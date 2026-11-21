"""Connection robustness: dead grants, concurrent refreshes, throttling, account switches, scopes.

Only the network is faked: ``microsoft_graph._http_request`` for Graph and ``_msal_app`` for the
login service. graph_request, graph_url, the token storage and the error handling all run for
real, against real Microsoft Calendar rows that the test transaction rolls back.
"""

from contextlib import contextmanager
from unittest.mock import MagicMock, patch

import frappe
from frappe.utils import add_to_date, now_datetime
from frappe.utils.password import get_decrypted_password

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.tests.base import BaseTestCase

SETTINGS = frappe._dict(
	enabled=1,
	client_id="cid",
	tenant_id="tid",
	default_scopes="",
	use_calendar=1,
	redirect_uri="https://site.example.com/cb",
)


def _response(status=200, payload=None, headers=None):
	resp = MagicMock()
	resp.status_code = status
	resp.headers = headers or {}
	resp.content = b"{}" if payload is None else b"x"
	resp.reason = "Mocked"
	resp.json.return_value = payload if payload is not None else {}
	return resp


class FakeMsal:
	"""Stands in for msal.ConfidentialClientApplication: the login service, not Graph."""

	def __init__(self, refresh_result=None, code_result=None):
		self.refresh_result = refresh_result
		self.code_result = code_result
		self.refresh_calls = 0

	def acquire_token_by_refresh_token(self, refresh_token, scopes=None):
		self.refresh_calls += 1
		result = self.refresh_result
		if isinstance(result, Exception):
			raise result
		return dict(result)

	def acquire_token_by_authorization_code(self, code, scopes=None, redirect_uri=None):
		return dict(self.code_result)


class ConnectionTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		graph.clear_token_cache()
		self.addCleanup(graph.clear_token_cache)
		for patcher in (
			# Commits would end the test transaction; the rollback is what cleans up.
			patch.object(frappe.db, "commit"),
			patch.object(graph, "get_settings", lambda: SETTINGS),
			patch.object(graph, "_settings_doc", lambda: SETTINGS),
		):
			patcher.start()
			self.addCleanup(patcher.stop)
		self.user = self._user("robust-owner@example.com")

	def _user(self, email):
		if not frappe.db.exists("User", email):
			frappe.get_doc(
				{
					"doctype": "User",
					"email": email,
					"first_name": "Robust",
					"send_welcome_email": 0,
					"user_type": "System User",
					"roles": [{"role": "System Manager"}],
				}
			).insert(ignore_permissions=True)
		return email

	def _calendar(self, name="_Test Robust Calendar", expired=False, **values):
		if frappe.db.exists("Microsoft Calendar", name):
			frappe.delete_doc("Microsoft Calendar", name, force=True, ignore_permissions=True)
		doc = frappe.get_doc(
			{
				"doctype": "Microsoft Calendar",
				"account_name": name,
				"user": self.user,
				"enabled": 1,
				"authorized": 1,
				"pull_from_microsoft_calendar": 1,
				"push_to_microsoft_calendar": 1,
				"access_token": "tok-stored",
				"refresh_token": "rt-stored",
				"token_expiry": add_to_date(now_datetime(), hours=-1 if expired else 1),
				**values,
			}
		).insert(ignore_permissions=True)
		return doc.name

	@contextmanager
	def _msal(self, **kwargs):
		fake = FakeMsal(**kwargs)
		with patch.object(graph, "_msal_app", lambda settings=None: fake):
			yield fake

	def _notifications(self, calendar):
		return frappe.get_all(
			"Notification Log",
			filters={"document_type": "Microsoft Calendar", "document_name": calendar, "for_user": self.user},
			pluck="name",
		)


# --- B8: token storage never saves the document; one refresh for concurrent workers -------


FRESH_TOKENS = {"access_token": "tok-fresh", "refresh_token": "rt-fresh", "expires_in": 3600}


class TestTokenStorage(ConnectionTestCase):
	def test_a_refresh_in_another_worker_does_not_break_a_save_of_the_same_calendar(self):
		"""The production error: TimestampMismatchError after a concurrent token refresh."""
		cal = self._calendar()
		held = frappe.get_doc("Microsoft Calendar", cal)  # the sync's copy, read first

		graph._store_tokens(cal, FRESH_TOKENS)  # another worker refreshes meanwhile

		held.last_error = "saved after the refresh"
		held.save(ignore_permissions=True)  # raised TimestampMismatchError before 1.2.0
		self.assertEqual(
			frappe.db.get_value("Microsoft Calendar", cal, "last_error"), "saved after the refresh"
		)

	def test_tokens_are_stored_encrypted_without_touching_modified(self):
		cal = self._calendar()
		modified = frappe.db.get_value("Microsoft Calendar", cal, "modified")

		graph._store_tokens(cal, FRESH_TOKENS)

		self.assertEqual(get_decrypted_password("Microsoft Calendar", cal, "access_token"), "tok-fresh")
		self.assertEqual(get_decrypted_password("Microsoft Calendar", cal, "refresh_token"), "rt-fresh")
		row = frappe.db.get_value("Microsoft Calendar", cal, ["access_token", "modified"], as_dict=True)
		self.assertEqual(row.access_token, "*" * len("tok-fresh"))
		self.assertEqual(row.modified, modified)
		# A form save afterwards keeps the token (an empty column would have deleted it).
		frappe.get_doc("Microsoft Calendar", cal).save(ignore_permissions=True)
		self.assertEqual(get_decrypted_password("Microsoft Calendar", cal, "access_token"), "tok-fresh")

	def test_a_refresh_is_committed_straight_away(self):
		cal = self._calendar(expired=True)
		with self._msal(refresh_result=FRESH_TOKENS), patch.object(graph, "_http_request"):
			self.assertEqual(graph.get_valid_access_token(cal), "tok-fresh")
		self.assertGreaterEqual(frappe.db.commit.call_count, 2)  # before the re-read, after the write

	def test_a_worker_that_waited_for_the_lock_uses_the_token_the_other_one_minted(self):
		cal = self._calendar(expired=True)
		real_lock = graph._refresh_lock

		@contextmanager
		def lock_won_by_another_worker(name):
			# While this worker waited, the other one refreshed and committed.
			graph._store_tokens(name, {"access_token": "tok-other-worker", "expires_in": 3600})
			graph.clear_token_cache(name)
			with real_lock(name):
				yield

		with (
			self._msal(refresh_result=FRESH_TOKENS) as msal,
			patch.object(graph, "_refresh_lock", lock_won_by_another_worker),
			patch.object(graph, "_http_request", return_value=_response(200, {"id": "me"})) as http,
		):
			graph.graph_request("GET", "/me", cal)

		self.assertEqual(msal.refresh_calls, 0)
		self.assertEqual(http.call_args.kwargs["headers"]["Authorization"], "Bearer tok-other-worker")

	def test_the_refresh_lock_is_exclusive_per_connection(self):
		cal = self._calendar()
		cache = frappe.cache()
		key = cache.make_key(f"microsoft365:token-refresh:{cal}")

		with graph._refresh_lock(cal):
			other = cache.lock(key, timeout=5)
			self.assertFalse(other.acquire(blocking=False))  # a second worker has to wait
			another = cache.lock(cache.make_key("microsoft365:token-refresh:another calendar"), timeout=5)
			self.assertTrue(another.acquire(blocking=False))  # another connection does not
			another.release()
		self.assertTrue(other.acquire(blocking=False))  # released on the way out
		self.addCleanup(lambda: other.locked() and other.release())

		# A lock nobody gives back (a worker killed mid-refresh) delays a refresh, never stops it.
		ran = []
		with patch.object(graph, "TOKEN_REFRESH_LOCK_TIMEOUT_SECONDS", 0.2):
			with graph._refresh_lock(cal):
				ran.append(True)
		self.assertEqual(ran, [True])
