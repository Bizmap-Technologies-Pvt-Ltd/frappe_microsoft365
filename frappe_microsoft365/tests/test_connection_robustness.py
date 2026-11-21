"""Connection robustness: dead grants, concurrent refreshes, throttling, account switches, scopes.

Only the network is faked: ``microsoft_graph._http_request`` for Graph and ``_msal_app`` for the
login service. graph_request, graph_url, the token storage and the error handling all run for
real, against real Microsoft Calendar rows that the test transaction rolls back.
"""

from contextlib import contextmanager
from datetime import datetime, timedelta, timezone
from email.utils import format_datetime
from unittest.mock import MagicMock, patch

import frappe
from frappe.utils import add_to_date, now_datetime
from frappe.utils.password import get_decrypted_password

from frappe_microsoft365 import doctor
from frappe_microsoft365 import microsoft_calendar_sync as sync
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar import microsoft_calendar as mc
from frappe_microsoft365.microsoft_graph import MsGraphAuthorizationLost, MsGraphError, MsGraphThrottled
from frappe_microsoft365.patches import copy_authorized_scopes_to_calendars
from frappe_microsoft365.tests.base import BaseTestCase

SETTINGS = frappe._dict(
	enabled=1,
	client_id="cid",
	tenant_id="tid",
	default_scopes="",
	use_calendar=1,
	redirect_uri="https://site.example.com/cb",
)

HEALTHY_QUEUE = {"ok": True, "reasons": [], "fixes": [], "message": "Background jobs are running."}


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


# --- B7: a dead grant is detected once and not retried -----------------------------------


DEAD_GRANT = {
	"error": "invalid_grant",
	"error_description": "AADSTS50173: The provided grant has expired due to it being revoked.",
}


class TestDeadGrant(ConnectionTestCase):
	def test_invalid_grant_marks_the_connection_unauthorised_and_stops_further_attempts(self):
		cal = self._calendar(expired=True)
		with self._msal(refresh_result=DEAD_GRANT) as msal, patch.object(graph, "_http_request") as http:
			with self.assertRaises(MsGraphAuthorizationLost):
				graph.graph_request("GET", "/me/events", cal)
			# Later calls in this run, and the next runs, never ask Microsoft again.
			for _attempt in range(3):
				with self.assertRaises(MsGraphError):
					graph.graph_request("GET", "/me/events", cal)

		self.assertEqual(msal.refresh_calls, 1)
		http.assert_not_called()
		row = frappe.db.get_value(
			"Microsoft Calendar", cal, ["authorized", "authorization_error"], as_dict=True
		)
		self.assertEqual(row.authorized, 0)
		self.assertIn("AADSTS50173", row.authorization_error)
		self.assertEqual(len(self._notifications(cal)), 1)

	def test_the_scheduler_and_the_sync_leave_a_dead_connection_alone(self):
		cal = self._calendar(expired=True)
		with (
			self._msal(refresh_result=DEAD_GRANT) as msal,
			patch.object(graph, "_http_request") as http,
			patch.object(frappe, "log_error"),
		):
			first = sync.sync_calendar(cal)
			second = sync.sync_calendar(cal)
			queued = sync.sync_all()

		self.assertEqual(msal.refresh_calls, 1)
		http.assert_not_called()
		self.assertFalse(first["ok"])
		self.assertIn("not authorized", second["message"])
		self.assertNotIn(cal, (queued or {}).get("queued", []))
		self.assertEqual(len(self._notifications(cal)), 1)

	def test_interaction_required_is_a_dead_grant_too(self):
		cal = self._calendar(expired=True)
		result = {"error": "interaction_required", "error_description": "AADSTS50076: MFA is required."}
		with self._msal(refresh_result=result), patch.object(graph, "_http_request"):
			with self.assertRaises(MsGraphAuthorizationLost):
				graph.get_valid_access_token(cal)

		self.assertEqual(frappe.db.get_value("Microsoft Calendar", cal, "authorized"), 0)

	def test_a_transient_failure_never_marks_a_live_connection_dead(self):
		cal = self._calendar(expired=True)
		transient = [
			{"error": "temporarily_unavailable", "error_description": "AADSTS90033: try later"},
			{"error": "server_error"},
			ConnectionError("login.microsoftonline.com timed out"),
		]
		for failure in transient:
			with self._msal(refresh_result=failure), patch.object(graph, "_http_request"):
				with self.assertRaises(Exception) as ctx:
					graph.get_valid_access_token(cal)
			self.assertNotIsInstance(ctx.exception, MsGraphAuthorizationLost)

		self.assertEqual(frappe.db.get_value("Microsoft Calendar", cal, "authorized"), 1)
		self.assertEqual(self._notifications(cal), [])


# --- B9: throttling ----------------------------------------------------------------------


class TestThrottling(ConnectionTestCase):
	def setUp(self):
		super().setUp()
		patcher = patch.object(graph, "get_valid_access_token", return_value="token")
		patcher.start()
		self.addCleanup(patcher.stop)

	def test_a_503_is_retried(self):
		responses = [_response(503), _response(200, {"value": []})]
		with (
			patch.object(graph, "_http_request", side_effect=responses) as http,
			patch.object(graph.time, "sleep") as slept,
		):
			self.assertEqual(graph.graph_request("GET", "/me/events", "cal"), {"value": []})

		self.assertEqual(
			[(c.args[0], c.args[1]) for c in http.call_args_list],
			[("GET", "https://graph.microsoft.com/v1.0/me/events")] * 2,
		)
		slept.assert_called_once_with(graph.DEFAULT_RETRY_AFTER_SECONDS)

	def test_backoff_doubles_without_a_retry_after(self):
		responses = [_response(503), _response(503), _response(503), _response(200, {})]
		with (
			patch.object(graph, "_http_request", side_effect=responses),
			patch.object(graph.time, "sleep") as slept,
		):
			graph.graph_request("GET", "/me/events", "cal")
		self.assertEqual([c.args[0] for c in slept.call_args_list], [2, 4, 8])

	def test_retry_after_as_an_http_date_is_honoured(self):
		when = format_datetime(datetime.now(timezone.utc) + timedelta(seconds=40), usegmt=True)
		responses = [_response(429, headers={"Retry-After": when}), _response(200, {})]
		with (
			patch.object(graph, "_http_request", side_effect=responses),
			patch.object(graph.time, "sleep") as slept,
		):
			graph.graph_request("GET", "/me/events", "cal")
		self.assertIn(slept.call_args.args[0], (39, 40, 41))

	def test_retry_after_parsing(self):
		past = format_datetime(datetime.now(timezone.utc) - timedelta(seconds=30), usegmt=True)
		cases = {"7": 7, "0": 1, past: 1, "not a date": graph.DEFAULT_RETRY_AFTER_SECONDS}
		for raw, expected in cases.items():
			self.assertEqual(
				graph._retry_after_seconds(_response(429, headers={"Retry-After": raw})), expected, raw
			)

	def test_a_long_retry_after_stops_further_calls_for_that_calendar_this_run(self):
		with (
			patch.object(
				graph, "_http_request", return_value=_response(429, headers={"Retry-After": "120"})
			) as http,
			patch.object(graph.time, "sleep") as slept,
		):
			with self.assertRaises(MsGraphThrottled) as ctx:
				graph.graph_request("POST", "/me/events", "cal", json={"subject": "x"})
			with self.assertRaises(MsGraphThrottled):
				graph.graph_request("POST", "/me/events", "cal", json={"subject": "y"})

		self.assertEqual(ctx.exception.retry_after, 120)
		self.assertEqual(http.call_count, 1)
		self.assertEqual(http.call_args.args[:2], ("POST", "https://graph.microsoft.com/v1.0/me/events"))
		self.assertEqual(http.call_args.kwargs["json"], {"subject": "x"})
		slept.assert_not_called()
		self.assertTrue(graph.is_throttled("cal"))
		self.assertFalse(graph.is_throttled("another calendar"))

		# Another connection is a different mailbox and still gets called.
		with patch.object(graph, "_http_request", return_value=_response(200, {"ok": True})) as http:
			self.assertEqual(graph.graph_request("GET", "/me", "another calendar"), {"ok": True})

	def test_a_person_waiting_is_not_held_for_long_waits(self):
		frappe.local.request = MagicMock()
		self.addCleanup(lambda: setattr(frappe.local, "request", None))
		with (
			patch.object(
				graph, "_http_request", return_value=_response(429, headers={"Retry-After": "40"})
			) as http,
			patch.object(graph.time, "sleep") as slept,
		):
			with self.assertRaises(MsGraphThrottled):
				graph.graph_request("GET", "/me/events", "cal")
		slept.assert_not_called()
		self.assertEqual(http.call_count, 1)

	def test_a_504_is_retried_for_a_get_but_never_for_a_post(self):
		with (
			patch.object(graph, "_http_request", side_effect=[_response(504), _response(200, {})]) as http,
			patch.object(graph.time, "sleep"),
		):
			graph.graph_request("GET", "/me/events", "cal")
		self.assertEqual(http.call_count, 2)

		with (
			patch.object(graph, "_http_request", return_value=_response(504)) as http,
			patch.object(graph.time, "sleep"),
		):
			with self.assertRaises(MsGraphError) as ctx:
				graph.graph_request("POST", "/me/events", "cal", json={"subject": "x"})
		# A POST that timed out may have created the meeting; sending it again would duplicate it.
		self.assertEqual(http.call_count, 1)
		self.assertNotIsInstance(ctx.exception, MsGraphThrottled)
		self.assertFalse(graph.is_throttled("cal"))

	def test_a_refusal_that_outlasts_the_retries_pauses_the_connection(self):
		with (
			patch.object(graph, "_http_request", return_value=_response(503)) as http,
			patch.object(graph.time, "sleep"),
		):
			with self.assertRaises(MsGraphThrottled):
				graph.graph_request("GET", "/me/events", "cal")
		self.assertEqual(http.call_count, graph.MAX_THROTTLE_RETRIES + 1)
		self.assertTrue(graph.is_throttled("cal"))


class TestPushStopsWhenThrottled(ConnectionTestCase):
	def test_the_push_stops_calling_a_throttled_mailbox(self):
		cal = self._calendar(pull_from_microsoft_calendar=0)
		with (
			patch.object(sync.background, "health", return_value=HEALTHY_QUEUE),
			patch.object(frappe, "enqueue"),
		):
			frappe.flags.in_microsoft_sync = True
			try:
				for i in range(3):
					frappe.get_doc(
						{
							"doctype": "Event",
							"subject": f"pending {i}",
							"event_type": "Private",
							"starts_on": "2026-12-01 10:00:00",
							"ends_on": "2026-12-01 10:30:00",
							"custom_sync_with_microsoft_calendar": 1,
							"custom_microsoft_calendar": cal,
						}
					).insert(ignore_permissions=True)
			finally:
				frappe.flags.in_microsoft_sync = False
			frappe.db.sql(
				"update `tabEvent` set owner=%s where custom_microsoft_calendar=%s", (self.user, cal)
			)
			doc = frappe.get_doc("Microsoft Calendar", cal)
			with (
				patch.object(
					graph, "_http_request", return_value=_response(429, headers={"Retry-After": "120"})
				) as http,
				patch.object(frappe, "log_error") as logged,
			):
				pushed = sync._push(doc)

		self.assertEqual(pushed, 0)
		self.assertEqual(http.call_count, 1)
		self.assertEqual(http.call_args.args[:2], ("POST", "https://graph.microsoft.com/v1.0/me/events"))
		logged.assert_not_called()
		self.assertIn("throttling", doc.flags.push_warning)


# --- B12 / B13: what a sign-in records -----------------------------------------------------


CURRENT_SCOPES = ("User.Read", "Calendars.ReadWrite", "OnlineMeetings.ReadWrite")


class TestSignInRecords(ConnectionTestCase):
	OLD_DELTA = "https://graph.microsoft.com/v1.0/me/calendarView/delta?$deltatoken=OLD"

	def _sign_in(self, cal, claims, me, calendars=None):
		"""Run the real callback with Microsoft's answers faked at the network boundary."""
		state = "state-" + frappe.generate_hash(length=8)
		frappe.db.set_value(
			"Microsoft Calendar",
			cal,
			{"oauth_state": state, "oauth_state_expiry": add_to_date(now_datetime(), minutes=5)},
		)
		code_result = {
			"access_token": "tok-new",
			"refresh_token": "rt-new",
			"expires_in": 3600,
			"id_token_claims": claims,
		}
		calendars = calendars or [{"id": "AAMk-NEW", "name": "Calendar", "isDefaultCalendar": True}]

		def http(method, url, **kwargs):
			if url.endswith("/me"):
				return _response(200, me)
			if "/me/calendars" in url:
				return _response(200, {"value": calendars})
			return _response(404, {"error": {"code": "ErrorItemNotFound"}})

		frappe.local.response = frappe._dict()
		with (
			self._msal(code_result=code_result),
			patch.object(graph, "_http_request", side_effect=http) as sent,
			patch.object(frappe, "respond_as_web_page") as failure_page,
			patch.object(frappe, "log_error"),
		):
			mc.callback(code="code", state=state)
		failure_page.assert_not_called()
		return sent

	def _connected(self, **values):
		# Server-managed fields are wiped by validate, so they are written the way the app writes
		# them: straight to the row.
		server = {k: values.pop(k) for k in ("microsoft_account_id", "authorization_error") if k in values}
		cal = self._calendar(
			microsoft_user_email="a.old@contoso.example",
			ms_calendar_id="AAMk-OLD",
			ms_calendar_name="Old calendar",
			last_sync=add_to_date(now_datetime(), hours=-1),
			**values,
		)
		frappe.db.set_value(
			"Microsoft Calendar",
			cal,
			{
				"delta_link": self.OLD_DELTA,
				"delta_window_end": add_to_date(now_datetime(), days=90),
				**server,
			},
			update_modified=False,
		)
		return cal

	def _event(self, cal, ms_id, pulled=0):
		frappe.flags.in_microsoft_sync = True
		try:
			return (
				frappe.get_doc(
					{
						"doctype": "Event",
						"subject": ms_id,
						"event_type": "Private",
						"starts_on": "2026-12-01 10:00:00",
						"ends_on": "2026-12-01 10:30:00",
						"custom_sync_with_microsoft_calendar": 1,
						"custom_microsoft_calendar": cal,
						"custom_microsoft_event_id": ms_id,
						"custom_pulled_from_microsoft": pulled,
					}
				)
				.insert(ignore_permissions=True)
				.name
			)
		finally:
			frappe.flags.in_microsoft_sync = False

	def test_signing_in_as_another_account_clears_the_old_sync_state(self):
		cal = self._connected(microsoft_account_id="oid-old")
		local = self._event(cal, "old-local-1")
		pulled = self._event(cal, "old-pulled-1", pulled=1)

		sent = self._sign_in(
			cal,
			{"oid": "oid-new", "tid": "tenant-new", "preferred_username": "b.new@fabrikam.example"},
			{"id": "oid-new", "mail": "b.new@fabrikam.example"},
		)

		row = frappe.db.get_value("Microsoft Calendar", cal, ["*"], as_dict=True)
		self.assertFalse(row.delta_link)
		self.assertIsNone(row.delta_window_end)
		self.assertIsNone(row.last_sync)
		self.assertEqual(row.ms_calendar_id, "AAMk-NEW")  # the new account's default calendar
		self.assertEqual(row.microsoft_account_id, "oid-new")
		self.assertEqual(row.microsoft_user_email, "b.new@fabrikam.example")
		self.assertEqual(row.authorized, 1)
		# Graph was really called with the new account's token for /me and its calendars.
		self.assertEqual(
			[
				(c.args[0], c.args[1].split("?")[0], c.kwargs["headers"]["Authorization"])
				for c in sent.call_args_list
			],
			[
				("GET", "https://graph.microsoft.com/v1.0/me", "Bearer tok-new"),
				("GET", "https://graph.microsoft.com/v1.0/me/calendars", "Bearer tok-new"),
			],
		)
		# A Frappe-made Event no longer points into the old mailbox, and is no longer pushed.
		self.assertEqual(
			frappe.db.get_value(
				"Event", local, ["custom_microsoft_event_id", "custom_sync_with_microsoft_calendar"]
			),
			(None, 0),
		)
		# A pulled copy of an old-mailbox event is left as it was.
		self.assertEqual(frappe.db.get_value("Event", pulled, "custom_microsoft_event_id"), "old-pulled-1")
		self.assertTrue(
			frappe.db.exists(
				"Comment",
				{
					"reference_doctype": "Microsoft Calendar",
					"reference_name": cal,
					"content": ["like", "%different Microsoft account%"],
				},
			)
		)

	def test_signing_in_again_as_the_same_account_keeps_the_sync_state(self):
		cal = self._connected(microsoft_account_id="oid-old")
		local = self._event(cal, "old-local-1")

		self._sign_in(
			cal, {"oid": "oid-old", "preferred_username": "a.old@contoso.example"}, {"id": "oid-old"}
		)

		row = frappe.db.get_value(
			"Microsoft Calendar", cal, ["delta_link", "ms_calendar_id", "last_sync"], as_dict=True
		)
		self.assertEqual(row.delta_link, self.OLD_DELTA)
		self.assertEqual(row.ms_calendar_id, "AAMk-OLD")
		self.assertIsNotNone(row.last_sync)
		self.assertEqual(frappe.db.get_value("Event", local, "custom_microsoft_event_id"), "old-local-1")

	def test_a_connection_from_before_the_upgrade_is_matched_by_any_of_its_addresses(self):
		"""No stored id yet; the ID token says the UPN while the form recorded /me's mail."""
		cal = self._connected()

		self._sign_in(
			cal,
			{"oid": "oid-old", "preferred_username": "a.old.upn@contoso.example"},
			{
				"id": "oid-old",
				"mail": "A.Old@contoso.example",
				"userPrincipalName": "a.old.upn@contoso.example",
			},
		)

		row = frappe.db.get_value(
			"Microsoft Calendar", cal, ["delta_link", "microsoft_account_id"], as_dict=True
		)
		self.assertEqual(row.delta_link, self.OLD_DELTA)
		self.assertEqual(row.microsoft_account_id, "oid-old")  # recorded from now on

	def test_a_sign_in_records_its_scopes_on_the_connection_and_clears_a_lost_authorization(self):
		cal = self._connected(authorized=0, authorization_error="AADSTS50173: revoked")

		self._sign_in(cal, {"oid": "oid-old"}, {"id": "oid-old", "mail": "a.old@contoso.example"})

		row = frappe.db.get_value(
			"Microsoft Calendar",
			cal,
			["authorized", "authorization_error", "authorized_scopes"],
			as_dict=True,
		)
		self.assertEqual(row.authorized, 1)
		self.assertFalse(row.authorization_error)
		self.assertEqual(row.authorized_scopes, " ".join(graph.get_scopes(SETTINGS)))

	def test_recorded_scopes_and_account_id_cannot_be_typed_in(self):
		cal = self._calendar()
		frappe.db.set_value(
			"Microsoft Calendar", cal, {"authorized_scopes": "User.Read", "microsoft_account_id": "oid-1"}
		)
		doc = frappe.get_doc("Microsoft Calendar", cal)
		doc.authorized_scopes = "User.Read OnlineMeetings.ReadWrite"
		doc.microsoft_account_id = "oid-2"
		doc.save(ignore_permissions=True)
		self.assertEqual(
			frappe.db.get_value("Microsoft Calendar", cal, ["authorized_scopes", "microsoft_account_id"]),
			("User.Read", "oid-1"),
		)


class TestPerConnectionScopes(BaseTestCase):
	def test_the_doctor_flags_a_connection_authorised_before_a_capability_was_added(self):
		found = doctor.check_connection_scopes(
			list(CURRENT_SCOPES),
			[
				{"name": "A", "authorized": 1, "authorized_scopes": "User.Read Calendars.ReadWrite"},
				{"name": "B", "authorized": 1, "authorized_scopes": " ".join(list(CURRENT_SCOPES))},
			],
		)
		self.assertEqual([(f["check"], f["target"]) for f in found], [("scopes.connection_missing", "A")])
		self.assertIn("OnlineMeetings.ReadWrite", found[0]["title"])

	def test_nothing_recorded_means_nothing_reported(self):
		self.assertEqual(
			doctor.check_connection_scopes(
				list(CURRENT_SCOPES), [{"name": "A", "authorized": 1, "authorized_scopes": ""}]
			),
			[],
		)

	def test_a_connection_microsoft_stopped_accepting_is_reported(self):
		found = doctor.check_connection_scopes(
			list(CURRENT_SCOPES),
			[
				{
					"name": "A",
					"authorized": 0,
					"authorized_scopes": "User.Read",
					"authorization_error": "AADSTS50173",
				},
				{"name": "B", "authorized": 0, "authorized_scopes": "User.Read"},
			],
		)
		self.assertEqual([(f["check"], f["target"]) for f in found], [("calendar.authorization_lost", "A")])


class TestCopyAuthorizedScopesPatch(ConnectionTestCase):
	def test_copies_the_site_wide_value_onto_authorised_connections_and_runs_twice(self):
		site_wide = "User.Read Calendars.ReadWrite"
		previous = frappe.db.get_single_value("Microsoft Settings", "authorized_scopes")
		frappe.db.set_single_value("Microsoft Settings", "authorized_scopes", site_wide)
		self.addCleanup(frappe.db.set_single_value, "Microsoft Settings", "authorized_scopes", previous or "")
		empty = self._calendar("_Test Robust Empty")
		recorded = self._calendar("_Test Robust Recorded")
		frappe.db.set_value("Microsoft Calendar", recorded, "authorized_scopes", "User.Read")
		unauthorised = self._calendar("_Test Robust Off", authorized=0)

		copy_authorized_scopes_to_calendars.execute()
		copy_authorized_scopes_to_calendars.execute()

		scopes = dict(
			frappe.get_all(
				"Microsoft Calendar",
				filters={"name": ["in", [empty, recorded, unauthorised]]},
				fields=["name", "authorized_scopes"],
				as_list=True,
			)
		)
		self.assertEqual(scopes[empty], site_wide)
		self.assertEqual(scopes[recorded], "User.Read")  # a post-upgrade sign-in is never overwritten
		self.assertFalse(scopes[unauthorised])
		# Nobody is told to re-authorise because of the upgrade.
		connections = doctor._connection_scope_rows()
		ours = [c for c in connections if c.name == empty]
		self.assertEqual(doctor.check_connection_scopes(graph.parse_scopes(site_wide), ours), [])

	def test_does_nothing_when_nothing_was_ever_recorded(self):
		previous = frappe.db.get_single_value("Microsoft Settings", "authorized_scopes")
		frappe.db.set_single_value("Microsoft Settings", "authorized_scopes", "")
		self.addCleanup(frappe.db.set_single_value, "Microsoft Settings", "authorized_scopes", previous or "")
		cal = self._calendar()

		copy_authorized_scopes_to_calendars.execute()

		self.assertFalse(frappe.db.get_value("Microsoft Calendar", cal, "authorized_scopes"))
