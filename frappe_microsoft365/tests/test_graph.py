"""Graph transport tests: pooling, token caching, paging, delta links, throttling, expiry.

Nothing here touches the network — ``microsoft_graph._http_request`` is the one place a Graph
call reaches a socket, and it is mocked at the module boundary.
"""

from unittest.mock import MagicMock, patch

import frappe
from frappe.utils import add_to_date, now_datetime

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError, MsGraphResyncRequired
from frappe_microsoft365.tests.base import BaseTestCase


def _response(status=200, payload=None, headers=None):
	resp = MagicMock()
	resp.status_code = status
	resp.headers = headers or {}
	resp.content = b"{}"
	resp.reason = "Mocked"
	resp.json.return_value = payload if payload is not None else {}
	resp.text = "body"
	return resp


class TestGraphPaging(BaseTestCase):
	def test_paged_follows_next_link(self):
		pages = [
			{"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/page2"},
			{"value": [{"id": "b"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/page3"},
			{"value": [{"id": "c"}]},
		]
		with patch.object(graph, "graph_request", side_effect=pages) as mocked:
			items = graph.graph_paged("/me/events", "cal")

		self.assertEqual([i["id"] for i in items], ["a", "b", "c"])
		self.assertEqual(mocked.call_count, 3)

	def test_paged_stops_at_max_pages(self):
		endless = {"value": [{"id": "x"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"}
		with patch.object(graph, "graph_request", return_value=endless) as mocked:
			items = graph.graph_paged("/me/events", "cal", max_pages=3)

		self.assertEqual(len(items), 3)
		self.assertEqual(mocked.call_count, 3)

	def test_delta_returns_link_only_after_last_page(self):
		pages = [
			{"value": [{"id": "a"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/page2"},
			{"value": [{"id": "b"}], "@odata.deltaLink": "https://graph.microsoft.com/v1.0/delta?token=2"},
		]
		with patch.object(graph, "graph_request", side_effect=pages):
			items, delta_link = graph.graph_delta("/me/calendarView/delta", "cal")

		self.assertEqual([i["id"] for i in items], ["a", "b"])
		self.assertEqual(delta_link, "https://graph.microsoft.com/v1.0/delta?token=2")

	def test_delta_link_is_none_when_paging_is_cut_short(self):
		"""A truncated read must NOT hand back a watermark, or the rest is skipped forever."""
		endless = {"value": [{"id": "x"}], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"}
		with patch.object(graph, "graph_request", return_value=endless):
			items, delta_link = graph.graph_delta("/me/calendarView/delta", "cal", max_pages=2)

		self.assertEqual(len(items), 2)
		self.assertIsNone(delta_link)


class TestGraphRequestBehaviour(BaseTestCase):
	def setUp(self):
		super().setUp()
		self.token_patcher = patch.object(graph, "get_valid_access_token", return_value="token")
		self.token_patcher.start()
		self.addCleanup(self.token_patcher.stop)

	def test_429_waits_for_retry_after_then_retries(self):
		responses = [
			_response(429, headers={"Retry-After": "2"}),
			_response(200, {"value": []}),
		]
		with patch.object(graph, "_http_request", side_effect=responses) as mocked, patch.object(
			graph.time, "sleep"
		) as slept:
			result = graph.graph_request("GET", "/me/events", "cal")

		self.assertEqual(result, {"value": []})
		self.assertEqual(mocked.call_count, 2)
		slept.assert_called_once_with(2)

	def test_429_with_long_backoff_is_left_to_the_next_run(self):
		with patch.object(
			graph, "_http_request", return_value=_response(429, headers={"Retry-After": "600"})
		) as mocked, patch.object(graph.time, "sleep") as slept:
			with self.assertRaises(MsGraphError):
				graph.graph_request("GET", "/me/events", "cal")

		self.assertEqual(mocked.call_count, 1)
		slept.assert_not_called()

	def test_410_raises_resync_required(self):
		payload = {"error": {"code": "syncStateNotFound", "message": "resync"}}
		with patch.object(graph, "_http_request", return_value=_response(410, payload)):
			with self.assertRaises(MsGraphResyncRequired):
				graph.graph_request("GET", "/me/calendarView/delta", "cal")

	def test_401_refreshes_token_and_retries_once(self):
		responses = [_response(401), _response(200, {"ok": True})]
		with patch.object(graph, "_http_request", side_effect=responses) as mocked, patch.object(
			frappe.db, "set_value"
		):
			result = graph.graph_request("GET", "/me", "cal")

		self.assertEqual(result, {"ok": True})
		self.assertEqual(mocked.call_count, 2)

	def test_error_body_never_leaks_the_token(self):
		payload = {"error": {"code": "ErrorAccessDenied", "message": "no"}}
		with patch.object(graph, "_http_request", return_value=_response(403, payload)):
			with self.assertRaises(MsGraphError) as ctx:
				graph.graph_request("GET", "/me/events", "cal")

		self.assertNotIn("token", str(ctx.exception).lower())
		self.assertIn("ErrorAccessDenied", str(ctx.exception))


class TestConnectionPooling(BaseTestCase):
	def setUp(self):
		super().setUp()
		self.token_patcher = patch.object(graph, "get_valid_access_token", return_value="token")
		self.token_patcher.start()
		self.addCleanup(self.token_patcher.stop)

	def test_the_session_is_built_once_and_keeps_a_connection_pool(self):
		"""A handshake per call measured 44 ms against 11 ms on a socket that is already open."""
		graph._session = None
		self.addCleanup(setattr, graph, "_session", None)

		first = graph._graph_session()
		second = graph._graph_session()

		self.assertIs(first, second)
		adapter = first.get_adapter("https://graph.microsoft.com/v1.0/me")
		self.assertEqual(adapter._pool_connections, graph.POOL_CONNECTIONS)
		self.assertEqual(adapter._pool_maxsize, graph.POOL_MAXSIZE)

	def test_the_adapter_carries_no_retry_policy_of_its_own(self):
		"""graph_request owns 401/429; an adapter retrying underneath would multiply the attempts."""
		graph._session = None
		self.addCleanup(setattr, graph, "_session", None)

		adapter = graph._graph_session().get_adapter("https://graph.microsoft.com/v1.0/me")

		self.assertEqual(adapter.max_retries.total, 0)

	def test_consecutive_graph_calls_go_out_on_the_same_session(self):
		"""The point of pooling: a 50-page delta run must not open 50 connections."""
		session = graph._graph_session()
		responses = [_response(200, {"page": 1}), _response(200, {"page": 2})]
		with patch.object(session, "request", side_effect=responses) as mocked:
			graph.graph_request("GET", "/me/events", "cal")
			graph.graph_request("GET", "/me/events", "cal")

		# A second session would not carry this patch, so both calls arriving here is the proof.
		self.assertEqual(mocked.call_count, 2)
		self.assertIs(graph._graph_session(), session)


class TestTokenCache(BaseTestCase):
	def setUp(self):
		super().setUp()
		graph.clear_token_cache()
		self.addCleanup(graph.clear_token_cache)

	def _calendar_doc(self):
		doc = MagicMock()
		doc.token_expiry = add_to_date(now_datetime(), hours=1)
		return doc

	def test_a_second_lookup_for_the_same_calendar_does_not_read_the_token_again(self):
		"""Resolving a token is two queries plus an AES decrypt; a 50-page delta paid it per page."""
		with patch.object(
			graph.frappe, "get_doc", return_value=self._calendar_doc()
		) as loaded, patch.object(graph, "get_decrypted_password", return_value="tok") as decrypted:
			first = graph.get_valid_access_token("cal")
			second = graph.get_valid_access_token("cal")

		self.assertEqual([first, second], ["tok", "tok"])
		self.assertEqual(loaded.call_count, 1)
		self.assertEqual(decrypted.call_count, 1)

	def test_each_calendar_is_cached_under_its_own_name(self):
		"""One shared entry would hand one person's connection the token of another's."""
		with patch.object(graph.frappe, "get_doc", return_value=self._calendar_doc()), patch.object(
			graph, "get_decrypted_password", side_effect=["tok-a", "tok-b"]
		):
			self.assertEqual(graph.get_valid_access_token("cal-a"), "tok-a")
			self.assertEqual(graph.get_valid_access_token("cal-b"), "tok-b")

	def test_the_cache_lives_on_frappe_local_so_it_cannot_outlive_the_request(self):
		"""A module global survives into the next request, and one worker serves every site."""
		graph._cache_token("cal", "tok", add_to_date(now_datetime(), hours=1))
		self.assertEqual(frappe.local.microsoft365_token_cache["cal"][0], "tok")

		frappe.local.microsoft365_token_cache = None  # what the next request starts from

		with patch.object(graph.frappe, "get_doc", return_value=self._calendar_doc()), patch.object(
			graph, "get_decrypted_password", return_value="fresh"
		):
			self.assertEqual(graph.get_valid_access_token("cal"), "fresh")

	def test_a_token_seconds_from_expiry_is_not_served_from_the_cache(self):
		"""Serving one with seconds left hands Graph a token that dies in the middle of the call."""
		graph._cache_token("cal", "stale", add_to_date(now_datetime(), seconds=5))

		with patch.object(graph.frappe, "get_doc", return_value=self._calendar_doc()), patch.object(
			graph, "get_decrypted_password", return_value="fresh"
		):
			self.assertEqual(graph.get_valid_access_token("cal"), "fresh")

	def test_storing_new_tokens_replaces_the_cached_one(self):
		"""A refresh that left the old token cached would send the dead one on the very next call."""
		graph._cache_token("cal", "old", add_to_date(now_datetime(), hours=1))

		with patch.object(graph.frappe, "get_doc", return_value=MagicMock()):
			graph._store_tokens("cal", {"access_token": "new", "expires_in": 3600})

		# Served from the cache — an unpatched get_doc here would be the tell if it were not.
		self.assertEqual(graph.get_valid_access_token("cal"), "new")

	def test_a_401_invalidates_the_cached_token_so_the_retry_sends_a_fresh_one(self):
		"""Without this the retry reads the cache back, re-sends what Graph just refused, and loops."""
		graph._cache_token("cal", "dead", add_to_date(now_datetime(), hours=1))
		responses = [_response(401), _response(200, {"ok": True})]
		with patch.object(graph, "_http_request", side_effect=responses) as mocked, patch.object(
			graph.frappe, "get_doc", return_value=self._calendar_doc()
		), patch.object(graph, "get_decrypted_password", return_value="fresh"), patch.object(
			frappe.db, "set_value"
		):
			result = graph.graph_request("GET", "/me", "cal")

		self.assertEqual(result, {"ok": True})
		sent = [c.kwargs["headers"]["Authorization"] for c in mocked.call_args_list]
		self.assertEqual(sent, ["Bearer dead", "Bearer fresh"])


class TestStreamedRequests(BaseTestCase):
	def setUp(self):
		super().setUp()
		self.token_patcher = patch.object(graph, "get_valid_access_token", return_value="token")
		self.token_patcher.start()
		self.addCleanup(self.token_patcher.stop)

	def test_stream_is_passed_through_to_the_underlying_request(self):
		"""A Teams recording runs to gigabytes; buffered, it is the worker's memory that pays."""
		with patch.object(graph, "_http_request", return_value=_response(200)) as mocked:
			graph.graph_request("GET", "/me/onlineMeetings/m/recordings/r/content", "cal", raw=True, stream=True)

		self.assertTrue(mocked.call_args.kwargs["stream"])

	def test_an_ordinary_call_does_not_stream(self):
		"""Streaming a small JSON body would leave its connection checked out of the pool."""
		with patch.object(graph, "_http_request", return_value=_response(200, {"ok": True})) as mocked:
			graph.graph_request("GET", "/me", "cal")

		self.assertFalse(mocked.call_args.kwargs["stream"])

	def test_a_streamed_response_is_handed_back_untouched(self):
		"""Reading .content to decide what to return would pull the whole recording into RAM."""
		resp = _response(200)
		with patch.object(graph, "_http_request", return_value=resp):
			returned = graph.graph_request("GET", "/me/x/content", "cal", raw=True, stream=True)

		self.assertIs(returned, resp)
		resp.json.assert_not_called()
		resp.close.assert_not_called()

	def test_the_401_retry_is_streamed_too(self):
		"""A retry that dropped stream would buffer the recording the first attempt was streaming."""
		responses = [_response(401), _response(200)]
		with patch.object(graph, "_http_request", side_effect=responses) as mocked, patch.object(
			frappe.db, "set_value"
		):
			graph.graph_request("GET", "/me/x/content", "cal", raw=True, stream=True)

		self.assertEqual([c.kwargs["stream"] for c in mocked.call_args_list], [True, True])

	def test_the_429_retry_is_streamed_too(self):
		"""Same trap on the throttling path — the second attempt has to stream as well."""
		responses = [_response(429, headers={"Retry-After": "1"}), _response(200)]
		with patch.object(graph, "_http_request", side_effect=responses) as mocked, patch.object(
			graph.time, "sleep"
		):
			graph.graph_request("GET", "/me/x/content", "cal", raw=True, stream=True)

		self.assertEqual([c.kwargs["stream"] for c in mocked.call_args_list], [True, True])

	def test_a_failed_stream_gives_its_pooled_connection_back(self):
		"""An unread streamed body keeps its connection out of the pool for the life of the process."""
		resp = _response(403, {"error": {"code": "ErrorAccessDenied", "message": "no"}})
		with patch.object(graph, "_http_request", return_value=resp):
			with self.assertRaises(MsGraphError):
				graph.graph_request("GET", "/me/x/content", "cal", raw=True, stream=True)

		resp.close.assert_called_once()
