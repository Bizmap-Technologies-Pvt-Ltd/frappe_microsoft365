"""Graph transport tests: paging, delta links, throttling, sync-state expiry.

Nothing here touches the network — ``requests.request`` is mocked at the module boundary.
"""

from unittest.mock import MagicMock, patch

import frappe

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
		with patch.object(graph.requests, "request", side_effect=responses) as mocked, patch.object(
			graph.time, "sleep"
		) as slept:
			result = graph.graph_request("GET", "/me/events", "cal")

		self.assertEqual(result, {"value": []})
		self.assertEqual(mocked.call_count, 2)
		slept.assert_called_once_with(2)

	def test_429_with_long_backoff_is_left_to_the_next_run(self):
		with patch.object(
			graph.requests, "request", return_value=_response(429, headers={"Retry-After": "600"})
		) as mocked, patch.object(graph.time, "sleep") as slept:
			with self.assertRaises(MsGraphError):
				graph.graph_request("GET", "/me/events", "cal")

		self.assertEqual(mocked.call_count, 1)
		slept.assert_not_called()

	def test_410_raises_resync_required(self):
		payload = {"error": {"code": "syncStateNotFound", "message": "resync"}}
		with patch.object(graph.requests, "request", return_value=_response(410, payload)):
			with self.assertRaises(MsGraphResyncRequired):
				graph.graph_request("GET", "/me/calendarView/delta", "cal")

	def test_401_refreshes_token_and_retries_once(self):
		responses = [_response(401), _response(200, {"ok": True})]
		with patch.object(graph.requests, "request", side_effect=responses) as mocked, patch.object(
			frappe.db, "set_value"
		):
			result = graph.graph_request("GET", "/me", "cal")

		self.assertEqual(result, {"ok": True})
		self.assertEqual(mocked.call_count, 2)

	def test_error_body_never_leaks_the_token(self):
		payload = {"error": {"code": "ErrorAccessDenied", "message": "no"}}
		with patch.object(graph.requests, "request", return_value=_response(403, payload)):
			with self.assertRaises(MsGraphError) as ctx:
				graph.graph_request("GET", "/me/events", "cal")

		self.assertNotIn("token", str(ctx.exception).lower())
		self.assertIn("ErrorAccessDenied", str(ctx.exception))
