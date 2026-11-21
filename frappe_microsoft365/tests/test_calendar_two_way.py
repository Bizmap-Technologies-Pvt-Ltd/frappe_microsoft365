"""Two-way calendar consistency: what each side ends up holding after the other one changed.

All-day dates, the sync window, mirrors Outlook no longer has, recurring series, what a push
sends, events moving between calendars, and the transaction boundaries of a sync.

Graph is faked at ``microsoft_graph._http_request`` wherever a request is sent, so the URL guard,
token lookup and error handling of ``graph_request`` run for real; only the token itself and the
network are stand-ins.
"""

import json
from unittest.mock import MagicMock, patch

import frappe
from frappe.utils import add_to_date, get_datetime, now_datetime

from frappe_microsoft365 import microsoft_calendar_sync as sync
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.tests.test_calendar_sync import (
	CALENDAR,
	HEALTHY_QUEUE,
	SyncTestCase,
	ms_attendee,
	ms_event,
)

GRAPH = "https://graph.microsoft.com/v1.0"


def _response(status=200, payload=None):
	resp = MagicMock()
	resp.status_code = status
	resp.headers = {}
	resp.content = b"" if payload is None else json.dumps(payload).encode()
	resp.reason = "Mocked"
	resp.json.return_value = payload if payload is not None else {}
	resp.text = json.dumps(payload or {})
	return resp


class FakeGraph:
	"""Stands in for the network under ``graph_request``; records every request it is sent.

	``routes`` maps (method, url-prefix) to a payload, a status, or a callable(url, body).
	Anything not routed answers 404, so a request nobody expected shows up as a failure.
	"""

	def __init__(self, routes=None):
		self.routes = dict(routes or {})
		self.calls = []

	def __call__(self, method, url, **kwargs):
		body = kwargs.get("json")
		self.calls.append((method, url, body))
		for (want_method, prefix), answer in self.routes.items():
			if method == want_method and url.startswith(prefix):
				if callable(answer):
					answer = answer(url, body)
				if isinstance(answer, int):
					return _response(answer, {"error": {"code": "ErrorItemNotFound", "message": "x"}} if answer >= 400 else None)
				return _response(200, answer)
		return _response(404, {"error": {"code": "ErrorItemNotFound", "message": "not routed"}})

	def of(self, method):
		return [call for call in self.calls if call[0] == method]


class GraphTestCase(SyncTestCase):
	"""A calendar that pushes, with Graph faked only at the network."""

	def setUp(self):
		super().setUp()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "push_to_microsoft_calendar", 1)
		self.calendar.reload()
		for patcher in (
			patch.object(graph, "get_valid_access_token", return_value="token"),
			patch.object(sync.background, "health", return_value=HEALTHY_QUEUE),
		):
			patcher.start()
			self.addCleanup(patcher.stop)

	def _network(self, fake):
		return patch.object(graph, "_http_request", side_effect=fake)

	def _quiet_event(self, ms_id=None, **kwargs):
		"""A Frappe-originated Event inserted without its save reaching Graph."""
		frappe.flags.in_microsoft_sync = True
		try:
			return self._local_event(ms_id, **kwargs)
		finally:
			frappe.flags.in_microsoft_sync = False


# --- B3: all-day events ----------------------------------------------------------------


class TestAllDayEvents(GraphTestCase):
	def _pull_all_day(self, tz, ms_id):
		with patch.object(sync, "get_system_timezone", return_value=tz):
			self._pull_with(
				[
					ms_event(
						ms_id,
						subject="Christmas",
						isAllDay=True,
						start={"dateTime": "2026-12-25T00:00:00.0000000", "timeZone": "UTC"},
						end={"dateTime": "2026-12-26T00:00:00.0000000", "timeZone": "UTC"},
					)
				]
			)
		return self._event_for(ms_id)

	def test_a_pulled_all_day_event_keeps_its_date_west_of_utc(self):
		event = self._pull_all_day("America/New_York", "ms-allday-ny")
		self.assertEqual(get_datetime(event.starts_on), get_datetime("2026-12-25 00:00:00"))
		self.assertEqual(get_datetime(event.ends_on), get_datetime("2026-12-25 23:59:59"))
		self.assertTrue(event.all_day)

	def test_a_pulled_all_day_event_keeps_its_date_east_of_utc(self):
		event = self._pull_all_day("Asia/Kolkata", "ms-allday-in")
		self.assertEqual(get_datetime(event.starts_on), get_datetime("2026-12-25 00:00:00"))
		self.assertEqual(get_datetime(event.ends_on), get_datetime("2026-12-25 23:59:59"))

	def test_a_three_day_event_ends_on_its_last_day(self):
		start, end = sync._all_day_to_system(
			{"dateTime": "2026-12-24T00:00:00", "timeZone": "UTC"},
			{"dateTime": "2026-12-27T00:00:00", "timeZone": "UTC"},
		)
		self.assertEqual((str(start), str(end)), ("2026-12-24 00:00:00", "2026-12-26 23:59:59"))

	def test_a_frappe_all_day_event_pushes_with_midnight_start_and_end(self):
		fake = FakeGraph({("POST", f"{GRAPH}/me/events"): {"id": "AAMk-allday"}})
		for tz in ("America/New_York", "Asia/Kolkata"):
			with patch.object(sync, "get_system_timezone", return_value=tz):
				self._quiet_event(
					subject=f"Holiday {tz}",
					all_day=1,
					starts_on="2026-12-25 00:00:00",
					ends_on="2026-12-25 23:59:59",
				)
				fake.calls.clear()
				with self._network(fake):
					sync._push(self.calendar)
			(method, url, body), = fake.of("POST")
			self.assertEqual((method, url), ("POST", f"{GRAPH}/me/events"))
			self.assertTrue(body["isAllDay"])
			self.assertEqual(body["start"], {"dateTime": "2026-12-25T00:00:00", "timeZone": tz})
			self.assertEqual(body["end"], {"dateTime": "2026-12-26T00:00:00", "timeZone": tz})

	def test_an_all_day_event_with_no_end_is_one_day(self):
		start, end = sync._all_day_to_ms(get_datetime("2026-12-25 00:00:00"), None)
		self.assertEqual((start["dateTime"], end["dateTime"]), ("2026-12-25T00:00:00", "2026-12-26T00:00:00"))

	def test_an_end_already_at_the_next_midnight_is_not_pushed_a_day_further(self):
		start, end = sync._all_day_to_ms(get_datetime("2026-12-25 00:00:00"), get_datetime("2026-12-27 00:00:00"))
		self.assertEqual(end["dateTime"], "2026-12-27T00:00:00")


class TestCorrectPulledAllDayEventsPatch(SyncTestCase):
	def _stored(self, ms_id, starts_on, ends_on, pulled=1):
		event = frappe.get_doc(
			{
				"doctype": "Event",
				"subject": ms_id,
				"event_type": "Private",
				"all_day": 1,
				"starts_on": starts_on,
				"ends_on": ends_on,
				"custom_microsoft_calendar": CALENDAR,
				"custom_microsoft_event_id": ms_id,
				"custom_pulled_from_microsoft": pulled,
				"custom_sync_with_microsoft_calendar": 1,
			}
		)
		frappe.flags.in_microsoft_sync = True
		try:
			event.insert(ignore_permissions=True)
		finally:
			frappe.flags.in_microsoft_sync = False
		return event.name

	def _dates(self, name):
		row = frappe.db.get_value("Event", name, ["starts_on", "ends_on"], as_dict=True)
		return str(row.starts_on), str(row.ends_on)

	def test_shifted_pulled_events_are_put_back_and_the_patch_is_idempotent(self):
		from frappe_microsoft365.patches import correct_pulled_all_day_events as patch_module

		with patch.object(patch_module, "get_system_timezone", return_value="Asia/Kolkata"):
			# How 1.1.x stored a pulled 25 Dec on a Kolkata site.
			shifted = self._stored("ms-shifted", "2026-12-25 05:30:00", "2026-12-26 05:30:00")
			# Already right: a UTC midnight is not what 00:00 Kolkata converts back to.
			right = self._stored("ms-right", "2026-12-25 00:00:00", "2026-12-25 23:59:59")
			# Created in Frappe: never Outlook's to correct, whatever it looks like.
			local = self._stored("ms-local", "2026-12-25 05:30:00", "2026-12-26 05:30:00", pulled=0)
			patch_module.execute()
			first = [self._dates(n) for n in (shifted, right, local)]
			patch_module.execute()
			second = [self._dates(n) for n in (shifted, right, local)]

		self.assertEqual(first[0], ("2026-12-25 00:00:00", "2026-12-25 23:59:59"))
		self.assertEqual(first[1], ("2026-12-25 00:00:00", "2026-12-25 23:59:59"))
		self.assertEqual(first[2], ("2026-12-25 05:30:00", "2026-12-26 05:30:00"))
		self.assertEqual(first, second)

	def test_new_york_shift_is_corrected_too(self):
		from frappe_microsoft365.patches.correct_pulled_all_day_events import corrected_dates

		fixed = corrected_dates("2026-12-24 19:00:00", "2026-12-25 19:00:00", "America/New_York")
		self.assertEqual([str(v) for v in fixed], ["2026-12-25 00:00:00", "2026-12-25 23:59:59"])

	def test_a_time_that_was_not_a_utc_midnight_is_left_alone(self):
		from frappe_microsoft365.patches.correct_pulled_all_day_events import corrected_dates

		self.assertIsNone(corrected_dates("2026-12-25 07:00:00", "2026-12-26 07:00:00", "Asia/Kolkata"))

	def test_a_site_with_no_all_day_events_does_nothing(self):
		from frappe_microsoft365.patches import correct_pulled_all_day_events as patch_module

		with patch.object(frappe.db, "set_value") as written:
			patch_module.execute()
		written.assert_not_called()
