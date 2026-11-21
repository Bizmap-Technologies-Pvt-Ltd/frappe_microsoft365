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
	_ensure_user,
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


# --- B4: the look-ahead ----------------------------------------------------------------


class TestTheLookAheadNeverShrinks(GraphTestCase):
	def _pull_at(self, when):
		"""Run one pull with the clock at ``when``; returns the first URL it asked Graph for."""
		fake = FakeGraph({("GET", f"{GRAPH}/me/calendarView/delta"): {"value": [], "@odata.deltaLink": f"{GRAPH}/me/calendarView/delta?$deltatoken=next"}})
		self.calendar.reload()
		with patch.object(sync, "now_datetime", return_value=get_datetime(when)), self._network(fake):
			sync._pull(self.calendar)
		return fake.calls[0][1]

	def test_an_event_inside_the_look_ahead_is_pulled_late_in_the_windows_life(self):
		"""Regression (#21): the window was renewed 14 days before its end, so on day 166 of a
		180-day window an event 100 days ahead was outside it and not pulled."""
		first = self._pull_at("2026-09-13 12:00:00")
		self.assertIn("startDateTime=", first)
		window_end = get_datetime(frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_window_end"))

		for day in range(1, 200, 7):
			now = add_to_date(get_datetime("2026-09-13 12:00:00"), days=day)
			url = self._pull_at(now)
			if "startDateTime=" in url:
				window_end = get_datetime(frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_window_end"))
			# Whatever the run, the window it read covers an event booked the full look-ahead out.
			self.assertGreaterEqual(window_end, add_to_date(now, days=sync.WINDOW_FUTURE_DAYS), f"day {day}")

	def test_a_window_from_before_the_fix_is_renewed_on_its_first_run(self):
		frappe.db.set_value(
			"Microsoft Calendar",
			CALENDAR,
			{
				"delta_link": f"{GRAPH}/me/calendarView/delta?$deltatoken=OLD",
				# What 1.1.x stored: connection day + 180, here 69 days on.
				"delta_window_end": add_to_date(now_datetime(), days=111),
			},
		)
		url = self._pull_at(now_datetime())
		self.assertIn("startDateTime=", url)
		end = get_datetime(frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_window_end"))
		self.assertGreaterEqual(end, add_to_date(now_datetime(), days=sync.WINDOW_FUTURE_DAYS))

	def test_a_window_that_still_covers_the_look_ahead_is_reused(self):
		old = f"{GRAPH}/me/calendarView/delta?$deltatoken=OLD"
		frappe.db.set_value(
			"Microsoft Calendar",
			CALENDAR,
			{"delta_link": old, "delta_window_end": add_to_date(now_datetime(), days=sync.WINDOW_FUTURE_DAYS + 5)},
		)
		self.assertEqual(self._pull_at(now_datetime()), old)


# --- B5: mirrors Outlook no longer has ---------------------------------------------------


class TestMirrorsMissingFromOutlook(GraphTestCase):
	OLD = f"{GRAPH}/me/calendarView/delta?$deltatoken=OLD"

	def setUp(self):
		super().setUp()
		self._delete_setting(0)
		# Three mirrors from an earlier sync, inside the window.
		self._pull_with([ms_event(f"ms-o{i}", start={"dateTime": "2026-12-02T09:00:00", "timeZone": "UTC"},
			end={"dateTime": "2026-12-02T09:30:00", "timeZone": "UTC"}) for i in (1, 2, 3)])

	def _delete_setting(self, value):
		frappe.db.set_single_value("Microsoft Settings", "delete_orphaned_calendar_events", value)
		self.addCleanup(frappe.db.set_single_value, "Microsoft Settings", "delete_orphaned_calendar_events", 0)

	def _resync(self, old_answer):
		"""A run whose stored link Graph answers with ``old_answer``; the fresh read has o1 and o3."""
		frappe.db.set_value(
			"Microsoft Calendar", CALENDAR, {"delta_link": self.OLD, "delta_window_end": add_to_date(now_datetime(), days=5)}
		)
		self.calendar.reload()
		fake = FakeGraph(
			{
				("GET", f"{GRAPH}/me/calendarView/delta?startDateTime="): {
					"value": [ms_event("ms-o1", start={"dateTime": "2026-12-02T09:00:00", "timeZone": "UTC"},
						end={"dateTime": "2026-12-02T09:30:00", "timeZone": "UTC"}),
						ms_event("ms-o3", start={"dateTime": "2026-12-02T09:00:00", "timeZone": "UTC"},
						end={"dateTime": "2026-12-02T09:30:00", "timeZone": "UTC"})],
					"@odata.deltaLink": f"{GRAPH}/me/calendarView/delta?$deltatoken=FRESH",
				},
				("GET", self.OLD): old_answer,
			}
		)
		with self._network(fake), patch.object(frappe.db, "commit"):
			result = sync._sync_locked(self.calendar)
		return result, fake

	def _flag(self, ms_id):
		event = self._event_for(ms_id)
		return None if event is None else event.custom_microsoft_orphaned

	def test_a_mirror_missing_from_a_full_resync_is_flagged_not_deleted(self):
		result, _fake = self._resync({"value": [], "@odata.deltaLink": self.OLD + "-2"})

		self.assertTrue(result["ok"], result)
		self.assertEqual(self._flag("ms-o2"), 1)
		self.assertEqual((self._flag("ms-o1"), self._flag("ms-o3")), (0, 0))
		self.assertIn("Missing From Outlook", result["message"])

	def test_with_the_setting_on_it_is_deleted(self):
		self._delete_setting(1)
		result, _fake = self._resync({"value": [], "@odata.deltaLink": self.OLD + "-2"})

		self.assertIsNone(self._event_for("ms-o2"))
		self.assertEqual(result["deleted"], 1)
		self.assertIsNotNone(self._event_for("ms-o1"))

	def test_a_renewal_first_reads_the_old_link_to_the_end(self):
		"""A deletion still waiting behind the old link is applied as a deletion, not a flag."""
		result, fake = self._resync({"value": [{"id": "ms-o2", "@removed": {"reason": "deleted"}}],
			"@odata.deltaLink": self.OLD + "-2"})

		self.assertIsNone(self._event_for("ms-o2"))
		self.assertEqual(result["deleted"], 1)
		self.assertIn(("GET", self.OLD, None), fake.calls)

	def test_after_a_410_the_mirror_is_flagged(self):
		result, fake = self._resync(410)

		self.assertEqual(self._flag("ms-o2"), 1)
		self.assertEqual([c[1] for c in fake.calls].count(self.OLD), 1, "an expired link is not asked twice")

	def test_a_frappe_event_is_never_flagged(self):
		local = self._quiet_event("ms-local-unseen", starts_on="2026-12-02 10:00:00", ends_on="2026-12-02 11:00:00")
		self._resync({"value": [], "@odata.deltaLink": self.OLD + "-2"})
		self.assertEqual(frappe.db.get_value("Event", local.name, "custom_microsoft_orphaned"), 0)

	def test_a_flagged_mirror_that_comes_back_is_unflagged(self):
		self._resync({"value": [], "@odata.deltaLink": self.OLD + "-2"})
		self.assertEqual(self._flag("ms-o2"), 1)
		self._pull_with([ms_event("ms-o2", subject="Back again")])
		self.assertEqual(self._flag("ms-o2"), 0)

	def test_the_doctor_and_the_upgrade_report_list_them(self):
		self._resync({"value": [], "@odata.deltaLink": self.OLD + "-2"})
		name = self._event_for("ms-o2").name

		for check, findings in (
			("calendar.orphaned_events", sync.orphaned_event_findings()),
			("upgrade.calendar_orphaned_events", sync.upgrade_check_orphaned_events()),
		):
			mine = [f for f in findings if f.get("target") == CALENDAR]
			self.assertEqual(len(mine), 1, findings)
			self.assertEqual(mine[0]["check"], check)
			self.assertIn(name, mine[0]["detail"])

	def test_the_upgrade_check_is_registered_for_1_2_0(self):
		hooks = frappe.get_hooks("microsoft365_upgrade_checks")
		paths = [p for paths in (hooks.get("1.2.0") or []) for p in ([paths] if isinstance(paths, str) else paths)]
		self.assertIn("frappe_microsoft365.microsoft_calendar_sync.upgrade_check_orphaned_events", paths)


# --- B8: transaction boundaries of a sync --------------------------------------------------


class TestASyncDoesNotHoldItsLocks(GraphTestCase):
	def test_pull_results_are_committed_before_the_push_begins(self):
		timeline = []
		fake = FakeGraph(
			{
				("GET", f"{GRAPH}/me/calendarView/delta"): lambda url, body: {
					"value": [ms_event("ms-b8-pulled")],
					"@odata.deltaLink": f"{GRAPH}/me/calendarView/delta?$deltatoken=2",
				},
				("POST", f"{GRAPH}/me/events"): lambda url, body: {"id": f"AAMk-b8-{len(fake.of('POST'))}"},
			}
		)

		def network(method, url, **kwargs):
			timeline.append(method)
			return fake(method, url, **kwargs)

		for n in range(2):
			self._quiet_event(subject=f"Pending {n}")
		with patch.object(graph, "_http_request", side_effect=network), patch.object(
			frappe.db, "commit", side_effect=lambda *a, **k: timeline.append("COMMIT")
		):
			result = sync._sync_locked(self.calendar)

		self.assertTrue(result["ok"], result)
		first_post = timeline.index("POST")
		self.assertIn("COMMIT", timeline[:first_post], timeline)
		# And one commit behind every pushed event, so none of them waits on the next one.
		self.assertEqual(timeline[first_post:], ["POST", "COMMIT", "POST", "COMMIT", "COMMIT"])


# --- B10: one Frappe Event, exactly one Outlook event --------------------------------------


class TestOneEventOneMeeting(GraphTestCase):
	def test_a_second_create_for_an_event_that_already_has_an_id_is_skipped(self):
		event = self._quiet_event(subject="Board meeting")
		# Another worker created it a moment ago; this one still holds the old copy.
		frappe.db.set_value("Event", event.name, "custom_microsoft_event_id", "AAMk-first", update_modified=False)
		fake = FakeGraph({("POST", f"{GRAPH}/me/events"): {"id": "AAMk-second"}})

		with self._network(fake):
			self.assertFalse(sync._create_in_graph(CALENDAR, event.name, live_doc=event))

		self.assertEqual(fake.calls, [])
		self.assertEqual(event.custom_microsoft_event_id, "AAMk-first")

	def test_a_form_opened_before_the_create_cannot_unlink_the_event(self):
		event = self._quiet_event(subject="Quarterly review")
		stale = frappe.get_doc("Event", event.name)
		fake = FakeGraph({("POST", f"{GRAPH}/me/events"): {"id": "AAMk-only"}})
		with self._network(fake), patch.object(frappe.db, "commit"):
			sync._push(self.calendar)

		stale.subject = "Quarterly review (edited)"
		with patch.object(frappe, "enqueue") as enqueued:
			stale.save()

		self.assertEqual(frappe.db.get_value("Event", event.name, "custom_microsoft_event_id"), "AAMk-only")
		methods = [c[0][0] for c in enqueued.call_args_list]
		self.assertEqual(methods, ["frappe_microsoft365.microsoft_calendar_sync.patch_event_in_graph"])
		self.assertEqual(len(fake.of("POST")), 1)

	def test_an_ordinary_new_event_is_created_after_its_save_commits(self):
		fake = FakeGraph({("POST", f"{GRAPH}/me/events"): {"id": "AAMk-later"}})
		with self._network(fake), patch.object(frappe, "enqueue") as enqueued:
			event = self._local_event(subject="Plain")

		self.assertEqual(fake.calls, [], "nothing reaches Outlook before the save commits")
		(call,) = enqueued.call_args_list
		self.assertEqual(call[0][0], "frappe_microsoft365.microsoft_calendar_sync.create_event_in_graph")
		self.assertTrue(call[1]["enqueue_after_commit"])

		with self._network(fake), patch.object(frappe.db, "commit"):
			sync.create_event_in_graph(call[1]["event_name"])
		(method, url, body), = fake.calls
		self.assertEqual((method, url, body["subject"]), ("POST", f"{GRAPH}/me/events", "Plain"))
		self.assertEqual(frappe.db.get_value("Event", event.name, "custom_microsoft_event_id"), "AAMk-later")

	def _teams_fake(self):
		return FakeGraph(
			{
				("POST", f"{GRAPH}/me/events"): {
					"id": "AAMk-teams",
					"onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/x"},
				},
				("DELETE", f"{GRAPH}/me/events/AAMk-teams"): 204,
			}
		)

	def test_a_teams_meeting_is_deleted_again_when_its_save_rolls_back(self):
		fake = self._teams_fake()
		with self._network(fake):
			event = self._local_event(subject="Teams", custom_add_teams_meeting=1)
			self.assertEqual(event.custom_teams_join_url, "https://teams.microsoft.com/l/meetup-join/x")
			frappe.db.rollback()

		self.assertEqual([c[:2] for c in fake.calls], [
			("POST", f"{GRAPH}/me/events"),
			("DELETE", f"{GRAPH}/me/events/AAMk-teams"),
		])

	def test_a_rollback_to_a_savepoint_is_caught_at_the_commit(self):
		fake = self._teams_fake()
		with self._network(fake):
			frappe.db.savepoint("teams_save")
			self._local_event(subject="Teams", custom_add_teams_meeting=1)
			frappe.db.rollback(save_point="teams_save")
			frappe.db.commit()

		self.assertEqual([c[0] for c in fake.calls], ["POST", "DELETE"])

	def test_a_teams_meeting_that_is_kept_is_not_deleted(self):
		fake = self._teams_fake()
		with self._network(fake):
			self._local_event(subject="Teams", custom_add_teams_meeting=1)
			frappe.db.commit()

		self.assertEqual([c[0] for c in fake.calls], ["POST"])


# --- B2: a push sends what changed in Frappe, and only that ---------------------------------


class TestOutlookEditsSurviveAPush(GraphTestCase):
	ID = "AAMk-b2"

	def setUp(self):
		super().setUp()
		self.xavier = _ensure_user("m365-b2-xavier@example.com")
		self.fake = FakeGraph(
			{
				("POST", f"{GRAPH}/me/events"): {"id": self.ID},
				("PATCH", f"{GRAPH}/me/events/{self.ID}"): {"id": self.ID},
				("GET", f"{GRAPH}/me/events/{self.ID}?$select=attendees"): {
					"attendees": [
						ms_attendee("Xavier", self.xavier, "accepted"),
						ms_attendee("Yusuf (added in Outlook)", "yusuf@partner.example"),
					]
				},
			}
		)
		# Created by the push, the way every Frappe event reaches Outlook.
		self.event = self._quiet_event(
			subject="Design review",
			description="<p>Frappe agenda</p>",
			event_participants=[{"reference_doctype": "User", "reference_docname": self.xavier}],
		)
		with self._network(self.fake), patch.object(frappe.db, "commit"):
			sync._push(self.calendar)
		self.assertEqual(frappe.db.get_value("Event", self.event.name, "custom_microsoft_event_id"), self.ID)
		self.fake.calls.clear()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "last_sync", add_to_date(now_datetime(), seconds=1))
		self.calendar.reload()

	def _pulled_reply(self):
		"""Outlook's copy: Yusuf added, the agenda rewritten, and Xavier has accepted."""
		return ms_event(
			self.ID,
			subject="Design review",
			preview="Agenda rewritten in Outlook",
			start={"dateTime": "2026-09-15T09:00:00.0000000", "timeZone": "UTC"},
			end={"dateTime": "2026-09-15T09:30:00.0000000", "timeZone": "UTC"},
			location={"displayName": ""},
			attendees=[
				ms_attendee("Xavier", self.xavier, "accepted"),
				ms_attendee("Yusuf (added in Outlook)", "yusuf@partner.example"),
			],
		)

	def _edit(self, **values):
		event = frappe.get_doc("Event", self.event.name)
		event.update(values)
		with patch.object(frappe, "enqueue"):
			event.save()
		with self._network(self.fake):
			sync.patch_event_in_graph(self.event.name)

	def test_an_attendee_reply_pulled_from_outlook_does_not_cause_a_patch(self):
		"""Regression (#19): the pull saved the reply, and the same run's push sent Frappe's
		participants and description back over Outlook's."""
		# Synced a minute ago, untouched in Frappe since: only the pull's own save moves it.
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "last_sync", add_to_date(now_datetime(), minutes=-1))
		self.calendar.reload()
		frappe.db.set_value(
			"Event",
			self.event.name,
			{
				"starts_on": "2026-09-15 09:00:00",
				"ends_on": "2026-09-15 09:30:00",
				"location": "",
				"modified": add_to_date(now_datetime(), minutes=-2),
			},
			update_modified=False,
		)
		with patch.object(sync, "get_system_timezone", return_value="UTC"):
			self._pull_with([self._pulled_reply()])
			self.assertGreater(
				get_datetime(frappe.db.get_value("Event", self.event.name, "modified")),
				get_datetime(self.calendar.last_sync),
				"the pull's save makes the Event a push candidate",
			)
			self.assertIn("accepted", frappe.db.get_value("Event", self.event.name, "custom_microsoft_attendees"))
			with self._network(self.fake), patch.object(frappe.db, "commit"):
				sync._push(self.calendar)

		self.assertEqual(self.fake.calls, [])

	def test_an_attendee_added_in_outlook_survives_a_title_edit_in_frappe(self):
		self._edit(subject="Design review (moved to Thursday)")

		(method, url, body), = self.fake.calls
		self.assertEqual((method, url), ("PATCH", f"{GRAPH}/me/events/{self.ID}"))
		self.assertEqual(body, {"subject": "Design review (moved to Thursday)"})

	def test_a_participant_added_in_frappe_is_merged_into_outlooks_list(self):
		zara = _ensure_user("m365-b2-zara@example.com")
		event = frappe.get_doc("Event", self.event.name)
		event.append("event_participants", {"reference_doctype": "User", "reference_docname": zara})
		with patch.object(frappe, "enqueue"):
			event.save()
		with self._network(self.fake):
			sync.patch_event_in_graph(self.event.name)

		self.assertEqual([c[0] for c in self.fake.calls], ["GET", "PATCH"])
		body = self.fake.calls[1][2]
		self.assertEqual(sorted(body), ["attendees"])
		self.assertEqual(
			[a["emailAddress"]["address"] for a in body["attendees"]],
			[self.xavier, "yusuf@partner.example", zara],
		)

	def test_a_participant_removed_in_frappe_is_the_only_one_removed(self):
		event = frappe.get_doc("Event", self.event.name)
		event.event_participants = []
		with patch.object(frappe, "enqueue"):
			event.save()
		with self._network(self.fake):
			sync.patch_event_in_graph(self.event.name)

		body = self.fake.of("PATCH")[0][2]
		self.assertEqual([a["emailAddress"]["address"] for a in body["attendees"]], ["yusuf@partner.example"])

	def test_an_event_from_before_the_fix_gets_its_state_from_the_version_before_the_edit(self):
		frappe.db.set_value("Event", self.event.name, "custom_microsoft_sync_state", None, update_modified=False)
		frappe.db.set_value("Event", self.event.name, "modified", add_to_date(now_datetime(), hours=-1), update_modified=False)

		self._edit(subject="Renamed")

		self.assertEqual(self.fake.of("PATCH")[0][2], {"subject": "Renamed"})


# --- B6: series this app pushed ---------------------------------------------------------------


class TestOurSeriesFollowsOutlook(GraphTestCase):
	MASTER = "AAMk-series-master"
	DELTA = f"{GRAPH}/me/calendarView/delta"

	def setUp(self):
		super().setUp()
		frappe.db.set_value(
			"Microsoft Calendar",
			CALENDAR,
			{
				"delta_link": f"{self.DELTA}?$deltatoken=OLD",
				"delta_window_end": add_to_date(now_datetime(), days=sync.WINDOW_FUTURE_DAYS + 20),
			},
		)
		self.calendar.reload()
		self.series = self._quiet_event(
			self.MASTER,
			subject="Standup",
			starts_on="2026-12-01 10:00:00",
			ends_on="2026-12-01 10:15:00",
			repeat_this_event=1,
			repeat_on="Weekly",
			tuesday=1,
		)

	def _occurrence(self, ms_id, kind="occurrence", **extra):
		extra.setdefault("subject", "Standup")
		return ms_event(ms_id, type=kind, seriesMasterId=self.MASTER, **extra)

	def _pull(self, items, master_status=200):
		fake = FakeGraph(
			{
				("GET", f"{GRAPH}/me/events/{self.MASTER}?$select=id"): {"id": self.MASTER} if master_status == 200 else master_status,
				("GET", self.DELTA): {"value": items, "@odata.deltaLink": f"{self.DELTA}?$deltatoken=NEXT"},
			}
		)
		self.calendar.reload()
		with self._network(fake):
			result = sync._pull(self.calendar)
		return result, fake

	def _master_field(self):
		return frappe.db.get_value("Event", self.series.name, "custom_microsoft_series_master_id")

	def test_seeing_an_occurrence_confirms_the_series_lazily(self):
		self.assertFalse(self._master_field())
		self._pull([self._occurrence("occ-1")])
		self.assertEqual(self._master_field(), self.MASTER)
		self.assertFalse(frappe.db.exists("Event", {"custom_microsoft_event_id": "occ-1"}))

	def test_an_occurrence_changed_in_outlook_is_mirrored_beside_its_series(self):
		self._pull([self._occurrence("occ-2", kind="exception", subject="Standup (moved)")])
		mirror = self._event_for("occ-2")
		self.assertIsNotNone(mirror)
		self.assertEqual((mirror.subject, mirror.custom_pulled_from_microsoft), ("Standup (moved)", 1))
		self.assertEqual(mirror.custom_microsoft_series_master_id, self.MASTER)

	def test_deleting_an_app_created_series_in_outlook_removes_the_frappe_event(self):
		"""Regression (#23): only the occurrences arrive as @removed, so nothing matched."""
		self._pull([self._occurrence("occ-1"), self._occurrence("occ-2", kind="exception")])
		removals = [{"id": f"occ-{n}", "@removed": {"reason": "deleted"}} for n in (1, 2, 3)]

		(_upserted, removed), fake = self._pull(removals, master_status=404)

		self.assertIn(("GET", f"{GRAPH}/me/events/{self.MASTER}?$select=id", None), fake.calls)
		self.assertFalse(frappe.db.exists("Event", self.series.name))
		self.assertIsNone(self._event_for("occ-2"))
		self.assertEqual(removed, 2)

	def test_a_series_that_still_exists_is_kept(self):
		self._pull([self._occurrence("occ-1")])
		self._pull([{"id": "occ-1", "@removed": {"reason": "deleted"}}], master_status=200)
		self.assertTrue(frappe.db.exists("Event", self.series.name))

	def test_an_unconfirmed_series_is_never_deleted(self):
		_result, fake = self._pull([{"id": "occ-9", "@removed": {"reason": "deleted"}}], master_status=404)
		self.assertTrue(frappe.db.exists("Event", self.series.name))
		self.assertEqual(len(fake.calls), 1, "nothing to ask about a series Graph never confirmed")

	def test_a_title_edit_does_not_resend_the_recurrence(self):
		fake = FakeGraph({("PATCH", f"{GRAPH}/me/events/{self.MASTER}"): {"id": self.MASTER}})
		state = sync._sync_state_of(sync._event_to_graph_body(self.series))
		sync._write_sync_state(self.series.name, state)

		event = frappe.get_doc("Event", self.series.name)
		event.subject = "Daily standup"
		with self._network(fake):
			sync._patch_in_graph(CALENDAR, event)
		self.assertEqual(fake.of("PATCH")[0][2], {"subject": "Daily standup"})

		event.thursday = 1
		fake.calls.clear()
		with self._network(fake):
			sync._patch_in_graph(CALENDAR, event)
		body = fake.of("PATCH")[0][2]
		self.assertEqual(sorted(body), ["recurrence"])
		self.assertIn("thursday", body["recurrence"]["pattern"]["daysOfWeek"])


# --- B11: moving an Event to another calendar, and unticking sync ---------------------------

OTHER_CALENDAR = "_Test MS Calendar Second"


class TestMovingAndUnticking(GraphTestCase):
	def setUp(self):
		super().setUp()
		if frappe.db.exists("Microsoft Calendar", OTHER_CALENDAR):
			frappe.delete_doc("Microsoft Calendar", OTHER_CALENDAR, force=True, ignore_permissions=True)
		frappe.get_doc(
			{
				"doctype": "Microsoft Calendar",
				"account_name": OTHER_CALENDAR,
				"user": "Administrator",
				"enabled": 1,
				"authorized": 1,
				"pull_from_microsoft_calendar": 0,
				"push_to_microsoft_calendar": 1,
			}
		).insert(ignore_permissions=True)
		self.addCleanup(self._drop_other_calendar)
		self.tokens = {CALENDAR: "token-first", OTHER_CALENDAR: "token-second"}
		token = patch.object(graph, "get_valid_access_token", side_effect=lambda name: self.tokens[name])
		token.start()
		self.addCleanup(token.stop)

	def _drop_other_calendar(self):
		with patch.object(frappe, "enqueue"):
			for name in frappe.get_all("Event", filters={"custom_microsoft_calendar": OTHER_CALENDAR}, pluck="name"):
				frappe.delete_doc("Event", name, force=True, ignore_permissions=True)
		frappe.delete_doc("Microsoft Calendar", OTHER_CALENDAR, force=True, ignore_permissions=True)

	def _network_by_token(self, fake):
		"""Like _network, but records which mailbox (token) each request went to."""

		def call(method, url, **kwargs):
			token = kwargs["headers"]["Authorization"].replace("Bearer ", "")
			fake.tokens.append((method, token))
			return fake(method, url, **kwargs)

		fake.tokens = []
		return patch.object(graph, "_http_request", side_effect=call)

	def _run_jobs(self, enqueued, fake):
		with self._network_by_token(fake), patch.object(frappe.db, "commit"):
			for call in enqueued.call_args_list:
				method = frappe.get_attr(call[0][0])
				kwargs = {k: v for k, v in call[1].items() if k not in ("queue", "timeout", "job_id", "deduplicate", "enqueue_after_commit")}
				method(**kwargs)

	def test_changing_the_calendar_recreates_the_event_in_the_new_mailbox(self):
		"""Regression (#28): the old id was kept, so every patch to the new mailbox was a 404."""
		event = self._quiet_event("AAMk-in-first", subject="Move me")
		fake = FakeGraph(
			{
				("DELETE", f"{GRAPH}/me/events/AAMk-in-first"): 204,
				("POST", f"{GRAPH}/me/events"): {"id": "AAMk-in-second"},
			}
		)

		event.reload()
		event.custom_microsoft_calendar = OTHER_CALENDAR
		with patch.object(frappe, "enqueue") as enqueued, self._network(fake):
			event.save()
		self.assertEqual(fake.calls, [], "nothing is sent before the save commits")
		self.assertFalse(frappe.db.get_value("Event", event.name, "custom_microsoft_event_id"))

		self._run_jobs(enqueued, fake)

		self.assertEqual(fake.tokens, [("DELETE", "token-first"), ("POST", "token-second")])
		self.assertEqual([c[:2] for c in fake.calls], [
			("DELETE", f"{GRAPH}/me/events/AAMk-in-first"),
			("POST", f"{GRAPH}/me/events"),
		])
		self.assertEqual(frappe.db.get_value("Event", event.name, "custom_microsoft_event_id"), "AAMk-in-second")

	def test_an_ordinary_edit_does_not_move_anything(self):
		event = self._quiet_event("AAMk-stays", subject="Stay")
		event.reload()
		event.subject = "Stay, renamed"
		with patch.object(frappe, "enqueue") as enqueued:
			event.save()
		self.assertEqual(
			[c[0][0] for c in enqueued.call_args_list],
			["frappe_microsoft365.microsoft_calendar_sync.patch_event_in_graph"],
		)
		self.assertEqual(frappe.db.get_value("Event", event.name, "custom_microsoft_event_id"), "AAMk-stays")

	def test_the_pull_leaves_an_unticked_event_alone(self):
		"""Regression (#28): the pull ticked Sync with Microsoft Calendar again on every change."""
		local = self._quiet_event("AAMk-unticked", subject="Do not sync", custom_sync_with_microsoft_calendar=0)
		self._pull_with([ms_event("AAMk-mirror", subject="Mirror")])
		frappe.db.set_value("Event", self._event_for("AAMk-mirror").name, "custom_sync_with_microsoft_calendar", 0)

		self._pull_with([
			ms_event("AAMk-unticked", subject="Renamed in Outlook"),
			ms_event("AAMk-mirror", subject="Mirror renamed in Outlook"),
		])

		local.reload()
		mirror = self._event_for("AAMk-mirror")
		self.assertEqual((local.custom_sync_with_microsoft_calendar, local.subject), (0, "Do not sync"))
		self.assertEqual((mirror.custom_sync_with_microsoft_calendar, mirror.subject), (0, "Mirror"))

	def test_an_unticked_mirror_is_not_deleted_by_outlook(self):
		self._pull_with([ms_event("AAMk-keep")])
		name = self._event_for("AAMk-keep").name
		frappe.db.set_value("Event", name, "custom_sync_with_microsoft_calendar", 0)

		self._pull_with([{"id": "AAMk-keep", "@removed": {"reason": "deleted"}}])

		self.assertTrue(frappe.db.exists("Event", name))

	def test_a_new_mirror_is_still_ticked(self):
		self._pull_with([ms_event("AAMk-new-mirror")])
		self.assertEqual(self._event_for("AAMk-new-mirror").custom_sync_with_microsoft_calendar, 1)
