"""The catch-up job and recurring meetings: when to ask, what to take, and when to stop.

Only the network is faked (tests/network.py), so graph_request, graph_url and the error text
every status is built from run for real, and the URLs asserted are the ones that would go out.
"""

import json
import re
from unittest.mock import patch

import frappe
from frappe.utils import add_to_date, get_datetime, now_datetime

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365 import microsoft_meeting_artifacts as artifacts
from frappe_microsoft365.microsoft_calendar_sync import _iso_utc
from frappe_microsoft365.tests.network import Response, error, fake_graph
from frappe_microsoft365.tests.test_meeting_artifacts import CALENDAR, JOIN_URL, ArtifactsTestCase

SPENT = len(artifacts.RETRY_MINUTES)


def utc(moment):
	"""A system-time datetime as Graph writes it."""
	return _iso_utc(moment)


class Graph:
	"""A tenant with one Teams meeting behind JOIN_URL. Every answer can be swapped per test."""

	def __init__(self, transcripts=(), recordings=(), resolve="meeting-1"):
		self.transcripts = list(transcripts)  # (id, created) pairs
		self.recordings = list(recordings)
		self.resolve = resolve  # a meeting id, None for no match, or a Response to answer with

	def __call__(self, method, url, kwargs):
		if "$filter=" in url:
			if isinstance(self.resolve, Response):
				return self.resolve
			return Response(200, {"value": [{"id": self.resolve}] if self.resolve else []})
		content = re.search(r"/transcripts/([^/]+)/content", url)
		if content:
			return Response(
				200, content=f"WEBVTT\n\n00:00:01.000 --> 00:00:02.000\n{content.group(1)}\n".encode()
			)
		if url.endswith("/transcripts"):
			return Response(200, {"value": [{"id": i, "createdDateTime": c} for i, c in self.transcripts]})
		if url.endswith("/recordings"):
			return Response(
				200,
				{"value": [{"id": i, "createdDateTime": c, "endDateTime": c} for i, c in self.recordings]},
			)
		if "/me/events/" in url:
			return Response(200, {"onlineMeeting": {"joinUrl": JOIN_URL}})
		return error(404, "NotFound", url)


class ScheduleTestCase(ArtifactsTestCase):
	filter_by_occurrence = True

	def setUp(self):
		super().setUp()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 1)
		self.addCleanup(
			frappe.db.set_value, "Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 0
		)
		settings = frappe.get_single("Microsoft Settings")
		restore = {"enabled": settings.enabled, "use_transcripts": settings.use_transcripts}
		frappe.db.set_value("Microsoft Settings", None, {"enabled": 1, "use_transcripts": 1})
		self.addCleanup(frappe.db.set_value, "Microsoft Settings", None, restore)

	def _meeting(self, ended_minutes_ago=20, minutes=30, **overrides):
		ends_on = add_to_date(now_datetime(), minutes=-ended_minutes_ago).replace(microsecond=0)
		return self._finished_meeting(
			starts_on=add_to_date(ends_on, minutes=-minutes), ends_on=ends_on, **overrides
		)

	def _fields(self, event):
		return frappe._dict(
			frappe.db.get_value(
				"Event",
				event.name,
				[
					"custom_microsoft_artifacts_attempts as attempts",
					"custom_microsoft_artifacts_status as status",
					"custom_microsoft_artifacts_checked_on as checked_on",
					"custom_microsoft_online_meeting_id as meeting_id",
					"custom_microsoft_recordings_data as recordings",
				],
				as_dict=True,
			)
		)

	def _files(self, event):
		return sorted(
			frappe.get_all(
				"File",
				filters={"attached_to_doctype": "Event", "attached_to_name": event.name},
				pluck="file_name",
			)
		)


class TestTemporaryErrorsKeepTheSchedule(ScheduleTestCase):
	"""B15: only a permanent answer ends the schedule, and every stop says why."""

	def _run(self, tenant):
		with fake_graph(tenant) as calls:
			artifacts.fetch_pending()
		return calls

	def test_a_503_resolving_the_join_link_leaves_the_event_scheduled(self):
		event = self._meeting(custom_microsoft_online_meeting_id=None)

		calls = self._run(Graph(resolve=error(503, "ServiceUnavailable", "Service unavailable")))

		self.assertEqual(calls[0][0], "GET")
		self.assertTrue(
			calls[0][1].startswith(
				"https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=JoinWebUrl%20eq%20"
			)
		)
		state = self._fields(event)
		self.assertEqual(state.attempts, 1)
		self.assertIn("503", state.status)
		self.assertIn("Trying again around", state.status)
		self.assertTrue(state.checked_on)
		self.assertIn(event.name, artifacts._due_events([CALENDAR], now=add_to_date(now_datetime(), hours=1)))

	def test_a_429_is_temporary_too(self):
		event = self._meeting(custom_microsoft_online_meeting_id=None)

		self._run(
			Graph(resolve=Response(429, {"error": {"code": "TooManyRequests"}}, headers={"Retry-After": "1"}))
		)

		state = self._fields(event)
		self.assertEqual(state.attempts, 1)
		self.assertIn("429", state.status)

	def test_a_token_that_would_not_refresh_is_temporary(self):
		event = self._meeting()

		with (
			fake_graph(Graph()),
			patch.object(
				graph, "get_valid_access_token", side_effect=graph.MsGraphError("token refresh failed")
			),
		):
			artifacts.fetch_pending()

		state = self._fields(event)
		self.assertEqual(state.attempts, 1)
		self.assertIn("token refresh failed", state.status)

	def test_a_meeting_from_another_organisation_ends_the_schedule_and_says_so(self):
		event = self._meeting(custom_microsoft_online_meeting_id=None)

		self._run(Graph(resolve=None))

		state = self._fields(event)
		self.assertEqual(state.attempts, SPENT)
		self.assertIn("Stopped checking", state.status)
		self.assertIn("another organisation", state.status)
		self.assertTrue(state.checked_on)

	def test_a_404_for_the_meeting_ends_the_schedule(self):
		event = self._meeting()

		self._run(lambda m, url, k: error(404, "NotFound", "meeting not found"))

		state = self._fields(event)
		self.assertEqual(state.attempts, SPENT)
		self.assertIn("404", state.status)

	def test_an_event_that_lost_its_teams_meeting_ends_the_schedule(self):
		event = self._meeting()
		event.db_set("custom_teams_join_url", None, update_modified=False)

		with patch.object(artifacts, "_due_events", return_value=[event.name]):
			self._run(Graph())

		state = self._fields(event)
		self.assertEqual(state.attempts, SPENT)
		self.assertIn("no Teams meeting", state.status)

	def test_the_last_temporary_failure_says_it_stopped(self):
		event = self._meeting(ended_minutes_ago=25 * 60, custom_microsoft_online_meeting_id=None)
		event.db_set("custom_microsoft_artifacts_attempts", SPENT - 1, update_modified=False)

		with patch.object(artifacts, "_due_events", return_value=[event.name]):
			self._run(Graph(resolve=error(503, "ServiceUnavailable")))

		state = self._fields(event)
		self.assertEqual(state.attempts, SPENT)
		self.assertIn(f"Stopped checking after {SPENT} attempts", state.status)

	def test_no_ending_leaves_a_spent_event_without_a_status(self):
		"""Production had Events at attempts=14 with no status, no checked time, no meeting id."""
		answers = [
			Graph(resolve=error(503, "ServiceUnavailable")),
			Graph(resolve=None),
			Graph(resolve=error(403, "Forbidden", "Insufficient privileges")),
			lambda m, url, k: error(404, "NotFound"),
			Graph(),
		]
		for tenant in answers:
			with self.subTest(tenant=tenant):
				event = self._meeting(custom_microsoft_online_meeting_id=None)
				self._run(tenant)
				state = self._fields(event)
				self.assertTrue(state.status)
				self.assertTrue(state.checked_on)
				frappe.delete_doc("Event", event.name, force=True, ignore_permissions=True)

	def test_the_button_calls_a_503_temporary_instead_of_quoting_the_filter(self):
		event = self._meeting(custom_microsoft_online_meeting_id=None)

		with fake_graph(Graph(resolve=error(503, "ServiceUnavailable"))):
			with self.assertRaises(graph.MsGraphError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("temporary", str(ctx.exception))
		self.assertNotIn("$filter", str(ctx.exception))


class TestOccurrences(ScheduleTestCase):
	"""B16: each occurrence of a recurring meeting gets only its own transcripts and recordings."""

	def _series(self):
		"""Three weekly occurrences of one Teams meeting, as Outlook sync creates them."""
		first = add_to_date(now_datetime(), days=-15).replace(hour=10, minute=0, second=0, microsecond=0)
		events = []
		for week in range(3):
			start = add_to_date(first, days=7 * week)
			events.append(self._finished_meeting(starts_on=start, ends_on=add_to_date(start, minutes=30)))
		transcripts = [(f"t{w}", utc(add_to_date(first, days=7 * w, minutes=20))) for w in range(3)]
		recordings = [(f"r{w}", utc(add_to_date(first, days=7 * w, minutes=1))) for w in range(3)]
		return events, Graph(transcripts=transcripts, recordings=recordings)

	def test_each_occurrence_gets_only_its_own_transcript(self):
		events, tenant = self._series()

		with fake_graph(tenant) as calls:
			for event in events:
				artifacts.fetch_meeting_artifacts(event.name)

		for week, event in enumerate(events):
			files = self._files(event)
			self.assertEqual(len(files), 1, files)
			day = tenant.transcripts[week][1][:10]
			self.assertTrue(files[0].startswith(f"teams-transcript-{day}"), (files, day))
		listed = [u for _m, u, _k in calls if u.endswith("/transcripts")]
		self.assertEqual(
			listed[0], "https://graph.microsoft.com/v1.0/me/onlineMeetings/meeting-1/transcripts"
		)

	def test_each_occurrence_gets_only_its_own_recording(self):
		events, tenant = self._series()

		with fake_graph(tenant):
			artifacts.fetch_meeting_artifacts(events[1].name)

		self.assertEqual([r["id"] for r in json.loads(self._fields(events[1]).recordings)], ["r1"])

	def test_a_late_refetch_of_the_first_week_stays_the_first_week(self):
		events, tenant = self._series()

		with fake_graph(tenant):
			artifacts.fetch_meeting_artifacts(events[0].name)
			artifacts.fetch_meeting_artifacts(events[0].name)

		self.assertEqual(len(self._files(events[0])), 1)

	def test_graph_times_are_utc_and_events_are_system_time(self):
		"""On an Asia/Kolkata site a 10:20 IST transcript is 04:50Z; both are the same moment."""
		event = self._meeting(ended_minutes_ago=120, minutes=60)
		start = get_datetime(event.starts_on)

		inside = utc(add_to_date(start, minutes=10))
		self.assertTrue(
			artifacts._in_window(
				{"created_date_time": inside}, artifacts._occurrence_window(*artifacts._occurrence(event))
			)
		)
		a_week_on = utc(add_to_date(start, days=7))
		self.assertFalse(
			artifacts._in_window(
				{"created_date_time": a_week_on}, artifacts._occurrence_window(*artifacts._occurrence(event))
			)
		)

	def test_graphs_seven_digit_fractions_are_read(self):
		self.assertIsNotNone(artifacts._graph_instant("2026-09-04T10:20:00.1234567Z"))


class TestARepeatingFrappeEvent(ScheduleTestCase):
	"""One Event with a repeat rule is the whole series; it is polled per occurrence."""

	def _weekly(self, weeks_ago=3, minutes=30, **extra):
		first = add_to_date(now_datetime(), days=-7 * weeks_ago, minutes=-60).replace(second=0, microsecond=0)
		return self._finished_meeting(
			starts_on=first,
			ends_on=add_to_date(first, minutes=minutes),
			repeat_this_event=1,
			repeat_on="Weekly",
			**extra,
		)

	def test_the_occurrence_is_the_latest_one_that_started(self):
		event = self._weekly()

		start, end = artifacts._occurrence(event)

		self.assertEqual(start, add_to_date(get_datetime(event.starts_on), days=21))
		self.assertEqual(end - start, get_datetime(event.ends_on) - get_datetime(event.starts_on))

	def test_daily_monthly_and_repeat_till(self):
		first = get_datetime("2026-01-31 09:00:00")
		daily = frappe._dict(
			starts_on=first, ends_on=add_to_date(first, hours=1), repeat_this_event=1, repeat_on="Daily"
		)
		monthly = frappe._dict(daily, repeat_on="Monthly")
		ended = frappe._dict(daily, repeat_till="2026-02-10")

		self.assertEqual(
			artifacts._occurrence(daily, get_datetime("2026-03-05 08:00"))[0],
			get_datetime("2026-03-04 09:00"),
		)
		self.assertEqual(
			artifacts._occurrence(monthly, get_datetime("2026-03-05 08:00"))[0],
			get_datetime("2026-02-28 09:00"),
		)
		self.assertEqual(
			artifacts._occurrence(ended, get_datetime("2026-03-05 08:00"))[0],
			get_datetime("2026-02-10 09:00"),
		)

	def test_a_weekly_event_is_due_again_the_week_after(self):
		event = self._weekly()
		# Spent and complete for its first week, as 1.1.x left it.
		event.db_set(
			{
				"custom_microsoft_artifacts_attempts": SPENT,
				"custom_microsoft_artifacts_checked_on": add_to_date(event.starts_on, hours=2),
				"custom_microsoft_transcript_fetched_on": add_to_date(event.starts_on, hours=2),
				"custom_microsoft_recordings_data": '[{"id": "old"}]',
			},
			update_modified=False,
		)

		self.assertIn(event.name, artifacts._due_events([CALENDAR]))

	def test_this_weeks_transcript_is_added_and_earlier_weeks_are_kept(self):
		event = self._weekly(weeks_ago=1)
		first = get_datetime(event.starts_on)
		tenant = Graph(
			transcripts=[
				("w0", utc(add_to_date(first, minutes=5))),
				("w1", utc(add_to_date(first, days=7, minutes=5))),
			],
			recordings=[("r0", utc(first)), ("r1", utc(add_to_date(first, days=7)))],
		)
		with (
			fake_graph(tenant),
			patch.object(artifacts, "now_datetime", return_value=add_to_date(first, hours=1)),
		):
			artifacts.fetch_meeting_artifacts(event.name)

		with fake_graph(tenant):
			artifacts.fetch_pending()

		files = self._files(event)
		self.assertEqual(len(files), 2, files)
		self.assertEqual([r["id"] for r in json.loads(self._fields(event).recordings)], ["r1"])


class TestTheJobDoesNotStarve(ScheduleTestCase):
	"""B17: only due Events are selected, missing-something first, and progress is committed."""

	def test_finished_events_do_not_crowd_out_a_due_one(self):
		finished = []
		for _ in range(4):
			event = self._meeting(ended_minutes_ago=180)
			event.db_set(
				{
					"custom_microsoft_transcript_fetched_on": now_datetime(),
					"custom_microsoft_recordings_data": '[{"id": "r"}]',
				},
				update_modified=False,
			)
			finished.append(event)
		due = self._meeting(ended_minutes_ago=20)

		with patch.object(artifacts, "MAX_EVENTS_PER_RUN", 3):
			selected = artifacts._due_events([CALENDAR])

		self.assertEqual(selected[0], due.name)
		self.assertEqual(len(selected), 3)

	def test_an_event_between_steps_is_not_selected(self):
		event = self._meeting(ended_minutes_ago=20)
		event.db_set("custom_microsoft_artifacts_attempts", 1, update_modified=False)  # next at +25

		self.assertNotIn(event.name, artifacts._due_events([CALENDAR]))
		later = add_to_date(now_datetime(), minutes=10)
		self.assertIn(event.name, artifacts._due_events([CALENDAR], now=later))

	def test_progress_is_committed_for_every_event(self):
		events = [self._meeting(ended_minutes_ago=20 + i) for i in range(3)]

		with fake_graph(Graph()):
			artifacts.fetch_pending()

		# A claim before Microsoft is asked, and the outcome after, for each one.
		self.assertGreaterEqual(self.commits.call_count, 2 * len(events))
		for event in events:
			self.assertEqual(self._fields(event).attempts, 1)

	def test_a_killed_worker_still_spends_the_attempt(self):
		"""What a SIGKILL looks like to this code: nothing after the Graph call runs."""
		event = self._meeting()

		class Killed(BaseException):
			pass

		def kill(*args):
			raise Killed()

		with fake_graph(kill), self.assertRaises(Killed):
			artifacts.fetch_pending()

		state = self._fields(event)
		self.assertEqual(state.attempts, 1)
		self.assertIn("interrupted", state.status)

	def test_a_later_recording_part_is_picked_up(self):
		event = self._meeting(ended_minutes_ago=60, minutes=300)
		start = get_datetime(event.starts_on)
		tenant = Graph(
			transcripts=[("t1", utc(add_to_date(start, minutes=1)))], recordings=[("r1", utc(start))]
		)
		with fake_graph(tenant):
			artifacts.fetch_pending()

		tenant.recordings.append(("r2", utc(add_to_date(start, hours=4))))
		with fake_graph(tenant) as calls:
			artifacts.fetch_pending()

		self.assertEqual([r["id"] for r in json.loads(self._fields(event).recordings)], ["r1", "r2"])
		# The transcript it already had was not downloaded again.
		self.assertFalse([u for _m, u, _k in calls if u.endswith("/content")])
