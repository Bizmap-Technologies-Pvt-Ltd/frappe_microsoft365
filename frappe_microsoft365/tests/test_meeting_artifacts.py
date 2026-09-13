"""Transcript and recording tests. Graph is mocked; no tenant, no network."""

import json
from unittest.mock import patch

import frappe
from frappe.utils import add_to_date, now_datetime
from werkzeug.wrappers import Response

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365 import microsoft_meeting_artifacts as artifacts
from frappe_microsoft365 import microsoft_transcripts as ms
from frappe_microsoft365.tests.base import BaseTestCase

CALENDAR = "_Test Artifacts Calendar"
JOIN_URL = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0"
VTT = "WEBVTT\n\n00:00:01.000 --> 00:00:04.000\nPriya: Let's begin.\n"
VTT_TWO = "WEBVTT\n\n01:00:01.000 --> 01:00:04.000\nPriya: Back after the break.\n"


class UnbufferableResponse:
	"""A Graph response that makes buffering fail loudly instead of quietly costing 1.5 GB.

	``.content`` and ``.text`` are the two ways requests reads a whole body into memory, and
	reading either is precisely the bug: a Teams recording part runs to 1.5 GB, which a web
	worker cannot hold. Raising here turns that regression into a failing test rather than a
	worker the OOM killer takes out in production.
	"""

	def __init__(self, chunks=(b"one", b"two", b"three"), headers=None):
		self.chunks = list(chunks)
		self.headers = {} if headers is None else headers
		self.chunk_size = None
		self.closed = False

	@property
	def content(self):
		raise AssertionError("the recording must never be read into memory in one piece")

	@property
	def text(self):
		raise AssertionError("the recording must never be read into memory in one piece")

	def iter_content(self, chunk_size=None):
		# Deliberately not a generator: the chunk size has to be recorded when it is asked for,
		# not on the first read, or a test that never iterates would see nothing.
		self.chunk_size = chunk_size
		return iter(self.chunks)

	def close(self):
		self.closed = True


#: A bench where jobs run, for the status sentences that are not about the queue.
#:
#: Every expectation in this file predates the background-health check and assumes something
#: is running the catch-up job. Left unpinned they would pass or fail on whether the developer
#: happened to have a worker up, which would say nothing about Graph — the actual subject.
HEALTHY_QUEUE = {"ok": True, "reasons": [], "fixes": []}

DEAD_QUEUE = {
	"ok": False,
	"reasons": ["The scheduler is off for this site, so nothing is queued on a schedule."],
	"fixes": ["Run `bench --site test enable-scheduler`."],
}


class ArtifactsTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		queue = patch.object(artifacts.background, "health", return_value=HEALTHY_QUEUE)
		queue.start()
		self.addCleanup(queue.stop)
		if not frappe.db.exists("Microsoft Calendar", CALENDAR):
			frappe.get_doc(
				{
					"doctype": "Microsoft Calendar",
					"account_name": CALENDAR,
					"user": "Administrator",
					"enabled": 1,
					"authorized": 1,
				}
			).insert(ignore_permissions=True)

	def _finished_meeting(self, ended_hours_ago=1, **overrides):
		values = {
			"doctype": "Event",
			"subject": "Retro",
			"starts_on": add_to_date(now_datetime(), hours=-ended_hours_ago - 1),
			"ends_on": add_to_date(now_datetime(), hours=-ended_hours_ago),
			"event_type": "Private",
			"custom_sync_with_microsoft_calendar": 1,
			"custom_microsoft_calendar": CALENDAR,
			"custom_teams_join_url": JOIN_URL,
			"custom_microsoft_online_meeting_id": "meeting-1",
		}
		values.update(overrides)
		doc = frappe.get_doc(values).insert(ignore_permissions=True)
		self.addCleanup(frappe.delete_doc, "Event", doc.name, force=True, ignore_permissions=True)
		return doc

	def _graph(self, transcripts=None, recordings=None, content=VTT):
		"""Patch the three Graph calls this module makes."""
		return (
			patch.object(ms, "list_transcripts", return_value=transcripts or []),
			patch.object(ms, "list_recordings", return_value=recordings or []),
			patch.object(ms, "get_transcript_content", return_value={"content": content}),
		)

	def _fetch(self, event, transcripts=None, recordings=None, content=VTT):
		a, b, c = self._graph(transcripts, recordings, content)
		with a, b, c:
			return artifacts.fetch_meeting_artifacts(event.name)


class TestFetchArtifacts(ArtifactsTestCase):
	def test_the_transcript_is_attached_to_the_event(self):
		event = self._finished_meeting()

		result = self._fetch(event, transcripts=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}])

		self.assertEqual(len(result["transcripts"]), 1)
		files = frappe.get_all(
			"File",
			filters={"attached_to_doctype": "Event", "attached_to_name": event.name},
			fields=["file_name"],
		)
		self.assertTrue(any(f.file_name.endswith(".vtt") for f in files))
		event.reload()
		self.assertTrue(event.custom_microsoft_transcript_fetched_on)

	def test_every_transcript_is_attached_not_just_the_latest(self):
		"""Transcription can be stopped and restarted; keeping only the last part loses the rest."""
		event = self._finished_meeting()

		with patch.object(
			ms,
			"list_transcripts",
			return_value=[
				{"id": "t2", "created_date_time": "2026-09-12T12:00:00Z"},
				{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"},
			],
		), patch.object(ms, "list_recordings", return_value=[]), patch.object(
			ms, "get_transcript_content", side_effect=[{"content": VTT}, {"content": VTT_TWO}]
		):
			result = artifacts.fetch_meeting_artifacts(event.name)

		self.assertEqual(len(result["transcripts"]), 2)
		names = frappe.get_all(
			"File",
			filters={"attached_to_doctype": "Event", "attached_to_name": event.name},
			pluck="file_name",
		)
		self.assertEqual(len(names), 2)
		self.assertTrue(any("part-1" in n for n in names), names)
		self.assertTrue(any("part-2" in n for n in names), names)

	def test_fetching_twice_does_not_leave_two_copies(self):
		event = self._finished_meeting()
		transcripts = [{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}]

		self._fetch(event, transcripts=transcripts)
		self._fetch(event, transcripts=transcripts)

		files = frappe.get_all(
			"File", filters={"attached_to_doctype": "Event", "attached_to_name": event.name}
		)
		self.assertEqual(len(files), 1)

	def test_recordings_are_noted_but_never_downloaded(self):
		"""A Teams recording is hundreds of megabytes; it stays with Microsoft."""
		event = self._finished_meeting()

		with patch.object(ms, "list_transcripts", return_value=[]), patch.object(
			ms, "list_recordings", return_value=[{"id": "r1", "created_date_time": "2026-09-12T11:00:00Z"}]
		), patch.object(ms, "get_transcript_content") as content:
			result = artifacts.fetch_meeting_artifacts(event.name)

		content.assert_not_called()
		self.assertEqual(result["recordings"], 1)
		event.reload()
		self.assertIn("Recorded", event.custom_microsoft_recordings)
		self.assertFalse(
			frappe.get_all("File", filters={"attached_to_name": event.name}),
			"the recording must not be copied into the site",
		)

	def test_a_long_meeting_keeps_every_recording_part(self):
		"""Teams splits a recording at 4 hours or 1.5 GB, so a 5-hour meeting is two files."""
		event = self._finished_meeting(ended_hours_ago=1)

		result = self._fetch(
			event,
			recordings=[
				{"id": "r1", "created_date_time": "2026-09-12T09:00:00Z", "end_date_time": "2026-09-12T13:00:00Z"},
				{"id": "r2", "created_date_time": "2026-09-12T13:00:00Z", "end_date_time": "2026-09-12T14:05:00Z"},
			],
		)

		self.assertEqual(result["recordings"], 2)
		event.reload()
		self.assertIn("Part 1 of 2", event.custom_microsoft_recordings)
		self.assertIn("Part 2 of 2", event.custom_microsoft_recordings)
		self.assertEqual([r["id"] for r in json.loads(event.custom_microsoft_recordings_data)], ["r1", "r2"])

	def test_a_missing_recording_permission_still_lands_the_transcript(self):
		"""Recordings need their own permission and licence; that must not cost the transcript."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._finished_meeting()

		with patch.object(
			ms, "list_transcripts", return_value=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}]
		), patch.object(ms, "get_transcript_content", return_value={"content": VTT}), patch.object(
			ms, "list_recordings", side_effect=MsGraphError("403 Forbidden")
		):
			result = artifacts.fetch_meeting_artifacts(event.name)

		self.assertEqual(len(result["transcripts"]), 1)
		self.assertEqual(result["recordings"], 0)

	def test_a_transcript_microsoft_will_not_hand_over_is_reported_not_hidden(self):
		"""Listing a transcript and then refusing its content is a refusal, not a delay.

		The refusal used to be swallowed by the loop that reads each part, so the fetch came back
		empty and the form said "Microsoft has not finished processing this meeting" — about a
		transcript that was finished and being withheld, with the hint that named the fix thrown
		away on the way past.
		"""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._finished_meeting()

		with patch.object(
			ms, "list_transcripts", return_value=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}]
		), patch.object(
			ms, "get_transcript_content", side_effect=MsGraphError("403 Forbidden: read the hint")
		), patch.object(ms, "list_recordings", return_value=[]):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("read the hint", str(ctx.exception))

	def test_one_refused_part_does_not_cost_the_parts_that_worked(self):
		"""Transcription restarts mid-meeting, so parts fail independently. Reporting the refusal
		must not go so far as to throw away an hour of transcript that came back fine."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._finished_meeting()

		with patch.object(
			ms,
			"list_transcripts",
			return_value=[
				{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"},
				{"id": "t2", "created_date_time": "2026-09-12T12:00:00Z"},
			],
		), patch.object(
			ms, "get_transcript_content", side_effect=[{"content": VTT}, MsGraphError("403 Forbidden")]
		), patch.object(ms, "list_recordings", return_value=[]):
			result = artifacts.fetch_meeting_artifacts(event.name)

		self.assertEqual(len(result["transcripts"]), 1)

	def test_an_error_from_microsoft_is_repeated_not_reinterpreted(self):
		"""The one thing worse than a Graph error is a guess about what it meant."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._finished_meeting()

		with patch.object(
			ms, "list_transcripts", side_effect=MsGraphError("ErrorAccessDenied: Access is denied")
		), patch.object(ms, "list_recordings", return_value=[]):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("ErrorAccessDenied", str(ctx.exception))


class TestWhatItSays(ArtifactsTestCase):
	"""An empty answer from Graph means different things at different times."""

	def test_minutes_after_the_meeting_it_says_still_processing(self):
		event = self._finished_meeting(ended_hours_ago=0)
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-5), update_modified=False)
		event.reload()

		result = self._fetch(event)

		self.assertEqual(result["state"], "processing")
		self.assertNotIn("no transcript", result["message"].lower())
		self.assertIn("processing", result["message"].lower())

	def test_it_never_promises_a_check_that_nobody_scheduled(self):
		"""With automatic fetching off, "checking again at 4pm" is simply untrue."""
		event = self._finished_meeting(ended_hours_ago=0)
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-5), update_modified=False)
		event.reload()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 0)

		message = self._fetch(event)["message"]

		self.assertIn("Nothing is checking automatically", message)

	def test_with_automatic_fetching_on_it_names_the_next_check(self):
		event = self._finished_meeting(ended_hours_ago=0)
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-5), update_modified=False)
		event.reload()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 1)
		self.addCleanup(
			frappe.db.set_value, "Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 0
		)
		frappe.clear_cache(doctype="Microsoft Calendar")

		message = self._fetch(event)["message"]

		self.assertIn("Checking again", message)

	def test_a_dead_scheduler_makes_the_promised_time_a_lie_and_it_says_so(self):
		"""The tickbox only decides whether the catch-up job would ask; something still has to
		run it. With nothing consuming the queue, "Checking again around 14:48" names a time
		that nothing can keep — the same class of lie as the tickbox being off, but worse,
		because it is specific."""
		event = self._finished_meeting(ended_hours_ago=0)
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-5), update_modified=False)
		event.reload()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 1)
		self.addCleanup(
			frappe.db.set_value, "Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 0
		)
		frappe.clear_cache(doctype="Microsoft Calendar")

		with patch.object(artifacts.background, "health", return_value=DEAD_QUEUE):
			message = self._fetch(event)["message"]

		self.assertNotIn("Checking again", message)
		self.assertIn("Nothing is checking automatically", message)
		self.assertIn("enable-scheduler", message)
		self.assertIn("Get Transcript & Recording", message)

	def test_a_day_later_it_says_it_was_never_recorded(self):
		event = self._finished_meeting(ended_hours_ago=30)
		event.db_set("custom_microsoft_artifacts_attempts", len(artifacts.RETRY_MINUTES), update_modified=False)
		event.reload()

		result = self._fetch(event)

		self.assertEqual(result["state"], "nothing_found")
		self.assertIn("never recorded", result["message"].lower())

	def test_a_day_later_it_also_says_what_to_go_and_look_at(self):
		"""The old ending, "you can still check by hand", named no hand and no thing to check.

		Graph answering 200-with-nothing all day looks identical whether nobody recorded the
		meeting or this connection cannot reach what exists — and this sentence is built from the
		Event's own fields, so it never sees the 403 the catch-up job stored. One look at the
		meeting in Teams separates the two, which makes it the only advice worth giving here.
		"""
		event = self._finished_meeting(ended_hours_ago=30)
		event.db_set("custom_microsoft_artifacts_attempts", len(artifacts.RETRY_MINUTES), update_modified=False)
		event.reload()

		message = self._fetch(event)["message"]

		self.assertIn("Recordings and Transcripts", message)
		self.assertIn("Run Diagnostics", message)
		self.assertNotIn("by hand", message)

	def test_after_sixty_days_it_says_the_meeting_is_too_old(self):
		event = self._finished_meeting(ended_hours_ago=24 * (artifacts.MEETING_EXPIRY_DAYS + 5))

		result = self._fetch(event)

		self.assertEqual(result["state"], "expired")
		self.assertIn("too old", result["message"].lower())

	def test_too_old_to_fetch_is_not_the_same_as_gone(self):
		"""Graph stops serving at 60 days; Teams keeps the files on a retention policy that
		defaults to 120. "Too old" alone sends someone away from a recording still sitting in
		the meeting's own tab."""
		event = self._finished_meeting(ended_hours_ago=24 * (artifacts.MEETING_EXPIRY_DAYS + 5))

		message = self._fetch(event)["message"]

		self.assertIn("may still exist", message)
		self.assertIn("Recordings and Transcripts", message)

	def test_one_artifact_that_never_arrived_points_at_the_meeting_not_the_setup(self):
		"""The transcript landing proves the connection and the permissions work, so the missing
		recording is about the meeting. Sending someone to Azure from here wastes an afternoon."""
		event = self._finished_meeting(ended_hours_ago=24 * (artifacts.MEETING_EXPIRY_DAYS + 5))

		result = self._fetch(event, transcripts=[{"id": "t1", "created_date_time": "2026-07-12T11:00:00Z"}])

		self.assertEqual(result["state"], "partial")
		self.assertIn("never started", result["message"])
		self.assertIn("Recordings and Transcripts", result["message"])

	def test_a_finished_fetch_says_what_landed(self):
		event = self._finished_meeting()

		result = self._fetch(
			event,
			transcripts=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}],
			recordings=[{"id": "r1", "created_date_time": "2026-09-12T11:00:00Z"}],
		)

		self.assertEqual(result["state"], "complete")
		self.assertIn("Transcript attached", result["message"])
		self.assertIn("1 recording", result["message"])

	def test_the_status_is_stored_on_the_event_for_the_form(self):
		event = self._finished_meeting()

		self._fetch(event)

		event.reload()
		self.assertTrue(event.custom_microsoft_artifacts_status)
		self.assertTrue(event.custom_microsoft_artifacts_checked_on)


class TestTheCatchUpJob(ArtifactsTestCase):
	def setUp(self):
		super().setUp()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 1)
		self.addCleanup(
			frappe.db.set_value, "Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 0
		)
		settings = frappe.get_single("Microsoft Settings")
		self._restore = {"enabled": settings.enabled, "use_transcripts": settings.use_transcripts}
		frappe.db.set_value("Microsoft Settings", None, {"enabled": 1, "use_transcripts": 1})
		self.addCleanup(frappe.db.set_value, "Microsoft Settings", None, self._restore)

	def test_the_backoff_widens_instead_of_polling(self):
		"""Fourteen attempts over a day, not ninety-six — and the first is ten minutes in."""
		self.assertEqual(artifacts.RETRY_MINUTES[0], 10)
		self.assertEqual(artifacts.RETRY_MINUTES[-1], 24 * 60)
		self.assertEqual(list(artifacts.RETRY_MINUTES), sorted(artifacts.RETRY_MINUTES))

	def test_a_meeting_that_just_ended_is_not_asked_about_yet(self):
		"""Asking at minute one wastes a call: Microsoft has not started processing."""
		event = self._finished_meeting()
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-2), update_modified=False)

		self.assertNotIn(event.name, artifacts._due_events([CALENDAR]))

	def test_a_meeting_past_its_first_step_is_due(self):
		event = self._finished_meeting()
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-20), update_modified=False)

		self.assertIn(event.name, artifacts._due_events([CALENDAR]))

	def test_a_meeting_with_both_artifacts_is_never_asked_about_again(self):
		event = self._finished_meeting()
		event.db_set(
			{
				"custom_microsoft_transcript_fetched_on": now_datetime(),
				"custom_microsoft_recordings_data": '[{"id": "r1"}]',
			},
			update_modified=False,
		)

		self.assertNotIn(event.name, artifacts._due_events([CALENDAR]))

	def test_the_job_gives_up_after_the_last_step(self):
		event = self._finished_meeting(ended_hours_ago=48)
		event.db_set(
			"custom_microsoft_artifacts_attempts", len(artifacts.RETRY_MINUTES), update_modified=False
		)

		self.assertNotIn(event.name, artifacts._due_events([CALENDAR]))

	def test_the_job_only_touches_calendars_that_opted_in(self):
		event = self._finished_meeting()
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-20), update_modified=False)
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "fetch_artifacts_automatically", 0)

		with patch.object(artifacts, "_fetch_into") as fetch:
			artifacts.fetch_pending()

		fetch.assert_not_called()

	def test_the_job_attaches_what_it_finds(self):
		event = self._finished_meeting()
		event.db_set("ends_on", add_to_date(now_datetime(), minutes=-20), update_modified=False)

		a, b, c = self._graph(transcripts=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}])
		with a, b, c:
			artifacts.fetch_pending()

		event.reload()
		self.assertTrue(event.custom_microsoft_transcript_fetched_on)

	def test_one_broken_event_does_not_stop_the_rest(self):
		broken = self._finished_meeting(custom_microsoft_calendar=CALENDAR)
		broken.db_set("ends_on", add_to_date(now_datetime(), minutes=-20), update_modified=False)
		broken.db_set("custom_teams_join_url", None, update_modified=False)

		# It is no longer addressable, so it is spent rather than retried every 15 minutes.
		artifacts.fetch_pending()

		broken.reload()
		self.assertNotIn(broken.name, artifacts._due_events([CALENDAR]))


class TestGuards(ArtifactsTestCase):
	def test_a_meeting_that_has_not_started_is_refused(self):
		event = self._finished_meeting(
			starts_on=add_to_date(now_datetime(), hours=1),
			ends_on=add_to_date(now_datetime(), hours=2),
		)

		with self.assertRaises(frappe.ValidationError) as ctx:
			artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("not started", str(ctx.exception))

	def test_a_meeting_that_ended_early_can_be_fetched_inside_its_own_slot(self):
		"""A calendar slot is a reservation, not a record.

		An eighty-minute booking used for a one-minute call has its transcript ready within
		minutes; gating on the booked end refused to look for another hour while the files sat
		waiting in Teams. Reported from a real meeting.
		"""
		event = self._finished_meeting(
			starts_on=add_to_date(now_datetime(), minutes=-10),
			ends_on=add_to_date(now_datetime(), hours=1),
		)

		result = self._fetch(
			event, transcripts=[{"id": "t1", "created_date_time": "2026-09-13T11:05:00Z"}]
		)

		self.assertEqual(len(result["transcripts"]), 1)

	def test_an_empty_answer_inside_the_slot_does_not_claim_the_meeting_is_over(self):
		"""Two things are true at once and it would be wrong to assert either."""
		event = self._finished_meeting(
			starts_on=add_to_date(now_datetime(), minutes=-10),
			ends_on=add_to_date(now_datetime(), hours=1),
		)

		message = self._fetch(event)["message"]

		self.assertIn("booked until", message)
		self.assertIn("if it has already ended", message)

	def test_an_event_without_a_teams_meeting_is_refused(self):
		event = self._finished_meeting(custom_teams_join_url=None)

		with self.assertRaises(frappe.ValidationError):
			artifacts.fetch_meeting_artifacts(event.name)

	def test_the_meeting_id_is_resolved_once_and_kept(self):
		event = self._finished_meeting(custom_microsoft_online_meeting_id=None)

		with patch.object(
			artifacts, "_resolve_online_meeting_id", return_value="resolved-1"
		) as resolve, patch.object(ms, "list_transcripts", return_value=[]), patch.object(
			ms, "list_recordings", return_value=[]
		):
			artifacts.fetch_meeting_artifacts(event.name)
			event.reload()
			self.assertEqual(event.custom_microsoft_online_meeting_id, "resolved-1")

			artifacts.fetch_meeting_artifacts(event.name)

		resolve.assert_called_once()

	def test_an_unmatchable_join_link_explains_the_expiry(self):
		event = self._finished_meeting(custom_microsoft_online_meeting_id=None)

		with patch.object(artifacts, "_resolve_online_meeting_id", return_value=None):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("60 days", str(ctx.exception))

	def test_a_recording_id_from_another_meeting_is_refused(self):
		"""Otherwise the download endpoint is a way to pull any recording by guessing ids."""
		event = self._finished_meeting()
		event.db_set("custom_microsoft_recordings_data", '[{"id": "r1"}]', update_modified=False)

		with self.assertRaises(frappe.PermissionError):
			artifacts.download_recording(event.name, recording_id="someone-elses")


class TestDownloadingARecording(ArtifactsTestCase):
	"""The bytes pass through; they are never collected on the way."""

	def _recorded(self, *recording_ids):
		event = self._finished_meeting()
		event.db_set(
			"custom_microsoft_recordings_data",
			json.dumps([{"id": r} for r in recording_ids]),
			update_modified=False,
		)
		return event

	def _download(self, event, recording_id=None, **response_kwargs):
		"""Returns the werkzeug Response, the fake Graph response, and the patched call."""
		fake = UnbufferableResponse(**response_kwargs)
		with patch.object(graph, "graph_request", return_value=fake) as requested:
			response = artifacts.download_recording(event.name, recording_id=recording_id)
		return response, fake, requested

	def test_the_recording_is_never_read_into_memory(self):
		"""It used to be resp.content — the whole file resident in a gunicorn worker, and copied
		a second time when Frappe assigned it to response.data."""
		event = self._recorded("r1")

		response, fake, requested = self._download(event)

		self.assertIsInstance(response, Response)
		self.assertTrue(requested.call_args.kwargs["raw"])
		self.assertTrue(requested.call_args.kwargs["stream"])
		self.assertTrue(response.is_streamed)
		self.assertTrue(response.direct_passthrough)
		self.assertEqual(fake.chunk_size, artifacts.RECORDING_CHUNK_BYTES)
		self.assertEqual(b"".join(response.response), b"onetwothree")

	def test_the_download_is_named_for_the_event(self):
		event = self._recorded("r1")

		response, _fake, _requested = self._download(event)

		disposition = response.headers["Content-Disposition"]
		self.assertTrue(disposition.startswith("attachment"), disposition)
		self.assertIn(f"teams-recording-{event.name}.mp4", disposition)
		self.assertEqual(response.headers["Content-Type"], "video/mp4")

	def test_a_part_of_a_long_meeting_is_named_for_its_part(self):
		"""Teams splits at 4 hours or 1.5 GB, so two downloads called the same thing are two
		files the person cannot tell apart."""
		event = self._recorded("r1", "r2")

		response, _fake, requested = self._download(event, recording_id="r2")

		self.assertIn(f"teams-recording-{event.name}-part-2.mp4", response.headers["Content-Disposition"])
		self.assertIn("/recordings/r2/content", requested.call_args.args[1])

	def test_the_first_part_is_sent_when_none_is_named(self):
		event = self._recorded("r1", "r2")

		response, _fake, requested = self._download(event)

		self.assertIn(f"teams-recording-{event.name}-part-1.mp4", response.headers["Content-Disposition"])
		self.assertIn("/recordings/r1/content", requested.call_args.args[1])

	def test_the_length_microsoft_states_is_passed_on(self):
		"""Without it the browser shows an unknown-duration spinner for a gigabyte download."""
		event = self._recorded("r1")

		response, _fake, _requested = self._download(event, headers={"Content-Length": "1610612736"})

		self.assertEqual(response.headers["Content-Length"], "1610612736")

	def test_a_length_is_never_invented(self):
		"""A Content-Length that does not match the body truncates the file the person keeps."""
		event = self._recorded("r1")

		response, _fake, _requested = self._download(event)

		self.assertNotIn("Content-Length", response.headers)

	def test_microsofts_own_content_type_wins(self):
		"""It encoded the file; mp4 is only what Teams has always produced, not a promise."""
		event = self._recorded("r1")

		response, _fake, _requested = self._download(event, headers={"Content-Type": "video/quicktime"})

		self.assertEqual(response.headers["Content-Type"], "video/quicktime")

	def test_the_connection_is_released_when_the_response_closes(self):
		"""A download someone cancels halfway must not pin a socket until the request times out."""
		event = self._recorded("r1")

		response, fake, _requested = self._download(event)
		response.close()

		self.assertTrue(fake.closed)

	def test_a_recording_id_from_another_meeting_never_reaches_microsoft(self):
		"""The ownership check has to happen before the fetch, or the refusal costs a Graph call
		on someone else's recording — which is half of what the guard is there to prevent."""
		event = self._recorded("r1")

		with patch.object(graph, "graph_request") as requested:
			with self.assertRaises(frappe.PermissionError):
				artifacts.download_recording(event.name, recording_id="someone-elses")

		requested.assert_not_called()

	def test_a_meeting_with_no_recording_says_so_instead_of_asking(self):
		event = self._finished_meeting()

		with patch.object(graph, "graph_request") as requested:
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.download_recording(event.name)

		requested.assert_not_called()
		self.assertIn("no recording", str(ctx.exception))

	def test_a_refusal_from_microsoft_mentions_the_expiry(self):
		"""By the time a download 404s, the likeliest cause is the meeting ageing out, and
		"not found" on its own sends people looking in OneDrive for a file that is still there."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._recorded("r1")

		with patch.object(graph, "graph_request", side_effect=MsGraphError("404: itemNotFound")):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.download_recording(event.name)

		self.assertIn("itemNotFound", str(ctx.exception))
		self.assertIn("60 days", str(ctx.exception))


class TestTheTenantSwitch(ArtifactsTestCase):
	"""Microsoft added a tenant control for Graph access to transcripts in 2026, shipped OFF.

	A tenant with every permission consented still gets 403 until an admin turns it on, so the
	old hint — "grant OnlineMeetingTranscript.Read.All with tenant-admin consent" — sent people
	to re-grant consent they already had. Reported from a live tenant.
	"""

	GRAPH_403 = (
		"Microsoft Graph GET /me/onlineMeetings/x/transcripts failed (403): "
		"Forbidden: Graph API access to transcripts is disabled for this tenant."
	)

	def test_the_tenant_switch_is_named_instead_of_consent(self):
		"""Tested where the translation happens — inside the transcripts module, not through a
		mock of the very function that does it."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		with self.assertRaises(frappe.ValidationError) as ctx:
			ms._wrap_403(MsGraphError(self.GRAPH_403), ms._transcript_perm_hint())

		said = str(ctx.exception)
		self.assertIn("Teams admin center", said)
		self.assertIn("Transcript API access", said)
		# The point is not that the word never appears — the message explains that consent is
		# already in place. It is that nobody is told to go and grant anything.
		self.assertNotIn("grant", said.lower(), "sending them to grant consent they have is a loop")

	def test_an_ordinary_403_still_gets_the_permission_hint(self):
		"""The carve-out must not swallow the case it was carved out of."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		with self.assertRaises(frappe.ValidationError) as ctx:
			ms._wrap_403(MsGraphError("Graph GET /x failed (403): Forbidden"), ms._transcript_perm_hint())

		self.assertIn("OnlineMeetingTranscript.Read.All", str(ctx.exception))

	def test_what_microsoft_said_reaches_the_user_unrewritten(self):
		"""Whatever the layers below decided to say, the dialog shows exactly that — once."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._finished_meeting()

		with patch.object(ms, "list_transcripts", side_effect=MsGraphError(self.GRAPH_403)), patch.object(
			ms, "list_recordings", return_value=[]
		):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		said = str(ctx.exception)
		self.assertIn("disabled for this tenant", said)
		self.assertEqual(said.count("disabled for this tenant"), 1, "said once, not stacked")

	def test_the_doctor_recognises_it_too(self):
		from frappe_microsoft365 import doctor

		explained = doctor.explain_error(self.GRAPH_403)

		self.assertTrue(explained["matched"])
		self.assertIn("Transcript API access", explained["detail"])

	def test_a_real_permission_403_still_says_permission(self):
		"""The new pattern must not swallow the case it was carved out of."""
		from frappe_microsoft365 import doctor

		explained = doctor.explain_error("Microsoft Graph GET /me/events failed (403): Forbidden")

		self.assertNotIn("Transcript API access", explained.get("detail") or "")


class TestWhichPermissionIsMissing(ArtifactsTestCase):
	"""Three Microsoft refusals that read identically and are fixed in three different places.

	The report behind this: a tenant that had granted transcript consent was told to grant
	transcript consent, by a 403 that was really about something else. Every message below has
	to send someone somewhere they have not already been.
	"""

	def _forbidden(self, path):
		from frappe_microsoft365.microsoft_graph import MsGraphError

		return MsGraphError(f"Microsoft Graph GET {path} failed (403): Forbidden")

	def test_a_recordings_403_is_not_answered_with_the_transcript_fix(self):
		"""Microsoft consents to reading video separately from reading words: a tenant happy with
		one often refuses the other, so the two 403s are different problems with different fixes."""
		with patch.object(
			graph, "graph_request", side_effect=self._forbidden("/me/onlineMeetings/m1/recordings")
		):
			with self.assertRaises(frappe.ValidationError) as ctx:
				ms.list_recordings(CALENDAR, "m1")

		said = str(ctx.exception)
		self.assertIn("OnlineMeetingRecording.Read.All", said)
		self.assertNotIn("OnlineMeetingTranscript.Read.All", said)
		self.assertIn("API permissions", said, "a permission nobody can find is not a fix")

	def test_a_transcripts_403_is_not_answered_with_the_recording_fix(self):
		with patch.object(
			graph, "graph_request", side_effect=self._forbidden("/me/onlineMeetings/m1/transcripts")
		):
			with self.assertRaises(frappe.ValidationError) as ctx:
				ms.list_transcripts(CALENDAR, "m1")

		said = str(ctx.exception)
		self.assertIn("OnlineMeetingTranscript.Read.All", said)
		self.assertNotIn("OnlineMeetingRecording.Read.All", said)

	def test_the_transcript_and_recording_advice_are_not_the_same_sentence(self):
		"""Both reach the user through one wrapper, and one shared hint is exactly how two causes
		became one indistinguishable message."""
		self.assertNotEqual(ms._transcript_perm_hint(), ms._recording_perm_hint())

	def test_a_403_resolving_the_join_link_names_the_meetings_permission(self):
		"""Mapping a join link to a meeting id is /me/onlineMeetings, which needs
		OnlineMeetings.ReadWrite — a permission neither transcript nor recording consent
		includes. People tick transcripts, grant that, and stop, so this is the step that fails
		for them; it used to surface as a raw Graph 403 against a path nobody recognises."""
		event = self._finished_meeting(custom_microsoft_online_meeting_id=None)

		with patch.object(
			artifacts,
			"_resolve_online_meeting_id",
			side_effect=self._forbidden("/me/onlineMeetings?$filter=JoinWebUrl eq 'x'"),
		):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		said = str(ctx.exception)
		self.assertIn("OnlineMeetings.ReadWrite", said)
		self.assertIn("Standalone Teams meetings", said, "name the tickbox that asks for it")
		self.assertNotIn("OnlineMeetingTranscript.Read.All", said)

	def test_a_join_link_that_matches_nothing_is_not_called_a_permission_problem(self):
		"""An empty answer is not a refusal. Microsoft never saw this meeting — most often
		because another organisation hosted it — and no permission will change that."""
		event = self._finished_meeting(custom_microsoft_online_meeting_id=None)

		with patch.object(artifacts, "_resolve_online_meeting_id", return_value=None):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		said = str(ctx.exception)
		self.assertIn("another organisation", said)
		self.assertNotIn("OnlineMeetings.ReadWrite", said)

	def test_a_403_downloading_a_recording_is_consent_not_expiry(self):
		"""Listing a recording and reading its bytes take the same permission, so a 403 that
		appears only at download time is consent that was never granted. The expiry note sent
		someone hunting through OneDrive for a file Microsoft was refusing, not missing."""
		event = self._finished_meeting()
		event.db_set("custom_microsoft_recordings_data", '[{"id": "r1"}]', update_modified=False)

		with patch.object(
			graph, "graph_request", side_effect=self._forbidden("/recordings/r1/content")
		):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.download_recording(event.name)

		said = str(ctx.exception)
		self.assertIn("OnlineMeetingRecording.Read.All", said)
		self.assertNotIn("60 days", said)


class TestTheConnectionItself(ArtifactsTestCase):
	"""Two states of one record that used to share a sentence, and share no fix."""

	def test_a_switched_off_calendar_names_the_tickbox(self):
		"""Saying "disabled or not authorized" made the reader check both, one of which was fine."""
		event = self._finished_meeting()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "enabled", 0)
		self.addCleanup(frappe.db.set_value, "Microsoft Calendar", CALENDAR, "enabled", 1)

		with self.assertRaises(frappe.ValidationError) as ctx:
			artifacts.fetch_meeting_artifacts(event.name)

		said = str(ctx.exception)
		self.assertIn("tick Enabled", said)
		self.assertNotIn("Authorize", said)

	def test_a_calendar_nobody_signed_into_names_the_button(self):
		event = self._finished_meeting()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "authorized", 0)
		self.addCleanup(frappe.db.set_value, "Microsoft Calendar", CALENDAR, "authorized", 1)

		with self.assertRaises(frappe.ValidationError) as ctx:
			artifacts.fetch_meeting_artifacts(event.name)

		said = str(ctx.exception)
		self.assertIn("Authorize Microsoft Access", said)
		self.assertNotIn("tick Enabled", said)


#: A dialog is read standing up, once. Two sentences of cause and action land; a paragraph is
#: skipped, which costs more than saying nothing at all. The longest message here is the
#: nothing-found one at a little over 300 characters, so this ceiling leaves room to rephrase
#: and none to append.
MAX_MESSAGE_CHARS = 340


class TestTheMessagesStayReadable(ArtifactsTestCase):
	"""Advice is only advice if someone reads it to the end."""

	def _state(self, kind, **overrides):
		"""A state dict by hand: describe_state is pure, and building every branch through the
		database would test the state machine again rather than the sentences."""
		state = {
			"state": kind,
			"has_transcript": False,
			"recordings": 0,
			"attempts": 0,
			"expires_on": None,
			"next_check": None,
			"auto": False,
			"still_booked": False,
			"booked_until": None,
		}
		state.update(overrides)
		return state

	def _every_message(self):
		later = add_to_date(now_datetime(), hours=2)
		described = {
			"not started": self._state("not_started"),
			"complete, both": self._state("complete", has_transcript=True, recordings=2),
			"complete, transcript only": self._state("complete", has_transcript=True),
			"complete, recording only": self._state("complete", recordings=1),
			"partial, still trying": self._state(
				"partial", has_transcript=True, next_check=later, auto=True
			),
			"partial, given up": self._state("partial", has_transcript=True),
			"processing, inside the slot": self._state(
				"processing", still_booked=True, booked_until=later, next_check=later, auto=True
			),
			"processing": self._state("processing", next_check=later, auto=True),
			"processing, nothing scheduled": self._state("processing", next_check=later),
			"nothing found": self._state("nothing_found"),
			"expired": self._state("expired"),
			"no meeting": self._state("no_meeting"),
		}
		messages = {name: artifacts.describe_state(state) for name, state in described.items()}
		messages.update(
			{
				"transcript permission": ms._transcript_perm_hint(),
				"recording permission": ms._recording_perm_hint(),
				"online meetings permission": ms._online_meetings_hint(),
				"tenant switch": ms._tenant_switch_hint(),
				"expiry note": artifacts._expiry_note(),
				"teams tab note": artifacts._teams_tab_note(),
				"processing note": artifacts._processing_note(),
			}
		)
		return messages

	def test_no_message_grows_into_a_wall_of_text(self):
		"""Every one of these was added to answer a real support question, and the next one will
		be too. The ceiling is what stops the answers accumulating into a paragraph nobody finishes.

		Graph's own text is excluded on purpose: this bounds what we wrote, not what Microsoft said.
		"""
		for name, message in self._every_message().items():
			with self.subTest(message=name):
				self.assertTrue(message.strip(), "a state with no sentence explains nothing")
				self.assertLessEqual(
					len(message),
					MAX_MESSAGE_CHARS,
					f"{name!r} is {len(message)} characters: say less, or say it elsewhere",
				)


class TestAFlappingTenantSwitch(ArtifactsTestCase):
	"""A tenant that has just had Graph access to transcripts switched on answers
	inconsistently for a while: one call 403s, the next returns an empty list.

	Reported live — the button said "Microsoft has not finished processing this meeting" and
	sixty seconds later the same call refused with the tenant-switch message. Reporting that
	empty list as processing sends someone off to wait for Microsoft when what they are waiting
	for is their own setting reaching every server.
	"""

	def _refused_once(self):
		event = self._finished_meeting()
		event.db_set(
			"custom_microsoft_artifacts_status",
			ms._tenant_switch_hint(),
			update_modified=False,
		)
		event.reload()
		return event

	def test_an_empty_answer_after_a_refusal_does_not_claim_processing(self):
		event = self._refused_once()

		message = self._fetch(event)["message"]

		self.assertIn("just switched it on", message)
		self.assertNotIn("has not finished processing", message)

	def test_the_memory_survives_the_next_attempt(self):
		"""The status field holds one message, so a memory that does not survive being rewritten
		is not a memory — the sentence has to recognise itself."""
		event = self._refused_once()

		self._fetch(event)
		second = self._fetch(event)["message"]

		self.assertIn("just switched it on", second)

	def test_it_says_the_empty_answer_proves_nothing(self):
		"""The honest part: we cannot tell 'no transcript' from 'setting not live yet'."""
		event = self._refused_once()

		self.assertIn("not evidence", self._fetch(event)["message"])

	def test_a_meeting_that_never_hit_the_switch_gets_the_ordinary_message(self):
		event = self._finished_meeting()

		message = self._fetch(event)["message"]

		self.assertNotIn("just switched it on", message)


class TestSomebodyElsesFileHook(ArtifactsTestCase):
	"""Saving a File runs every after_insert hook any installed app put on File.

	Seen live: frappe_s3_attachment was installed without AWS credentials, so the transcript —
	which Microsoft had handed over perfectly — died in botocore on the way to disk, and Frappe
	reported the failure as coming from this app. Hours of Microsoft debugging were spent on a
	missing AWS key.
	"""

	def test_a_storage_failure_is_not_reported_as_a_microsoft_failure(self):
		event = self._finished_meeting()

		# Patched where the hook actually blows up — inside save_file — not on the function that
		# carries the guard, which would replace the thing under test with the mock.
		with patch(
			"frappe.utils.file_manager.save_file",
			side_effect=RuntimeError("Unable to locate credentials"),
		):
			with self.assertRaises(frappe.ValidationError) as ctx:
				self._fetch(event, transcripts=[{"id": "t1", "created_date_time": "2026-09-13T11:00:00Z"}])

		said = str(ctx.exception)
		self.assertIn("could not store the file", said)
		self.assertIn("not a Microsoft one", said)
		self.assertIn("Unable to locate credentials", said, "the real cause must survive")

	def test_the_transcript_being_fetched_is_stated_not_implied(self):
		"""The one fact that stops somebody debugging the wrong system."""
		event = self._finished_meeting()

		with patch("frappe.utils.file_manager.save_file", side_effect=RuntimeError("disk full")):
			with self.assertRaises(frappe.ValidationError) as ctx:
				self._fetch(event, transcripts=[{"id": "t1", "created_date_time": "2026-09-13T11:00:00Z"}])

		self.assertIn("Microsoft gave us the transcript", str(ctx.exception))
