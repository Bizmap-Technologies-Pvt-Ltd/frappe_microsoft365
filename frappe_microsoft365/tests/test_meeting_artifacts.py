"""Transcript and recording tests. Graph is mocked; no tenant, no network."""

import json
from unittest.mock import patch

import frappe
from frappe.utils import add_to_date, now_datetime

from frappe_microsoft365 import microsoft_meeting_artifacts as artifacts
from frappe_microsoft365 import microsoft_transcripts as ms
from frappe_microsoft365.tests.base import BaseTestCase

CALENDAR = "_Test Artifacts Calendar"
JOIN_URL = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0"
VTT = "WEBVTT\n\n00:00:01.000 --> 00:00:04.000\nPriya: Let's begin.\n"
VTT_TWO = "WEBVTT\n\n01:00:01.000 --> 01:00:04.000\nPriya: Back after the break.\n"


class ArtifactsTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
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

	def test_a_day_later_it_says_it_was_never_recorded(self):
		event = self._finished_meeting(ended_hours_ago=30)
		event.db_set("custom_microsoft_artifacts_attempts", len(artifacts.RETRY_MINUTES), update_modified=False)
		event.reload()

		result = self._fetch(event)

		self.assertEqual(result["state"], "nothing_found")
		self.assertIn("never recorded", result["message"].lower())

	def test_after_sixty_days_it_says_the_meeting_is_too_old(self):
		event = self._finished_meeting(ended_hours_ago=24 * (artifacts.MEETING_EXPIRY_DAYS + 5))

		result = self._fetch(event)

		self.assertEqual(result["state"], "expired")
		self.assertIn("too old", result["message"].lower())

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
	def test_a_meeting_that_has_not_finished_is_refused(self):
		event = self._finished_meeting(
			starts_on=add_to_date(now_datetime(), hours=1),
			ends_on=add_to_date(now_datetime(), hours=2),
		)

		with self.assertRaises(frappe.ValidationError) as ctx:
			artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("not finished", str(ctx.exception))

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
