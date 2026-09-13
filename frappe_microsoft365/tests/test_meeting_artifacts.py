"""Transcript and recording tests. Graph is mocked; no tenant, no network."""

from unittest.mock import patch

import frappe
from frappe.utils import add_to_date, now_datetime

from frappe_microsoft365 import microsoft_meeting_artifacts as artifacts
from frappe_microsoft365 import microsoft_transcripts as ms
from frappe_microsoft365.tests.base import BaseTestCase

CALENDAR = "_Test Artifacts Calendar"
JOIN_URL = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0"
VTT = "WEBVTT\n\n00:00:01.000 --> 00:00:04.000\nPriya: Let's begin.\n"


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

	def _finished_meeting(self, **overrides):
		values = {
			"doctype": "Event",
			"subject": "Retro",
			"starts_on": add_to_date(now_datetime(), hours=-2),
			"ends_on": add_to_date(now_datetime(), hours=-1),
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


class TestFetchArtifacts(ArtifactsTestCase):
	def test_the_transcript_is_attached_to_the_event(self):
		event = self._finished_meeting()

		with patch.object(
			ms, "list_transcripts", return_value=[{"id": "t1", "created_date_time": "2026-09-20T11:00:00Z"}]
		), patch.object(ms, "get_transcript_content", return_value={"content": VTT}), patch.object(
			ms, "list_recordings", return_value=[]
		):
			result = artifacts.fetch_meeting_artifacts(event.name)

		self.assertTrue(result["transcript"])
		files = frappe.get_all(
			"File",
			filters={"attached_to_doctype": "Event", "attached_to_name": event.name},
			fields=["file_name"],
		)
		self.assertTrue(any(f.file_name.endswith(".vtt") for f in files))
		event.reload()
		self.assertTrue(event.custom_microsoft_transcript_fetched_on)

	def test_fetching_twice_does_not_leave_two_copies(self):
		event = self._finished_meeting()
		calls = dict(
			list_transcripts=[{"id": "t1", "created_date_time": "2026-09-20T11:00:00Z"}],
			get_transcript_content={"content": VTT},
			list_recordings=[],
		)

		for _ in range(2):
			with patch.object(ms, "list_transcripts", return_value=calls["list_transcripts"]), patch.object(
				ms, "get_transcript_content", return_value=calls["get_transcript_content"]
			), patch.object(ms, "list_recordings", return_value=calls["list_recordings"]):
				artifacts.fetch_meeting_artifacts(event.name)

		files = frappe.get_all(
			"File", filters={"attached_to_doctype": "Event", "attached_to_name": event.name}
		)
		self.assertEqual(len(files), 1)

	def test_recordings_are_noted_but_never_downloaded(self):
		"""A Teams recording is hundreds of megabytes; it stays with Microsoft."""
		event = self._finished_meeting()

		with patch.object(ms, "list_transcripts", return_value=[]), patch.object(
			ms, "list_recordings", return_value=[{"id": "r1", "created_date_time": "2026-09-20T11:00:00Z"}]
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

	def test_nothing_found_says_so_and_mentions_retention(self):
		event = self._finished_meeting()

		with patch.object(ms, "list_transcripts", return_value=[]), patch.object(
			ms, "list_recordings", return_value=[]
		):
			result = artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("retention", result["message"].lower())

	def test_a_missing_recording_permission_still_lands_the_transcript(self):
		"""Recordings need their own permission and licence; that must not cost the transcript."""
		from frappe_microsoft365.microsoft_graph import MsGraphError

		event = self._finished_meeting()

		with patch.object(
			ms, "list_transcripts", return_value=[{"id": "t1", "created_date_time": "2026-09-20T11:00:00Z"}]
		), patch.object(ms, "get_transcript_content", return_value={"content": VTT}), patch.object(
			ms, "list_recordings", side_effect=MsGraphError("403 Forbidden")
		):
			result = artifacts.fetch_meeting_artifacts(event.name)

		self.assertTrue(result["transcript"])
		self.assertEqual(result["recordings"], 0)


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

	def test_an_unmatchable_join_link_explains_retention(self):
		event = self._finished_meeting(custom_microsoft_online_meeting_id=None)

		with patch.object(artifacts, "_resolve_online_meeting_id", return_value=None):
			with self.assertRaises(frappe.ValidationError) as ctx:
				artifacts.fetch_meeting_artifacts(event.name)

		self.assertIn("retention", str(ctx.exception).lower())
