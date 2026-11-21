"""Clearing the pre-1.0.0 scope list keeps the capabilities it covered (#46)."""

import frappe

from frappe_microsoft365 import doctor
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.patches import clear_legacy_default_scopes as patch_module
from frappe_microsoft365.tests.base import BaseTestCase

LEGACY = "User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite OnlineMeetingTranscript.Read.All"
FIELDS = ("enabled", "use_calendar", "use_teams", "use_transcripts", "default_scopes")


class LegacyScopesTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		before = {f: frappe.db.get_single_value("Microsoft Settings", f) for f in FIELDS}
		self.addCleanup(self._settings, **before)
		for name in frappe.get_all("Microsoft Calendar", pluck="name"):
			frappe.delete_doc("Microsoft Calendar", name, force=True, ignore_permissions=True)

	def _settings(self, **values):
		for field, value in values.items():
			frappe.db.set_single_value("Microsoft Settings", field, value)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")

	def _get(self, field):
		return frappe.db.get_single_value("Microsoft Settings", field)

	def _calendar(self, authorized=1):
		return frappe.get_doc(
			{
				"doctype": "Microsoft Calendar",
				"account_name": f"Legacy {frappe.generate_hash(length=6)}",
				"user": "Administrator",
				"enabled": 1,
				"authorized": authorized,
			}
		).insert(ignore_permissions=True)


class TestPatch(LegacyScopesTestCase):
	def test_the_legacy_value_keeps_teams_and_transcripts_on(self):
		self._calendar()
		self._settings(enabled=1, use_calendar=1, use_teams=0, use_transcripts=0, default_scopes=LEGACY)

		patch_module.execute()
		patch_module.execute()

		self.assertEqual(self._get("default_scopes") or "", "")
		self.assertEqual((self._get("use_teams"), self._get("use_transcripts")), (1, 1))
		# The old value's scopes are still asked for.
		self.assertTrue(set(LEGACY.split()) <= set(graph.get_scopes()))

	def test_nothing_is_ticked_where_nobody_ever_authorised(self):
		self._calendar(authorized=0)
		self._settings(use_teams=0, use_transcripts=0, default_scopes=LEGACY)

		patch_module.execute()

		self.assertEqual(self._get("default_scopes") or "", "")
		self.assertEqual((self._get("use_teams"), self._get("use_transcripts")), (0, 0))

	def test_an_edited_value_is_an_override_and_left_alone(self):
		self._calendar()
		self._settings(use_teams=0, use_transcripts=0, default_scopes="User.Read Calendars.ReadWrite")

		patch_module.execute()

		self.assertEqual(self._get("default_scopes"), "User.Read Calendars.ReadWrite")
		self.assertEqual((self._get("use_teams"), self._get("use_transcripts")), (0, 0))

	def test_an_empty_value_is_a_no_op(self):
		self._calendar()
		self._settings(use_teams=0, use_transcripts=0, default_scopes="")

		patch_module.execute()

		self.assertEqual((self._get("use_teams"), self._get("use_transcripts")), (0, 0))


class TestUpgradeCheck(LegacyScopesTestCase):
	def setUp(self):
		super().setUp()
		from frappe_microsoft365.setup import create_event_custom_fields

		create_event_custom_fields()

	def _event(self, **values):
		doc = frappe.get_doc(
			{
				"doctype": "Event",
				"subject": "Legacy meeting",
				"starts_on": "2026-11-20 10:00:00",
				"event_type": "Private",
			}
		).insert(ignore_permissions=True)
		if values:
			frappe.db.set_value("Event", doc.name, values, update_modified=False)
		return doc

	def _columns(self):
		return all(
			frappe.db.has_column("Event", c)
			for c in ("custom_microsoft_online_meeting_id", "custom_microsoft_transcript_fetched_on")
		)

	def test_meeting_data_with_the_capabilities_off_is_reported(self):
		if not self._columns():
			self.skipTest("meeting fields are not on Event on this site")
		self._event(
			custom_microsoft_online_meeting_id="MID-1",
			custom_microsoft_transcript_fetched_on="2026-11-20 11:00:00",
		)
		self._settings(use_teams=0, use_transcripts=0)

		found = patch_module.upgrade_check_meeting_capabilities_off()

		self.assertEqual(len(found), 2)
		self.assertTrue(
			all(f["check"] == "upgrade.meeting_capability_off" and f["status"] == doctor.WARN for f in found)
		)

	def test_quiet_when_the_capabilities_are_on_or_never_used(self):
		if not self._columns():
			self.skipTest("meeting fields are not on Event on this site")
		self._settings(use_teams=0, use_transcripts=0)
		frappe.db.sql(
			"update `tabEvent` set custom_microsoft_online_meeting_id=null, custom_microsoft_transcript_fetched_on=null"
		)
		self.assertEqual(patch_module.upgrade_check_meeting_capabilities_off(), [])

		self._event(
			custom_microsoft_online_meeting_id="MID-1",
			custom_microsoft_transcript_fetched_on="2026-11-20 11:00:00",
		)
		self._settings(use_teams=1, use_transcripts=1)
		self.assertEqual(patch_module.upgrade_check_meeting_capabilities_off(), [])
