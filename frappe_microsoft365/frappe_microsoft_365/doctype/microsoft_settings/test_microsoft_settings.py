"""Smoke tests — assert the app installs cleanly and its core surface is present."""

import frappe

from frappe_microsoft365.tests.base import BaseTestCase


class TestMicrosoftSettings(BaseTestCase):
	def test_doctypes_exist(self):
		for dt in ("Microsoft Settings", "Microsoft Calendar"):
			self.assertTrue(frappe.db.exists("DocType", dt), f"{dt} should be installed")

	def test_settings_is_single(self):
		self.assertTrue(frappe.get_meta("Microsoft Settings").issingle)

	def test_core_modules_import(self):
		from frappe_microsoft365 import (
			microsoft_calendar_sync,
			microsoft_graph,
			microsoft_meetings,
			microsoft_transcripts,
			permissions,
		)

		self.assertTrue(hasattr(microsoft_graph, "graph_request"))
		self.assertTrue(hasattr(microsoft_meetings, "create_meeting"))

	def test_redirect_uri_built(self):
		"""With no override configured, the callback URI is derived from the site URL."""
		from frappe_microsoft365.microsoft_graph import get_redirect_uri

		uri = get_redirect_uri(frappe._dict(redirect_uri=None))
		self.assertIn("microsoft_calendar.callback", uri)

	def test_configured_redirect_uri_wins(self):
		from frappe_microsoft365.microsoft_graph import get_redirect_uri

		uri = get_redirect_uri(frappe._dict(redirect_uri="https://example.com/cb"))
		self.assertEqual(uri, "https://example.com/cb")

	def test_transcripts_no_longer_depend_on_standalone_meetings(self):
		"""Transcripts find their meeting with OnlineMeetings.Read, so unticking standalone
		meetings must leave them on — and stop asking for ReadWrite on a site never granted it."""
		from frappe_microsoft365.microsoft_graph import derive_scopes

		settings = frappe.get_single("Microsoft Settings")
		self.addCleanup(frappe.db.rollback)
		settings.update({"enabled": 0, "use_teams": 1, "use_transcripts": 1})
		settings.save(ignore_permissions=True)

		settings.use_teams = 0
		settings.save(ignore_permissions=True)

		self.assertTrue(settings.use_transcripts)
		scopes = derive_scopes(settings)
		self.assertIn("OnlineMeetings.Read", scopes)
		self.assertNotIn("OnlineMeetings.ReadWrite", scopes)
