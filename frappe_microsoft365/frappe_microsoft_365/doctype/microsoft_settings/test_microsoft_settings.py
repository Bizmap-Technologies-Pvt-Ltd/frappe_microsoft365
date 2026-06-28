"""Smoke tests — assert the app installs cleanly and its core surface is present."""

import frappe
from frappe.tests.utils import FrappeTestCase


class TestMicrosoftSettings(FrappeTestCase):
	def test_doctypes_exist(self):
		for dt in ("Microsoft Settings", "Microsoft Calendar"):
			self.assertTrue(frappe.db.exists("DocType", dt), f"{dt} should be installed")

	def test_settings_is_single(self):
		self.assertTrue(frappe.get_meta("Microsoft Settings").issingle)

	def test_core_modules_import(self):
		from frappe_microsoft365 import (  # noqa: F401
			microsoft_graph,
			microsoft_calendar_sync,
			microsoft_meetings,
			microsoft_transcripts,
			permissions,
		)

		self.assertTrue(hasattr(microsoft_graph, "graph_request"))
		self.assertTrue(hasattr(microsoft_meetings, "create_meeting"))

	def test_redirect_uri_built(self):
		from frappe_microsoft365.microsoft_graph import get_redirect_uri

		uri = get_redirect_uri()
		self.assertIn("microsoft_calendar.callback", uri)
