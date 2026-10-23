"""Desk scripts reach browsers after a deploy."""

import os

import frappe

from frappe_microsoft365 import hooks
from frappe_microsoft365.tests.base import BaseTestCase


class TestDeskScripts(BaseTestCase):
	def test_desk_scripts_are_bundles(self):
		# A plain /assets path keeps its name across builds, and nginx lets browsers cache
		# /assets for a year: a deploy would never reach anyone who had the page open before.
		for script in hooks.app_include_js:
			self.assertFalse(script.startswith("/assets"), script)
			self.assertIn(".bundle.", script)

	def test_every_bundle_has_its_source(self):
		public_js = frappe.get_app_path("frappe_microsoft365", "public", "js")
		for script in hooks.app_include_js:
			self.assertTrue(os.path.exists(os.path.join(public_js, script)), script)
