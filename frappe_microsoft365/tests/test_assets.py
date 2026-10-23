"""Desk scripts reach browsers after a deploy."""

import glob
import os
import re

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

	def test_desk_scripts_only_call_what_frappe_makes_global(self):
		# Frappe's desk bundle wraps its files in one closure, so a top-level helper such as
		# prettyDate is only reachable where Frappe hands it out (window.X, or frappe.datetime.X).
		# A bare call works when the file is read on its own, and throws in the desk.
		private = frappe_private_functions()
		self.assertIn("prettyDate", private)
		public_js = frappe.get_app_path("frappe_microsoft365", "public", "js")
		for script in hooks.app_include_js:
			with open(os.path.join(public_js, script)) as f:
				source = f.read()
			code = re.sub(r"//.*", "", source)
			own = set(re.findall(r"\bfunction (\w+)\s*\(", code))
			own |= set(re.findall(r"\b(?:const|let|var) (\w+)\s*=", code))
			own |= set(re.findall(r"^\s*(?:async\s+)?(\w+)\s*\([^)]*\)\s*\{", code, re.M))  # methods
			for name in sorted(private - own):
				self.assertIsNone(re.search(rf"(?<![\w.$]){name}\s*\(", code), f"{script}: {name}")


def frappe_private_functions():
	"""Top-level functions in Frappe's desk scripts that Frappe never puts on window."""
	declared, exported = set(), set()
	for path in glob.glob(os.path.join(frappe.get_app_path("frappe", "public", "js"), "**", "*.js"), recursive=True):
		if "node_modules" in path:
			continue
		with open(path, errors="ignore") as f:
			source = f.read()
		declared |= set(re.findall(r"^function (\w+)\s*\(", source, re.M))
		exported |= set(re.findall(r"\bwindow\.(\w+)\s*=", source))
		for block in re.findall(r"Object\.assign\(window,\s*\{(.*?)\}\)", source, re.S):
			exported |= set(re.findall(r"(\w+)\s*[,:\n]", block))
	return declared - exported
