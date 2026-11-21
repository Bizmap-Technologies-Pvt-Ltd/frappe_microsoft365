"""Shared test base: works on both Frappe v15 and v16 test runners."""

try:  # Frappe v16+
	from frappe.tests import IntegrationTestCase as _FrappeTestCase
except ImportError:  # Frappe v15
	from frappe.tests.utils import FrappeTestCase as _FrappeTestCase

__all__ = ["BaseTestCase"]


class BaseTestCase(_FrappeTestCase):
	def setUp(self):
		super().setUp()
		# Tokens and throttling pauses are request-local, and a test run is one long request:
		# a 429 answered in one test would otherwise pause that calendar for every test after it.
		from frappe_microsoft365 import microsoft_graph

		microsoft_graph.clear_token_cache()
		self.addCleanup(microsoft_graph.clear_token_cache)
