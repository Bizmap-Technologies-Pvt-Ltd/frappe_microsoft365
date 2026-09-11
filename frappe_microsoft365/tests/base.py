"""Shared test base: works on both Frappe v15 and v16 test runners."""

try:  # Frappe v16+
	from frappe.tests import IntegrationTestCase as BaseTestCase
except ImportError:  # Frappe v15
	from frappe.tests.utils import FrappeTestCase as BaseTestCase

__all__ = ["BaseTestCase"]
