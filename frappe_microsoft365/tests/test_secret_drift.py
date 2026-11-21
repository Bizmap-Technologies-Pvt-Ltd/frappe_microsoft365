"""A rotated client secret is noticed wherever a copy of the old one was left behind (#44).

Set Up copies the secret in Microsoft Settings into the mail Connected App and the sign-in
Social Login Key. Rotating it in Settings alone used to break mail and sign-in with nothing on
the doctor or the plan to say why.
"""

import json

import frappe
from frappe.utils.password import get_decrypted_password, set_encrypted_password

from frappe_microsoft365 import doctor, provisioning
from frappe_microsoft365.tests.base import BaseTestCase

TENANT = "11111111-2222-3333-4444-555555555555"
CLIENT = "aaaaaaaa-0000-0000-0000-000000000001"
APP_CLIENT = "bbbbbbbb-0000-0000-0000-000000000002"
OLD, NEW = "OLD-secret-value~abc", "NEW-secret-value~xyz"


class TestSecretDrift(BaseTestCase):
	def setUp(self):
		super().setUp()
		frappe.set_user("Administrator")
		fields = (
			"tenant_id",
			"client_id",
			"use_sso",
			"use_mail",
			"use_calendar",
			"mail_flow",
			"app_client_id",
		)
		before = {f: frappe.db.get_single_value("Microsoft Settings", f) for f in fields}
		self.addCleanup(self._settings, **before)
		self._clear()
		self.addCleanup(self._clear)
		self._settings(
			tenant_id=TENANT, client_id=CLIENT, use_sso=1, use_mail=1, mail_flow="Delegated", app_client_id=""
		)
		self._secret(OLD)
		created = provisioning.apply()["created"]
		self.assertEqual({c["capability"] for c in created}, {"mail", "sso"})
		self.app = frappe.db.get_value("Connected App", {"provider_name": provisioning.MAIL_APP_NAME})
		self.key = frappe.db.get_value("Social Login Key", {"social_login_provider": "Office 365"})

	def _settings(self, **values):
		for field, value in values.items():
			frappe.db.set_single_value("Microsoft Settings", field, value)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")

	def _secret(self, value, field="client_secret"):
		set_encrypted_password("Microsoft Settings", "Microsoft Settings", value, field)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")

	def _clear(self):
		for name in frappe.get_all("Social Login Key", {"social_login_provider": "Office 365"}, pluck="name"):
			frappe.delete_doc("Social Login Key", name, force=True, ignore_permissions=True)
		names = set(
			frappe.get_all("Connected App", {"client_id": ["in", [CLIENT, APP_CLIENT]]}, pluck="name")
		)
		names.update(
			frappe.get_all("Connected App", {"provider_name": provisioning.MAIL_APP_NAME}, pluck="name")
		)
		for name in names:
			frappe.delete_doc("Connected App", name, force=True, ignore_permissions=True)

	def _drift(self):
		return [f for f in doctor.run_diagnostics()["findings"] if f["check"] == "secret.drift"]

	def test_matching_copies_are_not_reported(self):
		self.assertEqual(self._drift(), [])
		self.assertEqual(provisioning.plan()["secret_drift"], [])

	def test_a_rotated_secret_is_reported_on_both_copies(self):
		self._secret(NEW)

		found = self._drift()

		self.assertEqual({f["target"] for f in found}, {self.app, self.key})
		self.assertTrue(all(f["status"] == doctor.WARN for f in found))
		self.assertTrue(all("Sync Secret" in f["fix"] for f in found))

	def test_no_secret_ever_leaves_the_server(self):
		self._secret(NEW)

		payload = json.dumps([doctor.run_diagnostics(), provisioning.plan()], default=str)

		self.assertNotIn(OLD, payload)
		self.assertNotIn(NEW, payload)

	def test_the_plan_marks_the_records_and_offers_the_sync(self):
		self._secret(NEW)

		result = provisioning.plan()

		self.assertEqual(
			sorted(result["secret_drift"], key=lambda d: d["doctype"]),
			[
				{"doctype": "Connected App", "name": self.app},
				{"doctype": "Social Login Key", "name": self.key},
			],
		)
		steps = {s["capability"]: s for s in result["steps"]}
		for capability in ("mail", "sso"):
			self.assertEqual(steps[capability]["action"], provisioning.DRIFT)
			self.assertIn("secret.drift", {f["check"] for f in steps[capability]["findings"]})

	def test_sync_copies_the_current_secret_to_both(self):
		self._secret(NEW)

		result = provisioning.sync_client_secret()

		self.assertEqual(len(result["updated"]), 2)
		self.assertNotIn(NEW, json.dumps(result))
		self.assertEqual(get_decrypted_password("Connected App", self.app, "client_secret"), NEW)
		self.assertEqual(get_decrypted_password("Social Login Key", self.key, "client_secret"), NEW)
		self.assertEqual(self._drift(), [])
		# Nothing left to do: a second sync changes nothing.
		self.assertEqual(provisioning.sync_client_secret()["updated"], [])

	def test_a_record_for_another_registration_is_not_compared_or_touched(self):
		frappe.db.set_value("Connected App", self.app, "client_id", "cccccccc-0000-0000-0000-000000000003")
		self._secret(NEW)

		self.assertEqual({f["target"] for f in self._drift()}, {self.key})
		provisioning.sync_client_secret()
		self.assertEqual(get_decrypted_password("Connected App", self.app, "client_secret"), OLD)

	def test_application_access_copies_compare_with_its_own_secret(self):
		self._settings(app_client_id=APP_CLIENT)
		self._secret("APP-secret", "app_client_secret")
		app_only = frappe.get_doc(
			{
				"doctype": "Connected App",
				"provider_name": "Microsoft 365 App Mail",
				"client_id": APP_CLIENT,
				"client_secret": "APP-secret",
				"authorization_uri": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
				"token_uri": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
			}
		).insert(ignore_permissions=True)
		self.assertEqual(self._drift(), [])

		self._secret("APP-secret-2", "app_client_secret")

		self.assertEqual({f["target"] for f in self._drift()}, {app_only.name})
		provisioning.sync_client_secret()
		self.assertEqual(
			get_decrypted_password("Connected App", app_only.name, "client_secret"), "APP-secret-2"
		)

	def test_sync_is_for_system_managers_only(self):
		self._secret(NEW)
		user = "secret.drift.guest@example.com"
		if not frappe.db.exists("User", user):
			frappe.get_doc(
				{"doctype": "User", "email": user, "first_name": "Drift", "send_welcome_email": 0}
			).insert(ignore_permissions=True)
		frappe.set_user(user)
		try:
			with self.assertRaises(frappe.PermissionError):
				provisioning.sync_client_secret()
		finally:
			frappe.set_user("Administrator")
		self.assertEqual(get_decrypted_password("Connected App", self.app, "client_secret"), OLD)

	def test_quiet_without_any_copy(self):
		"""Production: no Connected App or Social Login Key made from these settings."""
		self._clear()
		self._secret(NEW)

		self.assertEqual(self._drift(), [])
		self.assertEqual(doctor._secret_copies(), [])
