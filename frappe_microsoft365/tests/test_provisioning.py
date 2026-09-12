"""Provisioning tests.

The value builders are pure, so the endpoints and scope strings that decide whether Microsoft
accepts the configuration are all verified offline. The plan/apply layer is tested for the two
properties that matter: it is modular (nothing is created for a capability you did not ask
for) and it is never destructive (an existing record is reported, never rewritten).
"""

import json

import frappe

from frappe_microsoft365 import doctor, provisioning
from frappe_microsoft365.doctor import APP_ONLY_SCOPE, DELEGATED_SCOPES, OFFLINE_ACCESS
from frappe_microsoft365.tests.base import BaseTestCase

TENANT = "11111111-2222-3333-4444-555555555555"


def configure_settings(**values):
	"""Set Microsoft Settings through the doc so the Password field is encrypted properly."""
	doc = frappe.get_doc("Microsoft Settings")
	doc.update(values)
	doc.save(ignore_permissions=True)
	frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")


def reset_settings():
	configure_settings(
		enabled=0,
		tenant_id=None,
		client_id=None,
		client_secret=None,
		redirect_uri=None,
		use_calendar=1,
		use_mail=0,
		use_sso=0,
		mail_flow="Delegated",
	)


def purge_mail_app():
	"""Clear every match, in case an earlier run left one behind."""
	for name in frappe.get_all(
		"Connected App", filters={"provider_name": provisioning.MAIL_APP_NAME}, pluck="name"
	):
		frappe.delete_doc("Connected App", name, force=True, ignore_permissions=True)


def settings(**overrides):
	values = {
		"enabled": 1,
		"tenant_id": TENANT,
		"client_id": "client-abc",
		"client_secret": "shhh",
		"redirect_uri": "https://site.example.com/api/method/callback",
		"use_calendar": 1,
		"use_mail": 0,
		"use_sso": 0,
		"mail_flow": "Delegated",
	}
	values.update(overrides)
	return values


class TestValueBuilders(BaseTestCase):
	def test_delegated_mail_app_carries_the_documented_scopes(self):
		values = provisioning.connected_app_values(settings(), "Delegated")

		self.assertIn(DELEGATED_SCOPES["imap"], values["scopes"])
		self.assertIn(DELEGATED_SCOPES["smtp"], values["scopes"])
		self.assertIn(OFFLINE_ACCESS, values["scopes"])

	def test_application_mail_app_uses_only_the_default_scope(self):
		values = provisioning.connected_app_values(settings(), "Application")

		self.assertEqual(values["scopes"], [APP_ONLY_SCOPE])

	def test_mail_app_uses_tenant_specific_v2_endpoints(self):
		values = provisioning.connected_app_values(settings(), "Delegated")

		self.assertEqual(values["authorization_uri"], f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize")
		self.assertEqual(values["token_uri"], f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token")

	def test_what_we_build_passes_our_own_doctor(self):
		"""The strongest check available offline: provisioning output must satisfy the rules."""
		for flow, app_only in (("Delegated", False), ("Application", True)):
			values = provisioning.connected_app_values(settings(), flow)
			config = {
				"name": values["provider_name"],
				"client_id": values["client_id"],
				# Frappe fills this in on save; simulate the saved state.
				"redirect_uri": "https://site.example.com/api/method/...callback/x",
				"authorization_uri": values["authorization_uri"],
				"token_uri": values["token_uri"],
				"scopes": values["scopes"],
			}
			with self.subTest(flow=flow):
				found = doctor.check_connected_app(config, settings(), app_only=app_only)
				self.assertEqual([f for f in found if f["status"] in (doctor.FAIL, doctor.WARN)], [])

	def test_sso_key_avoids_frappes_common_v1_defaults(self):
		"""Frappe's built-in Office 365 provider ships /common/ + v1.0, which single-tenant apps reject."""
		values = provisioning.social_login_key_values(settings())

		self.assertNotIn("/common/", values["authorize_url"])
		self.assertIn(TENANT, values["authorize_url"])
		self.assertIn("/oauth2/v2.0/", values["authorize_url"])
		self.assertIn("/oauth2/v2.0/", values["access_token_url"])

	def test_sso_key_does_not_let_strangers_self_register(self):
		self.assertEqual(provisioning.social_login_key_values(settings())["sign_ups"], "Deny")
		self.assertIn("openid", json.loads(provisioning.social_login_key_values(settings())["auth_url_data"])["scope"])

	def test_sso_key_passes_our_own_doctor(self):
		values = provisioning.social_login_key_values(settings())
		key = {
			"name": "Office 365",
			"enable_social_login": values["enable_social_login"],
			"client_id": values["client_id"],
			"authorize_url": values["authorize_url"],
			"access_token_url": values["access_token_url"],
		}

		self.assertEqual(doctor.check_social_login_key(key, settings()), [])


class TestCapabilitySelection(BaseTestCase):
	def test_nothing_is_selected_by_default_except_calendar(self):
		self.assertEqual(provisioning.selected_capabilities(settings()), ["calendar"])

	def test_each_capability_is_independent(self):
		self.assertEqual(
			provisioning.selected_capabilities(settings(use_calendar=0, use_sso=1)), ["sso"]
		)
		self.assertEqual(
			provisioning.selected_capabilities(settings(use_calendar=0, use_mail=1)), ["mail"]
		)
		self.assertEqual(
			provisioning.selected_capabilities(settings(use_mail=1, use_sso=1)),
			["calendar", "mail", "sso"],
		)

	def test_every_capability_documents_what_azure_needs(self):
		for capability in provisioning.capabilities():
			with self.subTest(capability=capability["id"]):
				self.assertTrue(capability["label"])
				self.assertTrue(capability["creates"])
				self.assertTrue(capability["azure"])
				self.assertTrue(capability["azure_type"])


class TestPlan(BaseTestCase):
	def setUp(self):
		super().setUp()
		purge_mail_app()
		self.addCleanup(reset_settings)

	def _set(self, **values):
		for field, value in values.items():
			frappe.db.set_single_value("Microsoft Settings", field, value)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")
		self.addCleanup(frappe.clear_document_cache, "Microsoft Settings", "Microsoft Settings")

	def test_unselected_capabilities_are_skipped_not_planned(self):
		self._set(use_calendar=1, use_mail=0, use_sso=0, tenant_id=TENANT, client_id="abc")

		steps = {s["capability"]: s for s in provisioning.plan()["steps"]}

		self.assertEqual(steps["mail"]["action"], provisioning.SKIP)
		self.assertEqual(steps["sso"]["action"], provisioning.SKIP)

	def test_selected_mail_is_planned_for_creation(self):
		self._set(use_mail=1, tenant_id=TENANT, client_id="abc")

		steps = {s["capability"]: s for s in provisioning.plan()["steps"]}

		self.assertEqual(steps["mail"]["action"], provisioning.CREATE)

	def test_missing_credentials_block_the_plan(self):
		self._set(use_mail=1, tenant_id="", client_id="")

		self.assertTrue(provisioning.plan()["blockers"])

	def test_plan_writes_nothing(self):
		self._set(use_mail=1, use_sso=1, tenant_id=TENANT, client_id="abc")
		before = frappe.db.count("Connected App")

		provisioning.plan()

		self.assertEqual(frappe.db.count("Connected App"), before)


class TestApply(BaseTestCase):
	def setUp(self):
		super().setUp()
		purge_mail_app()
		self.addCleanup(purge_mail_app)
		self.addCleanup(reset_settings)
		configure_settings(
			use_calendar=1,
			use_mail=1,
			use_sso=0,
			tenant_id=TENANT,
			client_id="client-abc",
			client_secret="test-secret",
			mail_flow="Delegated",
		)

	def test_creates_the_mail_connected_app_with_correct_scopes(self):

		result = provisioning.apply()

		self.assertEqual([c["capability"] for c in result["created"]], ["mail"])
		name = frappe.db.exists("Connected App", {"provider_name": provisioning.MAIL_APP_NAME})
		self.assertTrue(name)
		doc = frappe.get_doc("Connected App", name)
		scopes = {row.scope for row in doc.scopes}
		self.assertIn(OFFLINE_ACCESS, scopes)
		self.assertIn(DELEGATED_SCOPES["imap"], scopes)
		self.assertIn("/oauth2/v2.0/token", doc.token_uri)

	def test_running_twice_creates_nothing_the_second_time(self):
		provisioning.apply()

		second = provisioning.apply()

		self.assertEqual(second["created"], [])

	def test_an_existing_record_is_reported_never_rewritten(self):
		"""A mail setup someone tuned by hand must survive provisioning untouched."""
		hand_tuned = frappe.get_doc(
			{
				"doctype": "Connected App",
				"provider_name": provisioning.MAIL_APP_NAME,
				"client_id": "someone-elses-id",
				"authorization_uri": "https://login.microsoftonline.com/common/oauth2/authorize",
				"token_uri": "https://login.microsoftonline.com/common/oauth2/token",
			}
		).insert(ignore_permissions=True)

		result = provisioning.apply()

		self.assertEqual(result["created"], [])
		hand_tuned.reload()
		self.assertEqual(hand_tuned.client_id, "someone-elses-id")
		self.assertEqual(hand_tuned.token_uri, "https://login.microsoftonline.com/common/oauth2/token")

		step = next(s for s in result["plan"]["steps"] if s["capability"] == "mail")
		self.assertEqual(step["action"], provisioning.DRIFT)
		self.assertTrue(step["findings"], "drift should be explained, not just flagged")

	def test_nothing_is_created_for_an_unselected_capability(self):
		configure_settings(use_mail=0)

		result = provisioning.apply()

		self.assertEqual(result["created"], [])
		self.assertFalse(frappe.db.exists("Connected App", {"provider_name": provisioning.MAIL_APP_NAME}))


class TestCapabilityAwareDiagnostics(BaseTestCase):
	def setUp(self):
		super().setUp()
		self.addCleanup(reset_settings)

	def _set(self, **values):
		for field, value in values.items():
			frappe.db.set_single_value("Microsoft Settings", field, value)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")
		self.addCleanup(frappe.clear_document_cache, "Microsoft Settings", "Microsoft Settings")

	def test_mail_is_not_nagged_about_when_it_was_never_wanted(self):
		self._set(use_mail=0, use_sso=0)

		checks = {f["check"] for f in doctor.run_diagnostics()["findings"]}

		self.assertNotIn("email_account.none", checks)

	def test_mail_is_checked_once_selected(self):
		self._set(use_mail=1)

		checks = {f["check"] for f in doctor.run_diagnostics()["findings"]}

		self.assertIn("email_account.none", checks)

	def test_sso_selected_but_missing_is_reported(self):
		self._set(use_sso=1)
		if frappe.db.exists("Social Login Key", {"social_login_provider": "Office 365"}):
			self.skipTest("site already has a Microsoft Social Login Key")

		checks = {f["check"] for f in doctor.run_diagnostics()["findings"]}

		self.assertIn("sso.missing", checks)


class TestSsoChecks(BaseTestCase):
	def test_common_authority_is_rejected_for_a_single_tenant_app(self):
		key = {
			"name": "Office 365",
			"enable_social_login": 1,
			"client_id": "client-abc",
			"authorize_url": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
			"access_token_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		}

		checks = {f["check"] for f in doctor.check_social_login_key(key, settings())}

		self.assertIn("sso.tenant_authorize_url", checks)

	def test_v1_endpoints_are_rejected(self):
		key = {
			"name": "Office 365",
			"enable_social_login": 1,
			"client_id": "client-abc",
			"authorize_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/authorize",
			"access_token_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		}

		checks = {f["check"] for f in doctor.check_social_login_key(key, settings())}

		self.assertIn("sso.version_authorize_url", checks)

	def test_client_id_mismatch_is_a_failure(self):
		key = {
			"name": "Office 365",
			"enable_social_login": 1,
			"client_id": "a-different-app",
			"authorize_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
			"access_token_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		}

		checks = {f["check"] for f in doctor.check_social_login_key(key, settings())}

		self.assertIn("sso.client_id", checks)

	def test_disabled_sign_in_is_a_warning(self):
		key = {
			"name": "Office 365",
			"enable_social_login": 0,
			"client_id": "client-abc",
			"authorize_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
			"access_token_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		}

		checks = {f["check"] for f in doctor.check_social_login_key(key, settings())}

		self.assertIn("sso.enabled", checks)
