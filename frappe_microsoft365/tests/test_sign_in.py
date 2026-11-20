"""Sign in with Microsoft accepts only the organisation's own directory.

Covers the three layers in sign_in.py: provisioning and the doctor refuse a multi-tenant
authority, the Frappe 16 hook refuses a token from another directory, and the upgrade patch
names the directory from the calendars' tokens when they agree.
"""

import base64
import json
from unittest.mock import MagicMock, patch

import frappe

from frappe_microsoft365 import doctor, provisioning, sign_in
from frappe_microsoft365.patches import fill_sign_in_tenant_id
from frappe_microsoft365.tests.base import BaseTestCase

TENANT = "11111111-2222-3333-4444-555555555555"
OTHER = "99999999-8888-7777-6666-555555555555"
CLIENT = "038b9c8e-2699-4858-b5f1-4f7a3d4077c4"
COMMON_AUTHORIZE = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
COMMON_TOKEN = "https://login.microsoftonline.com/common/oauth2/v2.0/token"


def jwt_with(claims):
	"""An unsigned token shaped like the access tokens Microsoft issues."""

	def part(data):
		return base64.urlsafe_b64encode(json.dumps(data).encode()).rstrip(b"=").decode()

	return f"{part({'alg': 'none', 'typ': 'JWT'})}.{part(claims)}.c2ln"


def set_settings(**values):
	for field, value in values.items():
		frappe.db.set_single_value("Microsoft Settings", field, value)
	frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")


class SignInTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		before = {
			f: frappe.db.get_single_value("Microsoft Settings", f)
			for f in ("tenant_id", "client_id", "use_sso", "use_calendar", "use_mail")
		}
		self.addCleanup(set_settings, **before)
		self._drop_key()
		self.addCleanup(self._drop_key)

	def _drop_key(self):
		for name in frappe.get_all(
			"Social Login Key", filters={"social_login_provider": "Office 365"}, pluck="name"
		):
			frappe.delete_doc("Social Login Key", name, force=True, ignore_permissions=True)

	def _key(self, authorize=COMMON_AUTHORIZE, token=COMMON_TOKEN, client_id=CLIENT, enabled=0):
		return frappe.get_doc(
			{
				"doctype": "Social Login Key",
				"social_login_provider": "Office 365",
				"provider_name": "Office 365",
				"client_id": client_id,
				"base_url": "https://login.microsoftonline.com",
				"authorize_url": authorize,
				"access_token_url": token,
				"redirect_url": "/api/method/frappe.integrations.oauth2_logins.login_via_office365",
				"enable_social_login": enabled,
				"client_secret": "test-secret",
			}
		).insert(ignore_permissions=True)


class TestTenantShape(BaseTestCase):
	def test_a_guid_or_a_domain_names_one_directory(self):
		self.assertTrue(sign_in.is_single_tenant(TENANT))
		self.assertTrue(sign_in.is_single_tenant("contoso.onmicrosoft.com"))
		self.assertTrue(sign_in.is_single_tenant("Contoso.COM"))

	def test_multi_tenant_authorities_and_junk_do_not(self):
		for tenant in ("", None, "common", "Organizations", "consumers", "contoso", "a b.com", "x/../y.com"):
			self.assertFalse(sign_in.is_single_tenant(tenant), tenant)


class TestProvisioningRefusesAnOpenSignIn(SignInTestCase):
	def test_common_and_organizations_block_sign_in(self):
		for tenant in ("common", "organizations", "consumers"):
			set_settings(tenant_id=tenant, client_id=CLIENT, use_sso=1)

			blockers = provisioning.plan()["blockers"]

			self.assertTrue(any("Sign in with Microsoft" in b for b in blockers), (tenant, blockers))

	def test_a_specific_tenant_is_not_blocked(self):
		for tenant in (TENANT, "contoso.onmicrosoft.com"):
			set_settings(tenant_id=tenant, client_id=CLIENT, use_sso=1)

			blockers = provisioning.plan()["blockers"]

			self.assertFalse(any("Sign in with Microsoft" in b for b in blockers), (tenant, blockers))

	def test_multi_tenant_does_not_block_capabilities_other_than_sign_in(self):
		set_settings(tenant_id="common", client_id=CLIENT, use_sso=0)

		blockers = provisioning.plan()["blockers"]

		self.assertFalse(any("Sign in with Microsoft" in b for b in blockers), blockers)

	def test_the_record_itself_is_never_created_on_common(self):
		with self.assertRaises(frappe.ValidationError):
			provisioning._create_social_login_key(
				{"tenant_id": "common", "client_id": CLIENT, "client_secret": "x"}
			)
		self.assertFalse(frappe.db.exists("Social Login Key", sign_in.PROVIDER))


class TestDoctorReportsAnOpenSignIn(SignInTestCase):
	def test_multi_tenant_tenant_id_is_a_failure(self):
		key = {"name": "office_365", "authorize_url": COMMON_AUTHORIZE, "access_token_url": COMMON_TOKEN}

		for tenant in ("common", "organizations", ""):
			found = doctor.check_sign_in_tenant(key, {"tenant_id": tenant})

			failures = {f["check"] for f in found if f["status"] == doctor.FAIL}
			self.assertIn("sso.multi_tenant", failures, tenant)
			self.assertIn("sso.tenant_authorize_url", failures, tenant)

	def test_organizations_url_is_caught_not_only_common(self):
		key = {
			"name": "office_365",
			"authorize_url": "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize",
			"access_token_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		}

		found = doctor.check_sign_in_tenant(key, {"tenant_id": TENANT})

		self.assertEqual([f["check"] for f in found], ["sso.tenant_authorize_url"])

	def test_a_single_tenant_key_is_clean(self):
		key = {
			"name": "office_365",
			"authorize_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
			"access_token_url": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		}

		self.assertEqual(doctor.check_sign_in_tenant(key, {"tenant_id": TENANT}), [])

	def test_run_diagnostics_flags_it_even_when_sign_in_is_not_ticked(self):
		self._key()
		set_settings(tenant_id="common", use_sso=0)

		findings = doctor.run_diagnostics()["findings"]

		self.assertIn(("sso.multi_tenant", doctor.FAIL), {(f["check"], f["status"]) for f in findings})

	def test_upgrade_check_reports_an_enabled_open_sign_in(self):
		self._key(enabled=1)
		set_settings(tenant_id="organizations")

		found = sign_in.upgrade_check()

		self.assertTrue(found)
		self.assertTrue(all(f["check"] == "upgrade.sign_in_tenant" for f in found))
		self.assertTrue(all(f["status"] == doctor.FAIL for f in found))
		self.assertTrue(all(f["fix"] for f in found))

	def test_upgrade_check_is_quiet_when_sign_in_is_single_tenant_or_absent(self):
		self.assertEqual(sign_in.upgrade_check(), [])
		self._key(
			authorize=f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
			token=f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
			enabled=1,
		)
		set_settings(tenant_id=TENANT)

		self.assertEqual(sign_in.upgrade_check(), [])


class TestTokenDirectoryCheck(BaseTestCase):
	"""sign_in.refusal, the pure rule the Frappe 16 hook applies."""

	def test_a_guid_tenant_refuses_any_other_directory(self):
		self.assertIsNotNone(sign_in.refusal({"tid": OTHER}, TENANT, COMMON_TOKEN))

	def test_a_guid_tenant_refuses_a_token_without_a_directory(self):
		self.assertIsNotNone(sign_in.refusal({}, TENANT, COMMON_TOKEN))

	def test_a_guid_tenant_accepts_its_own_directory_in_any_case(self):
		self.assertIsNone(sign_in.refusal({"tid": TENANT.upper()}, TENANT, COMMON_TOKEN))

	def test_a_domain_tenant_trusts_its_own_endpoint(self):
		endpoint = "https://login.microsoftonline.com/contoso.com/oauth2/v2.0/token"

		self.assertIsNone(sign_in.refusal({"tid": OTHER}, "Contoso.com", endpoint))

	def test_a_domain_tenant_accepts_a_guid_endpoint_matching_the_token(self):
		endpoint = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"

		self.assertIsNone(sign_in.refusal({"tid": TENANT}, "contoso.com", endpoint))
		self.assertIsNotNone(sign_in.refusal({"tid": OTHER}, "contoso.com", endpoint))

	def test_a_domain_tenant_on_a_multi_tenant_endpoint_is_refused(self):
		self.assertIsNotNone(sign_in.refusal({"tid": TENANT}, "contoso.com", COMMON_TOKEN))

	def test_multi_tenant_keeps_signing_people_in_during_1_1_1(self):
		"""Nobody is locked out by the upgrade; enforcement starts in 1.2.0."""
		for tenant in ("common", "organizations", "", None):
			self.assertIsNone(sign_in.refusal({"tid": OTHER}, tenant, COMMON_TOKEN), tenant)


class TestSignInHook(SignInTestCase):
	def test_the_hook_is_registered(self):
		self.assertIn("frappe_microsoft365.sign_in.on_social_login", frappe.get_hooks("on_social_login"))

	def test_other_providers_are_left_alone(self):
		set_settings(tenant_id=TENANT)

		sign_in.on_social_login(user="a@example.com", provider="google", claims={"tid": OTHER})

	def test_a_token_from_another_directory_is_refused(self):
		set_settings(tenant_id=TENANT)

		with self.assertRaises(frappe.AuthenticationError):
			sign_in.on_social_login(user="a@example.com", provider="office_365", claims={"tid": OTHER})

		sign_in.on_social_login(user="a@example.com", provider="office_365", claims={"tid": TENANT})

	def test_frappe_login_stops_before_the_session_is_created(self):
		"""End to end through Frappe 16's login_oauth_user: no session for a foreign directory."""
		from frappe.utils import oauth

		if not hasattr(oauth, "create_oauth_state") or "on_social_login" not in (
			oauth.login_oauth_user.__code__.co_consts
		):
			self.skipTest("this Frappe has no on_social_login hook (Frappe 15)")

		set_settings(tenant_id=TENANT)
		self._key(
			authorize=f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
			token=f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
			enabled=1,
		)
		login_manager = MagicMock()
		claims = {"email": "administrator@example.com", "sub": "s", "tid": OTHER, "given_name": "A"}

		with (
			patch.object(frappe.local, "login_manager", login_manager, create=True),
			patch("frappe.utils.oauth.update_oauth_user", return_value=None),
		):
			with self.assertRaises(frappe.AuthenticationError):
				oauth.login_oauth_user(claims, provider="office_365", state=oauth.create_oauth_state(None))
			login_manager.login_as.assert_not_called()

			claims["tid"] = TENANT
			with patch("frappe.utils.oauth.redirect_post_login"):
				oauth.login_oauth_user(claims, provider="office_365", state=oauth.create_oauth_state(None))
			login_manager.login_as.assert_called_once()


class TestTokenDirectory(BaseTestCase):
	def test_reads_tid_from_a_jwt(self):
		self.assertEqual(sign_in.tid_of(jwt_with({"tid": TENANT.upper()})), TENANT)

	def test_opaque_or_broken_tokens_give_nothing(self):
		for token in (
			None,
			"",
			"EwBwA8l6BAAU",
			"a.b",
			"a.!!!.c",
			jwt_with({"tid": "not-a-guid"}),
			jwt_with([1]),
		):
			self.assertIsNone(sign_in.tid_of(token), token)


class TestFillTenantPatch(SignInTestCase):
	def setUp(self):
		super().setUp()
		# Only this test's calendars count; the class's transaction is rolled back afterwards.
		for name in frappe.get_all("Microsoft Calendar", pluck="name"):
			frappe.delete_doc("Microsoft Calendar", name, force=True, ignore_permissions=True)
		set_settings(tenant_id="common", client_id=CLIENT)

	def _calendar(self, token, authorized=1):
		return frappe.get_doc(
			{
				"doctype": "Microsoft Calendar",
				"account_name": f"Tenant patch {frappe.generate_hash(length=6)}",
				"user": "Administrator",
				"enabled": 1,
				"authorized": authorized,
				"access_token": token,
			}
		).insert(ignore_permissions=True)

	def _tenant(self):
		return frappe.db.get_single_value("Microsoft Settings", "tenant_id")

	def test_agreeing_tokens_fill_tenant_id_and_close_the_urls(self):
		self._key(enabled=1)
		self._calendar(jwt_with({"tid": TENANT}))
		self._calendar(jwt_with({"tid": TENANT}))

		fill_sign_in_tenant_id.execute()
		fill_sign_in_tenant_id.execute()  # idempotent

		self.assertEqual(self._tenant(), TENANT)
		key = frappe.db.get_value(
			"Social Login Key", sign_in.PROVIDER, ["authorize_url", "access_token_url"], as_dict=True
		)
		self.assertEqual(
			key.authorize_url, f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize"
		)
		self.assertEqual(
			key.access_token_url, f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
		)
		self.assertEqual(sign_in.upgrade_check(), [])

	def test_disagreeing_tokens_change_nothing_and_are_reported(self):
		self._key(enabled=1)
		self._calendar(jwt_with({"tid": TENANT}))
		self._calendar(jwt_with({"tid": OTHER}))

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "common")
		self.assertTrue(sign_in.upgrade_check())

	def test_an_unreadable_token_changes_nothing(self):
		self._key()
		self._calendar(jwt_with({"tid": TENANT}))
		self._calendar("EwBwA8l6BAAU-opaque")

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "common")

	def test_personal_accounts_change_nothing(self):
		self._key()
		self._calendar(jwt_with({"tid": sign_in.CONSUMER_TENANT}))

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "common")

	def test_no_authorised_calendar_changes_nothing(self):
		self._key()
		self._calendar(jwt_with({"tid": TENANT}), authorized=0)

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "common")

	def test_a_specific_tenant_is_never_touched(self):
		self._key()
		set_settings(tenant_id="contoso.onmicrosoft.com")
		self._calendar(jwt_with({"tid": TENANT}))

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "contoso.onmicrosoft.com")

	def test_a_site_without_microsoft_sign_in_is_left_alone(self):
		self._calendar(jwt_with({"tid": TENANT}))

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "common")

	def test_a_key_naming_another_directory_is_left_for_a_person(self):
		self._key(
			authorize=f"https://login.microsoftonline.com/{OTHER}/oauth2/v2.0/authorize",
			token=f"https://login.microsoftonline.com/{OTHER}/oauth2/v2.0/token",
		)
		self._calendar(jwt_with({"tid": TENANT}))

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), "common")

	def test_a_key_for_another_registration_keeps_its_urls(self):
		self._key(client_id="00000000-0000-4000-8000-000000000001")
		self._calendar(jwt_with({"tid": TENANT}))

		fill_sign_in_tenant_id.execute()

		self.assertEqual(self._tenant(), TENANT)
		self.assertEqual(
			frappe.db.get_value("Social Login Key", sign_in.PROVIDER, "authorize_url"), COMMON_AUTHORIZE
		)
