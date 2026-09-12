"""Connection doctor tests.

Every rule is exercised against plain config dicts — no tenant, no network, no fixtures.
That is deliberate: the doctor has to be trustworthy before anyone has a tenant to try it on.
"""

from frappe_microsoft365 import doctor
from frappe_microsoft365.doctor import (
	APP_ONLY_SCOPE,
	DELEGATED_SCOPES,
	FAIL,
	OFFLINE_ACCESS,
	PASS,
	SKIP,
	WARN,
)
from frappe_microsoft365.tests.base import BaseTestCase

TENANT = "11111111-2222-3333-4444-555555555555"


def ids(findings, status=None):
	return {f["check"] for f in findings if status is None or f["status"] == status}


def good_settings(**overrides):
	values = {
		"enabled": 1,
		"tenant_id": TENANT,
		"client_id": "client-abc",
		"has_client_secret": True,
		"redirect_uri": f"https://site.example.com/api/method/{doctor.__name__}.callback",
	}
	values.update(overrides)
	return values


def good_delegated_app(**overrides):
	values = {
		"name": "Microsoft Mail",
		"client_id": "client-abc",
		"redirect_uri": good_settings()["redirect_uri"],
		"authorization_uri": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/authorize",
		"token_uri": f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token",
		"scopes": [DELEGATED_SCOPES["imap"], DELEGATED_SCOPES["smtp"], OFFLINE_ACCESS],
	}
	values.update(overrides)
	return values


def good_account(**overrides):
	values = {
		"name": "Support",
		"email_id": "support@example.com",
		"auth_method": "OAuth",
		"connected_app": "Microsoft Mail",
		"connected_user": "agent@example.com",
		"backend_app_flow": 0,
		"login_id": None,
		"login_id_is_different": 0,
		"enable_incoming": 1,
		"enable_outgoing": 1,
		"use_imap": 1,
		"use_ssl": 1,
		"use_starttls": 0,
		"email_server": "outlook.office365.com",
		"smtp_server": "smtp.office365.com",
		"imap_folder": 1,
		"service": "Outlook.com",
	}
	values.update(overrides)
	return values


class TestSettingsChecks(BaseTestCase):
	def test_healthy_settings_produce_nothing(self):
		self.assertEqual(doctor.check_settings(good_settings()), [])

	def test_missing_credentials_are_reported_together(self):
		found = doctor.check_settings(good_settings(client_id="", has_client_secret=False))

		self.assertIn("settings.credentials", ids(found, FAIL))
		detail = next(f for f in found if f["check"] == "settings.credentials")["detail"]
		self.assertIn("Client ID", detail)
		self.assertIn("Client Secret", detail)

	def test_common_tenant_is_flagged_for_app_only(self):
		"""Client credentials cannot use the multi-tenant authority."""
		found = doctor.check_settings(good_settings(tenant_id="common"))

		self.assertIn("settings.tenant_id", ids(found, WARN))

	def test_plain_http_redirect_is_rejected_but_localhost_is_fine(self):
		self.assertIn(
			"settings.redirect_uri",
			ids(doctor.check_settings(good_settings(redirect_uri="http://site.example.com/cb")), FAIL),
		)
		self.assertNotIn(
			"settings.redirect_uri",
			ids(doctor.check_settings(good_settings(redirect_uri="http://m365.localhost:8000/cb"))),
		)

	def test_disabled_integration_is_a_warning_not_a_failure(self):
		found = doctor.check_settings(good_settings(enabled=0))
		self.assertIn("settings.enabled", ids(found, WARN))


class TestConnectedAppChecks(BaseTestCase):
	def test_healthy_delegated_app_has_no_problems(self):
		found = doctor.check_connected_app(good_delegated_app(), good_settings())

		self.assertEqual([f for f in found if f["status"] in (FAIL, WARN)], [])

	def test_missing_offline_access_is_caught(self):
		"""The documented cause of 'it works for a few hours then stops'."""
		app = good_delegated_app(scopes=[DELEGATED_SCOPES["imap"], DELEGATED_SCOPES["smtp"]])

		found = doctor.check_connected_app(app, good_settings())

		self.assertIn("connected_app.offline_access", ids(found, FAIL))

	def test_no_scopes_at_all(self):
		found = doctor.check_connected_app(good_delegated_app(scopes=[]), good_settings())
		self.assertIn("connected_app.scopes", ids(found, FAIL))

	def test_wrong_scope_family_for_delegated(self):
		app = good_delegated_app(scopes=["https://graph.microsoft.com/Mail.Read", OFFLINE_ACCESS])

		found = doctor.check_connected_app(app, good_settings())

		self.assertIn("connected_app.scopes_delegated", ids(found, FAIL))

	def test_app_only_must_use_default_scope(self):
		app = good_delegated_app()

		found = doctor.check_connected_app(app, good_settings(), app_only=True)

		self.assertIn("connected_app.scopes_app_only", ids(found, FAIL))

	def test_app_only_with_correct_scope_is_clean(self):
		app = good_delegated_app(scopes=[APP_ONLY_SCOPE])

		found = doctor.check_connected_app(app, good_settings(), app_only=True)

		self.assertEqual([f for f in found if f["status"] in (FAIL, WARN)], [])

	def test_app_only_extra_scopes_warn(self):
		app = good_delegated_app(scopes=[APP_ONLY_SCOPE, OFFLINE_ACCESS])

		found = doctor.check_connected_app(app, good_settings(), app_only=True)

		self.assertIn("connected_app.scopes_app_only_extra", ids(found, WARN))

	def test_v1_endpoint_is_rejected(self):
		app = good_delegated_app(
			token_uri=f"https://login.microsoftonline.com/{TENANT}/oauth2/token"
		)

		found = doctor.check_connected_app(app, good_settings())

		self.assertIn("connected_app.endpoint_version", ids(found, FAIL))

	def test_tenant_mismatch_between_settings_and_endpoints(self):
		other = "99999999-8888-7777-6666-555555555555"
		app = good_delegated_app(
			token_uri=f"https://login.microsoftonline.com/{other}/oauth2/v2.0/token"
		)

		found = doctor.check_connected_app(app, good_settings())

		self.assertIn("connected_app.tenant_token", ids(found, FAIL))

	def test_the_second_redirect_uri_is_surfaced_for_azure_registration(self):
		"""Frappe computes this endpoint from the record name; Azure needs it registered too."""
		app = good_delegated_app(redirect_uri="https://site.example.com/api/method/...callback/abc")

		found = doctor.check_connected_app(app, good_settings())

		notice = next(f for f in found if f["check"] == "connected_app.redirect_uri_registration")
		self.assertEqual(notice["status"], SKIP)
		self.assertIn("callback/abc", notice["detail"])

	def test_missing_redirect_uri_is_a_failure(self):
		found = doctor.check_connected_app(good_delegated_app(redirect_uri=""), good_settings())

		self.assertIn("connected_app.redirect_uri", ids(found, FAIL))


class TestEmailAccountChecks(BaseTestCase):
	def test_healthy_account_produces_nothing(self):
		self.assertEqual(doctor.check_email_account(good_account()), [])

	def test_basic_auth_account_is_skipped_with_a_deadline_note(self):
		found = doctor.check_email_account(good_account(auth_method="Basic"))

		self.assertIn("email_account.auth_method", ids(found, SKIP))
		self.assertIn("2026", found[0]["detail"])

	def test_oauth_without_connected_app(self):
		found = doctor.check_email_account(good_account(connected_app=None))
		self.assertIn("email_account.connected_app", ids(found, FAIL))

	def test_delegated_without_connected_user(self):
		found = doctor.check_email_account(good_account(connected_user=None))
		self.assertIn("email_account.connected_user", ids(found, FAIL))

	def test_app_only_does_not_need_a_connected_user(self):
		account = good_account(connected_user=None, backend_app_flow=1)

		self.assertNotIn("email_account.connected_user", ids(doctor.check_email_account(account)))

	def test_shared_mailbox_identity_conflict_is_flagged(self):
		"""IMAP wants the mailbox address, SMTP wants the signing-in user; one field cannot do both."""
		account = good_account(login_id_is_different=1, login_id="agent@example.com")

		found = doctor.check_email_account(account)

		self.assertIn("email_account.shared_mailbox_identity", ids(found, WARN))

	def test_no_identity_conflict_when_only_incoming_is_enabled(self):
		account = good_account(
			login_id_is_different=1, login_id="agent@example.com", enable_outgoing=0
		)

		self.assertNotIn(
			"email_account.shared_mailbox_identity", ids(doctor.check_email_account(account))
		)

	def test_no_identity_conflict_on_app_only(self):
		account = good_account(
			login_id_is_different=1, login_id="agent@example.com", backend_app_flow=1
		)

		self.assertNotIn(
			"email_account.shared_mailbox_identity", ids(doctor.check_email_account(account))
		)

	def test_imap_without_folder(self):
		found = doctor.check_email_account(good_account(imap_folder=0))
		self.assertIn("email_account.imap_folder", ids(found, FAIL))

	def test_both_tls_modes_enabled(self):
		found = doctor.check_email_account(good_account(use_starttls=1))
		self.assertIn("email_account.tls", ids(found, WARN))

	def test_wrong_incoming_host(self):
		found = doctor.check_email_account(good_account(email_server="imap.gmail.com"))
		self.assertIn("email_account.imap_host", ids(found, WARN))


class TestErrorDecoder(BaseTestCase):
	def test_recognises_the_common_microsoft_failures(self):
		cases = {
			"AADSTS50011: The redirect URI specified does not match": "Redirect URI",
			"AADSTS65001: The user or administrator has not consented": "consent",
			"AADSTS7000215: Invalid client secret provided": "secret",
			"invalid_grant: token expired": "refresh token",
			"smtplib.SMTPAuthenticationError (535, 5.7.3 Authentication unsuccessful)": "SMTP",
			"A01 NO AUTHENTICATE failed.": "IMAP",
			"451 4.7.0 Temporary server error": "temporary block",
			"Client host rejected: TLS required": "TLS",
		}
		for text, expected in cases.items():
			result = doctor.explain_error(text)
			with self.subTest(text=text):
				self.assertTrue(result["matched"], f"should recognise: {text}")
				self.assertIn(expected.lower(), (result["title"] + result["detail"]).lower())

	def test_unknown_error_is_honest_about_it(self):
		result = doctor.explain_error("something entirely unrelated")

		self.assertFalse(result["matched"])
		self.assertIn("Unrecognised", result["title"])

	def test_empty_input_does_not_crash(self):
		self.assertFalse(doctor.explain_error("")["matched"])
		self.assertFalse(doctor.explain_error(None)["matched"])


class TestPowerShellGenerator(BaseTestCase):
	def test_looks_up_object_id_instead_of_asking_for_it(self):
		"""Microsoft's documented #1 trap: the App Registration Object ID is the wrong one."""
		script = doctor.powershell_for_app_only("client-abc", mailboxes=["support@example.com"])

		self.assertIn("appId eq '$appId'", script)
		self.assertIn("New-ServicePrincipal", script)
		self.assertIn("Add-MailboxPermission", script)
		self.assertIn("support@example.com", script)

	def test_uses_a_supplied_object_id_when_known(self):
		script = doctor.powershell_for_app_only("client-abc", enterprise_object_id="obj-123")

		self.assertIn('$objectId = "obj-123"', script)
		self.assertNotIn("Get-MgServicePrincipal", script)

	def test_send_as_permission_is_opt_in(self):
		without = doctor.powershell_for_app_only("c", mailboxes=["a@b.com"])
		with_send = doctor.powershell_for_app_only("c", mailboxes=["a@b.com"], send_as=True)

		self.assertNotIn("Add-RecipientPermission", without)
		self.assertIn("Add-RecipientPermission", with_send)

	def test_every_mailbox_is_granted_individually(self):
		script = doctor.powershell_for_app_only("c", mailboxes=["a@b.com", "c@d.com"])

		self.assertEqual(script.count("Add-MailboxPermission"), 2)


class TestDiagnosticsEndpoint(BaseTestCase):
	def test_runs_against_the_real_site_without_raising(self):
		result = doctor.run_diagnostics()

		self.assertIn("findings", result)
		self.assertIn("counts", result)
		for item in result["findings"]:
			self.assertIn(item["status"], (PASS, WARN, FAIL, SKIP))
			self.assertTrue(item["check"])
			self.assertTrue(item["title"])
