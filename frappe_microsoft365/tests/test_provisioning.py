"""Provisioning tests.

The value builders are pure, so the endpoints and scope strings that decide whether Microsoft
accepts the configuration are all verified offline. The plan/apply layer is tested for the two
properties that matter: it is modular (nothing is created for a capability you did not ask
for) and it is never destructive (an existing record is reported, never rewritten).
"""

import json

import frappe

from frappe_microsoft365 import doctor, provisioning
from frappe_microsoft365 import microsoft_graph as graph
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
		use_teams=0,
		use_transcripts=0,
		use_mail=0,
		use_sso=0,
		mail_flow="Delegated",
		default_scopes="",
		authorized_scopes="",
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
		"use_teams": 0,
		"use_transcripts": 0,
		"use_mail": 0,
		"use_sso": 0,
		"mail_flow": "Delegated",
		"default_scopes": "",
		"authorized_scopes": "",
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


class TestScopeDerivation(BaseTestCase):
	"""Pure derivation, plain dicts: no site, no tenant, no network."""

	def test_calendar_alone_does_not_ask_for_teams_permissions(self):
		"""The bug this split exists for.

		Creating an event with a Teams link is a calendar write (isOnlineMeeting on
		POST /me/events), so wanting the calendar must never drag OnlineMeetings or transcript
		consent along — admins were deleting them from the scope field by hand.
		"""
		scopes = graph.derive_scopes({"use_calendar": 1})

		self.assertEqual(scopes, ["User.Read", "Calendars.ReadWrite"])

	def test_nothing_ticked_still_identifies_the_account(self):
		self.assertEqual(graph.derive_scopes({}), ["User.Read"])

	def test_standalone_meetings_add_the_meeting_permission(self):
		scopes = graph.derive_scopes({"use_calendar": 1, "use_teams": 1})

		self.assertIn("OnlineMeetings.ReadWrite", scopes)
		self.assertNotIn("OnlineMeetingTranscript.Read.All", scopes)

	def test_transcripts_add_only_the_transcript_permission(self):
		scopes = graph.derive_scopes({"use_teams": 1, "use_transcripts": 1})

		self.assertEqual(
			scopes, ["User.Read", "OnlineMeetings.ReadWrite", "OnlineMeetingTranscript.Read.All"]
		)

	def test_everything_ticked_is_the_full_list(self):
		scopes = graph.derive_scopes({"use_calendar": 1, "use_teams": 1, "use_transcripts": 1})

		self.assertEqual(
			scopes,
			[
				"User.Read",
				"Calendars.ReadWrite",
				"OnlineMeetings.ReadWrite",
				"OnlineMeetingTranscript.Read.All",
			],
		)

	def test_reserved_scopes_are_never_derived(self):
		"""build_authorize_url adds them; MSAL refuses them in the token calls."""
		scopes = graph.derive_scopes({"use_calendar": 1, "use_teams": 1, "use_transcripts": 1})

		self.assertEqual([s for s in scopes if s in graph.RESERVED_SCOPES], [])


class TestScopeOverride(BaseTestCase):
	def test_no_override_derives_from_the_capabilities(self):
		self.assertEqual(
			graph.get_scopes(settings(use_calendar=1)), ["User.Read", "Calendars.ReadWrite"]
		)

	def test_override_is_returned_verbatim(self):
		"""A tenant that consented to a hand-picked list must be asked for exactly that."""
		scopes = graph.get_scopes(settings(use_calendar=1, default_scopes="User.Read Mail.Read"))

		self.assertEqual(scopes, ["User.Read", "Mail.Read"])

	def test_override_is_not_widened_by_a_ticked_capability(self):
		scopes = graph.get_scopes(
			settings(use_calendar=1, use_teams=1, default_scopes="User.Read Calendars.ReadWrite")
		)

		self.assertNotIn("OnlineMeetings.ReadWrite", scopes)

	def test_override_accepts_commas_and_line_breaks(self):
		scopes = graph.get_scopes(settings(default_scopes="User.Read, Calendars.ReadWrite\nMail.Read"))

		self.assertEqual(scopes, ["User.Read", "Calendars.ReadWrite", "Mail.Read"])

	def test_blank_override_is_not_an_override(self):
		scopes = graph.get_scopes(settings(use_calendar=1, default_scopes="   \n "))

		self.assertEqual(scopes, ["User.Read", "Calendars.ReadWrite"])


class TestCapabilityAzurePermissions(BaseTestCase):
	"""The Set Up dialog renders capability["azure"], so it has to be what is really requested."""

	def _capability(self, capability_id):
		return next(c for c in provisioning.capabilities() if c["id"] == capability_id)

	def test_each_graph_capability_lists_exactly_what_it_will_request(self):
		for capability_id, field in provisioning.GRAPH_CAPABILITY_FIELDS.items():
			with self.subTest(capability=capability_id):
				listed = set(self._capability(capability_id)["azure"])

				self.assertEqual(
					listed - {OFFLINE_ACCESS}, set(graph.derive_scopes({field: 1}))
				)
				# Requested on every sign-in and consented like any other permission.
				self.assertIn(OFFLINE_ACCESS, listed)

	def test_the_combined_lists_match_the_derived_scopes(self):
		everything = dict.fromkeys(provisioning.GRAPH_CAPABILITY_FIELDS.values(), 1)
		listed = set()
		for capability_id in provisioning.GRAPH_CAPABILITY_FIELDS:
			listed |= set(self._capability(capability_id)["azure"])

		self.assertEqual(listed - {OFFLINE_ACCESS}, set(graph.derive_scopes(everything)))

	def test_graph_capabilities_do_not_advertise_openid_or_profile(self):
		"""Those are sign-in scopes, requested automatically; only offline_access is consented."""
		for capability_id in provisioning.GRAPH_CAPABILITY_FIELDS:
			with self.subTest(capability=capability_id):
				listed = set(self._capability(capability_id)["azure"])

				self.assertEqual(listed & {"openid", "profile"}, set())


class TestScopeChecks(BaseTestCase):
	"""doctor.check_scopes / check_authorized_scopes — pure, plain dicts."""

	def _ids(self, found):
		return {f["check"] for f in found}

	def test_derived_scopes_are_never_flagged(self):
		self.assertEqual(doctor.check_scopes(settings(use_calendar=1, use_teams=1)), [])

	def test_override_missing_a_needed_scope_names_it(self):
		found = doctor.check_scopes(
			settings(use_calendar=1, use_teams=1, default_scopes="User.Read Calendars.ReadWrite")
		)

		self.assertIn("scopes.override_incomplete", self._ids(found))
		self.assertIn("OnlineMeetings.ReadWrite", " ".join(f["title"] for f in found))

	def test_a_complete_override_is_silent(self):
		found = doctor.check_scopes(
			settings(use_calendar=1, default_scopes="User.Read Calendars.ReadWrite Mail.Read")
		)

		self.assertEqual(found, [])

	def test_reserved_scopes_in_the_override_are_a_failure(self):
		found = doctor.check_scopes(
			settings(use_calendar=1, default_scopes="offline_access User.Read Calendars.ReadWrite")
		)

		self.assertIn("scopes.override_reserved", self._ids(found))


class TestReauthorizationWarning(BaseTestCase):
	CURRENT = ("User.Read", "Calendars.ReadWrite", "OnlineMeetings.ReadWrite")

	def _ids(self, found):
		return {f["check"] for f in found}

	def test_unchanged_scopes_say_nothing(self):
		found = doctor.check_authorized_scopes(self.CURRENT, list(self.CURRENT), connections=2)

		self.assertEqual(found, [])

	def test_a_widened_scope_list_asks_for_re_authorisation(self):
		"""A token carries the permissions it was issued with; ticking a box does not widen it."""
		found = doctor.check_authorized_scopes(
			self.CURRENT, ["User.Read", "Calendars.ReadWrite"], connections=1
		)

		self.assertIn("scopes.changed_since_authorization", self._ids(found))
		self.assertIn("OnlineMeetings.ReadWrite", found[0]["detail"])

	def test_a_narrowed_scope_list_is_reported_too(self):
		found = doctor.check_authorized_scopes(["User.Read"], list(self.CURRENT), connections=1)

		self.assertIn("scopes.changed_since_authorization", self._ids(found))

	def test_nothing_recorded_means_no_guess(self):
		"""Connections that predate the recording get no invented verdict."""
		self.assertEqual(doctor.check_authorized_scopes(self.CURRENT, [], connections=3), [])

	def test_no_authorised_connection_means_nothing_to_re_authorise(self):
		found = doctor.check_authorized_scopes(self.CURRENT, ["User.Read"], connections=0)

		self.assertEqual(found, [])


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
