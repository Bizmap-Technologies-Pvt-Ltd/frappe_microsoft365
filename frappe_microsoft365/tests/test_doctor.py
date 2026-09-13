"""Connection doctor tests.

Every rule is exercised against plain config dicts — no tenant, no network, no fixtures.
That is deliberate: the doctor has to be trustworthy before anyone has a tenant to try it on.
"""

from frappe_microsoft365 import background, doctor
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


class TestCredentialShapes(BaseTestCase):
	"""The Secret ID sits next to the secret Value in Azure and only the Value works.

	This is the mistake that actually happened, and Microsoft only reports it at sign-in as
	AADSTS7000215, long after the person has moved on.
	"""

	def test_a_guid_in_the_secret_field_is_the_secret_id(self):
		found = doctor.check_credentials(
			good_settings(client_secret="a1b2c3d4-1234-5678-9abc-def012345678")
		)

		self.assertIn("settings.secret_is_the_id", ids(found, FAIL))
		self.assertIn("Value", found[0]["fix"])

	def test_a_real_secret_value_passes(self):
		found = doctor.check_credentials(good_settings(client_secret="8kQ~Xs9aBc.dEf-2Gh3IjK4lMn5OpQ6rSt"))

		self.assertNotIn("settings.secret_is_the_id", ids(found))

	def test_a_masked_stored_secret_is_not_judged(self):
		"""Frappe shows a saved Password field as asterisks; that is not a wrong value."""
		found = doctor.check_credentials(good_settings(client_secret="**********"))

		self.assertNotIn("settings.secret_is_the_id", ids(found))

	def test_a_client_id_that_is_not_a_guid_is_flagged(self):
		found = doctor.check_credentials(good_settings(client_id="my-app"))

		self.assertIn("settings.client_id_shape", ids(found, WARN))

	def test_a_tenant_domain_is_accepted(self):
		for tenant in ("contoso.onmicrosoft.com", "common", TENANT):
			with self.subTest(tenant=tenant):
				found = doctor.check_credentials(good_settings(tenant_id=tenant))
				self.assertNotIn("settings.tenant_id_shape", ids(found))

	def test_a_nonsense_tenant_is_flagged(self):
		found = doctor.check_credentials(good_settings(tenant_id="bizmap"))

		self.assertIn("settings.tenant_id_shape", ids(found, WARN))

	def test_a_redirect_uri_pointing_elsewhere_is_flagged(self):
		found = doctor.check_credentials(good_settings(redirect_uri="https://site.example.com/"))

		self.assertIn("settings.redirect_target", ids(found, WARN))

	def test_healthy_credentials_produce_nothing(self):
		settings = good_settings(
			client_secret="8kQ~Xs9aBc.dEf",
			client_id="038b9c8e-2699-4858-b5f1-4f7a3d4077c4",
			redirect_uri="https://site.example.com/api/method/" + doctor.CALLBACK_METHOD,
		)

		self.assertEqual(doctor.check_credentials(settings), [])


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


class TestMissingCustomFields(BaseTestCase):
	"""Regression: a plain install left these absent and sync died with a raw SQL error."""

	def test_missing_fields_are_reported_with_the_fix(self):
		found = doctor.check_event_custom_fields(["custom_sync_with_microsoft_calendar"])

		self.assertIn("fields.event", ids(found, FAIL))
		self.assertIn("migrate", found[0]["fix"].lower())

	def test_nothing_reported_when_all_present(self):
		self.assertEqual(doctor.check_event_custom_fields([]), [])

	def test_the_event_id_column_fits_a_real_graph_id(self):
		"""Regression: Graph ids are ~152 chars and Frappe's Data default is varchar(140)."""
		import frappe

		graph_id = "A" * 152
		event = frappe.get_doc(
			{
				"doctype": "Event",
				"subject": "_microsoft365 id width probe",
				"starts_on": "2026-09-20 10:00:00",
				"event_type": "Private",
			}
		).insert(ignore_permissions=True)
		self.addCleanup(
			frappe.delete_doc, "Event", event.name, force=True, ignore_permissions=True
		)

		frappe.db.set_value(
			"Event", event.name, "custom_microsoft_event_id", graph_id, update_modified=False
		)

		self.assertEqual(
			frappe.db.get_value("Event", event.name, "custom_microsoft_event_id"), graph_id
		)

	def test_this_site_has_them(self):
		"""The app is installed here, so the fields must exist."""
		import frappe

		missing = [f for f in doctor.EVENT_CUSTOM_FIELDS if not frappe.db.has_column("Event", f)]
		self.assertEqual(missing, [])

	def test_the_raw_sql_error_is_decoded(self):
		result = doctor.explain_error(
			"(1054, \"Unknown column 'custom_sync_with_microsoft_calendar' in 'WHERE'\")"
		)

		self.assertTrue(result["matched"])
		self.assertIn("migrate", result["detail"].lower())


def background_state(*problems, **overrides):
	"""A dict shaped like ``background.health()`` output, with nothing wrong by default."""
	values = {
		"ok": not problems,
		"problems": list(problems),
		"reasons": [p["reason"] for p in problems],
		"fixes": [p["fix"] for p in problems],
		"scheduler_inactive": False,
		"redis": True,
		"workers": 2,
		"queued": 0,
		"queues": {"default": 0},
		"jobs": [
			{
				"name": "frappe_microsoft365.microsoft_calendar_sync.sync_all",
				"method": "frappe_microsoft365.microsoft_calendar_sync.sync_all",
				"last_execution": "2026-09-13 10:00:00",
				"minutes_ago": 5,
				"stale": False,
			}
		],
		"backlog": False,
		"message": "",
	}
	values.update(overrides)
	return values


def problem(code, reason="something is wrong", fix="run something"):
	return {"code": code, "reason": reason, "fix": fix}


class TestBackgroundJobsCheck(BaseTestCase):
	"""Queued work that never runs is the one fault in this app that produces no error at all.

	No traceback, no log line, no wrong answer — just a calendar that quietly stops moving. If
	the doctor does not say it, nothing does.
	"""

	def test_a_working_queue_says_nothing(self):
		self.assertEqual(doctor.check_background_jobs(background_state()), [])

	def test_a_dead_scheduler_carries_the_command_that_revives_it(self):
		found = doctor.check_background_jobs(
			background_state(
				problem(background.SCHEDULER_OFF, fix="Run `bench --site x enable-scheduler`."),
				scheduler_inactive=True,
			)
		)

		self.assertIn("background.scheduler_off", ids(found, FAIL))
		self.assertIn("enable-scheduler", found[0]["fix"])

	def test_zero_workers_with_redis_up_is_a_failure_not_a_warning(self):
		"""It looks like success from every other angle, which is what makes it worth a FAIL."""
		found = doctor.check_background_jobs(
			background_state(
				problem(background.NO_WORKER, fix="Start one with `bench worker --queue default`."),
				workers=0,
			)
		)

		self.assertIn("background.no_worker", ids(found, FAIL))
		self.assertIn("bench worker", found[0]["fix"])
		self.assertIn("sit in the queue forever", found[0]["detail"])

	def test_an_unreadable_worker_count_is_only_a_warning(self):
		found = doctor.check_background_jobs(
			background_state(problem(background.WORKERS_UNKNOWN), workers=None)
		)

		self.assertIn("background.workers_unknown", ids(found, WARN))

	def test_stale_jobs_are_reported_even_when_everything_else_looks_fine(self):
		"""The scheduler process not running is invisible to the other four readings."""
		found = doctor.check_background_jobs(
			background_state(
				problem(background.JOBS_STALE, fix="Check `bench doctor`."),
				jobs=[
					{
						"name": "sync_all",
						"method": "frappe_microsoft365.microsoft_calendar_sync.sync_all",
						"last_execution": None,
						"minutes_ago": None,
						"stale": True,
					}
				],
			)
		)

		self.assertIn("background.jobs_stale", ids(found, FAIL))
		self.assertIn("has never run", found[0]["detail"])

	def test_every_finding_carries_all_five_readings(self):
		"""'No worker' and '203 waiting' and 'last ran never' are one sentence told three ways;
		an admin shown one without the others tends to fix the wrong end of it."""
		detail = doctor.check_background_jobs(
			background_state(problem(background.NO_WORKER), workers=0, queued=203)
		)[0]["detail"]

		for expected in ("Scheduler", "Redis", "Workers", "203", "sync_all"):
			self.assertIn(expected, detail)

	def test_two_faults_get_two_commands(self):
		found = doctor.check_background_jobs(
			background_state(
				problem(background.SCHEDULER_OFF, fix="enable-scheduler"),
				problem(background.NO_WORKER, fix="bench worker"),
				scheduler_inactive=True,
				workers=0,
			)
		)

		self.assertEqual(len(found), 2)
		self.assertEqual({f["fix"] for f in found}, {"enable-scheduler", "bench worker"})

	def test_a_backlog_is_worth_saying_even_on_an_otherwise_healthy_bench(self):
		"""203 jobs deep is not a fault, but it is why the sync someone is staring at has not
		happened yet."""
		found = doctor.check_background_jobs(background_state(queued=203, backlog=True))

		self.assertIn("background.backlog", ids(found, WARN))
		self.assertIn("203", found[0]["title"])

	def test_the_real_probe_and_this_check_agree_on_shape(self):
		"""A renamed key must fail here rather than in production at the moment a queue dies."""
		for item in doctor.check_background_jobs(background.health(refresh=True)):
			self.assertIn(item["status"], (WARN, FAIL))
			self.assertTrue(item["title"])
			self.assertTrue(item["fix"])
			self.assertTrue(item["detail"])


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

	def test_graphs_property_validation_error_is_explained(self):
		result = doctor.explain_error(
			"Microsoft Graph POST /me/events failed (400): ErrorPropertyValidationFailure"
		)

		self.assertTrue(result["matched"])
		self.assertIn("end time", result["detail"].lower())

	def test_unknown_error_is_honest_about_it(self):
		result = doctor.explain_error("something entirely unrelated")

		self.assertFalse(result["matched"])
		self.assertIn("Unrecognised", result["title"])

	def test_empty_input_does_not_crash(self):
		self.assertFalse(doctor.explain_error("")["matched"])
		self.assertFalse(doctor.explain_error(None)["matched"])


class TestTeamsErrorDecoder(BaseTestCase):
	"""Three of these hit one tenant in a single hour, and the app answered the first wrongly.

	Every 403 this app can produce says "Forbidden" and none of them share a fix, so ordering is
	behaviour here rather than tidiness: the wrong answer sent a real admin to grant consent that
	was already granted, which is a loop with no exit.
	"""

	def test_the_tenant_switch_403_is_not_diagnosed_as_a_consent_problem(self):
		"""The loop with no exit: every permission consented, and Graph still says Forbidden."""
		result = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings/MSpiOTM1ZTY3NS01ZTY3/transcripts failed "
			"(403): Forbidden: Graph API access to transcripts is disabled for this tenant.\n"
			"Transcripts need OnlineMeetingTranscript.Read.All: check it is listed and consented "
			"under Entra ID > App registrations > your app > API permissions."
		)

		self.assertTrue(result["matched"])
		self.assertIn("Teams admin center", result["detail"])
		self.assertIn("Transcript API access", result["detail"])
		self.assertNotIn("consent", result["detail"].lower())

	def test_the_tenant_switch_is_decoded_before_every_other_403(self):
		"""First match wins, and three later patterns also match the word Forbidden."""
		self.assertIn("GraphAccessToTranscriptsDisabled", doctor.error_patterns()[0][0])

	def test_the_apps_own_tenant_switch_message_is_recognised_too(self):
		"""What a person pastes back in is our sentence, not Microsoft's: the app replaces it."""
		from frappe_microsoft365 import microsoft_transcripts

		result = doctor.explain_error(microsoft_transcripts._tenant_switch_hint())

		self.assertTrue(result["matched"])
		self.assertIn("Meeting settings", result["detail"])

	def test_a_genuine_consent_failure_still_gets_the_consent_answer(self):
		"""Ordering the tenant switch first must not cost us the error it was hiding behind."""
		result = doctor.explain_error(
			"AADSTS65001: The user or administrator has not consented to use the application "
			"with ID '038b9c8e-2699-4858-b5f1-4f7a3d4077c4'."
		)

		self.assertTrue(result["matched"])
		self.assertIn("consent", result["title"].lower())

	def test_a_recording_403_and_a_transcript_403_give_different_advice(self):
		"""Microsoft consents to the two separately; one answer for both is wrong half the time."""
		recording = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings/MSpiOTM1ZTY3/recordings failed (403): Forbidden"
		)
		transcript = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings/MSpiOTM1ZTY3/transcripts failed (403): Forbidden"
		)

		self.assertNotEqual(recording["title"], transcript["title"])
		self.assertIn("OnlineMeetingRecording.Read.All", recording["detail"])
		self.assertNotIn("OnlineMeetingTranscript", recording["detail"])
		self.assertIn("OnlineMeetingTranscript.Read.All", transcript["detail"])
		self.assertNotIn("OnlineMeetingRecording", transcript["detail"])

	def test_a_recording_403_names_the_permission_before_the_licence(self):
		"""A tenant that cannot record has no recordings, not a 403 — so the licence is a footnote,
		and leading with it would send people to the wrong page."""
		result = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings/MSp/recordings/7e31db25/content failed (403): "
			"Forbidden"
		)

		self.assertLess(
			result["detail"].index("OnlineMeetingRecording.Read.All"),
			result["detail"].index("Meeting recording"),
		)
		self.assertIn("never as a 403", result["detail"])

	def test_the_join_link_lookup_names_the_permission_people_miss(self):
		"""A key to a door you cannot walk to: they grant the transcript permission and stop."""
		result = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings?$filter=JoinWebUrl eq "
			"'https://teams.microsoft.com/l/meetup-join/19%3ameeting_Yzg5@thread.v2/0' failed "
			"(403): Forbidden"
		)

		self.assertTrue(result["matched"])
		self.assertIn("OnlineMeetings.ReadWrite", result["detail"])

	def test_a_standalone_meeting_is_told_it_can_never_have_a_transcript(self):
		"""Not 'not yet' — Graph never serves transcripts for a meeting with no calendar event."""
		result = doctor.explain_error(
			"Standalone meeting is not calendar-associated; transcripts are unavailable."
		)

		self.assertTrue(result["matched"])
		self.assertIn("never", result["detail"])
		self.assertIn("Add Teams meeting", result["detail"])

	def test_an_aged_out_meeting_is_not_blamed_on_permissions(self):
		"""Graph stops serving a meeting's artifacts ~60 days on, and says only NotFound."""
		result = doctor.explain_error(
			"Microsoft could not match this join link to a meeting — either another organisation "
			"hosted it, or it has aged out."
		)

		self.assertTrue(result["matched"])
		self.assertIn("60 days", result["detail"])
		self.assertNotIn("permission", result["detail"].lower())

	def test_a_404_on_a_meeting_path_lands_on_the_same_answer(self):
		result = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings/MSpiOTM1/transcripts failed (404): NotFound"
		)

		self.assertTrue(result["matched"])
		self.assertIn("no longer has this meeting", result["title"])

	def test_a_deleted_outlook_event_is_named_as_deleted(self):
		result = doctor.explain_error(
			"Microsoft Graph PATCH /me/events/AAMkAGI2 failed (404): ErrorItemNotFound: The "
			"specified object was not found in the store."
		)

		self.assertTrue(result["matched"])
		self.assertIn("no longer exists", result["title"])

	def test_a_mailbox_refusal_points_at_the_calendar_permission(self):
		result = doctor.explain_error(
			"Microsoft Graph GET /me/calendarView failed (403): ErrorAccessDenied: Access is "
			"denied. Check credentials and try again."
		)

		self.assertTrue(result["matched"])
		self.assertIn("Calendars.ReadWrite", result["detail"])

	def test_throttling_reads_as_temporary_rather_than_broken(self):
		"""The app already backs off; a 429 in the log is not something to go and fix."""
		for text in (
			"Microsoft Graph rate limit hit (429). The next scheduled sync will retry.",
			"Microsoft Graph GET /me/events failed (429): TooManyRequests: Application is over "
			"its MailboxConcurrency limit.",
		):
			with self.subTest(text=text):
				result = doctor.explain_error(text)
				self.assertTrue(result["matched"])
				self.assertIn("throttling", result["title"].lower())
				self.assertIn("Nothing is misconfigured", result["detail"])

	def test_a_mailbox_with_no_exchange_online_licence(self):
		result = doctor.explain_error(
			"Microsoft Graph POST /me/sendMail failed (403): MailboxNotEnabledForRESTAPI: REST "
			"API is not yet supported for this mailbox."
		)

		self.assertTrue(result["matched"])
		self.assertIn("Licenses and apps", result["detail"])

	def test_the_1406_that_filled_the_error_log_on_every_sign_in(self):
		"""A real Graph calendar id is 152 characters; the column was Frappe's default 140, and
		the write was lost silently while sign-in reported success."""
		result = doctor.explain_error(
			"(1406, \"Data too long for column 'ms_calendar_id' at row 1\")"
		)

		self.assertTrue(result["matched"])
		self.assertIn("migrate", result["detail"].lower())

	def test_an_expired_delta_watermark_is_not_reported_as_a_fault(self):
		result = doctor.explain_error(
			"Microsoft Graph sync state expired (syncStateNotFound: The sync state generation is "
			"not found). A full re-sync is required."
		)

		self.assertTrue(result["matched"])
		self.assertIn("nothing to fix", result["detail"].lower())

	def test_a_tenant_that_does_not_exist(self):
		result = doctor.explain_error("AADSTS90002: Tenant 'bizmap' not found.")

		self.assertTrue(result["matched"])
		self.assertIn("Directory (tenant) ID", result["detail"])

	def test_an_unrecognised_teams_error_still_falls_through_to_the_honest_answer(self):
		"""Twenty-five patterns is not omniscience, and pretending otherwise costs trust."""
		result = doctor.explain_error(
			"Microsoft Graph GET /me/onlineMeetings/MSp/attendanceReports failed (500): "
			"UnknownError"
		)

		self.assertFalse(result["matched"])
		self.assertIn("Unrecognised", result["title"])

	def test_every_pattern_compiles_and_says_something(self):
		import re

		for pattern, title, detail in doctor.error_patterns():
			with self.subTest(pattern=pattern):
				re.compile(pattern)
				self.assertTrue(title)
				self.assertTrue(detail)


class TestManualSetupSteps(BaseTestCase):
	"""The half of the setup that happens in Microsoft's portals and cannot be read from here.

	These must read as "confirm this", never as "this is broken": we genuinely do not know, and
	telling someone their finished step is broken is the same loop by another route.
	"""

	def test_the_tenant_switch_step_names_the_teams_admin_center_path(self):
		step = next(
			s for s in doctor.manual_setup_steps() if s["id"] == "transcript_api_access"
		)

		self.assertIn("Teams admin center", step["where"])
		self.assertIn("Meetings > Meeting settings > Transcript API access", step["where"])
		self.assertIn("Microsoft Graph access On", step["where"])

	def test_the_consent_step_names_the_entra_path_and_the_button(self):
		step = next(s for s in doctor.manual_setup_steps() if s["id"] == "admin_consent")

		self.assertIn("Microsoft Entra admin center", step["where"])
		self.assertIn("Grant admin consent", step["where"])
		self.assertIn("Re-authorize", step["verify"])

	def test_the_recording_step_names_the_policy_and_the_licence(self):
		"""A permission alone does not make a recording exist; a licence that records does."""
		step = next(s for s in doctor.manual_setup_steps() if s["id"] == "recording_allowed")

		self.assertIn("Meeting policies", step["where"])
		self.assertIn("licence", step["where"])

	def test_every_step_says_where_it_is_what_it_unlocks_and_how_to_tell(self):
		"""The contract the Settings form renders; a step missing one of these is unusable."""
		for step in doctor.manual_setup_steps():
			with self.subTest(step=step["id"]):
				for field in ("title", "where", "unlocks", "verify"):
					self.assertTrue(step.get(field), f"{step['id']} has no {field}")

	def test_only_the_capabilities_this_site_ticked_are_listed(self):
		"""A calendar-only site being told to flip Teams switches is noise, and noise gets skipped."""
		calendar_only = doctor.manual_setup_steps({"use_calendar": 1})

		self.assertEqual([s["id"] for s in calendar_only], ["admin_consent"])

	def test_transcripts_bring_the_teams_steps_with_them(self):
		ids = [s["id"] for s in doctor.manual_setup_steps({"use_transcripts": 1})]

		self.assertIn("transcript_api_access", ids)
		self.assertIn("recording_allowed", ids)

	def test_a_site_with_nothing_ticked_is_told_nothing(self):
		self.assertEqual(doctor.manual_setup_steps({}), [])

	def test_they_are_notes_never_problems(self):
		"""SKIP, because nothing was measured. A FAIL here would be a verdict on an unread setting."""
		for item in doctor.manual_step_findings({"use_transcripts": 1}):
			with self.subTest(check=item["check"]):
				self.assertEqual(item["status"], SKIP)
				self.assertTrue(item["check"].startswith("manual."))
				self.assertIn("Not checked", item["detail"])

	def test_the_finding_carries_the_portal_path_and_the_proof(self):
		item = next(
			f
			for f in doctor.manual_step_findings({"use_transcripts": 1})
			if f["check"] == "manual.transcript_api_access"
		)

		self.assertIn("Transcript API access", item["fix"])
		self.assertIn("You will know it worked", item["fix"])
		self.assertIn("microsoftteams", item["doc"])


class TestSetupGuideEndpoint(BaseTestCase):
	def test_it_returns_the_shape_the_form_consumes(self):
		result = doctor.setup_guide()

		self.assertIn("steps", result)
		self.assertIn("capabilities", result)
		for step in result["steps"]:
			for field in ("title", "where", "unlocks", "verify"):
				self.assertIn(field, step)
		for capability in result["capabilities"]:
			for field in ("id", "label", "enabled", "permissions"):
				self.assertIn(field, capability)

	def test_the_capabilities_carry_the_permissions_sign_in_will_request(self):
		"""Derived from the same table get_scopes uses, so the guide cannot advertise a
		permission the sign-in never asks for."""
		from frappe_microsoft365 import microsoft_graph as graph

		by_field = {field: list(scopes) for field, scopes in graph.CAPABILITY_SCOPES}

		for capability in doctor.setup_guide()["capabilities"]:
			with self.subTest(capability=capability["id"]):
				self.assertEqual(capability["permissions"], by_field[capability["field"]])

	def test_only_the_capabilities_this_site_ticked_are_described(self):
		"""The dialog prints these as "Covers:", so an unticked one there describes somebody
		else's setup."""
		summary = doctor._capability_summary({"use_transcripts": 1})

		self.assertEqual([c["id"] for c in summary], ["transcripts"])
		self.assertEqual(doctor._capability_summary({}), [])

	def test_it_is_system_manager_only(self):
		"""It names tenants, portals and what is not yet granted; that is not for every user.

		Frappe 15's ``only_for`` returns early whenever ``frappe.flags.in_test`` is set, so on
		that version the guard cannot fire under a test runner at all — the assertion below
		would pass on 16 and fail on 15 while the code was identical and correct. The flag is
		dropped for the length of the call so both versions actually exercise the guard, rather
		than skipping the test on the version that cannot demonstrate it.
		"""
		import frappe

		in_test = frappe.flags.in_test
		frappe.flags.in_test = False
		self.addCleanup(setattr, frappe.flags, "in_test", in_test)

		frappe.set_user("Guest")
		self.addCleanup(frappe.set_user, "Administrator")

		with self.assertRaises(frappe.PermissionError):
			doctor.setup_guide()


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

	def test_the_script_ships_with_its_own_undo(self):
		"""Granting standing mailbox access is reversible; the reverse belongs in the output."""
		script = doctor.powershell_for_app_only("c", mailboxes=["a@b.com"], send_as=True)

		self.assertIn("TO UNDO", script)
		self.assertIn("Remove-MailboxPermission", script)
		self.assertIn("Remove-RecipientPermission", script)
		self.assertIn("Remove-ServicePrincipal", script)

	def test_undo_block_can_be_suppressed(self):
		script = doctor.powershell_for_app_only("c", mailboxes=["a@b.com"], include_undo=False)

		self.assertNotIn("TO UNDO", script)
		self.assertIn("Add-MailboxPermission", script)

	def test_undo_lines_are_commented_so_nothing_runs_by_accident(self):
		script = doctor.powershell_for_app_only("c", mailboxes=["a@b.com"])
		undo = script[script.index("TO UNDO"):]

		for line in undo.splitlines():
			if "Remove-" in line:
				self.assertTrue(line.strip().startswith("#"), f"undo line must be commented: {line}")

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

	def test_the_manual_steps_are_reported_alongside_the_automatic_checks(self):
		"""Run Diagnostics is the button people press when stuck, and three of the four faults a
		real tenant hit were invisible to every automatic check in this module."""
		from unittest.mock import patch

		settings = {
			"enabled": 1,
			"tenant_id": TENANT,
			"client_id": "038b9c8e-2699-4858-b5f1-4f7a3d4077c4",
			"has_client_secret": True,
			"redirect_uri": f"https://site.example.com/api/method/{doctor.CALLBACK_METHOD}",
			"use_transcripts": 1,
			"default_scopes": "",
			"authorized_scopes": "",
			"mail_flow": "Delegated",
		}

		with patch.object(doctor, "_settings_config", return_value=settings):
			result = doctor.run_diagnostics()

		self.assertIn("manual.transcript_api_access", ids(result["findings"]))
		self.assertIn("manual.admin_consent", ids(result["findings"]))

	def test_a_manual_step_never_counts_as_a_problem(self):
		"""They are unread settings, not failing ones. Counting them would cry wolf every run."""
		for item in doctor.run_diagnostics()["findings"]:
			if item["check"].startswith("manual."):
				self.assertEqual(item["status"], SKIP)
