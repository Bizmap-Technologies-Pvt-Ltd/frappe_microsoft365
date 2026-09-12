"""One-click setup for the Microsoft capabilities you actually want.

A tenant admin registers ONE Azure application, but Frappe then needs it configured in up to
three unrelated places — this app's Microsoft Settings, a `Connected App` for mail, and a
`Social Login Key` for sign-in — each with different endpoints and scope strings, none of
which tell you when they are wrong. That is the fifteen-step setup people complain about.

This module builds those records from the single Azure registration already in Microsoft
Settings. Two rules govern it:

**Modular.** Calendar, mail and sign-in are independent. Take one, two or all three; nothing
is created for a capability you did not ask for, and none of them depend on each other.

**Never destructive.** Provisioning only ever CREATES missing records. If something already
exists it is left untouched and reported, with any drift from Microsoft Settings spelled out,
so a working mail setup someone tuned by hand is never rewritten underneath them.

The value builders are pure functions so every endpoint and scope string is testable without
a tenant, a network or a site.
"""

import json

import frappe
from frappe import _
from frappe.utils.password import get_decrypted_password

from frappe_microsoft365 import doctor
from frappe_microsoft365.doctor import APP_ONLY_SCOPE, DELEGATED_SCOPES, OFFLINE_ACCESS

#: The Connected App we create for mail. Named so it is obvious where it came from.
MAIL_APP_NAME = "Microsoft 365 Mail"
SSO_KEY_NAME = "Office 365"

CREATE, EXISTS, DRIFT, SKIP = "create", "exists", "drift", "skip"


# --- pure value builders --------------------------------------------------------------

def authority(tenant_id):
	return f"https://login.microsoftonline.com/{(tenant_id or 'common').strip()}"


def connected_app_values(settings, flow="Delegated"):
	"""Field values for the mail Connected App.

	Delegated mail needs the Outlook resource URLs plus offline_access — without the latter
	Microsoft returns no refresh token and the connection dies after the first access token
	expires. Application (client credentials) mail takes exactly one scope, `.default`.
	"""
	tenant = (settings.get("tenant_id") or "").strip()
	base = authority(tenant)

	if flow == "Application":
		scopes = [APP_ONLY_SCOPE]
	else:
		scopes = [
			DELEGATED_SCOPES["imap"],
			DELEGATED_SCOPES["smtp"],
			OFFLINE_ACCESS,
		]

	return {
		"doctype": "Connected App",
		"provider_name": MAIL_APP_NAME,
		"client_id": (settings.get("client_id") or "").strip(),
		"client_secret": settings.get("client_secret") or "",
		# No redirect_uri: Connected App.validate() computes it from the record name.
		"authorization_uri": f"{base}/oauth2/v2.0/authorize",
		"token_uri": f"{base}/oauth2/v2.0/token",
		"scopes": scopes,
	}


def social_login_key_values(settings):
	"""Field values for Microsoft sign-in.

	Frappe's built-in Office 365 provider defaults to the `/common/` authority and the v1.0
	endpoints. Neither works for a single-tenant app registration, which is what Microsoft
	Settings holds, so the tenant-specific v2.0 endpoints are written explicitly.
	"""
	tenant = (settings.get("tenant_id") or "").strip()
	base = authority(tenant)
	secret = settings.get("client_secret") or ""

	return {
		"doctype": "Social Login Key",
		"social_login_provider": "Office 365",
		"provider_name": SSO_KEY_NAME,
		# Frappe throws ClientSecretNotSetError if sign-in is enabled with no secret.
		"enable_social_login": 1 if secret else 0,
		"client_id": (settings.get("client_id") or "").strip(),
		"client_secret": secret,
		"base_url": "https://login.microsoftonline.com",
		"custom_base_url": 0,
		"icon": "fa fa-windows",
		"authorize_url": f"{base}/oauth2/v2.0/authorize",
		"access_token_url": f"{base}/oauth2/v2.0/token",
		"redirect_url": "/api/method/frappe.integrations.oauth2_logins.login_via_office365",
		"auth_url_data": json.dumps({"response_type": "code", "scope": "openid email profile"}),
		"sign_ups": "Deny",
	}


# --- capability registry --------------------------------------------------------------

def capabilities():
	"""What each capability needs from Azure, and what this app will create for it."""
	return [
		{
			"id": "calendar",
			"label": _("Outlook calendar and Teams meetings"),
			"creates": _("Nothing — this app talks to Graph directly."),
			"azure": [
				"User.Read",
				"Calendars.ReadWrite",
				"OnlineMeetings.ReadWrite",
				"OnlineMeetingTranscript.Read.All",
				"offline_access",
			],
			"azure_type": _("Microsoft Graph, Delegated"),
			"note": _("Each user authorises their own Microsoft Calendar."),
		},
		{
			"id": "mail",
			"label": _("Outlook mail"),
			"creates": _("A Connected App named '{0}' for Frappe's Email Account.").format(MAIL_APP_NAME),
			"azure": [
				DELEGATED_SCOPES["imap"],
				DELEGATED_SCOPES["smtp"],
				"offline_access",
			],
			"azure_type": _("Office 365 Exchange Online, Delegated"),
			"note": _(
				"Mail itself stays with Frappe's Email Account. For shared mailboxes choose the "
				"Application method, which also needs IMAP.AccessAsApp / SMTP.SendAsApp plus the "
				"Exchange setup script."
			),
		},
		{
			"id": "sso",
			"label": _("Sign in with Microsoft"),
			"creates": _("A Social Login Key named '{0}'.").format(SSO_KEY_NAME),
			"azure": ["openid", "email", "profile"],
			"azure_type": _("Microsoft Graph, Delegated"),
			"note": _(
				"New visitors cannot self-register by default; existing users with a matching "
				"email can sign in."
			),
		},
	]


def selected_capabilities(settings):
	"""The capability ids ticked in Microsoft Settings."""
	chosen = []
	if settings.get("use_calendar"):
		chosen.append("calendar")
	if settings.get("use_mail"):
		chosen.append("mail")
	if settings.get("use_sso"):
		chosen.append("sso")
	return chosen


# --- plan / apply ---------------------------------------------------------------------

def _step(capability, action, target, detail, findings=None):
	return {
		"capability": capability,
		"action": action,
		"target": target,
		"detail": detail,
		"findings": findings or [],
	}


def _settings():
	doc = frappe.get_cached_doc("Microsoft Settings")
	return {
		"client_secret": get_decrypted_password(
			"Microsoft Settings", "Microsoft Settings", "client_secret", raise_exception=False
		)
		or "",
		"enabled": doc.enabled,
		"tenant_id": doc.tenant_id,
		"client_id": doc.client_id,
		"redirect_uri": doc.redirect_uri or "",
		"use_calendar": doc.use_calendar,
		"use_mail": doc.use_mail,
		"use_sso": doc.use_sso,
		"mail_flow": doc.mail_flow or "Delegated",
	}


def _plan_mail(settings):
	flow = settings.get("mail_flow") or "Delegated"
	existing = frappe.db.exists("Connected App", {"provider_name": MAIL_APP_NAME})

	if not existing:
		return _step(
			"mail",
			CREATE,
			MAIL_APP_NAME,
			_("Create a Connected App for {0} mail sign-in.").format(flow.lower()),
		)

	# Exists: report, never rewrite.
	current = doctor._connected_app_config(existing)
	findings = doctor.check_connected_app(current, settings, app_only=(flow == "Application"))
	problems = [f for f in findings if f["status"] in (doctor.FAIL, doctor.WARN)]
	if problems:
		return _step(
			"mail",
			DRIFT,
			existing,
			_("Already exists and does not match Microsoft Settings. Left untouched."),
			problems,
		)
	return _step("mail", EXISTS, existing, _("Already set up correctly. Nothing to do."))


def _plan_sso(settings):
	existing = frappe.db.exists("Social Login Key", {"social_login_provider": "Office 365"})
	if not existing:
		return _step("sso", CREATE, SSO_KEY_NAME, _("Create a Social Login Key for Microsoft sign-in."))

	wanted = social_login_key_values(settings)
	doc = frappe.get_doc("Social Login Key", existing)
	drift = []
	for field in ("authorize_url", "access_token_url", "client_id"):
		if (doc.get(field) or "") != wanted[field]:
			drift.append(
				doctor.finding(
					f"sso.{field}",
					doctor.WARN,
					_("Sign-in {0} differs from Microsoft Settings").format(field),
					_("Currently {0}").format(doc.get(field) or _("(empty)")),
					_("Expected {0}").format(wanted[field]),
					target=existing,
				)
			)
	if drift:
		return _step("sso", DRIFT, existing, _("Already exists and differs. Left untouched."), drift)
	return _step("sso", EXISTS, existing, _("Already set up correctly. Nothing to do."))


@frappe.whitelist()
def plan():
	"""Dry run: what would be created, what exists, what has drifted. Writes nothing."""
	frappe.only_for("System Manager")
	settings = _settings()
	chosen = selected_capabilities(settings)
	steps = []

	for capability in capabilities():
		if capability["id"] not in chosen:
			steps.append(
				_step(capability["id"], SKIP, "", _("Not selected in Microsoft Settings."))
			)
			continue

		if capability["id"] == "calendar":
			steps.append(
				_step(
					"calendar",
					EXISTS,
					"",
					_("Nothing to create. Grant the Azure permissions listed, then add a Microsoft Calendar."),
				)
			)
		elif capability["id"] == "mail":
			steps.append(_plan_mail(settings))
		elif capability["id"] == "sso":
			steps.append(_plan_sso(settings))

	blockers = []
	if not (settings.get("tenant_id") or "").strip() or not (settings.get("client_id") or "").strip():
		blockers.append(_("Tenant ID and Client ID must be filled in before anything can be created."))
	if not settings.get("client_secret") and ("mail" in chosen or "sso" in chosen):
		blockers.append(
			_("A Client Secret is required for mail and sign-in. Add it in Microsoft Settings first.")
		)

	return {
		"steps": steps,
		"capabilities": capabilities(),
		"selected": chosen,
		"blockers": blockers,
		"mail_flow": settings.get("mail_flow"),
	}


def _create_connected_app(settings):
	values = connected_app_values(settings, settings.get("mail_flow") or "Delegated")
	scopes = values.pop("scopes")
	doc = frappe.get_doc(values)
	for scope in scopes:
		doc.append("scopes", {"scope": scope})
	doc.insert(ignore_permissions=True)
	return doc.name


def _create_social_login_key(settings):
	doc = frappe.get_doc(social_login_key_values(settings))
	doc.insert(ignore_permissions=True)
	return doc.name


@frappe.whitelist()
def apply():
	"""Create whatever the plan says is missing. Never edits or deletes anything."""
	frappe.only_for("System Manager")
	settings = _settings()
	result = plan()
	if result["blockers"]:
		frappe.throw("<br>".join(result["blockers"]))

	created = []
	for step in result["steps"]:
		if step["action"] != CREATE:
			continue
		if step["capability"] == "mail":
			created.append({"capability": "mail", "name": _create_connected_app(settings)})
		elif step["capability"] == "sso":
			created.append({"capability": "sso", "name": _create_social_login_key(settings)})

	if created:
		frappe.db.commit()

	# Re-plan so the caller sees the settled state, and run the doctor over it.
	return {"created": created, "plan": plan(), "diagnostics": doctor.run_diagnostics()}
