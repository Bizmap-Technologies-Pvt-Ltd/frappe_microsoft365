"""Microsoft connection doctor.

Connecting Frappe to Microsoft 365 fails in a dozen ways that all surface as the same
unhelpful string — ``AUTHENTICATE failed``, ``535 5.7.3``, ``invalid_grant`` — with no
indication which of the fifteen setup steps was wrong. This module inspects the
configuration and says what is actually wrong, in the admin's language.

It only ever READS configuration. It does not touch mail sending or receiving, the Email
Account doctype's behaviour, or the email queue: Frappe's own IMAP/SMTP + OAuth path stays
exactly as it is, and several mailboxes keep working the way they always did.

Every rule below is derived from a documented requirement, cited inline:

* Microsoft, "Authenticate an IMAP, POP or SMTP connection using OAuth"
  https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth
* Microsoft Entra authentication error codes
  https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes

The check functions take plain dicts rather than Documents so they stay pure and fully
testable offline. ``run_diagnostics`` is the thin layer that loads real records into them.
"""

import re

import frappe
from frappe import _
from frappe.utils.password import get_decrypted_password

from frappe_microsoft365 import background
from frappe_microsoft365 import microsoft_graph as graph

# microsoft_graph does not import this module, so this direction is safe.
from frappe_microsoft365.microsoft_graph import CALLBACK_METHOD

# --- documented constants -------------------------------------------------------------

#: Delegated (a user signs in) scopes, per Microsoft's protocol table.
DELEGATED_SCOPES = {
	"imap": "https://outlook.office.com/IMAP.AccessAsUser.All",
	"pop": "https://outlook.office.com/POP.AccessAsUser.All",
	"smtp": "https://outlook.office.com/SMTP.Send",
}

#: App-only (client credentials) token scope. Microsoft: "You must use
#: https://outlook.office365.com/.default in the scope property in the body payload".
APP_ONLY_SCOPE = "https://outlook.office365.com/.default"

#: Admin consent for POP/IMAP application permissions uses a DIFFERENT scope than SMTP.
ADMIN_CONSENT_SCOPE_POP_IMAP = "https://ps.outlook.com/.default"
ADMIN_CONSENT_SCOPE_SMTP = "https://outlook.office365.com/.default"

OFFLINE_ACCESS = "offline_access"

EXCHANGE_IMAP_HOST = "outlook.office365.com"
EXCHANGE_SMTP_HOST = "smtp.office365.com"

MS_OAUTH_DOC = (
	"https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/"
	"how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth"
)
ENTRA_ERROR_DOC = "https://learn.microsoft.com/en-us/entra/identity-platform/reference-error-codes"

PASS, WARN, FAIL, SKIP = "pass", "warn", "fail", "skip"


def finding(check, status, title, detail="", fix="", doc="", target=None):
	"""One diagnostic result. Kept as a plain dict so it crosses the API boundary cleanly."""
	return {
		"check": check,
		"status": status,
		"title": title,
		"detail": detail,
		"fix": fix,
		"doc": doc,
		"target": target,
	}


def _norm_scopes(scopes):
	return [s.strip() for s in (scopes or []) if s and s.strip()]


def _tenant_of(uri):
	"""Pull the tenant segment out of a login.microsoftonline.com URI, if present."""
	match = re.search(r"login\.microsoftonline\.com/([^/]+)/", uri or "")
	return match.group(1) if match else None


# --- Microsoft Settings ---------------------------------------------------------------

def check_settings(settings):
	"""``settings``: dict of enabled, tenant_id, client_id, has_client_secret, redirect_uri."""
	out = []
	tenant = (settings.get("tenant_id") or "").strip()

	if not settings.get("enabled"):
		out.append(
			finding(
				"settings.enabled",
				WARN,
				_("Microsoft 365 integration is disabled"),
				_("Live operations will refuse to run while this is off."),
				_("Tick Enabled in Microsoft Settings once the rest of the setup checks out."),
			)
		)

	missing = [
		label
		for label, present in (
			("Tenant ID", tenant),
			("Client ID", (settings.get("client_id") or "").strip()),
			("Client Secret", settings.get("has_client_secret")),
		)
		if not present
	]
	if missing:
		out.append(
			finding(
				"settings.credentials",
				FAIL,
				_("Azure application details are incomplete"),
				_("Missing: {0}.").format(", ".join(missing)),
				_("Copy them from the Azure app registration into Microsoft Settings."),
			)
		)

	# Client credentials (app-only) cannot use the multi-tenant /common authority.
	if tenant.lower() in ("common", "organizations", "consumers"):
		out.append(
			finding(
				"settings.tenant_id",
				WARN,
				_("Tenant is set to '{0}'").format(tenant),
				_(
					"App-only (client credentials) access requires a specific tenant id. "
					"'{0}' works for interactive sign-in only."
				).format(tenant),
				_("Use the Directory (tenant) ID from the Azure app registration Overview page."),
				MS_OAUTH_DOC,
			)
		)

	redirect = (settings.get("redirect_uri") or "").strip()
	if redirect and not (redirect.startswith("https://") or "localhost" in redirect):
		out.append(
			finding(
				"settings.redirect_uri",
				FAIL,
				_("Redirect URI is not HTTPS"),
				_("Azure rejects plain HTTP redirect URIs except on localhost."),
				_("Serve the site over HTTPS, or test on http://...localhost."),
			)
		)

	return out


# --- delegated scopes -----------------------------------------------------------------

def check_scopes(settings):
	"""Is the Graph scope list going to ask for what the ticked capabilities need?

	``settings``: dict of the capability ticks plus ``default_scopes`` — the override field,
	which keeps its original fieldname so no site loses the value it already had.

	Scopes are derived from the capabilities, so the only way to end up asking for the wrong
	thing is the override, which exists for tenants that consent to a hand-picked list. An
	override that is missing something stays silent until the feature is used and Graph answers
	403, so it is worth saying out loud here.
	"""
	out = []
	override = graph.parse_scopes(settings.get("default_scopes"))
	if not override:
		return out

	reserved = [s for s in override if s in graph.RESERVED_SCOPES]
	if reserved:
		# MSAL rejects the reserved scopes in its `scopes` argument, so listing them here does
		# not widen consent — it breaks the token calls outright.
		out.append(
			finding(
				"scopes.override_reserved",
				FAIL,
				_("The scope override lists a reserved scope"),
				_("Found: {0}").format(", ".join(reserved)),
				_(
					"Remove them. {0} are requested automatically at sign-in and are refused when "
					"passed explicitly."
				).format(", ".join(graph.RESERVED_SCOPES)),
			)
		)

	missing = [s for s in graph.derive_scopes(settings) if s not in override]
	if missing:
		out.append(
			finding(
				"scopes.override_incomplete",
				WARN,
				_("The scope override is missing {0}").format(", ".join(missing)),
				_(
					"The capabilities ticked in Microsoft Settings need {0}, but the override asks "
					"only for {1}. Whatever is missing fails with a 403 the first time it is used."
				).format(", ".join(graph.derive_scopes(settings)), ", ".join(override)),
				_(
					"Add the missing scope to the override, or clear the override entirely and let "
					"the capabilities decide."
				),
			)
		)

	return out


def check_authorized_scopes(requested, authorized, connections):
	"""Have the scopes changed since the existing connections signed in?

	``requested`` is what sign-in asks for now, ``authorized`` what it asked for at the last
	successful authorisation, ``connections`` how many Microsoft Calendars are authorised.

	A token carries the permissions consented when it was issued. Ticking another capability
	does not widen a token that already exists, so the new feature fails with a 403 that names
	nothing until every connection has been re-authorised.

	If nothing was recorded — every connection predates this check — say nothing rather than
	invent a comparison.
	"""
	if not authorized or not connections:
		return []

	added = sorted(set(requested) - set(authorized))
	removed = sorted(set(authorized) - set(requested))
	if not added and not removed:
		return []

	changes = []
	if added:
		changes.append(_("now also asking for {0}").format(", ".join(added)))
	if removed:
		changes.append(_("no longer asking for {0}").format(", ".join(removed)))

	return [
		finding(
			"scopes.changed_since_authorization",
			WARN,
			_("The scopes changed after {0} connection(s) were authorised").format(connections),
			_("Last authorised with {0}; {1}.").format(", ".join(authorized), "; ".join(changes)),
			_(
				"Open each Microsoft Calendar and click Re-authorize. Existing tokens keep the "
				"permissions they were issued with, so they will not pick this up on their own."
			),
		)
	]


#: Azure identifiers are GUIDs. A secret VALUE never is, which is what makes the mix-up
#: detectable: the Secret ID sitting next to it in the same table is.
GUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)

#: Frappe shows a stored Password field as a mask rather than the value.
MASKED = re.compile(r"^\*+$")


def looks_like_guid(value):
	return bool(GUID.match((value or "").strip()))


def check_credentials(settings):
	"""Catch the Azure values people actually paste into the wrong box.

	Azure's Certificates & secrets table puts the secret **Value** next to its **Secret ID**,
	and only the Value works. Copying the wrong one is the single most common setup mistake,
	and Microsoft answers it with AADSTS7000215 long after the fact. The two are trivially
	distinguishable: a Secret ID is a GUID, a secret value never is.
	"""
	out = []
	secret = (settings.get("client_secret") or "").strip()

	if secret and not MASKED.match(secret) and looks_like_guid(secret):
		out.append(
			finding(
				"settings.secret_is_the_id",
				FAIL,
				_("The Client Secret looks like the Secret ID"),
				_("A GUID was entered, and a secret value is never a GUID."),
				_(
					"In Azure, Certificates & secrets shows Value next to Secret ID. Copy the "
					"Value column. It is only visible immediately after you create the secret, "
					"so if you have navigated away, create a new one."
				),
				MS_OAUTH_DOC,
			)
		)

	client_id = (settings.get("client_id") or "").strip()
	if client_id and not looks_like_guid(client_id):
		out.append(
			finding(
				"settings.client_id_shape",
				WARN,
				_("The Client ID is not a GUID"),
				_("Azure's Application (client) ID always is."),
				_("Copy Application (client) ID from the app registration Overview page."),
			)
		)

	tenant = (settings.get("tenant_id") or "").strip()
	if (
		tenant
		and not looks_like_guid(tenant)
		and tenant.lower() not in ("common", "organizations", "consumers")
		and "." not in tenant
	):
		out.append(
			finding(
				"settings.tenant_id_shape",
				WARN,
				_("The Tenant ID is neither a GUID nor a domain"),
				_("Found: {0}").format(tenant),
				_("Use Directory (tenant) ID from the Overview page, or your tenant's domain."),
			)
		)

	redirect = (settings.get("redirect_uri") or "").strip()
	if redirect and CALLBACK_METHOD.rsplit(".", 1)[-1] not in redirect:
		out.append(
			finding(
				"settings.redirect_target",
				WARN,
				_("The Redirect URI does not point at this app's callback"),
				_("Found: {0}").format(redirect),
				_("It should end with /api/method/{0}").format(CALLBACK_METHOD),
			)
		)

	return out


# --- Connected App --------------------------------------------------------------------

def check_connected_app(app, settings=None, app_only=False):
	"""``app``: dict of name, client_id, redirect_uri, authorization_uri, token_uri, scopes."""
	out = []
	settings = settings or {}
	name = app.get("name")
	scopes = _norm_scopes(app.get("scopes"))

	if not scopes:
		out.append(
			finding(
				"connected_app.scopes",
				FAIL,
				_("Connected App has no scopes"),
				_("Without scopes Microsoft issues a token that mail protocols will reject."),
				_("Add {0}.").format(APP_ONLY_SCOPE if app_only else ", ".join(DELEGATED_SCOPES.values())),
				MS_OAUTH_DOC,
				name,
			)
		)

	if app_only:
		# Microsoft: app-only tokens must be requested at the .default scope.
		if scopes and APP_ONLY_SCOPE not in scopes:
			out.append(
				finding(
					"connected_app.scopes_app_only",
					FAIL,
					_("App-only flow is using delegated scopes"),
					_("Found: {0}").format(", ".join(scopes)),
					_("Replace them with the single scope {0}.").format(APP_ONLY_SCOPE),
					MS_OAUTH_DOC,
					name,
				)
			)
		elif scopes and len(scopes) > 1:
			out.append(
				finding(
					"connected_app.scopes_app_only_extra",
					WARN,
					_("Extra scopes alongside .default"),
					_("Client credentials ignores everything except the .default scope."),
					_("Leave only {0}.").format(APP_ONLY_SCOPE),
					MS_OAUTH_DOC,
					name,
				)
			)
	elif scopes:
		delegated = set(DELEGATED_SCOPES.values())
		if not (set(scopes) & delegated):
			out.append(
				finding(
					"connected_app.scopes_delegated",
					FAIL,
					_("No Outlook protocol scopes found"),
					_("Found: {0}").format(", ".join(scopes)),
					_(
						"Delegated mail needs the full resource URLs, e.g. {0} for IMAP and {1} "
						"for sending."
					).format(DELEGATED_SCOPES["imap"], DELEGATED_SCOPES["smtp"]),
					MS_OAUTH_DOC,
					name,
				)
			)
		if OFFLINE_ACCESS not in scopes:
			# Without offline_access there is no refresh token, so the connection dies
			# after the first access token expires — the "loses access after a few hours"
			# symptom reported repeatedly on the forum.
			out.append(
				finding(
					"connected_app.offline_access",
					FAIL,
					_("offline_access scope is missing"),
					_(
						"Microsoft only returns a refresh token when offline_access is requested. "
						"Without it the account works for about an hour and then needs "
						"re-authorising, repeatedly."
					),
					_("Add the scope {0} to the Connected App.").format(OFFLINE_ACCESS),
					MS_OAUTH_DOC,
					name,
				)
			)

	token_uri = (app.get("token_uri") or "").strip()
	auth_uri = (app.get("authorization_uri") or "").strip()

	if token_uri and "/oauth2/v2.0/" not in token_uri:
		out.append(
			finding(
				"connected_app.endpoint_version",
				FAIL,
				_("Connected App points at the v1.0 token endpoint"),
				_("Token URI: {0}").format(token_uri),
				_("Use the v2.0 endpoints: .../oauth2/v2.0/token and .../oauth2/v2.0/authorize."),
				MS_OAUTH_DOC,
				name,
			)
		)

	tenant = (settings.get("tenant_id") or "").strip()
	for label, uri in (("authorization", auth_uri), ("token", token_uri)):
		uri_tenant = _tenant_of(uri)
		if tenant and uri_tenant and uri_tenant.lower() != tenant.lower():
			out.append(
				finding(
					f"connected_app.tenant_{label}",
					FAIL,
					_("Connected App {0} URI points at a different tenant").format(label),
					_("Microsoft Settings says {0}, the {1} URI says {2}.").format(
						tenant, label, uri_tenant
					),
					_("Point both URIs at https://login.microsoftonline.com/{0}/oauth2/v2.0/...").format(
						tenant
					),
					target=name,
				)
			)

	# Frappe COMPUTES this field in Connected App.validate() and it contains the record
	# name, so it is a different endpoint from this app's own callback and cannot be known
	# before the record exists. Azure therefore needs BOTH URIs registered — a trap that
	# surfaces later as AADSTS50011 and is documented nowhere.
	app_redirect = (app.get("redirect_uri") or "").strip()
	if not app_redirect:
		out.append(
			finding(
				"connected_app.redirect_uri",
				FAIL,
				_("Connected App has no redirect URI"),
				_("Frappe normally fills this in on save."),
				_("Re-save the Connected App."),
				target=name,
			)
		)
	else:
		out.append(
			finding(
				"connected_app.redirect_uri_registration",
				SKIP,
				_("Register this redirect URI in Azure"),
				_("Mail sign-in returns to {0}").format(app_redirect),
				_(
					"It is a different endpoint from this app's own callback, so Azure needs both "
					"registered under Web redirect URIs. A missing one fails with AADSTS50011."
				),
				ENTRA_ERROR_DOC,
				name,
			)
		)

	return out


# --- Email Account --------------------------------------------------------------------

def check_email_account(account):
	"""``account``: dict of the Email Account fields the OAuth path depends on."""
	out = []
	name = account.get("name") or account.get("email_id")

	if account.get("auth_method") != "OAuth":
		return [
				finding(
					"email_account.auth_method",
					SKIP,
					_("{0} uses Basic authentication").format(name),
					_(
						"Microsoft disables Basic auth for SMTP by default at the end of December "
						"2026, and it is unavailable for tenants created after that."
					),
					_("Move this account to OAuth before then."),
					MS_OAUTH_DOC,
					name,
				)
			]

	app_only = bool(account.get("backend_app_flow"))

	if not account.get("connected_app"):
		out.append(
			finding(
				"email_account.connected_app",
				FAIL,
				_("{0} is set to OAuth but has no Connected App").format(name),
				"",
				_("Link the Connected App that holds the Azure credentials."),
				target=name,
			)
		)

	if not app_only and not account.get("connected_user"):
		out.append(
			finding(
				"email_account.connected_user",
				FAIL,
				_("{0} has no Connected User").format(name),
				_("Delegated OAuth stores the token against a Frappe user; without one there is "
				  "no token to authenticate with."),
				_("Set Connected User, then click Authorize API Access while logged in as them."),
				target=name,
			)
		)

	# The shared-mailbox identity conflict. Microsoft: for shared mailbox access over
	# IMAP the XOAUTH2 user= field must be the SHARED mailbox address, while SMTP must
	# authenticate as the signing-in user. Frappe sends `login_id or email_id` for both,
	# so one field cannot satisfy both protocols on the same account.
	if (
		not app_only
		and account.get("enable_incoming")
		and account.get("enable_outgoing")
		and account.get("login_id_is_different")
		and (account.get("login_id") or "").strip()
		and (account.get("login_id") or "").strip().lower() != (account.get("email_id") or "").strip().lower()
	):
		out.append(
			finding(
				"email_account.shared_mailbox_identity",
				WARN,
				_("{0} sends one identity to both IMAP and SMTP").format(name),
				_(
					"For a shared mailbox Microsoft wants the mailbox address in the IMAP "
					"XOAUTH2 string but the signing-in user for SMTP. Frappe sends Login Id to "
					"both, so incoming and outgoing cannot both be right on this account."
				),
				_(
					"Either split it into two Email Accounts (one incoming, one outgoing), or "
					"switch to the app-only flow, where no user identity is involved."
				),
				MS_OAUTH_DOC,
				name,
			)
		)

	if account.get("enable_incoming") and account.get("use_imap") and not account.get("imap_folder"):
		out.append(
			finding(
				"email_account.imap_folder",
				FAIL,
				_("{0} has no IMAP folder configured").format(name),
				"",
				_("Add at least one folder (usually INBOX)."),
				target=name,
			)
		)

	if account.get("use_ssl") and account.get("use_starttls"):
		out.append(
			finding(
				"email_account.tls",
				WARN,
				_("{0} has both SSL and STARTTLS enabled").format(name),
				_("These are alternatives; enabling both is a common cause of 'TLS required'."),
				_("Use SSL for IMAP on 993, or STARTTLS for SMTP on 587 — not both."),
				target=name,
			)
		)

	server = (account.get("email_server") or "").strip().lower()
	if account.get("enable_incoming") and server and EXCHANGE_IMAP_HOST not in server:
		out.append(
			finding(
				"email_account.imap_host",
				WARN,
				_("{0} incoming server is {1}").format(name, server),
				_("Microsoft 365 mailboxes use {0}.").format(EXCHANGE_IMAP_HOST),
				_("Set the incoming server to {0}.").format(EXCHANGE_IMAP_HOST),
				target=name,
			)
		)

	return out


def check_social_login_key(key, settings=None):
	"""``key``: dict of enable_social_login, client_id, authorize_url, access_token_url."""
	out = []
	settings = settings or {}
	name = key.get("name") or "Office 365"
	tenant = (settings.get("tenant_id") or "").strip()

	if not key.get("enable_social_login"):
		out.append(
			finding(
				"sso.enabled",
				WARN,
				_("Microsoft sign-in is configured but switched off"),
				"",
				_("Tick Enable Social Login on the Social Login Key."),
				target=name,
			)
		)

	client_id = (settings.get("client_id") or "").strip()
	if client_id and (key.get("client_id") or "").strip() != client_id:
		out.append(
			finding(
				"sso.client_id",
				FAIL,
				_("Sign-in uses a different Client ID"),
				_("Microsoft Settings has {0}.").format(client_id),
				_("Point both at the same Azure app registration, or sign-in will fail."),
				target=name,
			)
		)

	for field, label in (("authorize_url", _("authorize")), ("access_token_url", _("token"))):
		uri = (key.get(field) or "").strip()
		if not uri:
			continue

		# Frappe's built-in Office 365 provider ships /common/ + v1.0 endpoints, which
		# cannot work with a single-tenant app registration.
		if "/common/" in uri and tenant and tenant.lower() not in ("common", "organizations"):
			out.append(
				finding(
					f"sso.tenant_{field}",
					FAIL,
					_("Sign-in {0} URL points at /common/").format(label),
					_("Your app is registered in tenant {0}, not the shared endpoint.").format(tenant),
					_("Use https://login.microsoftonline.com/{0}/oauth2/v2.0/...").format(tenant),
					target=name,
				)
			)
		if "/oauth2/v2.0/" not in uri:
			out.append(
				finding(
					f"sso.version_{field}",
					FAIL,
					_("Sign-in {0} URL uses the v1.0 endpoint").format(label),
					_("Currently {0}").format(uri),
					_("Use the v2.0 endpoint instead."),
					target=name,
				)
			)

	return out


#: Custom fields the sync reads and writes. Without them every query fails with a raw SQL
#: error ("Unknown column ... in WHERE"), which tells an admin nothing.
EVENT_CUSTOM_FIELDS = [
	"custom_sync_with_microsoft_calendar",
	"custom_microsoft_calendar",
	"custom_microsoft_event_id",
	"custom_pulled_from_microsoft",
]


def check_event_custom_fields(missing):
	"""``missing``: list of custom fieldnames absent from the Event doctype."""
	if not missing:
		return []
	return [
		finding(
			"fields.event",
			FAIL,
			_("Calendar sync fields are missing from Event"),
			_("Missing: {0}").format(", ".join(missing)),
			_(
				"Run `bench --site <site> migrate` (on Frappe Cloud, use Migrate in the site "
				"dashboard). Until then syncing fails with an Unknown column error."
			),
		)
	]


# --- background jobs ------------------------------------------------------------------

def _background_summary(state):
	"""The five readings, in one line, attached to whichever finding fires.

	They travel together on purpose: "no worker" and "203 jobs waiting" and "last ran never"
	are the same sentence told three ways, and an admin who sees one of them without the
	others tends to fix the wrong end of it.
	"""
	workers = state.get("workers")
	scheduler = state.get("scheduler_inactive")
	redis = state.get("redis")
	jobs = [
		_("{0} last ran {1} minutes ago").format(job["method"], job["minutes_ago"])
		if job["minutes_ago"] is not None
		else _("{0} has never run").format(job["method"])
		for job in state.get("jobs") or []
	]
	# None means the probe could not find out, which is a third answer and not a quiet yes.
	return _("Scheduler: {0}. Redis: {1}. Workers: {2}. Jobs waiting: {3}.{4}").format(
		_("unknown") if scheduler is None else (_("inactive") if scheduler else _("active")),
		_("unknown") if redis is None else (_("reachable") if redis else _("unreachable")),
		_("unknown") if workers is None else workers,
		state.get("queued", 0),
		(" " + "; ".join(jobs) + ".") if jobs else "",
	)


def check_background_jobs(state):
	"""``state``: the dict from ``background.health()``.

	Queued work that never runs is the one fault in this app that produces no error at all —
	no traceback, no log line, no wrong answer, just a calendar that quietly stops moving. The
	doctor exists to name the failing step instead of leaving people with a bare error, and
	this step does not even manage a bare error.

	Returns nothing on a healthy bench, like every other check here. The backlog line is the
	one thing it can say while ``ok`` is still True — a deep queue is not a fault, but it is
	worth knowing about before someone concludes their sync is broken.
	"""
	out = []
	detail = _background_summary(state)
	# Each finding carries the fix for its own cause, so a site with two faults gets two
	# commands rather than one paragraph containing both. Looked up by code, never by the
	# wording: the reasons are translated and the codes are not.
	fixes = {problem["code"]: problem["fix"] for problem in state.get("problems") or []}

	titles = {
		background.SCHEDULER_OFF: _("The scheduler is off, so nothing runs on a schedule"),
		background.REDIS_DOWN: _("Redis cannot be reached, so no job can be queued at all"),
		background.NO_WORKER: _("No background worker is running"),
		background.WORKERS_UNKNOWN: _("The number of running workers could not be read"),
		background.JOBS_MISSING: _("This app's scheduled jobs are not registered"),
		background.JOBS_STALE: _("This app's scheduled jobs have stopped running"),
		background.CHECK_FAILED: _("Whether background jobs run could not be established"),
	}
	details = {
		# Spelled out because this is the failure that looks like success: the queue accepts
		# everything and starts nothing.
		background.NO_WORKER: _("{0} Jobs are accepted and then sit in the queue forever.").format(detail),
		background.JOBS_STALE: _(
			"{0} They are registered to run every 15 minutes, so this is the evidence that survives "
			"when every other check looks fine."
		).format(detail),
	}
	# "could not be established" is not the same as "broken", and a FAIL for it would teach
	# people that this section cries wolf.
	severity = {background.WORKERS_UNKNOWN: WARN, background.CHECK_FAILED: WARN}

	for code, fix in fixes.items():
		out.append(
			finding(
				f"background.{code}",
				severity.get(code, FAIL),
				titles.get(code, _("Background jobs are not running")),
				details.get(code, detail),
				fix,
			)
		)

	if state.get("backlog"):
		out.append(
			finding(
				"background.backlog",
				WARN,
				_("{0} jobs are waiting in the queue").format(state.get("queued", 0)),
				detail,
				# purge-jobs takes its own --site rather than bench's, and without one it empties
				# the queue for every site on the bench.
				_("They drain once a worker runs. `bench purge-jobs --site <site>` drops them instead."),
			)
		)

	return out


# --- error decoder --------------------------------------------------------------------

def error_patterns():
	"""Ordered because some strings are substrings of others; first match wins.

	Built per call rather than held in a module-level constant: the messages go through
	frappe._(), which resolves against the CURRENT site and language. A global would freeze
	whichever site imported the module first and hand its translations to every other site.
	"""
	return [
	(
		r"AADSTS50011",
		_("Redirect URI mismatch"),
		_(
			"The redirect URI in the Connected App does not exactly match one registered on the "
			"Azure app — scheme, host, port and path all have to match."
		),
	),
	(
		r"AADSTS65001|consent",
		_("Admin consent has not been granted"),
		_(
			"An administrator must grant consent for the requested permissions. For POP/IMAP "
			"application permissions the consent URL uses scope "
			"https://ps.outlook.com/.default; for SMTP it uses "
			"https://outlook.office365.com/.default."
		),
	),
	(
		r"AADSTS7000215|invalid_client",
		_("Client secret is wrong or expired"),
		_("Azure client secrets expire. Create a new one and paste it into Microsoft Settings."),
	),
	(
		r"AADSTS700016|application with identifier",
		_("The application is not present in this tenant"),
		_("Check the tenant id, or have an admin consent the app into the tenant first."),
	),
	(
		r"invalid_grant",
		_("The refresh token is no longer valid"),
		_(
			"It expires after long inactivity, or when consent or the password changes. "
			"Re-authorise the account. If this recurs within hours, offline_access is probably "
			"missing from the Connected App scopes."
		),
	),
	(
		r"535 5\.7\.3|535 5\.7\.139|SMTPAuthenticationError",
		_("SMTP rejected the token"),
		_(
			"Usually SMTP AUTH is disabled for that mailbox in Exchange, or the identity in the "
			"XOAUTH2 string is not the user the token was issued for. For app-only sending, the "
			"service principal also needs SendAs via Add-RecipientPermission."
		),
	),
	(
		r"451 4\.7\.0",
		_("Exchange applied a temporary block"),
		_(
			"Usually throttling or a tenant-level restriction rather than a wrong credential. "
			"Retry, and confirm SMTP AUTH is enabled for the mailbox."
		),
	),
	(
		r"AUTHENTICATE failed|A01 NO",
		_("IMAP rejected the token"),
		_(
			"The token was issued but the mailbox refused it. Check that IMAP.AccessAsUser.All "
			"(delegated) or IMAP.AccessAsApp (app-only) is granted and consented, that the "
			"service principal is registered in Exchange, and that the mailbox itself was granted "
			"to it with Add-MailboxPermission."
		),
	),
	(
		r"TLS required|STARTTLS",
		_("TLS negotiation failed"),
		_("Check the SSL and STARTTLS flags — one or the other, not both."),
	),
	(
		r"ErrorPropertyValidationFailure",
		_("Microsoft rejected one of the event's values"),
		_(
			"Most often the end time is not after the start time. Frappe pre-fills both from the "
			"current moment and does not enforce the order, so an event saved without touching "
			"the times can end before it starts. Check the start and end on the event."
		),
	),
	(
		r"Unknown column 'custom_.*microsoft",
		_("The app's custom fields are missing from this site"),
		_(
			"They are created on install and on migrate. Run `bench --site <site> migrate`, or "
			"use Migrate in the Frappe Cloud site dashboard, then sync again."
		),
	),
	(
		r"Please Authorize OAuth",
		_("No token is stored for this account yet"),
		_("Open the Email Account and click Authorize API Access, signed in as the Connected User."),
	),
]


def explain_error(text):
	"""Translate a Microsoft/IMAP/SMTP error into a cause and a next step."""
	text = text or ""
	for pattern, title, detail in error_patterns():
		if re.search(pattern, text, re.IGNORECASE):
			return {"matched": True, "title": title, "detail": detail, "doc": ENTRA_ERROR_DOC}
	return {
		"matched": False,
		"title": _("Unrecognised error"),
		"detail": _("No known Microsoft cause matches this message."),
		"doc": ENTRA_ERROR_DOC,
	}


# --- Exchange PowerShell for app-only access ------------------------------------------

def powershell_for_app_only(client_id, enterprise_object_id=None, mailboxes=None, send_as=False, include_undo=True):
	"""The Exchange Online commands that app-only mailbox access requires.

	Microsoft's biggest documented trap: New-ServicePrincipal wants the Object ID from the
	ENTERPRISE APPLICATION blade, not the one shown on the App Registration page. Using the
	wrong one fails at authentication time with no useful message, so the generated script
	looks it up by AppId instead of asking anyone to copy it.
	"""
	mailboxes = mailboxes or []
	lines = [
		"# Run in Exchange Online PowerShell as a tenant admin.",
		"Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser",
		"Import-Module ExchangeOnlineManagement",
		"Connect-ExchangeOnline",
		"",
		"# Look the service principal up by AppId so the correct (Enterprise Application)",
		"# Object ID is used — copying the App Registration one causes silent auth failures.",
		f'$appId = "{client_id or "<CLIENT_ID>"}"',
	]

	if enterprise_object_id:
		lines.append(f'$objectId = "{enterprise_object_id}"')
	else:
		lines += [
			"$sp = Get-MgServicePrincipal -Filter \"appId eq '$appId'\"   # or Get-AzureADServicePrincipal",
			"$objectId = $sp.Id",
		]

	lines += [
		"",
		"New-ServicePrincipal -AppId $appId -ObjectId $objectId -DisplayName \"Frappe mail access\"",
		"$exoSp = Get-ServicePrincipal -Identity \"Frappe mail access\"",
		"",
		"# Grant only the mailboxes this app should ever read. Access is scoped per mailbox,",
		"# so the app cannot reach anything not listed here.",
	]

	for mailbox in mailboxes or ["<shared@yourdomain.com>"]:
		lines.append(
			f'Add-MailboxPermission -Identity "{mailbox}" -User $exoSp.Identity -AccessRights FullAccess'
		)
		if send_as:
			lines.append(
				f'Add-RecipientPermission -Identity "{mailbox}" -Trustee $exoSp.Identity '
				f"-AccessRights SendAs -Confirm:$false"
			)

	if include_undo:
		# Everything above is reversible. Ship the reverse commands with the script so nobody
		# has to work them out under pressure later.
		lines += [
			"",
			"# " + "-" * 72,
			"# TO UNDO. Removes the application's access again; run as a tenant admin.",
			"# " + "-" * 72,
		]
		for mailbox in mailboxes or ["<shared@yourdomain.com>"]:
			if send_as:
				lines.append(
					f'# Remove-RecipientPermission -Identity "{mailbox}" -Trustee $exoSp.Identity '
					f"-AccessRights SendAs -Confirm:$false"
				)
			lines.append(
				f'# Remove-MailboxPermission -Identity "{mailbox}" -User $exoSp.Identity '
				f"-AccessRights FullAccess -Confirm:$false"
			)
		lines.append("# Remove-ServicePrincipal -Identity $exoSp.Identity   # revokes it everywhere")

	return "\n".join(lines)


# --- collectors -----------------------------------------------------------------------

def _settings_config():
	settings = frappe.get_cached_doc("Microsoft Settings")
	secret = get_decrypted_password(
		"Microsoft Settings", "Microsoft Settings", "client_secret", raise_exception=False
	)
	return {
		"enabled": settings.enabled,
		"tenant_id": settings.tenant_id,
		"client_id": settings.client_id,
		"has_client_secret": bool(secret),
		"client_secret": secret,
		"redirect_uri": settings.redirect_uri,
		"use_calendar": settings.use_calendar,
		# .get() rather than attribute access: the doctor is what people run when a site is
		# half-upgraded, and it must not itself blow up on a field the site has not migrated yet.
		"use_teams": settings.get("use_teams"),
		"use_transcripts": settings.get("use_transcripts"),
		"use_mail": settings.use_mail,
		"use_sso": settings.use_sso,
		"mail_flow": settings.mail_flow or "Delegated",
		"default_scopes": settings.get("default_scopes") or "",
		"authorized_scopes": settings.get("authorized_scopes") or "",
	}


def _connected_app_config(name):
	doc = frappe.get_doc("Connected App", name)
	return {
		"name": doc.name,
		"client_id": doc.client_id,
		"redirect_uri": doc.redirect_uri,
		"authorization_uri": doc.authorization_uri,
		"token_uri": doc.token_uri,
		"scopes": [row.scope for row in (doc.scopes or [])],
	}


EMAIL_ACCOUNT_FIELDS = [
	"name",
	"email_id",
	"auth_method",
	"connected_app",
	"connected_user",
	"backend_app_flow",
	"login_id",
	"login_id_is_different",
	"enable_incoming",
	"enable_outgoing",
	"use_imap",
	"use_ssl",
	"use_starttls",
	"email_server",
	"smtp_server",
	"service",
]


def _email_account_configs():
	accounts = frappe.get_all("Email Account", fields=EMAIL_ACCOUNT_FIELDS)
	for account in accounts:
		account["imap_folder"] = frappe.db.count("IMAP Folder", {"parent": account["name"]})
	return accounts


@frappe.whitelist()
def run_diagnostics():
	"""Inspect the Microsoft 365 mail/identity configuration. Read-only. System Manager only."""
	frappe.only_for("System Manager")

	settings = _settings_config()
	findings = check_settings(settings) + check_credentials(settings)

	if any(settings.get(field) for field, _scope in graph.CAPABILITY_SCOPES):
		findings += check_scopes(settings)
		findings += check_authorized_scopes(
			graph.get_scopes(settings),
			graph.parse_scopes(settings.get("authorized_scopes")),
			# Only authorised connections hold a token; an unauthorised one picks up the new
			# scopes the first time it signs in, so it is not a problem to report.
			frappe.db.count("Microsoft Calendar", {"authorized": 1}),
		)

	if settings.get("use_calendar"):
		findings += check_event_custom_fields(
			[f for f in EVENT_CUSTOM_FIELDS if not frappe.db.has_column("Event", f)]
		)

	# Only where something actually depends on the queue. refresh=True because this is the one
	# place a person looks *after* starting a worker to find out whether it worked, and a
	# minute-old "no workers" would send them round the loop again.
	if settings.get("use_calendar") or settings.get("use_transcripts"):
		findings += check_background_jobs(background.health(refresh=True))

	wants_mail = bool(settings.get("use_mail"))
	wants_sso = bool(settings.get("use_sso"))

	accounts = _email_account_configs() if wants_mail else []
	oauth_accounts = [a for a in accounts if a.get("auth_method") == "OAuth"]

	checked_apps = set()
	for account in oauth_accounts:
		app_name = account.get("connected_app")
		if app_name and app_name not in checked_apps and frappe.db.exists("Connected App", app_name):
			checked_apps.add(app_name)
			findings += check_connected_app(
				_connected_app_config(app_name), settings, app_only=bool(account.get("backend_app_flow"))
			)

	for account in accounts:
		findings += check_email_account(account)

	if wants_sso:
		sso_name = frappe.db.exists("Social Login Key", {"social_login_provider": "Office 365"})
		if sso_name:
			key = frappe.db.get_value(
				"Social Login Key",
				sso_name,
				["name", "enable_social_login", "client_id", "authorize_url", "access_token_url"],
				as_dict=True,
			)
			findings += check_social_login_key(key, settings)
		else:
			findings.append(
				finding(
					"sso.missing",
					FAIL,
					_("Microsoft sign-in is selected but not set up"),
					"",
					_("Use Set Up on Microsoft Settings to create the Social Login Key."),
				)
			)

	if wants_mail and not oauth_accounts:
		findings.append(
			finding(
				"email_account.none",
				WARN,
				_("No Email Account is using OAuth"),
				_("{0} mail account(s) found, none on OAuth.").format(len(accounts)),
				_("Microsoft disables Basic auth for SMTP by default from the end of December 2026."),
				MS_OAUTH_DOC,
			)
		)

	if not any(f["status"] in (FAIL, WARN) for f in findings):
		findings.append(
			finding("all.ok", PASS, _("No configuration problems found"), "", "", "")
		)

	counts = {status: len([f for f in findings if f["status"] == status]) for status in (PASS, WARN, FAIL, SKIP)}
	return {"findings": findings, "counts": counts, "checked_connected_apps": sorted(checked_apps)}


@frappe.whitelist()
def run_for_email_account(email_account: str):
	"""Check one mail account, plus the Connected App it depends on. Read-only."""
	frappe.only_for("System Manager")

	settings = _settings_config()
	account = frappe.db.get_value("Email Account", email_account, EMAIL_ACCOUNT_FIELDS, as_dict=True)
	if not account:
		frappe.throw(_("Email Account {0} not found").format(email_account))
	account["imap_folder"] = frappe.db.count("IMAP Folder", {"parent": email_account})

	findings = check_email_account(account)
	app_name = account.get("connected_app")
	if app_name and frappe.db.exists("Connected App", app_name):
		findings += check_connected_app(
			_connected_app_config(app_name), settings, app_only=bool(account.get("backend_app_flow"))
		)

	if not findings:
		findings = [finding("account.ok", PASS, _("No configuration problems found for {0}").format(email_account))]

	counts = {status: len([f for f in findings if f["status"] == status]) for status in (PASS, WARN, FAIL, SKIP)}
	return {"findings": findings, "counts": counts}


@frappe.whitelist()
def explain(error_text: str | None = None):
	"""Whitelisted wrapper so the error decoder can be used from the Desk."""
	frappe.only_for("System Manager")
	return explain_error(error_text)


@frappe.whitelist()
def app_only_powershell(mailboxes: str | list | None = None, send_as: int = 0):
	"""Generate the Exchange Online setup script for app-only mailbox access."""
	frappe.only_for("System Manager")
	if isinstance(mailboxes, str):
		mailboxes = [m.strip() for m in mailboxes.replace(",", "\n").split("\n") if m.strip()]
	settings = _settings_config()
	return {
		"script": powershell_for_app_only(
			settings.get("client_id"), mailboxes=mailboxes, send_as=frappe.utils.cint(send_as)
		),
		"warning": _(
			"This grants the application permanent access to the listed mailboxes, with no "
			"one signed in. Nothing runs from here: you paste it into Exchange Online "
			"yourself, and the script ends with the commands that reverse it."
		),
		"mailboxes": mailboxes or [],
	}
