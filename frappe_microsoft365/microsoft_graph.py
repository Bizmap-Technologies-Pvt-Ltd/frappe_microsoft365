"""Microsoft Graph + MSAL core for frappe_microsoft365.

Generic, reusable: any Frappe/ERPNext site configures Microsoft Settings once, each user
authorizes a Microsoft Calendar, and this module handles OAuth (MSAL), token refresh, and
authenticated Graph v1.0 calls. Secrets/tokens are never logged or returned to clients.

See docs/graph-api-reference.md for the verified endpoint/permission contract.
"""

import time

import frappe
import requests
from frappe import _
from frappe.utils import add_to_date, get_datetime, get_url, now_datetime
from frappe.utils.password import get_decrypted_password

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

#: Requested on every sign-in regardless of capabilities: /me is how the connection resolves
#: which Microsoft account it is attached to.
BASE_SCOPE = "User.Read"

#: Capability field on Microsoft Settings -> the delegated scopes it needs. Ordered so the
#: derived list reads the same way every time.
#:
#: The split matters because the permissions are not interchangeable, verified against the
#: endpoints this app calls (docs/graph-api-reference.md):
#:
#: * Creating an event with isOnlineMeeting/teamsForBusiness on POST /me/events is a calendar
#:   write — Microsoft mints the Teams link as part of the event — so the Teams tickbox on an
#:   Event needs Calendars.ReadWrite and nothing else.
#: * OnlineMeetings.ReadWrite is needed only for standalone meetings (POST /me/onlineMeetings)
#:   and for resolving a join URL to a meeting id.
#: * OnlineMeetingTranscript.Read.All is needed only for transcripts, and recordings need
#:   their own OnlineMeetingRecording.Read.All on top of it — Microsoft grants the two
#:   separately, and a tenant that consents to reading words often will not consent to
#:   reading video.
#:
#: Bundling them forced admins to hand-edit the scope field back down to what their tenant had
#: actually consented to, which is the failure this split removes.
CAPABILITY_SCOPES = (
	("use_calendar", ("Calendars.ReadWrite",)),
	("use_teams", ("OnlineMeetings.ReadWrite",)),
	("use_transcripts", ("OnlineMeetingTranscript.Read.All", "OnlineMeetingRecording.Read.All")),
)

#: Added by build_authorize_url, never derived and never valid in an override: MSAL treats
#: these as reserved and refuses them in the `scopes` argument of the token calls, so an
#: admin who lists them here would break sign-in rather than widen it.
RESERVED_SCOPES = ("offline_access", "openid", "profile")

CALLBACK_METHOD = (
	"frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback"
)

#: Graph asks us to back off with a Retry-After header. Wait inline only for short waits;
#: anything longer is left to the next scheduled run.
MAX_RETRY_AFTER_SECONDS = 10
DEFAULT_RETRY_AFTER_SECONDS = 2

#: Safety valve for link-following loops (nextLink / deltaLink paging).
MAX_PAGES = 50


class MsGraphError(frappe.ValidationError):
	pass


class MsGraphResyncRequired(MsGraphError):
	"""Graph invalidated our delta token (410 Gone). The caller must restart from a full sync."""


# --- settings helpers ----------------------------------------------------------------

def _settings_doc():
	"""The Microsoft Settings single, WITHOUT the enabled check (pure config reads)."""
	return frappe.get_cached_doc("Microsoft Settings")


def get_settings():
	"""Settings for LIVE operations — requires the integration to be enabled."""
	s = _settings_doc()
	if not s.enabled:
		frappe.throw(_("Microsoft 365 integration is disabled. Enable it in Microsoft Settings."), MsGraphError)
	return s


def _client_secret():
	val = get_decrypted_password("Microsoft Settings", "Microsoft Settings", "client_secret", raise_exception=False)
	if not val:
		frappe.throw(_("Microsoft Settings: client secret is not set."), MsGraphError)
	return val


def get_authority(settings=None):
	settings = settings or _settings_doc()
	tenant = (settings.tenant_id or "common").strip()
	return f"https://login.microsoftonline.com/{tenant}"


def parse_scopes(raw):
	"""Split a space- or comma-separated scope string. Order kept, duplicates dropped."""
	scopes = []
	for scope in (raw or "").replace(",", " ").split():
		scope = scope.strip()
		if scope and scope not in scopes:
			scopes.append(scope)
	return scopes


def derive_scopes(capabilities):
	"""The delegated scopes implied by the ticked capabilities.

	Pure: takes a plain dict (or any .get() mapping, including the Settings Document) so the
	rule is testable without a site or a tenant.

	Nothing reserved is returned. offline_access/openid/profile are added once, centrally, in
	build_authorize_url; including them here would put them in the MSAL token calls too, which
	reject them.

	Asking for a permission nothing uses is not free — it shows up on the consent screen, and
	plenty of tenants refuse OnlineMeetings or transcript consent outright — so a capability
	that is not ticked contributes no scope at all.
	"""
	capabilities = capabilities or {}
	scopes = [BASE_SCOPE]
	for field, field_scopes in CAPABILITY_SCOPES:
		if not capabilities.get(field):
			continue
		for scope in field_scopes:
			if scope not in scopes:
				scopes.append(scope)
	return scopes


def get_scopes(settings=None):
	"""The delegated scopes sign-in will ask Microsoft for.

	An override wins verbatim: a tenant that consents to a hand-picked list has to be asked for
	exactly that list, and second-guessing it would put the admin back to editing free text.
	With no override the scopes follow the capability tickboxes, so the two can never disagree.
	"""
	settings = settings or _settings_doc()
	override = parse_scopes(settings.get("default_scopes"))
	if override:
		return override
	return derive_scopes(settings)


@frappe.whitelist()
def preview_scopes(capabilities: dict | str | None = None, override: str | None = None):
	"""What sign-in would request for the capabilities currently on screen. Reads nothing else.

	The Settings form shows this rather than asking an admin to work it out, and it is computed
	by the same functions get_scopes uses so the display cannot drift from what is requested.
	"""
	frappe.only_for("System Manager")
	if isinstance(capabilities, str):
		capabilities = frappe.parse_json(capabilities) or {}
	derived = derive_scopes(capabilities)
	overridden = parse_scopes(override)
	return {
		"scopes": overridden or derived,
		"derived": derived,
		"overridden": bool(overridden),
		"always_added": list(RESERVED_SCOPES),
		"missing_from_override": [s for s in derived if s not in overridden] if overridden else [],
	}


def get_redirect_uri(settings=None):
	settings = settings or _settings_doc()
	if settings.redirect_uri:
		return settings.redirect_uri.strip()
	return f"{get_url()}/api/method/{CALLBACK_METHOD}"


def _msal_app(settings=None):
	import msal

	settings = settings or get_settings()
	return msal.ConfidentialClientApplication(
		settings.client_id,
		client_credential=_client_secret(),
		authority=get_authority(settings),
	)


# --- OAuth lifecycle -----------------------------------------------------------------

def build_authorize_url(state):
	"""Build the Microsoft sign-in URL for the auth-code flow."""
	import urllib.parse

	settings = get_settings()
	params = {
		"client_id": settings.client_id,
		"response_type": "code",
		"redirect_uri": get_redirect_uri(settings),
		"response_mode": "query",
		"scope": " ".join([*RESERVED_SCOPES, *get_scopes(settings)]),
		"state": state,
		"prompt": "select_account",
	}
	return f"{get_authority(settings)}/oauth2/v2.0/authorize?" + urllib.parse.urlencode(params)


def exchange_code(code):
	"""Exchange an authorization code for tokens (confidential client, no PKCE)."""
	settings = get_settings()
	scopes = get_scopes(settings)
	result = _msal_app(settings).acquire_token_by_authorization_code(
		code, scopes=scopes, redirect_uri=get_redirect_uri(settings)
	)
	_raise_on_token_error(result)
	_record_authorized_scopes(scopes)
	return result


def _record_authorized_scopes(scopes):
	"""Remember the scope list that was in force at the last successful authorisation.

	A token carries the permissions consented when it was issued; ticking another capability
	afterwards does not widen a token that already exists, and Graph answers the newly enabled
	call with a bare 403. Recording this is what lets the doctor say "re-authorise" instead of
	leaving that to be discovered in production.

	What we requested is stored, not the `scope` Microsoft echoes back: that comes back as
	fully-qualified resource URIs plus the reserved scopes, so it is not comparable with the
	short names the settings hold.

	Written straight to the Singles row so the OAuth callback — which runs as whoever signed
	in, not necessarily a System Manager — is never blocked by permissions on Settings.
	"""
	frappe.db.set_single_value("Microsoft Settings", "authorized_scopes", " ".join(scopes))


def refresh_tokens(refresh_token):
	settings = get_settings()
	result = _msal_app(settings).acquire_token_by_refresh_token(refresh_token, scopes=get_scopes(settings))
	_raise_on_token_error(result)
	return result


def _raise_on_token_error(result):
	if not result or "access_token" not in result:
		err = (result or {}).get("error_description") or (result or {}).get("error") or "unknown error"
		# never include tokens; error_description from MS is safe (no secrets)
		frappe.throw(f"Microsoft sign-in failed: {err}", MsGraphError)


# --- per-calendar token management ---------------------------------------------------

def _store_tokens(calendar_name, result):
	expiry = add_to_date(now_datetime(), seconds=(result.get("expires_in") or 3600) - 300)
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	doc.access_token = result.get("access_token")
	if result.get("refresh_token"):
		doc.refresh_token = result["refresh_token"]
	doc.token_expiry = expiry
	claims = result.get("id_token_claims") or {}
	if claims.get("preferred_username") and not doc.microsoft_user_email:
		doc.microsoft_user_email = claims.get("preferred_username")
	doc.flags.ignore_permissions = True
	doc.save()
	return doc


def get_valid_access_token(calendar):
	"""Return a non-expired access token for a Microsoft Calendar doc/name, refreshing if needed."""
	name = calendar if isinstance(calendar, str) else calendar.name
	doc = frappe.get_doc("Microsoft Calendar", name)
	token = get_decrypted_password("Microsoft Calendar", name, "access_token", raise_exception=False)
	expiry = get_datetime(doc.token_expiry) if doc.token_expiry else None
	if token and expiry and expiry > now_datetime():
		return token
	# refresh
	refresh = get_decrypted_password("Microsoft Calendar", name, "refresh_token", raise_exception=False)
	if not refresh:
		frappe.throw(f"Microsoft Calendar '{name}' is not authorized. Click Authorize.", MsGraphError)
	result = refresh_tokens(refresh)
	_store_tokens(name, result)
	return result.get("access_token")


# --- authenticated Graph requests ----------------------------------------------------

def graph_request(method, path, calendar, json=None, params=None, headers=None, raw=False, _retried=False):
	"""Authenticated Graph v1.0 call. Refreshes the token once on 401 and retries.

	`path` is relative to GRAPH_BASE (e.g. '/me/events') or an absolute graph URL.
	Returns parsed JSON (or the requests.Response when raw=True).
	"""
	name = calendar if isinstance(calendar, str) else calendar.name
	token = get_valid_access_token(name)
	url = path if path.startswith("http") else f"{GRAPH_BASE}{path}"
	req_headers = {"Authorization": f"Bearer {token}"}
	if headers:
		req_headers.update(headers)
	resp = requests.request(method, url, json=json, params=params, headers=req_headers, timeout=30)

	if resp.status_code == 401 and not _retried:
		# force refresh then retry once
		frappe.db.set_value("Microsoft Calendar", name, "token_expiry", add_to_date(now_datetime(), seconds=-60))
		return graph_request(method, path, name, json=json, params=params, headers=headers, raw=raw, _retried=True)

	if resp.status_code == 429 and not _retried:
		# Honour Retry-After for short waits; longer backoffs are left to the next run.
		wait = _retry_after_seconds(resp)
		if wait is not None:
			time.sleep(wait)
			return graph_request(
				method, path, name, json=json, params=params, headers=headers, raw=raw, _retried=True
			)

	if resp.status_code == 429:
		frappe.throw(
			_("Microsoft Graph rate limit hit (429). The next scheduled sync will retry."), MsGraphError
		)

	if resp.status_code == 410:
		# Delta token expired/invalid — the caller has to restart with a full sync.
		frappe.throw(
			f"Microsoft Graph sync state expired ({_safe_error(resp)}). A full re-sync is required.",
			MsGraphResyncRequired,
		)

	if resp.status_code >= 400:
		detail = _safe_error(resp)
		frappe.throw(f"Microsoft Graph {method} {path} failed ({resp.status_code}): {detail}", MsGraphError)

	if raw:
		return resp
	if resp.status_code == 204 or not resp.content:
		return {}
	return resp.json()


def _safe_error(resp):
	"""A short, safe description of a Graph failure (never echoes request headers/body).

	Both halves matter: the code is what you search for, and the message is the only part
	that says WHICH property Graph objected to. "ErrorPropertyValidationFailure" on its own
	sends people hunting through a payload by hand.
	"""
	try:
		err = (resp.json() or {}).get("error", {})
		code = (err.get("code") or "").strip()
		message = (err.get("message") or "").strip()
		if code and message and message != code:
			return f"{code}: {message}"
		return code or message or resp.reason
	except Exception:
		return resp.reason


def _retry_after_seconds(resp):
	"""Seconds to wait per the Retry-After header, or None when the wait is too long to hold."""
	raw = (resp.headers or {}).get("Retry-After")
	try:
		wait = int(float(raw)) if raw else DEFAULT_RETRY_AFTER_SECONDS
	except (TypeError, ValueError):
		wait = DEFAULT_RETRY_AFTER_SECONDS
	if wait > MAX_RETRY_AFTER_SECONDS:
		return None
	return max(wait, 1)


def graph_paged(path, calendar, headers=None, max_pages=MAX_PAGES):
	"""GET every page of a Graph collection, following ``@odata.nextLink``.

	Graph caps page size (``$top`` is a hint, not a guarantee), so any collection read that
	is not explicitly "first N" MUST go through this instead of graph_request.
	"""
	items = []
	next_path = path
	for _page in range(max_pages):
		resp = graph_request("GET", next_path, calendar, headers=headers)
		items.extend(resp.get("value") or [])
		next_path = resp.get("@odata.nextLink")
		if not next_path:
			break
	return items


def graph_delta(path, calendar, headers=None, max_pages=MAX_PAGES):
	"""Follow a delta query to the end. Returns ``(items, delta_link)``.

	``delta_link`` is the ``@odata.deltaLink`` to pass back on the next run; it is only
	returned once every page has been consumed, so a partial read never advances the
	watermark. Raises MsGraphResyncRequired (410) when the token is no longer valid.
	"""
	items = []
	next_path = path
	delta_link = None
	for _page in range(max_pages):
		resp = graph_request("GET", next_path, calendar, headers=headers)
		items.extend(resp.get("value") or [])
		delta_link = resp.get("@odata.deltaLink")
		next_path = resp.get("@odata.nextLink")
		if delta_link or not next_path:
			break
	return items, delta_link


def whoami(calendar):
	return graph_request("GET", "/me", calendar)
