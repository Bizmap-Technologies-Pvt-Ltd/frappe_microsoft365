"""Microsoft Graph + MSAL core for frappe_microsoft365.

Generic, reusable: any Frappe/ERPNext site configures Microsoft Settings once, each user
authorizes a Microsoft Calendar, and this module handles OAuth (MSAL), token refresh, and
authenticated Graph v1.0 calls. Secrets/tokens are never logged or returned to clients.

See docs/graph-api-reference.md for the verified endpoint/permission contract.
"""

import datetime

import frappe
import requests
from frappe.utils import get_url, now_datetime, add_to_date, get_datetime
from frappe.utils.password import get_decrypted_password

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = [
	"User.Read",
	"Calendars.ReadWrite",
	"OnlineMeetings.ReadWrite",
	"OnlineMeetingTranscript.Read.All",
]
CALLBACK_METHOD = (
	"frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback"
)


class MsGraphError(frappe.ValidationError):
	pass


# --- settings helpers ----------------------------------------------------------------

def get_settings():
	s = frappe.get_cached_doc("Microsoft Settings")
	if not s.enabled:
		frappe.throw("Microsoft 365 integration is disabled. Enable it in Microsoft Settings.", MsGraphError)
	return s


def _client_secret():
	val = get_decrypted_password("Microsoft Settings", "Microsoft Settings", "client_secret", raise_exception=False)
	if not val:
		frappe.throw("Microsoft Settings: client secret is not set.", MsGraphError)
	return val


def get_authority(settings=None):
	settings = settings or get_settings()
	tenant = (settings.tenant_id or "common").strip()
	return f"https://login.microsoftonline.com/{tenant}"


def get_scopes(settings=None):
	settings = settings or get_settings()
	raw = (settings.default_scopes or "").strip()
	if raw:
		return [s.strip() for s in raw.replace(",", " ").split() if s.strip()]
	return list(DEFAULT_SCOPES)


def get_redirect_uri(settings=None):
	settings = settings or get_settings()
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
		"scope": " ".join(["offline_access", "openid", "profile"] + get_scopes(settings)),
		"state": state,
		"prompt": "select_account",
	}
	return f"{get_authority(settings)}/oauth2/v2.0/authorize?" + urllib.parse.urlencode(params)


def exchange_code(code):
	"""Exchange an authorization code for tokens (confidential client, no PKCE)."""
	settings = get_settings()
	result = _msal_app(settings).acquire_token_by_authorization_code(
		code, scopes=get_scopes(settings), redirect_uri=get_redirect_uri(settings)
	)
	_raise_on_token_error(result)
	return result


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
		# brief, single retry honouring Retry-After is left to the scheduler; surface clearly here
		frappe.throw("Microsoft Graph rate limit hit (429). Try again shortly.", MsGraphError)

	if resp.status_code >= 400:
		detail = _safe_error(resp)
		frappe.throw(f"Microsoft Graph {method} {path} failed ({resp.status_code}): {detail}", MsGraphError)

	if raw:
		return resp
	if resp.status_code == 204 or not resp.content:
		return {}
	return resp.json()


def _safe_error(resp):
	try:
		body = resp.json()
		err = body.get("error", {})
		return err.get("code") or err.get("message") or resp.reason
	except Exception:
		return resp.reason


def whoami(calendar):
	return graph_request("GET", "/me", calendar)
