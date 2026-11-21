"""Sign in with Microsoft: only accounts from the organisation's own directory.

Frappe's Office 365 login takes the email address from Microsoft's ID token and signs in the
Frappe user who has it. On a multi-tenant authority (``common``, ``organizations``) Microsoft
issues that token to an account in *any* directory, and a directory's owner chooses its users'
addresses — so the address alone proves nothing. What does is the token's ``tid``, the directory
that vouches for the account. This module makes sure it is the organisation's own.

The layers:

* Provisioning refuses to create the sign-in record unless Tenant ID names one directory, and
  the doctor reports a sign-in record on a multi-tenant authority as a failure.
* ``login_via_office365`` replaces Frappe's Microsoft callback (``override_whitelisted_methods``)
  on every version. It exchanges the code exactly as Frappe does, applies the directory rule to
  the token's claims before Frappe touches the user record, and then hands over to Frappe's own
  ``login_oauth_user``. Frappe 15 has no ``on_social_login`` hook, so on Frappe 15 this is the
  only place the rule can run; Frappe 15 also does not check the OAuth ``state`` against anything
  stored, so there the wrapper does that too, the way Frappe 16 does.
* On Frappe 16, ``on_social_login`` (called with the token's claims before the session is
  created) applies the same rule again for any other route into ``login_oauth_user``.
* An upgrade patch filled Tenant ID in for existing sites in 1.1.1, from the directory their
  authorised calendars already belong to. From 1.2.0 a multi-tenant Tenant ID refuses sign-in.
"""

import base64
import json
import re

import frappe
from frappe import _

#: Social Login Key name of Frappe's Office 365 provider: ``frappe.scrub("Office 365")``. It is
#: also the provider ``login_via_office365`` passes to every hook, whatever the record says.
PROVIDER = "office_365"

#: Authority segments that accept accounts from more than one directory. Empty means Frappe's
#: default, which is ``common``.
MULTI_TENANT = ("", "common", "organizations", "consumers")

#: The directory Microsoft puts personal (consumer) accounts in. A token from it says nothing
#: about which organisation a site belongs to.
CONSUMER_TENANT = "9188040d-6c67-4c5b-b112-36a304b66dad"

GUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)
DOMAIN = re.compile(r"^(?=.{4,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$", re.IGNORECASE)

GUID_TENANT, DOMAIN_TENANT, MULTI_TENANT_KIND, INVALID_TENANT = "guid", "domain", "multi", "invalid"


def tenant_kind(tenant):
	"""Classify a Tenant ID: one directory by id, one directory by domain, any directory, or junk."""
	tenant = (tenant or "").strip()
	if tenant.lower() in MULTI_TENANT:
		return MULTI_TENANT_KIND
	if GUID.match(tenant):
		return GUID_TENANT
	if DOMAIN.match(tenant):
		return DOMAIN_TENANT
	return INVALID_TENANT


def is_single_tenant(tenant):
	"""True when ``tenant`` names exactly one directory, as sign-in requires."""
	return tenant_kind(tenant) in (GUID_TENANT, DOMAIN_TENANT)


def authority_tenant(url):
	"""The tenant segment of a login.microsoftonline.com URL; None for any other URL."""
	match = re.search(r"login\.microsoftonline\.com/([^/?#]*)", url or "", re.IGNORECASE)
	return match.group(1) if match else None


def refusal(claims, tenant, token_url):
	"""Why a Microsoft sign-in must be refused, or None to let it through. Pure.

	``claims``: the ID token's claims. ``tenant``: Tenant ID in Microsoft Settings.
	``token_url``: the token endpoint Frappe exchanged the sign-in code at.

	* A GUID tenant is compared with ``tid`` directly.
	* A domain cannot be: ``tid`` is always the directory's GUID, and turning a domain into one
	  needs a call to Microsoft in the middle of a login. It does not need to, though. A
	  tenant-specific authority only ever issues tokens for its own directory, and Frappe
	  exchanged the code at the token endpoint itself, over TLS, so the token's directory is the
	  endpoint's. A domain tenant is therefore accepted when the token endpoint is that domain's
	  (or a GUID equal to ``tid``), and refused when it is a multi-tenant one, which would let
	  any directory through.
	* A multi-tenant Tenant ID cannot tell this organisation's directory from any other, so it
	  is refused. 1.1.1 still let it through so that nobody was locked out by that upgrade (the
	  doctor and the upgrade report said so); from 1.2.0 the sign-in is refused, with a message
	  that tells the person what their administrator has to change.
	"""
	tid = (claims.get("tid") or "").strip().lower()
	tenant = (tenant or "").strip()
	kind = tenant_kind(tenant)

	if kind == GUID_TENANT:
		if tid != tenant.lower():
			return _("This Microsoft account belongs to a different organisation's directory.")
		return None

	if kind == DOMAIN_TENANT:
		endpoint = (authority_tenant(token_url) or "").strip().lower()
		if endpoint == tenant.lower():
			return None
		if GUID.match(endpoint) and endpoint == tid:
			return None
		return _("This Microsoft account could not be confirmed as one of your organisation's.")

	if kind == INVALID_TENANT:
		return _("Microsoft sign-in is not set up correctly: Tenant ID is not a directory ID or domain.")

	return _(
		"Microsoft sign-in is turned off on this site until it is limited to your organisation's "
		"own directory. Please ask your administrator to set Tenant ID in Microsoft Settings to "
		"your Directory (tenant) ID."
	)


def _token_endpoint():
	"""The token URL Frappe used for this sign-in, as Frappe itself builds it."""
	key = frappe.db.get_value(
		"Social Login Key", PROVIDER, ["base_url", "custom_base_url", "access_token_url"], as_dict=True
	)
	if not key:
		return ""
	if key.custom_base_url:
		from frappe.utils.oauth import build_oauth_url

		return build_oauth_url(key.base_url, key.access_token_url)
	return key.access_token_url or ""


def check_claims(user, claims):
	"""Why the Microsoft sign-in of ``user`` with these ID token claims must be refused, or None.

	The rule both entry points apply. It only reads, so running it twice in one login (the
	callback wrapper, then Frappe 16's hook) gives the same answer twice and changes nothing.
	"""
	claims = claims or {}
	tenant = frappe.db.get_single_value("Microsoft Settings", "tenant_id")
	problem = refusal(claims, tenant, _token_endpoint())
	if problem:
		frappe.logger("frappe_microsoft365").warning(
			f"Microsoft sign-in refused for {user}: token directory {claims.get('tid')!r}, "
			f"Tenant ID {tenant!r}"
		)
	return problem


def on_social_login(user, provider, claims, **kwargs):
	"""Frappe 16 hook, run after Microsoft vouches for an account and before the session starts.

	Raising here stops the login; the request is rolled back, so nothing the login wrote to the
	user record survives. A sign-in through this app's callback has already been checked before
	Frappe saw it; this catches any other route into ``login_oauth_user``.
	"""
	if provider != PROVIDER:
		return

	problem = check_claims(user, claims)
	if problem:
		frappe.throw(problem, frappe.AuthenticationError, title=_("Sign-in refused"))


# --- the Microsoft callback ----------------------------------------------------------------

#: Frappe's Microsoft sign-in callback, which ``hooks.override_whitelisted_methods`` points here.
#: The Social Login Key's redirect URL names it, on Frappe 15 and 16 alike.
FRAPPE_CALLBACK = "frappe.integrations.oauth2_logins.login_via_office365"
CALLBACK = "frappe_microsoft365.sign_in.login_via_office365"

#: Where a Frappe 15 sign-in's ``state`` token is remembered between the login page and the
#: callback. Ten minutes, as Frappe 16 allows.
STATE_PREFIX = "frappe_microsoft365:login_state"
STATE_TTL = 600


def _oauth():
	"""Frappe's OAuth module. A function so tests can stand Frappe 15's in for it."""
	from frappe.utils import oauth

	return oauth


def frappe_checks_login_state():
	"""Whether Frappe itself accepts a sign-in ``state`` only once, from a server-side store.

	Frappe 16 does (``consume_oauth_state``). Frappe 15 only checks that the state it gets back
	has a token in it, which anyone can write.
	"""
	return hasattr(_oauth(), "consume_oauth_state")


def frappe_runs_sign_in_hook():
	"""Whether Frappe's ``login_oauth_user`` calls the ``on_social_login`` hooks.

	Read from the function's own constants rather than the version number, because the hook is
	what matters: a Frappe that has it calls ``on_social_login`` by that name.
	"""
	code = getattr(_oauth().login_oauth_user, "__code__", None)
	return bool(code) and "on_social_login" in code.co_consts


def callback_is_ours():
	"""Whether Frappe routes the Microsoft callback to this module, not to another app's override."""
	return frappe.override_whitelisted_method(FRAPPE_CALLBACK) == CALLBACK


def _state_token(state):
	"""The ``token`` inside a Frappe 15 sign-in state (base64 JSON), or None if it has none."""
	try:
		raw = state.encode() if isinstance(state, str) else bytes(state)
		data = json.loads(base64.b64decode(raw).decode("utf-8"))
	except (ValueError, TypeError, AttributeError):
		return None
	token = data.get("token") if isinstance(data, dict) else None
	return token if isinstance(token, str) and token else None


def _state_key(token):
	return frappe.cache.make_key(f"{STATE_PREFIX}:{token}")


def remember_login_state(context):
	"""``update_website_context`` hook: on Frappe 15, store the state the login page hands out.

	Frappe 15's login page builds the Microsoft button's URL, with a fresh ``state``, in
	``get_context``; this runs right after it, on the same request. Remembering the state's token
	here is what lets the callback accept it only once, as Frappe 16 does. On Frappe 16, and on
	every page but the login page, it does nothing.
	"""
	logins = context.get("provider_logins") if hasattr(context, "get") else None
	if not logins or frappe_checks_login_state():
		return
	from urllib.parse import parse_qs, urlsplit

	for login in logins:
		if not isinstance(login, dict) or login.get("name") != PROVIDER:
			continue
		state = (parse_qs(urlsplit(login.get("auth_url") or "").query).get("state") or [None])[0]
		token = _state_token(state)
		if token:
			# _state_key is site-scoped (make_key); raw set and delete give the atomic consume below.
			frappe.cache.set(_state_key(token), "1", ex=STATE_TTL)  # nosemgrep


def consume_login_state(state):
	"""True once for a state this site's login page issued in the last ten minutes, then never.

	The delete is the check: Redis removes the key and says whether it was there in one step,
	so two callbacks racing with the same state cannot both get through.
	"""
	token = _state_token(state)
	if not token:
		return False
	# Site-scoped key (see _state_key).
	return bool(frappe.cache.delete(_state_key(token)))  # nosemgrep


# Guest by necessity: it is the sign-in callback, called before anyone is signed in, exactly like
# the Frappe method it overrides.
@frappe.whitelist(allow_guest=True)  # nosemgrep
def login_via_office365(code: str, state: str):
	"""Frappe's Microsoft sign-in callback, with this app's checks in front of Frappe's login.

	The same steps as Frappe's ``login_via_office365``: exchange the code for an ID token with
	Frappe's ``get_info_via_oauth`` (so the claims are decoded exactly as Frappe decodes them),
	then sign in with Frappe's ``login_oauth_user``. In between, the directory rule runs, and
	a refusal is shown as a page saying why instead of an error. On Frappe 15 the ``state`` is
	also checked against the one the login page issued, before Microsoft is asked anything.
	"""
	from frappe.integrations.oauth2_logins import decoder_compat

	oauth = _oauth()
	if not frappe_checks_login_state() and not consume_login_state(state):
		frappe.respond_as_web_page(
			_("Invalid Request"),
			_("Your login attempt is invalid or has expired. Please try again."),
			http_status_code=417,
		)
		return

	info = oauth.get_info_via_oauth(PROVIDER, code, decoder_compat, id_token=True)

	problem = check_claims(oauth.get_email(info), info)
	if problem:
		frappe.respond_as_web_page(
			_("Sign-in refused"), problem, success=False, http_status_code=403, indicator_color="red"
		)
		return

	oauth.login_oauth_user(info, provider=PROVIDER, state=state)


# --- reading a token's directory, for the upgrade patch ----------------------------------


def tid_of(token):
	"""The ``tid`` claim of a JWT, read without verifying it; None when there is none.

	Only for working out which directory a token we stored ourselves came from — never to
	decide whether to trust it. Personal-account tokens are opaque rather than JWTs, and give
	None.
	"""
	parts = (token or "").split(".")
	if len(parts) != 3:
		return None
	payload = parts[1] + "=" * (-len(parts[1]) % 4)
	try:
		claims = json.loads(base64.urlsafe_b64decode(payload.encode()))
	except (ValueError, TypeError):
		return None
	tid = claims.get("tid") if isinstance(claims, dict) else None
	return tid.strip().lower() if isinstance(tid, str) and GUID.match(tid.strip()) else None


def calendar_tenant():
	"""The one directory every authorised Microsoft Calendar's token belongs to, if there is one.

	None when there are no authorised calendars, when any token cannot be read, when they
	disagree, or when they are personal accounts: in each case guessing could cut someone off.
	"""
	if not frappe.db.table_exists("Microsoft Calendar"):
		return None
	from frappe.utils.password import get_decrypted_password

	tids = set()
	names = frappe.get_all("Microsoft Calendar", filters={"authorized": 1}, pluck="name")
	for name in names:
		try:
			token = get_decrypted_password("Microsoft Calendar", name, "access_token", raise_exception=False)
		except Exception:
			# A rotated encryption key leaves a token nobody can read; it proves nothing.
			return None
		tid = tid_of(token)
		if not tid:
			return None
		tids.add(tid)
	if len(tids) != 1:
		return None
	tid = tids.pop()
	return None if tid == CONSUMER_TENANT else tid


# --- upgrade report -------------------------------------------------------------------------


def upgrade_check():
	"""Read-only: is an enabled Microsoft sign-in still open to every directory?"""
	from frappe_microsoft365.doctor import check_sign_in_tenant

	if not frappe.db.table_exists("Social Login Key"):
		return []
	key = frappe.db.get_value(
		"Social Login Key",
		PROVIDER,
		["name", "enable_social_login", "authorize_url", "access_token_url"],
		as_dict=True,
	)
	if not key or not key.enable_social_login:
		return []
	tenant = frappe.db.get_single_value("Microsoft Settings", "tenant_id")
	findings = check_sign_in_tenant(key, {"tenant_id": tenant})
	for f in findings:
		f["check"] = "upgrade.sign_in_tenant"
	return findings


def upgrade_check_multi_tenant_refused():
	"""Read-only, 1.2.0: Microsoft sign-in that this release now refuses for everyone.

	The same condition as the 1.1.1 check, but a different consequence. 1.1.1 kept such a site
	signing people in and said so; 1.2.0 refuses every Microsoft sign-in until Tenant ID names
	the organisation's directory. A site without an enabled Microsoft sign-in hears nothing.
	"""
	from frappe_microsoft365 import doctor

	if not frappe.db.table_exists("Social Login Key"):
		return []
	key = frappe.db.get_value("Social Login Key", PROVIDER, ["name", "enable_social_login"], as_dict=True)
	if not key or not key.enable_social_login:
		return []
	tenant = (frappe.db.get_single_value("Microsoft Settings", "tenant_id") or "").strip()
	if is_single_tenant(tenant):
		return []
	return [
		doctor.finding(
			"upgrade.sign_in_refused",
			doctor.FAIL,
			_("Sign in with Microsoft is now refused for everyone on this site"),
			_(
				"Tenant ID is '{0}', which accepts accounts from any organisation's directory. From "
				"1.2.0 every Microsoft sign-in is refused until it names your directory; people see "
				"a page asking them to contact their administrator. Password sign-in is not affected."
			).format(tenant or "common"),
			_(
				"Set Tenant ID in Microsoft Settings to the Directory (tenant) ID on the app "
				"registration's Overview page in Microsoft Entra, then point the Social Login Key's "
				"Authorize URL and Access Token URL at "
				"https://login.microsoftonline.com/<tenant id>/oauth2/v2.0/... ."
			),
			doctor.MS_OAUTH_DOC,
			target=key.name,
		)
	]
