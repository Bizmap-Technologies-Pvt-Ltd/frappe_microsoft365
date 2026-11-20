"""Sign in with Microsoft: only accounts from the organisation's own directory.

Frappe's Office 365 login takes the email address from Microsoft's ID token and signs in the
Frappe user who has it. On a multi-tenant authority (``common``, ``organizations``) Microsoft
issues that token to an account in *any* directory, and a directory's owner chooses its users'
addresses — so the address alone proves nothing. What does is the token's ``tid``, the directory
that vouches for the account. This module makes sure it is the organisation's own.

Three layers, because Frappe 15 runs only the first two:

* Provisioning refuses to create the sign-in record unless Tenant ID names one directory, and
  the doctor reports a sign-in record on a multi-tenant authority as a failure.
* On Frappe 16, ``on_social_login`` (called with the token's claims before the session is
  created) refuses a token from any other directory.
* An upgrade patch fills Tenant ID in for existing sites, from the directory their authorised
  calendars already belong to.
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
DOMAIN = re.compile(
	r"^(?=.{4,253}$)(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$", re.IGNORECASE
)

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
	* A multi-tenant Tenant ID cannot tell this organisation's directory from any other. In
	  1.1.1 such a site keeps signing people in, so nobody is locked out by an upgrade: the
	  doctor reports it and the upgrade patch fixes what it can. It is refused from 1.2.0.
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

	return None


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


def on_social_login(user, provider, claims, **kwargs):
	"""Frappe 16 hook, run after Microsoft vouches for an account and before the session starts.

	Raising here stops the login; the request is rolled back, so nothing the login wrote to the
	user record survives. Frappe 15 has no such hook, and relies on provisioning and the doctor.
	"""
	if provider != PROVIDER:
		return

	tenant = frappe.db.get_single_value("Microsoft Settings", "tenant_id")
	problem = refusal(claims or {}, tenant, _token_endpoint())
	if problem:
		frappe.logger("frappe_microsoft365").warning(
			f"Microsoft sign-in refused for {user}: token directory {claims.get('tid')!r}, "
			f"Tenant ID {tenant!r}"
		)
		frappe.throw(problem, frappe.AuthenticationError, title=_("Sign-in refused"))


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
