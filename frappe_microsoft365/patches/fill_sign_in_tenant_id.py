"""Name the organisation's directory on sites whose Microsoft sign-in accepts any directory.

With Tenant ID left at ``common`` or ``organizations``, Sign in with Microsoft accepts an account
from any organisation's directory, matched to a Frappe user by email address alone. The fix is
to name the site's own directory, and most such sites already say which one it is: every
authorised Microsoft Calendar holds an access token whose ``tid`` claim is the directory the
person signed in from.

When every one of those tokens names the same directory, Tenant ID is set to it, and the sign-in
record's multi-tenant URLs are pointed at it too — on Frappe 15, which cannot check the token
at sign-in, the URLs are the only protection there is. Anything less certain changes nothing:
no calendars, a token that cannot be read, tokens from several directories, personal accounts,
or a sign-in record that already names a different directory. The upgrade report then says
what to do by hand.

Reads stored tokens only; makes no call to Microsoft.
"""

import frappe

from frappe_microsoft365 import sign_in


def execute():
	if not frappe.db.exists("DocType", "Microsoft Settings") or not frappe.db.table_exists(
		"Social Login Key"
	):
		return

	tenant = frappe.db.get_single_value("Microsoft Settings", "tenant_id")
	if sign_in.tenant_kind(tenant) != sign_in.MULTI_TENANT_KIND:
		return

	# Only sites that sign in with Microsoft: elsewhere the multi-tenant authority opens nothing,
	# and changing it would only change which endpoint the calendars refresh at.
	key = frappe.db.get_value(
		"Social Login Key",
		sign_in.PROVIDER,
		["name", "client_id", "authorize_url", "access_token_url"],
		as_dict=True,
	)
	if not key:
		return

	tid = sign_in.calendar_tenant()
	if not tid:
		return

	stale = []
	for field in ("authorize_url", "access_token_url"):
		segment = sign_in.authority_tenant(key.get(field))
		if segment is None:
			continue
		if segment.lower() in sign_in.MULTI_TENANT:
			stale.append((field, segment))
		elif segment.lower() != tid:
			# The sign-in record already names a directory, and it is not the calendars' one (or
			# is a domain, which cannot be compared offline). Choosing between them is a person's
			# call; the upgrade report asks for it.
			return

	frappe.db.set_single_value("Microsoft Settings", "tenant_id", tid)

	# The URLs belong to whichever registration the record was made for; rewrite them only when
	# that is the one in Microsoft Settings, whose directory the calendars just told us.
	client_id = (frappe.db.get_single_value("Microsoft Settings", "client_id") or "").strip()
	if client_id and (key.client_id or "").strip() == client_id:
		for field, segment in stale:
			url = key.get(field)
			marker = f"login.microsoftonline.com/{segment}"
			start = url.lower().index(marker.lower())
			fixed = url[:start] + f"login.microsoftonline.com/{tid}" + url[start + len(marker) :]
			frappe.db.set_value("Social Login Key", key.name, field, fixed, update_modified=False)

	frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")
