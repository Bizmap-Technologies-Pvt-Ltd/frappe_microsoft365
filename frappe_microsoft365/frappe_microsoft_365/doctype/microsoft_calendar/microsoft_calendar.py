"""Microsoft Calendar — a per-user Microsoft 365 account connection.

Handles the OAuth authorize/callback round-trip (MSAL) and exposes the account for calendar sync
(Batch M2), Teams meetings (M3) and transcripts (M4). Mirrors Frappe's Google Calendar doctype.
"""

import frappe
from frappe import _
from frappe.model.document import Document
from frappe.utils import add_to_date, get_datetime, now_datetime

from frappe_microsoft365 import microsoft_graph as graph

#: An authorize link that is never followed should not stay usable forever.
OAUTH_STATE_TTL_MINUTES = 15


class MicrosoftCalendar(Document):
	def before_insert(self):
		if not self.user:
			self.user = frappe.session.user

	def get_access_token(self):
		return graph.get_valid_access_token(self.name)


def _check_owner(doc):
	if frappe.session.user == "Administrator" or "System Manager" in frappe.get_roles():
		return
	if doc.user and doc.user != frappe.session.user:
		frappe.throw(_("You can only manage your own Microsoft Calendar."), frappe.PermissionError)


# --- OAuth round-trip ----------------------------------------------------------------

@frappe.whitelist(methods=["POST"])
def authorize_access(calendar_name, reauthorize=0):
	"""Return the Microsoft sign-in URL for this calendar; the form redirects the user to it."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	state = frappe.generate_hash(length=32)
	frappe.db.set_value(
		"Microsoft Calendar",
		calendar_name,
		{
			"oauth_state": state,
			"oauth_state_expiry": add_to_date(now_datetime(), minutes=OAUTH_STATE_TTL_MINUTES),
		},
		update_modified=False,
	)
	return {"url": graph.build_authorize_url(state)}


@frappe.whitelist()
def callback(code=None, state=None, error=None, error_description=None, **kwargs):
	"""OAuth redirect target. Exchanges the code for tokens and stores them on the matching doc."""
	if error:
		frappe.local.response["type"] = "redirect"
		frappe.local.response["location"] = f"/app/microsoft-calendar?error={frappe.utils.quoted(error)}"
		return

	name = frappe.db.get_value("Microsoft Calendar", {"oauth_state": state}) if state else None
	if not name:
		frappe.throw(_("Invalid or expired authorization state. Please try authorizing again."))

	doc = frappe.get_doc("Microsoft Calendar", name)
	_check_owner(doc)

	if not doc.oauth_state_expiry or get_datetime(doc.oauth_state_expiry) < now_datetime():
		frappe.db.set_value("Microsoft Calendar", name, "oauth_state", "", update_modified=False)
		# GET requests are never auto-committed, and the throw below would roll this back.
		frappe.db.commit()  # nosemgrep
		frappe.throw(_("This authorization link has expired. Please click Authorize again."))

	result = graph.exchange_code(code)
	graph._store_tokens(name, result)  # stores access/refresh/expiry + email from claims
	frappe.db.set_value(
		"Microsoft Calendar", name, {"authorized": 1, "oauth_state": "", "oauth_state_expiry": None}
	)
	# OAuth callback is a GET; without this the tokens we just stored would be discarded.
	frappe.db.commit()  # nosemgrep

	# best-effort: fill account email + default calendar
	try:
		_fill_account_details(name)
	except Exception:
		frappe.log_error(title="MS Calendar post-auth detail fetch failed")

	frappe.local.response["type"] = "redirect"
	frappe.local.response["location"] = f"/app/microsoft-calendar/{name}"


def _fill_account_details(calendar_name):
	me = graph.whoami(calendar_name)
	updates = {}
	email = me.get("mail") or me.get("userPrincipalName")
	if email:
		updates["microsoft_user_email"] = email
	# default calendar
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	if not doc.ms_calendar_id:
		cals = graph.graph_request("GET", "/me/calendars?$select=id,name,isDefaultCalendar", calendar_name)
		default = next((c for c in cals.get("value", []) if c.get("isDefaultCalendar")), None) \
			or (cals.get("value") or [None])[0]
		if default:
			updates["ms_calendar_id"] = default["id"]
			updates["ms_calendar_name"] = default.get("name")
	if updates:
		frappe.db.set_value("Microsoft Calendar", calendar_name, updates)
		# Still inside the GET callback request, which Frappe does not auto-commit.
		frappe.db.commit()  # nosemgrep


@frappe.whitelist()
def test_connection(calendar_name):
	"""Verify the stored token works by calling /me. Returns the account, never the token."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	me = graph.whoami(calendar_name)
	return {"ok": True, "account": me.get("mail") or me.get("userPrincipalName"), "display_name": me.get("displayName")}


@frappe.whitelist(methods=["POST"])
def disconnect(calendar_name):
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	frappe.db.set_value("Microsoft Calendar", calendar_name, {
		"authorized": 0, "access_token": "", "refresh_token": "", "token_expiry": None,
		"oauth_state": "", "oauth_state_expiry": None,
		"delta_link": "", "delta_window_end": None, "last_error": "",
	})
	return {"disconnected": True}


@frappe.whitelist()
def sync(calendar_name=None):
	"""Two-way sync entrypoint (M2). Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	from frappe_microsoft365.microsoft_calendar_sync import sync_calendar
	return sync_calendar(calendar_name)
