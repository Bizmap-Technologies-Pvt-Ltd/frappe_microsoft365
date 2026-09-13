"""Microsoft Calendar — a per-user Microsoft 365 account connection.

Handles the OAuth authorize/callback round-trip (MSAL) and exposes the account for calendar sync
(Batch M2), Teams meetings (M3) and transcripts (M4). Mirrors Frappe's Google Calendar doctype.
"""

import frappe
from frappe import _
from frappe.model.document import Document
from frappe.utils import add_to_date, get_datetime, get_url_to_form, now_datetime

from frappe_microsoft365 import doctor
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
def authorize_access(calendar_name: str, reauthorize: int = 0):
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
def callback(
	code: str | None = None,
	state: str | None = None,
	error: str | None = None,
	error_description: str | None = None,
	**kwargs,
):
	"""OAuth redirect target. Exchanges the code for tokens and stores them on the matching doc.

	A browser lands here, so nothing may escape as a traceback. Microsoft's own failures are
	readable (AADSTS7000215 says in words that the client secret is wrong); dumping a Python
	stack on top of that hides the one useful sentence on the page.
	"""
	name = None
	try:
		if error:
			raise MicrosoftAuthError(error_description or error)

		name = frappe.db.get_value("Microsoft Calendar", {"oauth_state": state}) if state else None
		if not name:
			raise MicrosoftAuthError(
				_("This sign-in link is not valid any more. Click Authorize again to start a fresh one.")
			)

		doc = frappe.get_doc("Microsoft Calendar", name)
		_check_owner(doc)

		if not doc.oauth_state_expiry or get_datetime(doc.oauth_state_expiry) < now_datetime():
			frappe.db.set_value("Microsoft Calendar", name, "oauth_state", "", update_modified=False)
			# GET requests are never auto-committed, and the raise below would roll this back.
			frappe.db.commit()  # nosemgrep
			raise MicrosoftAuthError(
				_("This sign-in link has expired. Click Authorize again.")
			)

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

	except Exception as e:
		_render_auth_failure(name, e)
		return

	frappe.local.response["type"] = "redirect"
	frappe.local.response["location"] = get_url_to_form("Microsoft Calendar", name)


class MicrosoftAuthError(frappe.ValidationError):
	"""A sign-in failure worth showing the person in the browser, not logging and hiding."""


def _render_auth_failure(calendar_name, exception):
	"""Show a readable page instead of a traceback, and record the reason on the connection."""
	message = str(exception) or _("Microsoft sign-in failed.")
	frappe.log_error(title=f"MS Calendar authorization failed: {calendar_name or 'unknown'}")

	hint = doctor.explain_error(message)
	detail = hint["detail"] if hint.get("matched") else ""

	if calendar_name:
		# Surfaced on the form as well, so the reason survives closing this page.
		frappe.db.set_value(
			"Microsoft Calendar", calendar_name, "last_error", message[:500], update_modified=False
		)
		# The callback is a GET, which Frappe never auto-commits, and this reason has to
		# outlive the request so the form can show it.
		frappe.db.commit()  # nosemgrep

	body = f"<p>{frappe.utils.escape_html(message)}</p>"
	if detail:
		body += f"<p class='text-muted'>{frappe.utils.escape_html(detail)}</p>"

	frappe.respond_as_web_page(
		_("Microsoft sign-in failed"),
		body,
		indicator_color="red",
		primary_action=get_url_to_form("Microsoft Calendar", calendar_name)
		if calendar_name
		else "/app/microsoft-calendar",
		primary_label=_("Back to the connection"),
	)


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
def test_connection(calendar_name: str):
	"""Verify the stored token works by calling /me. Returns the account, never the token."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	me = graph.whoami(calendar_name)
	return {"ok": True, "account": me.get("mail") or me.get("userPrincipalName"), "display_name": me.get("displayName")}


@frappe.whitelist(methods=["POST"])
def disconnect(calendar_name: str):
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	frappe.db.set_value("Microsoft Calendar", calendar_name, {
		"authorized": 0, "access_token": "", "refresh_token": "", "token_expiry": None,
		"oauth_state": "", "oauth_state_expiry": None,
		"delta_link": "", "delta_window_end": None, "last_error": "",
	})
	# Tokens are also held in a request-local cache to keep a paging run from re-reading and
	# re-decrypting them on every call. Wiping the row without wiping that cache would leave a
	# live token in memory for the rest of this request — harmless today, because nothing here
	# calls Graph afterwards, and exactly the kind of thing that stops being harmless quietly.
	graph.clear_token_cache(calendar_name)
	return {"disconnected": True}


@frappe.whitelist()
def sync(calendar_name: str | None = None, run_inline: int = 0):
	"""Two-way sync entrypoint (M2). Owner-checked.

	Stays in the request whenever it can. The pulled/deleted/pushed counts are the reason
	anyone presses Sync Now, and a job queued in the background can only ever answer
	"queued" — so an incremental run, which is a second or two, is worth holding the request
	open for.

	A sync that is known in advance to be slow goes to a worker instead. A first sync walks up
	to 50 Graph pages plus a call per pending push; it outlives the gunicorn timeout, and the
	browser is handed a 504 while the sync quietly finishes server-side. Nothing is broken and
	it looks entirely broken. The reasons travel back with the answer so the form can say why
	it has no counts for the person this time.
	"""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	from frappe_microsoft365.microsoft_calendar_sync import (
		background_reasons,
		enqueue_sync,
		sync_blocked_by_dead_queue,
		sync_calendar,
	)

	# "Run now anyway", pressed by someone who has just been told the queue is dead. They are
	# choosing a long wait with their eyes open, which is a different thing from us choosing it
	# for them, so this is the one path that ignores how slow the sync is expected to be.
	if frappe.utils.cint(run_inline):
		return sync_calendar(calendar_name)

	reasons = background_reasons(doc)
	if reasons:
		# Asked before queueing, not after: frappe.enqueue succeeds perfectly well against a
		# Redis nobody is listening to, so a job id here would be a receipt for work that will
		# never happen. This is the one place a person is waiting to be told that.
		blocked = sync_blocked_by_dead_queue()
		if blocked:
			return {"queued": False, "blocked": True, "health": blocked, "reasons": reasons}
		# enqueue_sync owns the "queued" verdict; hard-coding True here would tell the person a
		# job had started even on the run where the queue refused it.
		return {**enqueue_sync(calendar_name), "reasons": reasons}

	return sync_calendar(calendar_name)
