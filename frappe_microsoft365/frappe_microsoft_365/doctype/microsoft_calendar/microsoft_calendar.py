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

#: Sync state that only the sync engine writes, always with frappe.db.set_value. The form shows
#: them read-only, but Frappe does not enforce read_only over the API, and the delta links are
#: URLs the worker later calls with the user's token attached — so a save never changes them.
SERVER_MANAGED_FIELDS = (
	"delta_link",
	"delta_window_end",
	"delta_next_link",
	"delta_next_window_end",
	# Written by the OAuth callback and the token refresh. The doctor trusts authorized_scopes
	# to say what a token can do, and the account id is how a re-authorisation tells whether it
	# is still the same mailbox, so neither may be typed in.
	"authorized_scopes",
	"microsoft_account_id",
	"authorization_error",
)

#: What a sync knows about one Microsoft mailbox. Signing in as a different account makes all of
#: it point at the wrong mailbox: the delta links would be replayed with the new account's token,
#: and the calendar id names a calendar the new account does not have.
ACCOUNT_SYNC_STATE = {
	"delta_link": "",
	"delta_window_end": None,
	"delta_next_link": None,
	"delta_next_window_end": None,
	"ms_calendar_id": "",
	"ms_calendar_name": "",
	# Push catches up edits made "since the last sync"; with another mailbox there has been none.
	"last_sync": None,
	"last_error": "",
}


class MicrosoftCalendar(Document):
	def before_insert(self):
		if not self.user:
			self.user = frappe.session.user

	def validate(self):
		self._keep_server_managed_fields()

	def _keep_server_managed_fields(self):
		"""Put the sync state back to what is stored, whatever the save carried.

		Restored rather than refused: a form opened before a sync finished carries the old
		values back on its next save, and that is no reason to fail the save — nor to let it
		roll the watermark back. A new connection starts with no sync state at all.
		"""
		meta = frappe.get_meta(self.doctype)
		fieldnames = [f for f in SERVER_MANAGED_FIELDS if meta.has_field(f)]
		if self.is_new():
			stored = {}
		else:
			# Read from the database when there is no snapshot (validate run by hand), never
			# assumed empty: blanking these on an ordinary save would restart the sync.
			stored = self.get_doc_before_save() or frappe.db.get_value(
				self.doctype, self.name, fieldnames, as_dict=True
			) or {}
		for fieldname in fieldnames:
			self.set(fieldname, stored.get(fieldname))

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
		previous = frappe.db.get_value(
			"Microsoft Calendar", name, _existing_fields("microsoft_account_id", "microsoft_user_email"),
			as_dict=True,
		) or {}
		# Committed below together with everything else a sign-in changes, so a failure part-way
		# can never leave the new account's tokens on the old account's sync state.
		graph._store_tokens(name, result, commit=False)
		me = _whoami_or_none(name)
		account_id = _account_id(result, me)

		updates = {
			"authorized": 1,
			"oauth_state": "",
			"oauth_state_expiry": None,
			"authorization_error": "",
			"authorized_scopes": " ".join(result.get("requested_scopes") or graph.get_scopes()),
			"microsoft_account_id": account_id or previous.get("microsoft_account_id") or "",
		}
		switched = _is_another_account(previous, account_id, result, me)
		if switched:
			updates.update(ACCOUNT_SYNC_STATE)
			email = _account_email(result, me)
			if email:
				updates["microsoft_user_email"] = email
		frappe.db.set_value("Microsoft Calendar", name, _only_existing(updates))
		unlinked = _unlink_events_of_previous_account(name) if switched else 0
		if switched:
			_note_account_switch(name, previous, updates.get("microsoft_user_email"), unlinked)
		# OAuth callback is a GET; without this the tokens we just stored would be discarded.
		frappe.db.commit()  # nosemgrep

		# best-effort: fill account email + default calendar
		try:
			_fill_account_details(name, me)
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


def _existing_fields(*fieldnames):
	"""The given fields that this site's Microsoft Calendar has (code can run before migrate)."""
	meta = frappe.get_meta("Microsoft Calendar")
	return [f for f in fieldnames if meta.has_field(f)]


def _only_existing(values):
	existing = set(_existing_fields(*values))
	return {k: v for k, v in values.items() if k in existing}


def _whoami_or_none(calendar_name):
	"""/me for the account that has just signed in, or None: identification must not fail sign-in."""
	try:
		return graph.whoami(calendar_name) or None
	except Exception:
		frappe.log_error(title=f"MS Calendar could not read the signed-in account: {calendar_name}")
		return None


def _account_id(result, me):
	"""The signed-in account's object id: from the ID token, else from /me (which is the same id)."""
	claims = result.get("id_token_claims") or {}
	return (claims.get("oid") or (me or {}).get("id") or "").strip()


def _account_email(result, me):
	me = me or {}
	claims = result.get("id_token_claims") or {}
	return me.get("mail") or me.get("userPrincipalName") or claims.get("preferred_username") or ""


def _account_names(result, me):
	"""Every address the signed-in account is known by, lower-cased."""
	me = me or {}
	claims = result.get("id_token_claims") or {}
	names = (
		me.get("mail"),
		me.get("userPrincipalName"),
		claims.get("preferred_username"),
		claims.get("email"),
		claims.get("upn"),
	)
	return {n.strip().lower() for n in names if n and n.strip()}


def _is_another_account(previous, account_id, result, me):
	"""Did this sign-in pick a different Microsoft account from the one the sync state belongs to?

	Compared by object id whenever both sides have one. A connection authorised before 1.2.0
	never stored an id, so it falls back to the address it recorded — matched against every
	address the new account goes by, because the form shows /me's mail while the ID token says
	the UPN, and one person's two addresses must not look like two people. With nothing to
	compare (a first sign-in) there is no old state to clear.
	"""
	stored_id = (previous.get("microsoft_account_id") or "").strip()
	if stored_id and account_id:
		return stored_id.lower() != account_id.lower()
	stored_email = (previous.get("microsoft_user_email") or "").strip().lower()
	names = _account_names(result, me)
	if stored_email and names:
		return stored_email not in names
	return False


def _unlink_events_of_previous_account(calendar_name):
	"""Detach the Events this connection had created in the previous account's mailbox.

	Their Microsoft ids name events in the OLD mailbox. Left alone, every edit would PATCH the
	new account with an id it does not have (404); cleared but still ticked, the next push would
	create each one again in the new mailbox and send every attendee a second invitation. So the
	link is dropped and the tick taken off: the meetings stay in the old mailbox, the Frappe
	Events stay as they are, and ticking one again sends it to the new account on purpose.

	Events pulled from the old mailbox are left untouched: a pull never pushes them back, and
	the new account's pull cannot match their ids, so they are inert history.
	"""
	filters = {
		"custom_microsoft_calendar": calendar_name,
		"custom_pulled_from_microsoft": 0,
		"custom_microsoft_event_id": ["is", "set"],
	}
	if not frappe.db.has_column("Event", "custom_microsoft_event_id"):
		return 0
	names = frappe.get_all("Event", filters=filters, pluck="name")
	for event_name in names:
		frappe.db.set_value(
			"Event",
			event_name,
			{"custom_microsoft_event_id": None, "custom_sync_with_microsoft_calendar": 0},
			update_modified=False,
		)
	return len(names)


def _note_account_switch(calendar_name, previous, new_email, unlinked):
	"""Leave a timeline comment: the person should not have to wonder why the sync restarted."""
	text = _(
		"Signed in as a different Microsoft account ({0}, previously {1}). The sync starts afresh "
		"from the new account's calendar."
	).format(new_email or _("unknown"), previous.get("microsoft_user_email") or _("unknown"))
	if unlinked:
		text += " " + _(
			"{0} Event(s) created in the previous account's calendar were unlinked and no longer "
			"sync; tick Sync with Microsoft Calendar on one to send it to the new account."
		).format(unlinked)
	try:
		frappe.get_doc("Microsoft Calendar", calendar_name).add_comment("Info", text)
	except Exception:
		frappe.log_error(title=f"MS Calendar account switch note failed: {calendar_name}")


def _fill_account_details(calendar_name, me=None):
	me = me or graph.whoami(calendar_name)
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
		**_resume_point_cleared(),
		**_only_existing({"authorization_error": ""}),
	})
	# Tokens are also held in a request-local cache to keep a paging run from re-reading and
	# re-decrypting them on every call. Wiping the row without wiping that cache would leave a
	# live token in memory for the rest of this request — harmless today, because nothing here
	# calls Graph afterwards, and exactly the kind of thing that stops being harmless quietly.
	graph.clear_token_cache(calendar_name)
	return {"disconnected": True}


def _resume_point_cleared():
	"""Values that drop an unfinished pull's resume point, on sites that have the columns."""
	meta = frappe.get_meta("Microsoft Calendar")
	return {
		fieldname: None
		for fieldname in ("delta_next_link", "delta_next_window_end")
		if meta.has_field(fieldname)
	}


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
