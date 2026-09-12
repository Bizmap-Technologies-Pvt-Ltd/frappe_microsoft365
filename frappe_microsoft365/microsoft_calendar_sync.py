"""Two-way calendar sync (Frappe Event <-> Microsoft Graph).

How the pull works
------------------
Graph is asked for a *delta* of the user's calendar view
(``/me/calendarView/delta``) rather than a "modified since" filter. That gives us three
things a filter cannot:

  * a stable id space — ``calendarView`` expands recurring series into occurrences, and a
    delta run keeps returning those same occurrence ids, so a recurring meeting is not
    duplicated on every run;
  * deletions — removed events arrive as ``@removed`` entries, so a hard-deleted Microsoft
    event no longer leaves an orphaned Frappe Event behind;
  * a watermark we only advance once every page has been consumed, so a failure mid-run
    means the next run repeats the work instead of skipping it.

Origin matters
--------------
An Event that originated in Frappe (we pushed it) keeps
``custom_pulled_from_microsoft = 0`` forever. The pull never flips that flag and never
overwrites the description of such an event, because Graph only hands back a truncated
plain-text ``bodyPreview``. Events that originated in Microsoft are mirrors and are fully
overwritten by the pull.

Attendees
---------
The pull also brings back the invitation itself: who organised the event, who was invited
and what each of them replied, plus our own ``responseStatus``. Those land on three
read-only Event fields so an Outlook invitation is legible from Frappe; replying to one is
``microsoft_rsvp``. The push sends Frappe's ``event_participants`` back as Graph attendees.

Third-party meeting links (Zoom, Google Meet) are *not* extracted: Microsoft only populates
``onlineMeeting`` for its own providers, and everything else sits as free text in the body.

All Graph calls go through ``microsoft_graph`` (auth/refresh/paging/clean errors).
Everything is guarded so an unconfigured / unauthorized site never raises on schedule.
"""

import frappe
from frappe import _
from frappe.utils import (
	add_to_date,
	get_datetime,
	get_system_timezone,
	now_datetime,
)

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError, MsGraphResyncRequired

EVENT_SELECT = (
	"id,subject,bodyPreview,start,end,location,isAllDay,isCancelled,"
	"onlineMeeting,webLink,type,lastModifiedDateTime,"
	# Graph returns only the properties asked for, so the invitation side of an event —
	# who was invited, who organised it and what we replied — has to be selected explicitly
	# or it never reaches Frappe at all.
	"attendees,organizer,isOrganizer,responseStatus,responseRequested"
)

#: How much of the calendar the delta window covers. Graph fixes the window when the delta
#: is initialised, so we re-initialise before the far edge gets close.
WINDOW_PAST_DAYS = 30
WINDOW_FUTURE_DAYS = 180
WINDOW_REFRESH_MARGIN_DAYS = 14

#: Ask Graph to hand back UTC so we never have to interpret a Windows timezone name.
UTC_PREFER = {"Prefer": 'odata.maxpagesize=50, outlook.timezone="UTC"'}

#: Graph may still echo a Windows timezone id (e.g. on data written by an Outlook client).
#: ZoneInfo cannot parse those, and silently falling back to UTC moves meetings by hours,
#: so the common ones are mapped explicitly.
WINDOWS_TO_IANA = {
	"UTC": "UTC",
	"GMT Standard Time": "Europe/London",
	"Greenwich Standard Time": "Atlantic/Reykjavik",
	"W. Europe Standard Time": "Europe/Berlin",
	"Central Europe Standard Time": "Europe/Budapest",
	"Central European Standard Time": "Europe/Warsaw",
	"Romance Standard Time": "Europe/Paris",
	"E. Europe Standard Time": "Europe/Chisinau",
	"FLE Standard Time": "Europe/Kiev",
	"GTB Standard Time": "Europe/Bucharest",
	"Turkey Standard Time": "Europe/Istanbul",
	"Israel Standard Time": "Asia/Jerusalem",
	"Arabian Standard Time": "Asia/Dubai",
	"Arab Standard Time": "Asia/Riyadh",
	"India Standard Time": "Asia/Kolkata",
	"Sri Lanka Standard Time": "Asia/Colombo",
	"Bangladesh Standard Time": "Asia/Dhaka",
	"SE Asia Standard Time": "Asia/Bangkok",
	"Singapore Standard Time": "Asia/Singapore",
	"China Standard Time": "Asia/Shanghai",
	"Tokyo Standard Time": "Asia/Tokyo",
	"Korea Standard Time": "Asia/Seoul",
	"AUS Eastern Standard Time": "Australia/Sydney",
	"AUS Central Standard Time": "Australia/Darwin",
	"W. Australia Standard Time": "Australia/Perth",
	"New Zealand Standard Time": "Pacific/Auckland",
	"Eastern Standard Time": "America/New_York",
	"US Eastern Standard Time": "America/Indiana/Indianapolis",
	"Central Standard Time": "America/Chicago",
	"Central Standard Time (Mexico)": "America/Mexico_City",
	"Mountain Standard Time": "America/Denver",
	"US Mountain Standard Time": "America/Phoenix",
	"Pacific Standard Time": "America/Los_Angeles",
	"Alaskan Standard Time": "America/Anchorage",
	"Hawaiian Standard Time": "Pacific/Honolulu",
	"Atlantic Standard Time": "America/Halifax",
	"SA Eastern Standard Time": "America/Cayenne",
	"E. South America Standard Time": "America/Sao_Paulo",
	"Argentina Standard Time": "America/Argentina/Buenos_Aires",
	"SA Pacific Standard Time": "America/Bogota",
	"South Africa Standard Time": "Africa/Johannesburg",
	"W. Central Africa Standard Time": "Africa/Lagos",
	"E. Africa Standard Time": "Africa/Nairobi",
	"Egypt Standard Time": "Africa/Cairo",
	"Morocco Standard Time": "Africa/Casablanca",
	"Russian Standard Time": "Europe/Moscow",
}


# --- datetime helpers ----------------------------------------------------------------

def _zone(tz_name):
	"""Resolve a Graph timezone name (IANA or Windows) to a ZoneInfo, defaulting to UTC."""
	from zoneinfo import ZoneInfo

	name = (tz_name or "UTC").strip()
	try:
		return ZoneInfo(name)
	except Exception:
		pass
	mapped = WINDOWS_TO_IANA.get(name)
	if mapped:
		try:
			return ZoneInfo(mapped)
		except Exception:
			pass
	# Unknown zone: log once per name so a wrong-by-hours event is traceable, not silent.
	frappe.log_error(title=f"MS Calendar: unknown timezone '{name}', assuming UTC")
	return ZoneInfo("UTC")


def _ms_dt_to_system(dt):
	"""Convert a Graph dateTimeTimeZone dict to a naive datetime in the system timezone."""
	from zoneinfo import ZoneInfo

	from dateutil import parser

	if not dt or not dt.get("dateTime"):
		return None
	parsed = parser.parse(dt["dateTime"])
	if parsed.tzinfo is None:
		parsed = parsed.replace(tzinfo=_zone(dt.get("timeZone")))
	return parsed.astimezone(ZoneInfo(get_system_timezone())).replace(tzinfo=None)


def _system_dt_to_ms(dt):
	"""Format a Frappe datetime as a Graph dateTimeTimeZone dict in the system timezone."""
	dt = get_datetime(dt)
	return {"dateTime": dt.isoformat(), "timeZone": get_system_timezone()}


def _iso_utc(dt):
	"""Render a naive/system datetime as a UTC ISO8601 string for Graph query params."""
	from zoneinfo import ZoneInfo

	dt = get_datetime(dt)
	if dt.tzinfo is None:
		dt = dt.replace(tzinfo=ZoneInfo(get_system_timezone()))
	return dt.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")


# --- concurrency ---------------------------------------------------------------------

def _calendar_lock(calendar_name):
	"""Per-calendar lock so a long sync is never overlapped by the next scheduled run."""
	try:
		from frappe.utils.synchronization import filelock

		return filelock(f"microsoft365_sync_{frappe.scrub(calendar_name)}", timeout=1)
	except ImportError:  # pragma: no cover - older Frappe without filelock
		from contextlib import nullcontext

		return nullcontext()


# --- entrypoints ---------------------------------------------------------------------

def sync_all():
	"""Scheduled entry: sync every enabled+authorized Microsoft Calendar. Never raises."""
	try:
		settings = frappe.get_cached_doc("Microsoft Settings")
		if not settings.enabled:
			return
	except Exception:
		return

	names = frappe.get_all(
		"Microsoft Calendar",
		filters={"enabled": 1, "authorized": 1},
		pluck="name",
	)
	results = []
	for name in names:
		try:
			results.append(sync_calendar(name))
		except Exception:
			frappe.log_error(title=f"MS Calendar sync failed: {name}")
	return results


def sync_calendar(calendar_name=None):
	"""Pull then push for a single Microsoft Calendar. Returns a JSON-serializable summary."""
	if not calendar_name:
		return {"ok": False, "pulled": 0, "deleted": 0, "pushed": 0, "message": "No calendar specified."}

	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	if not doc.enabled or not doc.authorized:
		return {
			"ok": False,
			"pulled": 0,
			"deleted": 0,
			"pushed": 0,
			"message": "Calendar is disabled or not authorized.",
		}

	from frappe.utils.file_lock import LockTimeoutError

	try:
		with _calendar_lock(calendar_name):
			return _sync_locked(doc)
	except LockTimeoutError:
		return {
			"ok": False,
			"pulled": 0,
			"deleted": 0,
			"pushed": 0,
			"message": "Another sync is already running for this calendar.",
		}


def _sync_locked(doc):
	calendar_name = doc.name
	pulled = deleted = pushed = 0
	messages = []
	pull_ok = True

	if doc.pull_from_microsoft_calendar:
		try:
			pulled, deleted = _pull(doc)
		except Exception as e:
			pull_ok = False
			messages.append(f"Pull failed: {e}")
			frappe.log_error(title=f"MS Calendar pull failed: {calendar_name}")

	if doc.push_to_microsoft_calendar:
		try:
			pushed = _push(doc)
		except Exception as e:
			messages.append(f"Push failed: {e}")
			frappe.log_error(title=f"MS Calendar push failed: {calendar_name}")

	# The watermark only moves when the pull actually completed. Advancing it after a
	# failure would silently skip every change made inside the failed window.
	updates = {"last_error": "; ".join(messages)[:500] or ""}
	if pull_ok:
		updates["last_sync"] = now_datetime()
	frappe.db.set_value("Microsoft Calendar", calendar_name, updates, update_modified=False)
	# sync_all loops over every calendar in one job: committing here keeps this calendar's
	# watermark even if a later calendar raises, so its window is not re-fetched forever.
	frappe.db.commit()  # nosemgrep

	return {
		"ok": not messages,
		"pulled": pulled,
		"deleted": deleted,
		"pushed": pushed,
		"message": "; ".join(messages) or f"Pulled {pulled}, deleted {deleted}, pushed {pushed}.",
	}


# --- pull (Graph -> Frappe) ----------------------------------------------------------

def _initial_delta_path():
	start = _iso_utc(add_to_date(now_datetime(), days=-WINDOW_PAST_DAYS))
	end = _iso_utc(add_to_date(now_datetime(), days=WINDOW_FUTURE_DAYS))
	return (
		f"/me/calendarView/delta?startDateTime={start}&endDateTime={end}"
		f"&$select={EVENT_SELECT}"
	)


def _delta_window_is_stale(doc):
	"""True when the stored delta window is close to its end and must be re-initialised."""
	if not doc.delta_window_end:
		return True
	margin = add_to_date(now_datetime(), days=WINDOW_REFRESH_MARGIN_DAYS)
	return get_datetime(doc.delta_window_end) <= margin


def _pull(doc):
	"""Pull changes from Microsoft into Frappe. Returns (upserted, deleted)."""
	reuse_delta = doc.delta_link and not _delta_window_is_stale(doc)
	path = doc.delta_link if reuse_delta else _initial_delta_path()

	try:
		items, delta_link = graph.graph_delta(path, doc.name, headers=UTC_PREFER)
	except MsGraphResyncRequired:
		# Token expired or the mailbox was moved: start the window again from scratch.
		items, delta_link = graph.graph_delta(_initial_delta_path(), doc.name, headers=UTC_PREFER)
		reuse_delta = False

	frappe.flags.in_microsoft_sync = True
	upserted = removed = 0
	try:
		for ev in items:
			try:
				outcome = _upsert_event(doc, ev)
				if outcome == "deleted":
					removed += 1
				elif outcome in ("created", "updated"):
					upserted += 1
			except Exception:
				frappe.log_error(title=f"MS event upsert failed: {ev.get('id')}")
	finally:
		frappe.flags.in_microsoft_sync = False

	# Only store the new watermark once every page was consumed without raising. Without a
	# delta link we simply redo this window next run, which repeats work but loses nothing.
	if delta_link:
		updates = {"delta_link": delta_link}
		if not reuse_delta:
			updates["delta_window_end"] = add_to_date(now_datetime(), days=WINDOW_FUTURE_DAYS)
		frappe.db.set_value("Microsoft Calendar", doc.name, updates, update_modified=False)

	return upserted, removed


def _is_removed(ev):
	return bool(ev.get("@removed")) or bool(ev.get("isCancelled"))


# --- attendees (Graph -> Frappe) -----------------------------------------------------

#: Graph's documented responseStatus.response values, which are also the options of the
#: custom_microsoft_my_response Select. Anything outside this set would fail Frappe's
#: Select validation and take the whole event's sync down with it over a display field,
#: so an unrecognised value is dropped instead of stored.
RESPONSE_STATUSES = (
	"none",
	"organizer",
	"tentativelyAccepted",
	"accepted",
	"declined",
	"notResponded",
)


def _response_labels():
	"""Graph ``responseStatus`` values in plain words.

	Built inside a function, never at module level: ``frappe._()`` evaluated at import time
	resolves once, in whatever language the worker booted with, and would then freeze that
	language into every site's summaries.
	"""
	return {
		"none": _("no response"),
		"organizer": _("organizer"),
		"notResponded": _("not responded"),
		"tentativelyAccepted": _("tentative"),
		"accepted": _("accepted"),
		"declined": _("declined"),
	}


def _attendee_line(attendee, labels):
	"""One Graph attendee as a single readable line, or None if it carries no address."""
	email_address = attendee.get("emailAddress") or {}
	address = (email_address.get("address") or "").strip()
	name = (email_address.get("name") or "").strip()

	# Angle brackets would be the conventional way to write this, but Frappe sanitises HTML
	# on save and silently eats "<asha@x.com>" as an unknown tag, losing the address
	# entirely. Parentheses survive.
	#
	# Outlook also fills `name` with the address itself for people outside the tenant, and
	# printing "asha@x.com (asha@x.com)" helps nobody.
	if name and address and name.lower() != address.lower():
		who = f"{name} ({address})"
	else:
		who = address or name
	if not who:
		return None

	# Rooms and equipment are invited exactly like people (type "resource"), and a decline
	# from a room means something different from a decline from a person, so the type is
	# worth showing whenever it is not the ordinary "required". It rides with the response
	# rather than after the address, to avoid a second bracketed group.
	kind = (attendee.get("type") or "").strip()
	suffix = kind if kind and kind != "required" else ""

	label = labels.get(((attendee.get("status") or {}).get("response") or "").strip())
	trailer = ", ".join(part for part in (suffix, label) if part)
	return f"{who} — {trailer}" if trailer else who


def _attendee_summary(ev):
	"""Everyone on the Microsoft invitation, one per line, with their reply."""
	labels = _response_labels()
	lines = []
	for attendee in ev.get("attendees") or []:
		line = _attendee_line(attendee, labels)
		if line:
			lines.append(line)
	return "\n".join(lines)


def _people_values(ev):
	"""Organizer / attendees / our own response as Frappe Event field values.

	Kept apart from the rest of the mapping so the caller can apply the origin rule: an
	empty value here means "Graph did not carry this" just as readily as "nobody is
	invited", and the payload does not distinguish the two.
	"""
	organizer = ((ev.get("organizer") or {}).get("emailAddress") or {}).get("address") or ""
	my_response = ((ev.get("responseStatus") or {}).get("response") or "").strip()
	return {
		"custom_microsoft_organizer": organizer.strip(),
		"custom_microsoft_attendees": _attendee_summary(ev),
		"custom_microsoft_my_response": my_response if my_response in RESPONSE_STATUSES else "",
	}


def _target_values(doc, ev, locally_originated):
	"""The Frappe Event field values a Microsoft event maps to."""
	values = {
		"subject": ev.get("subject") or "(No subject)",
		"all_day": 1 if ev.get("isAllDay") else 0,
		"custom_microsoft_calendar": doc.name,
		"custom_sync_with_microsoft_calendar": 1,
	}
	starts_on = _ms_dt_to_system(ev.get("start"))
	ends_on = _ms_dt_to_system(ev.get("end"))
	if starts_on:
		values["starts_on"] = starts_on
	if ends_on:
		values["ends_on"] = ends_on

	location = (ev.get("location") or {}).get("displayName")
	if location is not None:
		values["location"] = location

	# A meeting organised in Outlook keeps its join link when it lands in Frappe.
	join_url = (ev.get("onlineMeeting") or {}).get("joinUrl")
	if join_url:
		values["custom_teams_join_url"] = join_url
		values["custom_add_teams_meeting"] = 1
	if ev.get("webLink"):
		values["custom_microsoft_web_link"] = ev["webLink"]

	# bodyPreview is a truncated plain-text preview. Writing it onto an Event that
	# originated in Frappe would destroy the real description, so mirrors only.
	if not locally_originated:
		values["description"] = ev.get("bodyPreview") or ""

	# The same origin rule covers the invitation fields. On a mirror, empty means empty and
	# the field is cleared; on an Event that originated in Frappe, empty far more likely
	# means Graph did not carry the property, and blanking would throw away what the user
	# can plainly see in Outlook.
	for field, value in _people_values(ev).items():
		if value or not locally_originated:
			values[field] = value

	return values


def _upsert_event(doc, ev):
	"""Create/update/delete the Frappe mirror of one Microsoft event.

	Returns "deleted", "updated", "created" or "skipped".
	"""
	ms_id = ev.get("id")
	if not ms_id:
		return "skipped"

	existing = frappe.db.get_value("Event", {"custom_microsoft_event_id": ms_id}, "name")

	if _is_removed(ev):
		if existing:
			frappe.delete_doc("Event", existing, ignore_permissions=True, force=True)
			return "deleted"
		return "skipped"

	# calendarView returns occurrences; a seriesMaster would duplicate all of them.
	if ev.get("type") == "seriesMaster":
		return "skipped"

	if existing:
		event = frappe.get_doc("Event", existing)
		locally_originated = not event.custom_pulled_from_microsoft
	else:
		event = frappe.new_doc("Event")
		event.custom_microsoft_event_id = ms_id
		event.custom_pulled_from_microsoft = 1
		event.event_type = "Private"
		locally_originated = False

	values = _target_values(doc, ev, locally_originated)
	changed = _apply(event, values)
	if not existing:
		changed = True
	if not changed:
		# Nothing moved. Saving anyway would bump `modified` and make the push step
		# re-patch this event on the next run, forever.
		return "skipped"

	event.flags.ignore_permissions = True
	event.flags.ignore_mandatory = True
	event.save()
	return "updated" if existing else "created"


def _apply(doc, values):
	"""Set values on a doc, returning True if anything actually changed."""
	changed = False
	for field, value in values.items():
		current = doc.get(field)
		if isinstance(value, str) or value is None:
			same = (current or "") == (value or "")
		elif field in ("starts_on", "ends_on"):
			same = bool(current) and get_datetime(current) == get_datetime(value)
		else:
			same = current == value
		if not same:
			doc.set(field, value)
			changed = True
	return changed


# --- push (Frappe -> Graph) ----------------------------------------------------------

def _push(doc):
	"""Create new Frappe events in Microsoft and patch ones edited while offline."""
	base_filters = {
		"custom_sync_with_microsoft_calendar": 1,
		"custom_microsoft_calendar": doc.name,
		"custom_pulled_from_microsoft": 0,
	}

	count = 0
	new_events = frappe.get_all(
		"Event",
		filters={**base_filters, "custom_microsoft_event_id": ["is", "not set"]},
		pluck="name",
	)
	for name in new_events:
		try:
			event = frappe.get_doc("Event", name)
			created = _create_graph_event(doc.name, event)
			if _store_graph_response(name, created, doc.name):
				count += 1
		except Exception:
			frappe.log_error(title=f"MS event push failed: {name}")

	# Edits made while the connection was down never reached doc_events; catch them up.
	if doc.last_sync:
		edited = frappe.get_all(
			"Event",
			filters={
				**base_filters,
				# NOT IN with NULL never matches in SQL; "is set" is the NULL-safe form.
				"custom_microsoft_event_id": ["is", "set"],
				"modified": [">", doc.last_sync],
			},
			pluck="name",
		)
		for name in edited:
			try:
				event = frappe.get_doc("Event", name)
				patched = graph.graph_request(
					"PATCH",
					f"/me/events/{event.custom_microsoft_event_id}",
					doc.name,
					json=_event_to_graph_body(event),
				)
				_store_graph_response(name, patched or {}, doc.name)
				count += 1
			except Exception:
				frappe.log_error(title=f"MS event patch failed: {name}")

	# The ids of events already created in Graph must survive a later failure in this job;
	# losing them would create duplicates in Outlook on the next run.
	frappe.db.commit()  # nosemgrep
	return count


def _email_fields(doctype):
	"""Data fields on a doctype declared as an Email, in field order.

	Resolved by fieldtype rather than from a hardcoded doctype list: an Event participant
	is a Dynamic Link and can point at anything — Contact, Lead, Employee, something from
	an app this one has never heard of — and a fixed list would quietly stop resolving the
	moment somebody links a doctype that is not on it.
	"""
	try:
		meta = frappe.get_meta(doctype)
	except Exception:
		return []
	return [f.fieldname for f in meta.fields if f.fieldtype == "Data" and (f.options or "") == "Email"]


def _participant_email(participant):
	"""Resolve one Event Participants row to an email address, or None.

	The row has its own ``email`` column, but Frappe only fills it in for links it can
	resolve to a Contact, rows written by other apps routinely leave it blank, and older
	Frappe does not populate it at all — so the referenced document is the fallback.
	"""
	email = (participant.get("email") or "").strip()
	if email:
		return email

	doctype = participant.get("reference_doctype")
	docname = participant.get("reference_docname")
	if not doctype or not docname:
		return None

	# A Frappe User is named by its email address, so the link itself is the answer.
	if doctype == "User":
		return (frappe.db.get_value("User", docname, "email") or docname or "").strip() or None

	fields = _email_fields(doctype)
	if not fields:
		return None
	row = frappe.db.get_value(doctype, docname, fields, as_dict=True) or {}
	for fieldname in fields:
		value = (row.get(fieldname) or "").strip()
		if value:
			return value
	return None


def _participant_name(participant):
	"""A display name for an attendee. Graph treats emailAddress.name as optional."""
	doctype = participant.get("reference_doctype")
	docname = participant.get("reference_docname")
	if not doctype or not docname:
		return None
	if doctype == "User":
		return frappe.db.get_value("User", docname, "full_name") or None
	try:
		title_field = frappe.get_meta(doctype).get_title_field()
	except Exception:
		return None
	if not title_field or title_field == "name":
		return docname
	return frappe.db.get_value(doctype, docname, title_field) or docname


def _graph_attendees(event):
	"""Frappe's Event Participants as Graph attendees, skipping any that cannot be addressed."""
	attendees = []
	seen = set()
	for participant in event.get("event_participants") or []:
		email = _participant_email(participant)
		if not email:
			# Graph rejects an attendee with no address, and guessing one would send a real
			# invitation to the wrong person. An unresolvable participant is left out.
			continue
		key = email.lower()
		if key in seen:
			# Two participant rows can point at different records with the same address
			# (a Contact and the User behind it); Outlook would show the person twice.
			continue
		seen.add(key)
		entry = {"emailAddress": {"address": email}, "type": "required"}
		name = _participant_name(participant)
		if name and name != email:
			entry["emailAddress"]["name"] = name
		attendees.append(entry)
	return attendees


def _end_after_start(starts_on, ends_on):
	"""Graph rejects an event whose end is not after its start, with ErrorPropertyValidationFailure.

	Frappe does not enforce the ordering, and its Event form pre-fills both from "now", so a
	record saved without touching the times can end a few seconds before it starts. Rather
	than fail the push over that, fall back to a half hour from the start, which is what an
	empty end already does.
	"""
	ends_on = get_datetime(ends_on) if ends_on else None
	if ends_on and ends_on > starts_on:
		return ends_on
	return add_to_date(starts_on, minutes=30)


def _event_to_graph_body(event):
	starts_on = get_datetime(event.starts_on)
	body = {
		"subject": event.subject or "(No subject)",
		"body": {"contentType": "HTML", "content": event.description or ""},
		"start": _system_dt_to_ms(starts_on),
		"end": _system_dt_to_ms(_end_after_start(starts_on, event.ends_on)),
	}
	if getattr(event, "all_day", 0):
		body["isAllDay"] = True
	if getattr(event, "location", None):
		body["location"] = {"displayName": event.location}

	# Only ever sent as True. Microsoft does not support turning an existing online meeting
	# back into a plain event, so sending False would silently do nothing and imply otherwise.
	if getattr(event, "custom_add_teams_meeting", 0):
		body["isOnlineMeeting"] = True
		body["onlineMeetingProvider"] = "teamsForBusiness"

	# Only sent when at least one participant resolved to an address. Graph reads an empty
	# attendees array as "remove everyone", and "this Event has no participants in Frappe"
	# is indistinguishable from "its attendees were added in Outlook", so sending [] would
	# quietly uninvite people nobody asked to uninvite.
	attendees = _graph_attendees(event)
	if attendees:
		body["attendees"] = attendees

	return body


def _store_graph_response(event_name, created, calendar_name=None, live_doc=None):
	"""Save the bits Graph fills in itself: the event id, join link and Outlook link.

	The id is the link between the two sides, and recording it is not optional. If it does not
	persist, the next run sees an Event with no id and creates the meeting AGAIN — which is
	how three copies of one meeting ended up in a real calendar when this column was too
	narrow to hold a Graph id.

	So the write is verified, and when it fails the event we just created is removed from the
	calendar again rather than left as an orphan for the next run to duplicate. Passing
	``calendar_name`` enables that rollback; without it the failure is only reported.
	``live_doc`` is the in-memory Event when we are inside its own save, so the values Graph
	filled in show up on screen without a reload.
	"""
	values = {}
	if created.get("id"):
		values["custom_microsoft_event_id"] = created["id"]
	join_url = (created.get("onlineMeeting") or {}).get("joinUrl")
	if join_url:
		values["custom_teams_join_url"] = join_url
		values["custom_add_teams_meeting"] = 1
	if created.get("webLink"):
		values["custom_microsoft_web_link"] = created["webLink"]

	if not values:
		return False

	ms_id = values.get("custom_microsoft_event_id")
	try:
		frappe.db.set_value("Event", event_name, values, update_modified=False)
	except Exception:
		frappe.log_error(title=f"MS event link could not be stored: {event_name}")
		_undo_orphaned_graph_event(calendar_name, ms_id, event_name)
		return False

	if not ms_id:
		return False

	# db.set_value writes the row, but the document being saved is still in memory and is what
	# the browser gets back. Without this the join link exists in the database and stays
	# invisible on screen until someone reloads the page.
	if live_doc is not None:
		for field, value in values.items():
			live_doc.set(field, value)

	# Read it back. A write can be accepted and still not round-trip (truncation, sanitising),
	# and a half-stored link is indistinguishable from no link on the next run.
	if frappe.db.get_value("Event", event_name, "custom_microsoft_event_id") != ms_id:
		frappe.log_error(
			title=f"MS event link did not round-trip: {event_name}",
			message=f"Graph returned id of {len(ms_id)} characters; reading it back gave something else.",
		)
		_undo_orphaned_graph_event(calendar_name, ms_id, event_name)
		return False

	return True


def _undo_orphaned_graph_event(calendar_name, ms_id, event_name):
	"""Remove an event we created but could not link, so the next run cannot duplicate it."""
	if not calendar_name or not ms_id:
		frappe.log_error(
			title=f"MS event orphaned in the calendar: {event_name}",
			message=(
				f"Created in Microsoft as {ms_id} but not linked in Frappe, and no calendar was "
				"available to undo it. Delete it in Outlook by hand, or the next sync will "
				"create another copy."
			),
		)
		return

	try:
		graph.graph_request("DELETE", f"/me/events/{ms_id}", calendar_name)
	except Exception:
		frappe.log_error(title=f"MS orphaned event could not be removed: {event_name}")


def _create_graph_event(calendar_name, event):
	return graph.graph_request(
		"POST", "/me/events", calendar_name, json=_event_to_graph_body(event)
	)


# --- doc_events (single Frappe Event lifecycle -> Graph) -----------------------------

def event_validate(doc, method=None):
	"""Refuse a time range Microsoft will reject, while the person can still fix it.

	Frappe does not enforce that an Event ends after it starts, and its form pre-fills both
	from the current moment, so an event saved without touching the times can end seconds
	before it begins. Graph answers ErrorPropertyValidationFailure, by which point the save
	has succeeded and the failure is buried in the Error Log.

	Only events actually bound for Microsoft are checked: this app has no business dictating
	what an unrelated Event may contain.
	"""
	# Never applied to data arriving FROM Microsoft. Outlook is authoritative during a pull,
	# and refusing what it sends would stall the sync on a record we cannot fix from here.
	# This is a guard on what a person types into Frappe, nothing else.
	if frappe.flags.in_microsoft_sync:
		return
	if not getattr(doc, "custom_sync_with_microsoft_calendar", 0):
		return
	if not doc.starts_on or not doc.ends_on:
		return

	if get_datetime(doc.ends_on) <= get_datetime(doc.starts_on):
		frappe.throw(
			_("This event ends before it starts, and Microsoft will not accept it. Set an end time after {0}.").format(
				frappe.utils.format_datetime(doc.starts_on)
			),
			title=_("Check the times"),
		)


def event_on_update(doc, method=None):
	"""Push/patch a single Event to Graph on save. Best-effort, never blocks the save."""
	if frappe.flags.in_microsoft_sync:
		return
	if not getattr(doc, "custom_sync_with_microsoft_calendar", 0):
		return
	if not getattr(doc, "custom_microsoft_calendar", None):
		return
	if getattr(doc, "custom_pulled_from_microsoft", 0):
		return

	try:
		cal = frappe.get_cached_doc("Microsoft Calendar", doc.custom_microsoft_calendar)
		if not cal.enabled or not cal.authorized or not cal.push_to_microsoft_calendar:
			return

		if doc.custom_microsoft_event_id:
			patched = graph.graph_request(
				"PATCH",
				f"/me/events/{doc.custom_microsoft_event_id}",
				cal.name,
				json=_event_to_graph_body(doc),
			)
			_store_graph_response(doc.name, patched or {}, cal.name, live_doc=doc)
		else:
			_store_graph_response(doc.name, _create_graph_event(cal.name, doc), cal.name, live_doc=doc)
	except Exception:
		frappe.log_error(title=f"MS Event on_update sync failed: {doc.name}")


def event_on_trash(doc, method=None):
	"""Delete the mirrored Microsoft event when the Frappe Event is deleted. Best-effort."""
	if frappe.flags.in_microsoft_sync:
		return
	ms_id = getattr(doc, "custom_microsoft_event_id", None)
	cal_name = getattr(doc, "custom_microsoft_calendar", None)
	if not ms_id or not cal_name:
		return
	try:
		cal = frappe.get_cached_doc("Microsoft Calendar", cal_name)
		if not cal.enabled or not cal.authorized:
			return
		graph.graph_request("DELETE", f"/me/events/{ms_id}", cal.name)
	except Exception:
		frappe.log_error(title=f"MS Event on_trash sync failed: {doc.name}")


# --- read-only fetch for external consumers (e.g. Bizmap CRM) ------------------------

@frappe.whitelist()
def fetch_events(calendar_name: str, start_datetime: str | None = None, end_datetime: str | None = None):
	"""Return events from Graph (calendarView) as a clean list. Owner-checked, fully paged."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)

	start = _iso_utc(start_datetime) if start_datetime else _iso_utc(now_datetime())
	end = _iso_utc(end_datetime) if end_datetime else _iso_utc(add_to_date(now_datetime(), days=30))

	path = (
		f"/me/calendarView?startDateTime={start}&endDateTime={end}"
		f"&$select={EVENT_SELECT}&$orderby=start/dateTime"
	)
	items = graph.graph_paged(path, calendar_name, headers=UTC_PREFER)

	out = []
	for ev in items:
		online = ev.get("onlineMeeting") or {}
		out.append(
			{
				"id": ev.get("id"),
				"subject": ev.get("subject"),
				"start": _ms_dt_to_system(ev.get("start")),
				"end": _ms_dt_to_system(ev.get("end")),
				"is_all_day": bool(ev.get("isAllDay")),
				"location": (ev.get("location") or {}).get("displayName"),
				"join_url": online.get("joinUrl"),
				"web_link": ev.get("webLink"),
			}
		)
	return out


def _check_owner(doc):
	if frappe.session.user == "Administrator" or "System Manager" in frappe.get_roles():
		return
	if doc.user and doc.user != frappe.session.user:
		frappe.throw(_("You can only manage your own Microsoft Calendar."), frappe.PermissionError)
