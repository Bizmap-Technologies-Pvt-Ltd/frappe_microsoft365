"""Two-way calendar sync (Frappe Event <-> Microsoft Graph). Stage M2.

Mirrors Frappe's Google Calendar sync:
  * Pull: Graph events changed since last_sync are upserted as Frappe Events.
  * Push: Frappe Events flagged for sync (and not pulled) are created in Graph.
  * doc_events keep individual Frappe Event edits/deletes in sync (best-effort).

All Graph calls go through ``microsoft_graph.graph_request`` (auth/refresh/clean errors).
Everything is guarded so an unconfigured / unauthorized site never raises on schedule.
"""

import datetime

import frappe
from frappe import _
from frappe.utils import (
	add_to_date,
	get_datetime,
	get_system_timezone,
	now_datetime,
)

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError

EVENT_SELECT = (
	"id,subject,bodyPreview,start,end,location,isAllDay,isCancelled,"
	"onlineMeeting,webLink,lastModifiedDateTime"
)


# --- datetime helpers ----------------------------------------------------------------

def _ms_dt_to_system(dt):
	"""Convert a Graph dateTimeTimeZone dict to a naive datetime in the system timezone."""
	from dateutil import parser
	from zoneinfo import ZoneInfo

	if not dt or not dt.get("dateTime"):
		return None
	raw = dt["dateTime"]
	tz = dt.get("timeZone") or "UTC"
	parsed = parser.parse(raw)
	if parsed.tzinfo is None:
		# Graph timeZone is an IANA/Windows name; UTC is the common default
		try:
			parsed = parsed.replace(tzinfo=ZoneInfo(tz))
		except Exception:
			parsed = parsed.replace(tzinfo=ZoneInfo("UTC"))
	return parsed.astimezone(ZoneInfo(get_system_timezone())).replace(tzinfo=None)


def _system_dt_to_ms(dt):
	"""Format a Frappe datetime as a Graph dateTimeTimeZone dict in the system timezone."""
	dt = get_datetime(dt)
	return {"dateTime": dt.isoformat(), "timeZone": get_system_timezone()}


def _iso_utc(dt):
	"""Render a naive/system datetime as a UTC ISO8601 string for $filter."""
	from zoneinfo import ZoneInfo

	dt = get_datetime(dt)
	if dt.tzinfo is None:
		dt = dt.replace(tzinfo=ZoneInfo(get_system_timezone()))
	return dt.astimezone(ZoneInfo("UTC")).strftime("%Y-%m-%dT%H:%M:%SZ")


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
		return {"ok": False, "pulled": 0, "pushed": 0, "message": "No calendar specified."}

	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	if not doc.enabled or not doc.authorized:
		return {
			"ok": False,
			"pulled": 0,
			"pushed": 0,
			"message": "Calendar is disabled or not authorized.",
		}

	pulled = pushed = 0
	messages = []

	if doc.pull_from_microsoft_calendar:
		try:
			pulled = _pull(doc)
		except MsGraphError as e:
			messages.append(f"Pull failed: {e}")
			frappe.log_error(title=f"MS Calendar pull failed: {calendar_name}")
		except Exception as e:
			messages.append(f"Pull error: {e}")
			frappe.log_error(title=f"MS Calendar pull error: {calendar_name}")

	if doc.push_to_microsoft_calendar:
		try:
			pushed = _push(doc)
		except MsGraphError as e:
			messages.append(f"Push failed: {e}")
			frappe.log_error(title=f"MS Calendar push failed: {calendar_name}")
		except Exception as e:
			messages.append(f"Push error: {e}")
			frappe.log_error(title=f"MS Calendar push error: {calendar_name}")

	frappe.db.set_value("Microsoft Calendar", calendar_name, "last_sync", now_datetime())
	frappe.db.commit()

	return {
		"ok": not messages,
		"pulled": pulled,
		"pushed": pushed,
		"message": "; ".join(messages) or f"Pulled {pulled}, pushed {pushed}.",
	}


# --- pull (Graph -> Frappe) ----------------------------------------------------------

def _pull(doc):
	"""Pull changed Microsoft events into Frappe. Returns count upserted/processed."""
	if doc.last_sync:
		since = _iso_utc(doc.last_sync)
		path = (
			f"/me/events?$select={EVENT_SELECT}&$top=50"
			f"&$orderby=lastModifiedDateTime desc"
			f"&$filter=lastModifiedDateTime ge {since}"
		)
		resp = graph.graph_request("GET", path, doc.name)
		events = resp.get("value", [])
	else:
		# first sync: a recent window via calendarView (expands recurrences, returns UTC)
		start = _iso_utc(add_to_date(now_datetime(), days=-30))
		end = _iso_utc(add_to_date(now_datetime(), days=60))
		path = (
			f"/me/calendarView?startDateTime={start}&endDateTime={end}"
			f"&$select={EVENT_SELECT}&$top=50&$orderby=start/dateTime"
		)
		resp = graph.graph_request(
			"GET", path, doc.name, headers={"Prefer": 'outlook.timezone="UTC"'}
		)
		events = resp.get("value", [])

	frappe.flags.in_microsoft_sync = True
	count = 0
	try:
		for ev in events:
			try:
				_upsert_event(doc, ev)
				count += 1
			except Exception:
				frappe.log_error(title=f"MS event upsert failed: {ev.get('id')}")
	finally:
		frappe.flags.in_microsoft_sync = False
	return count


def _upsert_event(doc, ev):
	ms_id = ev.get("id")
	if not ms_id:
		return

	existing = frappe.db.get_value(
		"Event", {"custom_microsoft_event_id": ms_id}, "name"
	)

	# cancelled in Microsoft -> remove the Frappe mirror
	if ev.get("isCancelled"):
		if existing:
			frappe.delete_doc("Event", existing, ignore_permissions=True, force=True)
		return

	starts_on = _ms_dt_to_system(ev.get("start"))
	ends_on = _ms_dt_to_system(ev.get("end"))
	subject = ev.get("subject") or "(No subject)"
	description = ev.get("bodyPreview") or ""

	if existing:
		event = frappe.get_doc("Event", existing)
	else:
		event = frappe.new_doc("Event")
		event.custom_microsoft_event_id = ms_id

	event.subject = subject
	if starts_on:
		event.starts_on = starts_on
	if ends_on:
		event.ends_on = ends_on
	event.all_day = 1 if ev.get("isAllDay") else 0
	event.description = description
	event.custom_microsoft_calendar = doc.name
	event.custom_sync_with_microsoft_calendar = 1
	event.custom_pulled_from_microsoft = 1
	event.event_type = event.event_type or "Private"
	event.flags.ignore_permissions = True
	event.flags.ignore_mandatory = True
	event.save()


# --- push (Frappe -> Graph) ----------------------------------------------------------

def _push(doc):
	"""Create unsynced Frappe events in Microsoft. Returns count pushed."""
	candidates = frappe.get_all(
		"Event",
		filters={
			"custom_sync_with_microsoft_calendar": 1,
			"custom_microsoft_calendar": doc.name,
			"custom_pulled_from_microsoft": 0,
			"custom_microsoft_event_id": ["in", ["", None]],
		},
		pluck="name",
	)
	count = 0
	for name in candidates:
		try:
			event = frappe.get_doc("Event", name)
			created = _create_graph_event(doc.name, event)
			if created.get("id"):
				frappe.db.set_value(
					"Event", name, "custom_microsoft_event_id", created["id"],
					update_modified=False,
				)
				count += 1
		except Exception:
			frappe.log_error(title=f"MS event push failed: {name}")
	frappe.db.commit()
	return count


def _event_to_graph_body(event):
	body = {
		"subject": event.subject or "(No subject)",
		"body": {"contentType": "HTML", "content": event.description or ""},
		"start": _system_dt_to_ms(event.starts_on),
		"end": _system_dt_to_ms(event.ends_on or add_to_date(get_datetime(event.starts_on), minutes=30)),
	}
	if getattr(event, "all_day", 0):
		body["isAllDay"] = True
	return body


def _create_graph_event(calendar_name, event):
	return graph.graph_request(
		"POST", "/me/events", calendar_name, json=_event_to_graph_body(event)
	)


# --- doc_events (single Frappe Event lifecycle -> Graph) -----------------------------

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
			graph.graph_request(
				"PATCH",
				f"/me/events/{doc.custom_microsoft_event_id}",
				cal.name,
				json=_event_to_graph_body(doc),
			)
		else:
			created = _create_graph_event(cal.name, doc)
			if created.get("id"):
				frappe.db.set_value(
					"Event", doc.name, "custom_microsoft_event_id", created["id"],
					update_modified=False,
				)
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
def fetch_events(calendar_name, start_datetime=None, end_datetime=None):
	"""Return upcoming events from Graph (calendarView) as a clean list. Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)

	start = _iso_utc(start_datetime) if start_datetime else _iso_utc(now_datetime())
	end = _iso_utc(end_datetime) if end_datetime else _iso_utc(add_to_date(now_datetime(), days=30))

	path = (
		f"/me/calendarView?startDateTime={start}&endDateTime={end}"
		f"&$select={EVENT_SELECT}&$top=100&$orderby=start/dateTime"
	)
	resp = graph.graph_request(
		"GET", path, calendar_name, headers={"Prefer": 'outlook.timezone="UTC"'}
	)

	out = []
	for ev in resp.get("value", []):
		online = ev.get("onlineMeeting") or {}
		out.append(
			{
				"id": ev.get("id"),
				"subject": ev.get("subject"),
				"start": _ms_dt_to_system(ev.get("start")),
				"end": _ms_dt_to_system(ev.get("end")),
				"is_all_day": bool(ev.get("isAllDay")),
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
