"""Teams online-meeting creation (stage M3).

Two paths:
  * create_calendar_event=True  -> POST /me/events with isOnlineMeeting (PREFERRED).
    This creates a Teams meeting AND associates it with a calendar event, which is the
    precondition for transcripts (see M4 / docs/graph-api-reference.md).
  * create_calendar_event=False -> POST /me/onlineMeetings (standalone). NOT calendar-
    associated, so transcripts will NOT be available for meetings created this way.

All entrypoints are whitelisted and owner-checked. Times are ISO 8601; naive datetimes are
interpreted in the system timezone.
"""

import frappe
from frappe import _
from frappe.utils import add_to_date, get_datetime

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_calendar_sync import (
	_check_owner,
	_iso_utc,
	_system_dt_to_ms,
)


def _attendees_array(attendees):
	out = []
	for a in attendees or []:
		if not a:
			continue
		out.append({"emailAddress": {"address": a}, "type": "required"})
	return out


@frappe.whitelist()
def create_meeting(
	calendar_name: str,
	subject: str,
	start_datetime: str,
	end_datetime: str | None = None,
	attendees: str | list | None = None,
	body: str | None = None,
	create_calendar_event: int = 1,
):
	"""Create a Teams online meeting. Owner-checked.

	Returns (calendar path): {event_id, web_link, join_url, online_meeting_id}
	Returns (standalone path): {online_meeting_id, join_url, calendar_associated: False}
	"""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)

	if isinstance(attendees, str):
		attendees = [a.strip() for a in attendees.replace(",", "\n").split("\n") if a.strip()]
	create_calendar_event = frappe.utils.cint(create_calendar_event)

	if not end_datetime:
		end_datetime = add_to_date(get_datetime(start_datetime), minutes=30)

	if create_calendar_event:
		payload = {
			"subject": subject,
			"body": {"contentType": "HTML", "content": body or ""},
			"start": _system_dt_to_ms(start_datetime),
			"end": _system_dt_to_ms(end_datetime),
			"isOnlineMeeting": True,
			"onlineMeetingProvider": "teamsForBusiness",
		}
		att = _attendees_array(attendees)
		if att:
			payload["attendees"] = att

		event = graph.graph_request("POST", "/me/events", calendar_name, json=payload)
		online = event.get("onlineMeeting") or {}
		join_url = online.get("joinUrl")

		online_meeting_id = None
		if join_url:
			# best-effort; may be empty immediately after creation, that's ok
			try:
				online_meeting_id = _resolve_online_meeting_id(calendar_name, join_url)
			except Exception:
				online_meeting_id = None

		return {
			"event_id": event.get("id"),
			"web_link": event.get("webLink"),
			"join_url": join_url,
			"online_meeting_id": online_meeting_id,
			"calendar_associated": True,
		}

	# standalone Teams meeting — NOT calendar-associated; transcripts won't be available
	payload = {
		"startDateTime": _iso_utc(start_datetime),
		"endDateTime": _iso_utc(end_datetime),
		"subject": subject,
	}
	meeting = graph.graph_request("POST", "/me/onlineMeetings", calendar_name, json=payload)
	return {
		"online_meeting_id": meeting.get("id"),
		"join_url": meeting.get("joinWebUrl"),
		"calendar_associated": False,
		"note": "Standalone meeting is not calendar-associated; transcripts are unavailable.",
	}


def _resolve_online_meeting_id(calendar_name, join_url):
	"""Map a Teams joinWebUrl to its onlineMeeting id (needed for transcripts)."""
	# OData string literal: escape single quotes by doubling them
	safe = (join_url or "").replace("'", "''")
	resp = graph.graph_request(
		"GET",
		f"/me/onlineMeetings?$filter=JoinWebUrl eq '{safe}'",
		calendar_name,
	)
	value = resp.get("value") or []
	return value[0]["id"] if value else None


@frappe.whitelist()
def resolve_online_meeting_id(calendar_name: str, join_url: str):
	"""Whitelisted wrapper around the join-URL -> onlineMeeting id mapping. Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	return {"online_meeting_id": _resolve_online_meeting_id(calendar_name, join_url)}
