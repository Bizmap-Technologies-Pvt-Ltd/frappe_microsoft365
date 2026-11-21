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

import re

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


#: Characters that mean something to a query string itself. Written raw into the $filter they
#: never reach Graph as part of the join URL: ``&`` starts another parameter, ``+`` is read as a
#: space and ``#`` ends the URL. Percent-encoded, they arrive as the characters they are.
#:
#: Deliberately only these three. Everything else is left for graph_url to encode the way
#: requests does, which is what old meetup-join links (full of %3a and %40) have been sent as
#: all along, and Graph matches them that way.
_QUERY_SPECIALS = (("&", "%26"), ("+", "%2B"), ("#", "%23"))

#: The numeric meeting id in a new-style join link: https://teams.microsoft.com/meet/<id>?p=...
_MEET_ID = re.compile(r"/meet/(\d+)")


def _filter_literal(value):
	"""``value`` as the inside of an OData string literal in a $filter written into the path."""
	# OData string literal: escape single quotes by doubling them.
	value = (value or "").replace("'", "''")
	for char, encoded in _QUERY_SPECIALS:
		value = value.replace(char, encoded)
	return value


def _first_meeting(calendar_name, condition):
	resp = graph.graph_request("GET", f"/me/onlineMeetings?$filter={condition}", calendar_name)
	value = resp.get("value") or []
	return value[0]["id"] if value else None


def _outlook_join_url(calendar_name, event_id):
	"""The join link Outlook itself holds for this event, or None.

	The link stored on a Frappe Event can be a different spelling of the same meeting — copied
	from an invitation, or written before Teams switched to /meet/ links — and Graph matches the
	string, not the meeting. The event's own onlineMeeting is the one Graph issued.

	A 404 or 403 here is not the lookup failing, only this second opinion being unavailable, so
	it answers None rather than replacing the real question's answer with its own error.
	"""
	try:
		event = graph.graph_request(
			"GET", f"/me/events/{event_id}?$select=onlineMeeting", calendar_name
		)
	except graph.MsGraphNotFound:
		return None
	except graph.MsGraphError as e:
		if "403" in str(e) or "Forbidden" in str(e):
			return None
		raise
	return (event.get("onlineMeeting") or {}).get("joinUrl")


def _resolve_online_meeting_id(calendar_name, join_url, event_id=None):
	"""Map a Teams joinWebUrl to its onlineMeeting id (needed for transcripts).

	Asked by join link first, as it always was. Only when that matches nothing does it try the
	Outlook event's own join link (``event_id`` is its Graph id), and then the numeric meeting
	id a /meet/ link carries — so a meeting that resolved before costs exactly the one call it
	always did. Errors are raised, never turned into "no match": a 503 is not a foreign meeting.
	"""
	candidates = [join_url] if join_url else []
	meeting_id = _first_meeting(calendar_name, f"JoinWebUrl eq '{_filter_literal(join_url)}'")
	if meeting_id:
		return meeting_id

	if event_id:
		outlook_url = _outlook_join_url(calendar_name, event_id)
		if outlook_url and outlook_url != join_url:
			candidates.append(outlook_url)
			meeting_id = _first_meeting(
				calendar_name, f"JoinWebUrl eq '{_filter_literal(outlook_url)}'"
			)
			if meeting_id:
				return meeting_id

	for url in candidates:
		match = _MEET_ID.search(url)
		if match:
			meeting_id = _first_meeting(
				calendar_name, f"joinMeetingIdSettings/joinMeetingId eq '{match.group(1)}'"
			)
			if meeting_id:
				return meeting_id
	return None


@frappe.whitelist()
def resolve_online_meeting_id(calendar_name: str, join_url: str):
	"""Whitelisted wrapper around the join-URL -> onlineMeeting id mapping. Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	return {"online_meeting_id": _resolve_online_meeting_id(calendar_name, join_url)}
