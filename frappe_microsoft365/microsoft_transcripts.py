"""Teams meeting transcripts and recordings (stage M4).

Transcripts/recordings are only available for meetings that are **associated with a calendar
event** (created via POST /me/events with isOnlineMeeting=true) and that have **not expired**.

Delegated permissions required (with tenant-admin consent):
  * transcripts -> OnlineMeetingTranscript.Read.All
  * recordings  -> OnlineMeetingRecording.Read.All  (also subject to Teams Premium / licensing)

All entrypoints are whitelisted and owner-checked. Graph errors surface as clean messages;
a 403 typically means the scope/consent is missing or the meeting is not calendar-associated.
"""

import frappe
from frappe import _

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError
from frappe_microsoft365.microsoft_calendar_sync import _check_owner
from frappe_microsoft365.microsoft_meetings import _resolve_online_meeting_id

_PERM_HINT = (
	"Transcripts require the OnlineMeetingTranscript.Read.All delegated permission with "
	"tenant-admin consent, and the meeting must be calendar-associated and not expired."
)
_REC_PERM_HINT = (
	"Recordings require the OnlineMeetingRecording.Read.All delegated permission with "
	"tenant-admin consent (and may require Teams Premium licensing)."
)


def _wrap_403(e, hint):
	msg = str(e)
	if "403" in msg or "Forbidden" in msg or "Authorization" in msg:
		frappe.throw(f"{msg}\n{hint}", MsGraphError)
	raise e


# --- transcripts ---------------------------------------------------------------------

@frappe.whitelist()
def list_transcripts(calendar_name, online_meeting_id):
	"""List transcripts for an online meeting. Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET", f"/me/onlineMeetings/{online_meeting_id}/transcripts", calendar_name
		)
	except MsGraphError as e:
		_wrap_403(e, _PERM_HINT)
	return [
		{
			"id": t.get("id"),
			"created_date_time": t.get("createdDateTime"),
			"transcript_content_url": t.get("transcriptContentUrl"),
		}
		for t in resp.get("value", [])
	]


@frappe.whitelist()
def get_transcript_content(calendar_name, online_meeting_id, transcript_id, fmt="text/vtt"):
	"""Return the raw transcript content (VTT by default). Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET",
			f"/me/onlineMeetings/{online_meeting_id}/transcripts/{transcript_id}/content",
			calendar_name,
			params={"$format": fmt},
			raw=True,
		)
	except MsGraphError as e:
		_wrap_403(e, _PERM_HINT)
	return {"format": fmt, "content": resp.text}


@frappe.whitelist()
def get_transcripts_for_join_url(calendar_name, join_url):
	"""Resolve a join URL -> online meeting, then list transcripts + fetch the latest VTT.

	Returns {online_meeting_id, transcripts:[...], latest_vtt}. Owner-checked.
	"""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)

	online_meeting_id = _resolve_online_meeting_id(calendar_name, join_url)
	if not online_meeting_id:
		return {"online_meeting_id": None, "transcripts": [], "latest_vtt": None}

	transcripts = list_transcripts(calendar_name, online_meeting_id)
	latest_vtt = None
	if transcripts:
		latest = sorted(
			transcripts, key=lambda t: t.get("created_date_time") or "", reverse=True
		)[0]
		try:
			latest_vtt = get_transcript_content(
				calendar_name, online_meeting_id, latest["id"]
			).get("content")
		except Exception:
			latest_vtt = None

	return {
		"online_meeting_id": online_meeting_id,
		"transcripts": transcripts,
		"latest_vtt": latest_vtt,
	}


# --- recordings ----------------------------------------------------------------------

@frappe.whitelist()
def list_recordings(calendar_name, online_meeting_id):
	"""List recordings for an online meeting. Owner-checked. See licensing note."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET", f"/me/onlineMeetings/{online_meeting_id}/recordings", calendar_name
		)
	except MsGraphError as e:
		_wrap_403(e, _REC_PERM_HINT)
	return [
		{
			"id": r.get("id"),
			"created_date_time": r.get("createdDateTime"),
			"recording_content_url": r.get("recordingContentUrl"),
		}
		for r in resp.get("value", [])
	]


@frappe.whitelist()
def get_recording_content(calendar_name, online_meeting_id, recording_id):
	"""Return the raw recording bytes' download via Graph. Owner-checked. See licensing note."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET",
			f"/me/onlineMeetings/{online_meeting_id}/recordings/{recording_id}/content",
			calendar_name,
			raw=True,
		)
	except MsGraphError as e:
		_wrap_403(e, _REC_PERM_HINT)
	return {
		"content_type": resp.headers.get("Content-Type"),
		"content_length": resp.headers.get("Content-Length"),
	}
