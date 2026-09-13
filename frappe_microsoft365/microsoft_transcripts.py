"""Teams meeting transcripts and recordings (stage M4).

Transcripts/recordings are only available for meetings that are **associated with a calendar
event** (created via POST /me/events with isOnlineMeeting=true) and that have **not expired**.

Delegated permissions required (with tenant-admin consent):
  * transcripts -> OnlineMeetingTranscript.Read.All
  * recordings  -> OnlineMeetingRecording.Read.All

Whether a meeting COULD be recorded at all is a separate tenant policy (Teams admin center,
Meetings > Meeting policies > Meeting recording), and it is not what a 403 here is about: a
tenant with recording switched off simply has no recordings, and Graph answers 200 with an
empty list. Naming a licence in a 403 message would send people to the wrong page.

All entrypoints are whitelisted and owner-checked. Graph errors surface as clean messages.

Three 403s look identical and are not
-------------------------------------
Every one of them says "Forbidden", and the fix for each is in a different place. Answering
all three with "grant transcript consent" is what sent a real tenant round a loop:

* **transcripts** -> the tenant switch below, far more often than consent.
* **recordings**  -> OnlineMeetingRecording.Read.All, which Microsoft consents to separately
  from the transcript permission. A tenant happy to let an app read words frequently refuses
  to let it read video, so this one is missing on its own more often than not.
* **the join-link lookup** -> OnlineMeetings.ReadWrite, which neither of the above includes.

So each has its own hint, and none of them mentions the others.
"""

import frappe
from frappe import _

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_calendar_sync import _check_owner
from frappe_microsoft365.microsoft_graph import MsGraphError
from frappe_microsoft365.microsoft_meetings import _resolve_online_meeting_id


def _transcript_perm_hint():
	"""Built per call: _() resolves against the current site and language, so a module-level
	constant would freeze whichever site imported this first."""
	return _(
		"Transcripts need OnlineMeetingTranscript.Read.All: check it is listed and consented "
		"under Entra ID > App registrations > your app > API permissions, then Re-authorize this "
		"Microsoft Calendar. A meeting created without a calendar event never has one at all."
	)


def _recording_perm_hint():
	return _(
		"Recordings need their own OnlineMeetingRecording.Read.All, which tenants often consent "
		"to separately from the transcript one: add it under Entra ID > App registrations > your "
		"app > API permissions, grant admin consent, then Re-authorize this Microsoft Calendar."
	)


def _online_meetings_hint():
	return _(
		"Finding a meeting from its join link needs OnlineMeetings.ReadWrite, which the transcript "
		"permission does not include: tick Standalone Teams meetings in Microsoft Settings, consent "
		"to it in Entra ID > App registrations, then Re-authorize this Microsoft Calendar."
	)


#: Microsoft added a tenant switch for this in 2026 and shipped it OFF. Every tenant now has
#: to turn it on explicitly, so a 403 here is far more often this than a missing permission —
#: and telling someone to grant consent they already granted sends them in a circle.
TENANT_SWITCH = "GraphAccessToTranscriptsDisabled"
TENANT_SWITCH_TEXT = "access to transcripts is disabled"


def _tenant_switch_hint():
	return _(
		"This is a tenant setting, not a permission: in the Teams admin center, go to "
		"Meetings > Meeting settings > Transcript API access and turn Microsoft Graph access "
		"On. Microsoft ships it off, so it has to be turned on once per tenant even when every "
		"permission is already consented."
	)


def _wrap_403(e, hint):
	msg = str(e)
	if TENANT_SWITCH in msg or TENANT_SWITCH_TEXT in msg:
		# Named cause, so say the named cause. The generic permission hint below is actively
		# misleading here — it asks for consent that is already granted.
		frappe.throw(_tenant_switch_hint(), MsGraphError, title=_("Transcripts are switched off for this tenant"))
	if "403" in msg or "Forbidden" in msg or "Authorization" in msg:
		frappe.throw(f"{msg}\n{hint}", MsGraphError)
	raise e


def _wrap_join_url_403(e):
	"""Explain a refusal to map a join link to a meeting id; anything else is re-raised as it is.

	Kept apart from _wrap_403 because the permission is a different one: this call is
	/me/onlineMeetings, not /transcripts, and it fails for tenants that consented to every
	transcript permission there is. Handing that person the transcript hint tells them to grant
	what the portal in front of them already shows as granted.
	"""
	msg = str(e)
	if "403" in msg or "Forbidden" in msg or "Authorization" in msg:
		frappe.throw(
			f"{msg}\n{_online_meetings_hint()}",
			MsGraphError,
			title=_("Microsoft would not look this meeting up"),
		)
	raise e


# --- transcripts ---------------------------------------------------------------------

@frappe.whitelist()
def list_transcripts(calendar_name: str, online_meeting_id: str):
	"""List transcripts for an online meeting. Owner-checked."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET", f"/me/onlineMeetings/{online_meeting_id}/transcripts", calendar_name
		)
	except MsGraphError as e:
		_wrap_403(e, _transcript_perm_hint())
	return [
		{
			"id": t.get("id"),
			"created_date_time": t.get("createdDateTime"),
			"transcript_content_url": t.get("transcriptContentUrl"),
		}
		for t in resp.get("value", [])
	]


@frappe.whitelist()
def get_transcript_content(calendar_name: str, online_meeting_id: str, transcript_id: str, fmt: str = "text/vtt"):
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
		_wrap_403(e, _transcript_perm_hint())
	return {"format": fmt, "content": resp.text}


@frappe.whitelist()
def get_transcripts_for_join_url(calendar_name: str, join_url: str):
	"""Resolve a join URL -> online meeting, then list transcripts + fetch the latest VTT.

	Returns {online_meeting_id, transcripts:[...], latest_vtt}. Owner-checked.
	"""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)

	try:
		online_meeting_id = _resolve_online_meeting_id(calendar_name, join_url)
	except MsGraphError as e:
		# Refused, not empty. Left unwrapped this surfaced as a bare Graph 403 naming the
		# /me/onlineMeetings path, which reads like the transcript permission failing.
		_wrap_join_url_403(e)
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
def list_recordings(calendar_name: str, online_meeting_id: str):
	"""List recordings for an online meeting. Owner-checked. Empty is not an error: see above."""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET", f"/me/onlineMeetings/{online_meeting_id}/recordings", calendar_name
		)
	except MsGraphError as e:
		_wrap_403(e, _recording_perm_hint())
	return [
		{
			"id": r.get("id"),
			"created_date_time": r.get("createdDateTime"),
			# A meeting past four hours comes back as several recordings; without the end time
			# there is no way to tell a caller which part covers which stretch of the meeting.
			"end_date_time": r.get("endDateTime"),
			"recording_content_url": r.get("recordingContentUrl"),
		}
		for r in resp.get("value", [])
	]


@frappe.whitelist()
def get_recording_content(calendar_name: str, online_meeting_id: str, recording_id: str):
	"""Describe a recording's content — type and size — without downloading it.

	``stream=True`` is the whole point of this function. Without it ``requests`` reads the
	entire body before returning, so asking "how big is this?" about a Teams recording pulled
	up to 1.5 GB into the worker and threw it away to read two headers. The headers arrive with
	the response; the body is closed unread.

	To actually fetch the bytes, use ``microsoft_meeting_artifacts.download_recording``, which
	streams them through to the browser a chunk at a time.
	"""
	doc = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(doc)
	try:
		resp = graph.graph_request(
			"GET",
			f"/me/onlineMeetings/{online_meeting_id}/recordings/{recording_id}/content",
			calendar_name,
			raw=True,
			stream=True,
		)
	except MsGraphError as e:
		_wrap_403(e, _recording_perm_hint())
	try:
		return {
			"content_type": resp.headers.get("Content-Type"),
			"content_length": resp.headers.get("Content-Length"),
		}
	finally:
		# An unread streamed body holds its pooled connection open until the process ends.
		resp.close()
