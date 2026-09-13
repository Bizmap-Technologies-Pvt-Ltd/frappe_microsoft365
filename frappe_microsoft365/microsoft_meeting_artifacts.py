"""Bring a finished Teams meeting's transcript and recording back to its Frappe Event.

The two are treated differently on purpose:

* **The transcript is attached.** A VTT is kilobytes, it is the part people actually search
  and quote, and once it is a File on the Event it survives whatever Microsoft later does with
  its own copy.
* **The recording is not.** A Teams recording is routinely hundreds of megabytes, and copying
  one into the site's file store per meeting is a bad trade. Microsoft keeps it; we record that
  it exists and fetch it on demand.

Why the recording cannot simply be a link: ``recordingContentUrl`` is a Graph endpoint that
requires a bearer token, so a browser given that URL gets 401, not a video. The download below
streams it through Frappe instead, authorised by Frappe's own permissions.

Graph also publishes no expiry for a recording (the callRecording resource has no such
property), so rather than invent a date, a fetch that comes back missing is reported as what it
is: gone, because recordings expire under the tenant's retention policy.

Requires the Teams and transcript capabilities: a transcript is addressed by online meeting id,
and resolving that from a join URL needs OnlineMeetings.ReadWrite.
"""

import frappe
from frappe import _
from frappe.utils import get_datetime, now_datetime

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_calendar_sync import _check_owner
from frappe_microsoft365.microsoft_graph import MsGraphError
from frappe_microsoft365.microsoft_meetings import _resolve_online_meeting_id

#: Attachments are matched by this prefix so a re-fetch replaces rather than accumulates.
TRANSCRIPT_PREFIX = "teams-transcript-"

def gone_hint():
	"""Built per call: _() resolves against the current site and language, so a module-level
	constant would freeze whichever site imported this first."""
	return _(
		"Microsoft no longer has this. Teams recordings and transcripts are deleted under your "
		"tenant's retention policy, and Microsoft does not publish the expiry date through the API."
	)


def _event_for_artifacts(event_name):
	"""The Event, its calendar and its online meeting id, with every precondition checked."""
	doc = frappe.get_doc("Event", event_name)
	doc.check_permission("read")

	calendar_name = doc.get("custom_microsoft_calendar")
	join_url = doc.get("custom_teams_join_url")
	if not calendar_name or not join_url:
		frappe.throw(_("This event has no Teams meeting."))

	calendar = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(calendar)
	if not calendar.enabled or not calendar.authorized:
		frappe.throw(_("The Microsoft Calendar for this event is disabled or not authorized."))

	# Nothing exists until the meeting has actually happened, and saying so is kinder than a
	# Graph call that returns an empty list for a reason the person cannot see.
	if doc.ends_on and get_datetime(doc.ends_on) > now_datetime():
		frappe.throw(_("This meeting has not finished yet."))

	meeting_id = doc.get("custom_microsoft_online_meeting_id")
	if not meeting_id:
		meeting_id = _resolve_online_meeting_id(calendar_name, join_url)
		if not meeting_id:
			frappe.throw(
				_("Microsoft could not match this join link to a meeting. {0}").format(gone_hint())
			)
		frappe.db.set_value(
			"Event", doc.name, "custom_microsoft_online_meeting_id", meeting_id, update_modified=False
		)

	return doc, calendar_name, meeting_id


def _describe_recordings(recordings):
	"""One readable line per recording. Graph gives no size and no expiry, so neither is shown."""
	lines = []
	for recording in recordings:
		created = recording.get("created_date_time") or ""
		when = frappe.utils.format_datetime(created) if created else _("unknown time")
		lines.append(_("Recorded {0}").format(when))
	return "\n".join(lines)


@frappe.whitelist(methods=["POST"])
def fetch_meeting_artifacts(event: str):
	"""Attach the transcript and note any recordings. POST: it writes a File and fields."""
	from frappe_microsoft365 import microsoft_transcripts as ms

	doc, calendar_name, meeting_id = _event_for_artifacts(event)
	attached = None
	recordings = []

	try:
		transcripts = ms.list_transcripts(calendar_name, meeting_id)
	except MsGraphError as e:
		frappe.throw(_("Could not read transcripts: {0}").format(str(e)))

	if transcripts:
		latest = sorted(transcripts, key=lambda t: t.get("created_date_time") or "", reverse=True)[0]
		content = ms.get_transcript_content(calendar_name, meeting_id, latest["id"]).get("content")
		if content:
			attached = _attach_transcript(doc, content, latest.get("created_date_time"))

	try:
		recordings = ms.list_recordings(calendar_name, meeting_id)
	except MsGraphError:
		# Recordings need their own permission and licence; a transcript should still land.
		recordings = []

	frappe.db.set_value(
		"Event",
		doc.name,
		{
			"custom_microsoft_recordings": _describe_recordings(recordings),
			"custom_microsoft_transcript_fetched_on": now_datetime() if attached else None,
		},
		update_modified=False,
	)

	return {
		"transcript": attached,
		"recordings": len(recordings),
		"message": _summary(attached, len(recordings)),
	}


def _summary(attached, recording_count):
	if attached and recording_count:
		return _("Transcript attached. {0} recording(s) available to download.").format(recording_count)
	if attached:
		return _("Transcript attached. No recording was found.")
	if recording_count:
		return _("No transcript was found. {0} recording(s) available to download.").format(recording_count)
	return _("Microsoft has no transcript or recording for this meeting. {0}").format(gone_hint())


def _attach_transcript(doc, content, created_at):
	"""Save the VTT as a File on the Event, replacing an earlier copy of the same meeting."""
	from frappe.utils.file_manager import save_file

	stamp = (created_at or "")[:10] or frappe.utils.nowdate()
	filename = f"{TRANSCRIPT_PREFIX}{stamp}.vtt"

	# Matched by prefix, not by exact name: Frappe appends a content hash when a file of that
	# name already exists, so re-fetching would otherwise pile up near-identical copies.
	for existing in frappe.get_all(
		"File",
		filters={
			"attached_to_doctype": "Event",
			"attached_to_name": doc.name,
			"file_name": ["like", f"{TRANSCRIPT_PREFIX}%"],
		},
		pluck="name",
	):
		frappe.delete_doc("File", existing, force=True, ignore_permissions=True)

	saved = save_file(filename, content.encode("utf-8"), "Event", doc.name, is_private=1)
	return {"file_name": saved.file_name, "file_url": saved.file_url}


@frappe.whitelist()
def download_recording(event: str, recording_id: str | None = None):
	"""Stream a recording from Microsoft through Frappe.

	Not stored: the bytes pass through and are handed to the browser. Graph's own URL cannot
	be given to a browser because it needs a bearer token.
	"""
	doc, calendar_name, meeting_id = _event_for_artifacts(event)
	from frappe_microsoft365 import microsoft_transcripts as ms

	if not recording_id:
		recordings = ms.list_recordings(calendar_name, meeting_id)
		if not recordings:
			frappe.throw(_("Microsoft has no recording for this meeting. {0}").format(gone_hint()))
		recording_id = sorted(
			recordings, key=lambda r: r.get("created_date_time") or "", reverse=True
		)[0]["id"]

	try:
		resp = graph.graph_request(
			"GET",
			f"/me/onlineMeetings/{meeting_id}/recordings/{recording_id}/content",
			calendar_name,
			raw=True,
		)
	except MsGraphError as e:
		frappe.throw(_("Could not download the recording: {0}. {1}").format(str(e), gone_hint()))

	frappe.local.response.filename = f"teams-recording-{doc.name}.mp4"
	frappe.local.response.filecontent = resp.content
	frappe.local.response.type = "download"
