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
proxies it through Frappe instead, authorised by Frappe's own permissions — chunk by chunk and
never whole. A single recording runs to 1.5 GB (see below), so holding one in a web worker
would cost that much resident memory per concurrent download and make the browser wait for the
last byte before it sees the first, which is past the default 120-second worker timeout.

Nothing is ready the moment a meeting ends
------------------------------------------
Microsoft publishes no SLA for this. Reported reality: five to thirty minutes for an ordinary
meeting, a few hours for a long or heavy one, and Graph lags the Teams UI — a transcript can be
readable in Teams while this endpoint still returns an empty list. So a fetch that comes back
empty is *not* evidence that nothing exists, and saying "there is no recording" at minute two
would be a lie. ``fetch_pending`` keeps asking on a widening backoff for a day, and every
message distinguishes "not ready yet" from "Microsoft has none" from "too old now".

Two different clocks end it
---------------------------
* **The meeting expires.** Graph's list-recordings endpoint "works only for a meeting that
  hasn't expired": 60 days after a one-off meeting's scheduled time, extended by 60 more
  whenever someone joins or edits it. After that these endpoints return nothing, whatever
  still exists in OneDrive.
* **The file expires.** Teams deletes recordings and transcripts on the tenant's retention
  policy — 120 days by default, admin-configurable from one day to never.

Long meetings come back in pieces
---------------------------------
A Teams recording stops and restarts at **4 hours or 1.5 GB**, whichever comes first, so a
five-hour meeting is two recordings, not one. Transcription can be stopped and restarted too.
Everything here is therefore plural: every transcript is attached, every recording is listed
and separately downloadable.

Requires the Teams and transcript capabilities: a transcript is addressed by online meeting id,
and resolving that from a join URL needs OnlineMeetings.ReadWrite.
"""

import json

import frappe
from frappe import _
from frappe.utils import add_to_date, cint, get_datetime, now_datetime
from werkzeug.wrappers import Response

from frappe_microsoft365 import background
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_calendar_sync import _check_owner
from frappe_microsoft365.microsoft_graph import MsGraphError
from frappe_microsoft365.microsoft_meetings import _resolve_online_meeting_id

#: Attachments are matched by this prefix so a re-fetch replaces rather than accumulates.
TRANSCRIPT_PREFIX = "teams-transcript-"

#: Minutes after the meeting's end at which the catch-up job asks Microsoft again.
#:
#: Widening rather than fixed, because the distribution is lopsided: most meetings are ready
#: inside half an hour and the tail is hours long. Fourteen attempts spread over a day cost
#: fourteen Graph calls; polling every fifteen minutes across the same day would cost ninety-six
#: and find it no sooner.
RETRY_MINUTES = (10, 25, 45, 75, 120, 180, 270, 360, 480, 600, 720, 960, 1200, 1440)

#: How long Graph keeps serving a meeting's artifacts, counted from the meeting itself. A join
#: or an edit pushes it out by another 60 days, so this is the earliest it can go, not a
#: promise that it has.
MEETING_EXPIRY_DAYS = 60

#: Bytes read from Graph, and written to the browser, per step of a recording download.
#:
#: The number only has to be big enough that per-chunk overhead disappears and small enough that
#: a worker's memory stays flat: at 256 KB a worst-case 1.5 GB part is ~6,000 reads holding a
#: quarter of a megabyte at a time, where 8 KB would be ~200,000 reads for no benefit and 16 MB
#: would put a visible sawtooth back into the worker's resident size.
RECORDING_CHUNK_BYTES = 256 * 1024


def _processing_note():
	"""Built per call: _() resolves against the current site and language, so a module-level
	constant would freeze whichever site imported this first."""
	return _(
		"Microsoft publishes no schedule for this. Most meetings are ready in 5-30 minutes; a "
		"long or heavy one can take a few hours."
	)


def _expiry_note():
	return _(
		"Microsoft stops serving a meeting's transcript and recording about {0} days after it "
		"happens, and deletes the files themselves on your tenant's retention policy."
	).format(MEETING_EXPIRY_DAYS)


def _event_for_artifacts(event_name, resolve=True):
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
	if not meeting_id and resolve:
		meeting_id = _resolve_online_meeting_id(calendar_name, join_url)
		if not meeting_id:
			frappe.throw(
				_("Microsoft could not match this join link to a meeting. {0}").format(_expiry_note())
			)
		frappe.db.set_value(
			"Event", doc.name, "custom_microsoft_online_meeting_id", meeting_id, update_modified=False
		)

	return doc, calendar_name, meeting_id


# --- state ---------------------------------------------------------------------------
#
# One function decides what is true about an Event's artifacts; the messages, the status field
# and the form buttons all read it. Working it out twice is how a form ends up offering
# "Get Transcript" for a meeting whose transcript is already attached.


@frappe.whitelist()
def artifact_state(event: str):
	"""What stage this Event's artifacts are at, for the form. Reads the Event, not Graph."""
	doc = frappe.get_doc("Event", event)
	doc.check_permission("read")
	state = _state(doc)
	return {**state, "message": describe_state(state)}


def _state(doc):
	ends_on = get_datetime(doc.ends_on) if doc.ends_on else None
	now = now_datetime()
	attempts = cint(doc.get("custom_microsoft_artifacts_attempts"))
	has_transcript = bool(doc.get("custom_microsoft_transcript_fetched_on"))
	recordings = _stored_recordings(doc)

	state = {
		"has_transcript": has_transcript,
		"recordings": len(recordings),
		"attempts": attempts,
		"expires_on": add_to_date(ends_on, days=MEETING_EXPIRY_DAYS) if ends_on else None,
		"next_check": None,
		"auto": _auto_enabled(doc),
	}

	if not doc.get("custom_teams_join_url"):
		return {**state, "state": "no_meeting"}
	if ends_on and ends_on > now:
		return {**state, "state": "not_finished"}

	if has_transcript and recordings:
		return {**state, "state": "complete"}

	# Past the meeting's own expiry there is nothing left to wait for, so stop promising.
	if state["expires_on"] and state["expires_on"] < now:
		return {**state, "state": "partial" if (has_transcript or recordings) else "expired"}

	# Still worth asking again? Only while the backoff has attempts left.
	state["next_check"] = _next_check_at(ends_on, attempts)
	if has_transcript or recordings:
		return {**state, "state": "partial" if state["next_check"] else "complete"}
	return {**state, "state": "processing" if state["next_check"] else "nothing_found"}


def _auto_enabled(doc):
	"""Is anything going to check on its own? Promising a next check when nothing is scheduled
	is the same class of lie as reporting an empty answer as 'never recorded'."""
	calendar = doc.get("custom_microsoft_calendar")
	if not calendar:
		return False
	return bool(
		frappe.get_cached_value("Microsoft Calendar", calendar, "fetch_artifacts_automatically")
	)


def _next_check_at(ends_on, attempts):
	"""When the catch-up job will try next, or None once the backoff is exhausted."""
	if not ends_on or attempts >= len(RETRY_MINUTES):
		return None
	return add_to_date(ends_on, minutes=RETRY_MINUTES[attempts])


def _stored_recordings(doc):
	raw = doc.get("custom_microsoft_recordings_data")
	if not raw:
		return []
	try:
		value = json.loads(raw)
	except (ValueError, TypeError):
		return []
	return value if isinstance(value, list) else []


def _next_step(state):
	"""Who is going to look again, and when — or that nobody is."""
	if not state["auto"]:
		return _(
			"Nothing is checking automatically: use Get Transcript & Recording, or tick Fetch "
			"transcripts after meetings on this Microsoft Calendar."
		)

	# The tickbox only decides whether the catch-up job *would* ask; something still has to run
	# it. On a bench with no scheduler or no worker it never runs, and "Checking again around
	# 14:48" is then exactly the same class of lie as the tickbox being off — worse, in fact,
	# because it names a time. Say what is wrong and how to end it, in the same breath.
	queue = background.health()
	if not queue["ok"]:
		return _(
			"Nothing is checking automatically: {0} {1} Use Get Transcript & Recording in the meantime."
		).format(queue["reasons"][0], queue["fixes"][0])

	if state["next_check"] and state["next_check"] > now_datetime():
		return _("Checking again around {0}.").format(
			frappe.utils.format_datetime(state["next_check"])
		)
	return _("Checking again shortly.")


def describe_state(state):
	"""The sentence a person reads. Every branch says something different, on purpose."""
	kind = state["state"]

	if kind == "not_finished":
		return _("This meeting has not finished yet.")

	if kind == "complete":
		if state["has_transcript"] and state["recordings"]:
			return _("Transcript attached. {0} recording(s) ready to download.").format(
				state["recordings"]
			)
		if state["has_transcript"]:
			return _("Transcript attached. Microsoft has no recording for this meeting.")
		return _("{0} recording(s) ready to download. Microsoft has no transcript.").format(
			state["recordings"]
		)

	if kind == "partial":
		missing = _("recording") if state["has_transcript"] else _("transcript")
		if not state["next_check"]:
			return _("Microsoft never produced the {0} for this meeting.").format(missing)
		return _("Microsoft has not finished the {0} yet. {1} {2}").format(
			missing, _next_step(state), _processing_note()
		)

	if kind == "processing":
		return _("Microsoft has not finished processing this meeting. {0} {1}").format(
			_next_step(state), _processing_note()
		)

	if kind == "nothing_found":
		return _(
			"Microsoft still has no transcript or recording a day after this meeting, so it was "
			"most likely never recorded or transcribed. You can still check by hand."
		)

	if kind == "expired":
		return _("This meeting is too old. {0}").format(_expiry_note())

	return _("This event has no Teams meeting.")


# --- fetching ------------------------------------------------------------------------


@frappe.whitelist(methods=["POST"])
def fetch_meeting_artifacts(event: str):
	"""Attach every transcript and note every recording. POST: it writes a File and fields."""
	doc, calendar_name, meeting_id = _event_for_artifacts(event)
	result = _fetch_into(doc, calendar_name, meeting_id)

	if result["error"]:
		# Microsoft said something; repeating it verbatim beats a guess about why.
		frappe.throw(
			_("Microsoft could not be asked for this meeting: {0}").format(result["error"]),
			title=_("Microsoft returned an error"),
		)

	doc.reload()
	state = _state(doc)
	return {
		"transcripts": result["transcripts"],
		"recordings": result["recordings"],
		"state": state["state"],
		"message": describe_state(state),
	}


def _fetch_into(doc, calendar_name, meeting_id):
	"""Ask Graph, write what came back, record the attempt. Never raises for 'not ready yet'.

	A transcript error and a recording error are kept apart: recordings need their own
	permission and licence, and a tenant that refuses them must not cost you the transcript.
	"""
	from frappe_microsoft365 import microsoft_transcripts as ms

	attached, recordings, error = [], [], None

	try:
		transcripts = ms.list_transcripts(calendar_name, meeting_id)
	except MsGraphError as e:
		transcripts, error = [], str(e)
	else:
		if transcripts:
			attached = _attach_transcripts(doc, calendar_name, meeting_id, transcripts)

	try:
		recordings = ms.list_recordings(calendar_name, meeting_id)
	except MsGraphError as e:
		# Surfaced only if the transcript side did not already fail — one error is a diagnosis,
		# two stacked together is noise.
		recordings = []
		if error is None and not transcripts:
			error = str(e)

	values = {
		"custom_microsoft_artifacts_attempts": cint(doc.get("custom_microsoft_artifacts_attempts")) + 1,
		"custom_microsoft_artifacts_checked_on": now_datetime(),
	}
	if attached:
		values["custom_microsoft_transcript_fetched_on"] = now_datetime()
	if recordings:
		values["custom_microsoft_recordings"] = _describe_recordings(recordings)
		values["custom_microsoft_recordings_data"] = json.dumps(
			[
				{
					"id": r.get("id"),
					"created": r.get("created_date_time"),
					"end": r.get("end_date_time"),
				}
				for r in recordings
			]
		)
	frappe.db.set_value("Event", doc.name, values, update_modified=False)

	# The status sentence is stored so the form can explain itself without calling Graph.
	doc.reload()
	frappe.db.set_value(
		"Event",
		doc.name,
		"custom_microsoft_artifacts_status",
		error or describe_state(_state(doc)),
		update_modified=False,
	)

	return {"transcripts": attached, "recordings": len(recordings), "error": error}


def _describe_recordings(recordings):
	"""One readable line per recording.

	Numbered, because a meeting past four hours comes back in pieces and "Recorded 09:00"
	twice tells you nothing about which piece is which.
	"""
	lines = []
	total = len(recordings)
	for index, recording in enumerate(recordings, start=1):
		created = recording.get("created_date_time") or ""
		when = frappe.utils.format_datetime(created) if created else _("unknown time")
		if total > 1:
			lines.append(_("Part {0} of {1} — recorded {2}").format(index, total, when))
		else:
			lines.append(_("Recorded {0}").format(when))
	return "\n".join(lines)


def _attach_transcripts(doc, calendar_name, meeting_id, transcripts):
	"""Save every transcript as a File on the Event, replacing earlier copies of this meeting.

	Every one, not only the latest: transcription can be stopped and restarted mid-meeting, and
	keeping just the last part would quietly lose the first hour of a long meeting.
	"""
	from frappe_microsoft365 import microsoft_transcripts as ms

	ordered = sorted(transcripts, key=lambda t: t.get("created_date_time") or "")
	contents = []
	for transcript in ordered:
		try:
			content = ms.get_transcript_content(calendar_name, meeting_id, transcript["id"]).get(
				"content"
			)
		except MsGraphError:
			continue
		if content:
			contents.append((transcript.get("created_date_time"), content))

	if not contents:
		return []

	_delete_previous_transcripts(doc)

	saved = []
	total = len(contents)
	for index, (created_at, content) in enumerate(contents, start=1):
		stamp = (created_at or "")[:10] or frappe.utils.nowdate()
		suffix = f"-part-{index}" if total > 1 else ""
		file_doc = _save_transcript(f"{TRANSCRIPT_PREFIX}{stamp}{suffix}.vtt", content, doc.name)
		saved.append({"file_name": file_doc.file_name, "file_url": file_doc.file_url})
	return saved


def _save_transcript(filename, content, event_name):
	from frappe.utils.file_manager import save_file

	return save_file(filename, content.encode("utf-8"), "Event", event_name, is_private=1)


def _delete_previous_transcripts(doc):
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


# --- the catch-up job ----------------------------------------------------------------


def fetch_pending():
	"""Scheduled: collect artifacts for meetings that finished recently. Never raises.

	Opt-in per connection, because this writes files onto Events by itself. Only meetings on a
	calendar with the tickbox on are ever asked about.
	"""
	try:
		settings = frappe.get_cached_doc("Microsoft Settings")
		if not settings.enabled or not settings.use_transcripts:
			return
	except Exception:
		return

	calendars = frappe.get_all(
		"Microsoft Calendar",
		filters={"enabled": 1, "authorized": 1, "fetch_artifacts_automatically": 1},
		pluck="name",
	)
	if not calendars:
		return

	done = []
	for name in _due_events(calendars):
		try:
			doc, calendar_name, meeting_id = _event_for_artifacts(name)
		except Exception:
			# An event that cannot even be addressed must not be retried every quarter of an
			# hour for a day; spend its budget in one go.
			frappe.db.set_value(
				"Event",
				name,
				"custom_microsoft_artifacts_attempts",
				len(RETRY_MINUTES),
				update_modified=False,
			)
			continue

		try:
			done.append({name: _fetch_into(doc, calendar_name, meeting_id)})
		except Exception:
			frappe.log_error(title=f"MS meeting artifacts failed: {name}")
	return done


def _due_events(calendars):
	"""Events whose next backoff step has come round. Cheap filter first, per-event maths after."""
	now = now_datetime()
	candidates = frappe.get_all(
		"Event",
		filters={
			"custom_microsoft_calendar": ["in", calendars],
			"custom_teams_join_url": ["is", "set"],
			"ends_on": ["between", [add_to_date(now, minutes=-RETRY_MINUTES[-1] - 60), now]],
			"custom_microsoft_artifacts_attempts": ["<", len(RETRY_MINUTES)],
		},
		fields=[
			"name",
			"ends_on",
			"custom_microsoft_artifacts_attempts",
			"custom_microsoft_transcript_fetched_on",
			"custom_microsoft_recordings_data",
		],
		limit=100,
	)

	due = []
	for row in candidates:
		if row.custom_microsoft_transcript_fetched_on and row.custom_microsoft_recordings_data:
			continue  # both sides landed; nothing left to ask for
		next_check = _next_check_at(
			get_datetime(row.ends_on), cint(row.custom_microsoft_artifacts_attempts)
		)
		if next_check and next_check <= now:
			due.append(row.name)
	return due


# --- download ------------------------------------------------------------------------


@frappe.whitelist()
def download_recording(event: str, recording_id: str | None = None):
	"""Stream one recording from Microsoft through Frappe.

	Not stored, and not buffered either: the bytes are pulled from Graph and pushed to the
	browser a chunk at a time. Graph's own URL cannot be given to a browser because it needs a
	bearer token, so this worker has to sit in the middle — but a Teams part reaches 1.5 GB, and
	reading ``resp.content`` would put all of that in the worker's memory before the download
	even began.

	Hence a werkzeug Response returned rather than ``frappe.local.response``: the response
	builder assigns the body to ``response.data`` (frappe/utils/response.py, ``as_raw``), which
	is a second full copy and cannot stream by construction. A whitelisted method that returns a
	Response instead has it passed through untouched — frappe/handler.py and
	frappe/api/__init__.py both check ``isinstance(data, Response)`` before doing anything else.
	"""
	doc, calendar_name, meeting_id = _event_for_artifacts(event)
	stored = _stored_recordings(doc)

	if not recording_id:
		if not stored:
			frappe.throw(_("Microsoft has no recording for this meeting. {0}").format(_expiry_note()))
		recording_id = stored[0]["id"]
	elif stored and recording_id not in [r.get("id") for r in stored]:
		# The id has to have come from this Event, or this endpoint becomes a way to pull any
		# recording the connected account can reach by guessing ids.
		frappe.throw(_("That recording does not belong to this meeting."), frappe.PermissionError)

	part = next((i for i, r in enumerate(stored, start=1) if r.get("id") == recording_id), 1)

	try:
		# stream=True defers only the body: the status line and headers have already arrived, so
		# graph_request still raises MsGraphError on a 4xx here, before any video is touched.
		# That is why this try/except is the whole of the error handling — once the Response is
		# returned there is no longer a request to fail.
		resp = graph.graph_request(
			"GET",
			f"/me/onlineMeetings/{meeting_id}/recordings/{recording_id}/content",
			calendar_name,
			raw=True,
			stream=True,
		)
	except MsGraphError as e:
		frappe.throw(
			_("Microsoft could not give us the recording: {0}. {1}").format(str(e), _expiry_note())
		)

	suffix = f"-part-{part}" if len(stored) > 1 else ""
	headers = resp.headers or {}

	response = Response(
		resp.iter_content(chunk_size=RECORDING_CHUNK_BYTES),
		# Without direct_passthrough werkzeug consumes the iterator into one buffer to measure
		# it, which is the whole thing this function exists to avoid.
		direct_passthrough=True,
		# Microsoft's own type first: it knows what it encoded. mp4 is the documented format and
		# the only one Teams has ever returned, so it is the fallback rather than a guess.
		content_type=headers.get("Content-Type") or "video/mp4",
	)
	response.headers.add(
		"Content-Disposition", "attachment", filename=f"teams-recording-{doc.name}{suffix}.mp4"
	)

	# Passed on only when Graph states it. A browser with a length draws a real progress bar for
	# what may be a twenty-minute download; a browser given a wrong one truncates the file, so
	# there is nothing sensible to invent when the header is absent.
	if headers.get("Content-Length"):
		response.headers["Content-Length"] = headers["Content-Length"]

	# The socket is ours to release: if the person cancels halfway, closing the response returns
	# the connection to the pool instead of leaving it pinned until the request times out.
	response.call_on_close(resp.close)

	return response
