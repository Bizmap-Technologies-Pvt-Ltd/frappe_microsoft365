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

Requires the transcript capability: a transcript is addressed by online meeting id, and
resolving that from a join URL needs OnlineMeetings.Read (ReadWrite on sites that already had it).
"""

import json
import re
from datetime import datetime, timedelta

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


def _teams_tab_note():
	"""Where a person can see for themselves what Microsoft actually holds.

	This app can only report what Graph answers, and Graph goes quiet for reasons that have
	nothing to do with the meeting — an expired one, a tenant switch, a permission. Teams shows
	the same meeting's files from the other side, so a single look there splits "nothing was
	ever recorded" from "the connection cannot reach it" without anyone guessing.
	"""
	return _(
		"Open the meeting in the Teams calendar and check its Recordings and Transcripts tab."
	)


class MeetingUnavailable(frappe.ValidationError):
	"""There is nothing to fetch for this Event, and asking again will not change that.

	No Teams meeting on it, or a join link Microsoft matches to nothing in this tenant (another
	organisation's meeting, or one that has aged out). These end the catch-up job's schedule for
	the Event; every other failure is retried on the normal backoff.
	"""


#: Graph answers that say "not now" rather than "no": throttling, and Microsoft's own outages.
_TEMPORARY_STATUS = re.compile(r"\((429|500|502|503|504)\)|rate limit", re.IGNORECASE)


def _is_temporary(error):
	"""A Graph error that says nothing about the meeting: throttling or an outage."""
	return bool(_TEMPORARY_STATUS.search(str(error)))


def _status_of(error):
	match = re.search(r"\((\d{3})\)", str(error))
	return match.group(1) if match else "429"


def _event_for_artifacts(event_name, resolve=True):
	"""The Event, its calendar and its online meeting id, with every precondition checked."""
	doc = frappe.get_doc("Event", event_name)
	doc.check_permission("read")

	calendar_name = doc.get("custom_microsoft_calendar")
	join_url = doc.get("custom_teams_join_url")
	if not calendar_name or not join_url:
		frappe.throw(_("This event has no Teams meeting."), MeetingUnavailable)

	calendar = frappe.get_doc("Microsoft Calendar", calendar_name)
	_check_owner(calendar)
	# Split, because from here the two are indistinguishable and their fixes are not: one is a
	# tickbox on a form, the other is a sign-in that has to happen in a browser. "Disabled or
	# not authorized" made the reader go and check both.
	if not calendar.enabled:
		frappe.throw(
			_("The Microsoft Calendar {0} is switched off. Open it and tick Enabled.").format(calendar_name)
		)
	if not calendar.authorized:
		frappe.throw(
			_(
				"The Microsoft Calendar {0} is not signed in to Microsoft. Open it and click "
				"Authorize Microsoft Access."
			).format(calendar_name)
		)

	# Gated on the start, never on the booked end. A calendar slot is a reservation, not a
	# record: people book an hour and talk for four minutes, and Microsoft has the transcript
	# ready minutes after the call actually ends — not after the slot they never used up. This
	# check once refused to look for another hour while the files sat waiting in Teams.
	#
	# Before the start there is genuinely nothing, and saying so is kinder than a Graph call
	# that comes back empty for a reason the person cannot see.
	if doc.starts_on and get_datetime(doc.starts_on) > now_datetime():
		frappe.throw(_("This meeting has not started yet."))

	meeting_id = doc.get("custom_microsoft_online_meeting_id")
	if not meeting_id and resolve:
		from frappe_microsoft365 import microsoft_transcripts as ms

		try:
			meeting_id = _resolve_online_meeting_id(
				calendar_name, join_url, event_id=doc.get("custom_microsoft_event_id")
			)
		except MsGraphError as e:
			# This step needs OnlineMeetings.Read, which the transcript permission does not
			# include — so it is the one that fails for people who granted transcript consent and
			# stopped. Left to fall through, the dialog showed a raw Graph 403 against a path
			# nobody recognises.
			if _is_temporary(e):
				# Said as what it is. The raw text quotes the whole $filter, join link and all, and
				# reads like the link is wrong when Microsoft simply did not answer this time.
				frappe.throw(
					_(
						"Microsoft did not answer the meeting lookup this time ({0}). That is "
						"temporary: try again in a few minutes."
					).format(_status_of(e)),
					type(e),
				)
			ms._wrap_join_url_403(e)
		if not meeting_id:
			# Matched nothing rather than refused. An external organiser's meeting lives in their
			# tenant and is not in this one at all, which is the cause that "too old" hides.
			frappe.throw(
				_(
					"Microsoft could not match this join link to a meeting — either another "
					"organisation hosted it, or it has aged out. {0}"
				).format(_expiry_note()),
				MeetingUnavailable,
			)
		frappe.db.set_value(
			"Event", doc.name, "custom_microsoft_online_meeting_id", meeting_id, update_modified=False
		)

	return doc, calendar_name, meeting_id


# --- occurrences ---------------------------------------------------------------------
#
# Every occurrence of a recurring Teams meeting shares one join link and one online meeting, so
# Graph lists the transcripts and recordings of the whole series under it. Which ones belong to
# this Event has to be decided here, by time: an item belongs to the occurrence it was created
# during. A one-off meeting is simply a series of one.

#: How long before its booked start an occurrence's items may begin. People open the meeting a
#: little early, and transcription started then is still that meeting's.
OCCURRENCE_LEAD = timedelta(hours=1)

#: How long after its booked end. Meetings overrun, and a recording that restarts at the four-hour
#: split is created well after the slot of a short booking. Kept far below a day, so a daily
#: series' next occurrence is never in reach.
OCCURRENCE_TAIL = timedelta(hours=6)

#: Frappe's repeat_on values that step by months, and by how many.
_MONTH_STEPS = {"Monthly": 1, "Quarterly": 3, "Half Yearly": 6, "Yearly": 12}

_WEEKDAYS = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")

#: The date a transcript File is named for: teams-transcript-2026-09-04[-part-2][hash].vtt
_STAMP = re.compile(r"^teams-transcript-(\d{4}-\d{2}-\d{2})")


def _occurrence(doc, now=None):
	"""(start, end) of the occurrence of this Event that artifacts are being asked about.

	For a one-off Event, the Event itself. A repeating Frappe Event is one record for the whole
	series — exactly as Graph keeps one series master — so the occurrence is the latest one that
	has started by ``now``: the one whose transcript can exist. It used to be the first one
	forever, so the catch-up job never looked at a repeating Event again after its first week.
	"""
	starts_on = get_datetime(doc.get("starts_on")) if doc.get("starts_on") else None
	ends_on = get_datetime(doc.get("ends_on")) if doc.get("ends_on") else None
	if not doc.get("repeat_this_event") or not starts_on:
		return starts_on, ends_on

	start = _latest_repeat_start(doc, starts_on, now or now_datetime())
	return start, (start + (ends_on - starts_on)) if ends_on else None


def _latest_repeat_start(doc, first, now):
	"""The start of the last occurrence of a repeating Event at or before ``now``.

	Follows the series as it was sent to Outlook (microsoft_calendar_sync._recurrence): a weekly
	series always includes its own start day, and repeat_till is the last day it can fall on.
	"""
	limit = now
	if doc.get("repeat_till"):
		limit = min(limit, datetime.combine(get_datetime(doc.repeat_till).date(), first.time()))
	if limit <= first:
		return first

	kind = doc.get("repeat_on")
	if kind in ("Daily", "Weekly"):
		days = set(range(7))
		if kind == "Weekly":
			days = {i for i, day in enumerate(_WEEKDAYS) if doc.get(day)} | {first.weekday()}
		day = limit.date()
		# A week back always finds one: the start day itself is in every weekly series.
		for _offset in range(8):
			candidate = datetime.combine(day, first.time())
			if candidate <= limit and candidate.weekday() in days:
				return max(candidate, first)
			day -= timedelta(days=1)
		return first

	step = _MONTH_STEPS.get(kind)
	if not step:
		return first
	months = ((limit.year - first.year) * 12 + limit.month - first.month) // step * step
	while months > 0:
		candidate = get_datetime(add_to_date(first, months=months))
		if candidate <= limit:
			return candidate
		months -= step
	return first


def _occurrence_window(start, end):
	"""The span an occurrence's transcripts and recordings are created in, in system time.

	(None, None) means unbounded on that side — an Event with no start cannot be placed at all.
	"""
	if not start:
		return None, None
	return start - OCCURRENCE_LEAD, (end or start) + OCCURRENCE_TAIL


def _graph_instant(value):
	"""A Graph timestamp (UTC, e.g. 2026-09-04T10:20:00.1234567Z) as naive system time, or None."""
	if not value:
		return None
	from zoneinfo import ZoneInfo

	from dateutil import parser
	from frappe.utils import get_system_timezone

	try:
		parsed = parser.isoparse(value)
	except (ValueError, TypeError, OverflowError):
		return None
	if parsed.tzinfo is None:
		parsed = parsed.replace(tzinfo=ZoneInfo("UTC"))
	return parsed.astimezone(ZoneInfo(get_system_timezone())).replace(tzinfo=None)


def _in_window(item, window):
	"""Was this transcript or recording created during the occurrence?

	One Graph does not date cannot be placed in any occurrence, so it is kept rather than lost:
	Graph always dates them, and dropping a transcript on a technicality is the worse mistake.
	"""
	lo, hi = window
	created = _graph_instant(item.get("created_date_time"))
	if created is None:
		return True
	return (lo is None or created >= lo) and (hi is None or created <= hi)


def _utc_day(moment):
	"""The UTC calendar date of a naive system-time datetime: the date transcripts are named by."""
	from zoneinfo import ZoneInfo

	from frappe.utils import get_system_timezone

	aware = moment.replace(tzinfo=ZoneInfo(get_system_timezone()))
	return aware.astimezone(ZoneInfo("UTC")).date()


def _stamp_day(file_name):
	match = _STAMP.match(file_name or "")
	if not match:
		return None
	try:
		return datetime.strptime(match.group(1), "%Y-%m-%d").date()
	except ValueError:
		return None


def _day_in_window(day, window):
	"""Could a transcript named for UTC date ``day`` have been created inside ``window``?

	The file name keeps only the date, so this asks whether the two overlap at all — generous by
	up to a day, which is never enough to reach the next week's occurrence.
	"""
	lo, hi = window
	return (lo is None or day >= _utc_day(lo)) and (hi is None or day <= _utc_day(hi))


def _day_in_some_occurrence(doc, day):
	"""Does a transcript named for UTC ``day`` fall in any occurrence of this Event?"""
	from zoneinfo import ZoneInfo

	from frappe.utils import get_system_timezone

	# The last moment of that UTC day, in system time, plus the lead: the latest occurrence that
	# can have started by then is the only one this day can belong to.
	end_of_day = datetime.combine(day, datetime.max.time()).replace(tzinfo=ZoneInfo("UTC"))
	moment = end_of_day.astimezone(ZoneInfo(get_system_timezone())).replace(tzinfo=None)
	return _day_in_window(day, _occurrence_window(*_occurrence(doc, now=moment + OCCURRENCE_LEAD)))


def _is_new_occurrence(doc, occurrence_start):
	"""Has a repeating Event moved on to an occurrence nobody has checked yet?"""
	if not doc.get("repeat_this_event") or not occurrence_start:
		return False
	checked_on = doc.get("custom_microsoft_artifacts_checked_on")
	return not checked_on or get_datetime(checked_on) < occurrence_start


def _counters(doc, occurrence_start):
	"""(attempts, has_transcript, recordings) for the occurrence being asked about.

	A repeating Event keeps one set of fields for the whole series. What they say belongs to the
	occurrence that was last checked; once a later one has started, they describe the past, and
	that occurrence starts from nothing — otherwise week two would inherit week one's spent
	attempts and "transcript attached" and never be asked about at all.
	"""
	attempts = cint(doc.get("custom_microsoft_artifacts_attempts"))
	fetched_on = doc.get("custom_microsoft_transcript_fetched_on")
	recordings = _stored_recordings(doc)
	if not doc.get("repeat_this_event") or not occurrence_start:
		return attempts, bool(fetched_on), recordings

	if _is_new_occurrence(doc, occurrence_start):
		return 0, False, []
	return attempts, bool(fetched_on and get_datetime(fetched_on) >= occurrence_start), recordings


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
	now = now_datetime()
	# The occurrence, not the Event: for a repeating Event the two differ from the second week on.
	starts_on, ends_on = _occurrence(doc, now)
	attempts, has_transcript, recordings = _counters(doc, starts_on)

	state = {
		"has_transcript": has_transcript,
		"recordings": len(recordings),
		"attempts": attempts,
		"expires_on": add_to_date(ends_on, days=MEETING_EXPIRY_DAYS) if ends_on else None,
		"next_check": None,
		"auto": _auto_enabled(doc),
		# A meeting can be over long before its slot is. Kept apart from the states below so a
		# message can mention it without any of them having to branch on it.
		"still_booked": bool(ends_on and ends_on > now),
		"booked_until": ends_on,
		# What the last attempt was told. A tenant that has just had Graph access to transcripts
		# switched on answers inconsistently for a while — one call 403s, the next returns an
		# empty list — and reporting that empty list as "still processing" sends someone off to
		# wait for Microsoft when the thing they are actually waiting for is their own setting
		# reaching every node.
		"last_refusal": _remembered_refusal(doc),
	}

	if not doc.get("custom_teams_join_url"):
		return {**state, "state": "no_meeting"}
	if starts_on and starts_on > now:
		return {**state, "state": "not_started"}

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


def _remembered_refusal(doc):
	"""The tenant-switch refusal, if that is what the last attempt hit.

	Matched against the sentence this app generates rather than Microsoft's, because ours is
	the one that got stored — and against a distinctive fragment rather than the whole thing,
	so a reworded hint degrades to the ordinary message instead of lying about a new one.
	"""
	from frappe_microsoft365.microsoft_transcripts import TENANT_SWITCH, TENANT_SWITCH_TEXT

	status = doc.get("custom_microsoft_artifacts_status") or ""
	if TENANT_SWITCH in status or TENANT_SWITCH_TEXT in status:
		return True
	# Our own replacement wording never quotes Microsoft, so match it too.
	return "Transcript API access" in status


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

	if kind == "not_started":
		return _("This meeting has not started yet.")

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
			# One side landed, so the connection and the permissions are demonstrably fine —
			# which makes "nobody started it" the likely answer and a setup hunt a waste.
			return _(
				"Microsoft never produced the {0} for this meeting, so it was most likely never "
				"started. {1}"
			).format(missing, _teams_tab_note())
		return _("Microsoft has not finished the {0} yet. {1} {2}").format(
			missing, _next_step(state), _processing_note()
		)

	if kind == "processing":
		if state.get("last_refusal"):
			# Names the setting for two reasons: it is where the reader has to go, and it is how
			# this sentence recognises itself on the next attempt. The status field holds exactly
			# one message, so a memory that does not survive being rewritten is not a memory.
			return _(
				"Microsoft refused this recently: Graph access to transcripts was off for the "
				"tenant (Teams admin center > Meetings > Meeting settings > Transcript API "
				"access). If you just switched it on it takes a while to reach every server, so "
				"an empty answer now is not evidence the meeting has no transcript."
			)
		if state.get("still_booked"):
			# The slot has not run out, but the call may well be over — Teams produces nothing
			# until it actually ends, so "nothing yet" here means one of two things and it
			# would be wrong to claim either.
			return _(
				"Nothing from Microsoft yet. This meeting is booked until {0}; if it has already "
				"ended, the files usually appear a few minutes later. {1}"
			).format(frappe.utils.format_datetime(state["booked_until"]), _next_step(state))
		return _("Microsoft has not finished processing this meeting. {0} {1}").format(
			_next_step(state), _processing_note()
		)

	if kind == "nothing_found":
		# "You can still check by hand" was the end of the road: true, and nothing a person could
		# act on. Graph answering 200-with-nothing all day is usually a meeting nobody recorded,
		# but it looks identical to a connection that cannot reach what exists — and this branch
		# is reached from the form, which never sees the 403 the catch-up job stored. One look in
		# Teams tells those apart, so name it and name what to do with each answer.
		return _(
			"Microsoft still has no transcript or recording a day after this meeting, so it was "
			"most likely never recorded or transcribed. {0} If either is there, the gap is in this "
			"connection: run Microsoft Settings > Troubleshoot > Run Diagnostics."
		).format(_teams_tab_note())

	if kind == "expired":
		# Graph stops serving long before Teams deletes anything, so "too old" on its own sends
		# people away from files that are still sitting there.
		return _("This meeting is too old, but the files may still exist. {0} {1}").format(
			_expiry_note(), _teams_tab_note()
		)

	return _("This event has no Teams meeting.")


# --- fetching ------------------------------------------------------------------------


@frappe.whitelist(methods=["POST"])
def fetch_meeting_artifacts(event: str):
	"""Attach every transcript and note every recording. POST: it writes a File and fields."""
	doc, calendar_name, meeting_id = _event_for_artifacts(event)
	result = _fetch_into(doc, calendar_name, meeting_id)

	if result["error"]:
		# Each layer that caught this on the way up added its own line to the message log, so
		# the dialog was showing the same 403 three times over, each copy slightly longer than
		# the last. Clear those and say it once — one error is a diagnosis, three stacked is
		# noise a person has to read past to find the sentence that helps.
		frappe.clear_messages()
		frappe.throw(result["error"], title=_("Microsoft returned an error"))

	doc.reload()
	state = _state(doc)
	return {
		"transcripts": result["transcripts"],
		"recordings": result["recordings"],
		"state": state["state"],
		"message": describe_state(state),
	}


def _fetch_into(doc, calendar_name, meeting_id, scheduled=False, attempt=None):
	"""Ask Graph, write what came back, record the attempt. Never raises for 'not ready yet'.

	A transcript error and a recording error are kept apart: recordings are consented
	separately, and a tenant that refuses them must not cost you the transcript — nor answer
	the refusal with advice about transcripts.

	Only what was created during this Event's own occurrence is taken. A recurring Teams meeting
	lists every occurrence's transcripts under the one meeting, and each occurrence's Event used
	to receive all of them, labelled part 1, part 2 …

	``scheduled`` is the catch-up job: it commits as it goes, so a worker killed mid-run keeps
	the transcripts it already saved, and it does not download transcripts it already has.
	``attempt`` is the number it already claimed for this visit; otherwise this is the next one.
	"""
	from frappe_microsoft365 import microsoft_transcripts as ms

	occurrence_start, occurrence_end = _occurrence(doc)
	window = _occurrence_window(occurrence_start, occurrence_end)
	attempts, has_transcript, _recordings = _counters(doc, occurrence_start)
	new_occurrence = _is_new_occurrence(doc, occurrence_start)
	attempt = attempt if attempt is not None else attempts + 1
	attached, recordings, error, gone = [], [], None, False

	try:
		transcripts = [t for t in ms.list_transcripts(calendar_name, meeting_id) if _in_window(t, window)]
	except MsGraphError as e:
		transcripts, error = [], str(e)
		# Graph no longer knows the meeting id at all; no later attempt will find it either.
		gone = isinstance(e, graph.MsGraphNotFound)
	else:
		if transcripts and not (scheduled and has_transcript and _already_attached(doc, window, transcripts)):
			try:
				attached = _attach_transcripts(doc, calendar_name, meeting_id, transcripts, window)
			except MsGraphError as e:
				# Raised only when every part was refused; anything partial has already been
				# saved. Routed through the same field as the others so one refusal is reported
				# once, with its own hint, rather than escaping as a bare traceback.
				error = str(e)
			if attached and scheduled:
				# The files are on disk the moment they are saved; their rows must not be the
				# part a timeout rolls back.
				frappe.db.set_value(
					"Event", doc.name, "custom_microsoft_transcript_fetched_on", now_datetime(),
					update_modified=False,
				)
				frappe.db.commit()  # nosemgrep: frappe-manual-commit -- per-event progress in a scheduled job
				frappe.db.savepoint(EVENT_SAVEPOINT)

	try:
		recordings = [r for r in ms.list_recordings(calendar_name, meeting_id) if _in_window(r, window)]
	except MsGraphError as e:
		# Surfaced only if the transcript side did not already fail — one error is a diagnosis,
		# two stacked together is noise.
		recordings = []
		if error is None and not transcripts:
			error = str(e)

	values = {
		# Spent in one go when the meeting is gone: asking a day long would be fourteen 404s.
		"custom_microsoft_artifacts_attempts": len(RETRY_MINUTES) if gone else attempt,
		"custom_microsoft_artifacts_checked_on": now_datetime(),
	}
	if new_occurrence and not recordings:
		# Last week's recordings are not this week's: an empty answer for a new occurrence of a
		# repeating Event must not leave them offered as though they were.
		values["custom_microsoft_recordings"] = None
		values["custom_microsoft_recordings_data"] = None
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


def _attach_transcripts(doc, calendar_name, meeting_id, transcripts, window=(None, None)):
	"""Save every transcript as a File on the Event, replacing earlier copies of this occurrence.

	Every one, not only the latest: transcription can be stopped and restarted mid-meeting, and
	keeping just the last part would quietly lose the first hour of a long meeting.
	"""
	from frappe_microsoft365 import microsoft_transcripts as ms

	ordered = sorted(transcripts, key=lambda t: t.get("created_date_time") or "")
	contents = []
	refusal = None
	for transcript in ordered:
		try:
			fetched = ms.get_transcript_content(calendar_name, meeting_id, transcript["id"])
		except MsGraphError as e:
			# One part failing must not cost the others, so this keeps going — but it is kept,
			# not dropped. Microsoft listing a transcript it will not then hand over is a
			# refusal, and swallowing it reported "Microsoft has not finished processing" about
			# a transcript that was finished and being withheld.
			refusal = refusal or e
			continue
		content = fetched.get("content")
		if content:
			contents.append(
				(transcript.get("created_date_time"), content, _transcript_extension(fetched.get("format")))
			)

	if not contents:
		if refusal:
			raise refusal
		return []

	_delete_previous_transcripts(doc, window)

	saved = []
	total = len(contents)
	for index, (created_at, content, extension) in enumerate(contents, start=1):
		stamp = (created_at or "")[:10] or frappe.utils.nowdate()
		suffix = f"-part-{index}" if total > 1 else ""
		file_doc = _save_transcript(f"{TRANSCRIPT_PREFIX}{stamp}{suffix}{extension}", content, doc.name)
		saved.append({"file_name": file_doc.file_name, "file_url": file_doc.file_url})
	return saved


def _transcript_extension(fmt):
	"""``.txt`` for the plain-text transcript a tenant without speaker attribution gets, ``.vtt``
	otherwise. Decided per part, from what Microsoft actually sent: the file has no voice tags,
	and calling it VTT would promise speaker names it does not contain."""
	from frappe_microsoft365.microsoft_transcripts import PLAIN_TEXT_FORMAT

	return ".txt" if fmt == PLAIN_TEXT_FORMAT else ".vtt"


def _save_transcript(filename, content, event_name):
	"""Write the transcript as a File on the Event, and be clear about whose fault it is if that fails.

	Saving a File runs every after_insert hook any installed app has put on File — an S3
	offloader, a virus scanner, a storage quota. When one of those is misconfigured the upload
	raises from inside this call stack, Frappe blames the app that happens to be on top of it,
	and the person is handed a botocore traceback naming this integration for a missing AWS
	credential. Seen live: the transcript had been fetched perfectly and the error read as
	though Microsoft had failed.

	The distinction is worth keeping because the two have nothing to do with each other and
	fixing one never fixes the other.
	"""
	from frappe.utils.file_manager import save_file

	try:
		return save_file(filename, content.encode("utf-8"), "Event", event_name, is_private=1)
	except Exception as e:
		frappe.log_error(title=f"MS transcript could not be stored: {event_name}")
		frappe.throw(
			_(
				"Microsoft gave us the transcript, but Frappe could not store the file: {0}. "
				"That is a file-storage problem on this site, not a Microsoft one — check the "
				"apps that handle attachments."
			).format(str(e)),
			title=_("The transcript arrived but could not be saved"),
		)


def _transcript_files(doc):
	return frappe.get_all(
		"File",
		filters={
			"attached_to_doctype": "Event",
			"attached_to_name": doc.name,
			"file_name": ["like", f"{TRANSCRIPT_PREFIX}%"],
		},
		fields=["name", "file_name"],
	)


def _delete_previous_transcripts(doc, window=(None, None)):
	"""Remove this occurrence's earlier copies before the new ones are saved.

	Matched by prefix, not by exact name: Frappe appends a content hash when a file of that name
	already exists, so re-fetching would otherwise pile up near-identical copies. And only the
	files named for a date inside this occurrence: a repeating Event keeps every earlier week's
	transcript, and files another occurrence's transcripts were wrongly attached under before
	1.2.0 are for a person to remove (remove_misplaced_transcripts), never for a fetch.
	"""
	for existing in _transcript_files(doc):
		day = _stamp_day(existing.file_name)
		if day is None or _day_in_window(day, window):
			frappe.delete_doc("File", existing.name, force=True, ignore_permissions=True)


def _already_attached(doc, window, transcripts):
	"""Does the Event already hold one file per transcript Microsoft lists for this occurrence?

	The catch-up job keeps asking until the schedule ends, for recording parts that come later;
	downloading and rewriting the same transcripts on each of those visits would be fourteen
	copies of the same work.
	"""
	held = 0
	for existing in _transcript_files(doc):
		day = _stamp_day(existing.file_name)
		if day and _day_in_window(day, window):
			held += 1
	return held == len(transcripts)


# --- the catch-up job ----------------------------------------------------------------


#: Most Events one run of the catch-up job asks about. Runs are a quarter of an hour apart, and
#: each Event is a handful of Graph calls, so this bounds a run well inside a worker's timeout.
MAX_EVENTS_PER_RUN = 100

#: Undone to on a failure, so only that Event's half-written changes go — never the claim
#: already committed, nor anything a caller had open.
EVENT_SAVEPOINT = "ms_meeting_artifacts_event"

#: How far back an Event that was never checked at all is still picked up — one stopped by an
#: earlier version and reopened on upgrade, or one on a calendar whose tickbox was turned on
#: since. Comfortably inside Microsoft's 60-day expiry.
CATCH_UP_DAYS = 7


def fetch_pending():
	"""Scheduled: collect artifacts for meetings that finished recently. Never raises.

	Opt-in per connection, because this writes files onto Events by itself. Only meetings on a
	calendar with the tickbox on are ever asked about.

	Each Event is committed on its own. A run that times out, or a worker that is killed, loses
	at most the Event it was on — before, the whole run's attempt counters went with it and the
	same meetings were asked about again and again while transcript files piled up on disk.
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
		frappe.db.savepoint(EVENT_SAVEPOINT)
		try:
			result = _check_one(name)
		except Exception:
			# Nothing is allowed out of a scheduled job, and one Event must never stop the rest.
			frappe.db.rollback(save_point=EVENT_SAVEPOINT)
			frappe.log_error(title=f"MS meeting artifacts failed: {name}")
			continue
		frappe.db.commit()  # nosemgrep: frappe-manual-commit -- per-event progress in a scheduled job
		if result is not None:
			done.append({name: result})
	return done


def _check_one(name):
	"""One Event's turn in the catch-up job. Every way it ends leaves a status saying why."""
	try:
		doc = frappe.get_doc("Event", name)
	except frappe.DoesNotExistError:
		return None

	start, end = _occurrence(doc)
	attempt = _counters(doc, start)[0] + 1
	if _is_new_occurrence(doc, start):
		# Cleared before the claim below makes this occurrence look already checked.
		frappe.db.set_value(
			"Event",
			name,
			{"custom_microsoft_recordings": None, "custom_microsoft_recordings_data": None},
			update_modified=False,
		)
	# Claimed before Microsoft is asked, and committed: a worker killed halfway through still
	# spends this attempt, so one meeting that kills workers is not retried every quarter of an
	# hour for a day — and the status says what was going on when it stopped.
	_record(
		name,
		attempt,
		_(
			"Checking with Microsoft (attempt {0} of {1}). If this stays here, that check was "
			"interrupted."
		).format(attempt, len(RETRY_MINUTES)),
	)
	frappe.db.commit()  # nosemgrep: frappe-manual-commit -- see above
	frappe.db.savepoint(EVENT_SAVEPOINT)

	try:
		doc, calendar_name, meeting_id = _event_for_artifacts(name)
		return _fetch_into(doc, calendar_name, meeting_id, scheduled=True, attempt=attempt)
	except MeetingUnavailable as e:
		_stop(name, e, attempt, permanent=True)
	except graph.MsGraphNotFound as e:
		_stop(name, e, attempt, permanent=True)
	except Exception as e:
		# Everything else is temporary until proven otherwise: a 503 or a 429 while resolving the
		# join link, a token refresh that failed for a moment, a file hook that was down. Each
		# used to spend the whole day's budget at once, so one blip after a meeting meant its
		# transcript was never fetched.
		frappe.db.rollback(save_point=EVENT_SAVEPOINT)
		if not isinstance(e, MsGraphError):
			frappe.log_error(title=f"MS meeting artifacts failed: {name}")
		_stop(name, e, attempt, permanent=False, end=end)
	return None


def _record(name, attempts, status):
	frappe.db.set_value(
		"Event",
		name,
		{
			"custom_microsoft_artifacts_attempts": attempts,
			"custom_microsoft_artifacts_checked_on": now_datetime(),
			"custom_microsoft_artifacts_status": status,
		},
		update_modified=False,
	)


def _stop(name, error, attempt, permanent, end=None):
	"""Record a failed attempt, and say in the status whether and when it will be tried again."""
	reason = frappe.utils.strip_html(str(error)).strip() or type(error).__name__
	if permanent:
		_record(name, len(RETRY_MINUTES), _("Stopped checking: {0}").format(reason))
		return
	next_check = _next_check_at(end, attempt)
	if next_check:
		status = _("{0} Trying again around {1}.").format(reason, frappe.utils.format_datetime(next_check))
	else:
		status = _("Stopped checking after {0} attempts. The last one failed: {1}").format(
			len(RETRY_MINUTES), reason
		)
	_record(name, attempt, status)


def _due_events(calendars, now=None):
	"""Events whose next backoff step has come round, the ones still missing something first.

	The whole "is it due" question is asked in SQL. It used to fetch any 100 candidates and
	then skip, in Python, the ones already complete or not yet due — so on a busy site the
	100 could be all skips and a meeting that was due was never reached.

	Events that already have a transcript and a recording are still asked about until their
	schedule ends: a meeting past four hours comes back as several recordings, and the later
	parts appear later. They sort last, behind every Event still waiting for its first file.
	"""
	from frappe.query_builder import Case, Criterion
	from frappe.query_builder.functions import Coalesce

	now = now or now_datetime()
	event = frappe.qb.DocType("Event")
	attempts = Coalesce(event.custom_microsoft_artifacts_attempts, 0)
	complete = (
		Case()
		.when(
			event.custom_microsoft_transcript_fetched_on.isnotnull()
			& (Coalesce(event.custom_microsoft_recordings_data, "") != ""),
			1,
		)
		.else_(0)
	)
	# ends_on + RETRY_MINUTES[attempts] <= now, as one comparison per step so it is plain SQL on
	# every database Frappe runs on.
	due = Criterion.any(
		[
			(attempts == step) & (event.ends_on <= add_to_date(now, minutes=-minutes))
			for step, minutes in enumerate(RETRY_MINUTES)
		]
	)
	recent = event.ends_on >= add_to_date(now, minutes=-RETRY_MINUTES[-1] - 60)
	never_checked = event.custom_microsoft_artifacts_checked_on.isnull() & (
		event.ends_on >= add_to_date(now, days=-CATCH_UP_DAYS)
	)
	rows = (
		frappe.qb.from_(event)
		.select(event.name, event.ends_on, complete.as_("complete"))
		.where(event.custom_microsoft_calendar.isin(calendars))
		.where(Coalesce(event.custom_teams_join_url, "") != "")
		.where(Coalesce(event.repeat_this_event, 0) == 0)
		.where(attempts < len(RETRY_MINUTES))
		.where(recent | never_checked)
		.where(due)
		.orderby(complete)
		.orderby(event.ends_on)
		.limit(MAX_EVENTS_PER_RUN)
		.run(as_dict=True)
	)
	rows = [(cint(r.complete), get_datetime(r.ends_on), r.name) for r in rows]
	rows += _due_repeating_events(calendars, now)
	rows.sort()
	return [name for _complete, _ends, name in rows[:MAX_EVENTS_PER_RUN]]


def _due_repeating_events(calendars, now):
	"""Repeating Events whose current occurrence is due, as (complete, ends, name) rows.

	Worked out in Python, because the occurrence is: the Event's own starts_on/ends_on are the
	first occurrence's, and the SQL above would never see a repeating Event again after it.
	There are few of these, and each is one cheap calculation.
	"""
	candidates = frappe.get_all(
		"Event",
		filters={
			"custom_microsoft_calendar": ["in", calendars],
			"custom_teams_join_url": ["is", "set"],
			"repeat_this_event": 1,
			"starts_on": ["<=", now],
		},
		or_filters=[
			["repeat_till", "is", "not set"],
			["repeat_till", ">=", add_to_date(now, days=-2).date()],
		],
		fields=[
			"name",
			"starts_on",
			"ends_on",
			"repeat_this_event",
			"repeat_on",
			"repeat_till",
			*_WEEKDAYS,
			"custom_microsoft_artifacts_attempts",
			"custom_microsoft_artifacts_checked_on",
			"custom_microsoft_transcript_fetched_on",
			"custom_microsoft_recordings_data",
		],
	)
	due = []
	for row in candidates:
		start, end = _occurrence(row, now)
		if not end or end > now or end < add_to_date(now, minutes=-RETRY_MINUTES[-1] - 60):
			continue
		attempts, has_transcript, recordings = _counters(row, start)
		next_check = _next_check_at(end, attempts)
		if next_check and next_check <= now:
			due.append((int(bool(has_transcript and recordings)), end, row.name))
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
			frappe.throw(
				_("Microsoft has no recording for this meeting, so most likely nobody started one. {0}").format(
					_teams_tab_note()
				)
			)
		recording_id = stored[0]["id"]
	elif recording_id not in [r.get("id") for r in stored]:
		# The id has to have come from this Event, or this endpoint becomes a way to pull any
		# recording the connected account can reach by guessing ids. That holds just as much
		# when nothing is stored yet: an empty list means no id belongs here, not that any does.
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
		detail = str(e)
		if "403" in detail or "Forbidden" in detail:
			# Listing a recording and reading its bytes take the same permission, so a 403 that
			# only appears here is consent that was never granted — not this meeting ageing out.
			# The expiry note would send someone hunting in OneDrive for a file Microsoft is
			# refusing rather than missing.
			from frappe_microsoft365 import microsoft_transcripts as ms

			frappe.throw(
				_("Microsoft refused this recording: {0}. {1}").format(detail, ms._recording_perm_hint())
			)
		frappe.throw(
			_("Microsoft could not give us the recording: {0}. {1}").format(detail, _expiry_note())
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
