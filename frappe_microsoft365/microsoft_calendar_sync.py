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

What waits for Graph, and what does not
--------------------------------------
Saving an Event that Microsoft has never seen calls Graph inside the save: the join link Graph
hands back is written onto the very document the browser is waiting for, and a link that turns
up "in a minute" is not one anybody can paste into the invitation they are writing. Every other
Graph call a save triggers — patching an event that already exists, deleting one — is queued
instead, because nothing on screen depends on the answer and a throttled Graph can otherwise
hold a save open for ten seconds. The scheduled pass queues as well, one job per calendar, so
that a slow calendar cannot spend the budget of the calendars behind it in the list.

All Graph calls go through ``microsoft_graph`` (auth/refresh/paging/clean errors).
Everything is guarded so an unconfigured / unauthorized site never raises on schedule.
"""

import time

import frappe
from frappe import _
from frappe.utils import (
	add_to_date,
	cint,
	get_datetime,
	get_system_timezone,
	now_datetime,
)

from frappe_microsoft365 import background
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError, MsGraphResyncRequired

EVENT_SELECT = (
	"id,subject,bodyPreview,start,end,location,isAllDay,isCancelled,"
	"onlineMeeting,webLink,type,lastModifiedDateTime,"
	# Graph returns only the properties asked for, so the invitation side of an event —
	# who was invited, who organised it and what we replied — has to be selected explicitly
	# or it never reaches Frappe at all.
	"attendees,organizer,isOrganizer,responseStatus,responseRequested,"
	# Which series an occurrence belongs to. Without it there is no way to tell an occurrence of
	# a series this app pushed from an ordinary Outlook event, and every one of them would be
	# mirrored back as a separate Frappe Event beside the repeating one that created them.
	"seriesMasterId"
)

#: How much of the calendar the delta window covers. Graph fixes the window when the delta
#: is initialised and never moves it, so its far edge comes closer every day.
WINDOW_PAST_DAYS = 30
#: The look-ahead the sync promises: an event starting less than this far ahead is always
#: inside the window. The window used to be re-initialised only 14 days before it ran out, so
#: by then it saw two weeks ahead, and a meeting booked for next quarter was not pulled at all.
WINDOW_FUTURE_DAYS = 180
#: How often the window is re-initialised. It is asked for this much further ahead than the
#: look-ahead, and renewed as soon as the look-ahead would drop below WINDOW_FUTURE_DAYS, so a
#: renewal (one full read of the window) happens about once a month.
WINDOW_RESET_DAYS = 30

#: How many Microsoft ids go into one prefetch query. A delta page can carry thousands of them,
#: and at that size an ``IN`` clause stops behaving like a lookup: the statement grows past what
#: the server wants to parse and the optimiser gives up on the index. Chunking keeps every
#: prefetch a short indexed read, and still costs one query per 500 events rather than one each.
PREFETCH_CHUNK = 500

#: The scheduled pass puts one job per calendar on the long queue. A first sync reads a 210-day
#: window page by page and pushes whatever is waiting, which does not fit the default queue's
#: 300 seconds; 1500 is the long queue's own budget.
SYNC_QUEUE = "long"
SYNC_JOB_TIMEOUT = 1500

#: Past this, a sync is no longer something to make a person sit and watch, so ``Sync Now`` sends
#: it to a worker instead. Chosen well above a warm incremental run (a second or two) and below
#: the point where a browser request starts to look hung rather than slow.
SLOW_SYNC_SECONDS = 20

#: Every event in the push backlog is its own Graph round trip, so a backlog of this size means
#: the run is measured in tens of seconds before it has even started.
LARGE_PUSH_BACKLOG = 25

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


# --- all-day events -------------------------------------------------------------------
#
# An all-day event is a run of calendar dates, not a span of time. Graph stores it as
# midnight-to-midnight with an exclusive end (25 Dec is 25 Dec 00:00 to 26 Dec 00:00) and hands
# it back in whatever timezone the request preferred, which for this sync is UTC. Converting
# that midnight into the site's timezone, as a timed event is converted, moved Christmas to
# 05:30 on a Kolkata site and to 19:00 on Christmas Eve on a New York one. So the dates are
# read and written as they are, with no conversion in either direction.
#
# Frappe keeps the end inclusive: its calendar stores an all-day event as 00:00:00 on the first
# day to 23:59:59 on the last, and draws it from those dates.


def _graph_date(dt):
	"""The calendar date of a Graph dateTimeTimeZone dict, read as written (no conversion)."""
	from dateutil import parser

	if not dt or not dt.get("dateTime"):
		return None
	return parser.parse(dt["dateTime"]).date()


def _all_day_to_system(start, end):
	"""Graph's all-day ``start``/``end`` dicts as Frappe's (first day 00:00, last day 23:59:59)."""
	import datetime

	first = _graph_date(start)
	if not first:
		return None, None
	after = _graph_date(end)
	# Graph's end is the midnight after the last day. An end on or before the start would be an
	# event of no days at all, which is read as the one day it starts on.
	last = after - datetime.timedelta(days=1) if after and after > first else first
	return (
		datetime.datetime.combine(first, datetime.time.min),
		datetime.datetime.combine(last, datetime.time(23, 59, 59)),
	)


def _all_day_to_ms(starts_on, ends_on):
	"""A Frappe all-day Event's dates as Graph's midnight ``start`` and exclusive midnight ``end``.

	Graph refuses an all-day event whose start or end is not a midnight, so an end of 23:59:59 —
	what Frappe's own calendar stores — is turned into the midnight after that day. An end that
	is already a later midnight is taken as the exclusive end it looks like.
	"""
	import datetime

	first = get_datetime(starts_on).date()
	end = get_datetime(ends_on) if ends_on else None
	if end and end.time() == datetime.time.min and end.date() > first:
		after = end.date()
	elif end:
		after = end.date() + datetime.timedelta(days=1)
	else:
		after = first + datetime.timedelta(days=1)
	if after <= first:
		after = first + datetime.timedelta(days=1)
	tz = get_system_timezone()
	return (
		{"dateTime": f"{first.isoformat()}T00:00:00", "timeZone": tz},
		{"dateTime": f"{after.isoformat()}T00:00:00", "timeZone": tz},
	)


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
	"""Scheduled entry: queue a sync for every enabled+authorized Microsoft Calendar.

	This used to sync them one after another inside the scheduler's own job. The scheduler
	hands a cron hook the default queue's 300 second timeout and skips the tick entirely while
	the previous run is still in flight, so at roughly fifteen seconds a calendar the twentieth
	one is already past the budget and RQ kills the job part-way down the list. The list order
	is stable, which makes it the same calendars at the tail that are cut off every single time
	— they would never sync at all. A job per calendar gives each one its own timeout and lets
	the workers run them in parallel.

	Returns what it queued and what it left alone. Never raises: this is a scheduled hook, and
	redis being unreachable is not a reason to take the whole pass down.
	"""
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
	queued, skipped = [], []
	for name in names:
		if enqueue_sync(name)["queued"]:
			queued.append(name)
		else:
			# Already in flight from a previous tick, or the queue would not take it. Either
			# way this calendar is somebody else's problem for the next fifteen minutes.
			skipped.append(name)
	return {"queued": queued, "skipped": skipped}


def enqueue_sync(calendar_name):
	"""Queue a full sync for this calendar; returns {"queued": bool, "job_id": str}.

	The job id is per calendar and deduplicated, so "sync this one now" arriving five times —
	a scheduled tick plus somebody pressing the button four times — is one sync rather than
	five jobs racing each other for the same filelock.

	Deliberately not whitelisted: it takes a calendar name and no owner check, and the callers
	that face a browser (the Sync Now button) already do that check on the doc first.
	"""
	job_id = f"m365-sync-{calendar_name}"
	# Who hears about the result: the person who pressed the button, when a person pressed one.
	# The scheduled pass has nobody waiting on it, so its result is addressed to the calendar's
	# owner instead of to whichever account the scheduler happens to run as.
	notify_user = frappe.session.user if getattr(frappe.local, "request", None) else None

	try:
		if _sync_job_running(job_id):
			return {"queued": False, "job_id": job_id}
		frappe.enqueue(
			"frappe_microsoft365.microsoft_calendar_sync.run_sync_job",
			queue=SYNC_QUEUE,
			timeout=SYNC_JOB_TIMEOUT,
			job_id=job_id,
			deduplicate=True,
			calendar_name=calendar_name,
			notify_user=notify_user,
		)
	except Exception:
		# A queue that is full, a redis that is down, a site with no workers at all: none of
		# those are worth an exception in a scheduled pass or under somebody's button.
		frappe.log_error(title=f"MS Calendar sync could not be queued: {calendar_name}")
		return {"queued": False, "job_id": job_id}

	return {"queued": True, "job_id": job_id}


def _queue_or_do_it_after_commit(method, *, job_id, queue, timeout=None, **kwargs):
	"""Queue this, or — when nothing would ever run it — do it once the transaction commits.

	The commit matters as much as the fallback. These jobs exist because ``on_update`` fires
	*before* the row is written, so running the work inline right here would read the record as
	it was before the save and push that to Microsoft. ``enqueue_after_commit`` is what the
	queued path uses for exactly this reason, and the inline path has to honour it too.

	Silence is the failure this guards. ``frappe.enqueue`` succeeds perfectly well against a
	Redis with no workers behind it: the save returns, the queue grows, and Outlook quietly
	stops matching Frappe until somebody notices months of drift.
	"""
	if background.is_available():
		frappe.enqueue(
			method, queue=queue, timeout=timeout, job_id=job_id, deduplicate=True,
			enqueue_after_commit=True, **kwargs
		)
		return

	# Same call the worker would have made, just later in this request. Dropped automatically
	# if the transaction rolls back, because the callback registry is reset with it.
	frappe.db.after_commit.add(
		lambda: background.enqueue_or_run(method, job_id=job_id, queue=queue, timeout=timeout, **kwargs)
	)


def sync_blocked_by_dead_queue():
	"""Why a queued sync would never run, or None when the queue is fine.

	Only for the paths with a person waiting on the answer. The scheduled pass is already
	executing inside a worker when it queues, so the question is answered by the fact that it
	is running at all — and refusing there would turn a slow sync into no sync.

	A sync is also the one job that must never quietly fall back to running inline: it is
	queued precisely because it was judged too slow to hold a request open, so doing it here
	anyway would trade a silent failure for a certain gateway timeout. The caller reports this
	and offers to run it anyway, which is the person's decision to make, not ours.
	"""
	state = background.health()
	if state["ok"]:
		return None
	return {"reasons": state["reasons"], "fixes": state["fixes"], "message": state["message"]}


def _sync_job_running(job_id):
	"""True when this calendar's sync job is already queued or running.

	Its own function, and deliberately forgiving: a redis we cannot reach should answer "not
	running" and let ``enqueue_sync`` report the real failure, rather than raise from a check
	that only exists to avoid queueing a duplicate.
	"""
	try:
		from frappe.utils.background_jobs import is_job_enqueued

		return bool(is_job_enqueued(job_id))
	except Exception:
		return False


def run_sync_job(calendar_name, notify_user=None):
	"""Background entry: sync one calendar, then tell whoever is waiting that it finished."""
	try:
		result = sync_calendar(calendar_name)
	except Exception as e:
		frappe.log_error(title=f"MS Calendar sync job failed: {calendar_name}")
		result = {
			"ok": False,
			"pulled": 0,
			"deleted": 0,
			"pushed": 0,
			"message": _("Sync failed: {0}").format(str(e)),
		}

	_publish_sync_done(calendar_name, result, notify_user)
	return result


def _publish_sync_done(calendar_name, result, notify_user=None):
	"""Tell one person their queued sync is done, so an open form can stop saying "queued".

	Addressed to a user rather than broadcast: this says what happened to one person's mailbox,
	and every other browser on the site has no use for it. The recipient is whoever asked, and
	failing that the calendar's own owner — the only user with a reason to have that form open
	when the scheduler is what started the run.
	"""
	user = notify_user or frappe.db.get_value("Microsoft Calendar", calendar_name, "user")
	if not user:
		return

	try:
		frappe.publish_realtime(
			"microsoft365_sync_done",
			{
				"calendar": calendar_name,
				"pulled": result.get("pulled") or 0,
				"deleted": result.get("deleted") or 0,
				"pushed": result.get("pushed") or 0,
				"message": result.get("message") or "",
			},
			user=user,
		)
	except Exception:
		# The sync itself has already been done and committed. Failing to announce it is a
		# missing toast, not a failed sync, and must not be reported as one.
		frappe.log_error(title=f"MS Calendar sync result could not be published: {calendar_name}")


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
	started = time.monotonic()
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
		if doc.flags.pull_warning:
			messages.append(doc.flags.pull_warning)

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

	# The wall clock, because that is the question being asked: can somebody sit and watch the
	# next one of these. Almost all of it is Graph's latency rather than ours, so counting our
	# own work would answer a different question. Monotonic, so an NTP step mid-sync cannot
	# record a negative run. Written only when the column is there: the field arrived after the
	# rest of this doctype, and on a site that has the app but not yet the migration the
	# watermark still has to land — that part is not optional, a duration for a hint is.
	if frappe.get_meta("Microsoft Calendar").has_field("last_sync_seconds"):
		updates["last_sync_seconds"] = round(time.monotonic() - started, 2)

	frappe.db.set_value("Microsoft Calendar", calendar_name, updates, update_modified=False)
	# Committed here rather than left to the end of the job. The run above may have created
	# meetings in Outlook and stored their ids, and if this job dies after that point — a
	# timeout, a worker restart — an id that never committed makes the next run create the
	# meeting all over again. The watermark belongs with them.
	frappe.db.commit()  # nosemgrep

	summary = f"Pulled {pulled}, deleted {deleted}, pushed {pushed}."
	if doc.flags.orphans_flagged:
		summary += " " + _(
			"{0} events pulled from this calendar are no longer in Outlook and were marked Missing "
			"From Outlook; Microsoft Settings > Troubleshoot > Run Diagnostics lists them."
		).format(doc.flags.orphans_flagged)
	if doc.flags.pull_unfinished:
		# Not an error: a large calendar is read over several runs, and this says so.
		summary += " " + _("More changes are waiting in Microsoft; the next sync carries on from here.")

	return {
		"ok": not messages,
		"pulled": pulled,
		"deleted": deleted,
		"pushed": pushed,
		"message": "; ".join(messages) or summary,
	}


# --- inline, or in the background? ----------------------------------------------------

def background_reasons(doc):
	"""Human-readable reasons this sync is likely to be slow, empty if it should run inline.

	Reasons rather than a boolean because the caller puts them on screen: "this will take a
	while" is a far easier thing to accept when it says which of these is true. An empty list
	means an ordinary incremental run, which a person can sit and wait for.
	"""
	if isinstance(doc, str):
		doc = frappe.get_doc("Microsoft Calendar", doc)

	reasons = []
	if doc.pull_from_microsoft_calendar:
		if doc.get("delta_next_link"):
			reasons.append(_("The last sync stopped part-way through a large calendar and carries on from there."))
		elif not doc.delta_link:
			reasons.append(_("This is a first sync, so the whole calendar window is read."))
		elif _delta_window_is_stale(doc):
			reasons.append(_("The sync window is due to be renewed, so the calendar is read again in full."))

	if doc.push_to_microsoft_calendar:
		pending = _pending_push_count(doc)
		if pending >= LARGE_PUSH_BACKLOG:
			reasons.append(_("{0} events are waiting to be sent to Microsoft.").format(pending))

	# Measured rather than guessed. How long a sync takes is decided by the mailbox at the other
	# end — its size, and how hard Microsoft is throttling this tenant today — and the last run
	# is the only evidence anyone has about that.
	last_run = doc.get("last_sync_seconds")
	if last_run and last_run >= SLOW_SYNC_SECONDS:
		reasons.append(_("The last sync took {0} seconds.").format(int(last_run)))

	return reasons


def _pending_push_count(doc):
	"""How many Events the next push would send to Graph.

	The same filters as ``_push``, on purpose: an estimate that counts something other than what
	the push actually does is worse than having no estimate at all.
	"""
	if not doc.user:
		return 0
	base_filters = _push_filters(doc)
	pending = frappe.db.count("Event", {**base_filters, "custom_microsoft_event_id": ["is", "not set"]})
	if doc.last_sync:
		pending += frappe.db.count(
			"Event",
			{
				**base_filters,
				"custom_microsoft_event_id": ["is", "set"],
				"modified": [">", doc.last_sync],
			},
		)
	return pending


# --- pull (Graph -> Frappe) ----------------------------------------------------------

def _initial_delta_path(now=None):
	now = now or now_datetime()
	start = _iso_utc(add_to_date(now, days=-WINDOW_PAST_DAYS))
	end = _iso_utc(_window_end_from(now))
	return (
		f"/me/calendarView/delta?startDateTime={start}&endDateTime={end}"
		f"&$select={EVENT_SELECT}"
	)


def _window_end_from(now):
	"""Where a window initialised at ``now`` ends."""
	return add_to_date(now, days=WINDOW_FUTURE_DAYS + WINDOW_RESET_DAYS)


def _window_start_of(window_end):
	"""Where the window that ends at ``window_end`` started (its round's ``now`` minus the past)."""
	return add_to_date(get_datetime(window_end), days=-(WINDOW_PAST_DAYS + WINDOW_FUTURE_DAYS + WINDOW_RESET_DAYS))


def _window_end_is_stale(window_end):
	"""True once a delta window ending at ``window_end`` no longer covers the full look-ahead.

	That is the schedule: a window starts WINDOW_RESET_DAYS longer than the look-ahead, so it is
	renewed about once a month and an event up to WINDOW_FUTURE_DAYS ahead is always inside it.
	A window from before 1.2.0 was sized to the look-ahead alone, so it is renewed on its first run.
	"""
	if not window_end:
		return True
	return get_datetime(window_end) < add_to_date(now_datetime(), days=WINDOW_FUTURE_DAYS)


def _delta_window_is_stale(doc):
	"""True when the stored delta window is close to its end and must be re-initialised."""
	return _window_end_is_stale(doc.delta_window_end)


def _has_resume_point_fields():
	"""Whether this site has migrated the resume-point columns yet.

	The sync is scheduled, so it can run between the code update and `bench migrate`. Writing a
	column that does not exist yet would fail the whole pull; without it the pull simply behaves
	as it did before — each run reads as far as the page cap allows.
	"""
	return frappe.get_meta("Microsoft Calendar").has_field("delta_next_link")


def _start_of_pull(doc):
	"""Where this run's pull starts: ``(path, round_window_end)``.

	A *round* is one full delta read, from its first page to the page that carries the
	deltaLink. A large calendar can take several runs to get through one, so where it stopped
	is kept as a resume point and the next run carries on from there instead of starting over.

	``round_window_end`` is set when the round re-initialises the window, and is the end of the
	window it asked Graph for — stored as ``delta_window_end`` once the round completes. It is
	None for an incremental round, which keeps the window the stored deltaLink already has.
	"""
	resume = doc.get("delta_next_link")
	resume_window_end = doc.get("delta_next_window_end")
	if resume:
		if resume_window_end:
			# Carrying on with a first read. Only worth it while that window is still good.
			if not _window_end_is_stale(resume_window_end):
				return resume, get_datetime(resume_window_end)
		elif doc.delta_link and not _delta_window_is_stale(doc):
			return resume, None

	if doc.delta_link and not _delta_window_is_stale(doc):
		return doc.delta_link, None
	return _new_round()


def _new_round():
	# The window end doubles as the round's key (see _mark_seen). Datetime columns keep the
	# microseconds, so it round-trips exactly and two rounds a moment apart still differ.
	now = now_datetime()
	return _initial_delta_path(now), _window_end_from(now)


def _pull(doc):
	"""Pull changes from Microsoft into Frappe. Returns (upserted, deleted).

	Reads at most graph.MAX_PAGES pages per run. When that is not enough the round is left
	unfinished: the link to the next page is stored as the resume point and the next run
	carries on from it, so a large calendar catches up over a few runs instead of re-reading
	the same first pages forever. ``doc.flags.pull_unfinished`` tells the caller this happened.

	A round that reads the whole window (a first sync, a renewal, a resync after Graph refused
	the old link) is also the one moment the full list of what Outlook still has is known, so it
	ends by settling the mirrors it did not see: see _settle_unseen_mirrors.
	"""
	path, round_window_end = _start_of_pull(doc)
	# The watermark of the last completed round. A renewal reads the window afresh, and the
	# changes still waiting behind this link — deletions above all — would otherwise be lost.
	previous_link = doc.delta_link

	try:
		items, delta_link, next_link = graph.graph_delta(path, doc.name, headers=UTC_PREFER)
	except MsGraphResyncRequired:
		# Token expired or the mailbox was moved: start the window again from scratch. A resume
		# point belongs to the round Graph has just refused, so it goes with it.
		if path == previous_link:
			previous_link = None
		path, round_window_end = _new_round()
		items, delta_link, next_link = graph.graph_delta(path, doc.name, headers=UTC_PREFER)

	# One query per 500 events instead of one per event. Asking "do I already have this one?"
	# separately for every incoming event meant 500 round trips for a 500-event page, all of
	# them answering a question the database can answer in a single indexed read.
	existing_map = _existing_event_map(ev.get("id") for ev in items)

	frappe.flags.in_microsoft_sync = True
	upserted = removed = 0
	try:
		for ev in items:
			try:
				outcome = _upsert_event(doc, ev, existing_map)
				if outcome == "deleted":
					removed += 1
				elif outcome in ("created", "updated"):
					upserted += 1
			except Exception:
				frappe.log_error(title=f"MS event upsert failed: {ev.get('id')}")
	finally:
		frappe.flags.in_microsoft_sync = False

	if round_window_end:
		_mark_seen(items, round_window_end)

	resumable = _has_resume_point_fields()

	# Only store the new watermark once every page of the round was consumed without raising.
	# Until then the old deltaLink stays exactly as it was: a round that never finishes repeats
	# work, but never skips any.
	if delta_link:
		if round_window_end:
			# The whole window has just been read. First what the old link still had to say, then
			# whatever Outlook no longer has that nothing said was deleted.
			removed += _drain_previous_round(doc, previous_link)
			removed += _settle_unseen_mirrors(doc, round_window_end)
		updates = {"delta_link": delta_link}
		if round_window_end:
			updates["delta_window_end"] = round_window_end
		if resumable:
			updates.update({"delta_next_link": None, "delta_next_window_end": None})
		frappe.db.set_value("Microsoft Calendar", doc.name, updates, update_modified=False)
	elif next_link and resumable:
		if next_link == path:
			# Graph handed back the very page this run started from, so carrying on from it would
			# read the same pages forever. Said out loud rather than repeated in silence.
			raise MsGraphError(
				_("Microsoft returned the same page again, so the calendar read is not moving forward.")
			)
		frappe.db.set_value(
			"Microsoft Calendar",
			doc.name,
			{"delta_next_link": next_link, "delta_next_window_end": round_window_end},
			update_modified=False,
		)
		doc.flags.pull_unfinished = True
	elif next_link:
		# Not migrated yet: nowhere to keep the resume point, so a calendar this large cannot
		# move past its first pages until the site is migrated. Reported, not hidden.
		doc.flags.pull_warning = _(
			"The calendar has more changes than one sync reads. Run bench migrate so the sync can "
			"carry on where it stopped."
		)
	else:
		# Graph ended the read with neither a next page nor a deltaLink. Nothing is lost — the
		# next run reads the same changes again — but it can never move forward like this.
		raise MsGraphError(
			_("Microsoft ended the calendar read without a sync link, so the next run reads it again.")
		)

	return upserted, removed


def _has_event_field(fieldname):
	"""Whether this site's Event already has one of the app's newer custom fields.

	The sync is scheduled, so it can run between the code update and `bench migrate`, before the
	fields exist. What needs them is skipped until then rather than failing the whole pull.
	"""
	return frappe.get_meta("Event").has_field(fieldname)


def _mark_seen(items, round_key):
	"""Record that the mirrors of these events were seen in the full read keyed ``round_key``.

	One UPDATE per PREFETCH_CHUNK ids rather than a save per event: the mark is bookkeeping, it
	must not count as an edit, and a full read touches every mirror in the window. Seeing a
	mirror also clears its Missing From Outlook mark — it is there after all.
	"""
	if not (_has_event_field("custom_microsoft_seen_in_round") and _has_event_field("custom_microsoft_orphaned")):
		return
	ids = [ev.get("id") for ev in items if ev.get("id") and not _is_removed(ev)]
	ids = list(dict.fromkeys(ids))
	for start in range(0, len(ids), PREFETCH_CHUNK):
		chunk = ids[start : start + PREFETCH_CHUNK]
		frappe.db.sql(
			"""
			update `tabEvent`
			set custom_microsoft_seen_in_round = %(round)s, custom_microsoft_orphaned = 0
			where custom_pulled_from_microsoft = 1 and custom_microsoft_event_id in %(ids)s
			""",
			{"round": round_key, "ids": tuple(chunk)},
		)


def _drain_previous_round(doc, link):
	"""Apply the deletions still waiting behind the link a renewal is about to replace.

	Read only for its deletions: what still exists has just been read afresh with the new window.
	Best effort — the link may have expired, or have more pages than one run reads — because
	anything it cannot deliver is caught by _settle_unseen_mirrors straight afterwards, just
	flagged instead of deleted. Returns how many mirrors it deleted.
	"""
	if not link or not graph.is_graph_url(link):
		return 0
	try:
		items, _delta, _next = graph.graph_delta(link, doc.name, headers=UTC_PREFER)
	except MsGraphError:
		return 0

	gone = [ev for ev in items if _is_removed(ev)]
	existing_map = _existing_event_map(ev.get("id") for ev in gone)
	removed = 0
	frappe.flags.in_microsoft_sync = True
	try:
		for ev in gone:
			try:
				if _upsert_event(doc, ev, existing_map) == "deleted":
					removed += 1
			except Exception:
				frappe.log_error(title=f"MS event upsert failed: {ev.get('id')}")
	finally:
		frappe.flags.in_microsoft_sync = False
	return removed


def _unseen_mirrors(calendar_name, round_key):
	"""Mirrors in the window of the full read ``round_key`` that the read did not return."""
	return frappe.db.sql(
		"""
		select name from `tabEvent`
		where custom_microsoft_calendar = %(calendar)s
			and custom_pulled_from_microsoft = 1
			and custom_sync_with_microsoft_calendar = 1
			and coalesce(custom_microsoft_event_id, '') != ''
			and starts_on >= %(start)s and starts_on < %(end)s
			and (custom_microsoft_seen_in_round is null or custom_microsoft_seen_in_round != %(round)s)
		order by starts_on
		""",
		{
			"calendar": calendar_name,
			"start": _window_start_of(round_key),
			"end": round_key,
			"round": round_key,
		},
		pluck=True,
	)


def _settle_unseen_mirrors(doc, round_key):
	"""Mark the mirrors a full read of the window did not return as Missing From Outlook.

	A full read lists everything Outlook has in the window, so a mirror it did not return is one
	whose event was deleted without the deletion ever reaching us — in the gap a resync or a
	window renewal leaves. They are flagged, and listed by the doctor, rather than deleted:
	other records may link to these Events. Deleting them is the admin's choice, made with
	Microsoft Settings' "Delete events missing from Outlook". Returns how many were deleted.
	"""
	if not (_has_event_field("custom_microsoft_seen_in_round") and _has_event_field("custom_microsoft_orphaned")):
		return 0

	names = _unseen_mirrors(doc.name, round_key)
	if not names:
		return 0

	frappe.db.sql(
		"update `tabEvent` set custom_microsoft_orphaned = 1 where name in %(names)s",
		{"names": tuple(names)},
	)
	if not _delete_orphans():
		doc.flags.orphans_flagged = len(names)
		return 0

	deleted = 0
	frappe.flags.in_microsoft_sync = True
	try:
		for name in names:
			# A savepoint each, so one Event that cannot go takes nothing else with it.
			frappe.db.savepoint("m365_orphan")
			try:
				# Not forced: an Event another record still links to stays, flagged, rather than
				# leaving that record pointing at nothing.
				frappe.delete_doc("Event", name, ignore_permissions=True)
				deleted += 1
			except Exception:
				frappe.db.rollback(save_point="m365_orphan")
				frappe.log_error(title=f"MS Calendar: event missing from Outlook could not be deleted: {name}")
	finally:
		frappe.flags.in_microsoft_sync = False
	if deleted < len(names):
		doc.flags.orphans_flagged = len(names) - deleted
	return deleted


def _delete_orphans():
	"""The admin's choice: delete mirrors missing from Outlook, or only flag them (the default)."""
	if not frappe.get_meta("Microsoft Settings").has_field("delete_orphaned_calendar_events"):
		return False
	return bool(cint(frappe.db.get_single_value("Microsoft Settings", "delete_orphaned_calendar_events")))


def _is_removed(ev):
	return bool(ev.get("@removed")) or bool(ev.get("isCancelled"))


def _existing_event_map(ms_ids):
	"""Microsoft event id -> Frappe Event name, for the ids of one delta page.

	Chunked rather than one enormous ``IN``: see PREFETCH_CHUNK. Ids that have no Frappe Event
	are simply absent from the map, which is the same answer the per-event lookup gave.
	"""
	mapping = {}
	# dict.fromkeys and not set(): a delta page is normally in a meaningful order, and keeping
	# it makes the chunk boundaries reproducible when this has to be debugged against a log.
	ids = [ms_id for ms_id in dict.fromkeys(ms_ids) if ms_id]

	for start in range(0, len(ids), PREFETCH_CHUNK):
		rows = frappe.get_all(
			"Event",
			filters={"custom_microsoft_event_id": ["in", ids[start : start + PREFETCH_CHUNK]]},
			fields=["name", "custom_microsoft_event_id"],
			# Spelled out rather than relying on get_all's default, because a page capped at 20
			# rows would not fail: the events past the cap would read as new and be created a
			# second time. This is the one argument here that must not be wrong.
			limit_page_length=0,
		)
		for row in rows:
			mapping[row.custom_microsoft_event_id] = row.name

	return mapping


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


def _ms_times(ev):
	"""A Graph event's start and end as Frappe datetimes: floating dates when it is all day."""
	if ev.get("isAllDay"):
		return _all_day_to_system(ev.get("start"), ev.get("end"))
	return _ms_dt_to_system(ev.get("start")), _ms_dt_to_system(ev.get("end"))


def _target_values(doc, ev, locally_originated):
	"""The Frappe Event field values a Microsoft event maps to."""
	values = {
		"subject": ev.get("subject") or "(No subject)",
		"all_day": 1 if ev.get("isAllDay") else 0,
		"custom_microsoft_calendar": doc.name,
		"custom_sync_with_microsoft_calendar": 1,
	}
	starts_on, ends_on = _ms_times(ev)
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


def _upsert_event(doc, ev, existing_map=None):
	"""Create/update/delete the Frappe mirror of one Microsoft event.

	``existing_map`` is the pull's prefetched id -> Event name map. Called without one — the
	single-event paths, and the tests — the original per-event lookup still answers, because an
	optimisation that only works when it is handed the right argument is a trap for the next
	caller rather than a speed-up.

	Returns "deleted", "updated", "created" or "skipped".
	"""
	ms_id = ev.get("id")
	if not ms_id:
		return "skipped"

	if existing_map is None:
		existing = frappe.db.get_value("Event", {"custom_microsoft_event_id": ms_id}, "name")
	else:
		existing = existing_map.get(ms_id)

	if _is_removed(ev):
		if existing:
			frappe.delete_doc("Event", existing, ignore_permissions=True, force=True)
			# The map is this page's picture of what exists, so it has to follow what the page
			# does to it. One delta page can carry the same id twice — an event changed twice
			# between runs — and a stale entry would then have the second copy update an Event
			# that has just been deleted.
			if existing_map is not None:
				existing_map.pop(ms_id, None)
			return "deleted"
		return "skipped"

	# calendarView returns occurrences; a seriesMaster would duplicate all of them.
	if ev.get("type") == "seriesMaster":
		return "skipped"

	# An occurrence of a series this app pushed is already represented in Frappe — by the
	# repeating Event that created it. Mirroring it back would put a separate Event beside the
	# repeating one for every occurrence, forever, and the two models do not even disagree:
	# Frappe keeps one record with a repeat rule, exactly as Graph keeps one series master.
	#
	# calendarView never returns that master, so this is the only place the relationship is
	# visible at all.
	series_id = ev.get("seriesMasterId")
	if series_id and _is_ours(series_id):
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
	if existing and event.get("custom_microsoft_orphaned"):
		# Outlook has just said something about it, so it is not missing after all.
		values["custom_microsoft_orphaned"] = 0
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

	if existing:
		return "updated"

	# Same reason as the delete above: a second copy of this id later in the page must update
	# the Event this call just created instead of creating another one beside it.
	if existing_map is not None:
		existing_map[ms_id] = event.name
	_give_it_to_its_owner(event.name, doc.user)
	return "created"


def _give_it_to_its_owner(event_name, user):
	"""A pulled event belongs to the person whose calendar it came from.

	Frappe stamps owner with whoever is running, and set_user_and_timestamp does it
	unconditionally on insert, so it cannot simply be assigned beforehand. The scheduled pass
	runs as Administrator, so every event pulled on a schedule was owned by Administrator —
	while the same event pulled by somebody pressing Sync Now was owned by them.

	The ownership is not cosmetic. These are created Private, and Frappe shows a private event
	to its owner, to people it is shared with, and to its participants — so an event owned by
	Administrator is invisible to the one person whose calendar it came out of. On a
	single-admin site nothing looks wrong; on the second connection the feature quietly stops
	working for everybody but the admin.
	"""
	if not user:
		return
	frappe.db.set_value("Event", event_name, "owner", user, update_modified=False)


def _is_ours(series_master_id):
	"""Did this app push the series behind this occurrence?

	Matched on the id Graph gave back when the series was created, which is stored on the
	repeating Event — so this is a lookup, not a guess.
	"""
	return bool(
		frappe.db.exists("Event", {"custom_microsoft_event_id": series_master_id})
	)


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

def _push_filters(doc):
	"""The Events this calendar sends to Microsoft: its own user's, and nobody else's.

	The owner is part of the filter because linking an Event to a calendar proves nothing about
	who may write to it, and the scheduled push runs as Administrator — so without it, an Event
	anyone linked to this calendar would be sent from this person's mailbox.
	"""
	return {
		"custom_sync_with_microsoft_calendar": 1,
		"custom_microsoft_calendar": doc.name,
		"custom_pulled_from_microsoft": 0,
		"owner": doc.user,
	}


def _push(doc):
	"""Create new Frappe events in Microsoft and patch ones edited while offline."""
	if not doc.user:
		# A connection with no user has no owner whose Events it could send.
		return 0
	base_filters = _push_filters(doc)

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


#: Frappe's repeat_on -> Graph's recurrence pattern. Quarterly and Half Yearly have no Graph
#: equivalent and are expressed as an absoluteMonthly pattern with a wider interval, which is
#: what they are.
REPEAT_PATTERNS = {
	"Daily": ("daily", 1),
	"Weekly": ("weekly", 1),
	"Monthly": ("absoluteMonthly", 1),
	"Quarterly": ("absoluteMonthly", 3),
	"Half Yearly": ("absoluteMonthly", 6),
	"Yearly": ("absoluteYearly", 1),
}

WEEKDAYS = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday")


def _recurrence(event):
	"""A Frappe repeating Event as a Graph recurrence, or None when it does not repeat.

	Graph wants the pattern (how often) and the range (until when) separately, and it validates
	them against the event's own start — a weekly series whose daysOfWeek does not include the
	start day is rejected outright, so the start day is always included.
	"""
	if not getattr(event, "repeat_this_event", 0):
		return None

	mapped = REPEAT_PATTERNS.get(getattr(event, "repeat_on", None))
	if not mapped:
		# An unmapped frequency is better sent as a single event than as a wrong series: a
		# wrong recurrence writes itself across somebody's real calendar for months.
		return None

	kind, interval = mapped
	starts_on = get_datetime(event.starts_on)
	pattern = {"type": kind, "interval": interval}

	if kind == "weekly":
		days = [day for day in WEEKDAYS if getattr(event, day, 0)]
		# Graph rejects a weekly series that does not repeat on its own start day.
		start_day = WEEKDAYS[starts_on.weekday()]
		if start_day not in days:
			days.append(start_day)
		pattern["daysOfWeek"] = days
	elif kind == "absoluteMonthly":
		pattern["dayOfMonth"] = starts_on.day
	elif kind == "absoluteYearly":
		pattern["dayOfMonth"] = starts_on.day
		pattern["month"] = starts_on.month

	range_ = {"type": "noEnd", "startDate": starts_on.strftime("%Y-%m-%d")}
	if getattr(event, "repeat_till", None):
		range_ = {
			"type": "endDate",
			"startDate": starts_on.strftime("%Y-%m-%d"),
			"endDate": get_datetime(event.repeat_till).strftime("%Y-%m-%d"),
		}
	range_["recurrenceTimeZone"] = get_system_timezone()

	return {"pattern": pattern, "range": range_}


def _event_to_graph_body(event):
	starts_on = get_datetime(event.starts_on)
	body = {
		"subject": event.subject or "(No subject)",
		"body": {"contentType": "HTML", "content": event.description or ""},
	}
	if getattr(event, "all_day", 0):
		body["start"], body["end"] = _all_day_to_ms(starts_on, event.ends_on)
		body["isAllDay"] = True
	else:
		body["start"] = _system_dt_to_ms(starts_on)
		body["end"] = _system_dt_to_ms(_end_after_start(starts_on, event.ends_on))
	if getattr(event, "location", None):
		body["location"] = {"displayName": event.location}

	# Sent only while the event does not yet have a meeting — never again afterwards.
	#
	# Microsoft treats isOnlineMeeting on a PATCH as "make one", not "keep one": it mints a
	# fresh Teams meeting and throws the old one away, with a new join URL and a new online
	# meeting id. Re-asserting it on every save therefore did two silent kinds of damage —
	# everybody holding the invitation's join link was left with a dead one, and the meeting
	# that was actually held, with its transcript and its recording, was orphaned where nothing
	# could find it again. Observed on a real meeting: three saves, three different meetings,
	# and a recording sitting in Teams that Graph could not match to any of them.
	#
	# The condition is the join URL rather than the tickbox, because that is the thing that
	# answers "does a meeting already exist" — true on create, true when somebody ticks the box
	# on an event Outlook already knows about, false on every ordinary edit thereafter.
	# Microsoft does not support turning an online meeting back into a plain event either, so
	# there is nothing to send in the other direction.
	if getattr(event, "custom_add_teams_meeting", 0) and not getattr(event, "custom_teams_join_url", None):
		body["isOnlineMeeting"] = True
		body["onlineMeetingProvider"] = "teamsForBusiness"

	# Only sent when at least one participant resolved to an address. Graph reads an empty
	# attendees array as "remove everyone", and "this Event has no participants in Frappe"
	# is indistinguishable from "its attendees were added in Outlook", so sending [] would
	# quietly uninvite people nobody asked to uninvite.
	attendees = _graph_attendees(event)
	if attendees:
		body["attendees"] = attendees

	# A repeating Frappe Event used to reach Outlook as a single one-off: the series was simply
	# dropped, with nothing to say so. Sent only when it maps to something Graph understands.
	recurrence = _recurrence(event)
	if recurrence:
		body["recurrence"] = recurrence

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


# --- whose calendar is it? -------------------------------------------------------------

def _calendar_user(calendar_name):
	"""The Frappe user a Microsoft Calendar belongs to, or None.

	Read from the database rather than the document cache: this answers a permission question,
	and a cached copy can lag a change of user by the rest of the request.
	"""
	if not calendar_name:
		return None
	return frappe.db.get_value("Microsoft Calendar", calendar_name, "user")


def _event_owner(doc):
	"""Who the Event belongs to, as stored — not whatever owner a client sent with the save.

	On insert Frappe stamps the owner itself, so the value on the document is trustworthy. On
	an update the document carries whatever the request sent, so the stored owner is used.
	"""
	before = doc.get_doc_before_save() if hasattr(doc, "get_doc_before_save") else None
	return (before.owner if before else None) or doc.owner


def _is_owners_calendar(owner, calendar_name):
	"""True when ``calendar_name`` is the Microsoft Calendar of ``owner``.

	An Event is only ever written to its owner's own mailbox. There is no exception for System
	Managers: they can repair, re-authorise or disconnect any connection, but sending an Event
	from somebody else's mailbox — with invitations going out in that person's name — is not
	something any role is allowed to do.
	"""
	user = _calendar_user(calendar_name)
	return bool(owner) and bool(user) and owner == user


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

	_validate_calendar_owner(doc)

	if not doc.starts_on or not doc.ends_on:
		return

	if get_datetime(doc.ends_on) <= get_datetime(doc.starts_on):
		frappe.throw(
			_("This event ends before it starts, and Microsoft will not accept it. Set an end time after {0}.").format(
				frappe.utils.format_datetime(doc.starts_on)
			),
			title=_("Check the times"),
		)


def _validate_calendar_owner(doc):
	"""Refuse to link an Event to a Microsoft calendar that is not its owner's.

	Frappe's Link check only proves the calendar exists, not that the person saving may write to
	it; and Microsoft Calendar names are whatever their owners typed, so they are easy to guess.
	Without this, anybody could have an Event created in a colleague's Outlook — invitations and
	all, in the colleague's name.

	Mirrors pulled from Microsoft are left alone: saving one never sends anything, and they
	belong to the calendar's own user already.
	"""
	calendar_name = getattr(doc, "custom_microsoft_calendar", None)
	if not calendar_name or getattr(doc, "custom_pulled_from_microsoft", 0):
		return
	if _is_owners_calendar(_event_owner(doc), calendar_name):
		return
	frappe.throw(
		_(
			"The Microsoft Calendar {0} does not belong to the owner of this event. An event can "
			"only be synced with its owner's own Microsoft calendar."
		).format(frappe.bold(calendar_name)),
		frappe.PermissionError,
		title=_("Not your Microsoft Calendar"),
	)


def event_on_update(doc, method=None):
	"""Create in Graph inside the save; queue the patch for an event that already exists.

	A creation is the one Graph call somebody is genuinely waiting for: the id and the Teams
	join link come back in its response and are written onto the document the browser is about
	to be handed, so doing it later would mean saving a Teams meeting whose link is not there
	yet. A patch answers with nothing anyone can see, and paying 300-900 milliseconds of Graph
	latency for it on every save — up to ten seconds when Microsoft throttles and the retry
	sleeps — is a cost with no buyer.
	"""
	if frappe.flags.in_microsoft_sync:
		return
	if not getattr(doc, "custom_sync_with_microsoft_calendar", 0):
		return
	if not getattr(doc, "custom_microsoft_calendar", None):
		return
	if getattr(doc, "custom_pulled_from_microsoft", 0):
		return
	# Checked again after validate: other code can save an Event with ignore_validate, and this
	# is the step that actually writes to someone's mailbox.
	if not _is_owners_calendar(_event_owner(doc), doc.custom_microsoft_calendar):
		return

	try:
		cal = frappe.get_cached_doc("Microsoft Calendar", doc.custom_microsoft_calendar)
		if not cal.enabled or not cal.authorized or not cal.push_to_microsoft_calendar:
			return

		if doc.custom_microsoft_event_id:
			_queue_or_do_it_after_commit(
				"frappe_microsoft365.microsoft_calendar_sync.patch_event_in_graph",
				queue="short",
				# Ten rapid saves are ten requests, each queueing after its own commit, so
				# without an id per Event they would be ten patches sending increasingly stale
				# bodies. One job per Event, and it reads the record when it runs.
				job_id=f"m365-event-patch-{doc.name}",
				# on_update runs before this transaction commits. A worker that started now
				# could read the row as it was before the save — or, on a brand new Event, not
				# find it at all — and would then patch Outlook back to the old text.
				event_name=doc.name,
			)
		else:
			_store_graph_response(doc.name, _create_graph_event(cal.name, doc), cal.name, live_doc=doc)
	except Exception:
		frappe.log_error(title=f"MS Event on_update sync failed: {doc.name}")


def patch_event_in_graph(event_name):
	"""Send a Frappe Event's current state to Graph. Runs in a background job.

	Everything is read here rather than carried in the job's arguments. Between the save that
	queued this and a worker picking it up the Event may have been edited again, moved to
	another calendar, unticked, or deleted outright — so this is an instruction to sync the
	record as it now stands, never a snapshot of how it looked when somebody pressed Ctrl+S.
	"""
	if not frappe.db.exists("Event", event_name):
		return

	doc = frappe.get_doc("Event", event_name)
	if not doc.custom_sync_with_microsoft_calendar or not doc.custom_microsoft_calendar:
		return
	if doc.custom_pulled_from_microsoft or not doc.custom_microsoft_event_id:
		return
	# The job runs later and as whoever the worker runs as, so the save's own check proves
	# nothing by now: the Event or the calendar may have changed hands in between.
	if not _is_owners_calendar(doc.owner, doc.custom_microsoft_calendar):
		return

	# The guard on the doc_events side reads a flag that lives in the request that set it, and
	# a worker has no such request: it starts with the flag unset. Setting it here keeps the
	# guard's meaning true inside the job — "Frappe is talking to Microsoft right now, do not
	# bounce anything back" — so no write this push makes can queue another patch behind it.
	previous = frappe.flags.in_microsoft_sync
	frappe.flags.in_microsoft_sync = True
	try:
		cal = frappe.get_cached_doc("Microsoft Calendar", doc.custom_microsoft_calendar)
		if not cal.enabled or not cal.authorized or not cal.push_to_microsoft_calendar:
			return
		patched = graph.graph_request(
			"PATCH",
			f"/me/events/{doc.custom_microsoft_event_id}",
			cal.name,
			json=_event_to_graph_body(doc),
		)
		_store_graph_response(doc.name, patched or {}, cal.name)
	except Exception:
		frappe.log_error(title=f"MS Event patch failed: {event_name}")
	finally:
		frappe.flags.in_microsoft_sync = previous


def event_on_trash(doc, method=None):
	"""Queue the removal of the mirrored Microsoft event. Best-effort, never blocks the delete."""
	if frappe.flags.in_microsoft_sync:
		return
	ms_id = getattr(doc, "custom_microsoft_event_id", None)
	cal_name = getattr(doc, "custom_microsoft_calendar", None)
	if not ms_id or not cal_name:
		return

	try:
		# Deleting in Microsoft is a write like any other, so it follows the push switch too.
		# Without this a pull-only connection — "bring Outlook in, never touch it" — removed the
		# Outlook event, and cancelled the meeting for everyone when its owner was the organiser,
		# whenever someone deleted the mirror in Frappe.
		if not frappe.db.get_value("Microsoft Calendar", cal_name, "push_to_microsoft_calendar"):
			return
		# Only the owner's own Events are removed from their own calendar, and only when the
		# owner is the one deleting. A Public Event can be deleted by anyone with Event access,
		# and that must not reach into its owner's mailbox and cancel the meeting there; nor may
		# a System Manager's delete, for the same reason. The Frappe record still goes — the
		# Outlook event is simply left for its owner.
		if frappe.session.user != doc.owner or not _is_owners_calendar(doc.owner, cal_name):
			return
		_queue_or_do_it_after_commit(
			"frappe_microsoft365.microsoft_calendar_sync.delete_event_in_graph",
			queue="short",
			# Keyed on the Microsoft id, not the Event name: the Event is what is going away,
			# and the id is the only half of this pair that still means something afterwards.
			job_id=f"m365-event-delete-{ms_id}",
			# A delete that is rolled back must not have already removed the meeting from
			# somebody's Outlook, so this job only exists once the deletion is real.
			calendar_name=cal_name,
			ms_event_id=ms_id,
			event_owner=doc.owner,
		)
	except Exception:
		frappe.log_error(title=f"MS Event on_trash sync failed: {doc.name}")


def delete_event_in_graph(calendar_name, ms_event_id, event_owner=None):
	"""Remove an event from Microsoft after its Frappe Event has been deleted.

	Takes two plain ids and no document reference, because by the time a worker runs this the
	Frappe Event is gone: there is nothing left to load, and a job that tried would find
	nothing and quietly leave the meeting sitting in the calendar forever.

	What can still be re-checked is re-checked. The connection may have been disconnected in
	the meantime, and some other Event may have come to carry this same Microsoft id — deleting
	it then would strand that mirror, and the next pull would delete it locally too, taking a
	live meeting off both sides.

	``event_owner`` is checked against the calendar's user once more, in case the connection
	changed hands since the delete. A job with no owner — queued by an older version, before
	this check existed — cannot be verified and does nothing: an Outlook event left behind can
	be deleted by hand, a meeting cancelled in the wrong mailbox cannot be undone.
	"""
	if not calendar_name or not ms_event_id:
		return
	if not _is_owners_calendar(event_owner, calendar_name):
		return

	try:
		if frappe.db.exists("Event", {"custom_microsoft_event_id": ms_event_id}):
			return
		cal = frappe.get_cached_doc("Microsoft Calendar", calendar_name)
		# Checked again here: push may have been switched off between the delete and this job.
		if not cal.enabled or not cal.authorized or not cal.push_to_microsoft_calendar:
			return
		graph.graph_request("DELETE", f"/me/events/{ms_event_id}", cal.name)
	except Exception:
		frappe.log_error(title=f"MS Event delete failed: {ms_event_id}")


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
		start, end = _ms_times(ev)
		out.append(
			{
				"id": ev.get("id"),
				"subject": ev.get("subject"),
				"start": start,
				"end": end,
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


# --- upgrade checks (read-only, run after migrate) -----------------------------------

#: How many Event names each upgrade finding lists; the count covers the rest.
UPGRADE_SAMPLE_SIZE = 10


def upgrade_check_event_owners():
	"""Events linked to a Microsoft Calendar that belongs to somebody other than their owner.

	Since 1.1.1 an Event is only pushed, patched or deleted in its owner's own Microsoft
	calendar, so these stop syncing and cannot be saved with sync ticked until they are
	re-linked. Nothing is changed here: it reads, counts per calendar and names a few of them.
	"""
	from frappe_microsoft365 import doctor

	if not frappe.db.table_exists("Microsoft Calendar") or not frappe.db.has_column(
		"Event", "custom_microsoft_calendar"
	):
		return []

	rows = _events_of_others()

	by_calendar = {}
	for row in rows:
		by_calendar.setdefault(row.calendar, {"user": row.calendar_user, "events": []})["events"].append(row)

	findings = []
	for calendar, group in by_calendar.items():
		events = group["events"]
		sample = ", ".join(
			f"{row.name} ({row.owner or _('no owner')})" for row in events[:UPGRADE_SAMPLE_SIZE]
		)
		more = len(events) - UPGRADE_SAMPLE_SIZE
		if more > 0:
			sample += " " + _("and {0} more").format(more)
		findings.append(
			doctor.finding(
				"upgrade.calendar_event_owner",
				doctor.WARN,
				_("{0} events on Microsoft Calendar {1} belong to someone else").format(len(events), calendar),
				_(
					"This connection belongs to {0}, but these events are owned by other users: {1}. "
					"Events are now only synced with their owner's own Microsoft calendar, so these "
					"no longer reach Outlook, are not deleted there, and cannot be saved with sync "
					"ticked. Nothing was changed."
				).format(group["user"] or _("no user"), sample),
				_(
					"If they were made for {0} (by an administrator, say), open Microsoft Calendar {1} "
					"and choose Microsoft > Give Linked Events to {0}. Otherwise have each event's owner "
					"link it to their own Microsoft Calendar, or untick Sync with Microsoft Calendar on it."
				).format(group["user"] or _("its user"), calendar),
				target=calendar,
			)
		)
	return findings


def orphaned_event_findings(check="calendar.orphaned_events"):
	"""Doctor findings: Events pulled from Outlook that a full read no longer found, per calendar.

	Read-only. The sync only flags these (see _settle_unseen_mirrors); this is where an admin
	finds them, decides, and learns about the setting that deletes them instead.
	"""
	from frappe_microsoft365 import doctor

	if not frappe.db.has_column("Event", "custom_microsoft_orphaned"):
		return []

	rows = frappe.get_all(
		"Event",
		filters={"custom_microsoft_orphaned": 1},
		fields=["name", "subject", "starts_on", "custom_microsoft_calendar as calendar"],
		order_by="custom_microsoft_calendar, starts_on",
		limit_page_length=0,
	)
	by_calendar = {}
	for row in rows:
		by_calendar.setdefault(row.calendar or _("no calendar"), []).append(row)

	findings = []
	for calendar, events in by_calendar.items():
		sample = ", ".join(f"{row.name} ({row.subject})" for row in events[:UPGRADE_SAMPLE_SIZE])
		more = len(events) - UPGRADE_SAMPLE_SIZE
		if more > 0:
			sample += " " + _("and {0} more").format(more)
		findings.append(
			doctor.finding(
				check,
				doctor.WARN,
				_("{0} events pulled from Microsoft Calendar {1} are no longer in Outlook").format(len(events), calendar),
				_(
					"A full read of the Outlook calendar did not find these, so they were most likely "
					"deleted there while the sync was resetting: {0}. They are marked Missing From "
					"Outlook and were not deleted, because other records may link to them."
				).format(sample),
				_(
					"Check them in Outlook and delete the ones that are gone. To have the sync delete "
					"such events itself, tick Delete events missing from Outlook in Microsoft Settings."
				),
				target=calendar,
			)
		)
	return findings


def upgrade_check_orphaned_events():
	"""Upgrade report: Events pulled from Outlook that are flagged Missing From Outlook.

	Events only get flagged by a full read of their calendar, which the first sync after
	upgrading to 1.2.0 does, so right after `bench migrate` this normally has nothing to say;
	Run Diagnostics shows the same list from then on.
	"""
	return orphaned_event_findings("upgrade.calendar_orphaned_events")


def _events_of_others(calendar=None):
	"""Events linked to a Microsoft Calendar whose owner is not that calendar's user.

	Only Events that would ever reach Microsoft: ticked for sync, or already carrying an id —
	deleting one of those is a write to the mailbox even when sync was later unticked.
	"""
	return frappe.db.sql(
		f"""
		select e.name, e.owner, e.custom_microsoft_calendar as calendar, c.`user` as calendar_user
		from `tabEvent` e
		inner join `tabMicrosoft Calendar` c on c.name = e.custom_microsoft_calendar
		where (e.custom_sync_with_microsoft_calendar = 1
				or coalesce(e.custom_microsoft_event_id, '') != '')
			and (coalesce(c.`user`, '') = '' or e.owner != c.`user`)
			{"and c.name = %(calendar)s" if calendar else ""}
		order by e.custom_microsoft_calendar, e.creation
		""",
		{"calendar": calendar},
		as_dict=True,
	)


@frappe.whitelist(methods=["POST"])
def give_linked_events_to_calendar_user(calendar_name: str, apply: int = 0):
	"""Make the calendar's user the owner of the Events linked to it that someone else owns.

	Events an administrator created for a colleague's calendar stopped syncing in 1.1.1, when
	only an Event's owner's own calendar could be written to, and Frappe offers no way to change
	an owner from the form. This is that way, for System Managers, one calendar at a time. With
	``apply`` unset it only counts, so the form can say what is about to change. Ownership decides
	who sees a private Event, so each one gets a comment naming its previous owner.
	"""
	frappe.only_for("System Manager")
	user = frappe.db.get_value("Microsoft Calendar", calendar_name, "user")
	if not user:
		frappe.throw(_("Microsoft Calendar {0} has no user to give events to.").format(calendar_name))
	events = _events_of_others(calendar_name)
	if not cint(apply):
		return {"user": user, "count": len(events)}
	for row in events:
		frappe.db.set_value("Event", row.name, "owner", user, update_modified=False)
		frappe.get_doc("Event", row.name).add_comment(
			"Info",
			_("Owner changed from {0} to {1} so the event syncs with {1}'s Microsoft calendar.").format(
				row.owner or _("no owner"), user
			),
		)
	return {"user": user, "count": len(events), "changed": [row.name for row in events]}

