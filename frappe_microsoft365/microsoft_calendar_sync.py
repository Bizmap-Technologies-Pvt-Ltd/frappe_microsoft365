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
	"attendees,organizer,isOrganizer,responseStatus,responseRequested"
)

#: How much of the calendar the delta window covers. Graph fixes the window when the delta
#: is initialised, so we re-initialise before the far edge gets close.
WINDOW_PAST_DAYS = 30
WINDOW_FUTURE_DAYS = 180
WINDOW_REFRESH_MARGIN_DAYS = 14

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

	return {
		"ok": not messages,
		"pulled": pulled,
		"deleted": deleted,
		"pushed": pushed,
		"message": "; ".join(messages) or f"Pulled {pulled}, deleted {deleted}, pushed {pushed}.",
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
		if not doc.delta_link:
			reasons.append(_("This is a first sync, so the whole calendar window is read."))
		elif _delta_window_is_stale(doc):
			reasons.append(_("The sync window has run out, so the calendar is read again in full."))

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
	base_filters = {
		"custom_sync_with_microsoft_calendar": 1,
		"custom_microsoft_calendar": doc.name,
		"custom_pulled_from_microsoft": 0,
	}
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
		)
	except Exception:
		frappe.log_error(title=f"MS Event on_trash sync failed: {doc.name}")


def delete_event_in_graph(calendar_name, ms_event_id):
	"""Remove an event from Microsoft after its Frappe Event has been deleted.

	Takes two plain ids and no document reference, because by the time a worker runs this the
	Frappe Event is gone: there is nothing left to load, and a job that tried would find
	nothing and quietly leave the meeting sitting in the calendar forever.

	What can still be re-checked is re-checked. The connection may have been disconnected in
	the meantime, and some other Event may have come to carry this same Microsoft id — deleting
	it then would strand that mirror, and the next pull would delete it locally too, taking a
	live meeting off both sides.
	"""
	if not calendar_name or not ms_event_id:
		return

	try:
		if frappe.db.exists("Event", {"custom_microsoft_event_id": ms_event_id}):
			return
		cal = frappe.get_cached_doc("Microsoft Calendar", calendar_name)
		if not cal.enabled or not cal.authorized:
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
