"""Will work handed to ``frappe.enqueue`` actually be done — and if not, what fixes it.

Slow work belongs in a background job. The trouble is that handing work to one succeeds in
three situations that look identical from the inside and are not:

* **The scheduler is off.** The cron hooks never fire, so the periodic passes simply never
  happen. Nothing raises, nothing is logged, and nobody finds out until someone notices a
  calendar that has not moved in a week.
* **Redis is unreachable.** ``frappe.enqueue`` raises. Loud, at least — but it reaches the
  person as a traceback rather than as a problem with a fix next to it.
* **Redis is up and no worker is running.** ``frappe.enqueue`` returns happily and the job
  sits in the queue forever. This is the dangerous one: every appearance of success and none
  of the work. The bench this was written on had 203 jobs waiting on nobody.

So nothing here enqueues blind. ``enqueue_or_run`` asks first, and runs the work in the
request when the queue cannot be trusted — the same trade core makes in
``frappe/desk/doctype/bulk_update``, which submits inline rather than queue while the
scheduler is inactive — and hands back the reasons so the caller can *also* say what is
broken. Doing the work and swallowing the fault would only move the next failure further
away from its cause, which is how a site ends up with 203 dead jobs in the first place.

This module deliberately does not decide how the person is told. Two callers already need to
say it differently — the doctor lists it beside the other setup faults, a meeting's status
line has to fold it into one honest sentence — so ``health`` returns the raw reasons and
fixes and lets each of them phrase it.

``health`` never raises, at all, under any circumstance. It opens a Redis connection that may
not open and reads a database that may not answer, and a check that takes down the page it
exists to protect is worse than no check.
"""

import frappe
from frappe import _
from frappe.utils import get_datetime, now_datetime
from frappe.utils.background_jobs import (
	get_job,
	get_queue,
	get_queue_list,
	get_redis_conn,
	get_workers,
)
from frappe.utils.scheduler import is_scheduler_inactive
from rq.job import JobStatus

#: Where the expensive half of the probe is kept.
#:
#: ``frappe.cache()`` prefixes every key with the site's database name, so two sites sharing
#: one bench — and one Redis — never read each other's answer. That matters here more than
#: usual: the scheduler is a per-site setting, so the verdict genuinely differs per site even
#: though the workers being counted are shared.
CACHE_KEY = "microsoft365:background_probe"

#: Seconds the live-worker count is trusted for.
#:
#: Sixty is a compromise between two costs that pull opposite ways. Counting workers takes
#: ~155 ms, which cannot be paid on every form load; a stale "yes there are workers" makes
#: this module queue into a void for at most that long after the last worker dies. Nothing
#: else in the probe is cached, so the answers a person changes by hand — enabling the
#: scheduler, in particular — take effect the moment they change them rather than a minute
#: later, which is exactly when someone is watching to see whether their fix worked.
CACHE_TTL = 60

#: This app's own scheduled jobs, as ``Scheduled Job Type.method`` stores them.
JOB_METHOD_LIKE = "%frappe_microsoft365%"

#: Minutes without a run after which one of our scheduled jobs is treated as not running.
#:
#: Both jobs are registered as ``*/15 * * * *`` (see hooks.py), so an hour is four missed
#: turns in a row. One missed turn explains itself — the scheduler only wakes every
#: ``DEFAULT_SCHEDULER_TICK`` (4 minutes), a busy queue delays the job further, and a slow run
#: pushes the next one out. Four in a row does not: past that, the honest reading is that
#: nothing is running the job at all. An hour also clears rq's 7-minute worker heartbeat TTL
#: comfortably, so a worker restart cannot trip it.
STALE_AFTER_MINUTES = 60

#: Jobs already waiting before the backlog is worth mentioning. Not a fault on its own — a
#: worker that is merely behind will still get there — so it never changes the verdict.
BACKLOG_NOTABLE = 50

#: Stable identifiers for the ways this goes wrong. The reason and the fix beside them are
#: translated sentences; these are not, so a caller can single out one cause without matching
#: on English that a site running in German will never produce.
SCHEDULER_OFF = "scheduler_off"
REDIS_DOWN = "redis_down"
NO_WORKER = "no_worker"
WORKERS_UNKNOWN = "workers_unknown"
JOBS_MISSING = "jobs_missing"
JOBS_STALE = "jobs_stale"
CHECK_FAILED = "check_failed"

_UNSET = object()


def _problem(code, reason, fix):
	"""One thing that is wrong, the sentence for it, and the command that ends it."""
	return {"code": code, "reason": reason, "fix": fix}


# --- the two cache layers --------------------------------------------------------------


def _request_store():
	"""Per-request memo for the expensive probe.

	Kept on ``frappe.local`` rather than in a module global because a worker process serves
	one site after another out of the same import; ``frappe.local`` is reset between them and
	a module global is not, which is how a multi-tenant bench starts answering for the wrong
	site. It also survives Redis being down, which the cache layer below cannot — and Redis
	being down is precisely one of the cases this module has to stay cheap in.
	"""
	store = getattr(frappe.local, "microsoft365_background", None)
	if store is None:
		store = {}
		frappe.local.microsoft365_background = store
	return store


def forget():
	"""Drop everything remembered about the queue, both layers. Never raises.

	``health(refresh=True)`` is the normal way to get an uncached answer; this exists for the
	moment after someone has actually fixed the bench and nothing should survive.
	"""
	_request_store().pop("workers", None)
	try:
		frappe.cache().delete_value(CACHE_KEY)
	except Exception:
		pass


def _count_workers():
	"""Live workers on this bench, or None if Redis could not be asked.

	The only expensive call in this module — ~155 ms against a warm local Redis, because rq's
	``Worker.all()`` verifies each registered worker key individually instead of trusting the
	set it is listed in. That per-key check is the whole point: a worker killed mid-job stays
	in the set until its heartbeat key expires seven minutes later, so counting set members
	would report a dead worker as alive for exactly the seven minutes it matters most.

	Counted across the bench rather than per queue: Frappe's standard worker consumes short,
	default and long together, so a per-queue count would report zero for a queue that is in
	fact being served and send people chasing a worker they already have.
	"""
	try:
		return len(get_workers())
	except Exception:
		return None


def _live_workers(refresh=False):
	"""``_count_workers`` behind the request memo and the site cache."""
	store = _request_store()
	if not refresh and "workers" in store:
		return store["workers"]

	value = _UNSET
	if not refresh:
		try:
			cached = frappe.cache().get_value(CACHE_KEY, expires=True)
		except Exception:
			cached = None
		# A dict, not the bare number: None is a real answer here ("could not ask"), and a
		# cache miss also reads back as None.
		if isinstance(cached, dict) and "workers" in cached:
			value = cached["workers"]

	if value is _UNSET:
		value = _count_workers()
		try:
			frappe.cache().set_value(CACHE_KEY, {"workers": value}, expires_in_sec=CACHE_TTL)
		except Exception:
			pass

	store["workers"] = value
	return value


# --- the individual checks -------------------------------------------------------------


def _site():
	return getattr(frappe.local, "site", None) or "<site>"


def _scheduler_fix():
	"""The command that actually turns it back on, which depends on *how* it is off.

	``is_scheduler_inactive`` collapses three different settings into one boolean, and telling
	someone to run ``enable-scheduler`` when the real cause is ``maintenance_mode`` sends them
	round a loop: core refuses to re-enable the scheduler while maintenance mode is on.
	"""
	site = _site()
	conf = getattr(frappe.local, "conf", None) or {}
	if conf.get("maintenance_mode"):
		return _("The site is in maintenance mode. Run `bench --site {0} set-maintenance-mode off`.").format(
			site
		)
	if conf.get("pause_scheduler"):
		return _("The scheduler is paused. Run `bench --site {0} scheduler resume`.").format(site)
	return _("Run `bench --site {0} enable-scheduler`.").format(site)


def _redis_reachable():
	"""Can the queue's Redis be talked to at all? ``get_redis_conn`` raises when it cannot."""
	try:
		get_redis_conn().ping()
		return True, ""
	except Exception as e:
		# str(e), never the exception object: passing one to .format() is blocked by semgrep,
		# and a ConnectionError's own repr is the useful half anyway.
		return False, str(e)


def _queue_depths():
	"""Jobs waiting per queue. Bench-wide, not per site — the queue is shared and counting
	only this site's jobs would mean reading every job body, which is not a cheap check."""
	depths = {}
	try:
		for name in get_queue_list():
			depths[name] = get_queue(name).count
	except Exception:
		pass
	return depths


def _job_rows():
	"""This app's rows in Scheduled Job Type, or None if they could not be read.

	``last_execution`` is read back in Python rather than filtered on in SQL: it is NULL for a
	job that has never run, and a never-run job is the single most interesting row here.
	``stopped`` jobs are left out — one that somebody deliberately switched off is not broken.
	"""
	try:
		return frappe.get_all(
			"Scheduled Job Type",
			filters={"method": ["like", JOB_METHOD_LIKE], "stopped": 0},
			fields=["name", "method", "last_execution", "creation"],
		)
	except Exception:
		return None


def job_status(rows, now=None, dormant=False):
	"""Each row with a staleness verdict. Pure — rows in, verdicts out — so the threshold can
	be argued with in a test rather than against a live scheduler."""
	now = now or now_datetime()
	jobs = []
	for row in rows:
		last = row.get("last_execution")
		minutes = int((now - get_datetime(last or row.get("creation"))).total_seconds() // 60)
		jobs.append(
			{
				"name": row.get("name"),
				"method": row.get("method"),
				"last_execution": last,
				"minutes_ago": minutes if last else None,
				# Only a job that ran and then stopped is proof of anything. "Never run" cannot
				# be dated: Scheduled Job Type.creation records when the row was first synced,
				# not when the scheduler was last switched on, so a site installed months ago
				# and enabled this morning looks identical to one that has been broken since
				# install. Measuring from creation called the first of those broken — right
				# after the admin did the thing we asked them to do, which is the fastest way
				# to teach someone to stop believing a diagnostic.
				"stale": bool(last) and (not dormant) and minutes > STALE_AFTER_MINUTES,
				"never_ran": not last,
			}
		)
	return jobs


def _scheduled_jobs():
	rows = _job_rows()
	return None if rows is None else job_status(rows, dormant=_dormant())


def _dormant():
	"""Frappe Cloud throttles a dormant site's scheduled jobs to once a day on purpose.

	Calling an hour-old run stale there would be a false alarm about a feature working as
	designed. Off Frappe Cloud this returns False immediately, so it costs nothing.
	"""
	try:
		from frappe.utils.scheduler import is_dormant

		return bool(is_dormant())
	except Exception:
		return False


# --- the verdict -----------------------------------------------------------------------


def health(refresh: bool = False) -> dict:
	"""See ``_health``. This wrapper exists only to make "never raises" absolute.

	Guarding each call individually would hold only until the next person adds a check and
	forgets, and the cost of being wrong is a traceback on a page whose whole purpose was to
	avoid one. A probe that cannot complete reports that background work is not running: that
	costs an inline run and never costs the work.
	"""
	try:
		return _health(refresh)
	except Exception as e:
		reason = _("Whether background jobs run here could not be established: {0}").format(str(e))
		fix = _("Run `bench doctor` to see the queue from the bench's side.")
		return {
			"ok": False,
			"problems": [_problem(CHECK_FAILED, reason, fix)],
			"reasons": [reason],
			"fixes": [fix],
			"scheduler_inactive": None,
			"redis": None,
			"workers": None,
			"queued": 0,
			"queues": {},
			"jobs": [],
			"backlog": False,
			"checked_on": None,
			"message": reason + " " + fix,
		}


def _health(refresh: bool = False) -> dict:
	"""Whether background work runs on this bench, why not, and what fixes it.

	Returns ``{"ok": bool, "reasons": [str], "fixes": [str], ...}`` plus the raw readings the
	doctor prints: ``scheduler_inactive``, ``redis``, ``workers``, ``queued``, ``queues``,
	``jobs``. ``reasons`` are sentences a person reads; ``fixes`` are the commands that end
	the problem; ``problems`` pairs each with a stable ``code`` so a caller can pick out one
	cause without matching on a sentence that a translated site will not produce.

	``refresh=True`` bypasses both cache layers. The doctor and any Run-now button want the
	truth, not a minute-old opinion — someone who just started a worker is about to press the
	button precisely to find out whether it took.

	Every problem flips ``ok``, including a scheduled job that has gone quiet while everything
	else looks fine. That last one is the conservative direction on purpose: it costs a caller
	an inline run it might not have needed, and the alternative costs the work entirely.
	"""
	problems = []

	scheduler_inactive = _scheduler_inactive()
	if scheduler_inactive:
		# An explicitly enqueued job would in principle still be picked up by a worker that is
		# running regardless of the scheduler. We follow core's rule anyway (bulk_update.py
		# runs inline on exactly this condition) because the scheduler and the workers are
		# started and stopped together — by `bench start`, by one supervisor group — so a
		# disabled scheduler is the loudest cheap evidence that nothing is consuming the queue.
		problems.append(
			_problem(
				SCHEDULER_OFF,
				_("The scheduler is off for this site, so nothing is queued on a schedule."),
				_scheduler_fix(),
			)
		)

	redis_ok, redis_error = _redis_reachable()
	workers, queues = None, {}
	if not redis_ok:
		problems.append(
			_problem(
				REDIS_DOWN,
				_("Redis, which holds the job queue, cannot be reached: {0}").format(redis_error),
				_("Start Redis and check `redis_queue` in common_site_config.json, then `bench doctor`."),
			)
		)
	else:
		workers = _live_workers(refresh)
		queues = _queue_depths()
		if workers == 0:
			# The silent one. Queuing succeeds, the job is accepted, and it is never started.
			problems.append(
				_problem(
					NO_WORKER,
					_("No background worker is running, so queued jobs are accepted and then never start."),
					_(
						"Start one with `bench worker --queue default`, or in production check "
						"`supervisorctl status` for the worker processes."
					),
				)
			)
		elif workers is None:
			problems.append(
				_problem(
					WORKERS_UNKNOWN,
					_("Redis answered, but how many workers are running could not be read."),
					_("Run `bench doctor` to see the queue from the bench's side."),
				)
			)

	jobs = _scheduled_jobs()
	if jobs is not None:
		problems += _job_problems(jobs, scheduler_inactive)

	queued = sum(queues.values())
	state = {
		"ok": not problems,
		"problems": problems,
		"reasons": [p["reason"] for p in problems],
		"fixes": [p["fix"] for p in problems],
		"scheduler_inactive": scheduler_inactive,
		"redis": redis_ok,
		"workers": workers,
		"queued": queued,
		"queues": queues,
		"jobs": jobs or [],
		"backlog": queued >= BACKLOG_NOTABLE,
		"checked_on": now_datetime(),
	}
	state["message"] = _message(state)
	return state


def _scheduler_inactive():
	try:
		# verbose=False: core's default prints to the terminal, which would put scheduler
		# chatter into the middle of every bench command that happens to load a meeting.
		return bool(is_scheduler_inactive(verbose=False))
	except Exception:
		# Unreadable settings are not evidence that the scheduler is off, and guessing "off"
		# here would push every caller into running inline on a healthy bench.
		return False


def _job_problems(jobs, scheduler_inactive):
	"""What this app's own scheduled jobs say about the bench.

	The most honest signal available, because it measures the thing itself rather than its
	preconditions: a job that should run four times an hour and last ran yesterday is not
	running, whatever the scheduler setting, Redis and the worker count claim.
	"""
	if not jobs:
		return [
			_problem(
				JOBS_MISSING,
				_("This app's scheduled jobs are not registered on this site."),
				_("Run `bench --site {0} migrate` to register them.").format(_site()),
			)
		]

	stale = [job for job in jobs if job["stale"]]
	# A stale job with the scheduler off is the same fault told twice; the scheduler already
	# said it, and said it with the command that fixes it.
	if not stale or scheduler_inactive:
		return []

	return [
		_problem(
			JOBS_STALE,
			_("{0} scheduled job(s) of this app last ran {1} minutes ago, not the expected 15.").format(
				len(stale), max(job["minutes_ago"] for job in stale)
			),
			_(
				"The scheduler is enabled but its process may not be running. Check `bench doctor`, "
				"and `supervisorctl status` in production."
			),
		)
	]


def never_ran_note(jobs):
	"""A job that has never run, for the doctor to mention without calling the site broken.

	Kept apart from the problems above on purpose: this one cannot tell a freshly enabled
	scheduler from a permanently broken one, so it is worth saying and not worth acting on.
	"""
	never = [job for job in jobs or [] if job.get("never_ran")]
	if not never:
		return None
	return _(
		"{0} scheduled job(s) of this app have not run yet. If the scheduler was only just "
		"enabled this is expected within 15 minutes; if it persists, check `bench doctor`."
	).format(len(never))


def _message(state):
	"""One sentence, for a caller that has room for exactly one."""
	if state["ok"]:
		return _("Background jobs are running.")
	return _("Background jobs are not running. {0} {1}").format(
		" ".join(state["reasons"]), " ".join(state["fixes"])
	)


def is_available() -> bool:
	"""Will a job handed to ``frappe.enqueue`` right now actually be executed?"""
	return health()["ok"]


# --- doing the work either way ---------------------------------------------------------


def enqueue_or_run(method, *, job_id, queue="default", timeout=None, **kwargs) -> dict:
	"""Queue the work when the queue works; do it here and now when it does not.

	Returns ``{"queued", "ran_inline", "health", "result"}``. Exactly one of the first two is
	ever True: a queue that cannot be trusted is not merely supplemented by an inline run, it
	is bypassed, and any job already sitting under this ``job_id`` is dropped first so the
	work cannot be done twice once someone eventually starts a worker.

	``job_id`` is required and has no default. Every caller here is a repeatable operation —
	sync this calendar, chase this meeting's transcript — and without an id two clicks queue
	the same Graph traffic twice; with one, ``deduplicate`` collapses them.

	Errors from the work itself are not caught. The point of the fallback is that the work
	happens *and* the fault is visible; swallowing the work's own failure on top of that
	would leave the caller unable to tell the two apart.
	"""
	state = health()

	if state["ok"]:
		try:
			frappe.enqueue(method, queue=queue, timeout=timeout, job_id=job_id, deduplicate=True, **kwargs)
			return {"queued": True, "ran_inline": False, "health": state, "result": None}
		except Exception:
			# Redis can die between the probe and this line, and the worker count may have been
			# a minute old. Re-ask without the cache: if the queue really is fine, this is our
			# bug — a bad queue name, an unimportable method — and hiding it behind an inline
			# run would make it invisible forever.
			state = health(refresh=True)
			if state["ok"]:
				raise

	# A worker already doing this exact work is the one case where running inline is the wrong
	# answer: health() can be wrong in this direction (the worker count is up to a minute old,
	# and a scheduler that is off does not stop a worker someone started by hand), and two
	# concurrent runs of the same sync is worse than one late one.
	if _is_running(job_id):
		return {"queued": True, "ran_inline": False, "health": state, "result": None}

	_drop_queued(job_id)
	# now=True routes through frappe.call, which is exactly what the worker would have done
	# with this method and these kwargs — same resolution of a dotted path, same dropping of
	# arguments the method does not take — so inline and queued cannot drift apart.
	#
	# It never touches Redis, so it still works when Redis is the thing that is broken — but
	# only as called here. Passing job_id or deduplicate would send frappe.enqueue through
	# get_job() BEFORE it checks `now`, which reaches for Redis and defeats the entire
	# fallback. Do not add them to this call.
	result = frappe.enqueue(method, now=True, **kwargs)
	return {"queued": False, "ran_inline": True, "health": state, "result": result}


def _is_running(job_id):
	"""True when a worker has this job in hand right now.

	Deliberately fails closed on an unreachable queue: if we cannot tell, we assume nothing is
	running, because the alternative is refusing to do the work at all on a bench whose Redis
	is down — which is exactly when the inline path is the only one left.
	"""
	try:
		job = get_job(job_id)
		return bool(job) and job.get_status(refresh=False) == JobStatus.STARTED
	except Exception:
		return False


def _drop_queued(job_id):
	"""Delete a job of this id that is still waiting, before running the same work inline.

	Without this, the fallback is not a fallback: the stale job outlives the fix, and the
	first worker anyone starts does a day's worth of queued work all over again.
	"""
	try:
		job = get_job(job_id)
		if job and job.get_status(refresh=False) == JobStatus.QUEUED:
			job.delete()
	except Exception:
		# Best effort by nature — if Redis is unreachable there is no queue to clean, and a
		# failure here must not stop the work from being done.
		pass
