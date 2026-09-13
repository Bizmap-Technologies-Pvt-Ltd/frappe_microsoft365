"""Background-queue health tests.

No worker is started and no job is really queued: every probe is pinned. That is the point —
the bench this was written on has the scheduler off, no worker and 203 stranded jobs, while CI
has a live scheduler and workers. A test that read the real bench would pass on one and fail
on the other while saying nothing about the code.

The one thing that is real is the inline fallback: it reaches ``record_run`` by dotted path
through ``frappe.call``, exactly as a worker would have, so "it ran inline" means the work
actually happened rather than that a flag was set.
"""

from contextlib import ExitStack
from unittest.mock import Mock, patch

import frappe
from frappe.utils import add_to_date, now_datetime
from rq.job import JobStatus

from frappe_microsoft365 import background
from frappe_microsoft365.tests.base import BaseTestCase

#: Appended to by ``record_run``. The only evidence that inline means inline.
RUNS = []

METHOD = "frappe_microsoft365.tests.test_background.record_run"

SYNC = "frappe_microsoft365.microsoft_calendar_sync.sync_all"


def record_run(marker=None):
	RUNS.append(marker)
	return marker


def job_row(method=SYNC, minutes_ago=5, registered_minutes_ago=24 * 60):
	"""One Scheduled Job Type row. ``minutes_ago=None`` is a job that has never run."""
	return {
		"name": method,
		"method": method,
		"last_execution": None if minutes_ago is None else add_to_date(now_datetime(), minutes=-minutes_ago),
		"creation": add_to_date(now_datetime(), minutes=-registered_minutes_ago),
	}


def codes(state):
	return {problem["code"] for problem in state["problems"]}


class BackgroundTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		RUNS.clear()
		# Both layers, or the previous test's pinned worker count answers this one.
		background.forget()
		self.addCleanup(background.forget)

	def pins(self, *, scheduler_inactive=False, redis=True, workers=2, queued=0, jobs=None):
		"""Patchers for every probe, in the order ``health`` calls them.

		``workers=None`` means the count could not be read at all, which is a different answer
		from zero and has to stay one.
		"""
		return [
			patch.object(background, "is_scheduler_inactive", return_value=scheduler_inactive),
			patch.object(background, "get_queue_list", return_value=["default"]),
			patch.object(background, "get_queue", return_value=Mock(count=queued)),
			patch.object(background, "_job_rows", return_value=[job_row()] if jobs is None else jobs),
			# Frappe Cloud's dormant-site throttle is a real exemption, tested on its own below.
			patch.object(background, "_dormant", return_value=False),
			patch.object(
				background,
				"get_redis_conn",
				**(
					{"return_value": Mock()}
					if redis
					else {"side_effect": ConnectionError("Error 61 connecting to localhost:11000.")}
				),
			),
			patch.object(
				background,
				"get_workers",
				**(
					{"return_value": [Mock()] * workers}
					if workers is not None
					else {"side_effect": ConnectionError("connection lost")}
				),
			),
		]

	def probe(self, **kwargs):
		"""Pin every probe for the rest of the test. Returns the ``get_workers`` mock, which the
		caching tests count calls on."""
		started = [self._start(pin) for pin in self.pins(**kwargs)]
		return started[-1]

	def _start(self, pin):
		mock = pin.start()
		self.addCleanup(pin.stop)
		return mock

	def site_conf(self, **overrides):
		"""The real site config with a value or two changed, rather than a bare dict: half of
		Frappe reads db_name out of it, including the cache key this module writes."""
		conf = frappe._dict(frappe.local.conf)
		conf.update(overrides)
		return patch.object(frappe.local, "conf", conf)


class TestHealth(BackgroundTestCase):
	def test_a_bench_that_works_is_available(self):
		self.probe()

		state = background.health()

		self.assertTrue(state["ok"])
		self.assertEqual(state["reasons"], [])
		self.assertTrue(background.is_available())

	def test_a_disabled_scheduler_names_the_command_that_enables_it(self):
		"""Silent from the outside: the cron hooks stop firing and nothing is logged."""
		self.probe(scheduler_inactive=True)

		state = background.health()

		self.assertFalse(state["ok"])
		self.assertIn(background.SCHEDULER_OFF, codes(state))
		self.assertIn("scheduler", state["reasons"][0].lower())
		self.assertIn("enable-scheduler", state["fixes"][0])
		self.assertIn(frappe.local.site, state["fixes"][0])

	def test_maintenance_mode_is_not_fixed_by_enable_scheduler(self):
		"""Core refuses to re-enable the scheduler while maintenance mode is on, so naming
		enable-scheduler there sends the admin round a loop that cannot close."""
		self.probe(scheduler_inactive=True)

		with self.site_conf(maintenance_mode=1):
			fix = background.health()["fixes"][0]

		self.assertIn("set-maintenance-mode off", fix)
		self.assertNotIn("enable-scheduler", fix)

	def test_a_paused_scheduler_is_resumed_not_enabled(self):
		self.probe(scheduler_inactive=True)

		with self.site_conf(pause_scheduler=1):
			fix = background.health()["fixes"][0]

		self.assertIn("scheduler resume", fix)

	def test_redis_being_unreachable_is_reported_rather_than_raised(self):
		"""get_redis_conn raises when Redis is gone. A health check that takes down the page it
		exists to protect is worse than no check at all."""
		self.probe(redis=False)

		state = background.health()

		self.assertFalse(state["ok"])
		self.assertIn(background.REDIS_DOWN, codes(state))
		self.assertIn("Redis", state["reasons"][0])
		self.assertIn("Error 61", state["reasons"][0])

	def test_redis_up_with_no_worker_is_caught(self):
		"""The dangerous one: frappe.enqueue succeeds, the job is accepted, nothing starts it,
		and every other signal on the bench says fine."""
		self.probe(workers=0)

		state = background.health()

		self.assertFalse(state["ok"])
		self.assertIn(background.NO_WORKER, codes(state))
		self.assertTrue(state["redis"])
		self.assertIn("bench worker --queue default", state["fixes"][0])
		self.assertIn("supervisorctl", state["fixes"][0])

	def test_a_worker_count_that_cannot_be_read_is_admitted_to(self):
		"""Not the same as zero, and pretending it is would be a guess in the loud direction."""
		self.probe(workers=None)

		state = background.health()

		self.assertIn(background.WORKERS_UNKNOWN, codes(state))
		self.assertIsNone(state["workers"])

	def test_a_backlog_is_reported_but_is_not_itself_a_fault(self):
		"""A worker that is merely behind still gets there; forcing inline runs on a busy queue
		would turn a slow bench into a stalled one."""
		self.probe(queued=203)

		state = background.health()

		self.assertTrue(state["ok"])
		self.assertTrue(state["backlog"])
		self.assertEqual(state["queued"], 203)

	def test_health_returns_a_verdict_even_when_every_probe_explodes(self):
		self.probe()
		for pin in (
			patch.object(background, "is_scheduler_inactive", side_effect=RuntimeError("no settings")),
			patch.object(background, "get_redis_conn", side_effect=RuntimeError("no redis")),
			patch.object(background, "_job_rows", side_effect=RuntimeError("no table")),
		):
			self._start(pin)

		state = background.health()

		self.assertFalse(state["ok"])
		self.assertTrue(state["message"])
		self.assertFalse(background.is_available())

	def test_a_probe_that_cannot_complete_at_all_still_answers(self):
		"""The clock itself needs the database in Frappe, so even now_datetime() can throw. The
		promise is absolute: a caller is told background work is not running, and pays one
		inline run rather than a traceback on the page this was meant to protect."""
		with patch.object(background, "_health", side_effect=RuntimeError("object is not bound")):
			state = background.health()

		self.assertFalse(state["ok"])
		self.assertEqual(codes(state), {background.CHECK_FAILED})
		self.assertIn("object is not bound", state["reasons"][0])
		self.assertIsNone(state["workers"])
		self.assertIsNone(state["scheduler_inactive"])


class TestScheduledJobStaleness(BackgroundTestCase):
	"""The signal that survives when scheduler, Redis and worker count all look fine.

	A `bench schedule` process that is not actually running is invisible to all three: workers
	drain the queue happily, and nothing ever puts the cron jobs into it.
	"""

	def test_a_job_that_ran_minutes_ago_is_fine(self):
		self.probe(jobs=[job_row(minutes_ago=5)])

		self.assertTrue(background.health()["ok"])

	def test_a_job_that_has_not_run_for_hours_is_stale(self):
		self.probe(jobs=[job_row(minutes_ago=240)])

		state = background.health()

		self.assertFalse(state["ok"])
		self.assertIn(background.JOBS_STALE, codes(state))
		self.assertIn("240", state["reasons"][0])

	def test_a_job_that_has_never_run_is_not_called_broken(self):
		"""A never-run job cannot be dated, so it must not be treated as proof of a fault.

		Scheduled Job Type.creation records when the row was first synced, not when the
		scheduler was last switched on. A site installed months ago and enabled this morning is
		indistinguishable from one broken since install — and calling the first of those broken
		means crying wolf at the exact moment the admin did what we told them to.
		"""
		self.probe(jobs=[job_row(minutes_ago=None, registered_minutes_ago=180)])

		state = background.health()

		self.assertTrue(state["ok"], "an undateable signal must not flip the verdict")
		self.assertNotIn(background.JOBS_STALE, codes(state))

	def test_a_job_that_has_never_run_is_still_worth_mentioning(self):
		"""Not a fault, but not nothing either — the doctor says it without acting on it."""
		jobs = background.job_status([job_row(minutes_ago=None, registered_minutes_ago=180)])

		note = background.never_ran_note(jobs)

		self.assertTrue(note)
		self.assertIn("not run yet", note)

	def test_a_job_that_ran_and_then_stopped_is_unambiguous(self):
		"""The other half of the same rule: this one IS dateable, so it does flip the verdict."""
		self.probe(jobs=[job_row(minutes_ago=240)])

		state = background.health()

		self.assertFalse(state["ok"])
		self.assertIn(background.JOBS_STALE, codes(state))

	def test_a_job_registered_minutes_ago_has_not_had_its_chance_yet(self):
		"""Right after a migrate every row has last_execution NULL, and shouting then would
		train people to ignore this check."""
		self.probe(jobs=[job_row(minutes_ago=None, registered_minutes_ago=5)])

		self.assertTrue(background.health()["ok"])

	def test_no_registered_jobs_at_all_points_at_migrate(self):
		self.probe(jobs=[])

		state = background.health()

		self.assertIn(background.JOBS_MISSING, codes(state))
		self.assertIn("migrate", state["fixes"][0])

	def test_unreadable_job_rows_are_not_treated_as_missing_jobs(self):
		"""A half-migrated site cannot answer the question, which is not the same as 'none'."""
		self.probe()

		with patch.object(background, "_job_rows", return_value=None):
			state = background.health()

		self.assertNotIn(background.JOBS_MISSING, codes(state))
		self.assertTrue(state["ok"])

	def test_a_dead_scheduler_does_not_also_report_its_own_symptom(self):
		"""One cause, one sentence: the stale job IS the disabled scheduler, and the scheduler
		line already carries the command that ends it."""
		self.probe(scheduler_inactive=True, jobs=[job_row(minutes_ago=600)])

		self.assertEqual(codes(background.health()), {background.SCHEDULER_OFF})

	def test_the_threshold_is_four_missed_quarter_hours(self):
		"""Both jobs are registered */15. One miss is explained by the four-minute scheduler
		tick and a busy queue; four in a row is not."""
		self.assertEqual(background.STALE_AFTER_MINUTES, 60)

		fresh, stale = background.job_status([job_row(minutes_ago=59), job_row(minutes_ago=61)])

		self.assertFalse(fresh["stale"])
		self.assertTrue(stale["stale"])

	def test_a_dormant_site_is_not_accused(self):
		"""Frappe Cloud throttles a dormant site's scheduled jobs to once a day on purpose."""
		jobs = background.job_status([job_row(minutes_ago=600)], dormant=True)

		self.assertFalse(jobs[0]["stale"])


class TestCaching(BackgroundTestCase):
	def test_the_expensive_worker_count_is_asked_for_once(self):
		"""~155 ms a call, because rq verifies each registered worker key individually rather
		than trusting the set. Paying that on every form load is not an option."""
		workers = self.probe()

		background.health()
		background.health()
		background.health()

		self.assertEqual(workers.call_count, 1)

	def test_refresh_goes_and_looks_again(self):
		"""The doctor and any Run-now button need the truth: someone presses them precisely to
		find out whether the worker they just started took."""
		workers = self.probe()

		background.health(refresh=True)
		background.health(refresh=True)

		self.assertEqual(workers.call_count, 2)

	def test_the_cheap_checks_are_never_cached(self):
		"""Someone who has just run enable-scheduler must not be told it is still off for
		another minute — that is exactly the window in which they conclude the fix failed."""
		self.probe(scheduler_inactive=True)
		self.assertFalse(background.health()["ok"])

		background.is_scheduler_inactive.return_value = False

		self.assertTrue(background.health()["ok"])

	def test_the_cached_answer_belongs_to_one_site(self):
		"""Workers are shared across a bench but the scheduler is a per-site setting, so a
		bench-wide key would hand one site's verdict to its neighbour."""
		key = frappe.cache().make_key(background.CACHE_KEY)

		self.assertIn(frappe.conf.db_name.encode(), key)

	def test_forget_drops_both_layers(self):
		workers = self.probe()
		background.health()

		background.forget()
		background.health()

		self.assertEqual(workers.call_count, 2)


class TestEnqueueOrRun(BackgroundTestCase):
	def test_a_working_queue_gets_the_job_and_nothing_runs_here(self):
		self.probe()

		with patch.object(frappe, "enqueue") as enqueue:
			result = background.enqueue_or_run(METHOD, job_id="m365-test", marker="a")

		self.assertTrue(result["queued"])
		self.assertFalse(result["ran_inline"])
		self.assertEqual(RUNS, [])
		self.assertEqual(enqueue.call_args.kwargs["job_id"], "m365-test")
		self.assertTrue(enqueue.call_args.kwargs["deduplicate"])
		self.assertNotIn("now", enqueue.call_args.kwargs)

	def test_a_dead_scheduler_does_the_work_here_and_says_so(self):
		self.probe(scheduler_inactive=True)

		result = background.enqueue_or_run(METHOD, job_id="m365-test", marker="a")

		self.assertTrue(result["ran_inline"])
		self.assertFalse(result["queued"])
		self.assertEqual(RUNS, ["a"])
		self.assertEqual(result["result"], "a")
		self.assertFalse(result["health"]["ok"])
		self.assertIn("enable-scheduler", result["health"]["fixes"][0])

	def test_no_worker_also_does_the_work_here(self):
		"""The case that used to lose the work entirely: queued, accepted, never started."""
		self.probe(workers=0)

		result = background.enqueue_or_run(METHOD, job_id="m365-test", marker="b")

		self.assertTrue(result["ran_inline"])
		self.assertEqual(RUNS, ["b"])

	def test_it_never_both_queues_and_runs(self):
		"""Two runs of the same Graph traffic is worse than none: it doubles the throttling and
		re-attaches every transcript."""
		for label, kwargs in (
			("healthy", {}),
			("scheduler off", {"scheduler_inactive": True}),
			("no worker", {"workers": 0}),
			("redis down", {"redis": False}),
			("stale jobs", {"jobs": [job_row(minutes_ago=600)]}),
		):
			with self.subTest(label), ExitStack() as stack:
				RUNS.clear()
				background.forget()
				for pin in self.pins(**kwargs):
					stack.enter_context(pin)

				pushed = []

				def spy(method, **kw):
					pushed.append(kw)
					if kw.get("now"):
						return frappe.call(method, **{k: v for k, v in kw.items() if k != "now"})

				stack.enter_context(patch.object(frappe, "enqueue", side_effect=spy))
				result = background.enqueue_or_run(METHOD, job_id="m365-test", marker=label)

				self.assertNotEqual(result["queued"], result["ran_inline"])
				self.assertEqual(RUNS, [label] if result["ran_inline"] else [])
				self.assertEqual(len(pushed), 1)

	def test_the_inline_fallback_drops_a_job_already_waiting_under_the_same_id(self):
		"""Otherwise the fallback is not a fallback: the stranded job outlives the fix, and the
		first worker anyone starts does the whole thing a second time."""
		self.probe(scheduler_inactive=True)
		stranded = Mock()
		stranded.get_status.return_value = JobStatus.QUEUED

		with patch.object(background, "get_job", return_value=stranded):
			background.enqueue_or_run(METHOD, job_id="m365-test", marker="c")

		stranded.delete.assert_called_once()
		self.assertEqual(RUNS, ["c"])

	def test_a_job_that_already_finished_is_left_alone(self):
		self.probe(scheduler_inactive=True)
		done = Mock()
		done.get_status.return_value = JobStatus.FINISHED

		with patch.object(background, "get_job", return_value=done):
			background.enqueue_or_run(METHOD, job_id="m365-test")

		done.delete.assert_not_called()

	def test_a_redis_failure_at_the_push_still_gets_the_work_done(self):
		"""The worker count can be a minute old, so Redis can go away inside the window between
		the check and the push."""
		self.probe()

		def enqueue(method, **kwargs):
			if kwargs.get("now"):
				return frappe.call(method, **{k: v for k, v in kwargs.items() if k != "now"})
			background.get_redis_conn.side_effect = ConnectionError("connection refused")
			raise ConnectionError("connection refused")

		with patch.object(frappe, "enqueue", side_effect=enqueue):
			result = background.enqueue_or_run(METHOD, job_id="m365-test", marker="d")

		self.assertTrue(result["ran_inline"])
		self.assertEqual(RUNS, ["d"])
		self.assertIn(background.REDIS_DOWN, codes(result["health"]))

	def test_our_own_bug_is_raised_rather_than_hidden_behind_an_inline_run(self):
		"""A bad queue name or an unimportable method is not the bench's fault, and quietly
		running inline forever is how it would stay invisible."""
		self.probe()

		with patch.object(frappe, "enqueue", side_effect=ValueError("Queue should be one of ...")):
			with self.assertRaises(ValueError):
				background.enqueue_or_run(METHOD, job_id="m365-test", marker="e")

		self.assertEqual(RUNS, [])


class _FakeJob:
	"""Just enough of an rq Job to answer "what state are you in" and record a delete."""

	def __init__(self, status):
		self._status = status
		self.deleted = False

	def get_status(self, refresh=False):
		return self._status

	def delete(self):
		self.deleted = True


def _job(status):
	return _FakeJob(status)


class TestAJobAlreadyInHand(BackgroundTestCase):
	"""The inline fallback exists because nothing is running the work. If something IS running
	it, doing it again here is the one outcome worse than waiting."""

	def test_work_a_worker_has_already_started_is_left_alone(self):
		"""health() can be wrong in this direction: the worker count is up to a minute old, and
		a scheduler that is off does not stop a worker somebody started by hand."""
		ran = []
		self.probe(scheduler_inactive=True)

		with patch.object(background, "get_job") as job, patch.object(
			frappe, "enqueue", side_effect=lambda *a, **kw: ran.append(a)
		):
			job.return_value = _job(JobStatus.STARTED)
			result = background.enqueue_or_run("some.method", job_id="m365-test-running")

		self.assertFalse(result["ran_inline"], "a running job must not be duplicated")
		self.assertEqual(ran, [], "nothing should have been executed here")

	def test_a_job_still_waiting_is_dropped_and_the_work_is_done_here(self):
		"""The stranded job must not outlive the fix and redo a day's work on the first worker."""
		self.probe(scheduler_inactive=True)
		queued = _job(JobStatus.QUEUED)

		with patch.object(background, "get_job", return_value=queued), patch.object(
			frappe, "enqueue"
		) as enqueued:
			result = background.enqueue_or_run("some.method", job_id="m365-test-queued")

		self.assertTrue(queued.deleted)
		self.assertTrue(result["ran_inline"])
		self.assertTrue(enqueued.call_args.kwargs.get("now"), "the inline path routes via now=True")

	def test_an_unreachable_queue_does_not_stop_the_work(self):
		"""Failing closed here would refuse to do anything on exactly the bench that needs it."""
		self.probe(scheduler_inactive=True)

		with patch.object(background, "get_job", side_effect=RuntimeError("redis is gone")), patch.object(
			frappe, "enqueue"
		) as enqueued:
			result = background.enqueue_or_run("some.method", job_id="m365-test-noredis")

		self.assertTrue(result["ran_inline"])
		self.assertTrue(enqueued.call_args.kwargs.get("now"))

	def test_the_inline_call_never_carries_a_job_id(self):
		"""Load-bearing: frappe.enqueue checks `deduplicate` via get_job BEFORE it checks `now`,
		so passing one would reach for redis and defeat the fallback exactly when redis is the
		thing that is broken."""
		self.probe(scheduler_inactive=True)

		with patch.object(background, "get_job", return_value=None), patch.object(
			frappe, "enqueue"
		) as enqueued:
			background.enqueue_or_run("some.method", job_id="m365-test-noid")

		kwargs = enqueued.call_args.kwargs
		self.assertNotIn("job_id", kwargs)
		self.assertNotIn("deduplicate", kwargs)
