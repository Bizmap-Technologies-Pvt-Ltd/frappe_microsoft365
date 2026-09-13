"""Sync engine tests — the bugs these cover are the ones that silently lose data.

Graph is mocked at ``microsoft_graph.graph_delta`` / ``graph_request``; no network, no tokens.
"""

from unittest.mock import patch

import frappe
from frappe.utils import add_to_date, get_datetime, now_datetime

from frappe_microsoft365 import microsoft_calendar_sync as sync
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError, MsGraphResyncRequired
from frappe_microsoft365.tests.base import BaseTestCase

#: A bench where queued work actually runs.
#:
#: Most of this file is about what the sync does with Graph, not about the developer's redis.
#: Left unpinned, tests would queue on a machine with a worker and fall back to running inline
#: on one without — passing or failing on something their subject has nothing to do with.
HEALTHY_QUEUE = {"ok": True, "reasons": [], "fixes": [], "message": "Background jobs are running."}

DEAD_QUEUE = {
	"ok": False,
	"reasons": ["No worker is running, so queued jobs will not be executed."],
	"fixes": ["Start a worker with `bench worker --queue default`."],
	"message": "Background jobs are not running.",
}

CALENDAR = "_Test MS Calendar"


def ms_event(event_id="ms-1", subject="Standup", preview="Graph preview", **overrides):
	event = {
		"id": event_id,
		"subject": subject,
		"bodyPreview": preview,
		"start": {"dateTime": "2026-09-15T09:00:00.0000000", "timeZone": "UTC"},
		"end": {"dateTime": "2026-09-15T09:30:00.0000000", "timeZone": "UTC"},
		"isAllDay": False,
		"isCancelled": False,
		"location": {"displayName": "Room 1"},
		"type": "singleInstance",
	}
	event.update(overrides)
	return event


def ms_attendee(name, address, response="none", kind="required"):
	"""One entry of the Graph event's attendees[] collection."""
	return {
		"emailAddress": {"name": name, "address": address},
		"type": kind,
		"status": {"response": response, "time": "2026-09-14T10:00:00Z"},
	}


class SyncTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		self._cleanup_events()
		if frappe.db.exists("Microsoft Calendar", CALENDAR):
			frappe.delete_doc("Microsoft Calendar", CALENDAR, force=True, ignore_permissions=True)
		self.calendar = frappe.get_doc(
			{
				"doctype": "Microsoft Calendar",
				"account_name": CALENDAR,
				"user": "Administrator",
				"enabled": 1,
				"authorized": 1,
				"pull_from_microsoft_calendar": 1,
				"push_to_microsoft_calendar": 0,
			}
		).insert(ignore_permissions=True)

	def tearDown(self):
		self._cleanup_events()
		super().tearDown()

	def _cleanup_events(self):
		# Deleting an Event queues a Microsoft delete. A teardown is not a thing anyone wants
		# jobs coming out of, and the runner has no worker to answer them anyway.
		#
		# Health is pinned alongside the queue for a subtler reason: on a bench with no worker
		# the delete would take the fallback instead, parking itself on frappe.db.after_commit
		# — where it survives teardown and detonates inside whichever test commits next,
		# counting as that test's work. Pinning here keeps a teardown from leaving anything at
		# all behind, and keeps the whole file's results the same with or without a worker.
		with patch.object(sync.background, "health", return_value=HEALTHY_QUEUE), patch.object(
			frappe, "enqueue"
		):
			for name in frappe.get_all(
				"Event", filters={"custom_microsoft_calendar": CALENDAR}, pluck="name"
			):
				frappe.delete_doc("Event", name, force=True, ignore_permissions=True)

	def _enable_integration(self, enabled):
		frappe.db.set_single_value("Microsoft Settings", "enabled", enabled)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")
		self.addCleanup(frappe.clear_document_cache, "Microsoft Settings", "Microsoft Settings")

	def _local_event(self, ms_id=None, **kwargs):
		"""An Event that originated in Frappe (as the push path leaves it)."""
		values = {
			"doctype": "Event",
			"subject": "Local meeting",
			"starts_on": "2026-09-15 09:00:00",
			"ends_on": "2026-09-15 09:30:00",
			"event_type": "Private",
			"description": "<p>Carefully written agenda</p>",
			"custom_sync_with_microsoft_calendar": 1,
			"custom_microsoft_calendar": CALENDAR,
			"custom_pulled_from_microsoft": 0,
		}
		if ms_id:
			values["custom_microsoft_event_id"] = ms_id
		values.update(kwargs)
		return frappe.get_doc(values).insert(ignore_permissions=True)

	def _event_for(self, ms_id):
		name = frappe.db.get_value("Event", {"custom_microsoft_event_id": ms_id}, "name")
		return frappe.get_doc("Event", name) if name else None

	def _pull_with(self, items, delta_link="DELTA-1"):
		with patch.object(graph, "graph_delta", return_value=(items, delta_link)) as mocked:
			result = sync._pull(self.calendar)
		return result, mocked


class TestPull(SyncTestCase):
	def test_creates_mirror_event(self):
		(upserted, deleted), _ = self._pull_with([ms_event()])

		self.assertEqual((upserted, deleted), (1, 0))
		event = self._event_for("ms-1")
		self.assertIsNotNone(event)
		self.assertEqual(event.subject, "Standup")
		self.assertEqual(event.description, "Graph preview")
		self.assertEqual(event.location, "Room 1")
		self.assertEqual(event.custom_pulled_from_microsoft, 1)
		self.assertEqual(event.custom_microsoft_calendar, CALENDAR)

	def test_stores_delta_link_as_the_watermark(self):
		self._pull_with([ms_event()], delta_link="DELTA-NEXT")

		self.assertEqual(
			frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_link"), "DELTA-NEXT"
		)

	def test_does_not_clobber_an_event_that_originated_in_frappe(self):
		"""Regression: bodyPreview is truncated plain text; it must not replace a real description."""
		local = self._local_event(ms_id="ms-local")

		self._pull_with([ms_event(event_id="ms-local", subject="Renamed in Outlook")])

		local.reload()
		self.assertEqual(local.description, "<p>Carefully written agenda</p>")
		self.assertEqual(local.subject, "Renamed in Outlook")

	def test_does_not_flip_the_origin_flag_on_a_pushed_event(self):
		"""Regression: flipping this flag made every later local edit stop syncing."""
		local = self._local_event(ms_id="ms-local")

		self._pull_with([ms_event(event_id="ms-local")])

		local.reload()
		self.assertEqual(local.custom_pulled_from_microsoft, 0)

	def test_removed_event_is_deleted_locally(self):
		self._pull_with([ms_event(event_id="ms-gone")])
		self.assertIsNotNone(self._event_for("ms-gone"))

		(upserted, deleted), _ = self._pull_with(
			[{"id": "ms-gone", "@removed": {"reason": "deleted"}}]
		)

		self.assertEqual((upserted, deleted), (0, 1))
		self.assertIsNone(self._event_for("ms-gone"))

	def test_cancelled_event_is_deleted_locally(self):
		self._pull_with([ms_event(event_id="ms-cancel")])

		(_, deleted), _ = self._pull_with([ms_event(event_id="ms-cancel", isCancelled=True)])

		self.assertEqual(deleted, 1)
		self.assertIsNone(self._event_for("ms-cancel"))

	def test_series_master_is_skipped_so_occurrences_are_not_duplicated(self):
		(upserted, deleted), _ = self._pull_with(
			[
				ms_event(event_id="ms-master", type="seriesMaster"),
				ms_event(event_id="ms-occurrence-1", type="occurrence"),
			]
		)

		self.assertEqual((upserted, deleted), (1, 0))
		self.assertIsNone(self._event_for("ms-master"))
		self.assertIsNotNone(self._event_for("ms-occurrence-1"))

	def test_unchanged_event_is_not_resaved(self):
		"""A pointless save bumps `modified`, which makes the push step patch it back forever."""
		self._pull_with([ms_event()])
		before = frappe.db.get_value("Event", {"custom_microsoft_event_id": "ms-1"}, "modified")

		(upserted, _), _ = self._pull_with([ms_event()])
		after = frappe.db.get_value("Event", {"custom_microsoft_event_id": "ms-1"}, "modified")

		self.assertEqual(upserted, 0)
		self.assertEqual(before, after)

	def test_changed_event_is_updated(self):
		self._pull_with([ms_event()])

		self._pull_with([ms_event(subject="Standup (moved)", start={"dateTime": "2026-09-15T11:00:00.0000000", "timeZone": "UTC"})])

		event = self._event_for("ms-1")
		self.assertEqual(event.subject, "Standup (moved)")

	def test_expired_delta_token_restarts_a_full_sync(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "delta_link", "STALE")
		self.calendar.reload()

		with patch.object(
			graph, "graph_delta", side_effect=[MsGraphResyncRequired, ([ms_event()], "DELTA-FRESH")]
		) as mocked:
			upserted, _ = sync._pull(self.calendar)

		self.assertEqual(upserted, 1)
		self.assertEqual(mocked.call_count, 2)
		self.assertEqual(
			frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_link"), "DELTA-FRESH"
		)

	def test_reuses_the_stored_delta_link_when_the_window_is_still_valid(self):
		frappe.db.set_value(
			"Microsoft Calendar",
			CALENDAR,
			{
				"delta_link": "https://graph.microsoft.com/v1.0/delta?token=abc",
				"delta_window_end": add_to_date(now_datetime(), days=90),
			},
		)
		self.calendar.reload()

		with patch.object(graph, "graph_delta", return_value=([], "DELTA-2")) as mocked:
			sync._pull(self.calendar)

		self.assertEqual(mocked.call_args[0][0], "https://graph.microsoft.com/v1.0/delta?token=abc")

	def test_reinitialises_when_the_delta_window_is_about_to_run_out(self):
		frappe.db.set_value(
			"Microsoft Calendar",
			CALENDAR,
			{"delta_link": "OLD", "delta_window_end": add_to_date(now_datetime(), days=2)},
		)
		self.calendar.reload()

		with patch.object(graph, "graph_delta", return_value=([], "DELTA-3")) as mocked:
			sync._pull(self.calendar)

		self.assertIn("/me/calendarView/delta", mocked.call_args[0][0])


class TestPrefetch(SyncTestCase):
	"""The pull asks the database once for the whole page, not once per incoming event."""

	def _page(self, tag):
		"""The four things one delta page can say about ids we may or may not already have."""
		return [
			ms_event(event_id=f"{tag}-same"),
			ms_event(event_id=f"{tag}-changed", subject="Standup (moved)"),
			ms_event(event_id=f"{tag}-new"),
			{"id": f"{tag}-gone", "@removed": {"reason": "deleted"}},
		]

	def test_the_prefetch_answers_exactly_as_the_per_event_lookup_does(self):
		"""It is an optimisation, so every outcome has to be the one the old path produced."""
		for tag in ("pre", "raw"):
			for suffix in ("same", "changed", "gone"):
				sync._upsert_event(self.calendar, ms_event(event_id=f"{tag}-{suffix}"))

		prefetched, _ = self._pull_with(self._page("pre"))
		# _pull sets the sync flag around its loop; calling the upsert directly does not, so
		# the removal in the page would queue a Microsoft delete this test has no use for.
		with patch.object(frappe, "enqueue"):
			per_event = [sync._upsert_event(self.calendar, ev) for ev in self._page("raw")]

		self.assertEqual(per_event, ["skipped", "updated", "created", "deleted"])
		# ...which is what _pull counts as one updated plus one created, and one deleted.
		self.assertEqual(prefetched, (2, 1))
		for suffix in ("same", "changed", "new"):
			with_map = self._event_for(f"pre-{suffix}")
			without_map = self._event_for(f"raw-{suffix}")
			self.assertEqual(with_map.subject, without_map.subject)
			self.assertEqual(get_datetime(with_map.starts_on), get_datetime(without_map.starts_on))
		self.assertIsNone(self._event_for("pre-gone"))
		self.assertIsNone(self._event_for("raw-gone"))

	def test_upsert_still_works_when_it_is_handed_no_map(self):
		"""The single-event callers pass none, and must still find the Event that exists."""
		sync._upsert_event(self.calendar, ms_event(event_id="ms-nomap"))

		outcome = sync._upsert_event(self.calendar, ms_event(event_id="ms-nomap", subject="Renamed"))

		self.assertEqual(outcome, "updated")
		self.assertEqual(self._event_for("ms-nomap").subject, "Renamed")
		self.assertEqual(
			len(frappe.get_all("Event", filters={"custom_microsoft_event_id": "ms-nomap"})), 1
		)

	def test_the_prefetch_is_chunked_rather_than_one_enormous_in_clause(self):
		with patch.object(frappe, "get_all", return_value=[]) as mocked:
			sync._existing_event_map([f"ms-bulk-{n}" for n in range(1201)])

		self.assertEqual(mocked.call_count, 3)
		for call in mocked.call_args_list:
			ids = call[1]["filters"]["custom_microsoft_event_id"][1]
			self.assertLessEqual(len(ids), sync.PREFETCH_CHUNK)

	def test_an_id_we_have_never_seen_is_simply_absent_from_the_map(self):
		sync._upsert_event(self.calendar, ms_event(event_id="ms-known"))

		mapping = sync._existing_event_map(["ms-known", "ms-unknown", None, ""])

		self.assertEqual(list(mapping), ["ms-known"])

	def test_the_same_id_twice_in_one_page_updates_rather_than_duplicates(self):
		"""The map has to learn about the Event the first copy created, or the second makes another."""
		(upserted, _), _ = self._pull_with(
			[ms_event(event_id="ms-twice"), ms_event(event_id="ms-twice", subject="Changed again")]
		)

		self.assertEqual(upserted, 2)
		self.assertEqual(
			len(frappe.get_all("Event", filters={"custom_microsoft_event_id": "ms-twice"})), 1
		)
		self.assertEqual(self._event_for("ms-twice").subject, "Changed again")

	def test_a_delete_earlier_in_the_page_leaves_no_stale_map_entry(self):
		"""Otherwise the entry points at a deleted Event and the recreation is lost."""
		self._pull_with([ms_event(event_id="ms-back")])

		(upserted, removed), _ = self._pull_with(
			[
				{"id": "ms-back", "@removed": {"reason": "deleted"}},
				ms_event(event_id="ms-back", subject="Recreated"),
			]
		)

		self.assertEqual((upserted, removed), (1, 1))
		self.assertEqual(self._event_for("ms-back").subject, "Recreated")


class TestWatermark(SyncTestCase):
	def test_failed_pull_does_not_advance_last_sync_or_delta_link(self):
		"""Regression: advancing the watermark after a failure skips that window forever."""
		last_sync = add_to_date(now_datetime(), hours=-3)
		frappe.db.set_value(
			"Microsoft Calendar", CALENDAR, {"delta_link": "KEEP-ME", "last_sync": last_sync}
		)
		self.calendar.reload()

		with patch.object(graph, "graph_delta", side_effect=MsGraphError("Graph exploded")), patch.object(
			frappe.db, "commit"
		):
			result = sync._sync_locked(self.calendar)

		self.assertFalse(result["ok"])
		self.assertEqual(frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_link"), "KEEP-ME")
		self.assertEqual(
			get_datetime(frappe.db.get_value("Microsoft Calendar", CALENDAR, "last_sync")),
			get_datetime(last_sync),
		)

	def test_successful_pull_advances_last_sync_and_clears_the_error(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "last_error", "previous failure")
		self.calendar.reload()

		with patch.object(graph, "graph_delta", return_value=([ms_event()], "DELTA-OK")), patch.object(
			frappe.db, "commit"
		):
			result = sync._sync_locked(self.calendar)

		self.assertTrue(result["ok"])
		self.assertIsNotNone(frappe.db.get_value("Microsoft Calendar", CALENDAR, "last_sync"))
		self.assertFalse(frappe.db.get_value("Microsoft Calendar", CALENDAR, "last_error"))

	def test_partial_page_read_does_not_store_a_watermark(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "delta_link", "KEEP-ME")
		self.calendar.reload()

		with patch.object(graph, "graph_delta", return_value=([ms_event()], None)):
			sync._pull(self.calendar)

		self.assertEqual(frappe.db.get_value("Microsoft Calendar", CALENDAR, "delta_link"), "KEEP-ME")


class TestEntrypoints(SyncTestCase):
	"""The public entry points, including the lock path (v15 and v16 both ship filelock)."""

	def test_sync_calendar_runs_the_whole_round_trip(self):
		with patch.object(graph, "graph_delta", return_value=([ms_event()], "DELTA-RT")), patch.object(
			frappe.db, "commit"
		):
			result = sync.sync_calendar(CALENDAR)

		self.assertTrue(result["ok"])
		self.assertEqual(result["pulled"], 1)
		self.assertIsNotNone(self._event_for("ms-1"))

	def test_sync_calendar_skips_a_calendar_that_is_already_syncing(self):
		"""A run slower than the 15-minute cron must not be overlapped by the next one."""
		from frappe.utils.file_lock import LockTimeoutError

		class Blocked:
			def __enter__(self):
				raise LockTimeoutError("already locked")

			def __exit__(self, *args):
				return False

		with patch.object(sync, "_calendar_lock", return_value=Blocked()), patch.object(
			graph, "graph_delta"
		) as mocked:
			result = sync.sync_calendar(CALENDAR)

		self.assertFalse(result["ok"])
		self.assertIn("already running", result["message"])
		mocked.assert_not_called()

	def test_sync_calendar_refuses_an_unauthorized_calendar(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "authorized", 0)

		with patch.object(graph, "graph_delta") as mocked:
			result = sync.sync_calendar(CALENDAR)

		self.assertFalse(result["ok"])
		mocked.assert_not_called()

	def test_sync_all_is_a_noop_while_the_integration_is_disabled(self):
		self._enable_integration(0)

		with patch.object(graph, "graph_delta") as mocked:
			sync.sync_all()

		mocked.assert_not_called()


class TestScheduledFanOut(SyncTestCase):
	def setUp(self):
		super().setUp()
		queue = patch.object(sync.background, "health", return_value=HEALTHY_QUEUE)
		queue.start()
		self.addCleanup(queue.stop)

	"""The scheduled pass queues a job per calendar instead of syncing them all in one.

	Regression: the cron hook gets one default-queue job with a 300 second timeout, and at
	roughly fifteen seconds a calendar it was killed part-way down a stable list — so the
	calendars at the tail of it never synced at all.
	"""

	def test_sync_all_queues_one_job_per_calendar(self):
		self._enable_integration(1)

		with patch.object(frappe, "enqueue") as enqueued, patch.object(
			sync, "_sync_job_running", return_value=False
		):
			result = sync.sync_all()

		self.assertIn(CALENDAR, result["queued"])
		job_ids = [call[1]["job_id"] for call in enqueued.call_args_list]
		self.assertIn(f"m365-sync-{CALENDAR}", job_ids)
		self.assertEqual(len(job_ids), len(result["queued"]))
		self.assertEqual(len(job_ids), len(set(job_ids)))

	def test_the_scheduled_pass_does_no_graph_work_of_its_own(self):
		self._enable_integration(1)

		with patch.object(frappe, "enqueue"), patch.object(
			sync, "_sync_job_running", return_value=False
		), patch.object(graph, "graph_delta") as mocked:
			sync.sync_all()

		mocked.assert_not_called()

	def test_a_calendar_already_syncing_is_not_queued_a_second_time(self):
		self._enable_integration(1)

		with patch.object(frappe, "enqueue") as enqueued, patch.object(
			sync, "_sync_job_running", return_value=True
		):
			result = sync.sync_all()

		enqueued.assert_not_called()
		self.assertIn(CALENDAR, result["skipped"])

	def test_a_queue_that_cannot_be_reached_is_not_an_exception(self):
		"""A scheduled hook that raises takes the whole pass down with it, redis or no redis."""
		self._enable_integration(1)

		with patch.object(frappe, "enqueue", side_effect=Exception("redis is down")), patch.object(
			sync, "_sync_job_running", return_value=False
		), patch.object(frappe, "log_error"):
			result = sync.sync_all()

		self.assertEqual(result["queued"], [])
		self.assertIn(CALENDAR, result["skipped"])

	def test_enqueue_sync_reports_the_job_it_queued(self):
		with patch.object(frappe, "enqueue") as enqueued, patch.object(
			sync, "_sync_job_running", return_value=False
		):
			result = sync.enqueue_sync(CALENDAR)

		self.assertEqual(result, {"queued": True, "job_id": f"m365-sync-{CALENDAR}"})
		self.assertEqual(enqueued.call_args[1]["calendar_name"], CALENDAR)
		self.assertTrue(enqueued.call_args[1]["deduplicate"])

	def test_enqueue_sync_says_so_when_it_queued_nothing(self):
		with patch.object(frappe, "enqueue") as enqueued, patch.object(
			sync, "_sync_job_running", return_value=True
		):
			result = sync.enqueue_sync(CALENDAR)

		enqueued.assert_not_called()
		self.assertEqual(result, {"queued": False, "job_id": f"m365-sync-{CALENDAR}"})

	def test_the_queued_job_does_the_same_work_as_a_sync_by_hand(self):
		with patch.object(graph, "graph_delta", return_value=([ms_event()], "DELTA-JOB")), patch.object(
			frappe.db, "commit"
		), patch.object(frappe, "publish_realtime") as published:
			result = sync.run_sync_job(CALENDAR, notify_user="Administrator")

		self.assertTrue(result["ok"])
		self.assertIsNotNone(self._event_for("ms-1"))
		self.assertEqual(published.call_args[0][0], "microsoft365_sync_done")
		self.assertEqual(published.call_args[1]["user"], "Administrator")
		payload = published.call_args[0][1]
		self.assertEqual(payload["calendar"], CALENDAR)
		self.assertEqual((payload["pulled"], payload["deleted"], payload["pushed"]), (1, 0, 0))
		self.assertIn("Pulled 1", payload["message"])

	def test_a_run_nobody_asked_for_is_reported_to_the_calendars_owner(self):
		"""The scheduled pass has no person behind it; the owner is who has that form open."""
		with patch.object(graph, "graph_delta", return_value=([], "DELTA-OWNER")), patch.object(
			frappe.db, "commit"
		), patch.object(frappe, "publish_realtime") as published:
			sync.run_sync_job(CALENDAR)

		self.assertEqual(published.call_args[1]["user"], "Administrator")


class TestEventSaveOffRequest(SyncTestCase):
	"""A save waits for Graph only where something on screen depends on the answer."""

	def setUp(self):
		super().setUp()
		queue = patch.object(sync.background, "health", return_value=HEALTHY_QUEUE)
		queue.start()
		self.addCleanup(queue.stop)
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "push_to_microsoft_calendar", 1)
		frappe.clear_document_cache("Microsoft Calendar", CALENDAR)
		self.addCleanup(frappe.clear_document_cache, "Microsoft Calendar", CALENDAR)
		self.calendar.reload()

	def _linked_event(self, ms_id):
		"""An Event Microsoft already knows about, inserted without queueing a patch for it."""
		with patch.object(frappe, "enqueue"), patch.object(graph, "graph_request"):
			return self._local_event(ms_id=ms_id)

	def _jobs_for(self, enqueued, function_name):
		"""Only the calls that queued one of ours.

		Saving a document is allowed to queue things of Frappe's own (a search index, say), and
		counting those as ours would make these tests fail for something unrelated.
		"""
		jobs = []
		for call in enqueued.call_args_list:
			method = call[0][0] if call[0] else call[1].get("method", "")
			if str(method).endswith(function_name):
				jobs.append(call)
		return jobs

	def test_creating_an_event_still_pushes_inside_the_save(self):
		"""The join link has to reach the document the browser gets back, not a worker."""
		created = {
			"id": "ms-live",
			"onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/live"},
		}

		with patch.object(graph, "graph_request", return_value=created) as mocked, patch.object(
			frappe, "enqueue"
		) as enqueued:
			event = self._local_event(custom_add_teams_meeting=1)

		self.assertEqual(mocked.call_args[0][0], "POST")
		self.assertEqual(self._jobs_for(enqueued, "patch_event_in_graph"), [])
		self.assertEqual(event.custom_microsoft_event_id, "ms-live")
		self.assertEqual(event.custom_teams_join_url, "https://teams.microsoft.com/l/meetup-join/live")

	def test_editing_an_event_queues_the_patch_instead_of_calling_graph(self):
		event = self._linked_event("ms-patch")

		with patch.object(graph, "graph_request") as mocked, patch.object(
			frappe, "enqueue"
		) as enqueued:
			event.subject = "Renamed in Frappe"
			event.save(ignore_permissions=True)

		mocked.assert_not_called()
		jobs = self._jobs_for(enqueued, "patch_event_in_graph")
		self.assertEqual(len(jobs), 1)
		self.assertEqual(jobs[0][1]["event_name"], event.name)
		self.assertEqual(jobs[0][1]["job_id"], f"m365-event-patch-{event.name}")
		# ...and not before the save it belongs to has actually committed.
		self.assertTrue(jobs[0][1]["enqueue_after_commit"])

	def test_ten_rapid_saves_queue_one_patch_and_not_ten(self):
		event = self._linked_event("ms-rapid")

		with patch.object(graph, "graph_request"), patch.object(frappe, "enqueue") as enqueued:
			for n in range(10):
				event.subject = f"Edit {n}"
				event.save(ignore_permissions=True)

		jobs = self._jobs_for(enqueued, "patch_event_in_graph")
		self.assertEqual(len({job[1]["job_id"] for job in jobs}), 1)
		self.assertTrue(all(job[1]["deduplicate"] for job in jobs))

	def test_the_queued_patch_sends_the_event_as_it_stands_when_it_runs(self):
		"""Not a snapshot: the Event can be edited again before a worker picks the job up."""
		event = self._linked_event("ms-late")
		frappe.db.set_value("Event", event.name, "subject", "Edited again while queued")

		with patch.object(graph, "graph_request", return_value={}) as mocked:
			sync.patch_event_in_graph(event.name)

		self.assertEqual(mocked.call_args[0][0], "PATCH")
		self.assertIn("ms-late", mocked.call_args[0][1])
		self.assertEqual(mocked.call_args[1]["json"]["subject"], "Edited again while queued")

	def test_a_patch_for_an_event_that_has_since_been_deleted_does_nothing(self):
		event = self._linked_event("ms-vanished")
		name = event.name
		with patch.object(frappe, "enqueue"):
			frappe.delete_doc("Event", name, force=True, ignore_permissions=True)

		with patch.object(graph, "graph_request") as mocked:
			sync.patch_event_in_graph(name)

		mocked.assert_not_called()

	def test_a_patch_for_an_event_no_longer_bound_to_microsoft_does_nothing(self):
		event = self._linked_event("ms-unticked")
		frappe.db.set_value("Event", event.name, "custom_sync_with_microsoft_calendar", 0)

		with patch.object(graph, "graph_request") as mocked:
			sync.patch_event_in_graph(event.name)

		mocked.assert_not_called()

	def test_the_queued_patch_cannot_feed_itself_another_job(self):
		"""in_microsoft_sync lives in the request that set it; a worker starts without it."""
		event = self._linked_event("ms-loop")
		seen = {}

		def remember(*args, **kwargs):
			seen["flag"] = frappe.flags.in_microsoft_sync
			return {}

		with patch.object(graph, "graph_request", side_effect=remember):
			sync.patch_event_in_graph(event.name)

		self.assertTrue(seen["flag"])
		self.assertFalse(frappe.flags.in_microsoft_sync)

	def test_deleting_an_event_queues_the_removal_with_plain_ids(self):
		"""A doc reference would be useless: the Event is gone before the job ever runs."""
		event = self._linked_event("ms-bye")

		with patch.object(frappe, "enqueue") as enqueued, patch.object(
			graph, "graph_request"
		) as mocked:
			frappe.delete_doc("Event", event.name, force=True, ignore_permissions=True)

		mocked.assert_not_called()
		jobs = self._jobs_for(enqueued, "delete_event_in_graph")
		self.assertEqual(len(jobs), 1)
		self.assertEqual(jobs[0][1]["calendar_name"], CALENDAR)
		self.assertEqual(jobs[0][1]["ms_event_id"], "ms-bye")
		self.assertEqual(jobs[0][1]["job_id"], "m365-event-delete-ms-bye")
		self.assertTrue(jobs[0][1]["enqueue_after_commit"])

	def test_the_queued_delete_runs_with_no_event_left_to_read(self):
		with patch.object(graph, "graph_request") as mocked:
			sync.delete_event_in_graph(CALENDAR, "ms-already-deleted")

		self.assertEqual(mocked.call_args[0][0], "DELETE")
		self.assertIn("ms-already-deleted", mocked.call_args[0][1])

	def test_a_delete_is_abandoned_when_an_event_still_claims_that_id(self):
		"""Removing it would strand that mirror, and the next pull would delete it locally too."""
		self._linked_event("ms-still-here")

		with patch.object(graph, "graph_request") as mocked:
			sync.delete_event_in_graph(CALENDAR, "ms-still-here")

		mocked.assert_not_called()


class TestSyncCost(SyncTestCase):
	"""What Sync Now asks before it decides between running inline and queueing."""

	def _make_it_warm(self):
		"""The state an ordinary incremental run leaves a calendar in."""
		frappe.db.set_value(
			"Microsoft Calendar",
			CALENDAR,
			{
				"delta_link": "https://graph.microsoft.com/v1.0/delta?token=warm",
				"delta_window_end": add_to_date(now_datetime(), days=90),
				"last_sync": now_datetime(),
			},
			update_modified=False,
		)
		self.calendar.reload()
		# background_reasons reads the duration off the document, so the test sets it there.
		self.calendar.last_sync_seconds = 1.4

	def test_a_warm_incremental_sync_runs_inline(self):
		self._make_it_warm()

		self.assertEqual(sync.background_reasons(self.calendar), [])

	def test_a_first_sync_is_a_reason_to_queue(self):
		reasons = sync.background_reasons(self.calendar)

		self.assertTrue(reasons)
		self.assertIn("first sync", reasons[0])

	def test_a_delta_window_about_to_run_out_is_a_reason_to_queue(self):
		self._make_it_warm()
		self.calendar.delta_window_end = add_to_date(now_datetime(), days=2)

		self.assertTrue(sync.background_reasons(self.calendar))

	def test_a_big_push_backlog_is_a_reason_to_queue(self):
		self._make_it_warm()
		self.calendar.push_to_microsoft_calendar = 1

		with patch.object(sync, "_pending_push_count", return_value=sync.LARGE_PUSH_BACKLOG):
			reasons = sync.background_reasons(self.calendar)

		self.assertTrue(reasons)
		self.assertIn(str(sync.LARGE_PUSH_BACKLOG), reasons[0])

	def test_a_handful_of_pending_events_is_not(self):
		self._make_it_warm()
		self.calendar.push_to_microsoft_calendar = 1

		with patch.object(sync, "_pending_push_count", return_value=sync.LARGE_PUSH_BACKLOG - 1):
			self.assertEqual(sync.background_reasons(self.calendar), [])

	def test_a_slow_last_run_is_a_reason_to_queue(self):
		self._make_it_warm()
		self.calendar.last_sync_seconds = sync.SLOW_SYNC_SECONDS + 5

		reasons = sync.background_reasons(self.calendar)

		self.assertTrue(reasons)
		self.assertIn(str(sync.SLOW_SYNC_SECONDS + 5), reasons[0])

	def test_the_backlog_count_is_what_the_push_would_actually_send(self):
		"""A count that does not match the push would make the estimate worse than none."""
		self._make_it_warm()
		self._local_event()
		self._pull_with([ms_event(event_id="ms-not-pushed-back")])

		self.assertEqual(sync._pending_push_count(self.calendar), 1)

	def test_a_completed_sync_records_how_long_it_took(self):
		if not frappe.get_meta("Microsoft Calendar").has_field("last_sync_seconds"):
			# The same condition the sync itself guards on, for the window before the column
			# has been migrated onto the doctype.
			self.skipTest("last_sync_seconds is not on this site's Microsoft Calendar yet")

		with patch.object(graph, "graph_delta", return_value=([ms_event()], "DELTA-TIMED")), patch.object(
			frappe.db, "commit"
		):
			sync._sync_locked(self.calendar)

		recorded = frappe.db.get_value("Microsoft Calendar", CALENDAR, "last_sync_seconds")
		self.assertIsNotNone(recorded)
		self.assertGreaterEqual(recorded, 0)


class TestPush(SyncTestCase):
	def test_new_local_event_is_created_in_graph(self):
		event = self._local_event()

		with patch.object(graph, "graph_request", return_value={"id": "ms-created"}) as mocked, patch.object(
			frappe.db, "commit"
		):
			pushed = sync._push(self.calendar)

		self.assertEqual(pushed, 1)
		self.assertEqual(mocked.call_args[0][0], "POST")
		event.reload()
		self.assertEqual(event.custom_microsoft_event_id, "ms-created")

	def test_event_edited_while_offline_is_patched(self):
		"""doc_events cannot fire while Graph is unreachable; the next sync must catch up."""
		self._local_event(ms_id="ms-edited")
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "last_sync", add_to_date(now_datetime(), hours=-1))
		self.calendar.reload()

		with patch.object(graph, "graph_request", return_value={}) as mocked, patch.object(
			frappe.db, "commit"
		):
			pushed = sync._push(self.calendar)

		self.assertEqual(pushed, 1)
		self.assertEqual(mocked.call_args[0][0], "PATCH")

	def test_mirrored_events_are_never_pushed_back(self):
		self._pull_with([ms_event(event_id="ms-mirror")])

		with patch.object(graph, "graph_request") as mocked, patch.object(frappe.db, "commit"):
			pushed = sync._push(self.calendar)

		self.assertEqual(pushed, 0)
		mocked.assert_not_called()

	def test_graph_body_carries_description_and_location(self):
		event = self._local_event(location="Board room")
		body = sync._event_to_graph_body(event)

		self.assertEqual(body["body"]["content"], "<p>Carefully written agenda</p>")
		self.assertEqual(body["location"], {"displayName": "Board room"})
		self.assertIn("dateTime", body["start"])


class TestTeamsMeetings(SyncTestCase):
	"""A Frappe Event can ask Outlook for a Teams meeting, the same as Outlook's own toggle."""

	def test_plain_event_is_not_an_online_meeting(self):
		body = sync._event_to_graph_body(self._local_event())

		self.assertNotIn("isOnlineMeeting", body)

	def test_ticking_the_box_asks_graph_for_a_teams_meeting(self):
		event = self._local_event(custom_add_teams_meeting=1)

		body = sync._event_to_graph_body(event)

		self.assertTrue(body["isOnlineMeeting"])
		self.assertEqual(body["onlineMeetingProvider"], "teamsForBusiness")

	def test_an_end_before_the_start_is_repaired_rather_than_rejected(self):
		"""Graph answers ErrorPropertyValidationFailure; Frappe pre-fills both times from now.

		A person can no longer save this (event_validate refuses it), but an Event inserted
		programmatically or predating that guard still must not break the push.
		"""
		event = frappe.get_doc(
			{
				"doctype": "Event",
				"subject": "Backwards",
				"starts_on": "2026-09-13 04:54:27",
				"ends_on": "2026-09-13 04:54:18",
				"event_type": "Private",
			}
		)

		body = sync._event_to_graph_body(event)

		self.assertGreater(body["end"]["dateTime"], body["start"]["dateTime"])

	def test_a_person_cannot_save_a_backwards_event_that_syncs(self):
		with self.assertRaises(frappe.ValidationError):
			self._local_event(starts_on="2026-09-13 04:54:27", ends_on="2026-09-13 04:54:18")

	def test_the_same_event_is_accepted_when_it_comes_from_microsoft(self):
		"""Outlook is authoritative during a pull; refusing its data would stall the sync."""
		frappe.flags.in_microsoft_sync = True
		self.addCleanup(lambda: setattr(frappe.flags, "in_microsoft_sync", False))

		event = self._local_event(starts_on="2026-09-13 04:54:27", ends_on="2026-09-13 04:54:18")

		self.assertTrue(event.name)

	def test_a_sensible_end_is_left_alone(self):
		event = self._local_event(
			starts_on="2026-09-13 10:00:00", ends_on="2026-09-13 11:30:00"
		)

		body = sync._event_to_graph_body(event)

		self.assertIn("11:30", body["end"]["dateTime"])

	def test_a_missing_end_still_defaults_to_half_an_hour(self):
		event = self._local_event(starts_on="2026-09-13 10:00:00", ends_on=None)

		body = sync._event_to_graph_body(event)

		self.assertIn("10:30", body["end"]["dateTime"])

	def test_join_link_and_outlook_link_are_stored_after_creation(self):
		event = self._local_event(custom_add_teams_meeting=1)
		created = {
			"id": "ms-teams-1",
			"webLink": "https://outlook.office365.com/owa/?itemid=abc",
			"onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/abc"},
		}

		sync._store_graph_response(event.name, created)

		event.reload()
		self.assertEqual(event.custom_microsoft_event_id, "ms-teams-1")
		self.assertEqual(event.custom_teams_join_url, "https://teams.microsoft.com/l/meetup-join/abc")
		self.assertIn("outlook.office365.com", event.custom_microsoft_web_link)

	def test_a_teams_meeting_pulled_from_outlook_keeps_its_join_link(self):
		self._pull_with(
			[
				ms_event(
					event_id="ms-online",
					onlineMeeting={"joinUrl": "https://teams.microsoft.com/l/meetup-join/xyz"},
					webLink="https://outlook.office365.com/owa/?itemid=xyz",
				)
			]
		)

		event = self._event_for("ms-online")
		self.assertEqual(event.custom_teams_join_url, "https://teams.microsoft.com/l/meetup-join/xyz")
		self.assertEqual(event.custom_add_teams_meeting, 1)
		self.assertIn("outlook.office365.com", event.custom_microsoft_web_link)

	def test_an_ordinary_pulled_event_has_no_join_link(self):
		self._pull_with([ms_event(event_id="ms-offline")])

		event = self._event_for("ms-offline")
		self.assertFalse(event.custom_teams_join_url)
		self.assertFalse(event.custom_add_teams_meeting)


class TestAttendeesPull(SyncTestCase):
	"""The invitation side of an Outlook event: who is on it, and what everyone replied."""

	def test_event_select_asks_graph_for_the_invitation_fields(self):
		"""Graph returns only the selected properties; a gap here means no attendees, ever."""
		for field in ("attendees", "organizer", "isOrganizer", "responseStatus", "responseRequested"):
			self.assertIn(field, sync.EVENT_SELECT)

	def test_attendee_summary_is_one_readable_line_per_person(self):
		self._pull_with(
			[
				ms_event(
					event_id="ms-invite",
					attendees=[
						ms_attendee("Asha Rao", "asha@example.com", response="accepted"),
						ms_attendee("Ben Muir", "ben@example.com", response="declined"),
						ms_attendee("Cara Lim", "cara@example.com", response="tentativelyAccepted"),
					],
				)
			]
		)

		lines = self._event_for("ms-invite").custom_microsoft_attendees.split("\n")
		self.assertEqual(lines[0], "Asha Rao (asha@example.com) — accepted")
		self.assertEqual(lines[1], "Ben Muir (ben@example.com) — declined")
		self.assertEqual(lines[2], "Cara Lim (cara@example.com) — tentative")

	def test_an_address_used_as_the_name_is_not_printed_twice(self):
		"""Outlook fills `name` with the address itself for people outside the tenant."""
		self._pull_with(
			[
				ms_event(
					event_id="ms-external",
					attendees=[ms_attendee("dev@partner.com", "dev@partner.com")],
				)
			]
		)

		self.assertEqual(
			self._event_for("ms-external").custom_microsoft_attendees, "dev@partner.com — no response"
		)

	def test_a_room_is_labelled_with_its_attendee_type(self):
		"""A room declining is a different problem from a person declining."""
		self._pull_with(
			[
				ms_event(
					event_id="ms-room",
					attendees=[
						ms_attendee("Board room", "board@example.com", response="declined", kind="resource")
					],
				)
			]
		)

		self.assertEqual(
			self._event_for("ms-room").custom_microsoft_attendees,
			"Board room (board@example.com) — resource, declined",
		)

	def test_organizer_is_mapped(self):
		self._pull_with(
			[
				ms_event(
					event_id="ms-organized",
					organizer={"emailAddress": {"name": "Dana Fox", "address": "dana@example.com"}},
					isOrganizer=False,
				)
			]
		)

		self.assertEqual(self._event_for("ms-organized").custom_microsoft_organizer, "dana@example.com")

	def test_my_response_is_stored_verbatim_as_graph_reports_it(self):
		"""Stored raw so it can be compared with a Graph payload without a lookup table."""
		self._pull_with(
			[
				ms_event(
					event_id="ms-mine",
					responseStatus={"response": "notResponded", "time": "0001-01-01T00:00:00Z"},
				)
			]
		)

		self.assertEqual(self._event_for("ms-mine").custom_microsoft_my_response, "notResponded")

	def test_an_event_with_nobody_invited_has_an_empty_summary(self):
		self._pull_with([ms_event(event_id="ms-solo")])

		self.assertFalse(self._event_for("ms-solo").custom_microsoft_attendees)

	def test_a_mirror_is_cleared_when_outlook_empties_the_invitation(self):
		self._pull_with(
			[
				ms_event(
					event_id="ms-cleared",
					attendees=[ms_attendee("Asha Rao", "asha@example.com", response="accepted")],
				)
			]
		)
		self.assertTrue(self._event_for("ms-cleared").custom_microsoft_attendees)

		self._pull_with([ms_event(event_id="ms-cleared", attendees=[])])

		self.assertFalse(self._event_for("ms-cleared").custom_microsoft_attendees)

	def test_a_frappe_originated_event_is_not_blanked_by_a_pull(self):
		"""A payload that says nothing about attendees is not a payload that says "nobody"."""
		local = self._local_event(ms_id="ms-local-invite")
		frappe.db.set_value(
			"Event",
			local.name,
			{
				"custom_microsoft_attendees": "Asha Rao (asha@example.com) — accepted",
				"custom_microsoft_organizer": "dana@example.com",
				"custom_microsoft_my_response": "organizer",
			},
			update_modified=False,
		)

		self._pull_with([ms_event(event_id="ms-local-invite", subject="Renamed in Outlook")])

		local.reload()
		self.assertEqual(local.custom_microsoft_attendees, "Asha Rao (asha@example.com) — accepted")
		self.assertEqual(local.custom_microsoft_organizer, "dana@example.com")
		self.assertEqual(local.custom_microsoft_my_response, "organizer")
		# ...and the fields Outlook *did* carry still win.
		self.assertEqual(local.subject, "Renamed in Outlook")


class TestAttendeesPush(SyncTestCase):
	"""Frappe's participants go out as Graph attendees, minus anyone we cannot address."""

	def _contact(self, first_name, email=None):
		# Contact is named after first_name, so a run that died before its cleanup would
		# otherwise make every later run fail on a duplicate name instead of on the bug.
		if frappe.db.exists("Contact", first_name):
			frappe.delete_doc("Contact", first_name, force=True, ignore_permissions=True)
		doc = frappe.get_doc({"doctype": "Contact", "first_name": first_name})
		if email:
			doc.append("email_ids", {"email_id": email, "is_primary": 1})
		doc.insert(ignore_permissions=True)
		self.addCleanup(frappe.delete_doc, "Contact", doc.name, force=True, ignore_permissions=True)
		return doc.name

	def _event_with_participants(self, *participants):
		event = self._local_event()
		for participant in participants:
			event.append("event_participants", participant)
		event.save(ignore_permissions=True)
		return event

	def test_push_body_carries_resolved_participants(self):
		contact = self._contact("_Test MS Asha", "asha@example.com")

		event = self._event_with_participants(
			{"reference_doctype": "Contact", "reference_docname": contact}
		)
		body = sync._event_to_graph_body(event)

		self.assertEqual(len(body["attendees"]), 1)
		self.assertEqual(body["attendees"][0]["emailAddress"]["address"], "asha@example.com")
		self.assertEqual(body["attendees"][0]["type"], "required")

	def test_a_participant_without_an_email_is_skipped_rather_than_faked(self):
		"""Graph rejects an attendee with no address; a guessed one invites a stranger."""
		reachable = self._contact("_Test MS Ben", "ben@example.com")
		unreachable = self._contact("_Test MS Nomail")

		event = self._event_with_participants(
			{"reference_doctype": "Contact", "reference_docname": reachable},
			{"reference_doctype": "Contact", "reference_docname": unreachable},
		)
		body = sync._event_to_graph_body(event)

		self.assertEqual(
			[a["emailAddress"]["address"] for a in body["attendees"]], ["ben@example.com"]
		)

	def test_the_attendees_key_is_omitted_when_nothing_resolves(self):
		"""Graph reads an empty attendees array as "remove everyone"."""
		unreachable = self._contact("_Test MS Nobody")

		event = self._event_with_participants(
			{"reference_doctype": "Contact", "reference_docname": unreachable}
		)

		self.assertNotIn("attendees", sync._event_to_graph_body(event))

	def test_an_event_with_no_participants_sends_no_attendees_key(self):
		self.assertNotIn("attendees", sync._event_to_graph_body(self._local_event()))

	def test_the_same_address_is_only_invited_once(self):
		"""A Contact and the User behind it are two rows and one person."""
		contact = self._contact("_Test MS Dup", "dup@example.com")

		event = self._event_with_participants(
			{"reference_doctype": "Contact", "reference_docname": contact},
			{
				"reference_doctype": "User",
				"reference_docname": "Administrator",
				"email": "DUP@example.com",
			},
		)

		self.assertEqual(len(sync._event_to_graph_body(event)["attendees"]), 1)

	def test_a_row_that_carries_its_own_email_needs_no_lookup(self):
		self.assertEqual(sync._participant_email({"email": " asha@example.com "}), "asha@example.com")

	def test_a_user_participant_resolves_through_the_user_record(self):
		expected = frappe.db.get_value("User", "Administrator", "email") or "Administrator"

		self.assertEqual(
			sync._participant_email(
				{"reference_doctype": "User", "reference_docname": "Administrator"}
			),
			expected,
		)

	def test_a_participant_linked_to_nothing_resolves_to_nothing(self):
		self.assertIsNone(sync._participant_email({}))


class TestAuthCallback(BaseTestCase):
	"""A browser lands on the callback, so nothing may escape as a traceback."""

	def _callback(self, **kwargs):
		from frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar import (
			microsoft_calendar as mc,
		)

		with patch.object(frappe, "respond_as_web_page") as page, patch.object(frappe, "log_error"):
			mc.callback(**kwargs)
		return page

	def test_an_unknown_state_renders_a_page_instead_of_raising(self):
		page = self._callback(code="x", state="not-a-real-state")

		page.assert_called_once()
		self.assertIn("sign-in failed", page.call_args[0][0].lower())

	def test_microsofts_own_error_is_shown_to_the_person(self):
		page = self._callback(
			error="invalid_client",
			error_description="AADSTS7000215: Invalid client secret provided.",
		)

		page.assert_called_once()
		body = page.call_args[0][1]
		self.assertIn("AADSTS7000215", body)
		# the doctor's explanation rides along, so the page says what to do about it
		self.assertIn("secret", body.lower())


class TestDuplicateGuard(SyncTestCase):
	"""Regression: a push that cannot record the id duplicates the meeting on every run.

	This is what put three copies of one meeting in a real calendar: the event was created
	in Outlook, the id came back too long for the column, and the next run saw an unlinked
	Event and created another.
	"""

	def test_a_link_that_cannot_be_written_removes_the_event_again(self):
		event = self._local_event()
		created = {"id": "ms-orphan", "webLink": "https://outlook.office365.com/x"}

		with patch.object(
			frappe.db, "set_value", side_effect=Exception("Data too long for column")
		), patch.object(graph, "graph_request") as mocked:
			stored = sync._store_graph_response(event.name, created, CALENDAR)

		self.assertFalse(stored)
		mocked.assert_called_once()
		self.assertEqual(mocked.call_args[0][0], "DELETE")
		self.assertIn("ms-orphan", mocked.call_args[0][1])

	def test_a_link_that_does_not_round_trip_removes_the_event_again(self):
		"""A write can be accepted and still not come back the same (truncation, sanitising)."""
		event = self._local_event()

		with patch.object(frappe.db, "set_value"), patch.object(
			frappe.db, "get_value", return_value="something-else"
		), patch.object(graph, "graph_request") as mocked:
			stored = sync._store_graph_response(event.name, {"id": "ms-truncated"}, CALENDAR)

		self.assertFalse(stored)
		self.assertEqual(mocked.call_args[0][0], "DELETE")

	def test_a_good_link_is_kept_and_nothing_is_deleted(self):
		event = self._local_event()

		with patch.object(graph, "graph_request") as mocked:
			stored = sync._store_graph_response(event.name, {"id": "ms-fine"}, CALENDAR)

		self.assertTrue(stored)
		mocked.assert_not_called()
		event.reload()
		self.assertEqual(event.custom_microsoft_event_id, "ms-fine")

	def test_without_a_calendar_it_reports_rather_than_guessing(self):
		event = self._local_event()

		with patch.object(
			frappe.db, "set_value", side_effect=Exception("boom")
		), patch.object(graph, "graph_request") as mocked:
			stored = sync._store_graph_response(event.name, {"id": "ms-x"})

		self.assertFalse(stored)
		mocked.assert_not_called()


class TestTimezones(BaseTestCase):
	def test_windows_timezone_name_is_mapped_not_assumed_utc(self):
		"""Regression: ZoneInfo cannot parse Windows ids; the old fallback shifted meetings."""
		from zoneinfo import ZoneInfo

		from frappe.utils import get_system_timezone

		converted = sync._ms_dt_to_system(
			{"dateTime": "2026-01-15T09:00:00.0000000", "timeZone": "Pacific Standard Time"}
		)
		expected = (
			get_datetime("2026-01-15 09:00:00")
			.replace(tzinfo=ZoneInfo("America/Los_Angeles"))
			.astimezone(ZoneInfo(get_system_timezone()))
			.replace(tzinfo=None)
		)

		self.assertEqual(converted, expected)

	def test_iana_timezone_still_works(self):
		from zoneinfo import ZoneInfo

		from frappe.utils import get_system_timezone

		converted = sync._ms_dt_to_system(
			{"dateTime": "2026-01-15T09:00:00.0000000", "timeZone": "UTC"}
		)
		expected = (
			get_datetime("2026-01-15 09:00:00")
			.replace(tzinfo=ZoneInfo("UTC"))
			.astimezone(ZoneInfo(get_system_timezone()))
			.replace(tzinfo=None)
		)

		self.assertEqual(converted, expected)

	def test_offset_in_the_payload_is_respected(self):
		converted = sync._ms_dt_to_system(
			{"dateTime": "2026-01-15T09:00:00+00:00", "timeZone": "Pacific Standard Time"}
		)
		self.assertIsNotNone(converted)


class TestADeadQueueDoesNotEatTheWork(SyncTestCase):
	"""frappe.enqueue succeeds against a Redis nobody is listening to.

	That is the whole failure: the save returns, the queue grows, and Outlook quietly stops
	matching Frappe. These tests pin the guard that turns that silence into either the work
	being done anyway, or a person being told.
	"""

	def setUp(self):
		super().setUp()
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "push_to_microsoft_calendar", 1)
		frappe.clear_document_cache("Microsoft Calendar", CALENDAR)
		self.addCleanup(frappe.clear_document_cache, "Microsoft Calendar", CALENDAR)
		self.calendar.reload()

	def _linked_event(self, ms_id):
		with patch.object(sync.background, "health", return_value=HEALTHY_QUEUE), patch.object(
			frappe, "enqueue"
		), patch.object(graph, "graph_request"):
			event = frappe.get_doc(
				{
					"doctype": "Event",
					"subject": "Standup",
					"starts_on": add_to_date(now_datetime(), hours=1),
					"ends_on": add_to_date(now_datetime(), hours=2),
					"event_type": "Private",
					"custom_sync_with_microsoft_calendar": 1,
					"custom_microsoft_calendar": CALENDAR,
					"custom_microsoft_event_id": ms_id,
				}
			).insert(ignore_permissions=True)
		self.addCleanup(self._forget, event.name)
		return event

	def _forget(self, name):
		# Pinned healthy: a teardown on a bench with no worker would otherwise register the
		# Microsoft delete as an after-commit callback, which then fires inside whichever test
		# commits next and counts as its work. Cleanup must leave nothing behind.
		with patch.object(sync.background, "health", return_value=HEALTHY_QUEUE), patch.object(
			frappe, "enqueue"
		), patch.object(graph, "graph_request"):
			if frappe.db.exists("Event", name):
				frappe.delete_doc("Event", name, force=True, ignore_permissions=True)

	def test_an_edit_is_done_after_commit_rather_than_queued_into_a_void(self):
		event = self._linked_event("AAMk-dead-queue-1")
		# Flush anything an earlier test left pending on the commit hook, so what fires below is
		# this test's work and only this test's work.
		frappe.db.commit()  # nosemgrep

		with patch.object(sync.background, "health", return_value=DEAD_QUEUE), patch.object(
			frappe, "enqueue"
		) as enqueued, patch.object(sync.background, "enqueue_or_run") as fallback:
			event.subject = "Standup, moved"
			event.save(ignore_permissions=True)

			# Not queued, and not run yet either — the row is not written until the commit, and
			# pushing before then would send Microsoft the record as it was.
			patches = [c for c in enqueued.call_args_list if "patch_event_in_graph" in str(c)]
			self.assertFalse(patches, "a dead queue must not be handed the work")
			fallback.assert_not_called()

			frappe.db.commit()  # nosemgrep

		# By name, not by count: any other after-commit work the request happens to carry is
		# not this test's business.
		pushed = [c for c in fallback.call_args_list if "patch_event_in_graph" in str(c)]
		self.assertEqual(len(pushed), 1, fallback.call_args_list)
		self.assertIn(event.name, str(pushed[0]))

	def test_the_work_is_dropped_when_the_save_is_rolled_back(self):
		"""The queued path uses enqueue_after_commit for this reason; the fallback must match."""
		event = self._linked_event("AAMk-dead-queue-2")
		frappe.db.commit()  # nosemgrep

		with patch.object(sync.background, "health", return_value=DEAD_QUEUE), patch.object(
			frappe, "enqueue"
		), patch.object(sync.background, "enqueue_or_run") as fallback:
			event.subject = "Standup, rolled back"
			event.save(ignore_permissions=True)
			frappe.db.rollback()

		self.assertFalse(
			[c for c in fallback.call_args_list if "patch_event_in_graph" in str(c)],
			"a rolled-back save must not reach Microsoft",
		)

	def test_a_healthy_queue_is_still_used(self):
		event = self._linked_event("AAMk-live-queue-1")

		with patch.object(sync.background, "health", return_value=HEALTHY_QUEUE), patch.object(
			frappe, "enqueue"
		) as enqueued, patch.object(sync.background, "enqueue_or_run") as fallback:
			event.subject = "Standup, queued"
			event.save(ignore_permissions=True)
			frappe.db.commit()  # nosemgrep

		self.assertTrue([c for c in enqueued.call_args_list if "patch_event_in_graph" in str(c)])
		fallback.assert_not_called()


class TestSyncNowRefusesToLie(SyncTestCase):
	"""A sync is the one job that must not silently fall back to running inline.

	It is queued precisely because it was judged too slow to hold a request open, so doing it
	here anyway would trade a silent failure for a certain gateway timeout. The person is told
	instead, and offered the choice.
	"""

	def _endpoint(self):
		from frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar import (
			microsoft_calendar as doctype,
		)

		return doctype

	def test_a_dead_queue_is_reported_instead_of_queued(self):
		doctype = self._endpoint()

		with patch.object(sync, "background_reasons", return_value=["This is a first sync"]), patch.object(
			sync.background, "health", return_value=DEAD_QUEUE
		), patch.object(sync, "enqueue_sync") as queued:
			result = doctype.sync(CALENDAR)

		queued.assert_not_called()
		self.assertTrue(result["blocked"])
		self.assertFalse(result["queued"])
		self.assertIn("worker", " ".join(result["health"]["reasons"]).lower())

	def test_run_inline_does_the_work_the_person_asked_for(self):
		"""Pressed after being told the queue is dead: their choice, so slowness is not ours."""
		doctype = self._endpoint()

		with patch.object(sync, "sync_calendar", return_value={"ok": True}) as ran, patch.object(
			sync, "background_reasons"
		) as reasons:
			result = doctype.sync(CALENDAR, run_inline=1)

		ran.assert_called_once_with(CALENDAR)
		reasons.assert_not_called()
		self.assertTrue(result["ok"])

	def test_a_healthy_queue_still_queues(self):
		doctype = self._endpoint()

		with patch.object(sync, "background_reasons", return_value=["This is a first sync"]), patch.object(
			sync.background, "health", return_value=HEALTHY_QUEUE
		), patch.object(sync, "enqueue_sync", return_value={"queued": True, "job_id": "x"}):
			result = doctype.sync(CALENDAR)

		self.assertTrue(result["queued"])
		self.assertNotIn("blocked", result)


class TestEditingDoesNotReplaceTheMeeting(SyncTestCase):
	"""Microsoft reads isOnlineMeeting on a PATCH as "make one", not "keep one".

	Found on a real meeting: it was recorded, then the Event was edited a couple of times, and
	each save minted a fresh Teams meeting with a new join URL and a new online meeting id. The
	recording stayed with the meeting that was actually held, which nothing could name any
	more — and everyone holding the original invitation had a dead link.
	"""

	def test_creating_asks_for_a_meeting(self):
		event = frappe._dict(
			subject="Kickoff",
			starts_on=add_to_date(now_datetime(), hours=1),
			ends_on=add_to_date(now_datetime(), hours=2),
			custom_add_teams_meeting=1,
			custom_teams_join_url=None,
		)

		body = sync._event_to_graph_body(event)

		self.assertTrue(body["isOnlineMeeting"])
		self.assertEqual(body["onlineMeetingProvider"], "teamsForBusiness")

	def test_editing_an_event_that_already_has_one_does_not_ask_again(self):
		event = frappe._dict(
			subject="Kickoff, moved",
			starts_on=add_to_date(now_datetime(), hours=1),
			ends_on=add_to_date(now_datetime(), hours=2),
			custom_add_teams_meeting=1,
			custom_teams_join_url="https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc/0",
		)

		body = sync._event_to_graph_body(event)

		self.assertNotIn("isOnlineMeeting", body, "re-asserting it mints a replacement meeting")
		self.assertNotIn("onlineMeetingProvider", body)

	def test_ticking_the_box_later_still_creates_one(self):
		"""An event Outlook already knows about, with the box ticked after the fact."""
		event = frappe._dict(
			subject="Kickoff",
			starts_on=add_to_date(now_datetime(), hours=1),
			ends_on=add_to_date(now_datetime(), hours=2),
			custom_add_teams_meeting=1,
			custom_teams_join_url="",
			custom_microsoft_event_id="AAMk-existing",
		)

		self.assertTrue(sync._event_to_graph_body(event)["isOnlineMeeting"])

	def test_an_ordinary_event_never_asks_for_one(self):
		event = frappe._dict(
			subject="Lunch",
			starts_on=add_to_date(now_datetime(), hours=1),
			ends_on=add_to_date(now_datetime(), hours=2),
			custom_add_teams_meeting=0,
			custom_teams_join_url=None,
		)

		self.assertNotIn("isOnlineMeeting", sync._event_to_graph_body(event))


class TestPulledEventsBelongToTheirPerson(SyncTestCase):
	"""The multi-user case, which a single-admin site cannot reveal.

	Pulled events are Private, and Frappe shows a private event to its owner, the people it is
	shared with, and its participants. Frappe stamps owner with whoever is running — and the
	scheduled pass runs as Administrator — so every event pulled on a schedule was owned by
	Administrator and invisible to the person whose calendar it came from. Pressing Sync Now
	made the same event visible, because that ran as them. Nothing looks wrong until a second
	person connects a calendar.
	"""

	def setUp(self):
		super().setUp()
		self.person = "m365-owner-test@example.com"
		if not frappe.db.exists("User", self.person):
			frappe.get_doc({
				"doctype": "User", "email": self.person, "first_name": "Calendar",
				"send_welcome_email": 0, "roles": [{"role": "Desk User"}],
			}).insert(ignore_permissions=True)
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "user", self.person)
		frappe.clear_document_cache("Microsoft Calendar", CALENDAR)
		self.addCleanup(frappe.clear_document_cache, "Microsoft Calendar", CALENDAR)
		self.calendar.reload()

	def _pull_one(self):
		frappe.flags.in_microsoft_sync = True
		self.addCleanup(lambda: frappe.flags.pop("in_microsoft_sync", None))
		sync._upsert_event(self.calendar, {
			"id": "AAMk-owner-1", "subject": "Their meeting",
			"start": {"dateTime": "2026-10-01T10:00:00.0000000", "timeZone": "UTC"},
			"end": {"dateTime": "2026-10-01T11:00:00.0000000", "timeZone": "UTC"},
			"isAllDay": False, "type": "singleInstance", "attendees": [],
			"organizer": {"emailAddress": {"address": "org@example.com"}},
			"responseStatus": {"response": "accepted"},
		})
		return frappe.db.get_value("Event", {"custom_microsoft_event_id": "AAMk-owner-1"}, "name")

	def test_the_event_belongs_to_whose_calendar_it_came_from(self):
		"""Not to Administrator, who merely happened to be running the scheduler."""
		name = self._pull_one()

		self.assertEqual(frappe.db.get_value("Event", name, "owner"), self.person)

	def test_that_person_can_actually_see_it(self):
		"""Ownership is the mechanism, visibility is the point — so assert the point."""
		name = self._pull_one()

		self.assertTrue(frappe.has_permission("Event", doc=name, user=self.person))

	def test_a_connection_with_no_user_is_left_alone(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "user", None)
		frappe.clear_document_cache("Microsoft Calendar", CALENDAR)
		self.calendar.reload()

		name = self._pull_one()

		self.assertTrue(frappe.db.get_value("Event", name, "owner"))
