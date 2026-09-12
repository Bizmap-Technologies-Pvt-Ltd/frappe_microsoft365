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
		for name in frappe.get_all(
			"Event", filters={"custom_microsoft_calendar": CALENDAR}, pluck="name"
		):
			frappe.delete_doc("Event", name, force=True, ignore_permissions=True)

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

	def _enable_integration(self, enabled):
		frappe.db.set_single_value("Microsoft Settings", "enabled", enabled)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")
		self.addCleanup(frappe.clear_document_cache, "Microsoft Settings", "Microsoft Settings")

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

	def test_sync_all_covers_enabled_authorized_calendars(self):
		self._enable_integration(1)

		with patch.object(graph, "graph_delta", return_value=([ms_event()], "DELTA-ALL")), patch.object(
			frappe.db, "commit"
		):
			results = sync.sync_all()

		self.assertTrue(any(r.get("pulled") == 1 for r in results or []))


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
