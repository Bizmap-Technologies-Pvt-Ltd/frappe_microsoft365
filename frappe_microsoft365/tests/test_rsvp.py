"""RSVP tests — Graph is mocked at ``microsoft_graph.graph_request``; no network, no tokens.

Every action is checked by the endpoint it posts to. The three replies differ only in the
last path segment, one of them is camelCase, and picking the wrong one is a 404 that reaches
the user as "the Microsoft integration is broken".
"""

from unittest.mock import patch

import frappe

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365 import microsoft_rsvp as rsvp
from frappe_microsoft365.tests.base import BaseTestCase

CALENDAR = "_Test MS RSVP Calendar"
MS_ID = "ms-invitation-1"


class RsvpTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		self._cleanup()
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
		# An invitation as the pull leaves it: mirrored from Outlook, not yet answered.
		self.event = frappe.get_doc(
			{
				"doctype": "Event",
				"subject": "Quarterly review",
				"starts_on": "2026-09-15 09:00:00",
				"ends_on": "2026-09-15 10:00:00",
				"event_type": "Private",
				"custom_sync_with_microsoft_calendar": 1,
				"custom_microsoft_calendar": CALENDAR,
				"custom_microsoft_event_id": MS_ID,
				"custom_pulled_from_microsoft": 1,
				"custom_microsoft_organizer": "dana@example.com",
				"custom_microsoft_my_response": "notResponded",
			}
		).insert(ignore_permissions=True)

	def tearDown(self):
		self._cleanup()
		super().tearDown()

	def _cleanup(self):
		for name in frappe.get_all("Event", filters={"custom_microsoft_calendar": CALENDAR}, pluck="name"):
			frappe.delete_doc("Event", name, force=True, ignore_permissions=True)
		if frappe.db.exists("Microsoft Calendar", CALENDAR):
			frappe.delete_doc("Microsoft Calendar", CALENDAR, force=True, ignore_permissions=True)

	def _respond(self, response, **kwargs):
		"""Answer the invitation with Graph mocked. 202 Accepted carries no body, hence {}."""
		with patch.object(graph, "graph_request", return_value={}) as mocked:
			result = rsvp.respond_to_event(self.event.name, response, **kwargs)
		return result, mocked

	def _refused(self, response="accept"):
		"""Assert the call is refused locally, before anything reaches Microsoft."""
		with patch.object(graph, "graph_request") as mocked:
			with self.assertRaises(frappe.ValidationError):
				rsvp.respond_to_event(self.event.name, response)
		mocked.assert_not_called()


class TestRespondToEvent(RsvpTestCase):
	def test_accept_posts_to_the_accept_action(self):
		result, mocked = self._respond("accept")

		self.assertEqual(mocked.call_args[0][0], "POST")
		self.assertEqual(mocked.call_args[0][1], f"/me/events/{MS_ID}/accept")
		self.assertEqual(mocked.call_args[0][2], CALENDAR)
		self.assertEqual(result["response"], "accepted")

	def test_decline_posts_to_the_decline_action(self):
		result, mocked = self._respond("decline")

		self.assertEqual(mocked.call_args[0][1], f"/me/events/{MS_ID}/decline")
		self.assertEqual(result["response"], "declined")

	def test_tentative_posts_to_microsofts_camelcase_action(self):
		"""/tentativelyAccept, not /tentative and not /tentativelyaccept."""
		result, mocked = self._respond("tentative")

		self.assertEqual(mocked.call_args[0][1], f"/me/events/{MS_ID}/tentativelyAccept")
		self.assertEqual(result["response"], "tentativelyAccepted")

	def test_an_empty_202_body_is_not_treated_as_a_failure(self):
		result, _ = self._respond("accept")

		self.assertTrue(result["ok"])

	def test_the_local_response_is_updated_without_waiting_for_the_next_sync(self):
		self._respond("accept")

		self.assertEqual(
			frappe.db.get_value("Event", self.event.name, "custom_microsoft_my_response"), "accepted"
		)

	def test_responding_does_not_bump_modified(self):
		"""A bumped `modified` is exactly what makes the sync re-patch an event forever."""
		before = frappe.db.get_value("Event", self.event.name, "modified")

		self._respond("decline")

		self.assertEqual(before, frappe.db.get_value("Event", self.event.name, "modified"))

	def test_a_comment_is_passed_through_to_graph(self):
		_, mocked = self._respond("decline", comment="Clashes with the board meeting")

		self.assertEqual(mocked.call_args[1]["json"]["comment"], "Clashes with the board meeting")

	def test_no_comment_key_is_sent_when_there_is_no_comment(self):
		_, mocked = self._respond("accept")

		self.assertNotIn("comment", mocked.call_args[1]["json"])

	def test_the_organizer_is_notified_by_default(self):
		_, mocked = self._respond("accept")

		self.assertTrue(mocked.call_args[1]["json"]["sendResponse"])

	def test_send_response_zero_replies_without_emailing_anyone(self):
		"""The tickbox arrives from the form as a string, so cint has to do the work."""
		_, mocked = self._respond("accept", send_response="0")

		self.assertFalse(mocked.call_args[1]["json"]["sendResponse"])


class TestRespondToEventGuards(RsvpTestCase):
	def test_an_unknown_response_is_refused_before_any_graph_call(self):
		self._refused("maybe-later")

	def test_an_empty_response_is_refused(self):
		self._refused("")

	def test_an_event_that_is_not_synced_cannot_be_answered(self):
		frappe.db.set_value("Event", self.event.name, "custom_microsoft_event_id", "")

		self._refused()

	def test_the_organizer_is_told_there_is_nothing_to_reply_to(self):
		"""Graph's own answer here is "ErrorInvalidRequest", which explains nothing."""
		frappe.db.set_value("Event", self.event.name, "custom_microsoft_my_response", "organizer")

		self._refused()

	def test_an_unauthorized_calendar_is_refused(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "authorized", 0)

		self._refused()

	def test_a_disabled_calendar_is_refused(self):
		frappe.db.set_value("Microsoft Calendar", CALENDAR, "enabled", 0)

		self._refused()

	def test_the_calendar_owner_check_runs_before_the_graph_call(self):
		"""Replying writes into a real mailbox, so it is the calendar owner's call alone."""
		with patch.object(rsvp, "_check_owner", side_effect=frappe.PermissionError) as checked, patch.object(
			graph, "graph_request"
		) as mocked:
			with self.assertRaises(frappe.PermissionError):
				rsvp.respond_to_event(self.event.name, "accept")

		checked.assert_called_once()
		mocked.assert_not_called()
