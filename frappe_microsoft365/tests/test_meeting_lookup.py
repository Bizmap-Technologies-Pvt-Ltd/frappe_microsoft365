"""Finding the online meeting behind a join link, as the request Graph would receive.

Only the network is faked (tests/network.py): graph_url's encoding and the URL guard run for
real, so the URLs asserted here are the ones that would go out.
"""

from unittest.mock import patch
from urllib.parse import parse_qsl, urlsplit

import frappe
from requests.utils import requote_uri

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365 import microsoft_transcripts as ms
from frappe_microsoft365.microsoft_meetings import _resolve_online_meeting_id
from frappe_microsoft365.tests.network import Response, error, fake_graph
from frappe_microsoft365.tests.test_meeting_artifacts import CALENDAR, ArtifactsTestCase

OLD_JOIN = (
	"https://teams.microsoft.com/l/meetup-join/19%3ameeting_NjA4ZTk2%40thread.v2/0"
	"?context=%7b%22Tid%22%3a%2211111111%22%2c%22Oid%22%3a%22aaaa%22%7d"
)
NEW_JOIN = "https://teams.microsoft.com/meet/29384756102938?p=AbCd+Ef&x=1#top"

FOUND = {"value": [{"id": "meeting-1"}]}
NOTHING = {"value": []}


def _filter_sent(url):
	"""The $filter Graph decodes out of a URL, which is what it compares with the meeting."""
	return dict(parse_qsl(urlsplit(url).query, keep_blank_values=True))["$filter"]


class TestTheJoinLinkFilter(ArtifactsTestCase):
	def test_an_old_meetup_join_link_is_sent_exactly_as_before(self):
		"""Production resolves these as 1.1.2 sends them, so not one byte of that may change."""
		with fake_graph(lambda *a: Response(200, FOUND)) as calls:
			self.assertEqual(_resolve_online_meeting_id(CALENDAR, OLD_JOIN), "meeting-1")

		[(method, url, kwargs)] = calls
		self.assertEqual(method, "GET")
		self.assertEqual(
			url,
			requote_uri(f"{graph.GRAPH_BASE}/me/onlineMeetings?$filter=JoinWebUrl eq '{OLD_JOIN}'"),
		)
		self.assertIsNone(kwargs["json"])
		self.assertEqual(kwargs["headers"]["Authorization"], "Bearer token")

	def test_ampersand_plus_and_hash_reach_graph_inside_the_literal(self):
		"""Raw, ``&`` would start a second parameter, ``+`` would become a space and ``#`` would
		cut the URL off; Graph would then be comparing the meeting with half a link."""
		with fake_graph(lambda *a: Response(200, FOUND)) as calls:
			_resolve_online_meeting_id(CALENDAR, NEW_JOIN)

		[(_method, url, _kwargs)] = calls
		self.assertEqual(
			url,
			"https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=JoinWebUrl%20eq%20"
			"'https://teams.microsoft.com/meet/29384756102938?p=AbCd%2BEf%26x=1%23top'",
		)
		self.assertEqual(urlsplit(url).fragment, "")
		self.assertEqual(_filter_sent(url), f"JoinWebUrl eq '{NEW_JOIN}'")

	def test_a_quote_in_the_link_is_doubled_for_odata(self):
		with fake_graph(lambda *a: Response(200, FOUND)) as calls:
			_resolve_online_meeting_id(CALENDAR, "https://teams.microsoft.com/meet/1?p=it's")

		self.assertIn("p=it''s'", calls[0][1])


class TestTheFallbacks(ArtifactsTestCase):
	def test_a_match_costs_one_call(self):
		with fake_graph(lambda *a: Response(200, FOUND)) as calls:
			_resolve_online_meeting_id(CALENDAR, NEW_JOIN, event_id="AAMk-1")

		self.assertEqual(len(calls), 1)

	def test_the_outlook_events_own_link_is_tried_next(self):
		outlook_link = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_OUTLOOK%40thread.v2/0"

		def router(method, url, kwargs):
			if "/me/events/" in url:
				return Response(200, {"onlineMeeting": {"joinUrl": outlook_link}})
			return Response(200, FOUND if "OUTLOOK" in url else NOTHING)

		with fake_graph(router) as calls:
			found = _resolve_online_meeting_id(CALENDAR, NEW_JOIN, event_id="AAMk-1")

		self.assertEqual(found, "meeting-1")
		self.assertEqual(
			calls[1][1], "https://graph.microsoft.com/v1.0/me/events/AAMk-1?$select=onlineMeeting"
		)
		self.assertIn("meeting_OUTLOOK", calls[2][1])

	def test_a_new_style_link_is_tried_by_its_meeting_id_last(self):
		def router(method, url, kwargs):
			return Response(200, FOUND if "joinMeetingId" in url else NOTHING)

		with fake_graph(router) as calls:
			found = _resolve_online_meeting_id(CALENDAR, NEW_JOIN)

		self.assertEqual(found, "meeting-1")
		self.assertEqual(
			calls[-1][1],
			"https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter="
			"joinMeetingIdSettings/joinMeetingId%20eq%20'29384756102938'",
		)

	def test_a_meeting_nobody_can_find_is_none(self):
		with fake_graph(lambda *a: Response(200, NOTHING)) as calls:
			self.assertIsNone(_resolve_online_meeting_id(CALENDAR, OLD_JOIN, event_id="AAMk-1"))

		# The join link, then the Outlook event (which answered without a link). An old
		# meetup-join link carries no numeric id to try.
		self.assertEqual(len(calls), 2)

	def test_an_outlook_event_that_is_gone_is_not_an_error(self):
		def router(method, url, kwargs):
			if "/me/events/" in url:
				return error(404, "ErrorItemNotFound")
			return Response(200, NOTHING)

		with fake_graph(router):
			self.assertIsNone(_resolve_online_meeting_id(CALENDAR, OLD_JOIN, event_id="AAMk-1"))

	def test_an_outage_is_raised_not_reported_as_no_match(self):
		"""A 503 says nothing about whether the meeting exists, so it must not read as 'foreign'."""
		with fake_graph(lambda *a: error(503, "ServiceUnavailable")):
			with self.assertRaises(graph.MsGraphError):
				_resolve_online_meeting_id(CALENDAR, OLD_JOIN)


class TestTheLookupPermission(ArtifactsTestCase):
	"""Transcripts ask for OnlineMeetings.Read, except where ReadWrite was already granted."""

	TRANSCRIPTS = {"use_calendar": 1, "use_transcripts": 1}
	BEFORE_120 = (
		"User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite "
		"OnlineMeetingTranscript.Read.All OnlineMeetingRecording.Read.All"
	)

	def _granted(self, value):
		previous = frappe.db.get_single_value("Microsoft Settings", "authorized_scopes")
		frappe.db.set_single_value("Microsoft Settings", "authorized_scopes", value)
		self.addCleanup(frappe.db.set_single_value, "Microsoft Settings", "authorized_scopes", previous)

	def test_a_new_setup_with_transcripts_alone_asks_for_read(self):
		self.assertEqual(
			graph.derive_scopes(self.TRANSCRIPTS),
			[
				"User.Read",
				"Calendars.ReadWrite",
				"OnlineMeetings.Read",
				"OnlineMeetingTranscript.Read.All",
				"OnlineMeetingRecording.Read.All",
			],
		)

	def test_a_site_already_granted_readwrite_keeps_asking_for_it(self):
		"""Swapping it for Read would be a scope nobody consented to: a prompt for everyone."""
		scopes = graph.derive_scopes(self.TRANSCRIPTS, graph.parse_scopes(self.BEFORE_120))

		self.assertIn("OnlineMeetings.ReadWrite", scopes)
		self.assertNotIn("OnlineMeetings.Read", scopes)

	def test_an_existing_site_requests_exactly_what_it_was_authorised_with(self):
		"""The 1.1.x setup (standalone meetings + transcripts) asks for the identical list, so
		the doctor sees no scope change and nobody is told to re-authorise."""
		self._granted(self.BEFORE_120)
		ticked = frappe._dict(use_calendar=1, use_teams=1, use_transcripts=1, default_scopes="")

		self.assertEqual(graph.get_scopes(ticked), graph.parse_scopes(self.BEFORE_120))

	def test_unticking_standalone_meetings_later_still_keeps_readwrite(self):
		self._granted(self.BEFORE_120)
		ticked = frappe._dict(self.TRANSCRIPTS, default_scopes="")

		self.assertEqual(graph.get_scopes(ticked), graph.parse_scopes(self.BEFORE_120))

	def test_standalone_meetings_ask_for_readwrite_and_never_both(self):
		scopes = graph.derive_scopes({"use_teams": 1, "use_transcripts": 1})

		self.assertIn("OnlineMeetings.ReadWrite", scopes)
		self.assertNotIn("OnlineMeetings.Read", scopes)

	def test_the_sign_in_url_of_an_existing_site_is_unchanged(self):
		self._granted(self.BEFORE_120)
		ticked = frappe._dict(use_calendar=1, use_teams=0, use_transcripts=1, default_scopes="")

		with patch.object(graph, "get_settings", return_value=ticked), patch.object(
			graph, "_settings_doc", return_value=ticked
		):
			ticked.update(client_id="cid", tenant_id="contoso.example", redirect_uri="https://x/cb")
			url = graph.build_authorize_url("state")

		scope = dict(parse_qsl(urlsplit(url).query))["scope"].split()
		self.assertIn("OnlineMeetings.ReadWrite", scope)
		self.assertNotIn("OnlineMeetings.Read", scope)

	def test_granted_scopes_also_reads_per_connection_storage(self):
		"""Whichever of the two places holds them, the answer is the same."""
		self._granted("")
		with patch.object(frappe.db, "has_column", return_value=True), patch.object(
			frappe, "get_all", return_value=["User.Read OnlineMeetings.ReadWrite", None]
		):
			self.assertEqual(graph.granted_scopes(), ["User.Read", "OnlineMeetings.ReadWrite"])

	def test_the_join_link_403_names_the_permission_in_use(self):
		self._granted("")
		with patch.object(graph, "_settings_doc", return_value=frappe._dict(self.TRANSCRIPTS)):
			self.assertIn("OnlineMeetings.Read,", ms._online_meetings_hint())
		self._granted(self.BEFORE_120)
		with patch.object(graph, "_settings_doc", return_value=frappe._dict(self.TRANSCRIPTS)):
			self.assertIn("OnlineMeetings.ReadWrite", ms._online_meetings_hint())

	def test_the_doctor_says_which_one_is_in_use(self):
		from frappe_microsoft365 import doctor

		new = doctor.check_meeting_lookup_scope(self.TRANSCRIPTS, graph.derive_scopes(self.TRANSCRIPTS))
		kept = doctor.check_meeting_lookup_scope(
			self.TRANSCRIPTS, graph.derive_scopes(self.TRANSCRIPTS, ["OnlineMeetings.ReadWrite"])
		)

		self.assertEqual(new[0]["title"], "Transcripts find their meeting with OnlineMeetings.Read")
		self.assertEqual(kept[0]["title"], "Transcripts find their meeting with OnlineMeetings.ReadWrite")
		self.assertIn("already granted", kept[0]["detail"])
		self.assertEqual(doctor.check_meeting_lookup_scope({"use_calendar": 1}, ["User.Read"]), [])

	def test_an_override_with_readwrite_is_not_told_read_is_missing(self):
		from frappe_microsoft365 import doctor

		found = doctor.check_scopes(
			{**self.TRANSCRIPTS, "default_scopes": self.BEFORE_120}
		)

		self.assertEqual(found, [])
