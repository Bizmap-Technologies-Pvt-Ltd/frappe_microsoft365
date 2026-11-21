"""Finding the online meeting behind a join link, as the request Graph would receive.

Only the network is faked (tests/network.py): graph_url's encoding and the URL guard run for
real, so the URLs asserted here are the ones that would go out.
"""

from urllib.parse import parse_qsl, urlsplit

from requests.utils import requote_uri

from frappe_microsoft365 import microsoft_graph as graph
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
