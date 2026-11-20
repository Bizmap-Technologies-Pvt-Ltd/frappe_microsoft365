"""Transcript content: the plain-text fallback for tenants without speaker attribution.

Graph is mocked at graph_request, so what is asserted is the request this app actually sends —
$format for VTT, the Accept header (and never $format) for the plain-text format.
"""

from unittest.mock import patch

import frappe

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365 import microsoft_meeting_artifacts as artifacts
from frappe_microsoft365 import microsoft_transcripts as ms
from frappe_microsoft365.microsoft_graph import MsGraphError
from frappe_microsoft365.tests.test_meeting_artifacts import CALENDAR, VTT, ArtifactsTestCase

PATH = "/me/onlineMeetings/meeting-1/transcripts/t1/content"
PLAIN = "00:00:01.000 --> 00:00:04.000\nLet's begin, Zoë.\n"

#: What Graph sends, as this app's error text carries it: the outer code and the message. The
#: SpeakerAttributionNotAllowed code itself sits in innerError, which that text leaves out.
ATTRIBUTION_REFUSAL = (
	f"Microsoft Graph GET {PATH} failed (403): Forbidden: Speaker-attributed transcript content "
	"is disabled for this tenant. Retry with Accept 'application/vnd.microsoft.graph.transcript+text'."
)
TENANT_SWITCH_REFUSAL = f"Microsoft Graph GET {PATH} failed (403): Forbidden: GraphAccessToTranscriptsDisabled"
BARE_403 = f"Microsoft Graph GET {PATH} failed (403): Forbidden"


class FakeResponse:
	def __init__(self, body, content_type):
		self.content = body.encode("utf-8")
		self.headers = {"Content-Type": content_type}

	@property
	def text(self):
		# What requests would do for a text/* type without a charset; reading it is the bug.
		return self.content.decode("latin-1")


def _graph(*answers):
	"""Patch graph_request to answer each call in turn: a FakeResponse, or an error to raise."""

	pending = list(answers)

	def side_effect(*args, **kwargs):
		answer = pending.pop(0)
		if isinstance(answer, Exception):
			raise answer
		return answer

	return patch.object(graph, "graph_request", side_effect=side_effect)


class TestTranscriptContent(ArtifactsTestCase):
	def _get(self, *answers):
		with _graph(*answers) as calls:
			try:
				return ms.get_transcript_content(CALENDAR, "meeting-1", "t1"), calls
			except Exception as e:
				e.calls = calls
				raise

	def test_vtt_is_still_asked_for_first_with_format(self):
		result, calls = self._get(FakeResponse(VTT, "text/vtt"))

		self.assertEqual(result, {"format": "text/vtt", "content": VTT})
		self.assertEqual(calls.call_count, 1)
		self.assertEqual(calls.call_args.kwargs.get("params"), {"$format": "text/vtt"})

	def test_speaker_attribution_refusal_retries_with_the_plain_text_accept_header(self):
		result, calls = self._get(
			MsGraphError(ATTRIBUTION_REFUSAL), FakeResponse(PLAIN, ms.PLAIN_TEXT_FORMAT)
		)

		self.assertEqual(calls.call_count, 2)
		retry = calls.call_args_list[1]
		self.assertEqual(retry.args[1], PATH)
		self.assertEqual(retry.kwargs.get("headers"), {"Accept": ms.PLAIN_TEXT_FORMAT})
		self.assertFalse(retry.kwargs.get("params"), "the plain-text format is header-only")
		self.assertEqual(result["format"], ms.PLAIN_TEXT_FORMAT)
		self.assertEqual(result["content"], PLAIN, "decoded as UTF-8, not guessed")

	def test_the_literal_inner_error_code_also_triggers_the_retry(self):
		result, calls = self._get(
			MsGraphError(f"{BARE_403}: SpeakerAttributionNotAllowed"),
			FakeResponse(PLAIN, ms.PLAIN_TEXT_FORMAT),
		)

		self.assertEqual(calls.call_count, 2)
		self.assertEqual(result["format"], ms.PLAIN_TEXT_FORMAT)

	def test_the_tenant_switch_is_not_retried(self):
		"""It refuses every format alike, so a second request only doubles the wait."""
		with self.assertRaises(frappe.ValidationError) as ctx:
			self._get(MsGraphError(TENANT_SWITCH_REFUSAL), FakeResponse(PLAIN, ms.PLAIN_TEXT_FORMAT))

		self.assertEqual(ctx.exception.calls.call_count, 1)
		self.assertIn("Transcript API access", str(ctx.exception))

	def test_a_404_is_not_retried(self):
		with self.assertRaises(graph.MsGraphNotFound) as ctx:
			self._get(graph.MsGraphNotFound(f"Microsoft Graph GET {PATH} failed (404): NotFound"))

		self.assertEqual(ctx.exception.calls.call_count, 1)

	def test_a_failed_fallback_is_explained_by_its_own_refusal(self):
		"""Attribution never blocks the plain-text format, so naming it then would mislead."""
		with self.assertRaises(frappe.ValidationError) as ctx:
			self._get(MsGraphError(ATTRIBUTION_REFUSAL), MsGraphError(BARE_403))

		said = str(ctx.exception)
		self.assertEqual(ctx.exception.calls.call_count, 2)
		self.assertIn("OnlineMeetingTranscript.Read.All", said)
		self.assertNotIn("Include speaker attribution", said)

	def test_asking_for_plain_text_directly_uses_the_header(self):
		with _graph(FakeResponse(PLAIN, ms.PLAIN_TEXT_FORMAT)) as calls:
			result = ms.get_transcript_content(CALENDAR, "meeting-1", "t1", fmt=ms.PLAIN_TEXT_FORMAT)

		self.assertEqual(calls.call_args.kwargs.get("headers"), {"Accept": ms.PLAIN_TEXT_FORMAT})
		self.assertFalse(calls.call_args.kwargs.get("params"))
		self.assertEqual(result["format"], ms.PLAIN_TEXT_FORMAT)

	def test_join_url_lookup_says_which_format_it_returns(self):
		with patch.object(ms, "_resolve_online_meeting_id", return_value="meeting-1"), patch.object(
			ms, "list_transcripts", return_value=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}]
		), _graph(MsGraphError(ATTRIBUTION_REFUSAL), FakeResponse(PLAIN, ms.PLAIN_TEXT_FORMAT)):
			result = ms.get_transcripts_for_join_url(CALENDAR, "https://teams.microsoft.com/l/x")

		self.assertEqual(result["latest_vtt"], PLAIN)
		self.assertEqual(result["latest_format"], ms.PLAIN_TEXT_FORMAT)


class TestPlainTextTranscriptIsAttached(ArtifactsTestCase):
	def _files(self, event):
		return sorted(
			frappe.get_all(
				"File",
				filters={"attached_to_doctype": "Event", "attached_to_name": event.name},
				pluck="file_name",
			)
		)

	def assertNamed(self, names, expected):
		"""Each name starts with its expected stem and keeps the extension. Frappe adds a short
		hash before the extension when another File already took the name, so exact names would
		depend on what earlier tests left behind."""
		self.assertEqual(len(names), len(expected), names)
		for name, (stem, extension) in zip(names, expected):
			self.assertTrue(name.startswith(stem) and name.endswith(extension), (name, stem, extension))

	def _fetch_with_attribution_off(self, event, transcripts):
		"""Every part refused as VTT, then served as plain text."""
		answers = []
		for transcript in transcripts:
			# Distinct bodies: Frappe renames a File whose content matches one already stored.
			body = f"{PLAIN}{transcript['id']}\n"
			answers += [MsGraphError(ATTRIBUTION_REFUSAL), FakeResponse(body, ms.PLAIN_TEXT_FORMAT)]
		with patch.object(ms, "list_transcripts", return_value=transcripts), patch.object(
			ms, "list_recordings", return_value=[]
		), _graph(*answers):
			return artifacts.fetch_meeting_artifacts(event.name)

	def test_it_is_stored_as_txt_and_the_fetch_counts_as_done(self):
		event = self._finished_meeting()

		result = self._fetch_with_attribution_off(
			event, [{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}]
		)

		self.assertEqual(len(result["transcripts"]), 1)
		self.assertNamed(self._files(event), [("teams-transcript-2026-09-12", ".txt")])
		event.reload()
		self.assertTrue(event.custom_microsoft_transcript_fetched_on)

	def test_several_parts_keep_their_part_names(self):
		event = self._finished_meeting()

		self._fetch_with_attribution_off(
			event,
			[
				{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"},
				{"id": "t2", "created_date_time": "2026-09-12T12:00:00Z"},
			],
		)

		self.assertNamed(
			self._files(event),
			[("teams-transcript-2026-09-12-part-1", ".txt"), ("teams-transcript-2026-09-12-part-2", ".txt")],
		)

	def test_a_refetch_replaces_an_earlier_vtt(self):
		"""The same prefix covers both extensions, so turning attribution off does not leave the
		old .vtt next to the new .txt."""
		event = self._finished_meeting()
		self._fetch(event, transcripts=[{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}])

		self._fetch_with_attribution_off(event, [{"id": "t1", "created_date_time": "2026-09-12T11:00:00Z"}])

		self.assertNamed(self._files(event), [("teams-transcript-2026-09-12", ".txt")])

	def test_a_content_answer_without_a_format_is_still_vtt(self):
		"""Callers and mocks that predate the format key keep getting .vtt."""
		self.assertEqual(artifacts._transcript_extension(None), ".vtt")
		self.assertEqual(artifacts._transcript_extension("text/vtt"), ".vtt")
		self.assertEqual(artifacts._transcript_extension(ms.PLAIN_TEXT_FORMAT), ".txt")


class TestDoctorReadsMicrosoftsWording(ArtifactsTestCase):
	def test_the_message_without_the_inner_code_is_still_recognised(self):
		from frappe_microsoft365 import doctor

		explained = doctor.explain_error(ATTRIBUTION_REFUSAL)

		self.assertTrue(explained["matched"])
		self.assertIn("speaker attribution", explained["title"].lower())
		self.assertIn(".txt", explained["detail"])
