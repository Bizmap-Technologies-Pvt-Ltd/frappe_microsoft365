"""SharePoint document storage, 1.2.0: what happens to a moved file over its life.

Kept local copies following the File's privacy and leaving with it, public files for guests, old
links, re-attachment, retries that must not duplicate, long paths, big files, resumable uploads,
recreated libraries, and link archiving that connects only where it checked.

Nothing here touches the network. Most tests use test_files.FakeGraph in place of graph_request;
the ones marked "network-level" patch only ``microsoft_graph._http_request`` and the app token,
so graph_request, graph_url and their error handling really run.
"""

import io
import json
import os
import socket
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import frappe

from frappe_microsoft365 import microsoft_files as files
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import MsGraphError
from frappe_microsoft365.tests.base import BaseTestCase
from frappe_microsoft365.tests.test_files import MAPPED, FilesTestCase, configure


def setUpModule():
	"""Create the File fields once, before any test switches storage on (see test_files)."""
	from frappe_microsoft365.setup import create_file_custom_fields

	create_file_custom_fields()
	frappe.db.commit()  # nosemgrep


class HttpResponse:
	"""What requests returns, as far as graph_request and the upload code read it."""

	def __init__(self, status=200, payload=None, headers=None, body=None):
		self.status_code = status
		self.reason = "Fake"
		self.headers = headers or {}
		if body is None:
			body = json.dumps(payload).encode() if payload is not None else b""
		self.content = body
		self.closed = False

	def json(self):
		return json.loads(self.content)

	def iter_content(self, chunk_size=1):
		for i in range(0, len(self.content), chunk_size):
			yield self.content[i : i + chunk_size]

	def close(self):
		self.closed = True


class NetworkFake:
	"""Stands in for microsoft_graph._http_request: records (method, url, kwargs), answers in order."""

	def __init__(self, answer):
		self.answer = answer
		self.calls = []

	def __call__(self, method, url, **kwargs):
		self.calls.append((method, url, kwargs))
		return self.answer(method, url, kwargs)


def network(test, answer):
	"""Patch only the transport and the app token; returns the NetworkFake."""
	fake = NetworkFake(answer)
	for target, kwargs in (
		("_http_request", {"side_effect": fake}),
		("get_app_access_token", {"return_value": "APP-TOKEN"}),
	):
		p = patch.object(graph, target, **kwargs)
		p.start()
		test.addCleanup(p.stop)
	return fake


GRAPH = "https://graph.microsoft.com/v1.0"


# --- S9: link archiving connects only to the address it checked ------------------------------


class TestLinkDownloadIsPinned(BaseTestCase):
	HOST = "rebind.example.test"
	PUBLIC = "93.184.216.34"
	PRIVATE = "169.254.169.254"

	def fetch_with_rebinding_dns(self, url):
		answers, connects = [], []
		real = socket.getaddrinfo

		def rebinding(host, port, *a, **k):
			if host == self.HOST:
				ip = self.PUBLIC if not answers else self.PRIVATE
				answers.append(ip)
				return [(socket.AF_INET, socket.SOCK_STREAM, 6, "", (ip, port))]
			return real(host, port, *a, **k)

		def connect(sock, address):
			connects.append(address)
			raise ConnectionRefusedError("no network in tests")

		with (
			patch.object(socket, "getaddrinfo", side_effect=rebinding),
			patch.object(socket.socket, "connect", connect),
		):
			with self.assertRaises(MsGraphError):
				with files._download_link(url):
					pass
		return answers, connects

	def test_a_name_that_answers_public_then_private_is_only_connected_at_the_public_address(self):
		answers, connects = self.fetch_with_rebinding_dns(f"http://{self.HOST}/latest/meta-data/")
		self.assertEqual(answers, [self.PUBLIC])  # resolved once, never again
		self.assertEqual([a[0] for a in connects], [self.PUBLIC])

	def test_https_connects_to_the_address_and_checks_the_certificate_for_the_name(self):
		with patch("urllib3.HTTPSConnectionPool") as pool_class:
			pool_class.return_value.urlopen.return_value = SimpleNamespace(status=200)
			with patch.object(
				socket,
				"getaddrinfo",
				return_value=[(socket.AF_INET, socket.SOCK_STREAM, 6, "", (self.PUBLIC, 443))],
			):
				files._pinned_get(f"https://{self.HOST}/a file.pdf?x=1")
		kwargs = pool_class.call_args.kwargs
		self.assertEqual(kwargs["host"], self.PUBLIC)
		self.assertEqual(kwargs["server_hostname"], self.HOST)
		self.assertEqual(kwargs["assert_hostname"], self.HOST)
		self.assertEqual(kwargs["cert_reqs"], "CERT_REQUIRED")
		call = pool_class.return_value.urlopen.call_args
		self.assertEqual(call.args, ("GET", "/a%20file.pdf?x=1"))
		self.assertEqual(call.kwargs["headers"]["Host"], self.HOST)
		self.assertFalse(call.kwargs["redirect"])

	def test_a_redirect_to_a_private_address_is_refused(self):
		first = MagicMock(status=302, headers={"Location": "http://10.0.0.5/secret"})
		with (
			patch.object(
				socket,
				"getaddrinfo",
				side_effect=lambda host, port, *a, **k: [
					(socket.AF_INET, socket.SOCK_STREAM, 6, "", (host if host[0].isdigit() else self.PUBLIC, port))
				],
			),
			patch("urllib3.HTTPConnectionPool") as pool_class,
		):
			pool_class.return_value.urlopen.return_value = first
			with self.assertRaises(MsGraphError):
				with files._download_link(f"http://{self.HOST}/start"):
					pass
		self.assertEqual(pool_class.call_count, 1)  # never connected to 10.0.0.5
