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
from frappe_microsoft365.tests.test_files import MAPPED, FilesTestCase, configure, upgrade_checks


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


# --- S8: a kept local copy follows the File's privacy and leaves with it ----------------------


class KeptCopyTestCase(FilesTestCase):
	def setUp(self):
		super().setUp()
		configure(self, files_keep_local_copy=1)

	def moved_public(self, content=None, name=None, shared_with=False):
		"""A public attachment moved with its local copy kept; optionally a second, unmoved File
		that shares its bytes (Frappe stores identical content once)."""
		content = content or f"kept {frappe.generate_hash()}".encode()
		name = name or f"kept-{frappe.generate_hash(length=6)}.txt"
		f = self.attach(content=content, name=name, is_private=0)
		local = f.file_url
		self.assertTrue(local.startswith("/files/"))
		if shared_with:
			other = frappe.get_doc({"doctype": "File", "file_name": name, "content": content, "is_private": 0}).insert()
			self.assertEqual(other.file_url, local)
		files.upload_file(f.name)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.STORED)
		for url in (local, "/private" + local):
			self.addCleanup(lambda p=files.local_path(url): os.path.exists(p) and os.remove(p))
		return f, local

	def kept(self, f):
		return frappe.db.get_value("File", f.name, "custom_microsoft_local_url")


class TestKeptCopyFollowsPrivacy(KeptCopyTestCase):
	def test_making_a_moved_file_private_moves_its_kept_copy_out_of_public_files(self):
		f, local = self.moved_public()
		doc = frappe.get_doc("File", f.name)
		doc.is_private = 1
		doc.save()
		self.assertFalse(os.path.exists(files.local_path(local)))
		self.assertEqual(self.kept(f), "/private" + local)
		self.assertTrue(os.path.exists(files.local_path("/private" + local)))

	def test_bring_back_restores_a_private_file_to_private_files(self):
		f, local = self.moved_public()
		# As 1.1.x left it: made private without the copy following.
		frappe.db.set_value("File", f.name, "is_private", 1)
		calls = len(self.graph.calls)
		target = files.bring_back(f.name)
		self.assertEqual(target, "/private" + local)
		self.assertEqual(frappe.db.get_value("File", f.name, "file_url"), target)
		self.assertFalse(os.path.exists(files.local_path(local)))
		self.assertEqual(self.graph.calls[calls:], [])  # the kept copy was reused, not downloaded

	def test_bytes_another_file_uses_are_copied_not_taken(self):
		f, local = self.moved_public(shared_with=True)
		doc = frappe.get_doc("File", f.name)
		doc.is_private = 1
		doc.save()
		self.assertTrue(os.path.exists(files.local_path(local)))  # still the other File's
		self.assertTrue(self.kept(f).startswith("/private/files/"))


class TestKeptCopyLeavesWithItsFile(KeptCopyTestCase):
	def test_deleting_a_moved_file_removes_its_kept_copy(self):
		f, local = self.moved_public()
		frappe.delete_doc("File", f.name)
		self.assertFalse(os.path.exists(files.local_path(local)))

	def test_a_kept_copy_another_file_still_uses_stays(self):
		f, local = self.moved_public(shared_with=True)
		frappe.delete_doc("File", f.name)
		self.assertTrue(os.path.exists(files.local_path(local)))

	def test_bytes_that_are_no_longer_the_files_are_never_removed(self):
		f, local = self.moved_public()
		with open(files.local_path(local), "wb") as fh:
			fh.write(b"someone else's file now")
		frappe.delete_doc("File", f.name)
		self.assertTrue(os.path.exists(files.local_path(local)))


class TestKeptCopyPatchAndCheck(KeptCopyTestCase):
	def patch(self):
		from frappe_microsoft365.patches import move_public_kept_copies_of_private_files

		move_public_kept_copies_of_private_files.execute()

	def test_the_patch_moves_public_copies_of_private_files_and_can_run_twice(self):
		f, local = self.moved_public()
		frappe.db.set_value("File", f.name, "is_private", 1)  # as 1.1.x left it
		self.patch()
		self.assertEqual(self.kept(f), "/private" + local)
		self.assertFalse(os.path.exists(files.local_path(local)))
		self.patch()
		self.assertEqual(self.kept(f), "/private" + local)
		self.assertTrue(os.path.exists(files.local_path("/private" + local)))

	def test_the_patch_leaves_public_files_alone(self):
		f, local = self.moved_public()
		self.patch()
		self.assertEqual(self.kept(f), local)
		self.assertTrue(os.path.exists(files.local_path(local)))

	def test_the_patch_does_nothing_where_storage_was_never_switched_on(self):
		f, local = self.moved_public()
		frappe.db.set_value("File", f.name, "is_private", 1)
		with patch.object(frappe.db, "has_column", return_value=False):
			self.patch()
		self.assertEqual(self.kept(f), local)

	def deleted_as_before(self, f):
		"""Delete a moved File the way 1.1.x did: its kept copy stays on disk."""
		with patch.object(files, "_drop_kept_copy"):
			frappe.delete_doc("File", f.name)

	def test_kept_copies_of_deleted_files_are_reported_not_deleted(self):
		f, local = self.moved_public()
		self.deleted_as_before(f)
		found = files.upgrade_check_orphaned_kept_copies()
		self.assertEqual(len(found), 1)
		self.assertEqual(found[0]["check"], "upgrade.sharepoint_orphaned_kept_copies")
		self.assertIn(local, found[0]["detail"])
		self.assertTrue(os.path.exists(files.local_path(local)))

	def test_a_copy_removed_since_or_used_by_another_file_is_not_reported(self):
		f, local = self.moved_public()
		self.deleted_as_before(f)
		os.remove(files.local_path(local))
		self.assertEqual(files.upgrade_check_orphaned_kept_copies(), [])

	def test_the_check_is_registered_for_the_upgrade_report(self):
		self.assertIn(
			"frappe_microsoft365.microsoft_files.upgrade_check_orphaned_kept_copies", upgrade_checks()
		)
