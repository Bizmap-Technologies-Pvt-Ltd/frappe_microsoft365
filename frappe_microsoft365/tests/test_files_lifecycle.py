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


#: The real graph_request, before any test replaces it with a FakeGraph.
GRAPH_REQUEST = graph.graph_request


def real_graph(test):
	"""Put the real graph_request back under a FilesTestCase, for a network-level test."""
	p = patch.object(graph, "graph_request", GRAPH_REQUEST)
	p.start()
	test.addCleanup(p.stop)


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


def remove(path):
	"""Delete a file a test left on disk; rolling the database back does not."""
	if os.path.exists(path):
		os.remove(path)


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
					(
						socket.AF_INET,
						socket.SOCK_STREAM,
						6,
						"",
						(host if host[0].isdigit() else self.PUBLIC, port),
					)
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
			other = frappe.get_doc(
				{"doctype": "File", "file_name": name, "content": content, "is_private": 0}
			).insert()
			self.assertEqual(other.file_url, local)
		files.upload_file(f.name)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.STORED)
		for url in (local, "/private" + local):
			self.addCleanup(remove, files.local_path(url))
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


# --- B20: public attachments stay public after the move -------------------------------------


class TestPublicFilesStayPublic(FilesTestCase):
	def moved(self, is_private):
		f = self.attach(name=f"logo-{frappe.generate_hash(length=6)}.png", is_private=is_private)
		files.upload_file(f.name)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.STORED)
		return f

	def as_guest(self):
		frappe.set_user("Guest")
		self.addCleanup(frappe.set_user, "Administrator")

	def test_a_guest_can_open_a_moved_public_file(self):
		f = self.moved(is_private=0)
		self.as_guest()
		with patch.object(files, "_stream", return_value="STREAM") as stream:
			self.assertEqual(files.open_file(f.name), "STREAM")
		stream.assert_called_once_with("DRIVE", "ITEM-" + f.file_name, f.file_name, download=0)

	def test_a_guest_still_cannot_open_a_moved_private_file(self):
		f = self.moved(is_private=1)
		self.as_guest()
		with patch.object(files, "_stream") as stream:
			with self.assertRaises(frappe.PermissionError):
				files.open_file(f.name)
		stream.assert_not_called()

	def test_open_file_is_reachable_without_signing_in(self):
		# Without allow_guest Frappe answers a guest with 403 before open_file runs.
		self.assertIn(files.open_file, frappe.guest_methods)


# --- B22: old /files links follow the database row --------------------------------------------


class TestOldLinks(FilesTestCase):
	USER = "m365-old-links@example.com"

	def setUp(self):
		super().setUp()
		if not frappe.db.exists("User", self.USER):
			frappe.get_doc(
				{"doctype": "User", "email": self.USER, "first_name": "Reader", "send_welcome_email": 0}
			).insert(ignore_permissions=True).add_roles("Desk User")
		self.mine = frappe.get_doc({"doctype": "ToDo", "description": "the reader's"}).insert()
		frappe.db.set_value("ToDo", self.mine.name, "owner", self.USER)

	def follow(self, path, user="Administrator"):
		"""Where redirect_moved_file sends GET path, as user; None when it lets Frappe answer."""
		from werkzeug.routing import RequestRedirect

		frappe.local.request = SimpleNamespace(method="GET", path=path)
		frappe.set_user(user)
		try:
			files.redirect_moved_file()
		except RequestRedirect as redirect:
			return redirect.new_url.split("file=", 1)[1]
		finally:
			frappe.set_user("Administrator")
			frappe.local.request = None
		return None

	def test_an_old_link_redirects_with_keep_local_on(self):
		configure(self, files_keep_local_copy=1)
		f = self.attach()
		old = f.file_url
		files.upload_file(f.name)
		self.addCleanup(remove, files.local_path(old))
		self.assertTrue(os.path.exists(files.local_path(old)))  # the kept copy
		self.assertEqual(self.follow(old), f.name)

	def test_a_shared_blob_redirects_to_the_row_the_user_may_read(self):
		content = f"shared {frappe.generate_hash()}".encode()
		for first, second in ((self.todo, self.mine), (self.mine, self.todo)):
			a = self.attach(content=content, attached_to_name=first.name)
			b = self.attach(content=content, attached_to_name=second.name)
			self.assertEqual(a.file_url, b.file_url)
			old = a.file_url
			files.upload_file(a.name)
			files.upload_file(b.name)
			readers = a if first is self.mine else b
			self.assertEqual(self.follow(old, self.USER), readers.name)
			frappe.db.delete("File", {"name": ["in", [a.name, b.name]]})
			content = f"shared {frappe.generate_hash()}".encode()

	def test_a_file_still_served_from_disk_is_left_to_frappe(self):
		content = f"shared {frappe.generate_hash()}".encode()
		a = self.attach(content=content, attached_to_name=self.mine.name)
		b = self.attach(content=content, attached_to_doctype="Note", attached_to_name="unmapped")
		frappe.db.set_value("File", b.name, "owner", self.USER)
		files.upload_file(a.name)
		self.assertTrue(os.path.exists(files.local_path(b.file_url)))  # b still uses the bytes
		self.assertIsNone(self.follow(b.file_url, self.USER))


# --- B21: record folders are never renamed ---------------------------------------------------


class TestRenamingARecord(FilesTestCase):
	def renamed(self):
		files.ensure_folder(MAPPED, self.todo.name)
		new = self.todo.name + "-renamed"
		frappe.rename_doc(MAPPED, self.todo.name, new, force=True)
		return new

	def test_the_folder_stays_linked_and_keeps_its_name(self):
		folder = files.ensure_folder(MAPPED, self.todo.name)
		new = self.renamed()
		self.assertIsNone(files.get_folder(MAPPED, self.todo.name))
		row = files.get_folder(MAPPED, new)
		self.assertEqual((row.item_id, row.folder_name), (folder.item_id, folder.folder_name))
		self.assertNotIn("PATCH", [m for m, _p in self.graph.calls])

	def test_a_row_frappe_did_not_rewrite_follows_without_renaming_the_folder(self):
		folder = files.ensure_folder(MAPPED, self.todo.name)
		old, new = self.todo.name, self.todo.name + "-renamed"
		frappe.db.set_value(MAPPED, old, "name", new)
		files.on_doc_rename(frappe.get_doc(MAPPED, new), old=old, new=new)
		self.assertEqual(files.get_folder(MAPPED, new).item_id, folder.item_id)
		self.assertNotIn("PATCH", [m for m, _p in self.graph.calls])

	def test_merging_into_a_record_with_its_own_folder_keeps_that_folder(self):
		other = frappe.get_doc({"doctype": "ToDo", "description": "merge target"}).insert()
		files.ensure_folder(MAPPED, self.todo.name)
		target = files.ensure_folder(MAPPED, other.name)
		files.on_doc_rename(other, old=self.todo.name, new=other.name, merge=True)
		self.assertIsNone(files.get_folder(MAPPED, self.todo.name))
		self.assertEqual(files.get_folder(MAPPED, other.name).item_id, target.item_id)


# --- B23: a recreated library or site is looked up again --------------------------------------


class TestRecreatedLibrary(FilesTestCase):
	def setUp(self):
		super().setUp()
		frappe.db.set_value("SharePoint Mapping", MAPPED, {"site_id": "OLD-SITE", "drive_id": "OLD-DRIVE"})
		files.clear_mapping_cache()

	def answer(self, method, url, kwargs):
		"""Graph as it answers once the library was deleted and created again (network-level)."""
		path = url[len(GRAPH) :]
		if path.startswith("/drives/OLD-DRIVE"):
			return HttpResponse(404, {"error": {"code": "itemNotFound", "message": "drive not found"}})
		if (method, path) == ("GET", "/sites/contoso.sharepoint.com:/sites/Sales"):
			return HttpResponse(200, {"id": "NEW-SITE"})
		if (method, path) == ("GET", "/sites/NEW-SITE/drives?$select=id,name,webUrl"):
			return HttpResponse(200, {"value": [{"id": "NEW-DRIVE", "name": "Documents"}]})
		if (method, path) == ("GET", "/drives/NEW-DRIVE/root:/Projects:"):
			return HttpResponse(200, {"id": "NEW-BASE", "folder": {}})
		if (method, path) == ("POST", "/drives/NEW-DRIVE/items/NEW-BASE/children"):
			body = kwargs["json"]
			return HttpResponse(
				201, {"id": "NEW-FOLDER", "name": body["name"], "folder": {}, "webUrl": "https://x/f"}
			)
		raise AssertionError(f"unexpected {method} {url}")

	def test_a_stale_drive_id_is_looked_up_again_once(self):
		real_graph(self)
		fake = network(self, self.answer)
		row = files.mapping_for(MAPPED)
		values = files.new_folder_values(row, MAPPED, self.todo.name)
		self.assertEqual((values["drive_id"], values["item_id"]), ("NEW-DRIVE", "NEW-FOLDER"))
		self.assertEqual(
			[(m, u[len(GRAPH) :]) for m, u, _k in fake.calls],
			[
				("GET", "/drives/OLD-DRIVE/root:/Projects:"),  # the whole path, then level by level
				("GET", "/drives/OLD-DRIVE/root:/Projects:"),
				("POST", "/drives/OLD-DRIVE/root/children"),
				("GET", "/drives/OLD-DRIVE"),  # is the library itself gone? yes
				("GET", "/sites/contoso.sharepoint.com:/sites/Sales"),
				("GET", "/sites/NEW-SITE/drives?$select=id,name,webUrl"),
				("GET", "/drives/NEW-DRIVE/root:/Projects:"),
				("POST", "/drives/NEW-DRIVE/items/NEW-BASE/children"),
			],
		)
		self.assertEqual(fake.calls[-1][2]["json"]["@microsoft.graph.conflictBehavior"], "rename")
		self.assertEqual(
			frappe.db.get_value("SharePoint Mapping", MAPPED, ["site_id", "drive_id"]),
			("NEW-SITE", "NEW-DRIVE"),
		)

	def test_a_404_inside_a_library_that_still_exists_is_not_retried(self):
		def answer(method, url, kwargs):
			path = url[len(GRAPH) :]
			if (method, path) == ("GET", "/drives/OLD-DRIVE"):
				return HttpResponse(200, {"id": "OLD-DRIVE"})
			return HttpResponse(404, {"error": {"code": "itemNotFound"}})

		real_graph(self)
		fake = network(self, answer)
		with self.assertRaises(graph.MsGraphNotFound):
			files.new_folder_values(files.mapping_for(MAPPED), MAPPED, self.todo.name)
		self.assertNotIn("/sites/", " ".join(u for _m, u, _k in fake.calls))


# --- B24: SharePoint's 400-character path limit, and deep folders -----------------------------


class TestPathLimits(FilesTestCase):
	def long_base(self, length):
		"""A base folder of segments of 100, making the library and base path ``length`` long."""
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		prefix = len("sites/Sales/Shared Documents/")
		base, left = [], length - prefix
		while left > 0:
			take = min(100, left)
			base.append("B" * take)
			left -= take + 1
		mapping.base_folder = "/".join(base)
		return mapping

	def test_the_library_path_is_counted_from_the_site(self):
		row = frappe._dict(site_url="https://contoso.sharepoint.com/sites/Sales", library="Documents")
		self.assertEqual(files.library_path(row), "sites/Sales/Shared Documents")
		row.update(use_base_folder=1, base_folder="Projects/2026")
		self.assertEqual(files.base_path_length(row), len("sites/Sales/Shared Documents/Projects/2026"))

	def test_a_mapping_whose_base_path_leaves_no_room_is_refused(self):
		mapping = self.long_base(files.MAX_PATH_LENGTH - files.MIN_ROOM_FOR_RECORDS + 5)
		with self.assertRaises(frappe.ValidationError) as caught:
			mapping.save()
		self.assertIn("400", str(caught.exception))
		self.long_base(files.MAX_PATH_LENGTH - files.MIN_ROOM_FOR_RECORDS - 5).save()

	def test_a_folder_is_not_created_under_a_base_path_that_is_too_long(self):
		row = files.mapping_for(MAPPED)
		row.base_folder = "/".join(["B" * 120] * 4)
		with self.assertRaises(frappe.ValidationError):
			files.new_folder_values(row, MAPPED, self.todo.name)
		self.assertNotIn("POST", [m for m, _p in self.graph.calls])

	def test_a_record_folder_name_leaves_room_for_its_files(self):
		self.long_base(files.MAX_PATH_LENGTH - files.MIN_ROOM_FOR_RECORDS).save()
		frappe.db.set_value(MAPPED, self.todo.name, "description", "D" * 200)
		frappe.db.set_value("SharePoint Mapping", MAPPED, "folder_pattern", "{description}")
		files.clear_mapping_cache()
		row = files.mapping_for(MAPPED)
		with patch.object(files, "ensure_path", return_value={"id": "BASE"}):
			values = files.new_folder_values(row, MAPPED, self.todo.name)
		used = files.base_path_length(row) + 1 + len(values["folder_name"])
		self.assertEqual(used, files.MAX_PATH_LENGTH - files.MIN_ROOM_FOR_FILE_NAME)

	def test_a_file_name_is_shortened_to_fit_keeping_its_extension(self):
		folder = frappe._dict(
			web_url="https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/" + "F" * 300
		)
		f = frappe._dict(file_name="quotation " * 20 + ".pdf", file_url="/private/files/q.pdf")
		name = files.upload_name(f, folder)
		self.assertTrue(name.endswith(".pdf"))
		self.assertEqual(
			len("sites/Sales/Shared Documents/" + "F" * 300) + 1 + len(name), files.MAX_PATH_LENGTH
		)

	def test_a_folder_with_no_room_left_gives_a_clear_error(self):
		folder = frappe._dict(
			web_url="https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/" + "F" * 395
		)
		with self.assertRaises(MsGraphError) as caught:
			files.upload_name(frappe._dict(file_name="spec.pdf"), folder)
		self.assertIn("SharePoint allows 400", str(caught.exception))

	def test_a_file_twelve_levels_down_in_the_records_folder_opens(self):
		folder = files.ensure_folder(MAPPED, self.todo.name)
		parent = folder.item_id
		for level in range(12):
			self.graph.items[f"SUB{level}"] = {"id": f"SUB{level}", "parent": parent}
			parent = f"SUB{level}"
		self.graph.items["DEEP"] = {"id": "DEEP", "parent": parent, "name": "deep.txt"}
		files._assert_inside(folder, "DEEP")
		self.graph.items["OUTSIDE"] = {"id": "OUTSIDE", "parent": "BASE"}
		with self.assertRaises(frappe.PermissionError):
			files._assert_inside(folder, "OUTSIDE")

	def test_mappings_close_to_the_limit_are_reported(self):
		self.assertEqual(files.upgrade_check_long_paths(), [])
		frappe.db.set_value("SharePoint Mapping", MAPPED, "base_folder", "/".join(["B" * 100] * 3))
		found = files.upgrade_check_long_paths()
		self.assertEqual([f["check"] for f in found], ["upgrade.sharepoint_long_paths"])
		self.assertEqual(found[0]["target"], MAPPED)
		self.assertIn("frappe_microsoft365.microsoft_files.upgrade_check_long_paths", upgrade_checks())


# --- B25: a retry adopts the upload a dead worker left behind ----------------------------------


class TestRetryAfterALostCommit(FilesTestCase):
	def setUp(self):
		super().setUp()
		files.add_folder_row(
			MAPPED,
			self.todo.name,
			{
				"drive_id": "DRIVE",
				"item_id": "FOLDER",
				"folder_name": "F",
				"folder_path": "Projects/F",
				"web_url": "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Projects/F",
			},
		)
		self.file = self.attach(name="report.csv")
		self.size = os.path.getsize(self.file.get_full_path())
		# The first attempt counted itself, uploaded, and died before recording the ids.
		frappe.db.set_value("File", self.file.name, "custom_microsoft_attempts", 1)
		real_graph(self)

	def left_behind(self, created_by):
		def answer(method, url, kwargs):
			if method == "GET" and url == f"{GRAPH}/drives/DRIVE/items/FOLDER:/report.csv:":
				return HttpResponse(
					200,
					{
						"id": "LEFT",
						"name": "report.csv",
						"size": self.size,
						"file": {},
						"webUrl": "https://x/report.csv",
						"parentReference": {"driveId": "DRIVE", "id": "FOLDER"},
						"createdBy": created_by,
					},
				)
			if method == "PUT" and url == f"{GRAPH}/drives/DRIVE/items/FOLDER:/report.csv:/content":
				return HttpResponse(
					201,
					{
						"id": "NEW",
						"name": "report (1).csv",
						"parentReference": {"driveId": "DRIVE", "id": "FOLDER"},
					},
				)
			raise AssertionError(f"unexpected {method} {url}")

		return network(self, answer)

	def test_the_retry_adopts_the_item_instead_of_uploading_again(self):
		client_id = frappe.get_single("Microsoft Settings").client_id
		fake = self.left_behind({"application": {"id": client_id, "displayName": "ERP"}})
		files.upload_file(self.file.name)
		self.assertEqual(
			[(m, u) for m, u, _k in fake.calls], [("GET", f"{GRAPH}/drives/DRIVE/items/FOLDER:/report.csv:")]
		)
		self.assertIn("size", fake.calls[0][2]["params"]["$select"])
		row = frappe.db.get_value(
			"File", self.file.name, ["custom_microsoft_status", "custom_microsoft_item_id"], as_dict=True
		)
		self.assertEqual((row.custom_microsoft_status, row.custom_microsoft_item_id), (files.STORED, "LEFT"))

	def test_a_file_a_person_put_there_is_not_taken(self):
		fake = self.left_behind({"user": {"id": "a-person", "displayName": "Asha"}})
		files.upload_file(self.file.name)
		self.assertEqual([m for m, _u, _k in fake.calls], ["GET", "PUT"])
		self.assertEqual(frappe.db.get_value("File", self.file.name, "custom_microsoft_item_id"), "NEW")

	def test_a_first_attempt_never_looks(self):
		frappe.db.set_value("File", self.file.name, "custom_microsoft_attempts", 0)
		fake = self.left_behind({})
		files.upload_file(self.file.name)
		self.assertEqual([m for m, _u, _k in fake.calls], ["PUT"])


# --- B26: a re-attached File's item moves with it ---------------------------------------------


class TestReattachedFileMoves(FilesTestCase):
	def setUp(self):
		super().setUp()
		# A name of its own: rows written by one test of a class are still there in the next.
		self.name = f"moved-{frappe.generate_hash(length=6)}.csv"
		self.file = self.attach(name=self.name)
		files.upload_file(self.file.name)  # into self.todo's folder, FOLDER (FakeGraph)
		self.item = frappe.db.get_value("File", self.file.name, "custom_microsoft_item_id")
		self.second = frappe.get_doc({"doctype": "ToDo", "description": "second"}).insert()
		files.add_folder_row(
			MAPPED,
			self.second.name,
			{"drive_id": "DRIVE", "item_id": "FOLDER2", "folder_name": "S", "folder_path": "Projects/S"},
		)
		self.jobs = patch.object(frappe, "enqueue").start()
		self.addCleanup(patch.stopall)
		real_graph(self)

	def reattach(self, doctype=MAPPED, name=None):
		doc = frappe.get_doc("File", self.file.name)
		doc.attached_to_doctype = doctype
		doc.attached_to_name = name or self.second.name
		doc.save()

	def answer(self, method, url, kwargs):
		path = url[len(GRAPH) :]
		if (method, path) == ("PATCH", f"/drives/DRIVE/items/{self.item}"):
			return HttpResponse(
				200,
				{
					"id": self.item,
					"webUrl": "https://x/S/m.csv",
					"parentReference": {"driveId": "DRIVE", "id": "FOLDER2"},
				},
			)
		if (method, path) == ("GET", f"/drives/DRIVE/items/{self.item}/content"):
			return HttpResponse(200, body=b"the bytes")
		if (method, path) == ("PUT", f"/drives/DRIVE/items/FOLDER2:/{self.name}:/content"):
			return HttpResponse(
				201, {"id": "OWN-COPY", "parentReference": {"driveId": "DRIVE", "id": "FOLDER2"}}
			)
		if (method, path) == ("DELETE", f"/drives/DRIVE/items/{self.item}"):
			return HttpResponse(204)
		raise AssertionError(f"unexpected {method} {url}")

	def location(self):
		return frappe.db.get_value(
			"File", self.file.name, ["custom_microsoft_item_id", "custom_microsoft_folder_id"], as_dict=True
		)

	def test_re_attaching_moves_the_item_to_the_new_records_folder(self):
		fake = network(self, self.answer)
		self.reattach()
		self.assertEqual(
			[(m, u) for m, u, _k in fake.calls], [("PATCH", f"{GRAPH}/drives/DRIVE/items/{self.item}")]
		)
		self.assertEqual(fake.calls[0][2]["json"], {"parentReference": {"id": "FOLDER2"}})
		self.assertEqual(fake.calls[0][2]["params"], {"@microsoft.graph.conflictBehavior": "rename"})
		self.assertEqual(
			self.location(), {"custom_microsoft_item_id": self.item, "custom_microsoft_folder_id": "FOLDER2"}
		)
		# Readers of the first record no longer find it through that record; readers of the second do.
		files.assert_app_item(frappe.get_doc("File", self.file.name))

	def test_an_item_shared_with_attachment_copies_is_copied_for_the_file_that_moved(self):
		frappe.get_doc("File", self.file.name).create_attachment_copy("Note", "anything")
		fake = network(self, self.answer)
		self.reattach()
		self.assertEqual(fake.calls, [])  # bytes are moved by a job, not in the save
		job = self.jobs.call_args.kwargs
		self.assertEqual((job["file_name"], job["old_name"]), (self.file.name, self.todo.name))
		files.follow_attachment(self.file.name, MAPPED, self.todo.name, in_job=True)
		self.assertEqual(
			[m for m, _u, _k in fake.calls], ["GET", "PUT"]
		)  # its own copy; the original stays for the copies
		self.assertEqual(
			self.location(), {"custom_microsoft_item_id": "OWN-COPY", "custom_microsoft_folder_id": "FOLDER2"}
		)

	def test_re_attached_to_an_unmapped_record_it_comes_back_to_this_server(self):
		fake = network(self, self.answer)
		self.reattach("Note", "unmapped")
		files.follow_attachment(self.file.name, MAPPED, self.todo.name, in_job=True)
		row = frappe.db.get_value(
			"File", self.file.name, ["file_url", "custom_microsoft_status"], as_dict=True
		)
		self.assertFalse(row.custom_microsoft_status)
		self.assertTrue(row.file_url.startswith("/private/files/"))
		self.addCleanup(remove, files.local_path(row.file_url))
		self.assertEqual([m for m, _u, _k in fake.calls], ["GET", "DELETE"])

	def test_an_item_that_was_not_in_the_old_records_folder_is_not_moved(self):
		third = frappe.get_doc({"doctype": "ToDo", "description": "third"}).insert()
		fake = network(self, self.answer)
		files.follow_attachment(self.file.name, MAPPED, third.name, in_job=True)
		self.assertEqual(fake.calls, [])

	def test_saving_without_re_attaching_does_nothing(self):
		fake = network(self, self.answer)
		doc = frappe.get_doc("File", self.file.name)
		doc.save()
		self.assertEqual(fake.calls, [])


# --- B30: get_content never reads a huge file into memory -------------------------------------


class Unread(HttpResponse):
	"""A response whose body must not be touched as a whole."""

	@property
	def content(self):
		raise AssertionError("the whole body was read")

	@content.setter
	def content(self, value):
		self._body = value

	def iter_content(self, chunk_size=1):
		for i in range(0, len(self._body), chunk_size):
			yield self._body[i : i + chunk_size]


class TestBigFilesAreNotReadWhole(FilesTestCase):
	def setUp(self):
		super().setUp()
		self.file = self.attach(name="clip.txt", content=f"small {frappe.generate_hash()}".encode())
		files.upload_file(self.file.name)
		real_graph(self)

	def serve(self, body, length=None):
		headers = {"Content-Length": str(length if length is not None else len(body))}
		return network(self, lambda method, url, kwargs: Unread(200, body=body, headers=headers))

	def test_a_file_over_the_limit_is_refused_before_it_is_read(self):
		fake = self.serve(b"x" * 10, length=6 * 1024**3)
		with self.assertRaises(frappe.ValidationError) as caught:
			frappe.get_doc("File", self.file.name).get_content()
		self.assertIn("MB", str(caught.exception))
		method, url, kwargs = fake.calls[0]
		self.assertEqual((method, url), ("GET", f"{GRAPH}/drives/DRIVE/items/ITEM-clip.txt/content"))
		self.assertTrue(kwargs["stream"])

	def test_a_body_that_runs_past_the_limit_is_refused_too(self):
		self.serve(b"x" * 64, length=0)
		with patch.object(files, "max_content_bytes", return_value=32):
			with self.assertRaises(frappe.ValidationError):
				files.fetch_content("DRIVE", "ITEM-clip.txt")

	def test_a_file_under_the_limit_is_returned(self):
		self.serve(b"hello from teams")
		self.assertEqual(frappe.get_doc("File", self.file.name).get_content(), "hello from teams")

	def test_the_limit_is_frappes_upload_limit_unless_the_site_sets_one(self):
		from frappe.core.api.file import get_max_file_size

		self.assertEqual(files.max_content_bytes(), get_max_file_size())
		with patch.dict(frappe.conf, {"microsoft365_max_content_bytes": 1234}):
			self.assertEqual(files.max_content_bytes(), 1234)

	def test_streaming_callers_get_every_piece(self):
		self.serve(b"y" * (files.DOWNLOAD_CHUNK_BYTES + 3))
		self.assertEqual(
			sum(len(p) for p in files.iter_content("DRIVE", "ITEM-clip.txt")), files.DOWNLOAD_CHUNK_BYTES + 3
		)


# --- F5: large uploads resume instead of restarting --------------------------------------------


class TestResumableUploads(BaseTestCase):
	"""Network-level: graph_request creates the session for real; only the transport is faked."""

	UPLOAD_URL = "https://contoso.sharepoint.com/_api/v2.0/drives/DRIVE/items/upload?tempauth=SECRET"
	CHUNK = files.UPLOAD_CHUNK_BYTES
	SIZE = 3 * files.UPLOAD_CHUNK_BYTES + 5  # four ranges

	def setUp(self):
		super().setUp()
		self.key = files._session_key("DRIVE", "FOLDER", "big.bin", self.SIZE)
		files._forget_session(self.key)
		self.addCleanup(files._forget_session, self.key)
		self.sleeps = patch.object(files.time, "sleep").start()
		self.addCleanup(patch.stopall)
		self.sessions = 0
		self.received = 0  # what SharePoint has, for the status query
		self.script = {}  # (start of range, how many times seen) -> response to give

	def answer(self, method, url, kwargs):
		if method == "POST" and url == f"{GRAPH}/drives/DRIVE/items/FOLDER:/big.bin:/createUploadSession":
			self.sessions += 1
			return HttpResponse(
				200, {"uploadUrl": self.UPLOAD_URL, "expirationDateTime": "2099-01-01T00:00:00Z"}
			)
		self.assertEqual(url, self.UPLOAD_URL)
		self.assertNotIn("Authorization", kwargs.get("headers") or {})
		if method == "GET":
			return HttpResponse(200, {"nextExpectedRanges": [f"{self.received}-"]})
		if method == "DELETE":
			return HttpResponse(204)
		start, end = map(int, kwargs["headers"]["Content-Range"].split(" ", 1)[1].split("/")[0].split("-"))
		seen = self.seen.get(start, 0) + 1
		self.seen[start] = seen
		scripted = self.script.get((start, seen))
		if scripted:
			return scripted() if callable(scripted) else scripted
		self.received = end + 1
		if end == self.SIZE - 1:
			return HttpResponse(201, {"id": "BIG", "name": "big.bin"})
		return HttpResponse(202, {"nextExpectedRanges": [f"{end + 1}-"]})

	def upload(self):
		return files.upload_stream("DRIVE", "FOLDER", "big.bin", io.BytesIO(bytes(self.SIZE)), self.SIZE)

	def ranges(self, fake):
		return [(m, (k.get("headers") or {}).get("Content-Range")) for m, _u, k in fake.calls[1:]]

	def go(self):
		self.seen = {}
		return network(self, self.answer)

	def test_a_throttled_range_waits_as_asked_and_is_sent_again(self):
		fake = self.go()
		self.script[(self.CHUNK, 1)] = HttpResponse(
			503, {"error": {"code": "serviceNotAvailable"}}, {"Retry-After": "5"}
		)
		self.assertEqual(self.upload(), {"id": "BIG", "name": "big.bin"})
		self.sleeps.assert_called_once_with(5)
		self.assertEqual(self.sessions, 1)
		self.assertEqual(
			[r for _m, r in self.ranges(fake)],
			[
				f"bytes 0-{self.CHUNK - 1}/{self.SIZE}",
				f"bytes {self.CHUNK}-{2 * self.CHUNK - 1}/{self.SIZE}",
				f"bytes {self.CHUNK}-{2 * self.CHUNK - 1}/{self.SIZE}",
				f"bytes {2 * self.CHUNK}-{3 * self.CHUNK - 1}/{self.SIZE}",
				f"bytes {3 * self.CHUNK}-{self.SIZE - 1}/{self.SIZE}",
			],
		)
		self.assertNotIn("DELETE", [m for m, _r in self.ranges(fake)])
		self.assertIsNone(frappe.cache().get_value(self.key, expires=True))

	def test_a_long_wait_keeps_the_session_and_the_next_attempt_resumes_it(self):
		fake = self.go()
		self.script[(2 * self.CHUNK, 1)] = HttpResponse(
			429, {"error": {"code": "tooManyRequests"}}, {"Retry-After": "600"}
		)
		with self.assertRaises(files.UploadInterrupted):
			self.upload()
		self.sleeps.assert_not_called()
		self.assertNotIn("DELETE", [m for m, _r in self.ranges(fake)])
		self.assertTrue(frappe.cache().get_value(self.key, expires=True))  # kept, encrypted
		self.assertNotIn("SECRET", str(frappe.cache().get_value(self.key, expires=True)))

		calls = len(fake.calls)
		self.assertEqual(self.upload()["id"], "BIG")
		self.assertEqual(self.sessions, 1)  # the same session, not a new one
		self.assertEqual(
			[(m, (k.get("headers") or {}).get("Content-Range")) for m, _u, k in fake.calls[calls:]],
			[
				("GET", None),  # where did we stop?
				("PUT", f"bytes {2 * self.CHUNK}-{3 * self.CHUNK - 1}/{self.SIZE}"),
				("PUT", f"bytes {3 * self.CHUNK}-{self.SIZE - 1}/{self.SIZE}"),
			],
		)

	def test_a_dropped_connection_asks_the_session_where_to_carry_on(self):
		import requests

		def dropped_after_arriving():
			self.received = 2 * self.CHUNK  # the range got there; the answer did not
			raise requests.ConnectionError("connection reset")

		fake = self.go()
		self.script[(self.CHUNK, 1)] = dropped_after_arriving
		self.assertEqual(self.upload()["id"], "BIG")
		self.assertEqual(
			[r for _m, r in self.ranges(fake)],
			[
				f"bytes 0-{self.CHUNK - 1}/{self.SIZE}",
				f"bytes {self.CHUNK}-{2 * self.CHUNK - 1}/{self.SIZE}",
				None,  # GET the session's status
				f"bytes {2 * self.CHUNK}-{3 * self.CHUNK - 1}/{self.SIZE}",
				f"bytes {3 * self.CHUNK}-{self.SIZE - 1}/{self.SIZE}",
			],
		)

	def test_a_refusal_that_will_not_pass_drops_the_session(self):
		fake = self.go()
		self.script[(0, 1)] = HttpResponse(400, {"error": {"code": "invalidRange"}})
		with self.assertRaises(MsGraphError):
			self.upload()
		self.assertEqual(fake.calls[-1][0], "DELETE")
		self.assertIsNone(frappe.cache().get_value(self.key, expires=True))
