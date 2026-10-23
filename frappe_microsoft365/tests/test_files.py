"""SharePoint document storage: naming, eligibility, uploads, opening, permissions, doctor checks.

Nothing here touches the network. Graph calls go through ``microsoft_graph.graph_request`` (and
upload ranges through ``microsoft_graph._http_request``), both replaced with fakes that answer
like Graph does.
"""

import os
from itertools import pairwise
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import frappe
from frappe.utils import add_to_date

from frappe_microsoft365 import doctor
from frappe_microsoft365 import microsoft_files as files
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import APP, MsGraphError, MsGraphNotFound
from frappe_microsoft365.tests.base import BaseTestCase

MAPPED = "ToDo"


def setUpModule():
	"""Create the File fields before any test switches storage on.

	Switching it on creates them (capabilities.py), and creating a column is DDL, which MariaDB
	commits together with whatever the test had written so far — settings a test means to roll
	back would stick. Done once here, the switch finds them present and alters nothing.
	"""
	from frappe_microsoft365.setup import create_file_custom_fields

	create_file_custom_fields()
	frappe.db.commit()  # nosemgrep: DDL above has committed already; make that explicit


STORAGE_OPTIONS = (
	"files_keep_local_copy",
	"files_archive_links",
	"files_delete_remote",
	"files_on_uninstall",
)


def configure(test, **overrides):
	"""Turn Document Storage on with ToDo mapped, restoring the previous settings afterwards.

	Overrides for SharePoint Settings options (files_*) go there; the rest to Microsoft Settings.
	Storage cannot be switched on before Frappe v16, so a test that needs it is skipped there.
	"""
	from frappe_microsoft365 import capabilities

	if capabilities.unmet(capabilities.get("files"), {}):
		test.skipTest("SharePoint document storage needs Frappe v16")
	doc = frappe.get_doc("Microsoft Settings")
	before = {k: doc.get(k) for k in ("enabled", "tenant_id", "client_id", "use_files")}
	storage = frappe.get_doc("SharePoint Settings")
	storage_before = {k: storage.get(k) for k in STORAGE_OPTIONS}

	doc.update(
		{
			"enabled": 1,
			"tenant_id": "9f0b1c2d-0000-4000-8000-000000000001",
			"client_id": "9f0b1c2d-0000-4000-8000-000000000002",
			"client_secret": "a-secret~value.not-a-guid",
			"use_files": 1,
		}
	)
	doc.update({k: v for k, v in overrides.items() if k not in STORAGE_OPTIONS})
	doc.flags.ignore_permissions = True
	doc.save()
	frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")

	storage.update({k: None if k == "files_on_uninstall" else 0 for k in STORAGE_OPTIONS})
	storage.update({k: v for k, v in overrides.items() if k in STORAGE_OPTIONS})
	storage.flags.ignore_permissions = True
	storage.save()
	frappe.clear_document_cache("SharePoint Settings", "SharePoint Settings")

	frappe.db.delete("SharePoint Mapping", {"name": MAPPED})
	frappe.db.delete("SharePoint Mapping Role", {"parent": MAPPED})
	frappe.db.delete("SharePoint Folder", {"parent": MAPPED})
	mapping = frappe.get_doc(
		{
			"doctype": "SharePoint Mapping",
			"enabled": 1,
			"reference_doctype": MAPPED,
			"site_url": "https://contoso.sharepoint.com/sites/Sales",
			"library": "Documents",
			"base_folder": "Projects",
			"folder_pattern": "{name}",
			"site_id": "SITE",
			"drive_id": "DRIVE",
		}
	).insert(ignore_permissions=True)

	def restore():
		frappe.db.delete("SharePoint Mapping", {"name": MAPPED})
		frappe.db.delete("SharePoint Mapping Role", {"parent": MAPPED})
		frappe.db.delete("SharePoint Folder", {"parent": MAPPED})
		files.clear_mapping_cache()
		d = frappe.get_doc("Microsoft Settings")
		d.update(before)
		d.flags.ignore_validate = True
		d.flags.ignore_permissions = True
		d.save()
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")
		s = frappe.get_doc("SharePoint Settings")
		s.update(storage_before)
		s.flags.ignore_permissions = True
		s.save()
		frappe.clear_document_cache("SharePoint Settings", "SharePoint Settings")

	test.addCleanup(restore)
	return mapping


class FakeGraph:
	"""Answers the Graph calls microsoft_files makes, and records them."""

	def __init__(self):
		self.calls = []
		self.items = {"FOLDER": {"id": "FOLDER", "parent": "BASE"}, "BASE": {"id": "BASE", "parent": "ROOT"}}
		self.fail_upload = None

	def __call__(
		self,
		method,
		path,
		caller,
		json=None,
		params=None,
		headers=None,
		raw=False,
		stream=False,
		data=None,
		timeout=30,
	):
		assert caller is APP, "document storage must always call as the application"
		self.calls.append((method, path))
		if method == "GET" and path == "/drives/DRIVE":
			return {"webUrl": "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents"}
		if method == "GET" and path == "/drives/DRIVE/root:/Projects/" + path.rsplit("/", 1)[-1]:
			return {"id": "FOLDER", "folder": {}, "webUrl": "https://contoso.sharepoint.com/f"}
		if method == "GET" and ":/Projects/" in path:
			return {"id": "FOLDER", "folder": {}, "webUrl": "https://contoso.sharepoint.com/f"}
		if method == "PUT" and path.endswith(":/content"):
			if self.fail_upload:
				raise self.fail_upload
			name = path.split(":/", 1)[1].rsplit(":/content", 1)[0]
			return {
				"id": "ITEM-" + name,
				"webUrl": "https://contoso.sharepoint.com/" + name,
				"parentReference": {"driveId": "DRIVE", "id": "FOLDER"},
			}
		if method == "GET" and path.startswith("/drives/DRIVE/items/") and params:
			item_id = path.rsplit("/", 1)[-1]
			item = self.items.get(item_id)
			if not item:
				raise MsGraphNotFound("404")
			return {
				"id": item_id,
				"name": item.get("name", item_id),
				"file": {},
				"parentReference": {"driveId": "DRIVE", "id": item.get("parent")},
			}
		raise AssertionError(f"unexpected Graph call {method} {path}")


class FilesTestCase(BaseTestCase):
	def setUp(self):
		super().setUp()
		frappe.set_user("Administrator")
		configure(self)
		self.todo = frappe.get_doc({"doctype": "ToDo", "description": "SharePoint test"}).insert()
		self.graph = FakeGraph()
		patcher = patch.object(graph, "graph_request", side_effect=self.graph)
		patcher.start()
		self.addCleanup(patcher.stop)
		# Jobs commit; a test must not.
		for name in ("commit",):
			p = patch.object(frappe.db, name)
			p.start()
			self.addCleanup(p.stop)
		p = patch.object(files, "enqueue_upload")
		self.enqueued = p.start()
		self.addCleanup(p.stop)

	def attach(self, content=None, name="spec.csv", **extra):
		# Unique bytes unless a test asks otherwise: Frappe stores identical content once, so a
		# blob shared with an earlier test's File would rightly never be deleted.
		content = content or f"a,b\n{frappe.generate_hash()},2\n".encode()
		return frappe.get_doc(
			{
				"doctype": "File",
				"file_name": name,
				"content": content,
				"attached_to_doctype": MAPPED,
				"attached_to_name": self.todo.name,
				"is_private": 1,
				**extra,
			}
		).insert()


# --- pure helpers ----------------------------------------------------------------------------


class TestNames(BaseTestCase):
	def test_slashes_from_naming_series_become_hyphens(self):
		self.assertEqual(files.safe_name("PO/2026/0042-A"), "PO-2026-0042-A")

	def test_characters_sharepoint_refuses_are_replaced(self):
		self.assertEqual(files.safe_name('a"b*c:d<e>f?g\\h|i#j%k'), "a-b-c-d-e-f-g-h-i-j-k")

	def test_reserved_names_and_leading_lock_prefix(self):
		self.assertEqual(files.safe_name("CON"), "_CON")
		self.assertEqual(files.safe_name("~$draft"), "draft")
		self.assertEqual(files.safe_name("  .hidden. "), "hidden")

	def test_empty_falls_back(self):
		self.assertEqual(files.safe_name("///"), "Untitled")

	def test_truncation_keeps_the_extension(self):
		name = files.safe_file_name("x" * 400 + ".pdf")
		self.assertTrue(name.endswith(".pdf"))
		self.assertLessEqual(len(name), 200)

	def test_pattern_uses_fields_and_ignores_unknown_tokens(self):
		doc = SimpleNamespace(name="PRJ/01", get=lambda k, d=None: {"customer": "ACME"}.get(k, d))
		self.assertEqual(files.render_pattern("{name} - {customer}{nope}", doc), "PRJ-01 - ACME")
		self.assertEqual(files.render_pattern("", doc), "PRJ-01")

	def test_base_segments_follow_the_tickbox(self):
		self.assertEqual(files.base_segments(frappe._dict(use_base_folder=1, base_folder="A/B")), ["A", "B"])
		self.assertEqual(files.base_segments(frappe._dict(use_base_folder=0, base_folder="A/B")), [])
		# Rows from before the tickbox existed behave as they always did.
		self.assertEqual(files.base_segments(frappe._dict(base_folder="A")), ["A"])

	def test_path_segments(self):
		self.assertEqual(files.path_segments("Projects / 2026"), ["Projects", "2026"])
		self.assertEqual(files.path_segments(""), [])

	def test_site_url_accepts_whatever_was_in_the_address_bar(self):
		self.assertEqual(
			files.parse_site_url(
				"https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Forms/AllItems.aspx"
			),
			("contoso.sharepoint.com", "/sites/Sales"),
		)
		self.assertEqual(
			files.parse_site_url("contoso.sharepoint.com/teams/Ops"), ("contoso.sharepoint.com", "/teams/Ops")
		)
		self.assertEqual(
			files.parse_site_url("https://contoso.sharepoint.com/"), ("contoso.sharepoint.com", "")
		)

	def test_upload_ranges_are_multiples_of_320_kib(self):
		ranges = files.chunk_ranges(25 * 1024 * 1024 + 7)
		for start, end in ranges[:-1]:
			self.assertEqual((end - start + 1) % (320 * 1024), 0)
		self.assertEqual(ranges[0][0], 0)
		self.assertEqual(ranges[-1][1], 25 * 1024 * 1024 + 6)
		for (_s1, e1), (s2, _e2) in pairwise(ranges):
			self.assertEqual(e1 + 1, s2)

	def test_private_addresses_are_never_fetched(self):
		for url in (
			"http://127.0.0.1/x",
			"http://169.254.169.254/latest/meta-data",
			"ftp://example.com/a",
			"file:///etc/passwd",
		):
			with self.assertRaises(MsGraphError):
				files.assert_public_url(url)


# --- app-only transport ----------------------------------------------------------------------


def _response(status=200, payload=None):
	resp = MagicMock()
	resp.status_code = status
	resp.headers = {}
	resp.content = b"{}"
	resp.reason = "Mocked"
	resp.json.return_value = payload if payload is not None else {}
	return resp


class TestAppOnly(BaseTestCase):
	def test_app_only_calls_use_the_client_credentials_token(self):
		with (
			patch.object(graph, "get_app_access_token", return_value="APPTOKEN") as tok,
			patch.object(graph, "get_valid_access_token") as delegated,
			patch.object(graph, "_http_request", return_value=_response(payload={"id": "x"})) as http,
		):
			self.assertEqual(graph.graph_request("GET", "/sites/root", APP), {"id": "x"})
		tok.assert_called_once()
		delegated.assert_not_called()
		self.assertEqual(http.call_args.kwargs["headers"]["Authorization"], "Bearer APPTOKEN")

	def test_401_mints_a_new_app_token_once(self):
		with (
			patch.object(graph, "get_app_access_token", side_effect=["OLD", "NEW"]),
			patch.object(graph, "clear_app_token_cache") as cleared,
			patch.object(
				graph, "_http_request", side_effect=[_response(401), _response(payload={"ok": 1})]
			) as http,
		):
			self.assertEqual(graph.graph_request("GET", "/sites/root", APP), {"ok": 1})
		cleared.assert_called_once()
		self.assertEqual(http.call_args.kwargs["headers"]["Authorization"], "Bearer NEW")

	def test_404_and_409_raise_their_own_errors(self):
		with (
			patch.object(graph, "get_app_access_token", return_value="T"),
			patch.object(
				graph, "_http_request", return_value=_response(404, {"error": {"code": "itemNotFound"}})
			),
		):
			with self.assertRaises(MsGraphNotFound):
				graph.graph_request("GET", "/drives/d/items/x", APP)
		with (
			patch.object(graph, "get_app_access_token", return_value="T"),
			patch.object(
				graph, "_http_request", return_value=_response(409, {"error": {"code": "nameAlreadyExists"}})
			),
		):
			with self.assertRaises(graph.MsGraphConflict):
				graph.graph_request("POST", "/drives/d/items/x/children", APP)

	def _app_token(self, settings, minted=("SECRET-TOKEN",), secret="main-secret"):
		"""get_app_access_token against these settings, with MSAL faked. Returns (tokens, msal mock)."""
		import msal

		app = MagicMock()
		app.acquire_token_for_client.side_effect = [
			{"access_token": token, "expires_in": 3600} for token in minted
		]
		with (
			patch.object(graph, "get_settings", return_value=settings),
			patch.object(graph, "_client_secret", return_value=secret),
			patch.object(graph, "get_decrypted_password", return_value="app-secret"),
			patch.object(msal, "ConfidentialClientApplication", return_value=app) as cca,
		):
			tokens = [graph.get_app_access_token() for _ in minted]
		return tokens, app, cca

	def test_app_token_is_cached_encrypted(self):
		graph.clear_app_token_cache()
		self.addCleanup(graph.clear_app_token_cache)
		settings = frappe._dict(enabled=1, tenant_id="contoso.onmicrosoft.com", client_id="MAIN")
		tokens, app, _cca = self._app_token(settings, minted=("SECRET-TOKEN", "UNUSED"))
		self.assertEqual(tokens, ["SECRET-TOKEN", "SECRET-TOKEN"])
		app.acquire_token_for_client.assert_called_once_with(scopes=[graph.APP_GRAPH_SCOPE])
		self.assertNotIn("SECRET-TOKEN", str(frappe.cache().get_value(graph.APP_TOKEN_CACHE_KEY, expires=True)))

	def test_separate_application_registration_is_used_when_set(self):
		graph.clear_app_token_cache()
		self.addCleanup(graph.clear_app_token_cache)
		settings = frappe._dict(
			enabled=1, tenant_id="contoso.onmicrosoft.com", client_id="MAIN", app_client_id="APPREG"
		)
		_tokens, _app, cca = self._app_token(settings)
		self.assertEqual(cca.call_args.args[0], "APPREG")
		self.assertEqual(cca.call_args.kwargs["client_credential"], "app-secret")

	def test_main_registration_is_used_when_no_separate_one(self):
		graph.clear_app_token_cache()
		self.addCleanup(graph.clear_app_token_cache)
		settings = frappe._dict(enabled=1, tenant_id="contoso.onmicrosoft.com", client_id="MAIN")
		_tokens, _app, cca = self._app_token(settings)
		self.assertEqual(cca.call_args.args[0], "MAIN")
		self.assertEqual(cca.call_args.kwargs["client_credential"], "main-secret")

	def test_a_cached_token_from_another_registration_is_not_served(self):
		graph.clear_app_token_cache()
		self.addCleanup(graph.clear_app_token_cache)
		tenant = "contoso.onmicrosoft.com"
		self._app_token(frappe._dict(enabled=1, tenant_id=tenant, client_id="OLD"), minted=("OLD-TOKEN",))
		tokens, _app, _cca = self._app_token(
			frappe._dict(enabled=1, tenant_id=tenant, client_id="NEW"), minted=("NEW-TOKEN",)
		)
		self.assertEqual(tokens, ["NEW-TOKEN"])

	def test_app_calls_need_a_specific_tenant(self):
		for tenant in ("", "common", "organizations"):
			with patch.object(
				graph, "get_settings", return_value=frappe._dict(enabled=1, tenant_id=tenant, client_id="X")
			):
				with self.assertRaises(MsGraphError):
					graph.app_credentials()

	def test_saving_microsoft_settings_forgets_the_app_token(self):
		frappe.cache().set_value(graph.APP_TOKEN_CACHE_KEY, {"owner": "x", "token": "y"})
		configure(self)
		self.assertIsNone(frappe.cache().get_value(graph.APP_TOKEN_CACHE_KEY, expires=True))

	def test_large_uploads_send_ranges_without_a_bearer_token(self):
		size = files.SIMPLE_UPLOAD_MAX_BYTES + 1
		fh = SimpleNamespace(read=lambda n=-1: b"x" * (n if n and n > 0 else size))
		session = {"uploadUrl": "https://contoso.sharepoint.com/upload?sig=1"}
		with (
			patch.object(graph, "graph_request", return_value=session) as gr,
			patch.object(graph, "_http_request", return_value=_response(201, {"id": "BIG"})) as http,
		):
			item = files.upload_stream("DRIVE", "FOLDER", "big.zip", fh, size, "application/zip")
		self.assertEqual(item, {"id": "BIG"})
		self.assertIn("createUploadSession", gr.call_args.args[1])
		headers = http.call_args.kwargs["headers"]
		self.assertNotIn("Authorization", headers)
		self.assertEqual(headers["Content-Range"], f"bytes 0-{size - 1}/{size}")


# --- eligibility and the upload job ---------------------------------------------------------


class TestUpload(FilesTestCase):
	def test_new_attachment_is_queued(self):
		f = self.attach()
		self.enqueued.assert_called_once_with(f.name)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.PENDING)

	def test_field_attachments_and_unmapped_doctypes_stay_local(self):
		self.assertFalse(
			files._eligible(
				frappe._dict(
					attached_to_doctype=MAPPED,
					attached_to_name="x",
					attached_to_field="image",
					file_url="/files/a.png",
				)
			)
		)
		self.assertFalse(
			files._eligible(
				frappe._dict(attached_to_doctype="Note", attached_to_name="x", file_url="/files/a.pdf")
			)
		)
		self.assertFalse(
			files._eligible(
				frappe._dict(
					attached_to_doctype=MAPPED, attached_to_name="x", file_url="https://example.com/a.pdf"
				)
			)
		)

	def test_links_are_eligible_only_when_archiving_is_on(self):
		configure(self, files_archive_links=1)
		self.assertTrue(
			files._eligible(
				frappe._dict(
					attached_to_doctype=MAPPED, attached_to_name="x", file_url="https://example.com/a.pdf"
				)
			)
		)

	def test_upload_moves_the_file_and_frees_the_disk(self):
		f = self.attach()
		path = f.get_full_path()
		self.assertTrue(os.path.exists(path))

		files.upload_file(f.name)

		row = frappe.db.get_value(
			"File",
			f.name,
			[
				"file_url",
				"custom_microsoft_status",
				"custom_microsoft_item_id",
				"custom_microsoft_drive_id",
				"custom_microsoft_local_url",
				"content_hash",
			],
			as_dict=True,
		)
		self.assertEqual(row.file_url, files.stored_url(f.name))
		self.assertEqual(row.custom_microsoft_status, files.STORED)
		self.assertEqual(row.custom_microsoft_item_id, "ITEM-" + f.file_name)
		self.assertEqual(row.custom_microsoft_drive_id, "DRIVE")
		self.assertEqual(row.custom_microsoft_local_url, f.file_url)  # old links redirect
		self.assertIsNone(row.content_hash)
		self.assertFalse(os.path.exists(path))

		folder = files.get_folder(MAPPED, self.todo.name)
		self.assertEqual(folder.item_id, "FOLDER")
		self.assertEqual(folder.folder_path, f"Projects/{files.safe_name(self.todo.name)}")

	def test_attachment_comment_link_follows_the_move(self):
		f = self.attach()
		old = f.file_url
		comment = frappe.get_all(
			"Comment",
			filters={"reference_name": self.todo.name, "comment_type": "Attachment"},
			fields=["name", "content"],
		)
		self.assertTrue(comment and old in comment[0].content)
		with patch.object(frappe, "publish_realtime") as published:
			files.upload_file(f.name)
		content = frappe.db.get_value("Comment", comment[0].name, "content")
		self.assertNotIn(old, content)
		self.assertIn(files.stored_url(f.name), content)
		self.assertEqual(published.call_args.args[0], files.MOVED_EVENT)
		self.assertEqual(published.call_args.args[1]["old_url"], old)

	def test_record_folders_can_go_straight_into_the_library(self):
		mapping = configure(self)
		mapping.use_base_folder = 0
		mapping.save()
		self.assertEqual(files.base_segments(files.mapping_for(MAPPED)), [])
		with patch.object(
			files, "ensure_path", return_value={"id": "TOP", "webUrl": "https://contoso.sharepoint.com/top"}
		) as ensure:
			folder = files.ensure_folder(MAPPED, self.todo.name)
		ensure.assert_called_once_with("DRIVE", [files.safe_name(self.todo.name)])
		self.assertEqual(folder.folder_path, files.safe_name(self.todo.name))

	def test_old_links_redirect_to_the_moved_file(self):
		from types import SimpleNamespace as NS

		from werkzeug.routing import RequestRedirect

		f = self.attach()
		old = f.file_url
		files.upload_file(f.name)
		frappe.local.request = NS(method="GET", path=old)
		self.addCleanup(setattr, frappe.local, "request", None)
		with self.assertRaises(RequestRedirect) as caught:
			files.redirect_moved_file()
		self.assertTrue(caught.exception.new_url.endswith(files.stored_url(f.name)))
		# anything else passes straight through
		frappe.local.request = NS(method="GET", path="/private/files/never-existed.txt")
		files.redirect_moved_file()
		frappe.local.request = NS(method="GET", path="/app/todo")
		files.redirect_moved_file()

	def test_keep_local_copy(self):
		configure(self, files_keep_local_copy=1)
		f = self.attach()
		local = f.file_url
		files.upload_file(f.name)
		self.assertTrue(os.path.exists(f.get_full_path()))
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_local_url"), local)

	def test_shared_blob_is_not_deleted_while_another_file_uses_it(self):
		f1 = self.attach(content=b"same bytes")
		f2 = self.attach(content=b"same bytes", attached_to_doctype="Note", attached_to_name="whatever")
		self.assertEqual(f1.file_url, f2.file_url)  # Frappe stores identical content once
		files.upload_file(f1.name)
		self.assertTrue(os.path.exists(f2.get_full_path()))

	def test_failure_leaves_the_file_where_it_was(self):
		self.graph.fail_upload = MsGraphError("SharePoint said no")
		f = self.attach()
		files.upload_file(f.name)
		row = frappe.db.get_value(
			"File",
			f.name,
			["file_url", "custom_microsoft_status", "custom_microsoft_error", "custom_microsoft_attempts"],
			as_dict=True,
		)
		self.assertEqual(row.file_url, f.file_url)
		self.assertEqual(row.custom_microsoft_status, files.FAILED)
		self.assertIn("SharePoint said no", row.custom_microsoft_error)
		self.assertEqual(row.custom_microsoft_attempts, 1)
		self.assertTrue(os.path.exists(f.get_full_path()))

	def test_a_long_file_name_is_recorded(self):
		# SharePoint's web URL percent-encodes every space: a 100-character name is well past 140.
		name = "Quotation for cold storage panels, Pune warehouse, revision 3, signed and returned.csv"
		f = self.attach(name=name)
		files.upload_file(f.name)
		row = frappe.db.get_value(
			"File", f.name, ["custom_microsoft_status", "custom_microsoft_web_url"], as_dict=True
		)
		self.assertEqual(row.custom_microsoft_status, files.STORED)
		self.assertGreater(len(row.custom_microsoft_web_url), 140)

	def test_rerunning_a_stored_file_does_nothing(self):
		f = self.attach()
		files.upload_file(f.name)
		calls = len(self.graph.calls)
		files.upload_file(f.name)
		self.assertEqual(len(self.graph.calls), calls)

	def test_storage_summary_counts_by_status(self):
		f = self.attach()
		files.upload_file(f.name)
		self.attach()  # left Pending: upload is mocked out
		summary = files.storage_summary()
		self.assertGreaterEqual(summary.get(files.STORED, 0), 1)
		self.assertGreaterEqual(summary.get(files.PENDING, 0), 1)

	def test_get_content_reads_from_sharepoint(self):
		f = self.attach(content=b"hello from teams", name="note.txt")
		files.upload_file(f.name)
		with patch.object(files, "fetch_content", return_value=b"hello from teams") as fetch:
			self.assertEqual(frappe.get_doc("File", f.name).get_content(), "hello from teams")
		fetch.assert_called_once_with("DRIVE", "ITEM-" + f.file_name)

	def test_deleting_the_record_keeps_the_sharepoint_folder_but_drops_the_link(self):
		f = self.attach()
		files.upload_file(f.name)
		self.assertTrue(files.get_folder(MAPPED, self.todo.name))
		frappe.delete_doc(MAPPED, self.todo.name, force=True)
		self.assertIsNone(files.get_folder(MAPPED, self.todo.name))
		self.assertNotIn("DELETE", [m for m, _p in self.graph.calls])

	def test_removing_an_attachment_leaves_sharepoint_alone_by_default(self):
		f = self.attach()
		files.upload_file(f.name)
		frappe.delete_doc("File", f.name, force=True)
		self.assertNotIn("DELETE", [m for m, _p in self.graph.calls])


# --- opening -------------------------------------------------------------------------------


class TestOpen(FilesTestCase):
	def test_open_file_streams_for_a_reader(self):
		f = self.attach()
		files.upload_file(f.name)
		with patch.object(files, "_stream", return_value="STREAM") as stream:
			self.assertEqual(files.open_file(f.name), "STREAM")
		stream.assert_called_once_with("DRIVE", "ITEM-" + f.file_name, f.file_name, download=0)

	def test_guests_are_refused(self):
		f = self.attach()
		files.upload_file(f.name)
		frappe.set_user("Guest")
		self.addCleanup(frappe.set_user, "Administrator")
		with self.assertRaises(frappe.PermissionError):
			files.open_file(f.name)

	def test_unknown_file_is_refused(self):
		with self.assertRaises(frappe.PermissionError):
			files.open_file("does-not-exist")

	def test_items_outside_the_records_folder_are_refused(self):
		files.ensure_folder(MAPPED, self.todo.name)
		folder = files.get_folder(MAPPED, self.todo.name)
		self.graph.items.update(
			{
				"INSIDE": {"parent": "SUB"},
				"SUB": {"parent": "FOLDER"},
				"ELSEWHERE": {"parent": "BASE"},
			}
		)
		files._assert_inside(folder, "INSIDE")  # two levels down: fine
		with self.assertRaises(frappe.PermissionError):
			files._assert_inside(folder, "ELSEWHERE")

	def test_force_download_types_are_never_inline(self):
		resp = MagicMock()
		resp.headers = {}
		resp.iter_content.return_value = iter([b"<svg/>"])
		with patch.object(graph, "graph_request", return_value=resp):
			out = files._stream("DRIVE", "X", "logo.svg")
		self.assertTrue(out.headers["Content-Disposition"].startswith("attachment"))
		self.assertEqual(out.headers["X-Content-Type-Options"], "nosniff")


# --- settings and doctor -------------------------------------------------------------------


class TestSettings(BaseTestCase):
	def test_changing_the_site_url_forgets_the_resolved_ids(self):
		mapping = configure(self)
		mapping.site_url = "https://contoso.sharepoint.com/sites/Other"
		mapping.save()
		mapping.reload()
		self.assertFalse(mapping.site_id)
		self.assertFalse(mapping.drive_id)

	def test_a_doctype_can_be_mapped_once(self):
		configure(self)
		with self.assertRaises(frappe.DuplicateEntryError):
			frappe.get_doc(
				{
					"doctype": "SharePoint Mapping",
					"reference_doctype": MAPPED,
					"site_url": "https://contoso.sharepoint.com/sites/X",
				}
			).insert()

	def test_child_tables_cannot_be_mapped(self):
		with self.assertRaises(frappe.ValidationError):
			frappe.get_doc(
				{
					"doctype": "SharePoint Mapping",
					"reference_doctype": "Has Role",
					"site_url": "https://contoso.sharepoint.com/sites/X",
				}
			).insert()

	def test_saving_a_mapping_refreshes_the_cache(self):
		mapping = configure(self)
		self.assertEqual(files.mapping_for(MAPPED).base_folder, "Projects")
		mapping.base_folder = "Archive"
		mapping.save()
		self.assertEqual(files.mapping_for(MAPPED).base_folder, "Archive")

	def test_doctype_picker_offers_only_mapped_doctypes(self):
		configure(self)
		self.assertEqual(files.mapped_doctype_query("DocType", "to", "name", 0, 20, {}), [[MAPPED]])
		self.assertEqual(files.mapped_doctype_query("DocType", "zzz", "name", 0, 20, {}), [])

	def test_boot_lists_mapped_doctypes(self):
		configure(self)
		boot = frappe._dict()
		files.boot_session(boot)
		self.assertEqual(boot.microsoft365_files_doctypes, [MAPPED])

	def test_addresses_names_and_ids_are_not_cut_at_140(self):
		# Frappe's Data default is varchar(140). SharePoint addresses, paths and names pass it
		# routinely, and the failed write came after the work was done in SharePoint.
		from frappe_microsoft365.setup import create_file_custom_fields

		fields = [
			(f["fieldname"], f.get("length"))
			for f in create_file_custom_fields(only_definitions=True)["File"]
			if f["fieldtype"] == "Data"
		]
		for doctype in ("SharePoint Mapping", "SharePoint Folder"):
			fields += [
				(f"{doctype}.{df.fieldname}", df.length)
				for df in frappe.get_meta(doctype).fields
				if df.fieldtype == "Data"
			]
		for fieldname, length in fields:
			self.assertGreaterEqual(length or 140, 255, fieldname)


class TestUninstall(FilesTestCase):
	def _moved(self, content=None):
		f = self.attach(content=content, name="report.csv")
		old = f.file_url
		files.upload_file(f.name)
		self.assertFalse(os.path.exists(files.local_path(old)))
		return f, old

	def _download(self, content):
		resp = MagicMock()
		resp.iter_content.return_value = iter([content])
		return resp

	def test_a_moved_file_comes_back_to_its_old_place(self):
		body = f"bring me home {frappe.generate_hash()}".encode()
		f, old = self._moved(body)
		with patch.object(graph, "graph_request", return_value=self._download(body)):
			self.assertEqual(files.bring_back(f.name), old)
		row = frappe.db.get_value(
			"File", f.name, ["file_url", "custom_microsoft_status", "content_hash"], as_dict=True
		)
		self.assertEqual(row.file_url, old)
		self.assertIsNone(row.custom_microsoft_status)
		self.assertTrue(row.content_hash)
		with open(files.local_path(old), "rb") as fh:
			self.assertEqual(fh.read(), body)
		comment = frappe.get_all(
			"Comment",
			filters={"reference_name": self.todo.name, "comment_type": "Attachment"},
			pluck="content",
		)
		self.assertTrue(any(old in c for c in comment))
		self.assertFalse(any(files.stored_url(f.name) in c for c in comment))

	def test_a_kept_local_copy_is_reused_without_downloading(self):
		configure(self, files_keep_local_copy=1)
		f = self.attach(name="kept.csv")
		old = f.file_url
		files.upload_file(f.name)
		with patch.object(graph, "graph_request") as gr:
			self.assertEqual(files.bring_back(f.name), old)
		gr.assert_not_called()

	def _uninstall_as(self, how):
		from frappe_microsoft365 import uninstall

		return patch.object(uninstall, "choice", return_value=how)

	def test_uninstall_stops_until_someone_chooses(self):
		from frappe_microsoft365 import uninstall

		self._moved()
		with self._uninstall_as(None):
			with self.assertRaises(frappe.ValidationError):
				uninstall.handle_files()

	def test_uninstall_will_not_fill_the_disk(self):
		from frappe_microsoft365 import uninstall

		self._moved()
		with (
			self._uninstall_as(uninstall.RESTORE),
			patch.object(uninstall, "free_space", return_value=1),
			patch.object(files, "bring_back") as back,
		):
			with self.assertRaises(frappe.ValidationError):
				uninstall.handle_files()
		back.assert_not_called()

	def test_leaving_files_in_sharepoint_turns_them_into_links(self):
		from frappe_microsoft365 import uninstall

		f, _old = self._moved()
		with self._uninstall_as(uninstall.LEAVE), patch.object(files, "bring_back") as back:
			uninstall.handle_files()
		back.assert_not_called()
		row = frappe.db.get_value("File", f.name, ["file_url", "custom_microsoft_status"], as_dict=True)
		self.assertTrue(row.file_url.startswith("https://contoso.sharepoint.com/"))
		self.assertIsNone(row.custom_microsoft_status)
		comments = frappe.get_all(
			"Comment",
			filters={"reference_name": self.todo.name, "comment_type": "Attachment"},
			pluck="content",
		)
		self.assertFalse(any(files.stored_url(f.name) in c for c in comments))

	def test_uninstall_stops_if_a_file_cannot_be_handled(self):
		from frappe_microsoft365 import uninstall

		self._moved()
		with (
			self._uninstall_as(uninstall.RESTORE),
			patch.object(uninstall, "free_space", return_value=10**12),
			patch.object(files, "bring_back", side_effect=MsGraphError("offline")),
			patch.object(frappe.db, "rollback"),
		):
			with self.assertRaises(frappe.ValidationError):
				uninstall.handle_files()

	def test_choice_follows_settings_then_site_config(self):
		from frappe_microsoft365 import uninstall

		configure(self, files_on_uninstall=uninstall.LEAVE)
		with patch.object(uninstall.sys, "stdin", None):
			self.assertEqual(uninstall.choice(), uninstall.LEAVE)
			with patch.dict(frappe.conf, {"microsoft365_uninstall_files": "restore"}):
				self.assertEqual(uninstall.choice(), uninstall.RESTORE)

	def test_dry_run_changes_nothing(self):
		from frappe_microsoft365 import uninstall

		self._moved()

		def remove_app(dry_run=True):
			uninstall.before_uninstall()

		with patch.object(uninstall, "choice") as asked, patch.object(uninstall, "clear_caches") as clear:
			remove_app()
		asked.assert_not_called()
		clear.assert_not_called()
		self.assertTrue(frappe.db.has_column("File", "custom_microsoft_status"))

	def test_every_custom_field_is_known_to_uninstall(self):
		from frappe_microsoft365.setup import app_custom_fields

		known = app_custom_fields()
		for doctype, fieldnames in known.items():
			for fieldname in fieldnames:
				self.assertTrue(
					frappe.db.exists("Custom Field", {"dt": doctype, "fieldname": fieldname}), fieldname
				)
		self.assertIn("custom_microsoft_status", known["File"])
		self.assertIn("custom_microsoft_event_id", known["Event"])


class TestOnDemand(FilesTestCase):
	def setUp(self):
		super().setUp()
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.ON_DEMAND
		mapping.save()

	def test_attachments_stay_put_until_the_record_has_a_folder(self):
		f = self.attach()
		self.enqueued.assert_not_called()
		self.assertFalse(frappe.db.get_value("File", f.name, "custom_microsoft_status"))
		self.assertIsNone(files.get_folder(MAPPED, self.todo.name))

	def test_create_folder_now_sends_the_records_files(self):
		f = self.attach()
		out = files.create_folder(MAPPED, self.todo.name)
		self.assertEqual(out["queued"], 1)
		self.enqueued.assert_called_once_with(f.name)
		# and from then on, new attachments follow
		g = self.attach()
		self.enqueued.assert_called_with(g.name)

	def test_panel_says_files_stay_here(self):
		self.assertTrue(files.list_folder(MAPPED, self.todo.name)["on_demand"])

	def add_row(self, **row):
		"""Add Row on the mapping's Folders table, then Save."""
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.append("folders", {"reference_name": self.todo.name, **row})
		mapping.save()
		return frappe.get_doc(
			"SharePoint Folder",
			{"parent": MAPPED, "reference_name": row.get("reference_name", self.todo.name)},
		)

	def test_a_folder_can_be_linked_by_path(self):
		f = self.attach()
		folder = self.add_row(existing_folder="Projects/Already there")
		self.assertEqual(folder.item_id, "FOLDER")
		self.assertEqual(folder.drive_id, "DRIVE")
		self.assertEqual(folder.reference_doctype, MAPPED)
		self.enqueued.assert_called_once_with(f.name)

	def test_a_folder_can_be_linked_by_its_sharepoint_address(self):
		folder = self.add_row(
			existing_folder="https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSales%2FShared%20Documents%2FProjects%2FOld"
		)
		self.assertEqual(folder.item_id, "FOLDER")
		self.assertIn(("GET", "/drives/DRIVE/root:/Projects/Old:"), self.graph.calls)

	def test_a_new_folder_is_created_when_none_is_given(self):
		folder = self.add_row()
		self.assertEqual(folder.folder_path, f"Projects/{files.safe_name(self.todo.name)}")

	def test_a_record_gets_one_folder(self):
		files.create_folder(MAPPED, self.todo.name)
		with self.assertRaises(frappe.DuplicateEntryError):
			self.add_row()

	def test_the_app_adds_its_rows_to_the_mapping(self):
		files.create_folder(MAPPED, self.todo.name)
		rows = frappe.get_doc("SharePoint Mapping", MAPPED).folders
		self.assertEqual([r.reference_name for r in rows], [self.todo.name])
		self.assertTrue(rows[0].item_id)

	def test_a_mapping_opened_before_a_folder_was_added_cannot_save_over_it(self):
		stale = frappe.get_doc("SharePoint Mapping", MAPPED)
		frappe.db.set_value(
			"SharePoint Mapping",
			MAPPED,
			"modified",
			add_to_date(stale.modified, seconds=-5),
			update_modified=False,
		)
		stale = frappe.get_doc("SharePoint Mapping", MAPPED)
		files.create_folder(MAPPED, self.todo.name)
		stale.base_folder = "Elsewhere"
		with self.assertRaises(frappe.TimestampMismatchError):
			stale.save()
		self.assertTrue(files.get_folder(MAPPED, self.todo.name))

	def test_deleting_a_row_forgets_the_folder_and_leaves_sharepoint_alone(self):
		files.create_folder(MAPPED, self.todo.name)
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.set("folders", [])
		mapping.save()
		self.assertIsNone(files.get_folder(MAPPED, self.todo.name))
		self.assertNotIn("DELETE", [m for m, _p in self.graph.calls])


class TestMappingOverrides(FilesTestCase):
	"""A mapping's Options override SharePoint Settings for its own DocType only."""

	def set_override(self, field, value):
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.set(field, value)
		mapping.save()

	def test_default_follows_sharepoint_settings(self):
		self.assertEqual(files.storage_option("files_keep_local_copy", MAPPED), 0)
		frappe.db.set_single_value("SharePoint Settings", "files_keep_local_copy", 1)
		frappe.clear_document_cache("SharePoint Settings", "SharePoint Settings")
		self.assertEqual(files.storage_option("files_keep_local_copy", MAPPED), 1)

	def test_on_and_off_win_over_the_settings(self):
		self.set_override("keep_local_copy", "On")
		self.assertEqual(files.storage_option("files_keep_local_copy", MAPPED), 1)
		self.assertEqual(files.storage_option("files_keep_local_copy", "Note"), 0)
		frappe.db.set_single_value("SharePoint Settings", "files_delete_remote", 1)
		frappe.clear_document_cache("SharePoint Settings", "SharePoint Settings")
		self.set_override("delete_remote", "Off")
		self.assertEqual(files.storage_option("files_delete_remote", MAPPED), 0)

	def test_a_kept_local_copy_follows_the_mapping(self):
		self.set_override("keep_local_copy", "On")
		f = self.attach()
		files.upload_file(f.name)
		f.reload()
		self.assertEqual(f.custom_microsoft_status, files.STORED)
		self.assertTrue(os.path.exists(files.local_path(f.custom_microsoft_local_url)))

	def test_off_keeps_the_sharepoint_copy_even_when_settings_say_remove(self):
		configure(self, files_delete_remote=1)
		self.set_override("delete_remote", "Off")
		f = self.attach()
		files.upload_file(f.name)
		frappe.delete_doc("File", f.name, force=True)
		self.assertNotIn("DELETE", [m for m, _p in self.graph.calls])

	def test_on_removes_it_even_when_settings_say_keep(self):
		self.set_override("delete_remote", "On")
		f = self.attach()
		files.upload_file(f.name)
		frappe.delete_doc("File", f.name, force=True)
		self.assertIn("DELETE", [m for m, _p in self.graph.calls])


class TestFolderPermissions(FilesTestCase):
	STAFF = "ms365-staff@example.com"

	ROLE = "SharePoint Folder Test Role"

	def setUp(self):
		super().setUp()
		if not frappe.db.exists("Role", self.ROLE):
			frappe.get_doc({"doctype": "Role", "role_name": self.ROLE, "desk_access": 1}).insert(
				ignore_permissions=True
			)
		if not frappe.db.exists("User", self.STAFF):
			frappe.get_doc(
				{"doctype": "User", "email": self.STAFF, "first_name": "Staff", "send_welcome_email": 0}
			).insert(ignore_permissions=True)
		user = frappe.get_doc("User", self.STAFF)
		user.set("roles", [])
		user.append("roles", {"role": "Desk User"})
		user.append("roles", {"role": self.ROLE})
		user.save(ignore_permissions=True)
		frappe.set_user(self.STAFF)
		self.addCleanup(frappe.set_user, "Administrator")
		# Assigned to them: newer Frappe lets only the assignee (or a ToDo role) read a ToDo.
		self.mine = frappe.get_doc(
			{"doctype": "ToDo", "description": "staff's own", "allocated_to": self.STAFF}
		).insert()
		frappe.set_user("Administrator")

	def test_system_managers_can_create_folders(self):
		self.assertTrue(files.can_create_folder(MAPPED, self.mine.name, user="Administrator"))

	def test_other_users_cannot_by_default(self):
		self.assertTrue(frappe.has_permission(MAPPED, "write", doc=self.mine.name, user=self.STAFF))
		self.assertFalse(files.can_create_folder(MAPPED, self.mine.name, user=self.STAFF))
		frappe.set_user(self.STAFF)
		self.assertFalse(files.list_folder(MAPPED, self.mine.name)["can_create"])
		with self.assertRaises(frappe.PermissionError):
			files.create_folder(MAPPED, self.mine.name)

	def test_a_listed_role_may_create_folders(self):
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.append("folder_roles", {"role": self.ROLE})
		mapping.save()
		self.assertTrue(files.can_create_folder(MAPPED, self.mine.name, user=self.STAFF))

	def test_a_listed_role_still_needs_edit_rights_on_the_record(self):
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.append("folder_roles", {"role": self.ROLE})
		mapping.save()
		others = frappe.get_doc({"doctype": "ToDo", "description": "admin's"}).insert()
		if not frappe.has_permission(MAPPED, "write", doc=others.name, user=self.STAFF):
			self.assertFalse(files.can_create_folder(MAPPED, others.name, user=self.STAFF))


class TestQuietLookups(BaseTestCase):
	def test_an_expected_404_leaves_no_message_for_the_desk(self):
		def not_found(*args, **kwargs):
			frappe.throw(
				"Microsoft Graph GET /drives/d/root:/Projects: failed (404): itemNotFound", MsGraphNotFound
			)

		frappe.local.message_log = []
		with patch.object(graph, "graph_request", side_effect=not_found):
			self.assertIsNone(files._get_by_path("d", ["Projects"]))
		self.assertEqual(frappe.local.message_log, [])


class TestTokenRoles(BaseTestCase):
	@staticmethod
	def _token(payload):
		import base64
		import json

		body = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")
		return f"header.{body}.signature"

	def test_roles_are_read_from_the_payload(self):
		self.assertEqual(files.token_roles(self._token({"roles": ["Sites.Selected"]})), ["Sites.Selected"])

	def test_no_roles_or_garbage_gives_empty(self):
		self.assertEqual(files.token_roles(self._token({"aud": "x"})), [])
		self.assertEqual(files.token_roles("not-a-jwt"), [])


class TestDoctorFiles(BaseTestCase):
	def test_nothing_reported_when_files_is_off(self):
		self.assertEqual(doctor.check_files({"use_files": 0}, []), [])

	def test_on_without_mappings_fails(self):
		out = doctor.check_files({"use_files": 1, "tenant_id": "t"}, [])
		self.assertEqual([f["check"] for f in out], ["files.mappings"])
		self.assertEqual(out[0]["status"], doctor.FAIL)

	def test_common_tenant_cannot_do_app_only(self):
		out = doctor.check_files(
			{"use_files": 1, "tenant_id": "common"},
			[{"enabled": 1, "reference_doctype": "Project", "site_url": "https://c.sharepoint.com/sites/S"}],
		)
		self.assertIn("files.tenant", [f["check"] for f in out])

	def test_failed_files_are_reported(self):
		out = doctor.check_files(
			{"use_files": 1, "tenant_id": "t"},
			[{"enabled": 1, "reference_doctype": "Project", "site_url": "https://c.sharepoint.com/sites/S"}],
			failed=3,
		)
		self.assertEqual([f["check"] for f in out], ["files.failed"])

	def test_site_grant_script_names_each_site_and_only_write(self):
		script = doctor.powershell_for_site_grant(
			"APP-ID",
			["https://c.sharepoint.com/sites/Sales/Shared%20Documents", "https://c.sharepoint.com/teams/Ops"],
		)
		self.assertIn('Get-MgSite -SiteId "c.sharepoint.com:/sites/Sales"', script)
		self.assertIn('Get-MgSite -SiteId "c.sharepoint.com:/teams/Ops"', script)
		self.assertIn('roles = @("write")', script)
		self.assertIn('$appId = "APP-ID"', script)
		self.assertNotIn('FullControl")', script.split("Connect-MgGraph", 1)[1].split("\n", 1)[1])

	def test_explain_recognises_the_sharepoint_401(self):
		out = doctor.explain_error(
			"Microsoft Graph GET /sites/contoso.sharepoint.com:/sites/Example failed (401): "
			"generalException: General exception while processing"
		)
		self.assertTrue(out["matched"])
		self.assertIn("Microsoft Graph", out["detail"])

	def test_setup_guide_lists_sites_selected(self):
		summary = doctor._capability_summary({"use_files": 1})
		self.assertEqual(summary[-1]["permissions"], ["Sites.Selected (Application)"])
		steps = [s["id"] for s in doctor.manual_setup_steps({"use_files": 1})]
		self.assertIn("files_site_grant", steps)
		self.assertIn("admin_consent", steps)


class TestOffMeansOff(BaseTestCase):
	"""A site that never switched storage on gets no File fields and no behaviour change."""

	def test_old_file_links_are_left_alone_without_the_fields(self):
		request = SimpleNamespace(method="GET", path="/private/files/nowhere-to-be-found.pdf")
		with (
			patch.object(frappe.local, "request", request, create=True),
			patch.object(files, "has_file_fields", return_value=False),
			patch.object(frappe.db, "get_value") as lookup,
		):
			files.redirect_moved_file()
		lookup.assert_not_called()

	def test_summary_is_empty_without_the_fields(self):
		with patch.object(files, "has_file_fields", return_value=False):
			self.assertEqual(files.storage_summary(), {})

	def test_hooks_do_nothing_when_off(self):
		off = frappe._dict(enabled=1, use_files=0)
		with patch.object(files, "_settings", return_value=off):
			self.assertEqual(files.mapped_doctypes(), [])
			self.assertFalse(files._eligible(frappe._dict(attached_to_doctype="ToDo", attached_to_name="x")))

	def test_switching_storage_on_creates_the_fields(self):
		with patch("frappe_microsoft365.setup.create_file_custom_fields") as create:
			from frappe_microsoft365 import capabilities

			capabilities.apply_changes({"use_files": 0}, {"use_files": 1})
		create.assert_called_once()


class TestInlineTypes(BaseTestCase):
	def test_safe_types_open_inline(self):
		for content_type in ("application/pdf", "image/png", "text/plain", "video/mp4"):
			self.assertTrue(files.opens_inline(content_type), content_type)

	def test_everything_else_downloads(self):
		for content_type in (
			None,
			"application/octet-stream",
			"text/html",
			"image/svg+xml",
			"application/xml",
			"application/javascript",
		):
			self.assertFalse(files.opens_inline(content_type), content_type)

	def test_extensionless_file_downloads_whatever_sharepoint_says(self):
		resp = MagicMock(headers={"Content-Type": "text/html"})
		resp.iter_content.return_value = iter([b"<script>alert(1)</script>"])
		with patch.object(graph, "graph_request", return_value=resp):
			response = files._stream("DRIVE", "ITEM", "report")
		self.assertIn("attachment", response.headers["Content-Disposition"])
		self.assertEqual(response.content_type, "application/octet-stream")
