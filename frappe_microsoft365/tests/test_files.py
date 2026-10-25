"""SharePoint document storage: naming, eligibility, uploads, opening, permissions, doctor checks.

Nothing here touches the network. Graph calls go through ``microsoft_graph.graph_request`` (and
upload ranges through ``microsoft_graph._http_request``), both replaced with fakes that answer
like Graph does.
"""

from contextlib import contextmanager
from itertools import pairwise
import os
from types import SimpleNamespace
from unittest.mock import MagicMock, patch
from urllib.parse import quote, unquote

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
		self.web_url_padding = ""

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
		if method == "GET" and path.startswith("/drives/DRIVE/items/FOLDER:/") and params:
			name = unquote(path.split("FOLDER:/", 1)[1].rsplit(":", 1)[0])
			for item in self.items.values():
				if item.get("name") == name and item.get("parent") == "FOLDER":
					return {"id": item["id"], "name": name, "file": {}, "webUrl": item.get("webUrl"), "parentReference": {"driveId": "DRIVE", "id": "FOLDER"}}
			raise MsGraphNotFound("404")
		if method == "GET" and ":/Projects/" in path:
			return {"id": "FOLDER", "folder": {}, "webUrl": "https://contoso.sharepoint.com/f"}
		if method == "PUT" and path.endswith(":/content"):
			if self.fail_upload:
				raise self.fail_upload
			name = unquote(path.split(":/", 1)[1].rsplit(":/content", 1)[0])
			parent = path.split("/items/", 1)[1].split(":/", 1)[0]
			web_url = "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Projects/FOLDER/" + self.web_url_padding + quote(name, safe="")
			self.items["ITEM-" + name] = {"id": "ITEM-" + name, "name": name, "parent": parent, "webUrl": web_url}
			return {
				"id": "ITEM-" + name,
				"webUrl": web_url,
				"parentReference": {"driveId": "DRIVE", "id": parent},
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
				"webUrl": item.get("webUrl"),
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


_TEST_STORAGE_TARGET = {}
_TEST_TARGETS = {}


def resolve_test_storage_target(doctype, name):
	target = _TEST_TARGETS.get(name)
	if not target:
		return None
	return {"anchor_doctype": doctype, "anchor_name": target[0], "path_segments": target[1]}


def resolve_mutable_test_storage_target(doctype, name):
	if name != _TEST_STORAGE_TARGET["source"]:
		return None
	return {"anchor_doctype": doctype, "anchor_name": _TEST_STORAGE_TARGET["anchor"]}


def resolve_invalid_test_storage_target(doctype, name):
	return {"anchor_name": name}


def resolve_missing_test_storage_target(doctype, name):
	return {"anchor_doctype": doctype, "anchor_name": "does-not-exist"}


def resolve_self_test_storage_target(doctype, name):
	return {"anchor_doctype": doctype, "anchor_name": name}


def resolve_raising_test_storage_target(doctype, name):
	raise RuntimeError("test provider failure")


def resolve_recursive_test_storage_target(doctype, name):
	from frappe_microsoft365 import storage_target

	return storage_target.resolve_storage_target(doctype, name)


@contextmanager
def test_storage_providers(providers):
	real_get_hooks = frappe.get_hooks
	with patch.object(
		frappe,
		"get_hooks",
		side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers
		if key == "microsoft365_storage_target_resolver"
		else real_get_hooks(key, default=default, **kwargs),
	):
		yield


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

	def test_pattern_title_is_the_doctypes_title_field(self):
		note = frappe.get_doc({"doctype": "Note", "title": "Scan-Line 920B"})
		note.name = "DEV-00042"
		self.assertEqual(files.render_pattern("{name} - {title}", note), "DEV-00042 - Scan-Line 920B")
		# No title: the name stands in, so the folder is never called just "DEV-00042 - ".
		note.title = None
		self.assertEqual(files.render_pattern("{title}", note), "DEV-00042")

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


class TestStorageTargetResolver(FilesTestCase):
	def test_no_provider_preserves_source_as_default_anchor(self):
		from frappe_microsoft365 import storage_target

		target = storage_target.resolve_storage_target(MAPPED, self.todo.name)
		self.assertEqual((target.source_doctype, target.source_name), (MAPPED, self.todo.name))
		self.assertEqual((target.anchor_doctype, target.anchor_name), (MAPPED, self.todo.name))
		self.assertIsNone(target.path_segments)
		self.assertFalse(target.dynamic)

	def test_multiple_providers_fail_closed(self):
		from frappe_microsoft365 import storage_target

		provider = "frappe_microsoft365.tests.test_files.resolve_test_storage_target"
		with test_storage_providers({MAPPED: [provider, provider]}), self.assertRaises(storage_target.StorageTargetError):
			storage_target.resolve_storage_target(MAPPED, self.todo.name)

	def test_invalid_response_missing_anchor_and_self_target_fail_closed(self):
		from frappe_microsoft365 import storage_target

		for provider in (
			"frappe_microsoft365.tests.test_files.resolve_invalid_test_storage_target",
			"frappe_microsoft365.tests.test_files.resolve_missing_test_storage_target",
			"frappe_microsoft365.tests.test_files.resolve_self_test_storage_target",
		):
			with self.subTest(provider=provider), test_storage_providers({MAPPED: provider}), self.assertRaises(
				storage_target.StorageTargetError
			):
				storage_target.resolve_storage_target(MAPPED, self.todo.name)

	def test_provider_exception_and_recursive_resolution_fail_closed(self):
		from frappe_microsoft365 import storage_target

		for provider in (
			"frappe_microsoft365.tests.test_files.resolve_raising_test_storage_target",
			"frappe_microsoft365.tests.test_files.resolve_recursive_test_storage_target",
		):
			with self.subTest(provider=provider), test_storage_providers({MAPPED: provider}), self.assertRaises(
				storage_target.StorageTargetError
			):
				storage_target.resolve_storage_target(MAPPED, self.todo.name)


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

	def test_dynamic_files_sharing_one_anchor_share_one_persisted_folder(self):
		from frappe_microsoft365 import storage_target

		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.AUTOMATIC
		mapping.save()
		anchor = frappe.get_doc({"doctype": MAPPED, "description": "canonical anchor"}).insert()
		second_source = frappe.get_doc({"doctype": MAPPED, "description": "another source"}).insert()
		providers = {
			MAPPED: "frappe_microsoft365.tests.test_files.resolve_test_storage_target",
		}
		self.addCleanup(setattr, __import__(__name__, fromlist=["_TEST_TARGETS"]), "_TEST_TARGETS", {})
		_TEST_TARGETS.update({self.todo.name: (anchor.name, None), second_source.name: (anchor.name, None)})
		get_hooks = frappe.get_hooks
		with patch.object(frappe, "get_hooks", side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers if key == "microsoft365_storage_target_resolver" else get_hooks(key, default=default, **kwargs)):
			first = self.attach(name="first.txt")
			second = self.attach(name="second.txt", attached_to_name=second_source.name)
			files.upload_file(first.name)
			files.upload_file(second.name)
			first_binding = storage_target.parse_binding(
				frappe.db.get_value("File", first.name, storage_target.BINDING_FIELD)
			)
			second_binding = storage_target.parse_binding(
				frappe.db.get_value("File", second.name, storage_target.BINDING_FIELD)
			)
		self.assertIsNotNone(
			first_binding,
			frappe.db.get_value("File", first.name, ["custom_microsoft_status", "custom_microsoft_error"], as_dict=True),
		)
		self.assertEqual(first_binding["anchor"], [MAPPED, anchor.name])
		self.assertEqual(second_binding["anchor"], [MAPPED, anchor.name])
		self.assertEqual(first_binding["path"], {"mode": "default"})
		self.assertEqual(first_binding["folder"], {"drive_id": "DRIVE", "item_id": "FOLDER"})
		self.assertEqual(frappe.db.count("SharePoint Folder", {"reference_doctype": MAPPED, "reference_name": anchor.name}), 1)

	def test_resolver_relink_to_same_physical_folder_fails_closed(self):
		from frappe_microsoft365 import storage_target

		first_anchor = frappe.get_doc({"doctype": MAPPED, "description": "anchor one"}).insert()
		second_anchor = frappe.get_doc({"doctype": MAPPED, "description": "anchor two"}).insert()
		state = {"source": self.todo.name, "anchor": first_anchor.name}
		providers = {MAPPED: "frappe_microsoft365.tests.test_files.resolve_mutable_test_storage_target"}
		get_hooks = frappe.get_hooks
		with patch.object(frappe, "get_hooks", side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers if key == "microsoft365_storage_target_resolver" else get_hooks(key, default=default, **kwargs)), patch(
			"frappe_microsoft365.tests.test_files._TEST_STORAGE_TARGET", state
		):
			f = self.attach()
			files.upload_file(f.name)
			# Different logical identity, deliberately the same physical drive+folder.
			files.add_folder_row(MAPPED, second_anchor.name, {"folder_name": "same", "drive_id": "DRIVE", "item_id": "FOLDER", "web_url": "https://contoso.sharepoint.com/f", "folder_path": "Projects/same"})
			state["anchor"] = second_anchor.name
			with self.assertRaises(storage_target.StorageTargetError):
				files.assert_stored_file_access(frappe.get_doc("File", f.name))
			with self.assertRaises(storage_target.StorageTargetError):
				storage_target.assert_binding_matches(frappe.get_doc("File", f.name), None)

	def test_dynamic_retry_after_graph_accepts_file_reuses_the_same_item(self):
		from frappe_microsoft365 import storage_target

		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.AUTOMATIC
		mapping.save()
		anchor = frappe.get_doc({"doctype": MAPPED, "description": "retry anchor"}).insert()
		_TEST_TARGETS[self.todo.name] = (anchor.name, None)
		providers = {MAPPED: "frappe_microsoft365.tests.test_files.resolve_test_storage_target"}
		get_hooks = frappe.get_hooks
		class Killed(BaseException):
			pass
		real_stream = files.upload_stream
		def stream_then_die(*args, **kwargs):
			real_stream(*args, **kwargs)
			raise Killed()
		with patch.object(frappe, "get_hooks", side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers if key == "microsoft365_storage_target_resolver" else get_hooks(key, default=default, **kwargs)), patch.object(files, "upload_stream", side_effect=stream_then_die):
			f = self.attach()
			with self.assertRaises(Killed):
				files.upload_file(f.name)
			binding = storage_target.parse_binding(frappe.db.get_value("File", f.name, storage_target.BINDING_FIELD))
			self.assertEqual(binding["folder"], {"drive_id": "DRIVE", "item_id": "FOLDER"})
			files.upload_file(f.name)
		self.assertEqual(len([path for method, path in self.graph.calls if method == "PUT"]), 1)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.STORED)

	def test_dynamic_failed_file_can_be_requeued_with_its_stored_binding(self):
		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.AUTOMATIC
		mapping.save()
		anchor = frappe.get_doc({"doctype": MAPPED, "description": "retry queue anchor"}).insert()
		_TEST_TARGETS[self.todo.name] = (anchor.name, None)
		providers = {MAPPED: "frappe_microsoft365.tests.test_files.resolve_test_storage_target"}
		with test_storage_providers(providers):
			f = self.attach()
			self.graph.fail_upload = MsGraphError("temporary Graph outage")
			files.upload_file(f.name)
			self.graph.fail_upload = None
			self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.FAILED)
			self.assertEqual(files.queue_record(MAPPED, self.todo.name), 1)
		self.assertEqual(self.enqueued.call_count, 2)
		self.enqueued.assert_called_with(f.name)

	def test_copy_of_stored_dynamic_file_is_denied_and_unbound(self):
		from frappe_microsoft365 import storage_target

		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.AUTOMATIC
		mapping.save()
		anchor = frappe.get_doc({"doctype": MAPPED, "description": "copy anchor"}).insert()
		_TEST_TARGETS[self.todo.name] = (anchor.name, None)
		providers = {MAPPED: "frappe_microsoft365.tests.test_files.resolve_test_storage_target"}
		get_hooks = frappe.get_hooks
		with patch.object(frappe, "get_hooks", side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers if key == "microsoft365_storage_target_resolver" else get_hooks(key, default=default, **kwargs)):
			original = self.attach()
			files.upload_file(original.name)
			copy_target = frappe.get_doc({"doctype": MAPPED, "description": "copy destination"}).insert()
			copy = frappe.get_doc("File", original.name).create_attachment_copy(MAPPED, copy_target.name)
			self.assertIsNone(copy.get(storage_target.BINDING_FIELD))
			self.assertEqual(frappe.db.get_value("File", copy.name, "custom_microsoft_status"), files.FAILED)
			with self.assertRaises(storage_target.StorageTargetError):
				files.open_file(copy.name)

	def test_copy_of_stored_default_file_keeps_legacy_shared_item_behavior(self):
		original = self.attach()
		files.upload_file(original.name)
		destination = frappe.get_doc({"doctype": MAPPED, "description": "default copy destination"}).insert()
		copy = frappe.get_doc("File", original.name).create_attachment_copy(MAPPED, destination.name)
		row = frappe.get_doc("File", copy.name)
		self.assertEqual(row.get("custom_microsoft_status"), files.STORED)
		self.assertEqual(row.get("custom_microsoft_drive_id"), "DRIVE")
		self.assertEqual(row.get("custom_microsoft_item_id"), "ITEM-" + original.file_name)
		self.assertEqual(row.file_url, files.stored_url(original.name))
		self.assertFalse(row.get("custom_microsoft_storage_binding"))

	def test_dynamic_explicit_path_is_normalized_and_compared(self):
		from frappe_microsoft365 import storage_target

		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.AUTOMATIC
		mapping.save()
		anchor = frappe.get_doc({"doctype": MAPPED, "description": "explicit anchor"}).insert()
		providers = {MAPPED: "frappe_microsoft365.tests.test_files.resolve_test_storage_target"}
		_TEST_TARGETS.clear()
		_TEST_TARGETS[self.todo.name] = (anchor.name, ["Live Test", "ORD 1"])
		get_hooks = frappe.get_hooks
		with patch.object(frappe, "get_hooks", side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers if key == "microsoft365_storage_target_resolver" else get_hooks(key, default=default, **kwargs)):
			f = self.attach()
			files.upload_file(f.name)
			target = storage_target.resolve_storage_target(MAPPED, self.todo.name)
			binding = storage_target.parse_binding(frappe.db.get_value("File", f.name, storage_target.BINDING_FIELD))
			self.assertEqual(target.path_segments, ("Live Test", "ORD 1"))
			self.assertEqual(binding["path"], {"mode": "explicit", "segments": ["Live Test", "ORD 1"]})
			self.assertEqual(binding["folder"], {"drive_id": "DRIVE", "item_id": "FOLDER"})
			self.assertEqual(files.get_folder(MAPPED, anchor.name).folder_path, "Projects/Live Test/ORD 1")

	def test_bound_content_requires_source_and_anchor_read_permission(self):
		from frappe_microsoft365 import storage_target

		mapping = frappe.get_doc("SharePoint Mapping", MAPPED)
		mapping.folder_creation = files.AUTOMATIC
		mapping.save()
		anchor = frappe.get_doc({"doctype": MAPPED, "description": "permission anchor"}).insert()
		_TEST_TARGETS[self.todo.name] = (anchor.name, None)
		providers = {MAPPED: "frappe_microsoft365.tests.test_files.resolve_test_storage_target"}
		get_hooks = frappe.get_hooks
		with patch.object(frappe, "get_hooks", side_effect=lambda key=None, default="_KEEP_DEFAULT_LIST", **kwargs: providers if key == "microsoft365_storage_target_resolver" else get_hooks(key, default=default, **kwargs)):
			f = self.attach()
			files.upload_file(f.name)
			checks = []
			def permission(doctype, ptype, doc=None, user=None, throw=False):
				checks.append((doctype, ptype, doc))
				if doc == anchor.name:
					raise frappe.PermissionError("anchor denied")
				return True
			with patch.object(frappe, "has_permission", side_effect=permission), patch.object(files, "fetch_content") as fetch:
				with self.assertRaises(frappe.PermissionError):
					frappe.get_doc("File", f.name).get_content()
			fetch.assert_not_called()
			self.assertEqual(checks, [(MAPPED, "read", self.todo.name), (MAPPED, "read", anchor.name)])

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

	def uploads(self):
		return [path for method, path in self.graph.calls if method == "PUT"]

	def test_a_failure_after_the_upload_never_uploads_again(self):
		# Recording the upload fails (here: an address wider than the column, a TEXT's 64 KB)
		# after SharePoint has the bytes. Before this was handled, every hourly retry uploaded
		# another copy.
		self.graph.web_url_padding = "x" * 70000
		f = self.attach()
		files.upload_file(f.name)
		row = frappe.db.get_value(
			"File",
			f.name,
			["file_url", "custom_microsoft_status", "custom_microsoft_error", "custom_microsoft_item_id"],
			as_dict=True,
		)
		self.assertEqual(row.custom_microsoft_status, files.FAILED)
		self.assertIn("too long", row.custom_microsoft_error)
		self.assertEqual(row.custom_microsoft_item_id, "ITEM-" + f.file_name)  # the upload is known
		self.assertEqual(row.file_url, f.file_url)  # and nothing was moved off the disk
		self.assertTrue(os.path.exists(f.get_full_path()))

		self.graph.web_url_padding = ""
		self.graph.items["ITEM-" + f.file_name]["webUrl"] = "https://contoso.sharepoint.com/spec"
		files.upload_file(f.name)
		self.assertEqual(len(self.uploads()), 1)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_status"), files.STORED)

	def test_an_attempt_counts_even_when_the_job_dies(self):
		class Killed(BaseException):
			"""What a job timeout or a worker restart looks like from inside the job."""

		f = self.attach()
		with patch.object(files, "_upload", side_effect=Killed):
			with self.assertRaises(Killed):
				files.upload_file(f.name)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_attempts"), 1)

	def test_an_upload_no_longer_in_the_folder_is_sent_again(self):
		f = self.attach()
		self.graph.items["MOVED"] = {"id": "MOVED", "parent": "SOMEWHERE-ELSE"}
		frappe.db.set_value(
			"File", f.name, {"custom_microsoft_item_id": "MOVED", "custom_microsoft_drive_id": "DRIVE"}
		)
		files.upload_file(f.name)
		self.assertEqual(len(self.uploads()), 1)
		self.assertEqual(frappe.db.get_value("File", f.name, "custom_microsoft_item_id"), "ITEM-" + f.file_name)

	def test_a_copy_does_not_take_over_the_originals_upload(self):
		f = self.attach()
		frappe.db.set_value(
			"File",
			f.name,
			{
				"custom_microsoft_status": files.FAILED,
				"custom_microsoft_attempts": 5,
				"custom_microsoft_item_id": "ITEM-elsewhere",
				"custom_microsoft_drive_id": "DRIVE",
			},
		)
		other = frappe.get_doc({"doctype": "ToDo", "description": "copy target"}).insert()
		copy = frappe.get_doc("File", f.name).create_attachment_copy(MAPPED, other.name)
		row = frappe.db.get_value(
			"File",
			copy.name,
			["custom_microsoft_status", "custom_microsoft_attempts", "custom_microsoft_item_id"],
			as_dict=True,
		)
		self.assertEqual(row.custom_microsoft_status, files.PENDING)
		self.assertEqual(row.custom_microsoft_attempts, 0)
		self.assertIsNone(row.custom_microsoft_item_id)

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
		self.assertIn(MAPPED, boot.microsoft365_files_doctypes)

	def tab_fields(self):
		return frappe.get_all(
			"Custom Field",
			filters={"dt": MAPPED, "fieldname": ["in", files.FILES_TAB_FIELDS]},
			pluck="fieldname",
			order_by="fieldname",
		)

	def panel(self, mapping, placement):
		mapping.files_panel = placement
		mapping.save()
		self.addCleanup(files.sync_files_tab, MAPPED, False)

	def test_files_panel_defaults_to_the_top_of_the_form(self):
		configure(self)
		boot = frappe._dict()
		files.boot_session(boot)
		self.assertEqual(boot.microsoft365_files_panel[MAPPED], files.PANEL_TOP)
		self.assertEqual(self.tab_fields(), [])

	def test_own_tab_adds_a_tab_to_the_form(self):
		mapping = configure(self)
		self.panel(mapping, files.PANEL_TAB)
		self.assertEqual(self.tab_fields(), sorted(files.FILES_TAB_FIELDS))
		meta = frappe.get_meta(MAPPED)
		self.assertEqual(meta.get_field(files.FILES_TAB_FIELD).fieldtype, "Tab Break")
		self.assertEqual(meta.fields[-1].fieldname, files.FILES_HTML_FIELD)  # at the end of the form
		self.assertFalse(frappe.db.has_column(MAPPED, files.FILES_HTML_FIELD))
		boot = frappe._dict()
		files.boot_session(boot)
		self.assertEqual(boot.microsoft365_files_panel[MAPPED], files.PANEL_TAB)

		mapping.save()  # saving again changes nothing
		self.assertEqual(len(self.tab_fields()), 2)

	def test_the_tab_goes_when_it_is_no_longer_wanted(self):
		mapping = configure(self)
		self.panel(mapping, files.PANEL_TAB)
		self.panel(mapping, files.PANEL_HIDDEN)
		self.assertEqual(self.tab_fields(), [])

		self.panel(mapping, files.PANEL_TAB)
		mapping.enabled = 0
		mapping.save()
		self.assertEqual(self.tab_fields(), [])

		mapping.enabled = 1
		mapping.save()
		self.assertEqual(len(self.tab_fields()), 2)
		mapping.delete()
		self.assertEqual(self.tab_fields(), [])

	def test_uninstall_removes_the_tab_too(self):
		from frappe_microsoft365.setup import app_custom_fields

		mapping = configure(self)
		self.panel(mapping, files.PANEL_TAB)
		self.assertEqual(sorted(app_custom_fields()[MAPPED]), sorted(files.FILES_TAB_FIELDS))

	def seed_cached_boot(self):
		# Frappe keeps each user's boot until they reload; stand in for a desk already open.
		frappe.cache().hset("bootinfo", "someone@example.com", {"from_cache": 1})
		self.addCleanup(frappe.cache().hdel, "bootinfo", "someone@example.com")

	def cached_boot(self):
		return frappe.cache().hget("bootinfo", "someone@example.com")

	def test_a_new_placement_reaches_desks_on_their_next_load(self):
		mapping = configure(self)
		self.seed_cached_boot()
		self.panel(mapping, files.PANEL_TAB)
		self.assertIsNone(self.cached_boot())
		# What the saving desk is handed straight away, and what the others rebuild.
		self.assertEqual(files.files_panel_boot()["microsoft365_files_panel"][MAPPED], files.PANEL_TAB)

		self.seed_cached_boot()
		mapping.delete()
		self.assertIsNone(self.cached_boot())
		self.assertNotIn(MAPPED, files.files_panel_boot()["microsoft365_files_doctypes"])
		self.assertIsNone(files.mapping_for(MAPPED))  # nor go on storing its files

	def test_switching_storage_on_or_off_reaches_desks_on_their_next_load(self):
		configure(self)
		settings = frappe.get_doc("Microsoft Settings")
		settings.flags.ignore_permissions = True

		self.seed_cached_boot()
		settings.save()  # storage still on: nothing for a desk to pick up
		self.assertIsNotNone(self.cached_boot())

		settings.use_files = 0
		settings.save()
		self.assertIsNone(self.cached_boot())
		self.assertEqual(files.files_panel_boot()["microsoft365_files_doctypes"], [])

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

	def test_file_gets_at_most_one_wide_varchar(self):
		# tabFile is shared with every app on the site and can already be near InnoDB's row
		# limit: a second wide varchar beside the indexed local_url failed migrate on a
		# COMPRESSED table. Anything long that is not indexed belongs in a Text field.
		from frappe_microsoft365.setup import create_file_custom_fields

		wide = [
			f["fieldname"]
			for f in create_file_custom_fields(only_definitions=True)["File"]
			if f["fieldtype"] == "Data" and (f.get("length") or 140) > 255
		]
		self.assertEqual(wide, ["custom_microsoft_local_url"])


	def test_file_list_shows_the_sharepoint_status(self):
		# Whether a file has moved should be visible in the File list, not only on the form.
		from frappe_microsoft365.setup import create_file_custom_fields

		create_file_custom_fields()
		frappe.clear_cache(doctype="File")
		status = frappe.get_meta("File").get_field("custom_microsoft_status")
		self.assertTrue(status.in_list_view)

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

	def test_site_grant_script_runs_when_pasted(self):
		# Pasted into a PowerShell window, a line ending inside @{ ... } leaves the shell at ">>"
		# waiting for more, and a last line without a line break is never run.
		script = doctor.powershell_for_site_grant(
			"APP-ID", ["https://c.sharepoint.com/sites/Sales", "https://c.sharepoint.com/teams/Ops"]
		)
		self.assertTrue(script.endswith("\n"))
		for line in script.splitlines():
			code = line.split("#", 1)[0].rstrip()
			self.assertFalse(code.endswith(("{", "(", "|", "`", ",", ";")), line)
			self.assertEqual(code.count("{"), code.count("}"), line)
			self.assertEqual(code.count("("), code.count(")"), line)

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
