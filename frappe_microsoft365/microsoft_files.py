"""SharePoint / Microsoft Teams document storage for Frappe attachments.

A site creates one SharePoint Mapping per DocType, saying where its files go:
a SharePoint site, a document library and a base folder (for a Teams channel, the channel's
folder in the team's library). From then on:

* every record of a mapped DocType gets its own folder, created the first time it is needed;
* a file attached in Frappe is moved into that folder by a background job, and its File row is
  repointed at ``open_file`` so it keeps opening from the form for anyone who may read the
  record — while the copy on this server's disk is removed, unless the site asks to keep it;
* files people drop into the folder from Teams or SharePoint are listed on the record, in the
  "Files in Microsoft 365" panel, and open through Frappe the same way.

Everything runs as the Azure app itself (client credentials), never as a signed-in user: a
shared library must not stop working because the person who connected it left. The app needs
the ``Sites.Selected`` application permission, and write access granted on each site; nothing
here asks for, or works with, tenant-wide ``Sites.ReadWrite.All``.

Graph contract used (all v1.0, see docs/graph-api-reference.md):

* ``GET /sites/{hostname}:/{server-relative-path}`` — site by URL
* ``GET /sites/{site-id}/drives`` — the site's document libraries
* ``GET /drives/{drive-id}/root:/{path}:`` and ``/items/{id}:/{name}:`` — item by path
* ``POST /drives/{drive-id}/items/{parent-id}/children`` — create folder
* ``PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content`` — simple upload
* ``POST .../{filename}:/createUploadSession`` then ``PUT {uploadUrl}`` in ranges — large upload
* ``GET /drives/{drive-id}/items/{item-id}/content`` — download (302 to a pre-signed URL)
* ``GET /drives/{drive-id}/items/{item-id}/children`` — folder listing
* ``PATCH /drives/{drive-id}/items/{item-id}`` with ``parentReference`` — move (re-attached file)
* ``DELETE /drives/{drive-id}/items/{item-id}`` — to the site recycle bin
"""

import ipaddress
import json
import mimetypes
import os
import re
import shutil
import socket
import tempfile
from urllib.parse import quote, urlparse

import frappe
from frappe import _
from frappe.utils import add_to_date, cint, now_datetime
from werkzeug.wrappers import Response

from frappe_microsoft365 import capabilities
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import APP, MsGraphConflict, MsGraphError, MsGraphNotFound

# --- constants ---------------------------------------------------------------------------

#: Below this a file goes up in one PUT; above it, through an upload session. Graph accepts a
#: single PUT up to 250 MB, but a single PUT also means the whole file in one request body and
#: one timeout for all of it; sessions resume and keep each request small.
SIMPLE_UPLOAD_MAX_BYTES = 4 * 1024 * 1024

#: Upload-session ranges must be multiples of 320 KiB (Microsoft's rule). 32 x 320 KiB = 10 MiB.
UPLOAD_CHUNK_BYTES = 32 * 320 * 1024

#: Streamed downloads are relayed in pieces this size, as microsoft_meeting_artifacts does.
DOWNLOAD_CHUNK_BYTES = 256 * 1024

#: A linked attachment is fetched before it is archived; nothing a record links to should be
#: anywhere near this, and a link to something bigger is more likely a mistake than a document.
MAX_LINK_DOWNLOAD_BYTES = 250 * 1024 * 1024

#: The longest folder name we generate. SharePoint allows 255 per segment and 400 for the whole
#: decoded path; a channel folder plus a record folder plus a file name has to fit in the latter.
MAX_FOLDER_NAME = 120

#: Background attempts per file before it is left as Failed for a person to look at.
MAX_ATTEMPTS = 5

#: Minutes a Pending/Failed file waits before the hourly sweep picks it up again.
RETRY_AFTER_MINUTES = 15

#: SharePoint's limit on a file's whole decoded path, counted from the site's address on
#: (``sites/Sales/Shared Documents/Projects/PRJ-0001/spec.pdf``).
MAX_PATH_LENGTH = 400

#: What a mapping's base folder must leave of MAX_PATH_LENGTH for a record's folder and the
#: files in it; a base path longer than that is refused when the mapping is saved.
MIN_ROOM_FOR_RECORDS = 100

#: What a record's folder must leave for the name of a file in it. A longer folder name is
#: shortened when the folder is created, and a file name that does not fit is shortened.
MIN_ROOM_FOR_FILE_NAME = 50

#: How far up the tree open_item walks to prove an item sits inside a record's folder. Not a
#: limit SharePoint has; its path limit allows at most this many levels (one character and a
#: slash each), so no file anyone can reach is ever refused for being deep.
MAX_FOLDER_DEPTH = MAX_PATH_LENGTH // 2

OPEN_METHOD = "frappe_microsoft365.microsoft_files.open_file"

# Status values written to File.custom_microsoft_status.
PENDING = "Pending"
STORED = "Stored"  # moved: the file now lives in SharePoint and file_url points at open_file
ARCHIVED = "Archived"  # a link attachment whose target was copied; the link itself is unchanged
FAILED = "Failed"

#: Characters SharePoint and OneDrive refuse in a name, plus # and % which it accepts but which
#: break more tools than they are worth.
INVALID_NAME_CHARS = re.compile(r'["*:<>?/\\|#%\x00-\x1f]')

#: Whole names SharePoint refuses outright.
RESERVED_NAMES = frozenset(
	[".lock", "con", "prn", "aux", "nul", "desktop.ini", "forms"]
	+ [f"com{i}" for i in range(10)]
	+ [f"lpt{i}" for i in range(10)]
)

PATTERN_TOKEN = re.compile(r"\{(\w+)\}")


# --- settings ----------------------------------------------------------------------------


def _settings():
	return graph._settings_doc()


#: Options a SharePoint Mapping may set for its own DocType: its field, and the SharePoint
#: Settings field that is the default.
MAPPING_OVERRIDES = {
	"files_keep_local_copy": "keep_local_copy",
	"files_delete_remote": "delete_remote",
}


def storage_option(fieldname, doctype=None):
	"""A SharePoint Settings option (files_keep_local_copy, files_archive_links, …).

	With a DocType, its mapping's override wins where it has one: On or Off on the mapping, or
	Default (or no mapping) to follow SharePoint Settings.
	"""
	override = MAPPING_OVERRIDES.get(fieldname)
	if override and doctype:
		# Disabled mappings count too: their files are still in SharePoint, and a mapping set to
		# keep them must keep doing so while paused.
		row = next((r for r in _mappings() if r.reference_doctype == doctype), None)
		choice = (row or {}).get(override)
		if choice == "On":
			return 1
		if choice == "Off":
			return 0
	return frappe.get_cached_doc("SharePoint Settings").get(fieldname)


def files_enabled(settings=None):
	return capabilities.is_enabled("files", settings or _settings())


def has_file_fields():
	"""Whether the File fields exist yet. They are created when storage is first switched on, so
	code that can run on a site that never switched it on must ask before touching them."""
	return frappe.get_meta("File").has_field("custom_microsoft_status")


MAPPING_CACHE_KEY = "frappe_microsoft365:sharepoint_mappings"

MAPPING_FIELDS = [
	"name",
	"enabled",
	"reference_doctype",
	"site_url",
	"library",
	"folder_creation",
	"use_base_folder",
	"base_folder",
	"folder_pattern",
	"site_id",
	"drive_id",
	"keep_local_copy",
	"delete_remote",
	"files_panel",
]


def _mappings():
	"""Every SharePoint Mapping, cached per site until one changes.

	Read on every attachment and, through the "*" hooks, on every delete and rename on the site,
	so it must not cost a query each time.
	"""
	cached = frappe.cache().get_value(MAPPING_CACHE_KEY)
	if cached is None:
		rows = [
			dict(r) for r in frappe.get_all("SharePoint Mapping", fields=MAPPING_FIELDS, order_by="name asc")
		]
		roles = {}
		if frappe.db.table_exists("SharePoint Mapping Role"):
			for r in frappe.get_all(
				"SharePoint Mapping Role",
				filters={"parenttype": "SharePoint Mapping"},
				fields=["parent", "role"],
			):
				roles.setdefault(r.parent, []).append(r.role)
		for row in rows:
			row["folder_roles"] = roles.get(row["name"], [])
		cached = rows
		frappe.cache().set_value(MAPPING_CACHE_KEY, cached)
	return [frappe._dict(r) for r in cached]


def can_create_folder(doctype, name, user=None):
	"""May this user create a SharePoint folder for this record?

	Needs edit rights on the record, and either System Manager or one of the roles the mapping
	lists under Who can create folders. Creating a folder moves the record's files out of the
	ERP, which is an administrator's decision unless the mapping says otherwise.
	"""
	user = user or frappe.session.user
	row = mapping_for(doctype)
	if not row or not frappe.has_permission(doctype, "write", doc=name, user=user):
		return False
	roles = set(frappe.get_roles(user))
	return "System Manager" in roles or bool(roles & set(row.get("folder_roles") or []))


def clear_mapping_cache():
	frappe.cache().delete_value(MAPPING_CACHE_KEY)


AUTOMATIC = "Automatic"
ON_DEMAND = "On demand"


def is_on_demand(row):
	"""On demand: a record gets a folder only when someone asks for one; until then its files stay put."""
	return (row.get("folder_creation") or AUTOMATIC) == ON_DEMAND


def mapping_for(doctype, settings=None):
	"""The enabled SharePoint Mapping for a DocType, or None."""
	if not doctype:
		return None
	for row in _mappings():
		if row.enabled and row.reference_doctype == doctype:
			return row
	return None


def mapped_doctypes(settings=None):
	settings = settings or _settings()
	if not files_enabled(settings):
		return []
	return sorted({row.reference_doctype for row in _mappings() if row.enabled})


@frappe.whitelist()
@frappe.validate_and_sanitize_search_inputs
def mapped_doctype_query(
	doctype: str, txt: str, searchfield: str, start: int, page_len: int, filters: dict | None = None
):
	"""Link-field search for DocTypes that have an enabled SharePoint Mapping."""
	txt = (txt or "").lower()
	names = [d for d in mapped_doctypes() if txt in d.lower()]
	start, page_len = cint(start), cint(page_len) or 20
	return [[d] for d in names[start : start + page_len]]


def boot_session(bootinfo):
	"""Tell the desk which DocTypes get the "Files in Microsoft 365" panel, and where."""
	try:
		bootinfo.microsoft365_files_doctypes = mapped_doctypes()
		bootinfo.microsoft365_files_panel = {
			doctype: panel_placement(mapping_for(doctype)) for doctype in bootinfo.microsoft365_files_doctypes
		}
	except Exception:
		# A broken settings row must never stop the desk from loading.
		bootinfo.microsoft365_files_doctypes = []
		bootinfo.microsoft365_files_panel = {}


def refresh_desk_boot():
	"""Make every desk pick up a change to which DocTypes have the panel, or where it goes.

	Frappe keeps each user's boot in the cache until they reload or sign in again, so without
	this a changed placement reaches nobody: "Own tab" adds the tab at once while the panel stays
	at the top of the form. The next page load rebuilds the boot.
	"""
	frappe.cache().delete_value("bootinfo")


@frappe.whitelist()
def files_panel_boot():
	"""The panel's boot values, for a desk that should not wait for a reload (see refresh_desk_boot)."""
	bootinfo = frappe._dict()
	boot_session(bootinfo)
	return {key: bootinfo[key] for key in ("microsoft365_files_doctypes", "microsoft365_files_panel")}


PANEL_TOP = "Top of form"
PANEL_TAB = "Own tab"
PANEL_HIDDEN = "Hidden"

#: The tab a mapping with Files Panel = Own tab adds to its DocType's form, and the field the panel
#: draws into. Neither has a column.
FILES_TAB_FIELD = "custom_microsoft365_files_tab"
FILES_HTML_FIELD = "custom_microsoft365_files"
FILES_TAB_FIELDS = (FILES_TAB_FIELD, FILES_HTML_FIELD)


def panel_placement(row):
	return (row or {}).get("files_panel") or PANEL_TOP


def sync_files_tab(doctype, wanted):
	"""Add or remove the Microsoft 365 Files tab on a DocType's form (idempotent).

	A tab is form layout, so it is a Custom Field like any other change to someone else's form;
	that is also what lets a System Manager move it with Customize Form.
	"""
	names = frappe.get_all(
		"Custom Field", filters={"dt": doctype, "fieldname": ["in", FILES_TAB_FIELDS]}, pluck="name"
	)
	if len(names) == (len(FILES_TAB_FIELDS) if wanted else 0):
		return
	for name in names:
		frappe.delete_doc("Custom Field", name, ignore_permissions=True, force=True)
	if wanted:
		last = [df.fieldname for df in frappe.get_meta(doctype).fields if df.fieldname not in FILES_TAB_FIELDS]
		fields = [
			{
				"fieldname": FILES_TAB_FIELD,
				"fieldtype": "Tab Break",
				"label": "Microsoft 365 Files",
				"insert_after": last[-1] if last else None,
				# Only on saved records of a DocType that is still mapped with storage on: the
				# boot list is empty when either is switched off.
				"depends_on": "eval:!doc.__islocal && (frappe.boot.microsoft365_files_doctypes || []).includes(doc.doctype)",
			},
			{"fieldname": FILES_HTML_FIELD, "fieldtype": "HTML", "insert_after": FILES_TAB_FIELD},
		]
		# Inserted directly rather than through create_custom_fields: that ends in a schema sync,
		# which commits, and neither field has a column to sync.
		in_create = frappe.flags.in_create_custom_fields
		frappe.flags.in_create_custom_fields = True
		try:
			for df in fields:
				field = frappe.get_doc({"doctype": "Custom Field", "dt": doctype, **df})
				field.flags.ignore_validate = True
				field.insert(ignore_permissions=True)
		finally:
			frappe.flags.in_create_custom_fields = in_create
	frappe.clear_cache(doctype=doctype)


# --- names and paths ---------------------------------------------------------------------


def safe_name(value, fallback="Untitled", max_length=MAX_FOLDER_NAME):
	"""A string SharePoint accepts as a file or folder name.

	Pure. ``PO/2026/0042-A`` -> ``PO-2026-0042-A``:
	slashes are by far the commonest offender, because naming series love them.
	"""
	name = INVALID_NAME_CHARS.sub("-", str(value if value is not None else ""))
	name = re.sub(r"\s+", " ", name)
	name = re.sub(r"-{2,}", "-", name)
	name = name.strip(" .-")
	if name.startswith("~$"):
		name = name[2:].strip(" .-")
	if name.lower() in RESERVED_NAMES or "_vti_" in name.lower():
		name = f"_{name}"
	if len(name) > max_length:
		name = name[:max_length].rstrip(" .-")
	return name or fallback


def safe_file_name(value):
	"""safe_name for a file: truncation keeps the extension, which is what opens it."""
	stem, ext = os.path.splitext(str(value or ""))
	ext = INVALID_NAME_CHARS.sub("", ext)[:16]
	return safe_name(stem, fallback="file", max_length=200 - len(ext)) + ext


def render_pattern(pattern, doc):
	"""The folder name for a record. ``{name}``, ``{title}`` or any ``{fieldname}``.

	Pure apart from get_title. A token that names no field renders empty rather than failing:
	a typo in settings should give a slightly odd folder name, not stop every upload.
	"""
	pattern = (pattern or "").strip() or "{name}"

	def value(match):
		token = match.group(1)
		if token == "name":
			return str(doc.name or "")
		if token == "title":
			title = doc.get_title() if hasattr(doc, "get_title") else doc.get("title")
			return str(title or doc.name or "")
		val = doc.get(token)
		return "" if val is None else str(val)

	return safe_name(PATTERN_TOKEN.sub(value, pattern), fallback=safe_name(doc.name))


def path_segments(path):
	"""Split a configured folder path into safe segments. ``"Projects / 2026"`` -> two."""
	return [safe_name(part) for part in re.split(r"[/\\]", path or "") if part.strip()]


def base_segments(row):
	"""The base-folder path of a mapping, or none when record folders go straight into the library."""
	if not cint(row.get("use_base_folder", 1)):
		return []
	return path_segments(row.base_folder)


def encode_path(segments):
	return "/".join(quote(segment, safe="") for segment in segments)


def library_path(row):
	"""The decoded path of a mapping's library, from the site's address on. Pure.

	Counted the way SharePoint counts its path limit. The library a Teams team keeps its files in
	is called Documents but lives at Shared Documents; any other library is assumed to live at its
	name, which is what SharePoint uses unless the library was renamed after it was made.
	"""
	_host, site_path = parse_site_url(row.get("site_url"))
	library = (row.get("library") or "").strip()
	if library.lower() in ("", "documents", "shared documents"):
		library = "Shared Documents"
	return "/".join(part for part in (site_path.strip("/"), library) if part)


def base_path_length(row):
	"""How much of SharePoint's path limit a mapping's library and base folder use. Pure."""
	return len("/".join([library_path(row), *base_segments(row)]))


def assert_base_path_fits(row):
	"""Refuse a base folder that leaves no room for record folders and their files."""
	used = base_path_length(row)
	if used > MAX_PATH_LENGTH - MIN_ROOM_FOR_RECORDS:
		frappe.throw(
			_(
				"The library and base folder path is {0} characters long, counted from the site address. SharePoint allows {1} characters for a file's whole path, so this leaves too little room for each record's folder and the files in it (at least {2} are needed). Use a shorter base folder."
			).format(used, MAX_PATH_LENGTH, MIN_ROOM_FOR_RECORDS),
			title=_("SharePoint path too long"),
		)


def folder_path_length(folder, row=None):
	"""How much of SharePoint's path limit a record folder's own path uses.

	Read from the folder's address when there is one (the real path, whatever the folder was
	called or moved to since), else from the mapping and the recorded path.
	"""
	from urllib.parse import unquote

	if folder.get("web_url"):
		return len(unquote(urlparse(folder.web_url).path).strip("/"))
	row = row or mapping_for(folder.get("reference_doctype")) or {}
	return len("/".join(part for part in (library_path(row), folder.get("folder_path") or "") if part))


def fit_file_name(name, room):
	"""``name`` shortened, keeping its extension, to ``room`` characters; None if it cannot be. Pure."""
	if len(name) <= room:
		return name
	stem, ext = os.path.splitext(name)
	keep = room - len(ext)
	if keep < 8:
		return None
	return stem[:keep].rstrip(" .-") + ext


def parse_site_url(url):
	"""``(hostname, server-relative path)`` for the site a URL belongs to.

	People paste whatever is in the address bar — a library, a folder, a file preview — so
	anything past ``/sites/<name>`` or ``/teams/<name>`` is dropped. Anything else is the
	tenant's root site, which has an empty path.
	"""
	raw = (url or "").strip()
	if not raw:
		frappe.throw(_("SharePoint Site URL is required."))
	if "://" not in raw:
		raw = f"https://{raw}"
	parsed = urlparse(raw)
	host = (parsed.hostname or "").lower()
	if not host:
		frappe.throw(_("{0} is not a SharePoint site URL.").format(url))
	match = re.match(r"^/(sites|teams)/([^/]+)", parsed.path or "", re.IGNORECASE)
	return host, f"/{match.group(1).lower()}/{match.group(2)}" if match else ""


def _forget_message():
	"""Drop the message an expected Graph error queued for the browser.

	graph_request raises through frappe.throw, which also queues the text for the desk to show.
	When the caller expects the error — "does this folder exist yet?" — and handles it, that
	queued text would still pop up as a red "itemNotFound" box over a perfectly good result.
	"""
	frappe.clear_last_message()


# --- locating the library ----------------------------------------------------------------


def resolve_drive(row):
	"""Return ``(site_id, drive_id)`` for a mapping row, looking them up once and caching.

	Microsoft ids are stable where URLs and display names are not, so everything downstream
	works with ids and only this function ever reads the URL.
	"""
	if row.get("site_id") and row.get("drive_id"):
		return row.site_id, row.drive_id

	host, path = parse_site_url(row.site_url)
	site = graph.graph_request(
		"GET",
		f"/sites/{host}:{path}" if path else f"/sites/{host}",
		APP,
		params={"$select": "id,webUrl,displayName"},
	)
	site_id = site["id"]
	drive_id = _find_drive(site_id, row.library)

	frappe.db.set_value(
		"SharePoint Mapping", row.name, {"site_id": site_id, "drive_id": drive_id}, update_modified=False
	)
	row.site_id, row.drive_id = site_id, drive_id
	clear_mapping_cache()
	return site_id, drive_id


def _drive_gone(drive_id):
	"""Ask for the library itself, by id: True when SharePoint says there is no such drive."""
	try:
		graph.graph_request("GET", f"/drives/{drive_id}", APP, params={"$select": "id"})
	except MsGraphNotFound:
		_forget_message()
		return True
	return False


def with_current_drive(row, work):
	"""``work(drive_id)`` against the mapping's library, looking the library up again once if the
	cached id has stopped existing.

	The site and drive ids are cached on the mapping (resolve_drive). A library or site that was
	deleted and created again under the same address gets new ids, and the cached drive id then
	answers 404 for everything, so every new folder fails until someone re-saves the mapping.
	On a 404 the drive itself is asked for; only if it is gone are the ids looked up again from
	the URL, and the work is tried once more against the new drive.
	"""
	cached = bool(row.get("site_id") and row.get("drive_id"))
	_site_id, drive_id = resolve_drive(row)
	try:
		return work(drive_id)
	except MsGraphNotFound:
		if not cached or not _drive_gone(drive_id):
			raise
		_forget_message()
	row.site_id = row.drive_id = None
	_site_id, fresh = resolve_drive(row)
	return work(fresh)


def _find_drive(site_id, library):
	library = (library or "").strip()
	drives = graph.graph_paged(f"/sites/{site_id}/drives?$select=id,name,webUrl", APP)
	wanted = library.lower()
	for drive in drives:
		if (drive.get("name") or "").lower() == wanted:
			return drive["id"]
		# The library a Teams team stores channel files in is called "Documents" but lives at
		# /Shared Documents, and people type whichever they saw last.
		web_tail = (drive.get("webUrl") or "").rstrip("/").rsplit("/", 1)[-1]
		if web_tail and quote(library) == web_tail:
			return drive["id"]
	if wanted in ("", "documents", "shared documents"):
		return graph.graph_request("GET", f"/sites/{site_id}/drive", APP, params={"$select": "id"})["id"]
	names = ", ".join(sorted(d.get("name") or "?" for d in drives)) or _("none visible")
	frappe.throw(
		_("No document library named {0} on that site. Libraries the app can see: {1}").format(
			library, names
		),
		MsGraphError,
	)


def _get_by_path(drive_id, segments, parent_id=None):
	"""The item at a path, or None. Relative to the library root, or to parent_id."""
	base = f"/drives/{drive_id}/items/{parent_id}" if parent_id else f"/drives/{drive_id}/root"
	path = f"{base}:/{encode_path(segments)}:" if segments else base
	try:
		return graph.graph_request(
			"GET", path, APP, params={"$select": "id,name,folder,webUrl,parentReference"}
		)
	except MsGraphNotFound:
		_forget_message()
		return None


def _create_folder(drive_id, parent_id, name, conflict="fail"):
	"""Create a folder.

	With ``conflict="fail"`` (base folders, the connection probe) a folder that is already there
	under that name is returned instead: a base folder is shared by every record, and a concurrent
	job may have created it a moment earlier. With ``"rename"`` SharePoint always creates a new
	folder, naming it ``Name 1`` when ``Name`` is taken; record folders are made that way.
	"""
	parent = f"/drives/{drive_id}/items/{parent_id}" if parent_id else f"/drives/{drive_id}/root"
	body = {"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": conflict}
	if conflict != "fail":
		return graph.graph_request("POST", f"{parent}/children", APP, json=body)
	try:
		return graph.graph_request("POST", f"{parent}/children", APP, json=body)
	except MsGraphConflict:
		_forget_message()
		existing = _get_by_path(drive_id, [name], parent_id=parent_id)
		if existing:
			return existing
		raise


def ensure_path(drive_id, segments):
	"""The folder at ``segments`` under the library root, created level by level if missing."""
	found = _get_by_path(drive_id, segments)
	if found:
		if "folder" not in found:
			frappe.throw(
				_("{0} exists in SharePoint but is a file, not a folder.").format("/".join(segments)),
				MsGraphError,
			)
		return found

	parent_id = None
	item = None
	for segment in segments:
		item = _get_by_path(drive_id, [segment], parent_id=parent_id) or _create_folder(
			drive_id, parent_id, segment
		)
		parent_id = item["id"]
	return item


# --- record folders ----------------------------------------------------------------------


def get_folder(doctype, name):
	"""The SharePoint Folder record for a document, or None."""
	return frappe.db.get_value(
		"SharePoint Folder",
		{"reference_doctype": doctype, "reference_name": name},
		["name", "reference_doctype", "drive_id", "item_id", "web_url", "folder_name", "folder_path"],
		as_dict=True,
	)


def ensure_folder(doctype, name):
	"""The folder for a record, created in SharePoint (and recorded here) on first use.

	Commits when it creates one. Each call that creates makes a new folder (see
	new_folder_values), so the record of it has to outlive whatever the caller does next: rolled
	back with a failed upload, the retry would make a second folder beside the first.
	"""
	existing = get_folder(doctype, name)
	if existing:
		return existing

	row = mapping_for(doctype)
	if not row:
		frappe.throw(_("{0} is not mapped to SharePoint in Microsoft Settings.").format(_(doctype)))

	from frappe.utils.synchronization import filelock

	# Several attachments of one record are uploaded by parallel jobs. One at a time here, or
	# each would create its own folder.
	with filelock(_folder_lock_name(doctype, name), timeout=60):
		# A fresh transaction, so that a folder another job committed while this one waited is seen.
		frappe.db.commit()  # nosemgrep
		existing = get_folder(doctype, name)
		if existing:
			return existing
		record = add_folder_row(doctype, name, new_folder_values(row, doctype, name))
		# The folder exists in SharePoint now; see the docstring.
		frappe.db.commit()  # nosemgrep
	return frappe._dict(record.as_dict())


def _folder_lock_name(doctype, name):
	"""A lock file name for one record. Hashed: record names may hold slashes and other characters."""
	import hashlib

	return "microsoft365-folder-" + hashlib.md5(f"{doctype}\0{name}".encode()).hexdigest()


def add_folder_row(doctype, name, values):
	"""Insert a row in the mapping's Folders table without saving the mapping.

	Saving the mapping would rewrite every row and fail whenever two records get their folders at
	once. Touching its modified time instead makes a mapping form opened before this row existed
	refuse to save ("modified after you opened it") rather than save without it.
	"""
	idx = (
		frappe.db.sql(
			"select max(idx) from `tabSharePoint Folder` where parent=%s and parenttype='SharePoint Mapping'",
			doctype,
		)[0][0]
		or 0
	)
	assert_folder_unclaimed(values.get("item_id"), doctype, name)
	record = frappe.get_doc(
		{
			"doctype": "SharePoint Folder",
			"parent": doctype,
			"parenttype": "SharePoint Mapping",
			"parentfield": "folders",
			"idx": idx + 1,
			"reference_doctype": doctype,
			"reference_name": name,
			**values,
		}
	)
	record.db_insert()
	touch_mapping(doctype)
	return record


def touch_mapping(doctype):
	"""Mark the mapping changed after its Folders rows were changed directly, so a mapping form
	opened before cannot save its stale rows back over them."""
	frappe.db.set_value("SharePoint Mapping", doctype, "modified", now_datetime(), update_modified=False)


def assert_folder_unclaimed(item_id, doctype, name):
	"""Refuse a folder that is already another record's folder.

	Whoever can read a record sees and opens everything in its folder, so two records on one
	folder would each show the other's files.
	"""
	if not item_id:
		return
	for other in frappe.get_all(
		"SharePoint Folder",
		filters={"item_id": item_id},
		fields=["reference_doctype", "reference_name"],
	):
		if (other.reference_doctype, other.reference_name) != (doctype, name):
			frappe.throw(
				_(
					"That SharePoint folder is already the folder of {0} {1}. Each record needs a folder of its own."
				).format(_(other.reference_doctype), other.reference_name),
				frappe.DuplicateEntryError,
			)


def new_folder_values(row, doctype, name):
	"""Create the record's own folder where its mapping says, and return the fields to store.

	Always a new folder. One already there under the same name belongs to someone else: another
	record with the same title, or a name that is the same once cleaned for SharePoint, or a
	deleted record whose folder was kept. So SharePoint is asked to pick a free name (``PRJ-0001
	1``) rather than hand that folder over. Using a folder that already exists is something an
	administrator does on purpose, with Existing Folder on the mapping's Folders table.
	"""
	doc = frappe.get_doc(doctype, name)
	assert_base_path_fits(row)
	base = base_segments(row)
	# The record's folder must leave room for the names of the files that go in it.
	room = MAX_PATH_LENGTH - base_path_length(row) - 1 - MIN_ROOM_FOR_FILE_NAME
	folder_name = safe_name(render_pattern(row.folder_pattern, doc), max_length=min(MAX_FOLDER_NAME, room))

	def create(drive_id):
		parent_id = ensure_path(drive_id, base)["id"] if base else None
		return drive_id, _create_folder(drive_id, parent_id, folder_name, conflict="rename")

	drive_id, item = with_current_drive(row, create)
	if "folder" not in item:
		frappe.throw(_("SharePoint did not create the folder {0}.").format(folder_name), MsGraphError)
	folder_name = item.get("name") or folder_name
	segments = [*base, folder_name]
	return {
		"folder_name": folder_name,
		"drive_id": drive_id,
		"item_id": item["id"],
		"web_url": item.get("webUrl"),
		"folder_path": "/".join(segments),
	}


def existing_folder_values(row, location):
	"""Find a folder someone points at — a SharePoint link, or a path inside the mapped library.

	Accepts what people actually copy: the folder's address from the browser (including the
	``AllItems.aspx?id=/sites/…`` form), a Copy link / Share link, or a plain path such as
	``Projects/PROJ-0001``. Only a folder is accepted, and only one the app can reach.
	"""
	import base64
	from urllib.parse import parse_qs, unquote

	value = (location or "").strip()

	def find(drive_id):
		item = None
		if value.lower().startswith(("http://", "https://")):
			parsed = urlparse(value)
			ids = parse_qs(parsed.query).get("id")
			path = unquote(ids[0] if ids else parsed.path).rstrip("/")
			drive = graph.graph_request("GET", f"/drives/{drive_id}", APP, params={"$select": "webUrl"})
			root = unquote(urlparse(drive.get("webUrl") or "").path).rstrip("/")
			if root and path.lower().startswith(root.lower() + "/"):
				item = _get_by_path(drive_id, [p for p in path[len(root) + 1 :].split("/") if p])
			else:
				# A sharing link (…/:f:/s/…) names no path; Graph resolves those through /shares.
				token = "u!" + base64.urlsafe_b64encode(value.encode()).decode().rstrip("=")
				try:
					item = graph.graph_request(
						"GET",
						f"/shares/{token}/driveItem",
						APP,
						params={"$select": "id,name,folder,webUrl,parentReference"},
					)
				except MsGraphError:
					_forget_message()
					item = None
		else:
			item = _get_by_path(drive_id, [p for p in value.replace("\\", "/").split("/") if p.strip()])

		if not item:
			# Raised bare, not thrown: nothing for the desk to show unless the retry fails too.
			raise MsGraphNotFound("folder not found")
		return drive_id, item

	try:
		drive_id, item = with_current_drive(row, find)
	except MsGraphNotFound:
		_forget_message()
		item = None
	if not item:
		frappe.throw(
			_(
				"Could not find that folder. Paste its address from SharePoint, or its path inside the {0} library, e.g. Projects/PRJ-0001. It must be on a site the app has been granted."
			).format(row.library or "Documents")
		)
	if "folder" not in item:
		frappe.throw(_("{0} is a file, not a folder.").format(item.get("name")))
	return {
		"folder_name": item.get("name"),
		"drive_id": (item.get("parentReference") or {}).get("driveId") or drive_id,
		"item_id": item["id"],
		"web_url": item.get("webUrl"),
		"folder_path": value,
	}


def queue_record(doctype, name):
	"""Queue the attachments of one record that are still on this server. Returns how many."""
	settings = _settings()
	queued = 0
	for row in frappe.get_all(
		"File",
		filters={
			"attached_to_doctype": doctype,
			"attached_to_name": name,
			"is_folder": 0,
			"custom_microsoft_status": ["in", ["", None, FAILED]],
		},
		fields=[
			"name",
			"file_url",
			"attached_to_doctype",
			"attached_to_name",
			"attached_to_field",
			"is_folder",
		],
	):
		if not _eligible(frappe._dict(row), settings):
			continue
		frappe.db.set_value("File", row.name, "custom_microsoft_status", PENDING, update_modified=False)
		enqueue_upload(row.name)
		queued += 1
	return queued


def _forget_folder(folder):
	"""Drop a folder record whose SharePoint folder is gone, so the next call recreates it."""
	frappe.db.delete("SharePoint Folder", {"name": folder.name})
	if folder.get("reference_doctype"):
		touch_mapping(folder.reference_doctype)


# --- uploads -----------------------------------------------------------------------------


def stored_url(file_name):
	"""The file_url a moved File carries. Relative, so it survives a domain change."""
	return f"/api/method/{OPEN_METHOD}?file={quote(file_name, safe='')}"


def local_path(url):
	"""Disk path for a /files or /private/files URL."""
	from frappe.utils import get_files_path

	if url.startswith("/private/files/"):
		return get_files_path(url.split("/private/files/", 1)[1], is_private=1)
	return get_files_path(url.split("/files/", 1)[1])


def is_local(file_doc):
	url = file_doc.file_url or ""
	return url.startswith(("/files/", "/private/files/"))


def is_link(file_doc):
	return (file_doc.file_url or "").startswith(("http://", "https://"))


def _eligible(file_doc, settings=None):
	"""Should this File be stored in SharePoint at all? Pure apart from settings."""
	settings = settings or _settings()
	if not files_enabled(settings) or file_doc.is_folder:
		return False
	if not (file_doc.attached_to_doctype and file_doc.attached_to_name):
		return False
	# A file behind a field (an image, a signature, an Attach field) is part of the record's
	# rendering — print formats and web pages fetch it without a desk session — so only loose
	# attachments move.
	if file_doc.attached_to_field:
		return False
	row = mapping_for(file_doc.attached_to_doctype, settings)
	if not row:
		return False
	if is_on_demand(row) and not get_folder(file_doc.attached_to_doctype, file_doc.attached_to_name):
		# Nobody has asked for this record to live in SharePoint yet.
		return False
	if is_local(file_doc):
		return True
	return bool(storage_option("files_archive_links")) and is_link(file_doc)


#: The File fields that say where an attachment lives in SharePoint. Only the app writes them,
#: always with frappe.db.set_value / db_set, which do not run validate; on_file_validate refuses
#: them from anyone else.
PROTECTED_FILE_FIELDS = (
	"custom_microsoft_status",
	"custom_microsoft_web_url",
	"custom_microsoft_error",
	"custom_microsoft_drive_id",
	"custom_microsoft_item_id",
	"custom_microsoft_folder_id",
	"custom_microsoft_local_url",
	"custom_microsoft_attempts",
	"custom_microsoft_content_hash",
)


def is_open_url(url):
	"""Whether a file_url hands the file to open_file. Pure."""
	from urllib.parse import unquote

	return OPEN_METHOD in unquote(url or "").lower()


def _same(a, b):
	return str(a or "") == str(b or "")


def on_file_validate(doc, method=None):
	"""File.validate: only this app says where a File lives in SharePoint.

	open_file, File.get_content and on_file_trash act on the drive and item a File names, as the
	app, so a File anyone could point at any item would let them read or delete anything the app
	can reach. Frappe does not enforce read_only over the API, and every user can create Files.
	Refused here, on insert and on update: setting or changing any PROTECTED_FILE_FIELDS, and a
	file_url that is, or was, the open_file address.

	The app itself never comes through here: it writes these fields with frappe.db.set_value and
	db_set, which skip validate. The one change allowed is a copy made by
	File.create_attachment_copy (or Duplicate): a new File that repeats, field for field, a File
	the user may read, as copies of a moved attachment have to.
	"""
	fields = PROTECTED_FILE_FIELDS if has_file_fields() else ()
	before = None
	if not doc.is_new():
		before = doc.get_doc_before_save() or frappe.db.get_value(
			"File", doc.name, ["file_url", *fields], as_dict=True
		)
	before = before or {}

	changed = [f for f in fields if not _same(doc.get(f), before.get(f))]
	url_changed = not _same(doc.file_url, before.get("file_url"))
	if url_changed and (is_open_url(doc.file_url) or is_open_url(before.get("file_url"))):
		changed.append("file_url")
	if not changed:
		return
	if doc.is_new() and _copies_a_readable_file(doc, fields):
		return
	frappe.throw(
		_("Where a file is stored in SharePoint is set by the Microsoft 365 app only ({0}).").format(
			", ".join(changed)
		),
		frappe.PermissionError,
	)


def _copies_a_readable_file(doc, fields):
	"""Is this new File a copy of an existing one the user may read? See on_file_validate."""
	if not doc.file_url:
		return False
	for source in frappe.get_all(
		"File",
		filters={"file_url": doc.file_url, "name": ["!=", doc.name or ""]},
		fields=["name", *fields],
	):
		if not all(_same(doc.get(f), source.get(f)) for f in fields):
			continue
		if doc.flags.ignore_permissions or frappe.has_permission("File", "read", doc=source.name):
			return True
	return False


def on_file_insert(doc, method=None):
	"""File.after_insert: queue the move. Never fails the upload itself."""
	try:
		if not _eligible(doc):
			return
		# A copy (File.create_attachment_copy) arrives with the original's SharePoint fields. A new
		# File has never been uploaded itself, so ids it carries are not its own to resume from.
		doc.db_set(
			{
				"custom_microsoft_status": PENDING,
				"custom_microsoft_error": None,
				"custom_microsoft_attempts": 0,
				"custom_microsoft_drive_id": None,
				"custom_microsoft_item_id": None,
			},
			update_modified=False,
		)
		enqueue_upload(doc.name)
	except Exception:
		frappe.log_error(title=_("Microsoft 365: could not queue file for SharePoint"))


def enqueue_upload(file_name):
	frappe.enqueue(
		"frappe_microsoft365.microsoft_files.upload_file",
		queue="long",
		file_name=file_name,
		job_id=f"microsoft365-file-{file_name}",
		deduplicate=True,
		enqueue_after_commit=True,
	)


def upload_file(file_name):
	"""Background job: put one File into its record's folder. Idempotent; safe to re-run."""
	if not frappe.db.exists("File", file_name):
		return
	file_doc = frappe.get_doc("File", file_name)
	if file_doc.get("custom_microsoft_status") in (STORED, ARCHIVED) or not _eligible(file_doc):
		return

	attempts = cint(file_doc.get("custom_microsoft_attempts")) + 1
	# Counted before anything can fail, so that whatever goes wrong, here or later in a way
	# nobody foresaw, still uses up one of the MAX_ATTEMPTS and the hourly retry stops.
	frappe.db.set_value("File", file_name, "custom_microsoft_attempts", attempts, update_modified=False)
	frappe.db.commit()  # nosemgrep

	# ensure_folder commits the folder it creates (see there), so it runs before the savepoint.
	try:
		folder = ensure_folder(file_doc.attached_to_doctype, file_doc.attached_to_name)
	except Exception as e:
		frappe.db.rollback()
		_record_failure(file_name, attempts, e)
		return

	# A savepoint, not a full rollback, on failure: the caller's own unrelated writes must not be
	# thrown away with the upload's.
	frappe.db.savepoint("microsoft365_upload")
	try:
		item = _already_uploaded(file_doc, folder)
		if not item and attempts > 1:
			item = _left_by_an_earlier_attempt(file_doc, folder)
		if not item:
			try:
				item = _upload(file_doc, folder)
			except MsGraphNotFound:
				_forget_message()
				# Someone deleted or moved the record's folder out of the library. Recreate it once.
				_forget_folder(folder)
				folder = ensure_folder(file_doc.attached_to_doctype, file_doc.attached_to_name)
				item = _upload(file_doc, folder)
			# The bytes are in SharePoint now. Say so durably before anything else can fail, or
			# a retry uploads them again and SharePoint keeps both, as "name (1).pdf".
			frappe.db.set_value(
				"File",
				file_name,
				{
					"custom_microsoft_drive_id": _drive_of(item, folder),
					"custom_microsoft_item_id": item["id"],
					"custom_microsoft_folder_id": folder.item_id,
				},
				update_modified=False,
			)
			frappe.db.commit()  # nosemgrep
	except Exception as e:
		_rollback_to("microsoft365_upload")
		_record_failure(file_name, attempts, e)
		return

	local_url = file_doc.file_url if is_local(file_doc) else None
	keep_local = bool(storage_option("files_keep_local_copy", file_doc.attached_to_doctype))
	values = {
		"custom_microsoft_status": STORED if local_url else ARCHIVED,
		"custom_microsoft_drive_id": _drive_of(item, folder),
		"custom_microsoft_item_id": item["id"],
		"custom_microsoft_folder_id": folder.item_id,
		"custom_microsoft_web_url": item.get("webUrl"),
		"custom_microsoft_error": None,
		"custom_microsoft_attempts": attempts,
	}
	if local_url:
		values["file_url"] = stored_url(file_name)
		values["content_hash"] = None
		# What the bytes were, so that whatever is at local_url later is served or brought back
		# only if it is still this file (see own_local_copy).
		try:
			values["custom_microsoft_content_hash"] = _hash_of(file_doc.get_full_path())
		except OSError:
			# Uploaded a moment ago, so this is unlikely; without a hash the kept copy is simply
			# trusted less (own_local_copy), which is no reason to fail the move.
			values["custom_microsoft_content_hash"] = None
		# Kept whether or not the disk copy is: links already handed out (a bookmark, an email,
		# a form opened before the move) still say this URL, and redirect_moved_file follows it.
		values["custom_microsoft_local_url"] = local_url
	frappe.db.savepoint("microsoft365_record")
	try:
		frappe.db.set_value("File", file_name, values, update_modified=False)
		# The row must point at SharePoint, durably, before the copy on disk is removed.
		frappe.db.commit()  # nosemgrep
	except Exception as e:
		# The upload is recorded already: the next attempt picks it up instead of sending it again.
		frappe.db.rollback(save_point="microsoft365_record")
		_record_failure(file_name, attempts, e)
		return

	if local_url:
		_repoint_links(file_doc, local_url, values["file_url"])
	if local_url and not keep_local:
		_remove_local_copy(local_url)


def _rollback_to(savepoint):
	"""Roll back to a savepoint, or the whole transaction if a commit since has released it.

	ensure_folder commits when it has to recreate a folder mid-upload, and MariaDB forgets
	savepoints on commit.
	"""
	try:
		frappe.db.rollback(save_point=savepoint)
	except Exception:
		frappe.db.rollback()


def _hash_of(path):
	"""MD5 hex digest of a file on disk, read in pieces. Frappe's content_hash is MD5 too."""
	import hashlib

	digest = hashlib.md5()
	# Callers pass local_path() of a /files URL.
	with open(path, "rb") as fh:  # nosemgrep
		for chunk in iter(lambda: fh.read(DOWNLOAD_CHUNK_BYTES), b""):
			digest.update(chunk)
	return digest.hexdigest()


def _drive_of(item, folder):
	return item.get("parentReference", {}).get("driveId") or folder.drive_id


def _already_uploaded(file_doc, folder):
	"""The driveItem an earlier attempt uploaded before it failed, if it is still in the folder."""
	item_id = file_doc.get("custom_microsoft_item_id")
	drive_id = file_doc.get("custom_microsoft_drive_id")
	if not (item_id and drive_id):
		return None
	try:
		item = graph.graph_request(
			"GET",
			f"/drives/{drive_id}/items/{item_id}",
			APP,
			params={"$select": "id,webUrl,parentReference"},
		)
	except MsGraphNotFound:
		_forget_message()
		# Removed from SharePoint since: upload it again.
		return None
	parent = item.get("parentReference") or {}
	if parent.get("id") != folder.item_id or (parent.get("driveId") or drive_id) != folder.drive_id:
		# Moved out of the record's folder since, or the folder was recreated: not this file's copy.
		return None
	return item


def _left_by_an_earlier_attempt(file_doc, folder):
	"""The item an earlier attempt uploaded but did not live to record, if it is in the folder.

	A worker killed between the upload and the commit of its ids leaves the bytes in SharePoint
	and nothing here saying so; the retry would upload them again and SharePoint keep both, as
	``report (1).pdf``. So a retry first looks in the record's folder for a file under the same
	name with the same size, uploaded by this app rather than by a person, that no other File
	claims, and adopts it.
	"""
	if not is_local(file_doc):
		return None  # a link's size is only known once it is downloaded again
	try:
		size = os.path.getsize(file_doc.get_full_path())
	except OSError:
		return None
	name = upload_name(file_doc, folder)
	try:
		item = graph.graph_request(
			"GET",
			f"/drives/{folder.drive_id}/items/{folder.item_id}:/{quote(name, safe='')}:",
			APP,
			params={"$select": "id,name,size,file,webUrl,parentReference,createdBy"},
		)
	except MsGraphNotFound:
		_forget_message()
		return None
	if "file" not in item or cint(item.get("size")) != size or not _uploaded_by_this_app(item):
		return None
	if frappe.db.exists("File", {"custom_microsoft_item_id": item["id"], "name": ["!=", file_doc.name]}):
		return None
	return item


def _uploaded_by_this_app(item):
	"""Whether a driveItem was created by this app, as far as SharePoint says: by its client id
	when SharePoint names the application, else by naming no signed-in user."""
	created = item.get("createdBy") or {}
	app_id = (created.get("application") or {}).get("id")
	if app_id:
		settings = graph.get_settings()
		return app_id in {(settings.get("app_client_id") or "").strip(), (settings.get("client_id") or "").strip()} - {""}
	return not (created.get("user") or {}).get("id")


def _record_failure(file_name, attempts, error):
	frappe.db.set_value(
		"File",
		file_name,
		{
			"custom_microsoft_status": FAILED,
			"custom_microsoft_error": str(error)[:1000],
			"custom_microsoft_attempts": attempts,
		},
		update_modified=False,
	)
	# The failure has to outlive the job that hit it, or the hourly retry never sees it.
	frappe.db.commit()  # nosemgrep
	if attempts >= MAX_ATTEMPTS:
		frappe.log_error(
			title=_("Microsoft 365: file not stored in SharePoint"), message=f"File {file_name}: {error}"
		)


MOVED_EVENT = "microsoft365_file_moved"


def redirect_moved_file():
	"""before_request: send an old /files or /private/files link of a moved file to its new home.

	Decided by the database, not the disk: with Keep Local Copy on the old path still exists, but
	no File carries it as its address any more, so Frappe would answer it with Forbidden. The
	moved File is found by its recorded old address; when several moved Files share one (Frappe
	stores identical content once, so attachments of different records can share a disk name),
	the redirect goes to one this user may read. The target, open_file, checks permission again
	exactly as the old URL would have, so this reveals nothing to someone who could not open the
	file before. Other requests cost a prefix check; old-style file paths one indexed lookup.
	"""
	request = getattr(frappe.local, "request", None)
	if not request or request.method not in ("GET", "HEAD"):
		return
	from urllib.parse import unquote

	path = unquote(request.path or "")
	if not path.startswith(("/private/files/", "/files/")) or ".." in path.split("/"):
		return
	if not has_file_fields():
		return
	moved = frappe.get_all(
		"File",
		filters={"custom_microsoft_local_url": path, "custom_microsoft_status": STORED},
		pluck="name",
		order_by="creation asc",
	)
	if not moved:
		return
	on_disk = os.path.exists(local_path(path))
	if on_disk and _readable_file_at(path):
		# The bytes there are also an unmoved File's own, which this user may read: Frappe serves it.
		return
	target = next((name for name in moved if frappe.has_permission("File", "read", doc=name)), None)
	if not target and not on_disk:
		# Nothing this user may read. open_file refuses them just as the old URL would have.
		target = moved[0]
	if target:
		from werkzeug.routing import RequestRedirect

		raise RequestRedirect(stored_url(target))


def _readable_file_at(url):
	"""Whether a File whose address is this local URL is one the user may read."""
	return any(
		frappe.has_permission("File", "read", doc=name)
		for name in frappe.get_all("File", filters={"file_url": url}, pluck="name")
	)


def _repoint_links(file_doc, old_url, new_url):
	"""Make links that already show the old /private/files URL point at SharePoint.

	The move happens a moment after the upload, by which time the form that made it has already
	drawn its sidebar, and the timeline has an "Attachment" comment whose link is frozen HTML.
	Both would lead to a Forbidden page once the disk copy is gone.
	"""
	try:
		for comment in frappe.get_all(
			"Comment",
			filters={
				"reference_doctype": file_doc.attached_to_doctype,
				"reference_name": file_doc.attached_to_name,
				"comment_type": "Attachment",
				"content": ["like", f"%{old_url}%"],
			},
			fields=["name", "content"],
		):
			frappe.db.set_value(
				"Comment",
				comment.name,
				"content",
				comment.content.replace(f"'{old_url}'", f"'{new_url}'").replace(
					f'"{old_url}"', f'"{new_url}"'
				),
				update_modified=False,
			)
		# Same transaction rule as the move itself.
		frappe.db.commit()  # nosemgrep
	except Exception:
		frappe.log_error(title=_("Microsoft 365: could not update attachment links"))

	frappe.publish_realtime(
		MOVED_EVENT,
		{
			"doctype": file_doc.attached_to_doctype,
			"docname": file_doc.attached_to_name,
			"file": file_doc.name,
			"file_url": new_url,
			"old_url": old_url,
		},
		doctype=file_doc.attached_to_doctype,
		docname=file_doc.attached_to_name,
	)


def upload_name(file_doc, folder):
	"""The name a File is stored under in its record's folder: made safe for SharePoint, and
	shortened (keeping its extension) where the whole path would pass SharePoint's limit."""
	name = safe_file_name(file_doc.file_name or os.path.basename(file_doc.file_url or ""))
	used = folder_path_length(folder)
	fitted = fit_file_name(name, MAX_PATH_LENGTH - used - 1)
	if not fitted:
		frappe.throw(
			_(
				"This record's SharePoint folder path is already {0} characters long, counted from the site address, and SharePoint allows {1} for a file's whole path, so no file fits in it. Shorten the mapping's base folder, or the folder's name in SharePoint."
			).format(used, MAX_PATH_LENGTH),
			MsGraphError,
		)
	return fitted


def _upload(file_doc, folder):
	name = upload_name(file_doc, folder)
	content_type = mimetypes.guess_type(name)[0] or "application/octet-stream"
	if is_local(file_doc):
		# get_full_path refuses anything outside the site's files directories.
		path = file_doc.get_full_path()
		with open(path, "rb") as fh:  # nosemgrep
			return upload_stream(
				folder.drive_id, folder.item_id, name, fh, os.path.getsize(path), content_type
			)
	with _download_link(file_doc.file_url) as (fh, size, remote_type):
		return upload_stream(folder.drive_id, folder.item_id, name, fh, size, remote_type or content_type)


def upload_stream(drive_id, parent_id, name, fh, size, content_type="application/octet-stream"):
	"""Upload a file-like object into a folder. Returns the new driveItem.

	An existing file of the same name is never overwritten: Microsoft renames the new one
	(``Spec (1).pdf``), which is what a person dropping a file into Teams would get too.
	"""
	target = f"/drives/{drive_id}/items/{parent_id}:/{quote(name, safe='')}:"
	if size <= SIMPLE_UPLOAD_MAX_BYTES:
		return graph.graph_request(
			"PUT",
			f"{target}/content",
			APP,
			data=fh.read(),
			params={"@microsoft.graph.conflictBehavior": "rename"},
			headers={"Content-Type": content_type},
			timeout=120,
		)

	session = graph.graph_request(
		"POST",
		f"{target}/createUploadSession",
		APP,
		json={
			"item": {"@microsoft.graph.conflictBehavior": "rename", "name": name},
		},
	)
	upload_url = session["uploadUrl"]
	try:
		return _send_ranges(upload_url, fh, size)
	except Exception:
		# An abandoned session holds its partial bytes on Microsoft's side until it expires.
		try:
			graph._release(graph._http_request("DELETE", upload_url, timeout=30))
		except Exception:
			pass
		raise


def chunk_ranges(size, chunk=UPLOAD_CHUNK_BYTES):
	"""``[(start, end_inclusive), ...]`` covering ``size`` bytes. Pure."""
	return [(start, min(start + chunk, size) - 1) for start in range(0, size, chunk)]


def _send_ranges(upload_url, fh, size):
	"""PUT each range to the pre-authenticated upload URL.

	No Authorization header: the URL carries its own credential, and Microsoft documents that
	sending a bearer token to it can make the request fail.
	"""
	for start, end in chunk_ranges(size):
		body = fh.read(end - start + 1)
		resp = graph._http_request(
			"PUT",
			upload_url,
			data=body,
			timeout=300,
			headers={
				"Content-Length": str(len(body)),
				"Content-Range": f"bytes {start}-{end}/{size}",
			},
		)
		if resp.status_code in (200, 201):
			return resp.json()
		if resp.status_code != 202:
			detail = graph._safe_error(resp)
			graph._release(resp)
			frappe.throw(
				_("SharePoint refused part of the upload ({0}): {1}").format(resp.status_code, detail),
				MsGraphError,
			)
		graph._release(resp)
	frappe.throw(_("SharePoint did not confirm the upload after the last part."), MsGraphError)


class _download_link:
	"""Fetch a linked attachment into a temporary file. Context manager: ``(fh, size, type)``."""

	def __init__(self, url):
		self.url = url
		self.fh = None

	def __enter__(self):
		url = self.url
		resp = _pinned_get(url)
		hops = 0
		while resp.status in REDIRECT_STATUSES and hops < 5:
			location = resp.headers.get("Location") or ""
			resp.release_conn()
			from urllib.parse import urljoin

			url = urljoin(url, location)
			# Every hop is checked and pinned again: a public page may redirect anywhere.
			resp = _pinned_get(url)
			hops += 1
		if resp.status != 200:
			resp.release_conn()
			frappe.throw(_("Could not download {0} ({1}).").format(self.url, resp.status), MsGraphError)

		self.fh = tempfile.TemporaryFile()
		size = 0
		try:
			for chunk in resp.stream(DOWNLOAD_CHUNK_BYTES):
				size += len(chunk)
				if size > MAX_LINK_DOWNLOAD_BYTES:
					frappe.throw(
						_("{0} is larger than {1} MB; not archived.").format(
							self.url, MAX_LINK_DOWNLOAD_BYTES // (1024 * 1024)
						),
						MsGraphError,
					)
				self.fh.write(chunk)
		finally:
			resp.release_conn()
		self.fh.seek(0)
		content_type = (resp.headers.get("Content-Type") or "").split(";")[0].strip() or None
		return self.fh, size, content_type

	def __exit__(self, *exc):
		if self.fh:
			self.fh.close()
		return False


REDIRECT_STATUSES = (301, 302, 303, 307, 308)


def _pinned_get(url, timeout=60):
	"""GET a public URL, connecting only to the address assert_public_url checked.

	Resolving the name a second time to connect, as any HTTP library does, would let a name that
	answers differently the second time (DNS rebinding) send the request to an internal address
	after the check passed. So the connection is made to the checked IP itself, with the Host
	header and, for https, the TLS server name and certificate check set to the name in the link.
	No proxy is used either, for the same reason: the proxy would resolve the name again.
	Returns a urllib3 response with its body still unread.
	"""
	import certifi
	import urllib3
	from requests.utils import requote_uri

	url = requote_uri(url)
	address = assert_public_url(url)
	parsed = urlparse(url)
	https = parsed.scheme == "https"
	host = parsed.hostname
	host_header = f"[{host}]" if ":" in host else host
	if parsed.port:
		host_header += f":{parsed.port}"
	options = {
		"host": address,
		"port": parsed.port or (443 if https else 80),
		"timeout": urllib3.Timeout(connect=min(timeout, 30), read=timeout),
		"retries": False,
		"maxsize": 1,
	}
	if https:
		pool = urllib3.HTTPSConnectionPool(
			cert_reqs="CERT_REQUIRED",
			ca_certs=certifi.where(),
			server_hostname=host,
			assert_hostname=host,
			**options,
		)
	else:
		pool = urllib3.HTTPConnectionPool(**options)
	target = (parsed.path or "/") + (f"?{parsed.query}" if parsed.query else "")
	try:
		return pool.urlopen(
			"GET",
			target,
			headers={"Host": host_header, "User-Agent": "frappe_microsoft365", "Accept-Encoding": "identity"},
			redirect=False,
			retries=False,
			preload_content=False,
		)
	except urllib3.exceptions.HTTPError as e:
		frappe.throw(_("Could not download {0}: {1}").format(url, e), MsGraphError)


def assert_public_url(url):
	"""Refuse to fetch anything but a public http(s) address; return the address to connect to.

	A link attachment is typed by a user and fetched by the server, so without this it is a way
	to make the server read its own metadata endpoint or a neighbour on the private network.
	Every address the name has must be public, and the caller connects to the one returned
	(see _pinned_get) rather than resolving the name again.
	"""
	parsed = urlparse(url or "")
	if parsed.scheme not in ("http", "https") or not parsed.hostname:
		frappe.throw(_("Only http and https links can be archived."), MsGraphError)
	try:
		infos = socket.getaddrinfo(
			parsed.hostname, parsed.port or (443 if parsed.scheme == "https" else 80), type=socket.SOCK_STREAM
		)
	except (socket.gaierror, ValueError):
		frappe.throw(_("Could not resolve {0}.").format(parsed.hostname), MsGraphError)
	if not infos:
		frappe.throw(_("Could not resolve {0}.").format(parsed.hostname), MsGraphError)
	for info in infos:
		address = ipaddress.ip_address(info[4][0].split("%", 1)[0])
		if not address.is_global:
			frappe.throw(
				_("{0} is not a public address; not archived.").format(parsed.hostname), MsGraphError
			)
	return infos[0][4][0]


def _remove_local_copy(local_url):
	"""Delete the blob on disk unless another File row still uses it.

	Frappe stores identical content once and points every File with that hash at the same
	path, so the blob for one attachment may well be another record's attachment too.
	"""
	from frappe.core.doctype.file.utils import delete_file

	if frappe.db.exists("File", {"file_url": local_url}):
		return
	try:
		delete_file(local_url)
	except Exception:
		frappe.log_error(title=_("Microsoft 365: could not remove local copy"), message=local_url)


def _privacy_prefix(file_doc):
	"""Where a File's bytes belong on disk, by its privacy as it is now."""
	return "/private/files/" if cint(file_doc.get("is_private")) else "/files/"


def _blob_users(url, except_name):
	"""Other File rows that use the bytes at a local URL, as their address or as their kept copy."""
	return frappe.get_all(
		"File",
		filters={"name": ["!=", except_name or ""]},
		or_filters={"file_url": url, "custom_microsoft_local_url": url},
		pluck="name",
	)


def follow_privacy(file_doc):
	"""Put a moved File's kept local copy where its privacy says: public/files or private/files.

	Frappe moves a File's bytes when its privacy changes, but a moved File's file_url is the
	open_file address, so Frappe leaves the kept copy alone. Left in public/files, the copy of a
	File made private would stay readable by anyone through the web server. Moved when no other
	File uses those bytes, copied when one does (that File's bytes are its own business).
	Returns the kept copy's new URL, or None when there was nothing to move.
	"""
	local = file_doc.get("custom_microsoft_local_url")
	if file_doc.get("custom_microsoft_status") != STORED or not local or ".." in local.split("/"):
		return None
	prefix = _privacy_prefix(file_doc)
	if local.startswith(prefix) or not own_local_copy(file_doc, local):
		# In place already, or nothing of this File's is there to move (keep-local off, or the
		# name belongs to another file by now).
		return None
	basename = local.rsplit("/", 1)[-1]
	target = prefix + basename
	if os.path.exists(local_path(target)) or _blob_users(target, file_doc.name):
		target = prefix + f"{file_doc.name}-{basename}"
		if os.path.exists(local_path(target)):
			target = prefix + f"{file_doc.name}-{frappe.generate_hash(length=8)}-{basename}"
	source, destination = local_path(local), local_path(target)
	if _blob_users(local, file_doc.name):
		shutil.copyfile(source, destination)
	else:
		os.replace(source, destination)
	frappe.db.set_value("File", file_doc.name, "custom_microsoft_local_url", target, update_modified=False)
	file_doc.custom_microsoft_local_url = target
	return target


def _drop_kept_copy(file_doc):
	"""Remove a moved File's kept local copy, unless another File still uses those bytes."""
	local = file_doc.get("custom_microsoft_local_url")
	if file_doc.get("custom_microsoft_status") != STORED or not local:
		return
	if _blob_users(local, file_doc.name) or not own_local_copy(file_doc, local):
		return
	try:
		# own_local_copy accepted the URL: a /files or /private/files path without "..".
		os.remove(local_path(local))  # nosemgrep
	except OSError:
		frappe.log_error(title=_("Microsoft 365: could not remove local copy"), message=local)


# --- deletes and renames -----------------------------------------------------------------


def on_file_trash(doc, method=None):
	"""File.on_trash: optionally send the SharePoint copy to the site's recycle bin.

	Only an item the app uploaded itself (Stored or Archived, in a record folder; see
	assert_app_item) is ever deleted. Failed or Pending rows may hold the id of a partial upload,
	which is left alone.
	"""
	item_id = doc.get("custom_microsoft_item_id")
	if not item_id or not doc.get("custom_microsoft_drive_id"):
		return
	if doc.get("custom_microsoft_status") not in (STORED, ARCHIVED):
		return
	settings = _settings()
	if not (files_enabled(settings) and storage_option("files_delete_remote", doc.attached_to_doctype)):
		return
	# Attachment copies share the item; the last one out removes it.
	others = frappe.db.count("File", {"custom_microsoft_item_id": item_id, "name": ["!=", doc.name]})
	if others:
		return
	try:
		assert_app_item(doc, for_delete=True)
		graph.graph_request("DELETE", f"/drives/{doc.custom_microsoft_drive_id}/items/{item_id}", APP)
	except frappe.PermissionError:
		_forget_message()
		frappe.log_error(
			title=_("Microsoft 365: SharePoint item not removed"),
			message=_(
				"File {0} names an item outside its own record's folder; it was left in SharePoint."
			).format(doc.name),
		)
	except MsGraphNotFound:
		_forget_message()
		pass
	except Exception:
		# Removing the attachment in Frappe must not fail because SharePoint was unreachable.
		frappe.log_error(title=_("Microsoft 365: could not remove file from SharePoint"))


def on_file_update(doc, method=None):
	"""File.on_update: keep a moved file's SharePoint copy and kept local copy in step with the File.

	Its privacy decides where the kept local copy lives (follow_privacy), and the record it is
	attached to decides which folder its SharePoint item is in (follow_attachment). Inserts are
	left alone: a new File has nothing to follow yet, and an attachment copy shares its item with
	the original, which stays where it is.
	"""
	before = doc.get_doc_before_save()
	if doc.flags.in_insert or not before:
		return
	status = doc.get("custom_microsoft_status")
	if status == STORED and doc.has_value_changed("is_private"):
		try:
			follow_privacy(doc)
		except Exception:
			frappe.log_error(title=_("Microsoft 365: could not move the kept local copy"), message=doc.name)
			if cint(doc.is_private):
				# A copy that could not follow the File into private/files must not stay public.
				# SharePoint holds the file; the kept copy is only a spare.
				_drop_kept_copy(doc)
	if status in (STORED, ARCHIVED) and (
		doc.has_value_changed("attached_to_doctype") or doc.has_value_changed("attached_to_name")
	):
		old = (before.get("attached_to_doctype"), before.get("attached_to_name"))
		try:
			follow_attachment(doc.name, *old)
		except Exception:
			frappe.log_error(title=_("Microsoft 365: could not move a re-attached file"), message=doc.name)
			enqueue_follow_attachment(doc.name, *old)


def enqueue_follow_attachment(file_name, old_doctype, old_name):
	frappe.enqueue(
		"frappe_microsoft365.microsoft_files.follow_attachment",
		queue="long",
		file_name=file_name,
		old_doctype=old_doctype,
		old_name=old_name,
		in_job=True,
		job_id=f"microsoft365-follow-{file_name}",
		deduplicate=True,
		enqueue_after_commit=True,
	)


def follow_attachment(file_name, old_doctype, old_name, in_job=False):
	"""Move a stored File's SharePoint item to the folder of the record it is attached to now.

	Whoever can read a record sees and opens everything in its folder, so a File re-attached to
	another record must not stay in the first record's folder for that record's readers. Only an
	item that is in the folder of the record the File was attached to before is moved: anything
	else is not this File's to move (see assert_app_item).

	Within one library the item is moved in place, at once. Copying it between libraries, giving
	a File its own copy of an item it shares with attachment copies (they stay with the original),
	and bringing it back to this server when its new record has no SharePoint folder to go to all
	transfer the bytes, so they run as a background job (in_job).
	"""
	file_doc = frappe.get_doc("File", file_name)
	drive_id = file_doc.get("custom_microsoft_drive_id")
	item_id = file_doc.get("custom_microsoft_item_id")
	if file_doc.get("custom_microsoft_status") not in (STORED, ARCHIVED) or not (drive_id and item_id):
		return
	doctype, name = file_doc.attached_to_doctype, file_doc.attached_to_name
	if (doctype, name) == (old_doctype, old_name):
		return

	holder = _holder_of(file_doc)
	if not holder:
		return
	if (holder.reference_doctype, holder.reference_name) == (doctype, name):
		return  # already in the new record's folder
	if (holder.reference_doctype, holder.reference_name) != (old_doctype, old_name):
		frappe.log_error(
			title=_("Microsoft 365: re-attached file not moved"),
			message=_("File {0} names an item in the folder of {1} {2}, not of the record it was attached to.").format(
				file_name, holder.reference_doctype, holder.reference_name
			),
		)
		return

	shared = frappe.db.exists("File", {"custom_microsoft_item_id": item_id, "name": ["!=", file_name]})
	target = _folder_to_follow_into(file_doc)
	if target and not shared and target.drive_id == drive_id:
		item = graph.graph_request(
			"PATCH",
			f"/drives/{drive_id}/items/{item_id}",
			APP,
			json={"parentReference": {"id": target.item_id}},
			params={"@microsoft.graph.conflictBehavior": "rename"},
		)
		_record_location(file_name, item, target)
		return
	if not in_job:
		enqueue_follow_attachment(file_name, old_doctype, old_name)
		return
	if target:
		_copy_into(file_doc, target)
	elif file_doc.get("custom_microsoft_status") == STORED:
		# Nowhere in SharePoint to go: the record's DocType is not mapped (or storage is off), so
		# the file comes back to this server like any other attachment of that DocType.
		bring_back(file_name)
	else:
		# An archived link: the link itself never moved, so only the copy has to leave.
		frappe.db.set_value(
			"File", file_name, {"custom_microsoft_status": None, "custom_microsoft_item_id": None}, update_modified=False
		)
	if not shared:
		try:
			graph.graph_request("DELETE", f"/drives/{drive_id}/items/{item_id}", APP)
		except MsGraphNotFound:
			_forget_message()


def _folder_to_follow_into(file_doc):
	"""The folder a re-attached File's item belongs in now, or None when its new record has none
	to give: an unmapped DocType, storage switched off, a field attachment (those stay on this
	server, see _eligible), or an On demand record nobody asked a folder for. Creating the folder
	commits (see ensure_folder)."""
	doctype, name = file_doc.attached_to_doctype, file_doc.attached_to_name
	if not (files_enabled() and doctype and name) or file_doc.attached_to_field:
		return None
	row = mapping_for(doctype)
	if not row or (is_on_demand(row) and not get_folder(doctype, name)):
		return None
	return ensure_folder(doctype, name)


def _holder_of(file_doc):
	"""The SharePoint Folder row of the record folder a File's item is in, or None."""
	drive_id = file_doc.get("custom_microsoft_drive_id")
	folders = {
		row.item_id: row
		for row in frappe.get_all(
			"SharePoint Folder",
			filters={"drive_id": drive_id},
			fields=["item_id", "reference_doctype", "reference_name"],
		)
	}
	holder = folders.get(file_doc.get("custom_microsoft_folder_id") or "")
	if not holder:
		try:
			holder = folders.get(
				_known_folder_holding(drive_id, file_doc.get("custom_microsoft_item_id"), folders) or ""
			)
		except MsGraphNotFound:
			_forget_message()
			return None
	return holder


def _record_location(file_name, item, folder):
	"""Write where a File's SharePoint item now is."""
	frappe.db.set_value(
		"File",
		file_name,
		{
			"custom_microsoft_drive_id": _drive_of(item, folder),
			"custom_microsoft_item_id": item["id"],
			"custom_microsoft_folder_id": folder.item_id,
			"custom_microsoft_web_url": item.get("webUrl"),
		},
		update_modified=False,
	)


def _copy_into(file_doc, folder):
	"""Give a File its own copy of its SharePoint item, in ``folder`` (another library, or an
	item it shares with attachment copies). Streams through a temporary file."""
	resp = graph.graph_request(
		"GET",
		f"/drives/{file_doc.custom_microsoft_drive_id}/items/{file_doc.custom_microsoft_item_id}/content",
		APP,
		raw=True,
		stream=True,
		timeout=120,
	)
	with tempfile.TemporaryFile() as fh:
		try:
			for chunk in resp.iter_content(DOWNLOAD_CHUNK_BYTES):
				fh.write(chunk)
		finally:
			resp.close()
		size = fh.tell()
		fh.seek(0)
		name = upload_name(file_doc, folder)
		item = upload_stream(
			folder.drive_id, folder.item_id, name, fh, size, mimetypes.guess_type(name)[0] or "application/octet-stream"
		)
	_record_location(file_doc.name, item, folder)


def on_file_delete(doc, method=None):
	"""File.after_delete: remove a moved file's kept local copy with it.

	Frappe deletes a File's bytes only when it has a content_hash, which the move clears, so
	without this the kept copy would stay on disk (in public/files, for a public file) forever.
	"""
	try:
		_drop_kept_copy(doc)
	except Exception:
		frappe.log_error(title=_("Microsoft 365: could not remove local copy"), message=doc.name)


def on_doc_trash(doc, method=None):
	"""Any document's on_trash: forget its folder record. The SharePoint folder is kept."""
	if doc.doctype in ("SharePoint Folder", "File") or doc.doctype not in mapped_doctypes():
		return
	filters = {"reference_doctype": doc.doctype, "reference_name": doc.name}
	if frappe.db.exists("SharePoint Folder", filters):
		frappe.db.delete("SharePoint Folder", filters)
		touch_mapping(doc.doctype)


def on_doc_rename(doc, method=None, old=None, new=None, merge=False):
	"""Any document's after_rename: keep the record's folder linked to it under its new name.

	The SharePoint folder itself is never renamed. It may be one an administrator linked by hand
	(a Teams channel's folder, shared by people who never use the ERP), people bookmark and sync
	folders by name, and the link does not need it: it is kept by the folder's item id. Frappe
	normally rewrites the folder row's reference_name (a Dynamic Link) before this runs; this
	covers a row it did not.
	"""
	if doc.doctype not in mapped_doctypes():
		return
	folder = get_folder(doc.doctype, old)
	if not folder:
		return
	touch_mapping(doc.doctype)
	if merge or get_folder(doc.doctype, new):
		# Merged into a record that has a folder of its own: that one stays its folder.
		frappe.db.delete("SharePoint Folder", {"name": folder.name})
		return
	frappe.db.set_value("SharePoint Folder", folder.name, "reference_name", new, update_modified=False)


# --- retries and migration ---------------------------------------------------------------


def retry_pending():
	"""Hourly: re-queue files whose move is still pending or failed (up to MAX_ATTEMPTS)."""
	doctypes = mapped_doctypes()
	if not doctypes:
		return
	cutoff = add_to_date(now_datetime(), minutes=-RETRY_AFTER_MINUTES)
	names = frappe.get_all(
		"File",
		filters={
			"custom_microsoft_status": ["in", [PENDING, FAILED]],
			"custom_microsoft_attempts": ["<", MAX_ATTEMPTS],
			"attached_to_doctype": ["in", doctypes],
			"modified": ["<", cutoff],
		},
		pluck="name",
		limit=200,
	)
	for name in names:
		enqueue_upload(name)


@frappe.whitelist()
def queue_existing(doctype: str | None = None, limit: int = 500):
	"""Queue already-attached local files of mapped DocTypes for SharePoint. System Manager."""
	frappe.only_for("System Manager")
	doctypes = [doctype] if doctype else mapped_doctypes()
	doctypes = [dt for dt in doctypes if mapping_for(dt)]
	if not doctypes:
		frappe.throw(_("No DocType is mapped to SharePoint."))
	settings = _settings()
	candidates = frappe.get_all(
		"File",
		filters={
			"attached_to_doctype": ["in", doctypes],
			"is_folder": 0,
			"custom_microsoft_status": ["in", ["", None]],
		},
		fields=[
			"name",
			"file_url",
			"attached_to_doctype",
			"attached_to_name",
			"attached_to_field",
			"is_folder",
		],
		limit=cint(limit) or 500,
	)
	queued = 0
	for row in candidates:
		if not _eligible(frappe._dict(row), settings):
			continue
		frappe.db.set_value("File", row.name, "custom_microsoft_status", PENDING, update_modified=False)
		enqueue_upload(row.name)
		queued += 1
	return {"queued": queued, "more": len(candidates) >= (cint(limit) or 500)}


@frappe.whitelist()
def storage_summary():
	"""Counts by status, for the Settings form. System Manager."""
	frappe.only_for("System Manager")
	if not has_file_fields():
		return {}
	# One count per status rather than a GROUP BY: newer Frappe refuses SQL functions written as
	# strings in fields, and the dict form it wants instead does not exist on v15.
	counts = {}
	for status in (STORED, ARCHIVED, PENDING, FAILED):
		count = frappe.db.count("File", {"custom_microsoft_status": status})
		if count:
			counts[status] = count
	return counts


# --- reading -----------------------------------------------------------------------------

#: The only types opened in the browser tab; everything else downloads. An allow-list, because
#: what is not on it — HTML, SVG, XML, and anything whose type is unknown — a browser may render
#: as active content on our origin. A file with no extension has no type of its own, and the
#: Content-Type SharePoint reports for it is whatever was uploaded, so it is never trusted.
INLINE_TYPES = ("application/pdf", "text/plain", "text/csv")
INLINE_TYPE_PREFIXES = ("image/", "audio/", "video/")


def opens_inline(content_type):
	"""Pure. Whether a file of this type may be shown in the browser rather than downloaded."""
	if not content_type or content_type == "image/svg+xml":
		return False
	return content_type in INLINE_TYPES or content_type.startswith(INLINE_TYPE_PREFIXES)


#: Always downloaded, whatever their type is guessed to be. Kept alongside the allow-list as a
#: second line: the same list Frappe applies to /private/files.
FORCE_DOWNLOAD_EXTENSIONS = (
	".svg",
	".svgz",
	".html",
	".htm",
	".xhtml",
	".xht",
	".shtml",
	".shtm",
	".mhtml",
	".mht",
	".xml",
	".xsl",
	".xslt",
	".swf",
)


def _stream(drive_id, item_id, filename, download=False):
	"""Relay a driveItem's bytes to the browser without holding them in memory.

	Graph answers /content with a 302 to a pre-signed URL; requests follows it and drops our
	Authorization header on the way, since the host changes.
	"""
	resp = graph.graph_request(
		"GET", f"/drives/{drive_id}/items/{item_id}/content", APP, raw=True, stream=True, timeout=60
	)
	headers = resp.headers or {}
	content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
	inline = (
		not download
		and opens_inline(content_type)
		and not filename.lower().endswith(FORCE_DOWNLOAD_EXTENSIONS)
	)

	response = Response(
		resp.iter_content(chunk_size=DOWNLOAD_CHUNK_BYTES), direct_passthrough=True, content_type=content_type
	)
	response.headers.add("Content-Disposition", "inline" if inline else "attachment", filename=filename)
	response.headers["X-Content-Type-Options"] = "nosniff"
	response.headers["Cache-Control"] = "private, max-age=0"
	if headers.get("Content-Length"):
		response.headers["Content-Length"] = headers["Content-Length"]
	response.call_on_close(resp.close)
	return response


def assert_app_item(file_doc, for_delete=False):
	"""Refuse to read or delete a File's SharePoint item unless the app put it in a record folder.

	A File the upload job moved records the folder it went into, and only the app writes that
	(see on_file_validate), so checking it against the SharePoint Folder rows costs no Graph call.
	Files moved before that was recorded are checked by asking SharePoint where the item is: one
	request for a file directly in its folder, one more per level of subfolder.

	An item in another record's folder is an attachment copy (Frappe copies attachments when one
	document is made from another) or a row that was pointed there before 1.1.1 refused such
	edits; the two look the same. Either way it is read only by someone who may read the record
	that folder belongs to, and never deleted through this File.

	Raises PermissionError when the item is anywhere else, and MsGraphNotFound when SharePoint has
	no such item at all.
	"""
	drive_id = file_doc.get("custom_microsoft_drive_id")
	item_id = file_doc.get("custom_microsoft_item_id")
	if file_doc.get("custom_microsoft_status") not in (STORED, ARCHIVED) or not (drive_id and item_id):
		raise frappe.PermissionError(_("{0} is not stored in SharePoint.").format(file_doc.file_name))

	folders = {
		row.item_id: row
		for row in frappe.get_all(
			"SharePoint Folder",
			filters={"drive_id": drive_id},
			fields=["item_id", "reference_doctype", "reference_name"],
		)
	}
	# A record folder is never a file of its own, whatever a File row says.
	holder = None
	if item_id not in folders:
		holder = folders.get(file_doc.get("custom_microsoft_folder_id") or "")
		if not holder:
			holder = folders.get(_known_folder_holding(drive_id, item_id, folders) or "")
	if not holder:
		raise frappe.PermissionError(_("{0} is not in a record's SharePoint folder.").format(file_doc.file_name))

	if (holder.reference_doctype, holder.reference_name) == (
		file_doc.get("attached_to_doctype"),
		file_doc.get("attached_to_name"),
	):
		return
	if for_delete or not frappe.has_permission(holder.reference_doctype, "read", holder.reference_name):
		raise frappe.PermissionError(
			_("{0} is in the SharePoint folder of {1} {2}.").format(
				file_doc.file_name, _(holder.reference_doctype), holder.reference_name
			)
		)


def _known_folder_holding(drive_id, item_id, known):
	"""Ask SharePoint which of the ``known`` folders holds a file, at any depth; None for none."""
	item = graph.graph_request(
		"GET", f"/drives/{drive_id}/items/{item_id}", APP, params={"$select": "id,file,parentReference"}
	)
	if "file" not in item:
		return None
	for _hop in range(MAX_FOLDER_DEPTH):
		parent = item.get("parentReference") or {}
		if parent.get("driveId") and parent["driveId"] != drive_id:
			return None
		current = parent.get("id")
		if not current:
			return None
		if current in known:
			return current
		try:
			item = graph.graph_request(
				"GET", f"/drives/{drive_id}/items/{current}", APP, params={"$select": "id,parentReference"}
			)
		except MsGraphNotFound:
			_forget_message()
			return None
	return None


def _folder_still_there(folder):
	"""Ask for the folder itself, by id: True unless SharePoint says it is gone."""
	try:
		graph.graph_request(
			"GET", f"/drives/{folder.drive_id}/items/{folder.item_id}", APP, params={"$select": "id"}
		)
	except MsGraphNotFound:
		_forget_message()
		return False
	return True


def own_local_copy(file_doc, url):
	"""Is the file at a /files or /private/files URL still this File's own bytes?

	Once a file has moved without a kept copy its disk name is free, and Frappe gives that exact
	name to the next upload with the same file name, which may be another record's confidential
	file. So: when the bytes were hashed before the move, they must still match. Files moved
	before that was recorded have nothing to compare, and their old path counts as theirs only
	while no other File claims it, as its address or as the kept copy of a file that has itself
	moved since.
	"""
	if not url or ".." in url.split("/") or not url.startswith(("/files/", "/private/files/")):
		return False
	path = local_path(url)
	if not os.path.isfile(path):
		return False
	recorded = file_doc.get("custom_microsoft_content_hash")
	if recorded:
		return _hash_of(path) == recorded
	others = {"name": ["!=", file_doc.name]}
	if frappe.db.exists("File", {**others, "file_url": url}):
		return False
	return not frappe.db.exists("File", {**others, "custom_microsoft_local_url": url})


@frappe.whitelist(methods=["GET"], allow_guest=True)
def open_file(file: str, download: int = 0):
	"""Open a File that was moved to SharePoint. The file_url of every moved File points here.

	Permission is the File's own, exactly as for /files and /private/files: any File row carrying
	this URL that the user may read will do, because attachment copies share one URL. A public
	File (is_private=0) was readable by anyone at /files before it moved — an image on a web
	page, a link in a customer email — so it still is, signed in or not; a private one needs a
	signed-in user who may read it.
	"""
	from frappe.core.doctype.access_log.access_log import make_access_log
	from frappe.core.doctype.file.utils import find_file_by_url

	guest = frappe.session.user == "Guest"
	file_doc = find_file_by_url(stored_url(file))
	if guest and (not file_doc or cint(file_doc.is_private)):
		raise frappe.PermissionError(_("You must be logged in to open this file."))
	if not file_doc:
		raise frappe.PermissionError(_("You don't have permission to access this file"))
	if file_doc.get("custom_microsoft_status") != STORED:
		raise frappe.PermissionError(_("{0} is not stored in SharePoint.").format(file_doc.file_name))

	local = file_doc.get("custom_microsoft_local_url")
	try:
		assert_app_item(file_doc)
		response = _stream(
			file_doc.custom_microsoft_drive_id,
			file_doc.custom_microsoft_item_id,
			file_doc.file_name,
			download=cint(download),
		)
	except MsGraphNotFound:
		_forget_message()
		# Gone from SharePoint. A kept local copy is still the file; serve that instead, if it is.
		if local and own_local_copy(file_doc, local):
			if local.startswith("/private/files/"):
				from frappe.utils.response import send_private_file

				return send_private_file(local.split("/private", 1)[1])
			return Response(status=302, headers={"Location": local})
		frappe.throw(
			_(
				"{0} is no longer in SharePoint — it may have been deleted or moved out of the record's folder."
			).format(file_doc.file_name),
			frappe.DoesNotExistError,
		)
	make_access_log(
		doctype="File", document=file_doc.name, file_type=os.path.splitext(file_doc.file_name)[-1][1:]
	)
	return response


def fetch_content(drive_id, item_id):
	"""The whole content of a driveItem as bytes. For File.get_content; not for big files."""
	resp = graph.graph_request(
		"GET", f"/drives/{drive_id}/items/{item_id}/content", APP, raw=True, timeout=120
	)
	try:
		return resp.content
	finally:
		graph._release(resp)


# --- the form panel ----------------------------------------------------------------------


def _check_record(doctype, name, ptype="read"):
	if not mapping_for(doctype):
		frappe.throw(_("{0} is not mapped to SharePoint in Microsoft Settings.").format(_(doctype)))
	frappe.has_permission(doctype, ptype, doc=name, throw=True)


@frappe.whitelist()
def list_folder(doctype: str, name: str, subfolder: str | None = None):
	"""What is in a record's SharePoint folder (or one of its subfolders), for the form panel."""
	_check_record(doctype, name)
	folder = get_folder(doctype, name)
	can_write = can_create_folder(doctype, name)
	if not folder:
		return {"exists": False, "can_create": can_write, "on_demand": is_on_demand(mapping_for(doctype))}

	parent_id = folder.item_id
	if subfolder:
		_assert_inside(folder, subfolder)
		parent_id = subfolder
	try:
		items = graph.graph_paged(
			f"/drives/{folder.drive_id}/items/{parent_id}/children"
			"?$select=id,name,size,file,folder,webUrl,lastModifiedDateTime,lastModifiedBy&$top=200",
			APP,
		)
	except MsGraphNotFound:
		_forget_message()
		if subfolder:
			raise
		# Forgetting the folder means the next attachment makes a new one beside it ("Name 1"),
		# splitting the record's files, so one 404 from a listing is not taken at its word.
		if _folder_still_there(folder):
			frappe.throw(
				_("SharePoint did not list this record's folder just now. Try again in a moment."),
				MsGraphError,
			)
		_forget_folder(folder)
		return {"exists": False, "can_create": can_write, "missing": True}

	linked = {
		r.custom_microsoft_item_id: r.name
		for r in frappe.get_all(
			"File",
			filters={
				"attached_to_doctype": doctype,
				"attached_to_name": name,
				"custom_microsoft_item_id": ["is", "set"],
			},
			fields=["name", "custom_microsoft_item_id"],
		)
	}
	out = []
	for item in items:
		out.append(
			{
				"id": item["id"],
				"name": item.get("name"),
				"is_folder": "folder" in item,
				"child_count": (item.get("folder") or {}).get("childCount"),
				"size": item.get("size"),
				"modified": item.get("lastModifiedDateTime"),
				"modified_by": ((item.get("lastModifiedBy") or {}).get("user") or {}).get("displayName"),
				"web_url": item.get("webUrl"),
				"attachment": linked.get(item["id"]),
			}
		)
	out.sort(key=lambda i: (not i["is_folder"], (i["name"] or "").lower()))
	return {
		"exists": True,
		"web_url": folder.web_url,
		"folder_name": folder.folder_name,
		"path": folder.folder_path,
		"items": out,
		"can_create": can_write,
	}


@frappe.whitelist(methods=["POST"])
def create_folder(doctype: str, name: str):
	"""Create a record's folder now, and send its attachments that are still here into it.

	For an On demand mapping this is how a record starts living in SharePoint; for an Automatic
	one it only does early what the first attachment would have done.
	"""
	_check_record(doctype, name, "write")
	if not get_folder(doctype, name) and not can_create_folder(doctype, name):
		raise frappe.PermissionError(
			_(
				"Only System Managers and the roles listed on the {0} SharePoint Mapping can create folders."
			).format(_(doctype))
		)
	folder = ensure_folder(doctype, name)
	queued = queue_record(doctype, name)
	return {"web_url": folder.web_url, "folder_name": folder.folder_name, "queued": queued}


def _assert_inside(folder, item_id):
	"""Prove item_id is the record's folder or somewhere beneath it, or refuse.

	The panel hands item ids to the browser, so the browser can hand back any id at all; this is
	what stops open_item being a way to read the whole library through one readable record.
	"""
	current = item_id
	for _hop in range(MAX_FOLDER_DEPTH + 1):
		if current == folder.item_id:
			return
		try:
			item = graph.graph_request(
				"GET",
				f"/drives/{folder.drive_id}/items/{current}",
				APP,
				params={"$select": "id,parentReference"},
			)
		except MsGraphNotFound:
			_forget_message()
			break
		parent = item.get("parentReference") or {}
		if parent.get("driveId") and parent["driveId"] != folder.drive_id:
			break
		current = parent.get("id")
		if not current:
			break
	raise frappe.PermissionError(_("That file is not in this record's folder."))


@frappe.whitelist(methods=["GET"])
def open_item(doctype: str, name: str, item_id: str, download: int = 0):
	"""Open any file in a record's folder — including ones added from Teams — via Frappe."""
	_check_record(doctype, name)
	folder = get_folder(doctype, name)
	if not folder:
		frappe.throw(_("This record has no SharePoint folder yet."), frappe.DoesNotExistError)
	_assert_inside(folder, item_id)
	item = graph.graph_request(
		"GET", f"/drives/{folder.drive_id}/items/{item_id}", APP, params={"$select": "id,name,file"}
	)
	if "file" not in item:
		frappe.throw(_("That is a folder, not a file."))
	return _stream(folder.drive_id, item_id, item.get("name") or "file", download=cint(download))


# --- connection test ---------------------------------------------------------------------

PROBE_FOLDER = "frappe-connection-test"


@frappe.whitelist(methods=["POST"])
def test_connection(mapping: str | None = None):
	"""Prove each mapping (or just ``mapping``) works end to end: token, site, library, folder, write.

	Writes on purpose — a read-only grant passes every read and then fails on the first upload,
	which is the failure this button exists to catch early. It creates and immediately removes a
	folder named PROBE_FOLDER in each base folder (it lands in the site recycle bin).
	"""
	from frappe_microsoft365.doctor import FAIL, PASS, SITES_SELECTED_DOC, WARN, explain_error, finding

	frappe.only_for("System Manager")
	settings = graph.get_settings()
	findings = []
	if not settings.get("use_files"):
		frappe.throw(_("Tick SharePoint document storage first."))

	try:
		graph.clear_app_token_cache()
		token = graph.get_app_access_token()
	except Exception as e:
		explained = explain_error(str(e))
		fix = (
			explained["detail"]
			if explained.get("matched")
			else _(
				"Check Tenant ID, Client ID and the client secret Value; app-only sign-in needs a specific tenant."
			)
		)
		return {
			"findings": [
				finding(
					"files.token",
					FAIL,
					_("The app could not sign in as itself"),
					str(e),
					fix,
					explained.get("doc") or "",
				)
			]
		}

	roles = token_roles(token)
	if not ({"Sites.Selected", "Sites.ReadWrite.All", "Sites.FullControl.All"} & set(roles)):
		# SharePoint answers a token with no site permission at all with a bare 401
		# "generalException", which reads like an outage. Say what it is instead.
		return {
			"findings": [
				finding(
					"files.consent",
					FAIL,
					_("The app signed in, but its token carries no SharePoint permission"),
					_("Application permissions in the token: {0}.").format(", ".join(roles) or _("none")),
					_(
						"In Entra > App registrations > this app > API permissions, add Microsoft Graph > "
						"Application permissions > Sites.Selected (Application, not Delegated), then click "
						"Grant admin consent so its status reads Granted. It must be listed under Microsoft "
						"Graph: the SharePoint API has a permission with the same name that this app cannot "
						"use. Run this test again a minute later."
					),
					SITES_SELECTED_DOC,
				)
			]
		}

	for row in _mappings():
		if mapping and row.name != mapping:
			continue
		if not row.enabled and not mapping:
			continue
		target = row.reference_doctype
		try:
			row.site_id = row.drive_id = None  # re-resolve: the point is to test the URL as typed
			_site_id, drive_id = resolve_drive(row)
		except Exception as e:
			findings.append(
				finding(
					"files.site",
					FAIL,
					_("Cannot reach {0}").format(row.site_url),
					str(e),
					_site_fix(str(e)),
					target=target,
				)
			)
			continue

		segments = base_segments(row)
		try:
			base = ensure_path(drive_id, segments) if segments else _get_by_path(drive_id, [])
			probe = _create_folder(drive_id, base["id"], PROBE_FOLDER)
			graph.graph_request("DELETE", f"/drives/{drive_id}/items/{probe['id']}", APP)
		except Exception as e:
			findings.append(
				finding(
					"files.write",
					FAIL,
					_("Can read {0} but not write to it").format(row.site_url),
					str(e),
					_site_fix(str(e)),
					target=target,
				)
			)
			continue

		where = "/".join([row.library or "Documents", *segments])
		findings.append(
			finding(
				"files.ok",
				PASS,
				_("{0} files go to {1}").format(target, where),
				base.get("webUrl") or "",
				"",
				target=target,
			)
		)

	if not findings:
		findings.append(
			finding(
				"files.none",
				WARN,
				_("No DocType is mapped yet"),
				"",
				_("Create a SharePoint Mapping for the DocType."),
			)
		)
	return {"findings": findings}


def token_roles(token):
	"""The application permissions ("roles" claim) in an access token. Pure; never logs the token.

	Only the payload is decoded and nothing is verified: this is for telling an admin what
	Microsoft granted, not for trusting the token, which Graph itself validates on every call.
	"""
	import base64
	import json

	try:
		payload = token.split(".")[1]
		payload += "=" * (-len(payload) % 4)
		return list(json.loads(base64.urlsafe_b64decode(payload)).get("roles") or [])
	except Exception:
		return []


def _site_fix(message):
	if "401" in message:
		return _(
			"SharePoint did not accept the app's token. Check that Sites.Selected is an Application "
			"permission under Microsoft Graph (not under SharePoint) with admin consent granted, then "
			"wait a minute and test again."
		)
	if "403" in message or "accessDenied" in message or "Forbidden" in message:
		return _(
			"The app has not been granted this site. Add the Sites.Selected application permission with "
			"admin consent, then run the Site Grant Script on SharePoint Settings as a SharePoint admin."
		)
	if "404" in message or "itemNotFound" in message:
		return _(
			"Check the site URL and the library name; Teams channel files are in the library called Documents."
		)
	return ""


# --- bringing files back (uninstall) -------------------------------------------------------


def bring_back(file_name):
	"""Put a moved file back on this server and point its File row at it again.

	The reverse of upload_file, for uninstall: once this app is gone nothing can answer
	open_file, so every attachment that lives in SharePoint has to come home first. The
	SharePoint copy is left where it is. Returns the local URL the file now has.
	"""
	file_doc = frappe.get_doc("File", file_name)
	if file_doc.get("custom_microsoft_status") != STORED:
		return file_doc.file_url

	old_url = file_doc.file_url
	# A kept copy in the wrong place for the File's privacy is moved first, so that what comes
	# back is where the File's privacy says, never at the old public address of a private file.
	follow_privacy(file_doc)
	target = file_doc.get("custom_microsoft_local_url")
	prefix = _privacy_prefix(file_doc)
	if not target or ".." in target.split("/") or not target.startswith(prefix):
		target = prefix + safe_file_name(
			(target or "").rsplit("/", 1)[-1] or file_doc.file_name or file_doc.name
		)
	path = local_path(target)

	# A kept copy is reused only if it is still this file's bytes (see own_local_copy).
	if not own_local_copy(file_doc, target):
		if os.path.exists(path) or frappe.db.exists(
			"File", {"file_url": target, "name": ["!=", file_doc.name]}
		):
			# The name is taken by another attachment; never write over someone else's file, and
			# never take it for this one.
			target = prefix + f"{file_doc.name}-" + safe_file_name(file_doc.file_name or "file")
			path = local_path(target)
		resp = graph.graph_request(
			"GET",
			f"/drives/{file_doc.custom_microsoft_drive_id}/items/{file_doc.custom_microsoft_item_id}/content",
			APP,
			raw=True,
			stream=True,
			timeout=120,
		)
		tmp = path + ".part"
		try:
			# The path comes from local_path() of a URL checked for ".." above.
			with open(tmp, "wb") as fh:  # nosemgrep
				for chunk in resp.iter_content(DOWNLOAD_CHUNK_BYTES):
					fh.write(chunk)
			os.replace(tmp, path)
		finally:
			resp.close()
			if os.path.exists(tmp):
				os.remove(tmp)

	frappe.db.set_value(
		"File",
		file_doc.name,
		{"file_url": target, "content_hash": _hash_of(path), "custom_microsoft_status": None},
		update_modified=False,
	)
	for comment in frappe.get_all(
		"Comment",
		filters={
			"reference_doctype": file_doc.attached_to_doctype,
			"reference_name": file_doc.attached_to_name,
			"comment_type": "Attachment",
			"content": ["like", f"%{old_url}%"],
		},
		fields=["name", "content"],
	):
		frappe.db.set_value(
			"Comment",
			comment.name,
			"content",
			comment.content.replace(old_url, target),
			update_modified=False,
		)
	return target


def leave_in_sharepoint(file_name):
	"""Turn a moved attachment into a plain link to its SharePoint copy (uninstall, no download).

	Once this app is gone nothing answers open_file, but a File whose file_url is an https link
	is something Frappe opens by itself: the browser goes straight to SharePoint, which lets in
	whoever has access to the site. Returns the URL the attachment now has.
	"""
	file_doc = frappe.get_doc("File", file_name)
	if file_doc.get("custom_microsoft_status") != STORED:
		return file_doc.file_url
	old_url = file_doc.file_url
	web_url = file_doc.get("custom_microsoft_web_url")
	if not web_url:
		item = graph.graph_request(
			"GET",
			f"/drives/{file_doc.custom_microsoft_drive_id}/items/{file_doc.custom_microsoft_item_id}",
			APP,
			params={"$select": "webUrl"},
		)
		web_url = item.get("webUrl")
	if not web_url:
		frappe.throw(_("SharePoint gave no address for {0}.").format(file_doc.file_name))
	frappe.db.set_value(
		"File", file_doc.name, {"file_url": web_url, "custom_microsoft_status": None}, update_modified=False
	)
	for comment in frappe.get_all(
		"Comment",
		filters={
			"reference_doctype": file_doc.attached_to_doctype,
			"reference_name": file_doc.attached_to_name,
			"comment_type": "Attachment",
			"content": ["like", f"%{old_url}%"],
		},
		fields=["name", "content"],
	):
		frappe.db.set_value(
			"Comment",
			comment.name,
			"content",
			comment.content.replace(old_url, web_url),
			update_modified=False,
		)
	return web_url


# --- upgrade checks (read only; run after `bench migrate`) -------------------------------------

#: How many names an upgrade finding lists before it says "and N more".
UPGRADE_LIST_LIMIT = 20


def _names(rows):
	names = [r if isinstance(r, str) else r.name for r in rows]
	shown = ", ".join(names[:UPGRADE_LIST_LIMIT])
	if len(names) > UPGRADE_LIST_LIMIT:
		shown += " " + _("and {0} more").format(len(names) - UPGRADE_LIST_LIMIT)
	return shown


def upgrade_check_shared_folders():
	"""SharePoint Folder rows of different records that point at one folder.

	Until 1.1.1 a record could be given a folder that was already another record's (same title,
	names equal once cleaned, a deleted record's kept folder). Each record's readers see the whole
	folder, so such records show each other's files. The app cannot tell which record the folder
	belongs to, so it only reports them. Read only.
	"""
	from frappe_microsoft365.doctor import WARN, finding

	groups = {}
	for row in frappe.get_all(
		"SharePoint Folder",
		filters={"item_id": ["is", "set"]},
		fields=["item_id", "reference_doctype", "reference_name", "folder_path"],
		order_by="reference_doctype asc, reference_name asc",
	):
		groups.setdefault(row.item_id, []).append(row)

	out = []
	for rows in groups.values():
		records = sorted({(r.reference_doctype, r.reference_name) for r in rows})
		if len(records) < 2:
			continue
		out.append(
			finding(
				"upgrade.sharepoint_shared_folders",
				WARN,
				_("{0} records share the SharePoint folder {1}").format(
					len(records), rows[0].folder_path or rows[0].item_id
				),
				", ".join(f"{_(dt)} {name}" for dt, name in records),
				_(
					"People who can read any of these records see every file in that folder. On the SharePoint Mapping, delete the Folders rows of the records that should not have it; their next attachment gets a folder of its own. Move their files in SharePoint first if needed."
				),
				target=rows[0].reference_doctype,
			)
		)
	return out


def upgrade_check_files_in_other_records_folders():
	"""Moved Files whose item is in the folder of a record other than the one they are attached to.

	From 1.1.2 such an item opens only for someone who may read the record that owns the folder.
	These are attachment copies (one document made from another, the copy pointing at the
	original's item) or rows pointed at another record's file before 1.1.1 refused that. Read only
	and without Graph: the recorded folder, or else the file's address, says whose folder it is.
	"""
	from frappe_microsoft365.doctor import WARN, finding

	if not has_file_fields():
		return []
	has_folder_id = frappe.get_meta("File").has_field("custom_microsoft_folder_id")
	fields = [
		"name",
		"attached_to_doctype",
		"attached_to_name",
		"custom_microsoft_drive_id",
		"custom_microsoft_web_url",
	] + (["custom_microsoft_folder_id"] if has_folder_id else [])
	rows = frappe.get_all(
		"File", filters={"custom_microsoft_status": ["in", [STORED, ARCHIVED]]}, fields=fields, order_by="name asc"
	)
	if not rows:
		return []

	folders = frappe.get_all(
		"SharePoint Folder", fields=["drive_id", "item_id", "web_url", "reference_doctype", "reference_name"]
	)
	by_id = {(f.drive_id, f.item_id): f for f in folders if f.item_id}
	by_url = {(f.drive_id, (f.web_url or "").rstrip("/").lower()): f for f in folders if f.web_url}

	def holder(row):
		found = by_id.get((row.custom_microsoft_drive_id, row.get("custom_microsoft_folder_id")))
		if found:
			return found
		url = (row.custom_microsoft_web_url or "").rstrip("/").lower()
		while "/" in url:
			url = url.rsplit("/", 1)[0]
			found = by_url.get((row.custom_microsoft_drive_id, url))
			if found:
				return found
		return None

	elsewhere = []
	for row in rows:
		folder = holder(row)
		if folder and (folder.reference_doctype, folder.reference_name) != (
			row.attached_to_doctype,
			row.attached_to_name,
		):
			elsewhere.append(row)
	if not elsewhere:
		return []
	return [
		finding(
			"upgrade.sharepoint_files_in_other_folders",
			WARN,
			_("{0} File(s) open only for users who can read the record whose folder holds them").format(
				len(elsewhere)
			),
			_names(elsewhere),
			_(
				"Each is attached to one record but stored in another record's SharePoint folder: usually an attachment copied when one document was made from another, otherwise a File edited to point at someone else's file before 1.1.1. Someone who cannot read the record that owns the folder can no longer open it. If a copy has to be readable on its own, attach the file to that record again."
			),
		)
	]


def upgrade_check_file_items():
	"""Files whose SharePoint item cannot be shown to sit in a known record folder.

	From 1.1.1 a File's SharePoint fields can only be set by the app, and opening, reading or
	deleting a moved file checks that its item is in a record folder. Rows written before could
	have been edited by anyone; this lists the ones that do not look like the app's own work, so
	an admin can review them. Read only and without Graph: where a row records no folder (every
	file moved before 1.1.1), its SharePoint address is compared with the folders' addresses.
	"""
	from frappe_microsoft365.doctor import FAIL, WARN, finding

	if not has_file_fields():
		return []
	has_folder_id = frappe.get_meta("File").has_field("custom_microsoft_folder_id")
	fields = [
		"name",
		"file_url",
		"custom_microsoft_status",
		"custom_microsoft_drive_id",
		"custom_microsoft_item_id",
		"custom_microsoft_web_url",
	]
	if has_folder_id:
		fields.append("custom_microsoft_folder_id")
	rows = [
		frappe._dict(
			name=r.name,
			file_url=r.file_url,
			status=r.custom_microsoft_status,
			drive_id=r.custom_microsoft_drive_id,
			item_id=r.custom_microsoft_item_id,
			web_url=r.custom_microsoft_web_url,
			folder_id=r.get("custom_microsoft_folder_id"),
		)
		for r in frappe.get_all(
			"File",
			or_filters={
				"custom_microsoft_status": ["in", [STORED, ARCHIVED]],
				"file_url": ["like", f"%{OPEN_METHOD}%"],
			},
			fields=fields,
			order_by="name asc",
		)
	]
	if not rows:
		return []

	folders = frappe.get_all("SharePoint Folder", fields=["drive_id", "item_id", "web_url"])
	folder_ids = {(f.drive_id, f.item_id) for f in folders if f.item_id}
	folder_urls = {(f.drive_id, (f.web_url or "").rstrip("/").lower()) for f in folders if f.web_url}

	def under_known_folder(row):
		if row.folder_id and (row.drive_id, row.folder_id) in folder_ids:
			return True
		url = (row.web_url or "").rstrip("/").lower()
		while "/" in url:
			url = url.rsplit("/", 1)[0]
			if (row.drive_id, url) in folder_urls:
				return True
		return False

	is_folder, odd_status, outside = [], [], []
	for row in rows:
		if (row.drive_id, row.item_id) in folder_ids:
			is_folder.append(row)
		elif row.status not in (STORED, ARCHIVED) or not (row.drive_id and row.item_id):
			odd_status.append(row)
		elif row.status == STORED and not is_open_url(row.file_url):
			odd_status.append(row)
		elif not under_known_folder(row):
			outside.append(row)

	fix = _(
		"Open each File and compare its Item ID and Open in SharePoint address with the record it is attached to. Files the app moved itself are fine; a File that names someone else's item should be deleted. Until then the app refuses to open or delete what it names unless SharePoint confirms the item is in a record folder."
	)
	out = []
	if is_folder:
		out.append(
			finding(
				"upgrade.sharepoint_file_items",
				FAIL,
				_("{0} File(s) name a record's SharePoint folder as their item").format(len(is_folder)),
				_names(is_folder),
				fix,
			)
		)
	if odd_status:
		out.append(
			finding(
				"upgrade.sharepoint_file_items",
				WARN,
				_("{0} File(s) carry SharePoint fields the app would not have written").format(
					len(odd_status)
				),
				_(
					"They point at the app's open address without being marked Stored, or are marked Stored or Archived without an item: {0}"
				).format(_names(odd_status)),
				fix,
			)
		)
	if outside:
		out.append(
			finding(
				"upgrade.sharepoint_file_items",
				WARN,
				_("{0} File(s) in SharePoint are not under any known record folder").format(len(outside)),
				_(
					"Judged by their recorded SharePoint address, so a folder renamed or removed since the move also shows up here: {0}"
				).format(_names(outside)),
				fix,
			)
		)
	return out


def upgrade_check_orphaned_kept_copies():
	"""Kept local copies of moved Files that were deleted before 1.2.0, still on disk.

	Frappe removes a File's bytes only when it has a content_hash, which the move clears, so until
	1.2.0 deleting a moved File left its kept copy behind (in public/files for a public file). The
	deleted rows are read back from Deleted Document; a copy is listed when it is still on disk,
	no File uses it, and, where the bytes were hashed at the move, they are still those bytes.
	Read only: nothing is deleted, because only a person can say nothing else needs the file.
	"""
	from frappe_microsoft365.doctor import WARN, finding

	if not has_file_fields() or not frappe.db.table_exists("Deleted Document"):
		return []
	filters = {"deleted_doctype": "File", "data": ["like", '%"custom_microsoft_local_url": "/%']}
	if frappe.get_meta("Deleted Document").has_field("restored"):
		filters["restored"] = 0
	left = []
	for row in frappe.get_all(
		"Deleted Document", filters=filters, fields=["deleted_name", "data"], order_by="creation asc"
	):
		try:
			data = frappe._dict(json.loads(row.data))
		except ValueError:
			continue
		local = data.get("custom_microsoft_local_url") or ""
		if data.get("custom_microsoft_status") != STORED or local in left:
			continue
		if ".." in local.split("/") or not local.startswith(("/files/", "/private/files/")):
			continue
		path = local_path(local)
		if not os.path.isfile(path) or _blob_users(local, None):
			continue
		recorded = data.get("custom_microsoft_content_hash")
		if recorded and _hash_of(path) != recorded:
			continue
		left.append(local)
	if not left:
		return []
	public = [url for url in left if url.startswith("/files/")]
	return [
		finding(
			"upgrade.sharepoint_orphaned_kept_copies",
			WARN,
			_("{0} kept local copies of deleted Files are still on this server ({1} in public files)").format(
				len(left), len(public)
			),
			_names(left),
			_(
				"These Files were moved to SharePoint with Keep Local Copy on and deleted later; their copies on disk were never removed, and the ones under /files/ can be downloaded by anyone who has the address. Nothing in Frappe uses them. Delete them from the site's public/files and private/files folders unless something outside Frappe needs them. Files deleted from 1.2.0 on take their kept copy with them."
			),
		)
	]


def upgrade_check_long_paths():
	"""SharePoint Mappings whose library and base folder leave little of SharePoint's path limit.

	SharePoint refuses a file whose whole path passes 400 characters. From 1.2.0 a mapping whose
	base path leaves too little room is refused when it is saved, and record folders are named to
	leave room for their files; mappings saved before are listed here. Read only, without Graph.
	"""
	from frappe_microsoft365.doctor import WARN, finding

	if not frappe.db.table_exists("SharePoint Mapping"):
		return []
	out = []
	for row in frappe.get_all(
		"SharePoint Mapping",
		fields=["name", "reference_doctype", "site_url", "library", "use_base_folder", "base_folder"],
		order_by="name asc",
	):
		try:
			used = base_path_length(row)
		except Exception:
			continue  # an unreadable site address is the doctor's to report
		if used <= MAX_PATH_LENGTH - MIN_ROOM_FOR_RECORDS:
			continue
		out.append(
			finding(
				"upgrade.sharepoint_long_paths",
				WARN,
				_("The {0} mapping's base folder uses {1} of SharePoint's {2} path characters").format(
					_(row.reference_doctype), used, MAX_PATH_LENGTH
				),
				"/".join([library_path(row), *base_segments(row)]),
				_(
					"Uploads into it fail once a record folder and a file name take the path past {0} characters. Shorten the Base Folder on the mapping (and rename or move the folder in SharePoint to match); the mapping cannot be saved again until its base path leaves at least {1} characters."
				).format(MAX_PATH_LENGTH, MIN_ROOM_FOR_RECORDS),
				target=row.reference_doctype,
			)
		)
	return out
