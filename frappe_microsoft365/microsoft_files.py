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
* ``PATCH /drives/{drive-id}/items/{item-id}`` — rename
* ``DELETE /drives/{drive-id}/items/{item-id}`` — to the site recycle bin
"""

import ipaddress
import mimetypes
import os
import re
import socket
import tempfile
from urllib.parse import quote, urlparse

import frappe
from frappe import _
from frappe.utils import add_to_date, cint, now_datetime
from werkzeug.wrappers import Response

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_graph import APP_ONLY, MsGraphConflict, MsGraphError, MsGraphNotFound

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

#: How far up the tree open_item walks to prove an item sits inside a record's folder.
MAX_FOLDER_DEPTH = 10

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
	settings = settings or _settings()
	return bool(settings.enabled and settings.get("use_files"))


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
	"""Tell the desk which DocTypes get the "Files in Microsoft 365" panel."""
	try:
		bootinfo.microsoft365_files_doctypes = mapped_doctypes()
	except Exception:
		# A broken settings row must never stop the desk from loading.
		bootinfo.microsoft365_files_doctypes = []


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
		APP_ONLY,
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


def _find_drive(site_id, library):
	library = (library or "").strip()
	drives = graph.graph_paged(f"/sites/{site_id}/drives?$select=id,name,webUrl", APP_ONLY)
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
		return graph.graph_request("GET", f"/sites/{site_id}/drive", APP_ONLY, params={"$select": "id"})["id"]
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
			"GET", path, APP_ONLY, params={"$select": "id,name,folder,webUrl,parentReference"}
		)
	except MsGraphNotFound:
		_forget_message()
		return None


def _create_folder(drive_id, parent_id, name):
	"""Create a folder, or return the one a concurrent job created a moment earlier."""
	parent = f"/drives/{drive_id}/items/{parent_id}" if parent_id else f"/drives/{drive_id}/root"
	try:
		return graph.graph_request(
			"POST",
			f"{parent}/children",
			APP_ONLY,
			json={
				"name": name,
				"folder": {},
				"@microsoft.graph.conflictBehavior": "fail",
			},
		)
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
	"""The folder for a record, created in SharePoint (and recorded here) on first use."""
	existing = get_folder(doctype, name)
	if existing:
		return existing

	row = mapping_for(doctype)
	if not row:
		frappe.throw(_("{0} is not mapped to SharePoint in Microsoft Settings.").format(_(doctype)))

	values = new_folder_values(row, doctype, name)

	# A second job may have finished first; one record per document, whoever wins.
	existing = get_folder(doctype, name)
	if existing:
		return existing

	return frappe._dict(add_folder_row(doctype, name, values).as_dict())


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


def new_folder_values(row, doctype, name):
	"""Create (or find) the record's folder where its mapping says, and return the fields to store."""
	_site_id, drive_id = resolve_drive(row)
	doc = frappe.get_doc(doctype, name)
	folder_name = render_pattern(row.folder_pattern, doc)
	segments = [*base_segments(row), folder_name]
	item = ensure_path(drive_id, segments)
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

	_site_id, drive_id = resolve_drive(row)
	value = (location or "").strip()
	item = None
	if value.lower().startswith(("http://", "https://")):
		parsed = urlparse(value)
		ids = parse_qs(parsed.query).get("id")
		path = unquote(ids[0] if ids else parsed.path).rstrip("/")
		drive = graph.graph_request("GET", f"/drives/{drive_id}", APP_ONLY, params={"$select": "webUrl"})
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
					APP_ONLY,
					params={"$select": "id,name,folder,webUrl,parentReference"},
				)
			except MsGraphError:
				_forget_message()
				item = None
	else:
		item = _get_by_path(drive_id, [p for p in value.replace("\\", "/").split("/") if p.strip()])

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


def on_file_insert(doc, method=None):
	"""File.after_insert: queue the move. Never fails the upload itself."""
	try:
		if not _eligible(doc):
			return
		doc.db_set(
			{"custom_microsoft_status": PENDING, "custom_microsoft_error": None}, update_modified=False
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
	# A savepoint, not a full rollback, on failure: a half-recorded folder must not survive, but
	# neither should the caller's own unrelated writes be thrown away with it.
	frappe.db.savepoint("microsoft365_upload")
	try:
		folder = ensure_folder(file_doc.attached_to_doctype, file_doc.attached_to_name)
		try:
			item = _upload(file_doc, folder)
		except MsGraphNotFound:
			_forget_message()
			# Someone deleted or moved the record's folder out of the library. Recreate it once.
			_forget_folder(folder)
			folder = ensure_folder(file_doc.attached_to_doctype, file_doc.attached_to_name)
			item = _upload(file_doc, folder)
	except Exception as e:
		frappe.db.rollback(save_point="microsoft365_upload")
		frappe.db.set_value(
			"File",
			file_name,
			{
				"custom_microsoft_status": FAILED,
				"custom_microsoft_error": str(e)[:1000],
				"custom_microsoft_attempts": attempts,
			},
			update_modified=False,
		)
		# The failure has to outlive the job that hit it, or the hourly retry never sees it.
		frappe.db.commit()  # nosemgrep
		if attempts >= MAX_ATTEMPTS:
			frappe.log_error(
				title=_("Microsoft 365: file not stored in SharePoint"), message=f"File {file_name}: {e}"
			)
		return

	local_url = file_doc.file_url if is_local(file_doc) else None
	keep_local = bool(storage_option("files_keep_local_copy", file_doc.attached_to_doctype))
	values = {
		"custom_microsoft_status": STORED if local_url else ARCHIVED,
		"custom_microsoft_drive_id": item.get("parentReference", {}).get("driveId") or folder.drive_id,
		"custom_microsoft_item_id": item["id"],
		"custom_microsoft_web_url": item.get("webUrl"),
		"custom_microsoft_error": None,
		"custom_microsoft_attempts": attempts,
	}
	if local_url:
		values["file_url"] = stored_url(file_name)
		values["content_hash"] = None
		# Kept whether or not the disk copy is: links already handed out (a bookmark, an email,
		# a form opened before the move) still say this URL, and redirect_moved_file follows it.
		values["custom_microsoft_local_url"] = local_url
	frappe.db.set_value("File", file_name, values, update_modified=False)
	# The row must point at SharePoint, durably, before the copy on disk is removed.
	frappe.db.commit()  # nosemgrep

	if local_url:
		_repoint_links(file_doc, local_url, values["file_url"])
	if local_url and not keep_local:
		_remove_local_copy(local_url)


MOVED_EVENT = "microsoft365_file_moved"


def redirect_moved_file():
	"""before_request: send an old /files or /private/files link of a moved file to its new home.

	The disk copy is gone, so Frappe would answer the old link with Forbidden (private) or Not
	Found (public). The redirect target, open_file, checks permission exactly as the old URL
	would have, so this reveals nothing to someone who could not open the file before.
	Costs nothing on other requests: a prefix check, then one indexed lookup only when the
	file is missing from disk.
	"""
	request = getattr(frappe.local, "request", None)
	if not request or request.method not in ("GET", "HEAD"):
		return
	from urllib.parse import unquote

	path = unquote(request.path or "")
	if not path.startswith(("/private/files/", "/files/")) or ".." in path.split("/"):
		return
	if os.path.exists(local_path(path)):
		return
	name = frappe.db.get_value(
		"File", {"custom_microsoft_local_url": path, "custom_microsoft_status": STORED}, "name"
	)
	if name:
		from werkzeug.routing import RequestRedirect

		raise RequestRedirect(stored_url(name))


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
		frappe.db.commit()  # nosemgrep: same transaction rule as the move itself
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


def _upload(file_doc, folder):
	name = safe_file_name(file_doc.file_name or os.path.basename(file_doc.file_url or ""))
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
			APP_ONLY,
			data=fh.read(),
			params={"@microsoft.graph.conflictBehavior": "rename"},
			headers={"Content-Type": content_type},
			timeout=120,
		)

	session = graph.graph_request(
		"POST",
		f"{target}/createUploadSession",
		APP_ONLY,
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
		assert_public_url(self.url)
		import requests

		resp = requests.get(
			self.url,
			stream=True,
			timeout=60,
			allow_redirects=False,
			headers={"User-Agent": "frappe_microsoft365"},
		)
		hops = 0
		while resp.is_redirect and hops < 5:
			location = resp.headers.get("Location") or ""
			resp.close()
			nxt = requests.compat.urljoin(self.url, location)
			assert_public_url(nxt)
			resp = requests.get(
				nxt,
				stream=True,
				timeout=60,
				allow_redirects=False,
				headers={"User-Agent": "frappe_microsoft365"},
			)
			hops += 1
		if resp.status_code != 200:
			resp.close()
			frappe.throw(_("Could not download {0} ({1}).").format(self.url, resp.status_code), MsGraphError)

		self.fh = tempfile.TemporaryFile()
		size = 0
		for chunk in resp.iter_content(DOWNLOAD_CHUNK_BYTES):
			size += len(chunk)
			if size > MAX_LINK_DOWNLOAD_BYTES:
				resp.close()
				frappe.throw(
					_("{0} is larger than {1} MB; not archived.").format(
						self.url, MAX_LINK_DOWNLOAD_BYTES // (1024 * 1024)
					),
					MsGraphError,
				)
			self.fh.write(chunk)
		resp.close()
		self.fh.seek(0)
		content_type = (resp.headers.get("Content-Type") or "").split(";")[0].strip() or None
		return self.fh, size, content_type

	def __exit__(self, *exc):
		if self.fh:
			self.fh.close()
		return False


def assert_public_url(url):
	"""Refuse to fetch anything but a public http(s) address.

	A link attachment is typed by a user and fetched by the server, so without this it is a way
	to make the server read its own metadata endpoint or a neighbour on the private network.
	"""
	parsed = urlparse(url or "")
	if parsed.scheme not in ("http", "https") or not parsed.hostname:
		frappe.throw(_("Only http and https links can be archived."), MsGraphError)
	try:
		infos = socket.getaddrinfo(parsed.hostname, parsed.port or (443 if parsed.scheme == "https" else 80))
	except socket.gaierror:
		frappe.throw(_("Could not resolve {0}.").format(parsed.hostname), MsGraphError)
	for info in infos:
		address = ipaddress.ip_address(info[4][0])
		if not address.is_global:
			frappe.throw(
				_("{0} is not a public address; not archived.").format(parsed.hostname), MsGraphError
			)


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


# --- deletes and renames -----------------------------------------------------------------


def on_file_trash(doc, method=None):
	"""File.on_trash: optionally send the SharePoint copy to the site's recycle bin."""
	item_id = doc.get("custom_microsoft_item_id")
	if not item_id or not doc.get("custom_microsoft_drive_id"):
		return
	settings = _settings()
	if not (files_enabled(settings) and storage_option("files_delete_remote", doc.attached_to_doctype)):
		return
	# Attachment copies share the item; the last one out removes it.
	others = frappe.db.count("File", {"custom_microsoft_item_id": item_id, "name": ["!=", doc.name]})
	if others:
		return
	try:
		graph.graph_request("DELETE", f"/drives/{doc.custom_microsoft_drive_id}/items/{item_id}", APP_ONLY)
	except MsGraphNotFound:
		_forget_message()
		pass
	except Exception:
		# Removing the attachment in Frappe must not fail because SharePoint was unreachable.
		frappe.log_error(title=_("Microsoft 365: could not remove file from SharePoint"))


def on_doc_trash(doc, method=None):
	"""Any document's on_trash: forget its folder record. The SharePoint folder is kept."""
	if doc.doctype in ("SharePoint Folder", "File") or not frappe.db.table_exists("SharePoint Folder"):
		return
	if doc.doctype not in mapped_doctypes():
		return
	filters = {"reference_doctype": doc.doctype, "reference_name": doc.name}
	if frappe.db.exists("SharePoint Folder", filters):
		frappe.db.delete("SharePoint Folder", filters)
		touch_mapping(doc.doctype)


def on_doc_rename(doc, method=None, old=None, new=None, merge=False):
	"""Any document's after_rename: keep the folder record, and rename the folder to match."""
	if doc.doctype not in mapped_doctypes():
		return
	folder = get_folder(doc.doctype, old)
	if not folder:
		return
	touch_mapping(doc.doctype)
	if merge or get_folder(doc.doctype, new):
		frappe.db.delete("SharePoint Folder", {"name": folder.name})
		return
	frappe.db.set_value("SharePoint Folder", folder.name, "reference_name", new, update_modified=False)
	row = mapping_for(doc.doctype)
	new_name = render_pattern(row.folder_pattern, frappe.get_doc(doc.doctype, new))
	if new_name == folder.folder_name:
		return
	try:
		item = graph.graph_request(
			"PATCH",
			f"/drives/{folder.drive_id}/items/{folder.item_id}",
			APP_ONLY,
			json={
				"name": new_name,
				"@microsoft.graph.conflictBehavior": "fail",
			},
		)
		frappe.db.set_value(
			"SharePoint Folder",
			folder.name,
			{"folder_name": new_name, "web_url": item.get("webUrl")},
			update_modified=False,
		)
	except Exception:
		# The record is renamed either way; the folder keeps its old name and still links.
		frappe.log_error(title=_("Microsoft 365: could not rename SharePoint folder"))


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
	# One count per status rather than a GROUP BY: newer Frappe refuses SQL functions written as
	# strings in fields, and the dict form it wants instead does not exist on v15.
	counts = {}
	for status in (STORED, ARCHIVED, PENDING, FAILED):
		count = frappe.db.count("File", {"custom_microsoft_status": status})
		if count:
			counts[status] = count
	return counts


# --- reading -----------------------------------------------------------------------------

#: Served as a download even when asked to open inline: a browser renders these as active
#: content on our origin. The same list Frappe applies to /private/files.
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
		"GET", f"/drives/{drive_id}/items/{item_id}/content", APP_ONLY, raw=True, stream=True, timeout=60
	)
	headers = resp.headers or {}
	content_type = (
		mimetypes.guess_type(filename)[0] or headers.get("Content-Type") or "application/octet-stream"
	)
	inline = not download and not filename.lower().endswith(FORCE_DOWNLOAD_EXTENSIONS)

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


@frappe.whitelist(methods=["GET"])
def open_file(file: str, download: int = 0):
	"""Open a File that was moved to SharePoint. The file_url of every moved File points here.

	Permission is the File's own, exactly as for /private/files: any File row carrying this URL
	that the user may read will do, because attachment copies share one URL.
	"""
	from frappe.core.doctype.access_log.access_log import make_access_log
	from frappe.core.doctype.file.utils import find_file_by_url

	if frappe.session.user == "Guest":
		raise frappe.PermissionError(_("You must be logged in to open this file."))

	file_doc = find_file_by_url(stored_url(file))
	if not file_doc:
		raise frappe.PermissionError(_("You don't have permission to access this file"))

	local = file_doc.get("custom_microsoft_local_url")
	try:
		response = _stream(
			file_doc.custom_microsoft_drive_id,
			file_doc.custom_microsoft_item_id,
			file_doc.file_name,
			download=cint(download),
		)
	except MsGraphNotFound:
		_forget_message()
		# Gone from SharePoint. A kept local copy is still the file; serve that instead.
		if local and os.path.exists(local_path(local)):
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
		"GET", f"/drives/{drive_id}/items/{item_id}/content", APP_ONLY, raw=True, timeout=120
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
			APP_ONLY,
		)
	except MsGraphNotFound:
		_forget_message()
		if subfolder:
			raise
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
				APP_ONLY,
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
		"GET", f"/drives/{folder.drive_id}/items/{item_id}", APP_ONLY, params={"$select": "id,name,file"}
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
			graph.graph_request("DELETE", f"/drives/{drive_id}/items/{probe['id']}", APP_ONLY)
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
	import hashlib

	file_doc = frappe.get_doc("File", file_name)
	if file_doc.get("custom_microsoft_status") != STORED:
		return file_doc.file_url

	old_url = file_doc.file_url
	target = file_doc.get("custom_microsoft_local_url")
	prefix = "/private/files/" if file_doc.is_private else "/files/"
	if not target or ".." in target.split("/"):
		target = prefix + safe_file_name(file_doc.file_name or file_doc.name)
	path = local_path(target)

	if not os.path.exists(path):
		if frappe.db.exists("File", {"file_url": target, "name": ["!=", file_doc.name]}):
			# The name is taken by another attachment; never write over someone else's file.
			target = prefix + f"{file_doc.name}-" + safe_file_name(file_doc.file_name or "file")
			path = local_path(target)
		resp = graph.graph_request(
			"GET",
			f"/drives/{file_doc.custom_microsoft_drive_id}/items/{file_doc.custom_microsoft_item_id}/content",
			APP_ONLY,
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

	digest = hashlib.md5()
	with open(path, "rb") as fh:  # nosemgrep
		for chunk in iter(lambda: fh.read(DOWNLOAD_CHUNK_BYTES), b""):
			digest.update(chunk)

	frappe.db.set_value(
		"File",
		file_doc.name,
		{"file_url": target, "content_hash": digest.hexdigest(), "custom_microsoft_status": None},
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
			APP_ONLY,
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
