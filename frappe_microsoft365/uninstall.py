"""Leave a site in a working state, with nothing left behind, when this app is removed.

Frappe's own uninstall deletes the app's DocTypes and module, but not custom fields added to
standard DocTypes (Event, File) or their columns, and it knows nothing about attachments that
were moved to SharePoint, whose links point at this app. This runs first and handles both:

1. Attachments that live in SharePoint are dealt with as the site chose in SharePoint Settings >
   *If this app is uninstalled* (or ``microsoft365_uninstall_files`` in site_config.json:
   ``restore`` / ``leave``):

   * **Bring files back to this server**: each is downloaded back and its File row pointed at it
     again. The free disk space is checked first, because bringing years of attachments home is
     exactly what can fill a hosted server; if it would not fit, nothing is downloaded.
   * **Leave files in SharePoint**: nothing is downloaded; each attachment becomes a link to its
     SharePoint copy, which opens for people with access to the site.
   * No choice made: asked in a terminal, otherwise the uninstall stops and says where to choose.

   Either way, if any file cannot be handled the uninstall stops and lists them, so no attachment
   is ever left pointing at an app that is gone. SharePoint copies are never deleted.
2. Every custom field this app added is deleted, and its column dropped.
3. Cached tokens and settings are cleared.

Frappe calls before_uninstall during ``--dry-run`` too, so a dry run only reports what would happen.
"""

import inspect
import os
import shutil
import sys

import frappe

RESTORE = "Bring files back to this server"
LEAVE = "Leave files in SharePoint"

#: Free space that must remain after bringing files back, on top of the files themselves.
DISK_MARGIN = 0.10


def _dry_run():
	"""Whether this uninstall is a dry run. Frappe does not pass it to the hook, so look for it."""
	frame = inspect.currentframe()
	while frame:
		if frame.f_code.co_name == "remove_app" and "dry_run" in frame.f_locals:
			return bool(frame.f_locals["dry_run"])
		frame = frame.f_back
	return False


def before_uninstall():
	dry_run = _dry_run()
	handle_files(dry_run=dry_run)
	remove_custom_fields(dry_run=dry_run)
	if not dry_run:
		clear_caches()


def _size(n):
	for unit in ("B", "KB", "MB", "GB", "TB"):
		if n < 1024 or unit == "TB":
			return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
		n /= 1024


def free_space():
	return shutil.disk_usage(frappe.get_site_path("private", "files")).free


def choice():
	"""How to handle files in SharePoint: RESTORE, LEAVE, or None when nobody has said."""
	conf = (frappe.conf.get("microsoft365_uninstall_files") or "").strip().lower()
	if conf in ("restore", "bring back"):
		return RESTORE
	if conf == "leave":
		return LEAVE
	if frappe.db.table_exists("Singles"):
		value = frappe.db.get_single_value("SharePoint Settings", "files_on_uninstall")
		if value in (RESTORE, LEAVE):
			return value
	if sys.stdin and sys.stdin.isatty():
		import click

		picked = click.prompt(
			"Attachments are stored in SharePoint. Bring them back to this server (restore), "
			"or leave them in SharePoint as links (leave)?",
			type=click.Choice(["restore", "leave"]),
		)
		return RESTORE if picked == "restore" else LEAVE
	return None


def handle_files(dry_run=False):
	if not frappe.db.has_column("File", "custom_microsoft_status"):
		return
	from frappe_microsoft365 import microsoft_files as files

	stored = frappe.get_all(
		"File", filters={"custom_microsoft_status": files.STORED}, fields=["name", "file_size"]
	)
	if not stored:
		return
	total = sum(int(r.file_size or 0) for r in stored)
	free = free_space()
	print(
		f"{len(stored)} attachment(s), {_size(total)} in all, are stored in SharePoint. "
		f"Free space on this server: {_size(free)}."
	)
	if dry_run:
		return

	how = choice()
	if not how:
		frappe.throw(
			"Uninstall stopped: {0} attachment(s) ({1}) are stored in SharePoint. Choose what happens "
			"to them first: SharePoint Settings > If this app is uninstalled "
			"(Bring files back to this server, or Leave files in SharePoint), then run the uninstall "
			"again.".format(len(stored), _size(total))
		)

	if how == RESTORE and total * (1 + DISK_MARGIN) > free:
		frappe.throw(
			"Uninstall stopped: bringing the files back needs about {0} and this server has {1} free. "
			"Nothing was downloaded. Free up space, or choose Leave files in SharePoint in SharePoint "
			"Settings > If this app is uninstalled.".format(_size(total * (1 + DISK_MARGIN)), _size(free))
		)

	action = files.bring_back if how == RESTORE else files.leave_in_sharepoint
	print(f"{how}…")
	failed = []
	for i, row in enumerate(stored, start=1):
		try:
			action(row.name)
			# Each file handled stays handled if a later one fails.
			frappe.db.commit()  # nosemgrep
		except Exception as e:
			frappe.db.rollback()
			failed.append(f"{row.name}: {e}"[:300])
		if i % 25 == 0:
			print(f"  {i}/{len(stored)}")

	if failed:
		frappe.throw(
			"Uninstall stopped: {0} attachment(s) could not be handled, so removing the app now would "
			"leave them unopenable here. Fix the cause (usually the connection: Microsoft Settings > "
			"Troubleshoot) and run the uninstall again; files already handled stay that way.\n{1}".format(
				len(failed), "\n".join(failed[:20])
			)
		)


def remove_custom_fields(dry_run=False):
	from frappe_microsoft365.setup import app_custom_fields

	for doctype, fieldnames in app_custom_fields().items():
		print(f"Removing {len(fieldnames)} custom field(s) from {doctype}")
		if dry_run:
			continue
		for fieldname in fieldnames:
			name = frappe.db.get_value("Custom Field", {"dt": doctype, "fieldname": fieldname})
			if name:
				frappe.delete_doc("Custom Field", name, ignore_permissions=True, force=True)
			if frappe.db.has_column(doctype, fieldname):
				frappe.db.sql_ddl(f"alter table `tab{doctype}` drop column `{fieldname}`")  # nosemgrep
		frappe.clear_cache(doctype=doctype)


def clear_caches():
	from frappe_microsoft365 import microsoft_files as files
	from frappe_microsoft365 import microsoft_graph as graph

	graph.clear_app_token_cache()
	files.clear_mapping_cache()
	frappe.db.delete("Singles", {"doctype": ["in", ["Microsoft Settings", "SharePoint Settings"]]})
