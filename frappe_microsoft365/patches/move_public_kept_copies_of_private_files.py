"""Move kept local copies of private moved files out of public/files.

Until 1.2.0, a moved file's kept local copy stayed where it was when the file moved. A File made
private afterwards kept its copy in public/files, which the web server hands to anyone who knows
the address. This moves each such copy to private/files (see microsoft_files.follow_privacy).

Nothing is deleted and nothing is fetched from SharePoint. A copy whose bytes are no longer the
File's own, or that another File still uses, is left where it is or copied, never taken. Kept
copies of Files that were deleted are only reported, by upgrade.sharepoint_orphaned_kept_copies.
A site that never switched SharePoint storage on has no such column, and this does nothing.
Safe to run again: a second run finds every copy in place.
"""

import frappe

from frappe_microsoft365 import microsoft_files as files


def execute():
	if not frappe.db.has_column("File", "custom_microsoft_local_url"):
		return
	for name in frappe.get_all(
		"File",
		filters={
			"custom_microsoft_status": files.STORED,
			"is_private": 1,
			"custom_microsoft_local_url": ["like", "/files/%"],
		},
		pluck="name",
		limit_page_length=0,
	):
		try:
			files.follow_privacy(frappe.get_doc("File", name))
		except Exception:
			frappe.log_error(title="Microsoft 365: could not move a kept local copy", message=name)
