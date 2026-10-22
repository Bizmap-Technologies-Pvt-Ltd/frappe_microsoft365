"""SharePoint Mapping — where one DocType's attachments are stored in SharePoint.

One record per DocType, named after it, so a DocType can never be mapped twice. Lives beside
Microsoft Settings rather than inside it, like Microsoft Calendar: the settings hold the
connection and the global storage options, each mapping holds one destination and can be
tested on its own.

Its Folders table holds one row per record that has a folder. See SharePointFolder for how rows
get there; a row added by hand is resolved here, on save.
"""

import frappe
from frappe import _
from frappe.model.document import Document


class SharePointMapping(Document):
	def _validate_links(self):
		# Every Folders row is for this mapping's DocType. Frappe checks the rows' Dynamic Links
		# before any hook runs (before_validate included), so the DocType is filled in here.
		for row in self.folders:
			row.reference_doctype = self.reference_doctype
		super()._validate_links()

	def validate(self):
		meta = frappe.get_meta(self.reference_doctype)
		if (
			meta.istable
			or meta.issingle
			or self.reference_doctype
			in (
				"File",
				"SharePoint Folder",
				"SharePoint Mapping",
			)
		):
			frappe.throw(_("{0} cannot have its own SharePoint folders.").format(self.reference_doctype))

		# The resolved ids are a cache of the URL and the library name; editing either has to
		# drop them, or files keep flowing to the old library with nothing on screen saying so.
		if not self.is_new() and (self.has_value_changed("site_url") or self.has_value_changed("library")):
			self.site_id = None
			self.drive_id = None

		self.set_up_new_folders()

	def set_up_new_folders(self):
		"""Create or link the SharePoint folder of each row added by hand."""
		from frappe_microsoft365 import microsoft_files as files

		self.flags.new_folders = []
		seen = set()
		for row in self.folders:
			if row.reference_name in seen:
				frappe.throw(
					_("Row {0}: {1} {2} is listed twice.").format(
						row.idx, _(self.reference_doctype), row.reference_name
					),
					frappe.DuplicateEntryError,
				)
			seen.add(row.reference_name)
			if row.item_id:
				continue  # resolved already

			if not row.reference_name:
				frappe.throw(_("Row {0}: pick the {1}.").format(row.idx, _(self.reference_doctype)))
			mapping = files.mapping_for(self.reference_doctype)
			if not mapping or self.is_new():
				frappe.throw(
					_("Save this mapping with Enabled ticked before adding folders by hand."),
					title=_("Not enabled yet"),
				)
			frappe.has_permission(self.reference_doctype, "write", doc=row.reference_name, throw=True)
			if files.get_folder(self.reference_doctype, row.reference_name):
				frappe.throw(
					_("{0} {1} already has a SharePoint folder.").format(
						_(self.reference_doctype), row.reference_name
					),
					frappe.DuplicateEntryError,
				)

			if (row.existing_folder or "").strip():
				values = files.existing_folder_values(mapping, row.existing_folder)
			else:
				values = files.new_folder_values(mapping, self.reference_doctype, row.reference_name)
			row.update(values)
			# resolve_drive may have just looked these up; keep them rather than save blanks over them
			self.site_id, self.drive_id = mapping.site_id, mapping.drive_id
			self.flags.new_folders.append(row.reference_name)

	def on_update(self):
		from frappe_microsoft365 import microsoft_files as files

		files.clear_mapping_cache()
		queued = sum(
			files.queue_record(self.reference_doctype, name) for name in self.flags.new_folders or []
		)
		if queued:
			frappe.msgprint(
				_("{0} attachment(s) are on their way to the new folder(s).").format(queued), alert=True
			)

	def on_trash(self):
		from frappe_microsoft365.microsoft_files import clear_mapping_cache

		clear_mapping_cache()
