"""File behaviour for attachments that were moved to SharePoint (Frappe v16 extend_doctype_class).

A moved File's file_url is ``/api/method/...open_file``, which the browser follows happily but
server-side code cannot: anything that reads an attachment's bytes — attaching it to an email,
a print, an import — calls ``File.get_content()``, which opens a path on disk. This mixin answers
that call from SharePoint instead, decoding text the same way Frappe does for a file on disk.

Frappe v15 has no extend_doctype_class hook, so there the moved File opens from the form as
usual but server-side readers see no content; docs/sharepoint-files.md says so.
"""

from frappe_microsoft365 import microsoft_files


class SharePointFile:
	def _stored_in_sharepoint(self):
		return (
			self.get("custom_microsoft_status") == microsoft_files.STORED
			and self.get("custom_microsoft_item_id")
			and self.get("custom_microsoft_drive_id")
			and (self.file_url or "").startswith("/api/method/" + microsoft_files.OPEN_METHOD)
		)

	def get_content(self, encodings=None):
		if self.is_folder or self.get("content") or not self._stored_in_sharepoint():
			return super().get_content(encodings=encodings)

		from frappe.core.doctype.file.file import FILE_ENCODING_OPTIONS, OLE_FILE_SIGNATURE

		content = microsoft_files.fetch_content(self.custom_microsoft_drive_id, self.custom_microsoft_item_id)
		if not content.startswith(OLE_FILE_SIGNATURE):
			for encoding in encodings or FILE_ENCODING_OPTIONS:
				try:
					content = content.decode(encoding)
					break
				except UnicodeDecodeError:
					continue
		self._content = content
		return content

	def exists_on_disk(self):
		if self._stored_in_sharepoint():
			return False
		return super().exists_on_disk()
