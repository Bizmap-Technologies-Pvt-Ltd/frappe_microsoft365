"""How SharePoint document storage behaves: local copies, linked files, deletes, uninstall.

Turning the capability on stays in Microsoft Settings, beside the others, because it decides what
the Azure app is asked for. Everything about running it lives here, next to its mappings and folders.
"""

from frappe.model.document import Document


class SharePointSettings(Document):
	pass
