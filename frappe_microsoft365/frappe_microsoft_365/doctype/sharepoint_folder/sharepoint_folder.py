"""One record's folder in SharePoint: a row of its DocType's SharePoint Mapping (Folders table).

Kept apart from the business document so that any DocType can be mapped without a schema change,
and keyed on the folder's item id, which survives someone renaming or moving the folder in Teams,
rather than on a path.

Rows are added two ways. The app inserts one directly when a record first needs its folder (an
Automatic mapping's first upload, or Create folder now on the record), touching the mapping's
modified time so that a mapping form opened earlier cannot save over the new row. People add one
with Add Row on the mapping; SharePointMapping.validate then creates or links the folder. Removing
a row only forgets the link: the folder and its files stay in SharePoint.
"""

from frappe.model.document import Document


class SharePointFolder(Document):
	pass
