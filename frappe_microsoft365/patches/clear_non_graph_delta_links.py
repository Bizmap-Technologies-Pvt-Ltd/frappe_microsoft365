"""Clear any stored calendar sync link that does not point at Microsoft Graph over HTTPS.

The sync calls its stored delta link with the calendar owner's token attached. Since 1.1.1 the
app refuses to send a token anywhere but https://graph.microsoft.com, and the link can no longer
be edited from outside the sync; this removes any link stored before that, so the calendar it
belongs to starts a fresh sync instead of failing on every run.

Links Graph handed out are always on that host, so on a normal site this changes nothing. Safe
to run again: a second run finds nothing left to clear.
"""

import frappe

from frappe_microsoft365.microsoft_graph import is_graph_url

#: Each link field, and the window field that only has meaning alongside it.
LINK_FIELDS = (
	("delta_link", "delta_window_end"),
	("delta_next_link", "delta_next_window_end"),
)


def execute():
	if not frappe.db.table_exists("Microsoft Calendar"):
		return

	for link_field, window_field in LINK_FIELDS:
		if not frappe.db.has_column("Microsoft Calendar", link_field):
			continue
		rows = frappe.get_all(
			"Microsoft Calendar",
			filters={link_field: ["is", "set"]},
			fields=["name", link_field],
			limit_page_length=0,
		)
		for row in rows:
			if is_graph_url(row.get(link_field)):
				continue
			updates = {link_field: None}
			if frappe.db.has_column("Microsoft Calendar", window_field):
				updates[window_field] = None
			frappe.db.set_value("Microsoft Calendar", row.name, updates, update_modified=False)
