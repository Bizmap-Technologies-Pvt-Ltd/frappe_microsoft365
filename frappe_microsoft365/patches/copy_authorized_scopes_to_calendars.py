"""Give every authorised Microsoft Calendar its own record of the scopes it signed in with.

Up to 1.1.x the app kept one site-wide "authorized scopes" value, overwritten by whoever signed
in last, and the doctor judged every connection by it. 1.2.0 records the scopes on each
connection instead. Connections that signed in before the upgrade never recorded theirs, so the
site-wide value is copied onto them: it is the best record there is, and starting them empty or
from the current settings would either hide a real gap or ask people to re-authorise because of
the upgrade alone.

Only fills connections that have no value yet, so it is safe to run again and never overwrites
what a sign-in after the upgrade recorded. No network calls.
"""

import frappe


def execute():
	if not frappe.db.table_exists("Microsoft Calendar"):
		return
	if not frappe.db.has_column("Microsoft Calendar", "authorized_scopes"):
		return

	site_wide = (frappe.db.get_single_value("Microsoft Settings", "authorized_scopes") or "").strip()
	if not site_wide:
		# Nothing was ever recorded: the doctor stays silent about these connections, as before.
		return

	names = frappe.get_all(
		"Microsoft Calendar",
		filters={"authorized": 1},
		or_filters=[["authorized_scopes", "is", "not set"], ["authorized_scopes", "=", ""]],
		pluck="name",
		limit_page_length=0,
	)
	for name in names:
		frappe.db.set_value("Microsoft Calendar", name, "authorized_scopes", site_wide, update_modified=False)
