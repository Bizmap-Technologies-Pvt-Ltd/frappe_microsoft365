"""Retire the old hard-coded scope string so derivation can take over.

`default_scopes` used to be a pre-filled list of every scope the app might ever want. It is
now an OVERRIDE: empty means "derive the scopes from the capabilities that are ticked".

On a site that upgrades, the old pre-filled value is still sitting there, and it would now be
read as a deliberate override — pinning that site to scopes it may not want and, worse,
requesting Teams permissions an admin never granted. That is exactly the failure this change
set out to remove.

Clearing it is only safe when the value is untouched boilerplate, so this matches the known
historical defaults exactly. Anything an admin actually edited is left alone: a real override
must keep working.
"""

import frappe

#: Every value the field shipped with. Anything else is assumed to be deliberate.
LEGACY_DEFAULTS = {
	"User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite OnlineMeetingTranscript.Read.All",
	"User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite OnlineMeetingTranscript.Read.All ",
}


def execute():
	if not frappe.db.exists("DocType", "Microsoft Settings"):
		return

	current = frappe.db.get_single_value("Microsoft Settings", "default_scopes")
	if not current:
		return

	normalised = " ".join(current.split())
	if normalised in {" ".join(value.split()) for value in LEGACY_DEFAULTS}:
		frappe.db.set_single_value("Microsoft Settings", "default_scopes", "")
