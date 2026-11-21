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

Clearing it also takes away the Teams and transcript scopes the old value asked for, and the
tickboxes that now ask for them (Standalone Teams meetings, Meeting transcripts) start unticked —
so transcript fetching stopped silently on sites that relied on it. Where people had already
authorised calendars with the old value, and so consented to those scopes, the capabilities the
old value covered are ticked as it is cleared. A site with no authorised calendar never asked
Microsoft for anything, and is left with nothing ticked, as before.

Sites that ran the first version of this patch lost the value before it could be read; the
1.2.0 upgrade check ``upgrade_check_meeting_capabilities_off`` tells them to re-tick by hand.
"""

import frappe
from frappe import _

#: Capabilities the legacy value covered, in the order they must be ticked (transcripts builds on
#: Teams meetings).
LEGACY_CAPABILITIES = ("use_teams", "use_transcripts")

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
		for field in capabilities_to_keep(normalised.split()):
			frappe.db.set_single_value("Microsoft Settings", field, 1)
		frappe.db.set_single_value("Microsoft Settings", "default_scopes", "")
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")


def capabilities_to_keep(scopes):
	"""The unticked capability fields whose Graph scopes the old value asked for, if consented.

	Consented means at least one Microsoft Calendar is authorised: its token was issued for the
	old value's scopes, so ticking the capability asks for nothing new it already had (apart from
	the recordings scope, added after this value was retired, which the doctor reports).
	"""
	from frappe_microsoft365 import capabilities

	if not frappe.db.table_exists("Microsoft Calendar") or not frappe.db.count(
		"Microsoft Calendar", {"authorized": 1}
	):
		return []
	meta = frappe.get_meta("Microsoft Settings")
	scopes = set(scopes)
	keep = []
	for capability in capabilities.all_capabilities():
		field = capability.field
		if field not in LEGACY_CAPABILITIES or not meta.has_field(field):
			continue
		if frappe.db.get_single_value("Microsoft Settings", field):
			continue
		if scopes & set(capability.graph_scopes):
			keep.append(field)
	return sorted(keep, key=LEGACY_CAPABILITIES.index)


def upgrade_check_meeting_capabilities_off():
	"""Read-only, 1.2.0: Teams or transcript data on Events while the capability is unticked.

	The first version of this patch cleared the old scope value without ticking Standalone Teams
	meetings or Meeting transcripts, so a site upgraded from before 1.0.0 stopped fetching
	transcripts without saying so. The cleared value is gone, but the Events remember: a stored
	online meeting id was only ever resolved with the Teams scope, and a transcript only fetched
	with the transcript scope. A site that never used either has no such Events and hears nothing.
	"""
	from frappe_microsoft365 import doctor

	if not frappe.db.exists("DocType", "Microsoft Settings"):
		return []
	checks = (
		(
			"use_teams",
			"custom_microsoft_online_meeting_id",
			_("Standalone Teams meetings"),
			_("{0} Event(s) hold a Teams online meeting id"),
		),
		(
			"use_transcripts",
			"custom_microsoft_transcript_fetched_on",
			_("Meeting transcripts and recordings"),
			_("{0} Event(s) have a fetched Teams transcript"),
		),
	)
	meta = frappe.get_meta("Microsoft Settings")
	found = []
	for field, column, label, evidence in checks:
		if not meta.has_field(field) or frappe.db.get_single_value("Microsoft Settings", field):
			continue
		if not frappe.db.has_column("Event", column):
			continue
		count = frappe.db.count("Event", {column: ["is", "set"]})
		if not count:
			continue
		found.append(
			doctor.finding(
				"upgrade.meeting_capability_off",
				doctor.WARN,
				_("{0} is switched off, but this site has used it").format(label),
				_(
					"{0}, but the tickbox is off, so nothing asks Microsoft for it any more. Sites "
					"upgraded from before 1.0.0 lost it when the old scope list was cleared."
				).format(evidence.format(count)),
				_(
					"If you still want it, tick {0} in Microsoft Settings (Meeting transcripts needs "
					"Standalone Teams meetings too), save, and have each user authorise their "
					"Microsoft Calendar again."
				).format(label),
				target="Microsoft Settings",
			)
		)
	return found
