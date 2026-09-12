"""Idempotent post-migrate setup for frappe_microsoft365.

Creates the custom fields on the standard Frappe ``Event`` doctype that let an Event be
mirrored to/from a Microsoft (Outlook) calendar. Safe to run on every migrate.
"""

import frappe
from frappe.custom.doctype.custom_field.custom_field import create_custom_fields


def after_install():
	"""Frappe fires after_install (not after_migrate) on `bench install-app`.

	Wiring only after_migrate meant a plain install left the Event custom fields missing,
	and the sync then failed with "Unknown column custom_sync_with_microsoft_calendar".
	"""
	setup()


def after_migrate():
	"""Runs on every `bench migrate`, so upgrades pick up new fields too."""
	setup()


def setup():
	"""Everything this app needs present on a site. Idempotent, safe to re-run."""
	create_event_custom_fields()


def create_event_custom_fields():
	"""Add Microsoft-sync custom fields to the Event doctype (idempotent)."""
	custom_fields = {
		"Event": [
			{
				"fieldname": "microsoft_calendar_section",
				"fieldtype": "Section Break",
				"label": "Microsoft Calendar",
				"insert_after": "sync_with_google_calendar"
				if frappe.db.has_column("Event", "sync_with_google_calendar")
				else "description",
				"collapsible": 1,
			},
			{
				"fieldname": "custom_sync_with_microsoft_calendar",
				"fieldtype": "Check",
				"label": "Sync with Microsoft Calendar",
				"insert_after": "microsoft_calendar_section",
			},
			{
				"fieldname": "custom_microsoft_calendar",
				"fieldtype": "Link",
				"label": "Microsoft Calendar",
				"options": "Microsoft Calendar",
				"insert_after": "custom_sync_with_microsoft_calendar",
				"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
			},
			{
				"fieldname": "custom_microsoft_calendar_column",
				"fieldtype": "Column Break",
				"insert_after": "custom_microsoft_calendar",
			},
			{
				"fieldname": "custom_microsoft_event_id",
				"fieldtype": "Data",
				"label": "Microsoft Event ID",
				"insert_after": "custom_microsoft_calendar_column",
				"read_only": 1,
				"no_copy": 1,
				# Graph event ids run to ~150 characters and occurrence ids from
				# calendarView/delta are longer again. Frappe's Data default is varchar(140),
				# so the write failed outright with "Data too long for column" AFTER the event
				# had already been created in Outlook.
				#
				# Microsoft documents no maximum for `id`, so any fixed width is a judgement.
				# Small Text would remove the ceiling but cannot be indexed, and this column is
				# looked up once per pulled event: unindexed that is a full scan of tabEvent
				# per event, 50 per delta batch, every 15 minutes. 768 is the widest a utf8mb4
				# column can be and still carry an index (InnoDB's 3072-byte key limit), so it
				# buys 5x headroom over observed ids without giving up the lookup.
				"length": 768,
				"search_index": 1,
			},
			{
				"fieldname": "custom_pulled_from_microsoft",
				"fieldtype": "Check",
				"label": "Pulled From Microsoft",
				"insert_after": "custom_microsoft_event_id",
				"read_only": 1,
				"hidden": 1,
				"no_copy": 1,
			},
		]
	}
	create_custom_fields(custom_fields, ignore_validate=True)
