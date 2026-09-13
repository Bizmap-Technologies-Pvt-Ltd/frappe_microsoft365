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
				# Sync is meaningless without knowing which connection to sync through, and
				# without this the Event saves happily and then does nothing at all.
				"mandatory_depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
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
			{
				"fieldname": "custom_add_teams_meeting",
				"fieldtype": "Check",
				"label": "Add Teams meeting",
				"insert_after": "custom_sync_with_microsoft_calendar",
				"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
				# Microsoft cannot turn an existing online meeting back into a plain event, so
				# once the meeting is real this control would be a lie. Lock it then rather
				# than leave a tickbox that quietly does nothing.
				"read_only_depends_on": "eval:doc.custom_teams_join_url",
				"description": (
					"Creates the event in Outlook as a Teams meeting. Locked once it exists: "
					"Microsoft cannot undo it."
				),
			},
			{
				"fieldname": "custom_teams_join_url",
				"fieldtype": "Data",
				"options": "URL",
				"label": "Join Meeting",
				"insert_after": "custom_pulled_from_microsoft",
				"read_only": 1,
				"no_copy": 1,
				# Join URLs run past the 140-char default. Nothing queries this column, so it
				# needs width but no index.
				"length": 1000,
				"description": "Use the Join Meeting button above.",
							"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
			},
			{
				"fieldname": "custom_microsoft_web_link",
				"fieldtype": "Small Text",
				"label": "Open in Outlook",
				"insert_after": "custom_teams_join_url",
				"read_only": 1,
				"no_copy": 1,
				# Small Text rather than Data: webLink can be very long and, unlike the event
				# id, is never looked up, so there is no index to preserve.
				"description": "Also available as a button above.",
							"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
			},
			{
				"fieldname": "custom_microsoft_organizer",
				"fieldtype": "Data",
				"options": "Email",
				"label": "Organizer",
				"insert_after": "custom_microsoft_web_link",
				"read_only": 1,
				"no_copy": 1,
				# RFC 5321 caps an address at 254 characters, which is past Frappe's
				# varchar(140) default. Nothing queries this column, so the width is free.
				"length": 254,
				"description": "The Microsoft account that organized this event.",
							"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
			},
			{
				"fieldname": "custom_microsoft_my_response",
				"fieldtype": "Select",
				"label": "My Response",
				"insert_after": "custom_microsoft_organizer",
				"read_only": 1,
				"no_copy": 1,
				# Graph's responseStatus.response values verbatim, so a stored value can be
				# compared with a Graph payload without a translation table in between. The
				# leading blank covers an event Microsoft has said nothing about yet.
				"options": "\nnone\norganizer\ntentativelyAccepted\naccepted\ndeclined\nnotResponded",
				"description": "Your reply to this invitation, as Microsoft has it.",
							"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
			},
			{
				"fieldname": "custom_microsoft_online_meeting_id",
				"fieldtype": "Data",
				"label": "Online Meeting ID",
				"insert_after": "custom_microsoft_web_link",
				"read_only": 1,
				"hidden": 1,
				"no_copy": 1,
				# Resolved from the join URL once and kept: transcripts and recordings are
				# addressed by meeting id, and resolving costs a Graph call every time.
				"length": 500,
			},
			{
				"fieldname": "custom_microsoft_transcript_fetched_on",
				"fieldtype": "Datetime",
				"label": "Transcript Fetched On",
				"insert_after": "custom_microsoft_online_meeting_id",
				"read_only": 1,
				"no_copy": 1,
				"depends_on": "eval:doc.custom_microsoft_transcript_fetched_on",
				"description": "The transcript is attached to this event.",
			},
			{
				"fieldname": "custom_microsoft_recordings",
				"fieldtype": "Small Text",
				"label": "Recordings",
				"insert_after": "custom_microsoft_transcript_fetched_on",
				"read_only": 1,
				"no_copy": 1,
				"depends_on": "eval:doc.custom_microsoft_recordings",
				"description": "Held by Microsoft, not copied here. Use Download Recording.",
			},
			{
				# What the buttons act on: the recording ids, so offering part 2 of a long
				# meeting costs no Graph call. The field above is the same thing for humans.
				"fieldname": "custom_microsoft_recordings_data",
				"fieldtype": "Long Text",
				"label": "Recordings Data",
				"insert_after": "custom_microsoft_recordings",
				"read_only": 1,
				"hidden": 1,
				"no_copy": 1,
			},
			{
				"fieldname": "custom_microsoft_artifacts_status",
				"fieldtype": "Small Text",
				"label": "Transcript Status",
				"insert_after": "custom_microsoft_recordings_data",
				"read_only": 1,
				"no_copy": 1,
				"depends_on": "eval:doc.custom_microsoft_artifacts_status",
			},
			{
				# Where the backoff has got to. Attempts, not a timer: the schedule is derived
				# from the meeting's end, so a missed run never shifts the whole sequence.
				"fieldname": "custom_microsoft_artifacts_attempts",
				"fieldtype": "Int",
				"label": "Artifact Checks",
				"insert_after": "custom_microsoft_artifacts_status",
				"read_only": 1,
				"hidden": 1,
				"no_copy": 1,
			},
			{
				"fieldname": "custom_microsoft_artifacts_checked_on",
				"fieldtype": "Datetime",
				"label": "Artifacts Checked On",
				"insert_after": "custom_microsoft_artifacts_attempts",
				"read_only": 1,
				"hidden": 1,
				"no_copy": 1,
			},
			{
				"fieldname": "custom_microsoft_attendees",
				"fieldtype": "Small Text",
				"label": "Attendees",
				"insert_after": "custom_microsoft_my_response",
				"read_only": 1,
				"no_copy": 1,
				# Small Text rather than Data: one line per attendee has no useful ceiling, and
				# unlike custom_microsoft_event_id nothing ever queries this column, so there is
				# no index to keep inside InnoDB's 3072-byte key limit.
				"description": "Everyone invited, with their reply. Refreshed by each sync.",
							"depends_on": "eval:doc.custom_sync_with_microsoft_calendar",
			},
		]
	}
	create_custom_fields(custom_fields, ignore_validate=True)
