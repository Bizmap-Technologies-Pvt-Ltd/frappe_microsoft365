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
	"""Everything this app needs present on a site. Idempotent, safe to re-run.

	What an optional capability needs (the File fields for SharePoint storage, say) is created
	when its tickbox is switched on, not here, so a site that never uses it never gets it. On
	migrate it is refreshed only for the capabilities that are on.
	"""
	from frappe_microsoft365 import capabilities

	create_event_custom_fields()
	capabilities.ensure_enabled_capabilities()


def create_event_custom_fields(only_definitions=False):
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
				# Two scheduled jobs filter tabEvent by this column every 15 minutes — the push
				# looking for events to send, the artifacts job for meetings whose files may
				# have landed — and EXPLAIN reported `type: ALL` for both: a full scan of every
				# Event on the site, twice a quarter hour, forever.
				#
				# This is the only narrow column either query has. Everything else they filter
				# on is a checkbox or a small integer, and an index over two or three distinct
				# values is worse than none — the optimiser ignores it while every Event write
				# still pays to maintain it. A site has a handful of Microsoft Calendars, so
				# narrowing on the link alone leaves few enough rows that the rest of each
				# filter costs nothing to apply.
				"search_index": 1,
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
	if only_definitions:
		return custom_fields
	create_custom_fields(custom_fields, ignore_validate=True)


FILE_STATUS_OPTIONS = "\nPending\nStored\nArchived\nFailed"


def create_file_custom_fields(only_definitions=False):
	"""Fields on File that record where an attachment lives in SharePoint (idempotent).

	Created when SharePoint document storage is switched on (capabilities.py), never before.

	Custom fields rather than a side table: every reader of File — permissions, copies made by
	create_attachment_copy, the File list — already carries them along, and a moved File is still
	just a File.
	"""
	custom_fields = {
		"File": [
			{
				"fieldname": "custom_microsoft_section",
				"fieldtype": "Section Break",
				"label": "Microsoft 365",
				"insert_after": "uploaded_to_google_drive"
				if frappe.db.has_column("File", "uploaded_to_google_drive")
				else "content_hash",
				"collapsible": 1,
				"depends_on": "eval:doc.custom_microsoft_status",
			},
			{
				"fieldname": "custom_microsoft_status",
				"fieldtype": "Select",
				"label": "SharePoint Status",
				"options": FILE_STATUS_OPTIONS,
				"insert_after": "custom_microsoft_section",
				"read_only": 1,
				"in_standard_filter": 1,
				"search_index": 1,
				"description": "<b>Pending:</b> waiting to move.<br><b>Stored:</b> moved to SharePoint and opened from there.<br><b>Archived:</b> a link whose target was copied to SharePoint.<br><b>Failed:</b> still on this server; see Last SharePoint Error. Retried hourly.",
			},
			{
				"fieldname": "custom_microsoft_web_url",
				"fieldtype": "Small Text",
				"label": "Open in SharePoint",
				"insert_after": "custom_microsoft_status",
				"read_only": 1,
				# A SharePoint web URL carries the site, library, folders and file name, percent-
				# encoded, and passes Frappe's 140-char default for any long file or customer
				# name. At 140 the write failed AFTER the upload, the job died before recording
				# the attempt, and the hourly retry uploaded the file again, every hour, as
				# "name (1).pdf", "name (2).pdf", ... Nothing queries it: width, no index.
				# Small Text, not a long Data: a varchar(1000) here beside the indexed
				# local_url varchar(768) took a busy tabFile (ROW_FORMAT=COMPRESSED, other
				# apps' columns and indexes) past InnoDB's 8126-byte row limit, and migrate
				# failed. TEXT has no width to outgrow and does not count against the limit.
			},
			{
				"fieldname": "custom_microsoft_error",
				"fieldtype": "Small Text",
				"label": "Last SharePoint Error",
				"insert_after": "custom_microsoft_web_url",
				"read_only": 1,
				"depends_on": "eval:doc.custom_microsoft_status == 'Failed'",
			},
			{
				"fieldname": "custom_microsoft_column",
				"fieldtype": "Column Break",
				"insert_after": "custom_microsoft_error",
			},
			{
				"fieldname": "custom_microsoft_drive_id",
				"fieldtype": "Data",
				"label": "Drive ID",
				"insert_after": "custom_microsoft_column",
				"read_only": 1,
				# Graph ids have no documented maximum; drive ids run to ~66 characters and item
				# ids to ~34. 255 leaves headroom without widening every File row for nothing.
				"length": 255,
			},
			{
				"fieldname": "custom_microsoft_item_id",
				"fieldtype": "Data",
				"label": "Item ID",
				"insert_after": "custom_microsoft_drive_id",
				"read_only": 1,
				"search_index": 1,
				"length": 255,
			},
			{
				"fieldname": "custom_microsoft_local_url",
				"fieldtype": "Data",
				"label": "Original URL",
				"insert_after": "custom_microsoft_item_id",
				"read_only": 1,
				"search_index": 1,
				# The local file URL before the move: as long as the file name. Indexed, because
				# every old link is looked up by it, so 768, the widest utf8mb4 column InnoDB can
				# still index.
				"length": 768,
				"description": "Where the file was on this server before it moved. Old links to it are redirected to SharePoint. With Keep a local copy on, the copy is still there and is served if the SharePoint copy disappears.",
			},
			{
				"fieldname": "custom_microsoft_attempts",
				"fieldtype": "Int",
				"label": "Upload Attempts",
				"insert_after": "custom_microsoft_local_url",
				"read_only": 1,
			},
		]
	}
	if only_definitions:
		return custom_fields
	create_custom_fields(custom_fields, ignore_validate=True)


def app_custom_fields():
	"""Every custom field this app adds, as ``{doctype: [fieldname, ...]}``. Used by uninstall."""
	from frappe_microsoft365.microsoft_files import FILES_TAB_FIELDS

	out = {}
	for definitions in (create_event_custom_fields(True), create_file_custom_fields(True)):
		for doctype, fields in definitions.items():
			out.setdefault(doctype, []).extend(f["fieldname"] for f in fields)
	# The Microsoft 365 Files tab, on whichever DocTypes a mapping put it.
	for row in frappe.get_all(
		"Custom Field", filters={"fieldname": ["in", FILES_TAB_FIELDS]}, fields=["dt", "fieldname"]
	):
		out.setdefault(row.dt, []).append(row.fieldname)
	return out
