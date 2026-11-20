"""Seed a site running the PREVIOUS release with the data a real site has, before upgrading it.

Run from the bench's ``sites`` directory with the bench's Python, while the app is checked out at
the last release tag::

	../env/bin/python <repo>/.github/upgrade/seed.py <site> <manifest.json>

This script is not part of the app: it runs against whatever version is installed, so it reads
the installed DocTypes' fields and only sets those that exist, and seeds SharePoint storage only
when that release has it (1.1.0 onwards, Frappe 16 only). It never calls Microsoft: records are
written the way the app itself writes pulled data, with the sync hooks switched off.

The manifest it writes is what ``assert_upgrade.py`` checks after ``bench migrate``.
"""

import base64
import json
import sys

import frappe
from frappe.utils import add_days, getdate, now_datetime

USERS = ("alice@example.com", "bob@example.com")
TENANT = "11111111-2222-3333-4444-555555555555"

GRAPH_DELTA = "https://graph.microsoft.com/v1.0/me/calendarView/delta?$deltatoken=seeded-token"
FOREIGN_DELTA = "http://169.254.169.254/latest/meta-data/?$deltatoken=x"


def only_existing(doctype, values):
	"""Keep the fields this release's DocType actually has."""
	meta = frappe.get_meta(doctype)
	standard = ("doctype", "name", "parent", "parenttype", "parentfield", "idx")
	return {k: v for k, v in values.items() if k in standard or meta.has_field(k)}


def fake_token(claims):
	"""An unsigned JWT-shaped token. Nothing verifies it; code that reads claims (tid) can."""

	def part(data):
		return base64.urlsafe_b64encode(json.dumps(data).encode()).decode().rstrip("=")

	return "{0}.{1}.".format(part({"alg": "none", "typ": "JWT"}), part(claims))


def seed_users():
	for email in USERS:
		if frappe.db.exists("User", email):
			continue
		user = frappe.get_doc(
			{
				"doctype": "User",
				"email": email,
				"first_name": email.split("@")[0].title(),
				"send_welcome_email": 0,
			}
		).insert(ignore_permissions=True)
		user.add_roles("System Manager" if email.startswith("alice") else "Desk User")


def seed_settings():
	# Straight to the database: saving the form would run provisioning and validation that
	# expects a real tenant. These are the values a configured site has.
	values = {
		"enabled": 1,
		# Multi-tenant authority with sign-in on: what the 1.1.1 tenant patch has to deal with.
		"tenant_id": "common",
		"client_id": "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
		"use_calendar": 1,
		"use_teams": 1,
		"use_sso": 1,
		"authorized_scopes": "User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite offline_access",
	}
	meta = frappe.get_meta("Microsoft Settings")
	for field, value in values.items():
		if meta.has_field(field):
			frappe.db.set_single_value("Microsoft Settings", field, value)
	from frappe.utils.password import set_encrypted_password

	set_encrypted_password("Microsoft Settings", "Microsoft Settings", "seeded-secret", "client_secret")


def seed_calendar(name, user, delta_link):
	if not frappe.db.exists("Microsoft Calendar", name):
		doc = frappe.get_doc(
			only_existing(
				"Microsoft Calendar",
				{
					"doctype": "Microsoft Calendar",
					"account_name": name,
					"user": user,
					"microsoft_user_email": user,
					"enabled": 1,
					"pull_from_microsoft_calendar": 1,
					"push_to_microsoft_calendar": 1,
					"fetch_artifacts_automatically": 1,
					"ms_calendar_id": "AAMk-" + name,
					"ms_calendar_name": "Calendar",
				},
			)
		)
		doc.insert(ignore_permissions=True)
	# What the OAuth callback writes once a user has signed in.
	from frappe.utils.password import set_encrypted_password

	set_encrypted_password(
		"Microsoft Calendar", name, fake_token({"tid": TENANT, "upn": user}), "access_token"
	)
	set_encrypted_password("Microsoft Calendar", name, "seeded-refresh-token", "refresh_token")
	frappe.db.set_value(
		"Microsoft Calendar",
		name,
		{
			"authorized": 1,
			"token_expiry": add_days(now_datetime(), 1),
			"last_sync": now_datetime(),
			"delta_link": delta_link,
			"delta_window_end": add_days(now_datetime(), 180),
		},
		update_modified=False,
	)


def seed_event(key, calendar, owner, **values):
	"""Insert an Event the way the sync does (hooks off), then give it its owner."""
	existing = frappe.db.get_value("Event", {"subject": key})
	if existing:
		return existing
	start = values.pop("starts_on")
	doc = {
		"doctype": "Event",
		"subject": key,
		"event_type": "Private",
		"starts_on": start,
		"custom_sync_with_microsoft_calendar": 1 if calendar else 0,
		"custom_microsoft_calendar": calendar,
		"custom_microsoft_event_id": ("AAMkEvt-" + key) if calendar else None,
		"custom_microsoft_web_link": ("https://outlook.office365.com/owa/?itemid=" + key)
		if calendar
		else None,
	}
	doc.update(values)
	frappe.flags.in_microsoft_sync = True
	try:
		event = frappe.get_doc(only_existing("Event", doc)).insert(ignore_permissions=True)
	finally:
		frappe.flags.in_microsoft_sync = False
	frappe.db.set_value("Event", event.name, "owner", owner, update_modified=False)
	return event.name


def seed_events():
	today = getdate()
	tomorrow = add_days(today, 1)
	names = {}
	names["timed"] = seed_event(
		"upgrade-seed timed",
		"Alice Calendar",
		USERS[0],
		starts_on=f"{tomorrow} 10:00:00",
		ends_on=f"{tomorrow} 11:00:00",
	)
	names["all_day"] = seed_event(
		"upgrade-seed all-day",
		"Alice Calendar",
		USERS[0],
		starts_on=f"{tomorrow} 00:00:00",
		ends_on=f"{tomorrow} 00:00:00",
		all_day=1,
		custom_pulled_from_microsoft=1,
	)
	names["repeating"] = seed_event(
		"upgrade-seed repeating",
		"Alice Calendar",
		USERS[0],
		starts_on=f"{tomorrow} 09:00:00",
		ends_on=f"{tomorrow} 09:30:00",
		repeat_this_event=1,
		repeat_on="Weekly",
		repeat_till=add_days(today, 90),
		monday=1,
		wednesday=1,
	)
	names["teams"] = seed_event(
		"upgrade-seed teams",
		"Alice Calendar",
		USERS[0],
		starts_on=f"{today} 08:00:00",
		ends_on=f"{today} 08:30:00",
		custom_add_teams_meeting=1,
		custom_teams_join_url="https://teams.microsoft.com/l/meetup-join/seeded",
		custom_microsoft_online_meeting_id="MSoxseeded",
	)
	# Owned by Bob but linked to Alice's calendar: what the 1.1.1 ownership check reports.
	names["other_owner"] = seed_event(
		"upgrade-seed other owner",
		"Alice Calendar",
		USERS[1],
		starts_on=f"{tomorrow} 14:00:00",
		ends_on=f"{tomorrow} 15:00:00",
	)
	names["unsynced"] = seed_event(
		"upgrade-seed unsynced",
		None,
		USERS[1],
		starts_on=f"{tomorrow} 16:00:00",
		ends_on=f"{tomorrow} 17:00:00",
	)
	return names


def seed_sharepoint():
	"""SharePoint storage, when the installed release has it (1.1.0 onwards, Frappe 16 only)."""
	if not frappe.get_meta("Microsoft Settings").has_field("use_files"):
		return None
	if not frappe.db.exists("DocType", "SharePoint Mapping"):
		return None
	if int(frappe.__version__.split(".")[0]) < 16:
		return None

	from frappe_microsoft365 import setup

	if hasattr(setup, "create_file_custom_fields"):
		setup.create_file_custom_fields()

	todos = []
	for i in range(2):
		todos.append(
			frappe.get_doc({"doctype": "ToDo", "description": f"upgrade-seed todo {i}"}).insert(
				ignore_permissions=True
			)
		)

	files = {}
	for status in ("Stored", "Failed", "Pending"):
		f = frappe.get_doc(
			{
				"doctype": "File",
				"file_name": f"upgrade-seed-{status.lower()}.txt",
				"content": f"seeded {status}".encode(),
				"is_private": 1,
				"attached_to_doctype": "ToDo",
				"attached_to_name": todos[0].name,
			}
		)
		# Inserted while storage is still off, so the file hooks queue nothing; the SharePoint
		# state below is what the move job leaves behind.
		f.insert(ignore_permissions=True)
		values = {"custom_microsoft_status": status, "custom_microsoft_attempts": 1}
		if status == "Stored":
			values.update(
				{
					"custom_microsoft_drive_id": "b!seeded-drive",
					"custom_microsoft_item_id": "01SEEDEDFILE",
					"custom_microsoft_local_url": f.file_url,
					"custom_microsoft_web_url": "https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/x.txt",
				}
			)
		if status == "Failed":
			values["custom_microsoft_error"] = "Seeded failure"
		frappe.db.set_value("File", f.name, values, update_modified=False)
		files[status] = f.name
	frappe.db.set_single_value("Microsoft Settings", "use_files", 1)
	frappe.db.set_single_value("Microsoft Settings", "app_client_id", "ffffffff-bbbb-cccc-dddd-eeeeeeeeeeee")

	mapping = frappe.get_doc(
		only_existing(
			"SharePoint Mapping",
			{
				"doctype": "SharePoint Mapping",
				"enabled": 1,
				"reference_doctype": "ToDo",
				"folder_creation": "Automatic",
				"site_url": "https://contoso.sharepoint.com/sites/Sales",
				"library": "Documents",
				"use_base_folder": 1,
				"base_folder": "Frappe/ToDo",
				"folder_pattern": "{name}",
				"keep_local_copy": "Default",
				"delete_remote": "Default",
				"files_panel": "Top of form",
				"site_id": "contoso.sharepoint.com,site-guid,web-guid",
				"drive_id": "b!seeded-drive",
			},
		)
	)
	# Validation would ask Graph for the site; this is what a saved mapping holds.
	mapping.name = "ToDo"
	mapping.db_insert()
	for i, todo in enumerate(todos):
		row = frappe.get_doc(
			only_existing(
				"SharePoint Folder",
				{
					"doctype": "SharePoint Folder",
					"parent": mapping.name,
					"parenttype": "SharePoint Mapping",
					"parentfield": "folders",
					"idx": i + 1,
					"reference_doctype": "ToDo",
					"reference_name": todo.name,
					"folder_name": todo.name,
					"web_url": f"https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Frappe/ToDo/{todo.name}",
					"drive_id": "b!seeded-drive",
					"item_id": f"01SEEDEDFOLDER{i}",
					"folder_path": f"Frappe/ToDo/{todo.name}",
				},
			)
		)
		row.db_insert()

	return {"mapping": mapping.name, "folders": len(todos), "files": files}


def main(site, manifest_path):
	frappe.init(site=site, sites_path=".")
	frappe.connect()
	frappe.set_user("Administrator")
	try:
		from frappe_microsoft365 import __version__

		seed_users()
		seed_settings()
		seed_calendar("Alice Calendar", USERS[0], GRAPH_DELTA)
		seed_calendar("Bob Calendar", USERS[1], FOREIGN_DELTA)
		events = seed_events()
		sharepoint = seed_sharepoint()
		frappe.db.commit()

		manifest = {
			"from_version": __version__,
			"frappe": frappe.__version__,
			"events": events,
			"calendars": {"Alice Calendar": GRAPH_DELTA, "Bob Calendar": FOREIGN_DELTA},
			"tenant": TENANT,
			"sharepoint": sharepoint,
			"counts": {
				"Event": frappe.db.count("Event", {"subject": ["like", "upgrade-seed %"]}),
				"Microsoft Calendar": frappe.db.count("Microsoft Calendar"),
			},
		}
		with open(manifest_path, "w") as f:
			json.dump(manifest, f, indent=1, default=str)
		print(json.dumps(manifest, indent=1, default=str))
	finally:
		frappe.destroy()


if __name__ == "__main__":
	main(sys.argv[1], sys.argv[2])
