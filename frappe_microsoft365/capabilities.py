"""The capabilities this app offers, defined once.

Microsoft Settings has one tickbox per capability. What a capability needs — the Azure
permissions, whether it acts as a person or as the app, what it depends on, which Frappe it
runs on, what has to exist on the site once it is on — used to be spelt out separately in the
scope derivation, Set Up, the doctor and the Settings form, and adding a capability meant
finding all of them. Everything now reads it from here.

Pure apart from is_enabled/apply_changes: the definitions are plain data, so every rule built
on them is testable without a site or a tenant.

Labels and notes are plain strings; callers translate them when they render them.
"""

from dataclasses import dataclass

import frappe
from frappe import _

#: Acts as whoever signed in (a Microsoft Calendar's delegated token).
DELEGATED = "delegated"
#: Acts as the Azure app itself (client credentials); nobody signs in for it.
APPLICATION = "application"
#: Frappe does the Microsoft side itself through a record this app creates (Connected App,
#: Social Login Key); the scopes live on that record, not in our sign-in.
FRAPPE_RECORD = "frappe_record"


@dataclass(frozen=True)
class Capability:
	id: str
	field: str
	label: str
	access: str
	#: Delegated Microsoft Graph scopes our own sign-in requests for this capability.
	graph_scopes: tuple = ()
	#: Permissions an admin grants in Azure for it, as shown in Set Up and the doctor.
	azure_permissions: tuple = ()
	azure_type: str = ""
	creates: str = "Nothing — this app talks to Graph directly."
	note: str = ""
	#: Capability ids that must be on for this one to work.
	requires: tuple = ()
	#: Lowest Frappe major version it runs on.
	min_frappe: int = 15
	#: Dotted path, called once when the capability is switched on (and on migrate while it is
	#: on). Idempotent. This is where a capability puts what it needs on the site — so a site
	#: that never switches it on never gets it.
	on_enable: str = ""


CAPABILITIES = (
	Capability(
		id="calendar",
		field="use_calendar",
		label="Outlook calendar",
		access=DELEGATED,
		graph_scopes=("Calendars.ReadWrite",),
		azure_type="Microsoft Graph, Delegated",
		note=(
			"Covers the Add Teams meeting tickbox on an Event too: Microsoft creates the "
			"Teams link as part of the event, so no meeting permission is involved. Each "
			"user authorises their own Microsoft Calendar."
		),
	),
	Capability(
		id="teams",
		field="use_teams",
		label="Standalone Teams meetings",
		access=DELEGATED,
		graph_scopes=("OnlineMeetings.ReadWrite",),
		azure_type="Microsoft Graph, Delegated",
		note=(
			"Only for meetings created outside a calendar event, and for resolving a join "
			"URL back to a meeting. Tenants commonly refuse this one; leave it off unless "
			"something asks for it."
		),
	),
	Capability(
		id="transcripts",
		field="use_transcripts",
		label="Meeting transcripts and recordings",
		access=DELEGATED,
		graph_scopes=("OnlineMeetingTranscript.Read.All", "OnlineMeetingRecording.Read.All"),
		azure_type="Microsoft Graph, Delegated",
		note=(
			"Transcripts are reached through the meeting behind a join URL, so this builds on "
			"standalone Teams meetings. The meeting must also be calendar-associated and not "
			"expired, or Graph returns nothing."
		),
		requires=("teams",),
	),
	Capability(
		id="mail",
		field="use_mail",
		label="Outlook mail",
		access=FRAPPE_RECORD,
		azure_permissions=(
			"https://outlook.office.com/IMAP.AccessAsUser.All",
			"https://outlook.office.com/SMTP.Send",
			"offline_access",
		),
		azure_type="Office 365 Exchange Online, Delegated",
		creates="A Connected App named '{mail_app}' for Frappe's Email Account.",
		note=(
			"Mail itself stays with Frappe's Email Account. For shared mailboxes choose the "
			"Application method, which also needs IMAP.AccessAsApp / SMTP.SendAsApp plus the "
			"Exchange setup script."
		),
	),
	Capability(
		id="sso",
		field="use_sso",
		label="Sign in with Microsoft",
		access=FRAPPE_RECORD,
		azure_permissions=("openid", "email", "profile"),
		azure_type="Microsoft Graph, Delegated",
		creates="A Social Login Key named '{sso_key}'.",
		note=(
			"New visitors cannot self-register by default; existing users with a matching "
			"email can sign in."
		),
	),
	Capability(
		id="files",
		field="use_files",
		label="SharePoint document storage",
		access=APPLICATION,
		azure_permissions=("Sites.Selected",),
		azure_type="Microsoft Graph, Application",
		creates=(
			"Fields on File that record where each attachment lives. Then add a SharePoint "
			"Mapping for each DocType whose attachments should go to SharePoint."
		),
		note=(
			"Runs as the app itself, never as a signed-in person, so files keep working when "
			"someone leaves. Sites.Selected reaches no site until a SharePoint admin grants the "
			"app each site (SharePoint Settings > Site Grant Script)."
		),
		min_frappe=16,
		on_enable="frappe_microsoft365.setup.create_file_custom_fields",
	),
)


def all_capabilities():
	return CAPABILITIES


def get(capability_id):
	for capability in CAPABILITIES:
		if capability.id == capability_id:
			return capability
	raise KeyError(capability_id)


def by_field(fieldname):
	for capability in CAPABILITIES:
		if capability.field == fieldname:
			return capability
	return None


def ticked(settings):
	"""The capabilities ticked in ``settings`` (any .get() mapping), in registry order."""
	return [c for c in CAPABILITIES if (settings or {}).get(c.field)]


def delegated_graph_scopes():
	"""``((field, scopes), ...)`` for every capability our own sign-in requests scopes for."""
	return tuple((c.field, c.graph_scopes) for c in CAPABILITIES if c.access == DELEGATED and c.graph_scopes)


def frappe_major_version():
	try:
		return int(str(frappe.__version__).split(".")[0])
	except (TypeError, ValueError):
		return 0


def unmet(capability, settings, frappe_major=None):
	"""Why ``capability`` cannot be on with these settings, or None. Pure given frappe_major."""
	frappe_major = frappe_major_version() if frappe_major is None else frappe_major
	if frappe_major and frappe_major < capability.min_frappe:
		return _("{0} needs Frappe v{1} or later; this site runs v{2}.").format(
			_(capability.label), capability.min_frappe, frappe_major
		)
	missing = [get(r) for r in capability.requires if not (settings or {}).get(get(r).field)]
	if missing:
		return _("{0} builds on {1}, which is off.").format(
			_(capability.label), ", ".join(_(m.label) for m in missing)
		)
	return None


def is_enabled(capability_id, settings=None):
	"""Is this capability live: the integration enabled and its tickbox on.

	Every hook a capability registers calls this first, so a site that has it off pays one
	cached read and nothing else.
	"""
	settings = settings or frappe.get_cached_doc("Microsoft Settings")
	return bool(settings.get("enabled") and settings.get(get(capability_id).field))


def run_on_enable(capability):
	if capability.on_enable:
		frappe.get_attr(capability.on_enable)()


def apply_changes(before, after):
	"""Run on_enable for every capability that ``after`` has on and ``before`` did not.

	Switching a capability off removes nothing: data it put on the site (moved files, folder
	records) still has to keep working. Removal is uninstall's job.
	"""
	for capability in ticked(after):
		if not (before or {}).get(capability.field):
			run_on_enable(capability)


def ensure_enabled_capabilities(settings=None):
	"""On migrate: re-run on_enable for what is on, so its definitions stay current."""
	if not frappe.db.exists("DocType", "Microsoft Settings"):
		return
	settings = settings or frappe.get_single("Microsoft Settings")
	for capability in ticked(settings):
		run_on_enable(capability)
