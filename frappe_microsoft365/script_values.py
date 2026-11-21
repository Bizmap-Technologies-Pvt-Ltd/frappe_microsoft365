"""Values that go into the PowerShell scripts this app generates.

The Site Grant and Exchange scripts are run by somebody else — the tenant's SharePoint or
Exchange administrator — with rights far beyond anything on this site. Every value written into
them comes from a form a System Manager filled in, so a value must never be able to become code
in that administrator's window.

Two layers, because either alone has gaps. Values are checked against the shape Microsoft gives
them (a client id is a GUID, a mailbox an address, a site a *.sharepoint.com URL), both when they
are saved and again when a script is generated. And whatever passes is written as a single-quoted
PowerShell string, the only kind PowerShell never expands.
"""

import re
from urllib.parse import urlsplit

import frappe
from frappe import _
from frappe.utils import escape_html

GUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.IGNORECASE)

#: An Exchange address. Deliberately narrower than RFC 5322, which allows quoted local parts and
#: characters no Exchange mailbox has. The apostrophe stays: O'Brien is a real mailbox, and the
#: quoting below makes it harmless.
MAILBOX = re.compile(
	r"^[A-Za-z0-9._%+'-]+@[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?"
	r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?)+$"
)

#: SharePoint Online lives under <tenant>.sharepoint.com, and <tenant>-my.sharepoint.com for
#: OneDrive; the app reaches nothing else.
SHAREPOINT_HOST = re.compile(r"^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.sharepoint\.com$")

#: A site's name as it appears in /sites/<name>. SharePoint itself allows only letters, digits,
#: hyphens, underscores and dots in a new site address; %-escapes cover older sites with spaces.
SITE_NAME = re.compile(r"^[A-Za-z0-9._~%-]+$")

#: PowerShell ends a single-quoted string at any of these, not only at the ASCII apostrophe, and
#: inside one each of them is escaped by doubling it.
_SINGLE_QUOTES = "'\u2018\u2019\u201a\u201b"


def ps_literal(value):
	"""``value`` as a single-quoted PowerShell string: literal text, never expanded.

	Inside single quotes PowerShell evaluates nothing — no ``$variable``, no ``$(...)``, no
	backtick escape — and the only special character is the quote itself, written twice.
	"""
	text = str(value or "")
	for quote in _SINGLE_QUOTES:
		text = text.replace(quote, quote * 2)
	return f"'{text}'"


def is_client_id(value):
	return bool(GUID.match((value or "").strip()))


def is_mailbox(value):
	return bool(MAILBOX.match((value or "").strip()))


def validate_client_id(value, label=None):
	"""Throw unless ``value`` is an Azure Application (client) ID, which is always a GUID."""
	value = (value or "").strip()
	if not is_client_id(value):
		frappe.throw(
			_(
				"{0} must be the Application (client) ID from the app registration's Overview page, a GUID such as 00000000-0000-0000-0000-000000000000."
			).format(label or _("Client ID")),
			title=_("Not a client ID"),
		)
	return value


def validate_mailbox(value):
	"""Throw unless ``value`` is a plain mailbox address."""
	value = (value or "").strip()
	if not is_mailbox(value):
		frappe.throw(
			_(
				"{0} is not a mailbox address. Enter addresses such as support@contoso.com, one per line."
			).format(frappe.bold(escape_html(value))),
			title=_("Not a mailbox"),
		)
	return value


def site_url_problem(url):
	"""Why ``url`` cannot be a SharePoint Online site, or None when it can. Pure.

	Checks the part the scripts and Graph calls use — the host and the /sites/ or /teams/ name —
	and nothing after it, since people paste the address of a library or a file as often as the
	site itself.
	"""
	raw = (url or "").strip()
	if not raw:
		return _("SharePoint Site URL is required.")
	if any(c.isspace() or ord(c) < 32 for c in raw):
		return _("The address contains spaces or line breaks.")
	if "://" not in raw:
		raw = f"https://{raw}"
	parsed = urlsplit(raw)
	if parsed.scheme.lower() != "https":
		return _("The address must start with https://.")
	host = (parsed.hostname or "").lower()
	if not SHAREPOINT_HOST.match(host) or parsed.port or parsed.username or parsed.password:
		return _("The site must be on <tenant>.sharepoint.com.")
	match = re.match(r"^/(sites|teams)/([^/]+)", parsed.path or "", re.IGNORECASE)
	if match and not SITE_NAME.match(match.group(2)):
		return _("The site name {0} contains characters a SharePoint site address cannot have.").format(
			match.group(2)
		)
	return None


def validate_site_url(url):
	"""Throw unless ``url`` belongs to a SharePoint Online site; return ``(host, path)``."""
	from frappe_microsoft365.microsoft_files import parse_site_url

	problem = site_url_problem(url)
	if problem:
		frappe.throw(
			_("{0} is not a SharePoint Online site address. {1}").format(
				frappe.bold(escape_html((url or "").strip())), problem
			),
			title=_("Not a SharePoint site"),
		)
	return parse_site_url(url)


# --- upgrade report -------------------------------------------------------------------------


def upgrade_check():
	"""Saved values that the stricter validation in 1.1.1 would now refuse. Read-only.

	Nothing stops working on upgrade: these values are only checked when the record is next
	saved or a script is generated. Listing them now means the next person to save the form is
	not the one who finds out.
	"""
	from frappe_microsoft365.doctor import WARN, finding

	out = []
	# Read field by field from tabSingles: a site that has not migrated app_client_id yet must
	# still get its report rather than an error.
	meta = frappe.get_meta("Microsoft Settings")
	values = {
		field: frappe.db.get_single_value("Microsoft Settings", field)
		for field in ("client_id", "app_client_id")
		if meta.has_field(field)
	}
	for field, label in (("client_id", _("Client ID")), ("app_client_id", _("Application Client ID"))):
		value = (values.get(field) or "").strip()
		if value and not is_client_id(value):
			out.append(
				finding(
					"upgrade.script_values",
					WARN,
					_("Microsoft Settings: {0} is not a GUID").format(label),
					_(
						"Found {0}. Saving Microsoft Settings, or generating a setup script, now requires the Application (client) ID, which is always a GUID."
					).format(value),
					_(
						"Copy Application (client) ID from the app registration's Overview page into {0}."
					).format(label),
					target="Microsoft Settings",
				)
			)

	if frappe.db.table_exists("SharePoint Mapping"):
		for row in frappe.get_all("SharePoint Mapping", fields=["name", "site_url"]):
			problem = site_url_problem(row.site_url) if row.site_url else None
			if problem:
				out.append(
					finding(
						"upgrade.script_values",
						WARN,
						_("SharePoint Mapping {0}: the site address will be refused").format(row.name),
						_(
							"{0} {1} The mapping cannot be saved again, and no Site Grant Script is generated, until it is corrected."
						).format(row.site_url, problem),
						_(
							"Open the mapping and paste the site's address, e.g. https://contoso.sharepoint.com/sites/Sales."
						),
						target=row.name,
					)
				)
	return out
