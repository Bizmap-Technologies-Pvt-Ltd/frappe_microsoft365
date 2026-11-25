"""After ``bench migrate`` from the previous release: did the upgrade keep the data, run the
patches and produce the upgrade report?

Run from the bench's ``sites`` directory with the bench's Python, with the app checked out at the
code under test::

	../env/bin/python <repo>/.github/upgrade/assert_upgrade.py <site> <manifest.json>

``manifest.json`` is what ``seed.py`` wrote before the upgrade. Exits non-zero with every failure
listed, so one run shows everything that is wrong.

Patch expectations are keyed by a fragment of the patch's dotted path in ``patches.txt``; a
release that adds a data patch adds its expectation to ``PATCH_EXPECTATIONS`` below.
"""

import json
import sys

import frappe
from frappe.utils.password import get_decrypted_password

FAILURES = []


def check(condition, message):
	print(("  ok    " if condition else "  FAIL  ") + message)
	if not condition:
		FAILURES.append(message)


def app_patches():
	"""The dotted paths in the app's patches.txt, in order."""
	path = frappe.get_app_path("frappe_microsoft365", "patches.txt")
	with open(path) as f:
		lines = [line.split("#", 1)[0].strip() for line in f]
	return [line for line in lines if line and not line.startswith("[")]


# --- data the upgrade must keep -------------------------------------------------------


def assert_data_kept(manifest):
	print("Data kept:")
	counts = manifest["counts"]
	check(
		frappe.db.count("Event", {"subject": ["like", "upgrade-seed %"]}) == counts["Event"],
		"every seeded Event is still there",
	)
	check(
		frappe.db.count("Microsoft Calendar") == counts["Microsoft Calendar"],
		"every Microsoft Calendar is still there",
	)

	events = manifest["events"]
	for key, name in events.items():
		check(frappe.db.exists("Event", name), f"Event {key} ({name}) exists")
	for key in ("timed", "all_day", "repeating", "teams", "other_owner"):
		event_id = frappe.db.get_value("Event", events[key], "custom_microsoft_event_id")
		check(bool(event_id), f"Event {key} keeps its Microsoft event id")
	check(
		frappe.db.get_value("Event", events["all_day"], "all_day") == 1, "the all-day Event is still all-day"
	)
	check(
		frappe.db.get_value("Event", events["repeating"], "repeat_this_event") == 1,
		"the repeating Event still repeats",
	)
	check(
		bool(frappe.db.get_value("Event", events["teams"], "custom_teams_join_url")),
		"the Teams Event keeps its join URL",
	)

	for name in manifest["calendars"]:
		check(
			bool(get_decrypted_password("Microsoft Calendar", name, "refresh_token", raise_exception=False)),
			f"{name}'s refresh token still decrypts",
		)
		check(
			frappe.db.get_value("Microsoft Calendar", name, "authorized") == 1, f"{name} is still authorised"
		)

	sharepoint = manifest.get("sharepoint")
	if not sharepoint:
		print("  (no SharePoint data: the previous release or this Frappe version has none)")
		return
	check(
		frappe.db.exists("SharePoint Mapping", sharepoint["mapping"]), "the SharePoint Mapping is still there"
	)
	check(
		frappe.db.count("SharePoint Folder", {"parent": sharepoint["mapping"]}) == sharepoint["folders"],
		"every SharePoint Folder row is still there",
	)
	for status, name in sharepoint["files"].items():
		check(
			frappe.db.get_value("File", name, "custom_microsoft_status") == status,
			f"the {status} File keeps its SharePoint status",
		)
	check(
		bool(frappe.db.get_value("File", sharepoint["files"]["Stored"], "custom_microsoft_item_id")),
		"the Stored File keeps its SharePoint item id",
	)


# --- patches --------------------------------------------------------------------------


def expect_foreign_delta_link_cleared(manifest):
	"""1.1.1: a stored delta_link that is not an https Graph URL is cleared; a real one is kept."""
	for name, seeded in manifest["calendars"].items():
		now = frappe.db.get_value("Microsoft Calendar", name, "delta_link")
		if seeded.startswith("https://graph.microsoft.com/"):
			check(now == seeded, f"{name}: the Graph delta_link is kept")
		else:
			check(not now, f"{name}: the non-Graph delta_link is cleared")


def expect_tenant_filled(manifest):
	"""1.1.1: Tenant ID `common` is replaced by the tid every authorised calendar token agrees on."""
	check(
		frappe.db.get_single_value("Microsoft Settings", "tenant_id") == manifest["tenant"],
		"Tenant ID was filled from the calendar tokens' tid",
	)
	key = frappe.db.get_value(
		"Social Login Key", {"social_login_provider": "Office 365"}, ["authorize_url", "access_token_url"], as_dict=True
	)
	check(bool(key), "the seeded Microsoft sign-in key still exists")
	if key:
		for field in ("authorize_url", "access_token_url"):
			check(
				f"/{manifest['tenant']}/" in (key.get(field) or "") and "/common/" not in key.get(field),
				f"the sign-in key's {field} now names the tenant",
			)


#: patches.txt fragment -> what the patch must have done to the seeded data.
PATCH_EXPECTATIONS = {
	"delta_link": expect_foreign_delta_link_cleared,
	"tenant": expect_tenant_filled,
}


def assert_patches(manifest):
	print("Patches:")
	patches = app_patches()
	for patch in patches:
		check(frappe.db.exists("Patch Log", {"patch": patch}), f"{patch} is recorded as run")
	before = set(manifest.get("patches_before") or [])
	for fragment, expectation in PATCH_EXPECTATIONS.items():
		matching = [patch for patch in patches if fragment in patch]
		if not matching:
			print(f"  (no patch matching '{fragment}' in this release)")
		elif all(patch in before for patch in matching):
			# Recorded as run when the previous release was installed, before the data existed.
			print(f"  ({', '.join(matching)} was already applied by the previous release)")
		else:
			expectation(manifest)


# --- upgrade report -------------------------------------------------------------------


def assert_upgrade_report(manifest):
	print("Upgrade report:")
	from frappe_microsoft365 import __version__, doctor, upgrade

	check(
		upgrade.reported_version() == __version__,
		f"the report was produced for {__version__} (recorded: {upgrade.reported_version()})",
	)
	report = upgrade.get_report() or {"findings": []}
	for f in report["findings"]:
		print(f"        [{f['status']}] {f.get('check')}: {f.get('title')} {f.get('target') or ''}")
	check(
		not [f for f in report["findings"] if f.get("check") == "upgrade.check_failed"],
		"no upgrade check failed during migrate",
	)
	if report["findings"]:
		subject = ["like", f"Microsoft 365 {report['version']}:%"]
		for user in upgrade.system_managers():
			count = frappe.db.count("Notification Log", {"for_user": user, "subject": subject})
			check(count == 1, f"{user} got exactly one notification for {report['version']} (got {count})")

	# Every registered check, of every release, run against the seeded data. Checks for a
	# release the code has not been bumped to yet are not run by migrate, so this is what proves
	# them before the release commit.
	registered = frappe.get_hooks(upgrade.HOOK) or {}
	for version in sorted(registered, key=upgrade.version_key):
		for path in registered[version]:
			try:
				result = frappe.get_attr(path)()
			except Exception as e:
				check(False, f"{version} {path} raised {e!r}")
				continue
			valid = isinstance(result, list) and all(
				isinstance(f, dict)
				and f.get("status") in (doctor.PASS, doctor.WARN, doctor.FAIL, doctor.SKIP)
				for f in result
			)
			check(valid, f"{version} {path} returns doctor findings ({len(result or [])})")
			for f in result if valid else []:
				print(f"        [{f['status']}] {f.get('check')}: {f.get('title')}")
	frappe.db.rollback()


def assert_tables_dynamic(manifest):
	"""Seeded compressed, tabFile and tabEvent are DYNAMIC after migrate and take an ALTER (#68)."""
	print("Row format:")
	for doctype, alter_failed in (manifest.get("compressed") or {}).items():
		table = f"tab{doctype}"
		fmt = frappe.db.sql(
			"""select row_format from information_schema.tables
			where table_schema = database() and table_name = %s""",
			table,
		)[0][0]
		check(fmt == "Dynamic", f"{table} is Dynamic (was seeded Compressed, ALTER failed then: {alter_failed})")
		frappe.db.sql_ddl("set session innodb_strict_mode = 1")
		try:
			frappe.db.sql_ddl(f"alter table `{table}` add column upgrade_assert_probe int")
			frappe.db.sql_ddl(f"alter table `{table}` drop column upgrade_assert_probe")
			check(True, f"{table} takes an ALTER")
		except Exception as e:
			check(False, f"{table} takes an ALTER ({e})")


def main(site, manifest_path):
	with open(manifest_path) as f:
		manifest = json.load(f)
	frappe.init(site=site, sites_path=".")
	frappe.connect()
	frappe.set_user("Administrator")
	try:
		from frappe_microsoft365 import __version__

		print(f"Upgraded {manifest['from_version']} -> {__version__} on Frappe {frappe.__version__}")
		assert_data_kept(manifest)
		assert_patches(manifest)
		assert_tables_dynamic(manifest)
		assert_upgrade_report(manifest)
	finally:
		frappe.destroy()

	if FAILURES:
		print(f"\n{len(FAILURES)} upgrade assertion(s) failed:")
		for message in FAILURES:
			print("  - " + message)
		sys.exit(1)
	print("\nUpgrade assertions passed.")


if __name__ == "__main__":
	main(sys.argv[1], sys.argv[2])
