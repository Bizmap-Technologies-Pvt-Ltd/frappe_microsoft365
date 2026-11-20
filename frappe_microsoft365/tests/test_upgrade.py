"""The post-migrate upgrade report: checks per release, run once, shown in three places."""

import io
import json
from contextlib import redirect_stdout
from typing import ClassVar
from unittest.mock import patch

import frappe

from frappe_microsoft365 import doctor, setup, upgrade
from frappe_microsoft365.tests.base import BaseTestCase

HERE = "frappe_microsoft365.tests.test_upgrade"
SYSTEM_MANAGER = "upgrade-report-sm@example.com"

CALLS = []


# Fake checks, registered through the hook by the tests below.
def warn_check():
	CALLS.append("warn")
	return [
		doctor.finding(
			"upgrade.test_warn",
			doctor.WARN,
			"3 Events <will> stop syncing",
			"detail",
			"Reassign them",
			target="Calendar A",
		)
	]


def fail_check():
	CALLS.append("fail")
	return [doctor.finding("upgrade.test_fail", doctor.FAIL, "Shared folder", "", "Split it")]


def empty_check():
	CALLS.append("empty")
	return []


def pass_check():
	CALLS.append("pass")
	return [doctor.finding("upgrade.test_pass", doctor.PASS, "Fine")]


def later_check():
	CALLS.append("later")
	return [doctor.finding("upgrade.test_later", doctor.WARN, "From a later release")]


def raising_check():
	CALLS.append("raising")
	raise RuntimeError("boom")


def writing_check():
	# A check must only read; one that writes and then fails must leave nothing behind.
	CALLS.append("writing")
	frappe.get_doc({"doctype": "ToDo", "description": "upgrade-check-left-this"}).insert(
		ignore_permissions=True
	)
	raise RuntimeError("after writing")


def garbage_check():
	CALLS.append("garbage")
	return "not a list"


_get_hooks = frappe.get_hooks


def hooks(registered):
	"""Patch frappe.get_hooks so only the upgrade hook is faked; every other hook (doc_events
	run by the Notification Log insert, say) still comes from the installed apps."""

	def fake(hook=None, *args, **kwargs):
		if hook == upgrade.HOOK:
			return registered
		return _get_hooks(hook, *args, **kwargs)

	return patch("frappe.get_hooks", side_effect=fake)


def run_quietly(**kwargs):
	out = io.StringIO()
	with redirect_stdout(out):
		result = upgrade.run(**kwargs)
	return result, out.getvalue()


class TestVersions(BaseTestCase):
	def test_compares_numerically(self):
		self.assertGreater(upgrade.version_key("1.1.10"), upgrade.version_key("1.1.9"))
		self.assertGreater(upgrade.version_key("1.10.0"), upgrade.version_key("1.9.9"))
		self.assertGreater(upgrade.version_key("2.0.0"), upgrade.version_key("1.99.99"))

	def test_equal_forms(self):
		self.assertEqual(upgrade.version_key("1.1"), upgrade.version_key("1.1.0"))
		self.assertEqual(upgrade.version_key("v1.1.1"), upgrade.version_key("1.1.1"))
		self.assertEqual(upgrade.version_key("1.2.0-dev"), upgrade.version_key("1.2.0"))

	def test_nothing_is_older_than_everything(self):
		self.assertEqual(upgrade.version_key(None), ())
		self.assertLess(upgrade.version_key(""), upgrade.version_key("0.0.1"))


class UpgradeTestCase(BaseTestCase):
	@classmethod
	def setUpClass(cls):
		super().setUpClass()
		if not frappe.db.exists("User", SYSTEM_MANAGER):
			user = frappe.get_doc(
				{
					"doctype": "User",
					"email": SYSTEM_MANAGER,
					"first_name": "Upgrade Report",
					"send_welcome_email": 0,
				}
			).insert(ignore_permissions=True)
			user.add_roles("System Manager")

	def setUp(self):
		CALLS.clear()
		self._saved = {
			field: frappe.db.get_single_value(upgrade.SETTINGS, field, cache=False)
			for field in (upgrade.VERSION_FIELD, upgrade.REPORT_FIELD)
		}
		upgrade._save(upgrade.VERSION_FIELD, "")
		upgrade._save(upgrade.REPORT_FIELD, "")
		self._delete_notifications()

	def tearDown(self):
		for field, value in self._saved.items():
			upgrade._save(field, value or "")
		self._delete_notifications()
		frappe.db.commit()

	def _delete_notifications(self):
		frappe.db.delete("Notification Log", {"subject": ["like", "Microsoft 365 %"]})

	def notifications(self, user=SYSTEM_MANAGER):
		return frappe.get_all(
			"Notification Log",
			filters={"for_user": user, "subject": ["like", "Microsoft 365 %"]},
			fields=["subject", "email_content", "document_type", "document_name"],
		)


class TestCollecting(UpgradeTestCase):
	def test_selects_releases_after_the_last_reported_up_to_current(self):
		registered = {
			"1.1.0": [f"{HERE}.empty_check"],
			"1.1.10": [f"{HERE}.later_check"],
			"1.1.9": [f"{HERE}.warn_check", f"{HERE}.warn_check"],
			"1.2.0": [f"{HERE}.fail_check"],
		}
		with hooks(registered):
			selected = upgrade.registered_checks("1.1.0", "1.1.10")
		# Oldest first, numerically; the repeated path runs once; 1.1.0 and 1.2.0 are outside.
		self.assertEqual(selected, [("1.1.9", [f"{HERE}.warn_check"]), ("1.1.10", [f"{HERE}.later_check"])])

	def test_never_reported_means_every_release_up_to_current(self):
		registered = {"1.1.1": [f"{HERE}.warn_check"], "1.2.0": [f"{HERE}.fail_check"]}
		with hooks(registered):
			self.assertEqual(upgrade.registered_checks(None, "1.1.1"), [("1.1.1", [f"{HERE}.warn_check"])])

	def test_real_hook_resolves(self):
		# Whatever the app registers must import, so a typo fails here and not on a customer's migrate.
		for version, paths in (frappe.get_hooks(upgrade.HOOK) or {}).items():
			self.assertTrue(upgrade.version_key(version), version)
			for path in paths:
				self.assertTrue(callable(frappe.get_attr(path)), path)


class TestRun(UpgradeTestCase):
	registered: ClassVar[dict] = {
		"1.1.1": [f"{HERE}.warn_check", f"{HERE}.empty_check", f"{HERE}.pass_check"],
		"1.2.0": [f"{HERE}.later_check"],
	}

	def test_findings_are_printed_stored_and_notified(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks(self.registered):
			findings, out = run_quietly(current="1.1.1")

		self.assertEqual(CALLS, ["warn", "empty", "pass"])
		# PASS is not something to report.
		self.assertEqual([f["check"] for f in findings], ["upgrade.test_warn"])
		self.assertEqual(findings[0]["version"], "1.1.1")

		self.assertIn("Microsoft 365 1.1.1: 1 item(s) to check", out)
		self.assertIn("3 Events <will> stop syncing (Calendar A)", out)
		self.assertIn("Reassign them", out)
		self.assertIn("Run Diagnostics", out)

		report = upgrade.get_report()
		self.assertEqual(report["version"], "1.1.1")
		self.assertEqual(report["from_version"], "1.1.0")
		self.assertEqual([f["check"] for f in report["findings"]], ["upgrade.test_warn"])
		self.assertEqual(upgrade.reported_version(), "1.1.1")

		notes = self.notifications()
		self.assertEqual(len(notes), 1)
		self.assertIn("1.1.1", notes[0].subject)
		self.assertEqual(notes[0].document_type, "Microsoft Settings")
		# The finding's title is escaped in the notification body.
		self.assertIn("&lt;will&gt;", notes[0].email_content)
		self.assertNotIn("<will>", notes[0].email_content)
		self.assertTrue(self.notifications("Administrator"))

	def test_migrate_twice_reports_once(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks(self.registered):
			run_quietly(current="1.1.1")
			CALLS.clear()
			second, out = run_quietly(current="1.1.1")
		self.assertIsNone(second)
		self.assertEqual(CALLS, [])
		self.assertEqual(out, "")
		self.assertEqual(len(self.notifications()), 1)
		self.assertTrue(upgrade.get_report())

	def test_nothing_to_report_sends_nothing(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks({"1.1.1": [f"{HERE}.empty_check", f"{HERE}.pass_check"]}):
			findings, out = run_quietly(current="1.1.1")
		self.assertEqual(findings, [])
		self.assertIn("nothing to report", out)
		self.assertIsNone(upgrade.get_report())
		self.assertEqual(self.notifications(), [])
		self.assertEqual(upgrade.reported_version(), "1.1.1")

	def test_skipped_releases_all_run(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks(self.registered):
			findings, _out = run_quietly(current="1.2.0")
		self.assertEqual(CALLS, ["warn", "empty", "pass", "later"])
		self.assertEqual({f["version"] for f in findings}, {"1.1.1", "1.2.0"})
		self.assertEqual(len(self.notifications()), 1)

	def test_next_release_runs_only_its_own_checks_and_keeps_undismissed_findings(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks(self.registered):
			run_quietly(current="1.1.1")
			CALLS.clear()
			run_quietly(current="1.2.0")
		self.assertEqual(CALLS, ["later"])
		report = upgrade.get_report()
		self.assertEqual(report["version"], "1.2.0")
		self.assertEqual(report["from_version"], "1.1.0")
		self.assertEqual(
			[f["check"] for f in report["findings"]], ["upgrade.test_later", "upgrade.test_warn"]
		)
		self.assertEqual(len(self.notifications()), 2)

	def test_downgrade_or_same_version_runs_nothing(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.2.0")
		with hooks(self.registered):
			self.assertIsNone(run_quietly(current="1.1.1")[0])
		self.assertEqual(CALLS, [])

	def test_a_failing_check_is_reported_and_the_rest_still_run(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		registered = {
			"1.1.1": [
				f"{HERE}.raising_check",
				f"{HERE}.garbage_check",
				f"{HERE}.no_such_check",
				f"{HERE}.fail_check",
			]
		}
		errors_before = frappe.db.count("Error Log")
		with hooks(registered):
			findings, out = run_quietly(current="1.1.1")

		self.assertEqual(CALLS, ["raising", "garbage", "fail"])
		checks = [f["check"] for f in findings]
		self.assertEqual(checks.count("upgrade.check_failed"), 3)
		self.assertIn("upgrade.test_fail", checks)
		failed = [f for f in findings if f["check"] == "upgrade.check_failed"]
		self.assertTrue(all(f["status"] == doctor.FAIL for f in failed))
		self.assertIn("raising_check", failed[0]["detail"])
		self.assertIn("boom", failed[0]["detail"])
		self.assertEqual(frappe.db.count("Error Log"), errors_before + 3)
		self.assertIn("could not run", out)
		self.assertEqual(upgrade.reported_version(), "1.1.1")

	def test_a_failing_check_leaves_no_writes(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks({"1.1.1": [f"{HERE}.writing_check", f"{HERE}.warn_check"]}):
			findings, _out = run_quietly(current="1.1.1")
		self.assertFalse(frappe.db.exists("ToDo", {"description": "upgrade-check-left-this"}))
		self.assertEqual([f["check"] for f in findings], ["upgrade.check_failed", "upgrade.test_warn"])
		self.assertEqual(upgrade.reported_version(), "1.1.1")

	def test_after_migrate_never_raises(self):
		with patch.object(upgrade, "run", side_effect=RuntimeError("broken")):
			out = io.StringIO()
			with redirect_stdout(out):
				upgrade.after_migrate()
		self.assertIn("could not be produced", out.getvalue())

	def test_setup_after_migrate_runs_the_report_after_setup(self):
		order = []
		with (
			patch.object(setup, "setup", side_effect=lambda: order.append("setup")),
			patch.object(upgrade, "run", side_effect=lambda: order.append("report")),
		):
			setup.after_migrate()
		self.assertEqual(order, ["setup", "report"])

	def test_long_reports_are_capped(self):
		many = [
			doctor.finding(f"upgrade.n{i}", doctor.WARN, f"n{i}") for i in range(upgrade.MAX_FINDINGS + 5)
		]
		upgrade._store("1.1.1", "1.1.0", many)
		report = upgrade.get_report()
		self.assertEqual(len(report["findings"]), upgrade.MAX_FINDINGS)
		self.assertEqual(report["omitted"], 5)


class TestFreshInstall(UpgradeTestCase):
	def test_install_starts_the_record_at_the_installed_version(self):
		with patch.object(setup, "setup"):
			setup.after_install()
		self.assertEqual(upgrade.reported_version(), upgrade.app_version())

	def test_first_migrate_after_install_reports_nothing(self):
		upgrade.mark_installed("1.1.1")
		with hooks({"1.1.1": [f"{HERE}.warn_check"]}):
			self.assertIsNone(run_quietly(current="1.1.1")[0])
		self.assertEqual(CALLS, [])
		self.assertEqual(self.notifications(), [])

	def test_install_does_not_rewind_an_existing_record(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.2.0")
		upgrade.mark_installed("1.1.1")
		self.assertEqual(upgrade.reported_version(), "1.2.0")

	def test_a_site_upgraded_before_the_report_existed_gets_the_checks(self):
		# No record at all: a 1.0.x or 1.1.0 site migrating to the first release with the runner.
		with hooks({"1.1.1": [f"{HERE}.warn_check"]}):
			findings, _out = run_quietly(current="1.1.1")
		self.assertEqual(len(findings), 1)


class TestDisplay(UpgradeTestCase):
	def _report(self):
		upgrade._save(upgrade.VERSION_FIELD, "1.1.0")
		with hooks({"1.1.1": [f"{HERE}.warn_check", f"{HERE}.fail_check"]}):
			run_quietly(current="1.1.1")

	def test_doctor_returns_the_report_separately(self):
		self._report()
		result = doctor.run_diagnostics()
		report = result["upgrade_report"]
		self.assertEqual(report["version"], "1.1.1")
		self.assertEqual({f["check"] for f in report["findings"]}, {"upgrade.test_warn", "upgrade.test_fail"})
		# Not mixed into the live checks or their counts.
		self.assertFalse(any(f["check"].startswith("upgrade.test") for f in result["findings"]))

	def test_doctor_without_a_report(self):
		self.assertIsNone(doctor.run_diagnostics()["upgrade_report"])

	def test_dismiss_hides_it_and_does_not_rerun(self):
		self._report()
		upgrade.dismiss_report()
		self.assertIsNone(upgrade.get_report())
		self.assertIsNone(doctor.run_diagnostics()["upgrade_report"])
		with hooks({"1.1.1": [f"{HERE}.warn_check"]}):
			self.assertIsNone(run_quietly(current="1.1.1")[0])

	def test_dismiss_is_for_system_managers_only(self):
		self._report()
		frappe.set_user("Guest")
		try:
			with self.assertRaises(frappe.PermissionError):
				upgrade.dismiss_report()
		finally:
			frappe.set_user("Administrator")
		self.assertTrue(upgrade.get_report())

	def test_settings_form_gets_a_pointer(self):
		self._report()
		doc = frappe.get_doc("Microsoft Settings")
		doc.run_method("onload")
		self.assertEqual(doc.get_onload().get("upgrade_report"), {"version": "1.1.1", "count": 2})

	def test_saving_settings_cannot_change_the_report(self):
		self._report()
		stored = frappe.db.get_single_value(upgrade.SETTINGS, upgrade.REPORT_FIELD, cache=False)
		doc = frappe.get_doc("Microsoft Settings")
		doc.set(upgrade.REPORT_FIELD, json.dumps({"version": "0.0.1", "findings": []}))
		doc.set(upgrade.VERSION_FIELD, "0.0.1")
		doc._keep_upgrade_state()
		self.assertEqual(doc.get(upgrade.REPORT_FIELD), stored)
		self.assertEqual(doc.get(upgrade.VERSION_FIELD), "1.1.1")

	def test_unreadable_report_is_ignored(self):
		upgrade._save(upgrade.REPORT_FIELD, "{not json")
		self.assertIsNone(upgrade.get_report())


class TestEveryCheckIsRegistered(BaseTestCase):
	"""A check that is written but not listed in hooks.py never runs, and nothing else notices.

	It has happened once already: three branches each added the hook, a merge kept all three
	assignments, and Python keeps only the last.
	"""

	def test_every_upgrade_check_function_is_in_the_hook(self):
		import ast
		import os

		package = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
		defined = set()
		for root, dirs, files in os.walk(package):
			dirs[:] = [d for d in dirs if d not in ("tests", "node_modules", "public")]
			for name in files:
				if not name.endswith(".py"):
					continue
				path = os.path.join(root, name)
				with open(path, encoding="utf-8") as f:
					tree = ast.parse(f.read())
				module = os.path.relpath(path, os.path.dirname(package))[:-3].replace(os.sep, ".")
				defined.update(
					f"{module}.{node.name}"
					for node in tree.body
					if isinstance(node, ast.FunctionDef) and node.name.startswith("upgrade_check")
				)

		registered = {path for paths in upgrade_hooks().values() for path in paths}
		self.assertTrue(defined)
		self.assertEqual(defined - registered, set())

	def test_hooks_assign_the_check_list_once(self):
		import ast
		import os

		hooks = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "hooks.py")
		with open(hooks, encoding="utf-8") as f:
			tree = ast.parse(f.read())
		assigned = [
			target.id
			for node in tree.body
			if isinstance(node, ast.Assign)
			for target in node.targets
			if isinstance(target, ast.Name) and target.id == "microsoft365_upgrade_checks"
		]
		self.assertEqual(len(assigned), 1)


def upgrade_hooks():
	from frappe_microsoft365 import hooks

	return hooks.microsoft365_upgrade_checks
