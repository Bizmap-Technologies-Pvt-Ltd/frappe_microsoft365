"""What an upgrade changed or found on this site, reported once per release.

Some fixes change behaviour for sites already using the app, and some find data an admin has to
look at (Events that will stop syncing, folders two records share). A patch can repair what is
safe to repair; everything else has to be *told* to someone, or it is found months later as
"sync just stopped".

Each release registers read-only checks in ``hooks.py``::

        microsoft365_upgrade_checks = {
            "1.1.1": ["frappe_microsoft365.<module>.<function>"],
        }

A check takes no arguments, only reads, and returns a list of ``doctor.finding(...)`` dicts
(``[]`` when there is nothing to report). After ``bench migrate`` the checks of every release
newer than the one last reported on this site, up to the app's own version, run once. Their
findings are printed in the migrate output, kept for the doctor to show at the top as *After
upgrading to X.Y.Z* until an admin dismisses them, and sent to each System Manager as one
Notification Log. A site with nothing to report gets no notification, and a fresh install runs
nothing: there is no earlier version for it to have upgraded from.

Where it is kept: two hidden fields on Microsoft Settings, ``upgrade_reported_version`` and
``upgrade_report`` (JSON). Not in Defaults, which Frappe sends to every signed-in user's browser
at boot; Microsoft Settings is readable by System Managers only. The controller restores both
fields from the database on every save, so a form left open across a migrate cannot write an
old report back, and nobody can set them through the API.
"""

import json
import re

import frappe
from frappe import _
from frappe.utils import escape_html, now

from frappe_microsoft365 import doctor

HOOK = "microsoft365_upgrade_checks"
SETTINGS = "Microsoft Settings"
VERSION_FIELD = "upgrade_reported_version"
REPORT_FIELD = "upgrade_report"

#: The report is stored as one JSON value. A check that returns one finding per row of a big
#: table must not be able to make it unreadable; past this the rest are counted, not listed.
MAX_FINDINGS = 100

#: Statuses worth telling someone about. PASS means nothing to report.
REPORTED = (doctor.FAIL, doctor.WARN, doctor.SKIP)


# --- versions -------------------------------------------------------------------------


def version_key(version):
	"""A comparable form of a version string: ``"1.1.10"`` sorts after ``"1.1.9"``.

	Plain string comparison gets that wrong, and so would float(). Anything after the numbers
	(``-dev``, ``b1``) is ignored, and trailing zeros are dropped so ``1.1`` equals ``1.1.0``.
	An empty or unreadable value is ``()``, older than every release.
	"""
	parts = []
	for piece in str(version or "").strip().lstrip("vV").split("."):
		match = re.match(r"\d+", piece)
		if not match:
			break
		parts.append(int(match.group()))
	while parts and parts[-1] == 0:
		parts.pop()
	return tuple(parts)


def app_version():
	from frappe_microsoft365 import __version__

	return __version__


# --- stored state ---------------------------------------------------------------------


def _has_fields():
	"""Whether Microsoft Settings has been migrated far enough to hold the report."""
	meta = frappe.get_meta(SETTINGS)
	return bool(meta.has_field(VERSION_FIELD) and meta.has_field(REPORT_FIELD))


def reported_version():
	"""The last release whose checks ran on this site, or None if none ever did."""
	return frappe.db.get_single_value(SETTINGS, VERSION_FIELD, cache=False) or None


def get_report():
	"""The stored report, or None when there is none or it was dismissed.

	``{"version", "from_version", "created_on", "findings": [...], "omitted"}``
	"""
	if not _has_fields():
		return None
	raw = frappe.db.get_single_value(SETTINGS, REPORT_FIELD, cache=False)
	if not raw:
		return None
	try:
		report = json.loads(raw)
	except ValueError:
		return None
	if not isinstance(report, dict) or not report.get("findings"):
		return None
	return report


def _save(field, value):
	# set_single_value leaves the request's value cache alone, which is why the reads above
	# pass cache=False: within one migrate they must see what was just written.
	frappe.db.set_single_value(SETTINGS, field, value)


# --- collecting and running checks ----------------------------------------------------


def registered_checks(after, upto):
	"""``[(version, [dotted path, ...]), ...]`` for releases in ``(after, upto]``, oldest first.

	``after`` None means no release was ever reported here, so every release up to ``upto``
	counts. A path registered twice (two apps, or a merge that repeated a line) runs once.
	"""
	hooks = frappe.get_hooks(HOOK) or {}
	low, high = version_key(after), version_key(upto)
	selected = []
	for version in sorted(hooks, key=version_key):
		key = version_key(version)
		if not (low < key <= high):
			continue
		paths = []
		for path in hooks[version] or []:
			if path and path not in paths:
				paths.append(path)
		if paths:
			selected.append((version, paths))
	return selected


def _check_failed(path, error):
	return doctor.finding(
		"upgrade.check_failed",
		doctor.FAIL,
		_("An upgrade check could not run"),
		_("{0} failed: {1}").format(path, error),
		_(
			"The other checks ran. The Error Log has the details; please report it to the app's "
			"maintainers, since whatever this check looks for has not been checked on this site."
		),
	)


def run_check(path, version):
	"""Run one check and return its findings, each tagged with the release it belongs to.

	A check that raises, or returns something that is not a list of findings, becomes a FAIL
	finding of its own and an Error Log entry. It never stops migrate, and anything it may have
	written before failing is rolled back to a savepoint.
	"""
	savepoint = "microsoft365_upgrade_check"
	frappe.db.savepoint(savepoint)
	try:
		result = frappe.get_attr(path)()
		if result is None:
			result = []
		if not isinstance(result, list | tuple) or not all(
			isinstance(f, dict) and f.get("status") for f in result
		):
			raise TypeError("expected a list of doctor findings, got {0!r}".format(result))
	except Exception as e:
		frappe.db.rollback(save_point=savepoint)
		frappe.log_error(title=_("Microsoft 365 upgrade check failed: {0}").format(path))
		result = [_check_failed(path, e)]

	findings = []
	for f in result:
		if f.get("status") not in REPORTED:
			continue
		f = dict(f)
		f["version"] = version
		findings.append(f)
	return findings


def run(current=None, previous=None):
	"""Run the checks this site has not reported yet; store, print and notify. Idempotent.

	``current`` defaults to the app's ``__version__``, ``previous`` to the version last reported
	on this site; tests pass both. Returns the new findings, or None when there was nothing to
	run (this release was already reported, or the site has not been migrated far enough).
	"""
	if not _has_fields():
		return None

	current = current or app_version()
	previous = previous if previous is not None else reported_version()
	if previous and version_key(previous) >= version_key(current):
		return None

	findings = []
	for version, paths in registered_checks(previous, current):
		for path in paths:
			findings += run_check(path, version)

	# Recorded first and in the same transaction as the report: a second migrate of the same
	# release must find it already done, whatever happens to the notification below.
	_save(VERSION_FIELD, current)
	if findings:
		_store(current, previous, findings)
		notify(current, findings)
	print_report(current, findings)
	return findings


def _store(current, previous, findings):
	"""Keep the findings for the doctor. Findings from an earlier release that nobody has
	dismissed yet are kept beneath the new ones, so a site that skipped reading them loses
	nothing by upgrading again."""
	earlier = get_report()
	if earlier:
		previous = earlier.get("from_version") or previous
		findings = findings + earlier.get("findings", [])
	omitted = max(len(findings) - MAX_FINDINGS, 0) + ((earlier or {}).get("omitted") or 0)
	report = {
		"version": current,
		"from_version": previous,
		"created_on": str(now()),
		"findings": findings[:MAX_FINDINGS],
		"omitted": omitted,
	}
	_save(REPORT_FIELD, json.dumps(report, default=str))


def mark_installed(version=None):
	"""A fresh install has nothing to upgrade from: start the record at this version."""
	if _has_fields() and not reported_version():
		_save(VERSION_FIELD, version or app_version())


# --- output ---------------------------------------------------------------------------


def _label(status):
	"""The doctor's own words for a status, so the console matches the dialog."""
	if status == doctor.FAIL:
		return _("Problem")
	if status == doctor.WARN:
		return _("Check")
	return _("Note")


def print_report(version, findings):
	"""A few lines in the migrate output: what to look at, and where the details are."""
	if not findings:
		print(_("Microsoft 365 {0}: upgrade checks found nothing to report.").format(version))
		return
	print(_("Microsoft 365 {0}: {1} item(s) to check after this upgrade").format(version, len(findings)))
	for f in findings:
		line = "  [{0}] {1}".format(_label(f["status"]), f.get("title") or f.get("check"))
		if f.get("target"):
			line += " ({0})".format(f["target"])
		print(line)
		if f.get("fix"):
			print("        {0}".format(f["fix"]))
	print("  " + _("Details: Microsoft Settings > Troubleshoot > Run Diagnostics."))


def system_managers():
	"""Enabled desk users holding System Manager, Administrator included."""
	users = frappe.get_all(
		"Has Role", filters={"role": "System Manager", "parenttype": "User"}, pluck="parent"
	)
	if not users:
		return []
	return sorted(
		set(
			frappe.get_all(
				"User",
				filters={"name": ["in", users], "enabled": 1, "user_type": "System User"},
				pluck="name",
			)
		)
	)


def notify(version, findings):
	"""One Notification Log per System Manager for this release. Best effort: a notification that
	cannot be created is logged, and the report is still in the doctor."""
	subject = _("Microsoft 365 {0}: {1} item(s) to check after the upgrade").format(version, len(findings))
	items = "".join(
		"<li>{0}</li>".format(escape_html(f.get("title") or f.get("check") or "")) for f in findings[:20]
	)
	content = "<p>{0}</p><ul>{1}</ul><p>{2}</p>".format(
		escape_html(_("The upgrade checks found the following on this site.")),
		items,
		escape_html(_("Open Microsoft Settings > Troubleshoot > Run Diagnostics for the details.")),
	)
	for user in system_managers():
		try:
			frappe.get_doc(
				{
					"doctype": "Notification Log",
					"for_user": user,
					"type": "Alert",
					"document_type": SETTINGS,
					"document_name": SETTINGS,
					"subject": subject,
					"email_content": content,
				}
			).insert(ignore_permissions=True)
		except Exception:
			frappe.log_error(title=_("Microsoft 365 upgrade notification failed for {0}").format(user))


# --- entry points ---------------------------------------------------------------------


def after_migrate():
	"""Called from setup.after_migrate. Never lets the report break a migrate."""
	try:
		run()
	except Exception:
		frappe.log_error(title=_("Microsoft 365 upgrade report failed"))
		print(_("Microsoft 365: the upgrade report could not be produced; see the Error Log."))


@frappe.whitelist(methods=["POST"])
def dismiss_report():
	"""Hide the stored report from the doctor. The checks are not run again for this release."""
	frappe.only_for("System Manager")
	if _has_fields():
		_save(REPORT_FIELD, "")
	return True
