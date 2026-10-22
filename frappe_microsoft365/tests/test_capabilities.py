"""The capability registry: one definition per tickbox, read by everything else."""

from unittest.mock import patch

import frappe

from frappe_microsoft365 import capabilities
from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365 import provisioning
from frappe_microsoft365.tests.base import BaseTestCase


class TestRegistry(BaseTestCase):
	def test_every_capability_has_a_tickbox_on_microsoft_settings(self):
		meta = frappe.get_meta("Microsoft Settings")
		for capability in capabilities.all_capabilities():
			field = meta.get_field(capability.field)
			self.assertIsNotNone(field, capability.field)
			self.assertEqual(field.fieldtype, "Check")

	def test_ids_and_fields_are_unique(self):
		ids = [c.id for c in capabilities.all_capabilities()]
		fields = [c.field for c in capabilities.all_capabilities()]
		self.assertEqual(len(ids), len(set(ids)))
		self.assertEqual(len(fields), len(set(fields)))

	def test_requirements_name_real_capabilities(self):
		for capability in capabilities.all_capabilities():
			for required in capability.requires:
				capabilities.get(required)

	def test_on_enable_hooks_resolve(self):
		for capability in capabilities.all_capabilities():
			if capability.on_enable:
				self.assertTrue(callable(frappe.get_attr(capability.on_enable)))

	def test_only_delegated_capabilities_add_sign_in_scopes(self):
		fields = [field for field, _scopes in graph.CAPABILITY_SCOPES]
		self.assertEqual(fields, ["use_calendar", "use_teams", "use_transcripts"])
		self.assertNotIn("Sites.Selected", graph.derive_scopes({"use_files": 1}))

	def test_ticked_follows_registry_order(self):
		settings = {"use_files": 1, "use_calendar": 1, "use_mail": 0}
		self.assertEqual([c.id for c in capabilities.ticked(settings)], ["calendar", "files"])
		self.assertEqual(provisioning.selected_capabilities(settings), ["calendar", "files"])

	def test_set_up_lists_every_capability(self):
		listed = {c["id"]: c for c in provisioning.capabilities()}
		self.assertEqual(list(listed), [c.id for c in capabilities.all_capabilities()])
		self.assertEqual(listed["files"]["azure"], ["Sites.Selected"])
		self.assertIn("Microsoft 365 Mail", listed["mail"]["creates"])


class TestRules(BaseTestCase):
	def test_version_requirement(self):
		files = capabilities.get("files")
		self.assertIn("v16", capabilities.unmet(files, {"use_files": 1}, frappe_major=15))
		self.assertIsNone(capabilities.unmet(files, {"use_files": 1}, frappe_major=16))

	def test_missing_prerequisite(self):
		transcripts = capabilities.get("transcripts")
		self.assertTrue(capabilities.unmet(transcripts, {"use_transcripts": 1}, frappe_major=16))
		self.assertIsNone(
			capabilities.unmet(transcripts, {"use_transcripts": 1, "use_teams": 1}, frappe_major=16)
		)

	def test_is_enabled_needs_the_integration_on_too(self):
		self.assertTrue(capabilities.is_enabled("files", frappe._dict(enabled=1, use_files=1)))
		self.assertFalse(capabilities.is_enabled("files", frappe._dict(enabled=0, use_files=1)))
		self.assertFalse(capabilities.is_enabled("files", frappe._dict(enabled=1, use_files=0)))


class TestSwitchingOn(BaseTestCase):
	def test_on_enable_runs_only_for_newly_ticked(self):
		with patch.object(capabilities, "run_on_enable") as run:
			capabilities.apply_changes({"use_files": 1}, {"use_files": 1, "use_calendar": 1})
		self.assertEqual([c.args[0].id for c in run.call_args_list], ["calendar"])

	def test_switching_off_runs_nothing(self):
		with patch.object(capabilities, "run_on_enable") as run:
			capabilities.apply_changes({"use_files": 1}, {"use_files": 0})
		run.assert_not_called()

	def test_v15_refuses_storage(self):
		settings = frappe.get_single("Microsoft Settings")
		self.addCleanup(frappe.db.rollback)
		settings.update({"enabled": 0, "use_files": 1})
		with patch.object(capabilities, "frappe_major_version", return_value=15):
			with self.assertRaises(frappe.ValidationError):
				settings.save(ignore_permissions=True)
