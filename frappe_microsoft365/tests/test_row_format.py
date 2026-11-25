"""A compressed tabFile or tabEvent is converted to DYNAMIC before migrate (#68)."""

import io
from contextlib import redirect_stdout
from unittest.mock import patch

import frappe

from frappe_microsoft365 import doctor, row_format
from frappe_microsoft365.tests.base import BaseTestCase

PROBE = "M365 Row Format Probe"
PROBE_TABLE = "tabM365 Row Format Probe"
CONVERT_PROBE = {PROBE: "alter table `tabM365 Row Format Probe` row_format=dynamic key_block_size=0"}


def _row_format(table):
	return frappe.db.sql(
		"""select row_format from information_schema.tables
		where table_schema = database() and lower(table_name) = %s""",
		table.lower(),
	)[0][0]


class TestRowFormat(BaseTestCase):
	def setUp(self):
		super().setUp()
		if frappe.db.db_type != "mariadb":
			self.skipTest("MariaDB only")
		frappe.db.sql_ddl("drop table if exists `tabM365 Row Format Probe`")
		self.addCleanup(frappe.db.sql_ddl, "drop table if exists `tabM365 Row Format Probe`")

	def _make_probe(self, options):
		frappe.db.sql_ddl(
			"create table `tabM365 Row Format Probe` (name varchar(140) primary key, value varchar(768),"
			" index value_index (value)) engine=InnoDB " + options
		)

	def test_compressed_table_is_converted_once(self):
		for options in ("row_format=compressed", "row_format=compressed key_block_size=8"):
			with self.subTest(options=options):
				frappe.db.sql_ddl("drop table if exists `tabM365 Row Format Probe`")
				self._make_probe(options)
				self.assertEqual(row_format.compressed_tables([PROBE]), [PROBE])

				out = io.StringIO()
				with redirect_stdout(out):
					converted, failed = row_format.make_dynamic(CONVERT_PROBE)
				self.assertEqual((converted, failed), ([PROBE], []))
				self.assertEqual(_row_format(PROBE_TABLE), "Dynamic")
				self.assertIn("converting `tabM365 Row Format Probe`", out.getvalue())

				# A second run finds nothing to do and says nothing.
				out = io.StringIO()
				with redirect_stdout(out):
					self.assertEqual(row_format.make_dynamic(CONVERT_PROBE), ([], []))
				self.assertEqual(out.getvalue(), "")

	def test_dynamic_table_is_left_alone(self):
		self._make_probe("row_format=dynamic")
		with patch.object(frappe.db, "sql_ddl") as ddl:
			self.assertEqual(row_format.make_dynamic(CONVERT_PROBE), ([], []))
		ddl.assert_not_called()

	def test_failed_conversion_reports_and_does_not_raise(self):
		self._make_probe("row_format=compressed")
		out = io.StringIO()
		with patch.object(frappe.db, "sql_ddl", side_effect=Exception("Lock wait timeout")), redirect_stdout(out):
			converted, failed = row_format.make_dynamic(CONVERT_PROBE)
		self.assertEqual((converted, failed), ([], [PROBE]))
		self.assertIn("could not convert", out.getvalue())
		self.assertIn("ROW_FORMAT=DYNAMIC KEY_BLOCK_SIZE=0", out.getvalue())

	def test_the_app_tables_are_dynamic_on_this_site(self):
		self.assertEqual(row_format.compressed_tables(), [])
		self.assertEqual(row_format.upgrade_check(), [])

	def test_upgrade_check_reports_a_compressed_table(self):
		with patch.object(row_format, "compressed_tables", return_value=["File"]):
			findings = row_format.upgrade_check()
		self.assertEqual(len(findings), 1)
		self.assertEqual(findings[0]["check"], "upgrade.compressed_tables")
		self.assertEqual(findings[0]["status"], doctor.FAIL)
		self.assertEqual(findings[0]["target"], "File")
		self.assertIn("KEY_BLOCK_SIZE=0", findings[0]["fix"])

	def test_hooks_run_before_install_and_migrate(self):
		from frappe_microsoft365 import hooks

		self.assertEqual(hooks.before_migrate, "frappe_microsoft365.row_format.before_migrate")
		self.assertEqual(hooks.before_install, "frappe_microsoft365.row_format.before_install")
		self.assertIn("frappe_microsoft365.row_format.upgrade_check", hooks.microsoft365_upgrade_checks["1.2.1"])
