"""Keep the tables this app widens out of MariaDB's compressed row format.

Since 1.1.0 the app adds an indexed ``varchar(768)`` to File (``custom_microsoft_local_url``) and
to Event (``custom_microsoft_event_id``). A table created by an older Frappe can still be stored
``ROW_FORMAT=COMPRESSED``; Frappe creates and declares these tables ``Dynamic`` today, but never
converts an existing one. In a compressed ``tabFile`` that index leaves MariaDB (strict mode)
unable to rebuild the table at all, so every later ``ALTER TABLE`` on it fails with "Row size too
large": this app's new columns, and Frappe's and other apps' own migrations too (#68).

So before migrate, and before the app is installed, each of these tables still compressed is
converted to ``Dynamic``, the format Frappe gives them. Nothing else about the table changes. A
conversion that fails is reported with the manual steps and migrate carries on; the upgrade check
keeps reporting the table until it is converted.
"""

import frappe
from frappe import _

# The tables this app adds wide, indexed columns to. Fixed statements, not built from names:
# KEY_BLOCK_SIZE=0 drops a compressed table's block size, which DYNAMIC refuses to keep (1478).
CONVERT = {
	"File": "alter table `tabFile` row_format=dynamic key_block_size=0",
	"Event": "alter table `tabEvent` row_format=dynamic key_block_size=0",
}


def manual_steps(doctype):
	return _(
		"Take a backup, then run `ALTER TABLE `tab{0}` ROW_FORMAT=DYNAMIC KEY_BLOCK_SIZE=0;` in `bench --site <site> mariadb` and migrate again. It rebuilds the table and holds writes to {0} while it runs."
	).format(doctype)


def compressed_tables(doctypes=None):
	"""The DocTypes among ``doctypes`` (default: those in CONVERT) whose table is COMPRESSED."""
	if frappe.db.db_type != "mariadb":
		return []
	doctypes = list(doctypes or CONVERT)
	by_table = {f"tab{dt}".lower(): dt for dt in doctypes}
	rows = frappe.db.sql(
		"""
		select lower(table_name)
		from information_schema.tables
		where table_schema = database()
			and lower(table_name) in %(tables)s
			and upper(row_format) = 'COMPRESSED'
		""",
		{"tables": tuple(by_table)},
	)
	return [by_table[row[0]] for row in rows]


def make_dynamic(statements=None):
	"""Convert each compressed table in ``statements`` ({doctype: ALTER}) to DYNAMIC.

	Returns ``(converted, failed)`` lists of DocTypes. Never raises: a table left compressed fails
	later in migrate with MariaDB's own error, and this has already printed what to do about it.
	"""
	statements = statements or CONVERT
	converted, failed = [], []
	for doctype in compressed_tables(statements):
		print(
			f"Microsoft 365: converting `tab{doctype}` from ROW_FORMAT=COMPRESSED to DYNAMIC, "
			"the format Frappe gives it (#68). This rebuilds the table once."
		)
		try:
			frappe.db.sql_ddl(statements[doctype])
		except Exception as e:
			failed.append(doctype)
			print(f"Microsoft 365: could not convert `tab{doctype}`: {e}")
			print(f"Microsoft 365: {manual_steps(doctype)}")
			continue
		converted.append(doctype)
	return converted, failed


def before_migrate():
	make_dynamic()


def before_install():
	make_dynamic()


def upgrade_check():
	"""``upgrade.compressed_tables``: a File or Event table still stored COMPRESSED."""
	from frappe_microsoft365.doctor import FAIL, finding

	return [
		finding(
			"upgrade.compressed_tables",
			FAIL,
			_("The {0} table is stored ROW_FORMAT=COMPRESSED").format(_(doctype)),
			_(
				"With this app's indexed columns, MariaDB cannot rebuild a compressed tab{0}, so any change to the {0} table fails with \"Row size too large\", including other apps' migrations. Migrate tries to convert it and could not."
			).format(doctype),
			manual_steps(doctype),
			target=doctype,
		)
		for doctype in compressed_tables()
	]
