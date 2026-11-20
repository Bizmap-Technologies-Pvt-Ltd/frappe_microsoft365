"""Values written into the PowerShell scripts an administrator runs are always literal text.

Also the two script fixes that go with it: the Site Grant Script names the registration that
makes the SharePoint calls, and the Exchange script signs in to Microsoft Graph before using it.
"""

from unittest.mock import patch

import frappe

from frappe_microsoft365 import doctor, script_values
from frappe_microsoft365.script_values import ps_literal
from frappe_microsoft365.tests.base import BaseTestCase

MAIN = "038b9c8e-2699-4858-b5f1-4f7a3d4077c4"
APP_ONLY = "9f0b1c2d-0000-4000-8000-000000000002"
SITE = "https://contoso.sharepoint.com/sites/Sales"


class TestLiteralStrings(BaseTestCase):
	def test_values_are_single_quoted(self):
		self.assertEqual(ps_literal("a$(whoami)b"), "'a$(whoami)b'")

	def test_quotes_are_doubled_including_the_typographic_ones_powershell_honours(self):
		self.assertEqual(ps_literal("O'Brien"), "'O''Brien'")
		self.assertEqual(ps_literal("a\u2019b"), "'a\u2019\u2019b'")
		self.assertEqual(ps_literal("a\u2018b\u201ac\u201bd"), "'a\u2018\u2018b\u201a\u201ac\u201b\u201bd'")

	def test_double_quotes_and_backticks_pass_through_inert(self):
		self.assertEqual(ps_literal('"; `n'), "'\"; `n'")


class TestValidation(BaseTestCase):
	def test_client_ids_must_be_guids(self):
		self.assertEqual(script_values.validate_client_id(f" {MAIN} "), MAIN)
		for bad in ("client-abc", f'{MAIN}"; Remove-Item x', f"{MAIN}$(x)", ""):
			with self.assertRaises(frappe.ValidationError, msg=bad):
				script_values.validate_client_id(bad)

	def test_mailboxes_must_be_addresses(self):
		for good in ("support@contoso.com", "o'brien@contoso.co.uk", "a.b+tag@sub.contoso.com"):
			self.assertEqual(script_values.validate_mailbox(good), good)
		for bad in (
			"a$(whoami)@contoso.com",
			'a"@contoso.com',
			"a@contoso.com; Remove-Mailbox x",
			"a@contoso.com\nRemove-Mailbox x",
			"a`b@contoso.com",
			"not-an-address",
		):
			with self.assertRaises(frappe.ValidationError, msg=bad):
				script_values.validate_mailbox(bad)

	def test_site_urls_must_be_sharepoint_online(self):
		for good in (SITE, "contoso.sharepoint.com/teams/Ops", "https://contoso.sharepoint.com/sites/A_b-c.d/Shared%20Documents/x"):
			self.assertIsNone(script_values.site_url_problem(good), good)
		for bad in (
			'https://contoso.sharepoint.com/sites/Sales"$(whoami)',
			"https://contoso.sharepoint.com/sites/a$(whoami)",
			"https://contoso.sharepoint.com/sites/a;b",
			"https://contoso.sharepoint.com/sites/a'b",
			"https://evil.example.com/sites/Sales",
			"https://contoso.sharepoint.com.evil.com/sites/Sales",
			"http://contoso.sharepoint.com/sites/Sales",
			"https://contoso.sharepoint.com/sites/Sales\n# x",
			"https://user@contoso.sharepoint.com/sites/Sales",
			"",
		):
			self.assertIsNotNone(script_values.site_url_problem(bad), bad)
			with self.assertRaises(frappe.ValidationError, msg=bad):
				script_values.validate_site_url(bad)


class TestScriptsRefuseOrQuote(BaseTestCase):
	def test_a_hostile_mailbox_is_refused(self):
		for bad in ('x"$(Remove-Mailbox -Identity ceo)"@contoso.com', "a@b.com; Remove-Mailbox ceo"):
			with self.assertRaises(frappe.ValidationError):
				doctor.powershell_for_app_only(MAIN, mailboxes=[bad])

	def test_an_apostrophe_in_a_mailbox_stays_inside_its_string(self):
		script = doctor.powershell_for_app_only(MAIN, mailboxes=["o'brien@contoso.com"], send_as=True)

		self.assertIn("Add-MailboxPermission -Identity 'o''brien@contoso.com' -User", script)
		self.assertIn("# Remove-MailboxPermission -Identity 'o''brien@contoso.com' -User", script)

	def test_a_hostile_site_url_is_refused(self):
		for bad in ('https://c.sharepoint.com/sites/S"$(whoami)', 'https://c.sharepoint.com";$(whoami)/sites/S'):
			with self.assertRaises(frappe.ValidationError):
				doctor.powershell_for_site_grant(MAIN, [bad])

	def test_a_hostile_client_id_is_refused_by_both_scripts(self):
		bad = f'{MAIN}"; Remove-Item C:\\'
		with self.assertRaises(frappe.ValidationError):
			doctor.powershell_for_site_grant(bad, [SITE])
		with self.assertRaises(frappe.ValidationError):
			doctor.powershell_for_app_only(bad)

	def test_no_inserted_value_sits_in_a_double_quoted_string(self):
		grant = doctor.powershell_for_site_grant(MAIN, [SITE + "/Shared%20Documents/Forms/AllItems.aspx"])
		exchange = doctor.powershell_for_app_only(MAIN, mailboxes=["support@contoso.com"], send_as=True)

		for script, values in ((grant, (MAIN, "contoso.sharepoint.com")), (exchange, (MAIN, "support@contoso.com"))):
			for line in script.splitlines():
				for value in values:
					if value in line and not line.lstrip().startswith("#"):
						self.assertIn(f"'{value}", line, line)
						self.assertNotIn(f'"{value}', line, line)

	def test_only_the_site_address_reaches_the_comment(self):
		script = doctor.powershell_for_site_grant(MAIN, [SITE + "/Shared%20Documents/Forms/AllItems.aspx?x=1"])

		self.assertIn(f"# {SITE}\n", script)
		self.assertNotIn("AllItems", script)

	def test_endpoint_refuses_bad_mailboxes_from_the_dialog(self):
		with self.assertRaises(frappe.ValidationError):
			doctor.app_only_powershell(mailboxes="good@contoso.com\nbad$(x)@contoso.com")


class TestSaveTimeValidation(BaseTestCase):
	def test_microsoft_settings_refuses_a_client_id_that_is_not_a_guid(self):
		doc = frappe.get_doc("Microsoft Settings")
		for field in ("client_id", "app_client_id"):
			doc.reload()
			doc.set(field, 'abc"$(whoami)')
			with self.assertRaises(frappe.ValidationError, msg=field):
				doc.save(ignore_permissions=True)
		frappe.clear_document_cache("Microsoft Settings", "Microsoft Settings")

	def test_sharepoint_mapping_refuses_a_hostile_site_url(self):
		doc = frappe.get_doc(
			{
				"doctype": "SharePoint Mapping",
				"reference_doctype": "ToDo",
				"enabled": 0,
				"site_url": 'https://c.sharepoint.com/sites/S"$(whoami)',
			}
		)
		with self.assertRaises(frappe.ValidationError):
			doc.insert(ignore_permissions=True)


class TestUpgradeReport(BaseTestCase):
	def test_lists_saved_values_that_would_now_be_refused(self):
		before = {f: frappe.db.get_single_value("Microsoft Settings", f) for f in ("client_id", "app_client_id")}
		self.addCleanup(lambda: [frappe.db.set_single_value("Microsoft Settings", f, v) for f, v in before.items()])
		frappe.db.set_single_value("Microsoft Settings", "client_id", "client-abc")
		frappe.db.set_single_value("Microsoft Settings", "app_client_id", APP_ONLY)

		with patch.object(
			frappe,
			"get_all",
			side_effect=lambda doctype, **kw: [
				frappe._dict(name="ToDo", site_url="https://c.sharepoint.com/sites/a$(x)"),
				frappe._dict(name="Note", site_url=SITE),
			]
			if doctype == "SharePoint Mapping"
			else [],
		):
			found = script_values.upgrade_check()

		self.assertEqual([f["target"] for f in found], ["Microsoft Settings", "ToDo"])
		self.assertTrue(all(f["check"] == "upgrade.script_values" for f in found))
		self.assertTrue(all(f["status"] == doctor.WARN and f["fix"] for f in found))

	def test_clean_site_reports_nothing(self):
		before = frappe.db.get_single_value("Microsoft Settings", "client_id")
		self.addCleanup(frappe.db.set_single_value, "Microsoft Settings", "client_id", before)
		frappe.db.set_single_value("Microsoft Settings", "client_id", MAIN)

		with patch.object(frappe, "get_all", return_value=[frappe._dict(name="Note", site_url=SITE)]):
			self.assertEqual(script_values.upgrade_check(), [])


class TestExchangeScriptRunsAsPasted(BaseTestCase):
	def test_signs_in_to_graph_before_using_it(self):
		script = doctor.powershell_for_app_only(MAIN, mailboxes=["support@contoso.com"])

		self.assertIn("Install-Module -Name Microsoft.Graph.Applications", script)
		self.assertIn("Connect-MgGraph -Scopes 'Application.Read.All'", script)
		self.assertLess(script.index("Connect-MgGraph"), script.index("Get-MgServicePrincipal"))

	def test_no_retired_azuread_module(self):
		script = doctor.powershell_for_app_only(MAIN)

		self.assertNotIn("AzureAD", script)

	def test_finds_the_exchange_principal_by_app_id_not_display_name(self):
		script = doctor.powershell_for_app_only(MAIN)

		self.assertIn("Get-ServicePrincipal -Identity $appId", script)
		self.assertNotIn('Get-ServicePrincipal -Identity "Frappe mail access"', script)
		self.assertIn("if (-not $exoSp) { $exoSp = New-ServicePrincipal -AppId $appId", script)

	def test_every_statement_is_on_one_line(self):
		script = doctor.powershell_for_app_only(MAIN, mailboxes=["a@contoso.com"], send_as=True, include_undo=False)
		self.assertTrue(script.endswith("\n"), "a last line without a line break is never run")
		for line in script.splitlines():
			code = line.split("#", 1)[0].rstrip()
			self.assertEqual(code.count("{"), code.count("}"), line)
			self.assertEqual(code.count("("), code.count(")"), line)
