"""Microsoft Settings — single config page for the Azure AD app (keys + scopes)."""

import frappe
from frappe import _
from frappe.model.document import Document


class MicrosoftSettings(Document):
	def validate(self):
		# Surface the effective redirect URI so the admin can register it in Azure.
		if not self.redirect_uri:
			from frappe_microsoft365.microsoft_graph import get_redirect_uri
			# don't persist a computed value silently if integration disabled; just hint via msg
			if self.enabled:
				frappe.msgprint(
					f"Register this Redirect URI in Azure: <code>{get_redirect_uri(self)}</code>",
					indicator="blue", alert=True,
				)
		if self.enabled and not (self.tenant_id and self.client_id):
			frappe.throw(_("Tenant ID and Client ID are required to enable the integration."))


@frappe.whitelist()
def get_effective_redirect_uri():
	"""Expose the default redirect URI for the Settings form help text."""
	frappe.only_for("System Manager")
	from frappe_microsoft365.microsoft_graph import get_redirect_uri
	return get_redirect_uri()
