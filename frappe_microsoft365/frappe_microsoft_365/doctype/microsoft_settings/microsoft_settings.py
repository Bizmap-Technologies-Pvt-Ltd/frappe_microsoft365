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
		if self.enabled:
			# The secret was missing from this check, so a half-configured integration could be
			# enabled and only fail later, at sign-in, with an Azure error code.
			missing = [
				label
				for label, value in (
					(_("Tenant ID"), self.tenant_id),
					(_("Client ID"), self.client_id),
					(_("Client Secret"), self.get_password("client_secret", raise_exception=False)),
				)
				if not value
			]
			if missing:
				frappe.throw(
					_("{0} are required to enable the integration.").format(", ".join(missing))
				)

		self._reject_the_secret_id()
		self._drop_capabilities_that_lost_their_prerequisite()
		self._refuse_capabilities_this_frappe_cannot_run()

	def on_update(self):
		from frappe_microsoft365 import capabilities
		from frappe_microsoft365.microsoft_graph import clear_app_token_cache

		# The tenant, client or secret may just have changed; a token minted under the old ones
		# must not be served again.
		clear_app_token_cache()
		before = self.get_doc_before_save()
		capabilities.apply_changes(before, self)
		if bool((before or {}).get("use_files")) != bool(self.use_files):
			# Which forms get the Files panel follows this tick; see refresh_desk_boot.
			from frappe_microsoft365.microsoft_files import refresh_desk_boot

			refresh_desk_boot()

	def _drop_capabilities_that_lost_their_prerequisite(self):
		"""A capability that builds on another cannot work without it — transcripts, for one,
		are reached through the meeting behind a join URL, so they need standalone meetings.

		The field is hidden when its prerequisite is off, and a hidden Check keeps whatever value
		it had — which would leave sign-in quietly asking Azure for transcript and recording
		consent that nothing on screen admits to wanting. Clear it instead, and say so.
		"""
		from frappe_microsoft365 import capabilities

		for capability in capabilities.ticked(self):
			missing = [r for r in capability.requires if not self.get(capabilities.get(r).field)]
			if not missing:
				continue
			self.set(capability.field, 0)
			frappe.msgprint(
				_("{0} was turned off: it builds on {1}.").format(
					_(capability.label),
					", ".join("<b>{0}</b>".format(_(capabilities.get(r).label)) for r in missing),
				),
				indicator="orange",
				alert=True,
			)

	def _refuse_capabilities_this_frappe_cannot_run(self):
		from frappe_microsoft365 import capabilities

		major = capabilities.frappe_major_version()
		for capability in capabilities.ticked(self):
			if major and major < capability.min_frappe:
				frappe.throw(capabilities.unmet(capability, self, major))

	def _reject_the_secret_id(self):
		"""Azure shows a secret's Value and its Secret ID together, and only the Value works.

		Pasting the wrong one is the commonest setup mistake there is, and Microsoft only says
		so at sign-in, with AADSTS7000215, by which point the person has moved on. The two are
		trivially distinguishable, so say it here instead.
		"""
		from frappe_microsoft365.doctor import MASKED, looks_like_guid

		secret = (self.client_secret or "").strip()
		if not secret or MASKED.match(secret):
			return

		if looks_like_guid(secret):
			frappe.throw(
				_(
					"That looks like the Secret ID, not the secret value. In Azure, "
					"Certificates & secrets shows <b>Value</b> next to <b>Secret ID</b>: copy "
					"the Value. It is shown only once, so create a new secret if you have "
					"navigated away."
				),
				title=_("Wrong column"),
			)


@frappe.whitelist()
def get_effective_redirect_uri():
	"""Expose the default redirect URI for the Settings form help text."""
	frappe.only_for("System Manager")
	from frappe_microsoft365.microsoft_graph import get_redirect_uri
	return get_redirect_uri()
