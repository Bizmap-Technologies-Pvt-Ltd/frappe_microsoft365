// Check one mail account's Microsoft configuration, from the form where it fails.
//
// Read-only. This adds a button; it changes nothing about how the account sends or receives.

frappe.ui.form.on("Email Account", {
	refresh(frm) {
		if (frm.is_new() || frm.doc.auth_method !== "OAuth") return;

		frm.add_custom_button(__("Check Microsoft Setup"), () => {
			frappe.call({
				method: "frappe_microsoft365.doctor.run_for_email_account",
				args: { email_account: frm.doc.name },
				freeze: true,
				freeze_message: __("Checking…"),
				callback: (r) => {
					frappe_microsoft365.show_findings(__("Microsoft Connection Doctor"), r.message);
				},
			});
		});
	},
});
