// Microsoft Settings — set up the capabilities you picked, and diagnose them.
//
// Set Up only ever CREATES records that are missing. Anything that already exists is
// reported and left alone, so a working setup is never rewritten underneath you.

frappe.ui.form.on("Microsoft Settings", {
	refresh(frm) {
		frm.add_custom_button(__("Set Up"), () => show_plan(frm)).addClass("btn-primary");
		frm.add_custom_button(__("Run Diagnostics"), () => run_diagnostics(), __("Troubleshoot"));
		frm.add_custom_button(__("Explain an Error"), () => explain_error(), __("Troubleshoot"));
		frm.add_custom_button(__("Exchange Setup Script"), () => powershell(), __("Troubleshoot"));

		if (!frm.doc.use_calendar && !frm.doc.use_mail && !frm.doc.use_sso) {
			frm.dashboard.set_headline(
				__("Pick at least one capability above, then click <b>Set Up</b>.")
			);
		}
	},
});

const ACTION_STYLE = {
	create: { colour: "blue", label: __("Will create") },
	exists: { colour: "green", label: __("Ready") },
	drift: { colour: "orange", label: __("Needs attention") },
	skip: { colour: "gray", label: __("Not selected") },
};

function show_plan(frm) {
	frappe.call({
		method: "frappe_microsoft365.provisioning.plan",
		freeze: true,
		freeze_message: __("Working out what is needed…"),
		callback: (r) => {
			const plan = r.message || {};
			const esc = frappe.utils.escape_html;
			const capabilities = {};
			(plan.capabilities || []).forEach((c) => (capabilities[c.id] = c));

			const blocked = (plan.blockers || []).length > 0;
			const will_create = (plan.steps || []).filter((s) => s.action === "create");

			const body = (plan.steps || [])
				.map((step) => {
					const capability = capabilities[step.capability] || {};
					const style = ACTION_STYLE[step.action] || ACTION_STYLE.skip;
					const rows = [
						`<div style="margin-bottom:14px;padding-left:10px;border-left:3px solid var(--${style.colour}-400,#ccc)">`,
						`<div><span class="indicator ${style.colour}">${style.label}</span> <b>${esc(
							capability.label || step.capability
						)}</b></div>`,
						`<div class="small" style="margin-top:4px">${esc(step.detail || "")}</div>`,
					];
					if (step.action !== "skip" && capability.azure) {
						rows.push(
							`<div class="small text-muted" style="margin-top:4px">${__(
								"Azure permissions"
							)} (${esc(capability.azure_type || "")}): <code>${esc(
								capability.azure.join(" ")
							)}</code></div>`
						);
					}
					if (step.action !== "skip" && capability.note) {
						rows.push(`<div class="small text-muted" style="margin-top:4px">${esc(capability.note)}</div>`);
					}
					(step.findings || []).forEach((f) => rows.push(frappe_microsoft365.render_finding(f)));
					rows.push("</div>");
					return rows.join("");
				})
				.join("");

			const blockers = blocked
				? `<div class="alert alert-warning small">${(plan.blockers || [])
						.map(esc)
						.join("<br>")}</div>`
				: "";

			const dialog = new frappe.ui.Dialog({
				title: __("Set Up Microsoft 365"),
				size: "large",
				fields: [{ fieldtype: "HTML", options: blockers + body }],
				primary_action_label: will_create.length
					? __("Create {0} item(s)", [will_create.length])
					: __("Close"),
				primary_action() {
					if (!will_create.length) return dialog.hide();
					dialog.hide();
					apply(frm);
				},
			});
			if (blocked) dialog.get_primary_btn().prop("disabled", true);
			dialog.show();
		},
	});
}

function apply(frm) {
	frappe.call({
		method: "frappe_microsoft365.provisioning.apply",
		freeze: true,
		freeze_message: __("Creating what is missing…"),
		callback: (r) => {
			const result = r.message || {};
			const created = result.created || [];
			frappe.show_alert({
				message: created.length
					? __("Created {0} item(s).", [created.length])
					: __("Nothing needed creating."),
				indicator: "green",
			});
			frappe_microsoft365.show_findings(__("Microsoft Connection Doctor"), result.diagnostics);
			frm.reload_doc();
		},
	});
}

function run_diagnostics() {
	frappe.call({
		method: "frappe_microsoft365.doctor.run_diagnostics",
		freeze: true,
		freeze_message: __("Checking your Microsoft configuration…"),
		callback: (r) => frappe_microsoft365.show_findings(__("Microsoft Connection Doctor"), r.message),
	});
}

function explain_error() {
	const dialog = new frappe.ui.Dialog({
		title: __("Explain an Error"),
		fields: [
			{
				fieldname: "error_text",
				fieldtype: "Small Text",
				label: __("Paste the error from the Error Log"),
				reqd: 1,
			},
			{ fieldname: "result", fieldtype: "HTML" },
		],
		primary_action_label: __("Explain"),
		primary_action(values) {
			frappe.call({
				method: "frappe_microsoft365.doctor.explain",
				args: { error_text: values.error_text },
				callback: (r) => {
					const m = r.message || {};
					const esc = frappe.utils.escape_html;
					dialog.fields_dict.result.$wrapper.html(
						`<div style="margin-top:10px"><b>${esc(m.title || "")}</b>
						<div class="small" style="margin-top:4px">${esc(m.detail || "")}</div></div>`
					);
				},
			});
		},
	});
	dialog.show();
}

function powershell() {
	const dialog = new frappe.ui.Dialog({
		title: __("Exchange Setup Script"),
		size: "large",
		fields: [
			{
				fieldname: "mailboxes",
				fieldtype: "Small Text",
				label: __("Mailboxes (one per line)"),
				description: __(
					"Access is granted per mailbox, so the application can never reach anything not listed here."
				),
			},
			{
				fieldname: "send_as",
				fieldtype: "Check",
				label: __("Also allow sending as these mailboxes"),
			},
			{ fieldname: "script", fieldtype: "Code", label: __("Run in Exchange Online PowerShell") },
		],
		primary_action_label: __("Generate"),
		primary_action(values) {
			frappe.call({
				method: "frappe_microsoft365.doctor.app_only_powershell",
				args: { mailboxes: values.mailboxes, send_as: values.send_as ? 1 : 0 },
				callback: (r) => dialog.set_value("script", (r.message || {}).script || ""),
			});
		},
	});
	dialog.show();
}
