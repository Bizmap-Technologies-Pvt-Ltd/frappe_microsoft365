// Microsoft Settings — diagnostics and Exchange setup helpers.
//
// These only read configuration and generate text. Nothing here changes how Frappe sends
// or receives mail.

const STATUS_STYLE = {
	fail: { colour: "red", label: __("Problem") },
	warn: { colour: "orange", label: __("Check") },
	pass: { colour: "green", label: __("OK") },
	skip: { colour: "gray", label: __("Note") },
};

frappe.ui.form.on("Microsoft Settings", {
	refresh(frm) {
		frm.add_custom_button(__("Run Diagnostics"), () => run_diagnostics(), __("Troubleshoot"));
		frm.add_custom_button(__("Explain an Error"), () => explain_error(), __("Troubleshoot"));
		frm.add_custom_button(__("Exchange Setup Script"), () => powershell(frm), __("Troubleshoot"));
	},
});

function run_diagnostics() {
	frappe.call({
		method: "frappe_microsoft365.doctor.run_diagnostics",
		freeze: true,
		freeze_message: __("Checking your Microsoft configuration…"),
		callback: (r) => {
			const result = r.message || {};
			const findings = result.findings || [];
			const counts = result.counts || {};

			const summary = __("{0} problem(s), {1} to check, {2} note(s).", [
				counts.fail || 0,
				counts.warn || 0,
				counts.skip || 0,
			]);

			const order = { fail: 0, warn: 1, skip: 2, pass: 3 };
			findings.sort((a, b) => order[a.status] - order[b.status]);

			new frappe.ui.Dialog({
				title: __("Microsoft Connection Doctor"),
				size: "large",
				fields: [
					{
						fieldtype: "HTML",
						options: `<p class="text-muted">${frappe.utils.escape_html(summary)}</p>
							${findings.map(render_finding).join("")}`,
					},
				],
				primary_action_label: __("Close"),
				primary_action(dialog) {
					dialog.hide();
				},
			}).show();
		},
	});
}

function render_finding(f) {
	const style = STATUS_STYLE[f.status] || STATUS_STYLE.skip;
	const esc = frappe.utils.escape_html;
	const parts = [
		`<div style="margin-bottom:12px;padding-left:10px;border-left:3px solid var(--${style.colour}-400,#ccc)">`,
		`<div><span class="indicator ${style.colour}">${style.label}</span> <b>${esc(f.title)}</b></div>`,
	];
	if (f.target) {
		parts.push(`<div class="text-muted small">${esc(f.target)}</div>`);
	}
	if (f.detail) {
		parts.push(`<div class="small" style="margin-top:4px">${esc(f.detail)}</div>`);
	}
	if (f.fix) {
		parts.push(`<div class="small" style="margin-top:4px"><b>${__("Fix")}:</b> ${esc(f.fix)}</div>`);
	}
	if (f.doc) {
		parts.push(
			`<div class="small" style="margin-top:4px"><a href="${esc(f.doc)}" target="_blank" rel="noopener">${__(
				"Microsoft documentation"
			)}</a></div>`
		);
	}
	parts.push("</div>");
	return parts.join("");
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

function powershell(frm) {
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
				callback: (r) => {
					dialog.set_value("script", (r.message || {}).script || "");
				},
			});
		},
	});
	dialog.show();
}
