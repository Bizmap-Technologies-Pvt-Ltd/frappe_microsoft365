// Microsoft Settings — set up the capabilities you picked, and diagnose them.
//
// Set Up only ever CREATES records that are missing. Anything that already exists is
// reported and left alone, so a working setup is never rewritten underneath you.

// The scope tickboxes. Anything ticked here changes what sign-in asks Microsoft for, so the
// derived list is re-rendered whenever one of them moves.
const SCOPE_FIELDS = ["use_calendar", "use_teams", "use_transcripts"];
const CAPABILITY_FIELDS = [...SCOPE_FIELDS, "use_mail", "use_sso"];

const scope_handlers = {};
[...SCOPE_FIELDS, "default_scopes"].forEach((field) => {
	scope_handlers[field] = (frm) => render_scopes(frm);
});

frappe.ui.form.on("Microsoft Settings", {
	...scope_handlers,
	refresh(frm) {
		frm.add_custom_button(__("Set Up"), () => show_plan(frm)).addClass("btn-primary");
		frm.add_custom_button(__("Setup Guide"), () => setup_guide());
		frm.add_custom_button(__("Run Diagnostics"), () => run_diagnostics(), __("Troubleshoot"));
		frm.add_custom_button(__("Explain an Error"), () => explain_error(), __("Troubleshoot"));
		frm.add_custom_button(__("Exchange Setup Script"), () => powershell(), __("Troubleshoot"));

		render_scopes(frm);

		if (!CAPABILITY_FIELDS.some((field) => frm.doc[field])) {
			frm.dashboard.set_headline(
				__(
					"Pick at least one capability above, then click <b>Set Up</b>. <b>Setup Guide</b> walks the Microsoft side in the order it has to happen."
				)
			);
		}
	},
});

// The Microsoft side spans two portals and Frappe, and the order is the whole point: a token
// carries only what was consented before it was issued, and the tenant switch transcripts need
// lives somewhere nobody thinks to look. The README says all of this, but the person who needs it
// is on this screen, not on GitHub.
function setup_guide() {
	frappe.call({
		method: "frappe_microsoft365.doctor.setup_guide",
		freeze: true,
		freeze_message: __("Working out your setup steps…"),
		callback: (r) => {
			const guide = r.message || {};
			const esc = frappe.utils.escape_html;

			const covered = (guide.capabilities || []).map((c) =>
				esc(typeof c === "string" ? c : c.label || c.id || "")
			);

			const head = [
				`<p class="small">${__(
					"Do these in order. A token carries only what was consented when it was issued, so anything ticked afterwards does nothing until the connection is re-authorised."
				)}</p>`,
			];
			if (covered.length) {
				head.push(
					`<p class="small text-muted">${__("Covers: {0}", [
						`<b>${covered.join(", ")}</b>`,
					])}</p>`
				);
			}

			const steps = (guide.steps || []).map((step, index) => {
				const rows = [
					`<div style="margin-bottom:14px;padding-left:10px;border-left:3px solid var(--blue-400,#ccc)">`,
					`<div><b>${index + 1}. ${esc(step.title || "")}</b></div>`,
				];
				// The portal path is what people scan for, so it never reads as prose.
				if (step.where) {
					rows.push(`<div style="margin-top:4px"><code>${esc(step.where)}</code></div>`);
				}
				if (step.unlocks) {
					rows.push(`<div class="small" style="margin-top:4px">${esc(step.unlocks)}</div>`);
				}
				if (step.verify) {
					rows.push(
						`<div class="small text-muted" style="margin-top:4px"><b>${__(
							"Check"
						)}:</b> ${esc(step.verify)}</div>`
					);
				}
				rows.push("</div>");
				return rows.join("");
			});

			new frappe.ui.Dialog({
				title: __("Setup Guide"),
				size: "large",
				fields: [{ fieldtype: "HTML", options: head.join("") + steps.join("") }],
				primary_action_label: __("Close"),
				primary_action(dialog) {
					dialog.hide();
				},
			}).show();
		},
	});
}

// Show exactly what sign-in will request, so nobody has to reverse-engineer it from the
// tickboxes — the guesswork that had admins editing the scope field by hand. Derived on the
// server by the same function get_scopes uses, so the display cannot drift from reality.
function render_scopes(frm) {
	const field = frm.fields_dict.effective_scopes_display;
	if (!field) return;

	const capabilities = {};
	SCOPE_FIELDS.forEach((name) => (capabilities[name] = frm.doc[name] ? 1 : 0));

	frappe.call({
		method: "frappe_microsoft365.microsoft_graph.preview_scopes",
		args: {
			capabilities: JSON.stringify(capabilities),
			override: frm.doc.default_scopes || "",
		},
		callback: (r) => {
			const result = r.message || {};
			const esc = frappe.utils.escape_html;
			const scopes = (result.scopes || []).join(" ");
			const added = (result.always_added || []).join(" ");
			const missing = result.missing_from_override || [];

			const rows = [
				`<div class="small text-muted">${__(
					"Grant these Microsoft Graph <b>Delegated</b> permissions in Azure, then Grant admin consent."
				)}</div>`,
				`<div style="margin-top:6px"><code>${esc(scopes)} ${esc(added)}</code></div>`,
				`<div class="small text-muted" style="margin-top:6px">${__(
					"{0} are requested automatically and never need listing below.",
					[`<code>${esc(added)}</code>`]
				)}</div>`,
			];
			if (result.overridden) {
				rows.push(
					`<div class="small" style="margin-top:6px"><span class="indicator orange">${__(
						"Overridden"
					)}</span> ${__("The capabilities would ask for {0}.", [
						`<code>${esc((result.derived || []).join(" "))}</code>`,
					])}</div>`
				);
			}
			if (missing.length) {
				rows.push(
					`<div class="alert alert-warning small" style="margin-top:8px">${__(
						"The override leaves out {0}, which a ticked capability needs. That feature will fail with a 403.",
						[`<code>${esc(missing.join(" "))}</code>`]
					)}</div>`
				);
			}
			field.$wrapper.html(rows.join(""));
		},
	});
}

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

			// The per-capability lists overlap, so also give the one line to paste into Azure.
			const graph_selected = (plan.selected || []).some((id) =>
				["calendar", "teams", "transcripts"].includes(id)
			);
			const scope_summary =
				graph_selected && (plan.graph_scopes || []).length
					? `<div class="small text-muted" style="margin-top:4px">${__(
							"All Graph permissions to consent to, combined"
					  )}: <code>${esc((plan.graph_scopes || []).join(" "))}</code></div>`
					: "";

			const dialog = new frappe.ui.Dialog({
				title: __("Set Up Microsoft 365"),
				size: "large",
				fields: [{ fieldtype: "HTML", options: blockers + body + scope_summary }],
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
	// This is the most consequential thing in the app, so the dialog says plainly what the
	// script does and will not generate it until that has been acknowledged. Nothing is
	// executed from here: the admin runs it in Exchange Online themselves.
	const warning = `
		<div class="alert alert-warning" style="margin-bottom:12px">
			<b>${__("Read before you run this")}</b>
			<ul style="margin:8px 0 0 16px;padding:0">
				<li>${__(
					"It gives the application <b>permanent access to the mailboxes you list</b>, readable with no one signed in."
				)}</li>
				<li>${__(
					"Access is granted one mailbox at a time. Mailboxes not listed here stay out of reach."
				)}</li>
				<li>${__(
					"It is reversible. The generated script ends with the commands that undo it."
				)}</li>
				<li>${__(
					"Nothing runs from this screen. You paste it into Exchange Online yourself."
				)}</li>
			</ul>
			<div class="small" style="margin-top:8px">${__(
				"You only need this for shared mailboxes, such as turning support@ into a ticket queue. Ordinary calendar and mail sync does not."
			)}</div>
		</div>`;

	const dialog = new frappe.ui.Dialog({
		title: __("Exchange Setup Script"),
		size: "large",
		fields: [
			{ fieldtype: "HTML", options: warning },
			{
				fieldname: "mailboxes",
				fieldtype: "Small Text",
				label: __("Mailboxes (one per line)"),
				description: __(
					"Access is granted per mailbox, so the application can never reach anything not listed here."
				),
				reqd: 1,
			},
			{
				fieldname: "send_as",
				fieldtype: "Check",
				label: __("Also allow sending as these mailboxes"),
			},
			{
				fieldname: "understood",
				fieldtype: "Check",
				label: __("I understand this grants standing access to the mailboxes listed above"),
			},
			{ fieldname: "script", fieldtype: "Code", label: __("Run in Exchange Online PowerShell") },
		],
		primary_action_label: __("Generate"),
		primary_action(values) {
			if (!values.understood) {
				frappe.msgprint({
					title: __("Confirm first"),
					message: __("Tick the confirmation box to generate the script."),
					indicator: "orange",
				});
				return;
			}
			frappe.call({
				method: "frappe_microsoft365.doctor.app_only_powershell",
				args: { mailboxes: values.mailboxes, send_as: values.send_as ? 1 : 0 },
				callback: (r) => dialog.set_value("script", (r.message || {}).script || ""),
			});
		},
	});
	dialog.show();
}
