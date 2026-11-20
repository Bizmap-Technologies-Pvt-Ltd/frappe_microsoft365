// Shared renderer for diagnostic findings, used by Microsoft Settings and Email Account.
//
// Read-only: nothing here changes how Frappe sends or receives mail.

frappe.provide("frappe_microsoft365");

frappe_microsoft365.STATUS_STYLE = {
	fail: { colour: "red", label: __("Problem") },
	warn: { colour: "orange", label: __("Check") },
	pass: { colour: "green", label: __("OK") },
	skip: { colour: "gray", label: __("Note") },
};

frappe_microsoft365.render_finding = function (f) {
	const style = frappe_microsoft365.STATUS_STYLE[f.status] || frappe_microsoft365.STATUS_STYLE.skip;
	const esc = frappe.utils.escape_html;
	const parts = [
		`<div style="margin-bottom:12px;padding-left:10px;border-left:3px solid var(--${style.colour}-400,#ccc)">`,
		`<div><span class="indicator ${style.colour}">${style.label}</span> <b>${esc(f.title)}</b></div>`,
	];
	if (f.target) parts.push(`<div class="text-muted small">${esc(f.target)}</div>`);
	if (f.detail) parts.push(`<div class="small" style="margin-top:4px">${esc(f.detail)}</div>`);
	if (f.fix) parts.push(`<div class="small" style="margin-top:4px"><b>${__("Fix")}:</b> ${esc(f.fix)}</div>`);
	if (f.doc) {
		parts.push(
			`<div class="small" style="margin-top:4px"><a href="${esc(f.doc)}" target="_blank" rel="noopener">${__(
				"Microsoft documentation"
			)}</a></div>`
		);
	}
	parts.push("</div>");
	return parts.join("");
};

frappe_microsoft365.show_findings = function (title, result) {
	const findings = (result || {}).findings || [];
	const counts = (result || {}).counts || {};
	const order = { fail: 0, warn: 1, skip: 2, pass: 3 };
	findings.sort((a, b) => order[a.status] - order[b.status]);

	const summary = __("{0} problem(s), {1} to check, {2} note(s).", [
		counts.fail || 0,
		counts.warn || 0,
		counts.skip || 0,
	]);

	const upgrade = frappe_microsoft365.render_upgrade_report((result || {}).upgrade_report);

	// Frappe calls primary_action with the dialog's *values*, not the dialog, so the handler
	// has to close over the instance; `primary_action(d) { d.hide() }` silently did nothing.
	const dialog = new frappe.ui.Dialog({
		title: title,
		size: "large",
		fields: [
			{
				fieldtype: "HTML",
				options: `${upgrade}<p class="text-muted">${frappe.utils.escape_html(summary)}</p>
					${findings.map(frappe_microsoft365.render_finding).join("")}`,
			},
		],
		primary_action_label: __("Close"),
		primary_action() {
			dialog.hide();
		},
	});
	dialog.$wrapper.on("click", ".microsoft365-upgrade-dismiss", () => {
		frappe.call({
			method: "frappe_microsoft365.upgrade.dismiss_report",
			callback: () => dialog.$wrapper.find(".microsoft365-upgrade-report").remove(),
		});
	});
	dialog.show();
};

// What the last `bench migrate` found, from upgrade.py. Shown above the live checks until a
// System Manager dismisses it, because it is about data the upgrade met, not the setup now.
frappe_microsoft365.render_upgrade_report = function (report) {
	const findings = ((report || {}).findings || []).slice();
	if (!findings.length) return "";
	const esc = frappe.utils.escape_html;
	const order = { fail: 0, warn: 1, skip: 2, pass: 3 };
	findings.sort((a, b) => order[a.status] - order[b.status]);
	const omitted = report.omitted
		? `<p class="small text-muted">${esc(__("{0} more not listed.", [report.omitted]))}</p>`
		: "";
	return `<div class="microsoft365-upgrade-report" style="margin-bottom:16px;padding-bottom:8px;border-bottom:1px solid var(--border-color)">
		<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
			<b>${esc(__("After upgrading to {0}", [report.version]))}</b>
			<button class="btn btn-xs btn-default microsoft365-upgrade-dismiss">${esc(__("Dismiss"))}</button>
		</div>
		${findings.map(frappe_microsoft365.render_finding).join("")}
		${omitted}
	</div>`;
};
