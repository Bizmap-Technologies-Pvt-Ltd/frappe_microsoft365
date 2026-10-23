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

	new frappe.ui.Dialog({
		title: title,
		size: "large",
		fields: [
			{
				fieldtype: "HTML",
				options: `<p class="text-muted">${frappe.utils.escape_html(summary)}</p>
					${findings.map(frappe_microsoft365.render_finding).join("")}`,
			},
		],
		primary_action_label: __("Close"),
		primary_action(d) {
			d.hide();
		},
	}).show();
};
