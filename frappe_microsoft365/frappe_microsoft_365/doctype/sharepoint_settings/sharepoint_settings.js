// SharePoint Settings — how document storage behaves, and the tools to set it up and check it.
//
// The capability itself is switched on in Microsoft Settings, beside the others, because it
// changes what the Azure app is asked for. Everything about running it is here.

frappe.ui.form.on("SharePoint Settings", {
	refresh(frm) {
		frappe.db.get_single_value("Microsoft Settings", "use_files").then((on) => {
			const $status = frm.fields_dict.status_html.$wrapper;
			if (!on) {
				$status.html(
					`<div class="alert alert-warning small">${__(
						"SharePoint document storage is off. Tick it under Capabilities in <a href='/app/microsoft-settings'>Microsoft Settings</a>, then come back here."
					)}</div>`
				);
				return;
			}
			$status.empty();
			show_grant_script(frm);
			frappe_microsoft365.files.render_summary($status);
			frm.add_custom_button(
				__("Test SharePoint Connection"),
				() => frappe_microsoft365.files.test_connection(frm),
				__("Troubleshoot")
			);
			frm.add_custom_button(__("Mappings"), () => frappe.set_route("List", "SharePoint Mapping"), __("Go to"));
			frm.add_custom_button(__("Move Existing Attachments"), () => move_existing(frm), __("Actions"));
		});
	},
});

// The grant script for every enabled mapping's site, on the page itself so the administrator
// who runs it can be sent straight here. Built from the mappings, so it names this site's own
// SharePoint sites and the app's client id.
function show_grant_script(frm) {
	const $wrapper = frm.fields_dict.grant_script_html.$wrapper;
	const intro = `<div class="text-muted small" style="margin-bottom:8px">${__(
		"Sites.Selected lets the app be granted individual sites; this script is the grant. It gives the app <b>write</b> access to the sites of the enabled SharePoint Mappings and to no other site. A SharePoint or Global administrator runs it in Microsoft Graph PowerShell. Nothing runs from this page."
	)}</div>`;
	frappe.db.count("SharePoint Mapping", { filters: { enabled: 1 } }).then((n) => {
		if (!n) {
			$wrapper.html(
				intro +
					`<div class="text-muted small">${__(
						"Add a SharePoint Mapping first; its site appears here."
					)}</div>`
			);
			return;
		}
		frappe.call({
			method: "frappe_microsoft365.doctor.site_grant_powershell",
			callback: (r) => {
				const script = (r.message || {}).script || "";
				$wrapper.html(
					intro +
						`<div style="position:relative">
							<button class="btn btn-xs btn-default ms365-copy" style="position:absolute;top:8px;right:8px">${__("Copy")}</button>
							<pre style="white-space:pre-wrap;font-size:12px;padding:12px;margin:0">${frappe.utils.escape_html(script)}</pre>
						</div>`
				);
				$wrapper.find(".ms365-copy").on("click", () => frappe.utils.copy_to_clipboard(script));
			},
		});
	});
}

function move_existing(frm) {
	frappe.confirm(
		__(
			"Queue every attachment already on records of the mapped DocTypes for SharePoint? Files move in the background; each keeps opening from its record throughout."
		),
		() =>
			frappe.call({
				method: "frappe_microsoft365.microsoft_files.queue_existing",
				freeze: true,
				callback: (r) => {
					const out = r.message || {};
					frappe.msgprint(
						out.more
							? __("{0} attachments queued. There are more: run this again once these are done.", [out.queued])
							: __("{0} attachments queued.", [out.queued])
					);
					frappe_microsoft365.files.render_summary(frm.fields_dict.status_html.$wrapper);
				},
			})
	);
}
