// "Files in Microsoft 365" — the SharePoint folder of a record, shown on its form.
//
// Loaded on every desk page (app_include_js) because which DocTypes are mapped is a setting,
// not code: the list arrives in the boot (frappe.boot.microsoft365_files_doctypes) and the
// panel draws itself only on those forms. Everything it shows comes from whitelisted methods
// that check the user can read the record; ids in the page are never trusted on their own.

frappe.provide("frappe_microsoft365.files");

(function () {
	const esc = (s) => frappe.utils.escape_html(s == null ? "" : String(s));
	const METHOD = "frappe_microsoft365.microsoft_files";

	function mapped(doctype) {
		return (frappe.boot.microsoft365_files_doctypes || []).includes(doctype);
	}

	function size(bytes) {
		if (bytes == null) return "";
		if (bytes < 1024) return `${bytes} B`;
		if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(0)} KB`;
		return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
	}

	function when(iso) {
		if (!iso) return "";
		return frappe.datetime.comment_when
			? frappe.datetime.comment_when(iso.replace("T", " ").replace("Z", ""), true)
			: iso.slice(0, 10);
	}

	function open_url(frm, item_id, download) {
		const args = new URLSearchParams({
			doctype: frm.doctype,
			name: frm.docname,
			item_id,
			download: download ? 1 : 0,
		});
		return `/api/method/${METHOD}.open_item?${args.toString()}`;
	}

	class Panel {
		constructor(frm) {
			this.frm = frm;
			this.trail = []; // [{id, name}] subfolders opened from the record's folder
			this.$body = this.mount();
			this.load();
		}

		// Its own block at the top of the form's first tab. Not frm.dashboard: on a DocType
		// with a Connections tab (Project, Customer, …) Frappe puts the dashboard inside that
		// tab, where nobody looks for files.
		mount() {
			const frm = this.frm;
			frm.$wrapper.find(".microsoft365-files").remove();
			const $section = $(`
				<div class="microsoft365-files" style="padding:12px var(--padding-md, 15px);border-bottom:1px solid var(--border-color)">
					<div style="font-weight:600;margin-bottom:6px">${__("Files in Microsoft 365")}</div>
					<div class="microsoft365-files-body"></div>
				</div>`);
			const tabs = frm.layout && frm.layout.tabs;
			const target =
				tabs && tabs.length ? tabs[0].wrapper : frm.layout.wrapper.find(".form-page");
			$(target).prepend($section);
			return $section.find(".microsoft365-files-body");
		}

		// The same actions in the toolbar, so they are found without scrolling to the panel.
		toolbar(data) {
			const frm = this.frm;
			const group = __("SharePoint");
			frm.remove_custom_button(__("Create folder"), group);
			frm.remove_custom_button(__("Open in SharePoint"), group);
			frm.remove_custom_button(__("Show files"), group);
			if (!data.exists && data.can_create) {
				frm.add_custom_button(__("Create folder"), () => this.create(), group);
			}
			if (data.exists && data.web_url) {
				frm.add_custom_button(
					__("Open in SharePoint"),
					() => window.open(data.web_url, "_blank"),
					group,
				);
			}
			frm.add_custom_button(
				__("Show files"),
				() => {
					frm.layout.tabs && frm.layout.tabs.length && frm.layout.tabs[0].set_active();
					frm.$wrapper
						.find(".microsoft365-files")[0]
						?.scrollIntoView({ behavior: "smooth" });
				},
				group,
			);
		}

		create() {
			frappe.call({
				method: `${METHOD}.create_folder`,
				args: { doctype: this.frm.doctype, name: this.frm.docname },
				freeze: true,
				freeze_message: __("Creating the folder in SharePoint…"),
				callback: (r) => {
					const queued = (r.message || {}).queued || 0;
					if (queued) {
						frappe.show_alert({
							message: __("{0} attachment(s) are moving to the folder.", [queued]),
							indicator: "green",
						});
					}
					this.load();
				},
			});
		}

		load() {
			const subfolder = this.trail.length ? this.trail[this.trail.length - 1].id : null;
			this.$body.html(`<div class="text-muted small">${__("Loading…")}</div>`);
			frappe.call({
				method: `${METHOD}.list_folder`,
				args: { doctype: this.frm.doctype, name: this.frm.docname, subfolder },
				callback: (r) => this.render(r.message || {}),
				error: () =>
					this.$body.html(
						`<div class="text-muted small">${__(
							"Could not read the SharePoint folder. Microsoft Settings > Troubleshoot > Test SharePoint Connection says why.",
						)}</div>`,
					),
			});
		}

		render(data) {
			if (!this.trail.length) this.toolbar(data);
			if (!data.exists) {
				const kind = __(this.frm.doctype);
				let msg = data.missing
					? __("This {0}'s SharePoint folder was deleted or moved out of the library.", [
							kind,
						])
					: data.on_demand
						? __(
								"This {0}'s files stay on this server. Create a SharePoint folder to move them there; new attachments will follow.",
								[kind],
							)
						: __("No SharePoint folder yet. It is created with the first attachment.");
				if (!data.can_create && (data.on_demand || data.missing)) {
					msg += " " + __("Ask a System Manager to create one.");
				}
				const btn = data.can_create
					? `<button class="btn btn-xs btn-default ms365-create">${__("Create folder now")}</button>`
					: "";
				this.$body.html(
					`<div class="small text-muted" style="display:flex;gap:10px;align-items:center">${esc(msg)} ${btn}</div>`,
				);
				this.$body.find(".ms365-create").on("click", () => this.create());
				return;
			}

			const crumbs = [
				`<a class="ms365-crumb" data-depth="0">${esc(data.folder_name || this.frm.docname)}</a>`,
			]
				.concat(
					this.trail.map(
						(t, i) =>
							`<a class="ms365-crumb" data-depth="${i + 1}">${esc(t.name)}</a>`,
					),
				)
				.join(" / ");

			const rows = (data.items || []).map((item) => {
				const icon = item.is_folder ? "📁" : "📄";
				const link = item.is_folder
					? `<a class="ms365-dir" data-id="${esc(item.id)}" data-name="${esc(item.name)}">${esc(item.name)}</a>`
					: `<a href="${esc(open_url(this.frm, item.id))}" target="_blank" rel="noopener">${esc(item.name)}</a>`;
				const tag = item.attachment
					? ` <span class="text-muted small" title="${esc(__("Attached in Frappe"))}">· ${__("attachment")}</span>`
					: "";
				const extra = item.is_folder
					? esc(__("{0} items", [item.child_count || 0]))
					: esc(size(item.size));
				const download = item.is_folder
					? ""
					: `<a href="${esc(open_url(this.frm, item.id, true))}" title="${esc(__("Download"))}" class="text-muted">⤓</a>`;
				return `<tr>
					<td style="width:24px">${icon}</td>
					<td>${link}${tag}</td>
					<td class="text-muted small text-nowrap">${esc(item.modified_by || "")}</td>
					<td class="text-muted small text-nowrap">${esc(when(item.modified))}</td>
					<td class="text-muted small text-right text-nowrap">${extra}</td>
					<td style="width:24px" class="text-right">${download}</td>
				</tr>`;
			});

			const table = rows.length
				? `<table class="table table-sm" style="margin:6px 0 0"><tbody>${rows.join("")}</tbody></table>`
				: `<div class="text-muted small" style="margin-top:6px">${__(
						"Empty. Attach a file here, or drop one into the folder in Teams.",
					)}</div>`;

			this.$body.html(`
				<div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
					<div class="small" style="flex:1">${crumbs}</div>
					<button class="btn btn-xs btn-default ms365-refresh">${__("Refresh")}</button>
					${
						data.web_url
							? `<a class="btn btn-xs btn-default" href="${esc(data.web_url)}" target="_blank" rel="noopener">${__(
									"Open in SharePoint",
								)} ↗</a>`
							: ""
					}
					<button class="btn btn-xs btn-default ms365-help" title="${esc(__("How this panel works"))}">?</button>
				</div>
				${table}`);

			this.$body.find(".ms365-refresh").on("click", () => this.load());
			this.$body.find(".ms365-help").on("click", () => show_help());
			this.$body.find(".ms365-dir").on("click", (e) => {
				const $a = $(e.currentTarget);
				this.trail.push({ id: $a.data("id"), name: $a.data("name") });
				this.load();
			});
			this.$body.find(".ms365-crumb").on("click", (e) => {
				this.trail = this.trail.slice(0, Number($(e.currentTarget).data("depth")));
				this.load();
			});
		}
	}

	function show_help() {
		frappe.msgprint({
			title: __("Files in Microsoft 365"),
			indicator: "blue",
			message: [
				`<p>${__(
					"This record's files are kept in its own folder in SharePoint (the Teams channel's Files tab), not on this server.",
				)}</p><ul>`,
				`<li>${__(
					"<b>Attach</b> files from the sidebar as usual. A few seconds later they move into this folder; the sidebar link keeps working.",
				)}</li>`,
				`<li>${__(
					"Files saved into the folder from <b>Teams or SharePoint</b> appear here after <b>Refresh</b>. Ones marked <i>attachment</i> were attached in this system.",
				)}</li>`,
				`<li>${__(
					"Click a name to open it, the arrow to download it, a folder to go into it. Anyone who can read this record can open its files here.",
				)}</li>`,
				`<li>${__(
					"<b>Open in SharePoint</b> opens the folder itself, for people with access to the channel. Renaming or moving the folder there does not break this link.",
				)}</li>`,
				`<li>${__(
					"Removing an attachment here does not delete the SharePoint copy unless an administrator has turned that on.",
				)}</li></ul>`,
			].join(""),
		});
	}

	frappe_microsoft365.files.Panel = Panel;

	// "Attachments in SharePoint: Stored: 12 · Failed: 1". Shown on SharePoint Settings, and under
	// the capability's tickbox on Microsoft Settings, where `link` adds the way to its settings.
	function render_summary($wrapper, opts = {}) {
		frappe.call({
			method: `${METHOD}.storage_summary`,
			callback: (r) => {
				const c = r.message || {};
				const parts = ["Stored", "Archived", "Pending", "Failed"]
					.filter((k) => c[k])
					.map((k) => `${__(k)}: <b>${c[k]}</b>`);
				if (c.Failed) {
					parts.push(`<a href="/app/file?custom_microsoft_status=Failed">${__("see failed")}</a>`);
				}
				const counts = parts.length
					? `${__("Attachments in SharePoint")}: ${parts.join(" · ")}`
					: __("No attachments in SharePoint yet.");
				const link = opts.link
					? `<div><a href="/app/sharepoint-settings">${__("SharePoint Settings")} →</a></div>`
					: "";
				const style = opts.link ? "margin:-4px 0 8px 24px" : "margin-bottom:8px";
				$wrapper.html(
					`<div class="small text-muted" style="${style}">${counts}${link}</div>`,
				);
			},
		});
	}

	// Test SharePoint Connection: on SharePoint Settings and under Troubleshoot on Microsoft Settings.
	function test_connection(frm) {
		if (frm.is_dirty()) {
			frappe.msgprint(__("Save first: the test uses the saved settings."));
			return;
		}
		frappe.call({
			method: "frappe_microsoft365.microsoft_files.test_connection",
			freeze: true,
			freeze_message: __("Signing in as the app and writing a test folder in each library…"),
			callback: (r) => {
				const result = r.message || {};
				const findings = result.findings || [];
				const counts = {};
				findings.forEach((f) => (counts[f.status] = (counts[f.status] || 0) + 1));
				frappe_microsoft365.show_findings(__("SharePoint Connection"), { findings, counts });
				frm.reload_doc();
			},
		});
	}

	frappe_microsoft365.files.render_summary = render_summary;
	frappe_microsoft365.files.test_connection = test_connection;

	// A file is moved a few seconds after it is attached, by which time the sidebar already
	// shows its old /private/files link. The server announces the move on the document's
	// realtime room; swap the URL in place so the link keeps working without a reload.
	frappe.realtime.on("microsoft365_file_moved", (data) => {
		const frm = window.cur_frm;
		if (
			!data ||
			!frm ||
			frm.doctype !== data.doctype ||
			cstr(frm.docname) !== cstr(data.docname)
		) {
			return;
		}
		const info = (frappe.model.docinfo[frm.doctype] || {})[frm.docname] || {};
		(info.attachments || []).forEach((a) => {
			if (a.name === data.file || a.file_url === data.old_url) a.file_url = data.file_url;
		});
		frm.attachments && frm.attachments.refresh();
		frm.timeline && frm.timeline.refresh();
		frm.microsoft365_files_panel && frm.microsoft365_files_panel.load();
	});

	$(document).on("form-refresh", (_e, frm) => {
		if (!frm || frm.is_new() || !mapped(frm.doctype)) return;
		// The dashboard is rebuilt on every refresh, taking any earlier panel with it.
		frm.microsoft365_files_panel = new Panel(frm);
	});
})();
