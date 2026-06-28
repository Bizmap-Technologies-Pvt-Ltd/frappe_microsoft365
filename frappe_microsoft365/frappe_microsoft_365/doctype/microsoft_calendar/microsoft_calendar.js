// Microsoft Calendar — form actions: Authorize, Test connection, Sync, Disconnect.
frappe.ui.form.on("Microsoft Calendar", {
	refresh(frm) {
		if (frm.is_new()) {
			frm.dashboard.set_headline(
				__("Save this connection, then click <b>Authorize</b> to sign in with Microsoft.")
			);
			return;
		}

		if (!frm.doc.authorized) {
			frm.add_custom_button(__("Authorize Microsoft Access"), () => authorize(frm)).addClass(
				"btn-primary"
			);
		} else {
			frm.dashboard.set_headline(
				__("Connected as {0}", [frm.doc.microsoft_user_email || "Microsoft account"])
			);
			frm.add_custom_button(__("Test Connection"), () => test_conn(frm), __("Microsoft"));
			frm.add_custom_button(__("Sync Now"), () => sync_now(frm), __("Microsoft"));
			frm.add_custom_button(__("Re-authorize"), () => authorize(frm), __("Microsoft"));
			frm.add_custom_button(__("Disconnect"), () => disconnect(frm), __("Microsoft"));
		}
	},
});

function authorize(frm) {
	frappe.call({
		method: "frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.authorize_access",
		args: { calendar_name: frm.doc.name },
		freeze: true,
		freeze_message: __("Redirecting to Microsoft…"),
		callback: (r) => {
			if (r.message && r.message.url) window.location.href = r.message.url;
		},
	});
}

function test_conn(frm) {
	frappe.call({
		method: "frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.test_connection",
		args: { calendar_name: frm.doc.name },
		freeze: true,
		callback: (r) => {
			if (r.message && r.message.ok) {
				frappe.show_alert({ message: __("Connected as {0}", [r.message.account]), indicator: "green" });
			}
		},
	});
}

function sync_now(frm) {
	frappe.call({
		method: "frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.sync",
		args: { calendar_name: frm.doc.name },
		freeze: true,
		freeze_message: __("Syncing…"),
		callback: (r) => {
			const m = r.message || {};
			frappe.msgprint({
				title: __("Sync complete"),
				message: m.message || __("Pulled {0}, pushed {1}.", [m.pulled || 0, m.pushed || 0]),
				indicator: "green",
			});
			frm.reload_doc();
		},
	});
}

function disconnect(frm) {
	frappe.confirm(__("Disconnect this Microsoft account? Tokens will be cleared."), () => {
		frappe.call({
			method: "frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.disconnect",
			args: { calendar_name: frm.doc.name },
			freeze: true,
			callback: () => frm.reload_doc(),
		});
	});
}
