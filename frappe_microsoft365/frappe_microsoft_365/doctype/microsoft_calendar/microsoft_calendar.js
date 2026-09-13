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
			if (frm.doc.last_error) {
				// A failed sync leaves the watermark untouched and retries next run; say so.
				frm.dashboard.set_headline_alert(
					__("Last sync did not complete: {0}", [frm.doc.last_error]),
					"orange"
				);
			}
			frm.add_custom_button(__("Test Connection"), () => test_conn(frm), __("Microsoft"));
			frm.add_custom_button(__("Sync Now"), () => sync_now(frm), __("Microsoft"));
			frm.add_custom_button(__("Re-authorize"), () => authorize(frm), __("Microsoft"));
			frm.add_custom_button(__("Disconnect"), () => disconnect(frm), __("Microsoft"));
		}

		// frappe.realtime.on appends, so binding this inside refresh without dropping the
		// previous registration would announce one finished sync once per visit to the form.
		// Passing the same named function to off removes ours and only ours.
		// The handler is bound once by reference, so it cannot close over this frm — it is
		// parked here instead. cur_frm would have answered the same question, but it is
		// deprecated and returns whatever form is open, which is not necessarily this one.
		open_form = frm;
		frappe.realtime.off("microsoft365_sync_done", on_sync_done);
		frappe.realtime.on("microsoft365_sync_done", on_sync_done);
	},
});

// Turning push on is the only setting here that writes into someone's real Outlook calendar,
// and deleting a Frappe Event then deletes the Microsoft one. Say so before it is on.
frappe.ui.form.on("Microsoft Calendar", {
	push_to_microsoft_calendar(frm) {
		if (!frm.doc.push_to_microsoft_calendar || frm.doc.__push_warning_shown) return;

		frappe.confirm(
			__(
				"This writes into the real Outlook calendar of <b>{0}</b>. Frappe Events marked for sync will be created as Microsoft events, edits are sent across, and deleting a Frappe Event deletes the Microsoft one. Leave it off if you only want to read Outlook into Frappe.",
				[frm.doc.microsoft_user_email || frm.doc.user || __("this account")]
			),
			() => {
				frm.doc.__push_warning_shown = true;
			},
			() => {
				frm.set_value("push_to_microsoft_calendar", 0);
			}
		);
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

function sync_now(frm, run_inline) {
	frappe.call({
		method: "frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.sync",
		args: { calendar_name: frm.doc.name, run_inline: run_inline ? 1 : 0 },
		freeze: true,
		freeze_message: run_inline ? __("Syncing here — this may take a while…") : __("Syncing…"),
		callback: (r) => {
			const m = r.message || {};

			// Nothing is going to run the job: no worker, no scheduler, or Redis is gone. Say
			// so with the fix, and offer to do it here anyway — a long wait the person chose
			// beats work that silently never happens.
			if (m.blocked) {
				report_blocked(frm, m);
				return;
			}

			// Two honest shapes come back. Counts mean the sync ran inside this request; a
			// queued job means the server judged it too slow to hold one open. The inline
			// shape has no `queued` key at all, so its presence — not its value — decides.
			if (m.queued !== undefined) {
				report_queued(m);
				return;
			}

			frappe.msgprint({
				title: m.ok ? __("Sync complete") : __("Sync incomplete"),
				message:
					m.message ||
					__("Pulled {0}, deleted {1}, pushed {2}.", [
						m.pulled || 0,
						m.deleted || 0,
						m.pushed || 0,
					]),
				indicator: m.ok ? "green" : "orange",
			});
			frm.reload_doc();
		},
	});
}

// Queued work on a bench with no worker is the failure this whole guard exists for: enqueue
// succeeds, the page says "running in the background", and nothing ever runs. Name the fault,
// give the command, and leave a way to get the work done in the meantime.
function report_blocked(frm, m) {
	const health = m.health || {};
	const lines = []
		.concat(health.reasons || [], health.fixes || [])
		.map((line) => `<li>${frappe.utils.escape_html(line)}</li>`)
		.join("");

	const dialog = new frappe.ui.Dialog({
		title: __("Background jobs are not running"),
		indicator: "red",
		primary_action_label: __("Run now anyway"),
		primary_action: () => {
			dialog.hide();
			// Explicitly chosen, so the slowness check is skipped. It may still time out on a
			// first sync; the server keeps going and the next attempt picks up where it left off.
			sync_now(frm, true);
		},
		secondary_action_label: __("Cancel"),
	});
	dialog.set_message(
		__("This sync was not queued, because nothing would have run it.") +
			(lines ? `<ul>${lines}</ul>` : "")
	);
	dialog.show();
}

// The request is already answered by the time this runs, so the freeze is gone and the page
// is usable again; the counts arrive later on microsoft365_sync_done.
function report_queued(m) {
	if (!m.queued) {
		frappe.msgprint({
			title: __("Sync not started"),
			message: __("The background job could not be queued. Try again in a moment."),
			indicator: "red",
		});
		return;
	}

	// Why, not just that: "running in the background" on its own reads like the button
	// failed. The reasons are phrased by the server, so they are shown rather than built.
	const reasons = (m.reasons || [])
		.map((reason) => `<li>${frappe.utils.escape_html(reason)}</li>`)
		.join("");

	frappe.msgprint({
		title: __("Syncing in the background"),
		message:
			__(
				"This one is slow enough to time the page out, so it is running as a background job. The counts will appear here when it finishes."
			) + (reasons ? `<ul>${reasons}</ul>` : ""),
		indicator: "blue",
	});
}

//: The Microsoft Calendar currently on screen, refreshed with the form. The realtime handler
//: below is a single stable reference — that is what lets it be unbound before it is rebound —
//: so the form it should act on cannot travel in a closure and is kept here instead.
let open_form = null;

// Where a backgrounded sync lands, minutes after the button was pressed. Bound in refresh so
// it is always live on an open form, and only ever once — see the off/on pair there.
function on_sync_done(data) {
	const frm = open_form;
	if (!data || !frm || frm.doctype !== "Microsoft Calendar" || frm.doc.name !== data.calendar) return;

	frappe.show_alert({
		message:
			data.message ||
			__("Pulled {0}, deleted {1}, pushed {2}.", [
				data.pulled || 0,
				data.deleted || 0,
				data.pushed || 0,
			]),
		indicator: "green",
	});

	// Long enough has passed that the person may be mid-edit by now, and reloading would throw
	// their unsaved changes away over a watermark they would see on the next save anyway.
	if (!frm.is_dirty()) frm.reload_doc();
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
