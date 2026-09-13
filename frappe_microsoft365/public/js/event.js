// Accept / Decline / Tentative for a Microsoft invitation, from the Event form.
//
// The buttons only show on an event that arrived from Outlook and that this user did not
// organize: Graph rejects a reply from the organizer, and an event Frappe created has
// nobody to reply to.

frappe.provide("frappe_microsoft365");

// Graph's own responseStatus value for "I am the organizer of this".
frappe_microsoft365.ORGANIZER_RESPONSE = "organizer";

frappe_microsoft365.is_microsoft_invitee = function (doc) {
	if (!doc || !doc.custom_microsoft_event_id) return false;
	// custom_pulled_from_microsoft = 0 means Frappe pushed this event out, so this user is
	// the organizer by construction.
	if (!doc.custom_pulled_from_microsoft) return false;
	return doc.custom_microsoft_my_response !== frappe_microsoft365.ORGANIZER_RESPONSE;
};

frappe_microsoft365.rsvp = function (frm, response, label) {
	frappe.prompt(
		[
			{
				fieldname: "comment",
				fieldtype: "Small Text",
				label: __("Comment (optional)"),
				description: __("Sent to the organizer along with your reply."),
			},
			{
				fieldname: "send_response",
				fieldtype: "Check",
				label: __("Send a reply to the organizer"),
				default: 1,
				description: __("Untick to record your answer in the calendar without emailing anyone."),
			},
		],
		(values) => {
			frappe.call({
				method: "frappe_microsoft365.microsoft_rsvp.respond_to_event",
				args: {
					event: frm.doc.name,
					response: response,
					comment: values.comment,
					send_response: values.send_response,
				},
				freeze: true,
				freeze_message: __("Sending your reply…"),
				callback: () => {
					frappe.show_alert({
						message: __("Reply sent: {0}", [label]),
						indicator: "green",
					});
					// The whitelisted method writes the new response straight onto the Event,
					// so reload rather than leaving a stale value on screen.
					frm.reload_doc();
				},
			});
		},
		__("{0} this invitation", [label]),
		label
	);
};

frappe.ui.form.on("Event", {
	refresh(frm) {
		if (frm.is_new()) return;

		// Joining is the thing people come to this page to do, so it is a primary button
		// rather than a 500-character URL they have to find, select and paste.
		if (frm.doc.custom_teams_join_url) {
			frm.add_custom_button(__("Join Meeting"), () => {
				window.open(frm.doc.custom_teams_join_url, "_blank", "noopener");
			}).addClass("btn-primary");
		}

		if (frm.doc.custom_microsoft_web_link) {
			frm.add_custom_button(
				__("Open in Outlook"),
				() => window.open(frm.doc.custom_microsoft_web_link, "_blank", "noopener"),
				__("Microsoft")
			);
		}

		// A finished Teams meeting can bring its transcript and recording back.
		if (frm.doc.custom_teams_join_url) {
			const past = frm.doc.ends_on && frappe.datetime.now_datetime() > frm.doc.ends_on;
			if (past) {
				frm.add_custom_button(
					__("Get Transcript"),
					() => frappe_microsoft365.fetch_artifacts(frm),
					__("Microsoft")
				);
			}
			if (frm.doc.custom_microsoft_recordings) {
				frm.add_custom_button(
					__("Download Recording"),
					() => {
						// Streamed through Frappe: Graph's own URL needs a bearer token, so a
						// browser given it would get 401 rather than a video.
						window.open(
							"/api/method/frappe_microsoft365.microsoft_meeting_artifacts.download_recording" +
								"?event=" + encodeURIComponent(frm.doc.name),
							"_blank",
							"noopener"
						);
					},
					__("Microsoft")
				);
			}
		}

		if (!frappe_microsoft365.is_microsoft_invitee(frm.doc)) return;

		const group = __("Microsoft");
		[
			["accept", __("Accept")],
			["tentative", __("Tentative")],
			["decline", __("Decline")],
		].forEach(([response, label]) => {
			frm.add_custom_button(label, () => frappe_microsoft365.rsvp(frm, response, label), group);
		});
	},
});

frappe_microsoft365.fetch_artifacts = function (frm) {
	frappe.call({
		method: "frappe_microsoft365.microsoft_meeting_artifacts.fetch_meeting_artifacts",
		args: { event: frm.doc.name },
		freeze: true,
		freeze_message: __("Asking Microsoft…"),
		callback: (r) => {
			const result = r.message || {};
			frappe.msgprint({
				title: __("Meeting files"),
				message: result.message,
				indicator: result.transcript ? "green" : "orange",
			});
			frm.reload_doc();
		},
	});
};
