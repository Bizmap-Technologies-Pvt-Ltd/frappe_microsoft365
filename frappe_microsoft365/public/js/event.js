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
