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
	custom_sync_with_microsoft_calendar(frm) {
		frappe_microsoft365.offer_a_connection(frm);
	},

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
			frappe_microsoft365.add_artifact_buttons(frm);
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

// Ticking the sync box makes the connection field mandatory. Somebody who has never connected
// therefore hits "Microsoft Calendar is required", with nothing on screen explaining what a
// Microsoft Calendar is or how to get one — a dead end at the exact moment they were trying to
// use the feature for the first time. Everyone in a company meets this once.
frappe_microsoft365.offer_a_connection = async function (frm) {
	if (!frm.doc.custom_sync_with_microsoft_calendar) return;
	if (frm.doc.custom_microsoft_calendar) return;

	// Permission-scoped by the server, so this only ever counts the person's own connections.
	const mine = await frappe.db.get_list("Microsoft Calendar", { limit: 1 });
	if (mine.length) return;

	const dialog = new frappe.ui.Dialog({
		title: __("Connect your Microsoft account"),
		indicator: "blue",
		primary_action_label: __("Connect now"),
		primary_action: () => {
			dialog.hide();
			// Their own connection, created for them: the doctype fills in the user, and
			// Authorize on the form is the only step left.
			frappe.new_doc("Microsoft Calendar", { account_name: frappe.session.user_fullname });
		},
		secondary_action_label: __("Not now"),
		secondary_action: () => {
			dialog.hide();
			frm.set_value("custom_sync_with_microsoft_calendar", 0);
		},
	});
	dialog.set_message(
		__(
			"You have not connected a Microsoft account yet. Each person connects their own, once — after that your events sync both ways and Teams meetings, transcripts and recordings work from here."
		)
	);
	dialog.show();
};

frappe_microsoft365.stored_recordings = function (frm) {
	try {
		return JSON.parse(frm.doc.custom_microsoft_recordings_data || "[]");
	} catch (e) {
		return [];
	}
};

// The buttons follow the state, so the form never offers to fetch what it already has, and
// never hides the manual check while Microsoft might still be processing.
frappe_microsoft365.add_artifact_buttons = function (frm) {
	const group = __("Microsoft");
	// Offered from the moment the meeting starts, not from the end of the slot it was booked
	// into. A one-minute call inside an eighty-minute booking has its transcript ready long
	// before the slot runs out, and hiding the button until then hides the feature.
	const started = !frm.doc.starts_on || frappe.datetime.now_datetime() > frm.doc.starts_on;
	if (!started) return;

	const recordings = frappe_microsoft365.stored_recordings(frm);
	const has_transcript = !!frm.doc.custom_microsoft_transcript_fetched_on;

	if (!has_transcript || !recordings.length) {
		// Named for what is still missing: "Get Transcript" on an event whose transcript is
		// already attached reads like the first fetch failed.
		const label = has_transcript
			? __("Check for Recording")
			: recordings.length
			? __("Check for Transcript")
			: __("Get Transcript & Recording");
		frm.add_custom_button(label, () => frappe_microsoft365.fetch_artifacts(frm), group);
	}

	if (recordings.length) {
		frm.add_custom_button(
			__("Download Recording"),
			() => frappe_microsoft365.download_recording(frm, recordings),
			group
		);
	}
};

frappe_microsoft365.download_recording = function (frm, recordings) {
	// Streamed through Frappe: Graph's own URL needs a bearer token, so a browser given it
	// would get 401 rather than a video.
	const fetch_one = (id) =>
		window.open(
			"/api/method/frappe_microsoft365.microsoft_meeting_artifacts.download_recording" +
				"?event=" + encodeURIComponent(frm.doc.name) +
				(id ? "&recording_id=" + encodeURIComponent(id) : ""),
			"_blank",
			"noopener"
		);

	if (recordings.length < 2) {
		fetch_one(recordings.length ? recordings[0].id : null);
		return;
	}

	// Teams splits a recording every 4 hours or 1.5 GB, so a long meeting has parts and the
	// person has to be asked which one they want.
	const dialog = new frappe.ui.Dialog({
		title: __("Which part?"),
		fields: [
			{
				fieldname: "part",
				fieldtype: "Select",
				label: __("Recording"),
				reqd: 1,
				options: recordings.map((r, i) => ({
					value: r.id,
					label: __("Part {0} of {1}", [i + 1, recordings.length]) +
						(r.created ? " — " + frappe.datetime.str_to_user(r.created.replace("T", " ").slice(0, 19)) : ""),
				})),
			},
		],
		primary_action_label: __("Download"),
		primary_action: (values) => {
			dialog.hide();
			fetch_one(values.part);
		},
	});
	dialog.show();
};

frappe_microsoft365.fetch_artifacts = function (frm) {
	frappe.call({
		method: "frappe_microsoft365.microsoft_meeting_artifacts.fetch_meeting_artifacts",
		args: { event: frm.doc.name },
		freeze: true,
		freeze_message: __("Asking Microsoft…"),
		callback: (r) => {
			const result = r.message || {};
			const landed = result.state === "complete" || result.state === "partial";
			frappe.msgprint({
				title: __("Meeting files"),
				message: result.message,
				indicator: landed ? "green" : "orange",
			});
			frm.reload_doc();
		},
	});
};
