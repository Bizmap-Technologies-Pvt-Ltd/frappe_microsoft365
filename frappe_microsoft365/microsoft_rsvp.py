"""Reply to a Microsoft calendar invitation from Frappe.

Graph exposes the three replies as POST actions on the event itself —
``/me/events/{id}/accept``, ``/decline`` and ``/tentativelyAccept``. Each needs only the
delegated ``Calendars.ReadWrite`` permission the calendar sync already holds, so replying
adds no new Azure consent and works on any calendar that is already authorized.

Each action takes an optional ``{"comment": str, "sendResponse": bool}`` body and answers
``202 Accepted`` with no content; ``graph_request`` turns an empty body into ``{}``, so
there is nothing to parse and success is simply "it did not raise".
"""

import frappe
from frappe import _
from frappe.utils import cint

from frappe_microsoft365 import microsoft_graph as graph
from frappe_microsoft365.microsoft_calendar_sync import _check_owner

#: response name -> (Graph action segment, the responseStatus value Microsoft is left with).
#: The segment casing is Microsoft's, not ours: /tentativelyAccept is camelCase while
#: /accept and /decline are not, and Graph 404s on the wrong spelling.
RESPONSE_ACTIONS = {
	"accept": ("accept", "accepted"),
	"decline": ("decline", "declined"),
	"tentative": ("tentativelyAccept", "tentativelyAccepted"),
}


# POST-only: this replies in someone's real mailbox and writes the reply back locally.
# Frappe auto-commits on POST/PUT/DELETE/PATCH but never on GET, so over a GET the local
# update below would be silently discarded.
@frappe.whitelist(methods=["POST"])
def respond_to_event(event: str, response: str, comment: str | None = None, send_response: int = 1):
	"""Accept, decline or tentatively accept the Microsoft event behind a Frappe Event.

	``event`` is the Frappe Event name, ``response`` one of accept / decline / tentative.
	``send_response`` mirrors Graph's ``sendResponse``: 0 records the reply in the calendar
	without emailing the organizer, which is what "respond without sending" does in Outlook.

	Returns ``{"ok": True, "response": <responseStatus value>}``.
	"""
	action = RESPONSE_ACTIONS.get((response or "").strip())
	if not action:
		frappe.throw(_("Unknown response {0}. Use accept, decline or tentative.").format(response or ""))
	segment, resulting_status = action

	doc = frappe.get_doc("Event", event)
	# Whoever may read the Event may reply to it; Frappe's own Event rules already scope
	# that to the owner, shared users and listed participants.
	doc.check_permission("read")

	ms_id = doc.get("custom_microsoft_event_id")
	calendar_name = doc.get("custom_microsoft_calendar")
	if not ms_id or not calendar_name:
		frappe.throw(_("This event is not synced with a Microsoft calendar."))

	# Graph rejects a reply from the organizer of an event ("ErrorInvalidRequest"), which
	# tells the user nothing. responseStatus is Microsoft's own word for it, so say so here.
	if doc.get("custom_microsoft_my_response") == "organizer":
		frappe.throw(_("You organized this event, so there is nothing to reply to."))

	calendar = frappe.get_doc("Microsoft Calendar", calendar_name)
	# Replying writes into someone's real mailbox and can email the organizer in their
	# name, so the calendar-owner rule the rest of the app uses applies here too.
	_check_owner(calendar)
	if not calendar.enabled or not calendar.authorized:
		frappe.throw(_("The Microsoft Calendar for this event is disabled or not authorized."))

	body = {"sendResponse": bool(cint(send_response))}
	if comment:
		body["comment"] = comment

	graph.graph_request("POST", f"/me/events/{ms_id}/{segment}", calendar_name, json=body)

	# Reflect the reply locally straight away. The next delta pull carries the same value
	# back from Graph, but that is up to fifteen minutes off and the user just clicked the
	# button. update_modified=False keeps the sync's dirty-check honest: bumping `modified`
	# is what makes the push step re-patch an event forever.
	frappe.db.set_value(
		"Event", doc.name, "custom_microsoft_my_response", resulting_status, update_modified=False
	)

	return {"ok": True, "response": resulting_status}
