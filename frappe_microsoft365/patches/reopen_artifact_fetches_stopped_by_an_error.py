"""Give meetings that 1.1.x stopped asking about after one error their transcripts back.

Before 1.2.0 the catch-up job spent an Event's whole day of attempts the moment anything went
wrong while looking its meeting up — a 503, a 429, a token refresh that failed for a moment —
and wrote nothing else: no status, no checked time, no meeting id. That signature (attempts at
the maximum, and all three empty) is only ever left by that path; every attempt that reached
Microsoft records when it ran. From the row alone the cause cannot be told apart from a
permanent one, so:

* Meetings that ended in the last 7 days are reopened. 1.2.0 retries a temporary failure and
  stops only on a permanent one, with a status saying which, so a reopened meeting either gets
  its transcript or says why not. Seven days is well inside the 60 days Graph keeps serving a
  meeting, and recent enough that the transcript is still wanted.
* Older ones are not asked about again, but are given a status saying what happened and that
  Get Transcript & Recording still works — so no Event is left stopped with nothing to say why.

No network calls. Idempotent: a row it has handled has a status and no longer matches.
"""

import frappe
from frappe import _
from frappe.utils import add_to_date, now_datetime

REOPEN_DAYS = 7


def execute():
	if not frappe.db.has_column("Event", "custom_microsoft_artifacts_attempts"):
		return
	from frappe_microsoft365.microsoft_meeting_artifacts import RETRY_MINUTES

	stopped = frappe.db.sql(
		"""
		select name, ends_on
		from `tabEvent`
		where coalesce(custom_teams_join_url, '') != ''
			and custom_microsoft_artifacts_attempts >= %(spent)s
			and coalesce(custom_microsoft_artifacts_status, '') = ''
			and custom_microsoft_artifacts_checked_on is null
			and coalesce(custom_microsoft_online_meeting_id, '') = ''
		""",
		{"spent": len(RETRY_MINUTES)},
		as_dict=True,
	)
	since = add_to_date(now_datetime(), days=-REOPEN_DAYS)
	for row in stopped:
		if row.ends_on and row.ends_on >= since:
			values = {
				"custom_microsoft_artifacts_attempts": 0,
				"custom_microsoft_artifacts_status": _(
					"An earlier version stopped checking this meeting after one error. It is being "
					"checked again."
				),
			}
		else:
			values = {
				"custom_microsoft_artifacts_status": _(
					"An earlier version stopped checking this meeting after one error, without "
					"saying which. Use Get Transcript & Recording to ask Microsoft again."
				),
			}
		frappe.db.set_value("Event", row.name, values, update_modified=False)
