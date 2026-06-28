"""Two-way calendar sync (Frappe Event <-> Microsoft Graph). Implemented in stage M2.

This stub keeps the Authorize/connect flow usable in M1 without erroring.
"""

import frappe


def sync_calendar(calendar_name=None):
	return {
		"ok": True,
		"pulled": 0,
		"pushed": 0,
		"message": "Calendar sync ships in stage M2. Authorization and connection are ready.",
	}
