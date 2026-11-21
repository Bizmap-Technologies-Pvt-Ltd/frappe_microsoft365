"""Put all-day events pulled from Outlook back on their own dates.

Before 1.2.0 the pull converted an all-day event's midnight, which Graph sends in UTC, into the
site's timezone, so a pulled Christmas was stored as 25 Dec 05:30 on a Kolkata site and as
24 Dec 19:00 on a New York one.

The damage can be undone exactly, without asking Microsoft again. Every such event was stored as
a UTC midnight converted into the site's timezone, so converting it back gives that midnight —
and its date — precisely. That is also the test: only an event whose start and end both come
back as an exact UTC midnight is corrected, so an event already on its own dates, or one stored
while the site had another timezone, is left as it is. The second kind is still put right by
the next sync, which in 1.2.0 re-reads every calendar's window once and takes Outlook's dates.

Only events pulled from Outlook are touched. An all-day event created in Frappe was never
converted and is not Outlook's to correct. No network calls; safe to run again, because a
corrected event no longer converts back to a UTC midnight.
"""

import datetime

import frappe
from frappe.utils import get_datetime, get_system_timezone


def execute():
	if not frappe.db.has_column("Event", "custom_pulled_from_microsoft"):
		return

	rows = frappe.get_all(
		"Event",
		filters={"all_day": 1, "custom_pulled_from_microsoft": 1},
		fields=["name", "starts_on", "ends_on"],
		limit_page_length=0,
	)
	if not rows:
		return

	tz = get_system_timezone()
	for row in rows:
		corrected = corrected_dates(row.starts_on, row.ends_on, tz)
		if corrected:
			frappe.db.set_value(
				"Event",
				row.name,
				{"starts_on": corrected[0], "ends_on": corrected[1]},
				update_modified=False,
			)


def _utc_midnight(value, zone):
	"""The UTC date ``value`` (naive, in ``zone``) was converted from, if it was a UTC midnight."""
	from zoneinfo import ZoneInfo

	if not value:
		return None
	local = get_datetime(value).replace(tzinfo=zone)
	utc = local.astimezone(ZoneInfo("UTC"))
	if utc.time() != datetime.time.min:
		return None
	# Converted forward again it has to land on exactly the stored value, or this was not the
	# conversion that produced it (a time inside a daylight-saving change, say).
	if utc.astimezone(zone).replace(tzinfo=None) != get_datetime(value).replace(tzinfo=None):
		return None
	return utc.date()


def corrected_dates(starts_on, ends_on, tz_name):
	"""``(starts_on, ends_on)`` as the 1.2.0 pull stores them, or None when there is nothing to fix."""
	from zoneinfo import ZoneInfo

	zone = ZoneInfo(tz_name)
	first = _utc_midnight(starts_on, zone)
	if not first:
		return None
	after = _utc_midnight(ends_on, zone) if ends_on else None
	if ends_on and not after:
		return None

	last = after - datetime.timedelta(days=1) if after and after > first else first
	corrected = (
		datetime.datetime.combine(first, datetime.time.min),
		datetime.datetime.combine(last, datetime.time(23, 59, 59)),
	)
	stored = (get_datetime(starts_on), get_datetime(ends_on) if ends_on else None)
	if stored == corrected:
		return None
	return corrected
