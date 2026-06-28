# frappe_microsoft365 — build progress

## M1 — Foundation (done, pre-existing)
- `Microsoft Settings` (Single) + `Microsoft Calendar` (per-user) doctypes.
- MSAL OAuth round-trip (authorize/callback), token storage + refresh.
- `microsoft_graph.py`: `graph_request`, `get_valid_access_token`, `whoami`, `MsGraphError`.
- Per-user row scoping (`permissions.py`).

## M2 — Two-way Event sync (done)
- `setup.py` `create_event_custom_fields()` adds idempotent custom fields to `Event`:
  `custom_sync_with_microsoft_calendar`, `custom_microsoft_calendar`,
  `custom_microsoft_event_id` (read-only), `custom_pulled_from_microsoft` (read-only, hidden).
  Wired via `after_migrate` hook (`setup.after_migrate`).
- `microsoft_calendar_sync.py`:
  - `sync_calendar(calendar_name)` — pull (changed-since `last_sync`, else 30/60-day
    `calendarView`) + push (unsynced Frappe Events), updates `last_sync`. Per-item errors
    are logged and skipped so one bad event never aborts the batch.
  - `sync_all()` — scheduler entry; loops enabled+authorized calendars; no-ops when the
    integration is disabled.
  - `event_on_update` / `event_on_trash` doc_events — single-Event push/patch/delete,
    best-effort, guarded by `frappe.flags.in_microsoft_sync` to avoid pull/push echo.
  - `fetch_events(calendar_name, start, end)` — whitelisted, owner-checked, clean list
    (id, subject, start, end, join_url, web_link) for external consumers (Bizmap CRM).
- Hooks: `doc_events["Event"]`, `scheduler_events` cron `*/15 * * * *` -> `sync_all`.

## M3 — Teams meeting creation (done)
- `microsoft_meetings.py` `create_meeting(...)`:
  - Preferred: POST `/me/events` with `isOnlineMeeting/teamsForBusiness` (calendar-
    associated -> transcripts work). Resolves `online_meeting_id` via JoinWebUrl filter
    (best-effort). Returns `{event_id, web_link, join_url, online_meeting_id}`.
  - Alternate (`create_calendar_event=False`): POST `/me/onlineMeetings` (standalone, NOT
    calendar-associated -> transcripts unavailable; documented in the return payload).

## M4 — Transcripts + recordings (done)
- `microsoft_transcripts.py`:
  - `list_transcripts`, `get_transcript_content` (VTT/plain via `?$format=`),
    `get_transcripts_for_join_url` (resolve -> list -> latest VTT).
  - `list_recordings`, `get_recording_content` (licensing noted).
  - 403 -> clear message about required scope/consent + calendar-association/expiry.

## Whitelisted method paths (for wiring Bizmap CRM)
- `frappe_microsoft365.microsoft_calendar_sync.fetch_events`
- `frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.sync`
- `frappe_microsoft365.microsoft_meetings.create_meeting`
- `frappe_microsoft365.microsoft_meetings.resolve_online_meeting_id`
- `frappe_microsoft365.microsoft_transcripts.list_transcripts`
- `frappe_microsoft365.microsoft_transcripts.get_transcript_content`
- `frappe_microsoft365.microsoft_transcripts.get_transcripts_for_join_url`
- `frappe_microsoft365.microsoft_transcripts.list_recordings`
- `frappe_microsoft365.microsoft_transcripts.get_recording_content`

All whitelisted methods are owner-checked (caller must own the Microsoft Calendar, or be a
System Manager / Administrator).
