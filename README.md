# Frappe Microsoft 365

Generic Microsoft 365 integration for any Frappe / ERPNext site — the Microsoft counterpart to
the built-in Google Calendar integration. Configure the Azure app registration once in
**Microsoft Settings**, let each user authorize a **Microsoft Calendar**, and the rest of the
platform gets real Outlook calendar, Teams meetings and transcripts through a small, reusable API.

## Features

- **Outlook calendar two-way sync** — Microsoft events ↔ Frappe `Event` (pull on a schedule,
  push on save/delete via doc events). Per-calendar `pull`/`push` toggles.
- **Teams meeting creation** — create a calendar-associated Teams online meeting and get back the
  join URL / event id / online-meeting id (calendar association is the precondition for transcripts).
- **Transcripts & recordings** — list and fetch Teams meeting transcripts (VTT) and recordings for
  meetings that are calendar-associated and not expired.
- **Secure by design** — MSAL auth-code flow, token refresh, CSRF state validation; secrets and
  tokens are stored as `Password` fields and never logged or returned to clients.

## Install

```bash
cd $PATH_TO_YOUR_BENCH
bench get-app https://github.com/<org>/frappe_microsoft365 --branch version-16
bench --site <your-site> install-app frappe_microsoft365
```

Python dependency: **`msal`** (declared in `pyproject.toml`; installed automatically by
`bench get-app`. If needed: `./env/bin/pip install msal`).

## Configure

1. Create an Azure app registration and a client secret, and grant the delegated Graph
   permissions. Full step-by-step: **[`docs/azure-setup.md`](docs/azure-setup.md)**.
2. In Frappe: **Microsoft Settings** → paste Tenant ID, Client ID, Client Secret → **Enable**.
3. Create a **Microsoft Calendar** record (one per user) and click **Authorize** to complete the
   Microsoft sign-in. The account email and default calendar are filled in automatically.

## How it's consumed

Other apps depend on this app and call its utilities (they never re-implement Graph):

| Purpose | Function |
| --- | --- |
| Settings / liveness | `frappe_microsoft365.microsoft_graph.get_settings()` |
| Authorize / callback / disconnect / test | `frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar` → `authorize_access(calendar_name)`, `callback`, `test_connection(calendar_name)`, `disconnect(calendar_name)` |
| Create a Teams meeting | `frappe_microsoft365.microsoft_meetings.create_meeting(calendar_name, subject, start_datetime, end_datetime, attendees=None, body=None, create_calendar_event=True)` → `{event_id, web_link, join_url, online_meeting_id}` |
| Transcripts for a join URL | `frappe_microsoft365.microsoft_transcripts.get_transcripts_for_join_url(calendar_name, join_url)` → `{online_meeting_id, transcripts, latest_vtt}` |
| Read calendar events | `frappe_microsoft365.microsoft_calendar_sync.fetch_events(calendar_name, start_datetime, end_datetime)` |

For example, **Bizmap OS** delegates its `graph_service` to these functions when this app is
installed and Microsoft Settings is enabled+configured, and falls back to stubs otherwise.

See **[`docs/graph-api-reference.md`](docs/graph-api-reference.md)** for the verified Graph
endpoint/permission contract.

## Contributing

This app uses `pre-commit` (ruff, eslint, prettier, pyupgrade):

```bash
cd apps/frappe_microsoft365
pre-commit install
```

## License

MIT
