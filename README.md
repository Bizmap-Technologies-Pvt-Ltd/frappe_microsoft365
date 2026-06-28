# Frappe Microsoft 365

Generic **Microsoft 365 integration for any Frappe / ERPNext site** — the Microsoft counterpart to
Frappe's built-in Google Calendar integration. Configure the Azure app registration once in
**Microsoft Settings**, let each user authorize a **Microsoft Calendar**, and the rest of the
platform gets real Outlook calendar sync, Teams meetings and transcripts through a small, reusable API.

## Features

- **Outlook calendar two-way sync** — Microsoft events ↔ Frappe `Event` (pull on a schedule,
  push on save/delete via doc events). Per-calendar `pull`/`push` toggles.
- **Teams meeting creation** — create a calendar-associated Teams online meeting and get back the
  join URL / event id / online-meeting id (calendar association is the precondition for transcripts).
- **Transcripts & recordings** — list and fetch Teams meeting transcripts (VTT) and recordings for
  meetings that are calendar-associated and not expired.
- **Secure by design** — MSAL auth-code flow, token refresh, CSRF state validation; secrets and
  tokens are stored as `Password` fields and never logged or returned to clients.

## Requirements

- A Frappe **v15 or v16** bench ([install guide](https://frappeframework.com/docs/user/en/installation)).
- Python ≥ 3.10. The only extra dependency is **`msal`** (declared in `pyproject.toml`, installed
  automatically by `bench get-app`).
- An **Azure AD (Microsoft Entra) app registration** — free; see [`docs/azure-setup.md`](docs/azure-setup.md).

## Install (existing bench)

```bash
cd /path/to/your/bench
bench get-app https://github.com/Bizmap-Technologies-Pvt-Ltd/frappe_microsoft365
bench --site your-site.localhost install-app frappe_microsoft365
bench --site your-site.localhost migrate
```

If `msal` didn't get pulled in automatically: `./env/bin/pip install msal`.

## Quick start from scratch (try it on a fresh local site)

```bash
# 1. (one-time) install bench + Frappe — see the Frappe install guide above, then:
bench new-site m365.localhost --admin-password admin
bench --site m365.localhost set-config developer_mode 1

# 2. get + install this app
bench get-app https://github.com/Bizmap-Technologies-Pvt-Ltd/frappe_microsoft365
bench --site m365.localhost install-app frappe_microsoft365
bench --site m365.localhost migrate

# 3. run it
bench start          # then open http://m365.localhost:8000/app
```

Now do the **Azure app registration** (5 minutes) following [`docs/azure-setup.md`](docs/azure-setup.md),
then configure (below). The one value Azure needs from you is the **Redirect URI**:

```
http://m365.localhost:8000/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback
```

(Use your real site URL/port. `http://...localhost` is accepted by Azure for local testing.)

## Configure

1. **Azure** → register an app, add the Redirect URI above (Web platform), create a client secret,
   and grant these **Delegated** Microsoft Graph permissions, then *Grant admin consent*:
   `offline_access openid profile User.Read Calendars.ReadWrite OnlineMeetings.ReadWrite OnlineMeetingTranscript.Read.All`
2. **Frappe Desk → Microsoft Settings** → paste Tenant ID, Client ID, Client Secret, the Redirect URI →
   tick **Enabled** → Save.
3. **Microsoft Calendar** → New → give it a name → Save → click **Authorize Microsoft Access** →
   sign in. The account email + default calendar fill in automatically. Use **Test Connection** /
   **Sync Now** to verify.

Full walkthrough with screenshots-worthy detail: [`docs/azure-setup.md`](docs/azure-setup.md).

## How it's consumed (for app developers)

Other apps depend on this app and call its utilities (they never re-implement Graph):

| Purpose | Function |
| --- | --- |
| Settings / liveness | `frappe_microsoft365.microsoft_graph.get_settings()` |
| Authorize / callback / disconnect / test | `…doctype.microsoft_calendar.microsoft_calendar` → `authorize_access(calendar_name)`, `callback`, `test_connection(calendar_name)`, `disconnect(calendar_name)` |
| Create a Teams meeting | `frappe_microsoft365.microsoft_meetings.create_meeting(calendar_name, subject, start_datetime, end_datetime, attendees=None, body=None, create_calendar_event=True)` → `{event_id, web_link, join_url, online_meeting_id}` |
| Transcripts for a join URL | `frappe_microsoft365.microsoft_transcripts.get_transcripts_for_join_url(calendar_name, join_url)` → `{online_meeting_id, transcripts, latest_vtt}` |
| Read calendar events | `frappe_microsoft365.microsoft_calendar_sync.fetch_events(calendar_name, start_datetime, end_datetime)` |

Example: **Bizmap OS** delegates its meeting/calendar features to these functions when this app is
installed and Microsoft Settings is configured, and falls back to stubs otherwise.

See [`docs/graph-api-reference.md`](docs/graph-api-reference.md) for the verified Graph
endpoint/permission contract, and [`docs/azure-setup.md`](docs/azure-setup.md) for Azure setup.

## Troubleshooting

- **Redirect URI mismatch (AADSTS50011):** the Redirect URI in Microsoft Settings must match the one
  registered in Azure *exactly* (scheme, host, port, path).
- **Transcript 403 / empty:** transcripts need `OnlineMeetingTranscript.Read.All` + admin consent, the
  meeting must be **calendar-associated** (create it via Microsoft Calendar / `create_meeting`, not a
  standalone online meeting) and **not expired**.
- **Token errors after a while:** click **Re-authorize** on the Microsoft Calendar; refresh tokens
  rotate automatically but a revoked consent requires re-auth.

## Contributing

```bash
cd apps/frappe_microsoft365
pre-commit install   # ruff, eslint, prettier, pyupgrade
```

Issues and PRs welcome.

## License

MIT
