# Frappe Microsoft 365

Generic **Microsoft 365 integration for any Frappe / ERPNext site** — the Microsoft counterpart to
Frappe's built-in Google Calendar integration. Configure the Azure app registration once in
**Microsoft Settings**, let each user authorize a **Microsoft Calendar**, and the rest of the
platform gets real Outlook calendar sync, Teams meetings and transcripts through a small, reusable API.

## Features

- **Outlook calendar two-way sync** — Microsoft events ↔ Frappe `Event` (delta pull on a
  schedule, push on save/delete via doc events). Per-calendar `pull`/`push` toggles.
  Handles recurring series, deletions, paging and throttling; see *How the sync behaves*.
- **Teams meeting creation** — create a calendar-associated Teams online meeting and get back the
  join URL / event id / online-meeting id (calendar association is the precondition for transcripts).
- **Transcripts & recordings** — list and fetch Teams meeting transcripts (VTT) and recordings for
  meetings that are calendar-associated and not expired.
- **Secure by design** — MSAL auth-code flow, token refresh, CSRF state validation; secrets and
  tokens are stored as `Password` fields and never logged or returned to clients.

## Requirements

- A Frappe **v15 or v16** bench ([install guide](https://frappeframework.com/docs/user/en/installation)).
- Python ≥ 3.10 (whatever your Frappe version requires: v15 runs on 3.10+, v16 needs 3.14).
  The only extra dependency is **`msal`** (declared in `pyproject.toml`, installed
  automatically by `bench get-app`).
- An **Azure AD (Microsoft Entra) app registration** — free; see [`docs/azure-setup.md`](docs/azure-setup.md).

## Compatibility

| Frappe / ERPNext | Branch | CI |
| --- | --- | --- |
| version-16 | `version-16` (or `main`) | every push and PR |
| version-15 | `version-15` (or `main`) | every push and PR |

One codebase supports both. `main` is the source of truth; `version-15` and `version-16` track
it so the usual `bench get-app --branch <version>` works. Nothing here is version-gated, and
every change is proven against **both** versions in CI (install, migrate, and the full test
suite on each) before it lands. [`docs/compatibility.md`](docs/compatibility.md) records the
API-by-API evidence, including the one place where version-specific code exists.

## Install (existing bench)

```bash
cd /path/to/your/bench

# either branch works; pick the one matching your bench
bench get-app https://github.com/Bizmap-Technologies-Pvt-Ltd/frappe_microsoft365 --branch version-16

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

## How the sync behaves

Worth knowing before you trust it with a real calendar:

- **Delta queries, not "modified since".** The pull runs `/me/calendarView/delta`, so recurring
  series arrive as individual occurrences with stable ids (no duplicates on later runs) and
  deletions arrive as `@removed` entries (no orphaned Frappe Events). The delta link is stored
  on the Microsoft Calendar and is the sync watermark.
- **A failed pull never advances the watermark.** If Graph errors halfway, the next run repeats
  that window instead of skipping it. The reason is written to **Last Sync Error** on the form.
- **Every page is read.** Collection reads follow `@odata.nextLink` to the end, so a calendar
  with hundreds of events in the window imports completely, not just the first page.
- **Origin decides who wins.** An Event that originated in Frappe keeps
  `custom_pulled_from_microsoft = 0` for life: the pull refreshes its subject, times and
  location from Outlook but never overwrites its description (Graph only returns a truncated
  plain-text preview) and never flips the flag, so later local edits keep syncing out.
  Events that originated in Microsoft are mirrors and are fully overwritten.
- **Nothing is saved when nothing changed**, so `modified` does not churn and the push step does
  not patch the same event back to Graph on every run.
- **The delta window** covers 30 days back and 180 days forward, and is re-initialised
  automatically when its far edge gets within 14 days.
- **Concurrency and throttling.** Each calendar syncs under a file lock, so a slow run is never
  overlapped by the next cron. A `429` is retried once after a short `Retry-After`; longer
  backoffs are left to the next scheduled run. A `410` (expired delta token) restarts a full
  sync automatically.
- **Timezones.** Graph is asked for UTC, and Windows timezone ids (`Pacific Standard Time`) are
  mapped to IANA rather than silently assumed to be UTC.

Access to `Microsoft Calendar` is granted to **System Manager** and **Desk User** (the same
pattern Frappe's Google Calendar uses), and each user only sees their own connection.

## Connection doctor

Connecting Frappe to Microsoft 365 has roughly fifteen steps across Azure, Exchange and
Frappe, and almost every mistake surfaces as the same unhelpful string — `AUTHENTICATE
failed`, `535 5.7.3`, `invalid_grant` — with no clue which step was wrong.

**Microsoft Settings → Troubleshoot** gives you three tools:

- **Run Diagnostics** inspects Microsoft Settings, the Connected App and every Email Account
  and reports what is actually wrong. It catches the failures people hit most: delegated
  scopes on an app-only flow (or the reverse), a missing `offline_access` scope — the reason
  a connection works for an hour then needs re-authorising forever — v1.0 endpoints, tenant
  mismatches between settings and endpoints, redirect-URI drift, IMAP without a folder, and
  the shared-mailbox identity conflict described below.
- **Explain an Error** turns a message from the Error Log into a cause and a next step.
- **Exchange Setup Script** generates the `New-ServicePrincipal` / `Add-MailboxPermission`
  commands for app-only mailbox access, looking the service principal up by AppId rather
  than asking you to copy an Object ID — Microsoft's own documentation warns that copying
  the one from the App Registration page (instead of the Enterprise Application page) causes
  authentication to fail with no useful error.

**The shared-mailbox identity conflict.** Microsoft requires the *shared mailbox address* in
the IMAP XOAUTH2 string but the *signing-in user* for SMTP. Frappe sends `login_id or
email_id` to both, so a single account cannot get incoming and outgoing right at the same
time — which is why "SMTP works but IMAP doesn't" recurs on the forum. The doctor flags the
configuration and suggests the two ways out: split incoming and outgoing into separate Email
Accounts, or use the app-only flow, where no user identity is involved.

### What this does NOT do

It does not send or receive mail, replace `Email Account`, or touch the email queue. Frappe's
own IMAP/SMTP + OAuth path is unchanged, several mail accounts keep working exactly as they
did, and uninstalling this app leaves your mail setup working. The doctor only reads
configuration and generates text.

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
