<div align="center">
	<img src=".github/logo.png" height="110" alt="Frappe Microsoft 365">
	<h2>Frappe Microsoft 365</h2>
	<p><b>Outlook calendar, Teams meetings and Microsoft setup that actually explains itself</b></p>

![Frappe](https://img.shields.io/badge/Frappe-v15%20%7C%20v16-2b3a8c)
![License](https://img.shields.io/badge/license-MIT-2b3a8c)
![Tests](https://img.shields.io/badge/tests-191-2b3a8c)

</div>

Generic **Microsoft 365 integration for any Frappe / ERPNext site** — the Microsoft counterpart to
Frappe's built-in Google Calendar integration. Register one Azure application, tick the
capabilities you actually want in **Microsoft Settings**, and let each user authorize their own
**Microsoft Calendar**.

You get Outlook calendar sync in both directions, Teams meetings created from a Frappe Event,
invitations you can reply to without leaving Frappe, and meeting transcripts — plus a
diagnostics tool for the part everyone loses time on, which is getting the connection to work
in the first place.

## Features

- **Outlook calendar, two ways** — Microsoft events ↔ Frappe `Event`. Delta pull on a schedule,
  push on save and delete. Recurring series arrive as individual occurrences without
  duplicating, deletions propagate, and a failed pull never silently skips a window.
  See *How the sync behaves*.
- **Teams meetings from a Frappe Event** — an **Add Teams meeting** tickbox, the same idea as
  Outlook's own. The join link comes back onto the Event, and a meeting organised in Outlook
  keeps its link when it syncs in, so people can join from either side. No extra Azure
  permission.
- **Attendees and RSVP** — organiser, the attendee list with everyone's reply, and your own
  response status on the Event, plus **Accept / Tentative / Decline** buttons on invitations you
  received. Frappe's event participants go out as Outlook attendees.
- **Transcripts and recordings land on the Event** — after a Teams meeting, **Get Transcript**
  attaches the VTT to the Event, where it is searchable and yours to keep. A recording is not
  copied into your file store — it stays with Microsoft and **Download Recording** streams it
  through Frappe on demand. When Microsoft's retention has already removed either, you are told
  that plainly. See *Transcripts and recordings*.
- **Take only the parts you want** — calendar, Teams meetings, transcripts, mail and sign-in are
  independent. The Azure scopes are **derived from what you tick** and shown before you
  authorise, so the permissions you grant in Azure and the ones the app requests cannot drift
  apart. Setting up mail or sign-in creates what each needs and never modifies anything that
  already exists.
- **A doctor for when it breaks** — connecting Frappe to Microsoft 365 takes fifteen-odd steps
  and almost every mistake surfaces as `AUTHENTICATE failed` or `535 5.7.3`. **Run Diagnostics**
  reads your configuration and names the failing step; **Explain an Error** decodes a message
  from the Error Log into a cause and a fix.
- **Secure by design** — MSAL auth-code flow, token refresh, CSRF state validation; secrets and
  tokens are stored as `Password` fields and never logged or returned to clients. The two
  actions with consequences outside Frappe warn and confirm first.

<details open>
<summary><b>View Screenshots</b></summary>

<br />

**Join a meeting, reply to an invitation, and see who else has — from the Event itself**
![Event with a Teams meeting](.github/screenshots/06-event-teams-meeting.png)

**Tick only the capabilities you want; the Azure permissions follow from what you tick**
![Microsoft Settings](.github/screenshots/01-microsoft-settings.png)

**Set Up shows a plan first, and never modifies anything that already exists**
![Set Up plan](.github/screenshots/02-setup-plan.png)

**The connection doctor names the failing step instead of leaving you with `AUTHENTICATE failed`**
![Connection doctor](.github/screenshots/03-connection-doctor.png)

**Exchange setup script for shared mailboxes, with the service principal looked up by AppId**
![Exchange setup script](.github/screenshots/04-exchange-setup-script.png)

**A per-user connection, with its own pull and push toggles**
![Microsoft Calendar](.github/screenshots/05-microsoft-calendar.png)

</details>

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

## Frappe Cloud

The app declares its supported Frappe range in `pyproject.toml`, which is what Frappe Cloud
reads when you add it to a bench:

```toml
[tool.bench.frappe-dependencies]
frappe = ">=15.0.0-dev,<17.0.0"
```

Without that section Frappe Cloud refuses the app with *"Could not find a compatible Frappe
version in pyproject.toml"*. The lower bound is anchored at `-dev` because pre-release builds
report versions like `15.0.0-dev`, which sort **below** `15.0.0` under semver.

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

1. **Frappe Desk → Microsoft Settings** → tick the capabilities you want. The **Delegated Scopes**
   section then shows the exact permission list to grant — for calendar only that is
   `offline_access openid profile User.Read Calendars.ReadWrite`, and nothing more.
2. **Azure** → register an app, add the Redirect URI above (Web platform), create a client secret,
   and grant exactly those **Delegated** Microsoft Graph permissions, then *Grant admin consent*.
   Back in **Microsoft Settings**, paste Tenant ID, Client ID, Client Secret, the Redirect URI →
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

## Set up only what you want

Calendar, Teams meetings, transcripts, mail and sign-in are **independent**. Tick the ones you
want in Microsoft Settings and leave the rest alone — nothing is created for a capability you
did not ask for, and none of them depend on each other. Use the calendar without mail, mail
without sign-in, sign-in on its own; all valid.

**Set Up** shows a plan first: what will be created, what already exists, which Azure
permissions each capability needs, and only then offers to apply it.

| Capability | What gets created | Azure permissions |
| --- | --- | --- |
| Outlook calendar | Nothing — this app talks to Graph directly | Graph delegated: `User.Read`, `Calendars.ReadWrite`, `offline_access` |
| Standalone Teams meetings | Nothing — this app talks to Graph directly | Graph delegated: `User.Read`, `OnlineMeetings.ReadWrite`, `offline_access` |
| Meeting transcripts and recordings | Nothing — this app talks to Graph directly | Graph delegated: `User.Read`, `OnlineMeetingTranscript.Read.All`, `OnlineMeetingRecording.Read.All`, `offline_access` |
| Outlook mail | A `Connected App` for Frappe's Email Account | Exchange delegated: `IMAP.AccessAsUser.All`, `SMTP.Send`, `offline_access` — or the `.default` app-only scope for shared mailboxes |
| Sign in with Microsoft | A `Social Login Key` | Graph delegated: `openid`, `email`, `profile` |

### The scopes follow the tickboxes

**You do not write a scope list.** Microsoft Settings derives the delegated scopes from the
capabilities above and shows the exact result under **Delegated Scopes**, so what you grant in
Azure and what sign-in asks for cannot drift apart.

The split is what makes that work. **Outlook calendar** is `Calendars.ReadWrite` and nothing
else — including the **Add Teams meeting** tickbox on an Event, because Microsoft mints the
Teams link as part of the event rather than as a separate meeting. `OnlineMeetings.ReadWrite`
is only for meetings created outside a calendar event and for resolving a join URL back to a
meeting; `OnlineMeetingTranscript.Read.All` and `OnlineMeetingRecording.Read.All` are only
for transcripts and recordings. All are permissions tenants routinely refuse — reading the
words and reading the video are consented separately — which is why wanting a calendar no
longer asks for any of them.

Transcripts build on standalone meetings, so their tickbox appears only once that one is on,
and turning that one back off turns transcripts off with it rather than leaving a hidden
capability asking Azure for consent.

`offline_access`, `openid` and `profile` are added automatically at sign-in and never belong in
a scope list of your own.

**Override Delegated Scopes** is the escape hatch, empty by default: fill it in only when your
tenant consents to a hand-picked list and sign-in has to ask for exactly that. An override is
used verbatim, and the doctor warns if it omits something a ticked capability needs.

Changing any of this affects new sign-ins only. Tokens carry the permissions that were
consented when they were issued, so after a change the doctor tells you which connections need
**Re-authorize** rather than letting the new feature fail with a bare 403.

**It never overwrites anything.** If a record already exists it is left exactly as it is and
reported, with any drift from Microsoft Settings spelled out, so a setup someone tuned by hand
survives untouched. Running Set Up twice does nothing the second time.

Two details it gets right that are easy to miss by hand:

- Frappe's built-in Office 365 sign-in provider defaults to the `/common/` authority and the
  **v1.0** endpoints, neither of which works with a single-tenant app registration. Provisioning
  writes tenant-specific v2.0 endpoints instead.
- The `Connected App` redirect URI is computed by Frappe from the record name, so it is a
  *different* endpoint from this app's callback and cannot be known before the record exists.
  Azure needs **both** registered; the doctor prints the exact URI to add. A missing one shows
  up later as `AADSTS50011`.

## Connection doctor

Connecting Frappe to Microsoft 365 has roughly fifteen steps across Azure, Exchange and
Frappe, and almost every mistake surfaces as the same unhelpful string — `AUTHENTICATE
failed`, `535 5.7.3`, `invalid_grant` — with no clue which step was wrong.

**Microsoft Settings → Troubleshoot** gives you three tools:

- **Run Diagnostics** inspects Microsoft Settings, the Connected App and every Email Account
  and reports what is actually wrong. It catches the failures people hit most: delegated
  scopes on an app-only flow (or the reverse), a missing `offline_access` scope — the reason
  a connection works for an hour then needs re-authorising forever — v1.0 endpoints, tenant
  mismatches between settings and endpoints, redirect-URI drift, IMAP without a folder, a
  scope override that leaves out something a ticked capability needs, scopes that changed
  after connections were authorised (their tokens predate the change and have to be renewed),
  and the shared-mailbox identity conflict described below.
- **Explain an Error** turns a message from the Error Log into a cause and a next step.
- **Wrong-box mistakes are caught as you type them.** Azure shows a secret's **Value** next to
  its **Secret ID**, and only the Value works; pasting the ID is the commonest setup mistake
  there is, and Microsoft only says so at sign-in, as `AADSTS7000215`. A secret that is a GUID
  is refused on save, because a secret value never is one. The same shape checks cover the
  client id, the tenant id and the redirect URI.
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

### Teams meetings from a Frappe Event

An Event carries an **Add Teams meeting** tickbox, the same idea as Outlook's own toggle.
Tick it, save, and the event is created in Outlook as a Teams meeting. A **Join Meeting**
button then appears at the top of the Event, with **Open in Outlook** beside it.

It works in both directions: a Teams meeting organised in Outlook keeps its join link when it
syncs into Frappe, so people can join from either side.

No extra Azure permission is needed. Microsoft creates the meeting as part of the event, so
`Calendars.ReadWrite` covers it — the **Outlook calendar** capability on its own is enough.
`OnlineMeetings.ReadWrite` is only required for standalone meetings and transcripts, which are
separate tickboxes precisely so a calendar-only setup never has to ask for them.

**One limitation, enforced rather than explained away:** Microsoft cannot turn an existing
online meeting back into a plain event. So the tickbox **locks itself once the meeting exists**
rather than sitting there doing nothing when you untick it. To remove a meeting, delete the
event and create it again. The app only ever sends `isOnlineMeeting: true`.

### Attendees and RSVP

An invitation that lands in someone's Outlook is usable from Frappe. Each synced Event
carries three read-only fields, refreshed by every sync:

- **Organizer** — the Microsoft account that created the event.
- **Attendees** — one line per invitee: `Asha Rao <asha@example.com> — accepted`. Rooms and
  equipment are labelled with their type, because a room declining is a different problem
  from a person declining.
- **My Response** — your own reply, stored exactly as Microsoft words it (`accepted`,
  `declined`, `tentativelyAccepted`, `notResponded`, `organizer`, `none`).

On an event you were invited to, the Event form grows **Accept**, **Tentative** and
**Decline** buttons under a *Microsoft* menu. Each one offers an optional comment for the
organizer and a tickbox to reply without emailing anyone — the same choice Outlook gives
you. The reply goes straight to Microsoft and the Event updates immediately rather than
waiting for the next scheduled sync. The buttons stay hidden on events you organized
yourself, because there is nothing to reply to.

No new Azure permission is needed: `/me/events/{id}/accept`, `/decline` and
`/tentativelyAccept` all run on the delegated `Calendars.ReadWrite` the calendar sync
already holds.

Going the other way, an Event's **participants** are sent to Outlook as attendees when the
event is pushed. A participant whose email cannot be resolved — no address on the row and
none on the record it links to — is left out rather than sent as something Microsoft would
reject. If nobody resolves, the attendee list is omitted from the request entirely: Graph
reads an empty list as *remove everyone*, and that would silently uninvite people who were
added in Outlook.

**One limitation.** Zoom and Google Meet links are **not** extracted. Microsoft only fills
in the structured `onlineMeeting` property for its own Teams meetings; a third-party link
is loose text in the event body, and guessing at it would produce wrong links more often
than right ones. Open the event in Outlook for those.

### Transcripts and recordings

Once a Teams meeting has finished, the Event grows **Get Transcript** under the *Microsoft*
menu. It fetches the latest transcript and **attaches it to the Event as a `.vtt` file**, and
notes any recordings it found. Fetching again replaces the attachment rather than piling up
copies.

The two are deliberately handled differently:

- **The transcript is attached.** A VTT is a few kilobytes, it is the part people actually
  search and quote, and once it is a File on the Event it survives whatever Microsoft later
  does with its own copy.
- **The recording is not.** A Teams recording routinely runs to hundreds of megabytes, and
  copying one into the site's file store per meeting is a bad trade. Microsoft keeps it; the
  Event records that it exists, and **Download Recording** fetches it when someone asks.

**Why the recording isn't just a link.** Graph's `recordingContentUrl` is an API endpoint that
requires a bearer token — paste it into a browser and you get `401`, not a video. So the
download streams through Frappe, authorised by Frappe's own permissions, and the token never
leaves the server.

**Expiry is reported, not guessed.** Teams recordings and transcripts are deleted under your
tenant's retention policy, and Graph publishes no expiry date for them — the `callRecording`
resource simply has no such property. Rather than invent a countdown, a fetch that comes back
empty says what is actually true: Microsoft no longer has this, and why.

Three read-only fields carry the state: the online meeting id (resolved once from the join
link and kept), when the transcript was last fetched, and one line per recording.

**Requirements.** This is the **Transcripts** capability, which needs
`OnlineMeetings.ReadWrite` — a transcript is addressed by online meeting id, and resolving
that from a join link needs it. Recordings additionally need
`OnlineMeetingRecording.Read.All` and a Teams licence that records; if that permission is
missing the transcript still lands rather than the whole action failing. Only meetings created
**as calendar events** have transcripts — standalone Teams meetings do not, which is why the
app creates them the calendar way.

### Sensitive actions are called out before they happen

Three things have consequences outside Frappe, and none of them happen quietly:

- **Writing into a real calendar is opt-in.** `Push Frappe events to Microsoft` is **off by
  default**; pull is the safe direction. Turning push on asks you to confirm, and says plainly
  that deleting a Frappe Event will delete the Microsoft one.
- **The Exchange setup script** grants the application standing access to the mailboxes you
  list, readable with nobody signed in. The dialog explains that and will not generate the
  script until you confirm; the app never runs it, and **the script ends with the commands
  that undo it**, commented out so nothing reverses by accident.
- **A time range Microsoft would reject is caught on save**, not after the fact. Frappe does
  not enforce that an Event ends after it starts and pre-fills both times from the current
  moment, so an event saved without touching them can end before it begins. You are told while
  you can still fix it. Data arriving *from* Outlook is never second-guessed this way.

Sync also refuses to be half-configured: ticking **Sync with Microsoft Calendar** requires
choosing a connection, and enabling the integration requires the Azure credentials, so nothing
saves in a state that silently does nothing.

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
| Reply to an invitation | `frappe_microsoft365.microsoft_rsvp.respond_to_event(event, response, comment=None, send_response=1)` — `response` is `accept`, `decline` or `tentative` |

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
