<div align="center">
	<img src=".github/logo.png" height="110" alt="Frappe Microsoft 365">
	<h1>Frappe Microsoft 365</h1>
	<p><b>Outlook calendar sync, Teams meetings and Microsoft sign-in for Frappe and ERPNext</b></p>

![Frappe](https://img.shields.io/badge/Frappe-v15%20%7C%20v16-2b3a8c)
![License](https://img.shields.io/badge/license-MIT-2b3a8c)
[![CI](https://github.com/Bizmap-Technologies-Pvt-Ltd/frappe_microsoft365/actions/workflows/ci.yml/badge.svg)](https://github.com/Bizmap-Technologies-Pvt-Ltd/frappe_microsoft365/actions/workflows/ci.yml)

[Setup](docs/setup.md) ·
[How it behaves](docs/how-it-works.md) ·
[Troubleshooting](docs/troubleshooting.md) ·
[FAQ](docs/faq.md)

</div>

Frappe ships a Google Calendar integration and nothing equivalent for Microsoft. This is that
counterpart: two-way Outlook calendar sync, Teams meetings created from a Frappe Event,
transcripts and recordings, Outlook mail over OAuth, and sign in with Microsoft.

Register one Azure application for the company, tick the capabilities you want, and each person
authorises their own connection. It follows the same model as Frappe's Google Calendar
integration — one app registration, one connection per user, two-way `Event` sync — against
Microsoft Graph.

<details open>
<summary><b>Screenshots</b></summary>

<br />

**Join a meeting, reply to an invitation, and see who else has — from the Event itself**
![Event with a Teams meeting](.github/screenshots/06-event-teams-meeting.png)

**Tick only the capabilities you want; the Azure permissions follow from what you tick**
![Microsoft Settings](.github/screenshots/01-microsoft-settings.png)

**The connection doctor names the failing step instead of leaving you with `AUTHENTICATE failed`**
![Connection doctor](.github/screenshots/03-connection-doctor.png)

</details>

## What it does

- **Outlook calendar, both ways.** Delta-based, so a recurring series arrives as individual
  occurrences without duplicating on later runs and deletions never leave orphans. A repeating
  Frappe Event goes back as a real Outlook series. Push is off by default, per connection.
- **Teams meetings from an Event.** A tickbox; the join link comes back onto the Event. Needs
  `Calendars.ReadWrite` and nothing else — Microsoft mints the link as part of the event.
- **The invitation side.** Organiser, the attendee list with everyone's reply, and
  Accept / Tentative / Decline on invitations you received.
- **Transcripts and recordings.** The transcript is attached to the Event as a `.vtt`. A
  recording stays with Microsoft and streams through Frappe on demand rather than filling your
  file store.
- **Sign in with Microsoft**, and **Outlook mail over OAuth** — it provisions the `Connected App`
  that Frappe's own `Email Account` needs, then stays out of the way.
- **A doctor for when it breaks.** Setup spans two Microsoft portals and almost every mistake
  surfaces as the same unhelpful string. **Run Diagnostics** names the failing step;
  **Explain an Error** decodes a message from the Error Log into a cause and a fix.

Each capability is independent, and the Azure scopes are **derived from what you tick**, so what
you grant in Azure and what sign-in asks for cannot drift apart.

## What it does not do

- It syncs the signed-in person's **default** calendar. Not a second calendar in the same
  mailbox, not a shared mailbox calendar, not a room calendar.
- It does not send or receive mail, and does not replace `Email Account`. Uninstalling leaves
  your mail setup working.
- Zoom and Google Meet links are not extracted — Microsoft only fills in the structured
  `onlineMeeting` property for its own meetings.
- Transcripts need a tenant switch Microsoft ships **off**, and recordings need their own
  permission and a licence that records. Both are covered in [Setup](docs/setup.md).

## Compatibility

| Frappe | Branch | Status |
| --- | --- | --- |
| v15 | `version-15` | supported |
| v16 | `version-16` | supported |

`main` carries the same code as both. Python 3.10+. `msal` is the only dependency that is not
already in a bench.

## Install

```bash
bench get-app https://github.com/Bizmap-Technologies-Pvt-Ltd/frappe_microsoft365
bench --site <your-site> install-app frappe_microsoft365
bench --site <your-site> migrate
```

Then follow [Setup](docs/setup.md) — the steps are in the order they have to happen, which
matters more than it sounds: a token carries only the permissions consented *before* it was
issued.

## Documentation

| | |
| --- | --- |
| [Setup](docs/setup.md) | Azure, the Teams admin center and Frappe, in order |
| [How it behaves](docs/how-it-works.md) | Sync rules, multi-user, Teams meetings, transcripts, the Python API |
| [Troubleshooting](docs/troubleshooting.md) | Every Microsoft error this app can produce, with its real cause |
| [FAQ](docs/faq.md) | What it does, what it costs, what it will not do |
| [Azure reference](docs/azure-setup.md) · [Graph endpoints](docs/graph-api-reference.md) | The detail behind both |

## Contributing

Issues and pull requests are welcome. Tests run against Frappe v15 and v16:

```bash
bench --site <your-site> run-tests --app frappe_microsoft365
```

If you find something wrong with the sync — a recurring series, a timezone, an event edited in
two places at once — an issue describing both calendars' states is the most useful thing you can
send.

## License

MIT
