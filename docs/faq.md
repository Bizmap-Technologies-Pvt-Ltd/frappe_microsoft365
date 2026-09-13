# FAQ

Short answers to what people ask before and during setup, including the ones where the answer
is no. Depth lives in the [README](../README.md), [`azure-setup.md`](azure-setup.md) and
[`compatibility.md`](compatibility.md) — this page points at them rather than repeating them.

---

## Does ERPNext sync with Outlook calendar?

With this app, yes, in both directions. Microsoft events become Frappe `Event` records and back
again: a delta pull runs every 15 minutes, and a push sends Frappe events to Outlook when they
are saved. Push is off by default on every connection; pull is the safe direction. See
[How the sync behaves](../README.md#how-outlook-calendar-sync-behaves).

## How do I connect ERPNext to Office 365?

Install the app, register one Azure (Entra) application for your tenant, tick the capabilities
you want in **Microsoft Settings**, grant in Azure exactly the permission list it prints, then
each person creates a **Microsoft Calendar** record and clicks Authorize. The order is not
decoration — a token carries only the permissions consented before it was issued. The
walkthrough is [Set it up, in order](../README.md#connect-frappe-to-microsoft-365-in-order); the Azure side is in
[`azure-setup.md`](azure-setup.md).

## Does Frappe support Microsoft Teams meetings?

With this app, yes. An Event gets an **Add Teams meeting** tickbox; save, and the join link
comes back onto the Event with a **Join Meeting** button. A meeting organised in Outlook keeps
its join link when it syncs in, so people can join from either side. No extra Azure permission
is involved — Microsoft mints the Teams link as part of the calendar event, so
`Calendars.ReadWrite` covers it.

## Can I get Teams transcripts and recordings into Frappe?

Yes, for meetings that exist as calendar events. The transcript is attached to the Event as a
`.vtt`; recordings stay with Microsoft and stream through Frappe on demand rather than being
copied into the site's file store. This is the most permission-sensitive part of the app and it
depends on a tenant switch Microsoft ships **off** — read
[Transcripts and recordings](../README.md#teams-meeting-transcripts-and-recordings) before relying on it.

## Is there a Microsoft equivalent of Frappe's Google Calendar integration?

This app is meant as one. Frappe ships Google Calendar in core (a per-user `Google Calendar`
record, two-way `Event` sync); this app does the same job against Microsoft Graph and follows
the same patterns — one connection record per person, the same `System Manager` / `Desk User`
permissions, its own custom fields on `Event` sitting next to the Google ones. On top of the
calendar it adds Teams meetings, RSVP, transcripts and recordings, and setup plus diagnostics
for Outlook mail OAuth and Microsoft sign-in.

## Does it send or receive email, or replace Email Account?

No to all three. It does not send or receive mail, does not replace `Email Account`, and does
not touch the email queue. The **Outlook mail** capability creates the `Connected App` that
Frappe's own IMAP/SMTP + OAuth path needs, and the doctor *reads* your mail configuration to
tell you what is wrong with it. Frappe still sends the mail. See
[What this does NOT do](../README.md#what-this-app-does-not-do).

## Does it do Microsoft SSO?

Yes, by configuring Frappe's own Office 365 social login rather than implementing a second one.
Ticking **Sign in with Microsoft** creates a `Social Login Key` with tenant-specific **v2.0**
endpoints, because Frappe's built-in provider defaults to the `/common/` authority and the v1.0
endpoints, neither of which works with a single-tenant app registration. Self-registration is
denied by default: existing users with a matching email can sign in, new visitors cannot create
accounts.

---

## Do I need a separate Azure app registration per person?

No. One app registration for the tenant, one connection per person. The admin registers the app
and grants consent once; each employee then creates their own **Microsoft Calendar** and
authorizes it, which takes about twenty seconds and needs no admin involvement.

## Can my colleagues see my calendar?

Ordinary users see only their own connection, and can only authorize, sync, disconnect or read
events through it. Events pulled from your calendar are created **Private** and owned by you, so
they show on your calendar and nobody else's — that ownership is set explicitly, because the
scheduled sync runs as Administrator. The honest exception: **System Manager** and
Administrator can see every `Microsoft Calendar` record and act on it, including reading events
through it. Tokens themselves are stored as `Password` fields and are never logged or returned
to a client.

## Does an admin have to set each employee up?

No. After the one-time Azure work, each person creates their own connection and signs in as
themselves. Ticking **Sync with Microsoft Calendar** on an Event when you have no connection
offers to create one there and then, which is where most people meet this for the first time.

---

## Does this cost extra?

The app is free. The Azure app registration is free, and the Graph calls behind calendar sync,
Teams meetings and RSVP add no Microsoft charge beyond the Microsoft 365 licences your people
already have. Recordings are the part worth checking with your licensing: a meeting is only
recorded when the organiser's Teams licensing and your tenant's recording policy allow it, and
where nothing was recorded Graph answers with an empty list rather than an error.

## What licence is the app under?

MIT — see [`license.txt`](../license.txt). Commercial use, modification and redistribution are
permitted.

## Does it send my data anywhere?

Only to Microsoft. The only hosts the code talks to are `login.microsoftonline.com` and
`graph.microsoft.com`. There is no telemetry and no analytics. Secrets and tokens are stored as
Frappe `Password` fields and are never logged or returned to clients.

---

## Will installing this break my existing email?

It is built not to. If you never tick **Outlook mail**, nothing mail-related is created at all.
If you do, **Set Up** shows a plan first and only ever creates records that are *missing* —
anything that already exists is reported with its drift and left exactly as it is, and a second
run does nothing. The diagnostics only read configuration and generate text; nothing in the
doctor writes to an `Email Account`.

## Will it write into my real Outlook calendar?

Only when you ask it to. **Push Frappe events to Microsoft** is off by default on every
connection. Turning it on asks you to confirm and says plainly that deleting a Frappe Event will
then delete the Microsoft one. The other two actions with consequences outside Frappe — the
Exchange setup script, and an event whose end time precedes its start — are also stopped and
explained first. See
[Sensitive actions](../README.md#sensitive-actions-are-called-out-before-they-happen).

## What happens if I uninstall it?

Your mail setup keeps working. The `Connected App`, `Social Login Key` and `Email Account`
records are Frappe's own doctypes, this app never modified how mail is sent, and none of that is
removed with it. `Microsoft Settings` and your `Microsoft Calendar` connections go with the app.
The custom fields it adds to `Event` carry no module, so a plain uninstall leaves them behind —
delete them by hand if you want them gone.

---

## Which Frappe and ERPNext versions does it work with?

Frappe v15 and v16, from one codebase. Nothing is version-gated, and CI does a real install,
migrate and full test run against **both** on every push and pull request. Use
`--branch version-15` or `--branch version-16` if you follow the Frappe convention; `main` is
the source of truth and works for either. The API-by-API evidence is in
[`compatibility.md`](compatibility.md).

## Does it need ERPNext?

No. It needs Frappe and the standard `Event` doctype, which is Frappe's own. It works on a plain
Frappe site and on an ERPNext one; nothing in the app imports ERPNext.

## Will it run on Frappe Cloud?

It declares its supported Frappe range in `pyproject.toml` (`>=15.0.0-dev,<17.0.0`), which is
what Frappe Cloud reads when you add the app to a bench — without that section Frappe Cloud
refuses the app. Add it from its GitHub URL. See [Frappe Cloud](../README.md#frappe-cloud).

## Can I use a personal Microsoft account (outlook.com, hotmail.com)?

No. This is built for Microsoft 365 work or school tenants: consumer Microsoft accounts cannot
use the Teams and transcript permissions at all, and the setup assumes a tenant where an
administrator can grant consent.

---

## Which Azure permissions do I actually need?

Only the ones your tickboxes imply. Microsoft Settings derives the delegated scope list from the
capabilities and shows it under **Delegated Scopes** before you authorise, so what you grant in
Azure and what sign-in requests cannot drift apart.

| Ticked in Frappe | Delegated Microsoft Graph permissions |
| --- | --- |
| Outlook calendar (covers Teams meetings on an Event, and RSVP) | `User.Read` `Calendars.ReadWrite` |
| Standalone Teams meetings | `User.Read` `OnlineMeetings.ReadWrite` |
| Meeting transcripts and recordings | the row above, **plus** `OnlineMeetingTranscript.Read.All` `OnlineMeetingRecording.Read.All` |

`offline_access`, `openid` and `profile` are requested automatically at sign-in, but
`offline_access` still has to be **granted** in Azure like any other permission — without it
Microsoft issues no refresh token and the connection dies when the first access token expires.
Mail and sign-in are not Graph permissions; their lists are in
[Set up only what you want](../README.md#set-up-only-what-you-want).

## Why does granting one of those not cover the others?

Because Microsoft consents to them separately, and tenants routinely grant one and refuse
another. `OnlineMeetingTranscript.Read.All` reads a transcript you can already address by
meeting id — but you start from a join URL, and turning that into a meeting id is
`OnlineMeetings.ReadWrite`. Reading the video is a third permission again. That split is why a
calendar-only setup never has to ask for any of them.

## Do I need to be a Global Administrator?

You need someone who can grant admin consent for the tenant; a Privileged Role Administrator or
Cloud Application Administrator is enough. Transcripts additionally need a switch in the
**Teams admin center** — a different portal from Entra — which Microsoft ships off. Consent is
no substitute for it, and re-granting consent does nothing.

## I ticked another capability and now it returns 403

A token carries the permissions consented when it was issued, and ticking a box does not widen
one that already exists. **Re-authorize** each Microsoft Calendar. **Scopes Last Authorised** on
Microsoft Settings records what the last sign-in asked for, and Run Diagnostics compares it
against what is requested now and names the connections that need it.

---

## How quickly does a change sync?

Pull runs every 15 minutes on a delta query, so that is the worst-case lag for something changed
in Outlook. In the other direction, creating an event calls Graph inside the save — the join
link has to be on the document the browser is waiting for — while edits and deletions are queued
and run moments later. There are no Graph change webhooks here, and the README explains why
under [Transcripts and recordings](../README.md#teams-meeting-transcripts-and-recordings).

## Does it sync every calendar in my mailbox?

No. It syncs the default calendar of the account that authorized the connection
(`/me/calendarView`). A second calendar in the same mailbox is neither read nor written, and
there is no calendar picker.

## Will it sync a shared mailbox or a room calendar?

No. The sync is the signed-in person's own calendar. A shared mailbox's calendar or a room
calendar needs delegated access to that mailbox or the application-permission path, and this app
sets up neither today. Shared mailboxes for *mail* are a different matter — the app-only mail
flow and the generated Exchange setup script cover that.

## Will a repeating Frappe Event become a recurring series in Outlook?

Yes. **Repeat On** is translated into a Graph recurrence — daily, weekly on the days you tick,
monthly, quarterly, half-yearly or yearly — ending on **Repeat Till**, or never if that is empty.
A weekly series always includes the day it starts on, because Graph rejects one that does not.

The occurrences Microsoft generates from that series are not mirrored back as separate Frappe
Events; the repeating Event already stands for them, in the same way one Graph series master
stands for its own occurrences.

In the other direction a recurring series organised **in Outlook** arrives as individual
occurrences, one Frappe Event each, since Frappe has nothing else to attach them to.

A frequency this app cannot map to Graph is sent as a single event rather than as a wrong
series: a wrong recurrence writes itself across a real calendar for months.

## Will it pick up Zoom or Google Meet links?

No. Microsoft fills in the structured `onlineMeeting` property only for its own Teams meetings;
a third-party link is loose text in the event body, and guessing at it would produce wrong links
more often than right ones. Open the event in Outlook for those.

## Can I remove a Teams meeting from an event after adding one?

No, and the tickbox locks itself once the meeting exists rather than sitting there doing nothing
when you untick it. Microsoft cannot turn an existing online meeting back into a plain event.
Delete the event and create it again. Re-asserting the tickbox would mint a *new* meeting and
orphan the old join link along with its transcript and recording, so the app never does.

## The meeting finished an hour ago and the transcript is still empty

Most likely processing. Microsoft publishes no schedule: 5-30 minutes is typical, a long or
heavy meeting can take hours, and Graph lags the Teams UI — a transcript can be readable in
Teams while the API still returns an empty list. The app keeps asking on a widening backoff for
a day, and the Event tells you which of "not ready yet", "Microsoft returned an error",
"probably never recorded" and "too old now" applies. A 403 is a different problem; see
[Troubleshooting](../README.md#troubleshooting-common-microsoft-errors).

---

## How does this relate to the other Frappe/Microsoft apps?

They solve different problems and can sit alongside this one. Summaries below are from each
project's own README — check them for current scope.

- **[ealecho/m365_sync](https://github.com/ealecho/m365_sync)** (and the related
  `m365_role_sync`) maps Microsoft 365 / Entra groups to ERPNext roles and keeps membership in
  step on a schedule. Identity and role provisioning; no calendar, mail or Teams.
- **[Aptitudetech/frappe-m365](https://github.com/Aptitudetech/frappe-m365)** is about Microsoft
  365 **Groups** — creating and managing groups and their members, a SharePoint site per group,
  and file storage and syncing. Its README states Frappe v13 and v14.
- **[castlecraft/microsoft_integration](https://github.com/castlecraft/microsoft_integration)**
  provides Microsoft sign-in for Frappe through a `Social Login Key` against Azure AD / Azure AD
  B2C, with configurable OAuth and JWKS endpoints.

This app's scope is Outlook calendar two-way sync, Teams meetings, RSVP, transcripts and
recordings, plus setup and diagnostics for Outlook mail OAuth and Microsoft sign-in. For
group-to-role provisioning or SharePoint groups, those apps are where to look.

## Can another app call this one instead of writing its own Graph code?

Yes, that is the intended way to use it. The whitelisted utilities — create a Teams meeting,
fetch events, fetch transcripts for a join URL, reply to an invitation — are listed under
[How it's consumed](../README.md#python-api-for-other-frappe-apps), with the verified
endpoint and permission contract in [`graph-api-reference.md`](graph-api-reference.md).

---

## Where do I go when it does not work?

**Microsoft Settings → Troubleshoot**. **Run Diagnostics** reads the whole configuration —
settings, the Connected App, every Email Account, the authorised scopes — and names the failing
step instead of the symptom. **Explain an Error** decodes a message from the Error Log into a
cause and a next step. The common symptoms are tabulated in
[Troubleshooting](../README.md#troubleshooting-common-microsoft-errors) and in
[`troubleshooting.md`](troubleshooting.md).
