# Setting up Microsoft 365 for Frappe

Connecting Frappe to Microsoft 365 spans two Microsoft portals and your own site. The steps
below are in the order they actually have to happen — a token carries only the permissions that
were consented *before* it was issued, so doing them out of order fails with a 403 that names
nothing.

If something goes wrong, [every Microsoft error this app can produce is written up with its real
cause](troubleshooting.md).

# Requirements

- A Frappe **v15 or v16** bench ([install guide](https://frappeframework.com/docs/user/en/installation)).
- Python ≥ 3.10 (whatever your Frappe version requires: v15 runs on 3.10+, v16 needs 3.14).
  The only extra dependency is **`msal`** (declared in `pyproject.toml`, installed
  automatically by `bench get-app`).
- An **Azure AD (Microsoft Entra) app registration** — free; walked through step by step below.

# Connect Frappe to Microsoft 365, in order

Setup spans **two Microsoft portals and Frappe**, and the order is not decoration. The permission
list depends on what you ticked in Frappe; the tenant switch in step 7 lives in a portal most
people never open; and the token you end up with carries only the permissions that were consented
*before* it was issued. Done out of order, every one of those bites you as a 403 that names
nothing.

The same checklist is available inside the app: **Microsoft Settings → Setup Guide**.

## 1. Frappe — tick the capabilities you want

`Microsoft Settings → Capabilities`

Calendar, Teams meetings, transcripts, mail and sign-in are independent — tick what you want and
leave the rest alone. **Delegated Scopes**, immediately below, then prints the exact permission
list to grant. Copy it: step 5 is nothing but pasting it into Entra.

Do this first. Which permissions to grant, and whether you have to visit the Teams admin center at
all, both follow from it.

## 2. Entra — register the application

`entra.microsoft.com → Entra ID → App registrations → New registration`

- **Name:** anything you will recognise later, e.g. `Frappe Microsoft 365`.
- **Supported account types:** *Accounts in this organizational directory only* unless people from
  other Microsoft work/school tenants have to connect. Personal Microsoft accounts cannot use the
  Teams or transcript permissions at all.
- Leave the redirect URI blank — it is the next step.

Afterwards, from the app's **Overview** page, copy **Application (client) ID** and **Directory
(tenant) ID**. You paste both into Frappe in step 8.

## 3. Entra — add the redirect URI

`Your app → Authentication → Add a platform → Web`

```
https://<your-site>/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback
```

The platform must be **Web** — this is a confidential client with a secret, not a SPA. The URI has
to match what Frappe sends *exactly*: scheme, host, port, path. `http://` is accepted for
`localhost` hosts, so local testing works. A mismatch surfaces at sign-in as `AADSTS50011` and
never says which character is wrong.

## 4. Entra — create a client secret

`Your app → Certificates & secrets → Client secrets → New client secret`

**Copy the Value, not the Secret ID.** Azure prints them side by side in the same row, only the
Value works, and it is shown once and never again. Microsoft's only comment on the mix-up is
`AADSTS7000215` at sign-in, long after you have moved on — which is why Microsoft Settings refuses
a secret that looks like a GUID on save.

Note the expiry date. Secrets expire and sign-in stops working the day they do.

## 5. Entra — add the delegated permissions

`Your app → API permissions → Add a permission → Microsoft Graph → Delegated permissions`

Add exactly what Microsoft Settings printed in step 1 — no more. Asking for scopes the tenant never
consented to is itself a way to make sign-in fail.

**Granting one does not imply the others.** They stack:

| Ticked in Frappe | Delegated Microsoft Graph permissions |
| --- | --- |
| Outlook calendar | `User.Read` `Calendars.ReadWrite` |
| Standalone Teams meetings | `User.Read` `OnlineMeetings.ReadWrite` |
| Meeting transcripts and recordings | the row above, **plus** `OnlineMeetingTranscript.Read.All` `OnlineMeetingRecording.Read.All` |

Transcripts are the row people get wrong, and all three permissions are load-bearing.
`OnlineMeetingTranscript.Read.All` reads a transcript you can already address by online meeting id
— but you start from a join URL, and turning that into a meeting id is `OnlineMeetings.ReadWrite`.
Recordings are consented separately again as `OnlineMeetingRecording.Read.All`, because Microsoft
treats reading the words and reading the video as different permissions and plenty of tenants grant
one and refuse the other. Frappe enforces the first half: *Meeting transcripts and recordings* only
appears once *Standalone Teams meetings* is on, and is cleared if you turn that back off.

`offline_access`, `openid` and `profile` are requested automatically at sign-in and never belong in
a list you maintain — but `offline_access` still has to be **granted here** like any other, or
Microsoft issues no refresh token and the connection dies when the first access token expires.

Mail and sign-in are not Graph permissions at all; their lists are in *Set up only what you want*.

## 6. Entra — grant admin consent, and check it took

`Your app → API permissions → Grant admin consent for <tenant>`

Then read the **Status** column. Every permission you added should show a green tick and
**Granted for &lt;tenant&gt;**. A row still reading *Not granted* is a permission that will 403 at
runtime however correctly the app asks for it.

If the button is greyed out, your account cannot consent — a Privileged Role Administrator or
Cloud Application Administrator has to do this step for you.

## 7. Teams admin center — turn on transcript API access

**Only if you ticked transcripts.** Skip this step otherwise.

`admin.teams.microsoft.com → Meetings → Meeting settings → Transcript API access`

This is a **different portal from Entra**, and consent is no substitute for it. Microsoft added a
tenant-level switch for Graph access to transcripts in 2026 and **ships it off**, so a tenant with
every permission granted and consented still gets:

```
403 Forbidden: Graph API access to transcripts is disabled for this tenant.
```

Turn **Microsoft Graph access** **On**. Then select **Configure** and turn **Include speaker
attribution** **On** if you want speaker names in the VTT — that one is off by default too, and
without it the transcript is text with nobody's name against it.

The PowerShell equivalent, if you would rather not click:

```powershell
Set-CsTeamsMeetingConfiguration -EnableGraphTranscriptAccess true -EnableAttributedTranscripts true -Identity Global
```

There is no request-side workaround and re-granting consent does nothing, which is why both the
error message and the doctor name this switch specifically rather than blaming permissions.

## 8. Frappe — paste the credentials and enable

`Microsoft Settings → Azure AD Application`

Paste **Tenant ID**, **Client ID** and the secret **Value** from step 4, tick **Enabled**, Save.

Leave **Redirect URI** blank unless the site's public URL is not what Frappe computes — blank uses
the site URL, which is right for most installs. If you do fill it in, it must match the URI you
registered in step 3 character for character.

## 9. Frappe — Set Up (only if you ticked mail or sign-in)

`Microsoft Settings → Set Up`

Calendar, Teams and transcripts need nothing created; this app talks to Graph directly. This step
exists for **Outlook mail**, which needs a `Connected App` for Frappe's Email Account, and **Sign
in with Microsoft**, which needs a `Social Login Key`. Set Up shows a plan first and only ever
creates what is missing — anything that already exists is reported and left exactly as it is.

The mail `Connected App` has a redirect URI of its own, computed by Frappe from the record name, so
it cannot be known until the record exists. Register that one in Entra too; **Run Diagnostics**
prints the exact URI. A missing one shows up later as `AADSTS50011`.

## 10. Frappe — authorize a connection

`Microsoft Calendar → New → Save → Authorize Microsoft Access`

Give the connection an **Account Name**, set **User** to the person it belongs to, Save, then
authorize and sign in as that person. The Microsoft account email and default calendar fill in
automatically.

Two toggles here are off on purpose. **Push Frappe events to Microsoft** stays off until you want
Frappe writing into a real calendar. **Fetch transcripts after meetings** stays off until you want
them collected without anyone pressing a button.

**If a connection already existed when you changed the capabilities, Re-authorize it.** A token
carries the permissions consented when it was issued; ticking another capability does not widen a
token that already exists, and the new feature fails with a 403 that explains nothing. This is why
step 1 is step 1. **Scopes Last Authorised** on Microsoft Settings records what the last sign-in
actually asked for, and Run Diagnostics compares it against what is requested now.

## 11. Verify

- **Microsoft Settings → Troubleshoot → Run Diagnostics.** It reads the whole configuration —
  settings, the Connected App, every Email Account, the authorised scopes — and names the failing
  step instead of the symptom.
- **The Delegated Scopes preview** shows exactly what sign-in will request. If it lists something
  you did not grant in step 5, sign-in fails; if you granted something it does not list, you
  granted more than this app will ever use.
- **Microsoft Calendar → Test Connection**, then **Sync Now**. A working sync reports *Pulled n,
  deleted n, pushed n*.
- **Transcripts:** open an Event whose Teams meeting has finished and click **Get Transcript &
  Recording**. Working looks like a `.vtt` attached to the Event and the recordings listed on it.
  *Microsoft has not finished processing this meeting* is a normal answer in the first half hour —
  see *Teams meeting transcripts and recordings*. A 403 naming the tenant is step 7; a 403 naming permissions is
  step 5 or 6.

Azure-side detail beyond this, including the app-only path for shared mailboxes:
[`docs/azure-setup.md`](azure-setup.md).

# Set up only what you want

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

## The Azure scopes follow the tickboxes

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

# Frappe Cloud

The app declares its supported Frappe range in `pyproject.toml`, which is what Frappe Cloud
reads when you add it to a bench:

```toml
[tool.bench.frappe-dependencies]
frappe = ">=15.0.0-dev,<17.0.0"
```

Without that section Frappe Cloud refuses the app with *"Could not find a compatible Frappe
version in pyproject.toml"*. The lower bound is anchored at `-dev` because pre-release builds
report versions like `15.0.0-dev`, which sort **below** `15.0.0` under semver.

# Quick start from scratch (try it on a fresh local site)

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

Now do the Microsoft side — *Connect Frappe to Microsoft 365, in order*, below. The one value Entra needs from you is
the **Redirect URI**:

```
http://m365.localhost:8000/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback
```

(Use your real site URL/port. `http://...localhost` is accepted by Azure for local testing.)

