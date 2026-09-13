# Troubleshooting — error by error

Connecting Frappe to Microsoft 365 spans two Microsoft portals, Exchange Online and Frappe, and
almost every mistake in it surfaces as one of a handful of strings that name nothing:
`AUTHENTICATE failed`, `535 5.7.3`, `invalid_grant`, `403 Forbidden`. This page is one section per
error, headed by the literal text you would paste into a search box, with what it actually means
and the exact place to fix it.

Two things are worth knowing before you read on.

**A 403 from Microsoft is not one diagnosis.** This app can produce six different ones — two tenant
switches, three permissions and a mailbox refusal. They are indistinguishable by status code and
their fixes live in three different places. Read the body of the message, never the number.
Microsoft's own advice on the transcript endpoints is to branch on `innerError.code`, not on the
message text, because the text is subject to change.

**The app decodes these itself.** `Microsoft Settings → Troubleshoot → Explain an Error` takes any
message from the Error Log and returns the same answer as this page; `Run Diagnostics` reads the
whole configuration and names the failing step. This page exists for the times you are looking at
the error somewhere else.

Setup instructions are in [`azure-setup.md`](azure-setup.md); the endpoint and permission contract
is in [`graph-api-reference.md`](graph-api-reference.md).

---

## Start here

Three causes account for most of the reports.

1. **[Graph API access to transcripts is disabled for this tenant](#transcripts-disabled-for-tenant)** —
   a tenant switch Microsoft ships **off**, in the Teams admin center. It is not a permission
   problem and no amount of consent fixes it. Most of what you will find online about 403s is
   wrong about this one.
2. **[A 403 on something you just granted](#403-after-consent)** — consent and the token are two
   different things. A token carries the permissions it was issued with. After granting anything,
   **Re-authorize** each Microsoft Calendar.
3. **[`AADSTS7000215`](#aadsts7000215)** — the client secret is wrong. In practice that almost
   always means the **Secret ID** was pasted instead of the secret **Value**; Azure prints them
   side by side and only the Value works. An *expired* secret is a
   [different code](#aadsts7000222).

---

## Contents

**Sign-in and OAuth**

- [`AADSTS50011`](#aadsts50011) — redirect URI mismatch
- [`AADSTS65001`](#aadsts65001) — no consent
- [`AADSTS7000215` / `invalid_client`](#aadsts7000215) — wrong client secret
- [`AADSTS7000222`](#aadsts7000222) — expired client secret
- [`AADSTS700016`](#aadsts700016) — application not in this tenant
- [`AADSTS90002`](#aadsts90002) — tenant not found
- [`invalid_grant`](#invalid-grant) — the grant is no longer usable
- [Works for an hour, then needs authorising again](#offline-access) — `offline_access` missing
- [`Microsoft sign-in failed: …`](#microsoft-sign-in-failed) — this app's wrapper around all of the above
- [`You cannot use any scope value that is reserved`](#reserved-scopes) — MSAL, from the scope override
- [`Microsoft Calendar '<name>' is not authorized. Click Authorize.`](#not-authorized)
- [`Microsoft 365 integration is disabled.`](#integration-disabled)

**Calendar sync**

- [`ErrorAccessDenied` / `Access is denied. Check credentials and try again.`](#erroraccessdenied)
- [`ErrorItemNotFound` / `The specified object was not found in the store.`](#erroritemnotfound)
- [`ErrorPropertyValidationFailure`](#errorpropertyvalidationfailure)
- [`This event ends before it starts, and Microsoft will not accept it.`](#ends-before-it-starts)
- [`syncStateNotFound` / `resyncRequired` / `Microsoft Graph sync state expired`](#syncstatenotfound)
- [`429` / `TooManyRequests` / `Microsoft Graph rate limit hit`](#throttling)
- [`MailboxNotEnabledForRESTAPI` / `REST API is not yet supported for this mailbox`](#mailboxnotenabledforrestapi)
- [The sync reports success and nothing moves](#sync-does-nothing)

**Teams meetings, transcripts and recordings**

- [`403 Forbidden: Graph API access to transcripts is disabled for this tenant`](#transcripts-disabled-for-tenant)
- [`403 Forbidden` / `SpeakerAttributionNotAllowed`](#speakerattributionnotallowed)
- [A 403 on a permission you have already granted](#403-after-consent)
- [`403` on `/transcripts`](#403-transcripts) — `OnlineMeetingTranscript.Read.All`
- [`403` on `/recordings`](#403-recordings) — `OnlineMeetingRecording.Read.All`
- [`403` on `/me/onlineMeetings`](#403-onlinemeetings) — `OnlineMeetings.ReadWrite`
- [`Microsoft could not match this join link to a meeting`](#join-link-no-match)
- [`Standalone meeting is not calendar-associated; transcripts are unavailable.`](#standalone-meeting)
- [`Microsoft has not finished processing this meeting.`](#still-processing)
- [`Microsoft still has no transcript or recording a day after this meeting`](#nothing-found)
- [`This meeting is too old, but the files may still exist.`](#meeting-expired)
- [Speaker names missing from the `.vtt`](#speaker-attribution)
- [`Microsoft gave us the transcript, but Frappe could not store the file`](#transcript-not-stored)
- [`This meeting has not started yet.` / `This event has no Teams meeting.`](#meeting-preconditions)
- [`The Microsoft Calendar … is switched off` / `… is not signed in to Microsoft`](#calendar-off-or-unauthorised)
- [`That recording does not belong to this meeting.`](#recording-not-this-meeting)
- [`Nothing is checking automatically`](#nothing-checking-automatically)

**Outlook mail (IMAP and SMTP)**

- [`AUTHENTICATE failed` / `A01 NO AUTHENTICATE failed`](#authenticate-failed)
- [`535 5.7.3 Authentication unsuccessful`](#535-573)
- [`535 5.7.139`](#535-57139)
- [`550 5.7.60 Client does not have permissions to send as this sender`](#550-5760)
- [`451 4.7.0 Temporary server error`](#451-470)
- [`TLS required` / STARTTLS failures](#tls-required)
- [`Please Authorize OAuth for Email Account …`](#please-authorize-oauth)
- [SMTP works but IMAP does not, on a shared mailbox](#shared-mailbox)
- [Basic authentication stopped working, or is about to](#basic-auth)

**Installation, database and background jobs**

- [`Unknown column 'custom_…microsoft…'`](#unknown-column)
- [`Data too long for column` / `(1406, …)`](#data-too-long)
- [`The scheduler is off` / `No background worker is running`](#background-jobs)
- [`This app's scheduled jobs have stopped running`](#jobs-stale)

[When nothing here matches](#nothing-matches)

---

## Sign-in and OAuth

<a id="aadsts50011"></a>

### `AADSTS50011: The redirect URI specified in the request does not match the redirect URIs configured for the application`

Also seen as *"The reply address is missing, misconfigured, or doesn't match reply addresses
configured for the app"* — Microsoft's short name for it is **InvalidReplyTo**. Either way, the
redirect URI the app sent is not one of the URIs registered on the Azure app registration.

**Fix.** Entra admin center → **Entra ID** → **App registrations** → your app →
**Authentication** → **Web** → **Redirect URIs**. The value must match what Frappe sends, exactly:

```
https://<your-site>/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback
```

Microsoft's matching rules, which is where the character usually goes wrong:

| | Rule |
| --- | --- |
| Scheme | Must be `https`, with exceptions for some `localhost` redirect URIs |
| Path | **Case-sensitive** — it must match the case of the URL path of the running application |
| Port | Ignored **only** for a localhost redirect URI. In every other case the port is part of the match |
| Trailing slash | A URI with no path segment is stored with one (`https://contoso.com` → `https://contoso.com/`); one with a path segment is not |

**The trap this app has of its own.** If you use the Outlook mail capability, Frappe's
`Connected App` computes a **second** redirect URI from its own record name, so it cannot be known
until that record exists, and it is a different endpoint from this app's callback. Azure needs
**both** registered. `Run Diagnostics` prints the exact URI to add; a missing one shows up here,
weeks later.

**What it is not.** It is not the Redirect URI field in Microsoft Settings being empty. Blank means
"use the site URL", which is right for most installs. Fill it in only when Frappe's computed URL is
not the public one, and then it has to match Azure character for character.

<a id="aadsts65001"></a>

### `AADSTS65001: The user or administrator has not consented to use the application with ID …`

Microsoft's short name: **DelegationDoesNotExist**. The permission was added to the app
registration but nobody granted consent for it, or it was never added at all.

**Fix.** Entra admin center → **Entra ID** → **App registrations** → **All applications** → your
app → **API permissions** (under **Manage**) → **Grant admin consent for &lt;tenant&gt;** → **Yes**.
Then read the **Status** column: each row should say **Granted for &lt;tenant&gt;**. A row still
reading *Not granted* will 403 at runtime however correctly the app asks for it.

If the button is greyed out, your account cannot consent. Microsoft Graph **application**
permissions need a Privileged Role Administrator; everything else can be done by Cloud Application
Administrator or Application Administrator.

**Then re-authorize.** Consent is not retroactive — see
[A 403 on a permission you have already granted](#403-after-consent).

**Mail has a wrinkle here.** For a **multi-tenant** app registration, Microsoft documents consent
via an `/adminconsent` URL whose `scope` differs by protocol: POP and IMAP application permissions
use `https://ps.outlook.com/.default`, and SMTP uses `https://outlook.office365.com/.default`
("only for SMTP", in Microsoft's words). For a single-tenant registration Microsoft's own guidance
is to use the portal button above instead and skip the URL entirely. Note this split applies to the
**consent** request; the **token** request for app-only mail always uses
`https://outlook.office365.com/.default`.

<a id="aadsts7000215"></a>

### `AADSTS7000215: Invalid client secret is provided` / `invalid_client`

The secret Frappe sent is not the app's secret. Microsoft's description says no more than that.

**In practice it is one of three things**, in descending order of how often it happens:

1. **The Secret ID was pasted instead of the secret Value.** Azure's *Certificates & secrets* table
   prints them side by side in the same row and only the Value works. Microsoft does not document
   this as a cause of `7000215` — the code only means "the secret is wrong" — so treat this as
   operational experience rather than a citable rule. It is also the one cause that is detectable
   from the outside: a Secret ID is a GUID and a secret value never is, which is why **Microsoft
   Settings refuses a secret that looks like a GUID on save**.
2. The secret was truncated, or carries whitespace from the copy.
3. The secret belongs to a different app registration than the Client ID.

**Fix.** Entra admin center → **App registrations** → your app → **Certificates & secrets** →
**Client secrets** → **New client secret**. Copy the **Value** column immediately — Microsoft's
wording is that *"this secret value is never displayed again after you leave this page"*. Paste it
into **Microsoft Settings → Azure AD Application → Client Secret Value**.

**An expired secret gives a different code.** See [`AADSTS7000222`](#aadsts7000222).

<a id="aadsts7000222"></a>

### `AADSTS7000222: The provided client secret keys are expired`

Microsoft's short name: **InvalidClientSecretExpiredKeysProvided**. The secret was right and has
now passed its expiry date. Azure caps secret lifetime at 24 months and recommends under 12, so
every install meets this eventually — on a day when nothing else changed, which is what makes it
confusing.

**Fix.** Create a new secret (same path as above) and paste the new **Value** into Microsoft
Settings. Nothing else needs redoing: consent, permissions and the Connected App are unaffected,
and connections start working again as soon as the secret is right.

This app decodes `7000222` through the same entry as `7000215`, because MSAL surfaces both under
the OAuth error `invalid_client`. The distinction matters when you are reading the raw text:
**7000215 means the secret is wrong, 7000222 means it has run out.**

<a id="aadsts700016"></a>

### `AADSTS700016: The application wasn't found in the directory/tenant`

Microsoft's short name: **UnauthorizedClient_DoesNotMatchRequest**. Microsoft looked for the Client
ID in the tenant named in the sign-in URL and it is not there.

**Fix.** Microsoft's own description names both candidates:

- **Wrong tenant, or a misconfigured identifier.** Copy **Directory (tenant) ID** and
  **Application (client) ID** from the app registration's **Overview** page into Microsoft
  Settings.
- **The application has never been installed or consented into this tenant.** This is the
  multi-tenant case: the app is registered elsewhere, and an administrator of *this* tenant has to
  consent it in before anyone here can use it.

<a id="aadsts90002"></a>

### `AADSTS90002: Tenant '<something>' not found`

Microsoft's short name: **InvalidTenantName** — "the tenant name wasn't found in the data store".
The tenant segment of the sign-in URL does not resolve. Usually the Tenant ID field holds a company
name, or a domain that is not registered on the tenant.

**Fix.** Use **Directory (tenant) ID** from the app registration **Overview** page. A verified
domain of the tenant also works. Microsoft Settings warns on save when the value is neither a GUID
nor a domain.

<a id="invalid-grant"></a>

### `invalid_grant`

The grant Frappe is presenting cannot be used, so no new access token can be minted.

`invalid_grant` is generic — Microsoft defines it as *"some of the authentication material (auth
code, refresh token, access token, PKCE challenge) was invalid, unparseable, missing, or otherwise
unusable"* — and the useful detail is in the `error_description` beside it. Two codes are worth
recognising there:

| In the description | Meaning |
| --- | --- |
| `AADSTS700082` | **ExpiredOrRevokedGrantInactiveToken** — the refresh token expired through inactivity. Microsoft calls this an expected part of the token lifecycle, not a fault |
| `AADSTS50173` | The grant was revoked and a fresh sign-in is needed. Microsoft names a password change or reset as a cause |

(Microsoft does not publish a table mapping those AADSTS codes onto the `invalid_grant` OAuth
error, so read whichever code is actually in front of you rather than assuming.)

**Fix.** Open the **Microsoft Calendar** and click **Re-authorize**, signed in as the user it
belongs to. For an Email Account, **Authorize API Access**, signed in as the Connected User.

**If it returns within hours rather than weeks**, the cause is different: see
[`offline_access`](#offline-access). A connection that needs re-authorising every few hours was
never issued a refresh token at all.

<a id="offline-access"></a>

### Works for about an hour, then needs authorising again — every time

There is no distinctive error for this one. It looks like `invalid_grant`, or like the connection
quietly dying, on a cycle roughly the length of an access token.

**What it means.** Microsoft only returns a refresh token when `offline_access` was requested. The
auth-code flow reference is blunt about it: the `refresh_token` response parameter is *"only
provided if `offline_access` scope was requested"*. Without one you get a single access token, good
for about an hour, and then nothing.

**Fix, for calendar, Teams and transcripts.** This app adds `offline_access` to every sign-in
automatically, so the *request* is never the problem — but the permission still has to exist on the
app registration. Entra → your app → **API permissions** → add **`offline_access`** (Microsoft
Graph, delegated), grant admin consent, then **Re-authorize**.

**Fix, for mail.** Frappe's `Connected App` holds its own scope list, and `offline_access` has to
be in it. `Run Diagnostics` reports a missing one as a hard failure, because it is the documented
cause behind the "loses access after a few hours" reports that recur on the Frappe forum.

<a id="microsoft-sign-in-failed"></a>

### `Microsoft sign-in failed: <something>`

This app's own wrapper. MSAL returned a result with no access token, and everything after the colon
is Microsoft's `error_description`, passed through unchanged. No token or secret is ever included.

**So the message is not the diagnosis: the `AADSTS` code inside it is.** Find that number and read
its section above. `Microsoft Settings → Troubleshoot → Explain an Error` does the same lookup if
you would rather paste it.

<a id="reserved-scopes"></a>

### `You cannot use any scope value that is reserved`

Raised by MSAL inside the site, before any request reaches Microsoft. MSAL refuses `openid`,
`profile` and `offline_access` in the `scopes` argument of its token calls — it adds them itself.

**Cause.** Somebody filled in **Microsoft Settings → Delegated Scopes → Override Delegated Scopes**
and listed one of the three.

**Fix.** Remove them from the override. They are requested at sign-in automatically, and listing
them breaks the token call rather than widening consent. `Run Diagnostics` catches this before you
hit it. If the override is not doing something you specifically need, clear it entirely and let the
capability tickboxes decide.

<a id="not-authorized"></a>

### `Microsoft Calendar '<name>' is not authorized. Click Authorize.`

No refresh token is stored against that connection: it was never authorised, or it was
disconnected.

**Fix.** Open the **Microsoft Calendar** record and click **Authorize Microsoft Access** while
signed in to Frappe as the user in the **User** field. The Microsoft account email and default
calendar fill in on return.

<a id="integration-disabled"></a>

### `Microsoft 365 integration is disabled. Enable it in Microsoft Settings.`

Every live operation checks this first. It is a deliberate stop, not a fault.

**Fix.** **Microsoft Settings** → tick **Enabled** → Save. Configuration reads — the diagnostics,
the setup guide, the scope preview — all work with it off, so the setup can be finished before it
is turned on.

---

## Calendar sync

<a id="erroraccessdenied"></a>

### `ErrorAccessDenied: Access is denied. Check credentials and try again.`

The token is valid — Microsoft accepted the sign-in — and the mailbox refused the request anyway.
Exchange defines the code as *"the calling account does not have the rights to perform the
requested action"*. (Graph's own 403 wording adds a second possibility: *"the user does not have
enough permission or does not have a required license"*.)

**Fix.** Three candidates, and the error does not distinguish them:

- **The calendar permission is missing.** Add **`Calendars.ReadWrite`** (Microsoft Graph,
  delegated) in Entra → your app → **API permissions**, grant admin consent, then **Re-authorize**
  the Microsoft Calendar.
- **It is not this account's calendar.** Reading somebody else's calendar needs them to share it
  with the connected account in Outlook first. No permission on the app registration substitutes
  for that.
- **The mailbox has no licence.** See [`MailboxNotEnabledForRESTAPI`](#mailboxnotenabledforrestapi)
  for the discriminating check.

**What it is not.** It is not a wrong secret or an expired token; those fail earlier and
differently.

<a id="erroritemnotfound"></a>

### `ErrorItemNotFound: The specified object was not found in the store.`

The Microsoft event id stored on a Frappe Event resolves to nothing. Usually a `PATCH` or `DELETE`
against an event that has gone.

**Careful with the obvious reading.** Exchange defines this as *"the item was not found **or you do
not have permission to access the item**"*. It is not proof the event was deleted. If a whole
calendar's worth of events start reporting it at once, suspect access rather than deletion — a
shared calendar that stopped being shared looks exactly like this.

**Causes, for a single event**, in rough order: deleted in Outlook; moved to a different calendar;
an occurrence of a recurring series that has since changed.

**Fix.** Nothing can patch an event Outlook does not have. Delete the Frappe copy, or untick **Sync
with Microsoft Calendar** on it so the sync stops trying.

<a id="errorpropertyvalidationfailure"></a>

### `ErrorPropertyValidationFailure`

Graph rejected a value in the event body. This app passes through both halves of Microsoft's error
on purpose — the code is what you search for, and the message is the only part that says *which*
property Graph objected to.

**Be aware that the documentation is no help here.** Exchange's response-code reference lists
`ErrorPropertyValidationFailure` as *"This response code is not used"*, which is plainly out of
date, and Graph's `event` resource documents no ordering constraint between `start` and `end` and
no error for violating one. So what follows is this app's operational experience, not a documented
rule.

**In practice the commonest cause is an end time that is not after the start time.** Frappe
pre-fills both `starts_on` and `ends_on` from the current moment and does not enforce the order, so
an event saved without touching the times can end before it begins.

**Fix.** Check the start and end on the Event. This app refuses that combination before it reaches
Microsoft (next entry), so a fresh `ErrorPropertyValidationFailure` is more likely to be an older
record being retried, or a genuinely different property — read the message.

<a id="ends-before-it-starts"></a>

### `This event ends before it starts, and Microsoft will not accept it. Set an end time after …`

This app's own guard, raised on save, before anything is sent. Title: *Check the times*.

**Fix.** Set `ends_on` after `starts_on`.

It fires only on events being synced to Microsoft, and never on data arriving *from* Microsoft.
Outlook is authoritative during a pull, and refusing what it sends would stall the sync on a record
you cannot fix from Frappe.

<a id="syncstatenotfound"></a>

### `syncStateNotFound` / `resyncRequired` / `Microsoft Graph sync state expired (…). A full re-sync is required.`

Microsoft invalidated the delta token this app uses as its sync watermark. Graph answers `410 Gone`
with a `Location` header pointing at a fresh, empty delta query.

**This is not a fault and there is nothing to fix.** Microsoft's own description is that it
*"usually happens to prevent data inconsistency due to internal maintenance or migration of the
target tenant, and is an indication that the application must restart with a full synchronization"*.
Delta tokens for Outlook entities have no fixed lifetime — it depends on the size of an internal
cache — so a long gap between syncs is enough. The next run starts a full sync by itself and the
watermark re-establishes.

**When to look further.** If it happens on *every* run, the delta link is not being stored. Check
the **Delta Link** field on the Microsoft Calendar record, and the Error Log for whatever is
failing to write it.

<a id="throttling"></a>

### `429` / `TooManyRequests` / `activityLimitReached` / `Microsoft Graph rate limit hit (429)`

Microsoft is throttling the tenant. Graph returns `429 Too Many Requests`, usually with a
`Retry-After` header. (`TooManyRequests` is the code in Graph's own 429 example;
`activityLimitReached` is the Files/OneDrive-scoped variant — this app's decoder matches both.)

**Nothing is misconfigured.** This app follows Microsoft's stated recovery — wait the number of
seconds in `Retry-After`, then retry — inline for short waits, up to 10 seconds, and leaves longer
backoffs to the next scheduled run. So it clears itself.

**If every run hits it**, the load is genuinely above what the tenant allows. Microsoft's published
Outlook limits are per **app ID and mailbox combination**:

| Limit | Value |
| --- | --- |
| Requests | 10,000 in a 10-minute period |
| Concurrency | 4 concurrent requests |
| Upload (PATCH/POST/PUT) | 150 MB in a 5-minute period |

Because the limit is per app *and* mailbox, exceeding it for one mailbox does not affect the
others — but several Microsoft Calendar records pointing at the *same* mailbox share one budget.
Sync fewer calendars, or narrow the sync window.

<a id="mailboxnotenabledforrestapi"></a>

### `MailboxNotEnabledForRESTAPI` / `REST API is not yet supported for this mailbox`

Graph will not serve this mailbox. Microsoft's KB documents the status as **404** and gives one
cause: *"the mailbox is on a dedicated Microsoft Exchange Server and is not a valid Microsoft 365
mailbox"* — in other words, the mailbox is not in Exchange Online.

The runtime message you may see alongside it mentions an inactive, soft-deleted or on-premises
mailbox. That wording is the service's, not the documentation's, but it points the same way.

**Fix.** Confirm the mailbox is actually hosted in Exchange Online rather than on-premises; if it
is not, it has to be migrated. If the account has no licence: Microsoft 365 admin center →
**Users** → **Active users** → select the user → **Licenses and Apps** → expand **Licenses** →
tick a licence that includes Exchange Online → **Save changes**. (Assigning licences needs at least
a License Administrator or User Administrator.)

<a id="sync-does-nothing"></a>

### The sync reports success and nothing moves

No error anywhere, `Last Sync` updating, and no events crossing.

Work through these in order:

| Check | Where |
| --- | --- |
| Is the direction switched on? | Microsoft Calendar → **Pull events from Microsoft** / **Push Frappe events to Microsoft**. Push is off by default, on purpose |
| Is the Frappe event marked for sync? | The Event's **Sync with Microsoft Calendar** tickbox |
| Is anything inside the window? | The pull covers 30 days back and 180 days forward. Events outside that are not pulled |
| Did the job actually run? | [Background jobs](#background-jobs) — a queued sync with no worker stays "queued" forever |
| Is the connection enabled and authorised? | Both tickboxes on the Microsoft Calendar record |

**Last Sync Error** on the Microsoft Calendar holds the most recent failure, and the Error Log holds
the traceback under a title beginning `MS Calendar`.

---

## Teams meetings, transcripts and recordings

<a id="transcripts-disabled-for-tenant"></a>

### `403 Forbidden: Graph API access to transcripts is disabled for this tenant` (`GraphAccessToTranscriptsDisabled`)

**This is not a permission problem, and it is the entry that costs people hours.** The obvious
reading of a 403 — missing permission, missing consent — is wrong here. Microsoft added a
**tenant-level switch** governing whether Graph may read Teams transcripts at all, and it ships
**off**. Microsoft's own wording: *"By default, Microsoft Graph access is off, so agents and apps
can't access meeting transcripts, **regardless of app-level permissions**."*

It also lives in a **different portal** from the permissions, which is why people who go looking in
Entra never find it. The Graph reference is equally blunt: *"There is no request-side workaround;
the app receives this response until an administrator re-enables access."*

The documented response body:

```json
{
  "error": {
    "code": "Forbidden",
    "message": "Graph API access to transcripts is disabled for this tenant.",
    "innerError": { "code": "GraphAccessToTranscriptsDisabled" }
  }
}
```

**Fix.** Teams admin center (`admin.teams.microsoft.com`) → **Meetings** → **Meeting settings** →
under **Transcript API access**, turn the **Microsoft Graph access** toggle **On**.

Or in Teams PowerShell:

```powershell
Set-CsTeamsMeetingConfiguration -Identity Global -EnableGraphTranscriptAccess $true
```

**Two things to expect afterwards.** Tenant settings take a while to reach every server, so an
empty answer in the first minutes after switching it on is not evidence that the meeting has no
transcript. And speaker names are a **second** switch — see
[`SpeakerAttributionNotAllowed`](#speakerattributionnotallowed).

**Why this app cannot check it for you.** The switch has no Microsoft Graph resource behind it:
only `Get-CsTeamsMeetingConfiguration` in Teams PowerShell can read it, and this app holds a
delegated Graph token, not a Teams admin session. `Run Diagnostics` therefore lists it as a step to
**confirm**, never as a fault — telling somebody their finished step is broken is the same loop by
another route.

**Dates, since "added in 2026" invites the question.** Microsoft's transcript-content reference
states that the related tenant administrator controls *"take effect at the end of July 2026"*.

<a id="speakerattributionnotallowed"></a>

### `403 Forbidden: Speaker-attributed transcript content is disabled for this tenant` (`SpeakerAttributionNotAllowed`)

A **second** tenant switch, separate from the one above, and one this app runs into by default.

**What it means.** Graph serves transcript content in two formats: `text/vtt`, which carries
`<v Speaker>` voice tags, and `application/vnd.microsoft.graph.transcript+text`, which is the same
transcript without speaker information. When the tenant administrator has disabled speaker
attribution, asking for the **attributed** format returns this 403. The unattributed format still
works — Microsoft's message even says so: *"Retry with Accept
'application/vnd.microsoft.graph.transcript+text'."*

**This app asks for `text/vtt`.** That is the format it attaches to the Event, and it is selected
with the `$format` query parameter. The unattributed format can only be selected with the `Accept`
header, so there is no way to fall back to it from here. On a tenant with attribution off, the
transcript fetch fails at the content step — after the list call succeeded, which makes it look
like a permission that half works.

**Fix.** Teams admin center → **Meetings** → **Meeting settings** → **Transcript API access** →
**Configure** → turn **Include speaker attribution** **On**. It is off by default and only
available once **Microsoft Graph access** is on. In PowerShell:

```powershell
Set-CsTeamsMeetingConfiguration -Identity Global -EnableAttributedTranscripts $true
```

**Only `/content` is affected.** Listing transcripts still works with attribution off, which is why
the Event can show that a transcript exists and then fail to fetch it.

<a id="403-after-consent"></a>

### A 403 on a permission you have already granted

The permission is on the app registration, the **Status** column says **Granted for &lt;tenant&gt;**,
and the call still fails.

**What it means.** An access token carries the permissions consented **when it was issued** —
Microsoft's wording is that *"encoded inside the access token is every permission that your
application is granted for that resource"*. Granting a new permission does not widen a token that
already exists, and the refresh token behind it was minted under the old consent. The new feature
therefore 403s until the user signs in again.

**Fix.** Open each **Microsoft Calendar** and click **Re-authorize**. For mail, **Authorize API
Access** on the Email Account.

**How to tell this is the cause.** **Microsoft Settings → Scopes Last Authorised** records the
scope list in force at the last successful sign-in. `Run Diagnostics` compares it with what sign-in
would request now and says how many connections are behind.

<a id="403-transcripts"></a>

### `403 Forbidden` on `GET /me/onlineMeetings/{id}/transcripts`

If the message names the tenant switch, it is [that one](#transcripts-disabled-for-tenant) — it
names itself, in `innerError.code`. If it does not, it is the permission.

**Fix.** Add **`OnlineMeetingTranscript.Read.All`** (Microsoft Graph, **delegated** — Microsoft
lists it as the least-privileged permission for this call and offers no alternative) in Entra →
your app → **API permissions**. Grant **admin consent**: this permission is marked as requiring it,
so the person signing in cannot consent to it themselves. Then **Re-authorize** the Microsoft
Calendar.

**If you are using application permissions rather than delegated**, there is an extra step
Microsoft documents and nothing in Entra hints at: a tenant administrator must create an
**application access policy** (`New-CsApplicationAccessPolicy`, then `Grant-CsApplicationAccessPolicy`)
authorising the app to read meetings on behalf of a user. Without one the call returns 403 with
*"No application access policy found for this app"*. Changes to those policies can take up to 30
minutes to take effect.

<a id="403-recordings"></a>

### `403 Forbidden` on `GET /me/onlineMeetings/{id}/recordings`

Recordings have their own permission, consented separately from transcripts. A tenant happy to let
an application read the words frequently refuses to let it read the video, so this is the one that
goes missing on its own.

**Fix.** Add **`OnlineMeetingRecording.Read.All`** (Microsoft Graph, delegated) in Entra → your app
→ **API permissions**, grant admin consent, then **Re-authorize**.

**What it is not.** It is not the tenant's recording policy. Whether the tenant is allowed to
record at all is a separate setting — Teams admin center → **Meetings** → **Meeting policies** →
the organiser's policy → **Recording & transcription** → **Meeting recording** — and that one shows
up as *no recordings*, never as a 403. Leading with it would send you to the wrong page.

Microsoft's product documentation does not publish a licence matrix for meeting recording, so
treat any list of "licences that can record" you find elsewhere with suspicion. What is documented:
the policy toggle above, and that education A1 users can record manually but cannot auto-record
from meeting options.

<a id="403-onlinemeetings"></a>

### `403 Forbidden` on `GET /me/onlineMeetings?$filter=JoinWebUrl eq '…'`

**The permission people miss.** They grant the transcript permission and stop. But a transcript is
addressed by online meeting id, and the only documented way to get that id from a join link is this
filtered call — which is a different permission. A key to a door you cannot walk to.

**Fix.** Tick **Standalone Teams meetings** in Microsoft Settings, which is what adds
**`OnlineMeetings.ReadWrite`** to the request, consent to it in Entra → your app → **API
permissions**, then **Re-authorize**.

Microsoft lists `OnlineMeetings.Read` as the least-privileged permission for the lookup itself and
`OnlineMeetings.ReadWrite` as the higher-privileged one that also works. This app requests
`ReadWrite` because the same capability also creates standalone meetings. Note that on the `/me`
path this call has **no** application-permission form at all — app-only has to use
`/users/{id}/onlineMeetings`, with an application access policy behind it.

**Why the app names this case specifically.** Left unwrapped, it surfaced as a bare Graph 403
against a path nobody recognises, which reads exactly like the transcript permission failing — and
sent a real administrator to re-grant consent that the portal in front of them already showed as
granted.

<a id="join-link-no-match"></a>

### `Microsoft could not match this join link to a meeting — either another organisation hosted it, or it has aged out.`

A `404`, not a `403`. Microsoft answered, and the meeting id does not resolve. Two causes, and
naming only one of them is wrong about half the time:

- **Another organisation hosted the meeting.** It lives in their tenant, not yours, and the
  connected account is a guest in it. Graph will not return it from `/me/onlineMeetings`.
- **It has aged out.** Both the transcripts and the recordings endpoints carry a caveat in
  Microsoft's reference — *"This API works for a meeting only if the meeting has not expired"* —
  and expiry is roughly 60 days. See [the two clocks](#meeting-expired).

**Fix.** Open the meeting in the Teams calendar and use its **Recordings and Transcripts** tab.
That view reads from the other side and settles which of the two it is: if the files are there,
this connection cannot reach them; if they are not, there is nothing to fetch.

**What it is not.** It is not permissions. A refusal is a 403 and says so.

<a id="standalone-meeting"></a>

### `Standalone meeting is not calendar-associated; transcripts are unavailable.`

Not "not yet" — never. Microsoft's own note on the list-transcripts endpoint: *"This API doesn't
support meetings created using the create onlineMeeting API that are not associated with an event
on the user's calendar."* A meeting created directly as an online meeting has no calendar event
behind it and will not have a transcript however long you wait.

**Fix.** Create the meeting from a Frappe Event with **Add Teams meeting** ticked. That is the whole
reason this app makes meetings that way: the event and the Teams link are created in one call
(`POST /me/events` with `isOnlineMeeting`), so the meeting is calendar-associated from the start.
That path needs only `Calendars.ReadWrite` — Microsoft lists no other permission for it.

<a id="still-processing"></a>

### `Microsoft has not finished processing this meeting. Checking again around hh:mm.`

Normal. Nothing is ready the moment a meeting ends, and **Microsoft publishes no schedule or SLA
for it** — the documentation says only that an app can fetch the artifacts "when it's generated
after the meeting ends". Reported reality is five to thirty minutes for an ordinary meeting and a
few hours for a long or heavy one, and Graph lags the Teams UI: a transcript can be readable in
Teams while this endpoint still returns an empty list.

**So an empty answer is not evidence that nothing exists.** The app keeps asking on a widening
backoff — 10, 25, 45 and 75 minutes, then 2, 3, 4½, 6, 8, 10, 12, 16, 20 and 24 hours — fourteen
attempts across a day, and **Get Transcript & Recording** stays on the Event for when you would
rather not wait.

If the message says *Nothing is checking automatically* instead, the backoff has nothing to run it:
see [background jobs](#nothing-checking-automatically).

<a id="nothing-found"></a>

### `Microsoft still has no transcript or recording a day after this meeting, so it was most likely never recorded or transcribed.`

Graph answered `200` with an empty list, all day. That usually means nobody pressed record — but
from this side it is indistinguishable from a connection that cannot reach files which do exist.

**The discriminating check**, and the reason the message names it: open the meeting in the Teams
calendar and look at its **Recordings and Transcripts** tab.

| What Teams shows | What it means |
| --- | --- |
| Nothing | Nothing was recorded or transcribed. There is nothing to fetch and no setup to fix |
| The files are there | The gap is in this connection. Run **Microsoft Settings → Troubleshoot → Run Diagnostics** |

Recording and transcription are not automatic. Somebody has to start them in the meeting (**More**
→ **Record and transcribe**), or the organiser has to set it in the meeting options in advance.

<a id="meeting-expired"></a>

### `This meeting is too old, but the files may still exist.`

Two different clocks end a meeting's artifacts, and they run at different speeds.

**The meeting expires**, and Graph stops serving its transcript and recording. Microsoft's
published table (currently marked public preview):

| Meeting type | Expires |
| --- | --- |
| Scheduled meeting, one time | 60 days after the scheduled time |
| Meet now, from a calendar or channel | 60 days after the link was created |
| Recurring with an end date | 60 days from the end date, or 60 days from the last occurrence, whichever is longer |
| Recurring with no end date | 1 year after it was last accessed, joined or updated |

*"When a meeting is joined or updated before its expiration limit, an extra 60 days are added"* —
except for Meet now meetings.

**The files expire**, and Teams deletes them. The default is **120 days**, set per meeting policy
(Teams admin center → **Meetings** → **Meeting policies** → **Recording & transcription**), and an
administrator can set it from 1 day to 99,999, or to never via PowerShell. Changes affect only
newly created recordings and transcripts. Education A1 defaults to 30 days.

Graph goes quiet long before Teams deletes anything, so *"too old"* on its own would send people
away from files that are still sitting there.

**Fix.** There is no API path back. Take the files from the meeting's **Recordings and
Transcripts** tab in the Teams calendar, or from the OneDrive/SharePoint location Teams stored them
in.

<a id="speaker-attribution"></a>

### Speaker names missing from the `.vtt` — the transcript is text with nobody's name against it

Two possibilities, and they look different in the log:

- **The fetch failed outright** with a 403 → that is
  [`SpeakerAttributionNotAllowed`](#speakerattributionnotallowed), and the fix is the **Include
  speaker attribution** toggle.
- **The fetch succeeded and the names are absent** → the transcript was generated before the
  setting was turned on, or in a tenant where it was off at the time. The setting affects
  transcripts produced after the change; existing ones do not gain names retrospectively.

<a id="transcript-not-stored"></a>

### `Microsoft gave us the transcript, but Frappe could not store the file: …`

Title: *The transcript arrived but could not be saved*.

**This one is deliberately worded to point away from Microsoft.** The fetch worked. Writing the
`File` failed — and saving a File runs every `after_insert` hook any installed app has put on
`File`: an S3 offloader, a virus scanner, a storage quota. When one of those is misconfigured the
exception is raised from inside this app's call stack, and without this wrapper Frappe hands you a
`botocore` traceback with this integration's name on it, for a missing AWS credential.

**Fix.** The message carries the underlying exception. Check the apps that handle attachments on
this site — storage backend credentials, quota, the antivirus hook. Microsoft is not involved and
re-authorising changes nothing.

<a id="meeting-preconditions"></a>

### `This event has no Teams meeting.` / `This meeting has not started yet.`

Precondition checks, raised before any Graph call.

**`This event has no Teams meeting.`** The Event has no Microsoft Calendar linked, or no Teams join
URL stored. Either it was never created as a Teams meeting, or it was created in Outlook and has
not been pulled in yet. Tick **Add Teams meeting** and save, or sync the calendar.

**`This meeting has not started yet.`** There is genuinely nothing to fetch. The check is on the
**start**, never the booked end: people book an hour and talk for four minutes, and Microsoft has
the transcript ready minutes after the call actually ends, not after the slot nobody used up.

<a id="calendar-off-or-unauthorised"></a>

### `The Microsoft Calendar <name> is switched off.` / `… is not signed in to Microsoft.`

Two messages, kept apart on purpose: one is a tickbox on a form, the other is a sign-in that has to
happen in a browser.

- **Switched off** → open the Microsoft Calendar and tick **Enabled**.
- **Not signed in** → open it and click **Authorize Microsoft Access**, as the user in the **User**
  field.

<a id="recording-not-this-meeting"></a>

### `That recording does not belong to this meeting.`

A permission check, not a fault. The recording id has to be one this app recorded against this
Event. Without the check, the download endpoint would be a way to pull any recording the connected
account can reach, by guessing ids.

**Fix.** Use **Download Recording** on the Event rather than calling the endpoint with an id from
somewhere else. If the Event's recording list is stale, click **Check for Recording** to refresh
it.

<a id="nothing-checking-automatically"></a>

### `Nothing is checking automatically: …`

The catch-up job that collects transcripts after meetings is opt-in per connection, but the tickbox
only decides whether the job *would* ask — something still has to run it. On a bench with no
scheduler or no worker, nothing does.

The message carries the specific reason and the command that ends it. See
[background jobs](#background-jobs).

Saying *"Checking again around 14:48"* on a bench with no worker would be the same class of lie as
the tickbox being off, and worse, because it names a time. **Get Transcript & Recording** still
works by hand in the meantime.

---

## Outlook mail (IMAP and SMTP)

This section is about Frappe's own Email Account OAuth path. This app does not replace it: it
diagnoses it, and leaves Frappe's IMAP/SMTP behaviour exactly as it is.

`Email Account → Check Microsoft Setup` runs the diagnostics for a single account.

The Microsoft 365 settings, for reference, are `outlook.office365.com:993` with SSL/TLS for IMAP4
and `smtp.office365.com:587` with STARTTLS for SMTP. Port 465 is not usable: Microsoft's guidance
is that a device defaulting to it "doesn't support the required versions of TLS for client SMTP
submission". TLS 1.2 or above is required, and the mailbox must be licensed.

<a id="authenticate-failed"></a>

### `AUTHENTICATE failed` / `A01 NO AUTHENTICATE failed.`

IMAP rejected the token. Microsoft issued it; the mailbox would not take it. The protocol gives
back nothing more specific, which is what makes this the most-searched and least-answered string in
the whole integration.

**We cannot tell which of these it is from the error alone.** Work down the list:

| Check | Where |
| --- | --- |
| Is the IMAP scope granted and consented? | `https://outlook.office.com/IMAP.AccessAsUser.All` (delegated) or `IMAP.AccessAsApp` (application), under **APIs my organization uses** → *Office 365 Exchange Online* |
| Is the scope the **full resource URL**? | Microsoft: *"Ensure to specify the full scopes, including Outlook resource URLs."* A bare `IMAP.AccessAsUser.All` is not the same scope |
| Is `offline_access` in the Connected App's scopes? | Without it you get an hour of working IMAP and then this, forever ([details](#offline-access)) |
| **App-only:** is the token requested at `.default`? | Microsoft: *"You must use `https://outlook.office365.com/.default` in the `scope` property in the body payload for the access token request."* |
| **App-only:** is the service principal registered in Exchange? | `New-ServicePrincipal` in Exchange Online PowerShell |
| **App-only:** was this mailbox granted to it? | `Add-MailboxPermission -Identity "<mailbox>" -User <ServicePrincipal_ID> -AccessRights FullAccess`. Access is per mailbox; the app cannot reach one that was never granted |
| Is the identity right? | See [shared mailboxes](#shared-mailbox) |

**The Exchange PowerShell trap, in Microsoft's own words.** On `New-ServicePrincipal`: *"The
OBJECT_ID is the Object ID from the Overview page of the Enterprise Application node… It is **not**
the Object ID from the Overview page of the App Registrations node. Using the incorrect Object ID
will cause an authentication failure."* They are different GUIDs, both visible, and picking the
wrong one fails at authentication time with no useful message — this error, in other words.

**Microsoft Settings → Troubleshoot → Exchange Setup Script** generates those commands with the
lookup done by `AppId`, so nobody has to copy an Object ID at all, and it ends with the commands
that reverse it. Note that `Add-MailboxPermission`'s `-User` wants the **Exchange** service
principal identity from `Get-ServicePrincipal`, not the Entra object id — another pair that look
interchangeable and are not.

<a id="535-573"></a>

### `535 5.7.3 Authentication unsuccessful`

SMTP rejected the credential. Microsoft groups this with `5.7.57 Client not authenticated to send
mail` and documents four causes:

| Cause | Fix |
| --- | --- |
| **SMTP AUTH is disabled on the mailbox** — the commonest, and it is off by default in modern tenants | Microsoft 365 admin center → **Users** → **Active users** → the user → **Mail** → **Manage email apps** → tick **Authenticated SMTP** → **Save changes**. Or `Set-CASMailbox -Identity <mailbox> -SmtpClientAuthenticationDisabled $false` |
| **Multi-factor authentication on the mailbox** | Basic-auth SMTP cannot satisfy MFA. Move the account to OAuth, which is what this app is for |
| **Azure Security Defaults are on** | Security defaults disable SMTP AUTH tenant-wide |
| **A Conditional Access policy blocks legacy authentication** | Check the policies that target legacy auth clients |

Two more worth adding from this app's side: the `https://outlook.office.com/SMTP.Send` scope
missing from the Connected App, and the mailbox not being licensed. Check `Get-CASMailbox -Identity
<address> | Format-List SmtpClientAuthenticationDisabled` first — it is one command and rules out
the commonest cause.

**Organisation-level versus mailbox-level.** `Set-TransportConfig -SmtpClientAuthenticationDisabled
$false` controls the tenant default, but *"the mailbox setting takes precedence over the
organization setting"*, so setting the mailbox alone is usually enough. And an authentication
policy that disables basic authentication for SMTP overrides both: clients then cannot use SMTP
AUTH even with the settings above enabled.

<a id="535-57139"></a>

### `535 5.7.139 Authentication unsuccessful`

A variant of the above, and the documentation is thinner than the internet suggests.

**What Microsoft documents** are two federation cases: *"federated STS service was unreachable"*
and *"the federated STS URL does not support HTTPS"*. Both mean the tenant uses federated identity
and the federation endpoint could not be used — nothing to do with the credential itself.

**What you will more often actually see** is a `5.7.139` naming basic authentication as disabled.
That variant appears widely in support threads and **is not covered by any Microsoft Learn
article**, so treat descriptions of it accordingly. If that is your text, read it at face value:
basic auth is off for that mailbox, and the fix is [OAuth](#basic-auth), not a new password.

<a id="550-5760"></a>

### `550 5.7.60 SMTP; Client does not have permissions to send as this sender`

The authenticated identity is not allowed to send from the address on the message.

Microsoft's rule: the application *"must send email from the same email address that you entered as
logon credentials"*; to send from another account, the logon account needs **Send As** permission
on it. Otherwise, Microsoft's documented alternative is to use SMTP relay instead of client
submission — client submission does not support that scenario.

**Fix.** Either make the Email Account's sending address the same as the account it authenticates
as, or grant Send As: `Add-RecipientPermission -Identity "<mailbox>" -Trustee <sender>
-AccessRights SendAs`. For the app-only flow, Microsoft's note is explicit — *"If you're trying to
use Client Credential Grant Flow with SendAs, you need to grant SendAs permissions to the sender"* —
and **Exchange Setup Script** emits that command when you ask it for send access.

<a id="451-470"></a>

### `451 4.7.0 Temporary server error. Please try again later.`

A `4.x.x` code is temporary by definition — the server is telling the client to come back.

**Be aware there is no Microsoft Learn article covering this code for Exchange Online SMTP client
submission.** What exists online is largely on-premises Exchange material for a different code
(`454 4.7.0`), and it does not apply. So the honest answer is that the specific cause is not
externally documented, and what is left is the discriminating checks.

**What to do.** Retry first — a genuine temporary error clears. If it persists, confirm SMTP AUTH
is enabled for the mailbox (the `Get-CASMailbox` command above), and check the sending volume
against Microsoft's published client-submission limits: **10,000 recipients per day** and
**30 messages per minute**. A Frappe email queue that suddenly flushes hits the per-minute limit
easily.

<a id="tls-required"></a>

### `TLS required` / STARTTLS failures

The connection's encryption settings contradict each other, or the client is offering a TLS version
Exchange no longer accepts.

**Fix.** One setting or the other, never both:

| Protocol | Server | Port | Encryption |
| --- | --- | --- | --- |
| IMAP4 | `outlook.office365.com` | 993 | SSL/TLS — tick **Use SSL** |
| POP3 | `outlook.office365.com` | 995 | SSL/TLS |
| SMTP | `smtp.office365.com` | 587 | STARTTLS — tick **Use STARTTLS** |

`Run Diagnostics` flags an Email Account with both **Use SSL** and **Use STARTTLS** ticked, because
that combination is a common cause of this error. Also check the host: use the DNS name, not an IP
address, and not the consumer `smtp-mail.outlook.com`, which is Outlook.com rather than Microsoft
365.

<a id="please-authorize-oauth"></a>

### `Please Authorize OAuth for Email Account <name>`

Frappe's own message, from `frappe/email/oauth.py`. The account is set to OAuth and no access token
is stored for it.

**Fix.** Open the **Email Account** and click **Authorize API Access**, signed in as the user in
**Connected User**. The token is stored against that Frappe user, so authorising as anybody else
puts it in the wrong place.

If the account uses the app-only (backend) flow there is no Connected User and no browser sign-in.
In that case this message means the client-credentials token call is failing — look for an `AADSTS`
code in the Error Log and read its section above.

<a id="shared-mailbox"></a>

### SMTP works but IMAP does not (or the reverse) on a shared mailbox

The commonest shape of this is a shared mailbox that receives fine and cannot send, or sends fine
and never receives.

**What Microsoft documents**, and it is less than the folklore:

- For shared-mailbox access over OAuth, *"an application needs to obtain the access token on behalf
  of a user but replace the **userName** field in the SASL XOAUTH2 encoded string with the email
  address of the shared mailbox"*. This is written protocol-agnostically — it is not stated as an
  IMAP-only rule.
- Separately, for SMTP client submission, the sending address must be the address the client
  authenticated as, unless that account holds **Send As** on the other one. See
  [`550 5.7.60`](#550-5760).

**Where Frappe sits.** Frappe computes one identity — `login_id or email_id` — and hands it to both
protocols (`email_account.py` builds it; `receive.py` and `smtp.py` both consume it). So a single
Email Account cannot present the shared mailbox address to IMAP and the signing-in user to SMTP.
Whether that is a problem depends on whether the signing-in account has Send As on the shared
mailbox.

**Fix — three ways out, in order of how little they disturb:**

1. **Grant Send As** on the shared mailbox to the account the Email Account signs in as
   (`Add-RecipientPermission … -AccessRights SendAs`). If that is all that was wrong, one account
   then works for both directions.
2. **Split the account in two.** One Email Account with **Enable Incoming** only, carrying the
   shared mailbox address; another with **Enable Outgoing** only, carrying the signing-in user.
3. **Use the app-only flow.** Client credentials involve no user identity at all, so the question
   disappears. It needs `New-ServicePrincipal` plus `Add-MailboxPermission` per mailbox in Exchange
   Online — **Troubleshoot → Exchange Setup Script** generates both, scoped to the mailboxes you
   name and nothing else, with the undo commands appended.

`Run Diagnostics` flags any account whose configuration puts it in this position, as a warning
rather than a failure, because whether it actually breaks depends on the Send As grant.

<a id="basic-auth"></a>

### Basic authentication stopped working, or is about to

**IMAP and POP: already gone.** Microsoft's position is unambiguous — *"Basic authentication is now
disabled in all tenants"*, covering Exchange ActiveSync, POP, IMAP, Remote PowerShell, EWS, OAB,
Autodiscover and the Outlook desktop clients, and *"no one (you or Microsoft support) can re-enable
Basic authentication in your tenant"*. If Frappe IMAP stopped working against Microsoft 365 with no
change on your side, this is why.

**SMTP AUTH: on its own timeline, and the dates have moved.** Microsoft Learn deliberately no
longer states them, deferring to the Exchange team's announcement. The current published plan is
that behaviour is unchanged until December 2026; at the end of December 2026 SMTP AUTH basic
authentication becomes **disabled by default for existing tenants**, with administrators still able
to re-enable it; tenants created after that have it **unavailable by default**; and a final removal
date is to be announced in the second half of 2027. Because this has been revised more than once,
**check Microsoft's current announcement before planning around any date, including this one.**

**Fix, and it is the same either way.** Move the Email Account to OAuth: set **Authentication
Method** to **OAuth**, link the `Connected App` (created by **Microsoft Settings → Set Up**), set
**Connected User**, and click **Authorize API Access**. `Run Diagnostics` lists every account still
on basic authentication as a note.

---

## Installation, database and background jobs

<a id="unknown-column"></a>

### `Unknown column 'custom_sync_with_microsoft_calendar' in 'WHERE'` (or any `custom_…microsoft…` column)

The app's custom fields are not on this site. Every calendar query reads them, so without them the
sync fails with a raw SQL error that tells an administrator nothing.

**Fix.**

```bash
bench --site <site> migrate
```

On Frappe Cloud, use **Migrate** in the site dashboard. The fields are created on install and on
migrate, so this means one or the other has not run since the app was added or updated.

<a id="data-too-long"></a>

### `Data too long for column '…' at row 1` / `(1406, …)`

A Microsoft identifier is longer than the Frappe column holding it. Graph ids run well past Frappe's
default 140 characters — one real calendar id measured 152.

**This is nastier than it looks.** The write is refused, the value is silently lost, and everything
on screen still reports success: sign-in "worked" and the calendar id was never stored.

**Fix.** `bench --site <site> migrate` (Frappe Cloud: **Migrate** in the site dashboard) to pick up
the widened columns, then **Re-authorize** the connection so the values are written again.

<a id="background-jobs"></a>

### `The scheduler is off` / `Redis cannot be reached` / `No background worker is running`

Queued work that never runs is the one failure in this app that produces **no error at all** — no
traceback, no log line, no wrong answer, just a calendar that quietly stops moving and transcripts
that never arrive. `Run Diagnostics` reports it explicitly for that reason.

| Reported | Fix |
| --- | --- |
| The scheduler is off for this site | `bench --site <site> enable-scheduler` |
| …and the site is in maintenance mode | `bench --site <site> set-maintenance-mode off` first — core refuses to enable the scheduler while it is on |
| Redis, which holds the job queue, cannot be reached | Start Redis; check `redis_queue` in `common_site_config.json`; then `bench doctor` |
| No background worker is running | `bench worker --queue default`, or in production check `supervisorctl status` for the worker processes |
| How many workers are running could not be read | `bench doctor` — Redis answered but the worker registry did not |

**`No background worker is running` is the one that looks like success.** The queue accepts every
job and starts none, so the interface says "queued" and means "never".

A backlog on its own — *n jobs are waiting in the queue* — is a warning, not a fault: the jobs drain
once a worker runs. `bench purge-jobs --site <site>` drops them instead. Pass `--site`, or it
empties the queue for every site on the bench.

<a id="jobs-stale"></a>

### `This app's scheduled jobs have stopped running` / `…are not registered`

Two different findings.

**Not registered** — the scheduled jobs are not on this site at all. `bench --site <site> migrate`
registers them.

**Stopped running** — they are registered, they are supposed to run every 15 minutes, and the last
run was much longer ago. This is the most honest signal available, because it measures the thing
itself rather than its preconditions: a job that should run four times an hour and last ran
yesterday is not running, whatever the scheduler setting, Redis and the worker count claim.

**Fix.** The scheduler is enabled but its process may not be running: `bench doctor`, and
`supervisorctl status` in production.

---

<a id="nothing-matches"></a>

## When nothing here matches

Twenty-five patterns is not omniscience, and pretending otherwise costs trust. **Explain an Error**
answers *Unrecognised error* rather than guessing, and so does this page.

What is worth doing next:

1. **`Microsoft Settings → Troubleshoot → Run Diagnostics`.** It reads the settings, the Connected
   App, every Email Account, the authorised scopes, the custom fields and the job queue, and names
   the failing step rather than the symptom.
2. **Read the whole Graph error, not the status code.** This app passes through both halves —
   the `code` is what you search for, the `message` is the only part that says *which* property or
   resource Graph objected to. On the transcript endpoints, Microsoft's own instruction is to branch
   on `innerError.code`, because the message text is subject to change.
3. **Check the Error Log.** This app's entries are titled `MS Calendar …`, `MS event …` or
   `MS transcript …`, so filtering on `MS ` finds them.
4. **Look at the meeting from the Teams side.** For anything involving transcripts or recordings,
   the **Recordings and Transcripts** tab on the meeting in the Teams calendar settles whether the
   files exist at all — which splits "nothing was ever recorded" from "this connection cannot reach
   it" without anyone guessing.

If you find a cause this page does not carry, the decoder lives in `error_patterns()` in
[`frappe_microsoft365/doctor.py`](../frappe_microsoft365/doctor.py). Adding a pattern there adds it
to **Explain an Error** for everyone.
