# Azure setup — step by step

This is the one-time Azure/Entra configuration that makes `frappe_microsoft365` work. You need a
Microsoft Entra ID (Azure AD) tenant where you can register an application and grant admin consent.

The endpoint/permission contract these steps satisfy is documented in
[`graph-api-reference.md`](graph-api-reference.md).

---

## 1. Register the application

1. Go to the **Microsoft Entra admin center** (<https://entra.microsoft.com>) or the Azure portal →
   **Microsoft Entra ID** → **App registrations** → **New registration**.
2. **Name:** anything recognizable, e.g. `Frappe Microsoft 365`.
3. **Supported account types:** *Accounts in this organizational directory only (single
   tenant)*, and paste your **Directory (tenant) ID** (or your tenant's domain, such as
   `contoso.onmicrosoft.com`) into **Tenant ID** in Frappe.
   - **Sign in with Microsoft requires single tenant.** Frappe signs in the user whose email
     address is in Microsoft's token. On the `common` or `organizations` authority that token can
     come from *any* organisation's directory, and whoever runs that directory chooses its users'
     addresses — including ones that match your Frappe users. Set Up refuses to create the sign-in
     record unless Tenant ID names your directory, the doctor reports an existing one on a
     multi-tenant authority as a failure, and on Frappe 16 a sign-in from any other directory is
     refused. People from partner organisations who need access should be invited into your
     directory as guests, which keeps them inside it.
   - App-only features (SharePoint document storage, app-only mail) need a specific tenant too.
   - Personal Microsoft accounts (consumer) are **not** supported for the Teams/transcript scopes.
4. Leave the redirect URI blank for now (added in the next step). Click **Register**.

## 2. Add the Web redirect URI

App → **Authentication** → **Add a platform** → **Web** → add the **Redirect URI** that points at
the app's OAuth callback:

```
https://<site>/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback
```

- Replace `<site>` with your real site host (e.g. `https://erp.example.com/...`).
- For **local development**, `http://` is allowed for `localhost`-style hosts, e.g.:
  ```
  http://bizmapos.localhost:8090/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback
  ```
- This must **exactly** match what Frappe sends. If you override **Redirect URI** in Microsoft
  Settings, the two must be identical (scheme, host, port, path).
- Platform must be **Web** (confidential client with a secret), not SPA/Public.

## 3. Create a client secret

App → **Certificates & secrets** → **Client secrets** → **New client secret**.

- Give it a description and an expiry (note the date — you'll need to rotate before it expires).
- **Copy the secret Value immediately** (it is shown only once; the "Secret ID" is not the value).

## 4. API permissions (Microsoft Graph, Delegated)

App → **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**,
then add all of:

```
offline_access
openid
profile
User.Read
Calendars.ReadWrite
OnlineMeetings.Read
OnlineMeetingTranscript.Read.All
```

(`OnlineMeetings.ReadWrite` instead of `OnlineMeetings.Read` if you also tick **Standalone Teams
meetings**, or if this site was set up before 1.2.0 and already holds it.)

Then click **Grant admin consent for <tenant>** (the green check). Admin consent is required —
`OnlineMeetingTranscript.Read.All` in particular will not work without it.

> Add `OnlineMeetingRecording.Read.All` (delegated) as well if you want Teams **recordings**.
> Microsoft consents to reading the words and reading the video separately, so the transcript scope
> does not cover recordings — and a tenant that refuses recordings still gets transcripts.
> Recordings may additionally require Teams Premium / appropriate licensing.

## 5. Copy the identifiers

From the app's **Overview** page, copy:

- **Application (client) ID**
- **Directory (tenant) ID**

…plus the **client secret value** from step 3.

## 6. Configure Frappe

1. In your Frappe site: open **Microsoft Settings**.
2. Paste **Tenant ID**, **Client ID**, **Client Secret**.
3. (Optional) Set **Redirect URI** only if you need to pin a non-default host — it must match the
   Azure value from step 2 exactly. Leave blank to use the site URL automatically.
4. (Optional) **Default Scopes** — defaults to the delegated scopes above; only change if you know
   what you're doing.
5. Tick **Enabled** and save.
6. Create a **Microsoft Calendar** record (set **User** to the person connecting), save, then click
   **Authorize**. You'll be redirected to Microsoft to sign in and consent; on return the record is
   marked **Authorized** and the account email / default calendar are filled in. Use **Test
   Connection** to verify.

---

## Transcript caveats (read this)

Teams transcripts are the most permission-sensitive feature. They only work when **all** of these
hold:

- **Graph access to transcripts is switched on for the tenant.** Microsoft added this control in
  2026 and **ships it off**, in a different portal from the permissions: Teams admin center >
  Meetings > Meeting settings > Transcript API access > Microsoft Graph access > On. Until it is
  on, every transcript call returns `403 Forbidden: Graph API access to transcripts is disabled for
  this tenant` (inner code `GraphAccessToTranscriptsDisabled`) **no matter how much consent has
  been granted**. Speaker names additionally need Configure > Include speaker attribution > On.
- The **delegated** scope `OnlineMeetingTranscript.Read.All` is granted **with tenant admin
  consent** (step 4).
- The **delegated** scope `OnlineMeetings.Read` is granted too (or `OnlineMeetings.ReadWrite`, on a
  site that already had it — **Run Diagnostics** says which one this site requests). A transcript
  is addressed by online meeting id, and the only way to get that from a join link is
  `GET /me/onlineMeetings?$filter=JoinWebUrl eq '...'`, which this scope covers. Granting the
  transcript scope alone is a key to a door you cannot walk to.

  A `403` is therefore not one diagnosis but several: read the message body rather than assuming
  consent is missing, or run **Run Diagnostics**, which decodes it.
- The meeting is **calendar-associated** — i.e. created via `POST /me/events` with
  `isOnlineMeeting=true` (this app's `create_meeting(..., create_calendar_event=True)`, which is the
  default). Standalone `POST /me/onlineMeetings` meetings are **not** calendar-associated and will
  never expose transcripts.
- The meeting has **not expired** — Graph serves a meeting's artifacts for about **60 days** after
  it happens (a join or an edit adds another 60). Separately, Teams deletes the files themselves on
  the tenant's retention policy, 120 days by default. Whichever comes first ends it.
- A transcript actually exists — transcription must have been turned on during the call, and there
  is a processing delay after the meeting ends before the VTT is available.

If you need to read transcripts for meetings the signed-in user did not organize (an
**application-permission** path rather than delegated), Microsoft additionally requires configuring
an **application access policy** in Teams PowerShell to scope which users' online meetings the app
can read. That path is outside the default delegated flow this app uses.

See [`graph-api-reference.md`](graph-api-reference.md) for the exact endpoints and the verified
permission matrix.
