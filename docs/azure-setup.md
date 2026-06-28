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
3. **Supported account types** — pick based on who signs in:
   - *Accounts in this organizational directory only (single tenant)* — recommended when every user
     belongs to your own tenant. Use your **Directory (tenant) ID** as the authority.
   - *Accounts in any organizational directory (multitenant)* — only if users from other Microsoft
     work/school tenants must connect. Multitenant apps typically authenticate against the
     `organizations` (or `common`) authority and may require admin consent in each guest tenant.
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
OnlineMeetings.ReadWrite
OnlineMeetingTranscript.Read.All
```

Then click **Grant admin consent for <tenant>** (the green check). Admin consent is required —
`OnlineMeetingTranscript.Read.All` in particular will not work without it.

> Optional: add `OnlineMeetingRecording.Read.All` (delegated) if you also want Teams **recordings**.
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

- The **delegated** scope `OnlineMeetingTranscript.Read.All` is granted **with tenant admin
  consent** (step 4). A `403`/`Forbidden` from Graph almost always means the scope or admin consent
  is missing.
- The meeting is **calendar-associated** — i.e. created via `POST /me/events` with
  `isOnlineMeeting=true` (this app's `create_meeting(..., create_calendar_event=True)`, which is the
  default). Standalone `POST /me/onlineMeetings` meetings are **not** calendar-associated and will
  never expose transcripts.
- The meeting has **not expired** — Graph only serves transcripts for recent meetings; very old
  meetings drop off.
- A transcript actually exists — transcription must have been turned on during the call, and there
  is a processing delay after the meeting ends before the VTT is available.

If you need to read transcripts for meetings the signed-in user did not organize (an
**application-permission** path rather than delegated), Microsoft additionally requires configuring
an **application access policy** in Teams PowerShell to scope which users' online meetings the app
can read. That path is outside the default delegated flow this app uses.

See [`graph-api-reference.md`](graph-api-reference.md) for the exact endpoints and the verified
permission matrix.
