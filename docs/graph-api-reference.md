# Microsoft Graph + MSAL — implementation reference (verified against learn.microsoft.com, v1.0)

This is the contract the app implements. Endpoints are Graph **v1.0**. Auth via **MSAL Python**.

## Azure AD app registration (the admin does this once)
- Register an app (single tenant or multi-tenant) in Entra/Azure portal.
- **Redirect URI** (Web platform): `https://<site>/api/method/frappe_microsoft365.frappe_microsoft_365.doctype.microsoft_calendar.microsoft_calendar.callback`
  (local: `http://bizmapos.localhost:8090/api/method/...callback` — http allowed for localhost).
- **Client secret** → pasted into Microsoft Settings (stored as Password).
- **Delegated API permissions** (Microsoft Graph) are DERIVED from the capabilities ticked in
  Microsoft Settings, not a fixed list. `offline_access`, `openid` and `profile` are always
  added by the app; on top of those:
  - Outlook calendar (including the Teams meeting tickbox on an Event): `User.Read`,
    `Calendars.ReadWrite`
  - Standalone Teams meetings: `+ OnlineMeetings.ReadWrite`
  - Transcripts and recordings: `+ OnlineMeetingTranscript.Read.All` (needs admin consent)

  The form prints the exact list to grant. Do not grant more than that: asking for scopes the
  tenant never consented to is what makes sign-in fail.

## MSAL (msal>=1.30, installed 1.37)
```python
import msal
app = msal.ConfidentialClientApplication(
    client_id,
    client_credential=client_secret,           # the secret string
    authority=f"https://login.microsoftonline.com/{tenant_id}",  # or /common, /organizations
)
# Authorize URL (we build/redirect):
#   {authority}/oauth2/v2.0/authorize?client_id=..&response_type=code&redirect_uri=..
#     &response_mode=query&scope=<space-joined>&state=<random>
# Callback: exchange code (no PKCE needed for confidential client):
result = app.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
# result keys: access_token, refresh_token, expires_in, id_token_claims{preferred_username,oid,...}, scope
# Refresh:
result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)
# On error: result has "error" + "error_description".
```
- SCOPES for token calls: the resource scopes WITHOUT reserved ones (msal injects openid/profile/offline_access).
  Built by `microsoft_graph.derive_scopes()` from the ticked capabilities, e.g. calendar-only
  gives `["User.Read","Calendars.ReadWrite"]`.
- Store `refresh_token` (Password), `access_token` (Password), and `token_expiry` (now + expires_in - 300s skew).
- Validate `state` ourselves (store random state on the Microsoft Calendar doc; verify in callback) for CSRF.

## Graph base
`https://graph.microsoft.com/v1.0`  · header `Authorization: Bearer <access_token>`.

### Who am I
`GET /me` → id, userPrincipalName, mail, displayName.

### Calendars
- `GET /me/calendars` → list (id, name, isDefaultCalendar).
- List events: `GET /me/events?$select=...&$top=50&$orderby=start/dateTime`
- Calendar view (expanded recurrences, by range):
  `GET /me/calendarView?startDateTime=<ISO>&endDateTime=<ISO>` (header `Prefer: outlook.timezone="UTC"`).
- Create event (also creates the Teams meeting when isOnlineMeeting):
  `POST /me/events`
  ```json
  {
    "subject": "...",
    "body": {"contentType": "HTML", "content": "..."},
    "start": {"dateTime": "2026-07-01T10:00:00", "timeZone": "Asia/Kolkata"},
    "end":   {"dateTime": "2026-07-01T10:30:00", "timeZone": "Asia/Kolkata"},
    "location": {"displayName": "..."},
    "attendees": [{"emailAddress": {"address": "x@y.com", "name": "X"}, "type": "required"}],
    "isOnlineMeeting": true,
    "onlineMeetingProvider": "teamsForBusiness"
  }
  ```
  Response: `id`, `webLink`, `onlineMeeting.joinUrl`, `iCalUId`. Scope: `Calendars.ReadWrite`.
- Update: `PATCH /me/events/{id}`. Delete: `DELETE /me/events/{id}`.

### Teams online meeting (standalone — NOT calendar-associated)
`POST /me/onlineMeetings` body `{startDateTime, endDateTime, subject}` → `{id, joinWebUrl}`.
Scope `OnlineMeetings.ReadWrite`. NOTE: transcripts API does NOT support meetings created this way.
**Therefore prefer creating a calendar Event with isOnlineMeeting=true for anything needing transcripts.**

### Map a join URL → onlineMeeting id (needed for transcripts)
`GET /me/onlineMeetings?$filter=JoinWebUrl eq '{joinUrl}'` → value[0].id.

### Transcripts (Teams)
- List: `GET /me/onlineMeetings/{onlineMeetingId}/transcripts` → value[].{id, transcriptContentUrl, createdDateTime}.
- Content: `GET /me/onlineMeetings/{omId}/transcripts/{id}/content?$format=text/vtt` (or text/plain).
- Permission (delegated): **OnlineMeetingTranscript.Read.All** (admin consent). Personal accounts unsupported.
- Caveats: only for meetings **associated with a calendar event**; meeting must **not be expired**;
  application-permission path needs an **application access policy** granted by a tenant admin.

### Recordings (similar shape)
`GET /me/onlineMeetings/{omId}/recordings` + `/content` (perm `OnlineMeetingRecording.Read.All`). Same caveats.

## Sync model (mirrors Frappe Google Calendar)
- `Microsoft Settings` (Single): enabled, tenant_id, client_id, client_secret(Password), redirect_uri, scopes.
- `Microsoft Calendar` (per user): user, email, calendar id/name, refresh/access tokens(Password), expiry,
  push_to_microsoft (Frappe Event → Graph), pull_from_microsoft (Graph → Frappe Event), last_sync, state.
- Frappe `Event` custom fields: `custom_microsoft_event_id`, `custom_microsoft_calendar`,
  `custom_sync_with_microsoft_calendar` (mirror Google's `google_calendar*` fields).
- Scheduled job (cron) calls a sync that pulls changed events and pushes Frappe events both ways,
  storing the Graph event id on the Frappe Event to dedupe. Use `@odata.deltaLink` / updated filters where possible.

## Errors / robustness
- 401 → refresh token once, retry. 403 → permission/consent missing (surface clearly).
- 429 → respect `Retry-After`. Never log tokens/secret. All Graph calls go through one helper that
  injects the bearer, refreshes on 401, and raises a clean error otherwise.
