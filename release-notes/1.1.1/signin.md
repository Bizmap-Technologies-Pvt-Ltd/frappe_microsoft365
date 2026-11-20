### Security
- Sign in with Microsoft only accepts accounts from the directory named in Microsoft Settings. Set Up refuses to create the sign-in record while Tenant ID is `common`, `organizations` or `consumers`, the doctor reports such a record as a failure, and on Frappe 16 a sign-in from any other directory is refused.
- Values written into the generated Site Grant and Exchange PowerShell scripts are always literal text: they are validated (client ids as GUIDs, mailboxes as addresses, SharePoint sites on `*.sharepoint.com`) when saved and again when a script is generated, and written as single-quoted strings.

### Fixed
- The Site Grant Script grants the Application Access registration when one is set, the app SharePoint calls are actually made as, instead of the main one, which left every upload at 403. (Fixes #36)
- The Exchange Setup Script runs as pasted: it signs in to Microsoft Graph (`Connect-MgGraph -Scopes Application.Read.All`) before looking up the service principal, no longer suggests the retired AzureAD module, and finds Exchange's record of the app by AppId rather than by a display name every site in the tenant shares. Running it a second time no longer fails. (Fixes #45)

### Changed
- Microsoft Settings refuses a Client ID or Application Client ID that is not a GUID, and SharePoint Mapping refuses a site address that is not on `<tenant>.sharepoint.com` or whose site name has characters a SharePoint address cannot have.
- The Exchange Setup Script refuses mailboxes that are not plain addresses.
- The doctor checks the Application Access registration: a client id with no secret, a secret that is the Secret ID, and a client id that is not a GUID are reported as failures.

### Upgrade notes
- **Sign-in on a multi-tenant Tenant ID.** On migrate, a site whose Tenant ID is `common`/`organizations` and that has a Microsoft Social Login Key gets Tenant ID filled in from the authorised Microsoft Calendars, when every one of their stored tokens names the same work or school directory. The Social Login Key's `/common/` or `/organizations/` URLs are pointed at that directory too, when the key uses the Client ID from Microsoft Settings. Nothing is sent to Microsoft. If that is not possible (no calendars, tokens from several directories, personal accounts, or a key naming another directory), nothing changes, sign-in keeps working in 1.1.1, and the upgrade report and the doctor show a failure with the fix: set Tenant ID to your Directory (tenant) ID, make the app registration single-tenant, and point the Social Login Key's URLs at `https://login.microsoftonline.com/<tenant id>/oauth2/v2.0/...`. **From 1.2.0, Microsoft sign-in on a multi-tenant Tenant ID is refused.**
- After the change, on Frappe 16, accounts from directories other than yours can no longer sign in. Partner users who need access should be invited into your directory as guests.
- Frappe 15 has no sign-in hook, so there the protection is the tenant-specific URLs on the Social Login Key; make sure the doctor shows no `sso.tenant_*` failure.
- **Separate Application Access registration.** If you use one for SharePoint, run the corrected Site Grant Script once (SharePoint Settings or each mapping); the old script granted the main registration.
- **Stricter values.** The upgrade report lists any saved Client ID, Application Client ID or SharePoint Mapping site address that would now be refused, so it can be corrected before someone next saves the form. Nothing stops working until then.
