### Security

- **Sign in with Microsoft on Frappe 15:** a sign-in is accepted only with a sign-in state that
  this site's login page issued in the last ten minutes, and only once, as Frappe 16 already does.
  Frappe 15 itself accepts any state.

### Fixed

- **Sign in with Microsoft on Frappe 15** now refuses accounts from another organisation's
  directory. The check ran only on Frappe 16, through a hook Frappe 15 does not have; it now runs
  in this app's sign-in callback on both versions. (#64)
- Rotating the client secret in Microsoft Settings no longer breaks mail or sign-in unnoticed:
  the doctor and Set Up report a Connected App or Social Login Key whose copy of the secret
  differs. Secrets are compared on the server and never shown. (#44)
- Upgrading from before 1.0.0 no longer switches Teams meetings and transcripts off: where
  calendars had been authorised with the old built-in scope list, clearing it ticks *Standalone
  Teams meetings* and *Meeting transcripts and recordings*. (#46)

### Changed

- **Sign in with Microsoft is refused while Tenant ID is `common`, `organizations`,
  `consumers` or empty**, as announced in 1.1.1. People see a page asking them to contact their
  administrator, who has to set Tenant ID to the directory's ID. (#65)
- A refused Microsoft sign-in now shows a page saying why, instead of Frappe's *Session
  Expired* page, and is refused before Frappe updates the user record.
- On Frappe 15 the doctor shows a warning while Microsoft sign-in is enabled: Frappe 15 does not
  protect other social login providers against login CSRF. It fails if another app has taken over
  the Microsoft sign-in callback, which would switch this app's checks off.

### Added

- **Sync Secret** in the Set Up dialog (System Managers): copies the client secret from
  Microsoft Settings to the Connected App and Social Login Key that still hold an old one. Only
  records made for the same app registration are touched, and only their secret.

### Upgrade notes

- **Manual steps**
  - **Before upgrading, if Sign in with Microsoft is enabled and Tenant ID is still `common`,
    `organizations`, `consumers` or empty, set Tenant ID first.** From 1.2.0 every Microsoft
    sign-in on such a site is refused (password sign-in is not affected). Put the *Directory
    (tenant) ID* from the app registration's Overview page into Microsoft Settings > Tenant ID,
    and point the Social Login Key's Authorize URL and Access Token URL at
    `https://login.microsoftonline.com/<tenant id>/oauth2/v2.0/authorize` and `.../token`. The
    upgrade report fails with `upgrade.sign_in_refused` on any site still in this state; 1.1.1's
    report (`upgrade.sign_in_tenant`) and the doctor already listed it. Sites without Microsoft
    sign-in see nothing.
  - Sites that upgraded from before 1.0.0 to 1.0.x–1.1.x may have had *Standalone Teams meetings*
    and *Meeting transcripts and recordings* switched off by the upgrade, so transcripts stopped
    being fetched. The old setting is gone and cannot be restored automatically. If the upgrade
    report lists `upgrade.meeting_capability_off`, tick the capabilities again in Microsoft
    Settings, save, and have each user authorise their Microsoft Calendar again.
- **Behaviour changes**
  - Frappe 15: Microsoft sign-ins from another organisation's directory are refused (see #64).
  - Frappe 15: Microsoft sign-in needs the state issued by Frappe's own login page. Anyone who
    had the login page open while the site was upgraded gets *Your login attempt is invalid or
    has expired* once, and signs in again. A site that replaced Frappe's login page with its own
    must still build the Microsoft button through the login page's `provider_logins` context, or
    every Microsoft sign-in is refused; see docs/compatibility.md.
  - The Microsoft sign-in callback (`frappe.integrations.oauth2_logins.login_via_office365`) is
    now routed through this app (`override_whitelisted_methods`) on Frappe 15 and 16. The Social
    Login Key's redirect URL does not change.
- **Patches that run on migrate:** `clear_legacy_default_scopes` (already listed, changed): on a
  site that has not run it yet, it now ticks Teams meetings and transcripts when it clears the
  pre-1.0.0 scope list and at least one Microsoft Calendar is authorised. Sites that ran it
  before are not touched again. The transcripts capability also asks for
  `OnlineMeetingRecording.Read.All`, which the old list did not, so on such a site the doctor asks
  for that permission to be consented and calendars to be authorised again.
- **Warnings to expect**
  - `upgrade.sign_in_refused` (FAIL): Microsoft sign-in is enabled with a multi-tenant Tenant ID
    and is now refused. See the first manual step.
  - `upgrade.meeting_capability_off` (WARN): Events hold Teams meeting ids or fetched transcripts
    while the matching capability is off.
  - The doctor may report `secret.drift` on sites that rotated their client secret in Microsoft
    Settings only. That is the existing breakage being surfaced; use *Set Up > Sync Secret* if
    Microsoft Settings holds the current secret.
  - Frappe 15 sites with Microsoft sign-in enabled see `sso.frappe15_state` (WARN) in the doctor
    from now on.
