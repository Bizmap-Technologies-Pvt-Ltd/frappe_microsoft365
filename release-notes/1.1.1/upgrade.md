### Added

- **Upgrade report.** After `bench migrate`, the app checks the site for anything the new release
  changes or needs an admin to look at, once per release. What it finds is printed at the end of
  the migrate output, shown at the top of Microsoft Settings > Troubleshoot > *Run Diagnostics*
  as *After upgrading to X.Y.Z* until a System Manager dismisses it, and sent to every System
  Manager as one notification. A site with nothing to report gets no notification, and a new
  install reports nothing. (#53)
- CI upgrades a site running the previous release, with calendars, Events and (on Frappe 16)
  SharePoint data in it, to every change before it is merged, on Frappe 15 and 16. (#54)
- Every release's changelog now ends with *Upgrade notes*: behaviour changes, patches that run on
  migrate, warnings to expect and manual steps. Pull requests say whether they change anything
  for existing sites. (#55)

### Upgrade notes

- `bench migrate` now ends with a `Microsoft 365 1.1.1:` line: either *nothing to report*, or the
  items to check. Each item is also listed in *Run Diagnostics* and in the System Managers'
  notifications; the line and the notification appear on the first migrate to 1.1.1 only.
- Microsoft Settings gets two hidden fields that record which release was last reported and what
  it found. Nothing to do.
