## What and why

<!-- What this changes, for whom, and why. `Fixes #N` for the issue it closes. -->

## Does this change anything for sites already running the app?

<!-- Tick one, and fill in what it asks for. RELEASING.md explains each. -->

- [ ] **none**
- [ ] **upgrade note**: a behaviour change an admin will notice. Line added under *Upgrade notes*:
- [ ] **patch**: data is changed on `bench migrate`. Patch module, what it changes, test that runs it twice:
- [ ] **migrate warning**: an upgrade check reports data an admin has to look at. Check registered under `microsoft365_upgrade_checks` for version:

## Checked

- [ ] Tests added or updated, and the suite passes locally
- [ ] Works on Frappe 15 (Python 3.10) and Frappe 16, or the capability is refused on the version that cannot run it
- [ ] `CHANGELOG.md` (or `release-notes/<version>/<topic>.md` on a release branch) updated, including *Upgrade notes*
