# Changelog

## Unreleased — SharePoint document storage

SharePoint document storage was built by David Eleben (@eleben) in his fork
(`eleben/frappe_microsoft365`, branch `feature/sharepoint-files`) and brought into this app as a
capability, with the changes listed under *Changed from the original*.

### Added

- **Document Storage**: a *SharePoint Mapping* per DocType points it at a SharePoint site,
  library and base folder (a Teams channel's Files tab is just that), with its own Test, Site
  Grant Script and Move Existing Attachments; the global switch and options live in Microsoft
  Settings.
- **Automatic or On demand** per mapping: on demand, a record's files move only once someone
  creates its folder (from the record, or a row in the mapping's Folders table); a row can
  also link a folder that already exists in SharePoint. Each record gets its own folder; new
  attachments are moved there by a background job and keep opening from the form, and the copy
  on disk is removed unless *Keep a local copy* is ticked. Large files go through resumable
  upload sessions. Failed moves are marked on the File and retried hourly.
- **Files in Microsoft 365** panel on mapped forms: the folder's contents, including files added
  from Teams or SharePoint, opened through Frappe with the record's own read permission.
  *Files Panel* on each mapping puts it at the top of the form (default), in its own
  *Microsoft 365 Files* tab, or hides it.
- **App-only Graph calls**: `graph_request(..., microsoft_graph.APP)` uses the client
  credentials flow, with the token cached encrypted per site. Document storage needs the
  `Sites.Selected` application permission and a per-site grant — never tenant-wide
  `Sites.ReadWrite.All`.
- **SharePoint Settings**: its own page in the sidebar for the storage options, attachment counts,
  tools and the site grant script; Microsoft Settings keeps only the capability tickbox.
- The app's sidebar is shipped in both v16 formats: `<module>/sidebar/` (Frappe 16.4x and later)
  and `workspace_sidebar/` (earlier v16 releases), so every page is listed on either.
- Record folders are the *Folders* table of their SharePoint Mapping (Add Row to create or link
  one, Delete to forget one); the sidebar's
  SharePoint section lists only the mappings.
- Per-mapping *Options*: a SharePoint Mapping can set *Keep a local copy* and *Remove from
  SharePoint when the attachment is deleted* to On or Off for its DocType, or follow SharePoint
  Settings (Default).
- [User guide](docs/sharepoint-user-guide.md) for SharePoint document storage: everyday use, setup on
  both the Microsoft and Frappe sides, administration and troubleshooting.
- Doctor: Document Storage checks, a *Test SharePoint Connection* that proves write access per
  mapping, and a *SharePoint Site Grant Script* that grants the app only the mapped sites.
- *Move Existing Attachments* queues files attached before a DocType was mapped. Optional
  archiving of linked attachments, SSRF-guarded to public addresses.

- **Clean uninstall**: attachments in SharePoint are brought back to the server (only if the disk
  has room) or turned into SharePoint links, as chosen in Microsoft Settings; the uninstall stops
  rather than guess or leave one broken. Then every custom field the app added to Event and File
  is deleted and its column dropped. A dry run only reports.

Nothing is deleted from SharePoint unless *Remove from SharePoint when the attachment is
deleted* is ticked, and then only to the site recycle bin. Deleting a record keeps its folder.

### Changed from the original

- **One capability registry** (`capabilities.py`): each tickbox on Microsoft Settings is defined
  once — label, whether it acts as a person or as the app, Azure permissions, prerequisites,
  minimum Frappe version, and what to create when it is switched on. Scopes, Set Up, the doctor
  and the Settings form all read it.
- **Off means off**: the File fields are created when storage is switched on, not on install, and
  every hook returns before doing any work while it is off.
- **App auth context** `microsoft_graph.APP` replaces the `APP_ONLY` marker; an optional separate
  *Application Access* registration on Microsoft Settings; the app token is keyed to the tenant
  and client that minted it and dropped whenever Microsoft Settings is saved.
- **Frappe v16 only** for storage: the tickbox is refused on v15, where moved files could not be
  read server-side.
- **Only known-safe types open inline** (PDF, images other than SVG, plain text, audio, video);
  everything else, including files without an extension, downloads.
- `File.get_content` no longer depends on a Frappe constant some v16 releases lack.
- The fork's intermediate-DocType migration patches are not carried over: the DocTypes are
  created in their final form, and a site upgrading from the fork already has them.

### Fixed from the original (found in testing)

- **A file was uploaded again every hour.** A SharePoint address longer than 140 characters
  (any long file or record name) could not be saved after the upload, the attempt was never
  counted, and the hourly retry uploaded the file again, as `name (1).pdf`, `name (2).pdf`, ...
  Addresses, paths and names now have room (1000), every attempt counts before any work, and a
  retry resumes from the upload it already made. Copies already in SharePoint are not deleted:
  remove the extras by hand.
- **The Files panel's Modified column showed raw HTML**, and its times were off by the site's
  UTC offset (a file saved two minutes ago showed as hours old on an India site).
- **Script changes did not reach browsers after a deploy.** The panel script is now a bundle with
  a new file name for every version, like the doctor script in 1.0.2.

## 1.0.2 — 2026-10-23

### Fixed

- **New versions reach browsers after a deploy.** The desk script was loaded by a file name
  that never changed, and browsers keep `/assets` files for up to a year, so after an update
  people kept running the old script — even after migrate, clearing the cache and hard
  refreshes. It is now built as a bundle with a new file name for every version. Run
  `bench build --app frappe_microsoft365` once after updating (`bench update` does it for you).

## 1.0.1 — 2026-10-22

### Fixed

- **Pull-only connections no longer delete in Outlook.** With *Push Frappe events to Microsoft*
  off, deleting an Event that had been pulled from Outlook still deleted the Outlook event —
  and, when the connection's owner organised the meeting, Microsoft sent every attendee a
  cancellation. Deleting is a write, so it now follows the push switch like creating and
  updating already did, both when the delete happens and again when the queued job runs.

## 1.0.0 — 2026-09-14

First stable release: two-way Outlook calendar sync including recurrence, Teams meetings from an
Event, RSVP, transcripts and recordings, Microsoft sign-in and Outlook mail OAuth, the connection
doctor and modular setup. 412 tests, CI on Frappe 15 and 16.

### Modular setup

#### Added

- **Capability selection** in Microsoft Settings: calendar, mail and sign-in are independent
  and individually optional. The doctor only checks what you selected, so it never nags about
  mail or sign-in you never wanted.
- **Set Up** builds what each selected capability needs from the single Azure registration
  already configured — a `Connected App` for mail, a `Social Login Key` for sign-in — with a
  plan shown first. It only ever creates missing records: anything that exists is reported
  with its drift and left untouched, and a second run does nothing.
- Sign-in is provisioned with tenant-specific **v2.0** endpoints, because Frappe's built-in
  Office 365 provider defaults to `/common/` and v1.0, which a single-tenant app rejects.
- New doctor rules for the `Social Login Key` (`/common/` authority, v1.0 endpoints, client-id
  mismatch, sign-in switched off) and a notice giving the exact second redirect URI Azure needs.
- **Check Microsoft Setup** button on `Email Account`, so one account can be diagnosed from
  the form where the failure appears.

Still read-only with respect to mail: no Graph mail, no replacement for `Email Account`, no
changes to the email queue.

### Connection doctor

#### Added

- **Microsoft connection doctor** (`Microsoft Settings` → Troubleshoot): inspects Microsoft
  Settings, the Connected App and every Email Account and reports what is wrong, instead of
  leaving admins with `AUTHENTICATE failed`. Rules cover scope/flow mismatches, the missing
  `offline_access` scope (the cause of "it works for an hour then stops"), v1.0 endpoints,
  tenant and redirect-URI drift, IMAP folders, TLS flag combinations, and the shared-mailbox
  identity conflict where Microsoft wants different identities for IMAP and SMTP.
- **Error explainer** mapping Microsoft/IMAP/SMTP failures (AADSTS codes, `535 5.7.3`,
  `451 4.7.0`, `invalid_grant`, `AUTHENTICATE failed`, TLS) to a cause and a next step.
- **Exchange setup script generator** for app-only mailbox access, which looks the service
  principal up by AppId rather than asking for an Object ID — Microsoft documents that using
  the App Registration one instead of the Enterprise Application one fails silently.
- 34 tests, all offline: every rule runs against plain config dicts, no tenant or network.

The doctor is read-only. It does not send or receive mail, replace `Email Account`, or touch
the email queue; Frappe's own IMAP/SMTP + OAuth path is untouched.

### Sync correctness

The calendar sync engine was reworked around Graph delta queries. If you are upgrading an
existing install, run `bench --site <site> migrate`; the first sync after the upgrade
re-initialises the delta window automatically.

#### Fixed

- **Pushed events stopped syncing after the first pull.** The pull flagged every event it
  touched as "pulled from Microsoft", including ones Frappe had just pushed, and the doc-event
  handler skips flagged events. Later local edits were silently dropped. Origin is now
  permanent.
- **The pull overwrote local descriptions** with Graph's truncated plain-text `bodyPreview`.
  Descriptions of Frappe-originated events are no longer touched.
- **A failed pull advanced the watermark**, so every change inside the failed window was lost
  forever. The watermark now moves only after a complete, successful read, and the failure is
  recorded in the new `last_error` field.
- **No pagination.** Reads were capped at the first page (`$top=50` / `$top=100`), so large
  calendars imported partially and reported success. All collection reads now follow
  `@odata.nextLink` to the end.
- **Recurring events were duplicated.** The first sync used `calendarView` (occurrences) and
  later syncs used `/me/events` (series masters) — two different id spaces. Both now use
  `calendarView/delta`, and series masters are skipped.
- **Deleted Microsoft events orphaned their Frappe mirror.** Delta `@removed` entries are now
  processed, alongside the existing `isCancelled` handling.
- **Windows timezone names silently became UTC**, moving meetings by hours. Graph is asked for
  UTC and Windows ids are mapped to IANA, with an Error Log entry for anything unrecognised.
- **Overlapping scheduled syncs.** Each calendar now syncs under a per-calendar file lock.
- **`429` was raised, not handled.** Short `Retry-After` waits are honoured with one retry;
  longer ones are deferred to the next run. `410` now triggers an automatic full re-sync.
- **Push was create-only.** Events edited while the connection was down are patched on the next
  sync instead of drifting forever.
- **`Microsoft Calendar` was readable/writable by the `All` role**, which includes Website
  Users. Replaced with `Desk User`, matching Frappe's Google Calendar.
- Pointless saves during the pull bumped `modified` on unchanged events, which made the push
  step patch them back to Graph on every run.

#### Added

- `graph_paged()` and `graph_delta()` helpers in `microsoft_graph`.
- `last_error`, `delta_window_end` and `oauth_state_expiry` fields on `Microsoft Calendar`;
  authorize links now expire after 15 minutes.
- Event `location` is synced in both directions.
- 31 tests covering the sync engine and Graph transport with mocked Graph responses.
- CI runs a Frappe `version-15` + `version-16` matrix.

#### Removed

- The unused `authorization_code` field on `Microsoft Calendar`.
