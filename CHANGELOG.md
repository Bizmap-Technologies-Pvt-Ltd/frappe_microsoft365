# Changelog

## 1.2.1 — 2026-11-25

A fix for sites whose `tabFile` is stored in MariaDB's compressed row format, where updating to
1.2.0 failed during migrate with "Row size too large". After updating, run
`bench --site <site> migrate` (`bench update` does it).

### Fixed

- **Migrate no longer fails with `Row size too large (> 8126)` on `ALTER TABLE `tabFile``.** A
  `tabFile` stored `ROW_FORMAT=COMPRESSED` (tables created by older Frappe versions can be) could
  not be rebuilt once it carried this app's indexed `custom_microsoft_local_url`, so every change
  to the File table failed, including Frappe's and other apps' own migrations. Before migrate and
  before install, `tabFile` and `tabEvent` are now converted to `ROW_FORMAT=DYNAMIC`, the format
  Frappe declares for them, when they are compressed. (#68)

### Upgrade notes

- **Behaviour changes**
  - On a site whose `tabFile` or `tabEvent` is compressed, the first migrate prints
    `Microsoft 365: converting `tabFile` from ROW_FORMAT=COMPRESSED to DYNAMIC …` and rebuilds that
    table once. The rebuild holds writes to File (or Event) while it runs; on a large table, update
    in a quiet window. Sites whose tables are already `DYNAMIC` (any site created by a current
    Frappe) see nothing.
- **Warnings to expect**
  - `upgrade.compressed_tables` (fail) if a table could not be converted, for instance because
    the rebuild timed out waiting for a lock. Migrate prints MariaDB's reason; run
    `ALTER TABLE `tabFile` ROW_FORMAT=DYNAMIC KEY_BLOCK_SIZE=0;` in `bench --site <site> mariadb`
    and migrate again (see docs/troubleshooting.md, *Row size too large*).
- **Manual steps**
  - A site stuck on a failed update to 1.2.0 can update straight to 1.2.1 and migrate. If
    `tabFile` was already converted by hand, nothing more is needed.

## 1.2.0 — 2026-11-21

Every bug found in the November audit, fixed and verified against the audit's reproductions:
two-way calendar sync that keeps Outlook's edits, all-day and recurring events, connections that
recover or say why they can't, Teams transcripts per meeting occurrence, SharePoint storage
hardening, and sign-in that only admits your own directory on Frappe 15 and 16. **If Sign in with
Microsoft is enabled, read the first manual step before updating.** After updating, run
`bench --site <site> migrate` and `bench build --app frappe_microsoft365` (`bench update` does
both), and read the `Microsoft 365 1.2.0:` line at the end of the migrate output.

### Security

- **Sign in with Microsoft on Frappe 15** accepts a sign-in only with a state this site's login
  page issued in the last ten minutes, and only once, as Frappe 16 already does.
- A Teams recording can only be downloaded from an Event it is stored on, including one with no
  recordings stored yet.
- SharePoint storage: archived links are downloaded from the address that was checked (no
  second name lookup, every redirect checked the same way).
- SharePoint storage: a moved file's kept local copy follows its File's privacy and is removed
  with the File when no other File uses it.

### Fixed

**Calendar**
- Changes made in Outlook are no longer overwritten: a push sends only the fields changed in
  Frappe, and attendees are merged into Outlook's list instead of replacing it. A reply pulled
  from Outlook no longer causes a push. (#19)
- All-day events keep their dates in both directions; they were stored hours off when pulled and
  refused by Microsoft when pushed. (#20)
- The sync always looks the full 180 days ahead; its window used to shrink to 14 days. (#21)
- Events deleted in Outlook while the sync reset no longer stay in Frappe unnoticed: after a full
  read, events Outlook no longer has are marked *Missing From Outlook* and listed by Run
  Diagnostics. (#22)
- Repeating Events pushed from Frappe: an occurrence moved or renamed in Outlook appears in
  Frappe, a series deleted in Outlook is marked *Missing From Outlook*, and the recurrence is only
  re-sent when it changes. (#23)
- A connection whose sign-in Microsoft no longer accepts (`invalid_grant`,
  `interaction_required`) is marked *not authorised* with Microsoft's reason, its user gets one
  notification, and it is no longer retried every 15 minutes. Timeouts and server errors never
  mark a connection. (#24)
- The pull no longer fails with `TimestampMismatchError: ... (Microsoft Calendar) has been
  modified after you have opened it`: tokens are stored without saving the connection, concurrent
  refreshes are serialised, and a long sync commits as it goes instead of holding row locks. (#25)
- Throttling is handled as Microsoft asks: 429 and 503 are retried per `Retry-After` (seconds or
  an HTTP date), 504 for requests safe to repeat, and a mailbox that stays throttled is left alone
  for the rest of the run. (#26)
- One Frappe Event creates exactly one Outlook event: the scheduled push and a person's save can
  no longer both create it, and a save that rolls back leaves no meeting behind. (#27)
- Moving an Event to another Microsoft Calendar creates it there and deletes it from the old one;
  unticking *Sync with Microsoft Calendar* sticks. (#28)
- Re-authorising a connection with a different Microsoft account starts a clean sync. (#29)
- Authorised scopes are tracked per connection, and the doctor checks each one. (#30)

**Teams transcripts and recordings**
- A temporary error no longer stops the catch-up job for good; only permanent answers end the
  schedule, and every stop leaves a status saying why. (#32)
- Each occurrence of a recurring meeting gets only its own transcripts and recordings, and a
  repeating Frappe Event is checked for every occurrence. (#33)
- The catch-up job selects only Events that are due, commits after each, and keeps checking for
  later recording parts until the schedule ends. (#34)
- Join links containing `&`, `+` or `#` reach Microsoft whole; other links are sent exactly as
  before. When a link matches nothing, the Outlook event's own link is tried. (#35)

**SharePoint storage**
- Public attachments moved to SharePoint open for visitors again, including ones moved before
  1.2.0. (#37)
- Old `/files` links redirect with *Keep Local Copy* on, and to a row the user may read. (#39)
- A library or site recreated under the same address is looked up again instead of breaking every
  new folder. (#40)
- Paths are checked against SharePoint's 400-character limit up front; record folders and file
  names leave room, and files more than 10 folders deep open again. (#41)
- A worker that died after uploading no longer leaves a duplicate: the retry adopts the upload.
  (#42)
- A moved File re-attached to another record takes its SharePoint file with it, in a background
  job after the save commits; meanwhile its status is *Moving*. (#43)
- `File.get_content()` streams and refuses a file over the limit instead of reading any size into
  memory. (#47)
- Large uploads resume instead of starting again from byte 0. (#60)

**Sign-in and setup**
- Sign in with Microsoft on Frappe 15 refuses accounts from another organisation's directory, as
  Frappe 16 already did. (#64)
- Rotating the client secret no longer breaks mail or sign-in unnoticed: the doctor and Set Up
  report a stale copy, compared on the server and never shown. (#44)
- Upgrading from before 1.0.0 no longer switches Teams meetings and transcripts off. (#46)

### Changed

- **Sign in with Microsoft is refused while Tenant ID is `common`, `organizations`, `consumers`
  or empty**, as announced in 1.1.1. A refused sign-in shows a page saying why. (#65)
- The Microsoft sign-in callback is routed through this app on Frappe 15 and 16; the Social Login
  Key's redirect URL does not change.
- Transcripts look up their meeting with `OnlineMeetings.Read` on new setups and no longer need
  *Standalone Teams meetings*; a site already granted `OnlineMeetings.ReadWrite` keeps requesting
  it. Run Diagnostics says which one is in use. (#58)
- A new Event without a Teams meeting is created in Outlook just after its save commits; its
  Outlook link appears after a reload. A new Teams meeting is still created in the save, so its
  join link shows at once.
- A record's SharePoint folder is never renamed (it never actually was); the docs no longer say
  it is. (#38)
- On Frappe 15 the doctor warns while Microsoft sign-in is enabled, since Frappe 15 does not
  protect other social login providers against login CSRF.

### Added

- **Delete events missing from Outlook** in Microsoft Settings (off by default), and the *Missing
  From Outlook* marker on Events.
- **Remove Other Meetings' Transcripts** on the Event form (System Managers), for Events holding
  transcripts dated outside their own meeting.
- **Sync Secret** in the Set Up dialog (System Managers).
- Server-managed fields: *Microsoft Account ID*, *Authorization Lost* and *Authorized Scopes* on
  Microsoft Calendar; SharePoint Status *Moving* on File.

### Upgrade notes

- **Manual steps**
  - **If Sign in with Microsoft is enabled and Tenant ID is still `common`, `organizations`,
    `consumers` or empty, set Tenant ID before updating.** From 1.2.0 every Microsoft sign-in on
    such a site is refused (password sign-in is not affected). Put the *Directory (tenant) ID* from
    the app registration's Overview page into Microsoft Settings > Tenant ID, and point the Social
    Login Key's Authorize URL and Access Token URL at
    `https://login.microsoftonline.com/<tenant id>/oauth2/v2.0/...`. The upgrade report fails with
    `upgrade.sign_in_refused` on any site still in this state. Sites without Microsoft sign-in
    see nothing.
  - If the report lists `upgrade.meeting_capability_off` (sites that upgraded from before 1.0.0):
    tick *Standalone Teams meetings* and *Meeting transcripts and recordings* again, save, and have
    each user authorise their Microsoft Calendar again.
  - Coming from 1.1.0 or 1.1.1: if the report lists `upgrade.calendar_event_owner` for events made
    for the calendar's own user, use *Microsoft > Give Linked Events to <user>* on that Microsoft
    Calendar (added in 1.1.2).
- **Behaviour changes**
  - **First sync after updating re-reads each calendar once** (30 days back to 210 ahead, about a
    first sync's work); after that the window renews about monthly.
  - Connections whose refresh token Microsoft already rejects are marked and their users notified
    the first time anything refreshes them; live connections are not affected.
  - *Missing From Outlook* appears after that first sync, not at migrate. Nothing is deleted
    unless *Delete events missing from Outlook* is ticked.
  - A token refresh is now a commit point, and a sync may sit out up to 60 seconds of Microsoft's
    requested wait in a background job.
  - Moving an Event between calendars cancels it in the old mailbox and invites again from the
    new one.
  - Frappe 15: Microsoft sign-in needs the state Frappe's own login page issues. Anyone with the
    login page open during the update signs in again once. A site that replaced the login page
    must build the Microsoft button through the login page's `provider_logins` context (see
    docs/compatibility.md).
  - Server-side readers of a moved SharePoint file read at most Frappe's upload limit (System
    Settings > Max File Size); set `microsoft365_max_content_bytes` in `site_config.json` for
    another limit. Opening a file from the form is unaffected.
  - Link archiving no longer goes through an HTTP(S) proxy set in the worker's environment.
  - Re-attaching a moved file with *Remove from SharePoint* off (the default) never sends
    anything to the recycle bin: the old file stays and a copy is made where it needs to go.
- **Patches that run on migrate** (none calls Microsoft; each is safe to run again)
  - `copy_authorized_scopes_to_calendars`: copies the site-wide authorised scopes onto every
    authorised Microsoft Calendar, so nobody is asked to re-authorise because of the update.
  - `correct_pulled_all_day_events`: puts pulled all-day events stored with shifted times back on
    their own date, only where the stored times are exactly a converted UTC midnight. All-day
    events created in Frappe are never touched.
  - `reopen_artifact_fetches_stopped_by_an_error`: meetings that 1.1.x stopped checking after one
    error are reopened if they ended in the last 7 days; older ones get a status saying to use
    *Get Transcript & Recording*.
  - `move_public_kept_copies_of_private_files`: moves kept local copies of private Files out of
    `public/files`. Nothing is deleted or fetched.
  - `clear_legacy_default_scopes` (changed): on a site that has not run it yet, it now keeps Teams
    meetings and transcripts on.
- **Warnings to expect** in the upgrade report
  - `upgrade.sign_in_refused` (FAIL): see the first manual step.
  - `upgrade.transcripts_outside_occurrence`: Events holding transcripts from other occurrences of
    the same meeting. Nothing is removed; use *Remove Other Meetings' Transcripts* on each.
  - `upgrade.meeting_capability_off`: Teams or transcript data exists while the matching
    capability is off.
  - `upgrade.calendar_orphaned_events`: normally empty at migrate; see *Missing From Outlook*.
  - `upgrade.sharepoint_orphaned_kept_copies`: kept local copies of Files deleted before 1.2.0,
    for you to delete from the site's files folders.
  - `upgrade.sharepoint_long_paths`: mappings whose base folder already uses more than 300 of
    SharePoint's 400 path characters; such a mapping cannot be saved again until shortened.
  - The doctor may report `secret.drift` on sites that rotated their client secret in Microsoft
    Settings only; use *Set Up > Sync Secret*. Frappe 15 sites with Microsoft sign-in see
    `sso.frappe15_state` (WARN).

## 1.1.2 — 2026-11-21

A fix for a regression in 1.1.1 that stopped Teams transcripts from being found for new
meetings, and two SharePoint storage protections from 1.1.1 completed for files moved before it.
Every site on 1.1.1 should update; sites on 1.1.0 or earlier should go straight to 1.1.2. After
updating, run `bench --site <site> migrate` and `bench build --app frappe_microsoft365`
(`bench update` does both), and read the `Microsoft 365 1.1.2:` line at the end of the migrate
output.

### Security

- **SharePoint storage:** an item stored in another record's folder is opened only by someone
  who may read that record, and is never deleted through a File attached elsewhere. 1.1.1 stopped
  File rows from being edited to point at another record's file, but a row edited that way
  before the upgrade could still open it.
- **SharePoint storage:** a file moved before 1.1.1 (which has no content hash) no longer falls
  back to a local path that another File now keeps as its own local copy.

### Fixed

- Teams transcripts and recordings are found again for meetings that have no meeting id stored
  yet. 1.1.1's check that tokens only go to Microsoft Graph refused the join-link lookup, whose
  query contains spaces, so it never reached Microsoft; the URL is now encoded exactly as 1.1.0
  sent it.
- The *Files in Microsoft 365* panel no longer forgets a record's folder after a single "not
  found" from SharePoint's listing: it asks for the folder itself first. Forgetting it made the
  next attachment create a second folder (`Name 1`) beside the first. (Part of #40.)

### Added

- **Give Linked Events to <user>** on a Microsoft Calendar (Microsoft menu, System Managers):
  makes the calendar's user the owner of the Events linked to it that someone else owns, after
  saying how many, with a comment on each naming its previous owner. Since 1.1.1 such Events —
  typically created by an administrator for a colleague — are not synced, and Frappe offers no
  way to change an Event's owner from the form.

### Upgrade notes

- **Behaviour changes**
  - An attachment copied from one record to another (for example when a document is made from
    another) still opens, but only for users who can also read the record whose SharePoint
    folder holds it.
- **Patches that run on migrate:** none.
- **Warnings to expect** in the upgrade report
  - `upgrade.sharepoint_files_in_other_folders`: Files attached to one record but stored in
    another record's folder — usually attachment copies, otherwise Files edited before 1.1.1 to
    point at someone else's file. Delete the ones that should not exist; attach the file to the
    record again where a copy must be readable on its own.
- **Manual steps**
  - If 1.1.1's upgrade report listed `upgrade.calendar_event_owner` (Events on a calendar that
    belong to someone else) and they were made for that calendar's user, open the Microsoft
    Calendar and choose *Microsoft > Give Linked Events to <user>*.
  - Meetings that ended while a site ran 1.1.1 may have stopped looking for their transcript.
    Press *Get Transcript & Recording* on those Events.

## 1.1.1 — 2026-11-21

Security fixes across calendar sync, SharePoint storage, sign-in and the generated setup
scripts, plus fixes found in an audit of 1.1.0 and the first release with an upgrade report.
Every site running 1.1.0 or earlier should update. After updating, run
`bench --site <site> migrate` and `bench build --app frappe_microsoft365` (`bench update` does
both), then read the `Microsoft 365 1.1.1:` line at the end of the migrate output.

### Security

- **Calendar:** an Event is only created, updated or deleted in its owner's own Microsoft
  calendar. Linking an Event to someone else's Microsoft Calendar is refused on save, for System
  Managers too.
- **Calendar:** Microsoft access tokens are only ever sent to Microsoft Graph
  (`https://graph.microsoft.com`); plain HTTP and any other host are refused. A calendar's sync
  state (delta link and sync window) can no longer be changed from the form or the API.
- **Sign in with Microsoft** only accepts accounts from the directory named in Microsoft
  Settings. Set Up refuses to create the sign-in record while Tenant ID is `common`,
  `organizations` or `consumers`, the doctor reports such a record as a failure, and on Frappe 16
  a sign-in from any other directory is refused.
- **SharePoint storage:** only the app decides where a File is stored. The Microsoft 365 fields on
  File and the `open_file` address are refused from the form and the API, and opening, reading
  (`File.get_content`) or deleting a moved file requires it to be stored by the app inside a
  known record folder.
- **SharePoint storage:** every record gets a folder of its own; an existing folder with the same
  name is never reused, and a folder that already belongs to one record cannot be linked to
  another.
- **SharePoint storage:** when a moved file's SharePoint copy is gone, a local copy is served (or
  restored on uninstall) only if it is still that file's own bytes.
- **Setup scripts:** values written into the generated Site Grant and Exchange PowerShell scripts
  are always literal text. They are validated (client ids as GUIDs, mailboxes as addresses,
  SharePoint sites on `*.sharepoint.com`) when saved and again when a script is generated, and
  written as single-quoted strings.

### Fixed

- Calendar sync no longer gets stuck on large calendars. A first sync or a backlog of more than
  about 2,500 changes used to re-read the same first items on every run and never reach the
  rest, without reporting anything. The sync now continues from where the last run stopped; if
  Microsoft stops returning a link that moves it forward, the reason is shown in *Last Sync
  Error*. (#18)
- Transcripts on tenants with speaker attribution switched off no longer fail with
  `403 SpeakerAttributionNotAllowed`: the app asks again for Microsoft's plain-text format and
  attaches it as a `.txt` without speaker names. (#31)
- The Site Grant Script grants the *Application Access* registration when one is set — the one
  SharePoint calls are actually made as — instead of the main one, which left every upload at
  403. (#36)
- The Exchange Setup Script runs as pasted: it connects to Microsoft Graph
  (`Connect-MgGraph -Scopes Application.Read.All`) before looking up the service principal, no
  longer suggests the retired AzureAD module, finds the app by AppId rather than by a display
  name every site in the tenant shares, and can be run twice. (#45)
- `docs/setup.md`: the transcript setup PowerShell used `true` for two Boolean parameters and
  failed as pasted; it uses `$true` and says to update the MicrosoftTeams module first. (#49)
- `docs/compatibility.md`, the FAQ and the README now state what differs between Frappe 15 and
  16 — SharePoint storage needs 16 — with a feature table. (#50)

### Changed

- Deleting an Event in Frappe deletes the Outlook event only when the Event's owner deletes it.
  When anyone else deletes it (a Public Event, a System Manager tidying up), the Frappe record is
  removed and the Outlook event stays.
- New SharePoint record folders are created with SharePoint's rename rule: when
  `Projects/PRJ-0001` already exists, the new record's folder is `Projects/PRJ-0001 1`. Base
  folders are shared as before. A record's folder is recorded as soon as it is created, and
  parallel uploads for one record wait for each other, so no second folder is left behind.
- The folder of a SharePoint Mapping *Folders* row can no longer be edited: delete the row and add
  a new one (with *Existing Folder*) to point a record elsewhere.
- Microsoft Settings refuses a Client ID or Application Client ID that is not a GUID, and
  SharePoint Mapping refuses a site address outside `<tenant>.sharepoint.com` or a site name
  with characters a SharePoint address cannot have.
- The doctor checks the Application Access registration (a client id with no secret, a secret
  that is the Secret ID, a client id that is not a GUID), recognises the speaker-attribution
  refusal from Microsoft's message text, and Graph errors keep Microsoft's nested error code.
- `get_transcript_content` returns the format Microsoft actually sent in `format`, and
  `get_transcripts_for_join_url` adds `latest_format` next to `latest_vtt`.

### Added

- **Upgrade report.** After `bench migrate` the app checks the site, once per release, for
  anything the new release changes or that an admin should look at. Findings are printed at the
  end of the migrate output, shown at the top of Microsoft Settings > Troubleshoot > *Run
  Diagnostics* as *After upgrading to X.Y.Z* until a System Manager dismisses them, and sent to
  each System Manager as one notification. A site with nothing to report gets no notification,
  and a new install reports nothing. (#53)
- CI upgrades a site running the previous release, seeded with calendars, Events and (on Frappe
  16) SharePoint data, to every change before it is merged, on Frappe 15 and 16. (#54)
- Every release ends with *Upgrade notes*, and pull requests say whether they change anything for
  existing sites. (#55)
- Hidden, server-managed fields: *Delta Resume Link* and *Delta Resume Window End* on Microsoft
  Calendar; *Folder ID* and *Content Hash Before Move* on File (storage-enabled sites only); the
  upgrade report's state on Microsoft Settings.

### Upgrade notes

- **Manual steps**
  - **Sign-in on a multi-tenant Tenant ID.** If the upgrade report or the doctor shows an
    `sso.tenant_*` or `upgrade.sign_in_tenant` failure: set Tenant ID in Microsoft Settings to
    your Directory (tenant) ID, make the app registration single-tenant, and point the Social
    Login Key's URLs at `https://login.microsoftonline.com/<tenant id>/oauth2/v2.0/...`.
    Sign-in keeps working in 1.1.1 until you do; **from 1.2.0, Microsoft sign-in on a
    multi-tenant Tenant ID is refused.** Partner users who need access should be invited into
    your directory as guests.
  - **Separate Application Access registration for SharePoint:** run the corrected Site Grant
    Script once (SharePoint Settings or each mapping).
  - **Delete jobs queued before the upgrade:** an Outlook delete queued by 1.1.0 that had not run
    yet does nothing after the upgrade, because it cannot be checked against the Event's owner.
    Delete that Outlook event by hand if needed.
- **Behaviour changes**
  - Events linked to someone else's Microsoft Calendar stop syncing, and cannot be saved with
    *Sync with Microsoft Calendar* ticked until they are linked to the owner's own calendar or
    sync is unticked. Events pulled from Outlook already belong to the calendar's user and are not
    affected.
  - On Frappe 16, accounts from directories other than the configured one can no longer sign in.
    Frappe 15 has no sign-in hook, so there the protection is the tenant-specific URLs on the
    Social Login Key: make sure the doctor shows no `sso.tenant_*` failure.
  - Calendars that were stuck catch up: on calendars with more than about 2,500 events or changes
    in their window, the first syncs after the upgrade may pull many events and take longer,
    over several 15-minute runs.
  - On tenants that keep speaker attribution off, transcripts arrive as
    `teams-transcript-<date>.txt` (no speaker names). Meetings from the last day are picked up
    by the catch-up job; for older ones press *Get Transcript & Recording* on the Event. For
    names, turn on *Include speaker attribution* in the Teams admin center, delete the `.txt` and
    fetch again.
  - Files moved to SharePoint before 1.1.1 have no recorded folder or content hash: opening or
    deleting one costs one extra Graph request to confirm where it is; one that SharePoint places
    outside every record folder no longer opens and is not deleted with its attachment (the event
    is logged); and if its SharePoint copy disappears, the local fallback is refused whenever
    another File now owns the old disk path.
  - Code that set the Microsoft 365 File fields with `doc.save()`/`insert()` now fails; the app
    writes them with `frappe.db.set_value`. Attachment copies still work when they repeat the
    original's fields unchanged. `microsoft_files.ensure_folder` now commits when it creates a
    folder.
- **Patches that run on migrate**
  - `clear_non_graph_delta_links` clears any stored delta link that is not an
    `https://graph.microsoft.com` URL; that calendar then does a fresh full sync. Normal installs
    have none, so nothing changes.
  - `fill_sign_in_tenant_id`, on sites whose Tenant ID is `common`/`organizations` and that have a
    Microsoft Social Login Key: fills Tenant ID from the authorised Microsoft Calendars when every
    stored token names the same work or school directory, and points the key's multi-tenant URLs
    at it when the key uses the Client ID from Microsoft Settings. Nothing is sent to Microsoft;
    when it cannot decide, it changes nothing.
  - Migrate adds the new fields listed under *Added*; no other data changes.
- **Warnings to expect** in the upgrade report
  - `upgrade.calendar_event_owner`: Events whose owner is not the user of their linked Microsoft
    Calendar, with a count per calendar.
  - `upgrade.sign_in_tenant`: sign-in still on a multi-tenant Tenant ID (see *Manual steps*).
  - `upgrade.script_values`: saved Client IDs or SharePoint Mapping site addresses that would now
    be refused, so they can be corrected before the form is next saved.
  - `upgrade.sharepoint_file_items`: File rows whose SharePoint item does not look like the app's
    own work. For files moved before 1.1.1 it compares recorded addresses, so a folder renamed
    since the move also appears here.
  - `upgrade.sharepoint_shared_folders`: records that share one SharePoint folder. They are not
    split automatically: delete the *Folders* rows of the records that should not have it, and
    their next attachment gets a new folder.

## 1.1.0 — 2026-10-24

SharePoint document storage, built by David Eleben (@eleben) in his fork
(`eleben/frappe_microsoft365`, branch `feature/sharepoint-files`) and brought into this app as a
capability, with the changes listed under *Changed from the original*. Tested on real tenants
by @eleben and @kgrininger.

After updating, run `bench --site <site> migrate` and `bench build --app frappe_microsoft365`
(`bench update` does both). Document storage stays off until *SharePoint document storage* is
ticked in Microsoft Settings; the [user guide](docs/sharepoint-user-guide.md) has the setup.

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
- The File list shows each file's **SharePoint Status** (Pending, Stored, Archived, Failed), so
  what has moved is visible without opening each file. (Suggested by @kgrininger)

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
  Addresses now have room (the SharePoint address is a text field, so a wide column can't take
  a busy `tabFile` past MariaDB's row-size limit on migrate), every attempt counts before any
  work, and a retry resumes from the upload it already made. Copies already in SharePoint are not deleted:
  remove the extras by hand.
- **The Files panel's Modified column showed raw HTML**, and its times were off by the site's
  UTC offset (a file saved two minutes ago showed as hours old on an India site).
- **Script changes did not reach browsers after a deploy.** The panel script is now a bundle with
  a new file name for every version, like the doctor script in 1.0.2.
- **A changed Files panel setting reached nobody until they signed in again.** Saving a mapping,
  or switching storage on or off, now takes effect on everyone's next page load, and at once on
  the desk that saved it. A deleted mapping also no longer goes on storing files. (@eleben)
- **The Site Grant Script left PowerShell waiting at `>>` when pasted.** Each grant is now one
  line and the script ends with a line break, so it runs as pasted; the mapping form's dialog has
  a Copy button too. (Found by @kgrininger)
- **The panel said "Empty" right after *Create folder now*** where realtime is blocked. It now
  says the attachments are moving and lists them as they arrive. (@eleben)

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
