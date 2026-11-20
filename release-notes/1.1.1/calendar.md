### Security
- An Event is now only created, updated or deleted in its owner's own Microsoft calendar. Linking an Event to someone else's Microsoft Calendar is refused on save, for System Managers too.
- Microsoft access tokens are only ever sent to Microsoft Graph (`https://graph.microsoft.com`); plain HTTP and any other host are refused.
- A calendar's sync state (delta link and sync window) can no longer be changed through the API or the form; only the sync itself writes it.

### Fixed
- Calendar sync no longer gets stuck on large calendars. A first sync or a backlog of more than about 2,500 changes used to re-read the same first items on every run and never reach the rest, without reporting anything. The sync now continues from where the last run stopped until it reaches the end. If Microsoft stops returning a link that moves the sync forward, the reason is shown in **Last Sync Error**. (#18)

### Changed
- Deleting an Event in Frappe deletes the Outlook event only when the Event's owner deletes it. If anyone else deletes it (for example a Public Event, or a System Manager tidying up), the Frappe record is removed and the Outlook event stays.
- When Microsoft has more changes than one sync can read, the **Sync Now** result says that the next sync continues from there.

### Added
- Two read-only fields on Microsoft Calendar, **Delta Resume Link** and **Delta Resume Window End**. They are hidden and only the sync sets them.

### Upgrade notes
- **Events linked to someone else's calendar stop syncing.** An Event whose owner is not the user of its linked Microsoft Calendar is no longer pushed, patched or deleted in Outlook, and it cannot be saved with *Sync with Microsoft Calendar* ticked until it is linked to the owner's own calendar or sync is unticked. After `bench migrate`, the upgrade report check `upgrade.calendar_event_owner` lists these Events, with a count per calendar. Migrate does not change any of them. Events pulled from Outlook already belong to the calendar's user and are not affected.
- **Calendars that were stuck now catch up.** If a calendar has more than about 2,500 events or changes in its window, the first syncs after the upgrade may pull many events and take longer than usual. This continues over several runs (every 15 minutes) until the calendar is complete.
- **Patch `clear_non_graph_delta_links`.** It clears any stored delta link that is not an `https://graph.microsoft.com` URL. That calendar then does a fresh full sync. Normal installs have no such links, so the patch changes nothing on them. It is safe to run more than once.
- **Delete jobs queued before the upgrade.** An Outlook delete that was queued by 1.1.0 but had not run yet does nothing after the upgrade, because it cannot be checked against the Event's owner. Delete that Outlook event by hand if needed.
- Migrate adds the new fields. No manual steps are needed.
