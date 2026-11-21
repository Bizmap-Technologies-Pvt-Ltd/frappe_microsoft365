### Fixed

- Calendar: changes made in Outlook are no longer overwritten by the next sync. A push now sends only the fields that changed in Frappe, and attendees are merged into Outlook's own list instead of replacing it, so people invited from Outlook stay invited and an agenda rewritten there survives a title change in Frappe. A reply pulled from Outlook no longer causes a push at all. (#19)
- Calendar: all-day events keep their dates in both directions. They were stored hours off when pulled (05:30 on a Kolkata site, 19:00 the day before on a New York one) and were rejected by Microsoft when pushed. (#20)
- Calendar: the sync always looks the full 180 days ahead. Its window used to shrink to 14 days before being renewed, so meetings booked months ahead were not pulled. (#21)
- Calendar: events deleted in Outlook while the sync was resetting no longer stay in Frappe unnoticed. A renewal reads the old link's deletions first, and after any full read the events Outlook no longer has are marked Missing From Outlook and listed by Run Diagnostics. (#22)
- Calendar: for a repeating Event pushed from Frappe, an occurrence moved or renamed in Outlook now appears in Frappe, deleting the whole series in Outlook marks the Frappe Event Missing From Outlook (listed by Run Diagnostics; deleted only with Delete events missing from Outlook ticked), and the recurrence is only re-sent when it changes in Frappe. (#23)
- Calendar: a long sync no longer holds row locks: the pull is committed before the push starts, and the push commits after each event. (#25, sync part)
- Calendar: one Frappe Event creates exactly one Outlook event. The scheduled push and a person's save can no longer both create it, a form opened before the create can no longer unlink it, and a save that rolls back leaves no meeting (with invitations sent) behind in Outlook. (#27)
- Calendar: moving an Event to another Microsoft Calendar creates it there and deletes it from the old one, instead of failing with 404 forever; unticking Sync with Microsoft Calendar now sticks. (#28)

### Changed

- Calendar: a new Event without a Teams meeting is created in Outlook just after its save commits (queued like an edit) rather than inside the save. Its Outlook link appears on the form after a reload. A new Teams meeting is still created inside the save so its join link is shown at once.
- Calendar: an Event with Sync with Microsoft Calendar unticked is left alone by the pull, including when it is deleted in Outlook.
- Calendar: changing an Event's Microsoft Calendar deletes it from the old mailbox and creates it in the new one; its attendees receive a cancellation and a new invitation.

### Added

- Microsoft Settings: **Delete events missing from Outlook** (off by default). When ticked, events Outlook no longer has — pulled events a full read of the calendar no longer finds, and repeating events pushed from Frappe whose Outlook series is gone — are deleted instead of only being marked. Events that other records link to are kept and stay marked.
- Event: **Missing From Outlook** marker, and Run Diagnostics lists such events per calendar.

### Upgrade notes

- **Migrate corrects all-day events already pulled.** A patch puts every all-day event that was pulled from Outlook with shifted times back on its own date, in place and without calling Microsoft. Only events whose stored times are exactly a UTC midnight converted into the site timezone are changed, so it is exact and safe to run again; all-day events created in Frappe are never touched. A site with no pulled all-day events (the patch makes one query) is unaffected.
- **First sync after upgrading re-reads each calendar once.** Windows from earlier releases are sized to the look-ahead alone, so every calendar renews its window on its first run: one full read of 30 days back to 210 days ahead (about the same work as a first sync). After that a renewal happens about once a month. Any pulled all-day event the patch could not correct (the site timezone was changed since it was pulled) is corrected by this read if it is in the window.
- **Missing From Outlook only appears after that first sync.** The upgrade report run by `bench migrate` cannot know which events Outlook has deleted, so it normally reports nothing for this; after the first sync, events Outlook no longer has are marked, the sync result says how many, and Microsoft Settings > Troubleshoot > Run Diagnostics lists them. Nothing is deleted unless an admin ticks Delete events missing from Outlook.
- **Edits are sent field by field from now on.** Existing Events get the record this needs lazily: on their next pull, or from the version before their next edit. An Event that had unsent Frappe edits at upgrade time is sent in full once more, as before, so those edits are not lost.
- **Recurring series** pushed from Frappe are recognised by their Outlook series on the first sync that sees one of their occurrences (the window re-read above does this for every series in the window); until then a series deleted in Outlook stays in Frappe unmarked. A series whose Outlook master answers "not found" is only marked Missing From Outlook, never deleted unless the setting is ticked, so re-authorising a connection as a different Microsoft account does not wipe the old account's series.
- **Moving an Event between calendars** now cancels it in the old mailbox and invites again from the new one. **Unticked Events** are no longer re-ticked by the pull.
- New Event custom fields are added by migrate. Until migrate has run, the sync keeps working and simply skips the features that need them.
