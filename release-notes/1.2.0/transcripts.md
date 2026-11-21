### Security
- A Teams recording can only be downloaded from an Event it is stored on, including an Event that has no recordings stored yet.

### Fixed
- A temporary error from Microsoft (503, 429, a token refresh that failed for a moment) no longer stops the catch-up job asking about a meeting for good. Only permanent answers end the schedule — no Teams meeting on the Event, a join link Microsoft matches to nothing (another organisation's meeting, or one that has aged out), or a meeting Graph answers 404 for — and every stop leaves a status on the Event saying why and whether it will be tried again. The Get Transcript & Recording button calls a 503/429 temporary instead of quoting the whole lookup. (#32)
- Each occurrence of a recurring Teams meeting now gets only its own transcripts and recordings: those created between an hour before its booked start and six hours after its booked end. A repeating Frappe Event is checked for every occurrence, not only the first, and keeps the transcripts of earlier weeks. (#33)
- The catch-up job selects only Events that are actually due, in SQL, with those still missing a transcript or recording first; it commits after every Event, so a timeout or a killed worker loses nothing already done; and it keeps checking for later recording parts of a long meeting until the schedule ends, without downloading the transcripts it already has again. (#34)
- Join links that contain `&`, `+` or `#` reach Microsoft as part of the link instead of being cut short. Links without them are sent exactly as before. When a link matches nothing, the Outlook event's own join link and then the meeting number of a `/meet/` link are tried. (#35)

### Changed
- Transcripts look up their meeting with `OnlineMeetings.Read`, the least privilege Microsoft allows, instead of `OnlineMeetings.ReadWrite`, and no longer need *Standalone Teams meetings* ticked. A site already granted `OnlineMeetings.ReadWrite` keeps requesting it; standalone meetings still use it. Run Diagnostics says which one this site uses. (#58)

### Added
- Event form, for System Managers: **Microsoft > Remove Other Meetings' Transcripts**, shown only on an Event holding transcripts dated outside its own meeting. It lists them, asks for confirmation of the count, and removes only those.

### Upgrade notes
- **No new consent prompt.** A site whose connections were authorised with `OnlineMeetings.ReadWrite` (every site that had transcripts on before 1.2.0) keeps requesting exactly the scopes it was authorised with. Re-authorising is not needed, and the doctor reports no scope change. Unticking *Standalone Teams meetings* afterwards no longer turns transcripts off, and still keeps `ReadWrite`.
- **Upgrade report:** "N of the M transcripts on Event X are from other meetings" is listed for each Event that holds transcripts dated outside its own occurrence (on the maintainer's site: the Events of one weekly series). Nothing is removed by the upgrade; use the Event form action above on each one if you want them gone.
- **Patch `reopen_artifact_fetches_stopped_by_an_error`:** Events that 1.1.x stopped checking after one error (attempts at the maximum with no status, no checked time and no meeting id) are reopened if the meeting ended in the last 7 days, so they get their transcript after the upgrade; older ones are left stopped but given a status saying to use Get Transcript & Recording. No network calls.
- The catch-up job now also picks up a meeting that ended in the last 7 days and was never checked at all (for instance on a calendar whose *Fetch transcripts after meetings* box was ticked later). Meetings that already have both a transcript and a recording are checked again on the normal schedule for a day, for later recording parts.
- A transcript already attached is no longer deleted and re-fetched when a refetch for a different occurrence runs: a refetch replaces only the files dated inside the Event's own occurrence.
