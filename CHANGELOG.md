# Changelog

## Unreleased — sync correctness

The calendar sync engine was reworked around Graph delta queries. If you are upgrading an
existing install, run `bench --site <site> migrate`; the first sync after the upgrade
re-initialises the delta window automatically.

### Fixed

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

### Added

- `graph_paged()` and `graph_delta()` helpers in `microsoft_graph`.
- `last_error`, `delta_window_end` and `oauth_state_expiry` fields on `Microsoft Calendar`;
  authorize links now expire after 15 minutes.
- Event `location` is synced in both directions.
- 31 tests covering the sync engine and Graph transport with mocked Graph responses.
- CI runs a Frappe `version-15` + `version-16` matrix.

### Removed

- The unused `authorization_code` field on `Microsoft Calendar`.
