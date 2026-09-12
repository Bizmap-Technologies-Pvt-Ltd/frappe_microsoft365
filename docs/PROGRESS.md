# Where this app stands

A short, honest statement of what is built, what is proven, and what is not. The README
describes the features; this file is for anyone deciding how far to trust them.

## Proven against a real Microsoft tenant

- **OAuth round trip.** Authorize, consent, token storage, refresh-token issuance.
- **Push.** A Frappe Event is created in Outlook and linked back.

## Built and tested, but only against mocked Graph

Everything else. The tests are thorough and offline by design, but they encode our reading of
Graph's contract, and that reading has been wrong before:

- **Delta pull.** The watermark, occurrence ids and `@removed` deletions have never completed a
  run against a live mailbox. This is the largest untested surface in the app.
- Teams meeting creation from an Event, attendees, RSVP, transcripts, recordings.
- Mail and sign-in provisioning, and every live-connection branch of the doctor.

## What live testing has already cost us

Four real bugs, none of which the mocked suite could find. Worth remembering before trusting a
green test run:

1. `after_migrate` alone left the Event custom fields missing on a plain install.
2. Graph event ids are ~152 characters; the column was `varchar(140)`, so push created the
   event in Outlook and then failed to link it — three copies of one meeting.
3. The OAuth callback threw a Python traceback at a browser instead of a readable failure.
4. The attendee summary wrote `Name <email>`, and Frappe's HTML sanitising ate the address.

## Known gaps

- **Argument type hints** are missing on whitelisted methods (Frappe validates and coerces
  arguments when they are present). Flagged by Frappe's own semgrep security rules, which CI
  does not currently gate on.
- **Two upstream Frappe bugs** remain unreported because we cannot yet demonstrate them:
  `#39099` (app-only Email Account validation) and `#23079` (IMAP and SMTP needing different
  identities for a shared mailbox).
- **No Desk workspace**; doctypes are reached through search.
- **Polling, not push.** Sync runs on a 15-minute cron. Graph change notifications would make
  it near-instant and are the obvious next infrastructure step.

## Verification a change is expected to pass

`bench --site <site> run-tests --app frappe_microsoft365`, `ruff check`, and Frappe's
`frappe_correctness` semgrep rules. CI runs all three against Frappe version-15 and version-16.
