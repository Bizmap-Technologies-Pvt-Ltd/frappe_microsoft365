### Fixed

- Transcripts: tenants with speaker attribution switched off no longer fail every transcript
  fetch with `403 SpeakerAttributionNotAllowed`. The app now retries with Microsoft's
  unattributed plain-text format and attaches it as a `.txt` without speaker names (#31).

- Docs: the transcript setup PowerShell in `docs/setup.md` used `true` for two Boolean
  parameters and failed as pasted; it now uses `$true` and says to update the MicrosoftTeams
  module first (#49).

### Changed

- `get_transcript_content` returns the format Microsoft actually sent in `format`, and
  `get_transcripts_for_join_url` adds `latest_format` next to `latest_vtt` (#31).
- Run Diagnostics recognises the speaker-attribution refusal from Microsoft's message text as
  well as its code (#31).

### Added

- A test that scans every PowerShell block in `docs/` and `README.md` for a Boolean parameter given
  a bare `true`/`false` (#49).

### Upgrade notes

- **Behaviour change (#31):** on tenants that keep speaker attribution off, transcripts now
  arrive as `teams-transcript-<date>.txt` (timestamped lines, no speaker names) instead of
  failing. Nothing to do on migrate. Meetings that ended within the last day are picked up by
  the catch-up job on its next step; for older ones press **Get Transcript & Recording** on the
  Event. To get `.vtt` files with names, turn on
  *Include speaker attribution* in the Teams admin center, then delete the `.txt` and fetch again.
