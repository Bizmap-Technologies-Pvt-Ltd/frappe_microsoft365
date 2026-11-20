# Frappe v15 / v16 compatibility

This app ships **one codebase for both versions**, and CI installs, migrates, builds and runs the
full suite against both on every push and pull request. Almost everything behaves the same on
both. The exceptions are few and deliberate, and this page lists them all: one capability is
v16-only, and the app finds out which version it is running on through a single function.

This document records *why* the rest is safe, so the next change does not have to re-derive it.

## What differs between v15 and v16

| | Frappe v15 | Frappe v16 |
| --- | --- | --- |
| Outlook calendar sync, Teams meetings on an Event, RSVP | yes | yes |
| Standalone Teams meetings | yes | yes |
| Meeting transcripts and recordings | yes | yes |
| Outlook mail over OAuth | yes | yes |
| Sign in with Microsoft | yes, with Frappe 15's own OAuth state handling | yes, the OAuth state is validated server-side |
| SharePoint document storage | **no** — Microsoft Settings refuses the tickbox | yes |
| Moved attachments readable by server-side code (`File.get_content()`) | n/a (needs SharePoint storage) | yes, through `extend_doctype_class` |
| Desk scripts (`microsoft_doctor.bundle.js`, `microsoft_files.bundle.js`) | built by `bench build` | built by `bench build` |

**SharePoint document storage needs Frappe v16.** Its capability declares `min_frappe=16` in
`frappe_microsoft365/capabilities.py`. `capabilities.frappe_major_version()` reads the major
version from `frappe.__version__`, and `Microsoft Settings` refuses to save with the tickbox on
when the site runs an older one, saying which version it needs. The reason is
`extend_doctype_class`: a moved attachment's bytes live in SharePoint, and only v16's hook lets
`File.get_content()` — which emails, prints and imports call — be answered from there. v15 has no
such hook (`hooks.py` declares it on both; v15 ignores it), so on v15 server-side readers would
get no content. See [SharePoint document storage](sharepoint-files.md#frappe-v15).

**Sign in with Microsoft uses Frappe's own social login.** This app creates the `Social Login
Key`; Frappe handles the redirect and the callback. Frappe 16 stores each sign-in's OAuth `state`
on the server and accepts it back only once. On Frappe 15, Microsoft sign-in relies on Frappe 15's
own `state` handling, which does not validate the `state` against anything stored on the server.
Connecting a Microsoft Calendar is a different flow, and the same on both versions: this app
issues its own `state`, stores it on the `Microsoft Calendar` with an expiry, and checks it in the
callback.

## Host requirements

| | Frappe version-15 | Frappe version-16 |
| --- | --- | --- |
| `requires-python` | `>=3.10,<3.15` | `>=3.14,<3.15` |
| Node | 18+ | 24+ |
| This app's floor | `requires-python = ">=3.10"` in `pyproject.toml`, ruff `target-version = "py310"` | same |

**There is a build step.** The desk-wide scripts are two bundles, `microsoft_doctor.bundle.js`
and `microsoft_files.bundle.js`, listed in `app_include_js` in `hooks.py`; `bench build --app
frappe_microsoft365` turns them into hashed files under `assets/frappe_microsoft365/dist/js/`, and
CI checks that both appear in `assets.json` on both versions. Bundles rather than plain
`/assets` paths, because a bundle gets a new file name on every build and browsers therefore pick
up a deploy. The form scripts (`event.js`, `email_account.js`, and the doctype scripts) are loaded
by `doctype_js` and the doctypes themselves, the same way on both. `msal` and `requests` are the
only declared Python dependencies.

## The calendar contract is identical in both

The `Event` doctype was compared field by field between the two branches:

- **No field added, removed, or retyped.** Every field this app writes — `subject`,
  `starts_on`, `ends_on`, `all_day`, `description`, `location`, `event_type`, `status` — has the
  same name and fieldtype in both, and `event_type` has the same options (`Private` / `Public`).
- `sync_with_google_calendar` exists in both, which is what `setup.py` anchors the Microsoft
  section to via `insert_after` (with a `has_column` fallback to `description`).
- Frappe's own **Google Calendar** integration, which this app mirrors, is functionally the same
  in both branches: same sync-token pattern, same `nextPageToken` paging, same
  `status == "cancelled"` deletion handling. The only v16 change is where the event `owner` is
  looked up, plus docstring reformatting.

That is the reason a single implementation is viable: the surface this app writes to did not
move between versions.

## Every Frappe API this app calls, verified in both

| API | v15 | v16 | Used by |
| --- | --- | --- | --- |
| `frappe.utils.get_system_timezone` | yes | yes | all datetime conversion |
| `frappe.utils.quoted` | yes | yes | OAuth error redirect |
| `frappe.utils.synchronization.filelock` | yes (same signature) | yes | per-calendar sync lock |
| `frappe.utils.file_lock.LockTimeoutError` | yes | yes | skip when already syncing |
| `frappe.log_error(title=...)` | falls back to `get_traceback` | same | every error path |
| `frappe.db.set_value(dt, dn, {dict}, update_modified=False)` | yes | yes | watermark + flags |
| `frappe.db.set_single_value` | yes | yes | tests |
| `frappe.clear_document_cache` | yes | yes | tests |
| db_query `["is", "set"]` / `["is", "not set"]` | yes | yes | push candidate filters |
| `frappe.delete_doc(force=, ignore_permissions=)` | yes | yes | removing mirrors |
| `frappe.generate_hash(length=)` | yes | yes | OAuth state |
| `frappe.scrub` | yes | yes | lock file name |
| `create_custom_fields(fields, ignore_validate=True)` | yes | yes | `after_migrate` |
| `scheduler_events` `cron` | yes | yes | 15-minute sync |
| `Desk User` role | yes (Google Calendar uses it) | yes | doctype permissions |

Runtime libraries the app relies on transitively — `filelock`, `python-dateutil`, `requests` —
are declared dependencies of Frappe itself in **both** branches, so they are always present.
`zoneinfo` is stdlib from Python 3.9, below both floors.

## Version-specific code

Every place the code knows which version it runs on:

- **`capabilities.frappe_major_version()`** and each capability's `min_frappe` — the only
  version check in app code. `capabilities.unmet()` and `Microsoft Settings` validation use it to
  keep SharePoint storage off below v16. A new v16-only feature belongs here too: give its
  capability `min_frappe=16` rather than adding a separate version check.
- **`extend_doctype_class`** in `hooks.py` (`file_extension.SharePointFile`) — a v16 hook that v15
  ignores. Only reached for files moved to SharePoint, which v15 cannot have.
- **`frappe_microsoft365/tests/base.py`**:

  ```python
  try:  # Frappe v16+
  	from frappe.tests import IntegrationTestCase as BaseTestCase
  except ImportError:  # Frappe v15
  	from frappe.tests.utils import FrappeTestCase as BaseTestCase
  ```

  v16 moved the integration test base class and deprecated the old import path; v15 only has the
  old one. Every test imports `BaseTestCase` from here.
- `microsoft_calendar_sync._calendar_lock` carries a defensive `ImportError` fallback for
  `filelock`, though both supported versions ship it.

## Branches

`main` is the source of truth. `version-15` and `version-16` track it, so the usual
`bench get-app <url> --branch version-15` works for people who follow the Frappe convention.
They are the same commits, not a fork: a fix lands on `main` and is fast-forwarded to both.

The v16-only behaviour above is gated at run time, not by branch, so the branches stay identical.
If the versions ever diverge in a way a capability's `min_frappe` cannot express — a v15 break, say
— that is the moment to let the branches differ, and this document should say where and why.

## Keeping it true

CI (`.github/workflows/ci.yml`) runs a matrix: `version-15` on Python 3.10 and `version-16` on
Python 3.14. Each job does a real `bench init`, `new-site`, `install-app`, `migrate`, `bench build
--app frappe_microsoft365` (checking both bundles were emitted) and the full test suite. A change that breaks either version fails before it merges. Do not drop a matrix leg
to make a build green.
