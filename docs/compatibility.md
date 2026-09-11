# Frappe v15 / v16 compatibility

This app ships **one codebase for both versions**. Nothing is version-gated, there are no
`frappe.__version__` checks in app code, and CI runs the full suite against both on every push
and pull request.

This document records *why* that is safe, so the next change does not have to re-derive it.

## Host requirements

| | Frappe version-15 | Frappe version-16 |
| --- | --- | --- |
| `requires-python` | `>=3.10,<3.15` | `>=3.14,<3.15` |
| Node | 18+ | 24+ (not used by this app — no frontend build) |
| This app's floor | `requires-python = ">=3.10"` in `pyproject.toml`, ruff `target-version = "py310"` | same |

The app has no build step and no JS bundle; the only client-side file is a doctype form script,
which both versions load the same way. `msal` and `requests` are the only declared dependencies.

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

## Version-specific code: exactly one place

`frappe_microsoft365/tests/base.py`:

```python
try:  # Frappe v16+
	from frappe.tests import IntegrationTestCase as BaseTestCase
except ImportError:  # Frappe v15
	from frappe.tests.utils import FrappeTestCase as BaseTestCase
```

v16 moved the integration test base class and deprecated the old import path; v15 only has the
old one. Every test imports `BaseTestCase` from here. `microsoft_calendar_sync._calendar_lock`
carries a second, defensive `ImportError` fallback for `filelock`, though both supported
versions ship it.

## Branches

`main` is the source of truth. `version-15` and `version-16` track it, so the usual
`bench get-app <url> --branch version-15` works for people who follow the Frappe convention.
They are the same commits, not a fork: a fix lands on `main` and is fast-forwarded to both.

If the versions ever genuinely diverge — a v16-only API worth adopting, or a v15 break — that is
the moment to let the branches differ, and this document should be updated to say where and why.

## Keeping it true

CI (`.github/workflows/ci.yml`) runs a matrix: `version-15` on Python 3.10 and `version-16` on
Python 3.14. Each job does a real `bench init`, `new-site`, `install-app`, `migrate` and the full
test suite. A change that breaks either version fails before it merges. Do not drop a matrix leg
to make a build green.
