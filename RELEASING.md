# Releasing

Every change reaches users the same way: a pull request into `main`, a version number, a
changelog entry, a tag and a GitHub release.

## Versions

`frappe_microsoft365/__init__.py` holds the version, [semver](https://semver.org):

| Change | Bump | Example |
| --- | --- | --- |
| A fix, nothing new to learn | patch | 1.0.1 → 1.0.2 |
| A new feature or option, nothing breaks | minor | 1.0.2 → 1.1.0 |
| Something existing sites must act on (removed setting, changed behaviour) | major | 1.4.0 → 2.0.0 |

## While working

- One branch per change (`fix/…`, `feature/…`, `docs/…`), one pull request into `main`.
- Each change that users would notice goes under `## Unreleased` at the top of `CHANGELOG.md`,
  in *Security*, *Fixed*, *Changed* or *Added*, written for the person running the site.
- A bug or request found in testing gets a GitHub issue; the PR says `Fixes #N`, so the issue
  closes when the PR merges.
- Every PR answers the template's question: **does this change anything for sites already
  running the app?** One of:

  | Answer | What the PR must contain |
  | --- | --- |
  | **none** | nothing more |
  | **upgrade note** | a behaviour change an admin will notice: a line under *Upgrade notes* |
  | **patch** | a data fix: a module in `frappe_microsoft365/patches/`, listed in `patches.txt` under `[post_model_sync]`; idempotent, no network calls, safe when the capability is off or the tables are empty, with a test that runs it twice; plus an *Upgrade notes* line saying what it changes |
  | **migrate warning** | data an admin has to look at: an *upgrade check* (below), plus an *Upgrade notes* line saying what the report will list |

### Upgrade checks

An upgrade check is how a release tells an admin about something on *their* site — Events that
will stop syncing, two records sharing a folder — without changing it. It is a function that
takes no arguments, only reads (SQL, no Graph calls, fast), and returns a list of
`doctor.finding(check, status, title, detail, fix, doc, target)` with `doctor.WARN` or
`doctor.FAIL`, or `[]` when there is nothing to report. Name it `upgrade.<topic>` and register it
under the release that introduces it:

```python
# hooks.py
microsoft365_upgrade_checks = {
	"1.1.1": [
		"frappe_microsoft365.<module>.<function>",
	],
}
```

After `bench migrate`, `frappe_microsoft365/upgrade.py` runs the checks of every release newer
than the one last reported on that site, up to `__version__`, once. Findings are printed in the
migrate output, shown at the top of *Run Diagnostics* as *After upgrading to X.Y.Z* until a
System Manager dismisses them, and sent to each System Manager as one notification. Checks for a
version above `__version__` do not run on migrate until the release commit bumps it; the CI
upgrade job runs every registered check against seeded data regardless, so they are exercised
before that.

### Releases built by several branches

When a release is prepared on a `release/X.Y.Z` branch by several people at once, nobody edits
`CHANGELOG.md` on their branch (it would conflict on every merge). Each branch writes its own
`release-notes/X.Y.Z/<topic>.md` instead, with the same headings as a changelog section —
*Security*, *Fixed*, *Changed*, *Added*, *Upgrade notes* — leaving out any that are empty.
At release time the maintainer folds them into one `## X.Y.Z` section: each heading once, in that
order, entries merged and edited into one voice, *Upgrade notes* ordered by what the admin does
first (manual steps before warnings to expect). Then `release-notes/X.Y.Z/` is deleted in the
release commit, so the changelog stays the single record.

## Releasing

1. **Release commit** on the branch, titled `chore: release X.Y.Z`:
   - set `__version__` in `frappe_microsoft365/__init__.py`;
   - rename `## Unreleased` to `## X.Y.Z — YYYY-MM-DD` (or fold `release-notes/X.Y.Z/` into it,
     above), and say what a site has to run after updating (`bench migrate`,
     `bench build --app frappe_microsoft365`) when it has to;
   - make sure the section ends with **Upgrade notes** (template below), even when the answer is
     "nothing changes for existing sites". Bumping `__version__` is what makes this release's
     upgrade checks run on migrate.

   ```markdown
   ## X.Y.Z — YYYY-MM-DD

   One or two sentences: what this release is about, and what to run after updating.

   ### Security
   ### Fixed
   ### Changed
   ### Added

   ### Upgrade notes

   - **Behaviour changes**: what an existing site will notice working differently.
   - **Patches that run on migrate**: what each one changes, and on which sites.
   - **Warnings to expect**: what the upgrade report (migrate output, Run Diagnostics, the
     notification) may list, and what to do about each.
   - **Manual steps**: anything an admin must do by hand, and when.
   ```

   Leave out empty headings, except *Upgrade notes*: when nothing applies it says so in one line.
2. **Merge** once CI is green on Frappe 15 and 16, with **Create a merge commit**. Not squash:
   squashing rewrites the commits, which drops co-author credit and breaks branches that
   already contain them.
3. **Version branches.** `version-15` and `version-16` follow `main` while both are supported:

   ```sh
   git fetch origin
   git rev-list --count origin/main..origin/version-15   # must be 0
   git rev-list --count origin/main..origin/version-16   # must be 0
   git push origin origin/main:refs/heads/version-15 origin/main:refs/heads/version-16
   ```

   If either has commits of its own, stop and merge instead of pushing over them.
4. **Tag and release** the merge commit:

   ```sh
   git tag -a vX.Y.Z <merge-commit> -m "X.Y.Z"
   git push origin vX.Y.Z
   # the changelog section for this version, without its heading; it includes the Upgrade
   # notes, so the GitHub release repeats them
   awk '/^## X\.Y\.Z /{f=1;next} /^## /{f=0} f' CHANGELOG.md > notes.md
   grep -q '^### Upgrade notes' notes.md || echo "notes.md has no Upgrade notes"
   gh release create vX.Y.Z --title "X.Y.Z" --notes-file notes.md
   ```

5. **Check** that CI is green on `main`, `version-15` and `version-16` after the push, including
   the *upgrade from the last release* jobs, which install the newest tag below the new version,
   seed it with data and migrate it to this one.

## Updating a site

Read the release's *Upgrade notes* first. After `bench migrate` the output ends with a
`Microsoft 365 X.Y.Z:` line: either nothing to report, or the items to check, which are also at
the top of Microsoft Settings > Troubleshoot > Run Diagnostics and in each System Manager's
notifications.

```sh
bench update --apps frappe_microsoft365
```

or by hand:

```sh
cd apps/frappe_microsoft365 && git pull && cd ../..
bench --site <site> migrate
bench build --app frappe_microsoft365
bench restart
```
