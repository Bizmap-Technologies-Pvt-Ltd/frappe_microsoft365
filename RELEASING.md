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
  in *Added*, *Changed* or *Fixed*, written for the person running the site.
- A bug or request found in testing gets a GitHub issue; the PR says `Fixes #N`, so the issue
  closes when the PR merges.

## Releasing

1. **Release commit** on the branch, titled `chore: release X.Y.Z`:
   - set `__version__` in `frappe_microsoft365/__init__.py`;
   - rename `## Unreleased` to `## X.Y.Z — YYYY-MM-DD`, and say what a site has to run after
     updating (`bench migrate`, `bench build --app frappe_microsoft365`) when it has to.
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
   # the changelog section for this version, without its heading
   awk '/^## X\.Y\.Z /{f=1;next} /^## /{f=0} f' CHANGELOG.md > notes.md
   gh release create vX.Y.Z --title "X.Y.Z" --notes-file notes.md
   ```

5. **Check** that CI is green on `main`, `version-15` and `version-16` after the push.

## Updating a site

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
