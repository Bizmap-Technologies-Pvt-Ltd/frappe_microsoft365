# SharePoint document storage — design

Status: in progress on `feature/sharepoint`.
Based on the SharePoint work by David Eleben (`eleben/frappe_microsoft365`, branch
`feature/sharepoint-files`), brought into this app as a capability and adapted to the rules below.

## What it does

Attachments of chosen DocTypes are stored in a SharePoint document library — usually the one
behind a Microsoft Teams channel — instead of on the Frappe server. Each record gets its own
folder. Files keep opening from the record, through Frappe's permissions, and files people add
to the folder from Teams show on the record in a "Files in Microsoft 365" panel.
The user-facing behaviour is described in [sharepoint-user-guide.md](sharepoint-user-guide.md).

## Rules every capability follows

1. **Off means off.** A site that never ticks *SharePoint document storage* gets no custom fields
   on File and no behaviour change. Every hook returns before doing any work when the capability
   is off. Unticking it later stops new moves; files already in SharePoint keep opening.
2. **Explicit auth context.** Every Graph call states who it acts as: a signed-in user (a
   Microsoft Calendar's delegated token) or the app itself (client credentials). There is no
   fallback from one to the other.
3. **Least privilege.** Storage uses the `Sites.Selected` application permission, granted per
   site. Nothing asks for tenant-wide `Sites.ReadWrite.All`.

## The capability registry

`frappe_microsoft365/capabilities.py` defines every tickbox on Microsoft Settings once: its label,
whether it acts as a person (`DELEGATED`), as the app (`APPLICATION`) or through a Frappe record
it creates (`FRAPPE_RECORD`), the Azure permissions, the capabilities it builds on, the lowest
Frappe version it runs on, and an `on_enable` hook that puts what it needs on the site. The
sign-in scopes, Set Up, the doctor and the Settings form all read it, so adding a capability is
one entry there plus its own module — not edits across five files. The Event fields of the
calendar capability predate this and are still created on install.

## What is kept from the original work

Kept as written, with the same DocType names so existing data carries over:

- DocTypes: SharePoint Settings, SharePoint Mapping, SharePoint Folder, SharePoint Mapping Role.
- Upload (simple PUT up to 4 MB, upload session in 10 MB ranges above), folder per record,
  folder name pattern, base folder, Automatic / On demand folder creation.
- Opening files through `open_file` / `open_item` with Frappe permissions, the
  "stay inside the record's folder" check, and the SSRF guard for archived links.
- Rename/delete handling, the hourly retry (5 attempts), Move Existing Attachments,
  Test connection, the site grant script, and the uninstall choice
  (bring files back / leave in SharePoint).

## What changes

| Area | Original | Here |
| --- | --- | --- |
| App-only auth | `APP_ONLY` marker passed in place of a calendar | `graph.APP` auth context; `graph_request(method, path, auth)` takes a calendar (delegated) or `APP` |
| App registration | Same client id/secret as delegated sign-in | Optional separate *Application access* client id and secret on Microsoft Settings (Microsoft advises not to mix delegated and application permissions in one registration); falls back to the main registration when left empty |
| App token cache | Cleared only by Test / uninstall | Keyed to tenant + client id, and cleared whenever Microsoft Settings is saved |
| File custom fields | Created on every install and migrate | Created when SharePoint storage is ticked (and refreshed on migrate only while it is ticked); removed on uninstall |
| Hooks when off | Mostly cheap, but the old-link redirect queries a File custom field | Every hook checks the capability (or the field's existence) first |
| Frappe v15 | Storage could be enabled; server-side `File.get_content` returned nothing for moved files | Storage requires v16 (`extend_doctype_class`); ticking it on v15 is refused with an explanation |
| Opening files inline | Extensionless files fell back to SharePoint's Content-Type | Only known-safe types (PDF, images except SVG, plain text, audio, video) open inline; everything else downloads |
| Migration patches | Four patches from intermediate fork-only DocTypes | Dropped: the DocTypes are created in their final form |

## Upgrading from the fork

A site running `eleben/feature/sharepoint-files` has the same DocTypes and File fields, so
switching the app to this branch and running `bench migrate` keeps its mappings, folders and
stored files. The fork's four patches are already recorded as run there and are simply absent here.

## Later

- The Microsoft account / resource split (one auth context reaching several resources) will
  replace the calendar-as-auth-context half of rule 2; `APP` stays as is.
- OneDrive (a user's own files: attach from / save to OneDrive) is a separate, delegated feature.
