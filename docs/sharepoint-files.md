# SharePoint document storage

*For a step-by-step guide aimed at the people who use and set it up, see the
[user guide](sharepoint-user-guide.md).*

Keep a DocType's attachments in SharePoint, or in a Microsoft Teams channel's Files tab, instead
of on the Frappe server. Each record gets its own folder. Files attached in Frappe move there,
and files people drop into the folder from Teams show on the record.

It is meant for sites where the documents attached to a record matter more than the record:
contracts and their amendments, project deliverables, supplier certificates, signed agreements. These files
are large, they keep coming, and the people who work on them already live in Teams. It also
frees disk space, which matters on hosted plans with a fixed quota (Frappe Cloud, for example).

## What happens

1. A **SharePoint Mapping**, one per DocType, points that DocType at a SharePoint site, a
   document library and a base folder. For a Teams channel, the library is `Documents` and the base
   folder is the channel's name.
2. A record gets a folder under the base folder, named from a pattern (`{name}` by default, so
   `PRJ/2026/001` becomes `Projects/PRJ-2026-001`). When depends on the mapping's **Folder
   Creation**:
   - **Automatic**: on the record's first attachment. Every record of the DocType ends up in
     SharePoint.
   - **On demand**: only when someone asks, with **Create folder now** on the record's panel or by
     adding a row to the mapping's **Folders** table. Records nobody asked about keep their files on the
     server. Creating the folder also sends that record's existing attachments into it.

   Who may create folders is set per mapping in **Who can create folders**: empty means System
   Managers only; listed roles may too. Either way the user must be able to edit the record.

   Each mapping's **Folders** table has one row per record with a folder; the app adds rows as
   records get theirs. To add one by hand, **Add Row**, pick the record and **Save**: the folder
   is created where the mapping says, or, with **Existing Folder** filled in, **a folder that
   already exists is linked**: paste its SharePoint address (from the browser or *Copy link*) or
   its path in the library, e.g. `Projects/PRJ-0001`. Deleting a row removes only the link; the
   folder and its files stay in SharePoint.
3. When someone attaches a file to the record, a background job uploads it to that folder.
   The `File` row is then pointed at `/api/method/frappe_microsoft365.microsoft_files.open_file`.
   The attachment still opens from the form, previews, prints and emails as before, and the copy
   on the server's disk is removed. Tick **Keep a local copy after upload** while trialling (on
   SharePoint Settings, or for one DocType on its mapping).
4. The record shows a **Files in Microsoft 365** panel. It lists everything in the folder,
   including subfolders and files added from Teams or SharePoint, and has an **Open in
   SharePoint** button.
5. Files open through Frappe, so anyone who can read the record can open its files, whether or
   not they have a SharePoint licence. People with SharePoint access can also use the folder
   directly.

Nothing is ever deleted from SharePoint by default. Removing an attachment in Frappe leaves the
SharePoint copy alone. **Remove from SharePoint when the attachment is deleted** changes that, and
even then the file goes to the site's recycle bin. Deleting a record keeps its folder too.

### Options, site-wide and per DocType

**SharePoint Settings** holds the defaults for every mapped DocType. Two of them can be set
differently for one DocType, under **Options** on its SharePoint Mapping:

| Option | Default (SharePoint Settings) | Per mapping |
| --- | --- | --- |
| Keep a local copy after upload | Off | Default / On / Off |
| Remove from SharePoint when the attachment is deleted | Off | Default / On / Off |
| Also copy linked files into SharePoint | Off | Site-wide only |
| If this app is uninstalled | (ask) | Site-wide only |

**Default** follows SharePoint Settings, and the mapping shows what that currently is. **On** or
**Off** applies to that DocType's attachments whatever SharePoint Settings says. Typical uses:
keep local copies for the one DocType being trialled while the others run fully in SharePoint,
or set removal **Off** for a DocType whose documents must be retained even if they are deleted
here. A disabled mapping's override still applies to its files already in SharePoint.

### What moves and what stays

| Attachment | Behaviour |
| --- | --- |
| A file attached to a record of a mapped DocType (sidebar / Attach button) | Moved |
| A file behind a field (image, signature, an `Attach` field) | Stays local: print formats and web pages fetch these without a session |
| A web link attached to a record | Stays a link. With **Also copy linked files into SharePoint** ticked, its target is downloaded once and archived in the folder; the link itself is kept |
| Attachments on unmapped DocTypes | Untouched |
| Attachments that existed before the DocType was mapped | Untouched until **Move Existing Attachments** on the mapping |

Uploads under 4 MB go up in a single request; larger ones go through a resumable upload session
in 10 MB parts. A failed upload leaves the file where it was, marked `Failed` with the reason,
and the hourly job retries it up to five times.

## Permissions: Sites.Selected, not Sites.ReadWrite.All

Document storage signs in as the **application** (OAuth client credentials), never as a person.
A shared library must not stop working because whoever connected it leaves the company.

The application permission it needs is **`Sites.Selected`**. On its own it grants access to no
site at all. A SharePoint administrator then grants the app access to each mapped site, one by
one. The tenant's other sites, including HR and finance sites, stay out of reach. The app never
asks for `Sites.ReadWrite.All`.

### Setup

1. **Entra admin center → App registrations →** the app already used by this integration **→ API
   permissions → Add a permission → Microsoft Graph → Application permissions →
   `Sites.Selected` → Grant admin consent.**
   Pick it under **Microsoft Graph**. The **SharePoint** API lists a permission with the same name,
   and that one does nothing for Graph calls: the symptom is a 401 `generalException` on the first
   request.
2. Make sure **Tenant ID** in Microsoft Settings is the directory (tenant) ID, not `common`.
   Client-credential sign-in has no `/common` endpoint.
3. In Microsoft Settings, tick **SharePoint document storage** and save. Its options (local
   copies, linked files, deletes, uninstall) are on **SharePoint Settings**, in the sidebar, which also has its tools. Then create a
   **SharePoint Mapping** (sidebar *SharePoint → Mappings*, then **Add**):
   - **DocType**: the DocType whose attachments go to SharePoint. The mapping is named after it,
     so each DocType can be mapped once.
   - **SharePoint Site URL**: in Teams, open the channel's *Files* tab and choose *Open in
     SharePoint*. Paste the address; anything after `/sites/<name>` is ignored.
   - **Document Library**: `Documents` for Teams.
   - **Base Folder**: the channel name, e.g. `Projects`. Nested paths like `Projects/2026` work.
     The folder is created if it is missing.
   - **Folder Name Pattern**: `{name}` by default. Use `{fieldname}` for any field, e.g.
     `{name} - {customer}`.
4. Save the mapping, then click **Site Grant Script** on it. A SharePoint or Global administrator
   runs the script in Microsoft Graph PowerShell. (The *Site Grant Script* section of
   *SharePoint Settings* gives one script for every mapped site.) For each site it runs:

   ```powershell
   Connect-MgGraph -Scopes "Sites.FullControl.All"
   $site = Get-MgSite -SiteId "contoso.sharepoint.com:/sites/Sales"
   New-MgSitePermission -SiteId $site.Id -BodyParameter @{
       roles = @("write")
       grantedToIdentities = @(@{ application = @{ id = "<client id>"; displayName = "Frappe" } })
   }
   ```

5. **Test** on the mapping (or *Troubleshoot → Test SharePoint Connection* on Microsoft Settings
   or SharePoint Settings, for all of them). It signs in as the app, resolves the site and library, and creates and
   removes a `frappe-connection-test` folder in the base folder, so a read-only grant shows up now
   rather than on the first upload.
6. Optionally, **Move Existing Attachments** on the mapping queues what is already attached.

### Private and shared channels

A standard channel keeps its files in the team's own site. A **private** or **shared** channel gets
a separate SharePoint site of its own, usually `https://<tenant>.sharepoint.com/sites/<Team>-<Channel>`,
with the channel's files in a folder of that site's `Documents` library. To use one:

- **Site URL**: the channel's own site. In Teams, open the channel's *Files* tab → *Open in
  SharePoint*.
- **Library**: `Documents`. Untick **Group record folders in a base folder**, so each record's
  folder sits directly in the channel's Files tab rather than inside an extra folder.
- **Grant the app that site.** A grant on the parent team's site does not reach a private
  channel's site. The mapping's Site Grant Script names that site.

The app reads and writes as itself, not as a channel member. Who can open the files from Frappe
is therefore decided by Frappe permissions on the DocType, not by channel membership. Restrict
the DocType's roles accordingly if the channel is private for a reason.

Uploads run on the `long` queue, so the site needs a background worker. That is always the case
on Frappe Cloud. **Run Diagnostics** checks it.

## Uninstalling

Before `bench --site <site> uninstall-app frappe_microsoft365` removes anything, it deals with the
attachments stored in SharePoint as chosen in **SharePoint Settings → If this app
is uninstalled** (or `"microsoft365_uninstall_files": "restore"` / `"leave"` in `site_config.json`):

- **Bring files back to this server**: each file is downloaded back and its attachment works as
  before. The free disk space is checked first. If the files would not fit with 10% to spare,
  nothing is downloaded and the uninstall stops, so it can never fill a hosted server.
- **Leave files in SharePoint**: nothing is downloaded. Each attachment becomes a link to its
  SharePoint copy, which opens for people with access to the site.
- **No choice made**: in a terminal it asks; otherwise (Frappe Cloud's dashboard, for example) the
  uninstall stops and says where to choose.

If any file cannot be handled (usually a broken connection), the uninstall stops and lists the
files rather than leave one unopenable. Fix the cause and run it again; files already handled
stay handled. SharePoint copies are never deleted. The app then deletes the custom fields it added
and drops their columns. `--dry-run` reports the file count, total size and free space, and
changes nothing.

## Using it from code

`frappe_microsoft365.microsoft_files` is plain Python that other apps can call:

```python
from frappe_microsoft365 import microsoft_files as m

folder = m.ensure_folder("Project", "PROJ-0001")    # created on first use
m.upload_stream(folder.drive_id, folder.item_id, "report.pdf", fh, size, "application/pdf")
m.list_folder("Project", "PROJ-0001")               # whitelisted; checks read permission
```

Any Graph call can be made as the application with
`microsoft_graph.graph_request(method, path, microsoft_graph.APP_ONLY, ...)`.

## Frappe v15

Everything works on v15 except one thing. Server-side code that reads a moved attachment's bytes
(`File.get_content()`, used for example to attach a file to an outgoing email) needs Frappe v16's
`extend_doctype_class` hook. On v15 those readers get no content, while opening and downloading
from the desk work as normal. If you rely on emailing attachments of a mapped DocType on v15,
keep a local copy.

## Graph endpoints used

All Graph v1.0, application permission `Sites.Selected`:

| Purpose | Request |
| --- | --- |
| Site by URL | `GET /sites/{hostname}:/{server-relative-path}` |
| Libraries | `GET /sites/{site-id}/drives`, `GET /sites/{site-id}/drive` |
| Item by path | `GET /drives/{drive-id}/root:/{path}:`, `GET /drives/{drive-id}/items/{id}:/{name}:` |
| Create folder | `POST /drives/{drive-id}/items/{parent-id}/children` (`conflictBehavior: fail`) |
| Upload ≤ 4 MB | `PUT /drives/{drive-id}/items/{parent-id}:/{name}:/content?@microsoft.graph.conflictBehavior=rename` |
| Upload > 4 MB | `POST …:/{name}:/createUploadSession`, then `PUT {uploadUrl}` in 320 KiB-multiple ranges, without a bearer token |
| Download | `GET /drives/{drive-id}/items/{item-id}/content` (302 to a pre-authenticated URL) |
| List | `GET /drives/{drive-id}/items/{item-id}/children` |
| Rename (record renamed) | `PATCH /drives/{drive-id}/items/{item-id}` |
| Remove (opt-in) | `DELETE /drives/{drive-id}/items/{item-id}`, which goes to the site recycle bin |
