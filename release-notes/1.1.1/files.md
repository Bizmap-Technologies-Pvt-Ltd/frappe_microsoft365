### Security
- SharePoint storage: only the app can set where a File is stored in SharePoint. The Microsoft 365 fields on File and the `open_file` address are refused from the form and the API, and opening, reading (`File.get_content`) or deleting a moved file now requires it to be Stored by the app and inside a known record folder.
- SharePoint storage: each record gets a folder of its own. A folder that already has the record folder's name is never reused, and a folder that is already another record's cannot be linked to a second one.

### Changed
- New record folders are created with SharePoint's "rename" rule: when `Projects/PRJ-0001` is already there, the new record's folder is `Projects/PRJ-0001 1`. Base folders are shared as before.
- The folder of a SharePoint Mapping Folders row can no longer be edited; delete the row and add a new one (with Existing Folder) to point a record at another folder. Ids typed into a new row are ignored.
- A record's folder is recorded (committed) as soon as it is created, and parallel uploads for one record wait for each other, so a failed or concurrent upload never leaves a second folder behind.

### Added
- A *Folder ID* field on File, filled by the upload job from now on with the record folder the file went into. It is created on migrate on sites with SharePoint storage switched on.

### Upgrade notes
- `bench migrate` adds the File field above (storage-enabled sites only) and changes no data.
- The upgrade report lists, for review: File rows whose SharePoint item does not look like the app's own work (`upgrade.sharepoint_file_items`: an item that is a record folder, SharePoint fields the app would not have written, or an address outside every known record folder); and records that share one SharePoint folder (`upgrade.sharepoint_shared_folders`). The file check reads no Graph: for files moved before 1.1.1 it compares each file's recorded SharePoint address with the folders' addresses, so a folder renamed since the move also appears there. Shared folders are not split automatically: delete the Folders rows of the records that should not have the folder; their next attachment gets a new one.
- Files moved before 1.1.1 have no recorded folder, so opening or deleting one costs one extra Graph request (more for files in subfolders) to confirm where it is. A file that SharePoint places outside every record folder (moved there by hand, or a tampered row) no longer opens and is not deleted with its attachment; the event is logged.
- Any code that set the Microsoft 365 File fields with `doc.save()`/`insert()` now fails; the app itself writes them with `frappe.db.set_value`/`db_set`. Attachment copies (`File.create_attachment_copy`, Duplicate) still work when they repeat the original's fields unchanged and the user may read the original.
- `microsoft_files.ensure_folder` now commits when it creates a folder.
