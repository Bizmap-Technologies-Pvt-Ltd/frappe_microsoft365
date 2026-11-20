### Security
- SharePoint storage: each record gets a folder of its own. A folder that already has the record folder's name is never reused, and a folder that is already another record's cannot be linked to a second one.

### Changed
- New record folders are created with SharePoint's "rename" rule: when `Projects/PRJ-0001` is already there, the new record's folder is `Projects/PRJ-0001 1`. Base folders are shared as before.
- The folder of a SharePoint Mapping Folders row can no longer be edited; delete the row and add a new one (with Existing Folder) to point a record at another folder. Ids typed into a new row are ignored.
- A record's folder is recorded (committed) as soon as it is created, and parallel uploads for one record wait for each other, so a failed or concurrent upload never leaves a second folder behind.

### Upgrade notes
- The upgrade report lists records that share one SharePoint folder (`upgrade.sharepoint_shared_folders`). Shared folders are not split automatically: delete the Folders rows of the records that should not have the folder; their next attachment gets a new one.
- `microsoft_files.ensure_folder` now commits when it creates a folder.
