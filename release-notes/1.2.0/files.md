### Security

- Archived links are downloaded from the address that was checked: the server no longer looks the name up a second time to connect, and every redirect is checked the same way.
- A moved file's kept local copy now follows the File's privacy (a File made private takes its copy out of public files) and is removed when the File is deleted and no other File uses it.

### Fixed

- Public attachments moved to SharePoint open for visitors who are not signed in again, as they did from `/files` before the move: images on web pages, links in customer emails. Private ones still need a signed-in user who may read them. This also fixes public files that were moved before 1.2.0. (#37)
- Old `/files` and `/private/files` links of a moved file redirect to it with Keep Local Copy on too (they used to answer Forbidden), and when several moved Files share one disk name the link goes to the one the user may read. (#39)
- A SharePoint library or site that was deleted and created again under the same address no longer breaks every new folder: when the cached library id answers 404 and the library itself is gone, the app looks the site and library up again from the mapping's address, once, and carries on. (#40)

- Long paths: SharePoint refuses a file whose whole path passes 400 characters, and uploads used to fail with a path error. A mapping whose base folder leaves too little room is now refused when saved, with the length in the message; record folders are named to leave room for their files; and a file name that would pass the limit is shortened, keeping its extension. Files more than 10 folders deep inside a record's folder open again from the panel. (#41)

### Changed

- A record's SharePoint folder is never renamed by the app. Renaming a record keeps its folder linked under the new name; the folder keeps the name it was created with. (In practice the rename never ran: Frappe had already moved the folder link to the new name before the app looked for it.) The user guide and the setup guide said a folder is renamed to match; they no longer do. (#38)

### Upgrade notes

- **Kept local copies (Keep Local Copy on).** Migrate moves the kept copy of every moved File that is private but whose copy is still in `public/files` to `private/files`. Nothing is deleted and nothing is fetched from SharePoint. Copies of Files that were deleted before 1.2.0 are not touched: the upgrade report lists them (`upgrade.sharepoint_orphaned_kept_copies`) so you can delete them from the site's files folders yourself. On a site that never switched SharePoint storage on, the patch and the check do nothing.
- **Link archiving** connects to the link's host directly and no longer goes through an HTTP(S) proxy set in the worker's environment; a proxy would resolve the name again.
- **Folder names.** Nothing changes for existing folders: they were never renamed, and still are not.
- **Long base folders.** The upgrade report lists mappings whose library and base folder already use more than 300 of SharePoint's 400 path characters (`upgrade.sharepoint_long_paths`). Such a mapping cannot be saved again until its Base Folder is shortened.
