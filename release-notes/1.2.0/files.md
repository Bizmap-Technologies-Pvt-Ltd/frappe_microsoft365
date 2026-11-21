### Security

- Archived links are downloaded from the address that was checked: the server no longer looks the name up a second time to connect, and every redirect is checked the same way.
- A moved file's kept local copy now follows the File's privacy (a File made private takes its copy out of public files) and is removed when the File is deleted and no other File uses it.

### Upgrade notes

- **Kept local copies (Keep Local Copy on).** Migrate moves the kept copy of every moved File that is private but whose copy is still in `public/files` to `private/files`. Nothing is deleted and nothing is fetched from SharePoint. Copies of Files that were deleted before 1.2.0 are not touched: the upgrade report lists them (`upgrade.sharepoint_orphaned_kept_copies`) so you can delete them from the site's files folders yourself. On a site that never switched SharePoint storage on, the patch and the check do nothing.
- **Link archiving** connects to the link's host directly and no longer goes through an HTTP(S) proxy set in the worker's environment; a proxy would resolve the name again.
