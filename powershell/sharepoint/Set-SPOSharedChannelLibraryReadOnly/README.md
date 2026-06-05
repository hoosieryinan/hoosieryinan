# Set-SPOSharedChannelLibraryReadOnly

PowerShell script for setting a SharePoint Online document library to read-only for archive purposes.

## Purpose

This script downgrades write-level permissions to Read for:

- Document library permissions
- Folder-level unique permissions
- File-level unique permissions

It is designed for SharePoint Online document libraries, including libraries located under Teams Shared Channel SharePoint sites.

## Important Notes

- This script does not manage Teams Shared Channel membership.
- This script does not restore permission inheritance.
- This script does not delete unique permissions.
- Owners are downgraded to Read by default, based on business archive requirements.
- Technical/system principals are skipped.

## Preview Mode

```powershell
.\Set-SPOSharedChannelLibraryReadOnly.ps1 `
  -LibraryUrl "https://tenant.sharepoint.com/sites/site-name/LibraryName/Forms/AllItems.aspx" `
  -ClientId "<client-id>" `
  -OutputFolder "C:\Temp\ReadOnlyPreview"
