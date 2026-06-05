# Set-SPOSharedChannelLibraryReadOnly.ps1

Freezes a single SharePoint Online document library as read-only for archive purposes. The script targets the document library represented by the supplied URL and any files/folders that already have unique permissions.

## Prerequisites

- PowerShell 7.
- `PnP.PowerShell` installed for the operator account:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

- A Microsoft Entra app/client ID that is allowed to authenticate with PnP interactive login.
- SharePoint permissions high enough to read and update library and item-level permissions.

The script authenticates with:

```powershell
Connect-PnPOnline -Url <SiteUrl> -Interactive -ClientId <ClientId>
```

## Preview Mode

Preview is the default. It parses the library URL, connects to the site, scans the library and unique item scopes, and writes all reports without changing permissions.

```powershell
.\Set-SPOSharedChannelLibraryReadOnly.ps1 `
  -LibraryUrl "https://inalfa0.sharepoint.com/sites/T-GL-ENG-TtMProjectsGlobal-NX5TLES/Prod%20-%20NA/Forms/AllItems.aspx" `
  -ClientId "<client-id>" `
  -OutputFolder "C:\Temp\NX5ReadOnly"
```

Preview mode clearly reports: `No permission changes were made.`

## Apply Mode

Apply mode removes write-level roles and adds `Read` for business users and groups, including normal Owners groups and Full Control assignments. Technical/system principals are protected and skipped.

```powershell
.\Set-SPOSharedChannelLibraryReadOnly.ps1 `
  -LibraryUrl "https://inalfa0.sharepoint.com/sites/T-GL-ENG-TtMProjectsGlobal-NX5TLES/Prod%20-%20NA/Forms/AllItems.aspx" `
  -ClientId "<client-id>" `
  -OutputFolder "C:\Temp\NX5ReadOnly" `
  -Apply
```

Apply mode requires the operator to type exactly:

```text
YES_DOWNGRADE_OWNERS_TO_READ
```

If the library itself inherits permissions, the script will not break inheritance. It prompts before continuing with item-level unique scopes only.

## Output Files

Each run creates timestamped files in `OutputFolder`:

- `ReadOnlyPermissionRun_yyyyMMdd_HHmmss.log`
- `ReadOnlyPermissionChanges_yyyyMMdd_HHmmss.csv`
- `ReadOnlyPermissionSkipped_yyyyMMdd_HHmmss.csv`
- `ReadOnlyPermissionSnapshot_yyyyMMdd_HHmmss.json`
- `ReadOnlyPermissionSummary_yyyyMMdd_HHmmss.txt`

The JSON snapshot captures the pre-change permission state for manual review or future restore-script development.

## Validate The Result

Run the script once in preview mode and review the change CSV before applying. After apply mode completes, run the same preview command again. A clean result should show no remaining write-to-read planned changes except permissions added after the apply run.

You can also validate in SharePoint by opening the library permissions page and checking that normal business users, groups, Members, Visitors, Owners, and project/security groups no longer have write-level roles such as `Full Control`, `Edit`, `Contribute`, `Design`, `Approve`, or custom write roles.

## Limitations

- The script does not change Microsoft Teams Shared Channel membership.
- The script does not add or remove Teams channel members.
- The script does not restore inheritance.
- The script does not delete unique permissions.
- The script does not break inheritance on inherited items.
- The script does not process `SharingLinks.*` principals.
- The script keeps technical/system principals protected.
