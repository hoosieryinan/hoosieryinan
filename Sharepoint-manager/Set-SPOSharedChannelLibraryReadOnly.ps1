<# 
.SYNOPSIS
Freezes one SharePoint Online document library as read-only by converting write-level role assignments to Read.

.DESCRIPTION
This script targets a single SharePoint Online document library, including the library permission scope and
all files/folders that already have unique permissions. It does not manage Microsoft Teams Shared Channel
membership, does not restore inheritance, does not delete unique permissions, and does not break inheritance
on inherited items.

Example preview:
  .\Set-SPOSharedChannelLibraryReadOnly.ps1 `
    -LibraryUrl "https://inalfa0.sharepoint.com/sites/T-GL-ENG-TtMProjectsGlobal-NX5TLES/Prod%20-%20NA/Forms/AllItems.aspx" `
    -ClientId "<client-id>" `
    -OutputFolder "C:\Temp\NX5ReadOnly"

Example apply:
  .\Set-SPOSharedChannelLibraryReadOnly.ps1 `
    -LibraryUrl "https://inalfa0.sharepoint.com/sites/T-GL-ENG-TtMProjectsGlobal-NX5TLES/Prod%20-%20NA/Forms/AllItems.aspx" `
    -ClientId "<client-id>" `
    -OutputFolder "C:\Temp\NX5ReadOnly" `
    -Apply
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LibraryUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [switch]$Apply,

    [ValidateRange(1, 5000)]
    [int]$PageSize = 500,

    [ValidateNotNullOrEmpty()]
    [string]$TargetRole = "Read",

    [string[]]$ProtectedPrincipalPatterns = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:RunStartedAt = Get-Date
$script:Mode = if ($Apply) { "Apply" } else { "Preview" }
$script:Timestamp = $script:RunStartedAt.ToString("yyyyMMdd_HHmmss")
$script:ChangeRecords = [System.Collections.Generic.List[object]]::new()
$script:SkippedRecords = [System.Collections.Generic.List[object]]::new()
$script:SnapshotRecords = [System.Collections.Generic.List[object]]::new()
$script:TotalItemsScanned = 0
$script:TotalFoldersScanned = 0
$script:TotalFilesScanned = 0
$script:TotalUniquePermissionObjects = 0
$script:TotalPlannedChanges = 0
$script:TotalAppliedChanges = 0
$script:TotalSkipped = 0
$script:TotalFailures = 0

$script:LogPath = $null
$script:ChangesPath = $null
$script:SkippedPath = $null
$script:SnapshotPath = $null
$script:SummaryPath = $null

$script:DefaultProtectedPrincipalPatterns = @(
    "^System Account$",
    "SHAREPOINT\\system",
    "AAD Service Account",
    "SharePoint App Principal",
    "Limited Access System Group",
    "^SharingLinks\.",
    "Company Administrator",
    "Global Administrator",
    "SharePoint Service Administrator",
    "Tenant Administrator",
    "c:0t\.c\|tenant\|",
    "i:0i\.t\|",
    "app@sharepoint",
    "spo-grid-all-users",
    "urn:spo:guest",
    "_layouts/",
    "00000003-0000-0ff1-ce00-000000000000"
)

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "VERBOSE")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date).ToString("s"), $Level, $Message
    if ($script:LogPath) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }

    switch ($Level) {
        "WARN"    { Write-Warning $Message }
        "ERROR"   { Write-Error $Message -ErrorAction Continue }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        "VERBOSE" { Write-Verbose $Message }
        default   { Write-Host $Message }
    }
}

function ConvertTo-ServerRelativeUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$UrlPath)

    $decoded = [System.Uri]::UnescapeDataString($UrlPath)
    if (-not $decoded.StartsWith("/")) {
        $decoded = "/$decoded"
    }
    return $decoded.TrimEnd("/")
}

function Parse-LibraryUrl {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Url)

    $uri = [System.Uri]::new($Url)
    $path = ConvertTo-ServerRelativeUrl -UrlPath $uri.AbsolutePath
    $segments = $path.Trim("/").Split("/", [System.StringSplitOptions]::RemoveEmptyEntries)

    if ((@($segments).Count) -lt 1) {
        throw "LibraryUrl does not contain a SharePoint path: $Url"
    }

    $siteRelativePath = ""
    if ($segments[0] -in @("sites", "teams", "personal") -and (@($segments).Count) -ge 2) {
        $siteRelativePath = "/{0}/{1}" -f $segments[0], $segments[1]
    }

    $siteUrl = if ([string]::IsNullOrWhiteSpace($siteRelativePath)) {
        $uri.GetLeftPart([System.UriPartial]::Authority)
    } else {
        "{0}{1}" -f $uri.GetLeftPart([System.UriPartial]::Authority), $siteRelativePath
    }

    $formsIndex = -1
    for ($i = 0; $i -lt (@($segments).Count); $i++) {
        if ($segments[$i] -ieq "Forms") {
            $formsIndex = $i
            break
        }
    }

    if ($formsIndex -gt 0) {
        $librarySegments = $segments[0..($formsIndex - 1)]
    } elseif ((@($segments).Count) -gt 2 -and $segments[0] -in @("sites", "teams", "personal")) {
        $librarySegments = $segments[0..2]
    } elseif ((@($segments).Count) -ge 1) {
        $librarySegments = @($segments[0])
    } else {
        throw "Unable to infer document library path from URL: $Url"
    }

    $libraryServerRelativePath = "/" + ($librarySegments -join "/")
    $libraryName = $librarySegments[-1]

    [pscustomobject]@{
        SiteUrl                   = $siteUrl
        LibraryServerRelativePath = $libraryServerRelativePath
        LibraryNameFromUrl        = $libraryName
        SourceUrl                 = $Url
    }
}

function Get-RoleNames {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$RoleAssignment)

    $roleNames = @()

    try {
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings | Out-Null

        foreach ($roleDefinition in @($RoleAssignment.RoleDefinitionBindings)) {
            try {
                Get-PnPProperty -ClientObject $roleDefinition -Property Name | Out-Null
                if (-not [string]::IsNullOrWhiteSpace($roleDefinition.Name)) {
                    $roleNames += [string]$roleDefinition.Name
                }
            } catch {
                Write-Verbose "Failed to load role definition name: $($_.Exception.Message)"
            }
        }
    } catch {
        Write-Verbose "Failed to load role definition bindings: $($_.Exception.Message)"
    }

    return @($roleNames | Sort-Object -Unique)
}

function Test-ProtectedPrincipal {
    [CmdletBinding()]
    param(
        [AllowNull()]$Principal,
        [string[]]$AdditionalPatterns = @()
    )

    if ($null -eq $Principal) {
        return $true
    }

    $name = if ($Principal.Title) { [string]$Principal.Title } else { "" }
    $login = if ($Principal.LoginName) { [string]$Principal.LoginName } else { "" }
    $principalType = if ($Principal.PrincipalType) { [string]$Principal.PrincipalType } else { "" }
    $candidate = "$name`n$login`n$principalType"
    $patterns = @($script:DefaultProtectedPrincipalPatterns + $AdditionalPatterns)

    foreach ($pattern in $patterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) {
            continue
        }
        if ($candidate -match $pattern) {
            return $true
        }
    }

    return $false
}

function Test-WriteRole {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RoleName)

    $readOnlyRoles = @("Read", "View Only", "Restricted View", "Limited Access")
    return ($RoleName -notin $readOnlyRoles)
}

function Test-OwnerOrFullControlBeforeChange {
    [CmdletBinding()]
    param(
        [AllowNull()]$Principal,
        [string[]]$RoleNames
    )

    $name = if ($Principal -and $Principal.Title) { [string]$Principal.Title } else { "" }
    $login = if ($Principal -and $Principal.LoginName) { [string]$Principal.LoginName } else { "" }
    return (($RoleNames -contains "Full Control") -or ($name -match "Owners") -or ($login -match "Owners"))
}

function Add-ChangeRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ObjectScope,
        [AllowNull()][string]$ItemId,
        [AllowNull()][string]$ObjectName,
        [AllowNull()][string]$ServerRelativeUrl,
        [bool]$HasUniquePermissions,
        [AllowNull()]$Principal,
        [string[]]$OriginalRoles,
        [string[]]$RemovedRoles,
        [AllowNull()][string]$AddedRole,
        [string[]]$FinalPlannedRoles,
        [Parameter(Mandatory = $true)][string]$Action,
        [Parameter(Mandatory = $true)][string]$Reason,
        [Parameter(Mandatory = $true)][string]$SourceUrl,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$LibraryTitle
    )

    $record = [pscustomobject]@{
        Timestamp                         = (Get-Date).ToString("s")
        Mode                              = $script:Mode
        ObjectScope                       = $ObjectScope
        ItemId                            = $ItemId
        ObjectName                        = $ObjectName
        ServerRelativeUrl                 = $ServerRelativeUrl
        HasUniquePermissions              = $HasUniquePermissions
        PrincipalName                     = if ($Principal -and $Principal.Title) { $Principal.Title } else { "" }
        PrincipalLoginName                = if ($Principal -and $Principal.LoginName) { $Principal.LoginName } else { "" }
        PrincipalType                     = if ($Principal -and $Principal.PrincipalType) { [string]$Principal.PrincipalType } else { "" }
        OriginalRoles                     = ($OriginalRoles -join "; ")
        RemovedRoles                      = ($RemovedRoles -join "; ")
        AddedRole                         = $AddedRole
        FinalPlannedRoles                 = ($FinalPlannedRoles -join "; ")
        IsOwnerOrFullControlBeforeChange  = if (Test-OwnerOrFullControlBeforeChange -Principal $Principal -RoleNames $OriginalRoles) { "Yes" } else { "No" }
        Action                            = $Action
        Reason                            = $Reason
        SourceUrl                         = $SourceUrl
        SiteUrl                           = $SiteUrl
        LibraryTitle                      = $LibraryTitle
    }

    $script:ChangeRecords.Add($record)
}

function Add-SkippedRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ObjectScope,
        [AllowNull()][string]$ItemId,
        [AllowNull()][string]$ObjectName,
        [AllowNull()][string]$ServerRelativeUrl,
        [AllowNull()]$Principal,
        [string[]]$CurrentRoles,
        [Parameter(Mandatory = $true)][string]$Reason,
        [Parameter(Mandatory = $true)][string]$SourceUrl,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$LibraryTitle
    )

    $record = [pscustomobject]@{
        Timestamp          = (Get-Date).ToString("s")
        ObjectScope        = $ObjectScope
        ItemId             = $ItemId
        ObjectName         = $ObjectName
        ServerRelativeUrl  = $ServerRelativeUrl
        PrincipalName      = if ($Principal -and $Principal.Title) { $Principal.Title } else { "" }
        PrincipalLoginName = if ($Principal -and $Principal.LoginName) { $Principal.LoginName } else { "" }
        PrincipalType      = if ($Principal -and $Principal.PrincipalType) { [string]$Principal.PrincipalType } else { "" }
        CurrentRoles       = ($CurrentRoles -join "; ")
        Reason             = $Reason
        SourceUrl          = $SourceUrl
        SiteUrl            = $SiteUrl
        LibraryTitle       = $LibraryTitle
    }

    $script:SkippedRecords.Add($record)
    $script:TotalSkipped++
}

function Add-SnapshotRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ObjectScope,
        [AllowNull()][string]$ItemId,
        [AllowNull()][string]$ObjectName,
        [AllowNull()][string]$ServerRelativeUrl,
        [bool]$HasUniquePermissions,
        [AllowNull()]$Principal,
        [string[]]$RoleNames,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$LibraryUrl,
        [Parameter(Mandatory = $true)][string]$LibraryTitle
    )

    $script:SnapshotRecords.Add([pscustomobject]@{
        SiteUrl                          = $SiteUrl
        LibraryUrl                       = $LibraryUrl
        LibraryTitle                     = $LibraryTitle
        ObjectScope                      = $ObjectScope
        ItemId                           = $ItemId
        ObjectName                       = $ObjectName
        ServerRelativeUrl                = $ServerRelativeUrl
        HasUniquePermissions             = $HasUniquePermissions
        PrincipalName                    = if ($Principal -and $Principal.Title) { $Principal.Title } else { "" }
        PrincipalLoginName               = if ($Principal -and $Principal.LoginName) { $Principal.LoginName } else { "" }
        PrincipalType                    = if ($Principal -and $Principal.PrincipalType) { [string]$Principal.PrincipalType } else { "" }
        RoleNames                        = @($RoleNames)
        IsOwnerOrFullControlBeforeChange = if (Test-OwnerOrFullControlBeforeChange -Principal $Principal -RoleNames $RoleNames) { "Yes" } else { "No" }
        CapturedAt                       = (Get-Date).ToString("s")
    })
}

function Convert-RoleAssignmentToRead {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)]$RoleAssignment,
        [Parameter(Mandatory = $true)]$ReadRoleDefinition,
        [Parameter(Mandatory = $true)][string]$ObjectScope,
        [AllowNull()][string]$ItemId,
        [AllowNull()][string]$ObjectName,
        [AllowNull()][string]$ServerRelativeUrl,
        [bool]$HasUniquePermissions,
        [Parameter(Mandatory = $true)][string]$SourceUrl,
        [Parameter(Mandatory = $true)][string]$SiteUrl,
        [Parameter(Mandatory = $true)][string]$LibraryTitle,
        [string[]]$AdditionalProtectedPatterns = @()
    )

    Get-PnPProperty -ClientObject $RoleAssignment -Property Member, RoleDefinitionBindings | Out-Null
    $principal = $RoleAssignment.Member
    $originalRoles = @(Get-RoleNames -RoleAssignment $RoleAssignment)

    Add-SnapshotRecord `
        -ObjectScope $ObjectScope `
        -ItemId $ItemId `
        -ObjectName $ObjectName `
        -ServerRelativeUrl $ServerRelativeUrl `
        -HasUniquePermissions $HasUniquePermissions `
        -Principal $principal `
        -RoleNames $originalRoles `
        -SiteUrl $SiteUrl `
        -LibraryUrl $SourceUrl `
        -LibraryTitle $LibraryTitle

    if (Test-ProtectedPrincipal -Principal $principal -AdditionalPatterns $AdditionalProtectedPatterns) {
        Add-SkippedRecord `
            -ObjectScope $ObjectScope `
            -ItemId $ItemId `
            -ObjectName $ObjectName `
            -ServerRelativeUrl $ServerRelativeUrl `
            -Principal $principal `
            -CurrentRoles $originalRoles `
            -Reason "Protected technical/system principal or matched ProtectedPrincipalPatterns." `
            -SourceUrl $SourceUrl `
            -SiteUrl $SiteUrl `
            -LibraryTitle $LibraryTitle
        return
    }

    $writeRoles = @($originalRoles | Where-Object { Test-WriteRole -RoleName $_ })
    if (@($writeRoles).Count -eq 0) {
        $reason = if ((@($originalRoles).Count -eq 1) -and ($originalRoles -contains "Limited Access")) {
            "Principal has only Limited Access. Left unchanged."
        } else {
            "Principal already has read-only roles only. Left unchanged."
        }

        Add-SkippedRecord `
            -ObjectScope $ObjectScope `
            -ItemId $ItemId `
            -ObjectName $ObjectName `
            -ServerRelativeUrl $ServerRelativeUrl `
            -Principal $principal `
            -CurrentRoles $originalRoles `
            -Reason $reason `
            -SourceUrl $SourceUrl `
            -SiteUrl $SiteUrl `
            -LibraryTitle $LibraryTitle
        return
    }

    $finalPlannedRoles = @(
        $originalRoles |
            Where-Object { $_ -notin $writeRoles } |
            Where-Object { $_ -ne $TargetRole }
        $TargetRole
    ) | Sort-Object -Unique

    if (-not $Apply) {
        $script:TotalPlannedChanges++
        Add-ChangeRecord `
            -ObjectScope $ObjectScope `
            -ItemId $ItemId `
            -ObjectName $ObjectName `
            -ServerRelativeUrl $ServerRelativeUrl `
            -HasUniquePermissions $HasUniquePermissions `
            -Principal $principal `
            -OriginalRoles $originalRoles `
            -RemovedRoles $writeRoles `
            -AddedRole $TargetRole `
            -FinalPlannedRoles $finalPlannedRoles `
            -Action "Planned" `
            -Reason "Write-level role(s) would be removed and $TargetRole would be added." `
            -SourceUrl $SourceUrl `
            -SiteUrl $SiteUrl `
            -LibraryTitle $LibraryTitle
        return
    }

    try {
        $bindingsToRemove = @()
        foreach ($binding in @($RoleAssignment.RoleDefinitionBindings)) {
            if ($binding.Name -in $writeRoles) {
                $bindingsToRemove += $binding
            }
        }

        foreach ($binding in @($bindingsToRemove)) {
            [void]$RoleAssignment.RoleDefinitionBindings.Remove($binding)
        }

        if ($originalRoles -notcontains $TargetRole) {
            [void]$RoleAssignment.RoleDefinitionBindings.Add($ReadRoleDefinition)
        }

        $RoleAssignment.Update()
        $Context.ExecuteQuery()
        $script:TotalAppliedChanges++

        Add-ChangeRecord `
            -ObjectScope $ObjectScope `
            -ItemId $ItemId `
            -ObjectName $ObjectName `
            -ServerRelativeUrl $ServerRelativeUrl `
            -HasUniquePermissions $HasUniquePermissions `
            -Principal $principal `
            -OriginalRoles $originalRoles `
            -RemovedRoles $writeRoles `
            -AddedRole $TargetRole `
            -FinalPlannedRoles $finalPlannedRoles `
            -Action "Changed" `
            -Reason "Write-level role(s) removed and $TargetRole added." `
            -SourceUrl $SourceUrl `
            -SiteUrl $SiteUrl `
            -LibraryTitle $LibraryTitle
    } catch {
        $script:TotalFailures++
        Add-ChangeRecord `
            -ObjectScope $ObjectScope `
            -ItemId $ItemId `
            -ObjectName $ObjectName `
            -ServerRelativeUrl $ServerRelativeUrl `
            -HasUniquePermissions $HasUniquePermissions `
            -Principal $principal `
            -OriginalRoles $originalRoles `
            -RemovedRoles $writeRoles `
            -AddedRole $TargetRole `
            -FinalPlannedRoles $finalPlannedRoles `
            -Action "Failed" `
            -Reason $_.Exception.Message `
            -SourceUrl $SourceUrl `
            -SiteUrl $SiteUrl `
            -LibraryTitle $LibraryTitle
        Write-Log -Level "WARN" -Message "Failed to update $ObjectScope '$ServerRelativeUrl' for principal '$($principal.Title)': $($_.Exception.Message)"
    }
}

function Export-Reports {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Parsed,
        [Parameter(Mandatory = $true)][string]$LibraryTitle
    )

    $changeColumns = @(
        "Timestamp",
        "Mode",
        "ObjectScope",
        "ItemId",
        "ObjectName",
        "ServerRelativeUrl",
        "HasUniquePermissions",
        "PrincipalName",
        "PrincipalLoginName",
        "PrincipalType",
        "OriginalRoles",
        "RemovedRoles",
        "AddedRole",
        "FinalPlannedRoles",
        "IsOwnerOrFullControlBeforeChange",
        "Action",
        "Reason",
        "SourceUrl",
        "SiteUrl",
        "LibraryTitle"
    )

    $skippedColumns = @(
        "Timestamp",
        "ObjectScope",
        "ItemId",
        "ObjectName",
        "ServerRelativeUrl",
        "PrincipalName",
        "PrincipalLoginName",
        "PrincipalType",
        "CurrentRoles",
        "Reason",
        "SourceUrl",
        "SiteUrl",
        "LibraryTitle"
    )

    if ($script:ChangeRecords.Count -gt 0) {
        $script:ChangeRecords |
            Select-Object $changeColumns |
            Export-Csv -LiteralPath $script:ChangesPath -NoTypeInformation -Encoding UTF8
    } else {
        Set-Content -LiteralPath $script:ChangesPath -Value ($changeColumns -join ",") -Encoding UTF8
    }

    if ($script:SkippedRecords.Count -gt 0) {
        $script:SkippedRecords |
            Select-Object $skippedColumns |
            Export-Csv -LiteralPath $script:SkippedPath -NoTypeInformation -Encoding UTF8
    } else {
        Set-Content -LiteralPath $script:SkippedPath -Value ($skippedColumns -join ",") -Encoding UTF8
    }

    $snapshotRoot = [pscustomobject]@{
        SiteUrl      = $Parsed.SiteUrl
        LibraryUrl   = $Parsed.SourceUrl
        LibraryTitle = $LibraryTitle
        CapturedAt   = (Get-Date).ToString("s")
        Entries      = @($script:SnapshotRecords)
    }
    $snapshotRoot |
        ConvertTo-Json -Depth 8 |
        Set-Content -LiteralPath $script:SnapshotPath -Encoding UTF8

    $summary = @(
        "Read-only permission run summary"
        "================================"
        "LibraryUrl: $($Parsed.SourceUrl)"
        "SiteUrl: $($Parsed.SiteUrl)"
        "LibraryTitle: $LibraryTitle"
        "Mode: $script:Mode"
        "Total items scanned: $script:TotalItemsScanned"
        "Total folders scanned: $script:TotalFoldersScanned"
        "Total files scanned: $script:TotalFilesScanned"
        "Total unique permission objects found: $script:TotalUniquePermissionObjects"
        "Total principals planned/changed: $(if ($Apply) { $script:TotalAppliedChanges } else { $script:TotalPlannedChanges })"
        "Total skipped: $script:TotalSkipped"
        "Total failed: $script:TotalFailures"
        ""
        "Output file paths:"
        "Log: $script:LogPath"
        "Changes: $script:ChangesPath"
        "Skipped: $script:SkippedPath"
        "Snapshot: $script:SnapshotPath"
        "Summary: $script:SummaryPath"
    )

    $summary | Set-Content -LiteralPath $script:SummaryPath -Encoding UTF8
}

try {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    $script:LogPath = Join-Path $OutputFolder "ReadOnlyPermissionRun_$script:Timestamp.log"
    $script:ChangesPath = Join-Path $OutputFolder "ReadOnlyPermissionChanges_$script:Timestamp.csv"
    $script:SkippedPath = Join-Path $OutputFolder "ReadOnlyPermissionSkipped_$script:Timestamp.csv"
    $script:SnapshotPath = Join-Path $OutputFolder "ReadOnlyPermissionSnapshot_$script:Timestamp.json"
    $script:SummaryPath = Join-Path $OutputFolder "ReadOnlyPermissionSummary_$script:Timestamp.txt"

    Write-Log -Message "Start time: $($script:RunStartedAt.ToString("s"))"
    Write-Log -Message "LibraryUrl: $LibraryUrl"
    Write-Log -Message "Mode: $script:Mode"

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "PnP.PowerShell is not installed. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }

    $parsed = Parse-LibraryUrl -Url $LibraryUrl
    Write-Log -Message "Parsed SiteUrl: $($parsed.SiteUrl)"
    Write-Log -Message "Parsed library server-relative path: $($parsed.LibraryServerRelativePath)"
    Write-Log -Message "Parsed library name from URL: $($parsed.LibraryNameFromUrl)"

    Write-Host ""
    Write-Host "Target parsed from input URL" -ForegroundColor Cyan
    Write-Host "  SiteUrl: $($parsed.SiteUrl)"
    Write-Host "  Library path: $($parsed.LibraryServerRelativePath)"
    Write-Host "  Library name from URL: $($parsed.LibraryNameFromUrl)"
    Write-Host ""

    if (-not $Apply) {
        Write-Warning "Preview mode. No permission changes were made."
    } else {
        Write-Warning "You are about to downgrade all write permissions to Read, including normal Owners groups and Full Control assignments. This will make the library read-only for business users and owners. Only technical/system principals defined as protected will be skipped."
        $confirmation = Read-Host "Type YES_DOWNGRADE_OWNERS_TO_READ to continue"
        if ($confirmation -ne "YES_DOWNGRADE_OWNERS_TO_READ") {
            throw "Confirmation phrase did not match. Stopping without making changes."
        }
    }

    Write-Log -Message "Connecting with PnP.PowerShell interactive authentication."
    Connect-PnPOnline -Url $parsed.SiteUrl -Interactive -ClientId $ClientId
    $web = Get-PnPWeb
    Write-Log -Level "SUCCESS" -Message "Connection status: connected to '$($web.Title)' at $($web.Url)"

    $ctx = Get-PnPContext
    $readRole = $ctx.Web.RoleDefinitions.GetByName($TargetRole)
    $ctx.Load($readRole)
    try {
        $ctx.ExecuteQuery()
    } catch {
        throw "Target role '$TargetRole' does not exist in this site. Stopping before scanning."
    }

    $lists = Get-PnPList -Includes RootFolder, Title, BaseTemplate, Hidden, ItemCount, HasUniqueRoleAssignments
    $targetList = $null
    foreach ($candidate in $lists) {
        Get-PnPProperty -ClientObject $candidate -Property RootFolder | Out-Null
        $rootUrl = ConvertTo-ServerRelativeUrl -UrlPath $candidate.RootFolder.ServerRelativeUrl
        if ($rootUrl -eq $parsed.LibraryServerRelativePath) {
            $targetList = $candidate
            break
        }
    }

    if ($null -eq $targetList) {
        throw "Could not find a document library whose root folder matches '$($parsed.LibraryServerRelativePath)' on site '$($parsed.SiteUrl)'."
    }

    Get-PnPProperty -ClientObject $targetList -Property RootFolder, HasUniqueRoleAssignments | Out-Null
    $libraryTitle = $targetList.Title
    $libraryRootUrl = ConvertTo-ServerRelativeUrl -UrlPath $targetList.RootFolder.ServerRelativeUrl

    if ($targetList.Hidden -or $targetList.BaseTemplate -ne 101) {
        throw "Target '$libraryTitle' is hidden or is not a standard document library. Hidden/system libraries are not modified."
    }

    Write-Host ""
    Write-Host "Resolved target library" -ForegroundColor Cyan
    Write-Host "  LibraryTitle: $libraryTitle"
    Write-Host "  RootFolder: $libraryRootUrl"
    Write-Host "  ItemCount: $($targetList.ItemCount)"
    Write-Host ""

    Write-Log -Message "Parsed Library: $libraryTitle"
    Write-Log -Message "Library item count: $($targetList.ItemCount)"

    if ($targetList.ItemCount -gt 5000) {
        Write-Log -Level "WARN" -Message "Library has more than 5,000 items. Continuing with paging. PageSize: $PageSize"
    }

    if (-not $targetList.HasUniqueRoleAssignments) {
        Write-Log -Level "WARN" -Message "The library inherits permissions. The script will not break inheritance or modify inherited library-level permissions."
        $continueInherited = Read-Host "Type CONTINUE_WITH_INHERITED_LIBRARY to scan items and process only existing unique item scopes"
        if ($continueInherited -ne "CONTINUE_WITH_INHERITED_LIBRARY") {
            throw "Operator did not confirm processing an inherited-permission library. Stopping without changes."
        }

        Add-SkippedRecord `
            -ObjectScope "Library" `
            -ItemId "" `
            -ObjectName $libraryTitle `
            -ServerRelativeUrl $libraryRootUrl `
            -Principal $null `
            -CurrentRoles @() `
            -Reason "Library inherits permissions. Library-level permissions were not modified because inheritance is not broken by this script." `
            -SourceUrl $LibraryUrl `
            -SiteUrl $parsed.SiteUrl `
            -LibraryTitle $libraryTitle
    } else {
        Write-Log -Message "Processing unique library-level role assignments."
        $script:TotalUniquePermissionObjects++
        Get-PnPProperty -ClientObject $targetList -Property RoleAssignments | Out-Null
        foreach ($roleAssignment in @($targetList.RoleAssignments)) {
            try {
                Convert-RoleAssignmentToRead `
                    -Context $ctx `
                    -RoleAssignment $roleAssignment `
                    -ReadRoleDefinition $readRole `
                    -ObjectScope "Library" `
                    -ItemId "" `
                    -ObjectName $libraryTitle `
                    -ServerRelativeUrl $libraryRootUrl `
                    -HasUniquePermissions $true `
                    -SourceUrl $LibraryUrl `
                    -SiteUrl $parsed.SiteUrl `
                    -LibraryTitle $libraryTitle `
                    -AdditionalProtectedPatterns $ProtectedPrincipalPatterns
            } catch {
                $script:TotalFailures++
                Write-Log -Level "WARN" -Message "Failed to process a library-level role assignment: $($_.Exception.Message)"
            }
        }
    }

    Write-Log -Message "Scanning document library items with PageSize $PageSize."
    $items = Get-PnPListItem -List $targetList -PageSize $PageSize -Fields "FileRef", "FileLeafRef", "FSObjType"
    foreach ($item in $items) {
        $script:TotalItemsScanned++

        $fileRef = if ($item.FieldValues.ContainsKey("FileRef")) { [string]$item.FieldValues["FileRef"] } else { "" }
        $fileLeafRef = if ($item.FieldValues.ContainsKey("FileLeafRef")) { [string]$item.FieldValues["FileLeafRef"] } else { "" }
        $fsObjType = if ($item.FieldValues.ContainsKey("FSObjType")) { [int]$item.FieldValues["FSObjType"] } else { -1 }
        $scope = if ($fsObjType -eq 1) { "Folder" } else { "File" }

        if ($scope -eq "Folder") {
            $script:TotalFoldersScanned++
        } else {
            $script:TotalFilesScanned++
        }

        try {
            Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments | Out-Null
            if (-not $item.HasUniqueRoleAssignments) {
                Write-Verbose "Inherited permissions: $fileRef"
                continue
            }

            $script:TotalUniquePermissionObjects++
            Write-Verbose "Processing unique permissions: $fileRef"
            Get-PnPProperty -ClientObject $item -Property RoleAssignments | Out-Null

            foreach ($roleAssignment in @($item.RoleAssignments)) {
                try {
                    Convert-RoleAssignmentToRead `
                        -Context $ctx `
                        -RoleAssignment $roleAssignment `
                        -ReadRoleDefinition $readRole `
                        -ObjectScope $scope `
                        -ItemId ([string]$item.Id) `
                        -ObjectName $fileLeafRef `
                        -ServerRelativeUrl $fileRef `
                        -HasUniquePermissions $true `
                        -SourceUrl $LibraryUrl `
                        -SiteUrl $parsed.SiteUrl `
                        -LibraryTitle $libraryTitle `
                        -AdditionalProtectedPatterns $ProtectedPrincipalPatterns
                } catch {
                    $script:TotalFailures++
                    Write-Log -Level "WARN" -Message "Failed to process role assignment on item '$fileRef': $($_.Exception.Message)"
                }
            }
        } catch {
            $script:TotalFailures++
            Write-Log -Level "WARN" -Message "Failed to inspect item '$fileRef': $($_.Exception.Message)"
        }
    }

    Export-Reports -Parsed $parsed -LibraryTitle $libraryTitle

    Write-Log -Message "Total scanned items: $script:TotalItemsScanned"
    Write-Log -Message "Total folders scanned: $script:TotalFoldersScanned"
    Write-Log -Message "Total files scanned: $script:TotalFilesScanned"
    Write-Log -Message "Total unique permission objects: $script:TotalUniquePermissionObjects"
    Write-Log -Message "Total planned changes: $script:TotalPlannedChanges"
    Write-Log -Message "Total applied changes: $script:TotalAppliedChanges"
    Write-Log -Message "Total skipped: $script:TotalSkipped"
    Write-Log -Message "Total failures: $script:TotalFailures"
    Write-Log -Message "End time: $((Get-Date).ToString("s"))"

    Write-Host ""
    if (-not $Apply) {
        Write-Warning "Preview complete. No permission changes were made."
    } else {
        Write-Host "Apply complete." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Output files" -ForegroundColor Cyan
    Write-Host "  Log: $script:LogPath"
    Write-Host "  Changes: $script:ChangesPath"
    Write-Host "  Skipped: $script:SkippedPath"
    Write-Host "  Snapshot: $script:SnapshotPath"
    Write-Host "  Summary: $script:SummaryPath"
} catch {
    $script:TotalFailures++
    if ($script:LogPath) {
        Write-Log -Level "ERROR" -Message $_.Exception.Message
        Write-Log -Message "End time: $((Get-Date).ToString("s"))"
    } else {
        Write-Error $_.Exception.Message
    }
    throw
}
