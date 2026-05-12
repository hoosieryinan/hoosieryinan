#requires -Version 7.0
# This script is read-only. It only reads SharePoint Online and OneDrive metadata and exports CSV reports. It does not delete, move, restore, update, relabel, change permissions, or modify any file, folder, library, or site.

<#
.SYNOPSIS
    Read-only SharePoint Online and OneDrive large/risk file type analysis.

.DESCRIPTION
    Scans SharePoint Online sites and OneDrive personal sites in a Microsoft 365 tenant for
    selected large/risk file extensions, then exports detail and summary CSV reports.

    Authentication uses PnP PowerShell interactive login with a Microsoft Entra application
    client ID. Certificate authentication and username/password authentication are not used.

.NOTES
    Requirements:
      - PowerShell 7+
      - PnP.PowerShell module
      - A Microsoft Entra ID application/client ID permitted for PnP interactive login
      - Tenant/site permissions sufficient to read site, library, and file metadata

    Safety:
      - Report-only script.
      - Reads metadata with PnP.PowerShell and exports reports.
      - No cleanup, content modification, permission change, label change, move, restore, or copy operations are performed.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId = '78aeeb14-45ae-4e8c-9f5a-233363c2683e',

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1.7976931348623157E+308)]
    [double]$MinSizeMB = 100,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOneDrive = $true,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSharePoint = $true,

    [Parameter(Mandatory = $false)]
    [switch]$PersistLogin,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$MaxSites
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptStartTime = Get-Date
$timestamp = $scriptStartTime.ToString('yyyyMMdd-HHmmss')
$resolvedOutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
New-Item -Path $resolvedOutputFolder -ItemType Directory -Force | Out-Null

$detailPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Detail-$timestamp.csv"
$summaryByCategoryPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Summary-ByCategory-$timestamp.csv"
$summaryByExtensionPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Summary-ByExtension-$timestamp.csv"
$summaryByUserPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Summary-ByUser-$timestamp.csv"
$summaryBySitePath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Summary-BySite-$timestamp.csv"
$summaryByLocationTypePath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Summary-ByLocationType-$timestamp.csv"
$errorPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Scan-Errors-$timestamp.csv"
$logPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-FileType-Scan-$timestamp.log"

$extensionCategoryMap = @{
    mp4     = 'Video'
    mov     = 'Video'
    avi     = 'Video'
    zip     = 'Archive'
    rar     = 'Archive'
    '7z'    = 'Archive'
    pst     = 'MailBackupImage'
    ost     = 'MailBackupImage'
    bak     = 'MailBackupImage'
    iso     = 'MailBackupImage'
    vhd     = 'MailBackupImage'
    vhdx    = 'MailBackupImage'
    dwg     = 'CADEngineering'
    step    = 'CADEngineering'
    stp     = 'CADEngineering'
    catpart = 'CADEngineering'
    sldprt  = 'CADEngineering'
    psd     = 'DesignSource'
    ai      = 'DesignSource'
    tif     = 'DesignSource'
}

$systemLibraryTitles = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
@('Site Assets', 'Style Library', 'Form Templates', 'Preservation Hold Library', 'Site Pages') | ForEach-Object { [void]$systemLibraryTitles.Add($_) }

$detailRows = [System.Collections.Generic.List[object]]::new()
$errorRows = [System.Collections.Generic.List[object]]::new()
$sitesScanned = 0
$librariesScanned = 0

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logPath -Value $line -Encoding utf8

    switch ($Level) {
        'ERROR' { Write-Error -Message $Message -ErrorAction Continue }
        'WARN'  { Write-Warning -Message $Message }
        default { Write-Host $line }
    }
}

function Add-ScanError {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$SiteUrl,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$LibraryTitle,

        [Parameter(Mandatory = $true)]
        [string]$ErrorStage,

        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    $errorRows.Add([pscustomobject]@{
        SiteUrl      = $SiteUrl
        LibraryTitle = $LibraryTitle
        ErrorStage   = $ErrorStage
        ErrorMessage = $ErrorMessage
    })

    $location = if ([string]::IsNullOrWhiteSpace($LibraryTitle)) { $SiteUrl } else { "$SiteUrl | $LibraryTitle" }
    Write-Log -Level 'ERROR' -Message "$ErrorStage failed for $location. $ErrorMessage"
}

function Connect-ReadOnlyPnP {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    $connectParameters = @{
        Url         = $Url
        Interactive = $true
        ClientId    = $ClientId
    }

    if ($PersistLogin.IsPresent) {
        $connectParameters.PersistLogin = $true
    }

    Connect-PnPOnline @connectParameters
}

function Get-LocationType {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    if ($SiteUrl -like '*-my.sharepoint.com/personal/*') {
        return 'OneDrive'
    }

    return 'SharePoint'
}

function Get-FieldValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$ListItem,

        [Parameter(Mandatory = $true)]
        [string[]]$FieldNames
    )

    if ($null -eq $ListItem -or $null -eq $ListItem.FieldValues) {
        return $null
    }

    foreach ($fieldName in $FieldNames) {
        if ($ListItem.FieldValues.ContainsKey($fieldName)) {
            return $ListItem.FieldValues[$fieldName]
        }
    }

    return $null
}

function Convert-ToInt64Safe {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return [int64]0
    }

    $result = [int64]0
    if ([int64]::TryParse([string]$Value, [ref]$result)) {
        return $result
    }

    return [int64]0
}

function Convert-UserValueToText {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return ''
    }

    foreach ($propertyName in @('Email', 'LookupValue', 'Title')) {
        $property = $Value.PSObject.Properties[$propertyName]
        if ($null -ne $property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return [string]$property.Value
        }
    }

    return [string]$Value
}

function Get-TopValue {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($Rows.Count -eq 0) {
        return ''
    }

    $topGroup = $Rows | Group-Object -Property $PropertyName | Sort-Object -Property Count -Descending | Select-Object -First 1
    if ($null -eq $topGroup) {
        return ''
    }

    return [string]$topGroup.Name
}

function New-EmptyCsvIfNeeded {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string[]]$Headers
    )

    if (-not (Test-Path -Path $Path)) {
        ($Headers -join ',') | Set-Content -Path $Path -Encoding utf8
    }
}

try {
    Write-Log -Message 'Starting read-only SharePoint Online / OneDrive file type scan.'
    Write-Log -Message "Start time: $($scriptStartTime.ToString('o'))"
    Write-Log -Message "Authentication method: Connect-PnPOnline interactive login with ClientId. PersistLogin: $($PersistLogin.IsPresent)"
    Write-Log -Message "Minimum file size: $MinSizeMB MB"
    Write-Log -Message "IncludeOneDrive: $IncludeOneDrive; IncludeSharePoint: $IncludeSharePoint"

    if (-not $IncludeOneDrive -and -not $IncludeSharePoint) {
        throw 'At least one of IncludeOneDrive or IncludeSharePoint must be enabled.'
    }

    Write-Log -Message "Connecting to tenant admin center: $TenantAdminUrl"
    Connect-ReadOnlyPnP -Url $TenantAdminUrl

    Write-Log -Message 'Retrieving tenant sites with Get-PnPTenantSite.'
    $tenantSites = @(Get-PnPTenantSite -IncludeOneDriveSites)
    Write-Log -Message "Tenant sites returned before filtering: $($tenantSites.Count)"

    $targetSites = @(
        $tenantSites |
            Where-Object {
                $url = [string]$_.Url
                $isOneDrive = $url -like '*-my.sharepoint.com/personal/*'
                ($IncludeOneDrive -and $isOneDrive) -or ($IncludeSharePoint -and -not $isOneDrive)
            } |
            Sort-Object -Property Url
    )

    if ($PSBoundParameters.ContainsKey('MaxSites')) {
        $targetSites = @($targetSites | Select-Object -First $MaxSites)
        Write-Log -Message "MaxSites enabled. Limiting scan to first $MaxSites site(s) after filtering."
    }

    Write-Log -Message "Sites selected for scanning: $($targetSites.Count)"

    for ($siteIndex = 0; $siteIndex -lt $targetSites.Count; $siteIndex++) {
        $site = $targetSites[$siteIndex]
        $siteUrl = [string]$site.Url
        $locationType = Get-LocationType -SiteUrl $siteUrl
        Write-Progress -Id 1 -Activity 'Scanning SharePoint Online / OneDrive sites' -Status $siteUrl -PercentComplete ((($siteIndex + 1) / [Math]::Max($targetSites.Count, 1)) * 100)

        try {
            Write-Log -Message "Connecting to site ($($siteIndex + 1)/$($targetSites.Count)): $siteUrl"
            Connect-ReadOnlyPnP -Url $siteUrl
            $sitesScanned++
        }
        catch {
            Add-ScanError -SiteUrl $siteUrl -LibraryTitle '' -ErrorStage 'SiteConnect' -ErrorMessage $_.Exception.Message
            continue
        }

        $libraries = @()
        try {
            $libraries = @(
                Get-PnPList -Includes BaseTemplate, BaseType, Hidden, Title, RootFolder, ItemCount |
                    Where-Object {
                        ($_.BaseTemplate -eq 101 -or [string]$_.BaseType -eq 'DocumentLibrary') -and
                        (-not $_.Hidden) -and
                        (-not $systemLibraryTitles.Contains([string]$_.Title))
                    } |
                    Sort-Object -Property Title
            )
            Write-Log -Message "Found $($libraries.Count) document library/libraries to scan in $siteUrl"
        }
        catch {
            Add-ScanError -SiteUrl $siteUrl -LibraryTitle '' -ErrorStage 'GetLibraries' -ErrorMessage $_.Exception.Message
            continue
        }

        for ($libraryIndex = 0; $libraryIndex -lt $libraries.Count; $libraryIndex++) {
            $library = $libraries[$libraryIndex]
            $libraryTitle = [string]$library.Title
            Write-Progress -Id 2 -ParentId 1 -Activity 'Scanning document libraries' -Status "$libraryTitle ($siteUrl)" -PercentComplete ((($libraryIndex + 1) / [Math]::Max($libraries.Count, 1)) * 100)

            try {
                $librariesScanned++
                Write-Log -Message "Scanning library ($($libraryIndex + 1)/$($libraries.Count)): $libraryTitle in $siteUrl"

                $query = @'
<View Scope="RecursiveAll">
  <Query>
    <Where>
      <Eq>
        <FieldRef Name="FSObjType" />
        <Value Type="Integer">0</Value>
      </Eq>
    </Where>
  </Query>
</View>
'@

                $items = Get-PnPListItem -List $libraryTitle -PageSize 500 -Fields 'FileLeafRef', 'FileRef', 'File_x0020_Size', 'SMTotalFileStreamSize', 'Created', 'Author', 'Modified', 'Editor', 'FSObjType' -Query $query

                foreach ($item in $items) {
                    try {
                        $fileName = [string](Get-FieldValue -ListItem $item -FieldNames @('FileLeafRef'))
                        if ([string]::IsNullOrWhiteSpace($fileName)) {
                            continue
                        }

                        $fileExtension = [System.IO.Path]::GetExtension($fileName).TrimStart('.').ToLowerInvariant()
                        if (-not $extensionCategoryMap.ContainsKey($fileExtension)) {
                            continue
                        }

                        $sizeBytes = Convert-ToInt64Safe -Value (Get-FieldValue -ListItem $item -FieldNames @('File_x0020_Size', 'SMTotalFileStreamSize'))
                        $sizeMB = [Math]::Round(($sizeBytes / 1MB), 4)
                        if ($sizeMB -lt $MinSizeMB) {
                            continue
                        }

                        $sizeGB = [Math]::Round(($sizeBytes / 1GB), 4)
                        $detailRows.Add([pscustomobject]@{
                            SiteUrl        = $siteUrl
                            LocationType   = $locationType
                            LibraryTitle   = $libraryTitle
                            FileName       = $fileName
                            FileExtension  = $fileExtension
                            FileCategory   = $extensionCategoryMap[$fileExtension]
                            FileUrl        = [string](Get-FieldValue -ListItem $item -FieldNames @('FileRef'))
                            SizeBytes      = [int64]$sizeBytes
                            SizeMB         = [double]$sizeMB
                            SizeGB         = [double]$sizeGB
                            CreatedDate    = Get-FieldValue -ListItem $item -FieldNames @('Created')
                            CreatedBy      = Convert-UserValueToText -Value (Get-FieldValue -ListItem $item -FieldNames @('Author'))
                            ModifiedDate   = Get-FieldValue -ListItem $item -FieldNames @('Modified')
                            ModifiedBy     = Convert-UserValueToText -Value (Get-FieldValue -ListItem $item -FieldNames @('Editor'))
                        })
                    }
                    catch {
                        Add-ScanError -SiteUrl $siteUrl -LibraryTitle $libraryTitle -ErrorStage 'ProcessFile' -ErrorMessage $_.Exception.Message
                        continue
                    }
                }
            }
            catch {
                Add-ScanError -SiteUrl $siteUrl -LibraryTitle $libraryTitle -ErrorStage 'ScanLibrary' -ErrorMessage $_.Exception.Message
                continue
            }
        }

        Write-Progress -Id 2 -Activity 'Scanning document libraries' -Completed
    }

    Write-Progress -Id 1 -Activity 'Scanning SharePoint Online / OneDrive sites' -Completed
}
finally {
    $scriptEndTime = Get-Date
    Write-Log -Message 'Exporting reports.'

    if ($detailRows.Count -gt 0) {
        $detailRows | Export-Csv -Path $detailPath -NoTypeInformation -Encoding utf8

        $detailRows |
            Group-Object -Property FileCategory |
            ForEach-Object {
                $rows = @($_.Group)
                [pscustomobject]@{
                    FileCategory = $_.Name
                    FileCount    = [int]$rows.Count
                    TotalSizeGB  = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Sum).Sum), 4)
                    AverageSizeMB = [double][Math]::Round((($rows | Measure-Object -Property SizeMB -Average).Average), 4)
                    LargestFileGB = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Maximum).Maximum), 4)
                }
            } | Sort-Object -Property FileCategory | Export-Csv -Path $summaryByCategoryPath -NoTypeInformation -Encoding utf8

        $detailRows |
            Group-Object -Property FileExtension, FileCategory |
            ForEach-Object {
                $rows = @($_.Group)
                [pscustomobject]@{
                    FileExtension = [string]$rows[0].FileExtension
                    FileCategory  = [string]$rows[0].FileCategory
                    FileCount     = [int]$rows.Count
                    TotalSizeGB   = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Sum).Sum), 4)
                    AverageSizeMB = [double][Math]::Round((($rows | Measure-Object -Property SizeMB -Average).Average), 4)
                    LargestFileGB = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Maximum).Maximum), 4)
                }
            } | Sort-Object -Property FileCategory, FileExtension | Export-Csv -Path $summaryByExtensionPath -NoTypeInformation -Encoding utf8

        $detailRows |
            Group-Object -Property ModifiedBy |
            ForEach-Object {
                $rows = @($_.Group)
                [pscustomobject]@{
                    ModifiedBy          = $_.Name
                    FileCount           = [int]$rows.Count
                    TotalSizeGB         = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Sum).Sum), 4)
                    TopFileCategory     = Get-TopValue -Rows $rows -PropertyName 'FileCategory'
                    TopFileExtension    = Get-TopValue -Rows $rows -PropertyName 'FileExtension'
                    OneDriveFileCount   = [int]@($rows | Where-Object { $_.LocationType -eq 'OneDrive' }).Count
                    SharePointFileCount = [int]@($rows | Where-Object { $_.LocationType -eq 'SharePoint' }).Count
                }
            } | Sort-Object -Property TotalSizeGB -Descending | Export-Csv -Path $summaryByUserPath -NoTypeInformation -Encoding utf8

        $detailRows |
            Group-Object -Property SiteUrl, LocationType |
            ForEach-Object {
                $rows = @($_.Group)
                [pscustomobject]@{
                    SiteUrl          = [string]$rows[0].SiteUrl
                    LocationType     = [string]$rows[0].LocationType
                    FileCount        = [int]$rows.Count
                    TotalSizeGB      = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Sum).Sum), 4)
                    TopFileCategory  = Get-TopValue -Rows $rows -PropertyName 'FileCategory'
                    TopFileExtension = Get-TopValue -Rows $rows -PropertyName 'FileExtension'
                }
            } | Sort-Object -Property TotalSizeGB -Descending | Export-Csv -Path $summaryBySitePath -NoTypeInformation -Encoding utf8

        $detailRows |
            Group-Object -Property LocationType |
            ForEach-Object {
                $rows = @($_.Group)
                [pscustomobject]@{
                    LocationType = $_.Name
                    SiteCount    = [int]@($rows | Select-Object -ExpandProperty SiteUrl -Unique).Count
                    FileCount    = [int]$rows.Count
                    TotalSizeGB  = [double][Math]::Round((($rows | Measure-Object -Property SizeGB -Sum).Sum), 4)
                }
            } | Sort-Object -Property LocationType | Export-Csv -Path $summaryByLocationTypePath -NoTypeInformation -Encoding utf8
    }
    else {
        New-EmptyCsvIfNeeded -Path $detailPath -Headers @('SiteUrl', 'LocationType', 'LibraryTitle', 'FileName', 'FileExtension', 'FileCategory', 'FileUrl', 'SizeBytes', 'SizeMB', 'SizeGB', 'CreatedDate', 'CreatedBy', 'ModifiedDate', 'ModifiedBy')
        New-EmptyCsvIfNeeded -Path $summaryByCategoryPath -Headers @('FileCategory', 'FileCount', 'TotalSizeGB', 'AverageSizeMB', 'LargestFileGB')
        New-EmptyCsvIfNeeded -Path $summaryByExtensionPath -Headers @('FileExtension', 'FileCategory', 'FileCount', 'TotalSizeGB', 'AverageSizeMB', 'LargestFileGB')
        New-EmptyCsvIfNeeded -Path $summaryByUserPath -Headers @('ModifiedBy', 'FileCount', 'TotalSizeGB', 'TopFileCategory', 'TopFileExtension', 'OneDriveFileCount', 'SharePointFileCount')
        New-EmptyCsvIfNeeded -Path $summaryBySitePath -Headers @('SiteUrl', 'LocationType', 'FileCount', 'TotalSizeGB', 'TopFileCategory', 'TopFileExtension')
        New-EmptyCsvIfNeeded -Path $summaryByLocationTypePath -Headers @('LocationType', 'SiteCount', 'FileCount', 'TotalSizeGB')
    }

    if ($errorRows.Count -gt 0) {
        $errorRows | Export-Csv -Path $errorPath -NoTypeInformation -Encoding utf8
    }
    else {
        New-EmptyCsvIfNeeded -Path $errorPath -Headers @('SiteUrl', 'LibraryTitle', 'ErrorStage', 'ErrorMessage')
    }

    Write-Log -Message "End time: $($scriptEndTime.ToString('o'))"
    Write-Log -Message "Sites scanned: $sitesScanned"
    Write-Log -Message "Libraries scanned: $librariesScanned"
    Write-Log -Message "Matching files found: $($detailRows.Count)"
    Write-Log -Message "Errors logged: $($errorRows.Count)"
    Write-Log -Message "Detail report: $detailPath"
    Write-Log -Message "Summary by category: $summaryByCategoryPath"
    Write-Log -Message "Summary by extension: $summaryByExtensionPath"
    Write-Log -Message "Summary by user: $summaryByUserPath"
    Write-Log -Message "Summary by site: $summaryBySitePath"
    Write-Log -Message "Summary by location type: $summaryByLocationTypePath"
    Write-Log -Message "Error log: $errorPath"
    Write-Log -Message "Text log: $logPath"
}
