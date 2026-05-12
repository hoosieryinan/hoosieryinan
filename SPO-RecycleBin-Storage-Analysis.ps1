#requires -Version 7.0
# This script only reads recycle bin metadata and exports reports. It does not delete, restore, move, or modify any files.

<#
.SYNOPSIS
    Read-only SharePoint Online and OneDrive for Business recycle bin storage analysis.

.DESCRIPTION
    This script connects to each supplied SharePoint Online or OneDrive personal site, reads
    recycle bin metadata with Get-PnPRecycleBinItem, and exports detailed and summary CSV
    reports to help identify large recycle bin items that may be cleanup candidates.

    This script only reads recycle bin metadata and exports reports. It does not delete,
    restore, move, or modify any files.

.NOTES
    Requirements:
      - PowerShell 7+
      - PnP.PowerShell module
      - A Microsoft Entra ID application/client ID that is allowed for PnP interactive login
      - Permissions to read recycle bin metadata for each target site

    Safety:
      - Report-only script.
      - No cleanup, delete, restore, move, or content modification operations are performed.
      - The script intentionally contains no recycle bin clearing command.
#>

[CmdletBinding()]
param(
    # Tenant admin URL is required for tenant context and auditability in logs.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    # PnP application/client ID used for interactive authentication.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    # Folder where timestamped CSV reports and the log file are written.
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    # Optional CSV containing a SiteUrl column. If both InputCsv and SiteUrl are omitted, the script prompts for an input CSV or site URL.
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsv,

    # Optional single SharePoint Online or OneDrive personal site URL. If both InputCsv and SiteUrl are omitted, the script prompts for an input CSV or site URL.
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SiteUrl,

    # Minimum size threshold for IsLargeFile and default detail export filtering.
    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 1.7976931348623157E+308)]
    [double]$MinSizeMB = 100,

    # Include second-stage recycle bin items when the installed PnP.PowerShell command supports it.
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSecondStage,

    # Export every recycle bin item. Without this switch, only items >= MinSizeMB are exported in detail.
    [Parameter(Mandatory = $false)]
    [switch]$ExportAll
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------
# Script-level output paths
# -----------------------------
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$resolvedOutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
New-Item -Path $resolvedOutputFolder -ItemType Directory -Force | Out-Null

$detailPath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-detail-$timestamp.csv"
$summaryBySitePath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-summary-by-site-$timestamp.csv"
$summaryByUserPath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-summary-by-user-$timestamp.csv"
$summaryByExtensionPath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-summary-by-extension-$timestamp.csv"
$summaryByStagePath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-summary-by-stage-$timestamp.csv"
$summaryByLocationPath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-summary-by-location-$timestamp.csv"
$logPath = Join-Path -Path $resolvedOutputFolder -ChildPath "recyclebin-analysis-$timestamp.log"

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped line to the console and the run log.
    #>
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

function Get-ObjectPropertyValue {
    <#
    .SYNOPSIS
        Safely reads the first matching property from an object.

    .DESCRIPTION
        PnP.PowerShell object shapes can vary across versions. This helper lets the script
        read common property names without failing when a property is absent.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames
    )

    if ($null -eq $InputObject) {
        return $null
    }

    foreach ($propertyName in $PropertyNames) {
        $property = $InputObject.PSObject.Properties[$propertyName]
        if ($null -ne $property) {
            return $property.Value
        }
    }

    return $null
}

function Convert-ToInt64Safe {
    <#
    .SYNOPSIS
        Converts numeric-like values to Int64, returning zero when empty or not convertible.
    #>
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

function Get-FileExtensionFromLeafName {
    <#
    .SYNOPSIS
        Returns a lowercase extension, including the leading dot, from LeafName.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$LeafName
    )

    if ([string]::IsNullOrWhiteSpace($LeafName)) {
        return ''
    }

    $extension = [System.IO.Path]::GetExtension($LeafName)
    if ([string]::IsNullOrWhiteSpace($extension)) {
        return ''
    }

    return $extension.ToLowerInvariant()
}

function Get-LocationType {
    <#
    .SYNOPSIS
        Classifies a site URL as OneDrive or SharePoint.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentSiteUrl
    )

    if ($CurrentSiteUrl -like '*-my.sharepoint.com/personal/*') {
        return 'OneDrive'
    }

    return 'SharePoint'
}

function New-DetailRecord {
    <#
    .SYNOPSIS
        Creates one normalized detail row for a recycle bin item or site-level error.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentSiteUrl,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$RecycleBinItem,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$StageFallback,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ErrorMessage
    )

    $leafName = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('LeafName', 'FileName', 'Name'))
    $sizeBytes = Convert-ToInt64Safe -Value (Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('Size', 'SizeInBytes', 'Length'))
    $sizeMB = [Math]::Round(($sizeBytes / 1MB), 4)
    $sizeGB = [Math]::Round(($sizeBytes / 1GB), 4)
    $stage = Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('Stage', 'RecycleBinStage')

    if ($null -eq $stage -or [string]::IsNullOrWhiteSpace([string]$stage)) {
        $stage = $StageFallback
    }

    [pscustomobject]@{
        SiteUrl        = $CurrentSiteUrl
        LocationType   = Get-LocationType -CurrentSiteUrl $CurrentSiteUrl
        Title          = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('Title'))
        LeafName       = $leafName
        DirName        = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('DirName', 'DirectoryName'))
        ItemType       = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('ItemType', 'Type'))
        FileExtension  = Get-FileExtensionFromLeafName -LeafName $leafName
        SizeBytes      = [int64]$sizeBytes
        SizeMB         = [double]$sizeMB
        SizeGB         = [double]$sizeGB
        IsLargeFile    = [bool]($sizeMB -ge $MinSizeMB)
        DeletedByName  = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('DeletedByName', 'DeletedBy'))
        DeletedByEmail = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('DeletedByEmail', 'DeletedByEmailAddress'))
        DeletedDate    = Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('DeletedDate', 'DeletedTime')
        Stage          = [string]$stage
        Id             = [string](Get-ObjectPropertyValue -InputObject $RecycleBinItem -PropertyNames @('Id', 'UniqueId'))
        Error          = [string]$ErrorMessage
    }
}

function Get-RecycleBinItemsForSite {
    <#
    .SYNOPSIS
        Retrieves first-stage and optionally second-stage recycle bin items for one site.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentSiteUrl
    )

    Write-Log -Message "Connecting to site: $CurrentSiteUrl"

    # Connect to each site individually as requested. -ReturnConnection lets downstream PnP
    # calls use the intended connection explicitly and avoids accidental cross-site reuse.
    $connection = Connect-PnPOnline -Url $CurrentSiteUrl -Interactive -ClientId $ClientId -ReturnConnection

    $records = New-Object System.Collections.Generic.List[object]

    Write-Log -Message "Retrieving first-stage recycle bin items for site: $CurrentSiteUrl"
    $firstStageItems = @(Get-PnPRecycleBinItem -Connection $connection)
    foreach ($item in $firstStageItems) {
        $records.Add((New-DetailRecord -CurrentSiteUrl $CurrentSiteUrl -RecycleBinItem $item -StageFallback 'FirstStage' -ErrorMessage $null))
    }

    if ($IncludeSecondStage.IsPresent) {
        $command = Get-Command -Name Get-PnPRecycleBinItem -ErrorAction Stop
        $supportsSecondStageParameter = $command.Parameters.ContainsKey('SecondStage')

        if ($supportsSecondStageParameter) {
            try {
                Write-Log -Message "Retrieving second-stage recycle bin items for site: $CurrentSiteUrl"
                $secondStageItems = @(Get-PnPRecycleBinItem -Connection $connection -SecondStage)
                foreach ($item in $secondStageItems) {
                    $records.Add((New-DetailRecord -CurrentSiteUrl $CurrentSiteUrl -RecycleBinItem $item -StageFallback 'SecondStage' -ErrorMessage $null))
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                Write-Log -Level 'ERROR' -Message "Failed retrieving second-stage recycle bin items for $CurrentSiteUrl. Error: $errorMessage"
                $records.Add((New-DetailRecord -CurrentSiteUrl $CurrentSiteUrl -RecycleBinItem $null -StageFallback 'SecondStage' -ErrorMessage $errorMessage))
            }
        }
        else {
            Write-Log -Level 'WARN' -Message "The installed PnP.PowerShell Get-PnPRecycleBinItem command does not expose a -SecondStage parameter. Only returned/default recycle bin items were captured for $CurrentSiteUrl."
        }
    }

    return $records
}

function Get-SiteUrlsFromInput {
    <#
    .SYNOPSIS
        Builds the de-duplicated site URL list from InputCsv and/or SiteUrl.

    .DESCRIPTION
        InputCsv and SiteUrl are intentionally optional parameters so the script can be
        started interactively. When both are omitted, the script prompts for a CSV path
        first and then prompts for one site URL if no CSV path is supplied.
    #>
    $siteUrls = New-Object System.Collections.Generic.List[string]
    $effectiveInputCsv = $InputCsv
    $effectiveSiteUrl = $SiteUrl

    if ([string]::IsNullOrWhiteSpace($effectiveInputCsv) -and [string]::IsNullOrWhiteSpace($effectiveSiteUrl)) {
        Write-Log -Message 'No -InputCsv or -SiteUrl was supplied. Prompting for target input.'
        $effectiveInputCsv = (Read-Host -Prompt 'Enter the full path to an input CSV with a SiteUrl column, or press Enter to provide a single SiteUrl').Trim().Trim([char]34).Trim([char]39)

        if ([string]::IsNullOrWhiteSpace($effectiveInputCsv)) {
            $effectiveSiteUrl = (Read-Host -Prompt 'Enter one SharePoint Online or OneDrive SiteUrl').Trim()
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($effectiveInputCsv)) {
        if (-not (Test-Path -Path $effectiveInputCsv -PathType Leaf)) {
            throw "InputCsv was specified but the file was not found: $effectiveInputCsv"
        }

        $rows = @(Import-Csv -Path $effectiveInputCsv)
        if ($rows.Count -gt 0 -and $null -eq $rows[0].PSObject.Properties['SiteUrl']) {
            throw "InputCsv must contain a column named 'SiteUrl'. File: $effectiveInputCsv"
        }

        foreach ($row in $rows) {
            if (-not [string]::IsNullOrWhiteSpace($row.SiteUrl)) {
                $siteUrls.Add($row.SiteUrl.Trim())
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($effectiveSiteUrl)) {
        $siteUrls.Add($effectiveSiteUrl.Trim())
    }

    $uniqueSiteUrls = @($siteUrls | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($uniqueSiteUrls.Count -eq 0) {
        throw "Provide at least one site URL by using -InputCsv with a SiteUrl column, entering an input CSV when prompted, using -SiteUrl, or entering a SiteUrl when prompted."
    }

    return $uniqueSiteUrls
}

function Export-CsvUtf8 {
    <#
    .SYNOPSIS
        Exports objects as UTF-8 CSV without type information.

    .DESCRIPTION
        Export-Csv does not create headers when there are no objects. The HeaderColumns
        parameter ensures every requested report file is created with a stable schema,
        even when a run finds no matching recycle bin items.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string[]]$HeaderColumns
    )

    if ($InputObject.Count -gt 0) {
        $InputObject | Select-Object -Property $HeaderColumns | Export-Csv -Path $Path -NoTypeInformation -Encoding utf8
        return
    }

    $headerLine = ($HeaderColumns | ForEach-Object { '"{0}"' -f ($_ -replace '"', '""') }) -join ','
    Set-Content -Path $Path -Value $headerLine -Encoding utf8
}

function Get-SumInt64 {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Rows,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $sum = [int64]0
    foreach ($row in $Rows) {
        $sum += [int64]$row.$PropertyName
    }

    return $sum
}

function Get-SummaryRows {
    <#
    .SYNOPSIS
        Creates all requested summary datasets from successful recycle bin detail rows.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Rows
    )

    $successfulRows = @($Rows | Where-Object { [string]::IsNullOrWhiteSpace($_.Error) })

    $bySite = @(
        $successfulRows |
            Group-Object -Property SiteUrl, LocationType |
            ForEach-Object {
                $groupRows = @($_.Group)
                $totalBytes = Get-SumInt64 -Rows $groupRows -PropertyName 'SizeBytes'
                [pscustomobject]@{
                    SiteUrl        = [string]$groupRows[0].SiteUrl
                    LocationType   = [string]$groupRows[0].LocationType
                    ItemCount      = [int]$groupRows.Count
                    LargeFileCount = [int]@($groupRows | Where-Object { $_.IsLargeFile }).Count
                    TotalSizeBytes = [int64]$totalBytes
                    TotalSizeMB    = [double][Math]::Round(($totalBytes / 1MB), 4)
                    TotalSizeGB    = [double][Math]::Round(($totalBytes / 1GB), 4)
                }
            } |
            Sort-Object -Property TotalSizeBytes -Descending
    )

    $byUser = @(
        $successfulRows |
            Group-Object -Property DeletedByName, DeletedByEmail |
            ForEach-Object {
                $groupRows = @($_.Group)
                $totalBytes = Get-SumInt64 -Rows $groupRows -PropertyName 'SizeBytes'
                [pscustomobject]@{
                    DeletedByName  = [string]$groupRows[0].DeletedByName
                    DeletedByEmail = [string]$groupRows[0].DeletedByEmail
                    ItemCount      = [int]$groupRows.Count
                    LargeFileCount = [int]@($groupRows | Where-Object { $_.IsLargeFile }).Count
                    TotalSizeGB    = [double][Math]::Round(($totalBytes / 1GB), 4)
                }
            } |
            Sort-Object -Property TotalSizeGB -Descending
    )

    $byExtension = @(
        $successfulRows |
            Group-Object -Property FileExtension |
            ForEach-Object {
                $groupRows = @($_.Group)
                $totalBytes = Get-SumInt64 -Rows $groupRows -PropertyName 'SizeBytes'
                [pscustomobject]@{
                    FileExtension = [string]$_.Name
                    ItemCount     = [int]$groupRows.Count
                    TotalSizeGB   = [double][Math]::Round(($totalBytes / 1GB), 4)
                }
            } |
            Sort-Object -Property TotalSizeGB -Descending
    )

    $byStage = @(
        $successfulRows |
            Group-Object -Property Stage |
            ForEach-Object {
                $groupRows = @($_.Group)
                $totalBytes = Get-SumInt64 -Rows $groupRows -PropertyName 'SizeBytes'
                [pscustomobject]@{
                    Stage       = [string]$_.Name
                    ItemCount   = [int]$groupRows.Count
                    TotalSizeGB = [double][Math]::Round(($totalBytes / 1GB), 4)
                }
            } |
            Sort-Object -Property TotalSizeGB -Descending
    )

    $byLocation = @(
        $successfulRows |
            Group-Object -Property LocationType |
            ForEach-Object {
                $groupRows = @($_.Group)
                $totalBytes = Get-SumInt64 -Rows $groupRows -PropertyName 'SizeBytes'
                [pscustomobject]@{
                    LocationType = [string]$_.Name
                    SiteCount    = [int]@($groupRows | Select-Object -ExpandProperty SiteUrl -Unique).Count
                    ItemCount    = [int]$groupRows.Count
                    TotalSizeGB  = [double][Math]::Round(($totalBytes / 1GB), 4)
                }
            } |
            Sort-Object -Property TotalSizeGB -Descending
    )

    return [pscustomobject]@{
        BySite      = $bySite
        ByUser      = $byUser
        ByExtension = $byExtension
        ByStage     = $byStage
        ByLocation  = $byLocation
    }
}

# -----------------------------
# Main execution
# -----------------------------
$scriptStartTime = Get-Date
Write-Log -Message "Recycle bin storage analysis started."
Write-Log -Message "TenantAdminUrl: $TenantAdminUrl"
Write-Log -Message "OutputFolder: $resolvedOutputFolder"
Write-Log -Message "MinSizeMB: $MinSizeMB"
Write-Log -Message "IncludeSecondStage: $($IncludeSecondStage.IsPresent)"
Write-Log -Message "ExportAll: $($ExportAll.IsPresent)"
Write-Log -Message "Safety mode: report-only; no content changes will be made."

Import-Module PnP.PowerShell -ErrorAction Stop

$allRows = New-Object System.Collections.Generic.List[object]
$targetSiteUrls = @(Get-SiteUrlsFromInput)
$totalSites = $targetSiteUrls.Count
Write-Log -Message "Total site URLs to process: $totalSites"

for ($index = 0; $index -lt $totalSites; $index++) {
    $currentSiteUrl = $targetSiteUrls[$index]
    $siteNumber = $index + 1
    $percentComplete = [int](($siteNumber / $totalSites) * 100)

    $progressStatus = 'Processing {0} of {1}: {2}' -f $siteNumber, $totalSites, $currentSiteUrl
    Write-Progress -Activity 'Analyzing SharePoint Online / OneDrive recycle bins' -Status $progressStatus -PercentComplete $percentComplete
    Write-Log -Message $progressStatus

    try {
        $siteRows = @(Get-RecycleBinItemsForSite -CurrentSiteUrl $currentSiteUrl)
        foreach ($row in $siteRows) {
            $allRows.Add($row)
        }

        $siteTotalBytes = Get-SumInt64 -Rows $siteRows -PropertyName 'SizeBytes'
        Write-Log -Message "Completed site: $currentSiteUrl"
        Write-Log -Message "Item count for site: $($siteRows.Count)"
        Write-Log -Message "Total recycle bin size for site in bytes: $siteTotalBytes"
        Write-Log -Message "Total recycle bin size for site in GB: $([Math]::Round(($siteTotalBytes / 1GB), 4))"
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Log -Level 'ERROR' -Message "Failed processing site $currentSiteUrl. Error: $errorMessage"
        $allRows.Add((New-DetailRecord -CurrentSiteUrl $currentSiteUrl -RecycleBinItem $null -StageFallback '' -ErrorMessage $errorMessage))
        continue
    }
}

Write-Progress -Activity 'Analyzing SharePoint Online / OneDrive recycle bins' -Completed

$allRowsArray = @($allRows)
if ($ExportAll.IsPresent) {
    $detailRowsToExport = $allRowsArray
}
else {
    # Keep error rows so failed sites are visible in the detailed report even when filtering
    # successful detail rows down to large files only.
    $detailRowsToExport = @($allRowsArray | Where-Object { $_.IsLargeFile -or -not [string]::IsNullOrWhiteSpace($_.Error) })
}

$summaries = Get-SummaryRows -Rows $allRowsArray

$detailColumns = @('SiteUrl', 'LocationType', 'Title', 'LeafName', 'DirName', 'ItemType', 'FileExtension', 'SizeBytes', 'SizeMB', 'SizeGB', 'IsLargeFile', 'DeletedByName', 'DeletedByEmail', 'DeletedDate', 'Stage', 'Id', 'Error')
$summaryBySiteColumns = @('SiteUrl', 'LocationType', 'ItemCount', 'LargeFileCount', 'TotalSizeBytes', 'TotalSizeMB', 'TotalSizeGB')
$summaryByUserColumns = @('DeletedByName', 'DeletedByEmail', 'ItemCount', 'LargeFileCount', 'TotalSizeGB')
$summaryByExtensionColumns = @('FileExtension', 'ItemCount', 'TotalSizeGB')
$summaryByStageColumns = @('Stage', 'ItemCount', 'TotalSizeGB')
$summaryByLocationColumns = @('LocationType', 'SiteCount', 'ItemCount', 'TotalSizeGB')

Export-CsvUtf8 -InputObject $detailRowsToExport -Path $detailPath -HeaderColumns $detailColumns
Export-CsvUtf8 -InputObject $summaries.BySite -Path $summaryBySitePath -HeaderColumns $summaryBySiteColumns
Export-CsvUtf8 -InputObject $summaries.ByUser -Path $summaryByUserPath -HeaderColumns $summaryByUserColumns
Export-CsvUtf8 -InputObject $summaries.ByExtension -Path $summaryByExtensionPath -HeaderColumns $summaryByExtensionColumns
Export-CsvUtf8 -InputObject $summaries.ByStage -Path $summaryByStagePath -HeaderColumns $summaryByStageColumns
Export-CsvUtf8 -InputObject $summaries.ByLocation -Path $summaryByLocationPath -HeaderColumns $summaryByLocationColumns

$scriptEndTime = Get-Date
Write-Log -Message "Detail CSV: $detailPath"
Write-Log -Message "Summary by site CSV: $summaryBySitePath"
Write-Log -Message "Summary by user CSV: $summaryByUserPath"
Write-Log -Message "Summary by extension CSV: $summaryByExtensionPath"
Write-Log -Message "Summary by stage CSV: $summaryByStagePath"
Write-Log -Message "Summary by location CSV: $summaryByLocationPath"
Write-Log -Message "Run log: $logPath"
Write-Log -Message "Recycle bin storage analysis ended."
Write-Log -Message "Start time: $($scriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log -Message "End time: $($scriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log -Message "Elapsed time: $($scriptEndTime - $scriptStartTime)"
