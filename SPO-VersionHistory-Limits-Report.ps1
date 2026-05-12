# This script is read-only. It only reads SharePoint Online / OneDrive site version history limit settings and exports reports. It does not change version history settings, trim versions, delete files, modify libraries, change permissions, or update any site.

<#
.SYNOPSIS
    Reports SharePoint Online and OneDrive site-level version history limit settings.

.DESCRIPTION
    Reads a CSV containing SharePoint Online site URLs and OneDrive personal site URLs,
    connects to the SharePoint Online Admin Center with Connect-SPOService, reads current
    site-level version history limit properties with Get-SPOSite, and exports the results
    to Excel when ImportExcel is available or to CSV files as a fallback.

    This script is report-only. It does not change version history settings, trim versions,
    delete files, modify libraries, change permissions, or update any site.

.NOTES
    Requirements:
      - Windows PowerShell 5.1 or PowerShell 7+
      - Microsoft.Online.SharePoint.PowerShell module
      - ImportExcel module for .xlsx export, optional
      - SharePoint Online permissions sufficient to read the target sites and tenant settings
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsv,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputFolder,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputExcelFile,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeTenantDefaults,

    [Parameter(Mandatory = $false)]
    [switch]$ExportCsvFallback
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptStartTime = Get-Date
$timestamp = $scriptStartTime.ToString('yyyyMMdd-HHmmss')
$resolvedOutputFolder = [System.IO.Path]::GetFullPath($OutputFolder)
New-Item -Path $resolvedOutputFolder -ItemType Directory -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($OutputExcelFile)) {
    $excelPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-VersionHistory-Limits-Report-$timestamp.xlsx"
}
elseif ([System.IO.Path]::IsPathRooted($OutputExcelFile)) {
    $excelPath = $OutputExcelFile
}
else {
    $excelPath = Join-Path -Path $resolvedOutputFolder -ChildPath $OutputExcelFile
}

if ([System.IO.Path]::GetExtension($excelPath) -ne '.xlsx') {
    $excelPath = [System.IO.Path]::ChangeExtension($excelPath, '.xlsx')
}

$siteCsvPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-VersionHistory-Limits-Report-$timestamp.csv"
$tenantCsvPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-VersionHistory-TenantDefaults-$timestamp.csv"
$errorCsvPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-VersionHistory-Errors-$timestamp.csv"
$logPath = Join-Path -Path $resolvedOutputFolder -ChildPath "SPO-VersionHistory-Limits-Report-$timestamp.log"

$siteRows = New-Object 'System.Collections.Generic.List[object]'
$tenantRows = New-Object 'System.Collections.Generic.List[object]'
$errorRows = New-Object 'System.Collections.Generic.List[object]'

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $logPath -Value $line -Encoding UTF8

    switch ($Level) {
        'ERROR' { Write-Error -Message $Message -ErrorAction Continue }
        'WARN'  { Write-Warning -Message $Message }
        default { Write-Host $line }
    }
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -ne $property) {
        return $property.Value
    }

    return $DefaultValue
}

function Convert-ToGigabytes {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Megabytes
    )

    if ($null -eq $Megabytes -or [string]::IsNullOrWhiteSpace([string]$Megabytes)) {
        return $null
    }

    $numericValue = 0.0
    if ([double]::TryParse([string]$Megabytes, [ref]$numericValue)) {
        return [Math]::Round(($numericValue / 1024), 2)
    }

    return $null
}

function Get-VersionPolicyMode {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$EnableAutoExpirationVersionTrim
    )

    if ($null -eq $EnableAutoExpirationVersionTrim -or [string]::IsNullOrWhiteSpace([string]$EnableAutoExpirationVersionTrim)) {
        return 'Inherit/NotSet'
    }

    $valueText = [string]$EnableAutoExpirationVersionTrim
    if ($valueText -eq 'True') {
        return 'Automatic'
    }

    if ($valueText -eq 'False') {
        return 'Manual'
    }

    return 'Inherit/NotSet'
}

function Get-EffectivePolicyNote {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('Automatic', 'Manual', 'Inherit/NotSet')]
        [string]$VersionPolicyMode
    )

    switch ($VersionPolicyMode) {
        'Automatic' { return 'Site-level automatic version history limits are configured.' }
        'Manual' { return 'Site-level manual version history limits are configured.' }
        default { return 'No site-level override detected. The site may inherit organization-level defaults.' }
    }
}

function Add-ReportError {
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$SiteUrl,

        [Parameter(Mandatory = $true)]
        [string]$Stage,

        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage
    )

    $errorRows.Add([pscustomobject]@{
        Timestamp    = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        SiteUrl      = $SiteUrl
        Stage        = $Stage
        ErrorMessage = $ErrorMessage
    })

    Write-Log -Level 'ERROR' -Message "$Stage failed for '$SiteUrl'. $ErrorMessage"
}

function Get-SiteVersionLimitRow {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl
    )

    $spoSite = Get-SPOSite -Identity $SiteUrl

    $url = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'Url'
    $owner = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'Owner'
    $template = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'Template'
    $storageUsageCurrent = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'StorageUsageCurrent'
    $storageQuota = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'StorageQuota'
    $storageQuotaWarningLevel = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'StorageQuotaWarningLevel'
    $lastContentModifiedDate = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'LastContentModifiedDate'
    $enableAutoExpirationVersionTrim = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'EnableAutoExpirationVersionTrim'
    $expireVersionsAfterDays = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'ExpireVersionsAfterDays'
    $majorVersionLimit = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'MajorVersionLimit'
    $majorWithMinorVersionsLimit = Get-ObjectPropertyValue -InputObject $spoSite -PropertyName 'MajorWithMinorVersionsLimit' -DefaultValue 'NotAvailable'

    $locationType = if ($SiteUrl -like '*-my.sharepoint.com/personal/*') { 'OneDrive' } else { 'SharePoint' }
    $versionPolicyMode = Get-VersionPolicyMode -EnableAutoExpirationVersionTrim $enableAutoExpirationVersionTrim

    return [pscustomobject]@{
        Url                             = $url
        Owner                           = $owner
        Template                        = $template
        StorageUsageCurrent             = $storageUsageCurrent
        StorageQuota                    = $storageQuota
        StorageQuotaWarningLevel        = $storageQuotaWarningLevel
        LastContentModifiedDate         = $lastContentModifiedDate
        EnableAutoExpirationVersionTrim = $enableAutoExpirationVersionTrim
        ExpireVersionsAfterDays         = $expireVersionsAfterDays
        MajorVersionLimit               = $majorVersionLimit
        MajorWithMinorVersionsLimit     = $majorWithMinorVersionsLimit
        LocationType                    = $locationType
        StorageUsedGB                   = Convert-ToGigabytes -Megabytes $storageUsageCurrent
        StorageQuotaGB                  = Convert-ToGigabytes -Megabytes $storageQuota
        StorageWarningGB                = Convert-ToGigabytes -Megabytes $storageQuotaWarningLevel
        VersionPolicyMode               = $versionPolicyMode
        EffectivePolicyNote             = Get-EffectivePolicyNote -VersionPolicyMode $versionPolicyMode
        Error                           = $null
    }
}

function Get-TenantDefaultsRow {
    $tenant = Get-SPOTenant

    $enableAutoExpirationVersionTrim = Get-ObjectPropertyValue -InputObject $tenant -PropertyName 'EnableAutoExpirationVersionTrim' -DefaultValue 'NotAvailable'
    $expireVersionsAfterDays = Get-ObjectPropertyValue -InputObject $tenant -PropertyName 'ExpireVersionsAfterDays' -DefaultValue 'NotAvailable'
    $majorVersionLimit = Get-ObjectPropertyValue -InputObject $tenant -PropertyName 'MajorVersionLimit' -DefaultValue 'NotAvailable'
    $versionPolicyMode = Get-VersionPolicyMode -EnableAutoExpirationVersionTrim $enableAutoExpirationVersionTrim

    return [pscustomobject]@{
        TenantAdminUrl                  = $TenantAdminUrl
        EnableAutoExpirationVersionTrim = $enableAutoExpirationVersionTrim
        ExpireVersionsAfterDays         = $expireVersionsAfterDays
        MajorVersionLimit               = $majorVersionLimit
        VersionPolicyMode               = $versionPolicyMode
        EffectivePolicyNote             = Get-EffectivePolicyNote -VersionPolicyMode $versionPolicyMode
        RetrievedAt                     = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    }
}

function Export-ReportData {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$UseCsv
    )

    if ($UseCsv) {
        $siteRows | Export-Csv -Path $siteCsvPath -NoTypeInformation -Encoding UTF8
        $errorRows | Export-Csv -Path $errorCsvPath -NoTypeInformation -Encoding UTF8

        if ($IncludeTenantDefaults.IsPresent) {
            $tenantRows | Export-Csv -Path $tenantCsvPath -NoTypeInformation -Encoding UTF8
        }

        return @($siteCsvPath, $(if ($IncludeTenantDefaults.IsPresent) { $tenantCsvPath }), $errorCsvPath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }

    if (Test-Path -Path $excelPath) {
        throw "Output Excel file '$excelPath' already exists. Provide a different OutputExcelFile value or remove the existing local report file before running again."
    }

    $siteRows | Export-Excel -Path $excelPath -WorksheetName 'SiteVersionLimits' -AutoSize -TableName 'SiteVersionLimits' -FreezeTopRow

    if ($IncludeTenantDefaults.IsPresent) {
        $tenantRows | Export-Excel -Path $excelPath -WorksheetName 'TenantDefaults' -AutoSize -TableName 'TenantDefaults' -FreezeTopRow -Append
    }

    $errorRows | Export-Excel -Path $excelPath -WorksheetName 'Errors' -AutoSize -TableName 'Errors' -FreezeTopRow -Append

    return @($excelPath)
}

Write-Log -Message 'Starting SharePoint Online / OneDrive version history limits report.'
Write-Log -Message "Start time: $($scriptStartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log -Message "TenantAdminUrl: $TenantAdminUrl"
Write-Log -Message "Output folder: $resolvedOutputFolder"

if (-not (Test-Path -Path $InputCsv -PathType Leaf)) {
    throw "Input CSV '$InputCsv' does not exist."
}

$inputRows = Import-Csv -Path $InputCsv
if ($null -eq $inputRows) {
    throw "Input CSV '$InputCsv' is empty."
}

$firstRow = $inputRows | Select-Object -First 1
if ($null -eq $firstRow.PSObject.Properties['SiteUrl']) {
    throw "Input CSV '$InputCsv' must contain a column named SiteUrl."
}

$siteUrls = @($inputRows | ForEach-Object { $_.SiteUrl } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() })
if ($siteUrls.Count -eq 0) {
    throw "Input CSV '$InputCsv' does not contain any non-empty SiteUrl values."
}

Write-Log -Message "Number of SiteUrls provided: $($siteUrls.Count)"

if ($null -eq (Get-Command -Name Connect-SPOService -ErrorAction SilentlyContinue)) {
    throw 'Connect-SPOService was not found. Install or load the Microsoft.Online.SharePoint.PowerShell module, then run this script again.'
}

if ($null -eq (Get-Command -Name Get-SPOSite -ErrorAction SilentlyContinue)) {
    throw 'Get-SPOSite was not found. Install or load the Microsoft.Online.SharePoint.PowerShell module, then run this script again.'
}

Write-Log -Message 'Connecting to SharePoint Online Admin Center. Complete the interactive sign-in prompt if the module requires it.'
Connect-SPOService -Url $TenantAdminUrl
Write-Log -Message 'Connected to SharePoint Online Admin Center.'

foreach ($siteUrl in $siteUrls) {
    Write-Log -Message "Checking site: $siteUrl"

    try {
        $row = Get-SiteVersionLimitRow -SiteUrl $siteUrl
        $siteRows.Add($row)
    }
    catch {
        $message = $_.Exception.Message
        Add-ReportError -SiteUrl $siteUrl -Stage 'Get-SPOSite' -ErrorMessage $message

        $locationType = if ($siteUrl -like '*-my.sharepoint.com/personal/*') { 'OneDrive' } else { 'SharePoint' }
        $siteRows.Add([pscustomobject]@{
            Url                             = $siteUrl
            Owner                           = $null
            Template                        = $null
            StorageUsageCurrent             = $null
            StorageQuota                    = $null
            StorageQuotaWarningLevel        = $null
            LastContentModifiedDate         = $null
            EnableAutoExpirationVersionTrim = $null
            ExpireVersionsAfterDays         = $null
            MajorVersionLimit               = $null
            MajorWithMinorVersionsLimit     = 'NotAvailable'
            LocationType                    = $locationType
            StorageUsedGB                   = $null
            StorageQuotaGB                  = $null
            StorageWarningGB                = $null
            VersionPolicyMode               = 'Inherit/NotSet'
            EffectivePolicyNote             = 'No site-level override detected. The site may inherit organization-level defaults.'
            Error                           = $message
        })
    }
}

if ($IncludeTenantDefaults.IsPresent) {
    if ($null -eq (Get-Command -Name Get-SPOTenant -ErrorAction SilentlyContinue)) {
        Add-ReportError -SiteUrl $null -Stage 'Get-SPOTenant' -ErrorMessage 'Get-SPOTenant was not found in the current session.'
    }
    else {
        try {
            $tenantRows.Add((Get-TenantDefaultsRow))
        }
        catch {
            Add-ReportError -SiteUrl $null -Stage 'Get-SPOTenant' -ErrorMessage $_.Exception.Message
        }
    }
}

$importExcelAvailable = $null -ne (Get-Module -ListAvailable -Name ImportExcel)
$useCsv = $ExportCsvFallback.IsPresent -or (-not $importExcelAvailable)

if (-not $importExcelAvailable) {
    Write-Log -Level 'WARN' -Message 'ImportExcel module is not installed or not available. Falling back to CSV export.'
}
elseif ($ExportCsvFallback.IsPresent) {
    Write-Log -Level 'WARN' -Message 'ExportCsvFallback was specified. Exporting CSV files instead of Excel.'
}

$outputPaths = Export-ReportData -UseCsv $useCsv

$successCount = @($siteRows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.Error) }).Count
$failedCount = @($siteRows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Error) }).Count
$oneDriveCount = @($siteRows | Where-Object { $_.LocationType -eq 'OneDrive' }).Count
$sharePointCount = @($siteRows | Where-Object { $_.LocationType -eq 'SharePoint' }).Count
$automaticCount = @($siteRows | Where-Object { $_.VersionPolicyMode -eq 'Automatic' }).Count
$manualCount = @($siteRows | Where-Object { $_.VersionPolicyMode -eq 'Manual' }).Count
$inheritCount = @($siteRows | Where-Object { $_.VersionPolicyMode -eq 'Inherit/NotSet' }).Count
$scriptEndTime = Get-Date

Write-Log -Message "Number of successful checks: $successCount"
Write-Log -Message "Number of failed checks: $failedCount"
Write-Log -Message "Output path: $($outputPaths -join '; ')"
Write-Log -Message "End time: $($scriptEndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Write-Log -Message 'Completed SharePoint Online / OneDrive version history limits report.'

Write-Host ''
Write-Host 'Summary'
Write-Host '-------'
Write-Host ("Total sites checked: {0}" -f $siteRows.Count)
Write-Host ("OneDrive count: {0}" -f $oneDriveCount)
Write-Host ("SharePoint count: {0}" -f $sharePointCount)
Write-Host ("Automatic policy count: {0}" -f $automaticCount)
Write-Host ("Manual policy count: {0}" -f $manualCount)
Write-Host ("Inherit/NotSet count: {0}" -f $inheritCount)
Write-Host ("Error count: {0}" -f $failedCount)
Write-Host ("Output path: {0}" -f ($outputPaths -join '; '))
Write-Host ("Log path: {0}" -f $logPath)
Write-Host ''
Write-Host 'Blank version history fields are not treated as errors. They may indicate that no site-level override is configured and the site may inherit organization-level defaults.'
Write-Host 'This report is read-only and uses only SharePoint Online read/reporting commands plus local import/export commands.'
