#Requires -Version 5.1
#Requires -Modules PnP.PowerShell
<#
.SYNOPSIS
Limits version history to a specified number of major versions for PowerPoint documents 
across all SharePoint sites in the tenant.
.DESCRIPTION
This script scans all SharePoint sites, document libraries, and identifies PowerPoint files 
(.pptx, .ppt). It trims version history keeping only the specified number of recent versions. 
By default, older versions are sent to the recycle bin. Use -SkipRecycleBin to permanently delete.
.PARAMETER MajorVersionsToKeep
Number of major versions to keep for each PowerPoint file. Default: 5. Set to 3-5 to 
aggressively free storage.
.PARAMETER TenantAdminUrl
Your tenant admin URL (e.g., https://yourtenant-admin.sharepoint.com)
.PARAMETER ClientId
Optional. Client ID for app-only authentication. If not provided, uses interactive auth.
.PARAMETER CertificatePath
Optional. Path to certificate (.pfx) for app-only authentication.
.PARAMETER CertificatePassword
Optional. Password for the certificate file.
.PARAMETER SkipRecycleBin
If specified, permanently deletes versions (bypasses recycle bin). Use with caution.
.PARAMETER LogDirectory
Directory for log files. Defaults to script directory.
.PARAMETER ThrottleDelayMs
Delay in milliseconds between site connections to avoid throttling. Default: 500.
.PARAMETER IncludeSites
Array of site URLs to include. If specified, only these sites are processed.
.PARAMETER ExcludeSites
Array of site URL patterns to exclude (supports wildcards). E.g., "*HR*", "*Legal*"
.PARAMETER ReportOnly
If specified, only generates a report of files and estimated storage savings without deleting.
.NOTES
- Requires PnP PowerShell module. Install with: Install-Module -Name PnP.PowerShell -Force
- Requires Global Admin or SharePoint Admin credentials
- Recommendation: Run with -WhatIf first to validate approach
- Run during off-hours (takes time to process all sites)

IMPACT NOTES:
- Audit Logs: Generates "Delete" events in M365 Unified Audit Log
- Storage: Quota updates take 24-48 hours. Recycle bin items still count for 93 days.
- Recycle Bin: May flood with version entries if -SkipRecycleBin is not used
.EXAMPLE
.\Limit-PowerPointVersions.ps1 -MajorVersionsToKeep 5 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com" -WhatIf
Runs in preview mode showing what would be deleted without making changes.
.EXAMPLE
.\Limit-PowerPointVersions.ps1 -MajorVersionsToKeep 5 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com" -ReportOnly
Generates a CSV report of all files and potential savings without making any changes.
.EXAMPLE
.\Limit-PowerPointVersions.ps1 -MajorVersionsToKeep 3 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com" -ExcludeSites "*HR*","*Legal*" -Confirm:$false
Runs cleanup keeping 3 versions, excluding HR and Legal sites.
.EXAMPLE
.\Limit-PowerPointVersions.ps1 -MajorVersionsToKeep 5 -TenantAdminUrl "https://yourtenant-admin.sharepoint.com" -IncludeSites "https://contoso.sharepoint.com/sites/Marketing"
Runs only on the Marketing site for testing.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = "High")]
param(
    [Parameter()]
    [ValidateRange(1, 500)]
    [int]$MajorVersionsToKeep = 5,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantAdminUrl,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$CertificatePath,

    [Parameter()]
    [SecureString]$CertificatePassword,

    [Parameter()]
    [switch]$SkipRecycleBin,

    [Parameter()]
    [string]$LogDirectory,

    [Parameter()]
    [ValidateRange(0, 10000)]
    [int]$ThrottleDelayMs = 500,

    [Parameter()]
    [string[]]$IncludeSites,

    [Parameter()]
    [string[]]$ExcludeSites,

    [Parameter()]
    [switch]$ReportOnly
)

# ============================================================================
# CONFIGURATION
# ============================================================================
$script:ScriptStartTime = Get-Date
$script:PowerPointExtensions = @(".pptx", ".ppt")

# Set log directory
if ([string]::IsNullOrEmpty($LogDirectory)) {
    $LogDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
}
$script:LogPath = Join-Path $LogDirectory "PowerPoint-VersionCleanup-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$script:ReportPath = Join-Path $LogDirectory "PowerPoint-VersionCleanup-$(Get-Date -Format 'yyyyMMdd-HHmmss')-Report.csv"

# Initialize tracking variables (script scope instead of global)
$script:Stats = @{
    SitesProcessed          = 0
    SitesSkipped            = 0
    LibrariesProcessed      = 0
    FilesProcessed          = 0
    FilesWithExcessVersions = 0
    VersionsDeleted         = 0
    StorageFreedBytes       = 0
    ErrorCount              = 0
}

# Report data collection
$script:ReportData = [System.Collections.ArrayList]::new()

# Store authentication parameters for connection reuse
$script:AuthParams = @{}

# ============================================================================
# FUNCTIONS
# ============================================================================
function Format-FileSize {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [long]$Bytes
    )
    
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes bytes"
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS", "SUMMARY")]
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    switch ($Level) {
        "ERROR"   { Write-Host $LogEntry -ForegroundColor Red }
        "WARNING" { Write-Host $LogEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $LogEntry -ForegroundColor Green }
        "SUMMARY" { Write-Host $LogEntry -ForegroundColor Cyan }
        default   { Write-Host $LogEntry }
    }
    
    Add-Content -Path $script:LogPath -Value $LogEntry -ErrorAction SilentlyContinue
}

function Test-AppOnlyAuthParameters {
    [CmdletBinding()]
    param(
        [string]$ClientId,
        [string]$CertificatePath
    )
    
    # If either is provided, both must be provided
    $hasClientId = -not [string]::IsNullOrEmpty($ClientId)
    $hasCertPath = -not [string]::IsNullOrEmpty($CertificatePath)
    
    if ($hasClientId -xor $hasCertPath) {
        Write-Log "App-only authentication requires both -ClientId and -CertificatePath" -Level ERROR
        return $false
    }
    
    if ($hasCertPath -and -not (Test-Path $CertificatePath)) {
        Write-Log "Certificate file not found: $CertificatePath" -Level ERROR
        return $false
    }
    
    return $true
}

function Initialize-AuthParameters {
    [CmdletBinding()]
    param(
        [string]$ClientId,
        [string]$CertificatePath,
        [SecureString]$CertificatePassword
    )
    
    # Store auth params for reuse across connections
    if ($ClientId -and $CertificatePath) {
        $script:AuthParams['ClientId'] = $ClientId
        $script:AuthParams['CertificatePath'] = $CertificatePath
        if ($CertificatePassword) {
            $script:AuthParams['CertificatePassword'] = $CertificatePassword
        }
        $script:AuthParams['UseAppOnly'] = $true
    }
    else {
        $script:AuthParams['Interactive'] = $true
        $script:AuthParams['UseAppOnly'] = $false
    }
}

function Connect-ToSharePointSite {
    <#
    .SYNOPSIS
    Connects to a SharePoint site, reusing authentication context where possible.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter()]
        [switch]$IsTenantAdmin
    )
    
    try {
        $connectParams = @{
            Url = $SiteUrl
        }

        # Add authentication parameters
        if ($script:AuthParams.UseAppOnly) {
            $connectParams['ClientId'] = $script:AuthParams.ClientId
            $connectParams['CertificatePath'] = $script:AuthParams.CertificatePath
            if ($script:AuthParams.CertificatePassword) {
                $connectParams['CertificatePassword'] = $script:AuthParams.CertificatePassword
            }
        }
        else {
            $connectParams['Interactive'] = $true
        }
        
        Connect-PnPOnline @connectParams -ErrorAction Stop
        
        if ($IsTenantAdmin) {
            Write-Log "Successfully connected to tenant admin" -Level SUCCESS
        }
        
        return $true
    }
    catch {
        $errorMsg = if ($IsTenantAdmin) { "Failed to connect to tenant admin: $_" } else { "Failed to connect to site '$SiteUrl': $_" }
        Write-Log $errorMsg -Level ERROR
        if (-not $IsTenantAdmin) {
            $script:Stats.ErrorCount++
        }
        return $false
    }
}

function Test-SiteExcluded {
    <#
    .SYNOPSIS
    Checks if a site should be excluded based on include/exclude lists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl
    )
    
    # If IncludeSites is specified, only process those sites
    if ($IncludeSites -and $IncludeSites.Count -gt 0) {
        $isIncluded = $IncludeSites | Where-Object { $SiteUrl -eq $_ }
        if (-not $isIncluded) {
            return $true  # Excluded because not in include list
        }
    }
    
    # Check exclude patterns
    if ($ExcludeSites -and $ExcludeSites.Count -gt 0) {
        foreach ($pattern in $ExcludeSites) {
            if ($SiteUrl -like $pattern) {
                return $true  # Excluded by pattern
            }
        }
    }
    
    return $false
}

function Get-AllSharePointSites {
    [CmdletBinding()]
    param()
    
    try {
        Write-Log "Retrieving all SharePoint sites..." -Level INFO
        
        # Get all sites and filter out system sites
        $Sites = Get-PnPTenantSite -Detailed | Where-Object {
            $_.Template -notlike "*REDIRECT*" -and
            $_.Template -notlike "*SRCHCEN*" -and
            $_.Template -notlike "*SPSMSITEHOST*" -and
            $_.Template -notlike "*APPCATALOG*" -and
            $_.Template -notlike "*POINTPUBLISHINGHUB*" -and
            $_.Template -notlike "*EDISC*" -and
            $_.Url -notlike "*-my.sharepoint.com*"
        }
        
        Write-Log "Found $($Sites.Count) sites (before include/exclude filtering)" -Level INFO
        return $Sites
    }
    catch {
        Write-Log "Failed to retrieve sites: $_" -Level ERROR
        return @()
    }
}

function Get-PowerPointFilesWithCAML {
    <#
    .SYNOPSIS
    Retrieves PowerPoint files from a library using server-side CAML filtering.
    This is much more efficient than client-side filtering.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ListName
    )
    
    $PowerPointFiles = [System.Collections.ArrayList]::new()
    
    try {
        # CAML query to filter for PowerPoint files on the server side
        # This dramatically reduces network traffic and memory usage
        $camlQuery = @"
<View Scope='RecursiveAll'>
    <Query>
        <Where>
            <Or>
                <Eq>
                    <FieldRef Name='File_x0020_Type'/>
                    <Value Type='Text'>pptx</Value>
                </Eq>
                <Eq>
                    <FieldRef Name='File_x0020_Type'/>
                    <Value Type='Text'>ppt</Value>
                </Eq>
            </Or>
        </Where>
    </Query>
    <RowLimit Paged='TRUE'>2000</RowLimit>
</View>
"@
        
        $AllItems = Get-PnPListItem -List $ListName -Query $camlQuery -PageSize 2000 -ErrorAction Stop
        
        foreach ($Item in $AllItems) {
            $FileName = $Item.FieldValues["FileLeafRef"]
            if ($FileName) {
                [void]$PowerPointFiles.Add($Item)
            }
        }
        
        return $PowerPointFiles
    }
    catch {
        # If CAML query fails (e.g., indexed column issue), fall back to client-side filtering
        Write-Log "CAML query failed for '$ListName', falling back to client-side filtering: $_" -Level WARNING
        return Get-PowerPointFilesClientSide -ListName $ListName
    }
}

function Get-PowerPointFilesClientSide {
    <#
    .SYNOPSIS
    Fallback: Retrieves PowerPoint files using client-side filtering.
    Used when CAML query fails (e.g., large list threshold issues).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ListName
    )
    
    $PowerPointFiles = [System.Collections.ArrayList]::new()
    
    try {
        $AllItems = Get-PnPListItem -List $ListName -PageSize 2000 -ErrorAction Stop
        
        foreach ($Item in $AllItems) {
            $FileName = $Item.FieldValues["FileLeafRef"]
            if (-not $FileName) { continue }
            
            $extension = [System.IO.Path]::GetExtension($FileName).ToLower()
            if ($script:PowerPointExtensions -contains $extension) {
                [void]$PowerPointFiles.Add($Item)
            }
        }
        
        return $PowerPointFiles
    }
    catch {
        Write-Log "Error retrieving files from library '$ListName': $_" -Level WARNING
        return @()
    }
}

function Remove-OldFileVersions {
    <#
    .SYNOPSIS
    Removes old versions from a file, keeping only the specified number of newest versions.
    Sorts by Created date (most reliable) instead of parsing version labels.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [object]$FileItem,

        [Parameter(Mandatory)]
        [int]$KeepVersions,

        [Parameter()]
        [switch]$SkipRecycleBin,

        [Parameter()]
        [switch]$ReportOnly,

        [Parameter()]
        [string]$SiteUrl,

        [Parameter()]
        [string]$LibraryName
    )
    
    $FileName = $FileItem.FieldValues["FileLeafRef"]
    $FileUrl = $FileItem.FieldValues["FileRef"]
    
    try {
        # Get all versions for this file
        $Versions = Get-PnPFileVersion -Url $FileUrl -ErrorAction Stop
        
        if (-not $Versions -or $Versions.Count -eq 0) {
            Write-Verbose "File '$FileName' has no version history"
            return @{ Deleted = 0; Skipped = $true }
        }
        
        $VersionCount = $Versions.Count
        
        if ($VersionCount -le $KeepVersions) {
            Write-Verbose "File '$FileName' has $VersionCount versions (keeping $KeepVersions) - no action needed"
            return @{ Deleted = 0; Skipped = $true }
        }
        
        # CRITICAL: Sort by Created date DESCENDING (newest first)
        # This is more reliable than parsing version labels which can have
        # unexpected formats (drafts, minor versions, etc.)
        $SortedVersions = $Versions | Sort-Object -Property Created -Descending
        
        $VersionsToRemove = $SortedVersions | Select-Object -Skip $KeepVersions
        $VersionsToDeleteCount = ($VersionsToRemove | Measure-Object).Count
        
        if ($VersionsToDeleteCount -eq 0) {
            return @{ Deleted = 0; Skipped = $true }
        }
        
        $script:Stats.FilesWithExcessVersions++
        
        # Calculate total size that would be freed
        $SizeToFree = ($VersionsToRemove | ForEach-Object { if ($_.Size) { [long]$_.Size } else { 0 } } | Measure-Object -Sum).Sum
        
        # Add to report data
        [void]$script:ReportData.Add([PSCustomObject]@{
            SiteUrl           = $SiteUrl
            Library           = $LibraryName
            FileName          = $FileName
            FileUrl           = $FileUrl
            TotalVersions     = $VersionCount
            VersionsToDelete  = $VersionsToDeleteCount
            VersionsToKeep    = $KeepVersions
            EstimatedSizeBytes = $SizeToFree
            EstimatedSize     = Format-FileSize -Bytes $SizeToFree
        })
        
        # If ReportOnly, just log and return
        if ($ReportOnly) {
            Write-Log "File '$FileName': Would delete $VersionsToDeleteCount of $VersionCount versions - $(Format-FileSize $SizeToFree)" -Level INFO
            $script:Stats.VersionsDeleted += $VersionsToDeleteCount
            $script:Stats.StorageFreedBytes += $SizeToFree
            return @{ Deleted = $VersionsToDeleteCount; Skipped = $false; ReportOnly = $true }
        }
        
        $DeletedCount = 0
        $SizeFreed = 0
        $Action = if ($SkipRecycleBin) { "permanently delete" } else { "recycle" }
        
        foreach ($Version in $VersionsToRemove) {
            $VersionLabel = $Version.VersionLabel
            $VersionSize = if ($Version.Size) { [long]$Version.Size } else { 0 }
            
            if ($PSCmdlet.ShouldProcess("$FileName (version $VersionLabel, $(Format-FileSize $VersionSize))", "Delete version")) {
                try {
                    $removeParams = @{
                        Url      = $FileUrl
                        Identity = $Version.ID
                        Force    = $true
                    }
                    
                    # Use recycle bin by default (safer)
                    if (-not $SkipRecycleBin) {
                        $removeParams['Recycle'] = $true
                    }
                    
                    Remove-PnPFileVersion @removeParams
                    
                    $DeletedCount++
                    $SizeFreed += $VersionSize
                    $script:Stats.VersionsDeleted++
                    $script:Stats.StorageFreedBytes += $VersionSize
                }
                catch {
                    Write-Log "Error deleting version $VersionLabel from '$FileName': $_" -Level WARNING
                    $script:Stats.ErrorCount++
                }
            }
            else {
                # WhatIf mode - count as would-be deleted and track size
                $DeletedCount++
                $SizeFreed += $VersionSize
                $script:Stats.VersionsDeleted++
                $script:Stats.StorageFreedBytes += $VersionSize
            }
        }
        
        if ($DeletedCount -gt 0) {
            $SizeFormatted = Format-FileSize -Bytes $SizeFreed
            Write-Log "File '$FileName': $Action $DeletedCount of $VersionCount versions (kept newest $KeepVersions) - $SizeFormatted" -Level INFO
        }
        
        return @{ Deleted = $DeletedCount; Skipped = $false }
    }
    catch {
        Write-Log "Error processing versions for '$FileName': $_" -Level WARNING
        $script:Stats.ErrorCount++
        return @{ Deleted = 0; Skipped = $true; Error = $true }
    }
}

function Invoke-LibraryVersionCleanup {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [object]$Library,

        [Parameter(Mandatory)]
        [int]$KeepVersions,

        [Parameter()]
        [switch]$SkipRecycleBin,

        [Parameter()]
        [switch]$ReportOnly
    )
    
    $LibraryTitle = $Library.Title
    
    try {
        Write-Verbose "Processing library: $LibraryTitle"
        
        # Use CAML query for efficient server-side filtering
        $PowerPointFiles = Get-PowerPointFilesWithCAML -ListName $LibraryTitle
        
        if ($PowerPointFiles.Count -eq 0) {
            Write-Verbose "No PowerPoint files found in library: $LibraryTitle"
            return
        }
        
        Write-Log "Library '$LibraryTitle': Found $($PowerPointFiles.Count) PowerPoint files" -Level INFO
        
        $fileIndex = 0
        foreach ($File in $PowerPointFiles) {
            $fileIndex++
            $script:Stats.FilesProcessed++
            
            # Show progress for files within library
            $fileName = $File.FieldValues["FileLeafRef"]
            Write-Progress -Activity "Processing Library: $LibraryTitle" `
                -Status "File $fileIndex of $($PowerPointFiles.Count): $fileName" `
                -PercentComplete (($fileIndex / $PowerPointFiles.Count) * 100) `
                -Id 2 -ParentId 1
            
            $removeParams = @{
                FileItem       = $File
                KeepVersions   = $KeepVersions
                SkipRecycleBin = $SkipRecycleBin
                ReportOnly     = $ReportOnly
                SiteUrl        = $SiteUrl
                LibraryName    = $LibraryTitle
            }
            
            Remove-OldFileVersions @removeParams | Out-Null
        }
        
        Write-Progress -Activity "Processing Library: $LibraryTitle" -Completed -Id 2
        
        $script:Stats.LibrariesProcessed++
    }
    catch {
        Write-Log "Error processing library '$LibraryTitle': $_" -Level ERROR
        $script:Stats.ErrorCount++
    }
}

function Write-Summary {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$WhatIfMode,

        [Parameter()]
        [switch]$ReportOnly
    )
    
    $ExecutionTime = (Get-Date) - $script:ScriptStartTime
    $StorageFormatted = Format-FileSize -Bytes $script:Stats.StorageFreedBytes
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "         EXECUTION SUMMARY              " -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    
    if ($ReportOnly) {
        Write-Log "[REPORT-ONLY MODE] No changes were made - report generated" -Level SUMMARY
        Write-Host ""
    }
    elseif ($WhatIfMode) {
        Write-Log "[WHAT-IF MODE] No changes were made - preview only" -Level SUMMARY
        Write-Host ""
    }
    
    Write-Log "Sites Processed:                 $($script:Stats.SitesProcessed)" -Level SUMMARY
    Write-Log "Sites Skipped (excluded):        $($script:Stats.SitesSkipped)" -Level SUMMARY
    Write-Log "Libraries Processed:             $($script:Stats.LibrariesProcessed)" -Level SUMMARY
    Write-Log "PowerPoint Files Scanned:        $($script:Stats.FilesProcessed)" -Level SUMMARY
    Write-Log "Files with Excess Versions:      $($script:Stats.FilesWithExcessVersions)" -Level SUMMARY
    
    if ($ReportOnly -or $WhatIfMode) {
        Write-Log "Versions That WOULD Be Deleted:  $($script:Stats.VersionsDeleted)" -Level SUMMARY
        Write-Log "Estimated Storage Savings:       $StorageFormatted" -Level SUMMARY
    }
    else {
        Write-Log "Versions Deleted:                $($script:Stats.VersionsDeleted)" -Level SUMMARY
        Write-Log "Storage Freed:                   $StorageFormatted" -Level SUMMARY
    }
    
    Write-Log "Errors Encountered:              $($script:Stats.ErrorCount)" -Level SUMMARY
    Write-Host "========================================" -ForegroundColor Green
    Write-Log "Execution Time: $($ExecutionTime.ToString('hh\:mm\:ss'))" -Level INFO
    Write-Log "Log file: $script:LogPath" -Level INFO
    
    # Export report if data was collected
    if ($script:ReportData.Count -gt 0) {
        $script:ReportData | Export-Csv -Path $script:ReportPath -NoTypeInformation -Encoding UTF8
        Write-Log "Report exported: $script:ReportPath" -Level INFO
    }
    
    Write-Host "========================================" -ForegroundColor Green
    
    # Impact warnings
    if (-not $ReportOnly -and -not $WhatIfMode -and $script:Stats.VersionsDeleted -gt 0) {
        Write-Host ""
        Write-Log "IMPORTANT NOTES:" -Level WARNING
        if (-not $SkipRecycleBin) {
            Write-Log "- Deleted versions are in site recycle bins (93 day retention)" -Level WARNING
            Write-Log "- Recycle bin items still count against storage quota" -Level WARNING
            Write-Log "- Consider emptying recycle bins to reclaim storage immediately" -Level WARNING
        }
        Write-Log "- Storage metrics update every 24-48 hours" -Level WARNING
        Write-Log "- Check M365 Unified Audit Log for 'Delete' events" -Level WARNING
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

# Validate app-only auth parameters if provided
if (-not (Test-AppOnlyAuthParameters -ClientId $ClientId -CertificatePath $CertificatePath)) {
    exit 1
}

# Initialize authentication parameters for reuse
Initialize-AuthParameters -ClientId $ClientId -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword

# Determine if running in WhatIf mode
$IsWhatIfMode = $WhatIfPreference -or $PSBoundParameters.ContainsKey('WhatIf')

Write-Log "=== PowerPoint Version History Cleanup Script ===" -Level INFO
Write-Log "Configuration: Keep $MajorVersionsToKeep versions per file" -Level INFO
Write-Log "Recycle Bin: $(if ($SkipRecycleBin) { 'DISABLED - Permanent deletion' } else { 'Enabled - Versions sent to recycle bin' })" -Level INFO
Write-Log "Throttle Delay: $ThrottleDelayMs ms between sites" -Level INFO

if ($IncludeSites) {
    Write-Log "Include Sites: $($IncludeSites.Count) site(s) specified" -Level INFO
}
if ($ExcludeSites) {
    Write-Log "Exclude Patterns: $($ExcludeSites -join ', ')" -Level INFO
}

if ($ReportOnly) {
    Write-Log "[REPORT-ONLY MODE] Generating report without making changes" -Level WARNING
}
elseif ($IsWhatIfMode) {
    Write-Log "[WHAT-IF MODE] Running in preview mode - no changes will be made" -Level WARNING
}

if ($SkipRecycleBin -and -not $IsWhatIfMode -and -not $ReportOnly) {
    Write-Log "WARNING: SkipRecycleBin is enabled. Deleted versions cannot be recovered!" -Level WARNING
}

# Connect to tenant admin
$connected = Connect-ToSharePointSite -SiteUrl $TenantAdminUrl -IsTenantAdmin
if (-not $connected) {
    Write-Log "Failed to connect to tenant. Exiting." -Level ERROR
    exit 1
}

# Get all sites
$AllSites = Get-AllSharePointSites
if ($AllSites.Count -eq 0) {
    Write-Log "No sites found to process" -Level WARNING
    Write-Summary -WhatIfMode:$IsWhatIfMode -ReportOnly:$ReportOnly
    exit 0
}

# Process each site
try {
    $siteIndex = 0
    $totalSites = $AllSites.Count
    
    foreach ($Site in $AllSites) {
        $siteIndex++
        $SiteUrl = $Site.Url
        
        # Check if site should be excluded
        if (Test-SiteExcluded -SiteUrl $SiteUrl) {
            Write-Verbose "Skipping excluded site: $SiteUrl"
            $script:Stats.SitesSkipped++
            continue
        }
        
        # Update progress
        Write-Progress -Activity "Processing SharePoint Sites" `
            -Status "Site $siteIndex of $totalSites`: $SiteUrl" `
            -PercentComplete (($siteIndex / $totalSites) * 100) `
            -Id 1
        
        Write-Log "=== [$siteIndex/$totalSites] Processing Site: $SiteUrl ===" -Level INFO
        
        # Add throttle delay between sites to avoid Microsoft throttling
        if ($siteIndex -gt 1 -and $ThrottleDelayMs -gt 0) {
            Start-Sleep -Milliseconds $ThrottleDelayMs
        }
        
        # Connect to the specific site
        $siteConnected = Connect-ToSharePointSite -SiteUrl $SiteUrl
        
        if (-not $siteConnected) {
            Write-Log "Skipping site due to connection failure: $SiteUrl" -Level WARNING
            continue
        }
        
        try {
            # Get all document libraries (exclude system libraries)
            $Libraries = Get-PnPList -Includes BaseTemplate | Where-Object {
                $_.BaseTemplate -eq 101 -and
                $_.Hidden -eq $false -and
                $_.Title -notlike "Form Templates" -and
                $_.Title -notlike "Style Library" -and
                $_.Title -notlike "Site Assets" -and
                $_.Title -notlike "Site Pages" -and
                $_.Title -notlike "Preservation Hold Library"
            }
            
            if (-not $Libraries -or $Libraries.Count -eq 0) {
                Write-Verbose "No document libraries found on site: $SiteUrl"
            }
            else {
                Write-Log "Site has $($Libraries.Count) document libraries" -Level INFO
                
                foreach ($Library in $Libraries) {
                    Invoke-LibraryVersionCleanup -SiteUrl $SiteUrl -Library $Library -KeepVersions $MajorVersionsToKeep -SkipRecycleBin:$SkipRecycleBin -ReportOnly:$ReportOnly
                }
            }
            
            $script:Stats.SitesProcessed++
        }
        catch {
            Write-Log "Error processing site '$SiteUrl': $_" -Level ERROR
            $script:Stats.ErrorCount++
        }
    }
    
    Write-Progress -Activity "Processing SharePoint Sites" -Completed -Id 1
    Write-Log "Processing completed" -Level SUCCESS
}
catch {
    Write-Log "Fatal error during execution: $_" -Level ERROR
    $script:Stats.ErrorCount++
}
finally {
    # Print summary
    Write-Summary -WhatIfMode:$IsWhatIfMode -ReportOnly:$ReportOnly
    
    # Disconnect
    try {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnect errors
    }
}
