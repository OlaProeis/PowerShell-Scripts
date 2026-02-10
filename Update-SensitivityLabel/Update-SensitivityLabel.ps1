<#
.SYNOPSIS
    Scans the entire tenant for files with a specific sensitivity label and replaces it.

.DESCRIPTION
    This script uses Microsoft Purview's Export-ContentExplorerData cmdlet to find all files
    across SharePoint Online and OneDrive that have a specific sensitivity label, then uses
    Microsoft Graph API to replace the label with a new one.

.PARAMETER OldLabelName
    The NAME of the old sensitivity label to replace (e.g., "Confidential - Pilot").

.PARAMETER OldLabelId
    The GUID of the old sensitivity label (alternative to OldLabelName).

.PARAMETER NewLabelId
    The GUID of the new sensitivity label to apply.

.PARAMETER DryRun
    If specified, only reports what would be changed without making any changes.

.PARAMETER DiscoveryOnly
    If specified, only lists files and sites - does not connect to Graph or attempt any changes.
    Use this to find which sites need permissions before running the actual migration.

.PARAMETER Workload
    Which workload to scan: ODB (OneDrive), SPO (SharePoint), or Both (default).

.EXAMPLE
    .\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Pilot" -DiscoveryOnly
    # Lists all files and sites - use this first to find where you need permissions

.EXAMPLE
    .\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Pilot" -NewLabelId "guid-here" -DryRun
    # Tests the migration without making changes

.EXAMPLE
    .\Update-SensitivityLabel.ps1 -OldLabelName "Confidential - Pilot" -NewLabelId "guid-here"
    # Executes the migration

.NOTES
    REQUIREMENTS:
    - ExchangeOnlineManagement module (for Connect-IPPSSession)
    - Microsoft.Graph module
    - Permissions: Files.ReadWrite.All, Sites.ReadWrite.All
    - Role: Data Classification Content Viewer (for Export-ContentExplorerData)
    
    FILE METADATA:
    The Graph API will update "Modified By" and "Modified Date" on files when labels are changed.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false)]
    [string]$OldLabelName,

    [Parameter(Mandatory = $false)]
    [string]$OldLabelId,

    [Parameter(Mandatory = $false)]
    [string]$NewLabelId,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$DiscoveryOnly,

    [Parameter(Mandatory = $false)]
    [ValidateSet("ODB", "SPO", "Both")]
    [string]$Workload = "Both",

    [Parameter(Mandatory = $false)]
    [string]$JustificationText = "Label migration: Replacing old label with new label",

    [Parameter(Mandatory = $false)]
    [string]$LogPath,

    [Parameter(Mandatory = $false)]
    [int]$PageSize = 100,

    [Parameter(Mandatory = $false)]
    [int]$ThrottleDelayMs = 500,

    # ============================================
    # CLIENT CREDENTIALS FOR METERED API BILLING
    # Required for assignSensitivityLabel to work!
    # ============================================
    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret
)

#region Configuration
$ErrorActionPreference = "Stop"
$script:SupportedExtensions = @(".docx", ".xlsx", ".pptx", ".pdf")
$script:TotalFound = 0
$script:ProcessedCount = 0
$script:SuccessCount = 0
$script:ErrorCount = 0
$script:SkippedCount = 0
$script:AlreadyUpdatedCount = 0
$script:Results = [System.Collections.ArrayList]::new()
#endregion

#region Logging
function Initialize-Logging {
    if ([string]::IsNullOrEmpty($LogPath)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $script:LogPath = Join-Path $PSScriptRoot "LabelMigration_$timestamp.log"
    }
    else {
        $script:LogPath = $LogPath
    }
    
    $logDir = Split-Path $script:LogPath -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    
    Write-Log "========================================" -NoTimestamp
    Write-Log "Sensitivity Label Migration" -NoTimestamp
    Write-Log "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -NoTimestamp
    Write-Log "========================================" -NoTimestamp
    Write-Log "Old Label: $(if ($OldLabelName) { $OldLabelName } else { $OldLabelId })"
    if (-not $DiscoveryOnly) {
        Write-Log "New Label ID: $NewLabelId"
    }
    Write-Log "Workload: $Workload"
    
    if ($DiscoveryOnly) {
        Write-Log "Mode: DISCOVERY ONLY (listing files and sites)" -Level "WARNING"
    }
    elseif ($DryRun) {
        Write-Log "Mode: DRY RUN (no changes)"
    }
    else {
        Write-Log "Mode: LIVE"
    }
    
    Write-Log "Log Path: $script:LogPath"
    Write-Log "----------------------------------------"
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        [switch]$NoTimestamp
    )
    
    $timestamp = if ($NoTimestamp) { "" } else { "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] " }
    $logMessage = "$timestamp[$Level] $Message"
    
    $color = switch ($Level) {
        "INFO"    { "White" }
        "WARNING" { "Yellow" }
        "ERROR"   { "Red" }
        "SUCCESS" { "Green" }
    }
    Write-Host $logMessage -ForegroundColor $color
    Add-Content -Path $script:LogPath -Value $logMessage -Encoding UTF8
}
#endregion

#region Connection Functions
function Connect-ToServices {
    Write-Log "Connecting to required services..."
    
    # Connect to Security & Compliance PowerShell
    # Check if Export-ContentExplorerData command is available (more reliable than session check)
    Write-Log "Connecting to Security & Compliance Center..."
    try {
        $cmdExists = Get-Command Export-ContentExplorerData -ErrorAction SilentlyContinue
        if ($null -eq $cmdExists) {
            Connect-IPPSSession -WarningAction SilentlyContinue
        }
        
        # Verify the command is now available
        $cmdExists = Get-Command Export-ContentExplorerData -ErrorAction SilentlyContinue
        if ($null -eq $cmdExists) {
            throw "Export-ContentExplorerData command not available after connection"
        }
        
        Write-Log "Connected to Security & Compliance Center" -Level "SUCCESS"
    }
    catch {
        Write-Log "Failed to connect: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
    
    # Connect to Microsoft Graph (skip in DiscoveryOnly mode)
    if (-not $DiscoveryOnly) {
        Write-Log "Connecting to Microsoft Graph..."
        try {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            
            if ($null -eq $context) {
                # Check if client credentials are provided (required for Metered API!)
                if (-not [string]::IsNullOrEmpty($ClientId) -and -not [string]::IsNullOrEmpty($ClientSecret) -and -not [string]::IsNullOrEmpty($TenantId)) {
                    # =====================================================
                    # CONFIDENTIAL CLIENT - Required for Metered API billing
                    # =====================================================
                    Write-Log "Using Client Credentials (Confidential Client) - Required for Metered API" -Level "INFO"
                    
                    $SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
                    $ClientCredential = New-Object System.Management.Automation.PSCredential($ClientId, $SecureSecret)
                    
                    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $ClientCredential -NoWelcome
                    $context = Get-MgContext
                    
                    Write-Log "Connected to Graph as App: $($context.AppName) [$($context.ClientId)]" -Level "SUCCESS"
                }
                else {
                    # =====================================================
                    # WARNING: Interactive login will NOT work with Metered API!
                    # =====================================================
                    Write-Log "WARNING: No Client Credentials provided!" -Level "WARNING"
                    Write-Log "Metered API (assignSensitivityLabel) requires a Confidential Client." -Level "WARNING"
                    Write-Log "Interactive login will result in 402 Payment Required errors." -Level "WARNING"
                    Write-Log "" -Level "WARNING"
                    Write-Log "To fix: Provide -TenantId, -ClientId, and -ClientSecret parameters" -Level "WARNING"
                    Write-Log "See SOLUTION_FIX.md for setup instructions." -Level "WARNING"
                    Write-Log "" -Level "WARNING"
                    
                    # Still connect interactively (works for DiscoveryOnly or DryRun to see what would happen)
                    Write-Log "Connecting interactively (limited functionality)..." -Level "WARNING"
                    Connect-MgGraph -Scopes "Files.ReadWrite.All", "Sites.ReadWrite.All" -NoWelcome
                    $context = Get-MgContext
                    
                    Write-Log "Connected to Graph as: $($context.Account)" -Level "SUCCESS"
                    Write-Log "NOTE: Label changes will fail with 402 error without Client Credentials!" -Level "WARNING"
                }
            }
            else {
                Write-Log "Already connected to Graph as: $($context.Account)" -Level "SUCCESS"
            }
        }
        catch {
            Write-Log "Failed to connect: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
    }
    else {
        Write-Log "Skipping Graph connection (Discovery mode)" -Level "INFO"
    }
}
#endregion

#region Content Explorer Functions
function Get-FilesWithLabel {
    param(
        [string]$LabelName,
        [string]$WorkloadFilter
    )
    
    Write-Log "Searching tenant for files with label: $LabelName"
    Write-Log "This may take several minutes for large tenants..."
    
    $allFiles = [System.Collections.ArrayList]::new()
    $workloads = if ($WorkloadFilter -eq "Both") { @("ODB", "SPO") } else { @($WorkloadFilter) }
    
    foreach ($wl in $workloads) {
        Write-Log "Scanning workload: $wl"
        
        try {
            $hasMore = $true
            $pageCookie = $null
            $pageNum = 0
            
            while ($hasMore) {
                $pageNum++
                Write-Log "  Fetching page $pageNum..."
                
                $params = @{
                    TagType  = "Sensitivity"
                    TagName  = $LabelName
                    Workload = $wl
                    PageSize = $PageSize
                }
                
                if ($null -ne $pageCookie) {
                    $params.PageCookie = $pageCookie
                }
                
                Write-Log "  Calling Export-ContentExplorerData with TagName='$LabelName', Workload='$wl'..."
                $result = Export-ContentExplorerData @params
                
                # Debug: Show what we got back
                $resultType = if ($null -eq $result) { "NULL" } else { $result.GetType().Name }
                $resultCount = if ($null -eq $result) { 0 } elseif ($result -is [array]) { $result.Count } else { 1 }
                Write-Log "  Result type: $resultType, Count: $resultCount"
                
                # Handle null or empty result
                if ($null -eq $result) {
                    Write-Log "  No results returned for $wl (result is null)" -Level "WARNING"
                    $hasMore = $false
                    continue
                }
                
                # Force to array to handle single object returns
                $resultArray = @($result)
                
                if ($resultArray.Count -eq 0) {
                    Write-Log "  No results returned for $wl (empty array)"
                    $hasMore = $false
                    continue
                }
                
                # Debug: Show first element properties
                $firstItem = $resultArray[0]
                if ($null -ne $firstItem) {
                    $props = ($firstItem.PSObject.Properties.Name | Select-Object -First 5) -join ", "
                    Write-Log "  First item properties: $props"
                }
                else {
                    Write-Log "  Warning: First element is null" -Level "WARNING"
                    $hasMore = $false
                    continue
                }
                
                # First element is metadata, rest are file results
                if ($resultArray.Count -gt 1) {
                    $items = @($resultArray | Select-Object -Skip 1)
                    
                    # Debug: Show ALL properties and values of the first file item
                    if ($items.Count -gt 0 -and $null -ne $items[0]) {
                        Write-Log "  --- First file item properties ---"
                        foreach ($prop in $items[0].PSObject.Properties) {
                            $val = if ($null -eq $prop.Value) { "(null)" } else { $prop.Value.ToString().Substring(0, [Math]::Min(100, $prop.Value.ToString().Length)) }
                            Write-Log "    $($prop.Name) = $val"
                        }
                        Write-Log "  -----------------------------------"
                    }
                    
                    $addedCount = 0
                    foreach ($item in $items) {
                        if ($null -eq $item) {
                            Write-Log "  Skipping null item" -Level "WARNING"
                            continue
                        }
                        
                        # Export-ContentExplorerData returns different property names
                        # Try multiple possible property names for file name
                        $fileName = $item.Name
                        if ([string]::IsNullOrEmpty($fileName)) { $fileName = $item.FileName }
                        if ([string]::IsNullOrEmpty($fileName)) { $fileName = $item.DocumentName }
                        
                        # Try multiple possible property names for file URL/path
                        # SiteUrl seems to be the most reliable for SPO/ODB
                        $fileUrl = $item.SiteUrl
                        if ([string]::IsNullOrEmpty($fileUrl)) { $fileUrl = $item.DocumentLink }
                        if ([string]::IsNullOrEmpty($fileUrl)) { $fileUrl = $item.ContentUri }
                        if ([string]::IsNullOrEmpty($fileUrl)) { $fileUrl = $item.FileUrl }
                        if ([string]::IsNullOrEmpty($fileUrl)) { $fileUrl = $item.Url }
                        if ([string]::IsNullOrEmpty($fileUrl)) { $fileUrl = $item.Path }
                        # Fallback: check any property containing 'url' or 'link' (case-insensitive)
                        if ([string]::IsNullOrEmpty($fileUrl)) {
                            foreach ($prop in $item.PSObject.Properties) {
                                if ($prop.Name -match '(url|link|uri|path)' -and -not [string]::IsNullOrEmpty($prop.Value)) {
                                    $fileUrl = $prop.Value
                                    break
                                }
                            }
                        }
                        if ([string]::IsNullOrEmpty($fileUrl)) { $fileUrl = $item.Location }
                        
                        # Try multiple possible property names for location/site
                        $location = $item.SiteUrl
                        if ([string]::IsNullOrEmpty($location)) { $location = $item.Location }
                        if ([string]::IsNullOrEmpty($location)) { $location = $item.Site }
                        
                        # Try multiple possible property names for last modified
                        $lastModified = $item.LastModifiedTime
                        if ($null -eq $lastModified) { $lastModified = $item.LastModifiedDate }
                        if ($null -eq $lastModified) { $lastModified = $item.Modified }
                        
                        # Try multiple possible property names for created by
                        $createdBy = $item.CreatedBy
                        if ([string]::IsNullOrEmpty($createdBy)) { $createdBy = $item.Author }
                        
                        if (-not [string]::IsNullOrEmpty($fileName)) {
                            [void]$allFiles.Add([PSCustomObject]@{
                                FileName     = $fileName
                                FileUrl      = if ([string]::IsNullOrEmpty($fileUrl)) { "N/A" } else { $fileUrl }
                                Location     = if ([string]::IsNullOrEmpty($location)) { $wl } else { $location }
                                Workload     = $wl
                                LastModified = $lastModified
                                CreatedBy    = $createdBy
                            })
                            $addedCount++
                        }
                        else {
                            Write-Log "  Skipping item - missing FileName" -Level "WARNING"
                        }
                    }
                    
                    Write-Log "  Found $($items.Count) items, added $addedCount files on page $pageNum"
                }
                else {
                    Write-Log "  Only metadata returned, no file items"
                }
                
                # Check for more pages safely
                $metadata = $resultArray[0]
                if ($null -ne $metadata -and $metadata.PSObject.Properties.Name -contains 'MorePagesAvailable') {
                    $hasMore = [bool]$metadata.MorePagesAvailable
                    if ($hasMore -and $metadata.PSObject.Properties.Name -contains 'PageCookie') {
                        $newPageCookie = $metadata.PageCookie
                        
                        # Safety check: If cookie didn't change, we're stuck in infinite loop
                        if ($newPageCookie -eq $pageCookie) {
                            Write-Log "  Warning: Page cookie unchanged - stopping to prevent infinite loop" -Level "WARNING"
                            $hasMore = $false
                        }
                        else {
                            $pageCookie = $newPageCookie
                        }
                    }
                    else {
                        $hasMore = $false
                    }
                }
                else {
                    Write-Log "  No 'MorePagesAvailable' property found in metadata"
                    $hasMore = $false
                }
            }
            
            Write-Log "Completed scanning $wl"
        }
        catch {
            Write-Log "Error scanning $wl workload: $($_.Exception.Message)" -Level "ERROR"
            Write-Log "  Stack trace: $($_.ScriptStackTrace)" -Level "ERROR"
        }
    }
    
    Write-Log "Total files found with label '$LabelName': $($allFiles.Count)" -Level "SUCCESS"
    $script:TotalFound = $allFiles.Count
    
    return $allFiles
}
#endregion

#region Graph API Functions
function Get-DriveItemFromUrl {
    param(
        [string]$FileUrl
    )
    
    try {
        # Convert URL to share token for Graph API
        $encodedUrl = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($FileUrl))
        $encodedUrl = $encodedUrl.TrimEnd('=').Replace('/', '_').Replace('+', '-')
        $shareToken = "u!$encodedUrl"
        
        $graphUri = "https://graph.microsoft.com/v1.0/shares/$shareToken/driveItem"
        $driveItem = Invoke-MgGraphRequest -Method GET -Uri $graphUri -ErrorAction Stop
        
        return @{
            DriveId = $driveItem.parentReference.driveId
            ItemId  = $driveItem.id
            Name    = $driveItem.name
        }
    }
    catch {
        return $null
    }
}

function Get-CurrentSensitivityLabel {
    param(
        [string]$DriveId,
        [string]$ItemId
    )
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/extractSensitivityLabels"
        $response = Invoke-MgGraphRequest -Method POST -Uri $uri -ErrorAction Stop
        
        if ($null -ne $response -and $null -ne $response.labels -and $response.labels.Count -gt 0) {
            return $response.labels[0].sensitivityLabelId
        }
        return $null
    }
    catch {
        # Return null if we can't get current label - will proceed with update attempt
        return $null
    }
}

function Set-FileSensitivityLabel {
    param(
        [string]$DriveId,
        [string]$ItemId,
        [string]$LabelId,
        [string]$FileName
    )
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/drives/$DriveId/items/$ItemId/assignSensitivityLabel"
        $body = @{
            sensitivityLabelId = $LabelId
            assignmentMethod   = "privileged"
            justificationText  = $JustificationText
        }
        
        $null = Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($body | ConvertTo-Json) -ContentType "application/json" -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log "Failed to set label on '$FileName': $($_.Exception.Message)" -Level "ERROR"
        
        # Try to extract detailed error from response body
        $responseBody = $null
        try {
            if ($_.Exception.Response) {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($null -ne $stream) {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $responseBody = $reader.ReadToEnd()
                    $reader.Close()
                }
            }
        } catch { }
        
        if (-not [string]::IsNullOrEmpty($responseBody)) {
            Write-Log "  Response body: $responseBody" -Level "ERROR"
        }
        
        # Also check ErrorDetails (PowerShell sometimes puts it here)
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
            Write-Log "  Error details: $($_.ErrorDetails.Message)" -Level "ERROR"
        }
        
        # Log the full exception for debugging
        Write-Log "  Exception type: $($_.Exception.GetType().FullName)" -Level "ERROR"
        Write-Log "  Stack: $($_.ScriptStackTrace)" -Level "ERROR"
        
        return $false
    }
}
#endregion

#region Main Processing
function Start-Migration {
    # Resolve label name
    $labelName = $OldLabelName
    if ([string]::IsNullOrEmpty($labelName) -and -not [string]::IsNullOrEmpty($OldLabelId)) {
        Write-Log "Looking up label name from ID..."
        try {
            $label = Get-Label | Where-Object { $_.Guid -eq $OldLabelId }
            if ($label) {
                $labelName = $label.Name
                Write-Log "Found label name: $labelName"
            }
            else {
                throw "Could not find label with ID: $OldLabelId"
            }
        }
        catch {
            Write-Log "Error looking up label: $($_.Exception.Message)" -Level "ERROR"
            throw "Please provide -OldLabelName instead of -OldLabelId"
        }
    }
    
    if ([string]::IsNullOrEmpty($labelName)) {
        throw "Either -OldLabelName or -OldLabelId must be provided"
    }
    
    # Find all files with the label
    $files = Get-FilesWithLabel -LabelName $labelName -WorkloadFilter $Workload
    
    if ($files.Count -eq 0) {
        Write-Log "No files found with label '$labelName'" -Level "WARNING"
        return
    }
    
    Write-Log ""
    Write-Log "========================================" -NoTimestamp
    Write-Log "Processing $($files.Count) files..." -NoTimestamp
    if ($DryRun) {
        Write-Log "MODE: DRY RUN - No changes will be made" -NoTimestamp
    }
    Write-Log "========================================" -NoTimestamp
    
    $fileNumber = 0
    foreach ($file in $files) {
        $fileNumber++
        $script:ProcessedCount++
        
        $progress = [math]::Round(($fileNumber / $files.Count) * 100, 1)
        Write-Progress -Activity "Processing files" -Status "$fileNumber of $($files.Count) - $($file.FileName)" -PercentComplete $progress
        
        $extension = [System.IO.Path]::GetExtension($file.FileName).ToLower()
        
        $result = [PSCustomObject]@{
            FileName   = $file.FileName
            FileUrl    = $file.FileUrl
            Location   = $file.Location
            Workload   = $file.Workload
            Extension  = $extension
            Action     = "None"
            Status     = "Pending"
            Error      = $null
        }
        
        # Check if supported file type
        if ($script:SupportedExtensions -notcontains $extension) {
            $result.Action = "Skipped"
            $result.Status = "UnsupportedType"
            $result.Error = "File type $extension not supported"
            $script:SkippedCount++
            Write-Log "Skipped (unsupported type): $($file.FileName)" -Level "WARNING"
            [void]$script:Results.Add($result)
            continue
        }
        
        # Resolve URL to drive item
        $driveItem = Get-DriveItemFromUrl -FileUrl $file.FileUrl
        
        if ($null -eq $driveItem) {
            $result.Action = "Failed"
            $result.Status = "ResolveError"
            $result.Error = "Could not resolve file URL (check permissions)"
            $script:ErrorCount++
            Write-Log "Failed to resolve: $($file.FileName) - Check if account has access" -Level "ERROR"
            [void]$script:Results.Add($result)
            continue
        }
        
        # Check current label to avoid processing stale index data
        $currentLabelId = Get-CurrentSensitivityLabel -DriveId $driveItem.DriveId -ItemId $driveItem.ItemId
        
        # If we can't read the label, we likely can't write it either (permission issue)
        if ($null -eq $currentLabelId) {
            $result.Action = "Failed"
            $result.Status = "ReadError"
            $result.Error = "Could not read current label (permission denied or no label exists)"
            $script:ErrorCount++
            Write-Log "Failed to read label: $($file.FileName)" -Level "ERROR"
            [void]$script:Results.Add($result)
            continue
        }
        
        if ($currentLabelId -eq $NewLabelId) {
            $result.Action = "Skipped"
            $result.Status = "AlreadyUpdated"
            $result.Error = "Label already matches target (index lag)"
            $script:AlreadyUpdatedCount++
            Write-Log "Skipped (already updated): $($file.FileName)" -Level "INFO"
            [void]$script:Results.Add($result)
            continue
        }
        
        $result.Action = "Replace"
        
        # DryRun check
        if ($DryRun) {
            Write-Log "[DRY RUN] Would update: $($file.FileName)" -Level "WARNING"
            $result.Status = "WouldUpdate"
            $script:SkippedCount++
        }
        else {
            # Actually perform the label change
            Write-Log "Updating label on: $($file.FileName)"
            $success = Set-FileSensitivityLabel -DriveId $driveItem.DriveId -ItemId $driveItem.ItemId -LabelId $NewLabelId -FileName $file.FileName
            
            if ($success) {
                Write-Log "Success: $($file.FileName)" -Level "SUCCESS"
                $result.Status = "Success"
                $script:SuccessCount++
            }
            else {
                $result.Status = "Failed"
                $result.Error = "Failed to apply new label"
                $script:ErrorCount++
            }
            
            # Throttle to avoid rate limiting
            Start-Sleep -Milliseconds $ThrottleDelayMs
        }
        
        [void]$script:Results.Add($result)
    }
    
    Write-Progress -Activity "Processing files" -Completed
}

function Write-DiscoverySummary {
    param(
        [array]$Files
    )
    
    Write-Log ""
    Write-Log "========================================" -NoTimestamp
    Write-Log "Discovery Summary" -NoTimestamp
    Write-Log "========================================" -NoTimestamp
    Write-Log "Total files found: $($Files.Count)"
    
    # Extract unique sites from file locations
    $sites = @{}
    foreach ($file in $Files) {
        # Parse location to get site URL
        $location = $file.Location
        if (-not [string]::IsNullOrEmpty($location)) {
            # Location format is usually the site/OneDrive path
            if (-not $sites.ContainsKey($location)) {
                $sites[$location] = @{
                    Workload = $file.Workload
                    FileCount = 0
                }
            }
            $sites[$location].FileCount++
        }
    }
    
    Write-Log ""
    Write-Log "Unique sites/locations requiring permissions: $($sites.Count)" -Level "WARNING"
    Write-Log ""
    
    # Group by workload
    $spSites = $sites.GetEnumerator() | Where-Object { $_.Value.Workload -eq "SPO" }
    $odSites = $sites.GetEnumerator() | Where-Object { $_.Value.Workload -eq "ODB" }
    
    if ($spSites.Count -gt 0) {
        Write-Log "SharePoint Sites ($($spSites.Count)):" -Level "INFO"
        foreach ($site in ($spSites | Sort-Object { $_.Value.FileCount } -Descending)) {
            Write-Log "  $($site.Key) ($($site.Value.FileCount) files)"
        }
        Write-Log ""
    }
    
    if ($odSites.Count -gt 0) {
        Write-Log "OneDrive Locations ($($odSites.Count)):" -Level "INFO"
        foreach ($site in ($odSites | Sort-Object { $_.Value.FileCount } -Descending)) {
            Write-Log "  $($site.Key) ($($site.Value.FileCount) files)"
        }
        Write-Log ""
    }
    
    # Export files list
    $csvPath = $script:LogPath -replace '\.log$', '_files.csv'
    if ($null -ne $Files -and $Files.Count -gt 0) {
        $Files | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "File list exported to: $csvPath"
    }
    else {
        Write-Log "No files to export to CSV" -Level "WARNING"
    }
    
    # Export sites list
    $sitesPath = $script:LogPath -replace '\.log$', '_sites.csv'
    if ($sites.Count -gt 0) {
        $sitesList = $sites.GetEnumerator() | ForEach-Object {
            [PSCustomObject]@{
                Location  = $_.Key
                Workload  = $_.Value.Workload
                FileCount = $_.Value.FileCount
            }
        }
        $sitesList | Export-Csv -Path $sitesPath -NoTypeInformation -Encoding UTF8
        Write-Log "Sites list exported to: $sitesPath"
    }
    else {
        Write-Log "No sites to export to CSV" -Level "WARNING"
    }
    
    Write-Log ""
    Write-Log "NEXT STEPS:" -Level "WARNING"
    Write-Log "1. Ensure you have access to the SharePoint sites listed above" -Level "WARNING"
    Write-Log "2. Ensure you have access to the OneDrive locations listed above" -Level "WARNING"
    Write-Log "3. Run script again with -NewLabelId parameter to migrate" -Level "WARNING"
    Write-Log "========================================" -NoTimestamp
}

function Write-Summary {
    Write-Log ""
    Write-Log "========================================" -NoTimestamp
    Write-Log "Migration Summary" -NoTimestamp
    Write-Log "========================================" -NoTimestamp
    Write-Log "Total files found: $($script:TotalFound)"
    Write-Log "Files processed: $($script:ProcessedCount)"
    
    if ($DryRun) {
        Write-Log "Would update: $(($script:Results | Where-Object Status -eq 'WouldUpdate').Count)" -Level "WARNING"
        Write-Log ""
        Write-Log "THIS WAS A DRY RUN - NO CHANGES WERE MADE" -Level "WARNING"
        Write-Log "Remove -DryRun parameter to execute the migration" -Level "WARNING"
    }
    else {
        Write-Log "Successfully updated: $($script:SuccessCount)" -Level "SUCCESS"
        Write-Log "Failed: $($script:ErrorCount)" -Level $(if ($script:ErrorCount -gt 0) { "ERROR" } else { "INFO" })
    }
    
    Write-Log "Already had new label: $($script:AlreadyUpdatedCount)"
    Write-Log "Skipped (other): $($script:SkippedCount)"
    
    # Export results
    $csvPath = $script:LogPath -replace '\.log$', '_results.csv'
    $script:Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log ""
    Write-Log "Detailed results exported to: $csvPath"
    Write-Log "========================================" -NoTimestamp
}
#endregion

#region Entry Point
try {
    # Validate parameters
    if ([string]::IsNullOrEmpty($OldLabelName) -and [string]::IsNullOrEmpty($OldLabelId)) {
        throw "Either -OldLabelName or -OldLabelId must be provided"
    }
    
    # NewLabelId is required unless in DiscoveryOnly mode
    if (-not $DiscoveryOnly -and [string]::IsNullOrEmpty($NewLabelId)) {
        throw "NewLabelId is required. Use -DiscoveryOnly to list files without migrating."
    }
    
    # Check for required modules
    $requiredModules = @(
        @{ Name = "ExchangeOnlineManagement"; InstallCmd = "Install-Module ExchangeOnlineManagement -Scope CurrentUser" }
    )
    
    # Graph module only needed if not in DiscoveryOnly mode
    if (-not $DiscoveryOnly) {
        $requiredModules += @{ Name = "Microsoft.Graph.Authentication"; InstallCmd = "Install-Module Microsoft.Graph -Scope CurrentUser" }
    }
    
    foreach ($mod in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $mod.Name)) {
            throw "Missing required module: $($mod.Name). Run: $($mod.InstallCmd)"
        }
    }
    
    # Import modules
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    if (-not $DiscoveryOnly) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }
    
    # Initialize logging
    Initialize-Logging
    
    # Connect to services
    Connect-ToServices
    
    # List available sensitivity labels for diagnostics
    Write-Log ""
    Write-Log "Available Sensitivity Labels in tenant:" -Level "INFO"
    Write-Log "----------------------------------------"
    try {
        $allLabels = Get-Label -ErrorAction Stop
        if ($null -eq $allLabels -or @($allLabels).Count -eq 0) {
            Write-Log "  No labels found - check if you have Data Classification permissions" -Level "WARNING"
        }
        else {
            foreach ($lbl in $allLabels) {
                $labelInfo = "  Name: '$($lbl.Name)'"
                if ($lbl.DisplayName -and $lbl.DisplayName -ne $lbl.Name) {
                    $labelInfo += " (Display: '$($lbl.DisplayName)')"
                }
                $labelInfo += " | GUID: $($lbl.Guid)"
                Write-Log $labelInfo
            }
        }
    }
    catch {
        Write-Log "  Could not retrieve labels: $($_.Exception.Message)" -Level "WARNING"
    }
    Write-Log "----------------------------------------"
    Write-Log ""
    
    # Discovery mode: just find files and list sites
    if ($DiscoveryOnly) {
        # Resolve label name
        $labelName = $OldLabelName
        if ([string]::IsNullOrEmpty($labelName) -and -not [string]::IsNullOrEmpty($OldLabelId)) {
            Write-Log "Looking up label name from ID..."
            $label = Get-Label | Where-Object { $_.Guid -eq $OldLabelId }
            if ($label) {
                $labelName = $label.Name
                Write-Log "Found label name: $labelName"
            }
            else {
                throw "Could not find label with ID: $OldLabelId"
            }
        }
        
        $files = Get-FilesWithLabel -LabelName $labelName -WorkloadFilter $Workload
        Write-DiscoverySummary -Files $files
    }
    else {
        # Run migration
        Start-Migration
        
        # Summary
        Write-Summary
    }
}
catch {
    if ($script:LogPath -and (Test-Path (Split-Path $script:LogPath -Parent) -ErrorAction SilentlyContinue)) {
        Write-Log "FATAL ERROR: $($_.Exception.Message)" -Level "ERROR"
    }
    else {
        Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    }
    exit 1
}
#endregion
