#Requires -Version 5.1
<#
.SYNOPSIS
Batch wrapper for Limit-PowerPointVersions.ps1. Processes sites from a file or CSV.
.DESCRIPTION
Reads a list of site URLs and runs the main cleanup script for each site individually.
Useful for splitting large tenant cleanups across multiple runs or machines.
.PARAMETER SiteListPath
Path to a text file with one site URL per line, or a CSV with a "Url" column.
.PARAMETER TenantAdminUrl
SharePoint tenant admin URL.
.PARAMETER MajorVersionsToKeep
Number of versions to keep (passed to main script).
.PARAMETER BatchSize
Number of sites to process before pausing. Set to 0 for no batching.
.PARAMETER ReportOnly
Pass-through to main script - only generate report.
.PARAMETER WhatIf
Pass-through to main script - preview mode.
.EXAMPLE
.\Invoke-BatchCleanup.ps1 -SiteListPath ".\sites.txt" -TenantAdminUrl "https://contoso-admin.sharepoint.com" -ReportOnly
.EXAMPLE
.\Invoke-BatchCleanup.ps1 -SiteListPath ".\sites.csv" -TenantAdminUrl "https://contoso-admin.sharepoint.com" -BatchSize 50
Process 50 sites, then pause for confirmation before continuing.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$SiteListPath,

    [Parameter(Mandatory)]
    [string]$TenantAdminUrl,

    [Parameter()]
    [int]$MajorVersionsToKeep = 5,

    [Parameter()]
    [int]$BatchSize = 0,

    [Parameter()]
    [switch]$ReportOnly,

    [Parameter()]
    [switch]$SkipRecycleBin,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$CertificatePath,

    [Parameter()]
    [SecureString]$CertificatePassword
)

$ErrorActionPreference = "Stop"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$mainScript = Join-Path $scriptDir "Limit-PowerPointVersions.ps1"

if (-not (Test-Path $mainScript)) {
    Write-Error "Main script not found: $mainScript"
    exit 1
}

# Load sites from file
$extension = [System.IO.Path]::GetExtension($SiteListPath).ToLower()
if ($extension -eq ".csv") {
    $sites = Import-Csv $SiteListPath | Select-Object -ExpandProperty Url
}
else {
    $sites = Get-Content $SiteListPath | Where-Object { $_ -match "^https://" }
}

if ($sites.Count -eq 0) {
    Write-Error "No valid sites found in $SiteListPath"
    exit 1
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Batch PowerPoint Version Cleanup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Sites to process: $($sites.Count)"
Write-Host "Versions to keep: $MajorVersionsToKeep"
Write-Host "Batch size: $(if ($BatchSize -eq 0) { 'No batching' } else { $BatchSize })"
Write-Host ""

# Track results
$results = [System.Collections.ArrayList]::new()
$logDir = Join-Path $scriptDir "BatchLogs-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
New-Item -ItemType Directory -Path $logDir -Force | Out-Null

$siteIndex = 0
$batchCount = 0

foreach ($siteUrl in $sites) {
    $siteIndex++
    $batchCount++
    
    Write-Host ""
    Write-Host "[$siteIndex/$($sites.Count)] Processing: $siteUrl" -ForegroundColor Yellow
    
    # Build parameters for main script
    $params = @{
        TenantAdminUrl     = $TenantAdminUrl
        MajorVersionsToKeep = $MajorVersionsToKeep
        IncludeSites       = @($siteUrl)
        LogDirectory       = $logDir
        Confirm            = $false
    }
    
    if ($ReportOnly) { $params['ReportOnly'] = $true }
    if ($SkipRecycleBin) { $params['SkipRecycleBin'] = $true }
    if ($WhatIfPreference) { $params['WhatIf'] = $true }
    if ($ClientId) { $params['ClientId'] = $ClientId }
    if ($CertificatePath) { $params['CertificatePath'] = $CertificatePath }
    if ($CertificatePassword) { $params['CertificatePassword'] = $CertificatePassword }
    
    $startTime = Get-Date
    $status = "Success"
    $errorMsg = $null
    
    try {
        & $mainScript @params
    }
    catch {
        $status = "Error"
        $errorMsg = $_.Exception.Message
        Write-Host "  ERROR: $errorMsg" -ForegroundColor Red
    }
    
    $duration = (Get-Date) - $startTime
    
    [void]$results.Add([PSCustomObject]@{
        SiteUrl   = $siteUrl
        Status    = $status
        Duration  = $duration.ToString('hh\:mm\:ss')
        Error     = $errorMsg
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    })
    
    # Batch pause
    if ($BatchSize -gt 0 -and $batchCount -ge $BatchSize -and $siteIndex -lt $sites.Count) {
        $batchCount = 0
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host "  Batch of $BatchSize sites completed" -ForegroundColor Cyan
        Write-Host "  Remaining: $($sites.Count - $siteIndex) sites" -ForegroundColor Cyan
        Write-Host "========================================" -ForegroundColor Cyan
        
        $continue = Read-Host "Continue with next batch? (Y/N)"
        if ($continue -notmatch "^[Yy]") {
            Write-Host "Stopping at user request." -ForegroundColor Yellow
            break
        }
    }
}

# Export results summary
$summaryPath = Join-Path $logDir "BatchSummary.csv"
$results | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  BATCH COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "Total sites processed: $($results.Count)"
Write-Host "Successful: $(($results | Where-Object Status -eq 'Success').Count)"
Write-Host "Errors: $(($results | Where-Object Status -eq 'Error').Count)"
Write-Host ""
Write-Host "Logs directory: $logDir"
Write-Host "Summary file: $summaryPath"
Write-Host "========================================" -ForegroundColor Green
