<#
.SYNOPSIS
    Bulk adds or removes a Secondary Site Collection Admin for SharePoint Online sites.

.DESCRIPTION
    Reads a CSV (headers: Type, Url) and adds/removes the specified user as a Site Collection Admin.
    Generates a log file with results.

.EXAMPLE
    .\Manage-SPOAdmin.ps1 -AdminUrl "https://contoso-admin.sharepoint.com" -UserEmail "admin@contoso.com" -CsvPath "C:\Temp\sites.txt" -Action Add
#>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$true)]
    [string]$AdminUrl,

    [Parameter(Mandatory=$true)]
    [string]$UserEmail,

    [Parameter(Mandatory=$true)]
    [string]$CsvPath,

    [Parameter(Mandatory=$true)]
    [ValidateSet("Add", "Remove")]
    [string]$Action
)

# 1. Validation
if (-not (Test-Path $CsvPath)) {
    Write-Error "CSV File not found at $CsvPath"
    return
}

# 2. Connection
try {
    Write-Host "Connecting to SPO Admin: $AdminUrl..." -Foreground Cyan
    Connect-SPOService -Url $AdminUrl -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to SharePoint Online. Check URL and permissions."
    return
}

# 3. Import Data
$targets = Import-Csv $CsvPath
$results = @()
$dateStamp = Get-Date -Format "yyyyMMdd-HHmm"
$logPath = ".\SPOAdminLog_$Action_$dateStamp.csv"

# Determine boolean based on Action
$isAdminStatus = ($Action -eq "Add")

# 4. Processing Loop
foreach ($t in $targets) {
    $siteUrl = $t.Url
    $status = "Success"
    $msg = "$Action $UserEmail"

    Write-Host "Processing [$Action]: $siteUrl" -NoNewline

    if ($PSCmdlet.ShouldProcess($siteUrl, "$Action Site Collection Admin")) {
        try {
            # The core logic
            Set-SPOUser -Site $siteUrl -LoginName $UserEmail -IsSiteCollectionAdmin $isAdminStatus -ErrorAction Stop
            Write-Host " [OK]" -Foreground Green
        }
        catch {
            $status = "Error"
            $msg = $_.Exception.Message
            Write-Host " [FAILED]" -Foreground Red
            Write-Warning "Error details: $msg"
        }
    }
    else {
        # This block runs only during -WhatIf
        $status = "Skipped"
        $msg = "WhatIf simulation"
        Write-Host " [WhatIf]" -Foreground Yellow
    }

    # Add to reporting object
    $results += [PSCustomObject]@{
        Time      = Get-Date
        SiteUrl   = $siteUrl
        User      = $UserEmail
        Action    = $Action
        Status    = $status
        Message   = $msg
    }
}

# 5. Export Logs
$results | Export-Csv -Path $logPath -NoTypeInformation
Write-Host "`n------------------------------------------------"
Write-Host "Operation Complete."
Write-Host "Log file saved to: $logPath" -Foreground Cyan
Write-Host "------------------------------------------------"