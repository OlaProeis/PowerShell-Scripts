#Requires -Version 5.1
<#
.SYNOPSIS
Simple GUI wrapper for Limit-PowerPointVersions.ps1
.DESCRIPTION
Provides a basic GUI to:
- Connect to SharePoint and retrieve all sites
- Select which sites to process
- Run the cleanup with progress tracking
- View logs for completed sites
.EXAMPLE
.\Start-CleanupGUI.ps1
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================================================
# CONFIGURATION
# ============================================================================
$script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$script:MainScript = Join-Path $script:ScriptDir "Limit-PowerPointVersions.ps1"
$script:LogDir = Join-Path $script:ScriptDir "GUILogs-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
$script:SiteData = @{}  # Stores site info and log paths
$script:IsProcessing = $false

# ============================================================================
# GUI CREATION
# ============================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerPoint Version Cleanup"
$form.Size = New-Object System.Drawing.Size(900, 650)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# --- Top Panel (Connection) ---
$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Location = New-Object System.Drawing.Point(10, 10)
$panelTop.Size = New-Object System.Drawing.Size(865, 60)

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text = "Tenant Admin URL:"
$lblTenant.Location = New-Object System.Drawing.Point(0, 8)
$lblTenant.AutoSize = $true

$txtTenant = New-Object System.Windows.Forms.TextBox
$txtTenant.Location = New-Object System.Drawing.Point(120, 5)
$txtTenant.Size = New-Object System.Drawing.Size(350, 25)
$txtTenant.Text = "https://yourtenant-admin.sharepoint.com"

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect && Load Sites"
$btnConnect.Location = New-Object System.Drawing.Point(480, 3)
$btnConnect.Size = New-Object System.Drawing.Size(130, 27)

$lblVersions = New-Object System.Windows.Forms.Label
$lblVersions.Text = "Versions to keep:"
$lblVersions.Location = New-Object System.Drawing.Point(620, 8)
$lblVersions.AutoSize = $true

$numVersions = New-Object System.Windows.Forms.NumericUpDown
$numVersions.Location = New-Object System.Drawing.Point(730, 5)
$numVersions.Size = New-Object System.Drawing.Size(60, 25)
$numVersions.Minimum = 1
$numVersions.Maximum = 500
$numVersions.Value = 5

$panelTop.Controls.AddRange(@($lblTenant, $txtTenant, $btnConnect, $lblVersions, $numVersions))

# --- Options Panel ---
$panelOptions = New-Object System.Windows.Forms.Panel
$panelOptions.Location = New-Object System.Drawing.Point(10, 75)
$panelOptions.Size = New-Object System.Drawing.Size(865, 30)

$chkReportOnly = New-Object System.Windows.Forms.CheckBox
$chkReportOnly.Text = "Report Only (no changes)"
$chkReportOnly.Location = New-Object System.Drawing.Point(0, 3)
$chkReportOnly.AutoSize = $true
$chkReportOnly.Checked = $true

$chkSkipRecycle = New-Object System.Windows.Forms.CheckBox
$chkSkipRecycle.Text = "Skip Recycle Bin (permanent delete)"
$chkSkipRecycle.Location = New-Object System.Drawing.Point(200, 3)
$chkSkipRecycle.AutoSize = $true
$chkSkipRecycle.ForeColor = [System.Drawing.Color]::DarkRed

$panelOptions.Controls.AddRange(@($chkReportOnly, $chkSkipRecycle))

# --- Site List ---
$lblSites = New-Object System.Windows.Forms.Label
$lblSites.Text = "Sites (select to process):"
$lblSites.Location = New-Object System.Drawing.Point(10, 110)
$lblSites.AutoSize = $true

$listSites = New-Object System.Windows.Forms.ListView
$listSites.Location = New-Object System.Drawing.Point(10, 130)
$listSites.Size = New-Object System.Drawing.Size(865, 350)
$listSites.View = [System.Windows.Forms.View]::Details
$listSites.CheckBoxes = $true
$listSites.FullRowSelect = $true
$listSites.GridLines = $true

# Columns
[void]$listSites.Columns.Add("Site URL", 400)
[void]$listSites.Columns.Add("Template", 120)
[void]$listSites.Columns.Add("Status", 100)
[void]$listSites.Columns.Add("Files", 60)
[void]$listSites.Columns.Add("Versions Deleted", 100)
[void]$listSites.Columns.Add("Storage Freed", 80)

# --- Button Panel ---
$panelButtons = New-Object System.Windows.Forms.Panel
$panelButtons.Location = New-Object System.Drawing.Point(10, 485)
$panelButtons.Size = New-Object System.Drawing.Size(865, 35)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(0, 3)
$btnSelectAll.Size = New-Object System.Drawing.Size(80, 27)

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Text = "Select None"
$btnSelectNone.Location = New-Object System.Drawing.Point(90, 3)
$btnSelectNone.Size = New-Object System.Drawing.Size(80, 27)

$btnProcess = New-Object System.Windows.Forms.Button
$btnProcess.Text = "Process Selected Sites"
$btnProcess.Location = New-Object System.Drawing.Point(300, 3)
$btnProcess.Size = New-Object System.Drawing.Size(150, 27)
$btnProcess.BackColor = [System.Drawing.Color]::LightGreen
$btnProcess.Enabled = $false

$btnViewLog = New-Object System.Windows.Forms.Button
$btnViewLog.Text = "View Log"
$btnViewLog.Location = New-Object System.Drawing.Point(460, 3)
$btnViewLog.Size = New-Object System.Drawing.Size(80, 27)
$btnViewLog.Enabled = $false

$btnExportList = New-Object System.Windows.Forms.Button
$btnExportList.Text = "Export Site List"
$btnExportList.Location = New-Object System.Drawing.Point(550, 3)
$btnExportList.Size = New-Object System.Drawing.Size(100, 27)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "Stop"
$btnStop.Location = New-Object System.Drawing.Point(780, 3)
$btnStop.Size = New-Object System.Drawing.Size(80, 27)
$btnStop.BackColor = [System.Drawing.Color]::LightCoral
$btnStop.Enabled = $false

$panelButtons.Controls.AddRange(@($btnSelectAll, $btnSelectNone, $btnProcess, $btnViewLog, $btnExportList, $btnStop))

# --- Progress Bar ---
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 525)
$progressBar.Size = New-Object System.Drawing.Size(865, 23)

# --- Status Label ---
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready. Enter tenant admin URL and click Connect."
$lblStatus.Location = New-Object System.Drawing.Point(10, 555)
$lblStatus.Size = New-Object System.Drawing.Size(865, 50)

# Add all controls to form
$form.Controls.AddRange(@($panelTop, $panelOptions, $lblSites, $listSites, $panelButtons, $progressBar, $lblStatus))

# ============================================================================
# EVENT HANDLERS
# ============================================================================

$btnConnect.Add_Click({
    $tenantUrl = $txtTenant.Text.Trim()
    if ([string]::IsNullOrEmpty($tenantUrl)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter the tenant admin URL.", "Error", "OK", "Error")
        return
    }
    
    $btnConnect.Enabled = $false
    $lblStatus.Text = "Connecting to $tenantUrl..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $listSites.Items.Clear()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Connect to tenant
        Connect-PnPOnline -Url $tenantUrl -Interactive -ErrorAction Stop
        
        $lblStatus.Text = "Loading sites..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Get sites (excluding system sites)
        $sites = Get-PnPTenantSite -Detailed | Where-Object {
            $_.Template -notlike "*REDIRECT*" -and
            $_.Template -notlike "*SRCHCEN*" -and
            $_.Template -notlike "*SPSMSITEHOST*" -and
            $_.Template -notlike "*APPCATALOG*" -and
            $_.Url -notlike "*-my.sharepoint.com*"
        }
        
        foreach ($site in $sites) {
            $item = New-Object System.Windows.Forms.ListViewItem($site.Url)
            $item.SubItems.Add($site.Template)
            $item.SubItems.Add("Pending")
            $item.SubItems.Add("-")
            $item.SubItems.Add("-")
            $item.SubItems.Add("-")
            $item.Tag = $site.Url
            [void]$listSites.Items.Add($item)
            
            # Store site data
            $script:SiteData[$site.Url] = @{
                Template = $site.Template
                Status   = "Pending"
                LogPath  = $null
            }
        }
        
        $lblStatus.Text = "Loaded $($sites.Count) sites. Select sites and click 'Process Selected Sites'."
        $btnProcess.Enabled = $true
    }
    catch {
        $lblStatus.Text = "Connection failed: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Failed to connect: $($_.Exception.Message)", "Error", "OK", "Error")
    }
    finally {
        $btnConnect.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnSelectAll.Add_Click({
    foreach ($item in $listSites.Items) {
        $item.Checked = $true
    }
})

$btnSelectNone.Add_Click({
    foreach ($item in $listSites.Items) {
        $item.Checked = $false
    }
})

$btnViewLog.Add_Click({
    $selected = $listSites.SelectedItems
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select a site first.", "Info", "OK", "Information")
        return
    }
    
    $siteUrl = $selected[0].Tag
    $siteInfo = $script:SiteData[$siteUrl]
    
    if ($siteInfo -and $siteInfo.LogPath -and (Test-Path $siteInfo.LogPath)) {
        Start-Process notepad.exe -ArgumentList $siteInfo.LogPath
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("No log available for this site yet.", "Info", "OK", "Information")
    }
})

$btnExportList.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text files (*.txt)|*.txt|CSV files (*.csv)|*.csv"
    $saveDialog.FileName = "sites.txt"
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        $sites = $listSites.Items | ForEach-Object { $_.Tag }
        $sites | Out-File -FilePath $saveDialog.FileName -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Exported $($sites.Count) sites to $($saveDialog.FileName)", "Success", "OK", "Information")
    }
})

$script:StopRequested = $false

$btnStop.Add_Click({
    $script:StopRequested = $true
    $lblStatus.Text = "Stop requested... finishing current site."
})

$btnProcess.Add_Click({
    $selectedSites = $listSites.Items | Where-Object { $_.Checked }
    
    if ($selectedSites.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one site.", "Info", "OK", "Information")
        return
    }
    
    # Confirm if not report-only
    if (-not $chkReportOnly.Checked) {
        $msg = "You are about to process $($selectedSites.Count) site(s).`n`n"
        if ($chkSkipRecycle.Checked) {
            $msg += "WARNING: Skip Recycle Bin is enabled - deletions are PERMANENT!`n`n"
        }
        $msg += "Continue?"
        
        $result = [System.Windows.Forms.MessageBox]::Show($msg, "Confirm", "YesNo", "Warning")
        if ($result -ne "Yes") { return }
    }
    
    # Create log directory
    if (-not (Test-Path $script:LogDir)) {
        New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
    }
    
    # Disable controls
    $script:IsProcessing = $true
    $script:StopRequested = $false
    $btnProcess.Enabled = $false
    $btnConnect.Enabled = $false
    $btnStop.Enabled = $true
    $progressBar.Value = 0
    $progressBar.Maximum = $selectedSites.Count
    
    $processed = 0
    $totalVersions = 0
    $totalStorage = 0
    
    foreach ($item in $selectedSites) {
        if ($script:StopRequested) {
            $lblStatus.Text = "Stopped by user."
            break
        }
        
        $siteUrl = $item.Tag
        $processed++
        
        $item.SubItems[2].Text = "Processing..."
        $item.EnsureVisible()
        $lblStatus.Text = "[$processed/$($selectedSites.Count)] Processing: $siteUrl"
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Build parameters
            $params = @{
                TenantAdminUrl      = $txtTenant.Text.Trim()
                MajorVersionsToKeep = [int]$numVersions.Value
                IncludeSites        = @($siteUrl)
                LogDirectory        = $script:LogDir
                Confirm             = $false
            }
            
            if ($chkReportOnly.Checked) { $params['ReportOnly'] = $true }
            if ($chkSkipRecycle.Checked) { $params['SkipRecycleBin'] = $true }
            
            # Capture output
            $output = & $script:MainScript @params 2>&1 | Out-String
            
            # Parse results from output (basic parsing)
            $filesMatch = [regex]::Match($output, "PowerPoint Files Scanned:\s+(\d+)")
            $versionsMatch = [regex]::Match($output, "Versions.*?:\s+(\d+)")
            $storageMatch = [regex]::Match($output, "Storage.*?:\s+([\d.]+\s*\w+)")
            
            $files = if ($filesMatch.Success) { $filesMatch.Groups[1].Value } else { "-" }
            $versions = if ($versionsMatch.Success) { $versionsMatch.Groups[1].Value } else { "0" }
            $storage = if ($storageMatch.Success) { $storageMatch.Groups[1].Value } else { "-" }
            
            $item.SubItems[2].Text = "Done"
            $item.SubItems[2].ForeColor = [System.Drawing.Color]::Green
            $item.SubItems[3].Text = $files
            $item.SubItems[4].Text = $versions
            $item.SubItems[5].Text = $storage
            
            if ($versions -match "^\d+$") { $totalVersions += [int]$versions }
            
            # Find log file for this run
            $logFiles = Get-ChildItem -Path $script:LogDir -Filter "*.log" | Sort-Object LastWriteTime -Descending
            if ($logFiles.Count -gt 0) {
                $script:SiteData[$siteUrl].LogPath = $logFiles[0].FullName
            }
            $script:SiteData[$siteUrl].Status = "Done"
        }
        catch {
            $item.SubItems[2].Text = "Error"
            $item.SubItems[2].ForeColor = [System.Drawing.Color]::Red
            $script:SiteData[$siteUrl].Status = "Error"
        }
        
        $progressBar.Value = $processed
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    # Re-enable controls
    $script:IsProcessing = $false
    $btnProcess.Enabled = $true
    $btnConnect.Enabled = $true
    $btnStop.Enabled = $false
    $btnViewLog.Enabled = $true
    
    $lblStatus.Text = "Complete! Processed $processed sites. Total versions affected: $totalVersions. Logs: $script:LogDir"
})

$listSites.Add_SelectedIndexChanged({
    $btnViewLog.Enabled = ($listSites.SelectedItems.Count -gt 0)
})

# ============================================================================
# RUN
# ============================================================================

# Check main script exists
if (-not (Test-Path $script:MainScript)) {
    [System.Windows.Forms.MessageBox]::Show(
        "Main script not found: $script:MainScript`n`nPlace this GUI script in the same folder as Limit-PowerPointVersions.ps1",
        "Error", "OK", "Error"
    )
    exit 1
}

[void]$form.ShowDialog()
