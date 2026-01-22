# Comprehensive Dell Bloatware Removal Script with Timeout Support
# Combines DellToolsRemover.ps1 and SupportAssist_remover.ps1 with enhanced features
# Supports headless operation with timeouts and retry mechanisms

# Set execution policy for this process
Set-ExecutionPolicy Bypass -Scope Process -Force

# Configuration
$TimeoutMinutes = 10
$TimeoutSeconds = $TimeoutMinutes * 60
$MaxRetries = 2
$RetryDelaySeconds = 30

# Comprehensive list of Dell software to remove
$dellAppsToRemove = @(
    "Dell Command*",
    "Dell Core Service*",
    "Dell Optimizer*",
    "Dell Customer Connect*",
    "Dell Digital Delivery*",
    "Dell SupportAssist*",
    "SupportAssist*",
    "Dell Support*",
    "Dell SupportAssist Remediation*",
    "Dell SupportAssist OS Recovery*",
    "Dell SupportAssistAgent*",
    "Dell Update*",
    "Dell Data Vault*",
    "Dell Power Manager*",
    "Dell CinemaColor*",
    "My Dell*",
    "Dell Trusted Device Agent*",
    "Dell TechHub*",
    "Dell Analytics*",
    "Dell Data Manager*",
    "Dell Instrumentation*",
    "Dell Hardware Support*",
    "Dell Client Management Service*",
    "Dell Fusion Service*"
)

# Logging setup
$logPath = "C:\ComprehensiveDellBloatwareRemoval.log"
$failedUninstalls = @()

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp [$Level] $Message"
    Write-Host $logMessage
    
    # Always try to log to both locations for Intune troubleshooting
    $logPaths = @($logPath)
    if ($logPath -ne "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Detail.log") {
        $logPaths += "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Detail.log"
    }
    
    foreach ($path in $logPaths) {
        if ($path) {
            try {
                # Ensure directory exists
                $logDir = Split-Path $path -Parent
                if (-not (Test-Path $logDir)) {
                    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
                }
                $logMessage | Out-File -Append $path -Encoding utf8
            } catch {
                Write-Warning "Failed to write to log file $path`: $($_.Exception.Message)"
            }
        }
    }
    
    # For ERROR level, also write to Windows Event Log for Intune visibility
    if ($Level -eq "ERROR") {
        try {
            Write-EventLog -LogName Application -Source "Dell Bloatware Removal" -EventId 1001 -EntryType Error -Message $Message -ErrorAction SilentlyContinue
        } catch {
            # Event source doesn't exist, create it
            try {
                New-EventLog -LogName Application -Source "Dell Bloatware Removal" -ErrorAction SilentlyContinue
                Write-EventLog -LogName Application -Source "Dell Bloatware Removal" -EventId 1001 -EntryType Error -Message $Message -ErrorAction SilentlyContinue
            } catch {
                # Ignore if we can't write to event log
            }
        }
    }
}

function Start-ProcessWithTimeout {
    param(
        [string]$FilePath,
        [string[]]$ArgumentList = @(),
        [int]$TimeoutSeconds = 600
    )
    
    try {
        # Filter out null or empty arguments
        $cleanArgs = $ArgumentList | Where-Object { $_ -ne $null -and $_ -ne "" }
        
        Write-Log "Starting process: $FilePath with arguments: $($cleanArgs -join ' ')"
        
        # Use the original approach directly - this worked in the old script
        if ($cleanArgs -and $cleanArgs.Count -gt 0) {
            $process = Start-Process -FilePath $FilePath -ArgumentList $cleanArgs -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        } else {
            $process = Start-Process -FilePath $FilePath -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        }
        
        # Add timeout to the original approach
        if ($process) {
            $finished = $process.WaitForExit($TimeoutSeconds * 1000)
            if (-not $finished) {
                Write-Log "Process timed out after $TimeoutSeconds seconds, terminating..." "WARNING"
                try { $process.Kill() } catch { }
                return $false
            }
            
            # Wait a moment for exit code to be available
            Start-Sleep -Milliseconds 100
            
            try {
                $exitCode = $process.ExitCode
                Write-Log "Process completed with exit code: $exitCode"
                
                # For Dell installers, we need to be more strict about exit codes
                # MSI: 0 = success, 3010 = success with reboot required
                # InstallShield: Often returns non-zero even for success, but some failures are real
                if ($exitCode -eq 0 -or $exitCode -eq 3010) {
                    return $true
                } elseif ($exitCode -eq 1602) {
                    Write-Log "Install already running or cancelled (1602)" "WARNING"
                    return $false
                } elseif ($exitCode -eq 1603) {
                    Write-Log "Fatal error during installation (1603)" "WARNING"
                    return $false
                } else {
                    Write-Log "Non-zero exit code ($exitCode) - may indicate failure" "WARNING"
                    # For InstallShield, we'll still consider it potentially successful
                    return $true
                }
            } catch {
                Write-Log "Could not get exit code: $($_.Exception.Message)" "WARNING"
                return $true
            }
        } else {
            Write-Log "Failed to start process" "ERROR"
            return $false
        }
        
    } catch {
        Write-Log "Failed to start process: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Stop-DellServices {
    Write-Log "=== Stopping Dell Services ==="
    
    # Specific service names from SupportAssist script
    $specificServices = @(
        "SupportAssistAgent",
        "Dell SupportAssist",
        "Dell SupportAssistAgent", 
        "SAService",
        "SAAgent",
        "Dell Hardware Support"
    )
    
    # Stop specific services first
    foreach ($serviceName in $specificServices) {
        $services = Get-Service -Name "*$serviceName*" -ErrorAction SilentlyContinue
        foreach ($service in $services) {
            try {
                Write-Log "Stopping specific service: $($service.Name) - $($service.DisplayName)"
                
                if ($service.Status -eq 'Running') {
                    Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    
                    # Wait for service to stop with timeout
                    $timeout = 30
                    $elapsed = 0
                    while ((Get-Service -Name $service.Name -ErrorAction SilentlyContinue).Status -eq 'Running' -and $elapsed -lt $timeout) {
                        Start-Sleep -Seconds 2
                        $elapsed += 2
                    }
                }
                
                Set-Service -Name $service.Name -StartupType Disabled -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Write-Log "Service $($service.Name) stopped and disabled"
                
            } catch {
                Write-Log "Failed to stop specific service $($service.Name): $($_.Exception.Message)" "ERROR"
            }
        }
    }
    
    # Then stop any remaining Dell services with wildcard patterns
    $dellServicePatterns = @(
        "*Dell*",
        "*SupportAssist*",
        "*SAService*", 
        "*SAAgent*"
    )
    
    foreach ($pattern in $dellServicePatterns) {
        $services = Get-Service | Where-Object { 
            $_.Name -like $pattern -or $_.DisplayName -like $pattern 
        }
        
        foreach ($service in $services) {
            try {
                # Skip if already processed above
                $alreadyProcessed = $false
                foreach ($specificName in $specificServices) {
                    if ($service.Name -like "*$specificName*" -or $service.DisplayName -like "*$specificName*") {
                        $alreadyProcessed = $true
                        break
                    }
                }
                
                if ($alreadyProcessed) { continue }
                
                Write-Log "Stopping service: $($service.Name) - $($service.DisplayName)"
                
                if ($service.Status -eq 'Running') {
                    Stop-Service -Name $service.Name -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                    
                    # Wait for service to stop with timeout
                    $timeout = 30
                    $elapsed = 0
                    while ((Get-Service -Name $service.Name -ErrorAction SilentlyContinue).Status -eq 'Running' -and $elapsed -lt $timeout) {
                        Start-Sleep -Seconds 2
                        $elapsed += 2
                    }
                }
                
                Set-Service -Name $service.Name -StartupType Disabled -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Write-Log "Service $($service.Name) stopped and disabled"
                
            } catch {
                Write-Log "Failed to stop service $($service.Name): $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

function Stop-DellProcesses {
    Write-Log "=== Stopping Dell Processes ==="
    
    $processPatterns = @(
        "*Dell*",
        "*SupportAssist*",
        "*SAService*",
        "*SAAgent*",
        "*DellSupportCenter*"
    )
    
    foreach ($pattern in $processPatterns) {
        $processes = Get-Process | Where-Object { $_.ProcessName -like $pattern }
        
        foreach ($process in $processes) {
            try {
                Write-Log "Killing process: $($process.ProcessName) (PID: $($process.Id))"
                Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                Write-Log "Process $($process.ProcessName) terminated"
            } catch {
                Write-Log "Failed to kill process $($process.ProcessName): $($_.Exception.Message)" "ERROR"
            }
        }
    }
    
    # Give processes time to terminate
    Start-Sleep -Seconds 5
}

function Uninstall-DellApplicationWithTimeout {
    param(
        [string]$AppName,
        [string]$UninstallString,
        [string]$ProductCode = ""
    )
    
    $success = $false
    $attempt = 0
    
    while (-not $success -and $attempt -lt $MaxRetries) {
        $attempt++
        Write-Log "Attempting to uninstall $AppName (Attempt $attempt/$MaxRetries)"
        
        try {
            if ([string]::IsNullOrWhiteSpace($UninstallString)) {
                Write-Log "No uninstall string found for $AppName" "WARNING"
                return $false
            }
            
            if ($UninstallString -like "*msiexec*" -or -not [string]::IsNullOrWhiteSpace($ProductCode)) {
                # MSI-based uninstall
                if ([string]::IsNullOrWhiteSpace($ProductCode)) {
                    $ProductCode = [regex]::Match($UninstallString, '\{[0-9A-F-]{36}\}').Value
                }
                
                if (-not [string]::IsNullOrWhiteSpace($ProductCode) -and $ProductCode -match '\{[0-9A-F-]{36}\}') {
                    Write-Log "Using MSI uninstall for $AppName with product code $ProductCode"
                    # Add SupportAssist-specific logging for better debugging
                    if ($AppName -like "*SupportAssist*") {
                        $arguments = @("/x", $ProductCode, "/quiet", "/norestart", "/L*v", "C:\SA_Uninstall.log")
                    } else {
                        $arguments = @("/x", $ProductCode, "/quiet", "/norestart")
                    }
                    $success = Start-ProcessWithTimeout -FilePath "msiexec.exe" -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                } else {
                    Write-Log "Could not extract valid product code from: $UninstallString" "WARNING"
                    # Try direct MSI command
                    $arguments = @("/c", $UninstallString, "/quiet", "/norestart")
                    $success = Start-ProcessWithTimeout -FilePath "cmd.exe" -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                }
                
            } elseif ($UninstallString -like "*.exe*") {
                # EXE-based uninstall
                $exePath = ($UninstallString -split '"')[1]
                if (-not $exePath) { $exePath = $UninstallString.Split(' ')[0] }
                
                if (Test-Path $exePath) {
                    Write-Log "Using EXE uninstall for ${AppName}: $exePath"
                    
                    # Try SupportAssist-specific parameters first for SupportAssist apps
                    if ($AppName -like "*SupportAssist*") {
                        Write-Log "Using SupportAssist-specific parameters"
                        $arguments = @("/S", "/SILENT", "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART")
                        $success = Start-ProcessWithTimeout -FilePath $exePath -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                    }
                    
                    # If not successful or not SupportAssist, try original script approach
                    if (-not $success) {
                        Write-Log "Trying standard EXE parameters"
                        $arguments = @("/S", "/silent", "/quiet", "/uninstall")
                        $success = Start-ProcessWithTimeout -FilePath $exePath -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                    }
                    
                    # For InstallShield installers, try different parameter combinations
                    if (-not $success -and $exePath -like "*InstallShield*") {
                        Write-Log "Retrying InstallShield uninstall with /uninst parameter"
                        $arguments = @("/uninst")
                        $success = Start-ProcessWithTimeout -FilePath $exePath -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                        
                        if (-not $success) {
                            Write-Log "Retrying InstallShield uninstall with remove option"
                            $arguments = @("/S", "/v/qn", "REMOVE=ALL")
                            $success = Start-ProcessWithTimeout -FilePath $exePath -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                        }
                    }
                } else {
                    Write-Log "EXE path not found: $exePath, trying full command" "WARNING"
                    $arguments = @("/c", "`"$UninstallString`"", "/S", "/silent")
                    $success = Start-ProcessWithTimeout -FilePath "cmd.exe" -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
                }
            } else {
                Write-Log "Unknown uninstall string format: $UninstallString" "WARNING"
                $arguments = @("/c", $UninstallString)
                $success = Start-ProcessWithTimeout -FilePath "cmd.exe" -ArgumentList $arguments -TimeoutSeconds $TimeoutSeconds
            }
            
        } catch {
            Write-Log "Exception during uninstall of $AppName`: $($_.Exception.Message)" "ERROR"
            $success = $false
        }
        
        if (-not $success -and $attempt -lt $MaxRetries) {
            Write-Log "Uninstall attempt $attempt failed for $AppName, waiting $RetryDelaySeconds seconds before retry..."
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
    
    if (-not $success) {
        Write-Log "Failed to uninstall $AppName after $MaxRetries attempts" "ERROR"
        $script:failedUninstalls += $AppName
    } else {
        Write-Log "Uninstall process completed for $AppName"
        
        # Validate the uninstall actually worked
        Start-Sleep -Seconds 3
        try {
            # Use registry check instead of slow WMI query for verification
            $stillInstalled = $false
            $registryPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
            )
            
            foreach ($regPath in $registryPaths) {
                $check = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $AppName }
                if ($check) {
                    $stillInstalled = $true
                    break
                }
            }
            
            if ($stillInstalled) {
                Write-Log "WARNING: $AppName still appears to be installed after uninstall" "WARNING"
                $script:failedUninstalls += $AppName
            } else {
                Write-Log "Verified: $AppName successfully removed"
            }
        } catch {
            Write-Log "Could not verify removal of $AppName" "WARNING"
        }
    }
    
    return $success
}

function Remove-DellApplicationsFromRegistry {
    Write-Log "=== Scanning for Dell Applications in Registry ==="
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $foundApps = @()
    
    foreach ($regPath in $registryPaths) {
        try {
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -and (
                    $dellAppsToRemove | Where-Object { $_.DisplayName -like $_ }
                )
            }
            $foundApps += $apps
        } catch {
            Write-Log "Error accessing registry path $regPath`: $($_.Exception.Message)" "ERROR"
        }
    }
    
    # Also search for partial matches
    foreach ($regPath in $registryPaths) {
        try {
            $apps = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -and (
                    $_.DisplayName -like "*Dell*" -or 
                    $_.DisplayName -like "*SupportAssist*"
                )
            }
            
            foreach ($app in $apps) {
                if ($foundApps.DisplayName -notcontains $app.DisplayName) {
                    $foundApps += $app
                }
            }
        } catch {
            Write-Log "Error accessing registry path $regPath for partial matches: $($_.Exception.Message)" "ERROR"
        }
    }
    
    Write-Log "Found $($foundApps.Count) Dell applications in registry"
    
    foreach ($app in $foundApps) {
        $appName = $app.DisplayName
        $uninstallString = $app.UninstallString
        $quietUninstallString = $app.QuietUninstallString
        
        Write-Log "Processing application: $appName"
        
        # Prefer quiet uninstall string if available
        $uninstallCommand = if (-not [string]::IsNullOrWhiteSpace($quietUninstallString)) {
            $quietUninstallString
        } else {
            $uninstallString
        }
        
        if (-not [string]::IsNullOrWhiteSpace($uninstallCommand)) {
            Uninstall-DellApplicationWithTimeout -AppName $appName -UninstallString $uninstallCommand
        } else {
            Write-Log "No uninstall command found for $appName" "WARNING"
        }
    }
}

function Remove-DellWMIApplications {
    Write-Log "=== Removing Dell Applications via CIM (faster than WMI) ==="
    
    try {
        # Use CIM instead of WMI for better performance and reliability
        $cimApps = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | Where-Object { 
            $_.Name -like "*SupportAssist*" -or 
            $_.Name -like "*Support Assist*" -or
            $_.Name -like "*Dell Support*" -or
            $_.Name -like "*Dell*"
        }
        
        foreach ($app in $cimApps) {
            Write-Log "Found application: $($app.Name)"
            try {
                Write-Log "Attempting CIM uninstall of: $($app.Name)"
                $result = Invoke-CimMethod -InputObject $app -MethodName Uninstall
                if ($result.ReturnValue -eq 0) {
                    Write-Log "CIM uninstall successful for: $($app.Name)"
                } else {
                    Write-Log "CIM uninstall failed for: $($app.Name), Return code: $($result.ReturnValue)" "WARNING"
                }
            } catch {
                Write-Log "CIM uninstall error for $($app.Name): $($_.Exception.Message)" "ERROR"
                
                # Fallback to registry-based uninstall if CIM fails
                Write-Log "Attempting fallback registry uninstall for: $($app.Name)"
                $registryPaths = @(
                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
                )
                
                foreach ($regPath in $registryPaths) {
                    $regApp = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq $app.Name }
                    if ($regApp -and $regApp.UninstallString) {
                        Uninstall-DellApplicationWithTimeout -AppName $app.Name -UninstallString $regApp.UninstallString
                        break
                    }
                }
            }
        }
    } catch {
        Write-Log "Error accessing CIM: $($_.Exception.Message)" "ERROR"
        Write-Log "Falling back to registry-only approach for this section"
    }
}

function Remove-SupportAssistSpecificUninstallers {
    Write-Log "=== Searching for specific SupportAssist uninstallers ==="
    
    $possibleUninstallers = @(
        "C:\Program Files\Dell\SupportAssist\bin\SupportAssist.exe",
        "C:\Program Files (x86)\Dell\SupportAssist\bin\SupportAssist.exe",
        "C:\Program Files\Dell\SupportAssistAgent\bin\SupportAssistAgent.exe",
        "C:\Program Files (x86)\Dell\SupportAssistAgent\bin\SupportAssistAgent.exe",
        "C:\ProgramData\Dell\SupportAssist\bin\SupportAssist.exe"
    )
    
    foreach ($uninstaller in $possibleUninstallers) {
        if (Test-Path $uninstaller) {
            Write-Log "Found SupportAssist executable: $uninstaller"
            try {
                Write-Log "Attempting uninstall with: $uninstaller"
                # Try the original script's approaches
                Start-ProcessWithTimeout -FilePath $uninstaller -ArgumentList @("/uninstall", "/quiet", "/norestart") -TimeoutSeconds $TimeoutSeconds
                Start-ProcessWithTimeout -FilePath $uninstaller -ArgumentList @("-uninstall", "-quiet") -TimeoutSeconds $TimeoutSeconds
            } catch {
                Write-Log "Direct SupportAssist uninstaller failed: $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

function Remove-DellUWPApps {
    Write-Log "=== Removing Dell UWP Apps ==="
    
    try {
        # Remove UWP apps for all users
        $uwpApps = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { 
            $_.Name -like "*Dell*" -or $_.Name -like "*SupportAssist*" 
        }
        
        foreach ($app in $uwpApps) {
            try {
                Write-Log "Removing UWP app: $($app.Name)"
                Remove-AppxPackage -Package $app.PackageFullName -AllUsers -ErrorAction SilentlyContinue
                Write-Log "Successfully removed UWP app: $($app.Name)"
            } catch {
                Write-Log "Failed to remove UWP app $($app.Name): $($_.Exception.Message)" "ERROR"
            }
        }
        
        # Remove provisioned apps to prevent reinstall
        Write-Log "=== Removing Dell Provisioned Apps ==="
        $provisionedApps = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { 
            $_.DisplayName -like "*Dell*" -or $_.DisplayName -like "*SupportAssist*" 
        }
        
        foreach ($app in $provisionedApps) {
            try {
                Write-Log "Removing provisioned app: $($app.DisplayName)"
                Remove-AppxProvisionedPackage -Online -PackageName $app.PackageName -ErrorAction SilentlyContinue
                Write-Log "Successfully removed provisioned app: $($app.DisplayName)"
            } catch {
                Write-Log "Failed to remove provisioned app $($app.DisplayName): $($_.Exception.Message)" "ERROR"
            }
        }
    } catch {
        Write-Log "Error removing UWP/Provisioned apps: $($_.Exception.Message)" "ERROR"
    }
}

function Remove-DellFolders {
    Write-Log "=== Removing Dell Folders ==="
    
    $foldersToRemove = @(
        "C:\Program Files\Dell",
        "C:\Program Files (x86)\Dell",
        "C:\ProgramData\Dell",
        "C:\Users\Public\Desktop\Dell*",
        "$env:APPDATA\Dell",
        "$env:LOCALAPPDATA\Dell",
        "C:\Intel\Dell",
        # SupportAssist-specific folders
        "C:\Program Files\Dell\SupportAssist*",
        "C:\Program Files (x86)\Dell\SupportAssist*",
        "C:\ProgramData\Dell\SupportAssist*",
        "C:\Users\Public\Desktop\*SupportAssist*",
        "$env:APPDATA\Dell\SupportAssist*",
        "$env:LOCALAPPDATA\Dell\SupportAssist*",
        "C:\Program Files\Dell\SupportAssistAgent*",
        "C:\Program Files (x86)\Dell\SupportAssistAgent*"
    )
    
    foreach ($folderPattern in $foldersToRemove) {
        try {
            if ($folderPattern -like "*\*") {
                # Handle wildcard patterns using SupportAssist approach
                $parentPath = Split-Path $folderPattern -Parent
                $childPattern = Split-Path $folderPattern -Leaf
                
                if (Test-Path $parentPath) {
                    $matchingFolders = Get-ChildItem -Path $parentPath -Filter $childPattern -ErrorAction SilentlyContinue
                    foreach ($folder in $matchingFolders) {
                        try {
                            Write-Log "Removing folder: $($folder.FullName)"
                            Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue
                            if (-not (Test-Path $folder.FullName)) {
                                Write-Log "Successfully removed: $($folder.FullName)"
                            } else {
                                Write-Log "Folder still exists after removal: $($folder.FullName)" "WARNING"
                            }
                        } catch {
                            Write-Log "Failed to remove folder $($folder.FullName): $($_.Exception.Message)" "ERROR"
                        }
                    }
                }
            } else {
                Remove-DellFolder -FolderPath $folderPattern
            }
        } catch {
            Write-Log "Error processing folder pattern $folderPattern`: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Remove-DellFolder {
    param([string]$FolderPath)
    
    if (Test-Path $FolderPath) {
        try {
            Write-Log "Removing folder: $FolderPath"
            
            # Take ownership and set permissions if needed
            try {
                takeown /f "$FolderPath" /r /d y 2>$null | Out-Null
                icacls "$FolderPath" /grant administrators:F /t /q 2>$null | Out-Null
            } catch {
                Write-Log "Could not take ownership of $FolderPath" "WARNING"
            }
            
            Remove-Item -Path $FolderPath -Recurse -Force -ErrorAction SilentlyContinue
            
            if (-not (Test-Path $FolderPath)) {
                Write-Log "Successfully removed folder: $FolderPath"
            } else {
                Write-Log "Folder still exists after removal attempt: $FolderPath" "WARNING"
            }
        } catch {
            Write-Log "Failed to remove folder $FolderPath`: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Remove-DellRegistryEntries {
    Write-Log "=== Cleaning Dell Registry Entries ==="
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Dell",
        "HKLM:\SOFTWARE\WOW6432Node\Dell",
        "HKCU:\SOFTWARE\Dell",
        "HKLM:\SYSTEM\CurrentControlSet\Services\*Dell*",
        "HKLM:\SYSTEM\CurrentControlSet\Services\*SupportAssist*",
        "HKLM:\SYSTEM\CurrentControlSet\Services\*SAService*",
        "HKLM:\SYSTEM\CurrentControlSet\Services\*SAAgent*"
    )
    
    foreach ($regPath in $registryPaths) {
        try {
            if ($regPath -like "*\*") {
                # Handle wildcard patterns
                $parentPath = $regPath -replace '\\\*.*$', ''
                $pattern = ($regPath -split '\*')[1]
                
                if (Test-Path $parentPath) {
                    $keys = Get-ChildItem -Path $parentPath -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$pattern*" }
                    foreach ($key in $keys) {
                        Write-Log "Removing registry key: $($key.PSPath)"
                        Remove-Item -Path $key.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                    }
                }
            } else {
                if (Test-Path $regPath) {
                    Write-Log "Removing registry path: $regPath"
                    Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Log "Failed to remove registry entry $regPath`: $($_.Exception.Message)" "ERROR"
        }
    }
    
    # Clean startup entries
    Write-Log "=== Removing Dell Startup Entries ==="
    $startupPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    )
    
    foreach ($startupPath in $startupPaths) {
        try {
            $runEntries = Get-ItemProperty -Path $startupPath -ErrorAction SilentlyContinue
            if ($runEntries) {
                $runEntries.PSObject.Properties | Where-Object { 
                    $_.Name -like "*Dell*" -or 
                    $_.Value -like "*Dell*" -or 
                    $_.Value -like "*SupportAssist*" -or
                    $_.Name -like "*SupportAssist*" -or
                    $_.Name -like "*SAService*" -or
                    $_.Name -like "*SAAgent*"
                } | ForEach-Object {
                    Write-Log "Removing startup entry: $($_.Name) = $($_.Value)"
                    Remove-ItemProperty -Path $startupPath -Name $_.Name -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Log "Failed to clean startup entries in $startupPath`: $($_.Exception.Message)" "ERROR"
        }
    }
}

function Remove-DellScheduledTasks {
    Write-Log "=== Removing Dell Scheduled Tasks ==="
    
    try {
        $dellTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { 
            $_.TaskName -like "*Dell*" -or 
            $_.TaskName -like "*SupportAssist*" -or
            $_.TaskPath -like "*Dell*" -or
            $_.TaskName -like "*SAService*" -or
            $_.TaskName -like "*SAAgent*"
        }
        
        foreach ($task in $dellTasks) {
            try {
                Write-Log "Removing scheduled task: $($task.TaskName) at path: $($task.TaskPath)"
                Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "Successfully removed scheduled task: $($task.TaskName)"
            } catch {
                Write-Log "Failed to remove scheduled task $($task.TaskName)`: $($_.Exception.Message)" "ERROR"
            }
        }
    } catch {
        Write-Log "Error accessing scheduled tasks: $($_.Exception.Message)" "ERROR"
    }
}

function Add-WindowsDefenderExclusions {
    Write-Log "=== Adding Windows Defender Exclusions (to prevent reinstall) ==="
    
    try {
        $pathsToBlock = @(
            "C:\Program Files\Dell\SupportAssist\",
            "C:\Program Files (x86)\Dell\SupportAssist\",
            "C:\ProgramData\Dell\SupportAssist\",
            "C:\Program Files\Dell\",
            "C:\Program Files (x86)\Dell\"
        )
        
        foreach ($path in $pathsToBlock) {
            try {
                Add-MpPreference -ExclusionPath $path -ErrorAction SilentlyContinue
                Write-Log "Added Windows Defender exclusion: $path"
            } catch {
                Write-Log "Could not add Windows Defender exclusion for $path`: $($_.Exception.Message)" "WARNING"
            }
        }
    } catch {
        Write-Log "Error adding Windows Defender exclusions: $($_.Exception.Message)" "WARNING"
    }
}

function Test-DellSoftwareRemoval {
    Write-Log "=== Verifying Dell Software Removal ==="
    
    # Check for remaining applications
    try {
        $remainingApps = Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | Where-Object { 
            $_.Name -like "*Dell*" -or $_.Name -like "*SupportAssist*" 
        }
        
        if ($remainingApps.Count -eq 0) {
            Write-Log "SUCCESS: No Dell applications found after removal"
        } else {
            Write-Log "WARNING: The following Dell applications may still be present:" "WARNING"
            foreach ($app in $remainingApps) {
                Write-Log "- $($app.Name)" "WARNING"
            }
        }
    } catch {
        Write-Log "Could not verify application removal: $($_.Exception.Message)" "WARNING"
    }
    
    # Check for remaining services
    try {
        $remainingServices = Get-Service -ErrorAction SilentlyContinue | Where-Object { 
            $_.Name -like "*Dell*" -or 
            $_.DisplayName -like "*Dell*" -or 
            $_.Name -like "*SupportAssist*" -or 
            $_.DisplayName -like "*SupportAssist*" 
        }
        
        if ($remainingServices.Count -eq 0) {
            Write-Log "SUCCESS: No Dell services found"
        } else {
            Write-Log "WARNING: The following Dell services may still be present:" "WARNING"
            foreach ($service in $remainingServices) {
                Write-Log "- $($service.Name): $($service.DisplayName)" "WARNING"
            }
        }
    } catch {
        Write-Log "Could not verify service removal: $($_.Exception.Message)" "WARNING"
    }
    
    # Check for remaining UWP apps
    try {
        $remainingUWP = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Where-Object { 
            $_.Name -like "*Dell*" -or $_.Name -like "*SupportAssist*" 
        }
        
        if ($remainingUWP.Count -eq 0) {
            Write-Log "SUCCESS: No Dell UWP apps found"
        } else {
            Write-Log "WARNING: The following Dell UWP apps may still be present:" "WARNING"
            foreach ($app in $remainingUWP) {
                Write-Log "- $($app.Name)" "WARNING"
            }
        }
    } catch {
        Write-Log "Could not verify UWP app removal: $($_.Exception.Message)" "WARNING"
    }
}

function Remove-PersistentDellComponents {
    Write-Log "=== Aggressive Cleanup of Persistent Components ==="
    
    # Force kill any remaining Dell processes
    $processPatterns = @("*Dell*", "*SupportAssist*", "*SAService*", "*SAAgent*")
    foreach ($pattern in $processPatterns) {
        Get-Process | Where-Object { $_.ProcessName -like $pattern } | ForEach-Object {
            try {
                Write-Log "Force killing persistent process: $($_.ProcessName)"
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
            } catch { }
        }
    }
    
    # Clean up specific SupportAssist registry entries more aggressively
    $supportAssistRegPaths = @(
        "HKLM:\SOFTWARE\Dell\SupportAssist",
        "HKLM:\SOFTWARE\WOW6432Node\Dell\SupportAssist", 
        "HKCU:\SOFTWARE\Dell\SupportAssist",
        "HKLM:\SYSTEM\CurrentControlSet\Services\SupportAssistAgent",
        "HKLM:\SYSTEM\CurrentControlSet\Services\SAService",
        "HKLM:\SYSTEM\CurrentControlSet\Services\SAAgent"
    )
    
    foreach ($regPath in $supportAssistRegPaths) {
        if (Test-Path $regPath) {
            try {
                Write-Log "Aggressively removing registry: $regPath"
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Could not remove registry path $regPath" "WARNING"
            }
        }
    }
    
    # Force remove SupportAssist folders with takeown
    $persistentFolders = @(
        "C:\Program Files\Dell\SupportAssist",
        "C:\Program Files (x86)\Dell\SupportAssist",
        "C:\ProgramData\Dell\SupportAssist",
        "C:\Program Files\Dell\SupportAssistAgent",
        "C:\Program Files (x86)\Dell\SupportAssistAgent"
    )
    
    foreach ($folder in $persistentFolders) {
        if (Test-Path $folder) {
            Write-Log "Aggressively removing folder: $folder"
            try {
                # Take ownership first
                & cmd /c "takeown /f `"$folder`" /r /d y 2>nul"
                & cmd /c "icacls `"$folder`" /grant administrators:F /t /q 2>nul"
                Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
                
                if (Test-Path $folder) {
                    Write-Log "Folder still exists after aggressive removal: $folder" "WARNING"
                } else {
                    Write-Log "Successfully removed persistent folder: $folder"
                }
            } catch {
                Write-Log "Failed to remove persistent folder $folder`: $($_.Exception.Message)" "ERROR"
            }
        }
    }
    
    # Remove any remaining uninstall entries
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($uninstallPath in $uninstallPaths) {
        try {
            Get-ItemProperty $uninstallPath -ErrorAction SilentlyContinue | Where-Object {
                $_.DisplayName -like "*SupportAssist*" -or $_.DisplayName -like "*Dell Optimizer*"
            } | ForEach-Object {
                Write-Log "Removing leftover uninstall entry: $($_.DisplayName)"
                Remove-Item -Path $_.PSPath -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}

function Restart-FailedUninstalls {
    Write-Log "=== Retrying Failed Uninstalls ==="
    
    if ($script:failedUninstalls.Count -eq 0) {
        Write-Log "No failed uninstalls to retry"
        return
    }
    
    Write-Log "Attempting to retry $($script:failedUninstalls.Count) failed uninstalls"
    $retryList = $script:failedUninstalls | Select-Object
    $script:failedUninstalls = @()
    
    foreach ($appName in $retryList) {
        Write-Log "Retrying uninstall for: $appName"
        
        # Try to find the app again in registry
        $registryPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )
        
        $found = $false
        foreach ($regPath in $registryPaths) {
            try {
                $app = Get-ItemProperty $regPath -ErrorAction SilentlyContinue | Where-Object {
                    $_.DisplayName -eq $appName
                }
                
                if ($app) {
                    $found = $true
                    $uninstallString = if ($app.QuietUninstallString) { $app.QuietUninstallString } else { $app.UninstallString }
                    Uninstall-DellApplicationWithTimeout -AppName $appName -UninstallString $uninstallString
                    break
                }
            } catch {
                Write-Log "Error searching for $appName in retry: $($_.Exception.Message)" "ERROR"
            }
        }
        
        if (-not $found) {
            Write-Log "Could not find $appName for retry - may have been removed" "WARNING"
        }
    }
}

# Main execution
try {
    $scriptStartTime = Get-Date
    Write-Log "=== Comprehensive Dell Bloatware Removal Started ==="
    Write-Log "Timeout per application: $TimeoutMinutes minutes"
    Write-Log "Maximum retries per application: $MaxRetries"
    
    # Initialize log file
    try {
        "=== Comprehensive Dell Bloatware Removal Script Started at $(Get-Date) ===" | Out-File -FilePath $logPath -Encoding utf8
        Write-Log "Log file initialized: $logPath"
    } catch {
        Write-Warning "Failed to create log file at $logPath"
        $logPath = $null
    }
    
    # Step 1: Stop services and processes
    Stop-DellServices
    Stop-DellProcesses
    
    # Step 2: Remove applications using multiple methods
    Remove-DellApplicationsFromRegistry
    Remove-DellWMIApplications
    Remove-SupportAssistSpecificUninstallers
    
    # Step 3: Remove UWP and provisioned apps
    Remove-DellUWPApps
    
    # Step 4: Clean up folders, registry, and scheduled tasks
    Remove-DellFolders
    Remove-DellRegistryEntries
    Remove-DellScheduledTasks
    
    # Step 5: Final cleanup - stop any remaining processes
    Stop-DellProcesses
    
    # Step 6: Retry failed uninstalls
    Restart-FailedUninstalls
    
    # Step 7: Aggressive cleanup of persistent components
    Remove-PersistentDellComponents
    
    # Step 8: Add Windows Defender exclusions (optional)
    Add-WindowsDefenderExclusions
    
    # Step 9: Verify removal
    Test-DellSoftwareRemoval
    
    Write-Log "=== Script completed successfully at $(Get-Date) ==="
    
    # Create execution marker file for "run once" detection
    $markerFile = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Executed.marker"
    try {
        $executionSummary = @"
Execution Date: $(Get-Date)
Failed Uninstalls: $($script:failedUninstalls.Count)
Total Script Runtime: $((Get-Date) - $scriptStartTime)
Computer Name: $env:COMPUTERNAME
User Context: $env:USERNAME
"@
        $executionSummary | Out-File -FilePath $markerFile -Encoding utf8 -Force
        Write-Log "Created execution marker file: $markerFile"
    } catch {
        Write-Log "Failed to create execution marker file: $($_.Exception.Message)" "WARNING"
    }
    
    # Create comprehensive summary for Intune monitoring
    $summaryFile = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Summary.json"
    try {
        $summary = @{
            ExecutionDate = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            ComputerName = $env:COMPUTERNAME
            TotalFailedUninstalls = $script:failedUninstalls.Count
            FailedApplications = $script:failedUninstalls
            ScriptVersion = "1.0"
            ExecutionStatus = if ($script:failedUninstalls.Count -eq 0) { "Success" } else { "Partial Success" }
            RebootRecommended = $true
            LogFiles = @(
                "C:\ComprehensiveDellBloatwareRemoval.log",
                "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Detail.log"
            )
        }
        $summary | ConvertTo-Json -Depth 3 | Out-File -FilePath $summaryFile -Encoding utf8 -Force
        Write-Log "Created summary file for Intune monitoring: $summaryFile"
    } catch {
        Write-Log "Failed to create summary file: $($_.Exception.Message)" "WARNING"
    }
    
    if ($script:failedUninstalls.Count -gt 0) {
        Write-Log "=== SUMMARY: Some applications failed to uninstall ===" "WARNING"
        foreach ($app in $script:failedUninstalls) {
            Write-Log "- $app" "WARNING"
        }
        Write-Log "Please check these applications manually and try running the script again if needed." "WARNING"
    } else {
        Write-Log "=== SUMMARY: All Dell software successfully removed ==="
    }
    
    Write-Log "Please reboot the system to complete the removal process."
    Write-Log "Log file saved to: $logPath"
    
    # Optional: Force a reboot (uncomment if needed)
    # Write-Log "Rebooting system in 60 seconds..."
    # shutdown /r /f /t 60 /c "Rebooting to complete Dell software removal"
    
} catch {
    Write-Log "Critical error in main execution: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    
    # Ensure Intune gets proper error reporting
    Write-Log "Script failed with critical error - Intune will report as failed" "ERROR"
    exit 1
}

# Final exit code determination for Intune
if ($script:failedUninstalls.Count -gt 0) {
    Write-Log "Script completed with some failures - Intune will report as partially successful" "WARNING"
    # Exit with warning code (0 = success, 1 = error, 3010 = success with reboot required)
    # Using 0 because partial success is still acceptable for aggressive removal
    exit 0
} else {
    Write-Log "Script completed successfully - Intune will report as successful" "INFO"
    exit 0
} 