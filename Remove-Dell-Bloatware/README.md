# Remove-Dell-Bloatware

Comprehensive PowerShell script to remove Dell pre-installed software (bloatware) from Windows machines. Battle-tested on 1000+ Dell devices.

## Files Included

| File | Description |
|------|-------------|
| `dellBloatware.ps1` | Main PowerShell script that performs the removal |
| `install.cmd` | Wrapper for Intune deployment |
| `detection.ps1` | Intune detection script for "run once" behavior |

## What It Removes

- Dell SupportAssist (all variants)
- Dell Command Update
- Dell Digital Delivery
- Dell Optimizer
- Dell Power Manager
- Dell Customer Connect
- My Dell
- Dell TechHub
- Dell Analytics
- And many more Dell utilities...

## Features

- **Multiple removal methods**: Registry-based, CIM/WMI, direct uninstallers, UWP apps
- **Timeout support**: Prevents hanging on stubborn uninstallers (10 min default)
- **Retry mechanism**: Automatically retries failed uninstalls
- **Service management**: Stops and disables Dell services before removal
- **Process termination**: Kills running Dell processes
- **Registry cleanup**: Removes Dell registry entries and startup items
- **Scheduled task removal**: Removes Dell scheduled tasks
- **Verification**: Confirms removal after execution
- **Comprehensive logging**: Detailed logs for troubleshooting

## Requirements

- Windows 10/11
- PowerShell 5.1 or later
- **Administrator privileges** (required)

## Usage

### Standalone Execution

```powershell
# Run as Administrator
Set-ExecutionPolicy Bypass -Scope Process -Force
.\dellBloatware.ps1
```

Or right-click the script and select "Run with PowerShell" (as Administrator).

### Intune Deployment

The script includes files for Microsoft Intune deployment:

#### Files

| File | Purpose |
|------|---------|
| `dellBloatware.ps1` | Main removal script |
| `install.cmd` | Wrapper script for Intune execution |
| `detection.ps1` | Detection script for "run once" behavior |

#### Setup as Win32 App

1. **Package the files** - Include all three files in your `.intunewin` package
2. **Install command**: `install.cmd`
3. **Uninstall command**: `cmd /c echo Uninstall not applicable`
4. **Detection rule**: Use `detection.ps1` as a custom script
   - Script runs in: 64-bit context
   - Enforce script signature check: No
   - Run script in 32-bit on 64-bit clients: No

#### How Detection Works

The detection script implements a "run once" pattern:
- **Exit 1** (Not Compliant): Script hasn't run yet → Intune will execute the install
- **Exit 0** (Compliant): Marker file exists → Intune skips execution

This prevents the script from running repeatedly on every Intune sync.

#### Marker File

After successful execution, the script creates:
```
C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Executed.marker
```

Delete this file if you need to force re-execution on a device.

## Configuration

You can adjust these settings at the top of the script:

```powershell
$TimeoutMinutes = 10        # Max time per application uninstall
$MaxRetries = 2             # Number of retry attempts
$RetryDelaySeconds = 30     # Wait time between retries
```

## Output Files

| File | Description |
|------|-------------|
| `C:\ComprehensiveDellBloatwareRemoval.log` | Main execution log |
| `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Detail.log` | Detailed log (Intune path) |
| `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Summary.json` | JSON summary for monitoring |
| `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Executed.marker` | Execution marker for detection |

## How It Works

1. **Stop Services** - Stops all Dell-related Windows services
2. **Kill Processes** - Terminates running Dell processes
3. **Uninstall Applications** - Removes apps via registry, CIM, and direct uninstallers
4. **Remove UWP Apps** - Removes Dell Store apps and provisioned packages
5. **Clean Folders** - Removes Dell directories from Program Files, ProgramData, etc.
6. **Clean Registry** - Removes Dell registry keys and startup entries
7. **Remove Scheduled Tasks** - Removes Dell scheduled tasks
8. **Retry Failed** - Retries any failed uninstalls
9. **Aggressive Cleanup** - Final pass for stubborn components
10. **Verify** - Confirms removal completed

## Supported Dell Software

The script targets these patterns:

- Dell Command*
- Dell Core Service*
- Dell Optimizer*
- Dell Customer Connect*
- Dell Digital Delivery*
- Dell SupportAssist*
- Dell Update*
- Dell Power Manager*
- My Dell*
- Dell TechHub*
- And more...

## Troubleshooting

### Script hangs
The script has built-in timeouts (10 minutes per app). If an uninstall hangs, it will be terminated and retried.

### Some apps remain after running
1. Reboot the machine
2. Run the script again
3. Check the log file for specific errors

### Permission errors
Ensure you're running as Administrator. The script requires elevated privileges to:
- Stop services
- Modify registry
- Remove system files

### Verification shows remaining items
Some Dell components may require a reboot to fully remove. The script recommends rebooting after completion.

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success (or partial success with some non-critical failures) |
| 1 | Critical error during execution |

## Notes

- **Reboot recommended** after running to complete removal
- The script adds Windows Defender exclusions to help prevent Dell software from being reinstalled
- Safe to run multiple times - already-removed items are skipped
- Does not remove Dell hardware drivers (only utility software)

## License

MIT License
