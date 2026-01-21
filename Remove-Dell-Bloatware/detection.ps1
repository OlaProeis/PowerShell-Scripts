# Intune detection script for Dell Bloatware Removal - "Run Once" Version
# Returns exit code 1 if script has NOT run yet (needs remediation)
# Returns exit code 0 if script has already run (compliant - don't retry)
try {
    # Check if the script has already been executed
    $markerFile = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Executed.marker"
    
    if (Test-Path $markerFile) {
        # Script has already run - report as compliant to prevent retry
        $markerContent = Get-Content $markerFile -ErrorAction SilentlyContinue
        Write-Output "Dell bloatware removal script already executed on: $markerContent"
        Write-Output "System compliant - no retry needed"
        exit 0  # Compliant - don't run again
    } else {
        # Script has not run yet - trigger execution
        Write-Output "Dell bloatware removal script has not been executed yet"
        Write-Output "Remediation needed - will execute removal script"
        exit 1  # Not compliant - needs remediation
    }

} catch {
    # If detection fails, assume remediation is needed to be safe
    Write-Output "Detection script error - assuming remediation needed"
    exit 1
}