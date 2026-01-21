@echo off
REM Intune wrapper for Dell Bloatware Removal
REM This wrapper adds Intune logging without modifying the core script
echo Starting Dell Bloatware Removal via Intune...
REM Create Intune log directory
if not exist "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs" mkdir "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
REM Log start to Intune location
echo %date% %time% [INFO] Starting Dell Bloatware Removal via Intune >> "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval.log"
REM Run the original, proven script without modification
PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0dellBloatware.ps1"
REM Capture exit code
set SCRIPT_EXIT_CODE=%ERRORLEVEL%
REM Log completion to Intune location
if %SCRIPT_EXIT_CODE%==0 (
    echo %date% %time% [INFO] Dell Bloatware Removal completed successfully >> "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval.log"
) else (
    echo %date% %time% [ERROR] Dell Bloatware Removal failed with exit code %SCRIPT_EXIT_CODE% >> "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval.log"
)
REM Copy main log to Intune location for easier access (backup copy)
if exist "C:\ComprehensiveDellBloatwareRemoval.log" (
    copy "C:\ComprehensiveDellBloatwareRemoval.log" "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\DellBloatwareRemoval_Main.log" >nul 2>&1
)
REM Exit with original script's exit code
exit /b %SCRIPT_EXIT_CODE%
