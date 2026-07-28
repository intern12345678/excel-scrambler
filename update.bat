@echo off
REM Excel Scrambler updater for Windows.
REM Pulls the latest version from GitHub and re-runs the installer so any new
REM dependencies are installed and the launcher/shortcut are regenerated.
cd /d "%~dp0"
echo == Excel Scrambler updater ==
git pull --ff-only
if errorlevel 1 (
    echo.
    echo Update failed - could not pull from GitHub. See the message above.
    pause
    exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
    echo.
    echo Update failed - see the message above.
    pause
    exit /b 1
)
echo.
echo Update complete - you're on the latest version.
pause
