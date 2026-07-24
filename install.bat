@echo off
REM Excel Scrambler installer for Windows.
REM Thin wrapper: all logic lives in install.ps1 (installs Python itself if
REM missing, then venv + dependencies + launcher + desktop shortcut).
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
    echo.
    echo Installation failed - see the message above.
)
pause
