@echo off
REM Excel Scrambler installer for Windows.
REM Creates a local venv, installs dependencies, and generates a launcher.
setlocal
cd /d "%~dp0"
echo == Excel Scrambler installer ==

REM --- find Python ------------------------------------------------------------
set "PYCMD="
py -3 -c "import sys" >nul 2>&1 && set "PYCMD=py -3"
if not defined PYCMD (
    python -c "import sys" >nul 2>&1 && set "PYCMD=python"
)
if not defined PYCMD (
    echo ERROR: Python 3 not found.
    echo Install it from https://www.python.org/downloads/ and be sure to check
    echo "Add python.exe to PATH" and keep "tcl/tk and IDLE" selected.
    pause
    exit /b 1
)
echo Using Python:
%PYCMD% --version

REM --- check tkinter ------------------------------------------------------------
%PYCMD% -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    echo ERROR: your Python was installed without tkinter ^(the GUI toolkit^).
    echo Re-run the Python installer, choose Modify, and enable "tcl/tk and IDLE".
    pause
    exit /b 1
)

REM --- venv + dependencies -------------------------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    %PYCMD% -m venv .venv
    if errorlevel 1 ( echo ERROR: could not create venv & pause & exit /b 1 )
)
echo Installing dependencies...
".venv\Scripts\python.exe" -m pip install --quiet --upgrade pip
".venv\Scripts\python.exe" -m pip install --quiet -r requirements.txt
".venv\Scripts\python.exe" -c "import tkinter, cryptography, openpyxl" >nul 2>&1
if errorlevel 1 ( echo ERROR: dependency check failed & pause & exit /b 1 )

REM --- launcher (no console window: pythonw) --------------------------------------
> "Excel Scrambler.bat" (
    echo @echo off
    echo start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0scrambler_app.py"
)

REM --- optional desktop shortcut ----------------------------------------------------
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Excel Scrambler.lnk'); $sc.TargetPath = '%~dp0.venv\Scripts\pythonw.exe'; $sc.Arguments = '\"%~dp0scrambler_app.py\"'; $sc.WorkingDirectory = '%~dp0'; $sc.Save()" >nul 2>&1

echo.
echo Done. Launch with "Excel Scrambler.bat" in this folder,
echo or the "Excel Scrambler" shortcut on your Desktop.
pause
