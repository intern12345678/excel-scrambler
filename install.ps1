# Excel Scrambler installer for Windows.
# Installs Python itself if missing (winget, else direct download from
# python.org), then creates a venv, installs dependencies, and generates
# a launcher + desktop shortcut. Run via install.bat or:
#   powershell -ExecutionPolicy Bypass -File install.ps1

$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Here
$PythonVersion = "3.13.7"   # used only for the direct-download fallback

Write-Host "== Excel Scrambler installer ==" -ForegroundColor Cyan

function Invoke-Py {
    # Runs a python command described as @{Exe=...; Args=@(...)} plus extra args.
    param($Cmd, [string[]]$Rest)
    & $Cmd.Exe @($Cmd.Args + $Rest)
}

function Test-PythonCmd {
    # $true only for a real Python 3.10+ (not the Microsoft Store alias).
    param($Cmd)
    try {
        $out = Invoke-Py $Cmd @("-c",
            "import sys; print(sys.version_info[0]*100+sys.version_info[1])") 2>$null
        return ($LASTEXITCODE -eq 0 -and [int]($out | Select-Object -Last 1) -ge 310)
    } catch { return $false }
}

function Find-Python {
    $candidates = @(
        @{Exe = "py";     Args = @("-3")},
        @{Exe = "python"; Args = @()}
    )
    foreach ($c in $candidates) {
        if (Get-Command $c.Exe -ErrorAction SilentlyContinue) {
            if (Test-PythonCmd $c) { return $c }
        }
    }
    # look in the default install locations (PATH may be stale right after install)
    $roots = @("$env:LocalAppData\Programs\Python", "$env:ProgramFiles") |
             Where-Object { Test-Path $_ }
    foreach ($root in $roots) {
        $exes = Get-ChildItem -Path $root -Filter python.exe -Recurse -Depth 2 `
                    -ErrorAction SilentlyContinue |
                Where-Object { $_.DirectoryName -match "Python3\d+" } |
                Sort-Object FullName -Descending
        foreach ($exe in $exes) {
            $c = @{Exe = $exe.FullName; Args = @()}
            if (Test-PythonCmd $c) { return $c }
        }
    }
    return $null
}

# --- 1. find or install Python ------------------------------------------------
$Py = Find-Python
if (-not $Py) {
    Write-Host "Python not found - installing it now..." -ForegroundColor Yellow
    $installed = $false

    # Preferred: winget (ships with Windows 10/11)
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "Installing Python via winget..."
        winget install -e --id Python.Python.3.13 --silent `
            --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -eq 0) { $installed = $true }
        else { Write-Host "winget install failed, falling back to direct download." }
    }

    # Fallback: official installer from python.org
    if (-not $installed) {
        $arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "arm64" } else { "amd64" }
        $url = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-$arch.exe"
        $tmp = Join-Path $env:TEMP "python-installer.exe"
        Write-Host "Downloading $url ..."
        Invoke-WebRequest -Uri $url -OutFile $tmp
        Write-Host "Running the Python installer (silent, per-user)..."
        Start-Process -FilePath $tmp -Wait -ArgumentList `
            "/quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1 Include_launcher=1"
        Remove-Item $tmp -ErrorAction SilentlyContinue
    }

    $Py = Find-Python
    if (-not $Py) {
        throw "Python was installed but cannot be located. Open a NEW terminal and re-run install.bat."
    }
}
$ver = Invoke-Py $Py @("--version")
Write-Host "Using $ver ($($Py.Exe))"

# --- 2. tkinter check -----------------------------------------------------------
Invoke-Py $Py @("-c", "import tkinter") 2>$null | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "This Python lacks tkinter (the GUI toolkit). Re-run the Python installer, choose Modify, and enable 'tcl/tk and IDLE'."
}

# --- 3. venv + dependencies -------------------------------------------------------
$VenvPy = Join-Path $Here ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPy)) {
    Write-Host "Creating virtual environment..."
    Invoke-Py $Py @("-m", "venv", ".venv")
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPy)) {
        throw "Could not create the virtual environment."
    }
}
Write-Host "Installing dependencies..."
& $VenvPy -m pip install --quiet --upgrade pip
& $VenvPy -m pip install --quiet -r requirements.txt
& $VenvPy -c "import tkinter, cryptography, openpyxl"
if ($LASTEXITCODE -ne 0) { throw "Dependency check failed." }

# --- 4. launcher + desktop shortcut --------------------------------------------------
$PyW = Join-Path $Here ".venv\Scripts\pythonw.exe"
@"
@echo off
start "" "$PyW" "$Here\scrambler_app.py"
"@ | Set-Content -Path (Join-Path $Here "Excel Scrambler.bat") -Encoding ASCII

try {
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Desktop")) "Excel Scrambler.lnk"))
    $sc.TargetPath = $PyW
    $sc.Arguments = "`"$Here\scrambler_app.py`""
    $sc.WorkingDirectory = $Here
    $sc.Save()
    $shortcut = $true
} catch { $shortcut = $false }

Write-Host ""
Write-Host "Done." -ForegroundColor Green
Write-Host "Launch with 'Excel Scrambler.bat' in this folder$(if ($shortcut) { ", or the 'Excel Scrambler' shortcut on your Desktop" })."
