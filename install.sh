#!/bin/bash
# Excel Scrambler installer for macOS / Linux.
# Creates a local venv, installs dependencies, and generates a launcher.
set -e
cd "$(dirname "$0")"
HERE="$(pwd)"

echo "== Excel Scrambler installer =="

# --- find (or install) a python3 --------------------------------------------
PY="$(command -v python3 || true)"
if [ -z "$PY" ]; then
    echo "python3 not found - installing it..."
    if [ "$(uname)" = "Darwin" ] && command -v brew >/dev/null; then
        brew install python
    elif command -v apt-get >/dev/null; then
        sudo apt-get update && sudo apt-get install -y python3 python3-venv python3-tk
    elif command -v dnf >/dev/null; then
        sudo dnf install -y python3 python3-tkinter
    else
        echo "ERROR: no supported package manager found. Install Python 3.10+"
        echo "manually (https://www.python.org/downloads/) and re-run."
        exit 1
    fi
    PY="$(command -v python3 || true)"
    [ -z "$PY" ] && { echo "ERROR: Python install failed."; exit 1; }
fi
echo "Using $PY ($($PY --version))"

# --- check (or install) tkinter ----------------------------------------------
if ! "$PY" -c "import tkinter" 2>/dev/null; then
    echo "tkinter missing - installing it..."
    if [ "$(uname)" = "Darwin" ] && command -v brew >/dev/null; then
        PYVER="$($PY -c 'import sys; print(f"{sys.version_info[0]}.{sys.version_info[1]}")')"
        brew install "python-tk@$PYVER"
    elif command -v apt-get >/dev/null; then
        sudo apt-get install -y python3-tk
    elif command -v dnf >/dev/null; then
        sudo dnf install -y python3-tkinter
    fi
    if ! "$PY" -c "import tkinter" 2>/dev/null; then
        echo "ERROR: could not install tkinter automatically. Install it for"
        echo "your Python and re-run this script."
        exit 1
    fi
fi

# --- venv + dependencies -----------------------------------------------------
if [ ! -x .venv/bin/python ]; then
    echo "Creating virtual environment..."
    "$PY" -m venv .venv
fi
echo "Installing dependencies..."
.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet -r requirements.txt
.venv/bin/python -c "import tkinter, cryptography, openpyxl" \
    || { echo "ERROR: dependency check failed"; exit 1; }

# --- launcher ----------------------------------------------------------------
if [ "$(uname)" = "Darwin" ]; then
    APP="$HERE/Excel Scrambler.app"
    mkdir -p "$APP/Contents/MacOS"
    cat > "$APP/Contents/Info.plist" << PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key><string>Excel Scrambler</string>
    <key>CFBundleDisplayName</key><string>Excel Scrambler</string>
    <key>CFBundleIdentifier</key><string>local.excel-scrambler</string>
    <key>CFBundleVersion</key><string>1.0</string>
    <key>CFBundleShortVersionString</key><string>1.0</string>
    <key>CFBundlePackageType</key><string>APPL</string>
    <key>CFBundleExecutable</key><string>launcher</string>
</dict>
</plist>
PLIST
    cat > "$APP/Contents/MacOS/launcher" << LAUNCH
#!/bin/bash
exec "$HERE/.venv/bin/python" "$HERE/scrambler_app.py"
LAUNCH
    chmod +x "$APP/Contents/MacOS/launcher"
    echo
    echo "Done. Double-click 'Excel Scrambler.app' in $HERE to launch"
    echo "(or drag it to the Dock)."
else
    cat > "$HERE/excel-scrambler.sh" << LAUNCH
#!/bin/bash
exec "$HERE/.venv/bin/python" "$HERE/scrambler_app.py"
LAUNCH
    chmod +x "$HERE/excel-scrambler.sh"
    echo
    echo "Done. Run ./excel-scrambler.sh to launch."
fi
