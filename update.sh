#!/bin/bash
# Excel Scrambler updater for macOS / Linux.
# Pulls the latest version from GitHub and re-runs the installer so any new
# dependencies are installed and the launcher is regenerated.
set -e
cd "$(dirname "$0")"

echo "== Excel Scrambler updater =="
git pull --ff-only
./install.sh
echo
echo "Update complete — you're on the latest version."
