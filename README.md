# Excel Scrambler

A small desktop app that encrypts ("scrambles") chosen columns of an Excel
workbook with a Fernet key, keeps the key in a local library, and can restore
the original values later. Works on macOS, Windows, and Linux.

## Install

Clone the repo and run the install script for your OS — it installs everything
needed, **including Python itself if it's missing**, then creates a private
virtual environment, installs all dependencies, and generates a
double-clickable launcher.

**macOS / Linux**
```
git clone https://github.com/intern12345678/excel-scrambler.git
cd excel-scrambler
./install.sh
```
Launch: double-click **Excel Scrambler.app** (macOS) or run
`./excel-scrambler.sh` (Linux).

**Windows**
```
git clone https://github.com/intern12345678/excel-scrambler.git
cd excel-scrambler
install.bat
```
Launch: **Excel Scrambler.bat** in the repo folder, or the
**Excel Scrambler** shortcut the installer puts on your Desktop.

No prerequisites beyond git: if Python 3.10+ (with tkinter) isn't already on
the machine, the installer gets it for you — via winget or the official
python.org installer on Windows, Homebrew on macOS, apt/dnf on Linux. If it
truly can't (no package manager at all), it prints exact instructions.

## Using it

- **Scramble** — pick an `.xlsx`/`.xlsm` file, pick the sheet and column(s),
  optionally exempt the first (label) row, click Scramble. A new
  `<name>_scrambled.xlsx` is written next to the original (the original is
  never touched) and a fresh Fernet key is saved to your key library along
  with the file name, sheet, columns, and timestamp.
- **Unscramble** — pick the scrambled file, then pick the matching key from
  the library (the app auto-selects the key whose recorded output matches the
  file and auto-fills the columns/header settings from the key). Writes
  `<name>_unscrambled.xlsx`.
- **Key library** — browse all saved keys, reveal a scrambled file in
  Finder/Explorer, or permanently delete a key (with a warning — a deleted
  key's file can never be unscrambled).

## Details

- Keys live in `~/.excel-scrambler/keys/` as one JSON file per scramble.
  Deleting a key makes the data in its file unrecoverable.
- Scrambled cells carry an `XSCRAMBLE1:` prefix so the app recognizes them,
  refuses to double-scramble, and skips non-scrambled cells on restore.
- Cell types survive the round trip (numbers come back as numbers, dates as
  dates) — values are JSON-typed before encryption.
- Encryption is Fernet (AES-128-CBC + HMAC-SHA256) from the `cryptography`
  package; every scramble run generates a brand-new key.
- Wrong key on unscramble → clear error, no output file written.
