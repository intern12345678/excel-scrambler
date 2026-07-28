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

## Updating

Click **⟳ Check for updates** on the app's main menu — it pulls the newest
code from GitHub, installs any new dependencies, and offers to restart into
the new version. Equivalently, run the updater in the repo folder (this route
also regenerates the launcher/desktop shortcut):

- **macOS / Linux**: `./update.sh`
- **Windows**: `update.bat`

## Using it

- **Scramble** — pick an `.xlsx`/`.xlsm` file, pick the sheet and column(s)
  (every column in the sheet is listed and selectable, whether or not it has
  a header), set how many rows from the top to skip (labels/headers —
  default 1, set 0 to scramble everything or higher to protect multi-row
  headers), click Scramble. A new `<name>_scrambled.xlsx` is written next to
  the original (the original is never touched) and a Fernet key is saved to
  your key library along with the file name, sheet, columns, and timestamp.
- **Sharing one key across files** — by default every scramble generates a
  brand-new key, but the Scramble screen lets you pick **"Reuse an existing
  key"** and choose any key from your library. Each file still gets its own
  library entry (so unscrambling auto-matches per file), they just all carry
  the same key. Combined with match-preserving mode, the same value scrambles
  to the same token *across files*, so cross-file matching keeps working.
- **Match-preserving mode** (toggle on the Scramble screen) — normally every
  cell gets a unique random token, even for identical values. With this mode
  on, identical values scramble to the *identical* token, so other programs
  can still group/match rows by the scrambled value. Trade-off: anyone can
  see which cells are equal to each other (though not what they say). Either
  kind of file unscrambles the same way.
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
  package; every scramble run generates a brand-new key. Match-preserving
  mode uses a Fernet-compatible deterministic variant (IV derived from the
  cell value) — tokens still decrypt with the ordinary key.
- Wrong key on unscramble → clear error, no output file written.
