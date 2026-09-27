# PST Merger

A Windows command-line tool that merges multiple Outlook `.pst` archive files into a
single destination `.pst`, preserving the folder hierarchy and skipping items that
already exist in the destination.

It drives **full desktop Microsoft Outlook** through its COM/MAPI automation
interface, which is the only supported way to read and write PST files.

---

## What it does

- Creates a new Unicode destination PST, or merges into an existing one.
- Walks the top-level folders of every source PST, creating matching folders in the
  destination as needed, and recursively copies every item and subfolder.
- **Deduplicates** items: anything already present in the destination folder is
  skipped rather than copied again, so re-running the tool is safe.
- Raises Outlook's PST size limits (up to ~390 GB) before merging, so large
  archives do not hit the default 50 GB cap.
- Detects and replaces an old **ANSI-format** destination PST (hard ~2 GB limit)
  with a new Unicode PST.
- Retries transient "RPC server unavailable" COM failures with backoff, which is
  common on long merges.
- Prints progress and a per-run summary to the console.

### How duplicates are detected

Each item is fingerprinted using the first identifier available:

1. `PR_SEARCH_KEY` (`0x300B0102`) — transport-unique, survives moves and copies.
2. `PR_INTERNET_MESSAGE_ID` (`0x1035001E`) — normalised, survives store moves.
3. A SHA-256 content hash of class + subject + sender + sent date + body, used only
   for items that carry neither of the above.

An item is considered a duplicate if **any** of its fingerprints is already present
in the destination folder. Destination fingerprints are computed with the exact same
function and cached per folder for the duration of the run.

---

## Requirements

- Windows (Windows-only: it relies on `winreg`, `tasklist`, and Outlook COM).
- **Full desktop Microsoft Outlook** (C2R or MSI). The Windows Store / "Mail" (UWP)
  version of Outlook does **not** support automation and will not work.
- Outlook must be installed and have been opened at least once so a mail profile
  exists.
- Python 3.8+ and `pywin32` if running from source.

```powershell
pip install pywin32
```

The prebuilt `dist\pst-merge.exe` already bundles everything and needs only Outlook.

---

## Usage

### Prebuilt executable

```powershell
dist\pst-merge.exe -d destination.pst -s source1.pst source2.pst source3.pst
```

### From source

```powershell
python merge_pst.py -d destination.pst -s source1.pst source2.pst source3.pst
```

### Options

| Flag | Required | Description |
| --- | --- | --- |
| `-d`, `--destination` | Yes | Path to the destination `.pst` file. Created if missing. |
| `-s`, `--sources` | Yes | One or more source `.pst` files to merge in. |

### Examples

Merge four archives into a new file:

```powershell
python merge_pst.py -d D:\Archive\all_mail.pst -s D:\Backup\2023.pst D:\Backup\2024.pst D:\Backup\2025.pst D:\Backup\2026.pst
```

Add another archive to an existing destination later (existing items are kept, new
ones are added, duplicates are skipped):

```powershell
python merge_pst.py -d D:\Archive\all_mail.pst -s D:\Backup\new.pst
```

Use absolute paths if you want to be explicit — relative paths are resolved against
the current working directory.

---

## Practical notes and behaviour

- **Close Outlook before running.** A running `OUTLOOK.EXE` has already cached the
  old PST size limits and can hold locks on the files. The tool prints a warning if
  it detects Outlook running, but does not stop it for you. A leftover attached
  store after a crash is harmless — it only exists in that Outlook session.
- **Administrator rights** may be required. The tool raises size limits under both
  `HKEY_CURRENT_USER` and `HKEY_LOCAL_MACHINE` (and the matching
  `Software\Policies\...` Group Policy locations for Office 15.0 and 16.0). Machine
  and policy keys usually need elevation, so run from an Administrator prompt if
  you want the limits applied system-wide.
- **Registry changes are permanent.** `fix_pst_size_limit()` writes
  `MaxLargeFileSize`, `WarnLargeFileSize`, `MaxFileSize`, and `WarnFileSize` on every
  run. It prints the resulting values afterwards.
- **An existing ANSI destination is deleted automatically.** If the destination
  exists and is ANSI format (`wVer < 23`), the tool removes it and recreates it as
  Unicode, because an ANSI file can never grow past ~2 GB. Back it up first if you
  care about it. If deletion fails, close Outlook and delete it manually.
- **Folder names are matched, not merged by structure.** Source top-level folders are
  mapped to destination folders by name; subfolders are matched the same way
  recursively. Renamed or differently-named folders are created as new folders.
- **Sources are attached read/write as normal stores.** Take a backup of your source
  PSTs before a large first run, in case anything interrupts the process.
- **Missing or unopenable sources are skipped** with a warning rather than aborting
  the whole run.
- **Failures are non-fatal.** Items that cannot be copied after retries are counted
  and reported in the final summary.

---

## Output

A typical run looks like:

```
WARNING: Outlook is already running.
  ...
Creating destination PST: D:\Archive\all_mail.pst
Ensuring large-PST (up to ~2TB) limit is enabled...
  PST size limits raised to ~390GB in: 4 registry location(s).
  [15.0] MaxLargeFileSize=399999 MB, ...
  [16.0] MaxLargeFileSize=399999 MB, ...
Processing: D:\Backup\2023.pst
  Folder: Inbox
    Progress: 100 copied, 0 skipped, 0 failed (100/500 items, 20%)
    ...
    Copied 500, skipped 0 duplicate
    Progress: 500 copied, 0 skipped, 0 failed of 500

Done. Total items copied: 500
Duplicates skipped: 0
Destination: D:\Archive\all_mail.pst
```

Progress and warnings go to stderr; status messages go to stdout.

---

## Rebuilding the executable

The project uses PyInstaller with the spec file in the repository root:

```powershell
pip install pyinstaller
pyinstaller pst-merge.spec
```

The result is written to `dist\pst-merge.exe`.

---

## Troubleshooting

| Symptom | Cause / fix |
| --- | --- |
| `Could not start or connect to Outlook` with `HRESULT: -2146959355` | Outlook is not installed, no mail profile exists, or a stuck `OUTLOOK.EXE`. Open Outlook once to create a profile; end all Outlook processes in Task Manager and retry. |
| `HRESULT: -2147352567` | The Windows Store (UWP) "Mail" version of Outlook was found. It does not support automation — install full desktop Outlook. Also check for a missing/locked mail profile or a policy blocking MAPI. |
| `pywin32 is not installed` | Run `pip install pywin32`. |
| `Could not open destination PST` | The file is locked, read-only, or the path is not writable. Close Outlook and check permissions. |
| `Could not find store for <path>` | Outlook did not attach the PST. Close Outlook, verify the file is a valid `.pst` and not open in another profile. |
| `FATAL: destination PST is still ANSI format` | This Outlook installation cannot create Unicode PSTs via automation. Reinstall full desktop Outlook. |
| `Warning: Could not copy item '...'` | Usually a transient RPC drop. Those items are reported as failed; re-running the tool will pick them up, since the rest of the archive is already deduplicated. |
| `Warning: could not update Software\...` | Registry key could not be written. Run as Administrator, or where a Group Policy blocks the location. |

---

## Project layout

```
merge_pst.py     the tool
pst-merge.spec   PyInstaller build spec
dist/            prebuilt pst-merge.exe
build/           PyInstaller intermediate output
```
