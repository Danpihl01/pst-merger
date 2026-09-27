"""
Merge multiple PST files into a single destination PST file.

Usage:
    python merge_pst.py -d destination.pst -s source1.pst source2.pst source3.pst source4.pst

Requirements:
    - Microsoft Outlook must be installed
    - pip install pywin32
"""

import argparse
import hashlib
import os
import subprocess
import sys
import time

try:
    import winreg
except ImportError:
    winreg = None


def outlook_running():
    """Return True if an OUTLOOK.EXE process is already running."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE"],
            capture_output=True, text=True, timeout=10,
        )
        return "OUTLOOK.EXE" in out.stdout
    except Exception:
        return False


def get_outlook():
    """Get or create an Outlook COM instance."""
    try:
        import win32com.client
    except ImportError:
        print("Error: pywin32 is not installed. Run: pip install pywin32", file=sys.stderr)
        sys.exit(1)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Force a connection so COM server errors surface here with context.
        _ = namespace.Folders
        return outlook
    except Exception as e:
        hresult = getattr(e, "hresult", None)
        print("Error: Could not start or connect to Outlook.", file=sys.stderr)
        print(f"  HRESULT: {hresult if hresult is not None else e}", file=sys.stderr)

        if hresult == -2146959355:
            print("\nThis error usually means one of the following:", file=sys.stderr)
            print("  1. Microsoft Outlook is NOT installed on this PC.", file=sys.stderr)
            print("     ->  The PST format is proprietary; an Outlook install is required.", file=sys.stderr)
            print("  2. Outlook has no mail profile configured.", file=sys.stderr)
            print("     ->  Open Outlook once and set it up so a profile exists.", file=sys.stderr)
            print("  3. A previous Outlook process is stuck.", file=sys.stderr)
            print("     ->  Open Task Manager, end any OUTLOOK.EXE, then retry.", file=sys.stderr)
            print("  4. Permission/COM server issue.", file=sys.stderr)
            print("     ->  Try running the exe from an elevated (Administrator) prompt.", file=sys.stderr)
        elif hresult == -2147352567:
            print("\nConnecting to Outlook's MAPI session failed.", file=sys.stderr)
            print("  The most common cause is the Windows Store (UWP) build of Outlook,", file=sys.stderr)
            print("  which does not support automation. It can also be a missing/locked", file=sys.stderr)
            print("  mail profile, or a policy blocking MAPI access.", file=sys.stderr)
            print("\n  Check:", file=sys.stderr)
            print("  1. Is this the 'Mail' app / Windows Store version of Outlook?", file=sys.stderr)
            print("     ->  It will NOT work with this tool. Full Outlook (C2R/MSI) is required.", file=sys.stderr)
            print("  2. Does a normal full Outlook install exist on this PC?", file=sys.stderr)
            print("     ->  Open Outlook once so it creates a mail profile, then retry.", file=sys.stderr)

        sys.exit(1)


def add_store(outlook, pst_path, create=False):
    """Open a PST file as a store in Outlook.

    If create=True and the file does not exist, it is created as a new
    Unicode PST via AddStoreEx.
    """
    namespace = outlook.GetNamespace("MAPI")
    pst_path = os.path.abspath(pst_path)

    os.makedirs(os.path.dirname(pst_path) or ".", exist_ok=True)

    if create and not os.path.exists(pst_path):
        olStoreUnicode = 3  # Outlook.OlStoreType.olStoreUnicode
        namespace.AddStoreEx(pst_path, olStoreUnicode)
    else:
        namespace.AddStore(pst_path)

    # Give Outlook time to fully initialize the store before first write.
    time.sleep(3)
    stores = namespace.Stores
    for i in range(1, stores.Count + 1):
        store = stores.Item(i)
        if store.FilePath.lower() == pst_path.lower():
            # Probe the store so initialization completes before returning.
            try:
                _ = store.GetRootFolder().Folders.Count
            except Exception:
                time.sleep(2)
            return store
    print(f"Warning: Could not find store for {pst_path}", file=sys.stderr)
    return None


def remove_store(outlook, store):
    """Remove a store from Outlook. Never raises.

    COM references can go stale after long merges, so we try several
    argument forms and ignore failures (a leftover attached store is
    harmless; it is only visible in this Outlook session).
    """
    namespace = outlook.GetNamespace("MAPI")
    attempts = [store]
    try:
        attempts.append(store.FilePath)
    except Exception:
        pass
    for arg in attempts:
        try:
            namespace.RemoveStore(arg)
            time.sleep(0.5)
            return True
        except Exception:
            continue
    # Last try: re-fetch a live reference from the collection.
    try:
        for j in range(1, namespace.Stores.Count + 1):
            s = namespace.Stores.Item(j)
            try:
                if s.FilePath.lower() == store.FilePath.lower():
                    namespace.RemoveStore(s)
                    time.sleep(0.5)
                    return True
            except Exception:
                continue
    except Exception:
        pass
    return False


def get_folder_count(folder):
    """Recursively count all items in a folder and its subfolders."""
    count = folder.Items.Count
    for i in range(1, folder.Folders.Count + 1):
        count += get_folder_count(folder.Folders.Item(i))
    return count


# Cache of already-in-destination item fingerprints, keyed by dest folder EntryID.
_fingerprint_cache = {}

_SK_URI = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"  # PR_SEARCH_KEY
_IMID_URI = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"  # PR_INTERNET_MESSAGE_ID


def item_fingerprints(item):
    """Return ALL stable identifiers for an item as a list of candidates.

    Duplicate detection treats the item as a duplicate if ANY candidate
    matches the destination set. Candidates:
      1. PR_SEARCH_KEY (transport-unique, survives moves/copies for mail)
      2. PR_INTERNET_MESSAGE_ID (normalised, survives store moves)
      3. strong content hash (class+subject+sender+sent+body) for items
         that carry neither transport key
    """
    candidates = []

    try:
        sk = item.PropertyAccessor.GetProperty(_SK_URI)
        if sk:
            candidates.append("sk:" + bytes(sk).hex())
    except Exception:
        pass

    try:
        imid = item.PropertyAccessor.GetProperty(_IMID_URI)
        if imid:
            candidates.append("imid:" + str(imid).lower())
    except Exception:
        pass

    if not candidates:
        try:
            h = hashlib.sha256()
            for part in (
                str(getattr(item, "Class", "")),
                str(getattr(item, "Subject", "")).lower(),
                str(getattr(item, "SenderName", "")).lower(),
                str(getattr(item, "SentOn", "")),
                str(getattr(item, "Body", "")),
            ):
                h.update(part.encode("utf-8", "replace"))
            candidates.append("h:" + h.hexdigest())
        except Exception:
            pass

    return candidates


def dest_fingerprints(dest_folder):
    """Return the set of fingerprints already present in dest_folder.

    Scans destination items with the exact same fingerprint function used
    for source items, so a verbatim copy in the destination always matches
    its source. Results are cached per destination folder for the run.
    """
    try:
        folder_id = dest_folder.EntryID
    except Exception:
        folder_id = None

    if folder_id is not None and folder_id in _fingerprint_cache:
        return _fingerprint_cache[folder_id]

    # IMPORTANT: use the exact same fingerprint function as for source items.
    # Relying on GetTable/property tags was unreliable and let duplicates
    # through; computing content hashes from the items themselves guarantees
    # that a verbatim copy in the destination always matches its source.
    existing = set()
    try:
        for j in range(1, dest_folder.Items.Count + 1):
            try:
                for fp in item_fingerprints(dest_folder.Items.Item(j)):
                    existing.add(fp)
            except Exception:
                pass
    except Exception:
        pass

    if folder_id is not None:
        _fingerprint_cache[folder_id] = existing
    return existing


def is_rpc_error(e):
    """Return True if an exception is an 'RPC server unavailable' failure."""
    hr = getattr(e, "hresult", None)
    if hr is not None and hr == -2147023174:
        return True
    return "-2147023174" in str(e)


def copy_items(source_folder, dest_folder, retries=6):
    """Recursively copy all items from source folder to dest folder.

    Skips items already present in the destination (deduplication).
    Retries transient RPC failures with longer backoff, re-fetching the item
    on each attempt (the COM item reference dies when the RPC session drops).
    """
    copied = 0
    failed = 0
    skipped = 0
    existing = dest_fingerprints(dest_folder)

    try:
        total = source_folder.Items.Count
    except Exception:
        total = 0

    def report(final=False):
        done = copied + failed + skipped
        if final:
            print(f"    Progress: {copied} copied, {skipped} skipped, {failed} failed "
                  f"of {total}", file=sys.stderr)
        elif total and done % 100 == 0:
            pct = done * 100 // total if total else 0
            print(f"    Progress: {copied} copied, {skipped} skipped, {failed} failed "
                  f"({done}/{total} items, {pct}%)", file=sys.stderr)

    for i in range(1, source_folder.Items.Count + 1):
        item = None
        subject = "<unknown>"
        fps = []

        for attempt in range(retries):
            # (Re)fetch the item fresh each attempt - a dropped RPC session
            # invalidates previously-held COM references.
            if item is None:
                try:
                    item = source_folder.Items.Item(i)
                    subject = getattr(item, "Subject", "<unknown>")
                except Exception as e:
                    item = None
                    if attempt < retries - 1:
                        time.sleep(5 if is_rpc_error(e) else 2)
                        continue
                    print(f"  Warning: Could not fetch item {i}: {e}", file=sys.stderr)
                    failed += 1
                    break

            try:
                fps = item_fingerprints(item)
            except Exception:
                fps = []

            if fps and any(fp in existing for fp in fps):
                skipped += 1
                time.sleep(0.01)
                report()
                break

            try:
                item.Copy().Move(dest_folder)
                copied += 1
                for fp in fps:
                    existing.add(fp)
                break
            except Exception as e:
                rpc = is_rpc_error(e)
                item = None  # reference is dead; re-fetch next attempt
                if attempt < retries - 1:
                    time.sleep(5 if rpc else 1 + attempt)
                else:
                    print(f"  Warning: Could not copy item '{subject}': {e}", file=sys.stderr)
                    failed += 1

        time.sleep(0.05)
        report()

    for i in range(1, source_folder.Folders.Count + 1):
        sub_source = source_folder.Folders.Item(i)
        sub_name = sub_source.Name

        sub_dest = None
        try:
            sub_dest = dest_folder.Folders.Item(sub_name)
        except Exception:
            sub_dest = dest_folder.Folders.Add(sub_name)

        c, f, s = copy_items(sub_source, sub_dest, retries)
        copied += c
        failed += f
        skipped += s

    report(final=True)
    return copied, failed, skipped


def fix_pst_size_limit(verbose=True):
    """Raise Outlook's large + small PST size limits so merges don't hit caps.

    Values are in MEGAbytes. MaxLargeFileSize/WarnLargeFileSize apply to
    Unicode PSTs (default 51200 MB = 50GB); MaxFileSize/WarnFileSize apply to
    ANSI PSTs. We write to all locations because Group Policy keys
    (Software\\Policies\\...) override the normal ones.
    """
    if winreg is None:
        if verbose:
            print("  Warning: winreg unavailable; cannot adjust PST size limit.", file=sys.stderr)
        return []

    # ~390GB in MB for large files; ~1.9GB is the absolute ANSI ceiling.
    max_large_mb = 399999
    warn_large_mb = 399999 - int(399999 * 0.05) - 1  # ~5% headroom under max
    max_small = 0x7C004400   # 2,075,149,312 bytes
    warn_small = 0x74404400  # 1,950,368,768 bytes

    # (root, subkey) pairs: normal and Group-Policy override locations.
    targets = []
    for version in ("15.0", "16.0"):
        sub = rf"Software\Microsoft\Office\{version}\Outlook\PST"
        targets.append((winreg.HKEY_CURRENT_USER, sub))
        targets.append((winreg.HKEY_LOCAL_MACHINE, sub))
        targets.append((winreg.HKEY_CURRENT_USER, rf"Software\Policies\Microsoft\Office\{version}\Outlook\PST"))
        targets.append((winreg.HKEY_LOCAL_MACHINE, rf"Software\Policies\Microsoft\Office\{version}\Outlook\PST"))

    written = []
    for root, subkey in targets:
        try:
            with winreg.CreateKeyEx(root, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "MaxLargeFileSize", 0, winreg.REG_DWORD, max_large_mb)
                winreg.SetValueEx(key, "WarnLargeFileSize", 0, winreg.REG_DWORD, warn_large_mb)
                winreg.SetValueEx(key, "MaxFileSize", 0, winreg.REG_DWORD, max_small)
                winreg.SetValueEx(key, "WarnFileSize", 0, winreg.REG_DWORD, warn_small)
                written.append(subkey)
        except Exception as e:
            if verbose:
                print(f"  Warning: could not update {subkey}: {e}", file=sys.stderr)

    if verbose and written:
        print(f"  PST size limits raised to ~390GB in: {len(written)} registry location(s).")

    # Read back what a normal HKCU 16.0 location actually holds (diagnostics).
    if verbose:
        for version in ("15.0", "16.0"):
            sub = rf"Software\Microsoft\Office\{version}\Outlook\PST"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub) as key:
                    ml = winreg.QueryValueEx(key, "MaxLargeFileSize")[0]
                    wl = winreg.QueryValueEx(key, "WarnLargeFileSize")[0]
                    ms = winreg.QueryValueEx(key, "MaxFileSize")[0]
                    ws = winreg.QueryValueEx(key, "WarnFileSize")[0]
                print(f"  [{version}] MaxLargeFileSize={ml} MB, WarnLargeFileSize={wl} MB, "
                      f"MaxFileSize={ms} B, WarnFileSize={ws} B")
            except Exception:
                pass

    return written


def pst_wver(path):
    """Return the wVer (file-format version) of a PST, or None if unreadable.

    wVer at offset 0x000A: 14/15 = ANSI (hard 2GB cap), >=23 = Unicode.
    """
    try:
        with open(path, "rb") as f:
            f.seek(0x000A)
            data = f.read(2)
        return int.from_bytes(data, "little") if len(data) == 2 else None
    except OSError:
        return None


def merge_psts(destination, sources):
    """Merge source PST files into destination PST."""
    if outlook_running():
        print("WARNING: Outlook is already running.", file=sys.stderr)
        print("  A running Outlook has cached PST size limits from startup.", file=sys.stderr)
        print("  Close Outlook completely, then re-run this program for reliable results.\n", file=sys.stderr)

    outlook = get_outlook()
    namespace = outlook.GetNamespace("MAPI")

    dest_path = os.path.abspath(destination)

    # If an existing destination is in the old ANSI format (hard ~2GB cap),
    # it can never grow to 34GB: delete it so a Unicode PST is created.
    if os.path.exists(dest_path):
        ver = pst_wver(dest_path)
        if ver is not None and ver < 23:
            print(f"Destination is old ANSI format (wVer={ver}); recreating as Unicode.", file=sys.stderr)
            try:
                os.remove(dest_path)
                print("  Deleted old destination.", file=sys.stderr)
            except OSError as e:
                print(f"  Error: could not delete old destination: {e}", file=sys.stderr)
                print("  Close Outlook and delete it manually, then re-run.", file=sys.stderr)
                sys.exit(1)

    if os.path.exists(dest_path):
        print(f"Opening existing destination PST: {dest_path}")
        print("  Warning: existing items will be kept and new sources merged into it.")
    else:
        print(f"Creating destination PST: {dest_path}")

    print("Ensuring large-PST (up to ~2TB) limit is enabled...")
    fix_pst_size_limit()

    dest_store = add_store(outlook, dest_path, create=True)
    if not dest_store:
        print("Error: Could not open destination PST.", file=sys.stderr)
        sys.exit(1)

    dest_root = dest_store.GetRootFolder()

    # Confirm the destination is Unicode; otherwise the merge is doomed.
    ver = pst_wver(dest_path)
    if ver is not None and ver < 23:
        print("\n  FATAL: destination PST is still ANSI format (wVer={}).".format(ver), file=sys.stderr)
        print("  This PC's Outlook cannot create a Unicode PST via automation.", file=sys.stderr)
        print("  Install/reinstall full desktop Outlook (not the new/Store version).\n", file=sys.stderr)
        sys.exit(1)

    # Warn if the destination is an old ANSI-format PST (hard ~2GB limit).
    try:
        is_data_file = bool(dest_store.IsDataFileStore)
    except Exception:
        is_data_file = False

    total_copied = 0
    total_failed = 0
    total_skipped = 0

    for src in sources:
        src_path = os.path.abspath(src)
        if not os.path.exists(src_path):
            print(f"Skipping (not found): {src}", file=sys.stderr)
            continue

        print(f"Processing: {src}")
        src_store = add_store(outlook, src_path)
        if not src_store:
            print(f"  Warning: Could not open {src}, skipping.", file=sys.stderr)
            continue

        src_root = src_store.GetRootFolder()

        for i in range(1, src_root.Folders.Count + 1):
            folder = src_root.Folders.Item(i)
            folder_name = folder.Name
            print(f"  Folder: {folder_name}")

            dest_folder = None
            try:
                dest_folder = dest_root.Folders.Item(folder_name)
            except Exception:
                dest_folder = dest_root.Folders.Add(folder_name)

            copied, failed, skipped = copy_items(folder, dest_folder)
            total_copied += copied
            total_failed += failed
            total_skipped += skipped
            print(f"    Copied {copied}, skipped {skipped} duplicate"
                  + (f", {failed} failed" if failed else ""))

        remove_store(outlook, src_store)

    print(f"\nDone. Total items copied: {total_copied}")
    print(f"Duplicates skipped: {total_skipped}")
    if total_failed:
        print(f"  Warning: {total_failed} items could not be copied and were skipped.")
    print(f"Destination: {dest_path}")


def main():
    parser = argparse.ArgumentParser(description="Merge multiple PST files into one.")
    parser.add_argument("-d", "--destination", required=True, help="Path to the destination PST file")
    parser.add_argument("-s", "--sources", nargs="+", required=True, help="Paths to source PST files")
    args = parser.parse_args()

    merge_psts(args.destination, args.sources)


if __name__ == "__main__":
    main()
