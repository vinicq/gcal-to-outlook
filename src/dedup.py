"""Remove duplicate [GCal] events from the Outlook Exchange calendar.

Run this once after a sync issue that created duplicate events.
It keeps the copy whose EntryID is recorded in sync_state.db and
deletes the extras. The database is updated when a stored ID is stale.

Usage:
  GCalSync.exe dedup   (via app.py)
  python src/dedup.py  (directly)
"""

import json
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

import win32com.client

OL_FOLDER_CALENDAR = 9

if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).resolve().parent.parent

CONFIG_PATH = ROOT / "config.json"
DB_PATH = ROOT / "sync_state.db"


def _find_exchange_calendar(ns, account_filter: str):
    exchange_candidate = None
    for store in ns.Stores:
        try:
            display = getattr(store, "DisplayName", "").lower()
            stype = getattr(store, "ExchangeStoreType", 0)

            if account_filter and (
                account_filter in display or display in account_filter
            ):
                print(f"Calendar: {store.DisplayName}")
                return store.GetDefaultFolder(OL_FOLDER_CALENDAR)

            if stype == 1 and exchange_candidate is None:
                exchange_candidate = store.GetDefaultFolder(OL_FOLDER_CALENDAR)
        except Exception:
            continue

    if exchange_candidate:
        print("Calendar: Exchange mailbox (type 1)")
        return exchange_candidate

    cal = ns.GetDefaultFolder(OL_FOLDER_CALENDAR)
    print("Calendar: default (fallback)")
    return cal


def main():
    if not CONFIG_PATH.exists():
        print("ERROR: config.json not found. Run the setup wizard first.")
        sys.exit(1)

    cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    tag_prefix = cfg["sync"].get("tag_prefix", "[GCal]")
    account = cfg.get("microsoft", {}).get("outlook_account", "").lower()

    if not DB_PATH.exists():
        print("ERROR: sync_state.db not found. Run a sync first.")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row

    # Load all known Outlook EntryIDs → google_id mapping from the DB
    rows = conn.execute("SELECT google_id, ms_id FROM event_map").fetchall()
    known_ids: dict[str, str] = {r["ms_id"]: r["google_id"] for r in rows}
    print(f"DB contains {len(known_ids)} known event IDs.")

    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    calendar = _find_exchange_calendar(ns, account)

    # Collect all tagged items; group by (subject, start minute) to detect duplicates
    items = calendar.Items
    items.Sort("[Start]")

    # key → list of EntryIDs
    groups: dict[tuple, list[str]] = defaultdict(list)

    print(f"\nScanning for '{tag_prefix}' events...")
    item = items.GetFirst()
    while item is not None:
        try:
            subject = getattr(item, "Subject", "") or ""
            if tag_prefix in subject:
                start = str(getattr(item, "Start", ""))
                key = (subject.strip(), start[:16])
                groups[key].append(item.EntryID)
        except Exception:
            pass
        item = items.GetNext()

    total_events = sum(len(v) for v in groups.values())
    dup_groups = {k: v for k, v in groups.items() if len(v) > 1}
    print(f"Found {total_events} tagged events in {len(groups)} unique slots.")
    print(f"Duplicate groups: {len(dup_groups)}")

    if not dup_groups:
        print("\nNo duplicates found. Calendar is clean.")
        conn.close()
        return

    removed = 0
    db_updated = 0

    for (subject, start), entry_ids in dup_groups.items():
        print(f"\n  [{len(entry_ids)}x] {subject[:70]} @ {start}")

        in_db = [eid for eid in entry_ids if eid in known_ids]
        not_in_db = [eid for eid in entry_ids if eid not in known_ids]

        if in_db:
            keep = in_db[0]
            to_delete = not_in_db + in_db[1:]
            print(f"    Keep (in DB): {keep[:32]}...")
        else:
            # DB was wiped - keep first copy and update DB entries that had old IDs
            keep = entry_ids[0]
            to_delete = entry_ids[1:]
            print(f"    Keep (DB update needed): {keep[:32]}...")

        for eid in to_delete:
            try:
                appt = ns.GetItemFromID(eid)
                appt.Delete()
                removed += 1
                print(f"    Deleted: {eid[:32]}...")
            except Exception as e:
                print(f"    Failed to delete {eid[:24]}...: {e}")

        # If the kept ID is not in DB, update any stale DB references
        if keep not in known_ids:
            for old_id in entry_ids:
                if old_id in known_ids:
                    google_id = known_ids[old_id]
                    conn.execute(
                        "UPDATE event_map SET ms_id=? WHERE google_id=?",
                        (keep, google_id),
                    )
                    db_updated += 1
            conn.commit()

    conn.close()
    print(f"\nDone. Removed {removed} duplicate events.")
    if db_updated:
        print(f"Updated {db_updated} stale EntryID references in sync_state.db.")
    print("Run a sync to verify everything is in order.")


if __name__ == "__main__":
    main()
