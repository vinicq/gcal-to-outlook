"""Main sync program: Google Calendar -> Microsoft / Teams.

Modes:
  python sync.py login    -> authenticate both Google and Microsoft, then exit
  python sync.py once      -> run one sync cycle and exit
  python sync.py run       -> run in a loop every poll_interval_seconds
  python sync.py reset     -> clear the syncToken to force a full reload on the next cycle

Direction: Google -> Microsoft only. Events created, edited, or deleted in Google
are mirrored to the Microsoft calendar. Changes made directly in Microsoft do NOT
flow back to Google - by design, to avoid loops and duplicates.
"""

import json
import logging
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

from google_client import GoogleClient
from mapping import google_to_ms, is_cancelled
from ms_client import MicrosoftClient
from ms_client_com import OutlookComClient
from store import Store

# Support both PyInstaller bundle (sync.exe) and plain script execution.
if getattr(sys, "frozen", False):
    ROOT_DIR = Path(sys.executable).parent
else:
    ROOT_DIR = Path(__file__).resolve().parent.parent  # go up from src/ to project root
BASE_DIR = ROOT_DIR
CONFIG_PATH = ROOT_DIR / "config.json"
DB_PATH = ROOT_DIR / "sync_state.db"
LOG_PATH = ROOT_DIR / "sync.log"

# Always log to the file. Also echo to stdout when it is a real stream: in a
# windowed (no-console) build sys.stdout may be None or redirected to the null
# device, in which case the file handler alone carries the log.
_log_handlers = [logging.FileHandler(LOG_PATH, encoding="utf-8")]
if sys.stdout is not None:
    _log_handlers.append(logging.StreamHandler(sys.stdout))
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=_log_handlers,
)
log = logging.getLogger("sync")


def load_config():
    if not CONFIG_PATH.exists():
        log.error("config.json not found. Copy config.example.json to config.json and fill it in.")
        sys.exit(1)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def resolve(path_str):
    """Resolve relative paths against the program directory."""
    p = Path(path_str).expanduser()
    return str(p if p.is_absolute() else BASE_DIR / p)


def build_clients(cfg):
    g = GoogleClient(
        credentials_file=resolve(cfg["google"]["credentials_file"]),
        token_file=resolve(cfg["google"]["token_file"]),
        calendar_id=cfg["google"].get("calendar_id", "primary"),
    )
    ms_mode = cfg.get("microsoft", {}).get("mode", "graph")
    if ms_mode == "com":
        m = OutlookComClient(
            outlook_account=cfg.get("microsoft", {}).get("outlook_account", ""),
            tag_prefix=cfg.get("sync", {}).get("tag_prefix", ""),
        )
    else:
        m = MicrosoftClient(
            client_id=cfg["microsoft"]["client_id"],
            tenant_id=cfg["microsoft"]["tenant_id"],
            token_cache_file=resolve(cfg["microsoft"]["token_cache_file"]),
            calendar_id=cfg["microsoft"].get("calendar_id", ""),
        )
    return g, m


def do_login(cfg):
    g, m = build_clients(cfg)
    log.info("Authenticating with Google...")
    g.authenticate()
    log.info("Google OK.")
    log.info("Authenticating with Microsoft...")
    m.authenticate()
    log.info("Microsoft OK. Login complete.")


def sync_once(cfg, g, m, store):
    sync_cfg = cfg["sync"]
    default_tz = sync_cfg.get("default_timezone", "UTC")
    tag_prefix = sync_cfg.get("tag_prefix", "")
    window_days = int(sync_cfg.get("window_days_ahead", 60))

    now = datetime.now(timezone.utc)
    time_min = now.isoformat()
    time_max = (now + timedelta(days=window_days)).isoformat()

    sync_token = store.get_state("google_sync_token")
    events, new_token = g.list_changes(sync_token, time_min, time_max)

    if events == "EXPIRED":
        log.warning("syncToken expired. Performing full reload.")
        store.set_state("google_sync_token", None)
        events, new_token = g.list_changes(None, time_min, time_max)

    created = updated = deleted = skipped = 0

    for ev in events:
        gid = ev.get("id")
        if not gid:
            continue

        if is_cancelled(ev):
            ms_id = store.get_ms_id(gid)
            if ms_id:
                try:
                    m.delete_event(ms_id)
                    store.delete_mapping(gid)
                    deleted += 1
                except Exception as e:
                    log.error("Failed to delete %s: %s", gid, e)
            continue

        payload = google_to_ms(ev, default_tz, tag_prefix)
        ms_id, stored_updated = store.get_mapping(gid)

        # Recover from a missing or reset local DB: if there is no mapping but an
        # Outlook event already carries this Google id's [sync-id] marker, adopt
        # that event instead of creating a duplicate.
        if not ms_id:
            existing = m.find_by_sync_id(gid)
            if existing:
                ms_id = existing
                stored_updated = None  # force one refresh; mapping is stored after

        try:
            if ms_id:
                ev_updated = ev.get("updated", "")
                if ev_updated and ev_updated == stored_updated:
                    continue  # unchanged since last sync, skip
                new_ms_id = m.update_event(ms_id, payload)
                store.put_mapping(gid, new_ms_id, ev_updated)
                updated += 1
            else:
                new_ms_id = m.create_event(payload)
                store.put_mapping(gid, new_ms_id, ev.get("updated", ""))
                created += 1
        except Exception as e:
            log.error("Failed to sync %s: %s", gid, e)
            skipped += 1

    if new_token:
        store.set_state("google_sync_token", new_token)

    log.info("Cycle complete: %d created, %d updated, %d deleted, %d errors.",
             created, updated, deleted, skipped)


def main():
    mode = sys.argv[1] if len(sys.argv) > 1 else "run"
    cfg = load_config()

    if mode == "login":
        do_login(cfg)
        return

    if mode == "reset":
        store = Store(str(DB_PATH))
        store.set_state("google_sync_token", None)
        store.close()
        log.info("syncToken cleared. Next cycle will perform a full reload.")
        return

    g, m = build_clients(cfg)
    log.info("Authenticating...")
    # Background cycles (run/once, spawned by the tray) must never open a browser:
    # with no usable token this raises and the cycle aborts with a logged error,
    # instead of popping an OAuth tab on every interval. Interactive login is the
    # job of the `login` mode above (triggered by the panel's Reconfigure).
    g.authenticate(allow_interactive=False)
    m.authenticate()
    store = Store(str(DB_PATH))

    if mode == "once":
        sync_once(cfg, g, m, store)
        store.close()
        return

    if mode == "run":
        interval = int(cfg["sync"].get("poll_interval_seconds", 300))
        log.info("Starting loop. Interval: %d s. Ctrl+C to stop.", interval)
        try:
            while True:
                try:
                    sync_once(cfg, g, m, store)
                except Exception as e:
                    log.error("Cycle error: %s", e)
                time.sleep(interval)
        except KeyboardInterrupt:
            log.info("Stopped by user.")
        finally:
            store.close()
        return

    log.error("Unknown mode: %s. Use login | once | run | reset.", mode)
    sys.exit(1)


if __name__ == "__main__":
    main()
