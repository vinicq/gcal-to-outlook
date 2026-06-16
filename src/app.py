"""Unified entry point: GCal -> Outlook/Teams Sync.

Automatic sequence when launched with no arguments:
  1. If not yet configured: opens the setup wizard in the terminal
  2. If already configured: opens the status window and manual sync

Command-line modes (for scripts and Task Scheduler):
  GCalSync.exe           -> wizard (if needed) + monitor
  GCalSync.exe setup     -> wizard only, does not open the monitor afterward
  GCalSync.exe run       -> background sync loop (no window)
  GCalSync.exe once      -> single sync cycle, then exit
  GCalSync.exe login     -> redo all logins
  GCalSync.exe reset     -> delete syncToken (forces full reload)
  GCalSync.exe dedup     -> scan Outlook for duplicate [GCal] events and remove them
"""

import ctypes
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).resolve().parent.parent

CONFIG = ROOT / "config.json"
GOOGLE_TOKEN = ROOT / "google_token.json"


def _set_console(visible: bool):
    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if hwnd:
        ctypes.windll.user32.ShowWindow(hwnd, 1 if visible else 0)


def _setup_done() -> bool:
    return CONFIG.exists() and GOOGLE_TOKEN.exists()


def main():
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else ""

    # Silent modes: no UI, invoked by Task Scheduler or SYNC.bat
    if mode in ("run", "once", "login", "reset"):
        _set_console(False)
        import sync as _sync
        _sync.main()
        return

    # Dedup mode: scan and remove duplicate events (shows console output)
    if mode == "dedup":
        _set_console(True)
        import dedup as _dedup
        _dedup.main()
        input("\nPress Enter to close.")
        return

    # Setup mode: wizard only, returns without opening the monitor (used by SYNC.bat)
    if mode == "setup":
        _set_console(True)
        import setup_wizard as _wiz
        _wiz.main()
        return

    # Default mode: wizard if needed, then monitor
    if not _setup_done():
        _set_console(True)
        import setup_wizard as _wiz
        _wiz.main()

    _set_console(False)
    from monitor import Monitor
    Monitor().run()


if __name__ == "__main__":
    main()
