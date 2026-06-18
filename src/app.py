"""Unified entry point: GCal -> Outlook/Teams Sync.

Automatic sequence when launched with no arguments:
  1. If not yet configured: opens the setup wizard in the terminal
  2. If already configured: opens the status window and manual sync

Command-line modes (for scripts and autostart):
  GCalSync.exe           -> wizard (if needed) + monitor
  GCalSync.exe tray      -> monitor minimized to tray + background sync.
                            This is what Windows autostart launches.
  GCalSync.exe setup     -> wizard only, does not open the monitor afterward
  GCalSync.exe run       -> headless background sync loop (no window). Manual/CLI
                            use only; autostart uses "tray", not this.
  GCalSync.exe once      -> single sync cycle, then exit
  GCalSync.exe login     -> redo all logins
  GCalSync.exe reset     -> delete syncToken (forces full reload)
  GCalSync.exe dedup     -> scan Outlook for duplicate [GCal] events and remove them
"""

import ctypes
import logging
import os
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    ROOT = Path(sys.executable).parent
else:
    ROOT = Path(__file__).resolve().parent.parent

CONFIG = ROOT / "config.json"
GOOGLE_TOKEN = ROOT / "google_token.json"


def _bind_stdio_to_console():
    """Point sys.stdin/stdout/stderr at the current console (CONIN$/CONOUT$).
    Callers must import the wizard/dedup AFTER this so their import-time console
    setup (e.g. enabling ANSI/VT processing) targets this console."""
    try:
        sys.stdin  = open("CONIN$",  encoding="utf-8")
        sys.stdout = open("CONOUT$", "w", encoding="utf-8")
        sys.stderr = open("CONOUT$", "w", encoding="utf-8")
    except OSError:
        pass


def _alloc_console():
    """The app is built windowed (no console subsystem), so a black terminal
    never appears for the GUI/sync paths. The interactive wizard and dedup do
    need a console for input()/print(). Prefer attaching to the parent's console
    (so output is inline when launched from a terminal or SYNC.bat); otherwise
    allocate a new one. No-op if this process already owns a console."""
    k = ctypes.windll.kernel32
    if k.GetConsoleWindow():
        return
    ATTACH_PARENT_PROCESS = -1
    if not k.AttachConsole(ATTACH_PARENT_PROCESS) and not k.AllocConsole():
        return
    _bind_stdio_to_console()


def _free_console():
    """Detach the console opened by _alloc_console (e.g. after first-run setup)
    so it does not linger behind the monitor window. Rebind stdio to the null
    device afterwards: leaving sys.stdout/stderr pointing at the freed console
    would crash the next write (monitor, update check, Tk/pystray stderr)."""
    try:
        ctypes.windll.kernel32.FreeConsole()
    except Exception:
        pass
    sys.stdin  = open(os.devnull)
    sys.stdout = open(os.devnull, "w")
    sys.stderr = open(os.devnull, "w")


def _setup_done() -> bool:
    return CONFIG.exists() and GOOGLE_TOKEN.exists()


def _relaunch_cmd():
    """Command to start a fresh monitor process after the first-run wizard.
    Frozen: re-exec the bundled exe with no mode (setup is now done, so it opens
    the monitor). Source: run app.py under pythonw (no console) when available."""
    if getattr(sys, "frozen", False):
        return [sys.executable]
    pyw = Path(sys.executable).with_name("pythonw.exe")
    runner = str(pyw) if pyw.exists() else sys.executable
    return [runner, str(Path(__file__).resolve())]


def _child_env():
    """Environment for a relaunched frozen child, with PyInstaller's _MEIPASS2
    stripped. In onefile mode the bootloader sets _MEIPASS2 so a re-exec reuses
    the parent's extracted temp dir; when the child outlives the parent (here it
    must - the parent exits to tear down its console), that dir is deleted and
    the child crashes on startup with "Tcl data directory ... not found".
    Dropping _MEIPASS2 makes the child extract its own copy and own its lifetime.
    """
    env = dict(os.environ)
    env.pop("_MEIPASS2", None)
    return env


def _relaunch_monitor():
    """Hand off to a fresh monitor process and let the caller exit. The wizard
    allocated a console; FreeConsole alone leaves its window orphaned (Enter
    does nothing and it never closes), but a process exit tears the console down
    cleanly. The new process runs with no console."""
    import subprocess
    creation = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        subprocess.Popen(_relaunch_cmd(), creationflags=creation,
                         close_fds=True, env=_child_env())
    except Exception:
        # If the handoff fails, open the monitor in-process so the user is not
        # left with nothing after completing setup.
        _free_console()
        from monitor import Monitor
        Monitor().run()


def _guard_stdio():
    """In a windowed build with no console, sys.stdout/stderr are None, so any
    stray print() in the sync path would crash. Point them at the null device.
    Modes that need real output (wizard/dedup) call _alloc_console afterwards,
    which rebinds stdio to the allocated console and overrides this."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")


def main():
    _guard_stdio()
    mode = sys.argv[1].lower() if len(sys.argv) > 1 else ""

    # Silent modes: no console, no window. Invoked by autostart/SYNC.bat.
    # These are also the modes the tray spawns per cycle (SYNC_CMD = [exe, once]),
    # so they must NOT be gated by the single-instance mutex below.
    if mode in ("run", "once", "login", "reset"):
        import sync as _sync
        # Without a console, an unhandled exception (e.g. the Outlook COM layer
        # raising when Outlook itself crashes) would reach the PyInstaller
        # windowed bootloader and pop a traceback window. Catch it here: log with
        # traceback to sync.log and exit non-zero instead. SystemExit and
        # KeyboardInterrupt propagate untouched (load_config's sys.exit, run's
        # Ctrl+C handling).
        try:
            _sync.main()
        except Exception:
            logging.getLogger("sync").exception("Fatal error in mode '%s'", mode)
            sys.exit(1)
        return

    # Dedup mode: scan and remove duplicate events (needs console output)
    if mode == "dedup":
        _alloc_console()
        import dedup as _dedup
        _dedup.main()
        input("\nPress Enter to close.")
        return

    # Setup mode: wizard only, returns without opening the monitor (used by SYNC.bat)
    if mode == "setup":
        _alloc_console()
        import setup_wizard as _wiz
        _wiz.main()
        return

    # Tray mode: launched by the Windows Startup folder. Opens the monitor
    # straight to the tray (no window popup on login) and lets it drive the
    # sync loop. If setup is somehow incomplete, show the window so the user
    # can finish it instead of starting hidden with nothing to act on.
    if mode == "tray":
        from single_instance import acquire_single_instance
        if not acquire_single_instance():
            return  # another tray/panel instance already runs the sync loop
        from monitor import Monitor
        Monitor(start_hidden=_setup_done()).run()
        return

    # Default mode: first run goes through the setup wizard in a temporary
    # console, then hands off to a fresh monitor process (so the console closes);
    # an already-configured launch opens the monitor directly, gated by the
    # single-instance mutex.
    if not _setup_done():
        _alloc_console()
        import setup_wizard as _wiz
        _wiz.main()
        # Exit this process so its allocated console is torn down; the relaunched
        # process opens the monitor and acquires the single-instance mutex. The
        # mutex is intentionally not held here so the relaunch can take it.
        _relaunch_monitor()
        return

    from single_instance import acquire_single_instance
    if not acquire_single_instance():
        return
    from monitor import Monitor
    Monitor().run()


if __name__ == "__main__":
    main()
