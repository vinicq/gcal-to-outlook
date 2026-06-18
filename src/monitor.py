"""Control panel: GCal -> Outlook/Teams Sync.

Features:
  - Google and Outlook account status with visual indicators
  - Reconfigure each account individually
  - Adjust sync interval (saved to config.json)
  - Enable/disable Windows autostart
  - Manual sync trigger
  - Log viewer
"""

import json
import os
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import urllib.request
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, simpledialog

import pystray
from PIL import Image, ImageDraw

from _version import GITHUB_REPO, RELEASES_URL, __version__

# ── Paths ─────────────────────────────────────────────────────────────────────
if getattr(sys, "frozen", False):
    ROOT_DIR = Path(sys.executable).parent
else:
    ROOT_DIR = Path(__file__).resolve().parent.parent

CONFIG_FILE  = ROOT_DIR / "config.json"
GOOGLE_TOKEN = ROOT_DIR / "google_token.json"
GOOGLE_CREDS = ROOT_DIR / "google_credentials.json"
LOG_FILE     = ROOT_DIR / "sync.log"

_STARTUP_FOLDER = Path(os.environ.get("APPDATA", "")) / \
    "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
_STARTUP_VBS = _STARTUP_FOLDER / "GCalSync-autorun.vbs"


def _asset(name: str) -> Path:
    """Resolve a bundled asset. Frozen builds extract data to sys._MEIPASS."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = ROOT_DIR
    return base / "assets" / name

# Command for a single sync cycle and for login
if getattr(sys, "frozen", False):
    _EXE = Path(sys.executable)
    SYNC_CMD  = [str(_EXE), "once"]
    LOGIN_CMD = [str(_EXE), "login"]
else:
    _PY  = ROOT_DIR / ".venv" / "Scripts" / "python.exe"
    _SC  = ROOT_DIR / "src" / "sync.py"
    SYNC_CMD  = [str(_PY), str(_SC), "once"]
    LOGIN_CMD = [str(_PY), str(_SC), "login"]

# Hide the console window of child sync/login processes (this is a tray app).
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

# ── Colors (Catppuccin Mocha) ─────────────────────────────────────────────────
BG    = "#1e1e2e"
BG2   = "#313244"
BG3   = "#45475a"
FG    = "#cdd6f4"
FG2   = "#a6adc8"
GREEN = "#a6e3a1"
RED   = "#f38ba8"
BLUE  = "#89b4fa"
YELL  = "#f9e2af"
SURF  = "#6c7086"


# ── Helpers: config ───────────────────────────────────────────────────────────
def _load_cfg() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _save_cfg(cfg: dict):
    CONFIG_FILE.write_text(
        json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8"
    )


# ── Helpers: account status ───────────────────────────────────────────────────
def _google_status() -> tuple[bool, str]:
    if not GOOGLE_TOKEN.exists():
        return False, "Not configured"
    try:
        import google.oauth2.credentials
        import requests as _req
        creds = google.oauth2.credentials.Credentials.from_authorized_user_file(
            str(GOOGLE_TOKEN)
        )
        if creds.expired and creds.refresh_token:
            from google.auth.transport.requests import Request
            creds.refresh(Request())
        r = _req.get(
            "https://www.googleapis.com/userinfo/v2/me",
            headers={"Authorization": f"Bearer {creds.token}"},
            timeout=6,
        )
        if r.status_code == 200:
            return True, r.json().get("email", "Connected")
    except Exception:
        pass
    return GOOGLE_TOKEN.exists(), "Connected" if GOOGLE_TOKEN.exists() else "Not configured"


def _ms_status() -> tuple[bool, str]:
    """Check Outlook connection using only DisplayName and ExchangeStoreType.

    Never accesses SmtpAddress, CurrentUser, or any address-book property -
    all of which trigger the Outlook Object Model Guard security popup.

    Runs on a worker thread, so COM must be initialized on this thread first.
    Without CoInitialize, Dispatch fails with CO_E_NOTINITIALIZED (-2147221008).
    """
    import pythoncom
    pythoncom.CoInitialize()
    try:
        import win32com.client
        app = win32com.client.Dispatch("Outlook.Application")
        ns  = app.GetNamespace("MAPI")
        cfg = _load_cfg()
        target = cfg.get("microsoft", {}).get("outlook_account", "").lower()

        for store in ns.Stores:
            try:
                display = getattr(store, "DisplayName", "")
                if not display:
                    continue
                stype = getattr(store, "ExchangeStoreType", 0)
                dlower = display.lower()
                # Match by configured account name (display name contains the email)
                if target and (target in dlower or dlower in target):
                    return True, display
                # Or accept any primary Exchange mailbox
                if stype == 1:
                    return True, display
            except Exception:
                continue

        # Fallback: Outlook is open - show configured account or generic label
        return True, target or "Connected"
    except Exception as e:
        return False, f"Disconnected ({str(e)[:40]})"
    finally:
        pythoncom.CoUninitialize()


# ── Helpers: last sync ────────────────────────────────────────────────────────
def _last_sync() -> tuple[str, str]:
    if not LOG_FILE.exists():
        return "Never", ""
    try:
        lines = LOG_FILE.read_text(encoding="utf-8", errors="ignore").splitlines()
        for line in reversed(lines):
            if "Cycle complete" in line:
                parts = line.split(" ")
                ts = f"{parts[0]} {parts[1].split(',')[0]}" if len(parts) >= 2 else "?"
                idx = line.find("Cycle complete")
                return ts, line[idx:]
        return "Never", ""
    except Exception:
        return "?", ""


def _recent_log(n=80) -> str:
    if not LOG_FILE.exists():
        return "(no log)"
    try:
        lines = LOG_FILE.read_text(encoding="utf-8", errors="ignore").splitlines()
        return "\n".join(lines[-n:])
    except Exception:
        return ""


# ── Helpers: autostart ────────────────────────────────────────────────────────
def _autostart_enabled() -> bool:
    # Autostart is the launcher dropped in the Windows Startup folder.
    # A legacy scheduled task may still exist (disabled) from older versions;
    # it does not count as enabled and is cleaned up by _autostart_set(False).
    return _STARTUP_VBS.exists()


def _remove_scheduled_task():
    """Delete the legacy ONLOGON task an older version (or the setup wizard) may
    have created. Silent if absent. The Startup-folder VBS is the single source
    of autostart now; this keeps the two mechanisms from both firing at login."""
    try:
        subprocess.run(
            ["schtasks", "/delete", "/tn", "GCal-Teams-Sync", "/f"],
            capture_output=True, timeout=5,
        )
    except Exception:
        pass


def _autostart_set(enable: bool):
    if enable:
        # Creates VBS in the Windows Startup folder (no admin required).
        # Launches the tray monitor (mode "tray"): it shows the tray icon AND
        # runs the sync loop itself, so one launch covers both. The old "run"
        # mode synced but was invisible, which read as "it did not start".
        # Also drop any legacy scheduled task so only one launcher fires.
        _remove_scheduled_task()
        if getattr(sys, "frozen", False):
            exe = str(Path(sys.executable))
            cmd_line = f'Chr(34) & "{exe}" & Chr(34) & " tray"'
        else:
            pyw = ROOT_DIR / ".venv" / "Scripts" / "pythonw.exe"
            app = ROOT_DIR / "src" / "app.py"
            cmd_line = (
                f'Chr(34) & "{pyw}" & Chr(34) & " " & '
                f'Chr(34) & "{app}" & Chr(34) & " tray"'
            )
        content = (
            'Dim shell\n'
            'Set shell = CreateObject("WScript.Shell")\n'
            f'shell.Run {cmd_line}, 0, False\n'
            'Set shell = Nothing\n'
        )
        _STARTUP_FOLDER.mkdir(parents=True, exist_ok=True)
        _STARTUP_VBS.write_text(content, encoding="utf-8")
    else:
        _STARTUP_VBS.unlink(missing_ok=True)
        _remove_scheduled_task()


# ── Helpers: update check ───────────────────────────────────────────────────
def _parse_version(tag: str) -> tuple:
    """'v1.2.3' / '1.2.3' -> (1, 2, 3). Non-numeric parts are dropped, so a
    malformed tag compares as lower than any real release rather than crashing."""
    nums = []
    for part in tag.lstrip("vV").split("."):
        digits = "".join(c for c in part if c.isdigit())
        if not digits:
            break
        nums.append(int(digits))
    return tuple(nums)


def _latest_release() -> str | None:
    """Return the latest release tag (e.g. 'v1.0.5') from GitHub, or None on any
    failure. Network/parse errors are swallowed: a missed check must never break
    the app or block the UI. Uses the public API, no token, no telemetry."""
    url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
    req = urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"GCalSync/{__version__}",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        tag = data.get("tag_name")
        return tag or None
    except Exception:
        return None


def _update_available() -> str | None:
    """Return the newer release tag if GitHub has one above the running version,
    else None (up to date, or the check could not run)."""
    tag = _latest_release()
    if not tag:
        return None
    if _parse_version(tag) > _parse_version(__version__):
        return tag
    return None


# ── Tray icon ─────────────────────────────────────────────────────────────────
def _make_tray_image(size=64) -> Image.Image:
    """Programmatic icon: blue circle with sync arrow."""
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    m = size // 2
    d.ellipse([3, 3, size - 3, size - 3], fill="#89b4fa")
    # Arrow body
    aw = size // 5
    ah = size // 10
    d.rectangle([m - aw, m - ah, m + ah, m + ah], fill="#1e1e2e")
    # Arrowhead
    d.polygon([
        (m + ah, m - aw // 2 - ah),
        (m + aw, m),
        (m + ah, m + aw // 2 + ah),
    ], fill="#1e1e2e")
    return img


# ── Widget helpers ────────────────────────────────────────────────────────────
def _label(parent, text, font=None, fg=FG, bg=BG, **kw):
    return tk.Label(parent, text=text,
                    font=font or ("Segoe UI", 9), fg=fg, bg=bg, **kw)


def _btn(parent, text, cmd, color=BG3, fg=FG, width=None):
    b = tk.Button(
        parent, text=text, command=cmd,
        font=("Segoe UI", 8), bg=color, fg=fg,
        activebackground=BG2, activeforeground=FG,
        relief="flat", padx=8, pady=4, cursor="hand2",
    )
    if width:
        b.configure(width=width)
    return b


# ── Main monitor ──────────────────────────────────────────────────────────────
class Monitor:
    def __init__(self, start_hidden: bool = False):
        self._sync_running = False
        self._log_visible  = False
        self._tray: pystray.Icon | None = None

        self.root = tk.Tk()
        self.root.title("GCal  Teams Sync")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)
        self._set_window_icon()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self._refresh_all()
        self.root.after(30_000, self._periodic_refresh)

        # Keep the Startup launcher pointing at the current command. Older
        # installs wrote a VBS that launched the headless "run" loop (no window,
        # no tray); refreshing it here migrates them to the visible tray on the
        # next launch without the user having to re-toggle the checkbox.
        if _autostart_enabled():
            try:
                _autostart_set(True)
            except Exception:
                pass

        # When launched by the Startup folder we open straight to the tray so
        # login is not interrupted by a window. This process also drives the
        # sync loop (spawning one sync cycle per interval), so a single launch
        # gives both the visible icon and background syncing.
        if start_hidden:
            self.root.after(0, self._hide_to_tray)
        self._schedule_background_sync(initial=True)
        self._check_update_async()

    def _set_window_icon(self):
        """Set the title-bar and taskbar icon from the bundled icon."""
        try:
            ico = _asset("icon.ico")
            if ico.exists():
                self.root.iconbitmap(default=str(ico))
                return
        except Exception:
            pass
        try:
            png = _asset("icon.png")
            if png.exists():
                self._win_icon = tk.PhotoImage(file=str(png))
                self.root.iconphoto(True, self._win_icon)
        except Exception:
            pass

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        P = 14

        # Header
        hdr = tk.Frame(self.root, bg=BG2, pady=12)
        hdr.pack(fill="x")
        _label(hdr, "GCal  →  Teams Sync",
               font=("Segoe UI", 13, "bold"), fg=FG, bg=BG2).pack()
        _label(hdr, f"v{__version__}",
               font=("Segoe UI", 8), fg=SURF, bg=BG2).pack()
        # Hidden until the update check finds a newer release. Clicking opens
        # the releases page in the browser.
        self.lbl_update = tk.Label(
            hdr, text="", font=("Segoe UI", 8, "underline"),
            fg=YELL, bg=BG2, cursor="hand2")
        self.lbl_update.bind("<Button-1>", lambda _e: webbrowser.open(RELEASES_URL))

        # ── Accounts ─────────────────────────────────────────────────────────
        frm_acc = tk.LabelFrame(self.root, text="  Accounts  ",
            font=("Segoe UI", 9), bg=BG, fg=FG2, bd=1, relief="groove",
            padx=P, pady=8)
        frm_acc.pack(fill="x", padx=P, pady=(P, 0))

        # Google
        rg = tk.Frame(frm_acc, bg=BG)
        rg.pack(fill="x", pady=3)
        self.dot_g = _label(rg, "●", font=("Segoe UI", 11), fg=SURF)
        self.dot_g.pack(side="left")
        _label(rg, "  Google Calendar", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.lbl_google = _label(rg, "  ...", fg=FG2)
        self.lbl_google.pack(side="left", padx=(4, 0))
        _btn(rg, "Reconfigure", self._reconfig_google).pack(side="right")

        # Microsoft
        rm = tk.Frame(frm_acc, bg=BG)
        rm.pack(fill="x", pady=3)
        self.dot_m = _label(rm, "●", font=("Segoe UI", 11), fg=SURF)
        self.dot_m.pack(side="left")
        _label(rm, "  Outlook / Teams", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.lbl_ms = _label(rm, "  ...", fg=FG2)
        self.lbl_ms.pack(side="left", padx=(4, 0))
        _btn(rm, "Reconfigure", self._reconfig_ms).pack(side="right")

        # Outlook requirement note
        _label(frm_acc,
               "Outlook for Windows desktop must be installed and signed in.",
               fg=SURF, font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 0))

        # ── Sync ──────────────────────────────────────────────────────────────
        frm_sync = tk.LabelFrame(self.root, text="  Sync  ",
            font=("Segoe UI", 9), bg=BG, fg=FG2, bd=1, relief="groove",
            padx=P, pady=8)
        frm_sync.pack(fill="x", padx=P, pady=(8, 0))

        self.lbl_last   = _label(frm_sync, "Last sync: ...", fg=FG2, anchor="w")
        self.lbl_last.pack(fill="x")
        self.lbl_result = _label(frm_sync, "", fg=GREEN, anchor="w")
        self.lbl_result.pack(fill="x")

        # Interval
        ri = tk.Frame(frm_sync, bg=BG)
        ri.pack(fill="x", pady=(6, 0))
        _label(ri, "Interval:", fg=FG2).pack(side="left")
        self._interval_var = tk.StringVar(value="5")
        sb = tk.Spinbox(ri, from_=1, to=120, width=4,
                        textvariable=self._interval_var,
                        font=("Segoe UI", 9), bg=BG2, fg=FG,
                        buttonbackground=BG3, relief="flat",
                        disabledbackground=BG2)
        sb.pack(side="left", padx=(6, 4))
        _label(ri, "minutes", fg=FG2).pack(side="left")
        _btn(ri, "Save", self._save_interval, color=BG3).pack(side="left", padx=(10, 0))
        self.lbl_interval_msg = _label(ri, "", fg=YELL)
        self.lbl_interval_msg.pack(side="left", padx=(6, 0))

        # ── Autostart ─────────────────────────────────────────────────────────
        frm_auto = tk.LabelFrame(self.root, text="  Autostart  ",
            font=("Segoe UI", 9), bg=BG, fg=FG2, bd=1, relief="groove",
            padx=P, pady=8)
        frm_auto.pack(fill="x", padx=P, pady=(8, 0))

        ra = tk.Frame(frm_auto, bg=BG)
        ra.pack(fill="x")
        self._auto_var = tk.BooleanVar(value=False)
        self.chk_auto = tk.Checkbutton(
            ra, text="  Start automatically when Windows starts",
            variable=self._auto_var, command=self._toggle_autostart,
            font=("Segoe UI", 9), bg=BG, fg=FG, selectcolor=BG2,
            activebackground=BG, activeforeground=FG,
            anchor="w", cursor="hand2", bd=0, highlightthickness=0,
        )
        self.chk_auto.pack(side="left")
        self.lbl_auto = _label(ra, "...", fg=SURF, font=("Segoe UI", 8, "bold"))
        self.lbl_auto.pack(side="right")

        # ── Main buttons ──────────────────────────────────────────────────────
        frm_btns = tk.Frame(self.root, bg=BG, pady=P)
        frm_btns.pack(padx=P)

        self.btn_sync = tk.Button(
            frm_btns, text="  Sync now  ",
            font=("Segoe UI", 9, "bold"),
            bg=BLUE, fg=BG, activebackground="#74c7ec", activeforeground=BG,
            relief="flat", padx=14, pady=7, cursor="hand2",
            command=self._sync_now,
        )
        self.btn_sync.pack(side="left", padx=(0, 8))

        self.btn_log = tk.Button(
            frm_btns, text="  View log  ",
            font=("Segoe UI", 9),
            bg=BG3, fg=FG, activebackground=BG2,
            relief="flat", padx=14, pady=7, cursor="hand2",
            command=self._toggle_log,
        )
        self.btn_log.pack(side="left")

        self.btn_export = tk.Button(
            frm_btns, text="  Export log  ",
            font=("Segoe UI", 9),
            bg=BG3, fg=FG, activebackground=BG2,
            relief="flat", padx=14, pady=7, cursor="hand2",
            command=self._export_log,
        )
        self.btn_export.pack(side="left", padx=(8, 0))

        # ── Log ───────────────────────────────────────────────────────────────
        self.frm_log = tk.Frame(self.root, bg=BG)
        self.txt_log = scrolledtext.ScrolledText(
            self.frm_log, font=("Consolas", 8),
            bg="#181825", fg=FG2, height=14, width=74,
            relief="flat", state="disabled",
        )
        self.txt_log.pack(padx=P, pady=(0, P))

        self.root.update_idletasks()

    # ── Data refresh ──────────────────────────────────────────────────────────
    def _refresh_all(self):
        threading.Thread(target=self._load_google, daemon=True).start()
        threading.Thread(target=self._load_ms, daemon=True).start()
        self._refresh_sync_labels()
        self._refresh_autostart()
        self._load_interval()

    def _periodic_refresh(self):
        self._refresh_sync_labels()
        self._refresh_autostart()
        self.root.after(30_000, self._periodic_refresh)

    # ── Background sync loop ──────────────────────────────────────────────────
    def _hide_to_tray(self):
        """Start minimized: withdraw the window and show only the tray icon."""
        self.root.withdraw()
        if self._tray is None:
            self._start_tray()

    def _schedule_background_sync(self, initial: bool = False):
        """Drive a sync cycle on the configured interval from the tray process.

        Replaces the separate headless "run" loop: the visible tray process now
        owns the cadence, so a single launch gives both the icon and background
        syncing. Uses Tk's after() (single-threaded scheduler); each cycle is a
        short-lived `once` subprocess spawned by _sync_now (process isolation
        keeps COM/pythoncom out of the UI thread). The child logs to sync.log
        regardless of stdout capture, so per-cycle errors stay diagnosable.
        """
        cfg = _load_cfg()
        secs = int(cfg.get("sync", {}).get("poll_interval_seconds", 300))
        # Floor at 60s; cap at 24h so the ms value stays within Tk after()'s
        # signed 32-bit limit (~24.8 days) even with a bad config value.
        secs = min(max(60, secs), 86_400)
        # On startup, kick off a first sync soon after the UI/auth settle.
        delay = 5_000 if initial else secs * 1000
        self.root.after(delay, self._background_tick)

    def _background_tick(self):
        # Reschedule in a finally so a failure in one cycle can never kill the
        # loop silently - that would recreate the "it stopped syncing" symptom.
        try:
            if not self._sync_running:
                self._sync_now()
        finally:
            self._schedule_background_sync()

    # ── Update check ──────────────────────────────────────────────────────────
    def _check_update_async(self):
        """Query GitHub for a newer release in the background and, if found,
        surface a clickable label (and a tray balloon when minimized). Runs once
        per launch; the network call is off the UI thread."""
        def _run():
            tag = _update_available()
            if tag:
                self.root.after(0, lambda: self._show_update(tag))

        threading.Thread(target=_run, daemon=True).start()

    def _show_update(self, tag: str):
        self.lbl_update.configure(text=f"Nova versão {tag} disponível - clique para baixar")
        self.lbl_update.pack(pady=(4, 0))
        if self._tray is not None:
            try:
                self._tray.notify(
                    f"Versão {tag} disponível. Abra o painel para baixar.",
                    "GCal → Teams Sync",
                )
            except Exception:
                pass

    def _load_google(self):
        ok, label = _google_status()
        text = label if ok else "Offline"
        self.root.after(0, lambda: (
            self.dot_g.configure(fg=GREEN if ok else RED),
            self.lbl_google.configure(text=f"  {text}"),
        ))

    def _load_ms(self):
        ok, label = _ms_status()
        # When not reachable, show a plain "Offline" instead of the raw COM error.
        # The detail stays available through the log and the Export log button.
        text = label if ok else "Offline"
        self.root.after(0, lambda: (
            self.dot_m.configure(fg=GREEN if ok else RED),
            self.lbl_ms.configure(text=f"  {text}"),
        ))

    def _refresh_sync_labels(self):
        ts, result = _last_sync()
        self.lbl_last.configure(text=f"Last sync: {ts}")
        self.lbl_result.configure(text=result)

    def _refresh_autostart(self):
        enabled = _autostart_enabled()
        self._auto_var.set(enabled)
        self.lbl_auto.configure(
            text="Enabled" if enabled else "Disabled",
            fg=GREEN if enabled else SURF,
        )

    def _load_interval(self):
        cfg = _load_cfg()
        secs = cfg.get("sync", {}).get("poll_interval_seconds", 300)
        mins = max(1, int(secs) // 60)
        self._interval_var.set(str(mins))

    # ── Actions: accounts ─────────────────────────────────────────────────────
    def _reconfig_google(self):
        if not messagebox.askyesno(
            "Reconfigure Google",
            "This will delete the current token and open the browser for login.\n\nContinue?",
            parent=self.root,
        ):
            return
        if GOOGLE_TOKEN.exists():
            GOOGLE_TOKEN.unlink()
        self.dot_g.configure(fg=SURF)
        self.lbl_google.configure(text="  Waiting for browser login...")

        def _run():
            subprocess.run(LOGIN_CMD, timeout=120, creationflags=_NO_WINDOW)
            self.root.after(0, lambda: threading.Thread(
                target=self._load_google, daemon=True).start())

        threading.Thread(target=_run, daemon=True).start()

    def _reconfig_ms(self):
        cfg = _load_cfg()
        current = cfg.get("microsoft", {}).get("outlook_account", "")
        novo = simpledialog.askstring(
            "Reconfigure Microsoft",
            "Outlook (Exchange) account email to sync with Teams:",
            initialvalue=current,
            parent=self.root,
        )
        if novo is None:
            return
        novo = novo.strip()
        if novo == current:
            return
        cfg.setdefault("microsoft", {})["outlook_account"] = novo
        _save_cfg(cfg)
        self.dot_m.configure(fg=SURF)
        self.lbl_ms.configure(text="  Checking...")
        threading.Thread(target=self._load_ms, daemon=True).start()

    # ── Action: interval ──────────────────────────────────────────────────────
    def _save_interval(self):
        try:
            mins = int(self._interval_var.get())
            if mins < 1 or mins > 120:
                raise ValueError
        except ValueError:
            self.lbl_interval_msg.configure(text="1-120 min", fg=RED)
            self.root.after(2000, lambda: self.lbl_interval_msg.configure(text=""))
            return
        cfg = _load_cfg()
        cfg.setdefault("sync", {})["poll_interval_seconds"] = mins * 60
        _save_cfg(cfg)
        self.lbl_interval_msg.configure(text="Saved! Restart sync to apply.", fg=YELL)
        self.root.after(3000, lambda: self.lbl_interval_msg.configure(text=""))

    # ── Action: autostart ─────────────────────────────────────────────────────
    def _toggle_autostart(self):
        # The Checkbutton toggles _auto_var before invoking this, so the
        # variable already holds the desired state.
        want = self._auto_var.get()
        try:
            _autostart_set(want)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self.root)
            self._auto_var.set(not want)  # revert on failure
        self._refresh_autostart()

    # ── Action: manual sync ───────────────────────────────────────────────────
    def _sync_now(self):
        if self._sync_running:
            return
        self._sync_running = True
        self.btn_sync.configure(
            text="  Syncing...  ", state="disabled", bg=BG3, fg=FG2)

        def _run():
            rc = -1
            try:
                proc = subprocess.run(SYNC_CMD, capture_output=True, timeout=120,
                                      creationflags=_NO_WINDOW)
                rc = proc.returncode
            except Exception:
                rc = -1
            self.root.after(0, lambda: self._sync_done(rc))

        threading.Thread(target=_run, daemon=True).start()

    def _sync_done(self, rc: int = 0):
        self._sync_running = False
        self.btn_sync.configure(
            text="  Sync now  ", state="normal", bg=BLUE, fg=BG)
        self._refresh_sync_labels()
        # A non-zero exit means the cycle aborted (e.g. Outlook crashed under the
        # COM layer). The detail is already in sync.log; surface it discreetly
        # here instead of a popup, and let the next successful cycle clear it.
        if rc != 0:
            self.lbl_result.configure(
                text="Last sync failed - check the log", fg=RED)
            if self._tray is not None:
                try:
                    self._tray.notify(
                        "A sync cycle failed. Open the panel and view the log.",
                        "GCal → Teams Sync",
                    )
                except Exception:
                    pass
        else:
            # Clear any prior failure styling once a cycle succeeds again.
            self.lbl_result.configure(fg=GREEN)
        if self._log_visible:
            self._update_log()

    # ── Action: log ───────────────────────────────────────────────────────────
    def _toggle_log(self):
        if self._log_visible:
            self.frm_log.pack_forget()
            self.btn_log.configure(text="  View log  ")
            self._log_visible = False
        else:
            self.frm_log.pack(fill="x")
            self.btn_log.configure(text="  Close log  ")
            self._log_visible = True
            self._update_log()

    def _update_log(self):
        text = _recent_log(80)
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.insert("end", text)
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def _export_log(self):
        if not LOG_FILE.exists():
            messagebox.showinfo(
                "Export log", "There is no log file yet.", parent=self.root)
            return
        dest = filedialog.asksaveasfilename(
            parent=self.root, title="Export log",
            defaultextension=".log", initialfile="gcalsync-sync.log",
            filetypes=[("Log file", "*.log"), ("All files", "*.*")],
        )
        if not dest:
            return
        try:
            shutil.copyfile(LOG_FILE, dest)
            messagebox.showinfo(
                "Export log", f"Log saved to:\n{dest}", parent=self.root)
        except Exception as e:
            messagebox.showerror("Export log", str(e), parent=self.root)

    # ── System tray ───────────────────────────────────────────────────────────
    def _on_close(self):
        """Closing the window minimizes to tray instead of exiting."""
        self.root.withdraw()
        if self._tray is None:
            self._start_tray()

    def _start_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Open", self._tray_show, default=True),
            pystray.MenuItem("Sync now", self._tray_sync),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._tray_quit),
        )
        self._tray = pystray.Icon(
            "GCalSync", self._tray_image(), "GCal → Teams Sync", menu
        )
        threading.Thread(target=self._tray.run, daemon=True).start()

    def _tray_image(self) -> Image.Image:
        """Use the bundled icon for the tray; fall back to the drawn icon."""
        try:
            png = _asset("icon.png")
            if png.exists():
                return Image.open(png)
        except Exception:
            pass
        return _make_tray_image()

    def _tray_show(self):
        self.root.after(0, self.root.deiconify)
        if self._tray:
            self._tray.stop()
            self._tray = None

    def _tray_sync(self):
        self.root.after(0, self._sync_now)

    def _tray_quit(self):
        if self._tray:
            self._tray.stop()
            self._tray = None
        self.root.after(0, self.root.destroy)

    # ─────────────────────────────────────────────────────────────────────────
    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    Monitor().run()
