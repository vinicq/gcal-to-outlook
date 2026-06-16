#!/usr/bin/env python3
"""Setup wizard for the GCal -> Microsoft Teams calendar synchronizer.

Automatically detects the best Microsoft access mode:
  - COM mode  : uses the installed Outlook desktop client (no admin, no Microsoft portal)
  - Graph mode: uses Microsoft Graph API (requires app registration in Entra)

Works both as a plain Python script and as setup.exe (PyInstaller bundle).
"""

import json
import os
import re
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

# ── Enable ANSI colors in Windows Console ─────────────────────────────────────
try:
    import ctypes
    ctypes.windll.kernel32.SetConsoleMode(
        ctypes.windll.kernel32.GetStdHandle(-11), 7
    )
except Exception:
    pass

# ── Detect execution mode (script vs compiled .exe) ───────────────────────────
_IS_BUNDLE = getattr(sys, "frozen", False)


def _resolve_sync_cmd(is_bundle, executable, venv_py=None, sync_py=None):
    """Command prefix used to run the sync engine as a child process.

    A frozen build is a single executable (GCalSync.exe): re-invoke it with a
    mode argument such as "login" or "once", routed by app.py. There is no
    separate sync.exe. From a source checkout, run the venv Python against
    src/sync.py.
    """
    if is_bundle:
        return [str(executable)]
    return [str(venv_py), str(sync_py)]


if _IS_BUNDLE:
    ROOT = Path(sys.executable).parent
    VENV = None
    VENV_PY = None
    _SYNC_CMD = _resolve_sync_cmd(True, sys.executable)
else:
    ROOT = Path(__file__).resolve().parent.parent
    VENV = ROOT / ".venv"
    VENV_PY = VENV / "Scripts" / "python.exe"
    _SYNC_PY = Path(__file__).resolve().parent / "sync.py"
    _SYNC_CMD = None  # set in step_venv()

# Run the login/sync child without its own console window. The child
# (GCalSync.exe) hides the console in login mode; without this flag it would
# inherit and hide the wizard's own console window.
_CHILD_FLAGS = getattr(subprocess, "CREATE_NO_WINDOW", 0)

GOOGLE_CREDS   = ROOT / "google_credentials.json"
GOOGLE_TOKEN   = ROOT / "google_token.json"
MS_TOKEN       = ROOT / "ms_token_cache.bin"
CONFIG_JSON    = ROOT / "config.json"
CONFIG_EXAMPLE = ROOT / "config.example.json"
REQUIREMENTS   = ROOT / "requirements.txt"

UUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
    re.IGNORECASE,
)

# ── ANSI ──────────────────────────────────────────────────────────────────────
G    = "\033[92m"
Y    = "\033[93m"
R    = "\033[91m"
B    = "\033[94m"
W    = "\033[97m"
DIM  = "\033[2m"
BOLD = "\033[1m"
RST  = "\033[0m"


# ── UI helpers ────────────────────────────────────────────────────────────────

def banner():
    os.system("cls" if os.name == "nt" else "clear")
    print(f"{B}{'=' * 62}{RST}")
    print(f"{BOLD}{W}  Google Calendar -> Microsoft Teams Synchronizer{RST}")
    print(f"{DIM}  Setup wizard{RST}")
    print(f"{B}{'=' * 62}{RST}")
    print()


def section(title: str):
    bar = "-" * max(0, 54 - len(title))
    print(f"\n{B}-- {title} {bar}{RST}\n")


def ok(msg: str):   print(f"  {G}[OK]{RST}    {msg}")
def warn(msg: str): print(f"  {Y}[!]{RST}     {msg}")
def fail(msg: str): print(f"  {R}[ERROR]{RST}  {msg}")
def info(msg: str): print(f"  {DIM}>{RST}       {msg}")


def ask(prompt: str) -> str:
    return input(f"\n  {BOLD}{Y}?{RST} {prompt}: ").strip()


def abort(msg: str):
    fail(msg)
    print()
    input("  Press Enter to exit...")
    sys.exit(1)


def _pause(msg: str = "Press Enter when done..."):
    input(f"\n  {DIM}{msg}{RST}")


def _open(url: str):
    print(f"\n  {B}Opening browser...{RST}")
    webbrowser.open(url)
    time.sleep(1)


def wait_for_file(path: Path, message: str):
    print(f"\n  {Y}{message}{RST}", end="", flush=True)
    while not path.exists():
        print(".", end="", flush=True)
        time.sleep(2)
    print(f" {G}found!{RST}")


def run_live(cmd: list, cwd=None) -> int:
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, cwd=cwd,
    )
    for line in proc.stdout:
        s = line.rstrip()
        if s:
            print(f"    {DIM}{s}{RST}")
    proc.wait()
    return proc.returncode


# ── Outlook COM detection ─────────────────────────────────────────────────────

def _outlook_com_available() -> bool:
    """Test Outlook access via COM using the Python that has pywin32.

    Uses the .venv Python (which already has pywin32 installed) via subprocess
    to avoid depending on the system Python having pywin32.
    """
    if _IS_BUNDLE:
        # pywin32 is bundled - test directly
        try:
            import win32com.client  # noqa: PLC0415
            app = win32com.client.Dispatch("Outlook.Application")
            ns = app.GetNamespace("MAPI")
            _ = ns.GetDefaultFolder(9).Items.Count  # touch to verify Outlook is reachable
            return True
        except Exception:
            return False

    # Use venv Python (already has pywin32), fall back to system Python
    py = str(VENV_PY) if VENV_PY and VENV_PY.exists() else sys.executable
    code = (
        "import win32com.client, pywintypes;"
        "a=win32com.client.Dispatch('Outlook.Application');"
        "n=a.GetNamespace('MAPI');"
        "n.GetDefaultFolder(9).Items.Count;"
        "print('ok')"
    )
    try:
        r = subprocess.run(
            [py, "-c", code],
            capture_output=True, text=True, timeout=20,
        )
        return r.returncode == 0 and "ok" in r.stdout
    except Exception:
        return False


# ── config.json generation ────────────────────────────────────────────────────

def _base_config() -> dict:
    if CONFIG_EXAMPLE.exists():
        return json.loads(CONFIG_EXAMPLE.read_text(encoding="utf-8"))
    return {
        "google": {
            "calendar_id": "primary",
            "credentials_file": "google_credentials.json",
            "token_file": "google_token.json",
        },
        "microsoft": {},
        "sync": {
            "poll_interval_seconds": 300,
            "window_days_ahead": 60,
            "default_timezone": "America/New_York",
            "tag_prefix": "[GCal]",
            "direction": "google_to_ms",
        },
    }


def _write_config_com():
    cfg = _base_config()
    cfg["microsoft"] = {"mode": "com"}
    CONFIG_JSON.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def _write_config_graph(client_id: str, tenant_id: str):
    cfg = _base_config()
    cfg["microsoft"]["mode"] = "graph"
    cfg["microsoft"]["client_id"] = client_id
    cfg["microsoft"]["tenant_id"] = tenant_id
    cfg["microsoft"].setdefault("token_cache_file", "ms_token_cache.bin")
    cfg["microsoft"].setdefault("calendar_id", "")
    CONFIG_JSON.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def _ms_mode() -> str:
    try:
        cfg = json.loads(CONFIG_JSON.read_text(encoding="utf-8"))
        return cfg.get("microsoft", {}).get("mode", "")
    except Exception:
        return ""


# ── Steps ─────────────────────────────────────────────────────────────────────

def step_python():
    section("Step 1 of 5 -- Python")
    if _IS_BUNDLE:
        ok("Executable build (.exe). Python is embedded internally.")
    else:
        v = sys.version.split()[0]
        ok(f"Python {v} detected.")


def step_google_creds():
    section("Step 2 of 5 -- Google Credentials")

    if GOOGLE_CREDS.exists():
        ok("google_credentials.json already found. Skipping.")
        return

    print("  We will create the credentials in the Google Cloud Console.")
    print(f"  {DIM}There are 4 sub-steps. The wizard opens the right page at each one.{RST}")

    # 2a ── create project ─────────────────────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[2a] Create a project in Google Cloud{RST}")
    print()
    print("  The project creation page will open.")
    print(f"  {BOLD}If prompted to log in:{RST} use the Google account you want to sync.")
    print()
    print("  On the page that opens:")
    print(f"    {BOLD}Project name:{RST}  any name, e.g.:  sync-calendar")
    print(f"    {BOLD}Location:{RST}      leave as is")
    print(f"    Click  {BOLD}Create{RST}")
    print()
    print("  Wait for the 'Project created' notification (top right corner).")
    print(f"  Click  {BOLD}Select project{RST}  in that notification.")
    _pause("Press Enter to open the project creation page...")
    _open("https://console.cloud.google.com/projectcreate")
    _pause()

    # 2b ── enable API ─────────────────────────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[2b] Enable the Google Calendar API{RST}")
    print()
    print("  On the page that opens:")
    print("    Confirm the project you just created is selected at the top.")
    print(f"    Click the blue  {BOLD}Enable{RST}  button.")
    print("    Wait for the page to show 'API enabled'.")
    _pause("Press Enter to open the API page...")
    _open("https://console.cloud.google.com/apis/library/calendar-json.googleapis.com")
    _pause()

    # 2c ── OAuth consent ──────────────────────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[2c] Configure the OAuth consent screen{RST}")
    print()
    print("  On the page that opens:")
    print(f"    If you see  {BOLD}Internal{RST}  and  {BOLD}External{RST}: choose  {BOLD}External{RST}  and click  {BOLD}Create{RST}")
    print()
    print("  In the form:")
    print(f"      {BOLD}App name:{RST}            any name, e.g.:  Sync Calendar")
    print(f"      {BOLD}User support email:{RST}  select your email from the dropdown")
    print(f"      {BOLD}Developer contact information{RST}  (section at the bottom):")
    print(f"        {BOLD}Email addresses:{RST}  enter your email")
    print()
    print(f"    Click  {BOLD}Save and continue{RST}")
    print(f"    Next screen (Scopes): click  {BOLD}Save and continue{RST}  without changes")
    print(f"    Next screen (Test users): click  {BOLD}Save and continue{RST}  without changes")
    print(f"    Last screen (Summary): click  {BOLD}Back to dashboard{RST}")
    _pause("Press Enter to open the OAuth consent page...")
    _open("https://console.cloud.google.com/apis/credentials/consent")
    _pause()

    # 2d ── create credential ──────────────────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[2d] Create the credential and download the JSON file{RST}")
    print()
    print("  On the page that opens:")
    print(f"    Click  {BOLD}+ Create credentials{RST}  (top bar)")
    print(f"    Select  {BOLD}OAuth client ID{RST}")
    print()
    print("  In the form:")
    print(f"      {BOLD}Application type:{RST}  select  {BOLD}Desktop app{RST}")
    print(f"      {BOLD}Name:{RST}              any name")
    print(f"    Click  {BOLD}Create{RST}")
    print()
    print(f"    In the dialog that appears, click  {BOLD}Download JSON{RST}")
    print()
    print(f"    {Y}After the download:{RST}")
    print(f"      Rename the file to:  {BOLD}google_credentials.json{RST}")
    print("      Move it to this folder:")
    print(f"      {DIM}{ROOT}{RST}")
    _pause("Press Enter to open the Credentials page...")
    _open("https://console.cloud.google.com/apis/credentials")

    wait_for_file(GOOGLE_CREDS, "Waiting for google_credentials.json in folder")
    ok("google_credentials.json received.")


def step_venv():
    """Step 3 - create the virtual environment and install dependencies (including pywin32)."""
    global _SYNC_CMD

    if _IS_BUNDLE:
        ok("Compiled executable. Python environment is not needed.")
        return

    section("Step 3 of 5 -- Python environment and dependencies")

    if VENV_PY.exists():
        info("Virtual environment (.venv) already exists.")
    else:
        info("Creating isolated virtual environment (.venv)...")
        rc = run_live([sys.executable, "-m", "venv", str(VENV)])
        if rc != 0:
            abort("Failed to create the virtual environment.")
        ok("Virtual environment created.")

    info("Installing dependencies (this may take 1-2 minutes)...")
    run_live([str(VENV_PY), "-m", "pip", "install", "--upgrade", "pip", "-q"])
    rc = run_live([str(VENV_PY), "-m", "pip", "install", "-r", str(REQUIREMENTS)])
    if rc != 0:
        abort("Failed to install dependencies. See the output above.")
    ok("Dependencies installed.")

    _SYNC_CMD = _resolve_sync_cmd(False, sys.executable, VENV_PY, _SYNC_PY)


def step_microsoft_creds():
    """Step 4 - detect Outlook COM or configure Graph API.

    Runs AFTER step_venv() so that pywin32 is available in the .venv.
    """
    section("Step 4 of 5 -- Microsoft Teams access")

    # Already configured?
    mode = _ms_mode()
    if mode == "com":
        ok("Outlook COM mode already configured. Skipping.")
        return
    if mode == "graph":
        try:
            cfg = json.loads(CONFIG_JSON.read_text(encoding="utf-8"))
            if UUID_RE.match(cfg.get("microsoft", {}).get("client_id", "")):
                ok("Graph API mode already configured. Skipping.")
                return
        except Exception:
            pass

    # Try Outlook COM (uses the venv which already has pywin32)
    print("  Checking if Outlook is accessible...")
    print()
    if _outlook_com_available():
        ok("Outlook detected and working!")
        info("Selected mode: Outlook COM.")
        info("The synchronizer creates events directly in Outlook - no admin,")
        info("no Microsoft portal, no Microsoft OAuth required.")
        _write_config_com()
        ok("config.json generated.")
        return

    # Outlook not accessible - fall back to Graph API
    warn("Outlook not detected via COM. Using Microsoft Graph API.")
    warn("This mode requires app registration in the Microsoft portal.")
    print()
    print("  We will register the app in Microsoft Entra.")
    print(f"  {DIM}There are 3 sub-steps. The wizard opens the right page at each one.{RST}")

    # 4a ── register the app ───────────────────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[4a] Register the application in Microsoft Entra{RST}")
    print()
    print("  The Microsoft Entra portal will open.")
    print(f"  {BOLD}If prompted to log in:{RST} use the Microsoft account for Teams.")
    print()
    print("  On the page that opens (app list):")
    print(f"    Click  {BOLD}+ New registration{RST}  (top bar)")
    print()
    print("  In the form:")
    print(f"      {BOLD}Name:{RST}                  any name, e.g.:  sync-calendar")
    print(f"      {BOLD}Supported account types:{RST} leave the first option selected")
    print(f"      {BOLD}Redirect URI:{RST}")
    print(f"        Dropdown:  {BOLD}Public client/native{RST}")
    print(f"        URL:  {BOLD}http://localhost{RST}")
    print()
    print(f"    Click  {BOLD}Register{RST}")
    _pause("Press Enter to open Microsoft Entra...")
    _open("https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade")
    _pause("Press Enter when registration is complete...")

    # 4b ── copy the IDs ───────────────────────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[4b] Copy the two application IDs{RST}")
    print()
    print("  On the app page you just created:")
    print(f"    {BOLD}Application (client) ID{RST}   format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")
    print(f"    {BOLD}Directory (tenant) ID{RST}     same format")
    print()

    while True:
        client_id = ask("Paste the  Application (client) ID  here")
        if UUID_RE.match(client_id):
            ok("Valid format.")
            break
        warn("Invalid format. Must contain hyphens: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")

    while True:
        tenant_id = ask("Paste the  Directory (tenant) ID  here")
        if UUID_RE.match(tenant_id):
            ok("Valid format.")
            break
        warn("Invalid format. Must contain hyphens: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx")

    # 4c ── authentication and permissions ────────────────────────────────────
    print()
    print(f"  {BOLD}{G}[4c] Configure authentication and permissions{RST}")
    print()
    print("  Still on the app page in Entra:")
    print()
    print(f"  {BOLD}Authentication:{RST}")
    print(f"    Left menu > click  {BOLD}Authentication{RST}")
    print(f"    Scroll to  {BOLD}Advanced settings{RST}")
    print(f"    'Allow public client flows' > select  {BOLD}Yes{RST}")
    print(f"    Click  {BOLD}Save{RST}")
    print()
    print(f"  {BOLD}API permissions:{RST}")
    print(f"    Left menu > click  {BOLD}API permissions{RST}")
    print(f"    Click  {BOLD}+ Add a permission{RST}")
    print(f"    Click  {BOLD}Microsoft Graph{RST}")
    print(f"    Click  {BOLD}Delegated permissions{RST}")
    print(f"    Search:  {BOLD}Calendars{RST}  and check  {BOLD}Calendars.ReadWrite{RST}")
    print(f"    Click  {BOLD}Add permissions{RST}")
    _pause()

    _write_config_graph(client_id, tenant_id)
    if not CONFIG_JSON.exists():
        abort("Failed to generate config.json.")
    ok("config.json generated in Graph API mode.")


def step_login():
    section("Step 5 of 5 -- Login")

    mode = _ms_mode()

    if mode == "com":
        if GOOGLE_TOKEN.exists() and MS_TOKEN.exists():
            ok("Login already configured. Skipping.")
            return

        print(f"  The browser will open  {BOLD}once{RST}: Google login.")
        print("  Microsoft does not require a separate login.")
        print(f"  {DIM}The synchronizer uses the Outlook already signed in on your computer.{RST}")
        print()
        _pause("Press Enter to open the browser...")

        rc = subprocess.run(_SYNC_CMD + ["login"], creationflags=_CHILD_FLAGS).returncode
        if rc != 0:
            print()
            warn("There was a problem with the Google login.")
            print()
            input("  Press Enter to exit...")
            sys.exit(1)

        # Sentinel: SYNC.bat checks this file to know setup is complete
        if not MS_TOKEN.exists():
            MS_TOKEN.write_bytes(b"com-mode")

        ok("Google login complete.")

    else:
        if GOOGLE_TOKEN.exists() and MS_TOKEN.exists():
            ok("Tokens already exist. Skipping login.")
            return

        print(f"  The browser will open  {BOLD}twice{RST}:")
        print(f"  {BOLD}1st time:{RST} Google login  -- authorize calendar read access")
        print(f"  {BOLD}2nd time:{RST} Microsoft login -- authorize Teams write access")
        print()
        _pause("Press Enter to open the browser...")

        rc = subprocess.run(_SYNC_CMD + ["login"], creationflags=_CHILD_FLAGS).returncode
        if rc != 0:
            print()
            warn("There was a problem during login.")
            warn("If you see 'access blocked by administrator',")
            warn("your organization's IT team must approve the app in Entra.")
            print()
            input("  Press Enter to exit...")
            sys.exit(1)

        ok("Logins complete. Tokens saved.")


def step_test():
    section("Test -- Running the first sync cycle")
    print("  Please wait...")
    print()
    subprocess.run(_SYNC_CMD + ["once"], creationflags=_CHILD_FLAGS)
    print()
    info("Open Teams > Calendar and look for events with the [GCal] prefix.")
    info("It may take up to 1 minute for events to appear in Teams.")


def step_schedule():
    section("Auto-start")
    print("  Do you want the synchronizer to start automatically when Windows boots?")
    print(f"  {DIM}Recommended: no manual steps needed after login.{RST}")
    print()
    resp = ask("Schedule auto-start? [y/n]").lower()

    if resp not in ("s", "sim", "y", "yes"):
        info("OK. Use option [3] in SYNC.bat whenever you want to schedule it.")
        return

    vbs = ROOT / "run-oculto.vbs"
    result = subprocess.run(
        ["schtasks", "/create",
         "/tn", "GCal-Teams-Sync",
         "/tr", f'wscript.exe "{vbs}"',
         "/sc", "ONLOGON",
         "/ru", os.environ.get("USERNAME", ""),
         "/f"],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        ok("Task registered in Windows Task Scheduler.")
        info("The synchronizer will start automatically on each Windows login.")
        info("To remove: use option [4] in SYNC.bat.")
    else:
        warn("Could not register the task automatically.")
        warn("Open SYNC.bat as Administrator and use option [3].")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    banner()
    if GOOGLE_CREDS.exists():
        print(f"  This wizard configures everything in  {BOLD}2 steps{RST}.")
        print(f"  {DIM}Google credentials already included. Only login is needed.{RST}")
        print(f"  {DIM}Estimated time: 2-3 minutes.{RST}")
    else:
        print(f"  This wizard configures everything in  {BOLD}5 steps{RST}.")
        print(f"  {DIM}You will need to access the Google Cloud Console once.{RST}")
        print(f"  {DIM}Estimated time: 10-15 minutes.{RST}")
    print()
    _pause("Press Enter to begin...")

    try:
        step_python()           # 1 - verify Python
        step_google_creds()     # 2 - Google credentials
        step_venv()             # 3 - virtual environment + pywin32 (BEFORE COM detection)
        step_microsoft_creds()  # 4 - detect Outlook COM or configure Graph API
        step_login()            # 5 - Google login (and Microsoft if Graph mode)
        step_test()             # initial test run
        step_schedule()         # schedule auto-start
    except KeyboardInterrupt:
        print(f"\n\n  {Y}Setup interrupted. Run the wizard again to resume.{RST}\n")
        sys.exit(1)

    print()
    print(f"{G}{'=' * 62}{RST}")
    print(f"{BOLD}{W}  Setup complete!{RST}")
    print(f"{G}{'=' * 62}{RST}")
    print()
    print(f"  Open  {BOLD}SYNC.bat{RST}  to test or manage the synchronizer.")
    print("  If auto-start was enabled, no manual steps are needed.")
    print()
    input("  Press Enter to close...")


if __name__ == "__main__":
    main()
