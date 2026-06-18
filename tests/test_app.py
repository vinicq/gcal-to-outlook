"""Tests for app.py's mode routing: the two guarantees that keep the fixes safe.

1. The single-instance mutex is acquired ONLY for the long-lived GUI modes
   (tray, default). It must NEVER be acquired for the modes the tray spawns per
   cycle (once/login) or the other silent/console modes (run/reset/dedup/setup),
   because the tray launches the same executable in `once` mode every interval
   (monitor.SYNC_CMD); gating those would silently kill syncing.

2. In a silent mode, an unhandled Exception (e.g. the Outlook COM layer raising
   when Outlook crashes) is logged and turned into exit code 1, instead of
   reaching the PyInstaller windowed bootloader and popping a traceback window.
   SystemExit/KeyboardInterrupt must still propagate untouched.

No GUI is constructed: monitor.Monitor is replaced by a fake (constructing the
real one needs a display, per test_monitor.py). sync.main / dedup.main /
setup_wizard.main are faked so no Google/Outlook/console is touched.

Oracle is app.py's documented routing, never a re-run of production code.
"""

import sys

import pytest

import app


class _Spy:
    """Records how many times it was called and returns a fixed value."""

    def __init__(self, ret=True):
        self.calls = 0
        self.ret = ret

    def __call__(self, *a, **k):
        self.calls += 1
        return self.ret


class _FakeMonitor:
    def __init__(self, **kwargs):
        self.kwargs = kwargs
        self.ran = False

    def run(self):
        self.ran = True


# ── Guarantee 1: the mutex never gates the spawned/silent modes ──────────────
@pytest.mark.parametrize("mode", ["once", "run", "login", "reset"])
def test_sync_modes_never_acquire_mutex(mode, monkeypatch):
    spy = _Spy()
    monkeypatch.setattr("single_instance.acquire_single_instance", spy)
    import sync
    monkeypatch.setattr(sync, "main", lambda: None)
    monkeypatch.setattr(sys, "argv", ["app", mode])

    app.main()

    # The armadilha: if this is ever > 0, the tray's per-cycle `once` subprocess
    # would detect the running tray's mutex and exit without syncing.
    assert spy.calls == 0


def test_dedup_mode_never_acquires_mutex(monkeypatch):
    spy = _Spy()
    monkeypatch.setattr("single_instance.acquire_single_instance", spy)
    monkeypatch.setattr(app, "_alloc_console", lambda: None)
    import dedup
    monkeypatch.setattr(dedup, "main", lambda: None)
    monkeypatch.setattr("builtins.input", lambda *a, **k: "")
    monkeypatch.setattr(sys, "argv", ["app", "dedup"])

    app.main()

    assert spy.calls == 0


def test_setup_mode_never_acquires_mutex(monkeypatch):
    spy = _Spy()
    monkeypatch.setattr("single_instance.acquire_single_instance", spy)
    monkeypatch.setattr(app, "_alloc_console", lambda: None)
    import setup_wizard
    monkeypatch.setattr(setup_wizard, "main", lambda: None)
    monkeypatch.setattr(sys, "argv", ["app", "setup"])

    app.main()

    assert spy.calls == 0


# ── Guarantee 1 (cont.): GUI modes DO gate ───────────────────────────────────
def test_tray_mode_acquires_mutex_and_runs_monitor(monkeypatch):
    spy = _Spy(ret=True)
    monkeypatch.setattr("single_instance.acquire_single_instance", spy)
    import monitor
    fake = _FakeMonitor()
    monkeypatch.setattr(monitor, "Monitor", lambda **k: fake)
    monkeypatch.setattr(app, "_setup_done", lambda: True)
    monkeypatch.setattr(sys, "argv", ["app", "tray"])

    app.main()

    assert spy.calls == 1
    assert fake.ran is True


def test_tray_second_instance_does_not_open_monitor(monkeypatch):
    monkeypatch.setattr(
        "single_instance.acquire_single_instance", lambda *a, **k: False)
    import monitor
    constructed = {"n": 0}

    def _fake_monitor(**k):
        constructed["n"] += 1
        return _FakeMonitor()

    monkeypatch.setattr(monitor, "Monitor", _fake_monitor)
    monkeypatch.setattr(app, "_setup_done", lambda: True)
    monkeypatch.setattr(sys, "argv", ["app", "tray"])

    app.main()

    # Gate said "already running" -> no second window/loop.
    assert constructed["n"] == 0


def test_default_mode_configured_second_instance_does_not_open_monitor(monkeypatch):
    # Already configured + gate says "already running": no second window/loop.
    monkeypatch.setattr(
        "single_instance.acquire_single_instance", lambda *a, **k: False)
    monkeypatch.setattr(app, "_setup_done", lambda: True)
    import monitor
    fake = _FakeMonitor()
    monkeypatch.setattr(monitor, "Monitor", lambda **k: fake)
    monkeypatch.setattr(sys, "argv", ["app"])

    app.main()

    assert fake.ran is False


# ── First-run handoff: wizard then relaunch (so the console closes) ───────────
def test_default_first_run_runs_wizard_then_hands_off_not_monitor(monkeypatch):
    wizard_ran = {"n": 0}
    relaunched = {"n": 0}
    monkeypatch.setattr(app, "_setup_done", lambda: False)
    monkeypatch.setattr(app, "_alloc_console", lambda: None)
    monkeypatch.setattr(app, "_relaunch_monitor",
                        lambda: relaunched.__setitem__("n", relaunched["n"] + 1))
    import setup_wizard
    monkeypatch.setattr(
        setup_wizard, "main", lambda: wizard_ran.__setitem__("n", 1))
    import monitor
    monitor_built = {"n": 0}

    def _fake_monitor(**k):
        monitor_built["n"] += 1
        return _FakeMonitor()

    monkeypatch.setattr(monitor, "Monitor", _fake_monitor)
    monkeypatch.setattr(sys, "argv", ["app"])

    app.main()

    # The wizard runs, then control hands off to a fresh process; the monitor
    # must NOT be opened in this (console-owning) process.
    assert wizard_ran["n"] == 1
    assert relaunched["n"] == 1
    assert monitor_built["n"] == 0


def test_relaunch_cmd_frozen_reexecs_the_exe(monkeypatch):
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", r"C:\Apps\GCalSync\GCalSync.exe")
    # Frozen: re-exec the bundled exe with no mode argument (setup is done now).
    assert app._relaunch_cmd() == [r"C:\Apps\GCalSync\GCalSync.exe"]


def test_relaunch_cmd_source_runs_app_py(monkeypatch):
    monkeypatch.setattr(sys, "frozen", False, raising=False)
    cmd = app._relaunch_cmd()
    # Source: a python runner plus the app.py path, nothing else.
    assert len(cmd) == 2
    assert cmd[1].endswith("app.py")


def test_child_env_strips_meipass2(monkeypatch):
    # A relaunched onefile child must NOT inherit _MEIPASS2: it would reuse the
    # parent's extracted temp dir, which is deleted when the parent exits,
    # crashing the child with "Tcl data directory ... not found".
    monkeypatch.setenv("_MEIPASS2", r"C:\Users\x\AppData\Local\Temp\_MEI31282")
    monkeypatch.setenv("PATH", "keep-me")
    env = app._child_env()
    assert "_MEIPASS2" not in env
    assert env["PATH"] == "keep-me"


def test_child_env_absent_meipass2_is_noop(monkeypatch):
    monkeypatch.delenv("_MEIPASS2", raising=False)
    env = app._child_env()
    assert "_MEIPASS2" not in env


# ── Guarantee 2: silent-mode exception -> log + exit 1, no traceback window ───
def test_silent_mode_exception_exits_with_code_1(monkeypatch):
    import sync

    def _boom():
        raise RuntimeError("Outlook gone")

    monkeypatch.setattr(sync, "main", _boom)
    monkeypatch.setattr(sys, "argv", ["app", "once"])

    with pytest.raises(SystemExit) as exc:
        app.main()
    assert exc.value.code == 1


def test_silent_mode_keyboardinterrupt_is_not_swallowed(monkeypatch):
    import sync

    def _interrupt():
        raise KeyboardInterrupt()

    # except Exception must let BaseException (Ctrl+C) propagate, not turn it
    # into a logged exit-1.
    monkeypatch.setattr(sync, "main", _interrupt)
    monkeypatch.setattr(sys, "argv", ["app", "run"])

    with pytest.raises(KeyboardInterrupt):
        app.main()
