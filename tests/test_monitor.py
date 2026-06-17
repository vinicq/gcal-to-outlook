"""Tests for monitor.py's pure update-check helpers and the autostart command.

What is covered here (no GUI, no network, no real Startup folder):
  - _parse_version: tag -> tuple-of-ints ordering and malformed-input handling.
  - _update_available: the newer-than-running comparison, with the network edge
    (_latest_release) faked.
  - _latest_release: tag extraction from a GitHub JSON payload, with urlopen
    faked; failure paths return None.
  - _autostart_set(True): the command string written into the Startup-folder VBS
    (the "run" -> "tray" migration), with the Startup paths redirected to tmp.

What is deliberately NOT tested (would be a false green to assert against a fake):
  - The tkinter Monitor window, tray icon, and clickable update label: they need
    a real display and event loop. Importing monitor is safe (Tk is only built
    inside Monitor.__init__), but constructing Monitor is not headless-safe.
  - The live GitHub call inside _latest_release against the real API: networked,
    flaky, and not the app's own logic. We fake the HTTP edge instead.
  - schtasks invocation in _autostart_set(False): shells out to a Windows tool.

Oracle for every value assertion is the docstring / spec of the function, never
a re-run of the production code (J2).
"""

import urllib.error
import urllib.request

import monitor


# ── _parse_version ───────────────────────────────────────────────────────────
def test_parse_version_strips_v_prefix_and_splits():
    # Spec: 'v1.2.3' -> (1, 2, 3). Value comes from the docstring, not the code.
    assert monitor._parse_version("v1.2.3") == (1, 2, 3)


def test_parse_version_accepts_bare_numeric_tag():
    assert monitor._parse_version("1.0.4") == (1, 0, 4)


def test_parse_version_uppercase_v_is_also_stripped():
    # lstrip("vV"): an uppercase prefix must parse the same as lowercase.
    assert monitor._parse_version("V2.0.0") == (2, 0, 0)


def test_parse_version_orders_like_a_real_release_bump():
    # The whole point of the tuple is comparison. 1.0.10 > 1.0.9 (numeric, not
    # lexical) and a minor bump beats a patch bump.
    assert monitor._parse_version("v1.0.10") > monitor._parse_version("v1.0.9")
    assert monitor._parse_version("v1.1.0") > monitor._parse_version("v1.0.99")


def test_parse_version_garbage_is_lower_than_any_real_release():
    # Spec: a malformed tag must compare as lower than any real release rather
    # than crash. Empty tuple < any populated tuple.
    garbage = monitor._parse_version("not-a-version")
    assert garbage == ()
    assert garbage < monitor._parse_version("v0.0.1")


def test_parse_version_stops_at_first_non_numeric_segment():
    # Spec: parsing breaks at the first part with no digits. '1.2.x.4' keeps
    # only the leading numeric run, dropping everything from 'x' on.
    assert monitor._parse_version("1.2.x.4") == (1, 2)


def test_parse_version_concatenates_digits_within_a_segment():
    # Spec: "digits within a part are kept" — and they are concatenated, not
    # re-split. So '3rc1' -> '31'. A patch tag with a suffix therefore compares
    # ABOVE the clean patch, which is the documented "malformed sorts low only
    # when it has no digits" trade-off, surfaced here so it is not a surprise.
    assert monitor._parse_version("v1.2.3rc1") == (1, 2, 31)


# ── _update_available (network edge faked via _latest_release) ───────────────
def test_update_available_returns_tag_when_remote_is_newer(monkeypatch):
    # Fake only the edge (J3): _latest_release is the network boundary.
    monkeypatch.setattr(monitor, "__version__", "1.0.4")
    monkeypatch.setattr(monitor, "_latest_release", lambda: "v1.0.5")
    assert monitor._update_available() == "v1.0.5"


def test_update_available_is_none_when_remote_equals_running(monkeypatch):
    monkeypatch.setattr(monitor, "__version__", "1.0.4")
    monkeypatch.setattr(monitor, "_latest_release", lambda: "v1.0.4")
    assert monitor._update_available() is None


def test_update_available_is_none_when_remote_is_older(monkeypatch):
    # Guards against an inverted comparison: an older remote must NOT prompt an
    # "update available". A == result-of-production test would not catch this.
    monkeypatch.setattr(monitor, "__version__", "1.0.4")
    monkeypatch.setattr(monitor, "_latest_release", lambda: "v1.0.3")
    assert monitor._update_available() is None


def test_update_available_is_none_when_check_fails(monkeypatch):
    # _latest_release returns None on any network/parse failure; the caller must
    # treat that as "no update", never raise.
    monkeypatch.setattr(monitor, "__version__", "1.0.4")
    monkeypatch.setattr(monitor, "_latest_release", lambda: None)
    assert monitor._update_available() is None


def test_update_available_handles_v_prefix_mismatch(monkeypatch):
    # Running version has no 'v', remote tag does. Both go through _parse_version
    # so the prefix must not make a same-version remote look newer.
    monkeypatch.setattr(monitor, "__version__", "1.0.4")
    monkeypatch.setattr(monitor, "_latest_release", lambda: "v1.0.4")
    assert monitor._update_available() is None


# ── _latest_release (HTTP edge faked via urllib) ─────────────────────────────
class _FakeResponse:
    def __init__(self, body: bytes):
        self._body = body

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def test_latest_release_extracts_tag_name(monkeypatch):
    # Fake the HTTP edge; assert the function pulls tag_name out of the payload.
    body = b'{"tag_name": "v9.9.9", "name": "ignored"}'
    monkeypatch.setattr(
        urllib.request, "urlopen", lambda *a, **k: _FakeResponse(body)
    )
    assert monitor._latest_release() == "v9.9.9"


def test_latest_release_returns_none_when_tag_absent(monkeypatch):
    # A payload with no tag_name (or empty) must yield None, not "" or KeyError.
    body = b'{"name": "no tag here"}'
    monkeypatch.setattr(
        urllib.request, "urlopen", lambda *a, **k: _FakeResponse(body)
    )
    assert monitor._latest_release() is None


def test_latest_release_swallows_network_error(monkeypatch):
    # Spec: network/parse errors are swallowed and return None so a failed check
    # never breaks the app. Raise from the faked edge and assert None, no bubble.
    def _boom(*a, **k):
        raise urllib.error.URLError("no network")

    monkeypatch.setattr(urllib.request, "urlopen", _boom)
    assert monitor._latest_release() is None


def test_latest_release_swallows_bad_json(monkeypatch):
    monkeypatch.setattr(
        urllib.request,
        "urlopen",
        lambda *a, **k: _FakeResponse(b"<html>not json</html>"),
    )
    assert monitor._latest_release() is None


# ── _autostart_set(True): the command written to the Startup VBS ─────────────
def _written_vbs(monkeypatch, tmp_path):
    """Redirect the Startup-folder constants to tmp and return the VBS text that
    _autostart_set(True) writes. Asserting on the on-disk content is asserting
    behaviour: that file is exactly what Windows executes at login (J5)."""
    folder = tmp_path / "Startup"
    vbs = folder / "GCalSync-autorun.vbs"
    monkeypatch.setattr(monitor, "_STARTUP_FOLDER", folder)
    monkeypatch.setattr(monitor, "_STARTUP_VBS", vbs)
    monitor._autostart_set(True)
    return vbs.read_text(encoding="utf-8")


def test_autostart_frozen_launches_tray_mode(monkeypatch, tmp_path):
    # Frozen build: the VBS must invoke the exe with the "tray" argument.
    monkeypatch.setattr(monitor.sys, "frozen", True, raising=False)
    monkeypatch.setattr(
        monitor.sys, "executable", r"C:\Apps\GCalSync\GCalSync.exe"
    )
    content = _written_vbs(monkeypatch, tmp_path)
    assert "GCalSync.exe" in content
    # The mode is the literal trailing token of the command shell.Run executes:
    # ... & Chr(34) & " tray", 0, False  -> the exe is invoked with "tray".
    assert '& Chr(34) & " tray", 0, False' in content


def test_autostart_frozen_does_not_use_run_mode(monkeypatch, tmp_path):
    # Direct guard against the regression this change fixes: the launcher must
    # NOT start the headless "run" loop (synced but invisible). Substring check
    # is unambiguous because the mode is the literal trailing token.
    monkeypatch.setattr(monitor.sys, "frozen", True, raising=False)
    monkeypatch.setattr(
        monitor.sys, "executable", r"C:\Apps\GCalSync\GCalSync.exe"
    )
    content = _written_vbs(monkeypatch, tmp_path)
    assert '" run"' not in content


def test_autostart_source_build_runs_app_py_in_tray_mode(monkeypatch, tmp_path):
    # Source checkout: pythonw runs src\app.py with "tray" (not sync.py / "run").
    monkeypatch.setattr(monitor.sys, "frozen", False, raising=False)
    content = _written_vbs(monkeypatch, tmp_path)
    assert "app.py" in content
    assert "tray" in content
    assert "sync.py" not in content
    assert '" run"' not in content


def test_autostart_writes_a_runnable_wscript_shell_invocation(monkeypatch, tmp_path):
    # The VBS must be a real shell.Run script, not just a path. Asserts the
    # structural contract WScript needs, so a malformed script is caught.
    monkeypatch.setattr(monitor.sys, "frozen", True, raising=False)
    monkeypatch.setattr(monitor.sys, "executable", r"C:\Apps\GCalSync\GCalSync.exe")
    content = _written_vbs(monkeypatch, tmp_path)
    assert 'CreateObject("WScript.Shell")' in content
    assert "shell.Run " in content


def test_autostart_disable_removes_the_vbs(monkeypatch, tmp_path):
    # _autostart_set(False) must delete the launcher; assert via the public
    # _autostart_enabled() predicate (behaviour, not the unlink call).
    folder = tmp_path / "Startup"
    folder.mkdir()
    vbs = folder / "GCalSync-autorun.vbs"
    vbs.write_text("placeholder", encoding="utf-8")
    monkeypatch.setattr(monitor, "_STARTUP_FOLDER", folder)
    monkeypatch.setattr(monitor, "_STARTUP_VBS", vbs)
    # Avoid shelling out to schtasks during the test.
    monkeypatch.setattr(monitor.subprocess, "run", lambda *a, **k: None)

    assert monitor._autostart_enabled() is True
    monitor._autostart_set(False)
    assert monitor._autostart_enabled() is False


# ── _sync_done: reaction to the cycle subprocess exit code ───────────────────
# These run the real _sync_done logic. We bypass Tk via Monitor.__new__ (so no
# display is needed, unlike constructing Monitor) and inject widget/tray doubles
# on the bare instance. The spec is: a non-zero exit code marks the result label
# red and fires a tray balloon; a zero exit clears the styling back to green and
# never notifies.
class _FakeWidget:
    """Records the last configure() kwargs so assertions read final state."""

    def __init__(self):
        self.config = {}

    def configure(self, **kw):
        self.config.update(kw)


class _FakeTray:
    def __init__(self):
        self.notes = []

    def notify(self, message, title):
        self.notes.append((message, title))


def _bare_monitor_for_sync_done(tray):
    m = monitor.Monitor.__new__(monitor.Monitor)  # no __init__, no Tk
    m._sync_running = True
    m.btn_sync = _FakeWidget()
    m.lbl_result = _FakeWidget()
    m._tray = tray
    m._log_visible = False
    # Collaborators _sync_done calls but that are out of scope here.
    m._refresh_sync_labels = lambda: None
    m._update_log = lambda: None
    return m


def test_sync_done_failure_marks_result_red_and_notifies():
    tray = _FakeTray()
    m = _bare_monitor_for_sync_done(tray)

    m._sync_done(1)  # subprocess exited non-zero -> failed cycle

    assert m.lbl_result.config.get("fg") == monitor.RED
    assert "failed" in m.lbl_result.config.get("text", "").lower()
    assert len(tray.notes) == 1
    assert m._sync_running is False


def test_sync_done_success_clears_to_green_and_does_not_notify():
    tray = _FakeTray()
    m = _bare_monitor_for_sync_done(tray)

    m._sync_done(0)  # clean cycle

    assert m.lbl_result.config.get("fg") == monitor.GREEN
    assert tray.notes == []


def test_sync_done_timeout_code_is_treated_as_failure_without_tray():
    # rc == -1 is the timeout/exception path from _sync_now; with the window
    # open (no tray) it must still mark red and not blow up on a None tray.
    m = _bare_monitor_for_sync_done(None)

    m._sync_done(-1)

    assert m.lbl_result.config.get("fg") == monitor.RED
