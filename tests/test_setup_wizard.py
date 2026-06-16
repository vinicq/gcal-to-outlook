"""Regression tests for the wizard's child-process command resolution.

Bug history: in a frozen build the wizard pointed at a separate `sync.exe`,
which the single-exe build never produces. Pressing Enter at the login step
ran a missing program, so the wizard errored and the window closed. The fix
re-invokes the running executable (GCalSync.exe) with a mode argument instead.

These tests exercise the pure resolver directly, so they run on any OS without
freezing the app. The bundle branch is otherwise only reachable from a compiled
build, which is why the original bug shipped untested.
"""

from setup_wizard import _resolve_sync_cmd


def test_bundle_reinvokes_the_running_executable():
    exe = r"C:\Users\me\AppData\Local\GCalSync\GCalSync.exe"
    cmd = _resolve_sync_cmd(True, exe)
    # The single-exe build must call itself; app.py routes the mode argument.
    assert cmd == [exe]


def test_bundle_never_points_at_a_separate_sync_exe():
    # Direct guard against the shipped regression: no standalone sync.exe exists.
    cmd = _resolve_sync_cmd(True, r"C:\app\GCalSync.exe")
    assert len(cmd) == 1
    assert not cmd[0].endswith("sync.exe")


def test_source_checkout_runs_python_against_the_sync_script():
    venv_py = r"C:\proj\.venv\Scripts\python.exe"
    sync_py = r"C:\proj\src\sync.py"
    cmd = _resolve_sync_cmd(False, "unused-executable", venv_py, sync_py)
    assert cmd == [venv_py, sync_py]
