"""Tests for single_instance.py (the named-mutex guard).

What is covered (no real mutex, ctypes faked):
  - is_already_running: the pure last-error -> bool decision.
  - acquire_single_instance: True on a fresh mutex, False when the OS reports
    ERROR_ALREADY_EXISTS, and fail-open (True) if the Win32 call blows up.

What is deliberately NOT tested (would need a second real process):
  - That a genuine second OS process actually sees the mutex. That is the
    kernel's contract, not ours; we assert our reading of GetLastError instead.

Oracle for every assertion is the function's spec/docstring (the ctypes contract
for CreateMutexW + ERROR_ALREADY_EXISTS), never a re-run of the production code.
"""

import ctypes

import single_instance


# ── is_already_running (pure) ────────────────────────────────────────────────
def test_is_already_running_true_on_error_already_exists():
    # Spec: ERROR_ALREADY_EXISTS == 183 is the "another instance owns it" signal.
    assert single_instance.is_already_running(183) is True


def test_is_already_running_false_on_success():
    # last error 0 = the mutex was created fresh by this process.
    assert single_instance.is_already_running(0) is False


def test_is_already_running_false_on_unrelated_error():
    # Any other error code is not the duplicate signal.
    assert single_instance.is_already_running(5) is False


# ── acquire_single_instance (ctypes faked) ───────────────────────────────────
class _FakeKernel:
    """Stand-in for the kernel32 WinDLL handle; CreateMutexW returns a fixed,
    valid handle in both the fresh and already-exists cases (mirrors the real
    Win32 behavior: the signal lives in GetLastError, not the return value)."""

    def __init__(self, handle):
        self._handle = handle

    def CreateMutexW(self, sec_attrs, initial_owner, name):
        return self._handle


def test_first_instance_acquires_and_keeps_handle(monkeypatch):
    monkeypatch.setattr(single_instance, "_held_handle", None, raising=False)
    monkeypatch.setattr(ctypes, "WinDLL", lambda *a, **k: _FakeKernel(handle=42))
    monkeypatch.setattr(ctypes, "get_last_error", lambda: 0)

    assert single_instance.acquire_single_instance("Local\\test") is True
    # Contract: the handle must be retained so the OS holds the mutex for the
    # whole process lifetime.
    assert single_instance._held_handle == 42


def test_second_instance_is_rejected(monkeypatch):
    monkeypatch.setattr(single_instance, "_held_handle", None, raising=False)
    monkeypatch.setattr(ctypes, "WinDLL", lambda *a, **k: _FakeKernel(handle=42))
    monkeypatch.setattr(ctypes, "get_last_error", lambda: 183)

    assert single_instance.acquire_single_instance("Local\\test") is False


def test_acquire_fails_open_when_win32_raises(monkeypatch):
    # If the Win32 layer is unavailable/broken, blocking the GUI from launching
    # is worse than allowing a rare second instance: spec says return True.
    def _boom(*a, **k):
        raise OSError("kernel32 unavailable")

    monkeypatch.setattr(ctypes, "WinDLL", _boom)
    assert single_instance.acquire_single_instance("Local\\test") is True
