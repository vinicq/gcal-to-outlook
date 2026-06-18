"""Single-instance guard for the long-running GUI process (tray / panel).

Windows-only. Uses a named mutex: the kernel releases it automatically when the
process dies, so there is no stale lock to clean up after a crash (the Outlook
COM layer this app drives crashes often enough that a PID-based lockfile would
leave orphans).

IMPORTANT: this guard must protect ONLY the long-lived GUI modes. The tray
spawns the same executable in `once`/`login` modes on every cycle
(monitor.SYNC_CMD); gating those would silently kill syncing. Callers therefore
acquire the mutex only in app.py's `tray` and default branches, after the mode
dispatch.
"""

import ctypes

# A handle valid in both the success and "already exists" cases; the signal is
# in the last error, not the return value.
_ERROR_ALREADY_EXISTS = 183

# Session-scoped (Local\, not Global\): the app installs per-user under
# %LOCALAPPDATA%, so two different Windows logins are two legitimate instances.
_MUTEX_NAME = "Local\\GCalSync-tray-singleton"

# Kept module-global on purpose: the handle must outlive this function so the OS
# holds the mutex for the whole process lifetime. If it were garbage-collected
# the mutex would be released and the guard would stop working.
_held_handle = None


def is_already_running(last_error: int) -> bool:
    """Pure decision: CreateMutexW returns a valid handle whether or not the
    mutex already existed; ERROR_ALREADY_EXISTS in the last error is the only
    reliable signal that another instance owns it."""
    return last_error == _ERROR_ALREADY_EXISTS


def acquire_single_instance(name: str = _MUTEX_NAME) -> bool:
    """Return True if this is the first instance, False if another already holds
    the mutex. On True, the handle is intentionally retained in a module global
    so the mutex lives as long as the process. Never raises: on any unexpected
    failure it returns True (fail-open), because blocking the GUI from launching
    is worse than allowing a rare second instance."""
    global _held_handle
    try:
        # use_last_error=True is required for ctypes.get_last_error() to be
        # reliable; without it ctypes' own internal calls can clobber the error
        # between CreateMutexW and the read.
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        handle = kernel32.CreateMutexW(None, False, name)
        last_error = ctypes.get_last_error()
    except Exception:
        return True
    if is_already_running(last_error):
        return False
    _held_handle = handle  # keep alive for the process lifetime
    return True
