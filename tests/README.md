# Tests

Unit tests for the pure, deterministic core of the Google -> Outlook sync.

Run them with the project venv:

```
.venv/Scripts/python.exe -m pytest -q
.venv/Scripts/python.exe -m ruff check tests
```

`pyproject.toml` sets `pythonpath = ["src"]`, so tests import the flat modules
directly (`import mapping`, `import store`, `from sync import sync_once`).

## What is covered

| File | Unit under test | Edges faked |
|---|---|---|
| `test_mapping.py` | `mapping.google_to_ms`, `mapping.is_cancelled` | none (pure functions) |
| `test_store.py` | `store.Store` against a real tmp SQLite file | none (real DB on `tmp_path`) |
| `test_ms_client_com.py` | `ms_client_com._to_local_str` (pure) | none |
| `test_sync.py` | `sync.sync_once` (real orchestration + real `Store`) | only the Google and Microsoft API clients |

In `test_sync.py`, `sync_once` is the unit, so it is never replaced. Only the
Google client (`list_changes`) and the Microsoft client (`create_event` /
`update_event` / `delete_event`) are faked, because those are the network/COM
edges. The store is the real `store.Store`.

## What is intentionally NOT unit-tested, and why

These modules are excluded because they require a live GUI, interactive console
input, or a running Outlook/network session. Unit tests there would either be
false greens (mocking the whole surface) or could not run headless.

- **`monitor.py`**: tkinter / system-tray GUI. No logic worth isolating from
  the event loop; needs a display.
- **`setup_wizard.py`**: interactive console wizard driven by `input()`
  prompts. Exercised manually during first-run setup. Its one pure helper,
  `_resolve_sync_cmd` (which child command the wizard runs), IS unit-tested in
  `test_setup_wizard.py` after a frozen-build-only bug shipped there.
- **`dedup.py`** and the COM methods of `ms_client_com.OutlookComClient`
  (`create_event`, `update_event`, `delete_event`, `_apply`, `_connect`,
  `_find_calendar`, `authenticate`) drive a live Outlook desktop session via
  `win32com`. The one pure helper in that module, `_to_local_str`, IS tested.
- **`google_client.py`** / **`ms_client.py`**: live Google Calendar and
  Microsoft Graph network clients (OAuth, HTTP). Faked at the edge in
  `test_sync.py` rather than hit for real.

The orchestration that ties these edges together (`sync.sync_once`) is fully
tested with the real `Store` and faked clients, so the decision logic (create /
update / skip / delete / token persistence / expired-token reload / error
handling) is covered without touching any live service.

---

Tests audited with the falsegreen framework (https://github.com/vinicq/falsegreen).
