"""Behavior tests for sync.sync_once orchestration.

sync_once IS the unit under test, so it is NEVER mocked (J3). Only the API
edges are faked: a fake Google client (scripted list_changes) and a fake
Microsoft client (records create/update/delete and returns ids). The Store is
the REAL store.Store backed by a tmp_path SQLite file.

Oracle for every expected value: the documented sync rules in sync.py's module
docstring and sync_once body — new events are created, changed events updated,
unchanged events skipped, cancelled events with a mapping deleted, the new
syncToken persisted, and an EXPIRED token forces a full reload with sync_token
None. None of these expectations re-implement production logic; they restate
the contract.
"""

import pytest

from store import Store
from sync import sync_once

# --------------------------------------------------------------------------
# Fakes for the API edges (NOT the unit under test)
# --------------------------------------------------------------------------

class FakeGoogleClient:
    """Scripts list_changes. Each scripted entry is (events, token).

    Records every (sync_token) it was called with so tests can assert the
    EXPIRED -> full-reload retry passed sync_token=None on the second call.
    """

    def __init__(self, scripted):
        self._scripted = list(scripted)
        self.calls = []  # list of sync_token values passed in

    def list_changes(self, sync_token, time_min, time_max):
        self.calls.append(sync_token)
        events, token = self._scripted.pop(0)
        return events, token


class FakeMicrosoftClient:
    """Records create/update/delete calls and returns ids.

    raise_on_create makes create_event raise, to exercise the error path.
    """

    def __init__(self, raise_on_create=False, existing=None):
        self.created = []   # list of payloads passed to create_event
        self.updated = []   # list of (entry_id, payload)
        self.deleted = []   # list of entry_ids
        # existing: maps google_id -> entry_id for events already in Outlook,
        # simulating the [sync-id] marker lookup when the local DB has no mapping.
        self.existing = dict(existing or {})
        self.find_calls = []  # google_ids passed to find_by_sync_id
        self._raise_on_create = raise_on_create
        self._counter = 0

    def create_event(self, payload):
        if self._raise_on_create:
            raise RuntimeError("MS create failed")
        self._counter += 1
        ms_id = f"ms-{self._counter}"
        self.created.append(payload)
        return ms_id

    def update_event(self, entry_id, payload):
        self.updated.append((entry_id, payload))
        return entry_id

    def delete_event(self, entry_id):
        self.deleted.append(entry_id)

    def find_by_sync_id(self, google_id):
        self.find_calls.append(google_id)
        return self.existing.get(google_id)


# --------------------------------------------------------------------------
# Fixtures / helpers
# --------------------------------------------------------------------------

@pytest.fixture
def cfg():
    return {
        "sync": {
            "default_timezone": "UTC",
            "tag_prefix": "[GCAL]",
            "window_days_ahead": 30,
        }
    }


@pytest.fixture
def store(tmp_path):
    s = Store(str(tmp_path / "sync_state.db"))
    try:
        yield s
    finally:
        s.close()


def _timed_event(gid, updated):
    return {
        "id": gid,
        "summary": "Standup",
        "status": "confirmed",
        "updated": updated,
        "start": {"dateTime": "2026-07-01T09:00:00-03:00", "timeZone": "America/Sao_Paulo"},
        "end": {"dateTime": "2026-07-01T09:15:00-03:00", "timeZone": "America/Sao_Paulo"},
    }


# --------------------------------------------------------------------------
# Tests
# --------------------------------------------------------------------------

def test_new_event_is_created_and_mapping_stored(cfg, store):
    event = _timed_event("g-new", "2026-06-16T10:00:00Z")
    g = FakeGoogleClient([([event], "tok-1")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    # exactly one create, with the subject the mapping layer produced
    assert len(m.created) == 1
    assert m.created[0]["subject"] == "[GCAL] Standup"
    assert m.updated == []
    assert m.deleted == []

    # mapping persisted with the returned ms id and the event's updated stamp
    ms_id, stored_updated = store.get_mapping("g-new")
    assert ms_id == "ms-1"
    assert stored_updated == "2026-06-16T10:00:00Z"


def test_no_mapping_but_existing_event_is_adopted_not_duplicated(cfg, store):
    # Regression: after a lost/reset DB, yesterday's events are still in Outlook
    # (carrying the [sync-id] marker). The sync must adopt them, never duplicate.
    # Local store is empty, but Outlook already has an event for this Google id.
    event = _timed_event("g-1", "2026-06-16T10:00:00Z")
    g = FakeGoogleClient([([event], "tok-1")])
    m = FakeMicrosoftClient(existing={"g-1": "ms-yesterday"})

    sync_once(cfg, g, m, store)

    # NO duplicate created; the pre-existing event is updated in place
    assert m.created == []
    assert len(m.updated) == 1
    assert m.updated[0][0] == "ms-yesterday"
    # and the mapping is now persisted so later cycles skip without scanning
    ms_id, stored_updated = store.get_mapping("g-1")
    assert ms_id == "ms-yesterday"
    assert stored_updated == "2026-06-16T10:00:00Z"


def test_no_mapping_and_no_existing_event_creates_once(cfg, store):
    # Counterpart: a genuinely new event (not present in Outlook) is created.
    event = _timed_event("g-new", "2026-06-16T10:00:00Z")
    g = FakeGoogleClient([([event], "tok-1")])
    m = FakeMicrosoftClient(existing={})  # nothing pre-existing

    sync_once(cfg, g, m, store)

    assert m.find_calls == ["g-new"]  # it looked before creating
    assert len(m.created) == 1
    assert m.updated == []


def test_changed_event_triggers_update_not_create(cfg, store):
    store.put_mapping("g-1", "ms-existing", "2026-06-16T10:00:00Z")
    # same id, but a newer 'updated' than what is stored
    event = _timed_event("g-1", "2026-06-16T12:00:00Z")
    g = FakeGoogleClient([([event], "tok-1")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    assert m.created == []
    assert len(m.updated) == 1
    entry_id, payload = m.updated[0]
    assert entry_id == "ms-existing"
    assert payload["subject"] == "[GCAL] Standup"

    # the stored updated stamp advances to the new value
    _, stored_updated = store.get_mapping("g-1")
    assert stored_updated == "2026-06-16T12:00:00Z"


def test_unchanged_event_is_skipped(cfg, store):
    store.put_mapping("g-1", "ms-existing", "2026-06-16T10:00:00Z")
    # identical 'updated' to what is stored -> no work
    event = _timed_event("g-1", "2026-06-16T10:00:00Z")
    g = FakeGoogleClient([([event], "tok-1")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    assert m.created == []
    assert m.updated == []
    assert m.deleted == []


def test_cancelled_event_with_mapping_is_deleted(cfg, store):
    store.put_mapping("g-1", "ms-existing", "2026-06-16T10:00:00Z")
    cancelled = {"id": "g-1", "status": "cancelled"}
    g = FakeGoogleClient([([cancelled], "tok-1")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    # the mapped MS event is deleted and the mapping removed
    assert m.deleted == ["ms-existing"]
    assert store.get_mapping("g-1") == (None, None)
    assert m.created == []
    assert m.updated == []


def test_cancelled_event_without_mapping_is_a_noop(cfg, store):
    cancelled = {"id": "g-unknown", "status": "cancelled"}
    g = FakeGoogleClient([([cancelled], "tok-1")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    # nothing to delete, nothing created/updated
    assert m.deleted == []
    assert m.created == []
    assert m.updated == []


def test_new_sync_token_is_persisted(cfg, store):
    event = _timed_event("g-new", "2026-06-16T10:00:00Z")
    g = FakeGoogleClient([([event], "tok-FRESH")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    assert store.get_state("google_sync_token") == "tok-FRESH"


def test_expired_token_forces_full_reload_with_none(cfg, store):
    # seed an existing token so the first call uses it
    store.set_state("google_sync_token", "tok-STALE")
    event = _timed_event("g-new", "2026-06-16T10:00:00Z")
    # first call returns EXPIRED, second (full reload) returns real data
    g = FakeGoogleClient([("EXPIRED", None), ([event], "tok-2")])
    m = FakeMicrosoftClient()

    sync_once(cfg, g, m, store)

    # called twice: first with the stale token, then again with None
    assert g.calls == ["tok-STALE", None]
    # the reload still processed the event and stored the new token
    assert len(m.created) == 1
    assert store.get_state("google_sync_token") == "tok-2"


def test_ms_create_error_is_counted_and_loop_continues(cfg, store):
    # first event fails on create; second must still be processed
    ev_fail = _timed_event("g-fail", "2026-06-16T10:00:00Z")
    ev_ok = _timed_event("g-ok", "2026-06-16T10:00:00Z")
    g = FakeGoogleClient([([ev_fail, ev_ok], "tok-1")])

    class FailFirstClient(FakeMicrosoftClient):
        def create_event(self, payload):
            # fail only for the first event, succeed afterwards
            if payload["subject"] == "[GCAL] Standup" and not self.created:
                # g-fail and g-ok share the same subject, so disambiguate by
                # the sync-id marker embedded in the body
                if "g-fail" in payload["body"]["content"]:
                    raise RuntimeError("MS create failed")
            return super().create_event(payload)

    m = FailFirstClient()

    # the loop must not crash despite the create error
    sync_once(cfg, g, m, store)

    # the failed event left no mapping; the OK one was created and stored
    assert store.get_mapping("g-fail") == (None, None)
    ok_ms_id, _ = store.get_mapping("g-ok")
    assert ok_ms_id is not None
    assert any("g-ok" in p["body"]["content"] for p in m.created)
    # token still persisted even though one event errored
    assert store.get_state("google_sync_token") == "tok-1"
