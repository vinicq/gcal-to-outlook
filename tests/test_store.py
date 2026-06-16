"""Behavior tests for store.Store against a real on-disk SQLite file.

Each test gets its own tmp_path db, so there is no shared mutable state and no
ordering dependency between tests (J6).
"""

import pytest

from store import Store


@pytest.fixture
def store(tmp_path):
    s = Store(str(tmp_path / "sync_state.db"))
    try:
        yield s
    finally:
        s.close()


def test_put_then_get_mapping_roundtrip(store):
    store.put_mapping("g-1", "ms-abc", "2026-06-16T10:00:00Z")

    ms_id, updated_at = store.get_mapping("g-1")
    assert ms_id == "ms-abc"
    assert updated_at == "2026-06-16T10:00:00Z"


def test_put_mapping_upserts_same_google_id(store):
    store.put_mapping("g-1", "ms-old", "2026-06-16T10:00:00Z")
    store.put_mapping("g-1", "ms-new", "2026-06-16T11:30:00Z")

    # upsert: the row reflects the latest values, not the first ones
    ms_id, updated_at = store.get_mapping("g-1")
    assert ms_id == "ms-new"
    assert updated_at == "2026-06-16T11:30:00Z"

    # and there is exactly one row for that google_id (no duplicate)
    count = store.conn.execute(
        "SELECT COUNT(*) FROM event_map WHERE google_id = ?", ("g-1",)
    ).fetchone()[0]
    assert count == 1


def test_get_ms_id_returns_value_and_none_for_unknown(store):
    store.put_mapping("g-1", "ms-abc", "2026-06-16T10:00:00Z")

    assert store.get_ms_id("g-1") == "ms-abc"
    assert store.get_ms_id("does-not-exist") is None


def test_delete_mapping_removes_the_row(store):
    store.put_mapping("g-1", "ms-abc", "2026-06-16T10:00:00Z")

    store.delete_mapping("g-1")

    # documented contract: get_mapping returns (None, None) when absent
    assert store.get_mapping("g-1") == (None, None)
    assert store.get_ms_id("g-1") is None


def test_set_and_get_state_roundtrip(store):
    store.set_state("google_sync_token", "TOKEN-123")

    assert store.get_state("google_sync_token") == "TOKEN-123"


def test_set_state_none_deletes_the_key(store):
    store.set_state("google_sync_token", "TOKEN-123")

    store.set_state("google_sync_token", None)

    # None means delete, so a subsequent read yields None
    assert store.get_state("google_sync_token") is None


def test_get_state_unknown_key_is_none(store):
    assert store.get_state("never-set") is None
