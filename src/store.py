"""Local SQLite storage.

Stores:
- the mapping between a Google event ID and the corresponding Microsoft event ID,
  along with the Google 'updated' timestamp of the last synced version;
- the Google syncToken, for incremental sync (changed items only).
"""

import sqlite3
from pathlib import Path


class Store:
    def __init__(self, db_path: str):
        self.db_path = str(Path(db_path).expanduser())
        self.conn = sqlite3.connect(self.db_path)
        self.conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self):
        self.conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS event_map (
                google_id   TEXT PRIMARY KEY,
                ms_id       TEXT NOT NULL,
                updated_at  TEXT,
                source      TEXT DEFAULT 'google'
            );
            CREATE INDEX IF NOT EXISTS idx_event_map_ms_id ON event_map(ms_id);
            CREATE TABLE IF NOT EXISTS state (
                key   TEXT PRIMARY KEY,
                value TEXT
            );
            """
        )
        self.conn.commit()

        # inline migration: add 'source' column to existing databases that lack it
        try:
            self.conn.execute(
                "ALTER TABLE event_map ADD COLUMN source TEXT DEFAULT 'google'"
            )
            self.conn.commit()
        except sqlite3.OperationalError:
            # column already exists - idempotent
            pass

        # inline migration: create index on existing databases that lack it
        try:
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_event_map_ms_id ON event_map(ms_id)"
            )
            self.conn.commit()
        except sqlite3.OperationalError:
            pass

    # ---- event mapping ----

    def get_ms_id(self, google_id: str) -> str | None:
        row = self.conn.execute(
            "SELECT ms_id FROM event_map WHERE google_id = ?", (google_id,)
        ).fetchone()
        return row["ms_id"] if row else None

    def get_mapping(self, google_id: str):
        """Returns (ms_id, updated_at) or (None, None) if the record does not exist."""
        row = self.conn.execute(
            "SELECT ms_id, updated_at FROM event_map WHERE google_id = ?", (google_id,)
        ).fetchone()
        if row:
            return row["ms_id"], row["updated_at"]
        return None, None

    def put_mapping(
        self,
        google_id: str,
        ms_id: str,
        updated_at: str = "",
        source: str = "google",
    ):
        self.conn.execute(
            "INSERT INTO event_map (google_id, ms_id, updated_at, source) "
            "VALUES (?, ?, ?, ?) "
            "ON CONFLICT(google_id) DO UPDATE SET "
            "ms_id=excluded.ms_id, "
            "updated_at=excluded.updated_at, "
            "source=excluded.source",
            (google_id, ms_id, updated_at, source),
        )
        self.conn.commit()

    def delete_mapping(self, google_id: str):
        self.conn.execute("DELETE FROM event_map WHERE google_id = ?", (google_id,))
        self.conn.commit()

    # ---- state (syncToken) ----

    def get_state(self, key: str) -> str | None:
        row = self.conn.execute(
            "SELECT value FROM state WHERE key = ?", (key,)
        ).fetchone()
        return row["value"] if row else None

    def set_state(self, key: str, value: str | None):
        if value is None:
            self.conn.execute("DELETE FROM state WHERE key = ?", (key,))
        else:
            self.conn.execute(
                "INSERT INTO state (key, value) VALUES (?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )
        self.conn.commit()

    def close(self):
        self.conn.close()
