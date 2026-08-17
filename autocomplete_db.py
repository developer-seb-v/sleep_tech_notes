"""
SQLite-backed autocomplete suggestion store.

Suggestions for a handful of free-text fields (mask used, ordering MD,
recording tech, medical history, medication list) are learned from
past tech notes rather than hard-coded: every value a user types and
saves gets recorded here, and the most-used values for a field are
offered back as suggestions the next time that field is edited.

The database lives next to this script as autocomplete.db.
"""

import os
import sqlite3
from datetime import datetime, timezone

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "autocomplete.db")


def _connect():
    return sqlite3.connect(DB_PATH)


def init_db(seed_defaults=None):
    """Create the suggestions table if it doesn't exist yet, and seed
    starter values for any field that has no data at all. A field that
    already has rows (i.e. real notes have been saved) is left
    untouched, so this is safe to call on every startup."""
    with _connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS suggestions (
                field_key TEXT NOT NULL,
                value TEXT NOT NULL,
                use_count INTEGER NOT NULL DEFAULT 0,
                last_used TEXT,
                PRIMARY KEY (field_key, value)
            )
            """
        )

        if not seed_defaults:
            return

        for field_key, values in seed_defaults.items():
            already_seeded = conn.execute(
                "SELECT 1 FROM suggestions WHERE field_key = ? LIMIT 1",
                (field_key,),
            ).fetchone()
            if already_seeded:
                continue
            conn.executemany(
                "INSERT OR IGNORE INTO suggestions (field_key, value, use_count, last_used) "
                "VALUES (?, ?, 0, NULL)",
                [(field_key, value) for value in values],
            )


def get_suggestions(field_key):
    """Return suggestion strings for a field, most-used first."""
    with _connect() as conn:
        rows = conn.execute(
            "SELECT value FROM suggestions WHERE field_key = ? "
            "ORDER BY use_count DESC, value ASC",
            (field_key,),
        ).fetchall()
    return [row[0] for row in rows]


def record_value(field_key, value):
    """Record that `value` was used for `field_key`, bumping its usage
    count so it ranks higher in future suggestions and is remembered
    even if it wasn't one of the seeded starter values."""
    value = value.strip()
    if not value:
        return
    now = datetime.now(timezone.utc).isoformat()
    with _connect() as conn:
        conn.execute(
            """
            INSERT INTO suggestions (field_key, value, use_count, last_used)
            VALUES (?, ?, 1, ?)
            ON CONFLICT (field_key, value) DO UPDATE SET
                use_count = use_count + 1,
                last_used = excluded.last_used
            """,
            (field_key, value, now),
        )


def record_values(field_key, values):
    """Record multiple values for `field_key` (see record_value)."""
    for value in values:
        record_value(field_key, value)
