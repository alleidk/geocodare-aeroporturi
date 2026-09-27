"""Bază de date SQLite: utilizatori, joburi, cache geocodare."""

import sqlite3
import time

from . import config

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY,
    username    TEXT UNIQUE NOT NULL COLLATE NOCASE,
    pw_hash     TEXT NOT NULL,
    is_admin    INTEGER NOT NULL DEFAULT 0,
    created_at  REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS jobs (
    id              TEXT PRIMARY KEY,
    user_id         INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    filename        TEXT NOT NULL,
    status          TEXT NOT NULL,           -- new, queued, running, done, failed, cancelled
    created_at      REAL NOT NULL,
    updated_at      REAL NOT NULL,
    sheet           TEXT,
    header_row      INTEGER,
    col_origin      INTEGER,
    col_dest        INTEGER,
    provider        TEXT,
    max_rows        INTEGER,
    groups          TEXT,                    -- JSON: prima coloană a fiecărui grup de 4
    lookup_sheet    TEXT,                    -- foaia cod localitate -> nume
    skip_done       INTEGER DEFAULT 1,       -- sare peste rândurile care au deja rezultat
    phase           TEXT,
    rows_count      INTEGER DEFAULT 0,
    already_done    INTEGER DEFAULT 0,
    unique_count    INTEGER DEFAULT 0,
    cache_hits      INTEGER DEFAULT 0,
    total           INTEGER DEFAULT 0,
    done            INTEGER DEFAULT 0,
    errors          INTEGER DEFAULT 0,
    addresses       INTEGER DEFAULT 0,       -- adrese scrise în rezultat
    found           INTEGER DEFAULT 0,       -- dintre ele, cu coordonate
    message         TEXT,
    output_name     TEXT,
    cancel_requested INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS geocache (
    provider    TEXT NOT NULL,
    key         TEXT NOT NULL,
    result      TEXT NOT NULL,
    created_at  REAL NOT NULL,
    PRIMARY KEY (provider, key)
);
"""

MIGRATIONS = [
    ("skip_done", "INTEGER DEFAULT 1"),
    ("already_done", "INTEGER DEFAULT 0"),
    ("groups", "TEXT"),
    ("lookup_sheet", "TEXT"),
    ("addresses", "INTEGER DEFAULT 0"),
    ("found", "INTEGER DEFAULT 0"),
]


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(config.DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init():
    config.DATA_DIR.mkdir(parents=True, exist_ok=True)
    config.JOBS_DIR.mkdir(parents=True, exist_ok=True)
    with connect() as conn:
        conn.execute("PRAGMA journal_mode = WAL")
        conn.executescript(SCHEMA)
        # Coloane adăugate după prima versiune (bazele de date existente nu le au)
        existing = {row["name"] for row in conn.execute("PRAGMA table_info(jobs)")}
        for name, decl in MIGRATIONS:
            if name not in existing:
                conn.execute(f"ALTER TABLE jobs ADD COLUMN {name} {decl}")


def query(sql: str, params=(), one=False):
    with connect() as conn:
        rows = conn.execute(sql, params).fetchall()
    if one:
        return rows[0] if rows else None
    return rows


def execute(sql: str, params=()):
    with connect() as conn:
        conn.execute(sql, params)


def update_job(job_id: str, **fields):
    fields["updated_at"] = time.time()
    cols = ", ".join(f"{k} = ?" for k in fields)
    execute(f"UPDATE jobs SET {cols} WHERE id = ?", (*fields.values(), job_id))
