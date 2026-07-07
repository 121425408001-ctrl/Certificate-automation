"""
db.py — Postgres-backed per-session storage for the certificate automation app.

Each browser gets a unique session_id (UUID stored in Flask cookie).
All state for that session lives in Postgres under that session_id.

Rolling buffer: before creating a new session, we check the DB size.
If it exceeds MAX_DB_BYTES (default 400 MB = 80% of Neon's 500 MB free tier),
we delete the oldest session (by last_accessed) until there is room.
"""

import os
import json
import psycopg2
import psycopg2.pool
from psycopg2.extras import Json

# ── connection ────────────────────────────────────────────────────────────────

DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://postgres:testpass@localhost:5432/certapp"
)

# 400 MB — gives 100 MB headroom below Neon's 500 MB free limit
MAX_DB_BYTES = int(os.environ.get("MAX_DB_SIZE_BYTES", str(400 * 1024 * 1024)))

_pool = None


def _get_pool():
    global _pool
    if _pool is None:
        _pool = psycopg2.pool.ThreadedConnectionPool(1, 10, DATABASE_URL)
    return _pool


class _Conn:
    """Context manager: borrows a connection from the pool."""
    def __enter__(self):
        self.conn = _get_pool().getconn()
        return self.conn

    def __exit__(self, exc_type, *_):
        if exc_type:
            try:
                self.conn.rollback()
            except Exception:
                pass
        _get_pool().putconn(self.conn)


# ── schema ────────────────────────────────────────────────────────────────────

def init_db():
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS sessions (
                    session_id    TEXT PRIMARY KEY,
                    created_at    TIMESTAMPTZ DEFAULT NOW(),
                    last_accessed TIMESTAMPTZ DEFAULT NOW()
                );

                CREATE TABLE IF NOT EXISTS session_state (
                    session_id TEXT REFERENCES sessions(session_id) ON DELETE CASCADE,
                    key        TEXT NOT NULL,
                    value      JSONB,
                    PRIMARY KEY (session_id, key)
                );

                CREATE TABLE IF NOT EXISTS session_template (
                    session_id     TEXT PRIMARY KEY
                                   REFERENCES sessions(session_id) ON DELETE CASCADE,
                    template_bytes BYTEA
                );
            """)
        conn.commit()


# ── rolling-buffer cleanup ────────────────────────────────────────────────────

def _db_size(cur) -> int:
    cur.execute("SELECT pg_database_size(current_database())")
    return cur.fetchone()[0]


def cleanup_if_needed() -> int:
    """Delete oldest sessions until DB is under MAX_DB_BYTES. Returns number deleted."""
    deleted = 0
    with _Conn() as conn:
        with conn.cursor() as cur:
            while _db_size(cur) > MAX_DB_BYTES:
                cur.execute("""
                    DELETE FROM sessions
                    WHERE session_id = (
                        SELECT session_id FROM sessions
                        ORDER BY last_accessed ASC
                        LIMIT 1
                    )
                """)
                if cur.rowcount == 0:
                    break
                deleted += 1
        conn.commit()
    return deleted


# ── session lifecycle ─────────────────────────────────────────────────────────

def ensure_session(session_id: str):
    """Create a session row if missing; run rolling-buffer cleanup first."""
    cleanup_if_needed()
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO sessions (session_id)
                VALUES (%s)
                ON CONFLICT (session_id) DO UPDATE
                    SET last_accessed = NOW()
            """, (session_id,))
        conn.commit()


def touch_session(session_id: str):
    """Update last_accessed so this session isn't the first to be evicted."""
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                UPDATE sessions SET last_accessed = NOW()
                WHERE session_id = %s
            """, (session_id,))
        conn.commit()


def delete_session(session_id: str):
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM sessions WHERE session_id = %s", (session_id,))
        conn.commit()


# ── state reads / writes (JSONB) ──────────────────────────────────────────────

_DEFAULTS = {
    "students": [],
    "next_id": 1,
    "excel_records": [],
    "excel_columns": [],
    "column_mapping": {},
    "editor_settings": {},
    "email_settings": {
        "smtp_host": "smtp.gmail.com",
        "smtp_port": 587,
        "sender_email": "",
        "app_password": "",
    },
    "email_template": {
        "subject": "Certificate of Participation",
        "body": (
            "Dear {name},\n\nThank you for participating.\n\n"
            "Your certificate is attached.\n\nRegards,\nEvent Team"
        ),
    },
}


def get_state(session_id: str) -> dict:
    """Return a complete state dict for this session (missing keys get defaults)."""
    import copy
    state = copy.deepcopy(_DEFAULTS)
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT key, value FROM session_state
                WHERE session_id = %s
            """, (session_id,))
            for key, value in cur.fetchall():
                state[key] = value
    return state


def save_key(session_id: str, key: str, value):
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO session_state (session_id, key, value)
                VALUES (%s, %s, %s)
                ON CONFLICT (session_id, key) DO UPDATE SET value = EXCLUDED.value
            """, (session_id, key, Json(value)))
        conn.commit()


def save_keys(session_id: str, updates: dict):
    with _Conn() as conn:
        with conn.cursor() as cur:
            for key, value in updates.items():
                cur.execute("""
                    INSERT INTO session_state (session_id, key, value)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (session_id, key) DO UPDATE SET value = EXCLUDED.value
                """, (session_id, key, Json(value)))
        conn.commit()


# ── template bytes (binary, stored separately) ────────────────────────────────

def get_template_bytes(session_id: str):
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                SELECT template_bytes FROM session_template
                WHERE session_id = %s
            """, (session_id,))
            row = cur.fetchone()
    return bytes(row[0]) if row else None


def save_template_bytes(session_id: str, data: bytes):
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO session_template (session_id, template_bytes)
                VALUES (%s, %s)
                ON CONFLICT (session_id) DO UPDATE
                    SET template_bytes = EXCLUDED.template_bytes
            """, (session_id, psycopg2.Binary(data)))
        conn.commit()


# ── utility ───────────────────────────────────────────────────────────────────

def get_db_stats() -> dict:
    with _Conn() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT pg_database_size(current_database())")
            size = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM sessions")
            sessions = cur.fetchone()[0]
    return {
        "db_size_bytes": size,
        "db_size_mb": round(size / (1024 * 1024), 2),
        "max_mb": round(MAX_DB_BYTES / (1024 * 1024), 2),
        "active_sessions": sessions,
        "pct_used": round((size / MAX_DB_BYTES) * 100, 1),
    }
