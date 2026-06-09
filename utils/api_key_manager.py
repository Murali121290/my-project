"""
utils/api_key_manager.py
------------------------
API key generation, validation, and management for external project access.
Keys are stored as SHA-256 hashes; the raw key is shown only once at creation.
"""

import hashlib
import logging
import os
import secrets
import sqlite3
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DB_PATH = os.environ.get(
    "GEMINI_USAGE_DB",
    os.path.join(_PROJECT_ROOT, "gemini_usage.db"),
)


def _conn() -> sqlite3.Connection:
    return sqlite3.connect(DB_PATH)


def ensure_table() -> None:
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS api_keys (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                key_hash     TEXT NOT NULL UNIQUE,
                project_name TEXT NOT NULL,
                created_at   TEXT NOT NULL,
                last_used    TEXT,
                active       INTEGER NOT NULL DEFAULT 1
            )
        """)


def _hash(raw_key: str) -> str:
    return hashlib.sha256(raw_key.encode()).hexdigest()


def generate_api_key(project_name: str) -> str:
    """Create a new API key for a project. Returns the raw key (shown once)."""
    raw_key = "pph_" + secrets.token_hex(32)
    ensure_table()
    with _conn() as c:
        c.execute(
            "INSERT INTO api_keys (key_hash, project_name, created_at) VALUES (?, ?, ?)",
            (_hash(raw_key), project_name, datetime.now(timezone.utc).isoformat()),
        )
    return raw_key


def validate_api_key(raw_key: str) -> dict | None:
    """Return {id, project_name} if key is valid and active, else None."""
    try:
        ensure_table()
        with _conn() as c:
            row = c.execute(
                "SELECT id, project_name FROM api_keys WHERE key_hash = ? AND active = 1",
                (_hash(raw_key),),
            ).fetchone()
        if row:
            _touch(row[0])
            return {"id": row[0], "project_name": row[1]}
        return None
    except Exception as exc:
        logger.warning("validate_api_key failed: %s", exc)
        return None


def _touch(key_id: int) -> None:
    try:
        with _conn() as c:
            c.execute(
                "UPDATE api_keys SET last_used = ? WHERE id = ?",
                (datetime.now(timezone.utc).isoformat(), key_id),
            )
    except Exception:
        pass


def list_api_keys() -> list:
    """Return all API keys (without the hash) for the admin UI."""
    try:
        ensure_table()
        with _conn() as c:
            rows = c.execute(
                "SELECT id, project_name, created_at, last_used, active FROM api_keys ORDER BY id DESC"
            ).fetchall()
        return [
            {
                "id": r[0],
                "project_name": r[1],
                "created_at": r[2],
                "last_used": r[3],
                "active": bool(r[4]),
            }
            for r in rows
        ]
    except Exception as exc:
        logger.warning("list_api_keys failed: %s", exc)
        return []


def revoke_api_key(key_id: int) -> None:
    """Deactivate an API key."""
    try:
        ensure_table()
        with _conn() as c:
            c.execute("UPDATE api_keys SET active = 0 WHERE id = ?", (key_id,))
    except Exception as exc:
        logger.warning("revoke_api_key failed: %s", exc)
