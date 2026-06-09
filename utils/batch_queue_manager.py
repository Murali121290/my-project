"""
utils/batch_queue_manager.py
-----------------------------
Batch queue for tracking all API calls from external projects.
Status lifecycle: pending → processing → completed | failed
"""

import logging
import os
import sqlite3
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
DB_PATH = os.environ.get(
    "GEMINI_USAGE_DB",
    os.path.join(_PROJECT_ROOT, "gemini_usage.db"),
)

STATUS_PENDING    = "pending"
STATUS_PROCESSING = "processing"
STATUS_COMPLETED  = "completed"
STATUS_FAILED     = "failed"


def _conn() -> sqlite3.Connection:
    return sqlite3.connect(DB_PATH)


def ensure_table() -> None:
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS batch_queue (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                ts           TEXT NOT NULL,
                project_name TEXT NOT NULL,
                endpoint     TEXT NOT NULL,
                method       TEXT NOT NULL,
                status       TEXT NOT NULL DEFAULT 'pending',
                payload_size INTEGER DEFAULT 0,
                result       TEXT,
                processed_at TEXT
            )
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_bq_ts      ON batch_queue(ts)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_bq_project ON batch_queue(project_name)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_bq_status  ON batch_queue(status)")


def enqueue(project_name: str, endpoint: str, method: str, payload_size: int = 0) -> int:
    """Add a new entry to the queue. Returns the queue row id."""
    try:
        ensure_table()
        with _conn() as c:
            cur = c.execute(
                "INSERT INTO batch_queue (ts, project_name, endpoint, method, status, payload_size)"
                " VALUES (?, ?, ?, ?, ?, ?)",
                (datetime.now(timezone.utc).isoformat(), project_name,
                 endpoint, method, STATUS_PENDING, payload_size),
            )
            return cur.lastrowid
    except Exception as exc:
        logger.warning("batch_queue.enqueue failed: %s", exc)
        return -1


def update_status(queue_id: int, status: str, result: str | None = None) -> None:
    """Update the status and optional result of a queued entry."""
    if queue_id < 0:
        return
    try:
        with _conn() as c:
            c.execute(
                "UPDATE batch_queue SET status = ?, result = ?, processed_at = ? WHERE id = ?",
                (status, result, datetime.now(timezone.utc).isoformat(), queue_id),
            )
    except Exception as exc:
        logger.warning("batch_queue.update_status failed: %s", exc)


def get_queue(status: str | None = None, project: str | None = None, limit: int = 50) -> list:
    """Return recent queue entries, optionally filtered by status or project."""
    try:
        ensure_table()
        query = "SELECT id, ts, project_name, endpoint, method, status, payload_size, result, processed_at FROM batch_queue"
        params = []
        conditions = []
        if status:
            conditions.append("status = ?")
            params.append(status)
        if project:
            conditions.append("project_name = ?")
            params.append(project)
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += " ORDER BY ts DESC LIMIT ?"
        params.append(limit)
        with _conn() as c:
            rows = c.execute(query, params).fetchall()
        return [
            {
                "id": r[0],
                "ts": r[1],
                "project_name": r[2],
                "endpoint": r[3],
                "method": r[4],
                "status": r[5],
                "payload_size": r[6],
                "result": r[7],
                "processed_at": r[8],
            }
            for r in rows
        ]
    except Exception as exc:
        logger.warning("batch_queue.get_queue failed: %s", exc)
        return []


def get_stats() -> dict:
    """Return counts per status and per project for the dashboard."""
    try:
        ensure_table()
        with _conn() as c:
            status_rows = c.execute(
                "SELECT status, COUNT(*) FROM batch_queue GROUP BY status"
            ).fetchall()
            project_rows = c.execute(
                "SELECT project_name, COUNT(*), "
                "SUM(CASE WHEN status='completed' THEN 1 ELSE 0 END), "
                "SUM(CASE WHEN status='failed' THEN 1 ELSE 0 END) "
                "FROM batch_queue GROUP BY project_name"
            ).fetchall()
        by_status = {r[0]: r[1] for r in status_rows}
        by_project = [
            {"project_name": r[0], "total": r[1], "completed": r[2], "failed": r[3]}
            for r in project_rows
        ]
        return {
            "total": sum(by_status.values()),
            "pending": by_status.get(STATUS_PENDING, 0),
            "processing": by_status.get(STATUS_PROCESSING, 0),
            "completed": by_status.get(STATUS_COMPLETED, 0),
            "failed": by_status.get(STATUS_FAILED, 0),
            "by_project": by_project,
        }
    except Exception as exc:
        logger.warning("batch_queue.get_stats failed: %s", exc)
        return {"total": 0, "pending": 0, "processing": 0, "completed": 0, "failed": 0, "by_project": []}
