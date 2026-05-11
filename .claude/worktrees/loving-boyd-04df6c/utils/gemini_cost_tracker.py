"""
utils/gemini_cost_tracker.py
----------------------------
Lightweight SQLite logger for Gemini API token usage and cost.
Importable from any context (Flask request, background thread, standalone script).
"""

import logging
import os
import sqlite3
from datetime import datetime, timezone

logger = logging.getLogger(__name__)

# Pricing in USD per token (as of 2025-04)
_PRICING: dict = {
    "gemini-2.0-flash":       {"input": 0.075  / 1_000_000, "output": 0.30  / 1_000_000},
    "gemini-2.0-flash-lite":  {"input": 0.075  / 1_000_000, "output": 0.30  / 1_000_000},
    "gemini-1.5-flash":       {"input": 0.075  / 1_000_000, "output": 0.30  / 1_000_000},
    "gemini-1.5-flash-8b":    {"input": 0.0375 / 1_000_000, "output": 0.15  / 1_000_000},
    "gemini-1.5-pro":         {"input": 1.25   / 1_000_000, "output": 5.00  / 1_000_000},
    "gemini-2.5-pro":         {"input": 1.25   / 1_000_000, "output": 10.00 / 1_000_000},
    "gemini-2.0-pro":         {"input": 1.25   / 1_000_000, "output": 10.00 / 1_000_000},
}
_DEFAULT_PRICING = {"input": 0.075 / 1_000_000, "output": 0.30 / 1_000_000}

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
            CREATE TABLE IF NOT EXISTS gemini_api_usage (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                ts            TEXT    NOT NULL,
                feature       TEXT    NOT NULL,
                model         TEXT    NOT NULL,
                input_tokens  INTEGER NOT NULL DEFAULT 0,
                output_tokens INTEGER NOT NULL DEFAULT 0,
                cost_usd      REAL    NOT NULL DEFAULT 0.0
            )
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_gau_ts      ON gemini_api_usage(ts)")
        c.execute("CREATE INDEX IF NOT EXISTS idx_gau_feature ON gemini_api_usage(feature)")


def log_usage(feature: str, model: str, input_tokens: int, output_tokens: int) -> None:
    """Record one Gemini API call. Silently swallows all errors."""
    try:
        p = _PRICING.get(model, _DEFAULT_PRICING)
        cost = input_tokens * p["input"] + output_tokens * p["output"]
        ensure_table()
        with _conn() as c:
            c.execute(
                "INSERT INTO gemini_api_usage (ts, feature, model, input_tokens, output_tokens, cost_usd)"
                " VALUES (?, ?, ?, ?, ?, ?)",
                (datetime.now(timezone.utc).isoformat(), feature, model,
                 input_tokens, output_tokens, cost),
            )
    except Exception as exc:
        logger.warning("gemini_cost_tracker.log_usage failed: %s", exc)


def get_stats() -> dict:
    """Return aggregated stats per feature for the dashboard."""
    try:
        ensure_table()
        today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        features = ["reference_conversion", "book_indexer"]
        result: dict = {}
        with _conn() as c:
            for feat in features:
                row = c.execute(
                    """
                    SELECT
                        SUM(input_tokens),
                        SUM(output_tokens),
                        SUM(cost_usd),
                        COUNT(*),
                        SUM(CASE WHEN ts >= ? THEN cost_usd ELSE 0.0 END),
                        SUM(CASE WHEN ts >= ? THEN 1 ELSE 0 END),
                        MAX(model)
                    FROM gemini_api_usage
                    WHERE feature = ?
                    """,
                    (today, today, feat),
                ).fetchone()
                inp, out, total_cost, calls, today_cost, today_calls, last_model = row
                inp = inp or 0
                out = out or 0
                total_cost = total_cost or 0.0
                calls = calls or 0
                today_cost = today_cost or 0.0
                today_calls = today_calls or 0
                result[feat] = {
                    "all_time": {
                        "input_tokens": inp,
                        "output_tokens": out,
                        "total_tokens": inp + out,
                        "total_calls": calls,
                        "cost": {"total_cost": total_cost},
                    },
                    "today": {
                        "calls": today_calls,
                        "cost": {"total_cost": today_cost},
                    },
                    "averages": {
                        "cost_per_call": (total_cost / calls) if calls else 0.0,
                    },
                    "pricing": {"model": last_model or "gemini-2.0-flash"},
                }
        # Totals across all features
        row = c.execute(
            """
            SELECT SUM(cost_usd), COUNT(*),
                   SUM(CASE WHEN ts >= ? THEN cost_usd ELSE 0.0 END),
                   SUM(CASE WHEN ts >= ? THEN 1 ELSE 0 END)
            FROM gemini_api_usage
            """,
            (today, today),
        ).fetchone()
        result["totals"] = {
            "all_time_cost": row[0] or 0.0,
            "all_time_calls": row[1] or 0,
            "today_cost": row[2] or 0.0,
            "today_calls": row[3] or 0,
        }
        return result
    except Exception as exc:
        logger.warning("gemini_cost_tracker.get_stats failed: %s", exc)
        return {}


def get_recent_usage(limit: int = 20) -> list:
    """Return the most recent API call rows for the history table."""
    try:
        ensure_table()
        with _conn() as c:
            rows = c.execute(
                "SELECT ts, feature, model, input_tokens, output_tokens, cost_usd"
                " FROM gemini_api_usage ORDER BY ts DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [
            {
                "ts": r[0],
                "feature": r[1],
                "model": r[2],
                "input_tokens": r[3],
                "output_tokens": r[4],
                "cost_usd": r[5],
            }
            for r in rows
        ]
    except Exception as exc:
        logger.warning("gemini_cost_tracker.get_recent_usage failed: %s", exc)
        return []
