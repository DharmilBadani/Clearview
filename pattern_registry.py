"""Small deterministic pattern registry for user-confirmed data corrections."""

from __future__ import annotations

import os
import sqlite3
from pathlib import Path


def _db_path() -> Path:
    configured = os.environ.get("CLEARVIEW_PATTERN_DB", "").strip()
    if configured:
        return Path(configured)
    return Path(__file__).resolve().parent / "output" / "patterns.sqlite3"


def _key(value: object) -> str:
    return " ".join(str(value).replace("\u00a0", " ").split()).casefold()


def _connect() -> sqlite3.Connection:
    path = _db_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS patterns (
            column_name TEXT NOT NULL,
            pattern_type TEXT NOT NULL,
            source_value TEXT NOT NULL,
            normalized_value TEXT NOT NULL,
            confirmations INTEGER NOT NULL DEFAULT 1,
            PRIMARY KEY (column_name, pattern_type, source_value)
        )
        """
    )
    return connection


def record_pattern(column_name: str, pattern_type: str, source_value: str, normalized_value: str) -> dict[str, object]:
    column_name = str(column_name).strip()
    pattern_type = str(pattern_type).strip().casefold()
    source_value = _key(source_value)
    normalized_value = str(normalized_value).strip()
    if not column_name or not pattern_type or not source_value or not normalized_value:
        raise ValueError("A pattern needs a column, type, source value, and normalized value.")
    connection = _connect()
    try:
        connection.execute(
            """
            INSERT INTO patterns (column_name, pattern_type, source_value, normalized_value)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(column_name, pattern_type, source_value) DO UPDATE SET
                normalized_value = excluded.normalized_value,
                confirmations = patterns.confirmations + 1
            """,
            (column_name, pattern_type, source_value, normalized_value),
        )
        row = connection.execute(
            """SELECT column_name, pattern_type, source_value, normalized_value, confirmations
               FROM patterns WHERE column_name = ? AND pattern_type = ? AND source_value = ?""",
            (column_name, pattern_type, source_value),
        ).fetchone()
        connection.commit()
    finally:
        connection.close()
    return {
        "column": row[0],
        "pattern_type": row[1],
        "source_value": row[2],
        "normalized_value": row[3],
        "confirmations": row[4],
    }


def confirmed_replacements(column_name: str, pattern_type: str, minimum_confirmations: int = 2) -> dict[str, str]:
    connection = _connect()
    try:
        rows = connection.execute(
            """SELECT source_value, normalized_value FROM patterns
               WHERE column_name = ? AND pattern_type = ? AND confirmations >= ?""",
            (str(column_name), str(pattern_type).casefold(), int(minimum_confirmations)),
        ).fetchall()
    finally:
        connection.close()
    return {row[0]: row[1] for row in rows}


def list_patterns() -> list[dict[str, object]]:
    connection = _connect()
    try:
        rows = connection.execute(
            """SELECT column_name, pattern_type, source_value, normalized_value, confirmations
               FROM patterns ORDER BY column_name, pattern_type, source_value"""
        ).fetchall()
    finally:
        connection.close()
    return [
        {
            "column": row[0],
            "pattern_type": row[1],
            "source_value": row[2],
            "normalized_value": row[3],
            "confirmations": row[4],
        }
        for row in rows
    ]
