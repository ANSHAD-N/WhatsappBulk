"""
SQLite database layer for WhatsApp Bulk Sender.
Stores contact numbers and their message delivery status.
Uses context managers for safe connection handling.
"""

import sqlite3
import os
from datetime import datetime
from contextlib import contextmanager

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_sender.db")


@contextmanager
def get_connection():
    """Context manager for safe database connections."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except sqlite3.Error as e:
        conn.rollback()
        raise e
    finally:
        conn.close()


def _now():
    """Get current timestamp string."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():
    """Initialize the database and create the messages table if it doesn't exist."""
    with get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS messages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                phone TEXT NOT NULL,
                status TEXT DEFAULT 'Pending',
                error_details TEXT,
                timestamp TEXT,
                batch_id TEXT
            )
        """)


def insert_record(name, phone, batch_id):
    """Insert a new contact record with 'Pending' status."""
    with get_connection() as conn:
        cursor = conn.execute(
            "INSERT INTO messages (name, phone, status, timestamp, batch_id) VALUES (?, ?, 'Pending', ?, ?)",
            (name, phone, _now(), batch_id)
        )
        return cursor.lastrowid


def update_record(record_id, status, error_details=None):
    """Update the status and error details of a record."""
    with get_connection() as conn:
        conn.execute(
            "UPDATE messages SET status = ?, error_details = ?, timestamp = ? WHERE id = ?",
            (status, error_details, _now(), record_id)
        )


def get_all_records(batch_id=None):
    """Get all records, optionally filtered by batch_id."""
    with get_connection() as conn:
        if batch_id:
            return conn.execute("SELECT * FROM messages WHERE batch_id = ? ORDER BY id", (batch_id,)).fetchall()
        return conn.execute("SELECT * FROM messages ORDER BY id DESC").fetchall()


def get_summary(batch_id):
    """Get a summary of statuses for a given batch."""
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT status, COUNT(*) as count FROM messages WHERE batch_id = ? GROUP BY status",
            (batch_id,)
        ).fetchall()
        return {row["status"]: row["count"] for row in rows}
