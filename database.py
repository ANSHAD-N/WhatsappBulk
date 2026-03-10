"""
SQLite database layer for WhatsApp Bulk Sender.
Stores contact numbers and their message delivery status.
"""

import sqlite3
import os
from datetime import datetime

DB_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "whatsapp_sender.db")


def get_connection():
    """Get a connection to the SQLite database."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Initialize the database and create the messages table if it doesn't exist."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
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
    conn.commit()
    conn.close()


def insert_record(name, phone, batch_id):
    """Insert a new contact record with 'Pending' status."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO messages (name, phone, status, timestamp, batch_id) VALUES (?, ?, 'Pending', ?, ?)",
        (name, phone, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), batch_id)
    )
    record_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return record_id


def update_record(record_id, status, error_details=None):
    """Update the status and error details of a record."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE messages SET status = ?, error_details = ?, timestamp = ? WHERE id = ?",
        (status, error_details, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), record_id)
    )
    conn.commit()
    conn.close()


def get_all_records(batch_id=None):
    """Get all records, optionally filtered by batch_id."""
    conn = get_connection()
    cursor = conn.cursor()
    if batch_id:
        cursor.execute("SELECT * FROM messages WHERE batch_id = ? ORDER BY id", (batch_id,))
    else:
        cursor.execute("SELECT * FROM messages ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows


def get_summary(batch_id):
    """Get a summary of statuses for a given batch."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT status, COUNT(*) as count FROM messages WHERE batch_id = ? GROUP BY status",
        (batch_id,)
    )
    summary = {row["status"]: row["count"] for row in cursor.fetchall()}
    conn.close()
    return summary
