"""
Constants and configuration for WhatsApp Bulk Sender.
"""

import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.txt")
DEFAULT_INPUT = os.path.join(SCRIPT_DIR, "contacts.xlsx")
CHROME_PROFILE_DIR = os.path.join(SCRIPT_DIR, "chrome_profile")

DELAY_BETWEEN_MESSAGES = 20
MAX_RETRIES = 2

# Status constants
STATUS_SENT = "Sent"
STATUS_FAILED = "Failed"
STATUS_NO_WA = "No WhatsApp"
STATUS_PENDING = "Pending"
STATUS_DRY_RUN = "Dry Run"

STATUS_ICONS = {
    STATUS_SENT: "[OK]",
    STATUS_FAILED: "[FAIL]",
    STATUS_NO_WA: "[NO WA]",
    STATUS_PENDING: "[...]",
    STATUS_DRY_RUN: "[TEST]",
}

STATUS_COLORS = {
    STATUS_SENT: "C8E6C9",
    STATUS_FAILED: "FFCDD2",
    STATUS_NO_WA: "FFE0B2",
    STATUS_PENDING: "E0E0E0",
    STATUS_DRY_RUN: "BBDEFB",
}

HEADER_COLOR = "1565C0"
