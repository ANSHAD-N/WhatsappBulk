"""
WhatsApp Bulk Message Sender
=============================
Reads phone numbers from an Excel file, sends a common message
(from config.txt) to all contacts via WhatsApp Web automatically.

Tracks status in:
  - SQLite database (whatsapp_sender.db)
  - Output Excel file (status_report_<timestamp>.xlsx)

Usage:
  python whatsapp_sender.py                     # Send messages for real
  python whatsapp_sender.py --dry-run           # Test without sending
  python whatsapp_sender.py --input other.xlsx  # Use a different input file
  python whatsapp_sender.py --retries 3         # Retry failed messages up to 3 times

Requirements:
  - pip install pywhatkit openpyxl
  - Chrome browser installed
  - WhatsApp Web logged in (first time: scan QR code)
"""

import os
import sys
import re
import time
import shutil
import argparse
import uuid
import subprocess
from datetime import datetime

# ─── Dependency Checks ───────────────────────────────────────

MISSING_DEPS = []
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    MISSING_DEPS.append("openpyxl")

try:
    import pywhatkit
except ImportError:
    MISSING_DEPS.append("pywhatkit")

if MISSING_DEPS:
    print(f"Missing dependencies: {', '.join(MISSING_DEPS)}")
    print(f"Run: pip install {' '.join(MISSING_DEPS)}")
    sys.exit(1)

from database import init_db, insert_record, update_record, get_all_records, get_summary

# ─── Constants ───────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.txt")
DEFAULT_INPUT = os.path.join(SCRIPT_DIR, "contacts.xlsx")
DELAY_BETWEEN_MESSAGES = 15
MAX_RETRIES = 2

# Status constants (single source of truth)
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
    STATUS_SENT: "C8E6C9",       # Green
    STATUS_FAILED: "FFCDD2",     # Red
    STATUS_NO_WA: "FFE0B2",      # Orange
    STATUS_PENDING: "E0E0E0",    # Grey
    STATUS_DRY_RUN: "BBDEFB",    # Blue
}

HEADER_COLOR = "1565C0"

# Error classification keywords
NO_WHATSAPP_KEYWORDS = [
    "not on whatsapp", "invalid", "not registered",
    "phone number shared via url is invalid",
    "couldn't find", "unable to find",
]
NETWORK_KEYWORDS = [
    "timeout", "timed out", "connection", "network",
    "no internet", "err_internet", "err_name_not_resolved",
]
BROWSER_KEYWORDS = [
    "browser", "chrome", "webdriver", "selenium",
    "no such window", "session not created",
]
WHATSAPP_NOT_OPEN_KEYWORDS = [
    "whatsapp", "qr", "scan", "not logged in",
    "landing", "startup", "retry",
]


# ─── Utility Helpers ────────────────────────────────────────

def log(msg, end="\n"):
    """Print a log message with flush."""
    print(msg, end=end, flush=True)


def fatal(msg):
    """Print error and exit."""
    log(f"ERROR: {msg}")
    sys.exit(1)


def check_chrome_installed():
    """Check if Chrome browser is available."""
    chrome_paths = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
    ]
    for path in chrome_paths:
        if os.path.exists(path):
            return True
    # Fallback: check if 'chrome' is on PATH
    return shutil.which("chrome") is not None or shutil.which("google-chrome") is not None


def check_internet():
    """Check if internet is available."""
    try:
        result = subprocess.run(
            ["ping", "-n", "1", "-w", "3000", "8.8.8.8"],
            capture_output=True, timeout=5
        )
        return result.returncode == 0
    except Exception:
        return False


# ─── Core Functions ──────────────────────────────────────────

def load_message():
    """Load the common message from config.txt."""
    if not os.path.exists(CONFIG_FILE):
        fatal(f"Message file not found: {CONFIG_FILE}\n   Create a config.txt file with your message.")

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        message = f.read().strip()

    if not message:
        fatal("config.txt is empty. Please write your message in it.")
    return message


def validate_phone(phone):
    """
    Validate and normalize phone number.
    Returns (is_valid, normalized_number, error_reason)
    """
    phone = re.sub(r"[\s\-\.\(\)]", "", str(phone).strip())

    if not phone.startswith("+"):
        if re.match(r"^[6-9]\d{9}$", phone):
            phone = "+91" + phone
        elif re.match(r"^91[6-9]\d{9}$", phone):
            phone = "+" + phone
        else:
            return False, phone, "Invalid format: must include country code (e.g., +919876543210)"

    digits = re.sub(r"\D", "", phone)
    if len(digits) < 10:
        return False, phone, f"Too short: {len(digits)} digits"
    if len(digits) > 15:
        return False, phone, f"Too long: {len(digits)} digits"

    return True, phone, None


def load_contacts(filepath):
    """Load contacts from Excel file. Returns list of (name, phone) tuples."""
    if not os.path.exists(filepath):
        fatal(f"Input file not found: {filepath}\n   Run: python generate_test_data.py")

    try:
        wb = load_workbook(filepath)
        ws = wb.active
    except Exception as e:
        fatal(f"Cannot read Excel file: {e}")

    headers = [str(c.value).strip().lower() if c.value else "" for c in ws[1]]

    name_col = next((i for i, h in enumerate(headers) if "name" in h), None)
    phone_col = next((i for i, h in enumerate(headers) if any(k in h for k in ("phone", "number", "mobile"))), None)

    if phone_col is None:
        fatal(f"No 'Phone'/'Number'/'Mobile' column found. Headers: {headers}")

    contacts = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[phone_col] is None:
            continue
        name = str(row[name_col]) if name_col is not None and row[name_col] else "Unknown"
        contacts.append((name, str(row[phone_col]).strip()))

    wb.close()
    return contacts


def classify_error(error_msg):
    """Classify an error message into a status and clean description."""
    msg = error_msg.lower()

    if any(k in msg for k in NO_WHATSAPP_KEYWORDS):
        return STATUS_NO_WA, error_msg
    if any(k in msg for k in WHATSAPP_NOT_OPEN_KEYWORDS):
        return STATUS_FAILED, f"WhatsApp Web not ready: {error_msg}"
    if any(k in msg for k in NETWORK_KEYWORDS):
        return STATUS_FAILED, f"Network error: {error_msg}"
    if any(k in msg for k in BROWSER_KEYWORDS):
        return STATUS_FAILED, f"Browser error: {error_msg}"

    return STATUS_FAILED, error_msg


def send_message(phone, message, dry_run=False, retries=MAX_RETRIES):
    """
    Send a WhatsApp message with retry logic.
    Returns (status, error_details)
    """
    if dry_run:
        return STATUS_DRY_RUN, None

    last_error = None
    for attempt in range(1, retries + 1):
        try:
            pywhatkit.sendwhatmsg_instantly(
                phone_no=phone,
                message=message,
                wait_time=15,
                tab_close=True,
                close_time=5,
            )
            return STATUS_SENT, None

        except pywhatkit.core.exceptions.CountryCodeException:
            return STATUS_FAILED, "Invalid country code"

        except pywhatkit.core.exceptions.CallTimeException:
            return STATUS_FAILED, "Call time error"

        except Exception as e:
            last_error = str(e)
            status, detail = classify_error(last_error)

            # Don't retry "No WhatsApp" — it won't change
            if status == STATUS_NO_WA:
                return status, detail

            # Retry for network/browser/transient errors
            if attempt < retries:
                log(f"\n   Attempt {attempt} failed: {detail}. Retrying in 10s...")
                time.sleep(10)
            else:
                return status, f"{detail} (after {retries} attempts)"

    return STATUS_FAILED, last_error


# ─── Excel Report ────────────────────────────────────────────

def _style_cell(cell, font=None, alignment=None, fill=None, border=None):
    """Apply styles to a cell (DRY helper)."""
    if font:
        cell.font = font
    if alignment:
        cell.alignment = alignment
    if fill:
        cell.fill = fill
    if border:
        cell.border = border
    return cell


def create_status_excel(batch_id, records, message):
    """Create output status report Excel file."""
    filename = f"status_report_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.xlsx"
    filepath = os.path.join(SCRIPT_DIR, filename)

    wb = Workbook()
    ws = wb.active
    ws.title = "Status Report"

    # Shared styles
    border = Border(*(Side(style="thin") for _ in range(4)))
    hdr_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    hdr_fill = PatternFill(start_color=HEADER_COLOR, end_color=HEADER_COLOR, fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center")
    data_font = Font(name="Calibri", size=11)
    center = Alignment(horizontal="center")

    # Headers
    headers = ["Sr No", "Name", "Phone", "Status", "Error Details", "Timestamp"]
    for col, header in enumerate(headers, 1):
        _style_cell(ws.cell(row=1, column=col, value=header),
                    font=hdr_font, fill=hdr_fill, alignment=hdr_align, border=border)

    # Data rows
    for i, rec in enumerate(records, 1):
        row = i + 1
        _style_cell(ws.cell(row=row, column=1, value=i), font=data_font, alignment=center, border=border)
        _style_cell(ws.cell(row=row, column=2, value=rec["name"]), font=data_font, border=border)
        _style_cell(ws.cell(row=row, column=3, value=rec["phone"]), font=data_font, border=border)

        status_val = rec["status"]
        color = STATUS_COLORS.get(status_val, "FFFFFF")
        _style_cell(
            ws.cell(row=row, column=4, value=status_val),
            font=Font(name="Calibri", size=11, bold=True),
            alignment=center,
            fill=PatternFill(start_color=color, end_color=color, fill_type="solid"),
            border=border,
        )
        _style_cell(ws.cell(row=row, column=5, value=rec["error_details"] or ""), font=data_font, border=border)
        _style_cell(ws.cell(row=row, column=6, value=rec["timestamp"]), font=data_font, border=border)

    # Common message at bottom
    msg_row = len(records) + 3
    _style_cell(ws.cell(row=msg_row, column=1, value="Common Message:"),
                font=Font(name="Calibri", size=11, bold=True))
    _style_cell(ws.cell(row=msg_row, column=2, value=message),
                font=Font(name="Calibri", size=11, italic=True))
    ws.merge_cells(start_row=msg_row, start_column=2, end_row=msg_row, end_column=6)

    # Column widths
    for col, w in zip("ABCDEF", [8, 25, 20, 15, 40, 22]):
        ws.column_dimensions[col].width = w

    wb.save(filepath)
    return filepath


# ─── Main ────────────────────────────────────────────────────

def run_preflight_checks(dry_run):
    """Run all checks before sending messages."""
    errors = []

    if not dry_run:
        if not check_chrome_installed():
            errors.append("Chrome browser not found. Please install Google Chrome.")
        if not check_internet():
            errors.append("No internet connection. Please check your network.")

    if errors:
        log("\n--- Pre-flight Check Failed ---")
        for err in errors:
            log(f"  - {err}")
        fatal("Fix the above issues and try again.")

    if not dry_run:
        log("Pre-flight checks passed (Chrome + Internet OK)")


def process_contact(name, phone, message, batch_id, dry_run, retries):
    """Process a single contact: validate, send, update DB. Returns status."""
    is_valid, normalized, error = validate_phone(phone)
    if not is_valid:
        rid = insert_record(name, phone, batch_id)
        update_record(rid, STATUS_FAILED, error)
        return STATUS_FAILED, error

    rid = insert_record(name, normalized, batch_id)

    try:
        status, error = send_message(normalized, message, dry_run=dry_run, retries=retries)
    except KeyboardInterrupt:
        update_record(rid, STATUS_FAILED, "Interrupted by user (Ctrl+C)")
        raise

    update_record(rid, status, error)
    return status, error


def print_summary(batch_id, total):
    """Print the final summary from the database."""
    summary = get_summary(batch_id)
    log("\nSummary:")
    log(f"  Total:          {total}")
    for status, count in summary.items():
        icon = STATUS_ICONS.get(status, "[?]")
        log(f"  {icon} {status:14s} {count}")


def main():
    parser = argparse.ArgumentParser(description="WhatsApp Bulk Message Sender")
    parser.add_argument("--dry-run", action="store_true", help="Test without actually sending messages")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Path to input Excel file (default: contacts.xlsx)")
    parser.add_argument("--retries", type=int, default=MAX_RETRIES, help=f"Max retry attempts per message (default: {MAX_RETRIES})")
    args = parser.parse_args()

    log("=" * 60)
    log("   WhatsApp Bulk Message Sender")
    log("=" * 60)
    if args.dry_run:
        log("   MODE: DRY RUN (no messages will be sent)")
    log("")

    # 1. Pre-flight checks
    run_preflight_checks(args.dry_run)

    # 2. Load message & contacts
    message = load_message()
    log(f"Message: \"{message[:80]}{'...' if len(message) > 80 else ''}\"")

    contacts = load_contacts(args.input)
    log(f"Loaded {len(contacts)} contacts from: {os.path.basename(args.input)}")

    if not contacts:
        log("No contacts found in the file.")
        sys.exit(0)

    # 3. Init DB + batch
    init_db()
    batch_id = str(uuid.uuid4())[:8]
    log(f"Batch ID: {batch_id}")
    log("-" * 60)

    # 4. Process contacts
    interrupted = False
    for i, (name, phone) in enumerate(contacts, 1):
        log(f"\n[{i}/{len(contacts)}] {name} ({phone}) ... ", end="")

        try:
            status, error = process_contact(name, phone, message, batch_id, args.dry_run, args.retries)
        except KeyboardInterrupt:
            log("INTERRUPTED!")
            log("\nSending stopped by user. Saving progress...")
            interrupted = True
            break

        icon = STATUS_ICONS.get(status, "[?]")
        log(f"{icon} {status}", end="")
        if error:
            log(f" - {error}", end="")
        log("")

        # Delay between messages (skip last + dry run)
        if not args.dry_run and i < len(contacts):
            log(f"   Waiting {DELAY_BETWEEN_MESSAGES}s before next message...")
            try:
                time.sleep(DELAY_BETWEEN_MESSAGES)
            except KeyboardInterrupt:
                log("\nSending stopped by user. Saving progress...")
                interrupted = True
                break

    # 5. Generate report
    log("\n" + "=" * 60)
    records = get_all_records(batch_id)
    try:
        report = create_status_excel(batch_id, records, message)
        log(f"Status report: {os.path.basename(report)}")
    except Exception as e:
        log(f"Warning: Could not create Excel report: {e}")

    # 6. Summary
    print_summary(batch_id, len(contacts))
    log(f"\nDatabase: whatsapp_sender.db (Batch: {batch_id})")
    if interrupted:
        log("Note: Sending was interrupted. Run again to continue with remaining contacts.")
    log("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("\n\nProgram terminated by user.")
        sys.exit(0)
    except Exception as e:
        log(f"\nUnexpected error: {e}")
        log("Please check your setup and try again.")
        sys.exit(1)
