"""
WhatsApp Bulk Message Sender
=============================
Reads phone numbers from an Excel file, sends a common message
(from config.txt) to all contacts via WhatsApp Web automatically.

Tracks status in:
  - SQLite database (whatsapp_sender.db)
  - Output Excel file (status_report_<timestamp>.xlsx)

Usage:
  python whatsapp_sender.py                  # Send messages for real
  python whatsapp_sender.py --dry-run        # Test without sending
  python whatsapp_sender.py --input other.xlsx  # Use a different input file

Requirements:
  - pip install pywhatkit openpyxl
  - Chrome browser installed
  - WhatsApp Web logged in (first time: scan QR code)
"""

import os
import sys
import re
import time
import argparse
import uuid
from datetime import datetime

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
except ImportError:
    print("❌ Error: openpyxl is not installed. Run: pip install openpyxl")
    sys.exit(1)

try:
    import pywhatkit
except ImportError:
    print("❌ Error: pywhatkit is not installed. Run: pip install pywhatkit")
    sys.exit(1)

from database import init_db, insert_record, update_record, get_all_records, get_summary

# ─── Configuration ───────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.txt")
DEFAULT_INPUT = os.path.join(SCRIPT_DIR, "contacts.xlsx")
DELAY_BETWEEN_MESSAGES = 15  # seconds between each message


# ─── Helpers ─────────────────────────────────────────────────

def load_message():
    """Load the common message from config.txt."""
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ Error: Message file not found: {CONFIG_FILE}")
        print("   Create a config.txt file with your message.")
        sys.exit(1)

    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        message = f.read().strip()

    if not message:
        print("❌ Error: config.txt is empty. Please write your message in it.")
        sys.exit(1)

    return message


def validate_phone(phone):
    """
    Validate and normalize phone number.
    Returns (is_valid, normalized_number, error_reason)
    """
    # Convert to string and strip whitespace
    phone = str(phone).strip()

    # Remove any spaces, dashes, dots
    phone = re.sub(r"[\s\-\.\(\)]", "", phone)

    # Must start with + and country code
    if not phone.startswith("+"):
        # Try to add +91 if it looks like a 10-digit Indian number
        if re.match(r"^[6-9]\d{9}$", phone):
            phone = "+91" + phone
        elif re.match(r"^91[6-9]\d{9}$", phone):
            phone = "+" + phone
        else:
            return False, phone, "Invalid format: must start with + and country code (e.g., +919876543210)"

    # Check if number has enough digits (min 10 digits after +country code)
    digits_only = re.sub(r"\D", "", phone)
    if len(digits_only) < 10:
        return False, phone, f"Too short: only {len(digits_only)} digits"
    if len(digits_only) > 15:
        return False, phone, f"Too long: {len(digits_only)} digits"

    return True, phone, None


def load_contacts(filepath):
    """Load contacts from Excel file. Returns list of (name, phone) tuples."""
    if not os.path.exists(filepath):
        print(f"❌ Error: Input file not found: {filepath}")
        print("   Run: python generate_test_data.py")
        sys.exit(1)

    try:
        wb = load_workbook(filepath)
        ws = wb.active
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        sys.exit(1)

    contacts = []
    headers = [str(cell.value).strip().lower() if cell.value else "" for cell in ws[1]]

    # Find Name and Phone columns
    name_col = None
    phone_col = None
    for i, h in enumerate(headers):
        if "name" in h:
            name_col = i
        if "phone" in h or "number" in h or "mobile" in h:
            phone_col = i

    if phone_col is None:
        print("❌ Error: No 'Phone' / 'Number' / 'Mobile' column found in Excel.")
        print(f"   Found headers: {headers}")
        sys.exit(1)

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[phone_col] is None:
            continue
        name = str(row[name_col]) if name_col is not None and row[name_col] else "Unknown"
        phone = str(row[phone_col]).strip()
        contacts.append((name, phone))

    wb.close()
    return contacts


def create_status_excel(batch_id, records, message):
    """Create output status report Excel file."""
    timestamp_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    filename = f"status_report_{timestamp_str}.xlsx"
    filepath = os.path.join(SCRIPT_DIR, filename)

    wb = Workbook()
    ws = wb.active
    ws.title = "Status Report"

    # --- Styling ---
    header_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    header_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    data_font = Font(name="Calibri", size=11)

    # Status colors
    status_colors = {
        "Sent": PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid"),       # Green
        "Failed": PatternFill(start_color="FFCDD2", end_color="FFCDD2", fill_type="solid"),      # Red
        "No WhatsApp": PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid"), # Orange
        "Pending": PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid"),     # Grey
        "Dry Run": PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid"),     # Blue
    }

    header_colors = {
        "Sr No": "1565C0",
        "Name": "1565C0",
        "Phone": "1565C0",
        "Status": "1565C0",
        "Error Details": "1565C0",
        "Timestamp": "1565C0",
    }

    # --- Headers ---
    headers = ["Sr No", "Name", "Phone", "Status", "Error Details", "Timestamp"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = PatternFill(start_color=header_colors[header], end_color=header_colors[header], fill_type="solid")
        cell.alignment = header_alignment
        cell.border = thin_border

    # --- Data rows ---
    for i, record in enumerate(records, 1):
        row = i + 1
        ws.cell(row=row, column=1, value=i).font = data_font
        ws.cell(row=row, column=1).alignment = Alignment(horizontal="center")
        ws.cell(row=row, column=1).border = thin_border

        ws.cell(row=row, column=2, value=record["name"]).font = data_font
        ws.cell(row=row, column=2).border = thin_border

        ws.cell(row=row, column=3, value=record["phone"]).font = data_font
        ws.cell(row=row, column=3).border = thin_border

        status_val = record["status"]
        status_cell = ws.cell(row=row, column=4, value=status_val)
        status_cell.font = Font(name="Calibri", size=11, bold=True)
        status_cell.alignment = Alignment(horizontal="center")
        status_cell.fill = status_colors.get(status_val, PatternFill())
        status_cell.border = thin_border

        ws.cell(row=row, column=5, value=record["error_details"] or "").font = data_font
        ws.cell(row=row, column=5).border = thin_border

        ws.cell(row=row, column=6, value=record["timestamp"]).font = data_font
        ws.cell(row=row, column=6).border = thin_border

    # --- Add message info at bottom ---
    msg_row = len(records) + 3
    ws.cell(row=msg_row, column=1, value="Common Message:").font = Font(name="Calibri", size=11, bold=True)
    ws.cell(row=msg_row, column=2, value=message).font = Font(name="Calibri", size=11, italic=True)
    ws.merge_cells(start_row=msg_row, start_column=2, end_row=msg_row, end_column=6)

    # --- Column widths ---
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 22

    wb.save(filepath)
    return filepath


def send_whatsapp_message(phone, message, dry_run=False):
    """
    Send a WhatsApp message to a phone number.
    Returns (status, error_details)
    """
    if dry_run:
        return "Dry Run", None

    try:
        # sendwhatmsg_instantly sends the message immediately
        # wait_time: seconds to wait for WhatsApp Web to load
        # tab_close: auto-close browser tab after sending
        # close_time: seconds to wait before closing the tab
        pywhatkit.sendwhatmsg_instantly(
            phone_no=phone,
            message=message,
            wait_time=12,      # Wait for WhatsApp Web to load
            tab_close=True,    # Auto-close browser tab
            close_time=5       # Wait 5 sec before closing tab
        )
        return "Sent", None

    except pywhatkit.core.exceptions.CountryCodeException:
        return "Failed", "Invalid country code"

    except pywhatkit.core.exceptions.CallTimeException:
        return "Failed", "Call time error - invalid time parameters"

    except Exception as e:
        error_msg = str(e).lower()

        # Detect "No WhatsApp" scenarios
        if any(keyword in error_msg for keyword in [
            "not on whatsapp", "invalid", "not registered",
            "phone number shared via url is invalid",
            "couldn't find", "unable to find"
        ]):
            return "No WhatsApp", str(e)

        # Network / timeout errors
        if any(keyword in error_msg for keyword in [
            "timeout", "timed out", "connection", "network",
            "no internet", "err_internet"
        ]):
            return "Failed", f"Network error: {e}"

        # Browser errors
        if any(keyword in error_msg for keyword in [
            "browser", "chrome", "webdriver", "selenium"
        ]):
            return "Failed", f"Browser error: {e}"

        # Generic failure
        return "Failed", str(e)


# ─── Main ────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="WhatsApp Bulk Message Sender")
    parser.add_argument("--dry-run", action="store_true", help="Test without actually sending messages")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Path to input Excel file (default: contacts.xlsx)")
    args = parser.parse_args()

    print("=" * 60)
    print("   📱 WhatsApp Bulk Message Sender")
    print("=" * 60)

    if args.dry_run:
        print("   🧪 DRY RUN MODE - No messages will be sent")
    print()

    # 1. Load message
    message = load_message()
    print(f"📝 Message: \"{message[:80]}{'...' if len(message) > 80 else ''}\"")
    print()

    # 2. Load contacts
    contacts = load_contacts(args.input)
    print(f"📊 Loaded {len(contacts)} contacts from: {os.path.basename(args.input)}")
    print()

    if not contacts:
        print("⚠️  No contacts found in the file.")
        sys.exit(0)

    # 3. Initialize database
    init_db()
    batch_id = str(uuid.uuid4())[:8]
    print(f"🆔 Batch ID: {batch_id}")
    print("-" * 60)

    # 4. Process each contact
    results = {"Sent": 0, "Failed": 0, "No WhatsApp": 0, "Dry Run": 0}

    for i, (name, phone) in enumerate(contacts, 1):
        print(f"\n[{i}/{len(contacts)}] 📞 {name} ({phone}) ... ", end="", flush=True)

        # Validate phone number
        is_valid, normalized_phone, error_reason = validate_phone(phone)

        if not is_valid:
            record_id = insert_record(name, phone, batch_id)
            update_record(record_id, "Failed", error_reason)
            results["Failed"] += 1
            print(f"❌ Invalid: {error_reason}")
            continue

        # Insert as pending
        record_id = insert_record(name, normalized_phone, batch_id)

        # Send message
        try:
            status, error = send_whatsapp_message(normalized_phone, message, dry_run=args.dry_run)
        except KeyboardInterrupt:
            update_record(record_id, "Failed", "Interrupted by user (Ctrl+C)")
            results["Failed"] += 1
            print("⛔ Interrupted!")
            print("\n⚠️  Sending stopped by user. Saving progress...")
            break

        update_record(record_id, status, error)
        results[status] = results.get(status, 0) + 1

        # Print status
        status_icons = {"Sent": "✅", "Failed": "❌", "No WhatsApp": "⚠️", "Dry Run": "🧪"}
        print(f"{status_icons.get(status, '❓')} {status}", end="")
        if error:
            print(f" - {error}", end="")
        print()

        # Delay between messages (skip on last message or dry run)
        if not args.dry_run and i < len(contacts):
            print(f"   ⏳ Waiting {DELAY_BETWEEN_MESSAGES}s before next message...", flush=True)
            try:
                time.sleep(DELAY_BETWEEN_MESSAGES)
            except KeyboardInterrupt:
                print("\n⚠️  Sending stopped by user. Saving progress...")
                break

    # 5. Generate status report Excel
    print("\n" + "=" * 60)
    records = get_all_records(batch_id)
    report_path = create_status_excel(batch_id, records, message)
    print(f"\n📄 Status report saved: {os.path.basename(report_path)}")

    # 6. Print summary
    db_summary = get_summary(batch_id)
    print("\n📊 Summary:")
    print(f"   Total:        {len(contacts)}")
    for status, count in db_summary.items():
        icon = {"Sent": "✅", "Failed": "❌", "No WhatsApp": "⚠️", "Dry Run": "🧪", "Pending": "⏳"}.get(status, "❓")
        print(f"   {icon} {status:14s} {count}")

    print(f"\n💾 Database: whatsapp_sender.db (Batch: {batch_id})")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⛔ Program terminated by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        print("   Please check your setup and try again.")
        sys.exit(1)
