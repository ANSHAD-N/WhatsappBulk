"""
WhatsApp Bulk Message Sender (Selenium-based)
==============================================
Sends a common message to all contacts from an Excel file via WhatsApp Web.

Usage:
  python whatsapp_sender.py                     # Send messages
  python whatsapp_sender.py --dry-run           # Test without sending
  python whatsapp_sender.py --input other.xlsx  # Different input file
  python whatsapp_sender.py --retries 3         # Retry failed messages

Requirements:
  pip install selenium webdriver-manager openpyxl
"""

import os
import sys
import time
import random
import argparse
import uuid

# ─── Dependency Checks ───────────────────────────────────────

MISSING_DEPS = []
for pkg in ["openpyxl", "selenium", "webdriver_manager"]:
    try:
        __import__(pkg)
    except ImportError:
        MISSING_DEPS.append(pkg.replace("_", "-"))

if MISSING_DEPS:
    print(f"Missing: {', '.join(MISSING_DEPS)}")
    print(f"Run: pip install {' '.join(MISSING_DEPS)}")
    sys.exit(1)

# ─── Imports from modules ───────────────────────────────────

from config import (
    DEFAULT_INPUT, MAX_RETRIES, DELAY_BETWEEN_MESSAGES,
    STATUS_SENT, STATUS_FAILED, STATUS_NO_WA, STATUS_DRY_RUN, STATUS_ICONS,
)
from utils import log, fatal, check_internet
from contacts import load_message, validate_phone, load_contacts
from driver import WhatsAppDriver
from report import create_status_excel
from database import init_db, insert_record, update_record, get_all_records, get_summary


# ─── Message Processing ─────────────────────────────────────

def process_contact(wa_driver, name, phone, message, batch_id, dry_run, retries):
    """Process a single contact. Returns (status, error)."""
    is_valid, normalized, error = validate_phone(phone)
    if not is_valid:
        rid = insert_record(name, phone, batch_id)
        update_record(rid, STATUS_FAILED, error)
        return STATUS_FAILED, error

    rid = insert_record(name, normalized, batch_id)

    if dry_run:
        update_record(rid, STATUS_DRY_RUN, None)
        return STATUS_DRY_RUN, None

    last_error = None
    for attempt in range(1, retries + 1):
        try:
            wa_driver.restart_if_needed()
            status, error = wa_driver.send_message(normalized, message)

            if status in (STATUS_SENT, STATUS_NO_WA):
                update_record(rid, status, error)
                return status, error

            last_error = error
            if attempt < retries:
                log(f"\n   Attempt {attempt} failed: {error}. Retrying in 10s...")
                time.sleep(10)
        except KeyboardInterrupt:
            update_record(rid, STATUS_FAILED, "Interrupted by user")
            raise
        except Exception as e:
            last_error = str(e)
            if attempt < retries:
                log(f"\n   Attempt {attempt} error: {last_error}. Retrying in 10s...")
                time.sleep(10)

    update_record(rid, STATUS_FAILED, f"{last_error} (after {retries} attempts)")
    return STATUS_FAILED, last_error


# ─── Main ────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="WhatsApp Bulk Message Sender")
    parser.add_argument("--dry-run", action="store_true", help="Test without sending")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Input Excel file")
    parser.add_argument("--retries", type=int, default=MAX_RETRIES, help="Retry attempts")
    args = parser.parse_args()

    log("=" * 60)
    log("   WhatsApp Bulk Message Sender (Selenium)")
    log("=" * 60)
    if args.dry_run:
        log("   MODE: DRY RUN")
    log("")

    if not args.dry_run and not check_internet():
        fatal("No internet connection.")

    message = load_message()
    log(f"Message: \"{message[:80]}{'...' if len(message) > 80 else ''}\"")
    contacts = load_contacts(args.input)
    log(f"Loaded {len(contacts)} contacts from: {os.path.basename(args.input)}")

    if not contacts:
        log("No contacts found.")
        sys.exit(0)

    init_db()
    batch_id = str(uuid.uuid4())[:8]
    log(f"Batch: {batch_id}")
    log("-" * 60)

    wa_driver = WhatsAppDriver()
    if not args.dry_run:
        wa_driver.start()

    interrupted = False
    for i, (name, phone) in enumerate(contacts, 1):
        log(f"\n[{i}/{len(contacts)}] {name} ({phone}) ... ", end="")

        try:
            status, error = process_contact(
                wa_driver, name, phone, message, batch_id, args.dry_run, args.retries
            )
        except KeyboardInterrupt:
            log("INTERRUPTED!")
            interrupted = True
            break

        icon = STATUS_ICONS.get(status, "[?]")
        log(f"{icon} {status}", end="")
        if error:
            log(f" - {error}", end="")
        log("")

        if not args.dry_run and i < len(contacts):
            delay = DELAY_BETWEEN_MESSAGES + random.uniform(-3, 5)
            log(f"   Waiting {delay:.0f}s before next...")
            try:
                time.sleep(delay)
            except KeyboardInterrupt:
                log("\nStopped by user.")
                interrupted = True
                break

    wa_driver.quit()

    log("\n" + "=" * 60)
    records = get_all_records(batch_id)
    try:
        report = create_status_excel(batch_id, records, message)
        log(f"Report: {os.path.basename(report)}")
    except Exception as e:
        log(f"Warning: Could not create report: {e}")

    summary = get_summary(batch_id)
    log("\nSummary:")
    log(f"  Total:          {len(contacts)}")
    for status, count in summary.items():
        log(f"  {STATUS_ICONS.get(status, '[?]')} {status:14s} {count}")

    log(f"\nDB: whatsapp_sender.db (Batch: {batch_id})")
    if interrupted:
        log("Note: Interrupted. Run again to continue.")
    log("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        log("\n\nTerminated by user.")
        sys.exit(0)
    except Exception as e:
        log(f"\nFatal error: {e}")
        sys.exit(1)
