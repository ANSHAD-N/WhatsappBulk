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
from database import init_db, insert_record, update_record, get_all_records, get_summary, get_completed_phones


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

def parse_arguments():
    """Step 1: Parse command-line settings the user provides."""
    parser = argparse.ArgumentParser(description="WhatsApp Bulk Message Sender")
    parser.add_argument("--dry-run", action="store_true", help="Test without actually sending messages")
    parser.add_argument("--input", default=DEFAULT_INPUT, help="Path to your Input Excel file")
    parser.add_argument("--retries", type=int, default=MAX_RETRIES, help="Number of times to retry if a message fails")
    parser.add_argument("--batch-size", type=int, default=50, help="Maximum number of new messages to send per run (0 means unlimited)")
    return parser.parse_args()

def filter_pending_contacts(all_contacts, batch_size):
    """Step 2: Check the database to remove people we already messaged."""
    # Get a list of phone numbers we already finished (either "Sent" or "No WhatsApp")
    completed_phones = get_completed_phones()
    
    pending_contacts = []
    
    # Loop through everyone in our Excel file
    for name, phone in all_contacts:
        # Check if their phone number format is valid
        is_valid, normalized_phone, _ = validate_phone(phone)
        
        # Use the normalized phone if it's valid, otherwise check the original phone
        check_phone = normalized_phone if is_valid else phone
        
        # If we haven't completed this phone number yet, add it to our pending list
        if check_phone not in completed_phones:
            pending_contacts.append((name, phone))
            
    # Calculate how many were skipped
    already_processed_count = len(all_contacts) - len(pending_contacts)
    log(f"Contacts already processed inside DB: {already_processed_count}")
    
    # Apply our daily limit (batch size) if one was provided
    if batch_size > 0:
        contacts_to_process = pending_contacts[:batch_size]
    else:
        contacts_to_process = pending_contacts
        
    log(f"Contacts to process in this exact run: {len(contacts_to_process)} (Daily Limit: {batch_size})")
    return contacts_to_process

def generate_final_reports(batch_id, message, total_contacts_count, processed_contacts_count):
    """Step 3: Save results to an Excel Status Report and print the summary."""
    log("\n" + "=" * 60)
    
    # Get all the records from our SQLite database for this specific run
    records = get_all_records(batch_id)
    
    try:
        # Create an Excel file containing the results for this batch
        report = create_status_excel(batch_id, records, message)
        log(f"Report Generated: {os.path.basename(report)}")
    except Exception as e:
        log(f"Warning: Could not create the Excel report: {e}")

    # Print a summary of how many succeeded, failed, etc.
    summary = get_summary(batch_id)
    log("\nSummary:")
    log(f"  Total Contacts in Excel:   {total_contacts_count}")
    log(f"  Processed in this run:     {processed_contacts_count}")
    
    for status, count in summary.items():
        # Print each status with its corresponding icon
        icon = STATUS_ICONS.get(status, "[?]")
        log(f"  {icon} {status:14s} {count}")

def main():
    """Main function that orchestrates the entire script."""
    # ─── 1. Setup Phase ───
    args = parse_arguments()

    log("=" * 60)
    log("   WhatsApp Bulk Message Sender (Selenium)")
    log("=" * 60)
    
    if args.dry_run:
        log("   MODE: DRY RUN (Testing Mode - No actual messages will be sent)")
    log("")

    # Check for internet connection before starting
    if not args.dry_run and not check_internet():
        fatal("No internet connection detected. Please connect to the internet and try again.")

    # Load the common message from config.txt
    message = load_message()
    log(f"Message Preview: \"{message[:80]}{'...' if len(message) > 80 else ''}\"")
    
    # Load all contacts from the Excel file
    all_contacts = load_contacts(args.input)
    log(f"Loaded {len(all_contacts)} total contacts from: {os.path.basename(args.input)}")

    if not all_contacts:
        log("No contacts found in the Excel file. Exiting.")
        sys.exit(0)

    # ─── 2. Filtering Phase ───
    # Initialize the database so we can read from it
    init_db()
    
    # Get the exact list of contacts we need to message today
    contacts_to_process = filter_pending_contacts(all_contacts, args.batch_size)

    if not contacts_to_process:
        log("No new contacts to process today. All done!")
        sys.exit(0)

    # Generate a unique Batch ID for this specific run to keep track of it in the database
    batch_id = str(uuid.uuid4())[:8]
    log(f"\nStarted Batch ID: {batch_id}")
    log("-" * 60)

    # ─── 3. WhatsApp Connection Phase ───
    # Open the Chrome Browser for WhatsApp Web
    wa_driver = WhatsAppDriver()
    if not args.dry_run:
        wa_driver.start()

    interrupted = False
    
    # ─── 4. Message Sending Phase ───
    # Loop through each contact and send the message
    for i, (name, phone) in enumerate(contacts_to_process, 1):
        log(f"\n[{i}/{len(contacts_to_process)}] {name} ({phone}) ... ", end="")

        try:
            # Let the process_contact function handle sending and updating the DB
            status, error = process_contact(
                wa_driver, name, phone, message, batch_id, args.dry_run, args.retries
            )
        except KeyboardInterrupt:
            # If the user presses Ctrl+C, stop safely
            log("INTERRUPTED BY USER!")
            interrupted = True
            break

        # Log the result of the message (Sent, Failed, etc.)
        icon = STATUS_ICONS.get(status, "[?]")
        log(f"{icon} {status}", end="")
        if error:
            log(f" - {error}", end="")
        log("")

        # Wait a random amount of seconds before messaging the next person to avoid getting banned
        if not args.dry_run and i < len(contacts_to_process):
            delay = DELAY_BETWEEN_MESSAGES + random.uniform(-3, 5)
            log(f"   Waiting {delay:.0f} seconds before sending the next one...")
            try:
                time.sleep(delay)
            except KeyboardInterrupt:
                log("\nStopped by user during sleep delay.")
                interrupted = True
                break

    # Close the WhatsApp Browser once done
    wa_driver.quit()

    # ─── 5. Reporting Phase ───
    # Generate the final reports and summary
    generate_final_reports(batch_id, message, len(all_contacts), len(contacts_to_process))

    log(f"\nDatabase File: whatsapp_sender.db (Batch ID: {batch_id})")
    if interrupted:
        log("Note: The script was interrupted early. You can run it again to continue from where it left off.")
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
