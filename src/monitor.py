import argparse
import json
import os
from pathlib import Path
import sys
import time
from dotenv import load_dotenv

# Ensure UTF-8 output encoding for terminal on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

STATE_FILE = BASE_DIR / "data" / "processed_messages.json"

# Import Gmail client helpers
from src.gmail_client import (
    get_gmail_service,
    list_recent_messages,
    get_message_details,
)

# Import Eligibility Matcher helper
from src.matcher import (
    process_mailbox1_email,
    TARGET_PHRASE,
)

# Import Mailbox 2 Excel Matcher helper
from src.excel_matcher import (
    process_mailbox2_email,
    diagnostic_scan_mailbox2,
)


def log_event(level, mailbox_label, message):
    """Print standard timestamped log line."""
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    prefix = f"[{timestamp}] [{level.upper():5}]"
    if mailbox_label:
        print(f"{prefix} [{mailbox_label}] {message}")
    else:
        print(f"{prefix} {message}")


def log_info(mailbox_label, message):
    log_event("INFO", mailbox_label, message)


def log_warn(mailbox_label, message):
    log_event("WARN", mailbox_label, message)


def log_error(mailbox_label, message):
    log_event("ERROR", mailbox_label, message)


def load_processed_messages(state_file=STATE_FILE):
    """Load the dictionary of processed message IDs from the local JSON file."""
    if not state_file.exists():
        return {"mailbox1": [], "mailbox2": []}
    try:
        with open(state_file, "r", encoding="utf-8") as f:
            data = json.load(f)
            if not isinstance(data, dict):
                return {"mailbox1": [], "mailbox2": []}
            return {
                "mailbox1": list(data.get("mailbox1", [])),
                "mailbox2": list(data.get("mailbox2", [])),
            }
    except Exception as e:
        log_warn(None, f"Failed to read {state_file.name} ({type(e).__name__}). Using fresh state.")
        return {"mailbox1": [], "mailbox2": []}


def save_processed_messages(processed_data, state_file=STATE_FILE):
    """Save the processed message IDs safely and atomically to the local JSON file."""
    try:
        state_file.parent.mkdir(parents=True, exist_ok=True)
        temp_file = state_file.with_suffix(".tmp")
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(processed_data, f, indent=2)
        temp_file.replace(state_file)
    except Exception as e:
        log_error(None, f"Failed to save processed messages state: {type(e).__name__} - {e}")


def get_monitoring_config():
    """Load monitoring settings from the root .env file."""
    load_dotenv(dotenv_path=BASE_DIR / ".env", override=True)
    raw_interval = os.getenv("CHECK_INTERVAL", "10").strip()
    try:
        interval = int(raw_interval)
    except ValueError:
        interval = 10

    mailbox1 = os.getenv("MAILBOX1", "Mailbox 1").strip()
    mailbox2 = os.getenv("MAILBOX2", "Mailbox 2").strip()

    return {
        "interval": interval,
        "mailbox1": mailbox1,
        "mailbox2": mailbox2,
    }


def check_mailbox(mailbox_id, email_address, processed_data, state_file=STATE_FILE):
    """
    Check a single mailbox for new un-processed emails.
    Handles Gmail API errors, token refresh, and malformed emails gracefully.
    Returns a list of newly detected message details.
    """
    label = f"Mailbox {mailbox_id}"
    mailbox_key = f"mailbox{mailbox_id}"
    log_info(label, f"Checking for new messages ({email_address})...")

    try:
        service = get_gmail_service(mailbox_id=mailbox_id, interactive=False)
    except Exception as e:
        log_warn(label, f"Failed to initialize service: {type(e).__name__} - {e}. Will retry next cycle.")
        return []

    if not service:
        log_warn(label, f"Skipping check: account is not authenticated (run 'python src/gmail_client.py {mailbox_id}' to connect).")
        return []

    try:
        recent_messages = list_recent_messages(service, max_results=5)
    except Exception as e:
        log_warn(label, f"Gmail API error fetching message list: {type(e).__name__} - {e}. Retrying on next cycle.")
        return []

    if not recent_messages:
        log_info(label, "No messages found in mailbox.")
        return []

    processed_set = set(processed_data.get(mailbox_key, []))
    new_messages = []
    skipped_duplicates = 0

    for msg_meta in recent_messages:
        msg_id = msg_meta.get("id")
        if not msg_id:
            continue

        if msg_id in processed_set:
            skipped_duplicates += 1
            continue

        # Fetch full message details safely
        try:
            details = get_message_details(service, msg_id)
        except Exception as e:
            log_warn(label, f"Could not read message {msg_id}: {type(e).__name__} - {e}. Skipping.")
            continue

        new_messages.append(details)
        log_info(label, f"New email detected (ID: {msg_id}) | Subject: '{details['subject']}' | From: {details['sender']}")

        # Process eligibility & shortlists per mailbox
        try:
            if mailbox_id == 1:
                match_result = process_mailbox1_email(details, source_mailbox=email_address)
                if match_result.get("matched"):
                    log_info(label, f"Matching email detected! Company: '{match_result.get('company')}'.")
                    if match_result.get("alert_sent"):
                        log_info(label, f"Telegram alert sent successfully for '{match_result.get('company')}'.")
                    else:
                        log_warn(label, f"Telegram alert delivery failed for '{match_result.get('company')}'.")
                else:
                    log_info(label, f"Processed '{details['subject']}' - target phrase not found.")

            elif mailbox_id == 2:
                match_result = process_mailbox2_email(service, details, source_mailbox=email_address)
                if match_result.get("matched"):
                    log_info(label, f"Matching email detected! Company: '{match_result.get('company')}', Match: {match_result.get('matched_identifier')}.")
                    if match_result.get("alert_sent"):
                        log_info(label, f"Telegram alert sent successfully for '{match_result.get('company')}'.")
                    else:
                        log_warn(label, f"Telegram alert delivery failed for '{match_result.get('company')}'.")
                else:
                    log_info(label, f"Processed '{details['subject']}' - no matching identifier.")

        except Exception as e:
            log_error(label, f"Unexpected error processing email {msg_id}: {type(e).__name__} - {e}. Continuing.")

        # Record message ID in processed state to prevent duplicate notifications
        processed_data[mailbox_key].append(msg_id)
        processed_set.add(msg_id)

    # Persist updated state if any new messages were recorded
    if new_messages:
        save_processed_messages(processed_data, state_file=state_file)
        log_info(label, f"Saved {len(new_messages)} newly processed message(s) to state file.")
    else:
        if skipped_duplicates > 0:
            log_info(label, f"Duplicate skipped: {skipped_duplicates} message(s) already processed.")
        else:
            log_info(label, "No new messages detected.")

    return new_messages


def recheck_mailbox1_emails(max_emails=5, custom_phrase=TARGET_PHRASE, send_alert=False):
    """
    Diagnostic function to inspect existing Mailbox 1 emails against the eligibility phrase.
    Useful for testing phrase matching on historical emails.
    """
    config = get_monitoring_config()
    print("=" * 65)
    print(f"Diagnostic Recheck of Mailbox 1: {config['mailbox1']}")
    print(f"Target Phrase: \"{custom_phrase}\"")
    print(f"Send Live Alert: {send_alert}")
    print("=" * 65)

    service = get_gmail_service(mailbox_id=1, interactive=False)
    if not service:
        print("Mailbox 1 service could not be initialized.")
        return

    messages = list_recent_messages(service, max_results=max_emails)
    print(f"Inspecting {len(messages)} recent message(s)...\n")

    matched_count = 0
    for idx, msg_meta in enumerate(messages, start=1):
        details = get_message_details(service, msg_meta["id"])
        print(f"[{idx}] Subject: {details['subject']}")
        print(f"    From: {details['sender']}")

        result = process_mailbox1_email(
            details,
            source_mailbox=config["mailbox1"],
            target_phrase=custom_phrase,
            send_alert=send_alert
        )
        if result["matched"]:
            matched_count += 1
            print(f"    👉 MATCH DETECTED! Company: '{result['company']}'")
            if send_alert:
                print("    📨 Telegram alert was dispatched.")
            else:
                print("    (Alert sending skipped in dry-run mode)")
        else:
            print(f"    ❌ Target phrase not found.")
        print()

    print(f"Recheck complete. Matched {matched_count} out of {len(messages)} email(s).")


def run_monitor(max_cycles=None, interval=None):
    """
    Run the Gmail monitoring loop continuously for Mailbox 1 and Mailbox 2.
    Isolated per-mailbox exception handling ensures one mailbox failure never blocks the other.
    If max_cycles is set, stops after that many cycles (ideal for testing).
    Handles keyboard interrupt (Ctrl+C) gracefully.
    """
    config = get_monitoring_config()
    poll_interval = interval if interval is not None else config["interval"]

    print("=" * 65)
    print("PlacementMonitor - Continuous Gmail Polling Engine")
    print(f"Check Interval: {poll_interval}s")
    print(f"Mailbox 1: {config['mailbox1']} (Eligibility Phrase Detection)")
    print(f"Mailbox 2: {config['mailbox2']} (Identifier & Excel Parsing)")
    if max_cycles:
        print(f"Mode: Test ({max_cycles} cycle(s))")
    else:
        print("Mode: Continuous (Press Ctrl+C to stop)")
    print("=" * 65)

    processed_data = load_processed_messages()
    cycle = 0

    try:
        while True:
            cycle += 1
            print(f"\n--- [Cycle {cycle}] Polling at {time.strftime('%Y-%m-%d %H:%M:%S')} ---")

            # Check Mailbox 1 with isolated exception guard
            try:
                check_mailbox(
                    mailbox_id=1,
                    email_address=config["mailbox1"],
                    processed_data=processed_data
                )
            except Exception as e:
                log_error("Mailbox 1", f"Unexpected check failure: {type(e).__name__} - {e}. Continuing to Mailbox 2.")

            # Check Mailbox 2 with isolated exception guard
            try:
                check_mailbox(
                    mailbox_id=2,
                    email_address=config["mailbox2"],
                    processed_data=processed_data
                )
            except Exception as e:
                log_error("Mailbox 2", f"Unexpected check failure: {type(e).__name__} - {e}. Continuing monitoring loop.")

            if max_cycles and cycle >= max_cycles:
                log_info(None, f"Completed {max_cycles} test cycle(s). Stopping monitoring engine.")
                break

            log_info(None, f"Waiting {poll_interval}s before next polling cycle... (Press Ctrl+C to stop)")
            time.sleep(poll_interval)

    except KeyboardInterrupt:
        print("\n")
        log_info(None, "Monitoring stopped by user (Ctrl+C). Exiting cleanly.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PlacementMonitor Gmail Polling Engine")
    parser.add_argument(
        "--cycles",
        type=int,
        default=None,
        help="Number of polling cycles to run before stopping (test mode)"
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=None,
        help="Override check interval in seconds"
    )
    parser.add_argument(
        "--reset-state",
        action="store_true",
        help="Reset processed_messages.json state before starting"
    )
    parser.add_argument(
        "--recheck-mailbox1",
        action="store_true",
        help="Diagnostic recheck of recent Mailbox 1 emails against eligibility phrase"
    )
    parser.add_argument(
        "--recheck-mailbox2",
        action="store_true",
        help="Diagnostic scan of recent Mailbox 2 emails and Excel attachments"
    )
    parser.add_argument(
        "--phrase",
        type=str,
        default=TARGET_PHRASE,
        help="Phrase to evaluate with --recheck-mailbox1"
    )
    parser.add_argument(
        "--send-alert",
        action="store_true",
        help="Send live Telegram alerts during diagnostic scans if match found"
    )
    args = parser.parse_args()

    if args.recheck_mailbox1:
        recheck_mailbox1_emails(custom_phrase=args.phrase, send_alert=args.send_alert)
        sys.exit(0)

    if args.recheck_mailbox2:
        diagnostic_scan_mailbox2(send_alert=args.send_alert)
        sys.exit(0)

    if args.reset_state and STATE_FILE.exists():
        STATE_FILE.unlink()
        log_info(None, "Cleared processed_messages.json state.")

    run_monitor(max_cycles=args.cycles, interval=args.interval)
