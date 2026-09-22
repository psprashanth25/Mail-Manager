import argparse
from datetime import datetime, timezone, timedelta
import json
import os
from pathlib import Path
import shutil
import sys
import time
from dotenv import load_dotenv

# Ensure UTF-8 output encoding for terminal on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

try:
    from zoneinfo import ZoneInfo
    IST = ZoneInfo("Asia/Kolkata")
except Exception:
    IST = timezone(timedelta(hours=5, minutes=30), name="IST")

STATE_FILE = BASE_DIR / "data" / "processed_messages.json"
CHECKPOINT_FILE = BASE_DIR / "data" / "monitor_state.json"

# Import Gmail client helpers
from src.gmail_client import (
    get_gmail_service,
    list_recent_messages,
    list_messages_in_time_range,
    get_message_details,
    to_epoch_seconds,
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


def format_ist_time(dt):
    """Format a datetime as 'YYYY-MM-DD HH:MM:SS IST'."""
    if dt is None:
        return "None"
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=IST)
    else:
        dt = dt.astimezone(IST)
    return dt.strftime("%Y-%m-%d %H:%M:%S IST")


def log_event(level, mailbox_label, message):
    """Print standard timestamped log line."""
    timestamp = datetime.now(IST).strftime("%Y-%m-%d %H:%M:%S")
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
    state_file = Path(state_file)
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
    state_file = Path(state_file)
    try:
        state_file.parent.mkdir(parents=True, exist_ok=True)
        temp_file = state_file.with_suffix(".tmp")
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(processed_data, f, indent=2)
        temp_file.replace(state_file)
    except Exception as e:
        log_error(None, f"Failed to save processed messages state: {type(e).__name__} - {e}")


def load_checkpoint(checkpoint_file=CHECKPOINT_FILE):
    """
    Load the last successful monitoring checkpoint timestamp from JSON.
    Returns timezone-aware datetime in IST, or None if no valid checkpoint exists.
    If the file is corrupted, preserves a backup copy and returns None to trigger a fresh 24-hour scan.
    """
    checkpoint_file = Path(checkpoint_file)
    if not checkpoint_file.exists():
        return None

    try:
        with open(checkpoint_file, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, dict):
            raise ValueError("Root JSON is not an object")

        raw_ts = data.get("last_successful_checkpoint")
        if not raw_ts or not isinstance(raw_ts, str):
            raise ValueError("Missing or invalid 'last_successful_checkpoint' field")

        dt = datetime.fromisoformat(raw_ts)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=IST)
        else:
            dt = dt.astimezone(IST)
        return dt

    except Exception as e:
        log_warn(None, f"Checkpoint file corrupted or unreadable at {checkpoint_file.name} ({type(e).__name__}: {e}).")
        try:
            backup_file = checkpoint_file.with_name(f"{checkpoint_file.stem}.corrupted.{int(time.time())}.json")
            shutil.copy2(checkpoint_file, backup_file)
            log_info(None, f"Preserved corrupted checkpoint file to: {backup_file.name}")
        except Exception as backup_err:
            log_warn(None, f"Could not preserve corrupted checkpoint backup: {backup_err}")

        log_info(None, "Initiating fresh initial 24-hour scan as safe recovery.")
        return None


def save_checkpoint(checkpoint_dt, checkpoint_file=CHECKPOINT_FILE):
    """
    Safely and atomically persist the monitoring checkpoint timestamp to JSON.
    Ensures timestamp is timezone-aware in IST and formatted as ISO-8601.
    """
    checkpoint_file = Path(checkpoint_file)
    try:
        if checkpoint_dt.tzinfo is None:
            checkpoint_dt = checkpoint_dt.replace(tzinfo=IST)
        else:
            checkpoint_dt = checkpoint_dt.astimezone(IST)

        checkpoint_file.parent.mkdir(parents=True, exist_ok=True)
        temp_file = checkpoint_file.with_suffix(".tmp")
        payload = {
            "last_successful_checkpoint": checkpoint_dt.isoformat()
        }
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        temp_file.replace(checkpoint_file)
        return True
    except Exception as e:
        log_error(None, f"Failed to save monitoring checkpoint: {type(e).__name__} - {e}")
        return False


def determine_scan_window(checkpoint_dt, current_time=None):
    """
    Determine the search window (start_time, end_time, mode) based on the checkpoint.
    - If checkpoint_dt is None (first-ever run or recovery): scan previous 24 hours.
    - If checkpoint_dt is valid: catch up from checkpoint to current_time.
    """
    if current_time is None:
        current_time = datetime.now(IST)
    elif current_time.tzinfo is None:
        current_time = current_time.replace(tzinfo=IST)
    else:
        current_time = current_time.astimezone(IST)

    if checkpoint_dt is None:
        start_time = current_time - timedelta(hours=24)
        end_time = current_time
        mode = "INITIAL 24-HOUR SCAN"
    else:
        if checkpoint_dt.tzinfo is None:
            checkpoint_dt = checkpoint_dt.replace(tzinfo=IST)
        else:
            checkpoint_dt = checkpoint_dt.astimezone(IST)

        if checkpoint_dt > current_time:
            log_warn(None, f"Saved checkpoint ({format_ist_time(checkpoint_dt)}) is in the future. Falling back to initial 24-hour scan.")
            start_time = current_time - timedelta(hours=24)
            end_time = current_time
            mode = "INITIAL 24-HOUR SCAN"
        else:
            start_time = checkpoint_dt
            end_time = current_time
            mode = "CATCH-UP"

    return start_time, end_time, mode


def get_monitoring_config():
    """Load monitoring settings from the root .env file."""
    load_dotenv(dotenv_path=BASE_DIR / ".env", override=True)
    raw_interval = os.getenv("CHECK_INTERVAL", "60").strip()
    try:
        interval = int(raw_interval)
    except ValueError:
        interval = 60

    mailbox1 = os.getenv("MAILBOX1", "Mailbox 1").strip()
    mailbox2 = os.getenv("MAILBOX2", "Mailbox 2").strip()

    return {
        "interval": interval,
        "mailbox1": mailbox1,
        "mailbox2": mailbox2,
    }


def check_mailbox(
    mailbox_id,
    email_address,
    processed_data,
    state_file=STATE_FILE,
    start_time=None,
    end_time=None,
    overlap_buffer_seconds=0
):
    """
    Check a single mailbox for new un-processed emails.
    Handles Gmail API errors, pagination, token refresh, and malformed emails gracefully.
    Returns a tuple of (new_messages: list, success: bool).
    """
    label = f"Mailbox {mailbox_id}"
    mailbox_key = f"mailbox{mailbox_id}"
    log_info(label, f"Checking for new messages ({email_address})...")

    try:
        service = get_gmail_service(mailbox_id=mailbox_id, interactive=False)
    except Exception as e:
        log_warn(label, f"Failed to initialize service: {type(e).__name__} - {e}. Will retry next cycle.")
        return [], False

    if not service:
        log_warn(label, f"Skipping check: account is not authenticated (run 'python src/gmail_client.py {mailbox_id}' to connect).")
        return [], False

    try:
        if start_time is not None:
            candidate_messages = list_messages_in_time_range(
                service,
                start_time=start_time,
                end_time=end_time,
                overlap_buffer_seconds=overlap_buffer_seconds
            )
        else:
            candidate_messages = list_recent_messages(service, max_results=5)
    except Exception as e:
        log_warn(label, f"Gmail API error fetching message list: {type(e).__name__} - {e}. Retrying on next cycle.")
        return [], False

    if not candidate_messages:
        log_info(label, "No messages found in mailbox.")
        return [], True

    processed_set = set(processed_data.get(mailbox_key, []))
    new_messages = []
    skipped_duplicates = 0

    # Reverse candidate messages so oldest messages are processed first in chronological order
    for msg_meta in reversed(candidate_messages):
        msg_id = msg_meta.get("id")
        if not msg_id:
            continue

        if msg_id in processed_set:
            skipped_duplicates += 1
            log_info(label, f"Duplicate message skipped: {msg_id}")
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
                if match_result.get("identifier_matched"):
                    matched_id = match_result.get("matched_identifier", "Registration number")
                    log_info(label, f"Identifier matched: {matched_id}")
                    if match_result.get("is_placement"):
                        log_info(label, "Placement relevance: TRUE")
                        log_info(label, "Classification: PLACEMENT")
                        log_info(label, f"Sending Telegram alert for '{match_result.get('company')}'")
                        if match_result.get("alert_sent"):
                            log_info(label, f"Telegram alert sent successfully for '{match_result.get('company')}'.")
                        else:
                            log_warn(label, f"Telegram alert delivery failed for '{match_result.get('company')}'.")
                    else:
                        log_info(label, "Placement relevance: FALSE")
                        log_info(label, "Classification: NON-PLACEMENT")
                        log_info(label, "Skipping Telegram placement alert")
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

    return new_messages, True


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


def run_monitor(max_cycles=None, interval=None, state_file=STATE_FILE, checkpoint_file=CHECKPOINT_FILE):
    """
    Run the Gmail monitoring engine with smart resume, catch-up monitoring, and 60-second polling.
    Isolated per-mailbox exception handling ensures one mailbox failure never blocks the other.
    If max_cycles is set, stops after that many cycles (ideal for testing).
    Handles keyboard interrupt (Ctrl+C) gracefully.
    """
    config = get_monitoring_config()
    poll_interval = interval if interval is not None else config["interval"]

    startup_time = datetime.now(IST)
    checkpoint = load_checkpoint(checkpoint_file=checkpoint_file)
    start_time, end_time, mode = determine_scan_window(checkpoint, startup_time)

    # 1. Startup Logging Banner
    print("=" * 60)
    print("PlacementMonitor - Gmail Placement Monitoring Engine")
    print("=" * 60)
    print()
    print("Current Time:")
    print(format_ist_time(startup_time))
    print()
    print("Previous Successful Checkpoint:")
    if checkpoint:
        print(format_ist_time(checkpoint))
    else:
        print("None (No previous monitoring checkpoint found. Performing initial 24-hour scan.)")
    print()
    print("Startup Mode:")
    print(mode)
    print()
    if mode == "INITIAL 24-HOUR SCAN":
        print("Lookback:")
    else:
        print("Catch-up Window:")
    print(format_ist_time(start_time))
    print("        →")
    print(format_ist_time(end_time))
    print()
    print("Polling Interval:")
    print(f"{poll_interval} seconds")
    print()
    print("=" * 60)

    processed_data = load_processed_messages(state_file=state_file)

    # 2. Perform Startup Initial / Catch-Up Scan
    log_info(None, f"Executing {mode} scan across mailboxes...")
    mb1_ok = False
    mb2_ok = False

    # Check Mailbox 1
    try:
        _, mb1_ok = check_mailbox(
            mailbox_id=1,
            email_address=config["mailbox1"],
            processed_data=processed_data,
            state_file=state_file,
            start_time=start_time,
            end_time=end_time,
            overlap_buffer_seconds=60
        )
    except Exception as e:
        log_error("Mailbox 1", f"Unexpected check failure: {type(e).__name__} - {e}. Continuing to Mailbox 2.")
        mb1_ok = False

    # Check Mailbox 2
    try:
        _, mb2_ok = check_mailbox(
            mailbox_id=2,
            email_address=config["mailbox2"],
            processed_data=processed_data,
            state_file=state_file,
            start_time=start_time,
            end_time=end_time,
            overlap_buffer_seconds=60
        )
    except Exception as e:
        log_error("Mailbox 2", f"Unexpected check failure: {type(e).__name__} - {e}. Continuing monitoring loop.")
        mb2_ok = False

    current_checkpoint = checkpoint
    if mb1_ok and mb2_ok:
        save_checkpoint(end_time, checkpoint_file=checkpoint_file)
        current_checkpoint = end_time
        print()
        print("=" * 60)
        print("Initial/catch-up scan completed successfully.")
        print("Monitoring checkpoint updated to:")
        print(format_ist_time(end_time))
        print()
        print("Entering continuous monitoring mode.")
        print(f"Polling interval: {poll_interval} seconds")
        print("=" * 60)
    else:
        log_warn(None, f"Startup scan encountered errors (Mailbox 1 ok: {mb1_ok}, Mailbox 2 ok: {mb2_ok}). Checkpoint retained at: {format_ist_time(checkpoint)}.")

    # 3. Continuous Monitoring Loop
    cycle = 0
    try:
        while True:
            if max_cycles and cycle >= max_cycles:
                log_info(None, f"Completed {max_cycles} cycle(s). Stopping monitoring engine.")
                break

            log_info(None, f"Waiting {poll_interval} seconds before next polling cycle... (Press Ctrl+C to stop)")
            time.sleep(poll_interval)

            cycle += 1
            cycle_time = datetime.now(IST)
            print(f"\n--- [Cycle {cycle}] Polling at {format_ist_time(cycle_time)} ---")

            cycle_start = current_checkpoint if current_checkpoint else (cycle_time - timedelta(seconds=poll_interval))
            cycle_end = cycle_time

            c_mb1_ok = False
            c_mb2_ok = False

            # Check Mailbox 1
            try:
                _, c_mb1_ok = check_mailbox(
                    mailbox_id=1,
                    email_address=config["mailbox1"],
                    processed_data=processed_data,
                    state_file=state_file,
                    start_time=cycle_start,
                    end_time=cycle_end,
                    overlap_buffer_seconds=60
                )
            except Exception as e:
                log_error("Mailbox 1", f"Unexpected check failure in cycle {cycle}: {type(e).__name__} - {e}.")
                c_mb1_ok = False

            # Check Mailbox 2
            try:
                _, c_mb2_ok = check_mailbox(
                    mailbox_id=2,
                    email_address=config["mailbox2"],
                    processed_data=processed_data,
                    state_file=state_file,
                    start_time=cycle_start,
                    end_time=cycle_end,
                    overlap_buffer_seconds=60
                )
            except Exception as e:
                log_error("Mailbox 2", f"Unexpected check failure in cycle {cycle}: {type(e).__name__} - {e}.")
                c_mb2_ok = False

            # Checkpoint only advances when required mailbox scans succeed
            if c_mb1_ok and c_mb2_ok:
                save_checkpoint(cycle_end, checkpoint_file=checkpoint_file)
                current_checkpoint = cycle_end
            else:
                log_warn(None, f"Cycle {cycle} scan encountered failures. Checkpoint not advanced, retained at {format_ist_time(current_checkpoint)}.")

    except KeyboardInterrupt:
        print("\n")
        log_info(None, "Monitoring stopped by user (Ctrl+C). Exiting cleanly.")


def main(argv=None):
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
        "--reset-checkpoint",
        action="store_true",
        help="Reset monitor_state.json checkpoint before starting"
    )
    parser.add_argument(
        "--reset-all",
        action="store_true",
        help="Reset both processed messages and checkpoint"
    )
    parser.add_argument(
        "--status",
        action="store_true",
        help="Show current monitoring checkpoint and processed message counts"
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
    args = parser.parse_args(argv)

    if args.status:
        config = get_monitoring_config()
        cp = load_checkpoint()
        data = load_processed_messages()
        print("=" * 60)
        print("PlacementMonitor Status")
        print("=" * 60)
        print(f"Current Time:                    {format_ist_time(datetime.now(IST))}")
        print(f"Configured Check Interval:       {config['interval']} seconds")
        print(f"Last Successful Checkpoint:      {format_ist_time(cp) if cp else 'None (First run pending)'}")
        print(f"Processed Mailbox 1 Messages:    {len(data.get('mailbox1', []))}")
        print(f"Processed Mailbox 2 Messages:    {len(data.get('mailbox2', []))}")
        print(f"Mailbox 1 Monitored Address:     {config['mailbox1']}")
        print(f"Mailbox 2 Monitored Address:     {config['mailbox2']}")
        print("=" * 60)
        return

    if args.recheck_mailbox1:
        recheck_mailbox1_emails(custom_phrase=args.phrase, send_alert=args.send_alert)
        return

    if args.recheck_mailbox2:
        diagnostic_scan_mailbox2(send_alert=args.send_alert)
        return

    if (args.reset_all or args.reset_state) and STATE_FILE.exists():
        STATE_FILE.unlink()
        log_info(None, "Cleared processed_messages.json state.")

    if (args.reset_all or args.reset_checkpoint) and CHECKPOINT_FILE.exists():
        CHECKPOINT_FILE.unlink()
        log_info(None, "Cleared monitor_state.json checkpoint.")

    run_monitor(max_cycles=args.cycles, interval=args.interval)


if __name__ == "__main__":
    main()
