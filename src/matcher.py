import argparse
import os
from pathlib import Path
import re
import sys
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# Ensure UTF-8 output encoding for terminal on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(dotenv_path=BASE_DIR / ".env", override=True)
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from src.telegram_bot import send_telegram_message

TARGET_PHRASE = "you are eligible for this company"


def get_clean_email_text(subject, body):
    """Combine subject and stripped body text into clean plain text."""
    subject_text = subject or ""
    body_text = ""
    if body:
        try:
            soup = BeautifulSoup(body, "html.parser")
            body_text = soup.get_text(separator=" ")
        except Exception:
            body_text = body

    return f"{subject_text}\n{body_text}"


def check_eligibility_phrase(text, target_phrase=TARGET_PHRASE):
    """
    Check if the target phrase exists case-insensitively in the text.
    Returns the target phrase if found, otherwise None.
    """
    if not text or not target_phrase:
        return None

    # Normalize whitespace for robust matching across newlines and multiple spaces
    normalized_text = re.sub(r"\s+", " ", text).lower()
    normalized_phrase = re.sub(r"\s+", " ", target_phrase).lower()

    if normalized_phrase in normalized_text:
        return target_phrase

    return None


def extract_company_name(subject, body=""):
    """
    Extract the company name from the email subject or body text.
    Prefers patterns:
      1. 'eligible for this company: <Company Name>'
      2. 'Congratulations! You're Eligible for <Company Name>'
      3. 'Eligible for <Company Name>'
      4. 'Company: <Company Name>'
    Falls back to 'Company name could not be determined'.
    """
    clean_body = ""
    if body:
        try:
            soup = BeautifulSoup(body, "html.parser")
            clean_body = soup.get_text(separator=" ")
        except Exception:
            clean_body = body

    search_texts = [subject or "", clean_body]

    patterns = [
        # Pattern 1: eligible for this company: <Company Name>
        r"eligible\s+for\s+this\s+company\s*[:\-]\s*([A-Za-z0-9\s&.,'\-]+?)(?:\s+(?:placement|drive|recruitment|process|batch)|[.!\n\r]|$)",

        # Pattern 2: Congratulations! You're Eligible for <Company Name>
        r"congratulations!?\s*(?:based\s+on\s+your\s+profile,?)?\s*(?:you(?:'re|\s+are)\s+)?eligible\s+for\s+([A-Za-z0-9\s&.,'\-]+?)(?:\s+(?:placement\s+drive|placement|drive|recruitment|process|batch)|[.!\n\r]|$)",

        # Pattern 3: Eligible for <Company Name>
        r"\beligible\s+for\s+([A-Za-z0-9\s&.,'\-]+?)(?:\s+(?:placement\s+drive|placement|drive|recruitment|process|batch)|[.!\n\r]|$)",

        # Pattern 4: Company: <Company Name>
        r"\bcompany\s*[:\-]\s*([A-Za-z0-9\s&.,'\-]+?)(?:\s+(?:placement\s+drive|placement|drive|recruitment|process|batch)|[.!\n\r]|$)"
    ]

    for text in search_texts:
        if not text:
            continue
        for pat in patterns:
            match = re.search(pat, text, re.IGNORECASE)
            if match:
                candidate = match.group(1).strip().strip(":.,-")
                candidate = re.sub(r"\s+", " ", candidate).strip()
                if candidate and len(candidate) > 1 and candidate.lower() not in ("this company", "the", "a"):
                    return candidate

    return "Company name could not be determined"


def format_eligibility_alert(
    company_name,
    source_mailbox,
    sender,
    subject,
    date,
    matched_phrase=TARGET_PHRASE
):
    """Format the Telegram notification message for an eligibility detection."""
    lines = [
        "🚨 Placement Eligibility Alert! 🚨",
        "",
        f"🏢 Company: {company_name}",
        f"📬 Source Mailbox: {source_mailbox}",
        f"👤 Sender: {sender}",
        f"📌 Subject: {subject}",
        f"📅 Date: {date}",
        f"🔍 Matched Phrase: \"{matched_phrase}\"",
    ]
    return "\n".join(lines)


def process_mailbox1_email(
    msg_details,
    source_mailbox=None,
    target_phrase=TARGET_PHRASE,
    send_alert=True
):
    """
    Inspect an email from Mailbox 1 for eligibility phrases and send a Telegram alert if matched.
    Returns a dict with detection results.
    """
    source_mailbox = source_mailbox or os.getenv("MAILBOX1", "Mailbox 1")
    subject = msg_details.get("subject", "")
    body = msg_details.get("body", "")
    sender = msg_details.get("sender", "(Unknown Sender)")
    date = msg_details.get("date", "(Unknown Date)")

    clean_text = get_clean_email_text(subject, body)
    matched = check_eligibility_phrase(clean_text, target_phrase=target_phrase)

    if not matched:
        return {
            "matched": False,
            "company": None,
            "alert_sent": False
        }

    company_name = extract_company_name(subject, body)
    alert_message = format_eligibility_alert(
        company_name=company_name,
        source_mailbox=source_mailbox,
        sender=sender,
        subject=subject,
        date=date,
        matched_phrase=matched
    )

    if send_alert:
        print(f"[Mailbox 1] 🎯 Match found! Company: '{company_name}'. Sending Telegram alert...")
        success = send_telegram_message(alert_message)
    else:
        print(f"[Mailbox 1] 🎯 Match found! Company: '{company_name}'. (Dry-run mode: Alert not sent)")
        success = False

    return {
        "matched": True,
        "company": company_name,
        "alert_sent": success,
        "alert_message": alert_message
    }


def run_unit_tests():
    """Run verification checks on phrase detection and company extraction logic."""
    print("=== Running Matcher Verification Checks ===")

    # Test 1: Exact target phrase matching
    test_text_1 = "Hello, you are eligible for this company: Infosys Limited. Please confirm."
    assert check_eligibility_phrase(test_text_1, TARGET_PHRASE) is not None, "Failed on exact phrase"
    comp_1 = extract_company_name("", test_text_1)
    print(f"Test 1 (Target Phrase in body): Match='{TARGET_PHRASE}', Company='{comp_1}'")

    # Test 2: Subject pattern
    test_subj_2 = "Congratulations! You're Eligible for PharmEasy Pvt Ltd Placement Drive"
    comp_2 = extract_company_name(test_subj_2, "")
    print(f"Test 2 (Subject extraction): Company='{comp_2}'")

    # Test 3: Mixed case and newlines
    test_text_3 = "Dear Candidate,\nYOU ARE ELIGIBLE FOR THIS COMPANY\nDrive Number: 123"
    assert check_eligibility_phrase(test_text_3, TARGET_PHRASE) is not None, "Failed on uppercase phrase"
    print("Test 3 (Case-insensitivity): Succeeded")

    # Test 4: Unmatched email
    test_text_4 = "Your meeting has been rescheduled to Friday."
    assert check_eligibility_phrase(test_text_4, TARGET_PHRASE) is None, "Should not match unrelated text"
    comp_4 = extract_company_name(test_text_4, "")
    assert comp_4 == "Company name could not be determined", "Should fallback to undetermined"
    print(f"Test 4 (No match fallback): Company='{comp_4}'")

    print("\nAll logic verification checks passed!\n")


def test_telegram_alert_simulation():
    """Simulate a detection event and deliver a live test notification to Telegram."""
    sample_email = {
        "id": "test-sample-1",
        "sender": "VIT - CDC Office <noreply.cdcinfo@vitstudent.ac.in>",
        "subject": "Congratulations! You're Eligible for PharmEasy Pvt Ltd Placement Drive",
        "date": "Thu, 17 Sep 2026 11:33:32 +0000",
        "body": "Dear Candidate,\n\nCongratulations! Based on your profile, you are eligible for this company: PharmEasy Pvt Ltd.\n\nPlease confirm your participation."
    }

    print("=== Testing Live Eligibility Alert to Telegram ===")
    source_mb = os.getenv("MAILBOX1", "Mailbox 1")
    result = process_mailbox1_email(sample_email, source_mailbox=source_mb, target_phrase=TARGET_PHRASE)
    if result["alert_sent"]:
        print("\n✅ Verification successful! Telegram notification was delivered.")
    else:
        print("\n❌ Notification could not be delivered.")
    return result["alert_sent"]


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PlacementMonitor Eligibility Matcher")
    parser.add_argument("--test-alert", action="store_true", help="Send a simulated live alert to Telegram")
    parser.add_argument("--verify", action="store_true", help="Run internal logic checks")
    args = parser.parse_args()

    if args.test_alert:
        test_telegram_alert_simulation()
    else:
        run_unit_tests()
        if not args.verify:
            print("Tip: Run with '--test-alert' to test delivering a simulated alert to Telegram:")
            print("     .\\.venv\\Scripts\\python.exe src/matcher.py --test-alert")
