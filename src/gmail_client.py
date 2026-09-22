import base64
from datetime import datetime, timezone, timedelta
import os
from pathlib import Path
import sys
from dotenv import load_dotenv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

try:
    from zoneinfo import ZoneInfo
    IST = ZoneInfo("Asia/Kolkata")
except Exception:
    IST = timezone(timedelta(hours=5, minutes=30), name="IST")

# Ensure UTF-8 output encoding for terminal on Windows
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Gmail read-only access scope
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(dotenv_path=BASE_DIR / ".env", override=True)
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_MAILBOX1 = BASE_DIR / "data" / "token_mailbox1.json"
TOKEN_MAILBOX2 = BASE_DIR / "data" / "token_mailbox2.json"


def extract_message_body(payload):
    """Extract plain text or HTML body from a Gmail message payload."""
    if not payload:
        return ""

    # Check multipart messages
    parts = payload.get("parts", [])
    if parts:
        # First preference: text/plain
        for part in parts:
            if part.get("mimeType") == "text/plain":
                data = part.get("body", {}).get("data")
                if data:
                    return base64.urlsafe_b64decode(data.encode("utf-8")).decode("utf-8", errors="replace")
            if "parts" in part:
                nested = extract_message_body(part)
                if nested:
                    return nested

        # Fallback preference: text/html
        for part in parts:
            if part.get("mimeType") == "text/html":
                data = part.get("body", {}).get("data")
                if data:
                    return base64.urlsafe_b64decode(data.encode("utf-8")).decode("utf-8", errors="replace")

    # Check single-part message
    body_data = payload.get("body", {}).get("data")
    if body_data:
        return base64.urlsafe_b64decode(body_data.encode("utf-8")).decode("utf-8", errors="replace")

    return ""


def authenticate_gmail(
    credentials_path=CREDENTIALS_FILE,
    token_path=TOKEN_MAILBOX1,
    account_label="Mailbox 1",
    login_hint=None,
    interactive=True
):
    """
    Authenticate with Gmail API using OAuth 2.0 desktop flow.
    Loads existing token if available and valid.
    If token is expired, attempts refresh.
    If no token exists and interactive=True, initiates browser login.
    Never prints credentials or token contents.
    """
    creds = None
    token_path = Path(token_path)
    credentials_path = Path(credentials_path)

    if token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        except Exception as e:
            print(f"[{account_label}] Failed to load cached token: {type(e).__name__}. Re-authenticating...")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                print(f"[{account_label}] Refreshing expired access token...")
                creds.refresh(Request())
            except Exception as e:
                print(f"[{account_label}] Token refresh failed: {type(e).__name__}.")
                creds = None

        if not creds:
            if not interactive:
                print(f"[{account_label}] Token not found ({token_path.name}). Run authentication test to connect.")
                return None

            if not credentials_path.exists():
                raise FileNotFoundError(
                    f"Credentials file not found at: {credentials_path.resolve()}"
                )

            print(f"\n[{account_label}] Initiating Google OAuth browser authentication...")
            if login_hint:
                print(f"[{account_label}] Please sign in using: {login_hint}")
            flow = InstalledAppFlow.from_client_secrets_file(
                str(credentials_path),
                scopes=SCOPES
            )
            creds = flow.run_local_server(port=0, prompt="consent")

        # Save credentials securely into the dedicated mailbox token file
        if creds and creds.valid:
            token_path.parent.mkdir(parents=True, exist_ok=True)
            with open(token_path, "w", encoding="utf-8") as token_file:
                token_file.write(creds.to_json())
            print(f"[{account_label}] Authentication token saved securely to: {token_path.relative_to(BASE_DIR)}")

    return creds


def get_gmail_service(mailbox_id=1, credentials_path=CREDENTIALS_FILE, interactive=True):
    """Build and return an authorized Gmail API service client for Mailbox 1 or Mailbox 2."""
    if str(mailbox_id).strip() in ("2", "mailbox2", "MAILBOX2"):
        token_path = TOKEN_MAILBOX2
        account_label = "Mailbox 2"
        login_hint = os.getenv("MAILBOX2", None)
    else:
        token_path = TOKEN_MAILBOX1
        account_label = "Mailbox 1"
        login_hint = os.getenv("MAILBOX1", None)

    creds = authenticate_gmail(
        credentials_path=credentials_path,
        token_path=token_path,
        account_label=account_label,
        login_hint=login_hint,
        interactive=interactive
    )
    if not creds:
        return None

    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    return service


def list_recent_messages(service, max_results=5, user_id="me"):
    """Fetch a list of recent messages from the user's mailbox."""
    response = service.users().messages().list(
        userId=user_id,
        maxResults=max_results
    ).execute()
    return response.get("messages", [])


def to_epoch_seconds(ts, default_tz=IST):
    """
    Convert a datetime, timestamp number, or ISO-8601 string to integer epoch seconds.
    Assumes Indian Standard Time (IST) if datetime is naive.
    """
    if ts is None:
        return None
    if isinstance(ts, (int, float)):
        return int(ts)
    if isinstance(ts, str):
        ts = datetime.fromisoformat(ts)
    if isinstance(ts, datetime):
        if ts.tzinfo is None:
            ts = ts.replace(tzinfo=default_tz)
        return int(ts.timestamp())
    raise TypeError(f"Unsupported timestamp format: {type(ts).__name__}")


def list_messages_in_time_range(
    service,
    start_time,
    end_time=None,
    user_id="me",
    overlap_buffer_seconds=0
):
    """
    Retrieve all message summaries from Gmail API within a given time range.
    Uses Gmail search query syntax ('after:<epoch> [before:<epoch>]') with full pagination.

    Args:
        service: Authorized Gmail API service resource.
        start_time: Starting timestamp (datetime, ISO string, or epoch seconds).
        end_time: Optional ending timestamp. If None, queries through present.
        user_id: Gmail user ID, default 'me'.
        overlap_buffer_seconds: Seconds subtracted from start_time to prevent boundary misses.

    Returns:
        List of message resource dicts: [{'id': '...', 'threadId': '...'}, ...]
    """
    if not service:
        raise ValueError("Gmail service client is None or not authenticated.")

    start_epoch = to_epoch_seconds(start_time)
    if overlap_buffer_seconds > 0:
        start_epoch = max(0, start_epoch - int(overlap_buffer_seconds))

    q_parts = [f"after:{start_epoch}"]

    if end_time is not None:
        end_epoch = to_epoch_seconds(end_time)
        q_parts.append(f"before:{end_epoch + 1}")

    query = " ".join(q_parts)

    messages = []
    seen_ids = set()
    page_token = None

    while True:
        try:
            response = service.users().messages().list(
                userId=user_id,
                q=query,
                pageToken=page_token,
                maxResults=100
            ).execute()
        except Exception as e:
            print(f"[Gmail Client Error] Query '{query}' failed: {type(e).__name__} - {e}")
            raise

        page_messages = response.get("messages", [])
        for msg in page_messages:
            mid = msg.get("id")
            if mid and mid not in seen_ids:
                seen_ids.add(mid)
                messages.append(msg)

        page_token = response.get("nextPageToken")
        if not page_token:
            break

    return messages


def extract_attachments_metadata(payload):
    """Recursively collect attachment metadata (filename, attachment_id, size) from message payload."""
    attachments = []

    def _walk(parts):
        for part in parts:
            filename = part.get("filename")
            body = part.get("body", {})
            att_id = body.get("attachmentId")
            if filename and att_id:
                attachments.append({
                    "filename": filename,
                    "attachment_id": att_id,
                    "size": body.get("size", 0)
                })
            if "parts" in part:
                _walk(part["parts"])

    if "parts" in payload:
        _walk(payload["parts"])
    elif payload.get("filename") and payload.get("body", {}).get("attachmentId"):
        attachments.append({
            "filename": payload["filename"],
            "attachment_id": payload["body"]["attachmentId"],
            "size": payload["body"].get("size", 0)
        })

    return attachments


def download_attachment(service, message_id, attachment_id, user_id="me"):
    """Fetch raw binary content for an attachment from Gmail API without modifying Gmail state."""
    attachment = service.users().messages().attachments().get(
        userId=user_id,
        messageId=message_id,
        id=attachment_id
    ).execute()
    data = attachment.get("data", "")
    return base64.urlsafe_b64decode(data.encode("utf-8"))


def get_message_details(service, message_id, user_id="me"):
    """Fetch email details (ID, Sender, Subject, Date, Snippet, Body, Attachments) without modifying it."""
    message = service.users().messages().get(
        userId=user_id,
        id=message_id,
        format="full"
    ).execute()

    payload = message.get("payload", {})
    headers = {h["name"]: h["value"] for h in payload.get("headers", [])}
    body = extract_message_body(payload)
    attachments = extract_attachments_metadata(payload)

    return {
        "id": message_id,
        "subject": headers.get("Subject", "(No Subject)"),
        "sender": headers.get("From", "(Unknown Sender)"),
        "date": headers.get("Date", "(Unknown Date)"),
        "snippet": message.get("snippet", ""),
        "body": body,
        "attachments": attachments
    }


def test_mailbox(mailbox_id=1):
    """Test OAuth authentication and display recent messages for a specific mailbox."""
    label = f"Mailbox {mailbox_id}"
    print(f"=== Gmail OAuth Test ({label}) ===")
    service = get_gmail_service(mailbox_id=mailbox_id, interactive=True)
    if not service:
        print(f"[{label}] Could not obtain service client.")
        return False

    # Fetch and display authenticated account email
    profile = service.users().getProfile(userId="me").execute()
    email_address = profile.get("emailAddress", "Unknown")
    print(f"[{label}] Authenticated Gmail address: {email_address}")

    # Fetch and display recent messages count and details
    messages = list_recent_messages(service, max_results=5)
    print(f"[{label}] Number of recent messages found: {len(messages)}")

    if messages:
        print(f"\n[{label}] Recent Messages (up to 5):")
        for idx, msg_item in enumerate(messages, start=1):
            details = get_message_details(service, msg_item["id"])
            print(f"  {idx}. Sender: {details['sender']}")
            print(f"     Subject: {details['subject']}")
    else:
        print(f"[{label}] No messages found in mailbox.")

    return True


def test_mailbox_1():
    """Test function for Mailbox 1."""
    return test_mailbox(1)


def test_mailbox_2():
    """Test function for Mailbox 2."""
    return test_mailbox(2)


if __name__ == "__main__":
    # Default to Mailbox 1 unless specified via CLI arguments (e.g. '2' or 'mailbox2')
    selected_mailbox = 1
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower().strip("-")
        if arg in ("2", "mailbox2", "mb2"):
            selected_mailbox = 2
        elif arg in ("all", "both"):
            test_mailbox(1)
            print()
            test_mailbox(2)
            sys.exit(0)

    test_mailbox(selected_mailbox)
