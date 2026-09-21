import os
from pathlib import Path
import requests
from dotenv import load_dotenv

# Ensure environment variables are loaded from the root .env file
BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = BASE_DIR / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=True)


def clean_env_value(val):
    """Strip whitespace and surrounding single or double quotes."""
    if val is None:
        return ""
    val_str = str(val).strip()
    if (val_str.startswith('"') and val_str.endswith('"')) or (val_str.startswith("'") and val_str.endswith("'")):
        val_str = val_str[1:-1].strip()
    return val_str


def get_telegram_config():
    """Retrieve and validate Telegram configuration from environment variables."""
    load_dotenv(dotenv_path=ENV_PATH, override=True)

    raw_token = os.getenv("TELEGRAM_BOT_TOKEN")
    raw_chat_id = os.getenv("TELEGRAM_CHAT_ID")

    bot_token = clean_env_value(raw_token)
    chat_id = clean_env_value(raw_chat_id)

    return bot_token, chat_id


def check_bot_status(bot_token):
    """Perform diagnostic request to getMe and report bot token validity and username."""
    if not bot_token:
        print("[Diagnostic] Cannot verify bot: Token is missing.")
        return False, None

    get_me_url = f"https://api.telegram.org/bot{bot_token}/getMe"
    try:
        res = requests.get(get_me_url, timeout=10)
        data = res.json()
        if res.status_code == 200 and data.get("ok"):
            bot_username = data.get("result", {}).get("username", "Unknown")
            bot_first_name = data.get("result", {}).get("first_name", "")
            print(f"[Diagnostic] Bot Token Status: Valid (Bot: @{bot_username}, Name: {bot_first_name})")
            return True, bot_username
        else:
            err_desc = data.get("description", "Unknown error")
            print(f"[Diagnostic] Bot Token Status: Invalid (Error: {err_desc})")
            return False, None
    except requests.RequestException as e:
        print(f"[Diagnostic] Bot Token Verification Failed: Network error ({type(e).__name__})")
        return False, None


def send_telegram_message(message):
    """Send a notification message to the configured Telegram chat."""
    bot_token, chat_id = get_telegram_config()
    if not bot_token:
        print("[Error] TELEGRAM_BOT_TOKEN is missing.")
        return False
    if not chat_id:
        print("[Error] TELEGRAM_CHAT_ID is missing.")
        return False

    api_url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    payload = {
        "chat_id": str(chat_id),
        "text": message,
    }

    try:
        response = requests.post(api_url, json=payload, timeout=15)
        response_data = response.json()

        if response.status_code == 200 and response_data.get("ok"):
            print("Telegram notification delivered successfully.")
            return True
        else:
            description = response_data.get("description", "Unknown error")
            print(f"[Telegram API Error] Status {response.status_code}: {description}")
            return False
    except requests.RequestException as e:
        print(f"[Network Error] Failed to send Telegram message ({type(e).__name__}).")
        return False


def test_telegram_message():
    """Run diagnostic checks and attempt to send the test message."""
    print("=== Telegram Diagnostics ===")
    print(f"Loading .env from: {ENV_PATH.resolve()}")

    bot_token, chat_id = get_telegram_config()

    token_exists = bool(bot_token)
    token_len = len(bot_token) if bot_token else 0
    if token_len >= 10:
        masked_token = f"{bot_token[:5]}...{bot_token[-5:]}"
    elif token_len > 0:
        masked_token = f"{bot_token[:2]}...{bot_token[-2:]}"
    else:
        masked_token = "N/A"

    print(f"Token Exists: {token_exists}")
    print(f"Token Length: {token_len}")
    print(f"Token (First 5 / Last 5): {masked_token}")
    print(f"Chat ID Value: {chat_id}")
    print(f"Chat ID Type: {type(chat_id).__name__}")

    print("\n--- Checking Bot Token Validity (/getMe) ---")
    is_valid, bot_username = check_bot_status(bot_token)

    print("\n--- Sending Test Message ---")
    test_message = "✅ PlacementMonitor Telegram test successful!"
    success = send_telegram_message(test_message)
    if success:
        print("Test passed: Message was successfully delivered!")
    else:
        print("Test failed: Message could not be delivered.")
    return success


if __name__ == "__main__":
    test_telegram_message()
