# PSP's Placement Monitor

A Python automation that monitors Gmail inboxes for placement eligibility and shortlist notifications and sends real-time alerts through Telegram.

---

## 1. Project Overview

**PSP's Placement Monitor** (repository: `Mail-Manager`) is a lightweight, reliable automation tool designed to bridge the gap between busy college placement communication channels and students. It continuously monitors two distinct Gmail accounts via the official Google Gmail API, analyzes emails and spreadsheet attachments for eligibility phrases and candidate identifiers, and immediately forwards structured, actionable notification summaries directly to a personal Telegram chat.

---

## 2. Problem Statement

During campus recruitment seasons, students face critical challenges:
- Important announcements and shortlist updates are fragmented across personal emails and official university/college inboxes.
- Placement shortlist files are frequently distributed as Excel spreadsheets (`.xlsx` or `.xls`) containing hundreds of roll numbers, requiring manual downloading and scanning.
- Deadlines for confirmations, technical tests, and interviews are often extremely tight (sometimes only a few hours).
- Constantly refreshing multiple email accounts on a computer is distracting and prone to missed alerts.

PlacementMonitor solves this by acting as an always-on watcher that parses inboxes and deep-scans spreadsheet attachments, sending instant push notifications to mobile devices via Telegram.

---

## 3. What the System Does

1. **Monitors Two Separate Gmail Inboxes:**
   - **Mailbox 1 (Personal):** Watches for company eligibility announcements.
   - **Mailbox 2 (College/University):** Watches for shortlist updates, interview calls, and schedule notices.
2. **Scans Email Content & Deep-Scans Attachments:** Inspects email subjects, plain text and HTML bodies, and iterates through every sheet and row in `.xlsx` and `.xls` attachments.
3. **Extracts Rich Context:** Automatically discovers company names, interview dates, reporting times, venues, instructions, registration links, and contact addresses.
4. **Sends Instant Telegram Push Notifications:** Dispatches structured, emoji-tagged alerts with essential details.
5. **Prevents Duplicate Alerts:** Tracks processed message IDs persistently to ensure each email is notified exactly once.

---

## 4. Key Features

- ⚡ **Dual Mailbox Monitoring:** Independent OAuth token management and isolated error-handling for personal and college Gmail accounts.
- 🔍 **Deep Excel Parsing:** Fully reads `.xlsx` (via `openpyxl`) and legacy `.xls` (via `xlrd`) attachments across all sheets and cells.
- 🎯 **Multi-Identifier Matching:** Detects student registration numbers, NeoPat/assessment IDs, or both simultaneously.
- 🏢 **Intelligent Company Extraction:** Regex-based detection patterns extract company names from subject headers, body text, or attachment names.
- 📋 **Automated Context Extraction:** Pulls test dates, interview venues, reporting times, required attire, laptop mandates, and contact information.
- 🛡️ **Duplicate Prevention:** Atomic state persistence prevents duplicate alerts even across script restarts.
- 🚀 **PowerShell Launcher:** Includes a turnkey `start_monitor.ps1` script for one-click startup on Windows.

---

## 5. System Workflow

1. **Cycle Start:** The engine begins a polling cycle based on `CHECK_INTERVAL`.
2. **Mailbox 1 Evaluation:** Queries Gmail API for recent messages. Checks whether each message ID exists in `data/processed_messages.json`. For new messages, extracts the body and evaluates for `"you are eligible for this company"` (case-insensitive). If matched, extracts the company name and sends an Eligibility Alert to Telegram.
3. **Mailbox 2 Evaluation:** Queries Mailbox 2 for recent messages. Skips previously processed IDs. Evaluates subject and body for student registration numbers and NeoPat IDs. Downloads any `.xlsx` or `.xls` attachments and scans all rows and columns across every sheet. If matched, extracts context and sends a Shortlist Alert to Telegram.
4. **State Persistence:** Adds newly processed message IDs to `data/processed_messages.json` using atomic file writing.
5. **Sleep & Repeat:** Waits for the configured interval before executing the next check cycle.

---

## 6. Architecture

```
                       ┌────────────────────────┐
                       │ Continuous Polling     │
                       │ Engine (monitor.py)    │
                       └───────────┬────────────┘
                                   │
         ┌─────────────────────────┴─────────────────────────┐
         ▼                                                   ▼
┌──────────────────┐                               ┌──────────────────┐
│ Mailbox 1 (Gmail)│                               │ Mailbox 2 (Gmail)│
│  Personal Inbox  │                               │  College Inbox   │
└────────┬─────────┘                               └────────┬─────────┘
         │ Gmail API (OAuth 2.0)                            │ Gmail API (OAuth 2.0)
         ▼                                                  ▼
┌──────────────────┐                               ┌──────────────────┐
│  Email Content   │                               │ Subject & Body   │
│  (Subject/Body)  │                               │    Matcher       │
└────────┬─────────┘                               └────────┬─────────┘
         │                                                  │
         ▼                                                  ├────────────────────────┐
┌──────────────────┐                                        ▼                        ▼
│Eligibility Match │                               ┌──────────────────┐     ┌─────────────────┐
│("you are eligible│                               │  Text Matcher    │     │  Excel Matcher  │
│for this company")│                               │(Reg No/NeoPat ID)│     │ (.xlsx & .xls)  │
└────────┬─────────┘                               └────────┬─────────┘     └────────┬────────┘
         │                                                  │                        │
         ▼                                                  └───────────┬────────────┘
┌──────────────────┐                                                    ▼
│ Company Extract  │                                           ┌──────────────────┐
└────────┬─────────┘                                           │Company & Context │
         │                                                     │    Extraction    │
         │                                                     └────────┬─────────┘
         │                                                              │
         └─────────────────────────┬────────────────────────────────────┘
                                   │
                                   ▼
                    ┌──────────────────────────────┐
                    │ Telegram Bot Dispatcher      │
                    │      (telegram_bot.py)       │
                    └──────────────┬───────────────┘
                                   │ Telegram Bot API (HTTPS POST)
                                   ▼
                    ┌──────────────────────────────┐
                    │     Mobile Telegram App      │
                    │   (Instant Push Alerts)      │
                    └──────────────────────────────┘

                    ┌──────────────────────────────┐
                    │ processed_messages.json      │
                    │ (Atomic Duplicate Protection)│
                    └──────────────────────────────┘
```

---

## 7. Supported Mailboxes

| Mailbox | Purpose | Target Account | Primary Match Criteria |
| :--- | :--- | :--- | :--- |
| **Mailbox 1** | Personal Placements | Personal Gmail Account | Phrase: `"you are eligible for this company"` |
| **Mailbox 2** | Campus Recruitment Drives | College / University Gmail | Student Registration Number & NeoPat ID across text and spreadsheets |

---

## 8. Detection Logic

### Mailbox 1 (Eligibility Detection)
- Evaluates cleaned text from subject and HTML body.
- Performs a normalized case-insensitive match for `"you are eligible for this company"`.
- Extracts company names using structured patterns (e.g., `eligible for this company: <Company>`, `Congratulations! You're Eligible for <Company>`).

### Mailbox 2 (Shortlist & Schedule Detection)
- Searches for student identifiers (`REGISTRATION_NUMBER` and `NEOPAT_ID`).
- Identifies whether one or both identifiers matched (`Registration number`, `NeoPat ID`, or `Both identifiers`).
- Extracts operational details:
  - **Company:** Extracted from subject headers, body text, or attachment filenames.
  - **Venue:** Identifies interview halls, block numbers, or campus locations.
  - **Test/Interview Date:** Parses `DD-MM-YYYY` or `DD/MM/YYYY` dates.
  - **Reporting Time:** Parses `HH:MM AM/PM` formats.
  - **Next Steps:** Flags technical interview rounds, online assessments, formal attire requirements, and laptop mandates.
  - **Links & Contacts:** Extracts placement portal links and coordinator email addresses.

---

## 9. Excel Attachment Processing

Spreadsheet processing handles both modern and legacy formats without requiring Microsoft Excel to be installed on the host system:

1. **Attachment Discovery:** Inspects email MIME payloads for attachments ending in `.xlsx` or `.xls`.
2. **Temporary Download:** Saves the binary file into `data/temp_attachments/` with sanitized filenames.
3. **Multi-Sheet Scanning:**
   - **`.xlsx` Files:** Scanned with `openpyxl` using `read_only=True` and `data_only=True` to minimize memory overhead.
   - **`.xls` Files:** Scanned with `xlrd`, iterating through all sheet indices.
4. **Cell-by-Cell Search:** Converts every cell value to a lowercase string and matches against candidate identifiers.
5. **Instant Deletion:** Safely removes the downloaded file immediately after inspection in a `finally` block.

---

## 10. Telegram Notifications

The bot sends clean, structured alerts with emoji markers for high readability on mobile devices.

### Placement Eligibility Alert
Real alert received when personal Mailbox 1 receives an eligibility notification:

![Placement Eligibility Alert](screenshots/telegram-eligibility-alert.png)

### Placement Shortlist Alert
Real alert received when Mailbox 2 detects student identifiers inside an attached Excel shortlist (`Exxonmobil interview shortlist.xlsx`), with extracted interview venue, date, time, and instructions:

![Placement Shortlist Alert](screenshots/telegram-shortlist-alert.png)

### Telegram Bot Output
Overview of the Telegram bot feed showing sequential placement notifications:

![Telegram Bot Output](screenshots/telegram-bot-overview.png)

---

## 11. Project Structure

```
Mail-Manager/
│
├── README.md                      # Comprehensive project documentation
├── requirements.txt               # Python package dependencies
├── .env.example                   # Environment variable template (no secrets)
├── .gitignore                     # Git exclusion rules for tokens, secrets, data
├── start_monitor.ps1              # Windows PowerShell one-click launcher
│
├── src/                           # Application source code
│   ├── __init__.py                # Package initializer
│   ├── main.py                    # Application entry point
│   ├── monitor.py                 # Continuous polling engine & state manager
│   ├── gmail_client.py            # Gmail API OAuth client & message fetcher
│   ├── matcher.py                 # Mailbox 1 eligibility phrase & company extractor
│   ├── excel_matcher.py           # Mailbox 2 identifier & Excel attachment parser
│   └── telegram_bot.py            # Telegram Bot API client & alert dispatcher
│
├── screenshots/                   # Real Telegram notification screenshots
│   ├── telegram-eligibility-alert.png
│   ├── telegram-shortlist-alert.png
│   └── telegram-bot-overview.png
│
├── docs/                          # Detailed technical documentation
│   └── project-overview.md        # In-depth system flow, logic & future roadmap
│
└── data/                          # Local runtime data (Git-ignored)
    ├── token_mailbox1.json        # OAuth token for Mailbox 1
    ├── token_mailbox2.json        # OAuth token for Mailbox 2
    ├── processed_messages.json    # Duplicate prevention state store
    └── temp_attachments/         # Ephemeral attachment download workspace
```

---

## 12. Technology Stack

- **Core Language:** Python 3.10+ (Tested on Python 3.11)
- **Gmail Integration:** `google-api-python-client`, `google-auth-oauthlib`, `google-auth-httplib2`
- **Spreadsheet Processing:** `openpyxl` (XLSX), `xlrd` (XLS)
- **HTML Parsing:** `beautifulsoup4`
- **HTTP Client:** `requests`
- **Configuration:** `python-dotenv`
- **Automation / Scripting:** PowerShell

---

## 13. Environment Variables

Create a local `.env` file in the project root based on `.env.example`:

```env
# Telegram Configuration
TELEGRAM_BOT_TOKEN=your_telegram_bot_token_here
TELEGRAM_CHAT_ID=your_telegram_chat_id_here

# Gmail Accounts to Monitor
MAILBOX1=your_personal_email@gmail.com
MAILBOX2=your_college_email@gmail.com

# Student Identifier Details
REGISTRATION_NUMBER=your_registration_number_here
NEOPAT_ID=your_neopat_id_here

# Polling Interval (in seconds)
CHECK_INTERVAL=300
```

---

## 14. Gmail API Setup

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (e.g., `PlacementMonitor`).
3. Navigate to **APIs & Services > Library** and enable the **Gmail API**.
4. Navigate to **APIs & Services > OAuth consent screen**:
   - Choose **External** user type.
   - Fill in application details and add your monitored email addresses under **Test Users**.
5. Navigate to **APIs & Services > Credentials**:
   - Click **Create Credentials > OAuth client ID**.
   - Application type: **Desktop app**.
   - Name: `PlacementMonitor Desktop`.
6. Download the generated client JSON, rename it to `credentials.json`, and place it in the project root directory.

> **Note:** The `credentials.json` file is ignored by git and must remain private on your local machine.

---

## 15. Telegram Bot Setup

1. Open Telegram and search for [@BotFather](https://t.me/BotFather).
2. Send `/newbot` and follow the prompts to create your bot.
3. Copy the provided **HTTP API Bot Token** and set it as `TELEGRAM_BOT_TOKEN` in `.env`.
4. Start a conversation with your newly created bot by clicking `/start`.
5. Retrieve your numerical **Chat ID**:
   - Send a message to [@userinfobot](https://t.me/userinfobot) or [@raw_data_bot](https://t.me/raw_data_bot).
   - Copy the numerical ID and set it as `TELEGRAM_CHAT_ID` in `.env`.

---

## 16. Local Installation

### Step 1: Clone the Repository
```bash
git clone https://github.com/psprashanth25/Mail-Manager.git
cd Mail-Manager
```

### Step 2: Create and Activate Virtual Environment

**Windows (PowerShell):**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**Linux / macOS:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 4: Configure Credentials
1. Place your `credentials.json` in the project root.
2. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```
3. Update `.env` with your actual Telegram bot credentials, emails, and student IDs.

---

## 17. Running the Project

### Option A: Using the PowerShell Launcher (Windows)
```powershell
.\start_monitor.ps1
```

### Option B: Using Python Directly
```bash
python src/main.py
```

### Initial Gmail Account Authentication
On the first run, authenticate each mailbox in your browser:
```powershell
# Authenticate Mailbox 1
python src/gmail_client.py 1

# Authenticate Mailbox 2
python src/gmail_client.py 2
```
This generates `data/token_mailbox1.json` and `data/token_mailbox2.json` locally. Subsequent runs will use and refresh these tokens automatically.

---

## 18. Configuration & CLI Utilities

The engine supports several diagnostic and testing flags:

```powershell
# Run a specific number of polling cycles (test mode)
python src/monitor.py --cycles 3

# Override polling interval (in seconds)
python src/monitor.py --interval 30

# Recheck recent Mailbox 1 emails against the eligibility phrase (dry-run)
python src/monitor.py --recheck-mailbox1

# Scan recent Mailbox 2 emails and Excel attachments (dry-run)
python src/monitor.py --recheck-mailbox2

# Test Telegram connectivity
python src/telegram_bot.py

# Run unit tests on eligibility matcher
python src/matcher.py --verify
```

---

## 19. Example Telegram Alerts

### Eligibility Alert Message
```text
🚨 Placement Eligibility Alert! 🚨

🏢 Company: PharmEasy Pvt Ltd
📬 Source Mailbox: [Redacted Personal Gmail]
👤 Sender: VIT - CDC Office <noreply.cdcinfo@vitstudent.ac.in>
📌 Subject: Congratulations! You're Eligible for PharmEasy Pvt Ltd Placement Drive
📅 Date: Thu, 17 Sep 2026 11:33:32 +0000
🔍 Matched Phrase: "you are eligible for this company"
```

### Shortlist Alert Message
```text
🎯 Placement Shortlist Alert! 🎯

🏢 Company: Exxonmobil
📬 Source Mailbox: [Redacted Student Gmail]
👤 Sender: 'No Reply CDC Info' via 2027 CDC <students.cdc2027@vitap.ac.in>
📌 Subject: Exxonmobil next round of selection process (Technical Interview) is scheduled on (18-09-2026) 08:00 AM @ CDC 717, SJT 7th floor - VIT Vellore
📅 Date: Thu, 17 Sep 2026 18:41:49 +0530
🆔 Matched Identifier: Both identifiers ([Registration No], [NeoPat ID])
📎 Attachment: Exxonmobil interview shortlist.xlsx

📋 Useful Context:
• Venue: CDC 717, SJT 7th floor - VIT Vellore
• Test/Interview Date: 18-09-2026
• Reporting Time: 08:00 AM
• Next Steps: Technical Interview; Carry printed copies of resume; Formal attire mandatory; Bring personal laptop
• Important Links: https://vit.ac.in/placement
• Contact Details: cdc@vit.ac.in
```

---

## 20. Security Notes

- **Zero Secrets Committed:** `.env`, `credentials.json`, and OAuth tokens in `data/` are strictly ignored by Git.
- **Read-Only Gmail Scope:** The application requests `https://www.googleapis.com/auth/gmail.readonly`. It cannot send emails, modify labels, delete messages, or alter mailbox contents.
- **Ephemeral Attachment Processing:** Attachments downloaded for scanning are unlinked immediately after parsing.
- **Sanitized Logging:** Console logs and diagnostics mask tokens and avoid printing sensitive authentication payload data.

---

## 21. Future Improvements (Planned Work)

- [ ] **Dockerization:** Containerize the application for containerized environments.
- [ ] **Cloud Hosting:** 24/7 background worker deployment on Oracle Cloud Always Free or GCP e2-micro compute instances.
- [ ] **Push-Based Architecture:** Gmail Pub/Sub webhook integration for event-driven trigger instead of polling.
- [ ] **PDF Shortlist Scanning:** Support for parsing text within attached PDF shortlists.

---

## 22. Troubleshooting

| Issue | Cause | Solution |
| :--- | :--- | :--- |
| `Credentials file not found` | Missing `credentials.json` | Download desktop client credentials from Google Cloud Console and place in project root. |
| `Token refresh failed` | Expired/revoked OAuth consent | Delete the corresponding `data/token_mailbox*.json` and re-run `python src/gmail_client.py <1 or 2>`. |
| `TELEGRAM_BOT_TOKEN is missing` | `.env` not populated | Ensure `.env` exists in the project root with valid `TELEGRAM_BOT_TOKEN` and `TELEGRAM_CHAT_ID`. |
| `Script execution policy error` | Windows PowerShell policy | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` before activating `.venv`. |

---

## 23. Author

**P. S. Prashanth**
- GitHub: [@psprashanth25](https://github.com/psprashanth25)
- Repository: [Mail-Manager](https://github.com/psprashanth25/Mail-Manager)
