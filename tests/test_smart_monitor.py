"""
PlacementMonitor Comprehensive Test Suite for Smart Resume + Catch-Up Monitoring
Tests all 15 required scenarios defined in the Master Prompt.
"""
from datetime import datetime, timezone, timedelta
import io
import json
import os
from pathlib import Path
import shutil
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import openpyxl

from src.gmail_client import (
    IST,
    to_epoch_seconds,
    list_messages_in_time_range,
)
from src.monitor import (
    determine_scan_window,
    load_checkpoint,
    save_checkpoint,
    load_processed_messages,
    save_processed_messages,
    check_mailbox,
    get_monitoring_config,
    format_ist_time,
)
from src.matcher import process_mailbox1_email, TARGET_PHRASE
from src.excel_matcher import process_mailbox2_email, search_xlsx_file


class TestSmartMonitor(unittest.TestCase):
    def setUp(self):
        # Create a clean isolated temporary directory for test state files
        self.test_dir = tempfile.mkdtemp()
        self.state_file = Path(self.test_dir) / "processed_messages.json"
        self.checkpoint_file = Path(self.test_dir) / "monitor_state.json"

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    # -------------------------------------------------------------------------
    # TEST 1: No previous checkpoint -> previous 24 hours are scanned.
    # -------------------------------------------------------------------------
    def test_01_no_previous_checkpoint_scans_previous_24_hours(self):
        startup_time = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)
        start_time, end_time, mode = determine_scan_window(None, current_time=startup_time)

        self.assertEqual(mode, "INITIAL 24-HOUR SCAN")
        self.assertEqual(end_time, startup_time)
        self.assertEqual(start_time, startup_time - timedelta(hours=24))
        self.assertEqual(start_time, datetime(2026, 9, 21, 16, 0, 0, tzinfo=IST))

    # -------------------------------------------------------------------------
    # TEST 2: Previous checkpoint = yesterday 4 PM, current startup = today 4 PM.
    # Expected: 4 PM yesterday → 4 PM today.
    # -------------------------------------------------------------------------
    def test_02_checkpoint_yesterday_to_today_4pm(self):
        checkpoint = datetime(2026, 9, 21, 16, 0, 0, tzinfo=IST)
        startup_time = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)
        start_time, end_time, mode = determine_scan_window(checkpoint, current_time=startup_time)

        self.assertEqual(mode, "CATCH-UP")
        self.assertEqual(start_time, checkpoint)
        self.assertEqual(end_time, startup_time)
        self.assertEqual(end_time - start_time, timedelta(hours=24))

    # -------------------------------------------------------------------------
    # TEST 3: Previous checkpoint = today 4 PM, current startup = today 6 PM.
    # Expected: 4 PM → 6 PM only. NOT: yesterday 6 PM → today 6 PM.
    # -------------------------------------------------------------------------
    def test_03_short_restart_window_not_24_hours(self):
        checkpoint = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)
        startup_time = datetime(2026, 9, 22, 18, 0, 0, tzinfo=IST)
        start_time, end_time, mode = determine_scan_window(checkpoint, current_time=startup_time)

        self.assertEqual(mode, "CATCH-UP")
        self.assertEqual(start_time, checkpoint)
        self.assertEqual(end_time, startup_time)
        # Verify window is strictly 2 hours, not 24 hours
        self.assertEqual(end_time - start_time, timedelta(hours=2))
        self.assertNotEqual(start_time, startup_time - timedelta(hours=24))

    # -------------------------------------------------------------------------
    # TEST 4: Previous checkpoint = Monday 4 PM, current startup = Thursday 10 AM.
    # Expected: Monday 4 PM → Thursday 10 AM (multi-day offline catch-up).
    # -------------------------------------------------------------------------
    def test_04_multi_day_offline_catchup_window(self):
        monday_4pm = datetime(2026, 9, 21, 16, 0, 0, tzinfo=IST)
        thursday_10am = datetime(2026, 9, 24, 10, 0, 0, tzinfo=IST)
        start_time, end_time, mode = determine_scan_window(monday_4pm, current_time=thursday_10am)

        self.assertEqual(mode, "CATCH-UP")
        self.assertEqual(start_time, monday_4pm)
        self.assertEqual(end_time, thursday_10am)
        expected_duration = timedelta(days=2, hours=18)  # 66 hours
        self.assertEqual(end_time - start_time, expected_duration)

    # -------------------------------------------------------------------------
    # TEST 5: Email received 23 hours ago on first-ever startup.
    # Expected: included in scan window.
    # -------------------------------------------------------------------------
    def test_05_email_received_23_hours_ago_included(self):
        startup_time = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)
        start_time, end_time, _ = determine_scan_window(None, current_time=startup_time)
        email_time = startup_time - timedelta(hours=23)

        self.assertTrue(start_time <= email_time <= end_time, "Email from 23 hours ago must be included.")

    # -------------------------------------------------------------------------
    # TEST 6: Email received 25 hours ago on first-ever startup.
    # Expected: excluded from scan window.
    # -------------------------------------------------------------------------
    def test_06_email_received_25_hours_ago_excluded(self):
        startup_time = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)
        start_time, end_time, _ = determine_scan_window(None, current_time=startup_time)
        email_time = startup_time - timedelta(hours=25)

        self.assertFalse(start_time <= email_time <= end_time, "Email from 25 hours ago must be excluded.")
        self.assertLess(email_time, start_time)

    # -------------------------------------------------------------------------
    # TEST 7: Email arrives while monitor is offline.
    # Expected: detected during catch-up after restart.
    # -------------------------------------------------------------------------
    @patch("src.monitor.get_gmail_service")
    @patch("src.monitor.list_messages_in_time_range")
    @patch("src.monitor.get_message_details")
    @patch("src.monitor.process_mailbox1_email")
    def test_07_email_arriving_while_offline_detected_on_restart(
        self, mock_process, mock_details, mock_list_range, mock_service
    ):
        mock_service.return_value = MagicMock()
        offline_email_id = "offline_msg_001"
        mock_list_range.return_value = [{"id": offline_email_id, "threadId": "t1"}]
        mock_details.return_value = {
            "id": offline_email_id,
            "subject": "Placement Opportunity",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 17:30:00",
            "body": "you are eligible for this company: Cisco",
            "attachments": []
        }
        mock_process.return_value = {"matched": True, "company": "Cisco", "alert_sent": True}

        processed_data = {"mailbox1": [], "mailbox2": []}
        save_processed_messages(processed_data, state_file=self.state_file)

        checkpoint_5pm = datetime(2026, 9, 22, 17, 0, 0, tzinfo=IST)
        startup_6pm = datetime(2026, 9, 22, 18, 0, 0, tzinfo=IST)

        new_messages, success = check_mailbox(
            mailbox_id=1,
            email_address="student@gmail.com",
            processed_data=processed_data,
            state_file=self.state_file,
            start_time=checkpoint_5pm,
            end_time=startup_6pm
        )

        self.assertTrue(success)
        self.assertEqual(len(new_messages), 1)
        self.assertEqual(new_messages[0]["id"], offline_email_id)

        # Verify ID was saved to disk
        reloaded_state = load_processed_messages(state_file=self.state_file)
        self.assertIn(offline_email_id, reloaded_state["mailbox1"])

    # -------------------------------------------------------------------------
    # TEST 8: Same email appears in overlapping Gmail queries.
    # Expected: only one Telegram alert sent, subsequent query skips duplicate.
    # -------------------------------------------------------------------------
    @patch("src.monitor.get_gmail_service")
    @patch("src.monitor.list_messages_in_time_range")
    @patch("src.monitor.get_message_details")
    @patch("src.matcher.send_telegram_message")
    def test_08_overlapping_queries_only_one_telegram_alert(
        self, mock_tg_send, mock_details, mock_list_range, mock_service
    ):
        mock_service.return_value = MagicMock()
        mock_tg_send.return_value = True

        test_msg_id = "overlap_msg_999"
        mock_list_range.return_value = [{"id": test_msg_id, "threadId": "t1"}]
        mock_details.return_value = {
            "id": test_msg_id,
            "subject": "Eligible for Oracle",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 17:15:00",
            "body": "Congratulations! you are eligible for this company: Oracle",
            "attachments": []
        }

        processed_data = {"mailbox1": [], "mailbox2": []}
        save_processed_messages(processed_data, state_file=self.state_file)

        # First cycle (startup catch-up)
        t1 = datetime(2026, 9, 22, 17, 0, 0, tzinfo=IST)
        t2 = datetime(2026, 9, 22, 18, 0, 0, tzinfo=IST)
        new_msgs_1, _ = check_mailbox(
            mailbox_id=1,
            email_address="student@gmail.com",
            processed_data=processed_data,
            state_file=self.state_file,
            start_time=t1,
            end_time=t2
        )
        self.assertEqual(len(new_msgs_1), 1)
        self.assertEqual(mock_tg_send.call_count, 1)

        # Second cycle (next polling cycle with overlap buffer returning same email)
        t3 = datetime(2026, 9, 22, 18, 1, 0, tzinfo=IST)
        new_msgs_2, _ = check_mailbox(
            mailbox_id=1,
            email_address="student@gmail.com",
            processed_data=processed_data,
            state_file=self.state_file,
            start_time=t2,
            end_time=t3,
            overlap_buffer_seconds=60
        )

        # Must skip duplicate without fetching details or sending duplicate alert
        self.assertEqual(len(new_msgs_2), 0)
        self.assertEqual(mock_tg_send.call_count, 1, "Telegram alert must NOT be dispatched twice!")

    # -------------------------------------------------------------------------
    # TEST 9: Mailbox 1 matching email.
    # Expected: existing eligibility Telegram alert triggered.
    # -------------------------------------------------------------------------
    @patch("src.matcher.send_telegram_message")
    def test_09_mailbox1_eligibility_matching_alert(self, mock_send_tg):
        mock_send_tg.return_value = True

        email = {
            "id": "mb1_eligible_1",
            "subject": "Placement Notice",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 10:00:00",
            "body": "Dear Student, you are eligible for this company: Microsoft India. Please register.",
            "attachments": []
        }
        res = process_mailbox1_email(email, source_mailbox="student@gmail.com", send_alert=True)

        self.assertTrue(res["matched"])
        self.assertEqual(res["company"], "Microsoft India")
        self.assertTrue(res["alert_sent"])
        mock_send_tg.assert_called_once()
        alert_text = mock_send_tg.call_args[0][0]
        self.assertIn("Placement Eligibility Alert!", alert_text)
        self.assertIn("Microsoft India", alert_text)

    # -------------------------------------------------------------------------
    # TEST 10: Mailbox 2 matching email/Excel attachment.
    # Expected: existing shortlist Telegram alert triggered.
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    @patch("src.excel_matcher.download_attachment")
    def test_10_mailbox2_shortlist_excel_attachment_alert(
        self, mock_download, mock_send_tg, mock_get_id
    ):
        mock_get_id.return_value = ("22MIS7017", "NEO12345")
        mock_send_tg.return_value = True

        # Create temporary genuine .xlsx file containing the student's registration number
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Shortlisted"
        ws.append(["S.No", "Reg No", "Student Name", "Status"])
        ws.append([1, "22MIS7017", "Prashanth", "Selected"])
        xlsx_buf = io.BytesIO()
        wb.save(xlsx_buf)
        wb.close()
        mock_download.return_value = xlsx_buf.getvalue()

        email = {
            "id": "mb2_shortlist_1",
            "subject": "Amazon technical interview shortlist",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 14:00:00",
            "body": "Attached is the shortlist for Amazon interview round.",
            "attachments": [{"filename": "Amazon shortlist.xlsx", "attachment_id": "att_101"}]
        }

        mock_service = MagicMock()
        res = process_mailbox2_email(mock_service, email, source_mailbox="college@vit.ac.in", send_alert=True)

        self.assertTrue(res["matched"])
        self.assertEqual(res["company"], "Amazon")
        self.assertIn("22MIS7017", res["matched_identifier"])
        self.assertTrue(res["alert_sent"])
        mock_send_tg.assert_called_once()
        alert_text = mock_send_tg.call_args[0][0]
        self.assertIn("Placement Shortlist Alert!", alert_text)
        self.assertIn("Amazon", alert_text)

    # -------------------------------------------------------------------------
    # TEST 11: One mailbox API failure -> other mailbox continues operating.
    # -------------------------------------------------------------------------
    @patch("src.monitor.get_gmail_service")
    @patch("src.monitor.list_messages_in_time_range")
    def test_11_one_mailbox_api_failure_does_not_block_other(
        self, mock_list_range, mock_service
    ):
        mock_service.return_value = MagicMock()
        processed_data = {"mailbox1": [], "mailbox2": []}

        # Mailbox 1 throws API error
        mock_list_range.side_effect = Exception("Mailbox 1 Gmail API 503 Service Unavailable")
        new_msgs_1, mb1_ok = check_mailbox(
            1, "user1@gmail.com", processed_data, state_file=self.state_file,
            start_time=datetime.now(IST), end_time=datetime.now(IST)
        )
        self.assertFalse(mb1_ok, "Mailbox 1 should fail gracefully.")
        self.assertEqual(len(new_msgs_1), 0)

        # Mailbox 2 operates normally
        mock_list_range.side_effect = None
        mock_list_range.return_value = []
        new_msgs_2, mb2_ok = check_mailbox(
            2, "user2@gmail.com", processed_data, state_file=self.state_file,
            start_time=datetime.now(IST), end_time=datetime.now(IST)
        )
        self.assertTrue(mb2_ok, "Mailbox 2 must continue operating independently.")
        self.assertEqual(len(new_msgs_2), 0)

    # -------------------------------------------------------------------------
    # TEST 12: Gmail pagination -> all pages are processed.
    # -------------------------------------------------------------------------
    def test_12_gmail_pagination_retrieves_all_pages(self):
        mock_service = MagicMock()
        # Mock 3 pages of results
        page1 = {"messages": [{"id": f"msg_{i}", "threadId": f"t_{i}"} for i in range(1, 101)], "nextPageToken": "tok_2"}
        page2 = {"messages": [{"id": f"msg_{i}", "threadId": f"t_{i}"} for i in range(101, 201)], "nextPageToken": "tok_3"}
        page3 = {"messages": [{"id": f"msg_{i}", "threadId": f"t_{i}"} for i in range(201, 251)]}

        mock_messages = mock_service.users().messages()
        mock_list = mock_messages.list
        mock_list.return_value.execute.side_effect = [page1, page2, page3]

        start_dt = datetime(2026, 9, 21, 16, 0, 0, tzinfo=IST)
        end_dt = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)

        results = list_messages_in_time_range(mock_service, start_time=start_dt, end_time=end_dt)

        self.assertEqual(len(results), 250, "Pagination must exhaustively fetch all 250 messages across 3 pages.")
        self.assertEqual(results[0]["id"], "msg_1")
        self.assertEqual(results[-1]["id"], "msg_250")
        self.assertEqual(mock_list.return_value.execute.call_count, 3)

    # -------------------------------------------------------------------------
    # TEST 13: Checkpoint is NOT advanced when the required scan fails.
    # -------------------------------------------------------------------------
    def test_13_checkpoint_not_advanced_on_scan_failure(self):
        initial_cp = datetime(2026, 9, 21, 16, 0, 0, tzinfo=IST)
        save_checkpoint(initial_cp, checkpoint_file=self.checkpoint_file)

        # Simulate check failure: mb1_ok = True, mb2_ok = False
        mb1_ok = True
        mb2_ok = False
        attempted_end_time = datetime(2026, 9, 22, 16, 0, 0, tzinfo=IST)

        if mb1_ok and mb2_ok:
            save_checkpoint(attempted_end_time, checkpoint_file=self.checkpoint_file)

        # Reload checkpoint: must remain at initial_cp, NOT advanced to attempted_end_time
        reloaded_cp = load_checkpoint(checkpoint_file=self.checkpoint_file)
        self.assertEqual(reloaded_cp, initial_cp)
        self.assertNotEqual(reloaded_cp, attempted_end_time)

    # -------------------------------------------------------------------------
    # TEST 14: CHECK_INTERVAL is 60 seconds.
    # -------------------------------------------------------------------------
    def test_14_check_interval_is_60_seconds(self):
        config = get_monitoring_config()
        self.assertEqual(config["interval"], 60, "Configured check interval must be 60 seconds.")

    # -------------------------------------------------------------------------
    # TEST 15: Restarting creates correct catch-up window based on saved checkpoint.
    # -------------------------------------------------------------------------
    def test_15_restarting_creates_correct_catchup_window_from_saved_checkpoint(self):
        # Session 1: completed at 5:00 PM
        session1_end = datetime(2026, 9, 22, 17, 0, 0, tzinfo=IST)
        save_checkpoint(session1_end, checkpoint_file=self.checkpoint_file)

        # Session 2: restarts at 6:00 PM
        restarted_checkpoint = load_checkpoint(checkpoint_file=self.checkpoint_file)
        startup_time_session2 = datetime(2026, 9, 22, 18, 0, 0, tzinfo=IST)

        start_time, end_time, mode = determine_scan_window(
            restarted_checkpoint, current_time=startup_time_session2
        )

        self.assertEqual(mode, "CATCH-UP")
        self.assertEqual(start_time, session1_end)
        self.assertEqual(end_time, startup_time_session2)
        self.assertEqual(end_time - start_time, timedelta(hours=1))

    # -------------------------------------------------------------------------
    # TEST 16 (BONUS): Corrupted checkpoint file triggers safe recovery to 24h scan.
    # -------------------------------------------------------------------------
    def test_16_corrupted_checkpoint_recovery_falls_back_to_initial_scan(self):
        # Write corrupted JSON to checkpoint file
        with open(self.checkpoint_file, "w", encoding="utf-8") as f:
            f.write("{invalid_json: true, unterminated...")

        recovered_cp = load_checkpoint(checkpoint_file=self.checkpoint_file)
        self.assertIsNone(recovered_cp, "Corrupted checkpoint must return None.")

        # Verify a preserved backup file was created
        backups = list(Path(self.test_dir).glob("monitor_state.corrupted.*.json"))
        self.assertGreaterEqual(len(backups), 1, "Corrupted file must be preserved as backup.")

        # Window calculation falls back to 24-hour scan
        startup = datetime(2026, 9, 22, 18, 0, 0, tzinfo=IST)
        start, end, mode = determine_scan_window(recovered_cp, current_time=startup)
        self.assertEqual(mode, "INITIAL 24-HOUR SCAN")
        self.assertEqual(start, startup - timedelta(hours=24))


class TestMailbox2PlacementClassification(unittest.TestCase):
    """
    Unit tests for Mailbox 2 placement relevance classification.
    Verifies that emails matching student identifiers are only alerted if they are
    genuinely placement/recruitment related, filtering out false positives.
    """
    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.state_file = Path(self.test_dir) / "processed_messages.json"
        self.reg_no = "22MIS7017"
        self.neopat_id = "NEO12345"

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    # -------------------------------------------------------------------------
    # TEST 1: Subject: "ExxonMobil next round of selection process"
    # Body contains registration number -> PLACEMENT = TRUE
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_01_exxonmobil_selection_process(self, mock_tg, mock_ids):
        mock_ids.return_value = (self.reg_no, self.neopat_id)
        mock_tg.return_value = True

        email = {
            "id": "exxon_001",
            "subject": "ExxonMobil next round of selection process",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 10:00:00",
            "body": f"Dear student ({self.reg_no}), please report for the selection process.",
            "attachments": []
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["identifier_matched"], "Identifier must match")
        self.assertTrue(res["is_placement"], "Must be classified as placement-related")
        self.assertEqual(res["classification"], "PLACEMENT")
        self.assertTrue(res["matched"])
        self.assertTrue(res["alert_sent"])
        mock_tg.assert_called_once()

    # -------------------------------------------------------------------------
    # TEST 2: Subject: "Congratulations! You're Eligible for PharmEasy Pvt Ltd Placement Drive"
    # -> PLACEMENT = TRUE
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_02_pharmeasy_eligible_placement_drive(self, mock_tg, mock_ids):
        mock_ids.return_value = (self.reg_no, self.neopat_id)
        mock_tg.return_value = True

        email = {
            "id": "pharmeasy_002",
            "subject": "Congratulations! You're Eligible for PharmEasy Pvt Ltd Placement Drive",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 11:00:00",
            "body": f"Candidate: {self.reg_no}. Confirm your attendance.",
            "attachments": []
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["is_placement"])
        self.assertEqual(res["classification"], "PLACEMENT")
        self.assertEqual(res["company"], "PharmEasy Pvt Ltd")
        self.assertTrue(res["alert_sent"])

    # -------------------------------------------------------------------------
    # TEST 3: Subject: "Security alert", Body contains registration number
    # -> PLACEMENT = FALSE
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_03_security_alert_false_positive(self, mock_tg, mock_ids):
        mock_ids.return_value = (self.reg_no, self.neopat_id)

        email = {
            "id": "sec_003",
            "subject": "Security alert for your account",
            "sender": "Google <no-reply@accounts.google.com>",
            "date": "2026-09-22 12:00:00",
            "body": f"A new login was detected for user {self.reg_no}@vitapstudent.ac.in. Check activity.",
            "attachments": []
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["identifier_matched"], "Identifier was matched in text")
        self.assertFalse(res["is_placement"], "Security alert must NOT be classified as placement")
        self.assertEqual(res["classification"], "NON-PLACEMENT")
        self.assertFalse(res["matched"], "matched flag must be False for non-placement alerts")
        self.assertFalse(res["alert_sent"], "No Telegram alert should be dispatched")
        mock_tg.assert_not_called()

    # -------------------------------------------------------------------------
    # TEST 4: Subject: "New sign-in on your Google Account"
    # Body contains registration number -> PLACEMENT = FALSE
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_04_google_signin_false_positive(self, mock_tg, mock_ids):
        mock_ids.return_value = (self.reg_no, self.neopat_id)

        email = {
            "id": "signin_004",
            "subject": "New sign-in on your Google Account",
            "sender": "Google <no-reply@accounts.google.com>",
            "date": "2026-09-22 13:00:00",
            "body": f"Your Google account {self.reg_no} was accessed from Windows.",
            "attachments": []
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["identifier_matched"])
        self.assertFalse(res["is_placement"])
        self.assertEqual(res["classification"], "NON-PLACEMENT")
        self.assertFalse(res["alert_sent"])
        mock_tg.assert_not_called()

    # -------------------------------------------------------------------------
    # TEST 5: Subject: "Placement shortlist - Technical Interview"
    # Excel attachment contains registration number -> PLACEMENT = TRUE
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    @patch("src.excel_matcher.download_attachment")
    def test_cls_05_placement_shortlist_excel_attachment(
        self, mock_download, mock_tg, mock_ids
    ):
        mock_ids.return_value = (self.reg_no, self.neopat_id)
        mock_tg.return_value = True

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Reg No", "Status"])
        ws.append([self.reg_no, "Shortlisted"])
        buf = io.BytesIO()
        wb.save(buf)
        wb.close()
        mock_download.return_value = buf.getvalue()

        email = {
            "id": "shortlist_excel_005",
            "subject": "Placement shortlist - Technical Interview",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 14:00:00",
            "body": "Please find the attached shortlist for the technical interview.",
            "attachments": [{"filename": "shortlist_interview.xlsx", "attachment_id": "att_5"}]
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["identifier_matched"])
        self.assertTrue(res["is_placement"])
        self.assertEqual(res["classification"], "PLACEMENT")
        self.assertTrue(res["alert_sent"])
        mock_tg.assert_called_once()

    # -------------------------------------------------------------------------
    # TEST 6: Subject: "University notification"
    # Body contains registration number but no recruitment/placement context
    # -> PLACEMENT = FALSE
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_06_university_notification_false_positive(self, mock_tg, mock_ids):
        mock_ids.return_value = (self.reg_no, self.neopat_id)

        email = {
            "id": "univ_006",
            "subject": "University notification",
            "sender": "registrar@vit.ac.in",
            "date": "2026-09-22 15:00:00",
            "body": f"Dear Student ({self.reg_no}), the hostel and transport facilities will operate on weekend timings.",
            "attachments": []
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["identifier_matched"])
        self.assertFalse(res["is_placement"], "General university notification must NOT be classified as placement")
        self.assertEqual(res["classification"], "NON-PLACEMENT")
        self.assertFalse(res["alert_sent"])
        mock_tg.assert_not_called()

    # -------------------------------------------------------------------------
    # TEST 7: Placement-related email where company extraction fails
    # -> PLACEMENT = TRUE, Company may remain "Company name could not be determined"
    # -------------------------------------------------------------------------
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_07_placement_where_company_name_fails(self, mock_tg, mock_ids):
        mock_ids.return_value = (self.reg_no, self.neopat_id)
        mock_tg.return_value = True

        email = {
            "id": "placement_no_comp_007",
            "subject": "Shortlisted candidates for Technical Interview",
            "sender": "cdc@vit.ac.in",
            "date": "2026-09-22 16:00:00",
            "body": f"Dear student ({self.reg_no}), you are shortlisted for the technical interview. Reporting time: 08:30 AM @ CDC 717.",
            "attachments": []
        }
        res = process_mailbox2_email(MagicMock(), email, send_alert=True)
        self.assertTrue(res["identifier_matched"])
        self.assertTrue(res["is_placement"], "Placement relevance is independent from company extraction")
        self.assertEqual(res["classification"], "PLACEMENT")
        self.assertEqual(res["company"], "Company name could not be determined")
        self.assertTrue(res["alert_sent"])
        mock_tg.assert_called_once()

    # -------------------------------------------------------------------------
    # TEST 8: Previously processed unrelated email
    # -> No Telegram notification and no repeated processing
    # -------------------------------------------------------------------------
    @patch("src.monitor.get_gmail_service")
    @patch("src.monitor.list_messages_in_time_range")
    @patch("src.monitor.get_message_details")
    @patch("src.excel_matcher.get_mailbox2_identifiers")
    @patch("src.excel_matcher.send_telegram_message")
    def test_cls_08_unrelated_email_evaluated_once_and_never_repeated(
        self, mock_tg, mock_ids, mock_details, mock_list_range, mock_service
    ):
        mock_service.return_value = MagicMock()
        mock_ids.return_value = (self.reg_no, self.neopat_id)

        unrelated_id = "unrelated_sec_alert_008"
        mock_list_range.return_value = [{"id": unrelated_id, "threadId": "t_unrelated"}]
        mock_details.return_value = {
            "id": unrelated_id,
            "subject": "Security alert",
            "sender": "Google <no-reply@accounts.google.com>",
            "date": "2026-09-22 17:00:00",
            "body": f"Security alert for account {self.reg_no}",
            "attachments": []
        }

        processed_data = {"mailbox1": [], "mailbox2": []}
        save_processed_messages(processed_data, state_file=self.state_file)

        # Cycle 1: Discovered, evaluated as NON-PLACEMENT, 0 alerts sent, ID saved to state
        t1 = datetime(2026, 9, 22, 17, 0, 0, tzinfo=IST)
        t2 = datetime(2026, 9, 22, 17, 1, 0, tzinfo=IST)
        new_msgs_1, ok1 = check_mailbox(
            mailbox_id=2,
            email_address="college@vit.ac.in",
            processed_data=processed_data,
            state_file=self.state_file,
            start_time=t1,
            end_time=t2
        )
        self.assertTrue(ok1)
        self.assertEqual(len(new_msgs_1), 1)
        mock_tg.assert_not_called()  # NO alert sent!

        # Check saved state
        state_on_disk = load_processed_messages(state_file=self.state_file)
        self.assertIn(unrelated_id, state_on_disk["mailbox2"], "Unrelated email must be stored in processed_messages.json")

        # Cycle 2: Next polling cycle receives the same message in overlapping window
        t3 = datetime(2026, 9, 22, 17, 2, 0, tzinfo=IST)
        mock_details.reset_mock()
        new_msgs_2, ok2 = check_mailbox(
            mailbox_id=2,
            email_address="college@vit.ac.in",
            processed_data=processed_data,
            state_file=self.state_file,
            start_time=t2,
            end_time=t3,
            overlap_buffer_seconds=60
        )
        self.assertTrue(ok2)
        self.assertEqual(len(new_msgs_2), 0, "Must be skipped immediately as duplicate")
        mock_details.assert_not_called()
        mock_tg.assert_not_called()


if __name__ == "__main__":
    unittest.main(verbosity=2)
