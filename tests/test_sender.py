"""Behavioral regression tests. All external message requests are mocked."""

import csv
import sys
from concurrent.futures import ThreadPoolExecutor
from copy import deepcopy
from dataclasses import replace
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from notifier import aligo, create_app, excel, preview
from notifier.delivery_store import DeliveryStore, fingerprint
from notifier.input_rules import normalize_percentage
from notifier.settings import Settings
from notifier.templates import TEST_TYPES


class SenderTests(TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory(prefix="sender-test-")
        self.addCleanup(self.temp.cleanup)
        self.network = patch.object(
            aligo.requests, "post", side_effect=AssertionError("External API forbidden in tests")
        )
        self.network.start()
        self.addCleanup(self.network.stop)
        settings = Settings(
            mode="demo",
            school_name="테스트학원",
            api_key="mock-key",
            user_id="mock-user",
            sender="01000000001",
            sender_key="mock-sender",
            template_none="MOCK_NONE",
            template_weekly="MOCK_WEEK",
            template_monthly="MOCK_MONTH",
            test_phone="01000000009",
        )
        self.app = create_app(settings, runtime_dir=Path(self.temp.name))
        self.app.config.update(TESTING=True, SECRET_KEY="test-only-session-secret")
        self.service = self.app.extensions["notifier"]
        self.client = self.app.test_client()
        self.students = excel.parse_student_db(self.service.sample_path)
        self.reports = excel.parse_daily_report(self.service.sample_path)
        self.rows = self.build_preview(self.students, self.reports)

    def build_preview(self, students, reports):
        return preview.build_preview(students, reports, self.service.settings)

    def validate(self, row):
        return preview.validate_row(row, self.service.settings)

    def post(self, path, data=None, client=None):
        client = client or self.client
        client.get("/")
        with client.session_transaction() as session:
            token = session["csrf"]
        form = {"csrf_token": token, "preview_revision": self.service.files.preview_revision()}
        form.update(data or {})
        return client.post(path, data=form, follow_redirects=True)

    def live(self):
        self.service.settings = replace(self.service.settings, mode="live", test_mode="N")
        self.rows = [
            self.validate(replace(row, parent_phone="01000000001", parent_name="가상보호자"))
            for row in self.rows
        ]

    def ready(self, rows=None):
        self.service.save_preview(rows if rows is not None else self.rows)

    def confirm_types(self, types):
        self.service.files.save_receipts(
            {
                self.service.receipt_signature(kind): {"requested": True, "confirmed": True}
                for kind in types
            }
        )

    @staticmethod
    def api_response(data):
        response = Mock()
        response.raise_for_status.return_value = None
        response.json.return_value = data
        return response

    def test_sample_ids_and_native_percentages(self):
        self.assertEqual(len(self.rows), 12)
        self.assertTrue(all(r.is_ready for r in self.rows))
        self.assertEqual(self.rows[0].student_name, self.rows[1].student_name)
        self.assertNotEqual(self.rows[0].student_id, self.rows[1].student_id)
        self.assertNotEqual(self.rows[0].parent_phone, self.rows[1].parent_phone)
        self.assertEqual(self.rows[0].homework_percent, "90%")
        self.assertEqual(self.rows[1].homework_percent, "100%")

    def test_namesake_without_id_is_blocked(self):
        reports = [replace(self.reports[0], student_id="")]
        rows = self.build_preview(self.students, reports)
        self.assertFalse(rows[0].is_ready)
        self.assertIn("동명이인", rows[0].missing_fields)
        self.assertEqual(rows[0].parent_phone, "")

    def test_id_name_mismatch_is_blocked(self):
        rows = self.build_preview(
            self.students, [replace(self.reports[0], student_name="다른 가상학생")]
        )
        self.assertFalse(rows[0].is_ready)
        self.assertEqual(rows[0].parent_phone, "")

    def test_duplicate_db_id_is_rejected(self):
        wb = excel.load_workbook(self.service.sample_path)
        wb["학생DB"]["A3"] = wb["학생DB"]["A2"].value
        with patch.object(excel, "load_workbook", return_value=wb):
            with self.assertRaisesRegex(ValueError, "중복 학생ID"):
                excel.parse_student_db(self.service.sample_path)

    def test_legacy_unique_name_still_matches(self):
        row = self.build_preview(self.students, [replace(self.reports[2], student_id="")])[0]
        self.assertTrue(row.is_ready)
        self.assertEqual(row.student_id, "STU003")

    def test_duplicate_report_blocks_both_rows(self):
        rows = self.build_preview(self.students, [self.reports[0], deepcopy(self.reports[0])])
        self.assertTrue(all(not r.is_ready for r in rows))

    def test_send_n_is_excluded(self):
        students = deepcopy(self.students)
        students["id:STU001"].send_yn = "N"
        self.assertEqual(len(self.build_preview(students, self.reports)), 11)

    def test_percent_formats_and_ranges(self):
        for value, fmt, expected in [
            (0.9, "0%", "90%"),
            (1, "0%", "100%"),
            (90, "General", "90%"),
            ("90%", "", "90%"),
            (1, "General", "1%"),
        ]:
            self.assertEqual(normalize_percentage(value, fmt), expected)
        for value in ["101%", "-1%", "abc", "NaN"]:
            with self.subTest(value=value):
                self.assertFalse(
                    self.validate(replace(self.rows[0], homework_percent=value)).is_ready
                )

    def test_invalid_numbers_dates_and_phones(self):
        cases = [
            {"word_total_count": "3.5"},
            {"word_wrong_count": "-1"},
            {"word_wrong_count": "31"},
            {"parent_phone": "123"},
            {"parent_phone": "010hello12345678"},
            {"date_text": "2026-02-30"},
            {"test_level": "3.5"},
            {"test_score": "101"},
            {"test_average": "NaN"},
        ]
        for change in cases:
            with self.subTest(change=change):
                self.assertFalse(self.validate(replace(self.rows[1], **change)).is_ready)

    def test_test_type_rules(self):
        row = self.validate(replace(self.rows[1], test_type=""))
        self.assertFalse(row.is_ready)
        row = self.validate(replace(self.rows[1], test_type="없음"))
        self.assertTrue(row.is_ready)
        self.assertNotIn("■ 주간 테스트", row.message)

    def test_real_mode_rejects_demo_numbers(self):
        self.service.settings = replace(self.service.settings, mode="live")
        self.assertFalse(self.validate(self.rows[0]).is_ready)

    def test_actual_flask_upload_and_preview(self):
        response = self.post(
            "/preview",
            {"combined_file": (BytesIO(self.service.sample_path.read_bytes()), "sample.xlsx")},
        )
        self.assertEqual(response.status_code, 200)
        self.assertEqual(len(self.service.load_preview()), 12)
        self.assertIn("STU001", response.get_data(as_text=True))
        self.assertIn("90%", response.get_data(as_text=True))

    def test_csrf_and_stale_preview_are_blocked(self):
        self.ready()
        self.assertEqual(self.client.post("/send", data={"mode": "all"}).status_code, 400)
        response = self.post("/send", {"mode": "all", "preview_revision": "stale"})
        self.assertIn("미리보기가 변경", response.get_data(as_text=True))
        self.assertFalse(self.service.deliveries.states([fingerprint(self.rows[0], "demo")]))

    def test_edited_message_requires_preview_review(self):
        self.ready()
        response = self.post(
            "/send", {"mode": "selected", "selected": "0", "attendance_0": "수정한 등원상태"}
        )
        self.assertIn("수정된 내용을", response.get_data(as_text=True))
        self.assertFalse(self.service.load_preview()[0].delivery_state)

    def test_demo_never_calls_api_and_repeat_is_blocked(self):
        self.ready()
        first = self.post("/send", {"mode": "selected", "selected": "0"})
        self.assertIn("접수 1건", first.get_data(as_text=True))
        second = self.post(
            "/send", {"mode": "selected", "selected": "0"}, client=self.app.test_client()
        )
        self.assertIn("접수 0건", second.get_data(as_text=True))
        self.assertIn("제외 1건", second.get_data(as_text=True))
        store = DeliveryStore(self.service.paths.ledger)
        self.assertEqual(
            store.states([fingerprint(self.rows[0], "demo")])[fingerprint(self.rows[0], "demo")],
            "simulated",
        )

    def test_bulk_and_selected_share_receipt_gate(self):
        self.live()
        self.ready(self.rows[:3])
        for form in [{"mode": "all"}, {"mode": "selected", "selected": ["0", "1", "2"]}]:
            response = self.post("/send", form)
            self.assertIn("수신 확인이 필요", response.get_data(as_text=True))

    def test_gate_requires_matching_template_configuration(self):
        self.live()
        self.confirm_types(TEST_TYPES)
        self.assertTrue(self.service.receipt_status()["passed"])
        self.service.settings = replace(self.service.settings, template_none="CHANGED_TEMPLATE")
        self.assertFalse(self.service.receipt_status()["passed"])

    def test_test_request_uses_own_phone_and_manual_receipt(self):
        self.live()
        self.ready(self.rows[:1])
        captured = []

        def fake(url, data, timeout):
            if url == aligo.ALIGO_TOKEN_URL:
                return self.api_response({"code": 0, "token": "mock-token"})
            captured.append(data)
            return self.api_response({"code": 0, "info": {"mid": "MOCK-1"}})

        with patch.object(aligo.requests, "post", side_effect=fake):
            self.post("/send-test", {"selected": "0"})
        self.assertEqual(captured[0]["receiver_1"], "01000000009")
        self.assertEqual(captured[0]["failover"], "N")
        self.assertFalse(self.service.receipt_status()["entries"][0]["confirmed"])
        self.post("/confirm-test", {"kind": "없음", "receipt_confirmed": "yes"})
        self.assertTrue(self.service.receipt_status()["entries"][0]["confirmed"])

    def test_accepted_request_prevents_repeat(self):
        self.live()
        self.ready(self.rows[:1])
        self.confirm_types(["없음"])
        fake = Mock(return_value={"code": 0, "info": {"mid": "MOCK-2"}})
        with patch.object(self.service.client, "send", fake):
            self.post("/send", {"mode": "all"})
            self.post("/send", {"mode": "all"})
        self.assertEqual(fake.call_count, 1)

    def test_timeout_remains_unknown_and_cannot_retry(self):
        self.live()
        self.ready(self.rows[:1])
        self.confirm_types(["없음"])
        fake = Mock(side_effect=aligo.SendUncertain("timeout"))
        with patch.object(self.service.client, "send", fake):
            self.post("/send", {"mode": "all"})
            self.post("/send", {"mode": "all"})
        self.assertEqual(fake.call_count, 1)
        self.assertEqual(self.service.load_preview()[0].delivery_state, "unknown")

    def test_rejected_request_can_retry(self):
        self.live()
        self.ready(self.rows[:1])
        self.confirm_types(["없음"])
        fake = Mock(
            side_effect=[aligo.ApiRejected("rejected"), {"code": 0, "info": {"mid": "MOCK-3"}}]
        )
        with patch.object(self.service.client, "send", fake):
            self.post("/send", {"mode": "all"})
            self.post("/send", {"mode": "all"})
        self.assertEqual(fake.call_count, 2)
        self.assertEqual(self.service.load_preview()[0].delivery_state, "accepted")

    def test_response_without_result_code_blocks_retry(self):
        self.live()
        self.ready(self.rows[:1])
        self.confirm_types(["없음"])
        send_calls = []

        def fake(url, data, timeout):
            if url == aligo.ALIGO_TOKEN_URL:
                return self.api_response({"code": 0, "token": "mock-token"})
            send_calls.append(data)
            return self.api_response({"message": "ambiguous response"})

        with patch.object(aligo.requests, "post", side_effect=fake):
            self.post("/send", {"mode": "all"})
            self.post("/send", {"mode": "all"})
        self.assertEqual(len(send_calls), 1)
        self.assertEqual(self.service.load_preview()[0].delivery_state, "unknown")

    def test_changed_content_and_mode_have_distinct_keys(self):
        row = self.rows[0]
        self.assertNotEqual(fingerprint(row, "demo"), fingerprint(row, "live"))
        changed = self.validate(replace(row, attendance="수정된 내용"))
        self.assertNotEqual(fingerprint(row, "live"), fingerprint(changed, "live"))

    def test_atomic_reservation_across_connections(self):
        store = DeliveryStore(self.service.paths.ledger)

        def reserve(_):
            return DeliveryStore(self.service.paths.ledger).reserve(["same-key"])[0]

        with ThreadPoolExecutor(max_workers=4) as pool:
            values = list(pool.map(reserve, range(4)))
        self.assertEqual(sum(map(len, values)), 1)
        self.assertEqual(store.states(["same-key"])["same-key"], "pending")

    def test_log_failure_does_not_enable_resend(self):
        self.ready(self.rows[:1])
        with patch.object(self.service.files, "append_log", side_effect=OSError("disk full")):
            self.post("/send", {"mode": "all"})
        response = self.post("/send", {"mode": "all"})
        self.assertIn("접수 0건", response.get_data(as_text=True))

    def test_300_rows_and_csv_logging(self):
        rows = [
            replace(self.rows[i % 12], student_id=f"VIRTUAL{i:03}", student_name=f"가상학생{i:03}")
            for i in range(300)
        ]
        self.ready(rows)
        response = self.post("/send", {"mode": "all"})
        self.assertIn("접수 300건", response.get_data(as_text=True))
        with (self.service.paths.logs / "alimtalk_send_log.csv").open(
            encoding="utf-8-sig", newline=""
        ) as f:
            entries = list(csv.DictReader(f))
        self.assertEqual(len(entries), 300)
        self.assertTrue(all(r["result_code"] == "DEMO" for r in entries))

    def test_batch_splitting_over_500(self):
        rows = [
            replace(self.rows[0], student_id=f"VIRTUAL{i:03}", student_name=f"가상학생{i:03}")
            for i in range(501)
        ]
        self.ready(rows)
        original = self.service.client.send
        counts = []

        def record(batch, code, settings):
            counts.append(len(batch))
            return original(batch, code, settings)

        with patch.object(self.service.client, "send", side_effect=record):
            self.post("/send", {"mode": "all"})
        self.assertEqual(counts, [500, 1])

    def test_clear_preview_preserves_duplicate_ledger(self):
        self.ready(self.rows[:1])
        self.post("/send", {"mode": "all"})
        self.post("/clear-preview")
        self.ready(self.rows[:1])
        response = self.post("/send", {"mode": "all"})
        self.assertIn("제외 1건", response.get_data(as_text=True))

    def test_invalid_selection_rejected(self):
        self.ready()
        self.assertEqual(
            self.post("/send", {"mode": "selected", "selected": "-1"}).status_code, 400
        )

    def test_app_instances_keep_preview_and_delivery_state_separate(self):
        self.ready(self.rows[:1])
        self.post("/send", {"mode": "all"})
        other_app = create_app(
            replace(self.service.settings, school_name="다른 가상학원"),
            runtime_dir=Path(self.temp.name) / "second-app",
        )
        other = other_app.extensions["notifier"]
        self.assertEqual(other.load_preview(), [])
        key = fingerprint(self.rows[0], "demo")
        self.assertEqual(other.deliveries.states([key]), {})
        self.assertEqual(self.service.deliveries.states([key])[key], "simulated")
        self.assertEqual(self.service.settings.school_name, "테스트학원")
        self.assertEqual(other.settings.school_name, "다른 가상학원")
