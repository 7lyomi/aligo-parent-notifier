"""Coordinate previews, receipt checks and durable message dispatch without Flask."""

import hashlib
import json
import threading
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from .aligo import MAX_RECIPIENTS, AligoClient, ApiRejected, BeforeSendError
from .delivery_store import DeliveryStore, fingerprint
from .excel import parse_daily_report, parse_student_db
from .preview import build_preview, validate_row
from .settings import Settings
from .storage import FileStore, RuntimePaths
from .templates import TEMPLATE_TEXTS, TEST_TYPES


@dataclass
class DispatchResult:
    accepted: int = 0
    failed: int = 0
    blocked: int = 0
    previous_accepted: int = 0
    log_failed: bool = False


class NotifierService:
    def __init__(self, settings: Settings, paths: RuntimePaths, sample_path: Path, client=None):
        self.settings = settings
        self.paths = paths
        self.sample_path = sample_path
        self.files = FileStore(paths)
        self.deliveries = DeliveryStore(paths.ledger)
        self.client = client if client is not None else AligoClient()
        self.lock = threading.RLock()

    def import_workbook(self, students_path, report_path):
        rows = build_preview(
            parse_student_db(students_path), parse_daily_report(report_path), self.settings
        )
        self.save_preview(rows)
        return rows

    def save_preview(self, rows):
        self.files.save_preview([validate_row(row, self.settings) for row in rows])

    def load_preview(self):
        rows = [validate_row(row, self.settings) for row in self.files.load_preview()]
        keys = [fingerprint(row, self.settings.delivery_mode) for row in rows]
        states = self.deliveries.states(keys)
        for row, key in zip(rows, keys):
            row.delivery_state = states.get(key, "")
        return rows

    def receipt_signature(self, kind):
        settings = self.settings
        payload = [
            settings.aligo_config(),
            settings.template_code(kind),
            TEMPLATE_TEXTS[kind],
            settings.academy_name,
            settings.test_phone,
        ]
        return hashlib.sha256(
            json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()
        ).hexdigest()

    def receipt_status(self):
        saved = self.files.load_receipts()
        entries = []
        for kind in TEST_TYPES:
            signature = self.receipt_signature(kind)
            item = saved.get(signature, {})
            entries.append(
                {
                    "kind": kind,
                    "requested": bool(item.get("requested")),
                    "confirmed": bool(item.get("confirmed")),
                    "signature": signature,
                }
            )
        return {"entries": entries, "passed": all(item["confirmed"] for item in entries)}

    def record_test_request(self, kind):
        saved = self.files.load_receipts()
        saved[self.receipt_signature(kind)] = {
            "requested": True,
            "confirmed": False,
            "sent_at": datetime.now().isoformat(),
        }
        self.files.save_receipts(saved)

    def confirm_receipt(self, kind, confirmed):
        signature = self.receipt_signature(kind)
        saved = self.files.load_receipts()
        if not saved.get(signature, {}).get("requested"):
            raise ValueError("해당 유형의 1건 테스트 요청을 먼저 진행해주세요.")
        if not confirmed:
            raise ValueError("실제 휴대전화 수신을 확인한 후 체크해주세요.")
        saved[signature]["confirmed"] = True
        self.files.save_receipts(saved)

    def missing_receipts(self, rows):
        confirmed = {
            entry["kind"] for entry in self.receipt_status()["entries"] if entry["confirmed"]
        }
        return {row.test_type for row in rows} - confirmed

    def dispatch(self, targets, mode_override=None):
        mode = mode_override or self.settings.delivery_mode
        unique = {}
        for row in targets:
            unique.setdefault(fingerprint(row, mode), row)
        prior = self.deliveries.states(unique)
        reserved, blocked = self.deliveries.reserve(unique)
        result = DispatchResult(
            blocked=len(targets) - len(reserved),
            previous_accepted=sum(prior.get(key) == "accepted" for key in blocked),
        )
        groups = defaultdict(list)
        for key in reserved:
            groups[unique[key].template_code].append((key, unique[key]))
        logs = []
        for code, items in groups.items():
            for offset in range(0, len(items), MAX_RECIPIENTS):
                batch = items[offset : offset + MAX_RECIPIENTS]
                keys, rows = [item[0] for item in batch], [item[1] for item in batch]
                sent_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                try:
                    response = self.client.send(rows, code, self.settings)
                except (BeforeSendError, ApiRejected) as error:
                    self.deliveries.finish(keys, "failed")
                    state, message, batch_id = "FAILED", str(error), ""
                    result.failed += len(rows)
                except Exception:
                    self.deliveries.finish(keys, "unknown")
                    state, message, batch_id = (
                        "UNKNOWN",
                        "응답 확인 불가: 알리고 내역 확인 필요",
                        "",
                    )
                    result.failed += len(rows)
                else:
                    info = response.get("info") if isinstance(response.get("info"), dict) else {}
                    batch_id = str(info.get("mid") or response.get("msg_id") or "")
                    # Persist acceptance before the auxiliary CSV is written.
                    self.deliveries.finish(
                        keys, "simulated" if self.settings.is_demo else "accepted", batch_id
                    )
                    state, message = (
                        ("DEMO", "데모 처리")
                        if self.settings.is_demo
                        else ("ACCEPTED", "요청 접수; 최종 수신 결과는 알리고에서 확인")
                    )
                    result.accepted += len(rows)
                logs.extend(self._log_row(row, sent_at, state, message, batch_id) for row in rows)
        if logs:
            try:
                self.files.append_log(logs)
            except OSError:
                result.log_failed = True
        return result

    def _log_row(self, row, sent_at, state, message, batch_id):
        return {
            "sent_at": sent_at,
            "student_name": row.student_name,
            "student_id": row.student_id,
            "parent_name": row.parent_name,
            "parent_phone": row.parent_phone,
            "test_type": row.test_type,
            "message": row.message,
            "template_code": row.template_code,
            "test_mode": "DEMO" if self.settings.is_demo else self.settings.test_mode,
            "estimated_cost": row.estimated_cost,
            "result_code": state,
            "result_message": message,
            "batch_id": batch_id,
        }
