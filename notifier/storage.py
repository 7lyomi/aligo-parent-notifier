"""Runtime paths, atomic JSON files, uploads and CSV logs."""

import csv
import hashlib
import json
import secrets
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from uuid import uuid4

from .models import PreviewRow


class PreviewReadError(ValueError):
    pass


def atomic_json(path: Path, value):
    temp = path.with_name(path.name + "." + uuid4().hex + ".tmp")
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


@dataclass(frozen=True)
class RuntimePaths:
    root: Path

    @property
    def data(self):
        return self.root / "data"

    @property
    def uploads(self):
        return self.data / "uploads"

    @property
    def logs(self):
        return self.root / "logs"

    @property
    def preview(self):
        return self.data / "preview.json"

    @property
    def receipts(self):
        return self.data / "test_status.json"

    @property
    def ledger(self):
        return self.root / "deliveries.sqlite3"

    @property
    def log(self):
        return self.logs / "alimtalk_send_log.csv"

    def prepare(self):
        for path in (self.data, self.uploads, self.logs):
            path.mkdir(parents=True, exist_ok=True)

    def session_secret(self):
        path = self.root / ".session_secret"
        if not path.exists():
            try:
                with path.open("x", encoding="utf-8") as file:
                    file.write(secrets.token_hex(32))
                path.chmod(0o600)
            except FileExistsError:
                pass
        return path.read_text(encoding="utf-8").strip()


class FileStore:
    def __init__(self, paths: RuntimePaths):
        self.paths = paths
        paths.prepare()

    def load_preview(self):
        if not self.paths.preview.exists():
            return []
        try:
            items = json.loads(self.paths.preview.read_text(encoding="utf-8"))
            names = {field.name for field in fields(PreviewRow)}
            return [PreviewRow(**{k: v for k, v in item.items() if k in names}) for item in items]
        except (ValueError, TypeError, AttributeError):
            raise PreviewReadError(
                "미리보기 파일을 읽지 못했습니다. 엑셀을 다시 불러와주세요."
            ) from None

    def save_preview(self, rows):
        atomic_json(self.paths.preview, [asdict(row) for row in rows])

    def clear_preview(self):
        self.paths.preview.unlink(missing_ok=True)

    def preview_revision(self):
        return (
            hashlib.sha256(self.paths.preview.read_bytes()).hexdigest()
            if self.paths.preview.exists()
            else "empty"
        )

    def load_receipts(self):
        try:
            return json.loads(self.paths.receipts.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def save_receipts(self, values):
        atomic_json(self.paths.receipts, values)

    def clear_receipts(self):
        self.paths.receipts.unlink(missing_ok=True)

    def save_upload(self, upload, fallback: Path):
        if not upload or not upload.filename:
            return fallback
        extension = Path(upload.filename).suffix.lower()
        if extension not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
            raise ValueError("엑셀 파일은 .xlsx 또는 .xlsm 형식으로 업로드해주세요.")
        path = self.paths.uploads / (uuid4().hex + extension)
        upload.save(path)
        return path

    def append_log(self, rows):
        path = self.paths.log
        is_new = not path.exists()
        fieldnames = [
            "sent_at",
            "student_id",
            "student_name",
            "parent_name",
            "parent_phone",
            "test_type",
            "message",
            "template_code",
            "test_mode",
            "estimated_cost",
            "result_code",
            "result_message",
            "batch_id",
        ]

        def safe(value):
            text = str(value)
            return "'" + text if text.lstrip().startswith(("=", "+", "-", "@")) else text

        with path.open("a", newline="", encoding="utf-8-sig") as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            if is_new:
                writer.writeheader()
            writer.writerows({key: safe(value) for key, value in row.items()} for row in rows)
