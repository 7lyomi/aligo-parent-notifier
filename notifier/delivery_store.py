"""Durable, atomic duplicate protection. No API credentials or message bodies stored."""

import hashlib
import json
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

BLOCKING = {"pending", "accepted", "unknown", "simulated"}


def fingerprint(row, mode):
    # Student identity separates namesakes with the same recipient.
    # A genuinely corrected message has a new key and may be sent after review.
    content = {
        "student": row.student_id or row.student_name,
        "phone": row.parent_phone,
        "date": row.date_text,
        "message": row.message.replace("\r\n", "\n").strip(),
        "mode": mode,
    }
    return hashlib.sha256(
        json.dumps(content, ensure_ascii=False, sort_keys=True).encode()
    ).hexdigest()


class DeliveryStore:
    def __init__(self, path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        with self.connection() as db:
            db.execute("""CREATE TABLE IF NOT EXISTS deliveries (
                fingerprint TEXT PRIMARY KEY, state TEXT NOT NULL,
                updated_at TEXT NOT NULL, batch_id TEXT NOT NULL DEFAULT ''
            )""")

    @contextmanager
    def connection(self):
        db = sqlite3.connect(self.path, timeout=10)
        try:
            yield db
            db.commit()
        except Exception:
            db.rollback()
            raise
        finally:
            db.close()

    def reserve(self, keys):
        reserved, blocked = [], []
        with self.connection() as db:
            db.execute("BEGIN IMMEDIATE")
            for key in keys:
                item = db.execute(
                    "SELECT state FROM deliveries WHERE fingerprint=?", (key,)
                ).fetchone()
                if item and item[0] in BLOCKING:
                    blocked.append(key)
                    continue
                db.execute(
                    """INSERT INTO deliveries VALUES (?, 'pending', ?, '')
                    ON CONFLICT(fingerprint) DO UPDATE SET
                    state='pending', updated_at=excluded.updated_at, batch_id=''""",
                    (key, datetime.now(timezone.utc).isoformat()),
                )
                reserved.append(key)
        return reserved, blocked

    def finish(self, keys, state, batch_id=""):
        if state not in BLOCKING | {"failed"}:
            raise ValueError("Unknown delivery state")
        with self.connection() as db:
            db.executemany(
                "UPDATE deliveries SET state=?,updated_at=?,batch_id=? WHERE fingerprint=?",
                [
                    (state, datetime.now(timezone.utc).isoformat(), str(batch_id), key)
                    for key in keys
                ],
            )

    def states(self, keys):
        # Kept small to avoid SQLite variable-count limits with large previews.
        with self.connection() as db:
            result = {}
            for key in set(keys):
                item = db.execute(
                    "SELECT state FROM deliveries WHERE fingerprint=?", (key,)
                ).fetchone()
                if item:
                    result[key] = item[0]
            return result
