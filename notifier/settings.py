"""Load configuration once per app instance; never expose credentials to the UI."""

import math
import os
import re
from dataclasses import dataclass
from typing import Mapping

from .input_rules import normalize_phone, normalize_text
from .templates import TEST_MONTHLY, TEST_NONE, TEST_TYPES, TEST_WEEKLY

DEFAULT_UNIT_PRICE = 6.5


def is_placeholder(value: str) -> bool:
    text = normalize_text(value)
    return not text or text.startswith("여기에_") or "승인후" in text or "승인_후" in text


@dataclass(frozen=True)
class Settings:
    mode: str = "demo"
    school_name: str = ""
    api_key: str = ""
    user_id: str = ""
    sender: str = ""
    sender_key: str = ""
    template_none: str = ""
    template_weekly: str = ""
    template_monthly: str = ""
    test_phone: str = ""
    test_mode: str = "Y"
    unit_price: float = DEFAULT_UNIT_PRICE

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None):
        env = os.environ if env is None else env
        try:
            price = float(env.get("ALIMTALK_UNIT_PRICE", str(DEFAULT_UNIT_PRICE)))
            if not math.isfinite(price) or price < 0:
                price = DEFAULT_UNIT_PRICE
        except ValueError:
            price = DEFAULT_UNIT_PRICE
        return cls(
            mode=env.get("APP_MODE", "demo").strip().lower(),
            school_name=env.get("ACADEMY_NAME", "").strip(),
            api_key=env.get("ALIGO_API_KEY", "").strip(),
            user_id=env.get("ALIGO_USER_ID", "").strip(),
            sender=normalize_phone(env.get("ALIGO_SENDER", "")),
            sender_key=env.get("ALIGO_SENDER_KEY", "").strip(),
            template_none=env.get("ALIGO_TEMPLATE_CODE_NONE", "").strip(),
            template_weekly=(
                env.get("ALIGO_TEMPLATE_CODE_WEEKLY", "") or env.get("ALIGO_TEMPLATE_CODE", "")
            ).strip(),
            template_monthly=env.get("ALIGO_TEMPLATE_CODE_MONTHLY", "").strip(),
            test_phone=normalize_phone(env.get("ALIGO_TEST_PHONE", "")),
            test_mode=env.get("ALIGO_TEST_MODE", "Y").strip().upper(),
            unit_price=price,
        )

    @property
    def is_demo(self):
        return self.mode == "demo"

    @property
    def academy_name(self):
        return "샘플학원" if self.is_demo else self.school_name

    @property
    def delivery_mode(self):
        return "demo" if self.is_demo else "api-test" if self.test_mode == "Y" else "live"

    def template_code(self, kind):
        if self.is_demo:
            return {
                TEST_NONE: "DEMO_NONE",
                TEST_WEEKLY: "DEMO_WEEKLY",
                TEST_MONTHLY: "DEMO_MONTHLY",
            }.get(kind, "")
        if kind == TEST_NONE:
            return self.template_none
        if kind == TEST_MONTHLY:
            return self.template_monthly
        return self.template_weekly

    def template_codes_are_unique(self):
        codes = [self.template_code(kind) for kind in TEST_TYPES]
        configured = [code for code in codes if not is_placeholder(code)]
        return len(configured) == len(set(configured))

    def aligo_config(self):
        return {
            "apikey": self.api_key,
            "userid": self.user_id,
            "sender": self.sender,
            "senderkey": self.sender_key,
            "test_mode": self.test_mode,
        }

    def public_config(self):
        return {
            "key_set": "미설정" if is_placeholder(self.api_key) else "설정됨",
            "sender_key_set": "미설정" if is_placeholder(self.sender_key) else "설정됨",
            "template_none": self.template_code(TEST_NONE) or "미설정",
            "template_weekly": self.template_code(TEST_WEEKLY) or "미설정",
            "template_monthly": self.template_code(TEST_MONTHLY) or "미설정",
            "test_mode": self.test_mode,
            "demo": self.is_demo,
            "academy_name": self.academy_name,
        }

    def validate_auth(self):
        fields = {
            "ALIGO_API_KEY": self.api_key,
            "ALIGO_USER_ID": self.user_id,
            "ALIGO_SENDER": self.sender,
            "ALIGO_SENDER_KEY": self.sender_key,
        }
        missing = [name for name, value in fields.items() if is_placeholder(value)]
        if self.mode not in {"demo", "live"}:
            missing.append("APP_MODE(demo/live)")
        if not re.fullmatch(r"\d{8,11}", self.sender):
            missing.append("ALIGO_SENDER 등록 발신번호")
        if missing:
            raise ValueError(".env에 다음 값을 설정해주세요: " + ", ".join(missing))
