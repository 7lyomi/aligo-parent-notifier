"""Aligo transport: authentication, payloads and unambiguous error categories."""

import re
from uuid import uuid4

import requests

from .input_rules import demo_phone
from .preview import validate_row
from .settings import Settings, is_placeholder

ALIGO_TOKEN_URL = "https://kakaoapi.aligo.in/akv10/token/create/30/s/"
ALIGO_ALIMTALK_SEND_URL = "https://kakaoapi.aligo.in/akv10/alimtalk/send/"
MAX_RECIPIENTS = 500


class BeforeSendError(ValueError):
    """No message request was sent; retry can be allowed."""


class ApiRejected(ValueError):
    """The API explicitly rejected the whole request."""


class SendUncertain(RuntimeError):
    """A message request may have been accepted; do not retry automatically."""


class AligoClient:
    def create_token(self, settings: Settings):
        config = settings.aligo_config()
        missing = [
            label
            for key, label in (("apikey", "ALIGO_API_KEY"), ("userid", "ALIGO_USER_ID"))
            if is_placeholder(config[key])
        ]
        if missing:
            raise ValueError(".env에 다음 값을 설정해주세요: " + ", ".join(missing))
        response = requests.post(
            ALIGO_TOKEN_URL,
            data={"apikey": config["apikey"], "userid": config["userid"]},
            timeout=20,
        )
        response.raise_for_status()
        data = response.json()
        if str(data.get("code")) != "0":
            raise ValueError("알리고 토큰 생성 실패")
        token = data.get("token") or data.get("urlencode")
        if not token:
            raise ValueError("알리고 응답에 token이 없습니다.")
        return str(token)

    def send(self, rows, template_code: str, settings: Settings):
        if not rows or len(rows) > MAX_RECIPIENTS:
            raise BeforeSendError("발송 대상은 1~500명이어야 합니다.")
        if any(not validate_row(row, settings).is_ready for row in rows):
            raise BeforeSendError("입력 오류가 있는 행은 발송할 수 없습니다.")
        if settings.is_demo:
            return {
                "code": 0,
                "message": "데모 시뮬레이션",
                "info": {"mid": "DEMO-" + uuid4().hex[:12]},
            }
        if any(demo_phone(row.parent_phone) for row in rows):
            raise BeforeSendError("가상 번호는 실제 발송할 수 없습니다.")
        try:
            settings.validate_auth()
            if not settings.academy_name or settings.academy_name == "샘플학원":
                raise ValueError("승인 템플릿의 학원명이 필요합니다.")
            if settings.test_mode not in {"Y", "N"}:
                raise ValueError("ALIGO_TEST_MODE는 Y 또는 N이어야 합니다.")
            token = self.create_token(settings)
        except Exception:
            raise BeforeSendError(
                "발송 전 인증·설정 확인에 실패했습니다. API 설정과 승인 학원명을 확인해주세요."
            ) from None
        config = settings.aligo_config()
        payload = {
            "apikey": config["apikey"],
            "userid": config["userid"],
            "token": token,
            "senderkey": config["senderkey"],
            "tpl_code": template_code,
            "sender": config["sender"],
            "failover": "N",
            "testMode": config["test_mode"],
        }
        for index, row in enumerate(rows, 1):
            payload.update(
                {
                    f"receiver_{index}": row.parent_phone,
                    f"recvname_{index}": row.parent_name,
                    f"subject_{index}": settings.academy_name + " 학습 현황",
                    f"message_{index}": row.message,
                }
            )
        try:
            response = requests.post(ALIGO_ALIMTALK_SEND_URL, data=payload, timeout=40)
            response.raise_for_status()
            data = response.json()
            if not isinstance(data, dict):
                raise ValueError("Unexpected response")
        except Exception:
            raise SendUncertain(
                "발송 응답을 확인하지 못했습니다. 재발송 전에 알리고 발송 내역을 확인해주세요."
            ) from None
        result_code = data.get("code", data.get("result_code"))
        success_code = "0" if "code" in data else "1"
        if str(result_code) != success_code:
            if not re.fullmatch(r"-?\d+", str(result_code)):
                raise SendUncertain(
                    "발송 응답에 유효한 결과 코드가 없습니다. 알리고 발송 내역을 먼저 확인해주세요."
                )
            raise ApiRejected("알리고가 요청을 거절했습니다. 결과 코드: " + str(result_code))
        return data
