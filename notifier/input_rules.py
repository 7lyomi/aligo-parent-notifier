"""Excel and form normalization. Invalid input is retained for an explicit error."""

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation


def normalize_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def number(value):
    if isinstance(value, bool):
        return None
    text = normalize_text(value)
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", text):
        return None
    try:
        result = Decimal(text)
        return result if result.is_finite() else None
    except InvalidOperation:
        return None


def number_text(value):
    return format(value.normalize(), "f") if value else "0"


def first_number(value):
    """Legacy labelled-cell reader only. Validators never use partial matching."""
    match = re.search(r"-?\d+(?:\.\d+)?", normalize_text(value))
    return match.group() if match else ""


def normalize_phone(value):
    text = normalize_text(value)
    if re.search(r"[^\d\s()+-]", text):
        return text
    digits = re.sub(r"\D", "", text)
    if text.startswith("+82"):
        digits = "0" + digits[2:]
    return digits


def valid_phone(value):
    return bool(re.fullmatch(r"010\d{8}|01[16789]\d{7,8}", normalize_text(value)))


def demo_phone(value):
    return bool(re.fullmatch(r"DEMO-\d{4}", normalize_text(value)))


def normalize_percentage(value, number_format=""):
    text = normalize_text(value)
    if not text:
        return ""
    numeric = number(text[:-1].strip() if text.endswith("%") else text)
    if numeric is None:
        return text
    # Only a numeric Excel cell with percentage formatting is a stored fraction.
    fmt = re.sub(r'"[^"]*"|\\.', "", number_format or "")
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool) and "%" in fmt:
        numeric *= 100
    return number_text(numeric) + "%"


def normalized_numeric(value, suffix=""):
    text = normalize_text(value)
    if suffix and text.endswith(suffix):
        text = text[: -len(suffix)].strip()
    numeric = number(text)
    return number_text(numeric) if numeric is not None else text


def normalize_count(value):
    return normalized_numeric(value, "개")


def normalize_score(value):
    return normalized_numeric(value, "점")


def normalize_test_level(value):
    text = re.sub(r"^(?:lv\.?|level)\s*", "", normalize_text(value), flags=re.I)
    return normalized_numeric(text)


def normalize_test_type(value, has_test_values=False):
    text = normalize_text(value).replace(" ", "").lower()
    if text in {"주간", "주간테스트", "weekly", "week", "w"}:
        return "주간"
    if text in {"월간", "월간테스트", "monthly", "month", "m"}:
        return "월간"
    if text in {"없음", "미실시", "안봄", "미응시", "none", "n", "x", "0"}:
        return "없음"
    if not text:
        # Scores without an explicit type must be reviewed, never assumed weekly.
        return "" if has_test_values else "없음"
    return normalize_text(value)


def format_date_text(value):
    if isinstance(value, (date, datetime)):
        return value.strftime("%Y-%m-%d")
    text = normalize_text(value)
    match = re.fullmatch(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})", text)
    if match:
        try:
            return date(*map(int, match.groups())).isoformat()
        except ValueError:
            pass
    return text


def valid_date(value):
    try:
        return date.fromisoformat(value).isoformat() == value
    except (ValueError, TypeError):
        return False


def numeric_in_range(value, minimum=0, maximum=None, integer=False, percent=False):
    text = normalize_text(value)
    if percent and text.endswith("%"):
        text = text[:-1]
    result = number(text)
    return (
        result is not None
        and result >= minimum
        and (maximum is None or result <= maximum)
        and (not integer or result == result.to_integral_value())
    )
