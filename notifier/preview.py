"""Student matching, input validation and message previews as pure functions."""

from collections import defaultdict
from dataclasses import asdict
from typing import Any

from .input_rules import (
    demo_phone,
    format_date_text,
    normalize_percentage,
    normalize_score,
    normalize_test_level,
    normalize_test_type,
    normalize_text,
    number,
    numeric_in_range,
    valid_date,
    valid_phone,
)
from .models import DailyReport, PreviewRow, StudentInfo
from .settings import Settings, is_placeholder
from .templates import (
    COMMON_REQUIRED_VARIABLES,
    TEMPLATE_TEXTS,
    TEST_MONTHLY,
    TEST_NONE,
    TEST_TYPES,
    TEST_WEEKLY,
)

ALIMTALK_MAX_LENGTH = 1000


def validate_row(row: PreviewRow, settings: Settings) -> PreviewRow:
    errors = [row.identity_error] if row.identity_error else []
    row.date_text = format_date_text(row.date_text)
    row.test_type = normalize_test_type(
        row.test_type, has_test_values=any((row.test_level, row.test_score, row.test_average))
    )
    if not valid_date(row.date_text):
        errors.append("날짜(YYYY-MM-DD)")
    if not row.student_name:
        errors.append("학생명")
    for attr, label in COMMON_REQUIRED_VARIABLES.items():
        if not normalize_text(getattr(row, attr)):
            errors.append(label)
    if not numeric_in_range(row.homework_percent, 0, 100, percent=True):
        errors.append("과제수행률 0~100%")
    for attr, label in (("word_total_count", "단어 전체"), ("word_wrong_count", "단어 오답")):
        if not numeric_in_range(getattr(row, attr), 0, integer=True):
            errors.append(label + " 0 이상 정수")
    total, wrong = number(row.word_total_count), number(row.word_wrong_count)
    if total is not None and wrong is not None and wrong > total:
        errors.append("단어 오답은 전체 수 이하")
    if row.test_type not in TEST_TYPES:
        errors.append("테스트구분 직접 선택(없음/주간/월간)")
    if row.test_type in {TEST_WEEKLY, TEST_MONTHLY}:
        if not numeric_in_range(row.test_level, 1, 7, integer=True):
            errors.append("테스트레벨 1~7 정수")
        for attr, label in (("test_score", "테스트점수"), ("test_average", "테스트평균")):
            if not numeric_in_range(getattr(row, attr), 0, 100):
                errors.append(label + " 0~100")
    if not valid_phone(row.parent_phone) and not (
        settings.is_demo and demo_phone(row.parent_phone)
    ):
        errors.append("학부모 휴대전화번호 형식")
    row.template_code = settings.template_code(row.test_type)
    row.template_label = {
        TEST_NONE: "테스트 없는 날",
        TEST_WEEKLY: "주간 테스트",
        TEST_MONTHLY: "월간 테스트",
    }.get(row.test_type, "유형 확인 필요")
    if is_placeholder(row.template_code):
        errors.append(row.template_label + " 템플릿 코드")
    elif not settings.template_codes_are_unique():
        errors.append("템플릿 코드 중복")
    template = TEMPLATE_TEXTS.get(row.test_type, TEMPLATE_TEXTS[TEST_NONE])
    row.message = template.format(**asdict(row), academy_name=settings.academy_name)
    row.message_length = len(row.message)
    if row.message_length > ALIMTALK_MAX_LENGTH:
        errors.append(f"본문 {ALIMTALK_MAX_LENGTH}자 이하")
    row.missing_fields = ", ".join(dict.fromkeys(errors))
    row.is_ready = not errors
    row.estimated_cost = settings.unit_price
    return row


def build_preview(
    student_map: dict[str, StudentInfo], reports: list[DailyReport], settings: Settings
) -> list[PreviewRow]:
    by_name = defaultdict(list)
    for student in student_map.values():
        by_name[student.student_name].append(student)
    rows, seen = [], set()
    for report in reports:
        error = ""
        if report.student_id:
            student = student_map.get("id:" + report.student_id)
            if not student:
                error = "학생DB에 없는 학생ID"
            elif student.student_name != report.student_name:
                error = "학생ID와 학생명이 일치하지 않음"
        else:
            matches = by_name.get(report.student_name, [])
            student = matches[0] if len(matches) == 1 else None
            if len(matches) > 1:
                error = "동명이인: 리포트에 학생ID 입력 필요"
            elif not matches:
                error = "학생DB에서 찾지 못한 학생"
        if student and student.send_yn != "Y":
            continue
        data = asdict(report)
        data["student_id"] = report.student_id or (student.student_id if student else "")
        data.update(
            parent_name=student.parent_name if student and not error else "",
            parent_phone=student.parent_phone if student and not error else "",
            identity_error=error,
        )
        row = validate_row(PreviewRow(**data), settings)
        identity = row.student_id or row.student_name
        record = (identity, row.date_text, row.test_type)
        if record in seen:
            row.identity_error = "동일 학생·날짜·테스트구분의 리포트 중복"
            validate_row(row, settings)
            # Neither duplicate is sendable: do not silently choose the first.
            for prior in rows:
                if (
                    prior.student_id or prior.student_name,
                    prior.date_text,
                    prior.test_type,
                ) == record:
                    prior.identity_error = row.identity_error
                    validate_row(prior, settings)
        seen.add(record)
        rows.append(row)
    return rows


def apply_variable_edits(rows: list[PreviewRow], form, settings: Settings) -> list[PreviewRow]:
    editable = [
        "attendance",
        "homework_percent",
        "missing_homework",
        "word_total_count",
        "word_wrong_count",
        "error_test_status",
        "test_type",
        "test_level",
        "test_score",
        "test_average",
    ]
    for idx, row in enumerate(rows):
        for attr in editable:
            key = f"{attr}_{idx}"
            if key not in form:
                continue
            value = normalize_text(form.get(key))
            if attr == "homework_percent":
                value = normalize_percentage(value)
            elif attr in {"word_total_count", "word_wrong_count", "test_score", "test_average"}:
                value = normalize_score(value)
            elif attr == "test_level":
                value = normalize_test_level(value)
            elif attr == "test_type":
                value = normalize_test_type(
                    value, has_test_values=any((row.test_level, row.test_score, row.test_average))
                )
            setattr(row, attr, value)
        row.missing_homework = row.missing_homework or "없음"
        validate_row(row, settings)
    return rows


def preview_stats(rows: list[PreviewRow], settings: Settings) -> dict[str, Any]:
    total = len(rows)
    ready_count = sum(1 for row in rows if row.is_ready)
    blocked_count = total - ready_count
    counts = {kind: sum(1 for row in rows if row.test_type == kind) for kind in TEST_TYPES}
    unit_price = settings.unit_price
    return {
        "total": total,
        "ready_count": ready_count,
        "blocked_count": blocked_count,
        "estimated_total": round(ready_count * unit_price, 1),
        "unit_price": unit_price,
        "none_count": counts[TEST_NONE],
        "weekly_count": counts[TEST_WEEKLY],
        "monthly_count": counts[TEST_MONTHLY],
    }
