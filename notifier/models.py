"""Student, report and preview records shared across modules."""

from dataclasses import dataclass


@dataclass
class StudentInfo:
    student_name: str
    parent_name: str
    parent_phone: str
    send_yn: str = "Y"
    note: str = ""
    student_id: str = ""


@dataclass
class DailyReport:
    date_text: str
    student_name: str
    attendance: str
    homework_percent: str
    word_total_count: str
    word_wrong_count: str
    error_test_status: str
    missing_homework: str
    test_type: str
    test_level: str
    test_score: str
    test_average: str
    special_note: str = ""
    student_id: str = ""


@dataclass
class PreviewRow:
    student_name: str
    parent_name: str
    parent_phone: str
    date_text: str
    attendance: str
    homework_percent: str
    missing_homework: str
    word_total_count: str
    word_wrong_count: str
    error_test_status: str
    test_type: str
    test_level: str
    test_score: str
    test_average: str
    special_note: str
    template_code: str = ""
    template_label: str = ""
    message: str = ""
    send_yn: bool = True
    is_ready: bool = False
    missing_fields: str = ""
    message_length: int = 0
    estimated_cost: float = 6.5
    student_id: str = ""
    identity_error: str = ""
    delivery_state: str = ""
