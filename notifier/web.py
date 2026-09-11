"""HTTP forms and user messages. Business logic lives in the service modules."""

import hmac
import secrets
from dataclasses import replace
from functools import wraps

from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from .delivery_store import BLOCKING
from .input_rules import valid_phone
from .preview import ALIMTALK_MAX_LENGTH, apply_variable_edits, preview_stats
from .storage import PreviewReadError
from .templates import TEST_TYPES

web = Blueprint("web", __name__)


def current_service():
    return current_app.extensions["notifier"]


def home():
    return redirect(url_for("web.index"))


def serialized(view):
    @wraps(view)
    def run(*args, **kwargs):
        with current_service().lock:
            return view(*args, **kwargs)

    return run


def read_preview():
    try:
        return current_service().load_preview()
    except PreviewReadError as error:
        flash(str(error), "error")
        return []


def check_revision():
    token = request.form.get("preview_revision", "")
    if not hmac.compare_digest(token, current_service().files.preview_revision()):
        flash("미리보기가 변경되었습니다. 새로고침 후 학생과 내용을 다시 확인해주세요.", "error")
        return False
    return True


def checked_form_rows():
    if not check_revision():
        return None
    service = current_service()
    rows = read_preview()
    if not rows:
        flash("먼저 엑셀 미리보기를 생성해주세요.", "error")
        return None
    before = [row.message for row in rows]
    rows = apply_variable_edits(rows, request.form, service.settings)
    service.save_preview(rows)
    if before != [row.message for row in rows]:
        flash("수정된 내용을 미리보기에 반영했습니다. 확인 후 다시 발송해주세요.", "error")
        return None
    return rows


def selected_indices(length):
    try:
        values = {int(value) for value in request.form.getlist("selected")}
        if any(index < 0 or index >= length for index in values):
            raise ValueError
        return sorted(values)
    except (ValueError, TypeError):
        abort(400, "학생 선택 값이 올바르지 않습니다.")


def log_warning(result):
    if result.log_failed:
        flash("CSV 로그 저장에 실패했습니다. 중복 방지 이력은 유지됩니다.", "error")


@web.before_app_request
def protect_local_forms():
    if "csrf" not in session:
        session["csrf"] = secrets.token_urlsafe(32)
    if request.method == "POST":
        if not hmac.compare_digest(request.form.get("csrf_token", ""), session["csrf"]):
            abort(400, "화면을 새로고침한 후 다시 시도해주세요.")


@web.app_context_processor
def form_context():
    return {"csrf_token": session.get("csrf", "")}


@web.get("/")
def index():
    service = current_service()
    rows = read_preview()
    return render_template(
        "index.html",
        preview_rows=rows,
        stats=preview_stats(rows, service.settings),
        config=service.settings.public_config(),
        has_log=service.paths.log.exists(),
        max_length=ALIMTALK_MAX_LENGTH,
        test_status=service.receipt_status(),
        test_types=TEST_TYPES,
        preview_revision=service.files.preview_revision(),
        blocking_states=BLOCKING,
    )


@web.post("/load-demo")
@serialized
def load_demo():
    service = current_service()
    if not service.settings.is_demo:
        abort(400, "데모 모드에서만 샘플을 불러올 수 있습니다.")
    service.import_workbook(service.sample_path, service.sample_path)
    flash("가상 학생 샘플을 불러왔습니다. 실제 메시지는 전송되지 않습니다.", "success")
    return home()


@web.get("/sample")
def download_sample():
    return send_file(
        current_service().sample_path, as_attachment=True, download_name="sample_students.xlsx"
    )


@web.post("/check-api")
def check_api():
    service = current_service()
    if service.settings.is_demo:
        flash("데모 모드에서는 외부 API를 호출하지 않습니다.", "success")
        return home()
    try:
        service.settings.validate_auth()
        service.client.create_token(service.settings)
        flash(
            "API 인증 요청 성공. 템플릿 승인·본문 일치는 각 유형의 1건 수신으로 확인해주세요.",
            "success",
        )
    except Exception:
        flash("API 인증을 확인하지 못했습니다. 등록 IP와 .env 설정을 확인해주세요.", "error")
    return home()


@web.post("/preview")
@serialized
def preview():
    service = current_service()
    files = service.files
    files.clear_preview()
    try:
        combined = request.files.get("combined_file")
        daily = request.files.get("daily_report_file")
        default_students = service.paths.data / "student_db.xlsx"
        default_report = service.paths.data / "daily_report.xlsx"
        if combined and combined.filename:
            students_path = files.save_upload(combined, default_students)
            report_path = (
                files.save_upload(daily, default_report)
                if daily and daily.filename
                else students_path
            )
        else:
            students_path = files.save_upload(
                request.files.get("student_db_file"), default_students
            )
            report_path = files.save_upload(daily, default_report)
        if not students_path.exists():
            raise ValueError("학생DB가 포함된 엑셀 파일을 업로드해주세요.")
        if not report_path.exists():
            raise ValueError("데일리 리포트가 포함된 엑셀 파일을 업로드해주세요.")
        rows = service.import_workbook(students_path, report_path)
        stats = preview_stats(rows, service.settings)
        flash(
            f"미리보기 생성 완료: 총 {stats['total']}건 / 테스트 없음 {stats['none_count']}건 / "
            f"주간 {stats['weekly_count']}건 / 월간 {stats['monthly_count']}건 / 발송 가능 {stats['ready_count']}건",
            "success",
        )
    except Exception as error:
        files.clear_preview()
        flash(
            str(error)
            if isinstance(error, ValueError)
            else "파일 처리 중 오류가 발생했습니다. 엑셀 형식과 내용을 확인해주세요.",
            "error",
        )
    return home()


@web.post("/save-preview")
@serialized
def save_preview_edits():
    service = current_service()
    rows = read_preview()
    if not rows:
        flash("저장할 미리보기가 없습니다.", "error")
        return home()
    if not check_revision():
        return home()
    rows = apply_variable_edits(rows, request.form, service.settings)
    service.save_preview(rows)
    stats = preview_stats(rows, service.settings)
    flash(
        f"변수 저장 완료: 발송 가능 {stats['ready_count']}건 / 확인 필요 {stats['blocked_count']}건",
        "success",
    )
    return home()


@web.post("/clear-preview")
@serialized
def clear_preview_route():
    current_service().files.clear_preview()
    flash("기존 미리보기를 비웠습니다.", "success")
    return home()


@web.post("/reset-test-status")
@serialized
def reset_test_status():
    current_service().files.clear_receipts()
    flash("1건 실발송 확인 상태를 초기화했습니다.", "success")
    return home()


@web.post("/confirm-test")
@serialized
def confirm_test():
    service = current_service()
    kind = request.form.get("kind")
    if kind not in TEST_TYPES or service.settings.is_demo or service.settings.test_mode != "N":
        abort(400)
    try:
        service.confirm_receipt(kind, request.form.get("receipt_confirmed") == "yes")
        flash(kind + " 템플릿의 수신 확인을 기록했습니다.", "success")
    except ValueError as error:
        flash(str(error), "error")
    return home()


@web.post("/send-test")
@serialized
def send_test():
    service = current_service()
    rows = checked_form_rows()
    if rows is None:
        return home()
    targets = [rows[index] for index in selected_indices(len(rows)) if rows[index].is_ready]
    if len(targets) != 1:
        flash("발송 가능한 학생을 정확히 1명 선택해주세요.", "error")
        return home()
    if service.settings.is_demo or service.settings.test_mode != "N":
        flash("1건 수신 확인은 실제 발송 모드에서 진행합니다.", "error")
        return home()
    phone = service.settings.test_phone
    if not valid_phone(phone):
        flash(".env의 ALIGO_TEST_PHONE에 본인 휴대전화번호를 입력하고 재시작해주세요.", "error")
        return home()
    row = replace(targets[0], parent_phone=phone, parent_name="테스트 수신자")
    result = service.dispatch(
        [row], mode_override="receipt-test:" + service.receipt_signature(row.test_type)
    )
    log_warning(result)
    if result.accepted or result.previous_accepted:
        service.record_test_request(row.test_type)
        flash(
            "본인 번호로 1건 요청을 접수했습니다. 휴대전화 수신 확인 후 해당 유형의 확인 버튼을 누르세요.",
            "success",
        )
    return home()


@web.post("/send")
@serialized
def send():
    service = current_service()
    rows = checked_form_rows()
    if rows is None:
        return home()
    targets = (
        [row for row in rows if row.is_ready]
        if request.form.get("mode") == "all"
        else [rows[index] for index in selected_indices(len(rows)) if rows[index].is_ready]
    )
    if not targets:
        flash("발송 가능한 대상을 선택해주세요.", "error")
        return home()
    if not service.settings.is_demo and service.settings.test_mode == "N":
        needed = service.missing_receipts(targets)
        if needed:
            flash(
                "먼저 본인 번호 1건 수신 확인이 필요합니다: " + ", ".join(sorted(needed)), "error"
            )
            return home()
    result = service.dispatch(targets)
    log_warning(result)
    prefix = (
        "데모 처리"
        if service.settings.is_demo
        else "API 테스트 요청"
        if service.settings.delivery_mode == "api-test"
        else "발송 요청"
    )
    flash(
        f"{prefix}: 접수 {result.accepted}건 / 중복·확인대기 제외 {result.blocked}건 / 오류 {result.failed}건",
        "error" if result.failed else "success",
    )
    return home()


@web.get("/download-log")
def download_log():
    path = current_service().paths.log
    if not path.exists():
        flash("아직 발송 로그가 없습니다.", "error")
        return home()
    return send_file(path, as_attachment=True, download_name="alimtalk_send_log.csv")
