import hashlib
import os
import re
import threading
import time
import uuid
from contextlib import contextmanager

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from main_app.services.graph_upload_session import GraphUploadSessionClient


EXCEL_DIR = os.path.join(os.getcwd(), "excel_files")
LOCK_DIR = os.path.join(EXCEL_DIR, ".locks")
os.makedirs(EXCEL_DIR, exist_ok=True)
os.makedirs(LOCK_DIR, exist_ok=True)

# In-process lock per school.
_SCHOOL_LOCKS: dict[str, threading.Lock] = {}

# OneDrive/SharePoint-invalid path characters.
_INVALID_PATH_CHARS = set('\\/:*?"<>|')


BASE_HEADERS = [
    "date", "time", "student_name", "gender", "grade", "user_role",
    "class_name", "teacher_name", "school_operation_region",
    "auto_correct_score_points",
]


def safe_name(name: str) -> str:
    """Validate a school name before using it as a local/remote path segment."""
    value = str(name or "").strip()

    if not value or value in {".", ".."} or ".." in value:
        raise ValueError(f"Unsafe school_name: {name!r}")

    if any(ch in value for ch in _INVALID_PATH_CHARS):
        raise ValueError(f"Invalid character in school_name: {name!r}")

    if value.endswith(".") or value.endswith(" "):
        raise ValueError(f"Unsafe school_name ending: {name!r}")

    return value[:120]


def safe_sheet_name(name: str) -> str:
    r"""
    Excel sheet rules:
    - max length 31
    - cannot contain: \ / ? * : [ ]
    """
    name = (name or "Sheet1").strip()
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"[\\/*?:\[\]]", "-", name)
    return name[:31] or "Sheet1"


def _get_lock_for_school(safe_school: str) -> threading.Lock:
    if safe_school not in _SCHOOL_LOCKS:
        _SCHOOL_LOCKS[safe_school] = threading.Lock()
    return _SCHOOL_LOCKS[safe_school]


@contextmanager
def _cross_process_school_lock(
    safe_school: str,
    timeout_seconds: int = 300,
    stale_after_seconds: int = 600,
):
    """
    Coordinate same-school writes across Gunicorn worker processes on the same
    Render instance using an atomic lock file.
    """
    digest = hashlib.sha256(safe_school.encode("utf-8")).hexdigest()
    lock_path = os.path.join(LOCK_DIR, f"{digest}.lock")
    deadline = time.monotonic() + timeout_seconds
    fd = None

    while True:
        try:
            fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, f"pid={os.getpid()} created={time.time()}".encode("utf-8"))
            break
        except FileExistsError:
            # Clean up a lock left by a crashed process after a generous timeout.
            try:
                age = time.time() - os.path.getmtime(lock_path)
                if age > stale_after_seconds:
                    os.remove(lock_path)
                    continue
            except FileNotFoundError:
                continue

            if time.monotonic() >= deadline:
                raise TimeoutError(
                    f"Timed out waiting for Excel lock for school: {safe_school}"
                )
            time.sleep(0.2)

    try:
        yield
    finally:
        if fd is not None:
            try:
                os.close(fd)
            except OSError:
                pass
        try:
            os.remove(lock_path)
        except FileNotFoundError:
            pass


def _style_header_cell(cell):
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        start_color="4F81BD", end_color="4F81BD", fill_type="solid"
    )
    cell.alignment = Alignment(horizontal="center", vertical="center")


def _style_data_row(ws, row_number: int):
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    fill_color = "DCE6F1" if row_number % 2 == 0 else "FFFFFF"

    for cell in ws[row_number]:
        cell.border = thin_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(
            start_color=fill_color, end_color=fill_color, fill_type="solid"
        )


def _ensure_headers(ws, answers: list[dict]) -> list[str]:
    """Return worksheet headers and append any newly seen question columns safely."""
    if ws.max_row < 1 or ws.cell(row=1, column=1).value is None:
        current_headers = []
    else:
        current_headers = [
            ws.cell(row=1, column=col).value
            for col in range(1, ws.max_column + 1)
        ]

    if not current_headers:
        current_headers = list(BASE_HEADERS)
        for col_num, header in enumerate(current_headers, 1):
            ws.cell(row=1, column=col_num, value=header)

    # Preserve incoming question order, but never duplicate a header.
    for ans in answers:
        q = ans.get("question_number")
        if q and q not in current_headers:
            current_headers.append(q)
            ws.cell(row=1, column=len(current_headers), value=q)

    for col_num, header in enumerate(current_headers, 1):
        cell = ws.cell(row=1, column=col_num)
        if cell.value is None:
            cell.value = header
        _style_header_cell(cell)
        ws.column_dimensions[cell.column_letter].width = 25

    return current_headers


def save_to_excel(data: dict):
    school_name = data.get("school_name", "")
    subject_raw = data.get("subject", "UnknownSubject")
    subject = safe_sheet_name(subject_raw)
    safe_school = safe_name(school_name)

    remote_folder = safe_school
    remote_filename = f"{safe_school}.xlsx"

    thread_lock = _get_lock_for_school(safe_school)

    with thread_lock:
        with _cross_process_school_lock(safe_school):
            client = GraphUploadSessionClient()

            # Use a unique per-request temp file so worker processes cannot share
            # or accidentally reuse stale workbook state.
            temp_filename = f"{hashlib.sha256(safe_school.encode('utf-8')).hexdigest()}-{uuid.uuid4().hex}.xlsx"
            file_path = os.path.join(EXCEL_DIR, temp_filename)
            wb = None

            try:
                # SAFETY RULE:
                # Always fetch the current OneDrive workbook first.
                # Only a confirmed 404 is allowed to create a new workbook.
                downloaded = client.download_file(
                    remote_folder, remote_filename, file_path
                )

                if downloaded:
                    print("[OneDrive] Downloaded current workbook before updating.")
                    wb = openpyxl.load_workbook(file_path)
                else:
                    print("[OneDrive] Workbook not found (404). Creating a new workbook.")
                    wb = openpyxl.Workbook()
                    default_sheet = wb.active
                    wb.remove(default_sheet)

                # Load or create subject sheet.
                if subject in wb.sheetnames:
                    ws = wb[subject]
                else:
                    ws = wb.create_sheet(title=subject)

                answers = data.get("answers", []) or []
                headers = _ensure_headers(ws, answers)

                base_values = {
                    "date": data.get("date"),
                    "time": data.get("time"),
                    "student_name": data.get("student_name"),
                    "gender": data.get("gender"),
                    "grade": data.get("grade"),
                    "user_role": data.get("user_role"),
                    "class_name": data.get("class_name"),
                    "teacher_name": data.get("teacher_name"),
                    "school_operation_region": data.get("school_operation_region"),
                    "auto_correct_score_points": data.get("auto_correct_score_points"),
                }
                answer_values = {
                    ans.get("question_number"): ans.get("answer_value")
                    for ans in answers
                    if ans.get("question_number")
                }

                # Append by header name, not by incoming answer-array position.
                row_values = [
                    base_values.get(header, answer_values.get(header))
                    for header in headers
                ]
                ws.append(row_values)
                _style_data_row(ws, ws.max_row)

                wb.save(file_path)
                wb.close()
                wb = None

                # Ensure file is flushed to disk before Graph upload.
                with open(file_path, "rb") as f:
                    os.fsync(f.fileno())

                # Retry short-lived 409/423/429/503 failures. Any final failure is
                # raised to the API; we never keep a stale local workbook for reuse.
                result = client.upload_large_file(
                    local_path=file_path,
                    remote_folder=remote_folder,
                    remote_filename=remote_filename,
                    chunk_size_mb=10,
                    max_retries=3,
                )

                print("[OneDrive Upload] Success.")
                return result

            except Exception as e:
                print(f"[Excel/OneDrive Error] {e}")
                raise

            finally:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

                # Never leave a workbook behind for the next request to reuse.
                try:
                    if os.path.exists(file_path):
                        os.remove(file_path)
                        print(f"[CLEANUP] Temp file deleted: {file_path}")
                except Exception as e:
                    print(f"[CLEANUP WARNING] Could not delete temp file: {e}")
