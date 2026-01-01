import os
import re
import threading
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from main_app.services.graph_upload_session import GraphUploadSessionClient

EXCEL_DIR = os.path.join(os.getcwd(), "excel_files")
os.makedirs(EXCEL_DIR, exist_ok=True)

# Lock لكل مدرسة لمنع تضارب حفظ/رفع نفس الملف إذا وصل إرسالين بسرعة
_SCHOOL_LOCKS: dict[str, threading.Lock] = {}


def safe_name(name: str) -> str:
    bad = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    for ch in bad:
        name = name.replace(ch, '-')
    return name.strip()[:120] or "UnknownSchool"


def safe_sheet_name(name: str) -> str:
    """
    Excel worksheet rules:
    - max length 31
    - cannot contain: \ / ? * [ ]
    - avoid leading/trailing quotes/spaces
    """
    name = (name or "Sheet1").strip()
    name = re.sub(r"\s+", "_", name)          # spaces -> _
    name = re.sub(r"[\\/*?:\[\]]", "-", name) # forbidden -> -
    return name[:31] or "Sheet1"


def _get_lock_for_school(safe_school: str) -> threading.Lock:
    if safe_school not in _SCHOOL_LOCKS:
        _SCHOOL_LOCKS[safe_school] = threading.Lock()
    return _SCHOOL_LOCKS[safe_school]


def save_to_excel(data: dict):
    # 1) Values
    school_name = data.get("school_name", "UnknownSchool")
    subject_raw = data.get("subject", "UnknownSubject")
    subject = safe_sheet_name(subject_raw)

    # 2) Safe naming for file/folder
    safe_school = safe_name(school_name)
    remote_folder = safe_school
    remote_filename = f"{safe_school}.xlsx"
    file_path = os.path.join(EXCEL_DIR, remote_filename)

    lock = _get_lock_for_school(safe_school)

    with lock:
        # 3) Load or create workbook
        if os.path.exists(file_path):
            wb = openpyxl.load_workbook(file_path)
        else:
            wb = openpyxl.Workbook()
            default_sheet = wb.active
            wb.remove(default_sheet)

        # 4) Load or create worksheet
        if subject in wb.sheetnames:
            ws = wb[subject]
        else:
            ws = wb.create_sheet(title=subject)

            base_headers = [
                "date", "time", "student_name", "class_name", "teacher_name",
                "school_operation_region", "auto_correct_score_points"
            ]
            answers = data.get("answers", [])
            question_headers = [ans.get("question_number") for ans in answers]
            ws.append(base_headers + question_headers)

            # Header styling
            for col_num in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_num)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(
                    start_color="4F81BD",
                    end_color="4F81BD",
                    fill_type="solid"
                )
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # 5) Append row
        row = [
            data.get("date"),
            data.get("time"),
            data.get("student_name"),
            data.get("class_name"),
            data.get("teacher_name"),
            data.get("school_operation_region"),
            data.get("auto_correct_score_points"),
        ]
        question_values = [ans.get("answer_value") for ans in data.get("answers", [])]
        ws.append(row + question_values)

        # 6) Formatting (borders + zebra)
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for idx, row_cells in enumerate(ws.iter_rows(), 1):
            if idx != 1:
                fill_color = "DCE6F1" if idx % 2 == 0 else "FFFFFF"
            else:
                fill_color = None

            for cell in row_cells:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if fill_color:
                    cell.fill = PatternFill(
                        start_color=fill_color,
                        end_color=fill_color,
                        fill_type="solid"
                    )

        # 7) Column widths
        for col in ws.columns:
            column_letter = col[0].column_letter
            ws.column_dimensions[column_letter].width = 25

        # 8) Save once
        wb.save(file_path)

        # ✅ Ensure file is flushed to disk before upload
        try:
            with open(file_path, "rb") as f:
                os.fsync(f.fileno())
        except Exception:
            # إذا النظام ما يدعم fsync بهالطريقة، نتجاوز بدون كسر
            pass

    #     # 9) Upload to OneDrive (chunked / replace)
    #     try:
    #         client = GraphUploadSessionClient()
    #         client.upload_large_file(
    #             local_path=file_path,
    #             remote_folder=remote_folder,
    #             remote_filename=remote_filename,
    #             chunk_size_mb=10
    #         )
    #     except Exception as e:
    #         print(f"[OneDrive Upload Error] {e}")

    # return file_path
    
        # 9) Upload to OneDrive (chunked / replace)
    try:
        client = GraphUploadSessionClient()
        client.upload_large_file(
            local_path=file_path,
            remote_folder=remote_folder,
            remote_filename=remote_filename,
            chunk_size_mb=10,
            max_retries=1  # مهم: لا نكرر داخل request
        )
    except Exception as e:
        msg = str(e)

        # إذا الملف مقفول، لا نفشل الطلب ولا نحاول كثير
        if "423" in msg or "Locked" in msg:
            print("[OneDrive Upload] Skipped (Locked). Will upload on next submission.")
        else:
            print(f"[OneDrive Upload Error] {e}")

