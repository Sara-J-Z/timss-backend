import os
import openpyxl
import threading

from main_app.services.graph_upload_session import GraphUploadSessionClient
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

EXCEL_DIR = os.path.join(os.getcwd(), "excel_files")
os.makedirs(EXCEL_DIR, exist_ok=True)

def safe_name(name: str) -> str:
    bad = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    for ch in bad:
        name = name.replace(ch, '-')
    return name.strip()[:120] or "UnknownSchool"

def save_to_excel(data: dict):
    # ✅ 1) get values first
    school_name = data.get("school_name", "UnknownSchool")
    subject = data.get("subject", "UnknownSubject")

    # ✅ 2) safe school name for OneDrive folder/filename
    safe_school = safe_name(school_name)
    remote_folder = safe_school
    remote_filename = f"{safe_school}.xlsx"

    # (اختياري لكن أنصح به) استخدمي safe_school أيضًا محليًا لتجنب مشاكل أسماء الملفات
    file_path = os.path.join(EXCEL_DIR, f"{safe_school}.xlsx")

    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
    else:
        wb = openpyxl.Workbook()
        default_sheet = wb.active
        wb.remove(default_sheet)

    if subject in wb.sheetnames:
        ws = wb[subject]
    else:
        ws = wb.create_sheet(title=subject)
        headers = [
            "date", "time", "student_name", "class_name", "teacher_name",
            "school_operation_region", "auto_correct_score_points"
        ]
        question_headers = [ans["question_number"] for ans in data.get("answers", [])]
        ws.append(headers + question_headers)

        for col_num, _ in enumerate(ws[1], 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    row = [
        data.get("date"),
        data.get("time"),
        data.get("student_name"),
        data.get("class_name"),
        data.get("teacher_name"),
        data.get("school_operation_region"),
        data.get("auto_correct_score_points")
    ]
    question_values = [ans.get("answer_value") for ans in data.get("answers", [])]
    ws.append(row + question_values)

    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # تطبيق الحدود والتنسيق وZebra Stripes
    for idx, row_cells in enumerate(ws.iter_rows(), 1):
        if idx != 1:
            fill_color = "DCE6F1" if idx % 2 == 0 else "FFFFFF"
        else:
            fill_color = None

        for cell in row_cells:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if fill_color:
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

    # ضبط عرض جميع الأعمدة إلى 25
    for col in ws.columns:
        column_letter = col[0].column_letter
        ws.column_dimensions[column_letter].width = 25

    # ✅ save once
        wb.save(file_path)

    # ✅ Upload in background so Storyline gets response immediately
    def _upload():
        try:
            client = GraphUploadSessionClient()
            client.upload_large_file(
                local_path=file_path,
                remote_folder=remote_folder,
                remote_filename=remote_filename,
                chunk_size_mb=10
            )
        except Exception as e:
            print(f"[OneDrive Upload Error] {e}")

    threading.Thread(target=_upload, daemon=True).start()

    return file_path

