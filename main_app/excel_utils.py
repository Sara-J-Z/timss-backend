import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

from main_app.services.graph_excel_append import GraphExcelAppender, safe_name

EXCEL_DIR = os.path.join(os.getcwd(), "excel_files")
os.makedirs(EXCEL_DIR, exist_ok=True)


def save_to_excel(data: dict):
    school_name = data.get("school_name", "UnknownSchool")
    subject = data.get("subject", "UnknownSubject")

    # headers ثابتة + أسئلة
    headers = [
        "date", "time", "student_name", "class_name", "teacher_name",
        "school_operation_region", "auto_correct_score_points"
    ]
    question_headers = [ans.get("question_number") for ans in data.get("answers", [])]
    all_headers = headers + question_headers

    row_values = [
        data.get("date"),
        data.get("time"),
        data.get("student_name"),
        data.get("class_name"),
        data.get("teacher_name"),
        data.get("school_operation_region"),
        data.get("auto_correct_score_points"),
    ] + [ans.get("answer_value") for ans in data.get("answers", [])]

    # (اختياري) حفظ محلي بسيط مثل قبل — إذا تبينه
    safe_school = safe_name(school_name)
    file_path = os.path.join(EXCEL_DIR, f"{safe_school}.xlsx")

    # حفظ محلي (نفس منطقك السابق — اختصرته)
    if os.path.exists(file_path):
        wb = openpyxl.load_workbook(file_path)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

    if subject in wb.sheetnames:
        ws = wb[subject]
    else:
        ws = wb.create_sheet(title=subject)
        ws.append(all_headers)

        for col_num, _ in enumerate(ws[1], 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.append(row_values)

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    for row_cells in ws.iter_rows():
        for cell in row_cells:
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 25

    wb.save(file_path)

    # ✅ هنا المهم: تحديث OneDrive "بالـ append" بدل replace
    try:
        appender = GraphExcelAppender()
        appender.ensure_and_append(
            school_name=school_name,
            subject=subject,
            headers=all_headers,
            row_values=row_values
        )
        print("[Graph Excel] appended row successfully")
    except Exception as e:
        print(f"[Graph Excel Error] {e}")

    return file_path
