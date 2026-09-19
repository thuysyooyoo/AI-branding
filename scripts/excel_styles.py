# -*- coding: utf-8 -*-
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

FONT_FAMILY = "Segoe UI"

# Common Fonts
TITLE_FONT = Font(name=FONT_FAMILY, size=14, bold=True, color="1F4E78")
HEADER_FONT = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name=FONT_FAMILY, size=10)
BOLD_BODY_FONT = Font(name=FONT_FAMILY, size=10, bold=True)
NOTE_FONT = Font(name=FONT_FAMILY, size=9, italic=True, color="595959")

# Common Fills
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")  # Navy Blue
ZEBRA_FILL = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")   # Light slate gray

# Common Borders
THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# Alignments
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)

def apply_header_style(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = ALIGN_CENTER
    cell.border = THIN_BORDER

def apply_body_style(cell, is_zebra=False, align_left=True, bold=False):
    cell.font = BOLD_BODY_FONT if bold else BODY_FONT
    cell.border = THIN_BORDER
    if is_zebra:
        cell.fill = ZEBRA_FILL
    cell.alignment = ALIGN_LEFT if align_left else ALIGN_CENTER

def safe_save_workbook(wb, path):
    """Lưu workbook an toàn. Nếu file đang bị Excel khóa, tự động đóng Excel và ghi đè."""
    import os
    import time
    try:
        wb.save(path)
        return True
    except PermissionError:
        print(f"⚠️ File {path} đang bị khóa bởi Excel. Đang tự động đóng tiến trình Excel...")
        os.system("taskkill /f /im excel.exe >nul 2>&1")
        time.sleep(1)
        try:
            wb.save(path)
            print(f"✅ Đã đóng Excel và lưu thành công tệp: {path}")
            return True
        except Exception as e:
            print(f"❌ Lỗi lưu tệp sau khi đóng Excel: {e}")
            return False
