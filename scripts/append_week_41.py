# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
import os
import shutil
import json
from datetime import datetime

file_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
json_path = r"d:\AI AGENT THUY\AI branding\scratch\w41_data.json"

with open(json_path, "r", encoding="utf-8") as f:
    w41_data = json.load(f)

wb = openpyxl.load_workbook(file_path)
ws1 = wb["Weekly Calendar"]

font_family = "Segoe UI"
bold_font = Font(name=font_family, size=10, bold=True)
regular_font = Font(name=font_family, size=10)

zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

cat_fills = {
    "CS": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "TN": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),
    "HT": PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid"),
    "FF": PatternFill(start_color="F5D6E0", end_color="F5D6E0", fill_type="solid"),
}

platform_fills = {
    "Facebook": PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid"),
    "YouTube": PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
}

center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)

# Update Sheet 1 Title
ws1["A1"] = "LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W41 CỘNG DỒN)"

# Check existing rows
current_max = ws1.max_row
existing_stts = [ws1.cell(row=r, column=1).value for r in range(5, current_max + 1)]

if 57 not in existing_stts:
    for idx, item in enumerate(w41_data):
        r_idx = current_max + 1 + idx
        fill = zebra_fill if r_idx % 2 == 0 else white_fill
        row_vals = [
            item["stt"], item["day"], item["platform"], item["category"],
            item["technique"], item["format"], item["opt_a"], item["opt_b"],
            item["opt_c"], item["source"], item["status"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.fill = fill
            cell.alignment = center_align if c_idx in [1, 2, 3, 11] else left_align
            
            if c_idx == 3:
                plat_key = "Facebook" if "Facebook" in val else "YouTube"
                cell.fill = platform_fills[plat_key]
                cell.font = bold_font
            elif c_idx == 4:
                cat_key = val[:2]
                if cat_key in cat_fills:
                    cell.fill = cat_fills[cat_key]
                cell.font = bold_font
            else:
                cell.font = regular_font

    # Re-apply Data Validation
    dv = DataValidation(
        type="list", 
        formula1='"Chờ duyệt,Duyệt Option A,Duyệt Option B,Duyệt Option C,Không duyệt"', 
        allow_blank=True
    )
    ws1.add_data_validation(dv)
    dv.add(f"K5:K{ws1.max_row}")

    tmp_path = file_path + ".tmp.xlsx"
    wb.save(tmp_path)
    os.replace(tmp_path, file_path)
    print(f"SUCCESS: Appended W41! New max row: {ws1.max_row}")
else:
    print("W41 already exists in Sheet 1!")
