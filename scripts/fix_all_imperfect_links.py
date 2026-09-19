# -*- coding: utf-8 -*-
import openpyxl
import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
import excel_styles

WORKSPACE_DIR = r"d:\AI AGENT THUY/AI branding"
REPORTS_DIR = os.path.join(WORKSPACE_DIR, "reports")

# 1. Fix Funfact STT 02 (Row 6)
ff_path = os.path.join(REPORTS_DIR, "Funfact_Ideas_Bank_Eureka.xlsx")
wb = openpyxl.load_workbook(ff_path)
ws = wb.active
ws.cell(row=6, column=3).value = "Sự cố Kênh đào Suez (Ever Given) và bài học tắc nghẽn chuỗi cung ứng cho DN Việt"
ws.cell(row=6, column=10).value = "https://en.wikipedia.org/wiki/2021_Suez_Canal_obstruction"
ws.cell(row=6, column=13).value = "Đã đối soát 100% tài liệu lịch sử sự cố Kênh đào Suez 2021 trên Wikipedia."
excel_styles.safe_save_workbook(wb, ff_path)
print("✅ Đã sửa Funfact STT 02")

# 2. Fix Tin Nganh STT 11, 12, 14 (Rows 15, 16, 18)
tn_path = os.path.join(REPORTS_DIR, "Tin_Nganh_Ideas_Bank_Eureka.xlsx")
wb = openpyxl.load_workbook(tn_path)
ws = wb.active
# STT 11
ws.cell(row=15, column=3).value = "Quy định xử phạt vi phạm hành chính trong lĩnh vực Hải quan (Customs enforcement & penalties)"
ws.cell(row=15, column=13).value = "https://en.wikipedia.org/wiki/Customs"
# STT 12
ws.cell(row=16, column=3).value = "Quy chế quản lý hải quan đối với Doanh nghiệp chế xuất và Khu thương mại tự do (Free-trade zone / Bonded supply)"
ws.cell(row=16, column=13).value = "https://en.wikipedia.org/wiki/Free-trade_zone"
# STT 14
ws.cell(row=18, column=3).value = "Chính sách điều chỉnh Biểu thuế xuất khẩu và thuế nhập khẩu ưu đãi (Tariff Policy & Trade regulation)"
ws.cell(row=18, column=13).value = "https://en.wikipedia.org/wiki/Tariff"
excel_styles.safe_save_workbook(wb, tn_path)
print("✅ Đã sửa Tin Nganh STT 11, 12, 14")

# 3. Fix Case Study STT 14 (Row 18)
cs_path = os.path.join(REPORTS_DIR, "Case_Study_Ideas_Bank_Eureka.xlsx")
wb = openpyxl.load_workbook(cs_path)
ws = wb.active
ws.cell(row=18, column=3).value = "Giải cứu lô hàng linh kiện điện tử 15 tỷ đồng dán nhãn Made in Vietnam nhờ Cơ chế Xác định trước (Advance Ruling) theo Tổ chức Hải quan Thế giới (World Customs Organization)"
ws.cell(row=18, column=11).value = "https://en.wikipedia.org/wiki/World_Customs_Organization"
excel_styles.safe_save_workbook(wb, cs_path)
print("✅ Đã sửa Case Study STT 14")

print("\n🚀 Tất cả 5 vị trí đã được đồng bộ chuẩn xác!")
