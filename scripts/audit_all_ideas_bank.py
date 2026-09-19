# -*- coding: utf-8 -*-
import os
import sys
import io
import openpyxl

try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except AttributeError:
    pass

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from url_validator import validate_url_deep

WORKSPACE_DIR = r"d:\AI AGENT THUY/AI branding"
REPORTS_DIR = os.path.join(WORKSPACE_DIR, "reports")

files = [
    ("Tin_Nganh_Ideas_Bank_Eureka.xlsx", 3, 13),
    ("Case_Study_Ideas_Bank_Eureka.xlsx", 3, 11),
    ("Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx", 3, 11),
    ("Funfact_Ideas_Bank_Eureka.xlsx", 3, 10)
]

print("="*90)
print("🔍 TOÀN DIỆN KIỂM TOÁN TẤT CẢ CÁC LINK TRONG 4 TỆP EXCEL KHO Ý TƯỞNG")
print("="*90)

total_checked = 0
total_passed = 0
total_failed = 0

for fn, title_col, url_col in files:
    path = os.path.join(REPORTS_DIR, fn)
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    
    print(f"\n📂 TỆP: {fn} (Tổng dòng: {ws.max_row - 4})")
    print("-" * 90)
    
    for r in range(5, ws.max_row + 1):
        stt = ws.cell(row=r, column=1).value
        title = ws.cell(row=r, column=title_col).value
        url = ws.cell(row=r, column=url_col).value
        
        if not url or not str(url).startswith("http"):
            continue
            
        total_checked += 1
        res = validate_url_deep(url, topic_title=str(title))
        
        if res["is_valid"]:
            total_passed += 1
            print(f"🟢 [Dòng {r:02d} | STT {stt:02d}] PASSED (HTTP {res['status_code']})")
            print(f"   👉 Title: {str(title)[:50]}...")
            print(f"   👉 Real Web Title: \"{res['page_title']}\"")
            print(f"   👉 Matched Entities: {res['matched_keywords']}")
        else:
            total_failed += 1
            print(f"🔴 [Dòng {r:02d} | STT {stt:02d}] FAILED!")
            print(f"   👉 Title: {str(title)[:50]}...")
            print(f"   👉 URL: {url}")
            print(f"   👉 Lý do: {res['error_reason']}")
            if res.get("page_title"):
                print(f"   👉 Real Web Title: \"{res['page_title']}\"")
        print("-" * 90)

print("\n" + "="*90)
print(f"📊 TỔNG KẾT KIỂM TOÁN: {total_passed}/{total_checked} LINK HỢP LỆ (THẤT BẠI: {total_failed})")
print("="*90)
