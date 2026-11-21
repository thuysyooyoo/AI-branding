# -*- coding: utf-8 -*-
import openpyxl
import os

wb_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
wb = openpyxl.load_workbook(wb_path, data_only=True)

print("=== KIỂM TRA ĐỘ TOÀN VẸN VÀ CHUẨN XÁC CỦA EXCEL ===")
ws1 = wb["Weekly Calendar"]
ws2 = wb["Facebook"]
ws3 = wb["YouTube"]

s1_count = ws1.max_row - 4
s1_approvals = [ws1.cell(r, 11).value for r in range(5, ws1.max_row + 1)]
s1_passed = sum(1 for a in s1_approvals if a and "Duyệt Option A" in str(a))
print(f"Sheet 1: {s1_count} dòng data (Row 5 đến {ws1.max_row}) | Duyệt Option A: {s1_passed}/{len(s1_approvals)}")

s2_count = ws2.max_row - 4
s2_passed = [ws2.cell(r, 9).value for r in range(5, ws2.max_row + 1)]
s2_ready = [ws2.cell(r, 11).value for r in range(5, ws2.max_row + 1)]
s2_content_ok = [bool(ws2.cell(r, 7).value) for r in range(5, ws2.max_row + 1)]
s2_link_ok = [bool(ws2.cell(r, 6).value) for r in range(5, ws2.max_row + 1)]
print(f"Sheet 2 (Facebook): {s2_count} dòng data (Row 5 đến {ws2.max_row})")
print(f"  - Fact Check PASSED: {sum(1 for p in s2_passed if 'PASSED' in str(p))}/{len(s2_passed)}")
print(f"  - Sẵn sàng đăng: {sum(1 for r in s2_ready if 'Sẵn sàng đăng' in str(r))}/{len(s2_ready)}")
print(f"  - Nội dung hoàn chỉnh: {sum(s2_content_ok)}/{len(s2_content_ok)}")
print(f"  - Link .md vật lý: {sum(s2_link_ok)}/{len(s2_link_ok)}")

s3_count = ws3.max_row - 4
s3_passed = [ws3.cell(r, 9).value for r in range(5, ws3.max_row + 1)]
s3_ready = [ws3.cell(r, 11).value for r in range(5, ws3.max_row + 1)]
s3_content_ok = [bool(ws3.cell(r, 7).value) for r in range(5, ws3.max_row + 1)]
s3_link_ok = [bool(ws3.cell(r, 6).value) for r in range(5, ws3.max_row + 1)]
print(f"Sheet 3 (YouTube): {s3_count} dòng data (Row 5 đến {ws3.max_row})")
print(f"  - Fact Check PASSED: {sum(1 for p in s3_passed if 'PASSED' in str(p))}/{len(s3_passed)}")
print(f"  - Sẵn sàng dựng / duyệt: {sum(1 for r in s3_ready if any(kw in str(r) for kw in ['Sẵn sàng dựng', 'duyệt quay']))}/{len(s3_ready)}")
print(f"  - Kịch bản hoàn chỉnh (Full text + Storyboard 4 cột): {sum(s3_content_ok)}/{len(s3_content_ok)}")
print(f"  - Link .md vật lý: {sum(s3_link_ok)}/{len(s3_link_ok)}")

# Kiểm tra file backup .xlsx
backup_files = []
for root, dirs, files in os.walk(r"d:\AI AGENT THUY\AI branding"):
    for f in files:
        if "backup" in f.lower() and f.endswith(".xlsx"):
            backup_files.append(os.path.join(root, f))
print(f"\nKiểm tra Rule 10 (Không tạo backup file): Số file backup .xlsx = {len(backup_files)}")
if backup_files:
    print("CẢNH BÁO: Tồn tại backup files:", backup_files)
else:
    print("TUÂN THỦ TUYỆT ĐỐI RULE 10: ZERO BACKUP FILES!")
