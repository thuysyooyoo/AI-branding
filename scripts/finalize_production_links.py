# -*- coding: utf-8 -*-
import openpyxl, os, shutil
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.styles import Font

# 1. Copy CM4-CD19 to OneDrive youtube folder
src_cm4 = r"d:\AI AGENT THUY\AI branding\reports\ThuyTalk_CM4\CM4-CD19_Cach_Doc_Va_Nghien_Cuu_Luat_Khong_Bi_Chan.md"
dst_cm4 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\youtube\CM4-CD19_Cach_Doc_Va_Nghien_Cuu_Luat_Khong_Bi_Chan.md"
if os.path.exists(src_cm4) and not os.path.exists(dst_cm4):
    shutil.copy2(src_cm4, dst_cm4)
    print("Copied CM4-CD19 to OneDrive youtube directory successfully.")
else:
    print("CM4-CD19 already exists in OneDrive youtube directory or src missing.")

# 2. Update Weekly_Content_Calendar.xlsx
cal_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
wb_cal = openpyxl.load_workbook(cal_path)
ws2 = wb_cal["Facebook"]
ws3 = wb_cal["YouTube"]

blue_font = Font(name="Arial", size=10, color="000000FF", underline="single")

fb_links = {
    113: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/facebook/FB-PO-2026-W46-HOT47.md",
    114: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/facebook/FB-PO-2026-W46-HOT48.md",
    115: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/facebook/FB-PO-2026-W46-HOT49.md",
    116: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/facebook/FB-PO-2026-W46-HOT50.md",
}

yt_links = {
    80: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/video_scripts/YS-PO-2026-W46-HOT47.md",
    81: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/video_scripts/YS-PO-2026-W46-HOT48.md",
    82: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/video_scripts/YS-PO-2026-W46-HOT49.md",
    83: "file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/video_scripts/YS-PO-2026-W46-HOT50.md",
}

for r, url in fb_links.items():
    c = ws2.cell(r, 6)
    c.value = url
    c.hyperlink = url
    c.font = blue_font

for r, url in yt_links.items():
    c = ws3.cell(r, 6)
    c.value = url
    c.hyperlink = url
    c.font = blue_font

# Atomic save for Weekly_Content_Calendar.xlsx
tmp_cal = cal_path + ".tmp.xlsx"
wb_cal.save(tmp_cal)
wb_cal.close()
os.replace(tmp_cal, cal_path)
print("Updated Weekly_Content_Calendar.xlsx atomically.")

# 3. Update Ke_Hoach_Phat_Song_YouTube_Series.xlsx
p2 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx"
wb_yt = openpyxl.load_workbook(p2)
yt_dir = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\youtube"
yt_files = os.listdir(yt_dir) if os.path.exists(yt_dir) else []

updated_count = 0
for sname in ["CM1_VAN_BAN_NONG", "CM2_DICH_LUAT_TIENG_NGUOI", "CM3_LICH_SU_THUONG_MAI", "CM4_TAM_TINH_THUC_CHIEN"]:
    sheet = wb_yt[sname]
    for r in range(5, sheet.max_row + 1):
        code = sheet.cell(r, 2).value
        status = sheet.cell(r, 12).value
        if code and "ĐÃ HOÀN THÀNH" in str(status or ""):
            matches = [f for f in yt_files if str(code) in f]
            if matches:
                file_name = matches[0]
                full_path = os.path.join(yt_dir, file_name)
                formula = f'=HYPERLINK("{full_path}", "👉 Xem Kịch Bản Chi Tiết")'
                c = sheet.cell(r, 13)
                c.value = formula
                c.font = blue_font
                updated_count += 1

tmp_yt = p2 + ".tmp.xlsx"
wb_yt.save(tmp_yt)
wb_yt.close()
os.replace(tmp_yt, p2)
print(f"Updated Ke_Hoach_Phat_Song_YouTube_Series.xlsx with {updated_count} hyperlinks atomically.")
