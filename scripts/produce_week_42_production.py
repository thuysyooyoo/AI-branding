# -*- coding: utf-8 -*-
"""
Script: produce_week_42_production.py
Mục tiêu: Thực hiện Giai đoạn 2 (Sản xuất, Thẩm định & Duyệt đăng) cho Tuần 42 (STT 64 - 70):
- Agent 04: Đóng gói 7 PO vật lý tại scratch/production_orders/PO-2026-W42-XX.md
- Agent 05: Soạn 4 bài viết Facebook chuẩn PBAN 2.0 (STT 64, 66, 68, 70)
- Agent 06: Soạn 3 kịch bản YouTube chuyên sâu + Storyboard 4 cột + Cliffhanger Outro (STT 65, 67, 69)
- Agent 08: Thẩm định Fact-Check SSoT Zero-Trust 100% (7 file FACT-CHECK-PO-2026-W42-XX.md)
- Cập nhật trực tiếp vào Weekly_Content_Calendar.xlsx (Sheet 1, Sheet 2, Sheet 3) với atomic_save (Zero Backup)
- Cập nhật weekly_calendar.md và content_status_report.md
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")

os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_FB_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)

print("Directories prepared successfully.")
