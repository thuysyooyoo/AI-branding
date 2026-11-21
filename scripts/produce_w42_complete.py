# -*- coding: utf-8 -*-
"""
Script: produce_w42_complete.py
Sản xuất toàn bộ nội dung Giai đoạn 2 cho Tuần W42 (STT 64 - 70):
- 7 PO vật lý (scratch/production_orders/PO-2026-W42-XX.md)
- 4 Bài viết Facebook chi tiết (scratch/production_orders/DRAFT-FB-PO-2026-W42-XX.md & reports/plan1/facebook/FB-2026-W42-XX.md)
- 3 Kịch bản YouTube dài + Storyboard 4 cột + Cliffhanger Outro (scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W42-XX.md & reports/plan1/video_scripts/YS-2026-W42-XX.md)
- 7 Báo cáo Fact-Check chuẩn Zero-Trust SSoT (scratch/production_orders/FACT-CHECK-PO-2026-W42-XX.md)
- Cập nhật trực tiếp Weekly_Content_Calendar.xlsx (Sheet 1, Sheet 2, Sheet 3) với atomic_save (Zero Backup)
- Cập nhật weekly_calendar.md và content_status_report.md
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")

os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_FB_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)

print("Starting production for Week 42 (STT 64 - 70)...")
