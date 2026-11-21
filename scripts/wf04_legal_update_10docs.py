# -*- coding: utf-8 -*-
"""
WORKFLOW 04: CẬP NHẬT 10 VĂN BẢN PHÁP LUẬT XNK MỚI (16/11/2026)
Hệ thống PBAN 2.0 - Eureka Logistics
Đồng bộ tự động vào: reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx
và OneDrive Marketing FB 2026.
"""

import os
import sys
import json
import subprocess

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
SCRATCH_JSON = os.path.join(BASE_DIR, "scratch", "wf04_10_new_legal_docs.json")
TRACKER_SCRIPT = os.path.join(BASE_DIR, "scripts", "legal_tracker.py")
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "Danh_sach_Van_ban_Phap_luat_XNK.xlsx")

def run_update():
    if not os.path.exists(SCRATCH_JSON):
        print(f"Error: {SCRATCH_JSON} not found!")
        return False
    
    cmd = [sys.executable, TRACKER_SCRIPT, "-i", SCRATCH_JSON, "-o", EXCEL_PATH]
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    print(res.stdout)
    if res.stderr:
        print("STDERR:", res.stderr)
    return res.returncode == 0

if __name__ == "__main__":
    run_update()
