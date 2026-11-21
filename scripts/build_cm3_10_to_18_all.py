# -*- coding: utf-8 -*-
"""
PIPELINE SẢN XUẤT PHÂN CẢNH YOUTUBE THUYTALK CM3 (CD10 ĐẾN CD18)
Tự động chuyển tiếp đến engine chuẩn hóa: scripts/update_all_15_storyboards.py
Đảm bảo 100% không trùng lặp, trực quan hóa số liệu, loại bỏ camera words.
"""

import os
import sys

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
sys.path.append(BASE_DIR)
sys.path.append(CURRENT_DIR)

from scripts.update_all_15_storyboards import main as update_all

if __name__ == "__main__":
    print("Executing master storyboard engine for ThuyTalk CM3...")
    update_all()
