# -*- coding: utf-8 -*-
"""
TỰ ĐỘNG ĐỒNG BỘ FILE EXCEL YOUTUBE SERIES KHI CHỊ THÚY ĐÓNG EXCEL
"""
import os
import shutil
import time

SRC = r"d:\AI AGENT THUY\AI branding\scratch\temp_series_copy.xlsx"
DST = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx"

def try_sync():
    if not os.path.exists(SRC):
        print("Không tìm thấy file nguồn tạm:", SRC)
        return False
    try:
        with open(DST, "r+b") as f:
            pass
        # Atomic replace
        tmp_dst = DST + ".tmp.xlsx"
        shutil.copy2(SRC, tmp_dst)
        os.replace(tmp_dst, DST)
        print("🟢 ĐÃ ĐỒNG BỘ THÀNH CÔNG VÀO FILE MASTER TRÊN ONEDRIVE!")
        return True
    except PermissionError:
        print("⚠️ File vẫn đang mở trong Excel.")
        return False
    except Exception as e:
        print("Lỗi:", e)
        return False

if __name__ == "__main__":
    try_sync()
