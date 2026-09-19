# -*- coding: utf-8 -*-
"""
TRÌNH HỖ TRỢ ĐĂNG NHẬP FACEBOOK (LƯU PHIÊN CHO AI AGENT)
Sử dụng Chromium độc lập để không xung đột ProcessSingleton với Chrome/Edge đang chạy.
"""

import os
import sys
import time
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_DIR = os.path.join(BASE_DIR, "fb_profile")

def login_facebook():
    print("=" * 65)
    print("   TRÌNH DUYỆT ĐĂNG NHẬP FACEBOOK (LƯU PHIÊN CHO AI AGENT)")
    print("=" * 65)
    print(f"[*] Thu muc luu phien: {PROFILE_DIR}")
    print("[*] Dang mo trinh duyet...")
    
    os.makedirs(PROFILE_DIR, exist_ok=True)
    
    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--no-first-run",
            "--no-default-browser-check",
            "--start-maximized"
        ]
        
        # Su dung Chromium doc lap de tranh xung dot lock file voi Chrome he thong
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=PROFILE_DIR,
                headless=False,
                args=args,
                no_viewport=True
            )
        except Exception as e:
            print(f"[!] Loi khoi dong: {e}")
            return
        
        page = context.pages[0] if context.pages else context.new_page()
        page.goto("https://www.facebook.com")
        
        print("\n" + "=" * 65)
        print("HUONG DAN DANG NHAP:")
        print("1. Trinh duyet Chromium da mo len tren man hinh.")
        print("2. Hay dang nhap tai khoan Facebook cua ban (mat khau, 2FA neu co).")
        print("3. Sau khi da vao duoc trang chu Bang tin Facebook thanh cong:")
        print("   -> Quay lai cua so den (CMD) nay va bam phim ENTER.")
        print("=" * 65 + "\n")
        
        try:
            input("--> Sau khi dang nhap xong, hay bam ENTER tai day de luu phien: ")
        except Exception:
            print("[*] Dang theo doi tien trinh dang nhap...")
            for _ in range(150):
                time.sleep(2)
                try:
                    cookies = [c['name'] for c in context.cookies()]
                    if 'c_user' in cookies:
                        print("[✓] Da phat hien dang nhap thanh cong!")
                        time.sleep(3)
                        break
                except Exception:
                    break
                    
        print("\n[*] Dang luu phien dang nhap va dong trinh duyet an toan...")
        time.sleep(2)
        try:
            context.close()
        except Exception:
            pass
        print("[✓] HOAN TAT! Da luu phien dang nhap Facebook vao thu muc 'fb_profile'.")
        print("=" * 65)

if __name__ == "__main__":
    login_facebook()
