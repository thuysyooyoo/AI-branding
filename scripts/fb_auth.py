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
        print("2. Chi Thuy vui long dang nhap tai khoan Facebook (nhap mat khau / 2FA).")
        print("3. He thong TU DONG THEO DOI va se tu dong luu phien ngay khi vao duoc Facebook!")
        print("   (Hoac co the quay lai cua so nay bam ENTER bat ky luc nao).")
        print("=" * 65 + "\n")
        
        import msvcrt
        print("[*] Dang cho Chi Thuy dang nhap tren trinh duyet...")
        max_wait_seconds = 600
        start_time = time.time()
        logged_in = False

        while time.time() - start_time < max_wait_seconds:
            # Kiem tra neu nguoi dung bam phim ENTER hoac phim bat ky
            if msvcrt.kbhit():
                key = msvcrt.getch()
                print("\n[✓] Nguoi dung da xac nhan hoan tat tren ban phim!")
                logged_in = True
                break

            # Tu dong quet cookie c_user
            try:
                cookies = [c['name'] for c in context.cookies()]
                if 'c_user' in cookies:
                    print("\n[✓] XAC NHAN: DA PHAT HIEN DANG NHAP THANH CONG (cookie c_user)!")
                    logged_in = True
                    time.sleep(2)
                    break
            except Exception:
                pass

            time.sleep(1.5)

        if not logged_in:
            print("[!] Het thoi gian cho dang nhap (10 phut).")
                    
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
