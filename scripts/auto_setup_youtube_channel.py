# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG CẤU HÌNH & QUẢN TRỊ KÊNH YOUTUBE STUDIO CHO SẾP THÚY (THUYTALK)
Tự động đồng bộ 100% nhận diện từ kênh https://www.youtube.com/@thuystalk:
1. Tên kênh: THÚY TALK
2. Mô tả kênh: Chuẩn nguyên văn từ @thuystalk
3. Avatar & Banner: Tự động tải lên từ tài nguyên trích xuất
4. Handle: Tự động điền @thuystalk (hoặc fallback thông minh nếu trùng)
5. Lưu vĩnh viễn phiên đăng nhập vào yt_profile để tự động upload video các lần sau
"""

import os
import sys
import time
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_DIR = os.path.join(BASE_DIR, "yt_profile")
ASSETS_DIR = os.path.join(BASE_DIR, "scratch", "yt_assets")
AVATAR_PATH = os.path.join(ASSETS_DIR, "avatar.jpg")
BANNER_PATH = os.path.join(ASSETS_DIR, "banner.jpg")

TARGET_CHANNEL_ID = "UChNjJbcX0n1gyKUGc3gdOBw"
STUDIO_URL = f"https://studio.youtube.com/channel/{TARGET_CHANNEL_ID}/editing/profile"

CHANNEL_NAME = "THÚY TALK"
CHANNEL_DESCRIPTION = """Chào bạn, mình là Thúy.
Đây là nơi Thúy chia sẻ những câu chuyện và phân tích xoay quanh nghề xuất nhập khẩu và thủ tục hải quan.
Không lý thuyết xa vời, không đao to búa lớn 
Mong rằng góc nhỏ này sẽ là một người bạn đồng hành quen thuộc, giúp bạn thấy công việc xuất nhập khẩu trở nên nhẹ nhàng và thú vị hơn.
Rất vui được kết nối và trò chuyện cùng bạn!"""

PREFERRED_HANDLES = ["thuystalk", "thuytalk.official", "thuytalk_official", "thuytalk.xnk"]

def log(msg):
    print(msg, flush=True)

def main():
    log("=" * 65)
    log("🚀 KHỞI ĐỘNG CỖ MÁY QUẢN TRỊ YOUTUBE STUDIO - THUYTALK")
    log("=" * 65)
    log(f"[+] Thư mục phiên (Profile): {PROFILE_DIR}")
    log(f"[+] Link đích: {STUDIO_URL}")
    log(f"[+] Avatar: {AVATAR_PATH} (Tồn tại: {os.path.exists(AVATAR_PATH)})")
    log(f"[+] Banner: {BANNER_PATH} (Tồn tại: {os.path.exists(BANNER_PATH)})")
    log("=" * 65)

    with sync_playwright() as p:
        log("[*] Đang mở trình duyệt Google Chrome trên màn hình...")
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            channel="chrome",
            args=[
                "--disable-blink-features=AutomationControlled",
                "--start-maximized"
            ],
            no_viewport=True
        )

        page = context.pages[0] if context.pages else context.new_page()
        page.goto(STUDIO_URL, wait_until="domcontentloaded", timeout=60000)

        # Kiểm tra xem có đang ở trang đăng nhập Google không
        log("[*] Đang kiểm tra trạng thái xác thực...")
        for _ in range(300): # Chờ tối đa 10 phút nếu cần đăng nhập
            curr_url = page.url
            if "accounts.google.com" in curr_url:
                log(f"[!] Đang ở trang đăng nhập Google ({curr_url}). Vui lòng đăng nhập trên cửa sổ Chrome vừa mở...")
                time.sleep(3)
            elif "studio.youtube.com" in curr_url:
                log(f"[✓] Đã nhận diện phiên YouTube Studio thành công! URL: {curr_url}")
                break
            else:
                time.sleep(2)
        else:
            log("[X] Hết thời gian chờ đăng nhập. Vui lòng thử lại!")
            return

        # Chờ trang tùy chỉnh tải xong
        log("[*] Đang chờ giao diện Tùy chỉnh kênh (Customization) tải hoàn tất...")
        time.sleep(5)

        # 1. ĐIỀN THÔNG TIN CƠ BẢN (Basic Info)
        log("\n--- [BƯỚC 1: ĐIỀN THÔNG TIN CƠ BẢN] ---")
        
        # Điền Tên Kênh
        try:
            name_input = page.wait_for_selector(
                "input#name-input, input[aria-label*='Tên'], input[aria-label*='Name'], #name-container input, ytcp-form-input-container input",
                timeout=15000
            )
            if name_input:
                name_input.click()
                # Xóa sạch text cũ
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                name_input.fill(CHANNEL_NAME)
                log(f"[✓] Đã cập nhật Tên kênh thành: '{CHANNEL_NAME}'")
        except Exception as e:
            log(f"[!] Không thể cập nhật Tên kênh: {e}")

        # Điền Mô Tả Kênh
        try:
            desc_textarea = page.wait_for_selector(
                "textarea#description-textarea, textarea[aria-label*='Mô tả'], textarea[aria-label*='Description'], #description-container textarea",
                timeout=10000
            )
            if desc_textarea:
                desc_textarea.click()
                page.keyboard.press("Control+A")
                page.keyboard.press("Backspace")
                desc_textarea.fill(CHANNEL_DESCRIPTION)
                log("[✓] Đã cập nhật Mô tả kênh (chuẩn nguyên văn @thuystalk)!")
        except Exception as e:
            log(f"[!] Không thể cập nhật Mô tả kênh: {e}")

        # Điền Handle (Thẻ người dùng)
        try:
            handle_input = page.wait_for_selector(
                "input#handle-input, input[aria-label*='Tên người dùng'], input[aria-label*='Handle'], #handle-container input",
                timeout=10000
            )
            if handle_input:
                for h in PREFERRED_HANDLES:
                    handle_input.click()
                    page.keyboard.press("Control+A")
                    page.keyboard.press("Backspace")
                    handle_input.fill(h)
                    time.sleep(1.5)
                    # Kiểm tra xem có thông báo lỗi trùng handle không
                    error_msg = page.query_selector(".error-message, ytcp-form-input-container[has-error]")
                    if not error_msg:
                        log(f"[✓] Đã chọn Handle hợp lệ: '@{h}'")
                        break
                    else:
                        log(f"[-] Handle '@{h}' đã có người dùng, thử handle tiếp theo...")
        except Exception as e:
            log(f"[!] Không thể cập nhật Handle: {e}")

        time.sleep(2)

        # 2. CHUYỂN SANG TAB XÂY DỰNG THƯƠNG HIỆU (Branding)
        log("\n--- [BƯỚC 2: TẢI LÊN AVATAR & BANNER] ---")
        try:
            branding_tab = page.query_selector(
                "tp-yt-paper-tab:has-text('Xây dựng thương hiệu'), tp-yt-paper-tab:has-text('Branding'), a[href*='/editing/images']"
            )
            if branding_tab:
                branding_tab.click()
                log("[✓] Đã chuyển sang tab 'Xây dựng thương hiệu'!")
                time.sleep(4)
            else:
                # Điều hướng thẳng URL hình ảnh
                page.goto(f"https://studio.youtube.com/channel/{TARGET_CHANNEL_ID}/editing/images")
                time.sleep(5)

            # Tìm các nút tải ảnh hoặc input file
            file_inputs = page.query_selector_all("input[type='file']")
            log(f"[*] Tìm thấy {len(file_inputs)} thẻ input upload ảnh.")

            # Upload Avatar (Mục đầu tiên)
            if os.path.exists(AVATAR_PATH):
                avatar_btn = page.query_selector(
                    "#avatar-section button:has-text('Tải lên'), #avatar-section button:has-text('Thay đổi'), button[aria-label*='ảnh hồ sơ'], button[aria-label*='profile picture']"
                )
                if avatar_btn:
                    with page.expect_file_chooser() as fc_info:
                        avatar_btn.click()
                    file_chooser = fc_info.value
                    file_chooser.set_files(AVATAR_PATH)
                    log("[✓] Đã nạp file Avatar!")
                    time.sleep(2)
                    done_btn = page.wait_for_selector("ytcp-button:has-text('Xong'), ytcp-button:has-text('Done'), #done-button", timeout=6000)
                    if done_btn:
                        done_btn.click()
                        log("[✓] Đã xác nhận crop Avatar thành công!")
                        time.sleep(2)
                elif len(file_inputs) >= 1:
                    file_inputs[0].set_files(AVATAR_PATH)
                    log("[✓] Đã gán trực tiếp file Avatar vào input 0!")

            # Upload Banner (Mục thứ hai)
            if os.path.exists(BANNER_PATH):
                banner_btn = page.query_selector(
                    "#banner-section button:has-text('Tải lên'), #banner-section button:has-text('Thay đổi'), button[aria-label*='biểu ngữ'], button[aria-label*='banner']"
                )
                if banner_btn:
                    with page.expect_file_chooser() as fc_info:
                        banner_btn.click()
                    file_chooser = fc_info.value
                    file_chooser.set_files(BANNER_PATH)
                    log("[✓] Đã nạp file Banner!")
                    time.sleep(2)
                    done_btn = page.wait_for_selector("ytcp-button:has-text('Xong'), ytcp-button:has-text('Done'), #done-button", timeout=6000)
                    if done_btn:
                        done_btn.click()
                        log("[✓] Đã xác nhận crop Banner thành công!")
                        time.sleep(2)
                elif len(file_inputs) >= 2:
                    file_inputs[1].set_files(BANNER_PATH)
                    log("[✓] Đã gán trực tiếp file Banner vào input 1!")

        except Exception as e:
            log(f"[!] Lỗi trong quá trình upload Avatar / Banner: {e}")

        # 3. BẤM NÚT XUẤT BẢN (Publish)
        log("\n--- [BƯỚC 3: XUẤT BẢN THAY ĐỔI] ---")
        try:
            publish_btn = page.query_selector(
                "ytcp-button#publish-button, button:has-text('Xuất bản'), button:has-text('Publish'), #publish-button button"
            )
            if publish_btn:
                is_disabled = publish_btn.get_attribute("disabled") or publish_btn.get_attribute("aria-disabled") == "true"
                if not is_disabled:
                    publish_btn.click()
                    log("[✓] ĐÃ BẤM NÚT 'XUẤT BẢN' THÀNH CÔNG!")
                    time.sleep(5)
                else:
                    log("[i] Nút 'Xuất bản' đang mờ (không có thay đổi mới cần lưu).")
        except Exception as e:
            log(f"[!] Lỗi khi bấm Xuất bản: {e}")

        # Chụp ảnh bằng chứng
        screenshot_path = os.path.join(ASSETS_DIR, "setup_completed.png")
        page.screenshot(path=screenshot_path)
        log(f"\n[✓] Đã chụp ảnh màn hình lưu tại: {screenshot_path}")

        log("=" * 65)
        log("🎉 HOÀN TẤT THIẾT LẬP KÊNH YOUTUBE STUDIO THUYTALK!")
        log("Phiên đăng nhập đã được lưu giữ trong yt_profile.")
        log("Từ nay, AI có thể tự động tải video và quản trị kênh hoàn toàn!")
        log("=" * 65)
        time.sleep(10)
        context.close()

if __name__ == "__main__":
    main()
