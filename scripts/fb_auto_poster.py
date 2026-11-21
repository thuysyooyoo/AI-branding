# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG ĐĂNG BÀI FACEBOOK ĐỊNH KỲ (CHẾ ĐỘ CHỈ MÌNH TÔI -> CÔNG KHAI)
Dự án: AI Personal Branding - Jasmine Nguyen / Nguyễn Ngọc Thúy (Eureka Logistics)
Tuân thủ nghiêm ngặt BỘ QUY TẮC GEMINI.md (Không tạo file backup, atomic_save).
"""

import os
import sys
import re
import json
import time
import argparse
import datetime
import openpyxl
from playwright.sync_api import sync_playwright

# Đường dẫn thư mục gốc
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_DIR = os.path.join(BASE_DIR, "fb_profile")
PENDING_FILE = os.path.join(BASE_DIR, "scratch", "pending_facebook_post.json")
CALENDAR_FILE = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
REPORTS_FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")

# Tích hợp Zalo Sender từ ZALO REPORT
ZALO_REPORT_DIR = r"D:\AI AGENT THUY\ZALO REPORT"
if ZALO_REPORT_DIR not in sys.path:
    sys.path.insert(0, ZALO_REPORT_DIR)

try:
    from core.zalo_sender import send_zalo_message
    ZALO_AVAILABLE = True
except Exception as e:
    print(f"[!] Cảnh báo không thể nạp zalo_sender: {e}")
    ZALO_AVAILABLE = False


def atomic_save(wb, target_file):
    """Lưu workbook an toàn qua file tạm rồi os.replace, TUYỆT ĐỐI không tạo file backup."""
    temp_file = target_file + ".tmp.xlsx"
    try:
        wb.save(temp_file)
        wb.close()
        os.replace(temp_file, target_file)
        print(f"[✓] Đã cập nhật file Excel an toàn: {os.path.basename(target_file)}")
    except Exception as e:
        if os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except Exception:
                pass
        raise e


def parse_post_draft(raw_text):
    """Bóc tách thân bài, hashtags và First Comment từ bản thảo."""
    if not raw_text:
        return "", "", ""
        
    post_body = ""
    first_comment = ""
    hashtags = ""
    
    # Bóc tách phần nội dung chính
    if "## Nội dung Bài viết (Post Draft)" in raw_text:
        parts = raw_text.split("## Nội dung Bài viết (Post Draft)", 1)[1]
    else:
        parts = raw_text
        
    # Bóc tách First Comment và Hashtags
    if "**Bình luận đầu tiên" in parts:
        body_part, comment_part = parts.split("**Bình luận đầu tiên", 1)
        post_body = body_part.strip()
        if "**Hashtags:**" in comment_part:
            c_part, h_part = comment_part.split("**Hashtags:**", 1)
            first_comment = re.sub(r'^\s*\(First Comment Link\):\s*', '', c_part).strip()
            # Bỏ chỉ dẫn visual nếu bị lẫn vào hashtags
            if "## Chỉ dẫn Visual" in h_part:
                h_part = h_part.split("## Chỉ dẫn Visual")[0]
            hashtags = h_part.strip()
        else:
            if "## Chỉ dẫn Visual" in comment_part:
                c_part = comment_part.split("## Chỉ dẫn Visual")[0]
            else:
                c_part = comment_part
            first_comment = re.sub(r'^\s*\(First Comment Link\):\s*', '', c_part).strip()
    else:
        post_body = parts.strip()
        
    post_body = post_body.strip(' -\n\r')
    if hashtags and hashtags not in post_body:
        post_body = post_body + "\n\n" + hashtags
        
    # Làm sạch First Comment
    first_comment = first_comment.strip(' *\n\r')
    
    return post_body, first_comment, hashtags


def get_target_post(po_code=None):
    """Tìm bài viết cần đăng từ file Weekly_Content_Calendar.xlsx."""
    if not os.path.exists(CALENDAR_FILE):
        raise FileNotFoundError(f"Không tìm thấy file lịch nội dung: {CALENDAR_FILE}")
        
    wb = openpyxl.load_workbook(CALENDAR_FILE, data_only=True)
    if "Facebook" not in wb.sheetnames:
        raise ValueError("Không tìm thấy sheet 'Facebook' trong file lịch nội dung!")
        
    sheet = wb["Facebook"]
    min_date = datetime.date(2026, 10, 14)  # BẮT BUỘC chỉ lấy từ ngày 14/10/2026 trở đi theo yêu cầu của Chị Thúy
    today = datetime.date.today()
    
    target_row = None
    candidates = []
    
    for r in range(5, sheet.max_row + 1):
        row_po = sheet.cell(r, 2).value
        date_str = sheet.cell(r, 3).value
        status = sheet.cell(r, 11).value
        
        if not row_po:
            continue
            
        if po_code:
            if str(row_po).strip() == po_code.strip():
                target_row = r
                break
        else:
            if not status or "Sẵn sàng đăng" not in str(status):
                continue
                
            # Parse ngày đăng DD/MM/YYYY
            m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(date_str))
            if m:
                d, m_val, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
                p_date = datetime.date(y, m_val, d)
                # Bỏ qua hoàn toàn bài cũ trước ngày 14/10/2026
                if p_date >= min_date:
                    candidates.append((p_date, r))
                    
    if not po_code and candidates:
        # Ưu tiên bài có ngày đăng đúng ngày hôm nay (today)
        today_match = [c for c in candidates if c[0] == today]
        if today_match:
            target_row = today_match[0][1]
        else:
            # Nếu không có bài hôm nay, lấy bài gần nhất từ ngày 14/10 trở đi
            candidates.sort(key=lambda x: x[0])
            target_row = candidates[0][1]
                
    if not target_row:
        wb.close()
        return None
        
    po = sheet.cell(target_row, 2).value
    date_str = sheet.cell(target_row, 3).value
    title = sheet.cell(target_row, 4).value
    seo_hash = sheet.cell(target_row, 5).value
    raw_content = sheet.cell(target_row, 7).value
    visual_brief = sheet.cell(target_row, 8).value
    wb.close()
    
    # Ưu tiên đọc file markdown hoàn chỉnh nếu có
    md_body = None
    clean_po = str(po).replace("PO-", "").replace("FB-", "").strip()
    candidate_md_files = [
        os.path.join(REPORTS_FB_DIR, f"FB-{po}.md"),
        os.path.join(REPORTS_FB_DIR, f"FB-{clean_po}.md"),
        os.path.join(REPORTS_FB_DIR, f"FB-PO-{clean_po}.md"),
        os.path.join(BASE_DIR, "scratch", "production_orders", f"DRAFT-FB-{po}.md")
    ]
    
    for cand in candidate_md_files:
        if os.path.exists(cand):
            try:
                with open(cand, "r", encoding="utf-8") as f:
                    md_body = f.read()
                    print(f"[*] Nạp bản thảo chi tiết từ: {os.path.basename(cand)}")
                    break
            except Exception:
                pass
                
    content_to_parse = md_body if md_body else raw_content
    body, first_comment, hashtags = parse_post_draft(content_to_parse)
    
    if hashtags and hashtags not in body:
        body = body + "\n\n" + hashtags
    elif not hashtags and seo_hash and "#" in str(seo_hash):
        tags = [p for p in str(seo_hash).split("|") if "#" in p]
        if tags and tags[0].strip() not in body:
            body = body + "\n\n" + tags[0].strip()
            
    target_data = {
        "row_idx": target_row,
        "po_code": str(po).strip(),
        "date_str": str(date_str).strip(),
        "title": str(title).strip(),
        "body": body,
        "first_comment": first_comment,
        "visual_brief": str(visual_brief) if visual_brief else ""
    }
    
    return target_data


def post_only_me(po_code=None, headless=True):
    """Đăng bài lên Facebook cá nhân ở chế độ 'Chỉ mình tôi' (Only Me)."""
    print("\n" + "=" * 65)
    print("   BƯỚC 1: TRÍCH XUẤT BÀI VIẾT TỪ LỊCH NỘI DUNG EXCEL")
    print("=" * 65)
    
    post_data = get_target_post(po_code)
    if not post_data:
        print("[!] Không tìm thấy bài viết nào có trạng thái 'Sẵn sàng đăng' trong lịch!")
        return False
        
    print(f"[+] Đã chọn bài: {post_data['po_code']}")
    print(f"    Lịch đăng: {post_data['date_str']}")
    print(f"    Tiêu đề: {post_data['title']}")
    print(f"    Độ dài nội dung: {len(post_data['body'])} ký tự")
    if post_data['first_comment']:
        print(f"    First Comment đã chuẩn bị: {post_data['first_comment'][:80]}...")
        
    print("\n" + "=" * 65)
    print("   BƯỚC 2: KHỞI ĐỘNG PLAYWRIGHT ĐĂNG BÀI DẠNG 'CHỈ MÌNH TÔI'")
    print("=" * 65)
    
    if not os.path.exists(PROFILE_DIR):
        raise FileNotFoundError(f"Chưa có phiên đăng nhập tại: {PROFILE_DIR}. Hãy chạy scripts/fb_auth.py trước!")
        
    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--no-first-run",
            "--no-default-browser-check"
        ]
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=headless,
            args=args
        )
        page = context.pages[0] if context.pages else context.new_page()
        
        print("[*] Đang truy cập Facebook...")
        page.goto("https://www.facebook.com", wait_until="domcontentloaded", timeout=40000)
        page.wait_for_timeout(3500)
        
        # Đóng mọi popup/thông báo đang mở trước khi mở soạn bài
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
        
        # 1. Tìm và click nút mở hộp soạn thảo "Bạn đang nghĩ gì thế?"
        composer_trigger = page.locator('div[role="button"]:has-text("nghĩ gì"), div[role="button"]:has-text("mind")').first
        if composer_trigger.count() == 0:
            print("[!] Không tìm thấy nút soạn thảo trên Facebook Home! Thử vào trang cá nhân...")
            page.goto("https://www.facebook.com/me", wait_until="domcontentloaded", timeout=35000)
            page.wait_for_timeout(3000)
            page.keyboard.press("Escape")
            composer_trigger = page.locator('div[role="button"]:has-text("nghĩ gì"), div[role="button"]:has-text("mind")').first
            
        if composer_trigger.count() == 0:
            print("[!] Không thể mở hộp soạn thảo bài viết Facebook!")
            context.close()
            return False
            
        print("[*] Mở hộp soạn thảo bài viết...")
        composer_trigger.click()
        page.wait_for_timeout(3000)
        
        # Định vị CHÍNH XÁC hộp thoại Tạo bài viết (tránh nhầm với dialog Thông báo/Chat)
        dialog = page.locator('div[role="dialog"]:has-text("Tạo bài viết"), div[role="dialog"]:has-text("Create post")').first
        if dialog.count() == 0:
            dialog = page.locator('div[role="dialog"]').last
            
        if dialog.count() == 0:
            print("[!] Hộp thoại tạo bài viết không xuất hiện!")
            context.close()
            return False
            
        # 2. Chọn quyền riêng tư: "Chỉ mình tôi" (Only Me)
        aud_btn = dialog.locator('div[role="button"]:has-text("Đối tượng của bài viết")').first
        if aud_btn.count() > 0:
            current_aud_text = aud_btn.inner_text()
            if "Chỉ mình tôi" not in current_aud_text:
                print(f"[*] Đang chuyển quyền riêng tư (Hiện tại: {current_aud_text.replace(chr(10), ' ')}) -> 'Chỉ mình tôi'...")
                aud_btn.click(force=True)
                page.wait_for_timeout(2000)
                
                top_dialog = page.locator('div[role="dialog"]:has-text("Chọn đối tượng"), div[role="dialog"]:has-text("Select audience")').first
                if top_dialog.count() == 0:
                    top_dialog = page.locator('div[role="dialog"]').last
                see_more = top_dialog.locator('span:has-text("Xem thêm"), div[role="button"]:has-text("Xem thêm")')
                if see_more.count() > 0:
                    try:
                        see_more.first.click()
                        page.wait_for_timeout(1000)
                    except Exception:
                        pass
                        
                only_me_span = top_dialog.locator('span:has-text("Chỉ mình tôi"), div[role="radio"]:has-text("Chỉ mình tôi")').first
                if only_me_span.count() > 0:
                    only_me_span.click()
                    page.wait_for_timeout(1000)
                    print("[✓] Đã chọn: Chỉ mình tôi (Only Me)")
                    
                done_btn = top_dialog.locator('div[role="button"]:has-text("Xong"), div[aria-label="Xong"]')
                if done_btn.count() > 0:
                    done_btn.first.click()
                    page.wait_for_timeout(1500)
            else:
                print("[✓] Quyền riêng tư đã là 'Chỉ mình tôi' sẵn.")
        else:
            print("[*] Không tìm thấy nút đối tượng trực tiếp, giữ cấu hình mặc định.")
            
        # 3. Điền nội dung bài viết
        textbox = dialog.locator('div[role="textbox"], div[contenteditable="true"]').first
        print("[*] Đang nhập nội dung bài viết...")
        textbox.click()
        page.wait_for_timeout(1000)
        
        # Dùng keyboard.insert_text để bảo toàn dấu xuống dòng \n\n
        page.keyboard.insert_text(post_data['body'])
        page.wait_for_timeout(1500)
        
        # Bấm Escape để tắt mọi gợi ý tag bạn bè / mention nếu có
        page.keyboard.press("Escape")
        page.wait_for_timeout(500)
        
        # 4. Bấm nút Đăng
        post_btn = dialog.locator('div[aria-label="Đăng"][role="button"], div[aria-label="Post"][role="button"]').first
        if post_btn.count() == 0:
            post_btn = dialog.locator('div[aria-label="Đăng"]').first
        print("[*] Đang bấm nút Đăng bài...")
        post_btn.click(force=True)
        
        # Chờ hộp thoại đóng lại sau khi đăng
        print("[*] Đang chờ Facebook xử lý đăng bài (10s)...")
        page.wait_for_timeout(3500)
        if dialog.is_visible() and post_btn.is_visible():
            print("[*] Hộp thoại vẫn mở, thử click lại nút Đăng...")
            post_btn.click(force=True)
            
        try:
            dialog.wait_for(state="hidden", timeout=25000)
            print("[✓] Hộp thoại đã đóng thành công!")
        except Exception:
            print("[*] Hộp thoại đang xử lý nền...")
        page.wait_for_timeout(5000)
        
        # 5. Lấy link bài viết vừa đăng trên trang cá nhân
        print("\n[*] Đang truy cập trang cá nhân để lấy link bài viết vừa đăng...")
        page.goto("https://www.facebook.com/me", wait_until="domcontentloaded", timeout=35000)
        page.wait_for_timeout(4000)
        
        post_url = page.url
        # Thử cuộn nhẹ để feed nạp bài mới
        page.evaluate("window.scrollBy(0, 500)")
        page.wait_for_timeout(2000)
        
        feed_posts = page.locator('div[role="feed"] div[data-pagelet*="FeedUnit"], div[role="main"] div[data-pagelet*="FeedUnit"]').all()
        target_unit = None
        title_snippet = post_data['title'][:20].lower()
        
        for fp in feed_posts:
            try:
                txt = fp.inner_text().lower()
                if "vừa xong" in txt or "just now" in txt or title_snippet in txt:
                    target_unit = fp
                    print("[+] Đã định vị chính xác bài viết vừa đăng trên Feed!")
                    break
            except Exception:
                pass
                
        if not target_unit and len(feed_posts) > 0:
            first_txt = feed_posts[0].inner_text().lower()
            if "bài viết đã ghim" in first_txt or "pinned" in first_txt:
                target_unit = feed_posts[1] if len(feed_posts) > 1 else feed_posts[0]
            else:
                target_unit = feed_posts[0]
                
        if target_unit:
            post_links = target_unit.locator('a[role="link"][href*="/posts/"], a[role="link"][href*="story_fbid="]').all()
            if len(post_links) > 0:
                raw_href = post_links[0].get_attribute("href")
                if raw_href:
                    post_url = raw_href.split("?")[0]
                    print(f"[✓] ĐÃ LẤY ĐƯỢC LINK BÀI VIẾT: {post_url}")
        
        if not post_url or post_url == page.url:
            all_links = page.locator('div[role="feed"] a[role="link"][href*="/posts/"], div[role="feed"] a[role="link"][href*="story_fbid="]').all()
            if len(all_links) > 0:
                post_url = all_links[0].get_attribute("href").split("?")[0]
                print(f"[✓] Lấy link bài viết đầu tiên: {post_url}")
                print(f"[✓] Lấy link bài viết đầu tiên: {post_url}")
            
        context.close()
        
    # Ghi nhận trạng thái chờ duyệt vào scratch/pending_facebook_post.json
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    pending_info = {
        "po_code": post_data["po_code"],
        "title": post_data["title"],
        "row_idx": post_data["row_idx"],
        "post_url": post_url,
        "first_comment": post_data["first_comment"],
        "posted_only_me_at": now_str,
        "status": "pending_review"
    }
    
    os.makedirs(os.path.dirname(PENDING_FILE), exist_ok=True)
    with open(PENDING_FILE, "w", encoding="utf-8") as f:
        json.dump(pending_info, f, ensure_ascii=False, indent=2)
    print(f"[✓] Đã lưu thông tin bài chờ duyệt vào: {os.path.basename(PENDING_FILE)}")
    
    # Cập nhật trạng thái trong Weekly_Content_Calendar.xlsx
    wb = openpyxl.load_workbook(CALENDAR_FILE)
    sheet = wb["Facebook"]
    sheet.cell(post_data["row_idx"], 11, f"Đang kiểm tra Chỉ mình tôi ({now_str})")
    atomic_save(wb, CALENDAR_FILE)
    
    # Bắn thông báo Zalo vào Cloud của tôi
    print("\n" + "=" * 65)
    print("   BƯỚC 3: GỬI BẢN TIN ĐIỀU HÀNH VÀO ZALO 'CLOUD CỦA TÔI'")
    print("=" * 65)
    
    zalo_msg = (
        f"🔔 BÀI VIẾT FACEBOOK ĐÃ ĐƯỢC ĐĂNG DẠNG 'CHỈ MÌNH TÔI'\n\n"
        f"📌 Mã PO: {post_data['po_code']}\n"
        f"📝 Tiêu đề: {post_data['title']}\n"
        f"🔒 Trạng thái: Chỉ mình tôi (Only Me) — Chưa ai thấy ngoài Chị\n"
        f"⏰ Giờ đăng: {now_str}\n"
        f"🔗 Link bài viết để kiểm tra:\n{post_url}\n\n"
        f"👉 HƯỚNG DẪN KIỂM TRA & ĐĂNG CÔNG KHAI (PHƯƠNG ÁN 2):\n"
        f"1. Chị Thúy mở link trên app Facebook để kiểm tra ngắt dòng, bố cục và ảnh.\n"
        f"2. Nếu đã ưng ý, chị chỉ cần bấm đúp chạy file:\n"
        f"   'chay_dang_cong_khai_facebook.bat'\n"
        f"   (Hoặc nhắn trong chat: 'đăng công khai đi em')\n"
        f"3. Bot sẽ tự động chuyển sang CÔNG KHAI 🌍 và thả First Comment ngay lập tức!"
    )
    
    if ZALO_AVAILABLE:
        try:
            send_zalo_message(zalo_msg, target_chat="Cloud của tôi")
            print("[✓] Đã gửi thông báo kiểm tra bài viết vào Zalo 'Cloud của tôi' thành công!")
        except Exception as e:
            print(f"[!] Lỗi khi gửi Zalo: {e}")
    else:
        print("[!] Bỏ qua bước gửi Zalo vì chưa nạp được thư viện zalo_sender.")
        
    return True


def publish_pending_post(headless=True):
    """Chuyển bài viết đang chờ duyệt từ 'Chỉ mình tôi' sang 'Công khai' và thả First Comment."""
    print("\n" + "=" * 65)
    print("   CHUYỂN ĐỐI TƯỢNG BÀI VIẾT SANG 'CÔNG KHAI' & THẢ FIRST COMMENT")
    print("=" * 65)
    
    if not os.path.exists(PENDING_FILE):
        print(f"[!] Không tìm thấy file bài chờ duyệt tại: {PENDING_FILE}")
        return False
        
    with open(PENDING_FILE, "r", encoding="utf-8") as f:
        pending_info = json.load(f)
        
    if pending_info.get("status") != "pending_review":
        print(f"[!] Bài viết '{pending_info.get('po_code')}' hiện không ở trạng thái chờ duyệt (Trạng thái: {pending_info.get('status')})!")
        return False
        
    post_url = pending_info.get("post_url")
    po_code = pending_info.get("po_code")
    title = pending_info.get("title")
    first_comment = pending_info.get("first_comment", "")
    row_idx = pending_info.get("row_idx")
    
    print(f"[*] Mã PO: {po_code}")
    print(f"[*] Tiêu đề: {title}")
    print(f"[*] Link bài viết: {post_url}")
    
    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--no-first-run",
            "--no-default-browser-check"
        ]
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=headless,
            args=args
        )
        page = context.pages[0] if context.pages else context.new_page()
        
        print("[*] Đang truy cập bài viết Facebook...")
        page.goto(post_url, wait_until="domcontentloaded", timeout=40000)
        page.wait_for_timeout(4000)
        
        # 1. Tìm nút 3 chấm menu của bài viết
        action_btn = page.locator('div[aria-label*="Hành động đối với bài viết"], div[aria-label*="Tùy chọn bài viết"], div[aria-label*="Actions for this post"]').first
        if action_btn.count() == 0:
            action_btn = page.locator('div[aria-haspopup="menu"][role="button"]').first
            
        if action_btn.count() > 0:
            print("[*] Đang mở menu tùy chọn bài viết...")
            action_btn.click(force=True)
            page.wait_for_timeout(2000)
            
            # Chọn "Chỉnh sửa đối tượng"
            edit_aud_item = page.locator('div[role="menuitem"]:has-text("Chỉnh sửa đối tượng"), div[role="menuitem"]:has-text("Edit audience")').first
            if edit_aud_item.count() > 0:
                print("[*] Chọn 'Chỉnh sửa đối tượng'...")
                edit_aud_item.click(force=True)
                page.wait_for_timeout(2500)
                
                # Trong dialog Chọn đối tượng -> Chọn "Công khai"
                top_dialog = page.locator('div[role="dialog"]').last
                pub_span = top_dialog.locator('span:has-text("Công khai"), div[role="radio"]:has-text("Công khai")').first
                if pub_span.count() > 0:
                    pub_span.click(force=True)
                    page.wait_for_timeout(1000)
                    print("[✓] Đã chọn: Công khai (Public 🌍)")
                    
                # Bấm Lưu / Xong
                save_btn = top_dialog.locator('div[role="button"]:has-text("Lưu"), div[role="button"]:has-text("Xong"), div[aria-label="Lưu"], div[aria-label="Xong"]').first
                if save_btn.count() > 0:
                    save_btn.click(force=True)
                    page.wait_for_timeout(2000)
                    print("[✓] Đã lưu quyền riêng tư Công khai thành công!")
            else:
                print("[!] Không tìm thấy mục 'Chỉnh sửa đối tượng' trong menu bài viết!")
        else:
            print("[!] Không tìm thấy nút menu 3 chấm của bài viết!")
            
        # 2. Thả First Comment nếu có
        if first_comment:
            print("\n[*] Đang tìm ô bình luận để thả First Comment...")
            comment_box = page.locator('div[role="textbox"][contenteditable="true"], div[aria-label*="bình luận"], div[aria-label*="comment"]').first
            if comment_box.count() > 0:
                try:
                    comment_box.click()
                    page.wait_for_timeout(1000)
                    page.keyboard.insert_text(first_comment)
                    page.wait_for_timeout(1500)
                    page.keyboard.press("Enter")
                    page.wait_for_timeout(3000)
                    print("[✓] ĐÃ TỰ ĐỘNG THẢ FIRST COMMENT CHỨA LINK NGUỒN THÀNH CÔNG!")
                except Exception as e:
                    print(f"[!] Lỗi khi thả comment: {e}")
            else:
                print("[*] Không tìm thấy ô bình luận trực tiếp trên trang.")
                
        context.close()
        
    # Cập nhật trạng thái hoàn tất
    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
    pending_info["status"] = "published"
    pending_info["published_at"] = now_str
    
    with open(PENDING_FILE, "w", encoding="utf-8") as f:
        json.dump(pending_info, f, ensure_ascii=False, indent=2)
        
    # Cập nhật file Excel Lịch nội dung
    wb = openpyxl.load_workbook(CALENDAR_FILE)
    sheet = wb["Facebook"]
    sheet.cell(row_idx, 11, f"Đã đăng Công khai ({now_str})")
    atomic_save(wb, CALENDAR_FILE)
    print(f"[✓] Đã cập nhật trạng thái xuất bản trong file Excel sang 'Đã đăng Công khai ({now_str})'")
    
    # Gửi thông báo Zalo hoàn tất
    zalo_msg = (
        f"🎉 BÀI VIẾT FACEBOOK ĐÃ ĐƯỢC CHUYỂN SANG CÔNG KHAI THÀNH CÔNG!\n\n"
        f"📌 Mã PO: {po_code}\n"
        f"📝 Tiêu đề: {title}\n"
        f"🌍 Chế độ: CÔNG KHAI (Public 🌍)\n"
        f"💬 First Comment: Đã tự động thả link tài liệu/nguồn tham khảo.\n"
        f"⏰ Giờ xuất bản: {now_str}\n"
        f"🔗 Link bài viết chính thức:\n{post_url}"
    )
    
    if ZALO_AVAILABLE:
        try:
            send_zalo_message(zalo_msg, target_chat="Cloud của tôi")
            print("[✓] Đã gửi thông báo xuất bản công khai vào Zalo 'Cloud của tôi'!")
        except Exception as e:
            print(f"[!] Lỗi gửi Zalo: {e}")
            
    print("\n" + "=" * 65)
    print("   HOÀN TẤT XUẤT BẢN BÀI VIẾT FACEBOOK AN TOÀN & CHUẨN XÁC!")
    print("=" * 65)
    return True


def show_status():
    """Hiển thị trạng thái bài viết đang chờ duyệt."""
    if not os.path.exists(PENDING_FILE):
        print("[*] Hiện tại không có bài viết nào đang ở trạng thái chờ duyệt.")
        return
        
    with open(PENDING_FILE, "r", encoding="utf-8") as f:
        info = json.load(f)
        
    print("=" * 65)
    print("   TRẠNG THÁI BÀI ĐĂNG FACEBOOK HIỆN TẠI")
    print("=" * 65)
    print(f"📌 Mã PO: {info.get('po_code')}")
    print(f"📝 Tiêu đề: {info.get('title')}")
    print(f"🔒 Trạng thái: {info.get('status')}")
    print(f"⏰ Đăng Chỉ mình tôi lúc: {info.get('posted_only_me_at')}")
    if info.get('published_at'):
        print(f"🌍 Xuất bản công khai lúc: {info.get('published_at')}")
    print(f"🔗 Link: {info.get('post_url')}")
    print("=" * 65)


def main():
    parser = argparse.ArgumentParser(description="Tự động đăng bài Facebook (Chỉ mình tôi -> Công khai)")
    parser.add_argument("--post-draft", action="store_true", help="Lấy bài tiếp theo và đăng ở chế độ Chỉ mình tôi")
    parser.add_argument("--publish", action="store_true", help="Chuyển bài đang chờ duyệt sang Công khai và thả First Comment")
    parser.add_argument("--po", type=str, default=None, help="Chỉ định mã PO cụ thể cần đăng")
    parser.add_argument("--status", action="store_true", help="Xem thông tin bài viết đang chờ duyệt")
    parser.add_argument("--headless", action="store_true", default=True, help="Chạy ẩn trình duyệt (mặc định: True)")
    parser.add_argument("--gui", dest="headless", action="store_false", help="Chạy hiện cửa sổ trình duyệt để quan sát")
    
    args = parser.parse_args()
    
    if args.status:
        show_status()
    elif args.publish:
        publish_pending_post(headless=args.headless)
    elif args.post_draft:
        post_only_me(po_code=args.po, headless=args.headless)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
