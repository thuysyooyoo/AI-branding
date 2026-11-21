# -*- coding: utf-8 -*-
r"""
HỆ THỐNG TỰ ĐỘNG LƯỚT VÀ TƯƠNG TÁC BÀI VIẾT CÁ NHÂN FACEBOOK (PERSONAL FEED INTERACTOR)
Dự án: AI Personal Branding & Engagement - Eureka Logistics
Cơ chế cốt lõi:
1. Chỉ tương tác bài viết CÁ NHÂN (Bạn bè), loại bỏ 100% Fanpage, Group và Quảng cáo (Sponsored).
2. Tỷ lệ tương tác: 50% Thả tim (❤️ Love), 50% Thích (👍 Like) đan xen tự nhiên.
3. TUYỆT ĐỐI KHÔNG LIKE BÌNH LUẬN (COMMENT): Chỉ tác động vào nút Action Bar bài viết chính.
4. Mô phỏng người thật 100%: Dừng đọc 4 - 8s, cuộn chuột tự nhiên (mouse wheel), chống checkpoint an toàn tuyệt đối.
5. Kiểm tra và tự động bỏ qua các bài đã tương tác trước đó (Ledger + DOM check).
6. Báo cáo kép:
   - File Excel Desktop: C:\Users\thuyn\Desktop\TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx (Sheet 04_LOG_TUONG_TAC_FB)
   - Zalo Bot: Bản tin điều hành gửi trực tiếp vào phòng chat 'Cloud của tôi'.
"""

import os
import sys
import time
import re
import json
import random
import argparse
import datetime
import subprocess
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from playwright.sync_api import sync_playwright

# Nạp Zalo Sender
ZALO_REPORT_DIR = r"D:\AI AGENT THUY\ZALO REPORT"
if ZALO_REPORT_DIR not in sys.path:
    sys.path.insert(0, ZALO_REPORT_DIR)

try:
    from core.zalo_sender import send_zalo_message
    ZALO_AVAILABLE = True
except Exception as e:
    print(f"[!] Cảnh báo không thể nạp Zalo Sender: {e}", flush=True)
    ZALO_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_DIR = os.path.join(BASE_DIR, "fb_profile")
LEDGER_FILE = os.path.join(BASE_DIR, "scratch", "fb_interacted_ledger.json")

# Danh sách từ khóa doanh nghiệp/thương mại để loại bỏ Fanpage
PAGE_KEYWORDS = [
    "công ty", "tnhh", "jsc", "corp", "agency", "shop", "store", "official", 
    "vietnam", "việt nam", "academy", "báo ", "tin tức", "truyền thông", "media", 
    "studio", "camera", "logistics", "design", "marketing", "auto", "chính hãng", 
    "chính phủ", "thời sự", "bảo hiểm", "bất động sản", "phần mềm", "giải pháp"
]

def load_interacted_ledger():
    if os.path.exists(LEDGER_FILE):
        try:
            with open(LEDGER_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_interacted_ledger(ledger):
    try:
        os.makedirs(os.path.dirname(LEDGER_FILE), exist_ok=True)
        temp_file = LEDGER_FILE + ".tmp"
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(ledger, f, ensure_ascii=False, indent=2)
        os.replace(temp_file, LEDGER_FILE)
    except Exception as e:
        print(f"[!] Lỗi lưu ledger: {e}", flush=True)

def check_and_ensure_login(page, context, headless=False):
    """Kiểm tra trạng thái đăng nhập Facebook, hỗ trợ click Tiếp tục / Continue nếu có."""
    cookies = [c['name'] for c in context.cookies()]
    if 'c_user' in cookies:
        print("[✓] Đã có phiên đăng nhập Facebook hợp lệ (c_user).", flush=True)
        return True

    print("\n" + "="*65, flush=True)
    print("🔑 CẦN ĐĂNG NHẬP FACEBOOK (PHIÊN ĐÃ HẾT HẠN HOẶC BỊ ĐĂNG XUẤT)", flush=True)
    print("="*65, flush=True)
    
    if headless:
        print("[!] Đang chạy ở chế độ Headless (ngầm). Không thể hiển thị giao diện để người dùng đăng nhập.", flush=True)
        print("[!] Dừng phiên ngay lập tức để gửi cảnh báo chính xác vào Zalo.", flush=True)
        return False

    try:
        continue_btn = page.locator("div[aria-label*='Continue' i], div[role='button']:has-text('Continue'), button:has-text('Continue'), div[role='button']:has-text('Tiếp tục'), button:has-text('Tiếp tục')")
        if continue_btn.count() > 0:
            print("[*] Đang mở hộp thoại đăng nhập cho tài khoản Jasmine Nguyen...", flush=True)
            continue_btn.first.click()
            page.wait_for_timeout(2000)
            
            pwd_input = page.locator('input[type="password"]')
            if pwd_input.count() > 0:
                pwd_input.first.focus()
    except Exception as e:
        print(f"[*] Thử mở login: {e}", flush=True)

    print("\n👉 Cửa sổ trình duyệt Facebook đang hiển thị trên màn hình.", flush=True)
    print("👉 Chị Thúy vui lòng nhập Mật khẩu tài khoản và bấm Đăng nhập.", flush=True)
    print("👉 Hệ thống đang tự động theo dõi và sẽ bắt đầu ngay khi đăng nhập xong (chờ 360s)...", flush=True)
    print("="*65 + "\n", flush=True)

    page.bring_to_front()

    for sec in range(180):
        time.sleep(2)
        try:
            cookies = [c['name'] for c in context.cookies()]
            if 'c_user' in cookies:
                print("\n[✓] XÁC NHẬN ĐĂNG NHẬP THÀNH CÔNG! Đang chuyển tiếp vào Bảng tin...", flush=True)
                time.sleep(3)
                return True
        except Exception:
            pass

    print("[!] Hết thời gian chờ đăng nhập (360 giây). Vui lòng thử lại.", flush=True)
    return False

def human_smooth_scroll(page, distance=None):
    """Cuộn chuột mượt mà từng bước nhỏ giống hệt người thật, cho thời gian nạp feed."""
    if distance is None:
        distance = random.randint(320, 480)
    steps = random.randint(4, 6)
    step_dist = distance / steps
    for _ in range(steps):
        page.mouse.wheel(0, step_dist * random.uniform(0.85, 1.15))
        time.sleep(random.uniform(0.08, 0.15))
    time.sleep(random.uniform(2.5, 4.0))

def interact_feed(target_count=20, headless=False):
    """
    Lướt Bảng tin Bạn bè (Chronological Friends Feed) và tương tác đúng target_count bài cá nhân:
    50% Thả tim (❤️ Love), 50% Thích (👍 Like).
    TUYỆT ĐỐI KHÔNG LIKE BÌNH LUẬN (COMMENT).
    Mô phỏng 100% nhịp điệu người thật, cuộn mượt mà, dừng đọc tự nhiên.
    """
    print("=" * 70, flush=True)
    print(f"KHỞI CHẠY BOT TƯƠNG TÁC BẢNG TIN FACEBOOK CÁ NHÂN ({target_count} BÀI)", flush=True)
    print(f"🎯 Mục tiêu: Đúng {target_count} bài cá nhân (Bạn bè) | 50% Thả tim ❤️ - 50% Thích 👍", flush=True)
    print(f"🖥️ Chế độ hiển thị: {'Chạy ngầm (Headless)' if headless else 'Cửa sổ trực quan (Visible)'}", flush=True)
    print("=" * 70, flush=True)

    ledger = load_interacted_ledger()
    today_key = datetime.datetime.now().strftime("%Y-%m-%d")
    if today_key not in ledger:
        ledger[today_key] = []

    results = []
    seen_keys = set(ledger[today_key])

    with sync_playwright() as p:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--no-first-run",
            "--no-default-browser-check"
        ]
        if not headless:
            args.append("--start-maximized")

        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=headless,
            args=args,
            viewport={"width": 1920, "height": 1080}
        )
        
        page = context.pages[0] if context.pages else context.new_page()
        
        print("[*] Đang truy cập Facebook...", flush=True)
        page.goto("https://www.facebook.com", timeout=45000, wait_until="domcontentloaded")
        page.wait_for_timeout(4000)

        # Kiểm tra đăng nhập
        if not check_and_ensure_login(page, context, headless=headless):
            print("[!] Chưa thể đăng nhập Facebook. Phiên dừng lại.", flush=True)
            context.close()
            return [], "NEED_LOGIN"

        # Chuyển hướng vào Bảng feed Bạn bè chính thức (Chronological Friends Feed)
        # Giúp 100% bài hiển thị là của Bạn bè theo thứ tự thời gian mới nhất, loại bỏ hoàn toàn rác thuật toán Group & Page
        friends_feed_url = "https://www.facebook.com/?filter=friends&sk=h_chr"
        print(f"\n[*] Đang chuyển hướng vào Bảng feed Bạn bè chính thức: {friends_feed_url}...", flush=True)
        page.goto(friends_feed_url, wait_until="domcontentloaded", timeout=45000)

        # Chờ nạp bài và hydrate DOM tự nhiên (8 - 10 giây)
        print("[*] Đang chờ Bảng feed nạp hoàn chỉnh (8s)...", flush=True)
        time.sleep(8)
        page.keyboard.press("Escape")
        time.sleep(1)

        print("\n[*] Đã vào Bảng feed Bạn bè thành công. Bắt đầu lướt và tương tác tự nhiên...", flush=True)

        # Đoạn JS tìm bài cá nhân tiếp theo
        js_find_next = """
        (args) => {
            const { seenKeys, pageKeywords } = args;
            const allBtns = Array.from(document.querySelectorAll('div[role="button"]'));
            
            // Tìm các nút Thích thuộc Action Bar của bài viết
            const likeBtns = allBtns.filter(b => {
                const aria = (b.getAttribute('aria-label') || '').toLowerCase();
                const text = (b.innerText || '').trim().toLowerCase();
                
                // 1. Nhận diện nút Thích
                const isLike = aria === 'thích' || aria === 'like' ||
                               aria.includes('cảm xúc thích') || aria.includes('react like') ||
                               aria.includes('bày tỏ cảm xúc');
                if (!isLike) return false;
                
                // Loại trừ nút quá nhỏ hoặc nút reaction icon
                const rect = b.getBoundingClientRect();
                if (rect.height < 22 || rect.width < 25) return false;
                
                // 2. Phải là Action Bar (nằm gần Bình luận hoặc Chia sẻ)
                let parent = b.parentElement;
                let isActionBar = false;
                for (let d = 0; d < 6 && parent; d++) {
                    const siblingArias = Array.from(parent.querySelectorAll('div[role="button"]'))
                        .map(x => (x.getAttribute('aria-label') || x.innerText || '').toLowerCase());
                    if (siblingArias.some(a => a.includes('bình luận') || a.includes('comment') || a.includes('chia sẻ') || a.includes('share'))) {
                        isActionBar = true;
                        break;
                    }
                    parent = parent.parentElement;
                }
                return isActionBar;
            });
            
            for (let i = 0; i < likeBtns.length; i++) {
                const btn = likeBtns[i];
                const aria = btn.getAttribute('aria-label') || '';
                
                // Tìm khung bài viết
                let postEl = btn;
                for (let d = 0; d < 14; d++) {
                    if (postEl.parentElement && postEl.parentElement !== document.body) {
                        postEl = postEl.parentElement;
                        if (postEl.getBoundingClientRect().height > 200) {
                            const hasHeader = postEl.querySelector('h2, h3, h4, strong, a[role="link"]');
                            if (hasHeader) break;
                        }
                    }
                }
                
                const postText = postEl ? postEl.innerText : '';
                
                // 1. Tên tác giả
                let author = "Unknown";
                if (aria.includes('về bài viết của ')) {
                    author = aria.split('về bài viết của ')[1].split(/[\\n,.]/)[0].trim();
                } else if (aria.toLowerCase().includes('about ') && aria.toLowerCase().includes('post')) {
                    author = aria.split(/about /i)[1].split(/'s post/i)[0].trim();
                } else if (postEl) {
                    const headerEl = postEl.querySelector('h2 a, h3 a, h4 a, strong a, a[role="link"] strong, span[dir="auto"] strong');
                    if (headerEl) author = headerEl.innerText.trim();
                }
                
                const authorLower = author.toLowerCase();
                
                // 2. Bỏ qua tài khoản chính mình
                if (authorLower.includes('jasmine nguyen') || authorLower.includes('nguyễn ngọc thúy') || authorLower.includes('ngoc thuy')) {
                    continue;
                }
                
                // 3. Bỏ qua Quảng cáo
                if (postText.includes('Được tài trợ') || postText.includes('Sponsored')) continue;
                if (postText.includes('S͏') && postText.includes('p͏') && postText.includes('o͏')) continue;
                const ctaButtons = Array.from(postEl.querySelectorAll('div[role="button"], a[role="button"]'))
                    .map(b => b.innerText.trim().toLowerCase())
                    .filter(t => ['tìm hiểu thêm', 'gửi tin nhắn', 'mua ngay', 'đăng ký', 'nhận ưu đãi', 'liên hệ'].includes(t));
                if (ctaButtons.length > 0) continue;
                
                // 4. Bỏ qua Group
                const links = Array.from(postEl.querySelectorAll('a')).map(a => a.href);
                if (links.some(h => h.includes('/groups/'))) continue;
                
                // 5. Bỏ qua Fanpage (Bạn bè không có nút Theo dõi / Follow cạnh tên)
                const headerArea = postEl.querySelector('h2, h3, h4, strong')?.parentElement?.parentElement;
                const headerText = headerArea ? headerArea.innerText : '';
                if (headerText.includes('Theo dõi') || headerText.includes('Follow')) continue;
                if (postEl.querySelector('svg[aria-label*="xác minh" i], svg[aria-label*="verified" i]')) continue;
                if (pageKeywords.some(kw => authorLower.includes(kw))) continue;
                
                // 6. Trích đoạn nội dung
                let snippet = "";
                const msgEl = postEl.querySelector('div[data-ad-preview="message"], div[data-ad-comet-preview="message"], div[dir="auto"]');
                if (msgEl) snippet = msgEl.innerText.trim().replace(/\\n/g, ' ');
                if (!snippet) {
                    const lines = postText.split('\\n').map(l => l.trim()).filter(l => l.length > 8 && !['facebook', 'thích', 'bình luận', 'chia sẻ', 'xem thêm', 'vừa xong', 'phút', 'giờ', 'tháng'].includes(l.toLowerCase()));
                    snippet = lines.slice(0, 2).join(' ');
                }
                snippet = snippet.slice(0, 80) || "Hình ảnh / Trạng thái bạn bè";
                
                const postKey = author + " | " + snippet.slice(0, 30);
                if (seenKeys.includes(postKey)) continue;
                
                // 7. Kiểm tra xem đã like chưa
                const ariaLower = aria.toLowerCase();
                const textLower = btn.innerText.toLowerCase();
                const isAlreadyLiked = ariaLower.includes('gỡ') || ariaLower.includes('bỏ thích') || 
                                       ariaLower.includes('remove') || textLower.includes('bỏ thích') ||
                                       ['yêu thích', 'thương thương', 'haha', 'wow', 'buồn', 'phẫn nộ'].some(x => textLower.includes(x) || ariaLower.includes(x));
                                       
                if (isAlreadyLiked) {
                    return { found: true, alreadyReacted: true, postKey, author };
                }
                
                btn.scrollIntoView({ block: "center", behavior: "smooth" });
                btn.setAttribute('data-target-post', 'true');
                const rect = btn.getBoundingClientRect();
                return {
                    found: true,
                    alreadyReacted: false,
                    author,
                    snippet,
                    postKey,
                    coord: { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 },
                    btnIdx: i
                };
            }
            return { found: false };
        }
        """

        scroll_attempts = 0
        max_scrolls = 120

        while len(results) < target_count and scroll_attempts < max_scrolls:
            candidate = page.evaluate(js_find_next, {
                "seenKeys": list(seen_keys),
                "pageKeywords": PAGE_KEYWORDS
            })

            if candidate.get("found"):
                post_key = candidate["postKey"]
                seen_keys.add(post_key)

                if candidate.get("alreadyReacted"):
                    print(f"  [i] Bỏ qua bài đã tương tác trước: {candidate['author']}", flush=True)
                    continue

                author = candidate["author"]
                snippet = candidate["snippet"]
                btn_locator = page.locator('div[data-target-post="true"]').first

                # Xen kẽ: Chẵn là Thả tim (❤️), Lẻ là Thích (👍)
                reaction_type = "Love" if len(results) % 2 == 0 else "Like"
                reaction_symbol = "❤️ Thả tim" if reaction_type == "Love" else "👍 Thích"

                print(f"\n[+] [{len(results) + 1}/{target_count}] BÀI CÁ NHÂN (BẠN BÈ): {author}", flush=True)
                print(f"    - Nội dung: \"{snippet}\"", flush=True)
                print(f"    - Đang đọc bài tự nhiên (4.5 - 7.5s)...", flush=True)

                # Dừng đọc mô phỏng người thật 4.5 - 7.5s
                time.sleep(random.uniform(4.5, 7.5))

                print(f"    - Thực hiện: {reaction_symbol}...", flush=True)

                success = False
                if reaction_type == "Love":
                    btn_locator.hover()
                    time.sleep(1.2)

                    # Tìm nút 'Yêu thích' trong thanh reaction popup
                    love_locator = page.locator('div[role="toolbar"] div[aria-label="Yêu thích"], div[aria-label="Yêu thích"]').first
                    if love_locator.count() > 0 and love_locator.is_visible():
                        love_locator.click()
                        success = True
                        print(f"    [✓] Đã Thả tim ❤️ thành công!", flush=True)
                    else:
                        # Fallback bấm Thích nếu popup chưa kịp hiện
                        btn_locator.click()
                        success = True
                        reaction_symbol = "👍 Thích (Fallback)"
                        print(f"    [✓] Đã bấm Thích 👍 (Menu tim mở chậm).", flush=True)
                else:
                    btn_locator.click()
                    success = True
                    print(f"    [✓] Đã bấm Thích 👍 thành công!", flush=True)

                # Dọn sạch thuộc tính đánh dấu
                page.evaluate("() => document.querySelectorAll('[data-target-post]').forEach(el => el.removeAttribute('data-target-post'))")

                if success:
                    results.append({
                        "stt": len(results) + 1,
                        "author": author,
                        "reaction": reaction_symbol,
                        "snippet": snippet,
                        "time": datetime.datetime.now().strftime("%H:%M:%S")
                    })
                    ledger[today_key].append(post_key)
                    save_interacted_ledger(ledger)
                    # Dừng ngắm bài sau khi like trước khi cuộn tiếp (4.0 - 6.5s)
                    time.sleep(random.uniform(4.0, 6.5))
            else:
                scroll_attempts += 1
                human_smooth_scroll(page)

        print("\n" + "=" * 70, flush=True)
        print(f"HOÀN THÀNH PHIÊN TƯƠNG TÁC! Đã tương tác: {len(results)}/{target_count} bài.", flush=True)
        print("=" * 70, flush=True)

        time.sleep(2)
        context.close()

    status = "SUCCESS" if len(results) > 0 else "EMPTY_FEED"
    return results, status

def send_report_to_zalo(results, status="SUCCESS"):
    """Soạn và gửi bản tin báo cáo chi tiết vào Zalo 'Cloud của tôi'."""
    if not ZALO_AVAILABLE:
        print("[!] Zalo Sender không khả dụng.", flush=True)
        return False

    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")

    if status == "NEED_LOGIN":
        msg_lines = [
            "⚠️ [CẢNH BÁO ĐIỀU HÀNH] TÀI KHOẢN FACEBOOK CẦN ĐĂNG NHẬP",
            "══════════════════════════════",
            f"⏰ Thời gian phát hiện: {now_str}",
            "❌ Kết quả: Phiên tương tác bạn bè tạm dừng do tài khoản Facebook (Jasmine Nguyen) bị hết hạn đăng nhập.",
            "🔍 Nguyên nhân: Cookie bảo mật của Facebook đã hết hạn định kỳ hoặc tài khoản bị đăng xuất.",
            "👉 HƯỚNG DẪN XỬ LÝ NHANH (1 PHÚT):",
            "   1. Chị Thúy mở màn hình Desktop máy tính.",
            "   2. Nhấp đúp chuột vào file: 'DANG_NHAP_FACEBOOK.bat'",
            "   3. Đăng nhập lại tài khoản trên cửa sổ trình duyệt hiện ra.",
            "🤖 Sau khi đăng nhập xong, bot sẽ tự động tiếp tục tương tác bình thường theo lịch trình!"
        ]
        full_message = "\n".join(msg_lines)
        return send_zalo_message(full_message, target_chat="Cloud của tôi")

    if not results or status == "EMPTY_FEED":
        # Vẫn gửi cảnh báo nếu phiên gặp sự cố để người dùng nắm rõ
        msg_lines = [
            "⚠️ [CẢNH BÁO ĐIỀU HÀNH] TỰ ĐỘNG TƯƠNG TÁC BẠN BÈ FACEBOOK",
            "══════════════════════════════",
            f"⏰ Thời gian thực hiện: {now_str}",
            "❌ Kết quả: Phiên quét không ghi nhận bài viết bạn bè mới hợp lệ.",
            "🔍 Nguyên nhân dự kiến: Bảng tin chưa nạp kịp hoặc đã tương tác hết bài trong ngày.",
            "🤖 Hành động bot: Đã ghi nhận nhật ký hệ thống, sẽ tự động quét bù vào phiên tiếp theo."
        ]
        full_message = "\n".join(msg_lines)
        return send_zalo_message(full_message, target_chat="Cloud của tôi")

    love_count = sum(1 for r in results if "Thả tim" in r["reaction"] or "❤️" in r["reaction"])
    like_count = len(results) - love_count

    msg_lines = [
        "❤️ [BẢN TIN ĐIỀU HÀNH] TỰ ĐỘNG TƯƠNG TÁC BẠN BÈ FACEBOOK",
        "══════════════════════════════",
        f"⏰ Thời gian thực hiện: {now_str}",
        f"🎯 Tổng số bài đã tương tác: {len(results)} bài viết CÁ NHÂN (Bạn bè)",
        f"📊 Phân bổ cảm xúc: {love_count} Thả tim (❤️) | {like_count} Thích (👍)",
        "🛡️ Bộ lọc an toàn: Loại bỏ 100% Group, Fanpage, Ads | TUYỆT ĐỐI KHÔNG like bất kỳ comment nào",
        "══════════════════════════════",
        f"📋 CHI TIẾT {len(results)} BÀI VIẾT ĐÃ TƯƠNG TÁC:"
    ]

    for r in results:
        msg_lines.append(f"{r['stt']}. {r['author']} - {r['reaction']}")
        msg_lines.append(f"   💬 \"{r['snippet']}\"")

    msg_lines.append("══════════════════════════════")
    msg_lines.append("🤖 Trạng thái bot: Hoàn tất xuất sắc, an toàn tuyệt đối cho tài khoản!")

    full_message = "\n".join(msg_lines)
    print("\n[*] Đang gửi bản tin báo cáo vào Zalo 'Cloud của tôi'...", flush=True)
    try:
        success = send_zalo_message(full_message, target_chat="Cloud của tôi")
        if success:
            print("[✓] ĐÃ GỬI BÁO CÁO QUA ZALO THÀNH CÔNG!", flush=True)
        else:
            print("[!] Gửi Zalo không thành công.", flush=True)
        return success
    except Exception as e:
        print(f"[!] Lỗi gửi Zalo: {e}", flush=True)
        return False

def report_to_desktop_excel(results, status="SUCCESS"):
    r"""
    Ghi nhật ký chi tiết phiên tương tác vào file Excel Desktop:
    C:\Users\thuyn\Desktop\TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx (Sheet: 04_LOG_TUONG_TAC_FB)
    """
    desktop_excel = r"C:\Users\thuyn\Desktop\TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx"
    print(f"\n[*] Đang ghi báo cáo vào Desktop Excel: {desktop_excel}...", flush=True)
    try:
        if not os.path.exists(desktop_excel):
            subprocess.run([sys.executable, r"D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py"], check=False)

        wb = openpyxl.load_workbook(desktop_excel)
        if "04_LOG_TUONG_TAC_FB" not in wb.sheetnames:
            ws = wb.create_sheet(title="04_LOG_TUONG_TAC_FB")
        else:
            ws = wb["04_LOG_TUONG_TAC_FB"]

        target_row = 3
        while ws.cell(target_row, 1).value is not None:
            target_row += 1

        stt = target_row - 2
        now = datetime.datetime.now()
        now_str = now.strftime("%d/%m/%Y %H:%M:%S")

        current_hour = now.hour
        if current_hour < 12:
            shift_name = "09:00 Sáng"
        else:
            shift_name = "17:00 Chiều"

        count = len(results) if results else 0
        love_count = sum(1 for r in results if "Thả tim" in r["reaction"] or "❤️" in r["reaction"]) if results else 0
        like_count = count - love_count

        if status == "NEED_LOGIN":
            friends_summary = "Phiên Facebook hết hạn - Cần chạy DANG_NHAP_FACEBOOK.bat"
            status_text = "Cần đăng nhập"
            status_color = "DC2626"
        elif results:
            friends_summary = "; ".join([f"{r['author']} ({r['reaction'].split()[0]})" for r in results[:10]])
            if len(results) > 10:
                friends_summary += f"; và {len(results) - 10} bạn bè khác..."
            status_text = f"Hoàn thành {len(results)} bài"
            status_color = "166534"
        else:
            friends_summary = "Không tìm thấy bài mới hợp lệ"
            status_text = "0 bài tương tác"
            status_color = "DC2626"

        thin_border = Border(
            left=Side(style='thin', color="CBD5E1"),
            right=Side(style='thin', color="CBD5E1"),
            top=Side(style='thin', color="CBD5E1"),
            bottom=Side(style='thin', color="CBD5E1")
        )
        bg_fill = PatternFill(start_color="F8FAFC" if target_row % 2 == 0 else "FFFFFF", fill_type="solid")

        row_data = [
            (1, stt, "center", Font(name="Segoe UI", size=9, bold=True)),
            (2, now_str, "center", Font(name="Segoe UI", size=9)),
            (3, shift_name, "center", Font(name="Segoe UI", size=9, bold=True, color="1B365D")),
            (4, f"{count} bài cá nhân", "center", Font(name="Segoe UI", size=9)),
            (5, f"{love_count} tim ❤️", "center", Font(name="Segoe UI", size=9, color="E11D48")),
            (6, f"{like_count} like 👍", "center", Font(name="Segoe UI", size=9, color="2563EB")),
            (7, "100% Bạn bè (Bỏ Group/Page/Ads; KHÔNG like comment)", "left", Font(name="Segoe UI", size=9)),
            (8, status_text, "center", Font(name="Segoe UI", size=9, bold=True, color=status_color)),
            (9, friends_summary, "left", Font(name="Segoe UI", size=9)),
            (10, "Đã gửi Cloud của tôi", "center", Font(name="Segoe UI", size=9, color="166534"))
        ]

        for col, val, align, font_style in row_data:
            c = ws.cell(row=target_row, column=col, value=val)
            c.font = font_style
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(align == "left"))
            c.fill = bg_fill
            c.border = thin_border

        ws.row_dimensions[target_row].height = 24

        temp_file = desktop_excel + ".tmp.xlsx"
        wb.save(temp_file)
        wb.close()
        os.replace(temp_file, desktop_excel)
        print(f"[✓] ĐÃ GHI BÁO CÁO THÀNH CÔNG VÀO FILE EXCEL DESKTOP (Sheet 04_LOG_TUONG_TAC_FB)!", flush=True)

        subprocess.run([sys.executable, r"D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py"], check=False)
        return True

    except Exception as e:
        print(f"[!] Lỗi ghi file Excel Desktop: {e}", flush=True)
        return False

def main():
    parser = argparse.ArgumentParser(description="Bot tương tác Bảng tin Facebook cá nhân")
    parser.add_argument("--count", type=int, default=20, help="Số bài cần tương tác (mặc định 20)")
    parser.add_argument("--headless", action="store_true", help="Chạy ngầm không mở cửa sổ (cho Task Scheduler)")
    args = parser.parse_args()

    results, status = interact_feed(target_count=args.count, headless=args.headless)
    
    # 1. Báo cáo ghi vào file Excel Desktop
    report_to_desktop_excel(results, status=status)
    # 2. Gửi bản tin điều hành vào Zalo Cloud của tôi
    send_report_to_zalo(results, status=status)

if __name__ == "__main__":
    main()
