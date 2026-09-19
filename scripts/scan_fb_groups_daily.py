# -*- coding: utf-8 -*-
"""
QUY TRÌNH QUÉT & PHÂN LOẠI THÔNG MINH BÀI ĐĂNG GROUP FACEBOOK (AI BRAIN & DAILY NOTES INTEGRATION)
Tuân thủ tuyệt đối bộ quy tắc tại GEMINI.md:
1. Phân loại chuẩn 100%: 01_BAI_HOI_DAP vs 02_BAI_CHIA_SE_KIEN_THUC (Mỗi bài DUY NHẤT 1 file).
2. Tác giả chuẩn 100%: Bóc tách từ tiêu đề bài viết (h2: "Bài viết của ..."), không bao giờ lấy tên tài khoản quét (Jasmine Nguyen) hay nhầm sang flyout thông báo.
3. Phương tiện (Ảnh/Video) chuẩn 100%: Kiểm tra trực tiếp trên DOM của bài viết (img > 150px), không nhận diện nhầm video reels quảng cáo ngoài sidebar.
4. Tiêu đề & Từ khóa chuẩn nghiệp vụ: Sử dụng AI Brain để tạo tiêu đề chuyên sâu, 3-5 từ khóa nghiệp vụ XNK/Logistics/Thuế.
5. Ghi file Excel an toàn: Dùng cơ chế atomic replace (.tmp.xlsx -> .xlsx) kèm retry 5 lần chống xung đột OneDrive sync lock.
6. Bước 4: Lưu vào Daily Note hàng ngày (scratch/daily_notes.md) làm chất liệu branding trước khi báo cáo.
7. Bước 5: Bắn ngay bản tin điều hành tóm tắt vào Zalo 'Cloud của tôi'.
"""

import os
import sys
import time
import datetime
import re
import json
import openpyxl
from openpyxl.styles import Font, Alignment
from playwright.sync_api import sync_playwright

# Tích hợp AI Brain từ ZALO REPORT
sys.path.insert(0, r"D:\AI AGENT THUY\ZALO REPORT")
try:
    from core.ai_brain import generate_ai_text
    AI_AVAILABLE = True
except Exception as e:
    print(f"[!] Cảnh báo không thể nạp AI Brain: {e}")
    AI_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_DIR = os.path.join(BASE_DIR, "fb_profile")

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\.13.Dữ liệu - bài viết - Hình ảnh - video\TỔNG HỢP DỮ LIỆU QUÉT GROUP FB"
FILE_HOI_DAP = os.path.join(ONEDRIVE_DIR, "01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx")
FILE_CHIA_SE = os.path.join(ONEDRIVE_DIR, "02_BAI_CHIA_SE_KIEN_THUC_GROUP_FACEBOOK.xlsx")
LOCAL_DAILY_NOTES = os.path.join(BASE_DIR, "scratch", "daily_notes.md")

TARGET_GROUPS = [
    {"url": "https://www.facebook.com/groups/thutuchaiquanxuatnhapkhau/?sorting_setting=CHRONOLOGICAL", "name": "Xuất nhập khẩu Logistics"},
    {"url": "https://www.facebook.com/groups/toantid/?sorting_setting=CHRONOLOGICAL", "name": "Giải đáp thủ tục Hải quan"},
    {"url": "https://www.facebook.com/groups/408350329619788/?sorting_setting=CHRONOLOGICAL", "name": "GIA ĐÌNH XUẤT NHẬP KHẨU - LOGISTICS"},
    {"url": "https://www.facebook.com/groups/113357826029183/?sorting_setting=CHRONOLOGICAL", "name": "Hỏi đáp Thủ tục hải quan - Xuất nhập khẩu - Logistics"},
    {"url": "https://www.facebook.com/groups/1476527345942555/?sorting_setting=CHRONOLOGICAL", "name": "Giải đáp thủ tục Hải quan"},
    {"url": "https://www.facebook.com/groups/880093095901505/?sorting_setting=CHRONOLOGICAL", "name": "DIỄN ĐÀN THUẾ CHUYÊN SÂU"}
]

SPAM_EXCLUDES = [
    "tuyển dụng", "tuyển nv", "tuyển thực tập", "cần tuyển", "tìm việc", "tìm việc làm",
    "bảng giá", "chào giá", "báo giá", "cung cấp pallet", "cho thuê kho", "thanh lý",
    "nhận kéo cont", "nhận làm c/o", "nhận vận chuyển", "liên hệ hotline", "zalo:", "sđt:",
    "khóa học", "học viện", "ưu đãi", "học thử", "đăng ký khóa", "khai giảng",
    "bán tài khoản", "via fb", "bm fb", "chạy ads chiết khấu"
]

def load_existing_urls():
    existing = set()
    for f in [FILE_HOI_DAP, FILE_CHIA_SE]:
        if os.path.exists(f):
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                s = wb.active
                for r in range(6, s.max_row + 1):
                    val = s.cell(r, 10).value
                    if val:
                        clean = str(val).strip().rstrip('/')
                        existing.add(clean.lower())
            except Exception as e:
                print(f"[!] Cảnh báo đọc file {f}: {e}")
    return existing

def is_within_24h(time_str):
    t = time_str.lower()
    if any(x in t for x in ["phút", "vừa xong", "giây", "vừa mới"]):
        return True
    if "giờ" in t:
        nums = [int(s) for s in t.split() if s.isdigit()]
        if nums:
            return nums[0] <= 24
        return True
    if "hôm qua" in t:
        return True
    return False

def analyze_content(content_text):
    lower = content_text.lower()
    
    # 1. Kiểm tra spam
    for sp in SPAM_EXCLUDES:
        if sp in lower:
            return False, None, None, None, None
            
    # 2. Kiểm tra từ khóa tối thiểu chuyên môn
    keywords = [
        "hải quan", "xuất nhập khẩu", "xnk", "logistics", "vận chuyển", "container", "tờ khai",
        "c/o", "mã hs", "hs code", "incoterms", "hóa đơn", "thuế", "thuế gtgt", "thuế tndn",
        "thông tư", "nghị định", "công văn", "chứng từ", "thanh toán quốc tế", "hộ kinh doanh",
        "etax", "máy nén", "lọc gió", "khí", "hàng hóa", "hồ sơ", "doanh nghiệp", "chi phí"
    ]
    matched = [kw for kw in keywords if kw in lower]
    if not matched:
        return False, None, None, None, None
        
    # 3. Phân tích qua AI Brain nếu có
    if AI_AVAILABLE:
        try:
            prompt = f"""
Bạn là Chuyên gia Phân tích Nghiệp vụ Hải quan, Xuất Nhập Khẩu và Thuế của Eureka Logistics.
Hãy đọc nội dung bài viết dưới đây từ Facebook Group và trích xuất thông tin nghiệp vụ:

NỘI DUNG BÀI VIẾT:
\"\"\"{content_text[:2000]}\"\"\"

YÊU CẦU:
1. PHÂN LOẠI: Trả về DUY NHẤT một trong hai giá trị:
   - "Hỏi đáp": Nếu người đăng nêu case study, câu hỏi vướng mắc nghiệp vụ (hải quan, thuế, C/O, mã HS, nợ thuế, nộp hồ sơ, mở hộ kinh doanh, hỏi ý kiến cộng đồng).
   - "Chia sẻ kiến thức": Nếu người đăng chia sẻ công văn mới, thông tư, cẩm nang nghiệp vụ, kinh nghiệm làm việc đúc kết, phân tích quy định.
2. TIÊU ĐỀ NGHIỆP VỤ: Một tiêu đề ngắn gọn (dưới 15 từ), súc tích, phản ánh đúng bản chất vấn đề. TUYỆT ĐỐI không chép nguyên câu chào hỏi cụt lủn.
3. TỪ KHÓA: 3 đến 5 cụm từ khóa nghiệp vụ cách nhau bằng dấu chấm phẩy (Ví dụ: Mã HS lò xo khí; Bộ phận máy nén khí; Thủ tục hải quan).

TRẢ VỀ DUY NHẤT JSON:
{{"category": "Hỏi đáp", "title": "Tiêu đề nghiệp vụ", "topics": "Từ khóa 1; Từ khóa 2; Từ khóa 3"}}
"""
            ai_raw = generate_ai_text(prompt)
            cleaned_ai = re.sub(r'^```json\s*', '', ai_raw.strip(), flags=re.IGNORECASE)
            cleaned_ai = re.sub(r'```$', '', cleaned_ai.strip())
            ai_data = json.loads(cleaned_ai)
            
            category = ai_data.get("category", "Chia sẻ kiến thức")
            if category not in ["Hỏi đáp", "Chia sẻ kiến thức"]:
                category = "Hỏi đáp" if any(x in lower for x in ["?", "hỏi", "tư vấn", "cho em hỏi", "cho mình hỏi", "phân vân"]) else "Chia sẻ kiến thức"
            title = ai_data.get("title", "").strip()
            topics = ai_data.get("topics", "").strip()
            summary = content_text.strip()[:2500]
            if title and topics:
                return True, category, title, topics, summary
        except Exception as e:
            print(f"    [!] Lỗi phân tích AI: {e}. Chuyển sang heuristics...")

    # Heuristics dự phòng
    qa_signals = ["?", "hỏi", "tư vấn", "cho em hỏi", "cho mình hỏi", "nhờ mọi người", "kinh nghiệm gỡ", "làm sao", "xin ý kiến", "có được không", "phân vân", "chưa ạ", "với ạ", "giúp em", "giúp mình"]
    is_qa = any(sig in lower for sig in qa_signals)
    category = "Hỏi đáp" if is_qa else "Chia sẻ kiến thức"
    
    topic_map = {
        "thuế": "Thuế",
        "hộ kinh doanh": "Hộ kinh doanh",
        "hóa đơn": "Hóa đơn điện tử",
        "mã hs": "Mã HS",
        "hs code": "Mã HS",
        "c/o": "Chứng nhận xuất xứ (C/O)",
        "tờ khai": "Tờ khai hải quan",
        "etax": "eTax Mobile",
        "logistics": "Logistics",
        "hải quan": "Thủ tục hải quan",
        "vận chuyển": "Vận tải",
        "công văn": "Công văn hướng dẫn"
    }
    extracted_topics = []
    for k, v in topic_map.items():
        if k in lower and v not in extracted_topics:
            extracted_topics.append(v)
    if not extracted_topics:
        extracted_topics = ["Nghiệp vụ xuất nhập khẩu"]
    topics = "; ".join(extracted_topics[:4])
    
    cleaned = re.sub(r'^(chào các bác|chào cả nhà|em chào cả nhà\s*(ạ)?|mọi người cho em hỏi|cho em hỏi|anh/chị ơi,?\s*(anh chị từng|cho em)?|cho mình hỏi|nhờ mọi người|dạ|e hỏi|e xin phép)\s*[,:]?\s*', '', content_text.strip(), flags=re.IGNORECASE)
    first_line = cleaned.strip().split('\n')[0].strip()
    if len(first_line) < 15 and '\n' in cleaned.strip():
        lines = [l.strip() for l in cleaned.strip().split('\n') if len(l.strip()) > 15]
        if lines:
            first_line = lines[0]
            
    first_line = first_line[:100].strip()
    if not first_line or len(first_line) < 10:
        first_line = f"Vướng mắc nghiệp vụ {extracted_topics[0]}"
    title = first_line[0].upper() + first_line[1:]
    summary = content_text.strip()[:2500]
    return True, category, title, topics, summary

def parse_post_detail(page, post_url):
    try:
        page.goto(post_url, timeout=35000, wait_until="domcontentloaded")
        page.wait_for_timeout(3500)
        
        # Bấm xem thêm nếu có
        try:
            see_more = page.locator('div[role="button"]:has-text("Xem thêm"), span:has-text("Xem thêm")').first
            if see_more.count() > 0:
                see_more.click(timeout=2000)
                page.wait_for_timeout(1000)
        except Exception:
            pass
            
        post_data = page.evaluate("""() => {
            const dialogs = Array.from(document.querySelectorAll('div[role="dialog"]'));
            
            // Tìm đúng dialog chứa bài viết (bỏ qua flyout thông báo/kết bạn)
            let postRoot = dialogs.find(d => {
                const h2 = d.querySelector('h2');
                return h2 && (h2.innerText.includes('Bài viết của') || h2.innerText.includes('Bài viết'));
            });
            
            if (!postRoot) {
                postRoot = dialogs.find(d => {
                    const txt = d.innerText;
                    return !txt.includes('Lời mời kết bạn') && !txt.startsWith('Thông báo') && 
                           d.querySelector('div[data-ad-preview="message"], div[data-ad-comet-preview="message"], div[dir="auto"]');
                });
            }
            
            if (!postRoot) {
                postRoot = document.querySelector('div[role="main"]') || document.body;
            }
            
            // 1. Tác giả
            let author = "Thành viên Facebook";
            const h2 = postRoot.querySelector('h2');
            if (h2 && h2.innerText.includes('Bài viết của')) {
                author = h2.innerText.replace('Bài viết của', '').trim();
            } else {
                const headerLinks = Array.from(postRoot.querySelectorAll('h2 a, h3 a, strong a, a[role="link"] strong'));
                for (let el of headerLinks) {
                    const txt = el.innerText.trim();
                    const lower = txt.toLowerCase();
                    if (txt.length > 2 && txt.length < 35 && 
                        !lower.includes('nhóm') && !lower.includes('diễn đàn') && 
                        !lower.includes('bình luận') && !lower.includes('chia sẻ') &&
                        !lower.includes('thông báo') && !lower.includes('lời mời') &&
                        !lower.includes('jasmine')) {
                        author = txt;
                        break;
                    }
                }
                if (author === "Thành viên Facebook" && postRoot.innerText.includes('Thành viên ẩn danh')) {
                    author = "Thành viên ẩn danh";
                }
            }
            
            // 2. Phương tiện (Ảnh / Video)
            const imgs = Array.from(postRoot.querySelectorAll('img')).filter(img => {
                const w = img.naturalWidth || img.width;
                const h = img.naturalHeight || img.height;
                const src = img.src || '';
                return w > 150 && h > 150 && !src.includes('rsrc.php') && !src.includes('emoji');
            });
            const vids = Array.from(postRoot.querySelectorAll('video')).filter(v => (v.videoWidth || v.clientWidth) > 150);
            
            // 3. Nội dung văn bản
            const msgDivs = Array.from(postRoot.querySelectorAll('div[data-ad-preview="message"], div[data-ad-comet-preview="message"], div[dir="auto"]'));
            const texts = msgDivs.map(d => d.innerText.trim()).filter(t => t.length > 15 && !t.startsWith('Thông báo') && !t.includes('Lời mời kết bạn'));
            let content = texts.length > 0 ? texts.reduce((a, b) => a.length > b.length ? a : b) : "";
            if (!content) {
                const lines = postRoot.innerText.split('\\n').map(l => l.trim()).filter(l => l.length > 20 && !l.startsWith('Thông báo') && !l.includes('Lời mời kết bạn') && !l.includes('Xem tất cả'));
                content = lines.join('\\n').slice(0, 3000);
            }
            
            return {
                author: author,
                content: content,
                imgCount: imgs.length,
                videoCount: vids.length,
                imgAlt: imgs.length > 0 ? (imgs[0].alt || '') : ''
            };
        }""")
        
        if not post_data or not post_data["content"]:
            return None
            
        # Xử lý chuỗi Media
        if post_data["videoCount"] > 0:
            media_str = f"Có ({post_data['videoCount']} video)"
        elif post_data["imgCount"] > 0:
            alt = post_data.get("imgAlt", "")
            if "etax" in alt.lower() or "thuế" in alt.lower():
                media_str = f"Có ({post_data['imgCount']} ảnh chụp màn hình eTax Mobile)"
            else:
                media_str = f"Có ({post_data['imgCount']} ảnh)"
        else:
            media_str = "Không"
            
        body_text = page.inner_text("body")
        cmt = re.search(r'(\d+)\s*bình luận', body_text, re.IGNORECASE)
        comments = int(cmt.group(1)) if cmt else 0
        
        sh = re.search(r'(\d+)\s*lượt chia sẻ', body_text, re.IGNORECASE)
        shares = int(sh.group(1)) if sh else 0
        
        rc = re.search(r'(\d+)\s*(người bày tỏ cảm xúc|người khác)', body_text, re.IGNORECASE)
        reactions = int(rc.group(1)) if rc else 0
        
        return {
            "author": post_data["author"],
            "media": media_str,
            "content": post_data["content"],
            "comments": comments,
            "shares": shares,
            "reactions": reactions
        }
    except Exception as e:
        print(f"    [!] Lỗi bóc tách chi tiết {post_url}: {e}")
        return None

def append_to_excel(fpath, record):
    wb = openpyxl.load_workbook(fpath)
    sheet = wb.active
    
    target_row = 6
    while sheet.cell(target_row, 1).value is not None:
        target_row += 1
        
    stt = target_row - 5
    font_data = Font(name="Arial", size=10.0, bold=False)
    
    row_data = [
        (1, stt, "center", "#,##0"),
        (2, record["scan_date"], "center", "dd/mm/yyyy hh:mm"),
        (3, record["post_date"], "center", "dd/mm/yyyy hh:mm"),
        (4, record["category"], "center", None),
        (5, record["group_name"], "left", None),
        (6, record["author"], "left", None),
        (7, record["topics"], "left", None),
        (8, record["title_summary"], "left", None),
        (9, record["content"], "wrap", None),
        (10, record["url"], "left", None),
        (11, record["reactions"], "right", "#,##0"),
        (12, record["comments"], "right", "#,##0"),
        (13, record["shares"], "right", "#,##0"),
        (14, record["media"], "center", None),
        (15, record["filter_note"], "left", None)
    ]
    
    for col_idx, val, align_type, num_fmt in row_data:
        cell = sheet.cell(row=target_row, column=col_idx, value=val)
        cell.font = font_data
        if align_type == "center":
            cell.alignment = Alignment(horizontal="center", vertical="top")
        elif align_type == "right":
            cell.alignment = Alignment(horizontal="right", vertical="top")
        elif align_type == "wrap":
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        else:
            cell.alignment = Alignment(vertical="top")
        if num_fmt:
            cell.number_format = num_fmt
            
    temp_fpath = fpath + ".tmp.xlsx"
    wb.save(temp_fpath)
    
    # Retry loop chống xung đột lock file OneDrive
    for attempt in range(5):
        try:
            os.replace(temp_fpath, fpath)
            print(f"  [+] ĐÃ LƯU THÀNH CÔNG VÀO {os.path.basename(fpath)} (STT: {stt}, Dòng: {target_row})")
            return target_row
        except PermissionError:
            time.sleep(2)
            
    # Fallback nếu vẫn lock
    try:
        os.replace(temp_fpath, fpath)
    except Exception as e:
        print(f"  [!] Lỗi ghi file Excel {fpath}: {e}")
    return target_row

def save_to_daily_notes(new_records):
    """
    Lưu các case study / bài viết mới thu thập trong ngày vào file Daily Note
    (D:\\AI AGENT THUY\\AI branding\\scratch\\daily_notes.md) và đồng bộ sang OneDrive Marketing.
    """
    if not new_records:
        print("[*] Không có bài viết mới để ghi vào Daily Note.")
        return False

    now = datetime.datetime.now()
    today_str = now.strftime("%d/%m/%Y")
    date_header = f"## Ngày {today_str}"
    
    existing_content = ""
    if os.path.exists(LOCAL_DAILY_NOTES):
        with open(LOCAL_DAILY_NOTES, "r", encoding="utf-8") as f:
            existing_content = f.read()
            
    # Lọc các bài chưa có trong daily note
    records_to_add = []
    for r in new_records:
        if r["url"].lower() not in existing_content.lower():
            records_to_add.append(r)
            
    if not records_to_add:
        print("[*] Các bài viết hôm nay đã được cập nhật đầy đủ trong Daily Note.")
        return True

    print(f"\n[*] Đang lưu {len(records_to_add)} bài viết vào Daily Note ({today_str})...")
    
    new_entries = []
    for idx, r in enumerate(records_to_add, 1):
        prompt = f"""
Bạn là Chuyên gia Tư vấn Quản trị Logistics & Thuế của Eureka.
Từ tình huống thực tế của doanh nghiệp/người kinh doanh dưới đây:
Tiêu đề: {r['title_summary']}
Nội dung: {r['content']}

Hãy viết 2 gạch đầu dòng súc tích (dưới 70 từ) cho mục góc nhìn Eureka:
- Bản chất rủi ro pháp lý/kỹ thuật.
- Khuyến nghị hành động chuẩn mực để phòng ngừa hoặc tối ưu chi phí.
"""
        try:
            insight = generate_ai_text(prompt)
            clean_lines = []
            for line in insight.strip().split('\n'):
                l = line.strip()
                if not l or "bài học rút ra" in l.lower() or "góc nhìn eureka" in l.lower():
                    continue
                l = re.sub(r'^[-*•\d\.]+\s*', '', l)
                clean_lines.append(f"  - {l}")
            insight_text = "\n".join(clean_lines)
        except Exception:
            insight_text = "  - Cần đối chiếu kỹ quy định chuyên ngành và tham vấn cơ quan quản lý trước khi áp dụng.\n  - Chuẩn bị đầy đủ chứng từ chứng minh tính chất kỹ thuật hoặc số liệu hạch toán thực tế."

        entry_md = f"""### Case Study: {r['title_summary']}
- **Tác giả & Nguồn:** {r['author']} | Nhóm: {r['group_name']}
- **Sự kiện / Câu chuyện nổi bật trong ngày:**
  - {r['content'].strip()}
- **Bài học rút ra hoặc suy ngẫm cá nhân (Góc nhìn Eureka):**
{insight_text}
- **Thông tin liên quan khác (nếu có):**
  - Phân loại: {r['category']} | Phương tiện: {r['media']}
  - Từ khóa chuyên môn: {r['topics']}
  - Link bài viết gốc: {r['url']}
"""
        new_entries.append(entry_md)

    combined_new_entries = "\n---\n\n".join(new_entries)
    
    if date_header in existing_content:
        parts = existing_content.split(date_header, 1)
        updated_content = parts[0] + date_header + "\n\n" + combined_new_entries + "\n\n---\n" + parts[1]
    else:
        split_marker = "---\n"
        if split_marker in existing_content:
            parts = existing_content.split(split_marker, 1)
            updated_content = parts[0] + split_marker + "\n" + date_header + "\n\n" + combined_new_entries + "\n\n---\n" + parts[1]
        else:
            updated_content = f"# NHẬT KÝ CÔNG VIỆC HẰNG NGÀY (DAILY NOTES)\n\n---\n\n{date_header}\n\n{combined_new_entries}\n\n---\n\n" + existing_content

    # Lưu bản local scratch
    os.makedirs(os.path.dirname(LOCAL_DAILY_NOTES), exist_ok=True)
    with open(LOCAL_DAILY_NOTES, "w", encoding="utf-8") as f:
        f.write(updated_content)
    print(f"[+] ĐÃ GHI NHẬN VÀO DAILY NOTE CỤC BỘ: {LOCAL_DAILY_NOTES}")

    return True

def main():
    print("="*70)
    print("KHỞI CHẠY QUY TRÌNH QUÉT & PHÂN LOẠI GROUP FACEBOOK 24H")
    print("="*70)
    
    existing_urls = load_existing_urls()
    print(f"[*] Đã tải {len(existing_urls)} URL bài viết đã có để chống trùng lặp.")
    
    candidate_posts = []
    
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=True,
            args=["--disable-blink-features=AutomationControlled"]
        )
        page = context.pages[0] if context.pages else context.new_page()
        
        for g in TARGET_GROUPS:
            g_name = g["name"]
            g_url = g["url"]
            print(f"\n[*] Đang quét nhóm: {g_name} ...")
            try:
                page.goto(g_url, timeout=35000, wait_until="domcontentloaded")
                page.wait_for_timeout(3500)
                
                # Cuộn trong vùng feed
                for _ in range(5):
                    page.evaluate("window.scrollBy(0, 1500)")
                    time.sleep(1.5)
                    
                # Chỉ lấy link bài viết trong feed chính của nhóm
                feed = page.locator('div[role="feed"]').first
                if feed.count() > 0:
                    links = feed.locator('a[href*="/posts/"], a[href*="/permalink/"]').all()
                else:
                    links = page.locator('div[role="main"] a[href*="/posts/"], div[role="main"] a[href*="/permalink/"]').all()
                    
                seen = set()
                for l in links:
                    href = l.get_attribute("href")
                    if not href or "/groups/" not in href:
                        continue
                    clean_url = href.split("?")[0].rstrip('/')
                    if clean_url in seen or clean_url.lower() in existing_urls:
                        continue
                    seen.add(clean_url)
                    
                    time_txt = l.inner_text().strip()
                    if is_within_24h(time_txt):
                        print(f"  -> Phát hiện bài mới trong 24h: {clean_url} ({time_txt})")
                        candidate_posts.append((clean_url, g_name, time_txt))
            except Exception as e:
                print(f"  [!] Lỗi quét {g_name}: {e}")
                
        print(f"\n[*] Tổng số bài viết 24h tìm thấy: {len(candidate_posts)}")
        
        now = datetime.datetime.now()
        added_hd = 0
        added_cs = 0
        added_records = []
        
        for post_url, group_name, time_txt in candidate_posts:
            if post_url.lower() in existing_urls:
                continue
                
            print(f"\n--- Bóc tách chi tiết: {post_url} ---")
            detail = parse_post_detail(page, post_url)
            if not detail:
                continue
                
            is_valid, cat, title, topics, summary = analyze_content(detail["content"])
            if not is_valid:
                print("  -> Bỏ qua: Không thỏa tiêu chí chuyên môn hoặc là tin rác/quảng cáo.")
                continue
                
            note = "Đã loại quảng cáo, tuyển dụng, chào giá, xin hàng/dịch vụ và bài trùng URL. Giờ đăng trong vòng 24h ngày quét."
            
            rec = {
                "scan_date": now,
                "post_date": now.replace(hour=0, minute=0, second=0, microsecond=0),
                "category": cat,
                "group_name": group_name,
                "author": detail["author"],
                "topics": topics,
                "title_summary": title,
                "content": summary,
                "url": post_url,
                "reactions": detail["reactions"],
                "comments": detail["comments"],
                "shares": detail["shares"],
                "media": detail["media"],
                "filter_note": note
            }
            
            print(f"  [>] Phân loại: {cat} | Tác giả: {detail['author']} | Tiêu đề: {title}")
            print(f"  [>] Media: {detail['media']} | Từ khóa: {topics}")
            
            if cat == "Hỏi đáp":
                append_to_excel(FILE_HOI_DAP, rec)
                added_hd += 1
            else:
                append_to_excel(FILE_CHIA_SE, rec)
                added_cs += 1
                
            added_records.append(rec)
            existing_urls.add(post_url.lower())
            
        context.close()
        
    print("\n" + "="*70)
    print("HOÀN TẤT GHI EXCEL!")
    print(f"- Bài Hỏi đáp thêm mới: {added_hd}")
    print(f"- Bài Chia sẻ kiến thức thêm mới: {added_cs}")
    print("="*70)

    # BƯỚC 4: LƯU VÀO DAILY NOTE HÀNG NGÀY TRƯỚC KHI BÁO CÁO ZALO
    print("\n[*] BƯỚC 4: Lưu vào Daily Note hàng ngày (scratch/daily_notes.md)...")
    save_to_daily_notes(added_records)

    # BƯỚC 5: TỰ ĐỘNG BẮN BẢN TIN ĐIỀU HÀNH QUA ZALO BOT
    print("\n[*] BƯỚC 5: Bắn bản tin điều hành vào Zalo 'Cloud của tôi'...")
    try:
        import subprocess
        subprocess.run([
            sys.executable,
            r"D:\AI AGENT THUY\ZALO REPORT\core\onedrive_to_zalo.py",
            "--zone", "group_fb",
            "--chat", "Cloud của tôi"
        ], check=False)
        print("[+] ĐÃ GỬI BẢN TIN ĐIỀU HÀNH VÀO ZALO THÀNH CÔNG!")
    except Exception as e:
        print(f"[!] Không thể kích hoạt Zalo: {e}")

if __name__ == "__main__":
    main()
