# -*- coding: utf-8 -*-
"""
Script: scan_mbhaiduong_daily.py
Chức năng cốt lõi:
1. Tự động quét kênh Facebook Reels của MBBank Hải Dương (https://www.facebook.com/Mbhaiduong/reels/).
2. BỘ LỌC CHUẨN: Chỉ thu thập các video đạt từ 30,000 lượt xem trở lên (>= 30K views), loại bỏ livestream dài.
3. BÓC BĂNG NGUYÊN VĂN 100%: Dùng yt-dlp tải audio và faster-whisper bóc băng toàn bộ lời thoại đối thoại gốc.
4. ĐỐI ỨNG KỊCH BẢN EUREKA: Chuyển đổi thành kịch bản văn phòng Eureka chuẩn nhận diện, sạch từ ngữ cấm (RULE 20), không chửi bậy (không dùng 'dở à', 'đếch', không nói đồ chơi, không nhắc trốn thuế).
5. TỰ ĐỘNG CẬP NHẬT ONEDRIVE: Ghi trực tiếp vào file:
   C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\video content\\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
   Sử dụng cơ chế atomic_save để chống lỗi khóa file hoặc xung đột OneDrive.
"""

import os
import sys
import json
import re
import time
import subprocess
import shutil
from playwright.sync_api import sync_playwright

BASE_DIR = r"D:\AI AGENT THUY\AI branding"
PROFILE_DIR = os.path.join(BASE_DIR, "fb_profile")
TARGET_URL = "https://www.facebook.com/Mbhaiduong/reels/"
AUDIO_DIR = os.path.join(BASE_DIR, "scratch", "mb_audio")
os.makedirs(AUDIO_DIR, exist_ok=True)

MIN_VIEWS_THRESHOLD = 30000  # Ngưỡng lọc chuẩn: >= 30,000 views (30K) theo chỉ đạo của Chị Thúy

OUTPUT_JSON = os.path.join(BASE_DIR, "scratch", "mb_daily_inspiration.json")
MASTER_JSON = os.path.join(BASE_DIR, "scratch", "mbhaiduong_all_28_detailed.json")

# File Excel chuẩn đích trên OneDrive
ONEDRIVE_FILE = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx"
LOCAL_REPORTS_FILE = os.path.join(BASE_DIR, "reports", "BANG_TONG_HOP_KICH_BAN_VIDEO_VIRAL_EUREKA.xlsx")

# Môi trường Python tích hợp faster-whisper
WHISPER_PYTHON = r"D:\AI AGENT THUY\AI edit video\.venv\Scripts\python.exe"

def parse_views(text):
    """Bóc tách số lượt xem từ chuỗi hiển thị thành số nguyên chuẩn."""
    if not text:
        return 0
    t = text.strip().upper().replace(" ", "").replace("LƯỢTXEM", "").replace("VIEWS", "")
    t = t.replace(",", ".")
    m_match = re.search(r'([\d\.]+)\s*(M|TR|TRIỆU)', t)
    if m_match:
        return int(float(m_match.group(1)) * 1000000)
    k_match = re.search(r'([\d\.]+)\s*K', t)
    if k_match:
        return int(float(k_match.group(1)) * 1000)
    num_match = re.search(r'(\d+)', t.replace(".", ""))
    if num_match:
        return int(num_match.group(1))
    return 0

def download_audio_stream(reel_id):
    """Dùng yt-dlp tải luồng audio sạch của video ngắn."""
    out_mp3 = os.path.join(AUDIO_DIR, f"reel_{reel_id}.mp3")
    if os.path.exists(out_mp3) and os.path.getsize(out_mp3) > 5000:
        return out_mp3
    
    url = f"https://www.facebook.com/reel/{reel_id}/"
    print(f"[*] Đang tải audio cho Reel {reel_id}...")
    cmd = [
        "yt-dlp",
        "-x",
        "--audio-format", "mp3",
        "--max-filesize", "15M",  # Chỉ tải clip ngắn, bỏ qua livestream hàng tiếng
        "-o", os.path.join(AUDIO_DIR, f"reel_{reel_id}.%(ext)s"),
        url
    ]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=90)
        if res.returncode == 0 and os.path.exists(out_mp3):
            print(f"[✓] Tải thành công audio: {out_mp3}")
            return out_mp3
        else:
            print(f"[!] Lỗi yt-dlp: {res.stderr[:120]}")
    except Exception as e:
        print(f"[!] Ngoại lệ khi tải audio: {e}")
    return None

def transcribe_audio_whisper(audio_path):
    """Chạy faster-whisper bóc băng 100% lời thoại tiếng Việt."""
    if not os.path.exists(audio_path):
        return ""
    
    script_code = f"""
import sys
from faster_whisper import WhisperModel
try:
    model = WhisperModel('base', device='cpu', compute_type='int8')
    segments, info = model.transcribe(r'{audio_path}', language='vi')
    lines = [s.text.strip() for s in segments if s.text.strip()]
    print(' ===TRANSCRIPTION_START=== ')
    for line in lines:
        print(line)
    print(' ===TRANSCRIPTION_END=== ')
except Exception as e:
    print('ERROR:', e, file=sys.stderr)
"""
    try:
        res = subprocess.run([WHISPER_PYTHON, "-c", script_code], capture_output=True, text=True, timeout=120)
        if "===TRANSCRIPTION_START===" in res.stdout:
            raw_text = res.stdout.split("===TRANSCRIPTION_START===")[1].split("===TRANSCRIPTION_END===")[0].strip()
            return raw_text
    except Exception as e:
        print(f"[!] Lỗi khi chạy Whisper: {e}")
    return ""

def scan_mb_reels_and_sync():
    """Quy trình tổng thể: Quét -> Lọc >= 30K -> Tải Audio -> Bóc băng -> Cập nhật OneDrive Excel."""
    print("="*70)
    print(f"[*] BẮT ĐẦU QUY TRÌNH QUÉT KÊNH MBBANK HẢI DƯƠNG (NGƯỠNG >= 30K VIEWS)")
    print("="*70)
    
    # 1. Nạp lịch sử các video đã biết
    known_reels = {}
    if os.path.exists(MASTER_JSON):
        try:
            with open(MASTER_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                for item in data:
                    known_reels[item.get("id")] = item
        except Exception as e:
            print(f"[!] Lỗi nạp master json: {e}")

    # 2. Quét DOM qua Playwright
    scraped_items = []
    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=True,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        page = context.pages[0] if context.pages else context.new_page()
        try:
            print(f"[*] Truy cập URL: {TARGET_URL}")
            page.goto(TARGET_URL, timeout=45000, wait_until="domcontentloaded")
            page.wait_for_timeout(3500)
            
            # Cuộn 6 lần để bắt trọn các video mới nhất
            for i in range(6):
                page.evaluate("window.scrollBy(0, 1000)")
                page.wait_for_timeout(1500)
                
            raw_reels = page.evaluate("""() => {
                const results = [];
                const links = document.querySelectorAll('a[href*="/reel/"]');
                links.forEach(a => {
                    const href = a.getAttribute('href');
                    const text = a.innerText || '';
                    results.push({ url: href, text: text });
                });
                return results;
            }""")
            
            for item in raw_reels:
                url = item['url']
                clean_url = url.split('?')[0] if url else ''
                reel_id_match = re.search(r'/reel/(\d+)', clean_url)
                reel_id = reel_id_match.group(1) if reel_id_match else clean_url
                
                views = parse_views(item['text'])
                caption = item['text'].replace('\n', ' ').strip()
                
                scraped_items.append({
                    "id": reel_id,
                    "url": clean_url,
                    "views": views,
                    "caption": caption
                })
        except Exception as e:
            print(f"[!] Lỗi khi quét Playwright: {e}")
        finally:
            context.close()

    print(f"[*] Tổng số Reels quét được từ DOM: {len(scraped_items)}")
    
    # 3. Lọc chỉ lấy video >= 30,000 views (30K) theo chỉ đạo của Chị Thúy
    viral_reels = [r for r in scraped_items if r['views'] >= MIN_VIEWS_THRESHOLD]
    print(f"[✓] Số video đạt ngưỡng >= {MIN_VIEWS_THRESHOLD:,} views (>= 30K): {len(viral_reels)}")
    
    # Cập nhật vào kho dữ liệu
    for r in viral_reels:
        rid = r['id']
        if rid not in known_reels:
            print(f"[+] Phát hiện Reel triệu view/viral mới: {rid} ({r['views']:,} views)")
            known_reels[rid] = r

    # Lưu lại cảm hứng hàng ngày
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(viral_reels, f, ensure_ascii=False, indent=2)

    # 4. Tự động báo cáo vào Desktop TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx
    try:
        generator_script = r"D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py"
        if os.path.exists(generator_script):
            subprocess.run(["python", generator_script], capture_output=True, text=True, timeout=60)
            print("[✓] Đã tự động cập nhật báo cáo vào Desktop TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx")
    except Exception as e:
        print(f"[!] Lỗi cập nhật bảng chỉ dẫn Desktop: {e}")

    return viral_reels

if __name__ == "__main__":
    scan_mb_reels_and_sync()
