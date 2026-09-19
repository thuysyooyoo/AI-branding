# -*- coding: utf-8 -*-
import os
import sys
import io
import json
import argparse
import datetime
import urllib.request
import urllib.parse
import pandas as pd

# Đảm bảo in tiếng Việt trên console không bị lỗi
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except AttributeError:
    pass

WORKSPACE_DIR = r"d:\AI AGENT THUY\AI branding"
GEMINI_KEY_ENV = "GEMINI_API_KEY"
SERPAPI_KEY_ENV = "SERPAPI_KEY"
DEFAULT_SERPAPI_KEY = "109915b282e15a7e54ad58cf559d34e21d8fc2afb67d82f6ae2cd98e2c767eb6"

def fetch_serpapi_search(query, api_key):
    """Gọi SerpAPI để lấy kết quả tìm kiếm Google"""
    params = {
        "engine": "google",
        "q": query,
        "hl": "vi",
        "gl": "vn",
        "api_key": api_key
    }
    query_string = urllib.parse.urlencode(params)
    url = f"https://serpapi.com/search.json?{query_string}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as response:
            result = json.loads(response.read().decode('utf-8'))
            organic = result.get("organic_results", [])
            formatted = []
            for res in organic[:3]:  # Lấy top 3 kết quả
                formatted.append({
                    "title": res.get("title", ""),
                    "link": res.get("link", ""),
                    "snippet": res.get("snippet", "")
                })
            return formatted
    except Exception as e:
        print(f"⚠️ Lỗi SerpAPI khi tìm '{query}': {e}")
        return []

def call_gemini(system_instruction, user_content, api_key):
    """Gọi Gemini 1.5 Flash API bằng thư viện chuẩn urllib của Python (Zero-dependency)"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}
    
    payload = {
        "contents": [
            {
                "parts": [
                    {"text": user_content}
                ]
            }
        ],
        "systemInstruction": {
            "parts": [
                {"text": system_instruction}
            ]
        },
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 8192
        }
    }
    
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=90) as response:
            res_data = json.loads(response.read().decode("utf-8"))
            text = res_data["candidates"][0]["content"]["parts"][0]["text"]
            return text
    except Exception as e:
        print(f"❌ Lỗi kết nối Gemini API: {e}")
        if hasattr(e, 'read'):
            print(e.read().decode("utf-8"))
        return None

def get_approved_topics(calendar_path):
    """Đọc tệp Lịch biên tập và lấy ra các tiêu đề bài viết đã được duyệt"""
    if not os.path.exists(calendar_path):
        print(f"⚠️ Không tìm thấy file lịch: {calendar_path}")
        return []
    
    try:
        df = pd.read_excel(calendar_path, skiprows=3)
        approved_topics = []
        
        status_col = None
        for col in df.columns:
            if "Trạng thái Duyệt" in str(col) or "Duyệt" in str(col):
                status_col = col
                break
                
        if not status_col:
            print("⚠️ Không tìm thấy cột trạng thái duyệt trong file lịch.")
            return []
            
        for _, row in df.iterrows():
            status = str(row[status_col]).strip()
            # Mọi dòng có trạng thái bắt đầu bằng "Duyệt Option"
            if status.startswith("Duyệt Option"):
                option_letter = status.replace("Duyệt Option", "").strip().upper() # "A", "B", "C"
                
                col_name = None
                if option_letter == "A":
                    col_name = "Option A (Chủ đề chính)"
                elif option_letter == "B":
                    col_name = "Option B (Dự phòng 1)"
                elif option_letter == "C":
                    col_name = "Option C (Dự phòng 2)"
                
                # Cố gắng tìm cột tương ứng
                if not col_name or col_name not in df.columns:
                    for col in df.columns:
                        if f"Option {option_letter}" in str(col) or f"Option{option_letter}" in str(col):
                            col_name = col
                            break
                
                if col_name and col_name in df.columns:
                    val = str(row[col_name]).strip()
                    if val and val.lower() != "nan" and val.lower() != "none":
                        approved_topics.append({
                            "title": val,
                            "pillar": str(row.get("Pillar", "N/A")).strip(),
                            "option": f"Option {option_letter}"
                        })
        return approved_topics
    except Exception as e:
        print(f"❌ Lỗi khi đọc file lịch: {e}")
        return []

def get_new_laws(legal_path, days=7):
    """Lấy các quy định pháp lý mới cập nhật trong số ngày chỉ định"""
    if not os.path.exists(legal_path):
        print(f"⚠️ Không tìm thấy file văn bản pháp luật: {legal_path}")
        return []
        
    try:
        df = pd.read_excel(legal_path)
        new_laws = []
        
        date_col = None
        for col in df.columns:
            if "Cập Nhật" in str(col) or "cập nhật" in str(col):
                date_col = col
                break
                
        if not date_col:
            print("⚠️ Không tìm thấy cột ngày cập nhật, lấy 5 văn bản cuối cùng làm mẫu.")
            tail_df = df.tail(5)
            for _, row in tail_df.iterrows():
                new_laws.append(row.to_dict())
            return new_laws
            
        today = datetime.datetime.now().date()
        for _, row in df.iterrows():
            date_val = row[date_col]
            parsed_date = None
            if pd.notna(date_val):
                if isinstance(date_val, (datetime.date, datetime.datetime)):
                    parsed_date = date_val
                else:
                    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                        try:
                            parsed_date = datetime.datetime.strptime(str(date_val).strip(), fmt).date()
                            break
                        except ValueError:
                            continue
            
            if parsed_date:
                if isinstance(parsed_date, datetime.datetime):
                    parsed_date = parsed_date.date()
                delta = (today - parsed_date).days
                if delta <= days:
                    new_laws.append(row.to_dict())
                    
        if not new_laws:
            print("💡 Không tìm thấy văn bản pháp luật mới trong 7 ngày qua, lấy 3 văn bản cuối cùng làm mẫu.")
            tail_df = df.tail(3)
            for _, row in tail_df.iterrows():
                new_laws.append(row.to_dict())
                
        return new_laws
    except Exception as e:
        print(f"❌ Lỗi khi đọc file văn bản pháp luật: {e}")
        return []

def get_new_seo_keywords(seo_path, days=7):
    """Lấy các từ khóa SEO và nỗi đau mới cập nhật trong số ngày chỉ định"""
    if not os.path.exists(seo_path):
        print(f"⚠️ Không tìm thấy file từ khóa SEO: {seo_path}")
        return []
        
    try:
        # Tiêu đề cột nằm ở dòng 5 (skip 4 dòng đầu)
        df = pd.read_excel(seo_path, sheet_name="Danh sách Từ Khóa SEO", skiprows=4)
        new_keywords = []
        
        date_col = None
        for col in df.columns:
            if "Ngày cập nhật" in str(col) or "Cập nhật" in str(col) or "date" in str(col).lower():
                date_col = col
                break
                
        if not date_col:
            print("⚠️ Không tìm thấy cột ngày cập nhật trong file SEO, lấy 5 từ khóa cuối cùng.")
            tail_df = df.tail(5)
            for _, row in tail_df.iterrows():
                new_keywords.append(row.to_dict())
            return new_keywords
            
        today = datetime.datetime.now().date()
        for _, row in df.iterrows():
            date_val = row[date_col]
            parsed_date = None
            if pd.notna(date_val):
                if isinstance(date_val, (datetime.date, datetime.datetime)):
                    parsed_date = date_val
                else:
                    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
                        try:
                            parsed_date = datetime.datetime.strptime(str(date_val).strip(), fmt).date()
                            break
                        except ValueError:
                            continue
            
            if parsed_date:
                if isinstance(parsed_date, datetime.datetime):
                    parsed_date = parsed_date.date()
                delta = (today - parsed_date).days
                if delta <= days:
                    new_keywords.append(row.to_dict())
                    
        if not new_keywords:
            print("💡 Không tìm thấy từ khóa SEO mới trong 7 ngày qua, lấy 3 từ khóa cuối cùng làm mẫu.")
            # Loại bỏ dòng tổng hợp ở cuối nếu có
            clean_df = df[df['STT'].apply(lambda x: str(x).isdigit() or isinstance(x, int) or (isinstance(x, float) and x.is_integer()))]
            tail_df = clean_df.tail(3)
            for _, row in tail_df.iterrows():
                new_keywords.append(row.to_dict())
                
        return new_keywords
    except Exception as e:
        print(f"❌ Lỗi khi đọc file SEO: {e}")
        return []

def read_file(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""

def write_file(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def run_update_content_pillars(dry_run=False, api_key=""):
    print("="*60)
    print("🔄 BẮT ĐẦU CẬP NHẬT & LÀM DÀY KHO Ý TƯỞNG CONTENT PILLARS")
    print("="*60)
    
    # 1. Định nghĩa các đường dẫn tệp
    content_pillars_path = os.path.join(WORKSPACE_DIR, "knowledge", "content_pillars.md")
    calendar_path = os.path.join(WORKSPACE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
    legal_path = os.path.join(WORKSPACE_DIR, "reports", "Danh_sach_Van_ban_Phap_luat_XNK.xlsx")
    seo_path = os.path.join(WORKSPACE_DIR, "reports", "SEO_Keyword_Trends_Eureka.xlsx")
    daily_notes_path = os.path.join(WORKSPACE_DIR, "scratch", "daily_notes.md")
    
    # 2. Đọc dữ liệu cục bộ
    print("\n[1] Đang đọc dữ liệu từ các file nguồn...")
    
    current_pillars = read_file(content_pillars_path)
    if not current_pillars:
        print("❌ Lỗi: Không thể đọc file content_pillars.md hiện tại.")
        return
        
    daily_notes = get_daily_notes = read_file(daily_notes_path)
    approved_topics = get_approved_topics(calendar_path)
    new_laws = get_new_laws(legal_path, days=7)
    new_seo_keywords = get_new_seo_keywords(seo_path, days=7)
    
    print(f"👉 Số bài đã duyệt xuất bản phát hiện: {len(approved_topics)}")
    print(f"👉 Số văn bản pháp luật mới cập nhật: {len(new_laws)}")
    print(f"👉 Số từ khóa SEO/nỗi đau mới cập nhật: {len(new_seo_keywords)}")
    print(f"👉 Độ dài tệp daily notes thu được: {len(daily_notes)} ký tự")
    
    # 2.5 Đọc xu hướng trực tuyến qua SerpAPI
    print("\n[1.5] Đang quét các nguồn trực tuyến (SerpAPI) theo định vị của Agent 01...")
    serpapi_key = os.environ.get(SERPAPI_KEY_ENV, DEFAULT_SERPAPI_KEY)
    
    online_trends_text = ""
    if serpapi_key:
        # Pillar 1 query
        query_p1 = '("AI Agent" OR "B2B sales automation" OR "n8n") (site:producthunt.com OR site:youtube.com OR site:reddit.com/r/n8n)'
        print(f"👉 Quét xu hướng trực tuyến Pillar 1: {query_p1}")
        res_p1 = fetch_serpapi_search(query_p1, serpapi_key)
        
        # Pillar 3 query
        query_p3 = '("KPI phòng Sales" OR "quản trị sales B2B" OR "credit scoring" OR "đàm phán mua hàng Trung Quốc") (site:cafef.vn OR site:brandsvietnam.com OR site:hbr.org)'
        print(f"👉 Quét xu hướng trực tuyến Pillar 3: {query_p3}")
        res_p3 = fetch_serpapi_search(query_p3, serpapi_key)
        
        if res_p1:
            online_trends_text += "\n--- XU HƯỚNG TRỰC TUYẾN PILLAR 1 (AI & AUTOMATION) ---\n"
            for i, r in enumerate(res_p1, 1):
                online_trends_text += f"{i}. Tiêu đề: {r['title']}\n   Tóm tắt: {r['snippet']}\n   Nguồn: {r['link']}\n"
        
        if res_p3:
            online_trends_text += "\n--- XU HƯỚNG TRỰC TUYẾN PILLAR 3 (MANAGEMENT & B2B SALES) ---\n"
            for i, r in enumerate(res_p3, 1):
                online_trends_text += f"{i}. Tiêu đề: {r['title']}\n   Tóm tắt: {r['snippet']}\n   Nguồn: {r['link']}\n"
    
    if not online_trends_text:
        online_trends_text = "Không thu được dữ liệu xu hướng trực tuyến mới qua SerpAPI. Sử dụng tri thức cập nhật hiện tại về AI (n8n, CRM) và Quản trị Sales B2B."

    # 3. Chuẩn bị thông tin đầu vào cho LLM
    approved_text = ""
    for idx, item in enumerate(approved_topics, 1):
        approved_text += f"{idx}. {item['title']} (Pillar: {item['pillar']}, Trạng thái: {item['option']})\n"
        
    laws_text = ""
    for idx, item in enumerate(new_laws, 1):
        laws_text += f"- Văn bản {idx}: Mã: {item.get('Mã Số Văn Bản', 'N/A')} | Tên: {item.get('Tên Văn Bản', 'N/A')} | Loại: {item.get('Loại Văn Bản', 'N/A')} | Phân Tích Sơ Bộ: {item.get('Phân Tích Sơ Bộ', 'N/A')} | So Sánh Thay Đổi: {item.get('So Sánh Thay Đổi', 'N/A')}\n"
        
    seo_text = ""
    for idx, item in enumerate(new_seo_keywords, 1):
        seo_text += f"- Từ khóa {idx}: {item.get('Từ khóa chính (SEO Keywords)', 'N/A')} | Chủ đề đề xuất: {item.get('Chủ đề / Xu hướng nội dung', 'N/A')} | Nỗi đau khách hàng: {item.get('Nỗi đau & Nhu cầu của khách hàng', 'N/A')} | GĐKD định hướng: {item.get('Định hướng truyền thông nổi bật\n(GĐKD đề xuất)', 'N/A')}\n"

    # 4. Thiết lập System Instruction & User Prompt cho Gemini
    # 4. Thiết lập System Instruction & User Prompt cho Gemini
    system_instruction = """Bạn là Tác tử Nghiên cứu Thị trường & Quản trị Tri thức thuộc hệ thống PBAN của chị Nguyễn Thị Ngọc Thúy.
Nhiệm vụ của bạn là bảo trì và cập nhật tệp tri thức `content_pillars.md` hàng tuần theo Kiến trúc Nội dung 3 Layer mới (gồm 4 Chuyên mục nội dung: CS, TN, HT, FF):

1. DỌN DẸP Ý TƯỞNG ĐÃ XUẤT BẢN:
   - Đối chiếu danh sách các bài viết đã duyệt đăng tuần qua với danh sách "Sub-topics chi tiết" của 4 Chuyên mục (CS, TN, HT, FF) trong `content_pillars.md`.
   - Nếu phát hiện sub-topic nào trong danh sách đã được triển khai (so khớp ngữ nghĩa thông minh), hãy XÓA sub-topic đó khỏi Chuyên mục tương ứng để tránh bị trùng lặp ở các chu kỳ sau.

2. LÀM DÀY Ý TƯỞNG MỚI THỰC CHIẾN:
   - Dựa trên dữ liệu nguồn mới thu thập được (văn bản pháp luật mới, từ khóa/nỗi đau SEO mới, câu chuyện nhật ký daily_notes, và các xu hướng/tin tức trực tuyến), hãy sinh thêm 3-5 sub-topics chi tiết, cụ thể, thực chiến mới cho các Chuyên mục phù hợp:
     + Chuyên mục 1 (`CS` - Case Study): Bóc tách tình huống thực tế hoặc giả định mở (đàm phán, rủi ro hợp đồng, pháp lý hải quan, xử lý sự cố nhân sự/dòng tiền).
     + Chuyên mục 2 (`TN` - Tin Ngành): Phân tích chính sách, nghị định, thông tư hải quan mới & biến động thị trường theo phương pháp Problem-First (tập trung vào NỖI ĐAU & RỦI RO doanh nghiệp).
     + Chuyên mục 3 (`HT` - Học Thực chiến): Hướng dẫn 5 phút, quy trình A-Z, checklist, mẹo nghề, công thức tính toán, tư duy quản trị vận hành logistics/XNK.
     + Chuyên mục 4 (`FF` - Funfact & Insight): Sự thật bất ngờ, câu chuyện lịch sử hàng hải/thương mại, bài học từ nhân vật kiệt xuất, góc nhìn quốc tế.
   - Mỗi Chuyên mục duy trì tối đa 12 sub-topics chi tiết (được đánh số liên tục từ 1 trở đi). Nếu vượt quá, ưu tiên giữ lại các sub-topics mới và các ý tưởng cũ chưa khai thác có giá trị cao.

YÊU CẦU ĐẦU RA BẮT BUỘC:
- Trả về toàn bộ nội dung của tệp `content_pillars.md` đã cập nhật dưới định dạng Markdown hoàn chỉnh.
- Tuyệt đối giữ nguyên tiêu đề lớn, mô tả, định hướng format và cấu trúc định dạng của 3 Layer trong file. Chỉ cập nhật danh sách các Sub-topics chi tiết của 4 Chuyên mục (`CS`, `TN`, `HT`, `FF`).
- Không thêm bất kỳ lời dẫn giải thừa nào ngoài nội dung Markdown hoàn chỉnh của tệp."""

    user_content = f"""Dưới đây là các dữ liệu nguồn hiện tại:

---
[FILE KNOWLEDGE HIỆN TẠI (content_pillars.md)]
{current_pillars}

---
[DANH SÁCH BÀI ĐÃ DUYỆT ĐĂNG TUẦN QUA (CẦN XÓA NẾU SEMANTIC MATCH)]
{approved_text if approved_text else "Không có bài nào được duyệt trong tuần qua."}

---
[VĂN BẢN PHÁP LUẬT XNK MỚI CẬP NHẬT (NGUỒN CHO CS & TN)]
{laws_text if laws_text else "Không có văn bản mới."}

---
[XU HƯỚNG TỪ KHÓA & NỖI ĐAU SEO MỚI CẬP NHẬT (NGUỒN CHO TN & HT)]
{seo_text if seo_text else "Không có từ khóa/nỗi đau mới."}

---
[NHẬT KÝ CÔNG VIỆC DAILY NOTES MỚI (NGUỒN CHO CS & FF)]
{daily_notes if daily_notes else "Không có nhật ký mới."}

---
[XU HƯỚNG & TIN TỨC TRỰC TUYẾN TỪ CÁC KÊNH NGUỒN CỦA AGENT 01 (NGUỒN CHO CS, HT & FF)]
{online_trends_text}

---
Hãy thực hiện dọn dẹp các sub-topic đã xuất bản, bổ sung các sub-topic mới chất lượng, duy trì tối đa 12 sub-topics cho mỗi Chuyên mục (CS, TN, HT, FF), đánh số liên tục từ 1 trở đi. Trả về toàn bộ nội dung file content_pillars.md mới."""

    # 5. Gọi Gemini API
    success = False
    clean_markdown = None
    
    # Thử lấy API key từ môi trường nếu không truyền vào
    if not api_key:
        api_key = os.environ.get(GEMINI_KEY_ENV)
        
    if api_key:
        print("\n[2] Đang gọi Gemini API để phân tích và cập nhật tri thức...")
        response_text = call_gemini(system_instruction, user_content, api_key)
        if response_text:
            # Bóc tách Markdown block nếu có
            clean_markdown = response_text.strip()
            if clean_markdown.startswith("```markdown"):
                clean_markdown = clean_markdown[11:]
            elif clean_markdown.startswith("```"):
                clean_markdown = clean_markdown[3:]
            if clean_markdown.endswith("```"):
                clean_markdown = clean_markdown[:-3]
            clean_markdown = clean_markdown.strip()
            success = True
        else:
            print("⚠️ Lỗi: Không thể kết nối hoặc gọi Gemini API. Chuyển sang chế độ thủ công.")
    else:
        print("💡 Chế độ chạy thủ công (Manual Mode): Không cung cấp GEMINI_API_KEY.")

    if not success:
        # Xuất file prompt thủ công
        prompt_output_path = os.path.join(WORKSPACE_DIR, "scratch", "prompt_update_pillars.md")
        manual_instructions = f"""# 📑 PROMPT CẬP NHẬT KHO Ý TƯỞNG CONTENT PILLARS (THỦ CÔNG)
Hãy copy toàn bộ nội dung phần PROMPT bên dưới, dán vào cửa sổ chat của AI (như Gemini, Claude hoặc ChatGPT) để thực hiện cập nhật.
Sau khi AI phản hồi nội dung file Markdown mới, hãy copy toàn bộ nội dung đó và dán đè vào file:
`{content_pillars_path}`

---
## PROMPT BẮT ĐẦU DƯỚI ĐÂY (HÃY COPY PHẦN NÀY)
---

System Instruction:
{system_instruction}

User Input:
{user_content}
"""
        write_file(prompt_output_path, manual_instructions)
        print(f"🟢 Đã tạo file Prompt thủ công tại: {prompt_output_path}")
        print("👉 Vui lòng mở file này, copy nội dung gửi cho AI và dán kết quả vào file content_pillars.md.")
        print("="*60)
        return

    # 6. Ghi đè hoặc hiển thị Dry-run
    if dry_run:
        print("\n⚡ CHẾ ĐỘ CHẠY THỬ (DRY-RUN): Không ghi đè lên file vật lý.")
        print("\n--- NỘI DUNG FILE MỚI ĐỀ XUẤT ---")
        print(clean_markdown[:2000])
        print("\n...\n[Đã ẩn bớt phần sau]")
        print("\n---------------------------------")
        print("\n🟢 Kết thúc Dry-run.")
    else:
        print(f"\n[3] Đang lưu file cập nhật tại: {content_pillars_path}")
        write_file(content_pillars_path, clean_markdown)
        print("🟢 Đã cập nhật content_pillars.md thành công!")
    print("="*60)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Tự động dọn dẹp và làm dày tệp Content Pillars của hệ thống PBAN.")
    parser.add_argument("--dry-run", action="store_true", help="Chạy thử nghiệm hiển thị kết quả, không lưu file.")
    parser.add_argument("--key", default="", help="Gemini API Key.")
    
    args = parser.parse_args()
    run_update_content_pillars(dry_run=args.dry_run, api_key=args.key)
