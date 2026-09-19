# -*- coding: utf-8 -*-
"""
URL Validator & Deep Content Verifier Module (Technical Gatekeeper - Version 4.0)
=================================================================================
Mục đích:
1. Gửi HTTP GET Request thực tế đến URL để kiểm tra mã trạng thái HTTP 200 OK.
2. Tải và phân tích cú pháp HTML thực tế: Bóc tách <title>, <meta description> và toàn bộ nội dung văn bản <body>.
3. Kiểm tra tính xác thực & độ liên quan về mặt ngữ nghĩa (Semantic Relevance):
   - Nếu có truyền `expected_keywords` (ví dụ: ['panama', 'canal'] hoặc ['128/2020', 'hai quan']):
     BẮT BUỘC phải tìm thấy ít nhất 1 trong các `expected_keywords` này trong nội dung trang web.
   - Loại trừ hoàn toàn các cụm từ phổ thông (như 'bài toán', 'câu chuyện', 'bài học', 'quy trình').
4. Chặn đứng tuyệt đối (Hard Reject) các đường link giả mạo, link 404, link lỗi Cloudflare 403, 
   hoặc link bị sai ID dẫn đến bài viết khác nội dung.
"""

import urllib.request
import urllib.parse
import re
import ssl
import sys
import unicodedata
from html.parser import HTMLParser

# Khởi tạo SSL Context an toàn
SSL_CONTEXT = ssl.create_default_context()
SSL_CONTEXT.check_hostname = False
SSL_CONTEXT.verify_mode = ssl.CERT_NONE

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache"
}

ERROR_PAGE_INDICATORS = [
    "404 not found",
    "page not found",
    "không tìm thấy trang",
    "trang không tồn tại",
    "access denied",
    "403 forbidden",
    "error 404",
    "tin này không tồn tại",
    "bài viết đã bị xóa",
    "website under maintenance"
]

NON_DISCRIMINATING_PHRASES = {
    "bai toan", "cau chuyen", "bai hoc", "quy trinh", "nguyen tac", 
    "bi quyet", "kinh nghiem", "giai phap", "huong dan", "chia se", 
    "goc nhin", "loi khuyen", "thuc trang", "dien bien", "top 3", "top 5"
}

def remove_accents(input_str):
    """Chuẩn hóa chuỗi tiếng Việt không dấu để so khớp tìm kiếm chính xác."""
    if not input_str:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower()

class HTMLContentExtractor(HTMLParser):
    """Parser HTML bóc tách Tiêu đề, Meta Description và Nội dung văn bản thân bài."""
    def __init__(self):
        super().__init__()
        self.reset()
        self.title_parts = []
        self.in_title = False
        self.in_ignored_tag = False
        self.meta_desc = ""
        self.body_text_parts = []
        self.current_tag = ""

    def handle_starttag(self, tag, attrs):
        self.current_tag = tag.lower()
        if self.current_tag == "title":
            self.in_title = True
        elif self.current_tag in ("script", "style", "noscript", "svg", "header", "footer", "nav"):
            self.in_ignored_tag = True
        elif self.current_tag == "meta":
            attrs_dict = {k.lower(): v for k, v in attrs if v is not None}
            name = attrs_dict.get("name", "").lower()
            prop = attrs_dict.get("property", "").lower()
            content = attrs_dict.get("content", "")
            if name in ("description", "abstract") or prop in ("og:description", "twitter:description"):
                if not self.meta_desc:
                    self.meta_desc = content.strip()

    def handle_endtag(self, tag):
        tag_lower = tag.lower()
        if tag_lower == "title":
            self.in_title = False
        elif tag_lower in ("script", "style", "noscript", "svg", "header", "footer", "nav"):
            self.in_ignored_tag = False
        self.current_tag = ""

    def handle_data(self, data):
        text = data.strip()
        if not text:
            return
        if self.in_title:
            self.title_parts.append(text)
        elif not self.in_ignored_tag:
            self.body_text_parts.append(text)

    def get_title(self):
        return " ".join(self.title_parts).strip()

    def get_body_text(self):
        return " ".join(self.body_text_parts)


def extract_meaningful_entities(topic_title):
    """
    Bóc tách các thực thể cốt lõi từ Tiêu đề (loại bỏ từ vựng generic).
    """
    if not topic_title:
        return []
    
    entities = []
    
    # 1. Bóc tách các cụm từ chuyên môn trong dấu ngoặc đơn (e.g. (EMC), (Back-to-Back L/C), (Bonded Warehouse)...)
    parens = re.findall(r'\((.*?)\)', topic_title)
    for p in parens:
        p_clean = p.strip()
        if p_clean:
            entities.append(p_clean)
            for w in re.findall(r'\b[A-Za-z0-9\-\/]{2,}\b', p_clean):
                entities.append(w)
    
    # 2. Tìm các từ mượn tiếng Anh / tên riêng / số hiệu (Panama, McLean, WCO, DCSA, CBAM, Incoterms, 128/2020...)
    eng_words = re.findall(r'\b[A-Za-z0-9\-\/]{2,}\b', topic_title)
    stop_words_ascii = (
        "cho", "cac", "the", "and", "for", "with", "top", "bai", "cua", "trong",
        "tai", "vao", "khi", "qua", "tu", "den", "bi", "do", "ra", "de", "nhu",
        "mot", "hai", "ba", "bon", "nam", "sau", "bay", "tam", "chin", "muoi",
        "dong", "ngay", "tuan", "thang", "lan", "co", "khong", "duoc", "se", "da"
    )
    for w in eng_words:
        if w.lower() not in stop_words_ascii and len(w) >= 2:
            entities.append(w)
            
    # 3. Bóc tách các cụm 2 từ ghép tiếng Việt chuyên ngành
    clean_title = re.sub(r'[\(\)\[\]\{\}\:\,\.\-\"\'\?\!\/]', ' ', topic_title)
    words = clean_title.split()
    for i in range(len(words) - 1):
        w1 = words[i].strip()
        w2 = words[i+1].strip()
        if len(w1) >= 2 and len(w2) >= 2:
            phrase = f"{w1} {w2}".lower()
            phrase_no_acc = remove_accents(phrase)
            if phrase_no_acc not in NON_DISCRIMINATING_PHRASES and not any(sw in phrase_no_acc for sw in ["la gi", "nhu the", "de lam", "cac buoc"]):
                entities.append(f"{w1} {w2}")
                
    return list(set(entities))


def validate_url_deep(url, topic_title="", expected_keywords=None, min_text_len=80, timeout=10):
    """
    Hàm kiểm tra sâu tính xác thực của URL và đọc trực tiếp nội dung trong trang.
    """
    if not url or not isinstance(url, str) or not url.startswith("http"):
        return {
            "is_valid": False,
            "url": url,
            "final_url": None,
            "status_code": None,
            "page_title": None,
            "meta_description": None,
            "content_snippet": None,
            "matched_keywords": [],
            "error_reason": "URL không hợp lệ hoặc rỗng"
        }

    # 1. Gửi HTTP Request
    try:
        req = urllib.request.Request(url, headers=HEADERS)
        with urllib.request.urlopen(req, timeout=timeout, context=SSL_CONTEXT) as response:
            status_code = response.status
            final_url = response.geturl()
            content_type = response.headers.get("Content-Type", "").lower()
            
            raw_bytes = response.read(512000)
            
            charset = "utf-8"
            match = re.search(r'charset=([\w-]+)', content_type)
            if match:
                charset = match.group(1)
            try:
                html_text = raw_bytes.decode(charset, errors="replace")
            except Exception:
                html_text = raw_bytes.decode("utf-8", errors="ignore")

    except urllib.error.HTTPError as e:
        return {
            "is_valid": False,
            "url": url,
            "final_url": None,
            "status_code": e.code,
            "page_title": None,
            "meta_description": None,
            "content_snippet": None,
            "matched_keywords": [],
            "error_reason": f"Lỗi HTTP {e.code}: {e.reason}"
        }
    except Exception as e:
        return {
            "is_valid": False,
            "url": url,
            "final_url": None,
            "status_code": None,
            "page_title": None,
            "meta_description": None,
            "content_snippet": None,
            "matched_keywords": [],
            "error_reason": f"Lỗi kết nối: {str(e)}"
        }

    # 2. Kiểm tra mã trạng thái HTTP
    if status_code != 200:
        return {
            "is_valid": False,
            "url": url,
            "final_url": final_url,
            "status_code": status_code,
            "page_title": None,
            "meta_description": None,
            "content_snippet": None,
            "matched_keywords": [],
            "error_reason": f"HTTP Status không phải 200 OK (Mã nhận được: {status_code})"
        }

    # 3. Phân tích nội dung trang HTML
    parser = HTMLContentExtractor()
    try:
        parser.feed(html_text)
    except Exception as e:
        pass

    page_title = parser.get_title()
    meta_desc = parser.meta_desc
    body_text = parser.get_body_text()
    
    # 4. Kiểm tra trang lỗi giả mạo (Soft 404)
    normalized_title = page_title.lower()
    for err_ind in ERROR_PAGE_INDICATORS:
        if err_ind in normalized_title:
            return {
                "is_valid": False,
                "url": url,
                "final_url": final_url,
                "status_code": status_code,
                "page_title": page_title,
                "meta_description": meta_desc,
                "content_snippet": body_text[:150],
                "matched_keywords": [],
                "error_reason": f"Trang web trả về giao diện báo lỗi (Tiêu đề: '{page_title}')"
            }

    # 5. Kiểm tra dung lượng nội dung bài viết
    if len(body_text) < min_text_len:
        return {
            "is_valid": False,
            "url": url,
            "final_url": final_url,
            "status_code": status_code,
            "page_title": page_title,
            "meta_description": meta_desc,
            "content_snippet": body_text[:150],
            "matched_keywords": [],
            "error_reason": f"Nội dung bài viết quá ngắn hoặc trang web rỗng (chỉ có {len(body_text)} ký tự)"
        }

    # 6. So khớp Ngữ nghĩa & Thực thể Cốt lõi
    full_search_corpus = f"{page_title} {meta_desc} {body_text}"
    corpus_no_accents = remove_accents(full_search_corpus)
    
    # Nếu người dùng/Agent truyền vào danh sách expected_keywords rõ ràng
    # -> Bắt buộc phải tìm thấy ít nhất 1 từ trong expected_keywords
    if expected_keywords:
        if isinstance(expected_keywords, str):
            check_kws = [k.strip() for k in expected_keywords.split(",") if k.strip()]
        else:
            check_kws = list(expected_keywords)
            
        matched_explicit = []
        for kw in check_kws:
            kw_no_acc = remove_accents(kw.strip())
            if not kw_no_acc:
                continue
            if " " in kw_no_acc:
                if kw_no_acc in corpus_no_accents:
                    matched_explicit.append(kw)
            else:
                if re.search(r'\b' + re.escape(kw_no_acc) + r'\b', corpus_no_accents):
                    matched_explicit.append(kw)
                    
        if not matched_explicit:
            return {
                "is_valid": False,
                "url": url,
                "final_url": final_url,
                "status_code": status_code,
                "page_title": page_title,
                "meta_description": meta_desc,
                "content_snippet": body_text[:150],
                "matched_keywords": [],
                "error_reason": f"Nội dung trong link không chứa bất kỳ từ khóa yêu cầu nào: {expected_keywords}. Tiêu đề thực tế đọc được: '{page_title}'."
            }
        matched = matched_explicit
    else:
        # Tự động bóc tách thực thể từ tiêu đề
        auto_entities = extract_meaningful_entities(topic_title)
        matched_auto = []
        for ent in auto_entities:
            ent_no_acc = remove_accents(ent.strip())
            if not ent_no_acc:
                continue
            if " " in ent_no_acc:
                if ent_no_acc in corpus_no_accents:
                    matched_auto.append(ent)
            else:
                if re.search(r'\b' + re.escape(ent_no_acc) + r'\b', corpus_no_accents):
                    matched_auto.append(ent)
                    
        if auto_entities and not matched_auto:
            return {
                "is_valid": False,
                "url": url,
                "final_url": final_url,
                "status_code": status_code,
                "page_title": page_title,
                "meta_description": meta_desc,
                "content_snippet": body_text[:150],
                "matched_keywords": [],
                "error_reason": f"Nội dung trong link hoàn toàn KHÔNG khớp với chủ đề '{topic_title}'. Tiêu đề thực tế đọc được: '{page_title}'."
            }
        matched = matched_auto

    content_clean_snippet = re.sub(r'\s+', ' ', body_text)[:200]

    return {
        "is_valid": True,
        "url": url,
        "final_url": final_url,
        "status_code": status_code,
        "page_title": page_title,
        "meta_description": meta_desc,
        "content_snippet": content_clean_snippet,
        "matched_keywords": list(set(matched)),
        "error_reason": None
    }


def enforce_strict_validation(ideas_list):
    """
    Rào chắn kỹ thuật (Technical Gatekeeper) cưỡng chế kiểm tra trước khi ghi vào Excel.
    """
    print("\n" + "="*80)
    print("🛡️ BẮT ĐẦU QUY TRÌNH THẨM ĐỊNH KỸ THUẬT SÂU TỪNG LINK (URL VALIDATOR 4.0)")
    print("="*80)
    
    passed_count = 0
    failed_items = []
    
    for item in ideas_list:
        stt = item.get("stt", 0)
        title = item.get("title", "")
        url = item.get("url", "")
        keywords = item.get("keywords", None)
        
        result = validate_url_deep(url, topic_title=title, expected_keywords=keywords)
        
        if result["is_valid"]:
            passed_count += 1
            print(f"🟢 [STT {stt}] PASSED (HTTP {result['status_code']})")
            print(f"   👉 URL: {url}")
            print(f"   👉 Tiêu đề trang web: \"{result['page_title']}\"")
            print(f"   👉 Đoạn trích nội dung: \"{result['content_snippet'][:120]}...\"")
            print(f"   👉 Từ khóa/Thực thể tìm thấy: {result['matched_keywords'][:5]}")
            print("-" * 80)
        else:
            failed_items.append((stt, title, url, result["error_reason"], result.get("page_title")))
            print(f"🔴 [STT {stt}] REJECTED!")
            print(f"   👉 URL: {url}")
            print(f"   👉 Lý do từ chối: {result['error_reason']}")
            if result.get("page_title"):
                print(f"   👉 Tiêu đề thực tế của link sai: \"{result['page_title']}\"")
            print("-" * 80)

    print(f"\n📊 KẾT QUẢ KIỂM THỬ: ĐẠT {passed_count}/{len(ideas_list)} link.")
    
    if failed_items:
        print(f"❌ CẢNH BÁO: Phát hiện {len(failed_items)} link KHÔNG ĐẠT tiêu chuẩn kỹ thuật!")
        print("❌ HỆ THỐNG ĐÃ TỰ ĐỘNG CHẶN KHÔNG CHO GHI VÀO FILE EXCEL.")
        raise ValueError(f"Technical Gatekeeper chặn ghi Excel do có {len(failed_items)} URL không hợp lệ!")
    
    print("✅ TOÀN BỘ LINK ĐỀU THẬT 100%, ĐỌC ĐƯỢC NỘI DUNG VÀ KHỚP VỚI CHỦ ĐỀ!")
    return True


if __name__ == "__main__":
    print("--- Test 1: Link đúng (Wikipedia Panama Canal) ---")
    res1 = validate_url_deep(
        url="https://en.wikipedia.org/wiki/Panama_Canal",
        topic_title="Kênh đào Panama và bài toán nước ngọt",
        expected_keywords=["panama", "canal"]
    )
    print("Kết quả 1:", res1["is_valid"], "| Tiêu đề:", res1["page_title"], "| Khớp:", res1["matched_keywords"])

    print("\n--- Test 2: Link sai (VnExpress ID Olympic Tin học gán cho Panama Canal) ---")
    res2 = validate_url_deep(
        url="https://vnexpress.net/kenh-dao-panama-doi-mat-kho-han-chua-tung-co-4648753.html",
        topic_title="Kênh đào Panama và bài toán nước ngọt",
        expected_keywords=["panama", "canal"]
    )
    print("Kết quả 2:", res2["is_valid"], "| Lý do lỗi:", res2["error_reason"])
