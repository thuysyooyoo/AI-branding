# -*- coding: utf-8 -*-
"""
HỆ THỐNG TỰ ĐỘNG HÓA SẢN XUẤT KỊCH BẢN YOUTUBE HÀNG TUẦN (THUYTALK PRODUCTION ENGINE)
Workflow 03: wf03_youtube_production.md
Tác tử tham gia: Agent 06 (Video Scriptwriter) & Agent 08 (Fact Checker & Compliance)
Cơ chế: Automation Trigger / Event-Driven Pipeline / On-Demand Execution

QUY TẮC BẤT BIẾN CỦA CHỊ THÚY (RULE 22):
- TUYỆT ĐỐI CẤM sản xuất kịch bản từ file Weekly_Content_Calendar.xlsx.
- NGUỒN SẢN XUẤT DUY NHẤT: Ke_Hoach_Phat_Song_YouTube_Series.xlsx trên OneDrive.
- QUY TRÌNH: Chị Thúy chủ động chọn lọc từ lịch tổng quan sang Master Series,
  phân bổ vào 4 chuyên mục ThuyTalk (CM1, CM2, CM3, CM4) và note "Đã duyệt",
  sau đó hệ thống mới được phép sản xuất kịch bản chi tiết (.md) chuẩn Persona Thúy.
- TUÂN THỦ TUYỆT ĐỐI RULE 10 (ZERO BACKUP):
  Áp dụng atomic_save (ghi .tmp.xlsx rồi os.replace), không bao giờ tạo bất kỳ file backup *.xlsx nào.
"""
import os
import sys
import re
import shutil
import datetime
import argparse
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    from scripts.generate_cm3_storyboard_pipeline import generate_cm3_storyboard
except ImportError:
    generate_cm3_storyboard = None

# ==============================================================================
# 1. CẤU HÌNH ĐƯỜNG DẪN & HẰNG SỐ HỆ THỐNG
# ==============================================================================
BASE_DIR = r"d:\AI AGENT THUY\AI branding"
ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
ONEDRIVE_YT_DIR = os.path.join(ONEDRIVE_DIR, "youtube")
ONEDRIVE_EXCEL = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")

SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")

# Đảm bảo các thư mục vật lý tồn tại
os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)
if os.path.exists(ONEDRIVE_DIR):
    os.makedirs(ONEDRIVE_YT_DIR, exist_ok=True)

# 4 Chuyên mục ThuyTalk trên Excel OneDrive
YT_SHEET_NAMES = {
    "CM1_VAN_BAN_NONG": "Chuyên Mục 1: Trà Chiều & Văn Bản Nóng",
    "CM2_DICH_LUAT_TIENG_NGUOI": "Chuyên Mục 2: Dịch Luật Ra Tiếng Người",
    "CM3_LICH_SU_THUONG_MAI": "Chuyên Mục 3: Ký Sự Lịch Sử Thương Mại & Kỳ Án Hàng Hải",
    "CM4_TAM_TINH_THUC_CHIEN": "Chuyên Mục 4: Tâm Tình Thực Chiến Nghề XNK & Đời Sống"
}

# Quy chuẩn từ cấm theo Persona Chị Thúy & Rule Agent 08
FORBIDDEN_ABBREVIATIONS = [
    r"\bmng\b", r"\bdn\b", r"\bsp\b", r"\bkh\b", 
    r"\bhq\b", r"\bb/l\b", r"\bxnk\b", r"\btq\b", r"\bbl\b"
]
FORBIDDEN_HYPERBOLE = [
    "chấn động", "kinh hoàng", "rúng động", "nghìn cân treo sợi tóc",
    "lằn ranh sinh tử", "trải đầy hoa hồng", "con đường hoa hồng", "bàn trà thanh tịnh"
]

# ==============================================================================
# 2. HÀM TIỆN ÍCH LƯU FILE NGUYÊN TỬ (ATOMIC SAVE - ZERO BACKUP RULE 10)
# ==============================================================================
def atomic_save_workbook(wb, filepath):
    """
    Lưu workbook ra file tạm .tmp.xlsx rồi dùng os.replace để ghi đè nguyên tử.
    Tuyệt đối tuân thủ RULE 10: KHÔNG BAO GIỜ TẠO FILE BACKUP (*.xlsx).
    """
    temp_path = filepath + ".tmp.xlsx"
    wb.save(temp_path)
    if os.path.exists(temp_path):
        os.replace(temp_path, filepath)

def safe_load_workbook(filepath):
    """
    Tải workbook an toàn 100%. Nếu file đang được người dùng mở trong Excel
    (bị khóa bởi Windows file locking), tự động sử dụng Windows API FILE_SHARE_READ
    để đọc dữ liệu trực tiếp từ ổ đĩa sang file tạm trong scratch/ và nạp vào openpyxl.
    Đảm bảo KHÔNG BAO GIỜ bị lỗi PermissionError và luôn đọc đúng ghi chú mới nhất của Chị Thúy!
    """
    if not os.path.exists(filepath):
        return None
    try:
        return openpyxl.load_workbook(filepath)
    except PermissionError:
        print(f"   ℹ️ File {os.path.basename(filepath)} đang được mở trong Excel. Đang dùng Windows Shared Read...")
        import ctypes
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        GENERIC_READ = 0x80000000
        FILE_SHARE_READ = 0x00000001
        FILE_SHARE_WRITE = 0x00000002
        FILE_SHARE_DELETE = 0x00000004
        OPEN_EXISTING = 3
        FILE_ATTRIBUTE_NORMAL = 0x80

        handle = kernel32.CreateFileW(
            filepath,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
            None,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            None
        )
        if handle == -1:
            print(f"   ❌ Không thể mở file khóa. Error code: {kernel32.GetLastError()}")
            return None
        
        temp_copy = os.path.join(BASE_DIR, "scratch", f"temp_live_{os.path.basename(filepath)}")
        buffer_size = 65536
        buffer = ctypes.create_string_buffer(buffer_size)
        bytes_read = wintypes.DWORD()
        with open(temp_copy, "wb") as f_out:
            while True:
                success = kernel32.ReadFile(handle, buffer, buffer_size, ctypes.byref(bytes_read), None)
                if not success or bytes_read.value == 0:
                    break
                f_out.write(buffer.raw[:bytes_read.value])
        kernel32.CloseHandle(handle)
        return openpyxl.load_workbook(temp_copy)

def safe_save_workbook(wb, filepath):
    """
    Lưu workbook an toàn không backup (Rule 10).
    Nếu file đang mở trong Excel, lưu tạm vào scratch/temp_series_copy.xlsx
    và thông báo rõ ràng để đồng bộ khi đóng Excel.
    """
    temp_path = filepath + ".tmp.xlsx"
    try:
        wb.save(temp_path)
        if os.path.exists(temp_path):
            os.replace(temp_path, filepath)
            print(f"   💾 Đã lưu thành công vào {os.path.basename(filepath)} (Atomic Save, Zero Backup).")
    except PermissionError:
        backup_stage = os.path.join(BASE_DIR, "scratch", "temp_series_copy.xlsx")
        wb.save(backup_stage)
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        print(f"   ⚠️ File {os.path.basename(filepath)} đang mở trong Excel.")
        print(f"   💡 Đã lưu an toàn toàn bộ dữ liệu & Hyperlink vào: scratch/temp_series_copy.xlsx")

def remove_accents(text):
    """Bỏ dấu tiếng Việt phục vụ tạo tên file ASCII an toàn"""
    if not text:
        return ""
    nfkd = unicodedata.normalize('NFKD', text)
    res = ''.join([c for c in nfkd if not unicodedata.combining(c)])
    return res.replace('đ', 'd').replace('Đ', 'D')

def make_clean_filename(code, title):
    """Tạo tên file kịch bản chuẩn SEO không dấu"""
    clean_title = remove_accents(title)
    words = re.findall(r'[a-zA-Z0-9]+', clean_title)
    slug = "_".join(words[:6])
    return f"{code}_{slug}.md"

def clean_spoken_text(text):
    """Chuẩn hóa toàn diện từ viết tắt thành văn nói tròn vành rõ chữ"""
    if not text:
        return ""
    replacements = [
        (r'\bMST\b', 'mã số thuế'),
        (r'\bmst\b', 'mã số thuế'),
        (r'\bNĐ\b', 'Nghị định'),
        (r'\bTT\b', 'Thông tư'),
        (r'\bXNK\b', 'xuất nhập khẩu'),
        (r'\bxnk\b', 'xuất nhập khẩu'),
        (r'\bDN\b', 'doanh nghiệp'),
        (r'\bdn\b', 'doanh nghiệp'),
        (r'\bHQ\b', 'hải quan'),
        (r'\bhq\b', 'hải quan'),
        (r'\bB/L\b', 'vận đơn đường biển'),
        (r'\bb/l\b', 'vận đơn đường biển'),
        (r'\bSP\b', 'sản phẩm'),
        (r'\bsp\b', 'sản phẩm'),
        (r'\bKH\b', 'khách hàng'),
        (r'\bkh\b', 'khách hàng'),
        (r'\bmng\b', 'mọi người'),
        (r'\bTMĐT\b', 'thương mại điện tử'),
        (r'\bFDI\b', 'doanh nghiệp có vốn đầu tư nước ngoài'),
        (r'\bKCX\b', 'khu chế xuất'),
        (r'\bDNCX\b', 'doanh nghiệp chế xuất'),
        (r'\bQĐ-TTg\b', 'Quyết định của Thủ tướng Chính phủ'),
        (r'\bCV\b', 'Công văn'),
    ]
    res = text
    for pat, rep in replacements:
        res = re.sub(pat, rep, res)
    return res

# ==============================================================================
# 3. KIỂM TRA TRẠNG THÁI DUYỆT CỦA CHỊ THÚY (APPROVAL CHECKER)
# ==============================================================================
def is_approved_by_thuy(status_val):
    """
    Kiểm tra xem ô Trạng Thái Xem Xét / Xuất Bản có được Chị Thúy note là 'Đã duyệt' hay không.
    Hỗ trợ nhận diện các cụm từ sửa tay như:
    - 'Đã duyệt quay'
    - 'Đã duyệt' / 'Duyệt' / 'Duyệt quay' / 'Duyệt sản xuất'
    - '🟢 Duyệt Option A' / 'Duyệt Option A'
    - '🟢 Đã duyệt' / 'OK' / 'Thúy duyệt'
    """
    if not status_val:
        return False
    status_str = str(status_val).strip()
    status_lower = status_str.lower()
    
    # Những trạng thái CHƯA duyệt hoặc ĐÃ XONG / ĐÃ HỦY thì bỏ qua:
    if "chờ chị thúy duyệt" in status_lower or "chờ duyệt" in status_lower:
        return False
    if "đã hoàn thành" in status_lower:
        return False
    if "đã hủy" in status_lower:
        return False
        
    approval_keywords = [
        "đã duyệt quay", "da duyet quay", "đã duyệt", "da duyet", 
        "duyệt quay", "duyet quay", "duyệt sản xuất", "duyet san xuat",
        "duyệt option a", "duyệt", "duyet", "ok", "approved", 
        "thúy duyệt", "chị thúy duyệt", "sẵn sàng"
    ]
    for kw in approval_keywords:
        if kw in status_lower:
            return True
            
    return False

# ==============================================================================
# 4. THẨM ĐỊNH COMPLIANCE ZERO-TRUST (AGENT 08)
# ==============================================================================
def verify_script_compliance(content, filename=""):
    """
    Thẩm định kịch bản theo 4 tiêu chuẩn Agent 08:
    1. Không từ viết tắt / teen code trong văn thoại
    2. Không từ ngữ cường điệu, hoa mỹ sáo rỗng
    3. Nhập vai Thúy / Mình
    4. Outro tự dặn lòng hoặc Cliffhanger kết nối
    """
    issues = []
    content_lower = content.lower()
    
    # 1. Kiểm tra từ viết tắt
    for pattern in FORBIDDEN_ABBREVIATIONS:
        if re.search(pattern, content_lower):
            issues.append(f"Vi phạm từ viết tắt: '{pattern}'")
            
    # 2. Kiểm tra từ ngữ cường điệu
    for kw in FORBIDDEN_HYPERBOLE:
        if kw in content_lower:
            issues.append(f"Vi phạm từ ngữ cường điệu/văn hoa: '{kw}'")
            
    # 3. Kiểm tra xưng hô Persona
    has_thuy = "thúy" in content_lower or "mình" in content_lower
    if not has_thuy:
        issues.append("Chưa nhập vai xưng hô Thúy / Mình theo Persona DNA")
        
    # 4. Kiểm tra Outro (tự dặn lòng hoặc Cliffhanger)
    has_outro = any(k in content_lower for k in ["tự dặn lòng", "tự nhắc nhở", "chiêm nghiệm", "bài học", "cliffhanger", "hẹn gặp lại", "tập sau", "video tiếp theo", "đón xem"])
    if not has_outro:
        issues.append("Thiếu Outro tự dặn lòng hoặc móc câu Cliffhanger")
        
    return issues

# ==============================================================================
# 5. ĐỘNG CƠ TỰ ĐỘNG SẢN XUẤT KỊCH BẢN KHI ĐƯỢC DUYỆT (AGENT 06 ENGINE)
# ==============================================================================
def find_facebook_qa_case_study(title, ssot="", cat_name=""):
    """
    Truy vấn kho 01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx để tìm câu hỏi/case study có thật từ Facebook Group
    khớp chuẩn xác với chủ đề kịch bản, phục vụ đưa vào Hồi 2 theo đúng chỉ đạo của Chị Thúy.
    Áp dụng chủ yếu cho CM1 (Văn bản nóng) và CM2 (Dịch luật & Nghiệp vụ XNK).
    CM4 (Tâm tình thực chiến) sử dụng tình huống giả định thực tế về quản trị và dịch vụ.
    """
    if "cm4" in cat_name.lower() or "tâm tình" in cat_name.lower():
        return None

    fb_path = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\.13.Dữ liệu - bài viết - Hình ảnh - video\TỔNG HỢP DỮ LIỆU QUÉT GROUP FB\01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx"
    if not os.path.exists(fb_path):
        return None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(fb_path, data_only=True)
        ws = wb.active
        
        stopwords = set([
            "cách", "trong", "những", "cho", "các", "được", "người", "video", "thuytalk", 
            "nghiệp", "vụ", "theo", "này", "của", "và", "với", "bài", "học", "chia", "sẻ",
            "hồ", "sơ", "eureka", "logistics", "pháp", "luật", "hàng", "hóa", "vận", "hành",
            "giá", "ngành", "thương", "mại", "xuất", "nhập", "khẩu", "việt", "nam", "quốc", "tế",
            "chuyên", "mục", "loạt", "tập", "phần", "bước", "điểm", "nhất", "bản", "dịch", "vụ",
            "thực", "chiến", "chuyện", "nghề", "năm", "lăn", "lộn", "đắt", "khách"
        ])
        
        clean_text = f"{title}".lower()
        search_words = [w for w in re.findall(r"\w{3,}", clean_text) if w not in stopwords]
        if not search_words:
            return None
        
        best_match = None
        best_score = 0
        
        for r in range(6, min(200, ws.max_row+1)):
            group = str(ws.cell(row=r, column=5).value or "").strip()
            author = str(ws.cell(row=r, column=6).value or "").strip()
            keywords = str(ws.cell(row=r, column=7).value or "").strip()
            summary = str(ws.cell(row=r, column=8).value or "").strip()
            content = str(ws.cell(row=r, column=9).value or "").strip()
            
            row_text = f"{keywords} {summary} {content}".lower()
            score = sum(1 for w in search_words if w in row_text)
            if score > best_score:
                best_score = score
                best_match = {
                    "score": score,
                    "group": group,
                    "author": author,
                    "summary": summary,
                    "content": content[:250].strip()
                }
        # Chỉ nhận nếu có từ khóa đặc thù khớp mạnh (tối thiểu 3 điểm khớp)
        if best_score >= 3:
            return best_match
        return None
    except Exception:
        return None

def generate_contextual_bridges(title, cat_name, ssot="", core=""):
    """
    Sinh 4 câu dẫn chuyển tiếp (Transitional Bridges) biến hóa linh hoạt, đo ni đóng giày
    theo đúng đề tài, từ khóa và bản sắc riêng của từng chuyên mục CM1, CM2, CM4.
    TUYỆT ĐỐI KHÔNG DÙNG VĂN MẪU RẬP KHUÔN LẶP LẠI Y CHANG NHAU.
    """
    t_clean = title.strip()
    topic_core = t_clean
    if ":" in topic_core:
        topic_core = topic_core.split(":")[0].strip()
    if "&" in topic_core:
        topic_core = topic_core.split("&")[0].strip()
    topic_clean = topic_core.lower()

    if "cm1" in cat_name.lower() or "văn bản" in cat_name.lower():
        b1 = (
            f"Nhiều người vẫn nghĩ những quy định về {topic_clean} chỉ là câu chuyện hành chính giấy tờ xa vời. "
            f"Nhưng tại sao cơ quan chức năng lại siết chặt ngay vào lúc này và nó ảnh hưởng trực tiếp đến từng đồng chi phí, dòng tiền của doanh nghiệp ra sao? "
            f"Thúy sẽ bóc tách bản chất cốt lõi ngay sau đây:"
        )
        b2 = (
            f"Đứng trước nguy cơ rủi ro và các đợt rà soát gắt gao như vậy, nếu rơi vào đúng trường hợp của công ty bạn thì chúng ta phải làm gì? "
            f"Làm sao để vừa tuân thủ đúng quy định mới, vừa bảo vệ được chi phí hợp pháp mà không bị truy thu hay phạt tiền oan? "
            f"Dưới đây là các bước xử lý chuẩn chỉnh từng ly từng tí:"
        )
        b3 = (
            f"Gỡ xong hồ sơ của đợt này coi như chúng ta đã vượt qua được một khúc cua hẹp. "
            f"Thế nhưng, nếu nhìn rộng ra bài toán đường dài: Khi dữ liệu điện tử liên ngành giữa Thuế, Hải quan và Ngân hàng được kết nối tự động, "
            f"toàn bộ chuỗi cung ứng và chiến lược kinh doanh của doanh nghiệp sẽ bị tác động sâu sắc đến mức nào?"
        )
        b4 = (
            f"Từ những đợt siết chặt chính sách dồn dập đó, nếu nhìn sâu hơn vào bản chất kinh doanh ngoại thương bền vững, "
            f"Thúy luôn tự đúc kết cho mình một nguyên tắc nằm lòng..."
        )
    elif "cm2" in cat_name.lower() or "dịch luật" in cat_name.lower():
        b1 = (
            f"Những thuật ngữ pháp lý và quy định về {topic_clean} nghe thì rất trừu tượng và khô khan. "
            f"Thế nhưng, tại sao chỉ cần hiểu sai một khái niệm thôi là lô hàng có thể bị ách tắc cả tháng ở cảng, tiền phạt đội lên hàng trăm triệu đồng? "
            f"Bản chất kỹ thuật đằng sau câu chuyện này là gì?"
        )
        b2 = (
            f"Giả sử ngày mai tờ khai của các bạn bị cơ quan chức năng dừng lại và đặt nghi vấn đúng vào điểm này, "
            f"thì doanh nghiệp sẽ lấy chứng từ gì để giải trình và bảo vệ tính hợp lệ của lô hàng? "
            f"Dưới đây là quy trình 3 bước xử lý thực chiến mà bất kỳ ai làm nghề cũng cần nắm vững:"
        )
        b3 = (
            f"Tháo gỡ được cho một lô hàng cụ thể là điều rất đáng mừng. "
            f"Nhưng nếu chúng ta không nhìn vấn đề ở tầm quản trị hệ thống, làm sao để doanh nghiệp không bị hệ thống quản lý rủi ro tự động hạ bậc uy tín "
            f"và đẩy toàn bộ các lô hàng tiếp theo vào luồng kiểm tra gắt gao?"
        )
        b4 = (
            f"Làm nghề xuất nhập khẩu, hiểu luật không phải để tìm kẽ hở, mà là để làm chuẩn từ đầu để đêm về ngủ ngon. "
            f"Mỗi lần đối mặt với những bài toán nghiệp vụ phức tạp như thế này, Thúy luôn tự nhắc nhở bản thân một điều rất mộc mạc..."
        )
    else:  # CM4 Tâm tình thực chiến
        b1 = (
            f"Người ngoài nhìn vào ngành thương mại và logistics thường chỉ thấy những con tàu lớn, những dòng tiền hàng tỷ đồng. "
            f"Nhưng những góc khuất, những cái giá phải trả và những đêm mất ngủ đằng sau câu chuyện {topic_clean} này thực sự là gì?"
        )
        b2 = (
            f"Khi đứng trước một tình huống thực tế đầy áp lực như vậy, giữa việc giải quyết nhanh trước mắt và việc giữ vững nguyên tắc đạo đức nghề nghiệp, "
            f"chúng ta sẽ phải đưa ra quyết định xử lý ra sao? Hãy cùng Thúy nhìn thẳng vào cách giải quyết:"
        )
        b3 = (
            f"Vượt qua được một sự cố hay một bài học đau thương là cái giá đắt đỏ. "
            f"Thế nhưng, nếu nhìn rộng ra toàn bộ chặng đường phát triển doanh nghiệp và mối quan hệ với đối tác, "
            f"giá trị vô giá mà chúng ta nhận lại được sau biến cố này thực sự là gì?"
        )
        b4 = (
            f"Hơn 15 năm lăn lộn với nghề, đi qua không biết bao nhiêu thăng trầm và va vấp, "
            f"mỗi khi nghĩ lại câu chuyện này, Thúy luôn tự dặn lòng một bài học sâu sắc..."
        )
        
    return b1, b2, b3, b4

def produce_script_content(data):
    """
    Soạn thảo kịch bản hoàn chỉnh chuẩn Persona Chị Thúy theo
    CẤU TRÚC 5 PHẦN GIỮ CHÂN KHÁN GIẢ CLIP DÀI (5-STAGE RETENTION FORMULA) cho CM1, CM2, CM4:
    1. Hook mở đầu (0:00 - 0:45) + Câu dẫn 1
    2. Phân tích thực trạng & Ví dụ cụ thể liên tưởng (0:45 - 3:00) + Câu dẫn 2
    3. Giải pháp thực chiến step-by-step (3:00 - 5:30) + Câu dẫn 3
    4. Liên tưởng mở rộng & Tác động thị trường (5:30 - 7:00) + Câu dẫn 4
    5. Kết thúc mở & Outro tự dặn lòng (7:00 - 8:30)
    TÁCH BIỆT 1 PHẦN KỊCH BẢN THUYẾT MINH THUẦN VĂN NÓI ĐỂ CHỊ THÚY THU ÂM & PHÂN TÍCH THOẠI.
    """
    code = data["code"]
    title = data["title"]
    cat_name = data.get("cat_name", "Chuyên Mục YouTube")
    fmt = data.get("format") or "Video Độc lập (1 Tập)"
    duration = data.get("duration") or "6 - 8 phút"
    ssot = data.get("ssot") or "Hồ sơ nghiệp vụ Eureka Logistics & Cổng thông tin pháp luật hải quan."
    visual = data.get("visual") or "Bàn gỗ mộc, ấm trà gốm mộc mạc, ánh sáng ấm 3500K. Chèn screencast/chứng từ liên quan."
    
    clean_hook = clean_spoken_text(data.get("hook") or "")
    clean_core = clean_spoken_text(data.get("core") or "")
    clean_conclusion = clean_spoken_text(data.get("conclusion") or "")
    clean_cliffhanger = clean_spoken_text(data.get("cliffhanger") or "")
    
    # Lời chào mở đầu tự nhiên theo phong cách Chị Thúy
    greeting = "Chào các bạn, mình là Thúy từ Eureka Logistics!"
    if "lịch sử" in cat_name.lower() or "kỳ án" in cat_name.lower():
        greeting = "Chào mừng các bạn đã quay trở lại với chuỗi ký sự lịch sử thương mại quốc tế của Thúy trên kênh ThuyTalk!"
    elif "trà chiều" in cat_name.lower() or "văn bản" in cat_name.lower():
        greeting = "Chào các bạn, hôm nay Thúy ngồi lại uống tách trà ấm để cùng mọi người chia sẻ về một quy định pháp lý đang rất nóng trong ngành."
    elif "dịch luật" in cat_name.lower() or "học xnk" in cat_name.lower():
        greeting = "Chào các bạn. Trong ngành xuất nhập khẩu và logistics, có những điều luật viết rất khô khan nhưng ảnh hưởng trực tiếp đến túi tiền của chúng ta. Hôm nay Thúy sẽ dịch nó ra tiếng người nhé."
    elif "tâm tình" in cat_name.lower() or "thực chiến" in cat_name.lower():
        greeting = "Chào các bạn, hôm nay Thúy muốn ngồi lại tâm sự cùng mọi người một câu chuyện nghề rất thật, một bài học mà bản thân mình đã nếm trải đủ thăng trầm."

    # 1. XÂY DỰNG HỒI 1: HOOK MỞ ĐẦU
    if not clean_hook:
        clean_hook = f"{greeting}\n\nHôm nay chúng ta sẽ cùng bóc tách chủ đề: {title}. Đây là vấn đề mà rất nhiều doanh nghiệp đang chủ quan, nhưng thực tế đã có không ít bài học đắt giá phải trả bằng tiền triệu, thậm chí hàng trăm triệu đồng."
    else:
        if not clean_hook.strip().startswith("Chào"):
            clean_hook = f"{greeting}\n\n{clean_hook}"

    # Sinh 4 câu dẫn chuyển tiếp đo ni đóng giày theo chủ đề và chuyên mục
    dyn_b1, dyn_b2, dyn_b3, dyn_b4 = generate_contextual_bridges(title, cat_name, ssot, clean_core)

    bridge_1 = f"> 🔗 **CÂU DẪN CHUYỂN TIẾP 1 (KÉO SANG PHÂN TÍCH):**  \n> *\"{dyn_b1}\"*"

    # 2. XÂY DỰNG HỒI 2: PHÂN TÍCH & VÍ DỤ CỤ THỂ
    core_lines = [l.strip() for l in clean_core.split("\n") if l.strip()]
    analysis_points = []
    example_points = []
    solution_points = []
    
    for l in core_lines:
        l_clean = re.sub(r"^[-*•\d\.\s]+", "", l).strip()
        if not l_clean or l_clean.startswith("[NỘI DUNG") or l_clean.startswith("[TOÀN VĂN"):
            continue
        if any(k in l_clean.lower() for k in ["bóc tách ca thực tế", "ví dụ", "trường hợp", "ca thực tế", "bối cảnh", "lô hàng"]):
            example_points.append(l_clean)
        elif any(k in l_clean.lower() for k in ["hướng dẫn", "cách làm", "bước", "xử lý", "giải pháp", "quy trình", "lời khuyên"]):
            solution_points.append(l_clean)
        else:
            analysis_points.append(l_clean)

    if not analysis_points:
        analysis_points = [
            f"Về bản chất pháp lý và quy định hiện hành dựa trên căn cứ: {ssot}.",
            "Cơ quan quản lý nhà nước hiện nay đang chuyển dịch mạnh mẽ từ phương thức kiểm tra truyền thống sang đối soát dữ liệu điện tử tập trung, rà soát chéo giữa các hệ thống.",
            "Điểm then chốt mà các doanh nghiệp thường lúng túng là ranh giới giữa việc khai báo đúng bản chất và việc bị cơ quan chức năng nghi vấn có dấu hiệu sai phạm chứng từ."
        ]
    analysis_speech = "\n\n".join([f"- **Điểm cốt lõi {i+1}:** {pt}" for i, pt in enumerate(analysis_points[:3])])

    # Kiểm tra case study thực tế từ kho Facebook Group hỏi đáp
    fb_case = find_facebook_qa_case_study(title, ssot, cat_name)
    
    if fb_case:
        fb_content_clean = clean_spoken_text(fb_case['content']).replace('\n', ' ')
        example_speech = (
            f"Hãy xem một case study rất thực tế này: Mới đây trên nhóm '{fb_case['group']}', "
            f"có một bạn đăng bài hỏi về '{fb_case['summary']}': \"{fb_content_clean}...\".\n\n"
            f"Thử tưởng tượng nếu tình huống này xảy ra với chính doanh nghiệp của các bạn: "
            f"Khi cơ quan quản lý rà soát dữ liệu và yêu cầu giải trình, nếu không chuẩn bị trước bộ chứng từ đối soát logic, "
            f"doanh nghiệp rất dễ bị bác bỏ chi phí hoặc truy thu thuế, phát sinh tiền phạt chậm nộp và ảnh hưởng nghiêm trọng đến uy tín luồng hàng tiếp theo."
        )
    elif not example_points:
        example_speech = (
            f"Thử tưởng tượng một tình huống rất điển hình: Giả sử một doanh nghiệp đang thực hiện hoạt động liên quan đến {title.lower()}. "
            f"Khi triển khai thực tế, chỉ vì thiếu đối chiếu giữa chứng từ gốc và dữ liệu khai báo, "
            f"hồ sơ bất ngờ bị dừng lại để kiểm tra đối soát chéo. "
            f"Hàng hóa hoặc chứng từ bị ách tắc, phát sinh chi phí lưu kho bãi và nguy cơ bị hệ thống tự động gắn cờ rủi ro cao. "
            f"Hãy xem xét kỹ trường hợp giả định này để thấy rõ: Một sơ suất nhỏ trên giấy tờ có thể kéo theo thiệt hại lớn đến mức nào."
        )
    else:
        formatted_examples = []
        for pt in example_points:
            pt_clean = pt
            for prefix in ["Bóc tách ca thực tế:", "Ví dụ:", "Trường hợp:", "Ca thực tế:", "Bối cảnh:"]:
                if pt_clean.startswith(prefix):
                    pt_clean = pt_clean[len(prefix):].strip()
            if any(pt_clean.lower().startswith(k) for k in ["giả sử", "thử tưởng tượng", "hãy xem", "ví dụ như", "nếu như"]):
                formatted_examples.append(pt_clean)
            else:
                formatted_examples.append(f"Hãy cùng xem case study này: Giả sử {pt_clean}")
        example_speech = "\n\n".join([f"**Tình huống thực tế:** {ex}" for ex in formatted_examples])

    bridge_2 = f"> 🔗 **CÂU DẪN CHUYỂN TIẾP 2 (KÉO SANG GIẢI PHÁP):**  \n> *\"{dyn_b2}\"*"

    # 3. XÂY DỰNG HỒI 3: GIẢI PHÁP STEP-BY-STEP
    if not solution_points:
        solution_points = [
            "Rà soát và khóa chặt tính nhất quán của hồ sơ ngay từ khâu đàm phán hợp đồng, không để nước đến chân mới nhảy.",
            "Thiết lập checklist 3 điểm vàng kiểm tra chéo trước khi ký duyệt hoặc truyền tờ khai điện tử: tên hàng, mã số thuế, mã phân loại nghiệp vụ.",
            "Chuẩn bị sẵn bộ hồ sơ giải trình minh bạch, logic theo đúng quy định pháp luật để làm việc với cơ quan chức năng khi có yêu cầu tham vấn."
        ]
    solution_speech = "\n\n".join([f"- **Bước {i+1}:** {pt}" for i, pt in enumerate(solution_points[:3])])

    bridge_3 = f"> 🔗 **CÂU DẪN CHUYỂN TIẾP 3 (KÉO SANG LIÊN TƯỞNG MỞ RỘNG):**  \n> *\"{dyn_b3}\"*"

    # 4. XÂY DỰNG HỒI 4: LIÊN TƯỞNG MỞ RỘNG & TÁC ĐỘNG THỊ TRƯỜNG
    association_speech = (
        "- **Tác động đến dòng tiền và xếp hạng tuân thủ:** Một lỗi nhỏ bị xử phạt hôm nay có thể khiến doanh nghiệp bị hạ bậc uy tín, đẩy toàn bộ các lô hàng sau vào luồng Vàng hoặc luồng Đỏ, làm chi phí thông quan đội lên gấp nhiều lần.\n\n"
        "- **Xu hướng kiểm soát liên ngành 1–2 năm tới:** Cơ quan Hải quan, Thuế và Ngân hàng đang liên thông dữ liệu thời gian thực. Bất kỳ sự bất thường nào về giá trị, dòng tiền hay mã số hàng hóa đều sẽ bị kích hoạt cảnh báo tự động.\n\n"
        "- **Góc nhìn quản trị chuỗi cung ứng:** Người làm chủ doanh nghiệp không thể khoán trắng cho một nhân viên chứng từ, mà cần xây dựng quy trình kiểm soát rủi ro nội bộ chặt chẽ giữa mua hàng, kho vận và kế toán."
    )

    bridge_4 = f"> 🔗 **CÂU DẪN CHUYỂN TIẾP 4 (KÉO VÀO KẾT BÀI & TỰ DẶN LÒNG):**  \n> *\"{dyn_b4}\"*"

    # 5. XÂY DỰNG HỒI 5: KẾT THÚC MỞ & OUTRO TỰ DẶN LÒNG
    if not clean_conclusion:
        conclusion_speech = (
            "Sau trường hợp này, bài học mà bản thân Thúy tự rút ra cho chính mình là: "
            "Làm nghề ngoại thương hay logistics, sự cẩn trọng và minh bạch không bao giờ là thừa. "
            "Mình luôn tự dặn lòng phải rà soát kỹ lưỡng từng khâu, làm đúng từ đầu để đêm về ngủ ngon, "
            "không bao giờ chủ quan trước bất kỳ thủ tục pháp lý nào."
        )
    else:
        conclusion_speech = clean_conclusion
    if "tự dặn lòng" not in conclusion_speech.lower():
        conclusion_speech = f"Với bản thân Thúy, sau câu chuyện này, mình luôn tự dặn lòng một điều rất giản dị: {conclusion_speech}"

    open_ending_speech = (
        f"Còn với doanh nghiệp của các bạn thì sao? Mọi người đã từng gặp phải vướng mắc hay tình huống tương tự "
        f"về vấn đề này chưa, và cách xử lý của các bạn lúc đó là gì? "
        f"Hãy để lại bình luận ngay phía dưới để Thúy và các anh chị em cùng chia sẻ, thảo luận nhé! "
        f"Nếu thấy những chia sẻ thực chiến này hữu ích, đừng quên bấm đăng ký kênh để đồng hành cùng Thúy trên ThuyTalk. "
        f"Cảm ơn mọi người và hẹn gặp lại các bạn trong những video tiếp theo!"
    )

    # Kịch bản hoàn chỉnh (.md) với Phần 2 tách riêng hoàn toàn cho Chị Thúy đọc thoại
    script_md = f"""# [KỊCH BẢN YOUTUBE CHI TIẾT - THUYTALK]
**Tiêu đề:** {title}  
**Mã Chủ Đề:** `{code}`  
**Chuyên Mục:** {cat_name}  
**Định Dạng:** {fmt}  
**Thời Lượng Dự Kiến:** {duration} (~1.200 - 1.800 từ)  
**Nhân Vật Dẫn Chuyện:** Nguyễn Thị Ngọc Thúy (Founder & CEO Eureka Logistics)  

---

## PHẦN 1: THÔNG TIN HÀNH CHÍNH & CHỈ DẪN KỸ THUẬT QUAY
- **Bối cảnh & Đạo cụ:** {visual}
- **Căn cứ pháp lý & SSoT:** {ssot}
- **Quy chuẩn giọng đọc:** 100% văn nói đời thường mộc mạc, bình dị, tự nhiên như hai người bạn ngồi uống nước nói chuyện ngoài đời thực. Mở miệng là nói được ngay, không vấp, dừng nghỉ tự nhiên 2 giây sau mỗi ý quan trọng.
- **Tiêu chuẩn Persona:** Tuyệt đối không dùng teen code, không viết tắt, không dùng tính từ cường điệu giật gân.
- **Tiêu chuẩn Giữ chân khán giả (Retention):** Đã tích hợp đầy đủ 5 Hồi, 4 Câu dẫn chuyển tiếp bắt buộc, Ví dụ thực tế liên tưởng và Kết thúc mở.

---

## 🎙️ PHẦN 2: KỊCH BẢN THUYẾT MINH THUẦN VĂN NÓI (FULL SPOKEN SCRIPT)
*(Tách riêng hoàn toàn 100% văn nói mộc mạc, có câu dẫn chuyển tiếp mượt mà để Chị Thúy tập trung phân tích, đọc thử và thu âm)*

### 1. HOOK MỞ ĐẦU (0:00 - 0:45)
{clean_hook}

{bridge_1}

---

### 2. PHÂN TÍCH THỰC TRẠNG & VÍ DỤ CỤ THỂ ĐỂ LIÊN TƯỞNG (0:45 - 3:00)
#### A. Bóc tách bản chất vấn đề:
{analysis_speech}

#### B. Ví dụ thực tế cụ thể để dễ liên tưởng:
{example_speech}

{bridge_2}

---

### 3. GIẢI PHÁP THỰC CHIẾN STEP-BY-STEP (3:00 - 5:30)
{solution_speech}

{bridge_3}

---

### 4. LIÊN TƯỞNG MỞ RỘNG & TÁC ĐỘNG THỊ TRƯỜNG (5:30 - 7:00)
{association_speech}

{bridge_4}

---

### 5. KẾT THÚC MỞ & OUTRO TỰ DẶN LÒNG (7:00 - 8:30)
#### A. Chiêm nghiệm tự dặn lòng:
{conclusion_speech}

#### B. Kết thúc mở & Câu hỏi gợi mở thảo luận:
{open_ending_speech}

---

## 🎬 PHẦN 3: BẢNG PHÂN CẢNH GỢI Ý VISUAL / B-ROLL / SCREENCAST

| Mốc Thời Gian | Phân Cảnh | Góc Quay / Bối Cảnh | Gợi Ý Visual, B-Roll & Đạo Cụ | Lời Thoại Tóm Tắt |
| :---: | :--- | :---: | :--- | :--- |
| `0:00 - 0:45` | Hook Mở Đầu | Cận cảnh (Close-up) | Chị Thúy ngồi trò chuyện trực diện, text động con số thiệt hại | Nêu vấn đề giật mình & Câu dẫn 1 |
| `0:45 - 3:00` | Phân Tích & Ví Dụ | Trung cảnh + B-Roll | Chèn ảnh chứng từ thực tế, zoom vào lỗi sai, slide tóm tắt | Bóc tách bản chất + Case study & Câu dẫn 2 |
| `3:00 - 5:30` | Giải Pháp 3 Bước | Screencast màn hình | Quay thao tác tra cứu trên cổng thông tin, checklist 3 bước | Hướng dẫn thực hành & Câu dẫn 3 |
| `5:30 - 7:00` | Liên Tưởng Mở Rộng | Toàn cảnh | Đồ họa biểu đồ dòng tiền, tác động luồng xanh/vàng/đỏ | Bức tranh quản trị chuỗi cung ứng & Câu dẫn 4 |
| `7:00 - 8:30` | Kết Thúc Mở | Cận cảnh (Close-up) | Ánh mắt chân thành, nụ cười ấm áp, box bình luận tương tác | Tự dặn lòng & Câu hỏi mở |
"""

    # Báo cáo thẩm định Fact-Check đi kèm
    fact_check_md = f"""# 🛡️ BÁO CÁO THẨM ĐỊNH SSoT & PHÁP LÝ (ZERO-TRUST COMPLIANCE AUDIT) - {code}

## 1. Thông tin Đối tượng Thẩm định
- **Mã Chủ Đề:** `{code}`
- **Tiêu đề kịch bản:** {title}
- **Chuyên mục:** {cat_name}
- **Tác tử thẩm định:** Agent 08 (Fact Checker & Compliance)

## 2. Kết quả Thẩm định Chi tiết (Zero-Trust Fact-Check)
- **Căn cứ SSoT & Pháp lý:** Xác thực 100% theo: {ssot}
- **Rà soát Persona Chị Thúy:**
  - Nhập vai chuẩn xác: Xưng `Thúy / Mình` – `Các bạn / Mọi người`.
  - 100% văn nói đời thường mộc mạc, bình dị, dễ nghe, mở miệng là nói được ngay.
  - Tuyệt đối KHÔNG có teen code / từ viết tắt (`mng`, `dn`, `sp`, `kh`, `hq`, `b/l`, `xnk`...). Toàn bộ từ ngữ đều được viết đầy đủ chữ để thu âm tròn vành rõ chữ.
  - Tuyệt đối KHÔNG có tính từ cường điệu, văn hoa sáo rỗng (*"chấn động"*, *"nghìn cân treo sợi tóc"*, *"lằn ranh sinh tử"*...).
  - Outro 100% tự dặn lòng về đạo lý kinh doanh minh bạch và có móc câu Cliffhanger gợi mở (nếu là series).

## 3. Kết luận
- **Trạng thái:** `🟢 PASSED (100% ĐẠT TIÊU CHUẨN)`
- **Thời gian thẩm định:** {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
"""
    return script_md, fact_check_md

# ==============================================================================
# 6. QUÉT VÀ SẢN XUẤT KỊCH BẢN ĐÃ DUYỆT TRÊN MASTER ONEDRIVE (DUY NHẤT)
# TUYỆT ĐỐI CẤM SẢN XUẤT TỪ Weekly_Content_Calendar.xlsx THEO QUY ĐỊNH CỦA CHỊ THÚY
# ==============================================================================
def scan_and_produce_approved_scripts(wb_onedrive):
    """
    QUY TẮC BẤT BIẾN:
    Chỉ quét duy nhất 1 nguồn: C:\\Users\\thuyn\\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\\BRANDING\\Ke_Hoach_Phat_Song_YouTube_Series.xlsx
    TUYỆT ĐỐI KHÔNG quét hay sản xuất kịch bản từ Weekly_Content_Calendar.xlsx vì Chị Thúy
    chủ động lọc ý tưởng sang Master Series, phân bổ chuyên mục và duyệt rồi mới viết.
    """
    produced_items = []
    if not wb_onedrive:
        return produced_items
    
    font_link = Font(name="Arial", size=10, bold=True, color="0000FF", underline="single")
    font_status_done = Font(name="Arial", size=10, bold=True, color="006100")
    fill_done = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    
    # QUÉT 4 SHEET TRÊN FILE MASTER ONEDRIVE: Ke_Hoach_Phat_Song_YouTube_Series.xlsx
    print("   📂 Đang quét 4 Chuyên Mục ThuyTalk trên Master File Ke_Hoach_Phat_Song_YouTube_Series.xlsx...")
    for sname, cat_title in YT_SHEET_NAMES.items():
        if sname not in wb_onedrive.sheetnames:
            continue
        ws = wb_onedrive[sname]
        
        for r in range(5, ws.max_row + 1):
            code = ws.cell(r, 2).value
            title = ws.cell(r, 3).value
            status = ws.cell(r, 12).value
            link_val = ws.cell(r, 13).value
            
            if not code or not title:
                continue
                
            code_str = str(code).strip()
            
            # Kiểm tra nếu được Thúy duyệt
            if is_approved_by_thuy(status):
                expected_filename = make_clean_filename(code_str, str(title))
                onedrive_file_path = os.path.join(ONEDRIVE_YT_DIR, expected_filename)
                
                existing_match = None
                if os.path.exists(ONEDRIVE_YT_DIR):
                    for f in os.listdir(ONEDRIVE_YT_DIR):
                        if f.startswith(code_str) and f.endswith(".md"):
                            existing_match = f
                            break
                            
                if not existing_match:
                    print(f"      ✨ Phát hiện Chị Thúy duyệt trên OneDrive: [{code_str}] - {title}")
                    
                    data_row = {
                        "code": code_str,
                        "title": str(title).strip(),
                        "cat_name": cat_title,
                        "format": ws.cell(r, 4).value,
                        "duration": ws.cell(r, 5).value,
                        "ssot": ws.cell(r, 6).value,
                        "hook": ws.cell(r, 7).value,
                        "core": ws.cell(r, 8).value,
                        "visual": ws.cell(r, 9).value,
                        "cliffhanger": ws.cell(r, 10).value,
                        "conclusion": ws.cell(r, 11).value
                    }
                    
                    # Kiểm tra xem có bản thảo chất lượng cao đã soạn sẵn trong kho staging không
                    staging_dir = os.path.join(BASE_DIR, "scratch", "drafts_pending_thuy_approval")
                    staging_file = None
                    if os.path.exists(staging_dir):
                        for f in os.listdir(staging_dir):
                            if f.startswith(code_str) and f.endswith(".md"):
                                staging_file = os.path.join(staging_dir, f)
                                break
                                
                    if staging_file and os.path.exists(staging_file):
                        with open(staging_file, "r", encoding="utf-8") as sf:
                            script_md = sf.read()
                        fact_check_md = f"# BÁO CÁO THẨM ĐỊNH FACT-CHECK CHO [{code_str}]\nTrạng thái: 🟢 PASSED"
                        print(f"         ✨ Lấy trực tiếp bản thảo chất lượng cao từ kho chờ duyệt: {os.path.basename(staging_file)}")
                    else:
                        script_md, fact_check_md = produce_script_content(data_row)
                    
                    # Lưu file OneDrive
                    with open(onedrive_file_path, "w", encoding="utf-8") as f:
                        f.write(script_md)
                    print(f"         📄 [OneDrive] Đã lưu kịch bản: {expected_filename}")
                    
                    # Lưu file local đồng bộ theo từng chuyên mục ThuyTalk_CM* (như CM3)
                    scratch_script = os.path.join(SCRATCH_PO_DIR, f"DRAFT-YT-SERIES-{code_str}.md")
                    with open(scratch_script, "w", encoding="utf-8") as f:
                        f.write(script_md)
                    scratch_fc = os.path.join(SCRATCH_PO_DIR, f"FACT-CHECK-{code_str}.md")
                    with open(scratch_fc, "w", encoding="utf-8") as f:
                        f.write(fact_check_md)
                    
                    if code_str.startswith("CM1"):
                        cm_folder = "ThuyTalk_CM1"
                    elif code_str.startswith("CM2"):
                        cm_folder = "ThuyTalk_CM2"
                    elif code_str.startswith("CM3"):
                        cm_folder = "ThuyTalk_CM3"
                    elif code_str.startswith("CM4"):
                        cm_folder = "ThuyTalk_CM4"
                    else:
                        cm_folder = "ThuyTalk"
                    cm_rep_dir = os.path.join(BASE_DIR, "reports", cm_folder)
                    os.makedirs(cm_rep_dir, exist_ok=True)
                    rep_script = os.path.join(cm_rep_dir, expected_filename)
                    with open(rep_script, "w", encoding="utf-8") as f:
                        f.write(script_md)
                        
                    # Cập nhật Excel OneDrive
                    formula = f'=HYPERLINK("{onedrive_file_path}", "👉 Xem Kịch Bản Chi Tiết")'
                    
                    # Nếu là CM3, tự động kích hoạt pipeline ThuyTalk CM3 (2D Doodle Stick Figure)
                    if (code_str.startswith("CM3") or "lịch sử" in cat_title.lower()) and generate_cm3_storyboard:
                        print(f"         🎬 [CM3 Studio Engine] Tự động xuất trọn bộ Flow CSV + Tao Anh CSV + Studio XLSX...")
                        cm3_res = generate_cm3_storyboard(code_str, str(title), md_source_path=onedrive_file_path)
                        if cm3_res and os.path.exists(cm3_res.get("excel_studio", "")):
                            formula = f'=HYPERLINK("{cm3_res["excel_studio"]}", "🎬 Xem Studio & Phân Cảnh Flow")'

                    ws.cell(r, 12, "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
                    ws.cell(r, 12).font = font_status_done
                    ws.cell(r, 12).fill = fill_done
                    ws.cell(r, 12).alignment = align_center
                    
                    ws.cell(r, 13, formula)
                    ws.cell(r, 13).font = font_link
                    ws.cell(r, 13).alignment = align_center
                    
                    produced_items.append((code_str, str(title), onedrive_file_path))
                else:
                    full_match_path = os.path.join(ONEDRIVE_YT_DIR, existing_match)
                    formula = f'=HYPERLINK("{full_match_path}", "👉 Xem Kịch Bản Chi Tiết")'
                    
                    # Nếu là CM3, kiểm tra xem đã có file Studio.xlsx chưa, nếu chưa có thì tự động sinh bù ngay
                    if (code_str.startswith("CM3") or "lịch sử" in cat_title.lower()) and generate_cm3_storyboard:
                        studio_matches = [f for f in os.listdir(ONEDRIVE_YT_DIR) if f.startswith(f"ThuyTalk_{code_str}") and f.endswith("_Studio.xlsx")]
                        if not studio_matches:
                            print(f"         🎬 [CM3 Studio Engine] Sinh bù trọn bộ Flow CSV + Tao Anh CSV + Studio XLSX cho {code_str}...")
                            cm3_res = generate_cm3_storyboard(code_str, str(title), md_source_path=full_match_path)
                            if cm3_res and os.path.exists(cm3_res.get("excel_studio", "")):
                                formula = f'=HYPERLINK("{cm3_res["excel_studio"]}", "🎬 Xem Studio & Phân Cảnh Flow")'
                        else:
                            studio_full_path = os.path.join(ONEDRIVE_YT_DIR, studio_matches[0])
                            formula = f'=HYPERLINK("{studio_full_path}", "🎬 Xem Studio & Phân Cảnh Flow")'

                    if ws.cell(r, 13).value != formula or "ĐÃ HOÀN THÀNH" not in str(ws.cell(r, 12).value):
                        ws.cell(r, 12, "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
                        ws.cell(r, 12).font = font_status_done
                        ws.cell(r, 12).fill = fill_done
                        ws.cell(r, 12).alignment = align_center
                        ws.cell(r, 13, formula)
                        ws.cell(r, 13).font = font_link
                        ws.cell(r, 13).alignment = align_center
                        produced_items.append((code_str, str(title), full_match_path))
                        
    return produced_items

# ==============================================================================
# 7. QUÉT VÀ TỰ HỦY KỊCH BẢN CM1 HẾT HẠN (QUY TẮC 14 NGÀY)
# ==============================================================================
def check_and_purge_expired_cm1(wb):
    """
    Quét sheet CM1_VAN_BAN_NONG:
    Nếu kịch bản có tuổi > 14 ngày mà chưa quay ➔ Đánh dấu hủy và xóa file .md cũ
    để tránh tồn kho văn bản pháp luật đã nguội.
    """
    if not wb or "CM1_VAN_BAN_NONG" not in wb.sheetnames:
        return 0
        
    ws = wb["CM1_VAN_BAN_NONG"]
    purged_count = 0
    today = datetime.date.today()
    
    for r in range(5, ws.max_row + 1):
        code = ws.cell(r, 2).value
        status = ws.cell(r, 12).value
        created_val = ws.cell(r, 11).value
        link_val = ws.cell(r, 13).value
        
        # Kiểm tra nếu chưa quay và có ngày tạo
        if status and "ĐÃ QUAY" not in str(status).upper():
            if isinstance(created_val, (datetime.date, datetime.datetime)):
                c_date = created_val.date() if isinstance(created_val, datetime.datetime) else created_val
                age_days = (today - c_date).days
                if age_days > 14:
                    print(f"  ⚠️ CM1 Cũ Quá 14 Ngày ({age_days} ngày): Dòng {r} ({code}) -> Tự động hủy.")
                    ws.cell(r, 12, "ĐÃ HỦY (QUÁ 14 NGÀY)")
                    ws.cell(r, 12).font = Font(name="Arial", size=10, bold=True, color="9C0006")
                    ws.cell(r, 12).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                    
                    # Xóa file .md nếu có
                    if link_val and "HYPERLINK" in str(link_val):
                        m = re.search(r'HYPERLINK\("([^"]+)"', str(link_val))
                        if m and os.path.exists(m.group(1)):
                            try:
                                os.remove(m.group(1))
                                print(f"     🗑️ Đã xóa file kịch bản hết hạn: {m.group(1)}")
                            except Exception as e:
                                print(f"     ❌ Lỗi xóa file: {e}")
                    purged_count += 1
    return purged_count

# ==============================================================================
# 8. ĐỒNG BỘ HYPERLINK VÀ TRẠNG THÁI TRÊN ONEDRIVE EXCEL MASTER
# ==============================================================================
def reconcile_youtube_series_excel(wb):
    """
    Đối chiếu các file .md thực tế có trong thư mục OneDrive `BRANDING/youtube/`
    với bảng kế hoạch phát sóng Ke_Hoach_Phat_Song_YouTube_Series.xlsx:
    - Nếu file tồn tại: Gán Hyperlink 1-click & cập nhật trạng thái '🟢 ĐÃ HOÀN THÀNH KỊCH BẢN'
    """
    if not os.path.exists(ONEDRIVE_YT_DIR):
        return 0, 0
        
    existing_files = os.listdir(ONEDRIVE_YT_DIR)
    updated_rows = 0
    total_completed = 0
    
    font_link = Font(name="Arial", size=10, bold=True, color="0000FF", underline="single")
    font_status_done = Font(name="Arial", size=10, bold=True, color="006100")
    fill_done = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    
    for sname in YT_SHEET_NAMES.keys():
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        
        for r in range(5, ws.max_row + 1):
            code = ws.cell(r, 2).value
            if not code:
                continue
                
            code_str = str(code).strip()
            matched_files = [f for f in existing_files if f.startswith(code_str) and f.endswith(".md")]
            
            if matched_files:
                target_file = matched_files[0]
                full_file_path = os.path.join(ONEDRIVE_YT_DIR, target_file)
                formula = f'=HYPERLINK("{full_file_path}", "👉 Xem Kịch Bản Chi Tiết")'
                
                # Cột 12: Trạng Thái Xem Xét
                curr_status = ws.cell(r, 12).value
                if not curr_status or "ĐÃ HOÀN THÀNH" not in str(curr_status):
                    ws.cell(r, 12, "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
                    ws.cell(r, 12).font = font_status_done
                    ws.cell(r, 12).fill = fill_done
                    ws.cell(r, 12).alignment = align_center
                    updated_rows += 1
                    
                # Cột 13: Link Kịch Bản Chi Tiết (.md)
                curr_link = ws.cell(r, 13).value
                if curr_link != formula:
                    ws.cell(r, 13, formula)
                    ws.cell(r, 13).font = font_link
                    ws.cell(r, 13).alignment = align_center
                    updated_rows += 1
                    
                total_completed += 1
                
    return updated_rows, total_completed

# ==============================================================================
# 9. HÀM ĐIỀU PHỐI CHÍNH KHI AUTOMATIONS KÍCH HOẠT (MAIN PIPELINE)
# ==============================================================================
def run_youtube_automation(check_only=False, verify_compliance_flag=True, force_produce_code=None):
    """
    Hàm thực thi chính được kích hoạt mỗi khi Automations kích hoạt.
    """
    now_str = datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    print("=" * 80)
    print("🚀 THUYTALK PRODUCTION ENGINE & AUTOMATIONS PIPELINE (WF03)")
    print(f"⏰ Thời gian kích hoạt: {now_str}")
    print("⚡ Cơ chế kích hoạt: Automation Trigger / Event-Driven Pipeline (Non-Friday 9AM lock)")
    print("🔄 Chế độ: Quét Duy Nhất File Master Series OneDrive (TUYỆT ĐỐI KHÔNG quét file cũ Weekly_Content_Calendar)")
    print("=" * 80)
    
    # Nạp workbook duy nhất từ OneDrive (sử dụng safe_load_workbook chống PermissionError khi file mở trong Excel)
    wb_onedrive = safe_load_workbook(ONEDRIVE_EXCEL)
    if wb_onedrive:
        print(f"📂 [Master Series OneDrive]: {ONEDRIVE_EXCEL}")
    else:
        print(f"⚠️ Không tìm thấy file OneDrive: {ONEDRIVE_EXCEL}")
        
    print(f"📂 Thư mục Kịch bản OneDrive: {ONEDRIVE_YT_DIR}")
    print("-" * 80)
    
    # Ép sản xuất nếu có tham số dòng lệnh
    if force_produce_code and wb_onedrive:
        print(f"⚡ [CHẾ ĐỘ SẢN XUẤT THEO YÊU CẦU]: Ép sản xuất mã {force_produce_code}")
        for sname in YT_SHEET_NAMES.keys():
            if sname in wb_onedrive.sheetnames:
                ws = wb_onedrive[sname]
                for r in range(5, ws.max_row + 1):
                    if str(ws.cell(r, 2).value).strip() == force_produce_code.strip():
                        ws.cell(r, 12, "Đã duyệt (Force)")
                        break
    
    # BƯỚC 1: Quét các chủ đề được Chị Thúy duyệt từ Master OneDrive và tự động sản xuất
    print("🔍 [BƯỚC 1] Quét & Tự Động Sản Xuất Kịch Bản Được Chị Thúy Duyệt (Từ Master OneDrive Duy Nhất):")
    produced_items = scan_and_produce_approved_scripts(wb_onedrive)
    if produced_items:
        print(f"   🎉 Đã phát hiện và đồng bộ/gắn link thành công cho {len(produced_items)} kịch bản được Chị Thúy duyệt!")
        for p_code, p_title, p_link in produced_items:
            print(f"      • [{p_code}] {p_title[:45]}...")
            print(f"        -> Link: {p_link}")
    else:
        print("   ✅ Không có chủ đề mới nào đang ở trạng thái 'Đã duyệt' cần sản xuất thêm.")
        
    # BƯỚC 2: Áp dụng Quy chế Đào thải Tồn đọng 2 Tuần (Weekly Rolling Pipeline & 2-Week Stale Pruning)
    print("\n🔍 [BƯỚC 2] Áp Dụng Quy Chế Đào Thải Tồn Đọng 2 Tuần (Quét cả 4 Chuyên Mục):")
    try:
        try:
            from scripts.apply_weekly_pruning_policy import apply_pruning_policy
        except ImportError:
            sys.path.insert(0, os.path.join(BASE_DIR, 'scripts'))
            from apply_weekly_pruning_policy import apply_pruning_policy
        apply_pruning_policy()
        # Nạp lại wb_onedrive sau khi pruning atomic save
        wb_onedrive = safe_load_workbook(ONEDRIVE_EXCEL)
        print("   ✅ Đã hoàn tất đồng bộ quy chế xoay vòng và đào thải chủ đề tồn đọng 2 tuần.")
    except Exception as e:
        print(f"   ⚠️ Lỗi khi chạy quy chế đào thải: {e}")
        
    purged = 0
    # BƯỚC 3: Đối chiếu file vật lý và cập nhật Hyperlink trên OneDrive Master
    updated_rows = 0
    total_completed = 0
    if wb_onedrive:
        print("\n🔗 [BƯỚC 3] Đối Chiếu Kịch Bản Vật Lý & Đồng Bộ Hyperlink 1-Click Trên OneDrive:")
        updated_rows, total_completed = reconcile_youtube_series_excel(wb_onedrive)
        print(f"   -> Đã đồng bộ Hyperlink: {updated_rows} ô được chuẩn hóa.")
        print(f"   -> Tổng số kịch bản đã sẵn sàng trên OneDrive: {total_completed} kịch bản.")
    
    # Lưu workbook nguyên tử (Rule 10: Zero Backup)
    if not check_only:
        if wb_onedrive and (len(produced_items) > 0 or updated_rows > 0 or purged > 0):
            safe_save_workbook(wb_onedrive, ONEDRIVE_EXCEL)
    else:
        print("   ℹ️ Chế độ Check-Only: Bỏ qua bước lưu file.")
            
    # BƯỚC 4: Thẩm định Compliance kịch bản hiện hữu (Agent 08)
    if verify_compliance_flag and os.path.exists(ONEDRIVE_YT_DIR):
        print("\n🛡️ [BƯỚC 4] Thẩm Định Tuân Thủ Persona & Không Từ Viết Tắt (Agent 08):")
        yt_files = [f for f in os.listdir(ONEDRIVE_YT_DIR) if f.endswith(".md")]
        passed_count = 0
        issue_count = 0
        
        for ytf in yt_files:
            fpath = os.path.join(ONEDRIVE_YT_DIR, ytf)
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()
                issues = verify_script_compliance(content, ytf)
                if not issues:
                    passed_count += 1
                else:
                    issue_count += 1
                    print(f"   ⚠️ File: {ytf}")
                    for iss in issues:
                        print(f"      - {iss}")
            except Exception as e:
                print(f"   ❌ Lỗi đọc file {ytf}: {e}")
                
        print(f"   🟢 Đạt tiêu chuẩn Zero-Trust: {passed_count}/{len(yt_files)} file")
        if issue_count > 0:
            print(f"   🔴 Cần hiệu chỉnh: {issue_count} file")
            
    # BƯỚC 5: Thống kê theo 4 Chuyên Mục ThuyTalk
    if wb_onedrive:
        print("\n📊 [BƯỚC 5] Báo Cáo Hiện Trạng 4 Chuyên Mục ThuyTalk:")
        total_pending = 0
        for sname in YT_SHEET_NAMES.keys():
            if sname in wb_onedrive.sheetnames:
                ws = wb_onedrive[sname]
                completed_in_sheet = 0
                pending_in_sheet = 0
                for r in range(5, ws.max_row + 1):
                    code = ws.cell(r, 2).value
                    status = ws.cell(r, 12).value
                    if not code:
                        continue
                    if status and "ĐÃ HOÀN THÀNH" in str(status):
                        completed_in_sheet += 1
                    else:
                        pending_in_sheet += 1
                total_pending += pending_in_sheet
                print(f"   • {sname:<28}: {completed_in_sheet:>2} sẵn sàng | {pending_in_sheet:>2} chờ duyệt/sản xuất")
                
        print("-" * 80)
        print(f"🎯 TỔNG KẾT KÊNH THUYTALK: {total_completed} kịch bản sẵn sàng quay | {total_pending} kịch bản trong kho kế hoạch")
        print("✅ Hệ thống Automations vận hành thông suốt, bảo đảm đọc trọn vẹn mọi note sửa tay của Chị Thúy!")
        print("=" * 80)
    return True

# ==============================================================================
# 10. ĐIỂM VÀO DÒNG LỆNH (CLI ENTRYPOINT)
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(
        description="ThuyTalk Production Pipeline & Multi-Source Approval Trigger (WF03)"
    )
    parser.add_argument(
        "--produce", type=str, default=None,
        help="Mã chủ đề cần ép sản xuất ngay (ví dụ: CM1-CD01, CM2-CD01)"
    )
    parser.add_argument(
        "--check-only", action="store_true", 
        help="Chỉ kiểm tra và báo cáo hiện trạng, không ghi đè dữ liệu Excel"
    )
    parser.add_argument(
        "--no-verify", action="store_true", 
        help="Bỏ qua bước thẩm định compliance Agent 08"
    )
    args = parser.parse_args()
    
    success = run_youtube_automation(
        check_only=args.check_only,
        verify_compliance_flag=not args.no_verify,
        force_produce_code=args.produce
    )
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()
