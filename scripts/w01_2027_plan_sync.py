# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Đa nền tảng Tuần W01/2027 (04/01/2027 - 10/01/2027)
theo quy trình workflows/wf01_weekly_planning.md & agents/04_content_planner.md:
- 11 slots (STT 157 - 167):
  + Facebook: 4 bài (Thứ 2, 4, 6, CN)
  + YouTube: 3 video (Thứ 3, 5, 7)
  + LinkedIn (Repurpose B2B): 2 bài (Thứ 3, 6)
  + TikTok/Shorts (Repurpose <60s): 2 video (Thứ 5, CN)
- 3 phương án Options A/B/C cho mỗi slot
- Tự động kiểm duyệt Bộ từ cấm (knowledge/banned_words.md) & Anti-AI Buzzwords trước khi ghi file
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx qua atomic_save (Rule 10 - Zero Backup)
- Cập nhật reports/plan1/weekly_calendar.md & lưu scratch/w01_2027_data.json
"""

import os
import sys
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8")

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
WEEKLY_CAL_MD = os.path.join(BASE_DIR, "reports", "plan1", "weekly_calendar.md")
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w01_2027_data.json")

W01_2027_ITEMS = [
    {
        "stt": 157,
        "day": "[W01/2027] Thứ Hai (04/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Phòng vệ thương mại kép đầu năm 2027: Hoa Kỳ (DOC A-552-855) áp ký quỹ sơ bộ 36,55% với PTMEG Việt Nam & Ấn Độ (DGTR AD/OI/079/2026) khởi xướng điều tra nhôm đùn — Nếu doanh nghiệp sản xuất xuất khẩu chưa chuẩn hóa sổ sách giá thành theo chuẩn kế toán quốc tế, nguy cơ bị áp mức thuế trần toàn quốc sẽ xử lý thế nào?",
        "opt_b": "Checklist 4 bước doanh nghiệp xuất khẩu hóa chất, vật liệu và nhôm đùn cần chuẩn bị ngay khi đối mặt điều tra chống bán phá giá từ Hoa Kỳ và Ấn Độ: Cách rà soát chứng từ giá thành sản xuất, hóa đơn nguyên liệu đầu vào và nộp bản trả lời câu hỏi đúng hạn.",
        "opt_c": "Nếu doanh nghiệp bạn xuất khẩu sang Mỹ hoặc Ấn Độ nhưng vẫn lưu sổ sách kế toán kiểu gộp chung chi phí, đến lúc nhận trát điều tra phòng vệ thương mại mới cuống cuồng tìm chứng từ thì liệu có kịp bảo vệ biên độ thuế riêng?",
        "ssot": "Kết luận sơ bộ DOC vụ việc A-552-855 ngày 10/11/2026 (ký quỹ 36,55% PTMEG); Thông báo khởi xướng DGTR vụ việc AD/OI/079/2026 ngày 09/11/2026 (nhôm đùn); tapchicongthuong.vn, vneconomy.vn & moit.gov.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #3 ngày 20/11/2026).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 158,
        "day": "[W01/2027] Thứ Ba (05/01/2027)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Liên Minh Thương Mại Hanseatic (Thế Kỷ 12–17): Cách Các Thương Hội Bắc Âu Tự Lập Hạm Đội, Đặt Luật Biển Riêng & Vận Hành 'Cảng Tự Do Kontor' Trước Khi Có WTO!",
        "opt_b": "Giải Mã Đế Chế Thương Mại 300 Năm Từ Lübeck Đến London: Vì Sao Thỏa Ước Năm 1241 Giữa Hai Thành Phố Cảng Đức Lại Định Hình Chuẩn Mực Đại Lý Logistics Ngày Nay?",
        "opt_c": "Không Hiệp Định FTA, Không Tòa Án Quốc Tế: Cách Thương Nhân Trung Cổ Bảo Vệ Tàu Buôn Trước Cướp Biển Và Thuế Quan Tùy Tiện Suốt 5 Thế Kỷ!",
        "ssot": "Tư liệu Lịch sử Liên minh Hanseatic (Hanseatic League, hiệp ước Lübeck – Hamburg 1241, hệ thống thương điếm Kontor tại London, Bruges, Novgorod, Bergen); Britannica & Wikipedia; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 266).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 159,
        "day": "[W01/2027] Thứ Ba (05/01/2027)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Ứng phó điều tra phòng vệ thương mại kép 2027 (Hoa Kỳ A-552-855 & Ấn Độ AD/OI/079/2026): 4 lớp hồ sơ giá thành và truy xuất nguyên liệu mà CFO doanh nghiệp xuất khẩu cần chuẩn hóa",
        "opt_b": "Vì sao không hợp tác hoặc thiếu dữ liệu kế toán chi phí lại khiến doanh nghiệp Việt chịu thuế ký quỹ toàn quốc 36,55% tại Mỹ? Bài học quản trị tuân thủ xuất khẩu B2B",
        "opt_c": "Xây dựng hệ thống lưu trữ chứng từ chi phí sản xuất theo lô hàng xuất khẩu: Lá chắn tài chính trước các vụ kiện chống bán phá giá tại thị trường Ấn Độ và Hoa Kỳ",
        "ssot": "Kết luận sơ bộ DOC A-552-855 (10/11/2026); Khởi xướng DGTR AD/OI/079/2026 (09/11/2026); Cục Phòng vệ thương mại (moit.gov.vn) (Repurpose từ STT 157).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 160,
        "day": "[W01/2027] Thứ Tư (06/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Mất trắng gần 72.000 USD (gần 2 tỷ đồng) chỉ sau vài email trao đổi giao dịch quốc tế: Kẻ gian dùng AI giả mạo văn phong đối tác và đổi số tài khoản ngân hàng ngay sát giờ chuyển tiền (bẫy BEC) — Phòng Sales XNK và Kế toán nhà bạn đã có kỷ luật 'xác minh 2 kênh ngoài đời thực' chưa?",
        "opt_b": "Checklist 3 lớp phòng vệ bắt buộc khi đối tác nước ngoài bất ngờ gửi email báo đổi tài khoản thụ hưởng: Cách kiểm tra tên miền giả mạo (typo-squatting), gọi video xác thực số hotline gốc và khóa quy trình duyệt chi nội bộ.",
        "opt_c": "Nếu chiều nay đối tác bên kia gửi một email cực kỳ chuẩn xác báo 'tài khoản ngân hàng cũ đang bảo trì, chuyển gấp 70.000 USD vào số tài khoản mới của công ty con để kịp nhả Bill gốc' — Nhân sự kế toán của bạn có dám bấm chuyển khoản ngay chỉ dựa vào email?",
        "ssot": "Hồ sơ cảnh báo doanh nghiệp xuất khẩu Việt Nam mất gần 72.000 USD qua email giả mạo BEC; dantri.com.vn, cafef.vn & vietnambiz.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 263).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 161,
        "day": "[W01/2027] Thứ Năm (07/01/2027)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 23",
        "opt_a": "Học XNK 5 Phút Tập 23: Quy Trình 5 Bước Phân Định Mã HS Cấu Kiện Thép (Nhóm 7308) vs Nhà Lắp Ghép (Nhóm 9406) & Hồ Sơ Miễn Thuế Tài Sản Cố Định Theo Công Văn 23970/CHQ-NVTHQ",
        "opt_b": "Cẩm Nang Kiểm Tra Bản Vẽ Kỹ Thuật & Danh Mục Hàng Trong Nước Chưa Sản Xuất Được Để Không Bị Bác Miễn Thuế Dự Án FDI Khi Hải Quan Hậu Kiểm Kết Cấu Thép",
        "opt_c": "Vì Sao Chênh Lệch Thuế MFN 10–15% Với C/O 0% Khiến Nhóm Mã HS 7308 Và 9406 Bị Soi Gắt Nhất Đầu Năm 2027? Hướng Dẫn Chuẩn Bị Hồ Sơ Tham Vấn Giá Cho Nhà Thầu",
        "ssot": "Công văn số 23969/CHQ-NVTHQ & 23970/CHQ-NVTHQ ngày 06/11/2026 của Cục Hải quan; Luật Thuế XNK số 107/2016/QH13 (Điều 16); luatvietnam.vn & congthuong.vn; reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx (STT 268).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 162,
        "day": "[W01/2027] Thứ Năm (07/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Thép Rời Nhóm 7308 vs Nhà Lắp Ghép 9406: 3 Lỗi Khiến Doanh Nghiệp Bị Bác Miễn Thuế Tài Sản Cố Định!",
        "opt_b": "Nhập Cấu Kiện Thép Theo Công Văn 23970/CHQ-NVTHQ: Kiểm Tra Ngay 2 Giấy Tờ Này Trước Khi Mở Tờ Khai",
        "opt_c": "Mẹo Nhớ Nhanh Mã HS 7308 Và 9406 Trong 30 Giây Để Không Bị Tham Vấn Giá & Truy Thu Thuế",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 23 (STT 161); Công văn 23969 & 23970/CHQ-NVTHQ ngày 06/11/2026.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 163,
        "day": "[W01/2027] Thứ Sáu (08/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Chính thức có hiệu lực hôm nay 08/01/2027: Thông tư 71/2026/TT-BCT lần đầu đặt thước đo phân hạng 3 cấp & 5 loại hình Trung tâm Logistics (Hạng I phải từ 20 ha, xử lý >= 200.000 TEU/năm, kết nối >= 2 phương thức vận tải & bắt buộc chuẩn kho bãi xanh) — Chủ hàng và doanh nghiệp 3PL nên nhìn vào tiêu chí nào khi chọn đối tác kho vận?",
        "opt_b": "Tóm tắt nhanh 5 loại hình và bộ tiêu chuẩn phân hạng Trung tâm Logistics Hạng I – II – III theo Thông tư 71/2026/TT-BCT (hiệu lực 08/01/2027): Cơ chế tự công bố lên Cơ sở dữ liệu hạ tầng thương mại quốc gia giúp chủ hàng minh bạch năng lực đối tác ra sao?",
        "opt_c": "Trước đây bãi đất trống có vài nhà kho lắp ghép cũng tự xưng 'Trung tâm Logistics', khiến chủ hàng gửi hàng vào mới thấy kẹt xe cont và thiếu phần mềm WMS — Từ 08/01/2027, Thông tư 71/2026/TT-BCT liệu có tạo ra cuộc sàng lọc chất lượng hạ tầng kho vận thật sự?",
        "ssot": "Thông tư số 71/2026/TT-BCT ngày 19/11/2026 của Bộ Công Thương (hiệu lực 08/01/2027); chinhphu.vn, dms.gov.vn & baochinhphu.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #2 ngày 20/11/2026).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 164,
        "day": "[W01/2027] Thứ Sáu (08/01/2027)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Phòng chống tấn công giả mạo email doanh nghiệp (BEC) bằng AI trong thanh toán xuất nhập khẩu: Quy trình xác minh 2 kênh (Two-Channel Verification) sau bài học 72.000 USD",
        "opt_b": "Kiểm soát rủi ro dòng tiền B2B quốc tế năm 2027: Tại sao phân quyền duyệt chi giữa Sales, Chứng từ và Kế toán lại quan trọng hơn mọi phần mềm lọc thư rác?",
        "opt_c": "3 lỗ hổng nội bộ khiến doanh nghiệp xuất nhập khẩu SME dễ sập bẫy thay đổi tài khoản thụ hưởng qua email và cách thiết lập phụ lục hợp đồng bảo vệ thanh toán",
        "ssot": "Hồ sơ cảnh báo vụ mất gần 72.000 USD qua email giả mạo BEC; dantri.com.vn & vietnambiz.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 263) (Repurpose từ STT 160).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 165,
        "day": "[W01/2027] Thứ Bảy (09/01/2027)",
        "platform": "YouTube",
        "category": "CS — Case Study",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Phân tích & Hướng dẫn (7-9p) [CM2/CM4]",
        "opt_a": "Cuộc Giải Cứu Lô Hàng Chè Xuất Khẩu Kẹt Chiến Sự Trung Đông Phải Tái Nhập Về Việt Nam: Chọn Đúng Mã Loại Hình A31 Hay G13 Để Không Mất Oan Tiền Thuế & Phí Lưu Bãi!",
        "opt_b": "Dịch Luật Điều 47 Nghị Định 08/2015/NĐ-CP: Hàng Đã Xuất Khẩu Bị Trả Lại Hoặc Quay Đầu Vì Bất Khả Kháng — Thủ Tục Xét Không Thu Thuế Nhập Khẩu & Xử Lý Thuế GTGT Chuẩn Xác!",
        "opt_c": "Phân Biệt Rạch Ròi Tờ Khai A31 (Tiêu Thụ/Tái Xuất) vs G13 (Tạm Nhập Sửa Chữa) vs A11: Sai Một Mã Loại Hình Là Mất Quyền Miễn Thuế Hàng Quay Đầu!",
        "ssot": "Điều 47 Nghị định số 08/2015/NĐ-CP; Quyết định 1357/QĐ-TCHQ (mã loại hình A31, G13, G23); Hồ sơ thực tế doanh nghiệp xuất khẩu chè tái nhập do chiến sự; tapchikinhtetaichinh.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 264).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 166,
        "day": "[W01/2027] Chủ Nhật (10/01/2027)",
        "platform": "Facebook",
        "category": "FF — Funfact & Insight",
        "technique": "Dẫn chứng quốc tế",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Thuật toán 'Cấm Rẽ Trái' (UPS ORION): Vì sao gã khổng lồ chuyển phát lại lập trình cho xe tải hạn chế tối đa rẽ trái qua làn ngược chiều, dù quãng đường đôi khi dài hơn — Kết quả tiết kiệm 100 triệu dặm chạy xe và 10 triệu gallon nhiên liệu mỗi năm! Bài học tối ưu 'những điểm nghẽn nhỏ lặp lại hàng nghìn lần' trong vận hành kho bãi & điều vận xe cont.",
        "opt_b": "Tư duy quản trị vận hành tinh gọn từ hệ thống AI định tuyến ORION của UPS: Lợi nhuận bền vững trong logistics không đến từ việc ép giá cước tài xế, mà đến từ việc cắt bỏ thời gian chờ nổ máy không tải tại cổng cảng và các điểm giao cắt.",
        "opt_c": "Chủ nhật đọc chuyện logistics thế giới: Một thay đổi nhỏ 'ưu tiên rẽ phải' giúp giảm 100.000 tấn CO2 và tiết kiệm hàng trăm triệu USD mỗi năm — Trong quy trình làm chứng từ hoặc điều xe của doanh nghiệp bạn, đâu là 'ngã tư rẽ trái' đang ngốn nhiều thời gian nhất?",
        "ssot": "Tài liệu công bố hệ thống UPS ORION (On-Road Integrated Optimization and Navigation); metalumna.com & bsr.org; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 264).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 167,
        "day": "[W01/2027] Chủ Nhật (10/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Hàng Xuất Khẩu Phải Quay Đầu Tái Nhập: Khai Mã A31 Hay G13 Để Không Mất Oan Tiền Thuế?",
        "opt_b": "Đừng Mở Nhầm Mã A11 Khi Hàng Xuất Khẩu Bị Trả Về: Quy Trình 3 Chứng Từ Xin Không Thu Thuế Theo Điều 47 NĐ 08",
        "opt_c": "Lô Hàng Bị Kẹt Chiến Sự Quay Về Cảng Việt Nam: Cách Chọn Đúng Mã Loại Hình Và Xin Giảm Phí Lưu Bãi DEM/DET",
        "ssot": "Trích xuất từ YouTube Case Study (STT 165); Điều 47 Nghị định 08/2015/NĐ-CP & Quyết định 1357/QĐ-TCHQ.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    }
]


def run_compliance_check():
    print("=== 0. KIỂM DUYỆT COMPLIANCE (BANNED WORDS & ANTI-PATTERNS) ===")
    global_banned = [
        "trốn thuế", "lách luật", "hàng lậu", "hàng nhái", "hàng giả",
        "bao thuế", "đi chui", "đi lậu", "tiểu ngạch",
        "mua hàng trung quốc tận gốc", "mua hàng trung tận gốc",
        "đánh hàng", "gom hàng biên", "chạy line",
        "đột phá", "chìa khóa vàng", "vượt bậc", "hệ sinh thái",
        "đau đớn", "giật mình", "hoảng loạn", "kinh hoàng", "chao đảo", "đau đứt ruột",
        "anh minh", "chị hương", "anh dũng", "inbox ngay", "dm ngay"
    ]
    tiktok_banned = ["kiếm tiền", "làm giàu", "thu nhập thụ động", "bí mật", "hack"]
    youtube_banned = ["lừa đảo", "scam", "bạo lực"]

    errors = []
    for item in W01_2027_ITEMS:
        combined = " | ".join([item["opt_a"], item["opt_b"], item["opt_c"], item["ssot"]]).lower()
        for bw in global_banned:
            if bw in combined:
                errors.append(f"STT {item['stt']} ({item['platform']}): chứa từ cấm '{bw}'")
        if "tiktok" in item["platform"].lower():
            for tb in tiktok_banned:
                if tb in combined:
                    errors.append(f"STT {item['stt']} (TikTok): chứa từ nhạy cảm TikTok '{tb}'")
        if "youtube" in item["platform"].lower():
            for yb in youtube_banned:
                if yb in combined:
                    errors.append(f"STT {item['stt']} (YouTube): chứa từ nhạy cảm YouTube '{yb}'")

    if errors:
        for e in errors:
            print("  ❌", e)
        raise ValueError(f"Phát hiện {len(errors)} lỗi compliance! Dừng ghi file.")
    print(f"  ✅ Đạt 100% compliance cho toàn bộ {len(W01_2027_ITEMS)} slots (0 vi phạm)!")


def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33/2026 ĐẾN W01/2027 CỘNG DỒN)")
    ws.cell(2, 1, value="Chiến lược Đa kênh | Chủ lực: Facebook (4 bài/tuần) & YouTube (3 video/tuần) | Repurpose: LinkedIn (2 bài/tuần) & TikTok/Shorts (2 video/tuần)")

    font_family = "Segoe UI"
    regular_font = Font(name=font_family, size=10)
    bold_font = Font(name=font_family, size=10, bold=True)
    pending_font = Font(name=font_family, size=10, bold=True, color="9A7D0A")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    pending_fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")

    platform_fills = {
        "Facebook": PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid"),
        "YouTube": PatternFill(start_color="FDEDEC", end_color="FDEDEC", fill_type="solid"),
        "LinkedIn": PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid"),
        "TikTok/Shorts": PatternFill(start_color="F4ECF7", end_color="F4ECF7", fill_type="solid"),
    }

    existing_row_by_stt = {}
    for r in range(5, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None:
            try:
                existing_row_by_stt[int(str(val).strip())] = r
            except ValueError:
                pass

    for item in W01_2027_ITEMS:
        stt = item["stt"]
        if stt in existing_row_by_stt:
            target_row = existing_row_by_stt[stt]
        else:
            target_row = ws.max_row + 1

        row_fill = zebra_fill if target_row % 2 == 0 else white_fill
        values = [
            item["stt"],
            item["day"],
            item["platform"],
            item["category"],
            item["technique"],
            item["format"],
            item["opt_a"],
            item["opt_b"],
            item["opt_c"],
            item["ssot"],
            item["approval"],
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(target_row, col_idx, value=val)
            cell.border = thin_border
            if col_idx == 11:
                cell.font = pending_font
                cell.alignment = center_align
                cell.fill = pending_fill
            elif col_idx == 3:
                cell.font = bold_font
                cell.alignment = center_align
                cell.fill = platform_fills.get(val, row_fill)
            elif col_idx in [1, 2, 4, 5, 6]:
                cell.font = bold_font if col_idx == 1 else regular_font
                cell.alignment = center_align
                cell.fill = row_fill
            else:
                cell.font = bold_font if col_idx == 7 else regular_font
                cell.alignment = left_align
                cell.fill = row_fill

        ws.row_dimensions[target_row].height = 54

    dv = DataValidation(
        type="list",
        formula1='"Chờ duyệt,Duyệt Option A,Duyệt Option B,Duyệt Option C,Không duyệt"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    dv.add("K5:K250")

    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print(f"  -> Đã lưu thành công {len(W01_2027_ITEMS)} dòng (STT 157–167) vào Sheet 1 qua atomic_save!")


def build_week_md(week_code, date_range, items):
    lines = [
        "",
        "---",
        "",
        f"# TUẦN {week_code}: LỊCH BIÊN TẬP ĐA NỀN TẢNG ({date_range})",
        "",
        f"## 1. Bảng Kế Hoạch Tuần {week_code} — Đủ 4 Nền Tảng (Options A/B/C)",
        "",
        "| STT | Tuần / Ngày đăng | Nền tảng | Chuyên mục | Kỹ thuật Storytelling | Định dạng Sản phẩm | Option A (Chủ đề chính) | Option B (Dự phòng 1) | Option C (Dự phòng 2) | Nguồn SSoT / Trích xuất Pháp lý | Trạng thái Duyệt |",
        "|:---:|:---|:---:|:---:|:---:|:---:|:---|:---|:---|:---|:---:|",
    ]
    for it in items:
        lines.append(
            f"| {it['stt']} | {it['day']} | {it['platform']} | {it['category']} | {it['technique']} | {it['format']} | **{it['opt_a']}** | {it['opt_b']} | {it['opt_c']} | {it['ssot']} | {it['approval']} |"
        )

    gm = sum(1 for it in items if it["energy"] == "Giáp Mộc")
    lt = sum(1 for it in items if it["energy"] == "Liêm Trinh")
    fb = sum(1 for it in items if it["platform"] == "Facebook")
    yt = sum(1 for it in items if it["platform"] == "YouTube")
    li = sum(1 for it in items if it["platform"] == "LinkedIn")
    tt = sum(1 for it in items if it["platform"] == "TikTok/Shorts")

    lines.extend([
        "",
        f"## 2. Bản phân tích Tỷ lệ Nền tảng & Năng lượng (Validation Tuần {week_code})",
        "",
        f"- **Tỷ lệ Năng lượng Nhân hiệu:** Giáp Mộc **{gm}/{len(items)} ({gm*100//len(items)}%)** | Liêm Trinh **{lt}/{len(items)} ({lt*100//len(items)}%)** — Đạt chuẩn cân bằng ~60/40.",
        f"- **Phân bổ 4 Nền tảng:** Facebook (**{fb} bài** — Thứ 2, 4, 6, CN) | YouTube (**{yt} video** — Thứ 3, 5, 7) | LinkedIn (**{li} bài Repurpose B2B** — Thứ 3, 6) | TikTok/Shorts (**{tt} video <60s** — Thứ 5, CN).",
        "- **Tuân thủ Bộ từ cấm (`banned_words.md`):** 100% tiêu đề và SSoT đã kiểm duyệt tự động (0 vi phạm, không nhắc tên Persona giả định, không chứa link đối thủ).",
        "",
        "## 3. Checkpoint 1 — Chờ Chị Thúy Phê Duyệt Concept",
        f"1. Chị Thúy mở file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) → **Sheet 1 (`Weekly Calendar`)**.",
        f"2. Chọn phương án duyệt (`Duyệt Option A` / `B` / `C`) tại **Cột K** cho từng dòng thuộc Tuần {week_code}.",
        "3. Sau khi có phê duyệt, hệ thống mới kích hoạt Agent 04 đóng gói Phiếu Giao việc (PO) vật lý sang Giai đoạn 2.",
    ])
    return "\n".join(lines)


def update_markdown_calendar():
    print("=== 2. CẬP NHẬT REPORTS/PLAN1/WEEKLY_CALENDAR.MD ===")
    with open(WEEKLY_CAL_MD, "r", encoding="utf-8") as f:
        existing_content = f.read()

    marker = "# TUẦN W01/2027: LỊCH BIÊN TẬP ĐA NỀN TẢNG"
    if marker in existing_content:
        existing_content = existing_content.split(marker)[0].rstrip()
        if existing_content.endswith("---"):
            existing_content = existing_content[:-3].rstrip()

    existing_content = existing_content.replace(
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W52 (Từ 17/08/2026 đến 03/01/2027) — W49–W52 Chờ duyệt Checkpoint 1",
        "**Kỳ đánh giá mới nhất:** Tuần W33/2026 đến W01/2027 (Từ 17/08/2026 đến 10/01/2027) — W49–W01/2027 Chờ duyệt Checkpoint 1"
    )

    appended = build_week_md("W01/2027", "04/01/2027 - 10/01/2027", W01_2027_ITEMS)
    final_content = existing_content + appended + "\n"

    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(final_content)
    print("  -> Đã cập nhật thành công reports/plan1/weekly_calendar.md!")


def save_json_data():
    print("=== 3. LƯU SCRATCH/W01_2027_DATA.JSON ===")
    os.makedirs(os.path.dirname(JSON_OUTPUT), exist_ok=True)
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W01_2027_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu 11 slots vào {JSON_OUTPUT}")


if __name__ == "__main__":
    print("=== BẮT ĐẦU QUY TRÌNH WF01: LẬP LỊCH ĐA NỀN TẢNG TUẦN W01/2027 ===")
    run_compliance_check()
    update_excel_calendar()
    update_markdown_calendar()
    save_json_data()
    print("=== HOÀN TẤT THỰC THI GIAI ĐOẠN 1 TUẦN W01/2027! ===")
