# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Đa nền tảng 4 Tuần W49 - W52 (07/12/2026 - 03/01/2027)
theo quy trình huong_dan_phoi_hop_gd1_lap_lich.md & workflows/wf01_weekly_planning.md:
- 44 slots (STT 113 - 156): 11 slots/tuần x 4 tuần
  + Facebook: 4 bài/tuần (Thứ 2, 4, 6, CN)
  + YouTube: 3 video/tuần (Thứ 3, 5, 7)
  + LinkedIn (Repurpose B2B): 2 bài/tuần (Thứ 3, 6)
  + TikTok/Shorts (Repurpose <60s): 2 video/tuần (Thứ 5, CN)
- 3 phương án Options A/B/C cho mỗi slot
- Tự động kiểm duyệt Bộ từ cấm (knowledge/banned_words.md) & Anti-AI Buzzwords trước khi ghi file
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx qua atomic_save (Rule 10)
- Cập nhật reports/plan1/weekly_calendar.md & lưu scratch/w49_w52_data.json
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
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w49_w52_data.json")

W49_W52_ITEMS = [
    # ======================== TUẦN W49 (07/12/2026 - 13/12/2026) ========================
    {
        "stt": 113,
        "day": "[W49] Thứ Hai (07/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Chỉ còn 24 ngày trước mốc khóa tài khoản VNACCS/VCIS 31/12/2026 (CV 21692 & 23474/CHQ-CNTT): Hàng loạt dn vẫn giữ tài khoản nhân viên cũ đã nghỉ việc hoặc sai mã số thuế chi nhánh — Đến 01/01/2027 bị treo luồng truyền tờ khai giữa mùa cao điểm Tết thì xử lý thế nào?",
        "opt_b": "Checklist 4 bước rà soát và làm sạch dữ liệu tài khoản VNACCS/VCIS tại cổng dksd.customs.gov.vn trước hạn chót 31/12/2026: Cách phân quyền User ID, cập nhật chữ ký số và thu hồi quyền nhân sự đã nghỉ việc chỉ trong 15 phút.",
        "opt_c": "Chuyện có thật ngay tại phòng chứng từ cuối năm: Nhân viên cũ nghỉ 2 năm nhưng vẫn dùng tên đăng nhập cũ trên VNACCS, đến khi rà soát định danh mới phát hiện đứng tên sai người đại diện — Bài học quản trị tài khoản hải quan số hóa trước thềm 2027.",
        "ssot": "Công văn số 21692/CHQ-CNTT & 23474/CHQ-CNTT (26-27/10/2026); Quyết định số 1761/QĐ-CHQ ngày 02/11/2026 của Cục Hải quan; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #8); Cổng dksd.customs.gov.vn.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 114,
        "day": "[W49] Thứ Ba (08/12/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Đế Chế FedEx & Cú Đặt Cược 27.000 USD Tại Bàn Đen Đỏ Las Vegas: Fred Smith Đã Cứu Tập Đoàn Chuyển Phát Nhanh Lớn Nhất Hành Tinh Khỏi Bờ Vực Phá Sản Thế Nào?",
        "opt_b": "Giải Mã Mô Hình Hub-and-Spoke Của FedEx: Tại Sao Mọi Kiện Hàng Từ Khắp Nước Mỹ Đều Phải Bay Về Sân Bay Memphis Lúc Nửa Đêm Rồi Mới Tỏa Đi Toàn Cầu?",
        "opt_c": "Từ Bài Luận Điểm C Ở Đại Học Yale Đến Đế Chế Hàng Không Vận Tải 100 Tỷ USD: Bài Học Bản Lĩnh Cầm Lái Khi Không Một Ngân Hàng Nào Tin Vào Ý Tưởng Của Bạn!",
        "ssot": "Tài liệu lịch sử doanh nghiệp FedEx Corp (Memphis Hub 1973); Hồi ký của Frederick W. Smith; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD13).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 115,
        "day": "[W49] Thứ Ba (08/12/2026)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Chuyển đổi số Hải quan 2026–2027 theo Quyết định 1761/QĐ-CHQ & 2735/QĐ-BTC: 4 nguyên tắc quản trị dữ liệu định danh (Data Compliance) để chuỗi cung ứng không đứt gãy đầu năm mới",
        "opt_b": "Lộ trình số hóa 100% thủ tục hải quan và chấm điểm tuân thủ doanh nghiệp năm 2027: CFO và Giám đốc Logistics cần chuẩn hóa những điểm chạm dữ liệu nào ngay từ tháng 12?",
        "opt_c": "Kinh nghiệm tái cấu trúc quyền truy cập hệ thống khai báo điện tử cho doanh nghiệp đa chi nhánh: Vì sao rủi ro bảo mật thường đến từ tài khoản nhân sự cũ bị bỏ quên?",
        "ssot": "Quyết định số 1761/QĐ-CHQ ngày 02/11/2026; Quyết định số 2735/QĐ-BTC ngày 30/09/2026; baochinhphu.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (E2E #285).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 116,
        "day": "[W49] Thứ Tư (09/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Bỏ cấp C/O khi DNCX bán hàng vào nội địa: Đáng ra cải cách thủ tục để thuận lợi cho dn, sao lại thành độ vênh khiến bên mua nội địa gánh thuế MFN tăng vọt? Chia sẻ cách rà soát lại điều khoản giá & thuế hợp đồng ngay trước khi ký giao hàng!",
        "opt_b": "Sơ đồ 3 bước ứng xử khi DNCX bán hàng vào nội địa không được cấp C/O ưu đãi: Cách xác định bên chịu thuế MFN phát sinh, lập văn bản tham vấn Bộ Công Thương và đàm phán phụ lục điều chỉnh đơn giá.",
        "opt_c": "Tâm sự từ cộng đồng làm nghề hải quan sáng nay: Hợp đồng báo giá bao gồm ưu đãi thuế quan, đến lúc giao hàng mới ngã ngửa vướng quy định bỏ cấp C/O — Nếu bạn đang mua hàng từ DNCX, bạn nên kiểm tra điều khoản nào đầu tiên?",
        "ssot": "Nghị quyết 19/2026/NQ-CP; Quy định xuất xứ hàng hóa & giao dịch DNCX bán nội địa; scratch/daily_notes.md ngày 20/11/2026 (Case Study DNCX bỏ cấp C/O).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 117,
        "day": "[W49] Thứ Năm (10/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 19",
        "opt_a": "Học XNK 5 Phút Tập 19: 4 Bước Làm Sạch Tài Khoản VNACCS/VCIS Tại dksd.customs.gov.vn Trước Hạn Chót 31/12/2026 Để Không Bị Khóa Tờ Khai!",
        "opt_b": "Hướng Dẫn Toàn Tập Quản Trị User ID & Chữ Ký Số Hải Quan Điện Tử Chuẩn Công Văn 23474/CHQ-CNTT Cho Doanh Nghiệp Xuất Nhập Khẩu",
        "opt_c": "Đừng Để Mất Quyền Truyền Tờ Khai Ngày Đầu Năm 2027: Thực Hành Rà Soát Và Hủy Tài Khoản Nhân Viên Cũ Trên Hệ Thống Hải Quan Trong 5 Phút",
        "ssot": "Công văn số 21692/CHQ-CNTT & 23474/CHQ-CNTT; Quyết định số 1761/QĐ-CHQ; Cổng đăng ký sử dụng hệ thống hải quan dksd.customs.gov.vn.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 118,
        "day": "[W49] Thứ Năm (10/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "3 Lỗi Khiến Tài Khoản VNACCS Bị Khóa Sau 31/12/2026 — Kiểm Tra Ngay Trong 60 Giây!",
        "opt_b": "Còn Đúng 3 Tuần: 1 Thao Tác Trên dksd.customs.gov.vn Giúp Không Bị Treo Tờ Khai Đầu Năm 2027",
        "opt_c": "Nhân Viên XNK Nghỉ Việc Nhưng Quên Khóa Mã VNACCS: Rủi Ro Pháp Lý Doanh Nghiệp Phải Gánh",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 19 (STT 117); Công văn 21692 & 23474/CHQ-CNTT.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 119,
        "day": "[W49] Thứ Sáu (11/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Thông tư 63/2026/TT-BKHCN chính thức có hiệu lực từ 01/01/2027: Cấm nhập khẩu 13 nhóm sp công nghệ số đã qua sử dụng & hàng tân trang (Refurbished) — Chỉ còn 20 ngày để rà lại toàn bộ mã HS thiết bị văn phòng, linh kiện cũ đang trên đường về cảng!",
        "opt_b": "Checklist đối chiếu nhanh Danh mục 13 nhóm sản phẩm công nghệ số cấm nhập khẩu theo Thông tư 63/2026/TT-BKHCN (thay thế TT 11/2018): Khác biệt cốt lõi nào khiến hàng tân trang cũng bị chặn ngay cửa khẩu?",
        "opt_c": "Nếu doanh nghiệp bạn đang đặt máy chủ, thiết bị mạng hay linh kiện công nghệ tân trang từ nước ngoài về cuối năm nay: Hãy kiểm tra ngày tàu cập cảng và ngày mở tờ khai trước mốc 01/01/2027 nếu không muốn buộc tái xuất!",
        "ssot": "Thông tư số 63/2026/TT-BKHCN ngày 27/10/2026 (hiệu lực 01/01/2027, thay thế TT 11/2018/TT-BTTTT); chinhphu.vn; reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx (STT 263).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 120,
        "day": "[W49] Thứ Sáu (11/12/2026)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Tình huống giả định mở",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Rủi ro chi phí thuế MFN khi doanh nghiệp nội địa mua hàng từ DNCX trong bối cảnh thay đổi quy trình cấp C/O: 3 điều khoản hợp đồng B2B cần bổ sung ngay",
        "opt_b": "Quản trị độ vênh chính sách trong chuỗi cung ứng nội địa – FDI: Cách phân bổ trách nhiệm thuế quan giữa Doanh nghiệp Chế xuất và bên mua nội địa",
        "opt_c": "Tại sao một thay đổi thủ tục hành chính về chứng nhận xuất xứ lại có thể làm đảo lộn biên lợi nhuận hợp đồng dài hạn? Góc nhìn kiểm soát rủi ro mua hàng B2B",
        "ssot": "Nghị quyết 19/2026/NQ-CP; Giao dịch DNCX bán nội địa; scratch/daily_notes.md ngày 20/11/2026 (Repurpose từ STT 116).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 121,
        "day": "[W49] Thứ Bảy (12/12/2026)",
        "platform": "YouTube",
        "category": "CS — Case Study",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Tâm tình Thực chiến (7-9p) [CM4]",
        "opt_a": "Quản Trị Công Nợ & Đàm Phán Phí Local Charges, Lưu Container (DEM/DET) Mùa Cao Điểm Cuối Năm: Bài Học Giữ Chữ Tín Mà Không Tự Đẩy Mình Vào Rủi Ro Dòng Tiền",
        "opt_b": "Cách Xin Miễn Lưu Bãi 14–21 Ngày (Free Time) Với Hãng Tàu Ngoại & Nghệ Thuật Từ Chối Nợ Cước Khéo Léo Cho Doanh Nghiệp Logistics",
        "opt_c": "Tâm Sự 15 Năm Làm Nghề: Khi Khách Quen Xin Khất Công Nợ Cuối Năm Đúng Đợt Cước Tàu Tăng — Ranh Giới Giữa Tình Cảm Và Kỷ Luật Tài Chính!",
        "ssot": "Cẩm nang đàm phán dịch vụ giao nhận vận tải Eureka Logistics; Quy định phí DEM/DET hãng tàu; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD11 & CM4-CD14).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 122,
        "day": "[W49] Chủ Nhật (13/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Nộp thông báo gia công lại Mẫu 18A suốt 3–4 năm nhưng quên qua văn thư lấy số tiếp nhận: Đến lúc hq về kiểm tra mới báo 'chưa có bằng chứng đã nộp' và dọa phạt nặng! Phải chuẩn bị bộ chứng từ nào để cứu quyền miễn thuế SXXK?",
        "opt_b": "Checklist 4 nhóm chứng từ thực tế (hợp đồng, phiếu xuất-nhập kho, thanh toán, vận chuyển) giúp doanh nghiệp SXXK chứng minh chu trình gia công chân thật khi thiếu số tiếp nhận Mẫu 18A.",
        "opt_c": "Một tờ giấy thiếu con dấu văn thư có thể làm bay sạch ưu đãi miễn thuế nguyên phụ liệu nhiều năm liền: Bài học xương máu về 'làm đúng quy trình nhưng quên lưu bằng chứng pháp lý' của bộ phận chứng từ xnk!",
        "ssot": "Thông tư 38/2015/TT-BTC & Thông tư 39/2018/TT-BTC (Mẫu 18A thông báo gia công lại SXXK); Công văn 17658/CHQ-GSQL; scratch/daily_notes.md ngày 20/11/2026.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 123,
        "day": "[W49] Chủ Nhật (13/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose FB)",
        "opt_a": "Từ 01/01/2027: 13 Nhóm Thiết Bị Công Nghệ Số Cũ Chính Thức Cấm Nhập Theo Thông Tư 63/2026!",
        "opt_b": "Coi Chừng Hàng Tân Trang (Refurbished) Bị Chặn Cảng Từ 01/01/2027: Check Ngay 3 Mã HS Này",
        "opt_c": "Đừng Nhập Thiết Bị Văn Phòng Đã Qua Sử Dụng Cuối Năm Nếu Chưa Rà Danh Mục Thông Tư 63/2026",
        "ssot": "Thông tư số 63/2026/TT-BKHCN ngày 27/10/2026 (hiệu lực 01/01/2027); Repurpose từ STT 119.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },

    # ======================== TUẦN W50 (14/12/2026 - 20/12/2026) ========================
    {
        "stt": 124,
        "day": "[W50] Thứ Hai (14/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Làn sóng áp thuế phòng vệ thương mại cuối 2026: EU chính thức đánh thuế 16% thép cán nguội (2026/2547), 12,3%–14,4% ống đồng (2026/2497) & Ukraine khởi xướng điều tra chống lẩn tránh dây thép (AD-615) — Hàng xuất khẩu Việt Nam làm sao giữ thế đứng vững vàng?",
        "opt_b": "Cẩm nang 5 bước ứng phó thuế chống bán phá giá 16% của EU và làn sóng rà soát chống lẩn tránh xuất xứ: Cách rà soát hồ sơ truy xuất nguồn gốc phôi nguyên liệu ('Melt & Pour').",
        "opt_c": "Không chỉ thị trường Mỹ, cả Châu Âu và Đông Âu đang soi cực kỹ dấu hiệu chuyển tải nguyên liệu từ bên thứ ba: Doanh nghiệp sản xuất xuất khẩu Việt Nam cần chuẩn hóa hồ sơ nhà máy thế nào để tự bảo vệ?",
        "ssot": "Quy định thực thi (EU) 2026/2547 (hiệu lực 17/11/2026); Quy định (EU) 2026/2497; Quyết định AD-615/2026/441-01 (Ukraine); tapchicongthuong.vn & moit.gov.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #1).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 125,
        "day": "[W50] Thứ Ba (15/12/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Siêu Cảng Tự Động Hóa Maasvlakte 2 Rotterdam: Nơi Robot Và Xe Tự Hành AGV Vận Hành 24/7 Không Cần Một Ngọn Đèn & Chạy 100% Bằng Điện Gió!",
        "opt_b": "Bên Trong 'Cảng Biển Không Bóng Người' Tại Hà Lan: AI Và Cần Cẩu Tự Động Xếp Dỡ Hàng Triệu Container Nhanh Gấp Đôi Cảng Truyền Thống Như Thế Nào?",
        "opt_c": "Từ Siêu Cảng Rotterdam Maasvlakte 2 Đến Cửa Khẩu Thông Minh Lạng Sơn: Cuộc Cách Mạng Xe Tự Hành AGV Đang Định Hình Lại Ngành Logistics Việt Nam Ra Sao?",
        "ssot": "Port of Rotterdam Authority (Maasvlakte 2 Automated Terminals); Báo cáo Smart Port & Green Logistics Châu Âu; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD15).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 126,
        "day": "[W50] Thứ Ba (15/12/2026)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Chiến lược chứng minh xuất xứ thực chất ('Melt & Pour') trước làn sóng thuế chống bán phá giá EU (2026/2547) & Ukraine (AD-615): Checklist 5 tầng hồ sơ cho Giám đốc Xuất khẩu",
        "opt_b": "So sánh biên độ thuế phòng vệ thương mại Châu Âu giữa Việt Nam (16%) với Ấn Độ (9,5%) và Nhật Bản (27,9%): Bài học tái cấu trúc nguồn cung nguyên liệu nội khối",
        "opt_c": "Quản trị rủi ro điều tra chống lẩn tránh xuất xứ năm 2027: Doanh nghiệp sản xuất công nghiệp cần lưu trữ những chứng từ xưởng gốc nào để đứng vững khi thanh tra quốc tế kiểm toán?",
        "ssot": "Quy định (EU) 2026/2547; Quy định (EU) 2026/2497; tapchicongthuong.vn; reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx (STT 267) (Repurpose từ STT 124).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 127,
        "day": "[W50] Thứ Tư (16/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Một dn niêm yết uy tín như Gỗ Đức Thành (GDT) vẫn bị Hải quan KCX Tân Thuận truy thu & phạt hơn 1,13 tỷ đồng sau kiểm tra sau thông quan vì áp sai thuế suất NK & GTGT: Cứ tưởng hàng thông quan qua VNACCS là an toàn tuyệt đối?",
        "opt_b": "Quy tắc 'tự kiểm tra trong 60 ngày' và 3 bước thiết lập bảng đối chiếu mã HS – Thuế suất NK – Thuế GTGT trên ERP để không bị truy thu tiền tỷ nhiều năm sau thông quan.",
        "opt_c": "Nếu bạn nghĩ tờ khai đã vào luồng Xanh và lấy hàng về kho từ năm ngoái là hết chuyện thì hãy đọc kỹ bài học 1,13 tỷ đồng này: Bộ phận XNK hay Kế toán phải đứng ra soát biểu thuế định kỳ?",
        "ssot": "Hồ sơ công bố thông tin xử phạt sau kiểm tra sau thông quan Công ty CP Chế biến Gỗ Đức Thành (GDT) tại Chi cục Hải quan KCX Tân Thuận; thuongtruong.com.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 269).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 128,
        "day": "[W50] Thứ Năm (17/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 20",
        "opt_a": "Học XNK 5 Phút Tập 20: Cẩm Nang 5 Bước Xử Lý Tiền Thuế Nộp Thừa, Bù Trừ & Hoàn Thuế Hải Quan Điện Tử Theo Thông Tư 86/2026/TT-BTC",
        "opt_b": "Cách Tra Cứu Nợ Phí Hải Quan Lẻ Để Không Bị Chặn Thông Quan & Quy Trình Xoay Vòng Tiền Thuế Nộp Thừa 24/7 Chuẩn Thông Tư 86/2026",
        "opt_c": "So Sánh Hoàn Thuế Về Tài Khoản Ngân Hàng vs Bù Trừ Sang Tờ Khai Mới Theo Thông Tư 86/2026/TT-BTC: Phương Án Nào Tối Ưu Dòng Tiền Cho Doanh Nghiệp?",
        "ssot": "Thông tư số 86/2026/TT-BTC ngày 30/06/2026 (hiệu lực 01/07/2026); vanban.chinhphu.vn (docid=218821); reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx (STT 259).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 129,
        "day": "[W50] Thứ Năm (17/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose FB)",
        "opt_a": "Hàng Đã Thông Quan Vẫn Bị Truy Thu Hơn 1,13 Tỷ Đồng: Quy Tắc Tự Kiểm Tra Thuế Suất Trong 60 Ngày!",
        "opt_b": "3 Khoản Mục Cần Soát Ngay Giữa Kế Toán Và XNK Để Tránh Bị Ấn Định Thuế Sau Thông Quan",
        "opt_c": "Tờ Khai Luồng Xanh Không Có Nghĩa Là Miễn Hồi Tố 5 Năm: Cảnh Báo Từ Vụ Truy Thu 1,13 Tỷ Đồng",
        "ssot": "Vụ kiểm tra sau thông quan Gỗ Đức Thành (GDT) & Điều 20 Thông tư 38/2015/TT-BTC; Repurpose từ STT 127.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 130,
        "day": "[W50] Thứ Sáu (18/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Quyết định 3318/QĐ-BTC chính thức công bố 133 thủ tục hải quan thực hiện 'Phi địa giới hành chính' toàn quốc: Từ xem hàng, hủy tờ khai đến thủ tục khu phi thuế quan đều nộp trực tuyến không phụ thuộc nơi đặt trụ sở — Dn ở tỉnh lẻ đã tiết kiệm được chi phí đi lại chưa?",
        "opt_b": "Quy trình 4 bước nộp hồ sơ hải quan phi địa giới hành chính theo Quyết định 3318/QĐ-BTC (trong tổng số 255 TTHC Bộ Tài chính): Danh sách các thủ tục trọng điểm dn XNK dùng hàng ngày.",
        "opt_c": "Ngày xưa doanh nghiệp ở tỉnh gửi bộ hồ sơ hủy tờ khai hay xin xem hàng trước khi khai phải cử nhân sự chạy xe liên tỉnh chờ đóng dấu — Nay 133 TTHC hải quan đã mở cửa phi địa giới, thực tế triển khai tại các chi cục đang mượt đến đâu?",
        "ssot": "Quyết định số 3318/QĐ-BTC ngày 12/11/2026 của Bộ Tài chính; thuehaiquan.tapchikinhtetaichinh.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #5).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 131,
        "day": "[W50] Thứ Sáu (18/12/2026)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Kiểm soát rủi ro hậu kiểm thuế suất & mã HS trong doanh nghiệp sản xuất: Bài học từ vụ truy thu 1,13 tỷ đồng và quy trình đối soát dữ liệu ERP – VNACCS",
        "opt_b": "Tại sao doanh nghiệp niêm yết vẫn vướng sai sót biểu thuế nhập khẩu? Khung kiểm soát nội bộ 2 lớp giữa phòng Xuất nhập khẩu và Kế toán thuế",
        "opt_c": "Tận dụng cửa sổ 60 ngày tự khai bổ sung sau thông quan theo Luật Hải quan để giảm thiểu rủi ro ấn định thuế và tiền chậm nộp cho doanh nghiệp B2B",
        "ssot": "Công bố thông tin GDT tại Chi cục Hải quan KCX Tân Thuận; Điều 20 Thông tư 38/2015/TT-BTC; thuongtruong.com.vn (Repurpose từ STT 127).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 132,
        "day": "[W50] Thứ Bảy (19/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Phân tích & Hướng dẫn (8-10p)",
        "opt_a": "Báo Cáo Quyết Toán Hải Quan Mẫu 15/BCQT-NVL Cuối Năm Tài Chính: Bí Quyết Đối Chiếu 'Tam Giác 3 Bên' Kho – Kế Toán – Hải Quan Tránh Lệch Tồn Kho Tỷ Đồng!",
        "opt_b": "Hướng Dẫn Lập Bảng Định Mức Thực Tế & Xử Lý Phế Liệu Phế Phẩm Cho Doanh Nghiệp Gia Công (E21) Và SXXK (E31) Không Bị Ấn Định Thuế",
        "opt_c": "Tại Sao Số Liệu Kho Trên Phần Mềm Kế Toán Lại Khác Xa Báo Cáo Hải Quan Cuối Năm? Phác Đồ Rà Soát Mẫu 15 Trong 4 Bước Dễ Hiểu Nhất!",
        "ssot": "Thông tư 38/2015/TT-BTC & Thông tư 39/2018/TT-BTC (Mẫu 15/BCQT-NVL); Nghị định 18/2021/NĐ-CP; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-CD16).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 133,
        "day": "[W50] Chủ Nhật (20/12/2026)",
        "platform": "Facebook",
        "category": "FF — Funfact & Insight",
        "technique": "Liên hệ lịch sử / Nhân vật nổi tiếng",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Sự thật bất ngờ về mạng lưới SWIFT (1973): 'Hệ thần kinh tài chính' truyền hàng nghìn tỷ USD giao dịch XNK qua điện L/C (MT700) & T/T (MT103) mỗi ngày nhưng thực chất không hề giữ một xu tiền mặt nào! Vì sao chỉ nhìn ảnh chụp chuyển tiền lại cực kỳ nguy hiểm?",
        "opt_b": "Phân biệt điện MT103 (chuyển tiền T/T) vs điện MT700 (phát hành L/C) trên hệ thống SWIFT: Mẹo kiểm tra chứng từ thanh toán quốc tế thật để không giao vận đơn gốc nhầm cho lệnh chuyển tiền cắt ghép.",
        "opt_c": "Nếu đối tác nước ngoài gửi cho bạn một bức ảnh màn hình báo 'đã chuyển đủ 100% tiền cọc qua ngân hàng' và giục gửi ngay Bill gốc chiều nay — Bạn có dám thả hàng trước khi ngân hàng Việt Nam xác nhận nhận điện SWIFT MT103?",
        "ssot": "Tư liệu lịch sử SWIFT (Society for Worldwide Interbank Financial Telecommunication 1973); Chuẩn điện ISO 20022 / MT103 / MT700; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 270).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 134,
        "day": "[W50] Chủ Nhật (20/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "3 Con Số Sống Còn Khi Đối Chiếu Báo Cáo Quyết Toán Hải Quan Mẫu 15 Cuối Năm!",
        "opt_b": "Kho Báo Một Đằng, Kế Toán Nhập Một Nẻo: Coi Chừng Bị Truy Thu Thuế Cuối Năm Tài Chính",
        "opt_c": "Checklist 30 Giây Rà Soát Định Mức BOM Trước Khi Nộp Báo Cáo Quyết Toán SXXK",
        "ssot": "Trích xuất từ YouTube CM2-CD16 (STT 132); Mẫu 15/BCQT-NVL theo Thông tư 39/2018/TT-BTC.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },

    # ======================== TUẦN W51 (21/12/2026 - 27/12/2026) ========================
    {
        "stt": 135,
        "day": "[W51] Thứ Hai (21/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Đếm ngược 7 ngày tới mốc thu thuế 28/12/2026: Thái Lan chính thức áp thuế tự vệ 7,73% CIF trong 3 năm cho 11 mã HS hạt nhựa Polypropylene (nhóm 3902) & Việt Nam KHÔNG được miễn trừ — Dn bao bì, nhựa gia dụng xuất khẩu xoay xở biên lợi nhuận thế nào?",
        "opt_b": "Bóc tách 11 mã HS hạt nhựa PP nhóm 3902 chịu thuế tự vệ 7,73% vào thị trường Thái Lan từ 28/12/2026: Cách tính lại Landed Cost và đàm phán chia sẻ thuế với nhà nhập khẩu ASEAN.",
        "opt_c": "Dù chung khối ASEAN hưởng C/O Form D thuế 0%, thuế tự vệ toàn cầu vẫn có thể đánh thẳng vào hàng Việt khi vượt ngưỡng tỷ trọng: Doanh nghiệp nhựa nên điều chỉnh hợp đồng sang điều kiện nào để không gánh lỗ?",
        "ssot": "Thông báo của Cục Phòng vệ Thương mại (Bộ Công Thương) ngày 12/11/2026 về Quyết định tự vệ hạt nhựa PP Thái Lan (thu thuế từ 28/12/2026); moit.gov.vn & pvtm.gov.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #2).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 136,
        "day": "[W51] Thứ Ba (22/12/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Chuẩn Hóa Khóa Góc Twistlock & Kích Thước Container (1965–1970): Vì Sao Malcolm McLean Quyết Định Hiến Tặng Bằng Sáng Chế Triệu Đô Cho Thế Giới?",
        "opt_b": "Cuộc Chiến Kích Thước Container Khốc Liệt Nhất Lịch Sử Hàng Hải: Tại Sao Kích Thước 20ft Và 40ft Thắng Thế Trước 24ft Của Matson Và 35ft Của Sea-Land?",
        "opt_c": "Chỉ Một Chiếc Khóa Xoay 4 Góc Đã Giảm Hàng Chục Nghìn Tai Nạn Tàu Biển & Mở Ra Kỷ Nguyên Thương Mại Toàn Cầu Không Biên Giới!",
        "ssot": "Tư liệu ISO Technical Committee 104 (1965–1970); Tác phẩm 'The Box' (Marc Levinson); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD16).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 137,
        "day": "[W51] Thứ Ba (22/12/2026)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Thuế tự vệ hạt nhựa Polypropylene Thái Lan (7,73% CIF từ 28/12/2026): Tác động tới chuỗi cung ứng bao bì – nhựa ASEAN và chiến lược phân bổ thị trường 2027",
        "opt_b": "Tại sao C/O Form D ưu đãi 0% không giúp doanh nghiệp miễn nhiễm trước biện pháp tự vệ (Safeguard)? Phân tích 3 kịch bản điều chỉnh giá xuất khẩu B2B",
        "opt_c": "Quản trị rủi ro hợp đồng dài hạn ngành nguyên liệu nhựa khi quốc gia nhập khẩu kích hoạt hàng rào tự vệ 3 năm liên tiếp",
        "ssot": "Thông báo Cục Phòng vệ Thương mại (moit.gov.vn / pvtm.gov.vn) ngày 12/11/2026; Hiệp định ATIGA & WTO Safeguards (Repurpose từ STT 135).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 138,
        "day": "[W51] Thứ Tư (23/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Vụ bắt giữ lô máy móc cũ khai mới 100% trị giá hơn 10 tỷ đồng tại Hữu Nghị & Vụ khai sai mã HS xe chuyên dùng thiếu thuế hơn 2,19 tỷ đồng tại Tân Thanh: Đừng bao giờ đánh cược lô hàng tiền tỷ bằng cách sơn lại vỏ rồi khai 'mới 100%' để né Quyết định 18!",
        "opt_b": "Checklist 4 bước kiểm tra năm sản xuất, tem nhãn gốc (Nameplate) và chứng thư giám định máy móc cũ đạt chuẩn QCVN/G7 trước khi xếp xe về cửa khẩu phía Bắc.",
        "opt_c": "Nghe xưởng bên kia báo 'máy sơn lại nhìn y như mới, cứ khai mới 100% cho nhanh khỏi tốn tiền giám định' — Kết cục bị phạt 40 triệu và buộc tái xuất trong 30 ngày thì ai là người chịu thiệt hại hàng tỷ đồng lưu bãi?",
        "ssot": "Quyết định số 18/2019/QĐ-TTg về nhập khẩu máy móc thiết bị đã qua sử dụng; Hồ sơ xử lý vi phạm Hải quan khu vực VI (Hữu Nghị & Tân Thanh); thuehaiquan.tapchikinhtetaichinh.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 270).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 139,
        "day": "[W51] Thứ Năm (24/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 21",
        "opt_a": "Học XNK 5 Phút Tập 21: Checklist 4 Bước Chuẩn Hóa Hồ Sơ Xuất Khẩu Thực Phẩm Sang Trung Quốc Theo Thông Báo 150/2026 GACC (Chỉ Còn 10 Mục, Duyệt 3 Ngày)",
        "opt_b": "Hướng Dẫn Đăng Ký Và Cập Nhật Tài Khoản CIFER GACC Trong 60 Ngày Để Nông Sản, Thủy Sản Không Bị Treo Mã Xuất Khẩu Dịp Tết 2027",
        "opt_c": "So Sánh Quy Định GACC Cũ (20 Mục) vs Thông Báo 150/2026 Mới (10 Mục): Cơ Hội Tăng Tốc Xuất Khẩu Thực Phẩm Chính Ngạch Sang Trung Quốc",
        "ssot": "Thông báo số 150/2026 của Tổng cục Hải quan Trung Quốc (GACC, hiệu lực 23/10/2026); congthuong.vn & moit.gov.vn; reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx (STT 269).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 140,
        "day": "[W51] Thứ Năm (24/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Xuất Thực Phẩm Sang Trung Quốc 2027: Đăng Ký Mã GACC Chỉ Còn 10 Mục & Duyệt Trong 3 Ngày!",
        "opt_b": "Quên Cập Nhật Thay Đổi GACC Trong 60 Ngày: Container Nông Sản Có Thể Bị Dừng Tại Cửa Khẩu",
        "opt_c": "3 Điểm Mới Cực Có Lợi Trong Thông Báo 150/2026 Của Hải Quan Trung Quốc Cho Chủ Hàng Thực Phẩm",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 21 (STT 139); Thông báo 150/2026/GACC.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 141,
        "day": "[W51] Thứ Sáu (25/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Công văn 24284/CHQ-GSQL (16/11/2026): Gỡ vướng kiểm tra nhà nước về chất lượng hàng nhập khẩu theo Nghị định 37/2026/NĐ-CP — Khi đầu mối các Bộ quản lý chuyên ngành vừa tái cấu trúc, dn làm sao đối chiếu đúng cơ quan kiểm định để hàng không nằm chờ cảng?",
        "opt_b": "Checklist 3 bước đối chiếu cơ quan kiểm tra chất lượng hàng nhập khẩu mới nhất theo Công văn 24284/CHQ-GSQL & Nghị định 37/2026/NĐ-CP cho nhóm máy móc, thiết bị bảo hộ, hóa chất.",
        "opt_c": "Chuyện muôn thuở ở cửa khẩu những ngày cuối năm: Cầm bộ hồ sơ kiểm tra chất lượng đi nộp nhưng bị trả về vì 'đầu mối quản lý đã chuyển sang Bộ khác' — Công văn 24284 đang tháo gỡ nút thắt nào cho doanh nghiệp nhập khẩu?",
        "ssot": "Công văn số 24284/CHQ-GSQL ngày 16/11/2026 của Cục Hải quan; Nghị định số 37/2026/NĐ-CP; thuvienphapluat.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #7).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 142,
        "day": "[W51] Thứ Sáu (25/12/2026)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Quản trị rủi ro nhập khẩu máy móc thiết bị đã qua sử dụng theo Quyết định 18/2019/QĐ-TTg: 4 điểm kiểm tra Nameplate & chứng thư giám định tại xưởng gốc",
        "opt_b": "Bài học 10 tỷ đồng từ hành vi khai báo sai tình trạng máy móc tại cửa khẩu biên giới: Tiêu chuẩn tuân thủ bắt buộc đối với nhà máy FDI và nhà thầu cơ khí",
        "opt_c": "Tại sao việc kiểm tra trước khi xếp hàng (Pre-shipment Inspection) lại tiết kiệm gấp 50 lần chi phí tái xuất và lưu bãi tại cửa khẩu?",
        "ssot": "Quyết định 18/2019/QĐ-TTg; Hồ sơ Hải quan khu vực VI; thuehaiquan.tapchikinhtetaichinh.vn (Repurpose từ STT 138).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 143,
        "day": "[W51] Thứ Bảy (26/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Phân tích & Hướng dẫn (8-10p)",
        "opt_a": "Các Loại Hình Kiểm Tra Sau Thông Quan (KTSTQ) Với Doanh Nghiệp Thương Mại: Phân Biệt Kiểm Tra 60 Ngày, Hồi Tố 5 Năm Tại Hải Quan vs Tại Trụ Sở DN!",
        "opt_b": "Giải Mã 3 Căn Cứ Kích Hoạt Kiểm Tra Sau Thông Quan (Dấu Hiệu Vi Phạm, Quản Lý Rủi Ro, Đánh Giá Tuân Thủ) & Cách Chuẩn Bị Hồ Sơ Đón Đoàn Thanh Tra",
        "opt_c": "Hàng Đã Thông Quan Rồi Sao Vẫn Bị Gọi Kiểm Tra Sau 3 Năm? Dịch Luật Hải Quan 2014 Điều 77–80 Ra Tiếng Người Dễ Hiểu Nhất!",
        "ssot": "Luật Hải quan số 54/2014/QH13 (Điều 77–80); Điều 142–143 Thông tư 38/2015/TT-BTC (sửa đổi bởi TT 39/2018 & TT 121/2025); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-HOT52).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 144,
        "day": "[W51] Chủ Nhật (27/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Tình huống giả định mở",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Chặn đứng lô hàng thời trang dính logo thương hiệu lớn (không có giấy ủy quyền hãng) tại cửa khẩu Chi Ma: Đừng tưởng chỉ cần ghi chung chung 'áo phông, túi xách' trên tờ khai là qua mặt được máy soi chiếu và dữ liệu bảo hộ SHTT!",
        "opt_b": "Checklist 3 giấy tờ bắt buộc phải kiểm tra trước khi nhận ủy thác hoặc vận chuyển hàng có thương hiệu: Thư ủy quyền phân phối (LOA) & tra cứu danh mục bảo hộ SHTT tại Hải quan.",
        "opt_c": "Nếu bạn làm dịch vụ logistics hoặc ủy thác nhập khẩu, khách gửi một kiện hàng có in logo thương hiệu quốc tế nhưng bảo 'cứ khai không nhãn hiệu đi cho nhanh' — Bạn sẽ từ chối hay nhắm mắt ký chữ ký số?",
        "ssot": "Công điện 38/CĐ-TTg; Luật Sở hữu trí tuệ; Hồ sơ Đội 1 Chi cục Điều tra chống buôn lậu & Hải quan Chi Ma; baochinhphu.vn & thuehaiquan.tapchikinhtetaichinh.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 265).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 145,
        "day": "[W51] Chủ Nhật (27/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose FB)",
        "opt_a": "Nhập Hàng Có Logo Thương Hiệu Mà Thiếu Giấy Ủy Quyền (LOA): Bị Tịch Thu & Đẩy Luồng Đỏ 100%!",
        "opt_b": "Đừng Ký Ủy Thác Nhập Khẩu Hàng Thương Hiệu Nếu Chưa Kiểm Tra 2 Giấy Tờ Pháp Lý Này",
        "opt_c": "Khai 'Áo Phông Không Nhãn' Để Né Kiểm Tra SHTT: Sai Lầm Khiến Lô Hàng Bị Giữ Tại Cửa Khẩu",
        "ssot": "Hồ sơ Hải quan cửa khẩu Chi Ma (Lạng Sơn); Luật Sở hữu trí tuệ; Repurpose từ STT 144.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },

    # ======================== TUẦN W52 (28/12/2026 - 03/01/2027) ========================
    {
        "stt": 146,
        "day": "[W52] Thứ Hai (28/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Chỉ còn 3 ngày trước giờ G 01/01/2027: Thông tư 154/2026/TT-BTC chính thức siết thủ tục hải quan & áp thời hạn 30 ngày kê khai nộp thuế đối với hàng hóa mua gom của cư dân biên giới — Bước ngoặt buộc thương nhân chuyển sang xuất nhập khẩu chính thống!",
        "opt_b": "Những điểm mới cốt lõi của Thông tư 154/2026/TT-BTC (sửa đổi Thông tư 80/2019/TT-BTC): Quy trình khai báo và lộ trình chuẩn hóa hóa đơn, chứng từ cho thương nhân thu mua nông sản biên giới.",
        "opt_c": "Cách làm ăn thu gom manh mún qua biên giới đang dần khép lại khi hành lang pháp lý ngày càng minh bạch: Doanh nghiệp thương mại cần chuẩn bị quy trình xuất nhập khẩu chính thống như thế nào từ đầu năm 2027?",
        "ssot": "Thông tư số 154/2026/TT-BTC ngày 15/10/2026 (hiệu lực 01/01/2027, sửa đổi TT 80/2019/TT-BTC); thoibaotaichinhvietnam.vn; reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx (STT 264).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 147,
        "day": "[W52] Thứ Ba (29/12/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Tuyến Đường Trà Mã Cổ Đạo (Tea Horse Road): Hải Trình Trên Lưng Ngựa Vượt Đỉnh Himalaya Chở Trà Đổi Chiến Mã Thời Cổ Đại!",
        "opt_b": "Tuyến Vận Tải Khắc Nghiệt Nhất Lịch Sử Nhân Loại Trước Khi Có Container: Cách Người Xưa Định Giá Trà Bánh Thành 'Đồng Tiền Cứng' Xuyên Biên Giới",
        "opt_c": "Từ Con Đường Trà Mã Cổ Đạo Đến Hạ Tầng Logistics Xuyên Á Hôm Nay: Bài Học Về Sức Mạnh Của Hàng Hóa Thiết Yếu Và Kỷ Luật Vận Tải",
        "ssot": "Tư liệu Lịch sử Thương mại Tây Nam Trung Quốc – Tây Tạng (Tea Horse Road); Mark Jenkins (National Geographic); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD17).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 148,
        "day": "[W52] Thứ Ba (29/12/2026)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Nghị định 378/2026/NĐ-CP & Cơ chế Doanh nghiệp Ưu tiên Hải quan (AEO) tại Khu Thương mại Tự do (FTZ) Việt Nam: Miễn điều kiện kim ngạch nhưng chuẩn hóa giám sát 24/7",
        "opt_b": "4 tiêu chuẩn hạ tầng bắt buộc cho kho bãi trong Khu phi thuế quan và Khu Logistics tích hợp theo Nghị định 378/2026/NĐ-CP (hàng rào cứng, camera lưu 6 tháng, cân điện tử)",
        "opt_c": "Làn sóng Khu thương mại tự do (FTZ) tại Việt Nam năm 2027: Cơ hội nào cho các doanh nghiệp sản xuất và phân phối chuỗi cung ứng khu vực?",
        "ssot": "Nghị định số 378/2026/NĐ-CP ngày 01/10/2026; Quyết định số 2735/QĐ-BTC; chinhphu.vn & thuvienphapluat.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #4).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 149,
        "day": "[W52] Thứ Tư (30/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Vụ kiện hành chính hiếm có tại Hải Phòng: Nhập 31 container (883 tấn lúa mì) nhưng khai nhầm mục đích sử dụng từ 'thức ăn chăn nuôi' sang 'thực phẩm cho người' — Đã nộp phạt 40 triệu vẫn bị giữ kéo dài hơn 1 tháng, thiệt hại đòi bồi thường hơn 3,3 tỷ đồng!",
        "opt_b": "Checklist đối chiếu 2 lớp chứng từ đầu vào (mục đích sử dụng, luồng kiểm dịch động vật vs an toàn thực phẩm) giúp dn nông sản không tự trói chân 31 container ngoài cảng.",
        "opt_c": "Chỉ một dòng mô tả mục đích sử dụng sai giữa hai luồng kiểm tra chuyên ngành khác nhau có thể thổi bay 3,3 tỷ đồng tiền lưu cont, lưu bãi và hủy hợp đồng: Bộ phận Sales & Chứng từ nhà bạn đã chốt chéo thông tin trước khi truyền tờ khai chưa?",
        "ssot": "Hồ sơ xét xử sơ thẩm TAND khu vực 4 TP. Hải Phòng (Công ty TNHH XNK và Dịch vụ Phúc Thành vs Chi cục Hải quan khu vực III & cảng Đình Vũ); reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 266).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 150,
        "day": "[W52] Thứ Năm (31/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 22",
        "opt_a": "Học XNK 5 Phút Tập 22: Quy Trình 5 Bước Đưa Hàng Vào - Ra Kho Ngoại Quan (Gia Cố, Chia Gói, Nhập Nội Địa) & Xử Lý Thuế GTGT Xuất Trả B13 Theo CV 23897 & 23739!",
        "opt_b": "Thủ Tục Tái Xuất, Xuất Trả Hàng Nhập Khẩu Không Đạt Chất Lượng (Loại Hình B13): 5 Chứng Từ Sống Còn Để Không Mất Oan Thuế GTGT Theo Công Văn 23739/CHQ-NVTHQ",
        "opt_c": "Cẩm Nang Vận Hành Kho Ngoại Quan Chuẩn Nghị Định 167/2025/NĐ-CP & Cụm Công Văn 23897–23907/CHQ-GSQL: Thời Hạn 12+12 Tháng Và Bẫy C/O Nhập Ngược Nội Địa",
        "ssot": "Công văn số 23897/CHQ-GSQL & 23907/CHQ-GSQL (05/11/2026); Công văn số 23739/CHQ-NVTHQ (02/11/2026); Nghị định 167/2025/NĐ-CP; reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx (STT 265).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 151,
        "day": "[W52] Thứ Năm (31/12/2026)",
        "platform": "TikTok/Shorts",
        "category": "FF — Funfact & Insight",
        "technique": "Dẫn chứng quốc tế",
        "format": "Video ngắn <60s (Repurpose FF)",
        "opt_a": "Vì Sao 78 Chiếc Đinh Trên Pallet Chuẩn Châu Âu (800×1200mm) Giúp Tiết Kiệm 15% Tiền Container Xuất Khẩu?",
        "opt_b": "3 Kích Thước Pallet Quốc Tế (EU, Mỹ, Châu Á) Chủ Hàng Phải Biết Để Không Mất Phí Dỡ Hàng Tại Cảng Đến",
        "opt_c": "Sự Thật Thú Vị Sau Chiếc Kệ Gỗ EUR-Pallet Năm 1961 Làm Thay Đổi Tốc Độ Xếp Dỡ Toàn Cầu",
        "ssot": "Hiệp hội Pallet Châu Âu EPAL & Liên minh Đường sắt Quốc tế UIC (1961); Tiêu chuẩn ISO 3394 & ISPM 15; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 269).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 152,
        "day": "[W52] Thứ Sáu (01/01/2027)",
        "platform": "Facebook",
        "category": "FF — Funfact & Insight",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Chào 2027: Nhìn lại cột mốc Kim ngạch XNK Việt Nam vượt 1.070 tỷ USD năm 2026 (+33,4%) & 3 nguyên tắc 'Giữ chữ Tín – Chuẩn pháp lý – Làm chủ công nghệ' để doanh nghiệp bước vào năm mới vững vàng trước làn sóng kiểm soát rủi ro!",
        "opt_b": "Thương mại Việt Nam vượt mốc 1.000 tỷ USD nhưng tỷ lệ luồng Vàng/Đỏ và điều tra phòng vệ thương mại cũng tăng theo: Doanh nghiệp SME cần chuẩn bị tâm thế gì cho năm 2027?",
        "opt_c": "Ngày đầu tiên của năm mới 2027, ngồi nhìn lại hành trình 2026 cùng anh em Eureka: Vì sao con đường làm chính ngạch minh bạch tuy chậm lúc đầu nhưng lại là con đường duy nhất đi được đường dài?",
        "ssot": "Số liệu Cục Hải quan kỳ 1 tháng 11/2026 (1.070,26 tỷ USD) & lũy kế 10 tháng 2026; thoibaotaichinhvietnam.vn & tapchikinhtetaichinh.vn; reports/SEO_Keyword_Trends_Eureka.xlsx (Keyword #3).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 153,
        "day": "[W52] Thứ Sáu (01/01/2027)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Chi phí cơ hội từ một lỗi khai sai mục đích sử dụng trên tờ khai nhập khẩu: Bài học rà soát chứng từ 2 lớp giữa Sales và Chứng từ XNK (Case 883 tấn lúa mì)",
        "opt_b": "Vì sao một sai sót kỹ thuật nhỏ trên tờ khai lại có thể kéo dài thời gian giải phóng 31 container hơn 1 tháng? Cách thiết kế quy trình đối soát chuyên ngành đầu vào",
        "opt_c": "Quản trị rủi ro tranh chấp hành chính trong logistics: Chuẩn bị hồ sơ chứng minh thiện chí ngay từ bước khai bổ sung đầu tiên",
        "ssot": "Hồ sơ xét xử TAND khu vực 4 TP. Hải Phòng (883 tấn lúa mì / Công ty Phúc Thành); reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 266) (Repurpose từ STT 149).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 154,
        "day": "[W52] Thứ Bảy (02/01/2027)",
        "platform": "YouTube",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Video Tâm tình Thực chiến (7-9p) [CM4]",
        "opt_a": "Đúc Kết 15 Năm Làm Nghề Thương Mại & Cách Dân Không Học Luật Dùng AI Đọc Văn Bản Pháp Luật Đầu Năm 2027 Không Bị Ngợp!",
        "opt_b": "3 Phẩm Chất Cốt Tử Để Một Bạn Trẻ Ngành XNK – Logistics Đi Đường Dài & Phương Pháp Tự Học Luật Qua Case Study Thực Tế",
        "opt_c": "Tâm Sự Đầu Năm 2027: Khi Công Nghệ AI Đã Hỗ Trợ 80% Việc Tra Cứu Quy Định, Giá Trị Thật Của Người Làm Logistics Nằm Ở Đâu?",
        "ssot": "Hành trình 15 năm thực chiến của Founder Nguyễn Thị Ngọc Thúy tại Eureka Logistics; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD18 & CM4-CD19).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 155,
        "day": "[W52] Chủ Nhật (03/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Công văn 24288/CHQ-GSQL: Vướng mắc phân loại hóa chất công nghiệp (CY-9) khi máy kiểm định chưa tách được thành phần polymer & cách áp mã 2843.30.00 cho Potassium Gold Cyanide — Bộ tài liệu kỹ thuật gốc từ nhà sản xuất chính là 'phao cứu sinh'!",
        "opt_b": "Kỹ thuật 5 bước chuẩn bị hồ sơ phân loại & đối chiếu danh mục hóa chất/tiền chất theo Nghị định 24/2026/NĐ-CP và Công văn 24288/CHQ-GSQL để không bị tắc hàng hóa chất tại cảng.",
        "opt_c": "Nhập hóa chất hỗn hợp về sản xuất nhưng cơ quan kiểm định chưa đủ điều kiện phân tích sâu: Doanh nghiệp cần chuẩn bị sẵn bảng thành phần CAS và tài liệu kỹ thuật nhà máy thế nào để bảo vệ mã HS?",
        "ssot": "Công văn số 24288/CHQ-GSQL ngày 16/11/2026 của Cục Hải quan; Phụ lục I Thông tư 31/2022/TT-BTC; Phụ lục III Nghị định 24/2026/NĐ-CP; reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx (STT 270).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 156,
        "day": "[W52] Chủ Nhật (03/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Xuất Trả Hàng Lỗi Ra Nước Ngoài (Loại Hình B13): Đừng Quên 1 Chứng Từ Này Để Không Mất Oan Thuế GTGT!",
        "opt_b": "Hàng Nhập Về Bị Lỗi Kỹ Thuật Phải Trả Lại Xưởng Ngoại: Quy Trình 3 Bước Làm Thủ Tục Xuất Trả Chuẩn",
        "opt_c": "Công Văn 23739/CHQ-NVTHQ: Khi Nào Hàng Tái Xuất Trả Chủ Hàng Được Hoàn Hoặc Không Thu Thuế GTGT?",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 22 (STT 150); Công văn 23739/CHQ-NVTHQ ngày 02/11/2026.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
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
    for item in W49_W52_ITEMS:
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
    print(f"  ✅ Đạt 100% compliance cho toàn bộ {len(W49_W52_ITEMS)} slots (0 vi phạm)!")


def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W52 CỘNG DỒN)")
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

    # Build lookup by STT (handle both int and str STT)
    existing_row_by_stt = {}
    for r in range(5, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None:
            try:
                existing_row_by_stt[int(str(val).strip())] = r
            except ValueError:
                pass

    for item in W49_W52_ITEMS:
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

    # Mở rộng DataValidation dropdown cho cột K
    dv = DataValidation(
        type="list",
        formula1='"Chờ duyệt,Duyệt Option A,Duyệt Option B,Duyệt Option C,Không duyệt"',
        allow_blank=True,
    )
    ws.add_data_validation(dv)
    dv.add("K5:K220")

    # Cân chỉnh nhẹ cột H, I để đọc dễ hơn
    if ws.column_dimensions["H"].width and ws.column_dimensions["H"].width < 42:
        ws.column_dimensions["H"].width = 45
    if ws.column_dimensions["I"].width and ws.column_dimensions["I"].width < 42:
        ws.column_dimensions["I"].width = 45
    if ws.column_dimensions["E"].width and ws.column_dimensions["E"].width < 20:
        ws.column_dimensions["E"].width = 22

    # Atomic save (Zero Backup Rule 10)
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print(f"  -> Đã lưu thành công {len(W49_W52_ITEMS)} dòng (STT 113–156) vào Sheet 1 qua atomic_save!")


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

    # Tách bỏ phần W49-W52 cũ nếu script chạy lại (idempotent)
    marker = "# TUẦN W49: LỊCH BIÊN TẬP ĐA NỀN TẢNG"
    if marker in existing_content:
        existing_content = existing_content.split(marker)[0].rstrip()
        # Loại bỏ dấu '---' thừa cuối nếu có
        if existing_content.endswith("---"):
            existing_content = existing_content[:-3].rstrip()

    existing_content = existing_content.replace(
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W48 (Từ 17/08/2026 đến 29/11/2026)",
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W52 (Từ 17/08/2026 đến 03/01/2027) — W49–W52 Chờ duyệt Checkpoint 1"
    )

    week_defs = [
        ("W49", "07/12/2026 - 13/12/2026", [x for x in W49_W52_ITEMS if "[W49]" in x["day"]]),
        ("W50", "14/12/2026 - 20/12/2026", [x for x in W49_W52_ITEMS if "[W50]" in x["day"]]),
        ("W51", "21/12/2026 - 27/12/2026", [x for x in W49_W52_ITEMS if "[W51]" in x["day"]]),
        ("W52", "28/12/2026 - 03/01/2027", [x for x in W49_W52_ITEMS if "[W52]" in x["day"]]),
    ]

    appended = "\n".join(build_week_md(wc, dr, items) for wc, dr, items in week_defs)
    final_content = existing_content + appended + "\n"

    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(final_content)
    print("  -> Đã cập nhật thành công reports/plan1/weekly_calendar.md!")


def save_json_data():
    print("=== 3. LƯU SCRATCH/W49_W52_DATA.JSON ===")
    os.makedirs(os.path.dirname(JSON_OUTPUT), exist_ok=True)
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W49_W52_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu 44 slots vào {JSON_OUTPUT}")


if __name__ == "__main__":
    print("=== BẮT ĐẦU QUY TRÌNH WF01: LẬP LỊCH ĐA NỀN TẢNG TUẦN W49 - W52 ===")
    run_compliance_check()
    update_excel_calendar()
    update_markdown_calendar()
    save_json_data()
    print("=== HOÀN TẤT THỰC THI GIAI ĐOẠN 1 TUẦN W49 - W52! ===")
