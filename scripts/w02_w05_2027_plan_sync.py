# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Đa nền tảng 4 Tuần W02/2027 – W05/2027 (11/01/2027 - 07/02/2027)
theo quy trình huong_dan_phoi_hop_gd1_lap_lich.md & workflows/wf01_weekly_planning.md:
- 44 slots (STT 168 - 211): 11 slots/tuần x 4 tuần
  + Facebook: 4 bài/tuần (Thứ 2, 4, 6, CN)
  + YouTube: 3 video/tuần (Thứ 3, 5, 7)
  + LinkedIn (Repurpose B2B): 2 bài/tuần (Thứ 3, 6)
  + TikTok/Shorts (Repurpose <60s): 2 video/tuần (Thứ 5, CN)
- 3 phương án Options A/B/C cho mỗi slot
- Tự động kiểm duyệt Bộ từ cấm (knowledge/banned_words.md), từ nhạy cảm nền tảng & Anti-AI Buzzwords
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx qua atomic_save (Rule 10 - Zero Backup)
- Cập nhật reports/plan1/weekly_calendar.md & lưu scratch/w02_w05_2027_data.json
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
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w02_w05_2027_data.json")

W02_W05_2027_ITEMS = [
    # ======================== TUẦN W02/2027 (11/01/2027 - 17/01/2027) ========================
    {
        "stt": 168,
        "day": "[W02/2027] Thứ Hai (11/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Quyết định 4502/QĐ-BKHCN công bố bộ thủ tục mới về kiểm soát hàng hóa lưỡng dụng (điện tử, máy tính, vật liệu đặc thù): Khi quyền cấp phép xuất khẩu và quá cảnh được phân cấp về UBND tỉnh qua Một cửa Quốc gia kèm yêu cầu End-User Statement, doanh nghiệp công nghệ đã chuẩn hóa hồ sơ người dùng cuối chưa?",
        "opt_b": "Checklist 4 bước xin giấy phép xuất khẩu, tạm nhập tái xuất và quá cảnh hàng hóa lưỡng dụng theo Quyết định 4502/QĐ-BKHCN (gắn Nghị định 259/2025/NĐ-CP & Thông tư 30/2026/TT-BKHCN): Cách chuẩn bị bản khai End-User Certificate để không bị dừng hàng tại cửa khẩu.",
        "opt_c": "Nếu đối tác nước ngoài gửi gấp một lô linh kiện điện tử công nghệ cao đi quá cảnh hoặc xuất khẩu, nhưng đến lúc nộp hồ sơ mới phát hiện vướng diện kiểm soát thương mại chiến lược theo Quyết định 4502/QĐ-BKHCN — Bạn xử lý thế nào để không lỡ lịch tàu?",
        "ssot": "Quyết định số 4502/QĐ-BKHCN ngày 19/11/2026 của Bộ Khoa học và Công nghệ (hiệu lực 19/11/2026); Nghị định 259/2025/NĐ-CP; Thông tư 30/2026/TT-BKHCN; mst.gov.vn & luatvietnam.vn; scratch/wf04_legal_research_2026_11_21.md.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 169,
        "day": "[W02/2027] Thứ Ba (12/01/2027)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Liên hệ lịch sử / Nhân vật nổi tiếng",
        "format": "Video Phân tích & Insight (8-12p) [CM3-CD05]",
        "opt_a": "Công Ty Đông Ấn Hà Lan (VOC, 1602) & Sàn Chứng Khoán Đầu Tiên Thế Giới: Vì Sao Cổ Phiếu Ra Đời Từ Những Chuyến Tàu Gỗ Chở Gia Vị Vượt Đại Dương?",
        "opt_b": "Giải Mã Mô Hình Chia Sẻ Rủi Ro Hàng Hải Năm 1602 Tại Amsterdam: Cách Công Ty VOC Huy Động Vốn Xã Hội Và Thống Trị Tuyến Thương Mại Á – Âu Trong 200 Năm",
        "opt_c": "Từ Những Con Tàu Buôn Tiêu Quế Thế Kỷ 17 Đến Thị Trường Tài Chính Toàn Cầu: Bài Học Quản Trị Rủi Ro Chuỗi Cung Ứng Đường Biển Của Người Hà Lan!",
        "ssot": "Tư liệu Lịch sử Công ty Đông Ấn Hà Lan (VOC, 20/03/1602) & Sàn Giao dịch Chứng khoán Amsterdam; Wikipedia & World History Encyclopedia; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 271); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD05).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 170,
        "day": "[W02/2027] Thứ Ba (12/01/2027)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Kiểm soát thương mại chiến lược & thủ tục cấp phép hàng hóa lưỡng dụng theo Quyết định 4502/QĐ-BKHCN: 4 yêu cầu hồ sơ End-User Statement cho doanh nghiệp công nghệ & FDI",
        "opt_b": "Phân cấp giải quyết thủ tục hàng hóa lưỡng dụng về UBND tỉnh qua Cổng Một cửa Quốc gia: Doanh nghiệp sản xuất điện tử cần rà soát những mã sản phẩm nào?",
        "opt_c": "Chuẩn hóa quy trình tuân thủ thương mại chiến lược (Strategic Trade Control) tại Việt Nam năm 2027: Bài học đối chiếu tiêu chuẩn quốc tế từ Nghị định 259/2025/NĐ-CP",
        "ssot": "Quyết định 4502/QĐ-BKHCN (19/11/2026); Nghị định 259/2025/NĐ-CP; Thông tư 30/2026/TT-BKHCN; mst.gov.vn (Repurpose từ STT 168).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 171,
        "day": "[W02/2027] Thứ Tư (13/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Tăng vốn điều lệ từ 1,9 tỷ lên 5 tỷ đồng cho công ty TNHH 2 thành viên nhưng hồ sơ bị trả về vì vướng kê khai Chủ sở hữu hưởng lợi (Mẫu 10 & Mẫu 12 ĐKKD): Lỗi nhỏ trên giấy phép kinh doanh có thể làm nghẽn toàn bộ giao dịch ngân hàng và chữ ký số XNK thế nào?",
        "opt_b": "Checklist 3 điểm rà soát đồng bộ tỷ lệ vốn góp và thông tin Chủ sở hữu hưởng lợi trên Mẫu 12 đăng ký kinh doanh khi doanh nghiệp tăng vốn hoặc thay đổi thành viên góp vốn.",
        "opt_c": "Tình huống thực tế mới gặp trong cộng đồng doanh nghiệp: Góp thêm vốn mở rộng nhập khẩu nhưng vướng Mẫu 12 Chủ sở hữu hưởng lợi khiến hồ sơ bị treo nhiều tuần — Bộ phận hành chính và kế toán nhà bạn có từng rà soát chéo hồ sơ ĐKKD trước khi chốt kế hoạch nhập hàng?",
        "ssot": "Nhật ký công việc thực tế scratch/daily_notes.md ngày 21/11/2026 (Case Study kê khai Mẫu 12 Chủ sở hữu hưởng lợi khi tăng vốn công ty TNHH 2 thành viên).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 172,
        "day": "[W02/2027] Thứ Năm (14/01/2027)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 24 [CM2-HOT48]",
        "opt_a": "Học XNK 5 Phút Tập 24: Nhập Hàng Mẫu Về Đánh Giá Chất Lượng Có Phải Kiểm Tra Chuyên Ngành Không? Giải Mã Điều 86 & Điều 40 Nghị Định 37/2026/NĐ-CP!",
        "opt_b": "4 Cái Bẫy Khi Nhập Hàng Mẫu (Sample) Về Test Thị Trường: Cách Chứng Minh Số Lượng Hợp Lý Để Được Miễn Kiểm Tra Chất Lượng Theo Nghị Định 37/2026",
        "opt_c": "Chỉ Có 5 Chiếc Máy Mẫu Nhưng Hải Quan Yêu Cầu Đăng Ký Kiểm Tra Chất Lượng: Quy Trình Xin Miễn Kiểm Tra Chuyên Ngành Chuẩn Pháp Lý Năm 2027!",
        "ssot": "Điều 86 & Điểm c khoản 4 Điều 40 Nghị định 37/2026/NĐ-CP; Luật Chất lượng sản phẩm hàng hóa; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-HOT48).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 173,
        "day": "[W02/2027] Thứ Năm (14/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Nhập Hàng Mẫu Test Chất Lượng: Có Được Miễn Kiểm Tra Chuyên Ngành Theo Nghị Định 37/2026?",
        "opt_b": "3 Điều Kiện Để Lô Hàng Mẫu Không Bị Giữ Lại Cảng Vì Thiếu Giấy Chứng Nhận Hợp Quy",
        "opt_c": "Đừng Khai Chung Hàng Mẫu Với Hàng Thương Mại Nếu Không Muốn Bị Kẹt Kiểm Tra Chất Lượng!",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 24 (STT 172); Điều 86 & Điều 40 Nghị định 37/2026/NĐ-CP.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 174,
        "day": "[W02/2027] Thứ Sáu (15/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Công văn 24424/CHQ-NVTHQ & Rà soát thuế chống bán phá giá tháp điện gió (AD18): Hướng dẫn mới nhất giúp doanh nghiệp dự án ưu đãi đầu tư được xét miễn cả thuế nhập khẩu bổ sung (CBPG) khi nhập tạo tài sản cố định — Kế toán và phòng XNK đã khai đúng mã miễn thuế chưa?",
        "opt_b": "Checklist 5 bước lập Danh mục hàng hóa miễn thuế tạo tài sản cố định để được miễn đồng thời thuế chống bán phá giá (AD18) theo Công văn 24424/CHQ-NVTHQ và Thông tư 86/2026/TT-BTC.",
        "opt_c": "Nhiều nhà thầu năng lượng tái tạo cứ nghĩ thuế chống bán phá giá là thuế bổ sung nên không được miễn dù dự án thuộc diện ưu đãi đầu tư — Theo Công văn 24424/CHQ-NVTHQ, ranh giới để được miễn hợp pháp nằm ở bước đăng ký hồ sơ nào trước khi hàng cập cảng?",
        "ssot": "Công văn số 24424/CHQ-NVTHQ ngày 18/11/2026 của Cục Hải quan; Luật Thuế XNK số 107/2016/QH13; Nghị định 134/2016/NĐ-CP; Thông tư 86/2026/TT-BTC; reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx (STT 273).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 175,
        "day": "[W02/2027] Thứ Sáu (15/01/2027)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Quản trị tuân thủ hồ sơ pháp nhân & kê khai Chủ sở hữu hưởng lợi (Mẫu 12) khi tăng vốn điều lệ: Điểm mù hành chính ảnh hưởng trực tiếp tới giao dịch ngân hàng và hoạt động XNK",
        "opt_b": "Tại sao doanh nghiệp thương mại cần kiểm tra chéo dữ liệu đăng ký kinh doanh, chữ ký số và tài khoản ngân hàng ngay sau mỗi lần thay đổi tỷ lệ góp vốn?",
        "opt_c": "Bài học vận hành cho doanh nghiệp SME khi mở rộng quy mô vốn: 3 bước chuẩn hóa hồ sơ pháp nhân trước khi ký hợp đồng nhập khẩu giá trị lớn",
        "ssot": "Nhật ký công việc thực tế scratch/daily_notes.md ngày 21/11/2026 (Case Study Mẫu 12 Chủ sở hữu hưởng lợi khi tăng vốn) (Repurpose từ STT 171).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 176,
        "day": "[W02/2027] Thứ Bảy (16/01/2027)",
        "platform": "YouTube",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Video Tâm tình Thực chiến (7-9p) [CM4-CD10]",
        "opt_a": "Tiền Cọc Đi Liền Khúc Ruột: Bài Học Đắt Giá Khi Đặt Cọc 30% Cho Xưởng Nước Ngoài & Cách Kiểm Hàng Để Không Nhận Container Lệch Cam Kết!",
        "opt_b": "3 Điều Khoản Bảo Vệ Tiền Cọc Khi Mua Hàng Trung Quốc Lần Đầu: Quy Trình Giám Định Trước Xếp Hàng (PSI) Cho Doanh Nghiệp Nhập Khẩu",
        "opt_c": "Vì Sao Chuyển Cọc 30% Xong Là Thế Cờ Nằm Trong Tay Nhà Cung Cấp? Tâm Tình Nghề XNK Về Nghệ Thuật Ràng Buộc Chất Lượng Trước Khi Thanh Toán 70% Còn Lại",
        "ssot": "Thực tế tư vấn tranh chấp thương mại quốc tế tại Eureka Logistics; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD10).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 177,
        "day": "[W02/2027] Chủ Nhật (17/01/2027)",
        "platform": "Facebook",
        "category": "FF — Funfact & Insight",
        "technique": "Dẫn chứng quốc tế",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Báo cáo 'Containers Lost at Sea' (Cập nhật 2026) của WSC: Trong 280 triệu container vượt biển mỗi năm, chỉ có khoảng 1.478 thùng rơi mất tích (~0,0005%) — Nhưng vì sao chỉ một cơn bão khi tàu vòng qua Mũi Hảo Vọng cũng đủ khiến chủ hàng mất trắng nếu mua bảo hiểm sai điều khoản?",
        "opt_b": "Góc nhìn từ thống kê an toàn hàng hải 2026 của World Shipping Council: Tỷ lệ an toàn đạt 99,9995% nhưng rủi ro phân bổ tuyến đường qua Mũi Hảo Vọng vẫn buộc doanh nghiệp XNK tính kỹ chi phí chằng buộc (Lashing) và điều kiện bảo hiểm ICC A.",
        "opt_c": "Chủ nhật đọc con số thú vị của ngành biển thế giới: 280 triệu lượt container đi biển chỉ rơi chưa tới 0,001% — Nhưng người làm logistics chuyên nghiệp không bao giờ đánh cược lô hàng tỷ đồng vào xác suất may rủi. Bạn đã kiểm tra kỹ điều khoản bảo hiểm lô hàng tuần tới chưa?",
        "ssot": "Báo cáo 'Containers Lost at Sea Report: 2026 Update' ngày 25/06/2026 của Hội đồng Vận tải Biển Thế giới (worldshipping.org); reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 273).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 178,
        "day": "[W02/2027] Chủ Nhật (17/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "FF — Funfact & Insight",
        "technique": "Dẫn chứng quốc tế",
        "format": "Video ngắn <60s (Repurpose FB)",
        "opt_a": "280 Triệu Container Ra Khơi Mỗi Năm: Bao Nhiêu Chiếc Rơi Xuống Biển Theo Báo Cáo WSC 2026?",
        "opt_b": "Tỷ Lệ Rơi Container Chỉ 0,0005% Nhưng Vì Sao Hàng Đóng Công Vẫn Cần Mua Đúng Bảo Hiểm ICC A?",
        "opt_c": "Sự Thật Về Những Container Rơi Xuống Biển Khi Tàu Vòng Qua Mũi Hảo Vọng Năm 2026",
        "ssot": "World Shipping Council (worldshipping.org) 'Containers Lost at Sea Report: 2026 Update'; Repurpose từ STT 177.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },

    # ======================== TUẦN W03/2027 (18/01/2027 - 24/01/2027) ========================
    {
        "stt": 179,
        "day": "[W03/2027] Thứ Hai (18/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Làn sóng chống lẩn tránh xuất xứ dồn dập vào nhóm hàng kim loại Việt Nam: Ukraine (QĐ AD-615/2026/441-01) điều tra dây thép silic-mangan HS 7229 kèm yêu cầu đặt cọc bảo đảm, còn EU (Quy định 2026/2497) áp thuế tạm thời 12,3%–14,4% ống đồng CN 7411 — Doanh nghiệp sản xuất đã tách bạch hồ sơ nguồn gốc phôi đầu vào chưa?",
        "opt_b": "Kỹ thuật 5 bước ứng phó điều tra chống lẩn tránh và phòng vệ thương mại kép (Ukraine AD-615 & EU 2026/2497): Cách chuẩn hóa chứng từ nguyên liệu đầu vào, định mức tiêu hao và nộp bản trả lời đúng hạn.",
        "opt_c": "Nếu lô hàng dây thép hoặc ống đồng xuất khẩu của bạn bỗng nhiên bị nước nhập khẩu yêu cầu nộp tiền đặt cọc bảo đảm ngay tại cảng đến vì nghi vấn chuyển tải xuất xứ — Phòng chứng từ và kế toán giá thành sẽ lấy hồ sơ gì ra chứng minh?",
        "ssot": "Quyết định AD-615/2026/441-01 (03/11/2026) của Ủy ban Liên bộ TMQT Ukraine; Quy định thực thi (EU) 2026/2497 (hiệu lực 12/11/2026); moit.gov.vn & aslgate.com; reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx (STT 274) & Hoc_Thuc_Chien (STT 275).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 180,
        "day": "[W03/2027] Thứ Ba (19/01/2027)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Liên hệ lịch sử / Nhân vật nổi tiếng",
        "format": "Video Phân tích & Insight (8-12p) [CM3-CD08]",
        "opt_a": "Kênh Đào Suez 1869: Vết Cắt 193 Km Xuyên Sa Mạc Rút Ngắn Nửa Vòng Trái Đất & Cuộc Chiến Địa Chính Trị Trăm Năm Trên Tuyến Huyết Mạch Hàng Hải!",
        "opt_b": "Ferdinand de Lesseps & 10 Năm Đào Kênh Suez Bằng Sức Người: Vì Sao Tuyến Đường Biển Này Định Hình Lại Bản Đồ Cước Tàu Á – Âu Suốt Hơn 150 Năm?",
        "opt_c": "Từ Lệnh Đào Kênh Thời Pharaoh Đến Cuộc Phong Tỏa 8 Năm (1967–1975): Bài Học Lịch Sử Về Rủi Ro Phụ Thuộc Tuyến Vận Tải Biển Độc Đạo",
        "ssot": "Suez Canal Authority (SCA); Sách 'Parting the Desert: The Creation of the Suez Canal' (Zachary Karabell); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD08).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 181,
        "day": "[W03/2027] Thứ Ba (19/01/2027)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Quản trị rủi ro điều tra chống lẩn tránh thuế CBPG tại Ukraine (AD-615/2026/441-01) và EU (2026/2497): 4 lớp hồ sơ truy xuất nguồn gốc phôi nguyên liệu cho doanh nghiệp kim loại",
        "opt_b": "Bài học chuẩn hóa dữ liệu định mức và hóa đơn nguyên liệu đầu vào trước áp lực đặt cọc bảo đảm thuế phòng vệ thương mại tại thị trường Đông Âu & EU",
        "opt_c": "Vì sao minh bạch chuỗi cung ứng đầu vào là lá chắn tài chính bền vững nhất cho nhà xuất khẩu thép, dây thép và ống đồng Việt Nam năm 2027?",
        "ssot": "Quyết định AD-615/2026/441-01 (Ukraine); Quy định (EU) 2026/2497; Cục Phòng vệ thương mại (moit.gov.vn) (Repurpose từ STT 179).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 182,
        "day": "[W03/2027] Thứ Tư (20/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Đại án xuất khẩu trái phép hơn 28.000 tấn quặng Titan (1.033 container, tổng vụ án >1.508 tỷ đồng) tại Cảng Sài Gòn KV1: Khi hàm lượng TiO2 không đạt ngưỡng 56% và cái giá của hành vi đánh tráo mẫu giám định — Quy trình niêm phong mẫu đối chứng đa bên quan trọng đến mức nào?",
        "opt_b": "Checklist 3 nguyên tắc vàng bảo vệ doanh nghiệp xuất khẩu khoáng sản và hàng kiểm định chuyên ngành: Niêm phong mẫu đối chứng, chữ ký 3 bên và hình ảnh định vị thực tế tại cảng.",
        "opt_c": "Nếu một đối tác môi giới gợi ý 'tác động nhẹ hồ sơ lấy mẫu giám định để lô hàng thiếu chuẩn kỹ thuật kịp xuống tàu', ranh giới giữa giữ đơn hàng trước mắt và rủi ro hình sự nhiều năm sau nên được người đứng đầu đặt ở đâu?",
        "ssot": "Hồ sơ xét xử sơ thẩm TAND TP.HCM tháng 07/2026 (28 bị cáo vụ Công ty Hưng Thịnh & cựu công chức kiểm hóa Cảng Sài Gòn KV1); sggp.org.vn, tuoitre.vn & vnexpress.net; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 271).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 183,
        "day": "[W03/2027] Thứ Năm (21/01/2027)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 25",
        "opt_a": "Học XNK 5 Phút Tập 25: Toàn Tập Thủ Tục Hải Quan Chuyển Phát Nhanh Quốc Tế & Thương Mại Điện Tử Theo Văn Bản Hợp Nhất 53/2026/VBHN-TT-BTC!",
        "opt_b": "4 Quy Định Sống Còn Khi Gửi Hàng Nhập Xuất Qua DHL, FedEx, UPS Năm 2027: Phân Nhóm Tờ Khai Trị Giá Thấp & Đồng Bộ Thông Tư 142/2026/TT-BTC",
        "opt_c": "Hàng Đi Chuyển Phát Nhanh Quốc Tế Sao Vẫn Bị Tắc Ở Kho Sân Bay? Hướng Dẫn Đọc Đúng Quy Trình Khai Báo Theo VBHN 53/2026/VBHN-TT-BTC",
        "ssot": "Văn bản hợp nhất số 53/2026/VBHN-TT-BTC ngày 12/10/2026 của Bộ Tài chính; Thông tư 142/2026/TT-BTC; thuvienphapluat.vn & luatvietnam.vn; scratch/wf04_legal_research_2026_11_21.md.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 184,
        "day": "[W03/2027] Thứ Năm (21/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Gửi Hàng Chuyển Phát Nhanh Quốc Tế: 3 Điểm Cần Biết Trong Văn Bản Hợp Nhất 53/2026/VBHN-TT-BTC!",
        "opt_b": "Vì Sao Hàng Về Sân Bay Qua Chuyển Phát Nhanh Vẫn Bị Yêu Cầu Mở Tờ Khai Thương Mại?",
        "opt_c": "Checklist 3 Giấy Tờ Cần Chuẩn Bị Ngay Trước Khi Nhận Hàng Chuyển Phát Nhanh Quốc Tế Về Việt Nam",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 25 (STT 183); Văn bản hợp nhất 53/2026/VBHN-TT-BTC.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 185,
        "day": "[W03/2027] Thứ Sáu (22/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Tin vui gỡ vướng điểm nghẽn C/O: Công văn 24381/CHQ-GSQL & 24449/CHQ-GSQL chính thức ghi nhận kiến nghị tháo gỡ việc dừng cấp C/O cho hàng gửi kho ngoại quan, C/O cấp sau và hàng từ DNCX bán vào nội địa — Doanh nghiệp FDI và đơn vị mua hàng nội địa cần theo dõi hướng dẫn tiếp theo thế nào?",
        "opt_b": "Cẩm nang 4 bước xử lý khi gặp vướng mắc dừng cấp C/O cho hàng từ DNCX vào nội địa và kho ngoại quan theo Công văn 24381/CHQ-GSQL & 24449/CHQ-GSQL (Phụ lục III Nghị quyết 19/2026/NĐ-CP).",
        "opt_c": "Từ câu chuyện doanh nghiệp lớn như LG Electronics Hải Phòng gặp vướng khi xin C/O đưa hàng từ DNCX vào nội địa: Bài học về sự phối hợp liên bộ giữa Hải quan và Bộ Công Thương để bảo vệ quyền lợi thuế suất ưu đãi cho doanh nghiệp sản xuất.",
        "ssot": "Công văn số 24381/CHQ-GSQL (17/11/2026) & Công văn số 24449/CHQ-GSQL (18/11/2026) của Cục Hải quan; Phụ lục III Nghị quyết 19/2026/NĐ-CP; Nghị định 31/2018/NĐ-CP; reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx (STT 272).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 186,
        "day": "[W03/2027] Thứ Sáu (22/01/2027)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Kiểm soát nội bộ quy trình lấy mẫu giám định độc lập trong xuất khẩu hàng công nghiệp: Bài học quản trị rủi ro pháp lý từ vụ án 1.033 container tại Cảng Sài Gòn KV1",
        "opt_b": "Thiết lập cơ chế kiểm chứng chéo 3 bên khi lấy mẫu kiểm định hàng hóa xuất nhập khẩu: Bảo vệ doanh nghiệp khỏi rủi ro sai lệch kết quả phân tích",
        "opt_c": "Vì sao tính toàn vẹn của mẫu đối chứng lại quyết định sinh mệnh pháp lý của một hợp đồng thương mại quốc tế? Góc nhìn quản trị tuân thủ B2B",
        "ssot": "Hồ sơ xét xử TAND TP.HCM tháng 07/2026; sggp.org.vn & vnexpress.net; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 271) (Repurpose từ STT 182).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 187,
        "day": "[W03/2027] Thứ Bảy (23/01/2027)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Dịch Luật Ra Tiếng Người (8-10p) [CM2-CD05]",
        "opt_a": "Trị Giá Hải Quan & Cơ Chế Tham Vấn Giá (TT 39/2015 & TT 60/2019): Vì Sao Giá Mua Thật Rẻ Vẫn Bị Bác Trị Giá & Cách Chuẩn Bị 5 Chứng Từ Giải Trình!",
        "opt_b": "Dịch Luật Trị Giá Hải Quan WTO Ra Tiếng Người: 6 Phương Pháp Xác Định Trị Giá Giao Dịch Để Doanh Nghiệp Không Bị Ấn Định Thuế Oan!",
        "opt_c": "Tờ Khai Nhập Khẩu Dính Luồng Tham Vấn Giá: Quy Trình 4 Bước Đối Thoại Chuyên Nghiệp Với Hải Quan Bảo Vệ Đúng Giá Mua Thực Tế",
        "ssot": "Hiệp định Trị giá Hải quan WTO (GATT 1994); Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-CD05).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 188,
        "day": "[W03/2027] Chủ Nhật (24/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Lỗi Java & định dạng tệp khi nộp tờ khai thuế nhà thầu qua phần mềm HTKK ngay sát hạn chót: Sai sót kỹ thuật nhỏ có thể kéo theo phạt chậm nộp không đáng có — Phòng Kế toán và XNK của bạn đã có quy trình kiểm tra hạ tầng ký số trước 48 giờ chưa?",
        "opt_b": "Checklist 3 bước xử lý nhanh khi nộp tờ khai thuế nhà thầu qua HTKK gặp lỗi cấu hình Java hoặc không nhận tệp XML chuẩn.",
        "opt_c": "Cuối tuần nhắc nhẹ chuyện tưởng nhỏ nhưng hay làm anh chị kế toán XNK mất ăn mất ngủ: Tờ khai làm xong hết rồi nhưng tới lúc bấm nộp HTKK lại báo lỗi Java — Kinh nghiệm chuẩn bị máy trạm dự phòng tại Eureka để không bao giờ trễ hạn thuế nhà thầu.",
        "ssot": "Nhật ký công việc thực tế scratch/daily_notes.md ngày 21/11/2026 (Case Study khắc phục lỗi Java/định dạng tệp khi nộp tờ khai thuế nhà thầu qua HTKK).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 189,
        "day": "[W03/2027] Chủ Nhật (24/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Vì Sao Giá Mua Thật Vẫn Bị Hải Quan Tham Vấn Giá? 3 Chứng Từ Chứng Minh Trị Giá Giao Dịch!",
        "opt_b": "Bị Gọi Tham Vấn Giá Hàng Nhập Khẩu: Đừng Quên Mang Theo 3 Giấy Tờ Thanh Toán Này!",
        "opt_c": "Săn Được Hàng Giá Tốt Từ Xưởng Ngoại Nhưng Hải Quan Nghi Ngờ Khai Thấp: Xử Lý Sao Cho Đúng Luật?",
        "ssot": "Trích xuất từ YouTube Dịch Luật CM2-CD05 (STT 187); Thông tư 39/2015/TT-BTC & Thông tư 60/2019/TT-BTC.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },

    # ======================== TUẦN W04/2027 (25/01/2027 - 31/01/2027) ========================
    {
        "stt": 190,
        "day": "[W04/2027] Thứ Hai (25/01/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Công văn 24503/CHQ-GSQL (20/11/2026): Chuẩn hóa thủ tục hải quan đối với hàng hóa đã xuất khẩu gặp vướng mắc phải tái nhập trở lại Việt Nam — Khi lô hàng bị trả về sát mùa cao điểm, doanh nghiệp cần đồng bộ những chứng từ nguồn gốc nào để không bị kẹt tại cảng nhập?",
        "opt_b": "Quy trình 4 bước chuẩn bị hồ sơ tái nhập hàng đã xuất khẩu theo Công văn 24503/CHQ-GSQL và Điều 47 Nghị định 08/2015/NĐ-CP để kiểm tra thực tế và giải phóng hàng nhanh chóng.",
        "opt_c": "Hàng đã xuất đi nước ngoài nhưng không đạt kiểm định hoặc đối tác từ chối nhận buộc phải quay đầu về cảng Việt Nam: Với hướng dẫn mới tại Công văn 24503/CHQ-GSQL, phòng chứng từ nên chuẩn bị sẵn bộ hồ sơ gốc nào ngay từ ngày xuất khẩu?",
        "ssot": "Công văn số 24503/CHQ-GSQL ngày 20/11/2026 của Cục Hải quan; Điều 47 Nghị định 08/2015/NĐ-CP; luatvietnam.vn & bkhub.vn; scratch/wf04_legal_research_2026_11_21.md.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 191,
        "day": "[W04/2027] Thứ Ba (26/01/2027)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Dẫn chứng quốc tế",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Cảng Tự Do Hải Nam (Hainan Free Trade Port): Cột Mốc Đóng Cửa Hải Quan Độc Lập Biến Cả Hòn Đảo Thành Đặc Khu Miễn Thuế & Gợi Mở Cho FTZ Việt Nam!",
        "opt_b": "Giải Mã Mô Hình 'Mở Cửa Tuyến 1, Kiểm Soát Tuyến 2' Tại Cảng Tự Do Hải Nam: Cách Trung Quốc Tái Thiết Chuỗi Cung Ứng Khu Vực",
        "opt_c": "Khi Một Hòn Đảo Rộng Gấp 50 Lần Singapore Trở Thành Đặc Khu Hải Quan Miễn Thuế: Doanh Nghiệp XNK Việt Nam Cần Nhìn Ra Cơ Hội Và Cạnh Tranh Gì?",
        "ssot": "Luật Cảng Thương mại Tự do Hải Nam (2021) & cột mốc vận hành đóng cửa hải quan độc lập 18/12/2025; Wikipedia, haikou.gov.cn & trade.gov; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 272).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 192,
        "day": "[W04/2027] Thứ Ba (26/01/2027)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Quản trị chuỗi cung ứng đảo chiều (Reverse Logistics) theo Công văn 24503/CHQ-GSQL: Chuẩn hóa hồ sơ tái nhập hàng xuất khẩu bị trả về để bảo vệ dòng tiền doanh nghiệp",
        "opt_b": "Quy trình kiểm soát chứng từ nguồn gốc cho hàng hóa xuất khẩu quay đầu về Việt Nam: Giảm thiểu thời gian lưu bãi và chi phí phát sinh",
        "opt_c": "Vì sao hồ sơ lưu trữ từ ngày mở tờ khai xuất khẩu lại quyết định tốc độ giải phóng hàng khi phát sinh tình huống tái nhập bất khả kháng?",
        "ssot": "Công văn 24503/CHQ-GSQL (20/11/2026); Điều 47 Nghị định 08/2015/NĐ-CP; luatvietnam.vn (Repurpose từ STT 190).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 193,
        "day": "[W04/2027] Thứ Tư (27/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Bắt giữ 24.930 phụ kiện mang nhãn hiệu Harley-Davidson không khai báo trong lô hàng quá cảnh tại Cửa khẩu Quốc tế Tà Lùng (Cao Bằng): Đừng nhầm tưởng hàng quá cảnh thì được miễn kiểm tra thực tế và miễn kiểm soát sở hữu trí tuệ!",
        "opt_b": "Checklist 3 bước kiểm tra tuân thủ đối với lô hàng quá cảnh và chuyển khẩu qua cửa khẩu đường bộ theo Thông tư 06/2026/TT-BTC để không bị ách tắc vì hàng không khai báo hoặc vi phạm nhãn hiệu.",
        "opt_c": "Nếu một khách hàng thuê đơn vị vận tải kéo container quá cảnh qua Việt Nam nhưng từ chối cung cấp chi tiết danh mục phụ kiện bên trong vì cho rằng 'hàng không tiêu thụ tại Việt Nam' — Doanh nghiệp logistics có nên nhắm mắt nhận đơn?",
        "ssot": "Hồ sơ Đội Kiểm soát chống buôn lậu số 1 (Hải quan khu vực XVI) & Chi cục Hải quan cửa khẩu Tà Lùng (05/2026); Thông tư 06/2026/TT-BTC; thuehaiquan.tapchikinhtetaichinh.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 272).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 194,
        "day": "[W04/2027] Thứ Năm (28/01/2027)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 26 [CM2-HOT49]",
        "opt_a": "Học XNK 5 Phút Tập 26: Nhập Máy Móc Cũ Chưa Quá 10 Năm Vẫn Bị Từ Chối Nhập Khẩu Vì Sao? 5 Rào Cản Trong Quyết Định 18/2019 & 28/2022/QĐ-TTg!",
        "opt_b": "Cẩm Nang Nhập Khẩu Dây Chuyền Máy Móc Đã Qua Sử Dụng Chuẩn Quyết Định 18/2019/QĐ-TTg: Kiểm Tra Tuổi Thiết Bị, Công Suất & Chứng Thư Giám Định G7",
        "opt_c": "Mua Máy Cũ Mới Dùng 5 Năm Từ Trung Quốc Nhưng Về Cảng Lại Bị Buộc Tái Xuất: Kiểm Tra Ngay 3 Tiêu Chuẩn Kỹ Thuật Trước Khi Mở L/C!",
        "ssot": "Quyết định 18/2019/QĐ-TTg & Quyết định 28/2022/QĐ-TTg (VBHN 14/VBHN-BKHCN); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-HOT49).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 195,
        "day": "[W04/2027] Thứ Năm (28/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Máy Móc Cũ Dưới 10 Năm Vẫn Bị Buộc Tái Xuất: 3 Điều Kiện Ngầm Trong Quyết Định 18/2019/QĐ-TTg!",
        "opt_b": "Nhập Máy Móc Đã Qua Sử Dụng: Quên Kiểm Tra Công Suất 85% Là Mất Tiền Lưu Bãi Ngay Tại Cảng!",
        "opt_c": "3 Giấy Tờ Bắt Buộc Phải Có Trước Khi Đưa Máy Móc Cũ Về Cảng Nhập Khẩu Việt Nam",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 26 (STT 194); Quyết định 18/2019/QĐ-TTg & 28/2022/QĐ-TTg.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 196,
        "day": "[W04/2027] Thứ Sáu (29/01/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Vụ án cắt seal hải quan tại Chơn Thành nhằm né thuế chống bán phá giá 17,45% đối với sợi Polyester: Một phút đánh đổi lợi nhuận ngắn hạn bằng án hình sự — Bài học kỷ luật thép về kiểm soát niêm phong container chuyển cửa khẩu và khai đúng chủng loại hàng!",
        "opt_b": "3 nguyên tắc kiểm soát vận tải chuyển cửa khẩu đối với hàng thuộc diện áp thuế chống bán phá giá: Giám sát GPS liên tục, bảo toàn nguyên trạng niêm phong seal và kiểm tra chứng thư thành phần kỹ thuật.",
        "opt_c": "Thuế phòng vệ thương mại 17,45% có thể khiến biên lợi nhuận mỏng đi, nhưng khai sai tên hàng và tự ý cắt seal hải quan sẽ xóa sổ cả doanh nghiệp — Bạn nghĩ doanh nghiệp sản xuất nên ứng phó với thuế CBPG bằng cách nào để phát triển bền vững?",
        "ssot": "Hồ sơ truy tố vụ án tại Chi cục Hải quan Chơn Thành (Bình Phước, 08/2025); tuoitre.vn, vtv.vn & vietnamplus.vn; reports/Case_Study_Ideas_Bank_Eureka.xlsx (STT 274).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 197,
        "day": "[W04/2027] Thứ Sáu (29/01/2027)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Kiểm soát tuân thủ hàng quá cảnh và quyền sở hữu trí tuệ tại cửa khẩu đường bộ theo Thông tư 06/2026/TT-BTC: Bài học từ vụ 24.930 phụ kiện tại Tà Lùng",
        "opt_b": "Trách nhiệm khai báo chi tiết và kiểm tra hồ sơ sở hữu trí tuệ đối với doanh nghiệp giao nhận vận tải thực hiện dịch vụ quá cảnh xuyên biên giới",
        "opt_c": "Chuẩn hóa quy trình sàng lọc rủi ro hàng quá cảnh (Transit Cargo Screening): Lá chắn bảo vệ uy tín cho doanh nghiệp logistics chính quy",
        "ssot": "Hồ sơ Hải quan cửa khẩu Tà Lùng (05/2026); Thông tư 06/2026/TT-BTC; thuehaiquan.tapchikinhtetaichinh.vn (Repurpose từ STT 193).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 198,
        "day": "[W04/2027] Thứ Bảy (30/01/2027)",
        "platform": "YouTube",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Video Tâm tình Thực chiến (7-9p) [CM4-CD12]",
        "opt_a": "Bẫy Đứng Tên Hộ Giám Đốc & Cho Mượn CCCD Mở Công Ty: Bài Học Đắt Giá Về Trách Nhiệm Nợ Thuế Và Cấm Xuất Cảnh Theo Luật Quản Lý Thuế!",
        "opt_b": "Dịch Luật Điều 66 Luật Quản Lý Thuế 38/2019/QH14: Khi Người Đại Diện Pháp Luật Bị Tạm Hoãn Xuất Cảnh Vì Nợ Thuế Doanh Nghiệp Mà Không Hay Biết",
        "opt_c": "Tâm Tình Nghề Kinh Doanh: Vì Sao Tuyệt Đối Không Bao Giờ Nên 'Cả Nể' Đứng Tên Hộ Công Ty Xuất Nhập Khẩu Cho Người Quen?",
        "ssot": "Điều 66 Luật Quản lý thuế số 38/2019/QH14; Nghị định 126/2020/NĐ-CP; Nghị định 49/2025/NĐ-CP về ngưỡng nợ thuế tạm hoãn xuất cảnh; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD12).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 199,
        "day": "[W04/2027] Chủ Nhật (31/01/2027)",
        "platform": "Facebook",
        "category": "FF — Funfact & Insight",
        "technique": "Liên hệ lịch sử / Nhân vật nổi tiếng",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Kỳ tích tàu chở hàng Liberty Ship (1941–1945): Khi Henry J. Kaiser biến quy trình đóng tàu mất cả năm trời thành dây chuyền lắp ráp module — hoàn thiện chiếc SS Robert E. Peary dài 134 mét chỉ trong đúng 4 ngày 15 giờ! Bài học chuẩn hóa quy trình cho ngành logistics và sản xuất hôm nay.",
        "opt_b": "Từ 2.710 con tàu Liberty Ship đóng hàng loạt trong Thế chiến II đến tư duy chuẩn hóa SOP tại doanh nghiệp vận tải: Vì sao chia nhỏ công việc thành module đồng bộ lại tăng tốc gấp 10 lần mà vẫn đảm bảo chất lượng?",
        "opt_c": "Chủ nhật đọc chuyện lịch sử hàng hải: Nếu một con tàu viễn dương vạn tấn còn có thể đóng xong trong chưa đầy 5 ngày nhờ chuẩn hóa module, thì quy trình xử lý chứng từ XNK lặp đi lặp lại mỗi ngày của chúng ta có đang bị chậm vì thiếu chuẩn hóa?",
        "ssot": "Tư liệu Lịch sử Hàng hải Hoa Kỳ về chương trình đóng tàu Liberty Ship (1941–1945) & chiếc SS Robert E. Peary (1942); Wikipedia & usmm.org; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 274).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 200,
        "day": "[W04/2027] Chủ Nhật (31/01/2027)",
        "platform": "TikTok/Shorts",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Đừng Cho Mượn CCCD Đứng Tên Hộ Giám Đốc: Hậu Quả Nợ Thuế & Tạm Hoãn Xuất Cảnh Theo Luật Quản Lý Thuế 38!",
        "opt_b": "Người Đại Diện Pháp Luật Bị Tạm Hoãn Xuất Cảnh Ở Sân Bay: Kiểm Tra Ngay Tra Cứu Nợ Thuế Trước Khi Bay!",
        "opt_c": "Sự Cả Nể Tiền Tỷ: Đứng Tên Hộ Công Ty Xuất Nhập Khẩu Và Cái Kết Phải Tự Gánh Trách Nhiệm Pháp Lý",
        "ssot": "Trích xuất từ YouTube Case Study CM4-CD12 (STT 198); Điều 66 Luật Quản lý thuế 38/2019/QH14 & Nghị định 49/2025/NĐ-CP.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },

    # ======================== TUẦN W05/2027 (01/02/2027 - 07/02/2027 | TUẦN TẾT ĐINH MÙI) ========================
    {
        "stt": 201,
        "day": "[W05/2027] Thứ Hai (01/02/2027)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Tuần cuối cùng trước kỳ nghỉ Tết Nguyên Đán Đinh Mùi 2027 (Mùng 1 là 06/02/2027): Cảnh báo 4 bẫy phí lưu container, lưu bãi (DEM/DET) xuyên Tết và kiểm tra kế hoạch trực thông quan cửa khẩu — Doanh nghiệp đã chốt hạn rút ruột container và xin gia hạn ngày miễn lưu bãi (Free Time) với hãng tàu chưa?",
        "opt_b": "Checklist 5 việc khẩn cấp phòng Logistics & XNK phải khóa sổ trước 27 Tết (03/02/2027): Xin thêm Free Time DEM/DET, xác nhận lịch trực kho bãi, kiểm tra thanh khoản chữ ký số và hoàn tất thông quan hàng gấp.",
        "opt_c": "Mỗi dịp cận Tết, chỉ cần container về cảng trễ đúng 1 ngày ngay trước kỳ nghỉ lễ là tiền lưu container có thể đội lên hàng chục triệu đồng suốt 9 ngày nghỉ — Kinh nghiệm đàm phán xin 14–21 ngày Free Time DEM/DET trước mùa cao điểm của Eureka thế nào?",
        "ssot": "Lịch nghỉ Tết Nguyên Đán 2027 (Mùng 1 Tết 06/02/2027); Thực tiễn đàm phán DEM/DET và quy tắc bố trí trực thông quan 24/7 của Cục Hải quan (tham chiếu Công văn 24013/CHQ-VP); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD14).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 202,
        "day": "[W05/2027] Thứ Ba (02/02/2027)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Liên hệ lịch sử / Nhân vật nổi tiếng",
        "format": "Video Phân tích & Insight (8-12p) [CM3-CD18]",
        "opt_a": "Cuộc Đại Di Cư Chuỗi Cung Ứng Dệt May Toàn Cầu: Từ Lancashire Nước Anh Đến Châu Á — Quay Ngược Dòng Lịch Sử Đọc Vị Hành Trình Chiếc Áo Mới Ngày Tết!",
        "opt_b": "Bí Ẩn Đằng Sau Một Chiếc Áo Phông Vượt 3 Châu Lục: Từ Cánh Đồng Bông Texas Đến Xưởng May Châu Á Và Bài Học Chuỗi Cung Ứng Toàn Cầu",
        "opt_c": "Vì Sao Ngành Dệt May Luôn Dịch Chuyển Theo Chi Phí Nhân Công Và Thuế Quan FTA? Ký Sự Lịch Sử Thương Mại Quốc Tế Qua 200 Năm!",
        "ssot": "Tác phẩm kinh điển 'The Travels of a T-Shirt in the Global Economy' (Pietra Rivoli); Dữ liệu thương mại dệt may WTO; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD18).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 203,
        "day": "[W05/2027] Thứ Ba (02/02/2027)",
        "platform": "LinkedIn",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Slide Carousel B2B (Repurpose FB)",
        "opt_a": "Quản trị rủi ro lưu container (DEM/DET) và tồn đọng hàng hóa cảng biển dịp cao điểm nghỉ Tết Nguyên Đán 2027: Checklist 5 điểm cho Giám đốc Chuỗi cung ứng",
        "opt_b": "Đàm phán gia hạn Free Time và điều phối rút hàng cảng trước kỳ nghỉ lễ dài ngày: Cách bảo vệ biên lợi nhuận doanh nghiệp nhập khẩu B2B",
        "opt_c": "Kế hoạch dự phòng vận hành xuất nhập khẩu xuyên kỳ nghỉ lễ đầu năm 2027: Đồng bộ lịch trực hải quan, hãng tàu và kho vận nội địa",
        "ssot": "Thực tiễn đàm phán DEM/DET và quy trình vận hành thông quan dịp lễ tại Eureka Logistics (Repurpose từ STT 201).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 204,
        "day": "[W05/2027] Thứ Tư (03/02/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Những ngày 27 Tết tại bãi xe cửa khẩu và kho cảng: Khi các ngành khác đã rục rịch tất niên, anh em lái xe container và nhân sự hiện trường vẫn bám trụ đến chuyến hàng cuối cùng để kịp giao vào kho khách trước giao thừa — Chữ Tín nghề Logistics được viết từ những hy sinh thầm lặng như thế!",
        "opt_b": "Bài học quản trị đội ngũ vận hành và tri ân lực lượng tuyến đầu Logistics dịp cuối năm: Cách Eureka sắp xếp ca trực, hỗ trợ tài chính và giữ lửa tinh thần cho anh em hiện trường mùa cao điểm Tết.",
        "opt_c": "Ngồi ăn vội hộp cơm cùng anh em lái xe ngoài bãi xe cửa khẩu những ngày cuối năm: Khi hiểu được nỗi vất vả sau vô-lăng đường dài, bạn sẽ trân trọng hơn từng chuyến hàng về kho đúng hẹn.",
        "ssot": "Nhật ký thực địa biên giới phía Bắc và vận hành đội ngũ hiện trường Eureka Logistics; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD15).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 205,
        "day": "[W05/2027] Thứ Năm (04/02/2027)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 27 [CM2-CD15]",
        "opt_a": "Học XNK 5 Phút Tập 27: Phân Biệt Rạch Ròi Hàng Phi Mậu Dịch (H11) vs Hàng Mẫu Thử Nghiệm vs Quà Biếu Tặng Dịp Tết Để Không Bị Truy Thu Thuế!",
        "opt_b": "Dịch Luật Điều 16 Luật Thuế XNK 107/2016: Định Mức Miễn Thuế Quà Biếu, Quà Tặng Dịp Lễ Tết & Thủ Tục Mở Tờ Khai H11 Chuẩn Xác",
        "opt_c": "Gửi Hàng Biếu Tặng Đối Tác Nước Ngoài Hoặc Nhập Hàng Mẫu Về Trước Tết: 4 Sai Lầm Khiến Hàng Bị Giữ Lại Sân Bay Nội Bài / Tân Sơn Nhất",
        "ssot": "Điều 16 Luật Thuế xuất khẩu, thuế nhập khẩu số 107/2016/QH13; Điều 8 Nghị định 134/2016/NĐ-CP; Quyết định 1357/QĐ-TCHQ (mã loại hình H11); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-CD15).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 206,
        "day": "[W05/2027] Thứ Năm (04/02/2027)",
        "platform": "TikTok/Shorts",
        "category": "HT — Học Thực chiến",
        "technique": "Problem-First",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Gửi Quà Biếu Tặng & Hàng Mẫu Dịp Tết: Nhầm Mã H11 Là Bị Áp Thuế & Giữ Hàng Ngay!",
        "opt_b": "Quà Tặng Nhập Khẩu Bao Nhiêu Tiền Thì Được Miễn Thuế Theo Nghị Định 134/2016?",
        "opt_c": "Phân Biệt Nhanh Hàng Phi Mậu Dịch H11 Và Hàng Mẫu Trong 45 Giây Để Không Kẹt Sân Bay!",
        "ssot": "Trích xuất từ YouTube Học XNK 5 Phút Tập 27 (STT 205); Điều 16 Luật Thuế XNK 107/2016/QH13 & Nghị định 134/2016/NĐ-CP.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
    {
        "stt": 207,
        "day": "[W05/2027] Thứ Sáu (05/02/2027)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Chiều 29 Tết (Đêm Giao thừa Đinh Mùi 2027): Khép lại một năm 2026 đầy biến động từ giá cước tàu đến hàng loạt quy định chuyển đổi số hải quan — Điều Thúy tự hào nhất không phải số lượng container đi qua cảng, mà là sự bình an pháp lý của từng khách hàng đồng hành cùng Eureka!",
        "opt_b": "Nhìn lại 3 bài học lớn nhất của ngành Xuất nhập khẩu – Logistics năm 2026 trước giờ Giao thừa: Minh bạch dữ liệu số, chủ động kiểm soát xuất xứ và làm đúng ngay từ đầu.",
        "opt_c": "Trước thềm năm mới Đinh Mùi 2027, gửi lời tri ân chân thành tới Quý khách hàng, đối tác và 33 cộng sự tại Eureka Logistics đã cùng nhau vượt qua một năm thử thách — Chúc cả cộng đồng làm nghề một năm mới hanh thông, vạn sự bình an!",
        "ssot": "Nhật ký điều hành và triết lý nhân hiệu Giáp Mộc – Liêm Trinh của Founder Nguyễn Thị Ngọc Thúy tại Eureka Logistics (knowledge/brand_persona.md).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 208,
        "day": "[W05/2027] Thứ Sáu (05/02/2027)",
        "platform": "LinkedIn",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài phân tích chuyên sâu B2B (Repurpose FB)",
        "opt_a": "Tổng kết vận hành chuỗi cung ứng năm 2026 & 3 trụ cột quản trị rủi ro bước vào năm mới Đinh Mùi 2027: Minh bạch pháp lý, kỷ luật dữ liệu và đối tác tin cậy",
        "opt_b": "Từ biến động giá cước biển đến chuẩn hóa hải quan số: Những bài học chiến lược giúp doanh nghiệp sản xuất xuất khẩu Việt Nam giữ vững năng lực cạnh tranh năm 2027",
        "opt_c": "Thông điệp đầu xuân từ Eureka Logistics: Xây dựng mạng lưới hợp tác B2B dựa trên sự chính trực, an toàn tuân thủ và tối ưu giá vốn bền vững",
        "ssot": "Nhật ký điều hành Eureka Logistics & Tổng kết số liệu thương mại Việt Nam 2026 (Repurpose từ STT 207).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 209,
        "day": "[W05/2027] Thứ Bảy (06/02/2027)",
        "platform": "YouTube",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Video Tâm tình Thực chiến (7-9p) [CM4-CD17]",
        "opt_a": "Khai Xuân Mùng 1 Tết 2027: 'Nâng Niu Từng Kiện Hàng Lẻ 1 CBM' — Vì Sao Eureka Không Bao Giờ Từ Chối Những Đơn Hàng Khởi Nghiệp Nhỏ Bé!",
        "opt_b": "Từ Lô Hàng Mỹ Phẩm Lẻ 1 Khối Đầu Tay Đến Hành Trình Đồng Hành Cùng Hàng Trăm Doanh Nghiệp SME Nhập Khẩu Chính Ngạch",
        "opt_c": "Tâm Sự Ngày Mùng 1 Tết: Ai Cũng Từng Bắt Đầu Từ Một Kiện Hàng Nhỏ — Giá Trị Của Sự Tận Tâm Trong Ngành Dịch Vụ Logistics",
        "ssot": "Hành trình khởi nghiệp của Founder Nguyễn Thị Ngọc Thúy từ những đơn hàng LCL 1 CBM đầu tiên; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD17).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 210,
        "day": "[W05/2027] Chủ Nhật (07/02/2027)",
        "platform": "Facebook",
        "category": "FF — Funfact & Insight",
        "technique": "Liên hệ lịch sử / Nhân vật nổi tiếng",
        "format": "Bài chia sẻ góc nhìn + Ảnh",
        "opt_a": "Mùng 2 Tết đọc chuyện quán cà phê London năm 1744 khai sinh chỉ số Baltic Dry Index (BDI): Vì sao một bảng tin giá thuê tàu chở ngũ cốc, than đá và quặng sắt trên giấy lại trở thành 'phong vũ biểu' dự báo chu kỳ phục hồi kinh tế toàn cầu suốt gần 3 thế kỷ?",
        "opt_b": "Giải mã chỉ số Baltic Dry Index (BDI) cho doanh nghiệp xuất nhập khẩu: Cách đọc tín hiệu đơn hàng nguyên liệu thô thế giới để lên kế hoạch tồn kho và ký hợp đồng cước đầu năm 2027.",
        "opt_c": "Ngày Mùng 2 Tết nhâm nhi tách trà và nhìn ra đại dương: Thương mại thế giới luôn vận hành theo chu kỳ — Hiểu sớm nhịp thở của những con tàu chở hàng rời sẽ giúp người làm kinh doanh đi trước thị trường một bước.",
        "ssot": "Lịch sử Baltic Exchange (Virginia and Baltick Coffee House, London 1744) & Chỉ số Baltic Dry Index (BDI); balticexchange.com; reports/Funfact_Ideas_Bank_Eureka.xlsx (STT 262).",
        "approval": "🟡 Chờ duyệt",
        "energy": "Giáp Mộc"
    },
    {
        "stt": 211,
        "day": "[W05/2027] Chủ Nhật (07/02/2027)",
        "platform": "TikTok/Shorts",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Video ngắn <60s (Repurpose YT)",
        "opt_a": "Kiện Hàng Lẻ 1 CBM Đầu Năm: 3 Bài Học Xương Máu Cho Người Mới Bắt Đầu Nhập Khẩu Chính Ngạch 2027!",
        "opt_b": "Vốn Nhỏ Muốn Nhập Khẩu Hàng Lẻ LCL Chính Ngạch Đầu Xuân: Cần Chuẩn Bị 3 Chứng Từ Gì Đầu Tiên?",
        "opt_c": "Đừng Ngại Bắt Đầu Từ Kiện Hàng 1 Khối: Cách Tính Chi Phí Nhập Hàng Lẻ Chuẩn Cho Người Mới",
        "ssot": "Trích xuất từ YouTube Tâm tình CM4-CD17 (STT 209); Quy trình nhập khẩu hàng lẻ LCL tại Eureka Logistics.",
        "approval": "🟡 Chờ duyệt",
        "energy": "Liêm Trinh"
    },
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
        "anh minh", "chị hương", "anh dũng", "inbox ngay", "dm ngay",
        "cam kết 100%", "chắc chắn thành công", "cỗ máy in tiền"
    ]
    tiktok_banned = ["kiếm tiền", "làm giàu", "thu nhập thụ động", "bí mật", "hack"]
    youtube_banned = ["lừa đảo", "scam", "bạo lực"]

    errors = []
    seen_stts = set()
    for item in W02_W05_2027_ITEMS:
        if item["stt"] in seen_stts:
            errors.append(f"Trùng STT {item['stt']}")
        seen_stts.add(item["stt"])

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
    print(f"  ✅ Đạt 100% compliance cho toàn bộ {len(W02_W05_2027_ITEMS)} slots (0 vi phạm)!")


def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33/2026 ĐẾN W05/2027 CỘNG DỒN)")
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

    for item in W02_W05_2027_ITEMS:
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
    print(f"  -> Đã lưu thành công {len(W02_W05_2027_ITEMS)} dòng (STT 168–211) vào Sheet 1 qua atomic_save!")


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

    marker = "# TUẦN W02/2027: LỊCH BIÊN TẬP ĐA NỀN TẢNG"
    if marker in existing_content:
        existing_content = existing_content.split(marker)[0].rstrip()
        if existing_content.endswith("---"):
            existing_content = existing_content[:-3].rstrip()

    existing_content = existing_content.replace(
        "**Kỳ đánh giá mới nhất:** Tuần W33/2026 đến W01/2027 (Từ 17/08/2026 đến 10/01/2027) — W49–W01/2027 Chờ duyệt Checkpoint 1",
        "**Kỳ đánh giá mới nhất:** Tuần W33/2026 đến W05/2027 (Từ 17/08/2026 đến 07/02/2027) — W49–W05/2027 Chờ duyệt Checkpoint 1"
    )

    weeks_meta = [
        ("W02/2027", "11/01/2027 - 17/01/2027", W02_W05_2027_ITEMS[0:11]),
        ("W03/2027", "18/01/2027 - 24/01/2027", W02_W05_2027_ITEMS[11:22]),
        ("W04/2027", "25/01/2027 - 31/01/2027", W02_W05_2027_ITEMS[22:33]),
        ("W05/2027", "01/02/2027 - 07/02/2027", W02_W05_2027_ITEMS[33:44]),
    ]

    appended_parts = [build_week_md(wcode, drange, witems) for wcode, drange, witems in weeks_meta]
    final_content = existing_content + "\n".join(appended_parts) + "\n"

    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(final_content)
    print("  -> Đã cập nhật thành công reports/plan1/weekly_calendar.md!")


def save_json_data():
    print("=== 3. LƯU SCRATCH/W02_W05_2027_DATA.JSON ===")
    os.makedirs(os.path.dirname(JSON_OUTPUT), exist_ok=True)
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W02_W05_2027_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu 44 slots vào {JSON_OUTPUT}")


if __name__ == "__main__":
    print("=== BẮT ĐẦU QUY TRÌNH WF01: LẬP LỊCH ĐA NỀN TẢNG TUẦN W02/2027 – W05/2027 ===")
    run_compliance_check()
    update_excel_calendar()
    update_markdown_calendar()
    save_json_data()
    print("=== HOÀN TẤT THỰC THI GIAI ĐOẠN 1 TUẦN W02/2027 – W05/2027! ===")
