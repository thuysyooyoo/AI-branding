# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
import os
import shutil
from datetime import datetime

file_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"

wb = openpyxl.load_workbook(file_path)
ws1 = wb["Weekly Calendar"]

font_family = "Segoe UI"
bold_font = Font(name=font_family, size=10, bold=True)
regular_font = Font(name=font_family, size=10)

zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

cat_fills = {
    "CS": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "TN": PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid"),
    "HT": PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid"),
    "FF": PatternFill(start_color="F5D6E0", end_color="F5D6E0", fill_type="solid"),
}

platform_fills = {
    "Facebook": PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid"),
    "YouTube": PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
}

center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)

w39_data = [
    {
        "stt": 43, "day": "[W39] Thứ Hai (28/09/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Problem-First", "format": "Bài tranh luận + Infographic",
        "opt_a": "Chỉ còn đúng 17 ngày trước hạn chót 15/10/2026: Quyết định 1434/QĐ-CHQ buộc chuẩn hóa dữ liệu chứng từ điện tử sang chuẩn JSON DataFormat — Nếu doanh nghiệp XNK và phần mềm chưa nâng cấp, luồng truyền tờ khai sẽ bị từ chối ngay tức thì!",
        "opt_b": "Checklist 4 việc bộ phận Xuất nhập khẩu và IT cần làm ngay với nhà cung cấp phần mềm hải quan (ECUS5) trước ngày 15/10/2026 để thông suốt đường truyền dữ liệu JSON.",
        "opt_c": "Từ Quyết định 748 đến Quyết định 1434/QĐ-CHQ: Bước ngoặt chuyển đổi số toàn diện của Hải quan Việt Nam và bài học số hóa chứng từ cho doanh nghiệp thương mại.",
        "source": "Quyết định số 1434/QĐ-CHQ ngày 26/08/2026 của Tổng cục Hải quan (thay thế QĐ 748 & QĐ 1143) & Thời báo Tài chính VN", "status": "Chờ duyệt"
    },
    {
        "stt": 44, "day": "[W39] Thứ Ba (29/09/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Kênh Đào Panama: Tại Sao Tuyến Huyết Mạch Cắt Ngang Châu Mỹ Lại Khát Nước Ngọt & Bí Mật Đấu Giá Hàng Triệu USD Để Được Ưu Tiên Qua Kênh Trong Mùa Cao Điểm?",
        "opt_b": "52 triệu gallon nước ngọt cho mỗi lượt tàu: Lịch sử 112 năm xây dựng Kênh đào Panama và bài toán biến đổi khí hậu đang đe dọa chuỗi cung ứng toàn cầu.",
        "opt_c": "Từ Kênh đào Panama đến Kênh đào Suez: Bài học phân bổ rủi ro hành trình hàng hải và cách doanh nghiệp XNK Việt Nam né bẫy phụ phí mùa cao điểm Q4.",
        "source": "Cơ quan Quản lý Kênh đào Panama (ACP), Báo cáo Drewry Maritime Research tháng 9/2026 & Lịch sử Hàng hải Thế giới", "status": "Chờ duyệt"
    },
    {
        "stt": 45, "day": "[W39] Thứ Tư (30/09/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Problem-First", "format": "Bài Case Study + Ảnh",
        "opt_a": "Cước tàu Bờ Tây Mỹ tăng vọt, hãng tàu liên tục tung chiêu 'Blank Sailings' cắt chuyến để neo giá: Doanh nghiệp xuất khẩu làm gì để không bị rớt cont (roll cont) sát mùa Giáng sinh?",
        "opt_b": "Chiến lược đặt chỗ booking trước 3-4 tuần và nghệ thuật đàm phán điều khoản cam kết không rớt cont (No-Roll Guarantee) với Forwarder & Hãng tàu.",
        "opt_c": "Ký hợp đồng cước dài hạn (Service Contract) hay tiếp tục thả nổi theo giá Spot khi thị trường biến động? Khung 3 bước ra quyết định cho CFO & Giám đốc Logistics.",
        "source": "Chỉ số Drewry World Container Index (WCI) ngày 03/09/2026, Báo cáo World Shipping Council & Thực tiễn điều hành Eureka Logistics", "status": "Chờ duyệt"
    },
    {
        "stt": 46, "day": "[W39] Thứ Năm (01/10/2026)", "platform": "YouTube", "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến", "format": "Video Học XNK 5 phút Tập 9",
        "opt_a": "Học XNK 5 Phút: 5 Bước Tra Cứu 31 Nhóm Hàng Lưỡng Dụng & Thao Tác Khai Mã VX/KO Trên VNACCS Theo Thông Tư 42 Để Không Bị Hạ Hàng Bỏ Lại Biên Giới",
        "opt_b": "Phân biệt rạch ròi Mã VX (có giấy phép) vs Mã KO (không thuộc diện cấp phép) trên tờ khai điện tử: Mẹo 3 phút tránh bị hải quan dừng thông quan kiểm tra đột xuất.",
        "opt_c": "Checklist 4 tiêu chí kỹ thuật (công suất laser, tần số, cấu tạo vi mạch) phân định giữa hàng dân sự thông thường và hàng kiểm soát lưỡng dụng.",
        "source": "Thông tư số 42/2026/TT-BCT của Bộ Công Thương, Công văn số 19916/CHQ-GSQL của Tổng cục Hải quan & Quy trình VNACCS/VCIS", "status": "Chờ duyệt"
    },
    {
        "stt": 47, "day": "[W39] Thứ Sáu (02/10/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế", "format": "Bài tranh luận + Infographic",
        "opt_a": "Nghị định 341/2026/NĐ-CP có hiệu lực: Nhập khẩu Router, Máy chủ, Camera thông minh mà thiếu Giấy phép Mật mã Dân sự là bị ách tắc ngay tại cảng nhập — Doanh nghiệp công nghệ đã nắm rõ quy định chưa?",
        "opt_b": "Checklist 4 bước xin Giấy phép kinh doanh & Giấy phép xuất nhập khẩu sản phẩm Mật mã Dân sự tại Ban Cơ yếu Chính phủ cho doanh nghiệp thương mại CNTT.",
        "opt_c": "Ranh giới kỹ thuật giữa thiết bị viễn thông thông thường và thiết bị sử dụng chức năng mã hóa bảo mật: Làm sao chứng minh sản phẩm không thuộc diện giấy phép chuyên ngành?",
        "source": "Nghị định số 341/2026/NĐ-CP ngày 20/07/2026 của Chính phủ (hiệu lực 01/09/2026), Luật An toàn thông tin mạng & Danh mục Mật mã dân sự", "status": "Chờ duyệt"
    },
    {
        "stt": 48, "day": "[W39] Thứ Bảy (03/10/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Đế Chế Cảng Biển Rotterdam: Nơi Không Một Bóng Người Bốc Xếp Nhưng Xử Lý Hơn 400 Triệu Tấn Hàng Mỗi Năm Nhờ Biệt Đội Robot Tự Hành AGV & AI!",
        "opt_b": "Tại sao Cảng Rotterdam được mệnh danh là 'Cửa ngõ của Châu Âu'? Lịch sử 700 năm từ một làng chài nhỏ đến thương cảng trung chuyển thông minh nhất hành tinh.",
        "opt_c": "Từ mô hình Cảng thông minh Rotterdam đến Hệ thống Cảng biển nước sâu Cái Mép - Thị Vải & Cần Giờ: Tương lai logistics xanh của Việt Nam trên bản đồ hàng hải thế giới.",
        "source": "Port of Rotterdam Authority (Havenbedrijf Rotterdam N.V.), Báo cáo Smart Port Transformation & Lịch sử Logistics Châu Âu", "status": "Chờ duyệt"
    },
    {
        "stt": 49, "day": "[W39] Chủ Nhật (04/10/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy", "format": "Bài Case Study + Ảnh",
        "opt_a": "Khi khách hàng giục 'khai bừa cho xong để kịp giờ tàu': Ranh giới bản lĩnh của người làm Logistics giữa chiều khách nhất thời và bảo vệ khách hàng an toàn trọn đời trước bão hậu kiểm!",
        "opt_b": "Bài học từ một hợp đồng ủy thác: Tại sao từ chối một đơn hàng sai luật hôm nay lại cứu doanh nghiệp khỏi khoản truy thu thuế hàng tỷ đồng sau 3 năm?",
        "opt_c": "Chữ Tín trong ngành logistics không mua được bằng chiết khấu hay lời hứa suông: Câu chuyện về văn hóa minh bạch và nguyên tắc thượng tôn pháp luật tại Eureka Logistics.",
        "source": "Nhật ký vận hành thực tế tại Eureka Logistics (scratch/daily_notes.md), Luật Hải quan 2014 & Brand Persona Giáp Mộc", "status": "Chờ duyệt"
    }
]

# Update Sheet 1 Title
ws1["A1"] = "LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33, W34, W35, W36, W37, W38 & W39 CỘNG DỒN)"

# Check existing rows
current_max = ws1.max_row
existing_stts = [ws1.cell(row=r, column=1).value for r in range(5, current_max + 1)]

if 43 not in existing_stts:
    for idx, item in enumerate(w39_data):
        r_idx = current_max + 1 + idx
        fill = zebra_fill if r_idx % 2 == 0 else white_fill
        row_vals = [
            item["stt"], item["day"], item["platform"], item["category"],
            item["technique"], item["format"], item["opt_a"], item["opt_b"],
            item["opt_c"], item["source"], item["status"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=val)
            cell.border = thin_border
            cell.fill = fill
            cell.alignment = center_align if c_idx in [1, 2, 3, 11] else left_align
            
            if c_idx == 3:
                plat_key = "Facebook" if "Facebook" in val else "YouTube"
                cell.fill = platform_fills[plat_key]
                cell.font = bold_font
            elif c_idx == 4:
                cat_key = val[:2]
                if cat_key in cat_fills:
                    cell.fill = cat_fills[cat_key]
                cell.font = bold_font
            else:
                cell.font = regular_font

    # Re-apply Data Validation
    dv = DataValidation(
        type="list", 
        formula1='"Chờ duyệt,Duyệt Option A,Duyệt Option B,Duyệt Option C,Không duyệt"', 
        allow_blank=True
    )
    ws1.add_data_validation(dv)
    dv.add(f"K5:K{ws1.max_row}")

    wb.save(file_path)
    print(f"SUCCESS: Appended W39! New max row: {ws1.max_row}")
else:
    print("W39 already exists in Sheet 1!")
