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

w38_data = [
    {
        "stt": 36, "day": "[W38] Thứ Hai (21/09/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Problem-First", "format": "Bài tranh luận + Infographic",
        "opt_a": "Quy định EUDR 2026 về chuỗi cung ứng không phá rừng: Thiếu tọa độ GPS vùng trồng từng hecta là hàng gỗ, cà phê, cao su bị trả về ngay tại cảng Rotterdam — Doanh nghiệp xuất khẩu đã số hóa dữ liệu nguồn gốc chưa?",
        "opt_b": "Checklist 4 tiêu chuẩn bắt buộc trong hồ sơ Thẩm định chuyên sâu (Due Diligence Statement) theo chuẩn EUDR để nông sản Việt Nam không bị chặn tại cửa khẩu Châu Âu.",
        "opt_c": "Từ cơ chế EUDR đến rào cản CBAM: Doanh nghiệp xuất khẩu nông lâm sản nên coi 'hộ chiếu xanh' là gánh nặng chi phí hay đòn bẩy chiếm lĩnh thị trường cấp cao?",
        "source": "Regulation (EU) 2023/1115 (EUDR), Thông tư 33/2023/TT-BTC & Cổng thông tin Cục Lâm nghiệp", "status": "Chờ duyệt"
    },
    {
        "stt": 37, "day": "[W38] Thứ Ba (22/09/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Chiếc 'Hộ Chiếu Quyền Lực' Của Hàng Hóa Quốc Tế: Làm Thế Nào Sổ Tạm Quản ATA Carnet Giúp Siêu Xe F1 Và Triển Lãm Nghìn Tỷ Chu Du 80 Nước Mà Không Cần Đặt Cọc 1 Đồng Tiền Thuế?",
        "opt_b": "Bí mật ngành triển lãm toàn cầu: Tại sao các tập đoàn công nghệ mang hàng mẫu nghìn tỷ qua biên giới trong 10 phút nhờ cuốn sổ màu xanh ATA Carnet?",
        "opt_c": "Từ Công ước Brussels 1961 đến Nghị định 64/2020: Bài học tối ưu chi phí lưu ký và dòng tiền cho doanh nghiệp tham gia hội chợ thương mại quốc tế.",
        "source": "Công ước Hải quan về Sổ Tạm quản ATA (WCO/ICC 1961), Công ước Istanbul 1990 & Nghị định 64/2020/NĐ-CP", "status": "Chờ duyệt"
    },
    {
        "stt": 38, "day": "[W38] Thứ Tư (23/09/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy", "format": "Bài Case Study + Ảnh",
        "opt_a": "Ngồi lại cùng hơn 30 Founder & CEO tại AIT Kết Nối: Phía sau những mái tóc bạc sớm và bài học xương máu là ngọn lửa kiên định với con đường kinh doanh chính ngạch minh bạch!",
        "opt_b": "Muốn đi nhanh thì đi một mình, muốn đi xa thì phải đi cùng nhau: Tại sao giữa thời buổi cạnh tranh khốc liệt, các chủ doanh nghiệp lại chọn sẻ chia công thức vận hành và công nghệ thay vì giấu nghề?",
        "opt_c": "Nhật ký người làm Logistics: 3 bài học đắt giá về tinh thần trách nhiệm và bản lĩnh vượt bão của các doanh chủ Việt Nam mà không trường lớp nào giảng dạy.",
        "source": "Nhật ký vận hành thực tế tại Eureka Logistics (scratch/daily_notes.md ngày 22/08/2026) & Brand Persona Giáp Mộc", "status": "Chờ duyệt"
    },
    {
        "stt": 39, "day": "[W38] Thứ Năm (24/09/2026)", "platform": "YouTube", "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến", "format": "Video Học XNK 5 phút Tập 8",
        "opt_a": "Học XNK 5 Phút: 5 Bước Khử Trùng & Soát Xét Tiêu Chuẩn Bao Bì Gỗ ISPM 15 Để Lô Hàng Xuất Khẩu Không Bị Hải Quan Mỹ, EU Buộc Tái Xuất Hay Tiêu Hủy Tại Cảng Đến",
        "opt_b": "Phân biệt rạch ròi Xử lý Nhiệt (HT) vs Xử lý Khí (MB) trên Dấu kiểm định IPPC: Mẹo 2 phút kiểm tra pallet gỗ xuất khẩu tránh bị phạt tiền tỷ.",
        "opt_c": "Checklist 4 lỗi sơ đẳng về bao bì gỗ khiến 20 tấn hàng xuất khẩu bị giữ tại cảng và quy trình cấp chứng thư kiểm dịch thực vật chuẩn quốc tế.",
        "source": "Tiêu chuẩn Quốc tế về Kiểm dịch Thực vật ISPM 15 (FAO/IPPC), TT 33/2014/TT-BNNPTNT & Quy chuẩn Kiểm dịch Thực vật VN", "status": "Chờ duyệt"
    },
    {
        "stt": 40, "day": "[W38] Thứ Sáu (25/09/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế", "format": "Bài tranh luận + Infographic",
        "opt_a": "Chỉ cần lệch 0.5% định mức nguyên liệu trên bảng BOM là bay sạch quyền miễn thuế nhập khẩu SXXK: Doanh nghiệp gia công đã rà soát Báo cáo Quyết toán mẫu số 15 chưa?",
        "opt_b": "Phân biệt rạch ròi Loại hình E31 (Nhập SXXK) vs E21 (Nhập Gia công): 3 cạm bẫy xử lý phế liệu, phế phẩm khiến doanh nghiệp bị truy thu thuế TNDN và thuế nhập khẩu.",
        "opt_c": "Tại sao các tập đoàn FDI luôn có 'Đội phản ứng nhanh BOM'? Bài học kiểm soát tỷ lệ hao hụt thực tế và quản trị rủi ro hậu kiểm hải quan cho dn sản xuất.",
        "source": "Luật Thuế XNK số 107/2016/QH13, Nghị định 18/2021/NĐ-CP (Điều 10, 12), Thông tư 39/2018/TT-BTC sửa đổi TT 38/2015/TT-BTC", "status": "Chờ duyệt"
    },
    {
        "stt": 41, "day": "[W38] Thứ Bảy (26/09/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Sự Kiện 11/9 Đã Tái Định Hình Hải Quan Toàn Cầu Như Thế Nào: Từ Khung Tiêu Chuẩn An Ninh SAFE Đến 'Tấm Thẻ Xanh' Doanh Nghiệp Ưu Tiên AEO Miễn Kiểm Tra Mọi Lô Hàng!",
        "opt_b": "Tại sao chỉ 1% doanh nghiệp XNK Việt Nam được cấp chứng chỉ AEO nhưng chiếm tới 40% kim ngạch cả nước? Quyền lực tối thượng của luồng xanh hải quan.",
        "opt_c": "Từ 'Vệ sĩ chuỗi cung ứng' WCO SAFE đến Hệ thống Quản trị Rủi ro Tự động: Tương lai ngành logistics quốc tế khi AI kết nối thông quan toàn cầu.",
        "source": "Khung Tiêu chuẩn SAFE của Tổ chức Hải quan Thế giới (WCO 2005), TT 72/2015/TT-BTC & TT 07/2019/TT-BTC về Doanh nghiệp Ưu tiên (AEO)", "status": "Chờ duyệt"
    },
    {
        "stt": 42, "day": "[W38] Chủ Nhật (27/09/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Problem-First", "format": "Bài Case Study + Ảnh",
        "opt_a": "Mở niêm chì nguyên vẹn nhưng bên trong mất sạch 120 thùng rượu ngoại trị giá 1.8 tỷ: Chiêu trò 'Cắt Seal tráo ruột' và bài học bảo vệ hàng giá trị cao cho dn nhập khẩu!",
        "opt_b": "Checklist 4 bước bắt buộc khi bàn giao seal container tại kho CFS: Làm gì khi số seal trên B/L lệch đúng 1 ký tự so với seal thực tế trên container?",
        "opt_c": "Ai chịu trách nhiệm khi container còn nguyên chì nhưng thiếu hàng bên trong? Bóc tách ranh giới pháp lý giữa Hãng tàu, Đơn vị kho CFS và Doanh nghiệp vận tải nội địa.",
        "source": "Bộ luật Hàng hải Việt Nam 2015 (Chương VII), Quy tắc Vận chuyển Hàng hóa Đường biển Hague-Visby & Án lệ Tranh chấp Tổn thất Hàng hóa", "status": "Chờ duyệt"
    }
]

# Update Sheet 1 Title
ws1["A1"] = "LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33, W34, W35, W36, W37 & W38 CỘNG DỒN)"

# Check existing rows
current_max = ws1.max_row
existing_stts = [ws1.cell(row=r, column=1).value for r in range(5, current_max + 1)]

if 36 not in existing_stts:
    for idx, item in enumerate(w38_data):
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
    print(f"SUCCESS: Appended W38! New max row: {ws1.max_row}")
else:
    print("W38 already exists in Sheet 1!")
