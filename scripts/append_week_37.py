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

w37_data = [
    {
        "stt": 29, "day": "[W37] Thứ Hai (14/09/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Problem-First", "format": "Bài tranh luận + Infographic",
        "opt_a": "Nghị định 08 & Quy chế Tạm nhập Tái xuất: Mượn khuôn mẫu nước ngoài 18 tỷ mà quên gia hạn tờ khai G12 là bị ấn định thuế tiền tỷ như chơi — Dn phụ trợ đã rà soát hạn dùng chưa?",
        "opt_b": "Checklist 4 mốc thời gian sống còn khi làm thủ tục Tạm nhập Tái xuất (G12/G22) khuôn mẫu, linh kiện mẫu tránh bị cưỡng chế thuế hải quan.",
        "opt_c": "Tại sao các tập đoàn FDI không mua đứt máy móc mà chỉ mượn tạm nhập? Bài học tối ưu dòng tiền và kiểm soát rủi ro thanh khoản tờ khai.",
        "source": "Nghị định 08/2015/NĐ-CP (Điều 49, 50), TT 38/2015/TT-BTC & Công ước Istanbul 1990", "status": "Chờ duyệt"
    },
    {
        "stt": 30, "day": "[W37] Thứ Ba (15/09/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Bí Mật Quán Cà Phê Edward Lloyd 1688: Nơi Mẩu Tin Đồn Vỉa Hè London Đã Khai Sinh Ngành Bảo Hiểm Hàng Hải Nghìn Tỷ USD Như Thế Nào?",
        "opt_b": "Tại sao 90% hợp đồng thương mại quốc tế ngày nay đều gắn liền với điều khoản bảo hiểm Lloyd's of London? Bài học quản trị rủi ro đại dương.",
        "opt_c": "Từ tiếng chuông Lutine Bell báo tàu đắm đến thị trường định giá rủi ro toàn cầu: Lịch sử 338 năm của tập đoàn bảo hiểm quyền lực nhất thế giới.",
        "source": "Lịch sử Hàng hải Quốc tế (Lloyd's Archives 1688) & Bộ quy tắc ICC 2009", "status": "Chờ duyệt"
    },
    {
        "stt": 31, "day": "[W37] Thứ Tư (16/09/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy", "format": "Bài Case Study + Ảnh",
        "opt_a": "Tiết kiệm 12 triệu phí bảo hiểm nhưng mất trắng lô hàng sữa bột 6.5 tỷ đồng khi container bị ngấm nước biển: Cái bẫy 'mua bảo hiểm cho có' của dn nhập khẩu!",
        "opt_b": "Phân biệt rạch ròi ICC A vs ICC B vs ICC C: 3 điều kiện bảo hiểm hàng hải dn SME tuyệt đối không được chọn bừa để tránh bị từ chối bồi thường.",
        "opt_c": "Hồ sơ đòi bồi thường bảo hiểm hàng hải (Claim): Tại sao biên bản giám định ROROC tại cảng là chứng từ quyết định tiền đền bù của bạn?",
        "source": "Quy tắc Bảo hiểm Hàng hải ICC 2009 (LMA/IUA) & Bộ luật Hàng hải Việt Nam 2015", "status": "Chờ duyệt"
    },
    {
        "stt": 32, "day": "[W37] Thứ Năm (17/09/2026)", "platform": "YouTube", "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến", "format": "Video Học XNK 5 phút Tập 7",
        "opt_a": "Học XNK 5 Phút: 5 Bước Đăng Ký Mã Số REX Trên Cổng Thông Tin EU Để Hưởng Thuế Ưu Đãi EVFTA 0% Tránh Phạt Truy Thu Tiền Tỷ",
        "opt_b": "Tự phát hành Lời văn khai báo xuất xứ trên hóa đơn thương mại (Invoice Declaration) chuẩn EVFTA: Không có mã REX thì bị bác thuế ngay lập tức!",
        "opt_c": "Mã số REX là gì? Hướng dẫn doanh nghiệp nông sản, dệt may hoàn thiện hồ sơ đăng ký REX trực tuyến qua Cổng Một cửa Quốc gia.",
        "source": "Hiệp định EVFTA, Nghị định 111/2020/NĐ-CP, TT 11/2020/TT-BCT & Quy chế REX (EU)", "status": "Chờ duyệt"
    },
    {
        "stt": 33, "day": "[W37] Thứ Sáu (18/09/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế", "format": "Bài tranh luận + Infographic",
        "opt_a": "Dính 1 cuộn chỉ hay vài chiếc cúc áo ngoại khối là bay sạch ưu đãi thuế CPTPP 0%: Dn dệt may đã biết cách vận dụng 'Phao cứu sinh' De Minimis 10% chưa?",
        "opt_b": "Cách tính tỷ lệ De Minimis chuẩn xác từng gram nguyên liệu: Công thức vàng giúp dn xuất khẩu không bị hải quan đối tác bác C/O.",
        "opt_c": "Nguyên tắc xuất xứ 'Từ sợi trở đi' (Yarn Forward) vs De Minimis: Rào cản kỹ thuật khắt khe hay động lực để dn chủ động nguồn cung trong nước?",
        "source": "Nghị định 31/2018/NĐ-CP về Xuất xứ hàng hóa, Hiệp định CPTPP (Chương 3) & TT 03/2019/TT-BCT", "status": "Chờ duyệt"
    },
    {
        "stt": 34, "day": "[W37] Thứ Bảy (19/09/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Vứt Hàng Xuống Biển Cứu Tàu: Luật Biển Đảo Rhodes 800 TCN & Chế Định 'Tổn Thất Chung' Buộc Mọi Chủ Hàng Phải Gánh Nợ Chung Như Thế Nào?",
        "opt_b": "Tại sao container của bạn không hề hỏng nhưng Hãng tàu vẫn bắt ký quỹ 50.000 USD mới cho lấy hàng? Giải mã bí ẩn General Average.",
        "opt_c": "Từ quy tắc ném hàng cứu tàu của thương gia Hy Lạp cổ đại đến Quy tắc York-Antwerp hiện đại: Bài học sống còn về quản trị rủi ro hàng hải.",
        "source": "Lịch sử Luật Hàng hải (Lex Rhodia 800 TCN), Quy tắc York-Antwerp 2016 & Bộ luật Hàng hải VN 2015", "status": "Chờ duyệt"
    },
    {
        "stt": 35, "day": "[W37] Chủ Nhật (20/09/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Problem-First", "format": "Bài Case Study + Ảnh",
        "opt_a": "Mất trắng 500 tấn thép trị giá 8 tỷ đồng vì cái bẫy 'SBLC kèm điều kiện xác nhận bên thứ ba': Bài học xương máu về thanh toán quốc tế cho dn SME!",
        "opt_b": "Checklist 5 điều khoản gài bẫy trong Thư tín dụng Dự phòng SBLC và L/C Thương mại mà người xuất khẩu tuyệt đối không được ký.",
        "opt_c": "Phân biệt rạch ròi Standby L/C (ISP98) vs Commercial L/C (UCP600): Đừng để đối tác ngoại 'ru ngủ' bằng những cam kết ngân hàng vô giá trị.",
        "source": "Quy tắc Thực hành Quốc tế ISP98 (ICC 590), URDG 758 & Án lệ tranh chấp SBLC", "status": "Chờ duyệt"
    }
]

# Update Sheet 1 Title
ws1["A1"] = "LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33, W34, W35, W36 & W37 CỘNG DỒN)"

# Check existing rows
current_max = ws1.max_row
existing_stts = [ws1.cell(row=r, column=1).value for r in range(5, current_max + 1)]

if 29 not in existing_stts:
    for idx, item in enumerate(w37_data):
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
    print(f"SUCCESS: Appended W37! New max row: {ws1.max_row}")
else:
    print("W37 already exists in Sheet 1!")
