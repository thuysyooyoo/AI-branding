# -*- coding: utf-8 -*-
"""
Script append Week 43 (STT 71 - 77) to Weekly_Content_Calendar.xlsx
Tuân thủ tuyệt đối Rule 10: Không tạo file backup, dùng atomic_save (.tmp.xlsx -> os.replace)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
import os
import json

file_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
json_path = r"d:\AI AGENT THUY\AI branding\scratch\w43_data.json"

w43_data = [
    {
        "stt": 71,
        "day": "[W43] Thứ Hai (26/10/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Thông tư 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP: Phân cấp thẩm quyền cấp C/O về địa phương — Cảnh báo bẫy ngộ nhận 'thủ tục ở tỉnh sẽ dễ hơn' khiến hàng loạt lô hàng xuất khẩu kẹt cứng vì tiêu chí xác minh xuất xứ bị siết chặt!",
        "opt_b": "Checklist 4 bước chuẩn bị hồ sơ xác minh cơ sở sản xuất và giải trình định mức nguyên liệu (BOM) theo quy chế phân cấp mới để được cấp C/O Form D, AK, VK trong 24 giờ.",
        "opt_c": "Khi thẩm quyền cấp C/O chuyển giao về địa phương: Bài học quản trị rủi ro 'bất đồng quan điểm thẩm định' giữa hệ thống eCoSys và cán bộ kiểm tra thực địa.",
        "source": "Thông tư số 26/2026/TT-BCT ngày 15/07/2026 của Bộ Công Thương, Nghị định số 146/2025/NĐ-CP, Nghị định số 31/2018/NĐ-CP về xuất xứ hàng hóa, Quy chế vận hành Hệ thống eCoSys & scratch/daily_notes.md ngày 09/10/2026",
        "status": "Chờ duyệt"
    },
    {
        "stt": 72,
        "day": "[W43] Thứ Ba (27/10/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Khủng Hoảng Hạn Hán Kênh Đào Panama: Cuộc Đấu Giá 4 Triệu USD Để Giành Một Slot Vượt Kênh, Nghịch Lý 'Thừa Tàu Thiếu Nước' & Bài Học Quản Trị Rủi Ro Tuyến Tuyến Hàng Hải Cho Doanh Nghiệp Xuất Khẩu Sang Bờ Đông Nước Mỹ",
        "opt_b": "Bí mật kỹ thuật đằng sau hệ thống âu tàu nâng nước hồ Gatun: Tại sao một con kênh nhân tạo 110 năm tuổi lại có thể làm đảo lộn chuỗi cung ứng nông sản và dệt may toàn cầu?",
        "opt_c": "Từ khủng hoảng Panama 2024-2026 đến bài toán đa dạng hóa tuyến đường: Tại sao doanh nghiệp logistics Việt Nam phải học cách phân bổ rủi ro qua kênh Suez, Mũi Hảo Vọng và cầu cạn xuyên Mỹ (IPI)?",
        "source": "Báo cáo Thường niên Cơ quan Quản lý Kênh đào Panama (ACP 2024-2026), Dữ liệu Phân tích Thị trường Vận tải Drewry WCI & Freightos, Báo cáo Chuỗi Cung ứng Toàn cầu Lloyd's List & Quy tắc York-Antwerp",
        "status": "Chờ duyệt"
    },
    {
        "stt": 73,
        "day": "[W43] Thứ Tư (28/10/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Doanh nghiệp FDI nằm trong Khu chế xuất (KCX) mua hàng từ nội địa: Cạm bẫy nhập nhằng giữa 'Doanh nghiệp chế xuất (DNCX) hưởng thuế quan ưu đãi' và 'Doanh nghiệp FDI thông thường' — Sai một bước là bị Thuế truy thu và tắc toàn bộ hóa đơn VAT!",
        "opt_b": "Sơ đồ 3 bước xác định chuẩn xác: Khi nào doanh nghiệp FDI trong KCX bắt buộc phải mở tờ khai hải quan XNK tại chỗ, và khi nào áp dụng hóa đơn GTGT 10% như thương mại nội địa thông thường?",
        "opt_c": "Tâm sự người làm nghề tư vấn: Phía sau câu chuyện kế toán cãi nhau với cán bộ thuế là sự thiếu rạch ròi giữa 'Địa giới hành chính KCX' và 'Tư cách pháp nhân DNCX'.",
        "source": "Điều 74 & Điều 86 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Điều 26 Nghị định số 35/2022/NĐ-CP về quản lý KCN & KKT, Công văn số 22411/CHQ-GSQL của Cục Hải quan & scratch/daily_notes.md ngày 07/10/2026",
        "status": "Chờ duyệt"
    },
    {
        "stt": 74,
        "day": "[W43] Thứ Năm (29/10/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 13",
        "opt_a": "Học XNK 5 Phút Tập 13: 5 Bước Đăng Ký Xác Định Trước Mã Số Hàng Hóa (Advance Ruling) Để 'Khóa' Rủi Ro Thuế Quan & Không Bị Hải Quan Bác Mã HS Khi Thông Quan",
        "opt_b": "Checklist 4 tài liệu kỹ thuật bắt buộc phải chuẩn bị (Catalogue gốc, tài liệu kỹ thuật dịch công chứng, kết quả kiểm nghiệm xưởng) để hồ sơ Advance Ruling được Cục Hải quan chấp thuận trong 30 ngày.",
        "opt_c": "Mẹo thực chiến cho doanh nghiệp nhập khẩu máy móc, thiết bị công nghệ mới: Cách gửi mẫu phân tích và xây dựng văn bản giải trình tính chất kỹ thuật chuẩn xác theo Điều 24 Nghị định 08/2015/NĐ-CP.",
        "source": "Điều 28 Luật Hải quan số 54/2014/QH13, Điều 24 Nghị định số 08/2015/NĐ-CP (sửa đổi tại Nghị định 59/2018/NĐ-CP), Điều 7 Thông tư số 38/2015/TT-BTC, Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan",
        "status": "Chờ duyệt"
    },
    {
        "stt": 75,
        "day": "[W43] Thứ Sáu (30/10/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Thông báo 22555/TB-CHQ vừa ban hành: Cảnh báo bẫy áp mã HS thiết bị đeo thông minh (Smartwatch) — Nhìn chữ 'Đồng hồ' vội áp nhóm 91.02 thay vì nhóm 85.17 là dính ngay bẫy thuế và xử phạt vi phạm hành chính!",
        "opt_b": "Bảng đối chiếu 4 tiêu chí kỹ thuật quyết định phân loại thiết bị đeo tay theo Chú giải 3 Phần XVI: Phân định rạch ròi giữa 'Đồng hồ có chức năng phụ trợ' và 'Thiết bị truyền/nhận dữ liệu đa năng'.",
        "opt_c": "Từ case study Smartwatch Huawei Watch GT 6: Bài học đắt giá cho dân làm chứng từ công nghệ — Hiểu bản chất vi mạch và giao thức Bluetooth/Wi-Fi quan trọng hơn nhìn thiết kế bề ngoài!",
        "source": "Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan về kết quả xác định trước mã số, Danh mục hàng hóa XNK Việt Nam (Thông tư 31/2022/TT-BTC), 6 Quy tắc Tổng quát giải thích HS (GIR) & scratch/daily_notes.md ngày 07/10/2026",
        "status": "Chờ duyệt"
    },
    {
        "stt": 76,
        "day": "[W43] Thứ Bảy (31/10/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Đế Chế Cảng Biển Tự Động Hóa Maasvlakte 2 (Rotterdam): Cỗ Máy Logistics 'Không Một Bóng Người' Bốc Xếp, Vận Hành Bằng 100% Điện Gió & Cuộc Cách Mạng Xanh Đang Thay Đổi Cuộc Đua Cảng Biển Thế Giới 2026",
        "opt_b": "Bí mật sau đội xe AGV tự hành và cần cẩu STS điều khiển từ xa cách 50km: Làm sao Rotterdam giải bài toán tăng gấp đôi năng suất bốc dỡ mà giảm 100% phát thải carbon tại bến cảng?",
        "opt_c": "Từ kỳ quan tự động hóa Rotterdam đến chiến lược cảng biển Việt Nam: Bài học nào cho các siêu cảng nước sâu Cần Giờ, Cái Mép - Thị Vải và Lạch Huyện trên hành trình trở thành Hub trung chuyển quốc tế?",
        "source": "Báo cáo Thường niên Port of Rotterdam Authority 2025-2026, Dữ liệu Công nghệ Cảng biển Tự động hóa APM Terminals Rotterdam, Báo cáo Nghiên cứu Chuyển đổi Xanh Hàng hải của Tổ chức Hàng hải Quốc tế (IMO / World Bank) & Tài liệu Phân tích Cảng biển Eureka",
        "status": "Chờ duyệt"
    },
    {
        "stt": 77,
        "day": "[W43] Chủ Nhật (01/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan: Cạm bẫy giao dịch Xuất nhập khẩu tại chỗ khi thiếu 'Văn bản chỉ định giao nhận của thương nhân nước ngoài' — Hàng chục tờ khai bị đình chỉ thông quan và nguy cơ truy thu thuế GTGT 10%!",
        "opt_b": "Checklist 3 điều khoản sống còn bắt buộc phải rà soát trong Hợp đồng mua bán 3 bên và Thư chỉ định giao hàng (Delivery Instruction) để được Hải quan chấp nhận thông quan XNK tại chỗ mượt mà.",
        "opt_c": "Bài học xương máu từ một đợt thanh tra sau thông quan: Tại sao thiếu một dòng chỉ định trong hợp đồng ngoại thương lại biến một giao dịch XNK hợp pháp thành hành vi mua bán nội địa bị phạt nặng?",
        "source": "Điều 47a Luật Hải quan (được bổ sung bởi Luật số 90/2025/QH15), Điều 35 Nghị định số 08/2015/NĐ-CP (sửa đổi, bổ sung tại Nghị định 167/2025/NĐ-CP), Thông tư 121/2025/TT-BTC sửa đổi Điều 86 Thông tư 38/2015/TT-BTC, Công văn số 22411/CHQ-GSQL ngày 30/09/2026 của Cục Hải quan & scratch/daily_notes.md ngày 01/10/2026",
        "status": "Chờ duyệt"
    }
]

# 1. Save scratch/w43_data.json
os.makedirs(os.path.dirname(json_path), exist_ok=True)
with open(json_path, "w", encoding="utf-8") as f:
    json.dump(w43_data, f, ensure_ascii=False, indent=2)
print(f"SUCCESS: Saved JSON data to {json_path}")

# 2. Append to Weekly_Content_Calendar.xlsx
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

# Update Sheet 1 Title
ws1["A1"] = "LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W43 CỘNG DỒN)"

# Check existing rows
current_max = ws1.max_row
existing_stts = [ws1.cell(row=r, column=1).value for r in range(5, current_max + 1)]

if 71 not in existing_stts:
    for idx, item in enumerate(w43_data):
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

    # Atomic Save to avoid OneDrive sync lock and ensure NO BACKUP files
    tmp_path = file_path + ".tmp.xlsx"
    wb.save(tmp_path)
    os.replace(tmp_path, file_path)
    print(f"SUCCESS: Appended W43! New max row: {ws1.max_row}")
else:
    print("W43 already exists in Sheet 1!")
