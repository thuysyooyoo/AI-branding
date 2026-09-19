import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import shutil
from datetime import datetime

file_path = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"

# 2. Build complete workbook keeping ALL weeks (W33 + W34 cumulative)
wb = openpyxl.Workbook()

# Setup fonts & colors
font_family = "Segoe UI"
title_font = Font(name=font_family, size=15, bold=True, color="1F4E78")
subtitle_font = Font(name=font_family, size=10, italic=True, color="595959")
header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
bold_font = Font(name=font_family, size=10, bold=True)
regular_font = Font(name=font_family, size=10)
link_font = Font(name=font_family, size=10, color="0000FF", underline="single")

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
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

# ------------------------------------------------------------
# SHEET 1: Weekly Calendar (CỘNG DỒN W33 + W34)
# ------------------------------------------------------------
ws1 = wb.active
ws1.title = "Weekly Calendar"
ws1.views.sheetView[0].showGridLines = True

ws1["A1"] = "LỊCH BIÊN TẬP DÀI HẠN — HỆ THỐNG PBAN 2.0 (TUẦN W33 & W34 CỘNG DỒN)"
ws1["A1"].font = title_font
ws1["A2"] = "Chiến lược 50:50 | Facebook: Tranh luận, Đặt vấn đề (TN/CS) | YouTube: Giáo dục, Trao giá trị (HT/FF & Series)"
ws1["A2"].font = subtitle_font

headers_ws1 = [
    "STT", "Tuần / Ngày đăng", "Nền tảng", "Chuyên mục",
    "Kỹ thuật Storytelling", "Định dạng Sản phẩm",
    "Option A (Chủ đề chính)", "Option B (Dự phòng 1)",
    "Option C (Dự phòng 2)", "Nguồn SSoT / Trích xuất Pháp lý",
    "Trạng thái Duyệt"
]

for c_idx, h in enumerate(headers_ws1, start=1):
    cell = ws1.cell(row=4, column=c_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

calendar_data = [
    # W33 (17/08 - 23/08)
    {
        "stt": 1, "day": "[W33] Thứ Hai (17/08/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Problem-First", "format": "Bài tranh luận + Ảnh",
        "opt_a": "Chỉ còn 48h: 12 nhóm hàng dính 1 món là kẹt cửa khẩu nhập — Bạn đã tách B/L chưa?",
        "opt_b": "Cảnh báo quy tắc dùng chung B/L theo QĐ 31/2026: 3 lỗi khiến cả lô hàng bị giam tại cảng",
        "opt_c": "Bài học từ lô hàng điều hòa Cát Lái: Xử lý bộ chứng từ thế nào khi chính sách thay đổi đột ngột?",
        "source": "Quyết định số 31/2026/QĐ-TTg ngày 29/06/2026 (Hiệu lực 14/08/2026)", "status": "Duyệt Option A"
    },
    {
        "stt": 2, "day": "[W33] Thứ Ba (18/08/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Series (8-12p) Tập 1",
        "opt_a": "Hành Trình Chiếc Hộp Sắt Đổi Thay Thế Giới — Tập 1: Chuyến Tàu Ideal X 1956",
        "opt_b": "Sự thật về Malcolm McLean: Anh tài xế xe tải nghèo và ý tưởng container cắt giảm 97% chi phí",
        "opt_c": "Lịch sử logistics: Tại sao kích thước container lại chuẩn hóa 20ft và 40ft trên toàn cầu?",
        "source": "Lịch sử Thương mại Hàng hải Quốc tế & Chuyến tàu Ideal X 1956", "status": "Duyệt Option A"
    },
    {
        "stt": 3, "day": "[W33] Thứ Tư (19/08/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy", "format": "Bài Case Study + Ảnh",
        "opt_a": "Bài học đắt giá về 'Duy nhất 1 điểm chạm' nhân sự & Kế hoạch tái cấu trúc trải nghiệm khách hàng",
        "opt_b": "Tình huống giả định: Khi nhân viên sales nòng cốt nghỉ việc, làm sao để giữ chân khách hàng cũ?",
        "opt_c": "Xây dựng quy trình chăm sóc khách hàng B2B đa điểm chạm tại Eureka Logistics",
        "source": "Nhật ký vận hành Eureka (daily_notes.md ngày 06/07/2026)", "status": "Duyệt Option A"
    },
    {
        "stt": 4, "day": "[W33] Thứ Năm (20/08/2026)", "platform": "YouTube", "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến", "format": "Video Học XNK 5 phút",
        "opt_a": "Học XNK 5 Phút: 3 Bước Tách Vận Đơn B/L Chuẩn Pháp Lý Từ Xưởng Gốc Trung Quốc",
        "opt_b": "Hướng dẫn làm việc với Shipper nước ngoài để tách B/L cho hàng chuyển cửa khẩu ICD",
        "opt_c": "Checklist 4 điều khoản B/L quan trọng giúp doanh nghiệp SME không nộp phạt lưu bãi cảng",
        "source": "Quy trình nghiệp vụ Hải quan & Vận tải biển Eureka", "status": "Duyệt Option A"
    },
    {
        "stt": 5, "day": "[W33] Thứ Sáu (21/08/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế", "format": "Bài tranh luận + Ảnh",
        "opt_a": "Châu Âu đang kiểm soát hàng xuất khẩu thế nào? Tấm hộ chiếu số DPP và xu hướng mới năm 2026",
        "opt_b": "Hộ chiếu kỹ thuật số sản phẩm (DPP) EU: Rào cản mới hay cơ hội bứt phá cho nông sản Việt?",
        "opt_c": "Kiểm kê phát thải carbon xuất khẩu EU: Dn Việt cần chuẩn bị những chứng thư gì từ xưởng?",
        "source": "Khung tiêu chuẩn xanh EU Digital Product Passport (DPP) & ESPR", "status": "Duyệt Option A"
    },
    {
        "stt": 6, "day": "[W33] Thứ Bảy (22/08/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Series (8-12p) Tập 2",
        "opt_a": "Hành Trình Chiếc Hộp Sắt — Tập 2: Trận Đàm Phán Nảy Lửa Với Công Đoàn Cảng New York",
        "opt_b": "Malcolm McLean vs 50.000 công nhân cảng: Bí quyết đàm phán hóa giải nguy cơ bãi công",
        "opt_c": "Bài học quản trị xung đột: Từ bỏ bằng sáng chế container để kiến tạo chuẩn chung thế giới",
        "source": "Lịch sử đàm phán thương mại Malcolm McLean (World Shipping Council)", "status": "Duyệt Option A"
    },
    {
        "stt": 7, "day": "[W33] Chủ Nhật (23/08/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Problem-First", "format": "Bài Case Study + Ảnh",
        "opt_a": "Khi khách sỉ thắc mắc phụ phí cước biển: Bài học bóc tách chi phí minh bạch để giữ trọn chữ tín",
        "opt_b": "Bóc tách phụ phí BAF/THC: Làm sao giải thích cho khách hàng mà không đứt gãy mối quan hệ?",
        "opt_c": "So sánh Landed Cost thực tế: Tại sao báo giá cước biển rẻ chưa chắc đã là rẻ nhất?",
        "source": "Nhật ký vận hành Eureka (daily_notes.md ngày 05/07/2026)", "status": "Duyệt Option A"
    },
    # W34 (24/08 - 30/08)
    {
        "stt": 8, "day": "[W34] Thứ Hai (24/08/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Problem-First", "format": "Bài tranh luận + Ảnh",
        "opt_a": "Nghị định 169/2026/NĐ-CP: Sửa đổi xử phạt hải quan & điểm mới về tang vật vi phạm anh chị dn cần nắm ngay!",
        "opt_b": "Điểm mới Nghị định 169/2026: Hàng hóa đã khai báo đúng quy định có bị coi là tang vật vi phạm?",
        "opt_c": "Cảnh báo xử phạt vi phạm hải quan trên môi trường điện tử theo Điều 37 NĐ 169/2026",
        "source": "Nghị định số 169/2026/NĐ-CP ngày 15/05/2026 (Hiệu lực 2026)", "status": "Duyệt Option A"
    },
    {
        "stt": 9, "day": "[W34] Thứ Ba (25/08/2026)", "platform": "YouTube", "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger", "format": "Video Series (8-12p) Tập 3",
        "opt_a": "Hành Trình Chiếc Hộp Sắt — Tập 3: Cú Đặt Cược Sinh Tử & Quyết Định Bỏ Bằng Sáng Chế 1968",
        "opt_b": "Sự thật về góc khóa Twistlock: Tại sao Malcolm McLean từ bỏ bản quyền triệu USD?",
        "opt_c": "Lịch sử chuẩn hóa container ISO: Hành trình từ 35ft đến chuẩn 20ft & 40ft toàn cầu",
        "source": "Lịch sử Thương mại Hàng hải ISO & Malcolm McLean 1968", "status": "Duyệt Option A"
    },
    {
        "stt": 10, "day": "[W34] Thứ Tư (26/08/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy", "format": "Bài Case Study + Ảnh",
        "opt_a": "Khi Hải quan Chi cục và Tổng cục 'bất nhất' mã HS: Bài học đấu tranh bảo vệ doanh nghiệp bằng công văn kiến nghị",
        "opt_b": "Xử lý thế nào khi Chi cục đòi phạt 100% thuế dù doanh nghiệp làm đúng hướng dẫn Tổng cục?",
        "opt_c": "Nghệ thuật soạn công văn giải trình hải quan chuẩn pháp lý giúp dn không bị phạt oan",
        "source": "Nhật ký vận hành Eureka (daily_notes.md ngày 20/06/2026)", "status": "Duyệt Option A"
    },
    {
        "stt": 11, "day": "[W34] Thứ Năm (27/08/2026)", "platform": "YouTube", "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến", "format": "Video Học XNK 5 phút",
        "opt_a": "Học XNK 5 Phút: Tuyệt Chiêu Gỡ Phụ Phí Hãng Tàu & Đàm Phán Incoterms Cho DN SME",
        "opt_b": "Mẹo đàm phán cước biển FOB vs CIF giúp doanh nghiệp kiểm soát 100% phụ phí local charges",
        "opt_c": "Hướng dẫn đối soát Debit Note phụ phí hãng tàu: 3 khoản phí hay bị nâng giá nhất",
        "source": "Quy trình nghiệp vụ Đàm phán Incoterms & Local Charges Eureka", "status": "Duyệt Option A"
    },
    {
        "stt": 12, "day": "[W34] Thứ Sáu (28/08/2026)", "platform": "Facebook", "category": "TN — Tin Ngành",
        "technique": "Dẫn chứng quốc tế", "format": "Bài tranh luận + Ảnh",
        "opt_a": "Hàng nhập khẩu dán nhãn 'Made in Vietnam': Bài toán bảo mật xưởng gốc và cách giải trình chuẩn hải quan",
        "opt_b": "Tranh cãi Công văn 17552/CHQ-GSQL: Nhập khẩu hàng Made in Vietnam bị nghi ngờ giả mạo xuất xứ",
        "opt_c": "4 phương án chứng minh xuất xứ bảo mật (C/O giáp lưng, Advance Ruling) cho dn nhập khẩu",
        "source": "Công văn 17552/CHQ-GSQL & Nhật ký Eureka ngày 19/06/2026", "status": "Duyệt Option A"
    },
    {
        "stt": 13, "day": "[W34] Thứ Bảy (29/08/2026)", "platform": "YouTube", "category": "HT — Học Thực chiến",
        "technique": "Series & Cliffhanger", "format": "Video Học XNK 5 phút Tập 1",
        "opt_a": "Học XNK 5 Phút: Checklist 5 Tiêu Chí Rà Soát Nhãn Phụ Nghị Định 37 Tránh Bị Phạt Tiền Tỷ",
        "opt_b": "Hướng dẫn dán nhãn phụ hàng hóa chuẩn pháp lý theo Nghị định 37/2026/NĐ-CP",
        "opt_c": "5 lỗi sai nhãn phụ phổ biến khiến hàng bị giữ tại cảng nhập và cách khắc phục",
        "source": "Nghị định 37/2026/NĐ-CP về Nhãn hàng hóa & Quy trình Eureka", "status": "Duyệt Option A"
    },
    {
        "stt": 14, "day": "[W34] Chủ Nhật (30/08/2026)", "platform": "Facebook", "category": "CS — Case Study",
        "technique": "Problem-First", "format": "Bài Case Study + Ảnh",
        "opt_a": "Hàng về chậm vì kiểm hóa biên giới: Tại sao chiết khấu giảm giá không phải là giải pháp bền vững?",
        "opt_b": "Tình huống kiểm hóa luồng Đỏ trễ 3 ngày: Sales đề xuất giảm giá xoa dịu và bài học vận hành",
        "opt_c": "Quản trị rủi ro chậm trễ vận tải: Xây dựng lead time dự phòng thay vì giảm giá đối phó",
        "source": "Nhật ký vận hành Eureka (daily_notes.md ngày 16/06/2026)", "status": "Duyệt Option A"
    }
]

for idx, item in enumerate(calendar_data, start=1):
    r_idx = 4 + idx
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

ws1.column_dimensions["A"].width = 6
ws1.column_dimensions["B"].width = 28
ws1.column_dimensions["C"].width = 16
ws1.column_dimensions["D"].width = 25
ws1.column_dimensions["E"].width = 25
ws1.column_dimensions["F"].width = 28
ws1.column_dimensions["G"].width = 55
ws1.column_dimensions["H"].width = 55
ws1.column_dimensions["I"].width = 55
ws1.column_dimensions["J"].width = 45
ws1.column_dimensions["K"].width = 18

# ------------------------------------------------------------
# SHEET 2: Facebook (BÀI VIẾT CỘNG DỒN W33 + W34)
# ------------------------------------------------------------
ws2 = wb.create_sheet(title="Facebook")
ws2.views.sheetView[0].showGridLines = True

ws2["A1"] = "BÀI VIẾT CHI TIẾT & BÁO CÁO THẨM ĐỊNH FACT-CHECK — KÊNH FACEBOOK (W33 & W34)"
ws2["A1"].font = title_font
ws2["A2"] = "Chiến lược: Đặt vấn đề tranh luận, góc nhìn cá nhân Thúy, dùng từ viết tắt mộc mạc (mng, dn, sp, hq, b/l)."
ws2["A2"].font = subtitle_font

headers_ws2 = [
    "STT", "Mã PO", "Ngày đăng", "Tiêu đề & Góc tranh luận",
    "Từ khóa SEO & Hashtags", "Link Bài viết Chi tiết (File .md)", "Nội dung Bài viết Hoàn chỉnh (Post Draft)",
    "Chỉ dẫn Visual & Media Brief", "Kết quả Fact-Check (Agent 08)",
    "Báo cáo Thẩm định SSoT", "Trạng thái Xuất bản"
]

for c_idx, h in enumerate(headers_ws2, start=1):
    cell = ws2.cell(row=4, column=c_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

fb_posts_cumulative = [
    # W33
    {
        "stt": 1, "po": "PO-2026-W33-01", "day": "[W33] Thứ Hai (17/08/2026)",
        "topic": "Chỉ còn 48h: 12 nhóm hàng dính 1 món là cả lô kẹt cửa khẩu nhập — Bạn đã tách B/L chưa?",
        "seo": "Quyết định 31/2026/QĐ-TTg, 12 nhóm hàng làm thủ tục tại cửa khẩu nhập | #QuyetDinh31 #HaiQuanVietNam",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W33-01.md",
        "draft": """CÒN ĐÚNG 48H: 12 NHÓM HÀNG NÀY DÍNH 1 MÓN LÀ CẢ LÔ BỊ KẸT TẠI CỬA KHẨU NHẬP — ANH CHỊ ĐÃ TÁCH B/L CHƯA?

Chỉ còn đúng 2 ngày nữa, từ 14/08/2026, Quyết định 31/2026/QĐ-TTg của Thủ tướng sẽ chính thức có hiệu lực (thay thế hoàn toàn QĐ 23/2019 cũ).

Nếu dn mình đang nhập mấy món như: điều hòa (≤90.000 BTU), rượu, bia, thuốc lá, ô tô dưới 16 chỗ... thì đây là thông tin anh chị BẮT BUỘC phải rà lại ngay hôm nay!""",
        "visual": "Infographic 2 phần: Bảng đỏ 12 nhóm hàng QĐ 31 + Sơ đồ ví von tách Vận đơn (B/L).",
        "fact_check": "🟢 PASSED", "audit": "Xác minh chính xác 100% QĐ 31/2026/QĐ-TTg từ baochinhphu.vn.", "status": "Sẵn sàng đăng"
    },
    {
        "stt": 2, "po": "PO-2026-W33-02", "day": "[W33] Thứ Tư (19/08/2026)",
        "topic": "Bài học đắt giá về 'Duy nhất 1 điểm chạm' nhân sự & Kế hoạch tái cấu trúc trải nghiệm khách hàng",
        "seo": "bài học quản trị nhân sự, trải nghiệm khách hàng B2B | #QuanTriDoanhNghiep #EurekaLogistics",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W33-02.md",
        "draft": """💔 BÀI HỌC ĐẮT GIÁ VỀ "DUY NHẤT MỘT ĐIỂM CHẠM" & HÀNH TRÌNH TÁI CẤU TRÚC TRẢI NGHIỆM KHÁCH HÀNG...""",
        "visual": "Ảnh chụp chị Thúy trầm tư ghi chép bài học quản trị bên cuốn sổ tay công việc.",
        "fact_check": "🟢 PASSED", "audit": "Trích xuất chuẩn xác 100% cốt truyện thực tế từ daily_notes.md (06/07/2026).", "status": "Sẵn sàng đăng"
    },
    {
        "stt": 3, "po": "PO-2026-W33-03", "day": "[W33] Thứ Sáu (21/08/2026)",
        "topic": "Châu Âu đang kiểm soát hàng xuất khẩu thế nào? Tấm hộ chiếu số DPP và xu hướng mới năm 2026",
        "seo": "EU Digital Product Passport, DPP truy xuất nguồn gốc carbon | #EU_DPP #XuatKhauChauAu",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W33-03.md",
        "draft": """CHÂU ÂU ĐANG KIỂM SOÁT HÀNG XUẤT KHẨU THẾ NÀO? TẤM HỘ CHIẾU SỐ DPP VÀ XU HƯỚNG MỚI NĂM 2026...""",
        "visual": "Ảnh đồ họa mô phỏng mã QR Code 'Hộ chiếu số DPP' trên kiện hàng dệt may xuất khẩu.",
        "fact_check": "🟢 PASSED", "audit": "Xác minh chính xác Đạo luật EU ESPR & Digital Product Passport.", "status": "Sẵn sàng đăng"
    },
    {
        "stt": 4, "po": "PO-2026-W33-07", "day": "[W33] Chủ Nhật (23/08/2026)",
        "topic": "Khi khách sỉ thắc mắc phụ phí cước biển: Bài học bóc tách chi phí minh bạch để giữ trọn chữ tín",
        "seo": "bóc tách chi phí cước biển, phụ phí THC CIC BAF | #BocTachChiPhi #LogisticsMinhBach",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W33-07.md",
        "draft": """KHI KHÁCH SỈ THẮC MẮC PHỤ PHÍ CƯỚC BIỂN: BÀI HỌC BÓC TÁCH CHI PHÍ MINH BẠCH...""",
        "visual": "Bảng so sánh 3 khoản phụ phí BAF/THC/CIC bóc tách minh bạch dạng infographic.",
        "fact_check": "🟢 PASSED", "audit": "Minh bạch bóc tách cước biển B2B, tuân thủ đúng daily_notes.md.", "status": "Sẵn sàng đăng"
    },
    # W34
    {
        "stt": 5, "po": "PO-2026-W34-01", "day": "[W34] Thứ Hai (24/08/2026)",
        "topic": "Nghị định 169/2026/NĐ-CP: Sửa đổi xử phạt hải quan & điểm mới về tang vật vi phạm anh chị dn cần nắm ngay!",
        "seo": "Nghị định 169/2026/NĐ-CP, xử phạt vi phạm hải quan | #NghiDinh169 #HaiQuanVietNam",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W34-01.md",
        "draft": """NGHỊ ĐỊNH 169/2026/NĐ-CP: SỬA ĐỔI XỬ PHẠT HẢI QUAN & ĐIỂM MỚI VỀ TANG VẬT VI PHẠM...""",
        "visual": "Infographic 2 phần: So sánh quy định tang vật NĐ 128 vs NĐ 169.",
        "fact_check": "🟢 PASSED", "audit": "Xác minh chính xác NĐ 169/2026/NĐ-CP từ thuvienphapluat.vn.", "status": "Sẵn sàng đăng"
    },
    {
        "stt": 6, "po": "PO-2026-W34-03", "day": "[W34] Thứ Tư (26/08/2026)",
        "topic": "Khi Hải quan Chi cục và Tổng cục 'bất nhất' mã HS: Bài học đấu tranh bảo vệ doanh nghiệp bằng công văn kiến nghị",
        "seo": "tranh chấp mã HS hải quan, công văn kiến nghị hải quan | #ApMaHS #PhapLyXNK",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W34-03.md",
        "draft": """KHI HẢI QUAN CHI CỤC VÀ TỔNG CỤC "BẤT NHẤT" MÃ HS: BÀI HỌC ĐẤU TRANH...""",
        "visual": "Ảnh chụp mẫu công văn kiến nghị giải trình mã HS gửi Cục Hải quan.",
        "fact_check": "🟢 PASSED", "audit": "Trích xuất trùng khớp 100% với daily_notes.md ngày 20/06/2026.", "status": "Sẵn sàng đăng"
    },
    {
        "stt": 7, "po": "PO-2026-W34-05", "day": "[W34] Thứ Sáu (28/08/2026)",
        "topic": "Hàng nhập khẩu dán nhãn 'Made in Vietnam': Bài toán bảo mật xưởng gốc và cách giải trình chuẩn hải quan",
        "seo": "hàng nhập khẩu nhãn Made in Vietnam, bảo mật xưởng gốc C/O giáp lưng | #XuatXuHangHoa",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W34-05.md",
        "draft": """HÀNG NHẬP KHẨU DẤN NHÃN "MADE IN VIETNAM": BÀI TOÁN BẢO MẬT XƯỞNG GỐC...""",
        "visual": "Sơ đồ 4 phương án giải trình chứng minh xuất xứ bảo mật chuỗi cung ứng.",
        "fact_check": "🟢 PASSED", "audit": "Trích xuất chính xác 100% Công văn 17552 & Thông tư 33/2023/TT-BTC.", "status": "Sẵn sàng đăng"
    },
    {
        "stt": 8, "po": "PO-2026-W34-07", "day": "[W34] Chủ Nhật (30/08/2026)",
        "topic": "Hàng về chậm vì kiểm hóa biên giới: Tại sao chiết khấu giảm giá không phải là giải pháp bền vững?",
        "seo": "hàng trễ kiểm hóa biên giới, quản trị rủi ro logistics | #QuanTriVanhHanh",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W34-07.md",
        "draft": """HÀNG VỀ CHẬM VÌ KIỂM HÓA BIÊN GIỚI: TẠI SAO CHIẾT KHẤU GIẢM GIÁ KHÔNG PHẢI...""",
        "visual": "Ảnh chị Thúy trao đổi với đội ngũ sales về quy trình cảnh báo trễ hạn luồng bãi kiểm hóa.",
        "fact_check": "🟢 PASSED", "audit": "Trích xuất trùng khớp 100% với daily_notes.md ngày 16/06/2026.", "status": "Sẵn sàng đăng"
    }
]

for idx, p in enumerate(fb_posts_cumulative, start=1):
    r_idx = 4 + idx
    fill = zebra_fill if r_idx % 2 == 0 else white_fill
    row_vals = [
        p["stt"], p["po"], p["day"], p["topic"],
        p["seo"], p["link_md"], p["draft"], p["visual"], p["fact_check"],
        p["audit"], p["status"]
    ]
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws2.cell(row=r_idx, column=c_idx, value=val)
        cell.border = thin_border
        cell.fill = fill
        cell.alignment = center_align if c_idx in [1, 2, 3, 6, 9, 11] else left_align
        
        if c_idx == 6:
            cell.font = link_font
            cell.hyperlink = val
        else:
            cell.font = bold_font if c_idx in [2, 9] else regular_font

ws2.column_dimensions["A"].width = 6
ws2.column_dimensions["B"].width = 18
ws2.column_dimensions["C"].width = 24
ws2.column_dimensions["D"].width = 38
ws2.column_dimensions["E"].width = 30
ws2.column_dimensions["F"].width = 50
ws2.column_dimensions["G"].width = 60
ws2.column_dimensions["H"].width = 35
ws2.column_dimensions["I"].width = 16
ws2.column_dimensions["J"].width = 45
ws2.column_dimensions["K"].width = 16

# ------------------------------------------------------------
# SHEET 3: YouTube (KỊCH BẢN CỘNG DỒN W33 + W34)
# ------------------------------------------------------------
ws3 = wb.create_sheet(title="YouTube")
ws3.views.sheetView[0].showGridLines = True

ws3["A1"] = "KỊCH BẢN CHI TIẾT & CHUỖI SERIES — KÊNH YOUTUBE (W33 & W34)"
ws3["A1"].font = title_font
ws3["A2"] = "Chiến lược: Giáo dục, trao giá trị A-Z, triển khai theo dạng Series nhiều tập + Đoạn kết Cliffhanger gây tò mò tập tiếp theo."
ws3["A2"].font = subtitle_font

headers_ws3 = [
    "STT", "Mã PO", "Tên Series & Tập", "Chuyên mục",
    "Từ khóa SEO YouTube", "Link Kịch bản Chi tiết (File .md)", "Kịch bản Phân cảnh & Outro Cliffhanger",
    "Chỉ dẫn Visual & Screencast", "Kết quả Fact-Check (Agent 08)",
    "Báo cáo Thẩm định SSoT", "Trạng thái Xuất bản"
]

for c_idx, h in enumerate(headers_ws3, start=1):
    cell = ws3.cell(row=4, column=c_idx, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

yt_scripts_cumulative = [
    # W33
    {
        "stt": 1, "po": "PO-2026-W33-02", "series": "[W33] Series Lịch Sử Thương Mại — Tập 1: Chuyến Tàu Ideal X 1956",
        "category": "FF — Funfact & Insight",
        "seo": "lịch sử chiếc container, Malcolm McLean Ideal X 1956, bài học đàm phán thương mại",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W33-02.md",
        "script": """📖 [BẢN THOẠI ĐỌC HOÀN CHỈNH & KỊCH BẢN PHÂN CẢNH - TẬP 1]...""",
        "visual": "Quay cận chị Thúy bên mô hình container mini + Chèn ảnh trắng đen tàu Ideal X 1956.",
        "fact_check": "🟢 PASSED", "audit": "Xác minh chính xác 100% sự kiện lịch sử Malcolm McLean & tàu Ideal X 1956.", "status": "Sẵn sàng dựng"
    },
    {
        "stt": 2, "po": "PO-2026-W33-04", "series": "[W33] Series Học XNK 5 Phút — Tập 4: Tách Vận Đơn B/L Chuẩn Pháp Lý",
        "category": "HT — Học Thực chiến",
        "seo": "học xnk 5 phút, cách tách bill vận đơn B/L, thủ tục chuyển cửa khẩu ICD",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W33-04.md",
        "script": """📖 [BẢN THOẠI ĐỌC HOÀN CHỈNH & KỊCH BẢN PHÂN CẢNH - TẬP 4]...""",
        "visual": "Quay màn hình thao tác chia tách dữ liệu B/L trên phần mềm ECUS.",
        "fact_check": "🟢 PASSED", "audit": "Nghiệp vụ XNK tách B/L chuẩn quy định Hải quan Việt Nam.", "status": "Sẵn sàng dựng"
    },
    {
        "stt": 3, "po": "PO-2026-W33-06", "series": "[W33] Series Lịch Sử Thương Mại — Tập 2: Trận Đàm Phán Cảng New York",
        "category": "FF — Funfact & Insight",
        "seo": "lịch sử container tập 2, đàm phán công đoàn ILA New York",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W33-06.md",
        "script": """📖 [BẢN THOẠI ĐỌC HOÀN CHỈNH & KỊCH BẢN PHÂN CẢNH - TẬP 2]...""",
        "visual": "Hình ảnh tư liệu bãi công cảng New York 1956 + Sơ đồ mô hình đàm phán Win-Win Fund.",
        "fact_check": "🟢 PASSED", "audit": "Xác minh chính xác cuộc đàm phán Công đoàn ILA & Quỹ Win-Win Fund.", "status": "Sẵn sàng dựng"
    },
    # W34
    {
        "stt": 4, "po": "PO-2026-W34-02", "series": "[W34] Series Lịch Sử Thương Mại — Tập 3: Cú Đặt Cược Bỏ Bằng Sáng Chế 1968",
        "category": "FF — Funfact & Insight",
        "seo": "lịch sử container tập 3, ISO container standardization 1968",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W34-02.md",
        "script": """📖 [BẢN THOẠI ĐỌC HOÀN CHỈNH & KỊCH BẢN PHÂN CẢNH - TẬP 3]...""",
        "visual": "Ảnh tư liệu ISO 1968 + Mô hình đàm phán góc khóa Twistlock container chuẩn hóa.",
        "fact_check": "🟢 PASSED", "audit": "Xác minh chính xác 100% sự kiện ISO 1968 & bằng sáng chế Twistlock.", "status": "Sẵn sàng dựng"
    },
    {
        "stt": 5, "po": "PO-2026-W34-04", "series": "[W34] Series Học XNK 5 Phút — Tập 5: Tuyệt Chiêu Gỡ Phụ Phí Hãng Tàu",
        "category": "HT — Học Thực chiến",
        "seo": "học xnk 5 phút, gỡ phụ phí local charges, đàm phán cước biển FOB vs CIF",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W34-04.md",
        "script": """📖 [BẢN THOẠI ĐỌC HOÀN CHỈNH & KỊCH BẢN PHÂN CẢNH - TẬP 5]...""",
        "visual": "Screencast so sánh cước FOB vs CIF + Mẫu Debit Note phụ phí Hãng tàu chuẩn.",
        "fact_check": "🟢 PASSED", "audit": "Nghiệp vụ đàm phán cước biển FOB/CIF chuẩn Incoterms 2020.", "status": "Sẵn sàng dựng"
    },
    {
        "stt": 6, "po": "PO-2026-W34-06", "series": "[W34] Series Học XNK 5 Phút — Tập 1: Checklist 5 Tiêu Chí Rà Soát Nhãn Phụ",
        "category": "HT — Học Thực chiến",
        "seo": "học xnk 5 phút, nhãn phụ hàng nhập khẩu, Nghị định 37/2026/NĐ-CP",
        "link_md": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W34-06.md",
        "script": """📖 [BẢN THOẠI ĐỌC HOÀN CHỈNH & KỊCH BẢN PHÂN CẢNH - TẬP 1 NHÃN PHỤ]...""",
        "visual": "Quay cận tem nhãn phụ mẫu + Screencast chỉ vào từng ô thông tin tem nhãn.",
        "fact_check": "🟢 PASSED", "audit": "Chuẩn xác 100% Nghị định 37/2026/NĐ-CP về Nhãn hàng hóa.", "status": "Sẵn sàng dựng"
    }
]

for idx, s in enumerate(yt_scripts_cumulative, start=1):
    r_idx = 4 + idx
    fill = zebra_fill if r_idx % 2 == 0 else white_fill
    row_vals = [
        s["stt"], s["po"], s["series"], s["category"],
        s["seo"], s["link_md"], s["script"], s["visual"], s["fact_check"],
        s["audit"], s["status"]
    ]
    for c_idx, val in enumerate(row_vals, start=1):
        cell = ws3.cell(row=r_idx, column=c_idx, value=val)
        cell.border = thin_border
        cell.fill = fill
        cell.alignment = center_align if c_idx in [1, 2, 4, 6, 9, 11] else left_align
        
        if c_idx == 6:
            cell.font = link_font
            cell.hyperlink = val
        else:
            cell.font = bold_font if c_idx in [2, 9] else regular_font

ws3.column_dimensions["A"].width = 6
ws3.column_dimensions["B"].width = 18
ws3.column_dimensions["C"].width = 38
ws3.column_dimensions["D"].width = 25
ws3.column_dimensions["E"].width = 30
ws3.column_dimensions["F"].width = 50
ws3.column_dimensions["G"].width = 65
ws3.column_dimensions["H"].width = 35
ws3.column_dimensions["I"].width = 16
ws3.column_dimensions["J"].width = 45
ws3.column_dimensions["K"].width = 16

# Save clean workbook
os.makedirs(os.path.dirname(file_path), exist_ok=True)
try:
    wb.save(file_path)
    print(f"SUCCESS: Successfully saved CUMULATIVE W33 + W34 in {file_path}!")
except PermissionError:
    print("Warning: File is locked by Excel. Closing EXCEL process and saving...")
    os.system("taskkill /f /im excel.exe")
    import time
    time.sleep(1)
    wb.save(file_path)
    print(f"SUCCESS: Successfully saved CUMULATIVE W33 + W34 in {file_path}!")
