# -*- coding: utf-8 -*-
"""
Script: append_daily_notes_2609_to_calendar.py
Mục tiêu: Cập nhật 4 bài viết Facebook và 3 kịch bản YouTube mới từ Daily Notes 26/09/2026 vào Weekly_Content_Calendar.xlsx:
- Sheet 2 (Facebook): Thêm STT 64, 65, 66, 67 (PO-2026-W39-HOT22 -> HOT25)
- Sheet 3 (YouTube): Thêm STT 40, 41, 42 (PO-2026-W39-HOT22, HOT23, HOT24)
Tuân thủ tuyệt đối GEMINI.md (Atomic Save, Zero Backup File, Chuẩn Typography Arial 10pt)
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

EXCEL_PATH = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"

def read_draft(fname):
    p = os.path.join(r"d:\AI AGENT THUY\AI branding\scratch\production_orders", fname)
    with open(p, "r", encoding="utf-8") as f:
        return f.read().strip()

fb_hot22 = read_draft("DRAFT-FB-HOT-22-LUAT-HAI-QUAN-11-2026.md")
fb_hot23 = read_draft("DRAFT-FB-HOT-23-DREWRY-BLANK-SAILINGS-GOLDEN-WEEK.md")
fb_hot24 = read_draft("DRAFT-FB-HOT-24-CBAM-MARKUP-10-PERCENT.md")
fb_hot25 = read_draft("DRAFT-FB-HOT-25-XUAT-KHAU-HALAL-LOGISTICS.md")

yt_hot22 = read_draft("DRAFT-YT-HOT-22-LUAT-HAI-QUAN-11-2026.md")
yt_hot23 = read_draft("DRAFT-YT-HOT-23-DREWRY-BLANK-SAILINGS-GOLDEN-WEEK.md")
yt_hot24 = read_draft("DRAFT-YT-HOT-24-CBAM-MARKUP-10-PERCENT.md")

NEW_FB_ROWS = [
    {
        "stt": 64,
        "po": "PO-2026-W39-HOT22",
        "date": "[W39] Thứ Bảy (26/09/2026)",
        "title": "Quốc hội thông qua Luật số 11/2026/QH16 sửa đổi Luật Hải quan & Lộ trình thay thế Nghị định 08, Thông tư 38/39: Chuyển dịch sang Hải quan số & Tự động phân luồng AI — Dn nên mừng hay lo?",
        "seo": "#LuatHaiQuan2026 #Luat11_2026_QH16 #HaiQuanSo #HaiQuanThongMinh #KiemTraSauThongQuan #ThongTu85 #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-22-LUAT-HAI-QUAN-11-2026.md",
        "content": fb_hot22,
        "visual": "Infographic so sánh 3 trụ cột đổi mới của Luật 11/2026/QH16: Phân luồng AI tự động + Liên thông dữ liệu lớn + Chuyển trọng tâm sang kiểm tra sau thông quan 5 năm; Sơ đồ quy trình đối soát hồ sơ kỹ thuật catalogue theo Thông tư 85.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Luật số 11/2026/QH16, Luật Hải quan 54/2014/QH13, Thông tư 85/2026/TT-BTC, QĐ 1434/QĐ-CHQ và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 65,
        "po": "PO-2026-W39-HOT23",
        "date": "[W39] Chủ Nhật (27/09/2026)",
        "title": "Chỉ số Drewry WCI đạt 4.468 USD/FEU & Áp lực 15 chuyến Blank Sailings trước thềm Golden Week (01-07/10): Tuyến Shanghai - Los Angeles neo 7.838 USD, rủi ro rớt cont diện rộng mùa cao điểm!",
        "seo": "#DrewryWCI #BlankSailings #GoldenWeek2026 #CuocBienTranspacific #RollCont #LogisticsVietNam #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-23-DREWRY-BLANK-SAILINGS-GOLDEN-WEEK.md",
        "content": fb_hot23,
        "visual": "Biểu đồ nến thị trường cước biển Drewry WCI tuần 24/09; Bản đồ luồng hàng Transpacific với 15 chuyến tàu bị hủy; Bảng checklist 3 giải pháp: Booking 21 ngày + No-Roll Guarantee + Gom cont LCL chính ngạch Eureka.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Báo cáo Drewry World Container Index tuần 24/09/2026, dữ liệu World Shipping Council và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 66,
        "po": "PO-2026-W39-HOT24",
        "date": "[W40] Thứ Hai (28/09/2026)",
        "title": "Quy định CBAM sửa đổi (Regulation 2026/1740) chính thức áp mức phạt Mark-up 10% cho giá trị phát thải mặc định: Doanh nghiệp xuất khẩu sắt thép, nhôm sang EU đối mặt bão chi phí nếu không tự đo khí thải!",
        "seo": "#CBAM #Regulation2026_1740 #KiemKeCarbon #GiaTriMacDinh #XuatKhauSatThep #GreenLogistics #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-24-CBAM-MARKUP-10-PERCENT.md",
        "content": fb_hot24,
        "visual": "Bảng so sánh chi phí chứng chỉ carbon EU ETS: Áp dữ liệu thực tế nhà máy vs Áp giá trị mặc định bị phạt thêm 10% Mark-up; Sơ đồ 3 tầng phát thải Scope 1, Scope 2 và Scope 3 Logistics chuẩn XML của Eureka.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Commission Implementing Regulation (EU) 2026/1740, Regulation (EU) 2023/956, Tiêu chuẩn ISO 14064-1 và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 67,
        "po": "PO-2026-W39-HOT25",
        "date": "[W40] Thứ Ba (29/09/2026)",
        "title": "Mở khóa thị trường Halal 2.000 tỷ USD: Hội nghị xúc tiến Bộ Công Thương (23/09) & Cạm bẫy 'nhiễm chéo' trong lưu kho, vận tải container khiến hàng bị trả về tại cảng đến!",
        "seo": "#XuatKhauHalal #ThiTruongHoiGiao #HalalLogistics #BoCongThuong #NongSanVietNam #NghiepVuLogistics #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-25-XUAT-KHAU-HALAL-LOGISTICS.md",
        "content": fb_hot25,
        "visual": "Infographic quy chuẩn Halal Logistics: Chuỗi cung ứng cách ly tuyệt đối từ xưởng đến mạn tàu; Bản đồ phân bố 3 chứng chỉ Halal quyền lực nhất (JAKIM Malaysia, BPJPH Indonesia, GSO Vùng Vịnh); Mẫu biên bản khử trùng vỏ cont.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Hội nghị xúc tiến thương mại Bộ Công Thương ngày 23/09/2026, tiêu chuẩn JAKIM / BPJPH / GSO Halal và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng đăng"
    }
]

NEW_YT_ROWS = [
    {
        "stt": 40,
        "po": "PO-2026-W39-HOT22",
        "series": "[W39] Series Quản Trị & Pháp Lý: Quốc Hội Thông Qua Luật Sửa Đổi Luật Hải Quan 11/2026/QH16 & Lộ Trình Thay Thế Nghị Định 08, Thông Tư 38/39",
        "category": "TN — Tin Ngành",
        "seo": "#LuatHaiQuan2026 #Luat11_2026_QH16 #HaiQuanSo #HaiQuanThongMinh #ThongTu85 #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-22-LUAT-HAI-QUAN-11-2026.md",
        "content": yt_hot22,
        "visual": "Quay tại studio Eureka; Trên tay cầm cuốn Luật Hải quan và màn hình phía sau hiển thị sơ đồ phân luồng AI; Screencast Cổng TTĐT Quốc hội và Báo Hải quan Online; Outro gợi mở bão cước Drewry WCI & Blank Sailings.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Luật số 11/2026/QH16, Luật Hải quan 54/2014/QH13, Thông tư 85/2026/TT-BTC, Quyết định 1434/QĐ-CHQ và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 41,
        "po": "PO-2026-W39-HOT23",
        "series": "[W39] Hot Topic Thực Chiến — Series Quản Trị Vận Tải: Bão Cước Drewry WCI 4.468 USD & Chiêu Trò 15 Chuyến Blank Sailings Trước Golden Week (Làm Sao Để Không Rớt Cont?)",
        "category": "TN — Tin Ngành",
        "seo": "#DrewryWCI #BlankSailings #GoldenWeek2026 #CuocBienTranspacific #RollCont #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-23-DREWRY-BLANK-SAILINGS-GOLDEN-WEEK.md",
        "content": yt_hot23,
        "visual": "Quay tại phòng làm việc Eureka bên đồ thị nến Drewry WCI; Animation mô phỏng 15 chuyến tàu mẹ bị hủy; Trình chiếu mẫu hợp đồng Booking Confirmation có điều khoản No-Roll Guarantee và giải pháp gom cont LCL Eureka; Outro CBAM.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Báo cáo Drewry World Container Index tuần 24/09/2026, dữ liệu World Shipping Council và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 42,
        "po": "PO-2026-W39-HOT24",
        "series": "[W39] Series Học XNK 5 Phút: Bóc Tách Quy Định CBAM Sửa Đổi (EU) 2026/1740: Mức Phạt Mark-Up 10% & 4 Bước Tự Đo Đạc Khí Thải Scope 1 & 2 Cứu Lợi Nhuận Doanh Nghiệp",
        "category": "HT — Học Thực chiến",
        "seo": "#CBAM #Regulation2026_1740 #KiemKeCarbon #GiaTriMacDinh #XuatKhauSatThep #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-24-CBAM-MARKUP-10-PERCENT.md",
        "content": yt_hot24,
        "visual": "Quay tại studio với bảng đồ họa Cơ chế CBAM Châu Âu; Cầm trên tay tấm mẫu chứng thư phát thải carbon và cuộn thép xuất khẩu mô hình; Screencast Regulation 2026/1740; Bảng tính phát thải Excel chuẩn XML; Outro Halal.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Commission Implementing Regulation (EU) 2026/1740, Regulation (EU) 2023/956, Tiêu chuẩn ISO 14064-1 và nhật ký scratch/daily_notes.md (26/09/2026).",
        "status": "Sẵn sàng dựng"
    }
]

def append_to_calendar():
    print(f"Đang nạp workbook từ: {EXCEL_PATH}...")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    
    font_name = "Arial"
    regular_font = Font(name=font_name, size=10)
    bold_font = Font(name=font_name, size=10, bold=True)
    link_font = Font(name=font_name, size=10, color="0000FF", underline="single")
    passed_font = Font(name=font_name, size=10, bold=True, color="006100")
    
    thin_border_side = Side(style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    # 1. Cập nhật Sheet 2: Facebook
    ws2 = wb["Facebook"]
    start_row_fb = ws2.max_row + 1
    print(f"Bắt đầu ghi Sheet 2 (Facebook) từ dòng {start_row_fb}...")
    
    for idx, item in enumerate(NEW_FB_ROWS):
        r = start_row_fb + idx
        ws2.cell(row=r, column=1, value=item["stt"])
        ws2.cell(row=r, column=2, value=item["po"])
        ws2.cell(row=r, column=3, value=item["date"])
        ws2.cell(row=r, column=4, value=item["title"])
        ws2.cell(row=r, column=5, value=item["seo"])
        
        c6 = ws2.cell(row=r, column=6, value=item["link"])
        c6.hyperlink = item["link"]
        c6.font = link_font
        
        ws2.cell(row=r, column=7, value=item["content"])
        ws2.cell(row=r, column=8, value=item["visual"])
        
        c9 = ws2.cell(row=r, column=9, value=item["factcheck"])
        c9.font = passed_font
        
        ws2.cell(row=r, column=10, value=item["ssot"])
        ws2.cell(row=r, column=11, value=item["status"])
        
        for c in range(1, 12):
            cell = ws2.cell(row=r, column=c)
            cell.border = thin_border
            if c not in [6, 9]:
                cell.font = regular_font
            if c in [1, 2, 3, 9, 11]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
            if c == 1:
                cell.number_format = "#,##0"
                
        ws2.row_dimensions[r].height = 65

    # 2. Cập nhật Sheet 3: YouTube
    ws3 = wb["YouTube"]
    start_row_yt = ws3.max_row + 1
    print(f"Bắt đầu ghi Sheet 3 (YouTube) từ dòng {start_row_yt}...")
    
    for idx, item in enumerate(NEW_YT_ROWS):
        r = start_row_yt + idx
        ws3.cell(row=r, column=1, value=item["stt"])
        ws3.cell(row=r, column=2, value=item["po"])
        ws3.cell(row=r, column=3, value=item["series"])
        ws3.cell(row=r, column=4, value=item["category"])
        ws3.cell(row=r, column=5, value=item["seo"])
        
        c6 = ws3.cell(row=r, column=6, value=item["link"])
        c6.hyperlink = item["link"]
        c6.font = link_font
        
        ws3.cell(row=r, column=7, value=item["content"])
        ws3.cell(row=r, column=8, value=item["visual"])
        
        c9 = ws3.cell(row=r, column=9, value=item["factcheck"])
        c9.font = passed_font
        
        ws3.cell(row=r, column=10, value=item["ssot"])
        ws3.cell(row=r, column=11, value=item["status"])
        
        for c in range(1, 12):
            cell = ws3.cell(row=r, column=c)
            cell.border = thin_border
            if c not in [6, 9]:
                cell.font = regular_font
            if c in [1, 2, 4, 9, 11]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
            if c == 1:
                cell.number_format = "#,##0"
                
        ws3.row_dimensions[r].height = 65

    # 3. Ghi tệp nguyên tử (Atomic save) tuân thủ GEMINI.md Rule 5 & 10
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    
    if os.path.exists(EXCEL_PATH):
        os.replace(tmp_path, EXCEL_PATH)
    else:
        os.rename(tmp_path, EXCEL_PATH)
    print(f"THÀNH CÔNG: Đã ghi đè nguyên tử (atomic replace) tệp chính {EXCEL_PATH}")

if __name__ == "__main__":
    append_to_calendar()
