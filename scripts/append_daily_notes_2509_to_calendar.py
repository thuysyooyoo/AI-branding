# -*- coding: utf-8 -*-
"""
Cập nhật 6 bài viết Facebook và 3 kịch bản YouTube mới từ Daily Notes 25/09/2026 vào Weekly_Content_Calendar.xlsx:
Sheet 2 (Facebook): Thêm STT 54, 55, 56, 57, 58, 59 (PO-2026-W39-HOT16 -> HOT21)
Sheet 3 (YouTube): Thêm STT 34, 35, 36 (PO-2026-W39-HOT16, HOT17, HOT20)
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

fb_hot16 = read_draft("DRAFT-FB-HOT-16-GAY-SEAL-CHUYEN-CANG.md")
fb_hot17 = read_draft("DRAFT-FB-HOT-17-TT121-GIA-HAN-KTCL.md")
fb_hot18 = read_draft("DRAFT-FB-HOT-18-VK010-VAT-HANH-LA-SAY.md")
fb_hot19 = read_draft("DRAFT-FB-HOT-19-THUE-XE-TU-LAI-7710.md")
fb_hot20 = read_draft("DRAFT-FB-HOT-20-HOA-DON-CONG-TY-MA.md")
fb_hot21 = read_draft("DRAFT-FB-HOT-21-HKD-AFFILIATE-TMDT.md")

yt_hot16 = read_draft("DRAFT-YT-HOT-16-GAY-SEAL-CHUYEN-CANG.md")
yt_hot17 = read_draft("DRAFT-YT-HOT-17-TT121-GIA-HAN-KTCL.md")
yt_hot20 = read_draft("DRAFT-YT-HOT-20-HOA-DON-CONG-TY-MA.md")

NEW_FB_ROWS = [
    {
        "stt": 54,
        "po": "PO-2026-W39-HOT16",
        "date": "[W39] Thứ Sáu (25/09/2026)",
        "title": "Cont bị gãy seal hãng tàu khi chuyển cảng Cát Lái sang Hiệp Phước, cảng đã bấm seal mới: Khi truyền tờ khai lấy hàng phải điền số seal nào để không bị giữ cont tại cổng?",
        "seo": "#ChuyenCangHaiQuan #ThaySealContainer #CatLai #TanCangHiepPhuoc #NghiepVuLogistics #ThongQuanChinhNgach #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-16-GAY-SEAL-CHUYEN-CANG.md",
        "content": fb_hot16,
        "visual": "Infographic chụp cận cảnh 2 loại chì niêm phong: Chì cáp hãng tàu bị đứt và chì cối Tân Cảng mới bấm; kèm quy trình 3 bước truyền số seal mới và đính kèm Biên bản bất thường lên v5 Hải quan.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Điều 34 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC), quy chế giám sát cảng Cát Lái - Hiệp Phước và scratch/daily_notes.md (25/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 55,
        "po": "PO-2026-W39-HOT17",
        "date": "[W39] Thứ Bảy (26/09/2026)",
        "title": "Hàng được mang về bảo quản theo Thông tư 121 nhưng phòng lab chưa có kết quả, Chi cục TĐC lại từ chối ký gia hạn: Làm sao để doanh nghiệp không bị phạt và tước quyền mang hàng trong 6 tháng?",
        "seo": "#ThongTu121 #KiemTraChatLuong #BaoQuanHangHoa #HaiQuanVietNam #ChiCucTieuChuanDoLuong #NghiepVuHaiQuan #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-17-TT121-GIA-HAN-KTCL.md",
        "content": fb_hot17,
        "visual": "Sơ đồ bóc tách xung đột quy trình giữa Hải quan (Điều 32 TT 39) và Chi cục TĐC; Checklist 3 bước: Thư xác nhận tiến độ từ phòng lab + Công văn giải trình Hải quan trước ngày 25 + Mời kiểm tra thực tế kho.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 6 Điều 32 Thông tư 38/2015 (sửa đổi TT 39/2018), Thông tư 121, Nghị định 128/2020/NĐ-CP và case study Hải quan (25/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 56,
        "po": "PO-2026-W39-HOT18",
        "date": "[W39] Chủ Nhật (27/09/2026)",
        "title": "Nhập khẩu hành lá sấy nhiệt cắt vụn từ Trung Quốc (HS 0712.90.90): Áp mã miễn thuế VAT 'VK010' theo Nghị định 181/2025/NĐ-CP được không hay dính bẫy truy thu thuế 5%?",
        "seo": "#NghiDinh181 #ThueVATNongSan #MaVK010 #HaiQuanVietNam #NhapKhauNongSan #MaHS0712 #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-18-VK010-VAT-HANH-LA-SAY.md",
        "content": fb_hot18,
        "visual": "Bảng so sánh kỹ thuật: Sơ chế thông thường (phơi sấy vật lý không gia vị -> VK010 không thuế VAT) vs Chế biến công nghiệp (sấy dầu, tẩm ướp phụ gia -> Chịu thuế VAT 5%/10%); mẫu bảng quy trình sản xuất nhà máy TQ.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 1 Điều 4 Nghị định 181/2025/NĐ-CP (Luật Thuế GTGT 2025), Biểu thuế XNK và scratch/daily_notes.md (25/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 57,
        "po": "PO-2026-W39-HOT19",
        "date": "[W40] Thứ Hai (28/09/2026)",
        "title": "Kinh doanh cho thuê xe ô tô tự lái: Có phải xin Giấy phép kinh doanh vận tải không? Cẩm nang đăng ký chuẩn mã ngành 7710 và tránh rủi ro trách nhiệm liên đới khi khách gây tai nạn!",
        "seo": "#ChoThueXeTuLai #MaNganh7710 #NghiDinh10 #PhapLyKinhDoanh #QuanTriRuiRo #TuVanDoanhNghiep #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-19-THUE-XE-TU-LAI-7710.md",
        "content": fb_hot19,
        "visual": "Checklist pháp lý 3 điểm: Đăng ký chuẩn mã ngành 7710 (không nhầm sang 4931/4932); Điều khoản hợp đồng chuyển giao trách nhiệm dân sự/hình sự; Quy định cấm khách dùng xe tự lái chở khách có thu tiền.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Nghị định 10/2020/NĐ-CP, Luật Đường bộ 2024, Quyết định 27/2018/QĐ-TTg và case study Diễn đàn Thuế chuyên sâu (25/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 58,
        "po": "PO-2026-W39-HOT20",
        "date": "[W40] Thứ Ba (29/09/2026)",
        "title": "'Lúc mua nó là công ty bình thường, sao giờ bảo nó là ma rồi bắt tôi đóng phạt?': Cách tự vệ trước bão thanh tra hóa đơn đầu vào và bảo vệ chi phí hợp lệ!",
        "seo": "#HoaDonCongTyMa #TrangThai06 #ThanhTraThue #ChiPhiHopLy #QuanTriTaiChinh #LuatQuanLyThue #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-20-HOA-DON-CONG-TY-MA.md",
        "content": fb_hot20,
        "visual": "Bộ 4 chứng từ thép bảo vệ chi phí: Hợp đồng kinh tế + Biên bản giao nhận/biển số xe tải + Thẻ kho nhập xuất + Ủy nhiệm chi ngân hàng; Màn hình tra cứu trạng thái MST trên tracuunnt.gdt.gov.vn.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Luật Quản lý thuế số 38/2019/QH14, Công văn 1798/TCT-TTKT, Nghị định 123/2020/NĐ-CP và case study Diễn đàn Thuế chuyên sâu (25/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 59,
        "po": "PO-2026-W39-HOT21",
        "date": "[W40] Thứ Tư (30/09/2026)",
        "title": "Vừa là chủ hộ kinh doanh cửa hàng vừa làm Affiliate Shopee/TikTok: Nên khai thuế riêng hay gộp doanh thu để không bị truy thu phạt và tối ưu thuế suất từ 10% xuống 7%?",
        "seo": "#ThueAffiliate #ThueThuongMaiDienTu #HoKinhDoanh #ShopeeAffiliate #TikTokShop #eTaxMobile #TuVanThue #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-21-HKD-AFFILIATE-TMDT.md",
        "content": fb_hot21,
        "visual": "Sơ đồ tối ưu hóa thuế TMĐT: Bổ sung ngành nghề quảng cáo vào HKD -> Cập nhật thông tin thuế lên Shopee/TikTok -> Kê khai tách 2 dòng trên Form 01/TKN-CNKD (Hàng hóa 1.5% + Dịch vụ Affiliate 7% thay vì 10% cá nhân).",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Thông tư 40/2021/TT-BTC (Mẫu 01/TKN-CNKD), Nghị định 125/2020/NĐ-CP và case study Diễn đàn Thuế chuyên sâu (25/09/2026).",
        "status": "Sẵn sàng đăng"
    }
]

NEW_YT_ROWS = [
    {
        "stt": 34,
        "po": "PO-2026-W39-HOT16",
        "series": "[W39] Series Học XNK 5 Phút: Container Bị Gãy Seal Khi Chuyển Cảng Cát Lái - Hiệp Phước (Khai Seal Nào?)",
        "category": "HT — Học Thực chiến",
        "seo": "#ChuyenCangHaiQuan #ThaySealContainer #CatLai #TanCangHiepPhuoc #NghiepVuLogistics #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-16-GAY-SEAL-CHUYEN-CANG.md",
        "content": yt_hot16,
        "visual": "Quay tại bãi container Cát Lái và văn phòng Eureka; Cầm trên tay chì cáp hãng tàu bị đứt và chì cối Tân Cảng; Screencast ECUS5-VNACCS thao tác khai báo số seal mới và đính kèm v5; Outro gợi mở bẫy 30 ngày TT 121.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Điều 34 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC) và scratch/daily_notes.md (25/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 35,
        "po": "PO-2026-W39-HOT17",
        "series": "[W39] Series Học XNK 5 Phút: Bẫy 30 Ngày Thông Tư 121 & Cẩm Nang Gia Hạn Kiểm Tra Chất Lượng",
        "category": "HT — Học Thực chiến",
        "seo": "#ThongTu121 #KiemTraChatLuong #BaoQuanHangHoa #HaiQuanVietNam #ChiCucTieuChuanDoLuong #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-17-TT121-GIA-HAN-KTCL.md",
        "content": yt_hot17,
        "visual": "Quay tại kho hàng nhập khẩu; Màn hình hiển thị đồng hồ đếm ngược 30 ngày; Screencast Điều 32 TT 39 và Thông tư 121; Trình chiếu mẫu Công văn giải trình và Giấy tiếp nhận mẫu Quatest 3; Outro gợi mở bẫy hóa đơn công ty ma.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 6 Điều 32 Thông tư 38/2015 (sửa đổi TT 39/2018), Thông tư 121 và scratch/daily_notes.md (25/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 36,
        "po": "PO-2026-W39-HOT20",
        "series": "[W40] Series Quản Trị & Pháp Lý: Bẫy Hóa Đơn Công Ty Ma & 4 Bằng Chứng Thép Tự Vệ Trước Thanh Tra Thuế",
        "category": "CS — Case Study",
        "seo": "#HoaDonCongTyMa #TrangThai06 #ThanhTraThue #ChiPhiHopLy #QuanTriTaiChinh #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-20-HOA-DON-CONG-TY-MA.md",
        "content": yt_hot20,
        "visual": "Quay tại phòng làm việc Eureka; Trực quan hóa sơ đồ dòng tiền vs dòng hàng hóa; Trình chiếu bộ 4 chứng từ thép (Hợp đồng, Phiếu cân xe tải, Thẻ kho, UNC); Hướng dẫn tra cứu 2 phút trên cổng Tổng cục Thuế; Outro TMĐT.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Luật Quản lý thuế số 38/2019/QH14, Công văn 1798/TCT-TTKT và scratch/daily_notes.md (25/09/2026).",
        "status": "Sẵn sàng dựng"
    }
]

def append_to_calendar():
    print(f"Đang tải workbook từ: {EXCEL_PATH}...")
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
