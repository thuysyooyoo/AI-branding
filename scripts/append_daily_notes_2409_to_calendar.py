# -*- coding: utf-8 -*-
"""
Cập nhật 4 bài viết và 2 kịch bản YouTube mới từ Daily Notes 24/09/2026 vào Weekly_Content_Calendar.xlsx:
Sheet 2 (Facebook): Thêm STT 50, 51, 52, 53 (PO-2026-W39-HOT12 -> HOT15)
Sheet 3 (YouTube): Thêm STT 32, 33 (PO-2026-W39-HOT12 & HOT13)
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
ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\.13.Dữ liệu - bài viết - Hình ảnh - video\TỔNG HỢP DỮ LIỆU QUÉT GROUP FB"

# Đọc nội dung từ các file draft đã tạo
def read_draft(fname):
    p = os.path.join(r"d:\AI AGENT THUY\AI branding\scratch\production_orders", fname)
    with open(p, "r", encoding="utf-8") as f:
        return f.read().strip()

fb_hot12 = read_draft("DRAFT-FB-HOT-12-PHAN-DAME-BOC-PHOT-ERP.md")
fb_hot13 = read_draft("DRAFT-FB-HOT-13-CV22082-BAO-BI-ATTP.md")
fb_hot14 = read_draft("DRAFT-FB-HOT-14-CHUYEN-DIA-CHI-QUYET-TOAN-THUE.md")
fb_hot15 = read_draft("DRAFT-FB-HOT-15-XAC-NHAN-THUE-MAU-24.md")

yt_hot12 = read_draft("DRAFT-YT-HOT-12-PHAN-DAME-BOC-PHOT-ERP.md")
yt_hot13 = read_draft("DRAFT-YT-HOT-13-CV22082-BAO-BI-ATTP.md")

NEW_FB_ROWS = [
    {
        "stt": 50,
        "po": "PO-2026-W39-HOT12",
        "date": "[W39] Thứ Năm (24/09/2026)",
        "title": "Bị khách bóc phốt tìm đối tác mới vì 'hàng về lâu lắc': Tại sao Thúy không đôi co mà dùng dữ liệu ERP để 'phản dame', và câu chuyện 'tâm thư bênh vực' chấn động của một khách hàng ruột!",
        "seo": "#XuLyKhungHoang #ERPLogistics #CustomerAdvocacy #NhapKhauChinhNgach #VanHoaPhungSu #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-12-PHAN-DAME-BOC-PHOT-ERP.md",
        "content": fb_hot12,
        "visual": "Bộ 5 ảnh chụp màn hình: Ảnh bài bóc phốt trên nhóm 200k mem; Ảnh comment cam kết của Thúy; Ảnh phiếu gửi hàng nội địa TQ không mã KH; Ảnh tra cứu hệ thống ERP Eureka; Ảnh tâm thư bảo vệ của khách hàng ruột.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% nhật ký scratch/daily_notes.md (24/09/2026), log quản lý hệ thống ERP Eureka và phiếu gửi hàng nội địa Trung Quốc.",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 51,
        "po": "PO-2026-W39-HOT13",
        "date": "[W39] Thứ Sáu (25/09/2026)",
        "title": "Công văn 22082/CHQ-GSQL & 1913/ATTP-KN vừa ban hành: Doanh nghiệp nhập khẩu túi vô trùng, bao bì tiếp xúc thực phẩm đừng lầm tưởng 'Tự công bố' là đương nhiên được miễn kiểm tra ATTP!",
        "seo": "#CongVan22082 #ThongTu27 #ThongTu15 #KiemTraATTP #BaoBiThucPham #AsepticBag #HaiQuanVietNam #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-13-CV22082-BAO-BI-ATTP.md",
        "content": fb_hot13,
        "visual": "Infographic bóc tách ranh giới pháp lý giữa 'Tự công bố sản phẩm' (theo TT 27/2026 & NĐ 15) và 'Kiểm tra nhà nước về ATTP' (theo TT 15/2024/TT-BYT & CV 22082/CHQ-GSQL). Kèm dấu mộc cảnh báo kẹt cảng.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Công văn 22082/CHQ-GSQL (23/09/2026), Công văn 1913/ATTP-KN (21/09/2026), Thông tư 27/2026/TT-BYT, Thông tư 15/2024/TT-BYT và Nghị định 37/2026/NĐ-CP.",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 52,
        "po": "PO-2026-W39-HOT14",
        "date": "[W39] Thứ Bảy (26/09/2026)",
        "title": "Chuyển địa chỉ kinh doanh khác phường / khác quận: Cẩm nang tránh bẫy 'thanh tra quyết toán thuế tại trụ sở' cho doanh nghiệp chưa quyết toán trên 2 năm!",
        "seo": "#ChuyenDiaChiKinhDoanh #QuyetToanThue #ThanhTraThue #Mau08MST #SoSachKeToan #TuVanThueChinhNgach #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-14-CHUYEN-DIA-CHI-QUYET-TOAN-THUE.md",
        "content": fb_hot14,
        "visual": "Sơ đồ quy trình 4 bước chuẩn bị hồ sơ dọn sạch sổ sách kế toán, rà soát hóa đơn F1-F2 và đối soát công nợ eTax trước khi nộp hồ sơ chuyển địa điểm kinh doanh theo Thông tư 105/2020/TT-BTC.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Luật Quản lý thuế số 38/2019/QH14, Thông tư 105/2020/TT-BTC, Nghị định 125/2020/NĐ-CP và case study Diễn đàn Thuế chuyên sâu (24/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 53,
        "po": "PO-2026-W39-HOT15",
        "date": "[W39] Chủ Nhật (27/09/2026)",
        "title": "Quy trình xin Giấy xác nhận hoàn thành nghĩa vụ thuế (Mẫu 24) nộp Sở Kế hoạch & Đầu tư: Bẫy 'thụ động chờ đợi' làm treo hồ sơ giải thể / thay đổi hàng tháng trời!",
        "seo": "#XacNhanNghiaVuThue #Mau24 #SoKeHoachDauTu #GiaiTheDoanhNghiep #eTax #ThuTucThue #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-15-XAC-NHAN-THUE-MAU-24.md",
        "content": fb_hot15,
        "visual": "Checklist 3 bước chốt sổ thuế điện tử: Tự tra cứu sạch nợ lẻ trên eTax -> Nộp đơn Mẫu 24 lấy số công văn -> Ký biên bản đối chiếu TMS trực tiếp với cán bộ thuế để lấy kết quả trong 3 ngày.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Thông tư 80/2021/TT-BTC (Mẫu 24), Luật Doanh nghiệp 2020, quy chế một cửa ngành Thuế và case study Diễn đàn Thuế chuyên sâu (24/09/2026).",
        "status": "Sẵn sàng đăng"
    }
]

NEW_YT_ROWS = [
    {
        "stt": 32,
        "po": "PO-2026-W39-HOT12",
        "series": "[W39] Hot Topic Thực Chiến — Series Quản Trị & Personal Branding: Nghệ Thuật 'Phản Dame' Khi Bị Bóc Phốt & ERP Eureka",
        "category": "CS — Case Study",
        "seo": "#XuLyKhungHoang #ERPLogistics #CustomerAdvocacy #NhapKhauChinhNgach #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-12-PHAN-DAME-BOC-PHOT-ERP.md",
        "content": yt_hot12,
        "visual": "Quay tại phòng làm việc Eureka; Screencast hệ thống phần mềm quản lý ERP Eureka tra cứu log đơn hàng; Đồ họa 3 bước xử lý khủng hoảng; Trích đoạn tâm thư khách hàng ruột; Outro Cliffhanger gợi mở Công văn 22082.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% nhật ký scratch/daily_notes.md (24/09/2026), log hệ thống ERP Eureka và phiếu gửi hàng nội địa Trung Quốc.",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 33,
        "po": "PO-2026-W39-HOT13",
        "series": "[W39] Series Học XNK 5 Phút: Cảnh Báo Nghẽn Cảng Bao Bì Thực Phẩm & Công Văn 22082/CHQ-GSQL",
        "category": "HT — Học Thực chiến",
        "seo": "#CongVan22082 #ThongTu27 #ThongTu15 #KiemTraATTP #BaoBiThucPham #AsepticBag #HaiQuanVietNam #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-13-CV22082-BAO-BI-ATTP.md",
        "content": yt_hot13,
        "visual": "Quay tại studio Eureka; Bối cảnh kho cảng và hình ảnh thực tế túi vô trùng Aseptic bag HS 3923.21.99; Screencast văn bản Công văn 22082 & CV 1913/ATTP-KN; Hướng dẫn thao tác NSW; Outro Cliffhanger đổi địa chỉ kinh doanh.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Công văn 22082/CHQ-GSQL (23/09/2026), Công văn 1913/ATTP-KN, Thông tư 27/2026/TT-BYT, Thông tư 15/2024/TT-BYT và Nghị định 37/2026/NĐ-CP.",
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
