# -*- coding: utf-8 -*-
"""
Script: append_daily_notes_0710_to_calendar.py
Mục tiêu:
1. Đồng bộ các bản thảo vào reports/plan1/facebook/ và reports/plan1/video_scripts/
2. Cập nhật 5 bài viết Facebook (HOT32 -> HOT36) và 4 kịch bản YouTube (HOT32 -> HOT35) vào Weekly_Content_Calendar.xlsx:
   - Sheet 2 (Facebook): Thêm STT 74, 75, 76, 77, 78
   - Sheet 3 (YouTube): Thêm STT 47, 48, 49, 50
3. Tuân thủ tuyệt đối GEMINI.md:
   - Atomic Save: ghi vào file tạm .tmp.xlsx rồi os.replace
   - Zero Backup File: không tạo bất kỳ file backup nào
   - Typography chuẩn Arial 10pt
"""

import os
import sys
import shutil
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

EXCEL_PATH = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"
PO_DIR = r"d:\AI AGENT THUY\AI branding\scratch\production_orders"
FB_PLAN_DIR = r"d:\AI AGENT THUY\AI branding\reports\plan1\facebook"
YT_PLAN_DIR = r"d:\AI AGENT THUY\AI branding\reports\plan1\video_scripts"

os.makedirs(FB_PLAN_DIR, exist_ok=True)
os.makedirs(YT_PLAN_DIR, exist_ok=True)

# 1. Đồng bộ sang reports/plan1
sync_fb_map = {
    "DRAFT-FB-HOT-32-CV22411-XNK-TAI-CHO.md": "FB-HOT-2026-32-CV22411-XNK-TAI-CHO.md",
    "DRAFT-FB-HOT-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md": "FB-HOT-2026-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md",
    "DRAFT-FB-HOT-34-MA-HS-SMARTWATCH-8517.md": "FB-HOT-2026-34-MA-HS-SMARTWATCH-8517.md",
    "DRAFT-FB-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md": "FB-HOT-2026-35-DN-FDI-TRONG-KCX-MUA-HANG.md",
    "DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md": "FB-HOT-2026-36-THUE-NHA-THAU-SOT-HOA-DON.md",
}

sync_yt_map = {
    "DRAFT-YT-HOT-32-CV22411-XNK-TAI-CHO.md": "YT-HOT-2026-32-CV22411-XNK-TAI-CHO.md",
    "DRAFT-YT-HOT-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md": "YT-HOT-2026-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md",
    "DRAFT-YT-HOT-34-MA-HS-SMARTWATCH-8517.md": "YT-HOT-2026-34-MA-HS-SMARTWATCH-8517.md",
    "DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md": "YT-HOT-2026-35-DN-FDI-TRONG-KCX-MUA-HANG.md",
}

print("Đang đồng bộ file sang reports/plan1/facebook và video_scripts...")
for src_name, dst_name in sync_fb_map.items():
    shutil.copy2(os.path.join(PO_DIR, src_name), os.path.join(FB_PLAN_DIR, dst_name))
    print(f" -> Đã sao chép FB: {dst_name}")

for src_name, dst_name in sync_yt_map.items():
    shutil.copy2(os.path.join(PO_DIR, src_name), os.path.join(YT_PLAN_DIR, dst_name))
    print(f" -> Đã sao chép YT: {dst_name}")

def read_draft(fname):
    p = os.path.join(PO_DIR, fname)
    with open(p, "r", encoding="utf-8") as f:
        return f.read().strip()

fb_hot32 = read_draft("DRAFT-FB-HOT-32-CV22411-XNK-TAI-CHO.md")
fb_hot33 = read_draft("DRAFT-FB-HOT-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md")
fb_hot34 = read_draft("DRAFT-FB-HOT-34-MA-HS-SMARTWATCH-8517.md")
fb_hot35 = read_draft("DRAFT-FB-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md")
fb_hot36 = read_draft("DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md")

yt_hot32 = read_draft("DRAFT-YT-HOT-32-CV22411-XNK-TAI-CHO.md")
yt_hot33 = read_draft("DRAFT-YT-HOT-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md")
yt_hot34 = read_draft("DRAFT-YT-HOT-34-MA-HS-SMARTWATCH-8517.md")
yt_hot35 = read_draft("DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md")

NEW_FB_ROWS = [
    {
        "stt": 74,
        "po": "PO-2026-W41-HOT32",
        "date": "[W41] Thứ Năm (01/10/2026)",
        "title": "Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan (sửa đổi): Thủ tục hải quan đối với hàng hóa xuất nhập khẩu tại chỗ chuyển biến nóng — Doanh nghiệp cần chuẩn bị hồ sơ gì?",
        "seo": "#XNKtaiCho #CongVan22411 #Dieu47aLuatHaiQuan #Luat90_2025 #ThuTucHaiQuan #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-32-CV22411-XNK-TAI-CHO.md",
        "content": fb_hot32,
        "visual": "Infographic so sánh quy trình XNK tại chỗ cũ (Điều 86 TT 38) vs quy trình mới theo Điều 47a Luật Hải quan (Luật 90/2025) & CV 22411/CHQ-GSQL; Sơ đồ 3 bên: Bên bán nội địa - Bên mua nội địa - Thương nhân nước ngoài; Bảng checklist 4 loại chứng từ cốt lõi.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Công văn 22411/CHQ-GSQL (29/09/2026), Luật Hải quan số 90/2025/QH15 (Điều 47a), Điều 86 Thông tư 38/2015/TT-BTC và nhật ký scratch/daily_notes.md (01/10/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 75,
        "po": "PO-2026-W41-HOT33",
        "date": "[W41] Thứ Ba (06/10/2026)",
        "title": "Tranh cãi Incoterms trong XNK tại chỗ: Bán chỉ định giao hàng tại kho đối tác thì dùng FCA hay EXW? Cước xe nội địa và phí bốc dỡ ai chịu?",
        "seo": "#Incoterms #FCA #EXW #XNKtaiCho #ChiPhiVanChuyen #HopDongThuongMai #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md",
        "content": fb_hot33,
        "visual": "Bảng ma trận so sánh EXW vs FCA trong giao dịch XNK tại chỗ: Điểm giao hàng, Chi phí cước xe nội địa, Trách nhiệm mở tờ khai xuất khẩu; Sơ đồ rủi ro phân định trách nhiệm khi phát sinh tổn thất hàng hóa trên xe.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Incoterms 2020 (ICC), Điều 86 Thông tư 38/2015/TT-BTC, Luật Thương mại 2005 và nhật ký scratch/daily_notes.md (06/10/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 76,
        "po": "PO-2026-W41-HOT34",
        "date": "[W41] Thứ Tư (07/10/2026)",
        "title": "Đồng hồ thông minh Smartwatch Huawei Watch GT 6 46mm mã HS 8517.62.59: Bài học xác định trước mã số theo Thông báo 22555/TB-CHQ giúp tránh tranh chấp áp mã 9102!",
        "seo": "#MaHS #Smartwatch #HuaweiWatch #ThongBao22555 #PhanLoaiHangHoa #HaiQuanVietNam #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-34-MA-HS-SMARTWATCH-8517.md",
        "content": fb_hot34,
        "visual": "Ảnh chụp sản phẩm đồng hồ thông minh hiển thị các tính năng: nhận thông báo Bluetooth, GPS, đo nhịp tim; Trích lục Thông báo xác định trước mã số 22555/TB-CHQ; Bảng so sánh đặc tính phân loại nhóm 8517 vs nhóm 9102 theo Chú giải 1(n) Chương 91.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Thông báo kết quả xác định trước mã số 22555/TB-CHQ (02/10/2026), Danh mục hàng hóa XNK Việt Nam (Thông tư 31/2022/TT-BTC), Chú giải 1(n) Chương 91 và nhật ký scratch/daily_notes.md (07/10/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 77,
        "po": "PO-2026-W41-HOT35",
        "date": "[W41] Thứ Tư (07/10/2026)",
        "title": "Doanh nghiệp FDI trong Khu chế xuất mua văn phòng phẩm, suất ăn công nghiệp, sửa chữa thiết bị từ nội địa: Có bắt buộc phải mở tờ khai hải quan không?",
        "seo": "#DNCX #KhuCheXuat #MoToKhai #Dieu74ThongTu38 #NghiDinh35 #KeToanXNK #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md",
        "content": fb_hot35,
        "visual": "Infographic phân chia 2 nhóm mặt hàng tại DNCX: Nhóm 1 (Hàng tiêu dùng phục vụ sinh hoạt: Lựa chọn mở hoặc không mở tờ khai theo Điều 74 TT 38) vs Nhóm 2 (Nguyên vật liệu sản xuất: Bắt buộc mở tờ khai); Bảng đối chiếu điều kiện thuế suất 0% GTGT.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 1 Điều 74 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC), Nghị định 35/2022/NĐ-CP, Thông tư 219/2013/TT-BTC (Điều 9) và nhật ký scratch/daily_notes.md (07/10/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 78,
        "po": "PO-2026-W41-HOT36",
        "date": "[W41] Thứ Tư (07/10/2026)",
        "title": "Hóa đơn thuế nhà thầu (FCT) phát sinh cùng ngày tờ khai hải quan nhưng kế toán bỏ sót chưa khai nộp: Hướng dẫn nộp bổ sung tránh phạt quá hạn 10 ngày!",
        "seo": "#ThueNhaThau #FCT #ThongTu103 #LuatQuanLyThue #NopThueBoSung #KeToanThue #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md",
        "content": fb_hot36,
        "visual": "Infographic mốc thời gian '10 ngày vàng' nộp tờ khai FCT theo từng lần phát sinh; Quy trình 3 bước xử lý khi phát hiện bỏ sót: Lập tờ khai bổ sung mẫu 01/NTNN, Nộp thuế kèm tiền chậm nộp 0.03%/ngày, Lưu trữ hồ sơ giải trình.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Thông tư 103/2014/TT-BTC, Điều 44 Luật Quản lý thuế số 38/2019/QH14, Nghị định 125/2020/NĐ-CP (Điều 13) và nhật ký scratch/daily_notes.md (07/10/2026).",
        "status": "Sẵn sàng đăng"
    }
]

NEW_YT_ROWS = [
    {
        "stt": 47,
        "po": "PO-2026-W41-HOT32",
        "series": "[W41] Series Pháp Lý & Thủ Tục: Công Văn 22411/CHQ-GSQL & Điều 47a Luật Hải Quan (Luật 90/2025): XNK Tại Chỗ Thay Đổi Thế Nào? Hướng Dẫn Hồ Sơ Thực Chiến Cho DN",
        "category": "PL — Pháp lý",
        "seo": "#XNKtaiCho #CongVan22411 #Dieu47a #LuatHaiQuan #HaiQuanVietNam #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-32-CV22411-XNK-TAI-CHO.md",
        "content": yt_hot32,
        "visual": "Quay tại studio Eureka; Trên bàn có văn bản CV 22411/CHQ-GSQL in sẵn và Luật số 90/2025/QH15; Screencast tra cứu văn bản hải quan; Sơ đồ luồng hàng luồng tiền tam giác XNK tại chỗ; Outro Incoterms FCA/EXW.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Công văn 22411/CHQ-GSQL (29/09/2026), Luật Hải quan số 90/2025/QH15 (Điều 47a), Điều 86 Thông tư 38/2015/TT-BTC và nhật ký scratch/daily_notes.md (01/10/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 48,
        "po": "PO-2026-W41-HOT33",
        "series": "[W41] Series Học XNK 5 Phút: XNK Tại Chỗ Dùng FCA Hay EXW? Bóc Tách Tranh Cãi Ai Chịu Cước Xe & Trách Nhiệm Thông Quan Xuất Khẩu",
        "category": "HT — Học Thực chiến",
        "seo": "#Incoterms #FCA #EXW #XNKtaiCho #LogisticsThucChien #HaiQuanVietNam #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-33-INCOTERMS-FCA-EXW-XNK-TAI-CHO.md",
        "content": yt_hot33,
        "visual": "Quay tại studio Eureka; Trên bàn có cuốn Incoterms 2020 gốc của ICC và mô hình xe tải chở hàng; Bảng so sánh 3 tiêu chí cốt lõi EXW vs FCA; Screencast mẫu hợp đồng mua bán ngoại thương 3 bên; Outro Mã HS Smartwatch.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Incoterms 2020 (ICC), Điều 86 Thông tư 38/2015/TT-BTC, Luật Thương mại 2005 và nhật ký scratch/daily_notes.md (06/10/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 49,
        "po": "PO-2026-W41-HOT34",
        "series": "[W41] Series Phân Loại & Mã HS: Smartwatch Phân Loại Mã HS Nào? Bài Học Áp Mã 8517 Thay Vì 9102 Từ Thông Báo 22555/TB-CHQ",
        "category": "HS — Mã HS & Thuế",
        "seo": "#MaHS #Smartwatch #HuaweiWatchGT6 #PhanLoaiHaiQuan #MaHS8517 #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-34-MA-HS-SMARTWATCH-8517.md",
        "content": yt_hot34,
        "visual": "Quay tại văn phòng Eureka; Cầm trên tay chiếc đồng hồ thông minh và bảng Thông báo kết quả xác định trước mã số số 22555/TB-CHQ; Trình chiếu Chú giải 1(n) Chương 91 và Quy tắc 3(b) phân loại; Outro Doanh nghiệp FDI KCX mua hàng nội địa.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Thông báo kết quả xác định trước mã số 22555/TB-CHQ (02/10/2026), Danh mục hàng hóa XNK Việt Nam (Thông tư 31/2022/TT-BTC), Chú giải 1(n) Chương 91 và nhật ký scratch/daily_notes.md (07/10/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 50,
        "po": "PO-2026-W41-HOT35",
        "series": "[W41] Series Quản Trị & Pháp Lý: Doanh Nghiệp FDI Trong Khu Chế Xuất Mua Hàng Nội Địa: Khi Nào Được Miễn Mở Tờ Khai? 3 Cụm Mặt Hàng 'Bẫy' Kế Toán Cần Tránh",
        "category": "CS — Case Study",
        "seo": "#DNCX #KhuCheXuat #HaiQuanKCX #Dieu74ThongTu38 #ThueSuat0VAT #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md",
        "content": yt_hot35,
        "visual": "Quay tại studio Eureka; Trên màn hình trình chiếu Điều 74 Thông tư 38/2015 (sửa đổi TT 39/2018) và Nghị định 35/2022/NĐ-CP; Bảng phân loại 2 nhóm: Hàng sinh hoạt vs Nguyên vật liệu sản xuất; Outro Hóa đơn thuế nhà thầu.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 1 Điều 74 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC), Nghị định 35/2022/NĐ-CP, Thông tư 219/2013/TT-BTC (Điều 9) và nhật ký scratch/daily_notes.md (07/10/2026).",
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
        ws2.row_dimensions[r].height = 45
        print(f" -> Sheet 2 Row {r}: STT {item['stt']} | {item['po']} | {item['title'][:40]}...")

    # 2. Cập nhật Sheet 3: YouTube
    ws3 = wb["YouTube"]
    start_row_yt = ws3.max_row + 1
    print(f"\nBắt đầu ghi Sheet 3 (YouTube) từ dòng {start_row_yt}...")
    
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
        ws3.row_dimensions[r].height = 45
        print(f" -> Sheet 3 Row {r}: STT {item['stt']} | {item['po']} | {item['series'][:40]}...")

    # Lưu nguyên tử (Atomic Save) tuân thủ triệt để GEMINI.md
    temp_path = EXCEL_PATH + ".tmp.xlsx"
    print(f"\nĐang thực hiện Atomic Save ra file tạm: {temp_path}...")
    wb.save(temp_path)
    wb.close()
    
    print(f"Đang thay thế nguyên tử vào file gốc: {EXCEL_PATH}...")
    os.replace(temp_path, EXCEL_PATH)
    print("✅ Cập nhật Weekly_Content_Calendar.xlsx thành công 100%! Không tạo file backup!")

if __name__ == "__main__":
    append_to_calendar()
