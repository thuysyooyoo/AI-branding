# -*- coding: utf-8 -*-
"""
Script: append_daily_notes_3009_to_calendar.py
Mục tiêu:
1. Đồng bộ các bản thảo vào reports/plan1/facebook/ và reports/plan1/video_scripts/
2. Cập nhật 6 bài viết Facebook (HOT26 -> HOT31) và 4 kịch bản YouTube (HOT26, HOT27, HOT29, HOT30) vào Weekly_Content_Calendar.xlsx:
   - Sheet 2 (Facebook): Thêm STT 68, 69, 70, 71, 72, 73
   - Sheet 3 (YouTube): Thêm STT 43, 44, 45, 46
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
    "DRAFT-FB-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md": "FB-HOT-2026-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md",
    "DRAFT-FB-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md": "FB-HOT-2026-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md",
    "DRAFT-FB-HOT-28-XUAT-HOA-DON-LE-CHUNG-CU.md": "FB-HOT-2026-28-XUAT-HOA-DON-LE-CHUNG-CU.md",
    "DRAFT-FB-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md": "FB-HOT-2026-29-CO-FORM-AI-BA-BEN-2-INVOICE.md",
    "DRAFT-FB-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md": "FB-HOT-2026-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md",
    "DRAFT-FB-HOT-31-KE-KHAI-THUE-THEO-QUY-HO-KINH-DOAN.md": "FB-HOT-2026-31-KE-KHAI-THUE-THEO-QUY-HO-KINH-DOAN.md",
}

sync_yt_map = {
    "DRAFT-YT-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md": "YT-HOT-2026-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md",
    "DRAFT-YT-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md": "YT-HOT-2026-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md",
    "DRAFT-YT-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md": "YT-HOT-2026-29-CO-FORM-AI-BA-BEN-2-INVOICE.md",
    "DRAFT-YT-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md": "YT-HOT-2026-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md",
}

print("Đang đồng bộ file sang reports/plan1/facebook và video_scripts...")
for src_name, dst_name in sync_fb_map.items():
    shutil.copy2(os.path.join(PO_DIR, src_name), os.path.join(FB_PLAN_DIR, dst_name))
    print(f" -> Đã sao chép: {dst_name}")

for src_name, dst_name in sync_yt_map.items():
    shutil.copy2(os.path.join(PO_DIR, src_name), os.path.join(YT_PLAN_DIR, dst_name))
    print(f" -> Đã sao chép: {dst_name}")

def read_draft(fname):
    p = os.path.join(PO_DIR, fname)
    with open(p, "r", encoding="utf-8") as f:
        return f.read().strip()

fb_hot26 = read_draft("DRAFT-FB-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md")
fb_hot27 = read_draft("DRAFT-FB-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md")
fb_hot28 = read_draft("DRAFT-FB-HOT-28-XUAT-HOA-DON-LE-CHUNG-CU.md")
fb_hot29 = read_draft("DRAFT-FB-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md")
fb_hot30 = read_draft("DRAFT-FB-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md")
fb_hot31 = read_draft("DRAFT-FB-HOT-31-KE-KHAI-THUE-THEO-QUY-HO-KINH-DOAN.md")

yt_hot26 = read_draft("DRAFT-YT-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md")
yt_hot27 = read_draft("DRAFT-YT-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md")
yt_hot29 = read_draft("DRAFT-YT-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md")
yt_hot30 = read_draft("DRAFT-YT-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md")

NEW_FB_ROWS = [
    {
        "stt": 68,
        "po": "PO-2026-W40-HOT26",
        "date": "[W40] Thứ Hai (28/09/2026)",
        "title": "Hóa đơn đầu vào ghi tên hàng chung chung, đầu ra ghi chi tiết đặc tính kỹ thuật: Cạm bẫy truy thu thuế & Bộ 'Khiên 3 chứng từ' tự vệ cho kế toán doanh nghiệp",
        "seo": "#HoaDonDienTu #NghiDinh123 #ThongTu78 #ThanhTraThue #ChiPhiHopLy #QuanTriKeToan #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md",
        "content": fb_hot26,
        "visual": "Infographic so sánh 2 tờ hóa đơn điện tử: Đầu vào chung chung ('Rèm cuốn') vs Đầu ra chi tiết ('Rèm cản sáng 80%'); Sơ đồ Bộ khiên 3 lớp chứng từ: Bảng quy chuẩn SKU nội bộ + Chứng từ kho vận + Phụ lục hợp đồng kỹ thuật.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Điều 10 Nghị định 123/2020/NĐ-CP, Thông tư 78/2021/TT-BTC, Luật Quản lý thuế số 38/2019/QH14 và nhật ký scratch/daily_notes.md (28/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 69,
        "po": "PO-2026-W40-HOT27",
        "date": "[W40] Thứ Ba (29/09/2026)",
        "title": "Cần hủy & sửa đổi Báo cáo Quyết toán mẫu số 15 tại Chi cục Hải quan Quản lý Hàng gia công TP.HCM (Cơ sở 15): Cảnh báo bẫy quá hạn 60 ngày Điều 60 Thông tư 39 khiến doanh nghiệp bị kiểm tra sau thông quan tại trụ sở!",
        "seo": "#BaoCaoQuyetToan #Mau15GiaCong #Dieu60ThongTu39 #HaiQuanCoSo15 #KiemTraSauThongQuan #GiaCongSXXK #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md",
        "content": fb_hot27,
        "visual": "Infographic dòng thời gian 60 ngày vàng sửa đổi Báo cáo Quyết toán theo Điều 60 Thông tư 39; Bảng so sánh rủi ro: Tự sửa trong 60 ngày (Miễn phạt) vs Quá 60 ngày (Bị phạt & Kiểm tra sau thông quan); Quy trình 4 bước nộp hồ sơ tại cơ sở 15.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 2 Điều 60 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC), Nghị định 128/2020/NĐ-CP và nhật ký scratch/daily_notes.md (29/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 70,
        "po": "PO-2026-W40-HOT28",
        "date": "[W40] Thứ Ba (29/09/2026)",
        "title": "Ở nhà chung cư không đăng ký Hộ kinh doanh được: Muốn xuất hóa đơn cho khách hàng doanh nghiệp thì làm thế nào? Cẩm nang xin cấp hóa đơn lẻ từng lần phát sinh từ Chi cục Thuế!",
        "seo": "#HoaDonLe #ChiCucThue #HoKinhDoanhChungCu #NghiDinh123 #ThongTu40 #KhoiNghiepKinhDoanh #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-28-XUAT-HOA-DON-LE-CHUNG-CU.md",
        "content": fb_hot28,
        "visual": "Infographic quy trình 3 bước xin cấp hóa đơn lẻ tại Chi cục Thuế cho cá nhân ở chung cư; Bảng tính thuế GTGT & TNCN theo tỷ lệ ngành nghề (Dịch vụ gia công 4.5% vs Bán lẻ 1.5%); Mẫu đơn 06/ĐN-PSĐT Nghị định 123.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Luật Nhà ở 2014 & Luật Nhà ở 2023 (Điều 3), Điều 13 Nghị định 123/2020/NĐ-CP, Thông tư 40/2021/TT-BTC và nhật ký scratch/daily_notes.md (29/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 71,
        "po": "PO-2026-W40-HOT29",
        "date": "[W40] Thứ Tư (30/09/2026)",
        "title": "C/O Form AI 3 bên thể hiện 2 trị giá FOB và 2 số Invoice (nhà sản xuất & bên thứ ba): Hải quan có chấp nhận không? Cẩm nang đối soát ô số 9, 10, 13 theo Hiệp định AIFTA tránh bị bác ưu đãi thuế!",
        "seo": "#COFormAI #XuatXuHangHoa #ThirdPartyInvoicing #AIFTA #HaiQuanVietNam #ThueUuDai #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md",
        "content": fb_hot29,
        "visual": "Hình ảnh minh họa mẫu C/O Form AI thực tế phóng to Ô số 7 (Bên thứ ba), Ô số 9 (Trị giá FOB kép), Ô số 10 (Invoice kép) và Ô số 13 (Tích Third Country Invoicing); Sơ đồ đối soát Bộ chứng từ giải trình 4 món.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Hiệp định AIFTA, Thông tư số 15/2010/TT-BCT, Thông tư số 33/2023/TT-BTC, Công văn số 12149/TCHQ-GSQL và nhật ký scratch/daily_notes.md (30/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 72,
        "po": "PO-2026-W40-HOT30",
        "date": "[W40] Thứ Tư (30/09/2026)",
        "title": "DNCX gửi hạt nhựa đi gia công biến tính nội địa không làm đổi mã HS: Ghi nhãn và khai xuất xứ 'Made in China' hay 'Việt Nam'? Cạm bẫy xử phạt xuất xứ theo Nghị định 111/2021/NĐ-CP & Thông tư 38/2015!",
        "seo": "#DNCX #GiaCongNoiDia #NghiDinh111 #NhanHangHoa #XuatXuHangHoa #HatNhuaBienTinh #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md",
        "content": fb_hot30,
        "visual": "Mẫu nhãn mác chuẩn theo Điều 15 Nghị định 111: 'Gia công tại Việt Nam từ nguyên liệu nhập khẩu Trung Quốc'; Sơ đồ phân định 2 tiêu chuẩn xuất xứ CTC và RVC; Bảng chỉ dẫn khai mã xuất xứ 'CN' và ghi chú tờ khai gia công ECUS5.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Nghị định 111/2021/NĐ-CP, Nghị định 31/2018/NĐ-CP, Điều 76 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018) và nhật ký scratch/daily_notes.md (30/09/2026).",
        "status": "Sẵn sàng đăng"
    },
    {
        "stt": 73,
        "po": "PO-2026-W40-HOT31",
        "date": "[W40] Thứ Tư (30/09/2026)",
        "title": "Kê khai thuế hộ kinh doanh & cá nhân theo quý: Lập tờ khai cộng dồn doanh thu 3 tháng hay lập từng tháng? 3 sai lầm khiến chủ shop bị phạt chậm nộp và ấn định thuế theo Thông tư 40/2021/TT-BTC!",
        "seo": "#KeKhaiThueTheoQuy #HoKinhDoanh #ThongTu40 #NghiDinh126 #eTaxMobile #QuyetToanThue #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-31-KE-KHAI-THUE-THEO-QUY-HO-KINH-DOAN.md",
        "content": fb_hot31,
        "visual": "Lịch hạn chót nộp thuế 4 quý chuẩn theo Luật Quản lý thuế: 30/04, 31/07, 31/10, 31/01; Screencast giao diện lập tờ khai mẫu 01/CNKD trên ứng dụng eTax Mobile; Bảng cảnh báo rủi ro dữ liệu liên thông với sàn TMĐT Shopee/TikTok Shop.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Điều 44 Luật Quản lý thuế số 38/2019/QH14, Điều 8, 9 Nghị định 126/2020/NĐ-CP, Thông tư 40/2021/TT-BTC và nhật ký scratch/daily_notes.md (30/09/2026).",
        "status": "Sẵn sàng đăng"
    }
]

NEW_YT_ROWS = [
    {
        "stt": 43,
        "po": "PO-2026-W40-HOT26",
        "series": "[W40] Series Quản Trị & Pháp Lý: Hóa Đơn Đầu Vào Ghi Chung Chung, Đầu Ra Ghi Chi Tiết: Bẫy Truy Thu Thuế Tiền Tỷ & Bộ Khiên Chắn 3 Chứng Từ Tự Vệ Cho Kế Toán",
        "category": "CS — Case Study",
        "seo": "#HoaDonDienTu #NghiDinh123 #ThanhTraThue #ChiPhiHopLy #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md",
        "content": yt_hot26,
        "visual": "Quay tại studio Eureka; Trên bàn để 2 tờ hóa đơn điện tử in sẵn và mô hình chiếc cân công lý; Screencast Diễn đàn Thuế chuyên sâu; Trình chiếu Bảng quy chuẩn mã SKU nội bộ và Phiếu kho vận; Outro BCQT mẫu 15.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Điều 10 Nghị định 123/2020/NĐ-CP, Thông tư 78/2021/TT-BTC, Luật Quản lý thuế số 38/2019/QH14 và nhật ký scratch/daily_notes.md (28/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 44,
        "po": "PO-2026-W40-HOT27",
        "series": "[W40] Series Học XNK 5 Phút: Cần Hủy & Sửa Đổi Báo Cáo Quyết Toán Mẫu 15: Cạm Bẫy Quá Hạn 60 Ngày Điều 60 Thông Tư 39 & Quy Trình Gỡ Khẩn Cấp",
        "category": "HT — Học Thực chiến",
        "seo": "#BaoCaoQuyetToan #Mau15GiaCong #Dieu60ThongTu39 #HaiQuanCoSo15 #KiemTraSauThongQuan #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md",
        "content": yt_hot27,
        "visual": "Quay tại văn phòng Eureka; Trên màn hình hiển thị mẫu số 15/BCQT-NVL/GSQL và dòng chữ cảnh báo 'LỆCH TỒN KHO 50 TẤN'; Screencast phân hệ Gia công ECUS5-VNACCS; Sơ đồ thời hạn 60 ngày vàng Điều 60; Outro Chung cư HKD.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Khoản 2 Điều 60 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC), Nghị định 128/2020/NĐ-CP và nhật ký scratch/daily_notes.md (29/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 45,
        "po": "PO-2026-W40-HOT29",
        "series": "[W40] Series Học XNK 5 Phút: C/O Form AI 3 Bên Thể Hiện 2 Trị Giá FOB & 2 Số Invoice: Hải Quan Có Chấp Nhận? Checklist 4 Điểm Sống Còn Tránh Bị Bác Thuế 0%",
        "category": "HT — Học Thực chiến",
        "seo": "#COFormAI #ThirdPartyInvoicing #AIFTA #HaiQuanVietNam #KiemTraCO #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md",
        "content": yt_hot29,
        "visual": "Quay tại studio Eureka; Cầm bộ C/O Form AI màu trắng xanh có dấu mộc EIC Ấn Độ; Đồ họa sơ đồ tam giác thương mại 3 bên; Phóng to Ô số 7, Ô số 9, Ô số 10 và Ô số 13; Trình chiếu Bộ chứng từ giải trình 4 món; Outro Gia công DNCX.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Hiệp định AIFTA, Thông tư số 15/2010/TT-BCT, Thông tư số 33/2023/TT-BTC, Công văn số 12149/TCHQ-GSQL và nhật ký scratch/daily_notes.md (30/09/2026).",
        "status": "Sẵn sàng dựng"
    },
    {
        "stt": 46,
        "po": "PO-2026-W40-HOT30",
        "series": "[W40] Series Học XNK 5 Phút: DNCX Gửi Hàng Gia Công Biến Tính Không Đổi Mã HS: Ghi Nhãn & Khai Xuất Xứ Thế Nào Tránh Phạt Tiền Tỷ Nghị Định 111?",
        "category": "CS — Case Study",
        "seo": "#DNCX #GiaCongNoiDia #NghiDinh111 #NhanHangHoa #HatNhuaBienTinh #EurekaLogistics #JasmineNguyen",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md",
        "content": yt_hot30,
        "visual": "Quay tại văn phòng Eureka; Trên bàn có 2 mẫu bao bì hạt nhựa: 'Made in China' vs 'Made in Vietnam'; Đồ họa chuỗi luân chuyển hàng gia công DNCX; Bảng so sánh tiêu chuẩn xuất xứ CTC và RVC; Mẫu nhãn mác chuẩn Nghị định 111; Outro Kê khai thuế quý.",
        "factcheck": "🟢 PASSED",
        "ssot": "Khớp 100% Nghị định 111/2021/NĐ-CP, Nghị định 31/2018/NĐ-CP, Điều 76 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018) và nhật ký scratch/daily_notes.md (30/09/2026).",
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
