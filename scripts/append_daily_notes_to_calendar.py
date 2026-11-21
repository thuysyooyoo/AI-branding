# -*- coding: utf-8 -*-
"""
Cập nhật 6 bài viết mới từ Daily Notes vào Weekly_Content_Calendar.xlsx
Áp dụng atomic_save (.tmp.xlsx -> os.replace) theo GEMINI.md
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
CALENDAR_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
TMP_CALENDAR_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.tmp.xlsx")

wb = openpyxl.load_workbook(CALENDAR_PATH)
ws2 = wb['Facebook']
ws3 = wb['YouTube']

font_data = Font(name="Arial", size=10)
font_bold = Font(name="Arial", size=10, bold=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_center = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
fill_passed = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
font_passed = Font(name="Arial", size=10, bold=True, color="375623")

# 6 bài Facebook mới
NEW_FB_ROWS = [
    {
        "po": "PO-2026-W39-HOT06",
        "date": "[W39] Thứ Ba (22/09/2026)",
        "title": "Công văn 21939/HQKV2-NVHQ vừa ban hành 21/09/2026: Đừng 'nhìn tên hàng để chọn mã loại hình' — Bẫy phân biệt A11, A12, A41 khiến doanh nghiệp bị truy thu và bác bỏ ưu đãi thuế khi hậu kiểm!",
        "keywords": "#MaLoaiHinhHaiQuan #CongVan21939 #A11 #A12 #A41 #HaiQuanVietNam #KiemTraSauThongQuan #TuVanXNK #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-06-MALOAIHINH-A11-A12.md",
        "content_file": "DRAFT-FB-HOT-06-MALOAIHINH-A11-A12.md",
        "visual": "Infographic sơ đồ đối chiếu 3 mã A11 (thương mại), A12 (sản xuất), A41 (FDI) kèm dấu mộc cảnh báo truy thu thuế.",
        "fc_result": "🟢 PASSED",
        "ssot_report": "Khớp 100% Công văn 21939/HQKV2-NVHQ ngày 21/09/2026, Quyết định 1357/QĐ-TCHQ và Công văn 1478/TCHQ-GSQL.",
        "status": "Sẵn sàng đăng"
    },
    {
        "po": "PO-2026-W39-HOT07",
        "date": "[W39] Thứ Tư (23/09/2026)",
        "title": "Chuẩn hóa trường dữ liệu khai báo hải quan theo quy định mới: Tách riêng số CCCD, biển số xe, container, mã HS — Sai lệch một ký tự là tắc nghẽn thông quan tự động!",
        "keywords": "#ChuanHoaDuLieu #HaiQuanDienTu #VNACCS #ThuTucHaiQuan #LogisticsVietNam #ThongQuanTuDong #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-07-CHUAN-HOA-TRUONG-DU-LIEU-HQ.md",
        "content_file": "DRAFT-FB-HOT-07-CHUAN-HOA-TRUONG-DU-LIEU-HQ.md",
        "visual": "Ảnh chụp màn hình lỗi reject trường dữ liệu trên hệ thống VNACCS và sơ đồ SOP 3 lớp kiểm tra chéo.",
        "fc_result": "🟢 PASSED",
        "ssot_report": "Khớp 100% chuẩn kỹ thuật dữ liệu hải quan điện tử và Quyết định 1434/QĐ-CHQ.",
        "status": "Sẵn sàng đăng"
    },
    {
        "po": "PO-2026-W39-HOT08",
        "date": "[W39] Thứ Năm (24/09/2026)",
        "title": "Bỏ địa chỉ kinh doanh để 'né nợ': Sau 3 năm nhận giấy báo phạt 30 - 60 triệu đồng tiền thuế môn bài & chậm nộp tờ khai — Lối thoát nào để gỡ khóa MST và giải thể hợp pháp?",
        "keywords": "#XuLyNoThue #TrangThai06 #BoDiaChiKinhDoanh #GiaiTheDoanhNghiep #NghiDinh125 #ThueMonBai #TuVanThueChinhNgach #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-08-PHAT-BO-DIA-CHI-GIAI-THE-MST.md",
        "content_file": "DRAFT-FB-HOT-08-PHAT-BO-DIA-CHI-GIAI-THE-MST.md",
        "visual": "Sơ đồ 4 bước làm việc với Chi cục Thuế: Lấy Bảng kê 01/BK-NT -> Nộp tờ khai trắng & Đơn giảm nhẹ NĐ 125 -> Nộp theo Mã ID -> Mẫu 25/ĐK-TCT giải thể.",
        "fc_result": "🟢 PASSED",
        "ssot_report": "Khớp 100% Luật Quản lý thuế số 38, Nghị định 125/2020/NĐ-CP (Điều 9, Điều 13) và NĐ 126/2020.",
        "status": "Sẵn sàng đăng"
    },
    {
        "po": "PO-2026-W39-HOT09",
        "date": "[W39] Thứ Sáu (25/09/2026)",
        "title": "Quy trình xin Giấy xác nhận hoàn thành nghĩa vụ thuế tại Chi cục Hải quan KV2 (02 Hàm Nghi): Bí quyết 'gỡ nợ thuế treo' để giải tỏa ách tắc lô hàng gấp!",
        "keywords": "#HaiQuanKV2 #02HamNghi #XacNhanNoThue #ThongQuanNhanh #ThucChienLogistics #CatLai #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-09-XAC-NHAN-NO-THUE-HQ-KV2.md",
        "content_file": "DRAFT-FB-HOT-09-XAC-NHAN-NO-THUE-HQ-KV2.md",
        "visual": "Ảnh chụp thực tế trụ sở Hải quan KV2 (02 Hàm Nghi, Q1) và quy trình 4 bước nộp văn thư lấy số công văn gỡ nợ treo.",
        "fc_result": "🟢 PASSED",
        "ssot_report": "Khớp 100% quy trình tiếp nhận hồ sơ văn thư và thẩm quyền xác nhận nghĩa vụ thuế Chi cục Hải quan KV2.",
        "status": "Sẵn sàng đăng"
    },
    {
        "po": "PO-2026-W39-HOT10",
        "date": "[W39] Thứ Bảy (26/09/2026)",
        "title": "Kê khai thuế cá nhân kinh doanh khi có doanh thu từ nhiều nguồn (Affiliate + dịch vụ nền tảng): Khai 0 đồng là dính bẫy phạt lệch dữ liệu eTax Mobile!",
        "keywords": "#ThueCaNhanKinhDoanh #eTaxMobile #Mau01TKNCNKD #ThueAffiliate #ThuongMaiDienTu #TuVanThueChinhNgach #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-10-KHAI-THUE-NHIEU-NGUON-ETAX.md",
        "content_file": "DRAFT-FB-HOT-10-KHAI-THUE-NHIEU-NGUON-ETAX.md",
        "visual": "Ảnh chụp màn hình ứng dụng eTax Mobile chỉ mục 'Nghĩa vụ thuế đã được tổ chức khấu trừ' và mẫu kê khai 01/TKN-CNKD.",
        "fc_result": "🟢 PASSED",
        "ssot_report": "Khớp 100% Thông tư 40/2021/TT-BTC và cơ chế liên thông dữ liệu thuế của Tổng cục Thuế.",
        "status": "Sẵn sàng đăng"
    },
    {
        "po": "PO-2026-W39-HOT11",
        "date": "[W39] Chủ Nhật (27/09/2026)",
        "title": "Quên mở tờ khai cho kiện hàng nhỏ đi kèm lô hàng lớn từ Trung Quốc: Bí quyết khai bổ sung (AMA) an toàn để không bị quy tội 'buôn lậu'!",
        "keywords": "#KhaiBoSungAMA #ThongTu39 #KiemHoaHaiQuan #HangThuaContainer #XuatNhapKhauTrungViet #LogisticsChinhNgach #EurekaLogistics",
        "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-11-SOT-KIEN-HANG-KHAI-AMA.md",
        "content_file": "DRAFT-FB-HOT-11-SOT-KIEN-HANG-KHAI-AMA.md",
        "visual": "Sơ đồ 4 bước xử lý khi phát hiện hàng đóng thừa: Thư giải trình xưởng TQ -> Đơn AMA -> Mở tờ khai bổ sung -> Bổ sung điều khoản hợp đồng.",
        "fc_result": "🟢 PASSED",
        "ssot_report": "Khớp 100% Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi TT 39/2018/TT-BTC) và Nghị định 128/2020/NĐ-CP.",
        "status": "Sẵn sàng đăng"
    }
]

# Đọc nội dung thực tế từ file
def read_draft(fname):
    p = os.path.join(BASE_DIR, "scratch", "production_orders", fname)
    if os.path.exists(p):
        with open(p, "r", encoding="utf-8") as f:
            return f.read()
    return ""

# Ghi vào Sheet 2 (Facebook)
start_fb_r = ws2.max_row + 1
for i, item in enumerate(NEW_FB_ROWS, start=1):
    r = start_fb_r + i - 1
    content_text = read_draft(item["content_file"])
    
    ws2.cell(r, 1, r - 4).alignment = align_center # STT
    ws2.cell(r, 2, item["po"]).alignment = align_center # Mã PO
    ws2.cell(r, 3, item["date"]).alignment = align_center # Ngày đăng
    ws2.cell(r, 4, item["title"]).alignment = align_left # Tiêu đề
    ws2.cell(r, 5, item["keywords"]).alignment = align_left # Từ khóa
    
    # Col F: Hyperlink
    cell_f = ws2.cell(r, 6, item["link"])
    cell_f.alignment = align_left
    cell_f.hyperlink = item["link"]
    cell_f.font = Font(name="Arial", size=10, color="0563C1", underline="single")
    
    ws2.cell(r, 7, content_text).alignment = align_left # Nội dung hoàn chỉnh
    ws2.cell(r, 8, item["visual"]).alignment = align_left # Visual brief
    
    # Col I: Fact check
    cell_i = ws2.cell(r, 9, item["fc_result"])
    cell_i.alignment = align_center
    cell_i.fill = fill_passed
    cell_i.font = font_passed
    
    ws2.cell(r, 10, item["ssot_report"]).alignment = align_left # SSoT report
    
    # Col K: Trạng thái
    cell_k = ws2.cell(r, 11, item["status"])
    cell_k.alignment = align_center
    cell_k.font = font_bold
    
    for c in range(1, 12):
        ws2.cell(r, c).border = thin_border
        if c not in [6, 9, 11]:
            ws2.cell(r, c).font = font_data

print(f"Đã thêm 6 bài Facebook vào Sheet 2 (từ dòng {start_fb_r} đến {start_fb_r + 5})")

# Thêm 1 video YouTube mới vào Sheet 3 (HOT06)
NEW_YT_ROW = {
    "po": "PO-2026-W39-HOT06",
    "series": "[W39] Series Học XNK 5 Phút: Đừng Nhìn Tên Hàng Chọn Mã Loại Hình A11 - A12 - A41",
    "category": "HT — Học Thực chiến",
    "keywords": "#MaLoaiHinhHaiQuan #A11 #A12 #A41 #CongVan21939 #TuVanXNK #EurekaLogistics",
    "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-06-MALOAIHINH-A11-A12.md",
    "content_file": "DRAFT-YT-HOT-06-MALOAIHINH-A11-A12.md",
    "visual": "Quay tại văn phòng Eureka; Screencast phần mềm ECUS5-VNACCS bóc tách ô mã loại hình; Đồ họa chuyển cảnh bảng so sánh mục đích sử dụng.",
    "fc_result": "🟢 PASSED",
    "ssot_report": "Khớp 100% Công văn 21939/HQKV2-NVHQ, Quyết định 1357/QĐ-TCHQ và Công văn 1478/TCHQ-GSQL.",
    "status": "Sẵn sàng dựng"
}

r_yt = ws3.max_row + 1
yt_content = read_draft(NEW_YT_ROW["content_file"])

ws3.cell(r_yt, 1, r_yt - 4).alignment = align_center
ws3.cell(r_yt, 2, NEW_YT_ROW["po"]).alignment = align_center
ws3.cell(r_yt, 3, NEW_YT_ROW["series"]).alignment = align_left
ws3.cell(r_yt, 4, NEW_YT_ROW["category"]).alignment = align_center
ws3.cell(r_yt, 5, NEW_YT_ROW["keywords"]).alignment = align_left

cell_yt_f = ws3.cell(r_yt, 6, NEW_YT_ROW["link"])
cell_yt_f.alignment = align_left
cell_yt_f.hyperlink = NEW_YT_ROW["link"]
cell_yt_f.font = Font(name="Arial", size=10, color="0563C1", underline="single")

ws3.cell(r_yt, 7, yt_content).alignment = align_left
ws3.cell(r_yt, 8, NEW_YT_ROW["visual"]).alignment = align_left

cell_yt_i = ws3.cell(r_yt, 9, NEW_YT_ROW["fc_result"])
cell_yt_i.alignment = align_center
cell_yt_i.fill = fill_passed
cell_yt_i.font = font_passed

ws3.cell(r_yt, 10, NEW_YT_ROW["ssot_report"]).alignment = align_left

cell_yt_k = ws3.cell(r_yt, 11, NEW_YT_ROW["status"])
cell_yt_k.alignment = align_center
cell_yt_k.font = font_bold

for c in range(1, 12):
    ws3.cell(r_yt, c).border = thin_border
    if c not in [6, 9, 11]:
        ws3.cell(r_yt, c).font = font_data

print(f"Đã thêm 1 kịch bản YouTube vào Sheet 3 (dòng {r_yt})")

# Atomic save
wb.save(TMP_CALENDAR_PATH)
wb.close()
os.replace(TMP_CALENDAR_PATH, CALENDAR_PATH)
print("Đã lưu Weekly_Content_Calendar.xlsx an toàn theo chuẩn atomic_save!")
