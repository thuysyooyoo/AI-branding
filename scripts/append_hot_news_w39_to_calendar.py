# -*- coding: utf-8 -*-
"""
Script: append_hot_news_w39_to_calendar.py
Mục tiêu: Cập nhật 5 bài viết Facebook và 3 Kịch bản YouTube vào Weekly_Content_Calendar.xlsx
Tuân thủ GEMINI.md: Atomic Save, Zero Backup, Giữ nguyên định dạng bảng tính
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# Configure UTF-8
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

CALENDAR_PATH = r"d:\AI AGENT THUY\AI branding\reports\plan1\Weekly_Content_Calendar.xlsx"

def append_hot_content():
    print(f"Đang tải workbook từ: {CALENDAR_PATH}...")
    wb = openpyxl.load_workbook(CALENDAR_PATH)
    
    font_family = "Arial"
    font_regular = Font(name=font_family, size=9, color="000000")
    font_bold = Font(name=font_family, size=9, bold=True, color="000000")
    font_link = Font(name=font_family, size=9, color="0563C1", underline="single")
    
    thin_side = Side(border_style="thin", color="D3D3D3")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    # -------------------------------------------------------------
    # 1. CẬP NHẬT SHEET FACEBOOK
    # -------------------------------------------------------------
    ws_fb = wb["Facebook"]
    start_r_fb = ws_fb.max_row + 1
    
    fb_entries = [
        {
            "stt": ws_fb.max_row - 3, # dynamic STT
            "po": "PO-2026-W39-HOT01",
            "date": "[W39] Thứ Hai (22/09/2026)",
            "title": "Đóng mã số thuế đuôi 888 (Affiliate sàn TMĐT): Doanh thu 1 triệu phạt gần 6 triệu vì chậm nộp tờ khai — Cơ quan thuế rà soát dữ liệu liên thông, mng cần xử lý ngay!",
            "seo": "dong ma so thue duoi 888, thue affiliate tiep thi lien ket, phat cham nop ho so khai thue, Nghi dinh 125 2020 ND CP | #ThueAffiliate #DongMaSoThue888 #NghiDinh125 #eTaxMobile #EurekaLogistics",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-01-MST888.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-FB-HOT-01-MST888.md",
            "visual": 'Infographic so sánh 2 cột: "Nghĩa vụ nộp thuế (sàn đã trừ 7%)" vs "Nghĩa vụ nộp tờ khai (Mẫu 01/TKN-CNKD bắt buộc nộp quý theo Điều 13 NĐ 125/2020)".',
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Điều 13 & Điều 9 Nghị định 125/2020/NĐ-CP, Thông tư 40/2021/TT-BTC, Luật Quản lý thuế số 38/2019/QH14 & Case study Diễn đàn Thuế 22/09/2026.",
            "status": "Sẵn sàng đăng"
        },
        {
            "stt": ws_fb.max_row - 2,
            "po": "PO-2026-W39-HOT02",
            "date": "[W39] Thứ Ba (23/09/2026)",
            "title": "Nhập khẩu rượu ngoại & danh mục 12 nhóm hàng cửa khẩu nhập (QĐ 31/2026): Cạm bẫy thiếu giấy phép & rủi ro kẹt cont lưu bãi cảng hàng trăm triệu mùa cuối năm!",
            "seo": "thu tuc nhap khau ruou ngoai, quyet dinh 31 2026 qd ttg, thong tu 28 2026 tt bct, giay phep phan phoi ruou, tu cong bo attp | #NhapKhauRuouNgoai #QuyetDinh31 #AnToanThucPham #EurekaLogistics",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-02-RUOU-NGOAI.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-FB-HOT-02-RUOU-NGOAI.md",
            "visual": 'Infographic thẻ dọc Eureka hiển thị sơ đồ 3 cửa ải bắt buộc: 1. Giấy phép phân phối rượu (Bộ Công Thương) -> 2. Tự công bố ATTP -> 3. Thông quan tại cửa khẩu nhập theo QĐ 31/2026.',
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Quyết định 31/2026/QĐ-TTg (hiệu lực 14/08/2026), Thông tư 28/2026/TT-BCT (hiệu lực 17/07/2026), Nghị định 105/2017/NĐ-CP, Nghị định 17/2020/NĐ-CP & Case study 22/09/2026.",
            "status": "Sẵn sàng đăng"
        },
        {
            "stt": ws_fb.max_row - 1,
            "po": "PO-2026-W39-HOT03",
            "date": "[W39] Thứ Tư (24/09/2026)",
            "title": "Đếm ngược 03 ngày (25/09/2026): Công văn 21692/CHQ-CNTT yêu cầu làm sạch dữ liệu tài khoản VNACCS — Anh chị dn khẩn trương rà soát tránh bị nghẽn luồng truyền tờ khai!",
            "seo": "cong van 21692 chq cntt, lam sach du lieu nguoi su dung hai quan, he thong dang ky dksd, xac thuc csdl dan cu | #CongVan21692 #LamSachDuLieuHaiQuan #VNACCS #EurekaLogistics",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-03-CV21692-VNACCS.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-FB-HOT-03-CV21692-VNACCS.md",
            "visual": 'Ảnh đồ họa đồng hồ đếm ngược "CÒN 72H" với checklist 4 điểm rà soát tài khoản người dùng VNACCS trên cổng dksd.customs.gov.vn.',
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Công văn số 21692/CHQ-CNTT ngày 15/09/2026 của Cục Hải quan, Cổng TTĐT Tổng cục Hải quan, CSDL Quốc gia về Dân cư.",
            "status": "Sẵn sàng đăng"
        },
        {
            "stt": ws_fb.max_row,
            "po": "PO-2026-W39-HOT04",
            "date": "[W39] Thứ Năm (25/09/2026)",
            "title": "Cước container tuyến Shanghai - New York phá đỉnh 10.394 USD/FEU & bão 79 chuyến Blank Sailings trước thềm Golden Week: Lối thoát nào cho chủ hàng xuất khẩu mùa cao điểm Q4?",
            "seo": "cuoc bien drewry wci, cuoc container shanghai new york 10394 usd, blank sailings 79 chuyen, ap luc golden week | #CuocBienDrewry #BlankSailings #XuatKhauMy #EurekaLogistics",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-04-DREWRY-WCI.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-FB-HOT-04-DREWRY-WCI.md",
            "visual": "Biểu đồ biến động giá cước Drewry WCI tuyến Transpacific Bờ Đông và đồ họa 3 Liên minh hàng hải hủy 79 chuyến trước Golden Week.",
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Báo cáo Drewry World Container Index ngày 17/09/2026, dữ liệu World Shipping Council và The Loadstar.",
            "status": "Sẵn sàng đăng"
        },
        {
            "stt": ws_fb.max_row + 1,
            "po": "PO-2026-W39-HOT05",
            "date": "[W39] Thứ Sáu (26/09/2026)",
            "title": "Doanh nghiệp 'Trạng thái 06' nộp tiền nợ thuế trên eTax nhưng không được gạch nợ: Cạm bẫy tiền treo Kho bạc & cách lấy chuẩn Mã ID khoản nộp (Mẫu 01/BK-NT) trong 24h!",
            "seo": "doanh nghiep trang thai 06, nop no thue etax, ma id khoan nop, bang ke 01 bk nt, thu tra soat c1 11 ns | #NoThueDoanhNghiep #DoanhNghiepTrangThai06 #eTaxMobile #MaIDKhoanNop #EurekaLogistics",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-05-DN-TRANG-THAI-06.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-FB-HOT-05-DN-TRANG-THAI-06.md",
            "visual": "Sơ đồ 3 bước gỡ nợ eTax cho DN Trạng thái 06: Chốt Bảng kê 01/BK-NT -> Lấy Mã ID khoản nộp -> Tra soát C1-11/NS gạch nợ tự động.",
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Luật Quản lý thuế số 38/2019/QH14 (Điều 66 về tạm hoãn xuất cảnh), Thông tư 19/2021/TT-BTC về thuế điện tử & Case study thực tế ngày 22/09/2026.",
            "status": "Sẵn sàng đăng"
        }
    ]

    for idx, item in enumerate(fb_entries):
        r = start_r_fb + idx
        fill = zebra_fill if r % 2 == 0 else white_fill
        
        with open(item["file_source"], "r", encoding="utf-8") as f:
            content_text = f.read()
            
        row_vals = [
            item["stt"], item["po"], item["date"], item["title"], item["seo"],
            item["link"], content_text, item["visual"], item["fact_check"],
            item["ssot"], item["status"]
        ]
        
        for c, val in enumerate(row_vals, 1):
            cell = ws_fb.cell(row=r, column=c, value=val)
            cell.fill = fill
            cell.border = cell_border
            if c in [1, 2, 3, 9, 11]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
                
            if c == 6:
                cell.font = font_link
                cell.hyperlink = val
            elif c in [1, 2, 9, 11]:
                cell.font = font_bold
            else:
                cell.font = font_regular
        ws_fb.row_dimensions[r].height = 90
    print(f"-> Đã thêm 5 bài viết vào Sheet Facebook (Max row mới: {ws_fb.max_row})")

    # -------------------------------------------------------------
    # 2. CẬP NHẬT SHEET YOUTUBE
    # -------------------------------------------------------------
    ws_yt = wb["YouTube"]
    start_r_yt = ws_yt.max_row + 1
    
    yt_entries = [
        {
            "stt": ws_yt.max_row - 3,
            "po": "PO-2026-W39-HOT01",
            "series": "[W39] Hot Topic Thực Chiến — Tập 1: Doanh Thu 1 Triệu Phạt 6 Triệu Khi Đóng MST 888 & Cẩm Nang Gỡ Phạt eTax",
            "cat": "CS — Case Study",
            "seo": "đóng mã số thuế 888, phạt chậm nộp hồ sơ khai thuế, Nghị định 125/2020/NĐ-CP, thuế affiliate Shopee TikTok, eTax Mobile",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-01-MST888.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-YT-HOT-01-MST888.md",
            "visual": "Quay cận mặt chị Thúy tại văn phòng + Screencast trực tiếp ứng dụng eTax Mobile thao tác từng nút bấm tra cứu nợ.",
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Điều 13 & 9 NĐ 125/2020, Thông tư 40/2021/TT-BTC, ứng dụng eTax Mobile Tổng cục Thuế.",
            "status": "Sẵn sàng dựng"
        },
        {
            "stt": ws_yt.max_row - 2,
            "po": "PO-2026-W39-HOT02",
            "series": "[W39] Hot Topic Thực Chiến — Tập 2: Cẩm Nang Thủ Tục Nhập Khẩu Rượu Ngoại 2026 & Bẫy Cửa Khẩu Nhập QĐ 31/2026",
            "cat": "HT — Học Thực chiến",
            "seo": "thủ tục nhập khẩu rượu ngoại, Giấy phép phân phối rượu Bộ Công Thương, Quyết định 31/2026/QĐ-TTg, tự công bố an toàn thực phẩm rượu",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-02-RUOU-NGOAI.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-YT-HOT-02-RUOU-NGOAI.md",
            "visual": "Chị Thúy tương tác cùng mô hình chai rượu vang và cảng container + Screencast phần mềm ECUS5-VNACCS đính kèm 7 chứng từ.",
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% QĐ 31/2026/QĐ-TTg, Thông tư 28/2026/TT-BCT, Nghị định 105/2017/NĐ-CP, Nghị định 15/2018/NĐ-CP.",
            "status": "Sẵn sàng dựng"
        },
        {
            "stt": ws_yt.max_row - 1,
            "po": "PO-2026-W39-HOT04",
            "series": "[W39] Hot Topic Thực Chiến — Tập 3: Cước Container New York Vượt 10.394 USD & Chiêu Blank Sailings Của Các Hãng Tàu",
            "cat": "TN — Tin Ngành",
            "seo": "cước container Drewry WCI, Shanghai New York 10394 USD, Blank Sailings hủy 79 chuyến, giữ booking xuất khẩu Mỹ Golden Week",
            "link": "file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-04-DREWRY-WCI.md",
            "file_source": r"d:\AI AGENT THUY\AI branding\scratch\production_orders\DRAFT-YT-HOT-04-DREWRY-WCI.md",
            "visual": "Chị Thúy bên đồ thị nến Drewry WCI + Bản đồ hải trình thế giới hiển thị 79 chuyến hủy tàu mẹ.",
            "fact_check": "🟢 PASSED",
            "ssot": "Xác minh chính xác 100% Báo cáo Drewry WCI ngày 17/09/2026, World Shipping Council, The Loadstar.",
            "status": "Sẵn sàng dựng"
        }
    ]

    for idx, item in enumerate(yt_entries):
        r = start_r_yt + idx
        fill = zebra_fill if r % 2 == 0 else white_fill
        
        with open(item["file_source"], "r", encoding="utf-8") as f:
            content_text = f.read()
            
        row_vals = [
            item["stt"], item["po"], item["series"], item["cat"], item["seo"],
            item["link"], content_text, item["visual"], item["fact_check"],
            item["ssot"], item["status"]
        ]
        
        for c, val in enumerate(row_vals, 1):
            cell = ws_yt.cell(row=r, column=c, value=val)
            cell.fill = fill
            cell.border = cell_border
            if c in [1, 2, 4, 9, 11]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
                
            if c == 6:
                cell.font = font_link
                cell.hyperlink = val
            elif c in [1, 2, 9, 11]:
                cell.font = font_bold
            else:
                cell.font = font_regular
        ws_yt.row_dimensions[r].height = 90
    print(f"-> Đã thêm 3 kịch bản vào Sheet YouTube (Max row mới: {ws_yt.max_row})")

    # Lưu nguyên tử tránh khóa file
    tmp_path = CALENDAR_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    
    if os.path.exists(CALENDAR_PATH):
        os.replace(tmp_path, CALENDAR_PATH)
    else:
        os.rename(tmp_path, CALENDAR_PATH)
    print(f"THÀNH CÔNG: Đã ghi đè nguyên tử (atomic replace) tệp {CALENDAR_PATH}")

if __name__ == "__main__":
    append_hot_content()
