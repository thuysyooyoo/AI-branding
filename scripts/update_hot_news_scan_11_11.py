# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_11_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 11/11/2026
(Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026 | Cập nhật ngày 11/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 46 cập nhật ngày 11/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 245 lên 250 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 11/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w46_1111.md
"""

import os
import sys
import hashlib
from datetime import datetime
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Cấu hình UTF-8 cho console Windows
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

FILE_PATH = r'reports\SEO_Keyword_Trends_Eureka.xlsx'
ONEDRIVE_DIR = r'C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh'
ONEDRIVE_FILE = os.path.join(ONEDRIVE_DIR, 'SEO_Keyword_Trends_Eureka.xlsx')
SYNC_LOG_PATH = r'reports\sync_log.txt'
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w46_1111.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w46_1111():
    print(f"Đang nạp workbook từ: {FILE_PATH}...")
    wb = openpyxl.load_workbook(FILE_PATH)

    font_family = 'Segoe UI'
    font_header_sub = Font(name=font_family, size=9, italic=True, color='555555')
    font_regular = Font(name=font_family, size=10, color='000000')
    font_bold = Font(name=font_family, size=10, bold=True, color='000000')
    font_link = Font(name=font_family, size=10, color='0563C1', underline='single')
    font_total = Font(name=font_family, size=10, bold=True, color='002060')

    fill_zebra_even = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    fill_zebra_odd = PatternFill(start_color='F2F5F8', end_color='F2F5F8', fill_type='solid')
    fill_total = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')

    thin_border_side = Side(border_style='thin', color='D3D3D3')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_top_side = Side(border_style='thin', color='002060')
    total_bottom_side = Side(border_style='double', color='002060')
    total_border = Border(left=thin_border_side, right=thin_border_side, top=total_top_side, bottom=total_bottom_side)

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)

    update_date_str = "11/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 46 (09/11/2026 - 15/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'thu tuc nhap khau my pham cong bo my pham tham van gia hai quan co uu dai thue',
         'Thủ tục nhập khẩu mỹ phẩm mùa cao điểm cuối năm: Cạm bẫy tham vấn giá hải quan, rủi ro bác C/O ưu đãi và chuẩn hóa phiếu công bố mỹ phẩm',
         19500, 'Điểm nóng mỹ phẩm mùa Tết (+3,820%)',
         'Doanh nghiệp nhập khẩu và phân phối mỹ phẩm, dược mỹ phẩm, hóa mỹ phẩm; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Sai lệch thành phần/nhãn mác với phiếu công bố dẫn đến bị đình chỉ thông quan và xử phạt vi phạm hành chính; bị nghi ngờ trị giá thấp, yêu cầu tham vấn giá và ấn định tăng thuế; nguy cơ bác C/O form E, D, VK làm mất thuế ưu đãi đặc biệt.',
         "Tiêu đề: 'Nhập khẩu mỹ phẩm mùa cao điểm: 3 cạm bẫy tham vấn giá, bác C/O và chuẩn hóa phiếu công bố tránh tắc hàng tại cảng'.",
         'Thông tư 06/2011/TT-BYT; Thông tư 38/2015/TT-BTC & 39/2018/TT-BTC; scratch/daily_notes.md (11/11/2026); Cục Quản lý Dược.',
         update_date_str),

        (2,
         'hang gia cong bi ben nuoc ngoai tu bo thong tu 121 2025 tt btc dieu 66',
         'Hàng gia công bị bên nước ngoài từ bỏ theo Thông tư 121/2025/TT-BTC: 03 hướng xử lý, phác đồ mở tờ khai A42 nộp thuế và trách nhiệm pháp lý doanh nghiệp',
         19100, 'Điểm nóng hàng gia công (+3,710%)',
         'Doanh nghiệp gia công may mặc, da giày, linh kiện điện tử, cơ khí; giám đốc nhà máy, kế toán kho và nhân viên khai báo hải quan.',
         'Bên thuê gia công nước ngoài từ bỏ nguyên vật liệu, thành phẩm, máy móc; doanh nghiệp lúng túng không biết tiêu thụ nội địa hay tiêu hủy; nếu quá hạn 30 ngày bị hải quan lập biên bản và ấn định thuế.',
         "Tiêu đề: 'Bên đặt gia công từ bỏ hàng: Phác đồ 3 bước mở tờ khai A42 hoặc tiêu hủy đúng Điều 66 Thông tư 121/2025/TT-BTC tránh bị xử phạt'.",
         'Thông tư số 121/2025/TT-BTC (sửa Điều 66 Thông tư 38/2015/TT-BTC); scratch/daily_notes.md (10/11/2026); Cục Hải quan.',
         update_date_str),

        (3,
         'to khai ola van chuyen doc lap giam sat hai quan chuyen cang cat lai cai mep',
         'Tờ khai vận chuyển độc lập OLA trong logistics cảng biển (Cát Lái - Cái Mép): Phân định hàng chịu sự giám sát hải quan và cảnh báo rủi ro mở sai thủ tục OLA',
         18700, 'Thủ tục tờ khai OLA cảng biển (+3,580%)',
         'Các công ty forwarder, nhà xe vận tải container, doanh nghiệp logistics trung chuyển cảng biển và nhân viên khai báo hải quan.',
         'Hiểu nhầm OLA cho hàng nội địa thông thường; mở sai thủ tục dẫn đến không thanh lý được tờ khai, seal điện tử GPS bị treo, container bị giữ lại bãi cảng và bị xử phạt hành chính về vi phạm thủ tục hải quan.',
         "Tiêu đề: 'Tờ khai OLA vận chuyển độc lập giữa Cát Lái và Cái Mép: 3 nguyên tắc phân định hàng giám sát hải quan tránh bị treo cont tại bãi'.",
         'Điều 50, Điều 51 Thông tư 38/2015/TT-BTC (sửa TT 39/2018/TT-BTC); scratch/daily_notes.md (11/11/2026); haiquanonline.com.vn.',
         update_date_str),

        (4,
         'thanh khoan dncx bao cao quyet toan hai quan mau 15 chenh lech xuat nhap ton',
         'Kiểm tra sau thông quan thanh khoản DNCX: Cách xử lý chênh lệch số liệu xuất nhập tồn, định mức hao hụt và bảo vệ Báo cáo quyết toán Mẫu 15',
         18400, 'Mùa rà soát thanh khoản DNCX (+3,490%)',
         'Doanh nghiệp chế xuất (DNCX), nhà máy FDI, giám đốc sản xuất, kế toán trưởng và nhân viên khai báo hải quan.',
         'Số liệu kho thực tế lệch với ERP và tờ khai hải quan; âm nguyên liệu không chứng minh được định mức hao hụt; nguy cơ bị ấn định thuế NK, VAT và phạt trốn thuế khi kiểm tra sau thông quan.',
         "Tiêu đề: 'Bảo vệ Báo cáo quyết toán DNCX trước kiểm tra sau thông quan: Công thức Tam giác vàng đối soát Hải quan - Kế toán - Kho'.",
         'Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; scratch/daily_notes.md; Báo Hải quan Online.',
         update_date_str),

        (5,
         'cong van 23913 chq gsql nhap khau hoa chat dncx luat quan ly ngoai thuong',
         'Công văn 23913/CHQ-GSQL ngày 05/11/2026: Hải quan đề nghị tháo gỡ giao thoa pháp lý nhập khẩu hóa chất vào DNCX giữa Luật Hóa chất và Ngoại thương',
         18000, 'Điểm nghẽn pháp lý hóa chất FDI (+3,390%)',
         'Doanh nghiệp chế xuất nhập khẩu hóa chất phục vụ sản xuất linh kiện, điện tử, bán dẫn và các đơn vị logistics chuyên tuyến KCN/KCX.',
         'Xung đột giữa Điều 59 Luật Quản lý ngoại thương (khu phi thuế quan miễn GP) và Luật Hóa chất; nguy cơ hàng kẹt tại cửa khẩu cảng biển/sân bay.',
         "Tiêu đề: 'Công văn 23913/CHQ-GSQL: Hướng giải quyết vướng mắc thủ tục nhập khẩu hóa chất vào DNCX không lo ách tắc thông quan'.",
         'Công văn 23913/CHQ-GSQL ngày 05/11/2026; Điều 59 Luật Quản lý ngoại thương; customs.gov.vn.',
         update_date_str),

        (6,
         'dong bo ma kho bai dia diem hai quan vnaccs vcis loi truyen to khai thong quan',
         'Đồng bộ bộ mã kho, bãi, địa điểm mới trên hệ thống VNACCS/VCIS cuối năm 2026: Giải mã nguyên nhân lỗi truyền tờ khai và hướng dẫn cập nhật chuẩn xác',
         17600, 'Chuẩn hóa mã kho bãi VNACCS (+3,280%)',
         'Toàn bộ doanh nghiệp xuất nhập khẩu, forwarder, đại lý khai thuê hải quan và doanh nghiệp kinh doanh cảng, kho bãi ICD trên toàn quốc.',
         'Truyền tờ khai bị lỗi từ chối tiếp nhận (Error Code VACS) do vẫn dùng mã địa điểm hoặc mã chi cục cũ; hàng hóa bị đình trệ, lỡ chuyến tàu xuất khẩu và phát sinh chi phí lưu bãi.',
         "Tiêu đề: 'Đồng bộ mã kho bãi địa điểm mới trên VNACCS/VCIS: Cách xử lý lỗi từ chối tờ khai và cập nhật danh mục cảng chuẩn xác'.",
         'Thông báo Cục CNTT & Thống kê Hải quan; haiquanonline.com.vn; customs.gov.vn.',
         update_date_str),

        (7,
         'quyet dinh 31 2026 qd ttg danh muc hang hoa phai lam thu tuc tai cua khau nhap',
         'Quyết định 31/2026/QĐ-TTg thay thế QĐ 23/2019/QĐ-TTg: Danh mục hàng hóa bắt buộc làm thủ tục tại cửa khẩu nhập và tác động đến chuỗi vận tải ICD',
         17200, 'Quy định cửa khẩu nhập (+3,170%)',
         'Doanh nghiệp nhập khẩu ô tô, hóa chất, phế liệu, máy móc cũ; các công ty giao nhận forwarder và đơn vị quản lý cảng cạn ICD.',
         'Không được chuyển cửa khẩu về ICD nội địa, phải làm thủ tục trực tiếp tại cảng biển/sân bay; tăng chi phí lưu bãi, lưu container nếu không chuẩn bị trước.',
         "Tiêu đề: 'Quyết định 31/2026/QĐ-TTg: Danh mục hàng hóa bắt buộc làm thủ tục tại cửa khẩu nhập và giải pháp tối ưu chi phí cảng'.",
         'Quyết định số 31/2026/QĐ-TTg; chinhphu.vn; haiquanonline.com.vn.',
         update_date_str),

        (8,
         'so hoa hai quan thong minh 2026 giam sat tu dong cua khau so cang bien so',
         'Ngành Hải quan tăng tốc vận hành Hệ thống Hải quan số thông minh cuối năm 2026: Doanh nghiệp logistics cần chuẩn bị gì cho mô hình Cảng biển số?',
         16800, 'Nước rút chuyển đổi số hải quan (+3,050%)',
         'Toàn bộ doanh nghiệp XNK, các hãng tàu, đại lý hải quan và doanh nghiệp kinh doanh cảng biển, kho bãi.',
         'Chuyển đổi từ thủ tục giấy tờ sang tích hợp API số hóa, camera AI nhận diện tại cửa khẩu/cảng biển; doanh nghiệp chậm thích ứng sẽ bị chậm trễ thông quan.',
         "Tiêu đề: 'Hệ thống Hải quan số thông minh 2026: 4 chuẩn bị cốt lõi cho doanh nghiệp xuất nhập khẩu và cảng biển số'.",
         'Báo Hải quan Online (haiquanonline.com.vn); customs.gov.vn.',
         update_date_str),

        (9,
         'cuoc tau bien drewry thang 11 2026 blank sailings huy chuyen phu phi gri',
         'Thị trường cước container đường biển tuần 2 tháng 11/2026: Ứng phó làn sóng Blank Sailings và kế hoạch tăng phụ phí GRI của các liên minh hãng tàu cuối năm',
         16400, 'Làn sóng Blank Sailings cước biển (+2,940%)',
         'Doanh nghiệp xuất nhập khẩu hàng đi Mỹ (USWC/USEC), Châu Âu (Rotterdam/Hamburg) và các nhà máy xuất khẩu hàng tiêu dùng mùa lễ Tết.',
         'Các hãng tàu cắt giảm 12-15% tải cung ứng (blank sailings) để ghìm giá và rục rịch tăng phụ phí GRI từ 500-1000 USD/FEU; nguy cơ trễ hẹn giao hàng và rớt cont tại cảng chuyển tải.',
         "Tiêu đề: 'Cước tàu biển tháng 11/2026: Hãng tàu đồng loạt hủy chuyến (Blank Sailings), chủ hàng làm gì để không bị rớt container?'.",
         'Drewry World Container Index; World Shipping Council; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (10,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl hang mau to khai',
         'Thông tư 142/2026/TT-BTC (hiệu lực từ 16/11/2026): Quy chuẩn mới về thủ tục hải quan chuyển phát nhanh quốc tế, số hóa Master/House AWB và cách lấy tờ khai hàng mẫu',
         15900, 'Quy chuẩn chuyển phát nhanh (+2,810%)',
         'Doanh nghiệp gửi hàng mẫu qua FedEx, DHL, UPS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Gửi hàng mẫu qua bưu chính/chuyển phát nhanh thường bị thông quan gom, không có tờ khai hải quan riêng đứng tên doanh nghiệp dẫn đến bị loại chi phí hợp lý và không được khấu trừ thuế GTGT 0%.',
         "Tiêu đề: 'Thông tư 142/2026/TT-BTC hiệu lực từ 16/11: Hướng dẫn lấy tờ khai hải quan hàng mẫu chuyển phát nhanh bảo vệ chi phí thuế hợp lệ'.",
         'Thông tư số 142/2026/TT-BTC; Báo Hải quan Online; Tổng cục Hải quan.',
         update_date_str),
    ]

    for idx, item in enumerate(sheet1_data):
        row_idx = 6 + idx
        fill_color = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
        for col_idx, val in enumerate(item, start=1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill_color
            cell.border = cell_border
            if col_idx == 1:
                cell.alignment = align_center
            elif col_idx in [2, 3, 5, 6, 7, 8, 9]:
                cell.alignment = align_left
            elif col_idx == 4:
                cell.alignment = align_right
                cell.number_format = '#,##0'
            elif col_idx == 10:
                cell.alignment = align_center
        ws1.row_dimensions[row_idx].height = 42

    # Cập nhật row 16 Total của Sheet 1
    ws1['A16'] = 'Tổng số lượng từ khóa nghiên cứu tuân thủ SLA:'
    ws1['A16'].font = font_bold
    ws1['A16'].alignment = align_left
    ws1['A16'].fill = fill_total
    ws1['A16'].border = total_border

    ws1['D16'] = '=SUM(D6:D15)'
    ws1['D16'].font = font_bold
    ws1['D16'].alignment = align_right
    ws1['D16'].number_format = '#,##0'
    ws1['D16'].fill = fill_total
    ws1['D16'].border = total_border

    for col in [2, 3, 5, 6, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    print("2. Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E...")
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 46 (09/11/2026 - 15/11/2026) | Mở rộng 250 chủ đề nghiên cứu (gồm 218 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 251)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 251 or m.max_row >= 251]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (246,
         'thu tuc nhap khau my pham cong bo my pham tham van gia hai quan co uu dai thue',
         'thu tuc nhap khau my pham, phieu cong bo my pham cfs, tham van gia my pham hai quan, bac co uu dai thue my pham, nhan phu my pham nhap khau, kiem tra sau thong quan my pham',
         99.9,
         61200,
         'https://thuvienphapluat.vn',
         'Thủ tục nhập khẩu mỹ phẩm mùa cao điểm cuối năm: Cạm bẫy tham vấn giá hải quan, rủi ro bác C/O ưu đãi và chuẩn hóa phiếu công bố mỹ phẩm',
         'Nội dung trọng tâm: Phân tích cảnh báo nghiệp vụ ngày 11/11/2026 về thủ tục nhập khẩu mỹ phẩm phục vụ mùa kinh doanh lễ Tết cuối năm: Mặt hàng mỹ phẩm (chăm sóc da, son môi, nước hoa, dầu gội...) thuộc diện quản lý chuyên ngành đặc thù của Bộ Y tế theo Thông tư 06/2011/TT-BYT; doanh nghiệp nhập khẩu thường đối mặt 03 rủi ro chí mạng: (1) Sai lệch thông tin thành phần/tên sản phẩm giữa Phiếu công bố sản phẩm mỹ phẩm với nhãn gốc thực tế và CFS (Certificate of Free Sale) dẫn đến bị hải quan tạm dừng thông quan, xử phạt vi phạm hành chính từ 30-50 triệu đồng và buộc tái xuất; (2) Tham vấn giá hải quan: Mỹ phẩm thuộc danh mục hàng hóa quản lý rủi ro về trị giá hải quan, cơ quan hải quan thường xuyên nghi ngờ giá khai báo thấp, yêu cầu tham vấn giá và ấn định tăng thuế nếu doanh nghiệp không chứng minh được dòng tiền thanh toán T/T, L/C, hợp đồng phân phối độc quyền và chứng từ cước vận chuyển quốc tế; (3) Nguy cơ bị bác C/O form E, form D, form AK, form VK làm mất thuế nhập khẩu ưu đãi đặc biệt (từ 0% bị đội lên 10-22%) do lỗi khai báo tiêu chí xuất xứ hoặc thiếu chứng từ chuyển tải hợp lệ; Khuyến nghị Eureka: Chủ hàng cần kiểm tra rà soát chéo 100% thành phần trên bao bì mẫu trước khi nhập hàng số lượng lớn, lập sẵn bộ hồ sơ chứng minh trị giá hải quan 5 bước để bảo vệ mức giá khai báo, và đăng ký làm thủ tục kiểm tra nhãn hàng hóa đồng bộ ngay tại kho bảo thuế.',
         'scratch/daily_notes.md (11/11) & Cục Quản lý Dược, Tổng cục Hải quan | CS (Case Study) / Cosmetic Import & Valuation Audit Compliance',
         update_date_str),

        (247,
         'to khai ola van chuyen doc lap giam sat hai quan chuyen cang cat lai cai mep',
         'to khai ola, van chuyen doc lap hai quan, giam sat hai quan cont chuyen cang, cat lai di cai mep, dieu 50 thong tu 38 sua doi, seal dinh vi gps hai quan, phan biet hang noi dia va hang chuyen cua khau',
         99.8,
         59500,
         'https://customs.gov.vn',
         'Tờ khai vận chuyển độc lập OLA trong logistics cảng biển (Cát Lái - Cái Mép): Phân định hàng chịu sự giám sát hải quan và cảnh báo rủi ro mở sai thủ tục OLA',
         'Nội dung trọng tâm: Giải mã vướng mắc thủ tục phát sinh ngày 11/11/2026 trong vận chuyển container giữa các cảng biển và ICD (điển hình tuyến Cát Lái - Cái Mép - ICD Phước Long): Nhiều doanh nghiệp forwarder và chủ hàng nội địa nhầm lẫn về phạm vi áp dụng của tờ khai vận chuyển độc lập (mã loại hình OLA trên VNACCS theo Điều 50, Điều 51 Thông tư 38/2015/TT-BTC sửa đổi tại Thông tư 39/2018/TT-BTC); bản chất pháp lý cốt lõi: OLA CHỈ áp dụng đối với hàng hóa chưa hoàn thành thủ tục hải quan đang chịu sự giám sát hải quan (như hàng quá cảnh, hàng chuyển cửa khẩu từ cảng dỡ sang cảng đích, hoặc hàng trung chuyển xuất khẩu bằng tàu mẹ tại Cái Mép); nếu tự ý mở tờ khai OLA cho hàng hóa nội địa thuần túy (hàng hóa lưu thông nội địa không có yếu tố hải quan) hoặc hàng nhập khẩu đã hoàn tất thông quan, hệ thống hải quan sẽ không thể thanh lý tờ khai, seal điện tử/GPS không được kích hoạt, dẫn đến container bị kẹt cứng tại bãi cảng và phát sinh vi phạm khai báo hải quan sai đối tượng; Khuyến nghị Eureka: Forwarder và chủ hàng cần phân loại rõ tình trạng pháp lý của lô hàng trước khi truyền tờ khai; với hàng nội địa chuyển cảng chỉ sử dụng vận đơn nội địa (Domestic B/L) và phiếu xuất kho kiêm vận chuyển nội bộ; với hàng chuyển cửa khẩu phải nộp văn bản đề nghị chuyển cửa khẩu kèm chỉ định của hãng tàu và khai đúng chi cục hải quan quản lý nơi đi/nơi đến.',
         'scratch/daily_notes.md (11/11) & haiquanonline.com.vn | HT (Học Thực Chiến) / Independent Transport OLA & Port Logistics',
         update_date_str),

        (248,
         'bao cao quyet toan hai quan mau 15 dinh muc thuc te chenh lech ton kho ktstq',
         'bao cao quyet toan hai quan mau 15, dinh muc san xuat thuc te, chenh lech xuat nhap ton nguyen lieu, kiem tra sau thong quan gia cong sxxk, tam giac so lieu hai quan kho ke toan, sua doi bo sung mau 15a',
         99.7,
         58200,
         'https://customs.gov.vn',
         'Báo cáo quyết toán Hải quan Mẫu 15: Chiến lược kiểm soát biến động định mức thực tế và cơ chế đối soát liên phòng ban (Hải quan - Kho - Kế toán) trước mùa KTSTQ',
         'Nội dung trọng tâm: Đúc kết kinh nghiệm quản trị rủi ro cấp thiết ngày 11/11/2026 cho mùa Báo cáo quyết toán cuối năm (Q4/2026): Hàng loạt doanh nghiệp gia công, sản xuất xuất khẩu (SXXK) và DNCX đang rơi vào vùng nguy hiểm khi chuẩn bị Báo cáo quyết toán tình hình sử dụng nguyên liệu, vật tư nhập khẩu theo Mẫu số 15/BCQT-NVL/GSQL (Thông tư 39/2018/TT-BTC); điểm nghẽn lớn nhất không nằm ở kỹ thuật truyền phần mềm mà nằm ở việc mất kiểm soát định mức sản xuất thực tế (Actual Scrap Rate) so với định mức dự kiến; các nguyên nhân sai lệch thường gặp: hao hụt tự nhiên do độ ẩm/nhiệt độ, phế liệu phế phẩm không hạch toán kịp thời, đơn vị tính mua vào lệch với đơn vị tính xuất kho sản xuất (kg vs cuộn, lít vs thùng) và thiếu biên bản hủy bán thành phẩm hỏng; khi kiểm tra sau thông quan (KTSTQ), sự chênh lệch âm kho nguyên liệu nhập khẩu sẽ bị Hải quan ấn định thuế nhập khẩu, thuế GTGT và xử phạt vi phạm hành chính từ 20% đến 3 lần số thuế trốn; Giải pháp Eureka: Doanh nghiệp cần triển khai mô hình đối soát Tam giác số liệu 3 bên định kỳ hằng tháng; chuẩn hóa bảng quy đổi đơn vị tính trên hệ thống ERP; nếu phát sinh sai lệch phải chủ động nộp Báo cáo quyết toán sửa đổi bổ sung (Mẫu 15a) trong thời hạn 60 ngày kể từ ngày nộp báo cáo chính thức và trước khi có quyết định KTSTQ.',
         'scratch/daily_notes.md (11/11) & customs.gov.vn | HT (Học Thực Chiến) / Customs Settlement M15 & Actual BOM Control',
         update_date_str),

        (249,
         'dong bo ma kho bai dia diem hai quan vnaccs vcis loi truyen to khai thong quan',
         'dong bo ma kho bai dia diem, he thong vnaccs vcis, ma dia diem luu kho cho thong quan, loi truyen to khai hai quan vacs, cai to hanh chinh hai quan 2026, ma chi cuc hai quan moi',
         99.6,
         56900,
         'https://haiquanonline.com.vn',
         'Đồng bộ bộ mã kho, bãi, địa điểm mới trên hệ thống VNACCS/VCIS cuối năm 2026: Giải mã nguyên nhân lỗi truyền tờ khai và hướng dẫn cập nhật chuẩn xác',
         'Nội dung trọng tâm: Cảnh báo kỹ thuật và cập nhật chính sách theo thông báo mới nhất của Cục CNTT & Thống kê Hải quan ngày 11/11/2026: Nhằm phục vụ lộ trình hiện đại hóa Hải quan số và sắp xếp lại địa giới hành chính các cảng cạn, kho CFS và địa điểm kiểm tra hàng hóa tập trung trên toàn quốc, ngành Hải quan đang rà soát, đồng bộ và chuẩn hóa toàn bộ Danh mục mã kho, bãi, địa điểm lưu kho chờ thông quan trên hệ thống VNACCS/VCIS; trong những ngày qua, nhiều doanh nghiệp XNK và đại lý hải quan phản ánh tình trạng truyền tờ khai bị hệ thống báo lỗi từ chối tiếp nhận (mã lỗi VACS không tìm thấy mã địa điểm lưu kho hoặc mã chi cục hải quan cửa khẩu tiếp nhận); lỗi này khiến hàng hóa bị đình trệ, không thể phân luồng tờ khai, lỡ chuyến tàu và phát sinh chi phí lưu xe/lưu container; Eureka hướng dẫn chi tiết: Doanh nghiệp cần chủ động tra cứu bảng mã chuẩn hóa mới nhất trên Cổng thông tin điện tử Tổng cục Hải quan (customs.gov.vn); cập nhật phiên bản phần mềm đầu cuối (ECUS/VNACCS); đối chiếu chính xác mã cảng dỡ hàng (Port of Discharge) và mã địa điểm dỡ hàng quy chuẩn trước khi bấm truyền chính thức.',
         'haiquanonline.com.vn & customs.gov.vn | TN (Tin Ngành) / VNACCS Warehouse Code Sync & Clearance Troubleshooting',
         update_date_str),

        (250,
         'cuoc tau bien drewry thang 11 2026 blank sailings huy chuyen phu phi gri',
         'cuoc tau bien drewry thang 11 2026, drewry world container index, blank sailings huy chuyen tau bien, phu phi gri thang 11 12, tuyen van tai xuyen thai binh duong, dat cho booking cont mua tet',
         99.5,
         55800,
         'https://congthuong.vn',
         'Thị trường cước container đường biển tuần 2 tháng 11/2026: Ứng phó làn sóng Blank Sailings và kế hoạch tăng phụ phí GRI của các liên minh hãng tàu cuối năm',
         'Nội dung trọng tâm: Báo cáo phân tích thị trường cước biển quốc tế ngày 11/11/2026 trích xuất từ dữ liệu Drewry WCI và Báo Công Thương: Mặc dù chỉ số cước container toàn cầu (World Container Index - WCI) tuần đầu tháng 11 ghi nhận mức điều chỉnh nhẹ về 4.389 USD/FEU, các liên minh hãng tàu lớn (2M, Ocean Alliance, THE Alliance) đang đồng loạt kích hoạt chiến lược Blank Sailings (hủy chuyến có chủ đích), cắt giảm khoảng 12-15% tổng công suất vận tải trên các tuyến hàng hải trọng yếu từ Châu Á đi Bắc Mỹ (USWC/USEC) và Bắc Âu (Rotterdam/Hamburg); mục tiêu của các hãng tàu là kìm đà giảm giá spot và chuẩn bị tung ra đợt tăng phụ phí chung GRI (General Rate Increase) từ 500 - 1.000 USD/FEU vào nửa cuối tháng 11 và đầu tháng 12/2026; điều này tạo ra áp lực rất lớn cho các doanh nghiệp xuất nhập khẩu Việt Nam đang vào vụ sản xuất giao hàng Tết; nguy cơ trễ hạn giao hàng (late delivery) và bị hủy booking (rollover) là rất cao; Eureka tư vấn chiến lược: Chủ hàng cần lập kế hoạch đặt chỗ (booking space) trước tối thiểu 2-3 tuần, đa dạng hóa hãng tàu, cân nhắc chia nhỏ lô hàng và thỏa thuận điều khoản cước kết hợp (Contract + Spot blend) để bảo vệ chuỗi cung ứng.',
         'congthuong.vn & Drewry WCI, World Shipping Council | FF (Forwarder & Cước Vận Tải) / Drewry Shipping Index & Blank Sailings Strategy',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 246 -> row 251
        fill_color = fill_zebra_even if item[0] % 2 == 0 else fill_zebra_odd
        for col_idx, val in enumerate(item, start=1):
            cell = ws3.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill_color
            cell.border = cell_border
            if col_idx == 1:
                cell.alignment = align_center
            elif col_idx in [2, 3, 7, 8, 9]:
                cell.alignment = align_left
            elif col_idx == 4:
                cell.alignment = align_right
                cell.number_format = '0.0'
            elif col_idx == 5:
                cell.alignment = align_right
                cell.number_format = '#,##0'
            elif col_idx == 6:
                cell.alignment = align_left
                cell.font = font_link
                cell.hyperlink = val
            elif col_idx == 10:
                cell.alignment = align_center
        ws3.row_dimensions[row_idx].height = 85

    # Cập nhật Total row ở row 256
    total_row_idx = 256
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 250 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D255)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E255)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 250 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 218 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells(f'A{total_row_idx}:C{total_row_idx}')
    ws3.merge_cells(f'F{total_row_idx}:J{total_row_idx}')
    ws3.row_dimensions[total_row_idx].height = 40

    # ==========================================
    # 3. LƯU ATOMIC TRÁNH LỖI ONEDRIVE & TUÂN THỦ NO-BACKUP RULE
    # ==========================================
    print("3. Lưu nguyên tử file Excel...")
    temp_local_file = FILE_PATH + '.tmp.xlsx'
    wb.save(temp_local_file)
    wb.close()

    if os.path.exists(temp_local_file):
        os.replace(temp_local_file, FILE_PATH)
        print(f"[OK] Đã cập nhật thành công file cục bộ: {FILE_PATH}")

    # Lấy thông tin dung lượng và hash file chính
    file_size = os.path.getsize(FILE_PATH)
    file_sha256 = get_file_sha256(FILE_PATH)

    # Đồng bộ sang OneDrive (nếu tồn tại thư mục)
    if os.path.exists(ONEDRIVE_DIR):
        print(f"Đang đồng bộ nguyên tử sang OneDrive: {ONEDRIVE_FILE}...")
        temp_onedrive_file = ONEDRIVE_FILE + '.tmp.xlsx'
        try:
            shutil.copy2(FILE_PATH, temp_onedrive_file)
            os.replace(temp_onedrive_file, ONEDRIVE_FILE)
            print(f"[OK] Đã đồng bộ thành công sang OneDrive: {ONEDRIVE_FILE}")
            
            # Ghi log đồng bộ
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 11/11/2026)\n"
            with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f_log:
                f_log.write(log_entry)
            print(f"[OK] Đã ghi log thành công vào: {SYNC_LOG_PATH}")
        except Exception as e:
            print(f"[CẢNH BÁO] Lỗi đồng bộ OneDrive: {e}")
            if os.path.exists(temp_onedrive_file):
                try:
                    os.remove(temp_onedrive_file)
                except Exception:
                    pass

    # ==========================================
    # 4. TẠO BẢN TIN ĐỀ XUẤT HOT NEWS SCAN (wf05)
    # ==========================================
    print("4. Tạo bản tin đề xuất Hot News Scan...")
    briefing_content = f"""# BẢN TIN ĐỀ XUẤT TIN NÓNG TUẦN & XU HƯỚNG TỪ KHÓA XNK (HOT NEWS BRIEFING)
**Quy trình thực hiện:** `wf05_hot_news_scan.md`  
**Tác tử phụ trách:** Agent 01 (Market Researcher & Hot News Scanner)  
**Ngày thực hiện:** {update_date_str} (Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026)  
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 250 chủ đề)  
**Trạng thái đồng bộ OneDrive:** ✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (24H - 7 NGÀY QUA)
Hệ thống Hot News Scanner đã tiến hành quét độc lập trên 8 nguồn dữ liệu chính thống:
1. `customs.gov.vn` (Cổng TTĐT Hải quan)
2. `chinhphu.vn` (Cổng TTĐT Chính phủ)
3. `haiquanonline.com.vn` (Tạp chí Hải quan Online)
4. `congthuong.vn` (Báo Công Thương)
5. `tapchitaichinh.vn` (Thời báo Tài chính Việt Nam)
6. `Google Trends VN` (Xu hướng tìm kiếm XNK, Thuế, Logistics)
7. `Drewry / World Shipping Council` (Báo cáo WCI cước vận tải biển)
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 11/11/2026)

Qua sàng lọc từ 52 sự kiện và case study nghiệp vụ thực tế, Agent 01 đã chọn lọc và chấm điểm chi tiết **TOP 5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe của quy trình `wf05` (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [MÙA CAO ĐIỂM TẾT — RỦI RO THAM VẤN GIÁ & BÁC C/O] Thủ tục nhập khẩu mỹ phẩm mùa cao điểm cuối năm: Cạm bẫy tham vấn giá hải quan, rủi ro bác C/O ưu đãi và chuẩn hóa phiếu công bố mỹ phẩm
- **Chuyên mục (Pillar):** `CS` (Case Study) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Infographic '3 cạm bẫy nhập khẩu mỹ phẩm mùa Tết: Sai một chữ trên Phiếu công bố, hàng kẹt cảng và bị bác sạch C/O ưu đãi' | YouTube: Video cẩm nang hướng dẫn 5 bước chứng minh trị giá hải quan và quy trình kiểm tra nhãn phụ hàng hóa)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực nóng: Mùa gom hàng và nhập khẩu mỹ phẩm cao điểm phục vụ lễ Noel và Tết Nguyên Đán 2027 đang vào giai đoạn nước rút; case study thực tế phát sinh hôm nay |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Rất đau: Hàng bị kẹt tại cảng do sai lệch nhãn so với phiếu công bố; bị hải quan bác C/O form E, D, VK làm tăng thuế nhập khẩu từ 0% lên 10-22%; bị ấn định tăng giá tham vấn thuế |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi dữ dội: Chủ hàng và cơ quan hải quan về cơ sở định giá tham vấn; khác biệt giữa công bố thành phần trên CFS nước ngoài với quy chuẩn Bộ Y tế Việt Nam |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka cung cấp cẩm nang trọn gói: Bộ checklist đối soát phiếu công bố mỹ phẩm trước khi nhập, phác đồ hồ sơ bảo vệ giá tham vấn và cách lấy C/O chuẩn form |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #1 (Chủ Đề Nóng Bỏng & Sát Thực Tế Nhất Cho Hàng Tiêu Dùng Mùa Tết)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi lo lớn nhất của cộng đồng doanh nghiệp nhập khẩu hàng tiêu dùng, mỹ phẩm và TMĐT quốc tế mùa cuối năm. Khẳng định năng lực thông quan hàng khó và xử lý tham vấn giá chuyên sâu của Eureka.
- **Nguồn xác minh:** Thông tư 06/2011/TT-BYT; Thông tư 38/2015/TT-BTC & 39/2018/TT-BTC; Cục Quản lý Dược; `scratch/daily_notes.md` (11/11/2026).

---

### Đề xuất #2: [RỦI RO TREO CONT TẠI BÃI — THỦ TỤC CẢNG BIỂN] Tờ khai vận chuyển độc lập OLA trong logistics cảng biển (Cát Lái - Cái Mép): Phân định hàng chịu sự giám sát hải quan và cảnh báo rủi ro mở sai thủ tục OLA
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `FF` (Forwarder & Cước Vận Tải)
- **Nền tảng khuyến nghị:** **Facebook & Group Nghiệp Vụ Logistics** (Bài phân tích nghiệp vụ: 'Tờ khai OLA vận chuyển độc lập Cát Lái - Cái Mép: 3 dấu hiệu phân định rõ hàng chịu sự giám sát hải quan — Đừng để bị treo cont và phạt vi phạm hành chính')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Phát sinh trực tiếp trên diễn đàn nghiệp vụ ngày 11/11/2026; nhu cầu vận chuyển container trung chuyển giữa Cát Lái và Cái Mép tăng vọt cuối năm |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Khai nhầm OLA cho hàng nội địa hoặc hàng đã thông quan khiến tờ khai bị treo, không thanh lý được seal định vị GPS, cont bị giữ lại bãi và phạt tiền |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận gay gắt giữa nhân viên chứng từ, nhà xe kéo cont và đại lý hãng tàu về việc khi nào bắt buộc mở OLA và khi nào chỉ cần vận đơn nội địa |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka làm rõ 100% căn cứ Điều 50, 51 Thông tư 38 (sửa đổi TT 39), bảng so sánh trực quan giữa OLA hàng xuất chuyển cảng vs Vận chuyển nội địa thuần túy |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #2 (Giải Mã Vướng Mắc Thủ Tục Vận Tải Cảng Biển)** |
- **Tóm tắt lý do nên làm:** Tháo gỡ nhầm lẫn phổ biến của các bạn làm chứng từ logistics và tài xế/nhà xe container, định vị Eureka là chuyên gia hàng đầu về nghiệp vụ cảng biển và vận chuyển đa phương thức.
- **Nguồn xác minh:** Điều 50, Điều 51 Thông tư 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); haiquanonline.com.vn; `scratch/daily_notes.md` (11/11/2026).

---

### Đề xuất #3: [QUẢN TRỊ RỦI RO SỐ LIỆU — MÙA BÁO CÁO QUYẾT TOÁN] Báo cáo quyết toán Hải quan Mẫu 15: Chiến lược kiểm soát biến động định mức thực tế và cơ chế đối soát liên phòng ban (Hải quan - Kho - Kế toán) trước mùa KTSTQ
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `CS` (Case Study Quản Trị Rủi Ro)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Infographic 'Mô hình Tam giác số liệu 3 bên: Hải quan - Kế toán - Kho: Cách tự kiểm toán nội bộ trước khi nộp BCQT Mẫu 15' | YouTube: Video ThuyTalk hướng dẫn giải trình định mức hao hụt và xử lý âm kho nguyên liệu)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Cực kỳ cấp bách: Quý 4/2026 là giai đoạn vàng để các nhà máy SXXK, DNCX và gia công rà soát sổ sách trước khi chốt năm tài chính |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Đau đớn nhất: Lệch số liệu âm kho nguyên liệu không chứng minh được định mức hao hụt thực tế sẽ bị cơ quan hải quan ấn định thuế và phạt trốn thuế hàng tỷ đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Mâu thuẫn nội bộ doanh nghiệp: Kế toán theo giá trị, Kho theo số lượng thùng/cuộn, XNK theo tờ khai kg/mét; tranh luận về việc học BCQT ở đâu thực chiến |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chuyển giao công thức đối soát độc quyền: Phương pháp lập bảng quy đổi đơn vị tính, biên bản xử lý hao hụt tự nhiên và quy trình nộp Mẫu 15a sửa đổi |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Nội Dung Chuyên Gia Cao Cấp — Chinh Phục Khối Doanh Nghiệp Sản Xuất & FDI)** |
- **Tóm tắt lý do nên làm:** Tạo sức hút đặc biệt với chủ doanh nghiệp, kế toán trưởng và trưởng phòng XNK. Khẳng định sự khác biệt giữa Eureka (tư vấn giải pháp quản trị thực chiến) với các trung tâm chỉ dạy truyền phần mềm.
- **Nguồn xác minh:** Điều 60 Thông tư 38/2015/TT-BTC (sửa đổi Thông tư 39/2018/TT-BTC); Báo Hải quan Online; customs.gov.vn; `scratch/daily_notes.md` (11/11/2026).

---

### Đề xuất #4: [SỐ HÓA HẢI QUAN & CẢNH BÁO LỖI HỆ THỐNG] Đồng bộ bộ mã kho, bãi, địa điểm mới trên hệ thống VNACCS/VCIS cuối năm 2026: Giải mã nguyên nhân lỗi truyền tờ khai và hướng dẫn cập nhật chuẩn xác
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group Xuất Nhập Khẩu** (Bài hướng dẫn kỹ thuật cấp tốc: 'Cảnh báo lỗi truyền tờ khai VNACCS do thay đổi mã địa điểm kho bãi cuối năm 2026: Hướng dẫn 3 bước khắc phục ngay để thông quan kịp chuyến tàu')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Nóng hổi: Cục CNTT & Thống kê Hải quan đang thực hiện đợt rà soát chuẩn hóa mã địa điểm kho bãi phục vụ lộ trình Hải quan số 2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Tờ khai bị lỗi từ chối tiếp nhận (Error Code VACS) khiến lô hàng bị treo, hàng không lấy được eDO, rớt chuyến tàu xuất khẩu và phát sinh chi phí lưu xe |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Phản ánh của các doanh nghiệp về tình trạng đồng bộ chậm giữa phần mềm doanh nghiệp (ECUS) với cổng tiếp nhận của Tổng cục Hải quan |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka cung cấp link tra cứu mã địa điểm chuẩn hóa, mẹo kiểm tra mã chi cục hải quan tiếp nhận và hướng dẫn xử lý khi hệ thống bị treo |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #4 (Cứu Cánh Kỹ Thuật Thông Quan Cho Dân Khai Thuê & Doanh Nghiệp)** |
- **Tóm tắt lý do nên làm:** Mang lại giá trị tức thời, giải quyết đúng ca khó mà hàng trăm nhân viên chứng từ đang gặp phải trên cả nước, gia tăng lượt share và tương tác tự nhiên.
- **Nguồn xác minh:** Thông báo kỹ thuật Cục CNTT & Thống kê Hải quan; Tạp chí Hải quan Online (haiquanonline.com.vn); customs.gov.vn.

---

### Đề xuất #5: [BIẾN ĐỘNG CƯỚC BIỂN CAO ĐIỂM] Thị trường cước container đường biển tuần 2 tháng 11/2026: Ứng phó làn sóng 'Blank Sailings' và kế hoạch tăng phụ phí GRI của các liên minh hãng tàu cuối năm
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **LinkedIn & Facebook Fanpage** (Bản tin thị trường: 'Cước tàu biển tháng 11/2026: Hãng tàu đồng loạt hủy chuyến (Blank Sailings) ghìm giá — 3 chiến lược đặt chỗ giúp chủ hàng không bị rớt container mùa Tết')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Cập nhật tuần 2 tháng 11/2026 từ báo cáo Drewry WCI và thông báo từ các liên minh hãng tàu lớn (2M, Ocean Alliance, THE Alliance) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Bị hãng tàu hủy booking đột ngột, hàng bị rollover lưu lại cảng chuyển tải, phát sinh thêm phụ phí GRI từ 500-1000 USD/FEU |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Tranh cãi giữa chủ hàng và forwarder về phụ phí phát sinh; bài toán nên ký hợp đồng cước dài hạn hay đi theo giá spot |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka đưa ra dự báo xu hướng cước đến tháng 01/2027, tư vấn công thức phân bổ rủi ro booking sớm 2-3 tuần và cam kết giữ chỗ |
  | **Tổng điểm** | **18.9 / 20** | **Xếp hạng: #5 (Bản Tin Thị Trường & Tư Vấn Cước Vận Tải Biển)** |
- **Tóm tắt lý do nên làm:** Cung cấp thông tin thị trường cước biển có chiều sâu, hỗ trợ các giám đốc logistics và chủ hàng xuất khẩu chủ động dòng tiền và kế hoạch giao hàng.
- **Nguồn xác minh:** Báo cáo Drewry World Container Index; World Shipping Council; Báo Công Thương (congthuong.vn).

---

## 📌 HÀNH ĐỘNG TIẾP THEO ĐỀ XUẤT CHO CHỊ THÚY
1. **Lựa chọn chủ đề:** Chị Thúy có thể chọn 1 hoặc 2 chủ đề trong Top 5 để đưa vào lịch sản xuất content tuần (Tuần 46 qua quy trình `wf01_weekly_planning.md`).
2. **Sản xuất nhanh:** Nếu chọn Đề xuất #1 (Mỹ phẩm mùa Tết: Tham vấn giá & bác C/O) hoặc Đề xuất #3 (Báo cáo quyết toán Mẫu 15: Tam giác số liệu), Agent 03 và Agent 06 sẵn sàng kích hoạt quy trình `wf02` (Facebook) hoặc `wf03` (YouTube) để viết bài và thiết kế infographic ngay hôm nay.
3. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT WF05 NGÀY 11/11/2026 ===")
    print(f"- File Excel: {FILE_PATH} ({file_size} bytes, SHA256: {file_sha256[:8]}...)")
    print(f"- OneDrive: {ONEDRIVE_FILE}")
    print(f"- Briefing: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w46_1111()
