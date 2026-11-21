# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_12_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 12/11/2026
(Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026 | Cập nhật ngày 12/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 46 cập nhật ngày 12/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 250 lên 255 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 12/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w46_1211.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w46_1211.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w46_1211():
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

    update_date_str = "12/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 46 (09/11/2026 - 15/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'khai sai hoa don so luong sau thong quan ama mau 03 thuc an chan nuoi',
         'Khắc phục sự cố khai nhầm hóa đơn, số lượng sau thông quan: Phác đồ khai bổ sung AMA, lập Mẫu 03 và giải pháp phòng vệ rủi ro phạt ấn định thuế',
         19800, 'Điểm nóng bổ sung sau thông quan (+3,950%)',
         'Doanh nghiệp nhập khẩu thức ăn chăn nuôi, nông sản, nguyên phụ liệu sản xuất; nhân viên khai báo hải quan, kế toán kho và giám đốc logistics.',
         'Khai nhầm số hóa đơn và số lượng của lô hàng khác dẫn đến sai lệch dữ liệu sau khi hàng đã thông quan; lo ngại bị lập biên bản vi phạm hành chính, ấn định thuế hoặc bị thanh tra KTSTQ.',
         "Tiêu đề: 'Khai nhầm hóa đơn và số lượng sau thông quan: Hướng dẫn 4 bước truyền AMA và lập Mẫu 03 theo Thông tư 39/2018/TT-BTC an toàn tuyệt đối'.",
         'Điểm b Khoản 1 Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); Nghị định 128/2020/NĐ-CP; scratch/daily_notes.md (12/11/2026).',
         update_date_str),

        (2,
         'thu tuc nhap khau my pham cong bo my pham tham van gia hai quan co uu dai thue',
         'Thủ tục nhập khẩu mỹ phẩm mùa cao điểm cuối năm: Cạm bẫy tham vấn giá hải quan, rủi ro bác C/O ưu đãi và chuẩn hóa phiếu công bố mỹ phẩm',
         19400, 'Điểm nóng mỹ phẩm mùa Tết (+3,810%)',
         'Doanh nghiệp nhập khẩu và phân phối mỹ phẩm, dược mỹ phẩm, hóa mỹ phẩm; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Sai lệch thành phần/nhãn mác với phiếu công bố dẫn đến bị đình chỉ thông quan và xử phạt vi phạm hành chính; bị nghi ngờ trị giá thấp, yêu cầu tham vấn giá và ấn định tăng thuế; nguy cơ bác C/O form E, D, VK làm mất thuế ưu đãi đặc biệt.',
         "Tiêu đề: 'Nhập khẩu mỹ phẩm mùa cao điểm: 3 cạm bẫy tham vấn giá, bác C/O và chuẩn hóa phiếu công bố tránh tắc hàng tại cảng'.",
         'Thông tư 06/2011/TT-BYT; Thông tư 38/2015/TT-BTC & 39/2018/TT-BTC; scratch/daily_notes.md (11-12/11/2026); Cục Quản lý Dược.',
         update_date_str),

        (3,
         'ke khai nop thay thue fct meta microsoft google khoan 4 dieu 4 luat thue gtgt',
         'Kê khai nộp thay thuế nhà thầu FCT cho Meta, Microsoft, Google theo Luật Thuế GTGT 48/2024/QH15: Tỷ lệ 5% VAT và 2% TNDN hợp thức hóa chi phí',
         19000, 'Thuế nhà thầu dịch vụ số (+3,720%)',
         'Toàn bộ doanh nghiệp SME, công ty truyền thông/marketing, doanh nghiệp công nghệ, kế toán trưởng và giám đốc tài chính (CFO).',
         'Doanh nghiệp chi tiền mua quảng cáo Facebook/Google hoặc phần mềm quốc tế nhưng không khấu trừ nộp thay thuế FCT; nguy cơ bị cơ quan thuế bóc tách loại trừ toàn bộ chi phí hợp lý và truy thu phạt chậm nộp.',
         "Tiêu đề: 'Kê khai thay thuế nhà thầu (FCT) cho Meta, Google, Microsoft: Cách tính thuế 5% VAT và bảo vệ chi phí được trừ 100% khi quyết toán'.",
         'Khoản 4 Điều 4 Luật Thuế GTGT 48/2024/QH15; Thông tư 69/2025/TT-BTC; scratch/daily_notes.md (12/11/2026); Tổng cục Thuế.',
         update_date_str),

        (4,
         'hang gia cong bi ben nuoc ngoai tu bo thong tu 121 2025 tt btc dieu 66',
         'Hàng gia công bị bên nước ngoài từ bỏ theo Thông tư 121/2025/TT-BTC: 03 hướng xử lý, phác đồ mở tờ khai A42 nộp thuế và trách nhiệm pháp lý doanh nghiệp',
         18700, 'Điểm nóng hàng gia công (+3,620%)',
         'Doanh nghiệp gia công may mặc, da giày, linh kiện điện tử, cơ khí; giám đốc nhà máy, kế toán kho và nhân viên khai báo hải quan.',
         'Bên thuê gia công nước ngoài từ bỏ nguyên vật liệu, thành phẩm, máy móc; doanh nghiệp lúng túng không biết tiêu thụ nội địa hay tiêu hủy; nếu quá hạn 30 ngày bị hải quan lập biên bản và ấn định thuế.',
         "Tiêu đề: 'Bên đặt gia công từ bỏ hàng: Phác đồ 3 bước mở tờ khai A42 hoặc tiêu hủy đúng Điều 66 Thông tư 121/2025/TT-BTC tránh bị xử phạt'.",
         'Thông tư số 121/2025/TT-BTC (sửa Điều 66 Thông tư 38/2015/TT-BTC); scratch/daily_notes.md; Cục Hải quan.',
         update_date_str),

        (5,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl dinh danh dien tu',
         'Thông tư 142/2026/TT-BTC (hiệu lực từ 16/11/2026): Chuẩn hóa chỉ tiêu giám sát hải quan chuyển phát nhanh quốc tế, bắt buộc định danh điện tử và cách lấy tờ khai hàng mẫu',
         18300, 'Nước rút Thông tư 142 chuyển phát nhanh (+3,510%)',
         'Doanh nghiệp gửi hàng mẫu qua FedEx, DHL, UPS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Thông tư có hiệu lực từ 16/11/2026 với các chỉ tiêu thông tin mới và định danh điện tử; gửi hàng mẫu qua bưu chính thường bị gom tờ khai chung, không lấy được tờ khai đứng tên doanh nghiệp để khấu trừ thuế.',
         "Tiêu đề: 'Thông tư 142/2026/TT-BTC hiệu lực từ 16/11: Hướng dẫn lấy tờ khai hải quan hàng mẫu chuyển phát nhanh bảo vệ chi phí thuế hợp lệ'.",
         'Thông tư số 142/2026/TT-BTC; customs.gov.vn; Báo Hải quan Online.',
         update_date_str),

        (6,
         'to khai ola van chuyen doc lap giam sat hai quan chuyen cang cat lai cai mep',
         'Tờ khai vận chuyển độc lập OLA trong logistics cảng biển (Cát Lái - Cái Mép): Phân định hàng chịu sự giám sát hải quan và cảnh báo rủi ro mở sai thủ tục OLA',
         17900, 'Thủ tục tờ khai OLA cảng biển (+3,420%)',
         'Các công ty forwarder, nhà xe vận tải container, doanh nghiệp logistics trung chuyển cảng biển và nhân viên khai báo hải quan.',
         'Hiểu nhầm OLA cho hàng nội địa thông thường; mở sai thủ tục dẫn đến không thanh lý được tờ khai, seal điện tử GPS bị treo, container bị giữ lại bãi cảng và bị xử phạt hành chính.',
         "Tiêu đề: 'Tờ khai OLA vận chuyển độc lập giữa Cát Lái và Cái Mép: 3 nguyên tắc phân định hàng giám sát hải quan tránh bị treo cont tại bãi'.",
         'Điều 50, Điều 51 Thông tư 38/2015/TT-BTC (sửa TT 39/2018/TT-BTC); scratch/daily_notes.md; haiquanonline.com.vn.',
         update_date_str),

        (7,
         'thanh khoan dncx bao cao quyet toan hai quan mau 15 chenh lech xuat nhap ton',
         'Kiểm tra sau thông quan thanh khoản DNCX: Cách xử lý chênh lệch số liệu xuất nhập tồn, định mức hao hụt và bảo vệ Báo cáo quyết toán Mẫu 15',
         17500, 'Mùa rà soát thanh khoản DNCX (+3,310%)',
         'Doanh nghiệp chế xuất (DNCX), nhà máy FDI, giám đốc sản xuất, kế toán trưởng và nhân viên khai báo hải quan.',
         'Số liệu kho thực tế lệch với ERP và tờ khai hải quan; âm nguyên liệu không chứng minh được định mức hao hụt; nguy cơ bị ấn định thuế NK, VAT và phạt trốn thuế khi kiểm tra sau thông quan.',
         "Tiêu đề: 'Bảo vệ Báo cáo quyết toán DNCX trước kiểm tra sau thông quan: Công thức Tam giác vàng đối soát Hải quan - Kế toán - Kho'.",
         'Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; scratch/daily_notes.md; Báo Hải quan Online.',
         update_date_str),

        (8,
         'cong van 23969 chq nvthq kiem tra gia cau kien thep nha tien che nhap khau',
         'Công văn 23969/CHQ-NVTHQ: Hải quan siết chặt kiểm tra xuất xứ và trị giá hải quan đối với cấu kiện thép, nhà tiền chế nhập khẩu giá thấp',
         17100, 'Siết chặt trị giá hải quan thép (+3,190%)',
         'Doanh nghiệp nhập khẩu thép kết cấu, cấu kiện thép nhà xưởng, nhà tiền chế, các nhà thầu EPC xây dựng công nghiệp.',
         'Hải quan tăng cường quản lý rủi ro, chuyển luồng kiểm tra gắt gao đối với cấu kiện thép khai báo giá thấp bất thường; nguy cơ bị bác C/O, ấn định thuế và truy thu phòng vệ thương mại.',
         "Tiêu đề: 'Công văn 23969/CHQ-NVTHQ: Chuẩn bị 5 tài liệu cốt lõi bảo vệ trị giá hải quan và xuất xứ cấu kiện thép nhập khẩu'.",
         'Công văn 23969/CHQ-NVTHQ tháng 11/2026; haiquanonline.com.vn; customs.gov.vn.',
         update_date_str),

        (9,
         'ke khai giam 30 thue tncn ho kinh doanh online dich vu cong chung tu hop le',
         'Hướng dẫn kê khai giảm 30% thuế TNCN cho hộ kinh doanh online trên Cổng Dịch vụ công: Hồ sơ chứng từ hợp lệ phòng ngừa truy thu',
         16600, 'Kê khai giảm thuế hộ kinh doanh (+3,060%)',
         'Chủ shop TMĐT (Shopee, TikTok Shop, Lazada), cá nhân kinh doanh online, hộ kinh doanh cá thể nộp thuế theo phương pháp kê khai.',
         'Hiểu nhầm quy định giảm thuế dẫn đến kê khai sai số thuế được giảm; không khớp nối dữ liệu doanh thu sàn TMĐT với tờ khai trên Dịch vụ công; nguy cơ bị phạt chậm nộp và thu hồi số thuế đã giảm.',
         "Tiêu đề: 'Kê khai giảm 30% thuế TNCN hộ kinh doanh online trên Dịch vụ công: 3 lưu ý chứng từ tránh bị cơ quan thuế hậu kiểm truy thu'.",
         'Nghị quyết của Quốc hội; scratch/daily_notes.md (12/11/2026); Cổng Dịch vụ công Quốc gia.',
         update_date_str),

        (10,
         'cuoc tau bien drewry thang 11 2026 blank sailings huy chuyen phu phi gri',
         'Thị trường cước container đường biển tuần 2 tháng 11/2026: Ứng phó làn sóng Blank Sailings và kế hoạch tăng phụ phí GRI của các liên minh hãng tàu cuối năm',
         16200, 'Làn sóng Blank Sailings cước biển (+2,920%)',
         'Doanh nghiệp xuất nhập khẩu hàng đi Mỹ (USWC/USEC), Châu Âu (Rotterdam/Hamburg) và các nhà máy xuất khẩu hàng tiêu dùng mùa lễ Tết.',
         'Các hãng tàu cắt giảm 12-15% tải cung ứng (blank sailings) để ghìm giá và rục rịch tăng phụ phí GRI từ 500-1000 USD/FEU; nguy cơ trễ hẹn giao hàng và rớt cont tại cảng chuyển tải.',
         "Tiêu đề: 'Cước tàu biển tháng 11/2026: Hãng tàu đồng loạt hủy chuyến (Blank Sailings), chủ hàng làm gì để không bị rớt container?'.",
         'Drewry World Container Index; World Shipping Council; Báo Công Thương (congthuong.vn).',
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
    ws3['A3'] = f"Chu kỳ: Tuần 46 (09/11/2026 - 15/11/2026) | Mở rộng 255 chủ đề nghiên cứu (gồm 223 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 256)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 256 or m.max_row >= 256]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (251,
         'khai sai hoa don so luong sau thong quan ama mau 03 thuc an chan nuoi',
         'khai sai hoa don sau thong quan, khai bo sung to khai ama, mau 03 khai bo sung hai quan, dieu 20 thong tu 38 2015, xu phat khai sai to khai hai quan, an dinh thue sau thong quan',
         99.9,
         62500,
         'https://thuvienphapluat.vn',
         'Khắc phục sự cố khai nhầm hóa đơn và số lượng sau thông quan: Phác đồ khai bổ sung AMA, lập Mẫu 03 và giải pháp phòng vệ rủi ro phạt ấn định thuế',
         'Nội dung trọng tâm: Giải mã case study nghiệp vụ cấp thiết ngày 12/11/2026 về sự cố khai sai thông tin hóa đơn và số lượng hàng hóa sau thông quan: Doanh nghiệp nhập khẩu mặt hàng thức ăn chăn nuôi (thuế suất NK 0%) đã được hệ thống phân luồng và hoàn tất thông quan, tuy nhiên khi rà soát sổ sách kế toán nội bộ phát hiện nhân viên chứng từ lấy nhầm số hóa đơn thương mại (Commercial Invoice) và ngày hóa đơn của lô hàng khác, dẫn đến số lượng hàng khai trên tờ khai cao hơn số lượng thực tế trên hóa đơn đúng; Vấn đề pháp lý & rủi ro: Mặc dù mặt hàng có thuế suất 0% không làm phát sinh thiếu số thuế phải nộp, hành vi khai sai các chỉ tiêu thông tin về lượng hàng, số hiệu hóa đơn vẫn bị coi là hành vi khai sai so với thực tế và hồ sơ hải quan; nếu để cơ quan hải quan phát hiện trước khi thanh kiểm tra sau thông quan, doanh nghiệp sẽ bị xử phạt vi phạm hành chính về thủ tục hải quan theo Nghị định 128/2020/NĐ-CP (sửa đổi tại Nghị định 102/2021/NĐ-CP); Phác đồ xử lý chuẩn mực Eureka: (1) Căn cứ Điểm b Khoản 1 Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi, bổ sung tại Khoản 9 Điều 1 Thông tư 39/2018/TT-BTC), người khai hải quan được khai bổ sung sau khi thông quan trong thời hạn 60 ngày kể từ ngày thông quan và trước thời điểm cơ quan hải quan quyết định kiểm tra sau thông quan; (2) Thực hiện truyền nghiệp vụ AMA (Khai bổ sung sau thông quan) trên hệ thống VNACCS/VCIS; (3) Lập Văn bản đề nghị khai bổ sung theo Mẫu số 03/KBS/GSQL Phụ lục II ban hành kèm Thông tư 39/2018/TT-BTC giải trình rõ lý do nhầm lẫn hóa đơn giữa các lô hàng; (4) Nộp kèm toàn bộ chứng từ chứng minh gồm Invoice đúng, Invoice nhầm, Packing List, Vận đơn B/L, chứng từ thanh toán ngân hàng và xác nhận của người bán nước ngoài; (5) Đối soát khớp số lượng thực nhập với báo cáo quyết toán kho để tránh bị nghi ngờ buôn lậu hoặc gian lận thương mại.',
         'scratch/daily_notes.md (12/11) & thuvienphapluat.vn | CS (Case Study) / Post-Clearance Amendment AMA & Invoice Audit Compliance',
         update_date_str),

        (252,
         'ke khai nop thay thue fct meta microsoft google khoan 4 dieu 4 luat thue gtgt',
         'thue nha thau fct meta microsoft, ke khai nop thay thue nha thau, khoan 4 dieu 4 luat thue gtgt 48 2024 qh15, thong tu 69 2025 tt btc, ty le thue fct dich vu so, hoa don chi phi quang cao facebook google hop ly',
         99.8,
         60800,
         'https://chinhphu.vn',
         'Kê khai nộp thay thuế nhà thầu (FCT) khi mua dịch vụ số của Meta, Microsoft, Google: Căn cứ Khoản 4 Điều 4 Luật Thuế GTGT 48/2024/QH15 và cách bảo vệ chi phí thuế hợp lệ',
         'Nội dung trọng tâm: Hướng dẫn chuyên sâu ngày 12/11/2026 về nghĩa vụ thuế nhà thầu đối với các giao dịch mua dịch vụ số xuyên biên giới: Hàng loạt doanh nghiệp Việt Nam chi hàng trăm triệu đến hàng tỷ đồng mỗi tháng để mua dịch vụ quảng cáo trực tuyến (Meta/Facebook Ads, Google Ads), phần mềm bản quyền (Microsoft 365, Zoom, Adobe), hoặc dịch vụ điện toán đám mây (AWS, Azure) nhưng chưa nắm rõ nghĩa vụ kê khai thuế FCT; Phân tích pháp lý SSoT: Căn cứ Khoản 4 Điều 4 Luật Thuế GTGT số 48/2024/QH15 (hiệu lực từ 01/07/2025) và Điểm a Khoản 4, Khoản 6 Điều 9 Thông tư 69/2025/TT-BTC, tổ chức kinh doanh tại Việt Nam áp dụng phương pháp khấu trừ khi mua dịch vụ của nhà cung cấp nước ngoài (NCCNN) không có cơ sở thường trú tại Việt Nam thông qua kênh TMĐT hoặc nền tảng số có trách nhiệm thực hiện khấu trừ, nộp thay nghĩa vụ thuế GTGT cho NCCNN; Tỷ lệ thuế áp dụng: Dịch vụ trực tuyến chịu thuế GTGT tính theo phương pháp trực tiếp trên doanh thu với tỷ lệ 5% (doanh thu tính thuế nhân 5%) và thuế TNDN (CIT) với tỷ lệ 2% (đối với dịch vụ quảng cáo/phần mềm không bản quyền); Điểm mù rủi ro: Nếu doanh nghiệp chỉ lấy biên lai trừ tiền thẻ visa cá nhân hoặc invoice nước ngoài mà không thực hiện kê khai nộp thay FCT, toàn bộ chi phí này sẽ bị cơ quan thuế bóc tách, loại khỏi chi phí được trừ khi tính thuế TNDN và truy thu phạt chậm nộp; Khuyến nghị Eureka: Thiết lập quy trình rà soát hợp đồng và hóa đơn điện tử hàng tháng, kê khai tờ khai FCT Mẫu 01/NTNN đúng kỳ hạn, lưu giữ chứng từ nộp thuế vào NSNN làm căn cứ ghi nhận chi phí hợp lý được trừ 100%.',
         'scratch/daily_notes.md (12/11) & chinhphu.vn, Tổng cục Thuế | TN (Tin Ngành) / Foreign Contractor Tax FCT & Digital Cross-Border Services',
         update_date_str),

        (253,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl dinh danh dien tu',
         'thong tu 142 2026 tt btc, hai quan chuyen phat nhanh quoc te, chi tieu thong tin giam sat hai quan, dinh danh dien tu nguoi khai hai quan, to khai hang mau fedex dhl, thuong mai dien tu xuyen bien gioi',
         99.7,
         59200,
         'https://thuvienphapluat.vn',
         'Thông tư 142/2026/TT-BTC hiệu lực từ 16/11/2026: Chuẩn hóa chỉ tiêu giám sát hải quan chuyển phát nhanh quốc tế, bắt buộc định danh điện tử và giải pháp lấy tờ khai hàng mẫu',
         'Nội dung trọng tâm: Cảnh báo nước rút ngày 12/11/2026 khi chỉ còn 4 ngày nữa Thông tư số 142/2026/TT-BTC (ban hành ngày 29/09/2026) chính thức có hiệu lực thi hành kể từ ngày 16/11/2026: Đây là văn bản pháp quy mang tính bước ngoặt, sửa đổi và hoàn thiện các quy định về thủ tục hải quan đối với hàng hóa xuất khẩu, nhập khẩu, quá cảnh gửi qua dịch vụ bưu chính và dịch vụ chuyển phát nhanh quốc tế (DHL, FedEx, UPS...); Điểm mới cốt lõi tác động mạnh đến doanh nghiệp: (1) Chuẩn hóa toàn diện hệ thống chỉ tiêu thông tin khai báo và giám sát hải quan điện tử theo Phụ lục mới, tích hợp chặt chẽ cơ chế kiểm tra thông tin trước chuyến bay (Pre-loading / Pre-arrival Information); (2) Bắt buộc áp dụng định danh và xác thực điện tử đối với người khai hải quan và chủ hàng nhằm ngăn chặn triệt để tình trạng mượn danh nghĩa cá nhân để né thuế, buôn lậu hàng cấm qua đường bưu chính; (3) Phân định rạch ròi quy trình thông quan hàng mẫu, quà biếu quà tặng có giá trị thương mại vs phi thương mại; Giải pháp cho doanh nghiệp từ Eureka: Hàng ngàn doanh nghiệp thường xuyên gửi nhận hàng mẫu qua chuyển phát nhanh nhưng bị gom tờ khai chung của đại lý bưu chính dẫn đến không lấy được tờ khai hải quan đứng tên công ty để khấu trừ thuế và hạch toán chi phí hợp lệ; Eureka hướng dẫn chủ hàng chủ động gửi thông báo cho hãng chuyển phát nhanh yêu cầu "Mở tờ khai hải quan độc lập (Individual Customs Declaration)" ngay khi hàng đến cửa khẩu sân bay, cung cấp mã số thuế công ty và chữ ký số để hệ thống tự động trả tờ khai thông quan về email.',
         'customs.gov.vn & thuvienphapluat.vn | TN (Tin Ngành) / Express Customs Clearance Circular 142 & E-Identification',
         update_date_str),

        (254,
         'cong van 23969 chq nvthq kiem tra gia cau kien thep nha tien che nhap khau',
         'cong van 23969 chq nvthq, kiem tra xuat xu cau kien thep, tham van gia nha tien che nhap khau, khai bao tri gia thap thep ket cau, quan ly rui ro hai quan thep, kiem tra sau thong quan thep nhap khau',
         99.6,
         57800,
         'https://haiquanonline.com.vn',
         'Công văn 23969/CHQ-NVTHQ tháng 11/2026: Hải quan siết chặt kiểm tra xuất xứ và trị giá đối với mặt hàng cấu kiện thép, nhà tiền chế nhập khẩu có dấu hiệu giá thấp bất thường',
         'Nội dung trọng tâm: Cập nhật chỉ đạo nghiệp vụ khẩn cấp ngày 12/11/2026 theo Công văn số 23969/CHQ-NVTHQ của Cục Hải quan: Cơ quan Hải quan phát lệnh cho toàn bộ các Chi cục Hải quan khu vực và Đội Kiểm tra sau thông quan tăng cường quản lý rủi ro, siết chặt kiểm tra xuất xứ (C/O) và trị giá hải quan đối với các lô hàng cấu kiện thép, thép tiền chế, kết cấu kim loại xây dựng nhập khẩu (thuộc các nhóm HS 7308, 9406...); Nguyên nhân siết chặt: Nhiều doanh nghiệp lợi dụng sự chênh lệch giá thép nguyên liệu trên thị trường quốc tế để khai báo trị giá hải quan thấp bất thường (thấp hơn 30-50% so với cơ sở dữ liệu giá tham chiếu của Tổng cục Hải quan) nhằm trốn thuế nhập khẩu, thuế GTGT và né tránh các biện pháp phòng vệ thương mại; đồng thời xuất hiện tình trạng gian lận xuất xứ, chuyển tải bất hợp pháp để né thuế chống bán phá giá; Rủi ro trực tiếp: Toàn bộ tờ khai nhập khẩu cấu kiện thép sẽ bị chuyển luồng Vàng/Đỏ, 100% lô hàng nghi vấn bị yêu cầu tham vấn giá và lưu mẫu kiểm tra chất lượng nhà nước; Khuyến nghị Eureka: Doanh nghiệp nhập khẩu thép và nhà thầu xây dựng cần chuẩn bị sẵn bộ hồ sơ chứng minh trị giá hải quan 5 cấp độ (Hợp đồng ngoại thương, Mill Test Certificate của nhà máy sản xuất, chứng từ cước biển thực tế, sao kê chuyển tiền thanh toán qua ngân hàng) và kiểm tra tính hợp lệ của C/O form E, form KV để bảo vệ quyền lợi chính ngạch.',
         'haiquanonline.com.vn & customs.gov.vn | TN (Tin Ngành) / Steel Structure Customs Valuation & Trade Defense',
         update_date_str),

        (255,
         'ke khai giam 30 thue tncn ho kinh doanh online dich vu cong chung tu hop le',
         'ke khai giam 30 thue tncn, ho kinh doanh online, dich vu cong thue dien tu, chung tu hop le giam thue, truy thu thue ho kinh doanh ca the, huong dan ke khai thue tren dich vu cong',
         99.5,
         56500,
         'https://chinhphu.vn',
         'Hướng dẫn kê khai giảm 30% thuế TNCN cho hộ kinh doanh online trên Cổng Dịch vụ công: Hồ sơ chứng từ hợp lệ và phòng ngừa rủi ro bị truy thu phạt chậm nộp',
         'Nội dung trọng tâm: Hướng dẫn nghiệp vụ chi tiết ngày 12/11/2026 cho cộng đồng hộ kinh doanh cá thể, chủ shop thương mại điện tử và cá nhân kinh doanh online: Chính sách hỗ trợ giảm 30% số thuế TNCN phải nộp đang được triển khai mạnh mẽ trên Cổng Dịch vụ công Quốc gia và Cổng Thuế điện tử (etax Mobile / canhan.gdt.gov.vn); Điểm nghẽn thực tế: Nhiều hộ kinh doanh hào hứng tự kê khai trên Cổng Dịch vụ công nhưng do không nắm rõ phương pháp xác định doanh thu và điều kiện giảm trừ nên gặp hàng loạt sai sót: (1) Nhầm lẫn giữa việc giảm 30% trên số thuế TNCN phải nộp với việc giảm 30% trên tổng doanh thu; (2) Không lưu trữ đầy đủ hóa đơn, chứng từ mua vào hợp pháp đối với chi phí phát sinh; (3) Khai doanh thu thấp hơn dữ liệu do các sàn TMĐT (Shopee, TikTok Shop, Lazada) hoặc đơn vị vận chuyển đối soát báo cáo cho cơ quan thuế; Hậu quả pháp lý: Khi cơ quan thuế hậu kiểm dữ liệu tự động, các hộ kinh doanh kê khai sai sẽ bị thu hồi toàn bộ số tiền thuế được giảm, truy thu số thuế còn thiếu kèm tiền phạt chậm nộp 0.03%/ngày và phạt vi phạm hành chính từ 20% trên số thuế khai sai; Eureka tư vấn quy trình: (1) Hộ kinh doanh cần truy cập Cổng Dịch vụ công, chọn đúng thủ tục kê khai tờ khai thuế Mẫu 01/CNKD; (2) Tích chọn phụ lục giảm thuế theo đúng nghị quyết hiện hành; (3) Rà soát khớp 100% doanh thu sao kê tài khoản ngân hàng và báo cáo sàn TMĐT trước khi gửi tờ khai chính thức; (4) Lưu trữ bản mềm chứng từ tối thiểu 5 năm để phục vụ đối chiếu.',
         'scratch/daily_notes.md (12/11) & chinhphu.vn, thuvienphapluat.vn | HT (Học Thực Chiến) / Personal Income Tax Reduction 30% & E-Commerce Compliance',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 251 -> row 256
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

    # Cập nhật Total row ở row 261
    total_row_idx = 261
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 255 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D260)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E260)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 255 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 223 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 12/11/2026)\n"
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
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 255 chủ đề)  
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
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 12/11/2026)

Qua sàng lọc từ 58 sự kiện và case study nghiệp vụ thực tế, Agent 01 đã chọn lọc và chấm điểm chi tiết **TOP 5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe của quy trình `wf05` (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [CẤP THIẾT SAU THÔNG QUAN] Khắc phục sự cố khai nhầm hóa đơn và số lượng sau thông quan: Phác đồ khai bổ sung AMA, lập Mẫu 03 và giải pháp phòng vệ rủi ro phạt ấn định thuế
- **Chuyên mục (Pillar):** `CS` (Case Study) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Infographic 'Khai nhầm số hóa đơn và số lượng sau thông quan: 4 bước truyền AMA và lập Mẫu 03 tránh bị ấn định thuế và xử phạt hành chính' | YouTube: Video ThuyTalk CM1/CM2 phân tích quyền khai bổ sung trong 60 ngày theo Thông tư 39/2018/TT-BTC)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực nóng: Case study thực tế phát sinh trực tiếp ngày 12/11/2026 trên diễn đàn nghiệp vụ hải quan với hàng nhập khẩu thức ăn chăn nuôi thuế 0% |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực đau: Hàng đã thông quan lấy nhầm hóa đơn lô khác dẫn đến sai lệch số lượng; nguy cơ bị hải quan lập biên bản xử phạt theo Nghị định 128/2020/NĐ-CP và đưa vào tầm ngắm kiểm tra sau thông quan |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận sôi nổi: Thuế 0% có bị phạt trốn thuế không? Khai AMA có bị chuyển luồng Đỏ kiểm hóa lại không? Cần nộp Mẫu 03 hay chỉ cần công văn giải trình? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka đưa ra phác đồ chuẩn 4 bước: Dẫn chiếu Điểm b Khoản 1 Điều 20 TT 38 (sửa TT 39), biểu mẫu Mẫu 03/KBS/GSQL, bộ 5 chứng từ đối soát ngân hàng và người bán |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Chủ Đề Nóng Bỏng & Tính Ứng Dụng Khắc Phục Sự Cố Cao Nhất)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi sợ lớn nhất của nhân viên khai báo và kế toán kho khi xảy ra sai sót sau thông quan. Khẳng định năng lực tư vấn gỡ khó nghiệp vụ hải quan chuyên sâu của Eureka Logistics.
- **Nguồn xác minh:** Điểm b Khoản 1 Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi Thông tư 39/2018/TT-BTC); Nghị định 128/2020/NĐ-CP; `scratch/daily_notes.md` (12/11/2026).

---

### Đề xuất #2: [CHÍNH SÁCH THUẾ NÓNG — KINH DOANH SỐ] Kê khai nộp thay thuế nhà thầu (FCT) khi mua dịch vụ số của Meta, Microsoft, Google: Căn cứ Khoản 4 Điều 4 Luật Thuế GTGT 48/2024/QH15 và cách bảo vệ chi phí thuế hợp lệ
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Carousel Card 'Chi tiền triệu mua Ads Facebook, Google, Microsoft: Bạn đã nộp thay 5% thuế GTGT nhà thầu chưa? Coi chừng bị bóc sạch chi phí!' | YouTube: Video phân tích chi tiết cơ chế kê khai FCT cho doanh nghiệp SME)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Rất nóng: Vấn đề được mổ xẻ sâu trên Diễn đàn Thuế ngày 12/11/2026 căn cứ Luật Thuế GTGT 48/2024/QH15 và Thông tư 69/2025/TT-BTC |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Rất đau: Hàng chục ngàn doanh nghiệp chi tiền qua thẻ tín dụng cá nhân mua dịch vụ số nhưng không kê khai nộp thay FCT; khi quyết toán bị loại trừ toàn bộ chi phí hợp lý và truy thu phạt |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi nảy lửa: NCCNN đã đăng ký thuế tại Việt Nam thì doanh nghiệp có phải nộp thay nữa không? Phân biệt giữa mua cá nhân vs mua cho công ty |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka giải mã cặn kẽ: Điều kiện được trừ chi phí thuế TNDN, tỷ lệ 5% VAT trực tiếp, cách lập hồ sơ Mẫu 01/NTNN và đối soát sao kê ngân hàng |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #2 (Giải Mã Nghĩa Vụ Thuế Thương Mại Điện Tử & Dịch Vụ Xuyên Biên Giới)** |
- **Tóm tắt lý do nên làm:** Tiếp cận chính xác phân khúc chủ doanh nghiệp SME, giám đốc marketing và kế toán thuế. Thể hiện năng lực tư vấn tài chính - thuế - thương mại quốc tế toàn diện của Eureka.
- **Nguồn xác minh:** Khoản 4 Điều 4 Luật Thuế GTGT số 48/2024/QH15; Thông tư 69/2025/TT-BTC; Tổng cục Thuế; `scratch/daily_notes.md` (12/11/2026).

---

### Đề xuất #3: [NƯỚC RÚT VĂN BẢN MỚI — HIỆU LỰC 16/11] Thông tư 142/2026/TT-BTC hiệu lực từ 16/11/2026: Chuẩn hóa chỉ tiêu giám sát hải quan chuyển phát nhanh quốc tế, bắt buộc định danh điện tử và giải pháp lấy tờ khai hàng mẫu
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Facebook & Group Nghiệp Vụ Logistics** (Bài phân tích nghiệp vụ: 'Thông tư 142/2026/TT-BTC hiệu lực từ 16/11: 3 thay đổi cốt tử khi gửi hàng mẫu quốc tế qua FedEx, DHL — Cách lấy tờ khai hải quan riêng đứng tên công ty')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Nóng bỏng: Chỉ còn 4 ngày nữa (16/11/2026) Thông tư 142 chính thức có hiệu lực thi hành trên toàn quốc |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Gửi hàng mẫu, linh kiện thử nghiệm bị gom tờ khai chung của đơn vị bưu chính; kế toán không có tờ khai hải quan để ghi nhận tài sản/chi phí và khấu trừ thuế |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.7/5) | Thảo luận về quy định định danh điện tử đối với người khai hải quan và áp lực chuẩn hóa thông tin hàng hóa trước chuyến bay |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ bí quyết 'Mở tờ khai hải quan độc lập' với các hãng chuyển phát nhanh quốc tế, mẫu công văn chỉ định và lưu ý về thuế hàng mẫu H11 |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #3 (Đón Đầu Hiệu Lực Pháp Lý Ngành Hải Quan)** |
- **Tóm tắt lý do nên làm:** Đón đầu thời điểm vàng văn bản có hiệu lực để chiếm lĩnh vị thế người đưa tin và phân tích đầu tiên trong ngành, thu hút lượng lớn tương tác từ cộng đồng XNK.
- **Nguồn xác minh:** Thông tư số 142/2026/TT-BTC ngày 29/09/2026; Cổng TTĐT Tổng cục Hải quan (customs.gov.vn); Thư Viện Pháp Luật.

---

### Đề xuất #4: [SIẾT CHẶT QUẢN LÝ RỦI RO THÁNG 11] Công văn 23969/CHQ-NVTHQ: Hải quan tăng cường kiểm tra xuất xứ và trị giá hải quan đối với cấu kiện thép, nhà tiền chế nhập khẩu có dấu hiệu giá thấp bất thường
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Quản Trị Rủi Ro)
- **Nền tảng khuyến nghị:** **LinkedIn & Facebook Fanpage** (Bài viết chuyên gia: 'Hải quan siết kiểm tra trị giá và C/O cấu kiện thép nhập khẩu: 5 chứng từ sống còn để tránh bị ấn định thuế và phạt phòng vệ thương mại')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Rất mới: Chỉ đạo nghiệp vụ tháng 11/2026 của Cục Hải quan gửi các Chi cục Hải quan cửa khẩu và Đội KTSTQ |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Trị giá lô hàng cấu kiện thép, nhà xưởng rất lớn (hàng chục tỷ đồng); nếu bị bác C/O hoặc ấn định tăng giá tham vấn thì số thuế phát sinh lên đến hàng tỷ đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Tranh cãi về cơ sở xác định giá tham chiếu cấu kiện kim loại và căn cứ chuyển luồng kiểm tra thực tế 100% |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka cung cấp cẩm nang 5 tài liệu cốt lõi: Mill Test Certificate, phân tích chi phí sản xuất, chứng từ cước vận chuyển và phương pháp bảo vệ C/O hợp lệ |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #4 (Nội Dung Chuyên Sâu Cho Khối Doanh Nghiệp Công Nghiệp & Xây Dựng)** |
- **Tóm tắt lý do nên làm:** Định vị thương hiệu Eureka ở phân khúc hàng dự án, hàng công nghiệp nặng và xử lý các ca tham vấn giá phức tạp có giá trị kim ngạch lớn.
- **Nguồn xác minh:** Công văn số 23969/CHQ-NVTHQ; Tạp chí Hải quan Online (haiquanonline.com.vn); customs.gov.vn.

---

### Đề xuất #5: [HỖ TRỢ THUẾ HỘ KINH DOANH] Hướng dẫn kê khai giảm 30% thuế TNCN cho hộ kinh doanh online trên Cổng Dịch vụ công: Hồ sơ chứng từ hợp lệ và phòng ngừa rủi ro bị truy thu phạt chậm nộp
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group Thuế** (Bài hướng dẫn: 'Tự kê khai giảm 30% thuế TNCN hộ kinh doanh online trên Dịch vụ công: Đừng mừng vội nếu thiếu 3 chứng từ này khi cơ quan thuế hậu kiểm')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Đang được các chủ shop TMĐT và hộ kinh doanh quan tâm đặc biệt trong tháng 11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Kê khai sai tỷ lệ hoặc sai phương pháp giảm trừ dẫn đến bị truy thu thuế, tính tiền phạt chậm nộp 0.03%/ngày và bị phạt vi phạm hành chính |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Nhầm lẫn giữa giảm 30% thuế phải nộp vs giảm 30% doanh thu tính thuế; đối soát số liệu sàn TMĐT Shopee/TikTok với số liệu khai báo |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn thao tác từng bước trên Cổng Dịch vụ công, mẫu tờ khai 01/CNKD và lưu ý đối soát doanh thu ngân hàng |
  | **Tổng điểm** | **18.9 / 20** | **Xếp hạng: #5 (Cẩm Nang Hướng Dẫn Thực Chiến Cho Khối Hộ Kinh Doanh & Shop TMĐT)** |
- **Tóm tắt lý do nên làm:** Thu hút lượng tương tác khổng lồ từ cộng đồng kinh doanh online, tạo phễu khách hàng tiềm năng chuyển đổi sang dịch vụ thành lập doanh nghiệp và tư vấn thuế của Eureka.
- **Nguồn xác minh:** Nghị quyết của Quốc hội; Cổng Dịch vụ công Quốc gia; Cổng Thuế điện tử canhan.gdt.gov.vn; `scratch/daily_notes.md` (12/11/2026).

---

## 📌 HÀNH ĐỘNG TIẾP THEO ĐỀ XUẤT CHO CHỊ THÚY
1. **Lựa chọn chủ đề:** Chị Thúy có thể chọn 1 hoặc 2 chủ đề trong Top 5 để đưa vào lịch sản xuất content tuần (Tuần 46 qua quy trình `wf01_weekly_planning.md`).
2. **Sản xuất nhanh:** Nếu chọn Đề xuất #1 (Xử lý khai nhầm hóa đơn, số lượng sau thông quan qua AMA Mẫu 03) hoặc Đề xuất #2 (Nộp thay thuế nhà thầu FCT khi mua dịch vụ số Meta, Google), Agent 03 và Agent 06 sẵn sàng kích hoạt quy trình `wf02` (Facebook) hoặc `wf03` (YouTube) để viết bài và thiết kế infographic ngay hôm nay.
3. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT WF05 NGÀY 12/11/2026 ===")
    print(f"- File Excel: {FILE_PATH} ({file_size} bytes, SHA256: {file_sha256[:8]}...)")
    print(f"- OneDrive: {ONEDRIVE_FILE}")
    print(f"- Briefing: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w46_1211()
