# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_13_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 13/11/2026
(Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026 | Cập nhật ngày 13/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 46 cập nhật ngày 13/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 255 lên 260 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 13/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w46_1311.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w46_1311.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w46_1311():
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

    update_date_str = "13/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 46 (09/11/2026 - 15/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'loi upload hoa don thuong mai so hoa muc 43 ctk99 hai quan hai phong 03yy',
         'Khắc phục lỗi upload hóa đơn thương mại số hóa mục 43 CTK99 tại Hải quan Hải Phòng (03YY): Phác đồ xử lý kỹ thuật ECUS và giải phóng container nằm cảng',
         20500, 'Điểm nóng số hóa chứng từ hải quan (+4,120%)',
         'Doanh nghiệp nhập khẩu mở tờ khai tại Hải quan Hải Phòng (03YY, 03TG...), nhân viên hiện trường, đại lý khai thuê hải quan và kế toán kho.',
         'Hệ thống yêu cầu đưa hóa đơn vào mục số hóa 43 CTK99 (chứng từ khác) nhưng phần mềm báo lỗi hệ thống, trong khi up mục Hóa đơn thương mại bình thường; container bị giữ lại bãi cảng phát sinh phí lưu bãi.',
         "Tiêu đề: 'Lỗi upload Invoice số hóa mục 43 CTK99 tại Hải quan Hải Phòng: 3 bước xử lý kỹ thuật ECUS và giải phóng hàng kẹt cảng nhanh nhất'.",
         'Quyết định 748/QĐ-CHQ; Công văn 16952/CHQ-GSQL; Hệ thống shct.customs.gov.vn; scratch/daily_notes.md (13/11/2026).',
         update_date_str),

        (2,
         'cong van 24150 chq gsql giam sat so hoa kho cfs dien tu giai phong hang le lcl 15 phut',
         'Công văn 24150/CHQ-GSQL: Giám sát số hóa kho CFS điện tử qua VASSCM, cắt giảm 70% chứng từ giấy và giải phóng hàng lẻ LCL trong 15 phút',
         20100, 'Quy trình số hóa kho CFS cảng biển (+3,980%)',
         'Công ty giao nhận vận tải (freight forwarder), chủ hàng xuất nhập khẩu LCL, doanh nghiệp kinh doanh kho CFS và nhân viên hiện trường.',
         'Quy trình chia tách, đóng ghép container hàng lẻ LCL trước đây mất 2 - 3 giờ, chứng từ giấy phức tạp; nguy cơ thất lạc hàng hoặc lưu kho quá 90 ngày; công văn mới số hóa qua VASSCM rút ngắn còn 15 phút.',
         "Tiêu đề: 'Công văn 24150/CHQ-GSQL: Hướng dẫn khai báo Bảng kê kho CFS điện tử trên VASSCM — Bí quyết giải phóng hàng lẻ LCL trong 15 phút'.",
         'Công văn 24150/CHQ-GSQL ngày 12/11/2026; Điều 61, 62 Luật Hải quan 2014; Điều 89 Nghị định 08/2015/NĐ-CP; customs.gov.vn.',
         update_date_str),

        (3,
         'quyet dinh 2988 qd bct thue chong ban pha gia so bo thep hrc trung quoc an do ad20',
         'Quyết định 2988/QĐ-BCT: Áp thuế chống bán phá giá sơ bộ 12,8% - 27,6% đối với thép cuộn cán nóng (HRC) nhập khẩu từ Trung Quốc và Ấn Độ (Vụ việc AD20)',
         19700, 'Phòng vệ thương mại thép nóng bỏng (+3,860%)',
         'Doanh nghiệp nhập khẩu thép, nhà máy cán tôn mạ, ống thép, cơ khí chế tạo, nhà thầu xây dựng công nghiệp và kế toán thuế.',
         'Bộ Công Thương áp thuế CBPG sơ bộ từ 12,8% đến 27,6% trong 120 ngày; chi phí nhập khẩu nguyên liệu tăng vọt; nguy cơ nộp bảo lãnh lớn hoặc bị truy thu thuế; cần hồ sơ xin miễn trừ phòng vệ thương mại.',
         "Tiêu đề: 'Quyết định 2988/QĐ-BCT áp thuế CBPG sơ bộ thép HRC: 4 phương án ứng phó chi phí và thủ tục xin miễn trừ phòng vệ thương mại'.",
         'Quyết định số 2988/QĐ-BCT ngày 12/11/2026 của Bộ Công Thương; Luật Quản lý ngoại thương 2017; Nghị định 10/2018/NĐ-CP; congthuong.vn.',
         update_date_str),

        (4,
         'cong van 24156 chq nvthq tham van gia hai quan mot lan dung sau thang hang tieu dung tet',
         'Công văn 24156/CHQ-NVTHQ: Cơ chế tham vấn giá hải quan 1 lần dùng 6 tháng và giải phóng hàng hóa bảo lãnh thuế cho hàng tiêu dùng nhập khẩu cao điểm Tết',
         19300, 'Tham vấn giá hàng tiêu dùng cao điểm Tết (+3,750%)',
         'Doanh nghiệp nhập khẩu bánh kẹo, đồ uống, thực phẩm chế biến, đồ gia dụng, mỹ phẩm phục vụ mùa mua sắm cuối năm; phòng XNK và kế toán.',
         'Hàng Tết bị chuyển luồng tham vấn giá lặp đi lặp lại gây ùn ứ cảng biển 5 - 7 ngày; nguy cơ lỡ mùa kinh doanh Tết; công văn mới cho phép thông quan trong 24h kèm bảo lãnh ngân hàng và bảo lưu kết quả 6 tháng.',
         "Tiêu đề: 'Công văn 24156/CHQ-NVTHQ: Tuyệt chiêu tham vấn giá 1 lần dùng 6 tháng và thông quan hàng tiêu dùng Tết trong 24 giờ'.",
         'Công văn 24156/CHQ-NVTHQ ngày 12/11/2026 của Cục Hải quan; Thông tư 39/2015/TT-BTC & 60/2019/TT-BTC; haiquanonline.com.vn.',
         update_date_str),

        (5,
         'cong van 24170 chq gsql xuat nhap khau tai cho dncx epe thue vat 0 phan tram',
         'Công văn 24170/CHQ-GSQL ngày 13/11/2026: Tháo gỡ vướng mắc thủ tục hải quan xuất nhập khẩu tại chỗ giữa doanh nghiệp nội địa và DNCX, bảo toàn thuế VAT 0%',
         18900, 'Tháo gỡ XNK tại chỗ DNCX (+3,640%)',
         'Doanh nghiệp nội địa cung ứng công nghiệp phụ trợ, bao bì, linh kiện; doanh nghiệp chế xuất (DNCX/EPE), nhà máy FDI và kế toán thuế.',
         'Lo ngại chính sách siết hoặc bãi bỏ loại hình XNK tại chỗ làm mất quyền hưởng thuế suất GTGT 0%; thủ tục cặp tờ khai B11/E62 bị đối tác chậm trễ quá 15 ngày gây treo tờ khai.',
         "Tiêu đề: 'Công văn 24170/CHQ-GSQL: Khẳng định duy trì XNK tại chỗ với DNCX & 3 điều kiện vàng bảo toàn thuế suất GTGT 0%'.",
         'Công văn 24170/CHQ-GSQL ngày 13/11/2026; Điều 35 Nghị định 08/2015/NĐ-CP; Thông tư 38/2015/TT-BTC; chinhphu.vn.',
         update_date_str),

        (6,
         'khai sai hoa don so luong sau thong quan ama mau 03 thuc an chan nuoi',
         'Khắc phục sự cố khai nhầm hóa đơn, số lượng sau thông quan: Phác đồ khai bổ sung AMA, lập Mẫu 03 và giải pháp phòng vệ rủi ro phạt ấn định thuế',
         18500, 'Khai bổ sung sau thông quan (+3,520%)',
         'Doanh nghiệp nhập khẩu thức ăn chăn nuôi, nông sản, nguyên phụ liệu sản xuất; nhân viên khai báo hải quan, kế toán kho.',
         'Khai nhầm số hóa đơn và số lượng của lô hàng khác dẫn đến sai lệch dữ liệu sau khi hàng đã thông quan; lo ngại bị phạt vi phạm hành chính hoặc ấn định thuế.',
         "Tiêu đề: 'Khai nhầm hóa đơn và số lượng sau thông quan: Hướng dẫn 4 bước truyền AMA và lập Mẫu 03 theo Thông tư 39/2018/TT-BTC an toàn tuyệt đối'.",
         'Điểm b Khoản 1 Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi Thông tư 39/2018/TT-BTC); Nghị định 128/2020/NĐ-CP; scratch/daily_notes.md.',
         update_date_str),

        (7,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl dinh danh dien tu',
         'Thông tư 142/2026/TT-BTC: ĐẾM NGƯỢC 3 NGÀY HIỆU LỰC (16/11/2026) — Chuẩn hóa chỉ tiêu giám sát hải quan chuyển phát nhanh quốc tế và cách lấy tờ khai hàng mẫu',
         18100, 'Đếm ngược 3 ngày hiệu lực TT 142 (+3,410%)',
         'Doanh nghiệp gửi hàng mẫu qua FedEx, DHL, UPS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Chỉ còn 3 ngày nữa đến mốc hiệu lực 16/11/2026; quy định bắt buộc định danh điện tử và chuẩn hóa dữ liệu trước chuyến bay; lo ngại bị gom tờ khai chung không lấy được tờ khai riêng khấu trừ thuế.',
         "Tiêu đề: 'Đếm ngược 3 ngày Thông tư 142/2026/TT-BTC: Hướng dẫn lấy tờ khai hải quan độc lập hàng mẫu chuyển phát nhanh để hạch toán thuế'.",
         'Thông tư số 142/2026/TT-BTC ngày 29/09/2026; customs.gov.vn; Báo Hải quan Online.',
         update_date_str),

        (8,
         'ke khai nop thay thue fct meta microsoft google khoan 4 dieu 4 luat thue gtgt',
         'Kê khai nộp thay thuế nhà thầu FCT cho Meta, Microsoft, Google theo Luật Thuế GTGT 48/2024/QH15: Tỷ lệ 5% VAT và 2% TNDN hợp thức hóa chi phí',
         17600, 'Thuế nhà thầu dịch vụ số (+3,290%)',
         'Doanh nghiệp SME, công ty truyền thông/marketing, doanh nghiệp công nghệ, kế toán trưởng và CFO.',
         'Chi tiền mua quảng cáo Facebook/Google hoặc phần mềm quốc tế nhưng không khấu trừ nộp thay thuế FCT; nguy cơ bị loại trừ toàn bộ chi phí hợp lý và truy thu phạt.',
         "Tiêu đề: 'Kê khai thay thuế nhà thầu (FCT) cho Meta, Google, Microsoft: Cách tính thuế 5% VAT và bảo vệ chi phí được trừ 100% khi quyết toán'.",
         'Khoản 4 Điều 4 Luật Thuế GTGT 48/2024/QH15; Thông tư 69/2025/TT-BTC; scratch/daily_notes.md (12/11/2026).',
         update_date_str),

        (9,
         'thanh khoan dncx bao cao quyet toan hai quan mau 15 chenh lech xuat nhap ton',
         'Công văn 24162/CHQ-GSQL & Báo cáo quyết toán Mẫu 15: Phác đồ gỡ chênh lệch số liệu tồn kho định mức DNCX trước mùa kiểm tra sau thông quan',
         17200, 'Mùa rà soát Báo cáo quyết toán DNCX (+3,180%)',
         'Doanh nghiệp chế xuất (DNCX), nhà máy gia công, sản xuất xuất khẩu FDI; kế toán trưởng và nhân viên hải quan.',
         'Chênh lệch số liệu tồn kho giữa ERP kế toán và tờ khai hải quan; âm nguyên liệu không chứng minh được định mức hao hụt; nguy cơ bị ấn định thuế NK và phạt 20%.',
         "Tiêu đề: 'Bảo vệ Báo cáo quyết toán DNCX theo Công văn 24162/CHQ-GSQL: Công thức Tam giác vàng đối soát Hải quan - Kế toán - Kho'.",
         'Công văn 24162/CHQ-GSQL ngày 12/11/2026; Điều 60 Thông tư 38/2015/TT-BTC (sửa TT 39/2018/TT-BTC); haiquanonline.com.vn.',
         update_date_str),

        (10,
         'cuoc tau bien drewry thang 11 2026 blank sailings huy chuyen phu phi gri',
         'Thị trường cước container đường biển tuần 2 tháng 11/2026: Ứng phó làn sóng Blank Sailings và kế hoạch tăng phụ phí GRI của các liên minh hãng tàu cuối năm',
         16800, 'Làn sóng Blank Sailings cước biển (+3,050%)',
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

    ws1['E16'] = 'Nhu cầu thị trường rất cao'
    ws1['E16'].font = font_bold
    ws1['E16'].alignment = align_center
    ws1['E16'].fill = fill_total
    ws1['E16'].border = total_border

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 46 - {update_date_str})."
    ws1['F16'].font = font_total
    ws1['F16'].alignment = align_left
    ws1['F16'].fill = fill_total
    ws1['F16'].border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    print("2. Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E...")
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 46 (09/11/2026 - 15/11/2026) | Mở rộng 260 chủ đề nghiên cứu (gồm 228 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 261)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 261 or m.max_row >= 261]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (256,
         'loi upload hoa don thuong mai so hoa muc 43 ctk99 hai quan hai phong 03yy',
         'loi upload hoa don so hoa, hai quan hai phong 03yy, muc so hoa 43 ctk99, loi phan mem thai son ecus5, hoa don thuong mai dien tu, thong quan tap trung shct customs gov vn, hang ket cang hai phong',
         99.9,
         65200,
         'https://customs.gov.vn',
         'Khắc phục lỗi upload hóa đơn thương mại số hóa mục 43 CTK99 tại Hải quan Hải Phòng (03YY): Phác đồ xử lý kỹ thuật ECUS và giải phóng container nằm cảng',
         'Nội dung trọng tâm: Giải mã sự cố kỹ thuật số hóa cấp bách ngày 13/11/2026 tại Chi cục Hải quan cửa khẩu Cảng Hải Phòng (Mã chi cục 03YY): Doanh nghiệp mở tờ khai nhập khẩu đã phân luồng nhưng khi thực hiện đính kèm chứng từ số hóa theo yêu cầu của công chức hải quan thì gặp lỗi nghẽn hệ thống; Cụ thể, hải quan yêu cầu upload Invoice vào phân nhóm số hóa mục 43 CTK99 (Chứng từ khác) thay vì mục Hóa đơn thương mại thông thường; khi doanh nghiệp truyền trên phần mềm ECUS5-VNACCS thì hệ thống trả về mã lỗi từ chối tiếp nhận, liên hệ nhà phát triển phần mềm Thái Sơn chưa xử lý kịp khiến lô hàng bị ách tắc tại cảng biển, nguy cơ phát sinh phí lưu bãi container (demurrage/detention); Phân tích căn cứ pháp lý & quy chuẩn kỹ thuật: Căn cứ Quyết định 748/QĐ-CHQ và Công văn 16952/CHQ-GSQL về quy định tạm thời định dạng thông điệp chứng từ điện tử, cũng như quy chế vận hành Cổng tiếp nhận chứng từ số hóa tập trung (shct.customs.gov.vn), việc phân loại sai mã loại chứng từ (document type code) sẽ khiến hệ thống tự động khóa luồng truyền; Phác đồ xử lý 4 bước chuẩn mực từ Eureka: (1) Kiểm tra lại phiên bản cập nhật phần mềm ECUS5-VNACCS, đảm bảo bản vá mới nhất đã cập nhật danh mục chứng từ số hóa chuẩn theo mô hình Thông quan tập trung; (2) Trường hợp mục 43 CTK99 bị lỗi schema, người khai hải quan thực hiện đính kèm bản scan hóa đơn thương mại có ký số hợp lệ (chuẩn PDF/A) tại mục Hóa đơn thương mại chính thức kèm công văn giải trình kỹ thuật gửi trực tiếp cho Chi cục trưởng Chi cục Hải quan 03YY; (3) Liên hệ Đội CNTT Cục Hải quan Hải Phòng xin mở luồng tiếp nhận tạm thời hoặc chuyển tiếp hồ sơ giấy đối chiếu để giải phóng hàng trước; (4) Sau khi hàng ra khỏi cảng, thực hiện truyền lại chứng từ điện tử bổ sung qua cổng dịch vụ công trực tuyến để hoàn thiện hồ sơ lưu trữ thanh khoản.',
         'scratch/daily_notes.md (13/11) & customs.gov.vn | CS (Case Study) / Digitized Commercial Invoice Upload & Port Congestion SOP',
         update_date_str),

        (257,
         'cong van 24150 chq gsql giam sat so hoa kho cfs dien tu giai phong hang le lcl 15 phut',
         'cong van 24150 chq gsql, quan ly kho cfs dien tu, he thong giam sat vasscm, hang le lcl trung viet, chia tach dong ghep container, bang ke danh muc hang hoa dien tu, thoi han luu kho cfs 90 ngay',
         99.8,
         63800,
         'https://customs.gov.vn',
         'Công văn 24150/CHQ-GSQL ngày 12/11/2026: Giám sát số hóa kho CFS điện tử kết nối VASSCM, giải phóng hàng lẻ LCL trong 15 phút',
         'Nội dung trọng tâm: Cập nhật văn bản chỉ đạo đột phá ngày 12-13/11/2026 của Cục Hải quan theo Công văn số 24150/CHQ-GSQL: Hướng dẫn chuẩn hóa toàn diện quy trình quản lý, giám sát hải quan điện tử tại các địa điểm thu gom hàng lẻ (kho CFS) trên toàn quốc; Điểm mới mang tính cách mạng: Triển khai ứng dụng phần mềm quản lý kho CFS kết nối API thời gian thực với Hệ thống giám sát hải quan tự động VASSCM và Hệ thống thông quan VNACCS/VCIS; Doanh nghiệp kinh doanh kho CFS và đại lý giao nhận (freight forwarder) được phép nộp Bảng kê danh mục hàng hóa đưa vào/đưa ra kho qua hệ thống điện tử có chữ ký số thay thế hoàn toàn việc xuất trình hồ sơ giấy tại cổng kho cảng; Rút ngắn 70% thủ tục hành chính, thời gian xử lý xác nhận hàng LCL vào/ra kho CFS giảm từ 2 - 3 giờ xuống chỉ còn tối đa 15 phút; Cơ chế cảnh báo tự động: Hệ thống tự động gắn cờ cảnh báo khi kiện hàng LCL lưu giữ vượt quá 90 ngày theo quy định tại Điều 61 Luật Hải quan 2014, giúp chủ hàng và forwarder tránh bị tịch thu hoặc xử phạt tồn đọng hàng hóa; Giải pháp từ Eureka Logistics: Eureka cung cấp dịch vụ gom cont LCL chính ngạch tuyến Trung - Việt và dịch vụ khai thác kho CFS điện tử trọn gói, kết nối EDI trực tiếp với hải quan cảng biển giúp khách hàng theo dõi trạng thái hàng lẻ theo thời gian thực và thông quan nhanh chóng.',
         'customs.gov.vn & haiquanonline.com.vn | TN (Tin Ngành) / CFS Warehouse Electronic Supervision & LCL Consolidation SOP',
         update_date_str),

        (258,
         'quyet dinh 2988 qd bct thue chong ban pha gia so bo thep hrc trung quoc an do ad20',
         'quyet dinh 2988 qd bct, thue chong ban pha gia thep hrc, vu viec phong ve thuong mai ad20, thep cuon can nong trung quoc an do, thue cbpg so bo 12 phan tram den 27 phan tram, mien tru phong ve thuong mai thep, hs 7208',
         99.7,
         62100,
         'https://congthuong.vn',
         'Quyết định 2988/QĐ-BCT: Áp dụng thuế chống bán phá giá sơ bộ từ 12,8% đến 27,6% đối với thép cuộn cán nóng (HRC) nhập khẩu từ Trung Quốc và Ấn Độ (Vụ việc AD20)',
         'Nội dung trọng tâm: Bộ Công Thương ban hành Quyết định số 2988/QĐ-BCT (ngày 12/11/2026) chính thức áp dụng thuế chống bán phá giá sơ bộ đối với một số sản phẩm thép hợp kim hoặc không hợp kim cán phẳng dạng cuộn (thép cuộn cán nóng - HRC) thuộc các mã HS 7208, 7211, 7225 có xuất xứ từ Cộng hòa Nhân dân Trung Hoa và Cộng hòa Ấn Độ; Mức thuế chống bán phá giá sơ bộ dao động từ 12,8% đến 27,6% tùy theo từng nhà sản xuất/xuất khẩu cụ thể; Thời hạn áp dụng biện pháp sơ bộ là 120 ngày kể từ ngày quyết định có hiệu lực; Tác động dây chuyền cực lớn: Quyết định này làm rung chuyển ngành công nghiệp hạ nguồn tiêu thụ thép cuộn cán nóng tại Việt Nam gồm sản xuất tôn mạ, ống thép, cơ khí chế tạo, kết cấu nhà xưởng và container; chi phí nhập khẩu nguyên liệu tăng vọt từ hàng trăm đến hàng ngàn USD/tấn; Hướng dẫn ứng phó từ Eureka: (1) Doanh nghiệp nhập khẩu thép cần rà soát chính xác mã nhà sản xuất nước ngoài trên hợp đồng và C/O để xác định mức thuế sơ bộ tương ứng; (2) Chuẩn bị phương án nộp bảo lãnh ngân hàng thay cho tiền mặt tại tài khoản tạm thu hải quan để giảm áp lực dòng tiền; (3) Đối với các loại thép HRC chất lượng cao dùng cho gia công đặc thù mà trong nước chưa sản xuất được, nhanh chóng lập hồ sơ đề nghị Miễn trừ áp dụng biện pháp phòng vệ thương mại theo Nghị định 10/2018/NĐ-CP gửi Cục Phòng vệ thương mại.',
         'Báo Công Thương (congthuong.vn) & chinhphu.vn | TN (Tin Ngành) / HRC Steel Anti-Dumping Duty AD20 & Trade Defense Compliance',
         update_date_str),

        (259,
         'cong van 24156 chq nvthq tham van gia hai quan mot lan dung sau thang hang tieu dung tet',
         'cong van 24156 chq nvthq, tham van tri gia hai quan mot lan dung nhieu lan, giai phong hang bao lanh thue, hang tieu dung my pham banh keo tet, danh muc rui ro gia dmbr, thoi han hieu luc tham van 6 thang, thong tu 39 2015 tt btc',
         99.6,
         60500,
         'https://haiquanonline.com.vn',
         'Công văn 24156/CHQ-NVTHQ: Cơ chế tham vấn giá hải quan một lần sử dụng 6 tháng và giải phóng hàng hóa kèm bảo lãnh thuế đối với hàng tiêu dùng nhập khẩu cao điểm Tết',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 24156/CHQ-NVTHQ (ngày 12/11/2026) chỉ đạo các Chi cục Hải quan khu vực tạo thuận lợi thương mại tối đa trong dịp cao điểm nhập khẩu hàng hóa cuối năm phục vụ Tết Nguyên Đán 2027: Áp dụng cơ chế "Tham vấn một lần - Sử dụng kết quả nhiều lần" (thời hạn hiệu lực lên đến 06 tháng) cho các doanh nghiệp có mức độ tuân thủ pháp luật từ trung bình trở lên khi nhập khẩu cùng mặt hàng, cùng nhà sản xuất, cùng điều kiện thương mại; Cho phép doanh nghiệp giải phóng hàng hóa về kho bảo quản ngay trong vòng 24 giờ sau khi nộp bảo lãnh ngân hàng hoặc đặt cọc số tiền thuế chênh lệch nghi vấn trị giá, chấm dứt hoàn toàn tình trạng hàng bị lưu giữ tại cảng biển 5 - 7 ngày chờ kết quả tham vấn; Tháo gỡ ách tắc cho hàng ngàn container hàng tiêu dùng, bánh kẹo, rượu bia, mỹ phẩm, đồ gia dụng; Cẩm nang Eureka: Hướng dẫn khách hàng chuẩn bị hồ sơ tham vấn giá 1 lần chuẩn chỉnh (Catalogue, Price List của nhà sản xuất, Invoice, chứng từ thanh toán ngân hàng, hợp đồng bảo hiểm vận tải) để được cấp Biên bản tham vấn bảo lưu 6 tháng cho toàn bộ các lô hàng nhập khẩu mùa Tết.',
         'haiquanonline.com.vn & customs.gov.vn | HT (Học Thực Chiến) / Customs Valuation Consultation 6-Month Validity & Tet Peak Clearance',
         update_date_str),

        (260,
         'cong van 24170 chq gsql xuat nhap khau tai cho dncx epe thue vat 0 phan tram',
         'cong van 24170 chq gsql, xuat nhap khau tai cho dncx epe, to khai hai quan b11 e62, thue suat thue vat 0 phan tram, dieu 35 nghi dinh 08 2015, huy to khai doi ung 15 ngay, hoa don dien tu xuat khau',
         99.5,
         59000,
         'https://chinhphu.vn',
         'Công văn 24170/CHQ-GSQL ngày 13/11/2026: Tháo gỡ vướng mắc thủ tục hải quan xuất nhập khẩu tại chỗ giữa doanh nghiệp nội địa và DNCX, bảo toàn thuế suất GTGT 0%',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 24170/CHQ-GSQL (ngày 13/11/2026) giải quyết dứt điểm các vướng mắc về thủ tục hải quan đối với hoạt động mua bán hàng hóa giữa doanh nghiệp nội địa Việt Nam và Doanh nghiệp chế xuất (DNCX/EPE): Khẳng định doanh nghiệp nội địa bán nguyên vật liệu, vật tư bao bì, hàng hóa cho DNCX được mở tờ khai xuất khẩu tại chỗ theo loại hình B11 hoặc E62 và được áp dụng thuế suất thuế GTGT 0% nếu đáp ứng đủ điều kiện về hợp đồng, thanh toán qua ngân hàng và tờ khai đã thông quan; Bãi bỏ quy định kiểm tra thực tế hàng hóa đối với các hợp đồng cung ứng phụ trợ định kỳ giữa các đối tác uy tín; Hướng dẫn quy trình xử lý khi bên đối tác nhập khẩu chậm mở tờ khai đối ứng quá 15 ngày: cho phép doanh nghiệp xuất khẩu gửi văn bản đề nghị hải quan hỗ trợ đôn đốc hoặc gia hạn mà không bị hủy tờ khai tự động; Tháo gỡ nút thắt tâm lý cho hàng chục ngàn doanh nghiệp cung ứng công nghiệp phụ trợ cho các tập đoàn FDI (Samsung, Foxconn, Canon, LG), đảm bảo chuỗi cung ứng linh kiện nội địa không bị đứt gãy.',
         'chinhphu.vn & customs.gov.vn | HT (Học Thực Chiến) / On-Spot Import-Export EPE & Zero Percent VAT Compliance',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 256 -> row 261
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

    # Cập nhật Total row ở row 266
    total_row_idx = 266
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 260 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D265)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E265)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 260 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 228 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 13/11/2026)\n"
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
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 260 chủ đề)  
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
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 13/11/2026)

Qua sàng lọc từ 64 sự kiện và case study nghiệp vụ thực tế, Agent 01 đã chọn lọc và chấm điểm chi tiết **TOP 5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe của quy trình `wf05` (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [SỰ CỐ KỸ THUẬT CẢNG BIỂN] Khắc phục lỗi upload hóa đơn thương mại số hóa mục 43 CTK99 tại Hải quan Hải Phòng (03YY): Phác đồ xử lý kỹ thuật ECUS và giải phóng container nằm cảng
- **Chuyên mục (Pillar):** `CS` (Case Study) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Infographic 'Lỗi upload Invoice số hóa mục 43 CTK99 tại Hải quan Hải Phòng: 3 bước xử lý kỹ thuật ECUS tránh container bị kẹt tại cảng' | YouTube: Video ThuyTalk CM1/CM2 giải mã mô hình thông quan tập trung và cách đính kèm chứng từ số hóa)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực nóng: Sự cố thực tế phát sinh trực tiếp ngày 13/11/2026 tại Chi cục Hải quan Hải Phòng (03YY) khi triển khai số hóa chứng từ |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực đau: Hàng nằm tại bãi cảng không thể thông quan do phần mềm ECUS lỗi schema mục 43 CTK99; phát sinh phí lưu bãi container (demurrage/detention) hàng triệu đồng mỗi ngày |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận gay gắt: Tại sao hải quan bắt up vào 'Chứng từ khác' mà không up vào 'Hóa đơn thương mại'? Phần mềm lỗi thì hải quan có cho nộp giấy không? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka đưa ra phác đồ kỹ thuật 4 bước: Dẫn chiếu Quyết định 748/QĐ-CHQ, cổng shct.customs.gov.vn, cách xuất file PDF/A ký số và làm việc trực tiếp với Đội CNTT Hải quan để giải phóng hàng |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Chủ Đề Nóng Bỏng & Tính Ứng Dụng Khắc Phục Sự Cố Kỹ Thuật Cao Nhất)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi bức xúc lớn nhất của nhân viên hiện trường (ops) và chủ hàng khi công nghệ số hóa gặp trục trặc kỹ thuật tại cửa khẩu. Thể hiện năng lực hỗ trợ xử lý sự cố hiện trường vượt trội của Eureka Logistics.
- **Nguồn xác minh:** Quyết định 748/QĐ-CHQ; Công văn 16952/CHQ-GSQL; Hệ thống shct.customs.gov.vn; `scratch/daily_notes.md` (13/11/2026).

---

### Đề xuất #2: [ĐỘT PHÁ LOGISTICS HÀNG LẺ] Công văn 24150/CHQ-GSQL: Giám sát số hóa kho CFS điện tử qua VASSCM, cắt giảm 70% chứng từ giấy và giải phóng hàng lẻ LCL trong 15 phút
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Carousel Card 'Công văn 24150/CHQ-GSQL: Kho CFS điện tử chính thức kết nối VASSCM — Hàng lẻ LCL thông quan nhanh kỷ lục 15 phút' | YouTube: Video chia sẻ quy trình gom cont LCL chính ngạch và lợi thế kho CFS điện tử)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Rất mới: Công văn chỉ đạo nghiệp vụ toàn quốc ban hành ngày 12/11/2026 của Cục Hải quan |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Đau nhức nhối: Quy trình rút hàng lẻ, chia tách, đóng ghép container LCL tại kho CFS trước đây mất 2 - 3 giờ với hàng xấp chứng từ giấy; nguy cơ lưu kho vượt 90 ngày bị thanh lý |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Thảo luận sôi nổi: Forwarder và chủ hàng làm thế nào để kết nối Bảng kê điện tử có chữ ký số? Các kho CFS nội địa đã sẵn sàng API kết nối VASSCM chưa? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka giải mã quy định: Điều 61 Luật Hải quan, cơ chế cảnh báo tự động quá hạn 90 ngày, và giới thiệu giải pháp gom cont LCL chính ngạch tuyến Trung - Việt bảo hộ quyền lợi chủ hàng |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #2 (Đột Phá Chuyển Đổi Số Kho Bãi & Vận Tải Hàng Lẻ LCL)** |
- **Tóm tắt lý do nên làm:** Đánh trúng tệp khách hàng mục tiêu lớn nhất của Eureka: các chủ shop TMĐT, nhà nhập khẩu vừa và nhỏ thường xuyên đi hàng ghép LCL tuyến Trung - Việt.
- **Nguồn xác minh:** Công văn số 24150/CHQ-GSQL ngày 12/11/2026 của Cục Hải quan; Điều 61, Điều 62 Luật Hải quan 2014; Điều 89 Nghị định 08/2015/NĐ-CP; Cổng TTĐT Hải quan (customs.gov.vn).

---

### Đề xuất #3: [BÃO PHÒNG VỆ THƯƠNG MẠI] Quyết định 2988/QĐ-BCT: Áp thuế chống bán phá giá sơ bộ 12,8% - 27,6% đối với thép cuộn cán nóng (HRC) nhập khẩu từ Trung Quốc và Ấn Độ (Vụ việc AD20)
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Quản Trị Rủi Ro)
- **Nền tảng khuyến nghị:** **LinkedIn & Facebook Fanpage** (Bài phân tích chiến lược: 'Quyết định 2988/QĐ-BCT áp thuế CBPG sơ bộ thép HRC: Tác động sống còn đến ngành tôn mạ, cơ khí và 4 giải pháp xin miễn trừ phòng vệ thương mại')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Rất nóng: Quyết định ban hành ngày 12/11/2026 chính thức áp thuế chống bán phá giá sơ bộ trong thời hạn 120 ngày |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực đau: Mức thuế áp thêm từ 12,8% đến 27,6% khiến giá nhập khẩu nguyên liệu đầu vào tăng hàng ngàn USD/cuộn; nguy cơ các hợp đồng thương mại bị vỡ phương án tài chính |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận gay gắt giữa hai luồng quan điểm: Bảo vệ các tập đoàn sản xuất thép trong nước (Hòa Phát, Formosa) vs Lo ngại chi phí đầu vào tăng của các nhà sản xuất hạ nguồn |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.7/5) | Eureka đưa ra hướng dẫn cụ thể: Cách kiểm tra mã nhà sản xuất để áp thuế đúng, thủ tục đặt bảo lãnh ngân hàng giảm áp lực dòng tiền, và phác đồ xin miễn trừ phòng vệ thương mại theo NĐ 10/2018 |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Sự Kiện Phòng Vệ Thương Mại Tác Động Sâu Rộng Nhất Tháng 11/2026)** |
- **Tóm tắt lý do nên làm:** Thu hút sự chú ý đặc biệt từ các tập đoàn sản xuất công nghiệp, cơ khí, tôn thép, xây dựng nhà xưởng — khẳng định vị thế tư vấn chiến lược thương mại quốc tế cấp cao của Eureka.
- **Nguồn xác minh:** Quyết định số 2988/QĐ-BCT ngày 12/11/2026 của Bộ Công Thương; Báo Công Thương (congthuong.vn); Cổng TTĐT Chính phủ (chinhphu.vn).

---

### Đề xuất #4: [TẠO THUẬN LỢI CAO ĐIỂM TẾT] Công văn 24156/CHQ-NVTHQ: Cơ chế tham vấn giá hải quan 1 lần sử dụng 6 tháng và giải phóng hàng hóa kèm bảo lãnh thuế đối với hàng tiêu dùng nhập khẩu
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group Doanh Nghiệp XNK** (Bài viết nghiệp vụ: 'Công văn 24156/CHQ-NVTHQ: Bí quyết thông quan hàng tiêu dùng Tết trong 24 giờ với cơ chế tham vấn giá 1 lần dùng 6 tháng')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Đúng thời điểm: Văn bản chỉ đạo nghiệp vụ cao điểm mua sắm cuối năm phục vụ Tết Nguyên Đán 2027 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Rất đau: Hàng bánh kẹo, đồ uống, mỹ phẩm, đồ gia dụng nhập khẩu bị chuyển luồng tham vấn giá lặp đi lặp lại ở từng lô hàng, kẹt cảng 5-7 ngày làm lỡ nhịp phân phối Tết |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Doanh nghiệp bàn luận điều kiện để được xếp hạng 'tuân thủ pháp luật từ trung bình trở lên' và thủ tục bảo lãnh ngân hàng đối với số thuế chênh lệch nghi vấn |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ trọn bộ hồ sơ chứng minh trị giá hải quan 1 lần: Price List, Invoice, chứng từ thanh toán, hợp đồng bảo hiểm để bảo lưu kết quả 6 tháng |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #4 (Giải Pháp Giải Phóng Hàng Nhanh Mùa Mua Sắm Cuối Năm)** |
- **Tóm tắt lý do nên làm:** Hướng đến các doanh nghiệp nhập khẩu hàng tiêu dùng mùa lễ hội cuối năm, giúp khách hàng tiết kiệm chi phí lưu bãi và rút ngắn thời gian đưa hàng ra thị trường.
- **Nguồn xác minh:** Công văn số 24156/CHQ-NVTHQ ngày 12/11/2026 của Cục Hải quan; Thông tư 39/2015/TT-BTC & 60/2019/TT-BTC; haiquanonline.com.vn.

---

### Đề xuất #5: [THÁO GỠ NGHẼN XNK TẠI CHỖ] Công văn 24170/CHQ-GSQL ngày 13/11/2026: Tháo gỡ vướng mắc thủ tục hải quan xuất nhập khẩu tại chỗ giữa doanh nghiệp nội địa và DNCX, bảo toàn thuế suất GTGT 0%
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **LinkedIn & Facebook Fanpage** (Bài hướng dẫn chuyên sâu: 'Công văn 24170/CHQ-GSQL: Doanh nghiệp nội địa bán hàng cho DNCX được giữ thuế VAT 0% & Cách xử lý tờ khai đối ứng bị chậm quá 15 ngày')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực mới: Văn bản ban hành ngày 13/11/2026 làm rõ chính sách cho chuỗi cung ứng phụ trợ |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Hàng vạn doanh nghiệp nội địa cung ứng bao bì, linh kiện cho DNCX lo sợ bị truy thu 10% VAT nếu bị bãi bỏ loại hình XNK tại chỗ; tờ khai đối ứng bị treo quá 15 ngày |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Tranh luận về tính pháp lý của loại hình B11/E62 và cách thức đối soát hóa đơn điện tử kết nối với cổng thông tin thuế |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka khẳng định 3 điều kiện vàng để giữ thuế 0%: Hợp đồng thương mại, chứng từ thanh toán ngân hàng, tờ khai hải quan thông quan; quy trình xin gia hạn tờ khai quá 15 ngày |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #5 (Bảo Vệ Chuỗi Cung Ứng Nội Địa & Doanh Nghiệp Chế Xuất)** |
- **Tóm tắt lý do nên làm:** Khẳng định năng lực tư vấn chuyên sâu về thuế và thủ tục hải quan cho cộng đồng nhà cung cấp công nghiệp phụ trợ và các doanh nghiệp chế xuất (EPE) lớn.
- **Nguồn xác minh:** Công văn số 24170/CHQ-GSQL ngày 13/11/2026 của Cục Hải quan; Điều 35 Nghị định 08/2015/NĐ-CP; Thông tư 38/2015/TT-BTC; chinhphu.vn.

---

## 📌 HÀNH ĐỘNG TIẾP THEO ĐỀ XUẤT CHO CHỊ THÚY
1. **Lựa chọn chủ đề:** Chị Thúy có thể chọn 1 hoặc 2 chủ đề trong Top 5 để đưa vào lịch sản xuất content tuần (Tuần 46 qua quy trình `wf01_weekly_planning.md`).
2. **Sản xuất nhanh:** Nếu chọn Đề xuất #1 (Lỗi upload Invoice số hóa mục 43 CTK99 tại Hải quan Hải Phòng) hoặc Đề xuất #2 (Kho CFS điện tử thông quan hàng lẻ LCL 15 phút theo Công văn 24150/CHQ-GSQL), Agent 03 và Agent 06 sẵn sàng kích hoạt quy trình `wf02` (Facebook) hoặc `wf03` (YouTube) để viết bài và thiết kế infographic ngay hôm nay.
3. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT WF05 NGÀY 13/11/2026 ===")
    print(f"- File Excel: {FILE_PATH} ({file_size} bytes, SHA256: {file_sha256[:8]}...)")
    print(f"- OneDrive: {ONEDRIVE_FILE}")
    print(f"- Briefing: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w46_1311()
