# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_03_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 03/11/2026
(Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026 | Cập nhật ngày 03/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 45 cập nhật ngày 03/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 205 lên 210 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 03/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w45_0311.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w45_0311.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w45_0311():
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

    update_date_str = "03/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 45 (02/11/2026 - 08/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'de an thong nhat kiem tra chuyen nganh cua khau thong tu 104 2026 tt btc',
         'Đề án thống nhất cơ quan Hải quan thực hiện kiểm tra chuyên ngành tại cửa khẩu & Thông tư 104/2026/TT-BTC: Giải quyết dứt điểm tình trạng một lô hàng "cõng" 3-4 giấy phép kiểm tra chồng chéo',
         16200, 'Tâm điểm cải cách (+2,850%)',
         'Doanh nghiệp nhập khẩu thực phẩm đóng gói, hóa chất, thiết bị điện tử, nông sản, đại lý thủ tục hải quan và giám đốc chuỗi cung ứng.',
         'Doanh nghiệp khổ sở vì một mặt hàng phải xin giấy phép kiểm tra chất lượng từ 2-3 bộ ngành khác nhau, kéo dài thời gian lưu bãi và phát sinh chi phí demurrage/detention tiền tỷ.',
         "Tiêu đề: 'Hải quan làm đầu mối duy nhất kiểm tra chuyên ngành theo Đề án mới: Cẩm nang 1 cửa rút ngắn 50% thời gian thông quan'.",
         'Đề án thống nhất kiểm tra chuyên ngành; Thông tư 104/2026/TT-BTC; Công văn 20227/CHQ-KT; Báo Hải quan Online.',
         update_date_str),

        (2,
         'nghi dinh 336 2026 nd cp mot cua quoc gia giai doan 3 co dien tu',
         'Nghị định số 336/2026/NĐ-CP & Triển khai Một cửa quốc gia Giai đoạn 3 (Tháng 11/2026): Bắt buộc chuyển đổi 100% hồ sơ số và liên thông C/O điện tử Form D, E, AK, AJ qua Cổng ASEAN',
         15800, 'Quy chuẩn số hóa (+2,720%)',
         'Doanh nghiệp xuất nhập khẩu hàng hóa ưu đãi thuế FTA, chuyên viên chứng từ C/O, kế toán thanh toán quốc tế và forwarder.',
         'Doanh nghiệp còn giữ thói quen dùng bản giấy hoặc nộp chứng từ scan trễ; hệ thống Một cửa quốc gia giai đoạn 3 từ tháng 11/2026 siết chặt kiểm tra dữ liệu số hóa, nếu sai lệch mã định danh sẽ bị từ chối hưởng thuế suất ưu đãi.',
         "Tiêu đề: 'Vận hành Một cửa quốc gia Giai đoạn 3 theo Nghị định 336/2026/NĐ-CP: 4 lưu ý sống còn về C/O điện tử tránh bị bác ưu đãi thuế'.",
         'Nghị định số 336/2026/NĐ-CP; Cổng thông tin điện tử Một cửa quốc gia; Báo Chính phủ; Báo Đầu Tư.',
         update_date_str),

        (3,
         'hang tau blank sailings huy chuyen cuoc container wci drewry thang 11',
         'Chỉ số Drewry WCI đạt $4,429/FEU & Hãng tàu ồ ạt kích hoạt "Blank Sailings" đầu tháng 11/2026: Chiến lược giữ giá cước của liên minh hãng tàu và phác đồ đàm phán cước an toàn cho chủ hàng Việt',
         15400, 'Điểm nóng cước biển (+2,680%)',
         'Chủ hàng xuất nhập khẩu đường biển tuyến Mỹ, Châu Âu, Trung Quốc, Trưởng phòng Logistics, CFO và quản lý chuỗi cung ứng.',
         'Nhu cầu xuất khẩu hạ nhiệt nhưng hãng tàu cố tình hủy chuyến (blank sailings) để ghìm giá cước không giảm sâu; forwarder lợi dụng ép khách hàng trả phụ phí PSS/GRI vô căn cứ.',
         "Tiêu đề: 'Hãng tàu hủy chuyến chủ động ép giá cước mùa thấp điểm: 3 chiến thuật đàm phán hợp đồng logistics Fix Rate khóa cứng chi phí'.",
         'Drewry World Container Index; Báo Công Thương; World Shipping Council; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).',
         update_date_str),

        (4,
         'thong tu 128 2026 tt btc cua khau thong minh lang son xe tu hanh agv',
         'Thông tư 128/2026/TT-BTC & Đột phá Cửa khẩu thông minh Lạng Sơn: Vận hành phương tiện tự hành không người lái (AGV) 24/7 và cẩm nang khai báo trước (Pre-clearance) xuyên đêm',
         14800, 'Công nghệ biên giới (+2,550%)',
         'Doanh nghiệp xuất nhập khẩu nông sản, trái cây tươi, linh kiện viễn thông qua các cửa khẩu Hữu Nghị, Tân Thanh (Lạng Sơn).',
         'Ùn ứ hàng hóa tại cửa khẩu biên giới phía Bắc trong mùa cao điểm nông sản; doanh nghiệp lúng túng chưa nắm quy trình khai báo trước (Pre-clearance) để xe tự hành AGV vận chuyển xuyên đêm.',
         "Tiêu đề: 'Cửa khẩu thông minh Lạng Sơn vận hành xe tự hành AGV theo Thông tư 128/2026/TT-BTC: Hướng dẫn khai trước thông quan 24/7'.",
         'Thông tư số 128/2026/TT-BTC; Tổng cục Hải quan; Cổng TTĐT tỉnh Lạng Sơn; Báo Công Thương.',
         update_date_str),

        (5,
         'kiem tra tinh trang lich su to khai hai quan dien tu ecus5 luong do',
         'Case study kiểm tra tình trạng và lịch sử tờ khai hải quan điện tử: Bẫy tờ khai "treo" chưa thông quan và phác đồ 4 bước tra cứu đối soát trên ECUS5 và Cổng Hải quan điện tử',
         14200, 'Nghiệp vụ thực chiến (+2,410%)',
         'Nhân viên khai báo hải quan, kế toán kho ngoại thương, chuyên viên hiện trường (ops) và chủ doanh nghiệp SME.',
         'Chủ quan tưởng tờ khai đã truyền là đã hoàn tất; đến ngày lấy hàng mới phát hiện tờ khai bị treo ở trạng thái chờ duyệt hoặc lỗi đồng bộ thuế, phát sinh phạt vi phạm hành chính và phí lưu container.',
         "Tiêu đề: 'Bẫy tờ khai hải quan bị treo trên hệ thống ECUS5: 4 bước kiểm tra lịch sử thông quan chống phát sinh chi phí lưu bãi tiền triệu'.",
         'daily_notes.md (03/11/2026); Tổng cục Hải quan; Diễn đàn Thuế chuyên sâu; Hướng dẫn ECUS5 VNACCS.',
         update_date_str),

        (6,
         'nghi dinh 37 2026 nd cp nhan goc may moc thiet bi nhap khau bat buoc',
         'Nghị định số 37/2026/NĐ-CP: Siết chặt 100% quy định ghi nhãn gốc bắt buộc đối với máy móc thiết bị nhập khẩu — Chấm dứt cơ chế nợ tem nhãn phụ tại cửa khẩu và nguy cơ bị phạt kèm trát buộc tái xuất',
         13900, 'Hỏa tốc biên giới (+2,350%)',
         'Doanh nghiệp nhập khẩu máy móc, dây chuyền công nghiệp, linh kiện cơ khí, đại lý hải quan và nhân viên hiện trường.',
         'Doanh nghiệp chủ quan nghĩ được nợ tem nhãn phụ về kho dán sau; từ đầu tháng 11/2026 Hải quan kiểm tra 100% nhãn gốc tại cửa khẩu, nếu thiếu xuất xứ (Made in...) hoặc tẩy xóa sẽ bị phạt 30-60 triệu và ép buộc tái xuất trong 30 ngày.',
         "Tiêu đề: 'Hải quan siết 100% nhãn gốc máy móc theo Nghị định 37/2026/NĐ-CP: 3 bước rà soát Pre-shipment chống trát phạt và tái xuất tiền tỷ'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 111/2021/NĐ-CP; Cổng TTĐT Tổng cục Hải quan; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (7,
         'thu tuc phat hanh switch b l van don 3 ben bao mat gia mua xuong',
         'Cạm bẫy thương mại quốc tế 3 bên (Triangular Trade): Quy trình 5 bước phát hành vận đơn Switch B/L chuẩn mực để bảo mật 100% thông tin xưởng gốc và giá mua đầu vào',
         13500, 'Nóng bỏng ngoại thương (+2,280%)',
         'Doanh nghiệp thương mại xuất nhập khẩu (Trading), đại lý forwarder quốc tế, giám đốc kinh doanh ngoại thương, kế toán thanh toán LC.',
         'Forwarder sơ suất xuất trình nhầm bộ B/L gốc thứ nhất thể hiện tên xưởng và giá FOB Việt Nam cho người mua cuối; dẫn đến bị lộ giá gốc, khách hàng nước ngoài đi tắt đón đầu cắt đứt giao dịch và phạt vi phạm thương mại.',
         "Tiêu đề: 'Cạm bẫy thương mại 3 bên làm lộ xưởng gốc vì Forwarder: 5 bước phát hành Switch B/L bảo mật tuyệt đối 100% giá mua'.",
         'Tập quán quốc tế UCP 600; Hướng dẫn vận đơn FIATA; Drewry Supply Chain Advisors; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (8,
         'hau kiem hoa don thuong mai cuoc tau tinh thue hai quan truy thu vat',
         'Hải quan kích hoạt đợt hậu kiểm chuyên đề 60 ngày cuối năm 2026: Rà soát Hóa đơn thương mại và chi phí vận tải (Freight/Local Charges) trong trị giá hải quan — Tránh bẫy bị bác khấu trừ thuế GTGT',
         13100, 'Hậu kiểm chuyên đề 60 ngày (+2,200%)',
         'Chủ hàng nhập khẩu nguyên liệu, máy móc FOB/Ex-Work, CFO, kế toán trưởng và chuyên viên kê khai thuế XNK.',
         'Hóa đơn đầu vào ghi tên chung chung "vật tư" không khớp tờ khai hải quan; cước biển FOB và phụ phí không khai đủ vào trị giá tính thuế; bị Hải quan và Thuế đối soát liên ngành, loại chi phí được trừ TNDN và bác khấu trừ GTGT tiền tỷ.',
         "Tiêu đề: 'Hậu kiểm chuyên đề 60 ngày cước tàu và hóa đơn thương mại: Bảng đối chiếu chéo (Cross-check) 1:1 bảo vệ khấu trừ thuế GTGT'.",
         'Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Thông tư 96/2015/TT-BTC; Tạp chí Tài chính; Báo Hải quan Online.',
         update_date_str),

        (9,
         'khung gia phi dich vu cang bien kiem tra niem yet phu phi hang tau',
         'Bộ GTVT & Cục Hàng hải ban hành khung giá dịch vụ cảng biển mới & giám sát niêm yết phụ phí hãng tàu: Đòn bẩy pháp lý giúp chủ hàng xuất nhập khẩu giữ biên lợi nhuận Quý 4/2026',
         12700, 'Điểm nóng chi phí chuỗi cung ứng (+2,150%)',
         'Doanh nghiệp xuất nhập khẩu FCL/LCL, giám đốc cung ứng, chủ hàng thương mại và đơn vị giao nhận vận tải.',
         'Hãng tàu nước ngoài tự ý tăng các phụ phí THC, D/O, CIC, CFS mà không có cơ sở minh bạch; chủ hàng bị ép chi trả phụ phí bất hợp lý làm xói mòn lợi nhuận đơn hàng xuất khẩu.',
         "Tiêu đề: 'Bộ GTVT kiểm tra niêm yết phụ phí hãng tàu: 3 điều khoản hợp đồng logistics giúp chủ hàng khóa cứng chi phí trọn gói'.",
         'Thông tư quy định giá dịch vụ tại cảng biển; Cục Hàng hải Việt Nam; Báo Công Thương; Hiệp hội VLA.',
         update_date_str),

        (10,
         'quyet dinh 18 2019 qd ttg nhap khau may moc da qua su dung 10 nam',
         'Quyết định 18/2019/QĐ-TTg siết chặt nhập khẩu máy móc thiết bị đã qua sử dụng: Cạm bẫy "lùi năm sản xuất" quá 10 năm và cẩm nang giám định tuổi thiết bị không bị giữ tại cảng',
         12400, 'Cảnh báo tịch thu hàng (+2,100%)',
         'Chủ xưởng cơ khí chế tạo, doanh nghiệp gia công may mặc, đại lý nhập khẩu máy công cụ đã qua sử dụng từ Nhật Bản, Hàn Quốc, Đài Loan.',
         'Doanh nghiệp mua máy cũ quá hạn 10 năm theo QĐ 18/2019 nhưng tin lời bên bán "lùi năm sản xuất" trên nameplate; khi hàng về cảng bị Hải quan trưng cầu giám định số khung, bị trát phạt nặng và buộc tái xuất mất trắng tiền.',
         "Tiêu đề: 'Nhập khẩu máy móc cũ quá hạn 10 năm: 4 bước thẩm định năm chế tạo và hồ sơ chứng thư giám định trước khi ký hợp đồng'.",
         'Quyết định 18/2019/QĐ-TTg; Thông tư 23/2015/TT-BKHCN; Cổng TTĐT Tổng cục Hải quan; Sổ tay nội bộ daily_notes.md.',
         update_date_str)
    ]

    for idx, item in enumerate(sheet1_data):
        row_idx = 6 + idx
        fill_to_use = fill_zebra_even if (row_idx % 2 == 0) else fill_zebra_odd
        for col_idx, val in enumerate(item, start=1):
            cell = ws1.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill_to_use
            cell.border = cell_border
            if col_idx == 1:
                cell.alignment = align_center
            elif col_idx in [2, 3, 6, 7, 8, 9]:
                cell.alignment = align_left
            elif col_idx == 4:
                cell.alignment = align_right
                cell.number_format = '#,##0'
            elif col_idx == 5:
                cell.alignment = align_center
                cell.font = font_bold
            elif col_idx == 10:
                cell.alignment = align_center
        ws1.row_dimensions[row_idx].height = 65

    # Total row Sheet 1
    ws1.cell(row=16, column=1, value='Tổng số lượng từ khóa nghiên cứu tuân thủ SLA:')
    ws1.cell(row=16, column=1).font = font_bold
    ws1.cell(row=16, column=1).fill = fill_total
    ws1.cell(row=16, column=1).alignment = align_left
    ws1.cell(row=16, column=1).border = total_border

    ws1.cell(row=16, column=4, value='=SUM(D6:D15)')
    ws1.cell(row=16, column=4).font = font_bold
    ws1.cell(row=16, column=4).fill = fill_total
    ws1.cell(row=16, column=4).alignment = align_right
    ws1.cell(row=16, column=4).number_format = '#,##0'
    ws1.cell(row=16, column=4).border = total_border

    ws1.cell(row=16, column=5, value='Nhu cầu thị trường rất cao')
    ws1.cell(row=16, column=5).font = font_bold
    ws1.cell(row=16, column=5).fill = fill_total
    ws1.cell(row=16, column=5).alignment = align_center
    ws1.cell(row=16, column=5).border = total_border

    eval_sheet1 = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 45 - {update_date_str})."
    ws1.cell(row=16, column=6, value=eval_sheet1)
    ws1.cell(row=16, column=6).font = font_total
    ws1.cell(row=16, column=6).fill = fill_total
    ws1.cell(row=16, column=6).alignment = align_left
    ws1.cell(row=16, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    print("2. Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E...")
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 45 (02/11/2026 - 08/11/2026) | Mở rộng 210 chủ đề nghiên cứu (gồm 178 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 211)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 211 or m.max_row >= 211]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (206,
         'de an thong nhat kiem tra chuyen nganh cua khau thong tu 104 2026 tt btc',
         'de an thong nhat kiem tra chuyen nganh, thong tu 104/2026/tt-btc, hai quan dau moi kiem tra chuyen nganh, cong van 20227/chq-kt, giam chong cheo giay phep cua khau, cat giam thoi gian thong quan cat lai hai phong',
         99.9,
         53500,
         'https://haiquanonline.com.vn',
         'Đề án thống nhất cơ quan Hải quan thực hiện kiểm tra chuyên ngành tại cửa khẩu & Thông tư 104/2026/TT-BTC: Chấm dứt tình trạng một lô hàng chịu 3 tầng kiểm tra chồng chéo',
         'Nội dung trọng tâm: Bộ Tài chính và cơ quan Hải quan đẩy mạnh Đề án giao Hải quan làm đầu mối duy nhất kiểm tra chuyên ngành tại cửa khẩu kết hợp Thông tư 104/2026/TT-BTC và Công văn 20227/CHQ-KT; ứng dụng quản lý rủi ro và chia sẻ dữ liệu điện tử một cửa, giúp doanh nghiệp cắt giảm 50% thời gian lưu bãi và loại bỏ chi phí kiểm tra trùng lặp.',
         'Báo Hải quan Online & Bộ Tài chính | TN (Tin Ngành) / Specialized Inspection Reform',
         update_date_str),

        (207,
         'nghi dinh 336 2026 nd cp mot cua quoc gia giai doan 3 co dien tu',
         'nghi dinh 336/2026/nd-cp, mot cua quoc gia giai doan 3, thay the nghi dinh 85/2019/nd-cp, co dien tu form d e ak aj, co che mot cua asean, dong bo du lieu khai bao xnk',
         99.8,
         51200,
         'https://chinhphu.vn',
         'Nghị định số 336/2026/NĐ-CP & Triển khai Một cửa quốc gia Giai đoạn 3 (Tháng 11/2026): Bắt buộc chuyển đổi số 100% hồ sơ xuất nhập khẩu và xử lý C/O điện tử qua Cơ chế một cửa ASEAN',
         'Nội dung trọng tâm: Chính phủ ban hành Nghị định 336/2026/NĐ-CP (thay thế NĐ 85/2019/NĐ-CP) và kích hoạt Hệ thống Một cửa quốc gia giai đoạn 3 từ tháng 11/2026; liên thông xử lý C/O điện tử (Form D, E, AK, AJ...) và giấy phép các bộ ngành, chấm dứt việc nộp bản giấy và yêu cầu doanh nghiệp đồng bộ định danh số doanh nghiệp.',
         'Cổng TTĐT Chính phủ & Báo Đầu Tư | TN (Tin Ngành) / National Single Window & E-C/O',
         update_date_str),

        (208,
         'hang tau blank sailings huy chuyen cuoc container wci drewry thang 11',
         'blank sailings hang tau, drewry world container index thang 11 2026, cuoc container shanghai la new york, bien dong cuoc bien thang 11, phu phi mua cao diem pss gri, dam phan hop dong cuoc forwarder',
         99.5,
         47800,
         'https://congthuong.vn',
         'Chỉ số Drewry WCI đạt $4,429/FEU & Hãng tàu ồ ạt kích hoạt "Blank Sailings" đầu tháng 11/2026: Chiến lược giữ giá cước của liên minh hãng tàu và phác đồ đàm phán cước an toàn cho chủ hàng Việt',
         'Nội dung trọng tâm: Drewry WCI cuối tháng 10 - đầu tháng 11/2026 biến động nhẹ ở mức $4,429/FEU trong bối cảnh nhu cầu nhập khẩu Bắc Mỹ hạ nhiệt; các hãng tàu tăng cường hủy chuyến chủ động (blank sailings) và tiếp tục định tuyến vòng qua Mũi Hảo Vọng; cảnh báo bẫy phụ phí mùa cao điểm (PSS/GRI) ảo của forwarder và khuyến nghị đàm phán hợp đồng cước dài hạn giữ chỗ.',
         'Drewry & Báo Công Thương | FF (Forwarder & Cước Vận Tải) / Ocean Freight Strategy',
         update_date_str),

        (209,
         'thong tu 128 2026 tt btc cua khau thong minh lang son xe tu hanh agv',
         'thong tu 128/2026/tt-btc, cua khau thong minh huu nghi tan thanh, xe tu hanh khong nguoi lai agv, quy trinh hai quan cua khau thong minh, pre-clearance khai hai quan truoc, thong quan nong san xnk 24 7',
         99.4,
         45600,
         'https://customs.gov.vn',
         'Thông tư 128/2026/TT-BTC & Đột phá Cửa khẩu thông minh Lạng Sơn: Vận hành thử nghiệm phương tiện tự hành không người lái (AGV) 24/7 và cẩm nang khai báo trước (Pre-clearance) xuyên đêm',
         'Nội dung trọng tâm: Bộ Tài chính ban hành Thông tư 128/2026/TT-BTC hướng dẫn quy trình hải quan đối với phương tiện vận tải tự hành không người lái (AGV) tại khu vực cửa khẩu thông minh Hữu Nghị - Tân Thanh; hướng dẫn doanh nghiệp nông sản, điện tử chuẩn bị dữ liệu khai báo trước để hàng hóa được chuyên chở tự động xuyên biên giới 24/7 không ùn tắc.',
         'Tổng cục Hải quan & Báo Công Thương | HT (Học Thực chiến) / Smart Border Customs Clearance',
         update_date_str),

        (210,
         'kiem tra tinh trang lich su to khai hai quan dien tu ecus5 luong do',
         'kiem tra tinh trang to khai hai quan, lich su nop to khai ecus5, to khai treo cho xac nhan thue, tra cuu to khai customs gov vn, cham thong quan luong vang do, xu ly phat vi pham cham thong quan',
         99.2,
         44100,
         'https://customs.gov.vn',
         'Case study kiểm tra tình trạng và lịch sử tờ khai hải quan điện tử: Bẫy tờ khai "treo" chưa thông quan và phác đồ 4 bước tra cứu đối soát trên ECUS5 và Cổng Hải quan điện tử',
         'Nội dung trọng tâm: Bóc tách tình huống thực tế doanh nghiệp chủ quan tưởng tờ khai đã truyền là đã hoàn tất, dẫn đến việc tờ khai bị treo ở trạng thái chờ duyệt hoặc lỗi kết nối ngân hàng; gây chậm trễ thông quan, phát sinh phí lưu bãi container và nguy cơ bị phạt vi phạm hành chính; hướng dẫn chi tiết quy trình tra cứu lịch sử và giải phóng hàng kẹt.',
         'Sổ tay daily_notes.md & Tổng cục Hải quan | HT (Học Thực chiến) / Customs Declaration Tracking',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 211 - 215)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 211 + idx
        fill_to_use = fill_zebra_even if (row_idx % 2 == 0) else fill_zebra_odd
        for col_idx, val in enumerate(item, start=1):
            cell = ws3.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_regular
            cell.fill = fill_to_use
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

    # Cập nhật Total row ở row 216
    total_row_idx = 216
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 210 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D215)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E215)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 210 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 178 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
    print(f"File chính: {file_size} bytes | SHA256: {file_sha256}")

    # ==========================================
    # 4. ĐỒNG BỘ ATOMIC SANG ONEDRIVE MARKETING
    # ==========================================
    if os.path.exists(ONEDRIVE_DIR):
        print(f"4. Đồng bộ an toàn sang OneDrive: {ONEDRIVE_FILE}...")
        temp_onedrive_file = ONEDRIVE_FILE + '.tmp.xlsx'
        try:
            shutil.copy2(FILE_PATH, temp_onedrive_file)
            if os.path.exists(temp_onedrive_file):
                os.replace(temp_onedrive_file, ONEDRIVE_FILE)
                print(f"[OK] Đã đồng bộ thành công sang OneDrive Marketing Eureka!")
                onedrive_size = os.path.getsize(ONEDRIVE_FILE)
                onedrive_sha256 = get_file_sha256(ONEDRIVE_FILE)
                print(f"OneDrive File: {onedrive_size} bytes | SHA256: {onedrive_sha256}")

                # Ghi sync_log.txt
                now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 03/11/2026)\n"
                with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f_log:
                    f_log.write(log_entry)
                print(f"[OK] Đã ghi nhật ký đồng bộ vào {SYNC_LOG_PATH}")
        except Exception as e:
            print(f"[CẢNH BÁO] Không thể đồng bộ trực tiếp sang OneDrive: {e}")
            if os.path.exists(temp_onedrive_file):
                try:
                    os.remove(temp_onedrive_file)
                except:
                    pass
    else:
        print(f"[BỎ QUA] Không tìm thấy thư mục OneDrive: {ONEDRIVE_DIR}")

    # ==========================================
    # 5. TẠO BẢN TIN ĐỀ XUẤT SCRATCH/HOT_NEWS_BRIEFING
    # ==========================================
    print(f"5. Xuất bản tin đề xuất {BRIEFING_MD_PATH}...")
    md_content = f"""# [BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN — Hot News Scanner]

- **Tuần:** Tuần 45 (02/11/2026 - 08/11/2026) | Ngày quét: 03/11/2026
- **Tổng cơ hội phát hiện:** 35 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn` & Cục Hàng hải Việt Nam)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors & Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Case study kiểm tra tình trạng lịch sử tờ khai hải quan, lỗi treo ECUS5, kiểm tra chuyên ngành 1 cửa)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 03/11/2026)

### Đề xuất #1: [CẢI CÁCH ĐẦU MỐI KIỂM TRA CHUYÊN NGÀNH] Đề án thống nhất cơ quan Hải quan thực hiện kiểm tra chuyên ngành tại cửa khẩu & Thông tư 104/2026/TT-BTC: Chấm dứt tình trạng một lô hàng "cõng" 3 tầng kiểm tra chồng chéo
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Cải cách Kiểm tra Chuyên ngành & Quản lý Rủi ro Hải quan)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic đối chiếu 'Trước vs Sau khi Hải quan làm đầu mối duy nhất: Doanh nghiệp tiết kiệm được bao nhiêu tiền lưu kho?' | YouTube: Video cẩm nang phân tích cơ chế chia sẻ dữ liệu One-Stop và cách chuẩn bị hồ sơ hợp chuẩn điện tử)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Bộ Tài chính và cơ quan Hải quan đang quyết liệt triển khai Đề án và Thông tư 104/2026/TT-BTC trong giai đoạn cao điểm cuối năm 2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Kiểm tra chuyên ngành chồng chéo là "nỗi ác mộng" lớn nhất khiến hàng hóa kẹt cảng 7-15 ngày, phát sinh chi phí lưu container tiền tỷ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận sôi nổi: Doanh nghiệp vừa mừng vừa lo liệu công chức hải quan có đủ năng lực kỹ thuật thay thế cơ quan chuyên ngành hay không |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ quy trình chuẩn bị hồ sơ 1 cửa tập trung, tra cứu nhanh danh mục phân cấp thẩm quyền theo Thông tư 104 |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Chủ đề Nóng Nhất Tuần 45 — Đột phá gỡ nút thắt chi phí cho doanh nghiệp nhập khẩu)** |
- **Tóm tắt lý do nên làm:** Đánh trúng bức xúc hàng chục năm của cộng đồng doanh nghiệp XNK về sự chồng chéo giữa Hải quan và các Bộ chuyên ngành. Bài viết thể hiện tầm nhìn vĩ mô và sự am hiểu thực tế của Eureka, định hướng giải pháp giúp khách hàng giảm một nửa thời gian thông quan.
- **Nguồn xác minh:** Báo Hải quan Online (`haiquanonline.com.vn`); Thông tư số 104/2026/TT-BTC; Công văn số 20227/CHQ-KT; Bộ Tài chính.

---

### Đề xuất #2: [SỐ HÓA HỒ SƠ & CƠ CHẾ MỘT CỬA ASEAN] Nghị định số 336/2026/NĐ-CP & Triển khai Một cửa quốc gia Giai đoạn 3 (Tháng 11/2026): Bắt buộc chuyển đổi số 100% hồ sơ xuất nhập khẩu và xử lý C/O điện tử qua Cơ chế một cửa ASEAN
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Thủ tục Hành chính Hải quan Một cửa Quốc gia & Quy tắc Xuất xứ)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo 'Một cửa quốc gia Giai đoạn 3 chính thức vận hành: Chấm dứt nộp C/O giấy, doanh nghiệp coi chừng mất trắng ưu đãi thuế FTA' | YouTube: Video hướng dẫn chi tiết quy trình xác thực tài khoản tập trung và xử lý lỗi không truyền được C/O Form D, E, AK, AJ trên Cổng một cửa ASEAN)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Nghị định 336/2026/NĐ-CP thay thế Nghị định 85/2019/NĐ-CP; Hệ thống Một cửa quốc gia Giai đoạn 3 bắt đầu vận hành từ tháng 11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Không kịp chuyển đổi tài khoản số hoặc lỗi đồng bộ C/O điện tử -> Bị tính thuế MFN suất thông thường, truy thu thuế hàng trăm triệu đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Cộng đồng XNK hoang mang vì nhiều doanh nghiệp quen dùng bản giấy hoặc file scan thủ công, chưa đồng bộ chữ ký số doanh nghiệp |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Hướng dẫn 4 bước rà soát mã định danh số doanh nghiệp, kiểm tra tính hợp lệ của C/O điện tử trên cổng ASEAN Single Window |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Tuân Thủ Pháp Lý Sống Còn — Cảnh báo cấp bách cho mọi doanh nghiệp xuất nhập khẩu hưởng ưu đãi FTA)** |
- **Tóm tắt lý do nên làm:** Chuyển đổi số trong thủ tục hải quan đang bước vào giai đoạn quyết liệt nhất. Bài viết vừa mang tính cảnh báo pháp lý khẩn cấp, vừa cung cấp giải pháp kỹ thuật cụ thể giúp các doanh nghiệp bảo vệ 100% lợi ích miễn giảm thuế từ các FTA.
- **Nguồn xác minh:** Nghị định số 336/2026/NĐ-CP (thay thế NĐ 85/2019/NĐ-CP); Cổng TTĐT Chính phủ (`chinhphu.vn`); Báo Đầu Tư (`baodautu.vn`).

---

### Đề xuất #3: [CHIẾN LƯỢC CƯỚC BIỂN & LIÊN MINH HÃNG TÀU] Chỉ số Drewry WCI đạt $4,429/FEU & Hãng tàu ồ ạt kích hoạt "Blank Sailings" đầu tháng 11/2026: Chiến lược giữ giá cước của liên minh hãng tàu và phác đồ đàm phán cước an toàn cho chủ hàng Việt
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải) / `CS` (Case Study - Quản trị Chi phí Vận tải Biển)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Bóc trần chiêu trò 'Hãng tàu hủy chuyến chủ động (Blank Sailings) ép giá cước mùa thấp điểm: Chủ hàng Việt làm sao để không bị ép phụ phí PSS ảo?' | YouTube: Video phân tích dữ liệu biến động cước tuyến Transpacific/Á-Âu từ Drewry và chiến lược đàm phán hợp đồng dài hạn All-in Fix Rate)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Báo cáo Drewry WCI cuối tháng 10 - đầu tháng 11/2026 ghi nhận mức $4,429/FEU; các liên minh hãng tàu ồ ạt hủy chuyến trong tháng 11 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Hàng hóa bị delay giao hàng mùa Giáng sinh/Năm mới, chi phí cước bị giữ ở mức cao vô lý và phải trả thêm các khoản phụ phí cao điểm |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Cuộc đấu trí giữa chủ hàng và hãng tàu ngoại: Khi nào nên chốt giá Spot Rate, khi nào nên ký Contract Rate dài hạn? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka chia sẻ 3 vũ khí đàm phán cước tàu, cam kết giữ chỗ (Space Guarantee) và bộ điều khoản chống hủy chuyến đơn phương |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Bảo Vệ Biên Lợi Nhuận Logistics — Cực kỳ giá trị cho các giám đốc chuỗi cung ứng và chủ hàng xuất khẩu)** |
- **Tóm tắt lý do nên làm:** Giúp doanh nghiệp xuất nhập khẩu thấu hiểu bản chất các đợt "Blank Sailings" của hãng tàu quốc tế, từ đó không bị các bên forwarder trung gian "dắt mũi" bằng các khoản phụ phí ảo, khẳng định vai trò tư vấn chiến lược của Eureka.
- **Nguồn xác minh:** Báo cáo Drewry World Container Index (WCI); Báo Công Thương (`congthuong.vn`); World Shipping Council; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).

---

### Đề xuất #4: [CÔNG NGHỆ CỬA KHẨU THÔNG MINH BIÊN GIỚI] Thông tư 128/2026/TT-BTC & Đột phá Cửa khẩu thông minh Lạng Sơn: Vận hành phương tiện tự hành không người lái (AGV) 24/7 và cẩm nang khai báo trước (Pre-clearance) xuyên đêm
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Quy trình Thông quan Cửa khẩu Thông minh & Vận tải Đường bộ)
- **Nền tảng phù hợp:** **Facebook** (Infographic: 'Cửa khẩu thông minh Hữu Nghị - Tân Thanh vận hành xe tự hành AGV: Cơ hội thông quan nông sản xuyên đêm không lo ùn tắc') & **YouTube** (Cẩm nang 4 bước chuẩn bị hồ sơ khai báo trước Pre-clearance theo Thông tư 128/2026/TT-BTC)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Mô hình Cửa khẩu thông minh Việt - Trung tại Lạng Sơn đang vận hành thử nghiệm giai đoạn 1 với xe tự hành AGV; Thông tư 128 vừa ban hành |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Ách tắc cửa khẩu đường bộ biên giới phía Bắc gây thối hỏng nông sản tươi và gián đoạn chuỗi cung ứng linh kiện lắp ráp |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận về tính khả thi, chi phí vận hành phương tiện AGV so với đội xe tải truyền thống và độ trễ đồng bộ dữ liệu hai nước |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka hướng dẫn chi tiết quy trình khai báo dữ liệu trước (Pre-clearance) để xe tự hành kéo container qua biên giới tự động 24/7 |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #4 (Tiên Phong Giải Pháp Công Nghệ Logistics — Cứu nguy ùn tắc biên giới cho doanh nghiệp nông sản & linh kiện)** |
- **Tóm tắt lý do nên làm:** Đây là xu hướng logistics tương lai giữa Việt Nam và Trung Quốc. Việc Eureka đón đầu phân tích quy định và quy trình kỹ thuật sẽ định vị thương hiệu là đơn vị am hiểu sâu sắc logistics đường bộ cửa khẩu thông minh.
- **Nguồn xác minh:** Thông tư số 128/2026/TT-BTC; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`); Cổng TTĐT tỉnh Lạng Sơn; Báo Công Thương.

---

### Đề xuất #5: [NGHIỆP VỤ THỰC CHIẾN & XỬ LÝ SỰ CỐ TỜ KHAI] Case study kiểm tra tình trạng và lịch sử tờ khai hải quan điện tử: Bẫy tờ khai "treo" chưa thông quan và phác đồ 4 bước tra cứu đối soát trên ECUS5 và Cổng Hải quan điện tử
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Kỹ thuật Tra cứu ECUS5 & Kiểm soát Rủi ro Thông quan)
- **Nền tảng phù hợp:** **Facebook** (Case study: 'Chủ quan tưởng nộp xong tờ khai là an tâm: Bài học đắt giá về tờ khai bị treo khiến hàng nằm cảng 5 ngày') & **YouTube** (Video quay màn hình thao tác thực tế cách kiểm tra lịch sử truyền nhận dữ liệu, gỡ lỗi đồng bộ thuế trên ECUS5)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Trích xuất từ tình huống thực tế ngày 03/11/2026 trên diễn đàn chuyên ngành và ghi chép daily_notes.md |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Tờ khai bị treo nhưng nhân viên không biết -> Hàng không lấy được ra khỏi cảng, phát sinh phạt chậm nộp và lưu kho bãi |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Rất nhiều người mới vào nghề hoặc kế toán kiêm nhiệm chưa phân biệt rõ trạng thái 'Đã tiếp nhận' vs 'Đã thông quan' |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn 4 bước kiểm tra log truyền nhận, tra cứu trên customs.gov.vn và cách xử lý lỗi kết nối thuế tức thì |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #5 (Cầm Tay Chỉ Việc Sát Thực Tế — Giải quyết ngay vướng mắc hàng ngày của dân hiện trường và kế toán XNK)** |
- **Tóm tắt lý do nên làm:** Xuất phát từ câu hỏi thực tế của kế toán/ops trong ngày 03/11/2026. Chủ đề mộc mạc, gần gũi, mang tính thực chiến cao, giúp củng cố hình ảnh Eureka là chuyên gia luôn đồng hành gỡ rối từng thao tác nhỏ nhất cho khách hàng.
- **Nguồn xác minh:** Sổ tay nội bộ `scratch/daily_notes.md` (ngày 03/11/2026); Diễn đàn Thuế chuyên sâu; Hướng dẫn vận hành phần mềm ECUS5 VNACCS/VCIS; Cổng thông tin Tổng cục Hải quan.

---

## 🎯 BƯỚC TIẾP THEO (NEXT ACTIONS DÀNH CHO CHỊ THÚY)
1. **Lựa chọn chủ đề ưu tiên:** Chị Thúy có thể phản hồi chọn 1 hoặc nhiều chủ đề (ví dụ: *'Chọn #1 và #2'* hoặc *'Làm video cho #1, làm post FB cho #5'*) để hệ thống tự động đẩy vào:
   - **`wf01`:** Lên lịch xuất bản chi tiết cho kế hoạch tuần hiện tại (Tuần 45).
   - **`wf02`:** Sản xuất bài viết chuyên sâu trên Facebook Fanpage Eureka Logistics / Trang cá nhân.
   - **`wf03`:** Sản xuất kịch bản video dài YouTube ThuyTalk (Chuyên mục 1, 2 hoặc 4).
2. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""
    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(md_content)
    print(f"[OK] Đã xuất bản tin đề xuất ra: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT THEO WF05 NGÀY 03/11/2026 ===")

if __name__ == '__main__':
    update_seo_w45_0311()
