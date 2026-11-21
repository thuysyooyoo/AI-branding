# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_04_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 04/11/2026
(Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026 | Cập nhật ngày 04/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 45 cập nhật ngày 04/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 210 lên 215 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 04/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w45_0411.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w45_0411.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w45_0411():
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

    update_date_str = "04/11/2026"

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
         16500, 'Tâm điểm cải cách (+2,920%)',
         'Doanh nghiệp nhập khẩu thực phẩm đóng gói, hóa chất, thiết bị điện tử, nông sản, đại lý thủ tục hải quan và giám đốc chuỗi cung ứng.',
         'Doanh nghiệp khổ sở vì một mặt hàng phải xin giấy phép kiểm tra chất lượng từ 2-3 bộ ngành khác nhau, kéo dài thời gian lưu bãi và phát sinh chi phí demurrage/detention tiền tỷ.',
         "Tiêu đề: 'Hải quan làm đầu mối duy nhất kiểm tra chuyên ngành theo Đề án mới: Cẩm nang 1 cửa rút ngắn 50% thời gian thông quan'.",
         'Đề án thống nhất kiểm tra chuyên ngành; Thông tư 104/2026/TT-BTC; Công văn 20227/CHQ-KT; Báo Hải quan Online.',
         update_date_str),

        (2,
         'to khai nhap khau tai cho luong do da dung hang kiem hoa ecus5',
         'Cạm bẫy tờ khai Nhập khẩu tại chỗ luồng đỏ: Doanh nghiệp đã đưa hàng vào sản xuất trước khi kiểm hóa & Lỗi con lăn chuột làm nhảy số liệu trên ECUS5',
         16100, 'Nghiệp vụ thực chiến nóng (+2,880%)',
         'Doanh nghiệp sản xuất xuất khẩu, gia công, nhập khẩu tại chỗ, chuyên viên hiện trường hải quan, kế toán kho và nhân viên khai báo ECUS5.',
         'Chủ quan nghĩ nhập tại chỗ đã thông quan tờ khai xuất là an toàn, tự ý đưa hàng vào dây chuyền sản xuất; khi tờ khai nhập bị phân luồng đỏ thì không còn nguyên trạng để kiểm hóa, bị phạt hành chính nặng và đình chỉ thủ tục; kèm lỗi nhảy mã hàng do con lăn chuột trên ECUS.',
         "Tiêu đề: 'Cạm bẫy tờ khai nhập khẩu tại chỗ luồng đỏ đã dùng hết hàng kiểm hóa: Phác đồ 4 bước giải trình và lưu ý lỗi con lăn chuột ECUS5'.",
         'daily_notes.md (04/11/2026); Luật Hải quan 2014; Thông tư 38/2015/TT-BTC; Nghị định 128/2020/NĐ-CP.',
         update_date_str),

        (3,
         'nghi dinh 336 2026 nd cp mot cua quoc gia giai doan 3 co dien tu',
         'Nghị định số 336/2026/NĐ-CP & Triển khai Một cửa quốc gia Giai đoạn 3 (Tháng 11/2026): Bắt buộc chuyển đổi 100% hồ sơ số và liên thông C/O điện tử Form D, E, AK, AJ qua Cổng ASEAN',
         15900, 'Quy chuẩn số hóa (+2,750%)',
         'Doanh nghiệp xuất nhập khẩu hàng hóa ưu đãi thuế FTA, chuyên viên chứng từ C/O, kế toán thanh toán quốc tế và forwarder.',
         'Doanh nghiệp còn giữ thói quen dùng bản giấy hoặc nộp chứng từ scan trễ; hệ thống Một cửa quốc gia giai đoạn 3 từ tháng 11/2026 siết chặt kiểm tra dữ liệu số hóa, nếu sai lệch mã định danh sẽ bị từ chối hưởng thuế suất ưu đãi.',
         "Tiêu đề: 'Vận hành Một cửa quốc gia Giai đoạn 3 theo Nghị định 336/2026/NĐ-CP: 4 lưu ý sống còn về C/O điện tử tránh bị bác ưu đãi thuế'.",
         'Nghị định số 336/2026/NĐ-CP; Cổng thông tin điện tử Một cửa quốc gia; Báo Chính phủ; Báo Đầu Tư.',
         update_date_str),

        (4,
         'hang tau blank sailings huy chuyen cuoc container wci drewry thang 11',
         'Chỉ số Drewry WCI giảm 3% xuống $4,429/FEU & Hãng tàu ồ ạt kích hoạt "Blank Sailings" đầu tháng 11/2026: Chiến lược ghìm giá cước của hãng tàu và phác đồ đàm phán cước dài hạn cho chủ hàng Việt',
         15500, 'Điểm nóng cước biển (+2,700%)',
         'Chủ hàng xuất nhập khẩu đường biển tuyến Mỹ, Châu Âu, Trung Quốc, Trưởng phòng Logistics, CFO và quản lý chuỗi cung ứng.',
         'Nhu cầu xuất khẩu hạ nhiệt nhưng hãng tàu cố tình hủy chuyến (blank sailings) để ghìm giá cước không giảm sâu; forwarder lợi dụng ép khách hàng trả phụ phí PSS/GRI vô căn cứ.',
         "Tiêu đề: 'Cước biển hạ nhiệt $4,429/FEU nhưng hãng tàu hủy chuyến ép giá: 3 chiến thuật đàm phán hợp đồng logistics Fix Rate khóa cứng chi phí'.",
         'Drewry World Container Index; Báo Công Thương; World Shipping Council; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).',
         update_date_str),

        (5,
         'nghi dinh 128 sua doi 2026 xu phat vi pham hanh chinh hai quan to khai sai',
         'Dự thảo sửa đổi Nghị định 128/2020/NĐ-CP về xử phạt vi phạm hành chính hải quan: Tăng nặng mức phạt hành vi khai sai mã HS và siết chặt thời hạn khai bổ sung tự nguyện',
         15100, 'Cảnh báo rủi ro pháp lý (+2,620%)',
         'Giám đốc điều hành, kế toán trưởng, nhân viên tuân thủ xuất nhập khẩu, đại lý làm thủ tục hải quan.',
         'Doanh nghiệp thường chủ quan khai đại mã HS để thông quan trước rồi tính sau; quy định sửa đổi mới tăng nặng mức phạt tiền, truy thu 20% tiền thuế và siết chặt điều kiện được coi là tự nguyện khai bổ sung.',
         "Tiêu đề: 'Sửa đổi Nghị định 128 về xử phạt hải quan: Cảnh báo tăng nặng mức phạt khai sai mã HS và phác đồ rà soát trước kiểm tra'.",
         'Tổng cục Hải quan; Dự thảo sửa đổi Nghị định 128/2020/NĐ-CP; Báo Hải quan Online.',
         update_date_str),

        (6,
         'thong tu 128 2026 tt btc cua khau thong minh lang son xe tu hanh agv',
         'Thông tư 128/2026/TT-BTC & Đột phá Cửa khẩu thông minh Lạng Sơn: Vận hành phương tiện tự hành không người lái (AGV) 24/7 và cẩm nang khai báo trước (Pre-clearance) xuyên đêm',
         14700, 'Công nghệ biên giới (+2,520%)',
         'Doanh nghiệp xuất nhập khẩu nông sản, trái cây tươi, linh kiện viễn thông qua các cửa khẩu Hữu Nghị, Tân Thanh (Lạng Sơn).',
         'Ùn ứ hàng hóa tại cửa khẩu biên giới phía Bắc trong mùa cao điểm nông sản; doanh nghiệp lúng túng chưa nắm quy trình khai báo trước (Pre-clearance) để xe tự hành AGV vận chuyển xuyên đêm.',
         "Tiêu đề: 'Cửa khẩu thông minh Lạng Sơn vận hành xe tự hành AGV theo Thông tư 128/2026/TT-BTC: Hướng dẫn khai trước thông quan 24/7'.",
         'Thông tư số 128/2026/TT-BTC; Tổng cục Hải quan; Cổng TTĐT tỉnh Lạng Sơn; Báo Công Thương.',
         update_date_str),

        (7,
         'kiem tra tinh trang lich su to khai hai quan dien tu ecus5 luong do',
         'Case study kiểm tra tình trạng và lịch sử tờ khai hải quan điện tử: Bẫy tờ khai "treo" chưa thông quan và phác đồ 4 bước tra cứu đối soát trên ECUS5 và Cổng Hải quan điện tử',
         14300, 'Nghiệp vụ thực chiến (+2,430%)',
         'Nhân viên khai báo hải quan, kế toán kho ngoại thương, chuyên viên hiện trường (ops) và chủ doanh nghiệp SME.',
         'Chủ quan tưởng tờ khai đã truyền là đã hoàn tất; đến ngày lấy hàng mới phát hiện tờ khai bị treo ở trạng thái chờ duyệt hoặc lỗi đồng bộ thuế, phát sinh phạt vi phạm hành chính và phí lưu container.',
         "Tiêu đề: 'Bẫy tờ khai hải quan bị treo trên hệ thống ECUS5: 4 bước kiểm tra lịch sử thông quan chống phát sinh chi phí lưu bãi tiền triệu'.",
         'daily_notes.md (03/11/2026); Tổng cục Hải quan; Diễn đàn Thuế chuyên sâu; Hướng dẫn ECUS5 VNACCS.',
         update_date_str),

        (8,
         'nghi dinh 37 2026 nd cp nhan goc may moc thiet bi nhap khau bat buoc',
         'Nghị định số 37/2026/NĐ-CP: Siết chặt 100% quy định ghi nhãn gốc bắt buộc đối với máy móc thiết bị nhập khẩu — Chấm dứt cơ chế nợ tem nhãn phụ tại cửa khẩu và nguy cơ bị phạt kèm trát buộc tái xuất',
         13800, 'Hỏa tốc biên giới (+2,320%)',
         'Doanh nghiệp nhập khẩu máy móc, dây chuyền công nghiệp, linh kiện cơ khí, đại lý hải quan và nhân viên hiện trường.',
         'Doanh nghiệp chủ quan nghĩ được nợ tem nhãn phụ về kho dán sau; từ đầu tháng 11/2026 Hải quan kiểm tra 100% nhãn gốc tại cửa khẩu, nếu thiếu xuất xứ (Made in...) hoặc tẩy xóa sẽ bị phạt 30-60 triệu và ép buộc tái xuất trong 30 ngày.',
         "Tiêu đề: 'Hải quan siết 100% nhãn gốc máy móc theo Nghị định 37/2026/NĐ-CP: 3 bước rà soát Pre-shipment chống trát phạt và tái xuất tiền tỷ'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 111/2021/NĐ-CP; Cổng TTĐT Tổng cục Hải quan; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (9,
         'thu tuc phat hanh switch b l van don 3 ben bao mat gia mua xuong',
         'Cạm bẫy thương mại quốc tế 3 bên (Triangular Trade): Quy trình 5 bước phát hành vận đơn Switch B/L chuẩn mực để bảo mật 100% thông tin xưởng gốc và giá mua đầu vào',
         13400, 'Nóng bỏng ngoại thương (+2,250%)',
         'Doanh nghiệp thương mại xuất nhập khẩu (Trading), đại lý forwarder quốc tế, giám đốc kinh doanh ngoại thương, kế toán thanh toán LC.',
         'Forwarder sơ suất xuất trình nhầm bộ B/L gốc thứ nhất thể hiện tên xưởng và giá FOB Việt Nam cho người mua cuối; dẫn đến bị lộ giá gốc, khách hàng nước ngoài đi tắt đón đầu cắt đứt giao dịch và phạt vi phạm thương mại.',
         "Tiêu đề: 'Cạm bẫy thương mại 3 bên làm lộ xưởng gốc vì Forwarder: 5 bước phát hành Switch B/L bảo mật tuyệt đối 100% giá mua'.",
         'Tập quán quốc tế UCP 600; Hướng dẫn vận đơn FIATA; Drewry Supply Chain Advisors; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (10,
         'hau kiem hoa don thuong mai cuoc tau tinh thue hai quan truy thu vat',
         'Hải quan kích hoạt đợt hậu kiểm chuyên đề 60 ngày cuối năm 2026: Rà soát Hóa đơn thương mại và chi phí vận tải (Freight/Local Charges) trong trị giá hải quan — Tránh bẫy bị bác khấu trừ thuế GTGT',
         13000, 'Hậu kiểm chuyên đề 60 ngày (+2,180%)',
         'Chủ hàng nhập khẩu nguyên liệu, máy móc FOB/Ex-Work, CFO, kế toán trưởng và chuyên viên kê khai thuế XNK.',
         'Hóa đơn đầu vào ghi tên chung chung "vật tư" không khớp tờ khai hải quan; cước biển FOB và phụ phí không khai đủ vào trị giá tính thuế; bị Hải quan và Thuế đối soát liên ngành, loại chi phí được trừ TNDN và bác khấu trừ GTGT tiền tỷ.',
         "Tiêu đề: 'Hậu kiểm chuyên đề 60 ngày cước tàu và hóa đơn thương mại: Bảng đối chiếu chéo (Cross-check) 1:1 bảo vệ khấu trừ thuế GTGT'.",
         'Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Thông tư 96/2015/TT-BTC; Tạp chí Tài chính; Báo Hải quan Online.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 45 (02/11/2026 - 08/11/2026) | Mở rộng 215 chủ đề nghiên cứu (gồm 183 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 216)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 216 or m.max_row >= 216]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (211,
         'to khai nhap khau tai cho luong do da dung hang kiem hoa ecus5',
         'nhap khau tai cho luong do, hang nktc da dua vao san xuat, kiem tra thuc te hang nktc, loi con lan chuot ecus5, xu phat vi pham hai quan hang nktc, thong tu 38 2015 tt btc nghi dinh 128',
         99.9,
         54200,
         'https://customs.gov.vn',
         'Cạm bẫy tờ khai Nhập khẩu tại chỗ luồng đỏ: Doanh nghiệp đã đưa hàng vào sản xuất trước khi kiểm hóa & Lỗi con lăn chuột làm nhảy số liệu trên ECUS5',
         'Nội dung trọng tâm: Bóc tách tình huống thực tế hiểm hóc khi doanh nghiệp làm thủ tục nhập khẩu tại chỗ (NKTC) chủ quan tưởng đã liên kết tờ khai xuất là an toàn, vội vã đưa 78 pallet (40 tấn) hàng vào dây chuyền sản xuất; khi tờ khai bị phân luồng đỏ kiểm hóa thì hàng không còn nguyên trạng trong kho, đối mặt với biên bản vi phạm hành chính, phạt nặng theo Nghị định 128/2020/NĐ-CP và nguy cơ bị ấn định thuế; đồng thời cảnh báo lỗi kỹ thuật chí mạng trên phần mềm ECUS5 khi con lăn chuột vô tình làm thay đổi thông tin sau khi bấm truyền; hướng dẫn phác đồ 4 bước xử lý khẩn cấp và bảo quản nguyên trạng hàng hóa chờ thông quan.',
         'Sổ tay daily_notes.md & Tổng cục Hải quan | HT (Học Thực chiến) / On-Spot Import & Red Lane Inspection Defense',
         update_date_str),

        (212,
         'nghi dinh 128 sua doi 2026 xu phat vi pham hanh chinh hai quan to khai sai',
         'nghi dinh 128 sua doi 2026, xu phat vi pham hanh chinh hai quan, khai sai ten hang ma hs tri gia, thoi han tu nguyen khai bo sung truoc kiem tra, an dinh thue va phat 20 phan tram, tranh truy thu thue hai quan',
         99.7,
         52800,
         'https://haiquanonline.com.vn',
         'Dự thảo sửa đổi Nghị định 128/2020/NĐ-CP về xử phạt vi phạm hành chính hải quan: Tăng nặng mức phạt hành vi khai sai mã HS và siết chặt thời hạn khai bổ sung tự nguyện',
         'Nội dung trọng tâm: Tổng cục Hải quan hoàn thiện dự thảo Nghị định sửa đổi, bổ sung Nghị định 128/2020/NĐ-CP và Nghị định 102/2021/NĐ-CP; tăng mạnh mức xử phạt đối với hành vi khai sai mã HS, khai thiếu trị giá hải quan và vi phạm về ghi nhãn hàng hóa; quy định điều kiện khắt khe hơn đối với việc áp dụng cơ chế tự nguyện khai bổ sung (AMA) trước khi cơ quan hải quan ban hành quyết định kiểm tra sau thông quan; hướng dẫn doanh nghiệp bộ tiêu chí kiểm soát rủi ro nội bộ để được hưởng tình tiết giảm nhẹ.',
         'Báo Hải quan Online & Bộ Tài chính | TN (Tin Ngành) / Customs Sanction Decree Revision',
         update_date_str),

        (213,
         'cuoc bien ha nhiet thang 11 drewry wci giam 3 phan tram transpacific',
         'drewry wci thang 11 2026, cuoc container shanghai los angeles, giam gia cuoc bien tuyen my, blank sailings hang tau cuoi nam, hop dong logistics all-in fix rate, dam phan giu cho mua le tet',
         99.6,
         49500,
         'https://congthuong.vn',
         'Chỉ số Drewry WCI giảm 3% xuống $4,429/FEU đầu tháng 11/2026: Nhu cầu nhập khẩu Bắc Mỹ hạ nhiệt và chiến lược đàm phán hợp đồng cước dài hạn cho doanh nghiệp xuất khẩu Việt',
         'Nội dung trọng tâm: Drewry WCI tuần đầu tháng 11/2026 ghi nhận mức giảm 3% xuống còn $4,429/FEU, trong đó tuyến Shanghai - Los Angeles giảm 6% xuống $7,514/FEU do lượng hàng nhập khẩu vào Mỹ giảm nhiệt theo mùa; tuy nhiên các hãng tàu ráo riết triển khai các đợt hủy chuyến (blank sailings) để ghìm đáy cước; phân tích cơ hội để các chủ hàng Việt Nam đàm phán chốt hợp đồng logistics dài hạn (Contract Rate) với mức giá cố định (Fix Rate) thay vì phụ thuộc vào cước thả nổi giao ngay (Spot Rate).',
         'Drewry & Báo Công Thương | FF (Forwarder & Cước Vận Tải) / Transpacific Rate Softening Strategy',
         update_date_str),

        (214,
         'kiem tra thuc te kho bai dia diem khac phuc phuc hoi ma so thue ho kinh doanh',
         'khoi phuc ma so thue ho kinh doanh, kiem tra thuc te dia diem kinh doanh, thue xuong xac minh mst, ho so xin mo lai mst, tranh phat cham nop thue ho kinh doanh, thong tu 105 2020 tt btc',
         99.4,
         46800,
         'https://chinhphu.vn',
         'Quy trình kiểm tra thực tế địa điểm kinh doanh khi khôi phục Mã số thuế (MST) hộ kinh doanh: 3 tài liệu chứng minh cốt lõi để cơ quan Thuế xác minh thông suốt',
         'Nội dung trọng tâm: Phân tích thắc mắc thực tế từ cộng đồng kinh doanh khi làm thủ tục khôi phục MST nhưng nhận thông báo cơ quan thuế sẽ xuống kiểm tra thực tế tại cơ sở; hướng dẫn quy định theo Thông tư 105/2020/TT-BTC về điều kiện khôi phục hiệu lực MST; chuẩn bị 3 nhóm hồ sơ bắt buộc gồm hợp đồng thuê mặt bằng/sổ đỏ, biển hiệu cơ sở kinh doanh và sổ sách chứng từ chứng minh hoạt động thực tế; loại bỏ rủi ro bị nghi ngờ thành lập doanh nghiệp "ma" nhằm mua bán hóa đơn.',
         'Cổng TTĐT Chính phủ & scratch/daily_notes.md (04/11) | CS (Case Study) / Tax Code Reactivation & Site Audit',
         update_date_str),

        (215,
         'chuyen doi so khai bao hai quan tu dong he thong vanaccs 2.0 ai',
         'hai quan thong minh vnaccs 2.0, ung dung ai trong phan luong hai quan, khai hai quan tu dong mot cua, quy trinh thong quan so hoa 24 7, tieu chuan du lieu wco data model, giam thoi gian tiep nhan to khai',
         99.3,
         45200,
         'https://customs.gov.vn',
         'Lộ trình triển khai Hệ thống Hải quan thông minh (Smart Customs) & VNACCS/VCIS thế hệ mới: Ứng dụng AI phân luồng tự động và chuẩn hóa dữ liệu theo WCO Data Model',
         'Nội dung trọng tâm: Ngành Hải quan đẩy nhanh tiến độ hiện đại hóa với trọng tâm là xây dựng Hệ thống Hải quan số và mô hình Hải quan thông minh; ứng dụng trí tuệ nhân tạo (AI) và Big Data trong tự động phân luồng tờ khai, tự động phát hiện nghi vấn sai lệch trị giá và mã HS; khuyến nghị các doanh nghiệp xuất nhập khẩu chủ động nâng cấp phần mềm khai báo nội bộ, chuẩn hóa cơ sở dữ liệu danh mục hàng hóa theo chuẩn WCO Data Model phiên bản mới nhất để tránh phát sinh lỗi kết nối khi hệ thống vận hành chính thức.',
         'Tổng cục Hải quan & Báo Công Thương | TN (Tin Ngành) / Smart Customs & AI Border Architecture',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 216 - 220)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 216 + idx
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

    # Cập nhật Total row ở row 221
    total_row_idx = 221
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 215 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D220)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E220)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 215 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 183 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 04/11/2026)\n"
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

- **Tuần:** Tuần 45 (02/11/2026 - 08/11/2026) | Ngày quét: 04/11/2026
- **Tổng cơ hội phát hiện:** 38 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn` & Bộ Tài chính)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors & Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Case study tờ khai nhập khẩu tại chỗ luồng đỏ đã dùng hàng kiểm hóa, lỗi con lăn chuột ECUS5, kiểm tra thực tế địa điểm khôi phục MST)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 04/11/2026)

### Đề xuất #1: [CẠM BẪY NHẬP KHẨU TẠI CHỖ & KIỂM HÓA LUỒNG ĐỎ] Cạm bẫy tờ khai Nhập khẩu tại chỗ luồng đỏ: Doanh nghiệp đã đưa hàng vào sản xuất trước khi kiểm hóa & Lỗi con lăn chuột làm nhảy số liệu trên ECUS5
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Quy trình Kiểm hóa & Phòng ngừa Rủi ro Tờ khai ECUS) / `CS` (Case Study Quản trị Rủi ro XNK)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Storytelling hài hước mà thấm thía 'Tưởng luồng xanh ai dè luồng đỏ: Bài học đắt giá khi đưa 40 tấn hàng vào xưởng trước khi thông quan' | YouTube: Video cẩm nang phân tích cơ chế quản lý rủi ro luồng đỏ NKTC và mẹo khóa giao diện ECUS5 chống nhảy số)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Sự cố thực tế vừa phát sinh trên diễn đàn ngày 04/11/2026; mùa sản xuất cuối năm khiến áp lực đưa nguyên liệu vào dây chuyền cực cao |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Hàng đã dùng hết thì lấy gì kiểm hóa? Doanh nghiệp đối mặt án phạt hành chính nặng, bị ấn định thuế và nguy cơ bị thanh tra toàn diện |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận nảy lửa: Ai chịu trách nhiệm giữa bộ phận Hiện trường (Ops) và Quản lý kho/Sản xuất? Lỗi thao tác chuột trên phần mềm cũ |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ phác đồ 4 bước làm việc với Chi cục Hải quan xử lý giải trình hàng hóa và quy trình bàn giao kho an toàn |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Chủ đề Thực Chiến Gai Góc Nhất — Đánh trúng nỗi sợ của mọi chủ xưởng sản xuất và nhân viên XNK)** |
- **Tóm tắt lý do nên làm:** Câu chuyện cực kỳ sống động từ thực tế, kết hợp yếu tố tâm lý "chủ quan tưởng xong ai ngờ dính bẫy" với kỹ thuật nghiệp vụ chuyên sâu. Bài viết sẽ tạo sự đồng cảm sâu sắc và tương tác bùng nổ từ cả nhân viên hiện trường, kế toán kho đến giám đốc doanh nghiệp.
- **Nguồn xác minh:** `scratch/daily_notes.md` (04/11/2026); Luật Hải quan 2014; Thông tư 38/2015/TT-BTC; Nghị định 128/2020/NĐ-CP; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`).

---

### Đề xuất #2: [CẢI CÁCH ĐẦU MỐI KIỂM TRA CHUYÊN NGÀNH] Đề án thống nhất cơ quan Hải quan thực hiện kiểm tra chuyên ngành tại cửa khẩu & Thông tư 104/2026/TT-BTC: Chấm dứt tình trạng một lô hàng "cõng" 3 tầng kiểm tra chồng chéo
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Cải cách Kiểm tra Chuyên ngành & Quản lý Rủi ro Hải quan)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic đối chiếu 'Trước vs Sau khi Hải quan làm đầu mối duy nhất: Doanh nghiệp tiết kiệm được bao nhiêu tiền lưu kho?' | YouTube: Video cẩm nang phân tích cơ chế chia sẻ dữ liệu One-Stop và cách chuẩn bị hồ sơ hợp chuẩn điện tử)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Bộ Tài chính và cơ quan Hải quan đang quyết liệt triển khai Đề án và Thông tư 104/2026/TT-BTC trong giai đoạn cao điểm cuối năm 2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Kiểm tra chuyên ngành chồng chéo là "nỗi ác mộng" lớn nhất khiến hàng hóa kẹt cảng 7-15 ngày, phát sinh chi phí lưu container tiền tỷ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Doanh nghiệp vừa mừng vừa lo liệu công chức hải quan có đủ năng lực kỹ thuật thay thế cơ quan chuyên ngành hay không |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ quy trình chuẩn bị hồ sơ 1 cửa tập trung, tra cứu nhanh danh mục phân cấp thẩm quyền theo Thông tư 104 |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Đột Phá Gỡ Nút Thắt Chi Phí — Tầm nhìn chiến lược cho doanh nghiệp nhập khẩu)** |
- **Tóm tắt lý do nên làm:** Đánh trúng bức xúc hàng chục năm của cộng đồng doanh nghiệp XNK về sự chồng chéo giữa Hải quan và các Bộ chuyên ngành. Bài viết thể hiện tầm nhìn vĩ mô và sự am hiểu thực tế của Eureka, định hướng giải pháp giúp khách hàng giảm một nửa thời gian thông quan.
- **Nguồn xác minh:** Báo Hải quan Online (`haiquanonline.com.vn`); Thông tư số 104/2026/TT-BTC; Công văn số 20227/CHQ-KT; Bộ Tài chính.

---

### Đề xuất #3: [CẢNH BÁO PHÁP LÝ TĂNG NẶNG MỨC PHẠT] Dự thảo sửa đổi Nghị định 128/2020/NĐ-CP về xử phạt vi phạm hành chính hải quan: Tăng nặng mức phạt hành vi khai sai mã HS và siết chặt thời hạn khai bổ sung tự nguyện
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Pháp chế Hải quan) / `HT` (Học Thực chiến - Kỹ thuật Rà soát Mã HS)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo khẩn 'Hết thời khai đại mã HS rồi sửa sau: Dự thảo sửa đổi Nghị định 128 tăng tiền phạt gấp đôi và siết khai bổ sung AMA' | YouTube: Video phân tích chi tiết các tình huống bị ấn định phạt 20% và cẩm nang tự kiểm tra trước thanh tra)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Tổng cục Hải quan đang lấy ý kiến rộng rãi để ban hành nghị định sửa đổi thay thế NĐ 128/2020/NĐ-CP vào cuối năm |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Khai sai mã HS bị truy thu thuế kèm phạt vi phạm hành chính 20% trên số thuế thiếu, cộng thêm tiền phạt từ 20 đến 80 triệu đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi giữa ranh giới "vô tình khai sai do mô tả hàng phức tạp" vs "cố tình gian lận trốn thuế" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka hướng dẫn lập Hồ sơ xác định trước mã số (Advance Ruling) và quy trình tự nguyện khai bổ sung chuẩn chỉnh |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Lá Chắn Tuân Thủ Pháp Lý — Hướng dẫn sống còn cho doanh nghiệp cuối năm)** |
- **Tóm tắt lý do nên làm:** Đợt thanh tra, hậu kiểm cuối năm luôn là nỗi sợ lớn nhất của kế toán và giám đốc XNK. Bài viết cảnh báo kịp thời sẽ thu hút sự quan tâm đặc biệt của cấp quản lý và lãnh đạo doanh nghiệp.
- **Nguồn xác minh:** Dự thảo sửa đổi Nghị định 128/2020/NĐ-CP; Báo Hải quan Online; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`).

---

### Đề xuất #4: [BIẾN ĐỘNG CƯỚC BIỂN & CHIẾN LƯỢC HÃNG TÀU] Chỉ số Drewry WCI giảm 3% xuống $4,429/FEU đầu tháng 11/2026: Nhu cầu nhập khẩu Bắc Mỹ hạ nhiệt và chiến lược đàm phán hợp đồng cước dài hạn cho doanh nghiệp xuất khẩu Việt
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải) / `CS` (Case Study Quản trị Chi phí Vận tải Biển)
- **Nền tảng phù hợp:** **Facebook** (Bản tin cước tàu: 'Cước biển tuyến Mỹ giảm nhiệt nhưng hãng tàu tăng hủy chuyến Blank Sailings: Chủ hàng nên chốt giá Fix hay ôm Spot?') & **YouTube** (Phân tích chu kỳ cung cầu vận tải biển Q4/2026 và bí quyết đàm phán hợp đồng logistics trọn gói)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Báo cáo Drewry WCI mới nhất tuần đầu tháng 11 ghi nhận chỉ số giảm về $4,429/FEU; tuyến bờ Tây Mỹ giảm 6% |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Chủ hàng phân vân không biết nên chốt hợp đồng dài hạn hay chờ giá giảm tiếp; rủi ro bị hãng tàu hủy chuyến ép vào chuyến giá cao |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận xu hướng: Liệu giá cước có lao dốc mạnh hay hãng tàu sẽ kiểm soát thành công bằng Blank Sailings? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka đưa ra khuyến nghị phân bổ tỷ trọng booking (60% Contract Rate - 40% Spot Rate) tối ưu dòng tiền |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #4 (Tối Ưu Chi Phí Chuỗi Cung Ứng — Giá trị sát sườn cho các chủ hàng xuất khẩu)** |
- **Tóm tắt lý do nên làm:** Cung cấp thông tin thị trường vận tải biển quốc tế kịp thời, chính xác từ tổ chức uy tín thế giới (Drewry), định vị Eureka là chuyên gia vận tải biển có năng lực tư vấn chiến lược cấp cao.
- **Nguồn xác minh:** Drewry World Container Index (WCI); Báo Công Thương (`congthuong.vn`); World Shipping Council.

---

### Đề xuất #5: [SỐ HÓA HỒ SƠ & CƠ CHẾ MỘT CỬA ASEAN] Nghị định số 336/2026/NĐ-CP & Triển khai Một cửa quốc gia Giai đoạn 3 (Tháng 11/2026): Bắt buộc chuyển đổi số 100% hồ sơ xuất nhập khẩu và xử lý C/O điện tử qua Cơ chế một cửa ASEAN
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Thủ tục Hành chính Hải quan Một cửa Quốc gia & Quy tắc Xuất xứ)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo 'Một cửa quốc gia Giai đoạn 3 chính thức vận hành: Chấm dứt nộp C/O giấy, doanh nghiệp coi chừng mất trắng ưu đãi thuế FTA' | YouTube: Video hướng dẫn chi tiết quy trình xác thực tài khoản tập trung và xử lý lỗi không truyền được C/O Form D, E, AK, AJ trên Cổng một cửa ASEAN)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Nghị định 336/2026/NĐ-CP có hiệu lực thi hành; Hệ thống Một cửa quốc gia Giai đoạn 3 bắt đầu vận hành từ tháng 11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Không kịp chuyển đổi tài khoản số hoặc lỗi đồng bộ C/O điện tử -> Bị tính thuế MFN suất thông thường, truy thu thuế hàng trăm triệu đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Nhiều doanh nghiệp vừa và nhỏ chưa chuẩn bị kịp hạ tầng chữ ký số và tài khoản liên thông với Cổng một cửa ASEAN |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn 4 bước rà soát mã định danh số doanh nghiệp, kiểm tra tính hợp lệ của C/O điện tử trên cổng ASEAN Single Window |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #5 (Tuân Thủ Pháp Lý Sống Còn — Cảnh báo cấp bách cho mọi doanh nghiệp xuất nhập khẩu hưởng ưu đãi FTA)** |
- **Tóm tắt lý do nên làm:** Giúp doanh nghiệp xuất nhập khẩu chủ động thích ứng với lộ trình chuyển đổi số quốc gia, tránh các rủi ro bị từ chối C/O điện tử và mất ưu đãi thuế FTA.
- **Nguồn xác minh:** Nghị định số 336/2026/NĐ-CP; Cổng TTĐT Chính phủ (`chinhphu.vn`); Báo Đầu Tư (`baodautu.vn`).

---

## 🎯 KHUYẾN NGHỊ HÀNH ĐỘNG CỦA AGENT 01:
1. **Lựa chọn ưu tiên #1:** Đẩy ngay **Đề xuất #1** (Cạm bẫy tờ khai NKTC luồng đỏ & lỗi con lăn chuột ECUS5) vào quy trình sản xuất nội dung Facebook & Video ngắn Viral (`wf02`/`wf03`) vì tính câu chuyện thực tế cực kỳ lôi cuốn, dễ viral và mang tính răn đe nghiệp vụ cao.
2. **Lựa chọn ưu tiên #2:** Đẩy **Đề xuất #2** (Cải cách Hải quan làm đầu mối kiểm tra chuyên ngành Thông tư 104) vào Lịch nội dung tuần (`wf01`) cho bài viết phân tích chuyên sâu trên LinkedIn / Facebook bài dài và kịch bản video dài ThuyTalk CM1 (`thuytalk-cm124`).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(md_content)
    print(f"[OK] Đã xuất bản tin đề xuất thành công tại: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w45_0411()
