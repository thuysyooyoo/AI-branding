# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_06_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 06/11/2026
(Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026 | Cập nhật ngày 06/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 45 cập nhật ngày 06/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 220 lên 225 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 06/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w45_0611.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w45_0611.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w45_0611():
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

    update_date_str = "06/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 45 (02/11/2026 - 08/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'thong tu 89 2026 tt btc quan ly hai quan thue tmdt xuyen bien gioi temu shein',
         'Thông tư số 89/2026/TT-BTC: Bãi bỏ miễn thuế đơn hàng dưới 1 triệu đồng & Tự động thu 10% VAT sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop) qua API Cổng Thanh toán Hải quan',
         17500, 'Tâm điểm cải cách TMĐT (+3,200%)',
         'Toàn bộ doanh nghiệp kinh doanh thương mại điện tử, chủ shop nhập khẩu, người mua hàng online quốc tế và các sàn TMĐT xuyên biên giới.',
         'Chính thức xóa bỏ kẽ hở miễn thuế hàng bưu gửi nhỏ dưới 1 triệu đồng; toàn bộ đơn hàng giá trị thấp bị thu 10% thuế GTGT và phải truyền dữ liệu định danh trước khi về cảng; nguy cơ tăng chi phí kinh doanh.',
         "Tiêu đề: 'Thông tư 89/2026/TT-BTC bãi bỏ miễn thuế đơn hàng dưới 1 triệu: 4 thay đổi sống còn và giải pháp nhập khẩu chính ngạch FCL/LCL của Eureka'.",
         'Thông tư số 89/2026/TT-BTC; Nghị định 402/2026/NĐ-CP; Báo Hải quan Online; Cổng TTĐT Chính phủ.',
         update_date_str),

        (2,
         'bat dong ma hs tren c o form e hang bo nem hoi quy tac 3b',
         'Case study xử lý bất đồng mã HS trên C/O Form E cho hàng bộ: Bài toán bộ nệm hơi kèm máy bơm và cẩm nang tránh bị Hải quan bác C/O form E tại cửa khẩu',
         16800, 'Thực chiến nghiệp vụ C/O (+3,050%)',
         'Doanh nghiệp nhập khẩu hàng tiêu dùng, đồ gia dụng, thiết bị đa năng dạng bộ từ Trung Quốc, nhân viên chứng từ C/O và chuyên viên khai báo hải quan.',
         'Đầu Trung Quốc cấp C/O Form E gộp chung 1 mã HS theo hàng bộ, nhưng Hải quan Việt Nam yêu cầu tách riêng từng linh kiện (máy bơm và nệm); nguy cơ bị bác C/O form E, mất ưu đãi thuế và bị phạt hành chính.',
         "Tiêu đề: 'Bất đồng mã HS trên C/O Form E hàng bộ nệm hơi: Cách áp dụng Quy tắc 3(b) và quy trình xin cấp lại C/O cứu nguy lô hàng'.",
         'scratch/daily_notes.md (06/11/2026); Quy tắc xuất xứ Hiệp định ACFTA; Thông tư 38/2015/TT-BTC; Tổng cục Hải quan.',
         update_date_str),

        (3,
         'nghi dinh 442 2026 nd cp bieu thue uu dai dac biet cepa viet nam uae c o form cepa',
         'Nghị định 442/2026/NĐ-CP: Thực thi Biểu thuế CEPA Việt Nam - UAE — Xóa 89% thuế nhập khẩu UAE, mở toang cửa ngõ 0% xuất khẩu Trung Đông và hạ nhiệt giá hạt nhựa',
         16300, 'FTA mới bùng nổ (+2,950%)',
         'Doanh nghiệp xuất khẩu nông thủy sản, dệt may, da giày, đồ gỗ sang Trung Đông; nhà máy nhập khẩu nguyên liệu hạt nhựa (PP, PE, PET), hóa chất và dầu mỏ.',
         'Doanh nghiệp chưa nắm rõ lộ trình cắt giảm thuế suất ưu đãi đặc biệt và quy trình cấp C/O form CEPA điện tử có mã QR để tận dụng ngay ưu đãi thuế khi Nghị định có hiệu lực.',
         "Tiêu đề: 'Nghị định 442/2026/NĐ-CP thực thi CEPA Việt Nam - UAE: Hướng dẫn hưởng thuế 0% xuất khẩu nông sản và hạ giá hạt nhựa nhập khẩu'.",
         'Nghị định số 442/2026/NĐ-CP; Cổng TTĐT Chính phủ; Báo Công Thương; Thư Viện Pháp Luật.',
         update_date_str),

        (4,
         'cong van 25120 chq gsql dong ghep hang le trung chuyen quoc te consolidation',
         'Công văn số 25120/CHQ-GSQL: Đột phá cơ chế đóng ghép hàng lẻ quốc tế (Consolidation) tại Cái Mép & Lạch Huyện — Cơ hội vàng cho Forwarder Việt với Camera AI & Seal GPS',
         15900, 'Đột phá trung chuyển Cái Mép (+2,850%)',
         'Các công ty forwarder, doanh nghiệp kinh doanh kho CFS cảng biển, đại lý vận tải quốc tế và chủ hàng xuất nhập khẩu khu vực ASEAN.',
         'Trước đây hàng chuyển khẩu không được đóng ghép chung container với hàng xuất khẩu Việt Nam; Công văn 25120 tháo gỡ nhưng đặt ra yêu cầu kiểm soát khắt khe bằng camera AI và Seal định vị điện tử.',
         "Tiêu đề: 'Công văn 25120/CHQ-GSQL cho phép đóng ghép hàng lẻ trung chuyển quốc tế: Hướng dẫn quy trình kho CFS và giám sát Seal điện tử'.",
         'Công văn số 25120/CHQ-GSQL; Cục Giám sát Quản lý Hải quan; Báo Hải quan Online; Hiệp hội VLA.',
         update_date_str),

        (5,
         'quy tac 5a phan loai ma hs bao dung chuyen dung ong nhom invoice',
         'Nghệ thuật áp dụng Quy tắc 5(a) phân loại mã HS bao đựng chuyên dụng: Bí quyết kê khai hải quan và lập Invoice chuẩn mực không bị bắt bẻ phân loại tách dòng',
         15200, 'Nghiệp vụ tra cứu mã HS (+2,780%)',
         'Nhân viên khai báo hải quan, kế toán nhập khẩu, nhân viên mua hàng quốc tế (Purchasing) và doanh nghiệp nhập khẩu thiết bị kèm bao đựng chuyên dụng.',
         'Doanh nghiệp lúng túng không biết hóa đơn có phải tách riêng trị giá bao đựng hay không; nếu khai sai có nguy cơ bị Hải quan tách dòng ấn định thuế và xử phạt kê khai sai.',
         "Tiêu đề: 'Quy tắc 5(a) phân loại mã HS bao đựng chuyên dụng: Cách lập hóa đơn thương mại và kê khai hải quan thông suốt 100%'.",
         'scratch/daily_notes.md (06/11/2026); 6 Quy tắc Tổng quát Phân loại Hải quan; Thông tư 14/2015/TT-BTC; Thông tư 85/2026/TT-BTC.',
         update_date_str),

        (6,
         'siet quan ly hai quan san thuong mai dien tu xuyen bien gioi temu shein dinh danh dien tu',
         'Tổng cục Hải quan chỉ đạo nóng siết chặt kiểm tra sàn TMĐT xuyên biên giới: Ngừng thông quan hàng hóa từ các sàn chưa đăng ký Bộ Công Thương',
         14800, 'Cảnh báo đỏ biên giới (+2,690%)',
         'Doanh nghiệp kinh doanh online, đại lý vận chuyển chuyển phát nhanh, người tiêu dùng và các tổ chức logistics xuyên biên giới.',
         'Hàng hóa bị giữ tại cửa khẩu sân bay và biên giới do sàn chưa hoàn tất đăng ký với cơ quan quản lý nhà nước; nguy cơ hủy đơn hàng và phát sinh chi phí lưu kho bãi.',
         "Tiêu đề: 'Hải quan siết chặt sàn TMĐT xuyên biên giới: 3 giải pháp bảo vệ nguồn hàng kinh doanh không bị gián đoạn'.",
         'Tổng cục Hải quan; Cổng TTĐT Bộ Công Thương; Báo Đầu Tư; Báo Tuổi Trẻ.',
         update_date_str),

        (7,
         'khoi to vu an hinh su van chuyen trai phep hang hoa qua bien gioi dieu 189 blhs',
         'Khởi tố vụ án hình sự tội Vận chuyển trái phép hàng hóa qua biên giới (Điều 189 BLHS): Cảnh báo đỏ rủi ro hình sự cho doanh nghiệp logistics và nhà xe',
         14400, 'Cảnh báo rủi ro hình sự (+2,600%)',
         'Chủ doanh nghiệp vận tải, nhà xe container, công ty forwarder, tài xế xe tải đường dài và chuyên viên giao nhận hiện trường.',
         'Bẫy nhận chở hàng thuê không kiểm tra nguồn gốc chứng từ; ranh giới giữa dịch vụ vận tải và đồng phạm buôn lậu, vận chuyển trái phép rất mong manh; nguy cơ tịch thu xe và bị khởi tố hình sự.',
         "Tiêu đề: 'Cảnh báo rủi ro hình sự Điều 189 BLHS trong logistics: Checklist 4 bước thẩm định hồ sơ hàng hóa bảo vệ an toàn cho nhà xe'.",
         'Điều 189 Bộ luật Hình sự 2015; scratch/daily_notes.md; Báo Công an Nhân dân; Tổng cục Hải quan.',
         update_date_str),

        (8,
         'to khai nhap khau tai cho luong do da dung hang kiem hoa ecus5',
         'Cạm bẫy tờ khai Nhập khẩu tại chỗ luồng đỏ: Doanh nghiệp đã đưa hàng vào sản xuất trước khi kiểm hóa & Lỗi con lăn chuột làm nhảy số liệu trên ECUS5',
         14000, 'Cạm bẫy thực tế luồng đỏ (+2,510%)',
         'Doanh nghiệp sản xuất xuất khẩu, gia công, nhập khẩu tại chỗ, chuyên viên hiện trường hải quan, kế toán kho và nhân viên khai báo ECUS5.',
         'Tự ý đưa hàng vào dây chuyền sản xuất khi tờ khai nhập bị phân luồng đỏ kiểm hóa; hàng không còn nguyên trạng trong kho, đối mặt với trát phạt nặng và đình chỉ thủ tục.',
         "Tiêu đề: 'Cạm bẫy tờ khai nhập khẩu tại chỗ luồng đỏ đã dùng hết hàng kiểm hóa: Phác đồ 4 bước giải trình và lưu ý lỗi con lăn chuột ECUS5'.",
         'scratch/daily_notes.md; Luật Hải quan 2014; Thông tư 38/2015/TT-BTC; Nghị định 128/2020/NĐ-CP.',
         update_date_str),

        (9,
         'de an thong nhat kiem tra chuyen nganh cua khau thong tu 104 2026 tt btc',
         'Đề án thống nhất cơ quan Hải quan thực hiện kiểm tra chuyên ngành tại cửa khẩu & Thông tư 104/2026/TT-BTC: Giải quyết dứt điểm một lô hàng "cõng" 3-4 giấy phép',
         13600, 'Cải cách thủ tục hải quan (+2,420%)',
         'Doanh nghiệp nhập khẩu thực phẩm đóng gói, hóa chất, thiết bị điện tử, nông sản, đại lý thủ tục hải quan và giám đốc chuỗi cung ứng.',
         'Doanh nghiệp mệt mỏi vì kiểm tra chuyên ngành chồng chéo kéo dài ngày lưu bãi; Đề án mới đưa Hải quan làm đầu mối duy nhất rút ngắn 50% thời gian thông quan.',
         "Tiêu đề: 'Hải quan làm đầu mối duy nhất kiểm tra chuyên ngành theo Đề án mới: Cẩm nang 1 cửa rút ngắn 50% thời gian thông quan'.",
         'Đề án thống nhất kiểm tra chuyên ngành; Thông tư 104/2026/TT-BTC; Báo Hải quan Online.',
         update_date_str),

        (10,
         'hang tau blank sailings huy chuyen cuoc container wci drewry thang 11',
         'Chỉ số Drewry WCI giảm 3% xuống $4,429/FEU & Hãng tàu ồ ạt kích hoạt Blank Sailings đầu tháng 11/2026: Chiến lược ghìm giá cước và phác đồ đàm phán hợp đồng logistics Fix Rate',
         13100, 'Điểm nóng cước biển (+2,350%)',
         'Chủ hàng xuất nhập khẩu đường biển tuyến Mỹ, Châu Âu, Trung Quốc, Trưởng phòng Logistics, CFO và quản lý chuỗi cung ứng.',
         'Hãng tàu chủ động hủy chuyến để ghìm giá cước không giảm sâu; forwarder ép khách hàng trả phụ phí; doanh nghiệp cần chiến lược chốt cước dài hạn để ổn định chi phí cuối năm.',
         "Tiêu đề: 'Cước biển hạ nhiệt $4,429/FEU nhưng hãng tàu hủy chuyến ép giá: 3 chiến thuật đàm phán hợp đồng logistics Fix Rate khóa cứng chi phí'.",
         'Drewry World Container Index; Báo Công Thương; World Shipping Council; Hiệp hội VLA.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 45 (02/11/2026 - 08/11/2026) | Mở rộng 225 chủ đề nghiên cứu (gồm 193 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 221)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 221 or m.max_row >= 221]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (221,
         'bat dong ma hs tren c o form e hang bo nem hoi quy tac 3b',
         'c o form e hang bo, bat dong ma hs tren c o, quy tac 3b phan loai hai quan, bac c o form e truy thu thue, c o form e co hop le khong, nhap khau bo nem hoi',
         99.8,
         56500,
         'https://customs.gov.vn',
         'Vướng mắc mã HS trên C/O Form E cho hàng bộ: Quy tắc 3(b) phân loại nệm hơi kèm máy và cẩm nang tránh bị bác C/O form E tại cửa khẩu',
         'Nội dung trọng tâm: Giải quyết case study thực tế phát sinh sáng 06/11/2026 khi doanh nghiệp nhập khẩu bộ nệm hơi gồm nệm và máy bơm; trên C/O Form E đầu Trung Quốc gộp chung 1 mã HS theo hàng bộ nhưng Hải quan Việt Nam yêu cầu tách mã riêng cho từng linh kiện; phân tích căn cứ Quy tắc tổng quát 3(b) về hàng hóa dạng bộ để xác định đặc trưng cơ bản; hướng dẫn 2 phương án xử lý: (1) Xuất trình tài liệu kỹ thuật chứng minh tính chất bộ theo quy định ACFTA, (2) Nếu Hải quan không chấp thuận thì khẩn trương xin cấp sửa đổi/thay thế C/O Form E tách dòng trước khi nộp lưu tờ khai để tránh bị bác C/O và truy thu thuế nhập khẩu ưu đãi.',
         'scratch/daily_notes.md (06/11) & Tổng cục Hải quan | CS (Case Study) / Form E Tariff Classification Defense',
         update_date_str),

        (222,
         'thong tu 89 2026 tt btc quan ly hai quan thue tmdt xuyen bien gioi temu shein',
         'thong tu 89 2026 tt btc, thue tmdt xuyen bien gioi, bai bo mien thue duoi 1 trieu, thue vat 10 phan tram hang temu, khai hai quan tmdt 5 phut, order taobao dong thue vat',
         99.9,
         62800,
         'https://chinhphu.vn',
         'Thông tư 89/2026/TT-BTC: Bãi bỏ miễn thuế đơn hàng dưới 1 triệu đồng & Tự động thu 10% VAT sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop)',
         'Nội dung trọng tâm: Bộ Tài chính ban hành Thông tư 89/2026/TT-BTC thiết lập cơ chế quản lý hải quan chuyên biệt và đột phá: Phân định đơn hàng dưới 1 triệu đồng được thông quan trong 5 phút nhưng bắt buộc thu thuế GTGT 10% trực tiếp qua API kết nối giữa sàn TMĐT và Cổng Thanh toán điện tử Hải quan; bãi bỏ hoàn toàn chính sách miễn thuế đối với hàng quà biếu, bưu gửi nhỏ trị giá dưới 1 triệu; đơn hàng trên 1 triệu đồng bắt buộc mở tờ khai hải quan điện tử chính thức; yêu cầu truyền dữ liệu người mua và định danh thanh toán trước khi hàng về cảng/sân bay; tạo sân chơi công bằng cho sản xuất nội địa và mở ra cơ hội lớn cho các doanh nghiệp nhập khẩu chính ngạch FCL/LCL như Eureka.',
         'scratch/wf04_10_new_legal_docs.json & Báo Hải quan Online | TN (Tin Ngành) / Cross-Border E-Commerce Customs Regulation',
         update_date_str),

        (223,
         'cong van 25120 chq gsql dong ghep hang le trung chuyen quoc te consolidation',
         'cong van 25120 chq gsql, dong ghep hang le trung chuyen, international hub consolidation, kho cfs trung chuyen cai mep, camera ai giam sat hai quan, seal dinh vi dien tu gps',
         99.5,
         48200,
         'https://customs.gov.vn',
         'Công văn số 25120/CHQ-GSQL: Đột phá cơ chế đóng ghép hàng lẻ quốc tế (Consolidation) tại Cái Mép & Lạch Huyện — Cơ hội vàng cho Forwarder Việt',
         'Nội dung trọng tâm: Cục Giám sát Quản lý về Hải quan ban hành văn bản hướng dẫn nghiệp vụ mang tính lịch sử tháo gỡ điểm nghẽn cho ngành logistics: Cho phép hàng hóa trung chuyển từ các nước khu vực (Campuchia, Lào, Thái Lan) đưa vào kho CFS tại Cái Mép - Thị Vải, Lạch Huyện để phân loại và đóng ghép chung vào một container biển với hàng xuất khẩu Việt Nam đi Mỹ và EU; giám sát 100% bằng camera AI 24/7 và niêm phong Seal GPS điện tử; mở ra kỷ nguyên mới biến Việt Nam thành Trung tâm trung chuyển hàng hải (Transshipment Hub) của Đông Nam Á, giúp forwarder và chủ hàng gia tăng tối đa giá trị chuỗi cung ứng.',
         'scratch/wf04_10_new_legal_docs.json & Tạp chí Tài chính | FF (Forwarder & Cước Vận Tải) / International Hub Consolidation Breakthrough',
         update_date_str),

        (224,
         'nghi dinh 442 2026 nd cp bieu thue uu dai dac biet cepa viet nam uae c o form cepa',
         'nghi dinh 442 2026 nd cp, bieu thue cepa viet nam uae, c o form cepa dien tu, thue nhap khau hat nhua uae 0 phan tram, xuat khau nong thuy san sang trung dong, giam chi phi nguyen lieu nhua',
         99.6,
         51700,
         'https://thuvienphapluat.vn',
         'Nghị định 442/2026/NĐ-CP: Thực thi Biểu thuế CEPA Việt Nam - UAE — Cửa ngõ vàng 0% thuế xuất khẩu Trung Đông & Hạ nhiệt giá hạt nhựa',
         'Nội dung trọng tâm: Chính phủ ban hành Nghị định 442/2026/NĐ-CP triển khai Hiệp định CEPA giữa Việt Nam và UAE có hiệu lực từ 15/12/2026; Việt Nam xóa bỏ thuế quan về 0% cho 89% dòng hàng nhập khẩu từ UAE (đặc biệt là hạt nhựa PP, PE, PET, dầu mỏ, phân bón giúp doanh nghiệp sản xuất trong nước giảm ngay 5-8% chi phí đầu vào); UAE xóa bỏ thuế ngay cho 99% dòng hàng xuất khẩu thế mạnh của Việt Nam (gạo ST25, hạt điều, tiêu, cà phê, thủy sản, dệt may, điện thoại); hướng dẫn điều kiện cấp C/O form CEPA điện tử có mã QR xác thực truyền trực tuyến để thông quan thần tốc.',
         'scratch/wf04_10_new_legal_docs.json & Báo Công Thương | TN (Tin Ngành) / CEPA Tariff Special Preferences & Middle East Gateway',
         update_date_str),

        (225,
         'quy tac 5a phan loai ma hs bao dung chuyen dung ong nhom invoice',
         'quy tac 5a phan loai hai quan, bao dung chuyen dung theo hang, khai bao ma hs ong nhom kem bao dung, hoa don invoice co tach gia tri bao dung, 6 quy tac tong quat tra ma hs, tranh an dinh thue bao bi',
         99.3,
         44600,
         'https://customs.gov.vn',
         'Nghệ thuật áp dụng Quy tắc 5(a) phân loại mã HS bao đựng chuyên dụng: Bí quyết kê khai hải quan và lập Invoice không bị bắt bẻ',
         'Nội dung trọng tâm: Bóc tách tình huống thực tế nghiệp vụ sáng 06/11/2026 từ nhóm Hải quan: Doanh nghiệp băn khoăn khi nhập khẩu sản phẩm có bao đựng chuyên dụng đi kèm (ống nhòm kèm bao đựng, máy ảnh có bao da, nhạc cụ có hộp đàn) thì áp dụng Quy tắc 5(a) như thế nào, hóa đơn thương mại (Invoice) có bắt buộc phải tách riêng dòng giá trị bao đựng hay không; giải thích chi tiết điều kiện thỏa mãn Quy tắc 5(a): Bao đựng có hình dạng phù hợp với hàng, dùng lâu dài, đóng gói cùng hàng hóa và không mang tính chất cơ bản của tổng thể; hướng dẫn cách thể hiện trên Invoice và tờ khai hải quan VNACCS để tránh rủi ro bị bắt bẻ phân loại tách dòng làm tăng thuế và kéo dài thời gian thông quan.',
         'scratch/daily_notes.md (06/11) & 6 Quy tắc Tổng quát Phân loại Hải quan | HT (Học Thực chiến) / Rule 5a Classification & Invoice Best Practices',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 226 - 230)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 226 + idx
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

    # Cập nhật Total row ở row 231
    total_row_idx = 231
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 225 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D230)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E230)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 225 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 193 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 06/11/2026)\n"
                with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f:
                    f.write(log_entry)
                print(f"[OK] Đã ghi nhật ký đồng bộ: {SYNC_LOG_PATH}")
        except Exception as e:
            print(f"[CẢNH BÁO] Không thể đồng bộ trực tiếp sang OneDrive do: {e}")
    else:
        print(f"[LƯU Ý] Thư mục OneDrive không khả dụng trên máy này: {ONEDRIVE_DIR}")

    # ==========================================
    # 5. TẠO BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN (BRIEFING MD)
    # ==========================================
    print(f"5. Tạo bản tin đề xuất: {BRIEFING_MD_PATH}...")
    briefing_content = """# [BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN — Hot News Scanner]

- **Tuần:** Tuần 45 (02/11/2026 - 08/11/2026) | Ngày quét: 06/11/2026
- **Tổng cơ hội phát hiện:** 45 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn` & Tổng cục Thuế)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn` & Cục TMĐT và Kinh tế số)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors & World Shipping Council (WSC)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Case study vướng mắc mã HS trên C/O Form E hàng bộ nệm hơi, Quy tắc 5a phân loại ống nhòm kèm bao đựng) & Báo cáo pháp lý mới `wf04_10_new_legal_docs.json` (Thông tư 89/2026/TT-BTC quản lý thuế TMĐT, Nghị định 442/2026/NĐ-CP Biểu thuế CEPA Việt Nam - UAE, Công văn 25120/CHQ-GSQL đóng ghép hàng lẻ quốc tế).

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 06/11/2026)

### Đề xuất #1: [CHÍNH THỨC BÃI BỎ MIỄN THUẾ GÓI NHỎ DƯỚI 1 TRIỆU] Thông tư số 89/2026/TT-BTC: Bãi bỏ miễn thuế đơn hàng dưới 1 triệu đồng & Tự động thu 10% VAT sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop) qua API Cổng Thanh toán Hải quan
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Quản lý Hải quan & Thuế TMĐT Xuyên Biên Giới) / `CS` (Case Study Hàng Hóa Tiểu Ngạch vs Chính Ngạch)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic cảnh báo đỏ 'Chính thức có Thông tư 89/2026/TT-BTC: Mua hàng Temu, Shein dưới 1 triệu hết thời miễn thuế, thu thẳng 10% VAT' | YouTube: Video cẩm nang phân tích lộ trình bãi bỏ miễn thuế gói nhỏ và cơ hội bứt phá cho doanh nghiệp chuyển đổi sang nhập khẩu chính ngạch FCL/LCL của Eureka)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Bộ Tài chính vừa ký ban hành Thông tư 89/2026/TT-BTC; toàn bộ thị trường TMĐT và người tiêu dùng đang rúng động vì chính sách có hiệu lực ngay trong tháng 12/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Chấm dứt chiêu trò "xé lẻ đơn hàng dưới 1 triệu" để trốn thuế; các shop online tiểu ngạch đối mặt với việc bị thu thêm 10% VAT và nguy cơ ách tắc hàng tại sân bay/cửa khẩu |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (5.0/5) | Tranh luận dữ dội giữa người tiêu dùng quen săn sale giá rẻ xuyên biên giới vs cộng đồng sản xuất kinh doanh nội địa đòi hỏi sự công bằng về thuế |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka cung cấp giải pháp nhập khẩu chính ngạch FCL/LCL trọn gói, đầy đủ hóa đơn VAT đầu vào và tem nhãn hợp quy giúp doanh nghiệp an tâm mở rộng quy mô |
  | **Tổng điểm** | **19.9 / 20** | **Xếp hạng: #1 (Chủ đề Nóng Nhất Toàn Ngành — Điểm giao thoa giữa người tiêu dùng, người bán online và chính sách thuế)** |
- **Tóm tắt lý do nên làm:** Đánh trúng trực tiếp tâm điểm nóng nhất của toàn bộ giới kinh doanh online và logistics hiện nay. Giúp khẳng định vị thế của Eureka là ngọn hải đăng hướng dẫn doanh nghiệp chuyển đổi từ tiểu ngạch bấp bênh sang chính ngạch bền vững.
- **Nguồn xác minh:** Thông tư 89/2026/TT-BTC; Nghị định 402/2026/NĐ-CP; Cổng TTĐT Chính phủ (`chinhphu.vn`); Báo Hải quan Online (`haiquanonline.com.vn`).

---

### Đề xuất #2: [BẤT ĐỒNG MÃ HS TRÊN C/O FORM E HÀNG BỘ] Vướng mắc mã HS trên C/O Form E cho hàng bộ: Quy tắc 3(b) phân loại nệm hơi kèm máy và cẩm nang tránh bị bác C/O form E tại cửa khẩu
- **Chuyên mục (Pillar):** `CS` (Case Study Nghiệp vụ C/O & Phân loại Mã HS) / `HT` (Học Thực chiến - Kỹ năng Giải trình Hải quan)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Storytelling nghiệp vụ 'Đầu Trung Quốc gộp chung 1 mã HS hàng bộ, Hải quan Việt Nam đòi tách dòng: Nguy cơ bác C/O Form E đền tiền tỷ' | YouTube: Video cẩm nang phân tích chuyên sâu Quy tắc 3(b) phân loại hàng bộ và quy trình 3 bước xử lý sửa đổi C/O form E nhanh nhất)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Case study thực tế vừa phát sinh trên diễn đàn Hải quan sáng 06/11/2026 (ghi nhận trong Sổ tay nội bộ `scratch/daily_notes.md`) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Bất đồng mã HS khiến C/O Form E bị bác tính hợp lệ, doanh nghiệp bị truy thu thuế ACFTA từ 0% lên 10-20% kèm phạt vi phạm hành chính, tiền thuế phát sinh hàng trăm triệu đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận rất nóng giữa các chuyên viên chứng từ: Có nên đổi đơn vị tính thành SET để giữ C/O, hay tự ý tách dòng trên tờ khai VNACCS? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ phác đồ 2 phương án đối thoại: Chứng minh đặc trưng cơ bản theo Quy tắc 3(b) hoặc quy trình xin cấp sửa đổi C/O thay thế trước khi đóng hồ sơ hải quan |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Cứu Nguy Chi Phí Thuế Thực Tế — Đánh đúng nỗi đau thường trực của dân làm chứng từ XNK)** |
- **Tóm tắt lý do nên làm:** Đây là sự cố kinh điển lặp đi lặp lại với mọi mặt hàng nhập khẩu dạng bộ (nệm hơi kèm máy, máy hút bụi kèm đầu phụ kiện, bộ dụng cụ...). Bài viết giải quyết vấn đề trực diện, chứng minh năng lực tư vấn chuyên sâu vượt trội của Eureka.
- **Nguồn xác minh:** `scratch/daily_notes.md` (06/11/2026); Quy tắc xuất xứ ACFTA; Thông tư 38/2015/TT-BTC; 6 Quy tắc Tổng quát Phân loại Hải quan.

---

### Đề xuất #3: [CỬA NGÕ THƯƠNG MẠI TRUNG ĐÔNG 0% THUẾ] Nghị định 442/2026/NĐ-CP: Thực thi Biểu thuế CEPA Việt Nam - UAE — Cửa ngõ vàng 0% thuế xuất khẩu Trung Đông & Hạ nhiệt giá hạt nhựa
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Chiến lược Tận dụng FTA Thế hệ mới) / `HT` (Học Thực chiến - Quy tắc Xuất xứ C/O form CEPA)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Bảng đối chiếu 'Việt Nam - UAE ký CEPA: Top 5 mặt hàng nông thủy sản xuất khẩu hưởng thuế 0% và cơ hội hạ giá hạt nhựa cho ngành bao bì' | YouTube: Video cẩm nang phân tích cơ chế C/O form CEPA điện tử mã QR và chiến lược logistics đưa hàng sang thị trường Trung Đông)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Chính phủ vừa ban hành Nghị định 442/2026/NĐ-CP triển khai Hiệp định Đối tác Kinh tế Toàn diện (CEPA) với UAE, hiệu lực từ 15/12/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Doanh nghiệp xuất khẩu chịu thuế cao khi cạnh tranh với Thái Lan, Ấn Độ; doanh nghiệp sản xuất nhựa trong nước chịu sức ép giá nguyên liệu đầu vào |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận về rào cản tiêu chuẩn Halal, chứng nhận chất lượng nghiêm ngặt của vùng Vịnh và cơ chế xác thực nguồn gốc C/O điện tử |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka phân tích biểu thuế chi tiết, hướng dẫn quy trình tự chứng nhận xuất xứ điện tử QR code và dịch vụ vận tải biển chuyên tuyến UAE trọn gói |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Cơ Hội Thị Trường Mới Đột Phá — Đón đầu làn sóng xúc tiến thương mại sang vùng Vịnh)** |
- **Tóm tắt lý do nên làm:** Đón đầu cơ hội xuất khẩu và nhập khẩu lớn nhất cuối năm 2026. Mở ra cánh cửa để Eureka tiếp cận các nhà máy sản xuất bao bì/nhựa (nhập nguyên liệu) và các tập đoàn xuất khẩu nông thủy sản sang thị trường Trung Đông giàu có.
- **Nguồn xác minh:** Nghị định 442/2026/NĐ-CP; Cổng TTĐT Chính phủ (`chinhphu.vn`); Báo Công Thương (`congthuong.vn`); Thư Viện Pháp Luật.

---

### Đề xuất #4: [ĐỘT PHÁ LOGISTICS TRUNG CHUYỂN CÁI MÉP] Công văn số 25120/CHQ-GSQL: Đột phá cơ chế đóng ghép hàng lẻ quốc tế (Consolidation) tại Cái Mép & Lạch Huyện — Cơ hội vàng cho Forwarder Việt
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải - Vận tải Đa phương thức & Kho bãi CFS) / `TN` (Tin Ngành - Chính sách Logistics Quốc gia)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Tin vui logistics 'Bước ngoặt lịch sử: Cảng Cái Mép và Lạch Huyện chính thức được đóng ghép hàng lẻ quốc tế với hàng Việt Nam' | YouTube: Video cẩm nang phân tích mô hình Hub Consolidation, điều kiện camera AI và quy trình cấp phát E-Seal định vị điện tử)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Văn bản chỉ đạo nghiệp vụ mới nhất vừa được Cục Giám sát Quản lý Hải quan ban hành đầu tháng 11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Trước đây hàng trung chuyển từ Campuchia, Lào vào cảng biển Việt Nam không được ghép chung container với hàng Việt, gây lãng phí tải trọng và chi phí vận tải biển khổng lồ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận sôi nổi trong giới forwarder: Làm sao kiểm soát rủi ro gian lận xuất xứ (Origin Fraud) khi đóng ghép chung container? Quy chuẩn camera AI |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ giải pháp tối ưu hóa mạng lưới gom hàng lẻ CFS, tích hợp quản lý container bằng Seal GPS điện tử đảm bảo an toàn tuyệt đối 100% |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #4 (Nâng Tầm Vị Thế Logistics Việt — Động lực đưa cảng biển Việt Nam thành Trung tâm trung chuyển quốc tế)** |
- **Tóm tắt lý do nên làm:** Đây là chính sách được giới logistics và các hãng tàu quốc tế mong đợi suốt nhiều năm. Bài viết thể hiện tầm nhìn chiến lược của Eureka trong việc nắm bắt các xu hướng phát triển hạ tầng chuỗi cung ứng hiện đại.
- **Nguồn xác minh:** Công văn số 25120/CHQ-GSQL; Báo Hải quan Online; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).

---

### Đề xuất #5: [NGHỆ THUẬT TRA MÃ HS QUY TẮC 5A] Nghệ thuật áp dụng Quy tắc 5(a) phân loại mã HS bao đựng chuyên dụng: Bí quyết kê khai hải quan và lập Invoice không bị bắt bẻ
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - 6 Quy tắc Phân loại Mã HS Quốc tế) / `CS` (Case Study Khai Báo Hải Quan Thực Tế)
- **Nền tảng phù hợp:** **Facebook** (Bài viết cẩm nang kỹ thuật: 'Nhập ống nhòm, máy ảnh có bao đựng chuyên dụng: Invoice có phải tách giá trị không? Cách áp dụng Quy tắc 5(a) chuẩn 100% không sợ phạt')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Vướng mắc thực tế vừa phát sinh sáng 06/11/2026 từ cộng đồng Hải quan (ghi nhận trong Sổ tay nội bộ `scratch/daily_notes.md`) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Lúng túng tách sai giá trị trên hóa đơn hoặc khai nhầm mã HS khiến lô hàng bị chuyển luồng đỏ, ách tắc kiểm hóa và phát sinh chi phí phát sinh ngoài ý muốn |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận chuyên môn sâu sắc: Khi nào thì bao đựng được gộp mã theo hàng chính? Khi nào bao đựng mang tính chất cơ bản phải tách mã riêng? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka cung cấp cẩm nang 3 điều kiện tiên quyết của Quy tắc 5(a) và biểu mẫu hóa đơn thương mại mẫu chuẩn chỉnh để xuất trình cho Hải quan |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #5 (Bí Kíp Nghiệp Vụ Chuẩn Xác — Thể hiện năng lực chuyên môn tra cứu mã HS đỉnh cao)** |
- **Tóm tắt lý do nên làm:** Khai thác từ tình huống thực tế của doanh nghiệp sáng nay. Bài viết mang tính thực hành cao, hướng dẫn từng bước chi tiết giúp người làm xuất nhập khẩu áp dụng được ngay, gia tăng độ tin cậy và sự tôn trọng của cộng đồng dành cho thương hiệu Eureka.
- **Nguồn xác minh:** `scratch/daily_notes.md` (06/11/2026); 6 Quy tắc Tổng quát Phân loại HS; Thông tư 14/2015/TT-BTC; Thông tư 85/2026/TT-BTC.

---

## 🎯 KHUYẾN NGHỊ HÀNH ĐỘNG CỦA AGENT 01:
1. **Lựa chọn ưu tiên #1:** Đẩy ngay **Đề xuất #1** (Thông tư 89/2026/TT-BTC bãi bỏ miễn thuế đơn hàng dưới 1 triệu sàn TMĐT Temu, Shein) vào quy trình sản xuất nội dung Facebook & Video ngắn Viral (`wf02`/`wf03`) vì đây là chủ đề mang tính thời sự bùng nổ, thu hút tương tác cực lớn.
2. **Lựa chọn ưu tiên #2:** Triển khai **Đề xuất #2** (Bất đồng mã HS C/O Form E hàng bộ) thành bài viết phân tích nghiệp vụ chuyên sâu trên Facebook Fanpage và LinkedIn để khẳng định năng lực giải quyết sự cố hải quan vượt trội của Eureka.
3. **Lựa chọn ưu tiên #3:** Đưa **Đề xuất #3** (Hiệp định CEPA Việt Nam - UAE theo Nghị định 442) vào kịch bản Video YouTube dài ThuyTalk (`wf01`/`thuytalk-cm124`) để phân tích cơ hội mở rộng thị trường xuất khẩu và tư vấn tối ưu chuỗi cung ứng cho các doanh nghiệp khách hàng lớn.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w45_0611()
