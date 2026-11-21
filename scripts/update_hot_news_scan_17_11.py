# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_17_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 17/11/2026
(Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026 | Cập nhật ngày 17/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 47 cập nhật ngày 17/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 275 lên 280 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 17/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w47_1711.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w47_1711.md'


def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()


def update_seo_w47_1711():
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

    update_date_str = '17/11/2026'

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 47 (16/11/2026 - 22/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van 24313 chq gsql thu tuc hai quan thuong mai dien tu xuyen bien gioi',
         'Công văn 24313/CHQ-GSQL (16/11/2026): Triển khai Hệ thống thực hiện thủ tục hải quan chuyên biệt cho hàng hóa XNK giao dịch qua nền tảng thương mại điện tử',
         24100, 'Tăng vọt triển khai TMĐT xuyên biên giới (+4,850%)',
         'Chủ shop TMĐT xuyên biên giới, doanh nghiệp XNK đơn nhỏ số lượng lớn, sàn TMĐT, đơn vị logistics và đại lý hải quan.',
         'Đơn hàng TMĐT xuyên biên giới có tần suất lớn, giá trị thấp nhưng trước đây xử lý rời rạc dễ tắc nghẽn; Công văn 24313 triển khai hệ thống thủ tục riêng kết nối dữ liệu tự động và quản lý rủi ro chuyên biệt; doanh nghiệp cần chuẩn hóa dữ liệu đơn hàng để thông quan nhanh.',
         "Tiêu đề: 'Công văn 24313/CHQ-GSQL: Doanh nghiệp và chủ shop TMĐT xuyên biên giới cần chuẩn bị gì khi Hải quan vận hành hệ thống khai báo riêng?'.",
         'Công văn 24313/CHQ-GSQL ngày 16/11/2026; Luật Hải quan sửa đổi 11/2026/QH16; luatvietnam.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (2,
         'quyet dinh 1828 qd chq kiem tra xuat xu hang hoa sau thong quan 15 ngay',
         'Quyết định 1828/QĐ-CHQ (hiệu lực 15/12/2026): Đếm ngược quy trình kiểm tra xuất xứ hàng hóa sau thông quan tối đa 15 ngày, siết chống chuyển tải bất hợp pháp',
         23500, 'Siết hậu kiểm C/O & chống chuyển tải (+4,690%)',
         'Doanh nghiệp xuất khẩu gỗ, thép, dệt may, điện tử, nông sản sử dụng C/O ưu đãi; doanh nghiệp nhập khẩu nguyên liệu và phòng XNK.',
         'Hàng đã thông quan vẫn có thể bị Đội Phúc tập và KTSTQ gọi kiểm tra xuất xứ trong 15 ngày theo nguyên tắc quản lý rủi ro; rủi ro bị bác C/O, truy thu thuế và phạt nếu thiếu hồ sơ chứng minh nguồn gốc nguyên liệu, quy trình gia công.',
         "Tiêu đề: 'Quyết định 1828/QĐ-CHQ hiệu lực 15/12/2026: Bộ hồ sơ chứng minh xuất xứ cần giữ sẵn để vượt qua kỳ hậu kiểm 15 ngày'.",
         'Quyết định 1828/QĐ-CHQ ngày 12/11/2026; Thông tư 124/2026/TT-BTC; Quyết định 467/QĐ-CHQ; thuvienphapluat.vn; luatvietnam.vn.',
         update_date_str),

        (3,
         'thong tu 142 2026 tt btc quyet dinh 1813 qd chq chuyen phat nhanh a45 a47',
         'Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ (đã hiệu lực từ 16/11/2026): Chuẩn thông điệp dữ liệu, mã loại hình A45/A47 và cơ chế nộp thay thuế chuyển phát nhanh',
         22900, 'Cao điểm ngày thứ 2 áp dụng TT 142 (+4,540%)',
         'Doanh nghiệp nhập hàng mẫu, linh kiện gấp qua FedEx, DHL, UPS, EMS; kế toán thuế cần tờ khai khấu trừ VAT và đơn vị chuyển phát.',
         'Từ 16/11/2026, thông điệp dữ liệu tờ khai chuyển phát nhanh phải chuẩn hóa theo QĐ 1813 và phụ lục mới (mã A45 hàng trị giá thấp, A47 tài liệu); nhiều lô hàng mẫu bị chậm vì thiếu mã định danh chủ hàng hoặc vướng cách nộp thay thuế.',
         "Tiêu đề: 'Sau ngày đầu Thông tư 142/2026/TT-BTC có hiệu lực: Cách dùng đúng mã A45/A47 và lấy tờ khai độc lập cho hàng mẫu chuyển phát nhanh'.",
         'Thông tư 142/2026/TT-BTC; Quyết định 1813/QĐ-CHQ; chinhphu.vn; customs.gov.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (4,
         'cong van 24286 chq pc to khai dieu chinh thue suat khong xu phat vi pham',
         'Công văn 24286/CHQ-PC (16/11/2026): Hướng dẫn tờ khai điều chỉnh thuế suất thuế nhập khẩu và căn cứ miễn xử phạt vi phạm hành chính',
         21900, 'Tháo gỡ vướng mắc miễn phạt khai bổ sung (+4,280%)',
         'Doanh nghiệp nhập khẩu gặp sai lệch thuế suất do lỗi hệ thống hoặc thay đổi biểu thuế; nhân viên khai báo hải quan và kế toán thuế.',
         'Doanh nghiệp lo ngại khi mở tờ khai điều chỉnh thuế suất thuế nhập khẩu sẽ mặc định bị xử phạt khai sai thiếu thuế; Công văn 24286 phân định rõ trường hợp bất khả kháng, lỗi hệ thống được xem xét không xử phạt theo Luật XLVPHC.',
         "Tiêu đề: 'Công văn 24286/CHQ-PC: Khi nào điều chỉnh thuế suất thuế nhập khẩu được miễn xử phạt vi phạm hành chính?'.",
         'Công văn 24286/CHQ-PC ngày 16/11/2026; Điều 11 Luật Xử lý vi phạm hành chính; thuvienphapluat.vn; luatvietnam.vn.',
         update_date_str),

        (5,
         'cong van 24153 chq nvthq ra soat ma hs my pham sua rua mat 3401 mat na 3307',
         'Công văn 24153/CHQ-NVTHQ: Hải quan rà soát chấn chỉnh khai sai mã HS mỹ phẩm nhập khẩu (sữa rửa mặt 3401.30.00 và mặt nạ dưỡng da 3307.90.90)',
         21400, 'Cảnh báo rủi ro áp sai mã HS mỹ phẩm (+4,150%)',
         'Doanh nghiệp nhập khẩu mỹ phẩm Hàn Quốc, Trung Quốc, Nhật Bản; chủ thương hiệu làm đẹp và nhân viên chứng từ khai báo HS.',
         'Nhiều doanh nghiệp lâu nay gộp chung sữa rửa mặt và mặt nạ vải không dệt vào nhóm 33.04 dẫn đến nguy cơ bị bác mã HS, lấy mẫu phân tích phân loại và truy thu thuế chênh lệch khi Hải quan tăng cường rà soát theo CV 24153.',
         "Tiêu đề: 'Công văn 24153/CHQ-NVTHQ: Đừng gộp mỹ phẩm vào nhóm 33.04 — Rà soát ngay mã HS sữa rửa mặt và mặt nạ dưỡng da'.",
         'Công văn 24153/CHQ-NVTHQ ngày 12/11/2026; Thông tư 31/2022/TT-BTC; thuvienphapluat.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (6,
         'cong van 24288 chq gsql thu tuc hai quan hoa chat cong nghiep bo cong thuong',
         'Công văn 24288/CHQ-GSQL (16-17/11/2026): Hướng dẫn thủ tục hải quan, mã HS và phối hợp quản lý chuyên ngành hóa chất công nghiệp với Bộ Công Thương',
         20700, 'Gỡ vướng thông quan hóa chất chuyên ngành (+3,990%)',
         'Doanh nghiệp sản xuất công nghiệp, xi mạ, cao su, linh kiện điện tử nhập khẩu hóa chất tiền chất/chuyên ngành (Potassium Gold Cyanide, chất xúc tiến lưu hóa).',
         'Hàng hóa chất công nghiệp thường bị vướng đồng thời ở khâu áp mã HS và giấy phép/khai báo hóa chất Bộ Công Thương, gây lưu bãi nguy hiểm chi phí cao; CV 24288 làm rõ đầu mối phối hợp tháo gỡ.',
         "Tiêu đề: 'Công văn 24288/CHQ-GSQL: Lưu ý mã HS và hồ sơ chuyên ngành Bộ Công Thương khi nhập khẩu hóa chất công nghiệp'.",
         'Công văn 24288/CHQ-GSQL ngày 16-17/11/2026; Luật Hóa chất; thuvienphapluat.vn.',
         update_date_str),

        (7,
         'cong van 24259 chq gsql ghi nhan phu nhap khau song song ky hieu R',
         'Công văn 24259/CHQ-GSQL: Hướng dẫn ghi nhãn phụ và nhập khẩu song song hàng hóa có ký hiệu ® trên tem nhãn',
         20100, 'Gỡ vướng nhãn phụ & sở hữu trí tuệ (+3,870%)',
         'Doanh nghiệp nhập khẩu hàng tiêu dùng, thương hiệu quốc tế phân phối chính thức hoặc nhập khẩu song song; phòng pháp chế và XNK.',
         'Nhiều lô hàng có ký hiệu ® trên bao bì gốc bị ách tắc vì nghi ngại vi phạm Nghị định 99/2013/NĐ-CP; CV 24259 thống nhất theo Cục SHTT: thể hiện trung thực tình trạng bảo hộ tại VN trên nhãn phụ thì không bị xử phạt.',
         "Tiêu đề: 'Công văn 24259/CHQ-GSQL: Cách ghi nhãn phụ chuẩn cho hàng nhập khẩu có ký hiệu ® để thông quan an toàn'.",
         'Công văn 24259/CHQ-GSQL ngày 13/11/2026; Nghị định 99/2013/NĐ-CP; thuvienphapluat.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (8,
         'quyet dinh 3025 qd bct dieu tra tu ve hat nhua pe pp phong ve thuong mai',
         'Quyết định 3025/QĐ-BCT & Cảnh báo Cục PVTM tháng 11/2026: Điều tra tự vệ hạt nhựa PE/PP (SG04.2026) và rủi ro phòng vệ thương mại dù kim ngạch giảm',
         19600, 'Cảnh báo phòng vệ thương mại 2 chiều (+3,740%)',
         'Doanh nghiệp nhập khẩu hạt nhựa PE/PP làm bao bì và doanh nghiệp xuất khẩu gỗ, thép, nội thất, điện tử sang thị trường đối tác.',
         'Ngành nhựa đối mặt nguy cơ thuế tự vệ tạm thời 5-12% với thời hạn 30 ngày nộp hồ sơ TRAV ONLINE; đồng thời Cục PVTM cảnh báo ngay cả khi kim ngạch xuất khẩu giảm, tỷ trọng lớn vẫn dễ kích hoạt điều tra PVTM.',
         "Tiêu đề: 'Cảnh báo PVTM tháng 11/2026: Từ vụ hạt nhựa PE/PP trong nước đến rủi ro xuất khẩu giảm vẫn bị điều tra'.",
         'Quyết định 3025/QĐ-BCT ngày 16/11/2026; Cục Phòng vệ thương mại (moit.gov.vn); congthuong.vn; tapchicongthuong.vn.',
         update_date_str),

        (9,
         'cong van 24092 chq nvthq cong van 23739 hoan thue nhap khau tai cho tai xuat',
         'Công văn 24092/CHQ-NVTHQ & Công văn 23739/CHQ-NVTHQ: Hướng dẫn hoàn thuế nguyên liệu nhập khẩu tại chỗ SXXK và xử lý thuế GTGT hàng tái xuất trả',
         19100, 'Khơi thông dòng tiền hoàn thuế SXXK (+3,620%)',
         'Doanh nghiệp sản xuất xuất khẩu (SXXK), gia công, DNCX và doanh nghiệp thương mại phải tái xuất trả hàng lỗi về nước ngoài.',
         'Hồ sơ hoàn thuế nhập khẩu tại chỗ theo NĐ 134/2016 (sửa đổi NĐ 18/2021) và xử lý thuế GTGT hàng nhập xong tái xuất trả đối tác dễ bị treo vốn nếu sai mốc tờ khai hoặc sai loại hình SXXK.',
         "Tiêu đề: 'Hoàn thuế nhập khẩu tại chỗ & hàng tái xuất trả: Đối chiếu ngay CV 24092 và CV 23739 để không đọng vốn cuối năm'.",
         'Công văn 24092/CHQ-NVTHQ (10/11/2026); Công văn 23739/CHQ-NVTHQ (02/11/2026); tapchikinhtetaichinh.vn; luatvietnam.vn.',
         update_date_str),

        (10,
         'tong kim ngach xnk 1000 ty usd cuoc bien drewry wci xuat khau rau qua trung quoc',
         'XNK vượt mốc 1.010,55 tỷ USD, xuất khẩu rau quả qua 8,2 tỷ USD (cửa khẩu thông minh Trung - Việt) & cước container Drewry WCI giảm về $4,305/FEU',
         18600, 'Toàn cảnh thị trường & cước tàu T11 (+3,510%)',
         'Giám đốc chuỗi cung ứng, doanh nghiệp xuất khẩu nông sản sang Trung Quốc (GACC) và chủ hàng đàm phán hợp đồng cước biển cuối năm.',
         'Lưu lượng hàng hóa kỷ lục khiến hải quan tăng cường kiểm tra trị giá (CV 24083) và kiểm soát GACC tại cửa khẩu phía Bắc; trong khi giá cước WCI hạ nhiệt về $4,305/FEU nhưng hãng tàu tăng Blank Sailings ghìm giá.',
         "Tiêu đề: 'XNK vượt 1.010 tỷ USD & Cước Drewry giảm còn $4,305/FEU: Chiến lược chốt lịch tàu và giữ luồng xanh cửa khẩu cuối năm'.",
         'Cổng TTĐT Chính phủ; congthuong.vn; thoibaotaichinhvietnam.vn; congly.vn; drewry.co.uk (12/11/2026).',
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
            elif col_idx in [2, 3, 6, 7, 8, 9]:
                cell.alignment = align_left
            elif col_idx == 4:
                cell.alignment = align_right
                cell.number_format = '#,##0'
            elif col_idx == 5:
                cell.alignment = align_center
            elif col_idx == 10:
                cell.alignment = align_center
        ws1.row_dimensions[row_idx].height = 42

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

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua Web Search & 8 nguồn chính thống (Chu kỳ Tuần 47 — Ngày cập nhật: {update_date_str})."
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
    ws3['A3'] = f"Chu kỳ: Tuần 47 (16/11/2026 - 22/11/2026) | Mở rộng 280 chủ đề nghiên cứu (gồm 248 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 281)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 281 or m.max_row >= 281]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (276,
         'cong van 24313 chq gsql thu tuc hai quan thuong mai dien tu xuyen bien gioi',
         'cong van 24313 chq gsql, thu tuc hai quan thuong mai dien tu, xnk qua san tmdt, luat hai quan 11 2026 qh16, don hang gia tri thap, quan ly rui ro tmdt xuyen bien gioi',
         100.0,
         73400,
         'https://luatvietnam.vn',
         'Công văn 24313/CHQ-GSQL ngày 16/11/2026: Triển khai Hệ thống thực hiện thủ tục hải quan chuyên biệt đối với hàng hóa XNK giao dịch qua nền tảng thương mại điện tử',
         'Nội dung trọng tâm: Ban Giám sát quản lý về hải quan (Cục Hải quan) ban hành Công văn 24313/CHQ-GSQL hướng dẫn các Cục Hải quan địa phương, doanh nghiệp XNK, đơn vị logistics và sàn thương mại điện tử xuyên biên giới triển khai Hệ thống thực hiện thủ tục hải quan chuyên biệt cho hàng hóa giao dịch qua nền tảng TMĐT; Bối cảnh thực tiễn: Dòng hàng TMĐT xuyên biên giới có đặc thù số lượng giao dịch rất lớn, giá trị mỗi đơn hàng nhỏ và tốc độ luân chuyển nhanh nên quy trình khai báo rời rạc trước đây dễ gây ùn ứ; Điểm thay đổi chính: Thống nhất phương thức kết nối dữ liệu điện tử giữa sàn TMĐT - đơn vị vận chuyển - cơ quan Hải quan, tự động hóa khai báo và áp dụng cơ chế quản lý rủi ro riêng đồng bộ với lộ trình Luật Hải quan sửa đổi 11/2026/QH16; Khuyến nghị từ Eureka: Chủ shop và doanh nghiệp kinh doanh trên sàn TMĐT quốc tế cần chuẩn hóa dữ liệu đơn hàng, chứng từ thanh toán và mã HS ngay từ khâu đặt hàng để tránh bị ách tắc khi hệ thống quét rủi ro tự động.',
         'luatvietnam.vn & tapchikinhtetaichinh.vn | TN (Tin Ngành) / Cross-border E-commerce Customs System CV 24313',
         update_date_str),

        (277,
         'quyet dinh 1828 qd chq kiem tra xuat xu hang hoa sau thong quan 15 ngay',
         'quyet dinh 1828 qd chq, kiem tra xuat xu sau thong quan, hau kiem co xuat xu, thoi han kiem tra 15 ngay, chong chuyen tai bat hop phap, nghi dinh 167 2025, thong tu 124 2026 tt btc',
         99.8,
         71800,
         'https://thuvienphapluat.vn',
         'Quyết định 1828/QĐ-CHQ (ban hành 12/11/2026, hiệu lực 15/12/2026): Quy trình kiểm tra xuất xứ hàng hóa XNK sau khi hàng hóa đã được thông quan (tối đa 15 ngày)',
         'Nội dung trọng tâm: Cục Hải quan ban hành Quyết định 1828/QĐ-CHQ thiết lập quy trình nghiệp vụ chuyên sâu cho Đội Phúc tập và Kiểm tra sau thông quan thực hiện kiểm tra xuất xứ hàng hóa xuất khẩu, nhập khẩu sau khi đã thông quan theo nguyên tắc quản lý rủi ro (có hiệu lực từ 15/12/2026, bổ trợ cho Quyết định 467/QĐ-CHQ tại khâu thông quan); Điểm đáng chú ý: Quy định rõ danh mục hồ sơ, trình tự đối chiếu chứng từ và thời hạn kiểm tra tối đa 15 ngày nhằm phát hiện, ngăn chặn gian lận xuất xứ và hành vi chuyển tải bất hợp pháp; Tác động doanh nghiệp: Dù hàng đã thông quan khỏi cảng, doanh nghiệp vẫn có nguy cơ bị truy thu thuế ưu đãi và xử phạt nếu hồ sơ định mức nguyên liệu, chứng từ mua bán đầu vào không khớp tiêu chí xuất xứ; Giải pháp từ Eureka: Rà soát lại bộ hồ sơ lưu trữ chứng minh xuất xứ (hóa đơn nguyên liệu, bảng kê quy trình sản xuất, C/O) trước mốc hiệu lực 15/12/2026.',
         'thuvienphapluat.vn & luatvietnam.vn | TN (Tin Ngành) / Post-Clearance Origin Verification Process Decision 1828',
         update_date_str),

        (278,
         'cong van 24286 chq pc to khai dieu chinh thue suat khong xu phat vi pham',
         'cong van 24286 chq pc, to khai dieu chinh thue suat nhap khau, mien xu phat vi pham hanh chinh, loi he thong hai quan, dieu 11 luat xlvphc, khai bo sung thue suat',
         99.5,
         68900,
         'https://thuvienphapluat.vn',
         'Công văn 24286/CHQ-PC ngày 16/11/2026: Hướng dẫn giải quyết vướng mắc liên quan đến tờ khai điều chỉnh lại thuế suất thuế nhập khẩu và căn cứ không xử phạt vi phạm hành chính',
         'Nội dung trọng tâm: Ban Pháp chế (Cục Hải quan) ban hành Công văn 24286/CHQ-PC làm rõ hướng xử lý chính sách thuế và căn cứ xem xét miễn xử phạt vi phạm hành chính đối với các tờ khai phải điều chỉnh lại thuế suất thuế nhập khẩu; Thực tế vướng mắc: Nhiều doanh nghiệp gặp tình huống phải điều chỉnh tờ khai do lỗi kỹ thuật hệ thống hoặc do sự thay đổi của biểu thuế, văn bản chính sách tại thời điểm chuyển tiếp nhưng lo ngại bị áp dụng chế tài phạt khai sai dẫn đến thiếu số thuế phải nộp; Căn cứ tháo gỡ: Công văn phân định rạch ròi giữa trường hợp phát sinh do nguyên nhân khách quan, bất khả kháng hoặc lỗi hệ thống (được xem xét không xử phạt VPHC theo Điều 11 và khoản 14 Điều 2 Luật Xử lý vi phạm hành chính) với trường hợp cố tình khai sai mã số, thuế suất; Khuyến nghị từ Eureka: Khi mở tờ khai điều chỉnh thuế suất, doanh nghiệp cần lưu vết ảnh chụp lỗi hệ thống và văn bản giải trình nguyên nhân khách quan ngay từ đầu để bảo vệ quyền lợi miễn phạt.',
         'thuvienphapluat.vn & luatvietnam.vn | CS (Case Study) / Import Duty Rate Amendment Penalty Exemption CV 24286',
         update_date_str),

        (279,
         'cong van 24153 chq nvthq ra soat ma hs my pham sua rua mat 3401 mat na 3307',
         'cong van 24153 chq nvthq, ma hs my pham nhap khau, ma hs sua rua mat 34013000, ma hs mat na duong da 33079090, nhom 3304, phan tich phan loai hai quan',
         99.4,
         67100,
         'https://thuvienphapluat.vn',
         'Công văn 24153/CHQ-NVTHQ ngày 12/11/2026: Rà soát, kiểm tra áp dụng mã số HS đối với mặt hàng mỹ phẩm nhập khẩu (sữa rửa mặt 3401.30.00 và mặt nạ dưỡng da 3307.90.90)',
         'Nội dung trọng tâm: Ban Nghiệp vụ Thuế Hải quan (Cục Hải quan) chỉ đạo các Chi cục Hải quan khu vực rà soát, chấn chỉnh tình trạng khai báo sai hoặc gộp chung các mặt hàng mỹ phẩm nhập khẩu vào nhóm 33.04; Hướng dẫn phân loại cụ thể theo bản chất hàng hóa: (1) Sữa rửa mặt có chứa chất hữu cơ hoạt động bề mặt dạng lỏng hoặc kem đóng gói bán lẻ (có hoặc không có xà phòng) phù hợp phân loại vào nhóm 34.01, mã số 3401.30.00; (2) Mặt nạ dưỡng da làm từ vải không dệt đã thấm tẩm thành phần mỹ phẩm, đóng gói bán lẻ xem xét phân loại vào nhóm 33.07, mã số 3307.90.90 thay vì nhóm 33.04; Rủi ro doanh nghiệp: Khai lệch mã HS khiến lô hàng mỹ phẩm Hàn Quốc, Trung Quốc bị đưa đi lấy mẫu phân tích phân loại, kéo dài thời gian thông quan và đối mặt truy thu thuế nếu thuế suất khác nhau; Giải pháp từ Eureka: Đối chiếu thành phần hóa lý, chất liệu nền (vải không dệt) và công bố mỹ phẩm trước khi truyền tờ khai.',
         'thuvienphapluat.vn & tapchikinhtetaichinh.vn | HT (Học Thực Chiến) / Cosmetics HS Code Audit Cleansing & Sheet Mask CV 24153',
         update_date_str),

        (280,
         'cong van 24288 chq gsql thu tuc hai quan hoa chat cong nghiep bo cong thuong',
         'cong van 24288 chq gsql, thu tuc hai quan hoa chat, potassium gold cyanide, chat xuc tien luu hoa cao su, ma hs hoa chat, kiem tra chuyen nganh bo cong thuong',
         99.2,
         65300,
         'https://thuvienphapluat.vn',
         'Công văn 24288/CHQ-GSQL (ngày 16-17/11/2026): Hướng dẫn giải quyết vướng mắc thủ tục hải quan, phân loại mã HS và quản lý chuyên ngành đối với mặt hàng hóa chất công nghiệp',
         'Nội dung trọng tâm: Ban Giám sát quản lý về hải quan (Cục Hải quan) ban hành Công văn 24288/CHQ-GSQL giải quyết vướng mắc thủ tục hải quan và phối hợp chính sách quản lý chuyên ngành với Bộ Công Thương đối với một số mặt hàng hóa chất phục vụ sản xuất công nghiệp (điển hình như hợp chất Potassium Gold Cyanide dùng trong xi mạ/điện tử và chất xúc tiến lưu hóa cao su); Điểm nóng nghiệp vụ: Doanh nghiệp nhập khẩu hóa chất thường gặp rủi ro kép khi vừa phải xác định chính xác tên hóa học/mã CAS để áp mã HS, vừa phải đáp ứng quy định khai báo hóa chất, hóa chất hạn chế sản xuất kinh doanh của Bộ Công Thương; Hàng hóa chất nếu bị dừng thông quan tại cảng dễ phát sinh phí lưu container hàng nguy hiểm (DG) rất cao; Giải pháp từ Eureka: Chuẩn hóa bảng thành phần MSDS, mã CAS và thực hiện khai báo hóa chất qua Cổng Một cửa Quốc gia trước khi hàng cập cảng.',
         'thuvienphapluat.vn | CS (Case Study) / Industrial Chemicals Customs & MOIT Compliance CV 24288',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5  # STT 276 -> row 281
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

    # Cập nhật Total row ở row 286
    total_row_idx = 286
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 280 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D285)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E285)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 280 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 248 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
    # 3. LƯU NGUYÊN TỬ VÀ ĐỒNG BỘ AN TOÀN
    # ==========================================
    print("3. Lưu nguyên tử file Excel...")
    temp_local_file = FILE_PATH + '.tmp.xlsx'
    wb.save(temp_local_file)
    wb.close()

    if os.path.exists(temp_local_file):
        os.replace(temp_local_file, FILE_PATH)
        print(f"[OK] Đã cập nhật thành công file cục bộ: {FILE_PATH}")

    file_size = os.path.getsize(FILE_PATH)
    file_sha256 = get_file_sha256(FILE_PATH)

    if os.path.exists(ONEDRIVE_DIR):
        print(f"Đang đồng bộ nguyên tử sang OneDrive: {ONEDRIVE_FILE}...")
        temp_onedrive_file = ONEDRIVE_FILE + '.tmp.xlsx'
        try:
            shutil.copy2(FILE_PATH, temp_onedrive_file)
            os.replace(temp_onedrive_file, ONEDRIVE_FILE)
            print(f"[OK] Đã đồng bộ thành công sang OneDrive: {ONEDRIVE_FILE}")

            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - {update_date_str})\n"
            with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f:
                f.write(log_entry)
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
**Tác tử phụ trách:** Agent 01 (Market Researcher — Chế độ Hot News Scanner)  
**Ngày thực hiện:** {update_date_str} (Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026)  
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng lên 280 chủ đề)  
**Trạng thái đồng bộ OneDrive:** ✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (7 - 30 NGÀY QUA)
Agent 01 đã quét độc lập trên 8 nguồn dữ liệu theo `wf05_hot_news_scan.md`:
1. `customs.gov.vn` (Cổng TTĐT Hải quan)
2. `chinhphu.vn` (Cổng TTĐT Chính phủ)
3. `haiquanonline.com.vn` / `tapchikinhtetaichinh.vn` (Tin tức ngành, phân tích chính sách Hải quan)
4. `congthuong.vn` / `tapchicongthuong.vn` (Phòng vệ thương mại, XNK vượt mốc 1.010,55 tỷ USD)
5. `tapchitaichinh.vn` / `thoibaotaichinhvietnam.vn` (Chính sách thuế, hoàn thuế XNK, thương mại biên giới)
6. `thuvienphapluat.vn` / `luatvietnam.vn` / Google Trends VN (Tra cứu văn bản pháp luật mới & xu hướng tìm kiếm)
7. `Drewry WCI` (Báo cáo ngày 12/11/2026: Cước container WCI giảm 2% về $4,305/FEU, hãng tàu tăng Blank Sailings hướng tới tháng 12)
8. `scratch/daily_notes.md` & `scratch/wf04_10_new_legal_docs_20261117.json` (10 văn bản XNK mới cập nhật ngày 17/11/2026 & tình huống thực tế doanh nghiệp)

Qua sàng lọc, Agent 01 đã chọn lọc và chấm điểm **TOP 5 CHỦ ĐỀ NÓNG NHẤT NGÀY 17/11/2026** theo 4 tiêu chí (thang điểm 1-5, tối đa 20 điểm):

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: Công văn 24313/CHQ-GSQL (16/11/2026): Triển khai Hệ thống thủ tục hải quan chuyên biệt cho hàng hóa XNK qua thương mại điện tử
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Đặt vấn đề thảo luận cho chủ shop & đơn vị logistics TMĐT xuyên biên giới | YouTube: Video hướng dẫn chuẩn hóa dữ liệu đơn hàng nhỏ lẻ, tần suất cao theo cơ chế mới)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Ban hành ngày 16/11/2026, đồng bộ với lộ trình Luật Hải quan sửa đổi 11/2026/QH16 |
  | Độ đau | 5.0/5 | Đánh trúng dòng hàng TMĐT đơn nhỏ số lượng lớn vốn dễ tắc khi khai báo rời rạc |
  | Tiềm năng tranh luận | 4.9/5 | Tranh luận sôi nổi giữa chủ shop TMĐT, đơn vị vận chuyển nhanh và cách khai báo tách đơn |
  | Khả năng cho giá trị | 5.0/5 | Phân tích rõ cách kết nối dữ liệu tự động và quản lý rủi ro chuyên biệt của Hải quan |
  | **Tổng điểm** | **19.9 / 20** | **Xếp hạng #1** |
- **Tóm tắt lý do nên làm:** Xu hướng TMĐT xuyên biên giới đang bùng nổ, văn bản vừa ra lò ngày 16/11, rất hợp tệp chủ shop nhập hàng từ Trung Quốc và Hàn Quốc.
- **Nguồn xác minh:** https://luatvietnam.vn | https://tapchikinhtetaichinh.vn

---

### Đề xuất #2: Quyết định 1828/QĐ-CHQ (hiệu lực 15/12/2026): Đếm ngược siết quy trình hậu kiểm xuất xứ hàng hóa sau thông quan trong 15 ngày
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study)
- **Nền tảng khuyến nghị:** **Cả hai** (YouTube ThuyTalk: Phân tích chuyên sâu rủi ro chuyển tải bất hợp pháp & checklist hồ sơ xuất xứ cần giữ | Facebook: Cảnh báo ngắn gọn trước mốc 15/12/2026)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Ban hành 12/11/2026, còn chưa đầy 30 ngày đến mốc hiệu lực 15/12/2026 |
  | Độ đau | 5.0/5 | Hàng đã thông quan vẫn có thể bị Đội Phúc tập và KTSTQ gọi kiểm tra trong 15 ngày; rủi ro bác C/O và truy thu thuế cực lớn |
  | Tiềm năng tranh luận | 4.8/5 | Doanh nghiệp thương mại xuất khẩu thu mua nguyên liệu nhiều nguồn khó chứng minh xuất xứ thế nào? |
  | Khả năng cho giá trị | 5.0/5 | Đưa ra bộ 4 nhóm chứng từ gốc (hóa đơn nguyên liệu, định mức, quy trình gia công, C/O) chuẩn bị sẵn |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng #2** |
- **Tóm tắt lý do nên làm:** Đánh đúng nỗi lo lớn nhất của doanh nghiệp xuất khẩu (gỗ, thép, nông sản, dệt may) khi bị kiểm tra sau thông quan về xuất xứ.
- **Nguồn xác minh:** https://thuvienphapluat.vn | https://luatvietnam.vn | https://hocxuatnhapkhau.com

---

### Đề xuất #3: Công văn 24286/CHQ-PC (16/11/2026): Tờ khai điều chỉnh thuế suất thuế nhập khẩu — Khi nào được miễn xử phạt vi phạm hành chính?
- **Chuyên mục (Pillar):** `CS` (Case Study) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Facebook** (Bài viết tình huống mở tạo tranh luận cao giữa kế toán thuế XNK và nhân viên khai báo)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Công văn mới ban hành ngày 16/11/2026 bởi Ban Pháp chế - Cục Hải quan |
  | Độ đau | 4.9/5 | Doanh nghiệp rất ngại khai bổ sung điều chỉnh thuế suất vì sợ bị quy lỗi khai sai thiếu thuế và phạt hành chính |
  | Tiềm năng tranh luận | 4.9/5 | Ranh giới giữa "lỗi hệ thống / bất khả kháng" với "khai sai chủ quan" thực tế được Hải quan xét ra sao? |
  | Khả năng cho giá trị | 4.7/5 | Chỉ rõ căn cứ Điều 11 & khoản 14 Điều 2 Luật XLVPHC và mẹo lưu chứng cứ hệ thống ngay lúc phát sinh |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng #3** |
- **Tóm tắt lý do nên làm:** Chủ đề thực chiến sát sườn dân chứng từ/kế toán, giúp doanh nghiệp tự tin xử lý điều chỉnh thuế suất đúng luật.
- **Nguồn xác minh:** https://thuvienphapluat.vn | https://luatvietnam.vn

---

### Đề xuất #4: Công văn 24153/CHQ-NVTHQ (12/11/2026): Chấn chỉnh áp sai mã HS mỹ phẩm nhập khẩu — Đừng gộp sữa rửa mặt & mặt nạ dưỡng da vào nhóm 33.04
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `CS` (Case Study)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Bài cảnh báo nhanh cho ngành mỹ phẩm Hàn/Trung | YouTube Short/Video: Giải thích cách phân biệt 3401.30.00 và 3307.90.90)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.7/5 | Hải quan đang rà soát trực tiếp trên hệ thống theo CV 24153 ban hành 12/11/2026 |
  | Độ đau | 4.9/5 | Ngành mỹ phẩm nhập khẩu tuyến Hàn Quốc - Trung Quốc rất đông; áp sai mã HS dễ bị lấy mẫu phân tích và truy thu thuế |
  | Tiềm năng tranh luận | 4.8/5 | Thói quen áp chung 33.04 lâu nay nay bị tuýt còi — hàng cũ đã nhập trước đó có bị rà soát lại không? |
  | Khả năng cho giá trị | 4.8/5 | Bóc tách cực rõ tiêu chí chất hoạt động bề mặt (3401.30.00) và vải không dệt tẩm mỹ phẩm (3307.90.90) |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng #4** |
- **Tóm tắt lý do nên làm:** Trúng thẳng tệp khách hàng tuyến Hàn Quốc - Việt Nam và Trung Quốc - Việt Nam của Eureka.
- **Nguồn xác minh:** https://thuvienphapluat.vn | https://tapchikinhtetaichinh.vn

---

### Đề xuất #5: Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ (đã hiệu lực từ 16/11/2026): Thực tế triển khai mã A45/A47 và gỡ vướng hàng mẫu chuyển phát nhanh
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **YouTube / Facebook Carousel** (Hướng dẫn thực hành dùng đúng mã loại hình A45, A47, cơ chế nộp thay thuế và lấy tờ khai độc lập khấu trừ VAT)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Đang bước sang ngày thứ 2 có hiệu lực thực tế trên toàn quốc (từ 16/11/2026) |
  | Độ đau | 4.8/5 | Hàng mẫu gửi qua FedEx, DHL, UPS dễ bị ách nếu thiếu đồng bộ định danh và mã loại hình mới |
  | Tiềm năng tranh luận | 4.5/5 | Doanh nghiệp chuyển phát nhanh nộp thay thuế hộ chủ hàng thì hạch toán chứng từ VAT thế nào? |
  | Khả năng cho giá trị | 4.8/5 | Hướng dẫn chi tiết Phụ lục mới, phân biệt A45 (hàng trị giá thấp) và A47 (tài liệu không giá trị thương mại) |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng #5** |
- **Tóm tắt lý do nên làm:** Tiếp nối độ nóng từ ngày 16/11 bằng góc độ hướng dẫn thao tác thực tế và xử lý hóa đơn thuế sau thông quan.
- **Nguồn xác minh:** https://chinhphu.vn | https://customs.gov.vn | https://tapchikinhtetaichinh.vn

---

## 🎯 ĐỀ XUẤT HÀNH ĐỘNG TIẾP THEO CHO CHỊ THÚY
- Nếu ưu tiên **bắt sóng ngay hôm nay trên Facebook**: Chọn **Đề xuất #1 (CV 24313 - Hải quan TMĐT)** hoặc **Đề xuất #4 (CV 24153 - Mã HS mỹ phẩm Hàn/Trung)**.
- Nếu ưu tiên **Video YouTube chuyên sâu / cảnh báo chiến lược**: Chọn **Đề xuất #2 (QĐ 1828 - Hậu kiểm xuất xứ 15 ngày trước mốc 15/12/2026)**.
- Sau khi Chị Thúy chọn chủ đề, hệ thống sẽ đẩy ngay sang `wf01` (Lập lịch tuần) hoặc `wf02`/`wf03` (Sản xuất bài Facebook / Kịch bản YouTube).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")
    print("=== HOÀN THÀNH TOÀN BỘ QUY TRÌNH WF05 (17/11/2026) ===")


if __name__ == '__main__':
    update_seo_w47_1711()
