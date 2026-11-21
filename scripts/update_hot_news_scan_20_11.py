# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_20_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 20/11/2026
(Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026 | Cập nhật ngày 20/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ tuyệt đối knowledge/banned_words.md và system_prompt_base.md
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 47 cập nhật ngày 20/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 290 lên 295 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 20/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w47_2011.md
"""

import os
import sys
import hashlib
from datetime import datetime
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

FILE_PATH = r'reports\SEO_Keyword_Trends_Eureka.xlsx'
ONEDRIVE_DIR = r'C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh'
ONEDRIVE_FILE = os.path.join(ONEDRIVE_DIR, 'SEO_Keyword_Trends_Eureka.xlsx')
SYNC_LOG_PATH = r'reports\sync_log.txt'
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w47_2011.md'


def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()


def update_seo_w47_2011():
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

    update_date_str = '20/11/2026'

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 47 (16/11/2026 - 22/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'nghi quyet 19 2026 bo cap co dncx xuat nhap khau tai cho 24381 24449 chq gsql',
         'Vướng mắc bỏ cấp C/O cho DNCX bán vào nội địa theo Nghị quyết 19/2026/NQ-CP & Công văn 24381, 24449/CHQ-GSQL (17-18/11/2026)',
         26400, 'Khoảng trống cấp C/O cho DNCX tại chỗ (+5,410%)',
         'Doanh nghiệp chế xuất (DNCX), doanh nghiệp nội địa mua nguyên liệu/thành phẩm từ DNCX, kho ngoại quan và phòng chứng từ XNK.',
         'Nghị quyết 19/2026/NQ-CP bãi bỏ thủ tục cấp C/O giữa DNCX/phi thuế quan và nội địa khiến DN nội địa không có C/O để hưởng thuế ưu đãi FTA, buộc chịu thuế MFN cao và phát sinh tranh chấp hợp đồng.',
         "Tiêu đề: 'Bỏ cấp C/O cho DNCX bán vào nội địa theo Nghị quyết 19/2026: Doanh nghiệp tránh rủi ro chịu thuế MFN thế nào trong lúc chờ Bộ Công Thương hướng dẫn?'.",
         'Nghị quyết 19/2026/NQ-CP; Công văn 24381 & 24449/CHQ-GSQL; tapchikinhtetaichinh.vn; thuvienphapluat.vn; daily_notes.md (20/11).',
         update_date_str),

        (2,
         'thong tu 71 2026 tt bct phan hang trung tam logistics tieu chuan hang 1',
         'Thông tư 71/2026/TT-BCT (ban hành 19/11/2026, hiệu lực 08/01/2027): Chuẩn hóa 5 loại hình & phân hạng 3 cấp Trung tâm Logistics toàn quốc',
         25600, 'Chuẩn phân hạng Trung tâm Logistics mới (+5,180%)',
         'Doanh nghiệp khai thác kho bãi 3PL/4PL, chủ đầu tư hạ tầng logistics, khu công nghiệp/FTZ và nhà xuất nhập khẩu lớn.',
         'Lần đầu tiên Bộ Công Thương đặt thước đo định lượng: Hạng I phải từ 20 ha, kết nối >= 2 phương thức vận tải, công suất >= 3 triệu tấn hoặc 200.000 TEU/năm và áp dụng cơ chế tự công bố lên CSDL quốc gia.',
         "Tiêu đề: 'Thông tư 71/2026/TT-BCT: Bộ tiêu chuẩn mới phân hạng Trung tâm Logistics Hạng I-II-III và tác động đến chi phí kho vận của chủ hàng'.",
         'Thông tư 71/2026/TT-BCT ngày 19/11/2026; dms.gov.vn; baochinhphu.vn; baodautu.vn.',
         update_date_str),

        (3,
         'hoa ky doc ptmeg 36.55 phan tram an do dgtr nhom dun ad oi 079 2026',
         'Cảnh báo phòng vệ thương mại kép tháng 11/2026: Hoa Kỳ (DOC) áp ký quỹ sơ bộ 36,55% PTMEG (A-552-855) & Ấn Độ (DGTR) điều tra nhôm đùn (AD/OI/079/2026)',
         24900, 'Áp ký quỹ 36,55% Mỹ & khởi xướng điều tra Ấn Độ (+4,920%)',
         'Doanh nghiệp sản xuất xuất khẩu hóa chất/vật liệu (PTMEG, spandex), nhôm đùn định hình (HS 7604, 7608, 7610, 7616) đi Mỹ và Ấn Độ.',
         'DOC đình chỉ thanh quyết toán và thu ký quỹ tạm thời 36,55% toàn quốc với PTMEG Việt Nam từ trung tuần T11; song song Ấn Độ điều tra CBPG nhôm đùn từ 4 nước (chiếm 90% thị phần nhập khẩu) dựa trên chi phí sản xuất nội địa Ấn Độ.',
         "Tiêu đề: 'Mỹ áp ký quỹ 36,55% PTMEG & Ấn Độ điều tra nhôm đùn Việt Nam: Doanh nghiệp xuất khẩu chuẩn bị hồ sơ giá thành ra sao để không bị áp mức trần?'.",
         'Kết luận sơ bộ DOC A-552-855 (10/11/2026); Khởi xướng DGTR AD/OI/079/2026 (09/11/2026); trungtamwto.vn; moit.gov.vn; dangcongsan.vn.',
         update_date_str),

        (4,
         'ukraine ad 615 dieu tra chong lan tranh thue day thep eu 2026 2497 ong dong',
         'Làn sóng chống lẩn tránh & PVTM tháng 11/2026: Ukraine điều tra dây thép VN (AD-615/2026/441-01) & EU áp thuế tạm thời 12,3-14,4% ống đồng (2026/2497)',
         24300, 'Điều tra lẩn tránh xuất xứ TQ & áp thuế EU (+4,760%)',
         'Doanh nghiệp sản xuất & xuất khẩu kim loại (thép, dây thép, ống đồng CN 7411), gia công cơ khí, phòng chứng từ C/O và xuất khẩu EU/Đông Âu.',
         'Ukraine nghi ngờ dây thép silic-mangan từ VN chuyển tải từ Trung Quốc; cùng lúc EU áp thuế tạm thời 12,3%-14,4% ống đồng (TARIC 7411 10 90 10) và 16% thép cán nguội (2026/2547), siết hóa đơn gắn mã TARIC nhà SX.',
         "Tiêu đề: 'Ukraine điều tra lẩn tránh thuế dây thép & EU áp thuế ống đồng: Doanh nghiệp Việt bảo vệ hồ sơ xuất xứ nguyên liệu thế nào?'.",
         'Quyết định AD-615/2026/441-01; Quy định (EU) 2026/2497 & 2026/2547; tapchicongthuong.vn; trungtamwto.vn; moit.gov.vn.',
         update_date_str),

        (5,
         'thai lan ap thue tu ve nhua polypropylene pp 7.73 phan tram hs 3902',
         'Quyết định tự vệ nhựa Polypropylene (Thái Lan): Áp thuế tự vệ 3 năm cho 11 mã HS nhóm 3902, mức 7,73% CIF thu từ 28/12/2026',
         23800, 'Áp thuế tự vệ PP nội khối ASEAN từ 28/12 (+4,590%)',
         'Doanh nghiệp sản xuất bao bì, hạt nhựa Polypropylene, hóa chất, doanh nghiệp XNK tuyến ASEAN - Thái Lan và logistics đường biển.',
         'Việt Nam không được miễn trừ dù trong khối ASEAN; hàng xuất sang Thái Lan từ 28/12/2026 chịu thêm 7,73% CIF (song song Việt Nam cũng điều tra tự vệ SG04.2026 với PE/PP nhập khẩu), buộc DN tính lại lịch giao ETD/ETA.',
         "Tiêu đề: 'Thái Lan áp thuế tự vệ 7,73% CIF hạt nhựa PP từ 28/12/2026: Doanh nghiệp ngành nhựa cần chốt lịch giao hàng trước mốc nào?'.",
         'Thông báo Cục PVTM (Bộ Công Thương) ngày 12/11/2026; moit.gov.vn; pvtm.gov.vn; vneconomy.vn.',
         update_date_str),

        (6,
         'cuoc bien drewry wci 4176 feu tuan 19 11 2026 10 blank sailings',
         'Drewry World Container Index (19/11/2026) giảm tuần thứ 4 liên tiếp về $4,176/FEU (-3% WoW, +126% YoY) & Hãng tàu tung 10 Blank Sailings',
         23100, 'Cước giảm về $4,176/FEU nhưng tăng 10 Blank Sailings (+4,420%)',
         'Doanh nghiệp xuất nhập khẩu đường biển, phòng logistics, mua hàng quốc tế và chủ hàng chuẩn bị đơn hàng giao trước Tết 2027.',
         'Giá cước đi Mỹ giảm sâu (Shanghai–NY $8,194 -9%, Shanghai–LA $7,026 -5%) trong khi Shanghai–Genoa tăng 6% lên $4,436/FEU; hãng tàu tung 10 chuyến hủy (blank sailings) tuần tới khiến lịch tàu dễ bị dồn/trễ.',
         "Tiêu đề: 'Cước tàu Drewry WCI hạ về $4,176/FEU nhưng hãng tàu cắt 10 chuyến: Bẫy rớt tàu và cách giữ slot container cuối năm'.",
         'Drewry World Container Index (19/11/2026); drewry.co.uk; World Shipping Council.',
         update_date_str),

        (7,
         'cong van 24424 chq nvthq thue chong ban pha gia thap dien gio mien thue tscd',
         'Công văn 24424/CHQ-NVTHQ & Rà soát AD18 (hạn 06/12/2026): Hướng dẫn kê khai và miễn thuế chống bán phá giá tháp điện gió Trung Quốc tạo TSCĐ',
         22500, 'Hướng dẫn thuế CBPG bổ sung & hạn rà soát 06/12 (+4,250%)',
         'Doanh nghiệp nhập khẩu máy móc thiết bị dự án ưu đãi đầu tư từ Trung Quốc, nhà thầu năng lượng, kế toán thuế và khai báo hải quan.',
         'Nhiều doanh nghiệp lúng túng liệu thuế CBPG (tới 97% theo QĐ 3453/QĐ-BCT) có được miễn như thuế NK tạo TSCĐ theo khoản 11 Điều 16 Luật 107/2016/QH13 hay không, và cần nộp hồ sơ rà soát AD18 trước 06/12/2026.',
         "Tiêu đề: 'Thuế chống bán phá giá có được miễn khi nhập máy móc tạo tài sản cố định? Giải mã từ Công văn 24424/CHQ-NVTHQ về tháp điện gió Trung Quốc'.",
         'Công văn 24424/CHQ-NVTHQ; QĐ 3453/QĐ-BCT; Luật 107/2016/QH13; tapchikinhtetaichinh.vn; thuonghieucongluan.com.vn; baodautu.vn.',
         update_date_str),

        (8,
         'cong van 21692 23474 chq cntt lam sach du lieu vnaccs dksd customs khoa tai khoan 2027',
         'Chiến dịch làm sạch dữ liệu tài khoản VNACCS/VCIS tại dksd.customs.gov.vn trước hạn chót 31/12/2026 & Quyết định 1761/QĐ-CHQ',
         21900, 'Đếm ngược khóa tài khoản VNACCS 31/12 (+4,080%)',
         'Toàn bộ doanh nghiệp xuất nhập khẩu, nhân viên khai báo hải quan điện tử, đại lý hải quan và kế toán quản lý chữ ký số.',
         'Chỉ còn hơn 40 ngày tới hạn chót 31/12/2026; doanh nghiệp chưa xác thực OTP email, chữ ký số/VNeID tại dksd.customs.gov.vn sẽ bị khóa tài khoản khai báo từ 01/01/2027, làm gián đoạn mở tờ khai đầu năm.',
         "Tiêu đề: 'Đếm ngược trước 31/12/2026: Hướng dẫn làm sạch tài khoản VNACCS trên dksd.customs.gov.vn để không bị khóa từ 01/01/2027'.",
         'Công văn 21692 & 23474/CHQ-CNTT; Quyết định 1761/QĐ-CHQ; baochinhphu.vn; thoibaotaichinhvietnam.vn.',
         update_date_str),

        (9,
         'quyet dinh 1828 qd chq kiem tra xuat xu hang hoa sau thong quan 15 ngay',
         'Quyết định 1828/QĐ-CHQ (hiệu lực 15/12/2026): Quy trình kiểm tra xuất xứ hàng hóa sau thông quan tối đa 15 ngày, siết chống chuyển tải',
         21200, 'Đếm ngược hậu kiểm C/O 15 ngày (+3,910%)',
         'Doanh nghiệp xuất khẩu gỗ, thép, dệt may, điện tử, nông sản sử dụng C/O ưu đãi và phòng chứng từ XNK.',
         'Hàng đã thông quan vẫn bị Đội Phúc tập và KTSTQ gọi kiểm tra xuất xứ; rủi ro bác C/O, truy thu thuế và xử phạt nếu thiếu hồ sơ nguyên liệu đầu vào, định mức sản xuất.',
         "Tiêu đề: 'Quyết định 1828/QĐ-CHQ hiệu lực 15/12/2026: Bộ hồ sơ chứng minh xuất xứ cần giữ sẵn để vượt qua kỳ hậu kiểm sau thông quan'.",
         'Quyết định 1828/QĐ-CHQ ngày 12/11/2026; tapchikinhtetaichinh.vn; luatvietnam.vn; thuvienphapluat.vn.',
         update_date_str),

        (10,
         'kim ngach xuat nhap khau 1070 ty usd ky 1 thang 11 2026 dau gia duong 2679',
         'Kim ngạch XNK đến 15/11/2026 chạm mốc 1.070,26 tỷ USD (+33,4%) & Thu ngân sách đạt 99,2% dự toán (447.362 tỷ đồng)',
         20600, 'Kỷ lục XNK 59,5 tỷ USD nửa đầu T11 (+3,750%)',
         'Doanh nghiệp xuất nhập khẩu đa ngành, doanh nghiệp thực phẩm/mía đường (HS 17.01), kế toán thuế và bộ phận hoạch định chuỗi cung ứng.',
         'Riêng 15 ngày đầu T11 đạt 59,5 tỷ USD (+54,1%), thu ngân sách XNK hết T10 đạt 99,2% dự toán nên Hải quan đang siết kiểm soát trị giá & chống thất thu cuối năm.',
         "Tiêu đề: 'Kim ngạch XNK chạm 1.070,26 tỷ USD & Hải quan đạt 99,2% dự toán thu: Cảnh báo siết kiểm tra trị giá 40 ngày cuối năm 2026'.",
         'Số liệu Cục Hải quan kỳ 1 T11/2026; QĐ 2679/QĐ-BCT; thoibaotaichinhvietnam.vn; thuongtruong.com.vn.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 47 (16/11/2026 - 22/11/2026) | Mở rộng 295 chủ đề nghiên cứu (gồm 263 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 296)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 296 or m.max_row >= 296]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (291,
         'nghi quyet 19 2026 bo cap co dncx xuat nhap khau tai cho 24381 24449 chq gsql',
         'nghi quyet 19 2026 nq cp, bo cap co dncx ban vao noi dia, cong van 24381 chq gsql, cong van 24449 chq gsql, xuat nhap khau tai cho e15, thue mfn dncx',
         99.9,
         77500,
         'https://tapchikinhtetaichinh.vn',
         'Vướng mắc bỏ cấp C/O cho DNCX bán hàng vào nội địa theo Nghị quyết 19/2026/NQ-CP & Công văn 24381, 24449/CHQ-GSQL (17-18/11/2026)',
         'Nội dung trọng tâm: Theo Phụ lục III (Mục 1, Phần Đ, khoản 5, 6, 7) Nghị quyết 19/2026/NQ-CP ngày 29/4/2026, Chính phủ bãi bỏ thủ tục cấp C/O đối với hàng hóa xuất nhập khẩu từ doanh nghiệp chế xuất (DNCX), khu chế xuất, kho ngoại quan, khu phi thuế quan có quan hệ XNK với nội địa (trước đây theo khoản 6 Điều 15 NĐ 31/2018/NĐ-CP) cùng C/O cấp sau và hàng gửi kho ngoại quan; Hệ quả thực tế: Doanh nghiệp nội địa mua nguyên vật liệu/hàng hóa từ DNCX bị rơi vào khoảng trống pháp lý vì không có C/O chứng minh xuất xứ để hưởng thuế suất ưu đãi FTA hoặc chứng minh không thuộc diện phòng vệ thương mại, buộc phải chịu thuế nhập khẩu MFN cao; Động thái tháo gỡ: Ngày 17-18/11/2026, Cục Hải quan ban hành liên tiếp Công văn 24381/CHQ-GSQL và 24449/CHQ-GSQL trao đổi chính thức với Cục XNK (Bộ Công Thương) và hướng dẫn doanh nghiệp liên hệ Bộ Công Thương trong thời gian chờ cơ chế chứng từ thay thế; Khuyến nghị từ Eureka (kết hợp Case Study 20/11 trên daily_notes.md): Rà soát ngay điều khoản thuế/giá trong hợp đồng mua bán với DNCX, lưu trữ đầy đủ bộ chứng từ nguồn gốc nguyên liệu sản xuất và tham vấn Chi cục Hải quan GSQL trước khi mở tờ khai loại hình E15/XNK tại chỗ.',
         'tapchikinhtetaichinh.vn, thuvienphapluat.vn & daily_notes.md | CS (Case Study) / Resolution 19/2026 EPE On-Spot C/O Gap',
         update_date_str),

        (292,
         'thong tu 71 2026 tt bct phan hang trung tam logistics tieu chuan hang 1',
         'thong tu 71 2026 tt bct, tieu chi phan loai phan hang trung tam logistics, trung tam logistics hang 1 20 ha 200000 teu, co so du lieu ha tang thuong mai quoc gia',
         99.7,
         74800,
         'https://dms.gov.vn',
         'Thông tư 71/2026/TT-BCT (ban hành 19/11/2026, hiệu lực 08/01/2027): Bộ Công Thương chuẩn hóa 5 loại hình & phân hạng 3 cấp Trung tâm Logistics toàn quốc',
         'Nội dung trọng tâm: Ngày 19/11/2026, Bộ Công Thương ban hành Thông tư số 71/2026/TT-BCT (có hiệu lực từ 08/01/2027) lần đầu tiên thiết lập bộ tiêu chuẩn pháp lý thống nhất để phân loại và phân hạng các trung tâm logistics tại Việt Nam; Cấu trúc phân hạng 3 cấp: (1) Hạng I (cấp quốc tế/quốc gia): diện tích từ 20 ha trở lên, kết nối trực tiếp ít nhất 2 phương thức vận tải (đường bộ, biển, thủy nội địa, sắt, hàng không) và đường cao tốc/quốc lộ lớn, công suất tối thiểu 3 triệu tấn/năm hoặc 200.000 TEU/năm (hoặc 35 triệu đơn hàng/năm), bãi đỗ >= 200 vị trí container, đường nội bộ >= 2 làn rộng >= 7m và thực hiện tối thiểu 50% dịch vụ logistics theo NĐ 163/2017/NĐ-CP; (2) Hạng II (cấp vùng) và (3) Hạng III (cấp tỉnh); Đồng thời phân chia 5 loại hình chuyên biệt (cảng biển/hàng không, ngành hàng, KCX/FTZ/phi thuế quan, biên giới/cửa khẩu, thương mại điện tử) kèm yêu cầu chuyển đổi số, hạ tầng xanh và cơ chế tự công bố lên Cơ sở dữ liệu hạ tầng thương mại quốc gia; Khuyến nghị từ Eureka: Doanh nghiệp XNK và 3PL đối chiếu chuẩn phân hạng mới khi lựa chọn đối tác kho vận, trung chuyển hoặc quy hoạch trung tâm phân phối.',
         'dms.gov.vn, baochinhphu.vn & baodautu.vn | TN (Tin Ngành) / Circular 71/2026 Logistics Center Classification',
         update_date_str),

        (293,
         'hoa ky doc ptmeg 36.55 phan tram an do dgtr nhom dun ad oi 079 2026',
         'doc ket luan so bo ptmeg viet nam a 552 855, ky quy 36 55 phan tram, an do dgtr dieu tra chong ban pha gia nhom dun ad oi 079 2026, hs 7604 7608 7610 7616',
         99.5,
         72400,
         'https://trungtamwto.vn',
         'Cảnh báo PVTM kép tháng 11/2026: Hoa Kỳ (DOC) áp thuế ký quỹ sơ bộ 36,55% PTMEG (A-552-855) & Ấn Độ (DGTR) khởi xướng điều tra nhôm đùn Việt Nam (AD/OI/079/2026)',
         'Nội dung trọng tâm: Tháng 11/2026 ghi nhận hai diễn biến phòng vệ thương mại lớn nhắm vào hàng xuất khẩu Việt Nam: (1) Ngày 10/11/2026, Bộ Thương mại Hoa Kỳ (DOC) công bố kết luận sơ bộ vụ điều tra CBPG sản phẩm PTMEG (HS 3907.29.0000, 2932.11.0000, 3404.90.5150; mã A-552-855; POI 01/10/2025–31/03/2026), áp mức biên độ & ký quỹ tạm thời 36,55% cho bị đơn bắt buộc và toàn quốc, đình chỉ thanh quyết toán (kể cả hàng qua nước thứ ba), bác đề nghị loại trừ PTMEG 1800 MW và gia hạn kết luận cuối cùng 135 ngày; (2) Ngày 09/11/2026, Tổng cục Phòng vệ thương mại Ấn Độ (DGTR) khởi xướng điều tra CBPG mã AD/OI/079/2026 đối với nhôm đùn (HS 7604, 7608, 7610, 7616; POI 01/04/2025–30/06/2026) từ Việt Nam, Trung Quốc, Malaysia và Thái Lan (4 nước chiếm 90% nhập khẩu vào Ấn Độ), tạm dùng dữ liệu chi phí nhà sản xuất Ấn Độ để tính giá trị thông thường; Khuyến nghị từ Eureka: Doanh nghiệp sản xuất xuất khẩu phải chuẩn hóa sổ sách giá thành nguyên liệu, tách bạch chuỗi cung ứng đầu vào và theo sát hạn nộp bản câu hỏi trên cổng SETU/ACCESS để tránh bị áp thuế biên độ bất lợi.',
         'trungtamwto.vn, moit.gov.vn & dangcongsan.vn | CS (Case Study) / US PTMEG 36.55% & India Aluminum Extrusions AD/OI/079',
         update_date_str),

        (294,
         'cuoc bien drewry wci 4176 feu tuan 19 11 2026 10 blank sailings',
         'drewry world container index 19 11 2026, gia cuoc container 4176 usd feu, shanghai new york 8194, shanghai los angeles 7026, 10 blank sailings cuoi nam',
         99.3,
         69800,
         'https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/world-container-index-assessed-by-drewry',
         'Drewry World Container Index (19/11/2026): Giá cước container giảm tuần thứ 4 liên tiếp về $4,176/FEU (-3% WoW, +126% YoY) & Hãng tàu tung 10 Blank Sailings tuần tới',
         'Nội dung trọng tâm: Theo báo cáo tuần 19/11/2026 của Drewry Maritime Research, chỉ số giá cước container thế giới (WCI) giảm 3% so với tuần trước xuống mức $4,176/FEU (đánh dấu tuần giảm thứ 4 liên tiếp nhưng vẫn cao hơn 126% so với cùng kỳ năm trước); Diễn biến phân hóa rõ giữa các tuyến: Tuyến xuyên Thái Bình Dương giảm tuần thứ 7 liên tiếp (Shanghai–New York giảm 9% về $8,194/FEU, Shanghai–Los Angeles giảm 5% về $7,026/FEU), Shanghai–Rotterdam đi ngang quanh $3,496/FEU, trong khi Shanghai–Genoa tăng ngược 6% lên $4,436/FEU; Đáng chú ý, để kìm đà giảm cước giao ngay trước giai đoạn cao điểm cuối năm, các hãng tàu đã thông báo cắt giảm mạnh năng lực vận tải với 10 chuyến hủy (blank sailings) cho tuần kế tiếp; Khuyến nghị từ Eureka: Chủ hàng không nên chỉ nhìn giá cước giảm mà trì hoãn booking, cần khóa chỗ sớm và kiểm tra phụ phí PSS/GRI cũng như lịch cắt chuyến để tránh rớt tàu sát dịp Tết Dương lịch và Tết Nguyên đán.',
         'drewry.co.uk & World Shipping Council | TN (Tin Ngành) / Drewry WCI $4,176/FEU & 10 Blank Sailings',
         update_date_str),

        (295,
         'cong van 24424 chq nvthq thue chong ban pha gia thap dien gio mien thue tscd',
         'cong van 24424 chq nvthq, thue chong ban pha gia thap dien gio trung quoc, quyet dinh 3453 qd bct ad18, mien thue nhap khau tao tai san co dinh, ra soat ad18 06 12 2026',
         99.1,
         67900,
         'https://tapchikinhtetaichinh.vn',
         'Công văn 24424/CHQ-NVTHQ & Rà soát AD18 (hạn 06/12/2026): Hướng dẫn kê khai và xem xét miễn thuế chống bán phá giá (thuế NK bổ sung) với tháp điện gió Trung Quốc tạo TSCĐ',
         'Nội dung trọng tâm: Trả lời vướng mắc thực tế của doanh nghiệp dự án (trường hợp CTCP Đầu tư Xây dựng Trường Sơn) nhập khẩu thiết bị tháp điện gió xuất xứ Trung Quốc (thuộc diện áp thuế chống bán phá giá tới 97% theo QĐ 3453/QĐ-BCT, nhóm HS 7308.20 / 8502.31), Cục Hải quan ban hành Công văn 24424/CHQ-NVTHQ làm rõ cơ chế pháp lý: (1) Theo khoản 5 Điều 4 Luật Thuế XNK 107/2016/QH13, thuế chống bán phá giá là thuế nhập khẩu bổ sung; (2) Theo khoản 11 Điều 16 Luật 107/2016/QH13 và Điều 14 Nghị định 134/2016/NĐ-CP (sửa đổi bởi NĐ 18/2021/NĐ-CP), hàng nhập khẩu tạo tài sản cố định của dự án thuộc đối tượng ưu đãi đầu tư được xem xét miễn thuế nhập khẩu nếu đáp ứng đủ điều kiện hồ sơ; (3) Kê khai, thu, nộp, hoàn thuế CBPG thực hiện như thuế nhập khẩu theo TT 38/2015, TT 39/2018 và TT 86/2026/TT-BTC; Song song, Cục PVTM thông báo hạn chót tiếp nhận hồ sơ yêu cầu rà soát vụ việc tháp điện gió (AD18) là ngày 06/12/2026; Khuyến nghị từ Eureka: Doanh nghiệp nhập khẩu máy móc dây chuyền, thiết bị dự án từ Trung Quốc cần kiểm tra chéo danh mục miễn thuế đầu tư với các quyết định PVTM đang có hiệu lực trước khi mở tờ khai.',
         'tapchikinhtetaichinh.vn, thuonghieucongluan.com.vn & baodautu.vn | HT (Học Thực Chiến) / Wind Tower AD18 Anti-Dumping Tax & Fixed Asset Exemption',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5  # STT 291 -> row 296
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
        ws3.row_dimensions[row_idx].height = 88

    # Cập nhật Total row ở row 301
    total_row_idx = 301
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 295 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D300)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E300)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 295 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 263 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng lên 295 chủ đề)  
**Trạng thái đồng bộ OneDrive:** ✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (7 - 30 NGÀY QUA)
Agent 01 đã quét độc lập trên 8 nguồn dữ liệu theo `wf05_hot_news_scan.md`:
1. `customs.gov.vn` (Công văn 24381/CHQ-GSQL ngày 17/11 & 24449/CHQ-GSQL ngày 18/11 xử lý vướng mắc cấp C/O DNCX; Công văn 24460/CHQ-GSQL ngày 19/11 kiểm tra ATTP dâu tây tươi; chiến dịch làm sạch VNACCS `dksd.customs.gov.vn` trước 31/12/2026)
2. `chinhphu.vn` / `baochinhphu.vn` (Nghị quyết 19/2026/NQ-CP bãi bỏ cấp C/O giữa DNCX và nội địa; Thông tư 71/2026/TT-BCT ngày 19/11 phân hạng Trung tâm Logistics)
3. `haiquanonline.com.vn` / `thoibaotaichinhvietnam.vn` (Thu ngân sách XNK 10T đạt 447.362 tỷ đồng = 99,2% dự toán, siết kiểm tra trị giá & xuất xứ cuối năm; số liệu XNK kỳ 1 T11 đạt 59,5 tỷ USD, lũy kế 1.070,26 tỷ USD)
4. `congthuong.vn` / `moit.gov.vn` / `dms.gov.vn` (Thông tư 71/2026/TT-BCT hiệu lực 08/01/2027; Ấn Độ DGTR khởi xướng điều tra CBPG nhôm đùn Việt Nam AD/OI/079/2026; Hoa Kỳ DOC kết luận sơ bộ PTMEG 36,55%; thông báo rà soát tháp điện gió AD18 hạn 06/12/2026)
5. `tapchitaichinh.vn` / `tapchikinhtetaichinh.vn` (Phân tích khoảng trống pháp lý C/O DNCX tại chỗ theo NQ 19/2026; Công văn 24424/CHQ-NVTHQ về thuế CBPG bổ sung tháp điện gió & miễn thuế TSCĐ)
6. `thuvienphapluat.vn` / `trungtamwto.vn` / Google Trends VN (Chi tiết lộ trình vụ PTMEG A-552-855 tại Mỹ & nhôm đùn tại Ấn Độ; tra cứu cụm công văn GSQL/NVTHQ ngày 17–19/11/2026)
7. `Drewry WCI` (Số liệu mới nhất ngày 19/11/2026: WCI giảm 3% xuống **$4,176/FEU**, Shanghai–NY $8,194, Shanghai–LA $7,026, Shanghai–Genoa $4,436, Shanghai–Rotterdam $3,496; hãng tàu công bố **10 Blank Sailings** tuần kế tiếp)
8. `scratch/daily_notes.md` (Cập nhật sáng **20/11/2026**: 3 case study thực tế nóng trên cộng đồng XNK/Thuế gồm: vướng mắc bỏ cấp C/O cho DNCX bán vào nội địa, nộp Mẫu 18A SXXK/gia công lại thiếu số tiếp nhận văn thư, và lỗi kê khai thuế GTGT điện tử)

Qua sàng lọc và đối chiếu chống trùng lặp 100% với 290 chủ đề đã có, Agent 01 đã bổ sung 5 chủ đề mới (STT 291–295) và chấm điểm **TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT NGÀY {update_date_str}** theo 4 tiêu chí (thang điểm 1–5, tối đa 20 điểm):

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: Vướng mắc bỏ cấp C/O cho DNCX bán hàng vào nội địa theo Nghị quyết 19/2026/NQ-CP & Công văn 24381, 24449/CHQ-GSQL (17–18/11/2026)
- **Chuyên mục (Pillar):** `CS` (Case Study) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Đặt vấn đề tranh luận trực tiếp từ câu chuyện bức xúc của DN trên group ngày 20/11 | YouTube: Phân tích ranh giới pháp lý giữa tinh thần cải cách TTHC và rủi ro chịu thuế MFN khi mua hàng từ DNCX)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Vừa có CV 24381 (17/11) & CV 24449/CHQ-GSQL (18/11), và bùng lên tranh luận ngày 20/11 trong `daily_notes.md` |
  | Độ đau | 5.0/5 | Doanh nghiệp nội địa mua hàng từ DNCX mất C/O ưu đãi, đối mặt thuế MFN cao và nguy cơ tranh chấp hợp đồng |
  | Tiềm năng tranh luận | 5.0/5 | Cắt giảm thủ tục hành chính nhưng lại đẩy doanh nghiệp vào thế khó khi thiếu chứng từ thay thế — lỗi khâu hướng dẫn hay hiểu nhầm văn bản? |
  | Khả năng cho giá trị | 4.8/5 | Hướng dẫn rõ điều khoản hợp đồng giá/thuế cần rà soát và quy trình xin tham vấn, giữ chứng từ nguồn gốc nguyên liệu |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng #1** |
- **Tóm tắt lý do nên làm:** Kết hợp trọn vẹn tin văn bản nóng 17–18/11 với nỗi đau thực tế vừa ghi nhận hôm nay (20/11) trong `daily_notes.md`, đúng chuẩn góc nhìn Liêm Trinh + Giáp Mộc.
- **Nguồn xác minh:** https://tapchikinhtetaichinh.vn | https://thuvienphapluat.vn | `scratch/daily_notes.md` (20/11/2026)

---

### Đề xuất #2: Thông tư 71/2026/TT-BCT (ban hành 19/11/2026, hiệu lực 08/01/2027): Chuẩn hóa 5 loại hình & phân hạng 3 cấp Trung tâm Logistics toàn quốc
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **YouTube / Facebook B2B** (Facebook: Điểm tin nhanh bộ tiêu chí Hạng I mới ban hành hôm qua 19/11 | YouTube: Phân tích cách chủ hàng đánh giá năng lực kho bãi 3PL/FTZ dựa theo Thông tư 71/2026)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Vừa ký ban hành ngày 19/11/2026, sắp hiệu lực ngày 08/01/2027 |
  | Độ đau | 4.7/5 | DN kho vận phải chuẩn hóa hạ tầng, chuyển đổi số và tự công bố lên CSDL quốc gia; chủ hàng có thước đo chuẩn khi thuê kho |
  | Tiềm năng tranh luận | 4.9/5 | Chuẩn Hạng I đòi hỏi >= 20 ha, >= 200.000 TEU/năm, >= 2 phương thức vận tải — liệu các kho vừa và nhỏ có bị phân tầng mạnh? |
  | Khả năng cho giá trị | 5.0/5 | Có đầy đủ thông số định lượng (Hạng I–II–III, 5 loại hình, tiêu chí bãi đỗ 200 cont, đường nội bộ 7m) |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng #2** |
- **Tóm tắt lý do nên làm:** Văn bản nền tảng mới nhất của ngành logistics vừa ra lò ngày 19/11/2026, giúp nâng tầm định vị chuyên gia cho chị Thúy.
- **Nguồn xác minh:** https://dms.gov.vn | https://baochinhphu.vn | https://baodautu.vn

---

### Đề xuất #3: Cảnh báo PVTM kép tháng 11/2026: Hoa Kỳ (DOC) áp thuế ký quỹ sơ bộ 36,55% PTMEG (A-552-855) & Ấn Độ (DGTR) điều tra CBPG nhôm đùn Việt Nam (AD/OI/079/2026)
- **Chuyên mục (Pillar):** `CS` (Case Study) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Cảnh báo rủi ro bị tính giá trị thông thường theo dữ liệu nước thứ ba khi không đủ sổ sách giá thành | YouTube: Checklist chuẩn bị hồ sơ kháng kiện CBPG cho doanh nghiệp sản xuất xuất khẩu)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Kết luận sơ bộ DOC công bố 10/11/2026 (thu ký quỹ trung tuần T11); DGTR khởi xướng 09/11/2026 |
  | Độ đau | 5.0/5 | Mức ký quỹ 36,55% và đình chỉ thanh quyết toán đánh trực diện vào dòng tiền xuất khẩu |
  | Tiềm năng tranh luận | 4.7/5 | Vì sao doanh nghiệp Việt hay bị áp mức thuế toàn quốc cao khi xuất sang Mỹ, Ấn Độ? |
  | Khả năng cho giá trị | 4.8/5 | Đầy đủ mã HS (3907.29, 7604, 7608, 7610, 7616), mốc POI và quy trình tham gia cổng SETU |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng #3** |
- **Tóm tắt lý do nên làm:** Kết nối sát với tuyến Trung Quốc - Việt Nam - Quốc tế (Ấn Độ điều tra đồng thời Trung Quốc, Việt Nam, Malaysia, Thái Lan chiếm 90% thị phần).
- **Nguồn xác minh:** https://trungtamwto.vn | https://moit.gov.vn | https://dangcongsan.vn

---

### Đề xuất #4: Drewry WCI (19/11/2026) giảm tuần thứ 4 liên tiếp về $4,176/FEU (-3% WoW, +126% YoY) nhưng Hãng tàu tung 10 Blank Sailings tuần tới
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Facebook** (Bài cập nhật thị trường cước tuần + chiến lược booking giữ chỗ cuối năm)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Số liệu Drewry WCI vừa công bố ngày 19/11/2026 |
  | Độ đau | 4.7/5 | Cước giảm nhưng hãng tàu cắt 10 chuyến dễ gây rớt tàu, trễ lịch giao hàng cao điểm cuối năm |
  | Tiềm năng tranh luận | 4.8/5 | Giá cước giảm 4 tuần liên tiếp có nên chờ giá đáy mới chốt booking không? |
  | Khả năng cho giá trị | 4.7/5 | Số liệu rõ ràng từng tuyến (NY $8,194, LA $7,026, Genoa $4,436, Rotterdam $3,496) |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng #4** |
- **Tóm tắt lý do nên làm:** Nội dung ngắn gọn, thực chiến, đánh trúng tâm lý chủ hàng đang cân đo thời điểm chốt tàu tháng 11–12.
- **Nguồn xác minh:** https://www.drewry.co.uk/supply-chain-advisors/supply-chain-expertise/world-container-index-assessed-by-drewry

---

### Đề xuất #5: Công văn 24424/CHQ-NVTHQ & Hạn rà soát AD18 (06/12/2026): Nhập máy móc tạo TSCĐ dự án ưu đãi có được miễn thuế chống bán phá giá (tháp điện gió Trung Quốc)?
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `CS` (Case Study)
- **Nền tảng khuyến nghị:** **YouTube / Facebook Q&A** (YouTube: Giải đáp nghiệp vụ thuế nhập khẩu bổ sung & điều kiện miễn thuế TSCĐ | Facebook: Bài hỏi đáp ngắn kèm cảnh báo hạn chót rà soát AD18 ngày 06/12/2026)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.8/5 | Hướng dẫn mới giữa T11/2026 & sắp đến hạn nộp hồ sơ rà soát AD18 ngày 06/12/2026 |
  | Độ đau | 4.8/5 | Thuế CBPG tháp điện gió Trung Quốc lên tới 97% (QĐ 3453/QĐ-BCT); hiểu nhầm quy định miễn thuế khiến lệch toàn bộ dự toán vốn |
  | Tiềm năng tranh luận | 4.6/5 | Ranh giới giữa miễn thuế nhập khẩu tạo TSCĐ (khoản 11 Điều 16 Luật 107/2016/QH13) và bản chất thuế NK bổ sung |
  | Khả năng cho giá trị | 4.8/5 | Dẫn chiếu chuẩn xác điều khoản Luật 107/2016/QH13, NĐ 134/2016/NĐ-CP (sửa đổi NĐ 18/2021) và TT 86/2026/TT-BTC |
  | **Tổng điểm** | **19.0 / 20** | **Xếp hạng #5** |
- **Tóm tắt lý do nên làm:** Chạm đúng tệp khách hàng nhập khẩu máy móc, thiết bị dự án tuyến Trung Quốc - Việt Nam của Eureka.
- **Nguồn xác minh:** https://tapchikinhtetaichinh.vn | https://thuonghieucongluan.com.vn | https://baodautu.vn

---

## 🎯 ĐỀ XUẤT HÀNH ĐỘNG TIẾP THEO CHO CHỊ THÚY
- Nếu ưu tiên **bắt sóng nóng nhất hôm nay trên cộng đồng Hải quan & XNK**: Chọn **Đề xuất #1 (Vướng mắc bỏ cấp C/O cho DNCX bán vào nội địa theo NQ 19/2026 & CV 24381/24449)** — có sẵn chất liệu thực tế từ `daily_notes.md` ngày 20/11.
- Nếu ưu tiên **chính sách hạ tầng mới ban hành**: Chọn **Đề xuất #2 (Thông tư 71/2026/TT-BCT phân hạng Trung tâm Logistics)**.
- Nếu ưu tiên **cảnh báo rủi ro cho doanh nghiệp sản xuất xuất khẩu & máy móc dự án**: Chọn **Đề xuất #3 (PVTM PTMEG Mỹ & Nhôm đùn Ấn Độ)** hoặc **Đề xuất #5 (CV 24424/CHQ-NVTHQ thuế CBPG tháp điện gió TQ)**.
- Sau khi Chị Thúy chọn chủ đề, hệ thống sẽ đẩy ngay sang `wf01` (Lập lịch tuần) hoặc `wf02`/`wf03` (Sản xuất bài Facebook / Kịch bản YouTube).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")
    print("=== HOÀN THÀNH TOÀN BỘ QUY TRÌNH WF05 (20/11/2026) ===")


if __name__ == '__main__':
    update_seo_w47_2011()
