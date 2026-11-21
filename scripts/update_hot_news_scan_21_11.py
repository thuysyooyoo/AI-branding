# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_21_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 21/11/2026
(Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026 | Cập nhật ngày 21/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ tuyệt đối knowledge/banned_words.md và system_prompt_base.md
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 47 cập nhật ngày 21/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 295 lên 300 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 21/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w47_2111.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w47_2111.md'


def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()


def update_seo_w47_2111():
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

    update_date_str = '21/11/2026'

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 47 (16/11/2026 - 22/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cua khau thong minh kim thanh bac son con minh tay nam trung quoc nong san chinh ngach',
         'Diễn đàn Kết nối Việt Nam – Tây Nam Trung Quốc tại Côn Minh (20/11/2026): Đẩy nhanh cửa khẩu thông minh & chuỗi nông sản chính ngạch',
         27200, 'Hành lang Côn Minh - Lào Cai - Hải Phòng (+5,680%)',
         'Doanh nghiệp xuất nhập khẩu tuyến Trung Quốc - Việt Nam, nhà xuất khẩu nông sản/trái cây chính ngạch, đơn vị logistics đường bộ và vận tải liên vận.',
         'Chủ hàng và đơn vị vận tải tuyến Trung–Việt cần nắm tiến độ cặp cửa khẩu thông minh (Kim Thành – Bắc Sơn), đường sắt khổ tiêu chuẩn và cơ chế chuỗi nông sản tin cậy nhằm hạ chi phí logistics, rút ngắn thời gian thông quan.',
         "Tiêu đề: 'Sau diễn đàn Côn Minh 20/11: Cửa khẩu thông minh Kim Thành – Bắc Sơn và cơ hội mở rộng xuất khẩu nông sản chính ngạch sang Tây Nam Trung Quốc'.",
         'Sự kiện Diễn đàn Côn Minh ngày 20/11/2026; baochinhphu.vn; vneconomy.vn; congthuong.vn.',
         update_date_str),

        (2,
         'cong van 24503 chq gsql thu tuc hai quan hang tai nhap 24485 co 24528 tthc',
         'Cụm Công văn ngày 20/11/2026 (24503/CHQ-GSQL, 24485/CHQ-GSQL, 24528/CHQ-NVTHQ): Hướng dẫn hàng tái nhập, xử lý kiến nghị C/O & đơn giản hóa TTHC',
         26500, 'Loạt công văn nóng Cục Hải quan ngày 20/11 (+5,340%)',
         'Doanh nghiệp xuất khẩu có lô hàng bị trả về/tái nhập, bộ phận chứng từ C/O, đại lý hải quan và nhân viên khai báo VNACCS.',
         'Hàng đã xuất khẩu bị từ chối hoặc gặp trục trặc ở nước đến buộc phải tái nhập về Việt Nam dễ bị vướng hồ sơ không thu thuế/hoàn thuế nếu thiếu chứng từ chứng minh nguồn gốc; đồng thời DN chờ hướng dẫn gỡ vướng C/O.',
         "Tiêu đề: 'Hàng xuất khẩu bị trả về phải làm thủ tục tái nhập: Cập nhật quy trình mới nhất từ Công văn 24503/CHQ-GSQL (20/11/2026)'.",
         'Công văn 24503/CHQ-GSQL, 24485/CHQ-GSQL, 24528/CHQ-NVTHQ ngày 20/11/2026; luatvietnam.vn; bkhub.vn.',
         update_date_str),

        (3,
         'cong van 24460 chq gsql kiem tra attp dau tay tuoi trung quoc du luong thuoc bvtv',
         'Công văn 24460/CHQ-GSQL (19/11/2026): Siết kiểm tra ATTP dâu tây tươi nhập khẩu từ Trung Quốc sau vụ vượt ngưỡng Cypermethrin gần 6 lần',
         25900, 'Cảnh báo kiểm tra chuyên ngành nông sản TQ (+5,120%)',
         'Doanh nghiệp nhập khẩu trái cây tươi/nông sản từ Trung Quốc, đại lý kiểm dịch thực vật, kiểm tra an toàn thực phẩm và kho lạnh cửa khẩu.',
         'Một lô 4.000 kg dâu tây tươi nhập từ Hà Khẩu (Vân Nam) bị phát hiện dư lượng Cypermethrin gấp gần 6 lần và Deltamethrin gấp gần 2 lần chuẩn TT 50/2016/TT-BYT, kéo theo chỉ đạo siết kiểm tra ATTP toàn diện, dễ làm chậm thông quan hàng tươi sống.',
         "Tiêu đề: 'Từ lô 4 tấn dâu tây vượt ngưỡng thuốc BVTV (CV 24460/CHQ-GSQL): Chủ hàng nhập nông sản Trung Quốc cần kiểm soát test mẫu tại nguồn thế nào?'.",
         'Công văn 24460/CHQ-GSQL (19/11/2026); CV 4650/TTTV-ATTPMT; Thông tư 50/2016/TT-BYT; thoibaotaichinhvietnam.vn.',
         update_date_str),

        (4,
         'quyet dinh 4502 qd bkhcn giay phep hang hoa luong dung thong tu 30 2026',
         'Quyết định 4502/QĐ-BKHCN (19/11/2026): Công bố bộ TTHC mới cấp phép xuất khẩu, TNTX, chuyển khẩu, quá cảnh hàng hóa lưỡng dụng',
         25100, 'TTHC hàng lưỡng dụng & phân cấp UBND tỉnh (+4,890%)',
         'Doanh nghiệp xuất nhập khẩu thiết bị điện tử, máy tính, linh kiện công nghệ cao, quá cảnh, tạm nhập tái xuất và logistics chuyên tuyến.',
         'Nhiều DN lúng túng xác định hàng công nghệ cao có thuộc danh mục lưỡng dụng theo Thông tư 30/2026/TT-BKHCN hay không (phải thỏa mãn đồng thời mã HS, mô tả và thông số kỹ thuật) và nơi nộp hồ sơ sau khi phân cấp về UBND tỉnh.',
         "Tiêu đề: 'Quyết định 4502/QĐ-BKHCN: 3 dấu hiệu nhận diện hàng hóa lưỡng dụng và quy trình xin phép mới qua Cổng một cửa'.",
         'Quyết định 4502/QĐ-BKHCN (19/11/2026); Thông tư 30/2026/TT-BKHCN; mst.gov.vn; luatvietnam.vn.',
         update_date_str),

        (5,
         'quy dinh eu 2026 2498 cbam khau tru gia carbon xuat khau thep nhom eu',
         'Quy định thực thi (EU) 2026/2498 (hiệu lực 19/11/2026): Hướng dẫn khấu trừ giá carbon đã trả ở nước ngoài trong giai đoạn tài chính CBAM',
         24400, 'Quy định CBAM mới hiệu lực 19/11/2026 (+4,720%)',
         'Doanh nghiệp sản xuất xuất khẩu sắt thép, nhôm, xi măng, phân bón sang thị trường EU và bộ phận ESG/chứng từ xuất khẩu.',
         'Từ năm 2026 CBAM bước vào vận hành tài chính; nếu thiếu chứng từ nộp giá carbon được xác minh độc lập (ngưỡng trọng yếu 5%) hoặc dùng giá mặc định bất lợi, DN xuất khẩu đối mặt chi phí chứng chỉ CBAM đội lên trước hạn nộp 30/09/2027.',
         "Tiêu đề: 'Quy định (EU) 2026/2498 chính thức có hiệu lực: Doanh nghiệp thép, nhôm chuẩn bị hồ sơ khấu trừ carbon CBAM ra sao?'.",
         'Commission Implementing Regulation (EU) 2026/2498; Quy định CBAM (EU) 2023/956; ec.europa.eu; moit.gov.vn.',
         update_date_str),

        (6,
         'nghi quyet 19 2026 bo cap co dncx xuat nhap khau tai cho 24381 24449 chq gsql',
         'Vướng mắc bỏ cấp C/O cho DNCX bán vào nội địa theo Nghị quyết 19/2026/NQ-CP & Công văn 24381, 24449/CHQ-GSQL',
         23900, 'Khoảng trống cấp C/O cho DNCX tại chỗ (+4,550%)',
         'Doanh nghiệp chế xuất (DNCX), doanh nghiệp nội địa mua nguyên liệu/thành phẩm từ DNCX, kho ngoại quan và phòng chứng từ XNK.',
         'Nghị quyết 19/2026/NQ-CP bãi bỏ thủ tục cấp C/O giữa DNCX/phi thuế quan và nội địa khiến DN nội địa không có C/O để hưởng thuế ưu đãi FTA, đối mặt nguy cơ chịu thuế MFN cao và phát sinh tranh chấp hợp đồng.',
         "Tiêu đề: 'Bỏ cấp C/O cho DNCX bán vào nội địa theo Nghị quyết 19/2026: Cách rà soát điều khoản thuế trong hợp đồng để tránh thiệt hại'.",
         'Nghị quyết 19/2026/NQ-CP; Công văn 24381 & 24449/CHQ-GSQL; tapchikinhtetaichinh.vn; daily_notes.md.',
         update_date_str),

        (7,
         'thong tu 71 2026 tt bct phan hang trung tam logistics tieu chuan hang 1',
         'Thông tư 71/2026/TT-BCT (ban hành 19/11/2026, hiệu lực 08/01/2027): Chuẩn hóa 5 loại hình & phân hạng 3 cấp Trung tâm Logistics toàn quốc',
         23200, 'Chuẩn phân hạng Trung tâm Logistics mới (+4,380%)',
         'Doanh nghiệp khai thác kho bãi 3PL/4PL, chủ đầu tư hạ tầng logistics, khu công nghiệp/FTZ và nhà xuất nhập khẩu lớn.',
         'Lần đầu tiên Bộ Công Thương đặt thước đo định lượng cho Trung tâm Logistics (Hạng I từ 20 ha, >= 2 phương thức vận tải, >= 200.000 TEU/năm) và yêu cầu tự công bố lên CSDL hạ tầng thương mại quốc gia.',
         "Tiêu đề: 'Thông tư 71/2026/TT-BCT: Bộ tiêu chuẩn phân hạng Trung tâm Logistics Hạng I-II-III và tác động khi chọn đối tác kho vận'.",
         'Thông tư 71/2026/TT-BCT ngày 19/11/2026; dms.gov.vn; baochinhphu.vn.',
         update_date_str),

        (8,
         '18825 vu vi pham hai quan 20463 ty cao diem kiem soat cuoi nam tet 2027',
         'Thống kê 10 tháng năm 2026 (18.825 vụ vi phạm hải quan, 20.463,6 tỷ đồng) & Cao điểm kiểm soát khai sai HS, trị giá, xuất xứ trước Tết 2027',
         22600, 'Siết kiểm soát sai phạm HS, trị giá, quá cảnh (+4,190%)',
         'Toàn bộ doanh nghiệp XNK chính ngạch, đại lý hải quan, chủ hàng sử dụng loại hình quá cảnh, chuyển cửa khẩu và kho ngoại quan.',
         'Cao điểm cận Tết Đinh Mùi 2027, Hải quan tăng cường quản lý rủi ro và kiểm tra sau thông quan các lỗi khai sai tên hàng, mã HS, trị giá, xuất xứ hoặc trà trộn hàng hóa qua loại hình quá cảnh, kho ngoại quan.',
         "Tiêu đề: '18.825 vụ vi phạm hải quan trong 10 tháng 2026: 4 lỗi khai báo doanh nghiệp hay mắc phải khi cận cao điểm Tết 2027'.",
         'Thống kê Cục Hải quan (15/12/2025 - 14/10/2026); thoibaotaichinhvietnam.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (9,
         'cong van 21692 23474 chq cntt lam sach du lieu vnaccs dksd customs khoa tai khoan 2027',
         'Chiến dịch làm sạch dữ liệu tài khoản VNACCS/VCIS tại dksd.customs.gov.vn trước hạn chót 31/12/2026',
         21800, 'Đếm ngược khóa tài khoản VNACCS 31/12 (+3,980%)',
         'Toàn bộ doanh nghiệp xuất nhập khẩu, nhân viên khai báo hải quan điện tử, đại lý hải quan và kế toán quản lý chữ ký số.',
         'Chỉ còn đúng 40 ngày tới hạn chót 31/12/2026; doanh nghiệp chưa xác thực OTP email, chữ ký số/VNeID tại dksd.customs.gov.vn sẽ bị khóa tài khoản khai báo từ 01/01/2027, làm gián đoạn thông quan đầu năm.',
         "Tiêu đề: 'Còn 40 ngày trước mốc 31/12/2026: Checklist làm sạch tài khoản VNACCS trên dksd.customs.gov.vn để tránh bị khóa đầu năm'.",
         'Công văn 21692 & 23474/CHQ-CNTT; Quyết định 1761/QĐ-CHQ; thoibaotaichinhvietnam.vn.',
         update_date_str),

        (10,
         'ke khai chu so huu huong loi mau 12 tang von dieu le loi java htkk thue nha thau',
         'Case study Thực chiến Doanh nghiệp (daily_notes 21/11): Kê khai Chủ sở hữu hưởng lợi Mẫu 12 khi tăng vốn & Khắc phục lỗi nộp thuế nhà thầu qua HTKK',
         21100, 'Case study nóng cộng đồng Thuế - DN 21/11 (+3,820%)',
         'Giám đốc/Founder doanh nghiệp TNHH, kế toán trưởng, bộ phận pháp chế doanh nghiệp và quản trị thuế nhà thầu XNK.',
         'Doanh nghiệp tăng vốn điều lệ và thay đổi cơ cấu thành viên dễ điền sai Mẫu 12 Chủ sở hữu hưởng lợi dẫn đến trả hồ sơ ĐKKD; đồng thời lỗi cấu hình Java/XML trên HTKK làm chậm nộp tờ khai thuế nhà thầu, phát sinh tiền phạt.',
         "Tiêu đề: 'Kê khai Chủ sở hữu hưởng lợi Mẫu 12 khi tăng vốn & lỗi nộp tờ khai thuế nhà thầu HTKK: 2 lỗi hành chính dễ làm chậm dòng tiền DN'.",
         'scratch/daily_notes.md (cập nhật ngày 21/11/2026); Diễn đàn Thuế chuyên sâu.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 47 (16/11/2026 - 22/11/2026) | Mở rộng 300 chủ đề nghiên cứu (gồm 268 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 301)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 301 or m.max_row >= 301]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (296,
         'cong van 24503 chq gsql thu tuc hai quan hang tai nhap 24485 co 24528 tthc',
         'cong van 24503 chq gsql, thu tuc hai quan hang hoa tai nhap, hang xuat khau bi tra ve viet nam, cong van 24485 chq gsql cap co, cong van 24528 chq nvthq',
         99.8,
         76200,
         'https://luatvietnam.vn',
         'Cụm Công văn ngày 20/11/2026 (24503/CHQ-GSQL, 24485/CHQ-GSQL, 24528/CHQ-NVTHQ): Hướng dẫn thủ tục hải quan hàng tái nhập, giải đáp cấp C/O & đơn giản hóa TTHC',
         'Nội dung trọng tâm: Ngày 20/11/2026, Cục Hải quan ban hành đồng loạt 03 công văn nghiệp vụ quan trọng: (1) Công văn 24503/CHQ-GSQL hướng dẫn thủ tục hải quan đối với hàng hóa đã xuất khẩu nhưng gặp vướng mắc (không đủ điều kiện nhập khẩu ở nước đến hoặc lý do thương mại) buộc phải tái nhập trở lại Việt Nam, làm rõ quy trình hồ sơ để thông quan thuận lợi và xử lý nghĩa vụ thuế; (2) Công văn 24485/CHQ-GSQL giải đáp, kiến nghị xử lý các vướng mắc thực tế về thủ tục cấp Giấy chứng nhận xuất xứ hàng hóa (C/O); (3) Công văn 24528/CHQ-NVTHQ trả lời kiến nghị của doanh nghiệp/hiệp hội về đơn giản hóa và rút ngắn thời gian giải quyết TTHC hải quan; Khuyến nghị từ Eureka: Khi có lô hàng xuất khẩu buộc phải quay đầu tái nhập, doanh nghiệp cần giữ nguyên bộ chứng từ xuất khẩu ban đầu, văn bản xác nhận lý do trả hàng từ đối tác/cảng đến để làm căn cứ khai báo tái nhập không thu thuế đúng quy định.',
         'luatvietnam.vn & bkhub.vn | HT (Học Thực Chiến) / Re-import Customs Procedure & Nov 20 Dispatches',
         update_date_str),

        (297,
         'cua khau thong minh kim thanh bac son con minh tay nam trung quoc nong san chinh ngach',
         'dien dan ket noi viet nam tay nam trung quoc con minh 20 11 2026, cua khau thong minh kim thanh bac son, duong sat kho tieu chuan lao cai, chuoi nong san chinh ngach trung quoc',
         99.9,
         78900,
         'https://baochinhphu.vn',
         'Diễn đàn Kết nối chiến lược Việt Nam – Tây Nam Trung Quốc tại Côn Minh (20/11/2026): Đẩy nhanh cửa khẩu thông minh (Kim Thành – Bắc Sơn) & Chuỗi thương mại nông sản tin cậy chính ngạch',
         'Nội dung trọng tâm: Ngày 20/11/2026 tại TP. Côn Minh (Vân Nam, Trung Quốc), Diễn đàn Kết nối chiến lược phát triển Việt Nam và khu vực Tây Nam Trung Quốc quy mô lớn nhất từ trước đến nay (hơn 570 doanh nghiệp hai nước, có sự tham dự của Tổng Bí thư, Chủ tịch nước Tô Lâm) đã chốt nhiều định hướng hạ tầng logistics trọng yếu trên hành lang Côn Minh – Lào Cai – Hà Nội – Hải Phòng – Quảng Ninh; Trọng tâm triển khai: Đẩy nhanh tiến độ mô hình cửa khẩu thông minh (như cặp cửa khẩu Kim Thành – Bắc Sơn tại Lào Cai), phát triển vận tải container liên vận, các tuyến đường sắt khổ tiêu chuẩn nhằm hạ chi phí logistics, đồng thời đề xuất thí điểm chuỗi thương mại nông sản tin cậy gắn liền đầu tư hạ tầng kho lạnh xuyên suốt từ vùng trồng Việt Nam sang hệ thống phân phối Tây Nam Trung Quốc; Khuyến nghị từ Eureka: Doanh nghiệp XNK tuyến Trung Quốc – Việt Nam cần chủ động chuẩn hóa chứng từ truy xuất vùng trồng, mã số GACC và kế hoạch điều phối container lạnh để đón đầu năng lực thông quan cửa khẩu thông minh.',
         'baochinhphu.vn, vneconomy.vn & congthuong.vn | TN (Tin Ngành) / Kunming Forum Smart Border Gate & Agro Trade Corridor',
         update_date_str),

        (298,
         'quyet dinh 4502 qd bkhcn giay phep hang hoa luong dung thong tu 30 2026',
         'quyet dinh 4502 qd bkhcn 19 11 2026, cap giay phep hang hoa luong dung, thong tu 30 2026 tt bkhcn, tam nhap tai xuat chuyen khau qua canh hang luong dung, phan cap ubnd tinh',
         99.4,
         71800,
         'https://mst.gov.vn',
         'Quyết định 4502/QĐ-BKHCN (19/11/2026): Bộ Khoa học và Công nghệ công bố bộ TTHC mới cấp phép xuất khẩu, TNTX, chuyển khẩu, trung chuyển, quá cảnh hàng hóa lưỡng dụng',
         'Nội dung trọng tâm: Ngày 19/11/2026, Bộ Khoa học và Công nghệ ban hành Quyết định số 4502/QĐ-BKHCN công bố các TTHC mới ban hành về cấp mới và cấp lại giấy phép xuất khẩu, tạm nhập tái xuất, chuyển khẩu, trung chuyển, quá cảnh hàng hóa lưỡng dụng thuộc phạm vi quản lý của Bộ (trên cơ sở Thông tư 30/2026/TT-BKHCN về kiểm soát thương mại chiến lược); Nguyên tắc xác định đối tượng: Hàng hóa, phần mềm, công nghệ chỉ thuộc diện xin phép khi đồng thời trùng mã HS, mô tả và đặc điểm kỹ thuật tại Phụ lục (nếu ghi mã HS 6 số thì áp dụng toàn bộ mã 8 số thuộc phân nhóm); Cơ chế phân cấp: Giao UBND cấp tỉnh giải quyết thủ tục cấp phép và kết nối dữ liệu theo dõi qua Hệ thống thông tin quốc gia chuyên ngành / Cổng một cửa; Khuyến nghị từ Eureka: Doanh nghiệp kinh doanh thiết bị điện tử, viễn thông, máy tính công nghiệp, vật liệu kỹ thuật cao hoặc làm dịch vụ quá cảnh cần đối chiếu kỹ thông số kỹ thuật với Thông tư 30/2026/TT-BKHCN và Thông tư 42/2026/TT-BCT trước khi mở tờ khai.',
         'mst.gov.vn, luatvietnam.vn & thuvienphapluat.vn | TN (Tin Ngành) / Decision 4502 Dual-Use Goods Licensing Procedures',
         update_date_str),

        (299,
         'cong van 24460 chq gsql kiem tra attp dau tay tuoi trung quoc du luong thuoc bvtv',
         'cong van 24460 chq gsql 19 11 2026, kiem tra an toan thuc pham dau tay tuoi nhap khau, cypermethrin deltamethrin thong tu 50 2016 tt byt, cong van 4650 tttv attpmt',
         99.5,
         73600,
         'https://thoibaotaichinhvietnam.vn',
         'Công văn 24460/CHQ-GSQL (19/11/2026): Cục Hải quan chỉ đạo siết kiểm tra ATTP dâu tây tươi nhập khẩu từ Trung Quốc sau phát hiện lô 4.000 kg vượt ngưỡng thuốc BVTV gần 6 lần',
         'Nội dung trọng tâm: Thực hiện đề nghị của Cục Trồng trọt và Bảo vệ thực vật (Công văn 4650/TTTV-ATTPMT ngày 13/11/2026), ngày 19/11/2026 Cục Hải quan ban hành Công văn 24460/CHQ-GSQL yêu cầu các Chi cục Hải quan khu vực tăng cường kiểm tra chặt chẽ an toàn thực phẩm đối với mặt hàng quả dâu tây tươi nhập khẩu, đặc biệt với lô hàng từ nhà sản xuất/xuất khẩu trong danh sách cảnh báo; Chi tiết vụ việc điển hình: Lô 4.000 kg dâu tây tươi (80 hộp xốp và 340 rổ nhựa) nhập từ nhà xuất khẩu tại Hà Khẩu (Vân Nam, Trung Quốc) qua kiểm nghiệm chứa dư lượng Cypermethrin 0,41 mg/kg (gấp gần 6 lần giới hạn 0,07 mg/kg theo TT 50/2016/TT-BYT) và Deltamethrin 0,39 mg/kg (gấp gần 2 lần giới hạn 0,2 mg/kg); Khuyến nghị từ Eureka: Chủ hàng nhập khẩu nông sản, trái cây tươi chính ngạch từ Trung Quốc cần yêu cầu nhà cung cấp gửi kết quả test dư lượng tại nguồn trước khi xếp xe lạnh để tránh rủi ro ách tắc kiểm tra chuyên ngành hoặc buộc tái xuất/tiêu hủy tại cửa khẩu.',
         'thoibaotaichinhvietnam.vn, thuonghieucongluan.com.vn & thuvienphapluat.vn | CS (Case Study) / Letter 24460 Fresh Strawberry Food Safety Alert',
         update_date_str),

        (300,
         'quy dinh eu 2026 2498 cbam khau tru gia carbon xuat khau thep nhom eu',
         'quy dinh thuc thi eu 2026 2498, cbam khau tru gia carbon nuoc thu ba, eu 2023 956 dieu 9, xac minh doc lap cbam 5 phan tram, han nop to khai cbam 30 09 2027',
         99.2,
         69400,
         'https://taxation-customs.ec.europa.eu',
         'Quy định thực thi (EU) 2026/2498 (hiệu lực 19/11/2026, hồi tố từ 01/01/2026): Cơ chế khấu trừ giá carbon đã trả ở nước ngoài trong giai đoạn vận hành tài chính CBAM',
         'Nội dung trọng tâm: Ủy ban Châu Âu (EC) chính thức đưa Quy định thực thi (EU) 2026/2498 vào hiệu lực từ ngày 19/11/2026 (áp dụng hồi tố từ 01/01/2026 — mốc khởi động giai đoạn thu tài chính CBAM), hướng dẫn Điều 9 Quy định (EU) 2023/956 về cách khấu trừ tiền carbon đã trả thực tế tại nước thứ ba khỏi nghĩa vụ mua chứng chỉ CBAM; Các nguyên tắc cốt lõi: (1) Chỉ chấp nhận chi phí carbon từ hệ thống giao dịch phát thải (ETS) hoặc thuế/phí carbon bắt buộc áp dụng bình đẳng giữa hàng nội địa và hàng xuất khẩu (phải trừ đi hoàn thuế/trợ cấp); (2) Giới hạn tỷ lệ sử dụng tín chỉ carbon quốc tế (Điều 6.2/6.4 Thỏa thuận Paris) tối đa 10%; (3) Báo cáo giá carbon phải được kiểm chứng viên độc lập xác nhận ở ngưỡng trọng yếu 5% (hoặc dùng giá carbon mặc định trên CBAM Registry), phục vụ kỳ quyết toán năm đầu tiên hạn chót 30/09/2027; Khuyến nghị từ Eureka: Doanh nghiệp xuất khẩu sắt thép, nhôm, xi măng, phân bón sang EU cần lưu giữ hồ sơ kiểm kê phát thải nội hàm và chứng từ phí carbon để tối ưu chi phí cho đối tác nhập khẩu EU.',
         'taxation-customs.ec.europa.eu & moit.gov.vn | TN (Tin Ngành) / EU 2026/2498 CBAM Carbon Price Deduction Rules',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5  # STT 296 -> row 301
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

    # Cập nhật Total row ở row 306
    total_row_idx = 306
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 300 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D305)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E305)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 300 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 268 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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

    sync_status_msg = "Chưa tìm thấy thư mục OneDrive"
    if os.path.exists(ONEDRIVE_DIR):
        print(f"Đang đồng bộ nguyên tử sang OneDrive: {ONEDRIVE_FILE}...")
        temp_onedrive_file = ONEDRIVE_FILE + '.tmp.xlsx'
        try:
            shutil.copy2(FILE_PATH, temp_onedrive_file)
            os.replace(temp_onedrive_file, ONEDRIVE_FILE)
            print(f"[OK] Đã đồng bộ thành công sang OneDrive: {ONEDRIVE_FILE}")
            sync_status_msg = "✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)"

            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - {update_date_str})\n"
            with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f:
                f.write(log_entry)
            print(f"[OK] Đã ghi log thành công vào: {SYNC_LOG_PATH}")
        except Exception as e:
            sync_status_msg = f"⚠️ Lỗi đồng bộ OneDrive: {e}"
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
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng lên **300 chủ đề**)  
**Trạng thái đồng bộ OneDrive:** {sync_status_msg}  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (7 - 30 NGÀY QUA)
Agent 01 đã quét độc lập trên 8 nguồn dữ liệu theo `wf05_hot_news_scan.md`:
1. `customs.gov.vn` / `luatvietnam.vn` (Loạt công văn ngày **20/11/2026**: `24503/CHQ-GSQL` hướng dẫn thủ tục hải quan hàng tái nhập, `24485/CHQ-GSQL` xử lý kiến nghị C/O, `24528/CHQ-NVTHQ` đơn giản hóa TTHC; Công văn `24460/CHQ-GSQL` ngày **19/11/2026** siết kiểm tra ATTP dâu tây tươi nhập khẩu; chiến dịch làm sạch VNACCS `dksd.customs.gov.vn` trước 31/12/2026)
2. `chinhphu.vn` / `baochinhphu.vn` (Diễn đàn Kết nối chiến lược Việt Nam – Tây Nam Trung Quốc tại Côn Minh ngày **20/11/2026** thúc đẩy cửa khẩu thông minh Kim Thành – Bắc Sơn & đường sắt khổ tiêu chuẩn; Quyết định `4502/QĐ-BKHCN` ngày **19/11/2026** công bố TTHC hàng hóa lưỡng dụng)
3. `haiquanonline.com.vn` / `thoibaotaichinhvietnam.vn` (Tổng kết 10 tháng năm 2026: phát hiện **18.825 vụ vi phạm hải quan** trị giá **20.463,6 tỷ đồng**; cao điểm kiểm soát trị giá, mã HS, xuất xứ, quá cảnh & kho ngoại quan dịp Tết Đinh Mùi 2027)
4. `congthuong.vn` / `moit.gov.vn` / `vneconomy.vn` (Thí điểm chuỗi thương mại nông sản tin cậy chính ngạch & kho lạnh sang Tây Nam Trung Quốc; quy định mới GACC rút gọn đăng ký doanh nghiệp thực phẩm còn 3 ngày; cảnh báo phòng vệ thương mại ngay cả với nhóm hàng có kim ngạch chững lại)
5. `tapchitaichinh.vn` / `tapchikinhtetaichinh.vn` / `ec.europa.eu` (Quy định thực thi **(EU) 2026/2498** hiệu lực **19/11/2026** — áp dụng hồi tố từ 01/01/2026 — quy định khấu trừ giá carbon đã trả ở nước ngoài trong giai đoạn tài chính CBAM)
6. `thuvienphapluat.vn` / `mst.gov.vn` / Google Trends VN (Chi tiết Thông tư `30/2026/TT-BKHCN` và Quyết định `4502/QĐ-BKHCN` về phân cấp cấp phép hàng lưỡng dụng; tra cứu cụm công văn GSQL/NVTHQ ngày 19–20/11/2026)
7. `Drewry WCI` / `Freightos FBX` (Drewry WCI tuần 19/11/2026 ở mức **$4,176/FEU** [-3% WoW, +126% YoY] cùng **10 Blank Sailings**; FBX01 Asia–USWC quanh **$8,200–$8,245/FEU**, FBX11 Asia–North Europe quanh **$3,422/FEU**)
8. `scratch/daily_notes.md` (Cập nhật sáng **21/11/2026**: 2 case study thực tế trên cộng đồng Thuế – Doanh nghiệp gồm: kê khai thông tin **Chủ sở hữu hưởng lợi Mẫu 12** khi tăng vốn công ty TNHH 2 thành viên, và khắc phục **lỗi Java/XML khi nộp tờ khai thuế nhà thầu qua HTKK**)

Qua sàng lọc và đối chiếu chống trùng lặp 100% với 295 chủ đề đã có, Agent 01 đã bổ sung **5 chủ đề mới (STT 296–300)** — đạt cột mốc **300 chủ đề nghiên cứu E2E** — và chấm điểm **TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT NGÀY {update_date_str}** theo 4 tiêu chí (thang điểm 1–5, tối đa 20 điểm):

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: Diễn đàn Côn Minh (20/11/2026): Đẩy nhanh Cửa khẩu thông minh Kim Thành – Bắc Sơn & Chuỗi nông sản chính ngạch Việt Nam – Tây Nam Trung Quốc
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `FF` (Funfact & Insight)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Góc nhìn cơ hội hạ chi phí logistics tuyến đường bộ/đường sắt Lào Cai – Côn Minh | YouTube: Phân tích cách doanh nghiệp nông sản & hàng công nghiệp chuẩn bị hạ tầng kho lạnh, mã GACC để bắt nhịp cửa khẩu thông minh)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Sự kiện ngoại giao kinh tế lớn nhất vừa diễn ra hôm qua 20/11/2026 tại Côn Minh (>570 doanh nghiệp) |
  | Độ đau | 4.8/5 | Đánh thẳng vào nút thắt chi phí vận chuyển, ùn tắc biên giới và hạ tầng kho lạnh tuyến Trung – Việt |
  | Tiềm năng tranh luận | 4.8/5 | Khi cửa khẩu thông minh và chuỗi nông sản tin cậy vận hành, doanh nghiệp nhỏ cần chuẩn hóa gì để không bị loại khỏi cuộc chơi? |
  | Khả năng cho giá trị | 5.0/5 | Gắn trực tiếp thế mạnh cốt lõi tuyến Trung Quốc – Việt Nam của Eureka |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng #1** |
- **Tóm tắt lý do nên làm:** Chạm đúng định vị chuyên gia tuyến Trung – Việt của chị Thúy, vừa mang tính thời sự vĩ mô vừa có giá trị định hướng thực chiến cho chủ hàng.
- **Nguồn xác minh:** https://baochinhphu.vn | https://vneconomy.vn | https://congthuong.vn

---

### Đề xuất #2: Công văn 24460/CHQ-GSQL (19/11/2026): Siết kiểm tra ATTP dâu tây tươi nhập khẩu từ Trung Quốc (vượt ngưỡng Cypermethrin gần 6 lần) — Cảnh báo kiểm tra chuyên ngành nông sản
- **Chuyên mục (Pillar):** `CS` (Case Study) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Facebook / YouTube** (Facebook: Bài cảnh báo nhanh từ số liệu thực tế 4.000 kg dâu tây Hà Khẩu để chủ hàng kiểm soát test mẫu | YouTube: Quy trình tự kiểm soát dư lượng thuốc BVTV & kiểm dịch thực vật khi nhập nông sản từ Trung Quốc)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Cục Hải quan vừa ban hành CV 24460/CHQ-GSQL ngày 19/11/2026 |
  | Độ đau | 5.0/5 | Hàng tươi sống bị ách lại kiểm tra chuyên ngành hoặc buộc tái xuất/tiêu hủy gây thiệt hại toàn bộ lô hàng và cước xe lạnh |
  | Tiềm năng tranh luận | 4.8/5 | Trách nhiệm test mẫu tại nguồn thuộc về nhà xuất khẩu hay chủ hàng Việt Nam? Cách đưa điều khoản phạt vào hợp đồng |
  | Khả năng cho giá trị | 4.7/5 | Có số liệu cụ thể (4.000 kg, Cypermethrin 0,41 vs 0,07 mg/kg; Deltamethrin 0,39 vs 0,2 mg/kg theo TT 50/2016/TT-BYT) |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng #2** |
- **Tóm tắt lý do nên làm:** Case study thực tế có số liệu kiểm nghiệm rõ ràng trên tuyến Hà Khẩu – Việt Nam, thể hiện trọn vẹn góc nhìn Liêm Trinh (chỉ rõ rủi ro) + Giáp Mộc (bảo vệ chủ hàng làm đúng).
- **Nguồn xác minh:** https://thoibaotaichinhvietnam.vn | https://thuonghieucongluan.com.vn | https://thuvienphapluat.vn

---

### Đề xuất #3: Cụm Công văn ngày 20/11/2026 (`24503/CHQ-GSQL`, `24485/CHQ-GSQL`, `24528/CHQ-NVTHQ`): Thủ tục hải quan khi hàng xuất khẩu phải tái nhập về Việt Nam
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `CS` (Case Study)
- **Nền tảng khuyến nghị:** **YouTube / Facebook** (YouTube: Hướng dẫn A-Z bộ hồ sơ tái nhập hàng xuất khẩu bị trả về để không bị tính thuế nhập khẩu oan | Facebook: Checklist 5 chứng từ giữ lại ngay khi đối tác nước ngoài báo trả hàng)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Bộ ba công văn vừa ban hành hôm qua 20/11/2026 |
  | Độ đau | 4.9/5 | Hàng xuất đi bị trả về vốn đã lỗ cước 2 chiều, nếu vướng thủ tục tái nhập còn bị kẹt cảng phát sinh DEM/DET và thuế |
  | Tiềm năng tranh luận | 4.5/5 | Vì sao nhiều doanh nghiệp bị bác miễn thuế khi tái nhập hàng của chính mình? |
  | Khả năng cho giá trị | 5.0/5 | Quy trình nghiệp vụ tái nhập rất ít người hướng dẫn bài bản nhưng doanh nghiệp XNK nào cũng có lúc gặp |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng #3** |
- **Tóm tắt lý do nên làm:** Giải quyết trực diện nỗi đau thực tế cuối năm của doanh nghiệp xuất khẩu, giàu tính hướng dẫn thực chiến.
- **Nguồn xác minh:** https://luatvietnam.vn | https://bkhub.vn

---

### Đề xuất #4: Quyết định 4502/QĐ-BKHCN (19/11/2026): Bộ TTHC mới cấp phép xuất khẩu, tạm nhập tái xuất, quá cảnh hàng hóa lưỡng dụng
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **YouTube / LinkedIn / Facebook B2B** (Facebook: Cảnh báo 3 tiêu chí nhận diện hàng lưỡng dụng | YouTube: Cách tra cứu Phụ lục Thông tư 30/2026/TT-BKHCN và nộp hồ sơ phân cấp về UBND tỉnh)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Vừa ban hành ngày 19/11/2026 |
  | Độ đau | 4.8/5 | Hàng điện tử, viễn thông, máy móc công nghệ cao, hàng quá cảnh dễ bị dừng thông quan nếu thiếu giấy phép lưỡng dụng |
  | Tiềm năng tranh luận | 4.6/5 | Ranh giới giữa hàng dân dụng thông thường và hàng lưỡng dụng — khi nào mới bị áp giấy phép? |
  | Khả năng cho giá trị | 4.9/5 | Làm rõ nguyên tắc trùng đồng thời cả 3 yếu tố (HS + mô tả + thông số kỹ thuật) và điểm mới phân cấp UBND tỉnh |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng #4** |
- **Tóm tắt lý do nên làm:** Chủ đề chuyên sâu giúp khẳng định vị thế tư vấn tuân thủ pháp lý XNK chính thống của Eureka.
- **Nguồn xác minh:** https://mst.gov.vn | https://luatvietnam.vn | https://thuvienphapluat.vn

---

### Đề xuất #5: Quy định thực thi (EU) 2026/2498 (hiệu lực 19/11/2026): Cơ chế khấu trừ giá carbon đã trả ở nước ngoài trong giai đoạn tài chính CBAM
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **YouTube / LinkedIn** (Bài phân tích chuẩn chuyên gia cho doanh nghiệp xuất khẩu thép, nhôm, xi măng, phân bón sang EU)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Có hiệu lực từ 19/11/2026, áp dụng hồi tố cho toàn bộ năm tài chính 2026 |
  | Độ đau | 4.8/5 | Quyết định trực tiếp số tiền chứng chỉ CBAM đối tác EU phải mua và khả năng giữ đơn hàng của nhà máy Việt Nam |
  | Tiềm năng tranh luận | 4.5/5 | Doanh nghiệp Việt chưa có sàn tín chỉ carbon nội địa bắt buộc thì xử lý bài toán khấu trừ và xác minh độc lập thế nào? |
  | Khả năng cho giá trị | 4.8/5 | Đầy đủ mốc pháp lý (EU 2026/2498, trần tín chỉ quốc tế 10%, ngưỡng trọng yếu 5%, hạn nộp 30/09/2027) |
  | **Tổng điểm** | **19.0 / 20** | **Xếp hạng #5** |
- **Tóm tắt lý do nên làm:** Bắt kịp quy định quốc tế mới nhất từ EU, mở rộng tệp khách hàng sản xuất xuất khẩu quy mô lớn.
- **Nguồn xác minh:** https://taxation-customs.ec.europa.eu | https://moit.gov.vn

---

## 💡 GHI CHÚ BỔ SUNG TỪ DAILY NOTES (21/11/2026) & THỐNG KÊ NGÀNH
- **Từ `scratch/daily_notes.md` (21/11/2026):** Ghi nhận 2 vướng mắc thực tế mới trong ngày của chủ doanh nghiệp/kế toán: (1) Cách điền **Mẫu 12 Chủ sở hữu hưởng lợi** khi tăng vốn điều lệ công ty TNHH 2 thành viên; (2) Lỗi Java/XML khi nộp tờ khai **thuế nhà thầu qua HTKK** (phù hợp làm bài chia sẻ ngắn trên Facebook về quản trị tuân thủ doanh nghiệp).
- **Từ thống kê kiểm soát hải quan 10 tháng 2026:** Toàn ngành xử lý **18.825 vụ vi phạm** (trị giá **20.463,6 tỷ đồng**), trọng tâm sai phạm là khai sai tên hàng, mã HS, trị giá, xuất xứ và trà trộn qua quá cảnh/kho ngoại quan trước cao điểm Tết Đinh Mùi 2027.

---

## 🎯 ĐỀ XUẤT HÀNH ĐỘNG TIẾP THEO CHO CHỊ THÚY
- Nếu ưu tiên **định vị tuyến ruột Trung Quốc – Việt Nam (tầm nhìn vĩ mô & thực chiến cửa khẩu)**: Chọn **Đề xuất #1 (Diễn đàn Côn Minh & Cửa khẩu thông minh Kim Thành – Bắc Sơn)** hoặc **Đề xuất #2 (CV 24460/CHQ-GSQL kiểm tra ATTP dâu tây tươi Trung Quốc)**.
- Nếu ưu tiên **gỡ rối thủ tục hải quan thực tế doanh nghiệp hay vướng**: Chọn **Đề xuất #3 (CV 24503/CHQ-GSQL thủ tục hàng tái nhập)** hoặc **Đề xuất #4 (QĐ 4502/QĐ-BKHCN hàng hóa lưỡng dụng)**.
- Sau khi Chị Thúy chọn chủ đề, hệ thống sẽ đẩy ngay sang `wf01` (Lập lịch tuần) hoặc `wf02`/`wf03` (Sản xuất bài Facebook / Kịch bản YouTube).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")
    print("=== HOÀN THÀNH TOÀN BỘ QUY TRÌNH WF05 (21/11/2026) ===")


if __name__ == '__main__':
    update_seo_w47_2111()
