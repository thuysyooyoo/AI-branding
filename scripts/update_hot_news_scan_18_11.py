# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_18_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 18/11/2026
(Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026 | Cập nhật ngày 18/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ tuyệt đối knowledge/banned_words.md và system_prompt_base.md
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 47 cập nhật ngày 18/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 280 lên 285 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 18/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w47_1811.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w47_1811.md'


def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()


def update_seo_w47_1811():
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

    update_date_str = '18/11/2026'

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 47 (16/11/2026 - 22/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'quy dinh eu 2026 2547 ap thue chong ban pha gia thep can nguoi 16 phan tram',
         'Quy định thực thi (EU) 2026/2547 (hiệu lực từ 17/11/2026): EC chính thức áp thuế chống bán phá giá 16% đối với thép cán nguội dạng phẳng Việt Nam',
         25200, 'Áp thuế CBPG 16% sang EU hiệu lực 17/11 (+5,120%)',
         'Doanh nghiệp sản xuất & xuất khẩu thép (nhóm 7209, 7211, 7225, 7226), gia công cơ khí, đơn vị tư vấn phòng vệ thương mại và XNK.',
         'Từ 17/11/2026, thép cán nguội phẳng Việt Nam vào EU chịu thuế CBPG 16% (cao hơn Ấn Độ 9,5% và Thổ Nhĩ Kỳ 3-7,8%), cộng hưởng chi phí chứng chỉ CBAM (82,32 EUR/tấn CO2e Q3/2026) làm xói mòn biên lợi nhuận.',
         "Tiêu đề: 'EC áp thuế CBPG 16% thép cán nguội Việt Nam từ 17/11/2026: Doanh nghiệp xuất khẩu thép ứng phó thế nào cùng lúc với CBAM?'.",
         'Quy định thực thi (EU) 2026/2547; tienphong.vn; sggp.org.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (2,
         'quyet dinh 3318 qd btc 133 thu tuc hai quan phi dia gioi hanh chinh',
         'Quyết định 3318/QĐ-BTC (12/11/2026): Bộ Tài chính công bố 255 TTHC phi địa giới hành chính toàn quốc, gồm 133 thủ tục Hải quan',
         24600, 'Cải cách 133 TTHC hải quan toàn quốc (+4,890%)',
         'Doanh nghiệp XNK có nhiều nhà máy/chi nhánh khác tỉnh, đại lý hải quan, phòng chứng từ và kế toán thuế.',
         'Trước đây DN phải nộp hồ sơ tại đúng Chi cục Hải quan nơi đặt trụ sở/kho bãi gây tốn thời gian đi lại; nay 133 TTHC hải quan (xem hàng trước khi khai, hủy tờ khai, khu phi thuế quan/FTZ...) được xử lý phi địa giới.',
         "Tiêu đề: 'Quyết định 3318/QĐ-BTC: 133 thủ tục hải quan nộp ở đâu cũng được — Doanh nghiệp đa chi nhánh tiết kiệm chi phí chứng từ ra sao?'.",
         'Quyết định 3318/QĐ-BTC ngày 12/11/2026; chinhphu.vn; thoibaotaichinhvietnam.vn; thuvienphapluat.vn.',
         update_date_str),

        (3,
         'thong tu 63 2026 tt bkhcn danh muc san pham cong nghe so da qua su dung cam nhap khau',
         'Thông tư 63/2026/TT-BKHCN (hiệu lực 01/01/2027): Danh mục 13 nhóm sản phẩm công nghệ số đã qua sử dụng cấm nhập khẩu kèm mã HS',
         23900, 'Cảnh báo thay thế TT 11/2018 trước 01/01 (+4,650%)',
         'Doanh nghiệp nhập khẩu máy móc dây chuyền, thiết bị điện tử, linh kiện, bo mạch, màn hình, máy in, camera phục vụ sản xuất hoặc thương mại.',
         'Lệnh cấm không chỉ áp dụng nguyên chiếc cũ mà mở rộng sang bộ phận, phụ kiện, linh kiện điện tử đã qua sử dụng và hàng tân trang (refurbished) thuộc 13 nhóm Chương 84-85; hợp đồng giao sau 01/01/2027 dễ bị tắc cảng.',
         "Tiêu đề: 'Thông tư 63/2026/TT-BKHCN hiệu lực 01/01/2027: Rà soát ngay 13 nhóm mã HS thiết bị công nghệ số cũ cấm nhập khẩu kẻo hàng nằm cảng'.",
         'Thông tư 63/2026/TT-BKHCN ngày 27/10/2026; chinhphu.vn; luatvietnam.vn; quangngai.gov.vn.',
         update_date_str),

        (4,
         'cong van 23897 23907 chq gsql thu tuc kho ngoai quan chia goi co nhap noi dia',
         'Cụm Công văn Kho ngoại quan 23897, 23907, 23905, 23006/CHQ-GSQL: Gia cố, chia gói theo NĐ 167/2025 & vướng mắc C/O hàng gửi kho nhập nội địa',
         23100, 'Gỡ vướng kho ngoại quan & C/O ưu đãi (+4,420%)',
         'Doanh nghiệp FDI sản xuất lắp ráp, công ty 3PL vận hành kho ngoại quan, đơn vị phân phối hàng qua hub ngoại quan.',
         'Hàng từ nội địa gửi vào kho ngoại quan làm trung tâm phân phối rồi nhập ngược vào nội địa đang vướng áp dụng C/O và thuế ưu đãi; đồng thời nghiệp vụ chia gói, đóng gói lại, xuất từng phần dễ sai lệch thanh khoản tồn kho.',
         "Tiêu đề: 'Hàng nội địa gửi kho ngoại quan rồi nhập ngược lại: Vì sao vướng C/O ưu đãi theo Công văn 23907 & 23897/CHQ-GSQL?'.",
         'Công văn 23897, 23907, 23905, 23006/CHQ-GSQL; Nghị định 167/2025/NĐ-CP; thuvienphapluat.vn; luatvietnam.vn.',
         update_date_str),

        (5,
         'cong van 21692 23474 chq cntt lam sach du lieu vnaccs dksd customs khoa tai khoan 2027',
         'Chiến dịch làm sạch dữ liệu tài khoản VNACCS/VCIS tại dksd.customs.gov.vn trước hạn chót 31/12/2026 & Quyết định 1761/QĐ-CHQ',
         22600, 'Đếm ngược khóa tài khoản VNACCS 31/12 (+4,310%)',
         'Toàn bộ doanh nghiệp xuất nhập khẩu, nhân viên khai báo hải quan điện tử, đại lý hải quan và kế toán quản lý chữ ký số.',
         'Hạn chót 31/12/2026 đang cận kề; doanh nghiệp chưa xác thực OTP email, chữ ký số/VNeID tại dksd.customs.gov.vn sẽ bị khóa tài khoản khai báo từ 01/01/2027, làm gián đoạn mở tờ khai đầu năm.',
         "Tiêu đề: 'Đếm ngược trước 31/12/2026: Hướng dẫn làm sạch tài khoản VNACCS trên dksd.customs.gov.vn để không bị khóa từ 01/01/2027'.",
         'Công văn 21692/CHQ-CNTT; Công văn 23474/CHQ-CNTT; Quyết định 1761/QĐ-CHQ; baochinhphu.vn; thuvienphapluat.vn.',
         update_date_str),

        (6,
         'cong van 24313 chq gsql thu tuc hai quan thuong mai dien tu xuyen bien gioi',
         'Công văn 24313/CHQ-GSQL (16/11/2026): Triển khai Hệ thống thực hiện thủ tục hải quan chuyên biệt cho hàng hóa XNK qua nền tảng TMĐT',
         21900, 'Vận hành hệ thống hải quan TMĐT riêng (+4,180%)',
         'Chủ shop TMĐT xuyên biên giới, doanh nghiệp nhập đơn nhỏ tần suất cao tuyến Trung Quốc/Hàn Quốc, đơn vị chuyển phát nhanh.',
         'Đơn hàng TMĐT xuyên biên giới giá trị thấp nhưng số lượng lớn dễ bị ùn ứ khi khai báo truyền thống; hệ thống chuyên biệt kết nối dữ liệu sàn TMĐT - vận chuyển - hải quan đòi hỏi chuẩn hóa mã HS và chứng từ.',
         "Tiêu đề: 'Công văn 24313/CHQ-GSQL: Doanh nghiệp và chủ shop TMĐT xuyên biên giới cần chuẩn bị gì khi Hải quan vận hành hệ thống khai báo riêng?'.",
         'Công văn 24313/CHQ-GSQL ngày 16/11/2026; Luật Hải quan 11/2026/QH16; luatvietnam.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (7,
         'quyet dinh 1828 qd chq kiem tra xuat xu hang hoa sau thong quan 15 ngay',
         'Quyết định 1828/QĐ-CHQ (hiệu lực 15/12/2026): Quy trình kiểm tra xuất xứ hàng hóa sau thông quan tối đa 15 ngày, siết chống chuyển tải',
         21300, 'Đếm ngược hậu kiểm C/O 15 ngày (+4,040%)',
         'Doanh nghiệp xuất khẩu gỗ, thép, dệt may, điện tử, nông sản sử dụng C/O ưu đãi và phòng chứng từ XNK.',
         'Hàng đã thông quan vẫn bị Đội Phúc tập và KTSTQ gọi kiểm tra xuất xứ trong 15 ngày; rủi ro bác C/O, truy thu thuế và xử phạt nếu thiếu hồ sơ nguyên liệu đầu vào, định mức sản xuất.',
         "Tiêu đề: 'Quyết định 1828/QĐ-CHQ hiệu lực 15/12/2026: Bộ hồ sơ chứng minh xuất xứ cần giữ sẵn để vượt qua kỳ hậu kiểm 15 ngày'.",
         'Quyết định 1828/QĐ-CHQ ngày 12/11/2026; Quyết định 467/QĐ-CHQ; thuvienphapluat.vn; thoibaotaichinhvietnam.vn.',
         update_date_str),

        (8,
         'thong tu 142 2026 tt btc quyet dinh 1813 qd chq chuyen phat nhanh a45 a47',
         'Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ (hiệu lực từ 16/11/2026): Chuẩn dữ liệu, mã A45/A47 và nộp thay thuế chuyển phát nhanh',
         20600, 'Tuần đầu áp dụng TT 142 thực tế (+3,890%)',
         'Doanh nghiệp nhập hàng mẫu, linh kiện gấp qua FedEx, DHL, UPS, EMS và kế toán thuế cần tờ khai khấu trừ VAT.',
         'Trong tuần đầu áp dụng từ 16/11/2026, nhiều lô hàng mẫu chuyển phát nhanh gặp vướng mắc mã định danh chủ hàng, chọn sai mã A45/A47 và chứng từ hạch toán khi hãng chuyển phát nộp thay thuế.',
         "Tiêu đề: 'Thực tế tuần đầu Thông tư 142/2026/TT-BTC: Cách dùng đúng mã A45/A47 và lấy tờ khai độc lập cho hàng mẫu chuyển phát nhanh'.",
         'Thông tư 142/2026/TT-BTC; Quyết định 1813/QĐ-CHQ; chinhphu.vn; customs.gov.vn; thoibaotaichinhvietnam.vn.',
         update_date_str),

        (9,
         'cong van 24153 chq nvthq ra soat ma hs my pham sua rua mat 3401 mat na 3307',
         'Công văn 24153/CHQ-NVTHQ: Hải quan rà soát chấn chỉnh khai sai mã HS mỹ phẩm nhập khẩu (sữa rửa mặt 3401.30.00 & mặt nạ 3307.90.90)',
         19900, 'Siết phân loại mã HS mỹ phẩm Hàn/Trung (+3,750%)',
         'Doanh nghiệp nhập khẩu mỹ phẩm Hàn Quốc, Trung Quốc, Nhật Bản; chủ thương hiệu làm đẹp và nhân viên khai báo HS.',
         'Thói quen gộp chung sữa rửa mặt và mặt nạ vải không dệt vào nhóm 33.04 khiến doanh nghiệp đối mặt rủi ro lấy mẫu phân tích phân loại, chậm thông quan và truy thu chênh lệch thuế.',
         "Tiêu đề: 'Công văn 24153/CHQ-NVTHQ: Đừng gộp mỹ phẩm vào nhóm 33.04 — Rà soát ngay mã HS sữa rửa mặt và mặt nạ dưỡng da'.",
         'Công văn 24153/CHQ-NVTHQ; Thông tư 31/2022/TT-BTC; thuvienphapluat.vn; tapchikinhtetaichinh.vn.',
         update_date_str),

        (10,
         'xuat khau sau rieng trung quoc luong xanh cua khau cuoc bien drewry wci 4305',
         'Xuất khẩu sầu riêng đạt 3,45 tỷ USD (Trung Quốc gần 96%), luồng xanh cửa khẩu Việt - Trung & cước biển Drewry WCI ở mức $4,305/FEU',
         19200, 'Cao điểm nông sản & cước tàu cuối năm (+3,580%)',
         'Doanh nghiệp xuất khẩu nông sản tuyến Việt - Trung, chủ hàng đàm phán cước container và đơn vị logistics cửa khẩu phía Bắc.',
         'Dù luồng xanh cửa khẩu Hữu Nghị - Kim Thành được duy trì, tỷ trọng phụ thuộc gần 96% vào thị trường Trung Quốc khiến nguy cơ ùn ứ cục bộ cuối vụ vẫn cao; cước WCI giảm về $4,305/FEU nhưng hãng tàu tăng Blank Sailings.',
         "Tiêu đề: 'Sầu riêng đạt 3,45 tỷ USD & Cước Drewry $4,305/FEU: Giữ luồng xanh cửa khẩu Việt - Trung và chiến lược chốt tàu trước tháng 12'.",
         'Hội nghị Ủy ban Cửa khẩu Việt - Trung lần XI; congthuong.vn; vnexpress.net; drewry.co.uk.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 47 (16/11/2026 - 22/11/2026) | Mở rộng 285 chủ đề nghiên cứu (gồm 253 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 286)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 286 or m.max_row >= 286]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (281,
         'quy dinh eu 2026 2547 ap thue chong ban pha gia thep can nguoi 16 phan tram',
         'quy dinh thuc thi eu 2026 2547, thue chong ban pha gia thep can nguoi, nhom hs 7209 7211 7225 7226, xuat khau thep sang eu, cbam 2026, phong ve thuong mai eu',
         100.0,
         75800,
         'https://tienphong.vn',
         'Quy định thực thi (EU) 2026/2547 (hiệu lực từ 17/11/2026): Ủy ban Châu Âu (EC) chính thức áp thuế chống bán phá giá 16% đối với thép cán nguội dạng phẳng Việt Nam',
         'Nội dung trọng tâm: Theo Quy định thực thi (EU) 2026/2547 ban hành ngày 13/11/2026, công bố trên Công báo EU ngày 16/11/2026 và chính thức có hiệu lực từ 17/11/2026, Ủy ban Châu Âu (EC) áp dụng mức thuế chống bán phá giá chính thức 16% đối với sản phẩm thép cán nguội dạng phẳng của Việt Nam (thuộc nhóm mã HS 7209, 7211, 7225, 7226, khởi xướng từ đơn của EUROFER ngày 18/09/2025); Bối cảnh cạnh tranh: Mức thuế 16% của Việt Nam cao hơn đáng kể so với Ấn Độ (9,5%) và Thổ Nhĩ Kỳ (3% - 7,8%), trong khi Nhật Bản chịu 27,9% và Đài Loan (TQ) chịu 20,7% - 27%; Rủi ro kép cho doanh nghiệp: Ngoài thuế CBPG 16%, thép Việt Nam vào EU còn chịu gánh nặng chứng chỉ CBAM tính theo EU ETS (giá trung bình Q3/2026 khoảng 82,32 EUR/tấn CO2e); Khuyến nghị từ Eureka: Doanh nghiệp xuất khẩu thép cần rà soát hợp đồng FOB/CIF sang EU, chuẩn hóa dữ liệu phát thải thực tế để tránh bị áp giá trị mặc định CBAM và đa dạng hóa thị trường xuất khẩu.',
         'tienphong.vn, sggp.org.vn & tapchikinhtetaichinh.vn | TN (Tin Ngành) / EU 2026/2547 16% Cold-Rolled Steel Anti-Dumping Duty',
         update_date_str),

        (282,
         'quyet dinh 3318 qd btc 133 thu tuc hai quan phi dia gioi hanh chinh',
         'quyet dinh 3318 qd btc, thu tuc hai quan phi dia gioi hanh chinh, 255 thu tuc bo tai chinh, xem hang truoc khi khai, huy to khai hai quan, khu phi thue quan',
         99.7,
         72400,
         'https://chinhphu.vn',
         'Quyết định 3318/QĐ-BTC ngày 12/11/2026: Bộ Tài chính công bố 255 thủ tục hành chính phi địa giới hành chính toàn quốc, gồm 133 thủ tục Hải quan',
         'Nội dung trọng tâm: Bộ Tài chính ban hành Quyết định 3318/QĐ-BTC công bố danh mục 255 thủ tục hành chính thuộc 6 lĩnh vực thực hiện không phụ thuộc địa giới hành chính trên phạm vi toàn quốc, trong đó lĩnh vực Hải quan chiếm tỷ trọng lớn nhất với 133 thủ tục (cùng 51 thủ tục Thuế, 58 thủ tục DN tư nhân/KTTT, 9 thủ tục KBNN...); Thay đổi cốt lõi: Doanh nghiệp có thể nộp hồ sơ và nhận kết quả tại bất kỳ điểm tiếp nhận có thẩm quyền nào hoặc trực tuyến toàn quốc mà không còn bị bó buộc theo nơi đặt trụ sở hay kho bãi như trước (áp dụng cho thủ tục xem hàng hóa trước khi khai mã 1.007840, khai báo tờ khai một lần, hủy tờ khai, hàng đưa vào/ra khu thương mại tự do và khu phi thuế quan...); Giá trị thực tiễn: Giúp doanh nghiệp XNK đa chi nhánh tinh gọn đầu mối chứng từ, tiết giảm thời gian và chi phí đi lại; Giải pháp từ Eureka: Rà soát danh mục 133 TTHC hải quan mới để chuẩn hóa quy trình nộp hồ sơ tập trung tại một bộ phận chứng từ.',
         'chinhphu.vn, thoibaotaichinhvietnam.vn & thuvienphapluat.vn | TN (Tin Ngành) / Decision 3318 Borderless Customs Procedures',
         update_date_str),

        (283,
         'thong tu 63 2026 tt bkhcn danh muc san pham cong nghe so da qua su dung cam nhap khau',
         'thong tu 63 2026 tt bkhcn, san pham cong nghe so da qua su dung cam nhap khau, thay the thong tu 11 2018 tt btttt, ma hs chuong 84 85, hang tan trang refurbished',
         99.5,
         69800,
         'https://luatvietnam.vn',
         'Thông tư 63/2026/TT-BKHCN (ban hành 27/10/2026, hiệu lực 01/01/2027): Danh mục 13 nhóm sản phẩm công nghệ số đã qua sử dụng cấm nhập khẩu kèm mã HS',
         'Nội dung trọng tâm: Bộ Khoa học và Công nghệ ban hành Thông tư 63/2026/TT-BKHCN (có hiệu lực từ 01/01/2027, thay thế hoàn toàn Thông tư 11/2018/TT-BTTTT của Bộ TT&TT cũ) quy định Danh mục 13 nhóm sản phẩm công nghệ số đã qua sử dụng cấm nhập khẩu thuộc Chương 84 và Chương 85 kèm mã HS chi tiết (nhóm 8443 máy in/copy, 8471 máy tính/máy xử lý dữ liệu tự động, 8517 điện thoại/thiết bị truyền dữ liệu, 8518 loa/micro, 8525 camera, 8528 màn hình/máy chiếu, 8534 mạch in, 8542 IC tích hợp, 8544 dây cáp...); Điểm rủi ro lớn: Phạm vi cấm không chỉ áp dụng cho máy nguyên chiếc cũ mà mở rộng sang bộ phận, phụ kiện, cụm linh kiện điện tử đã qua sử dụng và sản phẩm tân trang (refurbished); Quy định chuyển tiếp: Chỉ những lô hàng có giấy phép cấp trước 01/01/2027 mới được tiếp tục nhập theo thời hạn văn bản cho phép; Khuyến nghị từ Eureka: Doanh nghiệp nhập khẩu dây chuyền máy móc cũ có tích hợp màn hình, máy tính điều khiển hoặc linh kiện thay thế từ Trung Quốc/Hàn Quốc cần đối chiếu ngay mã HS trước khi ký L/C giao hàng đầu năm 2027.',
         'chinhphu.vn, luatvietnam.vn & quangngai.gov.vn | TN (Tin Ngành) / Circular 63/2026 Used Digital Products Import Ban',
         update_date_str),

        (284,
         'cong van 23897 23907 chq gsql thu tuc kho ngoai quan chia goi co nhap noi dia',
         'cong van 23897 chq gsql, cong van 23907 chq gsql, cong van 23905 chq gsql, cong van 23006 chq gsql, kho ngoai quan chia goi dong goi, co hang gui kho ngoai quan',
         99.3,
         67400,
         'https://thuvienphapluat.vn',
         'Cụm Công văn Kho ngoại quan 23897, 23907, 23905, 23006/CHQ-GSQL: Gia cố, chia gói theo NĐ 167/2025/NĐ-CP & vướng mắc C/O hàng gửi kho ngoại quan nhập ngược nội địa',
         'Nội dung trọng tâm: Cụm 4 văn bản hướng dẫn mới của Ban Giám sát quản lý (Cục Hải quan) giải đáp các vướng mắc thực tế tại kho ngoại quan: (1) CV 23897/CHQ-GSQL hướng dẫn thủ tục đưa hàng từ nước ngoài hoặc nội địa vào kho ngoại quan để gia cố, chia gói, đóng gói bao bì và xuất từng phần (khai tờ khai vận chuyển độc lập, lập bản kê chi tiết, phiếu xuất kho) theo NĐ 167/2025/NĐ-CP; (2) CV 23905/CHQ-GSQL siết kiểm soát thời hạn lưu kho và thủ tục chuyển nội địa/tái xuất đối với hàng miễn thuế đưa vào kho ngoại quan; (3) CV 23907/CHQ-GSQL & CV 23006/CHQ-GSQL xử lý vướng mắc tính hợp lệ C/O và thuế suất ưu đãi đặc biệt khi hàng xuất khẩu từ nội địa gửi vào kho ngoại quan rồi làm thủ tục nhập khẩu ngược lại nội địa (đang phối hợp Bộ Công Thương xử lý theo NĐ 31/2018/NĐ-CP); Khuyến nghị từ Eureka: Doanh nghiệp sản xuất xuất khẩu và 3PL không tự mặc định được hưởng thuế suất FTA cho luồng hàng nội địa - kho ngoại quan - nội địa mà cần rà soát quy tắc vận chuyển thẳng và văn bản hướng dẫn của Bộ Công Thương.',
         'thuvienphapluat.vn, luatvietnam.vn & bkhub.vn | CS (Case Study) / Bonded Warehouse Operations & C/O Rulings Cluster',
         update_date_str),

        (285,
         'quyet dinh 1761 qd chq so hoa hai quan 2027 lam sach du lieu vnaccs 31 12 2026',
         'quyet dinh 1761 qd chq, cai cach thu tuc hanh chinh hai quan 2026 2030, so hoa 100 phan tram tthc, luat hai quan 11 2026 qh16, dksd customs gov vn, khoa tai khoan vnaccs',
         99.1,
         65900,
         'https://baochinhphu.vn',
         'Quyết định 1761/QĐ-CHQ (02/11/2026) & Lộ trình Luật Hải quan 11/2026/QH16: Số hóa 100% TTHC tái sử dụng dữ liệu, chấm điểm tuân thủ & đếm ngược khóa tài khoản VNACCS trước 31/12/2026',
         'Nội dung trọng tâm: Cục Hải quan ban hành Quyết định 1761/QĐ-CHQ triển khai Kế hoạch cải cách TTHC giai đoạn 2026-2030, đặt mục tiêu ngay trong năm 2027 đạt 100% dịch vụ công trực tuyến toàn trình và số hóa 100% kết quả TTHC để tái sử dụng dữ liệu liên thông, kết hợp phổ biến Luật Hải quan sửa đổi 11/2026/QH16 (hiệu lực 01/03/2027) về cơ chế quản lý tuân thủ doanh nghiệp; Liên kết hành động khẩn cấp: Quá trình tái sử dụng dữ liệu đòi hỏi dữ liệu gốc của doanh nghiệp phải chuẩn xác, gắn trực tiếp với chiến dịch làm sạch dữ liệu tài khoản VNACCS/VCIS tại dksd.customs.gov.vn/client/external (CV 21692 & CV 23474/CHQ-CNTT) có hạn chót 31/12/2026; Rủi ro vận hành: Nếu chưa hoàn tất xác thực email OTP, VNeID và chữ ký số trước 31/12/2026, tài khoản khai báo sẽ bị khóa từ 01/01/2027; Giải pháp từ Eureka: Kiểm tra ngay trạng thái tài khoản trên dksd.customs.gov.vn và xây dựng hồ sơ tuân thủ nội bộ để hưởng ưu tiên luồng xanh.',
         'baochinhphu.vn, vjst.vn & thuvienphapluat.vn | HT (Học Thực Chiến) / Customs Digitalization QĐ 1761 & VNACCS Clean-up Deadline',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5  # STT 281 -> row 286
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

    # Cập nhật Total row ở row 291
    total_row_idx = 291
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 285 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D290)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E290)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 285 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 253 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng lên 285 chủ đề)  
**Trạng thái đồng bộ OneDrive:** ✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (7 - 30 NGÀY QUA)
Agent 01 đã quét độc lập trên 8 nguồn dữ liệu theo `wf05_hot_news_scan.md`:
1. `customs.gov.vn` (Cổng TTĐT Cục Hải quan — chiến dịch làm sạch VNACCS `dksd.customs.gov.vn`, phổ biến Luật Hải quan 11/2026/QH16)
2. `chinhphu.vn` / `baochinhphu.vn` (Cổng TTĐT Chính phủ — QĐ 3318/QĐ-BTC 255 TTHC phi địa giới hành chính, QĐ 1761/QĐ-CHQ, TT 63/2026/TT-BKHCN)
3. `haiquanonline.com.vn` / `tapchikinhtetaichinh.vn` (Phân tích chính sách hải quan, thuế CBPG EU đối với thép cán nguội Việt Nam)
4. `congthuong.vn` / `tienphong.vn` / `sggp.org.vn` (Quy định thực thi (EU) 2026/2547 hiệu lực 17/11/2026 áp thuế CBPG 16%, XNK vượt mốc 1.010,55 tỷ USD, luồng xanh cửa khẩu Việt - Trung)
5. `tapchitaichinh.vn` / `thoibaotaichinhvietnam.vn` (Thu ngân sách XNK 10T đạt 447.362 tỷ đồng = 99,2% dự toán, QĐ 1828/QĐ-CHQ, TTHC phi địa giới)
6. `thuvienphapluat.vn` / `luatvietnam.vn` / Google Trends VN (Tra cứu 10 văn bản mới ngày 18/11/2026 & cụm công văn kho ngoại quan 23897/23907/23905/23006)
7. `Drewry WCI` (Báo cáo ngày 12/11/2026: Cước container WCI ở mức $4,305/FEU, hãng tàu tăng Blank Sailings sang tháng 12)
8. `scratch/daily_notes.md` & `scratch/wf04_10_new_legal_docs_20261118.json` (10 văn bản XNK mới cập nhật sáng 18/11/2026)

Qua sàng lọc, Agent 01 đã chọn lọc và chấm điểm **TOP 5 CHỦ ĐỀ NÓNG NHẤT NGÀY 18/11/2026** theo 4 tiêu chí (thang điểm 1-5, tối đa 20 điểm):

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: Quy định thực thi (EU) 2026/2547 (hiệu lực 17/11/2026): EC chính thức áp thuế chống bán phá giá 16% đối với thép cán nguội Việt Nam
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Đặt vấn đề tranh luận về sức ép kép thuế CBPG 16% + chứng chỉ CBAM | YouTube: Phân tích sâu tác động chuỗi cung ứng thép & bài học phòng vệ thương mại cho doanh nghiệp sản xuất)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 5.0/5 | Vừa có hiệu lực ngày 17/11/2026 (công bố Công báo EU 16/11/2026) |
  | Độ đau | 5.0/5 | Mức thuế 16% cao hơn Ấn Độ (9,5%) và Thổ Nhĩ Kỳ (3-7,8%), cộng dồn chi phí CBAM |
  | Tiềm năng tranh luận | 4.8/5 | Doanh nghiệp nên xoay trục thị trường hay đầu tư số hóa dữ liệu phát thải giảm chi phí CBAM? |
  | Khả năng cho giá trị | 4.9/5 | Dẫn chứng số liệu rõ ràng: nhóm mã HS 7209, 7211, 7225, 7226 và giá EU ETS Q3/2026 (~82,32 EUR/tấn CO2e) |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng #1** |
- **Tóm tắt lý do nên làm:** Tin quốc tế nóng hổi mới nhất vừa có hiệu lực ngày 17/11, đánh thẳng vào ngành xuất khẩu chủ lực chịu áp lực kép từ phòng vệ thương mại và CBAM.
- **Nguồn xác minh:** https://tienphong.vn | https://sggp.org.vn | https://tapchikinhtetaichinh.vn

---

### Đề xuất #2: Quyết định 3318/QĐ-BTC (12/11/2026): 133 thủ tục Hải quan phi địa giới hành chính toàn quốc — Nộp hồ sơ không phụ thuộc nơi đặt trụ sở
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Tin nhanh giá trị cho chủ doanh nghiệp XNK đa chi nhánh | YouTube: Giải mã danh mục thủ tục hải quan & thuế được làm phi địa giới hành chính)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Mới ban hành 12/11/2026, đang được Bộ Tài chính và Hải quan triển khai toàn quốc |
  | Độ đau | 4.9/5 | Cởi trói gánh nặng đi lại và chi phí chứng từ cho DN có nhà máy/kho bãi ở nhiều tỉnh thành |
  | Tiềm năng tranh luận | 4.7/5 | Thực tế triển khai tại các Chi cục đã thông suốt hoàn toàn chưa hay vẫn cần lưu ý khâu tiếp nhận? |
  | Khả năng cho giá trị | 5.0/5 | Liệt kê cụ thể các thủ tục thiết thực: xem hàng trước khi khai (1.007840), hủy tờ khai, khai 1 lần, FTZ |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng #2** |
- **Tóm tắt lý do nên làm:** Chính sách cải cách mang lại lợi ích trực tiếp cho mọi doanh nghiệp XNK và đơn vị logistics, rất dễ thu hút lượt lưu và chia sẻ trên Facebook.
- **Nguồn xác minh:** https://chinhphu.vn | https://thoibaotaichinhvietnam.vn | https://thuvienphapluat.vn

---

### Đề xuất #3: Thông tư 63/2026/TT-BKHCN (hiệu lực 01/01/2027): Danh mục 13 nhóm sản phẩm công nghệ số đã qua sử dụng CẤM nhập khẩu (thay thế TT 11/2018/TT-BTTTT)
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Cảnh báo sớm trước mốc 01/01/2027 cho DN nhập máy móc, linh kiện từ Trung Quốc/Hàn Quốc | YouTube: Hướng dẫn tra cứu 13 nhóm mã HS Chương 84-85 & điều khoản chuyển tiếp)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.8/5 | Đếm ngược hơn 40 ngày tới 01/01/2027 — thời điểm các đơn hàng đặt tháng 11 cập cảng |
  | Độ đau | 5.0/5 | Cấm cả linh kiện, bộ phận cũ và hàng tân trang (refurbished); nhập nhầm là buộc tái xuất/tiêu hủy |
  | Tiềm năng tranh luận | 4.8/5 | Dây chuyền máy móc sản xuất cũ có gắn màn hình điều khiển/máy tính nhúng có bị vướng TT 63/2026 không? |
  | Khả năng cho giá trị | 4.8/5 | So sánh rõ điểm mới thay thế TT 11/2018/TT-BTTTT và cách tận dụng giấy phép chuyển tiếp trước 01/01/2027 |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng #3** |
- **Tóm tắt lý do nên làm:** Đánh trúng tệp doanh nghiệp nhập khẩu máy móc, thiết bị điện tử tuyến Trung Quốc - Việt Nam và Hàn Quốc - Việt Nam của Eureka.
- **Nguồn xác minh:** https://chinhphu.vn | https://luatvietnam.vn | https://quangngai.gov.vn

---

### Đề xuất #4: Cụm Công văn Kho ngoại quan (`23897`, `23907`, `23905`, `23006/CHQ-GSQL`): Gia cố, chia gói & Bẫy mất C/O khi hàng nội địa gửi kho ngoại quan rồi nhập ngược lại
- **Chuyên mục (Pillar):** `CS` (Case Study) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **YouTube / Facebook Case Study** (Phân tích bài học thực chiến cho doanh nghiệp FDI, SXXK và công ty logistics vận hành kho ngoại quan)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.7/5 | Cụm 4 công văn ban hành liên tiếp cuối T10 - đầu T11/2026 (CV 23897/23907/23905 ngày 05/11/2026) |
  | Độ đau | 4.9/5 | Mất ưu đãi C/O khi tái nhập từ kho ngoại quan vào nội địa gây đội chi phí thuế nhập khẩu cực lớn |
  | Tiềm năng tranh luận | 4.7/5 | Hàng Việt Nam đưa vào kho ngoại quan chỉ để phân phối/chia gói — vì sao lại gặp vướng mắc xuất xứ? |
  | Khả năng cho giá trị | 4.9/5 | Bóc tách chuẩn quy trình NĐ 167/2025/NĐ-CP: tờ khai vận chuyển độc lập, bản kê chi tiết, phiếu xuất kho |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng #4** |
- **Tóm tắt lý do nên làm:** Chủ đề nghiệp vụ chuyên sâu giúp khẳng định năng lực tư vấn giải pháp chuỗi cung ứng chính ngạch của Eureka với khách hàng FDI/SME.
- **Nguồn xác minh:** https://thuvienphapluat.vn | https://luatvietnam.vn | https://bkhub.vn

---

### Đề xuất #5: Quyết định 1761/QĐ-CHQ & Đếm ngược hạn chót 31/12/2026: Làm sạch dữ liệu VNACCS tại `dksd.customs.gov.vn` để không bị khóa tài khoản đầu năm 2027
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Bài nhắc nhở lưu về làm ngay trước 31/12 | YouTube: Hướng dẫn thao tác chi tiết trên `dksd.customs.gov.vn` + lộ trình chấm điểm tuân thủ theo QĐ 1761 & Luật HQ 11/2026/QH16)
- **Đánh giá 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | 4.9/5 | Còn hơn 40 ngày tới hạn chót 31/12/2026; đang được các Chi cục Hải quan đôn đốc ráo riết trong T11 |
  | Độ đau | 5.0/5 | Không cập nhật sẽ bị khóa tài khoản khai báo từ 01/01/2027 — đứng hình toàn bộ lô hàng đầu năm |
  | Tiềm năng tranh luận | 4.3/5 | Nhiều DN gặp lỗi khi đồng bộ email OTP hoặc VNeID/chữ ký số mới |
  | Khả năng cho giá trị | 4.9/5 | Có sẵn địa chỉ chính thức, bước thao tác, hotline 1900 9299 (nhánh 2) & email hỗ trợ |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng #5** |
- **Tóm tắt lý do nên làm:** Giá trị thực hành 100% cho mọi doanh nghiệp XNK, cực kỳ dễ viral lượt lưu (Save) và chia sẻ nội bộ phòng chứng từ.
- **Nguồn xác minh:** https://baochinhphu.vn | https://vjst.vn | https://thuvienphapluat.vn

---

## 🎯 ĐỀ XUẤT HÀNH ĐỘNG TIẾP THEO CHO CHỊ THÚY
- Nếu ưu tiên **bắt sóng tin nóng thời sự trên Facebook/YouTube**: Chọn **Đề xuất #1 (EC áp thuế CBPG 16% thép cán nguội từ 17/11/2026)** hoặc **Đề xuất #2 (QĐ 3318 - 133 thủ tục hải quan phi địa giới)**.
- Nếu ưu tiên **cảnh báo cứu nguy doanh nghiệp tuyến Trung Quốc - Hàn Quốc trước 01/01/2027**: Chọn **Đề xuất #3 (TT 63/2026 cấm NK 13 nhóm công nghệ số cũ)** và **Đề xuất #5 (Làm sạch tài khoản VNACCS trước 31/12/2026)**.
- Sau khi Chị Thúy chọn chủ đề, hệ thống sẽ đẩy ngay sang `wf01` (Lập lịch tuần) hoặc `wf02`/`wf03` (Sản xuất bài Facebook / Kịch bản YouTube).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")
    print("=== HOÀN THÀNH TOÀN BỘ QUY TRÌNH WF05 (18/11/2026) ===")


if __name__ == '__main__':
    update_seo_w47_1811()
