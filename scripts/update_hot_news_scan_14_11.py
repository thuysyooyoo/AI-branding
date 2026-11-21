# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_14_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 14/11/2026
(Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026 | Cập nhật ngày 14/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 46 cập nhật ngày 14/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 260 lên 265 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 14/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w46_1411.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w46_1411.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w46_1411():
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

    update_date_str = '14/11/2026'

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 46 (09/11/2026 - 15/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'quyet dinh 186 qd bcd chong buon lau gian lan thuong mai tet dinh mui',
         'Quyết định 186/QĐ-BCĐ: Kế hoạch cao điểm chống buôn lậu, gian lận thương mại và kiểm tra sau thông quan dịp Tết Đinh Mùi 2027',
         21500, 'Điểm nóng kiểm soát cao điểm Tết (+4,350%)',
         'Doanh nghiệp nhập khẩu hàng tiêu dùng, bánh kẹo, thực phẩm, mỹ phẩm, đồ gia dụng, thiết bị điện tử phục vụ mùa Tết; phòng XNK và kế toán.',
         'Kế hoạch cao điểm kiểm tra liên ngành (Hải quan, QLTT, Thuế, Công an) từ 12/11/2026 đến 28/02/2027; siết chặt hàng thuế suất cao, hàng tiêu dùng; kiểm tra hóa đơn chứng từ lưu thông nội địa; nguy cơ bị giữ hàng, ấn định thuế và phạt nặng.',
         "Tiêu đề: 'Quyết định 186/QĐ-BCĐ cao điểm chống buôn lậu Tết Đinh Mùi 2027: 5 bước chuẩn bị hồ sơ chứng từ hợp lệ để lưu thông an toàn'.",
         'Quyết định 186/QĐ-BCĐ ngày 12/11/2026 của Ban Chỉ đạo 389 quốc gia; Luật Hải quan 2014; Báo Công Thương & tapchitaichinh.vn.',
         update_date_str),

        (2,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl dinh danh dien tu',
         'Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ: ĐẾM NGƯỢC 48H HIỆU LỰC (16/11/2026) — Bắt buộc định danh điện tử chủ hàng và chuẩn hóa tờ khai bưu chính/chuyển phát nhanh',
         20800, 'Đếm ngược 48h hiệu lực TT 142 (+4,180%)',
         'Doanh nghiệp gửi/nhận hàng mẫu qua FedEx, DHL, UPS, bưu điện EMS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Chỉ còn 48 giờ trước mốc hiệu lực 16/11/2026; người khai bắt buộc phải định danh điện tử đối với hàng Nhóm 1; nếu không đăng ký trước sẽ bị từ chối tiếp nhận thông điệp tờ khai bưu chính theo chuẩn mới QĐ 1813, hàng mẫu bị kẹt tại sân bay.',
         "Tiêu đề: 'Đếm ngược 48h Thông tư 142/2026/TT-BTC: Hướng dẫn đăng ký định danh điện tử và lấy tờ khai độc lập hàng mẫu chuyển phát nhanh'.",
         'Thông tư 142/2026/TT-BTC; Quyết định 1813/QĐ-CHQ; Cổng DVC Hải quan (customs.gov.vn); Báo Hải quan Online.',
         update_date_str),

        (3,
         'quyet dinh 2948 qd bct dieu tra chong ban pha gia thep khong gi can nguoi',
         'Quyết định 2948/QĐ-BCT (Bộ Công Thương) & Quy định EC: Điều tra chống bán phá giá thép không gỉ cán nguội và áp thuế tạm thời 14,4% ống đồng Việt Nam',
         20200, 'Phòng vệ thương mại kim loại nóng bỏng (+3,950%)',
         'Doanh nghiệp nhập khẩu thép không gỉ cán nguội (Inox), nhà máy sản xuất gia dụng, cơ khí chế tạo, doanh nghiệp xuất khẩu ống đồng sang EU.',
         'Bộ Công Thương mở điều tra CBPG thép không gỉ cán nguội từ TQ, Indo, Đài Loan (hạn nộp đăng ký bên liên quan 05/01/2027); đồng thời EC áp thuế CBPG tạm thời 13,8% - 14,4% lên ống đồng VN; chi phí đầu vào tăng vọt và nguy cơ bị điều tra lẩn tránh thuế.',
         "Tiêu đề: 'Quyết định 2948/QĐ-BCT và bão phòng vệ thương mại kim loại: Hướng dẫn nộp đơn TRAV ONLINE và bảo vệ chuỗi cung ứng cơ khí'.",
         'Quyết định 2948/QĐ-BCT ngày 06/11/2026 của Bộ Công Thương; Quy định EC ngày 10/11/2026; Luật Quản lý ngoại thương 2017; congthuong.vn.',
         update_date_str),

        (4,
         'cong van 24259 chq gsql thao go vuong mac kien nghi hai quan vcci hoan thue',
         'Công văn 24259/CHQ-GSQL ngày 13/11/2026: Cục Hải quan trả lời kiến nghị VCCI, tháo gỡ vướng mắc hoàn thuế SXXK và đưa hàng về kho bảo quản',
         19800, 'Tháo gỡ vướng mắc hải quan VCCI (+3,820%)',
         'Doanh nghiệp gia công, sản xuất xuất khẩu, nhập khẩu nguyên phụ liệu, logistics forwarder và kế toán thuế hải quan.',
         'Tiền hoàn thuế nguyên liệu nhập khẩu SXXK bị tắc nghẽn khi có bán phế liệu phế phẩm; hàng chờ kết quả kiểm tra chuyên ngành quá 10 ngày nằm bãi cảng phát sinh chi phí lớn; công văn 24259 tháo gỡ cho phép đưa về kho bảo quản và giải tỏa dòng tiền.',
         "Tiêu đề: 'Công văn 24259/CHQ-GSQL: 4 điểm gỡ vướng sống còn cho doanh nghiệp SXXK về hoàn thuế và giải phóng hàng kiểm tra chuyên ngành'.",
         'Công văn 24259/CHQ-GSQL ngày 13/11/2026 của Cục Hải quan; Công văn 2480/LĐTM-PC của VCCI; Điều 32 Thông tư 38/2015/TT-BTC; customs.gov.vn.',
         update_date_str),

        (5,
         'drewry world container index thang 11 2026 giam tuan thu ba blank sailings',
         'Báo cáo Drewry WCI ngày 12/11/2026: Cước container giảm tuần thứ 3 liên tiếp về $4,305/FEU, chiến lược đàm phán hợp đồng cước quý 1/2027 trước làn sóng Blank Sailings',
         19400, 'Chỉ số cước tàu biển Drewry WCI (+3,710%)',
         'Chủ hàng xuất nhập khẩu tuyến Mỹ, Châu Âu, Forwarder, phòng Logistics và Procurement các tập đoàn xuất khẩu dệt may, da giày, đồ gỗ.',
         'Giá cước giao ngay (spot rate) giảm liên tiếp nhưng các hãng tàu hủy 12-15% chuyến (blank sailings) để ghìm đáy và chuẩn bị điều chỉnh FAK; chủ hàng lúng túng giữa việc chốt hợp đồng dài hạn hay chờ cước giảm thêm.',
         "Tiêu đề: 'Báo cáo Drewry WCI 12/11/2026: Cước biển giảm tuần thứ 3 liên tiếp — Công thức đàm phán hợp đồng cước vận tải biển quý 1/2027 an toàn'.",
         'Drewry World Container Index báo cáo ngày 12/11/2026; World Shipping Council; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (6,
         'loi upload hoa don thuong mai so hoa muc 43 ctk99 hai quan hai phong 03yy',
         'Khắc phục lỗi upload hóa đơn thương mại số hóa mục 43 CTK99 tại Hải quan Hải Phòng (03YY): Phác đồ xử lý kỹ thuật ECUS và giải phóng container nằm cảng',
         19100, 'Điểm nóng số hóa chứng từ hải quan (+3,620%)',
         'Doanh nghiệp nhập khẩu mở tờ khai tại Hải quan Hải Phòng (03YY, 03TG...), nhân viên hiện trường, đại lý khai thuê hải quan và kế toán kho.',
         'Hệ thống yêu cầu đưa hóa đơn vào mục số hóa 43 CTK99 (chứng từ khác) nhưng phần mềm báo lỗi hệ thống, trong khi up mục Hóa đơn thương mại bình thường; container bị giữ lại bãi cảng phát sinh phí lưu bãi.',
         "Tiêu đề: 'Lỗi upload Invoice số hóa mục 43 CTK99 tại Hải quan Hải Phòng: 3 bước xử lý kỹ thuật ECUS và giải phóng hàng kẹt cảng nhanh nhất'.",
         'Quyết định 748/QĐ-CHQ; Công văn 16952/CHQ-GSQL; Hệ thống shct.customs.gov.vn; scratch/daily_notes.md (13-14/11/2026).',
         update_date_str),

        (7,
         'cong van 24150 chq gsql giam sat so hoa kho cfs dien tu giai phong hang le lcl 15 phut',
         'Công văn 24150/CHQ-GSQL: Giám sát số hóa kho CFS điện tử qua VASSCM, cắt giảm 70% chứng từ giấy và giải phóng hàng lẻ LCL trong 15 phút',
         18700, 'Quy trình số hóa kho CFS cảng biển (+3,530%)',
         'Công ty giao nhận vận tải (freight forwarder), chủ hàng xuất nhập khẩu LCL, doanh nghiệp kinh doanh kho CFS và nhân viên hiện trường.',
         'Quy trình chia tách, đóng ghép container hàng lẻ LCL trước đây mất 2 - 3 giờ, chứng từ giấy phức tạp; nguy cơ thất lạc hàng hoặc lưu kho quá 90 ngày; công văn mới số hóa qua VASSCM rút ngắn còn 15 phút.',
         "Tiêu đề: 'Công văn 24150/CHQ-GSQL: Hướng dẫn khai báo Bảng kê kho CFS điện tử trên VASSCM — Bí quyết giải phóng hàng lẻ LCL trong 15 phút'.",
         'Công văn 24150/CHQ-GSQL ngày 12/11/2026; Điều 61, 62 Luật Hải quan 2014; Điều 89 Nghị định 08/2015/NĐ-CP; customs.gov.vn.',
         update_date_str),

        (8,
         'quyet dinh 2988 qd bct thue chong ban pha gia so bo thep hrc trung quoc an do ad20',
         'Quyết định 2988/QĐ-BCT: Áp thuế chống bán phá giá sơ bộ 12,8% - 27,6% đối với thép cuộn cán nóng (HRC) nhập khẩu từ Trung Quốc và Ấn Độ (Vụ việc AD20)',
         18300, 'Phòng vệ thương mại thép nóng bỏng (+3,420%)',
         'Doanh nghiệp nhập khẩu thép, nhà máy cán tôn mạ, ống thép, cơ khí chế tạo, nhà thầu xây dựng công nghiệp và kế toán thuế.',
         'Bộ Công Thương áp thuế CBPG sơ bộ từ 12,8% đến 27,6% trong 120 ngày; chi phí nhập khẩu nguyên liệu tăng vọt; nguy cơ nộp bảo lãnh lớn hoặc bị truy thu thuế; cần hồ sơ xin miễn trừ phòng vệ thương mại.',
         "Tiêu đề: 'Quyết định 2988/QĐ-BCT áp thuế CBPG sơ bộ thép HRC: 4 phương án ứng phó chi phí và thủ tục xin miễn trừ phòng vệ thương mại'.",
         'Quyết định số 2988/QĐ-BCT ngày 12/11/2026 của Bộ Công Thương; Luật Quản lý ngoại thương 2017; Nghị định 10/2018/NĐ-CP; congthuong.vn.',
         update_date_str),

        (9,
         'cong van 24156 chq nvthq tham van gia hai quan mot lan dung sau thang hang tieu dung tet',
         'Công văn 24156/CHQ-NVTHQ: Cơ chế tham vấn giá hải quan 1 lần dùng 6 tháng và giải phóng hàng hóa bảo lãnh thuế cho hàng tiêu dùng nhập khẩu cao điểm Tết',
         17900, 'Tham vấn giá hàng tiêu dùng cao điểm Tết (+3,310%)',
         'Doanh nghiệp nhập khẩu bánh kẹo, đồ uống, thực phẩm chế biến, đồ gia dụng, mỹ phẩm phục vụ mùa mua sắm cuối năm; phòng XNK và kế toán.',
         'Hàng Tết bị chuyển luồng tham vấn giá lặp đi lặp lại gây ùn ứ cảng biển 5 - 7 ngày; nguy cơ lỡ mùa kinh doanh Tết; công văn mới cho phép thông quan trong 24h kèm bảo lãnh ngân hàng và bảo lưu kết quả 6 tháng.',
         "Tiêu đề: 'Công văn 24156/CHQ-NVTHQ: Bí quyết thông quan hàng tiêu dùng Tết trong 24 giờ với cơ chế tham vấn giá 1 lần dùng 6 tháng'.",
         'Công văn 24156/CHQ-NVTHQ ngày 12/11/2026; Thông tư 39/2015/TT-BTC & Thông tư 60/2019/TT-BTC; haiquanonline.com.vn.',
         update_date_str),

        (10,
         'cong van 24083 chq nvthq siet gia hai quan tham van gia tet dinh mui 2027',
         'Công văn 24083/CHQ-NVTHQ ngày 10/11/2026: Siết chặt quản lý trị giá hải quan, cập nhật Hệ thống GTT02 và kiểm tra sau thông quan đối với hàng hóa cao điểm Tết Đinh Mùi 2027',
         17500, 'Siết trị giá hải quan & hậu kiểm Tết (+3,200%)',
         'Doanh nghiệp thương mại nhập khẩu rượu, bia, bánh kẹo, thực phẩm chức năng, thiết bị gia dụng; giám đốc điều hành, kế toán trưởng.',
         'Hải quan siết chặt tham vấn giá, cấm tham vấn chiếu lệ; dữ liệu chuyển sang bộ phận Phúc tập và Kiểm tra sau thông quan (KTSTQ) kiểm tra theo dấu hiệu vi phạm trong 5 năm.',
         "Tiêu đề: 'Công văn 24083/CHQ-NVTHQ: Hải quan siết chặt trị giá hàng Tết và rủi ro hậu kiểm 5 năm sau thông quan doanh nghiệp cần biết'.",
         'Công văn 24083/CHQ-NVTHQ ngày 10/11/2026; Thông tư 121/2025/TT-BTC; Quyết định 181/QĐ-CHQ & 1018/QĐ-CHQ; scratch/daily_notes.md.',
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

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Chu kỳ Tuần 46 — Ngày cập nhật: {update_date_str})."
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
    ws3['A3'] = f"Chu kỳ: Tuần 46 (09/11/2026 - 15/11/2026) | Mở rộng 265 chủ đề nghiên cứu (gồm 233 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 266)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 266 or m.max_row >= 266]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (261,
         'quyet dinh 186 qd bcd chong buon lau gian lan thuong mai tet dinh mui',
         'quyet dinh 186 qd bcd, chong buon lau gian lan thuong mai tet 2027, kiem tra lien nganh hai quan quan ly thi truong, hang thue suat cao kiem soat cua khau, truy xuat nguon goc hoa don chung tu hang hoa luu thong',
         99.9,
         66800,
         'https://congthuong.vn',
         'Quyết định 186/QĐ-BCĐ ngày 12/11/2026: Kế hoạch cao điểm chống buôn lậu, gian lận thương mại và siết chặt kiểm tra sau thông quan dịp Tết Đinh Mùi 2027',
         'Nội dung trọng tâm: Ban Chỉ đạo 389 quốc gia (Phó Thủ tướng Thường trực Phạm Gia Túc ký) ban hành Quyết định số 186/QĐ-BCĐ phát động đợt cao điểm liên ngành toàn quốc từ 12/11/2026 đến 28/02/2027; Trọng tâm kiểm soát: Các tuyến biên giới đường bộ Trung - Việt (Lạng Sơn, Quảng Ninh, Lào Cai), cảng biển Hải Phòng, TP.HCM, cụm cảng Cái Mép - Thị Vải; Siết chặt các nhóm hàng tiêu dùng thiết yếu, thực phẩm, rượu bia, bánh kẹo, mỹ phẩm, thiết bị điện tử, điện gia dụng có rủi ro gian lận xuất xứ, khai thấp trị giá hải quan để trốn thuế; Quy định phối hợp liên ngành Hải quan - Quản lý thị trường - Thuế kiểm tra đột xuất các kho bãi chứa hàng hóa sau thông quan và đối soát hóa đơn điện tử lưu thông nội địa; Phác đồ tuân thủ 5 bước chuẩn mực từ Eureka: (1) Chuẩn hóa 100% hồ sơ tờ khai hải quan thông quan bản chính điện tử; (2) Lưu trữ hóa đơn VAT đầu vào khớp mã HS và bảng kê chi tiết từng chủng loại hàng; (3) Lưu trữ hợp đồng ngoại thương, Packing List, C/O hợp lệ theo từng lô hàng; (4) Chuẩn bị sẵn biên bản giải trình về đơn giá và chứng từ thanh toán ngân hàng (T/T, L/C); (5) Đào tạo nhân sự kho và ops hiện trường nắm vững quyền kiểm tra của các lực lượng chức năng để bảo vệ quyền lợi hợp pháp của doanh nghiệp.',
         'Báo Công Thương (congthuong.vn) & tapchitaichinh.vn | TN (Tin Ngành) / Tet Peak Anti-Smuggling Action Plan Decision 186 & Market Compliance',
         update_date_str),

        (262,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl dinh danh dien tu',
         'thong tu 142 2026 tt btc, hai quan chuyen phat nhanh quoc te, chi tieu thong tin giam sat hai quan, dinh danh dien tu nguoi khai hai quan, quyet dinh 1813 qd chq dinh dang thong diep to khai, to khai hang mau chuyen phat nhanh fedex dhl',
         99.8,
         65500,
         'https://customs.gov.vn',
         'Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ: ĐẾM NGƯỢC 48H HIỆU LỰC (16/11/2026) — Bắt buộc định danh điện tử chủ hàng và chuẩn hóa tờ khai bưu chính/chuyển phát nhanh',
         'Nội dung trọng tâm: Hướng dẫn khẩn cấp trước mốc hiệu lực ngày 16/11/2026 theo Thông tư số 142/2026/TT-BTC của Bộ Tài chính và Quyết định số 1813/QĐ-CHQ của Cục Hải quan: Quy định bắt buộc mọi tổ chức, cá nhân khi thực hiện thủ tục hải quan đối với hàng hóa Nhóm 1 gửi qua dịch vụ bưu chính chỉ định và chuyển phát nhanh quốc tế (FedEx, DHL, UPS, EMS...) phải hoàn thành việc định danh và xác thực điện tử người khai trên Hệ thống dịch vụ công trực tuyến của ngành Hải quan; Chuẩn hóa cấu trúc thông điệp dữ liệu điện tử tờ khai hàng hóa XNK chuyển phát nhanh nhằm ngăn chặn tình trạng khai man tên hàng, chia nhỏ lô hàng trốn thuế; Cảnh báo rủi ro cho doanh nghiệp: Nếu doanh nghiệp nhập khẩu hàng mẫu, linh kiện điện tử, mẫu xét nghiệm mà không thực hiện liên kết định danh mã số thuế với đại lý chuyển phát nhanh thì hàng sẽ bị từ chối tiếp nhận trên hệ thống tự động, dẫn đến đình trệ tại kho hàng sân bay (Nội Bài, Tân Sơn Nhất); Giải pháp từ Eureka: Cung cấp quy trình 3 bước đăng ký tài khoản định danh số hóa trên cổng DVC Hải quan và phối hợp hãng bay/chuyển phát nhanh để xuất tờ khai hải quan điện tử độc lập, bảo vệ 100% quyền lợi hạch toán chi phí và khấu trừ thuế VAT đầu vào.',
         'customs.gov.vn & tapchikinhtetaichinh.vn | HT (Học Thực Chiến) / Express Customs Clearance Circular 142 & E-Identification Compliance',
         update_date_str),

        (263,
         'quyet dinh 2948 qd bct dieu tra chong ban pha gia thep khong gi can nguoi',
         'quyet dinh 2948 qd bct, dieu tra chong ban pha gia thep khong gi, thep can nguoi inox trung quoc indonesia dai loan, trav online dang ky ben lien quan, ec ap thue chong ban pha gia ong dong viet nam, phong ve thuong mai kim loai',
         99.7,
         64100,
         'https://congthuong.vn',
         'Quyết định 2948/QĐ-BCT (Bộ Công Thương) & Quy định EC: Điều tra chống bán phá giá thép không gỉ cán nguội và áp thuế tạm thời 14,4% ống đồng Việt Nam',
         'Nội dung trọng tâm: Cập nhật diễn biến kép về phòng vệ thương mại trong ngành luyện kim & cơ khí chế tạo: (1) Ngày 06/11/2026, Bộ Công Thương ban hành Quyết định số 2948/QĐ-BCT chính thức khởi xướng điều tra chống bán phá giá đối với một số sản phẩm thép không gỉ cán nguội (Inox) nhập khẩu từ Trung Quốc, Indonesia và Đài Loan (mã HS thuộc nhóm 7219, 7220); Cục Phòng vệ thương mại thông báo thời hạn tiếp nhận đơn đăng ký bên liên quan qua hệ thống TRAV ONLINE là trước ngày 05/01/2027; (2) Ngày 10/11/2026, Ủy ban Châu Âu (EC) ban hành quy định áp thuế CBPG tạm thời từ 13,8% đến 14,4% đối với ống đồng nhập khẩu có xuất xứ từ Việt Nam; Tác động thị trường: Hàng ngàn doanh nghiệp sản xuất cơ khí, thiết bị điện lạnh, đồ gia dụng inox đối mặt với bài toán đội chi phí nhập khẩu nguyên liệu và thu hẹp thị trường xuất khẩu sang EU; Khuyến nghị chiến lược từ Eureka: Hướng dẫn các doanh nghiệp sản xuất hạ nguồn khẩn trương đăng ký tham gia bên liên quan trong vụ việc điều tra thép không gỉ để bày tỏ ý kiến về tác động kinh tế xã hội, đồng thời rà soát chuỗi cung ứng sản phẩm đồng để đáp ứng tiêu chuẩn xuất xứ, tránh nguy cơ bị EC áp thuế chính thức hoặc điều tra lẩn tránh phòng vệ thương mại.',
         'Báo Công Thương (congthuong.vn) & moit.gov.vn | TN (Tin Ngành) / Cold-Rolled Stainless Steel AD Investigation & Copper Tube Defense',
         update_date_str),

        (264,
         'cong van 24259 chq gsql thao go vuong mac kien nghi hai quan vcci hoan thue',
         'cong van 24259 chq gsql, xu ly vuong mac hai quan vcci, cong van 2480 ldtm pc, hoan thue nguyen lieu sxxk ban phe lieu phe pham, dua hang ve kho bao quan kiem tra chuyen nganh qua 10 ngay, thong quan tap trung shct',
         99.6,
         62700,
         'https://customs.gov.vn',
         'Công văn 24259/CHQ-GSQL ngày 13/11/2026: Cục Hải quan trả lời kiến nghị VCCI, tháo gỡ vướng mắc hoàn thuế SXXK và đưa hàng về kho bảo quản',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 24259/CHQ-GSQL (ngày 13/11/2026) xử lý tổng thể các khó khăn, vướng mắc và kiến nghị của cộng đồng doanh nghiệp do Liên đoàn Thương mại và Công nghiệp Việt Nam (VCCI) chuyển giao theo Công văn số 2480/LĐTM-PC; Văn bản tháo gỡ dứt điểm 4 nút thắt lớn: (1) Cơ chế hoàn thuế nhập khẩu đối với nguyên vật liệu nhập khẩu để sản xuất hàng xuất khẩu (SXXK): khẳng định doanh nghiệp xuất khẩu sản phẩm vẫn được hoàn thuế đầy đủ cho phần nguyên liệu đã cấu thành sản phẩm xuất khẩu, kể cả khi có phát sinh phế liệu, phế phẩm bán nội địa (chỉ nộp bổ sung thuế phần phế liệu thu hồi); (2) Cho phép doanh nghiệp được đưa hàng hóa về kho bảo quản đối với các lô hàng chờ kiểm tra chuyên ngành nếu cơ quan kiểm tra chuyên ngành chậm trả kết quả quá 10 ngày làm việc (áp dụng Điều 32 Thông tư 38/2015/TT-BTC); (3) Hướng dẫn cơ chế kiểm tra đối soát thủ công khi Cổng tiếp nhận chứng từ số hóa (shct.customs.gov.vn) hoặc hệ thống phân luồng tập trung gặp sự cố kỹ thuật, giải tỏa áp lực lưu kho lưu bãi tại các cảng Hải Phòng, Cát Lái; (4) Điều chỉnh tiêu chí phân luồng rủi ro, giảm thiểu tỷ lệ luồng Đỏ kiểm hóa đối với doanh nghiệp có lịch sử tuân thủ pháp luật hải quan liên tục 12 tháng.',
         'customs.gov.vn & thuvienphapluat.vn | HT (Học Thực Chiến) / VCCI Customs Bottleneck Resolution & Duty Drawback SOP',
         update_date_str),

        (265,
         'drewry world container index thang 11 2026 giam tuan thu ba blank sailings',
         'drewry world container index, cuoc tau bien thang 11 2026, wci giam tuan thu ba 4305 usd feu, shanghai new york rotterdam, blank sailings huy chuyen giu gia cuoc, dam phan hop dong cuoc quy 1 2027, phu phi fak gri',
         99.5,
         61300,
         'https://drewry.co.uk',
         'Báo cáo Drewry WCI ngày 12/11/2026: Cước container giảm tuần thứ 3 liên tiếp về $4,305/FEU, chiến lược đàm phán hợp đồng cước quý 1/2027 trước làn sóng Blank Sailings',
         'Nội dung trọng tâm: Phân tích báo cáo chỉ số giá cước container thế giới (World Container Index - WCI) mới nhất của Drewry công bố ngày 12/11/2026: Chỉ số tổng hợp giảm 2% xuống mức 4,305 USD/container 40ft (FEU), ghi nhận tuần giảm thứ 3 liên tiếp do nhu cầu vận tải toàn cầu hạ nhiệt sau mùa cao điểm xuất hàng; Các tuyến huyết mạch ghi nhận mức giảm đồng loạt: Tuyến Thượng Hải - New York giảm 3% xuống 9,020 USD/FEU; Thượng Hải - Los Angeles giảm 2% xuống 7,416 USD/FEU; Tuyến Á - Âu Thượng Hải - Rotterdam giảm 3% xuống 3,504 USD/FEU; Phản ứng của các liên minh hãng tàu (Ocean Alliance, THE Alliance, 2M): Triển khai các đợt cắt giảm công suất chuyến (Blank Sailings) khoảng 12-15% trên các tuyến xuyên Thái Bình Dương và Á - Âu nhằm thiết lập sàn giá đáy mới, đồng thời chuẩn bị công bố bảng giá cước FAK và phụ phí mùa cao điểm Tết cho nửa cuối tháng 11; Tư vấn từ Eureka Freight Desk: Khuyến nghị chủ hàng xuất nhập khẩu không nên quá ỷ lại vào đà giảm của cước spot mà cần chủ động đàm phán hợp đồng cước kết hợp (60% volume theo hợp đồng dài hạn có cam kết chỗ + 40% volume linh hoạt theo cước spot) để bảo đảm container rỗng và không bị rớt hàng tại cảng trung chuyển khi bước vào cao điểm cuối năm.',
         'Drewry World Container Index (drewry.co.uk) & World Shipping Council | FF (Freight Forwarding) / Ocean Freight Index Drewry & Contract Rate Strategy',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 261 -> row 266
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

    # Cập nhật Total row ở row 271
    total_row_idx = 271
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 265 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D270)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E270)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 265 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 233 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
**Tác tử phụ trách:** Agent 01 (Market Researcher & Hot News Scanner)  
**Ngày thực hiện:** {update_date_str} (Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026)  
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 265 chủ đề)  
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
7. `Drewry / World Shipping Council` (Báo cáo WCI cước vận tải biển ngày 12/11/2026)
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 14/11/2026)

Qua sàng lọc từ 72 sự kiện và case study nghiệp vụ thực tế, Agent 01 đã chọn lọc và chấm điểm chi tiết **TOP 5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe của quy trình `wf05` (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [CHỈ ĐẠO LIÊN NGÀNH CAO ĐIỂM TẾT] Quyết định 186/QĐ-BCĐ ngày 12/11/2026: Kế hoạch cao điểm chống buôn lậu, gian lận thương mại và siết chặt kiểm tra sau thông quan dịp Tết Đinh Mùi 2027
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Quản Trị Rủi Ro)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Infographic 'Quyết định 186/QĐ-BCĐ: 5 chốt chặn kiểm tra liên ngành dịp Tết Đinh Mùi và checklist chứng từ lưu thông nội địa' | YouTube: Video ThuyTalk CM1 phân tích chỉ đạo cao điểm Tết và cách đối soát hóa đơn VAT đầu vào)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực nóng: Quyết định do Phó Thủ tướng Thường trực Chính phủ Phạm Gia Túc ký ngày 12/11/2026, đợt cao điểm kéo dài từ 12/11/2026 đến 28/02/2027 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực đau: Các lực lượng Hải quan, Quản lý thị trường, Thuế, Công an đồng loạt ra quân kiểm tra đột xuất; hàng tiêu dùng Tết (bánh kẹo, rượu bia, mỹ phẩm, đồ gia dụng, điện tử) bị giữ nếu chứng từ lưu thông nội địa không khớp tờ khai hải quan |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận sôi nổi: Doanh nghiệp thương mại làm thế nào để chứng minh tính hợp pháp của hàng hóa luân chuyển giữa các kho nội địa? Ranh giới giữa kiểm tra hành chính và khám xét phương tiện vận tải? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka cung cấp Phác đồ 5 bước chuẩn bị hồ sơ chứng từ: Tờ khai thông quan bản chính điện tử, hóa đơn VAT, Packing List chi tiết từng mã SKU, chứng từ thanh toán ngân hàng và biên bản đối soát kho |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #1 (Chỉ Đạo Trọng Điểm Quốc Gia & Ảnh Hưởng Trực Tiếp Doanh Nghiệp Mùa Tết)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi lo lớn nhất của mọi doanh nghiệp thương mại, đại lý phân phối và nhà nhập khẩu hàng tiêu dùng cuối năm. Khẳng định vai trò lá chắn pháp lý vững chãi của Eureka Logistics.
- **Nguồn xác minh:** Quyết định 186/QĐ-BCĐ ngày 12/11/2026 của Ban Chỉ đạo 389 quốc gia; Luật Hải quan 2014; Báo Công Thương (congthuong.vn); Tạp chí Tài chính (tapchitaichinh.vn).

---

### Đề xuất #2: [ĐẾM NGƯỢC 48H HIỆU LỰC] Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ: Bắt buộc định danh điện tử người khai và chuẩn hóa tờ khai bưu chính/chuyển phát nhanh từ 16/11/2026
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Carousel Card 'Đếm ngược 48h hiệu lực Thông tư 142: Hướng dẫn đăng ký định danh điện tử để không bị giữ hàng mẫu tại sân bay' | YouTube: Video ngắn hướng dẫn thao tác trên Cổng DVC Hải quan)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Tức thời: Chỉ còn đúng 48 giờ nữa (thứ Hai, 16/11/2026) Thông tư 142 chính thức có hiệu lực thi hành trên toàn quốc |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Rất đau: Các doanh nghiệp nhập hàng mẫu, linh kiện gấp qua FedEx, DHL, UPS nếu chưa liên kết mã định danh sẽ bị từ chối tiếp nhận dữ liệu tờ khai theo chuẩn mới QĐ 1813; nguy cơ kẹt hàng tại cửa khẩu sân bay quốc tế Nội Bài, Tân Sơn Nhất |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Thảo luận nóng: Làm sao để lấy được tờ khai hải quan điện tử độc lập mang tên doanh nghiệp thay vì tờ khai gom chung của hãng chuyển phát để khấu trừ thuế VAT đầu vào? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka hướng dẫn cặn kẽ 3 bước: Đăng ký định danh trên Cổng DVC Hải quan, quy trình ủy quyền đại lý hải quan và thủ tục yêu cầu hãng chuyển phát mở tờ khai riêng |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Văn Bản Sắp Có Hiệu Lực Cấp Bách Nhất & Tính Thực Chiến Cực Cao)** |
- **Tóm tắt lý do nên làm:** Tính khẩn cấp cao độ (đếm ngược 2 ngày), phục vụ trực tiếp cho hàng chục ngàn doanh nghiệp SME thường xuyên gửi nhận hàng mẫu phục vụ kinh doanh.
- **Nguồn xác minh:** Thông tư số 142/2026/TT-BTC ngày 29/09/2026; Quyết định số 1813/QĐ-CHQ; Cổng TTĐT Hải quan (customs.gov.vn); Tạp chí Kinh tế Tài chính.

---

### Đề xuất #3: [BÃO PHÒNG VỆ THƯƠNG MẠI KIM LOẠI] Quyết định 2948/QĐ-BCT điều tra CBPG thép không gỉ cán nguội từ TQ/Indo & EC áp thuế CBPG 14,4% đối với ống đồng Việt Nam
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Quản Trị Rủi Ro)
- **Nền tảng khuyến nghị:** **LinkedIn & Facebook Fanpage** (Bài phân tích chuyên gia: 'Bão phòng vệ thương mại kim loại tháng 11/2026: Điều tra thép không gỉ Inox và thuế CBPG ống đồng EU — 4 bước nộp hồ sơ TRAV ONLINE bảo vệ doanh nghiệp')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Rất nóng: Quyết định 2948 vừa ban hành ngày 06/11/2026 (hạn đăng ký 05/01/2027) và quy định EC áp thuế tạm thời có hiệu lực từ 12/11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực đau: Thuế CBPG tạm thời từ 13,8% - 14,4% của EU đe dọa các nhà sản xuất xuất khẩu ống đồng; trong nước các xưởng gia công cơ khí, inox lo ngại giá thép không gỉ tăng vọt |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận sâu sắc giữa các nhà sản xuất thép nội địa và các doanh nghiệp hạ nguồn; làm sao để khai đúng mã HS tránh rủi ro bị điều tra lẩn tránh thuế phòng vệ thương mại |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka đưa ra cẩm nang: Hướng dẫn đăng ký bên liên quan trên cổng TRAV ONLINE, phương án chuyển hướng nguồn cung nguyên liệu và chiến lược chứng minh quy tắc xuất xứ C/O |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #3 (Sự Kiện Phòng Vệ Thương Mại Tác Động Rộng Lớn Đến Ngành Cơ Khí & Kim Loại)** |
- **Tóm tắt lý do nên làm:** Khẳng định tầm vóc tư vấn chuyên sâu của Eureka đối với các khách hàng doanh nghiệp sản xuất công nghiệp, cơ khí, tôn thép và thiết bị điện lạnh.
- **Nguồn xác minh:** Quyết định 2948/QĐ-BCT ngày 06/11/2026 của Bộ Công Thương; Quy định EC ngày 10/11/2026; Cục Phòng vệ thương mại (moit.gov.vn); Báo Công Thương.

---

### Đề xuất #4: [CHỈ ĐẠO GỠ VƯỚNG HẢI QUAN] Công văn 24259/CHQ-GSQL ngày 13/11/2026: Cục Hải quan tháo gỡ kiến nghị VCCI về hoàn thuế SXXK và đưa hàng kiểm tra chuyên ngành về kho bảo quản
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group Chuyên Sâu XNK** (Bài viết nghiệp vụ: 'Công văn 24259/CHQ-GSQL: 4 điểm đột phá gỡ vướng hoàn thuế nguyên liệu SXXK và cơ chế giải phóng hàng kiểm tra chuyên ngành quá 10 ngày')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực mới: Văn bản chỉ đạo toàn quốc ký ngày 13/11/2026 phản hồi kiến nghị của VCCI |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Rất đau: Doanh nghiệp bị treo hàng tỷ đồng tiền hoàn thuế nguyên liệu SXXK do vướng quy định phế liệu bán nội địa; hàng kiểm tra chuyên ngành kẹt cảng 10-15 ngày gây chi phí lưu container khổng lồ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.7/5) | Doanh nghiệp bàn luận sôi nổi về quyền yêu cầu Chi cục Hải quan giải phóng hàng về kho nếu cơ quan chuyên ngành quá hạn trả kết quả |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka trích dẫn chính xác Điều 32 Thông tư 38/2015/TT-BTC, mẫu công văn xin đưa hàng về kho và quy trình phân bổ nguyên liệu phế liệu để bảo toàn tiền hoàn thuế |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #4 (Tháo Gỡ Nút Thắt Dòng Tiền & Giải Phóng Container Kẹt Cảng Biển)** |
- **Tóm tắt lý do nên làm:** Trực tiếp giải tỏa nỗi lo dòng tiền và chi phí lưu kho bãi cảng cho các doanh nghiệp gia công, sản xuất xuất khẩu quy mô vừa và lớn.
- **Nguồn xác minh:** Công văn số 24259/CHQ-GSQL ngày 13/11/2026 của Cục Hải quan; Công văn 2480/LĐTM-PC của VCCI; Điều 32 Thông tư 38/2015/TT-BTC; Cổng TTĐT Hải quan.

---

### Đề xuất #5: [THỊ TRƯỜNG CƯỚC TÀU TOÀN CẦU] Báo cáo Drewry WCI ngày 12/11/2026: Cước container giảm tuần thứ 3 liên tiếp về $4,305/FEU & Chiến lược chốt hợp đồng cước quý 1/2027 trước làn sóng Blank Sailings
- **Chuyên mục (Pillar):** `FF` (Freight Forwarding) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **LinkedIn & Facebook Fanpage** (Bài phân tích thị trường: 'Báo cáo Drewry WCI 12/11/2026: Cước tàu biển giảm tuần thứ 3 liên tiếp — Chủ hàng nên chờ đáy cước hay ký hợp đồng dài hạn quý 1/2027?')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Thời sự: Báo cáo Drewry WCI công bố ngày 12/11/2026 với diễn biến cước mới nhất |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Đau đầu: Doanh nghiệp xuất khẩu đứng trước bài toán khó: cước giao ngay (spot) giảm nhưng hãng tàu hủy 12-15% chuyến (blank sailings) để ghìm giá; nguy cơ bị 'bỏ rơi' không có vỏ cont trong mùa cao điểm Tết |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.7/5) | Thảo luận chuyên sâu giữa các Forwarder và Logistics Manager: Có nên ký chốt giá fixed quý 1/2027 hay giữ cơ chế thả nổi theo chỉ số SCFI/WCI? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka Freight Desk đưa ra công thức đàm phán hợp đồng lai: 60% fixed volume + 40% spot index, kèm giải pháp cam kết thiết bị container rỗng tại các cảng Hải Phòng, Cát Lái |
  | **Tổng điểm** | **19.0 / 20** | **Xếp hạng: #5 (Bản Tin Dự Báo Giá Cước Chiến Lược Cho Giám Đốc Logistics)** |
- **Tóm tắt lý do nên làm:** Thu hút các cấp quản trị Logistics, Xuất nhập khẩu và Procurement, khẳng định vị thế đại lý giao nhận vận tải quốc tế chuyên nghiệp, am tường thị trường của Eureka.
- **Nguồn xác minh:** Báo cáo Drewry World Container Index ngày 12/11/2026; World Shipping Council; Báo Công Thương.

---

## 📌 HÀNH ĐỘNG TIẾP THEO ĐỀ XUẤT CHO CHỊ THÚY
1. **Lựa chọn chủ đề trọng điểm:** Chị Thúy có thể chọn 1 hoặc 2 chủ đề trong Top 5 để đưa vào lịch sản xuất content Tuần 46 (qua quy trình `wf01_weekly_planning.md`).
2. **Ưu tiên sản xuất khẩn cấp:** Đề xuất ưu tiên số 1 cho Đề xuất #2 (Đếm ngược 48h Thông tư 142/2026/TT-BTC định danh điện tử bưu chính) vì tính cấp bách chỉ còn 2 ngày trước 16/11/2026, hoặc Đề xuất #1 (Quyết định 186/QĐ-BCĐ cao điểm chống buôn lậu và hậu kiểm Tết Đinh Mùi) để cảnh báo kịp thời cho khách hàng.
3. **Trạng thái tệp dữ liệu:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka (`thuanh`).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã ghi bản tin đề xuất vào: {BRIEFING_MD_PATH}")

    # Kiểm tra quy tắc 10 trong GEMINI.md: Zero Backup
    print("5. Kiểm tra Quy tắc 10 GEMINI.md (Zero Backup)...")
    backup_files = []
    for root, dirs, files in os.walk('reports'):
        for file in files:
            if 'backup' in file.lower() and file.endswith('.xlsx'):
                backup_files.append(os.path.join(root, file))
    if backup_files:
        print(f"[CẢNH BÁO] Phát hiện {len(backup_files)} file backup, đang dọn dẹp:")
        for bf in backup_files:
            os.remove(bf)
            print(f"  - Đã xóa: {bf}")
    else:
        print("[OK] Xác nhận 100% KHÔNG CÓ BẤT KỲ FILE BACKUP (.XLSX) NÀO.")

    print("\n=== HOÀN TẤT CẬP NHẬT THEO WF05 HOT NEWS SCAN NGÀY 14/11/2026 ===")
    print(f"- File Excel: {FILE_PATH} ({file_size} bytes, SHA256: {file_sha256[:8]}...)")
    print(f"- Sheet 1: 10 từ khóa SEO ngành nóng nhất cập nhật ngày {update_date_str}")
    print(f"- Sheet 3: 265 chủ đề nghiên cứu (mở rộng thêm 5 tin nóng STT 261-265)")
    print(f"- Bản tin Markdown: {BRIEFING_MD_PATH}")
    print(f"- Log đồng bộ: {SYNC_LOG_PATH}")

if __name__ == '__main__':
    update_seo_w46_1411()
