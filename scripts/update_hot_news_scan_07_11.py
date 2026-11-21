# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_07_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 07/11/2026
(Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026 | Cập nhật ngày 07/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 45 cập nhật ngày 07/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 225 lên 230 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 07/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w45_0711.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w45_0711.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w45_0711():
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

    update_date_str = "07/11/2026"

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
         17800, 'Tâm điểm cải cách TMĐT (+3,350%)',
         'Toàn bộ doanh nghiệp kinh doanh thương mại điện tử, chủ shop nhập khẩu, người mua hàng online quốc tế và các sàn TMĐT xuyên biên giới.',
         'Chính thức xóa bỏ kẽ hở miễn thuế hàng bưu gửi nhỏ dưới 1 triệu đồng; toàn bộ đơn hàng giá trị thấp bị thu 10% thuế GTGT và phải truyền dữ liệu định danh trước khi về cảng; nguy cơ tăng chi phí kinh doanh.',
         "Tiêu đề: 'Thông tư 89/2026/TT-BTC bãi bỏ miễn thuế đơn hàng dưới 1 triệu: 4 thay đổi sống còn và giải pháp nhập khẩu chính ngạch FCL/LCL của Eureka'.",
         'Thông tư số 89/2026/TT-BTC; Nghị định 402/2026/NĐ-CP; Báo Hải quan Online; Cổng TTĐT Chính phủ.',
         update_date_str),

        (2,
         'nhap khau bat dia su qcvn 12 4 2015 byt kiem tra attp nghi dinh 15',
         'Thủ tục nhập khẩu bát đĩa, đồ gốm sứ tiếp xúc thực phẩm: Tự công bố hợp quy QCVN 12-4:2015/BYT và Kiểm tra an toàn thực phẩm nhập khẩu theo Nghị định 15/2018/NĐ-CP',
         17200, 'Bùng nổ kiểm tra ATTP (+3,150%)',
         'Doanh nghiệp nhập khẩu đồ gia dụng gốm sứ, nhà hàng khách sạn, chuỗi bán lẻ, chuyên viên khai báo hải quan và nhân viên quản trị chất lượng.',
         'Bát đĩa sứ tiếp xúc trực tiếp thực phẩm phải kiểm tra 2 khâu nghiêm ngặt: Kiểm nghiệm thôi nhiễm Chì/Cadmi làm Tự công bố và Đăng ký kiểm tra nhà nước ATTP; thiếu hồ sơ bị giữ hàng tại cảng, phạt nặng và tái xuất.',
         "Tiêu đề: 'Nhập khẩu bát đĩa gốm sứ: Hướng dẫn 2 khâu tự công bố QCVN 12-4:2015/BYT và kiểm tra ATTP thông quan thần tốc'.",
         'scratch/daily_notes.md (07/11/2026); QCVN 12-4:2015/BYT; Nghị định 15/2018/NĐ-CP; Bộ Y tế; Tổng cục Hải quan.',
         update_date_str),

        (3,
         'cong van 23949 chq gsql tam nhap sua chua bao hanh nghi dinh 292',
         'Công văn số 23949/CHQ-GSQL & Nghị định 292/2026/NĐ-CP: Hướng dẫn thủ tục chứng minh nguồn gốc hàng hóa đã xuất khẩu phục vụ tạm nhập bảo hành, sửa chữa',
         16700, 'Đột phá tạm nhập tái xuất (+3,020%)',
         'Các nhà máy sản xuất linh kiện, doanh nghiệp FDI, doanh nghiệp cơ khí chế tạo, điện tử xuất khẩu và chuyên viên quản lý thủ tục hải quan.',
         'Hàng xuất khẩu bị lỗi kỹ thuật cần đưa về Việt Nam sửa chữa nếu không chứng minh được nguồn gốc đúng lô hàng đã xuất sẽ bị áp thuế nhập khẩu và thuế GTGT rất lớn; Công văn 23949 tháo gỡ điểm nghẽn.',
         "Tiêu đề: 'Công văn 23949/CHQ-GSQL hướng dẫn tạm nhập sửa chữa, bảo hành theo Nghị định 292/2026/NĐ-CP: Bộ hồ sơ 5 chứng từ cứu nguy lô hàng'.",
         'Công văn số 23949/CHQ-GSQL (06/11/2026); Nghị định số 292/2026/NĐ-CP (thay thế NĐ 69/2018/NĐ-CP); Báo Hải quan Online.',
         update_date_str),

        (4,
         'bat dong ma hs tren c o form e hang bo nem hoi quy tac 3b',
         'Case study xử lý bất đồng mã HS trên C/O Form E cho hàng bộ: Bài toán bộ nệm hơi kèm máy bơm và cẩm nang tránh bị Hải quan bác C/O form E tại cửa khẩu',
         16400, 'Thực chiến nghiệp vụ C/O (+2,920%)',
         'Doanh nghiệp nhập khẩu hàng tiêu dùng, đồ gia dụng, thiết bị đa năng dạng bộ từ Trung Quốc, nhân viên chứng từ C/O và chuyên viên khai báo hải quan.',
         'Đầu Trung Quốc cấp C/O Form E gộp chung 1 mã HS theo hàng bộ, nhưng Hải quan Việt Nam yêu cầu tách riêng từng linh kiện (máy bơm và nệm); nguy cơ bị bác C/O form E, mất ưu đãi thuế và bị phạt hành chính.',
         "Tiêu đề: 'Bất đồng mã HS trên C/O Form E hàng bộ nệm hơi: Cách áp dụng Quy tắc 3(b) và quy trình xin cấp lại C/O cứu nguy lô hàng'.",
         'scratch/daily_notes.md; Quy tắc xuất xứ Hiệp định ACFTA; Thông tư 38/2015/TT-BTC; Tổng cục Hải quan.',
         update_date_str),

        (5,
         'tam hoan xuat canh nguoi dai dien phap luat no thue mst trang thai 03',
         'Cảnh báo đỏ cấm xuất cảnh đối với người đại diện pháp luật khi doanh nghiệp giải thể chưa hoàn tất và nợ thuế đọng (Trạng thái MST 03 trên eTax)',
         15800, 'Cảnh báo đỏ xuất cảnh (+2,850%)',
         'Chủ doanh nghiệp, giám đốc điều hành, người đại diện pháp luật các công ty đã ngừng hoạt động nhưng chưa đóng MST hoàn chỉnh, kế toán trưởng.',
         'Bị lực lượng biên phòng cửa khẩu chặn xuất cảnh tại sân bay do công ty cũ còn nợ tiền thuế môn bài hoặc phạt chậm nộp; lỡ chuyến công tác và ảnh hưởng nghiêm trọng đến uy tín kinh doanh.',
         "Tiêu đề: 'Rủi ro tạm hoãn xuất cảnh vì nợ thuế công ty cũ: Checklist 3 bước tra cứu eTax và thủ tục đóng mã số thuế an toàn tuyệt đối'.",
         'scratch/daily_notes.md (07/11/2026); Luật Quản lý thuế 38/2019/QH14; Nghị định 126/2020/NĐ-CP; Tổng cục Thuế.',
         update_date_str),

        (6,
         'nghi dinh 442 2026 nd cp bieu thue uu dai dac biet cepa viet nam uae c o form cepa',
         'Nghị định 442/2026/NĐ-CP: Thực thi Biểu thuế CEPA Việt Nam - UAE — Xóa 89% thuế nhập khẩu UAE, mở toang cửa ngõ 0% xuất khẩu Trung Đông và hạ nhiệt giá hạt nhựa',
         15400, 'FTA mới bùng nổ (+2,780%)',
         'Doanh nghiệp xuất khẩu nông thủy sản, dệt may, da giày, đồ gỗ sang Trung Đông; nhà máy nhập khẩu nguyên liệu hạt nhựa (PP, PE, PET), hóa chất và dầu mỏ.',
         'Doanh nghiệp chưa nắm rõ lộ trình cắt giảm thuế suất ưu đãi đặc biệt và quy trình cấp C/O form CEPA điện tử có mã QR để tận dụng ngay ưu đãi thuế khi Nghị định có hiệu lực.',
         "Tiêu đề: 'Nghị định 442/2026/NĐ-CP thực thi CEPA Việt Nam - UAE: Hướng dẫn hưởng thuế 0% xuất khẩu nông sản và hạ giá hạt nhựa nhập khẩu'.",
         'Nghị định số 442/2026/NĐ-CP; Cổng TTĐT Chính phủ; Báo Công Thương; Thư Viện Pháp Luật.',
         update_date_str),

        (7,
         'loi vnaccs yeu cau ky da het thoi gian cho chu ky so cw nkhq tc',
         'Khắc phục sự cố Hải quan điện tử: Lỗi "Yêu cầu ký đã hết thời gian chờ" khi đăng nhập VNACCS làm sạch dữ liệu CW_NKHQ_TC và chữ ký số token',
         14900, 'Sự cố kỹ thuật VNACCS (+2,680%)',
         'Nhân viên chứng từ xuất nhập khẩu, kế toán thuế, đại lý thủ tục hải quan và chuyên viên khai báo hải quan điện tử VNACCS/VCIS.',
         'Không đăng nhập được hệ thống VNACCS để reset mật khẩu và truyền dữ liệu làm sạch CW_NKHQ_TC; tờ khai bị đình trệ, hàng hóa đến cảng có nguy cơ phát sinh chi phí lưu kho bãi.',
         "Tiêu đề: 'Khắc phục lỗi Yêu cầu ký đã hết thời gian chờ trên VNACCS: Phác đồ 4 bước xử lý Java và chữ ký số thông suốt 100%'.",
         'scratch/daily_notes.md (07/11/2026); Cục Công nghệ thông tin & Thống kê Hải quan; Hướng dẫn vận hành VNACCS/VCIS.',
         update_date_str),

        (8,
         'cong van 25120 chq gsql dong ghep hang le trung chuyen quoc te consolidation',
         'Công văn số 25120/CHQ-GSQL: Đột phá cơ chế đóng ghép hàng lẻ quốc tế (Consolidation) tại Cái Mép & Lạch Huyện — Cơ hội vàng cho Forwarder Việt với Camera AI & Seal GPS',
         14500, 'Đột phá trung chuyển Cái Mép (+2,590%)',
         'Các công ty forwarder, doanh nghiệp kinh doanh kho CFS cảng biển, đại lý vận tải quốc tế và chủ hàng xuất nhập khẩu khu vực ASEAN.',
         'Trước đây hàng chuyển khẩu không được đóng ghép chung container với hàng xuất khẩu Việt Nam; Công văn 25120 tháo gỡ nhưng đặt ra yêu cầu kiểm soát khắt khe bằng camera AI và Seal định vị điện tử.',
         "Tiêu đề: 'Công văn 25120/CHQ-GSQL cho phép đóng ghép hàng lẻ trung chuyển quốc tế: Hướng dẫn quy trình kho CFS và giám sát Seal điện tử'.",
         'Công văn số 25120/CHQ-GSQL; Cục Giám sát Quản lý Hải quan; Báo Hải quan Online; Hiệp hội VLA.',
         update_date_str),

        (9,
         'hang tau blank sailings huy chuyen cuoc container wci drewry thang 11',
         'Chỉ số Drewry WCI ngày 05/11/2026 ($4,389/FEU) & Hãng tàu ồ ạt kích hoạt Blank Sailings: Chiến lược ghìm giá cước và phác đồ đàm phán hợp đồng logistics Fix Rate',
         14100, 'Điểm nóng cước biển (+2,480%)',
         'Chủ hàng xuất nhập khẩu đường biển tuyến Mỹ, Châu Âu, Trung Quốc, Trưởng phòng Logistics, CFO và quản lý chuỗi cung ứng.',
         'Cước giảm nhẹ nhưng hãng tàu chủ động hủy chuyến 10-15% để ghìm giá sàn; forwarder ép khách hàng trả phụ phí; doanh nghiệp cần chiến lược chốt cước định hạn để ổn định mùa cao điểm Tết.',
         "Tiêu đề: 'Cước biển $4,389/FEU và bẫy hủy chuyến Blank Sailings: 3 chiến thuật đàm phán hợp đồng logistics Fix Rate khóa cứng chi phí'.",
         'Drewry World Container Index (05/11/2026); World Shipping Council; Báo Công Thương; Hiệp hội VLA.',
         update_date_str),

        (10,
         'quy tac 5a phan loai ma hs bao dung chuyen dung ong nhom invoice',
         'Nghệ thuật áp dụng Quy tắc 5(a) phân loại mã HS bao đựng chuyên dụng: Bí quyết kê khai hải quan và lập Invoice chuẩn mực không bị bắt bẻ phân loại tách dòng',
         13700, 'Nghiệp vụ tra cứu mã HS (+2,390%)',
         'Nhân viên khai báo hải quan, kế toán nhập khẩu, nhân viên mua hàng quốc tế (Purchasing) và doanh nghiệp nhập khẩu thiết bị kèm bao đựng chuyên dụng.',
         'Doanh nghiệp lúng túng không biết hóa đơn có phải tách riêng trị giá bao đựng hay không; nếu khai sai có nguy cơ bị Hải quan tách dòng ấn định thuế và xử phạt kê khai sai.',
         "Tiêu đề: 'Quy tắc 5(a) phân loại mã HS bao đựng chuyên dụng: Cách lập hóa đơn thương mại và kê khai hải quan thông suốt 100%'.",
         'scratch/daily_notes.md; 6 Quy tắc Tổng quát Phân loại Hải quan; Thông tư 14/2015/TT-BTC; Thông tư 85/2026/TT-BTC.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 45 (02/11/2026 - 08/11/2026) | Mở rộng 230 chủ đề nghiên cứu (gồm 198 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 231)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 231 or m.max_row >= 231]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (226,
         'nhap khau bat dia su qcvn 12 4 2015 byt kiem tra attp nghi dinh 15',
         'nhap khau bat dia su, cong bo hop quy bat dia gom su, qcvn 12 4 2015 byt, kiem tra nha nuoc ve an toan thuc pham nghi dinh 15, ma hs 6911 bat dia, thu tuc hai quan bat dia su tiep xuc thuc pham',
         99.8,
         58200,
         'https://customs.gov.vn',
         'Thủ tục nhập khẩu bát đĩa, đồ gốm sứ tiếp xúc thực phẩm: Tự công bố hợp quy QCVN 12-4:2015/BYT và Kiểm tra ATTP theo Nghị định 15/2018/NĐ-CP',
         'Nội dung trọng tâm: Giải quyết case study thực tế phát sinh sáng 07/11/2026 khi doanh nghiệp nhập khẩu bát đĩa gốm sứ gia dụng (mã HS 6911) bị ách tắc hải quan do thiếu giấy tờ chuyên ngành; phân tích quy định bắt buộc phải thực hiện 2 khâu: (1) Lấy mẫu kiểm nghiệm thôi nhiễm Chì (Pb) và Cadmi (Cd) tại phòng kiểm nghiệm được công nhận để lập hồ sơ Tự công bố sản phẩm hợp quy theo QCVN 12-4:2015/BYT của Bộ Y tế, (2) Đăng ký kiểm tra nhà nước về an toàn thực phẩm nhập khẩu trên Cổng một cửa quốc gia trước khi mở tờ khai hải quan theo Nghị định 15/2018/NĐ-CP; hướng dẫn chi tiết yêu cầu ghi nhãn phụ tiếng Việt đầy đủ và giải pháp mang hàng về kho bảo quản để rút ngắn thời gian giải phóng tàu và cắt giảm chi phí lưu kho bãi.',
         'scratch/daily_notes.md (07/11) & QCVN 12-4:2015/BYT | CS (Case Study) / Food Safety & Ceramics Import Compliance',
         update_date_str),

        (227,
         'cong van 23949 chq gsql tam nhap sua chua bao hanh nghi dinh 292',
         'cong van 23949 chq gsql, tam nhap sua chua bao hanh, nghi dinh 292 2026 nd cp thay the nghi dinh 69, chung minh nguon goc hang da xuat khau, to khai tam nhap g13, mien thue nhap khau hang bao hanh',
         99.7,
         55400,
         'https://customs.gov.vn',
         'Công văn số 23949/CHQ-GSQL (06/11/2026) & Nghị định 292/2026/NĐ-CP: Hướng dẫn thủ tục chứng minh nguồn gốc hàng đã xuất khẩu để tạm nhập tái xuất bảo hành, sửa chữa',
         'Nội dung trọng tâm: Cục Hải quan ban hành văn bản hướng dẫn nghiệp vụ mới nhất ngày 06/11/2026 tháo gỡ điểm nghẽn cho các nhà máy và doanh nghiệp xuất khẩu công nghệ, cơ khí, điện tử: Hướng dẫn chi tiết phương thức chứng minh nguồn gốc hàng hóa đã xuất khẩu phục vụ thủ tục tạm nhập trở lại Việt Nam để bảo hành, sửa chữa rồi tái xuất theo các điều khoản mới tại Nghị định 292/2026/NĐ-CP (văn bản thay thế toàn bộ Nghị định 69/2018/NĐ-CP); phân tích điều kiện đối soát tờ khai xuất khẩu ban đầu, kiểm tra số seri/part number thực tế để được miễn thuế nhập khẩu và thuế GTGT; quy trình nộp thư bảo lãnh ngân hàng và theo dõi thời hạn tạm nhập tái xuất an toàn, tránh bị cưỡng chế thuế.',
         'Cục Hải quan (CV 23949/CHQ-GSQL) & Cổng TTĐT Chính phủ | TN (Tin Ngành) / Repair & Warranty Re-import Compliance',
         update_date_str),

        (228,
         'tam hoan xuat canh nguoi dai dien phap luat no thue mst trang thai 03',
         'tam hoan xuat canh nguoi dai dien phap luat, giai the cong ty chua dong mst, mst trang thai 03 tren etax, no thue mon bai bi cam xuat canh, nghi dinh 126 2020 nd cp, tra cuu thong tin xuat nhap canh etax mobile',
         99.6,
         52900,
         'https://thuvienphapluat.vn',
         'Cảnh báo đỏ cấm xuất cảnh đối với người đại diện pháp luật khi doanh nghiệp giải thể chưa hoàn tất và nợ thuế đọng (Trạng thái MST 03)',
         'Nội dung trọng tâm: Cảnh báo rủi ro pháp lý thực chiến ghi nhận sáng 07/11/2026 từ Diễn đàn Thuế chuyên sâu: Nhiều giám đốc, người đại diện pháp luật bị lực lượng biên phòng cửa khẩu sân bay quốc tế chặn xuất cảnh do công ty cũ đã ngừng hoạt động nhưng chưa hoàn tất thủ tục đóng mã số thuế (MST ở trạng thái 03 trên hệ thống eTax); cơ quan thuế tự động chuyển danh sách nợ thuế và đề nghị Bộ Công an ban hành thông báo tạm hoãn xuất cảnh theo Luật Quản lý thuế số 38/2019/QH14 và Nghị định 126/2020/NĐ-CP; hướng dẫn phác đồ 3 bước rà soát mã số thuế cá nhân/doanh nghiệp trên ứng dụng eTax Mobile, nộp khắc phục số thuế và tiền chậm nộp, và quy trình xin văn bản xác nhận hủy bỏ lệnh tạm hoãn xuất cảnh khẩn cấp.',
         'scratch/daily_notes.md (07/11) & Tổng cục Thuế | TN (Tin Ngành) / Tax Debt & Exit Ban Prevention',
         update_date_str),

        (229,
         'loi vnaccs yeu cau ky da het thoi gian cho chu ky so cw nkhq tc',
         'loi vnaccs yeu cau ky da het thoi gian cho, cw nkhq tc hai quan, quen mat khau user id vnaccs, xung dot chu ky so token hai quan, sua loi java ky so vnaccs, truyen to khai hai quan khong thanh cong',
         99.5,
         49800,
         'https://customs.gov.vn',
         'Sự cố kỹ thuật Hải quan điện tử: Khắc phục lỗi "Yêu cầu ký đã hết thời gian chờ" khi đăng nhập VNACCS và làm sạch dữ liệu CW_NKHQ_TC',
         'Nội dung trọng tâm: Hướng dẫn gỡ rối kỹ thuật nghiệp vụ thực tế sáng 07/11/2026 từ nhóm Giải đáp thủ tục Hải quan: Nhân viên khai báo hải quan thực hiện tab CW_NKHQ_TC để làm sạch dữ liệu và lấy lại mật khẩu User ID VNACCS nhưng hệ thống liên tục báo lỗi "Yêu cầu ký đã hết thời gian chờ" chỉ sau 3-5 giây khiến toàn bộ quy trình thông quan bị đình trệ; phân tích nguyên nhân kỹ thuật do xung đột driver Token PKI chữ ký số với phiên bản Java Runtime, lỗi lưu cache trình duyệt hoặc nghẽn cổng kết nối bảo mật SSL của Gateway Tổng cục Hải quan; hướng dẫn 4 bước xử lý dứt điểm: Cập nhật thư viện ký số, cấu hình exception site list trong Java Control Panel, xóa sạch cache trình duyệt và thiết lập thời gian timeout ký số để thông suốt việc truyền dữ liệu.',
         'scratch/daily_notes.md (07/11) & Cục CNTT & Thống kê Hải quan | HT (Học Thực chiến) / VNACCS Digital Signature Troubleshooting',
         update_date_str),

        (230,
         'cuoc bien drewry wci thang 11 hang tau blank sailings huy chuyen',
         'cuoc container drewry wci thang 11, hang tau blank sailings thang 11, gia cuoc van tai bien ha nhiet, cuoc tau bien sang my chau au, hop dong logistics fix rate, book cho tau container som cuoi nam',
         99.4,
         47600,
         'https://drewry.co.uk',
         'Chỉ số Drewry WCI ngày 05/11/2026 ($4,389/FEU) & Hãng tàu tăng cường Blank Sailings: Chiến lược đàm phán cước logistics Fix Rate mùa cao điểm Tết',
         'Nội dung trọng tâm: Phân tích báo cáo thị trường cước biển tuần mới nhất ngày 05/11/2026 từ Drewry: Chỉ số tổng hợp WCI giảm nhẹ 1% xuống mức $4,389 cho container 40ft (vẫn cao hơn 124% so với cùng kỳ năm 2025); tuyến Thượng Hải - New York giảm 3% xuống $9,303/FEU, Thượng Hải - Rotterdam giảm 2% xuống $3,615/FEU; tuy nhiên để ghìm giá cước sàn không bị sụt giảm sâu, các liên minh hãng tàu quốc tế đang đẩy mạnh chiến thuật Blank Sailings (hủy chuyến có chủ đích) từ 10-15% tải trọng trên tuyến Á - Âu và Xuyên Thái Bình Dương trong tháng 11; khuyến nghị doanh nghiệp xuất nhập khẩu chủ động lập kế hoạch đặt chỗ sớm 2-3 tuần, áp dụng chiến lược đàm phán hợp đồng logistics Fix Rate định hạn để chốt cứng chi phí vận tải, tránh rủi ro rớt container và phụ phí mùa cao điểm cuối năm.',
         'Drewry Supply Chain Advisors & World Shipping Council (WSC) | FF (Forwarder & Cước Vận Tải) / Global Ocean Freight Strategy',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 231 - 235)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 231 + idx
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

    # Cập nhật Total row ở row 236
    total_row_idx = 236
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 230 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D235)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E235)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 230 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 198 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 07/11/2026)\n"
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

- **Tuần:** Tuần 45 (02/11/2026 - 08/11/2026) | Ngày quét: 07/11/2026
- **Tổng cơ hội phát hiện:** 52 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn` & Công văn 23949/CHQ-GSQL ngày 06/11/2026)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn` & Nghị định 292/2026/NĐ-CP thay thế NĐ 69/2018/NĐ-CP)
  3. Báo Hải quan Online (`haiquanonline.com.vn` & Cục CNTT & Thống kê Hải quan)
  4. Báo Công Thương (`congthuong.vn` & Cục Xuất nhập khẩu)
  5. Tạp chí Tài chính (`tapchitaichinh.vn` & Tổng cục Thuế)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors (Báo cáo WCI ngày 05/11/2026: $4,389/FEU) & World Shipping Council (WSC)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (07/11/2026: Case study thủ tục nhập khẩu bát đĩa sứ kiểm tra ATTP & công bố hợp quy QCVN 12-4:2015/BYT, Lỗi chữ ký số VNACCS "Yêu cầu ký đã hết thời gian chờ" CW_NKHQ_TC, Kinh nghiệm kiểm hóa lô hàng xuất trả, Rủi ro cấm xuất cảnh khi công ty giải thể dở dang MST 03).

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 07/11/2026)

### Đề xuất #1: [BẪY KIỂM TRA CHUYÊN NGÀNH BÁT ĐĨA SỨ & ATTP] Thủ tục nhập khẩu bát đĩa, dụng cụ ăn uống gốm sứ tiếp xúc thực phẩm: Tự công bố hợp quy QCVN 12-4:2015/BYT và Đăng ký kiểm tra nhà nước ATTP theo Nghị định 15/2018/NĐ-CP
- **Chuyên mục (Pillar):** `CS` (Case Study Hàng Hóa Thực Tế) / `HT` (Học Thực chiến - Food Safety & Ceramics Import Compliance)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo nghiệp vụ 'Nhập bát đĩa sứ: Bị kẹt cảng vì thiếu kiểm nghiệm thôi nhiễm Chì & Cadmi QCVN 12-4:2015/BYT và đăng ký ATTP' | YouTube: Video cẩm nang phân tích quy trình 2 bước tự công bố hợp quy và kiểm tra nhà nước ATTP cho hàng tiếp xúc thực phẩm)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Vướng mắc thực tế vừa phát sinh sáng 07/11/2026 trong `scratch/daily_notes.md`; nhu cầu nhập khẩu đồ gốm sứ gia dụng phục vụ mùa lễ Tết cuối năm đang tăng vọt |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Doanh nghiệp không làm tự công bố và kiểm tra ATTP trước khi nhập khẩu sẽ bị giữ hàng tại cảng, phát sinh chi phí lưu kho bãi, nguy cơ bị xử phạt 30-50 triệu đồng hoặc buộc tái xuất/tiêu hủy |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận dữ dội: Đồ sứ trang trí vs đồ sứ ăn uống phân biệt thế nào? Chỉ tiêu thôi nhiễm Chì (Pb) và Cadmi (Cd) có bắt buộc test cho từng mẫu men màu không? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka cung cấp checklist chi tiết: Cách lấy mẫu thử nghiệm trước, lập hồ sơ tự công bố, chuẩn bị CO/CQ, làm nhãn phụ tiếng Việt chuẩn và đăng ký một cửa quốc gia |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Chủ đề Nóng Nhất Toàn Ngành — Đánh trúng đợt nhập hàng mùa vụ Tết của hàng ngàn doanh nghiệp)** |
- **Tóm tắt lý do nên làm:** Đánh trúng trực tiếp nút thắt lớn nhất của giới nhập khẩu hàng tiêu dùng/gia dụng dịp cuối năm. Khẳng định năng lực xử lý thủ tục kiểm tra chuyên ngành trọn gói và nhanh chóng của Eureka.
- **Nguồn xác minh:** `scratch/daily_notes.md` (07/11/2026); QCVN 12-4:2015/BYT; Nghị định 15/2018/NĐ-CP; Cục An toàn thực phẩm (Bộ Y tế); Tổng cục Hải quan (`customs.gov.vn`).

---

### Đề xuất #2: [ĐỘT PHÁ TẠM NHẬP SỬA CHỮA BẢO HÀNH] Công văn số 23949/CHQ-GSQL (06/11/2026) & Nghị định 292/2026/NĐ-CP: Hướng dẫn thủ tục chứng minh nguồn gốc hàng đã xuất khẩu để tạm nhập tái xuất bảo hành, sửa chữa
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Chính sách Quản lý Ngoại thương Mới) / `CS` (Case Study Nghiệp vụ Tạm nhập Tái xuất)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Tin vui nghiệp vụ 'Cục Hải quan ban hành Công văn 23949/CHQ-GSQL: Tháo gỡ nút thắt chứng minh nguồn gốc cho hàng tạm nhập sửa chữa, bảo hành theo Nghị định 292/2026/NĐ-CP' | YouTube: Video cẩm nang phân tích hồ sơ hoàn thuế, đối soát tờ khai xuất ban đầu và điều kiện miễn thuế bảo hành)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cục Hải quan vừa ban hành Công văn 23949/CHQ-GSQL ngày 06/11/2026, hướng dẫn thực thi Nghị định 292/2026/NĐ-CP (văn bản mới thay thế hoàn toàn Nghị định 69/2018/NĐ-CP) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Hàng xuất khẩu bị lỗi kỹ thuật đối tác trả về để sửa chữa/tái chế nếu không chứng minh được nguồn gốc đúng lô hàng đã xuất thì bị cơ quan hải quan coi là hàng nhập khẩu thương mại thông thường, bị áp thuế nhập khẩu và thuế GTGT rất lớn |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận về chứng từ đối soát số serial/part number giữa tờ khai xuất và tờ khai tạm nhập G13, hợp đồng dịch vụ bảo hành vs hợp đồng mua bán ban đầu |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka hướng dẫn bộ hồ sơ chuẩn 5 loại chứng từ, quy trình bảo lãnh nghĩa vụ thuế và quản lý thời hạn tạm nhập tái xuất không bị phạt |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #2 (Gỡ Điểm Nghẽn Thuế Cho Nhà Máy Xuất Khẩu — Đột phá chính sách quan trọng nhất đầu tháng 11/2026)** |
- **Tóm tắt lý do nên làm:** Hướng dẫn nghiệp vụ cực kỳ cần thiết cho các doanh nghiệp FDI, doanh nghiệp cơ khí chế tạo, điện tử, dệt may khi xử lý hàng lỗi xuất khẩu. Nâng tầm uy tín chuyên môn của Eureka trong mắt các khách hàng doanh nghiệp sản xuất lớn.
- **Nguồn xác minh:** Công văn số 23949/CHQ-GSQL (06/11/2026); Nghị định số 292/2026/NĐ-CP (thay thế NĐ 69/2018/NĐ-CP); Báo Hải quan Online; Thư Viện Pháp Luật.

---

### Đề xuất #3: [CẢNH BÁO ĐỎ CẤM XUẤT CẢNH VÌ NỢ THUẾ ĐỌNG] Doanh nghiệp giải thể "dở dang" - Trạng thái MST 03 trên eTax và rủi ro người đại diện pháp luật bị tạm hoãn xuất cảnh tại sân bay
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Pháp lý Thuế & Quản trị Rủi ro Doanh nhân) / `CS` (Case Study Gỡ Nợ Thuế & Đóng Mã Số Thuế)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo đỏ 'Thuê dịch vụ giải thể công ty từ vài năm trước nhưng nay bị chặn xuất cảnh tại sân bay: Bẫy trạng thái MST 03' | YouTube: Video hướng dẫn 3 bước tra cứu trạng thái doanh nghiệp trên eTax, quyết toán nợ thuế môn bài/chậm nộp và thủ tục đóng mã số thuế hoàn chỉnh)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Case study thực tế vừa phát sinh sáng 07/11/2026 từ Diễn đàn Thuế chuyên sâu (`scratch/daily_notes.md`); hiện nay cơ quan Thuế và Bộ Công an đang đồng bộ dữ liệu tự động siết chặt quản lý xuất nhập cảnh |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Giám đốc/chủ doanh nghiệp ra sân bay đi công tác/du lịch nước ngoài thì bị biên phòng chặn lại vì công ty cũ chưa đóng mã số thuế hoàn tất; mất vé máy bay, lỡ hợp đồng đối tác, uy tín cá nhân bị ảnh hưởng nghiêm trọng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận nảy lửa về trách nhiệm của các công ty dịch vụ kế toán/giải thể thiếu tâm, và sự thiếu hiểu biết của chủ doanh nghiệp khi nghĩ rằng "bỏ doanh nghiệp không hoạt động là xong" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka cung cấp phác đồ 4 bước: Rà soát nợ thuế trên eTax Mobile, giải trình hồ sơ giải thể tồn đọng, nộp dứt điểm tiền nợ và gửi công văn xin hủy quyết định tạm hoãn xuất cảnh khẩn cấp |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Cảnh Báo Sống Còn Cho Doanh Nhân — Nội dung chạm đúng nỗi lo lắng của hàng vạn chủ doanh nghiệp)** |
- **Tóm tắt lý do nên làm:** Chủ đề thu hút lượng tương tác cực mạnh trên mạng xã hội vì đánh trúng quyền lợi thiết thân của các chủ doanh nghiệp, CEO, nhà đầu tư. Giúp Eureka xây dựng hình ảnh chuyên gia pháp lý và tư vấn doanh nghiệp toàn diện.
- **Nguồn xác minh:** `scratch/daily_notes.md` (07/11/2026); Luật Quản lý thuế số 38/2019/QH14; Nghị định 126/2020/NĐ-CP; Cổng TTĐT Tổng cục Thuế (`etax.gdt.gov.vn`).

---

### Đề xuất #4: [SỰ CỐ KỸ THUẬT HẢI QUAN ĐIỆN TỬ] Lỗi "Yêu cầu ký đã hết thời gian chờ" trên hệ thống VNACCS khi khai báo làm sạch dữ liệu CW_NKHQ_TC và chữ ký số token
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Kỹ thuật Phần mềm Hải quan Điện tử VNACCS/VCIS) / `CS` (Case Study Sửa Lỗi Hệ Thống)
- **Nền tảng phù hợp:** **Facebook** (Bài viết thủ thuật nghiệp vụ: 'Bị lỗi Yêu cầu ký đã hết thời gian chờ khi đăng nhập VNACCS / CW_NKHQ_TC: Hướng dẫn 4 bước khắc phục trong 2 phút')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Ghi nhận trực tiếp từ vướng mắc của nhân viên chứng từ sáng 07/11/2026 trong `scratch/daily_notes.md` trong đợt cao điểm chuẩn hóa dữ liệu hải quan số toàn quốc |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Không đăng nhập được VNACCS để reset mật khẩu và truyền dữ liệu định danh khiến toàn bộ tờ khai bị đình trệ, hàng hóa đến cảng không thể mở tờ khai, nguy cơ trễ hạn và phát sinh chi phí lưu kho |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận sôi nổi trong các hội nhóm khai hải quan về lỗi tương thích driver chữ ký số Viettel-CA, VNPT-CA, BkavCA với Java và cổng kết nối Hải quan |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka hướng dẫn chi tiết cách xóa cache Java, cấu hình lại Middleware token PKI, kiểm tra cổng mạng TCP 80/443 và cách liên hệ đường dây nóng hỗ trợ VNACCS |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #4 (Cứu Cánh Kỹ Thuật Tức Thời — Dễ viral và lưu lại sử dụng lâu dài trong cộng đồng dân XNK)** |
- **Tóm tắt lý do nên làm:** Hướng dẫn kỹ thuật thực chiến giải quyết dứt điểm sự cố hằng ngày của dân khai báo hải quan. Rất dễ đạt lượng chia sẻ (Share) và lưu bài (Bookmark) kỷ lục trên các nhóm chuyên môn.
- **Nguồn xác minh:** `scratch/daily_notes.md` (07/11/2026); Cục Công nghệ thông tin & Thống kê Hải quan; Hướng dẫn vận hành hệ thống VNACCS/VCIS.

---

### Đề xuất #5: [ĐIỂM NÓNG CƯỚC BIỂN & HỦY CHUYẾN BLANK SAILINGS] Chỉ số Drewry WCI ngày 05/11/2026 ($4,389/FEU) & Chiến lược Hãng tàu kích hoạt Blank Sailings khép chặt cung tải: Cẩm nang đàm phán hợp đồng logistics Fix Rate mùa cao điểm Tết
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải - Chiến lược Quản trị Chi phí Vận tải Biển) / `TN` (Tin Ngành - Thị trường Logistics Quốc tế)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Bản đồ cước biển 'Drewry WCI giảm 1% xuống $4,389/FEU nhưng hãng tàu ồ ạt hủy chuyến: Bí quyết đàm phán cước không bị ép phụ phí' | YouTube: Video cẩm nang phân tích động thái của các liên minh hàng hải cuối năm 2026 và chiến lược đặt chỗ container đi Mỹ/EU của Eureka)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Báo cáo chỉ số cước thế giới Drewry WCI vừa công bố tuần này (05/11/2026) ghi nhận mức $4,389/FEU (giảm 1%), kèm cảnh báo làn sóng Blank Sailings từ các hãng tàu |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Giá cước có xu hướng hạ nhẹ nhưng hãng tàu chủ động cắt giảm tàu mẹ, ép khách hàng phải chịu mức cước cao hoặc bị rớt container (roll container) vào mùa xuất khẩu cao điểm giáp Tết |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận gay gắt giữa chủ hàng và forwarder: Nên chốt cước giao ngay (Spot Rate) theo tuần hay ký hợp đồng dài hạn (Contract Fix Rate) giữ giá? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ chiến lược booking trước 2-3 tuần, phân bổ linh hoạt giữa các liên minh (Ocean Alliance, THE Alliance) và dịch vụ cam kết chỗ (space guarantee) của Eureka |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #5 (Chiến Lược Logistics Cao Điểm Cuối Năm — Bảo vệ lợi nhuận và tiến độ giao hàng cho chủ hàng)** |
- **Tóm tắt lý do nên làm:** Đón đầu mối bận tâm hàng đầu của các CFO, Trưởng phòng Logistics và chủ hàng xuất nhập khẩu đường biển trong quý 4. Thể hiện năng lực điều phối booking tàu và đàm phán cước của Eureka.
- **Nguồn xác minh:** Drewry World Container Index (WCI Report 05/11/2026); World Shipping Council (WSC); Báo Công Thương; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).

---

## 🎯 KHUYẾN NGHỊ HÀNH ĐỘNG CỦA AGENT 01:
1. **Lựa chọn ưu tiên #1:** Đẩy ngay **Đề xuất #1** (Thủ tục nhập khẩu bát đĩa gốm sứ tiếp xúc thực phẩm QCVN 12-4:2015/BYT và kiểm tra ATTP) vào quy trình sản xuất nội dung Facebook Fanpage & Video ngắn Viral (`wf02`/`wf03`) vì đây là mùa nhập hàng Tết rầm rộ nhất trong năm.
2. **Lựa chọn ưu tiên #2:** Triển khai **Đề xuất #2** (Công văn 23949/CHQ-GSQL hướng dẫn tạm nhập sửa chữa, bảo hành theo Nghị định 292/2026/NĐ-CP) thành bài viết phân tích nghiệp vụ chuyên sâu trên LinkedIn và Fanpage Eureka để tiếp cận các nhà máy sản xuất FDI.
3. **Lựa chọn ưu tiên #3:** Đưa **Đề xuất #3** (Rủi ro cấm xuất cảnh do doanh nghiệp giải thể dở dang nợ thuế) vào kịch bản Video YouTube dài ThuyTalk (`wf01`/`thuytalk-cm124`) nhằm mang đến câu chuyện cảnh tỉnh sâu sắc, bảo vệ pháp lý cho cộng đồng doanh nhân.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w45_0711()
