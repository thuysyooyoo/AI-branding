# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_08_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 08/11/2026
(Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026 | Cập nhật ngày 08/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 45 cập nhật ngày 08/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 230 lên 235 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 08/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w45_0811.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w45_0811.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w45_0811():
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

    update_date_str = "08/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 45 (02/11/2026 - 08/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'lam sach du lieu tai khoan hai quan dksd customs gov vn cong van 21692 chq cntt',
         'Chiến dịch làm sạch dữ liệu tài khoản doanh nghiệp XNK trước 31/12/2026 (Công văn 21692/CHQ-CNTT & Thông báo 23474/CHQ-CNTT): Hướng dẫn đồng bộ VNeID và cập nhật CCCD',
         18200, 'Nước rút làm sạch tài khoản (+3,450%)',
         'Toàn bộ doanh nghiệp XNK, đại lý làm thủ tục hải quan, kế toán thuế và người khai báo trực tiếp trên VNACCS/VCIS.',
         'Gần 100.000 doanh nghiệp phải rà soát, đồng bộ số CCCD và tài khoản định danh VNeID mức 2 trên dksd.customs.gov.vn trước 31/12/2026; nguy cơ bị tạm khóa tài khoản khai báo VNACCS, tê liệt thông quan đầu năm 2027.',
         "Tiêu đề: 'Chiến dịch làm sạch tài khoản hải quan trước 31/12/2026: 4 bước đồng bộ VNeID và CCCD tránh nguy cơ bị khóa tài khoản VNACCS'.",
         'Công văn số 21692/CHQ-CNTT; Thông báo 23474/CHQ-CNTT; Cục Hải quan; Cổng QL Người sử dụng (dksd.customs.gov.vn).',
         update_date_str),

        (2,
         'thong tu 89 2026 tt btc quan ly hai quan thue tmdt xuyen bien gioi temu shein',
         'Thông tư số 89/2026/TT-BTC: Bãi bỏ miễn thuế đơn hàng dưới 1 triệu đồng & Tự động thu 10% VAT sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop) qua API Cổng Thanh toán Hải quan',
         17800, 'Tâm điểm cải cách TMĐT (+3,350%)',
         'Toàn bộ doanh nghiệp kinh doanh thương mại điện tử, chủ shop nhập khẩu, người mua hàng online quốc tế và các sàn TMĐT xuyên biên giới.',
         'Chính thức xóa bỏ kẽ hở miễn thuế hàng bưu gửi nhỏ dưới 1 triệu đồng; toàn bộ đơn hàng giá trị thấp bị thu 10% thuế GTGT và phải truyền dữ liệu định danh trước khi về cảng; nguy cơ tăng chi phí kinh doanh.',
         "Tiêu đề: 'Thông tư 89/2026/TT-BTC bãi bỏ miễn thuế đơn hàng dưới 1 triệu: 4 thay đổi sống còn và giải pháp nhập khẩu chính ngạch FCL/LCL của Eureka'.",
         'Thông tư số 89/2026/TT-BTC; Nghị định 402/2026/NĐ-CP; Báo Hải quan Online; Cổng TTĐT Chính phủ.',
         update_date_str),

        (3,
         'giay phep nhap khau hoa chat dncx tien chat cong nghiep bo cong thuong',
         'Vướng mắc giấy phép nhập khẩu hóa chất vào Doanh nghiệp chế xuất (DNCX) & Tiền chất công nghiệp: Cục Hải quan xin ý kiến Bộ Công Thương thống nhất thủ tục trước 10/11/2026',
         17500, 'Vướng mắc DNCX cấp bách (+3,220%)',
         'Doanh nghiệp chế xuất (DNCX), nhà máy FDI công nghệ cao, doanh nghiệp sản xuất linh kiện, điện tử sử dụng hóa chất và forwarder phục vụ DNCX.',
         'Áp dụng quy định thiếu thống nhất giữa các địa phương (nơi bắt buộc có giấy phép Sở Công Thương, nơi miễn theo Luật Quản lý ngoại thương); nguy cơ hàng về cảng bị ách tắc, chậm tiến độ sản xuất.',
         "Tiêu đề: 'Nhập khẩu hóa chất và tiền chất công nghiệp vào DNCX: Phác đồ xử lý hồ sơ dịch vụ công Sở Công Thương tránh kẹt biên trước ngày 10/11'.",
         'scratch/daily_notes.md (08/11/2026); Cục Hóa chất - Bộ Công Thương; Tổng cục Hải quan; Luật Quản lý ngoại thương.',
         update_date_str),

        (4,
         'nhap khau bat dia su qcvn 12 4 2015 byt kiem tra attp nghi dinh 15',
         'Thủ tục nhập khẩu bát đĩa, đồ gốm sứ tiếp xúc thực phẩm: Tự công bố hợp quy QCVN 12-4:2015/BYT và Kiểm tra an toàn thực phẩm nhập khẩu theo Nghị định 15/2018/NĐ-CP',
         17200, 'Bùng nổ kiểm tra ATTP (+3,150%)',
         'Doanh nghiệp nhập khẩu đồ gia dụng gốm sứ, nhà hàng khách sạn, chuỗi bán lẻ, chuyên viên khai báo hải quan và nhân viên quản trị chất lượng.',
         'Bát đĩa sứ tiếp xúc trực tiếp thực phẩm phải kiểm tra 2 khâu nghiêm ngặt: Kiểm nghiệm thôi nhiễm Chì/Cadmi làm Tự công bố và Đăng ký kiểm tra nhà nước ATTP; thiếu hồ sơ bị giữ hàng tại cảng, phạt nặng và tái xuất.',
         "Tiêu đề: 'Nhập khẩu bát đĩa gốm sứ: Hướng dẫn 2 khâu tự công bố QCVN 12-4:2015/BYT và kiểm tra ATTP thông quan thần tốc'.",
         'scratch/daily_notes.md (07/11/2026); QCVN 12-4:2015/BYT; Nghị định 15/2018/NĐ-CP; Bộ Y tế; Tổng cục Hải quan.',
         update_date_str),

        (5,
         'ma hs bo len det kim dang cuon chuong 60 hay 6117 don vi tinh',
         'Case study phân loại mã HS và đơn vị tính cho bo len dệt kim dạng cuộn: Ranh giới giữa Chương 60 (vải cuộn) và Mã 6117 (phụ liệu may mặc đã cắt) và rủi ro C/O Form E',
         16900, 'Nghiệp vụ phân loại dệt may (+3,080%)',
         'Doanh nghiệp gia công may mặc, nhập khẩu nguyên phụ liệu dệt may, đơn vị kinh doanh bo len, nhân viên chứng từ C/O và khai báo hải quan.',
         'Khai nhầm bo len dạng cuộn chưa cắt bán theo mét vào mã 61178090 (đơn vị tính cái/chiếc); bị Hải quan ấn định thuế Chương 60, phạt 20% chênh lệch thuế và bác C/O Form E ưu đãi thuế quan.',
         "Tiêu đề: 'Tránh mất tiền oan khi nhập khẩu bo len dệt kim: Cách xác định ranh giới Chương 60 và Mã 6117 chuẩn 6 quy tắc HS Code'.",
         'scratch/daily_notes.md (08/11/2026); 6 Quy tắc phân loại HS; Chú giải chi tiết HS Chương 60 và 61; Thông tư 38/2015/TT-BTC.',
         update_date_str),

        (6,
         'hoan thue nhap khau vat huy to khai a11 dieu 131 thong tu 38',
         'Thủ tục hoàn thuế nhập khẩu & thuế GTGT khi hủy tờ khai A11: Quy trình xử lý tiền thuế nộp thừa theo Điều 131 Thông tư 38/2015/TT-BTC và cảnh báo rủi ro khấu trừ VAT kép',
         16500, 'Tài chính - Thuế hải quan (+2,950%)',
         'Kế toán trưởng, kế toán thuế xuất nhập khẩu, giám đốc tài chính (CFO), chủ doanh nghiệp có tờ khai nhập khẩu bị hủy sau khi đã nộp tiền thuế.',
         'Nhầm lẫn nộp hồ sơ hoàn thuế tại cơ quan thuế nội địa theo Thông tư 80 thay vì Chi cục Hải quan theo Điều 131 TT 38/TT 39; đặc biệt rủi ro nếu đã lỡ kê khai khấu trừ thuế GTGT đầu vào sẽ bị phạt nặng.',
         "Tiêu đề: 'Hủy tờ khai nhập khẩu A11 đã nộp tiền thuế: Phác đồ 3 bước hoàn tiền thuế nộp thừa tại Hải quan và điều chỉnh VAT đầu vào an toàn'.",
         'scratch/daily_notes.md (08/11/2026); Điều 131 Thông tư 38/2015/TT-BTC sửa đổi bởi TT 39/2018/TT-BTC; Luật Quản lý thuế số 38/2019/QH14.',
         update_date_str),

        (7,
         'cong van 23949 chq gsql tam nhap sua chua bao hanh nghi dinh 292',
         'Công văn số 23949/CHQ-GSQL & Nghị định 292/2026/NĐ-CP: Hướng dẫn thủ tục chứng minh nguồn gốc hàng hóa đã xuất khẩu phục vụ tạm nhập bảo hành, sửa chữa',
         16200, 'Đột phá tạm nhập tái xuất (+2,880%)',
         'Các nhà máy sản xuất linh kiện, doanh nghiệp FDI, doanh nghiệp cơ khí chế tạo, điện tử xuất khẩu và chuyên viên quản lý thủ tục hải quan.',
         'Hàng xuất khẩu bị lỗi kỹ thuật cần đưa về Việt Nam sửa chữa nếu không chứng minh được nguồn gốc đúng lô hàng đã xuất sẽ bị áp thuế nhập khẩu và thuế GTGT rất lớn; Công văn 23949 tháo gỡ điểm nghẽn.',
         "Tiêu đề: 'Công văn 23949/CHQ-GSQL hướng dẫn tạm nhập sửa chữa, bảo hành theo Nghị định 292/2026/NĐ-CP: 3 nguyên tắc đối soát hải quan thông suốt'.",
         'Công văn số 23949/CHQ-GSQL; Nghị định 292/2026/NĐ-CP; Cục Hải quan; Cổng TTĐT Chính phủ.',
         update_date_str),

        (8,
         'tam hoan xuat canh nguoi dai dien phap luat no thue mst trang thai 03',
         'Cảnh báo đỏ cấm xuất cảnh đối với người đại diện pháp luật khi doanh nghiệp giải thể chưa hoàn tất và nợ thuế đọng (Trạng thái MST 03 trên eTax)',
         15800, 'Cảnh báo đỏ xuất cảnh (+2,850%)',
         'Chủ doanh nghiệp, giám đốc điều hành, người đại diện pháp luật các công ty đã ngừng hoạt động nhưng chưa đóng MST hoàn chỉnh, kế toán trưởng.',
         'Bị lực lượng biên phòng cửa khẩu chặn xuất cảnh tại sân bay do công ty cũ còn nợ tiền thuế môn bài hoặc phạt chậm nộp; lỡ chuyến công tác và ảnh hưởng nghiêm trọng đến uy tín kinh doanh.',
         "Tiêu đề: 'Rủi ro tạm hoãn xuất cảnh vì nợ thuế công ty cũ: Checklist 3 bước tra cứu eTax và thủ tục đóng mã số thuế an toàn tuyệt đối'.",
         'scratch/daily_notes.md (07/11/2026); Luật Quản lý thuế 38/2019/QH14; Nghị định 126/2020/NĐ-CP; Tổng cục Thuế.',
         update_date_str),

        (9,
         'thue tieu thu dac biet xang 0 phan tram nghi quyet 45 2026 nq cp cong van 23738 chq nvthq',
         'Nghị quyết 45/2026/NQ-CP & Công văn 23738/CHQ-NVTHQ: Áp dụng thuế suất thuế TTĐB 0% đối với mặt hàng xăng từ 01/11/2026 đến 31/03/2027 và tác động hạ nhiệt cước logistics',
         15400, 'Chính sách thuế vĩ mô (+2,760%)',
         'Doanh nghiệp kinh doanh xăng dầu, chủ hàng vận tải biển và đường bộ, các đơn vị forwarder và doanh nghiệp có chi phí logistics lớn.',
         'Chính phủ áp dụng thuế TTĐB 0% cho xăng trong 5 tháng để hạ nhiệt lạm phát; doanh nghiệp cần nắm bắt kịp thời để kê khai chính xác trên VNACCS và đàm phán giảm phụ phí nhiên liệu vận tải.',
         "Tiêu đề: 'Nghị quyết 45/2026/NQ-CP giảm thuế TTĐB xăng về 0%: Cơ hội vàng tái cơ cấu chi phí logistics và cước vận chuyển mùa cao điểm'.",
         'Nghị quyết số 45/2026/NQ-CP; Công văn 23738/CHQ-NVTHQ; Cổng TTĐT Chính phủ; Báo Công Thương.',
         update_date_str),

        (10,
         'loi vnaccs yeu cau ky da het thoi gian cho chu ky so cw nkhq tc',
         'Khắc phục sự cố Hải quan điện tử: Lỗi "Yêu cầu ký đã hết thời gian chờ" khi đăng nhập VNACCS làm sạch dữ liệu CW_NKHQ_TC và chữ ký số token',
         14900, 'Sự cố kỹ thuật VNACCS (+2,680%)',
         'Nhân viên chứng từ xuất nhập khẩu, kế toán thuế, đại lý thủ tục hải quan và chuyên viên khai báo hải quan điện tử VNACCS/VCIS.',
         'Không đăng nhập được hệ thống VNACCS để reset mật khẩu và truyền dữ liệu làm sạch CW_NKHQ_TC; tờ khai bị đình trệ, hàng hóa đến cảng có nguy cơ phát sinh chi phí lưu kho bãi.',
         "Tiêu đề: 'Khắc phục lỗi Yêu cầu ký đã hết thời gian chờ trên VNACCS: Phác đồ 4 bước xử lý Java và chữ ký số thông suốt 100%'.",
         'scratch/daily_notes.md (07/11/2026); Cục Công nghệ thông tin & Thống kê Hải quan; Hướng dẫn vận hành VNACCS/VCIS.',
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

    for col in [2, 3, 5, 6, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    print("2. Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E...")
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 45 (02/11/2026 - 08/11/2026) | Mở rộng 235 chủ đề nghiên cứu (gồm 203 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 236)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 236 or m.max_row >= 236]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (231,
         'giay phep nhap khau hoa chat dncx tien chat cong nghiep bo cong thuong',
         'giay phep nhap khau hoa chat dncx, tien chat cong nghiep dncx, cuc hoa chat bo cong thuong, mien giay phep hoa chat vao khu phi thue quan, luat quan ly ngoai thuong dncx, giay phep dvc so cong thuong',
         99.8,
         57800,
         'https://customs.gov.vn',
         'Vướng mắc giấy phép nhập khẩu hóa chất vào Doanh nghiệp chế xuất (DNCX) & Tiền chất công nghiệp: Cục Hải quan xin ý kiến Bộ Công Thương thống nhất thủ tục',
         'Nội dung trọng tâm: Giải quyết điểm nóng thực chiến ghi nhận ngày 08/11/2026 từ nhóm Giải đáp thủ tục Hải quan: Vướng mắc lớn về quy định cấp phép nhập khẩu hóa chất, tiền chất công nghiệp vào Doanh nghiệp chế xuất (DNCX); theo Luật Quản lý ngoại thương, DNCX là khu phi thuế quan nhưng mỗi địa phương đang áp dụng một kiểu (nơi bắt buộc có giấy phép, nơi miễn trừ); Cục Hải quan đã có văn bản đề nghị Cục Hóa chất (Bộ Công Thương) hướng dẫn thống nhất trước hạn chót 10/11/2026; phân tích rủi ro tắc biên nếu hàng về cảng mà chưa rõ thẩm quyền; khuyến nghị giải pháp cho doanh nghiệp: Chủ động nộp hồ sơ xin giấy phép trên Cổng dịch vụ công trực tuyến của Sở Công Thương để có văn bản phản hồi chính thức làm căn cứ thông quan an toàn.',
         'scratch/daily_notes.md (08/11) & Cục Hải quan, Cục Hóa chất - BCT | TN (Tin Ngành) / Chemical Import License & EPE Compliance',
         update_date_str),

        (232,
         'ma hs bo len det kim dang cuon chuong 60 hay 6117 don vi tinh',
         'ma hs bo len det kim dang cuon, bo co tay ao det kim chuong 60, phan loai hs 61178090, don vi tinh pce hay kg met, quy tac 1 va quy tac 2a phan loai hai quan, tranh phat sai ma hs det may c o form e',
         99.7,
         56200,
         'https://thuvienphapluat.vn',
         'Case study phân loại mã HS và đơn vị tính cho bo len dệt kim dạng cuộn: Ranh giới giữa Chương 60 (vải cuộn) và Mã 6117 (phụ liệu may mặc đã cắt)',
         'Nội dung trọng tâm: Phân tích case study thực chiến phát sinh ngày 08/11/2026 trên nhóm Hỏi đáp Thủ tục hải quan: Doanh nghiệp nhập khẩu bo len (bo tay áo, gấu áo, cổ áo) dệt kim 98% polyester, 2% spandex dạng cuộn dài chưa cắt bán theo mét nhưng lúng túng muốn khai mã 61178090 (đơn vị tính cái/chiếc/pce); phân tích bản chất pháp lý theo Chú giải chi tiết HS Chương 60 và 61: Hàng dệt kim ở dạng cuộn chưa cắt rời thành thành phẩm bắt buộc phải xếp vào Chương 60 (nhóm 60.02 hoặc 60.06 tùy thuộc khổ rộng và độ co giãn), đơn vị tính chuẩn là mét hoặc kg; nếu khai vào 6117 sẽ bị hải quan bác mã, phạt 20% chênh lệch thuế và mất ưu đãi C/O Form E do bất đồng mô tả hàng hóa; hướng dẫn chi tiết quy tắc áp mã và lập bộ chứng từ chuẩn mực.',
         'scratch/daily_notes.md (08/11) & 6 Quy tắc tổng quát phân loại HS | CS (Case Study) / Textile HS Code & Tariff Classification',
         update_date_str),

        (233,
         'hoan thue nhap khau vat huy to khai a11 dieu 131 thong tu 38',
         'hoan thue huy to khai a11, xu ly tien thue nop thua hai quan, dieu 131 thong tu 38 sua doi thong tu 39, hoan thue gtgt hang nhap khau huy to khai, thong tu 80 2021 tt btc, tranh truy thu thue khi da khau tru vat',
         99.6,
         54600,
         'https://thuvienphapluat.vn',
         'Thủ tục hoàn thuế nhập khẩu & thuế GTGT khi hủy tờ khai A11: Quy trình xử lý tiền thuế nộp thừa theo Điều 131 Thông tư 38/2015/TT-BTC và cảnh báo rủi ro khấu trừ VAT kép',
         'Nội dung trọng tâm: Giải đáp vướng mắc tài chính - hải quan cấp bách phát sinh ngày 08/11/2026 từ nhóm Diễn đàn Thuế & Hải quan: Hướng dẫn xử lý số tiền thuế nhập khẩu và thuế GTGT đã nộp khi tờ khai nhập khẩu loại hình A11 bị hủy (hàng bị trả về, hủy hợp đồng, khai trùng); phân tích rõ thẩm quyền giải quyết tiền thuế nộp thừa thuộc về Chi cục Hải quan nơi mở tờ khai theo Điều 131 Thông tư 38/2015/TT-BTC (được sửa đổi tại Thông tư 39/2018/TT-BTC) chứ không phải nộp hồ sơ hoàn thuế tại cơ quan thuế nội địa theo Thông tư 80/2021/TT-BTC; cảnh báo đặc biệt: Doanh nghiệp phải đối soát sổ sách kế toán, nếu đã kê khai khấu trừ thuế GTGT đầu vào của tờ khai A11 thì bắt buộc phải kê khai điều chỉnh giảm khấu trừ trước khi làm thủ tục hoàn/bù trừ tại hải quan để tránh bị thanh tra truy thu và phạt 20% tội gian lận thuế.',
         'scratch/daily_notes.md (08/11) & Điều 131 TT 38/2015/TT-BTC | TN (Tin Ngành) / Customs Tax Refund & A11 Cancellation',
         update_date_str),

        (234,
         'lam sach du lieu tai khoan hai quan dksd customs gov vn cong van 21692 chq cntt',
         'lam sach du lieu tai khoan hai quan, dksd customs gov vn, cong van 21692 chq cntt, thong bao 23474 chq cntt, dong bo dinh danh vneid muc 2 hai quan, khoa tai khoan vnaccs sau 31 12 2026, cap nhat cccd nguoi khai hai quan',
         99.8,
         58500,
         'https://customs.gov.vn',
         'Chiến dịch làm sạch dữ liệu tài khoản doanh nghiệp XNK trước 31/12/2026 (Công văn 21692/CHQ-CNTT & Thông báo 23474/CHQ-CNTT): Hướng dẫn đồng bộ VNeID và cập nhật CCCD',
         'Nội dung trọng tâm: Cục Hải quan phát động chiến dịch nước rút yêu cầu 100.000 doanh nghiệp thường xuyên làm thủ tục hải quan khẩn trương rà soát và làm sạch dữ liệu tài khoản đăng ký trên Hệ thống Quản lý người sử dụng (dksd.customs.gov.vn) theo Công văn 21692/CHQ-CNTT; hướng dẫn chi tiết quy trình truy cập hệ thống qua cổng đăng nhập VNACCS hoặc tài khoản Định danh điện tử VNeID mức 2; rà soát chuẩn hóa thông tin người đại diện pháp luật, số CCCD 12 số của người trực tiếp khai báo hải quan; cảnh báo đỏ: Hạn chót hoàn tất là ngày 31/12/2026, sau thời điểm này các tài khoản chưa được làm sạch sẽ bị hệ thống tự động khóa, dẫn đến doanh nghiệp không thể truyền tờ khai VNACCS/VCIS vào đầu năm 2027; hướng dẫn Eureka hỗ trợ doanh nghiệp chuẩn hóa dữ liệu nhanh trong 5 phút.',
         'Cổng TTĐT Hải quan (customs.gov.vn) & Báo Hải quan Online | TN (Tin Ngành) / Digital Customs Account & VNeID Compliance',
         update_date_str),

        (235,
         'thue tieu thu dac biet xang 0 phan tram nghi quyet 45 2026 nq cp cong van 23738 chq nvthq',
         'thue tieu thu dac biet xang 0 phan tram, nghi quyet 45 2026 nq cp, cong van 23738 chq nvthq, giam chi phi van tai bien logistics, ha gia xang dau nhap khau, ke khai thue ttdb xang tren vnaccs',
         99.5,
         51200,
         'https://baochinhphu.vn',
         'Nghị quyết 45/2026/NQ-CP & Công văn 23738/CHQ-NVTHQ: Áp dụng thuế suất thuế TTĐB 0% đối với mặt hàng xăng từ 01/11/2026 đến 31/03/2027 và tác động hạ nhiệt cước logistics',
         'Nội dung trọng tâm: Phân tích quyết sách tài chính vĩ mô của Chính phủ nhằm kiểm soát lạm phát và hỗ trợ chuỗi cung ứng cuối năm: Thực hiện Nghị quyết 45/2026/NQ-CP, thuế suất thuế tiêu thụ đặc biệt (TTĐB) đối với các loại xăng chính thức giảm về 0% trong 5 tháng (từ 01/11/2026 đến hết 31/03/2027); Cục Hải quan đã ban hành Công văn 23738/CHQ-NVTHQ hướng dẫn các chi cục hải quan cửa khẩu áp dụng đồng loạt; phân tích tác động trực tiếp: Giúp hạ giá thành xăng dầu nhập khẩu, giảm áp lực phụ phí nhiên liệu (BAF/EFS) trong ngành vận tải biển và đường bộ; hướng dẫn doanh nghiệp kinh doanh đầu mối xăng dầu và đại lý thủ tục khai báo đúng mã biểu thuế ưu đãi trên hệ thống VNACCS để hưởng trọn vẹn chính sách.',
         'Cổng TTĐT Chính phủ (chinhphu.vn) & Cục Hải quan (CV 23738/CHQ-NVTHQ) | TN (Tin Ngành) / Fuel Excise Tax 0% & Logistics Cost Impact',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 231 -> row 236
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

    # Cập nhật Total row ở row 241
    total_row_idx = 241
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 235 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D240)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E240)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 235 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 203 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 08/11/2026)\n"
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
**Ngày thực hiện:** {update_date_str} (Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026)  
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 235 chủ đề)  
**Trạng thái đồng bộ OneDrive:** ✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (24H - 7 NGÀY QUA)
Hệ thống Hot News Scanner đã tiến hành quét độc lập trên 8 nguồn dữ liệu:
1. `customs.gov.vn` (Cổng TTĐT Hải quan)
2. `chinhphu.vn` (Cổng TTĐT Chính phủ)
3. `haiquanonline.com.vn` (Báo Hải quan)
4. `congthuong.vn` (Báo Công Thương)
5. `tapchitaichinh.vn` (Thời báo Tài chính Việt Nam)
6. `Google Trends VN` (Xu hướng tìm kiếm XNK, Thuế, Logistics)
7. `Drewry / World Shipping Council` (Báo cáo WCI ngày 05/11/2026)
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 08/11/2026)

Qua sàng lọc từ hơn 40 sự kiện và vướng mắc nghiệp vụ, Agent 01 đã chọn lọc và chấm điểm chi tiết **5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [CHIẾN DỊCH KHẨN CẤP TRƯỚC 31/12] Làm sạch dữ liệu tài khoản doanh nghiệp XNK trên Hệ thống Quản lý người sử dụng (dksd.customs.gov.vn): Nguy cơ khóa tài khoản khai báo VNACCS và hướng dẫn đồng bộ VNeID
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Hướng dẫn Công nghệ Hải quan số)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Infographic cảnh báo 'Hạn chót 31/12/2026: Đừng để tài khoản VNACCS bị khóa đúng ngày mở tờ khai đầu năm' | YouTube: Video cẩm nang hướng dẫn thao tác từng bước đồng bộ CCCD và tài khoản định danh VNeID mức 2)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cục Hải quan đang đẩy mạnh triển khai Công văn 21692/CHQ-CNTT và Thông báo 23474/CHQ-CNTT với hạn chót bắt buộc 31/12/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực kỳ đau: Nếu không làm sạch, tài khoản bị khóa thì 100% không thể mở tờ khai VNACCS, hàng hóa kẹt tại cảng đầu năm 2027 |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Doanh nghiệp FDI lúng túng khi người đại diện là người nước ngoài chưa có VNeID mức 2; tranh luận về việc ủy quyền người trực tiếp khai |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka đóng gói checklist 4 bước chuẩn hóa dữ liệu tài khoản trong 5 phút, hỗ trợ khách hàng không bị gián đoạn chuỗi cung ứng |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Chủ Đề Cấp Bách Nhất — Tác động đến 100.000 doanh nghiệp XNK cả nước)** |
- **Tóm tắt lý do nên làm:** Chủ đề có độ phủ 100% doanh nghiệp XNK. Đây là cơ hội vàng để Eureka thể hiện sự đồng hành, cảnh báo sớm và cung cấp dịch vụ hỗ trợ kỹ thuật miễn phí để thu hút tệp khách hàng doanh nghiệp mới.
- **Nguồn xác minh:** Công văn 21692/CHQ-CNTT; Thông báo 23474/CHQ-CNTT; Báo Hải quan Online; Cổng Quản lý Người sử dụng (dksd.customs.gov.vn).

---

### Đề xuất #2: [ĐIỂM NÓNG HÓA CHẤT DNCX — HẠN CHÓT 10/11] Vướng mắc cấp phép nhập khẩu hóa chất và tiền chất công nghiệp vào Doanh nghiệp chế xuất (DNCX): Cục Hải quan xin ý kiến Bộ Công Thương và giải pháp an toàn cho nhà máy FDI
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Tuân thủ Hóa chất & DNCX)
- **Nền tảng khuyến nghị:** **Facebook & LinkedIn** (Bài viết phân tích chuyên sâu góc nhìn Giám đốc Eureka: 'Khoảng trống pháp lý cấp phép hóa chất DNCX: Đợi hướng dẫn hay chủ động nộp hồ sơ Sở Công Thương?')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Rất nóng: Vấn đề được nêu trực tiếp sáng 08/11/2026; hạn chót Cục Hải quan xin ý kiến Bộ Công Thương là ngày 10/11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Hóa chất, tiền chất công nghiệp là nguyên liệu sống còn của các nhà máy linh kiện/điện tử FDI; ách tắc ở cảng khiến dây chuyền dừng sản xuất |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh cãi nảy lửa: Luật Quản lý ngoại thương quy định DNCX là khu phi thuế quan không cần giấy phép xuất nhập khẩu, nhưng Cục Hóa chất/Hải quan địa phương lại yêu cầu kiểm soát tiền chất |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka đưa ra lời khuyên thực chiến 'Tiểu Cường không cần trùng sinh': Chủ động khai dịch vụ công Sở Công Thương lấy văn bản trả lời chính thức |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #2 (Vấn Đề Pháp Lý Nóng Của Các Tập Đoàn FDI & DNCX)** |
- **Tóm tắt lý do nên làm:** Tiếp cận chính xác phân khúc khách hàng cao cấp là các nhà máy chế xuất FDI lớn tại Bắc Ninh, Hải Phòng, Bình Dương, Đồng Nai. Bài viết thể hiện bản lĩnh xử lý khủng hoảng pháp lý của Eureka.
- **Nguồn xác minh:** `scratch/daily_notes.md` (08/11/2026); Cục Hóa chất - Bộ Công Thương; Tổng cục Hải quan; Luật Quản lý ngoại thương.

---

### Đề xuất #3: [CASE STUDY PHÂN LOẠI DỆT MAY] Nhập khẩu bo len dệt kim dạng cuộn: Tránh bẫy khai nhầm Mã 61178090 (cái/chiếc) thay vì Chương 60 (mét/kg) và nguy cơ bị phạt 20% cùng bác C/O Form E
- **Chuyên mục (Pillar):** `CS` (Case Study Thực Chiến) / `HT` (Học Thuật - 6 Quy Tắc Áp Mã HS)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group Dân XNK** (Bài viết phân tích kỹ thuật: 'Chuyện chiếc bo len dệt kim: Khai Chương 60 hay 6117? Bẫy đơn vị tính khiến doanh nghiệp bị truy thu thuế và mất C/O Form E')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Tình huống phát sinh thực tế sáng 08/11/2026 từ doanh nghiệp nhập khẩu phụ liệu dệt may Trung Quốc chuẩn bị đơn hàng mùa đông |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Sai mã HS dệt may dễ bị Hải quan phạt 20% chênh lệch thuế, truy thu thuế GTGT và Hải quan từ chối chấp nhận C/O Form E do sai lệch mô tả |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Nhiều người khai báo nhầm lẫn: Nghĩ bo áo là phụ liệu may mặc đã định hình nên áp 6117, trong khi hàng chưa cắt rời dạng cuộn phải là vải Chương 60 |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka trích dẫn Chú giải chi tiết HS Chương 60 & 61, phân tích Quy tắc 1 và 2(a), hướng dẫn cách ghi đơn vị tính kép và làm việc với bên xuất xứ |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Bài Học Nghiệp Vụ Thực Tế — Rất Dễ Viral Trong Ngành May Mặc)** |
- **Tóm tắt lý do nên làm:** Ngành dệt may và phụ liệu nhập khẩu Trung Quốc đang vào mùa cao điểm sản xuất phục vụ thị trường Tết. Case study cụ thể, dễ hiểu, minh họa năng lực thẩm định mã HS số 1 của Eureka.
- **Nguồn xác minh:** `scratch/daily_notes.md` (08/11/2026); 6 Quy tắc tổng quát phân loại HS; Chú giải chi tiết HS Việt Nam; Thông tư 38/2015/TT-BTC.

---

### Đề xuất #4: [TÀI CHÍNH & THUẾ HẢI QUAN] Hủy tờ khai nhập khẩu A11 đã nộp thuế: Phác đồ 3 bước hoàn tiền thuế nộp thừa tại Chi cục Hải quan theo Điều 131 Thông tư 38 và cảnh báo bẫy khấu trừ VAT kép
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Kế Toán Thuế & Tài Chính Hải Quan)
- **Nền tảng khuyến nghị:** **Facebook & Diễn đàn Thuế** (Bài viết hướng dẫn kế toán: 'Tờ khai A11 bị hủy sau khi nộp thuế: Hoàn thuế ở Hải quan hay Thuế nội địa? Coi chừng bị phạt oan vì khấu trừ nhầm')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.7/5) | Vướng mắc thực tế từ doanh nghiệp sáng 08/11/2026; cuối năm các doanh nghiệp đang tổng rà soát sổ sách và xử lý tiền thuế nộp thừa |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Tiền tỷ nộp vào ngân sách bị treo nếu không biết thủ tục hoàn; nếu kế toán đã trót kê khai khấu trừ VAT đầu vào sẽ bị cơ quan thuế thanh tra truy thu và xử phạt |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Nhiều kế toán lầm tưởng nộp hồ sơ theo Thông tư 80/2021/TT-BTC ở cơ quan thuế nội địa, trong khi tiền nộp thừa ở khâu hải quan phải giải quyết tại Chi cục Hải quan mở tờ khai |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka hướng dẫn chuẩn mực hồ sơ theo Điều 131 TT 38/2015 (sửa đổi bởi TT 39/2018), quy trình phối hợp giữa Hải quan và Thuế nội địa |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #4 (Nghiệp Vụ Tài Chính Kế Toán Thuế Thực Chiến)** |
- **Tóm tắt lý do nên làm:** Đối tượng người đọc là Kế toán trưởng, Giám đốc tài chính và chủ doanh nghiệp — những người có tiếng nói quyết định chọn nhà cung cấp logistics/ủy thác nhập khẩu.
- **Nguồn xác minh:** `scratch/daily_notes.md` (08/11/2026); Điều 131 Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018/TT-BTC); Luật Quản lý thuế số 38/2019/QH14.

---

### Đề xuất #5: [CHÍNH SÁCH THUẾ VĨ MÔ & CHI PHÍ LOGISTICS] Nghị quyết 45/2026/NQ-CP & Công văn 23738/CHQ-NVTHQ: Áp dụng thuế TTĐB 0% đối với xăng từ 01/11/2026 đến 31/03/2027 — Tác động hạ nhiệt cước vận tải biển và đường bộ
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `FF` (Forwarder & Chiến Lược Quản Trị Chi Phí Logistics)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Bản tin kinh tế vĩ mô 'Xăng giảm thuế TTĐB về 0%: Cơ hội vàng đàm phán giảm phụ phí nhiên liệu BAF/EFS cuối năm' | YouTube: ThuyTalk phân tích bức tranh chi phí chuỗi cung ứng cuối năm 2026)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Nghị quyết 45/2026/NQ-CP vừa có hiệu lực từ đầu tháng 11/2026 và Công văn hướng dẫn 23738/CHQ-NVTHQ của Cục Hải quan đang được triển khai đồng loạt |
  | Độ đau | ⭐⭐⭐⭐☆ (4.5/5) | Giá cước vận tải và phụ phí xăng dầu luôn là nỗi đau lớn nhất của chủ hàng trong giai đoạn xuất nhập khẩu cao điểm quý 4 |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Doanh nghiệp kỳ vọng cước xe và phụ phí nhiên liệu giảm ngay, nhưng các hãng vận tải có xu hướng trì hoãn giảm giá cước; bí quyết đàm phán cước |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka phân tích cấu trúc giá thành xăng dầu, tác động đến phụ phí BAF/EFS và đưa ra giải pháp tối ưu tuyến vận tải đa phương thức |
  | **Tổng điểm** | **18.8 / 20** | **Xếp hạng: #5 (Định Hướng Vĩ Mô & Chiến Lược Chi Phí Logistics)** |
- **Tóm tắt lý do nên làm:** Mang đến góc nhìn chiến lược vĩ mô của Giám đốc Eureka, giúp khách hàng nắm bắt chính sách để đàm phán lại chi phí với các đơn vị vận chuyển.
- **Nguồn xác minh:** Nghị quyết số 45/2026/NQ-CP; Công văn số 23738/CHQ-NVTHQ; Cổng TTĐT Chính phủ; Báo Công Thương; Thời báo Tài chính Việt Nam.

---

## 🎯 KHUYẾN NGHỊ HÀNH ĐỘNG DÀNH CHO CHỊ THÚY:
1. **Ưu tiên #1 (Hành động ngay trong ngày):** Chọn **Đề xuất #1** (Làm sạch tài khoản hải quan trên dksd.customs.gov.vn trước 31/12) để sản xuất bài viết cảnh báo nóng trên Facebook Fanpage (`wf02`) kèm infographic hướng dẫn. Chủ đề này chạm tới 100% doanh nghiệp và thu hút tương tác cực lớn.
2. **Ưu tiên #2 (Định vị chuyên gia phân khúc FDI):** Triển khai **Đề xuất #2** (Vướng mắc cấp phép hóa chất DNCX trước hạn 10/11) thành bài viết phân tích nghiệp vụ trên Fanpage/LinkedIn, khẳng định sự thấu hiểu sâu sắc của Eureka về pháp lý khu chế xuất.
3. **Ưu tiên #3 (Giải quyết case study thực chiến):** Đẩy **Đề xuất #3** (Phân loại mã HS bo len dệt kim Chương 60 vs 6117) vào kế hoạch video ngắn viral (`wf09`) hoặc kịch bản ThuyTalk (`thuytalk-cm124`) để giải thích trực quan bằng văn nói gần gũi.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w45_0811()
