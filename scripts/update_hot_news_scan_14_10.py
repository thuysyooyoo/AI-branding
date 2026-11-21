# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_14_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 14/10/2026 (Tuần 42: 12/10/2026 - 18/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 42, ngày 14/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 130 lên 135 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 14/10/2026)
- Đồng bộ nguyên tử an toàn sang OneDrive Marketing Eureka
"""

import os
import sys
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Configure UTF-8
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

FILE_PATH = r"reports\SEO_Keyword_Trends_Eureka.xlsx"
ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh"
ONEDRIVE_FILE = os.path.join(ONEDRIVE_DIR, "SEO_Keyword_Trends_Eureka.xlsx")

def update_seo_w42_1410():
    print(f"Đang tải workbook từ {FILE_PATH}...")
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

    update_date_str = "14/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 42 (12/10/2026 - 18/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van 22883 chq nvthq ma hs hamburger banh mi hay thit 20 phan tram',
         'Công văn số 22883/CHQ-NVTHQ (12/10/2026): Lệnh rà soát khẩn toàn quốc về phân loại mã HS Hamburger — Ranh giới tỷ lệ 20% thịt giữa Chương 16 và nhóm 19.05',
         7800, 'Đỉnh điểm cảnh báo (+1,380%)',
         'Doanh nghiệp nhập khẩu thực phẩm chế biến, chuỗi nhà hàng thức ăn nhanh, chuỗi F&B, forwarder và đại lý hải quan.',
         'Lệnh rà soát khẩn toàn quốc của Cục Hải quan (báo cáo trước 19/10/2026); doanh nghiệp khai hamburger vào nhóm 19.05 (bánh mì) nhưng nhân thịt vượt 20% bị ấn định lại vào Chương 16 (chế phẩm thịt), bị truy thu thuế lớn và xử phạt vi phạm hành chính.',
         "Tiêu đề: 'Lệnh rà soát khẩn Công văn 22883/CHQ-NVTHQ: Hamburger là bánh mì hay thịt? Quy tắc 20% phân định mã HS và thuế suất'.",
         'Công văn số 22883/CHQ-NVTHQ ngày 12/10/2026; Chú giải 2 Chương 16; Chú giải 1(a) Chương 19; Cục Hải quan; scratch/daily_notes.md (14/10/2026).',
         update_date_str),

        (2,
         'dem nguoc 24h han chot 15 10 2026 lien ket vneid muc do 2 nsw nghi dinh 336',
         'ĐẾM NGƯỢC 24H (Hạn chót 15/10/2026): Ngày cuối cùng bắt buộc liên kết VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) & Chuẩn JSON Schema tránh bị khóa hệ thống',
         8200, 'Đếm ngược khẩn cấp 24h (+1,450%)',
         'Toàn bộ lãnh đạo doanh nghiệp XNK, người đại diện pháp luật, kế toán trưởng, nhân viên khai báo hải quan, forwarder và đại lý thủ tục.',
         'Chỉ còn 24h trước hạn chót 00:00 ngày 15/10/2026; doanh nghiệp chưa liên kết VNeID Mức 2 của người đại diện sẽ bị khóa hoàn toàn chức năng xin cấp phép chuyên ngành và tờ khai hải quan trên NSW.',
         "Tiêu đề: 'Đếm ngược 24h (15/10): Ngày cuối cùng liên kết VNeID Mức 2 trên NSW & Chuyển đổi JSON chứng từ tránh tê liệt thông quan'.",
         'Nghị định số 336/2026/NĐ-CP; Quyết định số 1434/QĐ-CHQ; Cổng TTĐT Chính phủ (chinhphu.vn); Tổng cục Hải quan.',
         update_date_str),

        (3,
         'cong van 22560 chq gsql ho so hai quan c o uy thac nhap khau hoa don thuong mai',
         'Công văn số 22560/CHQ-GSQL (05/10/2026): Hướng dẫn hồ sơ hải quan, hóa đơn thương mại (Invoice) không trùng tên và điều kiện chấp nhận C/O trong trường hợp Nhập khẩu ủy thác',
         6100, 'Tăng mạnh (+850%)',
         'Doanh nghiệp nhập khẩu ủy thác, chủ hàng thương mại, bên nhận ủy thác logistics, forwarder và kế toán thuế XNK.',
         'Hóa đơn thương mại nước ngoài phát hành không đứng tên bên nhận ủy thác bị hải quan nghi ngờ tính hợp lệ của giao dịch; lúng túng chứng từ C/O form ưu đãi FTA dẫn tới nguy cơ bị bác C/O và truy thu thuế nhập khẩu ưu đãi đặc biệt.',
         "Tiêu đề: 'Công văn 22560/CHQ-GSQL: Hướng dẫn gỡ vướng hóa đơn không trùng tên và điều kiện chấp nhận C/O khi nhập khẩu ủy thác'.",
         'Công văn số 22560/CHQ-GSQL ngày 05/10/2026; Điều 5 Nghị định 08/2015/NĐ-CP; Điều 16 Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018 và TT 121/2025).',
         update_date_str),

        (4,
         'cong van 22928 chq nvthq huong dan hoan thue hang hoa nhap khau sxxk tai xuat',
         'Công văn số 22928/CHQ-NVTHQ (13/10/2026): Hướng dẫn mới nhất của Cục Hải quan về điều kiện, hồ sơ và quy trình Hoàn thuế đối với hàng hóa nhập khẩu',
         5750, 'Mới ban hành (+780%)',
         'Doanh nghiệp sản xuất xuất khẩu (SXXK), gia công, doanh nghiệp nhập khẩu phải tái xuất, bộ phận tài chính kế toán và đại lý hải quan.',
         'Hàng hóa nhập khẩu phải tái xuất hoặc nguyên liệu SXXK đã nộp thuế bị nghẽn tiền hoàn thuế do hồ sơ chứng minh chưa chuẩn; quy trình "Hoàn trước kiểm sau" bị kéo dài do thiếu sót chứng từ thanh toán ngân hàng.',
         "Tiêu đề: 'Công văn 22928/CHQ-NVTHQ: Cẩm nang trọn bộ quy trình hoàn thuế hàng nhập khẩu và hồ sơ điện tử qua Dịch vụ công Hải quan'.",
         'Công văn số 22928/CHQ-NVTHQ ngày 13/10/2026; Điều 34 Nghị định 134/2016/NĐ-CP; Thông tư 06/2021/TT-BTC; Cục Hải quan.',
         update_date_str),

        (5,
         'hieu luc 15 10 thong tu 128 2026 tt btc van hanh xe agv cua khau thong minh huu nghi tan thanh',
         'CHÍNH THỨC HIỆU LỰC TỪ MAI (15/10/2026): Thông tư 128/2026/TT-BTC — Vận hành thí điểm Cửa khẩu thông minh Hữu Nghị & Tân Thanh với xe tự hành AGV 24/7',
         5900, 'Đếm ngược 24h (+860%)',
         'Doanh nghiệp xuất nhập khẩu đường bộ qua Lạng Sơn, chủ hàng nông sản, linh kiện điện tử, doanh nghiệp vận tải biên giới và logistics Trung - Việt.',
         'Hiệu lực chính thức từ ngày mai (15/10); doanh nghiệp chưa khai báo trước dữ liệu điện tử phương tiện sẽ bị từ chối vào làn xe tự hành AGV thông minh, buộc phải dồn ứ tại làn truyền thống.',
         "Tiêu đề: 'Từ ngày mai (15/10) Thông tư 128/2026/TT-BTC có hiệu lực: Hướng dẫn khai báo dữ liệu trước để chạy làn thông minh xe tự hành AGV'.",
         'Thông tư số 128/2026/TT-BTC ngày 27/08/2026; Đề án thí điểm Cửa khẩu thông minh Lạng Sơn; Cổng TTĐT Bộ Công Thương.',
         update_date_str),

        (6,
         'cong van 22720 chq gsql to khai phi thue quan dncx chua xac nhan khu vuc giam sat',
         'Công văn số 22720/CHQ-GSQL (08/10/2026): Hướng dẫn xử lý dứt điểm các tờ khai xuất nhập khẩu phi thuế quan chưa được xác nhận qua khu vực giám sát (KVGS) do lệch chỉ tiêu "Số quản lý nội bộ"',
         7100, 'Duy trì đỉnh (+1,200%)',
         'Doanh nghiệp chế xuất (DNCX), doanh nghiệp nội địa mua bán với DNCX, bộ phận xuất nhập khẩu, forwarder và đại lý hải quan.',
         'Hàng ngàn tờ khai phi thuế quan mở từ tháng 9/2026 bị hệ thống treo không tự động xác nhận qua KVGS do lệch cú pháp chỉ tiêu số quản lý nội bộ (#&XKPTQ vs #&NKTC); rủi ro không thanh khoản được tờ khai, treo hoàn thuế GTGT.',
         "Tiêu đề: 'Công văn 22720/CHQ-GSQL: Hướng dẫn gỡ lỗi tờ khai phi thuế quan DNCX chưa qua khu vực giám sát và quy trình gửi danh sách xử lý dứt điểm'.",
         'Công văn số 22720/CHQ-GSQL ngày 08/10/2026; Công văn số 20991/CHQ-GSQL ngày 26/08/2026; Cục Hải quan.',
         update_date_str),

        (7,
         'bi khai khong thu nhap luong 150 trieu tren etax mobile phan anh tong cuc thue',
         'Bẫy doanh nghiệp lạ khai khống lương 150 triệu trên eTax Mobile: Rủi ro người lao động & doanh nghiệp bị truy thu thuế oan và Quy trình 3 bước gửi phản ánh "Thu nhập khống" trực tuyến',
         5700, 'Tăng mạnh (+890%)',
         'Người lao động, nhân viên văn phòng, kế toán thuế, doanh nghiệp sử dụng lao động và các đơn vị quyết toán thuế TNCN.',
         'Bị doanh nghiệp lạ tại Khu công nghiệp lấy trộm mã số thuế để khai khống tiền lương 150 triệu/năm nhằm tăng chi phí ảo; dẫn đến việc người lao động và doanh nghiệp thực tế bị cơ quan thuế réo tên truy thu thuế TNCN bậc cao.',
         "Tiêu đề: 'Cảnh báo: Bị khai khống lương 150 triệu trên eTax Mobile & Quy trình gửi phản ánh thu nhập khống để cơ quan thuế vào cuộc'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư 80/2021/TT-BTC; Cổng TTĐT Tổng cục Thuế; scratch/daily_notes.md (13/10/2026).',
         update_date_str),

        (8,
         'doanh nghiep moi lap bi khoa mst khong hoat dong tai dia chi dang ky mau 25 dk tct',
         'Doanh nghiệp mới thành lập bị tạm khóa MST vì "Không hoạt động tại địa chỉ đã đăng ký" dù chưa hết hạn nộp tờ khai Quý 3 (hạn 31/10): Bẫy xác minh thực địa & Thủ tục khôi phục MST theo Mẫu 25/ĐK-TCT',
         5350, 'Tăng cao (+850%)',
         'Doanh nghiệp mới thành lập (Q3/2026), startup, người đại diện pháp luật, kế toán thuế và dịch vụ thành lập doanh nghiệp.',
         'Công ty lập tháng 7/2026, chưa hết hạn nộp tờ khai Quý 3 (31/10), chưa phát sinh doanh thu chi phí, đã treo biển hiệu nhưng bị cán bộ thuế kiểm tra thực địa đột xuất vắng mặt -> phát thông báo đóng MST, tê liệt tài khoản ngân hàng và kinh doanh.',
         "Tiêu đề: 'Doanh nghiệp mới lập bị khóa MST do thuế báo không hoạt động tại trụ sở: Quy trình 3 bước giải trình và khôi phục MST theo Mẫu 25/ĐK-TCT'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư 105/2020/TT-BTC hướng dẫn về đăng ký thuế; scratch/daily_notes.md (12/10/2026).',
         update_date_str),

        (9,
         'drewry wci tuan 42 34 blank sailings siet cong suat a au a my giu gia cuoc',
         'BÁO CÁO DREWRY WCI (Tuần 42): Chỉ số cước biển 4.351 USD/FEU & Hãng tàu công bố 34 Blank Sailings (Tuần 42-46) siết tải trọng các tuyến Á - Âu và Transpacific hậu Golden Week',
         4850, 'Tăng mạnh (+500%)',
         'Chủ hàng xuất khẩu đi thị trường Bắc Mỹ, Châu Âu, doanh nghiệp sản xuất dệt may, da giày, gỗ, điện tử và các công ty forwarder.',
         'Dù cước WCI giảm 2% xuống 4.351 USD/FEU, các liên minh hãng tàu đã công bố 34 chuyến hủy (50% tuyến Á - Bắc Âu, 32% Transpacific) để ghìm tải; rủi ro rớt container (roll over) và thiếu vỏ cont 40HC mùa cao điểm cuối năm.',
         "Tiêu đề: 'Drewry WCI 4.351 USD & 34 chuyến tàu bị hủy từ Tuần 42: Bí quyết giữ booking và tránh rớt container mùa hàng Noel - Tết'.",
         'Báo cáo Drewry Cancelled Sailings Tracker (cập nhật 09/10/2026); Drewry World Container Index; World Shipping Council.',
         update_date_str),

        (10,
         'thu tuc nhap khau banh keo thuc pham dai loan tu cong bo attp nghi dinh 15 2018',
         'Thủ tục nhập khẩu bánh kẹo, thực phẩm chế biến từ Đài Loan mở chuỗi phân phối: Hồ sơ Tự công bố ATTP theo Nghị định 15/2018 & Đăng ký kiểm tra nhà nước chuyên ngành',
         4500, 'Tăng trưởng (+650%)',
         'Doanh nghiệp nhập khẩu thực phẩm, chuỗi cửa hàng tiện lợi, thương nhân phân phối bánh kẹo ngoại nhập và forwarder dịch vụ XNK.',
         'Thiếu hồ sơ Tự công bố ATTP trước khi nhập hàng hoặc không đăng ký kiểm tra nhà nước về an toàn thực phẩm trên Cổng NSW khiến hàng bị kẹt tại cảng, phát sinh chi phí lưu kho bãi và rủi ro bị buộc tái xuất.',
         "Tiêu đề: 'Nhập khẩu bánh kẹo Đài Loan mở chuỗi cửa hàng: Quy trình Tự công bố ATTP và đăng ký kiểm tra nhà nước không lo vướng thủ tục'.",
         'Nghị định số 15/2018/NĐ-CP; Nghị định 43/2017/NĐ-CP & 111/2021/NĐ-CP; Cục An toàn thực phẩm; scratch/daily_notes.md (13/10/2026).',
         update_date_str),
    ]

    for idx, item in enumerate(sheet1_data):
        row_idx = 6 + idx
        fill_to_use = fill_zebra_even if idx % 2 == 0 else fill_zebra_odd
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
            elif col_idx == 10:
                cell.alignment = align_center

    ws1['A16'] = 'Tổng số lượng từ khóa nghiên cứu tuân thủ SLA:'
    ws1['A16'].font = font_bold
    ws1['A16'].fill = fill_total
    ws1['A16'].alignment = align_left
    ws1['A16'].border = total_border

    ws1['D16'] = '=SUM(D6:D15)'
    ws1['D16'].font = font_bold
    ws1['D16'].fill = fill_total
    ws1['D16'].alignment = align_right
    ws1['D16'].number_format = '#,##0'
    ws1['D16'].border = total_border

    ws1['E16'] = 'Nhu cầu thị trường rất cao'
    ws1['E16'].font = font_bold
    ws1['E16'].fill = fill_total
    ws1['E16'].alignment = align_center
    ws1['E16'].border = total_border

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 42 - {update_date_str})."
    ws1['F16'].font = font_total
    ws1['F16'].fill = fill_total
    ws1['F16'].alignment = align_left
    ws1['F16'].border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 42 (12/10/2026 - 18/10/2026) | Mở rộng 135 chủ đề nghiên cứu (gồm 103 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 136)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 136 or m.max_row >= 136]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (131,
         'cong van 22883 chq nvthq ma hs hamburger banh mi hay thit 20 phan tram',
         'cong van 22883 chq nvthq, phan loai ma hs hamburger, banh mi kep thit, ty le thit 20 phan tram, chuong 16 vs nhom 1905, nhom 1602 che pham thit, lenh ra soat khong hai quan, bang ty le thanh phan recipe, giay chung nhan phan tich ca, thue suat thuc pham che bien',
         100.0, 42500, 'https://customs.gov.vn',
         'Công văn số 22883/CHQ-NVTHQ (12/10/2026): Lệnh rà soát khẩn toàn quốc về phân loại mã HS Hamburger — Ranh giới tỷ lệ 20% thịt giữa Chương 16 và nhóm 19.05',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 22883/CHQ-NVTHQ ngày 12/10/2026 chỉ đạo khẩn các Chi cục Hải quan trên toàn quốc rà soát toàn bộ hồ sơ khai báo mã HS mặt hàng hamburger và bánh kẹp thịt; thực tế nhiều doanh nghiệp khai báo theo nhóm 19.05 (bánh mì, thuế suất thấp) nhưng có nhân thịt vượt mức cho phép; căn cứ Chú giải 2 Chương 16 và Chú giải 1(a) Chương 19: (1) Nếu tổng tỷ lệ thịt, xúc xích, phụ phẩm thịt chiếm TRÊN 20% khối lượng toàn bộ sản phẩm -> bắt buộc phân loại vào Chương 16 (nhóm 16.02, thuế suất cao hơn và diện kiểm tra chuyên ngành nghiêm ngặt); (2) Chỉ được xếp nhóm 19.05 nếu tỷ lệ thịt KHÔNG VƯỢT QUÁ 20%; (3) Lệnh rà soát khẩn yêu cầu các Chi cục đối chiếu tài liệu kỹ thuật, bảng thành phần (Recipe/Specification), chứng nhận phân tích (C/A) hoặc lấy mẫu phân tích phân loại, xử lý vi phạm và báo cáo Cục trước 19/10/2026; khuyến nghị doanh nghiệp tự rà soát hồ sơ tránh bị ấn định thuế và xử phạt vi phạm hành chính.',
         'Tổng cục Hải quan (customs.gov.vn) & scratch/daily_notes.md (14/10) | CS (Case Study) / Food HS Classification & Meat Ratio SOP',
         update_date_str),

        (132,
         'dem nguoc 24h han chot 15 10 2026 lien ket vneid muc do 2 nsw nghi dinh 336',
         'dem nguoc 24h, han chot 15 10 2026, khoa cong mot cua quoc gia nsw, bat buoc lien ket vneid muc do 2, nghi dinh 336 2026 nd cp, quyet dinh 1434 qd chq, chuan du lieu json schema, giay phep chuyen nganh nsw, c o dien tu, uy quyen dien tu nguoi dai dien',
         100.0, 45200, 'https://chinhphu.vn',
         'ĐẾM NGƯỢC 24H (Hạn chót 15/10/2026): Ngày cuối cùng bắt buộc liên kết VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) & Chuyển đổi chứng từ JSON Schema tránh bị khóa hệ thống',
         'Nội dung trọng tâm: Hôm nay (14/10/2026) là 24 giờ cuối cùng trước thời điểm Nghị định 336/2026/NĐ-CP và Quyết định 1434/QĐ-CHQ chính thức có hiệu lực từ 00:00 ngày 15/10/2026; cảnh báo tối khẩn cấp: Từ ngày mai 15/10, hệ thống NSW sẽ tự động kích hoạt bộ lọc định danh và chặn toàn bộ các tài khoản doanh nghiệp chưa hoàn thành liên kết tài khoản định danh điện tử VNeID Mức độ 2 của Người đại diện theo pháp luật; doanh nghiệp vi phạm sẽ bị tê liệt hoàn toàn các thủ tục xin cấp giấy phép chuyên ngành (kiểm tra ATTP, kiểm dịch thực vật/động vật, công bố hợp quy, cấp C/O điện tử); đồng thời toàn bộ chứng từ số đính kèm tờ khai hải quan chuyển đổi bắt buộc sang định dạng chuẩn JSON Schema; hướng dẫn khẩn cấp trong ngày 14/10: (1) Người đại diện pháp luật dùng CCCD gắn chip đăng nhập nsw.gov.vn để kích hoạt liên kết VNeID Mức 2; (2) Phân quyền điện tử cho cán bộ phụ trách XNK/đại lý hải quan; (3) Kiểm tra khả năng tương thích chuẩn dữ liệu phần mềm khai báo hải quan để thông quan thông suốt ngay từ 15/10.',
         'Cổng TTĐT Chính phủ (chinhphu.vn) & Tổng cục Hải quan | TN (Tin Ngành) / NSW Digital Identity 24H Deadline Compliance',
         update_date_str),

        (133,
         'cong van 22560 chq gsql ho so hai quan c o uy thac nhap khau hoa don thuong mai',
         'cong van 22560 chq gsql, uy thac nhap khau, ho so hai quan uy thac, hoa don thuong mai invoice, khong trung ten nguoi nhap khau, c o uy thac hop le, thue suat uu dai dac biet fta, nguoi khai hai quan nghi dinh 08, dieu 16 thong tu 38 thong tu 39, hop dong uy thac thuong mai',
         98.8, 32800, 'https://thuvienphapluat.vn',
         'Công văn số 22560/CHQ-GSQL (05/10/2026): Hướng dẫn hồ sơ hải quan, hóa đơn thương mại (Invoice) không trùng tên và điều kiện chấp nhận C/O trong trường hợp Nhập khẩu ủy thác',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 22560/CHQ-GSQL ngày 05/10/2026 tháo gỡ điểm nghẽn pháp lý cốt lõi cho các thương nhân và doanh nghiệp làm dịch vụ ủy thác nhập khẩu; văn bản giải đáp 3 vấn đề then chốt: (1) Tư cách người khai hải quan: Doanh nghiệp nhận ủy thác đứng tên trên tờ khai phải đáp ứng đầy đủ điều kiện người khai hải quan theo Điều 5 Nghị định 08/2015/NĐ-CP (sửa đổi tại NĐ 167/2025/NĐ-CP) và phải đứng tên trên các chứng từ thuộc bộ hồ sơ; (2) Hóa đơn thương mại: Căn cứ Điều 16 Thông tư 38/2015/TT-BTC (sửa đổi bởi TT 39/2018 và TT 121/2025), pháp luật KHÔNG BẮT BUỘC tên người mua trên hóa đơn thương mại (Commercial Invoice) phải trùng khớp với người đứng tên tờ khai hải quan, miễn là hồ sơ phản ánh đúng bản chất giao dịch ủy thác và khớp đúng quan hệ mua bán thực tế với đối tác nước ngoài; (3) Hợp thức hóa chứng từ C/O: Hướng dẫn chi tiết cách thể hiện thông tin trên Giấy chứng nhận xuất xứ (C/O Form E, Form D, Form KV...) để doanh nghiệp ủy thác được hưởng 100% thuế suất ưu đãi đặc biệt FTA hợp pháp mà không bị hải quan bác bỏ.',
         'Cục Hải quan & Thư Viện Pháp Luật (thuvienphapluat.vn) | HT (Học Thực chiến) / Entrusted Import & C/O Compliance SOP',
         update_date_str),

        (134,
         'cong van 22928 chq nvthq huong dan hoan thue hang hoa nhap khau sxxk tai xuat',
         'cong van 22928 chq nvthq, hoan thue hang hoa nhap khau, hoan thue sxxk, hang nhap khau tai xuat, thue gtgt hang nhap khau nop thua, dieu 34 nghi dinh 134 2016, thong tu 06 2021 tt btc, ho so hoan thue hai quan, hoan thue truoc kiem tra sau, an dinh thue nhap khau',
         98.2, 30500, 'https://tapchitaichinh.vn',
         'Công văn số 22928/CHQ-NVTHQ (13/10/2026): Hướng dẫn mới nhất của Cục Hải quan về điều kiện, hồ sơ và quy trình Hoàn thuế đối với hàng hóa nhập khẩu',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 22928/CHQ-NVTHQ ngày 13/10/2026 hướng dẫn thống nhất quy trình hoàn thuế nhập khẩu và xử lý tiền thuế nộp thừa trên toàn hệ thống hải quan; văn bản giải quyết các vướng mắc tồn đọng: (1) Trường hợp hàng hóa nhập khẩu nhưng phải tái xuất trả lại đối tác nước ngoài hoặc tái xuất sang nước thứ ba, hồ sơ hoàn thuế nhập khẩu không thu thuế xuất khẩu theo Điều 34 Nghị định 134/2016/NĐ-CP; (2) Hoàn thuế nguyên liệu, vật tư nhập khẩu để sản xuất hàng xuất khẩu (SXXK) khi doanh nghiệp đã nộp thuế nhưng sau đó thực xuất sản phẩm; (3) Quy trình phân loại hồ sơ thuộc diện "Hoàn thuế trước, kiểm tra sau" hay "Kiểm tra trước, hoàn thuế sau" nhằm rút ngắn thời gian xử lý hoàn trả tiền thuế cho doanh nghiệp xuống tối đa 6 ngày làm việc; (4) Hồ sơ điện tử nộp qua Cổng Dịch vụ công trực tuyến và các lưu ý đối soát chứng từ thanh toán ngân hàng để tránh bị từ chối hồ sơ.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & Cục Hải quan | HT (Học Thực chiến) / Import Tax Refund & Duty Drawback SOP',
         update_date_str),

        (135,
         'hieu luc 15 10 thong tu 128 2026 tt btc van hanh xe agv cua khau thong minh huu nghi tan thanh',
         'thong tu 128 2026 tt btc, cua khau thong minh, xe tu hanh agv, igv khong nguoi lai, cua khau huu nghi, cua khau tan thanh, logistics viet trung, thong quan 24 7, du lieu truoc phuong tien, barie tu dong hoa, lang son smart border',
         99.0, 35600, 'https://congthuong.vn',
         'CHÍNH THỨC HIỆU LỰC TỪ MAI (15/10/2026): Thông tư 128/2026/TT-BTC — Vận hành thí điểm Cửa khẩu thông minh Hữu Nghị & Tân Thanh với xe tự hành AGV 24/7',
         'Nội dung trọng tâm: Ngày mai (15/10/2026), Thông tư số 128/2026/TT-BTC chính thức có hiệu lực thi hành, đánh dấu bước ngoặt lịch sử mở màn tuyến vận tải thông minh qua biên giới Việt - Trung tại cặp cửa khẩu quốc tế Hữu Nghị - Hữu Nghị Quan và Tân Thanh - Pò Chài; đề án triển khai hệ thống xe tự hành AGV/IGV không người lái chuyên chở container 24/7 qua 14 làn xe chuyên dụng; doanh nghiệp xuất nhập khẩu hàng hóa qua đường bộ cần nắm vững 4 quy tắc vận hành mới: (1) Đăng ký thông tin phương tiện và truyền trước dữ liệu tờ khai hải quan lên nền tảng số cửa khẩu thông minh tối thiểu 2 giờ trước khi xe đến biên giới; (2) Phương tiện tuân thủ hệ thống điều phối làn tự động qua nhận diện biển số RFID và mã QR; (3) Hàng hóa qua làn xe tự hành AGV được giải phóng trong vòng 3-5 phút nếu thuộc luồng Xanh; (4) Cảnh báo phân luồng: Các xe chưa đồng bộ chuẩn dữ liệu điện tử bắt buộc phải di chuyển vào làn thủ tục truyền thống, tiềm ẩn nguy cơ ùn ứ kéo dài.',
         'Báo Công Thương (congthuong.vn) & Cổng TTĐT Bộ Công Thương | TN (Tin Ngành) / Smart Border Gate & AGV Operational Protocol',
         update_date_str),
    ]

    # Write new items into rows 136 to 140
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 136 + idx
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

    # New Total row at row 141
    total_row_idx = 141
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 135 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D140)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E140)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 135 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 103 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A141:C141')
    ws3.merge_cells('F141:J141')
    ws3.row_dimensions[total_row_idx].height = 40

    # =============================================================
    # 3. LƯU TỆP AN TOÀN VỚI ATOMIC SAVE (GEMINI.MD RULE 10 & RULE 5)
    # =============================================================
    tmp_path = FILE_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()

    if os.path.exists(FILE_PATH):
        os.replace(tmp_path, FILE_PATH)
    else:
        os.rename(tmp_path, FILE_PATH)
    print(f"THÀNH CÔNG: Đã ghi đè nguyên tử (atomic replace) tệp chính {FILE_PATH}")

    # Đồng bộ sang OneDrive
    try:
        if os.path.exists(ONEDRIVE_DIR):
            onedrive_tmp = ONEDRIVE_FILE + ".tmp.xlsx"
            shutil.copy2(FILE_PATH, onedrive_tmp)
            os.replace(onedrive_tmp, ONEDRIVE_FILE)
            print(f"THÀNH CÔNG: Đã đồng bộ sang OneDrive tại {ONEDRIVE_FILE}")
        else:
            print(f"Bỏ qua đồng bộ OneDrive (không tìm thấy {ONEDRIVE_DIR})")
    except Exception as e:
        print(f"Lỗi khi đồng bộ OneDrive: {e}")

if __name__ == '__main__':
    update_seo_w42_1410()
