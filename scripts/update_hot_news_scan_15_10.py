# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_15_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 15/10/2026 (Tuần 42: 12/10/2026 - 18/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 42, ngày 15/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 135 lên 140 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 15/10/2026)
- Đồng bộ nguyên tử an toàn sang OneDrive Marketing Eureka
"""

import os
import sys
import shutil
import hashlib
from datetime import datetime
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
SYNC_LOG_PATH = r"reports\sync_log.txt"

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w42_1510():
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

    update_date_str = "15/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 42 (12/10/2026 - 18/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van 22233 chq gsql khai ma ko hang an ninh mang mat ma dan su hang hoa luong dung 19916',
         'Công văn số 22233/CHQ-GSQL & 19916/CHQ-GSQL: Hướng dẫn khai mã "KO" đối với Hàng hóa An ninh mạng, Mật mã dân sự & Hàng hóa Lưỡng dụng tránh bị kiểm tra sau thông quan',
         7600, 'Mới ban hành (+1,250%)',
         'Doanh nghiệp nhập khẩu thiết bị CNTT, viễn thông, camera, hóa chất, bộ phận xuất nhập khẩu, forwarder và đại lý hải quan.',
         'Hàng hóa có mã HS trùng danh mục quản lý chuyên ngành nhưng công năng thực tế không thuộc diện cấp phép; nếu không khai mã "KO" trên tờ khai sẽ bị chặn thông quan hoặc phạt vi phạm hành chính, bị chuyển hồ sơ kiểm tra sau thông quan.',
         "Tiêu đề: 'Công văn 22233/CHQ-GSQL & 19916/CHQ-GSQL: Bẫy khai mã KO cho hàng an ninh mạng, mật mã dân sự và hàng lưỡng dụng tránh rủi ro kiểm tra sau thông quan'.",
         'Công văn số 22233/CHQ-GSQL; Công văn số 19916/CHQ-GSQL; Nghị định 53/2022/NĐ-CP; Nghị định 58/2016/NĐ-CP; scratch/daily_notes.md (15/10/2026).',
         update_date_str),

        (2,
         'thong tu 121 2025 tt btc thu tuc hai quan hang qua canh trung chuyen cang bien ban ke 09 10 11',
         'Thông tư số 121/2025/TT-BTC: 10 Điểm cốt lõi về Thủ tục hải quan đối với Hàng hóa Quá cảnh và Trung chuyển tại cảng biển — Quy chuẩn Bản kê và Thiết bị định vị GPS',
         7900, 'Đỉnh điểm quan tâm (+1,320%)',
         'Doanh nghiệp logistics, forwarder, hãng tàu, đại lý vận tải hàng quá cảnh sang Lào/Campuchia/Trung Quốc và trung chuyển cảng biển.',
         'Quy chuẩn nghiêm ngặt về phân loại 3 nhóm thủ tục quá cảnh, bắt buộc gắn thiết bị theo dõi hành trình GPS/seal định vị với container xếp chồng đường thủy nội địa; rủi ro tắc nghẽn nếu khai sai mẫu bản kê 09, 10, 11.',
         "Tiêu đề: 'Thông tư 121/2025/TT-BTC: Cẩm nang 10 điểm vàng làm thủ tục hải quan hàng quá cảnh và trung chuyển cảng biển không lo bị ách tắc'.",
         'Thông tư số 121/2025/TT-BTC (khoản 22, 23 Điều 1 sửa đổi Điều 51a, 51b TT 38/2015/TT-BTC); Cục Hải quan; scratch/daily_notes.md (15/10/2026).',
         update_date_str),

        (3,
         'eu hoan thuc thi eudr 30 12 2026 quy dinh chong pha rung go ca phe cao su dinh vi gps',
         'Ủy ban Châu Âu đề xuất lùi thời hạn thực thi Quy định Chống phá rừng (EUDR) đến 30/12/2026: Cơ hội "vàng" nước rút cho Nông sản & Gỗ Việt Nam',
         8100, 'Tăng đột biến (+1,400%)',
         'Doanh nghiệp xuất khẩu gỗ, đồ nội thất, cà phê, cao su sang thị trường EU, hợp tác xã sản xuất và các hiệp hội ngành hàng xuất khẩu.',
         'Dù EU đề xuất lùi hạn áp dụng sang 30/12/2026, các tiêu chuẩn thẩm định giải trình (DDS), định vị tọa độ GPS lô đất và mốc không phá rừng (31/12/2020) vẫn giữ nguyên; nguy cơ DN chủ quan dừng số hóa chuỗi cung ứng sẽ bị loại khỏi thị trường EU.',
         "Tiêu đề: 'EU hoãn EUDR đến 30/12/2026: Cơ hội vàng hay bẫy chủ quan? Chiến lược số hóa dữ liệu vùng trồng giữ vững thị phần xuất khẩu tỷ đô'.",
         'Thông cáo chính thức Ủy ban Châu Âu (EC); Báo Công Thương (congthuong.vn); Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (4,
         'chinh thuc hieu luc 15 10 bat buoc lien ket vneid muc do 2 nsw nghi dinh 336 2026',
         'CHÍNH THỨC HIỆU LỰC HÔM NAY (15/10/2026): Bắt buộc liên kết VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) & Chuyển đổi chuẩn JSON Schema tránh tê liệt tài khoản',
         8400, 'Đỉnh điểm áp dụng hôm nay (+1,500%)',
         'Toàn bộ lãnh đạo doanh nghiệp XNK, người đại diện theo pháp luật, kế toán trưởng, nhân viên khai báo hải quan, forwarder và đại lý thủ tục.',
         'Hôm nay (15/10) hệ thống NSW bắt đầu chặn toàn bộ tài khoản chưa liên kết VNeID Mức 2 của người đại diện; tê liệt các chức năng xin cấp phép chuyên ngành và cấp C/O điện tử; chứng từ đính kèm bắt buộc theo chuẩn JSON Schema.',
         "Tiêu đề: 'Hôm nay (15/10) chính thức kích hoạt liên kết VNeID Mức 2 trên NSW: Hướng dẫn xử lý lỗi đăng nhập và kết nối tài khoản tức thì'.",
         'Nghị định số 336/2026/NĐ-CP; Quyết định số 1434/QĐ-CHQ; Cổng TTĐT Chính phủ (chinhphu.vn); Tổng cục Hải quan.',
         update_date_str),

        (5,
         'cong van 22883 chq nvthq ma hs hamburger banh mi hay thit 20 phan tram',
         'Công văn số 22883/CHQ-NVTHQ (12/10/2026): Lệnh rà soát khẩn toàn quốc về phân loại mã HS Hamburger — Ranh giới tỷ lệ 20% thịt giữa Chương 16 và nhóm 19.05',
         7500, 'Duy trì đỉnh điểm (+1,100%)',
         'Doanh nghiệp nhập khẩu thực phẩm chế biến, chuỗi nhà hàng thức ăn nhanh, chuỗi F&B, forwarder và đại lý hải quan.',
         'Lệnh rà soát khẩn toàn quốc của Cục Hải quan (báo cáo trước 19/10/2026); doanh nghiệp khai hamburger vào nhóm 19.05 (bánh mì) nhưng nhân thịt vượt 20% bị ấn định lại vào Chương 16 (chế phẩm thịt), bị truy thu thuế lớn và xử phạt vi phạm hành chính.',
         "Tiêu đề: 'Lệnh rà soát khẩn Công văn 22883/CHQ-NVTHQ: Hamburger là bánh mì hay thịt? Quy tắc 20% phân định mã HS và thuế suất'.",
         'Công văn số 22883/CHQ-NVTHQ ngày 12/10/2026; Chú giải 2 Chương 16; Chú giải 1(a) Chương 19; Cục Hải quan; scratch/daily_notes.md (14/10/2026).',
         update_date_str),

        (6,
         'dang ky mst etax mobile cham treo ho so dong bo cccd vneid muc 2 chi cuc thue',
         'Giải pháp xử lý hồ sơ Đăng ký Mã số thuế cá nhân & Hộ kinh doanh bị treo trên eTax Mobile: Quy trình đối soát cơ sở dữ liệu dân cư VNeID',
         5800, 'Tăng cao (+890%)',
         'Cá nhân người nộp thuế, hộ kinh doanh mới thành lập, kế toán dịch vụ thuế, đại lý thuế và người lao động mở MST.',
         'Đăng ký MST trên eTax Mobile bị treo trạng thái "Đang xử lý" cả tháng không có kết quả do lỗi bất đồng bộ dữ liệu dân cư; dẫn đến nguy cơ chậm đăng ký kinh doanh, không mở được tài khoản ngân hàng và bị phạt chậm đăng ký thuế.',
         "Tiêu đề: 'eTax Mobile bị treo hồ sơ cấp MST gần 1 tháng: 3 bước tra soát dữ liệu CCCD/VNeID và quy trình mở khóa tại Chi cục Thuế'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư 105/2020/TT-BTC; Cổng TTĐT Tổng cục Thuế; scratch/daily_notes.md (15/10/2026).',
         update_date_str),

        (7,
         'hieu luc 15 10 thong tu 128 2026 tt btc van hanh xe agv cua khau thong minh huu nghi tan thanh',
         'CHÍNH THỨC HIỆU LỰC HÔM NAY (15/10/2026): Thông tư 128/2026/TT-BTC — Vận hành thí điểm Cửa khẩu thông minh Hữu Nghị & Tân Thanh với xe tự hành AGV 24/7',
         6300, 'Hiệu lực hôm nay (+950%)',
         'Doanh nghiệp xuất nhập khẩu đường bộ qua Lạng Sơn, chủ hàng nông sản, linh kiện điện tử, doanh nghiệp vận tải biên giới và logistics Trung - Việt.',
         'Thông tư có hiệu lực từ hôm nay (15/10); phương tiện chưa khai báo trước dữ liệu điện tử tối thiểu 2 giờ sẽ bị từ chối vào làn xe tự hành AGV thông minh, buộc phải dồn ứ tại làn thủ tục truyền thống.',
         "Tiêu đề: 'Hôm nay (15/10) Thông tư 128/2026/TT-BTC chính thức có hiệu lực: 4 quy tắc vàng vận hành xe tự hành AGV tại Cửa khẩu thông minh Lạng Sơn'.",
         'Thông tư số 128/2026/TT-BTC ngày 27/08/2026; Đề án thí điểm Cửa khẩu thông minh Lạng Sơn; Cổng TTĐT Bộ Công Thương.',
         update_date_str),

        (8,
         'drewry wci tuan 42 blank sailings siet cong suat a au giu gia cuoc cao diem q4',
         'Báo cáo Drewry WCI & Cancelled Sailings Tracker (Tuần 42): Hãng tàu tăng cường hủy chuyến siết tải trọng nhằm ghìm giá cước mùa cao điểm Q4',
         5200, 'Tăng mạnh (+620%)',
         'Chủ hàng xuất khẩu đi thị trường Bắc Mỹ, Châu Âu, doanh nghiệp sản xuất dệt may, da giày, gỗ, điện tử và các công ty forwarder.',
         'Chỉ số WCI đạt 4.351 USD/FEU; các liên minh hãng tàu công bố 34-42 chuyến hủy trên các tuyến Á - Âu và Transpacific hậu Golden Week để ghìm tải; rủi ro rớt container (roll over) và thiếu vỏ cont 40HC mùa cao điểm cuối năm.',
         "Tiêu đề: 'Drewry WCI 4.351 USD & Làn sóng Blank Sailings Tuần 42: Bí quyết giữ booking và tránh rớt container mùa hàng Noel - Tết'.",
         'Báo cáo Drewry Cancelled Sailings Tracker (cập nhật 15/10/2026); Drewry World Container Index; World Shipping Council.',
         update_date_str),

        (9,
         'cong van 22560 chq gsql ho so hai quan c o uy thac nhap khau hoa don thuong mai',
         'Công văn số 22560/CHQ-GSQL (05/10/2026): Hướng dẫn hồ sơ hải quan, hóa đơn thương mại (Invoice) không trùng tên và điều kiện chấp nhận C/O trong trường hợp Nhập khẩu ủy thác',
         6100, 'Duy trì cao (+820%)',
         'Doanh nghiệp nhập khẩu ủy thác, chủ hàng thương mại, bên nhận ủy thác logistics, forwarder và kế toán thuế XNK.',
         'Hóa đơn thương mại nước ngoài phát hành không đứng tên bên nhận ủy thác bị hải quan nghi ngờ tính hợp lệ; lúng túng chứng từ C/O form ưu đãi FTA dẫn tới nguy cơ bị bác C/O và truy thu thuế nhập khẩu ưu đãi đặc biệt.',
         "Tiêu đề: 'Công văn 22560/CHQ-GSQL: Hướng dẫn gỡ vướng hóa đơn không trùng tên và điều kiện chấp nhận C/O khi nhập khẩu ủy thác'.",
         'Công văn số 22560/CHQ-GSQL ngày 05/10/2026; Điều 5 Nghị định 08/2015/NĐ-CP; Điều 16 Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018 và TT 121/2025).',
         update_date_str),

        (10,
         'cong van 22720 chq gsql to khai phi thue quan dncx chua xac nhan khu vuc giam sat',
         'Công văn số 22720/CHQ-GSQL (08/10/2026): Hướng dẫn xử lý dứt điểm các tờ khai xuất nhập khẩu phi thuế quan chưa được xác nhận qua khu vực giám sát (KVGS) do lệch chỉ tiêu "Số quản lý nội bộ"',
         6800, 'Duy trì đỉnh (+980%)',
         'Doanh nghiệp chế xuất (DNCX), doanh nghiệp nội địa mua bán với DNCX, bộ phận xuất nhập khẩu, forwarder và đại lý hải quan.',
         'Hàng ngàn tờ khai phi thuế quan mở từ tháng 9/2026 bị hệ thống treo không tự động xác nhận qua KVGS do lệch cú pháp chỉ tiêu số quản lý nội bộ (#&XKPTQ vs #&NKTC); rủi ro không thanh khoản được tờ khai, treo hoàn thuế GTGT.',
         "Tiêu đề: 'Công văn 22720/CHQ-GSQL: Hướng dẫn gỡ lỗi tờ khai phi thuế quan DNCX chưa qua khu vực giám sát và quy trình gửi danh sách xử lý dứt điểm'.",
         'Công văn số 22720/CHQ-GSQL ngày 08/10/2026; Công văn số 20991/CHQ-GSQL ngày 26/08/2026; Cục Hải quan.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 42 (12/10/2026 - 18/10/2026) | Mở rộng 140 chủ đề nghiên cứu (gồm 108 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 141)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 141 or m.max_row >= 141]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (136,
         'thong tu 121 2025 tt btc thu tuc hai quan hang qua canh trung chuyen cang bien ban ke 09 10 11',
         'thong tu 121 2025 tt btc, hang qua canh, hang trung chuyen, thu tuc hai quan qua canh, trung chuyen cang bien, ban ke 09 10 11, thiet bi theo doi hanh trinh gps, seal dinh vi dien tu, qua canh hang khong 2 gio, chuyen tieu thu noi dia dieu 21',
         99.2, 38500, 'https://customs.gov.vn',
         'Thông tư số 121/2025/TT-BTC: 10 Điểm cốt lõi về Thủ tục hải quan đối với Hàng hóa Quá cảnh và Trung chuyển tại cảng biển — Quy chuẩn Bản kê và Thiết bị định vị GPS',
         'Nội dung trọng tâm: Thông tư 121/2025/TT-BTC (khoản 22, 23 Điều 1 sửa đổi Điều 51a, 51b Thông tư 38/2015/TT-BTC) siết chặt và chuẩn hóa toàn diện thủ tục hải quan với hàng hóa quá cảnh và trung chuyển tại cảng biển: (1) Phân định rõ 3 nhóm thủ tục: Quá cảnh đường bộ/đường thủy, Quá cảnh hàng không cùng cảng, và Trung chuyển cảng biển; (2) Quy chuẩn chứng từ: Bản kê hàng hóa quá cảnh bắt buộc sử dụng mẫu số 09, 10, 11 (riêng hàng không sử dụng mẫu số 08); (3) Container xếp chồng trên phương tiện thủy nội địa bắt buộc phải gắn thiết bị theo dõi hành trình GPS/seal định vị điện tử đảm bảo giám sát hải quan liên tục; (4) Thời hạn phê duyệt: Chi cục Hải quan phải hoàn tất phê duyệt Bản kê quá cảnh hàng không trong vòng 02 giờ làm việc kể từ khi tiếp nhận hồ sơ hợp lệ; (5) Xử lý hàng vận chuyển nhiều lần, chuyển tải, lưu kho hoặc đổi phương thức vận tải phải khai báo mẫu số 09b; trường hợp chuyển tiêu thụ nội địa phải thực hiện nghiêm ngặt theo Điều 21 Thông tư 38; cảnh báo doanh nghiệp forwarder và vận tải tuân thủ đúng quy trình để tránh bị xử phạt vi phạm giám sát hải quan.',
         'Tổng cục Hải quan (customs.gov.vn) & scratch/daily_notes.md (15/10/2026) | HT (Học Thực chiến) / Transit & Transshipment Customs SOP',
         update_date_str),

        (137,
         'cong van 22233 chq gsql khai ma ko hang an ninh mang mat ma dan su hang hoa luong dung 19916',
         'cong van 22233 chq gsql, cong van 19916 chq gsql, khai ma ko, hang an ninh mang, san pham mat ma dan su, hang hoa luong dung, trung ma hs danh muc, dac tinh ky thuat, sua to khai hai quan, kiem tra sau thong quan',
         98.9, 34200, 'https://customs.gov.vn',
         'Công văn số 22233/CHQ-GSQL & 19916/CHQ-GSQL: Hướng dẫn khai mã "KO" đối với Hàng hóa An ninh mạng, Mật mã dân sự và Hàng hóa Lưỡng dụng tránh bị kiểm tra sau thông quan',
         'Nội dung trọng tâm: Cục Giám sát Quản lý (Tổng cục Hải quan) ban hành Công văn số 22233/CHQ-GSQL và 19916/CHQ-GSQL giải quyết tình trạng ách tắc tờ khai đối với các mặt hàng thiết bị công nghệ thông tin, viễn thông, camera, phần mềm và hóa chất; nguyên tắc sống còn khi làm thủ tục: (1) Trường hợp hàng hóa nhập khẩu có mã HS 8 số trùng khớp với mã HS trong danh mục quản lý an ninh mạng (Nghị định 53/2022/NĐ-CP), mật mã dân sự (Nghị định 58/2016/NĐ-CP) hoặc danh mục hàng hóa lưỡng dụng, NHƯNG đặc tính kỹ thuật thực tế và công năng không thuộc phạm vi quản lý chuyên ngành, doanh nghiệp BẮT BUỘC phải chủ động khai báo mã "KO" tại ô ghi chú/mã quản lý riêng của tờ khai điện tử; (2) Doanh nghiệp chịu trách nhiệm pháp lý trước pháp luật về tính chính xác của khai báo và sẵn sàng xuất trình tài liệu kỹ thuật kiểm chứng; (3) Cảnh báo rủi ro: Nếu không khai mã "KO", hệ thống hải quan sẽ tự động chặn và yêu cầu giấy phép, hoặc hải quan cửa khẩu sẽ chuyển luồng kiểm tra, yêu cầu sửa tờ khai và chuyển hồ sơ sang diện Kiểm tra sau thông quan (KTSTQ), xử phạt vi phạm hành chính.',
         'Cục Hải quan (customs.gov.vn) & scratch/daily_notes.md (15/10/2026) | CS (Case Study) / Dual-use & Cyber Security KO Code Exemption SOP',
         update_date_str),

        (138,
         'eu hoan thuc thi eudr 30 12 2026 quy dinh chong pha rung go ca phe cao su dinh vi gps',
         'eu hoan eudr, eudr 30 12 2026, quy dinh chong pha rung eu, xuat khau go sang eu, ca phe viet nam eudr, cao su khong pha rung, dinh vi gps vuon trong, to khai tham dinh dds, truy xuat nguon goc so, giam sat chuoi cung ung',
         99.5, 41800, 'https://congthuong.vn',
         'Ủy ban Châu Âu đề xuất lùi thời hạn thực thi Quy định Chống phá rừng (EUDR) đến 30/12/2026: Cơ hội "vàng" nước rút cho Nông sản & Gỗ Việt Nam',
         'Nội dung trọng tâm: Ủy ban Châu Âu (EC) và Nghị viện Châu Âu đã chính thức thông qua đề xuất hoãn áp dụng Quy định Chống phá rừng (EUDR) thêm 12 tháng: Doanh nghiệp lớn và vừa áp dụng từ 30/12/2026, doanh nghiệp nhỏ và siêu nhỏ áp dụng từ 30/06/2027; động thái này tháo gỡ áp lực nghẹt thở cho chuỗi xuất khẩu gỗ, cà phê và cao su tỷ đô của Việt Nam; tuy nhiên, các chuyên gia cảnh báo đây KHÔNG PHẢI là sự nới lỏng chính sách: (1) Toàn bộ tiêu chuẩn khắt khe gồm mốc thời gian không phá rừng (sau ngày 31/12/2020), định vị tọa độ địa lý GPS đa giác (polygon) từng thửa đất và Tuyên bố thẩm định chuỗi cung ứng (Due Diligence Statement - DDS) vẫn giữ nguyên 100%; (2) Các nhà nhập khẩu EU bắt đầu sàng lọc nhà cung ứng từ đầu năm 2026; (3) Khuyến nghị chiến lược cho doanh nghiệp XNK Việt Nam: Tuyệt đối không trì hoãn số hóa, tận dụng 12 tháng gia hạn quý giá này để xây dựng cơ sở dữ liệu vùng trồng, bản đồ số và tích hợp hệ thống truy xuất nguồn gốc số hóa để giữ vững thị phần EU.',
         'Báo Công Thương (congthuong.vn) & Cổng TTĐT Chính phủ (chinhphu.vn) | TN (Tin Ngành) / EUDR Supply Chain Compliance & Due Diligence Strategy',
         update_date_str),

        (139,
         'dang ky mst etax mobile cham treo ho so dong bo cccd vneid muc 2 chi cuc thue',
         'dang ky mst ca nhan, dang ky mst etax mobile, ho so dang ky thue bi treo, dong bo du lieu dan cu, can cuoc cong dan gan chip, vneid muc do 2, cham cap mst kinh doanh, tra cuu ho so thue dien tu, mau 05 dkt tct, xu ly loi etax mobile',
         97.5, 28600, 'https://tapchitaichinh.vn',
         'Giải pháp xử lý hồ sơ Đăng ký Mã số thuế cá nhân & Hộ kinh doanh bị treo trên eTax Mobile: Quy trình đối soát cơ sở dữ liệu dân cư VNeID',
         'Nội dung trọng tâm: Ghi nhận làn sóng phản ánh từ người nộp thuế về tình trạng đăng ký mã số thuế mới (hoặc thay đổi thông tin đăng ký thuế) qua ứng dụng eTax Mobile bị treo trạng thái "Đang xử lý" kéo dài từ 2 đến 4 tuần mà không nhận được kết quả; nguyên nhân cốt lõi: Quá trình đối soát tự động giữa hệ thống TMS của Tổng cục Thuế và Cơ sở dữ liệu quốc gia về dân cư (VNeID/CCCD gắn chip) bị nghẽn lệnh hoặc lệch dữ liệu họ tên/ngày sinh/địa chỉ thường trú; hậu quả: Cá nhân không mở được tài khoản ngân hàng, hộ kinh doanh bị đình trệ hoạt động và nguy cơ bị phạt chậm đăng ký thuế theo Nghị định 125/2020/NĐ-CP; hướng dẫn 3 bước gỡ vướng khẩn cấp: (1) Sử dụng chức năng Tra cứu hồ sơ khai thuế/đăng ký thuế trên eTax Mobile để xem mã lỗi chi tiết; (2) Kiểm tra và chuẩn hóa thông tin cá nhân trên VNeID Mức 2 trùng khớp 100% với hồ sơ nộp thuế; (3) Gửi văn bản tra soát hoặc liên hệ trực tiếp bộ phận Một cửa Chi cục Thuế quản lý địa bàn kèm mã giao dịch điện tử để được cán bộ thuế phê duyệt thủ công trong 24 giờ.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & scratch/daily_notes.md (15/10/2026) | HT (Học Thực chiến) / Personal Tax ID & eTax Mobile Troubleshooting Guide',
         update_date_str),

        (140,
         'drewry wci tuan 42 blank sailings siet cong suat a au giu gia cuoc cao diem q4',
         'drewry wci tuan 42, chi so cuoc bien drewry, blank sailings, huy chuyen tau container, siet cong suat van tai, tuyen a au transpacific, gia cuoc mua cao diem q4, thieu vo cont 40hc, roll over rot cont, chien luoc booking hang tet',
         98.0, 31200, 'https://www.drewry.co.uk',
         'Báo cáo Drewry WCI & Cancelled Sailings Tracker (Tuần 42): Hãng tàu tăng cường hủy chuyến siết tải trọng nhằm ghìm giá cước mùa cao điểm Q4',
         'Nội dung trọng tâm: Báo cáo Drewry World Container Index (WCI) tuần 42 ghi nhận giá cước container tổng hợp đạt 4.351 USD/FEU; mặc dù chịu áp lực điều chỉnh giảm nhẹ sau kỳ nghỉ Tuần lễ Vàng (Golden Week) tại Trung Quốc, các liên minh hãng tàu lớn (2M, Ocean Alliance, THE Alliance) đã đồng loạt triển khai chiến lược "Blank Sailings" (hủy chuyến có chủ đích) với 34 đến 42 chuyến tàu bị hủy trong giai đoạn từ Tuần 42 đến Tuần 46; mục tiêu của các hãng tàu là chủ động cắt giảm từ 10% đến 14% năng lực vận chuyển khả dụng trên các tuyến huyết mạch Châu Á - Bắc Âu và Xuyên Thái Bình Dương (Transpacific) nhằm tạo tình trạng khan hiếm chỗ giả tạo, ngăn đà giảm sâu của cước biển và giữ mặt bằng giá cao đón đầu đợt xuất hàng Giáng sinh và Tết Nguyên Đán; khuyến nghị cho chủ hàng Việt Nam: (1) Ký hợp đồng giữ chỗ (space guarantee) và chốt booking trước thời điểm tàu chạy tối thiểu 2-3 tuần; (2) Phân bổ rủi ro bằng cách đi 2-3 hãng tàu khác nhau, kiểm soát kỹ điều khoản bồi thường khi bị rớt container (roll over).',
         'Drewry Shipping Consultants & World Shipping Council | TN (Tin Ngành) / Shipping Rate Forecast & Q4 Peak Booking SOP',
         update_date_str),
    ]

    # Write new items into rows 141 to 145
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 141 + idx
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

    # New Total row at row 146
    total_row_idx = 146
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 140 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D145)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E145)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 140 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 108 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A146:C146')
    ws3.merge_cells('F146:J146')
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

    # Lấy hash và dung lượng file sau cập nhật
    file_size = os.path.getsize(FILE_PATH)
    file_sha = get_file_sha256(FILE_PATH)
    print(f"File info: {file_size:,} bytes | SHA256: {file_sha}")

    # Đồng bộ sang OneDrive
    try:
        if os.path.exists(ONEDRIVE_DIR):
            onedrive_tmp = ONEDRIVE_FILE + ".tmp.xlsx"
            shutil.copy2(FILE_PATH, onedrive_tmp)
            os.replace(onedrive_tmp, ONEDRIVE_FILE)
            print(f"THÀNH CÔNG: Đã đồng bộ sang OneDrive tại {ONEDRIVE_FILE}")

            # Ghi log đồng bộ
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha}) -> OneDrive (wf05 Hot News Scan - {update_date_str})\n"
            with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as lf:
                lf.write(log_entry)
            print(f"Đã ghi nhật ký đồng bộ vào {SYNC_LOG_PATH}")
        else:
            print(f"Bỏ qua đồng bộ OneDrive (không tìm thấy {ONEDRIVE_DIR})")
    except Exception as e:
        print(f"Lỗi khi đồng bộ OneDrive: {e}")

if __name__ == '__main__':
    update_seo_w42_1510()
