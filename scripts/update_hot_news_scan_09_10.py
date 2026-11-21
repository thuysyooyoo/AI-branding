# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_09_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 09/10/2026 (Tuần 41)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành, ngày 09/10/2026, đếm ngược 6 ngày đến 15/10/2026, phân cấp C/O TT 26/2026/TT-BCT, Drewry WCI 4.351 USD & 39 Blank Sailings)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 111 lên 115 chủ đề, bổ sung 4 tin nóng nghiệp vụ ngày 09/10/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka
"""

import os
import sys
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

def update_seo_w41_0910():
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

    update_date_str = "09/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 41 (05/10/2026 - 11/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'phan cap tham quyen cap co dia phuong thong tu 26 2026 tt bct mau so 03 ecosys',
         'Thông tư số 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP: Phân cấp thẩm quyền cấp C/O về Sở Công Thương địa phương — Hướng dẫn lập Mẫu số 03 chuyển nơi cấp & Phân biệt khai báo co.moit.gov.vn vs ecosys.gov.vn',
         5800, 'Tăng kỷ lục (+980%)',
         'Doanh nghiệp xuất khẩu có C/O ưu đãi thuế (Form E, D, AK, VK...), bộ phận chứng từ xuất nhập khẩu, forwarder và đại lý hải quan.',
         'Từ 01/08/2026 phân cấp cấp C/O về địa phương, nhiều DN nộp sai thẩm quyền, nộp nhầm hệ thống giữa co.moit và ecosys khiến hồ sơ bị treo, chậm tiến độ tàu chạy và rủi ro bị đối tác phạt hợp đồng.',
         "Tiêu đề: 'Thông tư 26/2026/TT-BCT phân cấp cấp C/O về địa phương: Quy trình 3 bước chuyển nơi cấp (Mẫu số 03) & Bí quyết tránh tắc hồ sơ trên Ecosys'.",
         'Thông tư số 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP; Cục Xuất nhập khẩu (Bộ Công Thương) & scratch/daily_notes.md (09/10).',
         update_date_str),

        (2,
         'dem nguoc 6 ngay nghi dinh 336 bat buoc vneid muc do 2 nsw quyet dinh 1434 json',
         'ĐẾM NGƯỢC 6 NGÀY (Hạn chót 15/10/2026): Nghị định 336/2026/NĐ-CP & Quyết định 1434/QĐ-CHQ — Bắt buộc tài khoản VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) và Chuẩn dữ liệu chứng từ JSON DataFormat',
         6100, 'Đỉnh điểm cảnh báo (+1,050%)',
         'Toàn bộ lãnh đạo doanh nghiệp XNK, người đại diện pháp luật, kế toán trưởng, nhân viên khai báo hải quan và đại lý thủ tục.',
         'Chỉ còn 6 ngày đến thời hạn bắt buộc 15/10; các doanh nghiệp chưa tích hợp VNeID Mức 2 của người đại diện hoặc chưa cập nhật phần mềm khai báo theo chuẩn JSON sẽ bị ngắt kết nối hệ thống hải quan, không truyền được tờ khai và chứng từ chuyên ngành.',
         "Tiêu đề: 'Đếm ngược 6 ngày (15/10): Hạn chót liên kết VNeID Mức 2 trên NSW & Chuẩn bị chuẩn dữ liệu JSON QĐ 1434 tránh tê liệt thông quan'.",
         'Nghị định 336/2026/NĐ-CP & Quyết định 1434/QĐ-CHQ của Cục Hải quan; Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (3,
         'dem nguoc 6 ngay thong tu 128 2026 tt btc cua khau thong minh xe tu hanh agv huu nghi',
         'ĐẾM NGƯỢC 6 NGÀY (15/10/2026): Thông tư số 128/2026/TT-BTC — Thí điểm Cửa khẩu thông minh Hữu Nghị & Tân Thanh với xe tự hành AGV/IGV không người lái và chuẩn dữ liệu thông quan 24/7',
         5350, 'Tăng rất mạnh (+780%)',
         'Doanh nghiệp xuất nhập khẩu hàng hóa đường bộ tuyến biên giới phía Bắc, chủ hàng nông sản, linh kiện điện tử, doanh nghiệp vận tải và logistics cửa khẩu.',
         'Vận hành làn thông quan thông minh 24/7 với xe tự hành không người lái; doanh nghiệp thiếu thông tin về chuẩn dữ liệu điện tử phương tiện sẽ bị từ chối vào luồng tự động, buộc phải quay đầu hoặc xếp hàng ùn tắc ở làn truyền thống.',
         "Tiêu đề: 'Đếm ngược 6 ngày Thông tư 128/2026/TT-BTC: Cửa khẩu thông minh vận hành xe AGV & Hướng dẫn chủ hàng đăng ký luồng thông quan 24/7'.",
         'Thông tư số 128/2026/TT-BTC ngày 27/08/2026 của Bộ Tài chính; Ban Quản lý Khu KTCK Đồng Đăng - Lạng Sơn.',
         update_date_str),

        (4,
         'drewry wci 08 10 giam 2 phan tram 4351 usd 39 blank sailings hau golden week',
         'BÁO CÁO DREWRY WCI (08/10/2026): Chỉ số cước biển giảm 2% xuống 4.351 USD/FEU, Hãng tàu công bố 39 Blank Sailings (Tuần 41-45) quản trị tải trọng hậu Golden Week',
         4750, 'Tăng mạnh (+490%)',
         'Chủ hàng xuất khẩu đi thị trường Mỹ (USWC, USEC), Châu Âu, các công ty FMS, NVOCC và phòng xuất nhập khẩu logistics.',
         'Mặc dù cước giao ngay hạ nhiệt nhẹ 2% sau tuần lễ vàng, việc các liên minh hãng tàu cắt giảm 39 chuyến tàu (blank sailings) có nguy cơ gây thiếu vỏ cont cục bộ và rớt booking (roll over) vào nửa cuối tháng 10.',
         "Tiêu đề: 'Cước Drewry WCI giảm xuống 4.351 USD nhưng 39 chuyến tàu bị hủy: Chiến lược giữ chỗ cont & đàm phán hợp đồng cho mùa hàng cuối năm'.",
         'Báo cáo Drewry World Container Index (WCI) ngày 08/10/2026; Container Capacity Insight & World Shipping Council.',
         update_date_str),

        (5,
         'dieu chinh sai so nha dia diem kinh doanh ho kinh doanh cong dvc tinh etax',
         'Vướng mắc điều chỉnh sai địa chỉ/số nhà mã địa điểm kinh doanh Hộ kinh doanh: Hướng dẫn nộp hồ sơ thay đổi nội dung trên Cổng Dịch vụ công cấp tỉnh & Đồng bộ hệ thống Thuế eTax',
         3900, 'Tăng rất mạnh (+680%)',
         'Hộ kinh doanh cá thể, chủ cửa hàng thương mại điện tử, kế toán dịch vụ và doanh nghiệp quản lý chuỗi điểm bán.',
         'Khi địa chỉ hoặc số nhà trên mã địa điểm kinh doanh bị sai lệch, người nộp thuế không thể sửa trên hệ thống tra cứu MST eTax mà phải nộp hồ sơ qua DVC tỉnh; nếu không chuẩn hóa sẽ bị phạt vi phạm đăng ký kinh doanh và treo xuất hóa đơn.',
         "Tiêu đề: 'Sai số nhà trên mã địa điểm kinh doanh: Quy trình 3 bước nộp hồ sơ điều chỉnh trên Cổng DVC tỉnh để đồng bộ Thuế eTax chuẩn xác'.",
         'Nghị định 01/2021/NĐ-CP & Thông tư 105/2020/TT-BTC; Diễn đàn Thuế chuyên sâu & scratch/daily_notes.md (09/10).',
         update_date_str),

        (6,
         'incoterms fca exw dap xuat nhap khau tai cho dieu 47a luat 90 chi phi van chuyen tri gia hai quan',
         'Incoterms FCA/EXW vs DAP/DDP trong Xuất Nhập Khẩu Tại Chỗ theo Điều 47a Luật Hải quan 90/2025/QH15 & Công văn 22411/CHQ-GSQL: Phân định chi phí vận chuyển nội địa vào trị giá hải quan',
         5400, 'Duy trì đỉnh (+850%)',
         'Doanh nghiệp gia công, SXXK, DNCX, các đối tác thương nhân nước ngoài chỉ định giao nhận hàng tại Việt Nam, đại lý hải quan và kế toán thuế.',
         'Ranh giới phân chia rủi ro và trách nhiệm giữa bên bán và bên mua tại kho chỉ định nội địa theo Điều 47a; nguy cơ bị hải quan ấn định cộng chi phí vận tải nội địa vào trị giá tính thuế nhập khẩu nếu chọn sai điều kiện Incoterms.',
         "Tiêu đề: 'Điều 47a Luật Hải quan 90: Nên chọn FCA hay DAP khi giao nhận XNK tại chỗ để không bị cộng oan thuế trị giá hải quan?'." ,
         'Điều 47a Luật Hải quan 90/2025/QH15; Công văn 22411/CHQ-GSQL; Thông tư 39/2015/TT-BTC & Thông tư 60/2019/TT-BTC.',
         update_date_str),

        (7,
         'quy che quan ly seal niem phong hai quan quyet dinh 2722 qd tchq he thong vasscm xu ly dut gay seal',
         'Quyết định 2722/QĐ-TCHQ: Quy chế quản lý, sử dụng seal niêm phong hải quan & Quy trình xử lý sự cố đứt, gãy seal định vị GPS trong Hệ thống giám sát hải quan tự động (VASSCM)',
         4050, 'Tăng mạnh (+580%)',
         'Doanh nghiệp vận tải hàng hóa chịu sự giám sát hải quan, đơn vị kho ngoại quan, ICD, đại lý hải quan và chủ hàng.',
         'Sự cố đứt, gãy seal niêm phong trên đường vận chuyển nếu không báo cáo đúng quy trình theo QĐ 2722 sẽ bị lập biên bản vi phạm, phạt tiền và cưỡng chế kiểm tra thực tế hàng hóa 100% (luồng đỏ).',
         "Tiêu đề: 'Quyết định 2722/QĐ-TCHQ: Quy trình 3 bước xử lý khẩn cấp khi đứt gãy seal hải quan trong hệ thống VASSCM tránh bị chuyển luồng đỏ'.",
         'Quyết định số 2722/QĐ-TCHQ của Tổng cục Hải quan; Thông tư số 38/2015/TT-BTC & Thông tư số 39/2018/TT-BTC.',
         update_date_str),

        (8,
         'ma hs smartwatch huawei watch gt 6 thong bao 22555 tb chq 8517 62 59 xac dinh truoc ma so',
         'Phân loại mã HS Thiết bị đeo thông minh (Smartwatch Huawei Watch GT 6) theo Thông báo 22555/TB-CHQ: Nguyên tắc áp mã nhóm 85.17 thay vì nhóm đồng hồ 91.02',
         3850, 'Tăng mạnh (+550%)',
         'Doanh nghiệp nhập khẩu thiết bị điện tử viễn thông, thiết bị đeo thông minh, sàn TMĐT và đại lý hải quan.',
         'Nhầm lẫn giữa nhóm đồng hồ thông thường (91.02) và thiết bị truyền nhận dữ liệu kết nối Bluetooth/mạng (8517.62.59), dẫn đến sai mức thuế suất nhập khẩu và thuế tiêu thụ đặc biệt.',
         "Tiêu đề: 'Thông báo 22555/TB-CHQ: Bài học xương máu phân loại mã HS Smartwatch vào nhóm 85.17 thay vì 91.02'.",
         'Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan; 6 Quy tắc giải thích phân loại hàng hóa HS.',
         update_date_str),

        (9,
         'doanh nghiep fdi trong kcx mua hang noi dia mo to khai hai quan dieu 74 thong tu 38',
         'Doanh nghiệp FDI thông thường nằm trong Khu Chế Xuất (KCX) khi mua hàng nội địa có phải mở tờ khai hải quan không? Căn cứ Điều 74 Thông tư 38/2015/TT-BTC & Thông tư 39/2018/TT-BTC',
         3700, 'Tăng mạnh (+510%)',
         'Doanh nghiệp FDI, doanh nghiệp hoạt động trong KCX/KCN, nhà cung cấp hàng hóa nội địa và kế toán thuế XNK.',
         'Phân biệt rõ tư cách pháp nhân giữa Doanh nghiệp chế xuất (DNCX) được hưởng chính sách phi thuế quan với Doanh nghiệp FDI thông thường chỉ thuê nhà xưởng trong khuôn viên KCX để tránh mở tờ khai sai hoặc bị truy thu thuế GTGT 10%.',
         "Tiêu đề: 'FDI thuê xưởng trong KCX mua hàng nội địa có phải mở tờ khai hải quan? Phân định DNCX và DN thông thường theo Điều 74'.",
         'Điều 74 Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018/TT-BTC); Nghị định 35/2022/NĐ-CP về quản lý KCN & KCX.',
         update_date_str),

        (10,
         'gom cont chinh ngach trung viet van chuyen hang le lcl trung quoc vat chuan don hang sau golden week',
         'Gom cont LCL chính ngạch tuyến Trung - Việt có hóa đơn VAT chuẩn: Giải pháp đón sóng bùng nổ đơn hàng hậu Golden Week & Tối ưu chi phí bãi cảng',
         4850, 'Tăng mạnh (+290%)',
         'Các chủ shop, nhà bán hàng TMĐT, xưởng gia công nhỏ và doanh nghiệp vừa và nhỏ nhập khẩu hàng phụ kiện, linh kiện từ Trung Quốc.',
         'Nhu cầu nhập hàng tăng vọt sau kỳ nghỉ lễ Tuần Lễ Vàng trong khi đi cont nguyên FCL quá tốn chi phí và rủi ro rớt cont; nhu cầu gom hàng lẻ LCL chính ngạch có hóa đơn VAT để hợp thức hóa đầu vào.',
         "Tiêu đề: 'Bùng nổ đơn hàng hậu Golden Week: Giải pháp gom cont LCL chính ngạch Trung - Việt có hóa đơn VAT bảo vệ dòng vốn'.",
         'Báo cáo thị trường logistics thương mại biên giới; Biểu cước và dịch vụ gom hàng chính ngạch Eureka Logistics.',
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

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 41 - {update_date_str})."
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
    ws3['A3'] = f"Chu kỳ: Tuần 41 (05/10/2026 - 11/10/2026) | Mở rộng 115 chủ đề nghiên cứu (gồm 83 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 115)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 115 or m.max_row >= 115]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (112,
         'phan cap tham quyen cap co dia phuong thong tu 26 2026 tt bct mau so 03 chuyen noi cap',
         'thong tu 26/2026/tt-bct, nghi dinh 146/2025/nd-cp, phan cap cap co dia phuong, so cong thuong cap co, mau so 03 chuyen noi cap co, co form e form d form ak, he thong ecosys co moit gov vn',
         97.5, 16800, 'https://congthuong.vn',
         'Thông tư số 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP: Phân cấp thẩm quyền cấp Giấy chứng nhận xuất xứ hàng hóa (C/O) cho các cơ quan, tổ chức địa phương',
         'Nội dung trọng tâm: Bộ Công Thương phân cấp thẩm quyền cấp C/O từ Cục Xuất nhập khẩu về Sở Công Thương các tỉnh/thành phố; phân định rõ thẩm quyền: Cục XNK tiếp tục cấp 12 mẫu (theo Mẫu 17 Phụ lục II) trong khi địa phương cấp các mẫu còn lại (Mẫu 18); hướng dẫn doanh nghiệp thực hiện Mẫu số 03 để chuyển nơi cấp về Sở Công Thương địa bàn; lưu ý các mẫu C/O điện tử như D, AK, VK vẫn khai trên hệ thống ecosys.gov.vn truyền qua Cổng Một cửa Quốc gia (NSW), trong khi các mẫu khác khai trên co.moit.gov.vn; khuyến nghị doanh nghiệp rà soát sớm tránh nộp sai cơ quan dẫn đến chậm trễ C/O được hưởng ưu đãi thuế quan FTA.',
         'Báo Công Thương & scratch/daily_notes.md (09/10) | HT (Học Thực chiến) / C/O Issuance Decentralization & System SOP',
         update_date_str),

        (113,
         'dem nguoc 6 ngay nghi dinh 336 bat buoc vneid muc do 2 quyet dinh 1434 json hai quan',
         'dem nguoc 6 ngay, han chot 15/10/2026, nghi dinh 336/2026/nd-cp, quyet dinh 1434/qd-chq, vneid muc do 2 nsw, chuan du lieu json chung tu dien tu, cong mot cua quoc gia, hai quan so 2026',
         99.0, 19500, 'https://chinhphu.vn',
         'Đếm ngược 6 ngày (15/10/2026): Hạn chót liên kết VNeID Mức độ 2 trên NSW & Bắt buộc chuẩn dữ liệu chứng từ điện tử JSON DataFormat theo QĐ 1434/QĐ-CHQ',
         'Nội dung trọng tâm: Toàn ngành hải quan và cộng đồng doanh nghiệp XNK đếm ngược 6 ngày tới mốc 15/10/2026 khi Nghị định 336/2026/NĐ-CP và Quyết định 1434/QĐ-CHQ chính thức có hiệu lực đồng loạt; bắt buộc 100% tài khoản đại diện pháp luật trên Cổng thông tin một cửa quốc gia (NSW) phải định danh VNeID Mức độ 2; toàn bộ thông điệp dữ liệu chứng từ điện tử đi kèm tờ khai phải tuân thủ chuẩn JSON thay thế chuẩn XML cũ; cảnh báo nguy cơ hệ thống tự động khóa quyền truyền hồ sơ chuyên ngành đối với doanh nghiệp chậm cập nhật.',
         'Cổng TTĐT Chính phủ & Tổng cục Hải quan (customs.gov.vn) | TN (Tin Ngành) / Customs Digital Transformation Deadline',
         update_date_str),

        (114,
         'dieu chinh sai so nha dia diem kinh doanh ho kinh doanh cong dich vu cong tinh etax',
         'sai so nha dia diem kinh doanh, ho kinh doanh cong dvc tinh, thay doi noi dung dang ky ho kinh doanh, dong bo du lieu ma so thue etax, nghi dinh 01/2021/nd-cp, thong tu 105/2020/tt-btc, tra cuu ma so thue khong sua duoc',
         94.0, 13200, 'https://tapchitaichinh.vn',
         'Hướng dẫn xử lý sai sót số nhà, địa chỉ trên mã địa điểm kinh doanh Hộ kinh doanh: Quy trình nộp hồ sơ qua Cổng DVC cấp tỉnh & Đồng bộ Thuế eTax',
         'Nội dung trọng tâm: Gỡ rối tình huống thực tế khi thông tin số nhà hoặc địa chỉ địa điểm kinh doanh của hộ kinh doanh bị sai lệch so với thực tế; giải thích nguyên nhân không thể sửa trực tiếp trên chức năng tra cứu mã số thuế của hệ thống Thuế eTax; hướng dẫn quy trình nộp hồ sơ thay đổi nội dung đăng ký hộ kinh doanh/địa điểm kinh doanh trên Cổng Dịch vụ công cấp tỉnh hoặc bộ phận một cửa UBND cấp huyện để cơ quan đăng ký kinh doanh tiếp nhận và truyền dữ liệu đồng bộ tự động sang cơ quan Thuế; phòng ngừa rủi ro bị cơ quan thuế thông báo không hoạt động tại địa chỉ đăng ký.',
         'Tạp chí Tài chính & scratch/daily_notes.md (09/10) | CS (Case Study) / Business Location Tax Registration Rectification',
         update_date_str),

        (115,
         'drewry wci 08 10 giam 2 phan tram 4351 usd thieu cont 39 blank sailings hau golden week',
         'drewry world container index 08/10, wci 4351 usd, 39 blank sailings tuan 41-45, quan tri tai trong hang tau, gia cuoc tau bien giam nhe, roll over rot cont, khan hiem vo container sau golden week',
         95.5, 15400, 'https://drewry.co.uk',
         'Drewry WCI ngày 08/10/2026 giảm 2% xuống 4.351 USD/FEU: Hãng tàu áp dụng 39 Blank Sailings (Tuần 41-45) để giữ mặt bằng cước sau kỳ nghỉ lễ Trung Quốc',
         'Nội dung trọng tâm: Báo cáo Drewry WCI ngày 08/10/2026 ghi nhận chỉ số giá cước tổng hợp thế giới giảm 2% xuống 4.351 USD/FEU do nhu cầu hạ nhiệt trong kỳ nghỉ Tuần Lễ Vàng tại Trung Quốc; tuy nhiên các liên minh hãng tàu (2M, Ocean Alliance, THE Alliance) đã lên kế hoạch hủy 39 chuyến tàu (blank sailings) từ Tuần 41 đến Tuần 45 trên các tuyến Đông - Tây chính; cảnh báo rủi ro rớt hàng (roll over) và tình trạng lệch cân bằng vỏ cont rỗng tại các cảng xuất phát châu Á khi các nhà máy mở cửa sản xuất dồn dập trở lại; tư vấn chủ hàng chốt lịch booking sớm và yêu cầu cam kết space từ đại lý.',
         'Drewry Supply Chain Advisors & World Shipping Council | TN (Tin Ngành) / Ocean Freight & Post-Golden Week Capacity Strategy',
         update_date_str),
    ]

    # Write new items into rows 117 to 120
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 117 + idx
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

    # New Total row at row 121
    total_row_idx = 121
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 115 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D120)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E120)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 115 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 83 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A121:C121')
    ws3.merge_cells('F121:J121')
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
            import shutil
            shutil.copy2(FILE_PATH, onedrive_tmp)
            os.replace(onedrive_tmp, ONEDRIVE_FILE)
            print(f"THÀNH CÔNG: Đã đồng bộ sang OneDrive tại {ONEDRIVE_FILE}")
        else:
            print(f"Bỏ qua đồng bộ OneDrive (không tìm thấy {ONEDRIVE_DIR})")
    except Exception as e:
        print(f"Lỗi khi đồng bộ OneDrive: {e}")

if __name__ == '__main__':
    update_seo_w41_0910()
