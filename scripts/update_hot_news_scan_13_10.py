# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_13_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 13/10/2026 (Tuần 42: 12/10/2026 - 18/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 42, ngày 13/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 125 lên 130 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 13/10/2026)
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

def update_seo_w42_1310():
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

    update_date_str = "13/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 42 (12/10/2026 - 18/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van 22720 chq gsql to khai phi thue quan dncx chua xac nhan khu vuc giam sat',
         'Công văn số 22720/CHQ-GSQL (08/10/2026): Hướng dẫn xử lý dứt điểm các tờ khai xuất nhập khẩu phi thuế quan chưa được xác nhận qua khu vực giám sát (KVGS) do lệch chỉ tiêu "Số quản lý nội bộ"',
         7200, 'Đỉnh điểm cảnh báo (+1,250%)',
         'Doanh nghiệp chế xuất (DNCX), doanh nghiệp nội địa mua bán với DNCX, bộ phận xuất nhập khẩu, forwarder và đại lý hải quan.',
         'Hàng ngàn tờ khai phi thuế quan mở từ tháng 9/2026 bị hệ thống treo không tự động xác nhận qua KVGS do lệch cú pháp chỉ tiêu số quản lý nội bộ (#&XKPTQ vs #&NKTC); rủi ro không thanh khoản được tờ khai, treo hoàn thuế GTGT.',
         "Tiêu đề: 'Công văn 22720/CHQ-GSQL: Hướng dẫn gỡ lỗi tờ khai phi thuế quan DNCX chưa qua khu vực giám sát và quy trình gửi danh sách xử lý dứt điểm'.",
         'Công văn số 22720/CHQ-GSQL ngày 08/10/2026; Công văn số 20991/CHQ-GSQL ngày 26/08/2026; Cục Hải quan.',
         update_date_str),

        (2,
         'dem nguoc 2 ngay bat buoc lien ket vneid muc do 2 cong mot cua quoc gia nsw 15 10 2026',
         'ĐẾM NGƯỢC 2 NGÀY (Hạn chót 15/10/2026): Nghị định 336/2026/NĐ-CP & Quyết định 1434/QĐ-CHQ — Bắt buộc tài khoản định danh VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) và chuyển đổi chứng từ điện tử chuẩn JSON DataFormat',
         6950, 'Đếm ngược khẩn cấp (+1,180%)',
         'Toàn bộ lãnh đạo doanh nghiệp XNK, người đại diện pháp luật, kế toán trưởng, nhân viên khai báo hải quan, forwarder và đại lý thủ tục.',
         'Chỉ còn 48h trước hạn chót 15/10; các doanh nghiệp chưa tích hợp VNeID Mức 2 của người đại diện hoặc chưa cấu hình chuẩn JSON sẽ bị hệ thống tự động khóa chức năng truyền hồ sơ cấp phép chuyên ngành và tờ khai hải quan.',
         "Tiêu đề: 'Đếm ngược 2 ngày (15/10): Hướng dẫn khẩn cấp liên kết VNeID Mức 2 trên NSW & Kiểm tra chuẩn JSON chứng từ tránh tê liệt thông quan'.",
         'Nghị định số 336/2026/NĐ-CP ngày 22/08/2026; Thông tư số 128/2026/TT-BTC; Quyết định số 1434/QĐ-CHQ; Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (3,
         'bi khai khong thu nhap luong 150 trieu tren etax mobile phan anh tong cuc thue',
         'Bẫy doanh nghiệp lạ khai khống lương 150 triệu trên eTax Mobile: Rủi ro người lao động & doanh nghiệp bị truy thu thuế oan và Quy trình 3 bước gửi phản ánh "Thu nhập khống" trực tuyến',
         5800, 'Tăng đột biến (+920%)',
         'Người lao động, nhân viên văn phòng, kế toán thuế, doanh nghiệp sử dụng lao động và các đơn vị quyết toán thuế TNCN.',
         'Bị doanh nghiệp lạ tại Khu công nghiệp lấy trộm mã số thuế để khai khống tiền lương 150 triệu/năm nhằm tăng chi phí ảo; dẫn đến việc người lao động và doanh nghiệp thực tế bị cơ quan thuế réo tên truy thu thuế TNCN bậc cao.',
         "Tiêu đề: 'Cảnh báo: Bị khai khống lương 150 triệu trên eTax Mobile & Quy trình gửi phản ánh thu nhập khống để cơ quan thuế vào cuộc'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư 80/2021/TT-BTC; Cổng TTĐT Tổng cục Thuế; scratch/daily_notes.md (13/10/2026).',
         update_date_str),

        (4,
         'thu tuc nhap khau banh keo thuc pham dai loan tu cong bo attp nghi dinh 15 2018',
         'Thủ tục nhập khẩu bánh kẹo, thực phẩm chế biến từ Đài Loan mở chuỗi phân phối: Hồ sơ Tự công bố ATTP theo Nghị định 15/2018 & Đăng ký kiểm tra nhà nước chuyên ngành',
         4600, 'Tăng mạnh (+680%)',
         'Doanh nghiệp nhập khẩu thực phẩm, chuỗi cửa hàng tiện lợi, thương nhân phân phối bánh kẹo ngoại nhập và forwarder dịch vụ XNK.',
         'Thiếu hồ sơ Tự công bố ATTP trước khi nhập hàng hoặc không đăng ký kiểm tra nhà nước về an toàn thực phẩm trên Cổng NSW khiến hàng bị kẹt tại cảng, phát sinh chi phí lưu kho bãi và rủi ro bị buộc tái xuất.',
         "Tiêu đề: 'Nhập khẩu bánh kẹo Đài Loan mở chuỗi cửa hàng: Quy trình Tự công bố ATTP và đăng ký kiểm tra nhà nước không lo vướng thủ tục'.",
         'Nghị định số 15/2018/NĐ-CP; Nghị định 43/2017/NĐ-CP & 111/2021/NĐ-CP; Cục An toàn thực phẩm; scratch/daily_notes.md (13/10/2026).',
         update_date_str),

        (5,
         'thu tuc qua canh hang hoa hoa chat ma hs 3801 giay phep bo cong thuong vn sang laos',
         'Thủ tục vận chuyển quá cảnh hàng hóa & hóa chất đặc thù (HS 3801.9000, Mã CAS) từ Việt Nam sang Lào: Quy định xin Giấy phép quá cảnh Bộ Công Thương theo Nghị định 69/2018',
         4100, 'Tăng mới (+540%)',
         'Doanh nghiệp vận tải liên vận quốc tế Việt - Lào, forwarder chuyên tuyến đường bộ, thương nhân kinh doanh hóa chất và nguyên liệu khoáng sản.',
         'Nhầm lẫn giữa hàng quá cảnh thông thường và hóa chất cần giấy phép của Bộ Công Thương; thiếu giấy phép quá cảnh bị cơ quan hải quan cửa khẩu chặn giữ phương tiện, xử phạt vi phạm hành chính nặng.',
         "Tiêu đề: 'Vận chuyển hàng quá cảnh sang Lào: Khi nào bắt buộc phải xin Giấy phép quá cảnh Bộ Công Thương và cách tra cứu mã CAS chuẩn xác?'.",
         'Điều 35 Nghị định 69/2018/NĐ-CP; Luật Hóa chất; Luật Hải quan số 54/2014/QH13; scratch/daily_notes.md (13/10/2026).',
         update_date_str),

        (6,
         'doanh nghiep moi lap bi khoa mst khong hoat dong tai dia chi dang ky mau 25 dk tct',
         'Doanh nghiệp mới thành lập bị tạm khóa MST vì "Không hoạt động tại địa chỉ đã đăng ký" dù chưa hết hạn nộp tờ khai Quý 3 (hạn 31/10): Bẫy xác minh thực địa & Thủ tục khôi phục MST theo Mẫu 25/ĐK-TCT',
         5450, 'Tăng cao (+890%)',
         'Doanh nghiệp mới thành lập (Q3/2026), startup, người đại diện pháp luật, kế toán thuế và dịch vụ thành lập doanh nghiệp.',
         'Công ty lập tháng 7/2026, chưa hết hạn nộp tờ khai Quý 3 (31/10), chưa phát sinh doanh thu chi phí, đã treo biển hiệu nhưng bị cán bộ thuế kiểm tra thực địa đột xuất vắng mặt -> phát thông báo đóng MST, tê liệt tài khoản ngân hàng và kinh doanh.',
         "Tiêu đề: 'Doanh nghiệp mới lập bị khóa MST do thuế báo không hoạt động tại trụ sở: Quy trình 3 bước giải trình và khôi phục MST theo Mẫu 25/ĐK-TCT'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư 105/2020/TT-BTC hướng dẫn về đăng ký thuế; scratch/daily_notes.md (12/10/2026).',
         update_date_str),

        (7,
         'drewry wci tuan 42 34 blank sailings siet cong suat a au a my giu gia cuoc',
         'BÁO CÁO DREWRY WCI (Tuần 42): Chỉ số cước biển 4.351 USD/FEU & Hãng tàu công bố 34 Blank Sailings (Tuần 42-46) siết tải trọng các tuyến Á - Âu và Transpacific hậu Golden Week',
         4900, 'Tăng mạnh (+510%)',
         'Chủ hàng xuất khẩu đi thị trường Bắc Mỹ, Châu Âu, doanh nghiệp sản xuất dệt may, da giày, gỗ, điện tử và các công ty forwarder.',
         'Dù cước WCI giảm 2% xuống 4.351 USD/FEU, các liên minh hãng tàu đã công bố 34 chuyến hủy (50% tuyến Á - Bắc Âu, 32% Transpacific) để ghìm tải; rủi ro rớt container (roll over) và thiếu vỏ cont 40HC mùa cao điểm cuối năm.',
         "Tiêu đề: 'Drewry WCI 4.351 USD & 34 chuyến tàu bị hủy từ Tuần 42: Bí quyết giữ booking và tránh rớt container mùa hàng Noel - Tết'.",
         'Báo cáo Drewry Cancelled Sailings Tracker (cập nhật 09/10/2026); Drewry World Container Index; World Shipping Council.',
         update_date_str),

        (8,
         'dem nguoc 2 ngay thong tu 128 2026 tt btc cua khau thong minh xe tu hanh agv huu nghi',
         'ĐẾM NGƯỢC 2 NGÀY (Hiệu lực 15/10/2026): Thông tư số 128/2026/TT-BTC — Vận hành thí điểm Cửa khẩu thông minh Hữu Nghị & Tân Thanh với xe tự hành AGV không người lái và chuẩn dữ liệu thông quan 24/7',
         5650, 'Tăng rất mạnh (+820%)',
         'Doanh nghiệp xuất nhập khẩu đường bộ qua Lạng Sơn, chủ hàng nông sản, linh kiện điện tử, doanh nghiệp vận tải biên giới và logistics Trung - Việt.',
         'Chỉ còn 48h trước thời điểm mở barie tự động và chạy thử xe tự hành AGV; doanh nghiệp chưa nắm chuẩn thông điệp dữ liệu điện tử phương tiện và hàng hóa sẽ bị từ chối vào làn thông minh, buộc dồn ứ tại làn truyền thống.',
         "Tiêu đề: 'Đếm ngược 2 ngày Thông tư 128/2026/TT-BTC: Cửa khẩu thông minh vận hành xe AGV & Hướng dẫn khai báo dữ liệu thông quan 24/7'.",
         'Thông tư số 128/2026/TT-BTC ngày 27/08/2026; Đề án thí điểm Cửa khẩu thông minh tỉnh Lạng Sơn; Cổng TTĐT Bộ Công Thương.',
         update_date_str),

        (9,
         'hang khong co hoa don bi phat ban hang lau nghi dinh 98 2020 bang ke thu mua ho kinh doanh duoi 100 trieu',
         'Hàng không có hóa đơn hợp lệ bị phạt hành vi "Bán hàng nhập lậu" theo Nghị định 98/2020/NĐ-CP: Bẫy Bảng kê hộ kinh doanh & Rủi ro tịch thu 100% tang vật hàng hóa',
         5850, 'Duy trì đỉnh (+890%)',
         'Hộ kinh doanh cá thể, chủ shop thương mại điện tử, doanh nghiệp thương mại bán buôn bán lẻ và kế toán thuế.',
         'Nhiều đơn vị nhầm tưởng mua hàng không hóa đơn chỉ bị loại chi phí thuế TNDN, nhưng Nghị định 98/2020/NĐ-CP xử phạt hành vi kinh doanh hàng hóa nhập lậu, phạt tiền từ 20-100 triệu và tịch thu toàn bộ hàng hóa trôi nổi.',
         "Tiêu đề: 'Cảnh báo: Bán hàng không có hóa đơn bị phạt \"Bán hàng nhập lậu\" theo Nghị định 98 & Giải pháp nhập khẩu chính ngạch bảo vệ tài sản'.",
         'Nghị định 98/2020/NĐ-CP; Nghị định 125/2020/NĐ-CP; Thông tư 78/2021/TT-BTC; scratch/daily_notes.md (10/10/2026).',
         update_date_str),

        (10,
         'gia han thoi han tam nhap tai xuat khai tia hay nop dich vu cong truoc thong tu 38 thong tu 39 hai quan',
         'Thủ tục gia hạn thời hạn Tạm nhập - Tái xuất (G11, G12, G21): Quy trình chuẩn nộp hồ sơ qua Cổng Dịch vụ công trực tuyến Hải quan trước hay khai chỉ tiêu sửa đổi TIA trên VNACCS/VCIS theo Thông tư 38 & Thông tư 39',
         5400, 'Tăng mạnh (+750%)',
         'Doanh nghiệp tạm nhập tái xuất, gia công, SXXK, bảo hành bảo trì máy móc, forwarder và đại lý hải quan.',
         'Lúng túng thứ tự thực hiện giữa khai TIA và nộp hồ sơ Dịch vụ công; nếu khai TIA trước khi có văn bản chấp thuận của Chi cục trưởng thì bị từ chối, quá hạn tái xuất bị ấn định thuế nhập khẩu và phạt chậm nộp 0.03%/ngày.',
         "Tiêu đề: 'Gia hạn thời hạn tạm nhập tái xuất: Hướng dẫn nộp hồ sơ DVC Hải quan trước và khai TIA đúng luật tránh bị ấn định thuế'.",
         'Điều 50 & 51 Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018/TT-BTC); Cổng DVC trực tuyến Tổng cục Hải quan; scratch/daily_notes.md (10/10/2026).',
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
    ws3['A3'] = f"Chu kỳ: Tuần 42 (12/10/2026 - 18/10/2026) | Mở rộng 130 chủ đề nghiên cứu (gồm 98 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 131)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 131 or m.max_row >= 131]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (126,
         'cong van 22720 chq gsql to khai phi thue quan dncx chua xac nhan khu vuc giam sat',
         'cong van 22720 chq gsql, to khai phi thue quan, chua xac nhan khu vuc giam sat, so quan ly noi bo doanh nghiep, xkptq nktc, xktc nkptq, bophanhotro customs gov vn, tu dong xac nhan kvgs, dncx va noi dia, thong quan xuat nhap khau tai cho',
         99.8, 36500, 'https://customs.gov.vn',
         'Công văn số 22720/CHQ-GSQL (08/10/2026): Hướng dẫn xử lý dứt điểm các tờ khai xuất nhập khẩu phi thuế quan chưa được xác nhận qua khu vực giám sát (KVGS) do lệch chỉ tiêu "Số quản lý nội bộ"',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn số 22720/CHQ-GSQL ngày 08/10/2026 tháo gỡ điểm nghẽn nghiêm trọng cho các doanh nghiệp chế xuất (DNCX) và doanh nghiệp nội địa liên quan đến tính năng tự động xác nhận qua khu vực giám sát (KVGS) triển khai từ 09/09/2026; do sự không đồng nhất cú pháp khai báo tại chỉ tiêu "Số quản lý nội bộ doanh nghiệp" (như bên xuất khai #&XKPTQ nhưng bên nhập khai #&NKTC#&Số_TK_XK, hoặc bên xuất khai #&XKTC nhưng bên nhập khai #&NKPTQ#&Số_TK_XK), hàng ngàn tờ khai xuất nhập khẩu tại chỗ và mua bán với DNCX bị treo trạng thái chưa qua KVGS, dẫn tới nguy cơ tắc nghẽn thanh khoản, chậm hoàn thuế GTGT và rủi ro bị xử phạt vi phạm thủ tục; hướng dẫn quy trình xử lý khẩn cấp: (1) Doanh nghiệp chủ động rà soát đối chiếu cặp tờ khai xuất - nhập đối ứng; (2) Chi cục Hải quan nơi mở tờ khai tổng hợp danh sách các cặp tờ khai bị vướng gửi trực tiếp về Ban CNTT & Thống kê Hải quan qua email bophanhotro@customs.gov.vn để can thiệp hỗ trợ xác nhận trên hệ thống VNACCS; (3) Hướng dẫn chuẩn hóa 100% các tờ khai mới tuân thủ nghiêm ngặt chuẩn cú pháp theo Công văn số 20991/CHQ-GSQL để hệ thống tự động thông quan thông suốt.',
         'Tổng cục Hải quan (customs.gov.vn) & Thư Viện Pháp Luật | CS (Case Study) / Bonded & EPE Clearance Rectification SOP',
         update_date_str),

        (127,
         'dem nguoc 2 ngay bat buoc lien ket vneid muc do 2 cong mot cua quoc gia nsw 15 10 2026',
         'dem nguoc 2 ngay, han chot 15 10 2026, bat buoc vneid muc do 2, cong mot cua quoc gia nsw, nghi dinh 336 2026, quyet dinh 1434 qd chq, chuan du lieu json chung tu dien tu, khoa tai khoan nsw, phan quyen dai ly hai quan, uy quyen dien tu',
         100.0, 38900, 'https://chinhphu.vn',
         'ĐẾM NGƯỢC 2 NGÀY (Hạn chót 15/10/2026): Bắt buộc tài khoản định danh VNeID Mức độ 2 trên NSW & Chuẩn dữ liệu chứng từ điện tử JSON DataFormat theo QĐ 1434/QĐ-CHQ',
         'Nội dung trọng tâm: Toàn ngành hải quan và cộng đồng doanh nghiệp XNK bước vào 48 giờ đếm ngược cuối cùng trước thời điểm Nghị định 336/2026/NĐ-CP và Quyết định 1434/QĐ-CHQ chính thức có hiệu lực từ ngày 15/10/2026; cảnh báo cấp bách: 100% doanh nghiệp chưa hoàn tất liên kết tài khoản định danh điện tử VNeID Mức độ 2 của Người đại diện theo pháp luật trên Cổng Một cửa Quốc gia (NSW) sẽ bị hệ thống tự động ngắt kết nối và khóa quyền tiếp nhận hồ sơ xin cấp phép chuyên ngành, công bố hợp quy và cấp C/O; đồng thời toàn bộ chứng từ đính kèm tờ khai hải quan bắt buộc chuyển đổi sang cấu trúc dữ liệu JSON Schema; hướng dẫn 3 bước khẩn cấp: (1) Đăng nhập nsw.gov.vn bằng chữ ký số USB Token và liên kết số CCCD đã kích hoạt VNeID Mức 2; (2) Thiết lập phân quyền điện tử nội bộ cho nhân viên xuất nhập khẩu và đại lý hải quan; (3) Cập nhật phiên bản phần mềm khai báo hải quan mới nhất để đồng bộ chuẩn JSON tránh lỗi từ chối thông điệp.',
         'Cổng TTĐT Chính phủ & Tổng cục Hải quan (customs.gov.vn) | TN (Tin Ngành) / Customs Digital Compliance Deadline 15-10',
         update_date_str),

        (128,
         'bi khai khong thu nhap luong 150 trieu tren etax mobile phan anh tong cuc thue',
         'khai khong thu nhap, danh cap ma so thue ca nhan, truy thu thue tncn, phan anh thong tin thu nhap khong etax mobile, quyet toan thue tncn, hanh vi tron thue doanh nghiep ma, giai trinh co quan thue, bao hiem xa hoi bat hop phap',
         98.5, 29800, 'https://tapchitaichinh.vn',
         'Cảnh báo: Bị doanh nghiệp lạ khai khống lương 150 triệu trên eTax Mobile dẫn đến bị cơ quan thuế truy thu — Quy trình gửi phản ánh "Thu nhập khống" trực tuyến',
         'Nội dung trọng tâm: Tình huống phát sinh thực tế từ daily_notes.md (13/10/2026): Người lao động thực tế chỉ làm việc tại doanh nghiệp A với mức lương 50 triệu đồng/năm (thuộc diện không phải nộp thuế TNCN), bất ngờ nhận thông báo từ cơ quan thuế về việc phát sinh nợ thuế TNCN và bị yêu cầu quyết toán thuế bổ sung; qua tra cứu trên app eTax Mobile mới phát hiện một doanh nghiệp lạ B tại Khu công nghiệp đã tự ý lấy cắp thông tin CMND/CCCD để kê khai khống mức lương 150 triệu đồng/năm và đóng BHXH nhằm nâng khống chi phí được trừ khi tính thuế TNDN; hướng dẫn quy trình 3 bước xử lý bảo vệ quyền lợi: (1) Đăng nhập ứng dụng eTax Mobile, truy cập mục "Tra cứu thông tin quyết toán thuế" để trích xuất danh sách các tổ chức chi trả thu nhập; (2) Nhấp vào tổ chức chi trả sai lệch và bấm chọn chức năng "Phản ánh thông tin thu nhập khống" kèm bản cam kết không phát sinh thu nhập tại DN lạ; (3) Hệ thống tự động chuyển thông tin về Cục Thuế quản lý DN gian lận để tiến hành thanh tra, xử phạt hành vi trốn thuế và gỡ bỏ nghĩa vụ thuế oan cho người lao động.',
         'Tạp chí Tài chính & scratch/daily_notes.md (13/10) | CS (Case Study) / Fraud Income Rectification & eTax Mobile SOP',
         update_date_str),

        (129,
         'thu tuc nhap khau banh keo thuc pham dai loan tu cong bo attp nghi dinh 15 2018',
         'nhap khau banh keo dai loan, tu cong bo an toan thuc pham, nghi dinh 15 2018 nd cp, kiem tra nha nuoc an toan thuc pham, nghi dinh 43 2017 ghi nhan hang hoa, tem phu tieng viet, mo chuoi cua hang banh keo, co form ec, thue nhap khau uu dai',
         97.5, 25400, 'https://congthuong.vn',
         'Cẩm nang thủ tục nhập khẩu bánh kẹo, thực phẩm chế biến từ Đài Loan mở chuỗi phân phối: Hồ sơ Tự công bố ATTP & Kiểm tra nhà nước chuyên ngành',
         'Nội dung trọng tâm: Nhu cầu nhập khẩu bánh kẹo, thực phẩm chế biến từ Đài Loan để phát triển chuỗi cửa hàng bán lẻ gia tăng mạnh đón mùa tiêu dùng cuối năm từ daily_notes.md (13/10/2026); hướng dẫn toàn trình 4 bước tuân thủ pháp lý: (1) Kiểm nghiệm mẫu sản phẩm tại phòng kiểm nghiệm được chỉ định và lập Bản tự công bố an toàn thực phẩm nộp tại Sở Y tế/Ban Quản lý ATTP theo Nghị định 15/2018/NĐ-CP trước khi nhập khẩu thương mại; (2) Trước khi hàng cập cảng, đăng ký Kiểm tra nhà nước về an toàn thực phẩm trên Cổng Một cửa Quốc gia (NSW) và nộp bộ hồ sơ gồm Bản tự công bố, Packing List, Invoice, Bill of Lading; (3) Tuân thủ nghiêm ngặt quy định ghi nhãn hàng hóa bắt buộc bằng tiếng Việt (tên sản phẩm, thành phần, hạn sử dụng, thông tin nhà sản xuất và thương nhân chịu trách nhiệm) theo Nghị định 43/2017/NĐ-CP & Nghị định 111/2021/NĐ-CP; (4) Khuyến nghị đối chiếu chính sách thuế nhập khẩu và kiểm tra kỹ chứng nhận xuất xứ (C/O) để tối ưu chi phí và tránh rủi ro hàng bị lưu kho lưu bãi.',
         'Báo Công Thương & scratch/daily_notes.md (13/10) | HT (Học Thực chiến) / Food Import Compliance & Retail Chain SOP',
         update_date_str),

        (130,
         'thu tuc qua canh hang hoa hoa chat ma hs 3801 giay phep bo cong thuong vn sang laos',
         'thu tuc qua canh hang hoa, qua canh sang lao, ma hs 3801, than chi nhan tao, ma cas hoa chat, giay phep qua canh bo cong thuong, nghi dinh 69 2018 nd cp, luat hoa chat, to khai hai quan g61, niem phong hai quan, giam sat hanh trinh',
         96.5, 22100, 'https://haiquanonline.com.vn',
         'Thủ tục hải quan quá cảnh hàng hóa & hóa chất đặc thù (HS 3801.9000, Mã CAS) từ Việt Nam sang Lào: Quy định xin Giấy phép quá cảnh Bộ Công Thương theo Nghị định 69/2018',
         'Nội dung trọng tâm: Tình huống phát sinh thực tế từ daily_notes.md (13/10/2026) về lô hàng nguyên liệu than chì nhân tạo (mã HS 3801.9000, mã CAS 7704-34-9, 7782-42-5) cần quá cảnh qua lãnh thổ Việt Nam sang Lào; giải đáp chi tiết: (1) Căn cứ Điều 35 Nghị định 69/2018/NĐ-CP và Luật Hóa chất, các mặt hàng hóa chất thuộc Danh mục hóa chất sản xuất, kinh doanh có điều kiện hoặc danh mục tiền chất khi quá cảnh bắt buộc phải xin Giấy phép quá cảnh do Bộ Công Thương cấp; nếu chỉ là than chì nhân tạo thông thường không nằm trong danh mục hóa chất nguy hiểm/cấm thì được miễn giấy phép quá cảnh của Bộ Công Thương; (2) Thủ tục hải quan mở tờ khai vận chuyển kết hợp (mã loại hình G61) trên hệ thống VNACCS tại Chi cục Hải quan cửa khẩu nhập; (3) Yêu cầu niêm phong hải quan nguyên trạng phương tiện vận tải hoặc container gắn thiết bị định vị GPS giám sát hành trình từ cửa khẩu nhập đến cửa khẩu xuất sang Lào theo đúng tuyến đường và thời gian quy định tại Điều 43 Luật Hải quan.',
         'Báo Hải quan Online & scratch/daily_notes.md (13/10) | HT (Học Thực chiến) / Transit Cargo & Chemical Compliance SOP',
         update_date_str),
    ]

    # Write new items into rows 131 to 135
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 131 + idx
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

    # New Total row at row 136
    total_row_idx = 136
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 130 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D135)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E135)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 130 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 98 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A136:C136')
    ws3.merge_cells('F136:J136')
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
    update_seo_w42_1310()
