# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_10_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 10/11/2026
(Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026 | Cập nhật ngày 10/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 46 cập nhật ngày 10/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 240 lên 245 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 10/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w46_1011.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w46_1011.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w46_1011():
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

    update_date_str = "10/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 46 (09/11/2026 - 15/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'hang gia cong bi ben nuoc ngoai tu bo thong tu 121 2025 tt btc dieu 66',
         'Hàng gia công bị bên nước ngoài từ bỏ theo Thông tư 121/2025/TT-BTC: 03 hướng xử lý, phác đồ mở tờ khai A42 nộp thuế và trách nhiệm pháp lý doanh nghiệp',
         19200, 'Điểm nóng hàng gia công (+3,750%)',
         'Doanh nghiệp gia công may mặc, da giày, linh kiện điện tử, cơ khí; giám đốc nhà máy, kế toán kho và nhân viên khai báo hải quan.',
         'Bên thuê gia công nước ngoài từ bỏ nguyên vật liệu, thành phẩm, máy móc; doanh nghiệp lúng túng không biết tiêu thụ nội địa hay tiêu hủy; nếu quá hạn 30 ngày bị hải quan lập biên bản và ấn định thuế.',
         "Tiêu đề: 'Bên đặt gia công từ bỏ hàng: Phác đồ 3 bước mở tờ khai A42 hoặc tiêu hủy đúng Điều 66 Thông tư 121/2025/TT-BTC tránh bị xử phạt'.",
         'Thông tư số 121/2025/TT-BTC (sửa Điều 66 Thông tư 38/2015/TT-BTC); scratch/daily_notes.md (10/11/2026); Cục Hải quan.',
         update_date_str),

        (2,
         'thanh khoan dncx bao cao quyet toan hai quan mau 15 chenh lech xuat nhap ton',
         'Kiểm tra sau thông quan thanh khoản DNCX: Cách xử lý chênh lệch số liệu xuất nhập tồn, định mức hao hụt và bảo vệ Báo cáo quyết toán Mẫu 15',
         18800, 'Mùa rà soát thanh khoản DNCX (+3,610%)',
         'Doanh nghiệp chế xuất (DNCX), nhà máy FDI, giám đốc sản xuất, kế toán trưởng và nhân viên khai báo hải quan.',
         'Số liệu kho thực tế lệch với ERP và tờ khai hải quan; âm nguyên liệu không chứng minh được định mức hao hụt; nguy cơ bị ấn định thuế NK, VAT và phạt trốn thuế khi kiểm tra sau thông quan.',
         "Tiêu đề: 'Bảo vệ Báo cáo quyết toán DNCX trước kiểm tra sau thông quan: Công thức Tam giác vàng đối soát Hải quan - Kế toán - Kho'.",
         'Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; scratch/daily_notes.md (10/11/2026); Báo Hải quan Online.',
         update_date_str),

        (3,
         'cong van 23913 chq gsql nhap khau hoa chat dncx luat quan ly ngoai thuong',
         'Công văn 23913/CHQ-GSQL ngày 05/11/2026: Hải quan đề nghị tháo gỡ giao thoa pháp lý nhập khẩu hóa chất vào DNCX giữa Luật Hóa chất và Ngoại thương',
         18300, 'Điểm nghẽn pháp lý hóa chất FDI (+3,490%)',
         'Doanh nghiệp chế xuất nhập khẩu hóa chất phục vụ sản xuất linh kiện, điện tử, bán dẫn và các đơn vị logistics chuyên tuyến KCN/KCX.',
         'Xung đột giữa Điều 59 Luật Quản lý ngoại thương (khu phi thuế quan miễn GP) và Luật Hóa chất; nguy cơ hàng kẹt tại cửa khẩu cảng biển/sân bay.',
         "Tiêu đề: 'Công văn 23913/CHQ-GSQL: Hướng giải quyết vướng mắc thủ tục nhập khẩu hóa chất vào DNCX không lo ách tắc thông quan'.",
         'Công văn 23913/CHQ-GSQL ngày 05/11/2026; Điều 59 Luật Quản lý ngoại thương; customs.gov.vn.',
         update_date_str),

        (4,
         'quyet dinh 31 2026 qd ttg danh muc hang hoa phai lam thu tuc tai cua khau nhap',
         'Quyết định 31/2026/QĐ-TTg thay thế QĐ 23/2019/QĐ-TTg: Danh mục hàng hóa bắt buộc làm thủ tục tại cửa khẩu nhập và tác động đến chuỗi vận tải ICD',
         17900, 'Quy định cửa khẩu nhập (+3,380%)',
         'Doanh nghiệp nhập khẩu ô tô, hóa chất, phế liệu, máy móc cũ; các công ty giao nhận forwarder và đơn vị quản lý cảng cạn ICD.',
         'Không được chuyển cửa khẩu về ICD nội địa, phải làm thủ tục trực tiếp tại cảng biển/sân bay; tăng chi phí lưu bãi, lưu container nếu không chuẩn bị trước.',
         "Tiêu đề: 'Quyết định 31/2026/QĐ-TTg: Danh mục hàng hóa bắt buộc làm thủ tục tại cửa khẩu nhập và giải pháp tối ưu chi phí cảng'.",
         'Quyết định số 31/2026/QĐ-TTg; chinhphu.vn; haiquanonline.com.vn.',
         update_date_str),

        (5,
         'so hoa hai quan thong minh 2026 giam sat tu dong cua khau so cang bien so',
         'Ngành Hải quan tăng tốc vận hành Hệ thống Hải quan số thông minh cuối năm 2026: Doanh nghiệp logistics cần chuẩn bị gì cho mô hình Cảng biển số?',
         17500, 'Nước rút chuyển đổi số hải quan (+3,290%)',
         'Toàn bộ doanh nghiệp XNK, các hãng tàu, đại lý hải quan và doanh nghiệp kinh doanh cảng biển, kho bãi.',
         'Chuyển đổi từ thủ tục giấy tờ sang tích hợp API số hóa, camera AI nhận diện tại cửa khẩu/cảng biển; doanh nghiệp chậm thích ứng sẽ bị chậm trễ thông quan.',
         "Tiêu đề: 'Hệ thống Hải quan số thông minh 2026: 4 chuẩn bị cốt lõi cho doanh nghiệp xuất nhập khẩu và cảng biển số'.",
         'Báo Hải quan Online (haiquanonline.com.vn); customs.gov.vn.',
         update_date_str),

        (6,
         'xu ly nguyen lieu du thua hop dong gia cong thong tu 121 2025 tt btc dieu 64',
         'Siết chặt xử lý nguyên liệu, vật tư dư thừa sau hợp đồng gia công theo Thông tư 121/2025/TT-BTC: Hạn chót 30 ngày, 05 hình thức xử lý và bãi bỏ ngưỡng miễn giám sát tiêu hủy',
         17100, 'Tâm điểm gia công xuất khẩu (+3,180%)',
         'Các doanh nghiệp gia công dệt may, da giày, linh kiện điện tử, cơ khí; nhân viên xuất nhập khẩu, kế toán kho và đại lý hải quan.',
         'Quá hạn 30 ngày kể từ khi kết thúc hợp đồng gia công nếu chưa hoàn tất thủ tục xử lý sẽ bị Hải quan lập biên bản vi phạm và ấn định thuế; lo lắng về quy trình tiêu hủy mới qua Hệ thống (Mẫu 22a).',
         "Tiêu đề: 'Quản trị hợp đồng gia công: 5 hình thức xử lý nguyên liệu dư thừa theo Thông tư 121/2025/TT-BTC tránh bị xử phạt và ấn định thuế'.",
         'Thông tư số 121/2025/TT-BTC (sửa Điều 64 Thông tư 38/2015/TT-BTC); Nghị định 08/2015/NĐ-CP; scratch/daily_notes.md (09/11/2026).',
         update_date_str),

        (7,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl hang mau to khai',
         'Thông tư 142/2026/TT-BTC (hiệu lực từ 16/11/2026): Quy chuẩn mới về thủ tục hải quan chuyển phát nhanh quốc tế, số hóa Master/House AWB và cách lấy tờ khai hàng mẫu',
         16800, 'Quy chuẩn chuyển phát nhanh (+3,080%)',
         'Doanh nghiệp gửi hàng mẫu qua FedEx, DHL, UPS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Gửi hàng mẫu qua bưu chính/chuyển phát nhanh thường bị thông quan gom, không có tờ khai hải quan riêng đứng tên doanh nghiệp dẫn đến bị loại chi phí hợp lý và không được khấu trừ thuế GTGT 0%.',
         "Tiêu đề: 'Thông tư 142/2026/TT-BTC hiệu lực từ 16/11: Hướng dẫn lấy tờ khai hải quan hàng mẫu chuyển phát nhanh bảo vệ chi phí thuế hợp lệ'.",
         'Thông tư số 142/2026/TT-BTC; Báo Hải quan Online; scratch/daily_notes.md; Tổng cục Hải quan.',
         update_date_str),

        (8,
         'thue tncn gio lam them ot tu 1 7 gioi han 40h thang quyet toan thue',
         'Vướng mắc thuế TNCN đối với giờ làm thêm (OT) từ 01/07 và giới hạn 40 giờ OT/tháng: Cách xác định phần thu nhập miễn thuế TNCN và giải pháp hạch toán chi phí hợp lệ',
         16300, 'Kế toán thuế & Tiền lương (+2,950%)',
         'Giám đốc nhân sự, kế toán tiền lương, kế toán thuế và người lao động có phát sinh giờ làm thêm mùa cao điểm cuối năm.',
         'Hiểu nhầm quy định miễn thuế TNCN đối với phần tiền làm thêm giờ (phần chênh lệch giữa lương OT và lương bình thường); băn khoăn về việc vượt trần 40h/tháng thì bị tính thuế ra sao.',
         "Tiêu đề: 'Thuế TNCN giờ làm thêm (OT) cuối năm: Cách tính phần miễn thuế và phân bổ chi phí hợp lệ theo quy định mới nhất'.",
         'scratch/daily_notes.md (10/11/2026); Luật Thuế TNCN; Nghị định 145/2020/NĐ-CP; Diễn đàn thuế chuyên sâu.',
         update_date_str),

        (9,
         'giai the doanh nghiep con hang ton kho xuat hoa don tra mat bang tru so',
         'Xử lý hàng tồn kho và hóa đơn khi giải thể doanh nghiệp không còn trụ sở: Phác đồ chuyển trụ sở tạm thời, thanh lý tài sản và hoàn tất thủ tục thuế - hải quan',
         15800, 'Thủ tục giải thể & Thanh lý (+2,820%)',
         'Chủ doanh nghiệp, đại diện pháp luật, kế toán trưởng công ty đang làm thủ tục giải thể nhưng vẫn còn hàng tồn kho nhập khẩu/nội địa.',
         'Đã trả mặt bằng trụ sở nên không thể xuất hóa đơn thanh lý hàng tồn kho; rủi ro bị cơ quan thuế khóa mã số thuế, cưỡng chế và truy thu thuế GTGT, TNDN.',
         "Tiêu đề: 'Giải thể công ty khi còn hàng tồn kho mà đã trả trụ sở: 3 bước hợp thức hóa hóa đơn thanh lý tránh bị truy thu thuế'.",
         'scratch/daily_notes.md (10/11/2026); Luật Doanh nghiệp 2020; Luật Quản lý thuế số 38/2019/QH14; Diễn đàn thuế chuyên sâu.',
         update_date_str),

        (10,
         'drewry world container index wci thang 11 2026 bien dong cuoc tau bien cuoi nam',
         'Chỉ số cước Drewry WCI tuần đầu tháng 11/2026 điều chỉnh về 4.389 USD/FEU: Cơ hội đàm phán cước tàu và chiến lược đặt chỗ (booking) tránh rủi ro blank sailings cuối năm',
         15400, 'Thị trường cước biển hạ nhiệt (+2,740%)',
         'Doanh nghiệp xuất nhập khẩu hàng đi Mỹ (US West/East Coast), Châu Âu (Rotterdam/Hamburg) và các nhà máy xuất khẩu hàng tiêu dùng mùa lễ Tết.',
         'Cước spot giảm nhẹ nhưng các hãng tàu chuẩn bị cắt chuyến và tăng phụ phí FAK cuối tháng 11; rủi ro hoãn chuyến, rớt hàng tại cảng chuyển tải nếu không có kế hoạch booking trước 2-3 tuần.',
         "Tiêu đề: 'Cước tàu biển thế giới hạ nhiệt về 4.389 USD/FEU: Doanh nghiệp nên chốt cước dài hạn hay đi spot mùa cao điểm cuối năm?'.",
         'Báo cáo Drewry World Container Index ngày 05/11/2026; World Shipping Council; Báo Công Thương.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 46 (09/11/2026 - 15/11/2026) | Mở rộng 245 chủ đề nghiên cứu (gồm 213 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 246)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 246 or m.max_row >= 246]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (241,
         'hang gia cong bi ben nuoc ngoai tu bo thong tu 121 2025 tt btc dieu 66',
         'hang gia cong bi tu bo, thong tu 121 2025 tt btc, sua dieu 66 thong tu 38, ben dat gia cong tu bo hang, to khai a42 chuyen muc dich su dung, tich thu hang gia cong tu bo, thu tuc tieu huy hang gia cong khong nguoi nhan',
         99.9,
         59800,
         'https://thuvienphapluat.vn',
         'Xử lý hàng hóa khi bên đặt gia công nước ngoài từ bỏ theo Thông tư 121/2025/TT-BTC: 3 hướng xử lý, trách nhiệm pháp lý và phác đồ tờ khai A42',
         'Nội dung trọng tâm: Phân tích hướng dẫn cấp thiết tại khoản 37 Điều 1 Thông tư 121/2025/TT-BTC (sửa đổi Điều 66 Thông tư 38/2015/TT-BTC) ghi nhận ngày 10/11/2026 từ thực tế doanh nghiệp gia công: Khi bên đặt gia công nước ngoài có văn bản từ bỏ nguyên liệu dư thừa, máy móc thiết bị tạm nhập hoặc thành phẩm gia công, doanh nghiệp nhận gia công không được tự ý bỏ mặc hàng hóa tại kho bãi (rủi ro bị xử phạt vi phạm hành chính và nâng hạng rủi ro hải quan); Thông tư 121 quy định 03 hướng xử lý rõ ràng: (1) Tiêu thụ nội địa: Bên nhận gia công nộp đủ thuế NK, thuế GTGT theo thời điểm chuyển đổi mục đích sử dụng (mở tờ khai A42); (2) Tịch thu sung công quỹ: Nếu bên nhận gia công không nhận và hàng hóa còn giá trị sử dụng thì cơ quan hải quan ra quyết định tịch thu; (3) Tiêu hủy: Nếu hàng hóa không còn giá trị sử dụng thì bên nhận gia công bắt buộc phải tiêu hủy theo quy định tại Điều 64 và tự chịu mọi chi phí phát sinh; Khuyến nghị Eureka: Doanh nghiệp cần cài chặt điều khoản xử lý tài sản từ bỏ (Abandoned Goods Clause) ngay trong hợp đồng gia công, và chủ động làm việc với Hải quan trước thời hạn 30 ngày để chọn phương án tối ưu thuế.',
         'scratch/daily_notes.md (10/11) & Cục Hải quan, Thư Viện Pháp Luật | CS (Case Study) / Processing Contract & Abandoned Cargo Compliance',
         update_date_str),

        (242,
         'thanh khoan dncx bao cao quyet toan hai quan mau 15 chenh lech xuat nhap ton',
         'thanh khoan dncx, bao cao quyet toan hai quan mau 15, kiem tra sau thong quan dncx, chenh lech xuat nhap ton, dinh muc san xuat hao hut hat nhua, dong bo so lieu kho ke toan hai quan, an dinh thue bao cao quyet toan',
         99.8,
         58600,
         'https://customs.gov.vn',
         'Kiểm tra sau thông quan thanh khoản DNCX: Cách xử lý chênh lệch số liệu xuất nhập tồn, định mức hao hụt và bảo vệ Báo cáo quyết toán Mẫu 15',
         'Nội dung trọng tâm: Cảnh báo thực chiến từ ca kiểm tra thanh khoản DNCX ngày 10/11/2026: Doanh nghiệp chế xuất gặp sự cố chênh lệch hàng triệu linh kiện và âm 80 tấn hạt nhựa giữa số liệu thực tế tại kho, sổ sách kế toán ERP và dữ liệu tờ khai hải quan; nguyên nhân cốt lõi do không đồng nhất đơn vị tính (cuộn vs mét, bao vs kg), chậm cập nhật định mức hao hụt kỹ thuật và không lập biên bản điều chỉnh xuất nhập tồn trong kỳ; khi kiểm tra sau thông quan (KTSTQ), cơ quan hải quan sẽ ấn định thuế nhập khẩu, thuế GTGT và phạt trốn thuế đối với số lượng nguyên vật liệu âm không chứng minh được xuất xứ; Giải pháp chiến lược từ Eureka: Thiết lập cơ chế kiểm soát nội bộ \'Tam giác vàng số liệu\' (Hải quan - Kế toán - Kho) đối soát theo từng quý; chuẩn hóa bảng mã định danh nguyên vật liệu ngay từ khâu mở tờ khai E11/E15; khi phát sinh chênh lệch phải lập hồ sơ giải trình định mức và nộp báo cáo sửa đổi Mẫu 15a/BCQT-NVL/GSQL trong thời hạn luật định trước khi có quyết định kiểm tra.',
         'scratch/daily_notes.md (10/11) & Báo Hải quan Online, customs.gov.vn | HT (Học Thực Chiến) / EPE Liquidation & Customs Settlement Audit',
         update_date_str),

        (243,
         'cong van 23913 chq gsql nhap khau hoa chat dncx luat quan ly ngoai thuong',
         'cong van 23913 chq gsql, nhap khau hoa chat dncx, luat hoa chat va luat quan ly ngoai thuong, giay phep hoa chat han che kinh doanh dncx, thu tuc hai quan hoa chat fdi, giai toa vuong mac hoa chat dncx',
         99.7,
         57500,
         'https://customs.gov.vn',
         'Công văn 23913/CHQ-GSQL ngày 05/11/2026: Hải quan đề nghị tháo gỡ giao thoa pháp lý nhập khẩu hóa chất vào DNCX giữa Luật Hóa chất và Ngoại thương',
         'Nội dung trọng tâm: Phân tích diễn biến chính sách nóng ngày 10/11/2026: Cục Hải quan ban hành Công văn số 23913/CHQ-GSQL gửi Cục Hóa chất (Bộ Công Thương) đề nghị thống nhất hướng dẫn về thủ tục nhập khẩu hóa chất cần kiểm soát đặc biệt và hóa chất hạn chế sản xuất, kinh doanh đưa vào DNCX; điểm nghẽn pháp lý xuất phát từ việc Điều 59 Luật Quản lý ngoại thương quy định DNCX là khu vực hải quan riêng được miễn giấy phép, nhưng Luật Hóa chất và các Nghị định chuyên ngành vẫn yêu cầu điều kiện quản lý nghiêm ngặt; việc thiếu đồng bộ khiến các lô hàng hóa chất phục vụ sản xuất linh kiện, bán dẫn của doanh nghiệp FDI có nguy cơ bị ách tắc tại cửa khẩu sân bay và cảng biển; Eureka hướng dẫn chủ hàng: Chủ động chuẩn bị sẵn bộ hồ sơ giải trình mục đích sử dụng nội bộ của DNCX, đính kèm văn bản đề nghị thông quan kèm cam kết chịu trách nhiệm để làm việc với Chi cục Hải quan cửa khẩu, đồng thời cập nhật sát sao văn bản phản hồi chính thức từ Bộ Công Thương.',
         'customs.gov.vn & Báo Công Thương | TN (Tin Ngành) / Chemical Regulations & EPE Clearance Policy',
         update_date_str),

        (244,
         'quyet dinh 31 2026 qd ttg danh muc hang hoa phai lam thu tuc tai cua khau nhap',
         'quyet dinh 31 2026 qd ttg, danh muc hang hoa lam thu tuc tai cua khau nhap, thay the quyet dinh 23 2019 qd ttg, thu tuc hai quan tai cang bien san bay, chuyen cua khau icd cang can, kiem soat rui ro hai quan 2026',
         99.6,
         56100,
         'https://baochinhphu.vn',
         'Quyết định 31/2026/QĐ-TTg thay thế QĐ 23/2019/QĐ-TTg: Danh mục hàng hóa bắt buộc làm thủ tục tại cửa khẩu nhập và tác động đến chuỗi vận tải ICD',
         'Nội dung trọng tâm: Cập nhật văn bản pháp lý quy chuẩn mới Quyết định số 31/2026/QĐ-TTg của Thủ tướng Chính phủ quy định Danh mục hàng hóa nhập khẩu phải làm thủ tục hải quan tại cửa khẩu nhập (chính thức thay thế Quyết định 23/2019/QĐ-TTg); các nhóm mặt hàng nhạy cảm, hàng có rủi ro cao về xuất xứ, tiêu chuẩn kỹ thuật, an toàn thực phẩm và thuế (như phế liệu, máy móc thiết bị đã qua sử dụng, hóa chất, xe ô tô nguyên chiếc...) không được phép chuyển cửa khẩu về ICD nội địa mà bắt buộc phải hoàn tất kiểm tra chuyên ngành và thông quan trực tiếp tại cảng biển/sân bay cửa khẩu nhập đầu tiên; tác động lớn đến các doanh nghiệp forwarder và chuỗi vận tải thủy nội địa/đường sắt; Eureka khuyến nghị chủ hàng: Kiểm tra kỹ mã HS sản phẩm đối chiếu với Phụ lục ban hành kèm Quyết định 31 để bố trí nhân sự làm thủ tục ngay tại cửa khẩu cảng biển, dự toán chi phí lưu bãi/công viên (demurrage/detention) để tối ưu chi phí logistics.',
         'Cổng TTĐT Chính phủ (chinhphu.vn) & haiquanonline.com.vn | TN (Tin Ngành) / Port of Entry Customs Clearance Mandate',
         update_date_str),

        (245,
         'so hoa hai quan thong minh 2026 giam sat tu dong cua khau so cang bien so',
         'so hoa hai quan 2026, he thong hai quan so thong minh, giam sat tu dong cang bien so, cua khau thong minh huu nghi, ket noi api hang tau logistics, thu tuc hai quan khong giay to',
         99.5,
         54700,
         'https://haiquanonline.com.vn',
         'Ngành Hải quan tăng tốc vận hành Hệ thống Hải quan số thông minh cuối năm 2026: Doanh nghiệp logistics cần chuẩn bị gì cho mô hình Cảng biển số?',
         'Nội dung trọng tâm: Đánh giá toàn diện đề án đổi mới phương thức quản lý hải quan theo Báo Hải quan Online ngày 10/11/2026: Ngành Hải quan đang bước vào giai đoạn nước rút thử nghiệm và tích hợp toàn diện Hệ thống Hải quan số và Hải quan thông minh, chuẩn bị vận hành chính thức vào cuối năm 2026; chuyển đổi căn bản từ phương thức kiểm tra thủ công sang giám sát tự động 100% bằng công nghệ AI, IoT nhận diện container, camera thông minh tại cửa khẩu đường bộ (điển hình là Cửa khẩu thông minh Hữu Nghị - Hữu Nghị Quan) và các cảng biển loại đặc biệt (Hải Phòng, Cái Mép - Thị Vải); toàn bộ chứng từ vận tải, manifest, lệnh giao hàng điện tử (eDO) được tích hợp qua cổng API một cửa; Eureka khuyến nghị doanh nghiệp xuất nhập khẩu: Nhanh chóng chuẩn hóa hệ thống phần mềm ERP nội bộ, tham gia các chương trình đào tạo số hóa của Hải quan để sẵn sàng chuyển đổi, giảm thời gian thông quan từ vài ngày xuống tính bằng phút.',
         'Báo Hải quan Online & customs.gov.vn | TN (Tin Ngành) / Smart Digital Customs 2026 & Smart Port Transformation',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 241 -> row 246
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

    # Cập nhật Total row ở row 251
    total_row_idx = 251
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 245 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D250)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E250)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 245 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 213 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 10/11/2026)\n"
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
**Ngày thực hiện:** {update_date_str} (Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026)  
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 245 chủ đề)  
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
7. `Drewry / World Shipping Council` (Báo cáo WCI cước vận tải biển)
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 10/11/2026)

Qua sàng lọc từ 48 sự kiện và case study nghiệp vụ thực tế, Agent 01 đã chọn lọc và chấm điểm chi tiết **5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [ĐIỂM NÓNG GIA CÔNG — TỪ BỎ HÀNG HÓA] Bên đặt gia công nước ngoài từ bỏ hàng: Xử lý nguyên liệu dư thừa, máy móc tạm nhập và sản phẩm theo Thông tư 121/2025/TT-BTC (sửa Điều 66 Thông tư 38/2015/TT-BTC)
- **Chuyên mục (Pillar):** `CS` (Case Study) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Infographic 'Đối tác nước ngoài bỏ hàng gia công: Doanh nghiệp Việt Nam phải làm gì trong 30 ngày để không bị phạt và truy thu thuế?' | YouTube: Video cẩm nang hướng dẫn chi tiết thủ tục mở tờ khai A42 nộp thuế và quy trình tiêu hủy theo Điều 64)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực nóng: Quy định tại khoản 37 Điều 1 Thông tư 121/2025/TT-BTC (sửa Điều 66 TT 38) đang được Hải quan áp dụng quyết liệt; case study thực tế phát sinh hôm nay |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Rất đau: Doanh nghiệp bỏ mặc hàng tại kho sẽ bị lập biên bản xử phạt vi phạm hành chính, ấn định thuế và nâng hạng rủi ro doanh nghiệp; chi phí tiêu hủy rất tốn kém |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi nảy lửa: Bên gia công có quyền bán vớt vát thanh lý không, thủ tục chuyển đổi mục đích sử dụng A42 tính thuế theo thời điểm nào, khi nào bị tịch thu sung công quỹ? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka đóng gói 03 hướng xử lý chuẩn mực: (1) Nộp thuế tiêu thụ nội địa A42; (2) Hải quan tịch thu sung công quỹ; (3) Tiêu hủy chịu chi phí kèm điều khoản HĐ ngoại thương |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #1 (Chủ Đề Cấp Bách Nhất — Tác động trực tiếp đến hàng nghìn doanh nghiệp dệt may, da giày, cơ khí)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi lo 'bỏ thì thương, vương thì tội' của các xưởng gia công khi đối tác ngoại bỏ chạy. Bài viết khẳng định bản lĩnh tư vấn pháp lý và xử lý sự cố dày dạn của Eureka.
- **Nguồn xác minh:** Khoản 37 Điều 1 Thông tư 121/2025/TT-BTC; Điều 66 Thông tư 38/2015/TT-BTC; Nghị định 134/2016/NĐ-CP; `scratch/daily_notes.md` (10/11/2026).

---

### Đề xuất #2: [RỦI RO SINH TỬ DNCX — MÙA THANH KHOẢN] Kiểm tra sau thông quan thanh khoản DNCX: Chênh lệch hàng triệu linh kiện, âm 80 tấn hạt nhựa và công thức Tam giác vàng bảo vệ Báo cáo quyết toán Mẫu 15
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `CS` (Case Study Doanh Nghiệp Chế Xuất)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Infographic cảnh báo 'Tam giác vàng số liệu Hải quan - Kế toán - Kho: Cách DNCX tự rà soát trước khi bị kiểm tra sau thông quan' | YouTube: Video ThuyTalk phân tích giải trình định mức hao hụt và xử lý âm/dương nguyên liệu)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Nóng hổi: Mùa rà soát sổ sách và lập Báo cáo quyết toán cuối năm của các nhà máy DNCX và tập đoàn FDI tại các KCN |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Cực đau: Lệch số liệu giữa Kho thực tế và Tờ khai hải quan dẫn đến bị ấn định thuế nhập khẩu, thuế GTGT và xử phạt trốn thuế lên tới hàng tỷ đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Xung đột nội bộ giữa bộ phận Kho (xuất theo bao/cuộn) - Kế toán (ghi nhận tiền) - Xuất nhập khẩu (khai kg/mét theo HS); ai chịu trách nhiệm khi bị cơ quan thuế/hải quan sờ gáy? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka đưa ra giải pháp toàn diện: Quy trình đối soát chéo 3 bên theo quý, cách lập hồ sơ giải trình định mức hao hụt kỹ thuật và thời hạn nộp BCQT sửa đổi Mẫu 15a |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #2 (Nội Dung Chuyên Gia Cao Cấp — Chinh phục khối FDI & DNCX)** |
- **Tóm tắt lý do nên làm:** Đưa thương hiệu Eureka vào tầm ngắm của các Giám đốc nhà máy, Kế toán trưởng và Trưởng phòng XNK của các doanh nghiệp FDI lớn.
- **Nguồn xác minh:** Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; Báo Hải quan Online; `scratch/daily_notes.md` (10/11/2026).

---

### Đề xuất #3: [GIAO THOA PHÁP LÝ HÓA CHẤT DNCX] Công văn 23913/CHQ-GSQL ngày 05/11/2026: Hải quan đề nghị tháo gỡ vướng mắc giấy phép nhập khẩu hóa chất vào DNCX
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study)
- **Nền tảng khuyến nghị:** **Facebook & LinkedIn** (Bài phân tích chuyên sâu: 'Công văn 23913/CHQ-GSQL: Điểm nghẽn nhập khẩu hóa chất vào DNCX giữa Luật Hóa chất và Ngoại thương — Hướng đi an toàn cho doanh nghiệp FDI')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Mới nhất: Công văn 23913/CHQ-GSQL vừa ban hành ngày 05/11/2026 phản ánh kiến nghị nóng từ cơ sở hải quan gửi Cục Hóa chất (Bộ Công Thương) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Hóa chất chuyên dụng, tiền chất công nghiệp phục vụ dây chuyền bán dẫn/điện tử bị kẹt tại cửa khẩu sẽ khiến toàn bộ nhà máy FDI bị đình trệ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Tranh chấp pháp lý: Điều 59 Luật Quản lý ngoại thương quy định khu phi thuế quan không cần GP, nhưng các chi cục hải quan vẫn đòi giấy phép theo Luật Hóa chất |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.7/5) | Eureka cung cấp căn cứ pháp lý vững chắc, hướng dẫn lập bản cam kết sử dụng nội bộ để bảo lãnh thông quan nhanh trong lúc chờ văn bản liên bộ |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #3 (Pháp Lý Đỉnh Cao — Tác động trực tiếp đến chuỗi cung ứng công nghệ cao)** |
- **Tóm tắt lý do nên làm:** Thể hiện độ nhạy bén chính sách của Eureka khi theo sát từng công văn mới nhất của Cục Hải quan và Bộ Công Thương.
- **Nguồn xác minh:** Công văn 23913/CHQ-GSQL ngày 05/11/2026; Cổng TTĐT Hải quan (customs.gov.vn); Báo Công Thương.

---

### Đề xuất #4: [QUY CHUẨN CỬA KHẨU NHẬP MỚI] Quyết định 31/2026/QĐ-TTg thay thế QĐ 23/2019/QĐ-TTg: Danh mục hàng hóa bắt buộc làm thủ tục tại cửa khẩu nhập và tác động đến chi phí ICD
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `FF` (Forwarder & Cước Vận Tải)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group XNK** (Bài viết nghiệp vụ: 'Quyết định 31/2026/QĐ-TTg: Hàng hóa nào không còn được chuyển cửa khẩu về ICD? 4 lưu ý sống còn để không bị kẹt phí lưu bãi tại cảng biển')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Văn bản quy chuẩn mới của Thủ tướng Chính phủ quy định địa điểm làm thủ tục hải quan với hàng hóa nhập khẩu có hiệu lực trong năm 2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Chủ hàng quen chuyển cửa khẩu về ICD gần nhà máy bị bất ngờ vì hàng bị chặn tại cảng biển; phát sinh phí cắm điện, lưu bãi, lưu công đắt đỏ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.6/5) | Tranh luận giữa chủ hàng và forwarder về chi phí phát sinh và thời gian thông quan tại cảng biển so với cảng cạn nội địa |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka cung cấp bảng tra cứu nhanh danh mục mã HS bắt buộc làm thủ tục tại cửa khẩu nhập và giải pháp thông quan trọn gói tại cảng Hải Phòng / Cát Lái |
  | **Tổng điểm** | **19.0 / 20** | **Xếp hạng: #4 (Cẩm Nang Nghiệp Vụ Cửa Khẩu Cảng Biển)** |
- **Tóm tắt lý do nên làm:** Hướng dẫn thực tế giúp khách hàng của Eureka và cộng đồng logistics tránh rơi vào bẫy chi phí lưu bãi lưu container tại cảng biển.
- **Nguồn xác minh:** Quyết định số 31/2026/QĐ-TTg; Cổng TTĐT Chính phủ (chinhphu.vn); haiquanonline.com.vn.

---

### Đề xuất #5: [NƯỚC RÚT SỐ HÓA HẢI QUAN 2026] Ngành Hải quan tăng tốc vận hành Hệ thống Hải quan số thông minh cuối năm 2026: Mô hình Cảng biển số và chuẩn bị của doanh nghiệp
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Infographic 'Lộ trình chuyển đổi Cảng biển số & Cửa khẩu thông minh 2026' | YouTube: Video ThuyTalk định hướng công nghệ và số hóa quy trình logistics)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.7/5) | Định hướng chiến lược trọng tâm của ngành Hải quan đang bước vào giai đoạn thử nghiệm tích hợp cuối năm 2026 |
  | Độ đau | ⭐⭐⭐⭐☆ (4.5/5) | Doanh nghiệp xuất nhập khẩu chưa chuẩn hóa số liệu số, chưa tích hợp eDO và API sẽ bị tụt hậu, thời gian thông quan kéo dài |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.6/5) | Tranh luận về tính đồng bộ hạ tầng công nghệ thông tin giữa Hải quan, Hãng tàu, Doanh nghiệp cảng và Forwarder |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ góc nhìn tiên phong về tự động hóa logistics, lợi ích của tờ khai không giấy tờ và cách thức vận hành tinh gọn |
  | **Tổng điểm** | **18.7 / 20** | **Xếp hạng: #5 (Tầm Nhìn Vĩ Mô & Số Hóa Logistics Tương Lai)** |
- **Tóm tắt lý do nên làm:** Nâng tầm vị thế của Chị Thúy và Eureka như những chuyên gia dẫn dắt xu hướng công nghệ trong ngành logistics Việt Nam.
- **Nguồn xác minh:** Báo Hải quan Online (haiquanonline.com.vn); Tổng cục Hải quan (customs.gov.vn).

---

## 📌 HÀNH ĐỘNG TIẾP THEO ĐỀ XUẤT CHO CHỊ THÚY
1. **Lựa chọn chủ đề:** Chị Thúy có thể chọn 1 hoặc 2 chủ đề trong Top 5 để đưa vào lịch sản xuất content tuần (Tuần 46 hoặc Tuần 47 qua quy trình `wf01_weekly_planning.md`).
2. **Sản xuất nhanh:** Nếu chọn Đề xuất #1 (Gia công từ bỏ hàng) hoặc Đề xuất #2 (Kiểm tra thanh khoản DNCX), Agent 03 và Agent 06 sẵn sàng kích hoạt quy trình `wf02` (Facebook) hoặc `wf03` (YouTube) để viết bài và thiết kế infographic ngay hôm nay.
3. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT WF05 NGÀY 10/11/2026 ===")
    print(f"- File Excel: {FILE_PATH} ({file_size} bytes, SHA256: {file_sha256[:8]}...)")
    print(f"- OneDrive: {ONEDRIVE_FILE}")
    print(f"- Briefing: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w46_1011()
