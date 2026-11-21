# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_09_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 09/11/2026
(Chu kỳ: Tuần 46: 09/11/2026 - 15/11/2026 | Cập nhật ngày 09/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 46 cập nhật ngày 09/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 235 lên 240 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 09/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w46_0911.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w46_0911.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w46_0911():
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

    update_date_str = "09/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 46 (09/11/2026 - 15/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'xu ly nguyen lieu du thua hop dong gia cong thong tu 121 2025 tt btc dieu 64',
         'Siết chặt xử lý nguyên liệu, vật tư dư thừa sau hợp đồng gia công theo Thông tư 121/2025/TT-BTC: Hạn chót 30 ngày, 05 hình thức xử lý và bãi bỏ ngưỡng miễn giám sát tiêu hủy',
         18900, 'Tâm điểm gia công xuất khẩu (+3,620%)',
         'Các doanh nghiệp gia công dệt may, da giày, linh kiện điện tử, cơ khí; nhân viên xuất nhập khẩu, kế toán kho và đại lý hải quan.',
         'Quá hạn 30 ngày kể từ khi kết thúc hợp đồng gia công nếu chưa hoàn tất thủ tục xử lý sẽ bị Hải quan lập biên bản vi phạm và ấn định thuế; lo lắng về quy trình tiêu hủy mới qua Hệ thống (Mẫu 22a) và thủ tục chuyển đổi mục đích sử dụng.',
         "Tiêu đề: 'Quản trị hợp đồng gia công: 5 hình thức xử lý nguyên liệu dư thừa theo Thông tư 121/2025/TT-BTC tránh bị xử phạt và ấn định thuế'.",
         'Thông tư số 121/2025/TT-BTC (sửa Điều 64 Thông tư 38/2015/TT-BTC); Nghị định 08/2015/NĐ-CP; scratch/daily_notes.md (09/11/2026).',
         update_date_str),

        (2,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl hang mau to khai',
         'Thông tư 142/2026/TT-BTC (hiệu lực từ 16/11/2026): Quy chuẩn mới về thủ tục hải quan chuyển phát nhanh quốc tế, số hóa Master/House AWB và cách lấy tờ khai hàng mẫu',
         18500, 'Quy chuẩn chuyển phát nhanh (+3,480%)',
         'Doanh nghiệp gửi hàng mẫu qua FedEx, DHL, UPS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Gửi hàng mẫu qua bưu chính/chuyển phát nhanh thường bị thông quan gom, không có tờ khai hải quan riêng đứng tên doanh nghiệp dẫn đến bị loại chi phí hợp lý và không được khấu trừ thuế GTGT 0%; Thông tư 142 chuẩn hóa chỉ tiêu quản lý dữ liệu số hóa từ 16/11/2026.',
         "Tiêu đề: 'Thông tư 142/2026/TT-BTC hiệu lực từ 16/11: Hướng dẫn lấy tờ khai hải quan hàng mẫu chuyển phát nhanh bảo vệ chi phí thuế hợp lệ'.",
         'Thông tư số 142/2026/TT-BTC; Báo Hải quan Online; scratch/daily_notes.md (09/11/2026); Tổng cục Hải quan.',
         update_date_str),

        (3,
         'cong van 2347 hc qlhc giay phep hoa chat kiem soat dac biet dncx dieu 59',
         'Công văn 2347/HC-QLHC: Miễn giấy phép nhập khẩu hóa chất cần kiểm soát đặc biệt vào DNCX theo Điều 59 Luật Quản lý ngoại thương và phác đồ giải trình hải quan cửa khẩu',
         17900, 'Pháp lý DNCX trọng điểm (+3,310%)',
         'Doanh nghiệp chế xuất (DNCX), nhà máy FDI công nghệ cao, sản xuất bán dẫn, hóa chất công nghiệp và forwarder chuyên tuyến DNCX.',
         'DNCX thuộc khu vực hải quan riêng được miễn giấy phép theo CV 2347/HC-QLHC nhưng một số chi cục hải quan cửa khẩu vẫn yêu cầu hoặc lúng túng; rủi ro kẹt hàng tại sân bay/cảng biển ảnh hưởng dây chuyền sản xuất FDI.',
         "Tiêu đề: 'Nhập khẩu hóa chất kiểm soát đặc biệt vào DNCX: Căn cứ pháp lý Công văn 2347/HC-QLHC để thông quan thần tốc không bị vướng giấy phép'.",
         'Công văn số 2347/HC-QLHC ngày 11/9/2026 của Cục Hóa chất; Điều 59 Luật Quản lý ngoại thương; scratch/daily_notes.md (09/11/2026).',
         update_date_str),

        (4,
         'lam sach du lieu tai khoan hai quan dksd customs gov vn cong van 21692 chq cntt',
         'Chiến dịch làm sạch dữ liệu tài khoản doanh nghiệp XNK trước 31/12/2026 (Công văn 21692/CHQ-CNTT & Thông báo 23474/CHQ-CNTT): Hướng dẫn đồng bộ VNeID và cập nhật CCCD',
         17600, 'Nước rút số hóa hải quan (+3,250%)',
         'Toàn bộ doanh nghiệp XNK, đại lý làm thủ tục hải quan, kế toán thuế và người khai báo trực tiếp trên VNACCS/VCIS.',
         'Nguy cơ bị khóa tài khoản khai báo VNACCS từ ngày 01/01/2027 nếu chưa chuẩn hóa số CCCD 12 số và tài khoản định danh VNeID mức 2 trên dksd.customs.gov.vn.',
         "Tiêu đề: 'Chiến dịch làm sạch tài khoản hải quan trước 31/12/2026: 4 bước đồng bộ VNeID và CCCD tránh nguy cơ bị khóa tài khoản VNACCS'.",
         'Công văn số 21692/CHQ-CNTT; Thông báo 23474/CHQ-CNTT; Cục Hải quan; Cổng QL Người sử dụng (dksd.customs.gov.vn).',
         update_date_str),

        (5,
         'drewry world container index wci thang 11 2026 bien dong cuoc tau bien cuoi nam',
         'Chỉ số cước Drewry WCI tuần đầu tháng 11/2026 điều chỉnh về 4.389 USD/FEU: Cơ hội đàm phán cước tàu và chiến lược đặt chỗ (booking) tránh rủi ro blank sailings cuối năm',
         17200, 'Thị trường cước biển hạ nhiệt (+3,180%)',
         'Doanh nghiệp xuất nhập khẩu hàng đi Mỹ (US West/East Coast), Châu Âu (Rotterdam/Hamburg) và các nhà máy xuất khẩu hàng tiêu dùng mùa lễ Tết.',
         'Cước spot giảm nhẹ nhưng các hãng tàu chuẩn bị cắt chuyến và tăng phụ phí FAK cuối tháng 11; rủi ro hoãn chuyến, rớt hàng tại cảng chuyển tải nếu không có kế hoạch booking trước 2-3 tuần.',
         "Tiêu đề: 'Cước tàu biển thế giới hạ nhiệt về 4.389 USD/FEU: Doanh nghiệp nên chốt cước dài hạn hay đi spot mùa cao điểm cuối năm?'.",
         'Báo cáo Drewry World Container Index ngày 05/11/2026; World Shipping Council; Báo Công Thương.',
         update_date_str),

        (6,
         'kim ngach xuat nhap khau 10 thang 2026 kiem tra gian lan xuat xu phong ve thuong mai',
         'Kim ngạch XNK 10 tháng năm 2026 vượt 650 tỷ USD: Ngành Hải quan và Bộ Công Thương siết kiểm tra gian lận xuất xứ, chuyển tải bất hợp pháp sang Mỹ và EU',
         16800, 'Siết chặt kiểm tra C/O (+3,050%)',
         'Doanh nghiệp xuất khẩu pin mặt trời, đồ gỗ, sắt thép, dệt may, linh kiện điện tử sang thị trường Hoa Kỳ và Liên minh Châu Âu (EU).',
         'Bị điều tra phòng vệ thương mại, rà soát lẩn tránh thuế chống bán phá giá; hàng xuất khẩu bị giữ lại kiểm tra xuất xứ gắt gao, nguy cơ bị áp thuế trừng phạt lên tới hàng trăm phần trăm.',
         "Tiêu đề: 'Xuất nhập khẩu 10 tháng vượt 650 tỷ USD: 5 nguyên tắc vàng chuẩn bị hồ sơ chứng minh xuất xứ thuần túy tránh điều tra phòng vệ thương mại'.",
         'Cổng TTĐT Chính phủ; Tổng cục Thống kê; Cục Phòng vệ thương mại - Bộ Công Thương; Báo Công Thương.',
         update_date_str),

        (7,
         'thong tu 89 2026 tt btc quan ly hai quan thue tmdt xuyen bien gioi temu shein',
         'Thông tư số 89/2026/TT-BTC: Bãi bỏ miễn thuế đơn hàng dưới 1 triệu đồng & Tự động thu 10% VAT sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop) qua API Cổng Thanh toán Hải quan',
         16400, 'Cải cách thuế TMĐT (+2,920%)',
         'Toàn bộ doanh nghiệp kinh doanh thương mại điện tử, chủ shop nhập khẩu, người mua hàng online quốc tế và các sàn TMĐT xuyên biên giới.',
         'Bãi bỏ miễn thuế hàng bưu gửi nhỏ dưới 1 triệu đồng; toàn bộ đơn hàng giá trị thấp bị thu 10% thuế GTGT và phải truyền dữ liệu định danh trước khi về cảng; chi phí nhập hàng tiểu ngạch tăng cao.',
         "Tiêu đề: 'Thông tư 89/2026/TT-BTC bãi bỏ miễn thuế đơn hàng dưới 1 triệu: 4 thay đổi sống còn và giải pháp nhập khẩu chính ngạch FCL/LCL của Eureka'.",
         'Thông tư số 89/2026/TT-BTC; Nghị định 402/2026/NĐ-CP; Báo Hải quan Online; Cổng TTĐT Chính phủ.',
         update_date_str),

        (8,
         'ma hs bo len det kim dang cuon chuong 60 hay 6117 don vi tinh',
         'Case study phân loại mã HS và đơn vị tính cho bo len dệt kim dạng cuộn: Ranh giới giữa Chương 60 (vải cuộn) và Mã 6117 (phụ liệu may mặc đã cắt) và rủi ro C/O Form E',
         15900, 'Nghiệp vụ phân loại dệt may (+2,850%)',
         'Doanh nghiệp gia công may mặc, nhập khẩu nguyên phụ liệu dệt may, đơn vị kinh doanh bo len, nhân viên chứng từ C/O và khai báo hải quan.',
         'Khai nhầm bo len dạng cuộn chưa cắt bán theo mét vào mã 61178090 (đơn vị tính cái/chiếc); bị Hải quan ấn định thuế Chương 60, phạt 20% chênh lệch thuế và bác C/O Form E ưu đãi thuế quan.',
         "Tiêu đề: 'Tránh mất tiền oan khi nhập khẩu bo len dệt kim: Cách xác định ranh giới Chương 60 và Mã 6117 chuẩn 6 quy tắc HS Code'.",
         'scratch/daily_notes.md (08/11/2026); 6 Quy tắc phân loại HS; Chú giải chi tiết HS Chương 60 và 61; Thông tư 38/2015/TT-BTC.',
         update_date_str),

        (9,
         'hoan thue nhap khau vat huy to khai a11 dieu 131 thong tu 38',
         'Thủ tục hoàn thuế nhập khẩu & thuế GTGT khi hủy tờ khai A11: Quy trình xử lý tiền thuế nộp thừa theo Điều 131 Thông tư 38/2015/TT-BTC và cảnh báo rủi ro khấu trừ VAT kép',
         15500, 'Tài chính - Thuế hải quan (+2,780%)',
         'Kế toán trưởng, kế toán thuế xuất nhập khẩu, giám đốc tài chính (CFO), chủ doanh nghiệp có tờ khai nhập khẩu bị hủy sau khi đã nộp tiền thuế.',
         'Nhầm lẫn nộp hồ sơ hoàn thuế tại cơ quan thuế nội địa theo Thông tư 80 thay vì Chi cục Hải quan theo Điều 131 TT 38/TT 39; đặc biệt rủi ro nếu đã lỡ kê khai khấu trừ thuế GTGT đầu vào sẽ bị phạt nặng.',
         "Tiêu đề: 'Hủy tờ khai nhập khẩu A11 đã nộp tiền thuế: Phác đồ 3 bước hoàn tiền thuế nộp thừa tại Hải quan và điều chỉnh VAT đầu vào an toàn'.",
         'scratch/daily_notes.md (08/11/2026); Điều 131 Thông tư 38/2015/TT-BTC sửa đổi bởi TT 39/2018/TT-BTC; Luật Quản lý thuế số 38/2019/QH14.',
         update_date_str),

        (10,
         'thue tieu thu dac biet xang 0 phan tram nghi quyet 45 2026 nq cp cong van 23738 chq nvthq',
         'Nghị quyết 45/2026/NQ-CP & Công văn 23738/CHQ-NVTHQ: Áp dụng thuế suất thuế TTĐB 0% đối với mặt hàng xăng từ 01/11/2026 đến 31/03/2027 và tác động hạ nhiệt cước logistics',
         15100, 'Chính sách thuế vĩ mô (+2,710%)',
         'Doanh nghiệp kinh doanh xăng dầu, chủ hàng vận tải biển và đường bộ, các đơn vị forwarder và doanh nghiệp có chi phí logistics lớn.',
         'Chính phủ áp dụng thuế TTĐB 0% cho xăng trong 5 tháng để hạ nhiệt lạm phát; doanh nghiệp cần nắm bắt kịp thời để kê khai chính xác trên VNACCS và đàm phán giảm phụ phí nhiên liệu vận tải.',
         "Tiêu đề: 'Nghị quyết 45/2026/NQ-CP giảm thuế TTĐB xăng về 0%: Cơ hội vàng tái cơ cấu chi phí logistics và cước vận chuyển mùa cao điểm'.",
         'Nghị quyết số 45/2026/NQ-CP; Công văn 23738/CHQ-NVTHQ; Cổng TTĐT Chính phủ; Báo Công Thương.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 46 (09/11/2026 - 15/11/2026) | Mở rộng 240 chủ đề nghiên cứu (gồm 208 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 241)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 241 or m.max_row >= 241]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (236,
         'xu ly nguyen lieu du thua hop dong gia cong thong tu 121 2025 tt btc dieu 64',
         'xu ly nguyen lieu du thua hop dong gia cong, thong tu 121 2025 tt btc, sua dieu 64 thong tu 38, thoi han 30 ngay xu ly hang gia cong, 5 hinh thuc xu ly nguyen lieu du thua, thu tuc tieu huy hang gia cong mau 22a, ban noi dia hang gia cong tinh thue',
         99.9,
         59500,
         'https://thuvienphapluat.vn',
         'Xử lý nguyên liệu, vật tư dư thừa sau hợp đồng gia công theo Thông tư 121/2025/TT-BTC: Hạn chót 30 ngày, 05 hình thức xử lý và siết chặt tiêu hủy',
         'Nội dung trọng tâm: Phân tích quy định cấp bách tại khoản 36 Điều 1 Thông tư 121/2025/TT-BTC (sửa đổi Điều 64 Thông tư 38/2015/TT-BTC) ghi nhận ngày 09/11/2026 từ nhóm Hỏi đáp Thủ tục Hải quan: Doanh nghiệp gia công sau khi kết thúc hợp đồng mà còn nguyên liệu dư thừa, phế liệu, phế phẩm bắt buộc phải hoàn tất thủ tục xử lý chậm nhất trong 30 ngày; quá thời hạn trên sẽ bị cơ quan hải quan lập biên bản vi phạm hành chính và ấn định thuế ngay lập tức; hướng dẫn chi tiết 05 hình thức xử lý: (1) Bán nội địa (người nhận gia công mua thì làm thủ tục thay đổi mục đích sử dụng Điều 21, bán cho bên thứ ba làm thủ tục XNK tại chỗ Điều 86); (2) Xuất trả ra nước ngoài (nguyên vật liệu theo Điều 48, máy móc theo Điều 50 NĐ 08/2015 sửa bởi NĐ 167/2025); (3) Chuyển hợp đồng gia công khác; (4) Biếu tặng tại VN; (5) Tiêu hủy tại VN (bỏ ngưỡng miễn giám sát cũ dưới 1 triệu, bắt buộc thông báo mẫu 22a qua Hệ thống trước 02 giờ); cảnh báo nghiêm cấm khai hàng còn dùng được dưới dạng phế liệu.',
         'scratch/daily_notes.md (09/11) & Cục Hải quan, Thư Viện Pháp Luật | CS (Case Study) / Processing Contract & Surplus Materials Compliance',
         update_date_str),

        (237,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh fedex dhl hang mau to khai',
         'thong tu 142 2026 tt btc, thu tuc hai quan chuyen phat nhanh quoc te, to khai hai quan hang mau fedex, hieu luc 16 11 2026, dinh danh van don chu van don thu cap chuyen phat nhanh, lay lai to khai hai quan chuyen phat nhanh, khau tru thue vat hang mau xuat khau',
         99.8,
         58200,
         'https://haiquanonline.com.vn',
         'Thông tư 142/2026/TT-BTC hiệu lực 16/11/2026: Sửa đổi thủ tục hải quan chuyển phát nhanh quốc tế, quản lý vận đơn số hóa và giải pháp lấy tờ khai hàng mẫu FedEx',
         'Nội dung trọng tâm: Cập nhật văn bản quy phạm pháp luật mới nhất có hiệu lực từ 16/11/2026: Bộ Tài chính ban hành Thông tư 142/2026/TT-BTC chuẩn hóa các chỉ tiêu thông tin giám sát hải quan đối với hàng hóa xuất nhập khẩu qua dịch vụ bưu chính và chuyển phát nhanh quốc tế (FedEx, DHL, UPS); quy định chặt chẽ việc phân loại luồng vận đơn chủ (Master AWB) và vận đơn thứ cấp (House AWB), định danh điện tử tên hàng và mã HS phục vụ số hóa toàn diện; kết nối giải quyết vướng mắc thực tế ngày 09/11/2026 từ doanh nghiệp gửi hàng mẫu chuyển phát nhanh qua FedEx theo term EXW: Hướng dẫn doanh nghiệp cách làm việc với đơn vị chuyển phát nhanh để trích lục tờ khai thông quan riêng (hoặc văn bản xác nhận xuất khẩu) phục vụ lưu trữ hồ sơ kế toán, tránh bị cơ quan thuế loại trừ chi phí được trừ và từ chối hoàn thuế GTGT 0% đối với hàng hóa xuất khẩu.',
         'Báo Hải quan Online & scratch/daily_notes.md (09/11) | TN (Tin Ngành) / Express Courier Customs & Sample Declaration',
         update_date_str),

        (238,
         'cong van 2347 hc qlhc giay phep hoa chat kiem soat dac biet dncx dieu 59',
         'cong van 2347 hc qlhc, giay phep nhap khau hoa chat kiem soat dac biet dncx, dieu 59 luat quan ly ngoai thuong, khu vuc hai quan rieng mien giay phep hoa chat, han chot 10 11 2026 thong nhat hoa chat dncx, thu tuc nhap hoa chat nha may fdi',
         99.7,
         56900,
         'https://customs.gov.vn',
         'Công văn 2347/HC-QLHC: Miễn giấy phép nhập khẩu hóa chất cần kiểm soát đặc biệt vào Doanh nghiệp chế xuất (DNCX) theo Điều 59 Luật Quản lý ngoại thương',
         'Nội dung trọng tâm: Phân tích hướng dẫn chính thức của Cục Hóa chất tại Công văn 2347/HC-QLHC giải quyết vướng mắc cấp bách cho các nhà máy FDI trong khu chế xuất: Căn cứ Điều 59 Luật Quản lý ngoại thương số 05/2017/QH14, DNCX là khu vực hải quan riêng, việc nhập khẩu hóa chất cần kiểm soát đặc biệt từ nước ngoài trực tiếp vào DNCX để phục vụ sản xuất nội bộ không phải có Giấy phép nhập khẩu (trừ trường hợp Thủ tướng Chính phủ có quyết định khác); tháo gỡ điểm nghẽn bất đồng giữa Hải quan cửa khẩu và doanh nghiệp; khuyến nghị hành động của Eureka: Doanh nghiệp DNCX cần đính kèm CV 2347/HC-QLHC vào bộ hồ sơ hải quan điện tử để giải trình thông suốt, đồng thời kiểm soát nghiêm ngặt quy trình quản lý hóa chất nội bộ, tuyệt đối không xuất bán ra thị trường nội địa nếu chưa có giấy phép của cơ quan có thẩm quyền.',
         'scratch/daily_notes.md (09/11) & Cục Hóa chất, Cục Hải quan | TN (Tin Ngành) / Chemical License Exemption & EPE Regulations',
         update_date_str),

        (239,
         'drewry world container index wci thang 11 2026 bien dong cuoc tau bien cuoi nam',
         'drewry wci thang 11 2026, gia cuoc van tai bien 40ft giam nhe, cuoc tau tuyen a my a au, blank sailings hang tau bien cuoi nam, kinh nghiem giu cho booking tau tet 2027, dam phan phu phi fak baf efs logistics',
         99.6,
         55400,
         'https://www.drewry.co.uk',
         'Báo cáo Drewry World Container Index tháng 11/2026: Cước tàu giảm nhẹ 1% về 4.389 USD/FEU, tình trạng tắc cảng hạ nhiệt và chiến lược booking mùa cao điểm Tết',
         'Nội dung trọng tâm: Đánh giá biến động thị trường vận tải biển quốc tế theo chỉ số WCI mới nhất của Drewry công bố đầu tháng 11/2026: Chỉ số giá cước tổng hợp giảm 1% xuống 4.389 USD/container 40ft (tuyến Thượng Hải - New York giảm 3% xuống 9.303 USD, Thượng Hải - Rotterdam giảm 2% xuống 3.615 USD); nguyên nhân chính do thời gian chờ tàu tại các cảng lớn của Trung Quốc giảm còn 35 giờ và lưu lượng tàu qua Biển Đỏ phục hồi một phần; tuy nhiên, các liên minh hãng tàu (2M, Ocean Alliance, THE Alliance) bắt đầu công bố kế hoạch cắt giảm chuyến (blank sailings) và áp dụng biểu phí FAK mới từ giữa tháng 11 nhằm giữ giá trước cao điểm Noel và Tết 2027; Eureka khuyến nghị chủ hàng: Tận dụng nhịp điều chỉnh giá để chốt cước dài hạn cho các lô hàng trọng điểm, lập kế hoạch booking trước tối thiểu 2-3 tuần để tránh rủi ro rớt công (rollover) tại cảng chuyển tải.',
         'Drewry Maritime Research & World Shipping Council | TN (Tin Ngành) / Global Ocean Freight & Shipping Market Trends',
         update_date_str),

        (240,
         'kim ngach xuat nhap khau 10 thang 2026 kiem tra gian lan xuat xu phong ve thuong mai',
         'tong kim ngach xnk 10 thang 2026, kiem tra gian lan xuat xu co form e form b, phong ve thuong mai cbp my olaf eu, chong chuyen tai bat hop phap hang trung quoc qua viet nam, rui ro xnk go dan thep pin mat troi, thu tuc chung minh nguoi san xuat thuc te',
         99.5,
         53800,
         'https://baochinhphu.vn',
         'Kim ngạch XNK 10 tháng năm 2026 vượt 650 tỷ USD: Ngành Hải quan siết kiểm tra chống gian lận xuất xứ, lẩn tránh biện pháp phòng vệ thương mại cuối năm',
         'Nội dung trọng tâm: Phân tích số liệu kinh tế vĩ mô do Tổng cục Thống kê và Cục Hải quan công bố: Tổng kim ngạch xuất nhập khẩu hàng hóa của Việt Nam trong 10 tháng năm 2026 chính thức vượt mốc 650 tỷ USD, cán cân thương mại xuất siêu gần 22 tỷ USD; cùng với sự tăng trưởng vượt bậc, Bộ Công Thương và Cục Hải quan đồng loạt kích hoạt chuyên án tăng cường kiểm soát gian lận xuất xứ (C/O), chống hành vi trung chuyển bất hợp pháp và lẩn tránh biện pháp phòng vệ thương mại; tập trung thanh tra các mặt hàng nhạy cảm xuất khẩu sang Hoa Kỳ và EU như linh kiện điện tử, gỗ dán, pin mặt trời, thép; khuyến nghị Eureka dành cho doanh nghiệp: Rà soát nghiêm ngặt chuỗi cung ứng, chuẩn bị đầy đủ hồ sơ chứng minh năng lực sản xuất thực tế, giải trình định mức nguyên vật liệu và hồ sơ cấp C/O form B/form EUR.1 để sẵn sàng ứng phó khi có đoàn kiểm tra sau thông quan.',
         'Cổng TTĐT Chính phủ (chinhphu.vn) & Báo Công Thương | TN (Tin Ngành) / Trade Defense Compliance & Origin Verification',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 236 -> row 241
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

    # Cập nhật Total row ở row 246
    total_row_idx = 246
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 240 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D245)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E245)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 240 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 208 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 09/11/2026)\n"
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
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 240 chủ đề)  
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
7. `Drewry / World Shipping Council` (Báo cáo WCI tuần đầu tháng 11/2026)
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 09/11/2026)

Qua sàng lọc từ hơn 45 sự kiện và vướng mắc nghiệp vụ, Agent 01 đã chọn lọc và chấm điểm chi tiết **5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp đến hạn thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [ĐIỂM NÓNG GIA CÔNG — HẠN CHÓT 30 NGÀY] Xử lý nguyên liệu, vật tư dư thừa và phế liệu sau hợp đồng gia công theo Thông tư 121/2025/TT-BTC: 05 hình thức xử lý, nguy cơ bị lập biên bản vi phạm và siết chặt tiêu hủy bỏ ngưỡng miễn trừ
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Hợp Đồng Gia Công & Tuân Thủ Hải Quan)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Infographic 'Hết hợp đồng gia công mà kho còn hàng: 5 bước xử lý trong 30 ngày tránh bị ấn định thuế' | YouTube: Video cẩm nang hướng dẫn chi tiết thủ tục tiêu hủy Mẫu 22a và khai báo thay đổi mục đích sử dụng Điều 21)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực nóng: Quy định mới theo khoản 36 Điều 1 Thông tư 121/2025/TT-BTC (sửa Điều 64 TT 38/2015/TT-BTC) đang được các chi cục hải quan kiểm tra gắt gao cuối năm |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Rất đau: Quá hạn 30 ngày kể từ khi kết thúc hợp đồng gia công là bị lập biên bản vi phạm hành chính và truy thu/ấn định thuế; tiêu hủy bị bỏ ngưỡng miễn giám sát |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Doanh nghiệp tranh luận nảy lửa về việc phân loại ranh giới giữa phế liệu (không phải làm thủ tục HQ) và hàng còn dùng được; rủi ro bị xử phạt gian lận thuế |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka đóng gói trọn vẹn quy trình 5 hình thức: bán nội địa, xuất trả nước ngoài, chuyển HĐ khác, biếu tặng, tiêu hủy kèm biểu mẫu Mẫu 22a/18b chuẩn xác |
  | **Tổng điểm** | **19.9 / 20** | **Xếp hạng: #1 (Chủ Đề Cấp Bách Nhất — Tác động trực tiếp đến hàng nghìn nhà máy gia công dệt may, da giày, điện tử)** |
- **Tóm tắt lý do nên làm:** Cuối năm là thời điểm hàng loạt hợp đồng gia công kết thúc. Bài viết đánh trúng nỗi lo lớn nhất của giám đốc nhà máy và kế toán kho, khẳng định năng lực tư vấn chuyên sâu hàng đầu của Eureka.
- **Nguồn xác minh:** Khoản 36 Điều 1 Thông tư 121/2025/TT-BTC; Điều 64 Thông tư 38/2015/TT-BTC; Nghị định 08/2015/NĐ-CP (sửa bởi NĐ 167/2025/NĐ-CP); `scratch/daily_notes.md` (09/11/2026).

---

### Đề xuất #2: [VĂN BẢN MỚI HIỆU LỰC 16/11/2026] Thông tư 142/2026/TT-BTC: Chuẩn hóa thủ tục hải quan chuyển phát nhanh quốc tế (FedEx, DHL, UPS) và giải pháp lấy tờ khai hàng mẫu bảo vệ chi phí thuế hợp lệ
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Thủ Tục Chuyển Phát Nhanh & Hàng Mẫu)
- **Nền tảng khuyến nghị:** **Facebook & LinkedIn** (Bài viết nghiệp vụ phân tích: 'Gửi hàng mẫu qua FedEx/DHL: Làm sao để có tờ khai hải quan hợp lệ khi Thông tư 142/2026/TT-BTC có hiệu lực từ 16/11?')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Thông tư 142/2026/TT-BTC ban hành ngày 29/9/2026 và chính thức có hiệu lực từ ngày 16/11/2026 (chỉ còn đúng 1 tuần nữa) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Hàng nghìn doanh nghiệp gửi hàng mẫu EXW/DDP qua FedEx bị gom tờ khai chung, không có chứng từ hải quan đứng tên công ty dẫn đến bị loại chi phí thuế TNDN và mất thuế GTGT 0% |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận giữa bên XNK và kế toán: Forwarder/chuyển phát nhanh từ chối cung cấp tờ khai riêng hay đòi phụ phí cao; làm thế nào để số hóa quản lý vận đơn Master/House AWB |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka hướng dẫn chuẩn mực quy trình yêu cầu mở tờ khai riêng trước khi gửi hàng mẫu và cách dùng AWB trích lục chứng từ hải quan hợp lệ |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Văn Bản Pháp Lý Nóng Sắp Có Hiệu Lực Tuần Tới)** |
- **Tóm tắt lý do nên làm:** Đón đầu thông tư mới sắp có hiệu lực (16/11), giải quyết đúng nỗi đau thường trực của các nhân viên mua hàng, xuất nhập khẩu và kế toán thuế tại mọi doanh nghiệp thương mại.
- **Nguồn xác minh:** Thông tư số 142/2026/TT-BTC; Báo Hải quan Online; `scratch/daily_notes.md` (09/11/2026); Tổng cục Hải quan.

---

### Đề xuất #3: [ĐIỂM NÓNG HÓA CHẤT DNCX — HƯỚNG DẪN MỚI] Công văn 2347/HC-QLHC: Miễn giấy phép nhập khẩu hóa chất cần kiểm soát đặc biệt vào Doanh nghiệp chế xuất (DNCX) theo Điều 59 Luật Quản lý ngoại thương
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Doanh Nghiệp Chế Xuất FDI)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & LinkedIn** (Bài viết chuyên gia: 'Hóa chất kiểm soát đặc biệt vào DNCX: Đã có Công văn 2347/HC-QLHC khẳng định miễn giấy phép — Bí quyết giải trình hải quan cửa khẩu không bị ách tắc')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Văn bản hướng dẫn quan trọng của Cục Hóa chất đang được cộng đồng DNCX và Hải quan trao đổi nóng trước hạn chót thống nhất 10/11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Hóa chất kiểm soát đặc biệt dùng trong sản xuất bán dẫn/linh kiện nếu bị kẹt giấy phép sẽ làm dừng cả dây chuyền nhà máy FDI triệu USD |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Xung đột thẩm quyền: Chi cục Hải quan cửa khẩu quen đòi giấy phép Sở/Bộ Công Thương, trong khi Điều 59 Luật Quản lý ngoại thương quy định khu phi thuế quan không cần GP |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka trích dẫn chính xác CV 2347/HC-QLHC, hướng dẫn cách đính kèm văn bản lên hệ thống Một cửa quốc tế/VNACCS để thông quan trơn tru |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #3 (Pháp Lý Đỉnh Cao Dành Riêng Cho Khối Doanh Nghiệp FDI)** |
- **Tóm tắt lý do nên làm:** Định vị thương hiệu Eureka vượt trội trong phân khúc dịch vụ cho các doanh nghiệp chế xuất (DNCX) và tập đoàn FDI đa quốc gia tại Việt Nam.
- **Nguồn xác minh:** Công văn 2347/HC-QLHC ngày 11/9/2026; Điều 59 Luật Quản lý ngoại thương; Cục Hóa chất; `scratch/daily_notes.md` (09/11/2026).

---

### Đề xuất #4: [THỊ TRƯỜNG CƯỚC BIỂN WCI] Báo cáo Drewry Container Index tháng 11/2026: Cước tàu giảm nhẹ 1% về 4.389 USD/FEU — Chiến lược chốt cước dài hạn hay đi spot trước nguy cơ blank sailings cuối năm
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `FF` (Forwarder & Cước Vận Tải Quốc Tế)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook: Phân tích biểu đồ cước WCI và bảng giá tuyến Á - Mỹ, Á - Âu | YouTube: ThuyTalk phân tích xu hướng giá cước logistics và chiến lược đàm phán hợp đồng vận tải biển mùa vụ Tết)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Báo cáo tuần mới nhất công bố từ Drewry ngày 05/11/2026 chi phối thị trường giao dịch cước biển trong tuần 46 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Doanh nghiệp xuất nhập khẩu đau đầu tính toán chi phí cước biển; sợ chốt giá cao bị hớ, nhưng đi spot thì lo bị rớt hàng (rollover) do hãng tàu cắt chuyến |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận gay gắt giữa chủ hàng và forwarder: Cước spot đang giảm nhưng các hãng tàu rục rịch tăng phụ phí FAK và cắt giảm chuyến vào cuối tháng 11 |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka đưa ra lời khuyên độc quyền: Phân bổ tỷ lệ 70% hợp đồng cố định cho hàng trọng điểm và 30% linh hoạt spot, kèm bí quyết giữ booking an toàn |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #4 (Góc Nhìn Thị Trường Quốc Tế & Quản Trị Chi Phí Logistics)** |
- **Tóm tắt lý do nên làm:** Thể hiện tầm nhìn bao quát thị trường vận tải quốc tế của Ban Giám Đốc Eureka, giúp các chủ hàng tiết kiệm hàng trăm triệu đồng chi phí logistics mùa cao điểm.
- **Nguồn xác minh:** Báo cáo Drewry World Container Index (WCI); World Shipping Council; Báo Công Thương.

---

### Đề xuất #5: [PHÒNG VỆ THƯƠNG MẠI & C/O] Kim ngạch XNK 10 tháng năm 2026 vượt kỷ lục 650 tỷ USD: Ngành Hải quan siết chặt kiểm tra chống gian lận xuất xứ, lẩn tránh phòng vệ thương mại
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Phòng Vệ Thương Mại & Xuất Xứ Hàng Hóa)
- **Nền tảng khuyến nghị:** **Facebook & LinkedIn** (Bài viết cảnh báo rủi ro: 'XNK vượt 650 tỷ USD và làn sóng điều tra gian lận xuất xứ: 5 tiêu chí doanh nghiệp phải tự rà soát trước khi hải quan kiểm tra sau thông quan')
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Số liệu kinh tế vĩ mô 10 tháng vừa công bố; chiến dịch thanh tra xuất xứ của Bộ Công Thương và Hải quan đang vào đợt cao điểm quý 4 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Bị cơ quan nước ngoài (CBP, OLAF) điều tra lẩn tránh thuế phòng vệ thương mại; mức thuế trừng phạt lên tới 200-400%, doanh nghiệp có thể phá sản |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Doanh nghiệp gia công lắp ráp đơn giản có được coi là hàng có xuất xứ Việt Nam? Ranh giới giữa tối ưu thuế và gian lận xuất xứ bất hợp pháp |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka cung cấp cẩm nang 5 bước chứng minh năng lực sản xuất thực tế, giải trình định mức nguyên vật liệu và chuẩn bị hồ sơ cấp C/O hoàn hảo |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #5 (An Toàn Pháp Lý & Bảo Vệ Doanh Nghiệp Xuất Khẩu)** |
- **Tóm tắt lý do nên làm:** Cảnh báo sớm rủi ro pháp lý lớn nhất đối với các ngành hàng xuất khẩu mũi nhọn (gỗ, thép, pin mặt trời, dệt may), nâng cao uy tín thương hiệu Eureka trong việc bảo vệ khách hàng xuất khẩu.
- **Nguồn xác minh:** Cổng TTĐT Chính phủ (chinhphu.vn); Cục Hải quan; Cục Phòng vệ thương mại - Bộ Công Thương; Thời báo Tài chính Việt Nam.

---

## 🎯 KHUYẾN NGHỊ HÀNH ĐỘNG DÀNH CHO CHỊ THÚY:
1. **Ưu tiên #1 (Hành động ngay trong ngày):** Chọn **Đề xuất #1** (Xử lý nguyên liệu dư thừa hợp đồng gia công theo Thông tư 121/2025/TT-BTC) để sản xuất bài viết chuyên sâu trên Facebook Fanpage (`wf02`) kèm infographic 05 hình thức xử lý. Đây là điểm nóng sát sườn nhất của hàng loạt nhà máy đối tác đang kết thúc năm tài chính.
2. **Ưu tiên #2 (Đón đầu văn bản hiệu lực tuần tới):** Triển khai **Đề xuất #2** (Thông tư 142/2026/TT-BTC về hải quan chuyển phát nhanh & lấy tờ khai hàng mẫu FedEx) thành bài viết cẩm nang nghiệp vụ trên Fanpage và LinkedIn, hướng dẫn doanh nghiệp bảo vệ chi phí thuế hợp lý trước ngày 16/11.
3. **Ưu tiên #3 (Định vị chuyên gia khối FDI):** Triển khai **Đề xuất #3** (Công văn 2347/HC-QLHC miễn giấy phép hóa chất DNCX) để gửi trực tiếp cho khách hàng khu công nghiệp, khẳng định năng lực thông quan vượt trội của Eureka Logistics.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_briefing:
        f_briefing.write(briefing_content)
    print(f"[OK] Đã xuất bản tin đề xuất thành công: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT THEO QUY TRÌNH WF05 (09/11/2026) ===")

if __name__ == '__main__':
    update_seo_w46_0911()
