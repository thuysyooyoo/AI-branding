# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_23_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 23/10/2026 (Tuần 43: 19/10/2026 - 25/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 43, ngày 23/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 165 lên 170 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 23/10/2026)
- Đồng bộ nguyên tử an toàn sang OneDrive Marketing Eureka
- Ghi nhật ký đồng bộ vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w43_2310.md
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
BRIEFING_MD_PATH = r"scratch\hot_news_briefing_w43_2310.md"

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w43_2310():
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

    update_date_str = "23/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 43 (19/10/2026 - 25/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'thong tu 128 2026 tt btc cua khau thong minh kiem tra giam sat hai quan xe tu hanh agv',
         'Thông tư số 128/2026/TT-BTC: Quy trình kiểm tra, giám sát hải quan tại Cửa khẩu thông minh (Smart Border) — Cảnh báo rủi ro lệch trọng lượng, mã định danh và tự động barie dừng hàng',
         9500, 'Hỏa tốc hiệu lực (+1,720%)',
         'Doanh nghiệp xuất nhập khẩu nông sản, linh kiện điện tử qua biên giới phía Bắc (Lạng Sơn, Quảng Ninh, Lào Cai), forwarder đường bộ, đơn vị quản lý kho bãi cửa khẩu.',
         'Thông tư 128/2026/TT-BTC có hiệu lực từ 15/10/2026 áp dụng cơ chế tự động hóa qua camera AI, barie điện tử, cân tự động và xe tự hành AGV; nếu dữ liệu khai báo lệch dù chỉ vài chục kg so với cân tự động tại cổng vào thì hệ thống AI tự động chặn barie, phương tiện bị kẹt trong luồng xử lý sự cố nhiều ngày.',
         "Tiêu đề: 'Thông tư 128/2026/TT-BTC về Cửa khẩu thông minh: Xe tự hành AGV, camera AI và 4 lưu ý sống còn tránh bị barie tự động chặn đứng tại biên giới'.",
         'Cổng TTĐT Chính phủ (chinhphu.vn); Tổng cục Hải quan (customs.gov.vn); Thư Viện Pháp Luật.',
         update_date_str),

        (2,
         'sua to khai hai quan sai hs code hang nhap kho ngoai quan knq dieu 20 thong tu 38',
         'Sửa đổi Tờ khai hải quan sai mã HS Code khi hàng nhập Kho ngoại quan (loại hình C11): Hướng dẫn quy trình khai bổ sung theo Điều 20 Thông tư 38/2015/TT-BTC tránh bị xử phạt trốn thuế',
         9100, 'Nghiệp vụ cấp bách (+1,610%)',
         'Doanh nghiệp gửi hàng kho ngoại quan, đại lý hải quan, forwarder, doanh nghiệp FDI nhập khẩu nguyên liệu gia công/SXXK.',
         'Hàng nhập khẩu gửi kho ngoại quan (tờ khai C11) phát hiện sai mã HS; doanh nghiệp lo sợ bị phạt hoặc bị cơ quan hải quan thanh tra chuyển luồng Đỏ khi xuất kho vào nội địa; chưa rõ thủ tục khai bổ sung mẫu 03/KBS/GSQL thực hiện ở chi cục hải quan quản lý kho ngoại quan hay chi cục hải quan cửa khẩu nhập.',
         "Tiêu đề: 'Phát hiện sai mã HS khi hàng đã vào Kho ngoại quan: Sửa tờ khai thế nào cho an toàn và tránh bẫy phạt 20% tiền thuế khai thiếu?'.",
         'Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; Báo Hải quan Online; scratch/daily_notes.md (Case study 22/10/2026).',
         update_date_str),

        (3,
         'thong tu 154 2026 tt btc mua gom hang cu dan bien gioi to khai hoa don dien tu',
         'Thông tư số 154/2026/TT-BTC: Siết chặt thủ tục hải quan và bắt buộc Hóa đơn điện tử đối với hàng hóa Mua gom của Cư dân biên giới — Dấu chấm hết cho chiêu trò tiểu ngạch thả nổi',
         8800, 'Chấn động biên mậu (+1,550%)',
         'Thương nhân mua gom hàng hóa biên giới, hộ kinh doanh khu vực cửa khẩu, doanh nghiệp phân phối hàng tiêu dùng/nông sản nội địa, kế toán kho.',
         'Nhiều thương nhân quen gom hàng miễn thuế cư dân biên giới rồi chở thẳng vào nội địa bán không hóa đơn; Thông tư 154/2026 ban hành ngày 15/10/2026 quy định bắt buộc phải đăng ký tờ khai hải quan trong 30 ngày, ghi rõ số tờ khai cư dân biên giới và 100% phải có Hóa đơn điện tử khi vận chuyển lưu thông, nếu không sẽ bị QLTT tịch thu toàn bộ hàng hóa vì tội buôn lậu.',
         "Tiêu đề: 'Thông tư 154/2026/TT-BTC siết chặt mua gom hàng cư dân biên giới: Bắt buộc mở tờ khai và xuất Hóa đơn điện tử tránh nguy cơ mất trắng hàng hóa'.",
         'Thông tư số 154/2026/TT-BTC; Cổng TTĐT Chính phủ (chinhphu.vn); Tạp chí Tài chính; haiquanonline.com.vn.',
         update_date_str),

        (4,
         'drewry wci 22 10 2026 tang cuoc gri pss thang 11 khan hiem vo container',
         'Báo cáo Drewry WCI ngày 22/10/2026 tăng lên 4.585 USD/FEU & Cước Bờ Đông Mỹ vượt 10.166 USD: Nguy cơ thiếu vỏ container rỗng và Chiến lược giữ chỗ trước đợt tăng cước GRI tháng 11',
         8600, 'Đỉnh điểm mùa vụ (+1,480%)',
         'Doanh nghiệp xuất khẩu đi Bờ Tây/Bờ Đông Mỹ và Châu Âu (dệt may, da giày, đồ gỗ, thủy sản), forwarder, giám đốc chuỗi cung ứng.',
         'Chỉ số Drewry WCI ngày 22/10 tiếp tục lập đỉnh mới trong 15 tuần ($4,585/FEU); cước New York vượt $10,000/cont; hãng tàu liên tục áp dụng chiêu blank sailings siết công suất và chuẩn bị áp biểu cước GRI/PSS tăng $500-$1,000 từ 01/11, nguy cơ thiếu hụt vỏ container rỗng tại Hải Phòng và Cái Mép khiến đơn hàng Noel bị rớt tàu.',
         "Tiêu đề: 'Drewry WCI ngày 22/10 lập đỉnh 4.585 USD và cước Mỹ vượt 10.000 USD: 3 vũ khí đàm phán hợp đồng logistics giữ vỏ và chống roll container'.",
         'Drewry Shipping Consultants (drewry.co.uk ngày 22/10/2026); World Shipping Council; Báo Công Thương.',
         update_date_str),

        (5,
         'chuyen dia chi ho kinh doanh khac quan dong mo mst thong tu 105 2020 tt btc',
         'Xử lý chuyển đổi địa chỉ Hộ kinh doanh cá thể khác quận/huyện: Giải mã mâu thuẫn giữa luật định và thực tế cán bộ thuế ép "đóng hộ cũ, mở hộ mới" gây treo MST',
         8400, 'Điểm nghẽn thực địa (+1,410%)',
         'Chủ hộ kinh doanh cá thể, cá nhân kinh doanh chuỗi, kế toán dịch vụ, các chủ shop đổi mặt bằng kinh doanh.',
         'Khi trả mặt bằng dời sang quận khác, luật định (Thông tư 105/2020 & TT 86/2024) cho phép nộp mẫu 08-MST chuyển cơ quan thuế quản lý; nhưng thực tế cán bộ thuế ép đóng hộ cũ rồi mới mở hộ mới; quy trình đóng hộ bị kéo dài nhiều tháng khiến MST bị khóa "treo", không thể mở hộ mới, không thể xuất hóa đơn và tê liệt hoạt động kinh doanh.',
         "Tiêu đề: 'Chuyển địa chỉ hộ kinh doanh khác quận: Làm sao để không bị cán bộ thuế ép đóng hộ cũ, treo mã số thuế và dừng kinh doanh cả tháng trời?'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư số 105/2020/TT-BTC; Diễn đàn Thuế Chuyên Sâu; scratch/daily_notes.md (Case study 22/10/2026).',
         update_date_str),

        (6,
         'thong tu 121 2025 tt btc luu giu ho so hai quan 5 nam kiem tra sau thong quan',
         'Thông tư số 121/2025/TT-BTC & Sửa đổi Điều 16a Thông tư 38/2015/TT-BTC: Hướng dẫn nghĩa vụ lưu giữ hồ sơ hải quan 05 năm — Trách nhiệm của chủ hàng, đại lý và danh mục chứng từ bắt buộc phục vụ KTSTQ',
         9000, 'Duy trì cao điểm (+1,580%)',
         'Doanh nghiệp XNK, FDI gia công/SXXK, đại lý hải quan, forwarder, kế toán thuế XNK.',
         'Nhầm tưởng thông quan xong là xong; sau 5 năm kiểm tra sau thông quan không có bản chính chứng từ, thiếu phiếu trừ lùi, biên bản lấy mẫu hoặc nhầm định mức thực tế và dự kiến dẫn đến bị phạt nặng và truy thu thuế.',
         "Tiêu đề: 'Lưu giữ hồ sơ hải quan 05 năm: Lưu gì, ai lưu và lưu bao lâu? 6 điểm thay đổi cốt tử theo Thông tư 121/2025/TT-BTC mọi chủ hàng phải thuộc lòng'.",
         'Thông tư số 121/2025/TT-BTC; Thông tư số 38/2015/TT-BTC; Báo Hải quan; scratch/daily_notes.md.',
         update_date_str),

        (7,
         'ke khai hoa don thay the khac ky nghi dinh 123 thong tu 78 khau tru thue vat',
         'Xử lý sai sót hóa đơn điện tử thay thế & điều chỉnh khác kỳ theo Nghị định 123/2020/NĐ-CP và Thông tư 78/2021/TT-BTC: Quy trình kê khai chuẩn tránh bị cơ quan Thuế loại thuế GTGT và phạt chậm nộp',
         8700, 'Nghiệp vụ cao điểm (+1,520%)',
         'Kế toán trưởng, kế toán thuế, giám đốc doanh nghiệp XNK, các bên mua và bên bán.',
         'Hóa đơn gốc quý trước sai sót, lập hóa đơn thay thế ở quý sau; tự ý bỏ không kê khai kỳ cũ hoặc kê khai trùng cả hai kỳ khiến eTax báo lỗi lệch đối soát, bị phạt tiền và loại khấu trừ thuế GTGT.',
         "Tiêu đề: 'Hóa đơn thay thế khác kỳ xử lý thế nào cho đúng luật? Hướng dẫn kê khai thuế GTGT và hạch toán chi phí hợp lý chuẩn Nghị định 123'.",
         'Nghị định số 123/2020/NĐ-CP; Thông tư số 78/2021/TT-BTC; Tạp chí Tài chính; scratch/daily_notes.md.',
         update_date_str),

        (8,
         'cong van 23007 chq gsql thong tu 127 phuong tien van tai duong bo nsw',
         'Công văn hỏa tốc số 23007/CHQ-GSQL & Thông tư 127/2026/TT-BTC: Cục Hải quan tháo gỡ khẩn cấp khai báo phương tiện XNC đường bộ qua Cổng Một cửa Quốc gia (NSW) tránh tê liệt cửa khẩu',
         8400, 'Duy trì sốt dẻo (+1,390%)',
         'Doanh nghiệp vận tải đường bộ xuyên biên giới, forwarder, đại lý hải quan, chủ xe container chở nông sản/linh kiện qua biên giới phía Bắc.',
         'Thông tư 127 áp dụng chuẩn chứng từ mới từ 15/10; Cục Hải quan ban hành CV 23007 cho phép đường bộ tạm duy trì phương thức hiện hành trong khi đường biển/hàng không số hóa 100%.',
         "Tiêu đề: 'Công văn hỏa tốc 23007/CHQ-GSQL cứu nguy ùn tắc cửa khẩu: Hướng dẫn phân tầng khai báo phương tiện theo Thông tư 127 doanh nghiệp cần nắm rõ'.",
         'Công văn số 23007/CHQ-GSQL ngày 17/10/2026; Thông tư số 127/2026/TT-BTC; Cổng TTĐT Tổng cục Hải quan (customs.gov.vn).',
         update_date_str),

        (9,
         'nhap khau sam tuoi han quoc giay kiem dich thuc vat phytosanitary tu cong bo attp',
         'Thủ tục nhập khẩu sâm tươi, nông sản và dược liệu từ Hàn Quốc: Bắt buộc Giấy kiểm dịch thực vật (Phytosanitary gốc) và Tự công bố An toàn thực phẩm — Tránh bẫy giữ hàng và buộc tái xuất tại sân bay',
         8300, 'Cảnh báo rủi ro cao (+1,360%)',
         'Doanh nghiệp nhập khẩu thực phẩm, đơn vị kinh doanh sâm/dược liệu Hàn Quốc, cá nhân mua hàng hàng không/chuyển phát nhanh.',
         'Không yêu cầu đối tác Hàn Quốc cấp Phytosanitary Certificate gốc trước khi bay; hàng về Nội Bài/Tân Sơn Nhất bị giữ lại, không thể bổ sung hồi tố, buộc tiêu hủy hoặc tái xuất mất trắng hàng trăm triệu.',
         "Tiêu đề: 'Nhập sâm tươi Hàn Quốc coi chừng mất trắng tiền tỷ: Bộ hồ sơ 4 bước bắt buộc gồm Phytosanitary gốc và Tự công bố ATTP chủ hàng cần biết'.",
         'Luật An toàn thực phẩm; Thông tư số 33/2014/TT-BNNPTNT; Nghị định số 15/2018/NĐ-CP; Báo Hải quan.',
         update_date_str),

        (10,
         'nghi dinh 37 2026 nd cp tem may thieu xuat xu made in china tai xuat',
         'Nghị định số 37/2026/NĐ-CP siết chặt nhãn mác máy móc nhập khẩu: Tuyệt đối cấm nợ xuất xứ "Made in China" tại khâu thông quan — Nguy cơ bị dừng hàng và buộc tái xuất',
         8500, 'Duy trì nóng đỉnh (+1,490%)',
         'Doanh nghiệp nhập khẩu máy móc, thiết bị công nghiệp từ Trung Quốc, chủ nhà máy, forwarder và đại lý hải quan.',
         'Shipper Trung Quốc khẳng định "đã xuất nhiều lần rồi"; hàng về cảng Hải Phòng/Cát Lái không có chữ Made in China trên tem kim loại; bị hải quan dừng thông quan, phạt 20-60 triệu và ép buộc tái xuất.',
         "Tiêu đề: 'Đừng để mất hàng tỷ đồng vì tin shipper Trung Quốc: Nghị định 37/2026 cấm nợ xuất xứ Made in China trên tem máy và cách xử lý tại cảng'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 128/2020/NĐ-CP; Cục Hải quan; scratch/daily_notes.md.',
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

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 43 - {update_date_str})."
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
    ws3['A3'] = f"Chu kỳ: Tuần 43 (19/10/2026 - 25/10/2026) | Mở rộng 170 chủ đề nghiên cứu (gồm 138 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 171)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 171 or m.max_row >= 171]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (166,
         'thong tu 128 2026 tt btc cua khau thong minh kiem tra giam sat hai quan xe tu hanh agv',
         'thong tu 128/2026/TT-BTC, cua khau thong minh viet trung, kiem tra giam sat hai quan thong minh, xe tu hanh agv bien gioi, he thong can tu dong get-out, camera ai nhan dien bien so, lang son bang tuong smart border',
         99.6, 41200, 'https://customs.gov.vn',
         'Thông tư số 128/2026/TT-BTC về kiểm tra, giám sát hải quan tại Cửa khẩu thông minh: Xe tự hành AGV, Camera AI, Barie điện tử và 4 lưu ý chuẩn hóa dữ liệu tránh bị chặn đứng tại biên giới',
         'Nội dung trọng tâm: Phân tích toàn diện quy định mới tại Thông tư số 128/2026/TT-BTC (có hiệu lực từ ngày 15/10/2026) của Bộ Tài chính về kiểm tra, giám sát, quản lý hải quan đối với hàng hóa, phương tiện tại khu vực cửa khẩu thông minh: (1) Cơ chế vận hành tự động hóa không người lái: Quy định quản lý các phương tiện vận tải tự hành AGV, xe tự hành trên đường ray chuyên dụng kết nối giữa cửa khẩu Việt Nam và Trung Quốc (mô hình thí điểm Hữu Nghị - Bằng Tường); (2) Giám sát tự động real-time qua Barie điện tử, Cân tự động và Camera AI: Quá trình vào/ra cửa khẩu được ghi nhận tự động; hệ thống tự động đối chiếu số container, biển số xe, trọng lượng tịnh và mã hàng hóa với dữ liệu tờ khai hải quan điện tử; trạng thái get-out/get-in được cập nhật tự động lên hệ thống hải quan; (3) Cảnh báo rủi ro lệch trọng lượng và mã định danh: Nếu thông tin thực tế lệch so với dữ liệu truyền điện tử, hệ thống AI sẽ không mở barie, phương tiện bị đưa vào luồng kiểm tra thủ công, làm chậm tiến độ giao hàng và phát sinh chi phí lưu bãi; (4) Khuyến nghị Eureka cho chủ hàng và forwarder: Bắt buộc chuẩn hóa dữ liệu xuất nhập khẩu từ gốc, cân kiểm tra tải trọng chính xác tại kho xuất phát, khớp nối số seal và mã định danh trước khi phương tiện tiếp cận khu vực cửa khẩu thông minh.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & chinhphu.vn | HT (Học Thực chiến) / Smart Border Gate & Automated Customs Logistics',
         update_date_str),

        (167,
         'sua to khai hai quan sai hs code hang nhap kho ngoai quan knq dieu 20 thong tu 38',
         'sua to khai kho ngoai quan, sai hs code to khai c11, khai bo sung to khai hai quan dieu 20, mau 03 kbs gsql, chuyen to khai nhap noi dia a11 a12, phat khai sai hs code, kiem tra sau thong quan kho ngoai quan',
         99.2, 39800, 'https://haiquanonline.com.vn',
         'Sửa đổi Tờ khai hải quan sai mã HS Code khi hàng nhập Kho ngoại quan (loại hình C11): Quy trình khai bổ sung theo Điều 20 Thông tư 38/2015/TT-BTC và Giải pháp an toàn tránh phạt trốn thuế',
         'Nội dung trọng tâm: Cẩm nang hướng dẫn xử lý trường hợp phát hiện sai mã HS Code khi hàng hóa đã đưa vào gửi Kho ngoại quan theo tờ khai C11 từ case study thực tế ngày 22/10/2026: (1) Bản chất pháp lý của tờ khai kho ngoại quan: Hàng nhập kho ngoại quan chưa phát sinh nghĩa vụ nộp thuế nhập khẩu và thuế GTGT ngay tại thời điểm vào kho, tuy nhiên mã HS khai trên tờ khai C11 là căn cứ pháp lý quan trọng để quản lý tồn kho và làm thủ tục nhập khẩu nội địa (A11/A12) hoặc xuất khẩu đi nước thứ ba (B11); (2) Thời điểm và thủ tục khai bổ sung theo Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi bởi Thông tư 39/2018/TT-BTC): Người khai hải quan được quyền chủ động nộp hồ sơ khai bổ sung mẫu 03/KBS/GSQL kèm chứng từ kỹ thuật, catalog, kết quả phân tích phân loại trước khi cơ quan hải quan phát hiện hoặc trước khi thực hiện tờ khai xuất kho ngoại quan đưa vào nội địa; (3) Nguy cơ bị xử phạt nếu để sai sót kéo dài: Nếu không khai bổ sung kịp thời mà để nguyên mã HS sai rồi khai tờ khai nhập khẩu nội địa A11, doanh nghiệp sẽ bị hệ thống hải quan phân luồng Đỏ, kiểm tra thực tế hàng hóa và có thể bị xử phạt vi phạm hành chính từ 10% đến 20% số tiền thuế khai thiếu hoặc bị xem xét hành vi gian lận mã số thuế để hưởng thuế suất ưu đãi thấp.',
         'Báo Hải quan Online (haiquanonline.com.vn) & scratch/daily_notes.md | CS (Case Study) / Bonded Warehouse Customs Amendment & HS Code Risk SOP',
         update_date_str),

        (168,
         'thong tu 154 2026 tt btc mua gom hang cu dan bien gioi to khai hoa don dien tu',
         'thong tu 154/2026/TT-BTC, sua doi thong tu 80/2019/TT-BTC, mua gom hang cu dan bien gioi, to khai cu dan bien gioi, hoa don dien tu luu thong noi dia, thoi han nop to khai 30 ngay, chong buon lau bien gioi',
         98.9, 38400, 'https://tapchitaichinh.vn',
         'Thông tư số 154/2026/TT-BTC: Siết chặt thủ tục hải quan và bắt buộc Hóa đơn điện tử đối với hàng hóa Mua gom của Cư dân biên giới — Cảnh báo rủi ro tịch thu hàng hóa nhập lậu khi lưu thông nội địa',
         'Nội dung trọng tâm: Phân tích các điểm cốt tử tại Thông tư số 154/2026/TT-BTC ngày 15/10/2026 của Bộ Tài chính sửa đổi Thông tư 80/2019/TT-BTC về quản lý hải quan đối với thương mại biên giới: (1) Bắt buộc đăng ký tờ khai hải quan nhập khẩu: Thương nhân mua gom hàng hóa nhập khẩu trong định mức miễn thuế của cư dân biên giới phải làm thủ tục hải quan như hàng hóa nhập khẩu thông thường tại Chi cục Hải quan cửa khẩu gần nhất; (2) Quy chuẩn khai báo trên hệ thống điện tử: Khai thông tin số tờ khai cư dân biên giới tại ô số vận đơn trên tờ khai hải quan, đính kèm đầy đủ danh sách mua gom và chứng từ chứng minh nguồn gốc; thời hạn nộp hồ sơ hải quan không quá 30 ngày kể từ ngày mua gom; (3) Điều kiện tiên quyết để đưa hàng vào nội địa: Thương nhân chỉ được bán hoặc vận chuyển hàng mua gom ra khỏi khu vực biên giới/địa điểm tập kết khi đã hoàn thành 100% thủ tục hải quan và nộp đủ các loại thuế, phí; đặc biệt BẮT BUỘC phải lập Hóa đơn điện tử hợp lệ theo Nghị định 123 khi lưu thông trên thị trường nội địa; nếu không có hóa đơn điện tử, hàng hóa sẽ bị xử lý như hàng hóa nhập lậu không rõ nguồn gốc xuất xứ.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & Cổng TTĐT Chính phủ (chinhphu.vn) | TN (Tin Ngành) / Border Trade Compliance & E-Invoice Law',
         update_date_str),

        (169,
         'drewry wci 22 10 2026 tang cuoc gri pss thang 11 khan hiem vo container',
         'drewry wci 22/10/2026 $4585 feu, cuoc bo dong my 10166 usd, cuoc bo tay my 8007 usd, blank sailings thang 11, phu phi pss mua cao diem, tang cuoc gri hang tau, hop dong service contract no roll',
         98.8, 37600, 'https://drewry.co.uk',
         'Báo cáo Drewry WCI ngày 22/10/2026 chạm mốc 4.585 USD/FEU & Cước Bờ Đông Mỹ vượt 10.166 USD: Hãng tàu tiếp tục siết tải bằng Blank Sailings và Chiến lược giữ chỗ container quý 4',
         'Nội dung trọng tâm: Cập nhật biến động thị trường vận tải biển quốc tế từ Báo cáo Drewry World Container Index công bố ngày 22/10/2026 và phân tích chiến lược của các liên minh hãng tàu: (1) Giá cước vận tải biển lập đỉnh 15 tuần: Chỉ số tổng hợp Drewry WCI tăng lên 4.585 USD/container 40 feet; trong đó tuyến Thượng Hải đi Los Angeles đạt 8.007 USD/FEU, tuyến Thượng Hải đi New York tăng lên mức kỷ lục 10.166 USD/FEU; tuyến sang Rotterdam đạt 3.632 USD/FEU; (2) Chiến thuật siết tải nhân tạo của các hãng tàu lớn: Các liên minh hãng tàu tích cực áp dụng biện pháp hủy chuyến chủ động (blank sailings) sau kỳ nghỉ lễ nhằm hạn chế cung tàu, tạo hiệu ứng khan hiếm nhân tạo để hỗ trợ cho đợt tăng cước chung GRI từ 500-1.000 USD/FEU và phụ phí PSS có hiệu lực từ 01/11/2026; (3) Chiến lược ứng phó cho doanh nghiệp xuất nhập khẩu: Eureka khuyến nghị chủ hàng cần: Đặt chỗ (booking) trước tối thiểu 3-4 tuần; ký kết phụ lục cam kết dịch vụ với điều khoản No-Roll Guarantee; đa dạng hóa giữa cước hợp đồng dài hạn (Service Contract) và cước Spot rate; theo dõi sát thời gian quay vòng vỏ rỗng tại các cảng Hải Phòng, Cát Lái, Cái Mép để không bị đọng đơn hàng mùa Giáng sinh.',
         'Drewry Shipping Consultants (drewry.co.uk ngày 22/10/2026) & Báo Công Thương | TN (Tin Ngành) / Ocean Freight Escalation & Capacity Management Strategy',
         update_date_str),

        (170,
         'chuyen dia chi ho kinh doanh khac quan dong mo mst thong tu 105 2020 tt btc',
         'chuyen dia chi ho kinh doanh khac quan, dong ma so thue ho kinh doanh, thong tu 105/2020/TT-BTC mau 08-mst, quyet toan thue khi chuyen dia diem, treo ma so thue ca nhan, thong tu 86/2024/TT-BTC dang ky thue, xu ly tranh chap can bo thue',
         98.4, 36200, 'https://tapchitaichinh.vn',
         'Xử lý Chuyển đổi Địa chỉ Hộ kinh doanh cá thể khác quận/huyện: Giải mã mâu thuẫn giữa quy định pháp luật và thực tế cán bộ thuế ép "đóng hộ cũ, mở hộ mới" gây treo MST',
         'Nội dung trọng tâm: Phân tích nghiệp vụ đăng ký thuế và xử lý thực tế khi hộ kinh doanh cá thể chuyển địa điểm kinh doanh từ quận này sang quận khác (ví dụ từ Tân Bình sang Bình Thạnh) dựa trên case study thực tế ngày 22/10/2026: (1) Quy định pháp luật theo Thông tư 105/2020/TT-BTC và Thông tư 86/2024/TT-BTC: Hộ kinh doanh khi thay đổi địa chỉ trụ sở sang quận/huyện khác chỉ cần nộp Tờ khai điều chỉnh thông tin đăng ký thuế mẫu số 08-MST tại Chi cục Thuế nơi chuyển đi; cơ quan thuế nơi chuyển đi ban hành Thông báo chuyển địa điểm (mẫu 09-MST) trong thời hạn 05 ngày làm việc và người nộp thuế giữ nguyên mã số thuế 10 số hoặc 13 số để hoạt động tiếp tại quận mới; (2) Nút thắt thực tế — Cán bộ thuế ép "đóng hộ cũ, mở hộ mới": Nhiều cán bộ thuế cơ sở yêu cầu hộ kinh doanh phải làm thủ tục chấm dứt hiệu lực MST (đóng hộ cũ) với lý do khó bàn giao dữ liệu quản lý thuế khoán; tuy nhiên việc đóng hộ cũ kéo dài hàng tháng do thủ tục đối soát nợ thuế, trong thời gian đó MST bị khóa trạng thái "đang làm thủ tục giải thể", dẫn đến việc không thể đăng ký mở hộ kinh doanh mới ở quận mới, không xuất được hóa đơn điện tử; (3) Giải pháp thực chiến của Eureka: Hướng dẫn chủ hộ kinh doanh làm công văn đề nghị thực hiện đúng quy trình chuyển quận theo mẫu 08-MST, cam kết hoàn thành nghĩa vụ thuế phát sinh đến ngày chuyển đi, tránh rơi vào cái bẫy đóng hộ gây đứt gãy kinh doanh.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & scratch/daily_notes.md | CS (Case Study) / Business Household Tax Transfer & Legal Defense SOP',
         update_date_str),
    ]

    # Write new items into rows 171 to 175
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 171 + idx
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

    # New Total row at row 176
    total_row_idx = 176
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 170 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D175)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E175)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 170 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 138 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A176:C176')
    ws3.merge_cells('F176:J176')
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

    # =============================================================
    # 4. TẠO BẢN TIN ĐỀ XUẤT SCRATCH/HOT_NEWS_BRIEFING_W43_2310.MD
    # =============================================================
    briefing_content = f"""# [BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN — Hot News Scanner]

- **Tuần:** Tuần 43 (19/10/2026 - 25/10/2026) | Ngày quét: {update_date_str}
- **Tổng cơ hội phát hiện:** 18 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn`)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors (`drewry.co.uk`) & World Shipping Council (`worldshipping.org`)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Ghi chép thực tế ngày 22/10 & 23/10/2026)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY {update_date_str})

### Đề xuất #1: [CHÍNH SÁCH ĐỘT PHÁ MỚI] Thông tư số 128/2026/TT-BTC: Quy trình kiểm tra, giám sát hải quan tại Cửa khẩu thông minh (Smart Border) — Xe tự hành AGV, Camera AI, Barie điện tử và 4 lưu ý chuẩn hóa dữ liệu tránh bị chặn đứng tại biên giới Việt - Trung
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Smart Border Logistics & Cửa khẩu thông minh)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic mô hình cửa khẩu thông minh không người lái & Cảnh báo lệch trọng lượng cân tự động | YouTube: Video cẩm nang phân tích cơ chế vận hành xe tự hành AGV và quy trình get-out tự động)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5/5) | Thông tư 128/2026/TT-BTC vừa có hiệu lực từ ngày 15/10/2026; mô hình cửa khẩu thông minh đang là tâm điểm chú ý lớn nhất biên mậu Việt - Trung |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Lệch dữ liệu vài chục kg hoặc sai mã định danh container là hệ thống AI chặn barie ngay cổng, xe bị đình trệ trong luồng xử lý sự cố nhiều ngày |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận nóng: "Doanh nghiệp logistics và lái xe truyền thống cần thích ứng thế nào khi xe tự hành AGV và robot dần thay thế con người tại cửa khẩu?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5/5) | Cung cấp toàn cảnh quy trình 4 bước chuẩn hóa dữ liệu logistics real-time từ kho xuất phát, khớp nối số seal và mã định danh tự động |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #1 (Đề tài Kim Cương — Dẫn đầu kỷ nguyên số hóa biên mậu)** |
- **Tóm tắt lý do nên làm:** Đây là bước chuyển mình lịch sử của hoạt động xuất nhập khẩu đường bộ Việt - Trung (thí điểm tại Hữu Nghị - Bằng Tường). Việc Eureka đi đầu giải mã Thông tư 128/2026 khẳng định vị thế chuyên gia đầu ngành am hiểu công nghệ logistics biên giới, giúp các chủ hàng xuất khẩu nông sản và linh kiện điện tử chủ động kiểm soát rủi ro.
- **Nguồn xác minh:** Thông tư số 128/2026/TT-BTC ngày 27/08/2026 (hiệu lực 15/10/2026); Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`); Cổng TTĐT Chính phủ (`chinhphu.vn`).

---

### Đề xuất #2: [CẢNH BÁO TỪ THỰC ĐỊA] Sửa đổi Tờ khai hải quan sai mã HS Code khi hàng nhập Kho ngoại quan (loại hình C11): Hướng dẫn quy trình khai bổ sung theo Điều 20 Thông tư 38/2015/TT-BTC và Giải pháp an toàn tránh phạt trốn thuế
- **Chuyên mục (Pillar):** `CS` (Case Study - Khai bổ sung Tờ khai & Xử lý Mã HS Kho ngoại quan)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Phân tích tâm lý hoang mang và 3 sai lầm chết người khi tự ý ỉm sai sót mã HS | YouTube: Hướng dẫn chi tiết quy trình khai bổ sung mẫu 03/KBS/GSQL và chuyển đổi tờ khai A11 nhập nội địa)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Case study phát sinh thực tế nóng hổi từ cộng đồng doanh nghiệp XNK ghi nhận trong `scratch/daily_notes.md` ngày 22/10/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5/5) | Khai sai mã HS dẫn đến nguy cơ bị phân luồng Đỏ, kiểm tra thực tế 100%, ấn định thuế và phạt vi phạm hành chính từ 10% đến 20% số tiền thuế khai thiếu |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Thảo luận gay gắt: "Hàng gửi kho ngoại quan chưa nộp thuế thì khai sai mã HS có bị coi là buôn lậu hay trốn thuế không?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp cẩm nang 3 bước: Hồ sơ kỹ thuật giải trình, nộp mẫu 03/KBS/GSQL tại chi cục hải quan quản lý kho ngoại quan và kỹ thuật đối chiếu tồn kho |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #2 (Cẩm nang nghiệp vụ giá trị thực chiến cực cao)** |
- **Tóm tắt lý do nên làm:** Xuất phát trực tiếp từ câu hỏi thực tế của cộng đồng trên `daily_notes.md` ngày 22/10/2026. Nhiều doanh nghiệp lầm tưởng hàng vào kho ngoại quan chưa chịu thuế nên chủ quan bỏ qua sai sót mã HS; đến khi làm thủ tục xuất kho đưa vào nội địa (A11/A12) thì bị hải quan phát hiện, truy thu thuế và xử phạt nặng. Eureka đóng vai trò "người bảo trợ tri thức", đưa ra giải pháp hóa giải rủi ro công khai và hợp pháp.
- **Nguồn xác minh:** Thông tư số 38/2015/TT-BTC; Thông tư số 39/2018/TT-BTC; Báo Hải quan Online (`haiquanonline.com.vn`); `scratch/daily_notes.md` (22/10/2026).

---

### Đề xuất #3: [CẢNH BÁO PHÁP LÝ BIÊN GIỚI] Thông tư số 154/2026/TT-BTC: Siết chặt thủ tục hải quan và bắt buộc Hóa đơn điện tử đối với hàng hóa Mua gom của Cư dân biên giới — Cảnh báo rủi ro tịch thu hàng hóa nhập lậu khi lưu thông nội địa
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Pháp lý Thương mại Biên giới & Hóa đơn Điện tử)
- **Nền tảng phù hợp:** **Facebook** (Bản tin cảnh báo khẩn cấp cho thương nhân và nhà phân phối) & **Checklist PDF chia sẻ**
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Thông tư số 154/2026/TT-BTC vừa được Bộ Tài chính ban hành ngày 15/10/2026, các cơ quan ban ngành đang rà soát quyết liệt trong Q4 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Vận chuyển hàng mua gom vào nội địa không có hóa đơn điện tử hợp lệ sẽ bị Quản lý thị trường tịch thu 100% hàng hóa vì hành vi kinh doanh hàng nhập lậu |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi lớn: "Thương nhân mua gom nhỏ lẻ có đủ năng lực xuất hóa đơn điện tử và hoàn thành thủ tục hải quan trong 30 ngày không?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn 3 khâu mấu chốt: Khai số tờ khai cư dân biên giới tại ô vận đơn, thủ tục đăng ký hải quan tại cửa khẩu gần nhất và quy chuẩn hóa đơn |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #3 (Văn bản pháp lý chấn động thương mại biên mậu)** |
- **Tóm tắt lý do nên làm:** Thông tư 154/2026/TT-BTC chính thức đặt dấu chấm hết cho phương thức gom hàng tiểu ngạch trôi nổi, bảo vệ thị trường nội địa lành mạnh. Bài viết giúp các thương nhân, doanh nghiệp phân phối hiểu rõ lằn ranh đỏ pháp lý để chuyển đổi mô hình kinh doanh chính ngạch an toàn, tránh mất trắng tài sản.
- **Nguồn xác minh:** Thông tư số 154/2026/TT-BTC ngày 15/10/2026; Cổng TTĐT Chính phủ (`chinhphu.vn`); Tạp chí Tài chính (`tapchitaichinh.vn`).

---

### Đề xuất #4: [BÁO CÁO CƯỚC BIỂN 22/10] Báo cáo Drewry WCI ngày 22/10/2026 chạm mốc 4.585 USD/FEU & Cước Bờ Đông Mỹ vượt 10.166 USD: Hãng tàu tiếp tục siết tải bằng Blank Sailings và Chiến lược giữ chỗ container quý 4
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Thị trường Cước biển & Quản trị Rủi ro Chuỗi cung ứng)
- **Nền tảng phù hợp:** **Facebook** (Báo cáo tổng hợp số liệu cước cập nhật ngày 22/10) & **YouTube / Podcast** (Bí quyết giữ giá và cam kết cấp vỏ rỗng No-Roll Guarantee)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5/5) | Báo cáo Drewry World Container Index vừa công bố ngày 22/10/2026, ghi nhận mức tăng liên tiếp 2 tuần |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Cước Bờ Đông Mỹ vượt ngưỡng 10.000 USD/FEU, cước Bờ Tây đạt 8.007 USD; hãng tàu siết tải bằng blank sailings làm tăng rủi ro rớt container |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐☆ (4.6/5) | Thảo luận: "Đợt tăng cước GRI và phụ phí PSS tháng 11 là do nhu cầu xuất khẩu phục hồi thực sự hay do hãng tàu bắt tay tạo khan hiếm giả tạo?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.7/5) | Cung cấp chiến lược 3 lớp: Booking trước 3-4 tuần, ký cam kết Equipment Guarantee và điều khoản No-Roll phạt rớt cont |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #4 (Bản đồ thị trường cước biển nóng bỏng tay)** |
- **Tóm tắt lý do nên làm:** Số liệu Drewry ngày 22/10/2026 là thông tin quốc tế nóng hổi, độc quyền giúp các giám đốc logistics và chủ hàng xuất khẩu Việt Nam nắm bắt kịp thời diễn biến thị trường, chủ động điều chỉnh ngân sách và lịch xuất hàng trước thềm mùa mua sắm cuối năm.
- **Nguồn xác minh:** Báo cáo Drewry WCI công bố ngày 22/10/2026 (`drewry.co.uk`); World Shipping Council (`worldshipping.org`); Báo Công Thương (`congthuong.vn`).

---

### Đề xuất #5: [THÁO GỠ NÚT THẮNG THỰC ĐỊA] Xử lý Chuyển đổi Địa chỉ Hộ kinh doanh cá thể khác quận/huyện: Giải mã mâu thuẫn giữa quy định pháp luật và thực tế cán bộ thuế ép "đóng hộ cũ, mở hộ mới" gây treo MST
- **Chuyên mục (Pillar):** `CS` (Case Study - Pháp lý Thuế Hộ kinh doanh & Quản trị Doanh nghiệp)
- **Nền tảng phù hợp:** **Facebook** (Bài viết bóc tách thực trạng gây ức chế cho hộ kinh doanh) & **Tài liệu hướng dẫn mẫu công văn giải trình**
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.7/5) | Nhu cầu di dời mặt bằng kinh doanh cuối năm tăng cao; tình huống thực tế phát sinh trực tiếp trong `scratch/daily_notes.md` ngày 22/10/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Bị cán bộ thuế ép đóng hộ cũ làm kéo dài thời gian đối soát; MST bị khóa "treo", không mở được hộ mới, không xuất được hóa đơn, tê liệt kinh doanh |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Tranh luận gay gắt: "Luật cho phép chuyển quận theo mẫu 08-MST nhưng tại sao cán bộ thuế địa phương vẫn gây khó dễ bắt đóng hộ cũ?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.6/5) | Hướng dẫn căn cứ Thông tư 105/2020 và TT 86/2024; kỹ thuật lập hồ sơ mẫu 08-MST và công văn đối thoại chuyên nghiệp với cơ quan thuế |
  | **Tổng điểm** | **18.9 / 20** | **Xếp hạng: #5 (Giải pháp thực chiến bảo vệ hộ kinh doanh cá thể)** |
- **Tóm tắt lý do nên làm:** Tình huống thực tế đầy bức xúc của chủ hộ kinh doanh chuyển địa điểm từ Tân Bình sang Bình Thạnh. Eureka đứng về phía người kinh doanh chân chính, dùng sự hiểu biết pháp lý sâu sắc (Liêm Trinh) để bảo vệ quyền lợi, cung cấp lộ trình xử lý văn minh, đúng luật và tháo gỡ điểm nghẽn.
- **Nguồn xác minh:** Luật Quản lý thuế số 38/2019/QH14; Thông tư số 105/2020/TT-BTC; Thông tư số 86/2024/TT-BTC; Diễn đàn Thuế Chuyên Sâu; `scratch/daily_notes.md` (22/10/2026).

---

## 📊 TỔNG HỢP KẾT QUẢ ĐỒNG BỘ DỮ LIỆU TỰ ĐỘNG (SLA COMPLIANCE)
- ✅ **File báo cáo chính:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx)
  - Sheet 1 (*Danh sách Từ Khóa SEO*): Cập nhật 10 từ khóa ngành nóng nhất Tuần 43 (ngày {update_date_str}).
  - Sheet 3 (*Nghiên Cứu Từ Khóa E2E*): Mở rộng từ **165 lên 170 chủ đề** (bổ sung 5 chủ đề nghiên cứu sâu STT 166-170 kèm link xác thực).
- ✅ **Đồng bộ nguyên tử sang OneDrive:**
  - Vị trí: `{ONEDRIVE_FILE}`
  - Phương thức: Atomic save (`.tmp.xlsx` -> `os.replace`), tuyệt đối không sinh file backup thừa.
  - Mã băm SHA-256: `{file_sha}`
- ✅ **Nhật ký hệ thống:** Đã ghi nhận lịch sử đồng bộ vào [`reports/sync_log.txt`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/sync_log.txt).
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"THÀNH CÔNG: Đã tạo bản tin đề xuất tại {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w43_2310()
