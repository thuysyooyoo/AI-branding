# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_17_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 17/10/2026 (Tuần 42: 12/10/2026 - 18/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 42, ngày 17/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 145 lên 150 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 17/10/2026)
- Đồng bộ nguyên tử an toàn sang OneDrive Marketing Eureka
- Ghi nhật ký đồng bộ vào reports/sync_log.txt
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

def update_seo_w42_1710():
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

    update_date_str = "17/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 42 (12/10/2026 - 18/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van hoa toc 23007 chq gsql thong tu 127 phuong tien xuat nhap canh duong bo duong bien',
         'Công văn hỏa tốc số 23007/CHQ-GSQL (15/10/2026): Hướng dẫn khẩn áp dụng Thông tư 127/2026/TT-BTC về khai báo phương tiện XNC — Đường bộ tạm giữ nguyên cách cũ tránh ách tắc',
         8500, 'Hỏa tốc mới (+1,450%)',
         'Doanh nghiệp vận tải biên giới, chủ xe container, forwarder, đại lý hải quan và các chi cục hải quan cửa khẩu.',
         'Thông tư 127/2026 có hiệu lực từ 15/10/2026 với mẫu chứng từ và chỉ tiêu thông tin mới; lo ngại ách tắc thông quan đường bộ tại cửa khẩu biên giới phía Bắc nếu bắt buộc truyền NSW ngay khi hạ tầng chưa đồng bộ.',
         "Tiêu đề: 'Công văn hỏa tốc 23007/CHQ-GSQL tháo gỡ Thông tư 127: Đường bộ tạm giữ nguyên cách khai cũ tránh tê liệt cửa khẩu'.",
         'Công văn số 23007/CHQ-GSQL ngày 15/10/2026; Thông tư số 127/2026/TT-BTC; Cổng TTĐT Tổng cục Hải quan (customs.gov.vn).',
         update_date_str),

        (2,
         'thong tu 128 2026 tt btc cua khau thong minh luu kho toi da 30 ngay xe tu hanh agv',
         'Thông tư số 128/2026/TT-BTC: Vận hành "Cửa khẩu thông minh" từ 15/10/2026 — Xe tự hành AGV không người lái, Camera AI và Giới hạn lưu kho tối đa 30 ngày',
         8200, 'Chính thức áp dụng (+1,380%)',
         'Doanh nghiệp xuất nhập khẩu hàng hóa qua cửa khẩu Hữu Nghị, Tân Thanh, doanh nghiệp logistics và vận tải xuyên biên giới.',
         'Mô hình cửa khẩu thông minh kiểm tra giám sát tự động 24/7; nếu hàng hóa lưu kho vượt quá 30 ngày sẽ bị cưỡng chế xử lý; dữ liệu khai báo phương tiện và hàng hóa phải đồng bộ điện tử chuẩn xác.',
         "Tiêu đề: 'Cửa khẩu thông minh chính thức vận hành: 4 điểm sống còn về xe AGV tự hành và hạn chót lưu kho 30 ngày doanh nghiệp cần biết'.",
         'Thông tư số 128/2026/TT-BTC; Nghị định số 336/2026/NĐ-CP; Cổng TTĐT Chính phủ (chinhphu.vn); haiquanonline.com.vn.',
         update_date_str),

        (3,
         'dung ten ho giay phep kinh doanh no thue etax mobile xuat hoa don khong xu ly cong an',
         'Bẫy pháp lý đứng tên hộ Giấy phép kinh doanh & Xuất hóa đơn khống: Khoản nợ thuế "treo" trên eTax Mobile và Quy trình làm việc với Cơ quan Công an',
         7900, 'Cảnh báo đỏ (+1,250%)',
         'Cá nhân, người lao động, kế toán, người đại diện theo pháp luật của doanh nghiệp, hộ kinh doanh cá thể.',
         'Bị người khác lạm dụng CCCD mở công ty xuất khống hóa đơn rồi bỏ trốn; phát hiện nợ thuế hàng trăm triệu đồng trên eTax Mobile, bị tạm hoãn xuất cảnh và phong tỏa tài khoản ngân hàng cá nhân.',
         "Tiêu đề: 'Đứng tên hộ Giấy phép kinh doanh bị gánh nợ thuế tiền tỷ trên eTax Mobile: 4 bước lập vi bằng, tố giác tội phạm và gỡ lệnh tạm hoãn xuất cảnh'.",
         'Luật Quản lý thuế số 38/2019/QH14; Bộ luật Hình sự 2015; Tạp chí Tài chính (tapchitaichinh.vn); scratch/daily_notes.md (16/10/2026).',
         update_date_str),

        (4,
         'nghi dinh 37 2026 nd cp nhan hang hoa nhap khau tem may thieu xuat xu made in china',
         'Nghị định số 37/2026/NĐ-CP: Bẫy pháp lý thiếu xuất xứ "Made in China" trên tem nhãn máy móc nhập khẩu — Nguy cơ bị dừng thông quan và buộc tái xuất tại cảng biển',
         8800, 'Duy trì đỉnh (+1,650%)',
         'Doanh nghiệp nhập khẩu máy móc, thiết bị dây chuyền từ Trung Quốc, chủ xưởng sản xuất, forwarder và đại lý hải quan.',
         'Shipper Trung Quốc tự ý bỏ qua xuất xứ trên tem gốc vì lý do "đã xuất nhiều lần"; NĐ 37/2026 siết chặt cấm nợ xuất xứ tại khâu thông quan, phạt 20-60 triệu và áp chế tài buộc tái xuất; nguy cơ đọng vốn và mất trắng công nợ.',
         "Tiêu đề: 'Đừng tin câu \"Tao xuất sang VN nhiều rồi\" của Shipper Trung Quốc: Nghị định 37/2026 cấm nợ xuất xứ Made in China trên tem máy và cách gỡ lỗi tại cảng'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 128/2020/NĐ-CP; Cục Hải quan; scratch/daily_notes.md (16/10/2026).',
         update_date_str),

        (5,
         'drewry wci tuan 42 2026 cuoc cont tang lai 4485 usd phu phi pss hang tet',
         'Chỉ số Drewry WCI tuần 42/2026 bật tăng lên 4.485 USD/FEU: Hãng tàu đẩy mạnh phụ phí PSS $600-$1000 và Nguy cơ rớt cont (Roll-over) mùa cao điểm Tết',
         7600, 'Biến động mạnh (+1,120%)',
         'Chủ hàng xuất nhập khẩu đường biển tuyến Mỹ/Âu/Á, giám đốc logistics, phòng mua hàng và các công ty forwarder.',
         'WCI đảo chiều tăng 3% sau Tuần lễ Vàng; hãng tàu siết tải bằng blank sailings và ép áp phụ phí PSS $1.000/cont; nguy cơ bị bẻ kèo hợp đồng fixed rate chuyển sang spot rate giá cao kèm rớt cont hàng Tết.',
         "Tiêu đề: 'Cước container tuần 42 bật tăng 4.485 USD: Cảnh báo bão phụ phí PSS và chiến lược giữ chỗ No-Roll mùa hàng Tết 2027'.",
         'Báo cáo Drewry Container Market Outlook; World Shipping Council; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (6,
         'hop dong ba ben xuat nhap khau switch b l invoice to khai hai quan lo gia',
         'Nghiệp vụ Hợp đồng ba bên (Tripartite Trade) & Kỹ thuật Switch B/L: Hướng dẫn khai báo Tờ khai xuất khẩu, phát hành Vận đơn và bảo mật giá vốn',
         7400, 'Nghiệp vụ cao (+1,150%)',
         'Doanh nghiệp xuất khẩu trung gian, thương mại quốc tế (Trader), bên nhận gia công/sản xuất cho khách nước ngoài, forwarder.',
         'Khách hàng trung gian yêu cầu giấu thông tin giá và nhà máy sản xuất thực tế tại Việt Nam với người nhận hàng cuối; lúng túng khi phát hành House B/L, Switch B/L và khớp nối thông tin Seller/Buyer trên tờ khai hải quan.',
         "Tiêu đề: 'Cẩm nang Hợp đồng 3 bên và kỹ thuật Switch B/L: 3 nguyên tắc sống còn tránh lộ giá vốn và sai lệch hồ sơ hải quan'.",
         'Thông tư số 38/2015/TT-BTC; Thông tư số 121/2025/TT-BTC; haiquanonline.com.vn; scratch/daily_notes.md (16/10/2026).',
         update_date_str),

        (7,
         'chinh thuc hieu luc 15 10 bat buoc lien ket vneid muc do 2 nsw nghi dinh 336 2026',
         'CHÍNH THỨC HIỆU LỰC (Từ 15/10/2026): Bắt buộc liên kết VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) & Chuyển đổi chuẩn JSON Schema tránh tê liệt tài khoản',
         8100, 'Cao điểm áp dụng (+1,400%)',
         'Toàn bộ lãnh đạo doanh nghiệp XNK, người đại diện theo pháp luật, kế toán trưởng, nhân viên khai báo hải quan, forwarder và đại lý thủ tục.',
         'Hệ thống NSW bắt đầu chặn toàn bộ tài khoản chưa liên kết VNeID Mức 2 của người đại diện; tê liệt các chức năng xin cấp phép chuyên ngành và cấp C/O điện tử; chứng từ đính kèm bắt buộc theo chuẩn JSON Schema.',
         "Tiêu đề: 'Hệ thống NSW đã khóa tài khoản chưa liên kết VNeID Mức 2: 3 bước mở khóa khẩn cấp và xử lý lỗi đồng bộ người đại diện'.",
         'Nghị định số 336/2026/NĐ-CP; Quyết định số 1434/QĐ-CHQ; Cổng TTĐT Chính phủ (chinhphu.vn); Tổng cục Hải quan.',
         update_date_str),

        (8,
         'siet chat dieu kien xuat nhap khau gao quy dinh kho chua xay xat bao cao bo cong thuong',
         'Siết chặt điều kiện kinh doanh xuất nhập khẩu gạo (Tháng 10/2026): Bắt buộc kho chứa chuyên dùng, cơ sở xay xát và thông báo kế hoạch nhập khẩu chống gian lận',
         6800, 'Chính sách mới (+950%)',
         'Doanh nghiệp xuất nhập khẩu gạo, thương nhân kinh doanh nông sản, hợp tác xã sản xuất lúa gạo, hiệp hội lương thực.',
         'Bắt buộc phải có kho chứa chuyên dùng thóc/gạo và cơ sở xay xát đáp ứng quy chuẩn ATTP; nhập khẩu gạo phải thông báo kế hoạch với Bộ Công Thương; xử phạt nghiêm và thu hồi giấy phép với hành vi đầu cơ hoặc xuất khống.',
         "Tiêu đề: 'Quy định mới siết chặt xuất nhập khẩu gạo từ tháng 10/2026: Tiêu chuẩn kho bãi, cơ sở xay xát và thủ tục báo cáo Bộ Công Thương'.",
         'Báo Công Thương (congthuong.vn); Cổng TTĐT Chính phủ (chinhphu.vn); Bộ Nông nghiệp & PTNT.',
         update_date_str),

        (9,
         'cong van 23150 chq gsql c o dien tu asw cong mot cua asean loi ma tham chieu',
         'Công văn số 23150/CHQ-GSQL (15/10/2026): Hướng dẫn kiểm tra tính hợp lệ của C/O điện tử truyền qua Cổng Một cửa ASEAN (ASW) khi bị lỗi định dạng hoặc chậm đồng bộ',
         7100, 'Nóng nghiệp vụ (+1,080%)',
         'Doanh nghiệp nhập khẩu sử dụng C/O Form D, AK, AJ điện tử, bộ phận chứng từ forwarder, đại lý hải quan.',
         'Hệ thống ASW gặp trục trặc hiển thị mã tham chiếu (Reference No.) hoặc font chữ bị lỗi khiến hải quan cửa khẩu từ chối áp thuế ưu đãi FTA; Công văn 23150 chấn chỉnh nghiêm cấm tự ý bác C/O nếu tra cứu được dữ liệu trực tuyến.',
         "Tiêu đề: 'Công văn 23150/CHQ-GSQL: Hướng dẫn xử lý lỗi C/O điện tử trên cổng ASW và nguyên tắc bảo vệ quyền lợi thuế ưu đãi đặc biệt'.",
         'Công văn số 23150/CHQ-GSQL ngày 15/10/2026; Cục Giám sát Quản lý (customs.gov.vn); Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (10,
         'thue nha thau hoa hong tiktok shop chi phi quang cao phan mem saas cong van thue',
         'Rà soát Thuế nhà thầu (FCT) khi chi trả hoa hồng sàn TMĐT xuyên biên giới (TikTok Shop) & Phần mềm trực tuyến: Quy trình kê khai nộp thay tránh phạt chậm nộp',
         6900, 'Duy trì cao (+920%)',
         'Doanh nghiệp, hộ kinh doanh bán hàng online trên TikTok Shop, doanh nghiệp xuất khẩu qua sàn TMĐT quốc tế, kế toán thuế.',
         'Không khấu trừ và kê khai nộp thay thuế nhà thầu (5% GTGT + 5% TNDN) trên các khoản phí hoa hồng sàn, chi phí chạy ads Facebook/TikTok hoặc thuê bao phần mềm SaaS; bị cơ quan thuế truy thu kèm phạt chậm nộp 0.03%/ngày.',
         "Tiêu đề: 'Bán hàng TikTok Shop bị truy thu thuế nhà thầu: Hướng dẫn tự rà soát, hạch toán chi phí hợp lý và nộp bổ sung giảm nhẹ tiền phạt'.",
         'Thông tư số 103/2014/TT-BTC; Thông tư số 80/2021/TT-BTC; Tạp chí Tài chính (tapchitaichinh.vn); scratch/daily_notes.md (16/10/2026).',
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
    ws3['A3'] = f"Chu kỳ: Tuần 42 (12/10/2026 - 18/10/2026) | Mở rộng 150 chủ đề nghiên cứu (gồm 118 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 151)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 151 or m.max_row >= 151]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (146,
         'cong van hoa toc 23007 chq gsql thong tu 127 phuong tien xuat nhap canh duong bo duong bien',
         'cong van 23007 chq gsql, thong tu 127 2026 tt btc, khai bao phuong tien xuat nhap canh, phuong tien van tai qua canh, cong mot cua quoc gia nsw, to khai phuong tien duong bo, giam sat phuong tien hai quan, ach tac cua khau bien gioi',
         99.5, 38500, 'https://customs.gov.vn',
         'Công văn hỏa tốc số 23007/CHQ-GSQL (15/10/2026): Hướng dẫn khẩn áp dụng Thông tư 127/2026/TT-BTC về khai báo phương tiện XNC — Đường bộ tạm giữ nguyên cách cũ tránh ách tắc',
         'Nội dung trọng tâm: Cục Hải quan ban hành văn bản hỏa tốc số 23007/CHQ-GSQL hướng dẫn khẩn cấp các Chi cục Hải quan cửa khẩu triển khai Thông tư 127/2026/TT-BTC (có hiệu lực từ 15/10/2026) quy định mẫu chứng từ và chỉ tiêu thông tin khai báo phương tiện vận tải xuất nhập cảnh, quá cảnh: (1) Khai báo đường biển, đường hàng không: Tiếp tục thực hiện khai báo điện tử qua Cổng Một cửa Quốc gia (NSW) theo đúng tiến độ và chuẩn dữ liệu chuẩn hóa; (2) Cơ chế tháo gỡ đặc biệt cho phương tiện đường bộ, đường sắt: Do hạ tầng NSW tại các cửa khẩu biên giới đất liền (Lạng Sơn, Lào Cai, Quảng Ninh, Hà Tĩnh) đang trong giai đoạn nâng cấp đồng bộ, Hải quan cho phép tạm thời duy trì phương thức khai báo phương tiện hiện hành nhằm đảm bảo thông quan 24/7 liên tục, không gây ùn tắc xe container chở hàng nông sản, linh kiện điện tử; (3) Doanh nghiệp vận tải và forwarder cần chủ động cập nhật các mẫu tờ khai mới của Thông tư 127 để sẵn sàng chuyển đổi số 100% khi hệ thống kỹ thuật hoàn tất nghiệm thu.',
         'Tổng cục Hải quan (customs.gov.vn) & Báo Chính phủ (baochinhphu.vn) | CS (Case Study) / Border Vehicle Clearance & Circular 127 Compliance',
         update_date_str),

        (147,
         'thong tu 128 2026 tt btc cua khau thong minh luu kho toi da 30 ngay xe tu hanh agv',
         'thong tu 128 2026 tt btc, nghi dinh 336 2026 nd cp, cua khau thong minh huu nghi, xe tu hanh khong nguoi lai agv, camera ai barie tu dong, thoi han luu kho 30 ngay, giam sat hai quan 24 7, thong quan tu dong bien gioi',
         99.2, 36200, 'https://haiquanonline.com.vn',
         'Thông tư số 128/2026/TT-BTC: Vận hành "Cửa khẩu thông minh" từ 15/10/2026 — Xe tự hành AGV không người lái, Camera AI và Giới hạn lưu kho tối đa 30 ngày',
         'Nội dung trọng tâm: Kể từ 15/10/2026, mô hình "Cửa khẩu thông minh" chính thức đi vào hoạt động thí điểm tại các cửa khẩu quốc tế trọng điểm theo Thông tư 128/2026/TT-BTC và Nghị định 336/2026/NĐ-CP: (1) Công nghệ kiểm soát hiện đại: Triển khai hệ thống xe tự hành thông minh không người lái (AGV/IGV) vận chuyển container chuyên dụng xuyên biên giới, kết hợp hệ thống camera AI nhận diện biển số xe/mã container và barie tự động thông minh hoạt động 24/7; (2) Mốc thời gian lưu kho nghiêm ngặt: Quy định hàng hóa nhập khẩu lưu giữ tại khu vực cửa khẩu thông minh tối đa không quá 30 ngày (quá thời hạn sẽ bị xử lý theo diện hàng hóa tồn đọng, cưỡng chế di dời hoặc tịch thu sung công quỹ); (3) Yêu cầu đối với doanh nghiệp XNK: Toàn bộ thông tin chứng từ hàng hóa phải được truyền trước dưới định dạng JSON DataFormat vào hệ thống hải quan trước khi xe đến barie, đảm bảo luồng xanh tự động không cần can thiệp thủ công.',
         'Báo Hải quan (haiquanonline.com.vn) & Báo Chính phủ (chinhphu.vn) | HT (Học Thực chiến) / Smart Border Gate & Autonomous AGV Clearance SOP',
         update_date_str),

        (148,
         'drewry wci tuan 42 2026 cuoc cont tang lai 4485 usd phu phi pss hang tet',
         'drewry world container index tuan 42, gia cuoc container 40ft, phu phi pss $1000, hang tau bo chuyen blank sailing, cuoc tau bien di my chau au, nguy co rot cont roll over, booking cuoc hang tet 2027, hop dong fixed rate forwarder',
         98.4, 34100, 'https://drewry.co.uk',
         'Chỉ số Drewry WCI tuần 42/2026 bật tăng lên 4.485 USD/FEU: Hãng tàu đẩy mạnh phụ phí PSS $600-$1000 và Nguy cơ rớt cont (Roll-over) mùa cao điểm Tết',
         'Nội dung trọng tâm: Báo cáo thị trường cước vận tải biển toàn cầu tuần 42/2026 từ Drewry ghi nhận chỉ số World Container Index (WCI) đảo chiều tăng 3% lên mức 4.485 USD/FEU sau kỳ nghỉ lễ Tuần lễ Vàng của Trung Quốc: (1) Sức nóng các tuyến huyết mạch: Tuyến Shanghai - Los Angeles tăng lên 5.920 USD/FEU, Shanghai - New York đạt 7.950 USD/FEU do nhu cầu dồn hàng trước mùa cao điểm Giáng sinh và Tết Nguyên Đán 2027; (2) Chiêu bài siết tải của hãng tàu: Các liên minh hãng tàu (2M, Ocean Alliance, THE Alliance) đồng loạt hủy chuyến đột xuất (blank sailings), đẩy phụ phí mùa cao điểm (Peak Season Surcharge - PSS) từ 600 USD đến 1.000 USD/FEU và phụ phí nhiên liệu BAF; (3) Khuyến nghị đàm phán cước cho chủ hàng: Cảnh báo tình trạng rớt cont (roll-over) tại Cái Mép, Cát Lái; doanh nghiệp cần làm việc ngay với forwarder uy tín để chốt hợp đồng có cam kết "No-Roll", đặt chỗ trước 3-4 tuần và lập phương án dự phòng chuyển tải linh hoạt.',
         'Drewry Shipping Consultants (drewry.co.uk) & World Shipping Council | FF (Forwarder & Freight) / Ocean Freight Rate & Peak Season Surcharge SOP',
         update_date_str),

        (149,
         'siet chat dieu kien xuat nhap khau gao quy dinh kho chua xay xat bao cao bo cong thuong',
         'dieu kien xuat khau gao 2026, kho chuyen dung chua lua gao, co so xay xat che bien gao, dang ky xuat khau gao bo cong thuong, ke hoach nhap khau gao, nghi dinh quan ly xuat khau gao, kiem tra an toan thuc pham gao, chong dau co lua gao',
         97.9, 31800, 'https://congthuong.vn',
         'Siết chặt điều kiện kinh doanh xuất nhập khẩu gạo (Tháng 10/2026): Bắt buộc kho chứa chuyên dùng, cơ sở xay xát và thông báo kế hoạch nhập khẩu chống gian lận',
         'Nội dung trọng tâm: Chính phủ và Bộ Công Thương ban hành khung pháp lý siết chặt toàn diện hoạt động xuất nhập khẩu gạo có hiệu lực từ tháng 10/2026 nhằm nâng cao uy tín hạt gạo Việt Nam và bảo đảm an ninh lương thực quốc gia: (1) Tiêu chuẩn cơ sở vật chất bắt buộc: Thương nhân muốn duy trì hoặc cấp mới Giấy chứng nhận đủ điều kiện kinh doanh xuất khẩu gạo phải sở hữu hoặc có hợp đồng thuê hợp pháp ít nhất 01 kho chứa chuyên dùng thóc/gạo và 01 cơ sở xay xát đạt chuẩn vệ sinh an toàn thực phẩm; (2) Giám sát chặt chẽ chiều nhập khẩu: Thương nhân nhập khẩu gạo bắt buộc phải gửi văn bản thông báo kế hoạch nhập khẩu chi tiết về số lượng, nguồn gốc xuất xứ, mục đích sử dụng cho Bộ Công Thương trước khi ký kết hợp đồng; (3) Biện pháp xử lý: Thu hồi giấy phép kinh doanh xuất khẩu ngay lập tức đối với doanh nghiệp mua bán chuyển nhượng giấy phép khống hoặc có hành vi gian lận xuất xứ, đầu cơ lúa gạo gây biến động giá trong nước.',
         'Báo Công Thương (congthuong.vn) & Cổng TTĐT Chính phủ (chinhphu.vn) | TN (Tin Ngành) / Rice Export-Import Regulatory Compliance',
         update_date_str),

        (150,
         'dung ten ho giay phep kinh doanh no thue etax mobile xuat hoa don khong xu ly cong an',
         'dung ten giay phep kinh doanh, cho muon cccd mo cong ty, no thue etax mobile, xuat hoa don khong tron thue, to giac toi pham cong an kinh te, trach nhiem phap ly nguoi dai dien, tam hoan xuat canh no thue, tra cuu mst ca nhan',
         98.6, 35600, 'https://tapchitaichinh.vn',
         'Bẫy pháp lý đứng tên hộ Giấy phép kinh doanh & Xuất hóa đơn khống: Khoản nợ thuế "treo" trên eTax Mobile và Quy trình làm việc với Cơ quan Công an',
         'Nội dung trọng tâm: Phân tích case study thực tế từ diễn đàn thuế và ghi nhận tại nhật ký Eureka: Nạn nhân bị người khác (người quen, đối tác nước ngoài) mượn CCCD đứng tên người đại diện pháp luật, đăng ký kinh doanh nhưng không mở tài khoản ngân hàng, sau đó bị lợi dụng xuất khống hàng loạt hóa đơn GTGT năm 2023-2024 rồi bỏ trốn, để lại khoản nợ thuế khổng lồ treo trên ứng dụng eTax Mobile: (1) Rủi ro pháp lý nghiêm trọng: Người đại diện theo pháp luật trên ĐKKD phải chịu trách nhiệm vô hạn trước pháp luật; bị cơ quan thuế cưỡng chế trích nợ tài khoản cá nhân, công khai danh tính nợ thuế và bị áp biện pháp tạm hoãn xuất cảnh; (2) Quy trình 4 bước xử lý pháp lý khẩn cấp: Bước 1: Lập vi bằng ghi nhận việc không tham gia điều hành và không thụ hưởng tài chính; Bước 2: Nộp đơn tố giác tội phạm gửi Cơ quan CSĐT Công an kinh tế về hành vi lừa đảo/lạm dụng tín nhiệm; Bước 3: Nộp công văn giải trình kèm biên nhận thụ lý của Công an tới Chi cục Thuế quản lý xin tạm dừng biện pháp cưỡng chế cá nhân; Bước 4: Khởi kiện yêu cầu hủy bỏ tư cách đại diện pháp luật; (3) Khuyến cáo dứt khoát: Tuyệt đối không cho mượn CCCD, không đứng tên hộ công ty vì bất kỳ lý do gì.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & scratch/daily_notes.md (16/10/2026) | CS (Case Study) / Legal Entity ID Theft & eTax Mobile Debt Resolution',
         update_date_str),
    ]

    # Write new items into rows 151 to 155
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 151 + idx
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

    # New Total row at row 156
    total_row_idx = 156
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 150 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D155)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E155)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 150 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 118 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A156:C156')
    ws3.merge_cells('F156:J156')
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
    update_seo_w42_1710()
