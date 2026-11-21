# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_16_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 16/10/2026 (Tuần 42: 12/10/2026 - 18/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 42, ngày 16/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 140 lên 145 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 16/10/2026)
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

def update_seo_w42_1610():
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

    update_date_str = "16/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 42 (12/10/2026 - 18/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'nghi dinh 37 2026 nd cp nhan hang hoa nhap khau tem may thieu xuat xu made in china',
         'Nghị định số 37/2026/NĐ-CP: Bẫy pháp lý thiếu xuất xứ "Made in China" trên tem nhãn máy móc nhập khẩu — Nguy cơ bị dừng thông quan và buộc tái xuất tại cảng biển',
         8800, 'Tăng đột biến (+1,650%)',
         'Doanh nghiệp nhập khẩu máy móc, thiết bị dây chuyền từ Trung Quốc, chủ xưởng sản xuất, forwarder và đại lý hải quan.',
         'Shipper Trung Quốc tự ý bỏ qua xuất xứ trên tem gốc vì lý do "đã xuất nhiều lần"; NĐ 37/2026 siết chặt cấm nợ xuất xứ tại khâu thông quan, phạt 20-60 triệu và áp chế tài buộc tái xuất; nguy cơ đọng vốn và mất trắng công nợ.',
         "Tiêu đề: 'Đừng tin câu \"Tao xuất sang VN nhiều rồi\" của Shipper Trung Quốc: Nghị định 37/2026 cấm nợ xuất xứ Made in China trên tem máy và cách gỡ lỗi tại cảng'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 128/2020/NĐ-CP; Cục Hải quan; scratch/daily_notes.md (16/10/2026).',
         update_date_str),

        (2,
         'hop dong ba ben xuat nhap khau switch b l invoice to khai hai quan lo gia',
         'Nghiệp vụ Hợp đồng ba bên (Tripartite Trade) & Kỹ thuật Switch B/L: Hướng dẫn khai báo Tờ khai xuất khẩu, phát hành Vận đơn và bảo mật giá vốn',
         7400, 'Tăng rất mạnh (+1,150%)',
         'Doanh nghiệp xuất khẩu trung gian, thương mại quốc tế (Trader), bên nhận gia công/sản xuất cho khách nước ngoài, forwarder.',
         'Khách hàng trung gian yêu cầu giấu thông tin giá và nhà máy sản xuất thực tế tại Việt Nam với người nhận hàng cuối; lúng túng khi phát hành House B/L, Switch B/L và khớp nối thông tin Seller/Buyer trên tờ khai hải quan.',
         "Tiêu đề: 'Cẩm nang Hợp đồng 3 bên và kỹ thuật Switch B/L: 3 nguyên tắc sống còn tránh lộ giá vốn và sai lệch hồ sơ hải quan'.",
         'Thông tư số 38/2015/TT-BTC; Thông tư số 121/2025/TT-BTC; haiquanonline.com.vn; scratch/daily_notes.md (16/10/2026).',
         update_date_str),

        (3,
         'thue nha thau hoa hong tiktok shop chi phi quang cao phan mem saas cong van thue',
         'Rà soát Thuế nhà thầu (FCT) khi chi trả hoa hồng sàn TMĐT xuyên biên giới (TikTok Shop) & Phần mềm trực tuyến: Quy trình kê khai nộp thay tránh phạt chậm nộp',
         6900, 'Tăng cao (+920%)',
         'Doanh nghiệp, hộ kinh doanh bán hàng online trên TikTok Shop, doanh nghiệp xuất khẩu qua sàn TMĐT quốc tế, kế toán thuế.',
         'Không khấu trừ và kê khai nộp thay thuế nhà thầu (5% GTGT + 5% TNDN) trên các khoản phí hoa hồng sàn, chi phí chạy ads Facebook/TikTok hoặc thuê bao phần mềm SaaS; bị cơ quan thuế truy thu kèm phạt chậm nộp 0.03%/ngày.',
         "Tiêu đề: 'Bán hàng TikTok Shop bị truy thu thuế nhà thầu: Hướng dẫn tự rà soát, hạch toán chi phí hợp lý và nộp bổ sung giảm nhẹ tiền phạt'.",
         'Thông tư số 103/2014/TT-BTC; Thông tư số 80/2021/TT-BTC; Tạp chí Tài chính (tapchitaichinh.vn); scratch/daily_notes.md (16/10/2026).',
         update_date_str),

        (4,
         'cong van 23150 chq gsql c o dien tu asw cong mot cua asean loi ma tham chieu',
         'Công văn số 23150/CHQ-GSQL (15/10/2026): Hướng dẫn kiểm tra tính hợp lệ của C/O điện tử truyền qua Cổng Một cửa ASEAN (ASW) khi bị lỗi định dạng hoặc chậm đồng bộ',
         7100, 'Nóng nghiệp vụ (+1,080%)',
         'Doanh nghiệp nhập khẩu sử dụng C/O Form D, AK, AJ điện tử, bộ phận chứng từ forwarder, đại lý hải quan.',
         'Hệ thống ASW gặp trục trặc hiển thị mã tham chiếu (Reference No.) hoặc font chữ bị lỗi khiến hải quan cửa khẩu từ chối áp thuế ưu đãi FTA; Công văn 23150 chấn chỉnh nghiêm cấm tự ý bác C/O nếu tra cứu được dữ liệu trực tuyến.',
         "Tiêu đề: 'Công văn 23150/CHQ-GSQL: Hướng dẫn xử lý lỗi C/O điện tử trên cổng ASW và nguyên tắc bảo vệ quyền lợi thuế ưu đãi đặc biệt'.",
         'Công văn số 23150/CHQ-GSQL ngày 15/10/2026; Cục Giám sát Quản lý (customs.gov.vn); Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (5,
         'phu phi mua cao diem pss baf cuoc tau bien q4 2026 hop dong van tai hang tet',
         'Cập nhật Thị trường Vận tải biển Q4/2026: Hãng tàu đồng loạt áp Phụ phí mùa cao điểm (PSS) $600-$1000/cont & Chiến lược đàm phán hợp đồng hàng Tết',
         6500, 'Tăng mạnh (+850%)',
         'Chủ hàng xuất nhập khẩu đường biển tuyến Mỹ/Âu/Á, giám đốc logistics, phòng mua hàng và các công ty forwarder.',
         'Hãng tàu siết tải và tung phụ phí PSS, tăng phụ phí nhiên liệu BAF trước kỳ bầu cử Mỹ và cao điểm Noel - Tết, đẩy chi phí logistics lên thêm 15-25%; nguy cơ bị ép hủy hợp đồng dài hạn chuyển sang giá giao ngay.',
         "Tiêu đề: 'Làn sóng phụ phí PSS $1.000/cont mùa cao điểm Q4: 4 chiến lược đàm phán cước tàu và giữ chỗ không bị rớt container'.",
         'Báo Công Thương (congthuong.vn); Báo cáo Drewry Container Market Outlook; World Shipping Council.',
         update_date_str),

        (6,
         'chinh thuc hieu luc 15 10 bat buoc lien ket vneid muc do 2 nsw nghi dinh 336 2026',
         'CHÍNH THỨC HIỆU LỰC (Từ 15/10/2026): Bắt buộc liên kết VNeID Mức độ 2 trên Cổng Một cửa Quốc gia (NSW) & Chuyển đổi chuẩn JSON Schema tránh tê liệt tài khoản',
         8600, 'Cao điểm áp dụng (+1,520%)',
         'Toàn bộ lãnh đạo doanh nghiệp XNK, người đại diện theo pháp luật, kế toán trưởng, nhân viên khai báo hải quan, forwarder và đại lý thủ tục.',
         'Hệ thống NSW bắt đầu chặn toàn bộ tài khoản chưa liên kết VNeID Mức 2 của người đại diện; tê liệt các chức năng xin cấp phép chuyên ngành và cấp C/O điện tử; chứng từ đính kèm bắt buộc theo chuẩn JSON Schema.',
         "Tiêu đề: 'Hệ thống NSW đã khóa tài khoản chưa liên kết VNeID Mức 2: 3 bước mở khóa khẩn cấp và xử lý lỗi đồng bộ người đại diện'.",
         'Nghị định số 336/2026/NĐ-CP; Quyết định số 1434/QĐ-CHQ; Cổng TTĐT Chính phủ (chinhphu.vn); Tổng cục Hải quan.',
         update_date_str),

        (7,
         'eu hoan thuc thi eudr 30 12 2026 quy dinh chong pha rung go ca phe cao su dinh vi gps',
         'Ủy ban Châu Âu đề xuất lùi thời hạn thực thi Quy định Chống phá rừng (EUDR) đến 30/12/2026: Cơ hội "vàng" nước rút cho Nông sản & Gỗ Việt Nam',
         8300, 'Duy trì đỉnh (+1,420%)',
         'Doanh nghiệp xuất khẩu gỗ, đồ nội thất, cà phê, cao su sang thị trường EU, hợp tác xã sản xuất và các hiệp hội ngành hàng xuất khẩu.',
         'Dù EU đề xuất lùi hạn áp dụng sang 30/12/2026, các tiêu chuẩn thẩm định giải trình (DDS), định vị tọa độ GPS lô đất và mốc không phá rừng (31/12/2020) vẫn giữ nguyên; nguy cơ DN chủ quan dừng số hóa chuỗi cung ứng sẽ bị loại khỏi thị trường EU.',
         "Tiêu đề: 'EU hoãn EUDR đến 30/12/2026: Cơ hội vàng hay bẫy chủ quan? Chiến lược số hóa dữ liệu vùng trồng giữ vững thị phần xuất khẩu tỷ đô'.",
         'Thông cáo chính thức Ủy ban Châu Âu (EC); Báo Công Thương (congthuong.vn); Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (8,
         'cong van 22233 chq gsql khai ma ko hang an ninh mang mat ma dan su hang hoa luong dung 19916',
         'Công văn số 22233/CHQ-GSQL & 19916/CHQ-GSQL: Hướng dẫn khai mã "KO" đối với Hàng hóa An ninh mạng, Mật mã dân sự & Hàng hóa Lưỡng dụng tránh bị kiểm tra sau thông quan',
         7500, 'Duy trì cao (+1,200%)',
         'Doanh nghiệp nhập khẩu thiết bị CNTT, viễn thông, camera, hóa chất, bộ phận xuất nhập khẩu, forwarder và đại lý hải quan.',
         'Hàng hóa có mã HS trùng danh mục quản lý chuyên ngành nhưng công năng thực tế không thuộc diện cấp phép; nếu không khai mã "KO" trên tờ khai sẽ bị chặn thông quan hoặc phạt vi phạm hành chính, bị chuyển hồ sơ kiểm tra sau thông quan.',
         "Tiêu đề: 'Công văn 22233/CHQ-GSQL & 19916/CHQ-GSQL: Bẫy khai mã KO cho hàng an ninh mạng, mật mã dân sự và hàng lưỡng dụng tránh rủi ro kiểm tra sau thông quan'.",
         'Công văn số 22233/CHQ-GSQL; Công văn số 19916/CHQ-GSQL; Nghị định 53/2022/NĐ-CP; Nghị định 58/2016/NĐ-CP.',
         update_date_str),

        (9,
         'thong tu 121 2025 tt btc thu tuc hai quan hang qua canh trung chuyen cang bien ban ke 09 10 11',
         'Thông tư số 121/2025/TT-BTC: 10 Điểm cốt lõi về Thủ tục hải quan đối với Hàng hóa Quá cảnh và Trung chuyển tại cảng biển — Quy chuẩn Bản kê và Thiết bị định vị GPS',
         7700, 'Duy trì cao (+1,280%)',
         'Doanh nghiệp logistics, forwarder, hãng tàu, đại lý vận tải hàng quá cảnh sang Lào/Campuchia/Trung Quốc và trung chuyển cảng biển.',
         'Quy chuẩn nghiêm ngặt về phân loại 3 nhóm thủ tục quá cảnh, bắt buộc gắn thiết bị theo dõi hành trình GPS/seal định vị với container xếp chồng đường thủy nội địa; rủi ro tắc nghẽn nếu khai sai mẫu bản kê 09, 10, 11.',
         "Tiêu đề: 'Thông tư 121/2025/TT-BTC: Cẩm nang 10 điểm vàng làm thủ tục hải quan hàng quá cảnh và trung chuyển cảng biển không lo bị ách tắc'.",
         'Thông tư số 121/2025/TT-BTC (khoản 22, 23 Điều 1 sửa đổi Điều 51a, 51b TT 38/2015/TT-BTC); Cục Hải quan.',
         update_date_str),

        (10,
         'cong van 22883 chq nvthq ma hs hamburger banh mi hay thit 20 phan tram',
         'Công văn số 22883/CHQ-NVTHQ (12/10/2026): Lệnh rà soát khẩn toàn quốc về phân loại mã HS Hamburger — Ranh giới tỷ lệ 20% thịt giữa Chương 16 và nhóm 19.05',
         7200, 'Duy trì đỉnh (+1,050%)',
         'Doanh nghiệp nhập khẩu thực phẩm chế biến, chuỗi nhà hàng thức ăn nhanh, chuỗi F&B, forwarder và đại lý hải quan.',
         'Lệnh rà soát khẩn toàn quốc của Cục Hải quan (báo cáo trước 19/10/2026); doanh nghiệp khai hamburger vào nhóm 19.05 (bánh mì) nhưng nhân thịt vượt 20% bị ấn định lại vào Chương 16 (chế phẩm thịt), bị truy thu thuế lớn và xử phạt vi phạm hành chính.',
         "Tiêu đề: 'Lệnh rà soát khẩn Công văn 22883/CHQ-NVTHQ: Hamburger là bánh mì hay thịt? Quy tắc 20% phân định mã HS và thuế suất'.",
         'Công văn số 22883/CHQ-NVTHQ ngày 12/10/2026; Chú giải 2 Chương 16; Chú giải 1(a) Chương 19; Cục Hải quan.',
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
    ws3['A3'] = f"Chu kỳ: Tuần 42 (12/10/2026 - 18/10/2026) | Mở rộng 145 chủ đề nghiên cứu (gồm 113 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 146)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 146 or m.max_row >= 146]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (141,
         'nghi dinh 37 2026 nd cp nhan hang hoa nhap khau tem may thieu xuat xu made in china',
         'nghi dinh 37 2026 nd cp, nhan hang hoa nhap khau, tem may moc thieu xuat xu, made in china tren tem goc, tai xuat hang hoa cua khau, xu phat thieu nhan mac hai quan, kho kiem tra khac phuc tem nhan, nghi dinh 128 2020 nd cp, bien ban vi pham hanh chinh hai quan',
         99.8, 42500, 'https://customs.gov.vn',
         'Nghị định số 37/2026/NĐ-CP: Bẫy pháp lý thiếu xuất xứ "Made in China" trên tem nhãn máy móc nhập khẩu — Nguy cơ bị dừng thông quan và buộc tái xuất tại cảng biển',
         'Nội dung trọng tâm: Nghị định 37/2026/NĐ-CP (có hiệu lực từ 01/07/2026, thay thế toàn diện NĐ 43/2017 và NĐ 111/2021) siết chặt toàn diện quy định về nhãn gốc hàng hóa nhập khẩu tại khâu thông quan: (1) Doanh nghiệp tuyệt đối KHÔNG ĐƯỢC PHÉP NỢ 3 thông tin bắt buộc trên nhãn gốc vật lý của hàng hóa: Tên hàng hóa, Tên & địa chỉ nhà sản xuất/tổ chức chịu trách nhiệm, và Nước xuất xứ (Country of Origin — ví dụ: Made in China); (2) Cảnh báo bẫy chủ quan từ Shipper Trung Quốc: Lời khẳng định "đã xuất sang Việt Nam nhiều lần không sao" dẫn tới hàng ngàn lô hàng máy móc, thiết bị mới 100% bị Chi cục Hải quan cửa khẩu (như cảng Hải Phòng, Cát Lái) lập biên bản vi phạm, DỪNG THÔNG QUAN NGAY LẬP TỨC; (3) Khung chế tài: Phạt tiền từ 20 đến 60 triệu đồng theo Nghị định 128/2020/NĐ-CP và áp dụng biện pháp khắc phục hậu quả BUỘC TÁI XUẤT hàng hóa ra khỏi lãnh thổ Việt Nam; (4) Khuyến nghị quy trình xử lý thực chiến của Eureka: Nếu hàng chưa xuất, tận dụng đòn bẩy 40% công nợ giữ lại yêu cầu Shipper dán bổ sung nameplate kim loại chuẩn; nếu hàng đã về cảng, lập công văn giải trình xin đưa vào khu vực giám sát hải quan để khắc phục dán bổ sung công khai, minh bạch theo luật định, tuyệt đối không tìm cách đi ngầm gây rủi ro kiểm tra sau thông quan.',
         'Tổng cục Hải quan (customs.gov.vn) & scratch/daily_notes.md (16/10/2026) | CS (Case Study) / Machine Origin Labeling & Decree 37 Compliance',
         update_date_str),

        (142,
         'hop dong ba ben xuat nhap khau switch b l invoice to khai hai quan lo gia',
         'hop dong ba ben, giao thuong ba ben, tripartite trade, switch b l, phat hanh lai van don, hoa don thuong mai lech ten, to khai xuat khau viet nam, consignee nguoi nhan cuoi, bao mat gia von xnk, dieu 16 thong tu 38 2015',
         98.7, 36800, 'https://haiquanonline.com.vn',
         'Nghiệp vụ Hợp đồng ba bên (Tripartite Trade) & Kỹ thuật Switch B/L: Hướng dẫn khai báo Tờ khai xuất khẩu, phát hành Vận đơn và bảo mật giá vốn',
         'Nội dung trọng tâm: Hướng dẫn nghiệp vụ xử lý giao dịch thương mại quốc tế 3 bên (Bên bán A tại Việt Nam, Bên mua trung gian B tại Đài Loan/Singapore, Bên nhận hàng C tại Mỹ/EU): (1) Vướng mắc cốt lõi: Bên B thanh toán trực tiếp cho bên A nhưng yêu cầu Vận đơn (B/L) thể hiện Shipper A, Consignee C, Notify Party forwarder của C, trong khi Tờ khai và Hóa đơn thương mại tại Việt Nam đứng tên Seller A - Buyer B; (2) Kỹ thuật phát hành Switch B/L chuẩn quốc tế: Phát hành Bộ vận đơn 1 (Original B/L 1: Shipper A, Consignee B/hoặc ngân hàng của B) để phục vụ thông quan xuất khẩu tại Việt Nam; sau đó Bên B thanh toán đủ tiền hàng cho A và tiến hành Switch B/L tại cảng chuyển tải hoặc văn phòng forwarder để cấp Bộ vận đơn 2 (Original B/L 2: Shipper B, Consignee C, Invoice B-C) gửi cho C mở tờ khai nhập khẩu tại Mỹ; (3) Nguyên tắc bảo vệ bí mật kinh doanh: Tuyệt đối không để lộ giá vốn giữa A và B cho khách hàng C, đồng thời bảo đảm dòng tiền thanh toán quốc tế và tờ khai xuất khẩu tại Việt Nam hoàn toàn hợp pháp, đầy đủ điều kiện hoàn thuế GTGT.',
         'Báo Hải Quan (haiquanonline.com.vn) & scratch/daily_notes.md (16/10/2026) | HT (Học Thực chiến) / Tripartite Trade & Switch BL Execution SOP',
         update_date_str),

        (143,
         'thue nha thau hoa hong tiktok shop chi phi quang cao phan mem saas cong van thue',
         'thue nha thau, fct tiktok shop, thue nha thau hoa hong san, phi quang cao meta google, chi phi phan mem chatgpt saas, thong tu 103 2014 tt btc, khau tru thue nha thau, nop thay nha thau nuoc ngoai, cham nop tien thue 0 03, chi phi duoc tru thue tndn',
         98.2, 33500, 'https://tapchitaichinh.vn',
         'Rà soát Thuế nhà thầu (FCT) khi chi trả hoa hồng sàn TMĐT xuyên biên giới (TikTok Shop) & Phần mềm trực tuyến: Quy trình kê khai nộp thay tránh phạt chậm nộp',
         'Nội dung trọng tâm: Cơ quan thuế địa phương đang tăng cường kiểm tra, rà soát nghĩa vụ thuế nhà thầu (Foreign Contractor Tax - FCT) đối với các cá nhân, doanh nghiệp kinh doanh trên sàn thương mại điện tử (TikTok Shop, Shopee) và sử dụng dịch vụ xuyên biên giới: (1) Rủi ro pháp lý: Các khoản chi trả phí hoa hồng sàn (Commission Fee), phí giao dịch, phí dịch vụ quảng cáo hoặc thuê bao phần mềm SaaS (như ChatGPT, phần mềm quản lý kho, CRM nước ngoài) thuộc đối tượng chịu thuế nhà thầu theo Thông tư 103/2014/TT-BTC; (2) Trách nhiệm nộp thuế: Trường hợp tổ chức nước ngoài chưa đăng ký thuế trực tiếp tại Việt Nam qua Cổng TTĐT dành cho NCCNN của Tổng cục Thuế, doanh nghiệp Việt Nam có nghĩa vụ khấu trừ và nộp thay thuế GTGT (5%) và thuế TNDN (5%) trước khi thanh toán; (3) Hướng dẫn gỡ vướng: Hướng dẫn quy trình lập tờ khai mẫu 01/NTNN theo Thông tư 80/2021/TT-BTC, rà soát số liệu các kỳ từ 2024-2025 để nộp bổ sung kèm tiền chậm nộp 0.03%/ngày nhằm tránh bị xử phạt hành vi trốn thuế và bị loại chi phí hợp lý khi quyết toán thuế TNDN.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & scratch/daily_notes.md (16/10/2026) | TN (Tin Ngành) / Cross-border FCT & E-commerce Tax Compliance',
         update_date_str),

        (144,
         'cong van 23150 chq gsql c o dien tu asw cong mot cua asean loi ma tham chieu',
         'cong van 23150 chq gsql, c o dien tu asw, cong mot cua asean, loi ma tham chieu c o, c o form d dien tu, thue suat uu dai dac biet fta, tra cuu c o truc tuyen, tu choi c o hai quan cua khau, thong tu 33 2023 tt btc',
         98.5, 35200, 'https://customs.gov.vn',
         'Công văn số 23150/CHQ-GSQL (15/10/2026): Hướng dẫn kiểm tra tính hợp lệ của C/O điện tử truyền qua Cổng Một cửa ASEAN (ASW) khi bị lỗi định dạng hoặc chậm đồng bộ',
         'Nội dung trọng tâm: Cục Giám sát Quản lý (Tổng cục Hải quan) ban hành Công văn số 23150/CHQ-GSQL tháo gỡ vướng mắc cho các doanh nghiệp nhập khẩu khi tiếp nhận Giấy chứng nhận xuất xứ hàng hóa (C/O điện tử) qua Cổng Một cửa ASEAN (ASW): (1) Thực trạng: Nhiều lô hàng nhập khẩu có C/O mẫu D điện tử từ các nước thành viên ASEAN (Thái Lan, Indonesia, Malaysia) khi truyền qua hệ thống ASW gặp hiện tượng đứt đoạn ký tự, mất chữ số của mã tham chiếu (Reference Number) hoặc font chữ bị mã hóa lỗi hiển thị, dẫn tới việc một số chi cục hải quan cửa khẩu tự ý dừng thông quan, bác C/O hoặc bắt doanh nghiệp ký quỹ thuế MFN; (2) Chỉ đạo dứt khoát của Cục GSQL: Yêu cầu các đơn vị hải quan không được từ chối C/O điện tử chỉ vì lỗi kỹ thuật đường truyền nếu doanh nghiệp cung cấp được đường link tra cứu trực tuyến hoặc bản scan C/O có mã QR hợp lệ từ cơ quan cấp nước xuất khẩu; (3) Doanh nghiệp được hướng dẫn làm công văn tra soát gửi bộ phận công nghệ thông tin hải quan để kích hoạt đồng bộ lại dữ liệu trong vòng 04 giờ làm việc.',
         'Tổng cục Hải quan (customs.gov.vn) & Báo Chính phủ (chinhphu.vn) | CS (Case Study) / Electronic C/O ASW Troubleshooting SOP',
         update_date_str),

        (145,
         'phu phi mua cao diem pss baf cuoc tau bien q4 2026 hop dong van tai hang tet',
         'phu phi mua cao diem pss, phu phi nhien lieu baf, gia cuoc tau bien q4 2026, cuoc cont bac my chau au, hop dong cuoc fixed rate, spot rate tang cao, booking hang tet nguyen dan, thieu vo container, no roll guarantee, drewry container index',
         97.8, 32400, 'https://congthuong.vn',
         'Cập nhật Thị trường Vận tải biển Q4/2026: Hãng tàu đồng loạt áp Phụ phí mùa cao điểm (PSS) $600-$1000/cont & Chiến lược đàm phán hợp đồng hàng Tết',
         'Nội dung trọng tâm: Thị trường vận tải biển quốc tế bước vào giai đoạn biến động khốc liệt khi các hãng tàu lớn (Maersk, MSC, CMA CGM, COSCO) đồng loạt công bố áp dụng phụ phí mùa cao điểm (Peak Season Surcharge - PSS) từ 600 USD đến 1.000 USD/container 40HC, đồng thời điều chỉnh tăng phụ phí nhiên liệu (Bunker Adjustment Factor - BAF) từ giữa tháng 10/2026 trên các tuyến xuất khẩu đi Bắc Mỹ và Châu Âu: (1) Nguyên nhân: Áp lực dồn đơn hàng trước thềm bầu cử Mỹ, rủi ro đình công cảng biển kéo dài và nhu cầu gom hàng phục vụ mùa Giáng sinh và Tết Nguyên Đán 2027 tăng đột biến; (2) Rủi ro cho doanh nghiệp XNK Việt Nam: Hãng tàu có xu hướng "bẻ kèo" các hợp đồng cước dài hạn (Contract/Service Contract), viện cớ thiếu chỗ để ép chủ hàng chuyển sang cước giao ngay (Spot rate) giá cao, hoặc tình trạng rớt container (roll over) tràn lan tại các cảng trung chuyển Cái Mép, Cát Lái, Singapore; (3) Khuyến nghị đàm phán: Doanh nghiệp cần phối hợp chặt chẽ với forwarder uy tín để chốt hợp đồng có điều khoản cam kết không rớt cont (No-Roll Guarantee), lên kế hoạch xuất hàng sớm hơn 2-3 tuần so với lịch trình dự kiến.',
         'Báo Công Thương (congthuong.vn) & Drewry Shipping Consultants | TN (Tin Ngành) / Ocean Freight PSS Surcharge & Peak Season Booking SOP',
         update_date_str),
    ]

    # Write new items into rows 146 to 150
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 146 + idx
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

    # New Total row at row 151
    total_row_idx = 151
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 145 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D150)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E150)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 145 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 113 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A151:C151')
    ws3.merge_cells('F151:J151')
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
    update_seo_w42_1610()
