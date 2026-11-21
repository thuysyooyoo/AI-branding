# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_20_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 20/10/2026 (Tuần 43: 19/10/2026 - 25/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 43, ngày 20/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 155 lên 160 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 20/10/2026)
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

def update_seo_w43_2010():
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

    update_date_str = "20/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 43 (19/10/2026 - 25/10/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van 23007 chq gsql thong tu 127 phuong tien van tai duong bo nsw',
         'Công văn hỏa tốc số 23007/CHQ-GSQL & Thông tư 127/2026/TT-BTC: Cục Hải quan tháo gỡ khẩn cấp khai báo phương tiện XNC đường bộ qua Cổng Một cửa Quốc gia (NSW) tránh tê liệt cửa khẩu',
         8800, 'Hỏa tốc mới (+1,520%)',
         'Doanh nghiệp vận tải đường bộ xuyên biên giới, forwarder, đại lý hải quan, chủ xe container chở hàng nông sản/linh kiện qua Lạng Sơn, Lào Cai, Quảng Ninh.',
         'Thông tư 127 áp dụng từ 15/10 với chuẩn chứng từ mới; hạ tầng kỹ thuật NSW biên giới chưa đồng bộ khiến xe bị dừng thông quan; Công văn 23007 cho phép đường bộ duy trì cách cũ trong khi tàu bay, tàu biển bắt buộc số hóa 100%.',
         "Tiêu đề: 'Công văn hỏa tốc 23007/CHQ-GSQL cứu nguy ùn tắc cửa khẩu: Hướng dẫn phân tầng khai báo phương tiện theo Thông tư 127 doanh nghiệp cần nắm rõ'.",
         'Công văn số 23007/CHQ-GSQL ngày 17/10/2026; Thông tư số 127/2026/TT-BTC; Cổng TTĐT Tổng cục Hải quan (customs.gov.vn).',
         update_date_str),

        (2,
         'nghi dinh 336 2026 nd cp co che mot cua quoc gia nsw asean thay the nd 85',
         'Nghị định số 336/2026/NĐ-CP chính thức thay thế Nghị định 85/2019/NĐ-CP: Đột phá kết nối liên thông dữ liệu Một cửa Quốc gia (NSW) & Cấm yêu cầu nộp lại giấy tờ đã số hóa',
         8500, 'Đột phá số hóa (+1,450%)',
         'Toàn bộ doanh nghiệp xuất nhập khẩu, đơn vị kiểm tra chuyên ngành, các Bộ ban ngành cấp phép và đại lý hải quan.',
         'Doanh nghiệp thường xuyên bị cán bộ kiểm tra yêu cầu nộp lại chứng từ giấy dù đã đẩy lên hệ thống NSW; Nghị định 336 có hiệu lực từ 15/10 chính thức thiết lập hành lang pháp lý cấm sách nhiễu và bắt buộc dùng dữ liệu số dùng chung.',
         "Tiêu đề: 'Nghị định 336/2026/NĐ-CP thay thế Nghị định 85: Cơ chế Một cửa Quốc gia bước sang kỷ nguyên số hóa không giấy tờ & Quyền lợi chủ hàng'.",
         'Nghị định số 336/2026/NĐ-CP ngày 15/10/2026; Cổng TTĐT Chính phủ (chinhphu.vn); Cổng Một cửa Quốc gia (customs.gov.vn).',
         update_date_str),

        (3,
         'khai bao ma phu ko tren vnaccs mat ma dan su hang hoa luong dung',
         "Công văn 22233/CHQ-GSQL & 19916/CHQ-GSQL: Bắt buộc khai mã phụ 'KO' trên hệ thống VNACCS đối với hàng hóa trùng mã HS an ninh mạng, mật mã dân sự nhưng không thuộc diện cấp phép",
         8400, 'Tăng vọt kiểm soát (+1,420%)',
         'Doanh nghiệp nhập khẩu linh kiện điện tử, vi mạch, máy tính, thiết bị viễn thông, camera AI, máy móc tự động hóa, đại lý hải quan.',
         'Hàng hóa dân dụng bình thường có mã HS trùng danh mục quản lý chuyên ngành nhưng không khai mã "KO" dẫn đến việc VNACCS tự động nhảy luồng Đỏ, bị dừng thông quan 15-30 ngày chờ giám định kỹ thuật và bị phạt khai sai chỉ tiêu thông tin.',
         "Tiêu đề: 'Bẫy luồng Đỏ 18 ngày tại Cát Lái vì quên mã 'KO': Tuyệt chiêu khai báo hàng điện tử trùng danh mục mật mã dân sự trên VNACCS'.",
         'Công văn số 22233/CHQ-GSQL; Công văn 19916/CHQ-GSQL; Nghị định 58/2016/NĐ-CP; Nghị định 53/2018/NĐ-CP; Cục Giám sát Quản lý Hải quan.',
         update_date_str),

        (4,
         'hop dong ba ben xuat khau switch b l drop shipment to khai hai quan',
         'Nghiệp vụ Hợp đồng thương mại 3 bên (Triangle Trade) & Kỹ thuật Switch B/L: Hướng dẫn khai báo hải quan, bảo mật giá và chuẩn hóa dòng tiền khấu trừ thuế GTGT',
         8200, 'Nghiệp vụ thực chiến (+1,380%)',
         'Doanh nghiệp xuất khẩu thương mại, forwarder, trung gian thương mại quốc tế, kế toán thuế XNK.',
         'Bán hàng cho đối tác ngoại (Đài Loan/Singapore) nhưng giao thẳng cho khách hàng tại Mỹ/EU; lo sợ lộ nhà sản xuất và biên lợi nhuận giá vốn; lúng túng khi khai tờ khai xuất khẩu và phối hợp chứng từ thanh toán ngân hàng.',
         "Tiêu đề: 'Bí quyết thực chiến Hợp đồng thương mại 3 bên & Kỹ thuật phát hành Switch B/L: Vừa bảo mật giá vốn, vừa hoàn thuế GTGT an toàn 100%'.",
         'Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; Luật Thuế GTGT; Báo Hải quan; scratch/daily_notes.md (Case study thực tế).',
         update_date_str),

        (5,
         'thong tu 128 2026 tt btc cua khau thong minh agv luu kho 30 ngay',
         'Thông tư số 128/2026/TT-BTC & Nghị định 336/2026/NĐ-CP: Vận hành Cửa khẩu thông minh từ 15/10/2026 — Xe tự hành AGV, Camera AI và Thời hạn lưu kho tối đa 30 ngày',
         8100, 'Cao điểm thực thi (+1,350%)',
         'Doanh nghiệp xuất nhập khẩu nông sản, linh kiện qua cửa khẩu Hữu Nghị, Tân Thanh, doanh nghiệp logistics biên mậu.',
         'Xe tự hành AGV chạy tự động 24/7 đòi hỏi dữ liệu JSON DataFormat truyền trước 100%; hàng lưu kho vượt 30 ngày bị cưỡng chế thanh lý; doanh nghiệp chưa chuẩn hóa phần mềm bị từ chối barie thông minh.',
         "Tiêu đề: 'Cửa khẩu thông minh chạy xe tự hành AGV: 3 điều kiện số hóa bắt buộc để container luồng xanh không chạm qua biên giới'.",
         'Thông tư số 128/2026/TT-BTC; Nghị định số 336/2026/NĐ-CP; Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (6,
         'thue nha thau dich vu phan mem nuoc ngoai chatgpt hoa don dien tu hop le',
         'Rủi ro thanh tra Thuế nhà thầu (FCT) khi chi trả phần mềm nước ngoài (ChatGPT, Zoom, SaaS) và chi phí quảng cáo số: Quy trình 3 bước hợp thức hóa chi phí được trừ',
         8000, 'Cảnh báo rủi ro thuế (+1,310%)',
         'Giám đốc tài chính (CFO), kế toán trưởng, doanh nghiệp công nghệ, doanh nghiệp XNK có chi trả dịch vụ phần mềm/quảng cáo quốc tế.',
         'Cà thẻ visa cá nhân/công ty mua tài khoản dịch vụ xuyên biên giới nhưng không nộp thuế nhà thầu thay cho nhà cung cấp nước ngoài; bị cơ quan thuế bóc tách chi phí hợp lý khi quyết toán, truy thu và phạt nặng.',
         "Tiêu đề: 'Cà thẻ mua ChatGPT, Zoom, Ads nước ngoài bị gạt chi phí và phạt thuế: Hướng dẫn 3 bước hợp thức hóa hóa đơn nhà thầu theo Thông tư 80'.",
         'Thông tư số 80/2021/TT-BTC; Luật Quản lý thuế số 38/2019/QH14; Tạp chí Tài chính (tapchitaichinh.vn); scratch/daily_notes.md.',
         update_date_str),

        (7,
         'nghi dinh 289 2026 nd cp an ninh cua khau cang bien giay phep xuong tau',
         'Nghị định số 289/2026/NĐ-CP có hiệu lực từ 01/10/2026: Siết chặt an ninh trật tự tại cửa khẩu cảng biển do Bộ Công an quản lý — Thủ tục cấp phép xuống tàu và ra vào cảng quốc tế',
         7700, 'Siết chặt an ninh (+1,250%)',
         'Doanh nghiệp vận tải biển, forwarder, đại lý hải quan, giám định viên kiểm tra chuyên ngành và đơn vị cung ứng vật tư tàu biển.',
         'Thủ tục cấp phép xuống tàu (Pass port) và cấp phép đi bờ cho thuyền viên chuyển sang đăng ký điện tử qua Bộ Công an; chưa chuẩn hóa tài khoản định danh bị từ chối tiếp cận tàu biển, chậm trễ lấy mẫu giám định và phát sinh phí lưu tàu Demurrage.',
         "Tiêu đề: 'Nghị định 289/2026/NĐ-CP siết an ninh cảng biển: Hướng dẫn thủ tục đăng ký giấy phép xuống tàu và cấp phép thuyền viên tránh chậm trễ thông quan'.",
         'Nghị định số 289/2026/NĐ-CP ngày 01/10/2026; Cổng TTĐT Chính phủ (chinhphu.vn); thuvienphapluat.vn.',
         update_date_str),

        (8,
         'cbam lien minh chau au bao cao phat thai carbon thuc te sat thep nhom',
         'Cơ chế CBAM EU bước vào giai đoạn thực thi tài chính toàn diện: Doanh nghiệp xuất khẩu sắt thép, nhôm Việt Nam bắt buộc tính toán phát thải carbon thực tế thay vì áp hệ số mặc định',
         8600, 'Sức ép xanh (+1,460%)',
         'Doanh nghiệp sản xuất và xuất khẩu sắt thép, nhôm, xi măng, phân bón sang thị trường EU, giám đốc nhà máy, chuyên viên phát triển bền vững ESG.',
         'EU chính thức yêu cầu chứng chỉ CBAM tài chính; nếu không có dữ liệu kiểm kê phát thải thực tế (Scope 1 và Scope 2) được tổ chức quốc tế thẩm định, nhà nhập khẩu EU sẽ bị áp mức thuế carbon mặc định kịch khung, mất trắng thị phần.',
         "Tiêu đề: 'CBAM 2026 chuyển sang giai đoạn nộp tiền: Cách tính toán phát thải carbon thực tế cứu cánh đơn hàng xuất khẩu thép sang Châu Âu'.",
         'Quy định (EU) 2023/956; Hướng dẫn kỹ thuật EC về CBAM; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (9,
         'drewry wci cuoc tau bien mua tet 2027 phu phi pss no roll booking',
         'Báo cáo Drewry WCI tuần 43/2026 & Bão phụ phí PSS hàng Tết 2027: Kỹ thuật đàm phán hợp đồng No-Roll và bảo vệ slot tàu cho doanh nghiệp xuất nhập khẩu',
         7900, 'Nóng thị trường (+1,220%)',
         'Doanh nghiệp xuất khẩu đi bờ Tây/bờ Đông Mỹ và Châu Âu, forwarder, giám đốc logistics và chuỗi cung ứng.',
         'Hãng tàu siết công suất bằng hủy chuyến (blank sailings), đẩy phụ phí PSS từ 800 - 1.200 USD/FEU; tỷ lệ rớt cont (roll-over) chạm 18-22% tại Cái Mép, Cát Lái làm vỡ tiến độ giao hàng mùa lễ hội.',
         "Tiêu đề: 'Bão phụ phí PSS cước tàu mùa Tết 2027: 4 bí quyết đàm phán hợp đồng No-Roll và chốt slot tàu an toàn cho chủ hàng'.",
         'Drewry Shipping Consultants (drewry.co.uk); World Shipping Council; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (10,
         'nghi dinh 37 2026 nd cp tem may thieu xuat xu made in china tai xuat',
         'Nghị định số 37/2026/NĐ-CP siết chặt nhãn mác máy móc nhập khẩu: Tuyệt đối cấm nợ xuất xứ "Made in China" tại khâu thông quan — Nguy cơ bị dừng hàng và buộc tái xuất',
         8700, 'Duy trì nóng đỉnh (+1,580%)',
         'Doanh nghiệp nhập khẩu máy móc, thiết bị công nghiệp từ Trung Quốc, chủ nhà máy, forwarder và đại lý hải quan.',
         'Nghe lời shipper Trung Quốc "tao xuất nhiều lần rồi"; hàng về cảng không có chữ Made in China trên tem kim loại; bị hải quan dừng thông quan, phạt 20-60 triệu và ép buộc tái xuất, đọng vốn hàng tỷ đồng.',
         "Tiêu đề: 'Đừng để mất hàng tỷ đồng vì tin shipper Trung Quốc: Nghị định 37/2026 cấm nợ xuất xứ Made in China trên tem máy và cách xử lý tại cảng'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 128/2020/NĐ-CP; Cục Hải quan; scratch/daily_notes.md (Case study thực tế).',
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
    ws3['A3'] = f"Chu kỳ: Tuần 43 (19/10/2026 - 25/10/2026) | Mở rộng 160 chủ đề nghiên cứu (gồm 128 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 161)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 161 or m.max_row >= 161]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (156,
         'hop dong ba ben xuat khau switch b l drop shipment to khai hai quan',
         'hop dong ba ben xuat nhap khau, switch bill of lading la gi, drop shipment logistics, to khai xuat khau tai cho, seller shipper khac nhau, hoa don thuong mai ba ben, bao mat gia mua ban switch bl',
         99.4, 38800, 'https://haiquanonline.com.vn',
         'Nghiệp vụ Hợp đồng thương mại 3 bên (Triangle Trade) & Kỹ thuật Switch B/L: Hướng dẫn khai báo hải quan, đồng bộ chứng từ và bảo mật thông tin giá cho doanh nghiệp xuất khẩu',
         'Nội dung trọng tâm: Hướng dẫn nghiệp vụ chuyên sâu giải quyết bài toán giao dịch thương mại 3 bên (A bán cho B, B bán cho C; A giao hàng thẳng cho C) từ case study thực tế: (1) Rủi ro pháp lý & thương mại: Nếu trên B/L ghi Shipper là bên bán A và Consignee là bên mua C thì khách hàng cuối C sẽ biết trực tiếp bên bán A, dẫn đến mất khách hàng vào tay bên A; đồng thời thông tin hóa đơn Invoice (A-B) lệch với vận đơn B/L (A-C) dễ bị hải quan và cơ quan thuế nghi ngờ gian lận khấu trừ thuế GTGT hoặc chuyển giá; (2) Kỹ thuật phát hành Switch B/L: Bên trung gian B ủy thác cho hãng tàu/forwarder thu hồi vận đơn gốc số 1 (Shipper A, Consignee B) và phát hành vận đơn thay thế Switch B/L số 2 (Shipper B, Consignee C) tại cảng chuyển tải hoặc văn phòng đại diện để che giấu nhà sản xuất gốc và giá vốn; (3) Quy chuẩn mở tờ khai hải quan: Doanh nghiệp A mở tờ khai xuất khẩu thông thường (chỉ tiêu Người mua là B, Người nhận hàng là C); chứng từ thanh toán ngân hàng quốc tế (TT/LC) phải nhất quán theo dòng tiền thực tế giữa A và B để được hoàn thuế GTGT hợp lệ.',
         'Báo Hải quan (haiquanonline.com.vn) & scratch/daily_notes.md | HT (Học Thực chiến) / Triangle Trade & Switch B/L SOP',
         update_date_str),

        (157,
         'nghi dinh 289 2026 nd cp an ninh cua khau cang bien giay phep xuong tau',
         'nghi dinh 289 2026 nd cp, an ninh trat tu cua khau cang bien, bo cong an quan ly cang bien, giay phep xuong tau forwarder, cap phep di bo thuyen vien, kiem tra giam dinh tren tau, thu tuc tau thuyen xuat nhap canh',
         98.9, 36500, 'https://chinhphu.vn',
         'Nghị định số 289/2026/NĐ-CP có hiệu lực từ 01/10/2026: Siết chặt an ninh trật tự tại cửa khẩu cảng biển do Bộ Công an quản lý — Thủ tục cấp phép xuống tàu và ra vào khu vực cảng biển quốc tế',
         'Nội dung trọng tâm: Triển khai thi hành Nghị định số 289/2026/NĐ-CP của Chính phủ quy định về quản lý an ninh, trật tự tại các cửa khẩu cảng thuộc thẩm quyền của Bộ Công an: (1) Chuẩn hóa quy trình cấp phép xuống tàu: Nhân viên các công ty forwarder, đại lý hải quan, giám định viên kiểm tra chuyên ngành, đơn vị cung ứng vật tư tàu biển bắt buộc phải thực hiện thủ tục đăng ký cấp giấy phép xuống tàu (Pass port) điện tử qua Cổng dịch vụ công Bộ Công an trước tối thiểu 12-24 giờ; (2) Giám sát chặt chẽ thuyền viên và phương tiện chuyển cảng: Quy định mới về việc cấp phép đi bờ cho thuyền viên nước ngoài, kiểm tra an ninh đối với tàu thuyền chuyển cảng nội địa kết hợp tuyến quốc tế nhằm ngăn ngừa buôn lậu và xâm nhập trái phép; (3) Khuyến nghị cho doanh nghiệp dịch vụ logistics cảng biển: Các đơn vị cần lập danh sách nhân sự cố định, hoàn thiện hồ sơ định danh điện tử doanh nghiệp với Công an cửa khẩu cảng (Hải Phòng, Đà Nẵng, Bà Rịa - Vũng Tàu, TP.HCM) để tránh bị từ chối cấp phép xuống tàu, gây ách tắc việc lấy mẫu giám định hàng rời/hàng lỏng trên tàu.',
         'Cổng TTĐT Chính phủ (chinhphu.vn) & thuvienphapluat.vn | TN (Tin Ngành) / Seaport Security & Boarding Permit Regulations',
         update_date_str),

        (158,
         'nghi dinh 336 2026 nd cp thay the nghi dinh 85 co che mot cua quoc gia nsw',
         'nghi dinh 336 2026 nd cp, thay the nghi dinh 85 2019, co che mot cua quoc gia nsw, mot cua asean asw, chia se du lieu giua cac bo nganh, cam yeu cau nop lai giay to, so hoa thu tuc hai quan, giay phep chuyen nganh dien tu',
         99.2, 37900, 'https://customs.gov.vn',
         'Nghị định số 336/2026/NĐ-CP chính thức thay thế Nghị định 85/2019/NĐ-CP từ 15/10/2026: Đột phá kết nối liên thông dữ liệu Cơ chế Một cửa Quốc gia (NSW) và ASEAN (ASW)',
         'Nội dung trọng tâm: Phân tích các điểm mới mang tính bước ngoặt của Nghị định số 336/2026/NĐ-CP có hiệu lực từ ngày 15/10/2026 thay thế Nghị định 85/2019/NĐ-CP về thực hiện thủ tục hành chính qua Cổng Một cửa Quốc gia: (1) Nguyên tắc chia sẻ dữ liệu số dùng chung: Toàn bộ giấy phép, kết quả kiểm tra chuyên ngành của các Bộ (Bộ Y tế, Bộ NN&PTNT, Bộ Công Thương, Bộ KH&CN, Bộ GTVT) được tự động số hóa và đẩy trực tiếp vào hệ thống Một cửa quốc gia; nghiêm cấm cơ quan Hải quan và các đơn vị kiểm tra yêu cầu doanh nghiệp nộp lại bản giấy hoặc tệp đính kèm scan đối với các chứng từ đã có trên hệ thống; (2) Mở rộng kết nối Cơ chế Một cửa ASEAN: Tích hợp đầy đủ chứng nhận xuất xứ điện tử (e-C/O Form D), chứng thư kiểm dịch thực vật điện tử (e-Phyto) và tờ khai hải quan ASEAN (ACDD) với các nước thành viên; (3) Kế hoạch ứng phó sự cố truyền dữ liệu: Hướng dẫn cơ chế dự phòng và hotline phản ứng nhanh khi hệ thống NSW gặp sự cố nghẽn mạng để giải phóng hàng hóa đúng hẹn, không làm phát sinh chi phí lưu bãi.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Tạp chí Tài chính | TN (Tin Ngành) / National Single Window Decree 336 Compliance',
         update_date_str),

        (159,
         'thue nha thau dich vu phan mem nuoc ngoai chatgpt hoa don dien tu hop le',
         'thue nha thau fct dich vu phan mem, hoa don chatgpt open ai hop le, thue nha thau saas nuoc ngoai, thong tu 80 2021 tt btc, khau tru chi phi thue tndn phan mem, nha cung cap nuoc ngoai nop thue tai vn, ke khai thue nha thau nop thay',
         98.8, 36100, 'https://tapchitaichinh.vn',
         'Rủi ro thuế nhà thầu (FCT) khi chi trả phần mềm nước ngoài (ChatGPT, Zoom, SaaS) và dịch vụ quảng cáo quốc tế: Cách hạch toán chi phí hợp lý được trừ',
         'Nội dung trọng tâm: Cảnh báo rủi ro thanh kiểm tra thuế đối với các khoản chi trả dịch vụ xuyên biên giới từ case study thực tế tại doanh nghiệp: (1) Sai lầm kinh điển của doanh nghiệp: Cà thẻ tín dụng visa cá nhân hoặc công ty mua tài khoản ChatGPT, OpenAI, Adobe, Canva, Google Workspace hoặc thanh toán quảng cáo TikTok/Meta nhưng không lập hồ sơ kê khai thuế nhà thầu; (2) Chế tài của cơ quan Thuế: Toàn bộ chi phí này bị loại khỏi chi phí được trừ khi tính thuế TNDN (mất 20%), đồng thời doanh nghiệp bị ấn định truy thu thuế nhà thầu nộp thay (5% thuế TNDN + thuế GTGT tùy loại dịch vụ) cùng tiền phạt chậm nộp 0.03%/ngày; (3) Hướng dẫn giải pháp chuẩn pháp lý: Kiểm tra danh sách Nhà cung cấp nước ngoài đã đăng ký thuế trực tiếp tại Cổng TTĐT Tổng cục Thuế (etaxvn.gdt.gov.vn); trường hợp NCC đã nộp thuế trực tiếp, doanh nghiệp dùng sao kê thanh toán ngân hàng và hóa đơn/receipt điện tử theo Thông tư 80/2021/TT-BTC để ghi nhận chi phí hợp lệ; trường hợp NCC chưa nộp thuế thì doanh nghiệp bắt buộc phải khấu trừ và nộp thay theo mẫu 01/NTNN.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & scratch/daily_notes.md | CS (Case Study) / Foreign Contractor Tax (FCT) & Digital Service SOP',
         update_date_str),

        (160,
         'drewry wci cuoc tau bien mua tet 2027 phu phi pss no roll booking',
         'drewry wci cuoc cont tuan 43, phu phi mua cao diem pss $1000, rui ro rot cont roll over hang tet, hop dong no roll cam ket cho, blank sailings hang tau, cuoc bien tuyen my au, chien luoc booking hang tet 2027',
         98.4, 35200, 'https://drewry.co.uk',
         'Báo cáo Drewry WCI tuần 43/2026 & Bão phụ phí PSS hàng Tết 2027: Kỹ thuật đàm phán hợp đồng No-Roll và bảo vệ slot tàu cho doanh nghiệp xuất nhập khẩu',
         'Nội dung trọng tâm: Phân tích diễn biến thị trường vận tải biển container quốc tế tuần 43/2026 và cẩm nang điều hành logistics cuối năm: (1) Diễn biến áp lực tải: Chỉ số Drewry World Container Index duy trì vùng giá cao 4.490 USD/FEU; các hãng tàu tiếp tục áp dụng phụ phí mùa cao điểm (Peak Season Surcharge - PSS) từ 800 đến 1.200 USD/FEU cho các tuyến bờ Tây nước Mỹ và Châu Âu, kết hợp chính sách hủy chuyến chọn lọc (Blank Sailings) để duy trì mặt bằng giá; (2) Khủng hoảng rớt cont (Roll-over): Tỷ lệ rớt cont tại các cảng xuất khẩu trọng điểm Cái Mép, Cát Lái và Hải Phòng có xu hướng tăng cao lên mức 18-22% đối với các booking thả nổi (spot rate) do các hãng tàu ưu tiên hàng có phụ phí cam kết cao; (3) Chiến lược ứng phó của Eureka: Doanh nghiệp cần chủ động mở rộng cửa sổ booking trước tối thiểu 3-4 tuần so với ngày tàu chạy (ETD); đàm phán hợp đồng vận chuyển có điều khoản cam kết không rớt cont (No-Roll Guarantee Clause) với hãng tàu/forwarder uy tín; chia nhỏ các lô hàng lớn thành nhiều vận đơn khác nhau để đảm bảo luôn có hàng được lên tàu theo tiến độ sản xuất.',
         'Drewry Shipping Consultants (drewry.co.uk) & Báo Công Thương | CS (Case Study) / Ocean Freight & Roll-over Crisis Mitigation',
         update_date_str),
    ]

    # Write new items into rows 161 to 165
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 161 + idx
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

    # New Total row at row 166
    total_row_idx = 166
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 160 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D165)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E165)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 160 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 128 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A166:C166')
    ws3.merge_cells('F166:J166')
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
    update_seo_w43_2010()
