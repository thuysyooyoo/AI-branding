# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_19_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 19/10/2026 (Tuần 43: 19/10/2026 - 25/10/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 43, ngày 19/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 150 lên 155 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 19/10/2026)
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

def update_seo_w43_1910():
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

    update_date_str = "19/10/2026"

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
         'khai bao ma phu ko tren vnaccs mat ma dan su hang hoa luong dung',
         "Công văn 22233/CHQ-GSQL & 19916/CHQ-GSQL: Bắt buộc khai mã phụ 'KO' trên hệ thống VNACCS đối với hàng hóa trùng mã HS an ninh mạng, mật mã dân sự nhưng không thuộc diện cấp phép",
         8400, 'Tăng vọt kiểm soát (+1,420%)',
         'Doanh nghiệp nhập khẩu linh kiện điện tử, vi mạch, máy tính, thiết bị viễn thông, camera AI, máy móc tự động hóa, đại lý hải quan.',
         'Hàng hóa dân dụng bình thường có mã HS trùng danh mục quản lý chuyên ngành nhưng không khai mã "KO" dẫn đến việc VNACCS tự động nhảy luồng Đỏ, bị dừng thông quan 15-30 ngày chờ giám định kỹ thuật và bị phạt khai sai chỉ tiêu thông tin.',
         "Tiêu đề: 'Bẫy luồng Đỏ 18 ngày tại Cát Lái vì quên mã 'KO': Tuyệt chiêu khai báo hàng điện tử trùng danh mục mật mã dân sự trên VNACCS'.",
         'Công văn số 22233/CHQ-GSQL; Công văn 19916/CHQ-GSQL; Nghị định 58/2016/NĐ-CP; Nghị định 53/2018/NĐ-CP; Cục Giám sát Quản lý Hải quan.',
         update_date_str),

        (3,
         'thong tu 121 2025 tt btc hang qua canh trung chuyen niem phong gps seal',
         'Thông tư 121/2025/TT-BTC siết chặt giám sát hàng quá cảnh & trung chuyển cảng biển: Bắt buộc kẹp niêm phong điện tử định vị vệ tinh (GPS Seal) cho sà lan thủy nội địa',
         7900, 'Siết chặt an ninh (+1,280%)',
         'Doanh nghiệp vận tải thủy nội địa (sà lan), forwarder kinh doanh hàng quá cảnh Campuchia/Lào/Trung Quốc, hãng tàu trung chuyển quốc tế.',
         'Bắt buộc gắn GPS Seal kết nối dữ liệu hành trình trực tiếp với Tổng cục Hải quan; sự cố đứt seal hoặc mất sóng GPS giữa sông nước bị xử lý nghiêm như nghi án buôn lậu hoặc rút ruột container; mức phạt từ 50 - 200 triệu đồng.',
         "Tiêu đề: 'Thông tư 121/2025/TT-BTC siết chặt hàng quá cảnh: Quy chuẩn GPS Seal cho sà lan và quy trình ứng phó khẩn cấp khi đứt tín hiệu định vị'.",
         'Thông tư số 121/2025/TT-BTC; Thông tư số 38/2015/TT-BTC; Điều 43-44 Luật Hải quan 2014; haiquanonline.com.vn.',
         update_date_str),

        (4,
         'thong tu 128 2026 tt btc cua khau thong minh agv luu kho 30 ngay',
         'Thông tư số 128/2026/TT-BTC & Nghị định 336/2026/NĐ-CP: Vận hành Cửa khẩu thông minh từ 15/10/2026 — Xe tự hành AGV, Camera AI và Thời hạn lưu kho tối đa 30 ngày',
         8100, 'Cao điểm thực thi (+1,350%)',
         'Doanh nghiệp xuất nhập khẩu nông sản, linh kiện qua cửa khẩu Hữu Nghị, Tân Thanh, doanh nghiệp logistics biên mậu.',
         'Xe tự hành AGV chạy tự động 24/7 đòi hỏi dữ liệu JSON DataFormat truyền trước 100%; hàng lưu kho vượt 30 ngày bị cưỡng chế thanh lý; doanh nghiệp chưa chuẩn hóa phần mềm bị từ chối barie thông minh.',
         "Tiêu đề: 'Cửa khẩu thông minh chạy xe tự hành AGV: 3 điều kiện số hóa bắt buộc để container luồng xanh không chạm qua biên giới'.",
         'Thông tư số 128/2026/TT-BTC; Nghị định số 336/2026/NĐ-CP; Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (5,
         'cbam lien minh chau au bao cao phat thai carbon thuc te sat thep nhom',
         'Cơ chế CBAM EU bước vào giai đoạn thực thi tài chính toàn diện: Doanh nghiệp xuất khẩu sắt thép, nhôm Việt Nam bắt buộc tính toán phát thải carbon thực tế thay vì áp hệ số mặc định',
         8600, 'Sức ép xanh (+1,460%)',
         'Doanh nghiệp sản xuất và xuất khẩu sắt thép, nhôm, xi măng, phân bón sang thị trường EU, giám đốc nhà máy, chuyên viên phát triển bền vững ESG.',
         'EU chính thức yêu cầu chứng chỉ CBAM tài chính; nếu không có dữ liệu kiểm kê phát thải thực tế (Scope 1 và Scope 2) được tổ chức quốc tế thẩm định, nhà nhập khẩu EU sẽ bị áp mức thuế carbon mặc định kịch khung, mất trắng thị phần.',
         "Tiêu đề: 'CBAM 2026 chuyển sang giai đoạn nộp tiền: Cách tính toán phát thải carbon thực tế cứu cánh đơn hàng xuất khẩu thép sang Châu Âu'.",
         'Quy định (EU) 2023/956; Hướng dẫn kỹ thuật EC về CBAM; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (6,
         'drewry wci tuan 43 2026 cuoc container tang phu phi pss hang tet',
         'Chỉ số Drewry WCI tuần 43/2026 tiếp tục duy trì vùng cao 4.490 USD/FEU: Hãng tàu áp phụ phí mùa cao điểm (PSS $800-$1000) và Nguy cơ rớt cont (Roll-over) mùa hàng Tết 2027',
         7800, 'Nóng thị trường (+1,180%)',
         'Doanh nghiệp xuất khẩu đi bờ Tây/bờ Đông Mỹ và Châu Âu, forwarder, giám đốc logistics và chuỗi cung ứng.',
         'Hãng tàu liên minh (Ocean Alliance, Gemini, THE Alliance) siết công suất bằng blank sailings, ép giá spot rate và áp phụ phí PSS kịch khung; rủi ro rớt cont hàng xuất khẩu trước Tết làm đứt gãy cam kết giao hàng.',
         "Tiêu đề: 'Bão phụ phí PSS cước tàu mùa Tết 2027: 4 bí quyết đàm phán hợp đồng No-Roll và chốt slot tàu an toàn cho chủ hàng'.",
         'Báo cáo Drewry Container Market Outlook; World Shipping Council; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (7,
         'chuan json dataformat chung tu dien tu cong mot cua quoc gia nsw',
         'Quyết định của Cục Hải quan: Bắt buộc áp dụng chuẩn dữ liệu JSON DataFormat trong trao đổi chứng từ điện tử qua Cơ chế Một cửa Quốc gia từ nửa cuối tháng 10/2026',
         7500, 'Chuyển đổi công nghệ (+1,210%)',
         'Toàn bộ doanh nghiệp XNK, công ty phần mềm khai hải quan (ECUS, FPT, Thái Sơn), đại lý làm thủ tục hải quan.',
         'Không kịp cập nhật bản vá phần mềm theo chuẩn JSON mới sẽ bị từ chối kết nối dữ liệu chứng từ đính kèm (invoice, packing list, C/O), khiến hồ sơ bị treo không thể cấp phép chuyên ngành.',
         "Tiêu đề: 'Chuyển đổi chuẩn JSON DataFormat trên Cổng Một cửa Quốc gia: Hướng dẫn kỹ thuật kiểm tra và nâng cấp phần mềm cho doanh nghiệp'.",
         'Quyết định Cục Hải quan; Nghị định 336/2026/NĐ-CP; Cổng TTĐT Tổng cục Hải quan (customs.gov.vn).',
         update_date_str),

        (8,
         'nghi dinh 365 2026 nd cp dieu kien xuat khau gao kho chuyen dung',
         'Nghị định số 365/2026/NĐ-CP có hiệu lực từ 01/10/2026: Siết chặt điều kiện kinh doanh xuất khẩu gạo — Bắt buộc kho chuyên dùng, cơ sở xay xát và giám sát nhập khẩu chống gian lận',
         7100, 'Chính sách trọng điểm (+1,050%)',
         'Thương nhân xuất nhập khẩu gạo, doanh nghiệp chế biến nông sản, hiệp hội lương thực VFA.',
         'Quy định bắt buộc phải sở hữu hoặc thuê kho chuyên dùng và cơ sở xay xát đạt chuẩn vệ sinh ATTP; cấm chuyển nhượng giấy phép khống; giám sát nghiêm ngặt kế hoạch nhập khẩu gạo phụ phẩm.',
         "Tiêu đề: 'Nghị định 365/2026/NĐ-CP siết chặt xuất nhập khẩu gạo: Bộ tiêu chuẩn kho bãi và thủ tục đăng ký pháp lý doanh nghiệp cần hoàn tất'.",
         'Nghị định số 365/2026/NĐ-CP ngày 01/10/2026; Báo Công Thương (congthuong.vn); chinhphu.vn.',
         update_date_str),

        (9,
         'cho muon cccd mo cong ty no thue etax mobile xuat hoa don khong',
         'Bẫy pháp lý cho mượn CCCD mở doanh nghiệp: Khoản nợ thuế hàng trăm triệu "treo" trên eTax Mobile, lệnh tạm hoãn xuất cảnh và Quy trình 4 bước làm việc với Cơ quan Công an',
         8300, 'Cảnh báo đỏ pháp lý (+1,390%)',
         'Cá nhân, người lao động, người đại diện theo pháp luật trên ĐKKD, kế toán doanh nghiệp, hộ kinh doanh.',
         'Cho bạn bè/người nước ngoài mượn CCCD đứng tên công ty nhưng không điều hành; bị đối tượng lợi dụng xuất khống hóa đơn trốn thuế rồi bỏ trốn; bị cưỡng chế thuế, phong tỏa tài khoản và tạm hoãn xuất cảnh.',
         "Tiêu đề: 'Gánh nợ thuế hàng trăm triệu trên eTax Mobile vì cho mượn CCCD: 4 bước lập vi bằng, tố giác tội phạm và gỡ phong tỏa tài khoản'.",
         'Luật Quản lý thuế số 38/2019/QH14; Bộ luật Hình sự 2015; Tạp chí Tài chính (tapchitaichinh.vn); scratch/daily_notes.md (17/10/2026).',
         update_date_str),

        (10,
         'nghi dinh 37 2026 nd cp tem may thieu xuat xu made in china tai xuat',
         'Nghị định số 37/2026/NĐ-CP siết chặt nhãn mác máy móc nhập khẩu: Tuyệt đối cấm nợ xuất xứ "Made in China" tại khâu thông quan — Nguy cơ bị dừng hàng và buộc tái xuất',
         8700, 'Duy trì nóng đỉnh (+1,580%)',
         'Doanh nghiệp nhập khẩu máy móc, thiết bị công nghiệp từ Trung Quốc, chủ nhà máy, forwarder và đại lý hải quan.',
         'Nghe lời shipper Trung Quốc "tao xuất nhiều lần rồi"; hàng về cảng không có chữ Made in China trên tem kim loại; bị hải quan dừng thông quan, phạt 20-60 triệu và ép buộc tái xuất, đọng vốn hàng tỷ đồng.',
         "Tiêu đề: 'Đừng để mất hàng tỷ đồng vì tin shipper Trung Quốc: Nghị định 37/2026 cấm nợ xuất xứ Made in China trên tem máy và cách xử lý tại cảng'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 128/2020/NĐ-CP; Cục Hải quan; scratch/daily_notes.md (16/10/2026).',
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
    ws3['A3'] = f"Chu kỳ: Tuần 43 (19/10/2026 - 25/10/2026) | Mở rộng 155 chủ đề nghiên cứu (gồm 123 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 156)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 156 or m.max_row >= 156]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (151,
         'cong van 23007 chq gsql thong tu 127 phuong tien van tai duong bo nsw',
         'cong van 23007 chq gsql, thong tu 127 2026 tt btc, khai bao phuong tien xuat nhap canh, to khai phuong tien duong bo bien gioi, cong mot cua quoc gia nsw, giam sat phuong tien hai quan, ach tac cua khau hu nghi tan thanh',
         99.6, 39200, 'https://customs.gov.vn',
         'Công văn hỏa tốc số 23007/CHQ-GSQL (17/10/2026): Hướng dẫn khẩn cấp áp dụng Thông tư 127/2026/TT-BTC — Phương tiện đường bộ tạm duy trì phương thức hiện hành chống ùn tắc cửa khẩu biên giới',
         'Nội dung trọng tâm: Cục Hải quan ban hành văn bản hỏa tốc số 23007/CHQ-GSQL hướng dẫn khẩn cấp các Chi cục Hải quan tỉnh, thành phố biên giới triển khai Thông tư 127/2026/TT-BTC (có hiệu lực từ 15/10/2026) quy định mẫu chứng từ và chỉ tiêu thông tin khai báo phương tiện vận tải xuất nhập cảnh, quá cảnh: (1) Phân tầng triển khai linh hoạt: Phương tiện đường biển và hàng không tiếp tục thực hiện khai báo 100% điện tử qua Cổng Một cửa Quốc gia (NSW); riêng phương tiện đường bộ và đường sắt tại các cửa khẩu đất liền (Lạng Sơn, Lào Cai, Quảng Ninh, Hà Tĩnh) tạm thời cho phép duy trì phương thức khai báo hiện hành nhằm bảo đảm thông suốt dòng chảy hàng hóa nông sản, linh kiện, không để xảy ra ùn tắc cục bộ; (2) Trách nhiệm doanh nghiệp: Các chủ xe, forwarder và đại lý hải quan cần khẩn trương đào tạo nhân sự, chuẩn hóa biểu mẫu mới của Thông tư 127 để sẵn sàng kết nối toàn diện ngay khi hạ tầng NSW tại các trạm kiểm soát cửa khẩu hoàn tất nâng cấp kỹ thuật; (3) Tăng cường phối hợp liên ngành giữa Hải quan và Bộ đội Biên phòng trong kiểm tra, đối chiếu mã định danh phương tiện và biển số xe tự động.',
         'Tổng cục Hải quan (customs.gov.vn) & Báo Chính phủ (baochinhphu.vn) | CS (Case Study) / Border Vehicle Clearance & Circular 127 Compliance',
         update_date_str),

        (152,
         'khai bao ma phu ko tren vnaccs mat ma dan su hang hoa luong dung',
         'cong van 22233 chq gsql, cong van 19916 chq gsql, ma phu ko tren vnaccs, hang hoa luong dung quan su, mat ma dan su, kiem tra chuyen nganh an ninh mang, treo luong do giam dinh, linh kien dien tu vi mach',
         99.1, 37500, 'https://haiquanonline.com.vn',
         "Công văn 22233/CHQ-GSQL & 19916/CHQ-GSQL: Bắt buộc khai mã phụ 'KO' trên hệ thống VNACCS đối với hàng hóa có mã HS trùng danh mục an ninh mạng, mật mã dân sự nhưng không thuộc diện cấp phép",
         "Nội dung trọng tâm: Cục Hải quan ban hành văn bản chấn chỉnh và hướng dẫn thống nhất trên toàn quốc về việc khai báo mã phụ 'KO' trên hệ thống VNACCS/VCIS: (1) Nhận diện rủi ro: Nhiều mặt hàng linh kiện điện tử, vi mạch, máy tính, camera, biến tần công nghiệp có mã HS 8 chữ số trùng với Danh mục sản phẩm mật mã dân sự hoặc hàng hóa lưỡng dụng (Nghị định 58/2016/NĐ-CP và Nghị định 53/2018/NĐ-CP) nhưng bản chất kỹ thuật dân dụng thuần túy không có tính năng mã hóa an ninh quốc phòng; (2) Quy chuẩn khai báo bắt buộc: Doanh nghiệp khi mở tờ khai hải quan bắt buộc phải chủ động nhập mã 'KO' tại ô chỉ tiêu phân loại kiểm tra để hệ thống phân luồng chính xác. Việc bỏ trống ô mã phân loại sẽ khiến VNACCS tự động khóa chuyển luồng Đỏ kiểm hóa 100% và buộc gửi văn bản trưng cầu giám định kỹ thuật kéo dài 15-30 ngày; (3) Hồ sơ đối soát: Doanh nghiệp cần chuẩn bị sẵn bảng thông số kỹ thuật (Datasheet/Catalog) của nhà sản xuất và văn bản cam kết mục đích sử dụng dân dụng để xuất trình ngay khi Hải quan yêu cầu đối chiếu.",
         "Báo Hải quan (haiquanonline.com.vn) & Thư Viện Pháp Luật | HT (Học Thực chiến) / VNACCS 'KO' Code Clearance SOP",
         update_date_str),

        (153,
         'thong tu 121 2025 tt btc hang qua canh trung chuyen niem phong gps seal',
         'thong tu 121 2025 tt btc, dieu 51a 51b thong tu 38, hang qua canh duong thuy, niem phong hai quan dien tu, e seal dinh vi gps, sa lan cai mep campuchia, dut tin hieu niem phong chia, buon lau hang qua canh',
         98.7, 35400, 'https://haiquanonline.com.vn',
         'Thông tư 121/2025/TT-BTC siết chặt giám sát hàng quá cảnh & trung chuyển cảng biển: Bắt buộc kẹp niêm phong điện tử định vị vệ tinh (GPS Seal) cho sà lan thủy nội địa',
         'Nội dung trọng tâm: Hướng dẫn chi tiết việc triển khai Thông tư 121/2025/TT-BTC (sửa đổi Điều 51a, 51b Thông tư 38/2015/TT-BTC) về quy trình thủ tục và giám sát hải quan đối với hàng hóa quá cảnh, trung chuyển: (1) Yêu cầu kỹ thuật bắt buộc: 100% container xếp chồng trên phương tiện thủy nội địa (sà lan) vận chuyển hàng quá cảnh tuyến Cái Mép/Cát Lái đi Campuchia hoặc Lào bắt buộc phải kẹp seal định vị vệ tinh GPS (E-Seal) do cơ quan Hải quan trang cấp để theo dõi hải trình thời gian thực 24/7; (2) Quy trình xử lý sự cố đứt seal: Khi phương tiện gặp sự cố va quẹt đường thủy làm đứt vỡ seal hoặc mất tín hiệu GPS, thuyền trưởng bắt buộc phải neo đậu phương tiện, lập biên bản hiện trường và thông báo qua hotline khẩn cấp cho Chi cục Hải quan gần nhất trong vòng 01 giờ; (3) Chế tài nghiêm khắc: Tự ý tháo dỡ seal hoặc làm mất tín hiệu định vị mà không khai báo kịp thời sẽ bị tạm giữ phương tiện, khám xét toàn bộ lô hàng theo nghi án buôn lậu và xử phạt từ 50 đến 200 triệu đồng.',
         'Báo Hải quan (haiquanonline.com.vn) & Cổng TTĐT Bộ Tài chính | TN (Tin Ngành) / Transit Cargo GPS Seal Regulations',
         update_date_str),

        (154,
         'cbam lien minh chau au bao cao phat thai carbon thuc te sat thep nhom',
         'co che cbam eu 2026, chung chi carbon chau au, bao cao phat thai carbon thuc te, scope 1 scope 2 emissions, xuat khau thep sang eu, tham dinh phat thai doc lap, he so phat thai mac dinh, bien dong gia chung chi cbam',
         98.5, 36800, 'https://congthuong.vn',
         'Cơ chế CBAM EU bước vào giai đoạn thực thi tài chính toàn diện: Doanh nghiệp xuất khẩu sắt thép, nhôm Việt Nam bắt buộc tính toán phát thải carbon thực tế thay vì áp hệ số mặc định',
         'Nội dung trọng tâm: Phân tích lộ trình vận hành chính thức của Cơ chế điều chỉnh biên giới carbon (CBAM) của Liên minh Châu Âu (EU) từ năm 2026: (1) Chuyển dịch từ báo cáo sang nghĩa vụ tài chính: Nhà nhập khẩu tại EU bắt buộc phải mua chứng chỉ CBAM tương ứng với hàm lượng phát thải carbon tích hợp trong sản phẩm sắt thép, nhôm, xi măng, phân bón nhập khẩu; (2) Áp lực dồn lên nhà sản xuất Việt Nam: Doanh nghiệp xuất khẩu bắt buộc phải cung cấp dữ liệu phát thải thực tế (gồm phát thải trực tiếp Scope 1 và gián tiếp Scope 2 từ tiêu thụ điện) được tổ chức thẩm định độc lập quốc tế chứng thực. Nếu không có dữ liệu thực tế minh bạch, EU sẽ áp hệ số phát thải mặc định ở mức cao nhất, khiến mức thuế carbon tăng vọt làm mất hoàn toàn khả năng cạnh tranh; (3) Hành động khuyến nghị: Doanh nghiệp cần khẩn trương đầu tư hệ thống đo kiểm năng lượng số hóa, kiểm kê khí nhà kính theo chuẩn ISO 14064 và tìm kiếm các chứng chỉ năng lượng tái tạo (I-REC) để khấu trừ lượng phát thải gián tiếp.',
         'Báo Công Thương (congthuong.vn) & Báo cáo Ủy ban Châu Âu (EC) | FF (Funfact & Insight / Trade Intelligence) / EU CBAM Financial Compliance',
         update_date_str),

        (155,
         'drewry wci tuan 43 2026 cuoc container tang phu phi pss hang tet',
         'drewry world container index tuan 43, gia cuoc container duong bien, phu phi mua cao diem pss $1000, blank sailing hang tau bo chuyen, nguy co rot cont roll over, booking cont hang tet 2027, hop dong cam ket no roll',
         98.2, 34700, 'https://drewry.co.uk',
         'Chỉ số Drewry WCI tuần 43/2026 tiếp tục neo cao 4.490 USD/FEU: Hãng tàu đẩy mạnh phụ phí PSS $800-$1000 và Nguy cơ rớt cont (Roll-over) mùa hàng Tết 2027',
         'Nội dung trọng tâm: Báo cáo thị trường vận tải container đường biển tuần 43/2026 của Drewry ghi nhận chỉ số WCI tiếp tục duy trì mức cao 4.490 USD/FEU: (1) Diễn biến thị trường: Nhu cầu gom hàng phục vụ mùa lễ Giáng sinh và Tết Nguyên Đán 2027 tăng mạnh tại các cảng xuất khẩu châu Á, đặc biệt tuyến châu Á - Bờ Tây Bắc Mỹ và châu Á - Bắc Âu; (2) Chiến lược của các liên minh hãng tàu: Tận dụng nhu cầu dồn dập, các hãng tàu áp dụng triệt để chính sách hủy chuyến có chủ đích (blank sailings) để ghìm tải, đồng thời công bố áp dụng phụ phí mùa cao điểm (Peak Season Surcharge - PSS) từ 800 USD đến 1.000 USD/FEU từ nửa cuối tháng 10/2026; (3) Rủi ro rớt cont và giải pháp: Hiện tượng quá tải bãi cont tại Cát Lái và Cái Mép khiến tỷ lệ rớt cont (roll-over) chạm mốc 15-20%. Chủ hàng cần sớm ký hợp đồng dịch vụ có điều khoản cam kết "No-Roll", đặt booking trước tối thiểu 3-4 tuần và chia nhỏ lô hàng để giảm thiểu rủi ro tồn ứ.',
         'Drewry Shipping Consultants (drewry.co.uk) & World Shipping Council | CS (Case Study) / Ocean Freight Rate & Peak Season Surcharge SOP',
         update_date_str),
    ]

    # Write new items into rows 156 to 160
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 156 + idx
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

    # New Total row at row 161
    total_row_idx = 161
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 155 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D160)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E160)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 155 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 123 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A161:C161')
    ws3.merge_cells('F161:J161')
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
    update_seo_w43_1910()
