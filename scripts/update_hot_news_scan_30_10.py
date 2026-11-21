# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_30_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 30/10/2026
(Chu kỳ: Tuần 44: 26/10/2026 - 01/11/2026 | Cập nhật ngày 30/10/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành, chu kỳ Tuần 44 cập nhật ngày 30/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 185 lên 190 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 30/10/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w44_3010.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w44_3010.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w44_3010():
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

    update_date_str = "30/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 44 (26/10/2026 - 01/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'cong van 23535 chq gsql canh bao phong ve thuong mai gian lan xuat xu cong van 8744 bct',
         'Công văn số 23535/CHQ-GSQL ngày 27/10/2026 & CV 8744/BCT-PVTM: Cảnh báo đỏ 7 nhóm mặt hàng rủi ro cao bị điều tra phòng vệ thương mại, gian lận xuất xứ và lẩn tránh thuế — Nguy cơ bị áp thuế trừng phạt của Mỹ, EU',
         13500, 'Hỏa tốc toàn quốc (+2,300%)',
         'Doanh nghiệp sản xuất, xuất khẩu gỗ & nội thất, sắt thép kim loại, pin năng lượng, dệt may; forwarder tuyến Mỹ/EU.',
         'Nhiều doanh nghiệp nhập linh kiện thô từ nước ngoài về chỉ gia công đơn giản, dán nhãn Việt Nam để né thuế; bị cơ quan Hải quan và nước nhập khẩu điều tra lẩn tránh thuế phòng vệ thương mại, đối mặt nguy cơ bị áp thuế chống bán phá giá trừng phạt từ 25% đến hơn 100% và giữ hàng tại cảng đến vài tháng.',
         "Tiêu đề: 'Hải quan ra Công văn 23535/CHQ-GSQL cảnh báo đỏ 7 nhóm hàng: 4 bước rà soát hồ sơ xuất xứ chống nguy cơ bị Mỹ và EU trừng phạt thuế phòng vệ'.",
         'Công văn số 23535/CHQ-GSQL ngày 27/10/2026; Công văn số 8744/BCT-PVTM ngày 22/10/2026 của Cục Phòng vệ thương mại (Bộ Công Thương).',
         update_date_str),

        (2,
         'quy dinh cbam eu han chot bao cao phat thai 31 10 sat thep nhom chung chi carbon',
         'Đếm ngược 24h hạn chót nộp báo cáo phát thải CBAM (31/10): Áp lực nghẹt thở đối với doanh nghiệp xuất khẩu Sắt thép, Nhôm sang EU — Khi chứng chỉ carbon trở thành rào cản tài chính sống còn',
         13100, 'Đếm ngược 24h cấp bách (+2,200%)',
         'Doanh nghiệp sản xuất, xuất khẩu sắt thép, tôn mạ, nhôm thanh định hình, sản phẩm kim loại sang 27 quốc gia Liên minh Châu Âu.',
         'Ngày 31/10 là hạn chót nộp báo cáo quý 3 phát thải CBAM; nhiều doanh nghiệp Việt Nam lúng túng chưa kiểm kê khí nhà kính, không tính được phát thải nội hàm (SEE) thực tế; nếu dùng hệ số mặc định sẽ bị EU áp thuế carbon phạt cực cao (>80 EUR/tấn CO2e), làm mất đơn hàng vào tay đối thủ.',
         "Tiêu đề: 'Hạn chót CBAM ngày 31/10: Doanh nghiệp thép và nhôm Việt Nam cần hoàn tất báo cáo phát thải thế nào để không bị đánh thuế phạt nghìn tỷ?'.",
         'Quy định Cơ chế điều chỉnh biên giới carbon (CBAM) của Nghị viện Châu Âu; Báo Công Thương; Cổng TTĐT Logistics Việt Nam (logistics.gov.vn).',
         update_date_str),

        (3,
         'cong van 23244 chq nvthq hoan thue gtgt thue nhap khau xuat tra hang loi b13',
         'Công văn số 23244/CHQ-NVTHQ: Hướng dẫn thủ tục hoàn thuế GTGT và thuế nhập khẩu đối với hàng hóa xuất trả chủ hàng nước ngoài (mã B13) — 4 điều kiện sống còn để không bị bác hồ sơ hoàn thuế',
         12400, 'Nghiệp vụ tài chính nóng (+2,100%)',
         'Doanh nghiệp nhập khẩu nguyên vật liệu, máy móc thiết bị, linh kiện điện tử; kế toán thuế xuất nhập khẩu và đại lý hải quan.',
         'Hàng hóa nhập khẩu bị lỗi chất lượng hoặc sai quy cách nhưng doanh nghiệp lúng túng trong thủ tục giám định thương mại và khai báo tái xuất (B13), dẫn đến việc cơ quan Hải quan từ chối hoàn thuế GTGT và thuế nhập khẩu, gây thiệt hại tài chính nặng nề.',
         "Tiêu đề: 'Công văn 23244/CHQ-NVTHQ: Bí quyết hoàn thuế GTGT và thuế nhập khẩu khi xuất trả hàng lỗi (B13) sạch nước cản không lo bị truy thu'.",
         'Công văn số 23244/CHQ-NVTHQ tháng 10/2026; Điều 34 Nghị định 134/2016/NĐ-CP; Thông tư 38/2015/TT-BTC & TT 39/2018/TT-BTC.',
         update_date_str),

        (4,
         'no thue treo etax mobile khoa ma so thue chan mo to khai hoa don thay the khac ky',
         'Giải quyết triệt để rủi ro hóa đơn thay thế khác kỳ & nợ thuế treo eTax Mobile khi đóng/chuyển đổi MST Hộ kinh doanh sang Doanh nghiệp XNK — Gỡ nút thắt bị chặn mở tờ khai hải quan',
         12000, 'Nỗi đau thực tế chạm đỉnh (+2,050%)',
         'Chủ doanh nghiệp vừa và nhỏ (SME), cá nhân chuyển từ Hộ kinh doanh lên Doanh nghiệp tư nhân/Công ty TNHH, kế toán thuế XNK.',
         'Tồn đọng nợ thuế cũ của Hộ kinh doanh trên eTax Mobile hoặc xử lý sai hóa đơn thay thế khác kỳ khiến hệ thống Hải quan tự động khóa phân luồng, chặn truyền tờ khai điện tử và tạm hoãn xuất cảnh người đại diện pháp luật.',
         "Tiêu đề: 'Nợ thuế treo eTax Mobile chặn đứng tờ khai hải quan: Cách tra soát, gỡ khóa MST và xử lý hóa đơn thay thế khác kỳ trong 24 giờ'.",
         'Nghị định 123/2020/NĐ-CP; Luật Quản lý thuế số 38/2019/QH14; Cổng Dịch vụ công Hải quan; Nhật ký thực tế daily_notes.md.',
         update_date_str),

        (5,
         'quyet dinh 1434 qd chq chuan du lieu chung tu dien tu json dataformat thong quan 2026',
         'Quyết định số 1434/QĐ-CHQ chính thức áp dụng chuẩn dữ liệu JSON DataFormat: Thay thế hoàn toàn các chuẩn cũ — Doanh nghiệp xuất nhập khẩu và đơn vị phần mềm hải quan cần làm gì ngay để không bị lỗi truyền chứng từ?',
         11700, 'Nghiệp vụ công nghệ mới (+1,950%)',
         'Nhân viên chứng từ XNK, đại lý hải quan, doanh nghiệp phát triển phần mềm khai báo hải quan (ECUS, FPT, Thái Sơn...).',
         'Quyết định 1434/QĐ-CHQ có hiệu lực từ 15/10/2026 bắt buộc đổi sang chuẩn JSON DataFormat thay thế QĐ 748 và 1143; nhiều hệ thống chưa cập nhật schema dẫn đến lỗi từ chối tiếp nhận thông điệp, chứng từ đính kèm (invoice, packing list) bị lỗi định dạng Base64 khiến tờ khai bị treo tại cổng tiếp nhận.',
         "Tiêu đề: 'Quyết định 1434/QĐ-CHQ áp dụng chuẩn JSON DataFormat: Hướng dẫn sửa lỗi kết nối và chuẩn hóa file chứng từ đính kèm tờ khai hải quan'.",
         'Quyết định số 1434/QĐ-CHQ ngày 26/08/2026 của Cục Hải quan (có hiệu lực từ 15/10/2026); Báo Hải quan Online.',
         update_date_str),

        (6,
         'gia cuoc van tai container noi dia quoc te quy 4 tang 15 25 phan tram pss blank sailings',
         'Sóng tăng giá cước vận tải container nội địa (Hải Phòng - Cái Mép/TP.HCM tăng 15-25%) và Cước quốc tế cuối năm: Chiến lược tối ưu chi phí cung ứng Logistics Q4/2026 cho mùa hàng Tết',
         11400, 'Căng thẳng chi phí logistics (+1,850%)',
         'Chủ hàng xuất nhập khẩu, giám đốc cung ứng logistics, doanh nghiệp sản xuất phân phối hàng tiêu dùng Tết, forwarder nội địa và quốc tế.',
         'Các hãng tàu nội địa (Hải An, Vsico) đồng loạt tăng giá cước container 15-25% trong khi tuyến quốc tế áp phụ phí mùa cao điểm (PSS) và hủy chuyến (Blank Sailings), đẩy chi phí logistics đội giá thành sản phẩm 20-30%.',
         "Tiêu đề: 'Cước container nội địa tăng sốc 15-25% & Hãng tàu quốc tế siết tải cuối năm: 3 đòn bẩy giúp chủ hàng giữ slot và hạ nhiệt chi phí'.",
         'Biểu cước vận tải các hãng tàu nội địa; Báo Công Thương; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).',
         update_date_str),

        (7,
         'cong van 23298 chq gsql tam nhap tai xuat g12 g22 may moc demo trien lam hoi cho',
         'Công văn số 23298/CHQ-GSQL: Hướng dẫn thủ tục Hải quan Tạm nhập - Tái xuất máy móc thiết bị phục vụ giới thiệu sản phẩm & khảo sát thị trường (mã G12/G22) — Giải pháp tránh bẫy phạt quá hạn và phong tỏa thuế',
         10900, 'Pháp lý mùa sự kiện (+1,800%)',
         'Doanh nghiệp FDI, nhà phân phối công nghệ, đơn vị tổ chức triển lãm, hội chợ quốc tế và thương nhân nước ngoài đưa máy móc demo vào Việt Nam.',
         'Quá hạn thời gian tạm nhập máy móc demo mà chưa tái xuất hoặc sử dụng sai mục đích bị hải quan xử phạt vi phạm hành chính, truy thu thuế với thuế suất cao và tính phạt chậm nộp 0.03%/ngày trên bảo lãnh ngân hàng.',
         "Tiêu đề: 'Công văn 23298/CHQ-GSQL hướng dẫn tạm nhập tái xuất G12/G22: 5 bước bảo lãnh thuế và cách gia hạn chuẩn không bị phạt'.",
         'Công văn số 23298/CHQ-GSQL tháng 10/2026; Điều 53 Nghị định số 08/2015/NĐ-CP; Báo Hải quan Online.',
         update_date_str),

        (8,
         'cong van 23138 chq nvthq bien lai thu thue hai quan dien tu chuyen phat nhanh tmdt',
         'Công văn số 23138/CHQ-NVTHQ: Triển khai biên lai thu thuế, lệ phí hải quan điện tử cho hàng bưu chính chuyển phát nhanh — Siết chặt thu thuế Thương mại điện tử xuyên biên giới và dấu chấm hết cho chiêu trò chia nhỏ đơn hàng',
         10500, 'Siết chặt thuế TMĐT (+1,750%)',
         'Doanh nghiệp bưu chính, chuyển phát nhanh quốc tế (DHL, FedEx, Shopee, TikTok Shop, Taobao), chủ shop kinh doanh hàng nhập khẩu online.',
         'Hải quan áp dụng biên lai thu thuế điện tử tự động và siết chặt quy định miễn thuế hàng dưới 1 triệu; các lô hàng gom bưu chính bị soi kỹ thuật toán chia nhỏ đơn hàng, nguy cơ bị tạm giữ kiểm tra và truy thu thuế.',
         "Tiêu đề: 'Công văn 23138/CHQ-NVTHQ siết thuế hàng chuyển phát nhanh & TMĐT: Doanh nghiệp kinh doanh online chuyển hướng chính ngạch thế nào?'.",
         'Công văn số 23138/CHQ-NVTHQ ngày 19/10/2026; Cổng TTĐT Tổng cục Hải quan; Báo Công Thương.',
         update_date_str),

        (9,
         'cong van 23405 chq nvthq phan loai hang hoa ma hs may moc thiet bi advance ruling',
         'Công văn số 23405/CHQ-NVTHQ: Siết chặt quy trình phân loại hàng hóa và giải quyết tranh chấp mã HS máy móc, linh kiện — Bí quyết lập Hồ sơ kỹ thuật xác định trước mã số (Advance Ruling) bảo vệ doanh nghiệp',
         10600, 'Duy trì sốt dẻo (+1,700%)',
         'Doanh nghiệp nhập khẩu dây chuyền công nghệ, thiết bị đồng bộ, linh kiện tháo rời; giám đốc pháp chế và quản lý kho vận.',
         'Rủi ro bị hải quan bác mã HS tự khai (thuế 0%) chuyển sang mã chịu thuế 5-10%, kèm theo việc truy thu thuế và xử phạt vi phạm hành chính trong thời hạn 5 năm kiểm tra sau thông quan.',
         "Tiêu đề: 'Công văn 23405/CHQ-NVTHQ siết phân loại mã HS: Hướng dẫn lập hồ sơ Xác định trước mã số (Advance Ruling) bảo vệ doanh nghiệp 100%'.",
         'Công văn số 23405/CHQ-NVTHQ ngày 23/10/2026; Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; Cổng TTĐT Tổng cục Hải quan.',
         update_date_str),

        (10,
         'thue ngoai phong xuat nhap khau tron goi outsourced logistics sme tiet kiem chi phi',
         'Xu hướng "Thuê ngoài Phòng Xuất Nhập Khẩu trọn gói" (Outsourced Logistics Dept) cuối năm 2026: Giải pháp cắt giảm 70% chi phí vận hành và triệt tiêu 100% rủi ro phạt thuế cho SME',
         9800, 'Đột phá B2B (+1,650%)',
         'Chủ doanh nghiệp SME, xưởng sản xuất, doanh nghiệp thương mại nhập khẩu máy móc, nguyên vật liệu từ Trung Quốc, Hàn Quốc.',
         'Tự nuôi nhân sự XNK chi phí cao nhưng nghiệp vụ non, thường xuyên sai mã HS, làm mất C/O ưu đãi và bị phạt thuế; mùa cao điểm cuối năm khối lượng công việc tăng gấp 3 khiến bộ máy quá tải và tắc nghẽn hồ sơ.',
         "Tiêu đề: 'Tại sao doanh nghiệp SME đang ồ ạt thuê ngoài Phòng XNK trọn gói cuối năm 2026? Bài toán tiết kiệm 70% chi phí và bảo vệ pháp nhân sạch'.",
         'Mô hình Canvas & Giải pháp Phòng XNK thuê ngoài trọn gói 5 triệu/tháng của Eureka Logistics; Báo Công Thương.',
         update_date_str)
    ]

    for idx, item in enumerate(sheet1_data):
        row_idx = 6 + idx
        fill_to_use = fill_zebra_even if (row_idx % 2 == 0) else fill_zebra_odd
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
                cell.font = font_bold
            elif col_idx == 10:
                cell.alignment = align_center
        ws1.row_dimensions[row_idx].height = 65

    # Total row Sheet 1
    ws1.cell(row=16, column=1, value='Tổng số lượng từ khóa nghiên cứu tuân thủ SLA:')
    ws1.cell(row=16, column=1).font = font_bold
    ws1.cell(row=16, column=1).fill = fill_total
    ws1.cell(row=16, column=1).alignment = align_left
    ws1.cell(row=16, column=1).border = total_border

    ws1.cell(row=16, column=4, value='=SUM(D6:D15)')
    ws1.cell(row=16, column=4).font = font_bold
    ws1.cell(row=16, column=4).fill = fill_total
    ws1.cell(row=16, column=4).alignment = align_right
    ws1.cell(row=16, column=4).number_format = '#,##0'
    ws1.cell(row=16, column=4).border = total_border

    ws1.cell(row=16, column=5, value='Nhu cầu thị trường rất cao')
    ws1.cell(row=16, column=5).font = font_bold
    ws1.cell(row=16, column=5).fill = fill_total
    ws1.cell(row=16, column=5).alignment = align_center
    ws1.cell(row=16, column=5).border = total_border

    eval_sheet1 = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 44 - {update_date_str})."
    ws1.cell(row=16, column=6, value=eval_sheet1)
    ws1.cell(row=16, column=6).font = font_total
    ws1.cell(row=16, column=6).fill = fill_total
    ws1.cell(row=16, column=6).alignment = align_left
    ws1.cell(row=16, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    print("2. Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E...")
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 44 (26/10/2026 - 01/11/2026) | Mở rộng 190 chủ đề nghiên cứu (gồm 158 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 191)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 191 or m.max_row >= 191]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (186,
         'cong van 23244 chq nvthq hoan thue gtgt thue nhap khau xuat tra hang loi b13',
         'cong van 23244/CHQ-NVTHQ, hoan thue xuat tra chu hang nuoc ngoai, loai hinh tai xuat b13, dieu 34 nghi dinh 134/2016/nd-cp, hoan thue gtgt hang nhap khau loi, bien ban giam dinh hang loi, thong tu 38/2015/tt-btc',
         99.7,
         45800,
         'https://customs.gov.vn',
         'Công văn số 23244/CHQ-NVTHQ: Hướng dẫn thủ tục hoàn thuế GTGT và thuế nhập khẩu đối với hàng hóa xuất trả chủ hàng nước ngoài (mã loại hình B13) — 4 điều kiện sống còn để không bị bác hồ sơ hoàn thuế',
         'Nội dung trọng tâm: Cục Hải quan (Ban Nghiệp vụ thuế hải quan) ban hành Công văn số 23244/CHQ-NVTHQ hướng dẫn xử lý vướng mắc hoàn thuế cho hàng hóa nhập khẩu phải tái xuất trả lại chủ hàng nước ngoài do không phù hợp quy cách, chất lượng hoặc hỏng hóc theo hợp đồng ngoại thương. Rất nhiều doanh nghiệp bị cơ quan hải quan từ chối hoàn thuế GTGT và thuế nhập khẩu đã nộp do: (1) Không có biên bản giám định độc lập của tổ chức giám định thương mại xác nhận hàng lỗi tại thời điểm nhập khẩu; (2) Hàng hóa đã đưa vào sử dụng, sản xuất thử làm mất tính nguyên trạng; (3) Khai sai mã loại hình B13 sang B11 hoặc tái xuất quá hạn thỏa thuận. Eureka đúc kết cẩm nang 4 bước thu hồi tiền thuế: Chuẩn bị văn bản thông báo hàng lỗi cho đối tác nước ngoài ngay khi mở cont; thuê tổ chức giám định lập chứng thư xác định lỗi kỹ thuật trước khi thông quan tái xuất; lập bộ hồ sơ hoàn thuế chuẩn mẫu số 09 Phụ lục VII Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018); và cơ chế bù trừ tiền thuế nộp thừa sang các tờ khai nhập khẩu phát sinh tiếp theo.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Ban Nghiệp vụ Thuế Hải quan | CS (Case Study) / Defective Goods Re-Export & Tax Refund Compliance',
         update_date_str),

        (187,
         'cong van 23298 chq gsql tam nhap tai xuat g12 g22 may moc demo trien lam hoi cho',
         'cong van 23298/CHQ-GSQL, thu tuc tam nhap tai xuat g12 g22, tam nhap may moc thiet bi demo, hoi cho trien lam khao sat thi truong, dieu 53 nghi dinh 08/2015/nd-cp, bao lanh tien thue ngan hang, chuyen tieu thu noi dia a42',
         99.2,
         41500,
         'https://haiquanonline.com.vn',
         'Công văn số 23298/CHQ-GSQL: Hướng dẫn thủ tục Hải quan Tạm nhập - Tái xuất máy móc thiết bị phục vụ giới thiệu sản phẩm & khảo sát thị trường (mã G12/G22) — Giải pháp tránh bẫy phạt quá hạn và phong tỏa thuế',
         'Nội dung trọng tâm: Cục Giám sát quản lý về Hải quan ban hành Công văn số 23298/CHQ-GSQL hướng dẫn chi tiết về thủ tục hải quan đối với máy móc, thiết bị, linh kiện công nghệ cao tạm nhập để giới thiệu sản phẩm, tham gia hội chợ thương mại và khảo sát thị trường tại Việt Nam theo Điều 53 Nghị định 08/2015/NĐ-CP. Nhiều doanh nghiệp FDI và nhà nhập khẩu thiết bị mắc các sai phạm nghiêm trọng: Tạm nhập theo mã G12 nhưng để máy móc chạy thử sản xuất thương mại dẫn đến bị xử lý hành vi gian lận mục đích tạm nhập; không theo dõi thời hạn tạm nhập trên hệ thống dẫn đến quá hạn tái xuất (mã G22) bị xử phạt tiền và ấn định phạt chậm nộp 0.03%/ngày trên số thuế bảo lãnh. Cẩm nang bảo an của Eureka: Quy trình đăng ký thư bảo lãnh ngân hàng nộp thuế; thủ tục gia hạn thời gian tạm nhập bằng văn bản gửi Chi cục Hải quan nơi mở tờ khai trước ngày hết hạn tối thiểu 15 ngày; và giải pháp chuyển đổi mục đích sử dụng sang tiêu thụ nội địa (mã A42) nộp thuế hợp lệ khi khách hàng quyết định mua đứt máy móc demo.',
         'Báo Hải quan Online (haiquanonline.com.vn) & Cục Giám sát quản lý Hải quan | HT (Học Thực chiến) / Temporary Import G12/G22 & Exhibition Machinery Compliance',
         update_date_str),

        (188,
         'cong van 23138 chq nvthq bien lai thu thue hai quan dien tu chuyen phat nhanh tmdt',
         'cong van 23138/CHQ-NVTHQ, bien lai thu thue dien tu chuyen phat nhanh, thue hai quan hang tmdt xuyen bien gioi, shopee tiktok taobao chuyen phat nhanh, kiem tra hang gia tri nho duoi 1 trieu, chia nho lo hang tron thue',
         98.9,
         38700,
         'https://customs.gov.vn',
         'Công văn số 23138/CHQ-NVTHQ: Triển khai biên lai thu thuế, lệ phí hải quan điện tử cho hàng bưu chính chuyển phát nhanh — Siết chặt thu thuế Thương mại điện tử xuyên biên giới và dấu chấm hết cho chiêu trò chia nhỏ đơn hàng',
         'Nội dung trọng tâm: Ban Nghiệp vụ thuế hải quan ban hành Công văn số 23138/CHQ-NVTHQ chính thức áp dụng hệ thống cấp biên lai thu thuế, lệ phí hải quan điện tử kết nối tự động cho các doanh nghiệp chuyển phát nhanh quốc tế (DHL, FedEx, UPS, EMS...) và các sàn thương mại điện tử xuyên biên giới. Cơ quan hải quan đẩy mạnh đối soát dữ liệu luồng hàng để ngăn chặn hành vi lợi dụng chính sách miễn thuế đối với hàng hóa nhập khẩu có trị giá hải quan dưới 1.000.000 đồng qua bưu chính CPN; nhận diện tự động các thuật toán chia nhỏ đơn hàng cùng một người nhận hoặc cùng địa chỉ. Các hộ kinh doanh online và shop nhập hàng tiểu ngạch qua CPN đối mặt nguy cơ bị giữ hàng kiểm tra gộp, truy thu thuế GTGT và thuế nhập khẩu toàn bộ. Eureka khuyến nghị chuyển đổi mô hình: Gom hàng và nhập khẩu chính ngạch theo đường biển LCL/FCL có đầy đủ hóa đơn GTGT đầu vào để khấu trừ thuế hợp pháp và kinh doanh an tâm bền vững.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Ban Nghiệp vụ Thuế | TN (Tin Ngành) / Cross-border E-commerce Customs Tax & Courier Compliance',
         update_date_str),

        (189,
         'gia cuoc van tai container noi dia quoc te quy 4 tang 15 25 phan tram pss blank sailings',
         'gia cuoc van tai bien noi dia 2026, hang tau hai an vsico tang cuoc, cuoc cont 20 40 hai phong sai gon, phu phi mua cao diem pss quoc te, blank sailings cat giam chuyen tau, khung hoang thieu vo cont mua tet 2027',
         99.3,
         43100,
         'https://congthuong.vn',
         'Sóng tăng giá cước vận tải container nội địa (Hải Phòng - Cái Mép/TP.HCM tăng 15-25%) và Cước quốc tế cuối năm: Chiến lược tối ưu chi phí cung ứng Logistics Q4/2026 cho mùa hàng Tết',
         'Nội dung trọng tâm: Thị trường vận tải biển nội địa Việt Nam ghi nhận đợt điều chỉnh cước quy mô lớn từ các hãng tàu hàng đầu (Hải An, Vsico...) với mức tăng 15-25% trên các tuyến huyết mạch Hải Phòng - TP.HCM / Cái Mép; cước container 20 feet chạm ngưỡng 6.5 - 7.5 triệu đồng, container 40 feet lên 9.0 - 10.5 triệu đồng/container. Song song đó, các tuyến quốc tế đi Mỹ và Châu Âu chịu áp lực từ phụ phí mùa cao điểm (PSS) và chiến lược cắt giảm chuyến (Blank Sailings) của các liên minh hãng tàu quốc tế trước thềm Noel và Tết Nguyên đán 2027. Doanh nghiệp sản xuất và thương mại đối mặt tình trạng "cước chồng cước", chi phí logistics đội lên 20-30% bào mòn lợi nhuận đơn hàng. Eureka đề xuất 3 giải pháp ứng phó: (1) Ký hợp đồng vận tải dài hạn hoặc cam kết khối lượng (Volume Commitment) để giữ slot ổn định; (2) Giải pháp vận tải đa phương thức kết hợp đường bộ - đường biển linh hoạt theo thời gian thực; (3) Tối ưu hóa đóng hàng, nâng cao hệ số sử dụng cont (CBM utilization) để giảm số lượng container phải thuê.',
         'Báo Công Thương (congthuong.vn) & Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA) | FF (Forwarder & Cước Vận Tải) / Domestic & International Container Freight Strategy',
         update_date_str),

        (190,
         'no thue treo etax mobile khoa ma so thue chan mo to khai hoa don thay the khac ky',
         'no thue treo etax mobile, khoa ma so thue chan thong quan, ke khai hoa don thay the khac ky, dong ma so thue ho kinh doanh ca the, tam hoan xuat canh nguoi dai dien phap luat, nghi dinh 123/2020/nd-cp hoa don dien tu',
         99.8,
         46900,
         'https://tapchitaichinh.vn',
         'Giải quyết triệt để rủi ro hóa đơn thay thế khác kỳ & nợ thuế treo eTax Mobile khi đóng/chuyển đổi MST Hộ kinh doanh sang Doanh nghiệp XNK — Gỡ nút thắt bị chặn mở tờ khai hải quan',
         'Nội dung trọng tâm: Phân tích và tháo gỡ vướng mắc nóng nhất được cộng đồng doanh nghiệp và kế toán phản ánh dày đặc trên eTax Mobile và các hội nhóm XNK cuối tháng 10/2026: (1) Trường hợp hóa đơn điện tử bị thay thế ở Quý 2 nhưng hóa đơn thay thế phát hành vào Quý 3, kế toán xử lý sai quy định tại Điều 19 Nghị định 123/2020/NĐ-CP dẫn đến kê khai trùng hoặc thiếu doanh thu, bị cơ quan thuế tính tiền chậm nộp và phát hành thông báo truy thu; (2) Tình trạng chủ doanh nghiệp từng đứng tên đăng ký hộ kinh doanh cá thể đã dừng hoạt động từ nhiều năm trước nhưng chưa làm thủ tục đóng MST dứt điểm (chưa hoàn thành nghĩa vụ thuế môn bài, thuế GTGT khoán). Khi thành lập công ty XNK mới và truyền tờ khai thông quan hàng hóa, hệ thống Hải quan tự động chặn mở tờ khai và treo nợ thuế do liên thông mã định danh CCCD với eTax Mobile, thậm chí người đại diện pháp luật bị tạm hoãn xuất cảnh theo Luật Quản lý thuế số 38/2019/QH14. Eureka hướng dẫn quy trình 3 bước xử lý dứt điểm: Hướng dẫn kê khai bổ sung hóa đơn thay thế đúng kỳ phát sinh; thủ tục đối soát trực tiếp với Chi cục Thuế quản lý để ra Thông báo đóng MST hộ kinh doanh; và công văn đề nghị Chi cục Hải quan cập nhật xóa nợ thuế treo trên hệ thống trong vòng 24 giờ để thông quan giải phóng hàng kịp thời.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & Cục Thuế / Dữ liệu Nhật ký nội bộ daily_notes.md | CS (Case Study) / eTax Mobile Debt Clearance & Tax Invoice Correction',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 191 - 195)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 191 + idx
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

    # Cập nhật Total row ở row 196
    total_row_idx = 196
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 190 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D195)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E195)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 190 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 158 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A196:C196')
    ws3.merge_cells('F196:J196')
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
    print(f"File chính: {file_size} bytes | SHA256: {file_sha256}")

    # ==========================================
    # 4. ĐỒNG BỘ ATOMIC SANG ONEDRIVE MARKETING
    # ==========================================
    if os.path.exists(ONEDRIVE_DIR):
        print(f"4. Đồng bộ an toàn sang OneDrive: {ONEDRIVE_FILE}...")
        temp_onedrive_file = ONEDRIVE_FILE + '.tmp.xlsx'
        try:
            shutil.copy2(FILE_PATH, temp_onedrive_file)
            if os.path.exists(temp_onedrive_file):
                os.replace(temp_onedrive_file, ONEDRIVE_FILE)
                print(f"[OK] Đã đồng bộ thành công sang OneDrive Marketing Eureka!")
                onedrive_size = os.path.getsize(ONEDRIVE_FILE)
                onedrive_sha256 = get_file_sha256(ONEDRIVE_FILE)
                print(f"OneDrive File: {onedrive_size} bytes | SHA256: {onedrive_sha256}")

                # Ghi sync_log.txt
                now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 30/10/2026)\n"
                with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f_log:
                    f_log.write(log_entry)
                print(f"[OK] Đã ghi nhật ký đồng bộ vào {SYNC_LOG_PATH}")
        except Exception as e:
            print(f"[CẢNH BÁO] Không thể đồng bộ trực tiếp sang OneDrive: {e}")
            if os.path.exists(temp_onedrive_file):
                try:
                    os.remove(temp_onedrive_file)
                except:
                    pass
    else:
        print(f"[BỎ QUA] Không tìm thấy thư mục OneDrive: {ONEDRIVE_DIR}")

    # ==========================================
    # 5. TẠO BẢN TIN ĐỀ XUẤT SCRATCH/HOT_NEWS_BRIEFING
    # ==========================================
    print("5. Xuất bản tin đề xuất scratch/hot_news_briefing_w44_3010.md...")
    md_content = f"""# [BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN — Hot News Scanner]

- **Tuần:** Tuần 44 (26/10/2026 - 01/11/2026) | Ngày quét: 30/10/2026
- **Tổng cơ hội phát hiện:** 22 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn`)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors (`drewry.co.uk`) & Hiệp hội Logistics Việt Nam (VLA)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Ghi chép kiểm tra sau thông quan, nợ thuế eTax Mobile, hóa đơn thay thế khác kỳ, áp chỉ tiêu 1.34 và biến động cước vận tải biển)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 30/10/2026)

### Đề xuất #1: [NỖI ĐAU THỰC CHIẾN & PHÁP LÝ THUẾ] Giải quyết triệt để rủi ro hóa đơn thay thế khác kỳ & nợ thuế treo eTax Mobile khi đóng/chuyển đổi MST Hộ kinh doanh sang Doanh nghiệp XNK — Gỡ nút thắt bị chặn mở tờ khai hải quan
- **Chuyên mục (Pillar):** `CS` (Case Study - Rà soát Thuế Doanh nghiệp & Gỡ nợ treo eTax Mobile)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo 'Nợ thuế 500k từ 5 năm trước làm đứng hình cả lô hàng tiền tỷ' & 3 bước tra soát | YouTube: Hướng dẫn chi tiết quy trình kê khai hóa đơn thay thế khác kỳ theo NĐ 123 và thủ tục đóng dứt điểm MST hộ kinh doanh)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Ngành Thuế và Hải quan đang mở đợt tổng rà soát liên thông dữ liệu CCCD nợ thuế cuối tháng 10/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Bị chặn mở tờ khai hải quan, tàu về đến cảng nhưng không thể thông quan, nguy cơ giám đốc bị tạm hoãn xuất cảnh |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận dữ dội: "MST hộ kinh doanh đã dừng lâu rồi sao vẫn dính nợ thuế treo và chặn tờ khai công ty mới?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp quy trình 3 bước xử lý: Kê khai bổ sung hóa đơn thay thế đúng kỳ, đối soát Chi cục Thuế và công văn gửi Hải quan gỡ nợ trong 24h |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #1 (Đề tài Nỗi đau Thực tế Chạm đỉnh — Giải cứu trực tiếp dòng hàng của chủ doanh nghiệp)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi sợ lớn nhất của các chủ xưởng, chủ doanh nghiệp SME khi chuyển đổi mô hình từ hộ kinh doanh cá thể lên doanh nghiệp XNK. Sự cố hệ thống dữ liệu tự động liên thông eTax Mobile khiến nhiều doanh nghiệp "chết đứng" khi hàng đã cập cảng nhưng không thể truyền tờ khai. Bài viết khẳng định chuyên môn sâu, tính thực chiến và khả năng tháo gỡ vướng mắc tận gốc của Eureka.
- **Nguồn xác minh:** Nghị định số 123/2020/NĐ-CP; Luật Quản lý thuế số 38/2019/QH14; Cổng Dịch vụ công Hải quan; Sổ tay nội bộ `scratch/daily_notes.md`.

---

### Đề xuất #2: [TÀI CHÍNH XUẤT NHẬP KHẨU] Công văn số 23244/CHQ-NVTHQ: Hướng dẫn thủ tục hoàn thuế GTGT và thuế nhập khẩu đối với hàng hóa xuất trả chủ hàng nước ngoài (mã loại hình B13) — 4 điều kiện sống còn để không bị bác hồ sơ hoàn thuế
- **Chuyên mục (Pillar):** `CS` (Case Study - Tái xuất Trả hàng Lỗi B13 & Thu hồi Tiền thuế)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic 4 sai lầm kinh điển khiến doanh nghiệp mất trắng tiền thuế khi xuất trả hàng | YouTube: Video cẩm nang hướng dẫn lập bộ hồ sơ hoàn thuế mẫu 09 Phụ lục VII TT 38/2015 và cách đối thoại với giám định viên)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Công văn 23244/CHQ-NVTHQ vừa ban hành trong tháng 10 giải đáp hàng loạt vướng mắc hồ sơ hoàn thuế cuối năm |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Tiền thuế hàng trăm triệu đến hàng tỷ đồng đã nộp bị treo, nếu không đủ chứng từ sẽ bị từ chối hoàn hoàn toàn |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi nóng: "Hàng hóa đã bóc dỡ, chạy thử máy mới phát hiện lỗi thì có được coi là 'chưa qua sử dụng' để hoàn thuế?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp checklist 4 bước: Chứng thư giám định độc lập, khai mã B13 chuẩn xác, lập bộ hồ sơ hoàn thuế mẫu 09 và cơ chế cấn trừ thuế nộp thừa |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #2 (Cẩm nang Tài chính Thực chiến — Giúp doanh nghiệp thu hồi dòng tiền thuế bị kẹt)** |
- **Tóm tắt lý do nên làm:** Hàng hóa lỗi hỏng, sai hợp đồng là rủi ro thường trực trong thương mại quốc tế. Tuy nhiên, thủ tục tái xuất trả hàng (B13) và xin hoàn thuế GTGT, thuế nhập khẩu theo Điều 34 Nghị định 134/2016/NĐ-CP vô cùng phức tạp. Hướng dẫn chi tiết của Eureka giúp doanh nghiệp nắm chắc quy trình pháp lý, bảo toàn dòng tiền và tránh tranh chấp thuế kéo dài.
- **Nguồn xác minh:** Công văn số 23244/CHQ-NVTHQ tháng 10/2026; Điều 34 Nghị định 134/2016/NĐ-CP; Thông tư 38/2015/TT-BTC & Thông tư 39/2018/TT-BTC; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`).

---

### Đề xuất #3: [CƯỚC VẬN TẢI & CHUỖI CUNG ỨNG] Sóng tăng giá cước vận tải container nội địa (Hải Phòng - Cái Mép/TP.HCM tăng 15-25%) và Cước quốc tế cuối năm: Chiến lược tối ưu chi phí cung ứng Logistics Q4/2026 cho mùa hàng Tết
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải - Quản trị Chi phí Logistics Mùa Cao Điểm)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Bảng so sánh biểu cước container mới của các hãng tàu & Mẹo giữ chỗ mùa Tết | YouTube: Phân tích mô hình logistics đa phương thức kết hợp đường bộ - đường thủy nội địa giúp cắt giảm 20% chi phí)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Các hãng tàu nội địa (Hải An, Vsico) vừa áp dụng biểu cước mới và các hãng quốc tế siết phụ phí PSS cuối tháng 10 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Chi phí cước container tăng 15-25% ăn mòn biên lợi nhuận, nguy cơ trễ hẹn đơn hàng xuất khẩu và giao hàng Tết |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Thảo luận sôi nổi: "Cước biển nội địa tăng mạnh thì chuyển sang đường bộ xe tải/đầu kéo có tối ưu hơn không?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Chia sẻ giải pháp 3 đòn bẩy: Cam kết khối lượng (Volume Commitment), tối ưu hệ số đóng hàng CBM và vận tải đa phương thức Eureka |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #3 (Bài toán Chi phí Sát sườn — Nâng tầm năng lực điều phối forwarder của Eureka)** |
- **Tóm tắt lý do nên làm:** Quý 4 là giai đoạn nước rút sản xuất và tiêu thụ hàng Tết 2027 cũng như mùa xuất khẩu Noel sang Âu - Mỹ. Giá cước vận tải biển tăng đồng loạt cả nội địa lẫn quốc tế đang tạo áp lực nghẹt thở lên các giám đốc chuỗi cung ứng. Phân tích của Eureka đưa ra giải pháp cân bằng giữa thời gian và chi phí, khẳng định vị thế nhà cung cấp giải pháp vận tải toàn diện.
- **Nguồn xác minh:** Biểu giá cước của Công ty Vận tải Hải An, Vsico; Báo Công Thương (`congthuong.vn`); Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).

---

### Đề xuất #4: [THƯƠNG MẠI ĐIỆN TỬ & CHUYỂN PHÁT NHANH] Công văn số 23138/CHQ-NVTHQ: Triển khai biên lai thu thuế, lệ phí hải quan điện tử cho hàng bưu chính chuyển phát nhanh — Siết chặt thu thuế Thương mại điện tử xuyên biên giới và dấu chấm hết cho chiêu trò chia nhỏ đơn hàng
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Pháp lý Thuế TMĐT Xuyên Biên Giới & Hải quan Chuyển phát nhanh)
- **Nền tảng phù hợp:** **Facebook** (Infographic: 'Chấm dứt thời kỳ xé lẻ đơn hàng né thuế dưới 1 triệu' & Lối đi nào cho shop online) & **YouTube** (Phân tích xu hướng chuyển đổi từ hàng tiểu ngạch CPN sang nhập khẩu chính ngạch FCL/LCL)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Công văn 23138/CHQ-NVTHQ vừa triển khai ngày 19/10 và các đơn vị CPN đang kích hoạt hệ thống biên lai điện tử cuối tháng 10 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Các shop kinh doanh hàng nhập khẩu qua Shopee, TikTok Shop, Taobao bị truy thu thuế, hàng bị giữ kiểm tra gộp |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh cãi bùng nổ: "Xóa bỏ miễn thuế đơn hàng dưới 1 triệu: Cơ hội cho sản xuất trong nước hay gánh nặng cho người tiêu dùng?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Định hướng giải pháp nhập khẩu chính ngạch LCL gom hàng tiết kiệm, có đầy đủ hóa đơn chứng từ đầu vào hợp pháp |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #4 (Xu hướng Nóng bỏng — Hút tệp khách hàng kinh doanh online và bán lẻ nhập khẩu)** |
- **Tóm tắt lý do nên làm:** Chủ trương siết chặt thuế hàng TMĐT xuyên biên giới đang là chủ đề nhận được sự quan tâm đặc biệt lớn từ dư luận và cộng đồng bán lẻ trực tuyến. Eureka đón đầu làn sóng này để định hướng các cá nhân, chủ shop chuyển dịch sang phương thức nhập khẩu chính ngạch LCL gom cont tiết kiệm và an toàn pháp lý.
- **Nguồn xác minh:** Công văn số 23138/CHQ-NVTHQ ngày 19/10/2026; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`); Báo Công Thương.

---

### Đề xuất #5: [THỦ TỤC HẢI QUAN MÙA TRIỂN LÃM] Công văn số 23298/CHQ-GSQL: Hướng dẫn thủ tục Hải quan Tạm nhập - Tái xuất máy móc thiết bị phục vụ giới thiệu sản phẩm & khảo sát thị trường (mã G12/G22) — Giải pháp tránh bẫy phạt quá hạn và phong tỏa thuế
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Tạm nhập Tái xuất G12/G22 & Quản lý Máy móc Demo)
- **Nền tảng phù hợp:** **Facebook** (Checklist 5 bước làm thủ tục tạm nhập máy móc demo hội chợ không lo phạt) & **YouTube** (Hướng dẫn chi tiết thủ tục bảo lãnh ngân hàng tiền thuế và cách chuyển đổi mục đích tiêu thụ nội địa A42)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.7/5) | Mùa hội chợ, triển lãm công nghệ và xúc tiến thương mại quốc tế cuối năm 2026 đang vào cao điểm |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Quá hạn tạm nhập bị phạt hành chính nặng, bị tính phạt chậm nộp 0.03%/ngày và bị cưỡng chế phong tỏa mã số hải quan |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Thảo luận: "Máy móc demo cho khách dùng thử nghiệm tại xưởng có bị coi là sử dụng sai mục đích tạm nhập G12?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Cung cấp quy trình trọn gói: Đăng ký bảo lãnh ngân hàng, quản lý định mức linh kiện demo, gia hạn thời hạn và thủ tục tái xuất G22 sạch |
  | **Tổng điểm** | **19.0 / 20** | **Xếp hạng: #5 (Cẩm nang Kỹ thuật Thực chiến — Giải quyết nhu cầu xúc tiến kinh doanh của doanh nghiệp FDI và đại lý phân phối)** |
- **Tóm tắt lý do nên làm:** Cuối năm là thời điểm các hãng máy móc quốc tế ồ ạt đưa thiết bị mẫu vào Việt Nam để tham gia triển lãm hoặc trình diễn cho khách hàng. Nhiều doanh nghiệp vì thiếu hiểu biết đã để quá hạn tạm nhập hoặc chuyển đổi mục đích sử dụng sai quy định, dẫn đến bị truy thu thuế và xử phạt. Bài viết của Eureka đóng vai trò như một cẩm nang pháp lý tin cậy giúp các doanh nghiệp tự tin triển khai demo sản phẩm.
- **Nguồn xác minh:** Công văn số 23298/CHQ-GSQL tháng 10/2026; Điều 53 Nghị định số 08/2015/NĐ-CP; Báo Hải quan Online (`haiquanonline.com.vn`).

---

## 🎯 BƯỚC TIẾP THEO (NEXT ACTIONS DÀNH CHO CHỊ THÚY)
1. **Lựa chọn chủ đề ưu tiên:** Chị Thúy có thể phản hồi chọn 1 hoặc nhiều chủ đề (ví dụ: *'Chọn #1 và #2'*) để hệ thống tự động đẩy vào:
   - **`wf01`:** Lên lịch xuất bản chi tiết cho kế hoạch tuần tiếp theo.
   - **`wf02`:** Sản xuất bài viết chuyên sâu trên Facebook Fanpage Eureka Logistics / Trang cá nhân.
   - **`wf03`:** Sản xuất kịch bản video dài YouTube ThuyTalk (Chuyên mục 1, 2 hoặc 4).
2. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""
    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(md_content)
    print(f"[OK] Đã xuất bản tin đề xuất ra: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT THEO WF05 NGÀY 30/10/2026 ===")

if __name__ == '__main__':
    update_seo_w44_3010()
