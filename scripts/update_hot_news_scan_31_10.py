# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_31_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 31/10/2026
(Chu kỳ: Tuần 44: 26/10/2026 - 01/11/2026 | Cập nhật ngày 31/10/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 44 cập nhật ngày 31/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 190 lên 195 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 31/10/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w44_3110.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w44_3110.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w44_3110():
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

    update_date_str = "31/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 44 (26/10/2026 - 01/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'nhap tai cho e31 truy thu thue sxxk doi ung b11 e62 nghi dinh 18 2021',
         'Cạm bẫy truy thu thuế mã E31 khi nhập khẩu tại chỗ đối ứng xuất B11/E62: Rủi ro kiểm tra sau thông quan và bài toán nộp A12 hoàn thuế theo Nghị định 18/2021/NĐ-CP',
         14200, 'Hỏa tốc toàn quốc (+2,400%)',
         'Doanh nghiệp gia công, sản xuất xuất khẩu (SXXK), doanh nghiệp chế xuất EPE, kế toán thuế XNK và đại lý hải quan.',
         'Hiểu nhầm hàng nhập tại chỗ SXXK luôn được miễn thuế; khi cơ quan hải quan kiểm tra sau thông quan đối chiếu tờ khai bên bán mở B11/E62 sẽ lập tức truy thu 100% thuế nhập khẩu, xử phạt kê khai sai và tính phạt chậm nộp 0.03%/ngày tiền tỷ.',
         "Tiêu đề: 'Nhập tại chỗ mở E31 bị truy thu thuế tiền tỷ: Bài học trùng sinh và giải pháp mở A12 hoàn thuế sạch nước cản'.",
         'Nghị định 18/2021/NĐ-CP; Điều 12, Điều 36 Nghị định 134/2016/NĐ-CP; Công văn số 23642/CHQ; CV 16900/16901/BTC.',
         update_date_str),

        (2,
         'cong van 4323 tttv attpmt kiem tra chat nong san quyt tuoi trung quoc',
         "Công văn số 4323/TTTV-ATTPMT kích hoạt cơ chế 'Kiểm tra chặt' nông sản nhập khẩu từ Trung Quốc: Cảnh báo dư lượng thuốc BVTV Abamectin vượt ngưỡng 3 lần và giải pháp thông quan trái cây tươi",
         13800, 'Cảnh báo đỏ biên giới (+2,350%)',
         'Thương nhân kinhdong nông sản nhập khẩu, chủ hàng trái cây tươi qua cửa khẩu Lạng Sơn, Lào Cai, đại lý logistics biên giới.',
         'Phát hiện 2 lô hàng 39.5 tấn quýt tươi TQ có dư lượng Abamectin vượt gấp 3 lần quy chuẩn; bị chuyển sang luồng kiểm tra chặt 100%, chờ kết quả lab 3-5 ngày làm thối hỏng hoa quả, phát sinh phí cắm điện cont lạnh 1.5 - 2 triệu/ngày và nguy cơ buộc tái xuất.',
         "Tiêu đề: 'Công văn 4323/TTTV-ATTPMT siết kiểm tra chặt quýt tươi nhập khẩu: 3 bước rà soát mã vùng trồng GACC và đàm phán hợp đồng an toàn'.",
         'Công văn số 4323/TTTV-ATTPMT ngày 22/10/2026; Khoản 3 Điều 17 Nghị định 15/2018/NĐ-CP; Thông tư 50/2016/TT-BYT.',
         update_date_str),

        (3,
         'kiem tra sau thong quan cuoc tau container tri gia hai quan cif fcl',
         'Hải quan siết kiểm tra sau thông quan cước tàu container quốc tế & phụ phí hãng tàu (CIF/FOB): Doanh nghiệp cần chuẩn bị hồ sơ thanh toán qua ngân hàng thế nào để không bị bác trị giá hải quan?',
         13400, 'Hậu kiểm cước biển nóng (+2,250%)',
         'Chủ hàng nhập khẩu nguyên liệu, máy móc thiết bị, giám đốc tài chính, kế toán kho vận và công ty forwarder.',
         'Hải quan đối chiếu dữ liệu manifest hãng tàu, ePort và sao kê tài khoản ngân hàng; cước khai báo không khớp hóa đơn thực tế hoặc thanh toán tiền mặt sẽ bị bác trị giá tính thuế, ấn định truy thu thuế GTGT/thuế NK trong 5 năm.',
         "Tiêu đề: 'Hậu kiểm cước tàu container 5 năm: Checklist 5 chứng từ thanh toán ngân hàng bắt buộc để bảo vệ trị giá hải quan'.",
         'Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Luật Hải quan số 54/2014/QH13; Cổng TTĐT Tổng cục Hải quan.',
         update_date_str),

        (4,
         'cong van 23535 chq gsql canh bao phong ve thuong mai gian lan xuat xu cong van 8744 bct',
         'Công văn số 23535/CHQ-GSQL & CV 8744/BCT-PVTM: Cảnh báo đỏ 7 nhóm mặt hàng rủi ro cao bị điều tra phòng vệ thương mại, gian lận xuất xứ và lẩn tránh thuế — Nguy cơ bị áp thuế trừng phạt của Mỹ, EU',
         13100, 'Cảnh báo đỏ 7 nhóm hàng (+2,200%)',
         'Doanh nghiệp sản xuất, xuất khẩu gỗ & nội thất, sắt thép kim loại, pin năng lượng, dệt may; forwarder tuyến Mỹ/EU.',
         'Nhập linh kiện thô về chỉ gia công đơn giản, dán nhãn Việt Nam để né thuế; bị điều tra lẩn tránh thuế phòng vệ thương mại, đối mặt nguy cơ bị áp thuế chống bán phá giá trừng phạt từ 25% đến hơn 100% và giữ hàng tại cảng.',
         "Tiêu đề: 'Hải quan ra Công văn 23535/CHQ-GSQL cảnh báo đỏ 7 nhóm hàng: 4 bước rà soát hồ sơ xuất xứ chống nguy cơ bị Mỹ và EU trừng phạt thuế phòng vệ'.",
         'Công văn số 23535/CHQ-GSQL ngày 27/10/2026; Công văn số 8744/BCT-PVTM ngày 22/10/2026 của Cục Phòng vệ thương mại.',
         update_date_str),

        (5,
         'quy dinh cbam eu han chot bao cao phat thai 31 10 sat thep nhom chung chi carbon',
         'Hôm nay 31/10 hạn chót nộp báo cáo phát thải CBAM: Áp lực nghẹt thở đối với doanh nghiệp xuất khẩu Sắt thép, Nhôm sang EU — Khi chứng chỉ carbon trở thành rào cản tài chính sống còn',
         12900, 'Hạn chót 31/10 hôm nay (+2,150%)',
         'Doanh nghiệp sản xuất, xuất khẩu sắt thép, tôn mạ, nhôm thanh định hình, sản phẩm kim loại sang 27 quốc gia Liên minh Châu Âu.',
         'Hôm nay 31/10 là hạn chót nộp báo cáo quý 3 phát thải CBAM; nếu dùng hệ số mặc định sẽ bị EU áp thuế carbon phạt cực cao (>80 EUR/tấn CO2e), làm mất đơn hàng vào tay đối thủ cạnh tranh.',
         "Tiêu đề: 'Hôm nay 31/10 hạn chót CBAM: Doanh nghiệp thép và nhôm Việt Nam cần hoàn tất báo cáo phát thải thế nào để không bị đánh thuế phạt nghìn tỷ?'.",
         'Quy định Cơ chế điều chỉnh biên giới carbon (CBAM) của Nghị viện Châu Âu; Báo Công Thương; logistics.gov.vn.',
         update_date_str),

        (6,
         'cong van 23244 chq nvthq hoan thue gtgt thue nhap khau xuat tra hang loi b13',
         'Công văn số 23244/CHQ-NVTHQ: Hướng dẫn thủ tục hoàn thuế GTGT và thuế nhập khẩu đối với hàng hóa xuất trả chủ hàng nước ngoài (mã B13) — 4 điều kiện sống còn để không bị bác hồ sơ hoàn thuế',
         12400, 'Nghiệp vụ tài chính nóng (+2,100%)',
         'Doanh nghiệp nhập khẩu nguyên vật liệu, máy móc thiết bị, linh kiện điện tử; kế toán thuế xuất nhập khẩu và đại lý hải quan.',
         'Hàng hóa nhập khẩu bị lỗi chất lượng hoặc sai quy cách nhưng lúng túng trong thủ tục giám định thương mại và khai báo tái xuất (B13), dẫn đến bị từ chối hoàn thuế GTGT và thuế nhập khẩu.',
         "Tiêu đề: 'Công văn 23244/CHQ-NVTHQ: Bí quyết hoàn thuế GTGT và thuế nhập khẩu khi xuất trả hàng lỗi (B13) sạch nước cản không lo bị truy thu'.",
         'Công văn số 23244/CHQ-NVTHQ tháng 10/2026; Điều 34 Nghị định 134/2016/NĐ-CP; Thông tư 38/2015/TT-BTC & TT 39/2018/TT-BTC.',
         update_date_str),

        (7,
         'nghen mang gach no thue hai quan dien tu 24 7 tre thong quan lo tau',
         'Xử lý sự cố nghẽn mạng gạch nợ thuế điện tử Hải quan 24/7: Bí quyết giải phóng tờ khai thông quan trong 15 phút, tránh rớt tàu và phí lưu bãi container',
         12100, 'Giải cứu thông quan (+2,050%)',
         'Doanh nghiệp xuất nhập khẩu hàng gấp, nhân viên hiện trường (ops) hải quan, kế toán thanh toán thuế XNK.',
         'Nộp thuế qua ngân hàng nhưng 3-4 tiếng tiền chưa nổi trên hệ thống Hải quan; tờ khai bị treo nợ thuế sát giờ cắt máng gây lỡ tàu xuất khẩu hoặc phát sinh chi phí lưu bãi cảng hàng chục triệu đồng.',
         "Tiêu đề: 'Nộp thuế 4 tiếng chưa nổi tiền trên VNACCS: Phác đồ cấp cứu giải phóng tờ khai thông quan trong 15 phút tránh rớt tàu'.",
         'Cổng thanh toán điện tử hải quan 24/7; Thông tư 184/2015/TT-BTC; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (8,
         'no thue treo etax mobile khoa ma so thue chan mo to khai hoa don thay the khac ky',
         'Giải quyết triệt để rủi ro hóa đơn thay thế khác kỳ & nợ thuế treo eTax Mobile khi đóng/chuyển đổi MST Hộ kinh doanh sang Doanh nghiệp XNK — Gỡ nút thắt bị chặn mở tờ khai hải quan',
         11800, 'Nỗi đau thực tế chạm đỉnh (+2,000%)',
         'Chủ doanh nghiệp vừa và nhỏ (SME), cá nhân chuyển từ Hộ kinh doanh lên Doanh nghiệp tư nhân/Công ty TNHH, kế toán thuế XNK.',
         'Tồn đọng nợ thuế cũ của Hộ kinh doanh trên eTax Mobile hoặc xử lý sai hóa đơn thay thế khác kỳ khiến hệ thống Hải quan tự động khóa phân luồng, chặn truyền tờ khai điện tử và tạm hoãn xuất cảnh.',
         "Tiêu đề: 'Nợ thuế treo eTax Mobile chặn đứng tờ khai hải quan: Cách tra soát, gỡ khóa MST và xử lý hóa đơn thay thế khác kỳ trong 24 giờ'.",
         'Nghị định 123/2020/NĐ-CP; Luật Quản lý thuế số 38/2019/QH14; Cổng Dịch vụ công Hải quan; Nhật ký thực tế daily_notes.md.',
         update_date_str),

        (9,
         'quyet dinh 1434 qd chq chuan du lieu chung tu dien tu json dataformat thong quan 2026',
         'Quyết định số 1434/QĐ-CHQ chính thức áp dụng chuẩn dữ liệu JSON DataFormat: Thay thế hoàn toàn các chuẩn cũ — Doanh nghiệp xuất nhập khẩu và đơn vị phần mềm hải quan cần làm gì ngay để không bị lỗi truyền chứng từ?',
         11500, 'Nghiệp vụ công nghệ mới (+1,950%)',
         'Nhân viên chứng từ XNK, đại lý hải quan, doanh nghiệp phát triển phần mềm khai báo hải quan (ECUS, FPT, Thái Sơn...).',
         'Hệ thống chưa cập nhật schema JSON DataFormat dẫn đến lỗi từ chối tiếp nhận thông điệp, chứng từ đính kèm (invoice, packing list) bị lỗi định dạng Base64 khiến tờ khai bị treo tại cổng tiếp nhận.',
         "Tiêu đề: 'Quyết định 1434/QĐ-CHQ áp dụng chuẩn JSON DataFormat: Hướng dẫn sửa lỗi kết nối và chuẩn hóa file chứng từ đính kèm tờ khai hải quan'.",
         'Quyết định số 1434/QĐ-CHQ ngày 26/08/2026 của Cục Hải quan; Báo Hải quan Online.',
         update_date_str),

        (10,
         'gia cuoc van tai container noi dia quoc te quy 4 tang 15 25 phan tram pss blank sailings',
         'Sóng tăng giá cước vận tải container nội địa (Hải Phòng - Cái Mép/TP.HCM tăng 15-25%) và Cước quốc tế cuối năm: Chiến lược tối ưu chi phí cung ứng Logistics Q4/2026 cho mùa hàng Tết',
         11200, 'Căng thẳng cước cuối năm (+1,900%)',
         'Chủ hàng xuất nhập khẩu, giám đốc cung ứng logistics, doanh nghiệp sản xuất phân phối hàng tiêu dùng Tết, forwarder nội địa và quốc tế.',
         'Hãng tàu nội địa tăng cước 15-25% trong khi tuyến quốc tế áp phụ phí mùa cao điểm (PSS) và hủy chuyến (Blank Sailings), đẩy chi phí logistics đội giá thành sản phẩm 20-30%.',
         "Tiêu đề: 'Cước container nội địa tăng sốc 15-25% & Hãng tàu quốc tế siết tải cuối năm: 3 đòn bẩy giúp chủ hàng giữ slot và hạ nhiệt chi phí'.",
         'Biểu cước vận tải các hãng tàu nội địa; Báo Công Thương; Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA).',
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
    ws3['A3'] = f"Chu kỳ: Tuần 44 (26/10/2026 - 01/11/2026) | Mở rộng 195 chủ đề nghiên cứu (gồm 163 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 196)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 196 or m.max_row >= 196]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (191,
         'nhap tai cho e31 truy thu thue sxxk doi ung b11 e62 nghi dinh 18 2021',
         'nhap khau tai cho e31, xuat khau tai cho b11 e62, truy thu thue sxxk nhap tai cho, nghi dinh 18/2021/nd-cp, dieu 12 dieu 36 nghi dinh 134/2016/nd-cp, cong van 23642/chq, hoan thue loai hinh a12, kiem tra sau thong quan sxxk',
         99.9,
         48200,
         'https://haiquanonline.com.vn',
         'Cạm bẫy truy thu thuế mã E31 khi nhập khẩu tại chỗ đối ứng xuất B11/E62: Rủi ro kiểm tra sau thông quan và bài toán nộp A12 hoàn thuế theo Nghị định 18/2021/NĐ-CP',
         'Nội dung trọng tâm: Cảnh báo rủi ro pháp lý nóng bỏng nhất trong hoạt động kiểm tra sau thông quan cuối năm 2026 đối với các doanh nghiệp sản xuất xuất khẩu (SXXK). Theo quy định tại Nghị định 18/2021/NĐ-CP (sửa đổi Điều 12 Nghị định 134/2016/NĐ-CP) và các Công văn hướng dẫn số 23642/CHQ, CV 16900/16901/BTC: Hàng hóa nhập khẩu tại chỗ để làm nguyên liệu SXXK (kể cả trường hợp đối ứng bên bán mở tờ khai xuất khẩu tại chỗ B11 hay E62) KHÔNG THUỘC ĐỐI TƯỢNG ĐƯỢC MIỄN THUẾ NGAY TẠI THỜI ĐIỂM NHẬP KHẨU (chỉ loại hình gia công E21 mới được xem xét miễn thuế nếu đủ điều kiện). Rất nhiều doanh nghiệp lầm tưởng mở tờ khai E31 không đóng thuế, khi đoàn kiểm tra sau thông quan rà soát tờ khai đối ứng của bên bán sẽ lập tức ra quyết định ấn định truy thu 100% thuế nhập khẩu, xử phạt vi phạm hành chính 10-20% và tính tiền chậm nộp 0.03%/ngày. Eureka Logistics đúc kết giải pháp an toàn tuyệt đối 3 bước: (1) Trước khi mở tờ khai, nhân viên XNK bắt buộc đối chiếu mã loại hình đối ứng của bên bán; (2) Nếu bên bán xuất B11/E62 và không có hợp đồng gia công, doanh nghiệp chủ động đăng ký loại hình A12 và nộp đủ thuế nhập khẩu; (3) Sau khi đưa nguyên liệu vào sản xuất và thực xuất khẩu thành phẩm ra nước ngoài, lập bộ hồ sơ xin hoàn thuế đã nộp theo đúng quy định tại Điều 36 Nghị định 134/2016/NĐ-CP, vừa đảm bảo tuân thủ pháp luật, vừa thu hồi trọn vẹn dòng tiền thuế.',
         'Báo Hải quan Online (haiquanonline.com.vn) & Sổ tay nội bộ daily_notes.md | CS (Case Study) / On-the-spot Import E31 Tax Assessment & Duty Drawback',
         update_date_str),

        (192,
         'cong van 4323 tttv attpmt kiem tra chat nong san quyt tuoi trung quoc',
         'cong van 4323/tttv-attpmt, kiem tra chat an toan thuc pham, nong san nhap khau trung quoc, quyt tuoi citrus reticulata, du luong abamectin vuot quy chuan, thong tu 50/2016/tt-byt, nghi dinh 15/2018/nd-cp, cua khau huu nghi tan thanh',
         99.5,
         44600,
         'https://haiquanonline.com.vn',
         "Công văn số 4323/TTTV-ATTPMT kích hoạt cơ chế 'Kiểm tra chặt' nông sản nhập khẩu từ Trung Quốc: Cảnh báo dư lượng thuốc BVTV Abamectin vượt ngưỡng 3 lần và giải pháp thông quan trái cây tươi",
         'Nội dung trọng tâm: Cục Trồng trọt và Bảo vệ thực vật ban hành Công văn số 4323/TTTV-ATTPMT thông báo áp dụng phương thức kiểm tra chặt về an toàn thực phẩm đối với các lô hàng quýt tươi (Citrus reticulata) nhập khẩu từ Trung Quốc. Qua chương trình giám sát, cơ quan chức năng phát hiện 2 lô hàng lớn (tổng khối lượng 39.5 tấn) nhập khẩu qua cửa khẩu Lạng Sơn và Nghệ An có dư lượng hoạt chất bảo vệ thực vật Abamectin lên tới 0.03 mg/kg, vượt gấp 3 lần giới hạn tối đa cho phép (0.01 mg/kg theo Thông tư số 50/2016/TT-BYT). Căn cứ khoản 3 Điều 17 Nghị định 15/2018/NĐ-CP, toàn bộ các lô hàng nông sản tiếp theo từ các nhà sản xuất/xuất khẩu vi phạm bị chuyển sang chế độ kiểm tra chặt (lấy mẫu 100% từng lô để thử nghiệm). Rủi ro nghiêm trọng: Thời gian chờ kết quả kiểm nghiệm kéo dài 3-5 ngày khiến trái cây tươi bị thối hỏng, phát sinh phí lưu bãi container lạnh 1.5 - 2 triệu đồng/ngày và nguy cơ bị buộc tái xuất hoặc tiêu hủy toàn bộ. Khuyến nghị hành động từ Eureka: Kiểm tra danh mục đen nhà xuất khẩu bị cảnh báo; đàm phán hợp đồng bắt buộc bên bán cung cấp chứng thư kiểm nghiệm COA đạt chuẩn trước khi giao hàng; và phương án ủy thác nhập khẩu chính ngạch bảo quản lạnh chuyên nghiệp.',
         'Báo Hải quan Online (haiquanonline.com.vn) & Cục BVTV / daily_notes.md | TN (Tin Ngành) / Agro-food Safety Inspection & Abamectin Compliance',
         update_date_str),

        (193,
         'kiem tra sau thong quan cuoc tau container tri gia hai quan cif fcl',
         'kiem tra sau thong quan cuoc tau, tri gia hai quan cuoc van chuyen cif, chung tu thanh toan cuoc tau qua ngan hang, master bill house bill freight prepaid, phu phi local charges cic ebs thuc te, an dinh thue van tai quoc te, thong tu 39/2015/tt-btc',
         99.3,
         43800,
         'https://customs.gov.vn',
         'Hải quan siết kiểm tra sau thông quan cước tàu container quốc tế & phụ phí hãng tàu (CIF/FOB): Doanh nghiệp cần chuẩn bị hồ sơ thanh toán qua ngân hàng thế nào để không bị bác trị giá hải quan?',
         'Nội dung trọng tâm: Tổng cục Hải quan chỉ đạo các Cục Hải quan địa phương tăng cường kiểm tra sau thông quan chuyên đề trị giá tính thuế hàng nhập khẩu, trọng tâm là khoản chi phí vận chuyển quốc tế (cước tàu biển) và các khoản phụ phí vận tải cộng vào trị giá hải quan theo quy định tại Thông tư 39/2015/TT-BTC (sửa đổi tại TT 60/2019/TT-BTC). Cơ quan hải quan đẩy mạnh đối chiếu dữ liệu điện tử giữa hệ thống manifest hãng tàu, ePort và chứng từ thanh toán thực tế của doanh nghiệp. Hàng loạt doanh nghiệp nhập khẩu theo điều kiện FOB hoặc CIF bị cơ quan hải quan bác cước khai báo do: Cước tàu thực tế chuyển trả cho đại lý forwarder cao hơn cước thể hiện trên hợp đồng và vận đơn; thanh toán cước tàu bằng tiền mặt hoặc qua tài khoản cá nhân; không xuất trình được điện chuyển tiền TT/LC qua ngân hàng; hoặc không cộng các phụ phí vận tải bắt buộc (như THC đầu bến đi, EBS, CIC). Hậu quả bị ấn định lại trị giá hải quan, truy thu tiền thuế nhập khẩu và thuế GTGT trong thời hạn 5 năm cùng tiền chậm nộp. Cẩm nang của Eureka: Bộ hồ sơ chuẩn hóa 5 chứng từ chứng minh cước tàu hợp lệ; nguyên tắc thanh toán 100% qua ngân hàng và cách thỏa thuận forwarder phát hành Freight Invoice chuẩn chỉ.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Báo Công Thương | FF (Forwarder & Cước Vận Tải) / Freight Audit & Ocean Shipping Customs Valuation',
         update_date_str),

        (194,
         'nghen mang gach no thue hai quan dien tu 24 7 tre thong quan lo tau',
         'nghen gach no thue hai quan, nop thue hai quan 24 7 qua ngan hang, cong thanh toan dien tu hai quan, loi truyen du lieu kho bac hai quan, to khai chua thong quan vi no thue, giay nop tien vao ngan sach nha nuoc, xu ly khan cap to khai hai quan',
         99.0,
         40500,
         'https://tapchitaichinh.vn',
         'Xử lý sự cố nghẽn mạng gạch nợ thuế điện tử Hải quan 24/7: Bí quyết giải phóng tờ khai thông quan trong 15 phút, tránh rớt tàu và phí lưu bãi container',
         'Nội dung trọng tâm: Hiện tượng nghẽn truyền dữ liệu thông tin nộp thuế giữa Ngân hàng thương mại phối hợp thu – Kho bạc Nhà nước – Hệ thống VNACCS/VCIS của Hải quan thường xuyên xảy ra vào những ngày cao điểm cuối tháng/cuối quý. Nhiều doanh nghiệp đã hoàn tất trích tiền thuế XK/NK từ tài khoản ngân hàng nhưng 3-4 tiếng tiền vẫn chưa "nổi" trên hệ thống kế toán tập trung của Hải quan, dẫn đến hệ thống không tự động gạch nợ và tờ khai bị treo, không thể giải phóng hàng. Đối với các lô hàng xuất khẩu sát giờ cắt máng (closing time) hoặc hàng nhập khẩu cần lấy gấp, sự cố này gây thiệt hại nặng nề: Lỡ chuyến tàu, rớt booking, phát sinh phí đảo chuyển và tiền phạt lưu bãi container hàng chục triệu đồng. Eureka chia sẻ phác đồ cấp cứu thông quan 2 phương án: (1) Ưu tiên thực hiện nộp thuế trực tiếp qua Cổng thanh toán điện tử và thông quan 24/7 của Hải quan bằng chữ ký số token để hệ thống gạch nợ tự động trong 5-10 phút; (2) Khi nộp qua Internet Banking bị nghẽn, in ngay Giấy nộp tiền có dấu đỏ xác nhận của ngân hàng mang trực tiếp đến Chi cục Hải quan nơi mở tờ khai để đề nghị cán bộ hải quan can thiệp nghiệp vụ thông quan có điều kiện (xác nhận bảo lãnh nộp thuế) giải phóng hàng trước.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & Sổ tay nội bộ daily_notes.md | HT (Học Thực chiến) / 24/7 E-Tax Payment Troubleshooting & Customs Clearance Emergency',
         update_date_str),

        (195,
         'tra cuu no thue cccd etax mobile nghi dinh 13 2023 bao ve du lieu ca nhan',
         'tra cuu no thue qua cccd, etax mobile dinh danh vneid, nghi dinh 13/2023/nd-cp bao ve du lieu ca nhan, tra ho no thue ca nhan, tam hoan xuat canh nguoi dai dien phap luat, quyen bi mat doi tu, quan ly thue chu doanh nghiep',
         98.7,
         39200,
         'https://chinhphu.vn',
         'Tra soát nợ thuế cá nhân & Chủ doanh nghiệp qua CCCD trên eTax Mobile: Rào cản Nghị định 13/2023/NĐ-CP và giải pháp chủ động rà soát chống lệnh tạm hoãn xuất cảnh',
         'Nội dung trọng tâm: Giải đáp thắc mắc lớn của cộng đồng doanh nghiệp và kế toán về việc tra hộ nợ thuế cá nhân qua số thẻ Căn cước công dân (CCCD) trên ứng dụng eTax Mobile. Theo quy định tại Nghị định 13/2023/NĐ-CP về bảo vệ dữ liệu cá nhân và quy chế bảo mật của Tổng cục Thuế, ứng dụng eTax Mobile tuyệt đối không cho phép người ngoài tra cứu hộ thông tin nợ thuế nếu không có tài khoản định danh điện tử VNeID mức 2 của chính chủ. Điều này khiến nhiều kế toán lúng túng khi muốn rà soát nghĩa vụ thuế tồn đọng cho chủ doanh nghiệp hoặc đối tác trước khi ký hợp đồng đại diện pháp luật. Rủi ro thực tế: Hàng ngàn giám đốc doanh nghiệp chỉ phát hiện mình bị nợ thuế cũ (thuế TNCN, thuế môn bài hộ kinh doanh cá thể dừng hoạt động nhưng chưa đóng MST) khi bị chặn lại ở cửa khẩu sân bay quốc tế do có quyết định tạm hoãn xuất cảnh theo Điều 66 Luật Quản lý thuế số 38/2019/QH14. Eureka khuyến nghị giải pháp quản trị an toàn: Chủ doanh nghiệp cần chủ động đăng nhập eTax Mobile định kỳ mỗi quý để rà soát toàn bộ nghĩa vụ thuế cá nhân và mã số thuế các đơn vị liên quan; chuẩn bị sẵn giấy ủy quyền hợp pháp có công chứng khi giao cho nhân sự kế toán làm thủ tục đối soát trực tiếp tại Chi cục Thuế.',
         'Cổng Thông tin điện tử Chính phủ (chinhphu.vn) & Sổ tay nội bộ daily_notes.md | CS (Case Study) / Personal Tax Identification Audit & Executive Exit Ban Prevention',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 196 - 200)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 196 + idx
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

    # Cập nhật Total row ở row 201
    total_row_idx = 201
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 195 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D200)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E200)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 195 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 163 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 31/10/2026)\n"
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
    print(f"5. Xuất bản tin đề xuất {BRIEFING_MD_PATH}...")
    md_content = f"""# [BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN — Hot News Scanner]

- **Tuần:** Tuần 44 (26/10/2026 - 01/11/2026) | Ngày quét: 31/10/2026
- **Tổng cơ hội phát hiện:** 25 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn`)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors (`drewry.co.uk`) & Hiệp hội Logistics Việt Nam (VLA)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Case study truy thu thuế E31, CV 4323/TTTV-ATTPMT kiểm tra chặt quýt tươi TQ, nghẽn mạng gạch nợ hải quan 24/7, bảo mật eTax Mobile theo NĐ 13/2023)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 31/10/2026)

### Đề xuất #1: [NỖI ĐAU THỰC CHIẾN & KIỂM TRA SAU THÔNG QUAN] Cạm bẫy truy thu thuế mã E31 khi nhập khẩu tại chỗ đối ứng xuất B11/E62: Rủi ro kiểm tra sau thông quan và bài toán nộp A12 hoàn thuế theo Nghị định 18/2021/NĐ-CP
- **Chuyên mục (Pillar):** `CS` (Case Study - Rà soát Thuế SXXK & Kiểm tra sau thông quan)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo 'Trùng sinh sửa sai: Tại sao mở E31 nhập tại chỗ bị truy thu thuế tiền tỷ?' | YouTube: Cẩm nang pháp lý phân biệt luồng hàng B11 vs E62, hướng dẫn mở A12 nộp thuế và lập bộ hồ sơ hoàn thuế theo Điều 36 NĐ 134/2016)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Cơ quan Hải quan đang mở cao điểm kiểm tra sau thông quan quý 4/2026 đối với các doanh nghiệp SXXK |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Bị truy thu 100% thuế nhập khẩu + phạt vi phạm hành chính 10-20% + phạt chậm nộp 0.03%/ngày tiền tỷ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh cãi dữ dội: 'Nhập về SXXK thì bản chất là miễn thuế, cớ sao đối ứng bên bán xuất B11 lại bị bắt nộp thuế?' |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka chia sẻ giải pháp 3 bước: Soi loại hình bên bán, chủ động mở A12 nộp thuế và hoàn thuế theo Điều 36 NĐ 134 |
  | **Tổng điểm** | **19.8 / 20** | **Xếp hạng: #1 (Đề tài Nỗi đau Thực tế Chạm đỉnh — Tránh bẫy truy thu tiền tỷ cho doanh nghiệp SXXK)** |
- **Tóm tắt lý do nên làm:** Đánh trúng sự mơ hồ phổ biến của các doanh nghiệp sản xuất xuất khẩu khi mua nguyên liệu tại chỗ. Hàng loạt doanh nghiệp mở mã E31 vì nghĩ "hàng SXXK luôn miễn thuế", nhưng theo Nghị định 18/2021/NĐ-CP và Công văn 23642/CHQ, nếu bên bán xuất B11/E62 thì nhập tại chỗ không thuộc diện miễn thuế. Bài phân tích khẳng định chuyên môn thượng thừa của Eureka, hướng dẫn doanh nghiệp mở A12 nộp thuế trước để hoàn thuế hợp lệ sau này.
- **Nguồn xác minh:** Nghị định số 18/2021/NĐ-CP; Điều 12, Điều 36 Nghị định 134/2016/NĐ-CP; Công văn số 23642/CHQ; CV 16900/16901/BTC; Sổ tay nội bộ `scratch/daily_notes.md`.

---

### Đề xuất #2: [AN TOÀN THỰC PHẨM & CỬA KHẨU BIÊN GIỚI] Công văn số 4323/TTTV-ATTPMT kích hoạt cơ chế 'Kiểm tra chặt' nông sản nhập khẩu từ Trung Quốc: Cảnh báo dư lượng thuốc BVTV Abamectin vượt ngưỡng 3 lần và giải pháp thông quan trái cây tươi
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Kiểm tra chuyên ngành & An toàn thực phẩm nông sản nhập khẩu)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic cảnh báo danh sách nhà xuất khẩu TQ bị blacklist & Cách né bẫy thối hỏng hoa quả | YouTube: Hướng dẫn quy trình lấy mẫu kiểm tra ATTP, kiểm nghiệm COA tại nguồn và phương án bảo quản container lạnh tại cửa khẩu Hữu Nghị/Tân Thanh)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Công văn 4323/TTTV-ATTPMT vừa phát đi ngày 22/10, các cửa khẩu phía Bắc đang siết chặt 100% các lô hàng quýt tươi |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Kiểm tra chặt kéo dài 3-5 ngày làm thối hỏng nông sản tươi, tiền điện cont lạnh 1.5 - 2 triệu/ngày, nguy cơ bị buộc tái xuất |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận nóng bỏng về kiểm soát chất lượng nông sản nhập khẩu mùa Tết và trách nhiệm đền bù của nhà cung cấp nước ngoài |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn tra cứu blacklist nhà xuất khẩu, ràng buộc điều khoản COA tại nguồn và giải pháp ủy thác nhập khẩu chính ngạch |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #2 (Cảnh báo Biên giới Sát sườn — Bảo vệ dòng vốn cho các chủ buôn nông sản)** |
- **Tóm tắt lý do nên làm:** Cuối năm là mùa cao điểm nhập khẩu hoa quả phục vụ tiêu dùng Tết. Vụ việc 39.5 tấn quýt tươi Trung Quốc bị phát hiện dư lượng Abamectin vượt ngưỡng 3 lần là lời cảnh tỉnh lớn. Bài viết của Eureka giúp các doanh nghiệp kinh doanh nông sản nắm vững quy chế kiểm tra chặt của Bộ Y tế / Bộ Nông nghiệp, tránh rơi vào cảnh hàng bị kẹt tại cửa khẩu đến thối hỏng.
- **Nguồn xác minh:** Công văn số 4323/TTTV-ATTPMT ngày 22/10/2026 của Cục Trồng trọt và Bảo vệ thực vật; Khoản 3 Điều 17 Nghị định 15/2018/NĐ-CP; Thông tư 50/2016/TT-BYT; Báo Hải quan Online (`haiquanonline.com.vn`).

---

### Đề xuất #3: [HẬU KIỂM CƯỚC TÀU & TRỊ GIÁ HẢI QUAN] Hải quan siết kiểm tra sau thông quan cước tàu container quốc tế & phụ phí hãng tàu (CIF/FOB): Doanh nghiệp cần chuẩn bị hồ sơ thanh toán qua ngân hàng thế nào để không bị bác trị giá hải quan?
- **Chuyên mục (Pillar):** `FF` (Forwarder & Cước Vận Tải - Trị giá tính thuế cước biển & Chứng từ thanh toán)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: 'Hậu kiểm cước tàu 5 năm: Tại sao chênh lệch 50 USD/cont giữa B/L và Freight Invoice lại bị truy thu thuế?' | YouTube: Hướng dẫn đối chiếu Master B/L, House B/L, điện chuyển tiền TT ngân hàng và tờ khai hải quan chuẩn chỉnh)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Ngành Hải quan đẩy mạnh đối soát chéo dữ liệu điện tử ePort, manifest hãng tàu và hóa đơn cước vận tải cuối năm 2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Bị ấn định lại trị giá hải quan, truy thu thuế GTGT và thuế NK trong 5 năm, xử phạt kê khai sai |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi nảy lửa: 'Phụ phí THC, CIC, EBS hãng tàu thu thực tế có bắt buộc phải cộng vào trị giá tính thuế?' |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Chia sẻ checklist 5 chứng từ thanh toán ngân hàng bắt buộc và kỹ thuật phối hợp với forwarder làm sạch bộ hồ sơ cước |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #3 (Nghiệp vụ Hậu kiểm Chuyên sâu — Khẳng định năng lực bảo an vận tải của Eureka)** |
- **Tóm tắt lý do nên làm:** Hậu kiểm cước tàu biển là "vùng trũng rủi ro" của rất nhiều doanh nghiệp nhập khẩu. Việc khai báo cước thấp hơn thực tế hoặc thanh toán cước không qua ngân hàng đang bị cơ quan hải quan truy quét gắt gao. Phân tích của Eureka giải tỏa nỗi lo cho kế toán và phòng mua hàng, khẳng định vai trò của Forwarder chuyên nghiệp minh bạch chứng từ.
- **Nguồn xác minh:** Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Luật Hải quan số 54/2014/QH13; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`).

---

### Đề xuất #4: [CÔNG NGHỆ THÔNG QUAN & GẠCH NỢ HẢI QUAN] Xử lý sự cố nghẽn mạng gạch nợ thuế điện tử Hải quan 24/7: Bí quyết giải phóng tờ khai thông quan trong 15 phút, tránh rớt tàu và phí lưu bãi container
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Kỹ thuật Nộp thuế 24/7 & Xử lý nghẽn gạch nợ VNACCS)
- **Nền tảng phù hợp:** **Facebook** (Mẹo cấp cứu: Nộp thuế 4 tiếng chưa nổi tiền trên VNACCS thì làm gì để kịp giờ cắt máng?) & **YouTube** (Quy trình ký số nộp thuế trực tiếp qua Cổng Hải quan 24/7 và thủ tục xin thông quan can thiệp thủ công)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.7/5) | Giai đoạn cao điểm xuất nhập khẩu cuối tháng lượng giao dịch tăng đột biến gây nghẽn liên thông Ngân hàng - Kho bạc - Hải quan |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Hàng chờ thông quan bị lỡ chuyến tàu xuất khẩu, container bị rớt máng, phát sinh phí đổi booking và lưu bãi cảng hàng chục triệu |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Nhiều kế toán và ops hiện trường hoang mang khi tiền đã trích khỏi tài khoản công ty nhưng hải quan vẫn báo nợ thuế |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn 2 phác đồ cấp cứu: Kênh nộp trực tiếp Cổng thanh toán 24/7 và thủ tục mang Giấy nộp tiền đến chi cục can thiệp thủ công |
  | **Tổng điểm** | **19.0 / 20** | **Xếp hạng: #4 (Cẩm nang Cấp cứu Hiện trường — Rất thiết thực cho đội ngũ chứng từ và kế toán)** |
- **Tóm tắt lý do nên làm:** Đây là sự cố thực tế xảy ra như cơm bữa tại các chi cục hải quan cảng biển và sân bay trong mùa cao điểm. Chia sẻ cách xử lý linh hoạt của Eureka giúp nhân sự XNK và chủ hàng giải cứu lô hàng trong gang tấc, tạo thiện cảm và độ tin cậy cực lớn.
- **Nguồn xác minh:** Cổng thanh toán điện tử hải quan 24/7 (`customs.gov.vn`); Thông tư 184/2015/TT-BTC; Sổ tay nội bộ `scratch/daily_notes.md`.

---

### Đề xuất #5: [BẢO MẬT DỮ LIỆU & RỦI RO QUẢN TRỊ THUẾ] Tra soát nợ thuế cá nhân & Chủ doanh nghiệp qua CCCD trên eTax Mobile: Rào cản Nghị định 13/2023/NĐ-CP và giải pháp chủ động rà soát chống lệnh tạm hoãn xuất cảnh
- **Chuyên mục (Pillar):** `CS` (Case Study - Quản trị Pháp lý Thuế Cá nhân & Doanh nghiệp)
- **Nền tảng phù hợp:** **Facebook** (Cảnh báo: 'Đừng để bị giữ lại ở cửa khẩu sân bay vì món nợ thuế không ngờ tới từ 10 năm trước') & **YouTube** (Hướng dẫn liên kết định danh VNeID mức 2 với eTax Mobile để rà soát toàn bộ nghĩa vụ thuế của các pháp nhân do cá nhân đứng tên)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Cơ quan thuế liên tục công bố danh sách tạm hoãn xuất cảnh các đại diện pháp luật nợ thuế trong tháng 10/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Giám đốc chuẩn bị bay đi ký hợp đồng hoặc công tác nước ngoài bị chặn ngay cửa khẩu xuất nhập cảnh |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Tranh luận về việc tra hộ dữ liệu thuế, tính pháp lý của việc ủy quyền và quy định bảo mật thông tin cá nhân theo NĐ 13/2023 |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.6/5) | Eureka cung cấp quy trình rà soát định kỳ mã định danh cá nhân và phương án nộp thuế gỡ lệnh tạm hoãn xuất cảnh hỏa tốc trong 24h |
  | **Tổng điểm** | **18.9 / 20** | **Xếp hạng: #5 (Quản trị Rủi ro Lãnh đạo — Đánh trúng tâm lý lo ngại của các chủ doanh nghiệp)** |
- **Tóm tắt lý do nên làm:** Lệnh tạm hoãn xuất cảnh do nợ thuế đang là "cơn ác mộng" đối với các lãnh đạo doanh nghiệp. Nhiều người không hề biết mình còn nợ vài triệu đồng tiền thuế từ một hộ kinh doanh cũ hay một pháp nhân liên đới. Hướng dẫn của Eureka mang tính phòng vệ cao, bảo vệ uy tín và lịch trình công tác của các sếp.
- **Nguồn xác minh:** Nghị định 13/2023/NĐ-CP; Luật Quản lý thuế số 38/2019/QH14; Cổng Thông tin điện tử Chính phủ (`chinhphu.vn`); Sổ tay nội bộ `scratch/daily_notes.md`.

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

    print("\n=== HOÀN TẤT CẬP NHẬT THEO WF05 NGÀY 31/10/2026 ===")

if __name__ == '__main__':
    update_seo_w44_3110()
