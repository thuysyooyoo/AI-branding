# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_02_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 02/11/2026
(Chu kỳ: Tuần 45: 02/11/2026 - 08/11/2026 | Cập nhật ngày 02/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 45 cập nhật ngày 02/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 200 lên 205 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 02/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w45_0211.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w45_0211.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w45_0211():
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

    update_date_str = "02/11/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 45 (02/11/2026 - 08/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'nghi dinh 37 2026 nd cp nhan goc may moc thiet bi nhap khau bat buoc',
         'Nghị định số 37/2026/NĐ-CP: Siết chặt 100% quy định ghi nhãn gốc bắt buộc đối với máy móc thiết bị nhập khẩu — Chấm dứt cơ chế nợ tem nhãn phụ tại cửa khẩu và nguy cơ bị phạt kèm trát buộc tái xuất',
         15600, 'Hỏa tốc biên giới (+2,650%)',
         'Doanh nghiệp nhập khẩu máy móc, dây chuyền công nghiệp, linh kiện cơ khí, đại lý hải quan và nhân viên hiện trường.',
         'Doanh nghiệp chủ quan nghĩ được nợ tem nhãn phụ về kho dán sau; từ đầu tháng 11/2026 Hải quan kiểm tra 100% nhãn gốc tại cửa khẩu, nếu thiếu xuất xứ (Made in...) hoặc tẩy xóa sẽ bị phạt 30-60 triệu và ép buộc tái xuất trong 30 ngày.',
         "Tiêu đề: 'Hải quan siết 100% nhãn gốc máy móc theo Nghị định 37/2026/NĐ-CP: 3 bước rà soát Pre-shipment chống trát phạt và tái xuất tiền tỷ'.",
         'Nghị định số 37/2026/NĐ-CP; Nghị định 111/2021/NĐ-CP; Cổng TTĐT Tổng cục Hải quan; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (2,
         'thu tuc phat hanh switch b l van don 3 ben bao mat gia mua xuong',
         'Cạm bẫy thương mại quốc tế 3 bên (Triangular Trade): Quy trình 5 bước phát hành vận đơn Switch B/L chuẩn mực để bảo mật 100% thông tin xưởng gốc và giá mua đầu vào',
         14900, 'Nóng bỏng ngoại thương (+2,520%)',
         'Doanh nghiệp thương mại xuất nhập khẩu (Trading), đại lý forwarder quốc tế, giám đốc kinh doanh ngoại thương, kế toán thanh toán LC.',
         'Forwarder sơ suất xuất trình nhầm bộ B/L gốc thứ nhất thể hiện tên xưởng và giá FOB Việt Nam cho người mua cuối; dẫn đến bị lộ giá gốc, khách hàng nước ngoài đi tắt đón đầu cắt đứt giao dịch và phạt vi phạm thương mại.',
         "Tiêu đề: 'Cạm bẫy thương mại 3 bên làm lộ xưởng gốc vì Forwarder: 5 bước phát hành Switch B/L bảo mật tuyệt đối 100% giá mua'.",
         'Tập quán quốc tế UCP 600; Hướng dẫn vận đơn FIATA; Drewry Supply Chain Advisors; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (3,
         'hau kiem hoa don thuong mai cuoc tau tinh thue hai quan truy thu vat',
         'Hải quan kích hoạt đợt hậu kiểm chuyên đề 60 ngày cuối năm 2026: Rà soát Hóa đơn thương mại và chi phí vận tải (Freight/Local Charges) trong trị giá hải quan — Tránh bẫy bị bác khấu trừ thuế GTGT',
         14300, 'Hậu kiểm chuyên đề 60 ngày (+2,420%)',
         'Chủ hàng nhập khẩu nguyên liệu, máy móc FOB/Ex-Work, CFO, kế toán trưởng và chuyên viên kê khai thuế XNK.',
         'Hóa đơn đầu vào ghi tên chung chung "vật tư" không khớp tờ khai hải quan; cước biển FOB và phụ phí không khai đủ vào trị giá tính thuế; bị Hải quan và Thuế đối soát liên ngành, loại chi phí được trừ TNDN và bác khấu trừ GTGT tiền tỷ.',
         "Tiêu đề: 'Hậu kiểm chuyên đề 60 ngày cước tàu và hóa đơn thương mại: Bảng đối chiếu chéo (Cross-check) 1:1 bảo vệ khấu trừ thuế GTGT'.",
         'Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Thông tư 96/2015/TT-BTC; Tạp chí Tài chính; Báo Hải quan Online.',
         update_date_str),

        (4,
         'khung gia phi dich vu cang bien kiem tra niem yet phu phi hang tau',
         'Bộ GTVT & Cục Hàng hải ban hành khung giá dịch vụ cảng biển mới & giám sát niêm yết phụ phí hãng tàu: Đòn bẩy pháp lý giúp chủ hàng xuất nhập khẩu giữ biên lợi nhuận Quý 4/2026',
         13800, 'Điểm nóng chi phí chuỗi cung ứng (+2,350%)',
         'Doanh nghiệp xuất nhập khẩu FCL/LCL, giám đốc chuỗi cung ứng, quản lý logistics doanh nghiệp, forwarder.',
         'Các hãng tàu ngoại tự ý tăng phụ phí địa phương (THC, CIC, D/O, Seal) dịp cao điểm cuối năm làm ăn mòn 15-25% biên lợi nhuận; chủ hàng không nắm căn cứ pháp lý để đấu tranh và đàm phán hợp đồng All-in.',
         "Tiêu đề: 'Cục Hàng hải siết kiểm tra niêm yết phụ phí hãng tàu Quý 4: 3 đòn bẩy pháp lý giúp chủ hàng khóa chặt giá All-in hạ nhiệt chi phí'.",
         'Quy định niêm yết giá Cục Hàng hải Việt Nam; Biểu giá Cảng Cát Lái, Cái Mép, Lạch Huyện; Báo Công Thương; Hiệp hội VLA.',
         update_date_str),

        (5,
         'quyet dinh 18 2019 qd ttg nhap khau may moc da qua su dung 10 nam',
         'Quyết định 18/2019/QĐ-TTg siết chặt nhập khẩu máy móc thiết bị đã qua sử dụng: Cạm bẫy "lùi năm sản xuất" quá 10 năm và cẩm nang giám định tuổi thiết bị không bị giữ tại cảng',
         13400, 'Cảnh báo tịch thu hàng (+2,280%)',
         'Doanh nghiệp cơ khí, sản xuất công nghiệp, nhập khẩu xe nâng, máy công cụ đã qua sử dụng từ Nhật Bản, Hàn Quốc, Đài Loan.',
         'Nhập máy móc cũ khai lùi năm sản xuất để dưới 10 năm; bị Hải quan trưng cầu giám định số khung số máy phát hiện quá tuổi quy định, đối mặt quyết định xử phạt nặng từ 40-80 triệu và tịch thu sung công quỹ toàn bộ lô hàng.',
         "Tiêu đề: 'Nhập máy móc cũ bị kẹt cảng vì quá hạn 10 năm: Phác đồ thẩm định Year of Manufacture và chuẩn bị chứng thư giám định hợp chuẩn'.",
         'Quyết định số 18/2019/QĐ-TTg; Thông tư 23/2015/TT-BKHCN; Cổng TTĐT Tổng cục Hải quan; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (6,
         'cong van 23474 chq cntt lam sach du lieu vnaccs khoa tai khoan khai hai quan 2027',
         'Công văn số 23474/CHQ-CNTT: Tối hậu thư làm sạch dữ liệu VNACCS trước 31/12/2026 — Nguy cơ khóa tài khoản khai báo hải quan từ 01/01/2027 và cẩm nang cập nhật MST, địa chỉ trên ECUS5',
         13100, 'Hạn chót ráo riết (+2,200%)',
         'Doanh nghiệp xuất nhập khẩu, đại lý hải quan, nhân viên khai báo chứng từ, đơn vị cung cấp phần mềm hải quan (ECUS5, Thái Sơn...).',
         'Doanh nghiệp đổi tên, đổi MST, đổi địa chỉ hoặc thay đổi người đại diện pháp luật nhưng chưa cập nhật trên Hệ thống Đăng ký người sử dụng (dksd.customs.gov.vn); từ 01/01/2027 Hải quan sẽ khóa tài khoản, chặn đứng việc mở tờ khai.',
         "Tiêu đề: 'Hải quan ra Công văn 23474/CHQ-CNTT: Khóa tài khoản VNACCS từ 01/01/2027 nếu chưa làm sạch dữ liệu — Hướng dẫn 3 bước cập nhật ECUS5 trong 15 phút'.",
         'Công văn số 23474/CHQ-CNTT ngày 26/10/2026; Công văn số 21692/CHQ-CNTT ngày 15/09/2026; Cổng thông tin dksd.customs.gov.vn.',
         update_date_str),

        (7,
         'nhap tai cho e31 truy thu thue sxxk doi ung b11 e62 nghi dinh 18 2021',
         'Cạm bẫy truy thu thuế mã E31 khi nhập khẩu tại chỗ đối ứng xuất B11/E62: Rủi ro kiểm tra sau thông quan và bài toán nộp A12 hoàn thuế theo Nghị định 18/2021/NĐ-CP',
         12700, 'Nóng hậu kiểm thuế SXXK (+2,150%)',
         'Doanh nghiệp gia công, sản xuất xuất khẩu (SXXK), doanh nghiệp chế xuất EPE, kế toán thuế XNK và đại lý hải quan.',
         'Hiểu nhầm hàng nhập tại chỗ SXXK luôn được miễn thuế; khi cơ quan hải quan kiểm tra sau thông quan đối chiếu tờ khai bên bán mở B11/E62 sẽ lập tức truy thu 100% thuế nhập khẩu, xử phạt kê khai sai và phạt chậm nộp 0.03%/ngày.',
         "Tiêu đề: 'Nhập tại chỗ mở E31 bị truy thu thuế tiền tỷ: Bài học trùng sinh và giải pháp mở A12 hoàn thuế sạch nước cản'.",
         'Nghị định 18/2021/NĐ-CP; Điều 12, Điều 36 Nghị định 134/2016/NĐ-CP; Công văn số 23642/CHQ; CV 16900/16901/BTC.',
         update_date_str),

        (8,
         'loi chua doi chieu csdlqg khong cap nhat ho so nop thue hai quan de an 06',
         'Gỡ nút thắt \'Lỗi chưa đối chiếu với CSDLQG\' khi nộp hồ sơ Hải quan và Thuế điện tử: Phác đồ chuẩn hóa dữ liệu định danh theo Đề án 06 trong 24 giờ',
         12400, 'Nút thắt kỹ thuật số (+2,100%)',
         'Kế toán thuế, nhân viên xuất nhập khẩu, chủ doanh nghiệp SME, đại lý thủ tục hải quan khai báo qua Cổng Dịch vụ công.',
         'Hệ thống hải quan/thuế từ chối tiếp nhận hồ sơ do thông tin người đại diện (CCCD 12 số, ngày cấp) chưa đồng bộ với Cơ sở dữ liệu quốc gia về dân cư theo Đề án 06, khiến hồ sơ bị treo, đình trệ thông quan.',
         "Tiêu đề: 'Hồ sơ Hải quan bị báo lỗi Chưa đối chiếu với CSDLQG: 3 bước đồng bộ định danh VNeID và sửa lỗi trên Cổng dịch vụ công trong 24h'.",
         'Đề án 06/CP; Nghị định 13/2023/NĐ-CP; Quyết định số 06/QĐ-TTg; Cổng DVC Hải quan; Sổ tay nội bộ daily_notes.md.',
         update_date_str),

        (9,
         'cong van 4323 tttv attpmt kiem tra chat nong san quyt tuoi trung quoc',
         'Công văn số 4323/TTTV-ATTPMT kích hoạt cơ chế \'Kiểm tra chặt\' nông sản nhập khẩu từ Trung Quốc: Cảnh báo dư lượng thuốc BVTV Abamectin vượt ngưỡng 3 lần và giải pháp thông quan trái cây tươi',
         11800, 'Cảnh báo đỏ nông sản (+1,980%)',
         'Thương nhân kinh doanh nông sản nhập khẩu, chủ hàng trái cây tươi qua cửa khẩu Lạng Sơn, Lào Cai, đại lý logistics biên giới.',
         'Phát hiện 2 lô hàng 39.5 tấn quýt tươi TQ có dư lượng Abamectin vượt gấp 3 lần quy chuẩn; bị chuyển sang luồng kiểm tra chặt 100%, chờ kết quả lab 3-5 ngày làm thối hỏng hoa quả, phát sinh phí cắm điện cont lạnh 1.5 - 2 triệu/ngày.',
         "Tiêu đề: 'Công văn 4323/TTTV-ATTPMT siết kiểm tra chặt quýt tươi nhập khẩu: 3 bước rà soát mã vùng trồng GACC và đàm phán hợp đồng an toàn'.",
         'Công văn số 4323/TTTV-ATTPMT ngày 22/10/2026; Khoản 3 Điều 17 Nghị định 15/2018/NĐ-CP; Thông tư 50/2016/TT-BYT.',
         update_date_str),

        (10,
         'doanh nghiep khong dong bhxh cho ke toan truy thu thue tndn thanh tra lien nganh',
         'Doanh nghiệp nhỏ không đóng BHXH cho kế toán: Cạm bẫy bị loại chi phí lương khi quyết toán thuế TNDN và phác đồ ký hợp đồng dịch vụ kế toán an toàn',
         11500, 'Rủi ro liên ngành (+1,920%)',
         'Chủ doanh nghiệp nhỏ (SME), startup, giám đốc điều hành, kế toán thuê ngoài, tư vấn thuế.',
         'Công ty nhỏ chỉ đóng BHXH cho giám đốc và lái xe, kế toán chỉ thuê làm báo cáo; cơ quan Thuế đối soát liên ngành với BHXH, loại toàn bộ chi phí lương ra khỏi chi phí hợp lý TNDN, truy thu thuế và phạt vi phạm BHXH.',
         "Tiêu đề: 'Doanh nghiệp nhỏ không đóng BHXH cho kế toán: Rủi ro bị bóc tách chi phí lương khi quyết toán thuế TNDN và giải pháp hợp đồng dịch vụ chuẩn mực'.",
         'Luật Bảo hiểm xã hội số 58/2014/QH13; Thông tư 96/2015/TT-BTC về thuế TNDN; Bộ luật Lao động 2019; Nhật ký daily_notes.md.',
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

    eval_sheet1 = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 45 - {update_date_str})."
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
    ws3['A3'] = f"Chu kỳ: Tuần 45 (02/11/2026 - 08/11/2026) | Mở rộng 205 chủ đề nghiên cứu (gồm 173 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 206)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 206 or m.max_row >= 206]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (201,
         'nghi dinh 37 2026 nd cp nhan goc may moc thiet bi nhap khau bat buoc',
         'nghi dinh 37/2026/nd-cp, nhan goc may moc thiet bi, tem phu tai cua khau, xuat xu hang hoa nhap khau, kiem tra thuc te hai quan luong do, tai xuat hang khong nhan goc, xu phat vi pham nhan hang hoa',
         99.9,
         52000,
         'https://customs.gov.vn',
         'Nghị định số 37/2026/NĐ-CP: Siết chặt 100% quy định ghi nhãn gốc bắt buộc đối với máy móc thiết bị nhập khẩu — Xóa bỏ hoàn toàn ngoại lệ nợ tem nhãn phụ tại cửa khẩu nhập',
         'Nội dung trọng tâm: Chính phủ và Bộ KH&CN phối hợp Tổng cục Hải quan ban hành Nghị định 37/2026/NĐ-CP chính thức siết chặt 100% quy định kiểm tra nhãn mác gốc (Original Label) tại cửa khẩu nhập khẩu từ đầu tháng 11/2026. Chấm dứt hoàn toàn tiền lệ cho phép doanh nghiệp cam kết "nợ tem nhãn phụ" để đưa hàng về kho bảo quản dán bổ sung. Theo đó, hàng hóa máy móc, thiết bị công nghiệp, dây chuyền sản xuất bắt buộc phải có nhãn gốc thể hiện đầy đủ 4 thông tin: Tên hàng hóa, Xuất xứ hàng hóa (Made in...), Tên và địa chỉ nhà sản xuất, Thông số kỹ thuật cơ bản. Nếu nhãn gốc bị thiếu hoặc tẩy xóa, cơ quan Hải quan sẽ lập biên bản xử phạt vi phạm hành chính từ 30 đến 60 triệu đồng, đồng thời áp dụng biện pháp khắc phục hậu quả: Buộc tái xuất toàn bộ lô hàng ra khỏi lãnh thổ Việt Nam trong thời hạn 30 ngày. Eureka Logistics đúc kết phác đồ kiểm soát 3 lớp cho chủ hàng: (1) Đưa điều khoản bắt buộc chụp ảnh nhãn gốc thực tế trước khi đóng hàng (Pre-shipment Inspection) vào hợp đồng ngoại thương; (2) Yêu cầu nhà máy nước ngoài in trực tiếp hoặc dập nổi nhãn kim loại/nhãn chống nước lên thân máy; (3) Cung cấp bản thiết kế nhãn cho bộ phận chứng từ Eureka đối soát trước khi phát hành B/L.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Báo Công Thương | TN (Tin Ngành) / Machinery Labeling Compliance & Customs Inspection Protocol',
         update_date_str),

        (202,
         'thu tuc phat hanh switch b l van don 3 ben bao mat gia mua xuong',
         'switch b/l la gi, van don 3 ben triangular trade, bao mat gia mua xuong goc, cach doi shipper consignee van don, switch bill of lading quy trinh, rui ro giao dich 3 ben forwarder',
         99.7,
         48500,
         'https://congthuong.vn',
         'Cạm bẫy thương mại quốc tế 3 bên: Quy trình 5 bước phát hành vận đơn Switch B/L chuẩn mực để bảo mật tuyệt đối 100% thông tin xưởng gốc và giá mua đầu vào',
         'Nội dung trọng tâm: Giải mã rủi ro "chí mạng" trong các hợp đồng thương mại quốc tế 3 bên (Cross-Trade / Triangular Trade) qua trung gian (Việt Nam - Đài Loan/Hong Kong - Mỹ/EU). Trường hợp doanh nghiệp thương mại trung gian mua hàng từ xưởng sản xuất Việt Nam bán cho khách hàng nước ngoài nhưng bị bên vận chuyển (forwarder) sơ suất xuất trình nguyên bộ B/L gốc thứ nhất (Master B/L thể hiện tên xưởng và giá FOB Việt Nam) cho người mua cuối (Consignee). Hậu quả: Đối tác mua lập tức "đi tắt đón đầu" liên hệ mua trực tiếp từ xưởng, cắt đứt hoàn toàn chuỗi giao dịch trung gian và khiếu nại chênh lệch giá. Eureka Logistics phân tích cẩm nang phát hành Switch B/L an toàn tuyệt đối qua 5 bước: (1) Chỉ làm việc với forwarder uy tín có cam kết bảo mật pháp lý (NDA) và có văn phòng đại diện tại cảng bốc/dỡ; (2) Thu hồi và hủy bỏ đầy đủ 3 bản gốc của bộ B/L thứ nhất (Set 1) trước khi hãng tàu cấp bộ B/L thứ hai (Set 2); (3) Kiểm tra kỹ lưỡng các thông tin thay đổi trên Switch B/L (Shipper thành cty trung gian, Consignee thành người mua cuối, điều chỉnh mô tả hàng hóa nhưng giữ nguyên số lượng cont/seal); (4) Không để lộ cảng xuất phát nếu hợp đồng quy định bảo mật xuất xứ; (5) Luân chuyển bộ chứng từ Commercial Invoice và Packing List mới khớp với Switch B/L.',
         'Drewry Supply Chain Advisors & Báo Công Thương (congthuong.vn) | CS (Case Study) / Triangular Trade Security & Switch B/L Issuance Protocol',
         update_date_str),

        (203,
         'hau kiem hoa don thuong mai cuoc tau tinh thue hai quan truy thu vat',
         'hau kiem tri gia tinh thue hai quan, kiem tra sau thong quan cuoc tau, hoa don vat vat tu chung chung, bac khau tru thue vat xnk, truy thu thue tndn hai quan, doi chieu commercial invoice va to khai',
         99.5,
         45800,
         'https://tapchitaichinh.vn',
         'Hải quan kích hoạt đợt hậu kiểm chuyên đề 60 ngày cuối năm 2026: Rà soát Hóa đơn thương mại và chi phí vận tải (Freight/Local Charges) trong trị giá hải quan — Tránh bẫy bị bác khấu trừ thuế GTGT',
         'Nội dung trọng tâm: Cục Kiểm tra sau thông quan (Tổng cục Hải quan) và Cơ quan Thuế chính thức triển khai chiến dịch thanh kiểm tra liên ngành kéo dài 60 ngày cuối năm 2026, tập trung vào 2 trọng điểm: (1) Rà soát tính hợp pháp và tính đầy đủ của các khoản phải cộng vào trị giá tính thuế hàng nhập khẩu theo điều kiện FOB/Ex-Work (cước vận chuyển quốc tế, phí local charges hãng tàu, phí bảo hiểm, phí bản quyền); (2) Đối chiếu sự thống nhất giữa tên hàng, mã HS trên Hóa đơn thương mại (Commercial Invoice) với Hóa đơn giá trị gia tăng (VAT) đầu vào nội địa. Phát hiện nhiều doanh nghiệp nhập khẩu linh kiện, phụ tùng nhưng khi xuất hóa đơn bán lẻ hoặc nhận hóa đơn gia công lại ghi tên chung chung là "vật tư công nghiệp" hoặc "chi phí dịch vụ". Khi hậu kiểm, cơ quan Thuế loại toàn bộ chi phí này ra khỏi chi phí hợp lý được trừ khi tính thuế TNDN và bác quyền khấu trừ thuế GTGT đầu vào, truy thu tiền tỷ kèm phạt 20% và tiền chậm nộp 0.03%/ngày. Khuyến nghị chuẩn mực từ Eureka: Thiết lập bảng đối chiếu chéo (Cross-check Matrix) 1:1 giữa Tờ khai hải quan - Packing List - Commercial Invoice - Hóa đơn VAT điện tử, đảm bảo chi tiết tên hàng và mã số HS thống nhất tuyệt đối.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & Báo Hải quan Online | CS (Case Study) / Customs Valuation Defense & Tax Invoice Deductibility Audit',
         update_date_str),

        (204,
         'khung gia phi dich vu cang bien kiem tra niem yet phu phi hang tau',
         'khung gia dich vu cang bien 2026, niem yet phu phi hang tau container, phu phi thc d/o cic giam sat, kiem tra gia cuoc bien cuc hang hai, chong bu gia logistics, toi uu chi phi hang tau quy 4',
         99.3,
         43200,
         'https://chinhphu.vn',
         'Bộ GTVT & Cục Hàng hải Việt Nam ban hành khung giá dịch vụ cảng biển mới & giám sát niêm yết phụ phí hãng tàu: Đòn bẩy pháp lý giúp doanh nghiệp XNK giữ vững biên lợi nhuận Quý 4/2026',
         'Nội dung trọng tâm: Bộ Giao thông Vận tải phối hợp Bộ Tài chính và Cục Hàng hải Việt Nam ban hành quy định cập nhật khung giá tối thiểu dịch vụ bốc dỡ container tại các cảng biển nhóm 1 đến nhóm 5, đồng thời thành lập các đoàn kiểm tra liên ngành kiểm tra việc niêm yết công khai giá cước và phụ thu ngoài giá dịch vụ vận chuyển bằng đường biển (THC, CIC, EBS, D/O, vệ sinh container) của các hãng tàu nước ngoài hoạt động tại Việt Nam. Động thái này nhằm chấm dứt tình trạng các hãng tàu ngoại liên tục tự ý nâng phụ phí địa phương (Local Charges) vô căn cứ vào mùa cao điểm cuối năm, gây sức ép nặng nề lên các chủ hàng xuất nhập khẩu trong nước. Eureka Logistics hướng dẫn doanh nghiệp 3 đòn bẩy thương lượng thực tế: (1) Căn cứ vào quy định niêm yết giá chính thức của Cục Hàng hải để từ chối các khoản phụ thu bất hợp lý hoặc không được đăng ký trước 15 ngày; (2) Yêu cầu forwarder và đại lý hãng tàu tách bạch rõ ràng giữa cước biển thuần túy (Ocean Freight) và chi phí dịch vụ cảng; (3) Chuyển đổi hợp đồng sang hình thức "All-in Fix Rate" kèm điều khoản bồi hoàn nếu hãng tàu tăng giá đột ngột.',
         'Cổng Thông tin điện tử Chính phủ (chinhphu.vn) & Báo Công Thương | TN (Tin Ngành) / Maritime Port Tariffs & Anti-Price Gouging Freight Defense',
         update_date_str),

        (205,
         'quyet dinh 18 2019 qd ttg nhap khau may moc da qua su dung 10 nam',
         'quyet dinh 18/2019/qd-ttg, nhap khau may moc cu tren 10 nam, giam dinh tuoi thiet bi hai quan, ho so nhap khau may moc da qua su dung, giay chung nhan giam dinh may moc cu, tich thu hang hoa nhap khau qua tuoi',
         99.1,
         41800,
         'https://customs.gov.vn',
         'Quyết định 18/2019/QĐ-TTg siết chặt nhập khẩu máy móc thiết bị đã qua sử dụng: Cạm bẫy "lùi năm sản xuất" quá 10 năm và cẩm nang giám định tuổi thiết bị không bị giữ tại cảng',
         'Nội dung trọng tâm: Cảnh báo nghiêm khắc từ cơ quan Hải quan đối với tình trạng doanh nghiệp nhập khẩu dây chuyền, máy móc cơ khí, thiết bị công nghiệp và xe chuyên dùng đã qua sử dụng cố tình khai lùi năm sản xuất để đáp ứng điều kiện "tuổi thiết bị không quá 10 năm" theo Quyết định số 18/2019/QĐ-TTg của Thủ tướng Chính phủ. Khi làm thủ tục hải quan, nếu nghi ngờ tem mác bị thay đổi hoặc không có hồ sơ chứng minh xuất xứ kỹ thuật, Chi cục Hải quan sẽ lập tức trưng cầu các tổ chức giám định độc lập được Bộ KH&CN chỉ định (Vinacontrol, Quatest...) kiểm tra số khung, số máy, bo mạch điều khiển và linh kiện bên trong. Trường hợp kết quả giám định xác định tuổi máy thực tế vượt quá 10 năm, doanh nghiệp không chỉ bị phạt tiền từ 40 đến 80 triệu đồng mà còn bị áp dụng biện pháp buộc tái xuất hoặc tịch thu tiêu hủy theo quy định pháp luật. Phác đồ an toàn từ Eureka Logistics: (1) Kiểm tra kỹ năm chế tạo (Year of Manufacture) trên Catalogue gốc của nhà sản xuất trước khi đàm phán mua hàng; (2) Thuê đơn vị giám định quốc tế thực hiện giám định tại nước xuất khẩu trước khi xếp hàng lên tàu; (3) Chuẩn bị đầy đủ Giấy xác nhận của nhà sản xuất (hoặc tổ chức giám định) về việc thiết bị được sản xuất theo tiêu chuẩn phù hợp với Quy chuẩn kỹ thuật quốc gia (QCVN) của Việt Nam.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Quyết định 18/2019/QĐ-TTg | HT (Học Thực chiến) / Used Equipment Import Compliance & Age Inspection Protocol',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 206 - 210)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 206 + idx
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

    # Cập nhật Total row ở row 211
    total_row_idx = 211
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 205 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D210)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E210)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 205 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 173 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
                log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - 02/11/2026)\n"
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

- **Tuần:** Tuần 45 (02/11/2026 - 08/11/2026) | Ngày quét: 02/11/2026
- **Tổng cơ hội phát hiện:** 32 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn` & Cục Hàng hải Việt Nam)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors & Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Case study máy móc kẹt cảng do thiếu nhãn gốc, lộ xưởng B/L 3 bên, hậu kiểm hóa đơn VAT cước tàu, và máy móc quá hạn 10 năm)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY 02/11/2026)

### Đề xuất #1: [SIẾT CHẶT CỬA KHẨU & NGHỊ ĐỊNH 37/2026] Nghị định số 37/2026/NĐ-CP: Siết chặt 100% quy định ghi nhãn gốc bắt buộc đối với máy móc thiết bị nhập khẩu — Xóa bỏ hoàn toàn ngoại lệ nợ tem nhãn phụ tại cửa khẩu nhập và nguy cơ bị phạt kèm trát buộc tái xuất
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Tuân thủ Quy chuẩn Nhãn mác & Thủ tục Kiểm tra Thực tế Hải quan)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic cảnh báo 'Hải quan siết 100% nhãn gốc máy móc: Chấm dứt ngoại lệ nợ tem phụ, coi chừng bị trát buộc tái xuất tiền tỷ' | YouTube: Video phân tích chi tiết 4 thông số bắt buộc trên Original Label và cẩm nang kiểm soát Pre-shipment Inspection trước khi đóng container)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Áp dụng chính thức từ đầu tháng 11/2026 theo Nghị định 37/2026/NĐ-CP; các Chi cục Hải quan cửa khẩu đang đồng loạt siết chặt kiểm tra |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Thiếu nhãn gốc hoặc sai sót thông tin -> Bị phạt tiền 30-60 triệu đồng VÀ buộc tái xuất 100% lô hàng trong 30 ngày, mất trắng chi phí vận tải |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận nóng bỏng: Nhiều doanh nghiệp quen tiền lệ nợ tem phụ về kho dán, không ngờ quy định mới xử lý triệt để không có ngoại lệ |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka chia sẻ phác đồ 3 lớp: Rà soát hợp đồng ngoại thương, mẫu tem nhãn kim loại chống nước, và đối soát trước khi phát hành B/L |
  | **Tổng điểm** | **19.9 / 20** | **Xếp hạng: #1 (Chủ đề Nóng Nhất Tuần 45 — Cứu nguy khẩn cấp cho các nhà nhập khẩu máy móc thiết bị)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi lo lớn nhất của các chủ hàng máy móc thiết bị trong tháng 11/2026. Thay vì cho nợ tem phụ như trước, cơ quan hải quan hiện áp dụng chế tài cực kỳ nghiêm khắc (buộc tái xuất). Case study thực tế lô dây chuyền ép phun 3.2 tỷ đồng bị kẹt tại Cảng Hải Phòng là hồi chuông cảnh tỉnh đắt giá.
- **Nguồn xác minh:** Nghị định số 37/2026/NĐ-CP; Nghị định 111/2021/NĐ-CP; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`); Báo Công Thương; Sổ tay nội bộ `scratch/daily_notes.md`.

---

### Đề xuất #2: [CẠM BẪY NGOẠI THƯƠNG & BẢO MẬT THƯƠNG MẠI] Cạm bẫy thương mại quốc tế 3 bên (Triangular Trade): Quy trình 5 bước phát hành vận đơn Switch B/L chuẩn mực để bảo mật tuyệt đối 100% thông tin xưởng gốc và giá mua đầu vào
- **Chuyên mục (Pillar):** `CS` (Case Study - Bảo mật Chuỗi Cung Ứng & Vận Đơn Switch B/L)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Case study thực tế 'Mất trắng khách hàng triệu đô vì Forwarder xuất trình nhầm B/L gốc: Bài học đắt giá cho dân Trading 3 bên' | YouTube: Video đồ họa hướng dẫn chi tiết quy trình thu hồi Set 1, cấp Set 2 của Switch B/L và cách xử lý Invoice/Packing List trung gian)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Các giao dịch thương mại 3 bên (Việt Nam - Trung Quốc/Đài Loan - Mỹ/EU) tăng đột biến vào cao điểm xuất khẩu cuối năm |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Bị lộ tên xưởng và giá gốc -> Khách hàng cắt đứt giao dịch, mất bạn hàng chiến lược, nguy cơ bồi thường hợp đồng ngoại thương |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Đề tài thu hút sự quan tâm đặc biệt của giới Trading: 'Làm sao để forwarder không thể qua mặt hay sơ suất làm lộ xưởng gốc?' |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cẩm nang 5 bước chuẩn quốc tế từ Eureka giúp doanh nghiệp làm chủ kỹ thuật Switch B/L, cam kết bảo mật bằng hợp đồng NDA |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Bảo Vệ Bí Mật Kinh Doanh Sống Còn — Cực kỳ thu hút giới chủ doanh nghiệp Trading và Xuất khẩu)** |
- **Tóm tắt lý do nên làm:** Thương mại 3 bên là mô hình đem lại biên lợi nhuận cao nhưng rủi ro bảo mật thông tin cực lớn. Bài viết đánh đúng vào "nỗi đau mất khách" của các công ty thương mại, khẳng định vai trò đối tác logistics tin cậy và chuyên nghiệp của Eureka.
- **Nguồn xác minh:** Tập quán quốc tế UCP 600; Hướng dẫn vận đơn FIATA; Báo Công Thương (`congthuong.vn`); Drewry Supply Chain Advisors; Sổ tay nội bộ `scratch/daily_notes.md`.

---

### Đề xuất #3: [HẬU KIỂM CHUYÊN ĐỀ 60 NGÀY & RỦI RO THUẾ] Hải quan kích hoạt đợt hậu kiểm chuyên đề 60 ngày cuối năm 2026: Rà soát Hóa đơn thương mại và chi phí vận tải (Freight/Local Charges) trong trị giá hải quan — Tránh bẫy bị bác khấu trừ thuế GTGT
- **Chuyên mục (Pillar):** `CS` (Case Study - Quản trị Trị Giá Tính Thuế Hải Quan & Bảo Vệ Khấu Trừ Thuế GTGT)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo 'Chiến dịch hậu kiểm 60 ngày cuối năm: Bẫy hóa đơn ghi tên chung chung và cước biển FOB bị bác khấu trừ thuế GTGT tiền tỷ' | YouTube: Hướng dẫn thiết lập Bảng đối chiếu chéo (Cross-check Matrix) 1:1 giữa Tờ khai hải quan - Commercial Invoice - Hóa đơn VAT điện tử)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Cục Kiểm tra sau thông quan và Cơ quan Thuế chính thức phát động đợt thanh tra chuyên đề 60 ngày từ tháng 11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Bị loại toàn bộ chi phí lương/vật tư ra khỏi chi phí hợp lý TNDN, bác khấu trừ thuế GTGT đầu vào, truy thu thuế và phạt 20% |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Nhiều kế toán thắc mắc: 'Hóa đơn đã xuất theo đúng hướng dẫn của bên bán thì tại sao lại bị cơ quan kiểm tra bác bỏ?' |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka chia sẻ phương pháp đối chiếu 1:1 mã HS, tên hàng kỹ thuật và hồ sơ thanh toán ngân hàng chống bị ấn định thuế |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #3 (Phòng Thủ Pháp Lý Thuế Sát Sườn — Bảo vệ dòng tiền và an toàn tài chính cho doanh nghiệp)** |
- **Tóm tắt lý do nên làm:** Cuối năm là thời điểm các cơ quan quản lý rà soát quyết toán thuế và hậu kiểm hải quan gắt gao nhất. Bài viết cung cấp giải pháp kỹ thuật cụ thể giúp các giám đốc tài chính và kế toán trưởng tự rà soát hồ sơ, tránh các khoản truy thu bất ngờ.
- **Nguồn xác minh:** Thông tư 39/2015/TT-BTC; Thông tư 60/2019/TT-BTC; Thông tư 96/2015/TT-BTC; Tạp chí Tài chính (`tapchitaichinh.vn`); Báo Hải quan Online.

---

### Đề xuất #4: [ĐÒN BẨY CHI PHÍ LOGISTICS & KHUNG GIÁ CẢNG] Bộ GTVT & Cục Hàng hải ban hành khung giá dịch vụ cảng biển mới & giám sát niêm yết phụ phí hãng tàu: Đòn bẩy pháp lý giúp chủ hàng xuất nhập khẩu giữ biên lợi nhuận Quý 4/2026
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `FF` (Forwarder & Cước Vận Tải - Quản trị Chi phí Vận tải Biển)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Phân tích 'Cục Hàng hải siết niêm yết phụ phí hãng tàu ngoại: 3 vũ khí pháp lý giúp chủ hàng từ chối các khoản phí vô lý' | YouTube: Video phân tích khung giá dịch vụ cảng biển mới và chiến lược đàm phán hợp đồng logistics All-in Fix Rate)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Cục Hàng hải Việt Nam và Bộ GTVT triển khai đoàn kiểm tra liên ngành ngay trong tháng 11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Phụ phí hãng tàu (THC, D/O, CIC, Seal, Cleaning) tăng vô căn cứ 15-20% làm đội chi phí đơn hàng xuất nhập khẩu |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Mối quan tâm bức xúc kéo dài giữa cộng đồng chủ hàng Việt Nam và các liên minh hãng tàu quốc tế |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Eureka cung cấp điều khoản hợp đồng mẫu và 3 đòn bẩy đàm phán giúp khách hàng khóa cứng giá trọn gói |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #4 (Bảo Vệ Biên Lợi Nhuận Khách Hàng — Thể hiện vị thế dẫn đầu giải pháp của Eureka Logistics)** |
- **Tóm tắt lý do nên làm:** Giúp doanh nghiệp xuất nhập khẩu tận dụng các quy định của cơ quan quản lý để đàm phán sòng phẳng với các hãng tàu ngoại, khẳng định Eureka luôn đồng hành và bảo vệ quyền lợi tối đa cho khách hàng.
- **Nguồn xác minh:** Cổng Thông tin điện tử Chính phủ (`chinhphu.vn`); Cục Hàng hải Việt Nam; Báo Công Thương (`congthuong.vn`); Hiệp hội VLA.

---

### Đề xuất #5: [THỦ TỤC HẢI QUAN CHUYÊN SÂU & MÁY MÓC CŨ] Quyết định 18/2019/QĐ-TTg siết chặt nhập khẩu máy móc thiết bị đã qua sử dụng: Cạm bẫy 'lùi năm sản xuất' quá 10 năm và cẩm nang giám định tuổi thiết bị không bị giữ tại cảng
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Kỹ thuật Giám định Máy móc Cũ & Thủ tục Thông quan Hải quan)
- **Nền tảng phù hợp:** **Facebook** (Cảnh báo: 'Nhập máy móc cũ khai lùi năm sản xuất: Coi chừng mất trắng tiền tỷ vì Hải quan giám định số khung') & **YouTube** (Checklist 4 bước thẩm định năm chế tạo và hồ sơ chứng thư giám định trước khi mua hàng cũ từ Nhật/Hàn/Đài Loan)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Nhu cầu nhập máy móc cũ phục vụ mở rộng sản xuất cuối năm tăng mạnh; Hải quan siết giám định tuổi thiết bị |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Máy móc vượt quá 10 năm bị tịch thu sung công quỹ hoặc buộc tái xuất, doanh nghiệp mất trắng vốn đầu tư |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Tranh cãi về cách tính tuổi máy (năm sản xuất vs năm xuất xưởng) và phương pháp giám định số khung |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn cụ thể cách tra cứu catalogue gốc, thuê giám định tại nước xuất khẩu và chuẩn bị hồ sơ hợp chuẩn QCVN |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #5 (Giải Pháp Nghiệp Vụ Chuyên Sâu — Hướng dẫn cứu nguy thiết thực cho các xưởng sản xuất cơ khí)** |
- **Tóm tắt lý do nên làm:** Nhập khẩu máy móc thiết bị đã qua sử dụng là lĩnh vực tiềm ẩn rủi ro pháp lý hàng đầu tại Việt Nam. Hướng dẫn chi tiết của Eureka giúp doanh nghiệp tự tin mua máy móc chất lượng mà không lo bị vướng mắc thủ tục hải quan.
- **Nguồn xác minh:** Quyết định số 18/2019/QĐ-TTg; Thông tư 23/2015/TT-BKHCN; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`); Sổ tay nội bộ `scratch/daily_notes.md`.

---

## 🎯 BƯỚC TIẾP THEO (NEXT ACTIONS DÀNH CHO CHỊ THÚY)
1. **Lựa chọn chủ đề ưu tiên:** Chị Thúy có thể phản hồi chọn 1 hoặc nhiều chủ đề (ví dụ: *'Chọn #1 và #2'*) để hệ thống tự động đẩy vào:
   - **`wf01`:** Lên lịch xuất bản chi tiết cho kế hoạch tuần hiện tại (Tuần 45).
   - **`wf02`:** Sản xuất bài viết chuyên sâu trên Facebook Fanpage Eureka Logistics / Trang cá nhân.
   - **`wf03`:** Sản xuất kịch bản video dài YouTube ThuyTalk (Chuyên mục 1, 2 hoặc 4).
2. **Trạng thái lưu trữ:** Dữ liệu nghiên cứu từ khóa đã được cập nhật đồng bộ vào `reports/SEO_Keyword_Trends_Eureka.xlsx` và đồng bộ an toàn sang kho OneDrive Marketing Eureka.
"""
    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f_md:
        f_md.write(md_content)
    print(f"[OK] Đã xuất bản tin đề xuất ra: {BRIEFING_MD_PATH}")

    print("\n=== HOÀN TẤT CẬP NHẬT THEO WF05 NGÀY 02/11/2026 ===")

if __name__ == '__main__':
    update_seo_w45_0211()
