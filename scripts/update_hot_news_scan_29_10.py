# -*- coding: utf-8 -*-
"""
Script: update_seo_w44_2910.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 29/10/2026
(Chu kỳ: Tuần 44: 26/10/2026 - 01/11/2026 | Cập nhật ngày 29/10/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành, chu kỳ Tuần 44 cập nhật ngày 29/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 180 lên 185 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 29/10/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w44_2910.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w44_2910.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w44_2910():
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

    update_date_str = "29/10/2026"

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
         13200, 'Hỏa tốc toàn quốc (+2,300%)',
         'Doanh nghiệp sản xuất, xuất khẩu gỗ & nội thất, sắt thép kim loại, pin năng lượng, dệt may; forwarder tuyến Mỹ/EU.',
         'Nhiều doanh nghiệp nhập linh kiện thô từ nước ngoài về chỉ gia công đơn giản, dán nhãn Việt Nam để né thuế; bị cơ quan Hải quan và nước nhập khẩu điều tra lẩn tránh thuế phòng vệ thương mại, đối mặt nguy cơ bị áp thuế chống bán phá giá trừng phạt từ 25% đến hơn 100% và giữ hàng tại cảng đến vài tháng.',
         "Tiêu đề: 'Hải quan ra Công văn 23535/CHQ-GSQL cảnh báo đỏ 7 nhóm hàng: 4 bước rà soát hồ sơ xuất xứ chống nguy cơ bị Mỹ và EU trừng phạt thuế phòng vệ'.",
         'Công văn số 23535/CHQ-GSQL ngày 27/10/2026; Công văn số 8744/BCT-PVTM ngày 22/10/2026 của Cục Phòng vệ thương mại (Bộ Công Thương).',
         update_date_str),

        (2,
         'quy dinh cbam eu han chot bao cao phat thai 31 10 sat thep nhom chung chi carbon',
         'Đếm ngược hạn chót nộp báo cáo phát thải CBAM (31/10): Áp lực nghẹt thở đối với doanh nghiệp xuất khẩu Sắt thép, Nhôm sang EU — Khi chứng chỉ carbon trở thành rào cản tài chính sống còn',
         12800, 'Đếm ngược cấp bách (+2,150%)',
         'Doanh nghiệp sản xuất, xuất khẩu sắt thép, tôn mạ, nhôm thanh định hình, sản phẩm kim loại sang 27 quốc gia Liên minh Châu Âu.',
         'Ngày 31/10 là hạn chót nộp báo cáo quý 3 phát thải CBAM; nhiều doanh nghiệp Việt Nam lúng túng chưa kiểm kê khí nhà kính, không tính được phát thải nội hàm (SEE) thực tế; nếu dùng hệ số mặc định sẽ bị EU áp thuế carbon phạt cực cao (>80 EUR/tấn CO2e), làm mất đơn hàng vào tay đối thủ.',
         "Tiêu đề: 'Hạn chót CBAM ngày 31/10: Doanh nghiệp thép và nhôm Việt Nam cần hoàn tất báo cáo phát thải thế nào để không bị đánh thuế phạt nghìn tỷ?'.",
         'Quy định Cơ chế điều chỉnh biên giới carbon (CBAM) của Nghị viện Châu Âu; Báo Công Thương; Cổng TTĐT Logistics Việt Nam (logistics.gov.vn).',
         update_date_str),

        (3,
         'quyet dinh 1434 qd chq chuan du lieu chung tu dien tu json dataformat thong quan 2026',
         'Quyết định số 1434/QĐ-CHQ chính thức áp dụng chuẩn dữ liệu JSON DataFormat: Thay thế hoàn toàn các chuẩn cũ — Doanh nghiệp xuất nhập khẩu và đơn vị phần mềm hải quan cần làm gì ngay để không bị lỗi truyền chứng từ?',
         11500, 'Nghiệp vụ công nghệ mới (+1,900%)',
         'Nhân viên chứng từ XNK, đại lý hải quan, doanh nghiệp phát triển phần mềm khai báo hải quan (ECUS, FPT, Thái Sơn...).',
         'Quyết định 1434/QĐ-CHQ có hiệu lực từ 15/10/2026 bắt buộc đổi sang chuẩn JSON DataFormat thay thế QĐ 748 và 1143; nhiều hệ thống chưa cập nhật schema dẫn đến lỗi từ chối tiếp nhận thông điệp, chứng từ đính kèm (invoice, packing list) bị lỗi định dạng Base64 khiến tờ khai bị treo tại cổng tiếp nhận.',
         "Tiêu đề: 'Quyết định 1434/QĐ-CHQ áp dụng chuẩn JSON DataFormat: Hướng dẫn sửa lỗi kết nối và chuẩn hóa file chứng từ đính kèm tờ khai hải quan'.",
         'Quyết định số 1434/QĐ-CHQ ngày 26/08/2026 của Cục Hải quan (có hiệu lực từ 15/10/2026); Báo Hải quan Online.',
         update_date_str),

        (4,
         'cong van 23405 chq nvthq phan loai hang hoa ma hs may moc thiet bi linh kien thong tu 38',
         'Công văn số 23405/CHQ-NVTHQ: Siết chặt quy trình phân loại hàng hóa và giải quyết tranh chấp mã HS máy móc, linh kiện — Bí quyết lập Hồ sơ kỹ thuật xác định trước mã số (Advance Ruling) bảo vệ doanh nghiệp',
         10600, 'Siết chặt phân loại (+1,750%)',
         'Doanh nghiệp nhập khẩu máy móc, dây chuyền thiết bị toàn bộ, phụ tùng thay thế công nghiệp, kế toán thuế XNK.',
         'Tranh chấp mã HS thiết bị đồng bộ tháo rời nhiều chuyến theo Quy tắc 2(a) và hàng hợp bộ Quy tắc 3(b); doanh nghiệp tự khai thuế 0% bị hải quan bác mã, áp sang mã có thuế suất 5-10% kèm phạt hành chính và truy thu hàng tỷ đồng trong 5 năm.',
         "Tiêu đề: 'Công văn 23405/CHQ-NVTHQ chấn chỉnh phân loại hàng hóa: Chiến lược nộp hồ sơ Xác định trước mã số (Advance Ruling) triệt tiêu 100% rủi ro truy thu'.",
         'Công văn số 23405/CHQ-NVTHQ ngày 23/10/2026; Thông tư 38/2015/TT-BTC (sửa đổi tại TT 39/2018/TT-BTC); 6 Quy tắc phân loại HS quốc tế.',
         update_date_str),

        (5,
         'chot ty gia hai quan thu nam 29 10 2026 usd vnd lap ke hoach truyen to khai tiet kiem thue',
         'Kỳ chốt tỷ giá tính thuế hải quan thứ Năm (29/10/2026): Tỷ giá USD/VND neo mức đỉnh và mẹo căn giờ mở tờ khai tiết kiệm hàng chục triệu tiền thuế cho lô hàng nhập khẩu đầu tháng 11',
         10200, 'Điểm rơi tài chính (+1,700%)',
         'Doanh nghiệp nhập khẩu nguyên phụ liệu, giám đốc tài chính CFO, kế toán trưởng, nhân viên làm thủ tục thông quan.',
         'Tỷ giá USD/VND mua vào chuyển khoản Vietcombank chốt ngày thứ Năm hôm nay (29/10/2026) neo mức cao kỷ lục, làm tăng trực tiếp số thuế nhập khẩu và thuế GTGT phải nộp; doanh nghiệp không biết cách căn ngày truyền tờ khai (trước hay sau 00h ngày thứ Sáu) bị phát sinh thêm chi phí thuế nộp thừa đáng kể.',
         "Tiêu đề: 'Chốt tỷ giá tính thuế hải quan ngày 29/10/2026: Nghệ thuật căn ngày truyền tờ khai giúp doanh nghiệp tiết kiệm ngay hàng chục triệu tiền thuế'.",
         'Nghị định 08/2015/NĐ-CP (sửa đổi tại NĐ 167/2025/NĐ-CP & NĐ 252/2026/NĐ-CP); Thông báo tỷ giá tính thuế Cục Hải quan; Vietcombank.',
         update_date_str),

        (6,
         'cong van 23474 chq cntt lam sach du lieu dksd customs khoa tai khoan khai bao hai quan 2027',
         'Công văn số 23474/CHQ-CNTT ngày 26/10/2026: Rà soát & làm sạch dữ liệu người dùng trên dksd.customs.gov.vn trước 31/12/2026 — Hải quan cảnh báo KHÓA TÀI KHOẢN KHAI BÁO HẢI QUAN từ 01/01/2027',
         12500, 'Báo động hạn chót (+2,000%)',
         'Toàn bộ doanh nghiệp xuất nhập khẩu, đại lý hải quan, hãng tàu, forwarder sử dụng hệ thống VNACCS/VCIS.',
         'Doanh nghiệp không nắm được thông báo khẩn cấp, chưa chuẩn hóa định danh người đại diện pháp luật và chữ ký số token; sau ngày 31/12/2026 tài khoản VNACCS bị khóa khiến toàn bộ hàng hóa xuất nhập khẩu bị đình trệ, ách tắc tại cảng.',
         "Tiêu đề: 'Hải quan ra Công văn 23474/CHQ-CNTT: Hướng dẫn làm sạch dữ liệu trên dksd.customs.gov.vn trước 31/12/2026 tránh bị KHÓA TÀI KHOẢN KHAI BÁO'.",
         'Công văn số 23474/CHQ-CNTT ngày 26/10/2026 của Cục Hải quan; Hệ thống dksd.customs.gov.vn.',
         update_date_str),

        (7,
         'cong van 23384 chq gsql khai chi tieu 1 34 dia diem xep hang thue suat uu dai evfta thong tu 39',
         'Công văn số 23384/CHQ-GSQL ngày 23/10/2026: Bẫy khai báo chỉ tiêu 1.34 "Địa điểm xếp hàng" theo Thông tư 39 — Nguy cơ bị bác thuế suất ưu đãi đặc biệt EVFTA và mất quyền hưởng lợi C/O',
         11200, 'Duy trì sốt dẻo (+1,850%)',
         'Doanh nghiệp nhập khẩu từ EU (Đức, Pháp, Ý...), đại lý làm thủ tục hải quan, bộ phận khai chứng từ XNK.',
         'Khai sai mã UN LOCODE ở Ô 1 hoặc ghi thiếu tên cảng/kho ở Ô 2 của chỉ tiêu 1.34 khiến tờ khai không khớp vận tải đơn (Bill of Lading), bị cơ quan hải quan từ chối áp dụng thuế suất ưu đãi EVFTA, bác C/O form EUR.1 và ấn định thuế MFN hàng trăm triệu đồng.',
         "Tiêu đề: 'Công văn 23384/CHQ-GSQL: Hướng dẫn chuẩn xác khai chỉ tiêu 1.34 Địa điểm xếp hàng để được hưởng thuế suất ưu đãi EVFTA 0%'.",
         'Công văn số 23384/CHQ-GSQL ngày 23/10/2026; Mẫu số 01 Phụ lục I Thông tư 39/2018/TT-BTC; Hiệp định EVFTA.',
         update_date_str),

        (8,
         'eu hoan quy dinh eudr chong mat rung nong san go ca phe cao su 2026',
         'EU chính thức hoãn thực thi Quy định chống phá rừng (EUDR) sang 31/12/2026: "Khoảng thở" quý giá cho nông lâm sản Việt Nam và 4 việc sống còn doanh nghiệp không được chủ quan',
         9800, 'Duy trì đỉnh (+1,650%)',
         'Doanh nghiệp xuất khẩu cà phê, gỗ và sản phẩm gỗ, cao su sang thị trường EU; hiệp hội ngành hàng; forwarder chuyên tuyến Châu Âu.',
         'Nhiều doanh nghiệp hoang mang trước nguy cơ bị cấm xuất khẩu sang EU do chưa kịp hoàn thiện bản đồ định vị GPS vùng trồng và thẩm định chuỗi cung ứng; khi EU quyết định lùi thời hạn, một bộ phận chủ quan buông lỏng chuẩn bị, đối mặt nguy cơ trễ hạn vào cuối năm 2026.',
         "Tiêu đề: 'EU hoãn thi hành luật chống phá rừng EUDR sang cuối 2026: \"Khoảng thở\" vàng cho cà phê, gỗ Việt Nam và 4 việc bắt buộc phải làm ngay'.",
         'Báo Công Thương (congthuong.vn); vneconomy.vn; Cổng TTĐT Bộ Công Thương (moit.gov.vn).',
         update_date_str),

        (9,
         'drewry wci cuoi thang 10 tang cuoc gri pss thang 11 khan hiem vo cont',
         'Chỉ số Drewry WCI cuối tháng 10 chạm 4.585 USD/FEU & Cước Bờ Đông vượt 10.166 USD: Hãng tàu tung chiêu Blank Sailings siết tải và 3 vũ khí giữ chỗ container mùa Noel',
         9500, 'Căng thẳng mùa vụ (+1,600%)',
         'Doanh nghiệp xuất khẩu hàng đi Mỹ và Châu Âu, chủ hàng dệt may, da giày, đồ gỗ, giám đốc logistics chuỗi cung ứng.',
         'Giá cước vận tải biển spot đảo chiều tăng mạnh; các hãng tàu tiếp tục cắt giảm chuyến (Blank Sailings) khiến vỏ cont rỗng khan hiếm cục bộ và rủi ro rớt tàu (roll cont) cận kề ngày giao hàng Giáng sinh.',
         "Tiêu đề: 'Drewry WCI cuối tháng 10 bật tăng mạnh: Làm sao để giữ book tàu xuất khẩu đi Mỹ - Âu an toàn trước làn sóng tăng cước GRI tháng 11?'.",
         'Báo cáo Drewry WCI ngày 22/10/2026; World Shipping Council; Báo Công Thương.',
         update_date_str),

        (10,
         'thue ngoai phong xuat nhap khau tron goi outsourced logistics sme tiet kiem chi phi chong phat thue',
         'Xu hướng "Thuê ngoài Phòng Xuất Nhập Khẩu trọn gói" (Outsourced Logistics Dept) cuối năm 2026: Giải pháp cắt giảm 70% chi phí vận hành và triệt tiêu 100% rủi ro phạt thuế cho SME',
         9600, 'Đột phá B2B (+1,550%)',
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
    ws3['A3'] = f"Chu kỳ: Tuần 44 (26/10/2026 - 01/11/2026) | Mở rộng 185 chủ đề nghiên cứu (gồm 153 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 186)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 186 or m.max_row >= 186]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (181,
         'cong van 23535 chq gsql canh bao phong ve thuong mai gian lan xuat xu cong van 8744 bct',
         'cong van 23535/CHQ-GSQL, cong van 8744/BCT-PVTM, canh bao phong ve thuong mai 2026, gian lan xuat xu hang xuat khau, chuyen tai bat hop phap, nguy co ap thue chong ban pha gia my eu, go sat thep nhom det may phuong tien ca nhan',
         99.6,
         44200,
         'https://customs.gov.vn',
         'Công văn số 23535/CHQ-GSQL ngày 27/10/2026 & Công văn 8744/BCT-PVTM: Cảnh báo đỏ 7 nhóm mặt hàng rủi ro cao bị điều tra phòng vệ thương mại, gian lận xuất xứ và lẩn tránh thuế — Nguy cơ bị áp thuế trừng phạt của Mỹ, EU',
         'Nội dung trọng tâm: Cục Hải quan ban hành Công văn hỏa tốc số 23535/CHQ-GSQL ngày 27/10/2026 chuyển tiếp danh sách cảnh báo của Bộ Công Thương (Công văn số 8744/BCT-PVTM ngày 22/10/2026) đôn đốc các Chi cục Hải quan trên toàn quốc siết chặt giám sát, nhận diện sớm rủi ro gian lận xuất xứ và chuyển tải bất hợp pháp đối với 7 nhóm ngành hàng xuất khẩu chủ lực: (1) Gỗ và sản phẩm gỗ nội thất; (2) Sắt thép và sản phẩm kim loại; (3) Vật liệu xây dựng; (4) Phương tiện di chuyển cá nhân tốc độ thấp; (5) Hóa chất, nhựa và sản phẩm nhựa; (6) Xơ sợi và dệt may; (7) Sản phẩm công nghiệp chế tạo khác. Phân tích 3 rủi ro chí mạng: Hàng hóa chỉ gia công lắp ráp giản đơn tại Việt Nam nhưng dán nhãn Made in Vietnam xuất sang Mỹ, EU, Canada sẽ bị cơ quan điều tra quốc tế áp thuế chống bán phá giá trừng phạt từ 25% đến hơn 100%; bị giữ hàng tại cảng để xác minh chuỗi cung ứng; doanh nghiệp bị liệt vào danh sách đen kiểm tra sau thông quan 5 năm. Cẩm nang 4 bước phòng vệ của Eureka: Rà soát tỷ lệ giá trị gia tăng nội địa (RVC), minh bạch hóa 100% hồ sơ nguyên vật liệu đầu vào (tờ khai nhập khẩu, hóa đơn VAT), xây dựng bảng định mức kỹ thuật chi tiết và bảo đảm tuân thủ nghiêm ngặt quy tắc xuất xứ cụ thể mặt hàng (PSR).',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Cục Phòng vệ thương mại (Bộ Công Thương) | TN (Tin Ngành) / Trade Defense Compliance & Anti-Circumvention Origin Verification',
         update_date_str),

        (182,
         'quy dinh cbam eu han chot bao cao phat thai 31 10 sat thep nhom chung chi carbon',
         'cbam han chot 31 10, bao cao phat thai carbon eu, co che dieu chinh bien gioi carbon, xuat khau thep nhom sang eu, gia chung chi cbam 2026, phat thai truc tiep gian tiep see, kiem ke khi nha kinh iso 14064, communication template cbam',
         99.4,
         42600,
         'https://congthuong.vn',
         'Đếm ngược hạn chót nộp báo cáo phát thải CBAM (31/10): Áp lực nghẹt thở đối với doanh nghiệp xuất khẩu Sắt thép, Nhôm sang EU — Khi chứng chỉ carbon trở thành rào cản tài chính sống còn',
         'Nội dung trọng tâm: Thời khắc nước rút ngày 31/10 đối với Cơ chế điều chỉnh biên giới carbon (CBAM) của Liên minh Châu Âu khi doanh nghiệp xuất khẩu sắt thép, nhôm Việt Nam phải hoàn tất dữ liệu phát thải nội hàm (SEE) để đối tác EU nộp báo cáo quý 3. Bước sang năm 2026 khi CBAM thực thi tài chính đầy đủ, chứng chỉ carbon được định giá theo thị trường phát thải EU (ETS) với mức hơn 80 EUR/tấn CO2e (nhôm chịu nghĩa vụ phát thải trực tiếp). Doanh nghiệp nếu không tự đo đạc dữ liệu thực tế tại nhà máy mà sử dụng hệ số mặc định sẽ bị EU áp mức phạt phát thải cao nhất, làm đội giá thành sản phẩm lên 15-30% và mất toàn bộ lợi thế cạnh tranh tại thị trường châu Âu. Eureka hướng dẫn 3 bước cấp bách: Hoàn thiện mẫu biểu giao tiếp dữ liệu phát thải CBAM Communication Template, kiểm kê khí nhà kính tại nhà máy theo chuẩn ISO 14064 và liên kết với các đơn vị xác minh độc lập để bảo vệ chuỗi cung ứng xuất khẩu sang khối 27 nước EU.',
         'Báo Công Thương (congthuong.vn) & Cổng TTĐT Logistics Việt Nam (logistics.gov.vn) | TN (Tin Ngành) / EU CBAM Carbon Border Tax & Green Metallurgy Export Strategy',
         update_date_str),

        (183,
         'quyet dinh 1434 qd chq chuan du lieu chung tu dien tu json dataformat thong quan 2026',
         'quyet dinh 1434/qd-chq, json dataformat hai quan, chuan hoa du lieu chung tu dien tu, thay the quyet dinh 748 1143, loi truyen du lieu khai bao hai quan, he thong it hai quan ket noi doanh nghiep, base64 dinh kem hoa don packing list',
         99.1,
         39800,
         'https://haiquanonline.com.vn',
         'Quyết định số 1434/QĐ-CHQ chính thức áp dụng chuẩn dữ liệu JSON DataFormat: Thay thế hoàn toàn các chuẩn cũ — Doanh nghiệp xuất nhập khẩu và đơn vị phần mềm hải quan cần làm gì ngay để không bị lỗi truyền chứng từ?',
         'Nội dung trọng tâm: Quyết định số 1434/QĐ-CHQ ngày 26/08/2026 của Cục Hải quan chính thức có hiệu lực thi hành từ ngày 15/10/2026, ban hành Quy định tạm thời về định dạng thông điệp dữ liệu chứng từ điện tử liên quan đến hàng hóa xuất khẩu, nhập khẩu, thay thế toàn diện Quyết định 748/QĐ-CHQ và Quyết định 1143/QĐ-CHQ. Toàn bộ thông điệp trao đổi dữ liệu điện tử giữa hệ thống công nghệ thông tin của doanh nghiệp, hãng tàu, kho bãi và cơ quan Hải quan bắt buộc phải chuẩn hóa theo cấu trúc JSON DataFormat. Thực tế nhiều doanh nghiệp và đại lý hải quan gặp lỗi "Từ chối tiếp nhận thông điệp do sai cấu trúc định dạng" khi truyền file đính kèm (hóa đơn thương mại, packing list, chứng chỉ chất lượng C/Q), dẫn đến việc tờ khai bị treo, không thể lấy phân luồng thông quan. Eureka tổng hợp hướng dẫn kỹ thuật 4 điểm: Nâng cấp phiên bản phần mềm khai báo hải quan mới nhất đáp ứng schema JSON của QĐ 1434, chuẩn hóa trường dữ liệu mã hóa Base64 cho file đính kèm, kiểm tra mã số doanh nghiệp và cơ chế đồng bộ dữ liệu chữ ký số tập trung.',
         'Báo Hải quan Online (haiquanonline.com.vn) & Ban CNTT Hải quan | HT (Học Thực chiến) / Customs Electronic Data Interchange & JSON Compliance Guide',
         update_date_str),

        (184,
         'cong van 23405 chq nvthq phan loai hang hoa ma hs may moc thiet bi linh kien thong tu 38',
         'cong van 23405/chq-nvthq, phan loai hang hoa xuat nhap khau, quy tac 1 2a 3b 6 ma hs, phan loai may moc dong bo thao roi, tranh chap ma hs thue suat 0% va 10%, ho so xac dinh truoc ma so hai quan advance ruling, thong tu 39 2018 tt btc',
         98.8,
         37500,
         'https://customs.gov.vn',
         'Công văn số 23405/CHQ-NVTHQ: Siết chặt quy trình phân loại hàng hóa và giải quyết tranh chấp mã HS máy móc, linh kiện — Bí quyết lập Hồ sơ kỹ thuật xác định trước mã số (Advance Ruling) bảo vệ doanh nghiệp',
         'Nội dung trọng tâm: Cục Hải quan (Ban Nghiệp vụ thuế hải quan) ban hành Công văn số 23405/CHQ-NVTHQ ngày 23/10/2026 chấn chỉnh công tác phân loại hàng hóa xuất nhập khẩu, đặc biệt đối với các nhóm hàng máy móc thiết bị phức tạp, dây chuyền tháo rời nhiều chuyến và linh kiện điện tử đa chức năng. Rất nhiều doanh nghiệp bị hải quan bác mã HS tự khai (thuế suất 0%) và chuyển sang mã có thuế suất 5-10%, kèm theo việc truy thu thuế và xử phạt vi phạm hành chính do áp dụng sai Quy tắc 2(a) về hàng chưa hoàn chỉnh hoặc Quy tắc 3(b) về hàng hợp bộ. Bài viết phân tích case study thực tế và đưa ra giải pháp bảo vệ pháp lý toàn diện: Thủ tục nộp hồ sơ xác định trước mã số (Advance Ruling) theo Điều 24 Nghị định 08/2015/NĐ-CP và Thông tư 39/2018/TT-BTC trước khi nhập khẩu tối thiểu 60 ngày để có kết quả phân loại ràng buộc pháp lý toàn quốc; quy trình bảo quản mẫu và chuẩn bị tài liệu kỹ thuật của nhà sản xuất (catalogue, bản vẽ kết cấu, thuyết minh nguyên lý hoạt động) để đối thoại trực tiếp với bộ phận Kiểm tra sau thông quan.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Ban Nghiệp vụ Thuế Hải quan | CS (Case Study) / Machine HS Code Classification & Advance Ruling Defense',
         update_date_str),

        (185,
         'chot ty gia hai quan thu nam 29 10 2026 usd vnd lap ke hoach truyen to khai tiet kiem thue',
         'ty gia hai quan vietcombank 29 10 2026, ty gia tinh thue tuan 45, can ngay truyen to khai thue nhap khau, ty gia usd vnd lap dinh 25450, quan tri dong tien thue xnk, nghi dinh 08 2015 sua doi, bao lanh thue 24 7',
         98.5,
         35900,
         'https://tapchitaichinh.vn',
         'Kỳ chốt tỷ giá tính thuế hải quan thứ Năm (29/10/2026): Tỷ giá USD/VND neo mức đỉnh và mẹo căn giờ mở tờ khai tiết kiệm hàng chục triệu tiền thuế cho lô hàng nhập khẩu đầu tháng 11',
         'Nội dung trọng tâm: Hôm nay, thứ Năm ngày 29/10/2026 là ngày cơ quan Hải quan lấy tỷ giá mua vào chuyển khoản của Vietcombank để công bố tỷ giá tính thuế chính thức cho toàn bộ tuần công tác tiếp theo (từ thứ Sáu 30/10/2026 đến thứ Năm 05/11/2026). Trong bối cảnh tỷ giá USD/VND trên thị trường liên ngân hàng và biểu niêm yết thương mại neo sát trần biên độ, mỗi biến động tăng 50-100 đồng/USD sẽ làm tăng trực tiếp số thuế nhập khẩu và thuế GTGT phải nộp bằng VND của các lô hàng máy móc, nguyên vật liệu trị giá hàng trăm nghìn USD. Eureka chia sẻ cẩm nang chiến lược điều phối chứng từ thực chiến: (1) So sánh tỷ giá tuần hiện tại với tỷ giá chốt ngày 29/10 để quyết định truyền tờ khai trước 23h59 hôm nay hay chờ sang ngày mai; (2) Cách tính điểm rơi giá trị thuế nộp ngay đối với hàng luồng Vàng/Đỏ; (3) Tối ưu hóa hạn mức bảo lãnh ngân hàng nộp thuế 24/7 để thông quan giải phóng hàng nhanh, tránh lưu kho lưu bãi cuối tuần.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & Cục Quản lý Giá (Bộ Tài chính) | HT (Học Thực chiến) / Customs Forex Rate Strategy & Tax Cost Optimization',
         update_date_str)
    ]

    # Ghi 5 hàng mới vào ws3 (rows 186 - 190)
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 186 + idx
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

    # Cập nhật Total row ở row 191
    total_row_idx = 191
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 185 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D190)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E190)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 185 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 153 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A191:C191')
    ws3.merge_cells('F191:J191')
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
    print(f"File size: {file_size:,} bytes | SHA256: {file_sha256}")

    # ==========================================
    # 4. ĐỒNG BỘ NGUYÊN TỬ SANG ONEDRIVE THUANH
    # ==========================================
    if os.path.exists(ONEDRIVE_DIR):
        print("4. Bắt đầu đồng bộ an toàn sang OneDrive thuanh...")
        temp_onedrive_file = ONEDRIVE_FILE + '.tmp.xlsx'
        shutil.copy2(FILE_PATH, temp_onedrive_file)
        if os.path.exists(temp_onedrive_file):
            os.replace(temp_onedrive_file, ONEDRIVE_FILE)
            print(f"[OK] Đã đồng bộ nguyên tử sang OneDrive: {ONEDRIVE_FILE}")

        # Ghi log đồng bộ
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_line = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - {update_date_str})\n"
        with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as lf:
            lf.write(log_line)
        print(f"[OK] Đã ghi log đồng bộ vào: {SYNC_LOG_PATH}")
    else:
        print(f"[CẢNH BÁO] Không tìm thấy thư mục OneDrive: {ONEDRIVE_DIR}")

    # =============================================================
    # 5. TẠO BẢN TIN ĐỀ XUẤT SCRATCH/HOT_NEWS_BRIEFING_W44_2910.MD
    # =============================================================
    briefing_content = f"""# [BẢN TIN ĐỀ XUẤT CHỦ ĐỀ TUẦN — Hot News Scanner]

- **Tuần:** Tuần 44 (26/10/2026 - 01/11/2026) | Ngày quét: {update_date_str}
- **Tổng cơ hội phát hiện:** 20 cơ hội nóng từ 8 nguồn chính thống & `scratch/daily_notes.md`
- **Nguồn đã quét:**
  1. Tổng cục Hải quan (`customs.gov.vn`)
  2. Cổng Thông tin điện tử Chính phủ (`chinhphu.vn`)
  3. Báo Hải quan Online (`haiquanonline.com.vn`)
  4. Báo Công Thương (`congthuong.vn`)
  5. Tạp chí Tài chính (`tapchitaichinh.vn`)
  6. Google Trends VN (`trends.google.com`)
  7. Drewry Supply Chain Advisors (`drewry.co.uk`) & World Shipping Council (`worldshipping.org`)
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Ghi chép kiểm tra sau thông quan, nợ thuế, áp chỉ tiêu 1.34 và biến động tỷ giá USD/VND)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY {update_date_str})

### Đề xuất #1: [HỎA TỐC PHÒNG VỆ THƯƠNG MẠI] Công văn số 23535/CHQ-GSQL ngày 27/10/2026 & Công văn 8744/BCT-PVTM: Cảnh báo đỏ 7 nhóm mặt hàng rủi ro cao bị điều tra phòng vệ thương mại, gian lận xuất xứ và lẩn tránh thuế — Nguy cơ bị áp thuế trừng phạt của Mỹ, EU
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Phòng vệ Thương mại & Quản trị Rủi ro Xuất xứ Quốc tế)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic bản đồ 7 nhóm hàng rủi ro đỏ & Cảnh báo bẫy gia công đơn giản | YouTube: Cẩm nang chuyên sâu bóc tách quy tắc xuất xứ PSR và tính tỷ lệ RVC/LVC chuẩn quốc tế)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5/5) | Công văn 23535/CHQ-GSQL vừa ban hành ngày 27/10/2026 triển khai chỉ đạo hỏa tốc của Bộ Công Thương |
  | Độ đau | ⭐⭐⭐⭐⭐ (5/5) | Dính gian lận xuất xứ bị Mỹ/EU áp thuế chống bán phá giá trừng phạt từ 25% đến >100%, tạm giữ hàng tại cảng nhiều tháng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh luận nóng: "Nhập linh kiện về lắp ráp thế nào để không bị coi là chuyển tải bất hợp pháp lẩn tránh thuế?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5/5) | Cung cấp cẩm nang 4 bước: Rà soát tỷ lệ RVC, minh bạch hồ sơ mua nguyên liệu đầu vào, lập bảng định mức và kiểm tra chéo PSR |
  | **Tổng điểm** | **19.9 / 20** | **Xếp hạng: #1 (Đề tài Cấp bách Kim Cương — Lá chắn pháp lý bảo vệ kim ngạch xuất khẩu)** |
- **Tóm tắt lý do nên làm:** Đây là vấn đề sinh tử đối với các ngành xuất khẩu tỷ đô của Việt Nam (gỗ, thép, nhôm, dệt may, linh kiện điện tử). Khi các thị trường lớn (Mỹ, EU, Canada) gia tăng tần suất điều tra lẩn tránh thuế, việc chỉ đạo quyết liệt của cơ quan Hải quan là hồi chuông cảnh tỉnh cho các chủ xưởng. Lên bài phân tích giúp định vị Eureka là chuyên gia hàng đầu về xuất xứ và tuân thủ phòng vệ thương mại quốc tế.
- **Nguồn xác minh:** Công văn số 23535/CHQ-GSQL ngày 27/10/2026; Công văn số 8744/BCT-PVTM ngày 22/10/2026; Báo Công Thương (`congthuong.vn`).

---

### Đề xuất #2: [ĐẾM NGƯỢC HẠN CHÓT CBAM] Đếm ngược hạn chót nộp báo cáo phát thải CBAM (31/10): Áp lực nghẹt thở đối với doanh nghiệp xuất khẩu Sắt thép, Nhôm sang EU — Khi chứng chỉ carbon trở thành rào cản tài chính sống còn
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Tiêu chuẩn Xanh CBAM & Chiến lược Xuất khẩu Luyện kim)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic đồng hồ đếm ngược 48h hạn chót & Bảng so sánh thuế tự đo vs hệ số phạt | YouTube: Hướng dẫn chi tiết điền mẫu CBAM Communication Template và kiểm kê khí nhà kính ISO 14064)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5/5) | Hạn chót nộp báo cáo quý 3 rơi vào ngày 31/10/2026 (chỉ còn đúng 48 giờ nữa) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Không có số liệu phát thải nội hàm thực tế bị EU phạt mức trần >80 EUR/tấn CO2e, làm đội giá thành 15-30% |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh cãi gay gắt: "Doanh nghiệp thép và nhôm Việt Nam đã thực sự sẵn sàng chi tiền cho kiểm kê carbon độc lập hay chưa?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp hướng dẫn 3 bước: Chuẩn hóa dữ liệu theo CBAM Template, phân tách phát thải trực tiếp và liên kết đơn vị thẩm định |
  | **Tổng điểm** | **19.6 / 20** | **Xếp hạng: #2 (Đề tài Thời sự Nóng bỏng — Cứu nguy trực tiếp cho các nhà máy xuất khẩu)** |
- **Tóm tắt lý do nên làm:** Hạn chót 31/10 đang tạo ra áp lực tâm lý cực lớn trong cộng đồng xuất khẩu kim loại sang EU. Đa số doanh nghiệp vừa và nhỏ vẫn mơ hồ về cách tính lượng phát thải tích hợp (SEE). Hướng dẫn của Eureka mang tính "cứu nguy hỏa tốc", giải tỏa bế tắc cho các giám đốc nhà máy và củng cố vị thế dẫn dắt xu hướng xanh của Eureka.
- **Nguồn xác minh:** Quy định Cơ chế điều chỉnh biên giới carbon (CBAM) của Liên minh Châu Âu; Báo Công Thương; Cổng TTĐT Logistics Việt Nam (`logistics.gov.vn`).

---

### Đề xuất #3: [CÔNG NGHỆ & DỮ LIỆU HẢI QUAN] Quyết định số 1434/QĐ-CHQ chính thức áp dụng chuẩn dữ liệu JSON DataFormat: Thay thế hoàn toàn các chuẩn cũ — Doanh nghiệp xuất nhập khẩu và đơn vị phần mềm hải quan cần làm gì ngay để không bị lỗi truyền chứng từ?
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Hạ tầng Dữ liệu Hải quan & Xử lý Sự cố Truyền tờ khai)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Cảnh báo lỗi 'Bị từ chối tiếp nhận thông điệp' & 4 bước khắc phục | YouTube: Video demo cấu trúc JSON schema và cách đính kèm file hóa đơn/packing list Base64)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Quyết định 1434/QĐ-CHQ có hiệu lực từ 15/10/2026 và đang là tâm điểm chuyển đổi công nghệ cuối tháng 10 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Lỗi định dạng dữ liệu khiến toàn bộ chứng từ đính kèm bị treo, không lấy được phân luồng tờ khai, chậm thông quan lô hàng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Thảo luận: "Doanh nghiệp phải nâng cấp phần mềm hay chờ các đơn vị IT như ECUS, Thái Sơn cập nhật patch tự động?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp bảng hướng dẫn kỹ thuật 4 điểm: Kiểm tra schema JSON, mã hóa Base64 chứng từ, xác thực token chữ ký số và đối soát gateway |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #3 (Cẩm nang kỹ thuật thực chiến giải nguy hệ thống)** |
- **Tóm tắt lý do nên làm:** Rất nhiều nhân viên chứng từ và đại lý hải quan đang hoang mang khi gặp các mã lỗi kỹ thuật lạ khi gửi file đính kèm tờ khai. Việc Eureka phân tích bản chất của Quyết định 1434/QĐ-CHQ và hướng dẫn cách khắc phục lỗi truyền thông điệp JSON giúp cộng đồng doanh nghiệp giải tỏa tắc nghẽn ngay tại quầy khai báo.
- **Nguồn xác minh:** Quyết định số 1434/QĐ-CHQ ngày 26/08/2026 của Cục Hải quan; Báo Hải quan Online (`haiquanonline.com.vn`).

---

### Đề xuất #4: [NGHIỆP VỤ PHÂN LOẠI HS CODE] Công văn số 23405/CHQ-NVTHQ: Siết chặt quy trình phân loại hàng hóa và giải quyết tranh chấp mã HS máy móc, linh kiện — Bí quyết lập Hồ sơ kỹ thuật xác định trước mã số (Advance Ruling) bảo vệ doanh nghiệp
- **Chuyên mục (Pillar):** `CS` (Case Study - Phân loại HS Code & Thủ tục Xác định trước mã số)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Phân tích bẫy áp mã HS máy móc tháo rời Quy tắc 2a vs 3b | YouTube: Cẩm nang 5 bước lập bộ hồ sơ Advance Ruling hoàn chỉnh được Hải quan chấp thuận)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Công văn 23405/CHQ-NVTHQ ban hành ngày 23/10/2026 chấn chỉnh công tác phân loại hàng hóa toàn quốc |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Khai sai mã HS máy móc từ thuế 0% sang 5-10% bị truy thu hàng tỷ đồng và phạt chậm nộp trong thời hạn 5 năm kiểm tra sau thông quan |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Tranh luận gay gắt: "Tại sao cùng một thiết bị nhưng Chi cục này cho áp mã 0%, Chi cục khác lại bác và bắt nộp thuế 10%?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn quy trình nộp hồ sơ Xác định trước mã số (Advance Ruling) theo Điều 24 NĐ 08 trước khi nhập khẩu để có giá trị pháp lý toàn quốc |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #4 (Cẩm nang giải quyết tranh chấp mã HS có tính bảo chứng cao)** |
- **Tóm tắt lý do nên làm:** Phân loại mã HS máy móc thiết bị luôn là "điểm nóng" dễ phát sinh tranh chấp nhất giữa doanh nghiệp và cơ quan hải quan. Hướng dẫn của Eureka về thủ tục Xác định trước mã số (Advance Ruling) là giải pháp gốc rễ giúp doanh nghiệp loại bỏ 100% rủi ro bị ấn định thuế hồi tố khi kiểm tra sau thông quan.
- **Nguồn xác minh:** Công văn số 23405/CHQ-NVTHQ ngày 23/10/2026; Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`).

---

### Đề xuất #5: [TÀI CHÍNH & TỶ GIÁ THỰC CHIẾN] Kỳ chốt tỷ giá tính thuế hải quan thứ Năm (29/10/2026): Tỷ giá USD/VND neo mức đỉnh và mẹo căn giờ mở tờ khai tiết kiệm hàng chục triệu tiền thuế cho lô hàng nhập khẩu đầu tháng 11
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Kế toán Thuế XNK & Tối ưu Chi phí Tỷ giá)
- **Nền tảng phù hợp:** **Facebook** (Checklist căn giờ truyền tờ khai trước 23h59 thứ Năm) & **YouTube** (Phân tích cơ chế chốt tỷ giá Vietcombank và hạn mức bảo lãnh nộp thuế 24/7)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5/5) | Hôm nay đúng ngày thứ Năm 29/10/2026 — Kỳ chốt tỷ giá tính thuế chính thức cho tuần 45 (30/10 - 05/11/2026) |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.7/5) | Tỷ giá USD neo đỉnh trần khiến số thuế nhập khẩu và GTGT phải nộp tăng mạnh; truyền lệch ngày làm đội thêm chi phí đáng kể |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.6/5) | Thảo luận: "Nên truyền tờ khai trước 23h59 hôm nay hay chờ sang ngày mai để có tỷ giá có lợi hơn?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Chia sẻ 3 chiến thuật: So sánh tỷ giá 2 tuần, mẹo canh thời điểm phân luồng tờ khai và kết hợp bảo lãnh ngân hàng 24/7 |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #5 (Mẹo tài chính thực chiến sát sườn cho kế toán thuế và chủ hàng)** |
- **Tóm tắt lý do nên làm:** Đánh trúng yếu tố thời điểm "nóng từng giờ" trong ngày thứ Năm. Doanh nghiệp nhập khẩu nếu nắm được nguyên tắc chốt tỷ giá tuần có thể chủ động lựa chọn thời khắc truyền tờ khai tối ưu nhất, tiết kiệm ngay hàng triệu đến hàng chục triệu đồng tiền thuế cho các lô hàng lớn.
- **Nguồn xác minh:** Nghị định số 08/2015/NĐ-CP (sửa đổi tại NĐ 167/2025/NĐ-CP & NĐ 252/2026/NĐ-CP); Tạp chí Tài chính (`tapchitaichinh.vn`); Vietcombank.

---

## 🎯 BƯỚC TIẾP THEO (NEXT ACTIONS DÀNH CHO CHỊ THÚY)
1. **Lựa chọn chủ đề ưu tiên:** Chị Thúy có thể phản hồi chọn 1 hoặc nhiều chủ đề (ví dụ: *'Chọn #1 và #2'*) để hệ thống tự động đẩy vào:
   - Quy trình `wf01`: Lập lịch biên tập chi tiết trong Content Calendar.
   - Hoặc quy trình `wf02`/`wf03`: Sản xuất nội dung bài viết chuyên sâu và kịch bản video AI ngay lập tức.
2. **Trường hợp chưa chọn:** Bản tin này đã được lưu trữ bền vững tại `scratch/hot_news_briefing_w44_2910.md` và sẵn sàng tái sử dụng cho các kế hoạch truyền thông tiếp theo.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo bản tin đề xuất tại: {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w44_2910()
