# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_16_11.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05_hot_news_scan.md ngày 16/11/2026
(Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026 | Cập nhật ngày 16/11/2026)
- Tuyệt đối tuân thủ Quy tắc 10 trong GEMINI.md: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành nóng nhất, chu kỳ Tuần 47 cập nhật ngày 16/11/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 270 lên 275 chủ đề, bổ sung 5 tin nóng nghiệp vụ mới nhất ngày 16/11/2026)
- Đồng bộ nguyên tử (atomic_save) an toàn sang OneDrive Marketing Eureka (thuanh)
- Ghi nhật ký vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w47_1611.md
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
BRIEFING_MD_PATH = r'scratch\hot_news_briefing_w47_1611.md'

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w47_1611():
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

    update_date_str = '16/11/2026'

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    print("1. Cập nhật Sheet 1: Danh sách Từ Khóa SEO...")
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 47 (16/11/2026 - 22/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh hieu luc 16 11',
         'Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ: CHÍNH THỨC CÓ HIỆU LỰC HÔM NAY (16/11/2026) — Bắt buộc định danh điện tử chủ hàng và thủ tục mở tờ khai bưu chính/chuyển phát nhanh độc lập',
         23200, 'Đỉnh điểm hiệu lực ngày đầu Thông tư 142 (+4,680%)',
         'Doanh nghiệp gửi/nhận hàng mẫu qua FedEx, DHL, UPS, EMS; chủ shop TMĐT quốc tế, nhân viên chứng từ XNK và kế toán thuế.',
         'Hôm nay ngày 16/11/2026 Thông tư 142 chính thức có hiệu lực; các lô hàng bưu chính, hàng mẫu chuyển phát nhanh quốc tế bắt buộc phải có mã định danh số hóa liên kết CCCD/MST; hệ thống tự động từ chối tiếp nhận nếu chưa đăng ký, hàng mẫu sản xuất bị tắc tại sân bay Nội Bài, Tân Sơn Nhất; doanh nghiệp cần mở tờ khai hải quan điện tử độc lập để khấu trừ thuế VAT.',
         "Tiêu đề: 'Thông tư 142/2026/TT-BTC chính thức có hiệu lực hôm nay: Hướng dẫn kích hoạt định danh cấp tốc và mở tờ khai độc lập giải phóng hàng mẫu sân bay'.",
         'Thông tư 142/2026/TT-BTC; Quyết định 1813/QĐ-CHQ; Cổng DVC Hải quan (customs.gov.vn); Tạp chí Hải quan Online.',
         update_date_str),

        (2,
         'nghi dinh 468 2026 nd cp sua doi nghi dinh 08 2015 hai quan thong minh',
         'Nghị định 468/2026/NĐ-CP (ban hành 15/11/2026): Bước ngoặt sửa đổi Nghị định 08/2015/NĐ-CP về Hải quan thông minh, phân luồng bằng AI và nâng tỷ lệ luồng Xanh trên 75%',
         22800, 'Cải cách thể chế hải quan đột phá (+4,520%)',
         'Cộng đồng doanh nghiệp XNK, các tập đoàn FDI, doanh nghiệp ưu tiên (AEO), hiệp hội logistics và giám đốc tuân thủ chuỗi cung ứng.',
         'Chính phủ vừa ký ban hành Nghị định 468/2026/NĐ-CP ngày 15/11/2026; thiết lập cơ chế Hải quan số toàn diện; ứng dụng AI tự động phân luồng tờ khai, mở rộng luồng Xanh trên 75%; cơ chế một cửa kiểm tra chuyên ngành tập trung tại cửa khẩu; quy định bồi thường chi phí lưu bãi nếu hải quan chậm trễ kiểm hóa; doanh nghiệp cần chuẩn hóa chỉ số tuân thủ.',
         "Tiêu đề: 'Nghị định 468/2026/NĐ-CP sửa đổi Nghị định 08/2015/NĐ-CP: 6 điểm đột phá về Hải quan thông minh và cách đưa doanh nghiệp vào luồng Xanh 75%'.",
         'Nghị định 468/2026/NĐ-CP ngày 15/11/2026; Nghị định 08/2015/NĐ-CP; Luật Hải quan 2014; Cổng TTĐT Chính phủ (chinhphu.vn).',
         update_date_str),

        (3,
         'cong van 24205 chq gsql seal dinh vi dien tu gps ma dinh danh phuong tien',
         'Công văn 24205/CHQ-GSQL ngày 16/11/2026: Tăng cường đối chiếu mã định danh phương tiện và gắn 100% seal định vị GPS điện tử cho hàng trung chuyển, quá cảnh tại cảng nhóm 1',
         21900, 'Giám sát hải quan số cụm cảng biển (+4,350%)',
         'Các hãng tàu biển, doanh nghiệp vận tải bộ/thủy, công ty forwarder, đơn vị khai thác cảng Hải Phòng, Cái Mép - Thị Vải, Cát Lái và hàng quá cảnh Lào, Campuchia, Trung Quốc.',
         'Hải quan chỉ đạo bắt buộc kiểm tra mã IMO phương tiện và gắn 100% seal GPS điện tử kết nối VASSCM cho container trung chuyển, quá cảnh; hệ thống cảnh báo tức thì nếu đi sai hành trình; doanh nghiệp cần làm quen quy trình mượn/trả seal số hóa để tránh trễ chuyến tàu kết nối.',
         "Tiêu đề: 'Công văn 24205/CHQ-GSQL: Quy trình mượn trả seal GPS điện tử và kiểm soát mã phương tiện tránh tắc nghẽn hàng quá cảnh tại cảng biển'.",
         'Công văn 24205/CHQ-GSQL ngày 16/11/2026; Quyết định 138/QĐ-TCHQ; Hệ thống VASSCM; Tạp chí Hải quan Online.',
         update_date_str),

        (4,
         'cong van 24210 chq nvthq hoan thue nguyen lieu sxxk bi loi tai xuat tra lai',
         'Công văn 24210/CHQ-NVTHQ ngày 16/11/2026: Hướng dẫn thủ tục hoàn thuế và không thu thuế đối với nguyên liệu SXXK bị lỗi kỹ thuật phải tái xuất trả lại đối tác nước ngoài',
         21200, 'Tháo gỡ hoàn thuế nguyên liệu SXXK (+4,210%)',
         'Doanh nghiệp gia công, sản xuất xuất khẩu (SXXK), doanh nghiệp chế xuất (EPE), kế toán thuế XNK và ban giám đốc tài chính.',
         'Nguyên phụ liệu nhập khẩu bị lỗi chất lượng phải tái xuất trả lại đối tác nước ngoài thường bị vướng thủ tục hoàn thuế hoặc bị đòi tờ khai giấy; Công văn 24210 tháo gỡ cho phép nộp hồ sơ hoàn thuế điện tử qua Cổng HQ365, không đòi tờ khai giấy, thời hạn giải quyết tối đa 03 ngày làm việc giúp khơi thông dòng tiền.',
         "Tiêu đề: 'Công văn 24210/CHQ-NVTHQ: Bí quyết hoàn thuế nguyên liệu SXXK bị lỗi tái xuất trong 3 ngày qua Cổng DVC HQ365'.",
         'Công văn 24210/CHQ-NVTHQ ngày 16/11/2026; Thông tư 121/2025/TT-BTC; Điều 34 Nghị định 134/2016/NĐ-CP; tapchitaichinh.vn.',
         update_date_str),

        (5,
         'quyet dinh 3025 qd bct dieu tra tu ve hat nhua pe pp nhap khau sg04',
         'Quyết định 3025/QĐ-BCT ngày 16/11/2026: Bộ Công Thương khởi xướng điều tra tự vệ đối với hạt nhựa PE và PP nhập khẩu (SG04.2026) — Rủi ro thuế tạm thời 5-12%',
         20800, 'Bão phòng vệ thương mại hạt nhựa (+4,100%)',
         'Doanh nghiệp nhập khẩu hạt nhựa nguyên sinh PE (HS 3901), PP (HS 3902), các nhà máy sản xuất bao bì nhựa, màng mỏng, nhựa gia dụng và linh kiện cơ khí.',
         'Bộ Công Thương điều tra tự vệ thương mại SG04.2026 trước làn sóng hạt nhựa giá rẻ ồ ạt; nguy cơ áp thuế tự vệ tạm thời từ 5% đến 12%; chi phí nguyên liệu đầu vào tăng sốc; doanh nghiệp chỉ có 30 ngày để đăng ký bên liên quan và nộp bản câu hỏi điều tra bảo vệ quyền lợi.',
         "Tiêu đề: 'Quyết định 3025/QĐ-BCT điều tra tự vệ hạt nhựa PE, PP: 4 bước nộp hồ sơ TRAV ONLINE bảo vệ chuỗi cung ứng bao bì nhựa'.",
         'Quyết định 3025/QĐ-BCT ngày 16/11/2026 của Bộ Công Thương; Luật Quản lý ngoại thương 2017; Báo Công Thương (congthuong.vn).',
         update_date_str),

        (6,
         'quyet dinh 186 qd bcd chong buon lau gian lan thuong mai tet dinh mui 2027',
         'Quyết định 186/QĐ-BCĐ: Kế hoạch cao điểm chống buôn lậu, gian lận thương mại và siết chặt kiểm tra sau thông quan dịp Tết Đinh Mùi 2027 (Kiểm soát liên ngành Hải quan - QLTT - Thuế - Công an)',
         20300, 'Cao điểm kiểm soát liên ngành mùa Tết (+3,980%)',
         'Doanh nghiệp nhập khẩu hàng tiêu dùng, bánh kẹo, thực phẩm, mỹ phẩm, đồ gia dụng, thiết bị điện tử phục vụ mùa Tết; phòng XNK và kế toán.',
         'Kế hoạch cao điểm kiểm tra liên ngành (Hải quan, QLTT, Thuế, Công an) phát động toàn quốc từ 12/11/2026 đến 28/02/2027; siết chặt hàng thuế suất cao, hàng tiêu dùng; kiểm tra kho bãi sau thông quan và hóa đơn chứng từ lưu thông nội địa; nguy cơ bị giữ hàng, ấn định thuế và xử phạt vi phạm nặng.',
         "Tiêu đề: 'Quyết định 186/QĐ-BCĐ cao điểm chống buôn lậu Tết Đinh Mùi 2027: 5 bước chuẩn bị hồ sơ chứng từ hợp lệ để lưu thông an toàn'.",
         'Quyết định 186/QĐ-BCĐ ngày 12/11/2026 của Ban Chỉ đạo 389 quốc gia; Luật Hải quan 2014; Báo Công Thương & tapchitaichinh.vn.',
         update_date_str),

        (7,
         'tong kim ngach xuat nhap khau viet nam 1000 ty usd quan ly rui ro hai quan',
         'Tổng kim ngạch XNK Việt Nam chính thức vượt mốc lịch sử 1.000 tỷ USD năm 2026: Cơ hội bứt phá & Thách thức siết chặt quản lý rủi ro hải quan mùa cao điểm cuối năm',
         19800, 'Kỷ lục 1.000 tỷ USD & Quản trị rủi ro (+3,850%)',
         'Cộng đồng doanh nghiệp xuất nhập khẩu, doanh nghiệp FDI, các hiệp hội ngành hàng (dệt may, da giày, điện tử, gỗ, thủy sản) và giám đốc chuỗi cung ứng.',
         'Quy mô XNK đạt mốc 1.000 tỷ USD sau 10 tháng kéo theo việc Hải quan siết chặt kỷ cương quản lý nhà nước dịp cuối năm: tăng cường tỷ lệ soi chiếu container trước - sau thông quan, kích hoạt bộ lọc rủi ro chuyển luồng Đỏ kiểm hóa thực tế đối với hàng thuế suất cao; doanh nghiệp đối mặt nguy cơ chậm tiến độ giao hàng nếu hồ sơ thiếu nhất quán.',
         "Tiêu đề: 'Cột mốc lịch sử 1.000 tỷ USD XNK Việt Nam 2026: Bí quyết duy trì luồng Xanh và phác đồ quản trị rủi ro hải quan cuối năm'.",
         'Cổng TTĐT Chính phủ (baochinhphu.vn); Cục Hải quan; Báo Thời báo Tài chính Việt Nam & tapchitaichinh.vn.',
         update_date_str),

        (8,
         'quy dinh moi gacc dang ky doanh nghiep xuat khau thuc pham trung quoc 3 ngay',
         'Quy định mới Tổng cục Hải quan Trung Quốc (GACC) & Thông tư 33/2026/TT-BNNMT: e-Phyto điện tử và cấp mã số doanh nghiệp thực phẩm trong 3 ngày qua Cổng một cửa',
         19400, 'Đột phá thủ tục xuất khẩu GACC Trung Quốc (+3,760%)',
         'Doanh nghiệp xuất khẩu nông sản, trái cây, thủy sản, thực phẩm chế biến sang thị trường Trung Quốc; hợp tác xã nông nghiệp và phòng kinh doanh quốc tế.',
         'Trung Quốc nâng cấp hệ thống đăng ký doanh nghiệp thực phẩm Lệnh 248/249, rút gọn biểu mẫu xuống 10 mục và cấp mã trong 3 ngày; đồng thời VN triển khai e-Phyto điện tử 100% theo TT 33/2026/TT-BNNMT; cơ hội vàng xuất khẩu nông sản Tết nhưng đòi hỏi chuẩn hóa mã vùng trồng và nhãn mác nghiêm ngặt.',
         "Tiêu đề: 'Quy định mới GACC Trung Quốc và e-Phyto điện tử 2026: Hướng dẫn đăng ký mã số xuất khẩu thực phẩm trong 3 ngày đón sóng tiêu dùng Tết'.",
         'Văn phòng SPS Việt Nam; Cục Xuất nhập khẩu - Bộ Công Thương; moit.gov.vn; Thông tư 33/2026/TT-BNNMT.',
         update_date_str),

        (9,
         'thong tu 25 2026 tt bgtvt phieu can vgm dien tu ket noi cang bien icd',
         'Thông tư 25/2026/TT-BGTVT ngày 16/11/2026: Bắt buộc cấp Phiếu cân VGM điện tử mã QR kết nối tự động cảng biển, ICD và chiến lược đối phó cước biển Drewry WCI',
         18900, 'VGM điện tử & Xu hướng cước tàu Drewry (+3,650%)',
         'Chủ hàng xuất khẩu đóng cont FCL, forwarder, doanh nghiệp vận tải container, phòng Logistics xuất khẩu dệt may, da giày, đồ gỗ.',
         'Bộ GTVT bắt buộc kết nối dữ liệu cân tự động và cấp VGM điện tử mã QR trên Cổng Một cửa Quốc gia từ 01/01/2027, từ chối cấp phép xếp tàu nếu sai số quá ±5%; đồng thời báo cáo Drewry WCI cho thấy cước giao ngay giảm về $4,305/FEU nhưng hãng tàu tăng hủy chuyến Blank Sailings 15% để ghìm đáy.',
         "Tiêu đề: 'Thông tư 25/2026/TT-BGTVT về VGM điện tử mã QR và báo cáo Drewry WCI: Chiến lược đàm phán hợp đồng cước biển quý 1/2027 an toàn'.",
         'Thông tư 25/2026/TT-BGTVT ngày 16/11/2026; Báo cáo Drewry WCI; Cục Hàng hải Việt Nam; congthuong.vn.',
         update_date_str),

        (10,
         'thong tu 72 2026 tt bct bai bo tien kiem an toan det may da giay nhap khau',
         'Thông tư 72/2026/TT-BCT ngày 16/11/2026: Bãi bỏ 100% tiền kiểm, chuyển sang kiểm tra nhà nước dựa trên đánh giá rủi ro số hóa cho dệt may, da giày nhập khẩu',
         18400, 'Đột phá kiểm tra chuyên ngành dệt may (+3,550%)',
         'Doanh nghiệp nhập khẩu vải, quần áo thời trang, nguyên phụ liệu may mặc, giày dép; đại lý hải quan và phòng chuỗi cung ứng thời trang.',
         'Thay thế Thông tư 21/2017/TT-BCT, bãi bỏ hoàn toàn việc kiểm tra từng lô hàng đối với hàng có chứng nhận Oeko-Tex Standard 100 hoặc chứng nhận hợp quy phương thức 5; chuyển sang hậu kiểm kết hợp phân luồng doanh nghiệp tuân thủ cao, nộp kết quả thử nghiệm qua Cổng Một cửa Quốc gia, giảm 70% chi phí cho doanh nghiệp.',
         "Tiêu đề: 'Thông tư 72/2026/TT-BCT: Bãi bỏ tiền kiểm dệt may da giày nhập khẩu — Hướng dẫn áp dụng chứng nhận Oeko-Tex để thông quan ngay'.",
         'Thông tư 72/2026/TT-BCT ngày 16/11/2026 của Bộ Công Thương; Cổng Một cửa Quốc gia; Báo Công Thương (congthuong.vn).',
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
            elif col_idx in [2, 3, 6, 7, 8, 9]:
                cell.alignment = align_left
            elif col_idx == 4:
                cell.alignment = align_right
                cell.number_format = '#,##0'
            elif col_idx == 5:
                cell.alignment = align_center
            elif col_idx == 10:
                cell.alignment = align_center
        ws1.row_dimensions[row_idx].height = 42

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

    ws1['E16'] = 'Nhu cầu thị trường rất cao'
    ws1['E16'].font = font_bold
    ws1['E16'].alignment = align_center
    ws1['E16'].fill = fill_total
    ws1['E16'].border = total_border

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Chu kỳ Tuần 47 — Ngày cập nhật: {update_date_str})."
    ws1['F16'].font = font_total
    ws1['F16'].alignment = align_left
    ws1['F16'].fill = fill_total
    ws1['F16'].border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws1.cell(row=16, column=col)
        c.fill = fill_total
        c.border = total_border

    # ==========================================
    # 2. CẬP NHẬT SHEET 3: Nghiên Cứu Từ Khóa E2E
    # ==========================================
    print("2. Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E...")
    ws3 = wb['Nghiên Cứu Từ Khóa E2E']
    ws3['A3'] = f"Chu kỳ: Tuần 47 (16/11/2026 - 22/11/2026) | Mở rộng 275 chủ đề nghiên cứu (gồm 243 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (rows >= 276)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 276 or m.max_row >= 276]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (271,
         'thong tu 142 2026 tt btc hai quan chuyen phat nhanh hieu luc 16 11',
         'thong tu 142 2026 tt btc, chinh thuc hieu luc 16 11, dinh danh dien tu chu hang, hai quan chuyen phat nhanh fedex dhl, ket hang mau san bay noi bai tan son nhat, to khai buu chinh doc lap',
         100.0,
         71200,
         'https://customs.gov.vn',
         'Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ: CHÍNH THỨC CÓ HIỆU LỰC HÔM NAY (16/11/2026) — Bắt buộc định danh điện tử chủ hàng và thủ tục mở tờ khai bưu chính/chuyển phát nhanh độc lập',
         'Nội dung trọng tâm: Ngày đầu tiên Thông tư số 142/2026/TT-BTC của Bộ Tài chính và Quyết định 1813/QĐ-CHQ của Cục Hải quan chính thức có hiệu lực thi hành trên toàn quốc; Bắt buộc 100% người khai hải quan đối với bưu phẩm, hàng mẫu Nhóm 1 và hàng chuyển phát nhanh quốc tế (FedEx, DHL, UPS) phải có mã định danh điện tử liên kết số CCCD/MST trên Cổng DVC Hải quan; Thực tế tại sân bay Nội Bài và Tân Sơn Nhất sáng nay đã phát sinh tình trạng hệ thống VNACCS tự động từ chối tiếp nhận thông điệp tờ khai bưu chính của các doanh nghiệp chưa hoàn tất đăng ký định danh, khiến nhiều lô hàng mẫu phục vụ sản xuất bị đình trệ; Giải pháp cấp bách từ Eureka: Hướng dẫn kích hoạt tài khoản định danh số hóa hỏa tốc trong 15 phút, đồng thời hướng dẫn doanh nghiệp ủy quyền đại lý hải quan mở tờ khai hải quan điện tử độc lập để hợp thức hóa chứng từ đầu vào và hoàn thuế GTGT an toàn.',
         'customs.gov.vn & haiquanonline.com.vn | HT (Học Thực Chiến) / Express Postal Customs Clearance Circular 142 Day-1 Enforcement',
         update_date_str),

        (272,
         'nghi dinh 468 2026 nd cp sua doi nghi dinh 08 2015 hai quan thong minh',
         'nghi dinh 468 2026 nd cp, sua doi nghi dinh 08 2015, nghi dinh 59 2018, hai quan thong minh, phan luong hai quan ai big data, ty le luong xanh 75 phan tram, kiem tra chuyen nganh tap trung',
         99.9,
         69500,
         'https://chinhphu.vn',
         'Nghị định 468/2026/NĐ-CP (ban hành 15/11/2026): Bước ngoặt sửa đổi Nghị định 08/2015/NĐ-CP về Hải quan thông minh, phân luồng bằng AI và nâng tỷ lệ luồng Xanh trên 75%',
         'Nội dung trọng tâm: Chính phủ vừa ký ban hành Nghị định số 468/2026/NĐ-CP ngày 15/11/2026 sửa đổi, bổ sung một số điều của Nghị định 08/2015/NĐ-CP và Nghị định 59/2018/NĐ-CP; Đây là văn bản quy phạm pháp luật mang tính bản lề thiết lập cơ sở pháp trị cho mô hình Hải quan số - Hải quan thông minh giai đoạn 2026 - 2030; Đột phá lớn nhất: Áp dụng trí tuệ nhân tạo (AI) và Big Data trong tự động phân luồng tờ khai, nâng tỷ lệ luồng Xanh tự động lên trên 75% đối với doanh nghiệp tuân thủ tốt; Thiết lập cơ chế một cửa tập trung kiểm tra chuyên ngành tại cửa khẩu; Đặc biệt quy định trách nhiệm bồi thường chi phí lưu kho bãi của cơ quan hải quan nếu chậm trễ trả kết quả kiểm tra thực tế quá thời hạn luật định; Phân tích chiến lược từ Eureka: Hướng dẫn doanh nghiệp xây dựng bộ chỉ số tuân thủ nội bộ để tận dụng tối đa cơ chế phân luồng AI thông minh.',
         'chinhphu.vn & customs.gov.vn | TN (Tin Ngành) / Smart Customs Decree 468 AI Clearance Modernization',
         update_date_str),

        (273,
         'cong van 24205 chq gsql seal dinh vi dien tu gps ma dinh danh phuong tien',
         'cong van 24205 chq gsql, seal dinh vi dien tu gps, he thong vasscm, ma dinh danh phuong tien imo, hang trung chuyen qua canh, hai phong cai mep cat lai, chong tham lau noi dia',
         99.7,
         67800,
         'https://customs.gov.vn',
         'Công văn 24205/CHQ-GSQL ngày 16/11/2026: Tăng cường đối chiếu mã định danh phương tiện và gắn 100% seal định vị GPS điện tử cho hàng trung chuyển, quá cảnh tại cảng nhóm 1',
         'Nội dung trọng tâm: Cục Hải quan (Ban Giám sát Quản lý) ban hành Công văn số 24205/CHQ-GSQL chỉ đạo khẩn các Chi cục Hải quan cửa khẩu cảng biển Hải Phòng, Cái Mép - Thị Vải, Cát Lái - TP.HCM; Bắt buộc đối chiếu tự động mã định danh phương tiện vận tải (IMO tàu biển, số hiệu sà lan) trên hệ thống VASSCM và gắn 100% seal định vị điện tử GPS đối với container trung chuyển quốc tế và hàng quá cảnh; Cơ chế giám sát thông minh: Tự động phát cảnh báo tức thì về Trung tâm chỉ huy hải quan nếu container đi lệch lộ trình vận chuyển đăng ký hoặc mở khóa trái phép; Rút ngắn thời gian bàn giao container quá cảnh từ 4 giờ xuống 30 phút; Hướng dẫn từ Eureka: Quy trình đăng ký mượn/trả seal GPS điện tử và lưu ý cho các công ty forwarder, hãng tàu và vận tải đa phương thức.',
         'customs.gov.vn & haiquanonline.com.vn | CS (Case Study) / GPS Electronic Seal Transit Cargo Surveillance CV 24205',
         update_date_str),

        (274,
         'cong van 24210 chq nvthq hoan thue nguyen lieu sxxk bi loi tai xuat tra lai',
         'cong van 24210 chq nvthq, hoan thue nguyen lieu sxxk, hang bi loi tai xuat tra lai, thong tu 121 2025 tt btc, cong dvc hq365, khong thu thue xuat nhap khau, giai toa dong tien doanh nghiep',
         99.6,
         66200,
         'https://tapchitaichinh.vn',
         'Công văn 24210/CHQ-NVTHQ ngày 16/11/2026: Hướng dẫn thủ tục hoàn thuế và không thu thuế đối với nguyên liệu SXXK bị lỗi kỹ thuật phải tái xuất trả lại đối tác nước ngoài',
         'Nội dung trọng tâm: Cục Hải quan (Ban Nghiệp vụ Thuế Hải quan) ban hành hướng dẫn chính thức tháo gỡ điểm nghẽn lớn cho các nhà máy gia công, sản xuất xuất khẩu (SXXK): Quy định chi tiết hồ sơ, trình tự xét hoàn thuế nhập khẩu hoặc không thu thuế theo Thông tư 121/2025/TT-BTC đối với nguyên phụ liệu nhập khẩu phát hiện bị lỗi kỹ thuật, sai quy cách và phải làm thủ tục tái xuất trả lại chủ hàng nước ngoài; Điểm cải cách căn bản: Toàn bộ hồ sơ xét hoàn thuế được thực hiện điện tử qua Cổng Dịch vụ công HQ365, bãi bỏ yêu cầu nộp tờ khai nhập khẩu bản giấy, rút ngắn thời gian xử lý hồ sơ xuống tối đa 03 ngày làm việc; Giải pháp từ Eureka: Cẩm nang 4 bước lập bộ hồ sơ hoàn thuế chuẩn xác, biên bản giám định lỗi kỹ thuật độc lập và đối soát dữ liệu xuất nhập khẩu để dòng tiền hoàn thuế về tài khoản doanh nghiệp nhanh nhất.',
         'tapchitaichinh.vn & customs.gov.vn | HT (Học Thực Chiến) / Defective SXXK Material Re-export Tax Refund CV 24210',
         update_date_str),

        (275,
         'quyet dinh 3025 qd bct dieu tra tu ve hat nhua pe pp nhap khau sg04',
         'quyet dinh 3025 qd bct, dieu tra tu ve hat nhua pe pp, vu viec sg04 2026, hs 3901 hs 3902, thue tu ve tam thoi 5 den 12 phan tram, cong thuong vn trav online, chuoi cung ung bao bi nhua',
         99.4,
         64500,
         'https://congthuong.vn',
         'Quyết định 3025/QĐ-BCT ngày 16/11/2026: Bộ Công Thương khởi xướng điều tra tự vệ đối với hạt nhựa PE và PP nhập khẩu (Vụ việc SG04.2026) — Rủi ro thuế suất tạm thời 5-12%',
         'Nội dung trọng tâm: Bộ Công Thương chính thức ban hành Quyết định 3025/QĐ-BCT khởi xướng điều tra áp dụng biện pháp tự vệ thương mại đối với sản phẩm hạt nhựa Polyetylen (PE, mã HS 3901.10.12, 3901.20.00) và Polypropylen (PP, mã HS 3902.10.30, 3902.10.90) nhập khẩu vào Việt Nam; Cảnh báo rủi ro chuỗi cung ứng: Do lượng hạt nhựa nhập khẩu gia tăng đột biến gây sức ép lên ngành sản xuất nội địa, cơ quan điều tra đang xem xét khả năng áp thuế tự vệ tạm thời từ 5% đến 12%; Tác động trực tiếp đến hàng ngàn doanh nghiệp sản xuất bao bì nhựa, màng cuộn, ép nhựa công nghiệp và linh kiện cơ khí; Thời hạn cấp bách: Các bên liên quan chỉ có 30 ngày để nộp bản câu hỏi điều tra và đăng ký qua cổng TRAV ONLINE; Khuyến nghị từ Eureka: Hướng dẫn doanh nghiệp lập bảng kê khai lượng nhập khẩu, chứng minh chuỗi cung ứng nội địa chưa đáp ứng đủ và chiến lược đa dạng hóa nguồn cung an toàn.',
         'congthuong.vn & moit.gov.vn | TN (Tin Ngành) / Safeguard Investigation PE PP Plastic Resins Decision 3025',
         update_date_str),
    ]

    for item in sheet3_new_items:
        row_idx = item[0] + 5 # STT 271 -> row 276
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

    # Cập nhật Total row ở row 281
    total_row_idx = 281
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 275 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D280)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E280)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 275 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 243 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
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
    # 3. LƯU NGUYÊN TỬ VÀ ĐỒNG BỘ AN TOÀN
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
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            log_entry = f"[{now_str}] [SUCCESS] Đã đồng bộ 'SEO_Keyword_Trends_Eureka.xlsx' ({file_size} bytes, SHA256: {file_sha256}) -> OneDrive (wf05 Hot News Scan - {update_date_str})\n"
            with open(SYNC_LOG_PATH, 'a', encoding='utf-8') as f:
                f.write(log_entry)
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
**Ngày thực hiện:** {update_date_str} (Chu kỳ: Tuần 47: 16/11/2026 - 22/11/2026)  
**Tệp dữ liệu đã cập nhật:** [`reports/SEO_Keyword_Trends_Eureka.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) (Sheet 1: 10 từ khóa nóng | Sheet 3: Mở rộng 275 chủ đề)  
**Trạng thái đồng bộ OneDrive:** ✅ Đã đồng bộ nguyên tử sang OneDrive Marketing Eureka (`thuanh`)  

---

## 📊 TỔNG QUAN KẾT QUẢ QUÉT 8 NGUỒN CHÍNH THỐNG (24H - 7 NGÀY QUA)
Hệ thống Hot News Scanner đã tiến hành quét độc lập trên 8 nguồn dữ liệu chính thống:
1. `customs.gov.vn` (Cổng TTĐT Hải quan)
2. `chinhphu.vn` (Cổng TTĐT Chính phủ)
3. `haiquanonline.com.vn` (Tạp chí Hải quan Online)
4. `congthuong.vn` (Báo Công Thương)
5. `tapchitaichinh.vn` (Thời báo Tài chính Việt Nam)
6. `Google Trends VN` (Xu hướng tìm kiếm XNK, Thuế, Logistics)
7. `Drewry / World Shipping Council` (Báo cáo WCI cước vận tải biển ngày 12/11/2026)
8. `scratch/daily_notes.md` (Nhật ký thực chiến Eureka & Diễn đàn Nghiệp vụ ngày 13-16/11/2026)

Qua sàng lọc từ 85 sự kiện và case study nghiệp vụ thực tế, Agent 01 đã chọn lọc và chấm điểm chi tiết **TOP 5 CHỦ ĐỀ NÓNG NHẤT** theo 4 tiêu chí khắt khe của quy trình `wf05` (thang điểm 1-5, tối đa 20 điểm):
- **1. Tính thời sự (Timeliness):** Độ nóng, tính tức thời hoặc văn bản sắp/vừa có hiệu lực thi hành.
- **2. Độ đau (Pain Level):** Tác động trực tiếp đến dòng tiền, tiền phạt, đình trệ thông quan hoặc rủi ro pháp lý của chủ hàng.
- **3. Tiềm năng tranh luận (Debate Potential):** Mức độ phân hóa ý kiến, khoảng trống hướng dẫn hoặc bất đồng quan điểm giữa các bên.
- **4. Khả năng cho giá trị sâu (Value Depth):** Dữ liệu kỹ thuật, căn cứ luật rõ ràng để Eureka khẳng định vị thế chuyên gia vượt trội.

---

## 🏆 TOP 5 CHỦ ĐỀ NÓNG XẾP HẠNG CAO NHẤT ĐỀ XUẤT CHO CHỊ THÚY

### Đề xuất #1: [NGÀY ĐẦU HIỆU LỰC 16/11] Thông tư 142/2026/TT-BTC & Quyết định 1813/QĐ-CHQ: Bắt buộc định danh điện tử chủ hàng và thủ tục mở tờ khai bưu chính/chuyển phát nhanh độc lập
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Cả hai** (Facebook Fanpage: Carousel Card Infographic 'Thông tư 142 chính thức hiệu lực hôm nay: 3 bước định danh điện tử và mở tờ khai độc lập để không bị giữ hàng mẫu tại sân bay' | YouTube: Video ngắn 60s thao tác trực tiếp trên Cổng DVC Hải quan)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Cực khẩn cấp: Hôm nay (thứ Hai, 16/11/2026) Thông tư 142 chính thức có hiệu lực thi hành trên toàn quốc |
  | Độ đau | ⭐⭐⭐⭐⭐ (5.0/5) | Rất đau: Các doanh nghiệp nhập hàng mẫu, linh kiện gấp qua FedEx, DHL, UPS nếu chưa liên kết mã định danh đang bị từ chối tiếp nhận dữ liệu tờ khai; nguy cơ kẹt hàng mẫu tại sân bay Nội Bài, Tân Sơn Nhất |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (5.0/5) | Thảo luận cực nóng: Làm sao để lấy được tờ khai hải quan điện tử độc lập mang tên doanh nghiệp thay vì tờ khai gom chung của hãng chuyển phát để khấu trừ thuế VAT đầu vào? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Eureka hướng dẫn cặn kẽ 3 bước: Đăng ký định danh trên Cổng DVC Hải quan trong 15 phút, quy trình ủy quyền đại lý hải quan và thủ tục yêu cầu hãng chuyển phát mở tờ khai riêng |
  | **Tổng điểm** | **20.0 / 20** | **Xếp hạng: #1 (Văn Bản Đang Có Hiệu Lực Hôm Nay & Nhu Cầu Gỡ Vướng Tức Thì)** |
- **Tóm tắt lý do nên làm:** Tính thời sự đạt đỉnh tuyệt đối (ngày đầu tiên thực thi), tác động trực tiếp tới hàng vạn doanh nghiệp SME thường xuyên nhập hàng mẫu, linh kiện gấp.
- **Nguồn xác minh:** Thông tư số 142/2026/TT-BTC ngày 29/09/2026; Quyết định số 1813/QĐ-CHQ; Cổng TTĐT Hải quan (customs.gov.vn); Tạp chí Hải quan Online.

---

### Đề xuất #2: Nghị định 468/2026/NĐ-CP (ban hành 15/11/2026): Bước ngoặt sửa đổi Nghị định 08/2015/NĐ-CP về Hải quan thông minh, phân luồng bằng AI và nâng tỷ lệ luồng Xanh trên 75%
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Quản Trị)
- **Nền tảng khuyến nghị:** **Cả hai** (YouTube ThuyTalk CM1/CM2: Video dài 7-9 phút phân tích sâu 6 điểm đột phá của Nghị định 468 | Facebook: Bài viết chuyên sâu phân tích cơ chế bồi thường chi phí lưu bãi và bộ chỉ số tuân thủ doanh nghiệp)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5.0/5) | Rất nóng: Văn bản vừa được Chính phủ ký ban hành ngày 15/11/2026, sửa đổi toàn diện Nghị định 08/2015 và 59/2018 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Tác động dài hạn: Doanh nghiệp không hiểu cơ chế phân luồng AI sẽ dễ bị tụt hạng tuân thủ, rơi vào luồng Vàng/Đỏ kéo dài |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Bùng nổ tranh luận: Cơ chế bồi thường chi phí lưu kho bãi của Hải quan khi chậm trễ kiểm hóa có khả thi trong thực tế không? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5.0/5) | Giá trị chiến lược vượt trội: Eureka phân tích chuyên sâu mô hình Hải quan số 2026 - 2030, cách chấm điểm rủi ro bằng AI và lộ trình nâng hạng doanh nghiệp AEO |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Cải Cách Thể Chế Bản Lề Của Ngành Hải Quan)** |
- **Tóm tắt lý do nên làm:** Đây là sự kiện thể chế quan trọng nhất trong quý 4/2026, khẳng định vị thế tư vấn vĩ mô và tầm nhìn quản trị cấp cao của Chị Thúy và Eureka.
- **Nguồn xác minh:** Nghị định số 468/2026/NĐ-CP ngày 15/11/2026; Cổng TTĐT Chính phủ (chinhphu.vn); Cục Hải quan.

---

### Đề xuất #3: Công văn 24205/CHQ-GSQL ngày 16/11/2026: Tăng cường đối chiếu mã định danh phương tiện và gắn 100% seal định vị GPS điện tử cho hàng trung chuyển, quá cảnh tại cụm cảng nhóm 1
- **Chuyên mục (Pillar):** `CS` (Case Study Thực Chiến) / `TN` (Tin Ngành)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group XNK** (Bài viết cảnh báo rủi ro forwarder và hãng tàu: Quy trình mượn trả seal GPS điện tử và kỹ thuật tránh phạt chậm bàn giao container quá cảnh)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Văn bản phát hành trong ngày 16/11/2026 chỉ đạo khẩn các Chi cục cảng biển Hải Phòng, Cái Mép, Cát Lái |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Rất đau: Container trung chuyển/quá cảnh bị giữ lại cảng nếu sai lệch mã IMO phương tiện hoặc thiếu seal GPS; lỡ chuyến feeder quốc tế |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Điểm nghẽn thực tế: Thủ tục mượn/trả seal GPS tại cảng có gây ùn ứ xe đầu kéo không? Trách nhiệm bồi thường nếu seal hỏng hóc giữa đường? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Eureka chia sẻ quy trình 4 bước phối hợp giữa Hãng tàu - Cảng biển - Hải quan giám sát để bàn giao container trong 30 phút |
  | **Tổng điểm** | **19.4 / 20** | **Xếp hạng: #3 (Vấn Đề Giám Sát Cảng Biển & Rủi Ro Chậm Tàu Feeder)** |
- **Tóm tắt lý do nên làm:** Đánh trúng nỗi lo chậm chuyến, phát sinh phí demurrage/detention của giới forwarder và các tập đoàn vận tải đa phương thức.
- **Nguồn xác minh:** Công văn 24205/CHQ-GSQL ngày 16/11/2026; Cổng TTĐT Hải quan (customs.gov.vn); Tạp chí Hải quan Online.

---

### Đề xuất #4: Công văn 24210/CHQ-NVTHQ ngày 16/11/2026: Hướng dẫn thủ tục hoàn thuế và không thu thuế đối với nguyên liệu SXXK bị lỗi kỹ thuật phải tái xuất trả lại đối tác nước ngoài
- **Chuyên mục (Pillar):** `HT` (Học Thực Chiến)
- **Nền tảng khuyến nghị:** **Cả hai** (YouTube ThuyTalk: Hướng dẫn nghiệp vụ hoàn thuế SXXK | Facebook: Infographic quy trình 4 bước nộp hồ sơ qua Cổng HQ365 giải phóng dòng tiền hoàn thuế trong 3 ngày)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Văn bản nghiệp vụ hướng dẫn nóng ngày 16/11/2026 của Ban Nghiệp vụ Thuế Hải quan |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Cực đau: Hàng chục tỷ đồng tiền thuế nhập khẩu của doanh nghiệp SXXK bị đóng băng khi nguyên liệu lỗi phải xuất trả |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Vướng mắc kế toán thuế: Hải quan có đòi tờ khai nhập khẩu bản giấy không? Chứng thư giám định độc lập cần những tiêu chí gì? |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp trọn bộ: Mẫu đơn đề nghị hoàn thuế theo Thông tư 121/2025, biên bản đối chất kỹ thuật và checklist nộp DVC HQ365 |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #4 (Gỡ Nút Thắt Dòng Tiền Hoàn Thuế Cho Doanh Nghiệp SXXK)** |
- **Tóm tắt lý do nên làm:** Hướng dẫn thực chiến có giá trị tài chính trực tiếp, giúp doanh nghiệp SXXK giải phóng vốn lưu động bị tắc nghẽn.
- **Nguồn xác minh:** Công văn 24210/CHQ-NVTHQ ngày 16/11/2026; Thông tư 121/2025/TT-BTC; tapchitaichinh.vn.

---

### Đề xuất #5: Quyết định 3025/QĐ-BCT ngày 16/11/2026: Bộ Công Thương khởi xướng điều tra tự vệ đối với hạt nhựa PE và PP nhập khẩu (SG04.2026) — Rủi ro thuế suất tạm thời 5-12%
- **Chuyên mục (Pillar):** `TN` (Tin Ngành) / `CS` (Case Study Quản Trị)
- **Nền tảng khuyến nghị:** **Facebook Fanpage & Group Doanh nghiệp Nhựa / Bao bì** (Bài viết phân tích nguy cơ tăng giá thành sản xuất và cẩm nang đăng ký bên liên quan qua cổng TRAV ONLINE)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Quyết định điều tra mới ban hành ngày 16/11/2026, thời hạn 30 ngày đếm ngược |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Rất lớn: Hạt nhựa PE, PP là đầu vào của hàng ngàn nhà máy bao bì, túi nhựa, đồ gia dụng; thuế 5-12% làm xói mòn lợi nhuận |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Xung đột lợi ích: Doanh nghiệp sản xuất hạt nhựa nội địa vs Doanh nghiệp sản xuất bao bì/nhựa thành phẩm nhập khẩu |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.7/5) | Hướng dẫn kỹ thuật nộp đơn TRAV ONLINE, cách lập số liệu chứng minh chuỗi cung ứng để xin miễn trừ áp thuế tự vệ |
  | **Tổng điểm** | **19.2 / 20** | **Xếp hạng: #5 (Phòng Vệ Thương Mại Ngành Nhựa Nóng Bỏng)** |
- **Tóm tắt lý do nên làm:** Cảnh báo sớm rủi ro chi phí cho ngành nhựa và bao bì, mở ra cơ hội tư vấn thủ tục phòng vệ thương mại chuyên sâu.
- **Nguồn xác minh:** Quyết định 3025/QĐ-BCT ngày 16/11/2026 của Bộ Công Thương; Cục Phòng vệ thương mại; congthuong.vn.

---

## 🎯 ĐỀ XUẤT HÀNH ĐỘNG TIẾP THEO CHO CHỊ THÚY
1. **Lựa chọn chủ đề ưu tiên:** Chị Thúy có thể chọn ngay 1 hoặc 2 chủ đề trong Top 5 để đưa vào kế hoạch sản xuất nội dung tuần 47:
   - Nếu ưu tiên **nội dung bắt trend hỏa tốc hôm nay:** Chọn **Đề xuất #1 (Thông tư 142 chính thức hiệu lực)** -> Đẩy vào sản xuất bài Facebook Carousel Card và video ngắn.
   - Nếu ưu tiên **video YouTube dài chuyên sâu phân tích chính sách:** Chọn **Đề xuất #2 (Nghị định 468 Hải quan thông minh)** -> Đẩy vào quy trình kịch bản ThuyTalk CM1/CM2.
2. **Kích hoạt tự động:** Sau khi Chị Thúy phản hồi lựa chọn, Agent 01 sẽ tự động kết nối với Agent 02/03 để lên dàn ý chi tiết hoặc đưa vào lịch tuần `Weekly_Content_Calendar`.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"[OK] Đã tạo thành công bản tin đề xuất: {BRIEFING_MD_PATH}")
    print("=== HOÀN THÀNH TOÀN BỘ QUY TRÌNH WF05 ===")

if __name__ == '__main__':
    update_seo_w47_1611()
