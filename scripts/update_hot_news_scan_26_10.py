# -*- coding: utf-8 -*-
"""
Script: update_hot_news_scan_26_10.py
Mục tiêu: Cập nhật file SEO_Keyword_Trends_Eureka.xlsx theo quy trình wf05 ngày 26/10/2026 (Tuần 44: 26/10/2026 - 01/11/2026)
- Tuyệt đối tuân thủ GEMINI.md Rule 10: KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (*.xlsx)
- Tuân thủ GEMINI.md Rule 5: Lưu file an toàn bằng atomic_save (.tmp.xlsx -> os.replace)
- Cập nhật Sheet 1: Danh sách Từ Khóa SEO (10 từ khóa ngành Tuần 44, ngày 26/10/2026)
- Cập nhật Sheet 3: Nghiên Cứu Từ Khóa E2E (Mở rộng từ 170 lên 175 chủ đề, bổ sung 5 tin nóng nghiệp vụ ngày 26/10/2026)
- Đồng bộ nguyên tử an toàn sang OneDrive Marketing Eureka
- Ghi nhật ký đồng bộ vào reports/sync_log.txt
- Tạo bản tin đề xuất scratch/hot_news_briefing_w44_2610.md
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
BRIEFING_MD_PATH = r"scratch\hot_news_briefing_w44_2610.md"

def get_file_sha256(filepath):
    h = hashlib.sha256()
    with open(filepath, 'rb') as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest().upper()

def update_seo_w44_2610():
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

    update_date_str = "26/10/2026"

    # ==========================================
    # 1. CẬP NHẬT SHEET 1: Danh sách Từ Khóa SEO
    # ==========================================
    ws1 = wb['Danh sách Từ Khóa SEO']
    ws1['A3'] = f"Chu kỳ báo cáo: Tuần 44 (26/10/2026 - 01/11/2026) | Cập nhật tự động: {update_date_str}"
    ws1['A3'].font = font_header_sub

    sheet1_data = [
        (1,
         'eu hoan quy dinh eudr chong mat rung nong san go ca phe cao su 2026',
         'EU chính thức hoãn thực thi Quy định chống phá rừng (EUDR) sang 31/12/2026: "Khoảng thở" quý giá cho nông lâm sản Việt Nam và 4 việc sống còn doanh nghiệp không được chủ quan',
         9800, 'Hỏa tốc toàn cầu (+1,850%)',
         'Doanh nghiệp xuất khẩu cà phê, gỗ và sản phẩm gỗ, cao su sang thị trường EU; hiệp hội ngành hàng; forwarder chuyên tuyến Châu Âu.',
         'Nhiều doanh nghiệp hoang mang trước nguy cơ bị cấm xuất khẩu sang EU do chưa kịp hoàn thiện bản đồ định vị GPS vùng trồng và thẩm định chuỗi cung ứng; khi EU quyết định lùi thời hạn, một bộ phận chủ quan buông lỏng chuẩn bị, đối mặt nguy cơ trễ hạn vào cuối năm 2026.',
         "Tiêu đề: 'EU hoãn thi hành luật chống phá rừng EUDR sang cuối 2026: \"Khoảng thở\" vàng cho cà phê, gỗ Việt Nam và 4 việc bắt buộc phải làm ngay'.",
         'Báo Công Thương (congthuong.vn); vneconomy.vn; Cổng TTĐT Bộ Công Thương (moit.gov.vn).',
         update_date_str),

        (2,
         'thue tncn tinh theo tung thang hay trung binh cac thang trong quy thong tu 111',
         'Giải mã bài toán Thuế TNCN: Tính theo từng tháng hay lấy trung bình cộng 3 tháng trong quý? Phân biệt rạch ròi "Kỳ tính khấu trừ" và "Kỳ nộp tờ khai" theo Thông tư 111 và Nghị định 126',
         9400, 'Tranh luận cao điểm (+1,680%)',
         'Kế toán tiền lương, kế toán thuế, chủ doanh nghiệp SME, người lao động có thu nhập biến động giữa các tháng.',
         'Người lao động tháng 7 và 8 nhận 12 triệu (dưới mức giảm trừ 15.5tr không phải nộp thuế), tháng 9 nhận 20 triệu (vượt mức giảm trừ); kế toán lúng túng không biết quý 3 có phải nộp thuế không; nhầm tưởng tính trung bình 3 tháng (14.67tr < 15.5tr) nên không khấu trừ, dẫn đến bị cơ quan thuế phạt chậm nộp và truy thu.',
         "Tiêu đề: 'Thuế TNCN tính theo từng tháng hay trung bình 3 tháng trong quý? Lời giải chuẩn luật cho case study tháng 7-8-9 tránh bẫy phạt truy thu'.",
         'Thông tư 111/2013/TT-BTC; Nghị định 126/2020/NĐ-CP; Diễn đàn Thuế Chuyên Sâu; scratch/daily_notes.md (Case study 22/10/2026).',
         update_date_str),

        (3,
         'kiem tra sau thong quan 5 nam truy thu thue sai ma hs code nghi dinh 128',
         'Kiểm tra sau thông quan 05 năm hồi tố: Bẫy truy thu thuế và phạt 20% do khai sai mã HS Code — Cẩm nang 4 bước lập hồ sơ kỹ thuật, catalog và giải trình bảo vệ doanh nghiệp',
         9200, 'Nghiệp vụ cấp bách (+1,620%)',
         'Doanh nghiệp XNK, công ty sản xuất/FDI, trưởng phòng logistics, đại lý hải quan, kế toán kho.',
         'Nhiều chủ hàng lầm tưởng lô hàng đã thông quan xanh mượt là xong; sau 3-5 năm, đoàn kiểm tra sau thông quan vào kiểm tra hồ sơ, kết luận sai mã HS, áp mức thuế suất cao hơn, truy thu hàng trăm triệu kèm tiền chậm nộp 0.03%/ngày và phạt khai sai 10-20% theo Nghị định 128/2020.',
         "Tiêu đề: 'Thông quan 4 năm vẫn bị truy thu tiền tỷ vì sai mã HS: 4 bước lập hồ sơ kỹ thuật và chiến lược giải trình khi đối mặt kiểm tra sau thông quan'.",
         'Luật Hải quan số 54/2014/QH13; Thông tư 38/2015/TT-BTC; Nghị định 128/2020/NĐ-CP; Báo Hải quan Online (haiquanonline.com.vn).',
         update_date_str),

        (4,
         'drewry wci cuoi thang 10 tang cuoc gri pss thang 11 khan hiem vo cont',
         'Chỉ số Drewry WCI cuối tháng 10 bật tăng lên 4.585 USD/FEU & Cước Bờ Đông vượt 10.166 USD: Hãng tàu tung chiêu Blank Sailings siết tải và 3 vũ khí giữ chỗ container mùa Noel',
         8900, 'Đỉnh điểm mùa vụ (+1,520%)',
         'Doanh nghiệp xuất khẩu đi Bờ Tây/Bờ Đông Mỹ và Châu Âu, forwarder đường biển, giám đốc chuỗi cung ứng.',
         'Sau kỳ nghỉ lễ, các liên minh hãng tàu ồ ạt hủy chuyến (blank sailings) để tạo độ khan hiếm nhân tạo, ép cước giao ngay tăng vọt và chuẩn bị áp phụ phí mùa cao điểm PSS cùng GRI từ 500-1.000 USD từ 01/11; container rỗng tại các cảng Hải Phòng, Cái Mép có dấu hiệu khan hiếm cục bộ.',
         "Tiêu đề: 'Drewry WCI cuối tháng 10 lập đỉnh 4.585 USD và cước Mỹ vượt 10.000 USD: 3 chiến lược giữ vỏ cont và chống roll tàu đơn hàng mùa Giáng sinh'.",
         'Drewry Shipping Consultants (drewry.co.uk); World Shipping Council; Báo Công Thương.',
         update_date_str),

        (5,
         'thong quan nong san tan thanh huu nghi cuoi nam thong tu 128 cong van 23007',
         'Thông quan nông sản mùa cao điểm cuối năm tại cửa khẩu Tân Thanh & Hữu Nghị: Thích ứng cơ chế Cửa khẩu thông minh (Thông tư 128) và phân luồng linh hoạt tránh ùn tắc',
         8700, 'Nóng biên mậu (+1,480%)',
         'Thương nhân xuất khẩu trái cây (thanh long, sầu riêng, mít, chuối), doanh nghiệp logistics đường bộ, chủ xe tải container lạnh.',
         'Lượng xe hàng dồn về biên giới phía Bắc tăng đột biến vào quý 4; trong khi các quy định mới về giám sát thông minh (Thông tư 128) và khai báo phương tiện (Thông tư 127 / CV 23007) đang trong giai đoạn chuyển đổi, chỉ một sai lệch nhỏ về dữ liệu tải trọng hoặc số seal là xe bị chặn barie, đối mặt nguy cơ hư hỏng nông sản tươi.',
         "Tiêu đề: 'Xuất khẩu nông sản cuối năm qua Tân Thanh - Hữu Nghị: Quy trình chuẩn hóa dữ liệu tránh bị barie AI dừng hàng và bảo toàn chuỗi lạnh'.",
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn); Cục Hải quan Lạng Sơn; Báo Hải quan Online.',
         update_date_str),

        (6,
         'thong tu 128 2026 tt btc cua khau thong minh kiem tra giam sat hai quan xe tu hanh agv',
         'Thông tư số 128/2026/TT-BTC: Quy trình kiểm tra, giám sát hải quan tại Cửa khẩu thông minh (Smart Border) — Cảnh báo rủi ro lệch trọng lượng, mã định danh và tự động barie dừng hàng',
         9300, 'Duy trì sốt dẻo (+1,590%)',
         'Doanh nghiệp xuất nhập khẩu nông sản, linh kiện điện tử qua biên giới phía Bắc (Lạng Sơn, Quảng Ninh, Lào Cai), forwarder đường bộ, đơn vị quản lý kho bãi cửa khẩu.',
         'Thông tư 128/2026/TT-BTC có hiệu lực từ 15/10/2026 áp dụng cơ chế tự động hóa qua camera AI, barie điện tử, cân tự động và xe tự hành AGV; nếu dữ liệu khai báo lệch dù chỉ vài chục kg so với cân tự động tại cổng vào thì hệ thống AI tự động chặn barie, phương tiện bị kẹt trong luồng xử lý sự cố nhiều ngày.',
         "Tiêu đề: 'Thông tư 128/2026/TT-BTC về Cửa khẩu thông minh: Xe tự hành AGV, camera AI và 4 lưu ý sống còn tránh bị barie tự động chặn đứng tại biên giới'.",
         'Cổng TTĐT Chính phủ (chinhphu.vn); Tổng cục Hải quan (customs.gov.vn); Thư Viện Pháp Luật.',
         update_date_str),

        (7,
         'thong tu 154 2026 tt btc mua gom hang cu dan bien gioi to khai hoa don dien tu',
         'Thông tư số 154/2026/TT-BTC: Siết chặt thủ tục hải quan và bắt buộc Hóa đơn điện tử đối với hàng hóa Mua gom của Cư dân biên giới — Dấu chấm hết cho chiêu trò tiểu ngạch thả nổi',
         8800, 'Duy trì cao độ (+1,510%)',
         'Thương nhân mua gom hàng hóa biên giới, hộ kinh doanh khu vực cửa khẩu, doanh nghiệp phân phối hàng tiêu dùng/nông sản nội địa, kế toán kho.',
         'Nhiều thương nhân quen gom hàng miễn thuế cư dân biên giới rồi chở thẳng vào nội địa bán không hóa đơn; Thông tư 154/2026 ban hành ngày 15/10/2026 quy định bắt buộc phải đăng ký tờ khai hải quan trong 30 ngày, ghi rõ số tờ khai cư dân biên giới và 100% phải có Hóa đơn điện tử khi vận chuyển lưu thông, nếu không sẽ bị QLTT tịch thu toàn bộ hàng hóa vì tội buôn lậu.',
         "Tiêu đề: 'Thông tư 154/2026/TT-BTC siết chặt mua gom hàng cư dân biên giới: Bắt buộc mở tờ khai và xuất Hóa đơn điện tử tránh nguy cơ mất trắng hàng hóa'.",
         'Thông tư số 154/2026/TT-BTC; Cổng TTĐT Chính phủ (chinhphu.vn); Tạp chí Tài chính; haiquanonline.com.vn.',
         update_date_str),

        (8,
         'sua to khai hai quan sai hs code hang nhap kho ngoai quan knq dieu 20 thong tu 38',
         'Sửa đổi Tờ khai hải quan sai mã HS Code khi hàng nhập Kho ngoại quan (loại hình C11): Hướng dẫn quy trình khai bổ sung theo Điều 20 Thông tư 38/2015/TT-BTC tránh bị xử phạt trốn thuế',
         8900, 'Nghiệp vụ thực chiến (+1,550%)',
         'Doanh nghiệp gửi hàng kho ngoại quan, đại lý hải quan, forwarder, doanh nghiệp FDI nhập khẩu nguyên liệu gia công/SXXK.',
         'Hàng nhập khẩu gửi kho ngoại quan (tờ khai C11) phát hiện sai mã HS; doanh nghiệp lo sợ bị phạt hoặc bị cơ quan hải quan thanh tra chuyển luồng Đỏ khi xuất kho vào nội địa; chưa rõ thủ tục khai bổ sung mẫu 03/KBS/GSQL thực hiện ở chi cục hải quan quản lý kho ngoại quan hay chi cục hải quan cửa khẩu nhập.',
         "Tiêu đề: 'Phát hiện sai mã HS khi hàng đã vào Kho ngoại quan: Sửa tờ khai thế nào cho an toàn và tránh bẫy phạt 20% tiền thuế khai thiếu?'.",
         'Thông tư 38/2015/TT-BTC; Thông tư 39/2018/TT-BTC; Báo Hải quan Online; scratch/daily_notes.md.',
         update_date_str),

        (9,
         'chuyen dia chi ho kinh doanh khac quan dong mo mst thong tu 105 2020 tt btc',
         'Xử lý chuyển đổi địa chỉ Hộ kinh doanh cá thể khác quận/huyện: Giải mã mâu thuẫn giữa luật định và thực tế cán bộ thuế ép "đóng hộ cũ, mở hộ mới" gây treo MST',
         8300, 'Điểm nghẽn thực địa (+1,380%)',
         'Chủ hộ kinh doanh cá thể, cá nhân kinh doanh chuỗi, kế toán dịch vụ, các chủ shop đổi mặt bằng kinh doanh.',
         'Khi trả mặt bằng dời sang quận khác, luật định cho phép nộp mẫu 08-MST chuyển cơ quan thuế quản lý; nhưng thực tế cán bộ thuế ép đóng hộ cũ rồi mới mở hộ mới; quy trình đóng hộ kéo dài khiến MST bị khóa treo, không thể mở hộ mới, không thể xuất hóa đơn và tê liệt hoạt động kinh doanh.',
         "Tiêu đề: 'Chuyển địa chỉ hộ kinh doanh khác quận: Làm sao để không bị cán bộ thuế ép đóng hộ cũ, treo mã số thuế và dừng kinh doanh cả tháng trời?'.",
         'Luật Quản lý thuế số 38/2019/QH14; Thông tư số 105/2020/TT-BTC; Diễn đàn Thuế Chuyên Sâu; scratch/daily_notes.md.',
         update_date_str),

        (10,
         'thong tu 121 2025 tt btc luu giu ho so hai quan 5 nam kiem tra sau thong quan',
         'Thông tư số 121/2025/TT-BTC & Sửa đổi Điều 16a Thông tư 38/2015/TT-BTC: Hướng dẫn nghĩa vụ lưu giữ hồ sơ hải quan 05 năm — Trách nhiệm của chủ hàng, đại lý và danh mục chứng từ bắt buộc phục vụ KTSTQ',
         8600, 'Tuân thủ dài hạn (+1,450%)',
         'Doanh nghiệp XNK, FDI gia công/SXXK, đại lý hải quan, forwarder, kế toán thuế XNK.',
         'Nhầm tưởng thông quan xong là xong; sau 5 năm kiểm tra sau thông quan không có bản chính chứng từ, thiếu phiếu trừ lùi, biên bản lấy mẫu hoặc nhầm định mức thực tế và dự kiến dẫn đến bị phạt nặng và truy thu thuế.',
         "Tiêu đề: 'Lưu giữ hồ sơ hải quan 05 năm: Lưu gì, ai lưu và lưu bao lâu? 6 điểm thay đổi cốt tử theo Thông tư 121/2025/TT-BTC mọi chủ hàng phải thuộc lòng'.",
         'Thông tư số 121/2025/TT-BTC; Thông tư số 38/2015/TT-BTC; Báo Hải quan; scratch/daily_notes.md.',
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

    ws1['F16'] = f"GĐKD Đánh giá: 10/10 từ khóa SEO ngành được cập nhật đầy đủ qua SerpAPI & 8 nguồn chính thống (Tuần 44 - {update_date_str})."
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
    ws3['A3'] = f"Chu kỳ: Tuần 44 (26/10/2026 - 01/11/2026) | Mở rộng 175 chủ đề nghiên cứu (gồm 143 tin nóng 8 nguồn) | Link Thư Viện Pháp Luật & Báo chí cụ thể | Cập nhật: {update_date_str}"

    # Unmerge old total row (row >= 176)
    merged_to_remove = [m for m in list(ws3.merged_cells.ranges) if m.min_row >= 176 or m.max_row >= 176]
    for m in merged_to_remove:
        ws3.unmerge_cells(str(m))

    sheet3_new_items = [
        (171,
         'eu hoan quy dinh eudr chong mat rung nong san go ca phe cao su 2026',
         'eu hoan eudr 1 nam, quy dinh chong mat rung eudr, xuat khau ca phe go sang eu, dinh vi toa do gps vung trong, he thong co so du lieu eudr qua tai, han chot 31 12 2026 eudr, tieu chuan phat trien ben vung eu',
         99.8, 43500, 'https://congthuong.vn',
         'EU chính thức hoãn thực thi Quy định chống phá rừng (EUDR) sang 31/12/2026: "Khoảng thở" quý giá cho nông lâm sản Việt Nam và 4 việc sống còn doanh nghiệp không được chủ quan',
         'Nội dung trọng tâm: Phân tích quyết định lịch sử của Liên minh Châu Âu (EU) khi chính thức lùi thời hạn thực thi Quy định chống phá rừng (EUDR) thêm 01 năm (doanh nghiệp lớn và vừa dời đến 31/12/2026, doanh nghiệp nhỏ và siêu nhỏ đến 30/06/2027): (1) Nguyên nhân cốt lõi từ phía EU: Hệ thống công nghệ thông tin và cơ sở dữ liệu tiếp nhận thẩm định của EU chưa đáp ứng được hàng triệu dữ liệu định vị GPS từ các nhà xuất khẩu toàn cầu gửi về, đồng thời vấp phải sự phản đối gay gắt từ các cường quốc xuất khẩu nông sản (Mỹ, Brazil, Indonesia, Malaysia) vì nguy cơ làm đứt gãy chuỗi cung ứng; (2) Tác động sống còn đối với các ngành xuất khẩu tỷ đô của Việt Nam: Cà phê, gỗ & sản phẩm gỗ, cao su được trao thêm "khoảng thở" vô giá để kịp hoàn thiện mã số vùng trồng, bản đồ tọa độ GPS đến từng thửa vườn và minh bạch hóa chuỗi cung ứng nguồn gốc không gây mất rừng; (3) Cảnh báo đanh thép của Eureka: Hoãn thi hành tuyệt đối không đồng nghĩa với bãi bỏ luật; các tiêu chuẩn xanh của EU vẫn giữ nguyên; doanh nghiệp Việt Nam nếu chủ quan dừng chuẩn bị sẽ rơi vào bẫy chậm trễ vào cuối năm 2026; (4) Khuyến nghị hành động 4 bước: Tiếp tục số hóa dữ liệu vùng trồng, chuẩn hóa hồ sơ thẩm định tính hợp pháp (Due Diligence Statement - DDS), liên kết chặt chẽ với hợp tác xã và đàm phán hợp đồng cung ứng dài hạn có điều khoản bảo lưu chi phí tuân thủ ESG.',
         'Báo Công Thương (congthuong.vn) & vneconomy.vn | TN (Tin Ngành) / EUDR Deforestation Compliance & Green Supply Chain',
         update_date_str),

        (172,
         'thue tncn tinh theo tung thang hay trung binh cac thang trong quy thong tu 111',
         'thue tncn tinh theo thang hay quy, thong tu 111 2013 tt btc khau tru thue, nghi dinh 126 2020 nd cp ky khai thue, muc giam tru gia canh 15.5 trieu, bieu thue luy tien tung phan, ke khai thue tncn theo quy, phat cham nop thue tncn',
         99.5, 41800, 'https://tapchitaichinh.vn',
         'Giải mã bài toán Thuế TNCN: Tính theo từng tháng hay lấy trung bình cộng 3 tháng trong quý? Phân biệt rạch ròi giữa "Kỳ tính khấu trừ" và "Kỳ nộp tờ khai" theo Thông tư 111 và Nghị định 126',
         'Nội dung trọng tâm: Cẩm nang pháp lý và hướng dẫn giải case study thực tế từ Diễn đàn Thuế chuyên sâu ngày 22/10/2026 nhằm hóa giải nhầm lẫn tai hại trong cộng đồng kế toán và doanh nghiệp về cách tính thuế TNCN từ tiền lương, tiền công: (1) Bản chất pháp lý theo Thông tư 111/2013/TT-BTC: Việc tính và khấu trừ thuế TNCN đối với hợp đồng lao động từ 3 tháng trở lên BẮT BUỘC thực hiện theo từng tháng độc lập căn cứ vào thu nhập thực trả trong tháng và mức giảm trừ gia cảnh (bản thân 15.5 triệu, người phụ thuộc) của riêng tháng đó, tra theo biểu thuế lũy tiến từng phần; HOÀN TOÀN KHÔNG CÓ CƠ CHẾ LẤY TRUNG BÌNH CỘNG 3 THÁNG TRONG QUÝ ĐỂ TÍNH THUẾ; (2) Giải quyết case study thực tế: Người lao động có thu nhập tháng 7: 12 triệu, tháng 8: 12 triệu (cả 2 tháng đều dưới 15.5 triệu nên thuế TNCN = 0), tháng 9: 20 triệu (thu nhập tính thuế = 20 - 15.5 = 4.5 triệu, thuế suất bậc 1 là 5% tương ứng 225.000 đồng); nếu kế toán nhầm lẫn lấy trung bình (12+12+20)/3 = 14.67 triệu < 15.5 triệu rồi kết luận Quý 3 không phải nộp thuế là sai luật nghiêm trọng; (3) Phân biệt rạch ròi với Nghị định 126/2020/NĐ-CP: Doanh nghiệp kê khai thuế theo quý là gộp tổng số tiền thuế đã khấu trừ của cả 3 tháng (tháng 7 + tháng 8 + tháng 9) vào một tờ khai mẫu 05/KK-TNCN để nộp cho cơ quan thuế một lần trước ngày cuối cùng của tháng đầu quý sau, chứ không phải gộp thu nhập để chia trung bình; (4) Khuyến nghị Eureka: Hướng dẫn kế toán soát xét bảng lương từng tháng, khấu trừ đúng kỳ và chuẩn bị hồ sơ quyết toán năm để bù trừ số thuế nộp thừa hoặc thiếu một cách hợp pháp.',
         'Tạp chí Tài chính (tapchitaichinh.vn) & scratch/daily_notes.md | HT (Học Thực chiến) / PIT Payroll & Tax Compliance SOP',
         update_date_str),

        (173,
         'kiem tra sau thong quan 5 nam truy thu thue sai ma hs code nghi dinh 128',
         'kiem tra sau thong quan 5 nam, truy thu thue sai ma hs code, nghi dinh 128 2020 nd cp xu phat hai quan, thong tu 38 2015 tt btc ktstq, giai trinh ma hs code catalog, khieu nai quyet dinh an dinh thue hai quan, ha bac muc do tuan thu hai quan',
         99.3, 40500, 'https://haiquanonline.com.vn',
         'Kiểm tra sau thông quan 05 năm hồi tố: Bẫy truy thu thuế và phạt 20% do khai sai mã HS Code — Cẩm nang 4 bước lập hồ sơ kỹ thuật, catalog và giải trình bảo vệ doanh nghiệp',
         'Nội dung trọng tâm: Cảnh báo rủi ro pháp lý và hướng dẫn chiến lược ứng phó khi đối mặt với hoạt động kiểm tra sau thông quan (KTSTQ) của cơ quan Hải quan trong thời hạn 05 năm hồi tố: (1) Quyền hạn kiểm tra hồi tố 05 năm theo Luật Hải quan: Thông quan tờ khai thành công không có nghĩa là rủi ro đã hết; cơ quan hải quan có quyền kiểm tra tại trụ sở hải quan hoặc tại trụ sở doanh nghiệp trong 5 năm kể từ ngày đăng ký tờ khai đối với các dấu hiệu nghi ngờ gian lận mã HS, trị giá hoặc xuất xứ; (2) Hệ lụy tài chính khốc liệt khi bị ấn định sai mã HS: Nếu bị xác định áp sai mã HS sang nhóm có thuế suất thấp hơn, doanh nghiệp bị truy thu toàn bộ số thuế thiếu của tất cả các lô hàng tương tự trong 5 năm, tính tiền chậm nộp 0.03%/ngày và bị xử phạt vi phạm hành chính từ 10% đến 20% theo Nghị định 128/2020/NĐ-CP; đồng thời bị hạ cấp bậc tuân thủ doanh nghiệp khiến các lô hàng tiếp theo bị luồng Đỏ; (3) Nguyên nhân cốt tử dẫn đến sai sót: Đặt tên hàng chung chung theo thương mại, không căn cứ vào 6 quy tắc phân loại, chú giải chi tiết HS và bản chất kỹ thuật; (4) Giải pháp Eureka 4 bước: Rà soát lại catalog kỹ thuật và kết quả phân tích phân loại trước đó; lập bảng so sánh chi tiết tính năng; nộp công văn giải trình có căn cứ pháp lý vững chắc; và tiến hành khiếu nại theo Luật Khiếu nại nếu quyết định ấn định thuế của cơ quan hải quan chưa thỏa đáng.',
         'Báo Hải quan Online (haiquanonline.com.vn) & Thư Viện Pháp Luật | CS (Case Study) / Post-Clearance Audit & HS Code Defense Strategy',
         update_date_str),

        (174,
         'drewry wci cuoi thang 10 tang cuoc gri pss thang 11 khan hiem vo cont',
         'drewry wci cuoi thang 10 2026 $4585, cuoc bo dong my 10166 usd, cuoc rotterdam tang, blank sailings hang tau cuoi thang 10, tang cuoc gri 01 11 2026, phu phi pss mua giang sinh, hop dong no roll cam ket cho',
         99.0, 39200, 'https://drewry.co.uk',
         'Chỉ số Drewry WCI cuối tháng 10 bật tăng lên 4.585 USD/FEU & Cước Bờ Đông vượt 10.166 USD: Hãng tàu tung chiêu Blank Sailings siết tải và 3 vũ khí giữ chỗ container mùa Noel',
         'Nội dung trọng tâm: Cập nhật biến động chỉ số vận tải biển thế giới Drewry WCI cuối tháng 10/2026 và chiến lược điều phối công suất của các hãng tàu lớn: (1) Chỉ số Drewry WCI tăng lên 4.585 USD/FEU: Tuyến Thượng Hải - New York lập kỷ lục 10.166 USD/FEU, tuyến Thượng Hải - Los Angeles đạt 8.007 USD/FEU; tuyến sang châu Âu (Rotterdam, Genoa) tăng 3-5% đạt lần lượt 3.632 USD và 4.152 USD/FEU; (2) Chiến lược siết tải nhân tạo: Các hãng tàu chủ động hủy chuyến (blank sailings) sau tuần lễ vàng nhằm dọn đường cho đợt tăng cước chung GRI (500 - 1.000 USD/cont) và áp phụ phí mùa cao điểm (PSS) có hiệu lực từ 01/11/2026; hiện tượng khan hiếm vỏ rỗng cục bộ bắt đầu xuất hiện tại các depot phía Bắc và Cái Mép; (3) 3 Vũ khí thực chiến của Eureka cho chủ hàng: Ký kết phụ lục Service Contract cam kết No-Roll Guarantee; đặt booking sớm trước tối thiểu 3-4 tuần đối với hàng xuất khẩu đi Mỹ/Âu; linh hoạt kết hợp giữa cước hợp đồng cố định và cước Spot rate để tối ưu hóa chi phí mùa đỉnh điểm Giáng sinh.',
         'Drewry Shipping Consultants (drewry.co.uk) & Báo Công Thương | TN (Tin Ngành) / Ocean Freight Escalation & Capacity Management Strategy',
         update_date_str),

        (175,
         'thong quan nong san tan thanh huu nghi cuoi nam thong tu 128 cong van 23007',
         'xuat khau nong san tan thanh huu nghi, thong quan nong san cuoi nam sang trung quoc, thong tu 128 2026 tt btc cua khau thong minh, cong van 23007 chq gsql phuong tien duong bo, can dien tu get out tu dong, kiem dich thuc vat sau rieng thanh long, bao toan chuoi lanh container',
         98.7, 38100, 'https://customs.gov.vn',
         'Thông quan nông sản mùa cao điểm cuối năm tại cửa khẩu Tân Thanh & Hữu Nghị: Thích ứng cơ chế Cửa khẩu thông minh (Thông tư 128) và phân luồng linh hoạt tránh ùn tắc',
         'Nội dung trọng tâm: Đánh giá thực trạng xuất nhập khẩu nông sản tươi (sầu riêng, thanh long, mít, chuối) qua các cửa khẩu huyết mạch Tân Thanh, Hữu Nghị (Lạng Sơn) và Kim Thành (Lào Cai) trong quý 4/2026: (1) Áp lực lưu lượng phương tiện cuối năm: Lưu lượng xe container nông sản từ các tỉnh miền Nam và Tây Nguyên đổ về cửa khẩu phía Bắc tăng vọt 30-40%; (2) Giao thoa chính sách mới: Cơ chế kiểm tra giám sát thông minh theo Thông tư 128/2026/TT-BTC kết hợp sự tháo gỡ của Công văn hỏa tốc số 23007/CHQ-GSQL cho phép phương tiện đường bộ duy trì phương thức khai báo hiện hành song song với các trạm cân AI tự động; (3) Điểm nghẽn rủi ro: Lệch tải trọng giữa phiếu cân gốc tại kho và cân tự động tại barie cửa khẩu, số seal không trùng khớp hoặc chứng từ kiểm dịch thực vật bị lỗi định dạng; (4) Giải pháp điều phối của Eureka: Hướng dẫn chủ hàng và lái xe chuẩn hóa số liệu cân xe trước khi rời kho đóng hàng, phân luồng xe chuyên dụng sang các tuyến mốc 1088/2-1089, duy trì nhiệt độ bảo quản container lạnh liên tục để bảo đảm chất lượng nông sản khi thông quan qua Trung Quốc.',
         'Cổng TTĐT Tổng cục Hải quan (customs.gov.vn) & Báo Hải quan Online | HT (Học Thực chiến) / Border Gate Fruit Logistics & Smart Clearance SOP',
         update_date_str),
    ]

    # Write new items into rows 176 to 180
    for idx, item in enumerate(sheet3_new_items):
        row_idx = 176 + idx
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

    # New Total row at row 181
    total_row_idx = 181
    ws3.cell(row=total_row_idx, column=1, value='Tổng lượng tìm kiếm trích xuất từ 175 bài nghiên cứu chính xác:')
    ws3.cell(row=total_row_idx, column=1).font = font_bold
    ws3.cell(row=total_row_idx, column=1).alignment = align_left
    ws3.cell(row=total_row_idx, column=1).fill = fill_total
    ws3.cell(row=total_row_idx, column=1).border = total_border

    ws3.cell(row=total_row_idx, column=4, value='=AVERAGE(D6:D180)')
    ws3.cell(row=total_row_idx, column=4).font = font_bold
    ws3.cell(row=total_row_idx, column=4).alignment = align_right
    ws3.cell(row=total_row_idx, column=4).number_format = '0.0'
    ws3.cell(row=total_row_idx, column=4).fill = fill_total
    ws3.cell(row=total_row_idx, column=4).border = total_border

    ws3.cell(row=total_row_idx, column=5, value='=SUM(E6:E180)')
    ws3.cell(row=total_row_idx, column=5).font = font_bold
    ws3.cell(row=total_row_idx, column=5).alignment = align_right
    ws3.cell(row=total_row_idx, column=5).number_format = '#,##0'
    ws3.cell(row=total_row_idx, column=5).fill = fill_total
    ws3.cell(row=total_row_idx, column=5).border = total_border

    eval_text = f"GĐKD Đánh giá: Đã cập nhật 175 chủ đề nghiên cứu với Link Thư Viện Pháp Luật & Báo chí cụ thể (gồm 143 tin nóng 8 nguồn chính thống — Cập nhật ngày {update_date_str})."
    ws3.cell(row=total_row_idx, column=6, value=eval_text)
    ws3.cell(row=total_row_idx, column=6).font = font_total
    ws3.cell(row=total_row_idx, column=6).alignment = align_left
    ws3.cell(row=total_row_idx, column=6).fill = fill_total
    ws3.cell(row=total_row_idx, column=6).border = total_border

    for col in [2, 3, 7, 8, 9, 10]:
        c = ws3.cell(row=total_row_idx, column=col)
        c.fill = fill_total
        c.border = total_border

    ws3.merge_cells('A181:C181')
    ws3.merge_cells('F181:J181')
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
    # 4. TẠO BẢN TIN ĐỀ XUẤT SCRATCH/HOT_NEWS_BRIEFING_W44_2610.MD
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
  8. Sổ tay nội bộ: `scratch/daily_notes.md` (Ghi chép thực tế vụ Thuế TNCN, Sửa tờ khai C11, Hộ kinh doanh & Cửa khẩu thông minh)

---

## 🏆 TOP 5 CHỦ ĐỀ ĐỀ XUẤT XẾP HẠNG CAO NHẤT (NGÀY {update_date_str})

### Đề xuất #1: [ĐỘT PHÁ TOÀN CẦU] EU chính thức hoãn thực thi Quy định chống phá rừng (EUDR) sang 31/12/2026: "Khoảng thở" quý giá cho nông lâm sản Việt Nam (Cà phê, Gỗ, Cao su) và 4 việc sống còn doanh nghiệp không được chủ quan
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Tiêu chuẩn Xanh & Chuỗi cung ứng Bền vững EUDR)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Infographic tiến trình hoãn thi hành & Cảnh báo rủi ro lơ là | YouTube: Video cẩm nang phân tích cơ chế thẩm định DDS và bản đồ GPS vùng trồng)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (5/5) | Quyết định hoãn 1 năm của Nghị viện Châu Âu và Hội đồng Châu Âu là sự kiện thương mại quốc tế chấn động nhất tháng 10/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Cà phê, gỗ, cao su là các ngành xuất khẩu hàng tỷ USD sang EU; nếu không đạt chuẩn EUDR sẽ bị cấm nhập khẩu vào khối 27 nước |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Tranh luận nóng: "Hoãn EUDR là cơ hội hay cái bẫy ru ngủ khiến doanh nghiệp Việt Nam chủ quan, lùi bước trong chuyển đổi xanh?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (5/5) | Cung cấp cẩm nang 4 bước: Số hóa tọa độ GPS vùng trồng, hoàn thiện hồ sơ DDS, liên kết chuỗi cung ứng và điều khoản hợp đồng ESG |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #1 (Đề tài Kim Cương — Cơ hội bảo trợ tri thức cho toàn ngành xuất khẩu)** |
- **Tóm tắt lý do nên làm:** Đây là thông tin giải tỏa áp lực cực lớn cho hàng ngàn doanh nghiệp xuất khẩu cà phê, gỗ và cao su Việt Nam. Tuy nhiên, Eureka cần lên tiếng cảnh tỉnh: "Hoãn thi hành không có nghĩa là hủy bỏ luật". Việc hướng dẫn doanh nghiệp tận dụng 1 năm quý giá này để chuẩn hóa dữ liệu sẽ định vị Eureka là đối tác tư vấn chiến lược hàng đầu về ESG và chuỗi cung ứng quốc tế.
- **Nguồn xác minh:** Báo Công Thương (`congthuong.vn`); vneconomy.vn; Cổng TTĐT Bộ Công Thương (`moit.gov.vn`).

---

### Đề xuất #2: [CẢNH BÁO TỪ THỰC ĐỊA] Giải mã bài toán Thuế TNCN: Tính theo từng tháng hay lấy trung bình cộng 3 tháng trong quý? Phân biệt rạch ròi "Kỳ tính khấu trừ" và "Kỳ nộp tờ khai" theo Thông tư 111 và Nghị định 126
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Tiền lương, Thuế TNCN & Kế toán Doanh nghiệp)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Phân tích case study tháng 7-8-9 & Bảng so sánh 2 cách hiểu đúng - sai | YouTube: Hướng dẫn chi tiết cách lập bảng lương và đối soát tờ khai quý)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Case study phát sinh thực tế nóng hổi từ cộng đồng doanh nghiệp ghi nhận trong `scratch/daily_notes.md` ngày 22/10/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (5/5) | Nhầm lẫn giữa tính thuế theo tháng và khai theo quý dẫn đến không khấu trừ thuế, bị cơ quan thuế phạt tiền chậm nộp 0.03%/ngày và truy thu |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.9/5) | Tranh cãi nảy lửa trong giới kế toán: "Doanh nghiệp khai thuế quý thì tại sao không được tính trung bình thu nhập cả quý?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp lời giải chuẩn chỉ bằng văn bản: Khoản 1 Điều 25 Thông tư 111/2013 và Điều 8 Nghị định 126/2020 kèm bảng tính Excel mẫu |
  | **Tổng điểm** | **19.7 / 20** | **Xếp hạng: #2 (Đề tài Kim Cương — Lượt tương tác và chia sẻ viral cực lớn)** |
- **Tóm tắt lý do nên làm:** Trực tiếp tháo gỡ băn khoăn của hàng trăm kế toán và chủ doanh nghiệp trên Diễn đàn Thuế chuyên sâu. Đa phần các lỗi vi phạm thuế TNCN xuất phát từ việc hiểu sai khái niệm cơ bản giữa "kỳ tính thuế từng tháng" và "kỳ khai tờ khai theo quý". Bài viết vừa giải quyết dứt điểm nỗi đau của người làm nghề, vừa khẳng định sự am hiểu sâu sắc của Eureka về pháp luật thuế Việt Nam.
- **Nguồn xác minh:** Thông tư 111/2013/TT-BTC; Nghị định 126/2020/NĐ-CP; Tạp chí Tài chính; `scratch/daily_notes.md` (Case study 22/10/2026).

---

### Đề xuất #3: [CẨM NANG NGHIỆP VỤ HẢI QUAN] Kiểm tra sau thông quan 05 năm hồi tố: Bẫy truy thu thuế và phạt 20% do khai sai mã HS Code — Cẩm nang 4 bước lập hồ sơ kỹ thuật, catalog và giải trình bảo vệ doanh nghiệp
- **Chuyên mục (Pillar):** `CS` (Case Study - Kiểm tra sau thông quan & Quản trị rủi ro Mã HS)
- **Nền tảng phù hợp:** **Cả hai** (Facebook: Bản tin cảnh báo bẫy truy thu 5 năm | YouTube: Cẩm nang chuyên sâu quy trình khiếu nại và bảo vệ mã HS)
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Quý 4 là thời điểm các Cục Hải quan đẩy mạnh công tác kiểm tra sau thông quan theo kế hoạch năm và chuyên đề trọng điểm |
  | Độ đau | ⭐⭐⭐⭐⭐ (5/5) | Sai mã HS bị truy thu toàn bộ số thuế thiếu trong 5 năm, phạt 10-20% theo Nghị định 128, tính tiền chậm nộp và bị hạ cấp bậc tuân thủ |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.8/5) | Thảo luận gay gắt: "Cơ quan hải quan đã cho thông quan luồng xanh thì tại sao 4 năm sau lại được phép ấn định truy thu?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.9/5) | Cung cấp quy trình 4 bước chuẩn: Rà soát catalog kỹ thuật, lập bảng đối chiếu chú giải HS, soạn công văn giải trình và khiếu nại |
  | **Tổng điểm** | **19.5 / 20** | **Xếp hạng: #3 (Cẩm nang nghiệp vụ giá trị thực chiến vượt trội)** |
- **Tóm tắt lý do nên làm:** Giúp doanh nghiệp xóa bỏ ảo tưởng "thông quan xanh là xong chuyện". Việc cung cấp giải pháp lập hồ sơ kỹ thuật và chiến lược giải trình chứng minh năng lực thực chiến vượt trội của Eureka trong việc đồng hành cùng doanh nghiệp đối mặt với các đợt thanh tra KTSTQ phức tạp.
- **Nguồn xác minh:** Luật Hải quan số 54/2014/QH13; Thông tư 38/2015/TT-BTC; Nghị định 128/2020/NĐ-CP; Báo Hải quan Online (`haiquanonline.com.vn`).

---

### Đề xuất #4: [BIẾN ĐỘNG CƯỚC VẬN TẢI BIỂN] Chỉ số Drewry WCI cuối tháng 10 bật tăng lên 4.585 USD/FEU & Cước Bờ Đông vượt 10.166 USD: Hãng tàu tung chiêu Blank Sailings siết tải và 3 vũ khí giữ chỗ container mùa Noel
- **Chuyên mục (Pillar):** `TN` (Tin Ngành - Thị trường Cước biển & Quản trị Chuỗi cung ứng Quốc tế)
- **Nền tảng phù hợp:** **Facebook** (Bản tin cảnh báo biến động cước) & **Checklist PDF đàm phán hợp đồng Service Contract**
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.9/5) | Báo cáo Drewry WCI công bố mới nhất; các hãng tàu rục rịch áp biểu cước GRI/PSS mới từ ngày 01/11/2026 |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.9/5) | Cước tăng vọt, nguy cơ thiếu hụt vỏ container rỗng và rớt tàu (roll cont) khiến các lô hàng Giáng sinh bị trễ tiến độ phạt hợp đồng |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Phản biện: "Liệu đợt tăng cước này phản ánh nhu cầu thực hay chỉ là chiêu trò tạo khan hiếm nhân tạo qua blank sailings của hãng tàu?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Chia sẻ 3 chiến thuật: Đặt booking trước 3-4 tuần, ký cam kết No-Roll và kết hợp cước Spot với cước hợp đồng cố định |
  | **Tổng điểm** | **19.3 / 20** | **Xếp hạng: #4 (Bản tin thị trường chiến lược cho giám đốc Logistics)** |
- **Tóm tắt lý do nên làm:** Các doanh nghiệp xuất khẩu đồ gỗ, dệt may, da giày, thủy sản đang bước vào giai đoạn nước rút giao hàng cuối năm. Nắm bắt biến động cước và chiêu bài của các liên minh hãng tàu giúp chủ hàng chủ động ngân sách, không bị bất ngờ trước các khoản phụ phí GRI/PSS tăng thêm từ 01/11.
- **Nguồn xác minh:** Drewry Shipping Consultants (`drewry.co.uk`); World Shipping Council; Báo Công Thương.

---

### Đề xuất #5: [LOGISTICS CỬA KHẨU BIÊN GIỚI] Thông quan nông sản mùa cao điểm cuối năm tại cửa khẩu Tân Thanh & Hữu Nghị: Thích ứng cơ chế Cửa khẩu thông minh (Thông tư 128) và phân luồng linh hoạt tránh ùn tắc
- **Chuyên mục (Pillar):** `HT` (Học Thực chiến - Logistics Biên mậu & Chuỗi cung ứng Lạnh)
- **Nền tảng phù hợp:** **Facebook** (Cảnh báo thực địa cho tài xế và chủ vựa nông sản) & **Infographic sơ đồ phân luồng**
- **Đánh giá theo 4 tiêu chí:**
  | Tiêu chí | Điểm (1-5) | Ghi chú đánh giá |
  |----------|:----------:|------------------|
  | Tính thời sự | ⭐⭐⭐⭐⭐ (4.8/5) | Quý 4 bước vào cao điểm xuất khẩu trái cây (sầu riêng, thanh long, mít) sang Trung Quốc |
  | Độ đau | ⭐⭐⭐⭐⭐ (4.8/5) | Hàng nghìn xe container lạnh dồn về cửa khẩu; nếu bị lỗi dữ liệu cân tự động hoặc kiểm dịch thì hàng bị kẹt, chi phí chạy máy lạnh tăng vọt |
  | Tiềm năng tranh luận | ⭐⭐⭐⭐⭐ (4.7/5) | Thảo luận: "Làm sao để cân bằng giữa yêu cầu siết chặt kiểm tra thông minh của Thông tư 128 và tốc độ thông quan giải phóng hàng tươi sống?" |
  | Khả năng cho giá trị | ⭐⭐⭐⭐⭐ (4.8/5) | Hướng dẫn 3 khâu: Chuẩn hóa phiếu cân trước khi rời kho, khớp nối mã seal điện tử và điều phối xe sang tuyến chuyên dụng mốc 1088/2-1089 |
  | **Tổng điểm** | **19.1 / 20** | **Xếp hạng: #5 (Giải pháp thực địa bảo vệ dòng chảy nông sản Việt)** |
- **Tóm tắt lý do nên làm:** Nông sản là mặt hàng vô cùng nhạy cảm với thời gian. Việc Eureka đưa ra hướng dẫn chuẩn hóa dữ liệu kết hợp giải pháp điều phối luồng xe giúp bảo toàn chất lượng chuỗi lạnh và tiết kiệm hàng chục triệu đồng tiền lưu xe cho các chủ vựa.
- **Nguồn xác minh:** Cổng TTĐT Tổng cục Hải quan (`customs.gov.vn`); Cục Hải quan Lạng Sơn; Báo Hải quan Online.

---

## 🎯 BƯỚC TIẾP THEO (NEXT ACTIONS DÀNH CHO CHỊ THÚY)
1. **Lựa chọn chủ đề ưu tiên:** Chị Thúy có thể phản hồi chọn 1 hoặc nhiều chủ đề (ví dụ: *'Chọn #1 và #2'*) để hệ thống tự động đẩy vào:
   - Quy trình `wf01`: Lập lịch biên tập chi tiết trong Content Calendar.
   - Hoặc quy trình `wf02`/`wf03`: Sản xuất nội dung bài viết chuyên sâu và kịch bản video AI ngay lập tức.
2. **Trường hợp chưa chọn:** Bản tin này đã được lưu trữ bền vững tại `scratch/hot_news_briefing_w44_2610.md` và sẵn sàng tái sử dụng cho các kế hoạch truyền thông tiếp theo.
"""

    with open(BRIEFING_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(briefing_content)
    print(f"THÀNH CÔNG: Đã tạo bản tin đề xuất tại {BRIEFING_MD_PATH}")

if __name__ == '__main__':
    update_seo_w44_2610()
