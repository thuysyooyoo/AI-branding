# -*- coding: utf-8 -*-
r"""
Tạo file Excel Quản lý Lịch phát sóng YouTube Series đa Sheet:
- Tách mỗi chuyên mục thành một Sheet riêng biệt (4 Sheet Chuyên mục: CM1, CM2, CM3, CM4).
- Mỗi chuyên mục có ĐỦ 9 CHỦ ĐỀ (1 hoặc nhiều tập), tổng cộng 36 chủ đề chi tiết.
- Sheet Tổng quan Lịch phát sóng & Sheet Checklist quay Batch.
- Lưu trữ tại: C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx
- Tuân thủ quy chuẩn: Atomic save (.tmp.xlsx -> os.replace), không sinh file backup, font Arial 10pt.
- Kết luận video 100% là chiêm nghiệm tự thân của Thúy cho chính mình.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TARGET_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
TARGET_FILE = os.path.join(TARGET_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
TMP_FILE = os.path.join(TARGET_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.tmp.xlsx")

def create_multisheet_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    # Colors
    NAVY_HEADER = "1B365D"
    GOLD_ACCENT = "D4AF37"
    LIGHT_BLUE = "E8EEF5"
    ZEBRA_FILL = "F9FAFC"
    SPECIAL_FILL = "FFFDF0"
    WHITE = "FFFFFF"
    BORDER_COLOR = "D0D7DE"

    font_title = Font(name="Arial", size=13, bold=True, color="1B365D")
    font_subtitle = Font(name="Arial", size=9.5, italic=True, color="555555")
    font_header = Font(name="Arial", size=10, bold=True, color=WHITE)
    font_bold = Font(name="Arial", size=10, bold=True, color="1B365D")
    font_regular = Font(name="Arial", size=10, color="000000")

    fill_header = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_light = PatternFill(start_color=LIGHT_BLUE, end_color=LIGHT_BLUE, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_FILL, end_color=ZEBRA_FILL, fill_type="solid")
    fill_special = PatternFill(start_color=SPECIAL_FILL, end_color=SPECIAL_FILL, fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_top_left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    align_top_center = Alignment(horizontal="center", vertical="top", wrap_text=True)

    headers_cm_sheet = [
        "STT", "Mã Chủ Đề", "Tên Chủ Đề & Tiêu Đề Video (SEO Title)", "Định Dạng & Số Tập",
        "Thời Lượng", "Nguồn Dữ Liệu Xác Thực (Fact-Check)", "Hook Mở Đầu (0:00 - 0:45)",
        "Cốt Lõi Kịch Bản & Luận Điểm Chi Tiết (Full Script / Key Points)",
        "Chỉ Dẫn Visual, Đạo Cụ & Góc Máy", "Móc Câu Cliffhanger (Nếu nhiều tập)",
        "Kết Luận & Chiêm Nghiệm Tự Thân Của Thúy", "Trạng Thái Xem Xét"
    ]

    col_widths_cm = [6, 12, 36, 18, 14, 32, 35, 65, 30, 32, 35, 16]

    # =========================================================================
    # SHEET 1: TỔNG QUAN LỊCH PHÁT SÓNG & QUY TRÌNH QUẢN TRỊ NỘI DUNG
    # =========================================================================
    ws_ov = wb.create_sheet(title="TONG_QUAN_LICH_PHAT_SONG")
    ws_ov.views.sheetView[0].showGridLines = True

    ws_ov.merge_cells("A1:K1")
    ws_ov["A1"] = "HỆ THỐNG QUẢN LÝ LỊCH PHÁT SÓNG YOUTUBE CHUẨN SERIES — KÊNH TÂM TÌNH THUYTALK"
    ws_ov["A1"].font = font_title
    ws_ov["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws_ov.merge_cells("A2:K2")
    ws_ov["A2"] = "Chiến lược: 1 tuần 3 bài (Thứ 2 - 4 - 6 lúc 20:00) | Dựng hoàn tất trước 1 tuần | Quay theo đợt nguyên kịch bản/series | 4 Chuyên mục phân chia thành 4 Sheet độc lập (36 chủ đề)"
    ws_ov["A2"].font = font_subtitle

    # Table 1: Tổng quan 4 Chuyên mục trụ cột
    ws_ov["A4"] = "I. BẢNG ĐIỀU HƯỚNG 4 CHUYÊN MỤC TRỤ CỘT (SHOW FORMATS) TRÊN CÁC SHEET"
    ws_ov["A4"].font = font_bold

    headers_ov1 = ["STT", "Tên Sheet", "Chuyên Mục", "Số Chủ Đề", "Định Dạng Tập", "Nguồn Dữ Liệu Xác Thực Chính", "Phong Cách Thể Hiện & Giá Trị Cốt Lõi"]
    for col_idx, h in enumerate(headers_ov1, 1):
        cell = ws_ov.cell(row=5, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    data_ov1 = [
        ("1", "CM1_VAN_BAN_NONG", "Trà Chiều & Văn Bản Nóng", "9 Chủ đề", "1 hoặc 2 Tập", "Văn bản pháp luật mới (Nghị định, Thông tư, Công văn) + Cào Facebook/Google các điểm nóng dư luận", "Bắt nhịp thời sự, bóc tách bản chất văn bản, cảnh báo rủi ro kịp thời; giọng điệu trấn an, điềm đạm, không đao to búa lớn."),
        ("2", "CM2_DICH_LUAT_TIENG_NGUOI", "Dịch Luật Ra Tiếng Người", "9 Chủ đề", "1 hoặc 2 Tập", "Luật Hải quan 2014, Luật Quản lý thuế 38, Incoterms 2020 ICC, Giáo trình Ngoại thương & Kế toán", "Bình dân hóa khái niệm khô khan (Hậu kiểm, Incoterms, Trốn thuế vs Tối ưu) bằng ví dụ đời sống sinh động, hóm hỉnh."),
        ("3", "CM3_LICH_SU_THUONG_MAI", "Ký Sự Lịch Sử Thương Mại", "9 Chủ đề", "1 Tập hoặc Series 3 Tập Liên Hoàn", "Wikipedia, website chính thống, sách kinh điển (The Box - Marc Levinson), Đại Việt Sử Ký Toàn Thư", "Kể chuyện lịch sử 100% tình tiết thật, tổ chức thành các chuỗi dài có Cliffhanger kích thích bấm xem tập tiếp theo."),
        ("4", "CM4_TAM_TINH_THUC_CHIEN", "Tâm Tình Thực Chiến Nghề XNK", "9 Chủ đề", "1 Tập (Kèm Clip Chào Sân)", "Daily Note thực tế công việc hàng ngày + Bài đăng trang cá nhân https://www.facebook.com/Th.xuatnhapkhau/", "Chia sẻ trải nghiệm va vấp thật, sự cố hiện trường; kết bài luôn là bài học THÚY TỰ RÚT RA CHO BẢN THÂN MÌNH.")
    ]

    for row_idx, row_data in enumerate(data_ov1, 6):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws_ov.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_center if col_idx in [1, 2, 4, 5] else align_left

    # Table 2: Nguyên tắc sản xuất & Quy chế phát sóng
    start_r2 = 12
    ws_ov.cell(row=start_r2, column=1, value="II. NGUYÊN TẮC SẢN XUẤT GỐI ĐẦU & BẢN QUYỀN NỘI DUNG").font = font_bold

    headers_ov2 = ["Nguyên tắc cốt lõi", "Quy định chi tiết", "Tiêu chuẩn kiểm duyệt", "Mục tiêu đạt được"]
    for col_idx, h in enumerate(headers_ov2, 1):
        cell = ws_ov.cell(row=start_r2+1, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    data_ov2 = [
        ("1. Sản xuất gối đầu trước 1 tuần", "Quay và dựng hoàn tất ở Tuần W để toàn bộ 3 video sẵn sàng lịch đăng tự động ở Tuần W+1.", "Video master xuất chuẩn 4K/1080p hoàn thành trước lịch chiếu tối thiểu 3 ngày.", "Chủ động 100% lịch phát sóng, không bao giờ rơi vào thế bị động hay trễ hạn."),
        ("2. Quay theo đợt (Batch Filming)", "Mỗi buổi quay (3-4 tiếng bên bàn trà), Chị Thúy quay trọn vẹn 1 Kịch bản nhiều tập (như Series Container 3 tập) hoặc 3 kịch bản đơn lẻ.", "Giữ nguyên mạch cảm xúc, phong thái ung dung, ánh sáng và setup bàn trà đồng nhất.", "Tiết kiệm tối đa thời gian makeup, setup máy quay và tối ưu hóa lịch trình của Giám đốc."),
        ("3. Kết luận thuộc về bản thân", "Phần kết của 100% video là lời tự dặn lòng, chiêm nghiệm của chính bản thân Thúy cho riêng mình.", "Tuyệt đối không dùng câu từ mang tính dạy đời, khuyên răn hay bảo người khác phải làm gì.", "Tạo dựng hình ảnh một người phụ nữ làm nghề tự tin, sâu sắc, khiêm tốn và thấu cảm."),
        ("4. Xác thực dữ liệu 100%", "Mọi số liệu công văn, nghị định, sự kiện lịch sử đều phải có nguồn dẫn chứng xác thực chính thống.", "Không bịa đặt, không phóng đại tình tiết lịch sử; trích dẫn rõ tên văn bản hoặc tên sách.", "Xây dựng uy tín học thuật và tính chuyên nghiệp cao nhất cho kênh YouTube ThuyTalk.")
    ]

    for row_idx, row_data in enumerate(data_ov2, start_r2+2):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws_ov.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_left

    col_widths_ov = [28, 45, 38, 40]
    for i, w in enumerate(col_widths_ov, 1):
        ws_ov.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # HELPER FUNCTION TO POPULATE CM SHEETS
    # =========================================================================
    def populate_cm_sheet(sheet_title, title_text, subtitle_text, data_rows):
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:L1")
        ws["A1"] = title_text
        ws["A1"].font = font_title
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")

        ws.merge_cells("A2:L2")
        ws["A2"] = subtitle_text
        ws["A2"].font = font_subtitle

        for c_idx, h in enumerate(headers_cm_sheet, 1):
            cell = ws.cell(row=4, column=c_idx)
            cell.value = h
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border

        for r_idx, r_data in enumerate(data_rows, 5):
            fill_curr = fill_zebra if r_idx % 2 == 0 else PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")
            if "Liên Hoàn" in str(r_data[3]) or "Chào Sân" in str(r_data[3]):
                fill_curr = fill_special

            for c_idx, val in enumerate(r_data, 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.value = val
                cell.font = font_regular
                cell.border = thin_border
                cell.fill = fill_curr

                if c_idx in [1, 2, 4, 5, 12]:
                    cell.alignment = align_top_center
                else:
                    cell.alignment = align_top_left

        for i, w in enumerate(col_widths_cm, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    # =========================================================================
    # SHEET 2: CM1_VAN_BAN_NONG (9 CHỦ ĐỀ)
    # =========================================================================
    data_cm1 = [
        (
            1, "CM1-CD01", "Bão rà soát đóng mã số thuế phụ 888 & Nợ thuế eTax Mobile: Khi sự vô tư phải trả giá bằng tiền triệu",
            "1 Tập", "7 - 8 phút",
            "Nghị định 125/2020/NĐ-CP (Điều 13 mức phạt chậm nộp hồ sơ khai thuế); Thông tư 40/2021/TT-BTC; Dữ liệu cào nhóm Hỏi đáp Thuế FB.",
            "Chào bạn. Nếu bạn đang làm affiliate trên Shopee, TikTok Shop, hoặc mở mã số thuế phụ đuôi 888 nhận hoa hồng... mà thời gian này mở eTax Mobile lên bỗng thấy báo nợ thuế, hoặc đi xin đóng mã số thuế mà nhận thông báo phạt chậm nộp hàng chục triệu... thì hôm nay Thúy ngồi lại chia sẻ góc nhìn thực tế nhé.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bóc tách ca thực tế: Cá nhân làm affiliate hoa hồng chỉ 1 triệu/năm, sàn đã khấu trừ 7% nộp kho bạc, nhưng khi đi đóng mã 888 thì bị cán bộ thuế in danh sách phạt 4 quý chậm nộp tờ khai trắng, dự kiến gần 24 triệu theo NĐ 125/2020.
- Giải thích pháp lý: Nghĩa vụ nộp hồ sơ khai thuế và nghĩa vụ nộp tiền thuế là 2 việc độc lập. Có MST phụ là bắt buộc phải nộp tờ khai quý dù doanh thu bằng 0.
- 3 bước gỡ rối: (1) Tra cứu các quý còn thiếu trên eTax Mobile; (2) Kê khai bổ sung mẫu 01/TKN-CNKD, ghi rõ số thuế đã khấu trừ; (3) Làm văn bản giải trình tình tiết giảm nhẹ theo Điều 9 Nghị định 125/2020.""",
            "Quay Chị Thúy bên bàn trà, chèn screencast ứng dụng eTax Mobile (che thông tin cá nhân). Bảng đồ họa các khung phạt Điều 13 Nghị định 125.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Sau trường hợp này, bài học mà bản thân Thúy tự rút ra cho chính mình là: 'Trong thời đại dữ liệu số liên thông, sự cẩn trọng với các thủ tục hành chính không bao giờ là thừa'. Bản thân mình luôn duy trì việc tự rà soát định kỳ các nghĩa vụ thuế cá nhân và doanh nghiệp, không bao giờ để một mã số thuế nào mở ra rồi bỏ quên." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            2, "CM1-CD02", "Công văn 22883 Cục Hải quan: Hamburger kẹp thịt là bánh mì hay thịt? Ranh giới 20% và bài học phân loại hàng hóa",
            "1 Tập", "6 - 8 phút",
            "Công văn số 22883/CHQ-NVTHQ ngày 12/10/2026 của Cục Hải quan; Chú giải chi tiết HS Chương 16 và Chương 19; Dữ liệu Daily Note 13/10/2026.",
            "Chào bạn. Có bao giờ bạn nghĩ một chiếc bánh Hamburger kẹp thịt thơm ngon chúng ta ăn mỗi ngày, khi đi qua cửa khẩu hải quan lại có thể khiến các doanh nghiệp đau đầu đến mức Cục Hải quan phải ra hẳn một công văn khẩn để chấn chỉnh không? Hamburger rốt cuộc là 'bánh mì' hay là 'thịt'? Hôm nay Thúy kể câu chuyện nghiệp vụ rất thú vị này nhé.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bối cảnh: Cục Hải quan ra Công văn 22883 ngày 12/10/2026 yêu cầu các Chi cục rà soát khẩn mặt hàng khai báo là Hamburger.
- Thực trạng: Nhiều doanh nghiệp quen khai vào nhóm 19.05 (bánh mì nướng) để hưởng thuế thấp và kiểm tra đơn giản.
- Quy tắc 20% định đoạt:
  + Nếu tổng thịt, xúc xích, phụ phẩm thịt vượt TRÊN 20% khối lượng bánh: Bị kéo sang Chương 16 (nhóm 16.02 chế phẩm thịt). Thuế suất thay đổi và phải kiểm dịch động vật ngặt nghèo.
  + Chỉ khi thịt KHÔNG VƯỢT QUÁ 20%: Mới được giữ ở nhóm 19.05 bánh mì.""",
            "Quay Chị Thúy cầm tách trà, bên cạnh có đĩa bánh hamburger mini mô phỏng. Slide đồ họa hiện infographic so sánh: Tỷ lệ thịt >20% (Chương 16) vs <=20% (Chương 19). Trích dẫn văn bản CV 22883 trên màn hình.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ câu chuyện chiếc bánh hamburger này, bài học mà bản thân Thúy tự nhắc nhở mình mỗi ngày khi làm nghề là: 'Trong ngoại thương, ranh giới giữa an toàn và rủi ro đôi khi chỉ là 1% định lượng kỹ thuật'. Bản thân mình luôn tự dặn lòng phải yêu cầu đối tác ngoại cấp Bản phân tích thành phần (C/A) chi tiết đến từng gram trước khi đóng hàng lên tàu." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            3, "CM1-CD03", "Sàn TMĐT khấu trừ thuế thay & Định danh tài khoản ngân hàng: Hồi kết của thời bán hàng online 'vô hình'",
            "1 Tập", "7 - 8 phút",
            "Nghị định quy định trách nhiệm thuế của sàn TMĐT; Thông tư khấu trừ thuế tại nguồn; Dữ liệu đối soát tài khoản Đề án 06.",
            "Chào bạn. Có bao giờ bạn tự hỏi: Tại sao những năm trước bán hàng online trăm đơn ngàn đơn một ngày chẳng thấy ai hỏi han, nhưng bước sang năm 2026, cơ quan thuế lại nắm rõ từng giao dịch, từng đồng tiền về tài khoản của bạn như trong lòng bàn tay? Thời của việc kinh doanh online 'vô hình' đã chính thức khép lại.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Phân tích cơ chế sàn Shopee/TikTok Shop tự động khấu trừ thuế VAT và TNCN trực tiếp khi hoàn tất đơn hàng.
- Hệ thống dữ liệu lớn (Big Data) đối soát dòng tiền tài khoản ngân hàng theo mã CCCD định danh.
- 2 việc chủ shop cần làm ngay: (1) Tính thuế 1.5% - 7% vào cơ cấu giá vốn từ đầu; (2) Lưu trữ hóa đơn chứng từ đầu vào chuẩn chỉ để sẵn sàng chuyển đổi mô hình doanh nghiệp.""",
            "Quay Chị Thúy bên bàn trà, màn hình chèn sơ đồ luồng tiền: Khách hàng -> Sàn TMĐT -> Khấu trừ thuế trực tiếp vào Kho bạc -> Trả tiền ròng cho Người bán. Đồ họa liên thông Đề án 06.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ sự chuyển dịch này của nền kinh tế số, bài học mà bản thân Thúy luôn tự dặn mình trong quản trị là: 'Đừng bao giờ xây dựng mô hình kinh doanh dựa trên việc trốn tránh một nghĩa vụ tất yếu'. Minh bạch ngay từ 1 đồng doanh thu đầu tiên là cách duy nhất để cơ nghiệp của mình lớn mạnh mà không nơm nớp lo sợ." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            4, "CM1-CD04", "Quyết định 31/2026/QĐ-TTg: 12 nhóm hàng bắt buộc làm thủ tục tại cửa khẩu nhập - Dính 1 món kẹt cả cont",
            "1 Tập", "7 - 8 phút",
            "Quyết định số 31/2026/QĐ-TTg của Thủ tướng Chính phủ quy định Danh mục hàng hóa nhập khẩu phải làm thủ tục hải quan tại cửa khẩu nhập.",
            "Chào bạn. Bạn nhập một container hàng bách hóa tổng hợp gồm 20 mặt hàng đồ chơi, đồ gia dụng rất bình thường. Nhưng khi cont vừa về đến cảng biển, bạn mới ngã ngửa vì không được chuyển cửa khẩu về ICD gần kho mà bị giữ lại kiểm tra toàn bộ tại cảng biển đích thân. Lý do chỉ vì trong cont dính đúng một món nằm trong Quyết định 31!",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Phân tích 12 nhóm hàng đặc thù theo QĐ 31 (rượu bia, thuốc lá, điều hòa, ô tô xe máy, phế liệu...).
- Cơ chế 'Một con sâu làm rầu nồi canh': Hàng hỗn hợp chỉ cần dính 1 mã hàng thuộc QĐ 31 là toàn bộ cont bị chặn không cho chuyển cửa khẩu.
- Giải pháp đàm phán Incoterms & kỹ thuật tách Bill (B/L) từ xưởng gốc nước ngoài trước khi book cont.""",
            "Quay Chị Thúy cầm danh mục QĐ 31 trên tay. Màn hình minh họa chiếc container hỗn hợp bị dừng tại cảng biển Hải Phòng/Cát Lái.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ những ca thực tế gỡ cont hàng dính Quyết định 31, Thúy tự rút ra bài học cho mình: 'Khâu kiểm duyệt danh mục hàng hóa trước khi đóng cont quan trọng gấp trăm lần việc xử lý sự cố khi hàng đã cập cảng'. Bản thân mình luôn rà soát từng mã HS trước khi cho phép đóng hàng chung cont." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            5, "CM1-CD05", "Tạm hoãn xuất cảnh vì nợ thuế: Khi những khoản nợ thuế môn bài, thuế đất từ 10 năm trước bất ngờ 'sống dậy'",
            "1 Tập", "7 - 8 phút",
            "Điều 66 Luật Quản lý thuế số 38/2019/QH14; Nghị định 126/2020/NĐ-CP; Thông báo tạm hoãn xuất cảnh của Cục Thuế các địa phương.",
            "Chào bạn. Có những người xách vali ra đến sân bay quốc tế, chuẩn bị bước lên chuyến bay đi công tác hoặc du lịch cùng gia đình thì bị chặn lại tại cửa an ninh xuất nhập cảnh. Lý do: Đang bị tạm hoãn xuất cảnh vì một khoản nợ thuế vài triệu đồng từ một công ty đã bỏ hoang từ 10 năm trước. Chuyện gì đang diễn ra vậy?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Cơ chế liên thông giữa Tổng cục Thuế và Cục Quản lý Xuất nhập cảnh Bộ Công an theo Điều 66 Luật Quản lý thuế 38.
- Ai là người bị áp dụng? Người đại diện theo pháp luật của doanh nghiệp đang bị cưỡng chế thuế; cá nhân kinh doanh chưa hoàn thành nghĩa vụ thuế.
- 3 bước tự tra soát tình trạng nợ thuế cá nhân trước khi mua vé máy bay xuất cảnh.""",
            "Quay Chị Thúy bên tách trà, màn hình chèn ảnh minh họa cửa kiểm soát an ninh sân bay và thông báo hoãn xuất cảnh mẫu.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học mà bản thân Thúy tự nhắc nhở mình: 'Pháp lý doanh nghiệp luôn gắn liền mật thiết với sự tự do cá nhân của người đại diện'. Bản thân mình không bao giờ đứng tên hộ hay để lại bất kỳ doanh nghiệp nào ở trạng thái 'ngừng hoạt động nhưng chưa đóng mã số thuế' để bảo vệ sự thảnh thơi của chính mình." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            6, "CM1-CD06", "Công văn 6810/CT về chênh lệch tỷ giá hạch toán hải quan: Bài học phòng ngừa rủi ro cho doanh nghiệp nhập khẩu",
            "1 Tập", "6 - 8 phút",
            "Công văn số 6810/CT về tỷ giá tính thuế và hạch toán chênh lệch tỷ giá ngoại tệ; Thông tư 200/2014/TT-BTC chuẩn mực kế toán.",
            "Chào bạn. Khi nhập khẩu một lô hàng giá trị 100.000 USD, tỷ giá hải quan áp trên tờ khai là 25.000, nhưng khi bạn chuyển tiền thanh toán cho đối tác qua ngân hàng thì tỷ giá thực tế lại là 25.400. Khoản chênh lệch 40 triệu đồng đó xử lý thế nào để không bị gạt khỏi chi phí hợp lý khi quyết toán thuế?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Giải thích sự khác biệt giữa Tỷ giá tính thuế của Hải quan (áp dụng cố định theo tuần của Vietcombank) và Tỷ giá giao dịch thực tế tại thời điểm thanh toán.
- Cách hạch toán tài khoản 413 (chênh lệch tỷ giá) và tài khoản 635 (chi phí tài chính) chuẩn chỉ theo Thông tư 200.
- Các công cụ phái sinh tiền tệ (Forward, Swap) giúp doanh nghiệp cố định tỷ giá khi biến động tiền tệ thế giới tăng cao.""",
            "Quay Chị Thúy với cuốn sổ tay ghi chép số liệu tài chính. Slide biểu đồ biến động tỷ giá USD/VND và bảng hạch toán kế toán chuẩn.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ những lần đau đầu vì biến động tỷ giá của các lô hàng lớn, bài học Thúy tự rút ra là: 'Quản trị rủi ro tài chính phải đi trước khâu vận hành logistics một bước'. Bản thân mình luôn chủ động làm việc với ngân hàng để khóa trước tỷ giá cho các đơn hàng lớn để bảo toàn lợi nhuận." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            7, "CM1-CD07", "Siết kiểm tra an toàn thực phẩm & kiểm tra bao bì (CV 22082): Khi chiếc hộp đựng cũng bị soi kiểm nghiệm",
            "1 Tập", "7 - 8 phút",
            "Nghị định 15/2018/NĐ-CP quy định chi tiết Luật An toàn thực phẩm; Công văn số 22082 về rà soát kiểm tra chuyên ngành bao bì tiếp xúc trực tiếp thực phẩm.",
            "Chào bạn. Nhiều doanh nghiệp nhập khẩu bánh kẹo, trà, cà phê cứ nghĩ chỉ cần kiểm nghiệm chất lượng sản phẩm bên trong là xong. Nhưng khi hàng về cảng, hải quan yêu cầu kiểm tra chuyên ngành cả chiếc hộp sắt, túi nilon hay màng nhôm bọc bên ngoài! Tại sao lại như vậy?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Khái niệm: 'Bao bì, dụng cụ tiếp xúc trực tiếp với thực phẩm' theo Nghị định 15/2018.
- Nguy cơ thôi nhiễm kim loại nặng hoặc vi nhựa từ bao bì vào thực phẩm khiến lô hàng bị đình chỉ thông quan.
- Quy trình tự công bố sản phẩm và chuẩn bị hồ sơ kiểm tra ATTP cho cả sản phẩm lẫn bao bì chứa đựng.""",
            "Quay Chị Thúy cầm một hộp trà kim loại và túi zipper thực phẩm. Infographic minh họa quy trình kiểm tra chuyên ngành theo Nghị định 15.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học mà bản thân Thúy tự nhắc nhở mình: 'Đừng bao giờ bỏ qua những chi tiết phụ kiện đi kèm trong một sản phẩm xuất nhập khẩu'. Sự chuẩn bị hồ sơ đồng bộ cả vỏ lẫn ruột là cách tốt nhất giúp hàng hóa của mình thông quan nhanh chóng." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            8, "CM1-CD08", "Siết hóa đơn điện tử từng lần bán & Rủi ro hóa đơn công ty 'ma' bỏ trốn: Cách bảo vệ chi phí được trừ",
            "1 Tập", "7 - 8 phút",
            "Nghị định 123/2020/NĐ-CP về hóa đơn, chứng từ; Danh sách các doanh nghiệp rủi ro cao về thuế do Tổng cục Thuế công bố.",
            "Chào bạn. Có bao giờ bạn rơi vào tình cảnh: 3 năm trước bạn mua hàng hóa thật, chuyển khoản ngân hàng thật, có hóa đơn đỏ đầy đủ. Đùng một cái, cơ quan thuế gửi thông báo: Công ty bán hàng cho bạn đã bỏ trốn khỏi địa chỉ kinh doanh và bị liệt vào danh sách doanh nghiệp ma. Toàn bộ chi phí đầu vào và thuế VAT bị yêu cầu loại bỏ!",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Thực trạng mua bán hóa đơn bất hợp pháp và cơ chế tự động quét chuỗi hóa đơn của cơ quan thuế qua hệ thống hóa đơn điện tử e-Invoice.
- 'Bộ chứng từ bảo vệ 4 lớp': Hợp đồng kinh tế + Biên bản giao nhận hàng hóa + Chứng từ vận tải (phiếu cân, định vị xe) + Chứng từ thanh toán qua ngân hàng.
- Khi có công văn giải trình, doanh nghiệp chứng minh giao dịch có thật như thế nào để được chấp nhận chi phí hợp lý.""",
            "Quay Chị Thúy bên bàn trà, lật mở bộ hồ sơ chứng từ 4 lớp mẫu. Slide đồ họa minh họa chuỗi quét hóa đơn của cơ quan thuế.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm của bản thân Thúy: 'Hóa đơn chỉ là một tờ giấy, bằng chứng vật chất về việc hàng hóa thực sự di chuyển mới là thứ bảo vệ mình trước pháp luật'. Bản thân mình luôn yêu cầu nhân sự lưu giữ trọn bộ chứng từ giao nhận thực tế đi kèm mỗi hóa đơn đầu vào." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            9, "CM1-CD09", "Làn sóng điều tra phòng vệ thương mại & Quy định siết gian lận xuất xứ (C/O) đi Mỹ, EU",
            "2 Tập Liên Hoàn", "16 - 18 phút (2 Tập)",
            "Luật Quản lý ngoại thương 2017 (Chương phòng vệ thương mại); Thông tư 31/2018/TT-BCT về xuất xứ hàng hóa; Báo cáo của Cục Phòng vệ Thương mại.",
            "Chào bạn. Hàng hóa Việt Nam xuất khẩu sang Mỹ, EU đang đối mặt với hàng chục vụ điều tra chống lẩn tránh thuế và gian lận xuất xứ. Chỉ vì một vài doanh nghiệp 'mượn đường' dán nhãn Made in Vietnam ăn xổi, mà cả một ngành hàng gỗ, nhôm, thép của nước ta đứng trước nguy cơ bị áp thuế trừng phạt hàng trăm phần trăm!",
            r"""[NỘI DUNG 2 TẬP LIÊN HOÀN - CHỊ THÚY]
=== TẬP 1: MƯỢN ĐƯỜNG ĐỘI LỐT & NGUY CƠ BỊ VẠ LÂY ===
Mổ xẻ các vụ án thực tế cont hàng ngoại nhập về Việt Nam chỉ qua công đoạn lắp ráp giản đơn (chưa đạt tiêu chuẩn chuyển đổi mã HS - CTC) rồi xin cấp C/O xuất đi Mỹ để né thuế chống bán phá giá. Hậu quả khi Bộ Thương mại Mỹ (DOC) mở cuộc điều tra toàn diện cả ngành hàng.
=== TẬP 2: CHIẾN LƯỢC BẢO VỆ CHUỖI CUNG ỨNG & CHỨNG CHỈ C/O CHUẨN CHỈ ===
Cách doanh nghiệp chứng minh tỷ lệ giá trị gia tăng nội địa (RVC), tiêu chí xuất xứ thuần túy (WO), và lưu trữ hồ sơ chuỗi cung ứng nguyên phụ liệu để vượt qua các đợt thẩm tra tại chỗ (on-site verification) của đoàn thanh tra quốc tế.""",
            "Quay Chị Thúy với bản đồ chuỗi cung ứng toàn cầu và chứng nhận C/O mẫu. Trích dẫn số liệu các vụ kiện phòng vệ thương mại của Bộ Công Thương.",
            "Cuối Tập 1 có Cliffhanger: 'Khi DOC ập vào kiểm tra đột xuất tại xưởng, điều gì đã cứu sống doanh nghiệp? Mời bạn bấm sang Tập 2 để theo dõi tiếp nhé!'",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học mà bản thân Thúy luôn khắc ghi: 'Trong thương mại quốc tế, sự trung thực về xuất xứ là danh dự của một doanh nghiệp và của cả một quốc gia'. Bản thân mình không bao giờ chấp nhận các đơn hàng có yếu tố 'đội lốt', bởi vì con đường chân chính mới là con đường duy nhất để đi xa." """,
            "Chờ Chị Thúy duyệt"
        )
    ]

    # =========================================================================
    # SHEET 3: CM2_DICH_LUAT_TIENG_NGUOI (9 CHỦ ĐỀ)
    # =========================================================================
    data_cm2 = [
        (
            1, "CM2-CD01", "Kiểm tra sau thông quan: Giải mã 5 năm 'hồi tố' theo Luật Hải quan 2014 & Cách kiểm soát định mức kho",
            "1 Tập", "8 - 10 phút",
            "Điều 77 đến 82 Luật Hải quan 2014; Thông tư 38/2015/TT-BTC & TT 39/2018; Giáo trình Thủ tục Hải quan.",
            "Chào bạn. Trong ngành xuất nhập khẩu, người ta hay nói đùa một câu: 'Hàng ra khỏi cảng chưa chắc đã là của mình'. Nhiều anh chị thấy tờ khai luồng Xanh, hàng đã về kho đem bán hết lấy tiền về rồi thì đinh ninh là xong xuôi. Nhưng 3 năm, thậm chí 5 năm sau, bỗng nhận quyết định Kiểm tra sau thông quan... Tại sao lại có cơ chế 5 năm này?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bản chất chuyển từ 'Tiền kiểm' sang 'Hậu kiểm' của hải quan hiện đại: Cho thông quan nhanh, kiểm tra hồ sơ và tồn kho trong vòng 5 năm.
- 2 hình thức: Kiểm tra tại trụ sở cơ quan hải quan vs Kiểm tra tại trụ sở doanh nghiệp.
- 3 tử huyệt soi kỹ nhất: Mã HS, Trị giá tính thuế, và Báo cáo quyết toán định mức nguyên phụ liệu gia công/SXXK.""",
            "Quay Chị Thúy ngồi trước giá sách tài liệu. Đồ họa màn hình: So sánh Tiền kiểm vs Hậu kiểm. Highlight mốc thời hạn 5 năm Điều 78 Luật Hải quan.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Với riêng bản thân Thúy: Thông quan tại cảng chưa phải là kết thúc; bản thân mình luôn rèn luyện thói quen lưu trữ hồ sơ ngăn nắp, đồng bộ kho với kế toán từng tháng để lúc nào cũng sẵn sàng và tự tin tiếp đón đoàn kiểm tra." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            2, "CM2-CD02", "Incoterms 2020: Vì sao 'Mua CIF Bán FOB' khiến doanh nghiệp Việt mất quyền tự quyết chi phí logistics?",
            "1 Tập", "7 - 8 phút",
            "Quy tắc Incoterms 2020 ICC; Báo cáo Logistics Việt Nam thường niên của Bộ Công Thương; Phân tích cước vận tải biển.",
            "Chào bạn. Suốt mấy chục năm qua, có một thói quen ăn sâu vào tiềm thức của phần lớn doanh nghiệp XNK Việt Nam: Xuất khẩu thì bán FOB, còn Nhập khẩu thì mua CIF. Người ta nghĩ thế là an nhàn, không phải đau đầu thuê tàu. Nhưng thực sự chúng ta đang được lợi hay đang tự trói tay mình?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bản chất Incoterms: Phân định rành mạch Chi phí, Rủi ro và Trách nhiệm chứng từ.
- Bẫy phụ phí local charges khi mua CIF: Hãng tàu ngoại hạ giá cước biển nhưng nâng phí cảng đến cắt cổ.
- Nhường miếng bánh logistics khi bán FOB: Mất quyền đàm phán cước khối lượng lớn. Lợi ích khi tự tin chuyển sang Mua FOB, Bán CIF.""",
            "Quay Chị Thúy bên tách trà, 2 mô hình tàu container mini gắn nhãn FOB và CIF. Infographic so sánh dòng dịch chuyển chi phí và rủi ro.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ trải nghiệm bị hãng tàu ngoại ép phụ phí địa phương trong những năm đầu khởi nghiệp, bài học Thúy tự rút ra là: 'Chủ động nắm quyền book tàu chính là chìa khóa để mình làm chủ chi phí và tiến độ lô hàng'. Với Eureka, Thúy luôn ưu tiên giá FOB để tự chỉ định đối tác vận tải tin cậy." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            3, "CM2-CD03", "Ranh giới giữa 'Trốn thuế' (Evasion) và 'Tối ưu chi phí thuế' (Planning) trước pháp luật",
            "1 Tập", "8 - 9 phút",
            "Điều 143 Luật Quản lý thuế số 38/2019; Điều 200 Bộ luật Hình sự 2015; Chuẩn mực Kế toán Việt Nam (VAS).",
            "Chào bạn. Có một ranh giới trong kinh doanh mỏng manh như sợi tóc, nhưng bước qua một bên là sự thông minh tài chính, còn trượt sang bên kia lại là vòng lao lý. Đó chính là ranh giới giữa 'Tối ưu thuế' và 'Trốn thuế'. Làm sao để người làm kinh doanh vừa tiết kiệm chi phí hợp pháp, vừa đêm về ngủ ngon giấc?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Phân biệt 3 cấp độ: Trốn thuế (cố ý phạm pháp, bị xử lý hình sự) vs Tránh thuế (lợi dụng kẽ hở, bị rà soát chống chuyển giá) vs Tối ưu thuế (hoàn toàn hợp pháp và được khuyến khích).
- Các phương pháp tối ưu thuế chuẩn: Tận dụng ưu đãi FTA, ưu đãi địa bàn khu công nghiệp, trích lập quỹ phát triển KHCN, hạch toán đúng chi phí được trừ.
- Ví von bình dân: Trốn thuế là vượt đèn đỏ; Tối ưu thuế là đi đúng làn đường cao tốc cho phép.""",
            "Quay Chị Thúy ánh mắt điềm tĩnh, nụ cười nhẹ nhàng. Màn hình chia 3 cột đồ họa: Đỏ (Trốn thuế - Cấm) | Vàng (Tránh thuế - Rủi ro) | Xanh (Tối ưu thuế - Hợp pháp).",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm lớn nhất mà bản thân Thúy đúc kết cho chính mình qua bao năm làm doanh nghiệp: 'Đồng tiền giữ lại được từ sự hiểu biết và tuân thủ pháp luật là đồng tiền bình an nhất'. Bản thân mình không bao giờ tìm cách đi đường tắt, bởi vì giấc ngủ ngon là tài sản vô giá." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            4, "CM2-CD04", "Mã HS Code: 8 con số định đoạt số phận lô hàng & 6 Quy tắc giải thích phân loại quốc tế (GIR) dễ hiểu nhất",
            "1 Tập", "8 - 10 phút",
            "Công ước Quốc tế về Hệ thống Hài hòa mô tả và mã hóa hàng hóa (Hệ thống HS của WCO); 6 Quy tắc Tổng quát GIR; Biểu thuế XNK Việt Nam.",
            "Chào bạn. 8 con số trong mã HS Code trông như một dãy số ngẫu nhiên, nhưng trong ngoại thương, nó quyết định bạn phải nộp thuế 0% hay 30%, lô hàng của bạn thông quan trong 5 phút hay bị giữ lại lấy mẫu phân tích 2 tuần. Hôm nay Thúy giải mã 6 Quy tắc phân loại HS bằng ngôn ngữ đời thường nhé.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Cấu trúc 8 chữ số: 2 số đầu (Chương), 4 số (Nhóm), 6 số (Phân nhóm quốc tế), 8 số (Phân nhóm quốc gia).
- 6 Quy tắc GIR diễn giải bình dân:
  + Quy tắc 1: Tên nhóm và chú giải quyết định tất cả.
  + Quy tắc 2: Hàng chưa hoàn chỉnh hoặc dạng tháo rời.
  + Quy tắc 3: Hàng hỗn hợp, hàng bộ (chọn thành phần mang đặc trưng cơ bản nhất).
  + Quy tắc 5: Bao bì và hộp đựng chuyên dụng đi kèm sản phẩm.
- Thủ tục Xác định trước mã số (Advance Ruling) bảo hiểm pháp lý cho doanh nghiệp.""",
            "Quay Chị Thúy bên cuốn Biểu thuế XNK dày cộp. Infographic trực quan giải phẫu dãy số 8 chữ số mã HS Code.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học mà bản thân Thúy tự dặn mình: 'Mã HS không phải là việc phán đoán theo cảm tính, mà là một môn khoa học pháp lý đòi hỏi tính chính xác tuyệt đối'. Thúy luôn tự mình kiểm tra tài liệu kỹ thuật của nhà sản xuất trước khi gõ phím phân loại mã." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            5, "CM2-CD05", "Trị giá hải quan (Customs Valuation): Cơ chế 'Tham vấn giá' là gì mà hải quan hay nghi vấn giá khai báo?",
            "1 Tập", "7 - 8 phút",
            "Hiệp định Trị giá Hải quan WTO (GATT 1994); Thông tư 39/2015/TT-BTC & Thông tư 60/2019/TT-BTC về xác định trị giá hải quan.",
            "Chào bạn. Bạn mua được lô hàng giá hời từ nhà cung cấp với giá 10 USD/chiếc. Nhưng khi khai hải quan, công chức hải quan gửi thông báo nghi vấn và yêu cầu đi 'Tham vấn giá', cho rằng giá thực tế trên thị trường phải là 18 USD. Cơ chế tham vấn giá hoạt động thế nào và làm sao để bảo vệ mức giá mua thật của mình?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bản chất của Trị giá giao dịch (Transaction Value) - Phương pháp 1 trong 6 phương pháp xác định trị giá hải quan.
- Cơ sở dữ liệu giá của ngành Hải quan (Danh mục quản lý rủi ro về giá).
- Hồ sơ chứng minh giá mua thật: Hợp đồng ngoại thương, Swift thanh toán quốc tế qua ngân hàng, bảng chào giá của nhà sản xuất, chứng từ cước vận chuyển và sổ sách kế toán nhập kho.""",
            "Quay Chị Thúy bên bàn trà, màn hình hiển thị biên bản làm việc tham vấn giá mẫu và bộ hồ sơ chứng minh thanh toán quốc tế.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ những lần đi tham vấn giá cùng khách hàng, bài học Thúy tự rút ra là: 'Sự thật và dòng tiền thanh toán minh bạch qua ngân hàng là vũ khí phản biện sắc bén nhất'. Bản thân mình không bao giờ chấp nhận khai báo 2 giá, vì sự trung thực luôn giúp mình giữ được quyền lợi chính đáng." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            6, "CM2-CD06", "Phân biệt Luồng Xanh, Luồng Vàng, Luồng Đỏ: Hệ thống quản lý rủi ro (RMS) hải quan chấm điểm doanh nghiệp ra sao?",
            "1 Tập", "6 - 8 phút",
            "Luật Hải quan 2014 (Điều 16, 17 về Quản lý rủi ro); Thông tư 81/2019/TT-BTC quy định quản lý rủi ro trong hoạt động nghiệp vụ hải quan.",
            "Chào bạn. Tại sao cùng một mặt hàng, doanh nghiệp A bấm tờ khai một giây sau là ra Luồng Xanh thông quan ngay, còn doanh nghiệp B lại liên tục bị nhảy Luồng Đỏ phải kéo cont ra bãi cắt chì kiểm hóa từng kiện hàng? Bí mật nằm ở Hệ thống Quản lý rủi ro (RMS) tự động chấm điểm tuân thủ của hải quan.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Ý nghĩa 3 luồng tờ khai: Xanh (miễn kiểm tra), Vàng (kiểm tra hồ sơ chứng từ), Đỏ (kiểm tra thực tế hàng hóa).
- 5 mức độ tuân thủ pháp luật hải quan của doanh nghiệp (từ Mức 1 - Doanh nghiệp ưu tiên đến Mức 5 - Không tuân thủ).
- Những 'điểm đen' khiến doanh nghiệp bị hạ bậc: Bị phạt vi phạm hành chính, khai sai mã nhiều lần, nợ thuế quá hạn, doanh nghiệp mới thành lập dưới 1 năm.""",
            "Quay Chị Thúy với giao diện phần mềm khai báo hải quan điện tử VNACCS/VCIS. Đồ họa mô phỏng hệ thống chấm điểm rủi ro AI của Hải quan.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bản thân Thúy luôn tự nhắc nhở đội ngũ Eureka: 'Điểm số tuân thủ của doanh nghiệp giống như điểm tín dụng cá nhân, xây dựng thì mất cả năm nhưng đánh mất chỉ sau một lần làm ẩu'. Giữ cho tờ khai luôn xanh là thành quả của sự kỷ luật mỗi ngày." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            7, "CM2-CD07", "Hàng gia công vs Sản xuất xuất khẩu (A11, A12 vs E21, E31): Khác biệt sống còn về thuế và báo cáo quyết toán mẫu 15",
            "1 Tập", "8 - 10 phút",
            "Thông tư 39/2018/TT-BTC sửa đổi Điều 60 Thông tư 38/2015; Luật Thuế xuất khẩu, thuế nhập khẩu số 107/2016/QH13.",
            "Chào bạn. Cùng là nhập khẩu vải về may áo rồi xuất khẩu ra nước ngoài, nhưng nếu bạn chọn loại hình Gia công (E21) thì quyền sở hữu nguyên liệu thuộc về bên ngoại, còn nếu chọn Sản xuất xuất khẩu (E31) thì bạn tự mua nguyên liệu và tự bán thành phẩm. Sự khác nhau này ảnh hưởng sống còn đến nghĩa vụ thuế và báo cáo quyết toán ra sao?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- So sánh bản chất pháp lý: Gia công (bán sức lao động, nhận tiền gia công) vs Sản xuất xuất khẩu (mua đứt bán đoạn).
- Chính sách thuế nhập khẩu: Miễn thuế nguyên phụ liệu nhập khẩu nếu thực xuất trong thời hạn quy định.
- Nỗi ám ảnh Báo cáo quyết toán Mẫu 15: Công thức cân đối Nhập - Xuất - Tồn kho. Rủi ro bị truy thu thuế khi tỷ lệ tiêu hao thực tế lệch với định mức đăng ký.""",
            "Quay Chị Thúy bên bàn làm việc với bảng báo cáo quyết toán Mẫu 15/BCQT-NVL/GSQL. Đồ họa luồng dịch chuyển nguyên phụ liệu.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ những đợt quyết toán kho căng thẳng cùng các nhà máy, Thúy tự rút ra: 'Quản trị nguyên phụ liệu gia công và SXXK không thể quản lý bằng sổ tay mà bắt buộc phải quản trị bằng hệ thống ERP chuẩn'. Tính chính xác đến từng mét vải, từng cuộn chỉ là thứ giữ an toàn cho nhà xưởng." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            8, "CM2-CD08", "Vận đơn đường biển (B/L) - Master Bill vs House Bill: Ai thực sự là chủ sở hữu lô hàng khi có tranh chấp?",
            "1 Tập", "7 - 8 phút",
            "Bộ luật Hàng hải Việt Nam 2015; Quy tắc Hague-Visby về vận đơn đường biển quốc tế; Giáo trình Vận tải quốc tế.",
            "Chào bạn. Cầm trên tay tờ vận đơn đường biển (Bill of Lading), bạn có chắc chắn 100% mình là người duy nhất có quyền lấy hàng tại cảng không? Nếu đơn vị forwarder phát hành cho bạn một tờ House Bill, trong khi Hãng tàu chỉ biết người cầm Master Bill... chuyện gì sẽ xảy ra nếu một trong hai bên phá sản hoặc tranh chấp cước?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bản chất 3 chức năng của B/L: Biên nhận hàng hóa, Hợp đồng vận chuyển, và Chứng từ sở hữu hàng hóa (Document of Title).
- Phân biệt Master B/L (do Hãng tàu phát hành) vs House B/L (do Công ty Forwarder phát hành).
- Các loại vận đơn rủi ro: Vận đơn đích danh (Straight B/L) vs Vận đơn theo lệnh (To Order B/L) vs Surrendered B/L (Điện giao hàng telex release). Cách ký hậu chuyển nhượng an toàn.""",
            "Quay Chị Thúy cầm trên tay bản gốc vận đơn Master Bill của Hãng tàu Maersk/COSCO và bản House Bill đối chiếu. Highlight các ô quan trọng (Shipper, Consignee, Notify Party).",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học mà bản thân Thúy luôn dặn mình khi kiểm tra chứng từ: 'Vận đơn B/L chính là chiếc chìa khóa két sắt của lô hàng'. Bản thân mình không bao giờ cho phép xuất điện giao hàng khi dòng tiền thanh toán chưa về trọn vẹn tài khoản công ty." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            9, "CM2-CD09", "Thuế chống bán phá giá (Anti-dumping) & Thuế tự vệ: Bản chất kinh tế và cách doanh nghiệp kháng kiện",
            "2 Tập Liên Hoàn", "16 - 18 phút (2 Tập)",
            "Hiệp định Chống bán phá giá WTO (ADA); Luật Quản lý ngoại thương 2017; Các phán quyết của Cục Phòng vệ thương mại Việt Nam.",
            "Chào bạn. Tại sao một doanh nghiệp bán hàng giá rẻ cho người tiêu dùng lại bị kiện, bị phạt thuế chống bán phá giá lên đến 40% - 50%? Giá rẻ chẳng phải là điều tốt sao? Hôm nay Thúy giải mã bản chất kinh tế và pháp lý của công cụ phòng vệ thương mại quyền lực này nhé.",
            r"""[NỘI DUNG 2 TẬP LIÊN HOÀN - CHỊ THÚY]
=== TẬP 1: BẢN CHẤT CỦA 'BÁN PHÁ GIÁ' LÀ GÌ? ===
Giải thích công thức: Giá xuất khẩu THẤP HƠN Giá bán tại thị trường nội địa của nước xuất khẩu. Mục đích thôn tính thị trường, tiêu diệt nền sản xuất nước sở tại rồi độc quyền nâng giá. Cơ chế tính biên độ phá giá và điều kiện áp thuế trừng phạt.
=== TẬP 2: CHIẾN LƯỢC THAM GIA KHÁNG KIỆN PHÒNG VỆ THƯƠNG MẠI ===
Cách doanh nghiệp xuất khẩu chuẩn bị hệ thống sổ sách kế toán chi phí sản xuất để trả lời Bản câu hỏi điều tra (Questionnaire) của cơ quan điều tra quốc tế, giành mức thuế suất riêng lẻ 0% thay vì chịu mức thuế suất toàn quốc.""",
            "Quay Chị Thúy bên tách trà, màn hình hiển thị biểu đồ so sánh giá bán nội địa vs giá xuất khẩu. Bảng phán quyết thuế chống bán phá giá mẫu của WTO.",
            "Cuối Tập 1 có Cliffhanger: 'Khi doanh nghiệp Việt bị kiện bán phá giá, nếu im lặng không trả lời bản câu hỏi thì điều tồi tệ gì sẽ xảy ra? Hãy bấm sang Tập 2 để xem chiến lược kháng kiện nhé!'",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm của bản thân Thúy: 'Cạnh tranh bằng giá rẻ không có sự hậu thuẫn của năng lực quản trị chi phí thực tế là con đường rủi ro nhất trong thương mại quốc tế'. Bản thân mình luôn hướng tới việc nâng cao giá trị chất lượng dịch vụ thay vì lao vào cuộc chiến dìm giá." """,
            "Chờ Chị Thúy duyệt"
        )
    ]

    # =========================================================================
    # SHEET 4: CM3_LICH_SU_THUONG_MAI (9 CHỦ ĐỀ)
    # =========================================================================
    data_cm3 = [
        (
            1, "CM3-CD01", "Đại Ký Sự Chiếc Thùng Container: Từ ý tưởng của anh tài xế xe tải nghèo đến cuộc cách mạng thu nhỏ hành tinh",
            "Ký Sự Dài (3 Tập Liên Hoàn: E01, E02, E03)", "25 - 30 phút (3 Tập)",
            "Sách 'The Box: How the Shipping Container Made the World Smaller and the World Economy Bigger' của Marc Levinson; ISO 668:1968; Wikipedia: Malcom McLean.",
            "Chào bạn. Chiếc áo bạn mặc, điện thoại bạn cầm... gần như chắc chắn từng nằm trong một chiếc hộp thép hình chữ nhật vượt đại dương. Bạn có biết: Trước 1956, chiếc hộp này hoàn toàn không tồn tại? Và người phát minh ra nó là một anh tài xế xe tải nghèo? Hôm nay Thúy mở màn Phần 1 của Đại Ký Sự lịch sử thương mại thế giới.",
            r"""[TOÀN VĂN KỊCH BẢN 3 TẬP LIÊN HOÀN - CHỊ THÚY]
=== TẬP 1: Ý TƯỞNG NGỚ NGẨN & CHUYẾN TÀU BÍ MẬT IDEAL X 1956 ===
26/4/1956 tại cảng Newark, tàu Ideal X chở 58 thùng container đầu tiên. Malcom McLean – tài xế xe tải nghèo năm 1934 – khi nhìn cảnh phu bốc từng bao bông sợi đã tự hỏi: 'Sao không bốc cả thùng xe tải lên tàu?'. Ông bán công ty xe tải số 2 nước Mỹ đánh cược tất cả. Kết quả chấn động: Chi phí bốc dỡ từ 5,86 USD/tấn rơi thẳng đứng xuống 0,16 USD/tấn (giảm 97%)!
=== TẬP 2: CUỘC CHIẾN SINH TỬ VỚI 50.000 PHU BẾN CẢNG & VỤ KIỆN THẾ KỶ ===
Nghiệp đoàn ILA của Teddy Gleason đại diện 50.000 phu bến cảng New York đình công tê liệt bờ Đông Mỹ vì sợ mất chén cơm; giới đường sắt kiện chống độc quyền. McLean đàm phán lập Quỹ bảo trợ thất nghiệp, trả phí trên mỗi tấn hàng bốc dỡ để phu bến cảng có trợ cấp và học nghề mới. Khi vừa yên ổn thì thảm họa kích thước ập đến: Mỗi hãng làm một cỡ (17ft, 24ft, 35ft), cần cẩu cảng quốc tế tê liệt!
=== TẬP 3: QUYẾT ĐỊNH BUÔNG BỎ TWISTLOCK & CHUẨN ISO 1968 THAY ĐỔI NHÂN LOẠI ===
Trước nguy cơ đổ vỡ, McLean nắm giữ bằng sáng chế then chốt: Khóa góc Twistlock. Ông làm hành động vĩ đại: Tuyên bố từ bỏ độc quyền, tặng miễn phí bản vẽ cho quốc tế! ISO ban hành chuẩn ISO 668 năm 1968 (chuẩn 20ft và 40ft). Thập niên 1960, container sang cảng Cam Ranh giải tỏa tàu quân nhu từ 20 ngày xuống 40 giờ, mở màn chuỗi cung ứng châu Á bùng nổ thống trị thế giới.""",
            "Quay Chị Thúy ngồi bên mô hình container mini và ấm trà gốm. Chèn ảnh tư liệu lịch sử: Chân dung Malcom McLean, tàu Ideal X 1956, biểu tình nghiệp đoàn ILA 1958, bản vẽ kỹ thuật khóa Twistlock, ảnh cảng Cam Ranh 1966. Dẫn chiếu rõ sách 'The Box' của Marc Levinson.",
            r"""[MÓC CÂU CLIFFHANGER GIỮA CÁC TẬP]:
- Cuối Tập 1 dẫn Tập 2: 'Thành công vang dội của Ideal X ngay lập tức kéo theo CƠN THỊNH NỘ KHỦNG KHIẾP từ 50.000 phu bến cảng New York và các tập đoàn đường sắt kiện tụng đòi bóp chết container! McLean đã chiến đấu thế nào? Bấm ngay vào Tập 2 trên màn hình để xem tiếp nhé!'
- Cuối Tập 2 dẫn Tập 3: 'Khi kích thước hỗn loạn khiến cần cẩu cảng quốc tế tê liệt, McLean đứng trước bờ vực phá sản. Lúc này ông đưa ra quyết định điên rồ nhất lịch sử: Buông bỏ bằng sáng chế độc quyền tỷ đô! Quyết định đó là gì? Hãy bấm vào Tập 3 kết thúc series để xem nhé!' """,
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học lớn nhất, sâu sắc nhất mà bản thân Thúy khắc cốt ghi tâm cho chính doanh nghiệp của mình từ câu chuyện của Malcom McLean: 'Đôi khi, đỉnh cao của chiến lược kinh doanh không phải là cố gắng giữ chặt lợi ích độc quyền nhỏ bé của mình, mà là dũng cảm buông bỏ để cùng kiến tạo nên một chuẩn mực lớn cho toàn ngành'." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            2, "CM3-CD02", "Tiệc trà Boston 1773 (Boston Tea Party): Khi một đạo luật thuế quan châm ngòi làm sụp đổ cả đế chế Anh",
            "2 Tập Liên Hoàn", "16 - 18 phút (2 Tập)",
            "National Archives Hoa Kỳ (Tài liệu Đạo luật Trà 1773 - Tea Act); History.com; Wikipedia: Boston Tea Party; Sách lịch sử thế giới cận đại.",
            "Chào bạn. Bạn có tin một tách trà thơm lừng trong phòng khách của người Anh quý tộc lại có thể là nguyên nhân làm bùng nổ cuộc chiến tranh giải phóng và khai sinh ra cường quốc số một thế giới hiện đại? Hôm nay Thúy mời bạn cùng uống trà và lắng nghe sự kiện thuế quan kịch tính nhất lịch sử nhân loại: Tiệc trà Boston 1773.",
            r"""[NỘI DUNG 2 TẬP LIÊN HOÀN - CHỊ THÚY]
=== TẬP 1: 342 THÙNG TRÀ DƯỚI LÀN NƯỚC LẠNH CẢNG BOSTON ===
Đêm 16/12/1773, nhóm 'Những người con của Tự do' cải trang thành thổ dân Mohawk đột nhập 3 tàu Dartmouth, Eleanor, Beaver, ném 342 thùng trà của Công ty Đông Ấn Anh xuống biển. Nguồn cơn từ Đạo luật Trà (Tea Act 1773) trao độc quyền xuất khẩu cho Công ty Đông Ấn và loại bỏ thương nhân địa phương.
=== TẬP 2: TUYÊN NGÔN BẤT HỦ 'NO TAXATION WITHOUT REPRESENTATION' ===
Phân tích nguyên tắc: Không có đại diện trong nghị viện thì không có quyền đánh thuế. Đòn trả đũa đóng cửa cảng Boston của Vua George III và sự bùng nổ của Cuộc chiến tranh giành độc lập 1776 khai sinh ra nước Mỹ.""",
            "Quay Chị Thúy bên ấm trà gốm mộc mạc. Tranh khắc lịch sử thế kỷ 18: Cảnh tàu buôn Dartmouth bị ném thùng trà, bản đồ 13 thuộc địa Bắc Mỹ, biểu tượng Công ty Đông Ấn Anh.",
            "Cuối Tập 1 có Cliffhanger: 'Hành động đổ 342 thùng trà trị giá 1,7 triệu USD xuống biển đã khiến Vua nước Anh phản ứng điên cuồng ra sao? Mời bạn bấm sang Tập 2 để theo dõi bước ngoặt lịch sử!'",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ tách trà Boston, bài học Thúy tự chiêm nghiệm: 'Thuế quan và thương mại không bao giờ chỉ đơn thuần là bài toán thu ngân sách; nó là thước đo của sự công bằng, lòng tin xã hội và sự tôn trọng đối với những người trực tiếp làm ra của cải vật chất'." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            3, "CM3-CD03", "Thương cảng Phố Hiến thế kỷ 17: Tiền nhân Đại Việt đã lập Ty Tuần ty và thu thuế tàu phương Tây bản lĩnh thế nào?",
            "2 Tập Liên Hoàn", "16 - 18 phút (2 Tập)",
            "Đại Việt Sử Ký Toàn Thư (Bản kỷ thực lục thời Hậu Lê - Chúa Trịnh); Thư tịch cổ Viện Sử học; Sách 'Phố Hiến - Thương cảng cổ Việt Nam'.",
            "Chào bạn. Dân gian ta có câu ca dao truyền đời: 'Thứ nhất Kinh Kỳ, thứ nhì Phố Hiến'. Trong tâm thức của nhiều người, chúng ta thường nghĩ thời xưa nước mình khép kín với phương Tây. Nhưng bạn có biết, ngay từ thế kỷ 17, cha ông ta tại Phố Hiến đã thiết lập cơ quan hải quan, thu thuế ngoại thương và kiểm dịch tàu buôn quốc tế vô cùng bài bản và đĩnh đạc?",
            r"""[NỘI DUNG 2 TẬP LIÊN HOÀN - CHỊ THÚY]
=== TẬP 1: THƯƠNG ĐIẾM PHƯƠNG TÂY NƠI CỬA SÔNG HỒNG ===
Khung cảnh Phố Hiến thế kỷ 17 tấp nập thuyền buôn Hà Lan (VOC), Bồ Đào Nha, Anh, Nhật Bản, Trung Hoa mang bạc và súng đổi tơ tằm, hồ tiêu. Hoạt động kiểm kê, cấp thị thực thương mại nghiêm ngặt của Ty Tuần ty cửa biển Ba Lạt.
=== TẬP 2: KIỂM DỊCH Y TẾ VÀ QUY CHẾ THU THUẾ BẢN LĨNH CỦA ĐẠI VIỆT ===
Quy định cách ly phòng dịch bệnh cho thủy thủ đoàn phương Tây ngoài bãi sông. Chuyện thuyền trưởng Hà Lan cậy tàu to súng lớn làm trái lệ bị quan Tuần ty kiên quyết cấm cập bến, buộc phải nộp phạt và cam kết theo đúng luật lệ triều đình.""",
            "Quay Chị Thúy trong trang phục trang nhã, bên cạnh có bình hoa sen nhỏ. Tranh khắc gỗ cổ Phố Hiến xưa, bản đồ dòng sông Hồng thế kỷ 17, tiêu bản tiền đồng và thư tịch chữ Hán về Ty Tuần ty.",
            "Cuối Tập 1 có Cliffhanger: 'Khi tàu chiến Hà Lan mang theo đại bác toan uy hiếp cửa sông, tiền nhân Đại Việt đã dùng mưu trí gì để buộc họ tuân thủ luật pháp? Hãy bấm sang Tập 2 để theo dõi nhé!'",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Mỗi khi nhìn lại lịch sử Phố Hiến, bản thân Thúy càng thấm thía: 'Tự tôn dân tộc và bản lĩnh ngoại thương đã nằm sẵn trong huyết quản của tiền nhân Đại Việt: Biết mở cửa bang giao để làm giàu non sông, nhưng luôn giữ vững kỷ cương luật pháp và chủ quyền kinh tế của đất nước'." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            4, "CM3-CD04", "Bộ luật Hammurabi 3.800 năm trước: Phiến đá cổ Babylon và nguyên tắc 'Mắt đền mắt' sinh ra để giới hạn sự trả thù trong giao thương",
            "1 Tập", "8 - 9 phút",
            "Bảo tàng Louvre Paris (Hiện vật phiến đá bazan khắc Bộ luật Hammurabi); Tài liệu Lịch sử Luật học cổ đại Đại học Yale; Wikipedia: Code of Hammurabi.",
            "Chào bạn. 3.800 năm trước tại vùng đất Lưỡng Hà cổ đại, một vị vua đã khắc 282 điều luật lên một khối đá bazan đen cao hơn 2 mét đặt giữa quảng trường thành Babylon. Nhiều người chỉ nhớ đến câu 'Mắt đền mắt, răng đền răng' và nghĩ đó là sự dã man. Nhưng sự thực, bộ luật cổ nhất nhân loại này được sinh ra để bảo vệ sự công bằng trong giao thương như thế nào?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bối cảnh xã hội Babylon cổ đại: Nhu cầu chuẩn mực hóa giao dịch vay mượn bạc, thuê mướn thuyền bè và mua bán gia súc.
- Nguyên tắc 'Mắt đền mắt' (Lex Talionis): Thời nguyên thủy, nếu một người bị móc mắt, họ sẽ giết cả bộ tộc kẻ thù. Hammurabi đặt ra luật để GIỚI HẠN SỰ TRẢ THÙ tương xứng với thiệt hại, ngăn chặn chiến tranh bạo lực.
- Các điều luật thương mại tiến bộ: Trách nhiệm bồi thường của thợ đóng thuyền nếu thuyền chìm làm hỏng hàng hóa; hợp đồng gửi giữ ngũ cốc phải có nhân chứng và con dấu đất sét niêm phong.""",
            "Quay Chị Thúy bên tách trà, màn hình hiển thị hình ảnh 3D phiến đá đen Hammurabi tại bảo tàng Louvre và bản dịch chữ hình nêm cổ Cuneiform.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm của Thúy: 'Từ thuở bình minh của nhân loại, con người đã hiểu rằng nếu không có luật lệ rõ ràng thì lòng tham và sự trả thù sẽ tàn phá mọi thành quả kinh tế'. Bản thân mình luôn xem hợp đồng bằng văn bản là cách nhân văn nhất để bảo vệ hòa khí giữa các bên." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            5, "CM3-CD05", "Công ty Đông Ấn Hà Lan (VOC 1602): Tập đoàn cổ phần đầu tiên trên thế giới và bài học về lòng tham kinh tế",
            "2 Tập Liên Hoàn", "16 - 18 phút (2 Tập)",
            "Lưu trữ Quốc gia Hà Lan (Nationaal Archief); Sách 'The Dutch East India Company: The History of the World's First Multinational Corporation'; Wikipedia: Dutch East India Company.",
            "Chào bạn. Một công ty tư nhân sở hữu quân đội riêng gần 50.000 lính, một hạm đội gần 200 tàu chiến, được quyền tự đúc tiền và tự tuyên chiến với các quốc gia khác. Giá trị vốn hóa của nó ở thời kỳ đỉnh cao ước tính lên đến 7,9 nghìn tỷ USD theo thời giá hiện nay – gấp đôi cả Apple và Microsoft cộng lại! Đó là Công ty Đông Ấn Hà Lan (VOC).",
            r"""[NỘI DUNG 2 TẬP LIÊN HOÀN - CHỊ THÚY]
=== TẬP 1: CỖ MÁY SIÊU LỢI NHUẬN TỪ HỒ TIÊU VÀ SỰ RA ĐỜI CỦA SÀN CHỨNG KHOÁN ===
Năm 1602, VOC ra đời và phát minh ra cổ phiếu có thể chuyển nhượng tự do đầu tiên trên thế giới tại sàn Amsterdam. Độc quyền buôn bán hương liệu nhục đậu khấu, đinh hương và tơ lụa châu Á với tỷ suất sinh lời hàng trăm phần trăm.
=== TẬP 2: SỰ SỤP ĐỔ CỦA GÃ KHỔNG LỒ TƯ BẢN BỞI THAM NHŨNG VÀ LÒNG THAM ===
Cơ chế quản lý quan liêu, tệ nạn buôn lậu của các giám đốc chi nhánh thương điếm, chi phí duy trì quân sự đắt đỏ và các khoản nợ khổng lồ dẫn đến sự phá sản tủi hổ năm 1799. Chữ viết tắt VOC bị người đời mỉa mai thành 'Vergaan Onder Corruptie' (Chết vì tham nhũng).""",
            "Quay Chị Thúy với mô hình thuyền buồm thương mại cổ thế kỷ 17. Tranh vẽ sàn giao dịch chứng khoán Amsterdam 1602 và biểu đồ vốn hóa khổng lồ của VOC.",
            "Cuối Tập 1 có Cliffhanger: 'Làm thế nào một công ty giàu hơn cả một đế chế lại rơi vào cảnh vỡ nợ không còn một xu dính túi? Hãy bấm sang Tập 2 để chứng kiến cái chết của gã khổng lồ!'",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học sâu sắc mà bản thân Thúy tự nhắc nhở mình trong kinh doanh: 'Quy mô càng lớn mà đạo đức quản trị và sự kiểm soát tài chính lỏng lẻo thì sự sụp đổ càng diễn ra nhanh chóng'. Phát triển bền vững dựa trên sự minh bạch luôn quan trọng hơn việc phình to quy mô bằng mọi giá." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            6, "CM3-CD06", "Con đường Tơ lụa cổ đại: Cửa ải Ngọc Môn Quan và cách đế chế Hán - La Mã quản lý thuế quan xuyên lục địa",
            "1 Tập", "8 - 9 phút",
            "UNESCO Silk Roads Programme; Sách 'The Silk Roads: A New History of the World' của Peter Frankopan; Di tích Ngọc Môn Quan Đôn Hoàng.",
            "Chào bạn. Hơn 2.000 năm trước, hàng ngàn đoàn thương buôn lạc đà chở tơ lụa, ngọc bích, đồ sứ từ Trung Hoa vượt qua sa mạc Gobi khắc nghiệt để tiến về La Mã cổ đại. Khi đến cửa ngõ Ngọc Môn Quan (Cửa ải Ngọc Bích) tại Đôn Hoàng, các thương nhân thời xưa đã phải dừng lại để làm thủ tục hải quan và nộp thuế ra sao?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Chức năng của Ngọc Môn Quan: Vừa là pháo đài quân sự canh giữ biên cương, vừa là trạm hải quan kiểm soát hộ chiếu (thẻ tre thông hành) và thu thuế hàng hóa.
- Cơ chế kiểm dịch thú y: Kiểm tra sức khỏe của đàn lạc đà và ngựa thồ hàng để ngăn ngừa mầm bệnh xâm nhập lãnh thổ.
- Sự trao đổi văn hóa và pháp lý: Con đường tơ lụa không chỉ là buôn bán tơ lụa, mà là nơi sinh ra các quy ước thương mại quốc tế đầu tiên giữa phương Đông và phương Tây.""",
            "Quay Chị Thúy với bình trà thảo mộc mộc mạc. Tranh tư liệu đoàn lạc đà sa mạc Đôn Hoàng, hình ảnh tàn tích lũy thành Ngọc Môn Quan và thẻ tre thông hành cổ.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm của Thúy: 'Thương mại quốc tế chính là sợi dây kết nối hòa bình bền chặt nhất giữa các nền văn minh'. Bản thân mình khi làm nghề xuất nhập khẩu luôn thấy vinh dự vì mỗi chuyến hàng mình đưa đi là một cây cầu nối liền sự hiểu biết giữa các quốc gia." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            7, "CM3-CD07", "Chiến tranh Thuốc phiện (1839 - 1842): Thâm hụt thương mại trà - bạc và bài học cay đắng về việc bế quan tỏa cảng",
            "1 Tập", "8 - 10 phút",
            "Sách 'The Opium War: Drugs, Dreams, and the Making of Modern China' của Julia Lovell; Thư tịch Lâm Tắc Từ tại Hổ Môn 1839; Wikipedia: First Opium War.",
            "Chào bạn. Một cuộc chiến tranh thảm khốc làm thay đổi toàn bộ lịch sử Á Đông hiện đại, nguyên nhân gốc rễ sâu xa lại bắt đầu từ... một sự mất cân bằng cán cân thương mại trà và bạc giữa hai quốc gia. Đó chính là Chiến tranh Thuốc phiện năm 1839. Bài học thương mại cay đắng nào vẫn còn nguyên giá trị đến hôm nay?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bối cảnh thâm hụt thương mại: Nước Anh phát cuồng vì trà của nhà Thanh, khiến hàng triệu ounce bạc ròng chảy ngược vào Trung Quốc. Để cứu vãn thâm hụt, Công ty Đông Ấn Anh đã buôn lậu thuốc phiện từ Ấn Độ sang Trung Quốc.
- Quan khâm sai Lâm Tắc Từ và biến cố tiêu hủy thuốc phiện tại Hổ Môn 1839: Đốt cháy hơn 1.000 tấn thuốc phiện tịch thu từ các thương thuyền Anh.
- Sự kiêu ngạo của triều đình Mãn Thanh với chính sách 'Bế quan tỏa cảng' khiến họ hoàn toàn mù tịt về sức mạnh công nghiệp và hải quân của thế giới phương Tây, dẫn đến thất bại thảm hại và các hiệp ước bất bình đẳng.""",
            "Quay Chị Thúy với ánh mắt suy tư sâu lắng. Tranh vẽ cảnh tiêu hủy thuốc phiện tại Hổ Môn 1839 và trận chiến hạm đội tàu hơi nước Nemesis trên sông Châu Giang.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học lịch sử mà bản thân Thúy luôn tự răn mình: 'Sự tự phụ và khép kín cánh cửa học hỏi là liều thuốc độc giết chết một doanh nghiệp nhanh nhất'. Trong kỷ nguyên hội nhập toàn cầu, cập nhật hiểu biết về luật chơi quốc tế là cách tự bảo vệ mình tốt nhất." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            8, "CM3-CD08", "Kênh đào Suez 1869: Vết cắt 193 km rút ngắn nửa vòng Trái Đất và cuộc chiến địa chính trị quanh tuyến hàng hải huyết mạch",
            "1 Tập", "8 - 9 phút",
            "Suez Canal Authority (SCA); Sách 'Parting the Desert: The Creation of the Suez Canal' của Zachary Karabell; Wikipedia: Suez Canal.",
            "Chào bạn. Hãy tưởng tượng nếu một ngày con đường đi từ nhà bạn đến công ty bị chặn lại và bạn phải đi vòng qua một cung đường xa gấp ba lần mất cả tuần lễ. Đó chính là những gì xảy ra với ngành vận tải biển toàn cầu mỗi khi Kênh đào Suez bị sự cố. Vết cắt nhân tạo 193 km xuyên qua sa mạc Ai Cập đã định hình lại chuỗi cung ứng thế giới như thế nào?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Kỳ tích đào kênh năm 1869 của kỹ sư Ferdinand de Lesseps: Rút ngắn 7.000 km hải trình từ châu Á sang châu Âu, không còn phải đi vòng qua Mũi Hảo Vọng đầy bão tố ở Nam Phi.
- Cú sốc siêu tàu Ever Given mắc cạn năm 2021: Tắc nghẽn 6 ngày làm gián đoạn gần 10 tỷ USD hàng hóa mỗi ngày, cước container vọt lên kỷ lục 15.000 USD/cont.
- Bài học đứt gãy chuỗi cung ứng và sự phụ thuộc vào các điểm nghẽn địa chính trị hàng hải (Chokepoints).""",
            "Quay Chị Thúy bên mô hình tàu container. Bản đồ vệ tinh tuyến hàng hải qua Kênh đào Suez vs tuyến đường vòng qua Mũi Hảo Vọng. Ảnh tư liệu tàu Ever Given mắc cạn.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ câu chuyện Kênh đào Suez, bài học Thúy tự rút ra cho việc quản trị chuỗi cung ứng của mình là: 'Không bao giờ được bỏ tất cả trứng vào một giỏ vận chuyển'. Bản thân mình luôn xây dựng sẵn các kịch bản dự phòng về tuyến đường và thời gian lưu kho an toàn cho khách hàng." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            9, "CM3-CD09", "Vụ trộm bức họa Mona Lisa 1911: Khi khoảng trống trên tường bảo tàng Louvre tạo nên kiệt tác đắt giá nhất nhân loại",
            "1 Tập", "8 - 9 phút",
            "Bảo tàng Louvre Paris (Tài liệu lưu trữ vụ trộm 21/08/1911); Sách 'Vanished Smile: The Mysterious Theft of the Mona Lisa' của R.A. Reily; Wikipedia: Mona Lisa.",
            "Chào bạn. Trước năm 1911, bức tranh Mona Lisa của Leonardo da Vinci chỉ là một bức họa bình thường treo lọt thỏm giữa hàng ngàn kiệt tác khác trong bảo tàng Louvre, hầu như không mấy ai chú ý. Nhưng sau một vụ trộm kỳ lạ của một người thợ mộc Ý, bức tranh bỗng chốc trở thành báu vật đắt giá nhất hành tinh. Chuyện gì đã xảy ra?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Diễn biến vụ trộm ngày 21/08/1911: Vincenzo Peruggia giấu bức tranh dưới áo khoác bước ra khỏi bảo tàng Louvre với lý do 'yêu nước muốn trả tranh về cho nước Ý'.
- Nghịch lý tâm lý loài người: Trong suốt 2 năm bức tranh biến mất, hàng chục ngàn người xếp hàng dài chỉ để vào bảo tàng ngắm nhìn... 4 chiếc đinh gỉ và một khoảng trống trên tường!
- Khi được tìm thấy năm 1913, Mona Lisa trở thành hiện tượng truyền thông toàn cầu. Bài học về sự hình thành giá trị tài sản vô hình và các đạo luật bảo tồn di sản văn hóa quốc tế.""",
            "Quay Chị Thúy bên tách trà thơm. Ảnh tư liệu trắng đen: Khoảng trống trên tường bảo tàng Louvre năm 1911, chân dung kẻ trộm Vincenzo Peruggia, và bức tranh Mona Lisa ngày nay.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm sâu sắc của Thúy: 'Đôi khi giá trị của một thứ chỉ thực sự được nhận ra khi nó mất đi'. Bản thân mình luôn tự nhắc nhở phải trân trọng những đối tác, những cộng sự và những giá trị bình dị đang kề vai sát cánh cùng mình mỗi ngày trước khi quá muộn." """,
            "Chờ Chị Thúy duyệt"
        )
    ]

    # =========================================================================
    # SHEET 5: CM4_TAM_TINH_THUC_CHIEN (9 CHỦ ĐỀ)
    # =========================================================================
    data_cm4 = [
        (
            1, "CM4-CD01", "Chào các bạn, mình là Thúy: Tốt nghiệp Ngân hàng, rẽ sang Xuất Nhập Khẩu & Lô hàng đầu tiên lãi đúng 2 triệu",
            "Clip Chào Sân (1 Tập)", "5 - 6 phút",
            "Câu chuyện thật của Thúy: Tốt nghiệp Học viện Ngân hàng (Kinh doanh quốc tế); Trải nghiệm làm công ty startup 'tiền đô nhưng bỏ túi tiền lẻ'; Lô mỹ phẩm đầu đời và kiếp nạn Cổng một cửa & C/O ủy thác; Tâm sự nghề Logistics trên Facebook.",
            "Chào các bạn, mình là Thúy. Hôm trước có người bạn hỏi mình: 'Hồi xưa học Học viện Ngân hàng, chuyên ngành Kinh doanh quốc tế, bạn bè ra trường vào ngân hàng làm hết, sao Thúy lại rẽ ngang sang làm xuất nhập khẩu làm gì cho cực vậy?' Nói thật với mọi người, cái nghề này nhìn từ ngoài vào ai cũng nghĩ nó sang chảnh lắm. Làm xuất nhập khẩu, làm logistics, nghe toàn tiền đô với đối tác nước ngoài. Nhưng chỉ ai làm rồi mới hiểu: mang tiếng làm tiền đô nhưng toàn bỏ túi tiền lẻ thôi. Mà rủi ro thì nhiều vô kể, từ buôn bán ngoài thị trường cho tới giấy tờ, thủ tục, pháp luật...",
            r"""[TOÀN VĂN LỜI THOẠI BÌNH DỊ 5 PHẦN - CHỊ THÚY]

1. XUẤT PHÁT ĐIỂM & LÀM Ở CÔNG TY STARTUP:
Hồi mới ra trường, chưa có kinh nghiệm nên đi xin việc chắc chỉ có mỗi vị trí sales là người ta nhận. May rủi thế nào mình lại vào đúng một công ty startup. 
Mà đã làm startup thì mang tiếng là sales, nhưng việc gì cũng đến tay: từ đi tìm nguồn hàng, mặc cả giá, đặt tàu bè, tự gõ chứng từ, tự lên tờ khai hải quan, cho đến làm cả hóa đơn sổ sách. Đúng nghĩa là chỉ thiếu mỗi bước tự bốc hàng vào kho nữa thôi.
Và lô hàng đầu tiên mình làm lên bờ xuống ruộng, tính đi tính lại, tiền lãi mang về cho công ty vỏn vẹn đúng 2 triệu đồng.

2. CÚ VẤP ĐẦU ĐỜI VỚI LÔ MỸ PHẨM:
2 triệu tiền lãi đó đổi lại cho mình cả rổ kiếp nạn. Đó là một lô hàng mỹ phẩm. Hồi đấy còn lơ ngơ, mình cứ nghĩ đơn giản là hồ sơ có đủ rồi thì cứ nộp lên Cổng một cửa quốc gia, rồi ngồi đợi kết quả là xong. Ai ngờ hệ thống từ chối thẳng, báo hồ sơ sai định dạng, thiếu nội dung mà lúc đấy mình chẳng hiểu nổi vì sao lại sai. Thế là mấy đêm liền phải thức trắng, lọ mọ đọc kỹ từng dòng thông tư, rồi đi tìm xem hồ sơ của các bên khác người ta làm thế nào để học theo mà sửa.
Chưa hết, vẫn là lô hàng đó, lại dính tiếp vụ làm C/O 3 bên. Nhưng bên bán họ lại dùng C/O ủy thác. Cái này mà hải quan bác một cái là bị phạt thuế và đền tiền ngay. Lúc đấy lo đến thót tim, nhưng cũng nhờ thế mà mình buộc phải mày mò tìm mọi cách xử lý cho bằng được.

3. ĐỌC BÀI VIẾT BỎ NGHỀ & TÂM NIỆM TRƯỚC TUỔI 30:
Cách đây mấy hôm, mình đọc được bài tâm sự của một bạn trẻ: 'Em quyết định bỏ nghề logistics sau 2 năm vì áp lực quá, muốn tìm công việc nào không liên quan đến pháp luật'.
Đọc xong, mình thấy đồng cảm lắm. Vì ai làm nghề này cũng từng nếm trải những đêm thức trắng đến 2-3 giờ sáng ngóng tàu, những hôm thót tim khi tờ khai bị nhảy luồng đỏ, hay nơm nớp lo: chỉ cần gõ sai một con số hay một chữ cái trên chứng từ thôi là có thể phải đền tiền triệu, tiền chục triệu rồi. Áp lực doanh số, áp lực đền hàng, nhiều lúc mệt mỏi thật sự.
Nhưng mình cũng thấy hơi tiếc cho bạn ấy.
Hồi mới đi làm, mình luôn tự dặn lòng một điều: Trước năm 30 tuổi, điều quan trọng chưa phải là chỗ đó trả cho mình bao nhiêu tiền, mà là chỗ đó có cho mình được làm, được va vấp, được thử mọi thứ hay không. Để sau 30 tuổi, những kinh nghiệm, những lần tự mình vượt qua khó khăn đó mới là chỗ đứng vững chắc nhất của mình. Tiền lúc đầu có thể chưa nhiều, nhưng qua mỗi lô hàng, mình thấy mình giàu hơn về mặt kiến thức. Và cái đó thì không ai lấy đi được.
Còn cái chuyện bạn ấy bảo muốn tìm nghề không dính đến pháp luật, nói thật lòng là: Đã sống và đi làm thì chẳng có nghề nào không dính đến pháp luật cả. Mở quán trà đá vỉa hè, đi làm văn phòng đóng thuế thu nhập, hay nhập lô hàng về bán... việc gì cũng có luật cả. Thay vì sợ và né tránh, thì thà mình biết sớm một chút, hiểu sâu một chút để tự bảo vệ mình, tránh những sai lầm làm mình mất tiền sau này.

4. VÌ SAO MÌNH MỞ KÊNH NÀY?
Đấy cũng là lý do vì sao mình mở kênh này.
Mình không lên đây để dạy bảo ai cả. Mình chỉ muốn ngồi lại kể cho các bạn nghe những kinh nghiệm thật, những bài học xương máu mình đã trải qua, và chia sẻ góc nhìn cá nhân về luật, về thuế, về xuất nhập khẩu — những thứ ai đi làm hay kinh doanh cũng nên biết một chút.
Nếu bạn là người mới vào nghề đang thấy chán nản, hay bạn đang muốn tự nhập khẩu hàng về bán mà chưa biết bắt đầu từ đâu, thì mình hi vọng những chia sẻ giản dị này sẽ giúp bạn đỡ bỡ ngỡ và tự tin hơn.""",
            "Quay Chị Thúy ngồi trò chuyện thoải mái, tự nhiên tại bàn làm việc. Ánh sáng tự nhiên, góc quay đối diện gần gũi. Nói chuyện như đang ngồi tâm sự với một người bạn uống cà phê, giọng điệu mộc mạc, không lên gân.",
            "Không áp dụng (Clip độc lập chào sân, mở đầu làm quen).",
            r"""[KẾT LUẬN TỰ DẶN LÒNG CỦA THÚY]
"Với bản thân Thúy, mỗi lần ngồi nói chuyện như thế này cũng là một lần tự nhắc nhở chính mình: làm nghề này phải luôn cẩn thận, tôn trọng pháp luật và chịu khó học hỏi mỗi ngày. Vì chỉ cần một chút chủ quan thôi là bài học phải trả bằng tiền thật.
Cảm ơn các bạn đã lắng nghe câu chuyện của mình nhé. Hẹn gặp lại mọi người ở những video sau!" """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            2, "CM4-CD02", "Sự cố gãy seal chuyển cảng: 3 bước lập biên bản hiện trường cứu lô hàng trị giá vài tỷ khỏi luồng Đỏ kiểm hóa 100%",
            "1 Tập", "7 - 8 phút",
            "Thực tế hiện trường giám sát hải quan cảng Cát Lái / Hải Phòng; Quy trình chuyển cửa khẩu Thông tư 38/2015 Điều 51; Ghi chép Daily Note & FB Th.xuatnhapkhau.",
            "Chào bạn. Có bao giờ bạn rơi vào cảm giác: Xe container chở lô hàng trị giá vài tỷ của công ty vừa về đến kho thì tài xế báo chì niêm phong (seal) của hãng tàu và hải quan bị gãy rồi? Khách hàng sợ hãi định tự ý lấy dây buộc lại. Hôm nay Thúy kể câu chuyện thực tế về chiếc kẹp chì nhỏ bé này.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bản chất pháp lý của chiếc chì niêm phong (seal bolt): Bằng chứng xác nhận tính nguyên trạng của hàng hóa chịu sự giám sát hải quan.
- Nguy cơ nếu tự ý cắt seal: Bị nghi ngờ đánh tráo hàng hóa buôn lậu, đình chỉ thông quan và chuyển kiểm hóa 100% luồng Đỏ.
- 3 bước xử lý chuẩn chỉ:
  1. Giữ nguyên hiện trường 100%, chụp ảnh góc rộng xe và góc cận chiếc seal gãy.
  2. Lập Biên bản ghi nhận sự cố bất khả kháng có chữ ký 3 bên (lái xe vận tải, đại diện kho, công chức hải quan giám sát).
  3. Phối hợp kiểm đếm đối chiếu packing list ngay tại chỗ và niêm phong lại.""",
            "Quay Chị Thúy trên tay cầm một chiếc chì niêm phong (seal bolt) chuẩn của hãng tàu. Bàn trà có tài liệu biên bản hiện trường mẫu. Chèn hình ảnh xe container và vị trí bấm seal cửa sau.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Trong nghề XNK, sự cố là điều không ai mong muốn. Nhưng bài học mà bản thân Thúy tự nhắc mình: 'Người làm nghề bản lĩnh không phải là người không bao giờ gặp sự cố, mà là người biết giữ sự bình tĩnh, tuân thủ nguyên tắc pháp lý để biến nguy cơ thành sự an toàn'. Một biên bản hiện trường minh bạch luôn cứu sống doanh nghiệp tốt hơn ngàn lời giải thích." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            3, "CM4-CD03", "Nỗi đau con ốc vít cơ khí: Thuế từ 0% nhảy lên 25% vì áp sai mã HS và bài học làm thủ tục xác định trước mã số (Advance Ruling)",
            "1 Tập", "7 - 8 phút",
            "Biểu thuế XNK Việt Nam (Chương 73 vs Chương 84); Điều 28 Luật Hải quan 2014 về Xác định trước mã số; Hồ sơ tư vấn thực tế của Eureka.",
            "Chào bạn. Cùng là một con ốc vít bằng thép, nhưng nếu là ốc vít thông thường thì thuế nhập khẩu 0%, còn nếu phân loại là linh kiện đặc thù của máy in công nghiệp thì thuế nhảy lên 25% kèm thuế tự vệ thương mại. Chỉ lệch đúng một con số trong mã 8 chữ số, doanh nghiệp bị truy thu và phạt chậm nộp hàng trăm triệu!",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Câu chuyện thực tế một doanh nghiệp nhập khẩu linh kiện cơ khí bị bác mã HS sau 2 năm thông quan.
- Tại sao cơ quan hải quan không chấp nhận lý do 'bên em không biết'? Nguyên tắc người khai hải quan tự khai tự chịu trách nhiệm.
- Quy trình làm thủ tục Xác định trước mã số (Advance Ruling): Nộp hồ sơ và mẫu sản phẩm lên Tổng cục Hải quan trước 60 ngày để có văn bản xác định mã số có giá trị pháp lý toàn quốc.""",
            "Quay Chị Thúy cầm vài con ốc vít cơ khí trên tay, đặt cạnh cuốn Biểu thuế XNK. Slide đồ họa so sánh mức thuế nhóm 73.18 vs 84.43.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học đắt giá mà bản thân Thúy tự rút ra cho mình: 'Chi phí cho sự cẩn trọng ban đầu luôn rẻ hơn gấp ngàn lần chi phí đi khắc phục hậu quả'. Bản thân mình không bao giờ tiếc thời gian làm thủ tục xác định trước mã số cho những mặt hàng nhạy cảm." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            4, "CM4-CD04", "Bóc tách minh bạch phụ phí cước biển (THC, CIC, BAF) khi khách hàng sỉ thắc mắc: Rõ ràng từ đầu để giữ trọn tình nghĩa",
            "1 Tập", "7 - 8 phút",
            "Bảng niêm yết phụ phí địa phương (Local Charges) của các Hãng tàu quốc tế; Ghi chép đàm phán hợp đồng logistics từ Daily Note.",
            "Chào bạn. Có bao giờ bạn báo giá cước biển cho khách hàng sỉ là 500 USD/cont, nhưng khi hàng về cảng, bạn gửi thêm bảng thông báo thu phụ phí THC, CIC, D/O, vệ sinh cont lên đến 15 triệu đồng khiến khách hàng giận dữ cho rằng bạn đang 'chặt chém' không? Làm sao để bóc tách chi phí mà vẫn giữ được lòng tin?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Bản chất các khoản phụ phí Hãng tàu thu tại cảng đến: Phí xếp dỡ tại cảng (THC), Phí mất cân bằng vỏ cont (CIC), Phí chứng từ (D/O).
- Tâm lý khách hàng B2B: Họ không sợ chi phí cao, họ chỉ sợ chi phí PHÁT SINH BẤT NGỜ ngoài dự tính.
- Kỹ thuật báo giá Landed Cost trọn gói minh bạch: Tách riêng giá cước biển (Ocean Freight) và phụ phí niêm yết của Hãng tàu kèm Debit Note chứng từ gốc.""",
            "Quay Chị Thúy bên tách trà, trên bàn có bản Debit Note phụ phí hãng tàu mẫu. Bảng so sánh báo giá mập mờ vs Báo giá bóc tách minh bạch.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm lớn nhất của Thúy trong bán hàng B2B: 'Sự rõ ràng và minh bạch ngay từ đầu là chiếc chìa khóa duy nhất để giữ trọn tình nghĩa với khách hàng'. Bản thân mình luôn chủ động giải thích từng dòng phụ phí cho khách trước khi ký hợp đồng để không bao giờ có sự hiểu lầm." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            5, "CM4-CD05", "'Duy nhất 1 điểm chạm': Bài học đắt giá về tái cấu trúc trải nghiệm khách hàng B2B khi nhân sự bàn giao việc lộn xộn",
            "1 Tập", "7 - 8 phút",
            "Nhật ký điều hành nội bộ Eureka; Case study tái cấu trúc trải nghiệm khách hàng B2B từ Daily Note 19/08/2026.",
            "Chào bạn. Một khách hàng gọi điện cho kế toán hỏi về hóa đơn thì kế toán bảo 'chị hỏi bạn sale', gọi cho sale thì sale bảo 'em đang hỏi bạn hiện trường làm thủ tục ở cảng'... Khách hàng như bị đá bóng giữa các phòng ban. Bạn có từng chứng kiến sự hỗn loạn này trong chính doanh nghiệp của mình?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Nỗi đau của khách hàng khi phải làm việc với quá nhiều đầu mối nhân sự trong một lô hàng XNK.
- Chiến lược 'Duy nhất 1 điểm chạm' (Single Point of Contact - SPOC): Một chuyên viên dịch vụ khách hàng (Account Manager) chịu trách nhiệm toàn diện từ đầu đến cuối cho mọi thắc mắc của khách.
- Xây dựng hệ thống CRM & AI Agent cập nhật trạng thái lô hàng tự động vào Zalo giúp khách hàng luôn nắm rõ tiến độ 24/7 mà không cần hỏi.""",
            "Quay Chị Thúy với góc nhìn lãnh đạo điềm tĩnh. Sơ đồ quy trình làm việc phân tán cũ vs Quy trình '1 Điểm Chạm' mới.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học mà bản thân Thúy tự chiêm nghiệm sâu sắc trong quản trị: 'Khách hàng không mua quy trình nội bộ của bạn, họ mua sự an tâm và cảm giác được chăm sóc chu đáo'. Thúy luôn tự nhắc mình phải tinh gọn bộ máy để khách hàng chỉ cần nhớ đến một người bạn đồng hành duy nhất." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            6, "CM4-CD06", "Nghệ thuật từ chối khi người thân nhờ đứng tên giùm doanh nghiệp hoặc mượn giấy phép XNK: Mất lòng trước, được lòng sau",
            "1 Tập", "7 - 8 phút",
            "Quy định về trách nhiệm của Người đại diện theo pháp luật trong Luật Doanh nghiệp 2020; Rủi ro pháp lý ủy thác xuất nhập khẩu.",
            "Chào bạn. Một người bạn thân hoặc người họ hàng thân thiết đến gặp bạn và nói: 'Bác cho cháu nhờ đứng tên đại diện công ty này một thời gian', hoặc 'Cho anh mượn pháp nhân công ty em để nhập lô hàng này về nhé, anh em tin nhau là chính'. Trong tình huống đó, bạn sẽ nhận lời vì cả nể hay dũng cảm từ chối?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Rủi ro pháp lý khôn lường: Người đại diện theo pháp luật chịu trách nhiệm vô hạn đối với các khoản nợ thuế, vi phạm buôn lậu và nghĩa vụ phá sản của doanh nghiệp.
- Rủi ro của việc cho mượn pháp nhân XNK: Nếu lô hàng chứa hàng cấm, hàng giả nhãn mác, công ty bạn sẽ bị khởi tố hình sự!
- Kỹ thuật từ chối tinh tế, giữ trọn tình nghĩa: Đưa ra lý do nguyên tắc kiểm toán và cơ chế quản trị công ty, đồng thời hỗ trợ họ kết nối dịch vụ ủy thác XNK chính ngạch chuyên nghiệp.""",
            "Quay Chị Thúy bên tách trà ấm, biểu cảm chân thành, sâu lắng. Chèn các điều khoản trích dẫn về trách nhiệm pháp lý trong Luật Doanh nghiệp.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm lớn nhất mà bản thân Thúy tự răn mình: 'Sự cả nể trong kinh doanh là mầm mống phá hủy mọi mối quan hệ thân thiết nhất'. Bản thân mình học được cách nói lời từ chối dứt khoát nhưng ấm áp: Mất lòng trước để được lòng sau, giữ an toàn cho mình và giữ gìn tình nghĩa lâu dài." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            7, "CM4-CD07", "Đàm phán với nhà cung cấp nước ngoài khi hàng về bị lỗi quy cách: Cách dùng hợp đồng để họ tự nguyện bồi thường không tranh cãi",
            "1 Tập", "7 - 8 phút",
            "Công ước Viên 1980 về mua bán hàng hóa quốc tế (CISG); Điều khoản Khiếu nại (Claim) và Trọng tài thương mại trong Hợp đồng Ngoại thương.",
            "Chào bạn. Khui container hàng nhập khẩu từ Trung Quốc hoặc châu Âu ra, phát hiện 30% sản phẩm bị trầy xước, sai màu hoặc lỗi kỹ thuật không thể bán được. Bạn nhắn tin cho xưởng thì họ im lặng hoặc đổ lỗi cho đơn vị vận tải biển. Làm thế nào để buộc đối tác ngoại phải hoàn tiền hoặc bù hàng trong đơn tiếp theo?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Sai lầm lớn nhất: Đàm phán bằng cảm xúc và lời than vãn.
- 'Bộ chứng từ khiếu nại quốc tế' chuẩn chỉ:
  1. Biên bản kiểm nghiệm độc lập của bên thứ ba (SGS, Vinacontrol) lập ngay tại cảng hoặc kho nhận.
  2. Video unboxing cont và ảnh chụp chi tiết số seri sản phẩm bị lỗi.
  3. Thư khiếu nại thương mại (Claim Letter) dẫn chiếu đúng Điều khoản Bảo hành và Chất lượng trong Hợp đồng gốc.
- Đàm phán win-win: Đề xuất trừ tiền trực tiếp vào công nợ đơn hàng tiếp theo kèm chính sách giảm giá bồi thường.""",
            "Quay Chị Thúy bên tách trà, tay cầm Thư khiếu nại Claim Letter mẫu và Chứng thư giám định hàng hóa của Vinacontrol/SGS.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học Thúy tự rút ra sau nhiều lần đàm phán quốc tế: 'Hợp đồng chặt chẽ là công cụ đàm phán hòa bình và hiệu quả nhất'. Bản thân mình luôn dành thời gian đàm phán kỹ điều khoản khiếu nại và giám định trước khi chuyển tiền đặt cọc cho bất kỳ xưởng nào." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            8, "CM4-CD08", "Tách Bill (B/L) cho hàng hỗn hợp dính Quyết định 31: Tuyệt chiêu gỡ phí Split Bill 150-200 USD của Hãng tàu",
            "1 Tập", "7 - 8 phút",
            "Quy trình tách vận đơn đường biển theo Quyết định 31/2026/QĐ-TTg; Thực tế làm việc với các Hãng tàu ngoại từ Daily Note W34.",
            "Chào bạn. Để cứu lô hàng hỗn hợp không bị giam tại cảng biển theo Quyết định 31, bạn làm thủ tục tách Bill (B/L) thành 2 bộ. Nhưng ngay lập tức Hãng tàu phía nước ngoài đòi thu phí tách bill lên đến 150 - 200 USD mỗi bộ! Làm sao để gỡ các khoản phụ phí vô lý này?",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Tại sao Hãng tàu lại thu phí tách bill? Chi phí sửa đổi manifest điện tử và phát hành lại bộ vận đơn gốc.
- 3 tuyệt chiêu thực chiến:
  1. Đàm phán điều khoản Incoterms FOB để doanh nghiệp Việt Nam trực tiếp chỉ định forwarder, kiểm soát bảng giá phụ phí niêm yết.
  2. Thỏa thuận điều khoản 'Free Split Bill' ngay từ đầu khi chốt sản lượng vận chuyển với forwarder.
  3. Yêu cầu đối soát Debit Note chứng từ gốc của Hãng tàu để tránh bị đơn vị trung gian nâng giá khống.""",
            "Quay Chị Thúy bên màn hình so sánh bảng giá CIF vs FOB, mẫu Debit Note phụ phí tách bill của Hãng tàu.",
            "Không áp dụng (Tập đơn lẻ).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ những ca thực chiến gỡ phụ phí, Thúy tự dặn mình: 'Hiểu rõ luật chơi và chủ động giữ quyền chỉ định vận tải là cách tốt nhất để bảo vệ từng đồng lợi nhuận của doanh nghiệp'. Bản thân mình không bao giờ chấp nhận nộp một khoản phí nào nếu chưa nhìn thấy chứng từ gốc." """,
            "Chờ Chị Thúy duyệt"
        ),
        (
            9, "CM4-CD09", "15 năm lăn lộn nghề Thương mại & XNK: 3 chân lý 'dĩ bất biến, ứng vạn biến' giúp Thúy giữ tâm bình thản",
            "1 Tập", "8 - 10 phút",
            "Đúc kết triết lý sống và làm nghề sau 15 năm điều hành doanh nghiệp XNK Eureka; Triết lý năng lượng Giáp Mộc vững chãi & Liêm Trinh minh bạch.",
            "Chào bạn. 15 năm đi qua đủ những thăng trầm của thương trường, từ những ngày đầu chập chững bước chân ra cảng biển cho đến khi trực tiếp điều hành hàng ngàn chuyến hàng xuyên biên giới... Nếu được đúc kết lại 3 chân lý sâu sắc nhất giúp Thúy giữ được sự bình thản trước mọi bão giông luật lệ, mình muốn ngồi lại uống chén trà và tâm tình cùng bạn.",
            r"""[NỘI DUNG CHI TIẾT - CHỊ THÚY]
- Chân lý 1: 'Làm đúng từ đầu luôn là con đường ngắn nhất và rẻ nhất'. Chi phí cho sự tuân thủ (compliance) có thể tốn vài triệu, nhưng chi phí khắc phục sai phạm hải quan và thuế luôn tính bằng trăm triệu và danh dự.
- Chân lý 2: 'Biết luật là để tự do, không phải để sợ hãi'. Khi hồ sơ chứng từ minh bạch, tờ khai chuẩn chỉ, bạn có thể tự tin ngẩng cao đầu trước mọi cuộc thanh kiểm tra sau thông quan.
- Chân lý 3: 'Thương mại là chiếc cầu nối của sự tử tế'. Mỗi container hàng xuất nhập khẩu không chỉ là lợi nhuận, mà còn là uy tín của con người và niềm tự hào của hàng hóa Việt Nam.""",
            "Quay Chị Thúy bên tách trà thơm bốc khói, góc nhìn hoàng hôn cảng biển hoặc ánh đèn phòng làm việc dịu ấm. Nụ cười nhẹ nhàng, phong thái tự tại, bao dung.",
            "Không áp dụng (Tập tổng kết triết lý nhân sinh).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Khép lại những chia sẻ thực chiến này, lời tự dặn lòng lớn nhất của Thúy là: 'Giữ cho tâm mình liêm chính và vững chãi như cây gỗ lớn, thì dù gió bão cuộc đời có đổi hướng ra sao, mình vẫn luôn tìm thấy sự an yên đích thực'. Cảm ơn mọi người đã lắng nghe và đồng hành cùng Thúy!" """,
            "Chờ Chị Thúy duyệt"
        )
    ]

    # Populate the 4 Chuyên mục sheets
    populate_cm_sheet(
        "CM1_VAN_BAN_NONG",
        "CHUYÊN MỤC 1: TRÀ CHIỀU & VĂN BẢN NÓNG (THỜI SỰ PHÁP LÝ & CHÍNH SÁCH NỔI CỘM)",
        "Nguồn: Văn bản pháp luật mới (Nghị định, Thông tư, Công văn) + Cào Facebook/Google các điểm nóng dư luận | 9 Chủ đề chi tiết",
        data_cm1
    )

    populate_cm_sheet(
        "CM2_DICH_LUAT_TIENG_NGUOI",
        "CHUYÊN MỤC 2: DỊCH LUẬT RA TIẾNG NGƯỜI (GIẢI MÃ KHÁI NIỆM & VĂN BẢN KHÔ KHAN)",
        "Nguồn: Luật Hải quan 2014, Luật QLT 38, Incoterms 2020 ICC, Giáo trình Ngoại thương & Kế toán chuẩn | 9 Chủ đề chi tiết",
        data_cm2
    )

    populate_cm_sheet(
        "CM3_LICH_SU_THUONG_MAI",
        "CHUYÊN MỤC 3: KÝ SỰ LỊCH SỬ THƯƠNG MẠI (BÀI HỌC THẾ GIỚI & ĐẠI VIỆT 100% TÌNH TIẾT THẬT)",
        "Nguồn: Wikipedia, sách The Box (Marc Levinson), National Archives US, Đại Việt Sử Ký Toàn Thư | 9 Chủ đề (Có Series 3 tập liên hoàn & Cliffhanger)",
        data_cm3
    )

    populate_cm_sheet(
        "CM4_TAM_TINH_THUC_CHIEN",
        "CHUYÊN MỤC 4: TÂM TÌNH THỰC CHIẾN NGHỀ XNK & CUỘC SỐNG",
        "Nguồn: Daily Note thực tế công việc + Bài đăng trang cá nhân https://www.facebook.com/Th.xuatnhapkhau/ | 9 Chủ đề chi tiết (Kèm Clip Chào Sân)",
        data_cm4
    )

    # =========================================================================
    # SHEET 6: CHECKLIST QUAY BATCH SERIES
    # =========================================================================
    ws_cl = wb.create_sheet(title="CHECKLIST_QUAY_BATCH")
    ws_cl.views.sheetView[0].showGridLines = True

    ws_cl.merge_cells("A1:G1")
    ws_cl["A1"] = "CHECKLIST QUẢN TRỊ QUAY BATCH SERIES (MỖI ĐỢT QUAY TRỌN 1 CHUYÊN MỤC HOẶC 1 SERIES)"
    ws_cl["A1"].font = font_title
    ws_cl["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws_cl.merge_cells("A2:G2")
    ws_cl["A2"] = "Mục tiêu: Tối ưu thời gian của Chị Thúy, 1 buổi quay (3-4 tiếng) hoàn tất trọn vẹn 1 series 3 tập liên hoàn hoặc 3 kịch bản đơn lẻ để dựng trước 1 tuần."
    ws_cl["A2"].font = font_subtitle

    headers_cl = ["Hạng mục chuẩn bị", "Chi tiết kỹ thuật & Đạo cụ", "Tiêu chuẩn nghiệm thu (Quality Gate)", "Người phụ trách", "Ghi chú thực chiến"]
    for col_idx, h in enumerate(headers_cl, 1):
        cell = ws_cl.cell(row=4, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    checklist_data = [
        ("1. Không gian & Ánh sáng", "Bàn gỗ mộc, đèn Key Light 3500K (vàng ấm), Fill Light dịu nhẹ, đèn viền tóc (Rim light). Phía sau là giá sách hoặc chậu cây tĩnh lặng.", "Không bị bóng gắt trên mặt, tạo cảm giác ấm cúng, thư thái như ngồi trò chuyện tại nhà.", "Đội ngũ Media / Kỹ thuật", "Tránh ánh sáng trắng xanh 5600K kiểu phòng mổ hoặc livestream bán hàng thông thường."),
        ("2. Âm thanh (Audio First)", "Micro cài áo không dây (Rode Wireless PRO / Hollyland Lark) gắn gọn gàng, có lọc gió nhẹ. Thu file backup WAV 24-bit 48kHz.", "Âm thanh trong trẻo, giọng đọc trầm ấm, không có tiếng quạt gió hay tạp âm đường phố.", "Chị Thúy & Kỹ thuật âm thanh", "Giọng nói từ tốn (130-150 từ/phút), ngắt nghỉ tự nhiên sau mỗi luận điểm quan trọng."),
        ("3. Đạo cụ theo từng Chuyên mục", "CM1: eTax Mobile + Đĩa bánh hamburger mini + Danh mục QĐ 31.\nCM2: Cuốn Luật Hải quan + Mô hình tàu FOB/CIF + Báo cáo Mẫu 15.\nCM3: Mô hình container mini + Bản vẽ khóa Twistlock + Ấm trà gốm mộc.\nCM4: Chiếc seal chì hải quan thật + ốc vít cơ khí + biên bản hiện trường mẫu.", "Đạo cụ thật 100%, đặt tự nhiên trên bàn để làm điểm tựa visual khi nói.", "Bộ phận Hậu cần", "Không để quá nhiều đồ lộn xộn trên bàn, giữ mặt bàn mộc mạc tinh tế."),
        ("4. Máy nhắc chữ (Teleprompter)", "Cài đặt file kịch bản lời thoại chi tiết vào iPad/máy nhắc chữ đặt thẳng trước ống kính Camera 1.", "Chữ chạy vừa tốc độ nói, mắt nhìn tự nhiên vào ống kính, không bị liếc đảo.", "Trợ lý nội dung", "Chỉ dùng để nhớ ý chính và các con số/điều luật, không đọc như tụng kinh."),
        ("5. Phân công 2 Góc máy", "Cam 1 (Chính diện Eye-Level): 50mm f/1.8 quay từ ngực trở lên (tâm tình chính).\nCam 2 (Góc nghiêng 45 độ): 85mm f/1.8 quay cận góc nghiêng khi trầm ngâm, chiêm nghiệm.", "Chuyển góc máy mượt mà mỗi 15-20 giây để giữ chân người xem YouTube.", "Camera Operator / Editor", "Đồng bộ Timecode giữa 2 máy để hậu kỳ cắt dựng nhanh chóng."),
        ("6. Quy trình Dựng trước 1 tuần", "Sau khi quay xong vào Thứ Bảy: Hậu kỳ dựng thô Chủ Nhật, hoàn thiện đồ họa/screencast Thứ Ba, duyệt bản cuối Thứ Năm, sẵn sàng lịch đăng Thứ Hai tuần kế tiếp.", "Video hoàn thành trước ngày phát sóng ít nhất 3 ngày để kiểm duyệt kỹ lưỡng.", "Video Editor (Remotion/CapCut)", "Luôn có bản master xuất chuẩn 4K hoặc 1080p sắc nét, âm lượng chuẩn YouTube -14 LUFS.")
    ]

    for row_idx, row_data in enumerate(checklist_data, 5):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws_cl.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_top_left

    col_widths_cl = [22, 45, 35, 25, 35]
    for i, w in enumerate(col_widths_cl, 1):
        ws_cl.column_dimensions[get_column_letter(i)].width = w

    # Save to tmp and atomic replace
    wb.save(TMP_FILE)
    if os.path.exists(TMP_FILE):
        os.replace(TMP_FILE, TARGET_FILE)
        print(f"Workbook successfully saved and atomic replaced at:\n{TARGET_FILE}")
    else:
        raise FileNotFoundError(f"Tmp file {TMP_FILE} was not created.")

if __name__ == "__main__":
    create_multisheet_workbook()
