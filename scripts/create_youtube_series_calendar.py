# -*- coding: utf-8 -*-
r"""
Tạo file Excel Quản lý Lịch phát sóng YouTube Series & ĐỦ 9 BẢN NHÁP KỊCH BẢN RIÊNG BIỆT
Lưu trữ tại: C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx
Đáp ứng chuẩn xác 100% yêu cầu của Chị Thúy:
1. ĐỦ 9 BẢN NHÁP KỊCH BẢN RIÊNG BIỆT (KB-01 đến KB-09).
2. Phân biệt rõ: Kịch bản khác, Tập khác. Một kịch bản có thể triển khai thành nhiều tập (Ví dụ: KB-08 là Ký sự dài gồm 3 tập liên hoàn có Cliffhanger nối tiếp).
3. Kịch bản 01 là Clip ĐỘC LẬP 100% GIỚI THIỆU BẢN THÂN & LÀM QUEN.
4. Kết video rút ra KẾT LUẬN & CHIÊM NGHIỆM TỰ THÂN CỦA THÚY CHO CHÍNH MÌNH (tuyệt đối không khuyên bảo/giảng giải người khác).
5. 1 tuần đăng 3 bài, dựng trước 1 tuần, mỗi lượt quay quay trọn vẹn 1 series / kịch bản nhiều tập.
6. Cập nhật dữ liệu từ 4 nguồn chuẩn xác (Văn bản luật + FB/Google, Sách chính thống/Giáo trình, Wikipedia tình tiết thật, Daily Note & FB Th.xuatnhapkhau).
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TARGET_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
TARGET_FILE = os.path.join(TARGET_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
TMP_FILE = os.path.join(TARGET_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.tmp.xlsx")

def create_calendar_workbook():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    NAVY_HEADER = "1B365D"
    GOLD_ACCENT = "D4AF37"
    LIGHT_BLUE = "E8EEF5"
    ZEBRA_FILL = "F9FAFC"
    SPECIAL_FILL = "FFFDF0"
    WHITE = "FFFFFF"
    BORDER_COLOR = "D0D7DE"

    font_title = Font(name="Arial", size=14, bold=True, color="1B365D")
    font_subtitle = Font(name="Arial", size=10, italic=True, color="555555")
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

    # ==========================================
    # SHEET 1: TỔNG QUAN LỊCH PHÁT SÓNG & QUY TRÌNH
    # ==========================================
    ws1 = wb.create_sheet(title="LỊCH PHÁT SÓNG & SERIES")
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:J1")
    ws1["A1"] = "KẾ HOẠCH PHÁT SÓNG YOUTUBE CHUẨN SERIES — KÊNH TÂM TÌNH THUYTALK"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws1.merge_cells("A2:J2")
    ws1["A2"] = "Chiến lược: 1 tuần đăng 3 bài (Thứ 2 - 4 - 6 lúc 20:00) | Dựng hoàn tất trước 1 tuần | Quay theo đợt nguyên kịch bản/series | Kết luận tự thân của Thúy"
    ws1["A2"].font = font_subtitle

    # Table 1: Quy chế 4 Chuyên mục & Nguồn thẩm định
    ws1["A4"] = "I. HỆ THỐNG 4 CHUYÊN MỤC TRỤ CỘT & NGUỒN DỮ LIỆU XÁC THỰC"
    ws1["A4"].font = font_bold
    
    headers_cm = ["STT", "Chuyên mục", "Tên định dạng", "Nguồn dữ liệu xác thực (Fact-Check)", "Mục tiêu truyền thông & Giọng điệu"]
    for col_idx, h in enumerate(headers_cm, 1):
        cell = ws1.cell(row=5, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    data_cm = [
        ("1", "Chuyên mục 1", "Trà Chiều & Văn Bản Nóng", "Văn bản pháp luật mới (Nghị định, Thông tư, Công văn) + Cào Facebook/Google các điểm nóng dư luận đang tranh cãi", "Bắt nhịp thời sự, bóc tách thực chất quy định, chia sẻ góc nhìn điềm đạm, không đao to búa lớn."),
        ("2", "Chuyên mục 2", "Dịch Luật Ra Tiếng Người", "Văn bản pháp luật chính thống (Luật HQ 2014, Luật QLT 38), giáo trình ngoại thương, sách chuyên khảo chuẩn", "Bình dân hóa khái niệm khô khan (Hậu kiểm, Incoterms, Trốn thuế vs Tối ưu) qua ví dụ đời thường sinh động."),
        ("3", "Chuyên mục 3", "Ký Sự Lịch Sử Thương Mại", "Wikipedia, website chính thống, sách sử chính thống (The Box - Marc Levinson); 100% tình tiết thật, có dẫn chiếu nguồn", "Được tổ chức thành SERIES NHIỀU TẬP LIÊN HOÀN có Cliffhanger để người xem bấm xem tập tiếp theo."),
        ("4", "Chuyên mục 4", "Tâm Tình Thực Chiến Nghề XNK", "Daily Note thực tế công việc hàng ngày + Bài đăng trang cá nhân https://www.facebook.com/Th.xuatnhapkhau/", "Chia sẻ trải nghiệm va vấp thật; kết bài luôn là bài học THÚY TỰ RÚT RA CHO BẢN THÂN MÌNH.")
    ]

    for row_idx, row_data in enumerate(data_cm, 6):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_center if col_idx == 1 else align_left

    # Table 2: Danh mục 9 Kịch bản độc lập để Chị Thúy xem xét
    start_row_kb = 12
    ws1.cell(row=start_row_kb, column=1, value="II. DANH MỤC 9 BẢN NHÁP KỊCH BẢN ĐỘC LẬP (MỖI ĐỢT QUAY SẼ QUAY HẾT 1 KỊCH BẢN/SERIES)").font = font_bold

    headers_kb_overview = ["Mã KB", "Tên Kịch Bản", "Số Tập Trong Kịch Bản", "Chuyên Mục", "Nguồn Dữ Liệu", "Thời Lượng", "Đặc Điểm Kết Nối / Cliffhanger", "Tuần Quay (Lead)", "Tuần Phát Sóng", "Trạng Thái"]
    for col_idx, h in enumerate(headers_kb_overview, 1):
        cell = ws1.cell(row=start_row_kb+1, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    data_kb_overview = [
        ("KB-01", "Chào sân: Người đàn bà sau những container & tờ khai hải quan", "1 Tập", "Chào Sân / Định Vị", "Trải nghiệm cá nhân 10 năm nghề XNK Eureka", "6 - 7 phút", "Tập độc lập mở màn làm quen, giới thiệu bản thân", "Tuần W42", "Tuần W43", "Sẵn sàng duyệt"),
        ("KB-02", "Bão rà soát đóng mã số thuế phụ 888 & Nợ thuế eTax Mobile", "1 Tập", "CM1: Văn Bản Nóng", "Nghị định 125/2020 & Nhóm FB thuế", "7 - 8 phút", "Bắt nhịp điểm nóng thuế cá nhân 4.0", "Tuần W42", "Tuần W43", "Sẵn sàng duyệt"),
        ("KB-03", "Công văn 22883 Cục Hải quan: Hamburger kẹp thịt là bánh mì hay thịt?", "1 Tập", "CM1: Văn Bản Nóng", "Công văn 22883/CHQ-NVTHQ 12/10/2026 & Daily Note", "6 - 8 phút", "Ranh giới 20% thịt kẹp phân định Chương 16 vs 19", "Tuần W42", "Tuần W43", "Sẵn sàng duyệt"),
        ("KB-04", "Sàn TMĐT khấu trừ thuế thay: Hồi kết của thời bán hàng online vô hình", "1 Tập", "CM1: Văn Bản Nóng", "Nghị định TMĐT mới & Big Data ngành thuế", "7 - 8 phút", "Đàm đạo dòng tiền bán hàng online và truy thu", "Tuần W43", "Tuần W44", "Sẵn sàng duyệt"),
        ("KB-05", "Kiểm tra sau thông quan: Giải mã 5 năm 'hồi tố' & Quản trị kho", "1 Tập", "CM2: Dịch Luật", "Luật Hải quan 2014 Điều 77-82 & Giáo trình", "8 - 10 phút", "Dịch thuật ngữ hậu kiểm ra đời sống thực tế", "Tuần W43", "Tuần W44", "Sẵn sàng duyệt"),
        ("KB-06", "Incoterms 2020: Vì sao 'Mua CIF Bán FOB' khiến DN Việt mất quyền tự quyết?", "1 Tập", "CM2: Dịch Luật", "Incoterms 2020 ICC & Báo cáo logistics", "7 - 8 phút", "Phân tích vì sao bị ép phụ phí local charges", "Tuần W43", "Tuần W44", "Sẵn sàng duyệt"),
        ("KB-07", "Ranh giới giữa 'Trốn thuế' và 'Tối ưu chi phí thuế' trước pháp luật", "1 Tập", "CM2: Dịch Luật", "Luật QLT 38 Điều 143 & BLHS 2015 Điều 200", "8 - 9 phút", "Ví von vượt đèn đỏ vs đi làn đường cao tốc", "Tuần W44", "Tuần W45", "Sẵn sàng duyệt"),
        ("KB-08", "Đại Ký Sự Chiếc Thùng Container: Cuộc Cách Mạng Thu Nhỏ Hành Tinh", "3 Tập Liên Hoàn", "CM3: Lịch Sử (Multi-part)", "Sách The Box (Marc Levinson) & ISO 668", "25 - 30 phút (3 tập)", "Có Cliffhanger móc câu hồi hộp dẫn sang tập sau!", "Tuần W44", "Tuần W45", "Sẵn sàng duyệt"),
        ("KB-09", "Gãy seal chuyển cảng & Con ốc vít HS Code: Bài học cứu sống lô hàng", "1 Tập", "CM4: Tâm Tình Thực Chiến", "Daily Note thực tế & FB Th.xuatnhapkhau", "8 - 9 phút", "Kinh nghiệm hiện trường giải cứu cont hàng", "Tuần W44", "Tuần W45", "Sẵn sàng duyệt")
    ]

    for row_idx, row_data in enumerate(data_kb_overview, start_row_kb+2):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws1.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_center if col_idx in [1, 3, 6, 8, 9, 10] else align_left

    col_widths_s1 = [12, 38, 22, 24, 35, 14, 38, 16, 16, 16]
    for i, w in enumerate(col_widths_s1, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # ==========================================
    # SHEET 2: ĐỦ 9 BẢN NHÁP KỊCH BẢN CHI TIẾT
    # ==========================================
    ws2 = wb.create_sheet(title="9 BẢN NHÁP KỊCH BẢN CHI TIẾT")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:M1")
    ws2["A1"] = "DANH SÁCH ĐỦ 9 BẢN NHÁP KỊCH BẢN YOUTUBE (KB-01 ĐẾN KB-09) — ĐỂ CHỊ THÚY XEM XÉT"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws2.merge_cells("A2:M2")
    ws2["A2"] = "Mỗi kịch bản là một đề tài độc lập (KB-08 gồm 3 tập liên hoàn có Cliffhanger) | Kết luận video 100% là chiêm nghiệm tự thân của Thúy"
    ws2["A2"].font = font_subtitle

    headers_drafts = [
        "Mã KB", "Tên Kịch Bản & Tiêu Đề Video", "Định Dạng & Số Tập", "Chuyên Mục",
        "Nguồn Dữ Liệu Xác Thực (Fact-Check)", "Thời Lượng Dự Kiến",
        "Hook Mở Đầu (0:00 - 0:45)", "Nội Dung Kịch Bản Lời Thoại Chi Tiết (Full Script)",
        "Chỉ Dẫn Visual & Đạo Cụ Góc Máy", "Móc Câu Cliffhanger (Nếu chia nhiều tập)",
        "Kết Luận & Chiêm Nghiệm Tự Thân Của Thúy", "Lịch Đăng Đề Xuất", "Trạng Thái Xem Xét"
    ]

    for col_idx, h in enumerate(headers_drafts, 1):
        cell = ws2.cell(row=4, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    # 9 Distinct Scripts
    scripts_data = [
        # --- KB-01: CLIP GIỚI THIỆU BẢN THÂN ---
        (
            "KB-01",
            "Chào bạn, mình là Thúy: Uống chén trà và nghe chuyện của người đàn bà sau những container hàng & tờ khai hải quan",
            "Độc lập (1 Tập)",
            "Chào Sân / Định Vị Kênh",
            "Trải nghiệm cá nhân thực tế hơn 10 năm điều hành doanh nghiệp XNK Eureka; Ghi chép nhật ký hành trình làm nghề và tư vấn pháp lý thương mại.",
            "6 - 7 phút",
            "Chào bạn. Có lẽ khi nhắc đến những từ như 'Xuất nhập khẩu', 'Thuế', hay 'Thủ tục hải quan'... trong đầu nhiều người sẽ hiện lên những xấp hồ sơ dày cộp khô khốc, những con số nhức đầu, hoặc những cuộc kiểm tra căng thẳng. Nhưng sau nhiều năm lăn lộn trên thương trường, trực tiếp đứng sau hàng ngàn chuyến hàng vượt đại dương... mình nhận ra đằng sau mỗi chiếc container, mỗi dòng mã HS, mỗi tờ khai thông quan... là mồ hôi, là bản lĩnh và những câu chuyện nhân sinh vô cùng sâu sắc.",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, mình là Thúy. Hôm nay là ngày đầu tiên góc nhỏ này mở cửa, Thúy pha một ấm trà ấm để ngồi xuống làm quen và tâm tình cùng các bạn.

Mình là một người phụ nữ làm kinh doanh thực chiến trong lĩnh vực thương mại quốc tế, logistics và giải quyết các bài toán thủ tục hải quan, thuế chuyên sâu. Hơn một thập kỷ qua, cuộc sống của mình gắn liền với những cảng biển, những kho bãi, những cuộc đàm phán cước tàu và những lần trực tiếp giải cứu những lô hàng bị tắc nghẽn nơi cửa khẩu.

Mình từng chứng kiến những giọt nước mắt bất lực của những chủ doanh nghiệp trẻ khi hàng bị giữ tại cảng, bị phạt truy thu hàng trăm triệu chỉ vì một con số áp sai mã HS. Mình cũng từng chứng kiến những người bạn làm ăn thân thiết, chỉ vì cả nể, vì nghĩ 'anh em với nhau cần gì giấy tờ' mà khi sự cố xảy ra thì đánh mất cả cơ nghiệp lẫn tình nghĩa.

Góc nhỏ này ra đời không phải để Thúy lên lớp giảng giải câu chữ luật khô khan. Mình mở kênh này đơn giản như một góc bàn trà sau một ngày làm việc mỏi mệt. Nơi Thúy muốn chia sẻ lại những kinh nghiệm thực tế, những nghiên cứu cá nhân và những góc nhìn điềm đạm về:
- Các văn bản pháp luật, chính sách thuế, hải quan mới ban hành đang tác động trực tiếp đến túi tiền của chúng ta.
- 'Dịch' những khái niệm luật pháp khô khan, hàn lâm thành những câu chuyện đời thường dễ hiểu nhất.
- Và kể cho các bạn nghe những câu chuyện lịch sử thương mại có thật hàng ngàn năm trước của thế giới và của cha ông ta...

Tất cả đều được nhìn qua lăng kính của sự điềm đạm, khiêm tốn và thấu hiểu.""",
            "Quay trực diện bàn trà gỗ mộc, ánh sáng vàng ấm 3500K. Trên bàn có tách trà bốc khói nhẹ, mô hình container nhỏ xíu bằng kim loại và cặp sổ da mộc. Góc máy chính diện Eye-level gần gũi, chuyển góc nghiêng 45 độ khi chia sẻ những trăn trở sâu lắng.",
            "Không áp dụng (Tập độc lập chào sân, mở ra lời hẹn đồng hành cho các số tiếp theo).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Với riêng bản thân Thúy, sau ngần ấy năm đi qua đủ thăng trầm của nghề, bài học lớn nhất mà mình tự rút ra cho chính mình là: 'Hiểu luật không phải để trở thành một người thích tranh cãi hơn thua, mà hiểu luật là để tự bảo vệ sự bình yên trong tâm trí'. Khi mình làm đúng từ đầu, hồ sơ ngay ngắn, thì dù bão giông ngoài thương trường lớn đến đâu, đêm về mình vẫn ngủ trọn một giấc ngon.

Cảm ơn bạn đã ghé thăm góc nhỏ của Thúy ngày hôm nay. Nếu bạn thấy đồng điệu, hãy để lại một lời chào dưới bình luận để chúng mình làm quen nhé. Hẹn gặp lại bạn ở những buổi trà chiều tiếp theo!" """,
            "Thứ 2, 19/10/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-02: THỜI SỰ VĂN BẢN NÓNG - MST 888 & ETAX MOBILE ---
        (
            "KB-02",
            "Đàm đạo bão rà soát đóng mã số thuế phụ 888 & Nợ thuế eTax Mobile: Khi sự vô tư phải trả giá bằng tiền triệu",
            "Độc lập (1 Tập)",
            "CM1: Trà Chiều & Văn Bản Nóng",
            "Nghị định 125/2020/NĐ-CP (Điều 13 mức phạt chậm nộp hồ sơ khai thuế); Thông tư 40/2021/TT-BTC; Dữ liệu cào nhóm Hỏi đáp Thuế FB & phản ánh thực tế đóng MST phụ 888.",
            "7 - 8 phút",
            "Chào bạn. Nếu bạn đang làm tiếp thị liên kết affiliate trên Shopee, TikTok Shop, hoặc đang mở mã số thuế phụ đuôi 888 để nhận hoa hồng dịch vụ... mà thời gian này mở eTax Mobile lên bỗng thấy báo nợ thuế, hoặc đi xin đóng mã số thuế mà nhận thông báo phạt chậm nộp hàng chục triệu... thì hôm nay Thúy ngồi lại chia sẻ góc nhìn thực tế nhé.",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, Thúy đây.

Cách đây vài ngày, một bạn chuyên viên làm affiliate có nhắn tin chia sẻ với Thúy trong trạng thái khá hoang mang. Bạn ấy nộp hồ sơ xin đóng mã số thuế phụ đuôi 888. Năm ngoái cả năm bạn chỉ có doanh thu hoa hồng đúng 1 triệu đồng và sàn TMĐT đã khấu trừ 7% thuế nộp vào kho bạc. Thế nhưng khi lên chi cục thuế, cán bộ thuế in ra thông báo phạt chậm nộp 4 quý tờ khai trắng với số tiền dự kiến gần 24 triệu đồng theo Nghị định 125/2020!

Nhiều người ngây thơ nghĩ rằng: 'Sàn đã trừ thuế của mình rồi, mình đâu nợ thuế nhà nước, sao lại phạt?'

Ở góc độ nghiên cứu của Thúy:
Theo Luật Quản lý thuế số 38, nghĩa vụ nộp hồ sơ khai thuế và nghĩa vụ nộp tiền thuế là hai việc hoàn toàn độc lập. Mã 888 là mã cá nhân kinh doanh, theo quý bắt buộc phải nộp tờ khai định kỳ dù doanh thu bằng 0. Nghị định 125 phạt hành vi chậm nộp thời hạn hồ sơ, chứ không phải phạt trốn thuế.

Khi gặp tình huống này, cách Thúy hướng dẫn bạn xử lý bình tĩnh:
- Bước 1: Tra cứu đúng các quý còn thiếu trên eTax Mobile.
- Bước 2: Nộp bổ sung tờ khai Mẫu 01/TKN-CNKD, kê đúng số thuế sàn đã khấu trừ để không bị tính trùng nợ.
- Bước 3: Làm văn bản giải trình tình tiết giảm nhẹ theo Điều 9 Nghị định 125/2020 (vi phạm lần đầu, không cố ý, số thuế phát sinh nhỏ và đã khấu trừ tại nguồn) để xin áp dụng mức phạt tối thiểu.""",
            "Quay Chị Thúy ngồi bên bàn trà, chèn màn hình screencast app eTax Mobile (che thông tin cá nhân). Bảng đồ họa trích dẫn các khung phạt Điều 13 Nghị định 125/2020.",
            "Không áp dụng (Tập đơn, có lời dẫn mở kêu gọi thảo luận phía dưới).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Sau trường hợp này, bài học mà bản thân Thúy tự rút ra cho chính mình là: 'Trong thời đại dữ liệu số liên thông, sự cẩn trọng với các thủ tục hành chính không bao giờ là thừa'. Bản thân mình luôn tự rà soát định kỳ các nghĩa vụ thuế cá nhân và doanh nghiệp, không bao giờ để một mã số thuế nào mở ra rồi bỏ quên. Chuẩn chỉ sổ sách là chiếc áo giáp an toàn nhất cho chính mình." """,
            "Thứ 4, 21/10/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-03: THỜI SỰ VĂN BẢN NÓNG - CÔNG VĂN 22883 HAMBURGER ---
        (
            "KB-03",
            "Công văn 22883 Cục Hải quan: Hamburger kẹp thịt là bánh mì hay thịt? Ranh giới 20% và bài học phân loại hàng hóa",
            "Độc lập (1 Tập)",
            "CM1: Trà Chiều & Văn Bản Nóng",
            "Công văn số 22883/CHQ-NVTHQ ngày 12/10/2026 của Cục Hải quan về rà soát phân loại mặt hàng Hamburger; Chú giải chi tiết HS Chương 16 (nhóm 16.02) và Chương 19 (nhóm 19.05); Dữ liệu Daily Note ngày 13/10/2026.",
            "6 - 8 phút",
            "Chào bạn. Có bao giờ bạn nghĩ một chiếc bánh Hamburger kẹp thịt thơm ngon chúng ta ăn mỗi ngày, khi đi qua cửa khẩu hải quan lại có thể khiến các doanh nghiệp đau đầu đến mức Cục Hải quan phải ra hẳn một công văn khẩn để chấn chỉnh không? Hamburger rốt cuộc là 'bánh mì' hay là 'thịt'? Hôm nay Thúy kể cho bạn nghe một câu chuyện nghiệp vụ rất thú vị nhé.",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào anh chị em, Thúy đây. 

Mấy ngày hôm nay, cộng đồng xuất nhập khẩu thực phẩm đang bàn tán xôn xao về Công văn số 22883 vừa được Cục Hải quan ban hành ngày 12/10/2026. Nội dung là yêu cầu các Chi cục hải quan rà soát khẩn toàn bộ các lô hàng khai báo là Hamburger để xác định lại mã HS và chính sách thuế.

Chuyện là thế này: Rất nhiều doanh nghiệp trước giờ nhập bánh Hamburger kẹp thịt bò, hamburger xúc xích... thường hồn nhiên khai báo vào nhóm 19.05 – tức là nhóm bánh mì, bột nhào nướng, với mức thuế suất và thủ tục kiểm tra chuyên ngành khá đơn giản.
Nhưng Hải quan cho rằng: Hamburger có phần thịt kẹp bên trong. Nếu tỷ lệ thịt lớn mà vẫn áp vào bánh mì thì chưa chuẩn xác!

Và đây là con số định đoạt mà Thúy muốn chia sẻ: RANH GIỚI 20%.
Theo Chú giải pháp lý HS:
1. Nếu tổng khối lượng của xúc xích, thịt, phụ phẩm thịt chiếm TRÊN 20% khối lượng toàn bộ chiếc bánh: Sản phẩm lập tức bị kéo sang Chương 16 (nhóm 16.02 - chế phẩm từ thịt). Khi sang Chương 16, không chỉ thuế suất thay đổi, mà quy trình kiểm dịch động vật và an toàn thực phẩm ngặt nghèo hơn gấp nhiều lần!
2. Chỉ khi tỷ lệ thịt kẹp KHÔNG VƯỢT QUÁ 20% và bản chất vẫn là sản phẩm bánh mì nướng: Bánh mới được giữ lại ở nhóm 19.05 của Chương 19.

Cục Hải quan đã yêu cầu các Chi cục kiểm tra hồ sơ kỹ thuật, nếu không đủ cơ sở thì lấy mẫu đi phân tích định lượng thành phần.""",
            "Quay Chị Thúy cầm tách trà, bên cạnh có đĩa bánh hamburger mini mô phỏng. Slide đồ họa hiện infographic so sánh: Tỷ lệ thịt >20% (Chương 16) vs <=20% (Chương 19). Trích dẫn văn bản CV 22883/CHQ-NVTHQ trên màn hình.",
            "Không áp dụng (Tập đơn, có câu hỏi mở cuối video về kinh nghiệm phân loại hàng thực phẩm).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ câu chuyện chiếc bánh hamburger này, bài học mà bản thân Thúy tự nhắc nhở mình mỗi ngày khi làm nghề là: 'Trong ngoại thương, ranh giới giữa an toàn và rủi ro đôi khi chỉ là 1% định lượng kỹ thuật'. Bản thân mình khi tư vấn hay làm việc với đối tác ngoại, luôn tự dặn lòng phải yêu cầu họ cung cấp Bản phân tích thành phần (C/A) chi tiết đến từng gram trước khi hàng lên tàu, để chính mình không bao giờ rơi vào thế bị động." """,
            "Thứ 6, 23/10/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-04: THỜI SỰ VĂN BẢN NÓNG - SÀN TMĐT KHẤU TRỪ THUẾ THAY ---
        (
            "KB-04",
            "Sàn TMĐT khấu trừ thuế thay & Định danh tài khoản ngân hàng: Hồi kết của thời bán hàng online 'vô hình'",
            "Độc lập (1 Tập)",
            "CM1: Trà Chiều & Văn Bản Nóng",
            "Nghị định quy định trách nhiệm quản lý thuế đối với sàn giao dịch TMĐT; Thông tư hướng dẫn kê khai khấu trừ thuế tại nguồn; Dữ liệu đối soát tài khoản ngân hàng theo Đề án 06.",
            "7 - 8 phút",
            "Chào bạn. Có bao giờ bạn tự hỏi: Tại sao những năm trước bán hàng online trăm đơn ngàn đơn một ngày chẳng thấy ai hỏi han, nhưng bước sang năm 2026, cơ quan thuế lại nắm rõ từng giao dịch, từng đồng tiền về tài khoản của bạn như trong lòng bàn tay? Thời của việc kinh doanh online 'vô hình' đã chính thức khép lại. Hôm nay Thúy ngồi lại đàm đạo về chính sách khấu trừ thuế sàn TMĐT nhé.",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, Thúy đây.

Những quy định mới nhất yêu cầu các sàn thương mại điện tử lớn như Shopee, TikTok Shop, Lazada... phải có trách nhiệm khấu trừ, kê khai và nộp thuế thay cho các nhà bán hàng cá nhân, kết hợp với việc định danh tài khoản ngân hàng theo mã CCCD đã tạo nên một bước ngoặt chưa từng có trong giới kinh doanh số.

Trước đây, rất nhiều nhà bán hàng chọn cách 'né' bằng cách mở nhiều tài khoản cá nhân, chia nhỏ dòng tiền sang người thân. Nhưng hiện nay, cơ sở dữ liệu quốc gia về dân cư liên thông trực tiếp với hệ thống của Tổng cục Thuế và Ngân hàng Nhà nước qua hệ thống Big Data. Mọi dòng tiền chuyển khoản có dấu hiệu kinh doanh đều được thuật toán tự động nhận diện và phân tích.

Khi các sàn tự động khấu trừ thuế VAT và TNCN ngay khi đơn hàng hoàn tất, nhiều người lo lắng: 'Liệu mình có bị khấu trừ trùng không? Nếu bị trả hàng hoàn tiền thì thuế xử lý ra sao?'

Góc nhìn của Thúy:
Cơ chế khấu trừ tại nguồn thực ra lại là một chiếc van an toàn cho chính người bán. Thay vì để 3–4 năm sau bị cơ quan thuế thanh tra và truy thu kèm tiền phạt chậm nộp hàng trăm triệu, việc sàn khấu trừ ngay từng đơn giúp dòng tiền thuế được làm sạch tức thì.
Điều người kinh doanh cần làm ngay lúc này:
1. Rà soát lại giá bán và biên lợi nhuận: Phải tính thuế 1.5% - 7% trực tiếp vào cơ cấu giá vốn từ đầu.
2. Lưu giữ đầy đủ hóa đơn chứng từ đầu vào của nguồn hàng để sau này có căn cứ chuyển đổi lên doanh nghiệp khi quy mô lớn dần.""",
            "Quay Chị Thúy bên bàn trà, màn hình chèn sơ đồ luồng tiền: Khách hàng -> Sàn TMĐT -> Khấu trừ thuế trực tiếp vào Kho bạc -> Trả tiền ròng cho Người bán. Đồ họa liên thông Đề án 06.",
            "Không áp dụng (Tập đơn).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ sự chuyển dịch này của nền kinh tế số, bài học mà bản thân Thúy luôn tự dặn mình trong quản trị là: 'Đừng bao giờ xây dựng mô hình kinh doanh dựa trên việc trốn tránh một nghĩa vụ tất yếu'. Bản thân mình nhận thấy: Minh bạch ngay từ 1 đồng doanh thu đầu tiên tuy có thể làm mình cảm thấy biên lợi nhuận mỏng hơn một chút, nhưng đổi lại là một cơ nghiệp bền vững, đàng hoàng và ngủ ngon giấc mỗi đêm." """,
            "Thứ 2, 26/10/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-05: DỊCH LUẬT RA TIẾNG NGƯỜI - HẬU KIỂM 5 NĂM ---
        (
            "KB-05",
            "Kiểm tra sau thông quan: Giải mã 5 năm 'hồi tố' theo Luật Hải quan 2014 & Cách kiểm soát định mức kho",
            "Độc lập (1 Tập)",
            "CM2: Dịch Luật Ra Tiếng Người",
            "Luật Hải quan số 54/2014/QH13 (Điều 77 đến Điều 82); Thông tư 38/2015/TT-BTC sửa đổi bổ sung bởi Thông tư 39/2018/TT-BTC; Giáo trình Thủ tục Hải quan Đại học Tài chính - Marketing.",
            "8 - 10 phút",
            "Chào bạn. Trong ngành xuất nhập khẩu, người ta hay nói đùa một câu: 'Hàng ra khỏi cảng chưa chắc đã là của mình'. Nhiều anh chị thấy tờ khai luồng Xanh, hàng đã về kho đem bán hết lấy tiền về rồi thì đinh ninh là xong xuôi. Nhưng 3 năm, thậm chí 5 năm sau, bỗng nhận quyết định Kiểm tra sau thông quan... Tại sao lại có cơ chế 5 năm này? Hôm nay Thúy cùng bạn bóc tách nhé.",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, Thúy đây. Hôm nay mình lại ngồi uống trà và nói chuyện nghề.

Khi nhắc đến 4 chữ 'Kiểm tra sau thông quan', rất nhiều chủ doanh nghiệp cảm thấy e ngại. Nhưng dưới góc nhìn nghiên cứu quản trị hải quan hiện đại, Thúy muốn chia sẻ để mọi người hiểu đúng bản chất thế này:

Ngày xưa, hải quan các nước áp dụng cơ chế 'Tiền kiểm' — tức là hàng hóa đến cảng phải kiểm tra kỹ lưỡng từng kiện hàng, từng dòng chứng từ rồi mới cho đi. Việc này khiến cảng biển tắc nghẽn nghiêm trọng, doanh nghiệp đọng vốn.
Vì vậy, theo chuẩn mực của Tổ chức Hải quan Thế giới (WCO) và Luật Hải quan Việt Nam 2014, chúng ta chuyển sang cơ chế 'Hậu kiểm': Cho phép doanh nghiệp tự khai tự chịu trách nhiệm, thông quan thật nhanh tại cửa khẩu để đưa hàng vào sản xuất kinh doanh. Đổi lại, trong thời hạn 5 năm kể từ ngày đăng ký tờ khai, cơ quan hải quan có quyền kiểm tra lại toàn bộ hồ sơ, sổ sách kế toán, chứng từ thanh toán và tồn kho thực tế của doanh nghiệp.

Kiểm tra sau thông quan gồm 2 hình thức:
1. Kiểm tra tại trụ sở cơ quan hải quan: Cơ quan hải quan yêu cầu gửi hồ sơ lên để đối chiếu.
2. Kiểm tra tại trụ sở doanh nghiệp: Đoàn kiểm tra đến trực tiếp nhà máy, kho bãi để rà soát hệ thống kế toán và dây chuyền sản xuất.

3 điểm 'tử huyệt' mà các đoàn hậu kiểm soi kỹ nhất:
- Mã HS và thuế suất ưu đãi C/O.
- Trị giá hải quan (nghi vấn khai báo giá thấp để né thuế).
- Báo cáo quyết toán và định mức tiêu hao nguyên phụ liệu (đặc biệt đối với hàng gia công, sản xuất xuất khẩu). Khi số liệu tồn kho thực tế lệch với số liệu sổ sách kế toán, doanh nghiệp sẽ bị ấn định thuế và phạt tiền rất nặng!""",
            "Quay Chị Thúy ngồi trước giá sách tài liệu. Đồ họa màn hình: So sánh mô hình Tiền kiểm vs Hậu kiểm. Highlight các mốc thời hạn 5 năm theo Điều 78 Luật Hải quan 2014.",
            "Không áp dụng (Tập đơn).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Với riêng bản thân Thúy, sau nhiều lần đồng hành cùng các đợt kiểm tra sau thông quan, bài học mà mình tự rèn luyện cho chính doanh nghiệp của mình là: 'Không bao giờ xem việc thông quan tại cảng là điểm kết thúc'. Bản thân mình luôn duy trì thói quen lưu trữ hồ sơ chứng từ một cách ngăn nắp, đồng bộ số liệu kho với kế toán từng tháng, để lúc nào mình cũng ở tâm thế sẵn sàng, đàng hoàng và tự tin." """,
            "Thứ 4, 28/10/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-06: DỊCH LUẬT RA TIẾNG NGƯỜI - MUA CIF BÁN FOB ---
        (
            "KB-06",
            "Incoterms 2020: Vì sao 'Mua CIF Bán FOB' khiến doanh nghiệp Việt mất quyền tự quyết chi phí logistics?",
            "Độc lập (1 Tập)",
            "CM2: Dịch Luật Ra Tiếng Người",
            "Quy tắc Incoterms 2020 của Phòng Thương mại Quốc tế (ICC); Báo cáo Logistics Việt Nam thường niên của Bộ Công Thương; Phân tích thực tế nghiệp vụ cước biển.",
            "7 - 8 phút",
            "Chào bạn. Suốt mấy chục năm qua, có một công thức dường như đã trở thành phản xạ ăn sâu vào tiềm thức của phần lớn doanh nghiệp XNK Việt Nam: Xuất khẩu thì bán FOB (giao tại mạn tàu ở cảng Việt Nam), còn Nhập khẩu thì mua CIF (để người bán nước ngoài lo trọn gói cước tàu về cảng Việt Nam). Người ta nghĩ thế là an nhàn, không phải đau đầu thuê tàu. Nhưng thực sự chúng ta đang được lợi hay đang tự trói tay mình?",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, Thúy đây. 

Trong giao thương quốc tế, các điều kiện Incoterms sinh ra để phân định rành mạch 3 thứ: Chi phí, Rủi ro, và Trách nhiệm chứng từ giữa người mua và người bán.

Tại sao doanh nghiệp Việt thích 'Mua CIF, Bán FOB'?
Lý do tâm lý rất đơn giản: Ngại va chạm, sợ rủi ro biến động giá cước biển, không có kinh nghiệm đàm phán với các hãng tàu quốc tế. Khi bán FOB, giao hàng lên tàu ở cảng Cát Lái hay Hải Phòng là thở phào. Khi mua CIF, cứ ngồi ở nhà đợi hàng cập cảng rồi ra lấy.

Nhưng cái giá phải trả của sự 'an nhàn' này là gì?
Thứ nhất: Khi bạn mua CIF, quyền chỉ định hãng tàu thuộc về người bán nước ngoài. Họ thường chọn các hãng tàu có giá cước biển (Freight) cực rẻ, nhưng hãng tàu đó lại âm thầm 'chém' phụ phí cảng đến (Local Charges tại Việt Nam như phí D/O, CIC, vệ sinh cont) với mức giá cắt cổ. Lúc này hàng đã về cảng, bạn không thể không đóng tiền!
Thứ hai: Khi bán FOB, quyền book tàu thuộc về đối tác ngoại. Bạn mất hoàn toàn cơ hội kết hợp đàm phán cước khối lượng lớn (volume discount) và nhường toàn bộ miếng bánh dịch vụ logistics giá trị gia tăng cho các tập đoàn vận tải ngoại.

Góc nhìn cá nhân của Thúy: Đã đến lúc chúng ta cần tự tin chuyển dịch sang 'Mua FOB, Bán CIF' hoặc CFR để nắm quyền chủ động vận tải.""",
            "Quay Chị Thúy bên tách trà, trên bàn có 2 mô hình tàu container mini gắn nhãn FOB và CIF. Infographic so sánh dòng dịch chuyển chi phí và rủi ro giữa 2 điều kiện giao hàng Incoterms 2020.",
            "Không áp dụng (Tập đơn).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ trải nghiệm thực tế bị các hãng tàu ngoại ép phụ phí địa phương trong những năm đầu khởi nghiệp, bài học mà bản thân Thúy tự rút ra là: 'Chủ động nắm quyền book tàu chính là chìa khóa để mình làm chủ chi phí và tiến độ lô hàng'. Với các đơn hàng của Eureka hiện nay, bản thân mình luôn ưu tiên đàm phán giá FOB khi nhập khẩu để chỉ định đối tác vận tải tin cậy tại Việt Nam." """,
            "Thứ 6, 30/10/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-07: DỊCH LUẬT RA TIẾNG NGƯỜI - TRỐN THUẾ VS TỐI ƯU THUẾ ---
        (
            "KB-07",
            "Ranh giới giữa 'Trốn thuế' và 'Tối ưu chi phí thuế': Điều gì bảo vệ người làm kinh doanh trước pháp luật?",
            "Độc lập (1 Tập)",
            "CM2: Dịch Luật Ra Tiếng Người",
            "Luật Quản lý thuế số 38/2019/QH14 (Điều 143 về Hành vi trốn thuế); Bộ luật Hình sự 2015 sửa đổi 2017 (Điều 200 về Tội trốn thuế); Chuẩn mực Kế toán Việt Nam (VAS).",
            "8 - 9 phút",
            "Chào bạn. Có một ranh giới trong kinh doanh mỏng manh như sợi tóc, nhưng bước qua một bên là sự thông minh tài chính, còn trượt sang bên kia lại là vòng lao lý. Đó chính là ranh giới giữa 'Tối ưu thuế' (Tax Planning) và 'Trốn thuế' (Tax Evasion). Làm sao để người làm kinh doanh vừa tiết kiệm được chi phí hợp pháp, vừa đêm về ngủ ngon giấc?",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, Thúy đây. Hôm nay chúng ta lại ngồi đàm đạo về một chủ đề vừa nhạy cảm nhưng lại vô cùng thiết thực.

Trong đời sống kinh doanh, có 3 khái niệm thường bị đánh đồng với nhau:
1. Trốn thuế (Tax Evasion): Đây là hành vi vi phạm pháp luật cố ý. Ví dụ: Dùng 2 hệ thống sổ sách kế toán, lập hóa đơn khống, mua bán hóa đơn của công ty ma, giấu doanh thu ngoài sổ sách, hoặc cố tình khai sai mã HS để trốn thuế nhập khẩu. Theo Điều 200 Bộ luật Hình sự, trốn thuế từ 100 triệu đồng trở lên đã có thể bị truy cứu trách nhiệm hình sự.
2. Tránh thuế (Tax Avoidance): Là việc doanh nghiệp lợi dụng các kẽ hở của luật pháp để giảm số thuế phải nộp. Dù chưa phạm luật hình sự, nhưng hành vi này thường bị cơ quan thuế bác bỏ và xếp vào diện điều tra chống chuyển giá.
3. Tối ưu chi phí thuế (Tax Planning): Đây là nghệ thuật quản trị hoàn toàn hợp pháp và được nhà nước khuyến khích! Đó là việc doanh nghiệp nghiên cứu kỹ các chính sách ưu đãi của pháp luật để vận dụng đúng:
- Chọn địa bàn đầu tư có thuế suất ưu đãi (khu công nghiệp, khu kinh tế khó khăn).
- Tận dụng các Hiệp định FTA (EVFTA, CPTPP) để xin cấp C/O hưởng thuế nhập khẩu 0%.
- Trích lập quỹ phát triển khoa học công nghệ, hạch toán đúng và đủ các chi phí hợp lý được trừ theo Luật Thuế TNDN.

Thúy hay ví von thế này cho dễ hiểu:
- Trốn thuế giống như bạn lái xe vượt đèn đỏ hoặc đi vào đường cấm.
- Tối ưu thuế là bạn nhìn biển báo giao thông thật kỹ, chọn đi vào làn đường cao tốc mà nhà nước cho phép lưu thông để về đích nhanh nhất và an toàn nhất.""",
            "Quay Chị Thúy với ánh mắt điềm tĩnh, nụ cười nhẹ nhàng. Màn hình chia 3 cột đồ họa: Trốn thuế (Màu đỏ - Cấm) | Tránh thuế (Màu vàng - Rủi ro) | Tối ưu thuế (Màu xanh - Hợp pháp & Khuyến khích).",
            "Không áp dụng (Tập đơn).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Chiêm nghiệm lớn nhất mà bản thân Thúy đúc kết cho chính mình qua bao năm làm doanh nghiệp: 'Đồng tiền giữ lại được từ sự hiểu biết và tuân thủ pháp luật là đồng tiền bình an nhất'. Bản thân mình không bao giờ tìm cách đi đường tắt, bởi vì giấc ngủ ngon và sự tự tin ngẩng cao đầu khi tiếp đón bất kỳ đoàn thanh tra nào mới là tài sản vô giá của người làm kinh doanh." """,
            "Thứ 2, 02/11/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-08: ĐẠI KÝ SỰ CHIẾC THÙNG CONTAINER (3 TẬP LIÊN HOÀN CÓ CLIFFHANGER) ---
        (
            "KB-08",
            "Đại Ký Sự Chiếc Thùng Container: Từ ý tưởng của anh tài xế xe tải nghèo đến cuộc cách mạng thu nhỏ hành tinh",
            "Ký Sự Dài (Gồm 3 Tập Liên Hoàn: E01, E02, E03)",
            "CM3: Ký Sự Lịch Sử Thương Mại",
            "Sách 'The Box: How the Shipping Container Made the World Smaller and the World Economy Bigger' của Marc Levinson (NXB Đại học Princeton); Wikipedia: Malcom McLean; Chuẩn ISO 668:1968; Dữ liệu lịch sử cảng Cam Ranh 1966.",
            "25 - 30 phút (Chia làm 3 Tập: Mỗi tập 8 - 10 phút)",
            "[Hook Tập 1]: Chào bạn. Chiếc áo bạn mặc, điện thoại bạn cầm... gần như chắc chắn từng nằm trong một chiếc hộp thép hình chữ nhật vượt đại dương. Bạn có biết: Trước 1956, chiếc hộp này hoàn toàn không tồn tại? Và người phát minh ra nó là một anh tài xế xe tải nghèo? Hôm nay Thúy mở màn Phần 1 của Đại Ký Sự lịch sử thương mại thế giới.",
            r"""[TOÀN VĂN KỊCH BẢN 3 TẬP LIÊN HOÀN - CHỊ THÚY]

=== TẬP 1: Ý TƯỞNG NGỚ NGẨN & CHUYẾN TÀU BÍ MẬT IDEAL X 1956 ===
Ngày 26/4/1956 tại cảng Newark, New Jersey, con tàu chở dầu Ideal X rời bến chở 58 thùng container đầu tiên. Malcom McLean – anh tài xế xe tải nghèo năm 1934 – khi nhìn cảnh phu bốc từng bao bông sợi chậm chạp đã tự hỏi: 'Tại sao không bốc cả thùng xe tải lên tàu?'. Ông bán công ty xe tải số 2 nước Mỹ đánh cược tất cả. Kết quả chấn động: Chi phí bốc dỡ từ 5,86 USD/tấn rơi thẳng đứng xuống 0,16 USD/tấn (giảm 97%)!
-> Kết luận tự thân Thúy: 'Đổi mới vĩ đại bắt nguồn từ việc quan sát nỗi đau thực tế mỗi ngày'.

=== TẬP 2: CUỘC CHIẾN SINH TỬ VỚI 50.000 PHU BẾN CẢNG & VỤ KIỆN THẾ KỶ ===
Ngay sau Ideal X, Nghiệp đoàn ILA của Teddy Gleason đại diện 50.000 phu bến cảng đình công tê liệt cảng bờ Đông Mỹ vì sợ mất chén cơm; giới đường sắt kiện chống độc quyền. McLean đàm phán thành lập Quỹ bảo trợ thất nghiệp, trả phí trên mỗi tấn hàng bốc dỡ để phu bến cảng có trợ cấp và học nghề mới. Nhưng khi vừa yên ổn thì thảm họa kỹ thuật ập đến: Mỗi hãng làm một kích thước (17ft, 24ft, 35ft), cần cẩu cảng quốc tế tê liệt!
-> Kết luận tự thân Thúy: 'Khi đổi mới đụng chạm số đông, đối đầu triệt hạ không có lối thoát; chỉ có thấu cảm và chia sẻ lợi ích mới đi tiếp được'.

=== TẬP 3: QUYẾT ĐỊNH BUÔNG BỎ TWISTLOCK & CHUẨN ISO 1968 THAY ĐỔI NHÂN LOẠI ===
Trước nguy cơ đổ vỡ toàn cầu, McLean nắm giữ bằng sáng chế then chốt: Khóa góc Twistlock. Ông đã làm hành động vĩ đại: Tuyên bố từ bỏ độc quyền, trao tặng miễn phí bản vẽ cho quốc tế! ISO ban hành chuẩn ISO 668 năm 1968 (chuẩn 20ft và 40ft). Thập niên 1960, container đổ bộ vào cảng Cam Ranh giải tỏa tàu quân nhu từ 20 ngày xuống 40 giờ, mở màn chuỗi cung ứng châu Á bùng nổ thống trị thế giới cho đến nay.
-> Kết luận tự thân Thúy: 'Đỉnh cao chiến lược kinh doanh không phải là giữ chặt lợi ích độc quyền nhỏ bé, mà là dũng cảm buông bỏ để cùng tạo ra một chuẩn mực lớn cho cả ngành'.""",
            "Quay Chị Thúy ngồi bên mô hình container mini và ấm trà gốm. Chèn ảnh tư liệu lịch sử: Chân dung Malcom McLean, tàu Ideal X 1956, biểu tình nghiệp đoàn ILA 1958, bản vẽ kỹ thuật khóa Twistlock, ảnh cảng Cam Ranh 1966. Dẫn chiếu rõ sách 'The Box' của Marc Levinson.",
            r"""[MÓC CÂU CLIFFHANGER GIỮA CÁC TẬP]:
- Cuối Tập 1 dẫn sang Tập 2: 'Thành công vang dội của Ideal X ngay lập tức kéo theo CƠN THỊNH NỘ KHỦNG KHIẾP từ 50.000 phu bến cảng New York và các tập đoàn đường sắt kiện tụng đòi bóp chết container! McLean đã chiến đấu thế nào? Bấm ngay vào Tập 2 đang hiện trên màn hình để xem tiếp nhé!'
- Cuối Tập 2 dẫn sang Tập 3: 'Khi kích thước hỗn loạn khiến cần cẩu cảng quốc tế tê liệt, McLean đứng trước bờ vực phá sản. Lúc này ông đã đưa ra quyết định điên rồ nhất lịch sử: Buông bỏ bằng sáng chế độc quyền tỷ đô! Quyết định đó là gì? Hãy bấm vào Tập 3 kết thúc series để xem nhé!' """,
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Bài học lớn nhất, sâu sắc nhất mà bản thân Thúy khắc cốt ghi tâm cho chính doanh nghiệp của mình từ câu chuyện của Malcom McLean: 'Đôi khi, đỉnh cao của chiến lược kinh doanh không phải là cố gắng giữ chặt lợi ích độc quyền nhỏ bé của mình, mà là dũng cảm buông bỏ để cùng kiến tạo nên một chuẩn mực lớn cho toàn ngành'. Bản thân mình khi làm nghề thương mại, luôn tự dặn lòng phải hướng tới những giá trị chung bền vững như thế." """,
            "Thứ 4, 04/11 - Thứ 6, 06/11 - Thứ 2, 09/11/2026 (Phát 3 buổi liên tục)",
            "Chờ Chị Thúy duyệt"
        ),

        # --- KB-09: TÂM TÌNH THỰC CHIẾN - SỰ CỐ HIỆN TRƯỜNG & MÃ HS CODE ---
        (
            "KB-09",
            "Gãy seal chuyển cảng & Con ốc vít HS Code: Bài học đắt giá cứu sống doanh nghiệp nơi cửa khẩu",
            "Độc lập (1 Tập)",
            "CM4: Tâm Tình Thực Chiến Nghề XNK",
            "Thực tế hiện trường giám sát hải quan cảng Cát Lái / Hải Phòng; Quy trình chuyển cửa khẩu theo Thông tư 38/2015 Điều 51; Ghi chép thực chiến từ Daily Note và bài đăng trang cá nhân https://www.facebook.com/Th.xuatnhapkhau/.",
            "8 - 9 phút",
            "Chào bạn. Có bao giờ bạn rơi vào cảm giác: Xe container trị giá vài tỷ vừa về đến kho thì tài xế báo chì niêm phong (seal) của hãng tàu và hải quan bị gãy? Hoặc cùng là một con ốc vít bằng thép, nhưng nếu áp sai 1 mã HS thì thuế từ 0% nhảy lên 25%, bị truy thu tiền tỷ? Hôm nay Thúy tâm sự về những sự cố thực tế thót tim mà mình đã trực tiếp cùng khách hàng vượt qua.",
            r"""[LỜI THOẠI CHI TIẾT - CHỊ THÚY]

Chào mọi người, Thúy đây. Hôm nay mình lại ngồi đàm đạo về những câu chuyện hiện trường rất thực.

Chuyện thứ nhất: Gãy seal chuyển cảng.
Một khách hàng của Thúy chuyển cont hàng từ Hải Phòng về kho ngoại quan Hà Nội, khi đến nơi phát hiện seal thép bị đứt do va quẹt. Khách hàng sợ hãi định tự ý cắt bỏ seal lấy dây buộc lại. Thúy lập tức can ngăn: Tuyệt đối giữ nguyên hiện trường! Nếu tự ý cắt seal, hải quan có quyền nghi ngờ đánh tráo hàng hóa trong quá trình giám sát, lô hàng sẽ bị đình chỉ và chuyển kiểm tra thực tế 100% luồng Đỏ.
Thúy hướng dẫn khách: Chụp ảnh hiện trường xe và seal, lập biên bản ghi nhận sự cố bất khả kháng có chữ ký 3 bên (tài xế vận tải, kho nhận, công chức hải quan giám sát). Nhờ vậy, lô hàng được niêm phong lại và thông quan an toàn sau nửa ngày.

Chuyện thứ hai: Con ốc vít và cái bẫy mã HS Code.
Ốc vít thông thường bằng thép thuộc nhóm 73.18 thuế suất 0% hoặc rất thấp. Nhưng nếu con ốc vít đó được nhà sản xuất thiết kế chuyên dụng cho máy in công nghiệp hoặc máy phay CNC, hải quan có thể phân loại theo máy chính ở Chương 84 với thuế suất và thuế tự vệ cao hơn hẳn! Lệch một con số trong mã 8 chữ số, doanh nghiệp bị bác mã và truy thu tiền tỷ!

Cách phòng ngừa rủi ro tốt nhất: Làm thủ tục Xác định trước mã số (Advance Ruling) theo quy định của Luật Hải quan trước khi hàng về, gửi mẫu thật hoặc tài liệu kỹ thuật lên Cục Hải quan để có văn bản xác định mã số pháp lý vững chắc.""",
            "Quay Chị Thúy trên tay cầm một chiếc chì niêm phong (seal bolt) chuẩn của hãng tàu và vài con ốc vít cơ khí. Bàn trà có cuốn Biểu thuế XNK dày cộp. Chèn hình ảnh biên bản hiện trường mẫu.",
            "Không áp dụng (Tập đơn).",
            r"""[KẾT LUẬN TỰ THÂN CỦA THÚY]
"Từ những va vấp thực tế đó, bài học lớn nhất mà bản thân Thúy tự nhắc nhở mình mỗi ngày là: 'Trong nghề xuất nhập khẩu, sự điềm tĩnh và tính tuân thủ pháp lý là phao cứu sinh lớn nhất khi sự cố xảy ra'. Bản thân mình không bao giờ cho phép mình làm ẩu hay che giấu sự cố, bởi vì một biên bản hiện trường minh bạch luôn có sức mạnh bảo vệ doanh nghiệp tốt hơn ngàn lời giải thích vòng vo." """,
            "Thứ 4, 11/11/2026 (20:00)",
            "Chờ Chị Thúy duyệt"
        )
    ]

    for row_idx, rdata in enumerate(scripts_data, 5):
        kb_code = rdata[0]
        # Color coding
        if kb_code == "KB-01":
            fill_curr = fill_light
        elif kb_code == "KB-08":
            fill_curr = fill_special
        elif int(kb_code.split("-")[1]) % 2 == 0:
            fill_curr = fill_zebra
        else:
            fill_curr = PatternFill(start_color=WHITE, end_color=WHITE, fill_type="solid")

        for col_idx, val in enumerate(rdata, 1):
            cell = ws2.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.fill = fill_curr
            
            if col_idx in [1, 3, 6, 12, 13]:
                cell.alignment = align_top_center
            else:
                cell.alignment = align_top_left

    col_widths_s2 = [10, 35, 18, 24, 32, 14, 35, 65, 30, 32, 35, 20, 18]
    for i, w in enumerate(col_widths_s2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ==========================================
    # SHEET 3: CHECKLIST QUAY BATCH SERIES
    # ==========================================
    ws3 = wb.create_sheet(title="CHECKLIST QUAY BATCH SERIES")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:G1")
    ws3["A1"] = "CHECKLIST QUẢN TRỊ QUAY BATCH SERIES (MỖI ĐỢT QUAY TRỌN 1 KỊCH BẢN / SERIES)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(horizontal="left", vertical="center")

    ws3.merge_cells("A2:G2")
    ws3["A2"] = "Mục tiêu: Tối ưu thời gian của Chị Thúy, 1 buổi quay (3-4 tiếng) hoàn tất trọn vẹn 1 kịch bản nhiều tập hoặc 3 kịch bản đơn lẻ để dựng trước 1 tuần."
    ws3["A2"].font = font_subtitle

    headers_cl = ["Hạng mục chuẩn bị", "Chi tiết kỹ thuật & Đạo cụ", "Tiêu chuẩn nghiệm thu (Quality Gate)", "Người phụ trách", "Ghi chú thực chiến"]
    for col_idx, h in enumerate(headers_cl, 1):
        cell = ws3.cell(row=4, column=col_idx)
        cell.value = h
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    checklist_data = [
        ("1. Không gian & Ánh sáng", "Bàn gỗ mộc, đèn Key Light 3500K (vàng ấm), Fill Light dịu nhẹ, đèn viền tóc (Rim light). Phía sau là giá sách hoặc chậu cây tĩnh lặng.", "Không bị bóng gắt trên mặt, tạo cảm giác ấm cúng, thư thái như ngồi trò chuyện tại nhà.", "Đội ngũ Media / Kỹ thuật", "Tránh ánh sáng trắng xanh 5600K kiểu phòng mổ hoặc livestream bán hàng thông thường."),
        ("2. Âm thanh (Audio First)", "Micro cài áo không dây (Rode Wireless PRO / Hollyland Lark) gắn gọn gàng, có lọc gió nhẹ. Thu file backup WAV 24-bit 48kHz.", "Âm thanh trong trẻo, giọng đọc trầm ấm, không có tiếng quạt gió hay tạp âm đường phố.", "Chị Thúy & Kỹ thuật âm thanh", "Giọng nói từ tốn (130-150 từ/phút), ngắt nghỉ tự nhiên sau mỗi luận điểm quan trọng."),
        ("3. Đạo cụ theo từng Kịch bản", "KB-01: Tách trà ấm + Sổ da mộc + Mô hình container mini.\nKB-02/03: eTax Mobile + Đĩa bánh hamburger mini.\nKB-05/06: Cuốn Luật Hải quan + Mô hình tàu FOB/CIF.\nKB-08: Bản vẽ khóa Twistlock + Mô hình container.\nKB-09: Chiếc seal chì hải quan thật + ốc vít cơ khí.", "Đạo cụ thật 100%, đặt tự nhiên trên bàn để làm điểm tựa visual khi nói.", "Bộ phận Hậu cần", "Không để quá nhiều đồ lộn xộn trên bàn, giữ mặt bàn mộc mạc tinh tế."),
        ("4. Máy nhắc chữ (Teleprompter)", "Cài đặt file kịch bản lời thoại chi tiết vào iPad/máy nhắc chữ đặt thẳng trước ống kính Camera 1.", "Chữ chạy vừa tốc độ nói, mắt nhìn tự nhiên vào ống kính, không bị liếc đảo.", "Trợ lý nội dung", "Chỉ dùng để nhớ ý chính và các con số/điều luật, không đọc như tụng kinh."),
        ("5. Phân công 2 Góc máy", "Cam 1 (Chính diện Eye-Level): 50mm f/1.8 quay từ ngực trở lên (tâm tình chính).\nCam 2 (Góc nghiêng 45 độ): 85mm f/1.8 quay cận góc nghiêng khi trầm ngâm, chiêm nghiệm.", "Chuyển góc máy mượt mà mỗi 15-20 giây để giữ chân người xem YouTube.", "Camera Operator / Editor", "Đồng bộ Timecode giữa 2 máy để hậu kỳ cắt dựng nhanh chóng."),
        ("6. Quy trình Dựng trước 1 tuần", "Sau khi quay xong vào Thứ Bảy: Hậu kỳ dựng thô Chủ Nhật, hoàn thiện đồ họa/screencast Thứ Ba, duyệt bản cuối Thứ Năm, sẵn sàng lịch đăng Thứ Hai tuần kế tiếp.", "Video hoàn thành trước ngày phát sóng ít nhất 3 ngày để kiểm duyệt kỹ lưỡng.", "Video Editor (Remotion/CapCut)", "Luôn có bản master xuất chuẩn 4K hoặc 1080p sắc nét, âm lượng chuẩn YouTube -14 LUFS.")
    ]

    for row_idx, row_data in enumerate(checklist_data, 5):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws3.cell(row=row_idx, column=col_idx)
            cell.value = val
            cell.font = font_regular
            cell.border = thin_border
            cell.alignment = align_top_left

    col_widths_s3 = [22, 45, 35, 25, 35]
    for i, w in enumerate(col_widths_s3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w

    # Save to tmp and atomic replace
    wb.save(TMP_FILE)
    if os.path.exists(TMP_FILE):
        os.replace(TMP_FILE, TARGET_FILE)
        print(f"Workbook successfully saved and atomic replaced at:\n{TARGET_FILE}")
    else:
        raise FileNotFoundError(f"Tmp file {TMP_FILE} was not created.")

if __name__ == "__main__":
    create_calendar_workbook()
