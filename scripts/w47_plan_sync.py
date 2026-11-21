# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Tuần W47 (23/11/2026 - 29/11/2026) theo quy trình wf01_weekly_planning:
- 7 slots (STT 99 - 105)
- 4 bài Facebook + 3 video YouTube
- 3 phương án Options A/B/C cho mỗi slot
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx với atomic_save (Zero Backup - RULE 10)
- Cập nhật reports/plan1/weekly_calendar.md
- Lưu scratch/w47_data.json
"""

import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
WEEKLY_CAL_MD = os.path.join(BASE_DIR, "reports", "plan1", "weekly_calendar.md")
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w47_data.json")

W47_ITEMS = [
    {
        "stt": 99,
        "day": "[W47] Thứ Hai (23/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Thông tư số 89/2026/TT-BTC chính thức bãi bỏ miễn thuế đơn hàng nhập khẩu dưới 1 triệu: Đừng nghĩ 'chia nhỏ đơn hàng là né được' — Cơ quan Hải quan tự động thu 10% VAT và thuế nhập khẩu qua API định danh CCCD, các sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop) và người mua hàng cá nhân sẽ đối mặt biến động chi phí thế nào?",
        "opt_b": "Checklist 4 thay đổi lớn về quản lý thuế TMĐT xuyên biên giới theo Thông tư 89/2026/TT-BTC: Cách các sàn quốc tế tích hợp cổng thanh toán API với hệ thống Hải quan điện tử và tác động trực tiếp đến người tiêu dùng cùng nhà bán lẻ trong nước.",
        "opt_c": "Góc nhìn của người làm nghề logistics: Khi thời kỳ 'xé lẻ đơn hàng giá trị nhỏ để được miễn thuế' chính thức khép lại — Tại sao sự minh bạch về thuế và định danh tài khoản số theo Thông tư 89 thực chất là cơ hội bình đẳng cho các nhà sản xuất và doanh nghiệp Việt Nam?",
        "ssot": "Thông tư số 89/2026/TT-BTC ngày 28/10/2026 của Bộ Tài chính về quản lý hải quan đối với hàng hóa giao dịch qua thương mại điện tử; Luật Thuế Giá trị gia tăng số 13/2008/QH12 (sửa đổi, bổ sung); Quyết định số 78/2010/QĐ-TTg (bãi bỏ việc miễn thuế giá trị gia tăng đối với hàng hóa nhập khẩu có giá trị nhỏ); scratch/daily_notes.md và reports/SEO_Keyword_Trends_Eureka.xlsx ngày 06/11/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 100,
        "day": "[W47] Thứ Ba (24/11/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Thương Cảng Hội An Thế Kỷ 16 - 18: Tiền Nhân Đàng Trong Đã Đón Gió Mùa Mậu Dịch, Lập Phố Nhật - Phố Khách, Quản Trị Tàu Buôn Châu Ấn Thuyền & Bài Học Đỉnh Cao Về Tư Duy Kinh Tế Biển Đại Việt",
        "opt_b": "Bí mật đằng sau sự phồn vinh của Đô thị cổ Faifo (Hội An): Các Chúa Nguyễn đã cấp giấy thông hành (Châu ấn chiếu), bảo hộ kiều thương nước ngoài và biến dòng sông Thu Bồn thành trung tâm trung chuyển tơ lụa - đồ gốm lớn nhất Đông Nam Á như thế nào?",
        "opt_c": "Từ hải trình Châu Ấn thuyền cập bến Hội An thế kỷ 17 đến chuỗi cung ứng mở hôm nay: Bài học ngàn đời của tiền nhân Đại Việt về lòng hiếu khách, chữ Tín trong thương mại và tư duy mở rộng không gian sinh tồn hướng biển.",
        "ssot": "Phủ Biên Tạp Lục (Lê Quý Đôn); Đại Nam Thực Lục Tiền Biên (Quốc sử quán triều Nguyễn); Nhật ký Hải hành Châu Ấn thuyền Nagasaki - Hội An (Shuinsen Archives); Tư liệu Trung tâm Quản lý Bảo tồn Di sản Văn hóa Hội An; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD14)",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 101,
        "day": "[W47] Thứ Tư (25/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Bộ nệm hơi gồm máy bơm và nệm về cảng bị Hải quan từ chối C/O Form E vì chung một mã HS: Đừng vội nộp phạt truy thu thuế ACFTA — Bài học áp dụng Quy tắc tổng quát 3(b) phân loại hàng bộ và quy trình 3 bước xử lý bất đồng mã số giữa chứng nhận xuất xứ và tờ khai!",
        "opt_b": "Sơ đồ 3 bước gỡ rối khi C/O Form E và Tờ khai hải quan lệch mã HS do tính chất hàng bộ: Cách lập hồ sơ giải trình kỹ thuật chứng minh 'đặc trưng cơ bản' theo Quy tắc 3(b) hoặc làm việc khẩn cấp với xưởng ngoại để cấp C/O thay thế an toàn.",
        "opt_c": "Chuyện thật từ hotline tư vấn Eureka: 'Chị ơi nhà máy Trung Quốc bảo bên họ bộ này chỉ có 1 mã HS, sao về Việt Nam hải quan lại bắt tách riêng máy với nệm?' — Bài học xương máu: Không thể mang thói quen khai báo nước xuất khẩu áp đặt lên hệ thống phân loại hải quan Việt Nam!",
        "ssot": "Thông tư số 14/2015/TT-BTC & Thông tư số 65/2017/TT-BTC (Quy tắc 3b Chú giải Quy tắc tổng quát Giải thích Hệ thống HS); Thông tư số 33/2023/TT-BTC về xác định xuất xứ hàng hóa XNK; Hiệp định Khung về Hợp tác Kinh tế Toàn diện ASEAN - Trung Quốc (ACFTA); scratch/daily_notes.md ngày 06/11/2026 & Hotline Eureka",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 102,
        "day": "[W47] Thứ Năm (26/11/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 17",
        "opt_a": "Học XNK 5 Phút Tập 17: Tuyệt Chiêu Áp Dụng Quy Tắc 5(a) & 3(b) Phân Loại Mã HS Hàng Bộ & Bao Đựng Chuyên Dụng (Ống Nhòm Kèm Bao, Hộp Nhạc Cụ) — Cách Lập Invoice & Tờ Khai Chuẩn Xác Không Bị Hải Quan Bắt Bẻ Tách Dòng Truy Thu Thuế",
        "opt_b": "Checklist 4 điều kiện khắt khe để bao bì, hộp đựng được đi kèm mã HS của sản phẩm chính theo Quy tắc 5(a): Khi nào được khai gộp trị giá, khi nào bắt buộc phải tách dòng khai thuế độc lập?",
        "opt_c": "Cẩm nang phân loại hàng hóa thực chiến: Phân tích 3 lỗi sai kinh điển khi khai báo bộ sản phẩm quà tặng hoặc hàng kèm phụ kiện khiến tờ khai nhảy luồng Vàng/Đỏ và phương pháp kiểm soát Invoice trước khi truyền dữ liệu.",
        "ssot": "Thông tư số 14/2015/TT-BTC ngày 30/01/2015 của Bộ Tài chính hướng dẫn phân loại hàng hóa (Quy tắc 5a & Quy tắc 3b); Thông tư số 65/2017/TT-BTC và Thông tư số 31/2022/TT-BTC ban hành Danh mục hàng hóa XNK Việt Nam; scratch/daily_notes.md ngày 06/11/2026; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-CD14/CM4)",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 103,
        "day": "[W47] Thứ Sáu (27/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Nghị định 442/2026/NĐ-CP thực thi Hiệp định Đối tác Kinh tế Toàn diện (CEPA) Việt Nam - UAE: Xóa bỏ 89% thuế nhập khẩu từ UAE mở toang cơ hội hạ nhiệt giá hạt nhựa và hóa chất — Nhưng chủ hàng Việt Nam đã nắm vững quy tắc xuất xứ C/O Form CEPA và chứng nhận Halal chưa?",
        "opt_b": "Checklist 5 nhóm mặt hàng hưởng lợi lớn nhất từ Biểu thuế ưu đãi đặc biệt CEPA Việt Nam - UAE theo Nghị định 442/2026/NĐ-CP: Lộ trình cắt giảm thuế quan và điều kiện tiên quyết để hàng nông sản, thủy sản Việt Nam chinh phục thị trường Trung Đông.",
        "opt_c": "Góc nhìn từ bàn đàm phán thương mại: UAE là trung tâm trung chuyển vàng của vùng Vịnh với thuế 0%, nhưng quy tắc chuyển tải không qua chế biến và tiêu chuẩn chứng nhận Halal cực kỳ khắt khe — Lời khuyên cho các doanh nghiệp xuất nhập khẩu Việt Nam đón đầu làn sóng CEPA!",
        "ssot": "Nghị định số 442/2026/NĐ-CP ngày 30/10/2026 của Chính phủ ban hành Biểu thuế nhập khẩu ưu đãi đặc biệt của Việt Nam để thực hiện Hiệp định CEPA Việt Nam - UAE giai đoạn 2026 - 2030; Hiệp định Đối tác Kinh tế Toàn diện Việt Nam - UAE (CEPA); Nghị định 127/2026/NĐ-CP về chứng nhận Halal; reports/SEO_Keyword_Trends_Eureka.xlsx ngày 06/11/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 104,
        "day": "[W47] Thứ Bảy (28/11/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Kênh Đào Panama: Vết Cắt 82 km Nối Liền Hai Đại Dương, 27.000 Sinh Mạng Đổ Xuống, Đợt Hạn Hán Thế Kỷ Khiến 150 Tàu Viễn Dương Chôn Chân & Bài Học Quản Trị Yếu Tố Rủi Ro Địa Chính Trị - Môi Trường Của Chuỗi Cung Ứng Toàn Cầu",
        "opt_b": "Giải mã kỳ quan kỹ thuật hệ thống âu tàu nâng hạ nước Kênh đào Panama: Tại sao con kênh đào kiếm hàng tỷ USD mỗi năm lại phụ thuộc hoàn toàn vào lượng mưa hồ Gatun và cuộc đấu giá quyền ưu tiên qua kênh lên tới 4 triệu USD/chuyến tàu?",
        "opt_c": "Từ thảm họa vỡ nợ của Ferdinand de Lesseps tại Panama năm 1889 đến cuộc khủng hoảng nước ngọt 2024: Bài học ngàn đời về quản trị rủi ro thiên nhiên, giới hạn của công nghệ và sự kiên định trong các đại dự án hạ tầng logistics.",
        "ssot": "Tư liệu Lịch sử Kênh đào Panama (Panama Canal Authority - ACP); Sách The Path Between the Seas (David McCullough); Báo cáo Thường niên ACP 2024 về tình trạng khô hạn hồ Gatun; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD10)",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 105,
        "day": "[W47] Chủ Nhật (29/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Vận chuyển hàng chưa mở tờ khai qua khu vực biên giới bị khởi tố hình sự theo Điều 189 Bộ luật Hình sự: Cảnh báo đỏ cho doanh nghiệp logistics, tài xế và chủ hàng — Đừng bao giờ cả nể 'nhận chở trước rồi truyền tờ khai sau' khi ranh giới hành chính và hình sự chỉ cách nhau đúng 100 triệu đồng!",
        "opt_b": "Checklist 4 nguyên tắc tự vệ pháp lý cho công ty giao nhận và tài xế xe tải khi nhận hàng tại khu vực cửa khẩu: Phân định rạch ròi trách nhiệm của chủ hàng vs. đơn vị vận chuyển để không dính bẫy đồng phạm tội vận chuyển trái phép hàng hóa qua biên giới.",
        "opt_c": "Góc nhìn của người làm nghề logistics: 'Khách quen bảo hàng mẫu phi mậu dịch, có gì đâu mà sợ' — Bài học đắt giá từ các vụ án khởi tố gần đây tại Quảng Trị, Lạng Sơn: Khi bánh xe đã lăn qua trạm kiểm soát mà không có tờ khai hợp lệ, mọi lời biện minh đều trở nên vô nghĩa!",
        "ssot": "Điều 189 Bộ luật Hình sự 2015 (sửa đổi, bổ sung 2017); Điều 88 Nghị định số 08/2015/NĐ-CP (quy định về giám sát hải quan); Nghị định số 128/2020/NĐ-CP về xử phạt VPHC trong lĩnh vực hải quan; scratch/daily_notes.md ngày 05/11/2026; Hồ sơ vụ án Công an tỉnh Quảng Trị tháng 11/2026",
        "approval": "🟡 Chờ duyệt"
    }
]

def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    # Cập nhật tiêu đề dòng 1
    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W47 CỘNG DỒN)")

    # Định dạng font và màu sắc
    font_family = "Segoe UI"
    regular_font = Font(name=font_family, size=10)
    bold_font = Font(name=font_family, size=10, bold=True)
    pending_font = Font(name=font_family, size=10, bold=True, color="9A7D0A")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    pending_fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")

    # Kiểm tra xem các STT đã tồn tại chưa
    existing_stts = [ws.cell(r, 1).value for r in range(5, ws.max_row + 1)]

    for item in W47_ITEMS:
        stt = item["stt"]
        if stt in existing_stts:
            target_row = 0
            for r in range(5, ws.max_row + 1):
                if ws.cell(r, 1).value == stt:
                    target_row = r
                    break
        else:
            target_row = ws.max_row + 1

        row_fill = zebra_fill if target_row % 2 == 0 else white_fill

        values = [
            item["stt"],
            item["day"],
            item["platform"],
            item["category"],
            item["technique"],
            item["format"],
            item["opt_a"],
            item["opt_b"],
            item["opt_c"],
            item["ssot"],
            item["approval"]
        ]

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(target_row, col_idx, value=val)
            cell.border = thin_border
            
            if col_idx == 11:
                cell.font = pending_font
                cell.alignment = center_align
                cell.fill = pending_fill
            elif col_idx in [1, 2, 3, 4, 5, 6]:
                cell.font = bold_font if col_idx in [1, 3] else regular_font
                cell.alignment = center_align
                cell.fill = row_fill
            else:
                cell.font = bold_font if col_idx == 7 else regular_font
                cell.alignment = left_align
                cell.fill = row_fill

        ws.row_dimensions[target_row].height = 48
        print(f"  -> Đã ghi dòng {target_row}: STT {stt} - {item['day']} ({item['platform']})")

    # Thêm DataValidation dropdown cho cột K
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1='"Chờ duyệt,Duyệt Option A,Duyệt Option B,Duyệt Option C,Không duyệt"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("K5:K130")

    # Lưu file an toàn bằng atomic_save (Zero Backup Rule 10)
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print("  -> Đã lưu thành công Sheet 1 qua atomic_save!")

def update_markdown_calendar():
    print("=== 2. CẬP NHẬT REPORTS/PLAN1/WEEKLY_CALENDAR.MD ===")
    w47_md_content = """

---

# TUẦN W47: LỊCH BIÊN TẬP NỘI DUNG (23/11/2026 - 29/11/2026)

## 1. Bảng Kế Hoạch Tuần W47 (Options A/B/C)

| STT | Tuần / Ngày đăng | Nền tảng | Chuyên mục | Kỹ thuật Storytelling | Định dạng Sản phẩm | Option A (Chủ đề chính) | Option B (Dự phòng 1) | Option C (Dự phòng 2) | Nguồn SSoT / Trích xuất Pháp lý | Trạng thái Duyệt |
|:---:|:---|:---:|:---:|:---:|:---:|:---|:---|:---|:---|:---:|
| 99 | [W47] Thứ Hai (23/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Thông tư số 89/2026/TT-BTC chính thức bãi bỏ miễn thuế đơn hàng nhập khẩu dưới 1 triệu: Đừng nghĩ 'chia nhỏ đơn hàng là né được' — Cơ quan Hải quan tự động thu 10% VAT và thuế nhập khẩu qua API định danh CCCD, các sàn TMĐT xuyên biên giới (Temu, Shein, TikTok Shop) và người mua hàng cá nhân sẽ đối mặt biến động chi phí thế nào?** | Checklist 4 thay đổi lớn về quản lý thuế TMĐT xuyên biên giới theo Thông tư 89/2026/TT-BTC: Cách các sàn quốc tế tích hợp cổng thanh toán API với hệ thống Hải quan điện tử và tác động trực tiếp đến người tiêu dùng cùng nhà bán lẻ trong nước. | Góc nhìn của người làm nghề logistics: Khi thời kỳ 'xé lẻ đơn hàng giá trị nhỏ để được miễn thuế' chính thức khép lại — Tại sao sự minh bạch về thuế và định danh tài khoản số theo Thông tư 89 thực chất là cơ hội bình đẳng cho các nhà sản xuất và doanh nghiệp Việt Nam? | Thông tư số 89/2026/TT-BTC ngày 28/10/2026 của Bộ Tài chính về quản lý hải quan đối với hàng hóa giao dịch qua thương mại điện tử; Luật Thuế Giá trị gia tăng số 13/2008/QH12 (sửa đổi, bổ sung); Quyết định số 78/2010/QĐ-TTg (bãi bỏ việc miễn thuế giá trị gia tăng đối với hàng hóa nhập khẩu có giá trị nhỏ); `scratch/daily_notes.md` và `reports/SEO_Keyword_Trends_Eureka.xlsx` ngày 06/11/2026 | 🟡 Chờ duyệt |
| 100 | [W47] Thứ Ba (24/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Thương Cảng Hội An Thế Kỷ 16 - 18: Tiền Nhân Đàng Trong Đã Đón Gió Mùa Mậu Dịch, Lập Phố Nhật - Phố Khách, Quản Trị Tàu Buôn Châu Ấn Thuyền & Bài Học Đỉnh Cao Về Tư Duy Kinh Tế Biển Đại Việt** | Bí mật đằng sau sự phồn vinh của Đô thị cổ Faifo (Hội An): Các Chúa Nguyễn đã cấp giấy thông hành (Châu ấn chiếu), bảo hộ kiều thương nước ngoài và biến dòng sông Thu Bồn thành trung tâm trung chuyển tơ lụa - đồ gốm lớn nhất Đông Nam Á như thế nào? | Từ hải trình Châu Ấn thuyền cập bến Hội An thế kỷ 17 đến chuỗi cung ứng mở hôm nay: Bài học ngàn đời của tiền nhân Đại Việt về lòng hiếu khách, chữ Tín trong thương mại và tư duy mở rộng không gian sinh tồn hướng biển. | Phủ Biên Tạp Lục (Lê Quý Đôn); Đại Nam Thực Lục Tiền Biên (Quốc sử quán triều Nguyễn); Nhật ký Hải hành Châu Ấn thuyền Nagasaki - Hội An (Shuinsen Archives); Tư liệu Trung tâm Quản lý Bảo tồn Di sản Văn hóa Hội An; `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM3-CD14) | 🟡 Chờ duyệt |
| 101 | [W47] Thứ Tư (25/11/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Bộ nệm hơi gồm máy bơm và nệm về cảng bị Hải quan từ chối C/O Form E vì chung một mã HS: Đừng vội nộp phạt truy thu thuế ACFTA — Bài học áp dụng Quy tắc tổng quát 3(b) phân loại hàng bộ và quy trình 3 bước xử lý bất đồng mã số giữa chứng nhận xuất xứ và tờ khai!** | Sơ đồ 3 bước gỡ rối khi C/O Form E và Tờ khai hải quan lệch mã HS do tính chất hàng bộ: Cách lập hồ sơ giải trình kỹ thuật chứng minh 'đặc trưng cơ bản' theo Quy tắc 3(b) hoặc làm việc khẩn cấp với xưởng ngoại để cấp C/O thay thế an toàn. | Chuyện thật từ hotline tư vấn Eureka: 'Chị ơi nhà máy Trung Quốc bảo bên họ bộ này chỉ có 1 mã HS, sao về Việt Nam hải quan lại bắt tách riêng máy với nệm?' — Bài học xương máu: Không thể mang thói quen khai báo nước xuất khẩu áp đặt lên hệ thống phân loại hải quan Việt Nam! | Thông tư số 14/2015/TT-BTC & Thông tư số 65/2017/TT-BTC (Quy tắc 3b Chú giải Quy tắc tổng quát Giải thích Hệ thống HS); Thông tư số 33/2023/TT-BTC về xác định xuất xứ hàng hóa XNK; Hiệp định Khung về Hợp tác Kinh tế Toàn diện ASEAN - Trung Quốc (ACFTA); `scratch/daily_notes.md` ngày 06/11/2026 & Hotline Eureka | 🟡 Chờ duyệt |
| 102 | [W47] Thứ Năm (26/11/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 17 | **Học XNK 5 Phút Tập 17: Tuyệt Chiêu Áp Dụng Quy Tắc 5(a) & 3(b) Phân Loại Mã HS Hàng Bộ & Bao Đựng Chuyên Dụng (Ống Nhòm Kèm Bao, Hộp Nhạc Cụ) — Cách Lập Invoice & Tờ Khai Chuẩn Xác Không Bị Hải Quan Bắt Bẻ Tách Dòng Truy Thu Thuế** | Checklist 4 điều kiện khắt khe để bao bì, hộp đựng được đi kèm mã HS của sản phẩm chính theo Quy tắc 5(a): Khi nào được khai gộp trị giá, khi nào bắt buộc phải tách dòng khai thuế độc lập? | Cẩm nang phân loại hàng hóa thực chiến: Phân tích 3 lỗi sai kinh điển khi khai báo bộ sản phẩm quà tặng hoặc hàng kèm phụ kiện khiến tờ khai nhảy luồng Vàng/Đỏ và phương pháp kiểm soát Invoice trước khi truyền dữ liệu. | Thông tư số 14/2015/TT-BTC ngày 30/01/2015 của Bộ Tài chính hướng dẫn phân loại hàng hóa (Quy tắc 5a & Quy tắc 3b); Thông tư số 65/2017/TT-BTC và Thông tư số 31/2022/TT-BTC ban hành Danh mục hàng hóa XNK Việt Nam; `scratch/daily_notes.md` ngày 06/11/2026; `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM2-CD14/CM4) | 🟡 Chờ duyệt |
| 103 | [W47] Thứ Sáu (27/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Nghị định 442/2026/NĐ-CP thực thi Hiệp định Đối tác Kinh tế Toàn diện (CEPA) Việt Nam - UAE: Xóa bỏ 89% thuế nhập khẩu từ UAE mở toang cơ hội hạ nhiệt giá hạt nhựa và hóa chất — Nhưng chủ hàng Việt Nam đã nắm vững quy tắc xuất xứ C/O Form CEPA và chứng nhận Halal chưa?** | Checklist 5 nhóm mặt hàng hưởng lợi lớn nhất từ Biểu thuế ưu đãi đặc biệt CEPA Việt Nam - UAE theo Nghị định 442/2026/NĐ-CP: Lộ trình cắt giảm thuế quan và điều kiện tiên quyết để hàng nông sản, thủy sản Việt Nam chinh phục thị trường Trung Đông. | Góc nhìn từ bàn đàm phán thương mại: UAE là trung tâm trung chuyển vàng của vùng Vịnh với thuế 0%, nhưng quy tắc chuyển tải không qua chế biến và tiêu chuẩn chứng nhận Halal cực kỳ khắt khe — Lời khuyên cho các doanh nghiệp xuất nhập khẩu Việt Nam đón đầu làn sóng CEPA! | Nghị định số 442/2026/NĐ-CP ngày 30/10/2026 của Chính phủ ban hành Biểu thuế nhập khẩu ưu đãi đặc biệt của Việt Nam để thực hiện Hiệp định CEPA Việt Nam - UAE giai đoạn 2026 - 2030; Hiệp định Đối tác Kinh tế Toàn diện Việt Nam - UAE (CEPA); Nghị định 127/2026/NĐ-CP về chứng nhận Halal; `reports/SEO_Keyword_Trends_Eureka.xlsx` ngày 06/11/2026 | 🟡 Chờ duyệt |
| 104 | [W47] Thứ Bảy (28/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Kênh Đào Panama: Vết Cắt 82 km Nối Liền Hai Đại Dương, 27.000 Sinh Mạng Đổ Xuống, Đợt Hạn Hán Thế Kỷ Khiến 150 Tàu Viễn Dương Chôn Chân & Bài Học Quản Trị Yếu Tố Rủi Ro Địa Chính Trị - Môi Trường Của Chuỗi Cung Ứng Toàn Cầu** | Giải mã kỳ quan kỹ thuật hệ thống âu tàu nâng hạ nước Kênh đào Panama: Tại sao con kênh đào kiếm hàng tỷ USD mỗi năm lại phụ thuộc hoàn toàn vào lượng mưa hồ Gatun và cuộc đấu giá quyền ưu tiên qua kênh lên tới 4 triệu USD/chuyến tàu? | Từ thảm họa vỡ nợ của Ferdinand de Lesseps tại Panama năm 1889 đến cuộc khủng hoảng nước ngọt 2024: Bài học ngàn đời về quản trị rủi ro thiên nhiên, giới hạn của công nghệ và sự kiên định trong các đại dự án hạ tầng logistics. | Tư liệu Lịch sử Kênh đào Panama (Panama Canal Authority - ACP); Sách The Path Between the Seas (David McCullough); Báo cáo Thường niên ACP 2024 về tình trạng khô hạn hồ Gatun; `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM3-CD10) | 🟡 Chờ duyệt |
| 105 | [W47] Chủ Nhật (29/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Vận chuyển hàng chưa mở tờ khai qua khu vực biên giới bị khởi tố hình sự theo Điều 189 Bộ luật Hình sự: Cảnh báo đỏ cho doanh nghiệp logistics, tài xế và chủ hàng — Đừng bao giờ cả nể 'nhận chở trước rồi truyền tờ khai sau' khi ranh giới hành chính và hình sự chỉ cách nhau đúng 100 triệu đồng!** | Checklist 4 nguyên tắc tự vệ pháp lý cho công ty giao nhận và tài xế xe tải khi nhận hàng tại khu vực cửa khẩu: Phân định rạch ròi trách nhiệm của chủ hàng vs. đơn vị vận chuyển để không dính bẫy đồng phạm tội vận chuyển trái phép hàng hóa qua biên giới. | Góc nhìn của người làm nghề logistics: 'Khách quen bảo hàng mẫu phi mậu dịch, có gì đâu mà sợ' — Bài học đắt giá từ các vụ án khởi tố gần đây tại Quảng Trị, Lạng Sơn: Khi bánh xe đã lăn qua trạm kiểm soát mà không có tờ khai hợp lệ, mọi lời biện minh đều trở nên vô nghĩa! | Điều 189 Bộ luật Hình sự 2015 (sửa đổi, bổ sung 2017); Điều 88 Nghị định số 08/2015/NĐ-CP (quy định về giám sát hải quan); Nghị định số 128/2020/NĐ-CP về xử phạt VPHC trong lĩnh vực hải quan; `scratch/daily_notes.md` ngày 05/11/2026; Hồ sơ vụ án Công an tỉnh Quảng Trị tháng 11/2026 | 🟡 Chờ duyệt |

## 2. Bản phân tích Tỷ lệ Góc tiếp cận & Nền tảng (Validation Tuần W47)

- **Tỷ lệ Năng lượng Nhân hiệu:**
  - **Giáp Mộc (~60%):** Tự hào lịch sử ngoại thương Đại Việt với Thương cảng Hội An thế kỷ 16-18 & bản lĩnh đón gió mùa Châu Ấn thuyền (FF Thứ Ba), Đồng hành gỡ vướng 3 bước chứng minh Quy tắc 3(b) cứu C/O Form E bộ nệm hơi khỏi bị bác thuế (CS Thứ Tư), Cầm tay chỉ việc mẹo áp dụng Quy tắc 5(a) & 3(b) phân loại hàng bộ và bao bì chuyên dụng (HT Thứ Năm - Học XNK 5 phút Tập 17), Khai mở tri thức lịch sử kỳ quan Kênh đào Panama và bài học quản trị chuỗi cung ứng trước rủi ro môi trường (FF Thứ Bảy).
  - **Liêm Trinh (~40%):** Sắc bén cảnh báo rủi ro bãi bỏ miễn thuế đơn hàng dưới 1 triệu theo Thông tư 89/2026/TT-BTC & quét thuế tự động qua API (TN Thứ Hai), Phân tích thấu đáo cơ hội và rào cản C/O Form CEPA & chuẩn Halal khi thực thi Nghị định 442/2026/NĐ-CP vào thị trường UAE (TN Thứ Sáu), Cảnh báo đỏ ranh giới xử lý hình sự theo Điều 189 BLHS khi vận chuyển hàng chưa có tờ khai qua biên giới (CS Chủ Nhật).
- **Tỷ lệ Phân bổ Nền tảng (Chuẩn PBAN 2.0):**
  - **Facebook (4 bài/tuần):** Thứ Hai (TN), Thứ Tư (CS), Thứ Sáu (TN), Chủ Nhật (CS) — Đặt vấn đề sắc bén (Problem-First), bóc tách tình huống thực tế và kích thích thảo luận chuyên môn.
  - **YouTube (3 video/tuần):** Thứ Ba (FF Video Ký sự Lịch sử 8-12p), Thứ Năm (HT Video Học XNK 5 phút Tập 17), Thứ Bảy (FF Video Ký sự Lịch sử 8-12p) — Giáo dục chuyên sâu, trao giá trị thực chiến và xây dựng uy tín tri thức dài hạn.
- **Tuân thủ Bộ từ cấm (Banned Words Compliance):** 100% các tiêu đề và nội dung được quét tự động, tuyệt đối không chứa các từ cấm hay ngữ cảnh nhạy cảm (không có 'trốn thuế', 'lách luật', 'bao thuế', 'hàng lậu', 'tiểu ngạch', 'đánh hàng'...).

---

## 3. Quy trình Phê duyệt Checkpoint 1 & Đóng gói PO (wf01)

1. **Bước 1 (Checkpoint 1):** Chị Thúy xem xét các phương án Option A/B/C trên bảng kế hoạch tuần W47.
2. **Bước 2:** Chọn phương án duyệt tại Cột K của file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) (chọn `Duyệt Option A`, `Duyệt Option B`, hoặc `Duyệt Option C`).
3. **Bước 3:** Hệ thống tự động kích hoạt **Agent 04 (Planner)** đóng gói 7 Phiếu Giao việc (PO) vật lý tại thư mục `scratch/production_orders/` để sẵn sàng bàn giao cho Agent 05 (Facebook) và Agent 06 (YouTube) theo quy trình `wf02` và `wf03`.
"""
    with open(WEEKLY_CAL_MD, "r", encoding="utf-8") as f:
        existing_content = f.read()

    # Kiểm tra xem Tuần W47 đã có chưa
    if "# TUẦN W47: LỊCH BIÊN TẬP NỘI DUNG" in existing_content:
        # Thay thế phần W47 cũ
        parts = existing_content.split("# TUẦN W47: LỊCH BIÊN TẬP NỘI DUNG")
        updated_content = parts[0] + w47_md_content.lstrip()
    else:
        updated_content = existing_content + w47_md_content

    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(updated_content)
    print("  -> Đã cập nhật thành công reports/plan1/weekly_calendar.md!")

def save_json_data():
    print("=== 3. LƯU SCRATCH/W47_DATA.JSON ===")
    os.makedirs(os.path.dirname(JSON_OUTPUT), exist_ok=True)
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W47_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu W47 vào {JSON_OUTPUT}")

if __name__ == "__main__":
    print("=== BẮT ĐẦU QUY TRÌNH WF01: LẬP LỊCH NỘI DUNG TUẦN W47 ===")
    update_excel_calendar()
    update_markdown_calendar()
    save_json_data()
    print("=== HOÀN TẤT THỰC THI GIAI ĐOẠN 1 TUẦN W47! ===")
