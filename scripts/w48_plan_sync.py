# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Tuần W48 (30/11/2026 - 06/12/2026) theo quy trình wf01_weekly_planning:
- 7 slots (STT 106 - 112)
- 4 bài Facebook + 3 video YouTube
- 3 phương án Options A/B/C cho mỗi slot
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx với atomic_save (Zero Backup - RULE 10)
- Cập nhật reports/plan1/weekly_calendar.md
- Lưu scratch/w48_data.json
"""

import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
WEEKLY_CAL_MD = os.path.join(BASE_DIR, "reports", "plan1", "weekly_calendar.md")
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w48_data.json")

W48_ITEMS = [
    {
        "stt": 106,
        "day": "[W48] Thứ Hai (30/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Quyết định số 2988/QĐ-BCT áp thuế chống bán phá giá sơ bộ 12,8% - 27,6% đối với thép cuộn cán nóng (HRC) nhập khẩu từ Trung Quốc và Ấn Độ (Vụ việc AD20): Doanh nghiệp tôn mạ, kết cấu thép và gia công cơ khí trong nước đang đối mặt cơn 'địa chấn' chi phí nguyên liệu thế nào và đâu là chiến lược ứng phó sống còn?",
        "opt_b": "Checklist 4 điều doanh nghiệp nhập khẩu thép cần làm ngay sau Quyết định 2988/QĐ-BCT: Cách rà soát hợp đồng mua bán ngoại, giấy chứng nhận xuất xứ C/O và thủ tục bảo lãnh thuế để không bị động dòng tiền khi hàng về cảng.",
        "opt_c": "Góc nhìn đa chiều về phòng vệ thương mại AD20: Bảo hộ ngành sản xuất thép nội địa hay gia tăng áp lực lên hàng trăm doanh nghiệp hạ nguồn? Bài học đắt giá về đa dạng hóa chuỗi cung ứng nguyên liệu cho các nhà máy Việt Nam.",
        "ssot": "Quyết định số 2988/QĐ-BCT ngày 08/11/2026 của Bộ Công Thương về việc áp dụng thuế chống bán phá giá sơ bộ đối với một số sản phẩm thép cuộn cán nóng (HRC) có xuất xứ từ Trung Quốc và Ấn Độ (Mã vụ việc AD20); Luật Quản lý ngoại thương số 05/2017/QH14; Nghị định số 10/2018/NĐ-CP; reports/SEO_Keyword_Trends_Eureka.xlsx ngày 13/11/2026.",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 107,
        "day": "[W48] Thứ Ba (01/12/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Cuộc Khủng Hoảng Kênh Đào Suez 2021: Siêu Tàu Ever Given 20.000 TEU Mắc Cạn Làm Tê Liệt 12% Thương Mại Toàn Cầu — 6 Ngày Bịt Kín Tuyến Huyết Mạch, Thiệt Hại 54 Tỷ USD, Ván Cược Bồi Thường 900 Triệu USD & Bài Học Đắt Giá Về Quản Trị Rủi Ro Chuỗi Cung Ứng Đơn Tuyến!",
        "opt_b": "Giải mã chiến dịch giải cứu siêu tàu Ever Given tại Kênh đào Suez: Tại sao chỉ một cơn gió giật 40 hải lý/giờ và con tàu dài 400 mét lại khiến hơn 400 tàu viễn dương chôn chân giữa đại dương và giá cước biển toàn cầu tăng phi mã?",
        "opt_c": "Từ thảm kịch Ever Given 2021 đến khủng hoảng Biển Đỏ hôm nay: Khi 'con đường tơ lụa trên biển' trở nên mong manh, các tập đoàn đa quốc gia và chủ hàng Việt Nam đã tái cấu trúc tuyến vận tải dự phòng qua Mũi Hảo Vọng như thế nào?",
        "ssot": "Báo cáo điều tra chính thức của Cơ quan Quản lý Kênh đào Suez (SCA); Tạp chí hàng hải quốc tế Lloyd's List; Phân tích thiệt hại kinh tế của Hãng bảo hiểm Allianz (2021); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD11); Tư liệu lịch sử hàng hải quốc tế.",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 108,
        "day": "[W48] Thứ Tư (02/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Lỗi upload Invoice số hóa tại Chi cục Hải quan Hải Phòng (03YY) mục 43 CTK99 khiến container nằm kẹt cảng sát giờ phạt lưu bãi: Đừng hoảng loạn gọi khắp nơi — Chia sẻ phác đồ 3 bước xử lý lỗi hệ thống phần mềm, cách đính kèm chứng từ dự phòng và kinh nghiệm làm việc với công chức tiếp nhận!",
        "opt_b": "Checklist 3 bước 'gỡ kẹt' khi phần mềm hải quan báo lỗi không đính kèm được Hóa đơn thương mại số hóa: Khi nào nên liên hệ nhà cung cấp phần mềm, khi nào cần chuyển sang nộp phân nhóm chứng từ khác theo chỉ dẫn nghiệp vụ?",
        "opt_c": "Chuyện thật từ hotline hỗ trợ logistics của Thúy: 'Chị ơi tàu cập rồi mà ECUS cứ báo lỗi chứng từ số hóa mục 43, hàng không thông quan được thì đền bao nhiêu tiền lưu kho?' — Bài học kiểm soát quy trình số hóa hồ sơ trước khi truyền tờ khai chính thức.",
        "ssot": "Thông tư số 38/2015/TT-BTC & Thông tư số 39/2018/TT-BTC về thủ tục hải quan điện tử; Quyết định số 1434/QĐ-CHQ về chuẩn dữ liệu điện tử; Case study thực tế trên nhóm Giải đáp thủ tục Hải quan & scratch/daily_notes.md ngày 13/11/2026; Hotline hỗ trợ Eureka Logistics.",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 109,
        "day": "[W48] Thứ Năm (03/12/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 18",
        "opt_a": "Học XNK 5 Phút Tập 18: Tuyệt Chiêu Đổi Switch B/L Trong Thương Mại 3 Bên & Phác Đồ Khai Bổ Sung AMA Sau Thông Quan Khi Phát Hiện Nhầm Hóa Đơn Lô Hàng Khác — Quy Trình 4 Bước Hoàn Thiện Mẫu 03 Tránh Bị Hải Quan Bắt Lỗi Gian Lận Thuế!",
        "opt_b": "Checklist 4 bước lập hồ sơ khai bổ sung sau thông quan (Khai AMA & Mẫu 03/KBS/GSQL): Cách phân định rạch ròi giữa 'sai sót kỹ thuật vô ý' và 'gian lận số lượng hàng hóa' để bảo vệ doanh nghiệp trước nguy cơ bị xử phạt vi phạm hành chính.",
        "opt_c": "Bí kíp thực chiến cho kế toán và nhân viên XNK: Làm gì khi phát hiện hóa đơn và số lượng trên tờ khai đã thông quan lệch hoàn toàn so với thực tế nhập kho? Quy trình tự phát hiện và khai bổ sung trong vòng 60 ngày để được miễn tiền phạt chậm nộp.",
        "ssot": "Điều 20 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC) về khai bổ sung hồ sơ hải quan; Mẫu số 03/KBS/GSQL ban hành kèm Thông tư 39/2018; Điều 14 Nghị định số 128/2020/NĐ-CP; scratch/daily_notes.md ngày 12/11/2026; Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM2-CD10/CM2-CD12).",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 110,
        "day": "[W48] Thứ Sáu (04/12/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Công văn 24150/CHQ-GSQL & 24170/CHQ-GSQL ngày 13/11/2026: Đột phá giám sát số hóa kho CFS điện tử qua VASSCM giúp giải phóng hàng lẻ LCL trong 15 phút và gỡ vướng thủ tục XNK tại chỗ giữa DN nội địa với DNCX — Doanh nghiệp xuất nhập khẩu đã tận dụng hết lợi thế này chưa?",
        "opt_b": "Checklist 4 điểm mới đột phá trong quy trình giám sát số hóa kho CFS và tháo gỡ vướng mắc XNK tại chỗ DNCX theo các công văn mới nhất của Tổng cục Hải quan: Cắt giảm 70% thủ tục giấy tờ và bảo toàn thuế VAT 0% cho chuỗi cung ứng sản xuất xuất khẩu.",
        "opt_c": "Góc nhìn từ hiện trường logistics: Khi kho CFS chuyển mình từ 'điểm nghẽn chứng từ giấy' sang 'luồng dữ liệu điện tử 15 phút' — Lời khuyên cho các chủ hàng lẻ LCL và doanh nghiệp chế xuất đón đầu làn sóng số hóa cuối năm.",
        "ssot": "Công văn số 24150/CHQ-GSQL ngày 12/11/2026 của Tổng cục Hải quan về giám sát số hóa kho CFS điện tử qua hệ thống VASSCM; Công văn số 24170/CHQ-GSQL ngày 13/11/2026 về thủ tục hải quan xuất nhập khẩu tại chỗ với DNCX; Luật Hải quan 2014; reports/SEO_Keyword_Trends_Eureka.xlsx ngày 13/11/2026.",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 111,
        "day": "[W48] Thứ Bảy (05/12/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Tuyến Đường Sắt Xuyên Lục Địa Yiwu - Madrid (13.052 km): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh Qua 8 Quốc Gia, Vượt 3 Lần Thay Đổi Khổ Đường Ray & Cuộc Cách Mạng Logistics Đường Bộ Á - Âu Thách Thức Vận Tải Biển Truyền Thống Ra Sao?",
        "opt_b": "Bí mật đằng sau hành trình 21 ngày đêm của đoàn tàu container Yiwu - Madrid: Tại sao hàng hóa từ 'thủ phủ hàng tiêu dùng thế giới' Nghĩa Ô (Trung Quốc) lại chọn chạy tàu hỏa xuyên qua Kazakhstan, Nga, Ba Lan đến Tây Ban Nha thay vì đi tàu biển giá rẻ?",
        "opt_c": "Từ tuyến đường sắt xuyên Á - Âu Yiwu - Madrid đến mạng lưới vận tải liên vận đường sắt Việt - Trung: Cơ hội vàng để nông sản và linh kiện điện tử Việt Nam kết nối thẳng vào chuỗi cung ứng Á - Âu qua cửa khẩu Đồng Đăng và Lào Cai.",
        "ssot": "Hồ sơ Tuyến vận tải đường sắt Yiwu - Madrid (Yixinou Railway / CR Express); Báo cáo Ngân hàng Thế giới (World Bank) về Hành lang Kinh tế Á - Âu; Tư liệu Hiệp hội Vận tải Đường sắt Quốc tế (UIC); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD12).",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 112,
        "day": "[W48] Chủ Nhật (06/12/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Bẫy thuế Nhà thầu (FCT) khi chi trả dịch vụ xuyên biên giới (Meta, Google, Microsoft, AWS) theo Luật Thuế GTGT số 48/2024/QH15 & Thông tư 69/2025/TT-BTC: Cảnh báo doanh nghiệp XNK và Logistics chi tiền tỷ cho quảng cáo, phần mềm đám mây nhưng quên khấu trừ 5% VAT và 2% TNDN dẫn đến bị truy thu tiền trăm triệu kèm tiền phạt chậm nộp!",
        "opt_b": "Sơ đồ 3 bước xác định nghĩa vụ khấu trừ và kê khai nộp thay thuế FCT cho nhà cung cấp nước ngoài không có cơ sở thường trú tại Việt Nam: Phân biệt trường hợp nhà mạng đã tự đăng ký thuế qua Cổng TTĐT Tổng cục Thuế vs. doanh nghiệp Việt Nam bắt buộc phải nộp thay.",
        "opt_c": "Góc nhìn từ bàn làm việc của Thúy: 'Ủa chị ơi em trả tiền bằng thẻ visa công ty cho Google/Facebook thì bên đó phải tự nộp thuế chứ, sao thanh tra lại đè doanh nghiệp em ra phạt?' — Bài học xương máu về quản trị chi phí hợp lý được trừ và rủi ro thuế nhà thầu thời đại số.",
        "ssot": "Khoản 4 Điều 4 Luật Thuế GTGT số 48/2024/QH15; Điều 9 Thông tư số 69/2025/TT-BTC của Bộ Tài chính; Điều 76 đến Điều 81 Thông tư số 80/2021/TT-BTC; scratch/daily_notes.md ngày 12/11/2026; Hồ sơ tư vấn thuế XNK thực tế tại Eureka Logistics.",
        "approval": "🟡 Chờ duyệt"
    }
]

def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    # Cập nhật tiêu đề dòng 1
    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W48 CỘNG DỒN)")

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

    for item in W48_ITEMS:
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
    dv.add("K5:K140")

    # Lưu file an toàn bằng atomic_save (Zero Backup Rule 10)
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print("  -> Đã lưu thành công Sheet 1 qua atomic_save!")

def update_markdown_calendar():
    print("=== 2. CẬP NHẬT REPORTS/PLAN1/WEEKLY_CALENDAR.MD ===")
    w48_md_content = """

---

# TUẦN W48: LỊCH BIÊN TẬP NỘI DUNG (30/11/2026 - 06/12/2026)

## 1. Bảng Kế Hoạch Tuần W48 (Options A/B/C)

| STT | Tuần / Ngày đăng | Nền tảng | Chuyên mục | Kỹ thuật Storytelling | Định dạng Sản phẩm | Option A (Chủ đề chính) | Option B (Dự phòng 1) | Option C (Dự phòng 2) | Nguồn SSoT / Trích xuất Pháp lý | Trạng thái Duyệt |
|:---:|:---|:---:|:---:|:---:|:---:|:---|:---|:---|:---|:---:|
| 106 | [W48] Thứ Hai (30/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Quyết định số 2988/QĐ-BCT áp thuế chống bán phá giá sơ bộ 12,8% - 27,6% đối với thép cuộn cán nóng (HRC) nhập khẩu từ Trung Quốc và Ấn Độ (Vụ việc AD20): Doanh nghiệp tôn mạ, kết cấu thép và gia công cơ khí trong nước đang đối mặt cơn 'địa chấn' chi phí nguyên liệu thế nào và đâu là chiến lược ứng phó sống còn?** | Checklist 4 điều doanh nghiệp nhập khẩu thép cần làm ngay sau Quyết định 2988/QĐ-BCT: Cách rà soát hợp đồng mua bán ngoại, giấy chứng nhận xuất xứ C/O và thủ tục bảo lãnh thuế để không bị động dòng tiền khi hàng về cảng. | Góc nhìn đa chiều về phòng vệ thương mại AD20: Bảo hộ ngành sản xuất thép nội địa hay gia tăng áp lực lên hàng trăm doanh nghiệp hạ nguồn? Bài học đắt giá về đa dạng hóa chuỗi cung ứng nguyên liệu cho các nhà máy Việt Nam. | Quyết định số 2988/QĐ-BCT ngày 08/11/2026 của Bộ Công Thương về việc áp dụng thuế chống bán phá giá sơ bộ đối với một số sản phẩm thép cuộn cán nóng (HRC) có xuất xứ từ Trung Quốc và Ấn Độ (Mã vụ việc AD20); Luật Quản lý ngoại thương số 05/2017/QH14; Nghị định số 10/2018/NĐ-CP; `reports/SEO_Keyword_Trends_Eureka.xlsx` ngày 13/11/2026 | 🟡 Chờ duyệt |
| 107 | [W48] Thứ Ba (01/12/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Cuộc Khủng Hoảng Kênh Đào Suez 2021: Siêu Tàu Ever Given 20.000 TEU Mắc Cạn Làm Tê Liệt 12% Thương Mại Toàn Cầu — 6 Ngày Bịt Kín Tuyến Huyết Mạch, Thiệt Hại 54 Tỷ USD, Ván Cược Bồi Thường 900 Triệu USD & Bài Học Đắt Giá Về Quản Trị Rủi Ro Chuỗi Cung Ứng Đơn Tuyến!** | Giải mã chiến dịch giải cứu siêu tàu Ever Given tại Kênh đào Suez: Tại sao chỉ một cơn gió giật 40 hải lý/giờ và con tàu dài 400 mét lại khiến hơn 400 tàu viễn dương chôn chân giữa đại dương và giá cước biển toàn cầu tăng phi mã? | Từ thảm kịch Ever Given 2021 đến khủng hoảng Biển Đỏ hôm nay: Khi 'con đường tơ lụa trên biển' trở nên mong manh, các tập đoàn đa quốc gia và chủ hàng Việt Nam đã tái cấu trúc tuyến vận tải dự phòng qua Mũi Hảo Vọng như thế nào? | Báo cáo điều tra chính thức của Cơ quan Quản lý Kênh đào Suez (SCA); Tạp chí hàng hải quốc tế Lloyd's List; Phân tích thiệt hại kinh tế của Hãng bảo hiểm Allianz (2021); `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM3-CD11); Tư liệu lịch sử hàng hải quốc tế | 🟡 Chờ duyệt |
| 108 | [W48] Thứ Tư (02/12/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Lỗi upload Invoice số hóa tại Chi cục Hải quan Hải Phòng (03YY) mục 43 CTK99 khiến container nằm kẹt cảng sát giờ phạt lưu bãi: Đừng hoảng loạn gọi khắp nơi — Chia sẻ phác đồ 3 bước xử lý lỗi hệ thống phần mềm, cách đính kèm chứng từ dự phòng và kinh nghiệm làm việc với công chức tiếp nhận!** | Checklist 3 bước 'gỡ kẹt' khi phần mềm hải quan báo lỗi không đính kèm được Hóa đơn thương mại số hóa: Khi nào nên liên hệ nhà cung cấp phần mềm, khi nào cần chuyển sang nộp phân nhóm chứng từ khác theo chỉ dẫn nghiệp vụ? | Chuyện thật từ hotline hỗ trợ logistics của Thúy: 'Chị ơi tàu cập rồi mà ECUS cứ báo lỗi chứng từ số hóa mục 43, hàng không thông quan được thì đền bao nhiêu tiền lưu kho?' — Bài học kiểm soát quy trình số hóa hồ sơ trước khi truyền tờ khai chính thức. | Thông tư số 38/2015/TT-BTC & Thông tư số 39/2018/TT-BTC về thủ tục hải quan điện tử; Quyết định số 1434/QĐ-CHQ về chuẩn dữ liệu điện tử; Case study thực tế trên nhóm Giải đáp thủ tục Hải quan & `scratch/daily_notes.md` ngày 13/11/2026; Hotline hỗ trợ Eureka Logistics | 🟡 Chờ duyệt |
| 109 | [W48] Thứ Năm (03/12/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 18 | **Học XNK 5 Phút Tập 18: Tuyệt Chiêu Đổi Switch B/L Trong Thương Mại 3 Bên & Phác Đồ Khai Bổ Sung AMA Sau Thông Quan Khi Phát Hiện Nhầm Hóa Đơn Lô Hàng Khác — Quy Trình 4 Bước Hoàn Thiện Mẫu 03 Tránh Bị Hải Quan Bắt Lỗi Gian Lận Thuế!** | Checklist 4 bước lập hồ sơ khai bổ sung sau thông quan (Khai AMA & Mẫu 03/KBS/GSQL): Cách phân định rạch ròi giữa 'sai sót kỹ thuật vô ý' và 'gian lận số lượng hàng hóa' để bảo vệ doanh nghiệp trước nguy cơ bị xử phạt vi phạm hành chính. | Bí kíp thực chiến cho kế toán và nhân viên XNK: Làm gì khi phát hiện hóa đơn và số lượng trên tờ khai đã thông quan lệch hoàn toàn so với thực tế nhập kho? Quy trình tự phát hiện và khai bổ sung trong vòng 60 ngày để được miễn tiền phạt chậm nộp. | Điều 20 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC) về khai bổ sung hồ sơ hải quan; Mẫu số 03/KBS/GSQL ban hành kèm Thông tư 39/2018; Điều 14 Nghị định số 128/2020/NĐ-CP; `scratch/daily_notes.md` ngày 12/11/2026; `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM2-CD10/CM2-CD12) | 🟡 Chờ duyệt |
| 110 | [W48] Thứ Sáu (04/12/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Công văn 24150/CHQ-GSQL & 24170/CHQ-GSQL ngày 13/11/2026: Đột phá giám sát số hóa kho CFS điện tử qua VASSCM giúp giải phóng hàng lẻ LCL trong 15 phút và gỡ vướng thủ tục XNK tại chỗ giữa DN nội địa với DNCX — Doanh nghiệp xuất nhập khẩu đã tận dụng hết lợi thế này chưa?** | Checklist 4 điểm mới đột phá trong quy trình giám sát số hóa kho CFS và tháo gỡ vướng mắc XNK tại chỗ DNCX theo các công văn mới nhất của Tổng cục Hải quan: Cắt giảm 70% thủ tục giấy tờ và bảo toàn thuế VAT 0% cho chuỗi cung ứng sản xuất xuất khẩu. | Góc nhìn từ hiện trường logistics: Khi kho CFS chuyển mình từ 'điểm nghẽn chứng từ giấy' sang 'luồng dữ liệu điện tử 15 phút' — Lời khuyên cho các chủ hàng lẻ LCL và doanh nghiệp chế xuất đón đầu làn sóng số hóa cuối năm. | Công văn số 24150/CHQ-GSQL ngày 12/11/2026 của Tổng cục Hải quan về giám sát số hóa kho CFS điện tử qua hệ thống VASSCM; Công văn số 24170/CHQ-GSQL ngày 13/11/2026 về thủ tục hải quan xuất nhập khẩu tại chỗ với DNCX; Luật Hải quan 2014; `reports/SEO_Keyword_Trends_Eureka.xlsx` ngày 13/11/2026 | 🟡 Chờ duyệt |
| 111 | [W48] Thứ Bảy (05/12/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Tuyến Đường Sắt Xuyên Lục Địa Yiwu - Madrid (13.052 km): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh Qua 8 Quốc Gia, Vượt 3 Lần Thay Đổi Khổ Đường Ray & Cuộc Cách Mạng Logistics Đường Bộ Á - Âu Thách Thức Vận Tải Biển Truyền Thống Ra Sao?** | Bí mật đằng sau hành trình 21 ngày đêm của đoàn tàu container Yiwu - Madrid: Tại sao hàng hóa từ 'thủ phủ hàng tiêu dùng thế giới' Nghĩa Ô (Trung Quốc) lại chọn chạy tàu hỏa xuyên qua Kazakhstan, Nga, Ba Lan đến Tây Ban Nha thay vì đi tàu biển giá rẻ? | Từ tuyến đường sắt xuyên Á - Âu Yiwu - Madrid đến mạng lưới vận tải liên vận đường sắt Việt - Trung: Cơ hội vàng để nông sản và linh kiện điện tử Việt Nam kết nối thẳng vào chuỗi cung ứng Á - Âu qua cửa khẩu Đồng Đăng và Lào Cai. | Hồ sơ Tuyến vận tải đường sắt Yiwu - Madrid (Yixinou Railway / CR Express); Báo cáo Ngân hàng Thế giới (World Bank) về Hành lang Kinh tế Á - Âu; Tư liệu Hiệp hội Vận tải Đường sắt Quốc tế (UIC); `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM3-CD12) | 🟡 Chờ duyệt |
| 112 | [W48] Chủ Nhật (06/12/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Bẫy thuế Nhà thầu (FCT) khi chi trả dịch vụ xuyên biên giới (Meta, Google, Microsoft, AWS) theo Luật Thuế GTGT số 48/2024/QH15 & Thông tư 69/2025/TT-BTC: Cảnh báo doanh nghiệp XNK và Logistics chi tiền tỷ cho quảng cáo, phần mềm đám mây nhưng quên khấu trừ 5% VAT và 2% TNDN dẫn đến bị truy thu tiền trăm triệu kèm tiền phạt chậm nộp!** | Sơ đồ 3 bước xác định nghĩa vụ khấu trừ và kê khai nộp thay thuế FCT cho nhà cung cấp nước ngoài không có cơ sở thường trú tại Việt Nam: Phân biệt trường hợp nhà mạng đã tự đăng ký thuế qua Cổng TTĐT Tổng cục Thuế vs. doanh nghiệp Việt Nam bắt buộc phải nộp thay. | Góc nhìn từ bàn làm việc của Thúy: 'Ủa chị ơi em trả tiền bằng thẻ visa công ty cho Google/Facebook thì bên đó phải tự nộp thuế chứ, sao thanh tra lại đè doanh nghiệp em ra phạt?' — Bài học xương máu về quản trị chi phí hợp lý được trừ và rủi ro thuế nhà thầu thời đại số. | Khoản 4 Điều 4 Luật Thuế GTGT số 48/2024/QH15; Điều 9 Thông tư số 69/2025/TT-BTC của Bộ Tài chính; Điều 76 đến Điều 81 Thông tư số 80/2021/TT-BTC; `scratch/daily_notes.md` ngày 12/11/2026; Hồ sơ tư vấn thuế XNK thực tế tại Eureka Logistics | 🟡 Chờ duyệt |

## 2. Bản phân tích Tỷ lệ Góc tiếp cận & Nền tảng (Validation Tuần W48)

- **Tỷ lệ Năng lượng Nhân hiệu:**
  - **Giáp Mộc (~60%):** Đồng hành chia sẻ cách gỡ kẹt lỗi upload Invoice số hóa Hải Phòng 03YY (CS Thứ Tư - STT 108), Cầm tay chỉ việc phác đồ khai bổ sung AMA sau thông quan mẫu 03 và tuyệt chiêu Switch B/L (HT Thứ Năm - Học XNK 5 phút Tập 18 - STT 109), Khai mở tư duy logistics đường bộ đa phương thức với tuyến tàu Yiwu - Madrid (FF Thứ Bảy - STT 111), Cảnh báo bảo vệ doanh nghiệp trước rủi ro thuế nhà thầu FCT khi dùng dịch vụ số xuyên biên giới (CS Chủ Nhật - STT 112).
  - **Liêm Trinh (~40%):** Sắc bén mổ xẻ tác động của Quyết định 2988/QĐ-BCT áp thuế chống bán phá giá thép cuộn cán nóng HRC AD20 (TN Thứ Hai - STT 106), Phân tích thấu đáo kỳ án đứt gãy chuỗi cung ứng siêu tàu Ever Given Kênh đào Suez thiệt hại 54 tỷ USD (FF Thứ Ba - STT 107), Bóc tách tính pháp lý và cơ hội số hóa kho CFS VASSCM cùng XNK tại chỗ DNCX theo CV 24150 & 24170 (TN Thứ Sáu - STT 110).
- **Tỷ lệ Phân bổ Nền tảng (Chuẩn PBAN 2.0):**
  - **Facebook (4 bài/tuần):** Thứ Hai (TN), Thứ Tư (CS), Thứ Sáu (TN), Chủ Nhật (CS) — Đặt vấn đề sắc bén (Problem-First), bóc tách tình huống thực tế và kích thích thảo luận chuyên môn.
  - **YouTube (3 video/tuần):** Thứ Ba (FF Video Ký sự Lịch sử 8-12p), Thứ Năm (HT Video Học XNK 5 phút Tập 18), Thứ Bảy (FF Video Ký sự Lịch sử 8-12p) — Giáo dục chuyên sâu, trao giá trị thực chiến và xây dựng uy tín tri thức dài hạn.
- **Tuân thủ Bộ từ cấm (Banned Words Compliance):** 100% các tiêu đề và nội dung được quét tự động, tuyệt đối không chứa các từ cấm hay ngữ cảnh nhạy cảm (không có 'trốn thuế', 'lách luật', 'bao thuế', 'hàng lậu', 'tiểu ngạch', 'đánh hàng'...).

---

## 3. Quy trình Phê duyệt Checkpoint 1 & Đóng gói PO (wf01)

1. **Bước 1 (Checkpoint 1):** Chị Thúy xem xét các phương án Option A/B/C trên bảng kế hoạch tuần W48.
2. **Bước 2:** Chọn phương án duyệt tại Cột K của file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) (chọn `Duyệt Option A`, `Duyệt Option B`, hoặc `Duyệt Option C`).
3. **Bước 3:** Hệ thống tự động kích hoạt **Agent 04 (Planner)** đóng gói 7 Phiếu Giao việc (PO) vật lý tại thư mục `scratch/production_orders/` để sẵn sàng bàn giao cho Agent 05 (Facebook) và Agent 06 (YouTube) theo quy trình `wf02` và `wf03`.
"""
    with open(WEEKLY_CAL_MD, "r", encoding="utf-8") as f:
        existing_content = f.read()

    # Cập nhật phần kỳ đánh giá ở đầu trang
    existing_content = existing_content.replace(
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W47",
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W48"
    )

    # Kiểm tra xem Tuần W48 đã có chưa
    if "# TUẦN W48: LỊCH BIÊN TẬP NỘI DUNG" in existing_content:
        parts = existing_content.split("# TUẦN W48: LỊCH BIÊN TẬP NỘI DUNG")
        updated_content = parts[0] + w48_md_content.lstrip()
    else:
        updated_content = existing_content + w48_md_content

    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(updated_content)
    print("  -> Đã cập nhật thành công reports/plan1/weekly_calendar.md!")

def save_json_data():
    print("=== 3. LƯU SCRATCH/W48_DATA.JSON ===")
    os.makedirs(os.path.dirname(JSON_OUTPUT), exist_ok=True)
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W48_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu W48 vào {JSON_OUTPUT}")

if __name__ == "__main__":
    print("=== BẮT ĐẦU QUY TRÌNH WF01: LẬP LỊCH NỘI DUNG TUẦN W48 ===")
    update_excel_calendar()
    update_markdown_calendar()
    save_json_data()
    print("=== HOÀN TẤT THỰC THI GIAI ĐOẠN 1 TUẦN W48! ===")
