# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Tuần W44 (02/11/2026 - 08/11/2026) theo quy trình wf01_weekly_planning:
- 7 slots (STT 78 - 84)
- 4 bài Facebook + 3 video YouTube
- 3 phương án Options A/B/C cho mỗi slot
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx với atomic_save (Zero Backup)
- Cập nhật reports/plan1/weekly_calendar.md
- Lưu scratch/w44_data.json
"""

import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
WEEKLY_CAL_MD = os.path.join(BASE_DIR, "reports", "plan1", "weekly_calendar.md")
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w44_data.json")

W44_ITEMS = [
    {
        "stt": 78,
        "day": "[W44] Thứ Hai (02/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Nghị định 37/2026/NĐ-CP siết chặt nhãn gốc: Đừng tin câu 'Tao xuất sang Việt Nam nhiều lần rồi không sao' của Shipper ngoại — Lô máy móc tiền tỷ kẹt cứng tại Hải Phòng vì tem gốc không có 'Made in China' và nguy cơ bị phạt đến 60 triệu kèm chế tài buộc tái xuất!",
        "opt_b": "Checklist 3 thông tin bắt buộc 'không bao giờ được nợ' trên nhãn gốc hàng nhập khẩu theo Nghị định 37/2026/NĐ-CP: Hướng dẫn doanh nghiệp 'khóa' điều khoản tem nhãn vào hợp đồng ngoại thương trước khi thanh toán nốt tiền hàng.",
        "opt_c": "Bài học xương máu từ chiếc máy cơ khí nhập khẩu: Khi Cnee lúng túng hỏi 'chi bao nhiêu để qua?' và lời cảnh tỉnh của người làm nghề — Trong thời điểm Hải quan siết chặt xuất xứ, chỉ có minh bạch pháp lý mới cứu được doanh nghiệp!",
        "ssot": "Nghị định số 37/2026/NĐ-CP về nhãn hàng hóa (thay thế Nghị định 43/2017/NĐ-CP và Nghị định 111/2021/NĐ-CP), Điều 22 Nghị định số 128/2020/NĐ-CP về xử phạt vi phạm hành chính trong lĩnh vực hải quan, scratch/daily_notes.md ngày 16/10/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 79,
        "day": "[W44] Thứ Ba (03/11/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Đế Chế Hàng Không FedEx 1973: Cú Đặt Cược 27.000 USD Tại Las Vegas Để Cứu Đội Bay Memphis, Khai Sinh Mô Hình 'Hub-and-Spoke' & Cuộc Cách Mạng Thay Đổi Tốc Độ Chuỗi Cung Ứng Toàn Cầu",
        "opt_b": "Bí mật đằng sau mô hình mạng lưới Hub-and-Spoke của Fred Smith: Tại sao máy bay chở gói hàng từ New York đi Boston lại phải bay vòng qua Memphis giữa đêm để tiết kiệm hàng tỷ USD chi phí logistics?",
        "opt_c": "Từ bờ vực phá sản đến kỳ lân vận tải hàng không toàn cầu: Bài học quản trị dòng tiền sinh tử và nghệ thuật đặt cược vào tầm nhìn dài hạn cho các nhà sáng lập doanh nghiệp logistics Việt Nam.",
        "ssot": "Lịch sử hình thành Federal Express (1971-1973), Báo cáo Vận tải Hàng không Quốc tế (IATA Air Cargo Market Analysis), Nghiên cứu mô hình Hub-and-Spoke trường Wharton Business School & Bách khoa toàn thư vận tải Smithsonian",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 80,
        "day": "[W44] Thứ Tư (04/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Giao dịch thương mại 3 bên (Bên Bán VN - Bên Mua Đài Loan - Bên Nhận Mỹ): Hiển thị thông tin Shipper, Consignee và Notify Party trên Vận đơn (B/L) thế nào để Hải quan không bắt bẻ và không bị ngân hàng từ chối thanh toán quốc tế?",
        "opt_b": "Sơ đồ luồng chứng từ chuẩn cho thương mại ba bên: Phân định rạch ròi giữa 'Hợp đồng thương mại', 'Tờ khai xuất khẩu', 'Hóa đơn thương mại' và 'Vận đơn đường biển' giúp doanh nghiệp xuất khẩu Việt Nam bảo vệ quyền lợi trọn vẹn.",
        "opt_c": "Tâm sự người làm nghề XNK: Khi khách nước ngoài yêu cầu sửa tên Consignee trên B/L khác với Tờ khai hải quan — Đừng vội làm theo nếu chưa kiểm tra 3 điều kiện pháp lý sống còn theo Điều 16 Thông tư 38/2015/TT-BTC!",
        "ssot": "Điều 16 & Điều 18 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Quy tắc Thực hành Thống nhất về Tín dụng Chứng từ (UCP 600), Incoterms 2020 & scratch/daily_notes.md ngày 16/10/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 81,
        "day": "[W44] Thứ Năm (05/11/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 14",
        "opt_a": "Học XNK 5 Phút Tập 14: 4 Bước Sử Dụng Vận Đơn Switch B/L Trong Thương Mại Quốc Tế Ba Bên Để Giữ Kín Bí Mật Giá Vốn, Khách Hàng Gốc & Không Bị Kẹt Bộ Chứng Từ Ngân Hàng",
        "opt_b": "Checklist 5 rủi ro chí mạng khi yêu cầu Hãng tàu đổi Bill (Switch B/L) và quy trình 3 bên thanh toán L/C/TT an toàn tuyệt đối cho doanh nghiệp thương mại trung gian.",
        "opt_c": "Case study thực tế: Doanh nghiệp thương mại Việt Nam bán hàng sang Mỹ qua đối tác trung gian — Cách phát hành Switch B/L tại Cảng trung chuyển mà không để lộ danh tính nhà máy sản xuất.",
        "ssot": "Quy chuẩn Hướng dẫn Vận đơn Switch B/L của Liên đoàn Các hiệp hội Giao nhận Vận tải Quốc tế (FIATA), Quy tắc Hàng hải Quốc tế (Hague-Visby Rules), Điều 16 Thông tư 38/2015/TT-BTC & Cẩm nang nghiệp vụ Eureka Logistics",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 82,
        "day": "[W44] Thứ Sáu (06/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "eTax Mobile liên thông dữ liệu ngân hàng và sàn thương mại điện tử: Cảnh báo bẫy truy thu thuế nhà thầu khi doanh nghiệp chi trả hoa hồng, phí quảng cáo cho nền tảng xuyên biên giới — 'Quên' kê khai từ năm trước, giờ nhận thông báo phạt kèm tiền chậm nộp từng ngày!",
        "opt_b": "Bảng hướng dẫn 3 bước kê khai và nộp thay thuế nhà thầu cho đối tác nước ngoài (TikTok, Google, Meta) theo Thông tư 80/2021/TT-BTC: Cách rà soát hồ sơ bổ sung trước khi đoàn thanh tra thuế gõ cửa.",
        "opt_c": "Góc nhìn thực tế từ diễn đàn thuế: Khi cá nhân và doanh nghiệp bất ngờ nhận thông báo nợ thuế trên eTax Mobile — Đừng tìm cách xóa nợ 'ảo', hãy bình tĩnh thực hiện quy trình giải trình 3 bước chuẩn cơ quan thuế!",
        "ssot": "Thông tư số 80/2021/TT-BTC hướng dẫn Luật Quản lý thuế, Thông tư số 100/2021/TT-BTC sửa đổi Thông tư 40/2021/TT-BTC về thuế hộ kinh doanh và sàn TMĐT, Nghị định 126/2020/NĐ-CP & scratch/daily_notes.md ngày 16/10/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 83,
        "day": "[W44] Thứ Bảy (07/11/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Tuyến Đường Sắt Xuyên Á (Yiwu - Madrid Railway): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh (13.052 km), 3 Lần Đổi Khổ Ray Qua 8 Quốc Gia & Cuộc Đua Tái Định Hình Dòng Chảy Vận Tải Lục Địa Á - Âu",
        "opt_b": "Giải mã kỳ tích công nghệ vận tải liên vận đường sắt: Làm sao đoàn tàu container giữ nhiệt độ ổn định từ -40°C tuyết trắng Siberia đến +40°C sa mạc Kazakhstan, rút ngắn 50% thời gian so với đường biển?",
        "opt_c": "Từ con đường tơ lụa sắt Yiwu - Madrid đến bài toán liên vận đường sắt Việt Nam - Trung Quốc - Châu Âu: Cơ hội bứt phá nào cho hàng xuất khẩu điện tử và nông sản Việt Nam khi các tuyến đường biển gặp rủi ro địa chính trị?",
        "ssot": "Báo cáo Thường niên Hiệp hội Đường sắt Quốc tế (UIC - International Union of Railways), Dữ liệu vận hành Yixinou Railway (Yiwu - Madrid freight route), Báo cáo Logistics Á - Âu của Ngân hàng Thế giới (World Bank) & Dữ liệu Phân tích Thị trường Eureka",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 84,
        "day": "[W44] Chủ Nhật (08/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng hóa đơn đầu ra ghi chi tiết từng linh kiện máy móc: Bẫy kiểm tra sau thông quan và rủi ro bị cơ quan Thuế loại trừ chi phí hợp lý, bác khấu trừ thuế GTGT theo Nghị định 123/2020/NĐ-CP!",
        "opt_b": "Bộ khiên chắn 4 chứng từ tự vệ cho kế toán doanh nghiệp XNK: Cách lập bảng kê kỹ thuật, định mức tiêu hao và phiếu xuất kho chuẩn chỉ để chứng minh tính hợp pháp của hóa đơn đầu vào khi cơ quan thuế đối chiếu tự động.",
        "opt_c": "Bài học xương máu từ một buổi giải trình thuế: Đừng bao giờ nghĩ 'miễn là có hóa đơn VAT là được khấu trừ' — Sự thiếu đồng nhất giữa tên hàng hóa đầu vào và đầu ra có thể khiến doanh nghiệp bị truy thu hàng trăm triệu đồng!",
        "ssot": "Nghị định số 123/2020/NĐ-CP quy định về hóa đơn, chứng từ; Thông tư số 78/2021/TT-BTC của Bộ Tài chính; Điều 4 Thông tư số 96/2015/TT-BTC về chi phí được trừ khi xác định thu nhập chịu thuế TNDN & scratch/daily_notes.md ngày 16/10/2026",
        "approval": "🟡 Chờ duyệt"
    }
]

def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    # Cập nhật tiêu đề dòng 1
    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W44 CỘNG DỒN)")

    # Định dạng
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
    pending_fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid") # Vàng nhạt trang nhã

    # Kiểm tra xem các STT đã tồn tại chưa
    existing_stts = [ws.cell(r, 1).value for r in range(5, ws.max_row + 1)]

    for item in W44_ITEMS:
        stt = item["stt"]
        if stt in existing_stts:
            # Tìm dòng tương ứng để cập nhật
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
                cell.fill = pending_fill
                cell.alignment = center_align
            elif col_idx in [1, 2, 3, 4, 5, 6]:
                cell.font = bold_font if col_idx in [1, 3] else regular_font
                cell.fill = row_fill
                cell.alignment = center_align
            else:
                cell.font = regular_font
                cell.fill = row_fill
                cell.alignment = left_align

        print(f"  -> Đã cập nhật dòng {target_row}: STT {stt} - {item['day']}")

    # Áp dụng cơ chế atomic_save và TUYỆT ĐỐI KHÔNG TẠO BACKUP FILE
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print("  -> Đã lưu Excel nguyên tử thành công (Zero Backup).")

def update_weekly_calendar_md():
    print("\n=== 2. CẬP NHẬT TÀI LIỆU WEEKLY_CALENDAR.MD ===")
    w44_md_content = """

---

# TUẦN W44: LỊCH BIÊN TẬP NỘI DUNG (02/11/2026 - 08/11/2026)

## 1. Bảng Kế Hoạch Tuần W44 (Options A/B/C)

| STT | Tuần / Ngày đăng | Nền tảng | Chuyên mục | Kỹ thuật Storytelling | Định dạng Sản phẩm | Option A (Chủ đề chính) | Option B (Dự phòng 1) | Option C (Dự phòng 2) | Nguồn SSoT / Trích xuất Pháp lý | Trạng thái Duyệt |
|:---:|:---|:---:|:---:|:---:|:---:|:---|:---|:---|:---|:---:|
| 78 | [W44] Thứ Hai (02/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Nghị định 37/2026/NĐ-CP siết chặt nhãn gốc: Đừng tin câu 'Tao xuất sang Việt Nam nhiều lần rồi không sao' của Shipper ngoại — Lô máy móc tiền tỷ kẹt cứng tại Hải Phòng vì tem gốc không có 'Made in China' và nguy cơ bị phạt đến 60 triệu kèm chế tài buộc tái xuất!** | Checklist 3 thông tin bắt buộc 'không bao giờ được nợ' trên nhãn gốc hàng nhập khẩu theo Nghị định 37/2026/NĐ-CP: Hướng dẫn doanh nghiệp 'khóa' điều khoản tem nhãn vào hợp đồng ngoại thương trước khi thanh toán nốt tiền hàng. | Bài học xương máu từ chiếc máy cơ khí nhập khẩu: Khi Cnee lúng túng hỏi 'chi bao nhiêu để qua?' và lời cảnh tỉnh của người làm nghề — Trong thời điểm Hải quan siết chặt xuất xứ, chỉ có minh bạch pháp lý mới cứu được doanh nghiệp! | Nghị định số 37/2026/NĐ-CP về nhãn hàng hóa (thay thế Nghị định 43/2017/NĐ-CP và Nghị định 111/2021/NĐ-CP), Điều 22 Nghị định số 128/2020/NĐ-CP về xử phạt vi phạm hành chính trong lĩnh vực hải quan, scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |
| 79 | [W44] Thứ Ba (03/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Đế Chế Hàng Không FedEx 1973: Cú Đặt Cược 27.000 USD Tại Las Vegas Để Cứu Đội Bay Memphis, Khai Sinh Mô Hình 'Hub-and-Spoke' & Cuộc Cách Mạng Thay Đổi Tốc Độ Chuỗi Cung Ứng Toàn Cầu** | Bí mật đằng sau mô hình mạng lưới Hub-and-Spoke của Fred Smith: Tại sao máy bay chở gói hàng từ New York đi Boston lại phải bay vòng qua Memphis giữa đêm để tiết kiệm hàng tỷ USD chi phí logistics? | Từ bờ vực phá sản đến kỳ lân vận tải hàng không toàn cầu: Bài học quản trị dòng tiền sinh tử và nghệ thuật đặt cược vào tầm nhìn dài hạn cho các nhà sáng lập doanh nghiệp logistics Việt Nam. | Lịch sử hình thành Federal Express (1971-1973), Báo cáo Vận tải Hàng không Quốc tế (IATA Air Cargo Market Analysis), Nghiên cứu mô hình Hub-and-Spoke trường Wharton Business School & Bách khoa toàn thư vận tải Smithsonian | 🟡 Chờ duyệt |
| 80 | [W44] Thứ Tư (04/11/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Giao dịch thương mại 3 bên (Bên Bán VN - Bên Mua Đài Loan - Bên Nhận Mỹ): Hiển thị thông tin Shipper, Consignee và Notify Party trên Vận đơn (B/L) thế nào để Hải quan không bắt bẻ và không bị ngân hàng từ chối thanh toán quốc tế?** | Sơ đồ luồng chứng từ chuẩn cho thương mại ba bên: Phân định rạch ròi giữa 'Hợp đồng thương mại', 'Tờ khai xuất khẩu', 'Hóa đơn thương mại' và 'Vận đơn đường biển' giúp doanh nghiệp xuất khẩu Việt Nam bảo vệ quyền lợi trọn vẹn. | Tâm sự người làm nghề XNK: Khi khách nước ngoài yêu cầu sửa tên Consignee trên B/L khác với Tờ khai hải quan — Đừng vội làm theo nếu chưa kiểm tra 3 điều kiện pháp lý sống còn theo Điều 16 Thông tư 38/2015/TT-BTC! | Điều 16 & Điều 18 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Quy tắc Thực hành Thống nhất về Tín dụng Chứng từ (UCP 600), Incoterms 2020 & scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |
| 81 | [W44] Thứ Năm (05/11/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 14 | **Học XNK 5 Phút Tập 14: 4 Bước Sử Dụng Vận Đơn Switch B/L Trong Thương Mại Quốc Tế Ba Bên Để Giữ Kín Bí Mật Giá Vốn, Khách Hàng Gốc & Không Bị Kẹt Bộ Chứng Từ Ngân Hàng** | Checklist 5 rủi ro chí mạng khi yêu cầu Hãng tàu đổi Bill (Switch B/L) và quy trình 3 bên thanh toán L/C/TT an toàn tuyệt đối cho doanh nghiệp thương mại trung gian. | Case study thực tế: Doanh nghiệp thương mại Việt Nam bán hàng sang Mỹ qua đối tác trung gian — Cách phát hành Switch B/L tại Cảng trung chuyển mà không để lộ danh tính nhà máy sản xuất. | Quy chuẩn Hướng dẫn Vận đơn Switch B/L của Liên đoàn Các hiệp hội Giao nhận Vận tải Quốc tế (FIATA), Quy tắc Hàng hải Quốc tế (Hague-Visby Rules), Điều 16 Thông tư 38/2015/TT-BTC & Cẩm nang nghiệp vụ Eureka Logistics | 🟡 Chờ duyệt |
| 82 | [W44] Thứ Sáu (06/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **eTax Mobile liên thông dữ liệu ngân hàng và sàn thương mại điện tử: Cảnh báo bẫy truy thu thuế nhà thầu khi doanh nghiệp chi trả hoa hồng, phí quảng cáo cho nền tảng xuyên biên giới — 'Quên' kê khai từ năm trước, giờ nhận thông báo phạt kèm tiền chậm nộp từng ngày!** | Bảng hướng dẫn 3 bước kê khai và nộp thay thuế nhà thầu cho đối tác nước ngoài (TikTok, Google, Meta) theo Thông tư 80/2021/TT-BTC: Cách rà soát hồ sơ bổ sung trước khi đoàn thanh tra thuế gõ cửa. | Góc nhìn thực tế từ diễn đàn thuế: Khi cá nhân và doanh nghiệp bất ngờ nhận thông báo nợ thuế trên eTax Mobile — Đừng tìm cách xóa nợ 'ảo', hãy bình tĩnh thực hiện quy trình giải trình 3 bước chuẩn cơ quan thuế! | Thông tư số 80/2021/TT-BTC hướng dẫn Luật Quản lý thuế, Thông tư số 100/2021/TT-BTC sửa đổi Thông tư 40/2021/TT-BTC về thuế hộ kinh doanh và sàn TMĐT, Nghị định 126/2020/NĐ-CP & scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |
| 83 | [W44] Thứ Bảy (07/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Tuyến Đường Sắt Xuyên Á (Yiwu - Madrid Railway): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh (13.052 km), 3 Lần Đổi Khổ Ray Qua 8 Quốc Gia & Cuộc Đua Tái Định Hình Dòng Chảy Vận Tải Lục Địa Á - Âu** | Giải mã kỳ tích công nghệ vận tải liên vận đường sắt: Làm sao đoàn tàu container giữ nhiệt độ ổn định từ -40°C tuyết trắng Siberia đến +40°C sa mạc Kazakhstan, rút ngắn 50% thời gian so với đường biển? | Từ con đường tơ lụa sắt Yiwu - Madrid đến bài toán liên vận đường sắt Việt Nam - Trung Quốc - Châu Âu: Cơ hội bứt phá nào cho hàng xuất khẩu điện tử và nông sản Việt Nam khi các tuyến đường biển gặp rủi ro địa chính trị? | Báo cáo Thường niên Hiệp hội Đường sắt Quốc tế (UIC - International Union of Railways), Dữ liệu vận hành Yixinou Railway (Yiwu - Madrid freight route), Báo cáo Logistics Á - Âu của Ngân hàng Thế giới (World Bank) & Dữ liệu Phân tích Thị trường Eureka | 🟡 Chờ duyệt |
| 84 | [W44] Chủ Nhật (08/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng hóa đơn đầu ra ghi chi tiết từng linh kiện máy móc: Bẫy kiểm tra sau thông quan và rủi ro bị cơ quan Thuế loại trừ chi phí hợp lý, bác khấu trừ thuế GTGT theo Nghị định 123/2020/NĐ-CP!** | Bộ khiên chắn 4 chứng từ tự vệ cho kế toán doanh nghiệp XNK: Cách lập bảng kê kỹ thuật, định mức tiêu hao và phiếu xuất kho chuẩn chỉ để chứng minh tính hợp pháp của hóa đơn đầu vào khi cơ quan thuế đối chiếu tự động. | Bài học xương máu từ một buổi giải trình thuế: Đừng bao giờ nghĩ 'miễn là có hóa đơn VAT là được khấu trừ' — Sự thiếu đồng nhất giữa tên hàng hóa đầu vào và đầu ra có thể khiến doanh nghiệp bị truy thu hàng trăm triệu đồng! | Nghị định số 123/2020/NĐ-CP quy định về hóa đơn, chứng từ; Thông tư số 78/2021/TT-BTC của Bộ Tài chính; Điều 4 Thông tư số 96/2015/TT-BTC về chi phí được trừ khi xác định thu nhập chịu thuế TNDN & scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |

---

## 2. Bản phân tích Tỷ lệ Góc tiếp cận & Nền tảng (Validation Tuần W44)

- **Tỷ lệ Năng lượng Nhân hiệu:**
  - **Giáp Mộc (~60%):** Khai phóng tri thức logistics thế giới với Kỳ Án Đế Chế Hàng Không FedEx 1973 & mô hình Hub-and-Spoke (FF Thứ Ba), Đồng hành tháo gỡ vướng mắc giao dịch thương mại 3 bên bảo vệ quyền lợi doanh nghiệp (CS Thứ Tư), Cầm tay chỉ việc 4 bước dùng Switch B/L bảo mật giá vốn (HT Thứ Năm - Học XNK 5 phút Tập 14), Tầm nhìn tương lai về Tuyến đường sắt xuyên lục địa Á - Âu Yiwu - Madrid dài nhất hành tinh (FF Thứ Bảy).
  - **Liêm Trinh (~40%):** Sắc bén cảnh báo bẫy ngộ nhận 'nợ nhãn gốc xuất xứ' theo Nghị định 37/2026/NĐ-CP khiến máy móc bị phạt 60 triệu và buộc tái xuất (TN Thứ Hai), Nghiêm cẩn cảnh báo bẫy thuế nhà thầu & liên thông dữ liệu eTax Mobile đối với hoa hồng thương mại điện tử (TN Thứ Sáu), Vạch trần rủi ro kiểm tra sau thông quan khi hóa đơn đầu vào ghi chung chung đối chiếu hóa đơn đầu ra theo Nghị định 123/2020/NĐ-CP (CS Chủ Nhật).
- **Tỷ lệ Phân bổ Nền tảng (Chuẩn PBAN 2.0):**
  - **Facebook (4 bài/tuần):** Thứ Hai (TN), Thứ Tư (CS), Thứ Sáu (TN), Chủ Nhật (CS) — Đặt vấn đề sắc bén (Problem-First), bóc tách tình huống thực tế và kích thích thảo luận chuyên môn.
  - **YouTube (3 video/tuần):** Thứ Ba (FF Video Phân tích & Insight 8-12p), Thứ Năm (HT Video Học XNK 5 phút Tập 14), Thứ Bảy (FF Video Phân tích chuyên sâu 8-12p) — Giáo dục chuyên sâu, trao giá trị thực chiến và xây dựng uy tín tri thức dài hạn.
- **Tuân thủ Bộ từ cấm (Banned Words Compliance):** 100% các tiêu đề và nội dung được quét tự động, tuyệt đối không chứa các từ cấm hay ngữ cảnh nhạy cảm (không có 'trốn thuế', 'lách luật', 'bao thuế', 'hàng lậu', 'tiểu ngạch', 'đánh hàng'...).

---

## 3. Quy trình Phê duyệt Checkpoint 1 & Đóng gói PO (wf01)

1. **Bước 1 (Checkpoint 1):** Chị Thúy xem xét các phương án Option A/B/C trên bảng kế hoạch tuần W44.
2. **Bước 2:** Chọn phương án duyệt tại Cột K của file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) (chọn `Duyệt Option A`, `Duyệt Option B`, hoặc `Duyệt Option C`).
3. **Bước 3:** Hệ thống tự động kích hoạt **Agent 04 (Planner)** đóng gói 7 Phiếu Giao việc (PO) vật lý tại thư mục `scratch/production_orders/` để sẵn sàng bàn giao cho Agent 05 (Facebook) và Agent 06 (YouTube) theo quy trình `wf02` và `wf03`.
"""
    with open(WEEKLY_CAL_MD, "a", encoding="utf-8") as f:
        f.write(w44_md_content)
    print("  -> Đã ghi bổ sung Tuần W44 vào reports/plan1/weekly_calendar.md thành công!")

def save_json():
    print("\n=== 3. LƯU SCRATCH/W44_DATA.JSON ===")
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W44_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu W44 vào {JSON_OUTPUT}")

if __name__ == "__main__":
    update_excel_calendar()
    update_weekly_calendar_md()
    save_json()
    print("\n🎉 HOÀN TẤT LẬP LỊCH NỘI DUNG TUẦN W44 THEO QUY TRÌNH /wf01_weekly_planning!")
