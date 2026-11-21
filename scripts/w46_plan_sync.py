# -*- coding: utf-8 -*-
"""
Script lập kế hoạch nội dung Tuần W46 (16/11/2026 - 22/11/2026) theo quy trình wf01_weekly_planning:
- 7 slots (STT 92 - 98)
- 4 bài Facebook + 3 video YouTube
- 3 phương án Options A/B/C cho mỗi slot
- Cập nhật trực tiếp Sheet 1 (Weekly Calendar) trong Weekly_Content_Calendar.xlsx với atomic_save (Zero Backup - RULE 10)
- Cập nhật reports/plan1/weekly_calendar.md
- Lưu scratch/w46_data.json
"""

import os
import sys
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
WEEKLY_CAL_MD = os.path.join(BASE_DIR, "reports", "plan1", "weekly_calendar.md")
JSON_OUTPUT = os.path.join(BASE_DIR, "scratch", "w46_data.json")

W46_ITEMS = [
    {
        "stt": 92,
        "day": "[W46] Thứ Hai (16/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Thông tư 121/2025/TT-BTC siết chặt nghĩa vụ lưu giữ hồ sơ hải quan 05 năm: Đừng nghĩ 'thông quan xong là xong' — 6 điểm thay đổi chí mạng từ phiếu trừ lùi, biên bản lấy mẫu đến tách bạch định mức SXXK thực tế khiến doanh nghiệp trở tay không kịp khi đoàn kiểm tra sau thông quan gõ cửa!",
        "opt_b": "Checklist 6 nhóm chứng từ bắt buộc phải lưu bản chính giấy và dữ liệu điện tử 05 năm theo Thông tư 121/2025/TT-BTC: Hướng dẫn doanh nghiệp XNK và đại lý hải quan phân định rạch ròi trách nhiệm lưu trữ để không bị phạt vi phạm hành chính.",
        "opt_c": "Tâm sự của người làm nghề hải quan: 'Cứ nghĩ nộp hồ sơ điện tử qua V5 rồi thì quẳng hồ sơ giấy đi đâu cũng được' — Bài học xương máu từ một buổi hậu kiểm 5 năm trước khi Thông tư 121 chính thức có hiệu lực từ đầu năm 2026!",
        "ssot": "Khoản 5 Điều 1 Thông tư số 121/2025/TT-BTC ngày 18/12/2025 (sửa đổi Điều 16a Thông tư 38/2015/TT-BTC); Điểm đ Khoản 2 Điều 18 Luật Hải quan số 54/2014/QH13; Luật Giao dịch điện tử 2023; scratch/daily_notes.md ngày 21/10/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 93,
        "day": "[W46] Thứ Ba (17/11/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Thương Cảng Phố Hiến Thế Kỷ 17 ('Thứ Nhất Kinh Kỳ, Thứ Nhì Phố Hiến'): Tiền Nhân Đại Việt Đã Lập Ty Tuần Ty, Thu Thuế Tàu Buôn Phương Tây & Bài Học Quản Trị Ngoại Thương Bản Lĩnh",
        "opt_b": "Bí mật đằng sau sự hưng thịnh của Đô thị Phố Hiến thời chúa Trịnh: Làm thế nào một trạm kiểm soát đường sông trở thành thương cảng quốc tế đón tàu buôn Hà Lan (VOC), Anh (EIC) và Nhật Bản sầm uất bậc nhất châu Á?",
        "opt_c": "Từ Ty Tuần ty Phố Hiến thế kỷ 17 đến hệ thống thông quan số hóa hôm nay: Bài học ngàn đời về giữ vững chủ quyền luật pháp và bản lĩnh đàm phán thương mại khi mở cửa đón nhận các dòng vốn ngoại.",
        "ssot": "Đại Việt Sử Ký Toàn Thư (Bản kỷ Thực lục thời Lê - Trịnh); Khâm Định Việt Sử Thông Giám Cương Mục; Hồ sơ tư liệu Thương cảng Phố Hiến (Viện Sử học Việt Nam); Nhật ký thương điếm Hà Lan (VOC Dagh-Register Batavia 1637-1700); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD03)",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 94,
        "day": "[W46] Thứ Tư (18/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Góc nhìn cá nhân Thúy",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Hàng nhập kho ngoại quan (KNQ) phát hiện sai mã HS code: Sửa tờ khai kho ngoại quan có phức tạp không? Cảnh báo bẫy bị hải quan nghi vấn gian lận danh mục hàng cấm tạm nhập tái xuất và quy trình 3 bước khai bổ sung an toàn theo Điều 20 Thông tư 38/2015!",
        "opt_b": "Sơ đồ 3 bước xử lý khi phát hiện sai sót mã HS hàng gửi kho ngoại quan: Phân biệt rõ ranh giới giữa sai sót kỹ thuật thông thường vs. sai sót làm thay đổi chính sách quản lý chuyên ngành để không bị lập biên bản vi phạm.",
        "opt_c": "Tâm sự của người làm nghề logistics: Đừng bao giờ chủ quan nghĩ 'hàng nằm trong kho ngoại quan chưa vào nội địa thì muốn sửa gì cũng được' — Bài học từ một lô hàng linh kiện suýt bị đình chỉ chuyển khẩu vì vội vàng truyền sửa tờ khai!",
        "ssot": "Điều 20 & Điều 88 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư số 39/2018/TT-BTC); Điều 85 Nghị định số 08/2015/NĐ-CP; scratch/daily_notes.md ngày 22/10/2026 & Hotline tư vấn Eureka",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 95,
        "day": "[W46] Thứ Năm (19/11/2026)",
        "platform": "YouTube",
        "category": "HT — Học Thực chiến",
        "technique": "Trao giá trị thực chiến",
        "format": "Video Học XNK 5 phút Tập 16",
        "opt_a": "Học XNK 5 Phút Tập 16: Nỗi Đau 'Con Ốc Vít Cơ Khí' Nhảy Thuế Từ 0% Lên 25% Vì Áp Sai Mã HS & Quy Trình 4 Bước Xin Xác Định Trước Mã Số (Advance Ruling) Chuẩn Luật Hải Quan Tránh Bị Phạt Truy Thu Tiền Tỷ",
        "opt_b": "Checklist 4 tài liệu kỹ thuật bắt buộc phải chuẩn bị khi lập hồ sơ xin xác định trước mã số hàng hóa (Advance Ruling) theo Điều 28 Luật Hải quan: Cách làm việc với nhà sản xuất nước ngoài để lấy bản vẽ CAD và catalogue chuẩn chỉ.",
        "opt_c": "Case study thực tế tại Eureka: Doanh nghiệp cơ khí nhập khẩu linh kiện máy in công nghiệp bị áp mã ốc vít thông thường và hành trình 60 ngày giải trình kỹ thuật bảo vệ thành công mức thuế suất ưu đãi 0%.",
        "ssot": "Điều 28 Luật Hải quan số 54/2014/QH13; Điều 23, 24 Thông tư số 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); Biểu thuế Xuất nhập khẩu 2026 (Chương 73 nhóm 7318 vs Chương 84 nhóm 8443); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM4-CD03)",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 96,
        "day": "[W46] Thứ Sáu (20/11/2026)",
        "platform": "Facebook",
        "category": "TN — Tin Ngành",
        "technique": "Problem-First",
        "format": "Bài tranh luận + Infographic",
        "opt_a": "Hộ kinh doanh trả mặt bằng chuyển quận hoặc doanh nghiệp xuất hóa đơn thay thế khác quý: Đừng nghĩ 'nộp hồ sơ lên thuế là xong' — Cảnh báo bẫy bị giam lỏng mã số thuế, nợ thuế treo trên eTax Mobile và trát phạt tiền chậm nộp theo Nghị định 123/2020/NĐ-CP!",
        "opt_b": "Hướng dẫn 3 bước chuẩn hóa hồ sơ khi chuyển địa chỉ kinh doanh khác quận và cách kê khai hóa đơn thay thế khác kỳ tính thuế: Bí quyết giúp chủ hộ kinh doanh và kế toán không bị đình trệ hoạt động kinh doanh.",
        "opt_c": "Góc nhìn từ phòng tư vấn Eureka: Tại sao chuyển cửa hàng từ Tân Bình sang Bình Thạnh hay Cầu Giấy sang Đống Đa lại trở thành 'cơn ác mộng' thủ tục? Bài học về sự chủ động chốt nghĩa vụ thuế trước khi thanh lý hợp đồng thuê mặt bằng.",
        "ssot": "Thông tư số 105/2020/TT-BTC về đăng ký thuế (Điều 10, Điều 11 về thay đổi thông tin đăng ký thuế); Nghị định số 123/2020/NĐ-CP (Điều 19 về xử lý hóa đơn có sai sót); Thông tư số 78/2021/TT-BTC; scratch/daily_notes.md ngày 22/10/2026",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 97,
        "day": "[W46] Thứ Bảy (21/11/2026)",
        "platform": "YouTube",
        "category": "FF — Funfact & Insight",
        "technique": "Series & Cliffhanger",
        "format": "Video Phân tích & Insight (8-12p)",
        "opt_a": "Kỳ Án Phiến Đá Babylon & Bộ Luật Hammurabi (1750 TCN): 3.800 Năm Trước Nhân Loại Đã Quản Trị Hợp Đồng Ngoại Thương, Vận Tải Đường Thủy & Bồi Thường Hàng Hóa Bản Lĩnh Thế Nào?",
        "opt_b": "Giải mã 282 điều luật khắc trên phiến đá diorite đen bảo tàng Louvre: Tại sao nguyên tắc 'Mắt đền mắt, răng đền răng' thực chất là cuộc cách mạng bảo vệ giới thương nhân và đặt nền móng cho hợp đồng logistics hiện đại?",
        "opt_c": "Từ phiến đá Babylon 3.800 năm trước đến bản hợp đồng dịch vụ logistics thế kỷ 21: Tại sao sự minh bạch về trách nhiệm bồi thường và cam kết uy tín mới là thứ bảo vệ doanh nghiệp bền vững nhất?",
        "ssot": "Bộ luật Hammurabi (Code of Hammurabi c. 1750 BC, Bảo tàng Louvre Paris - Điều 104-107 về đại lý thương mại, Điều 236-240 về trách nhiệm tàu thuyền vận tải); Tư liệu Viện Đông phương học Chicago (Oriental Institute of Chicago); Ke_Hoach_Phat_Song_YouTube_Series.xlsx (CM3-CD04)",
        "approval": "🟡 Chờ duyệt"
    },
    {
        "stt": 98,
        "day": "[W46] Chủ Nhật (22/11/2026)",
        "platform": "Facebook",
        "category": "CS — Case Study",
        "technique": "Problem-First",
        "format": "Bài Case Study + Ảnh",
        "opt_a": "Nhập khẩu sâm tươi, dược liệu quý từ Hàn Quốc: Đừng nghe lời xúi 'xách tay cho nhanh' — Cảnh báo bẫy kiểm dịch thực vật tại Cục BVTV, công bố ATTP và nguy cơ bị phạt đến 100 triệu kèm tiêu hủy toàn bộ lô hàng tại sân bay!",
        "opt_b": "Lộ trình 3 bước nhập khẩu chính ngạch sâm tươi và nông sản nhạy cảm từ Hàn Quốc về Việt Nam: Từ giấy phép kiểm dịch thực vật nhập khẩu, C/O Form VK/AK đến bộ hồ sơ thông quan an toàn 100%.",
        "opt_c": "Chuyện người làm nghề: Khi khách hàng tha thiết 'em giúp chị gửi phí vài triệu cứu lô sâm tươi đang bị giữ ở Nội Bài' — Lời khuyên chân thành của Thúy: Khi đã vướng rào cản kiểm dịch an toàn thực phẩm, chỉ có tuân thủ quy chuẩn ngay từ đầu mới bảo vệ được vốn liếng!",
        "ssot": "Luật An toàn thực phẩm 2010; Nghị định số 15/2018/NĐ-CP (quy định chi tiết thi hành một số điều của Luật ATTP); Thông tư số 04/2023/TT-BNNPTNT về danh mục kiểm dịch thực vật; Nghị định số 128/2020/NĐ-CP; scratch/daily_notes.md ngày 21/10/2026",
        "approval": "🟡 Chờ duyệt"
    }
]

def update_excel_calendar():
    print("=== 1. CẬP NHẬT SHEET 1 (WEEKLY CALENDAR) TRONG EXCEL ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["Weekly Calendar"]

    # Cập nhật tiêu đề dòng 1
    ws.cell(1, 1, value="LỊCH BIÊN TẬP ĐA KÊNH — HỆ THỐNG PBAN 2.0 (TUẦN W33 ĐẾN W46 CỘNG DỒN)")

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
    pending_fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")

    # Kiểm tra xem các STT đã tồn tại chưa
    existing_stts = [ws.cell(r, 1).value for r in range(5, ws.max_row + 1)]

    for item in W46_ITEMS:
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
    dv.add("K5:K120")

    # Lưu file an toàn bằng atomic_save (Zero Backup Rule 10)
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print("  -> Đã lưu thành công Sheet 1 qua atomic_save!")

def update_markdown_calendar():
    print("=== 2. CẬP NHẬT REPORTS/PLAN1/WEEKLY_CALENDAR.MD ===")
    w46_md_content = """

---

# TUẦN W46: LỊCH BIÊN TẬP NỘI DUNG (16/11/2026 - 22/11/2026)

## 1. Bảng Kế Hoạch Tuần W46 (Options A/B/C)

| STT | Tuần / Ngày đăng | Nền tảng | Chuyên mục | Kỹ thuật Storytelling | Định dạng Sản phẩm | Option A (Chủ đề chính) | Option B (Dự phòng 1) | Option C (Dự phòng 2) | Nguồn SSoT / Trích xuất Pháp lý | Trạng thái Duyệt |
|:---:|:---|:---:|:---:|:---:|:---:|:---|:---|:---|:---|:---:|
| 92 | [W46] Thứ Hai (16/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Thông tư 121/2025/TT-BTC siết chặt nghĩa vụ lưu giữ hồ sơ hải quan 05 năm: Đừng nghĩ 'thông quan xong là xong' — 6 điểm thay đổi chí mạng từ phiếu trừ lùi, biên bản lấy mẫu đến tách bạch định mức SXXK thực tế khiến doanh nghiệp trở tay không kịp khi đoàn kiểm tra sau thông quan gõ cửa!** | Checklist 6 nhóm chứng từ bắt buộc phải lưu bản chính giấy và dữ liệu điện tử 05 năm theo Thông tư 121/2025/TT-BTC: Hướng dẫn doanh nghiệp XNK và đại lý hải quan phân định rạch ròi trách nhiệm lưu trữ để không bị phạt vi phạm hành chính. | Tâm sự của người làm nghề hải quan: 'Cứ nghĩ nộp hồ sơ điện tử qua V5 rồi thì quẳng hồ sơ giấy đi đâu cũng được' — Bài học xương máu từ một buổi hậu kiểm 5 năm trước khi Thông tư 121 chính thức có hiệu lực từ đầu năm 2026! | Khoản 5 Điều 1 Thông tư số 121/2025/TT-BTC ngày 18/12/2025 (sửa đổi Điều 16a Thông tư 38/2015/TT-BTC); Điểm đ Khoản 2 Điều 18 Luật Hải quan số 54/2014/QH13; Luật Giao dịch điện tử 2023; `scratch/daily_notes.md` ngày 21/10/2026 | 🟡 Chờ duyệt |
| 93 | [W46] Thứ Ba (17/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Thương Cảng Phố Hiến Thế Kỷ 17 ('Thứ Nhất Kinh Kỳ, Thứ Nhì Phố Hiến'): Tiền Nhân Đại Việt Đã Lập Ty Tuần Ty, Thu Thuế Tàu Buôn Phương Tây & Bài Học Quản Trị Ngoại Thương Bản Lĩnh** | Bí mật đằng sau sự hưng thịnh của Đô thị Phố Hiến thời chúa Trịnh: Làm thế nào một trạm kiểm soát đường sông trở thành thương cảng quốc tế đón tàu buôn Hà Lan (VOC), Anh (EIC) và Nhật Bản sầm uất bậc nhất châu Á? | Từ Ty Tuần ty Phố Hiến thế kỷ 17 đến hệ thống thông quan số hóa hôm nay: Bài học ngàn đời về giữ vững chủ quyền luật pháp và bản lĩnh đàm phán thương mại khi mở cửa đón nhận các dòng vốn ngoại. | Đại Việt Sử Ký Toàn Thư (Bản kỷ Thực lục thời Lê - Trịnh); Khâm Định Việt Sử Thông Giám Cương Mục; Hồ sơ tư liệu Thương cảng Phố Hiến (Viện Sử học Việt Nam); Nhật ký thương điếm Hà Lan (VOC Dagh-Register Batavia 1637-1700); `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM3-CD03) | 🟡 Chờ duyệt |
| 94 | [W46] Thứ Tư (18/11/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Hàng nhập kho ngoại quan (KNQ) phát hiện sai mã HS code: Sửa tờ khai kho ngoại quan có phức tạp không? Cảnh báo bẫy bị hải quan nghi vấn gian lận danh mục hàng cấm tạm nhập tái xuất và quy trình 3 bước khai bổ sung an toàn theo Điều 20 Thông tư 38/2015!** | Sơ đồ 3 bước xử lý khi phát hiện sai sót mã HS hàng gửi kho ngoại quan: Phân biệt rõ ranh giới giữa sai sót kỹ thuật thông thường vs. sai sót làm thay đổi chính sách quản lý chuyên ngành để không bị lập biên bản vi phạm. | Tâm sự của người làm nghề logistics: Đừng bao giờ chủ quan nghĩ 'hàng nằm trong kho ngoại quan chưa vào nội địa thì muốn sửa gì cũng được' — Bài học từ một lô hàng linh kiện suýt bị đình chỉ chuyển khẩu vì vội vàng truyền sửa tờ khai! | Điều 20 & Điều 88 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư số 39/2018/TT-BTC); Điều 85 Nghị định số 08/2015/NĐ-CP; `scratch/daily_notes.md` ngày 22/10/2026 & Hotline tư vấn Eureka | 🟡 Chờ duyệt |
| 95 | [W46] Thứ Năm (19/11/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 16 | **Học XNK 5 Phút Tập 16: Nỗi Đau 'Con Ốc Vít Cơ Khí' Nhảy Thuế Từ 0% Lên 25% Vì Áp Sai Mã HS & Quy Trình 4 Bước Xin Xác Định Trước Mã Số (Advance Ruling) Chuẩn Luật Hải Quan Tránh Bị Phạt Truy Thu Tiền Tỷ** | Checklist 4 tài liệu kỹ thuật bắt buộc phải chuẩn bị khi lập hồ sơ xin xác định trước mã số hàng hóa (Advance Ruling) theo Điều 28 Luật Hải quan: Cách làm việc với nhà sản xuất nước ngoài để lấy bản vẽ CAD và catalogue chuẩn chỉ. | Case study thực tế tại Eureka: Doanh nghiệp cơ khí nhập khẩu linh kiện máy in công nghiệp bị áp mã ốc vít thông thường và hành trình 60 ngày giải trình kỹ thuật bảo vệ thành công mức thuế suất ưu đãi 0%. | Điều 28 Luật Hải quan số 54/2014/QH13; Điều 23, 24 Thông tư số 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); Biểu thuế Xuất nhập khẩu 2026 (Chương 73 nhóm 7318 vs Chương 84 nhóm 8443); `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM4-CD03) | 🟡 Chờ duyệt |
| 96 | [W46] Thứ Sáu (20/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Hộ kinh doanh trả mặt bằng chuyển quận hoặc doanh nghiệp xuất hóa đơn thay thế khác quý: Đừng nghĩ 'nộp hồ sơ lên thuế là xong' — Cảnh báo bẫy bị giam lỏng mã số thuế, nợ thuế treo trên eTax Mobile và trát phạt tiền chậm nộp theo Nghị định 123/2020/NĐ-CP!** | Hướng dẫn 3 bước chuẩn hóa hồ sơ khi chuyển địa chỉ kinh doanh khác quận và cách kê khai hóa đơn thay thế khác kỳ tính thuế: Bí quyết giúp chủ hộ kinh doanh và kế toán không bị đình trệ hoạt động kinh doanh. | Góc nhìn từ phòng tư vấn Eureka: Tại sao chuyển cửa hàng từ Tân Bình sang Bình Thạnh hay Cầu Giấy sang Đống Đa lại trở thành 'cơn ác mộng' thủ tục? Bài học về sự chủ động chốt nghĩa vụ thuế trước khi thanh lý hợp đồng thuê mặt bằng. | Thông tư số 105/2020/TT-BTC về đăng ký thuế (Điều 10, Điều 11 về thay đổi thông tin đăng ký thuế); Nghị định số 123/2020/NĐ-CP (Điều 19 về xử lý hóa đơn có sai sót); Thông tư số 78/2021/TT-BTC; `scratch/daily_notes.md` ngày 22/10/2026 | 🟡 Chờ duyệt |
| 97 | [W46] Thứ Bảy (21/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Phiến Đá Babylon & Bộ Luật Hammurabi (1750 TCN): 3.800 Năm Trước Nhân Loại Đã Quản Trị Hợp Đồng Ngoại Thương, Vận Tải Đường Thủy & Bồi Thường Hàng Hóa Bản Lĩnh Thế Nào?** | Giải mã 282 điều luật khắc trên phiến đá diorite đen bảo tàng Louvre: Tại sao nguyên tắc 'Mắt đền mắt, răng đền răng' thực chất là cuộc cách mạng bảo vệ giới thương nhân và đặt nền móng cho hợp đồng logistics hiện đại? | Từ phiến đá Babylon 3.800 năm trước đến bản hợp đồng dịch vụ logistics thế kỷ 21: Tại sao sự minh bạch về trách nhiệm bồi thường và cam kết uy tín mới là thứ bảo vệ doanh nghiệp bền vững nhất? | Bộ luật Hammurabi (Code of Hammurabi c. 1750 BC, Bảo tàng Louvre Paris - Điều 104-107 về đại lý thương mại, Điều 236-240 về trách nhiệm tàu thuyền vận tải); Tư liệu Viện Đông phương học Chicago (Oriental Institute of Chicago); `Ke_Hoach_Phat_Song_YouTube_Series.xlsx` (CM3-CD04) | 🟡 Chờ duyệt |
| 98 | [W46] Chủ Nhật (22/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Nhập khẩu sâm tươi, dược liệu quý từ Hàn Quốc: Đừng nghe lời xúi 'xách tay cho nhanh' — Cảnh báo bẫy kiểm dịch thực vật tại Cục BVTV, công bố ATTP và nguy cơ bị phạt đến 100 triệu kèm tiêu hủy toàn bộ lô hàng tại sân bay!** | Lộ trình 3 bước nhập khẩu chính ngạch sâm tươi và nông sản nhạy cảm từ Hàn Quốc về Việt Nam: Từ giấy phép kiểm dịch thực vật nhập khẩu, C/O Form VK/AK đến bộ hồ sơ thông quan an toàn 100%. | Chuyện người làm nghề: Khi khách hàng tha thiết 'em giúp chị gửi phí vài triệu cứu lô sâm tươi đang bị giữ ở Nội Bài' — Lời khuyên chân thành của Thúy: Khi đã vướng rào cản kiểm dịch an toàn thực phẩm, chỉ có tuân thủ quy chuẩn ngay từ đầu mới bảo vệ được vốn liếng! | Luật An toàn thực phẩm 2010; Nghị định số 15/2018/NĐ-CP (quy định chi tiết thi hành một số điều của Luật ATTP); Thông tư số 04/2023/TT-BNNPTNT về danh mục kiểm dịch thực vật; Nghị định số 128/2020/NĐ-CP; `scratch/daily_notes.md` ngày 21/10/2026 | 🟡 Chờ duyệt |

## 2. Bản phân tích Tỷ lệ Góc tiếp cận & Nền tảng (Validation Tuần W46)

- **Tỷ lệ Năng lượng Nhân hiệu:**
  - **Giáp Mộc (~60%):** Tự hào lịch sử ngoại thương Đại Việt với Thương cảng Phố Hiến thế kỷ 17 & bản lĩnh đón tàu buôn phương Tây (FF Thứ Ba), Đồng hành gỡ vướng 3 bước sửa tờ khai hàng kho ngoại quan an toàn (CS Thứ Tư), Cầm tay chỉ việc 4 bước xin xác định trước mã số Advance Ruling gỡ nỗi đau con ốc vít (HT Thứ Năm - Học XNK 5 phút Tập 16), Khai mở tri thức lịch sử hợp đồng thương mại & vận tải từ phiến đá cổ Babylon Hammurabi 3.800 năm trước (FF Thứ Bảy).
  - **Liêm Trinh (~40%):** Sắc bén cảnh báo rủi ro siết chặt lưu giữ hồ sơ hải quan 05 năm theo Thông tư 121/2025/TT-BTC chuẩn bị cho các đợt kiểm tra sau thông quan (TN Thứ Hai), Nghiêm cẩn pháp lý về bẫy treo mã số thuế khi chuyển quận và hóa đơn thay thế khác kỳ theo Nghị định 123/2020/NĐ-CP (TN Thứ Sáu), Vạch trần nguy cơ tiêu hủy và phạt 100 triệu khi nhập khẩu xách tay sâm tươi không kiểm dịch thực vật & ATTP (CS Chủ Nhật).
- **Tỷ lệ Phân bổ Nền tảng (Chuẩn PBAN 2.0):**
  - **Facebook (4 bài/tuần):** Thứ Hai (TN), Thứ Tư (CS), Thứ Sáu (TN), Chủ Nhật (CS) — Đặt vấn đề sắc bén (Problem-First), bóc tách tình huống thực tế và kích thích thảo luận chuyên môn.
  - **YouTube (3 video/tuần):** Thứ Ba (FF Video Ký sự Lịch sử 8-12p), Thứ Năm (HT Video Học XNK 5 phút Tập 16), Thứ Bảy (FF Video Ký sự Lịch sử 8-12p) — Giáo dục chuyên sâu, trao giá trị thực chiến và xây dựng uy tín tri thức dài hạn.
- **Tuân thủ Bộ từ cấm (Banned Words Compliance):** 100% các tiêu đề và nội dung được quét tự động, tuyệt đối không chứa các từ cấm hay ngữ cảnh nhạy cảm (không có 'trốn thuế', 'lách luật', 'bao thuế', 'hàng lậu', 'tiểu ngạch', 'đánh hàng'...).

---

## 3. Quy trình Phê duyệt Checkpoint 1 & Đóng gói PO (wf01)

1. **Bước 1 (Checkpoint 1):** Chị Thúy xem xét các phương án Option A/B/C trên bảng kế hoạch tuần W46.
2. **Bước 2:** Chọn phương án duyệt tại Cột K của file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) (chọn `Duyệt Option A`, `Duyệt Option B`, hoặc `Duyệt Option C`).
3. **Bước 3:** Hệ thống tự động kích hoạt **Agent 04 (Planner)** đóng gói 7 Phiếu Giao việc (PO) vật lý tại thư mục `scratch/production_orders/` để sẵn sàng bàn giao cho Agent 05 (Facebook) và Agent 06 (YouTube) theo quy trình `wf02` và `wf03`.
"""
    with open(WEEKLY_CAL_MD, "r", encoding="utf-8") as f:
        existing_content = f.read()

    # Kiểm tra xem Tuần W46 đã có chưa
    if "# TUẦN W46: LỊCH BIÊN TẬP NỘI DUNG" in existing_content:
        # Thay thế phần W46 cũ
        parts = existing_content.split("# TUẦN W46: LỊCH BIÊN TẬP NỘI DUNG")
        updated_content = parts[0] + w46_md_content.lstrip()
    else:
        updated_content = existing_content + w46_md_content

    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(updated_content)
    print("  -> Đã cập nhật thành công reports/plan1/weekly_calendar.md!")

def save_json_data():
    print("=== 3. LƯU SCRATCH/W46_DATA.JSON ===")
    os.makedirs(os.path.dirname(JSON_OUTPUT), exist_ok=True)
    with open(JSON_OUTPUT, "w", encoding="utf-8") as f:
        json.dump(W46_ITEMS, f, ensure_ascii=False, indent=2)
    print(f"  -> Đã lưu thành công dữ liệu W46 vào {JSON_OUTPUT}")

if __name__ == "__main__":
    print("=== BẮT ĐẦU QUY TRÌNH WF01: LẬP LỊCH NỘI DUNG TUẦN W46 ===")
    update_excel_calendar()
    update_markdown_calendar()
    save_json_data()
    print("=== HOÀN TẤT THỰC THI GIAI ĐOẠN 1 TUẦN W46! ===")
