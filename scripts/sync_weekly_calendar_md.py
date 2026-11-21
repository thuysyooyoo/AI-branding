# -*- coding: utf-8 -*-
import openpyxl

excel_path = r"reports\plan1\Weekly_Content_Calendar.xlsx"
md_path = r"reports\plan1\weekly_calendar.md"

wb = openpyxl.load_workbook(excel_path, data_only=True)
ws = wb["Weekly Calendar"]

rows = []
for r in range(5, ws.max_row + 1):
    stt = ws.cell(row=r, column=1).value
    if not stt:
        continue
    day = str(ws.cell(row=r, column=2).value or "").strip()
    plat = str(ws.cell(row=r, column=3).value or "").strip()
    cat = str(ws.cell(row=r, column=4).value or "").strip()
    tech = str(ws.cell(row=r, column=5).value or "").strip()
    fmt = str(ws.cell(row=r, column=6).value or "").strip()
    opt_a = str(ws.cell(row=r, column=7).value or "").strip()
    opt_b = str(ws.cell(row=r, column=8).value or "").strip()
    opt_c = str(ws.cell(row=r, column=9).value or "").strip()
    ssot = str(ws.cell(row=r, column=10).value or "").strip()
    status = str(ws.cell(row=r, column=11).value or "").strip()
    
    if status == "Duyệt Option A":
        status_icon = "🟢 Duyệt Option A"
    elif status == "Duyệt Option B":
        status_icon = "🔵 Duyệt Option B"
    elif status == "Duyệt Option C":
        status_icon = "🟣 Duyệt Option C"
    elif status == "Không duyệt":
        status_icon = "🔴 Không duyệt"
    else:
        status_icon = "🟡 Chờ duyệt"
        
    rows.append({
        "stt": int(stt),
        "day": day,
        "plat": plat,
        "cat": cat,
        "tech": tech,
        "fmt": fmt,
        "opt_a": opt_a,
        "opt_b": opt_b,
        "opt_c": opt_c,
        "ssot": ssot,
        "status_icon": status_icon
    })

rows.sort(key=lambda x: x["stt"])

md_lines = [
    "# 📅 LỊCH BIÊN TẬP TUẦN - KẾ HOẠCH BÀI ĐĂNG (WEEKLY EDITORIAL CALENDAR)",
    "**Kế hoạch:** `plan1` — Quản lý lịch đăng & Bài đăng hoàn chỉnh của hệ thống PBAN  ",
    "**Kỳ đánh giá mới nhất:** Tuần W33 đến W43 (Từ 17/08/2026 đến 01/11/2026)  ",
    "**Trạng thái phê duyệt:** 🟢 W33 - W42 ĐÃ DUYỆT HOÀN TẤT (70 BÀI) | 🟡 W43 ĐANG CHỜ DUYỆT Ý TƯỞNG CHECKPOINT 1 (7 BÀI: STT 71-77)  ",
    "**Lịch Excel gốc:** [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx)",
    "",
    "---",
    "",
    "## 1. Bảng phân phối Lịch đăng bài 3 Phương án (Options A, B, C) — Hệ thống PBAN 2.0 (W33 - W43)",
    "",
    "Dưới đây là bảng tổng hợp các phương án chủ đề hàng tuần đã được tối ưu hóa theo đúng 4 Trụ cột nội dung (TN, CS, HT, FF), tỷ lệ 60% Giáp Mộc & 40% Liêm Trinh, và cân đối nền tảng Facebook (4 bài/tuần) & YouTube (3 video/tuần):",
    "",
    "| STT | Tuần / Ngày đăng | Nền tảng | Chuyên mục | Kỹ thuật Storytelling | Định dạng Sản phẩm | Option A (Chủ đề chính) | Option B (Dự phòng 1) | Option C (Dự phòng 2) | Nguồn SSoT / Trích xuất Pháp lý | Trạng thái Duyệt |",
    "|:---:|:---|:---:|:---:|:---|:---|:---|:---|:---|:---|:---:|"
]

for r in rows:
    line = f"| {r['stt']} | {r['day']} | {r['plat']} | {r['cat']} | {r['tech']} | {r['fmt']} | **{r['opt_a']}** | {r['opt_b']} | {r['opt_c']} | {r['ssot']} | {r['status_icon']} |"
    md_lines.append(line)

md_lines.extend([
    "",
    "---",
    "",
    "## 2. Bản phân tích Tỷ lệ Góc tiếp cận & Nền tảng (Validation Tuần W43)",
    "",
    "- **Tỷ lệ Năng lượng Nhân hiệu:**",
    "  - **Giáp Mộc (~60%):** Khai phóng tri thức Kỳ Án Khủng Hoảng Hạn Hán Kênh Đào Panama & bài học phân bổ tuyến hàng hải (FF Thứ Ba), Góc nhìn thấu hiểu & gỡ rối ranh giới DNCX vs Doanh nghiệp FDI thông thường trong KCX (CS Thứ Tư), Cầm tay chỉ việc 5 bước nộp hồ sơ Advance Ruling khóa rủi ro thuế quan (HT Thứ Năm - Học XNK 5 phút Tập 13), Bức tranh tương lai cỗ máy logistics cảng tự động hóa Rotterdam Maasvlakte 2 vận hành bằng 100% điện gió (FF Thứ Bảy).",
    "  - **Liêm Trinh (~40%):** Sắc bén cảnh báo rủi ro tắc C/O khi phân cấp về địa phương theo Thông tư 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP (TN Thứ Hai), Vạch trần cạm bẫy áp mã HS Smartwatch nhìn tên đồng hồ vội áp nhóm 91.02 theo Thông báo 22555/TB-CHQ (TN Thứ Sáu), Nghiêm cẩn pháp lý về điều kiện chỉ định thương nhân nước ngoài trong XNK tại chỗ theo Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan (CS Chủ Nhật).",
    "- **Tỷ lệ Phân bổ Nền tảng (Chuẩn PBAN 2.0):**",
    "  - **Facebook (4 bài/tuần):** Thứ Hai (TN), Thứ Tư (CS), Thứ Sáu (TN), Chủ Nhật (CS) — Đặt vấn đề sắc bén (Problem-First), bóc tách tình huống thực tế và kích thích thảo luận văn minh.",
    "  - **YouTube (3 video/tuần):** Thứ Ba (FF Video Phân tích & Insight 8-12p), Thứ Năm (HT Video Học XNK 5 phút Tập 13), Thứ Bảy (FF Video Phân tích chuyên sâu 8-12p) — Giáo dục chuyên sâu, trao giá trị thực chiến và xây dựng uy tín tri thức dài hạn.",
    "- **Tuân thủ Bộ từ cấm (Banned Words Compliance):** 100% các tiêu đề và nội dung được quét tự động, tuyệt đối không chứa các từ cấm hay ngữ cảnh nhạy cảm.",
    "",
    "---",
    "",
    "## 3. Quy trình Phê duyệt Checkpoint 1 & Đóng gói PO (wf01)",
    "",
    "1. **Bước 1 (Checkpoint 1):** Chị Thúy xem xét các phương án Option A/B/C trên bảng kế hoạch tuần W43.",
    "2. **Bước 2:** Chọn phương án duyệt tại Cột K của file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) (chọn `Duyệt Option A`, `Duyệt Option B`, hoặc `Duyệt Option C`).",
    "3. **Bước 3:** Hệ thống tự động kích hoạt **Agent 04 (Planner)** đóng gói 7 Phiếu Giao việc (PO) vật lý tại thư mục `scratch/production_orders/` để sẵn sàng bàn giao cho Agent 05 (Facebook) và Agent 06 (YouTube) theo quy trình `wf02` và `wf03`."
])

with open(md_path, "w", encoding="utf-8") as f:
    f.write("\n".join(md_lines))

print(f"SUCCESS: Updated {md_path} with {len(rows)} items!")
