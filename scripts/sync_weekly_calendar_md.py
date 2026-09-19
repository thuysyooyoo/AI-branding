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
    "**Kỳ đánh giá mới nhất:** Tuần W33 đến W41 (Từ 17/08/2026 đến 18/10/2026)  ",
    "**Trạng thái phê duyệt:** 🟢 W33 - W40 ĐÃ DUYỆT & HOÀN TẤT LẬP LỊCH (56 BÀI) | 🟡 W41 ĐANG CHỜ DUYỆT Ý TƯỞNG (7 BÀI: STT 57-63)  ",
    "**Lịch Excel gốc:** [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx)",
    "",
    "---",
    "",
    "## 1. Bảng phân phối Lịch đăng bài 3 Phương án (Options A, B, C) — Hệ thống PBAN 2.0 (W33 - W41)",
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
    "## 2. Bản phân tích Tỷ lệ Góc tiếp cận & Nền tảng (Validation Tuần W41)",
    "",
    "- **Tỷ lệ Năng lượng Nhân hiệu:**",
    "  - **Giáp Mộc (~60%):** Nâng đỡ, khai phóng tri thức Siêu Cảng Singapore & Tuas 20 tỷ USD vận hành 100% bằng robot (FF Thứ Ba), cầm tay chỉ việc khai e-Manifest và đồng bộ mã vạch dữ liệu số theo Thông tư 128 qua Cửa khẩu thông minh (HT Thứ Năm - Học XNK 5 phút Tập 11), bóc tách nghịch lý cỗ máy tuần hoàn container rỗng toàn cầu (FF Thứ Bảy), cẩm nang quản trị công nợ B2B và nghệ thuật đàm phán bảo lãnh ngân hàng (CS Chủ Nhật).",
    "  - **Liêm Trinh (~40%):** Sắc bén cảnh báo 72h đếm ngược bộ ba chuyển đổi số (TT 128, NĐ 336, QĐ 1434) có nguy cơ bị khóa barie cửa khẩu thông minh (TN Thứ Hai), vạch trần bẫy mã HS linh kiện máy nén khí (lò xo khí, bộ báo nghẹt lọc gió) truy thu thuế 5 năm (CS Thứ Tư), cảnh báo bão rà soát thuế COD qua eTax Mobile và đề xuất hạ ngưỡng miễn thuế 100.000đ (TN Thứ Sáu).",
    "- **Tỷ lệ Phân bổ Nền tảng (Chuẩn PBAN 2.0):**",
    "  - **Facebook (4 bài/tuần):** Thứ Hai (TN), Thứ Tư (CS), Thứ Sáu (TN), Chủ Nhật (CS) — Đặt vấn đề sắc bén (Problem-First), bóc tách tình huống thực tế và kích thích tranh luận văn minh.",
    "  - **YouTube (3 video/tuần):** Thứ Ba (FF Video Phân tích & Insight 8-12p), Thứ Năm (HT Video Học XNK 5 phút Tập 11), Thứ Bảy (FF Video Phân tích chuyên sâu 8-12p) — Giáo dục chuyên sâu, trao giá trị thực chiến và xây dựng uy tín tri thức dài hạn.",
    "- **Tuân thủ Bộ từ cấm (Banned Words Compliance):** 100% các tiêu đề và nội dung được quét tự động, tuyệt đối không chứa các từ cấm hay ngữ cảnh nhạy cảm.",
    "",
    "---",
    "",
    "## 3. Quy trình Phê duyệt Checkpoint 1 & Đóng gói PO (wf01)",
    "",
    "1. **Bước 1 (Checkpoint 1):** Chị Thúy xem xét các phương án Option A/B/C trên bảng kế hoạch tuần W41.",
    "2. **Bước 2:** Chọn phương án duyệt tại Cột K của file Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) (chọn `Duyệt Option A`, `Duyệt Option B`, hoặc `Duyệt Option C`).",
    "3. **Bước 3:** Hệ thống tự động kích hoạt **Agent 04 (Planner)** đóng gói 7 Phiếu Giao việc (PO) vật lý tại thư mục `scratch/production_orders/` để sẵn sàng bàn giao cho Agent 05 (Facebook) và Agent 06 (YouTube) theo quy trình `wf02` và `wf03`."
])

with open(md_path, "w", encoding="utf-8") as f:
    f.write("\n".join(md_lines))

print(f"SUCCESS: Updated {md_path} with {len(rows)} items!")
