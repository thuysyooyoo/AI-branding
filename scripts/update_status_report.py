# -*- coding: utf-8 -*-
import os

STATUS_REP_MD = r"d:\AI AGENT THUY\AI branding\reports\plan1\content_status_report.md"

with open(STATUS_REP_MD, "r", encoding="utf-8") as f:
    text = f.read()

# 1. Update Section 1
old_sec1 = """## 📌 1. TỔNG QUAN TIẾN ĐỘ SẢN XUẤT TOÀN GIAI ĐOẠN (W33 - W42 & DAILY NOTES HOT TOPICS)

- **Tổng số chủ đề trên Lịch Biên tập (Sheet 1):** 70 chủ đề (W33 đến W42, từ 17/08/2026 đến 25/10/2026)
- **Tình trạng Phê duyệt Ý tưởng (Checkpoint 1):** 70/70 chủ đề (`🟢 Duyệt Option A` 100%)
- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 78 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 38 bài HOT Daily Notes) (36 bài theo lịch chính + 2 bài W38-HOT + 5 bài W39-HOT + 6 bài HOT 23/09 + 4 bài HOT 24/09 + 6 bài HOT 25/09 + 4 bài HOT 26/09: HOT22-HOT25 + 6 bài HOT 28-30/09: HOT26-HOT31 + 5 bài HOT mới 01-07/10/2026: HOT32-HOT36). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check.
  - **Kênh YouTube (Sheet 3):** 50 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 20 kịch bản HOT Daily Notes) (27 kịch bản theo lịch chính W33-W41 + 3 kịch bản W39-HOT + 1 kịch bản HOT06 + 2 kịch bản HOT12, HOT13 + 3 kịch bản HOT16, HOT17, HOT20 + 3 kịch bản HOT 26/09: HOT22-HOT24 + 4 kịch bản HOT 28-30/09: HOT26, HOT27, HOT29, HOT30 + 4 kịch bản HOT mới 01-07/10/2026: HOT32-HOT35). 100% có kịch bản chi tiết, Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 128/128 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`.
- **Rà soát Ghi chú Cần sửa đổi:** Không có bài nào bị đánh dấu `🔴 REVISION REQUIRED` hoặc có ghi chú yêu cầu chỉnh sửa từ Chị Thúy.
- **Trạng thái Xuất bản:** 100% bài Facebook ở trạng thái **Sẵn sàng đăng**; 100% kịch bản YouTube ở trạng thái **Sẵn sàng dựng / Đã duyệt quay**."""

new_sec1 = """## 📌 1. TỔNG QUAN TIẾN ĐỘ SẢN XUẤT TOÀN GIAI ĐOẠN (W33 - W43 & DAILY NOTES HOT TOPICS)

- **Tổng số chủ đề trên Lịch Biên tập (Sheet 1):** 77 chủ đề (W33 đến W43, từ 17/08/2026 đến 01/11/2026)
- **Tình trạng Phê duyệt Ý tưởng (Checkpoint 1):** 77/77 chủ đề (`🟢 Duyệt Option A` 100%)
- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 82 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 4 bài W43 + 38 bài HOT Daily Notes). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check SSoT.
  - **Kênh YouTube (Sheet 3):** 53 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 3 kịch bản W43 + 20 kịch bản HOT Daily Notes). 100% có kịch bản lời thoại đầy đủ, Storyboard 4 cột (Audio/Visual/SFX) và Cliffhanger Outro kết nối tập sau.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 135/135 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`.
- **Rà soát Ghi chú Cần sửa đổi:** 100% các bài viết từ W33 đến W43 đều đạt chuẩn SSoT Zero-Trust, không có bài nào bị đánh dấu `🔴 REVISION REQUIRED` hoặc có ghi chú yêu cầu chỉnh sửa từ Chị Thúy.
- **Trạng thái Xuất bản:** 100% bài Facebook ở trạng thái **Sẵn sàng đăng**; 100% kịch bản YouTube ở trạng thái **Sẵn sàng dựng / Đã duyệt quay**."""

# 2. Update Header 2 and add rows
old_fb_head = "## 📋 2. BẢNG TỔNG HỢP NỘI DUNG KÊNH FACEBOOK (SHEET 2 — AGENT 05) — 78 BÀI"
new_fb_head = "## 📋 2. BẢNG TỔNG HỢP NỘI DUNG KÊNH FACEBOOK (SHEET 2 — AGENT 05) — 82 BÀI"

fb_rows_add = """| 78 | `PO-2026-W41-HOT36` | [W41] Thứ Tư (07/10/2026) | Hóa đơn thuế nhà thầu (FCT) phát sinh cùng ngày tờ khai hải quan nhưng kế toán bỏ sót chưa khai nộp: Hướng dẫn nộp bổ sung tránh phạt quá hạn 10 ngày! | [DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 79 | `PO-2026-W43-71` | [W43] Thứ Hai (26/10/2026) | Thông tư 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP: Phân cấp thẩm quyền cấp C/O về địa phương — Cảnh báo bẫy ngộ nhận 'thủ tục ở tỉnh sẽ dễ hơn' khiến hàng loạt lô hàng xuất khẩu kẹt cứng vì tiêu chí xác minh xuất xứ bị siết chặt! | [DRAFT-FB-PO-2026-W43-71.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W43-71.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 80 | `PO-2026-W43-73` | [W43] Thứ Tư (28/10/2026) | Doanh nghiệp FDI nằm trong Khu chế xuất (KCX) mua hàng từ nội địa: Cạm bẫy nhập nhằng giữa 'Doanh nghiệp chế xuất (DNCX) hưởng thuế quan ưu đãi' và 'Doanh nghiệp FDI thông thường' — Sai một bước là bị Thuế truy thu và tắc toàn bộ hóa đơn VAT! | [DRAFT-FB-PO-2026-W43-73.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W43-73.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 81 | `PO-2026-W43-75` | [W43] Thứ Sáu (30/10/2026) | Thông báo 22555/TB-CHQ vừa ban hành: Cảnh báo bẫy áp mã HS thiết bị đeo thông minh (Smartwatch) — Nhìn chữ 'Đồng hồ' vội áp nhóm 91.02 thay vì nhóm 85.17 là dính ngay bẫy thuế và xử phạt vi phạm hành chính! | [DRAFT-FB-PO-2026-W43-75.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W43-75.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 82 | `PO-2026-W43-77` | [W43] Chủ Nhật (01/11/2026) | Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan: Cạm bẫy giao dịch Xuất nhập khẩu tại chỗ khi thiếu 'Văn bản chỉ định giao nhận của thương nhân nước ngoài' — Hàng chục tờ khai bị đình chỉ thông quan và nguy cơ truy thu thuế GTGT 10%! | [DRAFT-FB-PO-2026-W43-77.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W43-77.md) | `🟢 PASSED` | **Sẵn sàng đăng** |"""

old_yt_head = "## 📋 3. BẢNG TỔNG HỢP NỘI DUNG KÊNH YOUTUBE (SHEET 3 — AGENT 06) — 50 KỊCH BẢN"
new_yt_head = "## 📋 3. BẢNG TỔNG HỢP NỘI DUNG KÊNH YOUTUBE (SHEET 3 — AGENT 06) — 53 KỊCH BẢN"

yt_rows_add = """| 50 | `PO-2026-W41-HOT35` | [W41] Series Quản Trị & Pháp Lý: Doanh Nghiệp FDI Trong Khu Chế Xuất Mua Hàng Nội Địa: Khi Nào Được Miễn Mở Tờ Khai? 3 Cụm Mặt Hàng 'Bẫy' Kế Toán Cần Tránh | `CS — Case Study` | [DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 51 | `PO-2026-W43-72` | [W43] Series Kỳ Án Hàng Hải — Tập 14: Khủng Hoảng Hạn Hán Kênh Đào Panama & Cuộc Đấu Giá 4 Triệu USD | `FF — Funfact & Insight` | [DRAFT-YT-SERIES-PO-2026-W43-72.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W43-72.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 52 | `PO-2026-W43-74` | [W43] Series Học XNK 5 Phút — Tập 13: 5 Bước Đăng Ký Xác Định Trước Mã Số Hàng Hóa (Advance Ruling) | `HT — Học Thực chiến` | [DRAFT-YT-SERIES-PO-2026-W43-74.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W43-74.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 53 | `PO-2026-W43-76` | [W43] Series Đột Phá Chuỗi Cung Ứng — Tập 15: Siêu Cảng Maasvlakte 2 Rotterdam Không Một Bóng Người | `FF — Funfact & Insight` | [DRAFT-YT-SERIES-PO-2026-W43-76.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W43-76.md) | `🟢 PASSED` | **Sẵn sàng dựng** |"""

old_sec4 = """1. **Kiểm tra File Master Calendar:**
   - Mở file [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx).
   - **Sheet 1 (`Weekly Calendar`):** 70 dòng đã duyệt `🟢 Duyệt Option A` 100%.
   - **Sheet 2 (`Facebook`):** 78 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 38 bài HOT Daily Notes) kèm Media Brief và Link SSoT.
   - **Sheet 3 (`YouTube`):** 50 kịch bản hoàn chỉnh kèm Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro."""

new_sec4 = """1. **Kiểm tra File Master Calendar:**
   - Mở file [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx).
   - **Sheet 1 (`Weekly Calendar`):** 77 dòng đã duyệt `🟢 Duyệt Option A` 100%.
   - **Sheet 2 (`Facebook`):** 82 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 4 bài W43 + 38 bài HOT Daily Notes) kèm Media Brief và Link SSoT.
   - **Sheet 3 (`YouTube`):** 53 kịch bản hoàn chỉnh kèm Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro."""

text = text.replace(old_sec1, new_sec1)
text = text.replace(old_fb_head, new_fb_head)
text = text.replace("""| 78 | `PO-2026-W41-HOT36` | [W41] Thứ Tư (07/10/2026) | Hóa đơn thuế nhà thầu (FCT) phát sinh cùng ngày tờ khai hải quan nhưng kế toán bỏ sót chưa khai nộp: Hướng dẫn nộp bổ sung tránh phạt quá hạn 10 ngày! | [DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-36-THUE-NHA-THAU-SOT-HOA-DON.md) | `🟢 PASSED` | **Sẵn sàng đăng** |""", fb_rows_add)
text = text.replace(old_yt_head, new_yt_head)
text = text.replace("""| 50 | `PO-2026-W41-HOT35` | [W41] Series Quản Trị & Pháp Lý: Doanh Nghiệp FDI Trong Khu Chế Xuất Mua Hàng Nội Địa: Khi Nào Được Miễn Mở Tờ Khai? 3 Cụm Mặt Hàng 'Bẫy' Kế Toán Cần Tránh | `CS — Case Study` | [DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-35-DN-FDI-TRONG-KCX-MUA-HANG.md) | `🟢 PASSED` | **Sẵn sàng dựng** |""", yt_rows_add)
text = text.replace(old_sec4, new_sec4)

with open(STATUS_REP_MD, "w", encoding="utf-8") as f:
    f.write(text)

print("SUCCESS: Cập nhật thành công content_status_report.md!")
