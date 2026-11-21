# -*- coding: utf-8 -*-
"""
Script: update_content_status_report_3009.py
Cập nhật content_status_report.md với các số liệu mới nhất:
- Kênh Facebook: 73 bài viết (thêm STT 68-73 HOT26-HOT31)
- Kênh YouTube: 46 kịch bản (thêm STT 43-46 HOT26, HOT27, HOT29, HOT30)
- Tổng sản phẩm: 119/119 đạt 🟢 PASSED
"""

import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

REPORT_PATH = r"d:\AI AGENT THUY\AI branding\reports\plan1\content_status_report.md"

with open(REPORT_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update Header overview
old_overview = """- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 67 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 27 bài HOT Daily Notes) (36 bài theo lịch chính + 2 bài W38-HOT + 5 bài W39-HOT + 6 bài HOT 23/09 + 4 bài HOT 24/09 + 6 bài HOT 25/09 + 4 bài HOT mới từ Daily Notes 26/09/2026: HOT22-HOT25). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check.
  - **Kênh YouTube (Sheet 3):** 42 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 12 kịch bản HOT Daily Notes) (27 kịch bản theo lịch chính W33-W41 + 3 kịch bản W39-HOT + 1 kịch bản HOT06 + 2 kịch bản HOT12, HOT13 + 3 kịch bản HOT16, HOT17, HOT20 + 3 kịch bản HOT mới từ Daily Notes 26/09/2026: HOT22, HOT23, HOT24). 100% có kịch bản chi tiết, Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 109/109 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`."""

new_overview = """- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 73 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 33 bài HOT Daily Notes) (36 bài theo lịch chính + 2 bài W38-HOT + 5 bài W39-HOT + 6 bài HOT 23/09 + 4 bài HOT 24/09 + 6 bài HOT 25/09 + 4 bài HOT 26/09: HOT22-HOT25 + 6 bài HOT mới từ Daily Notes 28-30/09/2026: HOT26-HOT31). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check.
  - **Kênh YouTube (Sheet 3):** 46 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 16 kịch bản HOT Daily Notes) (27 kịch bản theo lịch chính W33-W41 + 3 kịch bản W39-HOT + 1 kịch bản HOT06 + 2 kịch bản HOT12, HOT13 + 3 kịch bản HOT16, HOT17, HOT20 + 3 kịch bản HOT 26/09: HOT22-HOT24 + 4 kịch bản HOT mới từ Daily Notes 28-30/09/2026: HOT26, HOT27, HOT29, HOT30). 100% có kịch bản chi tiết, Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 119/119 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`."""

content = content.replace(old_overview, new_overview)

# Update Facebook section header
content = content.replace("## 📋 2. BẢNG TỔNG HỢP NỘI DUNG KÊNH FACEBOOK (SHEET 2 — AGENT 05) — 67 BÀI",
                          "## 📋 2. BẢNG TỔNG HỢP NỘI DUNG KÊNH FACEBOOK (SHEET 2 — AGENT 05) — 73 BÀI")

# Update YouTube section header
content = content.replace("## 📋 3. BẢNG TỔNG HỢP NỘI DUNG KÊNH YOUTUBE (SHEET 3 — AGENT 06) — 42 KỊCH BẢN",
                          "## 📋 3. BẢNG TỔNG HỢP NỘI DUNG KÊNH YOUTUBE (SHEET 3 — AGENT 06) — 46 KỊCH BẢN")

# 2. Append rows to Facebook Table
fb_tail_target = '| 67 | `PO-2026-W39-HOT25` | [W40] Thứ Bảy (03/10/2026) | Xuất khẩu thực phẩm sang Trung Đông & Malaysia: Đừng để mất trắng đơn hàng vì \'chứng chỉ Halal dỏm\' — Nhận diện 5 tổ chức chứng nhận Halal được công nhận toàn cầu và quy chuẩn logistics Halal chuỗi lạnh! | [DRAFT-FB-HOT-25-XUAT-KHAU-HALAL-LOGISTICS.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-25-XUAT-KHAU-HALAL-LOGISTICS.md) | `🟢 PASSED` | **Sẵn sàng đăng** |\n\n---'

fb_rows_to_append = """| 67 | `PO-2026-W39-HOT25` | [W40] Thứ Bảy (03/10/2026) | Xuất khẩu thực phẩm sang Trung Đông & Malaysia: Đừng để mất trắng đơn hàng vì 'chứng chỉ Halal dỏm' — Nhận diện 5 tổ chức chứng nhận Halal được công nhận toàn cầu và quy chuẩn logistics Halal chuỗi lạnh! | [DRAFT-FB-HOT-25-XUAT-KHAU-HALAL-LOGISTICS.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-25-XUAT-KHAU-HALAL-LOGISTICS.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 68 | `PO-2026-W40-HOT26` | [W40] Thứ Hai (28/09/2026) | Hóa đơn đầu vào ghi tên hàng chung chung, đầu ra ghi chi tiết đặc tính kỹ thuật: Cạm bẫy truy thu thuế & Bộ 'Khiên 3 chứng từ' tự vệ cho kế toán doanh nghiệp | [DRAFT-FB-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 69 | `PO-2026-W40-HOT27` | [W40] Thứ Ba (29/09/2026) | Cần hủy & sửa đổi Báo cáo Quyết toán mẫu số 15 tại Chi cục Hải quan Quản lý Hàng gia công TP.HCM (Cơ sở 15): Cảnh báo bẫy quá hạn 60 ngày Điều 60 Thông tư 39 khiến doanh nghiệp bị kiểm tra sau thông quan tại trụ sở! | [DRAFT-FB-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 70 | `PO-2026-W40-HOT28` | [W40] Thứ Ba (29/09/2026) | Ở nhà chung cư không đăng ký Hộ kinh doanh được: Muốn xuất hóa đơn cho khách hàng doanh nghiệp thì làm thế nào? Cẩm nang xin cấp hóa đơn lẻ từng lần phát sinh từ Chi cục Thuế! | [DRAFT-FB-HOT-28-XUAT-HOA-DON-LE-CHUNG-CU.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-28-XUAT-HOA-DON-LE-CHUNG-CU.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 71 | `PO-2026-W40-HOT29` | [W40] Thứ Tư (30/09/2026) | C/O Form AI 3 bên thể hiện 2 trị giá FOB và 2 số Invoice (nhà sản xuất & bên thứ ba): Hải quan có chấp nhận không? Cẩm nang đối soát ô số 9, 10, 13 theo Hiệp định AIFTA tránh bị bác ưu đãi thuế! | [DRAFT-FB-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 72 | `PO-2026-W40-HOT30` | [W40] Thứ Tư (30/09/2026) | DNCX gửi hạt nhựa đi gia công biến tính nội địa không làm đổi mã HS: Ghi nhãn và khai xuất xứ 'Made in China' hay 'Việt Nam'? Cạm bẫy xử phạt xuất xứ theo Nghị định 111/2021/NĐ-CP & Thông tư 38/2015! | [DRAFT-FB-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 73 | `PO-2026-W40-HOT31` | [W40] Thứ Tư (30/09/2026) | Kê khai thuế hộ kinh doanh & cá nhân theo quý: Lập tờ khai cộng dồn doanh thu 3 tháng hay lập từng tháng? 3 sai lầm khiến chủ shop bị phạt chậm nộp và ấn định thuế theo Thông tư 40/2021/TT-BTC! | [DRAFT-FB-HOT-31-KE-KHAI-THUE-THEO-QUY-HO-KINH-DOAN.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-31-KE-KHAI-THUE-THEO-QUY-HO-KINH-DOAN.md) | `🟢 PASSED` | **Sẵn sàng đăng** |

---"""

content = content.replace(fb_tail_target, fb_rows_to_append)

# 3. Append rows to YouTube Table
yt_tail_target = '| 42 | `PO-2026-W39-HOT24` | [W40] Series Đột Phá Chuỗi Cung Ứng: Bẫy Phạt 10% CBAM Châu Âu & Lộ Trình Kiểm Toán Carbon | `HT — Học Thực chiến` | [DRAFT-YT-HOT-24-CBAM-MARKUP-10-PERCENT.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-24-CBAM-MARKUP-10-PERCENT.md) | `🟢 PASSED` | **Sẵn sàng dựng** |\n\n---'

yt_rows_to_append = """| 42 | `PO-2026-W39-HOT24` | [W40] Series Đột Phá Chuỗi Cung Ứng: Bẫy Phạt 10% CBAM Châu Âu & Lộ Trình Kiểm Toán Carbon | `HT — Học Thực chiến` | [DRAFT-YT-HOT-24-CBAM-MARKUP-10-PERCENT.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-24-CBAM-MARKUP-10-PERCENT.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 43 | `PO-2026-W40-HOT26` | [W40] Series Quản Trị & Pháp Lý: Hóa Đơn Đầu Vào Ghi Chung Chung, Đầu Ra Ghi Chi Tiết: Bẫy Truy Thu Thuế Tiền Tỷ & Bộ Khiên Chắn 3 Chứng Từ Tự Vệ Cho Kế Toán | `CS — Case Study` | [DRAFT-YT-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-26-HOA-DON-DAU-VAO-CHUNG-CHUNG.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 44 | `PO-2026-W40-HOT27` | [W40] Series Học XNK 5 Phút: Cần Hủy & Sửa Đổi Báo Cáo Quyết Toán Mẫu 15: Cạm Bẫy Quá Hạn 60 Ngày Điều 60 Thông Tư 39 & Quy Trình Gỡ Khẩn Cấp | `HT — Học Thực chiến` | [DRAFT-YT-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-27-HUY-SUA-BAO-CAO-QUYET-TOAN-MAU-15.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 45 | `PO-2026-W40-HOT29` | [W40] Series Học XNK 5 Phút: C/O Form AI 3 Bên Thể Hiện 2 Trị Giá FOB & 2 Số Invoice: Hải Quan Có Chấp Nhận? Checklist 4 Điểm Sống Còn Tránh Bị Bác Thuế 0% | `HT — Học Thực chiến` | [DRAFT-YT-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-29-CO-FORM-AI-BA-BEN-2-INVOICE.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 46 | `PO-2026-W40-HOT30` | [W40] Series Học XNK 5 Phút: DNCX Gửi Hàng Gia Công Biến Tính Không Đổi Mã HS: Ghi Nhãn & Khai Xuất Xứ Thế Nào Tránh Phạt Tiền Tỷ Nghị Định 111? | `CS — Case Study` | [DRAFT-YT-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-30-DNCX-GIA-CONG-BIEN-TINH-KHAI-XUAT-XU.md) | `🟢 PASSED` | **Sẵn sàng dựng** |

---"""

content = content.replace(yt_tail_target, yt_rows_to_append)

# 4. Update section 4 summary counts
content = content.replace("67 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 27 bài HOT Daily Notes)",
                          "73 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 33 bài HOT Daily Notes)")
content = content.replace("42 kịch bản hoàn chỉnh",
                          "46 kịch bản hoàn chỉnh")

with open(REPORT_PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("✅ Đã cập nhật thành công content_status_report.md!")
