# -*- coding: utf-8 -*-
"""
Cập nhật content_status_report.md sau khi sản xuất xong loạt bài từ Daily Notes 25/09/2026.
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

FILE_PATH = r"d:\AI AGENT THUY\AI branding\reports\plan1\content_status_report.md"

with open(FILE_PATH, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update Header & Summary
old_summary = """- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 53 bài viết hoàn chỉnh (36 bài theo lịch chính + 2 bài W38-HOT + 5 bài W39-HOT + 6 bài HOT 23/09 + 4 bài HOT mới từ Daily Notes 24/09/2026). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check.
  - **Kênh YouTube (Sheet 3):** 33 kịch bản video hoàn chỉnh (27 kịch bản theo lịch chính W33-W41 + 3 kịch bản W39-HOT + 1 kịch bản HOT06 + 2 kịch bản HOT12, HOT13). 100% có kịch bản chi tiết, Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 86/86 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`."""

new_summary = """- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 59 bài viết hoàn chỉnh (36 bài theo lịch chính + 2 bài W38-HOT + 5 bài W39-HOT + 6 bài HOT 23/09 + 4 bài HOT 24/09 + 6 bài HOT mới từ Daily Notes 25/09/2026). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check.
  - **Kênh YouTube (Sheet 3):** 36 kịch bản video hoàn chỉnh (27 kịch bản theo lịch chính W33-W41 + 3 kịch bản W39-HOT + 1 kịch bản HOT06 + 2 kịch bản HOT12, HOT13 + 3 kịch bản HOT16, HOT17, HOT20). 100% có kịch bản chi tiết, Storyboard 4 cột, Audio/Visual/SFX và Cliffhanger Outro.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 95/95 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`."""

content = content.replace("24/09/2026 17:15", "25/09/2026 17:30")
content = content.replace(old_summary, new_summary)
content = content.replace("AGENT 05) — 53 BÀI", "AGENT 05) — 59 BÀI")
content = content.replace("AGENT 06) — 33 KỊCH BẢN", "AGENT 06) — 36 KỊCH BẢN")
content = content.replace("53 bài viết hoàn chỉnh kèm Media Brief", "59 bài viết hoàn chỉnh kèm Media Brief")
content = content.replace("33 kịch bản hoàn chỉnh kèm Storyboard", "36 kịch bản hoàn chỉnh kèm Storyboard")

# 2. Add Facebook rows 54 -> 59
fb_addition = """| 54 | `PO-2026-W39-HOT16` | [W39] Thứ Sáu (25/09/2026) | Cont bị gãy seal hãng tàu khi chuyển cảng Cát Lái sang Hiệp Phước, cảng đã bấm seal mới: Khi truyền tờ khai lấy hàng phải điền số seal nào để không bị giữ cont tại cổng? | [DRAFT-FB-HOT-16-GAY-SEAL-CHUYEN-CANG.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-16-GAY-SEAL-CHUYEN-CANG.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 55 | `PO-2026-W39-HOT17` | [W39] Thứ Bảy (26/09/2026) | Hàng được mang về bảo quản theo Thông tư 121 nhưng phòng lab chưa có kết quả, Chi cục TĐC lại từ chối ký gia hạn: Làm sao để doanh nghiệp không bị phạt và tước quyền mang hàng trong 6 tháng? | [DRAFT-FB-HOT-17-TT121-GIA-HAN-KTCL.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-17-TT121-GIA-HAN-KTCL.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 56 | `PO-2026-W39-HOT18` | [W39] Chủ Nhật (27/09/2026) | Nhập khẩu hành lá sấy nhiệt cắt vụn từ Trung Quốc (HS 0712.90.90): Áp mã miễn thuế VAT 'VK010' theo Nghị định 181/2025/NĐ-CP được không hay dính bẫy truy thu thuế 5%? | [DRAFT-FB-HOT-18-VK010-VAT-HANH-LA-SAY.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-18-VK010-VAT-HANH-LA-SAY.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 57 | `PO-2026-W39-HOT19` | [W40] Thứ Hai (28/09/2026) | Kinh doanh cho thuê xe ô tô tự lái: Có phải xin Giấy phép kinh doanh vận tải không? Cẩm nang đăng ký chuẩn mã ngành 7710 và tránh rủi ro trách nhiệm liên đới khi khách gây tai nạn! | [DRAFT-FB-HOT-19-THUE-XE-TU-LAI-7710.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-19-THUE-XE-TU-LAI-7710.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 58 | `PO-2026-W39-HOT20` | [W40] Thứ Ba (29/09/2026) | 'Lúc mua nó là công ty bình thường, sao giờ bảo nó là ma rồi bắt tôi đóng phạt?': Cách tự vệ trước bão thanh tra hóa đơn đầu vào và bảo vệ chi phí hợp lệ! | [DRAFT-FB-HOT-20-HOA-DON-CONG-TY-MA.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-20-HOA-DON-CONG-TY-MA.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 59 | `PO-2026-W39-HOT21` | [W40] Thứ Tư (30/09/2026) | Vừa là chủ hộ kinh doanh cửa hàng vừa làm Affiliate Shopee/TikTok: Nên khai thuế riêng hay gộp doanh thu để không bị truy thu phạt và tối ưu thuế suất từ 10% xuống 7%? | [DRAFT-FB-HOT-21-HKD-AFFILIATE-TMDT.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-21-HKD-AFFILIATE-TMDT.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
"""

target_row_fb = "| 53 | `PO-2026-W39-HOT15` | [W39] Chủ Nhật (27/09/2026) | Quy trình xin Giấy xác nhận hoàn thành nghĩa vụ thuế (Mẫu 24) nộp Sở Kế hoạch & Đầu tư: Bẫy 'thụ động chờ đợi' làm treo hồ sơ giải thể / thay đổi hàng tháng trời! | [DRAFT-FB-HOT-15-XAC-NHAN-THUE-MAU-24.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-HOT-15-XAC-NHAN-THUE-MAU-24.md) | `🟢 PASSED` | **Sẵn sàng đăng** |\n"
content = content.replace(target_row_fb, target_row_fb + fb_addition)

# 3. Add YouTube rows 34 -> 36
yt_addition = """| 34 | `PO-2026-W39-HOT16` | [W39] Series Học XNK 5 Phút: Container Bị Gãy Seal Khi Chuyển Cảng Cát Lái - Hiệp Phước (Khai Seal Nào?) | `HT — Học Thực chiến` | [DRAFT-YT-HOT-16-GAY-SEAL-CHUYEN-CANG.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-16-GAY-SEAL-CHUYEN-CANG.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 35 | `PO-2026-W39-HOT17` | [W39] Series Học XNK 5 Phút: Bẫy 30 Ngày Thông Tư 121 & Cẩm Nang Gia Hạn Kiểm Tra Chất Lượng | `HT — Học Thực chiến` | [DRAFT-YT-HOT-17-TT121-GIA-HAN-KTCL.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-17-TT121-GIA-HAN-KTCL.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 36 | `PO-2026-W39-HOT20` | [W40] Series Quản Trị & Pháp Lý: Bẫy Hóa Đơn Công Ty Ma & 4 Bằng Chứng Thép Tự Vệ Trước Thanh Tra Thuế | `CS — Case Study` | [DRAFT-YT-HOT-20-HOA-DON-CONG-TY-MA.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-20-HOA-DON-CONG-TY-MA.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
"""

target_row_yt = "| 33 | `PO-2026-W39-HOT13` | [W39] Series Học XNK 5 Phút: Cảnh Báo Nghẽn Cảng Bao Bì Thực Phẩm & Công Văn 22082/CHQ-GSQL | `HT — Học Thực chiến` | [DRAFT-YT-HOT-13-CV22082-BAO-BI-ATTP.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-HOT-13-CV22082-BAO-BI-ATTP.md) | `🟢 PASSED` | **Sẵn sàng dựng** |\n"
content = content.replace(target_row_yt, target_row_yt + yt_addition)

with open(FILE_PATH, "w", encoding="utf-8") as f:
    f.write(content)

print("Đã cập nhật thành công content_status_report.md")
