# -*- coding: utf-8 -*-
"""
Script điều phối đồng bộ toàn diện Giai đoạn 2 cho Tuần W43 (STT 71 - 77):
1. Xuất file vật lý:
   - 7 PO tại scratch/production_orders/PO-2026-W43-XX.md
   - 4 DRAFT FB tại scratch/production_orders/DRAFT-FB-PO-2026-W43-XX.md & reports/plan1/facebook/FB-2026-W43-XX.md
   - 3 DRAFT YT tại scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W43-XX.md & reports/plan1/video_scripts/YS-2026-W43-XX.md
   - 7 FACT-CHECK tại scratch/production_orders/FACT-CHECK-PO-2026-W43-XX.md
2. Cập nhật Weekly_Content_Calendar.xlsx (Sheet 1, Sheet 2, Sheet 3) với atomic_save (Zero Backup)
3. Cập nhật reports/plan1/weekly_calendar.md
4. Cập nhật reports/plan1/content_status_report.md
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Import data modules
from w43_item_71 import DATA_71, PO_71_MD, DRAFT_FB_71, FACT_CHECK_71_MD
from w43_item_72 import DATA_72, PO_72_MD, DRAFT_YT_72, FACT_CHECK_72_MD
from w43_item_73 import DATA_73, PO_73_MD, DRAFT_FB_73, FACT_CHECK_73_MD
from w43_item_74 import DATA_74, PO_74_MD, DRAFT_YT_74, FACT_CHECK_74_MD
from w43_item_75 import DATA_75, PO_75_MD, DRAFT_FB_75, FACT_CHECK_75_MD
from w43_item_76 import DATA_76, PO_76_MD, DRAFT_YT_76, FACT_CHECK_76_MD
from w43_item_77 import DATA_77, PO_77_MD, DRAFT_FB_77, FACT_CHECK_77_MD

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
WEEKLY_CAL_MD = os.path.join(BASE_DIR, "reports", "plan1", "weekly_calendar.md")
STATUS_REP_MD = os.path.join(BASE_DIR, "reports", "plan1", "content_status_report.md")

os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_FB_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)

print("=== 1. XUẤT CÁC TỆP TIN VẬT LÝ ===")

# Danh sách bài viết Facebook
fb_items = [
    (DATA_71, PO_71_MD, DRAFT_FB_71, FACT_CHECK_71_MD),
    (DATA_73, PO_73_MD, DRAFT_FB_73, FACT_CHECK_73_MD),
    (DATA_75, PO_75_MD, DRAFT_FB_75, FACT_CHECK_75_MD),
    (DATA_77, PO_77_MD, DRAFT_FB_77, FACT_CHECK_77_MD)
]

for data, po_md, draft, fc_md in fb_items:
    po_id = data["po_id"]
    # Save PO
    po_file = os.path.join(SCRATCH_PO_DIR, f"{po_id}.md")
    with open(po_file, "w", encoding="utf-8") as f:
        f.write(po_md)
    # Save Draft in scratch
    draft_scratch = os.path.join(SCRATCH_PO_DIR, f"DRAFT-FB-{po_id}.md")
    with open(draft_scratch, "w", encoding="utf-8") as f:
        f.write(draft)
    # Save Draft in reports
    draft_rep = os.path.join(REP_FB_DIR, f"FB-{po_id}.md")
    with open(draft_rep, "w", encoding="utf-8") as f:
        f.write(draft)
    # Save Fact-Check
    fc_file = os.path.join(SCRATCH_PO_DIR, f"FACT-CHECK-{po_id}.md")
    with open(fc_file, "w", encoding="utf-8") as f:
        f.write(fc_md)
    print(f"  -> Đã tạo đầy đủ 4 tệp cho Facebook: {po_id}")

# Danh sách kịch bản YouTube
yt_items = [
    (DATA_72, PO_72_MD, DRAFT_YT_72, FACT_CHECK_72_MD),
    (DATA_74, PO_74_MD, DRAFT_YT_74, FACT_CHECK_74_MD),
    (DATA_76, PO_76_MD, DRAFT_YT_76, FACT_CHECK_76_MD)
]

for data, po_md, draft, fc_md in yt_items:
    po_id = data["po_id"]
    # Save PO
    po_file = os.path.join(SCRATCH_PO_DIR, f"{po_id}.md")
    with open(po_file, "w", encoding="utf-8") as f:
        f.write(po_md)
    # Save Draft in scratch
    draft_scratch = os.path.join(SCRATCH_PO_DIR, f"DRAFT-YT-SERIES-{po_id}.md")
    with open(draft_scratch, "w", encoding="utf-8") as f:
        f.write(draft)
    # Save Draft in reports
    draft_rep = os.path.join(REP_YT_DIR, f"YS-{po_id}.md")
    with open(draft_rep, "w", encoding="utf-8") as f:
        f.write(draft)
    # Save Fact-Check
    fc_file = os.path.join(SCRATCH_PO_DIR, f"FACT-CHECK-{po_id}.md")
    with open(fc_file, "w", encoding="utf-8") as f:
        f.write(fc_md)
    print(f"  -> Đã tạo đầy đủ 4 tệp cho YouTube: {po_id}")

print("\n=== 2. CẬP NHẬT FILE EXCEL WEEKLY_CONTENT_CALENDAR.XLSX ===")

wb = openpyxl.load_workbook(EXCEL_PATH)

font_family = "Segoe UI"
bold_font = Font(name=font_family, size=10, bold=True)
regular_font = Font(name=font_family, size=10)
italic_font = Font(name=font_family, size=9, italic=True)

green_pass_fill = PatternFill(start_color="D5F5E3", end_color="D5F5E3", fill_type="solid") # Xanh lá nhạt
ready_pub_fill = PatternFill(start_color="E8F8F5", end_color="E8F8F5", fill_type="solid")
zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9")
)

center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

# 2.1 CẬP NHẬT SHEET 1 (Weekly Calendar)
ws1 = wb["Weekly Calendar"]
# Cập nhật cột K (Approval) cho dòng 75 đến 81 (STT 71 đến 77)
for r in range(5, ws1.max_row + 1):
    stt_val = ws1.cell(r, 1).value
    if stt_val in [71, 72, 73, 74, 75, 76, 77]:
        ws1.cell(r, 11, value="🟢 Duyệt Option A")
        ws1.cell(r, 11).fill = green_pass_fill
        ws1.cell(r, 11).font = bold_font
        ws1.cell(r, 11).alignment = center_align
print("  -> Sheet 1: Đã cập nhật 7 dòng W43 thành '🟢 Duyệt Option A'")

# 2.2 CẬP NHẬT SHEET 2 (Facebook)
ws2 = wb["Facebook"]
current_fb_max = ws2.max_row
existing_fb_pos = [ws2.cell(r, 2).value for r in range(5, current_fb_max + 1)]

# Tìm STT tiếp theo cho Facebook
last_fb_stt = 0
for r in range(5, current_fb_max + 1):
    val = ws2.cell(r, 1).value
    if isinstance(val, int) and val > last_fb_stt:
        last_fb_stt = val

for data, po_md, draft, fc_md in fb_items:
    po_id = data["po_id"]
    if po_id not in existing_fb_pos:
        last_fb_stt += 1
        new_row = ws2.max_row + 1
        fill = zebra_fill if new_row % 2 == 0 else white_fill
        link_md = f"file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-{po_id}.md"
        
        row_vals = [
            last_fb_stt,
            po_id,
            data["date"],
            data["title"],
            data["seo"],
            link_md,
            draft,
            data["visual_brief"],
            "🟢 PASSED",
            data["ssot_report"],
            data["pub_status"]
        ]
        
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws2.cell(row=new_row, column=c_idx, value=val)
            cell.border = thin_border
            cell.fill = fill
            cell.alignment = center_align if c_idx in [1, 2, 3, 9, 11] else left_align
            
            if c_idx == 9: # Fact check
                cell.fill = green_pass_fill
                cell.font = bold_font
            elif c_idx == 11: # Pub status
                cell.fill = ready_pub_fill
                cell.font = bold_font
            elif c_idx == 6: # Link md
                cell.font = italic_font
            elif c_idx in [1, 2, 4]:
                cell.font = bold_font
            else:
                cell.font = regular_font
                
        print(f"  -> Sheet 2 (Facebook): Đã thêm dòng mới STT {last_fb_stt} - {po_id}")
    else:
        print(f"  -> Sheet 2 (Facebook): {po_id} đã tồn tại, bỏ qua.")

# 2.3 CẬP NHẬT SHEET 3 (YouTube)
ws3 = wb["YouTube"]
current_yt_max = ws3.max_row
existing_yt_pos = [ws3.cell(r, 2).value for r in range(5, current_yt_max + 1)]

# Tìm STT tiếp theo cho YouTube
last_yt_stt = 0
for r in range(5, current_yt_max + 1):
    val = ws3.cell(r, 1).value
    if isinstance(val, int) and val > last_yt_stt:
        last_yt_stt = val

for data, po_md, draft, fc_md in yt_items:
    po_id = data["po_id"]
    if po_id not in existing_yt_pos:
        last_yt_stt += 1
        new_row = ws3.max_row + 1
        fill = zebra_fill if new_row % 2 == 0 else white_fill
        link_md = f"file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-{po_id}.md"
        
        row_vals = [
            last_yt_stt,
            po_id,
            data["series"],
            data["cat"],
            data["seo"],
            link_md,
            draft,
            data["visual_brief"],
            "🟢 PASSED",
            data["ssot_report"],
            data["pub_status"]
        ]
        
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws3.cell(row=new_row, column=c_idx, value=val)
            cell.border = thin_border
            cell.fill = fill
            cell.alignment = center_align if c_idx in [1, 2, 4, 9, 11] else left_align
            
            if c_idx == 9: # Fact check
                cell.fill = green_pass_fill
                cell.font = bold_font
            elif c_idx == 11: # Pub status
                cell.fill = ready_pub_fill
                cell.font = bold_font
            elif c_idx == 6: # Link md
                cell.font = italic_font
            elif c_idx in [1, 2, 3]:
                cell.font = bold_font
            else:
                cell.font = regular_font
                
        print(f"  -> Sheet 3 (YouTube): Đã thêm dòng mới STT {last_yt_stt} - {po_id}")
    else:
        print(f"  -> Sheet 3 (YouTube): {po_id} đã tồn tại, bỏ qua.")

# ATOMIC SAVE TUÂN THỦ RULE 10 (KHÔNG TẠO BACKUP FILE)
tmp_path = EXCEL_PATH + ".tmp.xlsx"
wb.save(tmp_path)
os.replace(tmp_path, EXCEL_PATH)
print("SUCCESS: Cập nhật thành công file Excel Weekly_Content_Calendar.xlsx!")

# 3. CẬP NHẬT WEEKLY_CALENDAR.MD
if os.path.exists(WEEKLY_CAL_MD):
    with open(WEEKLY_CAL_MD, "r", encoding="utf-8") as f:
        content = f.read()
    # Thay thế trạng thái W43 từ 🟡 Chờ duyệt sang 🟢 Duyệt Option A
    content_updated = content.replace("| 71 | [W43] Thứ Hai (26/10/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Thông tư 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP: Phân cấp thẩm quyền cấp C/O về địa phương — Cảnh báo bẫy ngộ nhận 'thủ tục ở tỉnh sẽ dễ hơn' khiến hàng loạt lô hàng xuất khẩu kẹt cứng vì tiêu chí xác minh xuất xứ bị siết chặt!** | Checklist 4 bước chuẩn bị hồ sơ xác minh cơ sở sản xuất và giải trình định mức nguyên liệu (BOM) theo quy chế phân cấp mới để được cấp C/O Form D, AK, VK trong 24 giờ. | Khi thẩm quyền cấp C/O chuyển giao về địa phương: Bài học quản trị rủi ro 'bất đồng quan điểm thẩm định' giữa hệ thống eCoSys và cán bộ kiểm tra thực địa. | Thông tư số 26/2026/TT-BCT ngày 15/07/2026 của Bộ Công Thương, Nghị định số 146/2025/NĐ-CP, Nghị định số 31/2018/NĐ-CP về xuất xứ hàng hóa, Quy chế vận hành Hệ thống eCoSys & scratch/daily_notes.md ngày 09/10/2026 | 🟡 Chờ duyệt |",
                              "| 71 | [W43] Thứ Hai (26/10/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Thông tư 26/2026/TT-BCT & Nghị định 146/2025/NĐ-CP: Phân cấp thẩm quyền cấp C/O về địa phương — Cảnh báo bẫy ngộ nhận 'thủ tục ở tỉnh sẽ dễ hơn' khiến hàng loạt lô hàng xuất khẩu kẹt cứng vì tiêu chí xác minh xuất xứ bị siết chặt!** | Checklist 4 bước chuẩn bị hồ sơ xác minh cơ sở sản xuất và giải trình định mức nguyên liệu (BOM) theo quy chế phân cấp mới để được cấp C/O Form D, AK, VK trong 24 giờ. | Khi thẩm quyền cấp C/O chuyển giao về địa phương: Bài học quản trị rủi ro 'bất đồng quan điểm thẩm định' giữa hệ thống eCoSys và cán bộ kiểm tra thực địa. | Thông tư số 26/2026/TT-BCT ngày 15/07/2026 của Bộ Công Thương, Nghị định số 146/2025/NĐ-CP, Nghị định số 31/2018/NĐ-CP về xuất xứ hàng hóa, Quy chế vận hành Hệ thống eCoSys & scratch/daily_notes.md ngày 09/10/2026 | 🟢 Duyệt Option A |")
    content_updated = content_updated.replace("| 72 | [W43] Thứ Ba (27/10/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Khủng Hoảng Hạn Hán Kênh Đào Panama: Cuộc Đấu Giá 4 Triệu USD Để Giành Một Slot Vượt Kênh, Nghịch Lý 'Thừa Tàu Thiếu Nước' & Bài Học Quản Trị Rủi Ro Tuyến Tuyến Hàng Hải Cho Doanh Nghiệp Xuất Khẩu Sang Bờ Đông Nước Mỹ** | Bí mật kỹ thuật đằng sau hệ thống âu tàu nâng nước hồ Gatun: Tại sao một con kênh nhân tạo 110 năm tuổi lại có thể làm đảo lộn chuỗi cung ứng nông sản và dệt may toàn cầu? | Từ khủng hoảng Panama 2024-2026 đến bài toán đa dạng hóa tuyến đường: Tại sao doanh nghiệp logistics Việt Nam phải học cách phân bổ rủi ro qua kênh Suez, Mũi Hảo Vọng và cầu cạn xuyên Mỹ (IPI)? | Báo cáo Thường niên Cơ quan Quản lý Kênh đào Panama (ACP 2024-2026), Dữ liệu Phân tích Thị trường Vận tải Drewry WCI & Freightos, Báo cáo Chuỗi Cung ứng Toàn cầu Lloyd's List & Quy tắc York-Antwerp | 🟡 Chờ duyệt |",
                                              "| 72 | [W43] Thứ Ba (27/10/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Khủng Hoảng Hạn Hán Kênh Đào Panama: Cuộc Đấu Giá 4 Triệu USD Để Giành Một Slot Vượt Kênh, Nghịch Lý 'Thừa Tàu Thiếu Nước' & Bài Học Quản Trị Rủi Ro Tuyến Tuyến Hàng Hải Cho Doanh Nghiệp Xuất Khẩu Sang Bờ Đông Nước Mỹ** | Bí mật kỹ thuật đằng sau hệ thống âu tàu nâng nước hồ Gatun: Tại sao một con kênh nhân tạo 110 năm tuổi lại có thể làm đảo lộn chuỗi cung ứng nông sản và dệt may toàn cầu? | Từ khủng hoảng Panama 2024-2026 đến bài toán đa dạng hóa tuyến đường: Tại sao doanh nghiệp logistics Việt Nam phải học cách phân bổ rủi ro qua kênh Suez, Mũi Hảo Vọng và cầu cạn xuyên Mỹ (IPI)? | Báo cáo Thường niên Cơ quan Quản lý Kênh đào Panama (ACP 2024-2026), Dữ liệu Phân tích Thị trường Vận tải Drewry WCI & Freightos, Báo cáo Chuỗi Cung ứng Toàn cầu Lloyd's List & Quy tắc York-Antwerp | 🟢 Duyệt Option A |")
    content_updated = content_updated.replace("| 73 | [W43] Thứ Tư (28/10/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Doanh nghiệp FDI nằm trong Khu chế xuất (KCX) mua hàng từ nội địa: Cạm bẫy nhập nhằng giữa 'Doanh nghiệp chế xuất (DNCX) hưởng thuế quan ưu đãi' và 'Doanh nghiệp FDI thông thường' — Sai một bước là bị Thuế truy thu và tắc toàn bộ hóa đơn VAT!** | Sơ đồ 3 bước xác định chuẩn xác: Khi nào doanh nghiệp FDI trong KCX bắt buộc phải mở tờ khai hải quan XNK tại chỗ, và khi nào áp dụng hóa đơn GTGT 10% như thương mại nội địa thông thường? | Tâm sự người làm nghề tư vấn: Phía sau câu chuyện kế toán cãi nhau với cán bộ thuế là sự thiếu rạch ròi giữa 'Địa giới hành chính KCX' và 'Tư cách pháp nhân DNCX'. | Điều 74 & Điều 86 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Điều 26 Nghị định số 35/2022/NĐ-CP về quản lý KCN & KKT, Công văn số 22411/CHQ-GSQL của Cục Hải quan & scratch/daily_notes.md ngày 07/10/2026 | 🟡 Chờ duyệt |",
                                              "| 73 | [W43] Thứ Tư (28/10/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Doanh nghiệp FDI nằm trong Khu chế xuất (KCX) mua hàng từ nội địa: Cạm bẫy nhập nhằng giữa 'Doanh nghiệp chế xuất (DNCX) hưởng thuế quan ưu đãi' và 'Doanh nghiệp FDI thông thường' — Sai một bước là bị Thuế truy thu và tắc toàn bộ hóa đơn VAT!** | Sơ đồ 3 bước xác định chuẩn xác: Khi nào doanh nghiệp FDI trong KCX bắt buộc phải mở tờ khai hải quan XNK tại chỗ, và khi nào áp dụng hóa đơn GTGT 10% như thương mại nội địa thông thường? | Tâm sự người làm nghề tư vấn: Phía sau câu chuyện kế toán cãi nhau với cán bộ thuế là sự thiếu rạch ròi giữa 'Địa giới hành chính KCX' và 'Tư cách pháp nhân DNCX'. | Điều 74 & Điều 86 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Điều 26 Nghị định số 35/2022/NĐ-CP về quản lý KCN & KKT, Công văn số 22411/CHQ-GSQL của Cục Hải quan & scratch/daily_notes.md ngày 07/10/2026 | 🟢 Duyệt Option A |")
    content_updated = content_updated.replace("| 74 | [W43] Thứ Năm (29/10/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 13 | **Học XNK 5 Phút Tập 13: 5 Bước Đăng Ký Xác Định Trước Mã Số Hàng Hóa (Advance Ruling) Để 'Khóa' Rủi Ro Thuế Quan & Không Bị Hải Quan Bác Mã HS Khi Thông Quan** | Checklist 4 tài liệu kỹ thuật bắt buộc phải chuẩn bị (Catalogue gốc, tài liệu kỹ thuật dịch công chứng, kết quả kiểm nghiệm xưởng) để hồ sơ Advance Ruling được Cục Hải quan chấp thuận trong 30 ngày. | Mẹo thực chiến cho doanh nghiệp nhập khẩu máy móc, thiết bị công nghệ mới: Cách gửi mẫu phân tích và xây dựng văn bản giải trình tính chất kỹ thuật chuẩn xác theo Điều 24 Nghị định 08/2015/NĐ-CP. | Điều 28 Luật Hải quan số 54/2014/QH13, Điều 24 Nghị định số 08/2015/NĐ-CP (sửa đổi tại Nghị định 59/2018/NĐ-CP), Điều 7 Thông tư số 38/2015/TT-BTC, Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan | 🟡 Chờ duyệt |",
                                              "| 74 | [W43] Thứ Năm (29/10/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 13 | **Học XNK 5 Phút Tập 13: 5 Bước Đăng Ký Xác Định Trước Mã Số Hàng Hóa (Advance Ruling) Để 'Khóa' Rủi Ro Thuế Quan & Không Bị Hải Quan Bác Mã HS Khi Thông Quan** | Checklist 4 tài liệu kỹ thuật bắt buộc phải chuẩn bị (Catalogue gốc, tài liệu kỹ thuật dịch công chứng, kết quả kiểm nghiệm xưởng) để hồ sơ Advance Ruling được Cục Hải quan chấp thuận trong 30 ngày. | Mẹo thực chiến cho doanh nghiệp nhập khẩu máy móc, thiết bị công nghệ mới: Cách gửi mẫu phân tích và xây dựng văn bản giải trình tính chất kỹ thuật chuẩn xác theo Điều 24 Nghị định 08/2015/NĐ-CP. | Điều 28 Luật Hải quan số 54/2014/QH13, Điều 24 Nghị định số 08/2015/NĐ-CP (sửa đổi tại Nghị định 59/2018/NĐ-CP), Điều 7 Thông tư số 38/2015/TT-BTC, Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan | 🟢 Duyệt Option A |")
    content_updated = content_updated.replace("| 75 | [W43] Thứ Sáu (30/10/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Thông báo 22555/TB-CHQ vừa ban hành: Cảnh báo bẫy áp mã HS thiết bị đeo thông minh (Smartwatch) — Nhìn chữ 'Đồng hồ' vội áp nhóm 91.02 thay vì nhóm 85.17 là dính ngay bẫy thuế và xử phạt vi phạm hành chính!** | Bảng đối chiếu 4 tiêu chí kỹ thuật quyết định phân loại thiết bị đeo tay theo Chú giải 3 Phần XVI: Phân định rạch ròi giữa 'Đồng hồ có chức năng phụ trợ' và 'Thiết bị truyền/nhận dữ liệu đa năng'. | Từ case study Smartwatch Huawei Watch GT 6: Bài học đắt giá cho dân làm chứng từ công nghệ — Hiểu bản chất vi mạch và giao thức Bluetooth/Wi-Fi quan trọng hơn nhìn thiết kế bề ngoài! | Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan về kết quả xác định trước mã số, Danh mục hàng hóa XNK Việt Nam (Thông tư 31/2022/TT-BTC), 6 Quy tắc Tổng quát giải thích HS (GIR) & scratch/daily_notes.md ngày 07/10/2026 | 🟡 Chờ duyệt |",
                                              "| 75 | [W43] Thứ Sáu (30/10/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Thông báo 22555/TB-CHQ vừa ban hành: Cảnh báo bẫy áp mã HS thiết bị đeo thông minh (Smartwatch) — Nhìn chữ 'Đồng hồ' vội áp nhóm 91.02 thay vì nhóm 85.17 là dính ngay bẫy thuế và xử phạt vi phạm hành chính!** | Bảng đối chiếu 4 tiêu chí kỹ thuật quyết định phân loại thiết bị đeo tay theo Chú giải 3 Phần XVI: Phân định rạch ròi giữa 'Đồng hồ có chức năng phụ trợ' và 'Thiết bị truyền/nhận dữ liệu đa năng'. | Từ case study Smartwatch Huawei Watch GT 6: Bài học đắt giá cho dân làm chứng từ công nghệ — Hiểu bản chất vi mạch và giao thức Bluetooth/Wi-Fi quan trọng hơn nhìn thiết kế bề ngoài! | Thông báo số 22555/TB-CHQ ngày 05/10/2026 của Cục Hải quan về kết quả xác định trước mã số, Danh mục hàng hóa XNK Việt Nam (Thông tư 31/2022/TT-BTC), 6 Quy tắc Tổng quát giải thích HS (GIR) & scratch/daily_notes.md ngày 07/10/2026 | 🟢 Duyệt Option A |")
    content_updated = content_updated.replace("| 76 | [W43] Thứ Bảy (31/10/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Đế Chế Cảng Biển Tự Động Hóa Maasvlakte 2 (Rotterdam): Cỗ Máy Logistics 'Không Một Bóng Người' Bốc Xếp, Vận Hành Bằng 100% Điện Gió & Cuộc Cách Mạng Xanh Đang Thay Đổi Cuộc Đua Cảng Biển Thế Giới 2026** | Bí mật sau đội xe AGV tự hành và cần cẩu STS điều khiển từ xa cách 50km: Làm sao Rotterdam giải bài toán tăng gấp đôi năng suất bốc dỡ mà giảm 100% phát thải carbon tại bến cảng? | Từ kỳ quan tự động hóa Rotterdam đến chiến lược cảng biển Việt Nam: Bài học nào cho các siêu cảng nước sâu Cần Giờ, Cái Mép - Thị Vải và Lạch Huyện trên hành trình trở thành Hub trung chuyển quốc tế? | Báo cáo Thường niên Port of Rotterdam Authority 2025-2026, Dữ liệu Công nghệ Cảng biển Tự động hóa APM Terminals Rotterdam, Báo cáo Nghiên cứu Chuyển đổi Xanh Hàng hải của Tổ chức Hàng hải Quốc tế (IMO / World Bank) & Tài liệu Phân tích Cảng biển Eureka | 🟡 Chờ duyệt |",
                                              "| 76 | [W43] Thứ Bảy (31/10/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Đế Chế Cảng Biển Tự Động Hóa Maasvlakte 2 (Rotterdam): Cỗ Máy Logistics 'Không Một Bóng Người' Bốc Xếp, Vận Hành Bằng 100% Điện Gió & Cuộc Cách Mạng Xanh Đang Thay Đổi Cuộc Đua Cảng Biển Thế Giới 2026** | Bí mật sau đội xe AGV tự hành và cần cẩu STS điều khiển từ xa cách 50km: Làm sao Rotterdam giải bài toán tăng gấp đôi năng suất bốc dỡ mà giảm 100% phát thải carbon tại bến cảng? | Từ kỳ quan tự động hóa Rotterdam đến chiến lược cảng biển Việt Nam: Bài học nào cho các siêu cảng nước sâu Cần Giờ, Cái Mép - Thị Vải và Lạch Huyện trên hành trình trở thành Hub trung chuyển quốc tế? | Báo cáo Thường niên Port of Rotterdam Authority 2025-2026, Dữ liệu Công nghệ Cảng biển Tự động hóa APM Terminals Rotterdam, Báo cáo Nghiên cứu Chuyển đổi Xanh Hàng hải của Tổ chức Hàng hải Quốc tế (IMO / World Bank) & Tài liệu Phân tích Cảng biển Eureka | 🟢 Duyệt Option A |")
    content_updated = content_updated.replace("| 77 | [W43] Chủ Nhật (01/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan: Cạm bẫy giao dịch Xuất nhập khẩu tại chỗ khi thiếu 'Văn bản chỉ định giao nhận của thương nhân nước ngoài' — Hàng chục tờ khai bị đình chỉ thông quan và nguy cơ truy thu thuế GTGT 10%!** | Checklist 3 điều khoản sống còn bắt buộc phải rà soát trong Hợp đồng mua bán 3 bên và Thư chỉ định giao hàng (Delivery Instruction) để được Hải quan chấp nhận thông quan XNK tại chỗ mượt mà. | Bài học xương máu từ một đợt thanh tra sau thông quan: Tại sao thiếu một dòng chỉ định trong hợp đồng ngoại thương lại biến một giao dịch XNK hợp pháp thành hành vi mua bán nội địa bị phạt nặng? | Điều 47a Luật Hải quan (được bổ sung bởi Luật số 90/2025/QH15), Điều 35 Nghị định số 08/2015/NĐ-CP (sửa đổi, bổ sung tại Nghị định 167/2025/NĐ-CP), Thông tư 121/2025/TT-BTC sửa đổi Điều 86 Thông tư 38/2015/TT-BTC, Công văn số 22411/CHQ-GSQL ngày 30/09/2026 của Cục Hải quan & scratch/daily_notes.md ngày 01/10/2026 | 🟡 Chờ duyệt |",
                                              "| 77 | [W43] Chủ Nhật (01/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan: Cạm bẫy giao dịch Xuất nhập khẩu tại chỗ khi thiếu 'Văn bản chỉ định giao nhận của thương nhân nước ngoài' — Hàng chục tờ khai bị đình chỉ thông quan và nguy cơ truy thu thuế GTGT 10%!** | Checklist 3 điều khoản sống còn bắt buộc phải rà soát trong Hợp đồng mua bán 3 bên và Thư chỉ định giao hàng (Delivery Instruction) để được Hải quan chấp nhận thông quan XNK tại chỗ mượt mà. | Bài học xương máu từ một đợt thanh tra sau thông quan: Tại sao thiếu một dòng chỉ định trong hợp đồng ngoại thương lại biến một giao dịch XNK hợp pháp thành hành vi mua bán nội địa bị phạt nặng? | Điều 47a Luật Hải quan (được bổ sung bởi Luật số 90/2025/QH15), Điều 35 Nghị định số 08/2015/NĐ-CP (sửa đổi, bổ sung tại Nghị định 167/2025/NĐ-CP), Thông tư 121/2025/TT-BTC sửa đổi Điều 86 Thông tư 38/2015/TT-BTC, Công văn số 22411/CHQ-GSQL ngày 30/09/2026 của Cục Hải quan & scratch/daily_notes.md ngày 01/10/2026 | 🟢 Duyệt Option A |")
    
    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(content_updated)
    print("SUCCESS: Cập nhật thành công file reports/plan1/weekly_calendar.md!")

print("\n=== HOÀN TẤT ĐỒNG BỘ TOÀN DIỆN GIAI ĐOẠN 2 CHO TUẦN W43! ===")
