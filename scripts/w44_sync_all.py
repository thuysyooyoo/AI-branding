# -*- coding: utf-8 -*-
"""
Script điều phối đồng bộ toàn diện Giai đoạn 2 cho Tuần W44 (STT 78 - 84):
1. Xuất file vật lý:
   - 7 PO tại scratch/production_orders/PO-2026-W44-XX.md
   - 4 DRAFT FB tại scratch/production_orders/DRAFT-FB-PO-2026-W44-XX.md & reports/plan1/facebook/FB-PO-2026-W44-XX.md
   - 3 DRAFT YT tại scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-XX.md & reports/plan1/video_scripts/YS-PO-2026-W44-XX.md
   - 7 FACT-CHECK tại scratch/production_orders/FACT-CHECK-PO-2026-W44-XX.md
2. Cập nhật Weekly_Content_Calendar.xlsx (Sheet 1, Sheet 2, Sheet 3) với atomic_save (Zero Backup - RULE 10)
3. Cập nhật reports/plan1/weekly_calendar.md
4. Cập nhật reports/plan1/content_status_report.md
"""
import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Import data modules for W44
from w44_item_78 import DATA_78, PO_78_MD, DRAFT_FB_78, FACT_CHECK_78_MD
from w44_item_79 import DATA_79, PO_79_MD, DRAFT_YT_79, FACT_CHECK_79_MD
from w44_item_80 import DATA_80, PO_80_MD, DRAFT_FB_80, FACT_CHECK_80_MD
from w44_item_81 import DATA_81, PO_81_MD, DRAFT_YT_81, FACT_CHECK_81_MD
from w44_item_82 import DATA_82, PO_82_MD, DRAFT_FB_82, FACT_CHECK_82_MD
from w44_item_83 import DATA_83, PO_83_MD, DRAFT_YT_83, FACT_CHECK_83_MD
from w44_item_84 import DATA_84, PO_84_MD, DRAFT_FB_84, FACT_CHECK_84_MD

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

print("=== 1. XUẤT CÁC TỆP TIN VẬT LÝ TUẦN W44 ===")

# Danh sách bài viết Facebook
fb_items = [
    (DATA_78, PO_78_MD, DRAFT_FB_78, FACT_CHECK_78_MD),
    (DATA_80, PO_80_MD, DRAFT_FB_80, FACT_CHECK_80_MD),
    (DATA_82, PO_82_MD, DRAFT_FB_82, FACT_CHECK_82_MD),
    (DATA_84, PO_84_MD, DRAFT_FB_84, FACT_CHECK_84_MD)
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
    (DATA_79, PO_79_MD, DRAFT_YT_79, FACT_CHECK_79_MD),
    (DATA_81, PO_81_MD, DRAFT_YT_81, FACT_CHECK_81_MD),
    (DATA_83, PO_83_MD, DRAFT_YT_83, FACT_CHECK_83_MD)
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
for r in range(5, ws1.max_row + 1):
    stt_val = ws1.cell(r, 1).value
    if stt_val in [78, 79, 80, 81, 82, 83, 84]:
        ws1.cell(r, 11, value="🟢 Duyệt Option A")
        ws1.cell(r, 11).fill = green_pass_fill
        ws1.cell(r, 11).font = bold_font
        ws1.cell(r, 11).alignment = center_align
print("  -> Sheet 1: Đã cập nhật 7 dòng W44 (STT 78-84) thành '🟢 Duyệt Option A'")

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
    
    # Cập nhật Header
    content_updated = content.replace(
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W43 (Từ 17/08/2026 đến 01/11/2026)  \n**Trạng thái phê duyệt:** 🟢 W33 - W43 ĐÃ DUYỆT HOÀN TẤT 100% (77 BÀI) | 🟢 ĐÃ SẢN XUẤT HOÀN TẤT 100% TRÊN SHEET 2 (FACEBOOK) & SHEET 3 (YOUTUBE)",
        "**Kỳ đánh giá mới nhất:** Tuần W33 đến W44 (Từ 17/08/2026 đến 08/11/2026)  \n**Trạng thái phê duyệt:** 🟢 W33 - W44 ĐÃ DUYỆT HOÀN TẤT 100% (84 BÀI) | 🟢 ĐÃ SẢN XUẤT HOÀN TẤT 100% TRÊN SHEET 2 (FACEBOOK) & SHEET 3 (YOUTUBE)"
    )
    content_updated = content_updated.replace(
        "Bảng phân phối Lịch đăng bài 3 Phương án (Options A, B, C) — Hệ thống PBAN 2.0 (W33 - W43)",
        "Bảng phân phối Lịch đăng bài 3 Phương án (Options A, B, C) — Hệ thống PBAN 2.0 (W33 - W44)"
    )
    
    # Cập nhật trạng thái duyệt từ 🟡 Chờ duyệt sang 🟢 Duyệt Option A cho STT 78-84
    w44_replacements = [
        ("| 78 | [W44] Thứ Hai (02/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Nghị định 37/2026/NĐ-CP siết chặt nhãn gốc: Đừng tin câu 'Tao xuất sang Việt Nam nhiều lần rồi không sao' của Shipper ngoại — Lô máy móc tiền tỷ kẹt cứng tại Hải Phòng vì tem gốc không có 'Made in China' và nguy cơ bị phạt đến 60 triệu kèm chế tài buộc tái xuất!** | Checklist 3 thông tin bắt buộc 'không bao giờ được nợ' trên nhãn gốc hàng nhập khẩu theo Nghị định 37/2026/NĐ-CP: Hướng dẫn doanh nghiệp 'khóa' điều khoản tem nhãn vào hợp đồng ngoại thương trước khi thanh toán nốt tiền hàng. | Bài học xương máu từ chiếc máy cơ khí nhập khẩu: Khi Cnee lúng túng hỏi 'chi bao nhiêu để qua?' và lời cảnh tỉnh của người làm nghề — Trong thời điểm Hải quan siết chặt xuất xứ, chỉ có minh bạch pháp lý mới cứu được doanh nghiệp! | Nghị định số 37/2026/NĐ-CP về nhãn hàng hóa (thay thế Nghị định 43/2017/NĐ-CP và Nghị định 111/2021/NĐ-CP), Điều 22 Nghị định số 128/2020/NĐ-CP về xử phạt vi phạm hành chính trong lĩnh vực hải quan, scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |",
         "| 78 | [W44] Thứ Hai (02/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **Nghị định 37/2026/NĐ-CP siết chặt nhãn gốc: Đừng tin câu 'Tao xuất sang Việt Nam nhiều lần rồi không sao' của Shipper ngoại — Lô máy móc tiền tỷ kẹt cứng tại Hải Phòng vì tem gốc không có 'Made in China' và nguy cơ bị phạt đến 60 triệu kèm chế tài buộc tái xuất!** | Checklist 3 thông tin bắt buộc 'không bao giờ được nợ' trên nhãn gốc hàng nhập khẩu theo Nghị định 37/2026/NĐ-CP: Hướng dẫn doanh nghiệp 'khóa' điều khoản tem nhãn vào hợp đồng ngoại thương trước khi thanh toán nốt tiền hàng. | Bài học xương máu từ chiếc máy cơ khí nhập khẩu: Khi Cnee lúng túng hỏi 'chi bao nhiêu để qua?' và lời cảnh tỉnh của người làm nghề — Trong thời điểm Hải quan siết chặt xuất xứ, chỉ có minh bạch pháp lý mới cứu được doanh nghiệp! | Nghị định số 37/2026/NĐ-CP về nhãn hàng hóa (thay thế Nghị định 43/2017/NĐ-CP và Nghị định 111/2021/NĐ-CP), Điều 22 Nghị định số 128/2020/NĐ-CP về xử phạt vi phạm hành chính trong lĩnh vực hải quan, scratch/daily_notes.md ngày 16/10/2026 | 🟢 Duyệt Option A |"),
        
        ("| 79 | [W44] Thứ Ba (03/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Đế Chế Hàng Không FedEx 1973: Cú Đặt Cược 27.000 USD Tại Las Vegas Để Cứu Đội Bay Memphis, Khai Sinh Mô Hình 'Hub-and-Spoke' & Cuộc Cách Mạng Thay Đổi Tốc Độ Chuỗi Cung Ứng Toàn Cầu** | Bí mật đằng sau mô hình mạng lưới Hub-and-Spoke của Fred Smith: Tại sao máy bay chở gói hàng từ New York đi Boston lại phải bay vòng qua Memphis giữa đêm để tiết kiệm hàng tỷ USD chi phí logistics? | Từ bờ vực phá sản đến kỳ lân vận tải hàng không toàn cầu: Bài học quản trị dòng tiền sinh tử và nghệ thuật đặt cược vào tầm nhìn dài hạn cho các nhà sáng lập doanh nghiệp logistics Việt Nam. | Lịch sử hình thành Federal Express (1971-1973), Báo cáo Vận tải Hàng không Quốc tế (IATA Air Cargo Market Analysis), Nghiên cứu mô hình Hub-and-Spoke trường Wharton Business School & Bách khoa toàn thư vận tải Smithsonian | 🟡 Chờ duyệt |",
         "| 79 | [W44] Thứ Ba (03/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Đế Chế Hàng Không FedEx 1973: Cú Đặt Cược 27.000 USD Tại Las Vegas Để Cứu Đội Bay Memphis, Khai Sinh Mô Hình 'Hub-and-Spoke' & Cuộc Cách Mạng Thay Đổi Tốc Độ Chuỗi Cung Ứng Toàn Cầu** | Bí mật đằng sau mô hình mạng lưới Hub-and-Spoke của Fred Smith: Tại sao máy bay chở gói hàng từ New York đi Boston lại phải bay vòng qua Memphis giữa đêm để tiết kiệm hàng tỷ USD chi phí logistics? | Từ bờ vực phá sản đến kỳ lân vận tải hàng không toàn cầu: Bài học quản trị dòng tiền sinh tử và nghệ thuật đặt cược vào tầm nhìn dài hạn cho các nhà sáng lập doanh nghiệp logistics Việt Nam. | Lịch sử hình thành Federal Express (1971-1973), Báo cáo Vận tải Hàng không Quốc tế (IATA Air Cargo Market Analysis), Nghiên cứu mô hình Hub-and-Spoke trường Wharton Business School & Bách khoa toàn thư vận tải Smithsonian | 🟢 Duyệt Option A |"),
        
        ("| 80 | [W44] Thứ Tư (04/11/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Giao dịch thương mại 3 bên (Bên Bán VN - Bên Mua Đài Loan - Bên Nhận Mỹ): Hiển thị thông tin Shipper, Consignee và Notify Party trên Vận đơn (B/L) thế nào để Hải quan không bắt bẻ và không bị ngân hàng từ chối thanh toán quốc tế?** | Sơ đồ luồng chứng từ chuẩn cho thương mại ba bên: Phân định rạch ròi giữa 'Hợp đồng thương mại', 'Tờ khai xuất khẩu', 'Hóa đơn thương mại' và 'Vận đơn đường biển' giúp doanh nghiệp xuất khẩu Việt Nam bảo vệ quyền lợi trọn vẹn. | Tâm sự người làm nghề XNK: Khi khách nước ngoài yêu cầu sửa tên Consignee trên B/L khác với Tờ khai hải quan — Đừng vội làm theo nếu chưa kiểm tra 3 điều kiện pháp lý sống còn theo Điều 16 Thông tư 38/2015/TT-BTC! | Điều 16 & Điều 18 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Quy tắc Thực hành Thống nhất về Tín dụng Chứng từ (UCP 600), Incoterms 2020 & scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |",
         "| 80 | [W44] Thứ Tư (04/11/2026) | Facebook | CS — Case Study | Góc nhìn cá nhân Thúy | Bài Case Study + Ảnh | **Giao dịch thương mại 3 bên (Bên Bán VN - Bên Mua Đài Loan - Bên Nhận Mỹ): Hiển thị thông tin Shipper, Consignee và Notify Party trên Vận đơn (B/L) thế nào để Hải quan không bắt bẻ và không bị ngân hàng từ chối thanh toán quốc tế?** | Sơ đồ luồng chứng từ chuẩn cho thương mại ba bên: Phân định rạch ròi giữa 'Hợp đồng thương mại', 'Tờ khai xuất khẩu', 'Hóa đơn thương mại' và 'Vận đơn đường biển' giúp doanh nghiệp xuất khẩu Việt Nam bảo vệ quyền lợi trọn vẹn. | Tâm sự người làm nghề XNK: Khi khách nước ngoài yêu cầu sửa tên Consignee trên B/L khác với Tờ khai hải quan — Đừng vội làm theo nếu chưa kiểm tra 3 điều kiện pháp lý sống còn theo Điều 16 Thông tư 38/2015/TT-BTC! | Điều 16 & Điều 18 Thông tư số 38/2015/TT-BTC (sửa đổi, bổ sung tại Thông tư 39/2018/TT-BTC), Quy tắc Thực hành Thống nhất về Tín dụng Chứng từ (UCP 600), Incoterms 2020 & scratch/daily_notes.md ngày 16/10/2026 | 🟢 Duyệt Option A |"),
        
        ("| 81 | [W44] Thứ Năm (05/11/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 14 | **Học XNK 5 Phút Tập 14: 4 Bước Sử Dụng Vận Đơn Switch B/L Trong Thương Mại Quốc Tế Ba Bên Để Giữ Kín Bí Mật Giá Vốn, Khách Hàng Gốc & Không Bị Kẹt Bộ Chứng Từ Ngân Hàng** | Checklist 5 rủi ro chí mạng khi yêu cầu Hãng tàu đổi Bill (Switch B/L) và quy trình 3 bên thanh toán L/C/TT an toàn tuyệt đối cho doanh nghiệp thương mại trung gian. | Case study thực tế: Doanh nghiệp thương mại Việt Nam bán hàng sang Mỹ qua đối tác trung gian — Cách phát hành Switch B/L tại Cảng trung chuyển mà không để lộ danh tính nhà máy sản xuất. | Quy chuẩn Hướng dẫn Vận đơn Switch B/L của Liên đoàn Các hiệp hội Giao nhận Vận tải Quốc tế (FIATA), Quy tắc Hàng hải Quốc tế (Hague-Visby Rules), Điều 16 Thông tư 38/2015/TT-BTC & Cẩm nang nghiệp vụ Eureka Logistics | 🟡 Chờ duyệt |",
         "| 81 | [W44] Thứ Năm (05/11/2026) | YouTube | HT — Học Thực chiến | Trao giá trị thực chiến | Video Học XNK 5 phút Tập 14 | **Học XNK 5 Phút Tập 14: 4 Bước Sử Dụng Vận Đơn Switch B/L Trong Thương Mại Quốc Tế Ba Bên Để Giữ Kín Bí Mật Giá Vốn, Khách Hàng Gốc & Không Bị Kẹt Bộ Chứng Từ Ngân Hàng** | Checklist 5 rủi ro chí mạng khi yêu cầu Hãng tàu đổi Bill (Switch B/L) và quy trình 3 bên thanh toán L/C/TT an toàn tuyệt đối cho doanh nghiệp thương mại trung gian. | Case study thực tế: Doanh nghiệp thương mại Việt Nam bán hàng sang Mỹ qua đối tác trung gian — Cách phát hành Switch B/L tại Cảng trung chuyển mà không để lộ danh tính nhà máy sản xuất. | Quy chuẩn Hướng dẫn Vận đơn Switch B/L của Liên đoàn Các hiệp hội Giao nhận Vận tải Quốc tế (FIATA), Quy tắc Hàng hải Quốc tế (Hague-Visby Rules), Điều 16 Thông tư 38/2015/TT-BTC & Cẩm nang nghiệp vụ Eureka Logistics | 🟢 Duyệt Option A |"),
        
        ("| 82 | [W44] Thứ Sáu (06/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **eTax Mobile liên thông dữ liệu ngân hàng và sàn thương mại điện tử: Cảnh báo bẫy truy thu thuế nhà thầu khi doanh nghiệp chi trả hoa hồng, phí quảng cáo cho nền tảng xuyên biên giới — 'Quên' kê khai từ năm trước, giờ nhận thông báo phạt kèm tiền chậm nộp từng ngày!** | Bảng hướng dẫn 3 bước kê khai và nộp thay thuế nhà thầu cho đối tác nước ngoài (TikTok, Google, Meta) theo Thông tư 80/2021/TT-BTC: Cách rà soát hồ sơ bổ sung trước khi đoàn thanh tra thuế gõ cửa. | Góc nhìn thực tế từ diễn đàn thuế: Khi cá nhân và doanh nghiệp bất ngờ nhận thông báo nợ thuế trên eTax Mobile — Đừng tìm cách xóa nợ 'ảo', hãy bình tĩnh thực hiện quy trình giải trình 3 bước chuẩn cơ quan thuế! | Thông tư số 80/2021/TT-BTC hướng dẫn Luật Quản lý thuế, Thông tư số 100/2021/TT-BTC sửa đổi Thông tư 40/2021/TT-BTC về thuế hộ kinh doanh và sàn TMĐT, Nghị định 126/2020/NĐ-CP & scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |",
         "| 82 | [W44] Thứ Sáu (06/11/2026) | Facebook | TN — Tin Ngành | Problem-First | Bài tranh luận + Infographic | **eTax Mobile liên thông dữ liệu ngân hàng và sàn thương mại điện tử: Cảnh báo bẫy truy thu thuế nhà thầu khi doanh nghiệp chi trả hoa hồng, phí quảng cáo cho nền tảng xuyên biên giới — 'Quên' kê khai từ năm trước, giờ nhận thông báo phạt kèm tiền chậm nộp từng ngày!** | Bảng hướng dẫn 3 bước kê khai và nộp thay thuế nhà thầu cho đối tác nước ngoài (TikTok, Google, Meta) theo Thông tư 80/2021/TT-BTC: Cách rà soát hồ sơ bổ sung trước khi đoàn thanh tra thuế gõ cửa. | Góc nhìn thực tế từ diễn đàn thuế: Khi cá nhân và doanh nghiệp bất ngờ nhận thông báo nợ thuế trên eTax Mobile — Đừng tìm cách xóa nợ 'ảo', hãy bình tĩnh thực hiện quy trình giải trình 3 bước chuẩn cơ quan thuế! | Thông tư số 80/2021/TT-BTC hướng dẫn Luật Quản lý thuế, Thông tư số 100/2021/TT-BTC sửa đổi Thông tư 40/2021/TT-BTC về thuế hộ kinh doanh và sàn TMĐT, Nghị định 126/2020/NĐ-CP & scratch/daily_notes.md ngày 16/10/2026 | 🟢 Duyệt Option A |"),
        
        ("| 83 | [W44] Thứ Bảy (07/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Tuyến Đường Sắt Xuyên Á (Yiwu - Madrid Railway): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh (13.052 km), 3 Lần Đổi Khổ Ray Qua 8 Quốc Gia & Cuộc Đua Tái Định Hình Dòng Chảy Vận Tải Lục Địa Á - Âu** | Giải mã kỳ tích công nghệ vận tải liên vận đường sắt: Làm sao đoàn tàu container giữ nhiệt độ ổn định từ -40°C tuyết trắng Siberia đến +40°C sa mạc Kazakhstan, rút ngắn 50% thời gian so với đường biển? | Từ con đường tơ lụa sắt Yiwu - Madrid đến bài toán liên vận đường sắt Việt Nam - Trung Quốc - Châu Âu: Cơ hội bứt phá nào cho hàng xuất khẩu điện tử và nông sản Việt Nam khi các tuyến đường biển gặp rủi ro địa chính trị? | Báo cáo Thường niên Hiệp hội Đường sắt Quốc tế (UIC - International Union of Railways), Dữ liệu vận hành Yixinou Railway (Yiwu - Madrid freight route), Báo cáo Logistics Á - Âu của Ngân hàng Thế giới (World Bank) & Dữ liệu Phân tích Thị trường Eureka | 🟡 Chờ duyệt |",
         "| 83 | [W44] Thứ Bảy (07/11/2026) | YouTube | FF — Funfact & Insight | Series & Cliffhanger | Video Phân tích & Insight (8-12p) | **Kỳ Án Tuyến Đường Sắt Xuyên Á (Yiwu - Madrid Railway): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh (13.052 km), 3 Lần Đổi Khổ Ray Qua 8 Quốc Gia & Cuộc Đua Tái Định Hình Dòng Chảy Vận Tải Lục Địa Á - Âu** | Giải mã kỳ tích công nghệ vận tải liên vận đường sắt: Làm sao đoàn tàu container giữ nhiệt độ ổn định từ -40°C tuyết trắng Siberia đến +40°C sa mạc Kazakhstan, rút ngắn 50% thời gian so với đường biển? | Từ con đường tơ lụa sắt Yiwu - Madrid đến bài toán liên vận đường sắt Việt Nam - Trung Quốc - Châu Âu: Cơ hội bứt phá nào cho hàng xuất khẩu điện tử và nông sản Việt Nam khi các tuyến đường biển gặp rủi ro địa chính trị? | Báo cáo Thường niên Hiệp hội Đường sắt Quốc tế (UIC - International Union of Railways), Dữ liệu vận hành Yixinou Railway (Yiwu - Madrid freight route), Báo cáo Logistics Á - Âu của Ngân hàng Thế giới (World Bank) & Dữ liệu Phân tích Thị trường Eureka | 🟢 Duyệt Option A |"),
        
        ("| 84 | [W44] Chủ Nhật (08/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng hóa đơn đầu ra ghi chi tiết từng linh kiện máy móc: Bẫy kiểm tra sau thông quan và rủi ro bị cơ quan Thuế loại trừ chi phí hợp lý, bác khấu trừ thuế GTGT theo Nghị định 123/2020/NĐ-CP!** | Bộ khiên chắn 4 chứng từ tự vệ cho kế toán doanh nghiệp XNK: Cách lập bảng kê kỹ thuật, định mức tiêu hao và phiếu xuất kho chuẩn chỉ để chứng minh tính hợp pháp của hóa đơn đầu vào khi cơ quan thuế đối chiếu tự động. | Bài học xương máu từ một buổi giải trình thuế: Đừng bao giờ nghĩ 'miễn là có hóa đơn VAT là được khấu trừ' — Sự thiếu đồng nhất giữa tên hàng hóa đầu vào và đầu ra có thể khiến doanh nghiệp bị truy thu hàng trăm triệu đồng! | Nghị định số 123/2020/NĐ-CP quy định về hóa đơn, chứng từ; Thông tư số 78/2021/TT-BTC của Bộ Tài chính; Điều 4 Thông tư số 96/2015/TT-BTC về chi phí được trừ khi xác định thu nhập chịu thuế TNDN & scratch/daily_notes.md ngày 16/10/2026 | 🟡 Chờ duyệt |",
         "| 84 | [W44] Chủ Nhật (08/11/2026) | Facebook | CS — Case Study | Problem-First | Bài Case Study + Ảnh | **Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng hóa đơn đầu ra ghi chi tiết từng linh kiện máy móc: Bẫy kiểm tra sau thông quan và rủi ro bị cơ quan Thuế loại trừ chi phí hợp lý, bác khấu trừ thuế GTGT theo Nghị định 123/2020/NĐ-CP!** | Bộ khiên chắn 4 chứng từ tự vệ cho kế toán doanh nghiệp XNK: Cách lập bảng kê kỹ thuật, định mức tiêu hao và phiếu xuất kho chuẩn chỉ để chứng minh tính hợp pháp của hóa đơn đầu vào khi cơ quan thuế đối chiếu tự động. | Bài học xương máu từ một buổi giải trình thuế: Đừng bao giờ nghĩ 'miễn là có hóa đơn VAT là được khấu trừ' — Sự thiếu đồng nhất giữa tên hàng hóa đầu vào và đầu ra có thể khiến doanh nghiệp bị truy thu hàng trăm triệu đồng! | Nghị định số 123/2020/NĐ-CP quy định về hóa đơn, chứng từ; Thông tư số 78/2021/TT-BTC của Bộ Tài chính; Điều 4 Thông tư số 96/2015/TT-BTC về chi phí được trừ khi xác định thu nhập chịu thuế TNDN & scratch/daily_notes.md ngày 16/10/2026 | 🟢 Duyệt Option A |")
    ]
    
    for old_txt, new_txt in w44_replacements:
        content_updated = content_updated.replace(old_txt, new_txt)
        
    with open(WEEKLY_CAL_MD, "w", encoding="utf-8") as f:
        f.write(content_updated)
    print("SUCCESS: Cập nhật thành công file reports/plan1/weekly_calendar.md!")

# 4. CẬP NHẬT CONTENT_STATUS_REPORT.MD
if os.path.exists(STATUS_REP_MD):
    with open(STATUS_REP_MD, "r", encoding="utf-8") as f:
        rep_content = f.read()
    
    # Cập nhật số liệu tổng quan
    rep_content_updated = rep_content.replace(
        "## 📌 1. TỔNG QUAN TIẾN ĐỘ SẢN XUẤT TOÀN GIAI ĐOẠN (W33 - W43 & DAILY NOTES HOT TOPICS)\n\n- **Tổng số chủ đề trên Lịch Biên tập (Sheet 1):** 77 chủ đề (W33 đến W43, từ 17/08/2026 đến 01/11/2026)\n- **Tình trạng Phê duyệt Ý tưởng (Checkpoint 1):** 77/77 chủ đề (`🟢 Duyệt Option A` 100%)\n- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**\n  - **Kênh Facebook (Sheet 2):** 82 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 4 bài W43 + 38 bài HOT Daily Notes). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check SSoT.\n  - **Kênh YouTube (Sheet 3):** 53 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 3 kịch bản W43 + 20 kịch bản HOT Daily Notes). 100% có kịch bản lời thoại đầy đủ, Storyboard 4 cột (Audio/Visual/SFX) và Cliffhanger Outro kết nối tập sau.\n- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 135/135 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`.\n- **Rà soát Ghi chú Cần sửa đổi:** 100% các bài viết từ W33 đến W43 đều đạt chuẩn SSoT Zero-Trust, không có bài nào bị đánh dấu `🔴 REVISION REQUIRED` hoặc có ghi chú yêu cầu chỉnh sửa từ Chị Thúy.",
        "## 📌 1. TỔNG QUAN TIẾN ĐỘ SẢN XUẤT TOÀN GIAI ĐOẠN (W33 - W44 & DAILY NOTES HOT TOPICS)\n\n- **Tổng số chủ đề trên Lịch Biên tập (Sheet 1):** 84 chủ đề (W33 đến W44, từ 17/08/2026 đến 08/11/2026)\n- **Tình trạng Phê duyệt Ý tưởng (Checkpoint 1):** 84/84 chủ đề (`🟢 Duyệt Option A` 100%)\n- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**\n  - **Kênh Facebook (Sheet 2):** 86 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 4 bài W43 + 4 bài W44 + 38 bài HOT Daily Notes). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check SSoT.\n  - **Kênh YouTube (Sheet 3):** 56 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 3 kịch bản W43 + 3 kịch bản W44 + 20 kịch bản HOT Daily Notes). 100% có kịch bản lời thoại đầy đủ, Storyboard 4 cột (Audio/Visual/SFX) và Cliffhanger Outro kết nối tập sau.\n- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 142/142 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`.\n- **Rà soát Ghi chú Cần sửa đổi:** 100% các bài viết từ W33 đến W44 đều đạt chuẩn SSoT Zero-Trust, không có bài nào bị đánh dấu `🔴 REVISION REQUIRED` hoặc có ghi chú yêu cầu chỉnh sửa từ Chị Thúy."
    )
    
    # Bổ sung 4 dòng Facebook W44 vào Bảng Facebook
    fb_w44_table_rows = """| 83 | `PO-2026-W44-78` | [W44] Thứ Hai (02/11/2026) | Nghị định 37/2026/NĐ-CP siết chặt nhãn gốc: Đừng tin câu 'Tao xuất sang Việt Nam nhiều lần rồi không sao' của Shipper ngoại — Lô máy móc tiền tỷ kẹt cứng tại Hải Phòng vì tem gốc không có 'Made in China' và nguy cơ bị phạt đến 60 triệu kèm chế tài buộc tái xuất! | [DRAFT-FB-PO-2026-W44-78.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-78.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 84 | `PO-2026-W44-80` | [W44] Thứ Tư (04/11/2026) | Giao dịch thương mại 3 bên (Bên Bán VN - Bên Mua Đài Loan - Bên Nhận Mỹ): Hiển thị thông tin Shipper, Consignee và Notify Party trên Vận đơn (B/L) thế nào để Hải quan không bắt bẻ và không bị ngân hàng từ chối thanh toán quốc tế? | [DRAFT-FB-PO-2026-W44-80.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-80.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 85 | `PO-2026-W44-82` | [W44] Thứ Sáu (06/11/2026) | eTax Mobile liên thông dữ liệu ngân hàng và sàn thương mại điện tử: Cảnh báo bẫy truy thu thuế nhà thầu khi doanh nghiệp chi trả hoa hồng, phí quảng cáo cho nền tảng xuyên biên giới — 'Quên' kê khai từ năm trước, giờ nhận thông báo phạt kèm tiền chậm nộp từng ngày! | [DRAFT-FB-PO-2026-W44-82.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-82.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 86 | `PO-2026-W44-84` | [W44] Chủ Nhật (08/11/2026) | Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng hóa đơn đầu ra ghi chi tiết từng linh kiện máy móc: Bẫy kiểm tra sau thông quan và rủi ro bị cơ quan Thuế loại trừ chi phí hợp lý, bác khấu trừ thuế GTGT theo Nghị định 123/2020/NĐ-CP! | [DRAFT-FB-PO-2026-W44-84.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-84.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
"""
    # Thay thế dòng cuối của bảng Facebook W43
    fb_target_w43_end = "| 82 | `PO-2026-W43-77` | [W43] Chủ Nhật (01/11/2026) | Công văn 22411/CHQ-GSQL & Điều 47a Luật Hải quan: Cạm bẫy giao dịch Xuất nhập khẩu tại chỗ khi thiếu 'Văn bản chỉ định giao nhận của thương nhân nước ngoài' — Hàng chục tờ khai bị đình chỉ thông quan và nguy cơ truy thu thuế GTGT 10%! | [DRAFT-FB-PO-2026-W43-77.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W43-77.md) | `🟢 PASSED` | **Sẵn sàng đăng** |"
    if fb_target_w43_end in rep_content_updated and "PO-2026-W44-78" not in rep_content_updated:
        rep_content_updated = rep_content_updated.replace(fb_target_w43_end, fb_target_w43_end + "\n" + fb_w44_table_rows)

    # Bổ sung 3 dòng YouTube W44 vào Bảng YouTube
    yt_w44_table_rows = """| 54 | `PO-2026-W44-79` | [W44] Thứ Ba (03/11/2026) | Kỳ Án Đế Chế Hàng Không FedEx 1973: Cú Đặt Cược 27.000 USD Tại Las Vegas Để Cứu Đội Bay Memphis, Khai Sinh Mô Hình 'Hub-and-Spoke' & Cuộc Cách Mạng Thay Đổi Tốc Độ Chuỗi Cung Ứng Toàn Cầu | [DRAFT-YT-SERIES-PO-2026-W44-79.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-79.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 55 | `PO-2026-W44-81` | [W44] Thứ Năm (05/11/2026) | Học XNK 5 Phút Tập 14: 4 Bước Sử Dụng Vận Đơn Switch B/L Trong Thương Mại Quốc Tế Ba Bên Để Giữ Kín Bí Mật Giá Vốn, Khách Hàng Gốc & Không Bị Kẹt Bộ Chứng Từ Ngân Hàng | [DRAFT-YT-SERIES-PO-2026-W44-81.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-81.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 56 | `PO-2026-W44-83` | [W44] Thứ Bảy (07/11/2026) | Kỳ Án Tuyến Đường Sắt Xuyên Á (Yiwu - Madrid Railway): Chuyến Tàu Chở Hàng Dài Nhất Hành Tinh (13.052 km), 3 Lần Đổi Khổ Ray Qua 8 Quốc Gia & Cuộc Đua Tái Định Hình Dòng Chảy Vận Tải Lục Địa Á - Âu | [DRAFT-YT-SERIES-PO-2026-W44-83.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-83.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
"""
    yt_target_w43_end = "| 53 | `PO-2026-W43-76` | [W43] Thứ Bảy (31/10/2026) | Đế Chế Cảng Biển Tự Động Hóa Maasvlakte 2 (Rotterdam): Cỗ Máy Logistics 'Không Một Bóng Người' Bốc Xếp, Vận Hành Bằng 100% Điện Gió & Cuộc Cách Mạng Xanh Đang Thay Đổi Cuộc Đua Cảng Biển Thế Giới 2026 | [DRAFT-YT-SERIES-PO-2026-W43-76.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W43-76.md) | `🟢 PASSED` | **Sẵn sàng dựng** |"
    if yt_target_w43_end in rep_content_updated and "PO-2026-W44-79" not in rep_content_updated:
        rep_content_updated = rep_content_updated.replace(yt_target_w43_end, yt_target_w43_end + "\n" + yt_w44_table_rows)

    with open(STATUS_REP_MD, "w", encoding="utf-8") as f:
        f.write(rep_content_updated)
    print("SUCCESS: Cập nhật thành công file reports/plan1/content_status_report.md!")

print("\n=== HOÀN TẤT ĐỒNG BỘ TOÀN DIỆN GIAI ĐOẠN 2 CHO TUẦN W44! ===")
