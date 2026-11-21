# -*- coding: utf-8 -*-
"""
HỆ THỐNG AGENT SẢN XUẤT NỘI DUNG GIAI ĐOẠN 2 — TUẦN W44 (STT 78 - 84)
Thực thi toàn bộ:
1. Tạo file PO vật lý trong scratch/production_orders/
2. Tạo file Draft vật lý trong scratch/production_orders/ & reports/plan1/
3. Tạo file Fact-Check trong scratch/production_orders/
4. Cập nhật Excel Weekly_Content_Calendar.xlsx (Sheet 1, Sheet 2, Sheet 3) với atomic_save (Zero Backup)
5. Cập nhật báo cáo tổng kết reports/plan1/content_status_report.md
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")
YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")
EXCEL_PATH = os.path.join(BASE_DIR, "reports", "plan1", "Weekly_Content_Calendar.xlsx")
STATUS_REPORT_PATH = os.path.join(BASE_DIR, "reports", "plan1", "content_status_report.md")

os.makedirs(PO_DIR, exist_ok=True)
os.makedirs(FB_DIR, exist_ok=True)
os.makedirs(YT_DIR, exist_ok=True)

# Import dữ liệu
sys.path.append(os.path.join(BASE_DIR, "scripts"))
import w44_data_all as d

ALL_ITEMS = [
    (d.DATA_78, d.PO_78_MD, d.DRAFT_FB_78, d.FACT_CHECK_78_MD),
    (d.DATA_79, d.PO_79_MD, d.DRAFT_YT_79, d.FACT_CHECK_79_MD),
    (d.DATA_80, d.PO_80_MD, d.DRAFT_FB_80, d.FACT_CHECK_80_MD),
    (d.DATA_81, d.PO_81_MD, d.DRAFT_YT_81, d.FACT_CHECK_81_MD),
    (d.DATA_82, d.PO_82_MD, d.DRAFT_FB_82, d.FACT_CHECK_82_MD),
    (d.DATA_83, d.PO_83_MD, d.DRAFT_YT_83, d.FACT_CHECK_83_MD),
    (d.DATA_84, d.PO_84_MD, d.DRAFT_FB_84, d.FACT_CHECK_84_MD),
]

def step1_write_physical_files():
    print("=== BƯỚC 1: XUẤT CÁC TỆP VẬT LÝ VÀO THƯ MỤC ===")
    for data, po_md, draft_content, fc_md in ALL_ITEMS:
        po_id = data["po_id"]
        platform = data["platform"]
        
        # 1. File PO
        po_path = os.path.join(PO_DIR, f"{po_id}.md")
        with open(po_path, "w", encoding="utf-8") as f:
            f.write(po_md)
        print(f"  -> Đã tạo PO: {po_path}")
        
        # 2. File Draft
        if platform == "Facebook":
            draft_name = f"DRAFT-FB-{po_id}.md"
            draft_path = os.path.join(PO_DIR, draft_name)
            with open(draft_path, "w", encoding="utf-8") as f:
                f.write(draft_content)
            
            # Copy sang reports/plan1/facebook
            fb_rep_path = os.path.join(FB_DIR, f"FB-{po_id}.md")
            with open(fb_rep_path, "w", encoding="utf-8") as f:
                f.write(draft_content)
            print(f"  -> Đã tạo FB Draft: {draft_path} & {fb_rep_path}")
        else:
            draft_name = f"DRAFT-YT-SERIES-{po_id}.md"
            draft_path = os.path.join(PO_DIR, draft_name)
            with open(draft_path, "w", encoding="utf-8") as f:
                f.write(draft_content)
            
            # Copy sang reports/plan1/video_scripts
            yt_rep_path = os.path.join(YT_DIR, f"YS-{po_id}.md")
            with open(yt_rep_path, "w", encoding="utf-8") as f:
                f.write(draft_content)
            print(f"  -> Đã tạo YT Script: {draft_path} & {yt_rep_path}")
            
        # 3. File Fact-Check
        fc_path = os.path.join(PO_DIR, f"FACT-CHECK-{po_id}.md")
        with open(fc_path, "w", encoding="utf-8") as f:
            f.write(fc_md)
        print(f"  -> Đã tạo Fact-Check: {fc_path}")

def step2_update_excel_calendar():
    print("\n=== BƯỚC 2: CẬP NHẬT WEEKLY_CONTENT_CALENDAR.XLSX ===")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    
    font_family = "Segoe UI"
    reg_font = Font(name=font_family, size=10)
    bold_font = Font(name=font_family, size=10, bold=True)
    passed_font = Font(name=font_family, size=10, bold=True, color="1E8449")
    link_font = Font(name=font_family, size=10, color="0000FF", underline="single")
    
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )
    
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    passed_fill = PatternFill(start_color="EAFAF1", end_color="EAFAF1", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9FBFD", end_color="F9FBFD", fill_type="solid")
    
    # ----------------------------------------------------
    # 1. Cập nhật Sheet 1 (Weekly Calendar): Duyệt STT 78-84
    # ----------------------------------------------------
    ws1 = wb["Weekly Calendar"]
    print("  -> Cập nhật Sheet 1 (Weekly Calendar)...")
    for r in range(5, ws1.max_row + 1):
        stt = ws1.cell(r, 1).value
        if stt in [78, 79, 80, 81, 82, 83, 84]:
            c11 = ws1.cell(r, 11, value="🟢 Duyệt Option A")
            c11.font = passed_font
            c11.fill = passed_fill
            c11.alignment = center_align
            c11.border = thin_border
            print(f"     STT {stt} -> 🟢 Duyệt Option A")
            
    # ----------------------------------------------------
    # 2. Cập nhật Sheet 2 (Facebook): Thêm STT 78, 80, 82, 84
    # ----------------------------------------------------
    ws2 = wb["Facebook"]
    print("  -> Cập nhật Sheet 2 (Facebook)...")
    existing_fb_pos = [ws2.cell(r, 2).value for r in range(5, ws2.max_row + 1)]
    
    # Tìm STT cao nhất hiện tại
    max_fb_stt = 0
    for r in range(5, ws2.max_row + 1):
        val = ws2.cell(r, 1).value
        try:
            if int(val) > max_fb_stt:
                max_fb_stt = int(val)
        except:
            pass
            
    fb_items = [item for item in ALL_ITEMS if item[0]["platform"] == "Facebook"]
    for data, _, draft_content, _ in fb_items:
        po_id = data["po_id"]
        if po_id in existing_fb_pos:
            print(f"     PO {po_id} đã tồn tại trong Sheet 2. Bỏ qua ghi đè.")
            continue
            
        max_fb_stt += 1
        target_row = ws2.max_row + 1
        row_fill = zebra_fill if target_row % 2 == 0 else white_fill
        
        draft_filename = f"DRAFT-FB-{po_id}.md"
        hyperlink_formula = f'=HYPERLINK("file:///d:/AI AGENT THUY/AI branding/scratch/production_orders/{draft_filename}", "👉 Xem Bài Viết Chi Tiết")'
        
        vals = [
            max_fb_stt,
            po_id,
            data["day"],
            data["title"],
            data["seo"],
            hyperlink_formula,
            draft_content,
            data["visual_brief"],
            "🟢 PASSED",
            data["ssot_report"],
            data["pub_status"]
        ]
        
        for col_idx, val in enumerate(vals, start=1):
            cell = ws2.cell(target_row, col_idx, value=val)
            cell.border = thin_border
            
            if col_idx in [1, 2, 3]:
                cell.font = bold_font if col_idx == 1 else reg_font
                cell.alignment = center_align
                cell.fill = row_fill
            elif col_idx == 6:
                cell.font = link_font
                cell.alignment = center_align
                cell.fill = row_fill
            elif col_idx == 9:
                cell.font = passed_font
                cell.fill = passed_fill
                cell.alignment = center_align
            elif col_idx == 11:
                cell.font = bold_font
                cell.alignment = center_align
                cell.fill = row_fill
            else:
                cell.font = reg_font
                cell.alignment = left_align
                cell.fill = row_fill
                
        print(f"     Đã thêm dòng {target_row}: STT {max_fb_stt} - {po_id}")
        
    # ----------------------------------------------------
    # 3. Cập nhật Sheet 3 (YouTube): Thêm STT 79, 81, 83
    # ----------------------------------------------------
    ws3 = wb["YouTube"]
    print("  -> Cập nhật Sheet 3 (YouTube)...")
    existing_yt_pos = [ws3.cell(r, 2).value for r in range(5, ws3.max_row + 1)]
    
    max_yt_stt = 0
    for r in range(5, ws3.max_row + 1):
        val = ws3.cell(r, 1).value
        try:
            if int(val) > max_yt_stt:
                max_yt_stt = int(val)
        except:
            pass
            
    yt_items = [item for item in ALL_ITEMS if item[0]["platform"] == "YouTube"]
    for data, _, draft_content, _ in yt_items:
        po_id = data["po_id"]
        if po_id in existing_yt_pos:
            print(f"     PO {po_id} đã tồn tại trong Sheet 3. Bỏ qua ghi đè.")
            continue
            
        max_yt_stt += 1
        target_row = ws3.max_row + 1
        row_fill = zebra_fill if target_row % 2 == 0 else white_fill
        
        draft_filename = f"DRAFT-YT-SERIES-{po_id}.md"
        hyperlink_formula = f'=HYPERLINK("file:///d:/AI AGENT THUY/AI branding/scratch/production_orders/{draft_filename}", "👉 Xem Kịch Bản Chi Tiết")'
        
        vals = [
            max_yt_stt,
            po_id,
            data["series"],
            data["category"],
            data["seo"],
            hyperlink_formula,
            draft_content,
            data["visual_brief"],
            "🟢 PASSED",
            data["ssot_report"],
            data["pub_status"]
        ]
        
        for col_idx, val in enumerate(vals, start=1):
            cell = ws3.cell(target_row, col_idx, value=val)
            cell.border = thin_border
            
            if col_idx in [1, 2, 4]:
                cell.font = bold_font if col_idx == 1 else reg_font
                cell.alignment = center_align
                cell.fill = row_fill
            elif col_idx == 6:
                cell.font = link_font
                cell.alignment = center_align
                cell.fill = row_fill
            elif col_idx == 9:
                cell.font = passed_font
                cell.fill = passed_fill
                cell.alignment = center_align
            elif col_idx == 11:
                cell.font = bold_font
                cell.alignment = center_align
                cell.fill = row_fill
            else:
                cell.font = reg_font
                cell.alignment = left_align
                cell.fill = row_fill
                
        print(f"     Đã thêm dòng {target_row}: STT {max_yt_stt} - {po_id}")

    # Áp dụng cơ chế atomic_save và TUYỆT ĐỐI KHÔNG TẠO BACKUP FILE
    tmp_path = EXCEL_PATH + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, EXCEL_PATH)
    print("  -> Đã lưu Excel nguyên tử thành công (Zero Backup).")

def step3_update_status_report():
    print("\n=== BƯỚC 3: CẬP NHẬT REPORTS/PLAN1/CONTENT_STATUS_REPORT.MD ===")
    
    # Đọc báo cáo hiện tại
    with open(STATUS_REPORT_PATH, "r", encoding="utf-8") as f:
        content = f.read()

    # Cập nhật số liệu tổng quan
    old_overview = """- **Tổng số chủ đề trên Lịch Biên tập (Sheet 1):** 77 chủ đề (W33 đến W43, từ 17/08/2026 đến 01/11/2026)
- **Tình trạng Phê duyệt Ý tưởng (Checkpoint 1):** 77/77 chủ đề (`🟢 Duyệt Option A` 100%)
- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 82 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 4 bài W43 + 38 bài HOT Daily Notes). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check SSoT.
  - **Kênh YouTube (Sheet 3):** 53 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 3 kịch bản W43 + 20 kịch bản HOT Daily Notes). 100% có kịch bản lời thoại đầy đủ, Storyboard 4 cột (Audio/Visual/SFX) và Cliffhanger Outro kết nối tập sau.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 135/135 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`.
- **Rà soát Ghi chú Cần sửa đổi:** 100% các bài viết từ W33 đến W43 đều đạt chuẩn SSoT Zero-Trust, không có bài nào bị đánh dấu `🔴 REVISION REQUIRED` hoặc có ghi chú yêu cầu chỉnh sửa từ Chị Thúy."""

    new_overview = """- **Tổng số chủ đề trên Lịch Biên tập (Sheet 1):** 84 chủ đề (W33 đến W44, từ 17/08/2026 đến 08/11/2026)
- **Tình trạng Phê duyệt Ý tưởng (Checkpoint 1):** 84/84 chủ đề (`🟢 Duyệt Option A` 100%)
- **Tình trạng Sản xuất Chi tiết (Giai đoạn 2):**
  - **Kênh Facebook (Sheet 2):** 86 bài viết hoàn chỉnh (36 bài W33-W41 + 4 bài W42 + 4 bài W43 + 4 bài W44 + 38 bài HOT Daily Notes). 100% đã có bản thảo chi tiết, file PO vật lý, file .md nháp, Media brief và Báo cáo Fact-Check SSoT.
  - **Kênh YouTube (Sheet 3):** 56 kịch bản video hoàn chỉnh (27 kịch bản W33-W41 + 3 kịch bản W42 + 3 kịch bản W43 + 3 kịch bản W44 + 20 kịch bản HOT Daily Notes). 100% có kịch bản lời thoại đầy đủ, Storyboard 4 cột (Audio/Visual/SFX) và Cliffhanger Outro kết nối tập sau.
- **Thẩm định 4 lớp (Agent 08 Fact-Checker):** 142/142 sản phẩm đạt `🟢 PASSED (100% ĐẠT CHUẨN XUẤT BẢN)`.
- **Rà soát Ghi chú Cần sửa đổi:** 100% các bài viết từ W33 đến W44 đều đạt chuẩn SSoT Zero-Trust, không có bài nào bị đánh dấu `🔴 REVISION REQUIRED` hoặc có ghi chú yêu cầu chỉnh sửa từ Chị Thúy."""

    if old_overview in content:
        content = content.replace(old_overview, new_overview)
    else:
        # Thay thế tiêu đề W33 - W43 thành W33 - W44
        content = content.replace("W33 - W43", "W33 - W44")
        content = content.replace("82 BÀI", "86 BÀI")
        content = content.replace("53 KỊCH BẢN", "56 KỊCH BẢN")

    # Bổ sung 4 bài Facebook vào bảng BẢNG TỔNG HỢP NỘI DUNG KÊNH FACEBOOK
    fb_additions = """| 83 | `PO-2026-W44-78` | [W44] Thứ Hai (02/11/2026) | Nghị định 37/2026/NĐ-CP siết chặt nhãn gốc: Đừng tin câu 'Tao xuất sang Việt Nam nhiều lần rồi không sao' của Shipper ngoại — Lô máy móc tiền tỷ kẹt cứng tại Hải Phòng vì tem gốc không có 'Made in China' và nguy cơ bị phạt đến 60 triệu kèm chế tài buộc tái xuất! | [DRAFT-FB-PO-2026-W44-78.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-78.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 84 | `PO-2026-W44-80` | [W44] Thứ Tư (04/11/2026) | Giao dịch thương mại 3 bên (Bên Bán VN - Bên Mua Đài Loan - Bên Nhận Mỹ): Hiển thị thông tin Shipper, Consignee và Notify Party trên Vận đơn (B/L) thế nào để Hải quan không bắt bẻ và không bị ngân hàng từ chối thanh toán quốc tế? | [DRAFT-FB-PO-2026-W44-80.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-80.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 85 | `PO-2026-W44-82` | [W44] Thứ Sáu (06/11/2026) | eTax Mobile liên thông dữ liệu ngân hàng và sàn thương mại điện tử: Cảnh báo bẫy truy thu thuế nhà thầu khi doanh nghiệp chi trả hoa hồng, phí quảng cáo cho nền tảng xuyên biên giới — 'Quên' kê khai từ năm trước, giờ nhận thông báo phạt kèm tiền chậm nộp từng ngày! | [DRAFT-FB-PO-2026-W44-82.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-82.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
| 86 | `PO-2026-W44-84` | [W44] Chủ Nhật (08/11/2026) | Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng hóa đơn đầu ra ghi chi tiết từng linh kiện máy móc: Bẫy kiểm tra sau thông quan và rủi ro bị cơ quan Thuế loại trừ chi phí hợp lý, bác khấu trừ thuế GTGT theo Nghị định 123/2020/NĐ-CP! | [DRAFT-FB-PO-2026-W44-84.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-FB-PO-2026-W44-84.md) | `🟢 PASSED` | **Sẵn sàng đăng** |
"""
    # Tìm dòng PO-2026-W43-77 trong bảng FB
    target_fb_anchor = "| 82 | `PO-2026-W43-77`"
    if target_fb_anchor in content and "PO-2026-W44-78" not in content:
        # Chèn vào sau dòng 82
        idx = content.find(target_fb_anchor)
        end_line_idx = content.find("\n", idx) + 1
        content = content[:end_line_idx] + fb_additions + content[end_line_idx:]
        print("  -> Đã chèn 4 bài Facebook W44 vào Báo cáo.")

    # Bổ sung 3 video YouTube vào bảng BẢNG TỔNG HỢP NỘI DUNG KÊNH YOUTUBE
    yt_additions = """| 54 | `PO-2026-W44-79` | [W44] Series Kỳ Án Hàng Không Logistics — Tập 1: Cú Đặt Cược 27.000 USD Tại Las Vegas Của Fred Smith & Cuộc Cách Mạng FedEx 1973 | `FF — Funfact & Insight` | [DRAFT-YT-SERIES-PO-2026-W44-79.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-79.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 55 | `PO-2026-W44-81` | [W44] Series Học XNK 5 Phút — Tập 14: 4 Bước Sử Dụng Vận Đơn Switch B/L Trong Thương Mại Quốc Tế Ba Bên | `HT — Học Thực chiến` | [DRAFT-YT-SERIES-PO-2026-W44-81.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-81.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
| 56 | `PO-2026-W44-83` | [W44] Series Đột Phá Chuỗi Cung Ứng — Tập 16: Kỳ Án Tuyến Đường Sắt Xuyên Á (Yiwu - Madrid Railway) | `FF — Funfact & Insight` | [DRAFT-YT-SERIES-PO-2026-W44-83.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/production_orders/DRAFT-YT-SERIES-PO-2026-W44-83.md) | `🟢 PASSED` | **Sẵn sàng dựng** |
"""
    target_yt_anchor = "| 53 | `PO-2026-W43-76`"
    if target_yt_anchor in content and "PO-2026-W44-79" not in content:
        idx = content.find(target_yt_anchor)
        end_line_idx = content.find("\n", idx) + 1
        content = content[:end_line_idx] + yt_additions + content[end_line_idx:]
        print("  -> Đã chèn 3 video YouTube W44 vào Báo cáo.")

    with open(STATUS_REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(content)
    print("  -> Đã lưu cập nhật content_status_report.md thành công.")

if __name__ == "__main__":
    step1_write_physical_files()
    step2_update_excel_calendar()
    step3_update_status_report()
    print("\n🎉 HOÀN TẤT 100% SẢN XUẤT NỘI DUNG GIAI ĐOẠN 2 CHO CÁC TOPICS CHƯA CÓ BÀI (TUẦN W44)!")
