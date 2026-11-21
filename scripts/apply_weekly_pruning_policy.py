# -*- coding: utf-8 -*-
r"""
Áp dụng Quy chế Xoay Vòng Hàng Tuần & Đào Thải Tồn Đọng 2 Tuần (Weekly Rolling Pipeline & 2-Week Stale Pruning):
File: C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx
- Mỗi tuần phát triển 9 chủ đề mới mỗi CM (+36 chủ đề/tuần).
- Tự động xóa các chủ đề đề xuất từ 2 tuần trước mà KHÔNG ĐƯỢC CHỊ THÚY BÁO DUYỆT VIẾT KỊCH BẢN CỤ THỂ.
- Bảo lưu vĩnh viễn 100% các chủ đề đã duyệt viết kịch bản hoặc đã hoàn thành.
- Thêm 2 cột quản lý: Col 14 'Tuần Đề Xuất' và Col 15 'Hạn Đào Thải (2 Tuần)'.
- Đánh lại STT liên tục, định dạng chuẩn mực, tuân thủ RULE 10 (Zero Backup - atomic_save).
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
EXCEL_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
TMP_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.tmp.xlsx")

def apply_pruning_policy():
    print("=== BẮT ĐẦU ÁP DỤNG QUY CHẾ XOAY VÒNG HÀNG TUẦN & ĐÀO THẢI TỒN ĐỌNG 2 TUẦN ===")
    
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Không tìm thấy file: {EXCEL_PATH}")
        return

    wb = openpyxl.load_workbook(EXCEL_PATH)
    
    # Định dạng styling
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
    
    regular_font = Font(name="Arial", size=10, bold=False, color="000000")
    bold_font = Font(name="Arial", size=10, bold=True, color="000000")
    approved_font = Font(name="Arial", size=10, bold=True, color="0E6251")
    pending_font = Font(name="Arial", size=10, bold=True, color="B9770E")
    
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    approved_fill = PatternFill(start_color="D4EFDF", end_color="D4EFDF", fill_type="solid")
    pending_fill = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")
    
    left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    cm_sheets = [
        "CM1_VAN_BAN_NONG",
        "CM2_DICH_LUAT_TIENG_NGUOI",
        "CM3_LICH_SU_THUONG_MAI",
        "CM4_TAM_TINH_THUC_CHIEN"
    ]
    
    sheet_stats = {}

    for s_name in cm_sheets:
        ws = wb[s_name]
        print(f"\n📂 Đang xử lý Sheet: {s_name} (Số dòng hiện tại: {ws.max_row})")
        
        # 1. Cập nhật Tiêu đề Header tại Row 4 (thêm Col 14 và Col 15)
        h14 = ws.cell(row=4, column=14, value="Tuần Đề Xuất")
        h14.font = header_font
        h14.fill = header_fill
        h14.alignment = center_align
        h14.border = thin_border
        
        h15 = ws.cell(row=4, column=15, value="Hạn Đào Thải (2 Tuần)")
        h15.font = header_font
        h15.fill = header_fill
        h15.alignment = center_align
        h15.border = thin_border
        
        ws.column_dimensions['N'].width = 16
        ws.column_dimensions['O'].width = 28

        # 2. Thu thập dữ liệu các dòng hiện tại
        rows_data = []
        for r in range(5, ws.max_row + 1):
            row_dict = {
                "values": [ws.cell(r, c).value for c in range(1, 14)], # col 1 to 13
                "code": ws.cell(r, 2).value,
                "title": ws.cell(r, 3).value,
                "status": ws.cell(r, 12).value,
                "link": ws.cell(r, 13).value
            }
            rows_data.append(row_dict)

        # 3. Phân loại Giữ (Keep) vs Đào thải (Prune)
        kept_rows = []
        pruned_rows = []

        for item in rows_data:
            code = item["code"] or ""
            status = item["status"] or ""
            
            # Kiểm tra xem mã thuộc đợt nào
            is_batch1 = False
            try:
                num = int(code.split("CD")[-1])
                if num <= 9:
                    is_batch1 = True
            except:
                pass
            
            is_approved = any(kw in status for kw in ["HOÀN THÀNH", "ĐÃ DUYỆT", "Duyệt"])
            
            if is_batch1 and not is_approved:
                # Đề xuất từ 2 tuần trước mà KHÔNG được Chị Thúy duyệt viết kịch bản -> ĐÀO THẢI
                pruned_rows.append(item)
                print(f"   🗑️ ĐÀO THẢI (Quá 2 tuần chưa duyệt): {code} - {item['title'][:45]}...")
            else:
                # Thuộc diện giữ lại
                if is_batch1 and is_approved:
                    item["proposed_week"] = "W44/2026"
                    item["deadline_info"] = "🛡️ Bảo lưu vĩnh viễn (Đã duyệt)"
                else: # Batch 2 (CD10 - CD18) mới đưa vào tuần W46
                    item["proposed_week"] = "W46/2026"
                    item["deadline_info"] = "⏳ Hết hạn sau Tuần W48"
                kept_rows.append(item)

        print(f"   -> Kết quả {s_name}: Giữ lại {len(kept_rows)} chủ đề, Đào thải {len(pruned_rows)} chủ đề tồn đọng.")
        sheet_stats[s_name] = {"kept": len(kept_rows), "pruned": len(pruned_rows)}

        # 4. Xóa trắng các dòng dữ liệu cũ từ dòng 5 trở đi
        max_r = ws.max_row
        if max_r >= 5:
            ws.delete_rows(5, max_r - 4)

        # 5. Ghi lại các dòng được GIỮ LẠI với STT liên tục
        for idx, item in enumerate(kept_rows, start=1):
            target_row = 4 + idx
            row_fill = zebra_fill if target_row % 2 == 0 else white_fill
            
            # Cập nhật lại STT mới ở cột 1
            vals = item["values"]
            vals[0] = idx # STT liên tục từ 1
            
            # Ghi 13 cột ban đầu
            for c_idx, val in enumerate(vals, start=1):
                cell = ws.cell(row=target_row, column=c_idx, value=val)
                cell.border = thin_border
                
                if c_idx == 12: # Trạng thái
                    status_str = str(val or "")
                    if "HOÀN THÀNH" in status_str or "ĐÃ DUYỆT" in status_str:
                        cell.font = approved_font
                        cell.fill = approved_fill
                    else:
                        cell.font = pending_font
                        cell.fill = pending_fill
                    cell.alignment = center_align
                elif c_idx in [1, 2, 4, 5]:
                    cell.font = bold_font if c_idx in [1, 2] else regular_font
                    cell.fill = row_fill
                    cell.alignment = center_align
                else:
                    cell.font = regular_font
                    cell.fill = row_fill
                    cell.alignment = left_align

            # Ghi Col 14 (Tuần Đề Xuất)
            c14 = ws.cell(row=target_row, column=14, value=item["proposed_week"])
            c14.font = bold_font
            c14.fill = row_fill
            c14.alignment = center_align
            c14.border = thin_border

            # Ghi Col 15 (Hạn Đào Thải)
            c15 = ws.cell(row=target_row, column=15, value=item["deadline_info"])
            if "Bảo lưu" in item["deadline_info"]:
                c15.font = approved_font
                c15.fill = row_fill
            else:
                c15.font = pending_font
                c15.fill = row_fill
            c15.alignment = center_align
            c15.border = thin_border

            ws.row_dimensions[target_row].height = 140.25

    # 6. Cập nhật Sheet 1: TONG_QUAN_LICH_PHAT_SONG
    if "TONG_QUAN_LICH_PHAT_SONG" in wb.sheetnames:
        ws_tq = wb["TONG_QUAN_LICH_PHAT_SONG"]
        total_kept = sum(s["kept"] for s in sheet_stats.values())
        total_pruned = sum(s["pruned"] for s in sheet_stats.values())
        
        ws_tq.cell(2, 1, value=f"BẢNG ĐIỀU HÀNH MASTER SERIES YOUTUBE THUYTALK — QUY CHẾ XOAY VÒNG HÀNG TUẦN & ĐÀO THẢI TỒN ĐỌNG 2 TUẦN (HIỆN CÓ {total_kept} CHỦ ĐỀ)")
        
        # Cập nhật số lượng tại bảng điều hướng (Rows 6-9, Col 4)
        cm_row_map = {
            "CM1_VAN_BAN_NONG": 6,
            "CM2_DICH_LUAT_TIENG_NGUOI": 7,
            "CM3_LICH_SU_THUONG_MAI": 8,
            "CM4_TAM_TINH_THUC_CHIEN": 9
        }
        for s_name, r_idx in cm_row_map.items():
            k_count = sheet_stats[s_name]["kept"]
            ws_tq.cell(r_idx, 4, value=f"{k_count} Chủ đề (Đã lọc đào thải)")

        print(f"\n✅ Đã cập nhật Sheet 1 TONG_QUAN_LICH_PHAT_SONG (Tổng giữ: {total_kept}, Tổng đào thải: {total_pruned}).")

    # 7. Lưu an toàn atomic_save trên ổ C:
    wb.save(TMP_PATH)
    wb.close()
    os.replace(TMP_PATH, EXCEL_PATH)
    print("\n🎉 HOÀN TẤT THỰC THI QUY CHẾ ĐÀO THẢI & XOAY VÒNG HÀNG TUẦN TRÊN ONEDRIVE!")

if __name__ == "__main__":
    apply_pruning_policy()
