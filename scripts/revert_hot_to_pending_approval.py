# -*- coding: utf-8 -*-
"""
Script khôi phục chuẩn xác quy trình kiểm duyệt kịch bản YouTube của Chị Thúy:
1. Đưa 5 chủ đề nóng W45 (HOT37 -> HOT41) trên Master Series Excel về đúng trạng thái `Chờ Chị Thúy duyệt`.
2. Xóa các Hyperlink Cột 13 (vì kịch bản chưa được Chị Thúy phê duyệt trên Excel).
3. Thiết lập Cột 15 thành `⏳ Hết hạn sau Tuần W47` (cơ chế đào thải sau 2 tuần nếu không được duyệt).
4. Chuyển các file kịch bản vật lý tạm thời ra khỏi thư mục OneDrive `BRANDING/youtube/` về `scratch/drafts_pending_thuy_approval/`.
5. Đảm bảo OneDrive chỉ lưu trữ và liên kết các kịch bản ĐÃ ĐƯỢC CHỊ THÚY DUYỆT.
"""

import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
ONEDRIVE_YT_DIR = os.path.join(ONEDRIVE_DIR, "youtube")
EXCEL_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
STAGING_DIR = r"d:\AI AGENT THUY\AI branding\scratch\drafts_pending_thuy_approval"

os.makedirs(STAGING_DIR, exist_ok=True)

# Styling chuẩn cho trạng thái chờ duyệt
font_pending = Font(name="Arial", size=10, bold=True, color="B9770E")
fill_none = PatternFill(fill_type=None)
align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

def atomic_save_workbook(wb, filepath):
    tmp_path = filepath + ".tmp.xlsx"
    if os.path.exists(tmp_path):
        try:
            os.remove(tmp_path)
        except Exception:
            pass
    wb.save(tmp_path)
    wb.close()
    os.replace(tmp_path, filepath)
    print(f"✅ Đã ghi đè an toàn vào: {filepath}")

def main():
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Không tìm thấy file: {EXCEL_PATH}")
        return

    # 1. Chuyển các file .md chưa được duyệt ra khỏi OneDrive về thư mục chờ (staging)
    hot_files = [
        "CM1-HOT37_Thong_tu_104_Kiem_tra_chuyen_nganh.md",
        "CM1-HOT38_Nghi_dinh_336_Mot_cua_quoc_gia_Giai_doan_3.md",
        "CM1-HOT40_Cua_khau_thong_minh_AGV_Thong_tu_128.md",
        "CM2-HOT41_To_khai_treo_ECUS5_Doi_soat_thue.md",
        "CM4-HOT39_Drewry_WCI_Blank_Sailings_Dam_phan_cuoc.md"
    ]
    
    print("📦 [1] Đang di chuyển các kịch bản chưa duyệt ra khỏi OneDrive BRANDING/youtube/...")
    for fname in hot_files:
        src = os.path.join(ONEDRIVE_YT_DIR, fname)
        dst = os.path.join(STAGING_DIR, fname)
        if os.path.exists(src):
            shutil.move(src, dst)
            print(f"   -> Đã chuyển về kho lưu tạm: {fname}")
        elif os.path.exists(dst):
            print(f"   -> Đã nằm sẵn trong kho lưu tạm: {fname}")

    # 2. Mở Master Excel và chuyển trạng thái về `Chờ Chị Thúy duyệt`
    print(f"📂 [2] Đang tải workbook: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=False)

    hot_codes_map = {
        "CM1-HOT37": ("CM1_VAN_BAN_NONG", 17),
        "CM1-HOT38": ("CM1_VAN_BAN_NONG", 18),
        "CM1-HOT40": ("CM1_VAN_BAN_NONG", 19),
        "CM2-HOT41": ("CM2_DICH_LUAT_TIENG_NGUOI", 18),
        "CM4-HOT39": ("CM4_TAM_TINH_THUC_CHIEN", 20),
    }

    for code, (sname, expected_row) in hot_codes_map.items():
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        
        target_row = None
        for r in range(5, ws.max_row + 1):
            if ws.cell(r, 2).value == code:
                target_row = r
                break
                
        if target_row:
            print(f"   🔄 Khôi phục trạng thái chờ duyệt cho [{code}] tại {sname} dòng {target_row}...")
            # Cột 12: Trạng Thái Xem Xét -> Chờ Chị Thúy duyệt
            ws.cell(target_row, 12, "Chờ Chị Thúy duyệt").alignment = align_center
            ws.cell(target_row, 12).font = font_pending
            ws.cell(target_row, 12).fill = fill_none
            
            # Cột 13: Link Kịch Bản -> None (Chờ duyệt)
            ws.cell(target_row, 13, None)
            
            # Cột 15: Hạn Đào Thải (2 Tuần) -> ⏳ Hết hạn sau Tuần W47
            ws.cell(target_row, 15, "⏳ Hết hạn sau Tuần W47").alignment = align_center
            ws.cell(target_row, 15).font = font_pending
            ws.cell(target_row, 15).fill = fill_none

    # 3. Cập nhật Sheet TONG_QUAN_LICH_PHAT_SONG
    if "TONG_QUAN_LICH_PHAT_SONG" in wb.sheetnames:
        ws_tq = wb["TONG_QUAN_LICH_PHAT_SONG"]
        cm1_count = wb["CM1_VAN_BAN_NONG"].max_row - 4
        cm2_count = wb["CM2_DICH_LUAT_TIENG_NGUOI"].max_row - 4
        cm3_count = wb["CM3_LICH_SU_THUONG_MAI"].max_row - 4
        cm4_count = wb["CM4_TAM_TINH_THUC_CHIEN"].max_row - 4
        
        ws_tq.cell(6, 4, f"{cm1_count} Chủ đề (3 đề xuất W45 chờ Thúy duyệt)")
        ws_tq.cell(7, 4, f"{cm2_count} Chủ đề (1 đề xuất W45 chờ Thúy duyệt)")
        ws_tq.cell(8, 4, f"{cm3_count} Chủ đề (Đã lọc đào thải)")
        ws_tq.cell(9, 4, f"{cm4_count} Chủ đề (1 đề xuất W45 chờ Thúy duyệt)")
        print(f"📊 Đã cập nhật Sheet TONG_QUAN: CM1={cm1_count}, CM2={cm2_count}, CM3={cm3_count}, CM4={cm4_count}")

    # 4. Ghi file an toàn
    atomic_save_workbook(wb, EXCEL_PATH)
    print("🎉 Hoàn tất khôi phục quyền duyệt cho Chị Thúy trên Master Series Excel!")

if __name__ == "__main__":
    main()
