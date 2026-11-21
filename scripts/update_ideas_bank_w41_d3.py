# -*- coding: utf-8 -*-
"""
================================================================================
SCRIPT: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS & FEEDBACK LOOP PHÁP LUẬT
PHIÊN BẢN: W41 DAY 3 — 08/10/2026 (STT 131 - 135)
TUÂN THỦ: wf06_ideas_bank_update.md & RÀO CHẮN KỸ THUẬT url_validator.py
QUY TẮC VĨNH VIỄN: TUYỆT ĐỐI KHÔNG TẠO FILE BACKUP (.XLSX)
================================================================================
"""

import sys
import os
import time
import io
from datetime import datetime

# Đảm bảo tiếng Việt in không lỗi trên console Windows
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except AttributeError:
    pass

sys.path.insert(0, r"d:\AI AGENT THUY\AI branding\lib")
sys.path.insert(0, r"d:\AI AGENT THUY\AI branding\scripts")

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from url_validator import enforce_strict_validation

from w41_d3_tn import TIN_NGANH_W41_D3
from w41_d3_cs import CASE_STUDY_W41_D3
from w41_d3_ht import HOC_THUC_CHIEN_W41_D3
from w41_d3_ff import FUNFACT_W41_D3
from w41_d3_legal import NEW_LEGAL_DOCS_W41_D3

REPORTS_DIR = r"d:\AI AGENT THUY\AI branding\reports"

# ------------------------------------------------------------------------------
# CẤU HÌNH STYLING CHUẨN DOANH NGHIỆP EUREKA
# ------------------------------------------------------------------------------
FONT_FAMILY = "Segoe UI"
BODY_FONT = Font(name=FONT_FAMILY, size=10, color="1F2937")
BOLD_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="111827")
STATUS_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="047857")

WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
STATUS_FILL = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="E5E7EB"),
    right=Side(style="thin", color="E5E7EB"),
    top=Side(style="thin", color="E5E7EB"),
    bottom=Side(style="thin", color="E5E7EB")
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

def atomic_save(wb, file_path, max_retries=5):
    """Lưu workbook an toàn qua file tạm để chống xung đột OneDrive sync với cơ chế thử lại (Tuyệt đối không tạo file backup)"""
    tmp_path = file_path + ".tmp.xlsx"
    wb.save(tmp_path)
    for attempt in range(max_retries):
        try:
            if os.path.exists(file_path):
                os.replace(tmp_path, file_path)
            else:
                os.rename(tmp_path, file_path)
            return True
        except PermissionError:
            time.sleep(0.5)
        except Exception:
            time.sleep(0.5)
    if os.path.exists(tmp_path):
        try:
            os.remove(tmp_path)
        except Exception:
            pass
    raise PermissionError(f"Không thể ghi đè file sau {max_retries} lần thử: {file_path}")


def run():
    print("=" * 80)
    print("🛡️ BẮT ĐẦU WORKFLOW 06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (W41 DAY 3 — 08/10/2026)")
    print("=" * 80)

    # 1. RÀO CHẮN KỸ THUẬT: VALIDATE TẤT CẢ 20 URLS TRƯỚC KHI GHI
    all_ideas = TIN_NGANH_W41_D3 + CASE_STUDY_W41_D3 + HOC_THUC_CHIEN_W41_D3 + FUNFACT_W41_D3
    enforce_strict_validation(all_ideas)
    print("\n🟢 TOÀN BỘ 20 LINK ĐÃ ĐƯỢC XÁC THỰC THÀNH CÔNG (PASSED 100%)!")

    # 2. CẬP NHẬT KHO Ý TƯỞNG TIN NGÀNH (TN)
    tn_path = os.path.join(REPORTS_DIR, "Tin_Nganh_Ideas_Bank_Eureka.xlsx")
    wb_tn = openpyxl.load_workbook(tn_path)
    s_tn = wb_tn.active
    stt_map_tn = {s_tn.cell(row=r, column=1).value: r for r in range(5, s_tn.max_row + 1)}
    print(f"\n📂 Đang đồng bộ Tin Ngành: 5 ý tưởng (STT 131 - 135)...")
    for item in TIN_NGANH_W41_D3:
        stt = item["stt"]
        r = stt_map_tn.get(stt, s_tn.max_row + 1)
        fill = ZEBRA_FILL if stt % 2 == 0 else WHITE_FILL
        row_vals = [
            item["stt"], item["week"], item["title"], item["agency"], item["timeline"],
            item["core_diff"], item["affected_target"], item["risk_penalty"], item["market_reaction"],
            item["pro_argument"], item["con_argument"], item["intl_ref"], item["url"],
            item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = s_tn.cell(row=r, column=c_idx, value=val)
            if c_idx == 1:
                cell.font = BOLD_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx == 2:
                cell.font = BODY_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx == 14:
                cell.font = STATUS_FONT
                cell.alignment = ALIGN_CENTER
                cell.fill = STATUS_FILL
            else:
                cell.font = BODY_FONT
                cell.alignment = ALIGN_LEFT
            if c_idx != 14:
                cell.fill = fill
            cell.border = THIN_BORDER
    atomic_save(wb_tn, tn_path)
    print(f"✅ Tin Ngành đã cập nhật xong! (Tổng dòng: {s_tn.max_row}, STT cuối: {s_tn.cell(s_tn.max_row, 1).value})")

    # 3. CẬP NHẬT KHO Ý TƯỞNG CASE STUDY (CS)
    cs_path = os.path.join(REPORTS_DIR, "Case_Study_Ideas_Bank_Eureka.xlsx")
    wb_cs = openpyxl.load_workbook(cs_path)
    s_cs = wb_cs.active
    stt_map_cs = {s_cs.cell(row=r, column=1).value: r for r in range(5, s_cs.max_row + 1)}
    print(f"\n📂 Đang đồng bộ Case Study: 5 ý tưởng (STT 131 - 135)...")
    for item in CASE_STUDY_W41_D3:
        stt = item["stt"]
        r = stt_map_cs.get(stt, s_cs.max_row + 1)
        fill = ZEBRA_FILL if stt % 2 == 0 else WHITE_FILL
        row_vals = [
            item["stt"], item["week"], item["title"], item["stakeholders"], item["context"],
            item["damage"], item["root_cause"], item["solution"], item["opposing_views"],
            item["lesson"], item["url"], item["series"], item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = s_cs.cell(row=r, column=c_idx, value=val)
            if c_idx == 1:
                cell.font = BOLD_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx in (2, 12):
                cell.font = BODY_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx == 13:
                cell.font = STATUS_FONT
                cell.alignment = ALIGN_CENTER
                cell.fill = STATUS_FILL
            else:
                cell.font = BODY_FONT
                cell.alignment = ALIGN_LEFT
            if c_idx != 13:
                cell.fill = fill
            cell.border = THIN_BORDER
    atomic_save(wb_cs, cs_path)
    print(f"✅ Case Study đã cập nhật xong! (Tổng dòng: {s_cs.max_row}, STT cuối: {s_cs.cell(s_cs.max_row, 1).value})")

    # 4. CẬP NHẬT KHO Ý TƯỞNG HỌC THỰC CHIẾN (HT)
    ht_path = os.path.join(REPORTS_DIR, "Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx")
    wb_ht = openpyxl.load_workbook(ht_path)
    s_ht = wb_ht.active
    stt_map_ht = {s_ht.cell(row=r, column=1).value: r for r in range(5, s_ht.max_row + 1)}
    print(f"\n📂 Đang đồng bộ Học Thực Chiến: 5 ý tưởng (STT 131 - 135)...")
    for item in HOC_THUC_CHIEN_W41_D3:
        stt = item["stt"]
        r = stt_map_ht.get(stt, s_ht.max_row + 1)
        fill = ZEBRA_FILL if stt % 2 == 0 else WHITE_FILL
        row_vals = [
            item["stt"], item["week"], item["title"], item["workflow"], item["formula"],
            item["top3_mistakes"], item["mnemonic_tip"], item["comparison"], item["persona"],
            item["checklist"], item["url"], item["series"], item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = s_ht.cell(row=r, column=c_idx, value=val)
            if c_idx == 1:
                cell.font = BOLD_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx in (2, 12):
                cell.font = BODY_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx == 13:
                cell.font = STATUS_FONT
                cell.alignment = ALIGN_CENTER
                cell.fill = STATUS_FILL
            else:
                cell.font = BODY_FONT
                cell.alignment = ALIGN_LEFT
            if c_idx != 13:
                cell.fill = fill
            cell.border = THIN_BORDER
    atomic_save(wb_ht, ht_path)
    print(f"✅ Học Thực Chiến đã cập nhật xong! (Tổng dòng: {s_ht.max_row}, STT cuối: {s_ht.cell(s_ht.max_row, 1).value})")

    # 5. CẬP NHẬT KHO Ý TƯỞNG FUNFACT (FF)
    ff_path = os.path.join(REPORTS_DIR, "Funfact_Ideas_Bank_Eureka.xlsx")
    wb_ff = openpyxl.load_workbook(ff_path)
    s_ff = wb_ff.active
    stt_map_ff = {s_ff.cell(row=r, column=1).value: r for r in range(5, s_ff.max_row + 1)}
    print(f"\n📂 Đang đồng bộ Funfact: 5 ý tưởng (STT 131 - 135)...")
    for item in FUNFACT_W41_D3:
        stt = item["stt"]
        r = stt_map_ff.get(stt, s_ff.max_row + 1)
        fill = ZEBRA_FILL if stt % 2 == 0 else WHITE_FILL
        row_vals = [
            item["stt"], item["week"], item["title"], item["hook"], item["history"],
            item["context"], item["modern_impact"], item["intl_comp"], item["business_lesson"],
            item["url"], item["series"], item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = s_ff.cell(row=r, column=c_idx, value=val)
            if c_idx == 1:
                cell.font = BOLD_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx in (2, 11):
                cell.font = BODY_FONT
                cell.alignment = ALIGN_CENTER
            elif c_idx == 12:
                cell.font = STATUS_FONT
                cell.alignment = ALIGN_CENTER
                cell.fill = STATUS_FILL
            else:
                cell.font = BODY_FONT
                cell.alignment = ALIGN_LEFT
            if c_idx != 12:
                cell.fill = fill
            cell.border = THIN_BORDER
    atomic_save(wb_ff, ff_path)
    print(f"✅ Funfact đã cập nhật xong! (Tổng dòng: {s_ff.max_row}, STT cuối: {s_ff.cell(s_ff.max_row, 1).value})")

    # 6. FEEDBACK LOOP: CẬP NHẬT BẢNG VĂN BẢN PHÁP LUẬT XNK
    legal_path = os.path.join(REPORTS_DIR, "Danh_sach_Van_ban_Phap_luat_XNK.xlsx")
    wb_leg = openpyxl.load_workbook(legal_path)
    s_leg = wb_leg.active
    existing_codes = {s_leg.cell(row=r, column=2).value for r in range(2, s_leg.max_row + 1)}
    last_stt = s_leg.cell(row=s_leg.max_row, column=1).value or 368
    docs_to_add = [doc for doc in NEW_LEGAL_DOCS_W41_D3 if doc["code"] not in existing_codes]

    print(f"\n⚖️ THỰC HIỆN FEEDBACK LOOP: Bổ sung {len(docs_to_add)} văn bản luật mới vào {legal_path}...")
    for idx, doc in enumerate(docs_to_add, start=1):
        stt_val = last_stt + idx
        row_data = [
            stt_val,
            doc["code"],
            doc["title"],
            doc["type"],
            doc["agency"],
            doc["issue_date"],
            doc["effective_date"],
            doc["expiry_date"],
            doc["replaced_doc"],
            doc["brief_analysis"],
            doc["comparative_change"],
            doc["update_date"]
        ]
        s_leg.append(row_data)
        r = s_leg.max_row
        fill = ZEBRA_FILL if stt_val % 2 == 0 else WHITE_FILL
        for c_idx in range(1, len(row_data) + 1):
            cell = s_leg.cell(row=r, column=c_idx)
            cell.font = BOLD_FONT if c_idx in (1, 2) else BODY_FONT
            cell.alignment = ALIGN_CENTER if c_idx in (1, 6, 7, 8, 12) else ALIGN_LEFT
            cell.fill = fill
            cell.border = THIN_BORDER

    if docs_to_add:
        atomic_save(wb_leg, legal_path)
    print(f"✅ Bảng Pháp luật XNK đã cập nhật xong! (Tổng dòng: {s_leg.max_row}, STT cuối: {s_leg.cell(s_leg.max_row, 1).value})")

    print("\n" + "=" * 80)
    print("🎉 HOÀN THÀNH WORKFLOW 06: ĐÃ CẬP NHẬT ĐẦY ĐỦ 20 Ý TƯỞNG MỚI (STT 131 - 135) CHO 4 PILLARS")
    print("🛡️ TOÀN BỘ 20 URL ĐÃ QUA THẨM ĐỊNH GATEKEEPER 100% TRƯỚC KHI LƯU.")
    print("🚫 TUYỆT ĐỐI KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (.XLSX) THEO ĐÚNG NGUYÊN TẮC.")
    print("=" * 80)


if __name__ == "__main__":
    run()
