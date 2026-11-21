#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EXPORT THUYTALK STORYBOARD & IMAGE PROMPTS (CM3 SPECIALIST)
Tuân thủ tuyệt đối RULE 13, RULE 14, RULE 16 trong GEMINI.md
Xuất đồng thời:
1. <Ten_Du_An>_Kich_Ban_Flow.csv (5 cột chuẩn Flow)
2. <Ten_Du_An>_Tao_Anh_Hang_Loat.csv (5 cột chuẩn tạo ảnh)
3. <Ten_Du_An>_Studio.xlsx (4 Sheet chuẩn Studio)
"""

import os
import sys
import csv
import json
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def safe_replace(src, dst, max_retries=5, delay=0.5):
    import time
    import subprocess
    for attempt in range(max_retries):
        if os.path.exists(dst):
            try:
                os.remove(dst)
            except Exception:
                try:
                    subprocess.run(['powershell', '-Command', f'Remove-Item -LiteralPath "{dst}" -Force -ErrorAction SilentlyContinue'], check=False)
                except Exception:
                    pass
        try:
            os.replace(src, dst)
            return
        except Exception:
            time.sleep(delay)
    # Fallback
    try:
        shutil.copy2(src, dst)
        if os.path.exists(src):
            os.remove(src)
    except Exception as e:
        print(f"Warning safe_replace fallback: {e}")

def export_thuytalk_project(project_data, output_dir="."):
    """
    project_data structure:
    {
        "project_name": "Malcolm_McLean_Container",
        "title": "Malcolm McLean và Cuộc Cách Mạng Thùng Container 1956",
        "category": "CM3 - Ký sự Lịch sử Thương mại & Logistics",
        "visual_style": "Hand-painted editorial 2.5D animation, cinematic lighting...",
        "character_anchor": "Cute modern female expert Thuy, high ponytail, navy blazer...",
        "voiceover_persona": "Thúy (Eureka Logistics) - giọng nói mộc mạc, tâm tình, tự nhiên",
        "scenes": [
            {
                "stt": 1,
                "title": "Scene 01",
                "prompt": "...",
                "dialogue": "...",
                "duration": 10
            },
            ...
        ],
        "image_prompts": [
            {
                "stt": 1,
                "title": "Khởi đầu",
                "prompt": "...",
                "shot_type": "Trung cảnh (Medium Shot)",
                "note": "..."
            },
            ...
        ]
    }
    """
    project_name = project_data.get("project_name", "ThuyTalk_CM3")
    os.makedirs(output_dir, exist_ok=True)
    
    # Silently close excel if open
    try:
        import subprocess
        subprocess.run(['powershell', '-Command', 'Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force'], check=False)
    except Exception:
        pass

    csv_flow_path = os.path.join(output_dir, f"{project_name}_Kich_Ban_Flow.csv")
    csv_img_path = os.path.join(output_dir, f"{project_name}_Tao_Anh_Hang_Loat.csv")
    xlsx_path = os.path.join(output_dir, f"{project_name}_Studio.xlsx")
    
    tmp_flow_csv = os.path.join(output_dir, f"{project_name}_Kich_Ban_Flow.tmp.csv")
    tmp_img_csv = os.path.join(output_dir, f"{project_name}_Tao_Anh_Hang_Loat.tmp.csv")
    tmp_xlsx_path = os.path.join(output_dir, f"{project_name}_Studio.tmp.xlsx")

    # 1. XUẤT CSV KỊCH BẢN FLOW (RULE 13 - 6 CỘT CHUẨN CÓ GÓC QUAY)
    img_shot_map = {img.get("stt"): img.get("shot_type") for img in project_data.get("image_prompts", [])}
    with open(tmp_flow_csv, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(["STT", "Tiêu đề", "Mô tả phân cảnh (Prompt)", "Góc quay (Shot Type)", "Lời thoại (Dialogue)", "Thời lượng (giây)"])
        for s in project_data.get("scenes", []):
            stt = s.get("stt", "")
            shot_type = s.get("shot_type") or img_shot_map.get(stt, "Trung cảnh (Medium Shot)")
            writer.writerow([
                stt,
                s.get("title", ""),
                s.get("prompt", ""),
                shot_type,
                s.get("dialogue", ""),
                s.get("duration", 10)
            ])
    safe_replace(tmp_flow_csv, csv_flow_path)
    print(f"✅ Đã xuất CSV Flow: {csv_flow_path}")

    # 2. XUẤT CSV TẠO ẢNH HÀNG LOẠT (RULE 14 - 5 CỘT)
    with open(tmp_img_csv, mode="w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"])
        for img in project_data.get("image_prompts", []):
            writer.writerow([
                img.get("stt", ""),
                img.get("title", ""),
                img.get("prompt", ""),
                img.get("shot_type", "Trung cảnh (Medium Shot)"),
                img.get("note", "")
            ])
    safe_replace(tmp_img_csv, csv_img_path)
    print(f"✅ Đã xuất CSV Tạo ảnh: {csv_img_path}")

    # 3. XUẤT EXCEL STUDIO 4 SHEET (RULE 16G)
    wb = openpyxl.Workbook()
    # Sheet 1: KICH_BAN_FLOW
    ws1 = wb.active
    ws1.title = "KICH_BAN_FLOW"
    
    # Sheet 2: TAO_ANH_HANG_LOAT
    ws2 = wb.create_sheet("TAO_ANH_HANG_LOAT")
    
    # Sheet 3: KICH_BAN_VOICEOVER_FULL
    ws3 = wb.create_sheet("KICH_BAN_VOICEOVER_FULL")
    
    # Sheet 4: TONG_QUAN_PROJECT_STYLE
    ws4 = wb.create_sheet("TONG_QUAN_PROJECT_STYLE")

    # Styles
    font_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=10)
    font_bold = Font(name="Arial", size=10, bold=True)
    font_title = Font(name="Arial", size=14, bold=True, color="1A365D")

    fill_flow = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")  # Navy Blue
    fill_img = PatternFill(start_color="975A16", end_color="975A16", fill_type="solid")    # Amber Bronze
    fill_accent = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid") # Royal Blue

    thin_border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # POPULATE SHEET 1: KICH_BAN_FLOW (6 CỘT CHUẨN CÓ GÓC QUAY)
    headers_s1 = ["STT", "Tiêu đề", "Mô tả phân cảnh (Prompt)", "Góc quay (Shot Type)", "Lời thoại (Dialogue)", "Thời lượng (giây)"]
    ws1.append(headers_s1)
    for col_num in range(1, 7):
        cell = ws1.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_flow
        cell.alignment = align_center

    for r_idx, s in enumerate(project_data.get("scenes", []), start=2):
        stt = s.get("stt", r_idx - 1)
        shot_type = s.get("shot_type") or img_shot_map.get(stt, "Trung cảnh (Medium Shot)")
        ws1.append([
            stt,
            s.get("title", f"Scene {r_idx-1:02d}"),
            s.get("prompt", ""),
            shot_type,
            s.get("dialogue", ""),
            s.get("duration", 10)
        ])
        ws1.cell(row=r_idx, column=1).alignment = align_center
        ws1.cell(row=r_idx, column=2).alignment = align_center
        ws1.cell(row=r_idx, column=3).alignment = align_left
        ws1.cell(row=r_idx, column=4).alignment = align_center
        ws1.cell(row=r_idx, column=5).alignment = align_left
        ws1.cell(row=r_idx, column=6).alignment = align_center
        for c in range(1, 7):
            ws1.cell(row=r_idx, column=c).font = font_body
            ws1.cell(row=r_idx, column=c).border = thin_border

    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 14
    ws1.column_dimensions['C'].width = 55
    ws1.column_dimensions['D'].width = 24
    ws1.column_dimensions['E'].width = 45
    ws1.column_dimensions['F'].width = 16

    # POPULATE SHEET 2: TAO_ANH_HANG_LOAT
    headers_s2 = ["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"]
    ws2.append(headers_s2)
    for col_num in range(1, 6):
        cell = ws2.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_img
        cell.alignment = align_center

    for r_idx, img in enumerate(project_data.get("image_prompts", []), start=2):
        ws2.append([
            img.get("stt", r_idx - 1),
            img.get("title", f"Frame {r_idx-1:02d}"),
            img.get("prompt", ""),
            img.get("shot_type", "Trung cảnh (Medium Shot)"),
            img.get("note", "")
        ])
        ws2.cell(row=r_idx, column=1).alignment = align_center
        ws2.cell(row=r_idx, column=2).alignment = align_left
        ws2.cell(row=r_idx, column=3).alignment = align_left
        ws2.cell(row=r_idx, column=4).alignment = align_center
        ws2.cell(row=r_idx, column=5).alignment = align_left
        for c in range(1, 6):
            ws2.cell(row=r_idx, column=c).font = font_body
            ws2.cell(row=r_idx, column=c).border = thin_border

    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 25
    ws2.column_dimensions['C'].width = 60
    ws2.column_dimensions['D'].width = 24
    ws2.column_dimensions['E'].width = 35

    # POPULATE SHEET 3: KICH_BAN_VOICEOVER_FULL
    ws3.append(["Mốc Thời Gian", "Phân Cảnh", "Lời Thoại Đọc Thuần Văn Nói (Narration)", "Chỉ Dẫn Biểu Cảm & Nhịp Thở"])
    for col_num in range(1, 5):
        cell = ws3.cell(row=1, column=col_num)
        cell.font = font_header
        cell.fill = fill_accent
        cell.alignment = align_center

    current_sec = 0
    for r_idx, s in enumerate(project_data.get("scenes", []), start=2):
        dur = s.get("duration", 10)
        time_str = f"{current_sec//60:02d}:{current_sec%60:02d} - {(current_sec+dur)//60:02d}:{(current_sec+dur)%60:02d}"
        current_sec += dur
        ws3.append([
            time_str,
            s.get("title", f"Scene {r_idx-1:02d}"),
            s.get("dialogue", ""),
            s.get("emotion_note", "Nói tự nhiên, nhấn giọng từ khóa, ngắt nghỉ đúng nhịp.")
        ])
        ws3.cell(row=r_idx, column=1).alignment = align_center
        ws3.cell(row=r_idx, column=2).alignment = align_center
        ws3.cell(row=r_idx, column=3).alignment = align_left
        ws3.cell(row=r_idx, column=4).alignment = align_left
        for c in range(1, 5):
            ws3.cell(row=r_idx, column=c).font = font_body
            ws3.cell(row=r_idx, column=c).border = thin_border

    ws3.column_dimensions['A'].width = 18
    ws3.column_dimensions['B'].width = 14
    ws3.column_dimensions['C'].width = 65
    ws3.column_dimensions['D'].width = 35

    # POPULATE SHEET 4: TONG_QUAN_PROJECT_STYLE
    overview_data = [
        ["HẠNG MỤC", "THÔNG TIN CHI TIẾT"],
        ["Tên dự án / Video", project_data.get("title", "")],
        ["Chuyên mục", project_data.get("category", "CM3 - Ký sự Lịch sử Thương mại ThuyTalk")],
        ["Thời lượng tổng cộng", f"{current_sec} giây (~{current_sec//60} phút {current_sec%60} giây)"],
        ["Tổng số phân cảnh", f"{len(project_data.get('scenes', []))} Scenes"],
        ["Bộ gen nhân vật (Persona)", project_data.get("character_anchor", "")],
        ["Phong cách nghệ thuật (Visual Anchor)", project_data.get("visual_style", "")],
        ["Tông giọng thuyết minh (Voiceover)", project_data.get("voiceover_persona", "")],
        ["Quy chuẩn Google Flow", "6 cột chuẩn: STT, Tiêu đề, Mô tả phân cảnh (Prompt), Góc quay (Shot Type), Lời thoại (Dialogue), Thời lượng (giây)"],
        ["Quy chuẩn Tạo ảnh hàng loạt", "5 cột chuẩn: STT, Tiêu đề, Prompt tạo ảnh, Góc quay, Ghi chú"]
    ]
    for r_idx, row in enumerate(overview_data, start=1):
        ws4.append(row)
        if r_idx == 1:
            ws4.cell(row=r_idx, column=1).font = font_header
            ws4.cell(row=r_idx, column=2).font = font_header
            ws4.cell(row=r_idx, column=1).fill = fill_flow
            ws4.cell(row=r_idx, column=2).fill = fill_flow
            ws4.cell(row=r_idx, column=1).alignment = align_center
            ws4.cell(row=r_idx, column=2).alignment = align_center
        else:
            ws4.cell(row=r_idx, column=1).font = font_bold
            ws4.cell(row=r_idx, column=2).font = font_body
            ws4.cell(row=r_idx, column=1).alignment = align_left
            ws4.cell(row=r_idx, column=2).alignment = align_left
            ws4.cell(row=r_idx, column=1).border = thin_border
            ws4.cell(row=r_idx, column=2).border = thin_border

    ws4.column_dimensions['A'].width = 30
    ws4.column_dimensions['B'].width = 75

    # LƯU ATOMIC TRÁNH LỖI ONEDRIVE SYNC & TUYỆT ĐỐI KHÔNG TẠO BACKUP
    wb.save(tmp_xlsx_path)
    safe_replace(tmp_xlsx_path, xlsx_path)
    print(f"✅ Đã xuất Excel Studio: {xlsx_path}")

    # ĐỒNG BỘ MCP FLOW NẾU CÓ THƯ MỤC
    mcp_flow_dir = r"C:\Users\thuyn\Documents\Flow_Storyboards"
    if os.path.exists(mcp_flow_dir):
        import shutil
        shutil.copy2(csv_flow_path, os.path.join(mcp_flow_dir, os.path.basename(csv_flow_path)))
        shutil.copy2(csv_img_path, os.path.join(mcp_flow_dir, os.path.basename(csv_img_path)))
        shutil.copy2(xlsx_path, os.path.join(mcp_flow_dir, os.path.basename(xlsx_path)))
        print(f"✅ Đã đồng bộ sang MCP Flow_Storyboards: {mcp_flow_dir}")

    return {
        "csv_flow": csv_flow_path,
        "csv_image": csv_img_path,
        "excel_studio": xlsx_path
    }

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1].endswith(".json"):
        with open(sys.argv[1], "r", encoding="utf-8") as f:
            data = json.load(f)
        export_thuytalk_project(data)
    else:
        print("Sử dụng: python export_thuytalk_storyboard.py <project_data.json>")
