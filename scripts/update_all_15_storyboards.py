# -*- coding: utf-8 -*-
"""
MASTER STORYBOARD GENERATOR & ENHANCER FOR THUYTALK CM3
Updates all 15 storyboard packages in reports/ThuyTalk_CM3:
1. CD03 Tap 1 (40 scenes)
2. CD03 Tap 2 (40 scenes)
3. CD06 (40 scenes)
4. CD07 (47 scenes)
5. CD08 (48 scenes)
6. CD09 (45 scenes)
7. CD10 (25 scenes)
8. CD11 (21 scenes)
9. CD12 (24 scenes)
10. CD13 (23 scenes)
11. CD14 (23 scenes)
12. CD15 (19 scenes)
13. CD16 (22 scenes)
14. CD17 (22 scenes)
15. CD18 (26 scenes)
Total scenes: 465 scenes.
Strict compliance:
- 0 camera words in prompt columns
- 100% unique, scene-specific narrative descriptions
- Explicit data-infused visualizations for numbers/metrics
- 100% consistent color palette: pure white background, crisp bold black outlines, vibrant yellow (#FFD700) and cyan blue (#00B4D8) accents, alert red for crises.
"""

import os
import sys
import re
import csv
import json
import shutil
import openpyxl

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.append(BASE_DIR)
sys.path.append(os.path.join(BASE_DIR, "scratch"))

from scripts.export_thuytalk_storyboard import export_thuytalk_project, safe_replace
import cm3_scenes_cd06_to_cd09 as mod_06_09
import cm3_scenes_cd10_to_cd18 as mod_10_18

REP_CM3 = os.path.join(BASE_DIR, "reports", "ThuyTalk_CM3")
REP_PLAN1 = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")
ONEDRIVE_YT = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\youtube"

with open(os.path.join(BASE_DIR, "scratch", "all_cm3_dialogues.json"), "r", encoding="utf-8") as f:
    ALL_DIALOGUES = json.load(f)

PROJECTS_CONFIG = [
    {
        "slug": "ThuyTalk_CM3-CD03_Tap_1_Thuong_diem_Phuong_Tay",
        "md_file": "CM3-CD03_Tap_1_Thuong_diem_Phuong_Tay.md",
        "type": "cd03_tap1",
        "title": "Thương Điểm Phương Tây & Hải Trình Đàng Ngoài (Tập 1)"
    },
    {
        "slug": "ThuyTalk_CM3-CD03_Tap_2_Kiem_dich_va_Thu_thue",
        "md_file": "CM3-CD03_Tap_2_Kiem_dich_va_Thu_thue.md",
        "type": "cd03_tap2",
        "title": "Kiểm Dịch & Thu Thuế Hải Quan Cổ Phố Hiến (Tập 2)"
    },
    {
        "slug": "ThuyTalk_CM3-CD06_Con_duong_To_lua_co_dai",
        "md_file": "CM3-CD06_Con_duong_To_lua_co_dai.md",
        "type": "scenes_list",
        "scenes_data": mod_06_09.CD06_SILK_ROAD_SCENES,
        "title": "Con Đường Tơ Lụa Cổ Đại & Cửa Ải Ngọc Môn Quan"
    },
    {
        "slug": "ThuyTalk_CM3-CD07_Chien_tranh_Thuoc_phien_1839_1842",
        "md_file": "CM3-CD07_Chien_tranh_Thuoc_phien_1839_1842.md",
        "type": "scenes_list",
        "scenes_data": mod_06_09.CD07_OPIUM_WAR_SCENES,
        "title": "Chiến Tranh Thuốc Phiện 1839 - 1842: Ngòi Nổ Từ Chén Trà"
    },
    {
        "slug": "ThuyTalk_CM3-CD08_Kenh_dao_Suez_1869_Vet_cat",
        "md_file": "CM3-CD08_Kenh_dao_Suez_1869_Vet_cat.md",
        "type": "scenes_list",
        "scenes_data": mod_06_09.CD08_SUEZ_1869_SCENES,
        "title": "Kênh Đào Suez 1869: Vết Cắt 193 Km Thay Đổi Hàng Hải"
    },
    {
        "slug": "ThuyTalk_CM3-CD09_Vu_trom_buc_hoa_Mona_Lisa",
        "md_file": "CM3-CD09_Vu_trom_buc_hoa_Mona_Lisa.md",
        "type": "scenes_list",
        "scenes_data": mod_06_09.CD09_MONA_LISA_SCENES,
        "title": "Vụ Trộm Bức Họa Mona Lisa 1911 & Nghệ Thuật Khoảng Trống"
    },
    {
        "slug": "ThuyTalk_CM3-CD10_Ky_An_Kenh_Dao_Panama_1914",
        "md_file": "CM3-CD10_Ky_An_Kenh_Dao_Panama_1914.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD10_PANAMA_SCENES,
        "title": "Kỳ Án Kênh Đào Panama 1914 & Khủng Hoảng Hạn Hán 2024"
    },
    {
        "slug": "ThuyTalk_CM3-CD11_Cuoc_Khung_Hoang_Kenh_Dao_Suez",
        "md_file": "CM3-CD11_Cuoc_Khung_Hoang_Kenh_Dao_Suez.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD11_EVER_GIVEN_SCENES,
        "title": "Cuộc Khủng Hoảng Siêu Tàu Ever Given Mắc Cạn Suez 2021"
    },
    {
        "slug": "ThuyTalk_CM3-CD12_Tuyen_Duong_Sat_Xuyen_Luc_Dia",
        "md_file": "CM3-CD12_Tuyen_Duong_Sat_Xuyen_Luc_Dia.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD12_RAILWAY_SCENES,
        "title": "Tuyến Đường Sắt Xuyên Lục Địa Á - Âu & Tàu Liên Vận Việt Nam"
    },
    {
        "slug": "ThuyTalk_CM3-CD13_De_Che_FedEx_Cu_Dat_Cuoc",
        "md_file": "CM3-CD13_De_Che_FedEx_Cu_Dat_Cuoc.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD13_FEDEX_SCENES,
        "title": "Đế Chế FedEx & Canh Bạc Las Vegas Cứu Nguy Năm 1973"
    },
    {
        "slug": "ThuyTalk_CM3-CD14_Thuong_Cang_Hoi_An_The_Ky",
        "md_file": "CM3-CD14_Thuong_Cang_Hoi_An_The_Ky.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD14_HOIAN_SCENES,
        "title": "Thương Cảng Hội An & Châu Ấn Thuyền Shuinsen Thế Kỷ 17"
    },
    {
        "slug": "ThuyTalk_CM3-CD15_Sieu_Cang_Tu_Dong_Hoa_Maasvlakte",
        "md_file": "CM3-CD15_Sieu_Cang_Tu_Dong_Hoa_Maasvlakte.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD15_MAASVLAKTE_SCENES,
        "title": "Siêu Cảng Tự Động Hóa Maasvlakte 2 Rotterdam Không Bóng Người"
    },
    {
        "slug": "ThuyTalk_CM3-CD16_Ky_An_Chuan_Hoa_Khoa_Twistlock",
        "md_file": "CM3-CD16_Ky_An_Chuan_Hoa_Khoa_Twistlock.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD16_TWISTLOCK_SCENES,
        "title": "Kỳ Án Chuẩn Hóa Khóa Góc Twistlock ISO Định Hình Thế Giới"
    },
    {
        "slug": "ThuyTalk_CM3-CD17_Ky_An_Tuyen_Duong_Tra_Ma",
        "md_file": "CM3-CD17_Ky_An_Tuyen_Duong_Tra_Ma.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD17_TRA_MA_SCENES,
        "title": "Kỳ Án Tuyến Đường Trà Mã Cổ Đạo & Danh Trà Phổ Nhĩ"
    },
    {
        "slug": "ThuyTalk_CM3-CD18_Cuoc_Dai_Di_Cu_Chuoi_Cung",
        "md_file": "CM3-CD18_Cuoc_Dai_Di_Cu_Chuoi_Cung.md",
        "type": "scenes_list",
        "scenes_data": mod_10_18.CD18_SUPPLY_CHAIN_SCENES,
        "title": "Cuộc Đại Di Cư 200 Năm Của Chuỗi Cung Ứng Dệt May Toàn Cầu"
    }
]

CAMERA_WORDS_REGEX = re.compile(r',\s*(?:Cận cảnh|Trung cảnh|Toàn cảnh)\s*\([^)]+\):?\s*|,\s*(?:Cận cảnh|Trung cảnh|Toàn cảnh)\s*\([^)]+\)\s*$|\b(?:Cận cảnh|Trung cảnh|Toàn cảnh)\s*\([^)]+\):?\s*', re.IGNORECASE)

def clean_cd03_flow_prompt(prompt):
    cleaned = CAMERA_WORDS_REGEX.sub(': ', prompt)
    cleaned = re.sub(r'\bClose-up of\b', 'Detailed macro view of', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\bMedium shot of\b', 'View of', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\bWide shot of\b', 'Panoramic view of', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s{2,}', ' ', cleaned)
    cleaned = re.sub(r':\s*:\s*', ': ', cleaned)
    return cleaned.strip()

def clean_cd03_image_prompt(prompt):
    cleaned = re.sub(r',\s*(?:Cận cảnh|Trung cảnh|Toàn cảnh)\s*\([^)]+\)\s*$', '', prompt, flags=re.IGNORECASE)
    cleaned = re.sub(r'\bClose-up of\b', 'Detailed view of', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\bMedium shot of\b', 'View of', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\bWide shot of\b', 'Panoramic view of', cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r'\s{2,}', ' ', cleaned)
    return cleaned.strip()

def build_project_package(cfg):
    slug = cfg["slug"]
    title = cfg["title"]
    p_type = cfg["type"]
    dialogues_list = ALL_DIALOGUES.get(slug, [])
    
    scenes = []
    image_prompts = []
    
    if p_type in ["cd03_tap1", "cd03_tap2"]:
        # Read existing flow CSV and clean camera words
        csv_path = os.path.join(REP_CM3, f"{slug}_Kich_Ban_Flow.csv")
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            reader = list(csv.reader(f))
        
        img_csv_path = os.path.join(REP_CM3, f"{slug}_Tao_Anh_Hang_Loat.csv")
        with open(img_csv_path, "r", encoding="utf-8-sig") as f:
            img_reader = list(csv.reader(f))
            
        for idx in range(1, len(reader)):
            r = reader[idx]
            stt = int(r[0])
            scene_title = r[1]
            raw_flow_prompt = r[2]
            shot_type = r[3]
            dialogue = r[4]
            duration = int(r[5]) if len(r) > 5 and r[5].isdigit() else 10
            
            cleaned_flow = clean_cd03_flow_prompt(raw_flow_prompt)
            
            scenes.append({
                "stt": stt,
                "title": scene_title,
                "prompt": cleaned_flow,
                "shot_type": shot_type,
                "dialogue": dialogue,
                "duration": duration
            })
            
            # Image prompt
            r_img = img_reader[idx] if idx < len(img_reader) else [stt, f"Frame {stt:02d}", "", shot_type, ""]
            raw_img_prompt = r_img[2]
            cleaned_img = clean_cd03_image_prompt(raw_img_prompt)
            
            image_prompts.append({
                "stt": stt,
                "title": r_img[1],
                "prompt": cleaned_img,
                "shot_type": shot_type,
                "note": r_img[4]
            })
    else:
        # scenes_list
        scenes_data = cfg["scenes_data"]
        for idx, s_def in enumerate(scenes_data, 1):
            stt = s_def["stt"]
            scene_title = f"Scene {stt:02d}"
            
            # Match dialogue and shot_type
            d_info = dialogues_list[idx-1] if idx-1 < len(dialogues_list) else {"dialogue": "", "shot_type": "Trung cảnh (Medium Shot)"}
            dialogue = d_info["dialogue"]
            shot_type = d_info["shot_type"]
            
            # Flow prompt
            flow_prompt = (
                f"{s_def['visual']}\n"
                f"Audio & SFX: {s_def['audio']}"
            )
            
            # Ensure no camera words in flow prompt
            flow_prompt = re.sub(r'\b(?:Cận cảnh|Trung cảnh|Toàn cảnh|Close-up|Medium Shot|Wide Shot)\b', '', flow_prompt)
            flow_prompt = re.sub(r'\s{2,}', ' ', flow_prompt)
            
            scenes.append({
                "stt": stt,
                "title": scene_title,
                "prompt": flow_prompt,
                "shot_type": shot_type,
                "dialogue": dialogue,
                "duration": 10
            })
            
            # Text-to-Image Prompt
            img_prompt = (
                f"Playful 2D minimalist doodle illustration: {s_def['visual']}, "
                f"clean vector line art, crisp bold black outlines, pure white background, "
                f"selective vibrant yellow and cyan accents, 8k resolution --ar 16:9"
            )
            # Remove any inadvertent camera words
            img_prompt = re.sub(r'\b(?:Cận cảnh|Trung cảnh|Toàn cảnh|Close-up|Medium Shot|Wide Shot)\b', '', img_prompt)
            img_prompt = re.sub(r'\s{2,}', ' ', img_prompt)
            
            image_prompts.append({
                "stt": stt,
                "title": s_def["frame_title"],
                "prompt": img_prompt,
                "shot_type": shot_type,
                "note": f"Start Frame S{stt:02d}: {s_def['note_vn']}"
            })

    project_data = {
        "project_name": slug,
        "title": title,
        "category": "CM3 - Ký Sự Lịch Sử Thương Mại & Logistics Quốc Tế",
        "visual_style": "Playful 2D minimalist doodle illustration, clean vector line art, crisp bold black outlines, pure white background, selective vibrant yellow and cyan accents, 8k resolution --ar 16:9",
        "character_anchor": "Cute 2D female stick figure named Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes",
        "voiceover_persona": "Thúy (Eureka Logistics) - giọng nói mộc mạc, tâm tình, tự nhiên, gần gũi",
        "scenes": scenes,
        "image_prompts": image_prompts
    }
    
    return project_data

def update_markdown_file(cfg, project_data):
    md_file = cfg.get("md_file")
    if not md_file:
        return
    md_path = os.path.join(REP_CM3, md_file)
    if not os.path.exists(md_path):
        return
        
    with open(md_path, "r", encoding="utf-8") as f:
        content = f.read()
        
    # We reconstruct the Flow table and Image table
    # 1. Flow Table rows
    flow_table_rows = []
    flow_table_rows.append("| STT | Tiêu đề | Mô tả phân cảnh 2D Doodle (Prompt) | Góc quay (Shot Type) | Lời thoại (Dialogue) | Thời lượng |")
    flow_table_rows.append("| :---: | :---: | :--- | :---: | :--- | :---: |")
    for s in project_data["scenes"]:
        p_clean = s["prompt"].replace("\n", " <br>").replace("|", "\\|")
        d_clean = s["dialogue"].replace("|", "\\|")
        flow_table_rows.append(f"| **{s['stt']}** | `{s['title']}` | {p_clean} | `{s['shot_type']}` | {d_clean} | {s['duration']}s |")
        
    flow_table_str = "\n".join(flow_table_rows)
    
    # 2. Image Table rows
    img_table_rows = []
    img_table_rows.append("| STT | Tiêu đề | Prompt tạo ảnh 2D Doodle (Image Prompt) | Góc quay (Shot Type) | Ghi chú |")
    img_table_rows.append("| :---: | :--- | :--- | :---: | :--- |")
    for img in project_data["image_prompts"]:
        p_clean = img["prompt"].replace("|", "\\|")
        t_clean = img["title"].replace("|", "\\|")
        n_clean = img["note"].replace("|", "\\|")
        img_table_rows.append(f"| **{img['stt']}** | `{t_clean}` | {p_clean} | `{img['shot_type']}` | {n_clean} |")
        
    img_table_str = "\n".join(img_table_rows)
    
    # Replace Table II (Flow)
    pattern_flow = r'(\| STT \| Tiêu đề \| Mô tả phân cảnh.*?)(?=\n---\n|\n## 🎨|\n## III|\Z)'
    if re.search(pattern_flow, content, flags=re.DOTALL):
        content = re.sub(pattern_flow, flow_table_str, content, flags=re.DOTALL)
    
    # Replace Table III (Image)
    pattern_img = r'(\| STT \| Tiêu đề \| Prompt tạo ảnh.*?)(?=\n---\n|\n## IV|\n## V|\Z)'
    if re.search(pattern_img, content, flags=re.DOTALL):
        content = re.sub(pattern_img, img_table_str, content, flags=re.DOTALL)
        
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(content)
        
    print(f"📄 Đã cập nhật Markdown: {md_path}")

def sync_to_targets(slug, md_file):
    # Sync to reports/plan1/video_scripts
    if os.path.exists(REP_PLAN1):
        for ext in ["_Kich_Ban_Flow.csv", "_Tao_Anh_Hang_Loat.csv", "_Studio.xlsx"]:
            src = os.path.join(REP_CM3, f"{slug}{ext}")
            dst = os.path.join(REP_PLAN1, f"{slug}{ext}")
            if os.path.exists(src):
                shutil.copy2(src, dst)
        if md_file:
            src_md = os.path.join(REP_CM3, md_file)
            dst_md = os.path.join(REP_PLAN1, md_file)
            if os.path.exists(src_md):
                shutil.copy2(src_md, dst_md)

    # Sync to OneDrive
    if os.path.exists(ONEDRIVE_YT):
        for ext in ["_Kich_Ban_Flow.csv", "_Tao_Anh_Hang_Loat.csv", "_Studio.xlsx"]:
            src = os.path.join(REP_CM3, f"{slug}{ext}")
            dst = os.path.join(ONEDRIVE_YT, f"{slug}{ext}")
            if os.path.exists(src):
                try:
                    shutil.copy2(src, dst)
                except Exception as e:
                    pass
        if md_file:
            src_md = os.path.join(REP_CM3, md_file)
            dst_md = os.path.join(ONEDRIVE_YT, md_file)
            if os.path.exists(src_md):
                try:
                    shutil.copy2(src_md, dst_md)
                except Exception as e:
                    pass

def main():
    print("🚀 BẮT ĐẦU CẬP NHẬT TOÀN DIỆN 15 DỰ ÁN STORYBOARD THUYTALK CM3...")
    for idx, cfg in enumerate(PROJECTS_CONFIG, 1):
        slug = cfg["slug"]
        print(f"\n[{idx}/15] Đang xử lý: {slug}...")
        project_data = build_project_package(cfg)
        
        # Export CSVs & Excel via export_thuytalk_project
        export_thuytalk_project(project_data, output_dir=REP_CM3)
        
        # Update Markdown
        update_markdown_file(cfg, project_data)
        
        # Sync to mirrors
        sync_to_targets(slug, cfg.get("md_file"))
        
    print("\n🎉 HOÀN TẤT CẬP NHẬT 15/15 DỰ ÁN THUYTALK CM3!")

if __name__ == "__main__":
    main()
