# -*- coding: utf-8 -*-
"""
Script: generate_youtube_5min_packaging.py
Mục tiêu:
1. Tạo kịch bản YouTube 5 phút (30 cảnh x 10s = 300s) chuẩn PO-2026-W42-64.
2. Đan xen chuẩn điện ảnh: 16 cảnh A-roll Talking Head (Chị Thúy chuyên gia độc lập) + 14 cảnh B-roll minh họa chuyên biệt (Cảng biển, Macro văn bản công văn, Sản phẩm thực tế, Phòng lab, Kho bãi).
3. Tuyệt đối không gắn tên bất kỳ doanh nghiệp nào (xây dựng thương hiệu cá nhân Chị Thúy - chuyên gia thực chiến độc lập).
4. Xuất bộ tệp chuẩn theo User Rule 13, Rule 14 và Rule 10 (Zero backup .xlsx):
   - File 1: reports/plan1/video_scripts/Bao_Bi_Thuc_Pham_5Phut_Kich_Ban_Flow.csv (5 cột)
   - File 2: reports/plan1/video_scripts/Bao_Bi_Thuc_Pham_5Phut_Tao_Anh_Hang_Loat.csv (5 cột, 31 ảnh)
   - File 3: reports/plan1/video_scripts/Bao_Bi_Thuc_Pham_5Phut_Studio.xlsx (4 sheets)
   - File 4: scratch/production_orders/KICH_BAN_YOUTUBE_5P_BAO_BI_THUC_PHAM.md
"""
import os
import sys
import csv
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
OUT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")
SCRATCH_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")

os.makedirs(OUT_DIR, exist_ok=True)
os.makedirs(SCRATCH_DIR, exist_ok=True)

# ==============================================================================
# DANH SÁCH 30 PHÂN CẢNH (300 GIÂY = 5 PHÚT CHUẨN)
# ĐAN XEN CHUẨN: 16 CẢNH A-ROLL TALKING HEAD + 14 CẢNH B-ROLL MINH HỌA
# MỖI CẢNH 10S: LỜI THOẠI 20 - 25 TỪ TIẾNG VIỆT
# ==============================================================================

SCENES = [
    # HỒI 1: HOOK & BÁO ĐỘNG KHẨN CẤP TẠI CẢNG (Cảnh 01 - 05)
    {
        "stt": 1,
        "title": "Scene 01",
        "shot_title": "[A-roll] Lời Chào & Nghịch Lý Ngành",
        "duration": 10,
        "dialogue": "Chào các bạn, mình là Thúy! Trong mảng xuất nhập khẩu bao bì thực phẩm, có một nghịch lý rất trớ trêu:",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy, elegant Vietnamese female trade & logistics expert in smart navy blazer, "
            "sits in modern sunlit executive office. She smiles warmly at camera with direct eye contact, "
            "then gestures thoughtfully with open hands. Camera slow gentle push-in. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chào các bạn, mình là Thúy! Trong mảng xuất nhập khẩu bao bì thực phẩm, có một nghịch lý rất trớ trêu:\" "
            "Audio & SFX: Warm inviting ambient corporate intro chime, subtle office atmosphere."
        ),
        "image_prompt": (
            "Cinematic talking head medium portrait of Ms. Thuy, Vietnamese female logistics expert in professional navy blazer, "
            "sitting at modern executive desk, warm engaging smile, direct eye contact with camera, soft background bokeh, natural studio lighting, 8k resolution --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy mở đầu video ấm áp, định danh chuyên gia độc lập và gợi mở nghịch lý"
    },
    {
        "stt": 2,
        "title": "Scene 02",
        "shot_title": "[B-roll] Cảnh Báo Kẹt Cảng & Thiệt Hại",
        "duration": 10,
        "dialogue": "Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần, phạt cả trăm triệu!",
        "prompt_en": (
            "Cinematic realistic B-roll: Busy Cat Lai container seaport terminal under harsh midday sun. "
            "Camera fast dolly push-in towards a 40ft refrigerated container with a heavy red customs padlock and detention seal. "
            "In background, giant gantry cranes idle, with an animated digital overlay of escalating demurrage penalty fees in glowing red numbers. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần, phạt cả trăm triệu!\" "
            "Audio & SFX: Distant seaport foghorn, heavy truck engine, sudden sharp warning alarm sound."
        ),
        "image_prompt": (
            "Cinematic wide documentary shot of seaport container terminal, rows of stacked cargo containers, "
            "a sharp red customs lock seal on container door, dramatic sun flare, gritty industrial realism, 8k --ar 16:9"
        ),
        "shot_type": "B-roll (Wide Shot)",
        "image_note": "B-roll cảnh báo kẹt cảng container thực tế, niêm phong kẹp chì đỏ"
    },
    {
        "stt": 3,
        "title": "Scene 03",
        "shot_title": "[B-roll] Cận Cảnh Túi Màng Nhôm Vô Trùng",
        "duration": 10,
        "dialogue": "Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!",
        "prompt_en": (
            "Macro product B-roll: Sterile multilayer aluminum aseptic packaging bags (Aseptic bag, HS code 3923.21.99) stacked inside clean industrial carton. "
            "Camera slow tilt-up from reflective metallic silver foil bag texture to an official corporate Self-Declaration dossier (Bản Tự công bố) with red stamp on desk. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!\" "
            "Audio & SFX: Metallic foil crinkling sound, subtle paper rustle."
        ),
        "image_prompt": (
            "Macro extreme close-up shot of sterile food packaging aluminum aseptic bags with blue sanitary spouts, "
            "clean commercial lighting, shallow depth of field, high-tech food manufacturing aesthetics, 8k --ar 16:9"
        ),
        "shot_type": "B-roll (Macro Shot)",
        "image_note": "B-roll cận cảnh sản phẩm túi màng nhôm vô trùng Aseptic bag"
    },
    {
        "stt": 4,
        "title": "Scene 04",
        "shot_title": "[B-roll] Bẫy Ngộ Nhận Thông Tư 27",
        "duration": 10,
        "dialogue": "Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!",
        "prompt_en": (
            "Document macro B-roll: Official Vietnamese legal document Circular 27/2026/TT-BYT open on executive conference table. "
            "A bright yellow highlighter marker emphasizes the regulatory clause: 'Nhóm có mức độ rủi ro trung bình - Thực hiện tự công bố sản phẩm'. "
            "Camera fast zoom on highlighted clause, followed by an overlay of a red stop sign. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!\" "
            "Audio & SFX: Marker squeak on paper, followed by harsh customs rejection stamp thud."
        ),
        "image_prompt": (
            "Macro top-down shot of official legal document Circular 27/2026/TT-BYT on wooden desk, yellow fluorescent highlighter emphasizing regulatory paragraph, "
            "reading glasses resting beside, warm desk lamp lighting, sharp focus --ar 16:9"
        ),
        "shot_type": "B-roll (Top-down Shot)",
        "image_note": "Chụp cận văn bản Thông tư 27/2026/TT-BYT điểm bẫy ngộ nhận pháp lý"
    },
    {
        "stt": 5,
        "title": "Scene 05",
        "shot_title": "[A-roll] Khẳng Định Bản Chất Văn Bản 22082",
        "duration": 10,
        "dialogue": "Công văn 22082 khẳng định: Tự công bố và Kiểm tra cửa khẩu là hai việc độc lập! Thúy sẽ gỡ rối trọn vẹn!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy seated in modern minimalist executive studio, looking directly into lens with authoritative reassuring confidence. "
            "She gestures decisively with two hands forming two separate pillars. Beside her, sleek digital floating cards display Official Dispatch 22082/CHQ-GSQL and Dispatch 1913/ATTP-KN with red seals. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Công văn 22082 khẳng định: Tự công bố và Kiểm tra cửa khẩu là hai việc độc lập! Thúy sẽ gỡ rối trọn vẹn!\" "
            "Audio & SFX: Clear affirmative bell chime, modern corporate synthesizer beat."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy in elegant navy blazer, smiling confidently with reassuring open-hand gesture, "
            "beside a floating digital glass card of Official Dispatch 22082 with red seal, modern executive studio backdrop --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy khẳng định văn bản 22082 và 1913, mở ra giải pháp toàn diện"
    },

    # HỒI 2: NGUYÊN LÝ 3 TẦNG QUẢN LÝ PHÁP LÝ (Cảnh 06 - 11)
    {
        "stt": 6,
        "title": "Scene 06",
        "shot_title": "[B-roll] Đồ Họa 3D Tháp 3 Tầng Pháp Lý",
        "duration": 10,
        "dialogue": "Để không bao giờ bị giữ hàng, bạn bắt buộc phải thấu suốt Nguyên lý 3 tầng quản lý pháp lý từ cẩm nang nghiệp vụ thực chiến!",
        "prompt_en": (
            "3D Motion Graphics B-roll: An elegant glassmorphism 3-tier pyramid floating against clean dark navy gradient background. "
            "Tier 1 (Base): 'Cửa khẩu Hải quan', Tier 2 (Middle): 'Thủ tục Tự công bố', Tier 3 (Apex): 'Quy chuẩn QCVN'. "
            "Camera smoothly orbits around the glowing pyramid structure as each layer pulses with neon light. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Để không bao giờ bị giữ hàng, bạn bắt buộc phải thấu suốt Nguyên lý 3 tầng quản lý pháp lý từ cẩm nang nghiệp vụ thực chiến!\" "
            "Audio & SFX: High-tech digital whoosh, holographic hum, gentle sweep."
        ),
        "image_prompt": (
            "Isometric 3D clean infographic visualization of a modern 3-tiered compliance pyramid structure, glowing neon layers, "
            "glassmorphism design, clean dark navy gradient background, sharp 8k render --ar 16:9"
        ),
        "shot_type": "B-roll (3D Animation)",
        "image_note": "Mô hình 3D trực quan Tháp 3 Tầng Quản Lý Pháp Lý"
    },
    {
        "stt": 7,
        "title": "Scene 07",
        "shot_title": "[A-roll] Giảng Giải Tầng 1 - Cửa Khẩu Hải Quan",
        "duration": 10,
        "dialogue": "Tầng 1 là Cửa khẩu Hải quan. Theo Thông tư 15 năm 2024, chỉ hai nhóm chất liệu bắt buộc kiểm tra tiền kiểm trước thông quan!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy seated at executive desk, looking intently into camera, holding a printed booklet of Circular 15/2024. "
            "She points downward indicating the foundational tier (Cửa khẩu), emphasizing the concept of 'Tiền kiểm' with clear precise articulation. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 1 là Cửa khẩu Hải quan. Theo Thông tư 15 năm 2024, chỉ hai nhóm chất liệu bắt buộc kiểm tra tiền kiểm trước thông quan!\" "
            "Audio & SFX: Subtle camera shutter sound, ambient office tone."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy seated at desk, holding a printed regulatory booklet of Circular 15/2024, "
            "looking directly at camera with clear informative demeanor, elegant navy blazer, 8k --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy phân tích Tầng 1 Cửa khẩu theo Thông tư 15"
    },
    {
        "stt": 8,
        "title": "Scene 08",
        "shot_title": "[B-roll] Cận Cảnh Nhựa PP & Núm Vú Cao Su",
        "duration": 10,
        "dialogue": "Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!",
        "prompt_en": (
            "Macro split B-roll: Left frame shows transparent plastic food container with prominent recycling resin identification code PP-05. "
            "Right frame shows baby silicone feeding bottle nipple labeled with HS code 4014.90.10. "
            "A red glowing warning border pulses around both items with bold text: 'BẮT BUỘC KIỂM TRA ATTP CỬA KHẨU'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!\" "
            "Audio & SFX: Dual alert chime, electronic barcode scanner beep."
        ),
        "image_prompt": (
            "Macro studio shot showing transparent PP food storage container beside baby silicone bottle nipple, "
            "clean white product studio background, sharp macro details, laboratory lighting, no text --ar 16:9"
        ),
        "shot_type": "B-roll (Macro Shot)",
        "image_note": "Cận cảnh nhựa tổng hợp PP/PE và núm vú cao su trẻ em"
    },
    {
        "stt": 9,
        "title": "Scene 09",
        "shot_title": "[B-roll] Danh Mục Sản Phẩm Miễn Cửa Khẩu",
        "duration": 10,
        "dialogue": "Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!",
        "prompt_en": (
            "Product showcase B-roll: Premium stainless steel 304 cooking pot, fine white ceramic plate, clear glass beverage bottle, and colorful silicone baking mold. "
            "Camera smooth cinematic panning across items on marble countertop. Bright green checkmark badges illuminate sequentially above each item with text: 'MIỄN KIỂM TRA CỬA KHẨU'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!\" "
            "Audio & SFX: Soft metallic chime, 4 crisp positive success pings."
        ),
        "image_prompt": (
            "Clean studio product showcase of stainless steel pot, white ceramic plate, glass jar and blue silicone baking mold, "
            "arranged neatly on marble countertop, bright softbox illumination, premium commercial look --ar 16:9"
        ),
        "shot_type": "B-roll (Medium Shot)",
        "image_note": "B-roll sản phẩm inox, sứ, thủy tinh được miễn kiểm tra cửa khẩu"
    },
    {
        "stt": 10,
        "title": "Scene 10",
        "shot_title": "[A-roll] Giảng Giải Tầng 2 - Thủ Tục Tự Công Bố",
        "duration": 10,
        "dialogue": "Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy seated in front of her laptop on executive desk, explaining the second administrative layer. "
            "She gestures with open palm towards the camera, maintaining poised professional posture, emphasizing legal circulation on the market. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!\" "
            "Audio & SFX: Mouse click sound, subtle keyboard tap, warm synthesizer pad."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy seated in front of laptop in modern office, explaining regulatory compliance with poised professional posture, "
            "warm natural lighting, executive background --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy hướng dẫn Tầng 2 Tự công bố lưu hành thị trường"
    },
    {
        "stt": 11,
        "title": "Scene 11",
        "shot_title": "[B-roll] Phòng Lab Thử Nghiệm Chuẩn QCVN",
        "duration": 10,
        "dialogue": "Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!",
        "prompt_en": (
            "Laboratory testing B-roll: Scientist in pristine white lab coat and blue gloves placing plastic food container sample into spectrophotometer analytical equipment. "
            "Computer monitor in background displays passing chemical migration curve labeled 'QCVN 12-1:2011/BYT - TEST PASSED'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!\" "
            "Audio & SFX: Clean laboratory machine beep, centrifuge hum, deep confident bass swell."
        ),
        "image_prompt": (
            "Cinematic documentary shot inside certified food contact safety testing laboratory, female analytical chemist holding test tube, "
            "scientific equipment and analytical balances in background, pristine sterile environment --ar 16:9"
        ),
        "shot_type": "B-roll (Medium Close-up)",
        "image_note": "B-roll phòng lab thử nghiệm chỉ tiêu thôi nhiễm QCVN"
    },

    # HỒI 3: PHÂN LOẠI CHẤT LIỆU & BẪY SILICONE NGƯỜI LỚN (Cảnh 12 - 16)
    {
        "stt": 12,
        "title": "Scene 12",
        "shot_title": "[A-roll] Cạm Bẫy Cao Su Trẻ Em vs Silicone",
        "duration": 10,
        "dialogue": "Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy leans slightly forward into camera lens, lowering her voice conspiratorially to share an insider secret. "
            "Her eyes are intense, focused and sharp. Animated red badge flashes beside her: 'ĐIỂM NGHẼN 80% DOANH NGHIỆP MẮC BẪY'. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!\" "
            "Audio & SFX: Dramatic chime, suspense whoosh."
        ),
        "image_prompt": (
            "Cinematic close-up portrait of Ms. Thuy, leaning forward with engaging, intense storytelling expression, "
            "dark executive boardroom background, soft dramatic lighting, 8k --ar 16:9"
        ),
        "shot_type": "A-roll (Close-up)",
        "image_note": "Chị Thúy hạ giọng chỉ ra điểm thắt cổ chai 80% mắc bẫy"
    },
    {
        "stt": 13,
        "title": "Scene 13",
        "shot_title": "[B-roll] Bóc Tách Nghịch Lý Khuôn Silicone Người Lớn",
        "duration": 10,
        "dialogue": "Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!",
        "prompt_en": (
            "Split document B-roll: Left frame displays Circular 15/2024 table without silicone mold HS code (marked 'Miễn cửa khẩu'). "
            "Right frame displays Circular 27/2026 Item 6: 'Bao bì dụng cụ cao su chứa đựng' (marked 'Bắt buộc Tự công bố'). "
            "A flexible orange silicone cake mold rests at the intersection of both documents. Animated arrows highlight the regulatory difference. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!\" "
            "Audio & SFX: Soft rubber flex sound, paper rustle, electronic highlight buzz."
        ),
        "image_prompt": (
            "Split graphic comparison on executive wooden table: official Circular 15 table contrasted with Circular 27 Item 6, "
            "colorful silicone cake mold resting on document corner, high resolution legal audit scene --ar 16:9"
        ),
        "shot_type": "B-roll (Top-down Shot)",
        "image_note": "Bóc tách nghịch lý silicone người lớn miễn cửa khẩu nhưng phải TCB"
    },
    {
        "stt": 14,
        "title": "Scene 14",
        "shot_title": "[B-roll] Đồ Gia Dụng Inox & Gốm Sứ Nhóm 1",
        "duration": 10,
        "dialogue": "Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!",
        "prompt_en": (
            "Product macro B-roll: Panning shot of gleaming stainless steel saucepan, fine porcelain soup bowl, and crystalline glass tumbler. "
            "A green official badge stamps over items with text: 'NHÓM 1: RỦI RO THẤP - KHÔNG PHẢI NỘP HỒ SƠ LÊN SỞ'. "
            "Sunlight reflects brightly off polished metal and glass surfaces. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!\" "
            "Audio & SFX: Crystal glass clink, metallic ring, positive confirmation tone."
        ),
        "image_prompt": (
            "Commercial macro product photography of polished stainless steel pan, elegant white ceramic bowl and clear glass tumbler, "
            "sunlit luxury kitchen studio set, pristine reflections, 8k resolution --ar 16:9"
        ),
        "shot_type": "B-roll (Close-up)",
        "image_note": "B-roll nồi inox, bát sứ, ly thủy tinh thuộc nhóm rủi ro thấp"
    },
    {
        "stt": 15,
        "title": "Scene 15",
        "shot_title": "[A-roll] Hướng Dẫn Hồ Sơ TCCS Phòng Thủ Hậu Kiểm",
        "duration": 10,
        "dialogue": "Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy holding an organized black corporate filing binder labeled 'HỒ SƠ TCCS NỘI BỘ'. "
            "She turns a page showing verified stamped lab report, smiling reassuringly at the camera, explaining internal compliance defense. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!\" "
            "Audio & SFX: Binder ring snap click, secure lock latch sound."
        ),
        "image_prompt": (
            "Medium close-up of Ms. Thuy holding an organized black corporate filing binder with official labels, "
            "confident advisory demeanor, modern office setting --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Close-up)",
        "image_note": "Chị Thúy hướng dẫn lưu trữ hồ sơ Tiêu chuẩn cơ sở nội bộ"
    },
    {
        "stt": 16,
        "title": "Scene 16",
        "shot_title": "[A-roll] Tối Ưu Tiết Kiệm Chi Phí Hàng Triệu Đồng",
        "duration": 10,
        "dialogue": "Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy smiles with pride and warmth, gesturing towards a digital financial savings graphic popping up on screen: "
            "'Tiết kiệm 1.000.000 đ/model Inox & Thủy tinh'. Camera slow push-in on her trustworthy expression. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!\" "
            "Audio & SFX: Cheerful register cha-ching sound, upbeat synth chime."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy with an encouraging smile, hands gesturing savings, "
            "modern executive office background with blurred trade books on shelf --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy nhấn mạnh giá trị tiết kiệm tiền bạc cho doanh nghiệp"
    },

    # HỒI 4: ĐIỂM HUYỆT - CHỨA ĐỰNG VS TIẾP XÚC (Cảnh 17 - 21)
    {
        "stt": 17,
        "title": "Scene 17",
        "shot_title": "[B-roll] So Sánh Dụng Cụ Chứa Đựng vs Tiếp Xúc",
        "duration": 10,
        "dialogue": "Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!",
        "prompt_en": (
            "Studio high-contrast product comparison setup on dark slate table: "
            "Left side: Plastic food storage box with sealed airtight locking lid (chứa đựng). "
            "Right side: Stainless steel dinner fork, spoon, wooden chopsticks and cooking spatula (tiếp xúc không chứa đựng). "
            "Camera slow tracking dolly shot between the two groups with dramatic side rim lighting. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!\" "
            "Audio & SFX: Smooth camera dolly sound, subtle low-frequency suspense beat."
        ),
        "image_prompt": (
            "Studio product comparison on dark slate surface: food storage box with sealed lid on left, metal spoon fork chopsticks on right, "
            "cinematic edge lighting, professional commercial aesthetic, clean and crisp --ar 16:9"
        ),
        "shot_type": "B-roll (Medium Shot)",
        "image_note": "So sánh dụng cụ chứa đựng vs dụng cụ tiếp xúc không chứa đựng"
    },
    {
        "stt": 18,
        "title": "Scene 18",
        "shot_title": "[B-roll] Kim Bài Công Văn 4372 Bộ Y Tế (Aeon)",
        "duration": 10,
        "dialogue": "Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!",
        "prompt_en": (
            "Macro document close-up: Official letterhead of Ministry of Health (Bộ Y tế) - Dispatch 4372/BYT-VPB6 dated 31/07/2019 addressing Aeon Vietnam Co., Ltd. "
            "Warm dramatic spotlight illuminates official ministry header, red national emblem stamp and incoming registration number. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!\" "
            "Audio & SFX: Paper rustle, spotlight switch on sound."
        ),
        "image_prompt": (
            "Document close-up of official Vietnamese Ministry of Health dispatch 4372/BYT-VPB6, authentic red stamp emblem, "
            "spotlight illumination on official seal, high legal credibility aesthetic --ar 16:9"
        ),
        "shot_type": "B-roll (Close-up)",
        "image_note": "Chụp cận Công văn 4372/BYT-VPB6 của Bộ Y tế trả lời Aeon"
    },
    {
        "stt": 19,
        "title": "Scene 19",
        "shot_title": "[B-roll] Thìa Đũa Dụng Cụ Thao Tác Cơ Học",
        "duration": 10,
        "dialogue": "Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!",
        "prompt_en": (
            "Fast dynamic flat-lay B-roll: Macro cutaways of stainless steel soup spoon, wooden chopsticks, chef fruit peeling knife, metal food tongs, and blue kitchen silicone gloves. "
            "Text overlay in bold green: 'DỤNG CỤ THAO TÁC CƠ HỌC - KHÔNG CÓ LÒNG CHỨA ĐỰNG BẢO QUẢN'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!\" "
            "Audio & SFX: Fast rhythm percussion, stainless cutlery clattering sound."
        ),
        "image_prompt": (
            "Collage flat lay of kitchen utensils: metal spoons, wooden chopsticks, fruit peeling knife, food tongs and kitchen gloves, "
            "neatly organized on culinary prep board, bright modern culinary lighting --ar 16:9"
        ),
        "shot_type": "B-roll (Top-down Shot)",
        "image_note": "B-roll các dụng cụ không chứa đựng: đũa, thìa, dĩa, dao gọt, kẹp gắp"
    },
    {
        "stt": 20,
        "title": "Scene 20",
        "shot_title": "[A-roll] Khẳng Định Miễn TCB & Miễn Kiểm Tra Cửa Khẩu",
        "duration": 10,
        "dialogue": "Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy makes two decisive outward hand gestures, nodding firmly with sharp authoritative confidence. "
            "Floating green approval stamp graphic appears beside her: 'MIỄN TỰ CÔNG BỐ & MIỄN KIỂM TRA CỬA KHẨU'. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!\" "
            "Audio & SFX: Double stamp thud sound, green channel notification chime."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy gesturing decisively with open hands, confident expert posture, "
            "green channel approval badge floating on screen, modern studio background, 8k portrait --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy khẳng định miễn TCB và miễn kiểm tra cửa khẩu"
    },
    {
        "stt": 21,
        "title": "Scene 21",
        "shot_title": "[B-roll] Xe Container Thông Quan Luồng Xanh Ra Cảng",
        "duration": 10,
        "dialogue": "Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!",
        "prompt_en": (
            "Seaport logistics B-roll: Heavy container reach stacker smoothly lowering a 40ft shipping container onto a modern delivery truck. "
            "Truck accelerates smoothly and drives out through the seaport terminal barrier gate under clear bright blue sky. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!\" "
            "Audio & SFX: Container reach stacker hydraulic hiss, truck engine accelerating, uplifting corporate harmony."
        ),
        "image_prompt": (
            "Dynamic action shot of reach stacker vehicle lifting shipping container onto modern delivery truck at seaport terminal, "
            "blue sunny sky, pristine logistics operation, motion blur on truck wheels, 8k --ar 16:9"
        ),
        "shot_type": "B-roll (Wide Shot)",
        "image_note": "B-roll xe container lấy hàng thông quan nhanh chóng ra khỏi cảng"
    },

    # HỒI 5: CHIẾN LƯỢC HÀNH ĐỘNG THỰC CHIẾN (Cảnh 22 - 26)
    {
        "stt": 22,
        "title": "Scene 22",
        "shot_title": "[B-roll] Cảnh Báo Mốc Thời Gian Lab 10-12 Ngày",
        "duration": 10,
        "dialogue": "Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!",
        "prompt_en": (
            "Laboratory timeline B-roll: Analytical test tube rack with colorful chemical extraction reagents. "
            "A digital calendar animation overlay flips from Day 1 through Day 12 consecutively with a rapidly ticking analog clock graphic. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!\" "
            "Audio & SFX: Rapid clock ticking SFX, electronic calendar flip sound."
        ),
        "image_prompt": (
            "Close-up of laboratory analytical test tube rack with colorful chemical solutions, digital countdown timer on wall, "
            "depth of field on scientist reviewing lab notebook, crisp modern science laboratory --ar 16:9"
        ),
        "shot_type": "B-roll (Close-up)",
        "image_note": "Cảnh báo mốc thời gian thử nghiệm phòng lab 10-12 ngày"
    },
    {
        "stt": 23,
        "title": "Scene 23",
        "shot_title": "[B-roll] Ác Mộng Hóa Đơn Phạt Dem/Det 15-30 Triệu",
        "duration": 10,
        "dialogue": "Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!",
        "prompt_en": (
            "Dramatic financial B-roll: Macro close-up of a shipping line demurrage/detention invoice stamped with overdue penalty: 'DEMURRAGE / DETENTION OVERDUE: 30.000.000 VND'. "
            "In background, a long line of idle containers sitting under dark overcast clouds at port. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!\" "
            "Audio & SFX: Heavy financial penalty buzzer, ominous low synth drone."
        ),
        "image_prompt": (
            "Dramatic close-up of demurrage fee penalty document held in hand against blurry port container background, "
            "overcast dramatic sky, realistic logistics crisis documentary feel --ar 16:9"
        ),
        "shot_type": "B-roll (Medium Close-up)",
        "image_note": "Cảnh báo tiền phạt lưu container lưu bãi Dem/Det 15-30 triệu"
    },
    {
        "stt": 24,
        "title": "Scene 24",
        "shot_title": "[A-roll] Tuyệt Chiêu Gửi Mẫu CPN Trước 25 Ngày",
        "duration": 10,
        "dialogue": "Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy smiles brightly, holding up an international air express envelope (CPN), revealing her practical insider secret. "
            "She speaks with warm mentor encouragement, pointing out the 25-day advance window. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!\" "
            "Audio & SFX: Parcel opening sound, upbeat brass chime."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy holding an air express document envelope, "
            "smiling enthusiastically, sharing practical logistics trade secret, bright executive office --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy cầm phong bì CPN chia sẻ tuyệt chiêu gửi mẫu trước 25 ngày"
    },
    {
        "stt": 25,
        "title": "Scene 25",
        "shot_title": "[B-roll] Màn Hình Cổng Một Cửa NSW Đạt Chuẩn",
        "duration": 10,
        "dialogue": "Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!",
        "prompt_en": (
            "Screen capture B-roll: Vietnam National Single Window (Cổng thông tin một cửa quốc gia NSW) interface screen. "
            "Status column displays: 'ĐÃ TIẾP NHẬN HỒ SƠ - KẾT QUẢ: ĐẠT YÊU CẦU NHẬP KHẨU (MẪU 05)'. "
            "Customs electronic signature stamp illuminates with a bright green verification badge. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!\" "
            "Audio & SFX: Electronic stamp sound, digital success fanfare."
        ),
        "image_prompt": (
            "Computer screen displaying Vietnam National Single Window clearance certificate with green approved checkmark, "
            "customs declaration specialist blurred in background, modern high-tech office --ar 16:9"
        ),
        "shot_type": "B-roll (Close-up)",
        "image_note": "Màn hình Một cửa quốc gia NSW cấp thông báo đạt trong 3 ngày"
    },
    {
        "stt": 26,
        "title": "Scene 26",
        "shot_title": "[B-roll] Xe Cont Cập Kho Riêng Cắt Phí Lưu Bãi",
        "duration": 10,
        "dialogue": "Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!",
        "prompt_en": (
            "Logistics warehouse B-roll: Heavy container delivery truck smoothly backing into a spacious, clean enterprise logistics warehouse loading dock. "
            "Automated warehouse roll-up door lifts open. Warehouse manager inspects intact customs seal and signs delivery note. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!\" "
            "Audio & SFX: Truck backing-up beep, warehouse automated door rolling up, air brake hiss."
        ),
        "image_prompt": (
            "Wide shot of modern clean distribution warehouse interior, container truck docking at loading bay, "
            "bright industrial LED lighting, professional supply chain operation --ar 16:9"
        ),
        "shot_type": "B-roll (Wide Shot)",
        "image_note": "Đưa container về kho riêng bảo quản cắt đứt tiền lưu kho bãi"
    },

    # HỒI 6: ĐÚC KẾT & KÊU GỌI HÀNH ĐỘNG (Cảnh 27 - 30)
    {
        "stt": 27,
        "title": "Scene 27",
        "shot_title": "[A-roll] Bản Tự Công Bố Có Giá Trị Vô Thời Hạn",
        "duration": 10,
        "dialogue": "Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy spreads hands wide, highlighting the immense long-term value. "
            "Floating golden graphic card appears beside her: 'HIỆU LỰC VÔ THỜI HẠN - LÀM 1 LẦN DÙNG MÃI MÃI'. "
            "She smiles encouragingly, radiating business wisdom. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!\" "
            "Audio & SFX: Golden chime sound, triumphant synthesizer chord."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy with hands open in a welcoming, expansive gesture, "
            "genuine smile, elegant executive office backdrop --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy nhấn mạnh bản Tự công bố có hiệu lực vô thời hạn"
    },
    {
        "stt": 28,
        "title": "Scene 28",
        "shot_title": "[B-roll] Lưu Ý Phiếu Kiểm Nghiệm Hạn 12 Tháng",
        "duration": 10,
        "dialogue": "Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!",
        "prompt_en": (
            "Document B-roll: Official laboratory testing certificate with certified ISO/IEC 17025 stamp and red official seal. "
            "A calendar bracket highlights the 12-month validity window. An elegant glass hourglass rests beside the document with sand trickling down gently. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!\" "
            "Audio & SFX: Soft clock ticking, gentle advisory reminder tone, woodblock chime."
        ),
        "image_prompt": (
            "Close-up of laboratory analytical certificate with blue accreditation stamp and red date stamp, "
            "resting on wooden executive desk with hourglass beside, warm lighting --ar 16:9"
        ),
        "shot_type": "B-roll (Close-up)",
        "image_note": "Lưu ý phiếu kiểm nghiệm có giá trị 12 tháng khi nộp hồ sơ"
    },
    {
        "stt": 29,
        "title": "Scene 29",
        "shot_title": "[A-roll] Đúc Kết Thông Điệp Triết Lý Thực Chiến",
        "duration": 10,
        "dialogue": "Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy looks straight into the camera lens with intense, inspiring conviction. "
            "Her expression embodies resilience and visionary leadership. Camera very slow push-in. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!\" "
            "Audio & SFX: Deep inspiring piano chords swell gracefully."
        ),
        "image_prompt": (
            "Cinematic close-up portrait of Ms. Thuy looking directly into lens, "
            "powerful leadership presence, visionary inspiring eyes, elegant corporate aesthetics, 8k --ar 16:9"
        ),
        "shot_type": "A-roll (Close-up)",
        "image_note": "Chị Thúy đúc kết thông điệp triết lý logistics thực chiến"
    },
    {
        "stt": 30,
        "title": "Scene 30",
        "shot_title": "[A-roll] Outro Kêu Gọi Subscribe & Tặng Cheat-Sheet",
        "duration": 10,
        "dialogue": "Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!",
        "prompt_en": (
            "A-roll Talking Head: Ms. Thuy waves hand gently with radiant warm smile, gesturing downwards to comment section. "
            "YouTube interactive graphics emerge beside her: Red Subscribe button with bell notification, "
            "and floating 3D digital book mockup labeled 'CẨM NANG NGHIỆP VỤ NHẬP KHẨU BAO BÌ THỰC PHẨM'. "
            "Voiceover (Vietnamese, confident female expert, pacing 10s): \"Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!\" "
            "Audio & SFX: Cheerful YouTube outro chime, bell notification ding, mouse click sound."
        ),
        "image_prompt": (
            "Medium shot of Ms. Thuy smiling radiantly, waving hand gracefully towards viewer, "
            "warm sunlit studio set, YouTube outro branding finish --ar 16:9"
        ),
        "shot_type": "A-roll (Medium Shot)",
        "image_note": "Chị Thúy tươi cười vẫy tay chào tạm biệt và tặng Cẩm nang"
    }
]

print(f"Tổng số phân cảnh: {len(SCENES)} cảnh.")
total_duration = sum(s["duration"] for s in SCENES)
total_words = sum(len(s["dialogue"].split()) for s in SCENES)
print(f"Tổng thời lượng: {total_duration} giây ({total_duration/60:.1f} phút).")
print(f"Tổng số từ lời thoại: {total_words} từ (trung bình {total_words/len(SCENES):.1f} từ/cảnh).")

aroll_count = sum(1 for s in SCENES if "[A-roll]" in s["shot_title"])
broll_count = sum(1 for s in SCENES if "[B-roll]" in s["shot_title"])
print(f"Số cảnh A-roll (Talking Head): {aroll_count} cảnh ({aroll_count/len(SCENES)*100:.1f}%).")
print(f"Số cảnh B-roll (Minh họa chuyên biệt): {broll_count} cảnh ({broll_count/len(SCENES)*100:.1f}%).")

# ==============================================================================
# DANH SÁCH 31 ẢNH TẠO HÀNG LOẠT (30 CẢNH + 1 YOUTUBE MASTER THUMBNAIL)
# ==============================================================================
IMAGES_31 = []
for s in SCENES:
    IMAGES_31.append({
        "stt": s["stt"],
        "title": f"Scene {s['stt']:02d} - {s['shot_title']}",
        "prompt": s["image_prompt"],
        "shot_type": s["shot_type"],
        "note": s["image_note"]
    })

# Ảnh thứ 31: Master YouTube Thumbnail / Keyframe Visual (Ảnh Bìa YouTube 16:9 thu hút CTR cao)
IMAGES_31.append({
    "stt": 31,
    "title": "Thumbnail - Ảnh Bìa YouTube Master 16:9",
    "prompt": (
        "Cinematic YouTube Master Thumbnail, ultra-high-resolution 16:9, dramatic commercial key visual. "
        "Left side: Professional Vietnamese female trade and logistics expert Ms. Thuy (mid-30s, sharp, confident, "
        "visionary expression, tailored navy blue blazer, arms folded confidently, direct engaging eye contact). "
        "Right side background: Dramatic atmospheric seaport of Cat Lai at twilight, towering gantry cranes, "
        "stacks of shipping containers under moody twilight sky, high-contrast rim lighting, and a floating bold red 3D stamp "
        "labeled 'CÔNG VĂN 22082: CẢNH BÁO RỦI RO HẢI QUAN'. Volumetric lighting, sharp 8k focus, commercial poster grade, "
        "clean negative space on top-right for high-CTR YouTube typography overlay --ar 16:9"
    ),
    "shot_type": "Key Visual / Master Thumbnail",
    "note": "Ảnh Bìa YouTube Master 16:9 (Key visual sắc nét, thu hút CTR cao, định vị Chị Thúy & cảnh báo rủi ro hải quan)"
})
print(f"Tổng số ảnh tạo hàng loạt: {len(IMAGES_31)} ảnh (30 Cảnh + 1 Master Thumbnail).")

# ==============================================================================
# XUẤT FILE 1: CSV KỊCH BẢN GOOGLE FLOW (5 CỘT CHUẨN - 30 CẢNH)
# ==============================================================================
csv_flow_path = os.path.join(OUT_DIR, "Bao_Bi_Thuc_Pham_5Phut_Kich_Ban_Flow.csv")
tmp_flow = csv_flow_path + ".tmp.csv"
with open(tmp_flow, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["STT", "Tiêu đề", "Mô tả phân cảnh (Prompt)", "Lời thoại (Dialogue)", "Thời lượng (giây)"])
    for s in SCENES:
        writer.writerow([
            s["stt"],
            s["title"],
            s["prompt_en"],
            s["dialogue"],
            s["duration"]
        ])
os.replace(tmp_flow, csv_flow_path)
print(f"SUCCESS: Đã xuất file CSV Flow -> {csv_flow_path}")

# ==============================================================================
# XUẤT FILE 2: CSV TẠO ẢNH HÀNG LOẠT (5 CỘT CHUẨN - 31 ẢNH)
# ==============================================================================
csv_image_path = os.path.join(OUT_DIR, "Bao_Bi_Thuc_Pham_5Phut_Tao_Anh_Hang_Loat.csv")
tmp_image = csv_image_path + ".tmp.csv"
with open(tmp_image, "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"])
    for img in IMAGES_31:
        writer.writerow([
            img["stt"],
            img["title"],
            img["prompt"],
            img["shot_type"],
            img["note"]
        ])
os.replace(tmp_image, csv_image_path)
print(f"SUCCESS: Đã xuất file CSV Tạo Ảnh (31 ảnh) -> {csv_image_path}")

# ==============================================================================
# XUẤT FILE 3: EXCEL STUDIO 4 SHEET THEO CHUẨN KỸ NĂNG AI-VIDEO-PRODUCER-PRO
# (TUÂN THỦ RULE 10: ATOMIC SAVE, KHÔNG TẠO BACKUP)
# ==============================================================================
excel_path = os.path.join(OUT_DIR, "Bao_Bi_Thuc_Pham_5Phut_Studio.xlsx")
wb = openpyxl.Workbook()

font_name = "Segoe UI"
header_font = Font(name=font_name, size=11, bold=True, color="FFFFFF")
cell_font = Font(name=font_name, size=10)
bold_cell_font = Font(name=font_name, size=10, bold=True)
title_font = Font(name=font_name, size=14, bold=True, color="1B365D")
sub_font = Font(name=font_name, size=10, italic=True, color="555555")

header_fill_blue = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
header_fill_teal = PatternFill(start_color="008080", end_color="008080", fill_type="solid")
header_fill_purple = PatternFill(start_color="4A235A", end_color="4A235A", fill_type="solid")
header_fill_dark = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")

zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="E2E8F0"),
    right=Side(style="thin", color="E2E8F0"),
    top=Side(style="thin", color="E2E8F0"),
    bottom=Side(style="thin", color="E2E8F0")
)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ------------------------------------------------------------------------------
# SHEET 1: KICH_BAN_FLOW (5 CỘT CHUẨN GOOGLE FLOW)
# ------------------------------------------------------------------------------
ws1 = wb.active
ws1.title = "KICH_BAN_FLOW"
ws1.views.sheetView[0].showGridLines = True

headers1 = ["STT", "Tiêu đề", "Mô tả phân cảnh (Prompt)", "Lời thoại (Dialogue)", "Thời lượng (giây)"]
for col_idx, h in enumerate(headers1, 1):
    c = ws1.cell(row=1, column=col_idx, value=h)
    c.font = header_font
    c.fill = header_fill_blue
    c.alignment = align_center
    c.border = thin_border
ws1.row_dimensions[1].height = 28

for row_idx, s in enumerate(SCENES, 2):
    fill = zebra_fill if row_idx % 2 == 0 else white_fill
    vals = [s["stt"], s["title"], s["prompt_en"], s["dialogue"], s["duration"]]
    for col_idx, val in enumerate(vals, 1):
        c = ws1.cell(row=row_idx, column=col_idx, value=val)
        c.font = cell_font
        c.border = thin_border
        c.fill = fill
        if col_idx in [1, 2, 5]:
            c.alignment = align_center
            if col_idx in [1, 2]:
                c.font = bold_cell_font
        else:
            c.alignment = align_left
    ws1.row_dimensions[row_idx].height = 70

ws1.column_dimensions["A"].width = 8
ws1.column_dimensions["B"].width = 14
ws1.column_dimensions["C"].width = 55
ws1.column_dimensions["D"].width = 45
ws1.column_dimensions["E"].width = 16

# ------------------------------------------------------------------------------
# SHEET 2: TAO_ANH_HANG_LOAT (5 CỘT CHUẨN KEYFRAMES & B-ROLL)
# ------------------------------------------------------------------------------
ws2 = wb.create_sheet(title="TAO_ANH_HANG_LOAT")
ws2.views.sheetView[0].showGridLines = True

headers2 = ["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"]
for col_idx, h in enumerate(headers2, 1):
    c = ws2.cell(row=1, column=col_idx, value=h)
    c.font = header_font
    c.fill = header_fill_teal
    c.alignment = align_center
    c.border = thin_border
ws2.row_dimensions[1].height = 28

for row_idx, img in enumerate(IMAGES_31, 2):
    fill = zebra_fill if row_idx % 2 == 0 else white_fill
    vals = [img["stt"], img["title"], img["prompt"], img["shot_type"], img["note"]]
    for col_idx, val in enumerate(vals, 1):
        c = ws2.cell(row=row_idx, column=col_idx, value=val)
        c.font = cell_font
        c.border = thin_border
        c.fill = fill
        if col_idx in [1, 4]:
            c.alignment = align_center
            if col_idx == 4:
                c.font = bold_cell_font
        elif col_idx == 2:
            c.alignment = align_left
            c.font = bold_cell_font
        else:
            c.alignment = align_left
    ws2.row_dimensions[row_idx].height = 65

ws2.column_dimensions["A"].width = 8
ws2.column_dimensions["B"].width = 30
ws2.column_dimensions["C"].width = 60
ws2.column_dimensions["D"].width = 22
ws2.column_dimensions["E"].width = 35

# ------------------------------------------------------------------------------
# SHEET 3: KICH_BAN_VOICEOVER_FULL (LỜI THOẠI LIÊN TỤC CÓ MỐC THỜI GIAN)
# ------------------------------------------------------------------------------
ws3 = wb.create_sheet(title="KICH_BAN_VOICEOVER_FULL")
ws3.views.sheetView[0].showGridLines = True

headers3 = ["STT", "Mốc thời gian", "Phân đoạn kịch bản", "Lời thoại Voiceover (Đo chuẩn từ)", "Ghi chú cảm xúc & Nhịp thở"]
for col_idx, h in enumerate(headers3, 1):
    c = ws3.cell(row=1, column=col_idx, value=h)
    c.font = header_font
    c.fill = header_fill_purple
    c.alignment = align_center
    c.border = thin_border
ws3.row_dimensions[1].height = 28

act_names = {
    1: "HỒI 1: HOOK & BÁO ĐỘNG KHẨN CẤP",
    6: "HỒI 2: NGUYÊN LÝ 3 TẦNG QUẢN LÝ",
    12: "HỒI 3: PHÂN LOẠI CHẤT LIỆU & BẪY SILICONE",
    17: "HỒI 4: ĐIỂM HUYỆT CHỨA ĐỰNG VS THAO TÁC",
    22: "HỒI 5: CHIẾN LƯỢC HÀNH ĐỘNG THỰC CHIẾN",
    27: "HỒI 6: ĐÚC KẾT & KÊU GỌI HÀNH ĐỘNG"
}

current_act = "HỒI 1: HOOK & BÁO ĐỘNG KHẨN CẤP"
for row_idx, s in enumerate(SCENES, 2):
    stt = s["stt"]
    if stt in act_names:
        current_act = act_names[stt]
    
    start_sec = (stt - 1) * 10
    end_sec = stt * 10
    time_str = f"{start_sec//60:02d}:{start_sec%60:02d} - {end_sec//60:02d}:{end_sec%60:02d}"
    
    # Emotion note
    if stt in [1, 2]:
        emotion = "Lời chào ấm áp, chỉ ra nghịch lý kịch tính, cuốn hút"
    elif stt in [3, 4]:
        emotion = "Kể chuyện case study thực tế, bóc trần ngộ nhận pháp lý"
    elif stt == 5:
        emotion = "Đĩnh đạc, dứt khoát, phong thái chuyên gia mở lối"
    elif stt in [6, 7, 8, 9, 10, 11]:
        emotion = "Giảng giải rõ ràng, rành mạch, ngắt nghỉ từng tầng"
    elif stt in [12, 13, 14, 15, 16]:
        emotion = "Hạ giọng bật mí, chỉ ra nghịch lý, tạo cảm giác 'à ra thế'"
    elif stt in [17, 18, 19, 20, 21]:
        emotion = "Nhấn mạnh từ khóa then chốt, đanh thép viện dẫn công văn"
    elif stt in [22, 23, 24, 25, 26]:
        emotion = "Cầm tay chỉ việc, truyền cảm hứng giải pháp thực chiến"
    else:
        emotion = "Ấm áp, tin cậy, nụ cười trong giọng nói, kêu gọi tự nhiên"

    fill = zebra_fill if row_idx % 2 == 0 else white_fill
    vals = [stt, time_str, current_act, s["dialogue"], emotion]
    for col_idx, val in enumerate(vals, 1):
        c = ws3.cell(row=row_idx, column=col_idx, value=val)
        c.font = cell_font
        c.border = thin_border
        c.fill = fill
        if col_idx in [1, 2]:
            c.alignment = align_center
            c.font = bold_cell_font
        elif col_idx in [3, 5]:
            c.alignment = align_left
            if col_idx == 3:
                c.font = bold_cell_font
        else:
            c.alignment = align_left
    ws3.row_dimensions[row_idx].height = 45

ws3.column_dimensions["A"].width = 8
ws3.column_dimensions["B"].width = 16
ws3.column_dimensions["C"].width = 30
ws3.column_dimensions["D"].width = 50
ws3.column_dimensions["E"].width = 35

# ------------------------------------------------------------------------------
# SHEET 4: TONG_QUAN_PROJECT_STYLE (HỒ SƠ NGHỆ THUẬT & HƯỚNG DẪN FLOW)
# ------------------------------------------------------------------------------
ws4 = wb.create_sheet(title="TONG_QUAN_PROJECT_STYLE")
ws4.views.sheetView[0].showGridLines = True

ws4["A1"] = "HỒ SƠ DỰ ÁN VIDEO AI: QUY ĐỊNH NHẬP KHẨU BAO BÌ THỰC PHẨM (5 PHÚT)"
ws4["A1"].font = title_font
ws4["A2"] = "Căn cứ theo PO-2026-W42-64 & Cẩm nang nghiệp vụ xuất nhập khẩu thực chiến"
ws4["A2"].font = sub_font

project_meta = [
    ("Tên Dự Án", "Video Phân Tích & Hướng Dẫn: Bóc Tách Quy Định Nhập Khẩu Bao Bì & Dụng Cụ Tiếp Xúc Thực Phẩm"),
    ("Mã Phiếu Giao Việc", "PO-2026-W42-64"),
    ("Thời Lượng Video & Khối Lượng Sản Xuất", "5 Phút (300 Giây) — Gồm 30 phân cảnh x 10 giây/cảnh | 31 Ảnh Prompts Hàng Loạt (30 Cảnh + 1 Master Thumbnail)"),
    ("Tỷ Lệ Khung Hình", "16:9 (1920x1080 Full HD / 4K) — Định dạng chuẩn YouTube"),
    ("Cấu Trúc Nhịp Dựng", "Đan xen điện ảnh: 16 Cảnh A-roll Talking Head (53%) + 14 Cảnh B-roll Minh Họa Chuyên Biệt (47%)"),
    ("Phong Cách Thị Giác (Visual Style)", "Chân dung Talking-Head Chuyên Gia Độc Lập đan xen B-roll Điện Ảnh (Seaport, Macro công văn có dấu đỏ, Phòng Lab, Mẫu vật thực tế)"),
    ("Nhân Vật Đại Diện (Host / Persona)", "Chị Thúy (Chuyên gia Thực chiến XNK & Logistics) — Phong thái chuyên gia đĩnh đạc, vest navy sang trọng, tư vấn sắc bén, truyền cảm hứng thực chiến"),
    ("Tông Giọng Voiceover", "Vietnamese Female Business Expert — Giọng nữ chuẩn, đĩnh đạc, dứt khoát, ấm áp và giàu năng lượng thực chiến"),
    ("Tốc Độ Đọc (Pacing)", "10 giây / 20 - 25 từ tiếng Việt (tương đương 135 - 145 từ/phút)"),
    ("Căn Cứ Pháp Lý & SSoT", "Công văn 22082/CHQ-GSQL (23/09/2026), Công văn 1913/ATTP-KN, Thông tư 27/2026/TT-BYT, Thông tư 15/2024/TT-BYT, Nghị định 37/2026/NĐ-CP, Nghị định 15/2018/NĐ-CP, Công văn 4372/BYT-VPB6 (Aeon), QCVN 12-1, 12-2, 12-3, 12-4"),
    ("Công Cụ Khuyên Dùng", "Google Flow (flow.google.com), Kling AI v1.5, Runway Gen-3 Alpha, Hailuo Minimax, Midjourney v6.1 / DALL-E 3"),
    ("Cách Tải Lên Google Flow", "1. Mở file Bao_Bi_Thuc_Pham_5Phut_Kich_Ban_Flow.csv. 2. Truy cập flow.google.com -> Create Project -> Import CSV. 3. Hệ thống tự động phân chia 30 cảnh với đầy đủ Prompt, Audio và Voiceover.")
]

ws4.row_dimensions[4].height = 24
for idx, (label, val) in enumerate(project_meta, 5):
    c1 = ws4.cell(row=idx, column=1, value=label)
    c1.font = bold_cell_font
    c1.fill = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
    c1.border = thin_border
    c1.alignment = align_left
    
    c2 = ws4.cell(row=idx, column=2, value=val)
    c2.font = cell_font
    c2.border = thin_border
    c2.alignment = align_left
    ws4.row_dimensions[idx].height = 28

ws4.column_dimensions["A"].width = 32
ws4.column_dimensions["B"].width = 85

# Atomic save to satisfy Rule 10 (Zero backup files)
tmp_excel = excel_path + ".tmp.xlsx"
wb.save(tmp_excel)
os.replace(tmp_excel, excel_path)
print(f"SUCCESS: Đã xuất file Excel 4 Sheet -> {excel_path}")

# ==============================================================================
# XUẤT FILE 4: MARKDOWN STORYBOARD CHI TIẾT
# ==============================================================================
md_path = os.path.join(SCRATCH_DIR, "KICH_BAN_YOUTUBE_5P_BAO_BI_THUC_PHAM.md")
with open(md_path, "w", encoding="utf-8") as f:
    f.write("# 🎬 KỊCH BẢN YOUTUBE 5 PHÚT: QUY ĐỊNH NHẬP KHẨU BAO BÌ & DỤNG CỤ TIẾP XÚC THỰC PHẨM\n\n")
    f.write("> **Mã PO:** `PO-2026-W42-64` | **Thời lượng:** 5 Phút (300 giây — 30 Cảnh x 10s) | **Tỷ lệ:** 16:9  \n")
    f.write("> **Chủ đề:** Bóc tách Công văn 22082/CHQ-GSQL & 1913/ATTP-KN, Tháp 3 Tầng Pháp Lý, Phân biệt Chứa đựng vs Tiếp xúc (CV 4372) & Tuyệt chiêu gửi mẫu CPN tránh phạt Dem/Det  \n")
    f.write("> **Cấu trúc nhịp dựng:** Đan xen 16 Cảnh A-roll Talking Head (53%) + 14 Cảnh B-roll Minh Họa Chuyên Biệt (47%)  \n")
    f.write("> **Khối lượng sản xuất:** 30 Phân cảnh video Flow + 31 Ảnh Prompts Hàng Loạt (30 Cảnh + 1 Master Thumbnail)  \n")
    f.write("> **Nhân vật chính:** Chị Thúy (Chuyên gia Thực chiến XNK & Logistics)  \n\n")
    f.write("---\n\n")
    
    f.write("## 📋 I. BẢNG PHÂN CẢNH VIDEO CHUẨN GOOGLE FLOW (5 CỘT)\n\n")
    f.write("| STT | Tiêu đề | Mô tả phân cảnh (Prompt) | Lời thoại (Dialogue) | Thời lượng (giây) |\n")
    f.write("| :---: | :--- | :--- | :--- | :---: |\n")
    for s in SCENES:
        clean_prompt = s["prompt_en"].replace("\n", " ")
        f.write(f"| {s['stt']} | {s['title']} | {clean_prompt} | {s['dialogue']} | {s['duration']} |\n")
    
    f.write("\n---\n\n")
    f.write("## 🎨 II. BẢNG PROMPT TẠO ẢNH HÀNG LOẠT (31 ẢNH: 30 PHÂN CẢNH + 1 MASTER THUMBNAIL) (5 CỘT)\n\n")
    f.write("| STT | Tiêu đề | Prompt tạo ảnh (Image Prompt) | Góc quay (Shot Type) | Ghi chú |\n")
    f.write("| :---: | :--- | :--- | :---: | :--- |\n")
    for img in IMAGES_31:
        clean_img_prompt = img["prompt"].replace("\n", " ")
        f.write(f"| {img['stt']} | {img['title']} | {clean_img_prompt} | {img['shot_type']} | {img['note']} |\n")

print(f"SUCCESS: Đã xuất file Markdown Storyboard -> {md_path}")
