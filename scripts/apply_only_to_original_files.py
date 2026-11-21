# -*- coding: utf-8 -*-
r"""
Script xóa sạch mọi file backup và ghi đè trực tiếp duy nhất lên 2 file gốc:
1. C:\Users\thuyn\Downloads\phân cảnh.xlsx
2. C:\Users\thuyn\Downloads\ảnh.xlsx
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# 1. Danh sách file rác / backup cần xóa
cleanup_files = [
    r"C:\Users\thuyn\Downloads\ảnh_chuẩn_tiếng_anh.xlsx",
    r"C:\Users\thuyn\Downloads\ảnh_chuẩn_khớp_dropdown.xlsx",
    r"C:\Users\thuyn\Downloads\ảnh_người que.xlsx",
    r"C:\Users\thuyn\Downloads\ảnh_người que_chuan_goc_quay.xlsx",
    r"C:\Users\thuyn\Downloads\phân cảnh_người que.xlsx",
    r"C:\Users\thuyn\Downloads\Bao_Bi_Thuc_Pham_Nguoi_Que_Studio.xlsx",
    r"C:\Users\thuyn\Downloads\Tao_Anh_Hang_Loat_English.csv",
    r"C:\Users\thuyn\Downloads\Tao_Anh_Hang_Loat_Khop_Dropdown.csv",
    r"C:\Users\thuyn\Downloads\Tao_Anh_Hang_Loat_Nguoi_Que.csv",
    r"C:\Users\thuyn\Downloads\Kich_Ban_Phan_Canh_Nguoi_Que.csv",
    r"C:\Users\thuyn\Downloads\phân cảnh.xlsx.tmp.xlsx",
    r"C:\Users\thuyn\Downloads\ảnh.xlsx.tmp.xlsx"
]

for p in cleanup_files:
    if os.path.exists(p):
        try:
            os.remove(p)
            print(f"Deleted backup file: {p}")
        except Exception as e:
            print(f"Error deleting {p}: {e}")

# 2. Dữ liệu chuẩn cho cả 2 file gốc
data_30_scenes = [
    {
        "stt": 1,
        "title_pc": "Scene 01",
        "title_anh": "Câu Đố Hóc Búa Bao Bì Thực Phẩm",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure detective named Thuy with an adorable ponytail and tiny bow ribbon, "
            "wearing cute round detective glasses and a tiny detective cap. She holds a giant cartoon magnifying glass and winks playfully at the camera, placing a stick hand on her hip. "
            "Three animated doodle items float and rotate around her with glowing yellow question marks: a plastic food box, an orange silicone cake mold, and a shiny metal spoon. "
            "A funny 3-second cartoon countdown timer ticks above her head as she challenges the viewers. Camera gentle smooth push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, playful energetic female narrator, pacing 10s): \"Đố bạn một câu cực hóc búa: Giữa hộp nhựa, khuôn silicone và chiếc thìa ăn cơm — cái nào hải quan cho thông quan ngay, cái nào bị xích lại phạt cả trăm triệu?\" "
            "Audio & SFX: Mystery detective harp intro, playful ticking timer clock, cartoon question pop ding, bouncy playful bassline."
        ),
        "dialogue": "Đố bạn một câu cực hóc búa: Giữa hộp nhựa, khuôn silicone và chiếc thìa ăn cơm — cái nào hải quan cho thông quan ngay, cái nào bị xích lại phạt cả trăm triệu?",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure detective named Thuy with ponytail and tiny bow ribbon, wearing cute round detective glasses, "
            "holding a giant cartoon magnifying glass and winking challengingly at camera. Three floating doodle items revolve around her with glowing yellow question marks: "
            "a plastic container, a silicone cake mold, and a metal spoon. Clean minimalist white background, comic motion lines, expressive cartoon eyes, "
            "crisp black ink line art with selective vibrant yellow and cyan accents, 8k resolution --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy hóa thân thành thám tử tinh nghịch, đưa ra câu đố hóc búa với 3 món đồ bay lơ lửng kèm dấu hỏi để thu hút người xem ngay 3 giây đầu"
    },
    {
        "stt": 2,
        "title_pc": "Scene 02",
        "title_anh": "Cảnh Báo Kẹt Cảng & Thiệt Hại",
        "visual_prompt": (
            "Split-screen 2D doodle animation: On left side, a happy stick figure drives a cute mini container truck zooming smoothly past a green traffic light with a '1 NGÀY' badge. "
            "On right side, a giant cargo container is locked with heavy red chains and an oversized padlock. "
            "Floating cartoon red smoke numbers display '100.000.000 VNĐ' with warning flashes. Camera fast whip pan between both sides. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần, phạt cả trăm triệu!\" "
            "Audio & SFX: Fast truck zoom whoosh, followed by loud metal chain rattle and comic alarm buzzer."
        ),
        "dialogue": "Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần, phạt cả trăm triệu!",
        "duration": 10,
        "image_prompt": (
            "Wide shot 2D doodle illustration of busy seaport container terminal: rows of stacked shipping containers, a sharp oversized red cartoon padlock and heavy chains locking the container doors, "
            "floating red smoke numbers '100.000.000 VND', dramatic industrial comic sketch aesthetic, clean white paper background, crisp black ink outlines, 8k --ar 16:9"
        ),
        "shot_type": "Toàn cảnh (Wide Shot)",
        "note": "B-roll cảnh báo kẹt cảng container thực tế dạng doodle, niêm phong kẹp chì đỏ và số tiền phạt bốc khói"
    },
    {
        "stt": 3,
        "title_pc": "Scene 03",
        "title_anh": "Cận Cảnh Túi Màng Nhôm Vô Trùng",
        "visual_prompt": (
            "Macro 2D doodle animation: Detailed close-up of sterile multilayer aluminum aseptic packaging bags (Aseptic bag, HS code 3923.21.99) stacked inside clean industrial carton. "
            "Camera slow tilt-up from reflective metallic silver foil bag texture to an official corporate Self-Declaration dossier (Bản Tự công bố) with red stamp on desk. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!\" "
            "Audio & SFX: Metallic foil crinkling sound, subtle paper rustle, sparkling star ding."
        ),
        "dialogue": "Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!",
        "duration": 10,
        "image_prompt": (
            "Macro close-up shot of sterile food packaging aluminum aseptic bags with blue sanitary spouts, clean industrial product doodle aesthetic, metallic silver foil reflections, "
            "an official corporate Self-Declaration document with red seal lying beside, shallow depth of field, high-tech food manufacturing vector art, 8k --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "B-roll cận cảnh sản phẩm túi màng nhôm vô trùng Aseptic bag và hồ sơ Tự công bố"
    },
    {
        "stt": 4,
        "title_pc": "Scene 04",
        "title_anh": "Bẫy Ngộ Nhận Thông Tư 27",
        "visual_prompt": (
            "Top-down 2D doodle animation: Official legal document Circular 27/2026/TT-BYT open on desk. A bright yellow fluorescent highlighter marker emphasizes the regulatory clause: "
            "'Nhóm có mức độ rủi ro trung bình - Thực hiện tự công bố sản phẩm'. A giant red comic rubber stamp suddenly slams down: 'DỪNG HÀNG KIỂM TRA!'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!\" "
            "Audio & SFX: Marker squeak on paper, followed by harsh customs rejection stamp thud."
        ),
        "dialogue": "Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!",
        "duration": 10,
        "image_prompt": (
            "Close-up top-down shot of official legal document Circular 27/2026/TT-BYT on wooden desk, yellow fluorescent highlighter emphasizing regulatory paragraph 'Rủi ro trung bình', "
            "a large red customs rubber stamp 'STOP' hitting the paper, reading glasses beside, warm studio lighting, sharp focus vector doodle art --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Chụp cận văn bản Thông tư 27/2026/TT-BYT điểm bẫy ngộ nhận pháp lý và dấu đỏ dừng hàng"
    },
    {
        "stt": 5,
        "title_pc": "Scene 05",
        "title_anh": "Khẳng Định Bản Chất Văn Bản 22082",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy with ponytail rolls up her stick sleeves with a confident, reassuring smile. "
            "She holds a teacher's pointer with a glowing star tip, tapping a huge messy scribble ball of yarn, which instantly unknots into neat parallel golden threads! "
            "Beside her, two clean illuminated legal scrolls unfurl with titles 'Công văn 22082' and 'Công văn 1913'. Camera smooth tracking dolly. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Thực tế, Văn bản 22082 và 1913 đã giải thích rõ. Đừng lo, hôm nay Thúy sẽ gỡ rối toàn diện cho bạn!\" "
            "Audio & SFX: Magical harp glissando, unravelling whoosh, uplifting chime."
        ),
        "dialogue": "Thực tế, Văn bản 22082 và 1913 đã giải thích rõ. Đừng lo, hôm nay Thúy sẽ gỡ rối toàn diện cho bạn!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl named Thuy with adorable ponytail and ribbon, smiling confidently with reassuring open-hand gesture, "
            "beside two floating glowing parchment scrolls of Official Dispatch 22082 and 1913 with red seals, neat golden threads untangled, clean minimalist background, 8k --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy khẳng định văn bản 22082 và 1913, gỡ rối cuộn len thành đường thẳng mở ra giải pháp"
    },
    {
        "stt": 6,
        "title_pc": "Scene 06",
        "title_anh": "Đồ Họa 3D Tháp 3 Tầng Pháp Lý",
        "visual_prompt": (
            "3D isometric doodle animation: A modern 3-tiered compliance pyramid structure rises up with glowing neon layers: Tier 1 (Cửa khẩu - Blue), Tier 2 (Tự công bố - Green), Tier 3 (QCVN - Orange). "
            "A sparkling golden key hovers above the apex of the pyramid with radiating light beams. Camera dynamic circular sweep around the pyramid. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chìa khóa vàng nằm ở việc nắm vững Tháp 3 tầng quản lý pháp lý mà hầu hết mọi người hay nhầm lẫn!\" "
            "Audio & SFX: Golden key sparkle sound, 3 ascending musical synth notes, triumphant chime."
        ),
        "dialogue": "Chìa khóa vàng nằm ở việc nắm vững Tháp 3 tầng quản lý pháp lý mà hầu hết mọi người hay nhầm lẫn!",
        "duration": 10,
        "image_prompt": (
            "Wide shot isometric 3D animation style clean infographic visualization of a modern 3-tiered compliance pyramid structure, glowing neon layers labeled Tier 1, Tier 2, Tier 3, "
            "floating oversized golden key hovering on top with magical sparkles, clean dark navy and white gradient background, sharp 8k render --ar 16:9"
        ),
        "shot_type": "Toàn cảnh (Wide Shot)",
        "note": "Mô hình 3D trực quan Tháp 3 Tầng Quản Lý Pháp Lý với chìa khóa vàng"
    },
    {
        "stt": 7,
        "title_pc": "Scene 07",
        "title_anh": "Giảng Giải Tầng 1 - Cửa Khẩu Hải Quan",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy with ponytail stands beside a cartoon customs barrier gate. "
            "She holds a neat guidebook titled 'THÔNG TƯ 15/2024' with her stick hands, flipping to a page where a glowing blue laser pointer highlights specific HS code boxes. "
            "Packages with matching codes move toward a scanner, while non-matching ones cruise right through. Camera medium shot with gentle pan. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 1 là Thủ tục Kiểm tra nhà nước tại Cửa khẩu. Chỉ những mặt hàng có mã HS trong Thông tư 15 mới bị kiểm tra!\" "
            "Audio & SFX: Book page flip sound, high-tech scanner beep, confident affirmative chime."
        ),
        "dialogue": "Tầng 1 là Thủ tục Kiểm tra nhà nước tại Cửa khẩu. Chỉ những mặt hàng có mã HS trong Thông tư 15 mới bị kiểm tra!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy with ponytail seated at modern desk, holding a printed regulatory booklet of Circular 15/2024, "
            "looking directly at camera with clear informative demeanor, miniature cartoon customs gate in background, clean minimalist line art --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy phân tích Tầng 1 Cửa khẩu theo Thông tư 15"
    },
    {
        "stt": 8,
        "title_pc": "Scene 08",
        "title_anh": "Cận Cảnh Nhựa PP & Núm Vú Cao Su",
        "visual_prompt": (
            "Macro split 2D doodle animation: Left frame shows transparent plastic food container with prominent recycling code PP-05. "
            "Right frame shows cute baby bottle with rubber pacifier labeled HS 4014.90.10. A red glowing warning border pulses around both items with bold text: 'BẮT BUỘC KIỂM TRA CỬA KHẨU'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!\" "
            "Audio & SFX: Double cartoon warning buzzer, siren whoop, barcode scanner beep."
        ),
        "dialogue": "Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!",
        "duration": 10,
        "image_prompt": (
            "Close-up macro shot showing transparent PP food storage container with recycling code PP-05 beside baby rubber silicone feeding bottle nipple labeled HS 4014.90.10, "
            "clean white product background, sharp macro details, red warning border graphic, no photographic noise, clean vector commercial look --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Cận cảnh nhựa tổng hợp PP/PE và núm vú cao su trẻ em"
    },
    {
        "stt": 9,
        "title_pc": "Scene 09",
        "title_anh": "Danh Mục Sản Phẩm Miễn Cửa Khẩu",
        "visual_prompt": (
            "2D doodle product showcase: Premium stainless steel cooking pot, fine ceramic plate, clear glass beverage bottle, and silicone baking spatula. "
            "Camera smooth panning across items. Bright green checkmark badges 'MIỄN KIỂM TRA CỬA KHẨU' illuminate sequentially above each item as they sprout tiny wings and fly past the customs gate. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!\" "
            "Audio & SFX: Soft metallic chime, 4 crisp positive success pings, upbeat acoustic strum."
        ),
        "dialogue": "Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!",
        "duration": 10,
        "image_prompt": (
            "Medium shot clean product showcase of stainless steel pot, white ceramic plate, glass jar, and blue silicone baking mold arranged neatly on marble countertop, "
            "bright green checkmark badges floating above with text 'MIỄN KIỂM TRA CỬA KHẨU', bright softbox illumination, premium commercial doodle look --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "B-roll sản phẩm inox, sứ, thủy tinh được miễn kiểm tra cửa khẩu"
    },
    {
        "stt": 10,
        "title_pc": "Scene 10",
        "title_anh": "Giảng Giải Tầng 2 - Thủ Tục Tự Công Bố",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy with ponytail sits at a cute desk with a laptop, explaining the second administrative layer. "
            "She stamps a folder 'TỰ CÔNG BỐ - NĐ 37', folds it into an origami paper airplane, and tosses it towards an animated city hall building labeled 'Ban Quản Lý ATTP'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!\" "
            "Audio & SFX: Fast cartoon typing sound, paper airplane swoosh, satisfying soft landing bell."
        ),
        "dialogue": "Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy with ponytail seated in front of laptop in modern office, explaining regulatory compliance with poised professional posture, "
            "stamping folder labeled 'TỰ CÔNG BỐ', clean natural lighting, executive doodle backdrop --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy hướng dẫn Tầng 2 Tự công bố lưu hành thị trường"
    },
    {
        "stt": 11,
        "title_pc": "Scene 11",
        "title_anh": "Phòng Lab Thử Nghiệm Chuẩn QCVN",
        "visual_prompt": (
            "2D doodle animation inside certified food safety laboratory: A chemical test tube sample transforms into a giant glowing emerald-green energy shield emblazoned with 'QCVN 12-1:2011/BYT - TEST PASSED'. "
            "Computer monitor in background displays passing chemical migration curve. Inspection arrows bounce off the shield harmlessly. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!\" "
            "Audio & SFX: Clean laboratory machine beep, centrifuge hum, deep confident bass swell."
        ),
        "dialogue": "Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot inside certified food contact safety testing laboratory in clean 2D doodle animation style: analytical test tubes and spectrophotometer, "
            "a glowing green holographic defense shield labeled 'QCVN 12-1:2011/BYT - PASSED', pristine sterile scientific environment, 8k --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "B-roll phòng lab thử nghiệm chỉ tiêu thôi nhiễm QCVN"
    },
    {
        "stt": 12,
        "title_pc": "Scene 12",
        "title_anh": "Cạm Bẫy Cao Su Trẻ Em vs Silicone",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy tiptoes close to camera lens, putting a stick finger over her lips with a comical hushed 'shhh' expression. "
            "The background dims with a dramatic red spotlight focusing on her. She pulls out a baby pacifier in one hand and a silicone cake mold in the other, with an animated neon trap sign flashing '80% MẮC BẪY!'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!\" "
            "Audio & SFX: Cartoon tiptoe woodblock sounds, mystery suspense whoosh, dramatic stinger chord."
        ),
        "dialogue": "Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of cute female stick figure girl Thuy with ponytail, leaning forward with engaging, intense storytelling expression, "
            "holding baby pacifier in left hand and silicone mold in right hand, animated red warning badge '80% MẮC BẪY', dark background with soft dramatic spotlight, 8k --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Cô bé người que Thúy hạ giọng chỉ ra điểm thắt cổ chai 80% mắc bẫy"
    },
    {
        "stt": 13,
        "title_pc": "Scene 13",
        "title_anh": "Bóc Tách Nghịch Lý Khuôn Silicone Người Lớn",
        "visual_prompt": (
            "Top-down 2D doodle animation: Split graphic comparison on wooden table: Left frame displays Circular 15 table without silicone mold HS code (marked 'Miễn cửa khẩu'). "
            "Right frame displays Circular 27 Item 6: 'Bao bì dụng cụ cao su chứa đựng' (marked 'Bắt buộc Tự công bố'). A flexible orange silicone cake mold rests at the intersection. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!\" "
            "Audio & SFX: Soft rubber flex sound, paper rustle, electronic highlight buzz."
        ),
        "dialogue": "Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!",
        "duration": 10,
        "image_prompt": (
            "Close-up split graphic comparison on executive wooden table: official Circular 15 table contrasted with Circular 27 Item 6, "
            "colorful silicone cake mold resting on document corner, high resolution legal audit vector scene, clean minimalist line art --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Bóc tách nghịch lý silicone người lớn miễn cửa khẩu nhưng phải TCB"
    },
    {
        "stt": 14,
        "title_pc": "Scene 14",
        "title_anh": "Đồ Gia Dụng Inox & Gốm Sứ Nhóm 1",
        "visual_prompt": (
            "Close-up 2D doodle animation: Gleaming stainless steel saucepan, fine porcelain soup bowl, and crystalline glass tumbler. "
            "A green official badge stamps over items with text: 'NHÓM 1: RỦI RO THẤP - KHÔNG PHẢI NỘP HỒ SƠ LÊN SỞ'. Sunlight reflects brightly off polished surfaces. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!\" "
            "Audio & SFX: Crystal glass clink, metallic ring, positive confirmation tone."
        ),
        "dialogue": "Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot commercial product illustration of polished stainless steel pan, elegant white ceramic bowl, and clear glass tumbler, "
            "green badge 'NHÓM 1: RỦI RO THẤP - MIỄN NỘP HỒ SƠ LÊN SỞ', pristine reflections, clean vector product art, 8k resolution --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "B-roll nồi inox, bát sứ, ly thủy tinh thuộc nhóm rủi ro thấp"
    },
    {
        "stt": 15,
        "title_pc": "Scene 15",
        "title_anh": "Hướng Dẫn Hồ Sơ TCCS Phòng Thủ Hậu Kiểm",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy holding an organized black corporate filing binder labeled 'HỒ SƠ TCCS NỘI BỘ'. "
            "She locks it into a cute mini-safe with a brass key, winking reassuringly at the camera, explaining internal compliance defense. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!\" "
            "Audio & SFX: Binder ring snap click, secure lock latch sound."
        ),
        "dialogue": "Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of cute female stick figure girl Thuy with ponytail holding an organized black corporate filing binder labeled 'HỒ SƠ TCCS NỘI BỘ', "
            "confident advisory demeanor, miniature cute mini-safe beside, clean modern doodle setting, 8k --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Cô bé người que Thúy hướng dẫn lưu trữ hồ sơ Tiêu chuẩn cơ sở nội bộ"
    },
    {
        "stt": 16,
        "title_pc": "Scene 16",
        "title_anh": "Tối Ưu Tiết Kiệm Chi Phí Hàng Triệu Đồng",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy smiles with pride and warmth, hugging an oversized smiling golden cartoon piggy bank. "
            "Digital financial savings graphic pops up: 'Tiết kiệm 1.000.000 đ/model Inox & Thủy tinh'. Shiny gold coins rain down joyfully. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!\" "
            "Audio & SFX: Cheerful register cha-ching sound, upbeat synth chime."
        ),
        "dialogue": "Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy with an encouraging smile, hands gesturing savings beside a giant golden piggy bank, "
            "floating graphic '+1.000.000 VNĐ / Model', cartoon money bills floating down, clean vector doodle art --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy nhấn mạnh giá trị tiết kiệm tiền bạc cho doanh nghiệp"
    },
    {
        "stt": 17,
        "title_pc": "Scene 17",
        "title_anh": "So Sánh Dụng Cụ Chứa Đựng vs Tiếp Xúc",
        "visual_prompt": (
            "Medium shot 2D doodle product comparison setup on dark slate table: Left side: Plastic food storage box with sealed airtight locking lid (chứa đựng). "
            "Right side: Stainless steel dinner fork, spoon, wooden chopsticks and cooking spatula (tiếp xúc không chứa đựng). Camera slow tracking dolly between the two groups. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!\" "
            "Audio & SFX: Smooth camera dolly sound, subtle low-frequency suspense beat."
        ),
        "dialogue": "Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!",
        "duration": 10,
        "image_prompt": (
            "Medium shot studio product comparison on dark slate surface: food storage box with sealed lid on left, metal spoon fork chopsticks on right, "
            "cinematic edge lighting, clean and crisp 2D vector product art, high contrast doodle aesthetic --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "So sánh dụng cụ chứa đựng vs dụng cụ tiếp xúc không chứa đựng"
    },
    {
        "stt": 18,
        "title_pc": "Scene 18",
        "title_anh": "Kim Bài Công Văn 4372 Bộ Y Tế (Aeon)",
        "visual_prompt": (
            "Close-up 2D doodle document shot: Official letterhead of Ministry of Health (Bộ Y tế) - Dispatch 4372/BYT-VPB6 dated 31/07/2019 addressing Aeon Vietnam Co., Ltd. "
            "Warm dramatic spotlight illuminates official ministry header, red national emblem stamp and incoming registration number. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!\" "
            "Audio & SFX: Paper rustle, spotlight switch on sound, royal trumpet chime."
        ),
        "dialogue": "Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of official Vietnamese Ministry of Health dispatch 4372/BYT-VPB6, authentic red stamp emblem, spotlight illumination on official seal, "
            "high legal credibility parchment scroll doodle, crisp black ink details --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Chụp cận Công văn 4372/BYT-VPB6 của Bộ Y tế trả lời Aeon"
    },
    {
        "stt": 19,
        "title_pc": "Scene 19",
        "title_anh": "Thìa Đũa Dụng Cụ Thao Tác Cơ Học",
        "visual_prompt": (
            "Top-down flat-lay 2D doodle animation: Stainless steel soup spoon, wooden chopsticks, fruit peeling knife, food tongs, silicone kitchen gloves. "
            "A bold red rubber stamp stamps across: 'KHÔNG PHẢI DỤNG CỤ CHỨA ĐỰNG!'. Utensils bounce with cartoon sound effects. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!\" "
            "Audio & SFX: Playful kitchen clinking, boing sound, rubber stamp thud."
        ),
        "dialogue": "Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!",
        "duration": 10,
        "image_prompt": (
            "Close-up flat lay shot collage of kitchen utensils: metal spoons, wooden chopsticks, fruit peeling knife, food tongs, silicone cooking gloves, "
            "bold red banner 'KHÔNG PHẢI DỤNG CỤ CHỨA ĐỰNG', studio lighting, clean 2D vector doodle --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "B-roll các dụng cụ không chứa đựng: đũa, thìa, dĩa, dao gọt, kẹp gắp"
    },
    {
        "stt": 20,
        "title_pc": "Scene 20",
        "title_anh": "Khẳng Định Miễn TCB & Miễn Kiểm Tra Cửa Khẩu",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy makes two decisive outward hand gestures, nodding firmly. "
            "Two red barrier signs 'TỰ CÔNG BỐ' and 'KIỂM TRA CỬA KHẨU' get crossed out and turn into glowing green checkmarks 'MIỄN 100%'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!\" "
            "Audio & SFX: Powerful double swoosh sound, magic transformation sparkle, triumphant chord."
        ),
        "dialogue": "Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy gesturing decisively with open hands, confident expert posture, "
            "green channel approval graphic 'MIỄN 100%', clean minimalist vector background, bright lighting --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy khẳng định miễn TCB và miễn kiểm tra cửa khẩu"
    },
    {
        "stt": 21,
        "title_pc": "Scene 21",
        "title_anh": "Xe Container Thông Quan Luồng Xanh Ra Cảng",
        "visual_prompt": (
            "Wide shot 2D doodle animation: Heavy container reach stacker smoothly lowering a shipping container onto a delivery truck. "
            "Truck accelerates smoothly and drives out through the seaport terminal barrier gate labeled 'LUỒNG XANH - THÔNG QUAN 1 NGÀY' under clear blue sky. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!\" "
            "Audio & SFX: Truck engine accelerating, container hydraulic hiss, uplifting corporate harmony."
        ),
        "dialogue": "Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!",
        "duration": 10,
        "image_prompt": (
            "Wide shot dynamic action doodle of reach stacker vehicle lifting shipping container onto modern delivery truck at seaport terminal, "
            "speeding past open green customs barrier 'LUỒNG XANH', bright blue sky, clean vector line art --ar 16:9"
        ),
        "shot_type": "Toàn cảnh (Wide Shot)",
        "note": "B-roll xe container lấy hàng thông quan nhanh chóng ra khỏi cảng"
    },
    {
        "stt": 22,
        "title_pc": "Scene 22",
        "title_anh": "Cảnh Báo Mốc Thời Gian Lab 10-12 Ngày",
        "visual_prompt": (
            "Close-up 2D doodle animation: Analytical test tube rack with colorful chemical extraction reagents. "
            "A digital calendar animation flips from Day 1 through Day 12 consecutively with a rapidly ticking analog clock graphic. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!\" "
            "Audio & SFX: Rapid clock ticking SFX, electronic calendar flip sound."
        ),
        "dialogue": "Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of laboratory analytical test tube rack with colorful chemical solutions, digital countdown calendar flip animation overlay showing 'Day 12', "
            "analog clock with spinning hands, clean scientific doodle aesthetic --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Cảnh báo mốc thời gian thử nghiệm phòng lab 10-12 ngày"
    },
    {
        "stt": 23,
        "title_pc": "Scene 23",
        "title_anh": "Ác Mộng Hóa Đơn Phạt Dem/Det 15-30 Triệu",
        "visual_prompt": (
            "Medium close-up 2D doodle animation: A shipping line demurrage/detention invoice stamped with overdue penalty: 'DEMURRAGE / DETENTION OVERDUE: 30.000.000 VND'. "
            "In background, a long line of idle containers sitting under dark overcast clouds at port. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!\" "
            "Audio & SFX: Heavy financial penalty buzzer, ominous low synth drone."
        ),
        "dialogue": "Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot dramatic doodle of demurrage fee penalty document stamped 'DEMURRAGE: 30.000.000 VND', held against blurry dark port container background, "
            "storm clouds and rain, high contrast comic lighting --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Cảnh báo tiền phạt lưu container lưu bãi Dem/Det 15-30 triệu"
    },
    {
        "stt": 24,
        "title_pc": "Scene 24",
        "title_anh": "Tuyệt Chiêu Gửi Mẫu CPN Trước 25 Ngày",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy with ponytail smiles brightly, holding up an international air express envelope (CPN) equipped with tiny wings. "
            "She points out the 25-day advance window on an animated calendar. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!\" "
            "Audio & SFX: Parcel opening sound, upbeat brass chime, rocket whoosh."
        ),
        "dialogue": "Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy with ponytail holding an air express document envelope (CPN), smiling enthusiastically, "
            "calendar overlay showing '-25 Ngày', sharing insider advice, clean minimalist background --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy cầm phong bì CPN chia sẻ tuyệt chiêu gửi mẫu trước 25 ngày"
    },
    {
        "stt": 25,
        "title_pc": "Scene 25",
        "title_anh": "Màn Hình Cổng Một Cửa NSW Đạt Chuẩn",
        "visual_prompt": (
            "Close-up 2D doodle screen capture: Vietnam National Single Window (Cổng thông tin một cửa quốc gia NSW) interface screen. "
            "Status column displays: 'ĐÃ TIẾP NHẬN HỒ SƠ - KẾT QUẢ: ĐẠT YÊU CẦU NHẬP KHẨU (MẪU 05)'. Customs electronic signature stamp illuminates with a bright green verification badge. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!\" "
            "Audio & SFX: Electronic stamp sound, digital success fanfare."
        ),
        "dialogue": "Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of computer screen displaying Vietnam National Single Window clearance certificate with green approved stamp 'ĐẠT YÊU CẦU - 3 NGÀY', "
            "clean modern office doodle setup, celebratory digital badge --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Màn hình Một cửa quốc gia NSW cấp thông báo đạt trong 3 ngày"
    },
    {
        "stt": 26,
        "title_pc": "Scene 26",
        "title_anh": "Xe Cont Cập Kho Riêng Cắt Phí Lưu Bãi",
        "visual_prompt": (
            "Wide shot 2D doodle animation: Container delivery truck smoothly backing into a spacious logistics warehouse loading dock labeled 'KHO RIÊNG BẢO QUẢN'. "
            "Automated warehouse roll-up door lifts open, then shuts securely cutting off demurrage fees. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!\" "
            "Audio & SFX: Truck backing-up beep, warehouse automated door rolling up, air brake hiss."
        ),
        "dialogue": "Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!",
        "duration": 10,
        "image_prompt": (
            "Wide shot of modern clean distribution warehouse interior in 2D vector style, container truck docking at loading bay labeled 'KHO RIÊNG BẢO QUẢN', "
            "bright logistics facility, cartoon padlock on shutter door, clean line art --ar 16:9"
        ),
        "shot_type": "Toàn cảnh (Wide Shot)",
        "note": "Đưa container về kho riêng bảo quản cắt đứt tiền lưu kho bãi"
    },
    {
        "stt": 27,
        "title_pc": "Scene 27",
        "title_anh": "Bản Tự Công Bố Có Giá Trị Vô Thời Hạn",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy spreads hands wide, highlighting the immense long-term value. "
            "Floating golden graphic card appears beside her: 'HIỆU LỰC VÔ THỜI HẠN - LÀM 1 LẦN DÙNG MÃI MÃI' with a sparkling infinity symbol '∞'. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!\" "
            "Audio & SFX: Golden chime sound, triumphant synthesizer chord."
        ),
        "dialogue": "Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy with hands open in a welcoming, expansive gesture, genuine smile, "
            "floating golden card with infinity symbol '∞' and text 'HIỆU LỰC VÔ THỜI HẠN', clean minimalist vector art --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy nhấn mạnh bản Tự công bố có hiệu lực vô thời hạn"
    },
    {
        "stt": 28,
        "title_pc": "Scene 28",
        "title_anh": "Lưu Ý Phiếu Kiểm Nghiệm Hạn 12 Tháng",
        "visual_prompt": (
            "Close-up 2D doodle document shot: Official laboratory testing certificate with certified stamp and red official seal. "
            "A calendar bracket highlights the 12-month validity window. An elegant glass hourglass rests beside the document with sand trickling down gently. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!\" "
            "Audio & SFX: Soft clock ticking, gentle advisory reminder tone, woodblock chime."
        ),
        "dialogue": "Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of laboratory analytical certificate with blue accreditation stamp and red date stamp labeled 'HẠN 12 THÁNG', "
            "resting beside an elegant glass hourglass with blue flowing sand, warm desk lamp lighting, clean vector art --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Lưu ý phiếu kiểm nghiệm có giá trị 12 tháng khi nộp hồ sơ"
    },
    {
        "stt": 29,
        "title_pc": "Scene 29",
        "title_anh": "Đúc Kết Thông Điệp Triết Lý Thực Chiến",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy looks straight into the camera lens with intense, inspiring conviction. "
            "Behind her, an unshakeable stone fortress wall labeled 'HÀNG RÀO PHÁP LÝ' stands strong shielding cargo containers. Camera very slow push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!\" "
            "Audio & SFX: Deep inspiring piano chords swell gracefully."
        ),
        "dialogue": "Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!",
        "duration": 10,
        "image_prompt": (
            "Close-up shot of cute female stick figure girl Thuy looking directly into lens, powerful leadership presence, "
            "cute captain hat, protective stone wall shielding cargo containers in background, visionary comic doodle style --ar 16:9"
        ),
        "shot_type": "Cận cảnh (Close-up)",
        "note": "Cô bé người que Thúy đúc kết thông điệp triết lý logistics thực chiến"
    },
    {
        "stt": 30,
        "title_pc": "Scene 30",
        "title_anh": "Outro Kêu Gọi Subscribe & Tặng Cẩm Nang",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy waves hand gently with radiant warm smile, gesturing downwards to comment section. "
            "Red Subscribe button with bell notification and floating gift handbook 'CẨM NANG NGHIỆP VỤ NHẬP KHẨU BAO BÌ THỰC PHẨM' emerge beside her. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!\" "
            "Audio & SFX: Cheerful YouTube outro chime, bell notification ding, mouse click sound."
        ),
        "dialogue": "Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!",
        "duration": 10,
        "image_prompt": (
            "Medium shot of cute female stick figure girl Thuy smiling radiantly, waving hand gracefully towards viewer, "
            "beside a red Subscribe button, ringing bell, and gift-wrapped handbook 'CẨM NANG BAO BÌ', hearts floating, clean vector doodle --ar 16:9"
        ),
        "shot_type": "Trung cảnh (Medium Shot)",
        "note": "Cô bé người que Thúy tươi cười vẫy tay chào tạm biệt và tặng Cẩm nang"
    }
]

# Định dạng chung
header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

# 3. Ghi trực tiếp đè vào file gốc: C:\Users\thuyn\Downloads\phân cảnh.xlsx
pc_headers = ["STT", "Tiêu đề", "Mô tả phân cảnh (Prompt)", "Lời thoại (Dialogue)", "Thời lượng (giây)"]
wb_pc = openpyxl.Workbook()
ws_pc = wb_pc.active
ws_pc.title = "KICH_BAN_FLOW"
ws_pc.append(pc_headers)
pc_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")

for c in range(1, 6):
    cell = ws_pc.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = pc_fill
    cell.alignment = center_align

for s in data_30_scenes:
    ws_pc.append([s["stt"], s["title_pc"], s["visual_prompt"], s["dialogue"], s["duration"]])

for r in range(2, len(data_30_scenes) + 2):
    ws_pc.cell(row=r, column=1).alignment = center_align
    ws_pc.cell(row=r, column=2).alignment = center_align
    ws_pc.cell(row=r, column=3).alignment = left_align
    ws_pc.cell(row=r, column=4).alignment = left_align
    ws_pc.cell(row=r, column=5).alignment = center_align
    for c in range(1, 6):
        cell = ws_pc.cell(row=r, column=c)
        cell.font = Font(name="Arial", size=10)
        cell.border = thin_border

ws_pc.column_dimensions['A'].width = 8
ws_pc.column_dimensions['B'].width = 15
ws_pc.column_dimensions['C'].width = 60
ws_pc.column_dimensions['D'].width = 45
ws_pc.column_dimensions['E'].width = 18

target_pc = r"C:\Users\thuyn\Downloads\phân cảnh.xlsx"
wb_pc.save(target_pc)
print(f"SUCCESS: Written directly to original file -> {target_pc}")

# 4. Ghi trực tiếp đè vào file gốc: C:\Users\thuyn\Downloads\ảnh.xlsx
anh_headers = ["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"]
wb_anh = openpyxl.Workbook()
ws_anh = wb_anh.active
ws_anh.title = "TAO_ANH_HANG_LOAT"
ws_anh.append(anh_headers)
anh_fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")

for c in range(1, 6):
    cell = ws_anh.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = anh_fill
    cell.alignment = center_align

for s in data_30_scenes:
    ws_anh.append([s["stt"], s["title_anh"], s["image_prompt"], s["shot_type"], s["note"]])

for r in range(2, len(data_30_scenes) + 2):
    ws_anh.cell(row=r, column=1).alignment = center_align
    ws_anh.cell(row=r, column=2).alignment = left_align
    ws_anh.cell(row=r, column=3).alignment = left_align
    ws_anh.cell(row=r, column=4).alignment = center_align
    ws_anh.cell(row=r, column=5).alignment = left_align
    for c in range(1, 6):
        cell = ws_anh.cell(row=r, column=c)
        cell.font = Font(name="Arial", size=10)
        cell.border = thin_border

ws_anh.column_dimensions['A'].width = 8
ws_anh.column_dimensions['B'].width = 35
ws_anh.column_dimensions['C'].width = 65
ws_anh.column_dimensions['D'].width = 25
ws_anh.column_dimensions['E'].width = 45

target_anh = r"C:\Users\thuyn\Downloads\ảnh.xlsx"
wb_anh.save(target_anh)
print(f"SUCCESS: Written directly to original file -> {target_anh}")
