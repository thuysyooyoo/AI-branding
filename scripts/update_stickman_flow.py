# -*- coding: utf-8 -*-
"""
Script chuyển đổi kịch bản Google Flow và Bảng Prompt tạo ảnh
từ phong cách Talking-Head sang Phong cách Hoạt hình Người que (Cute Female Stick Figure).
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import csv
import os

scenes_data = [
    # Scene 1
    {
        "stt": 1,
        "title_pc": "Scene 01",
        "title_anh": "[Người que] Câu Đố Hóc Búa Về Bao Bì Thực Phẩm",
        "shot_type": "Medium Shot (Detective Pose)",
        "dialogue": "Đố bạn một câu cực hóc búa: Giữa hộp nhựa, khuôn silicone và chiếc thìa ăn cơm — cái nào hải quan cho thông quan ngay, cái nào bị xích lại phạt cả trăm triệu?",
        "duration": 10,
        "note": "Cô bé người que Thúy hóa thân thành thám tử tinh nghịch, đưa ra câu đố hóc búa với 3 món đồ bay lơ lửng kèm dấu hỏi để thu hút người xem ngay 3 giây đầu",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure detective named Thuy with an adorable ponytail and tiny bow ribbon, "
            "wearing cute round detective glasses and a tiny detective cap. She holds a giant cartoon magnifying glass and winks playfully at the camera, placing a stick hand on her hip. "
            "Three animated doodle items float and rotate around her with glowing yellow question marks: a plastic food box, an orange silicone cake mold, and a shiny metal spoon. "
            "A funny 3-second cartoon countdown timer ticks above her head as she challenges the viewers. Camera gentle smooth push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, playful energetic female narrator, pacing 10s): \"Đố bạn một câu cực hóc búa: Giữa hộp nhựa, khuôn silicone và chiếc thìa ăn cơm — cái nào hải quan cho thông quan ngay, cái nào bị xích lại phạt cả trăm triệu?\" "
            "Audio & SFX: Mystery detective harp intro, playful ticking timer clock, cartoon question pop ding, bouncy playful bassline."
        ),
        "image_prompt": (
            "Playful 2D doodle animation illustration: Cute female stick figure detective named Thuy with ponytail and tiny bow ribbon, "
            "wearing cute round detective glasses, holding a giant cartoon magnifying glass and winking challengingly at camera. "
            "Three floating doodle items revolve around her with glowing yellow question marks: a plastic container, a silicone cake mold, and a metal spoon. "
            "Clean minimalist white background, comic motion lines, expressive cartoon eyes, crisp black ink line art with selective vibrant yellow and cyan accents, 8k resolution --ar 16:9"
        )
    },
    # Scene 2
    {
        "stt": 2,
        "title_pc": "Scene 02",
        "title_anh": "[Người que] Cảnh Báo Kẹt Cảng & Thiệt Hại Phạt Tiền",
        "shot_type": "Wide Shot (Split Scene)",
        "dialogue": "Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần, phạt cả trăm triệu!",
        "duration": 10,
        "note": "Lật mở nghịch lý từ câu đố: bên trái xe tải thông quan 1 ngày bon bon, bên phải cô bé người que toát mồ hôi bên container bị xích phạt 100 triệu",
        "visual_prompt": (
            "Split-screen 2D doodle animation revealing the riddle paradox: On left side, a happy stick figure drives a cute mini container truck zooming smoothly past a green traffic light with a '1 NGÀY' badge. "
            "On right side, our cute female stick figure with ponytail looks shocked with wide comical cartoon eyes and sweat drops beside a giant shipping container locked with heavy red chains and an oversized padlock. "
            "Floating cartoon red smoke numbers display '100.000.000 VNĐ' with warning flashes. Camera fast whip pan between both sides. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Cùng là bao bì, có doanh nghiệp thông quan trong 1 ngày, nhưng có bên bị kẹt cứng cả tuần, phạt cả trăm triệu!\" "
            "Audio & SFX: Fast truck zoom whoosh, followed by loud metal chain rattle and comic alarm buzzer."
        ),
        "image_prompt": (
            "Humorous split-screen 2D doodle illustration: on left, a stick figure truck speeding past a green checkmark; on right, cute female stick figure girl with ponytail holding her head in comic shock beside a giant locked cargo container wrapped in red chains with smoking numbers '100.000.000 VND', clean vector doodle style, white background --ar 16:9"
        )
    },
    # Scene 3
    {
        "stt": 3,
        "title_pc": "Scene 03",
        "title_anh": "[Người que] Case Study 4 Container Túi Màng Nhôm",
        "shot_type": "Medium Shot",
        "dialogue": "Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!",
        "duration": 10,
        "note": "Cô bé người que cầm túi màng nhôm ánh bạc và vẫy tờ Tự công bố với vẻ mặt tự tin đinh ninh",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure girl holding an oversized shiny silver cartoon aseptic pouch in one hand and proudly waving an official stamped paper labeled 'BẢN TỰ CÔNG BỐ' in the other. "
            "In background, 4 stacked cartoon shipping containers stamped with cute milk carton icons. She winks confidently with closed-eyes smile and a glowing cartoon star glint. Camera smooth tilt-up. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Điển hình là lô 4 container túi màng nhôm vô trùng đóng sữa vừa qua. Doanh nghiệp đinh ninh có bản Tự công bố là xong!\" "
            "Audio & SFX: Foil crinkling sound effect, comic paper flap, sparkling star ding."
        ),
        "image_prompt": (
            "Cute female stick figure girl with ponytail proudly presenting an oversized shiny silver aseptic food packaging bag in one hand and a cartoon stamped certificate in the other, smiling with a confident wink, four colorful cartoon shipping containers with milk icons in background, clean 2D vector line art --ar 16:9"
        )
    },
    # Scene 4
    {
        "stt": 4,
        "title_pc": "Scene 04",
        "title_anh": "[Người que] Bẫy Ngộ Nhận Thông Tư 27 & Dừng Hàng",
        "shot_type": "Close-up",
        "dialogue": "Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!",
        "duration": 10,
        "note": "Cô bé người que soi kính lúp vào Thông tư 27 thì bị chú hải quan người que giơ biển STOP và đóng dấu dừng hàng",
        "visual_prompt": (
            "2D stick figure comic animation: Cute female stick figure examines a giant cartoon document labeled 'Thông tư 27' with an oversized magnifying glass, highlighting 'RỦI RO TRUNG BÌNH' in bright yellow. "
            "Suddenly, a funny cartoon customs officer stick figure with a cap pops up, raising a giant red STOP paddle, and slams down a giant red rubber stamp that reads 'DỪNG HÀNG KIỂM TRA!'. "
            "Our stick girl's ponytail stands up in surprise. Camera dramatic punch-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Họ ngộ nhận Thông tư 27 xếp bao bì vào mức rủi ro trung bình, cho tới khi hải quan dừng hàng kiểm tra nhà nước!\" "
            "Audio & SFX: Highlighter squeak, sudden whistle blow, heavy comic rubber stamp thud."
        ),
        "image_prompt": (
            "Dynamic 2D doodle animation scene: cute female stick figure with ponytail inspecting an open legal document with a giant magnifying glass, while a humorous stick figure customs officer slams a giant red 'STOP' stamp onto the desk, comic motion lines, clean vector aesthetic --ar 16:9"
        )
    },
    # Scene 5
    {
        "stt": 5,
        "title_pc": "Scene 05",
        "title_anh": "[Người que] Gỡ Rối Toàn Diện Với Văn Bản 22082 & 1913",
        "shot_type": "Medium Shot",
        "dialogue": "Thực tế, Văn bản 22082 và 1913 đã giải thích rõ. Đừng lo, hôm nay Thúy sẽ gỡ rối toàn diện cho bạn!",
        "duration": 10,
        "note": "Cô bé người que Thúy cầm đũa phép/que chỉ bảng gỡ cuộn len rối thành những đường thẳng tắp bên 2 văn bản giải thích",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy rolls up her stick sleeves with a determined, reassuring smile. "
            "She holds a teacher's pointer with a glowing star tip, tapping a huge messy scribble ball of yarn, which instantly unknots into neat parallel golden threads! "
            "Beside her, two clean illuminated legal scrolls unfurl with titles 'Công văn 22082' and 'Công văn 1913'. Camera smooth tracking dolly. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Thực tế, Văn bản 22082 và 1913 đã giải thích rõ. Đừng lo, hôm nay Thúy sẽ gỡ rối toàn diện cho bạn!\" "
            "Audio & SFX: Magical harp glissando, unravelling whoosh, uplifting chime."
        ),
        "image_prompt": (
            "Cute female stick figure girl with ponytail holding a magical glowing pointer wand, untangling a messy scribble yarn ball into neat straight glowing lines beside two floating official scrolls labeled '22082' and '1913', cheerful reassuring smile, clean minimalist doodle art --ar 16:9"
        )
    },
    # Scene 6
    {
        "stt": 6,
        "title_pc": "Scene 06",
        "title_anh": "[Người que] Mô Hình Tháp 3 Tầng Pháp Lý",
        "shot_type": "Medium Shot",
        "dialogue": "Chìa khóa vàng nằm ở việc nắm vững Tháp 3 tầng quản lý pháp lý mà hầu hết mọi người hay nhầm lẫn!",
        "duration": 10,
        "note": "Cô bé người que giơ cao chiếc chìa khóa vàng khổng lồ bên cạnh mô hình Tháp 3 tầng rực rỡ sắc màu",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy stands proudly beside a colorful 3-tiered isometric step pyramid. "
            "She holds up an oversized sparkling golden key above her head with pride. Tier 1 glows blue (Cửa khẩu), Tier 2 glows green (Tự công bố), and Tier 3 glows orange (QCVN). "
            "Animated sparkles and stars erupt around the key as she points to each tier. Camera dynamic circular sweep around the scene. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chìa khóa vàng nằm ở việc nắm vững Tháp 3 tầng quản lý pháp lý mà hầu hết mọi người hay nhầm lẫn!\" "
            "Audio & SFX: Golden key sparkle sound, 3 ascending musical synth notes, triumphant chime."
        ),
        "image_prompt": (
            "Charming 2D doodle art of cute female stick figure with ponytail proudly holding a glowing oversized golden key beside a colorful 3-tiered step pyramid labeled Tier 1, Tier 2, Tier 3, floating cartoon sparkles and stars, clean infographic animation style --ar 16:9"
        )
    },
    # Scene 7
    {
        "stt": 7,
        "title_pc": "Scene 07",
        "title_anh": "[Người que] Tầng 1 - Cửa Khẩu & Thông Tư 15",
        "shot_type": "Medium Shot",
        "dialogue": "Tầng 1 là Thủ tục Kiểm tra nhà nước tại Cửa khẩu. Chỉ những mặt hàng có mã HS trong Thông tư 15 mới bị kiểm tra!",
        "duration": 10,
        "note": "Cô bé người que đứng tại cổng hải quan, lật cuốn sổ Thông tư 15 và chiếu đèn laser xanh vào các mã HS quy định",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy stands beside a cartoon customs border gate. "
            "She holds a neat guidebook titled 'THÔNG TƯ 15/2024' with her stick hands, flipping to a page where a glowing blue laser pointer highlights specific HS code boxes. "
            "Packages with matching codes move toward a scanner, while non-matching ones cruise right through. Camera medium shot with gentle pan. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 1 là Thủ tục Kiểm tra nhà nước tại Cửa khẩu. Chỉ những mặt hàng có mã HS trong Thông tư 15 mới bị kiểm tra!\" "
            "Audio & SFX: Book page flip sound, high-tech scanner beep, confident affirmative chime."
        ),
        "image_prompt": (
            "Cute female stick figure girl with ponytail standing at a miniature cartoon customs barrier gate, holding and pointing at an open regulatory booklet 'Thông tư 15', high-tech friendly scanner beam effect, clean minimalist line art with yellow and blue highlights --ar 16:9"
        )
    },
    # Scene 8
    {
        "stt": 8,
        "title_pc": "Scene 08",
        "title_anh": "[Người que] Cảnh Báo Nhựa PP/PE & Núm Ti Trẻ Em",
        "shot_type": "Medium Shot",
        "dialogue": "Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!",
        "duration": 10,
        "note": "Cô bé người que chỉ tay cảnh báo hai bên: hộp nhựa PP/PE và bình sữa núm ti trẻ em có đèn còi nhấp nháy đỏ",
        "visual_prompt": (
            "2D stick figure animation: Cute stick figure Thuy stands center, wagging her finger with a serious yet cute warning expression. "
            "On her left, animated cartoon plastic food boxes stamped 'PP-05' and 'PET' bounce into view. On her right, an adorable baby bottle with rubber pacifier stamped 'HS 4014.90.10' appears. "
            "Red flashing police sirens and warning triangles pop up above them with bold cartoon text: 'BẮT BUỘC KIỂM TRA CỬA KHẨU!'. Camera zoom-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Đó là Bao bì nhựa như PP, PE, PET và Núm vú cao su trẻ em mã HS 4014.90.10. Thiếu đăng ký kiểm tra là dừng thông quan!\" "
            "Audio & SFX: Double cartoon warning buzzer, siren whoop, barcode scanner beep."
        ),
        "image_prompt": (
            "Expressive 2D doodle illustration: cute female stick figure girl gesturing caution towards floating cute icons of plastic food boxes (PP/PE) and a baby pacifier bottle, glowing red warning sirens and exclamation marks above, bold clean line art --ar 16:9"
        )
    },
    # Scene 9
    {
        "stt": 9,
        "title_pc": "Scene 09",
        "title_anh": "[Người que] Danh Mục Miễn Cửa Khẩu: Inox, Sứ, Thủy Tinh",
        "shot_type": "Wide Shot",
        "dialogue": "Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!",
        "duration": 10,
        "note": "Cô bé người que vui vẻ giơ 2 ngón tay cái, các đồ vật inox, gốm sứ, thủy tinh mọc cánh bay vèo qua cổng hải quan",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy beams with a joyful smile, flashing two thumbs up. "
            "Around her, cute cartoon icons of a shiny stainless steel pot, a ceramic soup bowl, a clear glass jug, and a silicone baking spatula sprout tiny angelic wings. "
            "Bright green checkmarks 'MIỄN KIỂM TRA' pop above them as they flutter happily straight through the customs gate barrier! Camera smooth panning shot. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Ngược lại, Inox, Gốm sứ, Thủy tinh và Cao su người lớn được miễn kiểm tra cửa khẩu, hải quan cho thông quan ngay!\" "
            "Audio & SFX: Angelic choir glissando, 4 consecutive happy pings, upbeat acoustic strum."
        ),
        "image_prompt": (
            "Joyful 2D doodle scene: cute female stick figure girl giving two thumbs up, surrounded by floating friendly cartoon items: steel pot, ceramic bowl, glass bottle, with bright green checkmark badges flying happily past a barrier gate, vibrant uplifting minimalist art --ar 16:9"
        )
    },
    # Scene 10
    {
        "stt": 10,
        "title_pc": "Scene 10",
        "title_anh": "[Người que] Tầng 2 - Thủ Tục Tự Công Bố (Nghị Định 37)",
        "shot_type": "Medium Shot",
        "dialogue": "Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!",
        "duration": 10,
        "note": "Cô bé người que ngồi gõ laptop, đóng dấu xanh 'TỰ CÔNG BỐ' và gấp thành máy bay giấy phóng về tòa nhà Ban ATTP",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy sits at a cute doodle desk with a laptop, typing rapidly with animated smoke puffs on keys. "
            "She takes a file folder stamped 'TỰ CÔNG BỐ - NĐ 37', folds it into an origami paper airplane, and tosses it gracefully across screen. "
            "The paper airplane glides smoothly and lands right on the desk of a cute cartoon building labeled 'Ban Quản Lý An Toàn Thực Phẩm'. Camera tracking motion. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 2 là Thủ tục Tự công bố theo Nghị định 37. Hồ sơ này nộp về Ban Quản lý An toàn thực phẩm để lưu hành hợp pháp!\" "
            "Audio & SFX: Fast cartoon typing sound, paper airplane swoosh, satisfying soft landing bell."
        ),
        "image_prompt": (
            "Cute female stick figure girl with ponytail sitting at a minimalist cartoon desk, stamping a folder labeled 'Tự Công Bố', paper airplane folding and flying towards a neat city hall building doodle, clean vector doodle art style --ar 16:9"
        )
    },
    # Scene 11
    {
        "stt": 11,
        "title_pc": "Scene 11",
        "title_anh": "[Người que] Tầng 3 - Lá Chắn Tiêu Chuẩn QCVN",
        "shot_type": "Medium Shot (Hero Stance)",
        "dialogue": "Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!",
        "duration": 10,
        "note": "Cô bé người que đeo kính bảo hộ phòng lab, giơ cao chiếc khiên phát sáng QCVN bảo vệ sản phẩm khỏi tên bắn hậu kiểm",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy wears cute oversized round lab goggles and a tiny white lab coat. "
            "She holds up a bubbling cartoon test tube, which transforms into a gigantic glowing emerald-green energy shield emblazoned with 'QCVN 12-1'. "
            "A series of cartoon arrows labeled 'HẬU KIỂM' bounce harmlessly off the shield like rubber balls. She strikes a confident superhero pose. Camera heroic low angle push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Tầng 3 là Nghĩa vụ QCVN. Mọi sản phẩm đưa ra thị trường đều bắt buộc đạt chuẩn an toàn thôi nhiễm để phòng thủ hậu kiểm!\" "
            "Audio & SFX: Sci-fi energy shield activation hum, comic ricochet boing sounds, triumphant bass swell."
        ),
        "image_prompt": (
            "Cute female stick figure girl wearing laboratory safety goggles and lab coat, holding a bubbling test tube, standing bravely behind a giant translucent glowing defense shield labeled 'QCVN', clean comic doodle art style --ar 16:9"
        )
    },
    # Scene 12
    {
        "stt": 12,
        "title_pc": "Scene 12",
        "title_anh": "[Người que] Cạm Bẫy Phân Loại: Cao Su Trẻ Em vs Silicone",
        "shot_type": "Close-up",
        "dialogue": "Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!",
        "duration": 10,
        "note": "Cô bé người que rón rén ghé sát ống kính làm điệu bộ bí mật 'suỵt', so sánh núm ti em bé và khuôn nướng bánh silicone",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy tiptoes up close to camera lens, putting a stick finger over her lips with a comical hushed 'shhh' expression. "
            "The background dims with a dramatic cartoon red spotlight focusing on her. She pulls out a baby pacifier in one hand and a silicone cake mold in the other, with an animated neon trap sign flashing '80% MẮC BẪY!'. "
            "Her eyes dart side to side conspiratorially. Camera slow dramatic push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bây giờ là điểm thắt cổ chai khiến 80% doanh nghiệp mất tiền: Cạm bẫy phân loại giữa Cao su trẻ em và Silicone người lớn!\" "
            "Audio & SFX: Cartoon tiptoe woodblock sounds, mystery suspense whoosh, dramatic stinger chord."
        ),
        "image_prompt": (
            "Dramatic close-up 2D doodle: cute female stick figure girl with ponytail whispering a secret with finger on lips, holding a baby pacifier in left hand and silicone cake mold in right hand, theatrical spotlight effect, bold minimalist line art --ar 16:9"
        )
    },
    # Scene 13
    {
        "stt": 13,
        "title_pc": "Scene 13",
        "title_anh": "[Người que] Bóc Tách Nghịch Lý Khuôn Silicone Người Lớn",
        "shot_type": "Medium Shot",
        "dialogue": "Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!",
        "duration": 10,
        "note": "Cô bé người que kéo giãn chiếc khuôn bánh silicone dẻo quẹo, gãi đầu bối rối giữa 2 cuốn sách luật rồi sáng bừng mắt",
        "visual_prompt": (
            "2D stick figure animation: Cute stick figure Thuy holds a flexible cartoon orange silicone cake mold, stretching it comically like elastic rubber while scratching her head in funny confusion. "
            "On her left: 'Thông tư 15' shows a green light (Miễn cửa khẩu); on her right: 'Thông tư 27 Mục 6' flashes an orange light (Bắt buộc Tự công bố). "
            "A lightbulb suddenly pops over her ponytail with a bright 'DING!' as she nods with understanding. Camera split tracking pan. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Khuôn silicone người lớn được miễn cửa khẩu thông quan ngay, nhưng Mục 6 Thông tư 27 vẫn bắt buộc làm Tự công bố!\" "
            "Audio & SFX: Rubber stretching boing sound, comic head scratch, bright lightbulb ding."
        ),
        "image_prompt": (
            "Humorous 2D doodle illustration: cute female stick figure girl stretching a cartoon orange silicone baking mold, standing between two giant cartoon legal books with contradictory arrows, lightbulb turning on above her ponytail, clean line art --ar 16:9"
        )
    },
    # Scene 14
    {
        "stt": 14,
        "title_pc": "Scene 14",
        "title_anh": "[Người que] Nhóm Rủi Ro Thấp: Nồi Inox, Bát Sứ, Ly Thủy Tinh",
        "shot_type": "Wide Shot",
        "dialogue": "Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!",
        "duration": 10,
        "note": "Cô bé người que tung hứng vui vẻ chiếc nồi inox, đĩa sứ và ly thủy tinh, tung tăng không cần nộp hồ sơ lên Sở",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy happily juggles cartoon icons of a gleaming stainless pot, a fine ceramic plate, and a sparkling glass cup like an expert circus performer! "
            "A giant green ribbon banners across top screen: 'NHÓM 1: RỦI RO THẤP - MIỄN NỘP HỒ SƠ LÊN SỞ'. "
            "A cartoon paper folder labeled 'Hồ sơ Sở' flies away into the clouds with little wings. She smiles playfully at the camera. Camera wide tracking shot. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Trong khi đó, Nồi Inox, Bát gốm sứ và Ly thủy tinh thuộc nhóm rủi ro thấp. Bạn không phải nộp hồ sơ Tự công bố lên Sở!\" "
            "Audio & SFX: Juggling swooshes, glass clink sound, cheerful cartoon whistle melody."
        ),
        "image_prompt": (
            "Playful 2D cartoon doodle: cute female stick figure girl juggling shiny metal pot, ceramic plate, and crystal glass tumbler, green celebratory ribbons and low-risk badge in air, happy carefree expression, clean minimalist vector art --ar 16:9"
        )
    },
    # Scene 15
    {
        "stt": 15,
        "title_pc": "Scene 15",
        "title_anh": "[Người que] Khóa An Toàn Hồ Sơ TCCS Nội Bộ",
        "shot_type": "Medium Close-up",
        "dialogue": "Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!",
        "duration": 10,
        "note": "Cô bé người que cất tập hồ sơ TCCS và phiếu test vào chiếc két sắt mini hoạt hình, cất chìa khóa vào túi nháy mắt",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy neatly slides a blue folder labeled 'HỒ SƠ TCCS NỘI BỘ' and a stamped lab test sheet into a cute cartoon mini-safe. "
            "She turns the heavy brass dial with a satisfying 'CLICK-CLACK', locks it shut, and pockets the little brass key with a confident wink at the camera. "
            "An animated cartoon inspector stick figure with a clipboard gives a nod of approval in background. Camera medium close-up. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Doanh nghiệp chỉ cần tự công bố Tiêu chuẩn cơ sở nội bộ, lưu phiếu test QCVN để xuất trình khi Quản lý thị trường kiểm tra!\" "
            "Audio & SFX: Heavy safe tumbler clicks, lock latch clank, cheerful comic wink sparkle."
        ),
        "image_prompt": (
            "Cute female stick figure girl with ponytail locking a neat blue folder labeled 'HỒ SƠ TCCS' into a cute cartoon mini-safe with a brass key, confident winking face, clean lines, minimalist corporate doodle art --ar 16:9"
        )
    },
    # Scene 16
    {
        "stt": 16,
        "title_pc": "Scene 16",
        "title_anh": "[Người que] Tiết Kiệm Hàng Triệu Đồng Chi Phí",
        "shot_type": "Medium Shot",
        "dialogue": "Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!",
        "duration": 10,
        "note": "Cô bé người que ôm chú heo đất vàng khổng lồ, tiền xu vàng rơi rinh rích 'Tiết kiệm 1.000.000đ/model'",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy happily hugs an oversized smiling golden cartoon piggy bank. "
            "Shiny gold coins and green bills rain down joyfully like confetti around her, with digital coin counters rolling up: '+1.000.000 VNĐ / Model!'. "
            "Her cartoon eyes turn into sparkling dollar-stars with pure joy. Camera slow push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Hiểu rõ điểm này, bạn tiết kiệm ngay hàng triệu đồng chi phí làm thủ tục không cần thiết cho hàng Inox và Thủy tinh!\" "
            "Audio & SFX: Cash register cha-ching, cheerful coin jingling rain, triumphant brass fanfares."
        ),
        "image_prompt": (
            "Cheerful 2D doodle illustration: cute female stick figure girl happily hugging a giant golden piggy bank, cartoon green money bills floating down gently, sparkling gold stars, uplifting clean vector art --ar 16:9"
        )
    },
    # Scene 17
    {
        "stt": 17,
        "title_pc": "Scene 17",
        "title_anh": "[Người que] Điểm Huyệt: Dụng Cụ Chứa Đựng vs Tiếp Xúc",
        "shot_type": "Medium Shot",
        "dialogue": "Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!",
        "duration": 10,
        "note": "Cô bé người que đeo kính tri thức, cầm thước chỉ lên bảng đen phân biệt 2 nhóm dụng cụ chứa đựng và không chứa đựng",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy wears cute round black glasses and holds a wooden teacher's ruler, tapping an animated blackboard. "
            "The blackboard divides in two: Left box shows a closed plastic lunchbox holding hot soup ('CHỨA ĐỰNG'); Right box shows a fork, spoon, and chopsticks touching a bowl but holding nothing ('TIẾP XÚC KHÔNG CHỨA ĐỰNG'). "
            "She points sharply to the difference with an intelligent nod. Camera smooth tracking dolly. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Điểm huyệt thứ hai là phân biệt giữa Dụng cụ CHỨA ĐỰNG tiếp xúc trực tiếp và Dụng cụ TIẾP XÚC nhưng KHÔNG CHỨA ĐỰNG!\" "
            "Audio & SFX: Chalk writing sound on blackboard, ruler tapping clack, subtle mystery tone."
        ),
        "image_prompt": (
            "Cute female stick figure girl wearing round glasses pointing with a pointer at a cute blackboard comparison: sealed storage box vs spoons and chopsticks, educational whiteboard doodle animation aesthetic, sharp lines --ar 16:9"
        )
    },
    # Scene 18
    {
        "stt": 18,
        "title_pc": "Scene 18",
        "title_anh": "[Người que] Kim Bài Công Văn 4372 Bộ Y Tế (Aeon)",
        "shot_type": "Close-up",
        "dialogue": "Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!",
        "duration": 10,
        "note": "Cô bé người que mở cuộn thư hoàng gia phát sáng Công văn 4372 của Bộ Y tế với con dấu đỏ chót uy quyền",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy rolls out a majestic glowing cartoon parchment scroll labeled 'CÔNG VĂN 4372/BYT-VPB6 - BỘ Y TẾ'. "
            "A warm dramatic golden spotlight beam shines down directly on the official red emblem stamp and the text 'Trả lời Công ty Aeon Việt Nam'. "
            "She smiles proudly, gesturing both hands toward the legal document like presenting a treasure. Camera slow zoom-in on the glowing red seal. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Căn cứ then chốt nằm tại Công văn 4372 năm 2019 của Bộ Y tế trả lời chính thức cho công ty Aeon Việt Nam!\" "
            "Audio & SFX: Royal trumpet fanfare fanfare, scroll unrolling whoosh, spotlight turn-on click."
        ),
        "image_prompt": (
            "Dramatic 2D doodle scene: cute female stick figure girl with ponytail holding open a giant illuminated legal dispatch scroll labeled 'Công văn 4372 BYT', bright red official seal glowing under a golden spotlight beam, clean vector line art --ar 16:9"
        )
    },
    # Scene 19
    {
        "stt": 19,
        "title_pc": "Scene 19",
        "title_anh": "[Người que] Danh Sách Thìa, Đũa, Dao, Găng Tay Cơ Học",
        "shot_type": "Wide Shot",
        "dialogue": "Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!",
        "duration": 10,
        "note": "Các đồ dùng thìa, đũa, dĩa, dao gọt, kẹp gắp hoạt hình nhảy múa quanh cô bé người que, một con dấu đỏ đóng sầm 'KHÔNG PHẢI CHỨA ĐỰNG'",
        "visual_prompt": (
            "2D stick figure animation: Dynamic comic collage: Cute stick figure Thuy stands center as cartoon kitchen utensils dance merrily around her: "
            "wooden chopsticks doing a tap dance, soup spoon bouncing, fork twirling, fruit peeler bowing, cooking gloves waving, and salad tongs snapping playfully. "
            "A massive comic rubber stamp slams right across the screen: 'KHÔNG PHẢI DỤNG CỤ CHỨA ĐỰNG!'. Camera wide dynamic frame. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bộ Y tế nêu rõ: Đũa, Thìa, Dĩa, Dao gọt hoa quả, Găng tay nấu ăn và Kẹp gắp KHÔNG PHẢI dụng cụ chứa đựng!\" "
            "Audio & SFX: Playful cartoon kitchen clinking sounds, utensil boings, heavy rubber stamp thud."
        ),
        "image_prompt": (
            "Dynamic 2D doodle collage: cute female stick figure girl surrounded by dancing cartoon kitchen utensils: chopsticks, spoon, fork, peeler knife, tongs, cooking glove, with bold red banner 'KHÔNG PHẢI DỤNG CỤ CHỨA ĐỰNG', clean vector art --ar 16:9"
        )
    },
    # Scene 20
    {
        "stt": 20,
        "title_pc": "Scene 20",
        "title_anh": "[Người que] Khẳng Định Miễn 100% TCB & Kiểm Tra Cửa Khẩu",
        "shot_type": "Medium Shot",
        "dialogue": "Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!",
        "duration": 10,
        "note": "Cô bé người que vung 2 tay dứt khoát gạt bỏ mọi rào cản, hai dấu X đỏ biến thành hai dấu tích xanh khổng lồ 'MIỄN 100%'",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy makes a powerful, decisive double outward waving gesture with both stick hands. "
            "Two barrier signs labeled 'TỰ CÔNG BỐ' and 'KIỂM TRA CỬA KHẨU' get stamped with giant red 'X' marks, which instantly dissolve into two sparkling emerald green checkmarks labeled 'MIỄN 100%!'. "
            "She smiles victoriously, standing tall with hands on hips. Camera medium shot with gentle push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Mặt hàng này HOÀN TOÀN MIỄN TỰ CÔNG BỐ, Sở từ chối tiếp nhận, và MIỄN LUÔN kiểm tra chuyên ngành tại cửa khẩu!\" "
            "Audio & SFX: Powerful double swoosh sound, magic transformation sparkle, triumphant chord."
        ),
        "image_prompt": (
            "Empowering 2D doodle illustration: cute female stick figure girl making decisive outward waving gesture, two giant red crossing out marks turning into glowing green checkmarks, victorious beaming smile, clean minimalist lines --ar 16:9"
        )
    },
    # Scene 21
    {
        "stt": 21,
        "title_pc": "Scene 21",
        "title_anh": "[Người que] Xe Container Luồng Xanh Bon Bon Về Kho",
        "shot_type": "Wide Shot",
        "dialogue": "Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!",
        "duration": 10,
        "note": "Cô bé người que ngồi trên nóc cabin xe container hoạt hình, vẫy tay tươi cười khi xe phóng qua barie luồng xanh trong ngày",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy happily perches on top of a cute blue cartoon container truck. "
            "She waves her hand cheerfully at viewers as the truck speeds effortlessly past an open green customs barrier labeled 'LUỒNG XANH - THÔNG QUAN 1 NGÀY'. "
            "A happy cartoon sun shines above, with zero fee badges ($0) fluttering in the wind. Camera fast tracking vehicle moving right. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Bạn nhập khẩu chúng như hàng hóa thông thường, không mất một đồng phí kiểm tra chuyên ngành và thông quan trong ngày!\" "
            "Audio & SFX: Cheerful truck horn 'beep-beep', breezy wind sound, uplifting sunny melody."
        ),
        "image_prompt": (
            "Cheerful 2D vector doodle: cute female stick figure girl waving happily from the window of a cute cartoon container truck speeding past an open green customs barrier, bright sun and speed lines, uplifting clean art --ar 16:9"
        )
    },
    # Scene 22
    {
        "stt": 22,
        "title_pc": "Scene 22",
        "title_anh": "[Người que] Cảnh Báo Mốc Thời Gian Lab 10-12 Ngày",
        "shot_type": "Medium Shot",
        "dialogue": "Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!",
        "duration": 10,
        "note": "Cô bé người que gõ gõ vào đồng hồ đeo tay, bên cạnh là tờ lịch bóc liên tục từ Ngày 1 đến Ngày 12 toát mồ hôi hột",
        "visual_prompt": (
            "2D stick figure animation: Cute stick figure Thuy taps her tiny cartoon wrist watch with a nervous, urgent expression. "
            "Beside her, a giant cartoon wall calendar flips frantically: 'Ngày 1... Ngày 5... Ngày 10... Ngày 12!'. "
            "A cartoon clock with spinning hands ticks loudly overhead, while giant comical sweat drops fly from her ponytail. Camera fast zoom on calendar. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chiến lược then chốt tiếp theo là kiểm soát thời gian. Thử nghiệm tại phòng lab mất từ 10 đến 12 ngày làm việc!\" "
            "Audio & SFX: Fast rapid clock ticking, flapping calendar pages sound, nervous gulp."
        ),
        "image_prompt": (
            "Humorous 2D doodle art: cute female stick figure girl with ponytail tapping an oversized wristwatch, looking anxiously at a giant cartoon calendar flipping pages rapidly to 'Day 12', spinning clock hands, expressive comic line art --ar 16:9"
        )
    },
    # Scene 23
    {
        "stt": 23,
        "title_pc": "Scene 23",
        "title_anh": "[Người que] Cơn Ác Mộng Tiền Phạt Dem/Det 15-30 Triệu",
        "shot_type": "Wide Shot (Dramatic)",
        "dialogue": "Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!",
        "duration": 10,
        "note": "Cô bé người que ôm đầu kinh hãi khi cơn mưa hóa đơn phạt đỏ lòm 30 triệu trút xuống, tàu hàng kẹt khói đen ở cảng",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy holds her head in cartoon horror with wavy spaghetti arms! "
            "A dark thunderstorm cloud rumbles overhead, raining heavy red penalty invoice sheets stamped with fiery text: 'PHẠT DEM/DET: 30.000.000 VNĐ!'. "
            "In background, a sad cartoon cargo ship sits stuck at port blowing puffs of dark smoke. Camera dramatic high-angle looking down. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Nếu chờ tàu cập cảng mới lấy mẫu gửi test, tiền phạt lưu cont lưu bãi sẽ lên tới 15 đến 30 triệu mỗi container!\" "
            "Audio & SFX: Distant cartoon thunder, loud invoice thud, dramatic low brass drone."
        ),
        "image_prompt": (
            "Dramatic comic doodle scene: cute female stick figure girl holding her head in horror under a rain of red penalty fee papers labeled '30 TRIỆU DEM/DET', cartoon stranded ship in background, expressive manga-style stick figure --ar 16:9"
        )
    },
    # Scene 24
    {
        "stt": 24,
        "title_pc": "Scene 24",
        "title_anh": "[Người que] Tuyệt Chiêu Gửi Mẫu CPN Trước 25 Ngày",
        "shot_type": "Medium Shot",
        "dialogue": "Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!",
        "duration": 10,
        "note": "Cô bé người que sáng bừng mắt với bóng đèn ý tưởng, phóng gói hàng CPN gắn cánh tên lửa bay vèo về trước 25 ngày",
        "visual_prompt": (
            "2D stick figure animation: An animated lightbulb pops over cute stick figure Thuy's ponytail with a bright 'DING!'. "
            "Her eyes light up with enthusiasm as she holds up a cute yellow air express parcel (CPN) equipped with tiny rocket thrusters and angel wings! "
            "A calendar rolls back 25 days before ship departure. She releases the package, and it zooms into the sky like a supersonic jet straight to the lab. She winks with a thumbs-up. Camera smooth tilt-up. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Kinh nghiệm thực chiến của Thúy: Hãy gửi chuyển phát nhanh 2 đến 3 mẫu từ xưởng về trước ngày tàu chạy 25 ngày!\" "
            "Audio & SFX: Idea ding, rocket thruster whoosh, joyful cheerful whistle."
        ),
        "image_prompt": (
            "Inspiring 2D doodle illustration: cute female stick figure girl with ponytail holding a flying winged express courier parcel with rocket boost, a calendar showing '-25 Days', confident wink and thumbs up, clean vector lines --ar 16:9"
        )
    },
    # Scene 25
    {
        "stt": 25,
        "title_pc": "Scene 25",
        "title_anh": "[Người que] Màn Hình Cổng Một Cửa NSW Duyệt Trong 3 Ngày",
        "shot_type": "Medium Shot",
        "dialogue": "Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!",
        "duration": 10,
        "note": "Cô bé người que nhảy cẫng ăn mừng trước màn hình máy tính Một Cửa Quốc Gia có con dấu xanh 'ĐẠT YÊU CẦU TRONG 3 NGÀY'",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy jumps in the air with both hands raised in victory! "
            "In front of her, an animated giant computer monitor shows the Vietnam National Single Window (NSW) portal. A huge glowing green digital seal stamps down: 'KẾT QUẢ: ĐẠT YÊU CẦU - THÔNG QUAN 3 NGÀY!'. "
            "Party poppers explode with colorful confetti and streamers all around her. Camera medium shot. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Khi tàu cập cảng, Tự công bố đã sẵn sàng. Bạn đăng ký kiểm tra nhà nước và thông quan chỉ mất đúng 3 ngày!\" "
            "Audio & SFX: Digital stamp chime, party popper pop sound, triumphant victory fanfare."
        ),
        "image_prompt": (
            "Triumphant 2D cartoon scene: cute female stick figure girl high-fiving the air in front of a giant computer screen with big green 'APPROVED - 3 DAYS' stamp, colorful confetti and party horns, clean doodle vector art --ar 16:9"
        )
    },
    # Scene 26
    {
        "stt": 26,
        "title_pc": "Scene 26",
        "title_anh": "[Người que] Giải Pháp Đưa Hàng Về Kho Riêng Bảo Quản",
        "shot_type": "Wide Shot",
        "dialogue": "Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!",
        "duration": 10,
        "note": "Cô bé người que đội mũ bảo hộ tí hon, cầm gậy chỉ huy điều phối xe cont lùi vào kho riêng, cầm kéo cắt phéng hóa đơn phạt",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy wears a cute yellow construction hardhat, blowing a whistle and waving orange marshaling batons like a cheerful airport signalman. "
            "A cartoon container truck smoothly backs into a spacious private warehouse labeled 'KHO BẢO QUẢN RIÊNG'. "
            "The warehouse roller door shuts securely with a padlock, and a giant cartoon scissors cuts a DEMURRAGE invoice in half! Camera wide tracking shot. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Trường hợp gấp, giải pháp là làm thủ tục xin đưa hàng về kho riêng bảo quản, triệt tiêu hoàn toàn phí lưu bãi!\" "
            "Audio & SFX: Whistle tweet, truck reverse beeping, snipping scissors snip, warehouse shutter roll."
        ),
        "image_prompt": (
            "Fun 2D doodle scene: cute female stick figure girl wearing tiny safety hardhat guiding a cartoon container truck into a private warehouse door, giant scissors cutting a fee receipt, clean vector line art --ar 16:9"
        )
    },
    # Scene 27
    {
        "stt": 27,
        "title_pc": "Scene 27",
        "title_anh": "[Người que] Bản Tự Công Bố Vô Thời Hạn (Làm 1 Lần Dùng Mãi)",
        "shot_type": "Medium Shot",
        "dialogue": "Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!",
        "duration": 10,
        "note": "Cô bé người que ôm cuốn sổ Tự công bố có biểu tượng vô cực phát sáng, những chuyến hàng tiếp theo bay qua với phí 0 đồng",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy hugs a golden glowing dossier folder with a sparkling infinity symbol '∞' on the cover. "
            "Floating glowing badges labeled '0 ĐỒNG PHÍ' and 'DÙNG MÃI MÃI' dance around her head like friendly fireflies. "
            "A conveyor belt of cartoon cargo boxes flows smoothly into the distance without paying a single penny. She radiates pure satisfaction. Camera slow push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Đặc biệt: Bản Tự công bố có hiệu lực vô thời hạn cho từng model. Từ lô thứ hai, bạn không tốn thêm một đồng nào!\" "
            "Audio & SFX: Magical infinity chime, gentle harp strum, warm corporate synth pad."
        ),
        "image_prompt": (
            "Heartwarming 2D doodle: cute female stick figure girl hugging an ornate glowing document binder with a sparkling infinity symbol '∞' on cover, '$0' fee badges floating with stars, joyful proud smile, clean vector art --ar 16:9"
        )
    },
    # Scene 28
    {
        "stt": 28,
        "title_pc": "Scene 28",
        "title_anh": "[Người que] Nhắc Nhở Hạn Phiếu Kiểm Nghiệm 12 Tháng",
        "shot_type": "Medium Close-up",
        "dialogue": "Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!",
        "duration": 10,
        "note": "Cô bé người que cầm đồng hồ cát chảy nhẹ nhàng và chỉ tay nhắc nhở hạn 12 tháng trên phiếu kiểm nghiệm",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy holds a delicate cartoon glass hourglass with trickling blue sand in one hand, "
            "and points her other stick finger at a laboratory test certificate stamped 'HẠN SỬ DỤNG 12 THÁNG'. "
            "A gentle cartoon ringing alarm clock wobbles on a nearby desk with soft warning pulses. She reminds the audience with a kind, caring mentor smile. Camera medium close-up. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Chỉ cần lưu ý: Phiếu kiểm nghiệm có giá trị đúng 12 tháng khi nộp hồ sơ. Đừng để quá hạn kẻo phải làm lại từ đầu!\" "
            "Audio & SFX: Soft sand trickling sound, gentle alarm bell chime, friendly advisory tone."
        ),
        "image_prompt": (
            "Thoughtful 2D doodle illustration: cute female stick figure girl with ponytail holding an hourglass and pointing at a test report labeled 'Valid 12 Months', gentle reminder expression, minimalist clean lines, yellow accent --ar 16:9"
        )
    },
    # Scene 29
    {
        "stt": 29,
        "title_pc": "Scene 29",
        "title_anh": "[Người que] Triết Lý Hàng Rào Pháp Lý Trước Khi Rời Xưởng",
        "shot_type": "Wide Shot (Heroic)",
        "dialogue": "Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!",
        "duration": 10,
        "note": "Cô bé người que đứng oai phong khoanh tay, sau lưng là bức tường thành vững chãi bảo vệ đoàn xe container khỏi bão tố",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy stands tall and proud in center frame, wearing a cute captain's hat and a flowing cape, arms folded heroically with determined eyes. "
            "Behind her, an unshakeable cartoon brick fortress wall labeled 'HÀNG RÀO PHÁP LÝ' stands strong, shielding a fleet of cheerful container trucks from stormy clouds and lightning in the distance. "
            "Camera heroic slow push-in. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Logistics thực chiến không phải chờ sự cố rồi chữa cháy, mà là xây hàng rào pháp lý trước khi container rời xưởng!\" "
            "Audio & SFX: Inspiring cinematic piano chord progression, heroic orchestral swell, gentle thunder rumble far away."
        ),
        "image_prompt": (
            "Inspiring heroic 2D doodle art: cute female stick figure girl standing proudly with hands on hips, standing in front of a protective castle wall shielding cargo containers from storm clouds, strong leadership presence, clean lines --ar 16:9"
        )
    },
    # Scene 30
    {
        "stt": 30,
        "title_pc": "Scene 30",
        "title_anh": "[Người que] Kêu Gọi Subscribe & Tặng Cẩm Nang",
        "shot_type": "Medium Shot",
        "dialogue": "Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!",
        "duration": 10,
        "note": "Cô bé người que vẫy tay chào tạm biệt với má hồng dễ thương, cầm cẩm nang quà tặng và nút Subscribe rung rinh",
        "visual_prompt": (
            "2D stick figure animation: Cute female stick figure Thuy waves both hands adorably at the camera with blushing pink cheeks and a beaming radiant smile. "
            "Beside her, a big red cartoon 'ĐĂNG KÝ KÊNH' button and a vibrating golden notification bell wobble cutely. "
            "She holds out a lovely gift-wrapped e-book titled 'CẨM NANG NHẬP KHẨU BAO BÌ THỰC PHẨM' tied with a red ribbon. Cute floating hearts rise into the air. Camera medium shot. "
            "Ms. Thuy's Voiceover (Vietnamese, confident female expert, pacing 10s): \"Hãy bấm Đăng ký kênh và bình luận BAO BÌ, Thúy sẽ gửi tặng bạn Cẩm nang nghiệp vụ nhập khẩu bao bì thực phẩm độc quyền nhé!\" "
            "Audio & SFX: Cheerful YouTube outro theme, bell notification ding, cute popping bubble hearts."
        ),
        "image_prompt": (
            "Adorable 2D doodle outro: cute female stick figure girl with ponytail waving goodbye with blushing cheeks, presenting a gift-wrapped handbook 'CẨM NANG BAO BÌ', cute red Subscribe button and ringing bell doodle, hearts floating, clean vector aesthetic --ar 16:9"
        )
    }
]

def create_excel_and_csv():
    target_pc_xlsx = r"C:\Users\thuyn\Downloads\phân cảnh.xlsx"
    target_anh_xlsx = r"C:\Users\thuyn\Downloads\ảnh.xlsx"
    target_pc_csv = r"C:\Users\thuyn\Downloads\Kich_Ban_Phan_Canh_Nguoi_Que.csv"
    target_anh_csv = r"C:\Users\thuyn\Downloads\Tao_Anh_Hang_Loat_Nguoi_Que.csv"

    # 1. Update/Create phan canh.xlsx
    wb_pc = openpyxl.Workbook()
    ws_pc = wb_pc.active
    ws_pc.title = "KICH_BAN_FLOW"

    # Headers
    pc_headers = ["STT", "Tiêu đề", "Mô tả phân cảnh (Prompt)", "Lời thoại (Dialogue)", "Thời lượng (giây)"]
    ws_pc.append(pc_headers)

    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid") # Dark blue
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for col_idx in range(1, len(pc_headers) + 1):
        cell = ws_pc.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align

    for item in scenes_data:
        ws_pc.append([
            item["stt"],
            item["title_pc"],
            item["visual_prompt"],
            item["dialogue"],
            item["duration"]
        ])

    for row_idx in range(2, len(scenes_data) + 2):
        ws_pc.cell(row=row_idx, column=1).alignment = center_align
        ws_pc.cell(row=row_idx, column=2).alignment = center_align
        ws_pc.cell(row=row_idx, column=3).alignment = left_align
        ws_pc.cell(row=row_idx, column=4).alignment = left_align
        ws_pc.cell(row=row_idx, column=5).alignment = center_align

        for col_idx in range(1, 6):
            c = ws_pc.cell(row=row_idx, column=col_idx)
            c.font = Font(name="Arial", size=10)
            c.border = thin_border

    # Set column widths
    ws_pc.column_dimensions['A'].width = 8
    ws_pc.column_dimensions['B'].width = 15
    ws_pc.column_dimensions['C'].width = 60
    ws_pc.column_dimensions['D'].width = 45
    ws_pc.column_dimensions['E'].width = 18

    # Save to dedicated new file and try target file
    pc_nguoi_que = r"C:\Users\thuyn\Downloads\phân cảnh_người que.xlsx"
    wb_pc.save(pc_nguoi_que)
    print(f"Saved {pc_nguoi_que} successfully!")
    try:
        tmp_pc = target_pc_xlsx + ".tmp.xlsx"
        wb_pc.save(tmp_pc)
        os.replace(tmp_pc, target_pc_xlsx)
        print("Overwritten phan canh.xlsx successfully!")
    except Exception as e:
        print(f"Note: Could not overwrite original {target_pc_xlsx} because it is open in Excel: {e}")

    # 2. Update/Create anh.xlsx
    wb_anh = openpyxl.Workbook()
    ws_anh = wb_anh.active
    ws_anh.title = "TAO_ANH_HANG_LOAT"

    anh_headers = ["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"]
    ws_anh.append(anh_headers)

    anh_header_fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid") # Dark green

    for col_idx in range(1, len(anh_headers) + 1):
        cell = ws_anh.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = anh_header_fill
        cell.alignment = center_align

    for item in scenes_data:
        ws_anh.append([
            item["stt"],
            item["title_anh"],
            item["image_prompt"],
            item["shot_type"],
            item["note"]
        ])

    for row_idx in range(2, len(scenes_data) + 2):
        ws_anh.cell(row=row_idx, column=1).alignment = center_align
        ws_anh.cell(row=row_idx, column=2).alignment = left_align
        ws_anh.cell(row=row_idx, column=3).alignment = left_align
        ws_anh.cell(row=row_idx, column=4).alignment = center_align
        ws_anh.cell(row=row_idx, column=5).alignment = left_align

        for col_idx in range(1, 6):
            c = ws_anh.cell(row=row_idx, column=col_idx)
            c.font = Font(name="Arial", size=10)
            c.border = thin_border

    ws_anh.column_dimensions['A'].width = 8
    ws_anh.column_dimensions['B'].width = 35
    ws_anh.column_dimensions['C'].width = 65
    ws_anh.column_dimensions['D'].width = 22
    ws_anh.column_dimensions['E'].width = 45

    anh_nguoi_que = r"C:\Users\thuyn\Downloads\ảnh_người que.xlsx"
    wb_anh.save(anh_nguoi_que)
    print(f"Saved {anh_nguoi_que} successfully!")
    try:
        tmp_anh = target_anh_xlsx + ".tmp.xlsx"
        wb_anh.save(tmp_anh)
        os.replace(tmp_anh, target_anh_xlsx)
        print("Overwritten anh.xlsx successfully!")
    except Exception as e:
        print(f"Note: Could not overwrite original {target_anh_xlsx} because it is open in Excel: {e}")

    # 2.5 Create Bao_Bi_Thuc_Pham_Nguoi_Que_Studio.xlsx (All-in-One Workbook)
    target_studio = r"C:\Users\thuyn\Downloads\Bao_Bi_Thuc_Pham_Nguoi_Que_Studio.xlsx"
    wb_studio = openpyxl.Workbook()
    
    # Sheet 1: KICH_BAN_FLOW
    ws1 = wb_studio.active
    ws1.title = "KICH_BAN_FLOW"
    ws1.append(pc_headers)
    for col_idx in range(1, len(pc_headers) + 1):
        c = ws1.cell(row=1, column=col_idx)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center_align
    for item in scenes_data:
        ws1.append([item["stt"], item["title_pc"], item["visual_prompt"], item["dialogue"], item["duration"]])
    for row_idx in range(2, len(scenes_data) + 2):
        ws1.cell(row=row_idx, column=1).alignment = center_align
        ws1.cell(row=row_idx, column=2).alignment = center_align
        ws1.cell(row=row_idx, column=3).alignment = left_align
        ws1.cell(row=row_idx, column=4).alignment = left_align
        ws1.cell(row=row_idx, column=5).alignment = center_align
        for col_idx in range(1, 6):
            c = ws1.cell(row=row_idx, column=col_idx)
            c.font = Font(name="Arial", size=10)
            c.border = thin_border
    ws1.column_dimensions['A'].width = 8
    ws1.column_dimensions['B'].width = 15
    ws1.column_dimensions['C'].width = 60
    ws1.column_dimensions['D'].width = 45
    ws1.column_dimensions['E'].width = 18

    # Sheet 2: TAO_ANH_HANG_LOAT
    ws2 = wb_studio.create_sheet(title="TAO_ANH_HANG_LOAT")
    ws2.append(anh_headers)
    for col_idx in range(1, len(anh_headers) + 1):
        c = ws2.cell(row=1, column=col_idx)
        c.font = header_font
        c.fill = anh_header_fill
        c.alignment = center_align
    for item in scenes_data:
        ws2.append([item["stt"], item["title_anh"], item["image_prompt"], item["shot_type"], item["note"]])
    for row_idx in range(2, len(scenes_data) + 2):
        ws2.cell(row=row_idx, column=1).alignment = center_align
        ws2.cell(row=row_idx, column=2).alignment = left_align
        ws2.cell(row=row_idx, column=3).alignment = left_align
        ws2.cell(row=row_idx, column=4).alignment = center_align
        ws2.cell(row=row_idx, column=5).alignment = left_align
        for col_idx in range(1, 6):
            c = ws2.cell(row=row_idx, column=col_idx)
            c.font = Font(name="Arial", size=10)
            c.border = thin_border
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 35
    ws2.column_dimensions['C'].width = 65
    ws2.column_dimensions['D'].width = 22
    ws2.column_dimensions['E'].width = 45

    wb_studio.save(target_studio)
    print(f"Saved {target_studio} successfully!")

    # 3. Export CSV files (UTF-8 BOM for Google Flow upload compatibility)
    with open(target_pc_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(pc_headers)
        for item in scenes_data:
            writer.writerow([
                item["stt"],
                item["title_pc"],
                item["visual_prompt"],
                item["dialogue"],
                item["duration"]
            ])
    print("Saved Kich_Ban_Phan_Canh_Nguoi_Que.csv successfully!")

    with open(target_anh_csv, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(anh_headers)
        for item in scenes_data:
            writer.writerow([
                item["stt"],
                item["title_anh"],
                item["image_prompt"],
                item["shot_type"],
                item["note"]
            ])
    print("Saved Tao_Anh_Hang_Loat_Nguoi_Que.csv successfully!")

if __name__ == '__main__':
    create_excel_and_csv()
