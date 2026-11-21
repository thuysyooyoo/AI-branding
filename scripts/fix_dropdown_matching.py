# -*- coding: utf-8 -*-
"""
Script chuẩn hóa 100% cột 'Góc quay (Shot Type)' để khớp chính xác với Dropdown
trên giao diện web:
- Trung cảnh (Medium Shot)
- Toàn cảnh (Wide Shot)
- Cận cảnh (Close-up)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import csv

scenes = [
    {
        "stt": 1,
        "title": "Câu Đố Hóc Búa Bao Bì Thực Phẩm",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure detective named Thuy with ponytail and tiny bow ribbon, wearing cute round detective glasses, holding a giant cartoon magnifying glass and winking challengingly at camera. Three floating doodle items revolve around her with glowing yellow question marks: a plastic container, a silicone cake mold, and a metal spoon. Clean minimalist white background, comic motion lines, expressive cartoon eyes, crisp black ink line art with selective vibrant yellow and cyan accents, 8k resolution --ar 16:9",
        "note": "Cô bé người que Thúy hóa thân thành thám tử tinh nghịch, đưa ra câu đố hóc búa với 3 món đồ bay lơ lửng kèm dấu hỏi để thu hút người xem ngay 3 giây đầu"
    },
    {
        "stt": 2,
        "title": "Cảnh Báo Kẹt Cảng & Thiệt Hại",
        "shot_vn": "Toàn cảnh (Wide Shot)",
        "shot_en": "Wide Shot",
        "prompt": "Wide shot 2D doodle illustration of busy seaport container terminal: rows of stacked shipping containers, a sharp oversized red cartoon padlock and heavy chains locking the container doors, floating red smoke numbers '100.000.000 VND', dramatic industrial comic sketch aesthetic, clean white paper background, crisp black ink outlines, 8k --ar 16:9",
        "note": "B-roll cảnh báo kẹt cảng container thực tế dạng doodle, niêm phong kẹp chì đỏ và số tiền phạt bốc khói"
    },
    {
        "stt": 3,
        "title": "Cận Cảnh Túi Màng Nhôm Vô Trùng",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Macro close-up shot of sterile food packaging aluminum aseptic bags with blue sanitary spouts, clean industrial product doodle aesthetic, metallic silver foil reflections, an official corporate Self-Declaration document with red seal lying beside, shallow depth of field, high-tech food manufacturing vector art, 8k --ar 16:9",
        "note": "B-roll cận cảnh sản phẩm túi màng nhôm vô trùng Aseptic bag và hồ sơ Tự công bố"
    },
    {
        "stt": 4,
        "title": "Bẫy Ngộ Nhận Thông Tư 27",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up top-down shot of official legal document Circular 27/2026/TT-BYT on wooden desk, yellow fluorescent highlighter emphasizing regulatory paragraph 'Rủi ro trung bình', a large red customs rubber stamp 'STOP' hitting the paper, reading glasses beside, warm studio lighting, sharp focus vector doodle art --ar 16:9",
        "note": "Chụp cận văn bản Thông tư 27/2026/TT-BYT điểm bẫy ngộ nhận pháp lý và dấu đỏ dừng hàng"
    },
    {
        "stt": 5,
        "title": "Khẳng Định Bản Chất Văn Bản 22082",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl named Thuy with adorable ponytail and ribbon, smiling confidently with reassuring open-hand gesture, beside two floating glowing parchment scrolls of Official Dispatch 22082 and 1913 with red seals, neat golden threads untangled, clean minimalist background, 8k --ar 16:9",
        "note": "Cô bé người que Thúy khẳng định văn bản 22082 và 1913, gỡ rối cuộn len thành đường thẳng mở ra giải pháp"
    },
    {
        "stt": 6,
        "title": "Đồ Họa 3D Tháp 3 Tầng Pháp Lý",
        "shot_vn": "Toàn cảnh (Wide Shot)",
        "shot_en": "Wide Shot",
        "prompt": "Wide shot isometric 3D animation style clean infographic visualization of a modern 3-tiered compliance pyramid structure, glowing neon layers labeled Tier 1, Tier 2, Tier 3, floating oversized golden key hovering on top with magical sparkles, clean dark navy and white gradient background, sharp 8k render --ar 16:9",
        "note": "Mô hình 3D trực quan Tháp 3 Tầng Quản Lý Pháp Lý với chìa khóa vàng"
    },
    {
        "stt": 7,
        "title": "Giảng Giải Tầng 1 - Cửa Khẩu Hải Quan",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy with ponytail seated at modern desk, holding a printed regulatory booklet of Circular 15/2024, looking directly at camera with clear informative demeanor, miniature cartoon customs gate in background, clean minimalist line art --ar 16:9",
        "note": "Cô bé người que Thúy phân tích Tầng 1 Cửa khẩu theo Thông tư 15"
    },
    {
        "stt": 8,
        "title": "Cận Cảnh Nhựa PP & Núm Vú Cao Su",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up macro shot showing transparent PP food storage container with recycling code PP-05 beside baby rubber silicone feeding bottle nipple labeled HS 4014.90.10, clean white product background, sharp macro details, red warning border graphic, no photographic noise, clean vector commercial look --ar 16:9",
        "note": "Cận cảnh nhựa tổng hợp PP/PE và núm vú cao su trẻ em"
    },
    {
        "stt": 9,
        "title": "Danh Mục Sản Phẩm Miễn Cửa Khẩu",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot clean product showcase of stainless steel pot, white ceramic plate, glass jar, and blue silicone baking mold arranged neatly on marble countertop, bright green checkmark badges floating above with text 'MIỄN KIỂM TRA CỬA KHẨU', bright softbox illumination, premium commercial doodle look --ar 16:9",
        "note": "B-roll sản phẩm inox, sứ, thủy tinh được miễn kiểm tra cửa khẩu"
    },
    {
        "stt": 10,
        "title": "Giảng Giải Tầng 2 - Thủ Tục Tự Công Bố",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy with ponytail seated in front of laptop in modern office, explaining regulatory compliance with poised professional posture, stamping folder labeled 'TỰ CÔNG BỐ', clean natural lighting, executive doodle backdrop --ar 16:9",
        "note": "Cô bé người que Thúy hướng dẫn Tầng 2 Tự công bố lưu hành thị trường"
    },
    {
        "stt": 11,
        "title": "Phòng Lab Thử Nghiệm Chuẩn QCVN",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot inside certified food contact safety testing laboratory in clean 2D doodle animation style: analytical test tubes and spectrophotometer, a glowing green holographic defense shield labeled 'QCVN 12-1:2011/BYT - PASSED', pristine sterile scientific environment, 8k --ar 16:9",
        "note": "B-roll phòng lab thử nghiệm chỉ tiêu thôi nhiễm QCVN"
    },
    {
        "stt": 12,
        "title": "Cạm Bẫy Cao Su Trẻ Em vs Silicone",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of cute female stick figure girl Thuy with ponytail, leaning forward with engaging, intense storytelling expression, holding baby pacifier in left hand and silicone mold in right hand, animated red warning badge '80% MẮC BẪY', dark background with soft dramatic spotlight, 8k --ar 16:9",
        "note": "Cô bé người que Thúy hạ giọng chỉ ra điểm thắt cổ chai 80% mắc bẫy"
    },
    {
        "stt": 13,
        "title": "Bóc Tách Nghịch Lý Khuôn Silicone Người Lớn",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up split graphic comparison on executive wooden table: official Circular 15 table contrasted with Circular 27 Item 6, colorful silicone cake mold resting on document corner, high resolution legal audit vector scene, clean minimalist line art --ar 16:9",
        "note": "Bóc tách nghịch lý silicone người lớn miễn cửa khẩu nhưng phải TCB"
    },
    {
        "stt": 14,
        "title": "Đồ Gia Dụng Inox & Gốm Sứ Nhóm 1",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot commercial product illustration of polished stainless steel pan, elegant white ceramic bowl, and clear glass tumbler, green badge 'NHÓM 1: RỦI RO THẤP - MIỄN NỘP HỒ SƠ LÊN SỞ', pristine reflections, clean vector product art, 8k resolution --ar 16:9",
        "note": "B-roll nồi inox, bát sứ, ly thủy tinh thuộc nhóm rủi ro thấp"
    },
    {
        "stt": 15,
        "title": "Hướng Dẫn Hồ Sơ TCCS Phòng Thủ Hậu Kiểm",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of cute female stick figure girl Thuy with ponytail holding an organized black corporate filing binder labeled 'HỒ SƠ TCCS NỘI BỘ', confident advisory demeanor, miniature cute mini-safe beside, clean modern doodle setting, 8k --ar 16:9",
        "note": "Cô bé người que Thúy hướng dẫn lưu trữ hồ sơ Tiêu chuẩn cơ sở nội bộ"
    },
    {
        "stt": 16,
        "title": "Tối Ưu Tiết Kiệm Chi Phí Hàng Triệu Đồng",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy with an encouraging smile, hands gesturing savings beside a giant golden piggy bank, floating graphic '+1.000.000 VNĐ / Model', cartoon money bills floating down, clean vector doodle art --ar 16:9",
        "note": "Cô bé người que Thúy nhấn mạnh giá trị tiết kiệm tiền bạc cho doanh nghiệp"
    },
    {
        "stt": 17,
        "title": "So Sánh Dụng Cụ Chứa Đựng vs Tiếp Xúc",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot studio product comparison on dark slate surface: food storage box with sealed lid on left, metal spoon fork chopsticks on right, cinematic edge lighting, clean and crisp 2D vector product art, high contrast doodle aesthetic --ar 16:9",
        "note": "So sánh dụng cụ chứa đựng vs dụng cụ tiếp xúc không chứa đựng"
    },
    {
        "stt": 18,
        "title": "Kim Bài Công Văn 4372 Bộ Y Tế (Aeon)",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of official Vietnamese Ministry of Health dispatch 4372/BYT-VPB6, authentic red stamp emblem, spotlight illumination on official seal, high legal credibility parchment scroll doodle, crisp black ink details --ar 16:9",
        "note": "Chụp cận Công văn 4372/BYT-VPB6 của Bộ Y tế trả lời Aeon"
    },
    {
        "stt": 19,
        "title": "Thìa Đũa Dụng Cụ Thao Tác Cơ Học",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up flat lay shot collage of kitchen utensils: metal spoons, wooden chopsticks, fruit peeling knife, food tongs, silicone cooking gloves, bold red banner 'KHÔNG PHẢI DỤNG CỤ CHỨA ĐỰNG', studio lighting, clean 2D vector doodle --ar 16:9",
        "note": "B-roll các dụng cụ không chứa đựng: đũa, thìa, dĩa, dao gọt, kẹp gắp"
    },
    {
        "stt": 20,
        "title": "Khẳng Định Miễn TCB & Miễn Kiểm Tra Cửa Khẩu",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy gesturing decisively with open hands, confident expert posture, green channel approval graphic 'MIỄN 100%', clean minimalist vector background, bright lighting --ar 16:9",
        "note": "Cô bé người que Thúy khẳng định miễn TCB và miễn kiểm tra cửa khẩu"
    },
    {
        "stt": 21,
        "title": "Xe Container Thông Quan Luồng Xanh Ra Cảng",
        "shot_vn": "Toàn cảnh (Wide Shot)",
        "shot_en": "Wide Shot",
        "prompt": "Wide shot dynamic action doodle of reach stacker vehicle lifting shipping container onto modern delivery truck at seaport terminal, speeding past open green customs barrier 'LUỒNG XANH', bright blue sky, clean vector line art --ar 16:9",
        "note": "B-roll xe container lấy hàng thông quan nhanh chóng ra khỏi cảng"
    },
    {
        "stt": 22,
        "title": "Cảnh Báo Mốc Thời Gian Lab 10-12 Ngày",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of laboratory analytical test tube rack with colorful chemical solutions, digital countdown calendar flip animation overlay showing 'Day 12', analog clock with spinning hands, clean scientific doodle aesthetic --ar 16:9",
        "note": "Cảnh báo mốc thời gian thử nghiệm phòng lab 10-12 ngày"
    },
    {
        "stt": 23,
        "title": "Ác Mộng Hóa Đơn Phạt Dem/Det 15-30 Triệu",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot dramatic doodle of demurrage fee penalty document stamped 'DEMURRAGE: 30.000.000 VND', held against blurry dark port container background, storm clouds and rain, high contrast comic lighting --ar 16:9",
        "note": "Cảnh báo tiền phạt lưu container lưu bãi Dem/Det 15-30 triệu"
    },
    {
        "stt": 24,
        "title": "Tuyệt Chiêu Gửi Mẫu CPN Trước 25 Ngày",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy with ponytail holding an air express document envelope (CPN), smiling enthusiastically, calendar overlay showing '-25 Ngày', sharing insider advice, clean minimalist background --ar 16:9",
        "note": "Cô bé người que Thúy cầm phong bì CPN chia sẻ tuyệt chiêu gửi mẫu trước 25 ngày"
    },
    {
        "stt": 25,
        "title": "Màn Hình Cổng Một Cửa NSW Đạt Chuẩn",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of computer screen displaying Vietnam National Single Window clearance certificate with green approved stamp 'ĐẠT YÊU CẦU - 3 NGÀY', clean modern office doodle setup, celebratory digital badge --ar 16:9",
        "note": "Màn hình Một cửa quốc gia NSW cấp thông báo đạt trong 3 ngày"
    },
    {
        "stt": 26,
        "title": "Xe Cont Cập Kho Riêng Cắt Phí Lưu Bãi",
        "shot_vn": "Toàn cảnh (Wide Shot)",
        "shot_en": "Wide Shot",
        "prompt": "Wide shot of modern clean distribution warehouse interior in 2D vector style, container truck docking at loading bay labeled 'KHO RIÊNG BẢO QUẢN', bright logistics facility, cartoon padlock on shutter door, clean line art --ar 16:9",
        "note": "Đưa container về kho riêng bảo quản cắt đứt tiền lưu kho bãi"
    },
    {
        "stt": 27,
        "title": "Bản Tự Công Bố Có Giá Trị Vô Thời Hạn",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy with hands open in a welcoming, expansive gesture, genuine smile, floating golden card with infinity symbol '∞' and text 'HIỆU LỰC VÔ THỜI HẠN', clean minimalist vector art --ar 16:9",
        "note": "Cô bé người que Thúy nhấn mạnh bản Tự công bố có hiệu lực vô thời hạn"
    },
    {
        "stt": 28,
        "title": "Lưu Ý Phiếu Kiểm Nghiệm Hạn 12 Tháng",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of laboratory analytical certificate with blue accreditation stamp and red date stamp labeled 'HẠN 12 THÁNG', resting beside an elegant glass hourglass with blue flowing sand, warm desk lamp lighting, clean vector art --ar 16:9",
        "note": "Lưu ý phiếu kiểm nghiệm có giá trị 12 tháng khi nộp hồ sơ"
    },
    {
        "stt": 29,
        "title": "Đúc Kết Thông Điệp Triết Lý Thực Chiến",
        "shot_vn": "Cận cảnh (Close-up)",
        "shot_en": "Close-up",
        "prompt": "Close-up shot of cute female stick figure girl Thuy looking directly into lens, powerful leadership presence, cute captain hat, protective stone wall shielding cargo containers in background, visionary comic doodle style --ar 16:9",
        "note": "Cô bé người que Thúy đúc kết thông điệp triết lý logistics thực chiến"
    },
    {
        "stt": 30,
        "title": "Outro Kêu Gọi Subscribe & Tặng Cheat-Sheet",
        "shot_vn": "Trung cảnh (Medium Shot)",
        "shot_en": "Medium Shot",
        "prompt": "Medium shot of cute female stick figure girl Thuy smiling radiantly, waving hand gracefully towards viewer, beside a red Subscribe button, ringing bell, and gift-wrapped handbook 'CẨM NANG BAO BÌ', hearts floating, clean vector doodle --ar 16:9",
        "note": "Cô bé người que Thúy tươi cười vẫy tay chào tạm biệt và tặng Cẩm nang"
    }
]

headers = ["STT", "Tiêu đề", "Prompt tạo ảnh (Image Prompt)", "Góc quay (Shot Type)", "Ghi chú"]

# 1. Xuất file Excel chuẩn theo nhãn Dropdown Web (VN): "Trung cảnh (Medium Shot)", "Toàn cảnh (Wide Shot)", "Cận cảnh (Close-up)"
wb_vn = openpyxl.Workbook()
ws_vn = wb_vn.active
ws_vn.title = "TAO_ANH_HANG_LOAT"
ws_vn.append(headers)

header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
header_fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

for c in range(1, 6):
    cell = ws_vn.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

for s in scenes:
    ws_vn.append([s["stt"], s["title"], s["prompt"], s["shot_vn"], s["note"]])

for r in range(2, len(scenes) + 2):
    ws_vn.cell(row=r, column=1).alignment = center_align
    ws_vn.cell(row=r, column=2).alignment = left_align
    ws_vn.cell(row=r, column=3).alignment = left_align
    ws_vn.cell(row=r, column=4).alignment = center_align
    ws_vn.cell(row=r, column=5).alignment = left_align
    for c in range(1, 6):
        cell = ws_vn.cell(row=r, column=c)
        cell.font = Font(name="Arial", size=10)
        cell.border = thin_border

ws_vn.column_dimensions['A'].width = 8
ws_vn.column_dimensions['B'].width = 35
ws_vn.column_dimensions['C'].width = 65
ws_vn.column_dimensions['D'].width = 25
ws_vn.column_dimensions['E'].width = 45

# 2. Xuất file Excel bản Tiếng Anh thuần (Wide Shot, Medium Shot, Close-up)
wb_en = openpyxl.Workbook()
ws_en = wb_en.active
ws_en.title = "TAO_ANH_HANG_LOAT"
ws_en.append(headers)
for c in range(1, 6):
    cell = ws_en.cell(row=1, column=c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_align

for s in scenes:
    ws_en.append([s["stt"], s["title"], s["prompt"], s["shot_en"], s["note"]])

for r in range(2, len(scenes) + 2):
    ws_en.cell(row=r, column=1).alignment = center_align
    ws_en.cell(row=r, column=2).alignment = left_align
    ws_en.cell(row=r, column=3).alignment = left_align
    ws_en.cell(row=r, column=4).alignment = center_align
    ws_en.cell(row=r, column=5).alignment = left_align
    for c in range(1, 6):
        cell = ws_en.cell(row=r, column=c)
        cell.font = Font(name="Arial", size=10)
        cell.border = thin_border

ws_en.column_dimensions['A'].width = 8
ws_en.column_dimensions['B'].width = 35
ws_en.column_dimensions['C'].width = 65
ws_en.column_dimensions['D'].width = 25
ws_en.column_dimensions['E'].width = 45

# Lưu các file
file_vn_xlsx = r"C:\Users\thuyn\Downloads\ảnh_chuẩn_khớp_dropdown.xlsx"
wb_vn.save(file_vn_xlsx)
print(f"Saved: {file_vn_xlsx}")

file_en_xlsx = r"C:\Users\thuyn\Downloads\ảnh_chuẩn_tiếng_anh.xlsx"
wb_en.save(file_en_xlsx)
print(f"Saved: {file_en_xlsx}")

# CSV bản VN (Song ngữ khớp chính xác text trong ảnh dropdown của user)
csv_vn = r"C:\Users\thuyn\Downloads\Tao_Anh_Hang_Loat_Khop_Dropdown.csv"
with open(csv_vn, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(headers)
    for s in scenes:
        w.writerow([s["stt"], s["title"], s["prompt"], s["shot_vn"], s["note"]])
print(f"Saved: {csv_vn}")

# CSV bản EN (Chuẩn tiếng Anh đơn giản: Wide Shot, Medium Shot, Close-up)
csv_en = r"C:\Users\thuyn\Downloads\Tao_Anh_Hang_Loat_English.csv"
with open(csv_en, 'w', encoding='utf-8-sig', newline='') as f:
    w = csv.writer(f)
    w.writerow(headers)
    for s in scenes:
        w.writerow([s["stt"], s["title"], s["prompt"], s["shot_en"], s["note"]])
print(f"Saved: {csv_en}")
