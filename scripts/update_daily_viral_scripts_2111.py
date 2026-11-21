# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_2111.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 21/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-82, 83, 84) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-83, 84, 85) vào Sheet 1 và 3 gợi ý (STT 79, 80, 81) vào Sheet 4 của File 2:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx
3. Tự động kiểm tra và xóa kịch bản có ghi chú "Đã sản xuất" hoặc "Đã quay" (nếu có).
4. Rà soát Gác cổng 3 Lớp (Từ cấm SSoT banned_words.md + Văn minh + Đếm từ Ping-Pong 2-5 từ).
5. Áp dụng kỹ thuật lưu nguyên tử (Atomic Save) qua file đệm .tmp.xlsx và os.replace.
6. Kích hoạt generator_chi_dan.py để đồng bộ Desktop dashboard.
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
import subprocess

FILE_1 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx"
FILE_2 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx"
GENERATOR_SCRIPT = r"D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py"

thin_side = Side(style='thin', color='D9D9D9')
grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

font_regular = Font(name='Arial', size=10, bold=False, color='000000')
font_bold = Font(name='Arial', size=10, bold=True, color='000000')
font_code = Font(name='Arial', size=10, bold=True, color='1F497D')

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left_top = Alignment(horizontal='left', vertical='top', wrap_text=True)

BANNED_TERMS = [
    "trốn thuế", "lách luật", "hàng lậu", "buôn lậu", "hàng nhái", "hàng giả",
    "bao thuế", "đi chui", "đi lậu", "tiểu ngạch", "mua hàng trung quốc tận gốc",
    "mua hàng trung tận gốc", "đánh hàng quảng châu", "đánh hàng trung quốc",
    "gom hàng biên", "chạy line", "chạy luồng", "bôi trơn", "lót tay",
    "hối lộ", "đi cửa sau", "bắn hàng qua biên", "thông chốt", "đồ chơi trẻ em",
    "đồ chơi", "kiếm tiền", "làm giàu", "thu nhập thụ động", "lừa đảo", "scam",
    "bạo lực", "inbox ngay", "dm ngay", "dở à", "đếch", "điên à",
    "bắt cóc", "đi tù", "công an hốt", "ai thay thế hoàn toàn con người",
    "100% không bị phạt"
]

# Dữ liệu 3 kịch bản mới cho File 1 (BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx)
SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-82",
        "title": "Khách dị ứng giọng nữ vì bị vợ mắng và màn chốt đơn khó đỡ (Tham khảo Reel 460K views)",
        "mb_raw": "• Nhân viên: Em chào anh ạ!\n• Khách: Đã bảo không cần chào rồi bao nhiêu lần gì! Cứ nghe thấy cái giọng phụ nữ là ám ảnh đau hết cả đầu!\n• Nhân viên: Dạ thế thì em... (hắng giọng trầm) Anh ạ!\n• Chú đồng nghiệp: Này, cái cậu khách kia hay bị vợ mắng nhỉ!\n• Khách: Tuần chú chỉ bị một lần thôi!\n• Nhân viên: Vâng, không sao, khách hàng là thượng đế chú ạ!\n• Chú đồng nghiệp: Con bé này công nhận tính nết dịu dàng thật đấy!\n• Nhân viên: Một tuần cứ có 10 khách như thế cháu cười đến tối cũng được!\n• Chú đồng nghiệp: Úi giời ơi, ngân hàng mình uy tín thì ngày nào chả đông khách mà!\n• Chú đồng nghiệp: Đặc điểm nhận dạng là gì để chú đưa vào quầy mời?\n• Nhân viên: Mỗi lần đến quát tháo xong mở một sổ 8 tỷ chú ạ!\n• Chú đồng nghiệp (quay xe): Gia trưởng thế này hơi khó tìm đấy!\n• (Sound: Aiyo mama! Cả quầy cười rộ)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales nữ (cúi chào): Em chào anh ạ!\n• Khách (xua tay, gắt): Đã bảo không chào!\n• Khách: Nghe giọng nữ ám ảnh!\n• [ESCALATION 4-15s]\n• Sales nữ (hắng giọng trầm): Dạ chào anh ạ!\n• Sếp Thúy (nhìn): Khách hay bị vợ mắng?\n• Khách: Tuần bị một lần thôi!\n• Sales nữ: Khách hàng là thượng đế!\n• Sếp Thúy: Em chịu khó thật đấy!\n• Sales nữ: Có mười khách thế này!\n• Sales nữ: Em cười đến tối luôn!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Đặc điểm nhận dạng đâu?\n• Sếp Thúy: Ghi sổ chăm sóc riêng!\n• Sales nữ: Đến quát tháo ầm ĩ!\n• Sales nữ: Xong chốt mười cont CNC!\n• Sales nữ: Đóng đủ hóa đơn VAT!\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Chốt mười cont?\n• Sếp Thúy (tỉnh bơ): Mẫu người hơi hiếm đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Bác khách chủ xưởng cơ khí (vẻ cộc cằn nhưng sộp) + Sales nữ Eureka (áo đồng phục, cười tươi hết cỡ) + Sếp Thúy (ngồi bàn giám đốc ghi sổ).",
        "scene": "Sảnh cửa kính & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, tập kẹp hồ sơ A4, hợp đồng ủy thác máy CNC, bút ký.",
        "twist": "Bác khách bước vào văn phòng xua tay cấm chào vì ở nhà hay bị vợ mắng nên ám ảnh giọng phụ nữ; nữ Sales vẫn niềm nở đổi giọng trầm tiếp đón và khen 'có mười khách thế này em cười đến tối'. Sếp Thúy hỏi đặc điểm nhận dạng để ghi sổ lưu ý, vừa nghe bác mỗi lần quát tháo xong lại chốt 10 cont máy CNC kèm 100% hóa đơn VAT thì Sếp Thúy quay xe tỉnh bơ khen 'Mẫu người hơi hiếm đấy!'.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-83",
        "title": "Khách đọc vị hết kịch bản tư vấn của Sales Eureka và pha tự chốt đơn (Tham khảo Reel 257K views)",
        "mb_raw": "• Nhân viên: Dạ anh ơi!\n• Khách: Im! Định nói là 10 triệu 8.2, 500 nhận lãi trước, đúng không?\n• Nhân viên: Đúng ạ!\n• Khách: Ừ! Định hỏi gửi tại quầy hay gửi online đúng không?\n• Nhân viên: Dạ vâng ạ!\n• Khách: Đây đến đây đấy, nhưng mà thích gửi online có được không? Không cần, đây tự thao tác được! Đi ra, đây không thích phụ nữ, phiền chết đi được đấy!\n• (Sound: Aiyo mama! Cả quầy đứng hình đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales: Dạ em chào anh!\n• Khách (giơ tay): Khoan đã!\n• Khách: Định báo gom một khối?\n• Sales (giật mình): Dạ đúng ạ!\n• [ESCALATION 4-15s]\n• Khách: Định mời gói năm triệu?\n• Sales (ngớ người): Sao anh biết hết?\n• Khách: Định bảo xe chạy ngày?\n• Khách: Về Hà Nội sáu ngày?\n• Sales: Dạ vâng đúng luôn!\n• Khách: Định hỏi đặt tại quầy?\n• Khách: Hay đặt trên Portal ERP?\n• Sales: Dạ vâng chuẩn quá!\n• [PUNCHLINE 16-25s]\n• Khách: Đến đây ngồi cho mát!\n• Khách: Chứ tự bấm xong rồi!\n• Khách: Thao tác online mười giây!\n• Khách: Em ra chỗ khác đi!\n• Khách: Nói nhiều phiền quá đi!\n• [DEFLATION 26-30s]\n• Sales (đơ 1.5s): Ơ hay! Đọc vị hết?\n• Sếp Thúy (tỉnh bơ): Khách thuộc bài hơn em!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Bác khách chủ xưởng (đeo kính râm, cầm điện thoại mở Portal ERP) + Sales nữ Eureka (cầm catalogue, ngơ ngác) + Sếp Thúy (ngồi bàn làm việc nhìn sang).",
        "scene": "Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, điện thoại mở Portal ERP v2.5.0, catalogue gói Phòng XNK 5 triệu, cốc trà.",
        "twist": "Sales vừa mở lời chào, bác khách đã giơ tay chặn họng và đọc vanh vách toàn bộ kịch bản tư vấn như biết trước tương lai (gom cont từ 1 khối, Phòng XNK 5 triệu, xe chạy hàng ngày 6-8 ngày về Hà Nội, đặt trên Portal ERP). Xong xuôi bác tỉnh bơ bảo tự thao tác online xong trong 10 giây rồi, ghé văn phòng ngồi cho mát và xua Sales ra chỗ khác kẻo phiền.",
        "duration": "22s"
    },
    {
        "code": "EKA-REEL-84",
        "title": "Chủ xưởng tới văn phòng chỉ để thể hiện đẳng cấp và cái kết ngoài cửa kính (Tham khảo Reel 33K views)",
        "mb_raw": "• Khách: Đẳng cấp như anh là chỉ sử dụng những sản phẩm dịch vụ tiên phong dẫn đầu thôi!\n• Nhân viên: À vâng ạ, MB bên em là ngân hàng tiên phong dẫn đầu về chuyển đổi số!\n• Khách: Biết rồi, thế nên mới đến đây để gửi tiết kiệm!\n• Nhân viên: Ôi anh ơi, anh không cần phải đến trực tiếp đâu! Vì hầu hết các dịch vụ đều có thể thao tác online trên app MB!\n• Khách: Khổ lắm, đã bảo người ta đẳng cấp người ta biết rồi! Đang thao tác gửi online đây! Xong rồi đây này, ra đây chỉ để thể hiện đẳng cấp thôi!\n• Nhân viên: Đẳng cấp như anh chắc phải đi xe hơi hiện đại lắm nhỉ?\n• Khách: Đẳng cấp như anh sao phải đi ô tô! Chơi hẳn xe... (đạp xe cọc cạch / xe máy chở đồ)\n• (Sound: Aiyo mama! Cả quầy đứng hình)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách: Anh chỉ dùng hàng đầu!\n• Sales: Eureka tiên phong công nghệ!\n• Khách: Biết rồi mới tới đây!\n• [ESCALATION 4-15s]\n• Sales: Không cần qua trực tiếp!\n• Sales: Tính cước ba mươi giây!\n• Sales: Lên web erktransport.com ngay!\n• Khách: Khổ lắm, biết hết rồi!\n• Khách: Đang tự bấm online đây!\n• Khách: Xong lô linh kiện rồi!\n• Sales: Thế bác qua làm gì?\n• Khách: Để thể hiện đẳng cấp!\n• [PUNCHLINE 16-25s]\n• Sales: Đẳng cấp thế này ạ?\n• Sales: Chắc bác đi siêu xe?\n• Khách: Sao phải đi ô tô?\n• (Khách bước ra cửa kính văn phòng, dắt chiếc xe đạp cà tàng kêu cọt kẹt)\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Siêu xe đây?\n• Khách (tỉnh bơ): Bảo vệ môi trường nhé!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Bác khách chủ xưởng (áo vest bóng bẩy, đeo kính râm, cầm smartphone) + Sales Eureka (áo đồng phục) + Sếp Thúy (nhìn ra cửa kính).",
        "scene": "Bàn tiếp khách & góc cửa kính sảnh văn phòng Eureka 44 Triều Khúc (nhìn xuyên ra ngoài thấy chiếc xe đạp cà tàng dựng trước cửa), màn hình web erktransport.com.",
        "twist": "Bác khách ăn mặc bảnh bao tuyên bố đẳng cấp chỉ dùng dịch vụ tiên phong công nghệ, dù biết web erktransport.com tính cước 30 giây và Portal ERP đặt đơn online không cần đến nơi nhưng vẫn lặn lội tới tận 44 Triều Khúc bấm app chỉ để... thể hiện đẳng cấp. Sales ngưỡng mộ hỏi chắc bác đi siêu xe, bác chốt 'Đẳng cấp sao đi ô tô' rồi dắt chiếc xe đạp cà tàng ra về.",
        "duration": "24s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-83",
        "motif": "Motif 4 & 2: Hài gia đình (Ám ảnh vợ mắng) & Chăm sóc khách hàng VIP 1-1",
        "title": "Khách dị ứng giọng nữ vì bị vợ mắng và màn chốt đơn khó đỡ (Tham khảo Reel 460K views)",
        "strength": "Dịch vụ Dedicated Team 1-1 (Quản lý + CSKH + Chứng từ) tận tâm phục vụ, ủy thác nhập khẩu máy móc CNC chính ngạch 100% hóa đơn VAT & bảo hiểm đền bù mất mát.",
        "roles_scene": "• Bối cảnh: Sảnh cửa kính & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Bác khách chủ xưởng cơ khí + Sales nữ Eureka + Sếp Thúy\n• Đạo cụ: Tập kẹp hồ sơ A4, hợp đồng ủy thác máy CNC, bút ký.",
        "pingpong": "• [HOOK 0-3s]\n• Sales nữ (cúi chào): Em chào anh ạ!\n• Khách (xua tay, gắt): Đã bảo không chào!\n• Khách: Nghe giọng nữ ám ảnh!\n• [ESCALATION 4-15s]\n• Sales nữ (hắng giọng trầm): Dạ chào anh ạ!\n• Sếp Thúy (nhìn): Khách hay bị vợ mắng?\n• Khách: Tuần bị một lần thôi!\n• Sales nữ: Khách hàng là thượng đế!\n• Sếp Thúy: Em chịu khó thật đấy!\n• Sales nữ: Có mười khách thế này!\n• Sales nữ: Em cười đến tối luôn!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Đặc điểm nhận dạng đâu?\n• Sếp Thúy: Ghi sổ chăm sóc riêng!\n• Sales nữ: Đến quát tháo ầm ĩ!\n• Sales nữ: Xong chốt mười cont CNC!\n• Sales nữ: Đóng đủ hóa đơn VAT!\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Chốt mười cont?\n• Sếp Thúy (tỉnh bơ): Mẫu người hơi hiếm đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Bác khách bước vào văn phòng xua tay cấm chào vì ở nhà hay bị vợ mắng nên ám ảnh giọng phụ nữ; nữ Sales vẫn niềm nở đổi giọng trầm tiếp đón và khen 'có mười khách thế này em cười đến tối'. Sếp Thúy hỏi đặc điểm nhận dạng để ghi sổ lưu ý, vừa nghe bác mỗi lần quát tháo xong lại chốt 10 cont máy CNC kèm 100% hóa đơn VAT thì Sếp Thúy quay xe tỉnh bơ khen 'Mẫu người hơi hiếm đấy!'.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-84",
        "motif": "Motif 6 & 5: Đấu trí đọc vị & Khách thuộc lòng hệ sinh thái số hóa",
        "title": "Khách đọc vị hết kịch bản tư vấn của Sales Eureka và pha tự chốt đơn (Tham khảo Reel 257K views)",
        "strength": "Gom cont LCL từ 1 khối (1 CBM) xe chạy hàng ngày 6-8 ngày về Hà Nội, Phòng XNK thuê ngoài 5 triệu/tháng & Portal ERP v2.5.0 đặt đơn online 10 giây minh bạch 0 phí ẩn.",
        "roles_scene": "• Bối cảnh: Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Bác khách chủ xưởng + Sales nữ Eureka + Sếp Thúy\n• Đạo cụ: Điện thoại mở Portal ERP v2.5.0, catalogue gói Phòng XNK 5 triệu, cốc trà.",
        "pingpong": "• [HOOK 0-3s]\n• Sales: Dạ em chào anh!\n• Khách (giơ tay): Khoan đã!\n• Khách: Định báo gom một khối?\n• Sales (giật mình): Dạ đúng ạ!\n• [ESCALATION 4-15s]\n• Khách: Định mời gói năm triệu?\n• Sales (ngớ người): Sao anh biết hết?\n• Khách: Định bảo xe chạy ngày?\n• Khách: Về Hà Nội sáu ngày?\n• Sales: Dạ vâng đúng luôn!\n• Khách: Định hỏi đặt tại quầy?\n• Khách: Hay đặt trên Portal ERP?\n• Sales: Dạ vâng chuẩn quá!\n• [PUNCHLINE 16-25s]\n• Khách: Đến đây ngồi cho mát!\n• Khách: Chứ tự bấm xong rồi!\n• Khách: Thao tác online mười giây!\n• Khách: Em ra chỗ khác đi!\n• Khách: Nói nhiều phiền quá đi!\n• [DEFLATION 26-30s]\n• Sales (đơ 1.5s): Ơ hay! Đọc vị hết?\n• Sếp Thúy (tỉnh bơ): Khách thuộc bài hơn em!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Sales vừa mở lời chào, bác khách đã giơ tay chặn họng và đọc vanh vách toàn bộ kịch bản tư vấn như biết trước tương lai (gom cont từ 1 khối, Phòng XNK 5 triệu, xe chạy hàng ngày 6-8 ngày về Hà Nội, đặt trên Portal ERP). Xong xuôi bác tỉnh bơ bảo tự thao tác online xong trong 10 giây rồi, ghé văn phòng ngồi cho mát và xua Sales ra chỗ khác kẻo phiền.",
        "duration": "22s"
    },
    {
        "code": "EKA-V2-85",
        "motif": "Motif 1 & 5: Slapstick đối lập cửa kính & Chơi chữ 'Đẳng cấp tiên phong'",
        "title": "Chủ xưởng tới văn phòng chỉ để thể hiện đẳng cấp và cái kết ngoài cửa kính (Tham khảo Reel 33K views)",
        "strength": "Website erktransport.com tính cước tự động trong 30 giây, Portal ERP theo dõi camera kho Quảng Châu - Bằng Tường 24/7 cho hàng linh kiện cơ khí & gia dụng thông minh.",
        "roles_scene": "• Bối cảnh: Bàn tiếp khách & góc cửa kính sảnh văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Bác khách chủ xưởng (áo vest, kính râm) + Sales Eureka + Sếp Thúy\n• Đạo cụ: Smartphone mở web erktransport.com, xe đạp cà tàng dựng ngoài cửa kính.",
        "pingpong": "• [HOOK 0-3s]\n• Khách: Anh chỉ dùng hàng đầu!\n• Sales: Eureka tiên phong công nghệ!\n• Khách: Biết rồi mới tới đây!\n• [ESCALATION 4-15s]\n• Sales: Không cần qua trực tiếp!\n• Sales: Tính cước ba mươi giây!\n• Sales: Lên web erktransport.com ngay!\n• Khách: Khổ lắm, biết hết rồi!\n• Khách: Đang tự bấm online đây!\n• Khách: Xong lô linh kiện rồi!\n• Sales: Thế bác qua làm gì?\n• Khách: Để thể hiện đẳng cấp!\n• [PUNCHLINE 16-25s]\n• Sales: Đẳng cấp thế này ạ?\n• Sales: Chắc bác đi siêu xe?\n• Khách: Sao phải đi ô tô?\n• (Khách bước ra cửa kính văn phòng, dắt chiếc xe đạp cà tàng kêu cọt kẹt)\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Siêu xe đây?\n• Khách (tỉnh bơ): Bảo vệ môi trường nhé!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Bác khách ăn mặc bảnh bao tuyên bố đẳng cấp chỉ dùng dịch vụ tiên phong công nghệ, dù biết web erktransport.com tính cước 30 giây và Portal ERP đặt đơn online không cần đến nơi nhưng vẫn lặn lội tới tận 44 Triều Khúc bấm app chỉ để... thể hiện đẳng cấp. Sales ngưỡng mộ hỏi chắc bác đi siêu xe, bác chốt 'Đẳng cấp sao đi ô tô' rồi dắt chiếc xe đạp cà tàng ra về.",
        "duration": "24s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 79,
        "title": "Khách dị ứng giọng nữ vì bị vợ mắng và màn chốt đơn khó đỡ (Reel 460K views)",
        "views": "460,000 views",
        "motif": "Motif 4 & 2: Hài gia đình (Ám ảnh vợ mắng) & CSKH 1-1",
        "trigger": "Khách bước vào quầy cấm nhân viên chào vì ám ảnh giọng phụ nữ do ở nhà hay bị vợ mắng, nhân viên vẫn tươi cười bảo có 10 khách thế này cười đến tối; quản lý hỏi đặc điểm nhận dạng, nghe xong 'mỗi lần quát xong mở sổ 8 tỷ' liền quay xe khen hiếm có.",
        "adaptation": "Bác chủ xưởng cơ khí tới Eureka xua tay cấm chào vì ám ảnh giọng nữ ở nhà, nữ Sales đổi giọng trầm phục vụ tận tình; Sếp Thúy ghi sổ chăm sóc riêng, nghe bảo mỗi lần quát xong chốt 10 cont máy CNC kèm VAT liền tấm tắc 'Mẫu người hơi hiếm đấy!' (EKA-V2-83)."
    },
    {
        "stt": 80,
        "title": "Khách đọc vị hết kịch bản tư vấn của Sales Eureka và pha tự chốt đơn (Reel 257K views)",
        "views": "257,000 views",
        "motif": "Motif 6 & 5: Đấu trí đọc vị & Khách thuộc lòng Portal ERP",
        "trigger": "Nhân viên vừa mở lời chào đã bị khách ngắt lời đọc vanh vách từng câu sắp tư vấn như đọc kịch bản, xong bảo tự thao tác online được rồi và xua nhân viên ra chỗ khác kẻo phiền.",
        "adaptation": "Bác khách đọc vị trước toàn bộ bài tư vấn của Sales (gom cont 1 khối, Phòng XNK 5 triệu, xe chạy ngày 6-8 ngày về HN, đặt trên Portal ERP), rồi bảo tự bấm xong trong 10 giây và xua Sales ra chỗ khác ngồi ké điều hòa cho mát (EKA-V2-84)."
    },
    {
        "stt": 81,
        "title": "Chủ xưởng tới văn phòng chỉ để thể hiện đẳng cấp và cái kết ngoài cửa kính (Reel 33K views)",
        "views": "33,000 views",
        "motif": "Motif 1 & 5: Slapstick đối lập cửa kính & Chơi chữ đẳng cấp",
        "trigger": "Khách tự nhận đẳng cấp chỉ dùng công nghệ số tiên phong, tự thao tác online xong xuôi rồi nhưng vẫn lặn lội ra tận quầy chỉ để thể hiện đẳng cấp, cuối cùng ra về bằng phương tiện bình dân trái ngược hoàn toàn.",
        "adaptation": "Bác khách khoe đẳng cấp dùng web erktransport.com tính cước 30s và Portal ERP đặt xong lô linh kiện online vẫn tới tận 44 Triều Khúc ngồi sofa thể hiện đẳng cấp; Sales hỏi chắc đi siêu xe thì bác chốt 'Đẳng cấp sao đi ô tô' rồi dắt xe đạp cà tàng về (EKA-V2-85)."
    }
]

def atomic_save(wb, target_path):
    """Ghi workbook an toàn theo RULE 10 (Atomic Save)."""
    temp_path = target_path + ".tmp.xlsx"
    wb.save(temp_path)
    wb.close()
    try:
        os.replace(temp_path, target_path)
        print(f"[✓] Atomic Save thành công: {os.path.basename(target_path)}")
        return True
    except PermissionError:
        print(f"[!] File đang bị mở bởi ứng dụng khác: {target_path}")
        print(f"[*] Đã lưu vào file đệm: {temp_path}")
        return False
    except Exception as e:
        print(f"[!] Lỗi khi đổi tên file đệm: {e}")
        return False

def validate_banned_words(items, label):
    """Rà soát gác cổng Lớp 1 & 2: Quét từ cấm SSoT và ngôn từ chưa văn minh."""
    print("\n" + "="*60)
    print(f"[*] RÀ SOÁT GÁC CỔNG LỚP 1 & 2 ({label}): BANNED WORDS & VĂN MINH")
    print("="*60)
    all_clean = True
    for it in items:
        code = it.get("code", it.get("title", ""))
        fields = [
            ("title", it.get("title", "")),
            ("eureka_trans", it.get("eureka_trans", "")),
            ("pingpong", it.get("pingpong", "")),
            ("roles", it.get("roles", "")),
            ("roles_scene", it.get("roles_scene", "")),
            ("scene", it.get("scene", "")),
            ("twist", it.get("twist", "")),
            ("strength", it.get("strength", "")),
            ("adaptation", it.get("adaptation", "")),
        ]
        for fname, fval in fields:
            low = fval.lower()
            for term in BANNED_TERMS:
                if term in low:
                    print(f"[!] VI PHẠM TỪ CẤM '{term}' tại [{code}] ({fname})")
                    all_clean = False
    if all_clean:
        print(f"[✓] 100% SẠCH TỪ CẤM & ĐẠT CHUẨN VĂN MINH B2B ({label})!")
    return all_clean

def validate_pingpong(scripts):
    """Rà soát gác cổng Lớp 3: Kiểm tra từng câu thoại có độ dài 2-5 từ."""
    print("\n" + "="*60)
    print("[*] RÀ SOÁT GÁC CỔNG LỚP 3: KIỂM TRA ĐỘ DÀI CÂU THOẠI PING-PONG (2-5 TỪ)")
    print("="*60)
    all_valid = True
    for sc in scripts:
        code = sc.get('code')
        lines = sc.get('eureka_trans', sc.get('pingpong', '')).split('\n')
        for line in lines:
            line_str = line.strip()
            if not line_str or line_str.startswith('• [') or line_str.startswith('• (') or line_str.startswith('('):
                continue
            if ':' in line_str:
                utterance = line_str.split(':', 1)[1].strip()
            else:
                utterance = line_str.replace('•', '').strip()
            words = [w for w in utterance.replace('!', '').replace('?', '').replace('.', '').replace(',', '').split() if w]
            count = len(words)
            if count < 2 or count > 5:
                print(f"[!] VI PHẠM ĐỘ DÀI ({count} từ): [{code}] {line_str}")
                all_valid = False
    if all_valid:
        print("[✓] 100% CÂU THOẠI ĐẠT CHUẨN PING-PONG (2-5 TỪ/CÂU)!")
    return all_valid

def update_file_1():
    print("\n" + "="*60)
    print(f"[*] ĐANG CẬP NHẬT FILE 1: {os.path.basename(FILE_1)}")
    print("="*60)

    wb = openpyxl.load_workbook(FILE_1)
    ws = wb['KICH_BAN_EUREKA_CHUAN_MB']

    current_max_row = ws.max_row
    last_stt = 0
    for r in range(4, current_max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, int):
            last_stt = max(last_stt, v)
        elif str(v).isdigit():
            last_stt = max(last_stt, int(v))

    print(f"[*] STT hiện tại trong File 1: {last_stt} (Row {current_max_row})")

    existing_codes = [str(ws.cell(r, 2).value) for r in range(4, current_max_row + 1)]

    for script in SCRIPTS_FILE_1:
        if script['code'] in existing_codes:
            print(f"[-] Mã {script['code']} đã tồn tại trong File 1, bỏ qua.")
            continue

        last_stt += 1
        new_row = ws.max_row + 1

        ws.cell(new_row, 1, last_stt).alignment = align_center
        ws.cell(new_row, 1).font = font_regular
        ws.cell(new_row, 1).border = grid_border

        ws.cell(new_row, 2, script['code']).alignment = align_center
        ws.cell(new_row, 2).font = font_code
        ws.cell(new_row, 2).border = grid_border

        ws.cell(new_row, 3, script['title']).alignment = align_left_top
        ws.cell(new_row, 3).font = font_bold
        ws.cell(new_row, 3).border = grid_border

        ws.cell(new_row, 4, script['mb_raw']).alignment = align_left_top
        ws.cell(new_row, 4).font = font_regular
        ws.cell(new_row, 4).border = grid_border

        ws.cell(new_row, 5, script['eureka_trans']).alignment = align_left_top
        ws.cell(new_row, 5).font = font_regular
        ws.cell(new_row, 5).border = grid_border

        ws.cell(new_row, 6, script['roles']).alignment = align_left_top
        ws.cell(new_row, 6).font = font_regular
        ws.cell(new_row, 6).border = grid_border

        ws.cell(new_row, 7, script['scene']).alignment = align_left_top
        ws.cell(new_row, 7).font = font_regular
        ws.cell(new_row, 7).border = grid_border

        ws.cell(new_row, 8, script['twist']).alignment = align_left_top
        ws.cell(new_row, 8).font = font_regular
        ws.cell(new_row, 8).border = grid_border

        ws.cell(new_row, 9, script['duration']).alignment = align_center
        ws.cell(new_row, 9).font = font_regular
        ws.cell(new_row, 9).border = grid_border

        print(f"[+] Đã thêm dòng mới File 1: STT {last_stt} | Mã {script['code']} | Row {new_row}")

    return atomic_save(wb, FILE_1)

def update_file_2():
    print("\n" + "="*60)
    print(f"[*] ĐANG CẬP NHẬT FILE 2: {os.path.basename(FILE_2)}")
    print("="*60)

    wb = openpyxl.load_workbook(FILE_2)

    # --- 1. Cập nhật Sheet 1: KICH_BAN_VIRAL_REELS ---
    ws1 = wb['KICH_BAN_VIRAL_REELS']
    current_max_row = ws1.max_row

    # Cơ chế tự động xóa kịch bản "Đã sản xuất" hoặc "Đã quay" nếu có
    rows_to_delete = []
    for r in range(5, current_max_row + 1):
        row_vals = [str(ws1.cell(r, c).value or '') for c in range(1, ws1.max_column + 1)]
        if any("đã sản xuất" in v.lower() or "đã quay" in v.lower() for v in row_vals):
            rows_to_delete.append(r)

    if rows_to_delete:
        for r in reversed(rows_to_delete):
            print(f"[!] Xóa kịch bản đã sản xuất tại Row {r}: Mã {ws1.cell(r, 2).value}")
            ws1.delete_rows(r)

    last_stt = 0
    for r in range(5, ws1.max_row + 1):
        v = ws1.cell(r, 1).value
        if isinstance(v, int):
            last_stt = max(last_stt, v)
        elif str(v).isdigit():
            last_stt = max(last_stt, int(v))

    print(f"[*] STT hiện tại trong File 2 Sheet 1: {last_stt} (Row {ws1.max_row})")

    existing_codes = [str(ws1.cell(r, 2).value) for r in range(5, ws1.max_row + 1)]

    for script in SCRIPTS_FILE_2:
        if script['code'] in existing_codes:
            print(f"[-] Mã {script['code']} đã tồn tại trong File 2, bỏ qua.")
            continue

        last_stt += 1
        new_row = ws1.max_row + 1

        ws1.cell(new_row, 1, last_stt).alignment = align_center
        ws1.cell(new_row, 1).font = font_regular
        ws1.cell(new_row, 1).border = grid_border

        ws1.cell(new_row, 2, script['code']).alignment = align_center
        ws1.cell(new_row, 2).font = font_code
        ws1.cell(new_row, 2).border = grid_border

        ws1.cell(new_row, 3, script['motif']).alignment = align_left_top
        ws1.cell(new_row, 3).font = font_regular
        ws1.cell(new_row, 3).border = grid_border

        ws1.cell(new_row, 4, script['title']).alignment = align_left_top
        ws1.cell(new_row, 4).font = font_bold
        ws1.cell(new_row, 4).border = grid_border

        ws1.cell(new_row, 5, script['strength']).alignment = align_left_top
        ws1.cell(new_row, 5).font = font_regular
        ws1.cell(new_row, 5).border = grid_border

        ws1.cell(new_row, 6, script['roles_scene']).alignment = align_left_top
        ws1.cell(new_row, 6).font = font_regular
        ws1.cell(new_row, 6).border = grid_border

        ws1.cell(new_row, 7, script['pingpong']).alignment = align_left_top
        ws1.cell(new_row, 7).font = font_regular
        ws1.cell(new_row, 7).border = grid_border

        ws1.cell(new_row, 8, script['twist']).alignment = align_left_top
        ws1.cell(new_row, 8).font = font_regular
        ws1.cell(new_row, 8).border = grid_border

        ws1.cell(new_row, 9, script['duration']).alignment = align_center
        ws1.cell(new_row, 9).font = font_regular
        ws1.cell(new_row, 9).border = grid_border

        print(f"[+] Đã thêm dòng mới File 2 Sheet 1: STT {last_stt} | Mã {script['code']} | Row {new_row}")

    # --- 2. Cập nhật Sheet 4: GOI_Y_Y_TUONG_MB ---
    ws4 = wb['GOI_Y_Y_TUONG_MB']
    last_stt_s4 = 0
    for r in range(5, ws4.max_row + 1):
        v = ws4.cell(r, 1).value
        if isinstance(v, int):
            last_stt_s4 = max(last_stt_s4, v)
        elif str(v).isdigit():
            last_stt_s4 = max(last_stt_s4, int(v))

    print(f"[*] STT hiện tại trong File 2 Sheet 4: {last_stt_s4} (Row {ws4.max_row})")
    existing_titles_s4 = [str(ws4.cell(r, 2).value) for r in range(5, ws4.max_row + 1)]

    for idea in IDEAS_SHEET_4:
        if idea['title'] in existing_titles_s4:
            print(f"[-] Ý tưởng '{idea['title']}' đã có trong Sheet 4, bỏ qua.")
            continue

        last_stt_s4 += 1
        new_row = ws4.max_row + 1

        ws4.cell(new_row, 1, last_stt_s4).alignment = align_center
        ws4.cell(new_row, 1).font = font_regular
        ws4.cell(new_row, 1).border = grid_border

        ws4.cell(new_row, 2, idea['title']).alignment = align_left_top
        ws4.cell(new_row, 2).font = font_bold
        ws4.cell(new_row, 2).border = grid_border

        ws4.cell(new_row, 3, idea['views']).alignment = align_center
        ws4.cell(new_row, 3).font = font_regular
        ws4.cell(new_row, 3).border = grid_border

        ws4.cell(new_row, 4, idea['motif']).alignment = align_left_top
        ws4.cell(new_row, 4).font = font_regular
        ws4.cell(new_row, 4).border = grid_border

        ws4.cell(new_row, 5, idea['trigger']).alignment = align_left_top
        ws4.cell(new_row, 5).font = font_regular
        ws4.cell(new_row, 5).border = grid_border

        ws4.cell(new_row, 6, idea['adaptation']).alignment = align_left_top
        ws4.cell(new_row, 6).font = font_regular
        ws4.cell(new_row, 6).border = grid_border

        print(f"[+] Đã thêm dòng mới File 2 Sheet 4: STT {last_stt_s4} | {idea['title'][:45]} | Row {new_row}")

    return atomic_save(wb, FILE_2)

def run_desktop_sync():
    print("\n" + "="*60)
    print("[*] ĐỒNG BỘ BÁO CÁO VÀO BẢNG CHỈ DẪN DESKTOP")
    print("="*60)
    if os.path.exists(GENERATOR_SCRIPT):
        try:
            res = subprocess.run(["python", GENERATOR_SCRIPT], capture_output=True, text=True, timeout=60)
            print("[✓] Kết quả chạy generator_chi_dan.py:")
            print(res.stdout.strip())
            if res.stderr:
                print("[!] Cảnh báo stderr:", res.stderr.strip())
        except Exception as e:
            print(f"[!] Lỗi khi chạy generator_chi_dan.py: {e}")
    else:
        print(f"[!] Không tìm thấy script: {GENERATOR_SCRIPT}")

if __name__ == "__main__":
    ok_bw1 = validate_banned_words(SCRIPTS_FILE_1, "FILE 1")
    ok_bw2 = validate_banned_words(SCRIPTS_FILE_2, "FILE 2")
    ok_bw4 = validate_banned_words(IDEAS_SHEET_4, "SHEET 4")
    ok_pp1 = validate_pingpong(SCRIPTS_FILE_1)
    ok_pp2 = validate_pingpong(SCRIPTS_FILE_2)

    if not (ok_bw1 and ok_bw2 and ok_bw4 and ok_pp1 and ok_pp2):
        print("\n[!] REJECT: Kịch bản chưa đạt chuẩn Gác cổng 3 Lớp. Dừng xuất bản!")
        sys.exit(1)

    ok1 = update_file_1()
    ok2 = update_file_2()
    if ok1 and ok2:
        print("\n[✓] CẢ 2 FILE EXCEL TRÊN ONEDRIVE ĐÃ ĐƯỢC CẬP NHẬT HOÀN TẤT VÀ AN TOÀN!")
        run_desktop_sync()
    else:
        print("\n[!] Cảnh báo: Có lỗi trong quá trình lưu file.")
        sys.exit(1)
