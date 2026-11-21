# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_1911.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 19/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-76, 77, 78) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-77, 78, 79) vào Sheet 1 và 3 gợi ý (STT 73, 74, 75) vào Sheet 4 của File 2:
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
    "đánh hàng quảng châu", "đánh hàng trung quốc", "gom hàng biên", "chạy line",
    "chạy luồng", "bôi trơn", "lót tay", "hối lộ", "đi cửa sau", "bắn hàng qua biên",
    "thông chốt", "đồ chơi trẻ em", "đồ chơi", "kiếm tiền", "làm giàu",
    "thu nhập thụ động", "lừa đảo", "scam", "bạo lực", "inbox ngay", "dm ngay",
    "dở à", "đếch", "điên à", "bắt cóc", "đi tù", "công an hốt",
    "ai thay thế hoàn toàn con người", "100% không bị phạt"
]

# Dữ liệu 3 kịch bản mới cho File 1 (BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx)
SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-76",
        "title": "Học cấp tốc tiếng Trung đón sếp lớn và cái kết đặt đơn online trên Portal ERP (Tham khảo Reel 1.1M views)",
        "mb_raw": "• Nhân viên (chắp tay cúi chào liên tục): Sa-wa-đi-khạp... Khọp khun khạp...\n• Khách: Tao mỏi cổ lắm rồi đấy!\n• Nhân viên: Ô! Anh biết nói tiếng Việt á?\n• Khách: Sao sếp em lại gặp khách Thái Lan nhỉ?\n• Nhân viên: Người Thái không việc mới sang bên đấy làm ăn kiếm được, nhưng mỗi tội phải đến tận MB Hải Dương mở tiết kiệm hơi xa!\n• Khách: Ô anh ơi MB chúng em chuyển đổi số mấy năm nay, người Việt ở nước ngoài vẫn gửi tiết kiệm online được anh ạ!\n• Nhân viên: Úi giời thế mà không bảo sớm, tiết kiệm tiền vé máy bay! Mà mấy khi được đón tiếp Việt kiều Thái đâu, chúng em còn học một khóa tiếng Thái để đón anh đấy! Giờ chúng em thực hành nói tiếng Thái với anh nhé: Mấy cái tiếng khạp ra khạp 2 khạp...\n• Khách: Quá khạp mà nào đó... Giờ về đến đây để còn đợi khạp... Bắn hết cái đầu!\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales 1 (cúi chào): Nỉ hảo lảo bản!\n• Sales 2 (cúi chào): Hoan nghênh sếp lớn!\n• Khách (xoa cổ, đơ): Mỏi cổ quá rồi!\n• Khách: Anh nói tiếng Việt!\n• [ESCALATION 4-15s]\n• Sales 1: Ôi biết tiếng Việt!\n• Sales 2: Tưởng đối tác ngoại!\n• Sales 1: Bay từ Quảng Châu?\n• Khách: Bay về ký cont!\n• Sales 2: Mất tiền vé máy!\n• Sếp Thúy: Mở Portal ERP ngay!\n• Sếp Thúy: Báo giá ba mươi giây!\n• Sếp Thúy: Gom từ một khối!\n• Khách (ngớ người): Đặt online được á?\n• [PUNCHLINE 16-25s]\n• Sales 1: Bọn em cày ngoại ngữ!\n• Sales 2: Học cả tuần đón sếp!\n• Sales 1: Xin thực hành tiếp!\n• Sales 2: Nỉ hảo lảo bản!\n• Khách (đơ 1.5s): Ơ hay! Thôi ngay!\n• [DEFLATION 26-30s]\n• Khách: Về đặt app ngay!\n• Sếp Thúy (tỉnh bơ): Lần sau nhớ mở app!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Bác khách chủ xưởng (đeo kính, đứng mỏi cổ) + 2 Sales Eureka (áo đồng phục, chắp tay cúi chào liên tục) + Sếp Thúy (cầm tablet mở Portal ERP).",
        "scene": "Sảnh cửa kính & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, máy tính bảng mở Portal ERP, cốc trà, tập kẹp file A4.",
        "twist": "Hai Sales Eureka tưởng bác khách bay từ Quảng Châu sang là đối tác nước ngoài nên chắp tay cúi chào xổ tiếng Trung liên hoàn khiến bác mỏi cổ, mãi mới thanh minh mình nói tiếng Việt và bay về tận nơi ký cont. Sếp Thúy hướng dẫn tra báo giá 30 giây và đặt gom cont 1 khối online ngay trên Portal ERP; hai Sales vẫn nằng nặc đòi thực hành tiếp câu tiếng Trung mới học khiến bác khách đơ mặt xin về mở app ngay.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-77",
        "title": "Sếp Thúy thách đố Sổ tay AI tra mã HS và câu hỏi quỹ liên hoan khó đỡ (Tham khảo Reel 81K views)",
        "mb_raw": "• Sếp: Từ mai, hai chị em không phải đến ngân hàng làm việc nữa đâu!\n• Nhân viên 1: Úi, mai cho em nghỉ á? Chứ dạo này đông khách vất vả quá chị ạ! Thế gần giữa trưa cho em đi làm sếp nhé!\n• Sếp: Không! Ngoài kia cũng không phải đến! Hai chị em mà cứ không chịu phấn đấu thì AI nó cũng vượt mặt đấy!\n• Nhân viên 1: Úi! Làm sao mà vượt mặt được bọn em? Chứ thương bọn em, AI thì phá sản cho sếp mà xem!\n• Sếp: Mệt mỏi thật đấy, đấy là cực chân tay, tôi đang nói ở đây là cực trí tuệ ạ! Chứ thích hỏi AI, đây nó hỏi gì hỏi đi, cái gì nó cũng trả lời được đấy!\n• Nhân viên 1: Có cách nào không làm mà vẫn có ăn không?\n• Giọng AI: Cái này là do ăn ở, bạn thử hỏi sếp của bạn xem!\n• Sếp: Con này đã bảo hỏi sếp đấy đấy!\n• Nhân viên 1: Mà khách toàn hỏi linh tinh, hỏi cái gì liên quan nghiệp vụ số mà!\n• Nhân viên 2: Em cũng không hỏi số mà... Bạn ơi cho tôi hỏi: sếp tôi đưa cho tôi 10 triệu bảo cuối tuần đưa cả phòng đi ăn liên hoan. Hôm qua tôi đi Aeon tiêu hết 9 triệu rưỡi rồi. Làm cách nào để không phải trả tiền mà vẫn được đi ăn liên hoan và không bị sếp mắng?\n• (Sound: Aiyo mama! Cả phòng đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sếp Thúy: Mai hai đứa nghỉ nhé!\n• Nhân viên 1: Sếp cho nghỉ phép á?\n• Nhân viên 2: Mùa hàng vất vả quá!\n• Sếp Thúy: Không! AI sắp vượt mặt!\n• [ESCALATION 4-15s]\n• Nhân viên 1: AI biết gom cont?\n• Sếp Thúy: Tự động tám mươi phần!\n• Sếp Thúy: Tra mã HS mười giây!\n• Sếp Thúy: Thử hỏi AI xem nào!\n• Nhân viên 1 (hỏi AI): Không làm mà có ăn?\n• Giọng AI: Do ăn ở! Hỏi sếp!\n• Sếp Thúy: Hỏi chuyện nghiệp vụ ấy!\n• [PUNCHLINE 16-25s]\n• Nhân viên 2: Em hỏi câu thực tế!\n• Nhân viên 2: Quỹ phòng có mười triệu!\n• Nhân viên 2: Em lỡ tiêu chín rưỡi!\n• Nhân viên 2: Giờ liên hoan kiểu gì?\n• Nhân viên 2: Mà không bị sếp mắng?\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Chịu hai đứa luôn!\n• Sếp Thúy (tỉnh bơ): Hoàn quỹ ngay cho chị!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (cầm laptop/tablet mở Sổ tay AI) + 2 Nhân viên Gen Z (ngồi bàn làm việc máy tính, nét mặt tỉnh bơ).",
        "scene": "Góc bàn làm việc máy tính tại văn phòng Eureka 44 Triều Khúc, màn hình Sổ tay AI Eureka, kẹp hồ sơ mã HS & biểu thuế A4.",
        "twist": "Sếp Thúy răn đe nhân viên không phấn đấu thì Sổ tay AI (tự động hóa 80% quy trình lặp lại, tra mã HS trong 10 giây) sẽ vượt mặt, bảo hai nhân viên thử hỏi AI nghiệp vụ khó. Ai ngờ nhân viên lại hỏi AI cách không làm mà có ăn, rồi quay sang nhờ AI giải cứu quỹ liên hoan 10 triệu của phòng đã lỡ tiêu hết 9 triệu rưỡi mà không bị Sếp Thúy mắng.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-78",
        "title": "Đăng ký khóa tu im lặng cho Sếp Thúy tĩnh tâm và cú chốt khách bất ngờ (Tham khảo Reel 65K views)",
        "mb_raw": "• Sếp: Đứa nào đăng ký khóa tu mùa đông cho tôi đây? Cái nào đánh hả?\n• Nhân viên 1: Em chỉ đưa thông tin thôi!\n• Sếp: Thì làm hả?\n• Nhân viên 2: Em chỉ điền thông tin thôi!\n• Sếp: À thế tao là học khách hả?\n• Nhân viên 1: Em chỉ ấn gửi thôi!\n• Sếp: Thế này thì giỏi rồi! Chị chả bảo là tuyền nói nhiều là chị đau đầu còn gì! Giờ chị đến khóa tu ấy nhé người ta quy định 1 tuần không ai được nói câu gì luôn! Lại còn không được sử dụng điện thoại cơ, tha hồ mà tĩnh lặng! Chị có 1 tuần được an nhiên này! Một tuần không được nói câu nào, một tuần không được sử dụng điện thoại để tao phát điên á?!\n• Nhân viên: Ui, sếp! Khách đến nóng trường rồi kìa!\n• Sếp: Đi nhanh lên không bàn điện thoại tận người ta bán cho! Đồ bọn em chỉ đi kháy hết rồi đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sếp Thúy (cầm giấy): Ai đăng ký khóa tu?\n• Nhân viên 1: Em chỉ tìm thông tin!\n• Sếp Thúy: Rồi ai nộp đơn?\n• Nhân viên 2: Em chỉ ấn nút gửi!\n• [ESCALATION 4-15s]\n• Sếp Thúy: Tôi đi tu hả?\n• Nhân viên 1: Sếp kêu đau đầu lắm!\n• Nhân viên 2: Khách gọi suốt ngày đêm!\n• Sếp Thúy: Khóa tu bảy ngày im!\n• Sếp Thúy: Cấm nói không điện thoại!\n• Sếp Thúy: Ai chốt gom cont đây?\n• Sếp Thúy: Ai duyệt công nợ khách?\n• [PUNCHLINE 16-25s]\n• Nhân viên 1: Sếp tĩnh tâm an nhiên!\n• Sếp Thúy: Ngồi thiền mất khách á?\n• Nhân viên 2 (chỉ cửa): Sếp ơi! Khách lớn kìa!\n• Sếp Thúy (đơ 1.5s): Đâu? Khách đặt năm cont?\n• [DEFLATION 26-30s]\n• Sếp Thúy (ôm điện thoại): Hủy ngay khóa tu!\n• Sếp Thúy (tỉnh bơ): Làm ăn thế đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (cầm tờ giấy đăng ký khóa tu + điện thoại) + 2 Nhân viên Chứng từ/Sales (đứng khoanh tay tỉnh bơ).",
        "scene": "Bàn làm việc Sếp Thúy & quầy tiếp khách văn phòng Eureka 44 Triều Khúc, tờ thông báo khóa tu A4, điện thoại smartphone, tập hồ sơ C/O.",
        "twist": "Thấy Sếp Thúy than đau đầu vì trực điện thoại chốt đơn, hai nhân viên Gen Z phối hợp nhịp nhàng đăng ký cho Sếp khóa tu mùa đông 7 ngày im lặng tuyệt đối và cấm dùng điện thoại. Sếp Thúy đọc nội quy xong sững sờ hỏi ai chốt gom cont và duyệt công nợ 40%; đúng lúc khách lớn đặt 5 cont bước vào, Sếp Thúy ôm chặt điện thoại tuyên bố hủy ngay khóa tu.",
        "duration": "23s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-77",
        "motif": "Motif 5: Chơi chữ ngoại ngữ & Công nghệ số hóa",
        "title": "Học cấp tốc tiếng Trung đón sếp lớn và cái kết đặt đơn online trên Portal ERP (Tham khảo Reel 1.1M views)",
        "strength": "Hệ sinh thái Portal ERP v2.5.0 & Web erktransport.com tính cước 30 giây, đặt lệnh gom cont LCL từ 1 khối & theo dõi camera kho Bằng Tường - Quảng Châu online 24/7.",
        "roles_scene": "• Bối cảnh: Sảnh cửa kính & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Bác khách chủ xưởng + 2 Sales Eureka + Sếp Thúy\n• Đạo cụ: Máy tính bảng mở Portal ERP, cốc trà, tập kẹp file A4.",
        "pingpong": "• [HOOK 0-3s]\n• Sales 1 (cúi chào): Nỉ hảo lảo bản!\n• Sales 2 (cúi chào): Hoan nghênh sếp lớn!\n• Khách (xoa cổ, đơ): Mỏi cổ quá rồi!\n• Khách: Anh nói tiếng Việt!\n• [ESCALATION 4-15s]\n• Sales 1: Ôi biết tiếng Việt!\n• Sales 2: Tưởng đối tác ngoại!\n• Sales 1: Bay từ Quảng Châu?\n• Khách: Bay về ký cont!\n• Sales 2: Mất tiền vé máy!\n• Sếp Thúy: Mở Portal ERP ngay!\n• Sếp Thúy: Báo giá ba mươi giây!\n• Sếp Thúy: Gom từ một khối!\n• Khách (ngớ người): Đặt online được á?\n• [PUNCHLINE 16-25s]\n• Sales 1: Bọn em cày ngoại ngữ!\n• Sales 2: Học cả tuần đón sếp!\n• Sales 1: Xin thực hành tiếp!\n• Sales 2: Nỉ hảo lảo bản!\n• Khách (đơ 1.5s): Ơ hay! Thôi ngay!\n• [DEFLATION 26-30s]\n• Khách: Về đặt app ngay!\n• Sếp Thúy (tỉnh bơ): Lần sau nhớ mở app!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Hai Sales Eureka tưởng bác khách bay từ Quảng Châu sang là đối tác nước ngoài nên chắp tay cúi chào xổ tiếng Trung liên hoàn khiến bác mỏi cổ, mãi mới thanh minh mình nói tiếng Việt và bay về tận nơi ký cont. Sếp Thúy hướng dẫn tra báo giá 30 giây và đặt gom cont 1 khối online ngay trên Portal ERP; hai Sales vẫn nằng nặc đòi thực hành tiếp câu tiếng Trung mới học khiến bác khách đơ mặt xin về mở app ngay.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-78",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z & Sổ tay AI)",
        "title": "Sếp Thúy thách đố Sổ tay AI tra mã HS và câu hỏi quỹ liên hoan khó đỡ (Tham khảo Reel 81K views)",
        "strength": "Hệ thống Sổ tay AI Eureka hỗ trợ tự động hóa 80% quy trình lặp lại, giảm thời gian tra mã HS từ 15 phút xuống 10 giây & Gói Phòng XNK thuê ngoài 5 triệu/tháng.",
        "roles_scene": "• Bối cảnh: Góc bàn làm việc máy tính tại văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + 2 Nhân viên Gen Z\n• Đạo cụ: Laptop/tablet mở Sổ tay AI Eureka, kẹp hồ sơ mã HS & biểu thuế A4.",
        "pingpong": "• [HOOK 0-3s]\n• Sếp Thúy: Mai hai đứa nghỉ nhé!\n• Nhân viên 1: Sếp cho nghỉ phép á?\n• Nhân viên 2: Mùa hàng vất vả quá!\n• Sếp Thúy: Không! AI sắp vượt mặt!\n• [ESCALATION 4-15s]\n• Nhân viên 1: AI biết gom cont?\n• Sếp Thúy: Tự động tám mươi phần!\n• Sếp Thúy: Tra mã HS mười giây!\n• Sếp Thúy: Thử hỏi AI xem nào!\n• Nhân viên 1 (hỏi AI): Không làm mà có ăn?\n• Giọng AI: Do ăn ở! Hỏi sếp!\n• Sếp Thúy: Hỏi chuyện nghiệp vụ ấy!\n• [PUNCHLINE 16-25s]\n• Nhân viên 2: Em hỏi câu thực tế!\n• Nhân viên 2: Quỹ phòng có mười triệu!\n• Nhân viên 2: Em lỡ tiêu chín rưỡi!\n• Nhân viên 2: Giờ liên hoan kiểu gì?\n• Nhân viên 2: Mà không bị sếp mắng?\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Chịu hai đứa luôn!\n• Sếp Thúy (tỉnh bơ): Hoàn quỹ ngay cho chị!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Sếp Thúy răn đe nhân viên không phấn đấu thì Sổ tay AI (tự động hóa 80% quy trình lặp lại, tra mã HS trong 10 giây) sẽ vượt mặt, bảo hai nhân viên thử hỏi AI nghiệp vụ khó. Ai ngờ nhân viên lại hỏi AI cách không làm mà có ăn, rồi quay sang nhờ AI giải cứu quỹ liên hoan 10 triệu của phòng đã lỡ tiêu hết 9 triệu rưỡi mà không bị Sếp Thúy mắng.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-79",
        "motif": "Motif 3: Trào phúng tâm linh (Khóa tu im lặng) & Chăm sóc khách hàng 1-1",
        "title": "Đăng ký khóa tu im lặng cho Sếp Thúy tĩnh tâm và cú chốt khách bất ngờ (Tham khảo Reel 65K views)",
        "strength": "Dịch vụ Dedicated Team 1-1 (Quản lý + CSKH + Chứng từ) trực chiến 24/7, vận chuyển liên vận 1-2 ngày/xe, cấp công nợ 10%-40% & 100% hóa đơn VAT chính ngạch.",
        "roles_scene": "• Bối cảnh: Bàn làm việc Sếp Thúy & quầy tiếp khách văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + 2 Nhân viên Chứng từ/Sales\n• Đạo cụ: Tờ thông báo khóa tu A4, điện thoại smartphone, tập hồ sơ C/O.",
        "pingpong": "• [HOOK 0-3s]\n• Sếp Thúy (cầm giấy): Ai đăng ký khóa tu?\n• Nhân viên 1: Em chỉ tìm thông tin!\n• Sếp Thúy: Rồi ai nộp đơn?\n• Nhân viên 2: Em chỉ ấn nút gửi!\n• [ESCALATION 4-15s]\n• Sếp Thúy: Tôi đi tu hả?\n• Nhân viên 1: Sếp kêu đau đầu lắm!\n• Nhân viên 2: Khách gọi suốt ngày đêm!\n• Sếp Thúy: Khóa tu bảy ngày im!\n• Sếp Thúy: Cấm nói không điện thoại!\n• Sếp Thúy: Ai chốt gom cont đây?\n• Sếp Thúy: Ai duyệt công nợ khách?\n• [PUNCHLINE 16-25s]\n• Nhân viên 1: Sếp tĩnh tâm an nhiên!\n• Sếp Thúy: Ngồi thiền mất khách á?\n• Nhân viên 2 (chỉ cửa): Sếp ơi! Khách lớn kìa!\n• Sếp Thúy (đơ 1.5s): Đâu? Khách đặt năm cont?\n• [DEFLATION 26-30s]\n• Sếp Thúy (ôm điện thoại): Hủy ngay khóa tu!\n• Sếp Thúy (tỉnh bơ): Làm ăn thế đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Thấy Sếp Thúy than đau đầu vì trực điện thoại chốt đơn, hai nhân viên Gen Z phối hợp nhịp nhàng đăng ký cho Sếp khóa tu mùa đông 7 ngày im lặng tuyệt đối và cấm dùng điện thoại. Sếp Thúy đọc nội quy xong sững sờ hỏi ai chốt gom cont và duyệt công nợ 40%; đúng lúc khách lớn đặt 5 cont bước vào, Sếp Thúy ôm chặt điện thoại tuyên bố hủy ngay khóa tu.",
        "duration": "23s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 73,
        "title": "Học cấp tốc tiếng Trung đón sếp lớn và cái kết đặt đơn online trên Portal ERP (Reel 1.1M views)",
        "views": "1,100,000 views",
        "motif": "Motif 5: Chơi chữ ngoại ngữ & Công nghệ số hóa",
        "trigger": "Nhân viên tưởng khách là người nước ngoài nên cày cấp tốc khóa ngoại ngữ chắp tay cúi chào liên tục khiến khách mỏi cổ, đến lúc biết khách là người Việt có thể giao dịch online vẫn đòi thực hành tiếp câu vừa học.",
        "adaptation": "Hai Sales Eureka tưởng bác chủ xưởng bay từ Quảng Châu về là đối tác ngoại nên xổ tiếng Trung chào liên thanh; Sếp Thúy giới thiệu báo giá 30s và đặt gom cont 1 khối online trên Portal ERP nhưng hai Sales vẫn xin thực hành tiếp câu ngoại ngữ (EKA-V2-77)."
    },
    {
        "stt": 74,
        "title": "Sếp Thúy thách đố Sổ tay AI tra mã HS và câu hỏi quỹ liên hoan khó đỡ (Reel 81K views)",
        "views": "81,000 views",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z & Sổ tay AI)",
        "trigger": "Sếp dọa nhân viên không phấn đấu thì AI sẽ vượt mặt và thách thử hỏi AI nghiệp vụ, nhân viên quay sang hỏi AI câu hỏi bá đạo về cách ăn tiêu hết quỹ liên hoan của phòng mà không bị sếp mắng.",
        "adaptation": "Sếp Thúy khoe Sổ tay AI tự động hóa 80% quy trình lặp lại và tra mã HS trong 10 giây; hai nhân viên lại hỏi AI cách không làm mà có ăn và cách giải cứu quỹ liên hoan 10 triệu đã lỡ tiêu hết 9 triệu rưỡi (EKA-V2-78)."
    },
    {
        "stt": 75,
        "title": "Đăng ký khóa tu im lặng cho Sếp Thúy tĩnh tâm và cú chốt khách bất ngờ (Reel 65K views)",
        "views": "65,000 views",
        "motif": "Motif 3: Trào phúng tâm linh (Khóa tu im lặng) & Chăm sóc khách hàng 1-1",
        "trigger": "Nhân viên đùn đẩy trách nhiệm chuyện lén đăng ký cho sếp đi khóa tu mùa đông cấm nói và cấm điện thoại 1 tuần để sếp bớt đau đầu; sếp tá hỏa nhưng thấy khách lớn bước vào liền ôm chặt điện thoại làm việc ngay.",
        "adaptation": "Hai nhân viên đăng ký cho Sếp Thúy khóa tu im lặng 7 ngày không điện thoại; Sếp Thúy lo ai chốt gom cont 1 khối và duyệt công nợ 40%, thấy khách đặt 5 cont bước vào liền hủy ngay khóa tu (EKA-V2-79)."
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
