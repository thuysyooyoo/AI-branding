# -*- coding: utf-8 -*-
"""
Script: update_daily_viral_scripts_1111.py
Ngày thực hiện: 11/11/2026
Mục tiêu theo đúng SOP d:\\AI AGENT THUY\\AI branding\\HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md:
1. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 58, 59, 60) vào File 1: BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 59, 60, 61) vào File 2: BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx (Sheet 1 và Sheet 4)
3. Tham khảo các video Reels MBBank Hải Dương (>= 30K views):
   - Reel 1831704221158873 (192,000 views)
   - Reel 1424816669776931 (232,000 views)
   - Reel 1108856341859376 (82,000 views)
4. Tuân thủ 100% Rule 20 và banned_words.md (không từ cấm, không chửi bậy, Ping-Pong 2-5 từ, Deadpan 1.5s, 100% văn phòng Eureka 44 Triều Khúc)
5. Kỹ thuật atomic_save chống lock file / xung đột OneDrive
6. Tự động đồng bộ báo cáo sang Desktop TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx
"""

import os
import sys
import shutil
import subprocess
import openpyxl
from copy import copy

FILE_1 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx"
FILE_2 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx"

# Backup / local mirror file paths
LOCAL_FILE_1 = r"D:\AI AGENT THUY\AI branding\reports\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx"
LOCAL_FILE_2 = r"D:\AI AGENT THUY\AI branding\reports\BANG_10_KICH_BAN_VIRAL_EUREKA_NEW.xlsx"

# Dữ liệu 3 kịch bản mới cho File 1 (BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx)
NEW_SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-58",
        "title": "Hỏi cước năm cont máy phay CNC và cái kết gửi xe đi phượt chùa (Tham khảo Reel 192K views)",
        "mb_orig": """• Khách: Từ từ vào gần hàng, nhưng mà có được gửi xe miễn phí không?
• Em: Bên em anh cứ gửi thoải mái, miễn phí mà lãi suất lại còn đang ưu đãi nữa cơ ạ!
• Khách: Lãi suất ưu đãi như nào?
• Em: 50 triệu 8.2, 1 tỷ 9.1 ạ!
• Khách: Thế còn 5 tỷ? Anh... nhưng mà thôi cứ cho anh gửi cái xe ở đây. Gửi sổ tiền tỷ, rồi anh quay lại!
• Em: Dạ vâng anh đi lấy tiền ạ? Anh cứ đi thong thả!
... (3 ngày sau)
• Em: Ôi anh ơi, anh bảo quay lại mà để em đợi 3 ngày rồi đấy!
• Khách: Thì giờ anh quay lại rồi, không quay lại mất xe à!
• Em: Thế giờ anh gửi sổ bao nhiêu?
• Khách: Tiền nong cái gì! Vừa đi du lịch thong thả mấy ngày về! May quá bên em cho gửi xe miễn phí! Chứ bên kia anh hỏi rồi, 3 ngày mất trăm bảy!
• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)""",
        "eureka": """• [HOOK 0-3s]
• Khách (đội mũ phượt): Đỗ xe ở đây?
• Sales: Dạ miễn phí bác!
• Khách: Có an toàn không?
• Sales: Camera hai tư trên bảy!

• [ESCALATION 4-15s]
• Khách: Bác nhập năm cont!
• Khách: Máy phay CNC!
• Sales: Dạ cước siêu tốt!
• Sales: C/O Form E chuẩn!
• Khách: Có cấp công nợ?
• Sales: Cấp bốn mươi phần trăm!
• Khách: Khóa xe bác gửi!
• Khách: Bác đi lấy tiền!
• Sales: Bác đi thong thả!

• [PUNCHLINE 16-25s]
• (3 ngày sau, bác khách đeo ba lô phượt bước vào)
• Sales (mừng rỡ): Ôi bác quay lại!
• Sales: Hợp đồng năm cont?
• Khách (cười tươi): Cont cọc gì đâu!
• Khách: Bác đi Mộc Châu!
• Khách: Đi phượt ba ngày!
• Sales (ngã ngửa): Ủa đi phượt ạ?
• Khách: Gửi xe ở đây!
• Khách: Đỡ ba trăm nghìn!

• [DEFLATION 26-30s]
• Sếp Thúy (bước ra): Ơ hay bác ơi!
• Sếp Thúy: Làm ăn thế đấy!
• Sếp Thúy: Tưởng năm cont CNC...
• Sếp Thúy: Hóa ra gửi xe chùa?!
• (Sound: Bụp! Aiyo mama! Bác khách nhe răng cười, cả phòng đơ mặt 1.5s)""",
        "cast": "Sếp Thúy (ngồi bàn giám đốc quan sát) + 1 Sales nam Eureka (cầm catalogue máy CNC) + 1 Bác khách chủ xưởng (đội mũ bảo hiểm phượt, tay cầm chìa khóa xe máy, ba lô to đùng)",
        "setting": "Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, catalogue máy phay CNC, chìa khóa xe máy",
        "twist": "Tưởng khách sộp ghé văn phòng Triều Khúc ký hợp đồng nhập 5 cont máy phay CNC tiền tỷ, hóa ra giả vờ gửi xe máy nhờ văn phòng có camera 24/7 để đi phượt Mộc Châu 3 ngày đỡ mất 300k tiền gửi bãi.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-59",
        "title": "Tự hào chỉ huy ngàn cỗ máy tự động và cú lật ve chai phế liệu (Tham khảo Reel 232K views)",
        "mb_orig": """• Khách: Em có biết anh là ai không?
• Em: Dạ em không ạ?
• Khách: Anh lãnh đạo cả ngàn đầu quân! Anh cho chúng nó ăn chúng nó mới được ăn! Hôm nào anh thích lên nhớ, anh cầm đầu thằng nào là thằng đấy mất mạng với anh!
• Em: Sao em nghe cứ thấy mùi giang hồ thế anh? Dân giang hồ nhiều tiền lắm... Bên em có chương trình 50 triệu 8.2, 1 tỷ 9.1... Anh muốn gửi kỳ hạn nào?
• Khách: Gửi gì? Tiền đâu ra mà gửi!
• Em: Ơ em tưởng anh lãnh đạo cả ngàn quân?
• Khách: Thì cái trang trại lợn này, cả ngàn con chứ gì! Nhưng đợt bão vừa rồi nước ngập trôi mất sạch rồi! Ngân hàng có cho vay tiền không?
• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)""",
        "eureka": """• [HOOK 0-3s]
• Khách (đập bàn dõng dạc): Biết tôi là ai?
• Sales: Dạ chưa biết bác!
• Khách: Tôi chỉ huy nghìn quân!
• Sales (mắt tròn xoe): Nghìn quân cơ ạ?

• [ESCALATION 4-15s]
• Khách: Nghìn cỗ máy nghiền!
• Khách: Chạy suốt ngày đêm!
• Khách: Thích nghiền thằng nào...
• Khách: Nát vụn thằng đấy!
• Sales (toát mồ hôi): Xưởng cơ khí lớn?
• Khách: Lớn nhất miền Bắc!
• Sales: Eureka bao trọn gói!
• Sales: Phòng XNK năm triệu!
• Sales: Cấp công nợ lớn!
• Sales: Nhập dây chuyền gì?

• [PUNCHLINE 16-25s]
• Khách (gãi đầu cười hì): Tôi nhập ve chai!
• Khách: Máy xay xơ mướp!
• Khách: Đời chín mươi tám!
• Khách: Nát bét rỉ sét!
• Sales (đứng hình): Ủa máy ve chai?
• Khách: Về rã bán sắt!
• Khách: Cho tôi nợ cước!

• [DEFLATION 26-30s]
• Sếp Thúy (bước tới): Ơ hay bác ơi!
• Sếp Thúy: Làm ăn thế đấy!
• Sếp Thúy: Quyết định mười tám...
• Sếp Thúy: Máy nát ai cho nhập?!
• (Sound: Bụp! Aiyo mama! Cậu Sales ôm hồ sơ thở dài, cả phòng đơ mặt 1.5s)""",
        "cast": "Sếp Thúy (ngồi bàn làm việc, tay cầm bút ký) + 1 Sales nữ Eureka (cầm tập hồ sơ giám định A4) + 1 Bác khách (mặt ngầu, áo khoác đen, đập bàn thị uy)",
        "setting": "Bàn tiếp khách văn phòng Eureka 44 Triều Khúc, tập hồ sơ giám định máy móc theo Quyết định 18, catalogue máy công nghiệp",
        "twist": "Tự xưng chỉ huy cả nghìn cỗ máy nghiền nát tan tành như ông trùm cơ khí, hóa ra là chủ bãi thu mua ve chai phế liệu muốn Eureka cho nợ cước để nhập mấy cỗ máy xay xơ mướp nát đời 1998, bị Sếp Thúy gạt phắt vì vi phạm Quyết định 18/2019/QĐ-TTg.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-60",
        "title": "Hỏi tư vấn phòng XNK trọn gói và cái kết xin gói trà đợi xe bus (Tham khảo Reel 82K views)",
        "mb_orig": """• Khách: Vàng ngân hàng quân đội nhận buồn nhận miệng thế nhỉ? Cho xin cái kẹo em ơi! Cho xin cốc nước em ơi!
• Em: Dạ đây kẹo ngọt nước mát đây anh!
• Khách: Thế đợt này ngân hàng có chương trình gì không?
• Em: Dạ bên em có chương trình tổng giá trị giải thưởng tới 13.5 tỷ đồng, gửi từ 1 tỷ nhận quà lớn, khách Private từ 3 tỷ được điểm thưởng nửa tỷ... Bây giờ em làm sổ cho anh nhé?
• Khách: Ngân hàng uy tín, nhiều ưu đãi đấy... nhưng mà hạ cái điều hòa xuống một tí, rét lắm!
• Em: Dạ em hạ điều hòa rồi, mình làm sổ anh nhé?
• Khách: Không, làm gì có tiền mở sổ!
• Em: Ơ thế nãy giờ anh hỏi...
• Khách: Thấy ngân hàng quân đội nhiều chương trình thì vào hỏi xem có thật không! Nhưng mà lạnh rồi, thêm cho anh gói kẹo mang về cho cháu nhé!
• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)""",
        "eureka": """• [HOOK 0-3s]
• Khách (bốc kẹo lạc): Kẹo lạc ngon thế!
• Sales: Dạ mời bác xơi!
• Khách: Rót thêm chén trà!
• Sales: Trà Thái Nguyên bác!

• [ESCALATION 4-15s]
• Khách: Eureka có gói gì?
• Sales: Phòng XNK thuê ngoài!
• Khách: Bao nhiêu một tháng?
• Sales: Năm triệu trọn gói!
• Khách: Có gom cont lẻ?
• Sales: Từ một khối bác!
• Khách: Có cấp công nợ?
• Sales: Bốn mươi phần trăm!
• Sales: Bác ký hợp đồng?

• [PUNCHLINE 16-25s]
• Khách (phẩy tay cười khà): Ký cọt gì đâu!
• Sales (ngẩn ngơ): Dạ không ký ạ?
• Khách: Đang đợi xe bus!
• Khách: Tuyến ba mươi chín!
• Khách: Khát nước tạt vào!
• Khách: Uống nhờ chén trà!
• Khách: Cho xin gói kẹo!
• Khách: Mang về cho cháu!

• [DEFLATION 26-30s]
• Sếp Thúy (bước lại): Ơ hay bác ơi!
• Sếp Thúy: Làm ăn thế đấy!
• Sếp Thúy: Tư vấn nửa tiếng...
• Sếp Thúy: Xin kẹo đợi xe bus?!
• (Sound: Bụp! Aiyo mama! Bác khách cầm gói kẹo chạy vội ra cửa, cả phòng đơ mặt 1.5s)""",
        "cast": "Sếp Thúy (đứng bên máy in quan sát) + 1 Sales nam Eureka (rót nước trà) + 1 Bác khách (quần áo giản dị, ngồi vắt chân chữ ngũ, tay bốc kẹo lạc ăn ngon lành)",
        "setting": "Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, ấm chén trà Thái Nguyên, đĩa kẹo lạc, tập báo giá dịch vụ XNK A4",
        "twist": "Ngồi sofa thưởng trà xơi sạch đĩa kẹo lạc, hỏi han dồn dập gói phòng XNK 5 triệu và công nợ 40% như đại gia chuẩn bị ký hợp đồng, hóa ra đứng đợi xe bus tuyến 39 khát nước tạt vào uống chùa chén trà và xin thêm gói kẹo mang về cho cháu.",
        "duration": "22s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx)
NEW_SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-59",
        "motif": "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Bác khách gửi xe đi phượt",
        "title": "Hỏi cước năm cont máy phay CNC và cái kết gửi xe đi phượt chùa (Tham khảo Reel 192K views)",
        "canvas": "Nhập khẩu chính ngạch máy phay CNC, tư vấn C/O Form E miễn thuế 0%, cấp hạn mức công nợ 10%-40%, hệ thống camera giám sát văn phòng kho bãi an toàn 24/7.",
        "cast_setting": "• Bối cảnh: Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, catalogue máy phay CNC, chìa khóa xe máy\n• Nhân vật: Sếp Thúy (ngồi bàn giám đốc quan sát), 1 Sales nam Eureka (cầm catalogue máy CNC), 1 Bác khách chủ xưởng (đội mũ bảo hiểm phượt, tay cầm chìa khóa xe máy, ba lô to đùng)",
        "script": """• [HOOK 0-3s]
• Khách (đội mũ phượt): Đỗ xe ở đây?
• Sales: Dạ miễn phí bác!
• Khách: Có an toàn không?
• Sales: Camera hai tư trên bảy!

• [ESCALATION 4-15s]
• Khách: Bác nhập năm cont!
• Khách: Máy phay CNC!
• Sales: Dạ cước siêu tốt!
• Sales: C/O Form E chuẩn!
• Khách: Có cấp công nợ?
• Sales: Cấp bốn mươi phần trăm!
• Khách: Khóa xe bác gửi!
• Khách: Bác đi lấy tiền!
• Sales: Bác đi thong thả!

• [PUNCHLINE 16-25s]
• (3 ngày sau, bác khách đeo ba lô phượt bước vào)
• Sales (mừng rỡ): Ôi bác quay lại!
• Sales: Hợp đồng năm cont?
• Khách (cười tươi): Cont cọc gì đâu!
• Khách: Bác đi Mộc Châu!
• Khách: Đi phượt ba ngày!
• Sales (ngã ngửa): Ủa đi phượt ạ?
• Khách: Gửi xe ở đây!
• Khách: Đỡ ba trăm nghìn!

• [DEFLATION 26-30s]
• Sếp Thúy (bước ra): Ơ hay bác ơi!
• Sếp Thúy: Làm ăn thế đấy!
• Sếp Thúy: Tưởng năm cont CNC...
• Sếp Thúy: Hóa ra gửi xe chùa?!
• (Sound: Bụp! Aiyo mama! Bác khách nhe răng cười, cả phòng đơ mặt 1.5s)""",
        "punchline": "Tưởng khách sộp ghé văn phòng Triều Khúc ký hợp đồng nhập 5 cont máy phay CNC tiền tỷ, hóa ra giả vờ gửi xe máy nhờ văn phòng có camera 24/7 để đi phượt Mộc Châu 3 ngày đỡ mất 300k tiền gửi bãi.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-60",
        "motif": "Motif 2: Drama nội bộ công sở & Đấu trí hồ sơ giám định máy móc cũ Quyết định 18",
        "title": "Tự hào chỉ huy ngàn cỗ máy tự động và cú lật ve chai phế liệu (Tham khảo Reel 232K views)",
        "canvas": "Tư vấn nghiệp vụ Giám định máy móc thiết bị cũ theo Quyết định 18/2019/QĐ-TTg (tuổi thiết bị dưới 10 năm, quy chuẩn an toàn, 100% hóa đơn chứng từ chính ngạch, từ chối hàng phế liệu ve chai).",
        "cast_setting": "• Bối cảnh: Bàn tiếp khách văn phòng Eureka 44 Triều Khúc, tập hồ sơ giám định máy móc theo Quyết định 18, catalogue máy công nghiệp\n• Nhân vật: Sếp Thúy (ngồi bàn làm việc, tay cầm bút ký), 1 Sales nữ Eureka (cầm tập hồ sơ giám định A4), 1 Bác khách (mặt ngầu, áo khoác đen, đập bàn thị uy)",
        "script": """• [HOOK 0-3s]
• Khách (đập bàn dõng dạc): Biết tôi là ai?
• Sales: Dạ chưa biết bác!
• Khách: Tôi chỉ huy nghìn quân!
• Sales (mắt tròn xoe): Nghìn quân cơ ạ?

• [ESCALATION 4-15s]
• Khách: Nghìn cỗ máy nghiền!
• Khách: Chạy suốt ngày đêm!
• Khách: Thích nghiền thằng nào...
• Khách: Nát vụn thằng đấy!
• Sales (toát mồ hôi): Xưởng cơ khí lớn?
• Khách: Lớn nhất miền Bắc!
• Sales: Eureka bao trọn gói!
• Sales: Phòng XNK năm triệu!
• Sales: Cấp công nợ lớn!
• Sales: Nhập dây chuyền gì?

• [PUNCHLINE 16-25s]
• Khách (gãi đầu cười hì): Tôi nhập ve chai!
• Khách: Máy xay xơ mướp!
• Khách: Đời chín mươi tám!
• Khách: Nát bét rỉ sét!
• Sales (đứng hình): Ủa máy ve chai?
• Khách: Về rã bán sắt!
• Khách: Cho tôi nợ cước!

• [DEFLATION 26-30s]
• Sếp Thúy (bước tới): Ơ hay bác ơi!
• Sếp Thúy: Làm ăn thế đấy!
• Sếp Thúy: Quyết định mười tám...
• Sếp Thúy: Máy nát ai cho nhập?!
• (Sound: Bụp! Aiyo mama! Cậu Sales ôm hồ sơ thở dài, cả phòng đơ mặt 1.5s)""",
        "punchline": "Tự xưng chỉ huy cả nghìn cỗ máy nghiền nát tan tành như ông trùm cơ khí, hóa ra là chủ bãi thu mua ve chai phế liệu muốn Eureka cho nợ cước để nhập mấy cỗ máy xay xơ mướp nát đời 1998, bị Sếp Thúy gạt phắt vì vi phạm Quyết định 18/2019/QĐ-TTg.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-61",
        "motif": "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Bác khách ké kẹo lạc đợi xe bus",
        "title": "Hỏi tư vấn phòng XNK trọn gói và cái kết xin gói trà đợi xe bus (Tham khảo Reel 82K views)",
        "canvas": "Phòng Xuất nhập khẩu thuê ngoài trọn gói chỉ từ 5 triệu/tháng, gom cont LCL từ 1 khối xe chạy hàng ngày, cấp hạn mức công nợ 10%-40%, tư vấn C/O Form E miễn thuế 0%.",
        "cast_setting": "• Bối cảnh: Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, ấm chén trà Thái Nguyên, đĩa kẹo lạc, tập báo giá dịch vụ XNK A4\n• Nhân vật: Sếp Thúy (đứng bên máy in quan sát), 1 Sales nam Eureka (rót nước trà), 1 Bác khách (quần áo giản dị, ngồi vắt chân chữ ngũ, tay bốc kẹo lạc ăn ngon lành)",
        "script": """• [HOOK 0-3s]
• Khách (bốc kẹo lạc): Kẹo lạc ngon thế!
• Sales: Dạ mời bác xơi!
• Khách: Rót thêm chén trà!
• Sales: Trà Thái Nguyên bác!

• [ESCALATION 4-15s]
• Khách: Eureka có gói gì?
• Sales: Phòng XNK thuê ngoài!
• Khách: Bao nhiêu một tháng?
• Sales: Năm triệu trọn gói!
• Khách: Có gom cont lẻ?
• Sales: Từ một khối bác!
• Khách: Có cấp công nợ?
• Sales: Bốn mươi phần trăm!
• Sales: Bác ký hợp đồng?

• [PUNCHLINE 16-25s]
• Khách (phẩy tay cười khà): Ký cọt gì đâu!
• Sales (ngẩn ngơ): Dạ không ký ạ?
• Khách: Đang đợi xe bus!
• Khách: Tuyến ba mươi chín!
• Khách: Khát nước tạt vào!
• Khách: Uống nhờ chén trà!
• Khách: Cho xin gói kẹo!
• Khách: Mang về cho cháu!

• [DEFLATION 26-30s]
• Sếp Thúy (bước lại): Ơ hay bác ơi!
• Sếp Thúy: Làm ăn thế đấy!
• Sếp Thúy: Tư vấn nửa tiếng...
• Sếp Thúy: Xin kẹo đợi xe bus?!
• (Sound: Bụp! Aiyo mama! Bác khách cầm gói kẹo chạy vội ra cửa, cả phòng đơ mặt 1.5s)""",
        "punchline": "Ngồi sofa thưởng trà xơi sạch đĩa kẹo lạc, hỏi han dồn dập gói phòng XNK 5 triệu và công nợ 40% như đại gia chuẩn bị ký hợp đồng, hóa ra đứng đợi xe bus tuyến 39 khát nước tạt vào uống chùa chén trà và xin thêm gói kẹo mang về cho cháu.",
        "duration": "22s"
    }
]

# Gợi ý ý tưởng cho Sheet 4 của File 2 (GOI_Y_Y_TUONG_MB)
NEW_IDEAS_SHEET_4 = [
    (
        "Hỏi cước năm cont máy phay CNC và cái kết gửi xe đi phượt chùa (Reel 192K views)",
        "192,000 views",
        "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Bác khách gửi xe đi phượt",
        "Khách vào ngân hàng hỏi gửi xe miễn phí, làm bộ hỏi gửi tiết kiệm 5 tỷ để gửi xe nhờ, 3 ngày sau mới quay lại bảo đi du lịch thong thả về, gửi xe ở ngân hàng để khỏi mất trăm bảy tiền bãi xe.",
        "Khách hỏi cước 5 cont máy phay CNC và làm C/O Form E miễn thuế, xin gửi nhờ xe máy ở văn phòng Triều Khúc. 3 ngày sau quay lại lấy xe bảo vừa đi phượt Mộc Châu về, gửi xe ở Eureka vừa có camera 24/7 vừa đỡ mất 300k tiền bãi (EKA-V2-59)"
    ),
    (
        "Tự hào chỉ huy ngàn cỗ máy tự động và cú lật ve chai phế liệu (Reel 232K views)",
        "232,000 views",
        "Motif 2: Drama nội bộ công sở & Đấu trí hồ sơ giám định máy móc cũ Quyết định 18",
        "Khách đập bàn tự hào lãnh đạo cả ngàn quân sinh sát trong tay, nhân viên tưởng đại ca xã hội đen giàu có tư vấn gói tiết kiệm tiền tỷ, hóa ra là chủ trại nuôi lợn bị bão trôi mất sạch, đến ngân hàng xin vay tiền.",
        "Khách vỗ ngực chỉ huy nghìn cỗ máy tự động chạy ngày đêm, tưởng chủ xưởng cơ khí đại gia, hóa ra chủ bãi ve chai muốn Eureka cho nợ cước nhập mấy cái máy xay xơ mướp rỉ sét đời 1998, bị Sếp Thúy gạt phắt vì vi phạm Quyết định 18 (EKA-V2-60)"
    ),
    (
        "Hỏi tư vấn phòng XNK trọn gói và cái kết xin gói trà đợi xe bus (Reel 82K views)",
        "82,000 views",
        "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Bác khách ké kẹo lạc đợi xe bus",
        "Khách vào ngân hàng xin kẹo ăn, xin nước ngọt uống, hỏi dồn dập các chương trình trúng thưởng 13.5 tỷ và khách VIP Private, đến lúc làm sổ bảo không có tiền, vào hỏi cho biết và xin thêm kẹo mang về cho cháu.",
        "Khách ngồi sofa văn phòng 44 Triều Khúc ăn sạch đĩa kẹo lạc, uống 3 chén trà Thái Nguyên, hỏi han phòng XNK 5 triệu và gom cont LCL từ 1 khối, nhân viên mang hợp đồng ra thì bảo đang đợi xe bus tuyến 39 khát nước tạt vào uống chùa và xin gói kẹo mang về (EKA-V2-61)"
    )
]

def atomic_save(wb, target_path):
    """Lưu workbook nguyên tử tránh lỗi lock file hoặc xung đột OneDrive."""
    tmp_path = target_path + ".tmp.xlsx"
    wb.save(tmp_path)
    wb.close()
    if os.path.exists(target_path):
        os.replace(tmp_path, target_path)
    else:
        shutil.move(tmp_path, target_path)
    print(f"[✓] Đã lưu thành công (atomic): {target_path}")

def update_file_1():
    print(f"\n[*] Đang cập nhật File 1: {FILE_1}")
    wb = openpyxl.load_workbook(FILE_1)
    ws = wb['KICH_BAN_EUREKA_CHUAN_MB']
    
    # 1. Kiểm tra kịch bản có note 'Đã sản xuất' để xóa nếu có
    rows_to_delete = []
    for r in range(4, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            val = str(ws.cell(r, c).value or "").lower()
            if "đã sản xuất" in val:
                rows_to_delete.append(r)
                break
    for r in sorted(rows_to_delete, reverse=True):
        print(f"[-] Xóa kịch bản đã sản xuất ở hàng {r}")
        ws.delete_rows(r)
        
    # 2. Tìm hàng cuối và STT cuối
    last_row = ws.max_row
    while last_row > 3 and ws.cell(last_row, 1).value is None and ws.cell(last_row, 2).value is None:
        last_row -= 1
        
    last_stt = 0
    for r in range(last_row, 3, -1):
        val = ws.cell(r, 1).value
        if isinstance(val, int):
            last_stt = val
            break
            
    print(f"[*] File 1 hàng cuối hiện tại: {last_row}, STT cuối: {last_stt}")
    ref_row = last_row
    
    for item in NEW_SCRIPTS_FILE_1:
        last_stt += 1
        new_row = ws.max_row + 1
        
        ws.cell(new_row, 1, last_stt)
        ws.cell(new_row, 2, item["code"])
        ws.cell(new_row, 3, item["title"])
        ws.cell(new_row, 4, item["mb_orig"])
        ws.cell(new_row, 5, item["eureka"])
        ws.cell(new_row, 6, item["cast"])
        ws.cell(new_row, 7, item["setting"])
        ws.cell(new_row, 8, item["twist"])
        ws.cell(new_row, 9, item["duration"])
        
        # Sao chép định dạng
        for col in range(1, 10):
            target_cell = ws.cell(new_row, col)
            ref_cell = ws.cell(ref_row, col)
            if ref_cell.font: target_cell.font = copy(ref_cell.font)
            if ref_cell.alignment: target_cell.alignment = copy(ref_cell.alignment)
            if ref_cell.border: target_cell.border = copy(ref_cell.border)
            if ref_cell.fill: target_cell.fill = copy(ref_cell.fill)
            
        print(f"[+] Thêm vào File 1: STT {last_stt} - {item['code']} - {item['title'][:40]}...")

    atomic_save(wb, FILE_1)
    
    # Sync mirror local nếu có
    if os.path.exists(os.path.dirname(LOCAL_FILE_1)):
        try:
            shutil.copy2(FILE_1, LOCAL_FILE_1)
            print(f"[✓] Đã sao lưu cục bộ vào: {LOCAL_FILE_1}")
        except Exception as e:
            print(f"[!] Không thể sao lưu mirror File 1: {e}")

def update_file_2():
    print(f"\n[*] Đang cập nhật File 2: {FILE_2}")
    wb = openpyxl.load_workbook(FILE_2)
    
    # --- Cập nhật Sheet 1: KICH_BAN_VIRAL_REELS ---
    ws1 = wb['KICH_BAN_VIRAL_REELS']
    
    # 1. Kiểm tra kịch bản có note 'Đã sản xuất'
    rows_to_delete = []
    for r in range(5, ws1.max_row + 1):
        for c in range(1, ws1.max_column + 1):
            val = str(ws1.cell(r, c).value or "").lower()
            if "đã sản xuất" in val:
                rows_to_delete.append(r)
                break
    for r in sorted(rows_to_delete, reverse=True):
        print(f"[-] Xóa kịch bản đã sản xuất ở hàng {r}")
        ws1.delete_rows(r)
        
    last_row = ws1.max_row
    while last_row > 4 and ws1.cell(last_row, 1).value is None and ws1.cell(last_row, 2).value is None:
        last_row -= 1
        
    last_stt = 0
    for r in range(last_row, 4, -1):
        val = ws1.cell(r, 1).value
        if isinstance(val, int):
            last_stt = val
            break
            
    print(f"[*] Sheet 1 hàng cuối hiện tại: {last_row}, STT cuối: {last_stt}")
    ref_row = last_row
    
    for item in NEW_SCRIPTS_FILE_2:
        last_stt += 1
        new_row = ws1.max_row + 1
        
        ws1.cell(new_row, 1, last_stt)
        ws1.cell(new_row, 2, item["code"])
        ws1.cell(new_row, 3, item["motif"])
        ws1.cell(new_row, 4, item["title"])
        ws1.cell(new_row, 5, item["canvas"])
        ws1.cell(new_row, 6, item["cast_setting"])
        ws1.cell(new_row, 7, item["script"])
        ws1.cell(new_row, 8, item["punchline"])
        ws1.cell(new_row, 9, item["duration"])
        
        for col in range(1, 10):
            target_cell = ws1.cell(new_row, col)
            ref_cell = ws1.cell(ref_row, col)
            if ref_cell.font: target_cell.font = copy(ref_cell.font)
            if ref_cell.alignment: target_cell.alignment = copy(ref_cell.alignment)
            if ref_cell.border: target_cell.border = copy(ref_cell.border)
            if ref_cell.fill: target_cell.fill = copy(ref_cell.fill)
            
        print(f"[+] Thêm vào Sheet 1: STT {last_stt} - {item['code']} - {item['title'][:40]}...")

    # --- Cập nhật Sheet 4: GOI_Y_Y_TUONG_MB ---
    ws4 = wb['GOI_Y_Y_TUONG_MB']
    last_row_s4 = ws4.max_row
    while last_row_s4 > 3 and ws4.cell(last_row_s4, 1).value is None and ws4.cell(last_row_s4, 2).value is None:
        last_row_s4 -= 1
        
    last_stt_s4 = 0
    for r in range(last_row_s4, 3, -1):
        val = ws4.cell(r, 1).value
        if isinstance(val, int):
            last_stt_s4 = val
            break
    print(f"[*] Sheet 4 hàng cuối: {last_row_s4}, STT cuối: {last_stt_s4}")
    ref_row_s4 = last_row_s4
    
    for idea in NEW_IDEAS_SHEET_4:
        last_stt_s4 += 1
        new_row_s4 = ws4.max_row + 1
        
        ws4.cell(new_row_s4, 1, last_stt_s4)
        ws4.cell(new_row_s4, 2, idea[0])
        ws4.cell(new_row_s4, 3, idea[1])
        ws4.cell(new_row_s4, 4, idea[2])
        ws4.cell(new_row_s4, 5, idea[3])
        ws4.cell(new_row_s4, 6, idea[4])
        
        for col in range(1, 7):
            target_cell = ws4.cell(new_row_s4, col)
            ref_cell = ws4.cell(ref_row_s4, col)
            if ref_cell.font: target_cell.font = copy(ref_cell.font)
            if ref_cell.alignment: target_cell.alignment = copy(ref_cell.alignment)
            if ref_cell.border: target_cell.border = copy(ref_cell.border)
            if ref_cell.fill: target_cell.fill = copy(ref_cell.fill)
            
        print(f"[+] Thêm vào Sheet 4: STT {last_stt_s4} - {idea[0][:40]}...")

    atomic_save(wb, FILE_2)
    
    # Sync mirror local nếu có
    if os.path.exists(os.path.dirname(LOCAL_FILE_2)):
        try:
            shutil.copy2(FILE_2, LOCAL_FILE_2)
            print(f"[✓] Đã sao lưu cục bộ vào: {LOCAL_FILE_2}")
        except Exception as e:
            print(f"[!] Không thể sao lưu mirror File 2: {e}")

def sync_desktop_report():
    print("\n[*] Đang kích hoạt generator_chi_dan.py để đồng bộ Desktop...")
    script_path = r"D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py"
    if os.path.exists(script_path):
        res = subprocess.run(["python", script_path], capture_output=True, text=True, timeout=60)
        print(res.stdout)
        if res.returncode == 0:
            print("[✓] Đã đồng bộ thành công vào Bảng chỉ dẫn Desktop!")
        else:
            print(f"[!] Lỗi khi chạy generator_chi_dan: {res.stderr}")
    else:
        print(f"[!] Không tìm thấy script: {script_path}")

if __name__ == "__main__":
    update_file_1()
    update_file_2()
    sync_desktop_report()
    print("\n[✓] HOÀN TẤT TOÀN BỘ QUY TRÌNH LẬP KỊCH BẢN VIRAL HÀNG NGÀY THEO CHUẨN!")
