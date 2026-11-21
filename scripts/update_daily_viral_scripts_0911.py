# -*- coding: utf-8 -*-
"""
Script: update_daily_viral_scripts_0911.py
Ngày thực hiện: 09/11/2026
Mục tiêu:
1. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 52, 53, 54) vào File 1: BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 53, 54, 55) vào File 2: BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx (Sheet 1 và Sheet 4)
3. Tham khảo các video Reels MBBank Hải Dương (>= 30K views):
   - Reel 1500204105497669 (354,000 views)
   - Reel 2314073276081828 (84,000 views)
   - Reel 1107074315581312 (254,000 views)
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
        "code": "EKA-REEL-52",
        "title": "Tặng giỏ quà tri ân đối tác và cái kết bà xã vào viện cấp cứu (Tham khảo Reel 354K views)",
        "mb_orig": "• Khách: Cho anh rút tiền! Rút hết!\n• Em: Ôi anh ơi, lãi suất đang cao mà... 50 triệu 8.2...\n• Khách: Biết là cao rồi nhưng mà con vợ nó cứ giãy đành đạch bắt anh rút tiền về!\n• Em: Rõ ràng hôm qua em còn chu đáo tặng hoa 20 tháng 10 cho chị cơ mà!\n• Khách: Sau cái lúc mà em tặng hoa ấy là lúc nó phải nhập viện! Lúc cứ giãy đành đạch bắt anh rút tiền về!\n• Em: Ủa mà em tặng hoa gì đấy?\n• Khách: Hoa sữa! Ông hàng xóm mới chặt nguyên cả cành... Em còn cẩn thận rải khắp nhà cho tha hồ man mác mùa thu Hà Nội! Bảo sao nó chả nhập viện! Thế thôi không cần rút tiền nữa, nó mà nằm im bất động thì không sao, nó mà tỉnh dậy nhớ sang mà trả viện phí!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka": "• [HOOK 0-3s]\n• Khách (đập bàn thở dốc): Hủy hợp đồng!\n• Khách: Trả cọc cho anh!\n• Sales: Dạ hủy lô nào?\n• Khách: Dây chuyền dập tôn!\n\n• [ESCALATION 4-15s]\n• Sales: Đang gom cont LCL!\n• Sales: Tiết kiệm tiền tỷ!\n• Khách: Vợ anh bắt hủy!\n• Khách: Bắt rút sạch cọc!\n• Sales: Ủa lạ thế bác?\n• Sales: Hôm qua em sang!\n• Sales: Tặng giỏ quà VIP!\n• Sales: Chị nhà thích lắm!\n• Khách (quát to): Thích cái nỗi gì!\n• Khách: Tặng xong đi viện!\n• Sales: Đi viện làm sao?\n• Khách: Dị ứng cấp cứu!\n• Sales (toát mồ hôi): Em tặng gì đâu?\n• Sales: Giỏ quà tri ân!\n\n• [PUNCHLINE 16-25s]\n• Khách: Tinh dầu hồi quế!\n• Khách: Chú dặn xông xưởng!\n• Khách: Đốt cả phòng ngủ!\n• Khách: Cay xè cả mắt!\n• Khách: Vợ sưng húp mặt!\n• Sales (mếu máo): Xông xưởng cơ mà...\n• Khách: Bà ấy vừa dặn!\n• Khách: Không sang trả viện phí...\n• Khách: Đừng hòng nhập cont!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Tinh dầu đuổi muỗi...\n• Sếp Thúy: Ai cho mang tặng?!\n• (Sound: Aiyo mama! Cậu Sales ôm đầu gãi tai, cả phòng đơ mặt 1.5s)",
        "cast": "Sếp Thúy (ngồi bàn giám đốc quan sát) + 1 Sales nam Eureka + 1 Bác chủ xưởng (đùng đùng xông vào, mặt mũi bơ phờ, tay cầm chiếc gối ôm bệnh viện hoặc đơn thuốc)",
        "setting": "Bàn tiếp khách văn phòng Eureka 44 Triều Khúc, tập hợp đồng ủy thác nhập khẩu A4, giỏ quà tri ân",
        "twist": "Bác chủ xưởng đòi hủy cọc cont máy dập tôn vì vợ bắt rút sau khi nhận quà tri ân. Hóa ra Sales tặng chai tinh dầu quế dặn xông xưởng, khách đốt cả phòng ngủ làm vợ dị ứng sưng mặt nhập viện bắt Sales sang trả tiền viện phí.",
        "duration": "25s"
    },
    {
        "code": "EKA-REEL-53",
        "title": "Tuyển chọn nhân sự trực Portal ERP cuối tuần và cái kết của hội F.A (Tham khảo Reel 84K views)",
        "mb_orig": "• Quản lý: Cuối tuần này đứa nào chưa có người yêu nhỉ?\n• Nhân viên 1: Em!\n• Quản lý: Thế cuối tuần này đứa nào chưa được ai dẫn đi chơi nhỉ?\n• Nhân viên 2: Em!\n• Quản lý: Tốt! Chị Phân ơi, ghi tên hai đứa này cho em cuối tuần tăng ca nhé!\n• Quản lý: Không có người yêu, cuối tuần về sớm làm gì cho nó tủi thân ra!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka": "• [HOOK 0-3s]\n• Sếp Thúy (bước tới quầy): Ai chưa người yêu?\n• Sales Nam (giơ tay nhanh): Dạ em Sếp!\n• Sales Nữ (giơ tay theo): Em cũng ế Sếp!\n\n• [ESCALATION 4-15s]\n• Sếp Thúy (mỉm cười dịu dàng): Tối nay rảnh không?\n• Sales Nam: Dạ rảnh lắm Sếp!\n• Sếp Thúy: Có ai rủ đi?\n• Sales Nữ: Không ai rủ luôn!\n• Sếp Thúy: Thích ngắm cảnh đẹp?\n• Sales Nam (mắt sáng rực): Dạ thích lắm Sếp!\n• Sales Nữ: Sếp dẫn đi chơi?\n• Sếp Thúy: Chuẩn luôn hai đứa!\n• Sếp Thúy: Cảnh quốc tế luôn!\n\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (chỉ laptop): Đây hai đứa!\n• Sếp Thúy: Trực app Portal ERP!\n• Sếp Thúy: Camera kho Bằng Tường!\n• Sếp Thúy: Hai mươi bốn trên bảy!\n• Sales Nam (đứng hình): Ủa trực ca đêm?\n• Sếp Thúy (tỉnh bơ): Không có người yêu...\n• Sếp Thúy: Về sớm tủi thân!\n• Sếp Thúy: Ngắm cont cho vui!\n\n• [DEFLATION 26-30s]\n• Sales Nữ (nhìn màn hình): Ơ hay Sếp ơi!\n• Sales Nữ: Máy nâng container...\n• Sales Nữ: Sao lãng mạn được!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Trực xong báo giá...\n• Sếp Thúy: Ba mươi giây nghe!\n• (Sound: Aiyo mama! Hai bạn Sales nhìn nhau mếu máo, cả phòng đơ mặt 1.5s)",
        "cast": "Sếp Thúy (phong thái lãnh đạo, mỉm cười quan tâm) + 2 Sales trẻ (1 nam, 1 nữ ngồi bấm máy tính)",
        "setting": "Bàn làm việc văn phòng Eureka 44 Triều Khúc, máy tính laptop, lịch để bàn, cốc trà",
        "twist": "Sếp Thúy ân cần hỏi han hội nhân viên F.A rảnh rỗi chưa có người yêu tưởng được Sếp bao đi chơi, ai ngờ được Sếp giao cho trực camera kho Bằng Tường trên app Portal ERP cả đêm cho đỡ tủi thân.",
        "duration": "22s"
    },
    {
        "code": "EKA-REEL-54",
        "title": "Hỏi book năm cont máy móc đòi mười bộ quà tri ân và cú quay xe (Tham khảo Reel 254K views)",
        "mb_orig": "• Khách: Gửi tiết kiệm được tặng một lọ nước hoa thơm phức đúng không?\n• Em: Dạ đúng rồi anh ạ! Từ một tỷ được tặng một lọ nước hoa cao cấp anh ạ!\n• Khách: Năm tỷ phải đủ năm lọ nhé!\n• Em: Bên em anh yên tâm, em tặng hẳn mười lọ luôn! Lãi suất lại đang ưu đãi nữa anh ạ, 50 triệu 8.2, 1 tỷ 9.1... Cứ làm sổ đi!\n• Khách: Rồi anh đâu anh nhận quà trước nhé! Người thơm phức rồi, phải đủ mười lọ đấy nhé!\n• Em: Ơ thế tiền đâu để làm sổ anh?\n• Khách: Làm gì có tiền, lấy quà trước đã!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka": "• [HOOK 0-3s]\n• Khách (mở túi du lịch): Eureka tặng quà VIP?\n• Sales: Dạ đúng rồi bác!\n• Khách: Quà cho khách lớn?\n• Sales: Dạ khách ký hợp đồng!\n\n• [ESCALATION 4-15s]\n• Khách: Bác book năm cont!\n• Khách: Dây chuyền cơ khí!\n• Khách: Phải đủ năm bộ!\n• Sales: Năm cont lớn thế?\n• Sales: Em xin Sếp Thúy!\n• Sales: Tặng hẳn mười bộ!\n• Sales: Áo mưa bình nước!\n• Sales: Ô dù cao cấp!\n• Khách (vơ vào túi): Lấy mười bộ nhé!\n• Khách: Cho cả xưởng dùng!\n• Sales: Dạ hợp đồng đây!\n• Sales: Bác ký hợp đồng!\n• Sales: Cọc ba mươi phần trăm!\n\n• [PUNCHLINE 16-25s]\n• Khách (kéo khóa túi lại): Cọc cạch gì đâu!\n• Sales (ngơ ngác): Dạ không cọc?\n• Khách: Bác tham khảo cước!\n• Khách: Năm sau mới nhập!\n• Khách: Lấy quà về trước!\n• Sales (chới với): Ơ bác trả quà!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (chặn cửa): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Năm sau mới nhập...\n• Sếp Thúy: Khuân sạch kho quà?!\n• (Sound: Aiyo mama! Bác khách ôm túi quà cười gượng, cả phòng đơ mặt 1.5s)",
        "cast": "Sếp Thúy (đứng kiểm tra hồ sơ) + 1 Sales nữ Eureka + 1 Bác khách hàng (tay cầm túi du lịch to, kính râm cài ngực, tác phong nhanh nhẹn)",
        "setting": "Quầy lễ tân/tiếp khách văn phòng 44 Triều Khúc, thùng quà tặng Eureka (áo mưa, bình nước, mũ bảo hiểm), tập hợp đồng ủy thác",
        "twist": "Khách khoe book 5 cont dây chuyền cơ khí để vòi vĩnh 10 bộ quà tặng tri ân VIP (áo mưa, bình giữ nhiệt, ô dù). Vơ xong 10 bộ quà vào túi du lịch thì thản nhiên bảo mới tham khảo cước năm sau mới nhập, bị Sếp Thúy chặn cửa đòi lại quà.",
        "duration": "24s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx)
NEW_SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-53",
        "motif": "Motif 4: Nỗi sợ vợ thực tế của chủ xưởng kinh doanh & Dịch vụ tri ân khách hàng",
        "title": "Tặng giỏ quà tri ân đối tác và cái kết bà xã vào viện cấp cứu (Tham khảo Reel 354K views)",
        "canvas": "Phòng Xuất nhập khẩu thuê ngoài trọn gói 5 triệu/tháng, gom cont LCL từ 1 khối xe chạy hàng ngày, dịch vụ chăm sóc đối tác VIP tận tâm chu đáo bảo vệ tối đa lợi ích doanh nghiệp.",
        "cast_setting": "• Bối cảnh: Bàn tiếp khách văn phòng Eureka 44 Triều Khúc, tập hợp đồng ủy thác nhập khẩu A4, giỏ quà tri ân\n• Nhân vật: Sếp Thúy (ngồi bàn giám đốc quan sát), 1 Sales nam Eureka, 1 Bác chủ xưởng (đùng đùng xông vào, mặt mũi bơ phờ, tay cầm chiếc gối ôm bệnh viện hoặc đơn thuốc)",
        "script": "• [HOOK 0-3s]\n• Khách (đập bàn thở dốc): Hủy hợp đồng!\n• Khách: Trả cọc cho anh!\n• Sales: Dạ hủy lô nào?\n• Khách: Dây chuyền dập tôn!\n\n• [ESCALATION 4-15s]\n• Sales: Đang gom cont LCL!\n• Sales: Tiết kiệm tiền tỷ!\n• Khách: Vợ anh bắt hủy!\n• Khách: Bắt rút sạch cọc!\n• Sales: Ủa lạ thế bác?\n• Sales: Hôm qua em sang!\n• Sales: Tặng giỏ quà VIP!\n• Sales: Chị nhà thích lắm!\n• Khách (quát to): Thích cái nỗi gì!\n• Khách: Tặng xong đi viện!\n• Sales: Đi viện làm sao?\n• Khách: Dị ứng cấp cứu!\n• Sales (toát mồ hôi): Em tặng gì đâu?\n• Sales: Giỏ quà tri ân!\n\n• [PUNCHLINE 16-25s]\n• Khách: Tinh dầu hồi quế!\n• Khách: Chú dặn xông xưởng!\n• Khách: Đốt cả phòng ngủ!\n• Khách: Cay xè cả mắt!\n• Khách: Vợ sưng húp mặt!\n• Sales (mếu máo): Xông xưởng cơ mà...\n• Khách: Bà ấy vừa dặn!\n• Khách: Không sang trả viện phí...\n• Khách: Đừng hòng nhập cont!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Tinh dầu đuổi muỗi...\n• Sếp Thúy: Ai cho mang tặng?!\n• (Sound: Aiyo mama! Cậu Sales ôm đầu gãi tai, cả phòng đơ mặt 1.5s)",
        "punchline": "Bác chủ xưởng đòi hủy cọc cont máy dập tôn vì vợ bắt rút sau khi nhận quà tri ân. Hóa ra Sales tặng chai tinh dầu quế dặn xông xưởng, khách đốt cả phòng ngủ làm vợ dị ứng sưng mặt nhập viện bắt Sales sang trả tiền viện phí.",
        "duration": "25s"
    },
    {
        "code": "EKA-V2-54",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z) & Trực app ERP 24/7",
        "title": "Tuyển chọn nhân sự trực Portal ERP cuối tuần và cái kết của hội F.A (Tham khảo Reel 84K views)",
        "canvas": "Hệ thống Portal ERP và camera giám sát kho Bằng Tường trực tiếp 24/7, cam kết báo giá cước siêu tốc trong 30 giây, đội ngũ hỗ trợ nhiệt tình xuyên suốt ngày đêm.",
        "cast_setting": "• Bối cảnh: Bàn làm việc văn phòng Eureka 44 Triều Khúc, máy tính laptop, lịch để bàn, cốc trà\n• Nhân vật: Sếp Thúy (phong thái lãnh đạo, mỉm cười quan tâm), 2 Sales trẻ (1 nam, 1 nữ ngồi bấm máy tính)",
        "script": "• [HOOK 0-3s]\n• Sếp Thúy (bước tới quầy): Ai chưa người yêu?\n• Sales Nam (giơ tay nhanh): Dạ em Sếp!\n• Sales Nữ (giơ tay theo): Em cũng ế Sếp!\n\n• [ESCALATION 4-15s]\n• Sếp Thúy (mỉm cười dịu dàng): Tối nay rảnh không?\n• Sales Nam: Dạ rảnh lắm Sếp!\n• Sếp Thúy: Có ai rủ đi?\n• Sales Nữ: Không ai rủ luôn!\n• Sếp Thúy: Thích ngắm cảnh đẹp?\n• Sales Nam (mắt sáng rực): Dạ thích lắm Sếp!\n• Sales Nữ: Sếp dẫn đi chơi?\n• Sếp Thúy: Chuẩn luôn hai đứa!\n• Sếp Thúy: Cảnh quốc tế luôn!\n\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (chỉ laptop): Đây hai đứa!\n• Sếp Thúy: Trực app Portal ERP!\n• Sếp Thúy: Camera kho Bằng Tường!\n• Sếp Thúy: Hai mươi bốn trên bảy!\n• Sales Nam (đứng hình): Ủa trực ca đêm?\n• Sếp Thúy (tỉnh bơ): Không có người yêu...\n• Sếp Thúy: Về sớm tủi thân!\n• Sếp Thúy: Ngắm cont cho vui!\n\n• [DEFLATION 26-30s]\n• Sales Nữ (nhìn màn hình): Ơ hay Sếp ơi!\n• Sales Nữ: Máy nâng container...\n• Sales Nữ: Sao lãng mạn được!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Trực xong báo giá...\n• Sếp Thúy: Ba mươi giây nghe!\n• (Sound: Aiyo mama! Hai bạn Sales nhìn nhau mếu máo, cả phòng đơ mặt 1.5s)",
        "punchline": "Sếp Thúy ân cần hỏi han hội nhân viên F.A rảnh rỗi chưa có người yêu tưởng được Sếp bao đi chơi, ai ngờ được Sếp giao cho trực camera kho Bằng Tường trên app Portal ERP cả đêm cho đỡ tủi thân.",
        "duration": "22s"
    },
    {
        "code": "EKA-V2-55",
        "motif": "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Bác khách săn quà tri ân VIP",
        "title": "Hỏi book năm cont máy móc đòi mười bộ quà tri ân và cú quay xe (Tham khảo Reel 254K views)",
        "canvas": "Ủy thác nhập khẩu máy móc thiết bị dây chuyền cơ khí trọn gói, xuất hóa đơn VAT 100%, tư vấn C/O Form E miễn thuế, chính sách chăm sóc tri ân khách hàng chuyên nghiệp bằng quà tặng chuẩn nhận diện.",
        "cast_setting": "• Bối cảnh: Quầy lễ tân/tiếp khách văn phòng 44 Triều Khúc, thùng quà tặng Eureka (áo mưa, bình nước, mũ bảo hiểm), tập hợp đồng ủy thác\n• Nhân vật: Sếp Thúy (đứng kiểm tra hồ sơ), 1 Sales nữ Eureka, 1 Bác khách hàng (tay cầm túi du lịch to, kính râm cài ngực, tác phong nhanh nhẹn)",
        "script": "• [HOOK 0-3s]\n• Khách (mở túi du lịch): Eureka tặng quà VIP?\n• Sales: Dạ đúng rồi bác!\n• Khách: Quà cho khách lớn?\n• Sales: Dạ khách ký hợp đồng!\n\n• [ESCALATION 4-15s]\n• Khách: Bác book năm cont!\n• Khách: Dây chuyền cơ khí!\n• Khách: Phải đủ năm bộ!\n• Sales: Năm cont lớn thế?\n• Sales: Em xin Sếp Thúy!\n• Sales: Tặng hẳn mười bộ!\n• Sales: Áo mưa bình nước!\n• Sales: Ô dù cao cấp!\n• Khách (vơ vào túi): Lấy mười bộ nhé!\n• Khách: Cho cả xưởng dùng!\n• Sales: Dạ hợp đồng đây!\n• Sales: Bác ký hợp đồng!\n• Sales: Cọc ba mươi phần trăm!\n\n• [PUNCHLINE 16-25s]\n• Khách (kéo khóa túi lại): Cọc cạch gì đâu!\n• Sales (ngơ ngác): Dạ không cọc?\n• Khách: Bác tham khảo cước!\n• Khách: Năm sau mới nhập!\n• Khách: Lấy quà về trước!\n• Sales (chới với): Ơ bác trả quà!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (chặn cửa): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Năm sau mới nhập...\n• Sếp Thúy: Khuân sạch kho quà?!\n• (Sound: Aiyo mama! Bác khách ôm túi quà cười gượng, cả phòng đơ mặt 1.5s)",
        "punchline": "Khách khoe book 5 cont dây chuyền cơ khí để vòi vĩnh 10 bộ quà tặng tri ân VIP (áo mưa, bình giữ nhiệt, ô dù). Vơ xong 10 bộ quà vào túi du lịch thì thản nhiên bảo mới tham khảo cước năm sau mới nhập, bị Sếp Thúy chặn cửa đòi lại quà.",
        "duration": "24s"
    }
]

# Gợi ý ý tưởng cho Sheet 4 của File 2 (GOI_Y_Y_TUONG_MB)
NEW_IDEAS_SHEET_4 = [
    (
        "Tặng giỏ quà tri ân đối tác và cái kết bà xã vào viện cấp cứu (Reel 354K views)",
        "354,000 views",
        "Motif 4: Nỗi sợ vợ thực tế của chủ xưởng kinh doanh & Dịch vụ tri ân khách hàng",
        "Khách đùng đùng đòi rút hết tiền về vì vợ giãy nảy bắt rút. Hóa ra nhân viên chu đáo tặng hoa sữa rải khắp nhà khiến vợ dị ứng sưng húp nhập viện cấp cứu, bắt nhân viên sang trả viện phí.",
        "Khách đòi hủy cọc cont máy dập tôn vì vợ bắt hủy sau khi nhận quà tri ân. Hóa ra Sales tặng chai tinh dầu quế dặn xông xưởng, khách đốt cả phòng ngủ làm vợ dị ứng nhập viện (EKA-V2-53)"
    ),
    (
        "Tuyển chọn nhân sự trực Portal ERP cuối tuần cho hội F.A (Reel 84K views)",
        "84,000 views",
        "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z) & Trực app ERP 24/7",
        "Quản lý ân cần hỏi han đứa nào chưa có người yêu, chưa ai rủ đi chơi, nhân viên hớn hở tưởng được bao đi chơi ai ngờ bị ghi tên ở lại cuối tuần tăng ca cho đỡ tủi thân.",
        "Sếp Thúy hỏi thăm nhân viên F.A rảnh rỗi cuối tuần tưởng được Sếp dẫn đi chơi, ai ngờ giao trực camera kho Bằng Tường trên app Portal ERP cả đêm cho đỡ tủi thân (EKA-V2-54)"
    ),
    (
        "Hỏi book năm cont máy móc đòi mười bộ quà tri ân VIP (Reel 254K views)",
        "254,000 views",
        "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Bác khách săn quà tri ân VIP",
        "Khách đòi gửi 5 tỷ để lấy đủ 5 lọ nước hoa quà tặng cao cấp, nhân viên hứa tặng 10 lọ, khách nhận quà xong bảo không có tiền gửi mà lấy quà trước.",
        "Khách khoe book 5 cont máy cơ khí để vòi vĩnh 10 bộ quà tri ân VIP Eureka (áo mưa, bình giữ nhiệt, ô dù). Vơ xong 10 bộ quà thì bảo mới tham khảo cước năm sau mới nhập, bị Sếp Thúy chặn cửa (EKA-V2-55)"
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
    last_stt_s4 = 0
    for r in range(last_row_s4, 4, -1):
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
