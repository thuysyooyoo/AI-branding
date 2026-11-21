# -*- coding: utf-8 -*-
"""
Script: update_daily_viral_scripts_today.py
Thực hiện:
1. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 43, 44, 45) vào File 1: BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 44, 45, 46) vào File 2: BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx (Sheet 1 và Sheet 4)
3. Áp dụng kỹ thuật lưu nguyên tử atomic_save (.tmp.xlsx -> os.replace)
4. Chạy generator_chi_dan.py để đồng bộ sang Desktop TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx
"""

import os
import sys
import shutil
import subprocess
import openpyxl
from copy import copy

FILE_1 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx"
FILE_2 = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx"

# Dữ liệu 3 kịch bản mới cho File 1
NEW_SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-43",
        "title": "Sếp nghỉ ốm một buổi và cú đăng bài tuyển Sếp mới (Tham khảo Reel 54K views)",
        "mb_orig": "• Quản lý: Quái đứa nào lại đăng tin tuyển Trưởng phòng khách hàng cá nhân thế này?\n• Em: Dạ em! Em vừa có mấy người ứng tuyển, em đang phỏng vấn luôn đây!\n• Quản lý: Chị vẫn đang ở đây mà mày đi tuyển người, thích chị làm cái gì?\n• Em: Hôm qua bọn em có nhiều giấy tờ cần ký, chị nghỉ chị có báo bọn em đâu! Bọn em phải tuyển thêm Sếp để có người ký duyệt chứ! Hôm trước đông khách quá, chị phải tăng ca đau đầu quá xin nghỉ một hôm, Sếp nghỉ là không được, ốm cũng phải có kế hoạch! Sếp muốn ốm phải báo trước để bọn em còn chuẩn bị!\n• Quản lý: Nhưng mà này, cái vị trí của chị ở ngoài thị trường đang được trả lương bao nhiêu?\n• Em: Giờ mà chị sang chỗ khác là phải học từ đầu đấy, có khi phải phấn đấu từ vị trí tạp vụ lễ tân!\n• Quản lý: Thôi mày xóa bài tuyển dụng đi cho chị!\n• Em: Đăng tốn 200k tiền quảng cáo rồi, giờ xóa kiểu gì hả Sếp?\n• (Sound: Aiyo mama!)",
        "eureka": "• [HOOK 0-3s]\n• Sếp Thúy: Đứa nào đăng bài này?\n• Sales: Em đăng tuyển đấy!\n• Sếp Thúy: Tuyển ai hả em?\n• Sales: Tuyển thêm Giám Đốc!\n\n• [ESCALATION 4-15s]\n• Sếp Thúy: Chị còn ở đây!\n• Sales: Hôm qua Sếp ốm!\n• Sales: Mười bộ tờ khai!\n• Sales: Cần ký C/O gấp!\n• Sales: Máy phay nằm chờ!\n• Sếp Thúy: Chị đi khám răng!\n• Sales: Ốm phải có kế hoạch!\n• Sales: Báo trước ba ngày!\n• Sếp Thúy: Lương định trả mấy?\n\n• [PUNCHLINE 16-25s]\n• Sales (chìa bảng tính): Năm mươi ngàn một khối!\n• Sếp Thúy: Khối cont gom á?\n• Sales: Dạ, chạy hàng ngày!\n• Sếp Thúy (đơ mặt 1.5s): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Tuyển Sếp ăn lương mét?\n\n• [DEFLATION 26-30s]\n• Sales: Tốn tiền chạy tin rồi!\n• Sếp Thúy: Trừ vào hoa hồng!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "cast": "Sếp Thúy (áo sơ mi công sở, tay cầm tập hồ sơ kiểm tra) + 1 Sales nữ Gen Z (tay lướt điện thoại phỏng vấn) + 1 Ứng viên ngồi sofa",
        "setting": "Bàn làm việc Sếp Thúy & quầy tiếp khách văn phòng Eureka 44 Triều Khúc, góc quay trung cận",
        "twist": "Sếp Thúy nghỉ nửa ngày đi khám răng, về công ty bắt quả tang Sales Gen Z tự ý đăng bài tuyển Giám Đốc mới để ký C/O cont máy phay, còn tự tin trả lương theo sản lượng cont gom LCL 50k một mét khối.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-44",
        "title": "Dụ Sales nhảy việc sang xưởng siêu xe và cái kết rằm tháng bảy (Tham khảo Reel 82K views)",
        "mb_orig": "• Khách: Dạo này làm ngân hàng nhiều việc thế em? Nhưng mà nhân viên làm ngân hàng vất vả mà lương thưởng thấp quá! Bên anh cũng kinh doanh tiền tệ này, tiền vào như nước, nhân viên nhà anh đếm tiền mỏi tay, thưởng nóng liên tục! Nhìn em nhanh nhẹn thế này, có thích sang bên anh làm không? Rằm mùng một bận lắm!\n• Em: Anh ơi, em vừa nộp đơn xin nghỉ việc ở đây rồi! Cho em sang bên anh thử việc luôn nhé! Thế bên mình làm ngân hàng quốc tế hay công ty tài chính gì hả anh?\n• Khách: Ngân hàng tài chính gì đâu em! Cửa hàng vàng mã to nhất làng! Ngày rằm mùng một đông khách lắm, đếm tiền hoa cả mắt!\n• (Sound: Aiyo mama!)",
        "eureka": "• [HOOK 0-3s]\n• Khách: Tư vấn cont mệt nhỉ?\n• Sales: Dạ, nhiều khách book!\n• Khách: Sang xưởng anh làm!\n• Sales: Xưởng gì hả bác?\n\n• [ESCALATION 4-15s]\n• Khách: Xưởng làm siêu xe!\n• Khách: Rolls-Royce triệu đô!\n• Khách: Cả du thuyền xịn!\n• Sales: Đơn hàng nhiều không?\n• Khách: Cả ngàn chiếc xuất!\n• Khách: Thiếu Trưởng phòng XNK!\n• Khách: Lương trăm củ khoai!\n• Sales (sáng mắt): Ký hợp đồng luôn!\n\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (bước ra): Đưa chị xem nào!\n• Sếp Thúy (cầm hồ sơ): Xưởng Duyên Xương à?\n• Khách: Dạ đúng rồi sếp!\n• Sếp Thúy: Chuyên dán vàng mã?\n• Khách: Dạ... cúng rằm tháng bảy!\n• Sales (đơ mặt): Ơ... siêu xe giấy á?\n\n• [DEFLATION 26-30s]\n• Sếp Thúy: Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Eureka chỉ đi chính ngạch!\n• Sếp Thúy: Không chở vàng mã!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "cast": "1 Bác chủ xưởng (áo bà ba cách tân, tay đeo vòng phong thủy) + 1 Sales nữ Eureka + Sếp Thúy",
        "setting": "Bàn sofa tiếp khách văn phòng Eureka 44 Triều Khúc, máy quay góc ngang lia từ bàn làm việc Sếp Thúy",
        "twist": "Bác chủ xưởng gạ gẫm Sales nữ nhảy việc sang làm Trưởng phòng XNK xưởng siêu xe triệu đô và du thuyền quốc tế, ai ngờ Sếp Thúy đọc hồ sơ lật tẩy là xưởng thủ công mỹ nghệ dán vàng mã cúng rằm tháng bảy.",
        "duration": "25s"
    },
    {
        "code": "EKA-REEL-45",
        "title": "Tập dượt thuật ngữ thị uy đối thủ và cú bắn nhầm bà xã (Tham khảo Reel 165K views)",
        "mb_orig": "• Chị: Cái con tiểu tam đối thủ kia hôm nay thách thức chị! Hai đứa xông lên cho chị!\n• Em: Sếp ơi nó đang thách cái gì thế?\n• Chị: Nó thách chửi nhau tay đôi! Chết rồi chị lại chưa học qua khóa chửi nhau nào cơ! Hai đứa có biết chửi nhau không?\n• Em: Bọn em một thân một mình làm sao biết chửi nhau hả sếp?\n• Chị: Lúc chửi nhau quan trọng nhất là phải tuôn mồm dồn dập, mà chị lại hay bị vấp! Bây giờ thế này: Hai đứa nghĩ ra bao nhiêu từ ngữ mạnh nhất, quen mồm nhất tuôn hết ra cho chị, để tí chị xông ra thị uy!\n• Em: 50 triệu 8,2%! 1 tỷ 9,1%! Miễn phí chuyển khoản trọn đời!\n• Chị: Ơ hay... chửi nhau bằng lãi suất tiết kiệm à?\n• (Sound: Aiyo mama!)",
        "eureka": "• [HOOK 0-3s]\n• Khách: Đối thủ chê bác!\n• Sales: Chê gì cơ bác?\n• Khách: Chê bác mù luật!\n• Sales: Để Eureka lo!\n\n• [ESCALATION 4-15s]\n• Khách: Dạy bác thị uy!\n• Khách: Mấy câu sắc bén!\n• Sales: Đứng thẳng lưng lên!\n• Sales: Bắn nhanh như chớp!\n• Sales: Mã HS chuẩn tám số!\n• Sales: Form E miễn thuế!\n• Sales: Hóa đơn VAT đủ!\n• Sales: Cấp nợ bốn mươi!\n• Khách: Cấp nợ bốn mươi!\n\n• [PUNCHLINE 16-25s]\n• Vợ (bước vào sau lưng): Cấp nợ cho ai?\n• Khách (giật mình quay lại): Cấp nợ bốn mươi!\n• Vợ (giơ guốc): Bốn mươi phần trăm?\n• Vợ: Cho con nào nợ?\n• Khách (toát mồ hôi hột): Kìa... bà xã ơi!\n• Khách: Bác đang tập thoại!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy: Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Bắn nhầm đối tượng rồi!\n• (Sound: Aiyo mama! Bác chủ xưởng giơ hai tay đầu hàng, cả phòng đơ mặt 1.5s)",
        "cast": "1 Bác chủ xưởng cơ khí (áo sơ mi, vẻ mặt nghiêm trọng) + 1 Bà vợ (tay cầm túi xách, giơ guốc) + 1 Sales nam Eureka + Sếp Thúy",
        "setting": "Sảnh cửa kính văn phòng Eureka 44 Triều Khúc lia vào quầy lễ tân",
        "twist": "Bác chủ xưởng nhờ Eureka huấn luyện loạt thuật ngữ hải quan đanh thép để sang thị uy đối thủ cạnh tranh, vừa hăng say hô lớn 'Cấp nợ bốn mươi phần trăm' thì bà vợ kế toán bước vào tưởng cấp vốn cho quỹ đen, giơ guốc rượt đuổi.",
        "duration": "24s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2
NEW_SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-44",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Nhân viên)",
        "title": "Sếp nghỉ ốm một buổi và cú đăng bài tuyển Sếp mới (Tham khảo Reel 54K views)",
        "canvas": "Phòng XNK thuê ngoài trọn gói từ 5 triệu/tháng, cơ chế ký duyệt số 24/7 không phụ thuộc cá nhân, thông quan cont máy phay máy dập đúng hạn.",
        "cast_setting": "• Bối cảnh: Bàn làm việc Sếp Thúy & quầy tiếp khách văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy (áo sơ mi công sở, tay cầm tập hồ sơ kiểm tra), 1 Sales nữ Gen Z (tay lướt điện thoại phỏng vấn), 1 Ứng viên ngồi sofa",
        "script": "• [HOOK 0-3s]\n• Sếp Thúy: Đứa nào đăng bài này?\n• Sales: Em đăng tuyển đấy!\n• Sếp Thúy: Tuyển ai hả em?\n• Sales: Tuyển thêm Giám Đốc!\n\n• [ESCALATION 4-15s]\n• Sếp Thúy: Chị còn ở đây!\n• Sales: Hôm qua Sếp ốm!\n• Sales: Mười bộ tờ khai!\n• Sales: Cần ký C/O gấp!\n• Sales: Máy phay nằm chờ!\n• Sếp Thúy: Chị đi khám răng!\n• Sales: Ốm phải có kế hoạch!\n• Sales: Báo trước ba ngày!\n• Sếp Thúy: Lương định trả mấy?\n\n• [PUNCHLINE 16-25s]\n• Sales (chìa bảng tính): Năm mươi ngàn một khối!\n• Sếp Thúy: Khối cont gom á?\n• Sales: Dạ, chạy hàng ngày!\n• Sếp Thúy (đơ mặt 1.5s): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Tuyển Sếp ăn lương mét?\n\n• [DEFLATION 26-30s]\n• Sales: Tốn tiền chạy tin rồi!\n• Sếp Thúy: Trừ vào hoa hồng!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "punchline": "Sếp Thúy nghỉ nửa ngày đi khám răng, về công ty bắt quả tang Sales Gen Z tự ý đăng bài tuyển Giám Đốc mới để ký C/O cont máy phay, còn tự tin trả lương theo sản lượng cont gom LCL 50k một mét khối.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-45",
        "motif": "Motif 3: Trào phúng tâm linh & Thói quen dân buôn xuất nhập khẩu",
        "title": "Dụ Sales nhảy việc sang xưởng siêu xe và cái kết rằm tháng bảy (Tham khảo Reel 82K views)",
        "canvas": "Dịch vụ ủy thác nhập khẩu chính ngạch linh kiện, phụ tùng máy móc cơ khí, bảo đảm 100% hóa đơn VAT hợp pháp, tư vấn mã HS và C/O Form E miễn thuế.",
        "cast_setting": "• Bối cảnh: Bàn sofa tiếp khách văn phòng Eureka 44 Triều Khúc\n• Nhân vật: 1 Bác chủ xưởng (áo bà ba cách tân, tay đeo vòng phong thủy), 1 Sales nữ Eureka, Sếp Thúy",
        "script": "• [HOOK 0-3s]\n• Khách: Tư vấn cont mệt nhỉ?\n• Sales: Dạ, nhiều khách book!\n• Khách: Sang xưởng anh làm!\n• Sales: Xưởng gì hả bác?\n\n• [ESCALATION 4-15s]\n• Khách: Xưởng làm siêu xe!\n• Khách: Rolls-Royce triệu đô!\n• Khách: Cả du thuyền xịn!\n• Sales: Đơn hàng nhiều không?\n• Khách: Cả ngàn chiếc xuất!\n• Khách: Thiếu Trưởng phòng XNK!\n• Khách: Lương trăm củ khoai!\n• Sales (sáng mắt): Ký hợp đồng luôn!\n\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (bước ra): Đưa chị xem nào!\n• Sếp Thúy (cầm hồ sơ): Xưởng Duyên Xương à?\n• Khách: Dạ đúng rồi sếp!\n• Sếp Thúy: Chuyên dán vàng mã?\n• Khách: Dạ... cúng rằm tháng bảy!\n• Sales (đơ mặt): Ơ... siêu xe giấy á?\n\n• [DEFLATION 26-30s]\n• Sếp Thúy: Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Eureka chỉ đi chính ngạch!\n• Sếp Thúy: Không chở vàng mã!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "punchline": "Bác chủ xưởng gạ gẫm Sales nữ nhảy việc sang làm Trưởng phòng XNK xưởng siêu xe triệu đô và du thuyền quốc tế, ai ngờ Sếp Thúy đọc hồ sơ lật tẩy là xưởng thủ công mỹ nghệ dán vàng mã cúng rằm tháng bảy.",
        "duration": "25s"
    },
    {
        "code": "EKA-V2-46",
        "motif": "Motif 6: Nhại điện ảnh điệp viên & Đấu trí hồ sơ chứng từ",
        "title": "Tập dượt thuật ngữ thị uy đối thủ và cú bắn nhầm bà xã (Tham khảo Reel 165K views)",
        "canvas": "Gói đòn bẩy tài chính cấp hạn mức công nợ 40%, tư vấn C/O Form E miễn thuế 100%, hồ sơ chuẩn chỉ bảo vệ pháp nhân 10 năm với hóa đơn VAT chính quy.",
        "cast_setting": "• Bối cảnh: Sảnh cửa kính văn phòng Eureka 44 Triều Khúc lia vào quầy lễ tân\n• Nhân vật: 1 Bác chủ xưởng cơ khí (áo sơ mi, vẻ mặt nghiêm trọng), 1 Bà vợ (tay cầm túi xách, giơ guốc), 1 Sales nam Eureka, Sếp Thúy",
        "script": "• [HOOK 0-3s]\n• Khách: Đối thủ chê bác!\n• Sales: Chê gì cơ bác?\n• Khách: Chê bác mù luật!\n• Sales: Để Eureka lo!\n\n• [ESCALATION 4-15s]\n• Khách: Dạy bác thị uy!\n• Khách: Mấy câu sắc bén!\n• Sales: Đứng thẳng lưng lên!\n• Sales: Bắn nhanh như chớp!\n• Sales: Mã HS chuẩn tám số!\n• Sales: Form E miễn thuế!\n• Sales: Hóa đơn VAT đủ!\n• Sales: Cấp nợ bốn mươi!\n• Khách: Cấp nợ bốn mươi!\n\n• [PUNCHLINE 16-25s]\n• Vợ (bước vào sau lưng): Cấp nợ cho ai?\n• Khách (giật mình quay lại): Cấp nợ bốn mươi!\n• Vợ (giơ guốc): Bốn mươi phần trăm?\n• Vợ: Cho con nào nợ?\n• Khách (toát mồ hôi hột): Kìa... bà xã ơi!\n• Khách: Bác đang tập thoại!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy: Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Bắn nhầm đối tượng rồi!\n• (Sound: Aiyo mama! Bác chủ xưởng giơ hai tay đầu hàng, cả phòng đơ mặt 1.5s)",
        "punchline": "Bác chủ xưởng nhờ Eureka huấn luyện loạt thuật ngữ hải quan đanh thép để sang thị uy đối thủ cạnh tranh, vừa hăng say hô lớn 'Cấp nợ bốn mươi phần trăm' thì bà vợ kế toán bước vào tưởng cấp vốn cho quỹ đen, giơ guốc rượt đuổi.",
        "duration": "24s"
    }
]

# Gợi ý ý tưởng cho Sheet 4 của File 2
NEW_IDEAS_SHEET_4 = [
    (
        "Quản lý nghỉ ốm nhân viên đăng tuyển Sếp mới (Reel 54K views)",
        "54,000 views",
        "Motif 2: Drama nội bộ công sở & Sếp vs Nhân viên",
        "Quản lý nghỉ ốm một buổi vì nhức đầu, nhân viên tự ý đăng bài tuyển Trưởng phòng mới vì không có ai ký hồ sơ và bảo Sếp ốm phải có kế hoạch trước.",
        "Sếp Thúy nghỉ nửa ngày đi khám răng, Sales tự ý đăng bài tuyển Giám đốc mới ký C/O máy phay, lương tính 50k/khối cont gom LCL (EKA-V2-44)"
    ),
    (
        "Gạ nhân viên nhảy việc sang xưởng kinh doanh tiền tệ (Reel 82K views)",
        "82,000 views",
        "Motif 3: Trào phúng tâm linh & Vỡ mộng đổi đời",
        "Khách chê ngân hàng lương thấp rủ nhân viên sang làm cho xưởng tiền tệ triệu đô của anh, nhân viên xin nghỉ việc hóa ra là xưởng vàng mã rằm mùng một.",
        "Khách rủ Sales sang làm Trưởng phòng XNK xưởng siêu xe Rolls-Royce du thuyền, Sếp Thúy lật tẩy xưởng vàng mã cúng rằm tháng bảy (EKA-V2-45)"
    ),
    (
        "Tập chửi nhau dồn dập bằng các gói lãi suất tiết kiệm (Reel 165K views)",
        "165,000 views",
        "Motif 6: Nhại điệp viên & Đấu trí hồ sơ chứng từ",
        "Chị sếp bị đối thủ khiêu chiến, bảo hai nhân viên tuôn hết từ ngữ mạnh nhất để học thuộc thị uy, nhân viên bắn liên thanh toàn gói gửi tiết kiệm 50 triệu 8,2%.",
        "Bác chủ xưởng nhờ Eureka dạy thuật ngữ nghiệp vụ đanh thép thị uy đối thủ, vừa hô 'Cấp nợ bốn mươi phần trăm' thì bà vợ bước vào giơ guốc (EKA-V2-46)"
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
    
    # 1. Kiểm tra kịch bản có note "Đã sản xuất" để xóa nếu có
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
            
    print(f"[*] Hàng cuối hiện tại: {last_row}, STT cuối: {last_stt}")
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

def update_file_2():
    print(f"\n[*] Đang cập nhật File 2: {FILE_2}")
    wb = openpyxl.load_workbook(FILE_2)
    
    # --- Cập nhật Sheet 1: KICH_BAN_VIRAL_REELS ---
    ws1 = wb['KICH_BAN_VIRAL_REELS']
    
    # 1. Kiểm tra kịch bản có note "Đã sản xuất"
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
