# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_1311.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 13/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-64, 65, 66) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-65, 66, 67) vào Sheet 1 và 3 gợi ý vào Sheet 4 của File 2:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx
3. Tự động kiểm tra và xóa kịch bản có ghi chú "Đã sản xuất" (nếu có).
4. Áp dụng kỹ thuật lưu nguyên tử (Atomic Save) qua file đệm .tmp.xlsx và os.replace.
5. Kích hoạt generator_chi_dan.py để đồng bộ Desktop dashboard.
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
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

# Dữ liệu 3 kịch bản mới cho File 1 (BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx)
SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-64",
        "title": "Khách tự tra cứu Portal ERP từ A đến Z và lý do bất ngờ ghé Eureka (Tham khảo Reel 110K views)",
        "mb_raw": "• Khách: Sáng nay anh đến gửi tiết kiệm hả?\n• Nhân viên: Dạ vâng, hôm nay ngày đẹp anh à!\n• Khách: Sao phải hôm nay? Ngày mai đến gửi được không?\n• Nhân viên: Dạ gửi lúc nào cũng được anh à! Nhưng anh gửi kỳ hạn bao lâu? 5 năm nhé?\n• Khách: Sao phải 4-5 năm? Ngân hàng uy tín lãi suất ưu đãi gửi 5 năm?\n• Nhân viên: Dạ bên em có chương trình ưu đãi, gửi từ 10 triệu đã được 8.5 rồi nha. Em mở sổ tại quầy cho anh nha?\n• Khách: Sao phải tại quầy? Bây giờ thích làm online cho lãi suất cao hơn nhé!\n• Nhân viên: Thế để em hướng dẫn anh mở online?\n• Khách: Đâu cần hướng dẫn! Bấm mở xong rồi nè!\n• Nhân viên: Ơ anh ơi, online mở tại nhà cũng được nha!\n• Khách: Sao phải ở nhà? Ở nhà cho tốn điện ra, ra đây có điều hòa cho nó mát!\n• Nhân viên: Dạ, thế em chào anh nhé!\n• Khách: Sao phải chào? Mai anh lại ra, chào sau!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales: Em chào bác!\n• Khách: Khỏi chào mất công!\n• Sales: Mời bác tra cứu?\n• Khách: Sao phải tra cứu?\n• [ESCALATION 4-15s]\n• Sales: Check cont máy dập?\n• Khách: Xong từ sáng sớm!\n• Sales: Bác tra bằng gì?\n• Khách: App Portal ERP!\n• Sales: Xe tới Bằng Tường?\n• Khách: Định vị chuẩn nét!\n• Sales: Đặt lệnh thông quan?\n• Khách: Tự bấm một chạm!\n• Sales: Cần em tính cước?\n• Khách: Hệ thống tự chia!\n• Sales: Bác siêu công nghệ!\n• [PUNCHLINE 16-25s]\n• Sales: Thế ghé đây chi?\n• Khách (đơ 1.5s): Ơ hay! Xưởng mất điện!\n• Khách: Nóng chảy mỡ người!\n• Khách: Ra đây ké lạnh!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Điều hòa mát chứ?\n• Khách: Mai bác lại sang!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (ngồi bàn giám đốc quan sát) + Sales Eureka (cầm tablet) + Bác khách chủ xưởng (tay cầm điện thoại mở app ERP, phe phẩy quạt nan).",
        "scene": "Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, màn hình tablet và điện thoại mở giao diện Portal ERP Eureka, kết thúc cận mặt Sếp Thúy tỉnh bơ.",
        "twist": "Khách thao tác cực mượt trên Portal ERP tự check vị trí xe cont tại kho Bằng Tường và lệnh thông quan không thiếu thứ gì, sales ngạc nhiên hỏi sao còn ghé văn phòng thì mới gãi đầu bảo xưởng mất điện nóng quá nên sang ngồi điều hòa ké.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-65",
        "title": "Đòi tăng hạn mức công nợ vì gương mặt chuẩn tổng tài và cái phán của Sếp Thúy (Tham khảo Reel 179K views)",
        "mb_raw": "• Nhân viên: MB Hải Dương rất vui được gặp anh ạ!\n• Khách: Vui hơi sớm đấy!\n• Nhân viên: Bên em đang có chương trình tặng thêm lãi suất cho những người đẹp trai đấy ạ!\n• Khách: Đẹp trai thì đây có thừa! Thế lãi suất gốc bao nhiêu?\n• Nhân viên: 50 triệu 8.2, 1 tỷ 9.1... nhưng với những anh soái ca tổng tài đẹp trai thì tặng thêm ạ!\n• Khách: Thế tầm này thì được tặng thêm bao nhiêu ạ?\n• Nhân viên (soi kỹ): Dạ... anh thông cảm, của anh giữ nguyên ạ!\n• Khách: Giữ nguyên hả? Chán thế nhờ!\n• (Sound: Aiyo mama! Cả quầy đứng hình đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales: Em chào tổng tài!\n• Khách: Chào hơi muộn đấy!\n• Sales: Anh nhập máy phay?\n• Khách: Máy phay năm trục!\n• [ESCALATION 4-15s]\n• Khách: Cấp công nợ nhé?\n• Sales: Bên em có cấp!\n• Khách: Cấp tối đa luôn?\n• Sales: Bốn mươi phần trăm!\n• Khách: Thẩm định hồ sơ?\n• Sales: Cần soi pháp nhân!\n• Khách: Mặt anh bảo lãnh!\n• Khách: Chuẩn nét tổng tài!\n• Sales: Để em xin Sếp!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (chìa hồ sơ CIC, đơ 1.5s): Ơ hay! Tổng tài?\n• Sếp Thúy: Mặt thì rất bảnh!\n• Sếp Thúy: CIC nợ xấu đỏ!\n• Khách (giật mình): Sếp soi ra à?\n• [DEFLATION 26-30s]\n• Sếp Thúy: Cọc một trăm phần trăm!\n• Sếp Thúy: Miễn trừ tổng tài!\n• (Sound: Aiyo mama! Khách méo mặt)",
        "roles": "Sếp Thúy (đứng cầm hồ sơ CIC doanh nghiệp) + Sales Eureka (cầm catalogue máy phay CNC) + Khách nam (áo sơ mi bóng bẩy, vuốt tóc, đeo kính râm đen phong cách tổng tài).",
        "scene": "Quầy lễ tân và bàn tiếp khách văn phòng Eureka 44 Triều Khúc, tập tài liệu thẩm định tín dụng CIC, cận cảnh Sếp Thúy nghiêm nghị chỉ vào điểm tín dụng đỏ.",
        "twist": "Khách nhập cont máy phay CNC 5 trục đòi cấp hạn mức công nợ 40% bằng gương mặt chuẩn tổng tài không cần thẩm định hồ sơ, Sếp Thúy cầm bản tra cứu CIC thấy nợ xấu nhảy nhóm đỏ rực liền yêu cầu cọc 100% không miễn trừ.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-66",
        "title": "Dõng dạc đòi xuất hóa đơn VAT tên bồ nhí và cú quay xe của Sếp Thúy (Tham khảo Reel 260K views)",
        "mb_raw": "• Khách: Hôm nay phải tất tay em ạ! Anh già rồi, tiếc gì nữa!\n• Khách: Anh đưa em đi mua sắm, thích gì cũng được!\n• Em: Em thích nhiều thứ lắm đó!\n• Khách: Anh biết rồi, thích cả trang sức tiệm vàng bạc chứ gì? Cứ mua thoải mái! Với người đẹp như em có đáng gì đâu! Tí anh chuyển mấy chục đi chơi cho tự tin!\n• Khách: Rồi anh từ ngân hàng về anh đón nhé! Đưa sổ tiết kiệm đây!\n• Em: Ủa, tớ đang nói chuyện với người yêu tớ cơ mà! Cho xin cái bông tai đi chơi với người yêu nhé!\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách (gọi điện thoại): Em yêu yên tâm!\n• Khách: Anh bao tất cả!\n• Sales: Bác nhập cont gì?\n• Khách: Dây chuyền dập khuôn!\n• [ESCALATION 4-15s]\n• Sales: Vận chuyển chính ngạch?\n• Khách: Chuẩn từng con ốc!\n• Sales: Hóa đơn đỏ VAT?\n• Khách: Xuất tên em yêu!\n• Sales: Pháp nhân công ty?\n• Khách: Công ty Vạn Lộc!\n• Khách: Ghi tặng bồ nhí!\n• Sales: Bác chịu chơi quá!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (cầm DKKD bước sang, đơ 1.5s): Ơ hay! Bồ nhí?\n• Sếp Thúy: Công ty Vạn Lộc...\n• Sếp Thúy: Chủ tịch là ai?\n• Khách (chột dạ): Là... bà xã anh!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Gửi Zalo bà xã?\n• Khách (xua tay rối rít): Thôi! Xuất tên xưởng!\n• (Sound: Aiyo mama! Bác khách toát mồ hôi)",
        "roles": "Sếp Thúy (cầm bản in Giấy phép DKKD và con dấu) + Sales/Chứng từ Eureka (ngồi máy tính phần mềm hóa đơn) + Bác khách chủ xưởng (ngồi sofa gọi điện thoại huênh hoang).",
        "scene": "Bàn làm việc chứng từ & sofa tiếp khách văn phòng Eureka 44 Triều Khúc, điện thoại smartphone đang bật loa ngoài, bản in DKKD có tên vợ khách làm đại diện pháp luật.",
        "twist": "Khách huênh hoang gọi điện hứa nhập dây chuyền máy dập khuôn xuất hóa đơn đỏ VAT cho công ty của bồ nhí để ghi điểm, Sếp Thúy tra cứu DKKD thấy Chủ tịch công ty chính là bà xã của khách, dọa gửi thẳng hóa đơn vào Zalo bà xã khiến khách sợ toát mồ hôi xin xuất tên xưởng cũ.",
        "duration": "24s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-65",
        "motif": "Motif 2: Drama đối thoại công sở & Khách nghiện công nghệ",
        "title": "Khách tự tra cứu Portal ERP từ A đến Z và lý do bất ngờ ghé Eureka (Tham khảo Reel 110K views)",
        "strength": "Ứng dụng Portal ERP theo dõi vị trí cont máy móc 24/7, định vị xe chạy thời gian thực kho Bằng Tường - Hà Nội, báo giá tự động không cần gọi điện.",
        "roles_scene": "• Bối cảnh: Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Bác khách chủ xưởng (cầm điện thoại mở app ERP, phẩy quạt nan)\n• Đạo cụ: Tablet hiển thị Portal ERP, điện thoại smartphone, quạt nan tre truyền thống.",
        "pingpong": "• [HOOK 0-3s]\n• Sales: Em chào bác!\n• Khách: Khỏi chào mất công!\n• Sales: Mời bác tra cứu?\n• Khách: Sao phải tra cứu?\n• [ESCALATION 4-15s]\n• Sales: Check cont máy dập?\n• Khách: Xong từ sáng sớm!\n• Sales: Bác tra bằng gì?\n• Khách: App Portal ERP!\n• Sales: Xe tới Bằng Tường?\n• Khách: Định vị chuẩn nét!\n• Sales: Đặt lệnh thông quan?\n• Khách: Tự bấm một chạm!\n• Sales: Cần em tính cước?\n• Khách: Hệ thống tự chia!\n• Sales: Bác siêu công nghệ!\n• [PUNCHLINE 16-25s]\n• Sales: Thế ghé đây chi?\n• Khách (đơ 1.5s): Ơ hay! Xưởng mất điện!\n• Khách: Nóng chảy mỡ người!\n• Khách: Ra đây ké lạnh!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Điều hòa mát chứ?\n• Khách: Mai bác lại sang!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Khách thao tác cực mượt trên Portal ERP tự check vị trí xe cont tại kho Bằng Tường và lệnh thông quan không thiếu thứ gì, sales ngạc nhiên hỏi sao còn ghé văn phòng thì mới gãi đầu bảo xưởng mất điện nóng quá nên sang ngồi điều hòa ké.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-66",
        "motif": "Motif 5: Chơi chữ hài hước & Đòn bẩy tài chính B2B",
        "title": "Đòi tăng hạn mức công nợ vì gương mặt chuẩn tổng tài và cái phán của Sếp Thúy (Tham khảo Reel 179K views)",
        "strength": "Cấp hạn mức công nợ 10% - 40% cho doanh nghiệp cơ khí máy móc chính ngạch; thẩm định pháp nhân minh bạch, không cảm tính.",
        "roles_scene": "• Bối cảnh: Quầy lễ tân & bàn tiếp khách văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Khách nam (áo sơ mi bóng bẩy, vuốt tóc, kính râm đen)\n• Đạo cụ: Bản thẩm định tín dụng CIC, catalogue máy phay CNC 5 trục, kính râm đen tổng tài.",
        "pingpong": "• [HOOK 0-3s]\n• Sales: Em chào tổng tài!\n• Khách: Chào hơi muộn đấy!\n• Sales: Anh nhập máy phay?\n• Khách: Máy phay năm trục!\n• [ESCALATION 4-15s]\n• Khách: Cấp công nợ nhé?\n• Sales: Bên em có cấp!\n• Khách: Cấp tối đa luôn?\n• Sales: Bốn mươi phần trăm!\n• Khách: Thẩm định hồ sơ?\n• Sales: Cần soi pháp nhân!\n• Khách: Mặt anh bảo lãnh!\n• Khách: Chuẩn nét tổng tài!\n• Sales: Để em xin Sếp!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (chìa hồ sơ CIC, đơ 1.5s): Ơ hay! Tổng tài?\n• Sếp Thúy: Mặt thì rất bảnh!\n• Sếp Thúy: CIC nợ xấu đỏ!\n• Khách (giật mình): Sếp soi ra à?\n• [DEFLATION 26-30s]\n• Sếp Thúy: Cọc một trăm phần trăm!\n• Sếp Thúy: Miễn trừ tổng tài!\n• (Sound: Aiyo mama! Khách méo mặt)",
        "twist": "Khách nhập cont máy phay CNC 5 trục đòi cấp hạn mức công nợ 40% bằng gương mặt chuẩn tổng tài không cần thẩm định hồ sơ, Sếp Thúy cầm bản tra cứu CIC thấy nợ xấu nhảy nhóm đỏ rực liền yêu cầu cọc 100% không miễn trừ.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-67",
        "motif": "Motif 4: Nỗi sợ gia đình & 100% Hóa đơn VAT chuẩn pháp nhân",
        "title": "Dõng dạc đòi xuất hóa đơn VAT tên bồ nhí và cú quay xe của Sếp Thúy (Tham khảo Reel 260K views)",
        "strength": "100% Hóa đơn VAT chính ngạch, tư vấn mã HS và C/O Form E chuẩn chỉ cho pháp nhân doanh nghiệp, phòng tránh rủi ro kiểm tra sau thông quan.",
        "roles_scene": "• Bối cảnh: Bàn làm việc chứng từ & sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Bác khách chủ xưởng (gọi điện thoại huênh hoang)\n• Đạo cụ: Bản in Giấy phép DKKD, con dấu đỏ công ty, điện thoại smartphone bật loa ngoài.",
        "pingpong": "• [HOOK 0-3s]\n• Khách (gọi điện thoại): Em yêu yên tâm!\n• Khách: Anh bao tất cả!\n• Sales: Bác nhập cont gì?\n• Khách: Dây chuyền dập khuôn!\n• [ESCALATION 4-15s]\n• Sales: Vận chuyển chính ngạch?\n• Khách: Chuẩn từng con ốc!\n• Sales: Hóa đơn đỏ VAT?\n• Khách: Xuất tên em yêu!\n• Sales: Pháp nhân công ty?\n• Khách: Công ty Vạn Lộc!\n• Khách: Ghi tặng bồ nhí!\n• Sales: Bác chịu chơi quá!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy (cầm DKKD bước sang, đơ 1.5s): Ơ hay! Bồ nhí?\n• Sếp Thúy: Công ty Vạn Lộc...\n• Sếp Thúy: Chủ tịch là ai?\n• Khách (chột dạ): Là... bà xã anh!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Gửi Zalo bà xã?\n• Khách (xua tay rối rít): Thôi! Xuất tên xưởng!\n• (Sound: Aiyo mama! Bác khách toát mồ hôi)",
        "twist": "Khách huênh hoang gọi điện hứa nhập dây chuyền máy dập khuôn xuất hóa đơn đỏ VAT cho công ty của bồ nhí để ghi điểm, Sếp Thúy tra cứu DKKD thấy Chủ tịch công ty chính là bà xã của khách, dọa gửi thẳng hóa đơn vào Zalo bà xã khiến khách sợ toát mồ hôi xin xuất tên xưởng cũ.",
        "duration": "24s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 61,
        "title": "Khách tự tra cứu Portal ERP từ A đến Z và lý do bất ngờ ghé Eureka (Reel 110K views)",
        "views": "110,000 views",
        "motif": "Motif 2: Drama đối thoại công sở & Khách nghiện công nghệ",
        "trigger": "Hỏi gì cũng vặn lại 'Sao phải...', nhân viên nhiệt tình hỗ trợ thì khách đã tự thao tác xong trên app online, lý do ghé là để ké điều hòa cho mát đỡ tốn điện xưởng.",
        "adaptation": "Khách nhập cont máy dập tự check vị trí xe Bằng Tường và tính cước trên app ERP của Eureka không thiếu thứ gì, nhân viên hỏi sao còn phi sang 44 Triều Khúc làm gì thì bảo xưởng mất điện nóng quá sang ngồi sofa hưởng điều hòa mát rượi (EKA-V2-65)"
    },
    {
        "stt": 62,
        "title": "Đòi tăng hạn mức công nợ vì gương mặt chuẩn tổng tài và cái phán của Sếp Thúy (Reel 179K views)",
        "views": "179,000 views",
        "motif": "Motif 5: Chơi chữ hài hước & Đòn bẩy tài chính B2B",
        "trigger": "Tưởng bở được ngân hàng cộng thêm lãi suất vì là 'soái ca tổng tài đẹp trai', nhân viên soi kỹ mặt xong phán 'Dạ của anh giữ nguyên ạ'.",
        "adaptation": "Khách nhập cont máy phay CNC 5 trục đòi cấp hạn mức công nợ 40% tối đa bằng gương mặt chuẩn tổng tài không cần thẩm định, Sếp Thúy soi ra lịch sử CIC nợ xấu đỏ rực bắt cọc 100% không miễn trừ ngoại lệ (EKA-V2-66)"
    },
    {
        "stt": 63,
        "title": "Dõng dạc đòi xuất hóa đơn VAT tên bồ nhí và cú quay xe của Sếp Thúy (Reel 260K views)",
        "views": "260,000 views",
        "motif": "Motif 4: Nỗi sợ gia đình & 100% Hóa đơn VAT chuẩn pháp nhân",
        "trigger": "Chém gió gọi điện với bồ hứa hẹn mua sắm tiệm vàng bạc tiệm trang sức bao trọn gói, đến lúc lấy sổ xong mới lộ ra là em gái xin tiền đi chơi với bạn trai.",
        "adaptation": "Khách gọi điện huênh hoang với bồ hứa nhập dây chuyền dập khuôn xuất hóa đơn đỏ VAT đứng tên công ty bồ cho oai, Sếp Thúy check thông tin doanh nghiệp thấy Chủ tịch pháp nhân là bà xã khách, dọa gửi thẳng Zalo bà xã khiến khách toát mồ hôi xin xuất tên xưởng chính (EKA-V2-67)"
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
    
    rows_added = 0
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
        
        rows_added += 1
        print(f"[+] Đã thêm dòng mới File 1: STT {last_stt} | Mã {script['code']} | Row {new_row}")

    success = atomic_save(wb, FILE_1)
    return success

def update_file_2():
    print("\n" + "="*60)
    print(f"[*] ĐANG CẬP NHẬT FILE 2: {os.path.basename(FILE_2)}")
    print("="*60)
    
    wb = openpyxl.load_workbook(FILE_2)
    
    # --- 1. Cập nhật Sheet 1: KICH_BAN_VIRAL_REELS ---
    ws1 = wb['KICH_BAN_VIRAL_REELS']
    current_max_row = ws1.max_row
    
    # Cơ chế tự động xóa kịch bản "Đã sản xuất" nếu có
    rows_to_delete = []
    for r in range(5, current_max_row + 1):
        row_vals = [str(ws1.cell(r, c).value or '') for c in range(1, ws1.max_column + 1)]
        if any("đã sản xuất" in v.lower() or "đã quay" in v.lower() for v in row_vals):
            rows_to_delete.append(r)
            
    if rows_to_delete:
        for r in reversed(rows_to_delete):
            print(f"[!] Xóa kịch bản đã sản xuất tại Row {r}: Mã {ws1.cell(r, 2).value}")
            ws1.delete_rows(r)
            
    # Tính lại STT sau khi kiểm tra/xóa
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
        
        print(f"[+] Đã thêm dòng mới File 2 Sheet 4: STT {last_stt_s4} | {idea['title'][:35]} | Row {new_row}")

    success = atomic_save(wb, FILE_2)
    return success

def run_desktop_sync():
    print("\n" + "="*60)
    print(f"[*] ĐỒNG BỘ BÁO CÁO VÀO BẢNG CHỈ DẪN DESKTOP")
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
    ok1 = update_file_1()
    ok2 = update_file_2()
    if ok1 and ok2:
        print("\n[✓] CẢ 2 FILE EXCEL TRÊN ONEDRIVE ĐÃ ĐƯỢC CẬP NHẬT HOÀN TẤT VÀ AN TOÀN!")
        run_desktop_sync()
    else:
        print("\n[!] Cảnh báo: Có lỗi trong quá trình lưu file.")
