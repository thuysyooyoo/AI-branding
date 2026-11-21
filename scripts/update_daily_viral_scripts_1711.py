# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_1711.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 17/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-73, 74, 75) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-74, 75, 76) vào Sheet 1 và 3 gợi ý (STT 70, 71, 72) vào Sheet 4 của File 2:
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
    "dở à", "đếch", "điên à", "bắt cóc", "đi tù", "công an hốt"
]

# Dữ liệu 3 kịch bản mới cho File 1 (BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx)
SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-73",
        "title": "Chiếc ghế xoay bị điếc chọn lọc ở Eureka và màn đổi chỗ đòi báo cáo (Tham khảo Reel 41K views)",
        "mb_raw": "• Nhân viên: Chị ơi cho em vay chục triệu để gửi tiết kiệm! Em mới có bốn chục mà lãi suất thì đang ưu đãi, 50 triệu 8.2!\n• Chị đồng nghiệp: Hả? Cái gì? Tao không nghe thấy cái gì!\n• Nhân viên: Cho em vay 10 triệu!\n• Chị đồng nghiệp: Hả? Cái ghế này làm sao ấy, ngồi lên không nghe thấy cái gì cả!\n• Nhân viên: Ơ cái ghế làm sao? Chị đi ra em ngồi xem nào! Nghe rõ mồn một thế này, cả tiếng điều hòa kêu cũng thấy mà!\n• Chị đồng nghiệp: Thế mày trả chị 5 triệu tháng trước mua bàn đi!\n• Nhân viên: Hả? Năm triệu tháng trước đấy! Công nhận cái ghế này nó lạ, từ lúc em ngồi vào đây em cũng không nghe thấy gì luôn!\n• Chị đồng nghiệp: Thôi mày đi ra chị ngồi xem nào!\n• (Sound: Aiyo mama! Cả phòng cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales: Sếp duyệt công nợ!\n• Sếp Thúy: Hả? Gì cơ?\n• Sales: Bốn mươi phần trăm!\n• Sếp Thúy: Không nghe thấy gì!\n• [ESCALATION 4-15s]\n• Sales: Khách chưa đủ hồ sơ!\n• Sếp Thúy: Ghế này bị điếc!\n• Sales: Ơ hay! Ghế nào điếc?\n• Sếp Thúy: Ngồi lên mất tiếng!\n• Sales: Đâu, em ngồi thử!\n• Sales: Nghe rõ mồn một!\n• Sales: Tiếng máy in luôn!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Thế năm bộ Form E?\n• Sales: Hả? Cái gì cơ?\n• Sếp Thúy: Nộp báo cáo ERP!\n• Sales (đơ 1.5s): Công nhận ghế điếc thật!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Đứng dậy làm việc!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (ngồi ghế xoay tại bàn làm việc giám đốc) + Sales Gen Z (cầm kẹp file đề xuất công nợ A4).",
        "scene": "Bàn làm việc Sếp Thúy tại văn phòng Eureka 44 Triều Khúc, chiếc ghế xoay văn phòng, máy in góc phòng, tập kẹp hồ sơ C/O Form E.",
        "twist": "Sales xin duyệt hạn mức công nợ 40% cho khách chưa đủ hồ sơ, Sếp Thúy ngồi trên ghế xoay giả vờ không nghe thấy vì 'ghế bị điếc'; Sales không tin đòi ngồi thử chứng minh ghế nghe rõ tiếng máy in, nhưng vừa nghe Sếp hỏi 5 bộ C/O Form E và báo cáo ERP thì Sales cũng tỉnh bơ 'Công nhận ghế điếc thật!'.",
        "duration": "22s"
    },
    {
        "code": "EKA-REEL-74",
        "title": "Khách tấm tắc khen chuyên viên Eureka xinh gái và cú chỉ tay phũ phàng (Tham khảo Reel 59K views)",
        "mb_raw": "• Khách: Quầy này chuyên mấy em xinh gái nhanh nhẹn, tư vấn lãi suất ưu đãi lắm à?\n• Nhân viên: Dạ anh ơi từ 10 triệu được lãi suất 8.2 rồi ạ!\n• Khách: Ừ, mấy bạn ấy xinh gái nhỉ!\n• Nhân viên: Ôi giời cũng bình thường thôi anh ạ! Được cái nhanh nhẹn chuyên nghiệp ạ!\n• Khách: Ừ xinh mà!\n• Nhân viên: Dạ lấy lại chuyên nghiệp của app nữa anh ạ!\n• Khách: Anh thấy ở ngoài cũng xinh mà!\n• Nhân viên: Ấy chết nên hôm nay anh mới đến đây gửi tiết kiệm ạ? Người anh nói thế em ngại lắm!\n• Khách: Ừ, nhưng mà bạn xinh gái cái bàn đằng sau ấy nhỉ?\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách: Quầy này toàn xinh gái?\n• Sales nữ: Dạ em cảm ơn!\n• [ESCALATION 4-15s]\n• Khách: Gom cont từ một khối?\n• Sales nữ: Chạy xe hàng ngày!\n• Khách: Phòng XNK năm triệu?\n• Sales nữ: Trọn gói một một!\n• Khách: Vừa giỏi vừa xinh nhỉ!\n• Sales nữ: Ôi giời bình thường thôi!\n• Khách: Ở ngoài còn xinh hơn!\n• Sales nữ: Anh nói em ngại quá!\n• [PUNCHLINE 16-25s]\n• Khách: Hôm nay ký hợp đồng!\n• Sales nữ: Dạ em đón anh!\n• Khách: Cho gặp bạn xinh kia!\n• Sales nữ (đơ 1.5s): Ơ hay! Bạn nào ạ?\n• Khách (chỉ tay ra sau): Bạn chứng từ bàn sau!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Chứng từ chuẩn mới chốt!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Bác khách chủ xưởng cơ khí (vui tính) + Sales nữ Eureka (áo đồng phục, cười e thẹn) + Chuyên viên Chứng từ (ngồi bàn phía sau ôm xấp tờ khai) + Sếp Thúy.",
        "scene": "Quầy tiếp khách & dãy bàn chứng từ phía sau tại văn phòng Eureka 44 Triều Khúc, tập hồ sơ tờ khai hải quan A4, cốc nước.",
        "twist": "Khách bước vào văn phòng tấm tắc khen nhân viên tư vấn gom cont 1 khối và phòng XNK 5 triệu vừa giỏi vừa xinh, ở ngoài còn xinh hơn trên livestream khiến nữ Sales e thẹn nhận vơ — đến lúc chốt ký hợp đồng thì khách chỉ thẳng sang bạn Chứng từ ngồi bàn phía sau.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-75",
        "title": "Nhờ khách quay hình ảnh uy tín nhất tại Eureka và cái kết tự sướng khó đỡ (Tham khảo Reel 47K views)",
        "mb_raw": "• Nhân viên: Ôi anh ơi! Cũng thích quay video à? Truyền thông đấy! Ngày nào chả quay cho vui dân vui xóm! Ôi may quá, bên em đang có chương trình ưu đãi cần truyền thông ạ!\n• Khách: Chương trình gì?\n• Nhân viên: MB 32! Gửi tiết kiệm từ 20 triệu kỳ hạn 6 tháng trên app MB nhận ngay 1 lượt quay ong vàng, cơ hội trúng sổ tiết kiệm 3,2 tỷ, chuyến du lịch Thổ Nhĩ Kỳ và iPhone 18 sắp ra mắt! Anh quay giùm em hình ảnh đẹp nhất có trong ngân hàng để em truyền thông với!\n• Khách: Hình ảnh đẹp nhất á? Ừ đưa đây! Đây, xong rồi đấy!\n• Nhân viên: Em cảm ơn ạ! Ủa... Thế này mà?!\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales: Bác thích quay phim?\n• Khách: Đam mê nghệ thuật!\n• [ESCALATION 4-15s]\n• Sales: Quay giúp Eureka nhé!\n• Khách: Truyền thông dịch vụ?\n• Sales: Giám định máy cũ!\n• Khách: Quyết định mười tám?\n• Sales: Chuẩn hóa đơn VAT!\n• Khách: Hơn ngàn mặt hàng?\n• Sales: Quay góc uy tín!\n• Sales: Hình ảnh đẹp nhất!\n• [PUNCHLINE 16-25s]\n• Khách: Đưa máy bác quay!\n• Khách (nhại slow-motion): Nghệ thuật là đây!\n• Khách: Xong rồi đấy em!\n• Sales: Dạ em xem nào!\n• Sales (đơ 1.5s): Ơ hay! Mặt bác à?\n• [DEFLATION 26-30s]\n• Khách (tỉnh bơ): Khách hàng là đẹp nhất!\n• Sếp Thúy: Bác tự tin thật đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sales Eureka (cầm smartphone) + Bác khách chủ xưởng (đeo kính, tạo dáng đạo diễn quay slow-motion) + Sếp Thúy (ngồi bàn họp nhìn màn hình điện thoại).",
        "scene": "Góc bàn họp & quầy chứng từ văn phòng Eureka 44 Triều Khúc, điện thoại smartphone, tập hồ sơ giám định Quyết định 18, biển bảng thương hiệu Eureka.",
        "twist": "Sales nhờ bác khách quay giúp 'hình ảnh đẹp và uy tín nhất tại Eureka' để truyền thông dịch vụ giám định máy cũ Quyết định 18 và 100% hóa đơn VAT; bác khách tạo dáng quay slow-motion đậm chất điện ảnh rồi trả máy, hóa ra bác chỉ quay cận cảnh tự sướng chính gương mặt mình và khẳng định 'Khách hàng là đẹp nhất!'.",
        "duration": "22s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-74",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z)",
        "title": "Chiếc ghế xoay bị điếc chọn lọc ở Eureka và màn đổi chỗ đòi báo cáo (Tham khảo Reel 41K views)",
        "strength": "Cấp hạn mức công nợ linh hoạt 10% - 40% hỗ trợ dòng tiền mùa cao điểm & Hệ thống Portal ERP quản lý chứng từ, C/O Form E minh bạch.",
        "roles_scene": "• Bối cảnh: Bàn làm việc Sếp Thúy tại văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Gen Z\n• Đạo cụ: Chiếc ghế xoay văn phòng, máy in góc phòng, tập kẹp hồ sơ C/O Form E A4.",
        "pingpong": "• [HOOK 0-3s]\n• Sales: Sếp duyệt công nợ!\n• Sếp Thúy: Hả? Gì cơ?\n• Sales: Bốn mươi phần trăm!\n• Sếp Thúy: Không nghe thấy gì!\n• [ESCALATION 4-15s]\n• Sales: Khách chưa đủ hồ sơ!\n• Sếp Thúy: Ghế này bị điếc!\n• Sales: Ơ hay! Ghế nào điếc?\n• Sếp Thúy: Ngồi lên mất tiếng!\n• Sales: Đâu, em ngồi thử!\n• Sales: Nghe rõ mồn một!\n• Sales: Tiếng máy in luôn!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Thế năm bộ Form E?\n• Sales: Hả? Cái gì cơ?\n• Sếp Thúy: Nộp báo cáo ERP!\n• Sales (đơ 1.5s): Công nhận ghế điếc thật!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Đứng dậy làm việc!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Sales xin duyệt hạn mức công nợ 40% cho khách chưa đủ hồ sơ, Sếp Thúy ngồi trên ghế xoay giả vờ không nghe thấy vì 'ghế bị điếc'; Sales không tin đòi ngồi thử chứng minh ghế nghe rõ tiếng máy in, nhưng vừa nghe Sếp hỏi 5 bộ C/O Form E và báo cáo ERP thì Sales cũng tỉnh bơ 'Công nhận ghế điếc thật!'.",
        "duration": "22s"
    },
    {
        "code": "EKA-V2-75",
        "motif": "Motif 5: Hiểu nhầm công sở & Tưởng bở nhận vơ",
        "title": "Khách tấm tắc khen chuyên viên Eureka xinh gái và cú chỉ tay phũ phàng (Tham khảo Reel 59K views)",
        "strength": "Phòng Xuất nhập khẩu thuê ngoài trọn gói từ 5 triệu/tháng (Ekip 1-1: Quản lý + CSKH + Chứng từ) & Gom cont LCL từ 1 khối (1 CBM) xe chạy hàng ngày.",
        "roles_scene": "• Bối cảnh: Quầy tiếp khách & dãy bàn chứng từ phía sau tại văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Bác khách chủ xưởng + Sales nữ Eureka + Chuyên viên Chứng từ + Sếp Thúy\n• Đạo cụ: Tập hồ sơ tờ khai hải quan A4, cốc nước.",
        "pingpong": "• [HOOK 0-3s]\n• Khách: Quầy này toàn xinh gái?\n• Sales nữ: Dạ em cảm ơn!\n• [ESCALATION 4-15s]\n• Khách: Gom cont từ một khối?\n• Sales nữ: Chạy xe hàng ngày!\n• Khách: Phòng XNK năm triệu?\n• Sales nữ: Trọn gói một một!\n• Khách: Vừa giỏi vừa xinh nhỉ!\n• Sales nữ: Ôi giời bình thường thôi!\n• Khách: Ở ngoài còn xinh hơn!\n• Sales nữ: Anh nói em ngại quá!\n• [PUNCHLINE 16-25s]\n• Khách: Hôm nay ký hợp đồng!\n• Sales nữ: Dạ em đón anh!\n• Khách: Cho gặp bạn xinh kia!\n• Sales nữ (đơ 1.5s): Ơ hay! Bạn nào ạ?\n• Khách (chỉ tay ra sau): Bạn chứng từ bàn sau!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Chứng từ chuẩn mới chốt!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Khách bước vào văn phòng tấm tắc khen nhân viên tư vấn gom cont 1 khối và phòng XNK 5 triệu vừa giỏi vừa xinh, ở ngoài còn xinh hơn trên livestream khiến nữ Sales e thẹn nhận vơ — đến lúc chốt ký hợp đồng thì khách chỉ thẳng sang bạn Chứng từ ngồi bàn phía sau.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-76",
        "motif": "Motif 6: Nhại điện ảnh Hollywood & Đạo cụ truyền thông",
        "title": "Nhờ khách quay hình ảnh uy tín nhất tại Eureka và cái kết tự sướng khó đỡ (Tham khảo Reel 47K views)",
        "strength": "Tư vấn & thực hiện giám định máy móc cũ Quyết định 18, 100% Hóa đơn VAT chính ngạch, tư vấn mã HS & C/O Form E chuẩn chỉ cho >1.000 mặt hàng.",
        "roles_scene": "• Bối cảnh: Góc bàn họp & quầy chứng từ văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sales Eureka + Bác khách chủ xưởng (tạo dáng đạo diễn) + Sếp Thúy\n• Đạo cụ: Điện thoại smartphone, tập hồ sơ giám định Quyết định 18, biển bảng thương hiệu Eureka.",
        "pingpong": "• [HOOK 0-3s]\n• Sales: Bác thích quay phim?\n• Khách: Đam mê nghệ thuật!\n• [ESCALATION 4-15s]\n• Sales: Quay giúp Eureka nhé!\n• Khách: Truyền thông dịch vụ?\n• Sales: Giám định máy cũ!\n• Khách: Quyết định mười tám?\n• Sales: Chuẩn hóa đơn VAT!\n• Khách: Hơn ngàn mặt hàng?\n• Sales: Quay góc uy tín!\n• Sales: Hình ảnh đẹp nhất!\n• [PUNCHLINE 16-25s]\n• Khách: Đưa máy bác quay!\n• Khách (nhại slow-motion): Nghệ thuật là đây!\n• Khách: Xong rồi đấy em!\n• Sales: Dạ em xem nào!\n• Sales (đơ 1.5s): Ơ hay! Mặt bác à?\n• [DEFLATION 26-30s]\n• Khách (tỉnh bơ): Khách hàng là đẹp nhất!\n• Sếp Thúy: Bác tự tin thật đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Sales nhờ bác khách quay giúp 'hình ảnh đẹp và uy tín nhất tại Eureka' để truyền thông dịch vụ giám định máy cũ Quyết định 18 và 100% hóa đơn VAT; bác khách tạo dáng quay slow-motion đậm chất điện ảnh rồi trả máy, hóa ra bác chỉ quay cận cảnh tự sướng chính gương mặt mình và khẳng định 'Khách hàng là đẹp nhất!'.",
        "duration": "22s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 70,
        "title": "Chiếc ghế xoay bị điếc chọn lọc ở Eureka và màn đổi chỗ đòi báo cáo (Reel 41K views)",
        "views": "41,000 views",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z)",
        "trigger": "Nhân viên hỏi vay tiền thì người ngồi ghế giả điếc 'cái ghế này bị điếc không nghe thấy gì'; đến khi nhân viên đòi ngồi thử để chứng minh ghế bình thường và bị hỏi đòi nợ ngược lại thì cũng giả điếc y hệt.",
        "adaptation": "Sales Gen Z xin duyệt hạn mức công nợ 40% cho khách chưa đủ hồ sơ, Sếp Thúy ngồi trên ghế xoay bảo ghế bị điếc mất tiếng; Sales ngồi thử thấy nghe rõ tiếng máy in nhưng vừa nghe Sếp hỏi 5 bộ C/O Form E và báo cáo ERP thì tỉnh bơ 'Công nhận ghế điếc thật!' (EKA-V2-74)."
    },
    {
        "stt": 71,
        "title": "Khách tấm tắc khen chuyên viên Eureka xinh gái và cú chỉ tay phũ phàng (Reel 59K views)",
        "views": "59,000 views",
        "motif": "Motif 5: Hiểu nhầm công sở & Tưởng bở nhận vơ",
        "trigger": "Khách hàng vào quầy tấm tắc khen nhân viên xinh gái, chuyên nghiệp, ở ngoài còn xinh hơn trên app khiến nữ nhân viên ngượng ngùng tự nhận — đến cuối khách mới chỉ ra bạn đồng nghiệp ngồi bàn phía sau.",
        "adaptation": "Khách đến quầy tiếp khách khen dịch vụ gom cont 1 khối và phòng XNK 5 triệu vừa giỏi vừa xinh, ở ngoài còn xinh hơn trên livestream khiến Sales nữ e thẹn nhận vơ; đến lúc đòi gặp để ký hợp đồng lại chỉ thẳng bạn Chứng từ ngồi bàn phía sau (EKA-V2-75)."
    },
    {
        "stt": 72,
        "title": "Nhờ khách quay hình ảnh uy tín nhất tại Eureka và cái kết tự sướng khó đỡ (Reel 47K views)",
        "views": "47,000 views",
        "motif": "Motif 6: Nhại điện ảnh Hollywood & Đạo cụ truyền thông",
        "trigger": "Nhân viên quảng bá chương trình ưu đãi lớn rồi nhờ khách quay giúp hình ảnh đẹp nhất tại chi nhánh để truyền thông — khách gật đầu cầm máy quay xong trả lại hóa ra quay trớt quớt / tự sướng.",
        "adaptation": "Sales nhờ bác khách quay thước phim hình ảnh uy tín nhất tại Eureka để truyền thông giám định máy cũ Quyết định 18 và 100% hóa đơn VAT; bác khách tạo dáng quay slow-motion như đạo diễn rồi quay cận mặt tự sướng của chính mình (EKA-V2-76)."
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
