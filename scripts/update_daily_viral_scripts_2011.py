# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_2011.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 20/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-79, 80, 81) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-80, 81, 82) vào Sheet 1 và 3 gợi ý (STT 76, 77, 78) vào Sheet 4 của File 2:
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
        "code": "EKA-REEL-79",
        "title": "Chủ xưởng Bắc Giang chọn mã khách hàng phong thủy và màn kéo 500 anh em (Tham khảo Reel 215K views)",
        "mb_raw": "• Khách (giơ tay phòng thủ): Thôi chị ơi, chị cho em xin chị, em có vợ rồi! Em chỉ đến đây để mở tài khoản doanh nghiệp thôi, chứ có xuống cái xó này để đi bằng ngại đâu!\n• Nhân viên: MB chuyển đổi số, anh mở tài khoản online mà không cần đến trực tiếp đâu anh ạ!\n• Khách: Nhưng cứ phải đến đây mặc cả cho nó chắc, số tài khoản phong thủy... Miễn phí cho anh luôn... Thế chuyển tiền 24/7 nhớ đâu nửa đêm nhập lộn về tên thì sao?\n• Nhân viên: Không đâu anh ơi! Bên em đang có gói ưu đãi dành cho ngành nghề nông nghiệp. Anh ở trên Bắc Giang hợp không được không? Được giảm tối đa tận 2% của anh ạ!\n• Khách: Thế để đợi tôi đi họp cổ động viên đã, giới thiệu cho 500 anh em trên Bắc mai xuống đây!\n• (Sound: Aiyo mama! Cả quầy đứng hình)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách (giơ tay): Thôi! Anh có vợ rồi!\n• Sales (ngớ người): Ơ, em chào khách!\n• Khách: Mở mã doanh nghiệp thôi!\n• [ESCALATION 4-15s]\n• Sales: Mở online trên Portal!\n• Sales: Không cần qua tận nơi!\n• Khách: Phải mặc cả cho chắc!\n• Khách: Xin mã số phong thủy!\n• Sales: Miễn phí toàn bộ luôn!\n• Khách: Nhỡ nửa đêm nhầm đơn?\n• Sếp Thúy: Hệ thống duyệt hai lớp!\n• Sếp Thúy: Đặc quyền xưởng cơ khí!\n• [PUNCHLINE 16-25s]\n• Khách: Anh nhập máy nông nghiệp!\n• Sếp Thúy: Miễn phí báo giá luôn!\n• Khách: Thế đợi anh chút nhé!\n• Sales: Bác gọi ai thế ạ?\n• Khách: Họp hội đồng hương luôn!\n• Khách: Gọi năm trăm anh em!\n• Khách: Kéo hết từ Bắc Giang!\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Kéo cả hội?\n• Sếp Thúy (tỉnh bơ): Vào việc luôn đi anh!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Bác khách chủ xưởng Bắc Giang (giơ tay phòng thủ) + Nữ Sales Eureka (áo đồng phục) + Sếp Thúy (cầm tablet mở Portal ERP).",
        "scene": "Cửa kính sảnh vào & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, tablet mở Portal ERP, tập kẹp file A4, cốc trà.",
        "twist": "Bác chủ xưởng Bắc Giang bước vào văn phòng đã giơ tay thề thốt 'có vợ rồi' vì tưởng được chào đón đặc biệt, hóa ra chỉ muốn mở mã khách hàng doanh nghiệp. Nghe giới thiệu Portal ERP đăng ký online miễn phí mà vẫn đòi đến tận nơi mặc cả chọn mã phong thủy; đến khi Sếp Thúy báo gói hỗ trợ công nợ 10%-40% cho ngành máy móc cơ khí và nông nghiệp, bác đòi họp ngay 500 anh em chủ xưởng Bắc Giang kéo xuống văn phòng.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-80",
        "title": "Khách để quên điện thoại ở quầy chứng từ và màn liên lạc đi vào lòng đất (Tham khảo Reel 47K views)",
        "mb_raw": "• Sếp: Ô, em chưa trả lại điện thoại cho khách à?\n• Nhân viên: Anh ấy về rồi! Anh ấy để quên điện thoại đây này!\n• Sếp: Ô kìa, anh ấy quên điện thoại thì em phải báo người ta đi chứ!\n• Nhân viên: Ơ, em nhắn từ sáng rồi người ta có trả lời em đâu!\n• Sếp: Ơ, nhắn không được thì phải gọi chứ!\n• Nhân viên: Ơ, em gọi có thấy nghe máy đâu!\n• (Sếp bấm gọi thử -> chính chiếc điện thoại để quên trên bàn đổ chuông reng reng!)\n• Sếp: Em làm thế này thì tốn khách của anh lắm!\n• (Sound: Aiyo mama! Cả quầy đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sếp Thúy (cầm máy): Chưa trả điện thoại khách?\n• Chứng từ: Anh ấy về mất rồi!\n• Chứng từ: Để quên trên bàn đây!\n• [ESCALATION 4-15s]\n• Sếp Thúy: Nhắn tin báo người ta!\n• Chứng từ: Em nhắn từ sáng rồi!\n• Chứng từ: Có thấy trả lời đâu!\n• Sếp Thúy: Nhắn không được thì gọi!\n• Chứng từ: Em gọi ba cuộc rồi!\n• Chứng từ: Có thấy nghe máy đâu!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Đưa số chị gọi thử!\n• Chứng từ: Đang đổ chuông đấy Sếp!\n• (Sound: Chuông điện thoại reo vang ngay trên bàn!)\n• Sếp Thúy (đơ 1.5s): Em gọi vào máy nào?\n• Chứng từ (tỉnh bơ): Dạ, đúng máy anh ấy!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Giữ C/O thế này á?\n• Sếp Thúy (tỉnh bơ): Chịu em luôn rồi đấy!\n• (Sound: Aiyo mama! Cả phòng đứng hình)",
        "roles": "Sếp Thúy (nét mặt nghiêm nghị, cầm điện thoại) + Nhân viên Chứng từ Gen Z (ôm tập hồ sơ C/O Form E & hóa đơn VAT, mặt tỉnh bơ).",
        "scene": "Quầy chứng từ & bàn làm việc văn phòng Eureka 44 Triều Khúc, tập hồ sơ C/O Form E, hóa đơn VAT, 2 chiếc điện thoại smartphone.",
        "twist": "Khách hàng ký xong bộ C/O Form E và hóa đơn VAT rồi bỏ quên điện thoại tại quầy chứng từ. Sếp Thúy hỏi sao chưa báo khách, nhân viên Gen Z tỉnh bơ bảo đã nhắn tin từ sáng và gọi ba cuộc đều không thấy ai trả lời — đến khi bấm gọi lại mới phát hiện nhân viên nhắn và gọi vào chính chiếc điện thoại khách đang bỏ quên ngay trên bàn.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-81",
        "title": "Sếp Thúy hỏi tối nay rảnh không và pha quay xe tốc độ ánh sáng của Gen Z (Tham khảo Reel 44K views)",
        "mb_raw": "• Nhân viên 1: Thôi chị đừng nói em ơi đi! Chị nào chị lừa mấy!\n• Nhân viên 2: Mặt em có hiện lên chữ đần đâu mà chị xuống nói em thế! Năm trước chị cũng hỏi em có rảnh không xong chị bắt em tăng ca 2 tiếng á! Em thuộc bài rồi!\n• Sếp: Thế tối nay mấy cô rảnh không?\n• Nhân viên 1: Không, hôm nay bận nhé! Bận lắm! Hôm nay em có việc quan trọng bận lắm!\n• Sếp: Chị định mời cả phòng đi ăn nếu rảnh thôi mà...\n• Nhân viên (quay xe): Rủ cả phòng mình đi ăn nhé! Hà mê ăn lắm rồi!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "eureka_trans": "• [HOOK 0-3s]\n• Nhân viên 1 (giơ file): Sếp đừng lừa bọn em!\n• Nhân viên 2: Mặt em không chữ ngây!\n• Sếp Thúy: Chị đã nói gì đâu?\n• [ESCALATION 4-15s]\n• Nhân viên 1: Tháng trước hỏi rảnh không!\n• Nhân viên 2: Xong bắt tăng ca luôn!\n• Nhân viên 1: Bọn em thuộc bài rồi!\n• Sếp Thúy: Có Sổ tay AI rồi!\n• Sếp Thúy: Tra mã HS mười giây!\n• Sếp Thúy: Nay có rảnh không nào?\n• Nhân viên 1: Không! Em bận lắm nhé!\n• Nhân viên 2: Em có việc cực gấp!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Thế thôi, chị đi vậy!\n• Sếp Thúy: Định khao cả phòng lẩu!\n• Sếp Thúy: Mừng vượt KPI gom cont!\n• Nhân viên 1 (quay xe): Lẩu Thái hả Sếp ơi?\n• Nhân viên 2: Em rảnh nhất công ty!\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Tưởng bận mà?\n• Sếp Thúy (tỉnh bơ): Làm ăn thế đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (bước tới bàn nhân viên, nét mặt điềm tĩnh) + 2 Nhân viên Sales & Chứng từ Gen Z (giơ tập kẹp file A4 phòng thủ).",
        "scene": "Dãy bàn làm việc máy tính tại văn phòng Eureka 44 Triều Khúc, màn hình Sổ tay AI Eureka, tập kẹp file tờ khai hải quan & C/O Form E.",
        "twist": "Sếp Thúy vừa bước tới bàn làm việc, hai nhân viên Gen Z đã giơ tập hồ sơ lên phòng thủ vì sợ bị gài tăng ca như tháng trước. Dù Sếp Thúy giải thích đã có Sổ tay AI tra mã HS trong 10 giây (tự động hóa 80% quy trình lặp lại) không phải thức đêm, cả hai vẫn nhất quyết báo bận việc gấp; vừa nghe Sếp bảo định khao cả phòng ăn lẩu mừng vượt KPI gom cont 1-2 ngày/xe, cả hai quay xe 180 độ nhận mình rảnh nhất công ty.",
        "duration": "23s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-80",
        "motif": "Motif 4 & 3: Hài gia đình (Sợ vợ) & Phong thủy B2B",
        "title": "Chủ xưởng Bắc Giang chọn mã khách hàng phong thủy và màn kéo 500 anh em (Tham khảo Reel 215K views)",
        "strength": "Đăng ký đối tác online 30 giây trên Portal ERP, miễn phí tra cước & mã HS, ưu đãi hạn mức công nợ 10% - 40% cho doanh nghiệp cơ khí & máy móc nông nghiệp.",
        "roles_scene": "• Bối cảnh: Cửa kính sảnh vào & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Bác khách chủ xưởng Bắc Giang + Nữ Sales Eureka + Sếp Thúy\n• Đạo cụ: Tablet mở Portal ERP, tập kẹp file A4, cốc trà.",
        "pingpong": "• [HOOK 0-3s]\n• Khách (giơ tay): Thôi! Anh có vợ rồi!\n• Sales (ngớ người): Ơ, em chào khách!\n• Khách: Mở mã doanh nghiệp thôi!\n• [ESCALATION 4-15s]\n• Sales: Mở online trên Portal!\n• Sales: Không cần qua tận nơi!\n• Khách: Phải mặc cả cho chắc!\n• Khách: Xin mã số phong thủy!\n• Sales: Miễn phí toàn bộ luôn!\n• Khách: Nhỡ nửa đêm nhầm đơn?\n• Sếp Thúy: Hệ thống duyệt hai lớp!\n• Sếp Thúy: Đặc quyền xưởng cơ khí!\n• [PUNCHLINE 16-25s]\n• Khách: Anh nhập máy nông nghiệp!\n• Sếp Thúy: Miễn phí báo giá luôn!\n• Khách: Thế đợi anh chút nhé!\n• Sales: Bác gọi ai thế ạ?\n• Khách: Họp hội đồng hương luôn!\n• Khách: Gọi năm trăm anh em!\n• Khách: Kéo hết từ Bắc Giang!\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Kéo cả hội?\n• Sếp Thúy (tỉnh bơ): Vào việc luôn đi anh!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Bác chủ xưởng Bắc Giang bước vào văn phòng đã giơ tay thề thốt 'có vợ rồi' vì tưởng được chào đón đặc biệt, hóa ra chỉ muốn mở mã khách hàng doanh nghiệp. Nghe giới thiệu Portal ERP đăng ký online miễn phí mà vẫn đòi đến tận nơi mặc cả chọn mã phong thủy; đến khi Sếp Thúy báo gói hỗ trợ công nợ 10%-40% cho ngành máy móc cơ khí và nông nghiệp, bác đòi họp ngay 500 anh em chủ xưởng Bắc Giang kéo xuống văn phòng.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-81",
        "motif": "Motif 1: Slapstick âm thanh & Đạo cụ văn phòng",
        "title": "Khách để quên điện thoại ở quầy chứng từ và màn liên lạc đi vào lòng đất (Tham khảo Reel 47K views)",
        "strength": "Quy trình bàn giao chứng từ gốc (C/O Form E, Hóa đơn VAT 100%, Tờ khai hải quan) chuẩn xác & Dịch vụ chăm sóc khách hàng Dedicated Team 1-1.",
        "roles_scene": "• Bối cảnh: Quầy chứng từ & bàn làm việc văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Nhân viên Chứng từ Gen Z\n• Đạo cụ: Tập hồ sơ C/O Form E, hóa đơn VAT, 2 chiếc điện thoại smartphone.",
        "pingpong": "• [HOOK 0-3s]\n• Sếp Thúy (cầm máy): Chưa trả điện thoại khách?\n• Chứng từ: Anh ấy về mất rồi!\n• Chứng từ: Để quên trên bàn đây!\n• [ESCALATION 4-15s]\n• Sếp Thúy: Nhắn tin báo người ta!\n• Chứng từ: Em nhắn từ sáng rồi!\n• Chứng từ: Có thấy trả lời đâu!\n• Sếp Thúy: Nhắn không được thì gọi!\n• Chứng từ: Em gọi ba cuộc rồi!\n• Chứng từ: Có thấy nghe máy đâu!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Đưa số chị gọi thử!\n• Chứng từ: Đang đổ chuông đấy Sếp!\n• (Sound: Chuông điện thoại reo vang ngay trên bàn!)\n• Sếp Thúy (đơ 1.5s): Em gọi vào máy nào?\n• Chứng từ (tỉnh bơ): Dạ, đúng máy anh ấy!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Giữ C/O thế này á?\n• Sếp Thúy (tỉnh bơ): Chịu em luôn rồi đấy!\n• (Sound: Aiyo mama! Cả phòng đứng hình)",
        "twist": "Khách hàng ký xong bộ C/O Form E và hóa đơn VAT rồi bỏ quên điện thoại tại quầy chứng từ. Sếp Thúy hỏi sao chưa báo khách, nhân viên Gen Z tỉnh bơ bảo đã nhắn tin từ sáng và gọi ba cuộc đều không thấy ai trả lời — đến khi bấm gọi lại mới phát hiện nhân viên nhắn và gọi vào chính chiếc điện thoại khách đang bỏ quên ngay trên bàn.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-82",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales & Chứng từ Gen Z)",
        "title": "Sếp Thúy hỏi tối nay rảnh không và pha quay xe tốc độ ánh sáng của Gen Z (Tham khảo Reel 44K views)",
        "strength": "Hệ thống Sổ tay AI Eureka hỗ trợ tự động hóa 80% quy trình lặp lại, tra mã HS từ 15 phút xuống 10 giây, đảm bảo tiến độ gom cont LCL 1-2 ngày/xe không quá tải.",
        "roles_scene": "• Bối cảnh: Dãy bàn làm việc máy tính tại văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + 2 Nhân viên Sales & Chứng từ Gen Z\n• Đạo cụ: Màn hình Sổ tay AI Eureka, tập kẹp file tờ khai hải quan & C/O Form E.",
        "pingpong": "• [HOOK 0-3s]\n• Nhân viên 1 (giơ file): Sếp đừng lừa bọn em!\n• Nhân viên 2: Mặt em không chữ ngây!\n• Sếp Thúy: Chị đã nói gì đâu?\n• [ESCALATION 4-15s]\n• Nhân viên 1: Tháng trước hỏi rảnh không!\n• Nhân viên 2: Xong bắt tăng ca luôn!\n• Nhân viên 1: Bọn em thuộc bài rồi!\n• Sếp Thúy: Có Sổ tay AI rồi!\n• Sếp Thúy: Tra mã HS mười giây!\n• Sếp Thúy: Nay có rảnh không nào?\n• Nhân viên 1: Không! Em bận lắm nhé!\n• Nhân viên 2: Em có việc cực gấp!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Thế thôi, chị đi vậy!\n• Sếp Thúy: Định khao cả phòng lẩu!\n• Sếp Thúy: Mừng vượt KPI gom cont!\n• Nhân viên 1 (quay xe): Lẩu Thái hả Sếp ơi?\n• Nhân viên 2: Em rảnh nhất công ty!\n• [DEFLATION 26-30s]\n• Sếp Thúy (đơ 1.5s): Ơ hay! Tưởng bận mà?\n• Sếp Thúy (tỉnh bơ): Làm ăn thế đấy!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Sếp Thúy vừa bước tới bàn làm việc, hai nhân viên Gen Z đã giơ tập hồ sơ lên phòng thủ vì sợ bị gài tăng ca như tháng trước. Dù Sếp Thúy giải thích đã có Sổ tay AI tra mã HS trong 10 giây (tự động hóa 80% quy trình lặp lại) không phải thức đêm, cả hai vẫn nhất quyết báo bận việc gấp; vừa nghe Sếp bảo định khao cả phòng ăn lẩu mừng vượt KPI gom cont 1-2 ngày/xe, cả hai quay xe 180 độ nhận mình rảnh nhất công ty.",
        "duration": "23s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 76,
        "title": "Chủ xưởng Bắc Giang chọn mã khách hàng phong thủy và màn kéo 500 anh em (Reel 215K views)",
        "views": "215,000 views",
        "motif": "Motif 4 & 3: Hài gia đình (Sợ vợ) & Phong thủy B2B",
        "trigger": "Khách bước vào quầy đã giơ tay thề 'có vợ rồi', đòi trực tiếp mặc cả chọn số tài khoản phong thủy dù mở được online, nghe có gói ưu đãi ngành nghề liền đòi kéo 500 anh em đồng hương xuống.",
        "adaptation": "Bác chủ xưởng Bắc Giang vào Eureka phòng thủ 'có vợ rồi', đòi chọn mã khách hàng phong thủy trên Portal ERP; nghe Sếp Thúy báo gói công nợ 10%-40% ngành máy cơ khí & nông nghiệp liền đòi kéo 500 anh em xuống văn phòng (EKA-V2-80)."
    },
    {
        "stt": 77,
        "title": "Khách để quên điện thoại ở quầy chứng từ và màn liên lạc đi vào lòng đất (Reel 47K views)",
        "views": "47,000 views",
        "motif": "Motif 1: Slapstick âm thanh & Đạo cụ văn phòng",
        "trigger": "Khách để quên điện thoại ở quầy, quản lý hỏi sao chưa báo thì nhân viên tỉnh bơ bảo đã nhắn và gọi nhưng không nghe máy — hóa ra gọi vào chính máy khách đang nằm trên bàn.",
        "adaptation": "Khách ký xong bộ C/O Form E & hóa đơn VAT bỏ quên điện thoại ở quầy chứng từ Eureka; nhân viên Gen Z bảo đã nhắn tin và gọi ba cuộc không ai nghe, Sếp Thúy bấm gọi thử thì chiếc máy trên bàn reo vang (EKA-V2-81)."
    },
    {
        "stt": 78,
        "title": "Sếp Thúy hỏi tối nay rảnh không và pha quay xe tốc độ ánh sáng của Gen Z (Reel 44K views)",
        "views": "44,000 views",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales & Chứng từ Gen Z)",
        "trigger": "Sếp hỏi tối nay rảnh không, nhân viên cảnh giác tưởng bị gài tăng ca như năm ngoái nên kiên quyết báo bận; nghe sếp bảo mời cả phòng đi ăn liền quay xe nhận rảnh nhất.",
        "adaptation": "Sếp Thúy hỏi tối nay rảnh không, hai nhân viên Gen Z giơ file C/O phòng thủ sợ tăng ca dù có Sổ tay AI tra mã HS 10 giây; nghe Sếp mời lẩu mừng vượt KPI gom cont 1-2 ngày/xe liền quay xe rảnh nhất công ty (EKA-V2-82)."
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
