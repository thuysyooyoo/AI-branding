# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_1611.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 16/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-70, 71, 72) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-71, 72, 73) vào Sheet 1 và 3 gợi ý vào Sheet 4 của File 2:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx
3. Tự động kiểm tra và xóa kịch bản có ghi chú "Đã sản xuất" hoặc "Đã quay" (nếu có).
4. Áp dụng kỹ thuật lưu nguyên tử (Atomic Save) qua file đệm .tmp.xlsx và os.replace.
5. Kích hoạt generator_chi_dan.py để đồng bộ Desktop dashboard.
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

# Dữ liệu 3 kịch bản mới cho File 1 (BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx)
SCRIPTS_FILE_1 = [
    {
        "code": "EKA-REEL-70",
        "title": "Tổng tài bước vào Eureka đàm phán cont và cú bóc mẽ chiếc bụng bia (Tham khảo Reel 269K views)",
        "mb_raw": "• Nhân viên: Em chào tổng tài! Nhìn cái mức xin phô là biết tổng tài nghiêng trời lệch đất rồi! Đẹp trai thế chứ lị! Tàu mạn sao đấy, nhìn cái tướng như thế bụng phệ rồi!\n• Khách: Má má đã bảo bụng thôi! Tiền vào đấy nhé, lãi suất thế nào em nhé! Hôm nay ngày vàng lãi suất 50 triệu 8.2, 1 tỷ 9.1 đấy nhé!\n• Khách: Ừ, làm nhanh cho anh, mở sổ năm tỷ nhé!\n• Nhân viên: Em thấy đẹp trai quá, tổng tài quá, đi từ sáng đến giờ như thế...\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách: Tổng tài tới đây!\n• Sales: Em chào Chủ tịch!\n• Khách: Ký cont hỏa tốc!\n• Sếp Thúy: Mời Bác vào bàn!\n• [ESCALATION 4-15s]\n• Khách: Bên cô uy tín?\n• Sếp Thúy: Mười năm pháp nhân!\n• Khách: Xe chạy thế nào?\n• Sales: Chạy liên vận ngày!\n• Khách: Soi hàng ở đâu?\n• Sales: Camera Bằng Tường!\n• Khách: Máy móc an toàn?\n• Sếp Thúy: Niêm phong nguyên đai!\n• Khách: Bụng tôi phong độ?\n• [PUNCHLINE 16-25s]\n• Sales: Bụng bia quá khổ!\n• Khách (đơ 1.5s): Ơ hay! Đừng soi!\n• Khách: Bụng chứa hợp đồng!\n• Khách: Chứ bia bọt gì!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Ký ngay cho thon!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (ngồi bàn làm việc giám đốc) + Sales Eureka (cầm catalogue và kẹp file A4) + Bác khách chủ xưởng (mặc vest bóng bẩy, đeo kính râm, bụng bia to tròn).",
        "scene": "Quầy lễ tân & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, catalogue thiết bị cơ khí, kẹp hồ sơ A4, con dấu công ty.",
        "twist": "Khách diện vest bảnh bao tự nhận phong thái tổng tài bước vào văn phòng Eureka đòi ký cont hỏa tốc, hỏi kiểm tra nghiệp vụ Eureka cực gắt xong bất ngờ hỏi Sales xem bụng mình có phong độ không, Sales tỉnh bơ chê bụng bia quá khổ khiến khách giật bắn mình chữa thẹn bụng chứa hợp đồng cont, Sếp Thúy giục ký nhanh cho thon bụng.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-71",
        "title": "Hai nhiệm vụ tối thượng của Sales Eureka và câu trả lời né KPI của Gen Z (Tham khảo Reel 261K views)",
        "mb_raw": "• Nhân viên: Chào mừng mọi người đến với MB Hải Dương!\n• Nhân viên: Ở đây chỉ có 2 công việc: 1 làm video ngắn giúp khách hàng vui cười, cười càng nhiều càng tốt! 2 giúp khách hàng làm sổ tiết kiệm, sổ càng to quà càng nhiều!\n• Quản lý: Thế ở đây làm giận nhiều không đạt KPI thì sao?\n• Nhân viên: Nói xin...\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales 1: Chào mừng đến Eureka!\n• Sales 2: Hai nhiệm vụ lớn!\n• [ESCALATION 4-15s]\n• Sales 1: Một là gom cont!\n• Sales 2: Từ một khối nhé!\n• Sales 1: Hai là phòng XNK!\n• Sales 2: Trọn gói năm triệu!\n• Sales 1: Hóa đơn chuẩn chỉnh!\n• Sales 2: Miễn phí Form E!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Thế KPI đâu?\n• Sales 1 (đơ 1.5s): Dạ... Sếp Thúy ơi!\n• Sếp Thúy: Chưa đủ mười cont?\n• Sales 2: Em xin tăng ca!\n• Sales 1: Xin nợ tháng sau!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Chịu hai đứa luôn!\n• Sếp Thúy: Vào bàn làm việc!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (khoanh tay đứng sau) + 2 Sales Gen Z (cầm điện thoại quay video và kẹp file báo cáo công việc).",
        "scene": "Góc làm việc máy tính & quầy lễ tân văn phòng Eureka 44 Triều Khúc, điện thoại smartphone gắn tripod, kẹp tài liệu báo cáo KPI, cốc trà.",
        "twist": "Hai Sales Gen Z quay video giới thiệu văn phòng Eureka cực kỳ hào nhoáng với 2 nhiệm vụ tối thượng (gom cont 1 khối và phòng XNK 5 triệu), Sếp Thúy bất thình lình bước ra hỏi chỉ tiêu KPI 10 cont tháng này đâu thì cả hai nhanh nhảu xin tăng ca và xin khất nợ KPI sang tháng sau khiến Sếp Thúy cạn lời.",
        "duration": "22s"
    },
    {
        "code": "EKA-REEL-72",
        "title": "Hỏi cước mười cont máy móc như đại gia và cái kết trải nghiệm dịch vụ (Tham khảo Reel 3.6M views)",
        "mb_raw": "• Khách: Lãi bây giờ bao nhiêu?\n• Nhân viên: Tám phẩy hai!\n• Khách: Một tỷ?\n• Nhân viên: Chín phẩy một ạ! Thẻ tín dụng tặng con gấu! Số tài khoản phong thủy miễn phí cho anh luôn!\n• Khách: Có rút tiền 24/7 được không?\n• Nhân viên: Có ạ!\n• Khách: Có nhạc chuông không?\n• Nhân viên: Có luôn!\n• Nhân viên: Dạ anh ơi hôm nay định gửi bao nhiêu tiền ạ?\n• Khách: Dạ có tiền đâu mà gửi! Tư vấn tốt, dịch vụ tốt, nhạc chuông nhà đẹp đấy!\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách: Cước cont bao nhiêu?\n• Sales: Giá tối ưu nhất!\n• Khách: Mười cont máy tiện?\n• Sếp Thúy: Giảm sâu cho Bác!\n• [ESCALATION 4-15s]\n• Khách: Có công nợ không?\n• Sếp Thúy: Cấp bốn mươi phần!\n• Khách: Camera Bằng Tường?\n• Sales: Soi nét 24/7!\n• Khách: Hóa đơn đỏ chuẩn?\n• Sếp Thúy: VAT chính ngạch!\n• Khách: Xe chạy hàng ngày?\n• Sales: Đúng giờ xuất bến!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Bác book mấy cont?\n• Khách (đơ 1.5s): Tôi chưa có xưởng!\n• Sales: Ơ hay! Thật ạ?\n• Khách: Đang chờ trúng số!\n• Khách: Thấy dịch vụ tốt!\n• Khách: Nên vào hỏi thử!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Bao giờ trúng hẵng hỏi!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (chủ trì tư vấn) + Sales Eureka (cầm tablet Portal ERP) + Bác khách trung niên (tay lắc chìa khóa xe, phong thái đại gia hỏi dồn dập).",
        "scene": "Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, tách trà nóng, catalogue máy tiện CNC, hợp đồng vận chuyển A4.",
        "twist": "Khách vào văn phòng hỏi dồn dập các chính sách VIP cho 10 cont máy tiện CNC như đại gia thực thụ (công nợ 40%, camera Bằng Tường, VAT chính ngạch); Sếp Thúy chuẩn bị hợp đồng hỏi book mấy cont thì khách cười trừ bảo chưa có xưởng đang chờ trúng vé số vào hỏi thử xem dịch vụ tốt như lời đồn không.",
        "duration": "24s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-71",
        "motif": "Motif 1: Slapstick hành động & Đạo cụ văn phòng",
        "title": "Tổng tài bước vào Eureka đàm phán cont và cú bóc mẽ chiếc bụng bia (Tham khảo Reel 269K views)",
        "strength": "Xe cont liên vận quốc tế chạy hàng ngày tuyến Bằng Tường - Hà Nội, thông quan hỏa tốc 24h, bảo vệ pháp nhân sạch 10 năm, kiểm đếm niêm phong nguyên đai nguyên kiện.",
        "roles_scene": "• Bối cảnh: Quầy lễ tân & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Bác khách chủ xưởng (mặc vest bóng bẩy, đeo kính râm, bụng bia to)\n• Đạo cụ: Kẹp hồ sơ A4, catalogue thiết bị cơ khí, con dấu công ty.",
        "pingpong": "• [HOOK 0-3s]\n• Khách: Tổng tài tới đây!\n• Sales: Em chào Chủ tịch!\n• Khách: Ký cont hỏa tốc!\n• Sếp Thúy: Mời Bác vào bàn!\n• [ESCALATION 4-15s]\n• Khách: Bên cô uy tín?\n• Sếp Thúy: Mười năm pháp nhân!\n• Khách: Xe chạy thế nào?\n• Sales: Chạy liên vận ngày!\n• Khách: Soi hàng ở đâu?\n• Sales: Camera Bằng Tường!\n• Khách: Máy móc an toàn?\n• Sếp Thúy: Niêm phong nguyên đai!\n• Khách: Bụng tôi phong độ?\n• [PUNCHLINE 16-25s]\n• Sales: Bụng bia quá khổ!\n• Khách (đơ 1.5s): Ơ hay! Đừng soi!\n• Khách: Bụng chứa hợp đồng!\n• Khách: Chứ bia bọt gì!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Ký ngay cho thon!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Khách diện vest bảnh bao tự nhận phong thái tổng tài bước vào văn phòng Eureka đòi ký cont hỏa tốc, hỏi kiểm tra nghiệp vụ Eureka cực gắt xong bất ngờ hỏi Sales xem bụng mình có phong độ không, Sales tỉnh bơ chê bụng bia quá khổ khiến khách giật bắn mình chữa thẹn bụng chứa hợp đồng cont, Sếp Thúy giục ký nhanh cho thon bụng.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-72",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z)",
        "title": "Hai nhiệm vụ tối thượng của Sales Eureka và câu trả lời né KPI của Gen Z (Tham khảo Reel 261K views)",
        "strength": "Phòng Xuất nhập khẩu thuê ngoài trọn gói chỉ từ 5 triệu/tháng, tư vấn mã HS và cấp C/O Form E miễn phí, hỗ trợ khách hàng 24/7 trên Portal ERP.",
        "roles_scene": "• Bối cảnh: Góc làm việc máy tính & quầy lễ tân văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + 2 Sales Gen Z\n• Đạo cụ: Smartphone gắn tripod, kẹp tài liệu báo cáo KPI, cốc trà.",
        "pingpong": "• [HOOK 0-3s]\n• Sales 1: Chào mừng đến Eureka!\n• Sales 2: Hai nhiệm vụ lớn!\n• [ESCALATION 4-15s]\n• Sales 1: Một là gom cont!\n• Sales 2: Từ một khối nhé!\n• Sales 1: Hai là phòng XNK!\n• Sales 2: Trọn gói năm triệu!\n• Sales 1: Hóa đơn chuẩn chỉnh!\n• Sales 2: Miễn phí Form E!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Thế KPI đâu?\n• Sales 1 (đơ 1.5s): Dạ... Sếp Thúy ơi!\n• Sếp Thúy: Chưa đủ mười cont?\n• Sales 2: Em xin tăng ca!\n• Sales 1: Xin nợ tháng sau!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Chịu hai đứa luôn!\n• Sếp Thúy: Vào bàn làm việc!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Hai Sales Gen Z quay video giới thiệu văn phòng Eureka cực kỳ hào nhoáng với 2 nhiệm vụ tối thượng (gom cont 1 khối và phòng XNK 5 triệu), Sếp Thúy bất thình lình bước ra hỏi chỉ tiêu KPI 10 cont tháng này đâu thì cả hai nhanh nhảu xin tăng ca và xin khất nợ KPI sang tháng sau khiến Sếp Thúy cạn lời.",
        "duration": "22s"
    },
    {
        "code": "EKA-V2-73",
        "motif": "Motif 5: Chơi chữ hài hước & Khách hỏi dịch vụ VIP",
        "title": "Hỏi cước mười cont máy móc như đại gia và cái kết trải nghiệm dịch vụ (Tham khảo Reel 3.6M views)",
        "strength": "Cấp hạn mức công nợ linh hoạt 10% - 40%, định vị camera kho Bằng Tường 24/7, xuất 100% hóa đơn VAT hợp pháp, xe cont chạy hàng ngày.",
        "roles_scene": "• Bối cảnh: Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Bác khách trung niên (tay lắc chìa khóa xe)\n• Đạo cụ: Tách trà nóng, catalogue máy tiện CNC, hợp đồng vận chuyển A4.",
        "pingpong": "• [HOOK 0-3s]\n• Khách: Cước cont bao nhiêu?\n• Sales: Giá tối ưu nhất!\n• Khách: Mười cont máy tiện?\n• Sếp Thúy: Giảm sâu cho Bác!\n• [ESCALATION 4-15s]\n• Khách: Có công nợ không?\n• Sếp Thúy: Cấp bốn mươi phần!\n• Khách: Camera Bằng Tường?\n• Sales: Soi nét 24/7!\n• Khách: Hóa đơn đỏ chuẩn?\n• Sếp Thúy: VAT chính ngạch!\n• Khách: Xe chạy hàng ngày?\n• Sales: Đúng giờ xuất bến!\n• [PUNCHLINE 16-25s]\n• Sếp Thúy: Bác book mấy cont?\n• Khách (đơ 1.5s): Tôi chưa có xưởng!\n• Sales: Ơ hay! Thật ạ?\n• Khách: Đang chờ trúng số!\n• Khách: Thấy dịch vụ tốt!\n• Khách: Nên vào hỏi thử!\n• [DEFLATION 26-30s]\n• Sếp Thúy: Bao giờ trúng hẵng hỏi!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Khách vào văn phòng hỏi dồn dập các chính sách VIP cho 10 cont máy tiện CNC như đại gia thực thụ (công nợ 40%, camera Bằng Tường, VAT chính ngạch); Sếp Thúy chuẩn bị hợp đồng hỏi book mấy cont thì khách cười trừ bảo chưa có xưởng đang chờ trúng vé số vào hỏi thử xem dịch vụ tốt như lời đồn không.",
        "duration": "24s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 67,
        "title": "Tổng tài bước vào Eureka đàm phán cont và cú bóc mẽ chiếc bụng bia (Reel 269K views)",
        "views": "269,000 views",
        "motif": "Motif 1: Slapstick hành động & Đạo cụ văn phòng",
        "trigger": "Khách tự nhận là tổng tài nghiêng trời lệch đất, thích nhân viên khen ngợi tâng bốc phong thái, nhân viên vừa khen vừa bóc mẽ bụng bia khiến khách giật mình chối bay.",
        "adaptation": "Bác khách diện vest sang trọng bước vào Eureka đàm phán cont hỏa tốc với phong thái Chủ tịch tập đoàn cơ khí, hỏi nghiệp vụ xong hỏi Sales xem bụng mình có phong độ không, Sales bóc mẽ bụng bia quá khổ khiến khách chống chế bụng chứa hợp đồng cont, Sếp Thúy chốt ký nhanh cho thon bụng (EKA-V2-71)."
    },
    {
        "stt": 68,
        "title": "Hai nhiệm vụ tối thượng của Sales Eureka và câu trả lời né KPI của Gen Z (Reel 261K views)",
        "views": "261,000 views",
        "motif": "Motif 2: Drama nội bộ công sở (Sếp Thúy vs Sales Gen Z)",
        "trigger": "Nhân viên làm clip khoe 2 nhiệm vụ cốt lõi của chi nhánh cực kỳ hoành tráng, đến khi quản lý hỏi KPI nếu không đạt thì sao thì trả lời tỉnh bơ xin lỗi/né tránh.",
        "adaptation": "Hai Sales Gen Z quay video giới thiệu 2 dịch vụ đỉnh của Eureka là gom cont 1 khối và phòng XNK 5 triệu, Sếp Thúy bước ra hỏi chỉ tiêu KPI 10 cont tháng này đâu thì hai bạn nhanh nhảu xin tăng ca và xin nợ KPI sang tháng sau (EKA-V2-72)."
    },
    {
        "stt": 69,
        "title": "Hỏi cước mười cont máy móc như đại gia và cái kết trải nghiệm dịch vụ (Reel 3.6M views)",
        "views": "3,600,000 views",
        "motif": "Motif 5: Chơi chữ hài hước & Khách hỏi dịch vụ VIP",
        "trigger": "Khách hỏi dồn dập hàng loạt ưu đãi VIP (lãi suất, thẻ tín dụng, phong thủy, phục vụ 24/7), nhân viên đon đả tư vấn nhiệt tình rồi hỏi gửi bao nhiêu tỷ thì khách tỉnh bơ bảo không có tiền chỉ vào trải nghiệm dịch vụ.",
        "adaptation": "Khách hỏi dồn dập chính sách vận chuyển 10 cont máy tiện CNC như đại gia (công nợ 40%, camera Bằng Tường, hóa đơn VAT), Sếp Thúy hỏi book mấy cont thì khách thú nhận chưa có xưởng đang chờ trúng số ghé vào hỏi thử dịch vụ (EKA-V2-73)."
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
            # Extract utterance after speaker prefix
            if ':' in line_str:
                utterance = line_str.split(':', 1)[1].strip()
            else:
                utterance = line_str.replace('•', '').strip()
            # Clean punctuation for word counting
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
    # Rà soát gác cổng Lớp 3
    validate_pingpong(SCRIPTS_FILE_1)
    
    ok1 = update_file_1()
    ok2 = update_file_2()
    if ok1 and ok2:
        print("\n[✓] CẢ 2 FILE EXCEL TRÊN ONEDRIVE ĐÃ ĐƯỢC CẬP NHẬT HOÀN TẤT VÀ AN TOÀN!")
        run_desktop_sync()
    else:
        print("\n[!] Cảnh báo: Có lỗi trong quá trình lưu file.")
