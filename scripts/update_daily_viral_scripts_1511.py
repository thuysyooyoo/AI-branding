# -*- coding: utf-8 -*-
r"""
Script: update_daily_viral_scripts_1511.py
Thực thi BƯỚC 5 theo chuẩn SOP HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md (Ngày 15/11/2026)
1. Cập nhật 3 kịch bản mới (EKA-REEL-67, 68, 69) vào File 1:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới (EKA-V2-68, 69, 70) vào Sheet 1 và 3 gợi ý vào Sheet 4 của File 2:
   C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx
3. Tự động kiểm tra và xóa kịch bản có ghi chú "Đã sản xuất" hoặc "Đã quay" (nếu có).
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
        "code": "EKA-REEL-67",
        "title": "Kéo nhân viên đi đàm phán hợp đồng cont và phản xạ nghề nghiệp của Sales (Tham khảo Reel 184K views)",
        "mb_raw": "• Quản lý: Cái đứa đối thủ này hôm nay biết tay chị! Hai đứa xông lên cho chị!\n• Nhân viên: Chị ơi nó thách thức kìa!\n• Quản lý: Nó thách chị hay thách hai đứa mày? Bọn này chỉ có giỏi nói to thôi! Chết rồi, chị lại quên chưa học qua khóa võ mồm nào! Hai đứa có biết cãi nhau không?\n• Nhân viên: Bọn em toàn đi thi học sinh giỏi, ai biết chửi nhau bao giờ!\n• Quản lý: Lúc đàm phán quan trọng nhất là phải mồm mép, chị lại hay bị vấp! Bây giờ như này: Hai đứa nhặt hết bao nhiêu ngôn từ sắc bén nhất, cứ từ nào quen mồm nhất thì bắn ra cho chị!\n• Nhân viên 1 (bắn liên thanh): Gửi tiết kiệm lãi suất ưu đãi!\n• Nhân viên 2: Năm mươi triệu tám chấm hai! Một tỷ chín chấm một!\n• Quản lý (đơ mặt 1.5s): Ơ hay! Chị bảo bắn từ sắc bén, ai mượn đọc lãi suất ngân hàng?!\n• (Sound: Aiyo mama! Cả quầy cạn lời đơ mặt 1.5s)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sếp Thúy: Hôm nay đàm phán!\n• Sếp Thúy: Gặp khách ép giá!\n• Sales 1: Sếp dặn gì ạ?\n• Sếp Thúy: Cứ câu quen mồm!\n• Sếp Thúy: Bắn liên thanh nhé!\n• Sales 2: Tụi em sẵn sàng!\n• [ESCALATION 4-15s]\n• Khách: Bên này cước đắt!\n• Khách: Anh cần giảm giá!\n• Sales 1: Gom cont một khối!\n• Sales 2: Hóa đơn VAT chuẩn!\n• Sales 1: Miễn phí Form E!\n• Sales 2: Camera Bằng Tường!\n• Sales 1: Xe chạy hàng ngày!\n• Sales 2: Công nợ bốn mươi!\n• [PUNCHLINE 16-25s]\n• Khách (đơ 1.5s): Ơ hay! Khoan đã!\n• Khách: Chưa kịp mặc cả!\n• Khách: Đã chốt xong luôn?!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Đẳng cấp Eureka mà!\n• Sếp Thúy: Ký hợp đồng nhé!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (dẫn đầu ekip đàm phán) + 2 Sales Eureka (cầm catalogue máy móc và tablet Portal ERP) + Khách chủ xưởng (hất hàm đòi ép giá cước cont).",
        "scene": "Phòng họp & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, tập kẹp hồ sơ A4, tablet hiển thị Portal ERP, con dấu công ty.",
        "twist": "Sếp Thúy dặn nhân viên chuẩn bị từ ngữ sắc bén để đấu trí chống đối tác ép giá cước, ai ngờ 2 sales quen phản xạ nghiệp vụ bắn liên thanh một tràng vũ khí Eureka (gom cont 1 khối, camera Bằng Tường, miễn phí Form E) khiến đối tác hoa mắt chóng mặt chưa kịp mặc cả câu nào đã vội vã chốt đơn.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-68",
        "title": "Khách đòi thử thách tầng tư duy và cái kết bái phục chuyên viên Eureka (Tham khảo Reel 67K views)",
        "mb_raw": "• Khách: Phải là người đẳng cấp cùng tầng tư duy mới tư vấn được cho anh! Em có đọc được suy nghĩ trong đầu anh không?\n• Nhân viên: Nhìn mặt anh là em thấy có vẻ khó rồi đấy, nhưng anh cho em cơ hội thì em thử!\n• Khách: Anh người thường không đọc được đâu! Thử đi!\n• Nhân viên: Em viết sẵn ba đáp án ra giấy nhé, rồi anh trả lời sau!\n• Khách: Câu một: Anh định gửi mấy tỷ?\n• Nhân viên (chìa giấy): Một tỷ!\n• Khách: Câu hai: Kỳ hạn bao lâu?\n• Nhân viên (lật giấy): Kỳ hạn một năm!\n• Khách: Câu ba: Trên dưới trái phải anh đi hướng nào?\n• Nhân viên (lật giấy): Đi lên trên!\n• Khách (đơ mặt 1.5s): Ơ hay! Trúng phóc cả ba! Em đỉnh đấy! Lần sau đến bói cho anh xem bao giờ lấy được vợ nhé!\n• (Sound: Aiyo mama! Cả quầy cạn lời)",
        "eureka_trans": "• [HOOK 0-3s]\n• Khách: Phải cùng tư duy!\n• Khách: Mới đàm phán được!\n• Sales: Bác cần nhập gì?\n• Khách: Đoán thử xem nào!\n• [ESCALATION 4-15s]\n• Sếp Thúy (ghi giấy): Ghi xong ba dòng!\n• Sếp Thúy: Mời bác mở bài!\n• Khách: Xưởng nhập máy gì?\n• Sếp Thúy (chìa giấy): Máy phay CNC!\n• Khách: Đi từ kho nào?\n• Sếp Thúy (lật giấy): Kho Bằng Tường!\n• Khách: Hình thức vận chuyển?\n• Sếp Thúy (lật giấy): Chính ngạch chuẩn chỉ!\n• [PUNCHLINE 16-25s]\n• Khách (đơ 1.5s): Ơ hay! Siêu thế?!\n• Khách: Đọc vị chuẩn đét!\n• Khách: Thế bói hộ anh!\n• Khách: Bao giờ hết ế?!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Chốt cont hết ế!\n• (Sound: Aiyo mama! Khách bái phục)",
        "roles": "Sếp Thúy (ngồi bàn làm việc, tay cầm bút ghi giấy kẹp file) + Sales Eureka (đứng hỗ trợ) + Bác khách chủ xưởng cơ khí (áo sơ mi, tự cao, vuốt tóc).",
        "scene": "Bàn làm việc Sếp Thúy & quầy tiếp khách văn phòng Eureka 44 Triều Khúc, tập kẹp file A4 ghi chú đáp án, bút ký, catalogue thiết bị cơ khí.",
        "twist": "Khách chủ xưởng cơ khí tự cao đòi Eureka phải cùng 'tầng tư duy' mới làm việc được, thách đố đoán bài toán xưởng mình; Sếp Thúy điềm tĩnh ghi sẵn giấy đoán trúng phóc từ loại máy phay CNC, kho Bằng Tường đến hình thức đi chính ngạch khiến khách bái phục xin bói luôn ngày xưởng hết ế hàng.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-69",
        "title": "Khách chăm chú nghe tư vấn giám định máy cũ và lý do bất ngờ tá túc sofa (Tham khảo Reel 76K views)",
        "mb_raw": "• Khách: Ngồi đây có chương trình gì hấp dẫn không em?\n• Nhân viên: Dạ bên em có chương trình sinh nhật, gửi từ 10 triệu lãi suất 8.2, kỳ hạn 6 đến 12 tháng, quay thưởng trúng xe hơi iPhone...\n• Khách: Ừ, nghe hấp dẫn đấy! Trình bày tiếp đi!\n• Nhân viên: Dạ anh gửi nhiều thì nhận quà càng khủng ạ! Thế anh định mở sổ bao nhiêu tỷ để em làm luôn?\n• Khách: Sổ gì tầm này! Trong đầu chỉ có lãi suất với tiền nong thế à?\n• Nhân viên: Dạ? Thế từ nãy giờ anh ngồi nghe chăm chú là sao ạ?\n• Khách: Đêm qua anh đi quẩy đến 3 giờ sáng, bị mất xe, về nhà vợ gào thét khóa trái cửa không cho vào! May quá thấy ngân hàng mở cửa ghé vào ngồi ké điều hòa ngủ tí! Mấy đứa đi ra cho anh chợp mắt tí đi!\n• (Sound: Aiyo mama! Cả quầy cạn lời)",
        "eureka_trans": "• [HOOK 0-3s]\n• Sales: Bác uống nước ấm!\n• Khách: Cảm ơn cô bé!\n• Sales: Máy dập khuôn cũ?\n• Khách: Nghe nói khó nhập?\n• [ESCALATION 4-15s]\n• Sales: Dưới mười năm tuổi!\n• Khách: Giám định thế nào?\n• Sales: Chuẩn Quyết định 18!\n• Khách: Hiệu suất máy móc?\n• Sales: Trên tám mươi phần!\n• Sales: Ký giám định nhé?\n• [PUNCHLINE 16-25s]\n• Khách (đơ 1.5s): Ký gì tầm này!\n• Sales: Ủa sao thế bác?\n• Khách: Đêm qua xem bóng!\n• Khách: Bị vợ đuổi ra!\n• Khách: Cho ngủ nhờ sofa!\n• Khách: Hai tiếng thôi nhé!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Sofa có tính cước!\n• Sếp Thúy: Một khối một triệu!\n• Khách: Bác trả luôn nhé!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "roles": "Sếp Thúy (bước từ phòng giám đốc ra) + Sales Eureka (cầm kẹp hồ sơ tư vấn) + Bác khách chủ xưởng (mắt thâm quầng, mặt bơ phờ, tay cầm cốc nước ấm).",
        "scene": "Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc, cốc nước nóng, hồ sơ giám định máy móc thiết bị cũ theo Quyết định 18/2019/QĐ-TTg.",
        "twist": "Khách chủ xưởng ngồi nghe tư vấn thủ tục giám định máy dập khuôn cũ theo Quyết định 18 cực kỳ chăm chú gật gù từng câu, Sales tưởng chốt hợp đồng ngay thì mới ngượng ngùng thú nhận xem bóng đá về muộn bị vợ đuổi xin ngủ ké sofa 2 tiếng, Sếp Thúy liền báo cước ngủ sofa theo mét khối cont LCL.",
        "duration": "24s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (KICH_BAN_VIRAL_REELS)
SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-68",
        "motif": "Motif 2: Drama đối thoại công sở & Đàm phán B2B",
        "title": "Kéo nhân viên đi đàm phán hợp đồng cont và phản xạ nghề nghiệp của Sales (Tham khảo Reel 184K views)",
        "strength": "Gom cont LCL từ 1 khối (1 CBM), 100% hóa đơn VAT chính ngạch, tư vấn cấp C/O Form E miễn phí, định vị camera kho Bằng Tường 24/7 và cấp hạn mức công nợ 10%-40%.",
        "roles_scene": "• Bối cảnh: Phòng họp & bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + 2 Sales Eureka + Khách chủ xưởng (đòi ép giá cước cont)\n• Đạo cụ: Kẹp hồ sơ A4, tablet Portal ERP, catalogue máy móc, con dấu công ty.",
        "pingpong": "• [HOOK 0-3s]\n• Sếp Thúy: Hôm nay đàm phán!\n• Sếp Thúy: Gặp khách ép giá!\n• Sales 1: Sếp dặn gì ạ?\n• Sếp Thúy: Cứ câu quen mồm!\n• Sếp Thúy: Bắn liên thanh nhé!\n• Sales 2: Tụi em sẵn sàng!\n• [ESCALATION 4-15s]\n• Khách: Bên này cước đắt!\n• Khách: Anh cần giảm giá!\n• Sales 1: Gom cont một khối!\n• Sales 2: Hóa đơn VAT chuẩn!\n• Sales 1: Miễn phí Form E!\n• Sales 2: Camera Bằng Tường!\n• Sales 1: Xe chạy hàng ngày!\n• Sales 2: Công nợ bốn mươi!\n• [PUNCHLINE 16-25s]\n• Khách (đơ 1.5s): Ơ hay! Khoan đã!\n• Khách: Chưa kịp mặc cả!\n• Khách: Đã chốt xong luôn?!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Đẳng cấp Eureka mà!\n• Sếp Thúy: Ký hợp đồng nhé!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Sếp Thúy dặn nhân viên chuẩn bị từ ngữ sắc bén để đấu trí chống đối tác ép giá cước, ai ngờ 2 sales quen phản xạ nghiệp vụ bắn liên thanh một tràng vũ khí Eureka (gom cont 1 khối, camera Bằng Tường, miễn phí Form E) khiến đối tác hoa mắt chóng mặt chưa kịp mặc cả câu nào đã vội vã chốt đơn.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-69",
        "motif": "Motif 5: Chơi chữ hài hước & Am hiểu kỹ thuật máy móc",
        "title": "Khách đòi thử thách tầng tư duy và cái kết bái phục chuyên viên Eureka (Tham khảo Reel 67K views)",
        "strength": "Am hiểu sâu sắc >1.000 dòng máy móc, cơ khí, linh kiện, dây chuyền sản xuất; tư vấn mã HS và quy trình vận tải quốc tế chính ngạch chính xác tuyệt đối.",
        "roles_scene": "• Bối cảnh: Bàn làm việc Sếp Thúy & sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Bác khách chủ xưởng cơ khí (tự cao, vuốt tóc)\n• Đạo cụ: Kẹp file A4 ghi chú đáp án, bút ký, catalogue thiết bị cơ khí.",
        "pingpong": "• [HOOK 0-3s]\n• Khách: Phải cùng tư duy!\n• Khách: Mới đàm phán được!\n• Sales: Bác cần nhập gì?\n• Khách: Đoán thử xem nào!\n• [ESCALATION 4-15s]\n• Sếp Thúy (ghi giấy): Ghi xong ba dòng!\n• Sếp Thúy: Mời bác mở bài!\n• Khách: Xưởng nhập máy gì?\n• Sếp Thúy (chìa giấy): Máy phay CNC!\n• Khách: Đi từ kho nào?\n• Sếp Thúy (lật giấy): Kho Bằng Tường!\n• Khách: Hình thức vận chuyển?\n• Sếp Thúy (lật giấy): Chính ngạch chuẩn chỉ!\n• [PUNCHLINE 16-25s]\n• Khách (đơ 1.5s): Ơ hay! Siêu thế?!\n• Khách: Đọc vị chuẩn đét!\n• Khách: Thế bói hộ anh!\n• Khách: Bao giờ hết ế?!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Chốt cont hết ế!\n• (Sound: Aiyo mama! Khách bái phục)",
        "twist": "Khách chủ xưởng cơ khí tự cao đòi Eureka phải cùng 'tầng tư duy' mới làm việc được, thách đố đoán bài toán xưởng mình; Sếp Thúy điềm tĩnh ghi sẵn giấy đoán trúng phóc từ loại máy phay CNC, kho Bằng Tường đến hình thức đi chính ngạch khiến khách bái phục xin bói luôn ngày xưởng hết ế hàng.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-70",
        "motif": "Motif 4: Nỗi sợ gia đình & Giám định máy cũ Quyết định 18",
        "title": "Khách chăm chú nghe tư vấn giám định máy cũ và lý do bất ngờ tá túc sofa (Tham khảo Reel 76K views)",
        "strength": "Tư vấn và thực hiện thủ tục giám định máy móc, thiết bị đã qua sử dụng theo Quyết định 18/2019/QĐ-TTg (tuổi thiết bị <= 10 năm, kiểm tra an toàn và bảo vệ môi trường trọn gói).",
        "roles_scene": "• Bối cảnh: Bàn tiếp khách sofa văn phòng Eureka 44 Triều Khúc\n• Nhân vật: Sếp Thúy + Sales Eureka + Bác khách chủ xưởng (mắt thâm quầng, mặt bơ phờ)\n• Đạo cụ: Cốc nước nóng, hồ sơ giám định máy móc theo Quyết định 18.",
        "pingpong": "• [HOOK 0-3s]\n• Sales: Bác uống nước ấm!\n• Khách: Cảm ơn cô bé!\n• Sales: Máy dập khuôn cũ?\n• Khách: Nghe nói khó nhập?\n• [ESCALATION 4-15s]\n• Sales: Dưới mười năm tuổi!\n• Khách: Giám định thế nào?\n• Sales: Chuẩn Quyết định 18!\n• Khách: Hiệu suất máy móc?\n• Sales: Trên tám mươi phần!\n• Sales: Ký giám định nhé?\n• [PUNCHLINE 16-25s]\n• Khách (đơ 1.5s): Ký gì tầm này!\n• Sales: Ủa sao thế bác?\n• Khách: Đêm qua xem bóng!\n• Khách: Bị vợ đuổi ra!\n• Khách: Cho ngủ nhờ sofa!\n• Khách: Hai tiếng thôi nhé!\n• [DEFLATION 26-30s]\n• Sếp Thúy (tỉnh bơ): Sofa có tính cước!\n• Sếp Thúy: Một khối một triệu!\n• Khách: Bác trả luôn nhé!\n• (Sound: Aiyo mama! Cả phòng cười rộ)",
        "twist": "Khách chủ xưởng ngồi nghe tư vấn thủ tục giám định máy dập khuôn cũ theo Quyết định 18 cực kỳ chăm chú gật gù từng câu, Sales tưởng chốt hợp đồng ngay thì mới ngượng ngùng thú nhận xem bóng đá về muộn bị vợ đuổi xin ngủ ké sofa 2 tiếng, Sếp Thúy liền báo cước ngủ sofa theo mét khối cont LCL.",
        "duration": "24s"
    }
]

# Dữ liệu 3 gợi ý ý tưởng cho File 2 Sheet 4 (GOI_Y_Y_TUONG_MB)
IDEAS_SHEET_4 = [
    {
        "stt": 64,
        "title": "Kéo nhân viên đi đàm phán hợp đồng cont và phản xạ nghề nghiệp của Sales (Reel 184K views)",
        "views": "184,000 views",
        "motif": "Motif 2: Drama đối thoại công sở & Đàm phán B2B",
        "trigger": "Chuẩn bị câu chữ đao to búa lớn đi dằn mặt đối thủ, dặn nhân viên từ nào quen mồm nhất thì bắn ra, nhân viên quen mồm bắn lãi suất ngân hàng 50 triệu 8.2 khiến cả phòng đứng hình.",
        "adaptation": "Sếp Thúy kéo 2 sales đi đàm phán với khách ép giá cước, dặn nhả vũ khí nghiệp vụ sắc bén nhất, 2 bạn quen phản xạ bắn liên thanh combo gom cont LCL 1 khối, camera Bằng Tường và C/O Form E khiến khách hoa mắt chóng mặt chưa kịp mặc cả đã vội vã chốt đơn (EKA-V2-68)."
    },
    {
        "stt": 65,
        "title": "Khách đòi thử thách tầng tư duy và cái kết bái phục chuyên viên Eureka (Reel 67K views)",
        "views": "67,000 views",
        "motif": "Motif 5: Chơi chữ hài hước & Am hiểu kỹ thuật máy móc",
        "trigger": "Khách tự nhận tầng tư duy cao đòi nhân viên đọc suy nghĩ trong đầu, nhân viên ghi sẵn đoán trúng phóc 3 câu hỏi khiến khách bái phục xin bói luôn ngày lấy vợ.",
        "adaptation": "Khách chủ xưởng cơ khí tự cao đòi Eureka cùng 'tầng tư duy' mới làm việc được, Sếp Thúy điềm tĩnh ghi sẵn 3 đáp án đoán trúng từ dòng máy phay CNC, kho Bằng Tường đến đi chính ngạch khiến khách ngỡ ngàng bái phục xin bói luôn ngày xưởng hết ế hàng (EKA-V2-69)."
    },
    {
        "stt": 66,
        "title": "Khách chăm chú nghe tư vấn giám định máy cũ và lý do bất ngờ tá túc sofa (Reel 76K views)",
        "views": "76,000 views",
        "motif": "Motif 4: Nỗi sợ vợ thực tế của chủ xưởng kinh doanh & Quyết định 18",
        "trigger": "Khách ngồi nghe nhân viên tư vấn chương trình gửi tiết kiệm say sưa, nhân viên tưởng gửi tiền khủng thì khách bảo đêm qua đi quẩy về muộn bị vợ khóa cửa vào ngủ nhờ sofa.",
        "adaptation": "Khách chủ xưởng nghe tư vấn thủ tục giám định máy dập cũ theo Quyết định 18 cực kỳ chăm chú, Sales tưởng ký hợp đồng thì khách gãi đầu bảo xem bóng đá về muộn bị vợ đuổi ra khỏi nhà xin ngủ nhờ sofa 2 tiếng, Sếp Thúy liền định giá cước ngủ sofa theo mét khối cont (EKA-V2-70)."
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
