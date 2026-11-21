# -*- coding: utf-8 -*-
"""
Script: update_daily_viral_scripts_0811.py
Ngày thực hiện: 08/11/2026
Mục tiêu:
1. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 49, 50, 51) vào File 1: BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx
2. Cập nhật 3 kịch bản mới đạt chuẩn 100% (STT 50, 51, 52) vào File 2: BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx (Sheet 1 và Sheet 4)
3. Tham khảo các video Reels MBBank Hải Dương (>= 30K views):
   - Reel 4772045489737141 (57,000 views)
   - Reel 1814131096393184 (60,000 views)
   - Reel 1751050975999933 (130,000 views)
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
        "code": "EKA-REEL-49",
        "title": "Khởi nghiệp 30 tuổi kiếm 30 tỷ cont máy móc và cái kết đi được một nửa (Tham khảo Reel 57K views)",
        "mb_orig": "• Khách: Mấy cái đứa này kỹ năng kém thế, anh chỉ cho này: Phải biết phân khúc khách hàng! Đối với những khách hàng có câu chuyện khởi nghiệp, hành trình kiếm được 30 tỷ đầu tiên ở tuổi 30, thì cách giao tiếp nó phải khác!\n• Em: Ôi 30 tỷ hả anh? Bên em đang có chương trình gửi tiết kiệm từ 10 triệu lãi suất 8,2%, mà 30 tỷ được 9% luôn anh ạ!\n• Khách: Đấy, giao tiếp kém lắm! Đối với những người thành công người ta có tiền tài danh vọng, mình phải ngưỡng mộ người ta!\n• Em: Dạ vâng, em rất ngưỡng mộ anh ạ! Anh có thể chia sẻ bí quyết thành công được không ạ?\n• Khách: Thực ra để đạt được như ngày hôm nay là cả câu chuyện dài vất vả, nên là anh mới đi được một nửa thôi!\n• Em: À thế là anh đã kiếm được 15 tỷ rồi á?\n• Khách: Không! Mới đạt được cái tuổi 30 thôi! Còn 30 tỷ chưa có đồng nào! Thôi mở cho anh cái tài khoản 50 nghìn đi!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka": "• [HOOK 0-3s]\n• Khách (chống cằm giảng đạo): Giao tiếp kém thế!\n• Sales: Dạ kém chỗ nào?\n• Khách: Phải biết phân khúc!\n• Khách: Phân khúc ba mươi tỷ!\n\n• [ESCALATION 4-15s]\n• Sales: Bác nhập ba mươi tỷ?\n• Khách: Dây chuyền cơ khí!\n• Khách: Ba mươi tuổi...\n• Khách: Kiếm ba mươi tỷ!\n• Sales: Ôi ngưỡng mộ anh!\n• Sales: Eureka nhận làm!\n• Sales: Phòng XNK trọn gói!\n• Sales: Năm triệu một tháng!\n• Sales: Form E miễn thuế!\n• Sales: Tiết kiệm tiền tỷ!\n• Sales: Anh chia sẻ bí quyết?\n\n• [PUNCHLINE 16-25s]\n• Khách (thở dài đăm chiêu): Vất vả lắm em!\n• Khách: Anh mới đi...\n• Khách: Được một nửa thôi!\n• Sales: Một nửa là sao?\n• Sales: Kiếm được mười lăm tỷ?\n• Khách (tỉnh bơ): Chưa đâu em!\n• Khách: Vừa tròn ba mươi tuổi!\n• Khách: Chưa có đồng nào!\n• Khách: Gom cont một khối đi!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Ba mươi tuổi tròn...\n• Sếp Thúy: Đòi dạy Sếp phân khúc?\n• (Sound: Aiyo mama! Cậu thanh niên ôm sổ gãi đầu, cả phòng đơ mặt 1.5s)",
        "cast": "Sếp Thúy (ngồi bàn giám đốc quan sát) + 1 Sales nữ Eureka + 1 Thanh niên khởi nghiệp (áo vest kẻ sành điệu, phong thái tổng tài giảng đạo)",
        "setting": "Bàn sofa tiếp khách Eureka 44 Triều Khúc, cuốn sổ ghi chú bìa da, tập hồ sơ Phòng XNK 5 triệu",
        "twist": "Thanh niên lên mặt giảng đạo về phân khúc khách hàng 30 tuổi kiếm 30 tỷ máy móc cơ khí, khoe mới đi được một nửa chặng đường. Tưởng kiếm được 15 tỷ, hóa ra mới đạt cái tuổi 30, còn 30 tỷ chưa có một đồng, xin gom cont lẻ 1 khối.",
        "duration": "24s"
    },
    {
        "code": "EKA-REEL-50",
        "title": "Thế chấp siêu xe nhập khẩu không giấy tờ và cú lật xe rùa chở phân (Tham khảo Reel 60K views)",
        "mb_orig": "• Khách: Anh đang thích ngân hàng lắm đấy, một ngày không ra ngân hàng là bứt rứt!\n• Em: Ở nhà thao tác online được mà anh!\n• Khách: Nhưng người ta muốn đến đây được hỗ trợ cơ! Xe nhà anh thì vay được bao nhiêu em nhỉ? Chắc cẩn thận cũng phải được 500 triệu... Xe mới mua đập hộp tháng trước để mở rộng đàn lợn...\n• Em: Ôi xe mới thì vay được nhiều lắm anh ạ! Anh cho em mượn giấy tờ xe để em điền thông tin ạ!\n• Khách: Giấy tờ gì? Mua bán làm gì có giấy tờ! Cứ đưa tiền mặt là xong luôn, anh đi ngoài đường chả ai dám hỏi han gì!\n• Em: Ôi anh quan hệ rộng thế? Xe nhà anh là xe Mercedes hay Maybach thế anh?\n• Khách: Tên văn minh gọi là xe rùa, còn ở quê bà con hay gọi là xe cút kít chở phân lợn!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka": "• [HOOK 0-3s]\n• Khách (quay tít móc khóa): Eureka làm thủ tục?\n• Sales: Dạ thủ tục gì?\n• Khách: Thông quan xe nhập!\n• Sales: Xe cơ giới hả bác?\n\n• [ESCALATION 4-15s]\n• Khách: Xe chuyên dụng công trình!\n• Sales: Eureka nhận tuốt!\n• Sales: Giám định Quyết định 18!\n• Sales: Đủ hóa đơn VAT!\n• Sales: Không lo rủi ro!\n• Khách: Xe mới đập hộp!\n• Khách: Mua tháng trước!\n• Sales: Bác cho xem giấy!\n• Sales: Giấy chứng nhận xuất xưởng!\n\n• [PUNCHLINE 16-25s]\n• Khách (xua tay gạt phắt): Giấy tờ gì đâu!\n• Sales: Dạ không giấy?\n• Khách: Tiền tươi thóc thật!\n• Khách: Chạy khắp đường làng!\n• Khách: Không ai dám hỏi!\n• Sales (mắt sáng rực): Xe tải hạng nặng?\n• Khách: Xe rùa hai bánh!\n• Khách: Xe cút kít chở phân!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Xe rùa chở phân...\n• Sếp Thúy: Đòi thông quan chính ngạch?\n• (Sound: Aiyo mama! Bác khách nhét móc khóa vào túi cười trừ, cả phòng đơ mặt 1.5s)",
        "cast": "Sếp Thúy (ngồi bàn làm việc kiểm tra hợp đồng) + 1 Sales nam Eureka + 1 Bác chủ xưởng (áo sơ mi cài nút lệch, tay cầm chìa khóa xe quay tít, nét mặt ngạo nghễ)",
        "setting": "Bàn tiếp khách văn phòng 44 Triều Khúc, tập hồ sơ giám định máy móc, chiếc chìa khóa xe to bản",
        "twist": "Bác chủ xưởng vào khoe xe chuyên dụng công trình mới đập hộp tháng trước, chạy khắp đường làng không ai dám hỏi, không cần giấy tờ. Sales hào hứng chuẩn bị hồ sơ xe cơ giới hạng nặng, ai ngờ là chiếc xe rùa cút kít chở phân ở quê.",
        "duration": "23s"
    },
    {
        "code": "EKA-REEL-51",
        "title": "Đố mẹo bài toán công nợ 40% và cú chốt kiểm toán thuế của Sếp (Tham khảo Reel 130K views)",
        "mb_orig": "• Khách: Định nói là gửi từ 10 triệu 8.2 đúng không?\n• Em: Dạ vâng, anh cần hỗ trợ gì cứ bảo bọn em ạ!\n• Khách: Gì cũng hỗ trợ được đúng không?\n• Em: Dạ vâng ạ!\n• Khách: Nhìn cái mặt hớn hở thế kia là biết mới lên lương đúng không? Bây giờ anh đố em xem khả năng thông minh của em đến đâu nhé! Giả sử đi. Em mới lên lương 10 triệu, em cho anh vay 5 triệu. Thế là em còn mấy?\n• Em: Dạ còn 5 triệu ạ!\n• Khách: Thế là em dốt rồi! Em còn mới à? Sao em lại cho anh vay 5 triệu được? Không biết tính à? Dốt thế!\n• Em: Ủa, em có cho anh vay đâu ạ?\n• Khách: Thế cái não em dốt thế, nhìn mặt thông minh mà não dốt... Thế mà bảo cần cái gì cũng hỗ trợ được, rồi lấy tiền đâu mà anh đi chơi Halloween!\n• (Sound: Aiyo mama! Cả phòng đứng hình đơ mặt 1.5s)",
        "eureka": "• [HOOK 0-3s]\n• Khách (bấm máy tính cọc cạch): Eureka cấp công nợ?\n• Sales: Dạ bốn mươi phần trăm!\n• Khách: Hỗ trợ tối đa?\n• Sales: Dạ đúng rồi bác!\n\n• [ESCALATION 4-15s]\n• Khách: Đố em câu này!\n• Khách: Xem tính nhanh không!\n• Sales: Bác cứ đố đi!\n• Khách: Lô máy dập nguội!\n• Khách: Giá một trăm triệu!\n• Khách: Eureka cấp nợ bốn mươi!\n• Khách: Bác cọc sáu mươi!\n• Khách: Máy về xưởng bán!\n• Khách: Bác lãi bao nhiêu?\n• Sales: Lãi tiền bán máy!\n• Sales: Trừ tiền cước cọc!\n\n• [PUNCHLINE 16-25s]\n• Khách (đập bàn đắc thắng): Tính toán quá dốt!\n• Sales (ngơ ngác): Cháu tính sai đâu?\n• Khách: Bác bùng bốn mươi!\n• Khách: Trả nợ làm gì?\n• Khách: Lãi thêm bốn mươi!\n• Sales (toát mồ hôi): Ơ bác tính bùng?\n• Khách: Thế mới giàu nhanh!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (gập laptop bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Tờ khai tên bác...\n• Sếp Thúy: Thuế đến gõ cửa...\n• Sếp Thúy: Chạy đằng trời hả bác?\n• (Sound: Aiyo mama! Bác khách rơi máy tính Casio, cả phòng đơ mặt 1.5s)",
        "cast": "Sếp Thúy (ngồi trước laptop làm việc) + 1 Sales nữ Eureka + 1 Bác khách hàng (cầm máy tính bấm số bỏ túi Casio, kính lão trễ mũi, vẻ mặt gian xảo toan tính)",
        "setting": "Góc làm việc văn phòng Eureka 44 Triều Khúc, máy tính Casio bỏ túi, tập hợp đồng ủy thác nhập khẩu A4, bút bi",
        "twist": "Bác khách đố Sales bài toán lợi nhuận khi được Eureka cấp hạn mức công nợ 40%, rồi chê Sales dốt vì bác tính bùng luôn 40% nợ của công ty để lãi to. Vừa dứt lời thì Sếp Thúy nhắc nhẹ: Tờ khai hải quan và hóa đơn VAT đứng tên pháp nhân của bác, bùng nợ thì Cục Thuế đến gõ cửa tận nhà.",
        "duration": "25s"
    }
]

# Dữ liệu 3 kịch bản mới cho File 2 Sheet 1 (BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx)
NEW_SCRIPTS_FILE_2 = [
    {
        "code": "EKA-V2-50",
        "motif": "Motif 5: Chơi chữ nghĩa đen & Khách tự tin thái quá / Khởi nghiệp 30 tỷ",
        "title": "Khởi nghiệp 30 tuổi kiếm 30 tỷ cont máy móc và cái kết đi được một nửa (Tham khảo Reel 57K views)",
        "canvas": "Phòng Xuất nhập khẩu thuê ngoài trọn gói chỉ từ 5 triệu/tháng, tư vấn mã HS và C/O Form E miễn phí, hỗ trợ các chủ xưởng và startup nhập khẩu máy móc thiết bị chính ngạch từ số vốn nhỏ.",
        "cast_setting": "• Bối cảnh: Bàn sofa tiếp khách Eureka 44 Triều Khúc, cuốn sổ ghi chú bìa da, tập hồ sơ Phòng XNK 5 triệu\n• Nhân vật: Sếp Thúy (ngồi bàn giám đốc quan sát), 1 Sales nữ Eureka, 1 Thanh niên khởi nghiệp (áo vest kẻ sành điệu, phong thái tổng tài giảng đạo)",
        "script": "• [HOOK 0-3s]\n• Khách (chống cằm giảng đạo): Giao tiếp kém thế!\n• Sales: Dạ kém chỗ nào?\n• Khách: Phải biết phân khúc!\n• Khách: Phân khúc ba mươi tỷ!\n\n• [ESCALATION 4-15s]\n• Sales: Bác nhập ba mươi tỷ?\n• Khách: Dây chuyền cơ khí!\n• Khách: Ba mươi tuổi...\n• Khách: Kiếm ba mươi tỷ!\n• Sales: Ôi ngưỡng mộ anh!\n• Sales: Eureka nhận làm!\n• Sales: Phòng XNK trọn gói!\n• Sales: Năm triệu một tháng!\n• Sales: Form E miễn thuế!\n• Sales: Tiết kiệm tiền tỷ!\n• Sales: Anh chia sẻ bí quyết?\n\n• [PUNCHLINE 16-25s]\n• Khách (thở dài đăm chiêu): Vất vả lắm em!\n• Khách: Anh mới đi...\n• Khách: Được một nửa thôi!\n• Sales: Một nửa là sao?\n• Sales: Kiếm được mười lăm tỷ?\n• Khách (tỉnh bơ): Chưa đâu em!\n• Khách: Vừa tròn ba mươi tuổi!\n• Khách: Chưa có đồng nào!\n• Khách: Gom cont một khối đi!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Ba mươi tuổi tròn...\n• Sếp Thúy: Đòi dạy Sếp phân khúc?\n• (Sound: Aiyo mama! Cậu thanh niên ôm sổ gãi đầu, cả phòng đơ mặt 1.5s)",
        "punchline": "Thanh niên lên mặt giảng đạo về phân khúc khách hàng 30 tuổi kiếm 30 tỷ máy móc cơ khí, khoe mới đi được một nửa chặng đường. Tưởng kiếm được 15 tỷ, hóa ra mới đạt cái tuổi 30, còn 30 tỷ chưa có một đồng, xin gom cont lẻ 1 khối.",
        "duration": "24s"
    },
    {
        "code": "EKA-V2-51",
        "motif": "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Thế chấp tài sản khủng",
        "title": "Thế chấp siêu xe nhập khẩu không giấy tờ và cú lật xe rùa chở phân (Tham khảo Reel 60K views)",
        "canvas": "Dịch vụ thông quan, kiểm tra chất lượng và làm thủ tục hải quan máy móc cơ giới, xe chuyên dùng theo Quyết định 18, tư vấn pháp lý xuất nhập khẩu chính ngạch an toàn, bảo vệ doanh nghiệp bằng hóa đơn VAT chuẩn.",
        "cast_setting": "• Bối cảnh: Bàn tiếp khách văn phòng 44 Triều Khúc, tập hồ sơ giám định máy móc, chiếc chìa khóa xe to bản\n• Nhân vật: Sếp Thúy (ngồi bàn làm việc kiểm tra hợp đồng), 1 Sales nam Eureka, 1 Bác chủ xưởng (áo sơ mi cài nút lệch, tay cầm chìa khóa xe quay tít, nét mặt ngạo nghễ)",
        "script": "• [HOOK 0-3s]\n• Khách (quay tít móc khóa): Eureka làm thủ tục?\n• Sales: Dạ thủ tục gì?\n• Khách: Thông quan xe nhập!\n• Sales: Xe cơ giới hả bác?\n\n• [ESCALATION 4-15s]\n• Khách: Xe chuyên dụng công trình!\n• Sales: Eureka nhận tuốt!\n• Sales: Giám định Quyết định 18!\n• Sales: Đủ hóa đơn VAT!\n• Sales: Không lo rủi ro!\n• Khách: Xe mới đập hộp!\n• Khách: Mua tháng trước!\n• Sales: Bác cho xem giấy!\n• Sales: Giấy chứng nhận xuất xưởng!\n\n• [PUNCHLINE 16-25s]\n• Khách (xua tay gạt phắt): Giấy tờ gì đâu!\n• Sales: Dạ không giấy?\n• Khách: Tiền tươi thóc thật!\n• Khách: Chạy khắp đường làng!\n• Khách: Không ai dám hỏi!\n• Sales (mắt sáng rực): Xe tải hạng nặng?\n• Khách: Xe rùa hai bánh!\n• Khách: Xe cút kít chở phân!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Xe rùa chở phân...\n• Sếp Thúy: Đòi thông quan chính ngạch?\n• (Sound: Aiyo mama! Bác khách nhét móc khóa vào túi cười trừ, cả phòng đơ mặt 1.5s)",
        "punchline": "Bác chủ xưởng vào khoe xe chuyên dụng công trình mới đập hộp tháng trước, chạy khắp đường làng không ai dám hỏi, không cần giấy tờ. Sales hào hứng chuẩn bị hồ sơ xe cơ giới hạng nặng, ai ngờ là chiếc xe rùa cút kít chở phân ở quê.",
        "duration": "23s"
    },
    {
        "code": "EKA-V2-52",
        "motif": "Motif 2: Drama nội bộ công sở & Bẫy đố mẹo công nợ 40%",
        "title": "Đố mẹo bài toán công nợ 40% và cú chốt kiểm toán thuế của Sếp (Tham khảo Reel 130K views)",
        "canvas": "Chính sách đòn bẩy tài chính: cấp hạn mức công nợ 10% đến 40% cho doanh nghiệp sản xuất, cam kết bảo vệ pháp nhân bằng 100% tờ khai hải quan và hóa đơn VAT chuẩn chỉnh.",
        "cast_setting": "• Bối cảnh: Góc làm việc văn phòng Eureka 44 Triều Khúc, máy tính Casio bỏ túi, tập hợp đồng ủy thác nhập khẩu A4, bút bi\n• Nhân vật: Sếp Thúy (ngồi trước laptop làm việc), 1 Sales nữ Eureka, 1 Bác khách hàng (cầm máy tính bấm số bỏ túi Casio, kính lão trễ mũi, vẻ mặt gian xảo toan tính)",
        "script": "• [HOOK 0-3s]\n• Khách (bấm máy tính cọc cạch): Eureka cấp công nợ?\n• Sales: Dạ bốn mươi phần trăm!\n• Khách: Hỗ trợ tối đa?\n• Sales: Dạ đúng rồi bác!\n\n• [ESCALATION 4-15s]\n• Khách: Đố em câu này!\n• Khách: Xem tính nhanh không!\n• Sales: Bác cứ đố đi!\n• Khách: Lô máy dập nguội!\n• Khách: Giá một trăm triệu!\n• Khách: Eureka cấp nợ bốn mươi!\n• Khách: Bác cọc sáu mươi!\n• Khách: Máy về xưởng bán!\n• Khách: Bác lãi bao nhiêu?\n• Sales: Lãi tiền bán máy!\n• Sales: Trừ tiền cước cọc!\n\n• [PUNCHLINE 16-25s]\n• Khách (đập bàn đắc thắng): Tính toán quá dốt!\n• Sales (ngơ ngác): Cháu tính sai đâu?\n• Khách: Bác bùng bốn mươi!\n• Khách: Trả nợ làm gì?\n• Khách: Lãi thêm bốn mươi!\n• Sales (toát mồ hôi): Ơ bác tính bùng?\n• Khách: Thế mới giàu nhanh!\n\n• [DEFLATION 26-30s]\n• Sếp Thúy (gập laptop bước ra): Ơ hay!\n• Sếp Thúy: Làm ăn thế đấy!\n• Sếp Thúy: Tờ khai tên bác...\n• Sếp Thúy: Thuế đến gõ cửa...\n• Sếp Thúy: Chạy đằng trời hả bác?\n• (Sound: Aiyo mama! Bác khách rơi máy tính Casio, cả phòng đơ mặt 1.5s)",
        "punchline": "Bác khách đố Sales bài toán lợi nhuận khi được Eureka cấp hạn mức công nợ 40%, rồi chê Sales dốt vì bác tính bùng luôn 40% nợ của công ty để lãi to. Vừa dứt lời thì Sếp Thúy nhắc nhẹ: Tờ khai hải quan và hóa đơn VAT đứng tên pháp nhân của bác, bùng nợ thì Cục Thuế đến gõ cửa tận nhà.",
        "duration": "25s"
    }
]

# Gợi ý ý tưởng cho Sheet 4 của File 2 (GOI_Y_Y_TUONG_MB)
NEW_IDEAS_SHEET_4 = [
    (
        "Khởi nghiệp 30 tuổi kiếm 30 tỷ và cái kết mới đi được một nửa tuổi (Reel 57K views)",
        "57,000 views",
        "Motif 5: Chơi chữ nghĩa đen & Khách tự tin thái quá / Khởi nghiệp 30 tỷ",
        "Khách vào giảng đạo phân khúc khách hàng 30 tuổi kiếm 30 tỷ, khoe mới đi được một nửa chặng đường. Nhân viên tưởng đã kiếm 15 tỷ, ai ngờ mới đạt tuổi 30 còn 30 tỷ chưa có đồng nào, xin mở tài khoản 50k.",
        "Thanh niên khởi nghiệp dạy Sếp Thúy phân khúc 30 tỷ dây chuyền máy móc cơ khí, hóa ra mới đạt tuổi 30 chưa có tiền, xin gom cont lẻ 1 khối (EKA-V2-50)"
    ),
    (
        "Thế chấp xe đập hộp không giấy tờ hóa ra xe rùa cút kít (Reel 60K views)",
        "60,000 views",
        "Motif 1: Slapstick hành động & Đạo cụ văn phòng / Thế chấp tài sản khủng",
        "Khách vào đòi vay thế chấp xe sang mới mua đập hộp tháng trước không cần giấy tờ, nhân viên tưởng Maybach hay Mercedes, ai ngờ là xe rùa cút kít chở phân ở quê.",
        "Bác chủ xưởng đòi làm thủ tục thông quan xe chuyên dụng công trình mới đập hộp không giấy tờ, Eureka tưởng xe tải hạng nặng ai ngờ xe rùa cút kít (EKA-V2-51)"
    ),
    (
        "Đố mẹo bài toán lên lương cho vay tiền và cái kết Halloween (Reel 130K views)",
        "130,000 views",
        "Motif 2: Drama nội bộ công sở & Bẫy đố mẹo công nợ 40%",
        "Khách đố nhân viên bài toán lên lương 10 triệu cho khách vay 5 triệu còn mấy, nhân viên bảo còn 5 triệu thì mắng dốt vì ai lại cho khách vay, làm khách không có tiền đi chơi Halloween.",
        "Khách đố Sales bài toán cấp nợ 40% máy dập 100 triệu, chê Sales dốt vì bác tính bùng nợ 40% để lãi to, bị Sếp Thúy nhắc Cục Thuế tìm tận nhà (EKA-V2-52)"
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
