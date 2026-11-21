# -*- coding: utf-8 -*-
r"""
Cập nhật toàn diện Chuyên mục 1 (CM1_VAN_BAN_NONG) trên Master Series OneDrive:
File: C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx
- 100% thay thế các văn bản cũ (2025/2020) bằng các văn bản, công văn, quyết định và chính sách NÓNG HỔI BAN HÀNH TRONG THÁNG 10/2026 (trong vòng 1 tháng trở lại đây).
- Căn cứ SSoT: Danh_sach_Van_ban_Phap_luat_XNK.xlsx, Tin_Nganh_Ideas_Bank_Eureka.xlsx và Cổng TTĐT Tổng cục Hải quan, Bộ Tài chính, Bộ Công Thương.
- Tuân thủ RULE 10 (Zero Backup - atomic_save).
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
EXCEL_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
TMP_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.tmp.xlsx")

CM1_HOT_ITEMS_OCT_2026 = [
    {
        "code": "CM1-CD10",
        "title": "Công văn 23535/CHQ-GSQL & CV 8744/BCT (30/10/2026): Hải quan & Bộ Công Thương phát cảnh báo đỏ 7 nhóm mặt hàng rủi ro cao bị điều tra phòng vệ thương mại đi Mỹ, EU!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn 23535/CHQ-GSQL ngày 30/10/2026 của Tổng cục Hải quan & Công văn 8744/BCT-PVTM ngày 30/10/2026 của Cục Phòng vệ thương mại (Bộ Công Thương).",
        "hook": "Nếu doanh nghiệp của các bạn đang xuất khẩu pin mặt trời, thép tấm, nhôm định hình hay đồ gỗ sang Mỹ và EU, thì công văn hỏa tốc phát đi ngày 30/10/2026 của Tổng cục Hải quan và Bộ Công Thương chính là 'hồi chuông báo động đỏ' buộc các bạn phải dừng lại kiểm tra lại toàn bộ hồ sơ xuất xứ ngay hôm nay!",
        "points": "1. Điểm mặt 7 nhóm mặt hàng có kim ngạch tăng đột biến bị đưa vào tầm ngắm điều tra lẩn tránh thuế và gian lận xuất xứ.\n2. Cơ chế phối hợp liên ngành: Hải quan soi rọi tỷ lệ giá trị gia tăng (RVC) và kiểm tra cơ sở sản xuất thực tế tại Việt Nam đột xuất không báo trước.\n3. Rủi ro bị hải quan nước nhập khẩu áp thuế trừng phạt lên tới 200 - 400% nếu phát hiện trung chuyển linh kiện nước ngoài đội lốt xuất xứ Việt Nam.\n4. Bộ 4 tài liệu tự vệ sống còn doanh nghiệp phải hoàn thiện để chứng minh năng lực sản xuất thực chất.",
        "visual": "A-roll Thúy cầm bản in Công văn 23535; B-roll screencast danh mục 7 nhóm hàng rủi ro cao, bản đồ luồng hàng xuất sang thị trường Mỹ/EU, hiệu ứng cảnh báo dấu chấm than đỏ.",
        "cliffhanger": "Khi thị trường Mỹ và EU siết chặt điều tra phòng vệ thương mại thì ở ngay các cửa khẩu đường bộ phía Bắc, một công văn hỏa tốc khác vừa được ban hành ngày 29/10/2026 để giải cứu hàng ngàn xe container nông sản. Đó là gì? Thúy sẽ chia sẻ ngay ở video tiếp theo.",
        "conclusion": "Làm xuất khẩu trong thời kỳ bão phòng vệ thương mại, Thúy luôn tự dặn lòng: Xuất xứ hàng hóa không phải là mảnh giấy C/O mua được, mà là bằng chứng sống về năng lực sản xuất thực tế của doanh nghiệp mình.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD11",
        "title": "Công văn hỏa tốc 23007/CHQ-GSQL (29/10/2026): 'Phao cứu sinh' tháo gỡ Thông tư 127/2026/TT-BTC cho phương tiện vận tải đường bộ và container hàng tại cửa khẩu biên giới!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn hỏa tốc số 23007/CHQ-GSQL ngày 29/10/2026 của Tổng cục Hải quan về việc tháo gỡ vướng mắc phát sinh khi triển khai Thông tư 127/2026/TT-BTC.",
        "hook": "Chỉ trong 48 giờ sau khi Thông tư 127/2026/TT-BTC có nguy cơ làm ách tắc hàng ngàn xe container tại các cửa khẩu Lạng Sơn và Lào Cai, Tổng cục Hải quan đã phát đi Công văn hỏa tốc 23007 ngày 29/10/2026. Đây thực sự là chiếc 'phao cứu sinh' mà mọi chủ hàng và nhà xe xuất nhập khẩu đường bộ phải nắm rõ ngay lập tức!",
        "points": "1. Bản chất vướng mắc: Quy định mới về khai báo điện tử trước khi phương tiện đến cửa khẩu làm phát sinh lỗi hệ thống và chậm trễ thông quan.\n2. Quyết định hỏa tốc của Hải quan: Tạm thời cho phép phương tiện vận tải đường bộ và đường sắt áp dụng phương thức khai báo linh hoạt song song để giải phóng hàng nông sản tươi sống.\n3. Danh mục các thông tin bắt buộc và phần được phép bổ sung sau khi xe vào bãi chờ.\n4. Khuyến nghị phối hợp giữa chủ hàng, lái xe và đại lý hải quan tại cửa khẩu Hữu Nghị, Tân Thanh, Kim Thành.",
        "visual": "A-roll bàn làm việc với tài liệu công văn 23007; B-roll hiện trường bãi xe container cửa khẩu Hữu Nghị, sơ đồ đối chiếu quy trình khai báo mới vs cơ chế tháo gỡ tạm thời.",
        "cliffhanger": "Ách tắc cửa khẩu đường bộ vừa được giải tỏa, nhưng ở các cảng biển lớn, một loạt lô hàng nhập khẩu lại đang bị ách tắc chỉ vì một ký hiệu chữ R nhỏ trên bao bì. Câu chuyện này là sao?",
        "conclusion": "Làm logistics cửa khẩu đường bộ, Thúy luôn tự dặn lòng: Chính sách có thể thay đổi bất ngờ trong từng ngày, sự linh hoạt và nhạy bén với công văn hướng dẫn chính là cứu cánh bảo vệ chất lượng từng tấn nông sản của khách hàng.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD12",
        "title": "Công văn 23616/CHQ-GSQL (28/10/2026): Cảnh báo bẫy ký hiệu chữ (®) trên bao bì hàng nhập khẩu — Hải quan siết chặt quyền sở hữu trí tuệ khiến hàng bị ách tắc tại cảng!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn số 23616/CHQ-GSQL ngày 28/10/2026 của Tổng cục Hải quan hướng dẫn xử lý vướng mắc đối với hàng hóa nhập khẩu có chứa ký hiệu bảo hộ nhãn hiệu (®) trên sản phẩm và bao bì.",
        "hook": "Các bạn có tin được không: Một lô hàng nhập khẩu trị giá hàng tỷ đồng đang nằm im lìm tại cảng Hải Phòng suốt 2 tuần qua, không phải vì sai mã HS hay thiếu kiểm tra chuyên ngành, mà chỉ vì trên bao bì xuất hiện một ký hiệu chữ R tròn nhỏ xíu! Công văn 23616 vừa ban hành ngày 28/10/2026 sẽ chỉ rõ chiếc bẫy này!",
        "points": "1. Quy định pháp lý: Ký hiệu chữ ® (Registered) khẳng định nhãn hiệu đã được bảo hộ độc quyền; nếu nhập khẩu mà không có văn bản ủy quyền của chủ sở hữu thì bị nghi vấn xâm phạm quyền sở hữu trí tuệ.\n2. Vì sao cơ quan hải quan siết chặt: Tránh tình trạng hàng giả mạo nhãn hiệu nước ngoài hoặc hàng gia công không phép trà trộn vào thị trường nội địa.\n3. Hướng dẫn tháo gỡ theo Công văn 23616: Bộ tài liệu chứng minh quyền sử dụng nhãn hiệu hợp pháp từ đối tác xuất khẩu.\n4. Checklist 3 điểm kiểm tra nhãn mác bao bì trước khi ký duyệt maket sản phẩm nhập khẩu.",
        "visual": "A-roll Thúy cầm vỏ hộp sản phẩm phóng to ký hiệu chữ ®, slide văn bản Công văn 23616, infographic phân biệt ™ (Trade Mark) vs ® (Registered).",
        "cliffhanger": "Không chỉ nhãn mác bao bì bị soi xét gắt gao, việc nhập khẩu máy móc thiết bị hiện nay cũng đang đối mặt với đợt siết chặt phân loại mã HS chưa từng có theo công văn chỉ đạo ngày 26/10/2026.",
        "conclusion": "Qua câu chuyện ký hiệu chữ R nhỏ, Thúy tự dặn lòng: Trong nghề ngoại thương, sự cẩn trọng không bao giờ là thừa; một chi tiết nhỏ trên bao bì nếu bỏ qua cũng có thể đánh đổi bằng cả chuỗi ngày nằm bãi tốn kém.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD13",
        "title": "Công văn 23405/CHQ-NVTHQ (26/10/2026): Tổng cục Hải quan chỉ đạo siết chặt phân loại mã HS máy móc đồng bộ & bán thành phẩm tháo rời — Hết thời tách tờ khai để né thuế!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn số 23405/CHQ-NVTHQ ngày 26/10/2026 của Tổng cục Hải quan hướng dẫn công tác phân loại hàng hóa và giải quyết bất đồng mã HS xuất nhập khẩu đối với máy móc thiết bị tháo rời.",
        "hook": "Nếu doanh nghiệp của các bạn đang nhập khẩu dây chuyền máy móc công nghiệp hoặc các cụm linh kiện tháo rời, hãy cẩn thận! Công văn 23405 vừa được Tổng cục Hải quan ban hành ngày 26/10/2026 đã chính thức đóng lại cánh cửa của những chiêu trò 'tách đơn, xé lẻ linh kiện' để né mức thuế suất cao!",
        "points": "1. Bản chất tranh chấp: Áp dụng Quy tắc 2a trong 6 Quy tắc phân loại HS — Hàng hóa ở dạng chưa lắp ráp hoặc tháo rời có đặc trưng cơ bản của sản phẩm hoàn chỉnh thì bắt buộc phải áp mã của sản phẩm hoàn chỉnh.\n2. Trường hợp doanh nghiệp cố tình khai tách mã con để hưởng thuế suất 0% thay vì mã tổ hợp 5 - 10% sẽ bị xử lý ra sao?\n3. Thủ tục đăng ký danh mục máy móc thiết bị đồng bộ nhập khẩu nhiều chuyến theo Thông tư 14/2015/TT-BTC & Thông tư 65/2017/TT-BTC.\n4. Rủi ro truy thu thuế và xử phạt hành chính 20% khi cơ quan kiểm tra sau thông quan đối chiếu định mức BOM và dây chuyền nhà máy.",
        "visual": "A-roll Thúy trước màn hình sơ đồ phân tách dây chuyền máy móc, B-roll hình ảnh cụm máy móc tháo rời trong container, highlight Quy tắc 2a phân loại HS.",
        "cliffhanger": "Sau khi phân loại mã HS chuẩn xác, điều kiện để truyền được tờ khai mượt mà là tài khoản người sử dụng VNACCS của doanh nghiệp phải 'sạch'. Thế nhưng hàng loạt tài khoản đang bị quét đóng trong đợt tổng rà soát này...",
        "conclusion": "Làm việc với máy móc tiền tỷ, Thúy luôn tự dặn lòng: Đừng vì vài phần trăm thuế suất chênh lệch trước mắt mà chọn cách khai lách luật; làm chuẩn quy tắc phân loại ngay từ đầu mới là cách bảo vệ doanh nghiệp an toàn dài lâu.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD14",
        "title": "Công văn 2281 & 23474/CHQ-CNTT (26-27/10/2026): Chiến dịch làm sạch tài khoản người sử dụng VNACCS/VCIS — Quét xóa tài khoản khai hải quan không chính chủ!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn số 2281/HQKV19-QLRR&CNTT ngày 27/10/2026 & Công văn 23474/CHQ-CNTT ngày 26/10/2026 của Tổng cục Hải quan về việc đôn đốc triển khai làm sạch dữ liệu doanh nghiệp xuất nhập khẩu trên hệ thống VNACCS/VCIS.",
        "hook": "Sáng nay, có một khách hàng hốt hoảng gọi cho Thúy vì tờ khai hải quan luồng xanh mở ra nhưng hệ thống VNACCS từ chối tiếp nhận với mã lỗi xác thực người dùng. Hóa ra, tài khoản của họ vừa bị khóa trong chiến dịch 'làm sạch dữ liệu người dùng' toàn quốc theo Công văn 2281 và 23474 vừa ban hành cuối tháng 10/2026!",
        "points": "1. Mục tiêu chiến dịch: Quét sạch các tài khoản khai thuê mượn danh, tài khoản do các nhân sự đã nghỉ việc tạo từ nhiều năm trước mà doanh nghiệp không hay biết.\n2. Quy định liên thông: Đối chiếu thông tin người đại diện pháp luật và chữ ký số với Cơ sở dữ liệu quốc gia về dân cư (VNeID/CCCD gắn chip).\n3. Các dấu hiệu tài khoản nằm trong diện bị treo hoặc thu hồi quyền truy cập hệ thống VNACCS/VCIS.\n4. Hướng dẫn 3 bước cập nhật, phân quyền lại tài khoản User ID, Terminal ID trên Cổng Một cửa Quốc gia chỉ mất 15 phút.",
        "visual": "A-roll Thúy thao tác screencast giao diện quản trị tài khoản VNACCS/VCIS, bảng thông báo lỗi từ chối tiếp nhận tài khoản, checklist các bước làm sạch tài khoản doanh nghiệp.",
        "cliffhanger": "Khi tài khoản hệ thống đã thông suốt, một kỳ hạn khắt khe từ thị trường Châu Âu đang đếm ngược từng ngày đối với các nhà xuất khẩu sắt thép, nhôm và phân bón. Đó là hạn chót báo cáo CBAM ngày 31/10...",
        "conclusion": "Quản trị tài khoản hải quan cũng như giữ chùm chìa khóa ngôi nhà của mình, Thúy tự dặn lòng: Định kỳ rà soát, dọn dẹp các tài khoản thừa và cập nhật thông tin chuẩn xác là nguyên tắc sống còn của bảo mật dữ liệu doanh nghiệp.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD15",
        "title": "Hạn chót 31/10 nộp báo cáo phát thải CBAM Quý 3/2026 & Quyết định 1434/QĐ-CHQ: Chạy đua với thời gian để tránh án phạt 50 EUR/tấn của Liên minh Châu Âu!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Quy chế Cơ chế điều chỉnh biên giới carbon (CBAM) của Liên minh Châu Âu & Quyết định số 1434/QĐ-CHQ ngày 30/10/2026 của Tổng cục Hải quan về chuẩn định dạng dữ liệu số hóa NSW.",
        "hook": "Hôm nay đã là ngày 30/10/2026, chỉ còn đúng 24 giờ nữa là đến hạn chót nộp Báo cáo phát thải khí nhà kính CBAM Quý 3/2026 vào hệ thống của Ủy ban Châu Âu. Nếu doanh nghiệp xuất khẩu sắt thép, nhôm hay phân bón của các bạn nộp trễ hoặc nộp sai, mức phạt có thể lên tới 50 Euro cho mỗi tấn phát thải không báo cáo!",
        "points": "1. Nhắc nhở mốc thời gian: Hạn chót 31/10/2026 cho toàn bộ lô hàng xuất khẩu thuộc diện CBAM cập cảng EU trong Quý 3.\n2. Sự thay đổi khắt khe: EU không còn chấp nhận giá trị phát thải mặc định (Default Values) mà bắt buộc phải có số liệu đo đạc thực tế từ dây chuyền nhà máy sản xuất tại Việt Nam.\n3. Những lỗi kỹ thuật phổ biến trên Cổng thông tin CBAM Registry khiến báo cáo bị trả về.\n4. Quyết định 1434/QĐ-CHQ của Hải quan Việt Nam hỗ trợ chuẩn hóa dữ liệu phát thải số hóa ra sao?",
        "visual": "A-roll Thúy trước đồng hồ đếm ngược 24h, B-roll giao diện CBAM Transitional Registry của EU, bảng công thức tính mức phạt 10 - 50 EUR/tấn phát thải.",
        "cliffhanger": "Vượt qua rào cản xanh của Châu Âu là bài toán lớn, nhưng với các lô hàng nhập khẩu bị đối tác giao sai quy cách phải tái xuất trả lại, làm sao để lấy lại được tiền thuế GTGT đã nộp? Công văn ngày 21/10/2026 sẽ giải đáp trọn vẹn...",
        "conclusion": "Đối mặt với tiêu chuẩn xanh toàn cầu, Thúy luôn tự dặn lòng: Xanh hóa chuỗi cung ứng không còn là khẩu hiệu xa vời, mà đã trở thành bài toán chi phí sống còn tính bằng từng con số tài chính mỗi ngày.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD16",
        "title": "Công văn 23244/CHQ-NVTHQ (21/10/2026): Hướng dẫn thủ tục không thu thuế & hoàn thuế GTGT cho hàng nhập khẩu bị lỗi phải tái xuất trả lại nước ngoài!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn số 23244/CHQ-NVTHQ ngày 21/10/2026 của Tổng cục Hải quan hướng dẫn thủ tục không thu thuế, hoàn thuế GTGT đối với hàng hóa nhập khẩu phải tái xuất trả lại chủ hàng nước ngoài.",
        "hook": "Hàng nhập về mở container ra mới phát hiện linh kiện bị gỉ sét, sai thông số kỹ thuật không thể sử dụng. Đối tác nước ngoài đồng ý nhận lại hàng, nhưng tiền thuế nhập khẩu và thuế GTGT hàng trăm triệu đồng doanh nghiệp đã nộp vào ngân sách nhà nước thì lấy lại bằng cách nào? Công văn 23244 vừa ban hành ngày 21/10/2026 đã tháo gỡ dứt điểm nút thắt này!",
        "points": "1. Điều kiện tiên quyết: Hàng hóa chưa qua sử dụng, gia công, chế biến tại Việt Nam và phải có văn bản chấp thuận nhận lại hàng của chủ hàng nước ngoài.\n2. Phân biệt thủ tục: Khi nào làm thủ tục Không thu thuế (chưa nộp) vs Hoàn thuế (đã nộp tiền vào Kho bạc).\n3. Bộ hồ sơ hoàn thuế chuẩn xác theo hướng dẫn của Công văn 23244: Tờ khai nhập khẩu ban đầu, tờ khai tái xuất (loại hình B13), biên bản giám định hiện trường và chứng từ thanh toán hoàn trả tiền.\n4. Thời hạn giải quyết hồ sơ hoàn thuế tại Chi cục Hải quan: 06 ngày làm việc đối với hồ sơ kiểm tra trước hoàn thuế sau.",
        "visual": "A-roll Thúy cầm bộ hồ sơ hoàn thuế mẫu, B-roll hàng hóa bị lỗi trong kho bãi, lưu đồ quy trình hoàn thuế GTGT hàng tái xuất từ bước giám định đến nhận tiền hoàn.",
        "cliffhanger": "Được hoàn thuế là niềm vui lớn, nhưng nếu chứng từ C/O của các bạn chỉ bị một lỗi chính tả nhỏ mà hải quan nghi ngờ và bác bỏ thì sao? Công văn ngày 20/10/2026 của Hải quan sẽ là lá chắn bảo vệ các bạn...",
        "conclusion": "Xử lý hàng lỗi trả lại, Thúy tự dặn lòng: Trong rủi ro luôn có quy trình pháp lý để bảo toàn dòng tiền; chỉ cần nắm vững quy định và chuẩn bị chứng từ minh bạch, doanh nghiệp sẽ không bao giờ chịu thiệt thòi vô lý.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD17",
        "title": "Công văn 23186/CHQ-GSQL (20/10/2026): Hướng dẫn xử lý sai lệch nhỏ (Minor Errors) trên C/O điện tử Một cửa ASEAN — Bí quyết bảo vệ thuế ưu đãi đặc biệt!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn số 23186/CHQ-GSQL ngày 20/10/2026 của Tổng cục Hải quan hướng dẫn kiểm tra tính hợp lệ của Giấy chứng nhận xuất xứ hàng hóa (C/O) điện tử khi có sai lệch nhỏ.",
        "hook": "Một dấu chấm phẩy bị đặt sai vị trí, tên cảng dỡ hàng bị viết tắt hay ngày cấp C/O lệch 1 ngày so với ngày tàu chạy... Đã có bao nhiêu lô hàng của các bạn bị hải quan từ chối cho hưởng thuế nhập khẩu ưu đãi đặc biệt 0% chỉ vì những lỗi li ti như vậy? Hãy đọc ngay Công văn 23186 ban hành ngày 20/10/2026 của Tổng cục Hải quan!",
        "points": "1. Định nghĩa chuẩn mực về 'Sai lệch nhỏ' (Minor Errors) theo các Hiệp định Thương mại Tự do (ATIGA, ACFTA, RCEP).\n2. Những sai lệch được phép chấp nhận: Lỗi chính tả tên công ty không làm thay đổi pháp nhân, khác biệt định dạng ngày tháng, sai sót nhỏ về đơn vị đo lường nếu tổng trọng lượng không đổi.\n3. Hướng dẫn của Tổng cục Hải quan tại Công văn 23186: Công chức hải quan không được tự ý từ chối C/O điện tử nếu bản chất xuất xứ hàng hóa không bị thay đổi.\n4. Mẫu công văn giải trình sai lệch nhỏ gửi Chi cục Hải quan để giải phóng hàng hưởng thuế ưu đãi ngay trong ngày.",
        "visual": "A-roll Thúy soi kính lúp vào một bản C/O Form D điện tử mẫu, B-roll màn hình Cổng Một cửa Quốc gia hiển thị dữ liệu C/O ASEAN, bảng tổng hợp các lỗi nhỏ được phép chấp nhận.",
        "cliffhanger": "Bảo vệ được C/O ưu đãi là một thắng lợi, nhưng có những mặt hàng điện tử mã HS rất bình thường nhưng lại dính vào danh mục An toàn thông tin và Mật mã dân sự. Không biết cách khai mã miễn trừ là bị phạt nặng ngay...",
        "conclusion": "Đấu tranh cho quyền lợi thuế ưu đãi của doanh nghiệp, Thúy luôn tự dặn lòng: Hiểu luật sâu sắc đến từng tiểu tiết chính là sức mạnh để đối thoại đàng hoàng, tự tin và thuyết phục với cơ quan quản lý.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    },
    {
        "code": "CM1-CD18",
        "title": "Công văn 22233 & 19916/CHQ-GSQL (Cập nhật 15/10/2026): Khai báo hàng An toàn thông tin mạng & Mật mã dân sự — Bẫy trùng mã HS bắt buộc phải khai mã 'KO'!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Công văn số 22233/CHQ-GSQL ngày 15/10/2026 & Công văn 19916/CHQ-GSQL của Cục Giám sát Quản lý về việc khai báo hàng hóa an toàn thông tin mạng, mật mã dân sự và hàng hóa lưỡng dụng.",
        "hook": "Các bạn nhập khẩu camera, bộ phát wifi, switch mạng hay máy tính nhúng mã HS 8471, 8517... Tờ khai mở ra bỗng nhiên bị chặn lại bắt nộp Giấy phép An toàn thông tin mạng hoặc Giấy phép Mật mã dân sự! Nếu không có giấy phép, hàng bị phạt và tái xuất. Nhưng có một 'bí mật khai báo' mang tên Mã KO theo Công văn 22233 cập nhật ngày 15/10/2026 mà nhiều người làm nghề vẫn chưa biết!",
        "points": "1. Bản chất chiếc bẫy: Mã HS của hàng hóa thông thường trùng với mã HS nằm trong danh mục quản lý chuyên ngành của Bộ Thông tin & Truyền thông và Ban Cơ yếu Chính phủ, nhưng tính năng kỹ thuật thực tế thì không thuộc diện phải xin phép.\n2. Khai mã 'KO' trên phần ghi chú tờ khai VNACCS: Căn cứ pháp lý theo Công văn 22233/CHQ-GSQL và Công văn 19916/CHQ-GSQL.\n3. Rủi ro khôn lường nếu không khai mã KO: Bị chuyển luồng đỏ, yêu cầu đi giám định kỹ thuật tốn kém hàng triệu đồng và kéo dài thời gian lưu kho bãi.\n4. Bảng checklist đối chiếu tính năng kỹ thuật và hướng dẫn lưu trữ hồ sơ sẵn sàng cho thanh tra sau thông quan 5 năm.",
        "visual": "A-roll Thúy cầm thiết bị switch mạng và bản catalog kỹ thuật, B-roll screencast ô ghi chú tờ khai VNACCS nhập mã KO, infographic so sánh hàng dính giấy phép vs hàng khai KO.",
        "cliffhanger": "Văn bản chính sách xuất nhập khẩu luôn chuyển động không ngừng nghỉ. Để không bỏ lỡ bất kỳ quy định nóng hổi nào mỗi tuần, đừng quên đăng ký kênh và theo dõi các bản tin tiếp theo cùng Thúy nhé!",
        "conclusion": "Làm xuất nhập khẩu công nghệ cao, Thúy luôn tự dặn lòng: Kỹ thuật khai báo tinh thông không chỉ giúp hàng hóa thông quan thần tốc, mà còn tiết kiệm cho doanh nghiệp hàng chục triệu đồng chi phí lưu bãi và giám định oan uổng.",
        "status": "Chờ Chị Thúy duyệt",
        "proposed_week": "W46/2026",
        "deadline_info": "⏳ Hết hạn sau Tuần W48"
    }
]

def update_cm1_oct_2026():
    print("=== BẮT ĐẦU CẬP NHẬT 9 CHỦ ĐỀ VĂN BẢN NÓNG THÁNG 10/2026 CHO CM1 ===")
    
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Không tìm thấy file: {EXCEL_PATH}")
        return

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb["CM1_VAN_BAN_NONG"]
    
    regular_font = Font(name="Arial", size=10, bold=False, color="000000")
    bold_font = Font(name="Arial", size=10, bold=True, color="000000")
    pending_font = Font(name="Arial", size=10, bold=True, color="B9770E")
    
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    pending_fill = PatternFill(start_color="FCF3CF", end_color="FCF3CF", fill_type="solid")
    
    left_align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    thin_border_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    # 3 dòng đầu (Row 5, 6, 7) là các kịch bản đã hoàn thành (CD01, CD08, CD09)
    # Các dòng từ Row 8 trở đi là nơi đặt 9 chủ đề mới của tuần W46 (CD10 -> CD18)
    
    # Xóa các dòng từ Row 8 trở đi nếu có
    if ws.max_row >= 8:
        ws.delete_rows(8, ws.max_row - 7)
        
    print(f"Đã dọn sạch các dòng cũ từ Row 8. Bắt đầu ghi 9 chủ đề văn bản nóng tháng 10/2026...")
    
    for idx, item in enumerate(CM1_HOT_ITEMS_OCT_2026, start=4):
        target_row = 4 + idx # idx=4 -> row 8
        row_fill = zebra_fill if target_row % 2 == 0 else white_fill
        
        row_values = [
            idx,                     # Col 1: STT (4, 5, 6... 12)
            item["code"],            # Col 2: Mã Chủ Đề
            item["title"],           # Col 3: Tiêu Đề Video (SEO)
            item["format"],          # Col 4: Định Dạng & Số Tập
            item["duration"],        # Col 5: Thời Lượng
            item["ssot"],            # Col 6: Nguồn Dữ Liệu Xác Thực
            item["hook"],            # Col 7: Hook Mở Đầu
            item["points"],          # Col 8: Cốt Lõi Kịch Bản / Key Points
            item["visual"],          # Col 9: Chỉ Dẫn Visual / Góc Máy
            item["cliffhanger"],     # Col 10: Móc Câu Cliffhanger
            item["conclusion"],      # Col 11: Kết Luận & Tự Dặn Lòng
            item["status"],          # Col 12: Trạng Thái Xem Xét
            None,                    # Col 13: Link Kịch Bản Chi Tiết
            item["proposed_week"],   # Col 14: Tuần Đề Xuất
            item["deadline_info"]    # Col 15: Hạn Đào Thải (2 Tuần)
        ]
        
        for c_idx, val in enumerate(row_values, start=1):
            cell = ws.cell(row=target_row, column=c_idx, value=val)
            cell.border = thin_border
            
            if c_idx == 12:
                cell.font = pending_font
                cell.fill = pending_fill
                cell.alignment = center_align
            elif c_idx == 15:
                cell.font = pending_font
                cell.fill = row_fill
                cell.alignment = center_align
            elif c_idx in [1, 2, 4, 5, 14]:
                cell.font = bold_font if c_idx in [1, 2, 14] else regular_font
                cell.fill = row_fill
                cell.alignment = center_align
            else:
                cell.font = regular_font
                cell.fill = row_fill
                cell.alignment = left_align
                
        ws.row_dimensions[target_row].height = 140.25
        print(f"   -> Đã ghi Row {target_row}: STT {idx} | {item['code']} | {item['title'][:55]}...")

    # Cập nhật Sheet 1: TONG_QUAN_LICH_PHAT_SONG
    if "TONG_QUAN_LICH_PHAT_SONG" in wb.sheetnames:
        ws_tq = wb["TONG_QUAN_LICH_PHAT_SONG"]
        ws_tq.cell(6, 4, value="12 Chủ đề (3 Đã duyệt + 9 Đề xuất nóng T10/2026)")
        print("\n✅ Đã đồng bộ số liệu Sheet 1.")

    # Lưu nguyên tử trên cùng ổ C: (RULE 10)
    wb.save(TMP_PATH)
    wb.close()
    os.replace(TMP_PATH, EXCEL_PATH)
    print("\n🎉 HOÀN TẤT CẬP NHẬT 9 CHỦ ĐỀ VĂN BẢN NÓNG THÁNG 10/2026 VÀO ONEDRIVE!")

if __name__ == "__main__":
    update_cm1_oct_2026()
