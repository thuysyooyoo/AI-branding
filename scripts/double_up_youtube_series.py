# -*- coding: utf-8 -*-
r"""
Script mở rộng nhân đôi số lượng mã chủ đề (Batch Doubling Protocol) cho Master Series:
File: C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx
- Khởi tạo ban đầu (Batch 1): 9 mã (CD01 - CD09) = 36 chủ đề
- Đợt mở rộng tiếp theo (Batch 2): Bổ sung 9 mã mới (CD10 - CD18) cho cả 4 chuyên mục = 72 chủ đề
- Tuyển chọn từ kho 62 video thực chiến Weekly_Content_Calendar, văn bản pháp luật nóng tháng 10/2026,
  câu hỏi nghiệp vụ Facebook Group và kỳ án lịch sử thương mại thế giới.
- Tuân thủ tuyệt đối RULE 10 (Zero Backup - atomic_save).
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
EXCEL_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
TMP_PATH = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.tmp.xlsx")

# ==============================================================================
# DỮ LIỆU BATCH 2 (CD10 -> CD18) CHO 4 CHUYÊN MỤC
# ==============================================================================

CM1_NEW_ITEMS = [
    {
        "stt": 10,
        "code": "CM1-CD10",
        "title": "Thông tư 121/2025/TT-BTC: Siết chặt nghĩa vụ lưu giữ hồ sơ hải quan 05 năm — 6 điểm thay đổi chí mạng doanh nghiệp phải nằm lòng!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Khoản 5 Điều 1 Thông tư số 121/2025/TT-BTC (sửa đổi Điều 16a Thông tư 38/2015/TT-BTC); Điểm đ Khoản 2 Điều 18 Luật Hải quan số 54/2014/QH13; Dữ liệu Daily Note 21/10/2026.",
        "hook": "Nhiều doanh nghiệp cứ nghĩ tờ khai thông quan luồng xanh xong là an toàn 100%. Nhưng Thông tư 121 vừa siết nghĩa vụ lưu trữ hồ sơ 5 năm: Thiếu một phiếu trừ lùi hay kết quả phân tích mẫu là bị truy thu tiền tỷ!",
        "points": "1. Nguyên tắc 'Lưu gì, ai lưu, lưu bao lâu' theo Thông tư 121.\n2. 6 nhóm chứng từ tử huyệt: Phiếu trừ lùi hạn ngạch, biên bản lấy mẫu phân loại, định mức BOM SXXK, phụ lục hợp đồng, điện chuyển tiền MT103, hồ sơ bán phế liệu.\n3. Quy chuẩn 'Folder điện tử 1-1' tại Eureka giúp doanh nghiệp ngủ ngon trước mọi đợt thanh tra sau thông quan.",
        "visual": "Bàn trà mộc, cuốn Luật Hải quan, tập hồ sơ lưu trữ mẫu dán mã vạch. Screencast trực quan phân tích Điều 16a Thông tư 121/2025.",
        "cliffhanger": "Ở tập tiếp theo, Thúy sẽ giải mã tiếp câu chuyện về công văn hỏa tốc 23007 vừa giải cứu hàng nghìn xe container nông sản tại cửa khẩu phía Bắc.",
        "conclusion": "Với bản thân Thúy, một chứng từ làm cẩu thả hôm nay có thể phải trả giá bằng cả tuần mất ngủ của 5 năm sau. Cẩn trọng với giấy tờ lưu trữ là cách bảo vệ doanh nghiệp bền vững nhất.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 11,
        "code": "CM1-CD11",
        "title": "Thông tư 127/2026/TT-BTC & Công văn hỏa tốc 23007: 'Phao cứu sinh' giải cứu xe container nông sản tại cửa khẩu đường bộ phía Bắc!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Thông tư số 127/2026/TT-BTC ngày 28/08/2026; Công văn hỏa tốc số 23007/CHQ-GSQL ngày 16/10/2026 của Cục Hải quan; Dữ liệu cửa khẩu Hữu Nghị, Tân Thanh.",
        "hook": "Ngày 15/10/2026, Thông tư 127 có hiệu lực siết khai báo điện tử phương tiện XNC qua NSW. Nguy cơ ùn tắc hàng ngàn xe container nông sản tại biên giới phía Bắc đã được tháo gỡ hỏa tốc thế nào?",
        "points": "1. Điểm nghẽn kỹ thuật khi áp dụng khai báo phương tiện số hóa đối với xe tải đường bộ.\n2. Quyết định điều hành linh hoạt tại Công văn 23007/CHQ-GSQL: Phân định rạch ròi hàng không/hàng hải vs đường bộ/đường sắt.\n3. Bài học về độ trễ chính sách: Cơ quan quản lý tạo điều kiện giãn tiến độ, nhưng doanh nghiệp tuyệt đối không được ngủ quên.",
        "visual": "Bản đồ các cửa khẩu Lạng Sơn, Lào Cai, mô hình xe container đầu kéo mini, giao diện Cổng Một cửa Quốc gia NSW.",
        "cliffhanger": "Đón xem tập tiếp theo để nắm rõ chính sách gia hạn giảm 2% thuế GTGT liên tục 18 tháng theo Nghị định 174 vừa ban hành.",
        "conclusion": "Chính sách tạo điều kiện chuyển tiếp là cơ hội vàng để doanh nghiệp chuẩn hóa đội ngũ. Ai đi trước một bước chuẩn bị công nghệ thì người đó nắm chắc phần thắng.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 12,
        "code": "CM1-CD12",
        "title": "Nghị định 174/2025/NĐ-CP kéo dài giảm 2% thuế GTGT: 'Phao cứu sinh' dòng tiền nhưng cẩn thận bẫy áp sai thuế suất ở 3 Phụ lục loại trừ!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Nghị định số 174/2025/NĐ-CP ngày 30/06/2025; Nghị quyết số 204/2025/QH15; Danh mục 3 Phụ lục hàng hóa không được giảm thuế GTGT.",
        "hook": "Giảm 2% thuế VAT kéo dài liên tục 18 tháng là niềm vui lớn cho doanh nghiệp nhập khẩu. Nhưng khai nhầm 1 mã HS thuộc diện loại trừ từ 10% xuống 8% là dính ngay án phạt trốn thuế!",
        "points": "1. Bức tranh dòng tiền hỗ trợ từ chính sách giảm thuế GTGT 8% kéo dài.\n2. Giải mã 3 Phụ lục loại trừ: Kim loại, hóa chất, viễn thông, bất động sản.\n3. Case study thực tế: Nhập khẩu hạt nhựa nguyên sinh vs hạt nhựa tái sinh và rủi ro bị truy thu phạt 20% khi hậu kiểm.",
        "visual": "Biểu thuế điện tử, bảng đối soát 3 Phụ lục loại trừ, bảng tính so sánh dòng tiền thuế GTGT.",
        "cliffhanger": "Tập tới Thúy sẽ bật mí cơ chế hoàn thuế xuất khẩu siêu tốc 15 ngày theo Nghị định 182.",
        "conclusion": "Hưởng ưu đãi chính sách là quyền lợi của doanh nghiệp, nhưng hiểu thấu đáo từng điều khoản loại trừ mới là bản lĩnh giúp chúng ta giữ trọn vẹn đồng lãi.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 13,
        "code": "CM1-CD13",
        "title": "Nghị định 182/2025/NĐ-CP số hóa hoàn thuế SXXK rút ngắn 15 ngày: Doanh nghiệp gia công 'mừng thầm' nhưng coi chừng bẫy định mức BOM!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Nghị định số 182/2025/NĐ-CP sửa đổi Nghị định 134/2016/NĐ-CP; Điều 55 & 60 Thông tư 38/2015/TT-BTC; Dữ liệu Daily Note 22/10/2026.",
        "hook": "Hoàn thuế nguyên phụ liệu sản xuất xuất khẩu rút ngắn từ 40 ngày xuống còn 15 ngày qua dịch vụ công. Thủ tục nhanh hơn gấp đôi, nhưng vì sao các đoàn kiểm tra sau thông quan lại vào cuộc quyết liệt hơn?",
        "points": "1. Quy trình hoàn thuế tự động hóa trên nền tảng số của cơ quan hải quan.\n2. Rủi ro vênh định mức thực tế (BOM) giữa phân xưởng sản xuất và hồ sơ kế toán kho.\n3. Công thức 'Tam giác kiểm soát 3 bên' tại Eureka giúp hồ sơ hoàn thuế bất khả xâm phạm.",
        "visual": "Sơ đồ luồng hoàn thuế 15 ngày, bảng định mức kỹ thuật BOM mẫu, phiếu xuất nhập kho nguyên phụ liệu.",
        "cliffhanger": "Đón xem tập tiếp theo về tiêu chuẩn xanh CBAM và chứng chỉ giảm phát thải carbon khi xuất khẩu sang châu Âu.",
        "conclusion": "Thủ tục hành chính số hóa càng nhanh thì trách nhiệm giải trình hậu kiểm của doanh nghiệp càng cao. Minh bạch kho xưởng từ đầu là chìa khóa để hoàn thuế an tâm.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 14,
        "code": "CM1-CD14",
        "title": "Tiêu chuẩn xanh CBAM & Quy định chống mất rừng EUDR 2026: Rào cản kỹ thuật mới cho hàng hóa xuất khẩu Việt Nam!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Quy định Cơ chế điều chỉnh biên giới carbon (CBAM) EU; Quy định chống suy thoái rừng EUDR; Quyết định số 13/2024/QĐ-TTg của Thủ tướng Chính phủ.",
        "hook": "Từ năm 2026, xuất khẩu sang châu Âu không chỉ cần chất lượng tốt giá rẻ mà còn phải mua chứng chỉ carbon CBAM và định vị GPS vùng trồng chống mất rừng EUDR. Do nghiệp Việt đã sẵn sàng chưa?",
        "points": "1. Bản chất cơ chế thuế carbon CBAM đối với sắt thép, nhôm, xi măng, phân bón.\n2. Quy định EUDR đối với cà phê, cao su, gỗ: Yêu cầu truy xuất nguồn gốc đến từng mét vuông đất.\n3. Lộ trình 3 bước chuẩn hóa hồ sơ kiểm kê khí nhà kính cho doanh nghiệp sản xuất Việt Nam.",
        "visual": "Bản đồ vệ tinh định vị vùng trồng cà phê, chứng chỉ carbon mẫu, sơ đồ tính toán lượng phát thải GHG.",
        "cliffhanger": "Tập tới Thúy sẽ cảnh báo bẫy kiểm dịch thực vật và nguy cơ tiêu hủy sâm tươi tại cửa khẩu sân bay.",
        "conclusion": "Xanh hóa chuỗi cung ứng không còn là khẩu hiệu, mà đã trở thành điều kiện sống còn để hàng hóa Việt Nam giữ vững vị thế trên thị trường quốc tế.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 15,
        "code": "CM1-CD15",
        "title": "Siết chặt kiểm dịch thực vật hàng xách tay: Vì sao sâm tươi, nông sản quý bị tiêu hủy tại sân bay và phạt tiền trăm triệu?",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Luật An toàn thực phẩm 2010; Thông tư số 04/2023/TT-BNNPTNT về Danh mục kiểm dịch thực vật; Nghị định số 128/2020/NĐ-CP và Nghị định 115/2018/NĐ-CP.",
        "hook": "Khách hàng tha thiết 'em giúp chị gửi vài triệu cứu lô sâm tươi Hàn Quốc đang bị giữ ở sân bay Nội Bài'. Vì sao Thúy phải từ chối thẳng thắn và bài học đắt giá về rào cản an toàn sinh học?",
        "points": "1. Bản chất kiểm soát dịch hại: Củ sâm tươi dính rễ và đất mang mầm bệnh nguy hại cho nông nghiệp nội địa.\n2. Rủi ro khi nghe lời xúi 'xách tay cho nhanh': Tiêu hủy toàn bộ lô hàng kèm mức phạt 30 đến 100 triệu đồng.\n3. Lộ trình 3 bước nhập khẩu chính ngạch an toàn: Xin giấy phép kiểm dịch thực vật Cục BVTV, C/O Form VK, kiểm nghiệm ATTP.",
        "visual": "Hộp sâm tươi Hàn Quốc bọc rêu ẩm, giấy chứng nhận kiểm dịch Phytosanitary mẫu, quyết định tiêu hủy tại sân bay.",
        "cliffhanger": "Đón xem tập tiếp theo về bẫy xuất hóa đơn thay thế khác quý và cách xử lý hóa đơn sai sót theo Nghị định 123.",
        "conclusion": "Khi đã chạm vào rào cản an toàn thực phẩm và kiểm dịch quốc gia, sự cả nể hay tiền bạc đều vô nghĩa. Chỉ có tuân thủ pháp lý chính ngạch mới giữ trọn vốn liếng.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 16,
        "code": "CM1-CD16",
        "title": "Nghị định 123/2020/NĐ-CP & Thông tư 78: Bẫy xuất hóa đơn thay thế khác quý và rủi ro bị loại trừ chi phí khi quyết toán thuế!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 19 Nghị định số 123/2020/NĐ-CP; Thông tư số 78/2021/TT-BTC; Điều 47 Luật Quản lý thuế số 38/2019/QH14; Dữ liệu Daily Note 22/10/2026.",
        "hook": "Phát hiện hóa đơn quý trước bị sai đơn giá, kế toán vội vàng xuất hóa đơn thay thế ở quý này nhưng quên nộp mẫu 04/SS. Kết quả là bị cơ quan thuế bác toàn bộ chi phí đầu vào!",
        "points": "1. Ranh giới pháp lý giữa Hóa đơn điều chỉnh vs Hóa đơn thay thế theo Nghị định 123.\n2. Thời điểm kê khai thuế GTGT và thuế TNDN khi hóa đơn thay thế nằm ở kỳ tính thuế khác.\n3. Checklist 3 bước rà soát hóa đơn điện tử định kỳ cuối tháng giúp doanh nghiệp không bị phạt chậm nộp.",
        "visual": "Giao diện hệ thống hóa đơn điện tử Tổng cục Thuế, mẫu 04/SS-HĐĐT, bảng đối chiếu kỳ kê khai bổ sung.",
        "cliffhanger": "Tập tới Thúy sẽ giải mã Hộ chiếu kỹ thuật số sản phẩm DPP mà EU sắp áp dụng từ năm 2026.",
        "conclusion": "Hóa đơn điện tử phản ánh nghiệp vụ kinh tế thực tế. Quản trị chứng từ ngăn nắp, làm đúng quy trình sửa đổi ngay từ đầu là chiếc khiên bảo vệ doanh nghiệp tốt nhất.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 17,
        "code": "CM1-CD17",
        "title": "Hộ chiếu số sản phẩm (Digital Product Passport - DPP) của EU: Thách thức hay cơ hội bứt phá cho hàng xuất khẩu Việt Nam?",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Khung tiêu chuẩn EU Ecodesign for Sustainable Products Regulation (ESPR); Tài liệu hướng dẫn kỹ thuật DPP của Ủy ban Châu Âu (EC); Sách trắng Logistics Xanh.",
        "hook": "Quét mã QR trên một chiếc áo hay cục pin xuất khẩu sang châu Âu là biết tường tận từ nguồn gốc sợi vải, lượng nước tiêu thụ, đến tỷ lệ vật liệu tái chế. Tấm hộ chiếu số DPP sắp làm thay đổi cuộc chơi xuất khẩu thế nào?",
        "points": "1. Hộ chiếu số sản phẩm DPP là gì và lộ trình áp dụng bắt buộc của EU từ 2026-2027.\n2. Tác động trực tiếp đến ngành dệt may, da giày, linh kiện điện tử và pin xe điện Việt Nam.\n3. Doanh nghiệp SME cần số hóa dữ liệu chuỗi cung ứng thế nào để không bị loại khỏi chuỗi giá trị toàn cầu.",
        "visual": "Mô hình đồ họa quét mã QR hộ chiếu số DPP, chuỗi khối dữ liệu truy xuất nguồn gốc, sản phẩm dệt may có gắn thẻ số.",
        "cliffhanger": "Đón xem tập tiếp theo về Lệnh 248 và 249 của Hải quan Trung Quốc đối với nông sản xuất khẩu.",
        "conclusion": "Minh bạch dữ liệu toàn diện không còn là xu hướng xa vời mà đã trở thành tấm vé thông hành bắt buộc để hàng hóa Việt Nam vươn ra biển lớn.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 18,
        "code": "CM1-CD18",
        "title": "Lệnh 248 & 249 của Hải quan Trung Quốc (GACC): Cập nhật quy định gia hạn mã số xuất khẩu nông sản năm 2026!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Lệnh số 248 về đăng ký doanh nghiệp sản xuất thực phẩm nhập khẩu nước ngoài; Lệnh số 249 về biện pháp quản lý an toàn thực phẩm XNK của GACC; Cổng thông tin CIFER.",
        "hook": "Nhiều doanh nghiệp xuất khẩu sầu riêng, chuối, mít sang Trung Quốc chủ quan không gia hạn mã số CIFER trước 6 tháng. Đến khi hàng ra cửa khẩu thì mã số bị khóa, hàng quay đầu chịu lỗ tiền tỷ!",
        "points": "1. Quy trình rà soát và nộp hồ sơ gia hạn mã số đăng ký doanh nghiệp trên hệ thống CIFER.\n2. Kiểm soát mã số vùng trồng và cơ sở đóng gói: Chống mượn mã, trà trộn hàng ngoài vùng quy hoạch.\n3. Bí quyết làm việc với cơ quan chức năng hai nước để thông quan nông sản tươi sống êm ru.",
        "visual": "Giao diện hệ thống CIFER Trung Quốc, giấy chứng nhận mã vùng trồng, hình ảnh đóng gói sầu riêng đạt chuẩn xuất khẩu.",
        "cliffhanger": "Ở các tập chuyên mục sau, chúng ta sẽ bước vào những bài học dịch luật thực chiến chuyên sâu.",
        "conclusion": "Thị trường Trung Quốc ngày nay kiểm soát an toàn thực phẩm khắt khe không kém gì Âu Mỹ. Làm ăn bài bản, giữ gìn uy tín vùng trồng là con đường bền vững duy nhất của nông sản Việt.",
        "status": "Chờ Chị Thúy duyệt"
    }
]

CM2_NEW_ITEMS = [
    {
        "stt": 10,
        "code": "CM2-CD10",
        "title": "Vận đơn Switch B/L trong thương mại 3 bên: 4 bước đổi Bill bảo mật nhà cung cấp Trung Quốc an toàn tuyệt đối!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Cẩm nang nghiệp vụ Switch B/L của Liên đoàn các Hiệp hội Giao nhận Quốc tế (FIATA); Luật Hàng hải Việt Nam 2015; Thực tế điều vận tại Eureka Logistics.",
        "hook": "Làm thương mại 3 bên, nỗi sợ lớn nhất là khách hàng mua phát hiện ra xưởng gốc bên Trung Quốc rồi đi đêm trực tiếp. Làm sao để sử dụng vận đơn Switch B/L đổi Shipper hợp pháp mà không bị hãng tàu từ chối?",
        "points": "1. Bản chất Switch B/L là gì: Thu hồi bộ Bill 1 (Set 1) phát hành bộ Bill 2 (Set 2).\n2. 4 điều kiện tiên quyết để hãng tàu đồng ý Switch: Thanh toán xong tiền cước, nộp đủ 3/3 Bill gốc, không thay đổi cảng dỡ hàng, văn bản bảo lãnh LOI.\n3. Những rủi ro lừa đảo thương mại cần phòng tránh khi dùng Switch B/L.",
        "visual": "Sơ đồ luồng hàng - luồng tiền - luồng chứng từ thương mại 3 bên, mô hình 2 bộ vận đơn Switch B/L đối chiếu.",
        "cliffhanger": "Tập tới Thúy sẽ hướng dẫn kỹ thuật tách Bill B/L cho hàng dính Quyết định 31.",
        "conclusion": "Switch B/L là công cụ đắc lực bảo vệ tài sản thông tin của thương nhân. Nhưng sự cẩn trọng và trung thực với hãng tàu là nguyên tắc sống còn để tránh bị quy kết gian lận chứng từ.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 11,
        "code": "CM2-CD11",
        "title": "Tách Bill (B/L) cho hàng hỗn hợp dính Quyết định 31/2026/QĐ-TTg: Làm sao để 1 món hàng không làm kẹt cứng cả container?",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Quyết định số 31/2026/QĐ-TTg của Thủ tướng Chính phủ; Điều 20 & 50 Thông tư số 38/2015/TT-BTC; Quy trình nghiệp vụ vận tải biển Eureka.",
        "hook": "Một container nhập khẩu 20 mặt hàng, chỉ vì 1 món máy móc cũ dính Quyết định 31 bắt buộc mở tờ khai tại cửa khẩu nhập mà cả 19 mặt hàng còn lại bị giam lỏng không thể chuyển cửa khẩu về ICD! Bí quyết tách Bill cứu hàng thế nào?",
        "points": "1. Rủi ro dùng chung 1 Master B/L cho hàng hỗn hợp khi có mặt hàng dính diện kiểm soát danh mục Quyết định 31.\n2. Kỹ thuật đàm phán với Shipper và Forwarder nước ngoài để tách House B/L ngay tại cảng xếp hàng (POL).\n3. Thủ tục khai báo điều chỉnh và mở 2 tờ khai hải quan độc lập cho từng phần hàng.",
        "visual": "Mô hình container chứa hàng hỗn hợp, sơ đồ tách 1 Master Bill thành 2 House Bill độc lập, văn bản Quyết định 31.",
        "cliffhanger": "Đón xem tập tiếp theo về cách sửa mã HS hàng đưa vào kho ngoại quan C11.",
        "conclusion": "Chủ động nghiên cứu chính sách mặt hàng trước khi đóng cont là cách tiết kiệm hàng chục triệu tiền lưu bãi cho doanh nghiệp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 12,
        "code": "CM2-CD12",
        "title": "Hàng gửi Kho ngoại quan (loại hình C11) phát hiện sai mã HS code: Quy trình 3 bước khai bổ sung AMA an toàn!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 20 & Điều 88 Thông tư số 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); Điều 85 Nghị định số 08/2015/NĐ-CP; Hotline tư vấn Eureka.",
        "hook": "Hàng linh kiện điện tử đã đưa vào kho ngoại quan C11, bỗng phát hiện invoice shipper gửi sai mã HS. Nhiều người nghĩ hàng chưa vào nội địa thì muốn sửa gì cũng được. Thực tế rủi ro bị nghi buôn lậu lớn thế nào?",
        "points": "1. Bản chất pháp lý của kho ngoại quan: Chưa nộp thuế nhưng chịu sự giám sát hải quan tuyệt đối.\n2. Bẫy mã HS mới thuộc danh mục hàng cấm tạm nhập tái xuất hoặc cấm gửi kho ngoại quan theo Nghị định 69.\n3. Quy trình 3 bước khai bổ sung AMA an toàn kèm công văn giải trình của shipper.",
        "visual": "Kho ngoại quan hiện đại, kiện hàng dán seal kiểm tra hải quan, mẫu công văn giải trình khai bổ sung.",
        "cliffhanger": "Tập tới Thúy sẽ mổ xẻ rủi ro hóa đơn đầu vào ghi tên chung chung nhưng đầu ra ghi chi tiết.",
        "conclusion": "Trong kho ngoại quan, trung thực và chủ động báo cáo sai sót kỹ thuật là cách duy nhất để bảo vệ quyền lợi doanh nghiệp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 13,
        "code": "CM2-CD13",
        "title": "Hóa đơn đầu vào ghi tên chung chung 'vật tư phụ tùng' nhưng đầu ra ghi chi tiết linh kiện: Rủi ro thanh tra thuế bác khấu trừ!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 10 Nghị định số 123/2020/NĐ-CP về nội dung hóa đơn; Thông tư số 78/2021/TT-BTC; Dữ liệu đối soát thuế doanh nghiệp sản xuất.",
        "hook": "Mua hàng thì nhận hóa đơn ghi chung chung 'vật tư phụ tùng máy may', nhưng khi bán ra cho khách lại ghi rõ 'ổ chao, suốt chỉ, kim may công nghiệp'. Vì sao đoàn thanh tra thuế coi đây là dấu hiệu mua bán hóa đơn bất hợp pháp?",
        "points": "1. Nguyên tắc tương thích danh mục hàng hóa giữa hóa đơn mua vào - định mức sản xuất - hóa đơn bán ra.\n2. Rủi ro bị loại trừ chi phí được trừ khi tính thuế TNDN và truy thu thuế GTGT khấu trừ.\n3. Cách lập Bảng kê chi tiết đính kèm hóa đơn điện tử chuẩn quy định pháp luật.",
        "visual": "Hai mẫu hóa đơn điện tử so sánh cạnh nhau, bảng kê mã định danh hàng hóa, trát quyết định xử phạt thuế mẫu.",
        "cliffhanger": "Đón xem tập tiếp theo về 5 lỗi sai chí mạng trên C/O Form E bên thứ ba.",
        "conclusion": "Sự tỉ mỉ và đồng nhất về tên gọi hàng hóa trên sổ sách là nền tảng vững chắc nhất cho sức khỏe tài chính doanh nghiệp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 14,
        "code": "CM2-CD14",
        "title": "C/O Form E bên thứ ba (Third-party Invoicing): 5 lỗi sai ngớ ngẩn khiến doanh nghiệp bị bác C/O, truy thu thuế tiền tỷ!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Thông tư số 12/2019/TT-BCT quy định Quy tắc xuất xứ hàng hóa trong ACFTA; Công văn hướng dẫn C/O Form E của Cục Giám sát Quản lý Hải quan.",
        "hook": "Hàng nhập từ Trung Quốc qua một công ty thương mại tại Hong Kong hoặc Singapore. C/O Form E mang về tưởng được giảm thuế từ 15% về 0%, ai ngờ hải quan cửa khẩu bác thẳng thừng chỉ vì tích nhầm 1 ô nhỏ!",
        "points": "1. Bản chất C/O bên thứ ba trong Hiệp định ACFTA: Shipper ở nước A, bên phát hành invoice ở nước B, người nhận hàng ở Việt Nam.\n2. 5 lỗi sai tử huyệt: Tên công ty phát hành hóa đơn không ghi ở Ô số 7, quên không tích Ô số 13 'Third Party Invoicing', số invoice ở Ô 10 không khớp invoice thanh toán.\n3. Quy trình xác minh xuất xứ (Verification) và cách chuẩn bị chứng từ bảo vệ Form E.",
        "visual": "Tờ C/O Form E phóng to từng ô, con dấu giáp lai, sơ đồ đối chiếu invoice 3 bên.",
        "cliffhanger": "Tập sau Thúy sẽ phân biệt rạch ròi giữa hàng phi mậu dịch H11 và hàng mẫu thử nghiệm.",
        "conclusion": "Ưu đãi thuế quan FTA là con dao hai lưỡi: Tiết kiệm được nhiều tiền thì khâu thẩm định tính hợp lệ của C/O càng phải chuẩn chỉ từng dấu chấm dấu phẩy.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 15,
        "code": "CM2-CD15",
        "title": "Phân biệt rạch ròi: 'Hàng phi mậu dịch' (H11) vs 'Hàng mẫu thử nghiệm' vs 'Hàng quà biếu tặng' — Khi nào được miễn thuế?",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 16 Luật Thuế Xuất khẩu, thuế Nhập khẩu số 107/2016/QH13; Nghị định số 134/2016/NĐ-CP (Điều 10, 11, 16); Thông tư 38/2015/TT-BTC.",
        "hook": "Đối tác gửi tặng một kiện hàng linh kiện mẫu, bảo bên kia thanh toán hết rồi. Về đến Việt Nam mở tờ khai H11 tưởng không phải đóng đồng nào, ai ngờ hải quan vẫn bắt đóng thuế và nộp phạt! Vì sao hàng phi mậu dịch vẫn bị tính thuế?",
        "points": "1. Khái niệm 'Phi mậu dịch' không đồng nghĩa với 'Miễn thuế': Bản chất trị giá tính thuế hải quan.\n2. Định mức miễn thuế quà biếu tặng (dưới 2 triệu đồng hoặc trên 2 triệu nộp thuế phần vượt).\n3. Điều kiện để hàng mẫu thử nghiệm được miễn thuế NK theo Điều 16 Luật 107.",
        "visual": "Kiện hàng mẫu dán nhãn Sample Not for Sale, biểu thuế tính thuế hàng biếu tặng, tờ khai H11 mẫu.",
        "cliffhanger": "Tập tới Thúy sẽ bật mí bí quyết lập Báo cáo quyết toán Mẫu 15 không lệch 1 con ốc.",
        "conclusion": "Hiểu đúng bản chất luật thuế giúp doanh nghiệp dự trù chính xác chi phí, không bao giờ bị bất ngờ trước các khoản thuế phát sinh.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 16,
        "code": "CM2-CD16",
        "title": "Báo cáo quyết toán hải quan Mẫu 15/BCQT-NVL: Cách rà soát 'Tam giác kiểm soát 3 bên' để lượng tồn kho khớp 100% tờ khai!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 60 Thông tư số 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); Quyết định số 1966/QĐ-TCHQ về quy trình kiểm tra báo cáo quyết toán.",
        "hook": "90 ngày sau khi kết thúc năm tài chính là thời hạn chót nộp Báo cáo quyết toán Mẫu 15. Nỗi sợ hãi lớn nhất của kế toán xnk là số liệu nguyên phụ liệu nhập khẩu tồn kho trên sổ sách kế toán vênh hàng tấn so với số liệu hải quan!",
        "points": "1. Bản chất Báo cáo quyết toán Mẫu 15: Kiểm tra tính cân đối Nhập - Xuất - Tồn của nguyên liệu gia công và SXXK.\n2. 4 nguyên nhân gây lệch số liệu: Hao hụt tự nhiên không khai báo, phế liệu tiêu hủy chưa làm thủ tục, sai sót mã nguyên liệu giữa ERP và VNACCS.\n3. Quy chuẩn 'Tam giác kiểm soát 3 bên' (Thủ kho - Kế toán - Nhân viên XNK) xử lý dứt điểm trước kỳ chốt sổ.",
        "visual": "Bảng báo cáo quyết toán Mẫu 15, phần mềm ERP quản trị kho xưởng, sơ đồ tam giác kiểm soát 3 bên.",
        "cliffhanger": "Đón xem tập tiếp theo về rủi ro bị cơ quan hải quan bác bỏ trị giá khai báo.",
        "conclusion": "Quản trị định mức và tồn kho chuẩn xác từng ngày là nền tảng để doanh nghiệp sản xuất xuất khẩu thảnh thơi vượt qua mọi kỳ hậu kiểm.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 17,
        "code": "CM2-CD17",
        "title": "Rủi ro 'Bơm giá' (Over-invoicing) hoặc 'Hạ giá' (Under-invoicing): Căn cứ tham vấn trị giá hải quan 6 phương pháp!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Thông tư số 39/2015/TT-BTC và Thông tư số 60/2019/TT-BTC quy định về trị giá hải quan; Cơ sở dữ liệu Danh mục rủi ro về trị giá của Tổng cục Hải quan.",
        "hook": "Khai giá mua thật từ xưởng 10 USD nhưng hải quan bác bỏ và ấn định lên 25 USD, bắt nộp thêm hàng trăm triệu tiền thuế! Cơ quan hải quan dựa vào đâu để nghi vấn và doanh nghiệp cần chuẩn bị bộ hồ sơ tham vấn giá thế nào?",
        "points": "1. Cơ sở dữ liệu giá rủi ro của Hải quan và nguyên tắc kích hoạt tham vấn trị giá.\n2. 6 phương pháp xác định trị giá hải quan theo chuẩn mực WTO: Phương pháp trị giá giao dịch, hàng giống hệt, hàng tương tự...\n3. Bộ chứng từ chứng minh giá mua thật: Email đàm phán giá, hợp đồng ngoại, điện chuyển tiền MT103, chi phí vận chuyển nội địa tại nước xuất khẩu.",
        "visual": "Bảng danh mục giá rủi ro, sơ đồ 6 phương pháp xác định trị giá hải quan, biên bản tham vấn giá mẫu.",
        "cliffhanger": "Tập tới Thúy sẽ giải thích thủ tục hàng tạm nhập tái xuất tham dự hội chợ triển lãm.",
        "conclusion": "Minh bạch chứng từ thanh toán ngân hàng và kiên trì giải trình bằng chứng vật chất là cách duy nhất bảo vệ trị giá giao dịch thực tế của doanh nghiệp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 18,
        "code": "CM2-CD18",
        "title": "Hàng tạm nhập tái xuất (G11/G21) phục vụ hội chợ, triển lãm hoặc máy móc cho mượn: Thủ tục gia hạn và xử lý tiền bảo lãnh!",
        "format": "Video Đơn lẻ (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 53, 54 Luật Hải quan số 54/2014/QH13; Điều 84 Nghị định số 08/2015/NĐ-CP; Thông tư số 38/2015/TT-BTC.",
        "hook": "Tạm nhập một dây chuyền máy móc đắt tiền về để phục vụ triển lãm hoặc dùng thử trong 6 tháng. Hết hạn chưa kịp tái xuất mà không làm đơn xin gia hạn thì hậu quả pháp lý và tiền thuế phạt sẽ khủng khiếp ra sao?",
        "points": "1. Bản chất loại hình tạm nhập tái xuất G11 và tái xuất G21: Cơ chế bảo lãnh thuế của ngân hàng thương mại.\n2. Hồ sơ và thủ tục xin gia hạn thời hạn lưu giữ tại Việt Nam trước khi hết hạn ít nhất 15 ngày.\n3. Cách chuyển đổi mục đích sử dụng sang tiêu dùng nội địa (mã A42) khi khách quyết định mua đứt máy móc.",
        "visual": "Máy móc công nghiệp tại triển lãm, thư bảo lãnh ngân hàng mẫu, tờ khai G11 và G21.",
        "cliffhanger": "Khép lại chuyên mục Dịch luật, chúng ta sẽ bước sang những ký sự lịch sử thương mại thế giới đầy mê hoặc.",
        "conclusion": "Theo dõi sát sao từng mốc thời hạn bảo lãnh là trách nhiệm cốt lõi của người làm logistics để dòng tiền doanh nghiệp luôn được an toàn.",
        "status": "Chờ Chị Thúy duyệt"
    }
]

CM3_NEW_ITEMS = [
    {
        "stt": 10,
        "code": "CM3-CD10",
        "title": "Kỳ Án Kênh Đào Panama (1914 & 2024): Đợt hạn hán thế kỷ hồ Gatun, 150 tàu viễn dương chôn chân và ván cược tỷ đô!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Hồ sơ Cơ quan Quản lý Kênh đào Panama (ACP); Tác phẩm 'The Path Between the Seas' của David McCullough; Báo cáo hạn hán khí hậu El Nino hồ Gatun 2023-2024.",
        "hook": "Để rút ngắn mười hai ngàn cây số vòng qua mũi Horn giông bão, nhân loại đã bỏ ra hai mươi bảy ngàn sinh mạng để xẻ đôi eo đất châu Mỹ. Nhưng năm 2024, một cuộc khủng hoảng kỳ lạ chưa từng có nổ ra: Kênh đào Panama cạn nước!",
        "points": "1. Lịch sử xây dựng bi tráng kênh đào Panama: Từ thất bại phá sản của người Pháp đến sự hoàn thành của người Mỹ năm 1914.\n2. Cấu trúc âu tàu nước ngọt độc nhất vô nhị: Mỗi lượt tàu qua ngốn hai trăm triệu lít nước ngọt từ hồ Gatun đổ ra biển.\n3. Vụ đấu giá quyền ưu tiên qua kênh lên tới 4 triệu USD/chuyến và bài học về tính dễ tổn thương của chuỗi cung ứng hàng hải toàn cầu.",
        "visual": "Hoạt hình người que 2D Doodle: Cô bé Thúy nơ vàng cầm kính lúp soi bản đồ eo đất Panama, âu tàu nâng hạ nước, hàng dài tàu container neo chờ trên biển.",
        "cliffhanger": "Ở tập tiếp theo, chúng ta sẽ đến với sự cố mắc cạn chấn động của siêu tàu Ever Given làm tê liệt kênh đào Suez.",
        "conclusion": "Trước sức mạnh của thiên nhiên và biến đổi khí hậu, không có tuyến đường huyết mạch nào là an toàn vĩnh cửu. Đa dạng hóa lộ trình logistics là bài học sống còn của thương mại hiện đại.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 11,
        "code": "CM3-CD11",
        "title": "Cuộc Khủng Hoảng Kênh Đào Suez 2021: Siêu tàu Ever Given 20.000 TEU mắc cạn làm tê liệt 12% thương mại toàn cầu!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Báo cáo điều tra chính thức của Cơ quan Quản lý Kênh đào Suez (SCA); Tạp chí hàng hải quốc tế Lloyd's List; Phân tích thiệt hại kinh tế của Hãng bảo hiểm Allianz.",
        "hook": "Một cơn gió giật sa mạc và một cú bẻ lái định mệnh đã biến con tàu dài 400 mét thành một chiếc nút chai khổng lồ bịt kín dòng chảy 10 tỷ USD mỗi ngày của kinh tế thế giới suốt 6 ngày đêm!",
        "points": "1. Sự cố sáng ngày 23/3/2021: Siêu tàu container Ever Given đâm sầm mũi vào bờ cát kênh Suez.\n2. Cuộc giải cứu nghẹt thở: Hàng triệu mét khối cát được nạo vét, hạm đội tàu kéo phối hợp cùng lực lượng thủy triều trăng rằm.\n3. Cuộc chiến bồi thường pháp lý gần 1 tỷ USD, khái niệm Tổn thất chung (General Average) và bài học về giới hạn kích thước siêu tàu.",
        "visual": "Doodle 2D: Thúy người que cầm thước đo chiều dài tàu Ever Given so với tháp Eiffel, hình ảnh chú máy xúc mini nạo vét cát hài hước, dòng tàu hàng ùn tắc.",
        "cliffhanger": "Tập tới Thúy sẽ kể cho các bạn nghe về chuyến tàu hỏa chở hàng dài nhất hành tinh từ Nghĩa Ô sang Madrid.",
        "conclusion": "Một mắt xích nhỏ bé bị đứt gãy có thể làm rung chuyển toàn bộ chuỗi cung ứng hành tinh. Quản trị rủi ro và bảo hiểm hàng hóa luôn là chiếc phao cứu sinh quý giá nhất.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 12,
        "code": "CM3-CD12",
        "title": "Tuyến Đường Sắt Xuyên Lục Địa Yiwu - Madrid (13.052 km): Chuyến tàu chở hàng dài nhất hành tinh và cách mạng logistics Á - Âu!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Hồ sơ vận hành Tuyến đường sắt Yixinou Railway (Yiwu - Madrid); Báo cáo Logistics Á - Âu của Ngân hàng Thế giới (World Bank); Dữ liệu Hải quan liên vận quốc tế.",
        "hook": "Đi qua 8 quốc gia, vượt mười ba ngàn cây số, và phải 3 lần nhấc bổng toàn bộ toa hàng để đổi khổ đường ray! Đó là hải trình trên đất liền kỳ vĩ của chuyến tàu hàng từ thủ phủ hàng tiêu dùng Nghĩa Ô đến thủ đô Madrid Tây Ban Nha.",
        "points": "1. Nguồn cảm hứng phục hưng Con đường Tơ lụa thời hiện đại: Tại sao lại chọn đường sắt thay vì đường biển hay hàng không?\n2. Thách thức kỹ thuật khổng lồ: 3 lần đổi khổ ray từ khổ hẹp sang khổ rộng của Nga rồi lại sang khổ tiêu chuẩn châu Âu.\n3. Rút ngắn thời gian vận chuyển từ 40 ngày đường biển xuống còn 16-18 ngày và mở ra cơ hội liên vận đường sắt cho hàng hóa Việt Nam.",
        "visual": "Doodle 2D: Thúy người que cầm cờ chạy dọc bản đồ Á - Âu, hình vẽ hài hước cần cẩu nhấc container đổi bánh xe đường ray, đoàn tàu chở hàng vút qua thảo nguyên tuyết trắng.",
        "cliffhanger": "Đón xem tập tiếp theo về ván cược sinh tử 27.000 USD tại sòng bạc Las Vegas cứu đế chế FedEx.",
        "conclusion": "Sự sáng tạo trong logistics là không giới hạn. Khi con người dám kết nối những rào cản kỹ thuật khác biệt, cả lục địa sẽ biến thành một thị trường giao thương thông suốt.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 13,
        "code": "CM3-CD13",
        "title": "Đế Chế FedEx & Cú Đặt Cược 27.000 USD Tại Las Vegas: Fred Smith đã cứu tập đoàn chuyển phát nhanh lớn nhất thế giới thế nào?",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Tự truyện 'Changing How the World Does Business' của Roger Frock; Hồ sơ lịch sử Tập đoàn FedEx; Phân tích tài chính khởi nghiệp Đại học Yale.",
        "hook": "Tài khoản công ty chỉ còn đúng năm ngàn đô la, không đủ tiền đổ xăng cho đội máy bay vào thứ Hai tuần tới. Thay vì đầu hàng, nhà sáng lập FedEx bay thẳng đến sòng bạc Las Vegas đánh cược ván bài Blackjack sinh tử!",
        "points": "1. Ý tưởng bị giáo sư Đại học Yale chấm điểm C: Mô hình mạng lưới trục bánh xe và nan hoa (Hub-and-Spoke) giao hàng qua đêm.\n2. Khủng hoảng giá dầu 1973 và bờ vực phá sản: Cú thắng cược 27.000 USD tại sòng bài đủ duy trì hoạt động thêm vài ngày để tìm được khoản đầu tư 11 triệu USD.\n3. Bài học về sự kiên định với tầm nhìn đột phá và bản lĩnh vượt qua nghịch cảnh của người thuyền trưởng logistics.",
        "visual": "Doodle 2D: Thúy người que cầm bài Blackjack nơ vàng nháy mắt, mô hình máy bay vận tải màu tím cam, bản đồ mạng lưới Hub trung chuyển Memphis.",
        "cliffhanger": "Tập tới Thúy sẽ đưa các bạn về thương cảng cổ Hội An rực rỡ của tiền nhân Đàng Trong.",
        "conclusion": "Thành công lớn trong logistics không bao giờ đến từ sự dễ dàng. Sự kiên trì bảo vệ một giải pháp đúng đắn cho khách hàng sẽ giúp chúng ta vượt qua những thời khắc đen tối nhất.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 14,
        "code": "CM3-CD14",
        "title": "Thương Cảng Hội An Thế Kỷ 16 - 18: Tiền nhân Đàng Trong đã đón gió mùa, mở phố Nhật - phố Khách và quản trị ngoại thương rực rỡ ra sao?",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Tác phẩm 'Phủ Biên Tạp Lục' của Lê Quý Đôn; Tư liệu nghiên cứu Thương cảng Faifo - Hội An của Viện Sử học Việt Nam; Châu bản triều Nguyễn.",
        "hook": "Nếu Đàng Ngoài có Phố Hiến sầm uất ven sông Hồng, thì xứ Đàng Trong thời các chúa Nguyễn lại sở hữu thương cảng Hội An — nơi từng đón hàng ngàn thương thuyền Châu Ấn Thuyền của Nhật Bản, tàu Bồ Đào Nha và đoàn buôn Trung Hoa tạo nên một kỳ tích giao thương rực rỡ suốt hai thế kỷ!",
        "points": "1. Vị thế địa kinh tế của cửa biển Đại Chiêm: Quy luật gió mùa Đông Bắc và Tây Nam biến Hội An thành trạm trung chuyển quốc tế lý tưởng.\n2. Mô hình quản trị đô thị mở: Chúa Nguyễn cho phép người Nhật, người Hoa lập phố riêng, cử thủ lĩnh tự quản, áp dụng quy chế hải quan thu thuế thông thoáng.\n3. Những mặt hàng xuất khẩu làm rạng danh Đại Việt: Tơ tằm vàng, trầm hương, kỳ nam, yến sào và gốm sứ Chu Đậu vươn ra toàn cầu.",
        "visual": "Doodle 2D: Thúy người que cầm đèn lồng Hội An, tranh vẽ Châu Ấn Thuyền (Red Seal Ship) cập bến sông Hoài, chùa Cầu và các thương gia ngoại quốc.",
        "cliffhanger": "Đón xem tập tiếp theo về siêu cảng robot tự động hóa 100% không bóng đèn tại Rotterdam Hà Lan.",
        "conclusion": "Tư duy mở cửa, tôn trọng sự khác biệt văn hóa và tạo điều kiện tối đa cho thương mại của tiền nhân là di sản vô giá cho tinh thần hội nhập kinh tế của Việt Nam hôm nay.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 15,
        "code": "CM3-CD15",
        "title": "Siêu Cảng Tự Động Hóa Maasvlakte 2 Rotterdam: Nơi robot và xe tự hành AGV vận hành 24/7 không bóng đèn!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Báo cáo thường niên của Cơ quan Cảng Rotterdam (Port of Rotterdam Authority); Tạp chí Cảng biển Thông minh Thế giới; Dữ liệu tự động hóa APM Terminals.",
        "hook": "Một bến cảng container khổng lồ đón những con tàu lớn nhất hành tinh, nhưng trên cầu bến tuyệt nhiên không một bóng người, không cần bật đèn vào ban đêm! Những cỗ máy tự hành AGV và cần cẩu điều khiển từ xa cách 50 km đang định hình tương lai logistics thế nào?",
        "points": "1. Dự án lấn biển vĩ đại Maasvlakte 2: Mở rộng cảng Rotterdam vươn ra biển Bắc với diện tích 2.000 héc-ta.\n2. Công nghệ 'Ghost Port' (Cảng ma): Xe tự hành AGV dẫn đường bằng nam châm ngầm dưới lòng đất và năng lượng điện xanh 100% từ gió biển.\n3. Bài học về chuyển đổi số cảng biển cho Việt Nam: Từ Cái Mép - Thị Vải, Lạch Huyện đến Cần Giờ trong kỷ nguyên tự động hóa.",
        "visual": "Doodle 2D: Thúy người que đội mũ bảo hộ nơ vàng đứng trước xe robot AGV tự chạy, cần cẩu tự động nhấc thùng container trên nền đêm đen lấp lánh mã code số.",
        "cliffhanger": "Tập sau Thúy sẽ kể về cuộc chiến chuẩn hóa kích thước container và phát minh chiếc khóa Twistlock cứu mạng hàng triệu thủy thủ.",
        "conclusion": "Tự động hóa không lấy đi việc làm của con người, mà giải phóng người lao động khỏi những môi trường nặng nhọc nguy hiểm để nâng tầm năng suất lên những đỉnh cao mới.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 16,
        "code": "CM3-CD16",
        "title": "Kỳ Án Chuẩn Hóa Khóa Twistlock & Kích Thước Container (1965 - 1970): Quyết định hiến tặng bằng sáng chế của Malcolm McLean!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Tác phẩm 'The Box: How the Shipping Container Made the World Smaller' của Marc Levinson; Biên bản Ủy ban Kỹ thuật ISO/TC 104; Lịch sử Tổ chức Tiêu chuẩn hóa Quốc tế.",
        "hook": "Nếu mỗi hãng tàu tự đẻ ra một kích thước thùng hàng riêng: hãng làm 35ft, hãng làm 24ft, hãng làm 17ft thì chuỗi cung ứng thế giới sẽ sụp đổ trong hỗn loạn. Làm sao mà chuẩn 20ft và 40ft cùng chiếc khóa góc Twistlock xoay 90 độ lại trở thành ngôn ngữ chung của nhân loại?",
        "points": "1. Cuộc chiến tiêu chuẩn khốc liệt giữa hãng Sea-Land của Malcolm McLean và hãng Matson Navigation.\n2. Phát minh vĩ đại của kỹ sư Keith Tantlinger: Bộ góc đúc container (Corner Casting) và chốt khóa Twistlock xoay 90 độ cố định hàng chục tấn hàng trên sóng dữ.\n3. Quyết định vĩ đại của Malcolm McLean: Từ bỏ độc quyền bằng sáng chế, hiến tặng miễn phí cho tổ chức ISO để chuẩn hóa toàn cầu.",
        "visual": "Doodle 2D: Thúy người que cầm mô hình khóa góc Twistlock nghe tiếng 'Cạch', hình ảnh các thùng container xếp tầng ngay ngắn trên tàu biển.",
        "cliffhanger": "Đón xem tập tiếp theo về huyền thoại Trà Mã Cổ Đạo chở trà đổi ngựa trên dãy Himalaya.",
        "conclusion": "Một doanh nghiệp vĩ đại không phải là kẻ tìm cách độc quyền thị trường cho riêng mình, mà là người sẵn sàng chia sẻ tiêu chuẩn chung để cả ngành công nghiệp cùng tiến bước.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 17,
        "code": "CM3-CD17",
        "title": "Kỳ Án Tuyến Đường Trà Mã Cổ Đạo (Tea Horse Road): Hải trình trên lưng ngựa vượt dãy Himalaya chở trà đổi chiến mã thời cổ đại!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Tư liệu Lịch sử Thương mại Tây Nam Trung Quốc; Khảo cứu của Viện Hàn lâm Khoa học Xã hội Trung Quốc; Tác phẩm 'The Ancient Tea Horse Road' của Forbes & Henley.",
        "hook": "Trước khi có container hay tàu hơi nước, có một tuyến đường vận tải hiểm trở nhất lịch sử nhân loại kéo dài hơn bốn ngàn cây số men theo vách đá tuyết phủ dãy Himalaya: Nơi những đoàn la mã chở bánh trà Phổ Nhĩ từ Vân Nam đổi lấy những con chiến mã kiêu hùng của Tây Tạng!",
        "points": "1. Nguồn gốc của giao thương Trà - Ngựa: Triều đình phong kiến Trung Hoa cần ngựa chiến phòng ngự biên ải, còn người Tây Tạng cần trà để hỗ trợ tiêu hóa trên vùng cao nguyên khắc nghiệt.\n2. Sự ra đời của trà bánh ép chặt: Giải pháp đóng gói logistics cổ xưa giúp trà không bị ẩm mốc và tiện chất lên lưng ngựa vượt qua những con đèo cao 5.000 mét.\n3. Bài học về sự bổ khuyết giá trị thương mại: Khi hai thị trường trao đổi những thứ đối phương khao khát nhất thì hiểm nguy trùng điệp cũng không ngăn được dòng chảy giao thương.",
        "visual": "Doodle 2D: Thúy người que dắt chú ngựa thồ hàng chở bánh trà tròn vượt đỉnh núi tuyết, bản đồ uốn lượn của con đường Trà Mã Cổ Đạo.",
        "cliffhanger": "Tập tới Thúy sẽ đưa các bạn theo chân cuộc đại di cư của ngành dệt may toàn cầu.",
        "conclusion": "Bản chất của logistics từ ngàn xưa đến nay không bao giờ thay đổi: Đó là vượt qua mọi rào cản địa lý để mang lại giá trị trọn vẹn và sự ấm no cho con người.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 18,
        "code": "CM3-CD18",
        "title": "Cuộc Đại Di Cư Chuỗi Cung Ứng Dệt May Toàn Cầu: Từ Lancashire nước Anh đến châu Á — Quy luật dòng chảy logistics!",
        "format": "Video Ký sự Lịch sử (8-10p)",
        "duration": "8 - 10 phút",
        "ssot": "Tác phẩm kinh điển 'The Travels of a T-Shirt in the Global Economy' của Giáo sư Pietra Rivoli; Báo cáo Tổ chức Thương mại Thế giới (WTO); Lịch sử Cách mạng Công nghiệp Anh.",
        "hook": "Một chiếc áo phông bạn đang mặc trên người có thể chứa bông vải trồng tại Texas Mỹ, dệt nhuộm tại Trung Quốc, may ráp tại Việt Nam hoặc Bangladesh, và bán lẻ tại London! Hành trình đại di cư của ngành dệt may qua 3 thế kỷ phản ánh quy luật nào của kinh tế toàn cầu?",
        "points": "1. Khởi nguồn từ các nhà máy hơi nước vùng Lancashire nước Anh thế kỷ 18: Nơi khai sinh sản xuất dệt may công nghiệp.\n2. Làn sóng dịch chuyển sang miền Nam nước Mỹ, sang Nhật Bản, Đài Loan, Hàn Quốc, và bùng nổ tại Trung Quốc rồi Đông Nam Á.\n3. Quy luật dòng chảy chuỗi cung ứng: Ngành thâm dụng lao động luôn di chuyển đến nơi có chi phí tối ưu và cơ sở hạ tầng logistics đáp ứng được tốc độ giao hàng.",
        "visual": "Doodle 2D: Thúy người que cầm chiếc áo thun gắn cờ các quốc gia chu du vòng quanh quả địa cầu, sơ đồ chuỗi cung ứng sợi - vải - may - vận tải biển.",
        "cliffhanger": "Khép lại những trang ký sự lịch sử, chúng ta sẽ quay trở lại với những tâm tình thực chiến đẫm mồ hôi của nghề xuất nhập khẩu hôm nay.",
        "conclusion": "Không có ngành sản xuất nào đứng yên một chỗ. Hiểu được quy luật di cư của chuỗi cung ứng giúp doanh nghiệp logistics luôn đi trước một bước để đón đầu các dòng chảy thương mại.",
        "status": "Chờ Chị Thúy duyệt"
    }
]

CM4_NEW_ITEMS = [
    {
        "stt": 10,
        "code": "CM4-CD10",
        "title": "'Tiền cọc đi liền khúc ruột': Bài học cay đắng khi cọc 30% cho xưởng ngoại và container nhận về toàn gạch vụn!",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Thực tế tư vấn tranh chấp thương mại quốc tế tại Eureka Logistics; Điều khoản thanh toán T/T quốc tế; Dịch vụ kiểm xưởng Giám định SGS/Intertek/CCIC.",
        "hook": "Tin tưởng xưởng chào hàng qua mạng với website lung linh, chuyển thẳng tiền cọc 30% bằng điện T/T gần 50.000 USD. Hai tháng sau mở cửa container tại cảng Cát Lái, chủ hàng ngã ngửa khi bên trong chỉ toàn bao cát và gạch vụn! Bài học xương máu gỡ rối thế nào?",
        "points": "1. Thủ đoạn tinh vi của các xưởng 'ma' xuyên biên giới trên các sàn thương mại điện tử quốc tế.\n2. Sai lầm chí mạng: Chuyển tiền cọc vào tài khoản cá nhân hoặc không thuê đơn vị kiểm định độc lập (Factory Audit) trước khi chuyển tiền.\n3. Quy trình 3 bước 'khóa rủi ro' của Eureka: Thẩm định pháp lý giấy phép kinh doanh xưởng gốc, giám định hàng tại bến bốc (Pre-shipment Inspection), và giữ lại 10-20% thanh toán sau khi có vận đơn gốc.",
        "visual": "Bàn trà ấm cúng, chén trà gốm mộc, tờ điện chuyển tiền MT103, ảnh chụp container hàng bị gian lận mẫu.",
        "cliffhanger": "Tập tới Thúy sẽ tâm sự về nghệ thuật quản trị công nợ khách hàng B2B trong mùa cao điểm.",
        "conclusion": "Trong ngoại thương, lòng tin phải được xây dựng trên nền tảng của quy trình kiểm soát chặt chẽ. Đừng bao giờ tiếc vài trăm đô la chi phí giám định để rồi mất trắng cả cơ nghiệp tích cóp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 11,
        "code": "CM4-CD11",
        "title": "Quản trị công nợ khách hàng B2B mùa bão giá: Làm sao để từ chối cho nợ cước vận chuyển mà không mất lòng đối tác?",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Nhật ký điều hành tài chính Eureka Logistics; Quản trị dòng tiền doanh nghiệp dịch vụ giao nhận B2B; Dữ liệu Daily Note 2026.",
        "hook": "Làm logistics, khó nhất không phải là tìm khách hàng hay đặt booking tàu, mà là lúc đòi nợ! Khách hàng thân thiết nợ tiền cước hàng trăm triệu dây dưa từ tháng này qua tháng khác. Làm sao để thu hồi công nợ dứt khoát mà vẫn giữ được tình nghĩa làm ăn lâu dài?",
        "points": "1. Cạm bẫy 'bán hàng bằng công nợ': Doanh số cao ngất ngưởng nhưng dòng tiền âm, doanh nghiệp đi vay lãi ngân hàng để trả trước tiền cước cho hãng tàu.\n2. Bộ tiêu chuẩn phân loại hạn mức tín dụng khách hàng (Credit Limit) 4 nhóm dựa trên sức khỏe tài chính thực tế.\n3. Nghệ thuật đàm phán thu hồi nợ ấm áp nhưng kiên quyết: Tách bạch rõ ràng giữa tình cảm cá nhân và kỷ cương tài chính hợp đồng.",
        "visual": "Sổ theo dõi công nợ B2B, biểu đồ dòng tiền kinh doanh, tách trà nóng trên bàn làm việc.",
        "cliffhanger": "Đón xem tập tiếp theo về bài học cay đắng khi đứng tên hộ Giấy phép kinh doanh.",
        "conclusion": "Một thương vụ chỉ thực sự hoàn thành khi tiền đã về tài khoản ngân hàng. Giữ cho dòng tiền khỏe mạnh là cách tốt nhất để doanh nghiệp bảo vệ chính mình và phục vụ khách hàng chu đáo nhất.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 12,
        "code": "CM4-CD12",
        "title": "Bẫy đứng tên hộ Giám đốc & Cho mượn CCCD: Bài học cay đắng của sự 'cả nể' gánh nợ thuế hàng trăm triệu trên eTax Mobile!",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Điều 66 Luật Quản lý thuế số 38/2019/QH14; Nghị định số 126/2020/NĐ-CP; Điều 12, 13 Luật Doanh nghiệp 2020; Dữ liệu Daily Note 17/10/2026.",
        "hook": "Anh em bạn bè nhờ: 'Em đứng tên hộ anh cái công ty, anh đứng sau điều hành hết, em chỉ ký giấy tờ thôi không phải chịu trách nhiệm gì đâu'. Hai năm sau ra sân bay đi du lịch thì bị công an cửa khẩu chặn lại vì nợ thuế 500 triệu và công ty đã bỏ trốn khỏi địa chỉ kinh doanh!",
        "points": "1. Sự thật pháp lý trần trụi: Trong luật pháp Việt Nam, Người đại diện theo pháp luật chịu trách nhiệm vô hạn trước pháp luật và nghĩa vụ thuế.\n2. Cơ chế tạm hoãn xuất cảnh theo Điều 66 Luật Quản lý thuế: Không cần biết bạn có trực tiếp điều hành hay không, cứ đứng tên là bị chặn tại barie sân bay.\n3. Lời khuyên chân thành của Thúy: Cách nói lời từ chối khéo léo nhưng dứt khoát để không đánh mất tự do của chính mình.",
        "visual": "Màn hình ứng dụng eTax Mobile báo nợ thuế, thông báo tạm hoãn xuất cảnh của cơ quan quản lý xuất nhập cảnh.",
        "cliffhanger": "Tập tới Thúy sẽ chia sẻ câu chuyện nhân sự nòng cốt nghỉ việc mang theo khách hàng.",
        "conclusion": "Sự cả nể trong pháp lý là mầm mống phá hủy mọi bình an trong cuộc sống. Hãy dũng cảm từ chối những việc trái quy chuẩn để bảo vệ sự thảnh thơi cho chính bản thân mình.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 13,
        "code": "CM4-CD13",
        "title": "Khi nhân sự kinh doanh nòng cốt nghỉ việc mang theo khách hàng: Xây dựng cơ chế 'Bảo vệ tài sản tri thức' tại Eureka!",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Nhật ký vận hành nội bộ Eureka Logistics; Case study tái cấu trúc trải nghiệm khách hàng B2B; Daily Note ngày 19/08/2026.",
        "hook": "Đào tạo một bạn sales từ số 0 thành thạo nghề sau 3 năm, bỗng một ngày bạn ấy nộp đơn nghỉ việc và mang theo toàn bộ danh sách khách hàng ruột sang công ty đối thủ. Là chủ doanh nghiệp, Thúy đã vượt qua cú sốc đó và tái cấu trúc quy trình thế nào?",
        "points": "1. Nỗi đau chung của các doanh nghiệp dịch vụ: Khách hàng chỉ biết đến nhân viên kinh doanh phụ trách trực tiếp mà không gắn kết với thương hiệu công ty.\n2. Xây dựng quy trình 'Duy nhất 1 điểm chạm thương hiệu' nhưng vận hành bằng 'Hệ thống hỗ trợ đa tầng' (CSKH, Hiện trường, Chứng từ, Kế toán).\n3. Bài học về văn hóa doanh nghiệp: Đối đãi chân thành với nhân viên bằng sự tử tế, nhưng quản trị bằng hệ thống dữ liệu CRM minh bạch không phụ thuộc vào bất kỳ cá nhân nào.",
        "visual": "Sơ đồ quản trị trải nghiệm khách hàng đa điểm chạm, giao diện hệ thống CRM, hình ảnh ấm áp đội ngũ nhân sự Eureka.",
        "cliffhanger": "Đón xem tập tiếp theo về bí quyết đàm phán Local Charges và hạn lưu container miễn phí với hãng tàu ngoại.",
        "conclusion": "Nhân sự đến và đi là quy luật tự nhiên của cuộc sống. Điều quan trọng nhất là chúng ta xây dựng được một hệ thống chuẩn chỉ và một văn hóa phụng sự để khách hàng luôn tin yêu thương hiệu.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 14,
        "code": "CM4-CD14",
        "title": "Đàm phán Local Charges và Demurrage/Detention với Hãng tàu ngoại: Bí quyết xin 14 - 21 ngày Free Dem/Det cho khách!",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Cẩm nang đàm phán hợp đồng dịch vụ giao nhận vận tải Eureka; Biểu cước phụ phí địa phương của các hãng tàu Maersk, MSC, COSCO, ONE.",
        "hook": "Hàng về đến cảng nhưng vướng kiểm tra chuyên ngành mất 10 ngày chưa thông quan được. Hãng tàu chỉ cho 3 ngày miễn phí lưu bãi (Free Demurrage), từ ngày thứ tư tính phạt 50 USD/ngày/cont! Làm sao để đàm phán xin trước 14 đến 21 ngày Free Time ngay từ khi đặt booking?",
        "points": "1. Bản chất các loại phụ phí địa phương: Demurrage (lưu cont tại bãi), Detention (lưu vỏ cont tại kho riêng), Storage (lưu bãi cảng).\n2. Thời điểm vàng để đàm phán Free Time: Phải chốt trước khi phát hành Booking Confirmation, không thể xin sau khi tàu đã chạy.\n3. Nghệ thuật sử dụng sản lượng tích lũy của Forwarder để thương lượng quyền lợi tối đa cho khách hàng nhỏ lẻ.",
        "visual": "Hóa đơn phụ phí Demurrage hãng tàu phát hành, bảng so sánh thời gian Free Time, mô hình bãi container xếp dỡ.",
        "cliffhanger": "Tập tới Thúy sẽ tâm sự về bữa cơm thân mật cùng các anh tài xế xe container tại cửa khẩu Tân Thanh.",
        "conclusion": "Người làm logistics giỏi là người biết lường trước những khúc quanh của thủ tục để chủ động xin trước thời gian dự phòng, bảo vệ từng đồng chi phí cho khách hàng.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 15,
        "code": "CM4-CD15",
        "title": "Bữa cơm cùng anh em lái xe cont tại bãi xe cửa khẩu Tân Thanh: Thấu hiểu những con người giữ mạch máu logistics!",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Nhật ký thực địa biên giới phía Bắc của Chị Thúy; Tư liệu hoạt động điều vận xe đường bộ Eureka; Daily Note tháng 10/2026.",
        "hook": "Ngồi ăn bát cơm bụi dã chiến bên vệ đường trong cái gió lạnh buốt của cửa khẩu Tân Thanh cùng những người tài xế mắt trũng sâu vì 3 đêm thức trắng canh xe hàng nông sản, Thúy mới thấu hiểu hết sự hy sinh thầm lặng của những người giữ nhịp thở chuỗi cung ứng.",
        "points": "1. Góc khuất phía sau vô-lăng xe container đường dài: Những đêm đông thức trắng trông máy lạnh cont sầu riêng, áp lực quay đầu xe trả vỏ đúng hạn.\n2. Triết lý quản trị nhân văn tại Eureka: Tôn trọng và chăm sóc đội ngũ điều vận hiện trường như những người thân trong gia đình.\n3. Tại sao sự thấu hiểu hiện trường thực tế giúp người làm văn phòng đưa ra những quyết sách điều hành chính xác và nhân văn nhất.",
        "visual": "Ảnh chụp thực tế bãi xe cửa khẩu Tân Thanh trong sương sớm, bữa cơm ấm cúng bên xe tải, nụ cười của người lái xe.",
        "cliffhanger": "Đón xem tập tiếp theo về cách xử lý khi hàng hóa bị tổn thất, đổ vỡ trên biển.",
        "conclusion": "Kinh doanh không chỉ có những con số doanh thu khô khan. Sự trân trọng và biết ơn những con người lao động trực tiếp là thứ giữ cho trái tim người làm nghề luôn ấm áp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 16,
        "code": "CM4-CD16",
        "title": "Hàng bị vỡ hỏng trong container đường biển: Khi nào bảo hiểm đền, khi nào hãng tàu bồi thường và cách tự vệ cho chủ hàng?",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Quy tắc Hague-Visby về giới hạn trách nhiệm người vận chuyển; Quy tắc bảo hiểm hàng hóa xuất nhập khẩu Viện London ICC (Điều kiện A, B, C); Luật Hàng hải 2015.",
        "hook": "Container gốm sứ nhập khẩu mở ra thấy non nửa số kiện bị đổ vỡ nát vụn do sóng to ngoài biển. Hãng tàu chìa ra điều khoản miễn trừ trách nhiệm thiên tai bão gió, còn bảo hiểm thì từ chối vì bao bì chèn lót không đạt chuẩn! Chủ hàng phải làm gì để bảo vệ quyền lợi?",
        "points": "1. Giới hạn trách nhiệm của hãng tàu theo Quy tắc Hague-Visby: Vì sao tiền bồi thường chỉ tính theo cân nặng tượng trưng nếu không khai báo trị giá đặc biệt.\n2. Ranh giới giữa Điều kiện bảo hiểm loại A (mọi rủi ro) vs loại C (chỉ rủi ro cơ bản) và bẫy 'lỗi chèn lót chằng buộc (Lashing) không đúng quy cách'.\n3. Checklist 4 việc phải làm ngay tại hiện trường: Mời giám định độc lập trong vòng 72 giờ, lập biên bản dỡ hàng hỏng có xác nhận của cảng, và phát thư khiếu nại Notice of Claim.",
        "visual": "Hình ảnh container hàng bị xô lệch, chứng thư giám định tổn thất hàng hải Survey Report, hợp đồng bảo hiểm ICC (A).",
        "cliffhanger": "Tập tới Thúy sẽ chia sẻ lý do vì sao Eureka luôn trân trọng từng kiện hàng lẻ 1 CBM của các bạn trẻ khởi nghiệp.",
        "conclusion": "Cẩn trọng trong khâu chằng buộc lashing và mua đúng loại bảo hiểm là chi phí đầu tư bắt buộc để bảo vệ an toàn cho thành quả lao động của doanh nghiệp.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 17,
        "code": "CM4-CD17",
        "title": "'Nâng niu từng kiện hàng lẻ 1 CBM': Vì sao Eureka không bao giờ từ chối các đơn hàng nhỏ của bạn trẻ khởi nghiệp?",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Hành trình khởi nghiệp từ lô hàng mỹ phẩm lãi 2 triệu đồng của Founder Thúy; Triết lý phụng sự và năng lượng Giáp Mộc bao bọc của Eureka Logistics; Daily Notes.",
        "hook": "Nhiều công ty logistics lớn chỉ thích làm việc với các tập đoàn FDI đi hàng chục container, từ chối thẳng thừng những kiện hàng lẻ 1 khối (1 CBM) vì lãi chẳng bõ công làm. Nhưng tại sao ở Eureka, chúng tôi luôn trân trọng và hướng dẫn từng bước cho các bạn trẻ khởi nghiệp?",
        "points": "1. Ký ức về lô hàng đầu đời của chính Thúy: Tốt nghiệp Ngân hàng rẽ ngang, lô mỹ phẩm lãi đúng 2 triệu đồng nhưng nếm đủ kiếp nạn Cổng một cửa và C/O ủy thác.\n2. Sự thấu hiểu nỗi bơ vơ của người mới bắt đầu: Cần một người bạn đồng hành hướng dẫn từng chứng từ, từng mã HS mà không bị phân biệt đối xử.\n3. Triết lý Giáp Mộc phụng sự: Cây đại thụ vững chãi bắt đầu từ hạt mầm nhỏ bé. Đồng hành cùng khách hàng từ những bước đi chập chững là niềm hạnh phúc lớn nhất của người làm nghề.",
        "visual": "Kiện hàng lẻ LCL dán nhãn cẩn thận tại kho, Thúy ngồi hướng dẫn khách hàng trẻ kiểm tra chứng từ, nụ cười chân thành.",
        "cliffhanger": "Ở tập cuối cùng của đợt này, Thúy xin đúc kết 3 phẩm chất cốt tử để đi đường dài với nghề xuất nhập khẩu.",
        "conclusion": "Cho đi sự chân thành và nâng đỡ người khác từ những điều nhỏ bé nhất là cách Eureka gieo mầm cho những mối quan hệ bền chặt vượt thời gian.",
        "status": "Chờ Chị Thúy duyệt"
    },
    {
        "stt": 18,
        "code": "CM4-CD18",
        "title": "Đúc kết 15 năm làm nghề: 3 phẩm chất cốt tử để một bạn trẻ đi đường dài và thành công trong ngành Logistics đầy sóng gió!",
        "format": "Video Tâm sự Nghề (7-9p)",
        "duration": "7 - 9 phút",
        "ssot": "Hành trình 15 năm thực chiến của Nguyễn Thị Ngọc Thúy (Founder & CEO Eureka Logistics); Triết lý sống kết hợp Giáp Mộc vững chãi & Liêm Trinh chính trực; Daily Notes.",
        "hook": "Ngành xuất nhập khẩu và logistics hào nhoáng với những con tàu triệu đô, nhưng cũng đầy cạm bẫy với những đêm mất ngủ, hàng kẹt cảng và những khoản phạt tiền tỷ. Sau 15 năm lăn lộn từ nhân viên hiện trường đến điều hành doanh nghiệp, đâu là 3 phẩm chất quan trọng nhất giúp chúng ta đứng vững?",
        "points": "1. Phẩm chất thứ nhất: Sự Cẩn trọng và Tôn trọng sự thật (Tâm thế Liêm Trinh minh bạch — Một chứng từ cẩu thả hôm nay là mầm mống tai họa 5 năm sau).\n2. Phẩm chất thứ hai: Tinh thần Học hỏi và Thích nghi không ngừng (Chính sách hải quan, thuế, công nghệ số hóa thay đổi từng ngày, ai dừng lại là bị đào thải ngay).\n3. Phẩm chất thứ ba: Lòng Trắc ẩn và Tinh thần Phụng sự bền bỉ (Tâm thế Giáp Mộc bao bọc — Xem hàng hóa của khách hàng như tài sản của chính gia đình mình để chăm sóc chu đáo).",
        "visual": "Ấm trà mộc khói nghi ngút, ánh mắt kiên định và ấm áp của Chị Thúy, hình ảnh con tàu vượt trùng khơi hướng về chân trời mới.",
        "cliffhanger": "Cảm ơn các bạn đã đồng hành cùng Thúy qua chuỗi chia sẻ này. Hẹn gặp lại mọi người ở những chuyến hải trình tri thức tiếp theo trên kênh ThuyTalk!",
        "conclusion": "Giữ cho tâm mình liêm chính và vững chãi như cây gỗ lớn, thì dù sóng gió thương trường có đổi hướng ra sao, chúng ta vẫn luôn tìm thấy sự bình an và thành công đích thực trong cuộc đời.",
        "status": "Chờ Chị Thúy duyệt"
    }
]

# ==============================================================================
# HÀM THỰC THI GHI VÀO EXCEL VỚI ATOMIC SAVE
# ==============================================================================

def double_up_youtube_series():
    print("=== BẮT ĐẦU QUY TRÌNH DOUBLE UP MASTER SERIES YOUTUBE (CD10 -> CD18) ===")
    
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ Không tìm thấy file: {EXCEL_PATH}")
        return

    wb = openpyxl.load_workbook(EXCEL_PATH)

    font_family = "Arial"
    regular_font = Font(name=font_family, size=10, bold=False, color="000000")
    bold_font = Font(name=font_family, size=10, bold=True, color="1B365D")
    pending_font = Font(name=font_family, size=10, bold=True, color="9A7D0A")

    thin_border = Border(
        left=Side(style="thin", color="D0D7DE"),
        right=Side(style="thin", color="D0D7DE"),
        top=Side(style="thin", color="D0D7DE"),
        bottom=Side(style="thin", color="D0D7DE")
    )

    zebra_fill = PatternFill(start_color="F9FAFC", end_color="F9FAFC", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    pending_fill = PatternFill(start_color="FEF9E7", end_color="FEF9E7", fill_type="solid")

    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

    sheet_configs = [
        ("CM1_VAN_BAN_NONG", CM1_NEW_ITEMS),
        ("CM2_DICH_LUAT_TIENG_NGUOI", CM2_NEW_ITEMS),
        ("CM3_LICH_SU_THUONG_MAI", CM3_NEW_ITEMS),
        ("CM4_TAM_TINH_THUC_CHIEN", CM4_NEW_ITEMS)
    ]

    total_added = 0

    for sheet_name, new_items in sheet_configs:
        ws = wb[sheet_name]
        print(f"\n📂 Đang xử lý Sheet: {sheet_name} (Số dòng hiện tại: {ws.max_row})")
        
        # Lấy danh sách mã chủ đề đã có
        existing_codes = [ws.cell(r, 2).value for r in range(5, ws.max_row + 1)]
        
        for item in new_items:
            code = item["code"]
            if code in existing_codes:
                print(f"   ⚠️ Mã {code} đã tồn tại trong {sheet_name}, bỏ qua.")
                continue

            target_row = ws.max_row + 1
            row_fill = zebra_fill if target_row % 2 == 0 else white_fill

            row_values = [
                item["stt"],
                item["code"],
                item["title"],
                item["format"],
                item["duration"],
                item["ssot"],
                item["hook"],
                item["points"],
                item["visual"],
                item["cliffhanger"],
                item["conclusion"],
                item["status"],
                None # Link kịch bản sẽ được sinh khi viết kịch bản chi tiết
            ]

            for c_idx, val in enumerate(row_values, start=1):
                cell = ws.cell(row=target_row, column=c_idx, value=val)
                cell.border = thin_border
                
                if c_idx == 12: # Cột Trạng Thái Xem Xét
                    cell.font = pending_font
                    cell.fill = pending_fill
                    cell.alignment = center_align
                elif c_idx in [1, 2, 4, 5]:
                    cell.font = bold_font if c_idx in [1, 2] else regular_font
                    cell.fill = row_fill
                    cell.alignment = center_align
                else:
                    cell.font = regular_font
                    cell.fill = row_fill
                    cell.alignment = left_align

            ws.row_dimensions[target_row].height = 140.25
            total_added += 1
            print(f"   -> Đã chèn thành công dòng {target_row}: STT {item['stt']} - {code} - {item['title'][:40]}...")

    # Cập nhật ghi chú trên Sheet TONG_QUAN_LICH_PHAT_SONG
    if "TONG_QUAN_LICH_PHAT_SONG" in wb.sheetnames:
        ws_tongquan = wb["TONG_QUAN_LICH_PHAT_SONG"]
        ws_tongquan.cell(2, 1, value="BẢNG ĐIỀU HÀNH MASTER SERIES YOUTUBE THUYTALK — BATCH 1 (CD01-CD09) & BATCH 2 (CD10-CD18) TỔNG CỘNG 72 CHỦ ĐỀ")
        print("\n✅ Đã cập nhật tiêu đề tổng quan Sheet 1.")

    # Lưu an toàn bằng atomic_save trên cùng ổ đĩa C: (Zero Backup Rule 10)
    wb.save(TMP_PATH)
    wb.close()
    os.replace(TMP_PATH, EXCEL_PATH)
    print(f"\n🎉 HOÀN THÀNH TOÀN DIỆN! Đã double thêm {total_added} chủ đề mới vào file Master Series OneDrive qua atomic_save!")

if __name__ == "__main__":
    double_up_youtube_series()
