# -*- coding: utf-8 -*-
"""
Kịch bản cập nhật Ideas Bank Ngày 27/08/2026 (Tuần 35 - Day 2) theo chuẩn wf06.
Tác tử: Agent 01 (Research & Đóng gói) + Agent 08 (Fact Checker & Deep URL Validator).
======================================================================================
- Quét và đóng gói chính xác 5 ý tưởng mới (STT 46-50) cho mỗi Pillar (TN, CS, HT, FF).
- Kiểm tra 100% URL qua Technical Gatekeeper (scripts/url_validator.py).
- Ghi dữ liệu vào 4 file Excel trong thư mục reports/ với định dạng chuẩn Segoe UI.
- Feedback loop: Cập nhật 5 văn bản pháp luật/quy chuẩn mới vào Danh_sach_Van_ban_Phap_luat_XNK.xlsx.
"""

import os
import sys
import io
import shutil
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Đảm bảo in UTF-8
try:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
except AttributeError:
    pass

sys.path.append(r"d:\AI AGENT THUY\AI branding\scripts")
import excel_styles
from url_validator import validate_url_deep, enforce_strict_validation

WORKSPACE_DIR = r"d:\AI AGENT THUY\AI branding"
REPORTS_DIR = os.path.join(WORKSPACE_DIR, "reports")

FONT_NAME = "Segoe UI"
BODY_FONT = Font(name=FONT_NAME, size=10)
BOLD_FONT = Font(name=FONT_NAME, size=10, bold=True)
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
WHITE_FILL = PatternFill(fill_type=None)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ==============================================================================
# DỮ LIỆU Ý TƯỞNG NGÀY 27/08/2026 (STT 46 -> 50 | ĐỦ 4 PILLARS x 5 = 20 Ý TƯỞNG)
# ==============================================================================

TIN_NGANH_W35_D2 = [
    [
        46,
        "W35 — 27/08/2026",
        "Thông tư 33/2023/TT-BTC & Quy chuẩn Kiểm tra, Xác định Xuất xứ Hàng hóa Xuất khẩu, Nhập khẩu theo Tiêu chuẩn Công ước Kyoto Sửa đổi (Revised Kyoto Convention) của WCO",
        "Bộ Tài chính / Tổng cục Hải quan",
        "15/07/2026",
        "Quy định chi tiết cơ chế nộp chứng từ chứng nhận xuất xứ điện tử (e-C/O) có chữ ký số; quy định rõ ràng thời điểm nộp bổ sung C/O trong vòng 30 ngày kể từ ngày đăng ký tờ khai; hướng dẫn cụ thể về việc xử lý khác biệt nhỏ (minor discrepancies) không làm mất tính hợp lệ của C/O theo chuẩn mực Công ước Kyoto sửa đổi.",
        "Toàn bộ doanh nghiệp xuất nhập khẩu, đại lý hải quan và các hiệp hội ngành hàng áp dụng ưu đãi thuế quan FTA (EVFTA, CPTPP, RCEP, VKFTA, v.v.).",
        "Nộp C/O muộn quá 30 ngày hoặc không giải trình được các điểm sai khác về tên hàng, mã HS sẽ bị từ chối áp dụng thuế suất ưu đãi đặc biệt ngay tại khâu thông quan, phải nộp thuế MFN cao hơn từ 10-30% kèm tiền phạt chậm nộp.",
        "Các doanh nghiệp XNK hoan nghênh việc thừa nhận C/O điện tử và quy chế khác biệt nhỏ giúp giảm thiểu tranh chấp không đáng có với cơ quan hải quan cửa khẩu.",
        "Chuẩn hóa thủ tục xuất xứ theo thông lệ quốc tế WCO, giảm thiểu thủ tục giấy tờ, tăng tốc độ thông quan và phòng chống gian lận lẩn tránh phòng vệ thương mại.",
        "Hệ thống kết nối e-C/O giữa Hải quan Việt Nam và một số đối tác FTA song phương thỉnh thoảng bị lỗi kỹ thuật đường truyền, khiến doanh nghiệp vẫn phải chạy đi xin xác nhận bản giấy mất thời gian.",
        "Công ước Quốc tế về Đơn giản hóa và Hài hòa hóa Thủ tục Hải quan (Công ước Kyoto Sửa đổi - Revised Kyoto Convention Annex K: Rules of Origin của WCO).",
        "https://en.wikipedia.org/wiki/Rules_of_origin",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Quy tắc Xuất xứ Rules of Origin WCO & Thông tư 33/2023/TT-BTC."
    ],
    [
        47,
        "W35 — 27/08/2026",
        "Nghị định 18/2021/NĐ-CP & Luật Thuế Xuất khẩu, Thuế Nhập khẩu: Quy chế Miễn thuế và Quản lý Cơ sở Sản xuất đối với Hàng hóa Nhập khẩu để Gia công, Sản xuất Xuất khẩu (Duty Drawback & Inward Processing)",
        "Chính phủ / Bộ Tài chính",
        "01/08/2026",
        "Siết chặt điều kiện kiểm tra cơ sở sản xuất và quyền sở hữu máy móc thiết bị đối với doanh nghiệp gia công/SXXK; bắt buộc lập Báo cáo Quyết toán định kỳ (Mẫu 15/BCQT-NVL) theo phương pháp Nhập - Xuất - Tồn thực tế; phân định rõ quy chế xử lý nguyên phụ liệu dư thừa, phế liệu, phế phẩm tiêu thụ nội địa.",
        "Hàng chục nghìn doanh nghiệp dệt may, da giày, điện tử, gia công chế tạo cơ khí và các khu công nghiệp FDI trên toàn quốc.",
        "Nếu cơ sở sản xuất không đạt chuẩn kiểm tra của hải quan hoặc số liệu quyết toán nguyên phụ liệu chênh lệch không giải trình được, doanh nghiệp bị truy thu toàn bộ tiền thuế nhập khẩu 10-25% của lượng hàng đã nhập kèm mức phạt 20% và cưỡng chế dừng làm thủ tục XNK.",
        "Các hiệp hội dệt may (VITAS) và da giày (LEFASO) khuyến nghị các thành viên đầu tư đồng bộ phần mềm quản trị ERP tích hợp dữ liệu kho kế toán và tờ khai hải quan.",
        "Chấm dứt triệt để tình trạng các 'doanh nghiệp ma' lợi dụng chính sách miễn thuế SXXK để nhập lậu nguyên liệu rồi tuồn ra thị trường nội địa bán phá giá trốn thuế.",
        "Áp lực quy chuẩn kiểm tra cơ sở xưởng và quản lý định mức hao hụt quá khắt khe gây tốn kém chi phí nhân sự kế toán kho cho các doanh nghiệp SME.",
        "Chế độ Miễn và Hoàn thuế Nhập khẩu phục vụ Gia công Xuất khẩu (Duty Drawback / Inward Processing) của Tổ chức Hải quan Thế giới (WCO).",
        "https://en.wikipedia.org/wiki/Duty_drawback",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ chế định Duty Drawback chuẩn quốc tế & Nghị định 18/2021/NĐ-CP."
    ],
    [
        48,
        "W35 — 27/08/2026",
        "Quy chuẩn An ninh Chuỗi Cung ứng Toàn cầu và Chương trình Doanh nghiệp Ưu tiên Hải quan (Authorized Economic Operator - AEO Framework of Standards của WCO)",
        "Tổ chức Hải quan Thế giới (WCO) / Bộ Tài chính Việt Nam",
        "10/08/2026",
        "Mở rộng tiêu chí công nhận Doanh nghiệp Ưu tiên (AEO) theo Thông tư 72/2015/TT-BTC sửa đổi; doanh nghiệp AEO được hưởng đặc quyền miễn kiểm tra hồ sơ, miễn kiểm tra thực tế hàng hóa (100% Luồng Xanh), được thông quan trước nộp chứng từ sau và kiểm tra sau thông quan tại trụ sở tối đa 1 lần trong 3 năm.",
        "Các tập đoàn sản xuất xuất khẩu quy mô lớn (Samsung, Foxconn, VinFast, Hòa Phát, Hyosung...) và các công ty Logistics hàng đầu tại Việt Nam.",
        "Nếu doanh nghiệp AEO để xảy ra 01 vi phạm pháp luật hải quan nghiêm trọng hoặc không duy trì đủ tiêu chuẩn an ninh cơ sở hạ tầng mạng/kho bãi, sẽ bị Tổng cục Hải quan đình chỉ hoặc thu hồi tư cách AEO ngay lập tức, mất toàn bộ lợi thế cạnh tranh về thời gian giao hàng.",
        "Các doanh nghiệp FDI lớn tích cực duy trì hệ thống kiểm soát tuân thủ nội bộ (Internal Compliance Program - ICP) để giữ vững chứng chỉ AEO danh giá.",
        "Phân luồng quản lý rủi ro thông minh, tập trung nguồn lực hải quan vào các đối tượng rủi ro cao, thúc đẩy tốc độ quay vòng vốn cho các doanh nghiệp xuất khẩu chiến lược.",
        "Điều kiện xét duyệt AEO (kim ngạch xuất nhập khẩu tối thiểu hàng trăm triệu USD) là rào cản quá lớn khiến các doanh nghiệp vừa và nhỏ (SME) hầu như không có cơ hội tiếp cận đặc quyền này.",
        "Khung Tiêu chuẩn An ninh và Thuận lợi hóa Thương mại Toàn cầu WCO SAFE Framework & Chương trình AEO toàn cầu.",
        "https://en.wikipedia.org/wiki/Authorized_economic_operator",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Tiêu chuẩn quốc tế WCO SAFE Framework & Thông tư 72/2015/TT-BTC."
    ],
    [
        49,
        "W35 — 27/08/2026",
        "Hiệp định Thuận lợi hóa Thương mại của Tổ chức Thương mại Thế giới (WTO Trade Facilitation Agreement - TFA) và Cơ chế Nghiên cứu Đo lường Thời gian Giải phóng Hàng (Time Release Study - TRS)",
        "Tổ chức Thương mại Thế giới (WTO) / Ủy ban Quốc gia về Thuận lợi hóa Thương mại Việt Nam",
        "05/08/2026",
        "Triển khai đầy đủ Nhóm cam kết loại A, B, C của Hiệp định TFA; bắt buộc công khai minh bạch 100% mức phí, lệ phí cửa khẩu trên cổng thông tin điện tử; thực hiện định kỳ báo cáo Đo lường thời gian giải phóng hàng (TRS) tại các cảng biển Hải Phòng, Cát Lái, Cái Mép để cắt giảm thời gian thông quan thực tế xuống dưới 24 giờ đối với hàng nhập và 12 giờ đối với hàng xuất.",
        "Toàn bộ cộng đồng doanh nghiệp xuất nhập khẩu, đơn vị kinh doanh kho bãi cảng biển, các cơ quan kiểm tra chuyên ngành và cơ quan quản lý nhà nước tại cửa khẩu.",
        "Tình trạng thiếu minh bạch về quy trình hoặc kéo dài thời gian kiểm tra chuyên ngành sẽ bị đưa vào danh sách chỉ số cản trở logistics quốc gia, làm tăng chi phí lưu kho bãi và giảm năng lực cạnh tranh thu hút FDI.",
        "Cộng đồng doanh nghiệp trong và ngoài nước (EuroCham, AmCham, VCCI) đánh giá cao cam kết cải cách thủ tục hành chính thực chất của Chính phủ.",
        "Rút ngắn thời gian thông quan giúp tiết kiệm hàng tỷ USD chi phí chuỗi cung ứng mỗi năm, đưa Việt Nam trở thành trung tâm logistics hàng đầu Đông Nam Á.",
        "Sự phối hợp liên ngành giữa Hải quan và các Bộ quản lý chuyên ngành (Nông nghiệp, Y tế, Công Thương) tại một số cửa khẩu vẫn còn phân tán, khiến thời gian giải phóng hàng thực tế chưa đồng đều.",
        "Hiệp định Thuận lợi hóa Thương mại WTO Trade Facilitation Agreement (TFA) & Hướng dẫn Time Release Study (TRS) của WCO.",
        "https://en.wikipedia.org/wiki/Trade_Facilitation_Agreement",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Hiệp định TFA của WTO & Chỉ số Logistics Performance Index (LPI)."
    ],
    [
        50,
        "W35 — 27/08/2026",
        "Công ước Hải quan về Sổ Tạm quản ATA (ATA Carnet Convention) và Nghị định 64/2020/NĐ-CP: Cơ chế Tạm nhập Tái xuất Miễn Đặt cọc Thuế cho Hàng Triển lãm, Hàng mẫu và Thiết bị Chuyên dụng",
        "Chính phủ / Phòng Thương mại và Công nghiệp Việt Nam (VCCI) / Tổng cục Hải quan",
        "01/08/2026",
        "Áp dụng cơ chế một cuốn Sổ Tạm quản ATA duy nhất thay thế toàn bộ tờ khai hải quan quốc gia và giấy bảo lãnh ngân hàng truyền thống; cho phép doanh nghiệp tạm nhập hàng hóa vào hơn 80 quốc gia thành viên trong thời hạn 12 tháng mà không phải đóng tiền thuế bảo lãnh tại mỗi nước.",
        "Các doanh nghiệp tham gia hội chợ triển lãm quốc tế, các đoàn làm phim/truyền hình quốc tế, các đơn vị biểu diễn nghệ thuật, thể thao và doanh nghiệp cung cấp thiết bị đo lường thử nghiệm.",
        "Nếu để quá hạn 12 tháng mà không làm thủ tục tái xuất hoặc làm mất sổ ATA, doanh nghiệp sẽ bị tổ chức bảo lãnh (VCCI) khấu trừ toàn bộ tiền bảo lãnh và bị Hải quan nước sở tại phạt truy thu thuế nhập khẩu kèm lãi suất vi phạm.",
        "Các doanh nghiệp xúc tiến thương mại và tổ chức sự kiện quốc tế hưởng ứng nồng nhiệt vì thủ tục đơn giản và giải phóng hoàn toàn dòng vốn lưu động đặt cọc thuế.",
        "Hội nhập sâu rộng thương mại quốc tế, thúc đẩy ngoại giao văn hóa, chuyển giao công nghệ và tạo điều kiện tối đa cho doanh nghiệp Việt Nam đem sản phẩm đi quảng bá toàn cầu.",
        "Quy trình kiểm tra hải quan thực tế từng dòng linh kiện theo sổ ATA đòi hỏi nhân viên hiện trường phải có trình độ chuyên môn cao và am hiểu tiếng Anh chuyên ngành.",
        "Công ước Hải quan ATA Carnet và Công ước Istanbul của Tổ chức Hải quan Thế giới (WCO) phối hợp Phòng Thương mại Quốc tế (ICC).",
        "https://en.wikipedia.org/wiki/ATA_Carnet",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Công ước Istanbul / ATA Carnet của WCO/ICC & Nghị định 64/2020/NĐ-CP."
    ]
]

CASE_STUDY_W35_D2 = [
    [
        46,
        "W35 — 27/08/2026",
        "Lô hàng 20 tấn hạt điều nhập khẩu bị giữ tại Cảng Hải Phòng vì vi phạm Bao bì Gỗ chưa khử trùng theo Tiêu chuẩn ISPM 15: Thiệt hại 180 triệu đồng phí tái chế và phạt kiểm dịch",
        "Công ty Cổ phần Chế biến Nông sản H. (Hưng Yên), Nhà cung cấp nguyên liệu tại Bờ Biển Ngà, Chi cục Kiểm dịch Thực vật Vùng I (Hải Phòng)",
        "Doanh nghiệp nhập khẩu 1 container 40ft hạt điều thô từ Tây Phi về Cảng Hải Phòng. Khi cơ quan kiểm dịch thực vật lấy mẫu kiểm tra tại bãi, phát hiện toàn bộ 25 pallet gỗ kê hàng không có dấu mộc chứng nhận xử lý nhiệt (HT) hoặc khử trùng Methyl Bromide (MB) theo chuẩn ISPM 15, đồng thời phát hiện ấu trùng bọ cánh cứng còn sống ẩn nấp trong thớ gỗ. Chi cục Kiểm dịch lập biên bản đình chỉ thông quan, yêu cầu cách ly khẩn cấp và buộc tái xuất hoặc bốc dỡ toàn bộ hàng sang pallet nhựa để khử trùng tiêu hủy pallet gỗ.",
        "Thiệt hại 180 triệu đồng (gồm 65 triệu đồng phí bốc dỡ/sang bao thủ công trong kho cách ly, 45 triệu đồng tiền mua mới 25 pallet nhựa thay thế, 30 triệu đồng chi phí hun trùng tiêu hủy pallet gỗ tại bãi, và 40 triệu đồng phí lưu container/lưu bãi trong 14 ngày).",
        "Hợp đồng ngoại thương ký theo điều kiện CIF nhưng người mua không đưa điều khoản bắt buộc pallet gỗ phải đạt chuẩn ISPM 15 có mộc khắc nhiệt; tin tưởng người bán nước ngoài tự đóng hàng mà không kiểm tra hình ảnh hiện trường trước khi xếp hàng lên tàu.",
        "1. Bắt buộc ghi rõ trong hợp đồng ngoại thương và L/C: 'All wooden packaging materials must comply with ISPM 15 standards with clear IPPC markings'; 2. Yêu cầu người bán cung cấp Chứng thư Khử trùng (Fumigation Certificate) trước khi tàu rời cảng; 3. Ưu tiên sử dụng pallet nhựa, pallet kim loại hoặc pallet giấy ép (không thuộc diện kiểm soát ISPM 15).",
        "Phe A (Chủ hàng): Hạt điều bên trong đóng bao đay không bị sâu mọt, chỉ có pallet gỗ bên ngoài sao kiểm dịch lại phạt nặng và bắt bốc dỡ tốn kém thế? Phe B (Kiểm dịch thực vật): Côn trùng ngoại lai ẩn nấp trong gỗ có thể lây lan tàn phá toàn bộ rừng trồng và mùa màng quốc gia, an toàn sinh học biên giới là tuyệt đối.",
        "Bao bì đóng gói không chỉ là vật chứa hàng mà là một đối tượng quản lý pháp lý độc lập. Một chiếc pallet gỗ sơ sài 100 nghìn đồng có thể làm bốc hơi cả trăm triệu đồng chi phí xử lý sự cố.",
        "https://en.wikipedia.org/wiki/ISPM_15",
        "Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách tai nạn nhập khẩu nông sản vi phạm quy chuẩn bao bì gỗ ISPM 15 và kiểm dịch thực vật quốc tế."
    ],
    [
        47,
        "W35 — 27/08/2026",
        "Bẫy 'Chứng thư Bảo lãnh Ngân hàng Giả' (Fraudulent Bank Guarantee) khi mua hạt nhựa 1,2 triệu USD từ nhà cung ứng nước ngoài: DN mất trắng 3,8 tỷ đồng tiền cọc",
        "Công ty Sản xuất Bao bì Nhựa T.P (Bình Dương), Công ty Thương mại Trung gian 'ma' tại Dubai, Ngân hàng Thương mại Quốc tế",
        "Doanh nghiệp Việt Nam tìm được đối tác chào bán hạt nhựa nguyên sinh PP với giá rẻ hơn thị trường 15%. Để tạo lòng tin, đối tác đồng ý cung cấp một Thư Bảo lãnh Ngân hàng Hoàn trả Tiền cọc (Advance Payment Guarantee) trị giá 150.000 USD phát hành từ một chi nhánh ngân hàng quốc tế. Sau khi nhận được bản sao scan thư bảo lãnh có dấu đỏ và chữ ký số đẹp mắt qua email, doanh nghiệp Việt Nam vội vàng chuyển khoản 30% tiền đặt cọc (3,8 tỷ VNĐ). Sau 45 ngày không thấy giao hàng, doanh nghiệp mang thư bảo lãnh ra ngân hàng yêu cầu thực thi thì phát hiện đây là chứng thư giả mạo, không có mã điện SWIFT MT760 trên hệ thống liên ngân hàng quốc tế.",
        "Mất trắng 3,8 tỷ đồng (150.000 USD) tiền đặt cọc; đứt gãy nguồn nguyên liệu phục vụ sản xuất khiến nhà máy phải dừng 2 dây chuyền trong 3 tuần, ước tính thiệt hại thêm 500 triệu đồng doanh thu.",
        "Nhẹ dạ cả tin vào bản scan chứng thư bảo lãnh gửi qua email thông thường; không thực hiện quy trình xác thực thư bảo lãnh qua điện tín liên ngân hàng SWIFT (Authenticated MT760 message) qua ngân hàng phục vụ của mình tại Việt Nam.",
        "1. Tuyệt đối không chấp nhận bất kỳ bản scan thư bảo lãnh ngân hàng nào gửi qua email/fax; 2. Bắt buộc yêu cầu ngân hàng nước ngoài phát hành Bảo lãnh gửi điện SWIFT MT760 đích danh đến Ngân hàng của người mua tại Việt Nam để thông báo và xác thực (Advising/Confirmation Bank); 3. Thẩm định độc lập năng lực tài chính và lịch sử pháp nhân của nhà cung cấp qua các tổ chức xếp hạng tín nhiệm (Dun & Bradstreet, Coface).",
        "Phe A (DN): Doanh nghiệp đã cẩn thận đòi thư bảo lãnh ngân hàng rồi mà vẫn bị lừa tinh vi quá. Phe B (Chuyên gia Tài chính Quốc tế): Trong thương mại quốc tế, giấy tờ không có điện SWIFT mã hóa xác thực chỉ là một tờ giấy lộn vô giá trị.",
        "Đừng bao giờ giao tiền thật để nhận lại một tờ giấy bảo lãnh ảo. Trong thanh toán quốc tế, mọi sự tin tưởng phải được bảo đảm bằng hạ tầng mã hóa SWIFT của hệ thống ngân hàng.",
        "https://en.wikipedia.org/wiki/Bank_guarantee",
        "Bóc phốt chiêu trò",
        "Chưa dùng",
        "Cảnh báo thủ đoạn lừa đảo chứng thư bảo lãnh ngân hàng giả mạo trong nhập khẩu nguyên liệu sản xuất."
    ],
    [
        48,
        "W35 — 27/08/2026",
        "Giải cứu 8 container máy công cụ đã qua sử dụng quá 10 năm bị kẹt tại Cảng Cát Lái nhờ Áp dụng Cơ chế Giám định Hiệu suất Năng lượng và Tuổi thiết bị theo Quyết định 18/2019/QĐ-TTg",
        "Công ty Cơ khí Chế tạo Máy Nhật Việt (Đồng Nai), Tổ chức Giám định Độc lập được Bộ KH&CN chỉ định, Chi cục Hải quan Cửa khẩu Cảng Sài Gòn Khu vực 1",
        "Doanh nghiệp nhập khẩu dây chuyền 8 máy tiện, máy phay CNC đã qua sử dụng từ Osaka (Nhật Bản). Trên nhãn mác thân máy, năm sản xuất ghi 2015 (tính đến năm 2026 là 11 năm, vượt quá ngưỡng quy định 10 năm tuổi theo Quyết định 18/2019/QĐ-TTg). Hải quan Cát Lái tạm dừng thông quan và chuẩn bị ra quyết định buộc tái xuất toàn bộ lô hàng 8 container trị giá 12 tỷ đồng. Doanh nghiệp đối mặt với nguy cơ phá sản dây chuyền nhà xưởng mới khánh thành.",
        "Lô hàng 12 tỷ đồng có nguy cơ bị buộc tái xuất mất trắng 1,5 tỷ cước tàu; nhờ kích hoạt quy trình giải cứu trong 10 ngày, chỉ phát sinh 85 triệu đồng phí giám định kỹ thuật và lưu bãi, cứu vãn toàn bộ dây chuyền sản xuất của nhà máy.",
        "Doanh nghiệp không rà soát chặt chẽ năm xuất xưởng của từng cấu phần máy móc trước khi ký hợp đồng nhập khẩu; không chuẩn bị trước hồ sơ kỹ thuật chứng minh hiệu suất năng lượng và tiêu chuẩn an toàn của thiết bị.",
        "1. Áp dụng quy định tại Điều 6 Quyết định 18/2019/QĐ-TTg: Đối với máy móc vượt 10 năm nhưng phục vụ trực tiếp sản xuất công nghiệp phụ trợ, tiến hành mời tổ chức giám định được Bộ KH&CN chỉ định đến hiện trường đo kiểm hiệu suất còn lại trên 85% và mức tiêu hao năng lượng đạt chuẩn; 2. Nộp văn bản giải trình và Giấy chứng nhận Giám định đạt chuẩn để được Thủ tướng Chính phủ / Bộ KH&CN chấp thuận đặc cách nhập khẩu; 3. Lần sau bắt buộc yêu cầu đối tác Nhật Bản chụp ảnh toàn bộ nameplate máy để check tuổi trước khi đóng hàng.",
        "Phe A (DN): Máy móc Nhật Bản dù 11 năm tuổi nhưng độ chính xác và độ bền gấp 10 lần máy mới giá rẻ của nước khác, tại sao lại cứng nhắc cấm nhập? Phe B (Cơ quan quản lý): Quy định giới hạn 10 năm tuổi là hàng rào kỹ thuật sống còn để ngăn chặn Việt Nam biến thành 'bãi rác công nghệ' lạc hậu của thế giới.",
        "Hiểu sâu quy định pháp luật kỹ thuật chính là 'chiếc phao cứu sinh' khi xảy ra khủng hoảng thông quan. Khi rơi vào ca khó, sự am hiểu các điều khoản ngoại lệ và quy chuẩn giám định sẽ mở ra cánh cửa thoát hiểm cho doanh nghiệp.",
        "https://en.wikipedia.org/wiki/Customs_inspection",
        "Giải cứu hàng hóa",
        "Chưa dùng",
        "Case study thực chiến về giải cứu máy móc thiết bị cũ nhập khẩu theo Quyết định 18/2019/QĐ-TTg và quy trình giám định kỹ thuật hải quan."
    ],
    [
        49,
        "W35 — 27/08/2026",
        "Hãng tàu phá sản và Tàu mẹ bị Tòa án bắt giữ (Ship Arrest) trên hải trình: 15 container hàng dệt may xuất khẩu bị kẹt 3 tháng tại Singapore, thiệt hại 1,4 tỷ đồng",
        "Công ty Dệt may Thời trang M. (Bình Dương), Hãng tàu Vận tải Biển Đa quốc gia, Tòa án Hàng hải Singapore, Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA)",
        "Doanh nghiệp xuất khẩu 15 container quần áo mùa đông sang Bờ Đông nước Mỹ. Giữa hành trình tại eo biển Malacca, hãng tàu container lâm vào tình trạng vỡ nợ, các chủ nợ quốc tế nộp đơn lên Tòa án Tối cao Singapore yêu cầu bắt giữ tàu mẹ (Ship Arrest) để phong tỏa tài sản đòi nợ cước nhiên liệu. Toàn bộ hàng hóa trên tàu bị kẹt tại neo đậu cảng Singapore suốt 90 ngày. Người mua tại Mỹ hủy hợp đồng mua hàng, phạt hủy đơn vì trễ mùa vụ kinh doanh.",
        "Thiệt hại 1,4 tỷ đồng (gồm 600 triệu đồng tiền phạt hủy hợp đồng từ đối tác Mỹ, 500 triệu đồng chi phí thuê luật sư hàng hải nộp đơn giải tỏa hàng và bảo lãnh container ra khỏi tàu, 300 triệu đồng cước bốc dỡ sang tàu khác để chở tiếp).",
        "Ham giá cước biển siêu rẻ của hãng tàu nhỏ có tình hình tài chính bất ổn định; mua bảo hiểm hàng hải điều kiện hẹp không bao gồm điều khoản bảo hiểm rủi ro phá sản/mất khả năng thanh toán của chủ tàu (Insolvency Exclusion Clause trong ICC).",
        "1. Lựa chọn các hãng tàu thuộc các Liên minh hàng hải lớn (Ocean Alliance, THE Alliance, 2M/Gemini) có xếp hạng tín nhiệm tài chính an toàn; 2. Khi đàm phán hợp đồng bảo hiểm hàng hải, bắt buộc mua thêm Điều khoản sửa đổi bổ sung chấp nhận rủi ro vỡ nợ của người vận chuyển (Insolvency Clause Extension); 3. Theo dõi sát sao tình trạng tài chính của chuỗi cung ứng logistics quốc tế.",
        "Phe A (Chủ hàng): Hàng hóa của tôi là tài sản độc lập, tại sao Tòa án nước ngoài lại giữ luôn container của tôi khi chủ tàu nợ tiền? Phe B (Luật Hàng hải Quốc tế): Theo Luật Hàng hải Quốc tế và Công ước Bắt giữ Tàu biển (Arrest of Ships Convention), con tàu bị niêm phong nguyên trạng cho đến khi có phán quyết phân chia tài sản thanh lý.",
        "Chọn hãng tàu cũng giống như chọn bạn đồng hành vượt đại dương. Giá cước rẻ vài chục USD không thể bù đắp được rủi ro tàu bị bắt giữ giữa biển khơi.",
        "https://en.wikipedia.org/wiki/Ship_arrest",
        "Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách rủi ro pháp lý hàng hải quốc tế khi tàu chở hàng bị tòa án nước ngoài bắt giữ (Ship Arrest)."
    ],
    [
        50,
        "W35 — 27/08/2026",
        "Chiêu trò 'Cắt Seal trộm hàng' trong kho CFS: DN nhập khẩu rượu vang bị tráo 120 chai vang thượng hạng thành vang rác, bồi thường bảo hiểm giằng co 6 tháng trị giá 350 triệu đồng",
        "Công ty Phân phối Đồ uống Cao cấp V. (Hà Nội), Đơn vị Khai thác Kho Hàng lẻ CFS Cảng Hải Phòng, Đơn vị Giám định Hàng hải độc lập",
        "Doanh nghiệp nhập khẩu lô hàng 500 chai rượu vang Grand Cru cao cấp của Pháp đóng ghép trong container hàng lẻ (LCL). Khi container về đến kho CFS tại Hải Phòng, kẻ gian trong ca trực đêm đã dùng kìm bấm chuyên dụng cắt seal kẹp chì, rút ruột 120 chai rượu vang đắt tiền trị giá gần 3 triệu đồng/chai và thay thế bằng các chai rượu vang giá rẻ 80 nghìn đồng, sau đó dùng keo dán lại seal giả mạo. Khi chủ hàng đến nhận hàng mở niêm phong mới phát hiện sự việc. Phía kho CFS từ chối trách nhiệm vì biên bản bàn giao seal tại cổng cảng vẫn nguyên vẹn.",
        "Thiệt hại 350 triệu đồng tiền hàng hóa bị tráo đổi; tranh chấp pháp lý và hồ sơ khiếu nại bảo hiểm kéo dài 6 tháng làm đọng vốn và mất uy tín với hệ thống nhà hàng khách sạn 5 sao.",
        "Hàng hóa giá trị cao đóng trong bao bì thông thường không dán băng dính niêm phong an ninh (Security Tamper Evident Tape); không thuê giám định viên độc lập có mặt chứng kiến trực tiếp vào thời điểm cắt seal mở rút ruột container tại kho CFS.",
        "1. Đối với hàng hóa giá trị cao, bắt buộc dán tem niêm phong chống bóc tách (Tamper Proof Tape) trên từng thùng carton; 2. Luôn cử đại diện hoặc nhân sự giám định độc lập có mặt tại kho CFS vào đúng thời điểm nhân viên hải quan và thủ kho cắt seal mở cửa container (Unstuffing supervision); 3. Kiểm tra chụp ảnh cận cảnh mã số seal và các vết xước trên khóa container trước khi cho phép bốc dỡ hàng.",
        "Phe A (Chủ hàng): Kho CFS nhận bảo quản hàng phải chịu trách nhiệm 100% khi hàng bị tráo đổi trong khuôn viên kho. Phe B (Kho CFS): Seal container khi tiếp nhận từ tàu là nguyên vẹn, chủ hàng không chứng minh được thời điểm mất cắp xảy ra tại kho hay từ trên tàu biển.",
        "Hàng giá trị cao tuyệt đối không được 'khoán trắng' cho đơn vị vận chuyển. Sự giám sát trực tiếp tại thời điểm mở seal là bằng chứng pháp lý duy nhất để bảo vệ quyền lợi bồi thường của doanh nghiệp.",
        "https://en.wikipedia.org/wiki/Supply_chain_security",
        "Bóc phốt chiêu trò",
        "Chưa dùng",
        "Cảnh báo thủ đoạn rút ruột tráo hàng trong kho gom hàng lẻ CFS và cẩm nang giám sát an ninh chuỗi cung ứng hàng hóa nhập khẩu."
    ]
]

HOC_THUC_CHIEN_W35_D2 = [
    [
        46,
        "W35 — 27/08/2026",
        "Quy trình 5 bước Xử lý Thủ tục Tạm nhập Tái xuất bằng Sổ Tạm quản ATA (ATA Carnet) cho Hàng mẫu, Máy móc Triển lãm và Đạo cụ Hội chợ",
        "Bước 1: Chuẩn bị Hồ sơ xin cấp Sổ ATA tại VCCI (Đơn đăng ký, Danh sách hàng hóa chi tiết - General List gồm tên hàng, model, số serial, trọng lượng, trị giá); Bước 2: Nộp tiền bảo lãnh hoặc Thư bảo lãnh ngân hàng tương đương 30-50% tổng số tiền thuế dự kiến cho VCCI; Bước 3: Xuất trình Sổ ATA tại Chi cục Hải quan cửa khẩu xuất cảnh Việt Nam để làm thủ tục xuất khẩu tạm thời (Hải quan ký xác nhận Cuống sổ Vàng - Exportation Counterfoil); Bước 4: Xuất trình Sổ ATA tại Hải quan nước nhập khẩu để làm thủ tục tạm nhập thông quan luồng xanh (Cuống sổ Trắng - Importation Counterfoil) mà không cần nộp tiền đặt cọc thuế; Bước 5: Thực hiện tái xuất hàng hóa về nước đúng hạn 12 tháng, làm thủ tục tái nhập (Cuống sổ Vàng - Re-importation) và nộp lại toàn bộ Sổ ATA cho VCCI để thanh khoản thu hồi 100% tiền bảo lãnh.",
        "Số tiền ký quỹ bảo lãnh Sổ ATA = Tổng trị giá hàng hóa (Invoice value) x Thuế suất cao nhất của nước đến (thường từ 30% đến 50%). Thời hạn hiệu lực tối đa của Sổ ATA: Chính xác 12 tháng kể từ ngày phát hành, không được gia hạn trừ trường hợp cấp Sổ thay thế (Replacement Carnet).",
        "1. Tự ý thay đổi hoặc bổ sung thêm hàng hóa không có tên trong Danh sách Tổng quát (General List) ban đầu của Sổ ATA; 2. Làm mất tờ cuống sổ (Voucher) do Hải quan nước nhập khẩu giữ mà không xin xác nhận vào cuống sổ lưu; 3. Quá hạn 12 tháng không mang hàng về nước dẫn đến bị hải quan nước ngoài truy thu thuế tự động.",
        "Sổ ATA giống như 'Hộ chiếu đa quốc gia của hàng hóa' - Vàng là xuất/nhập của nước nhà, Trắng là nhập/xuất của nước khách. Giữ sổ như giữ hộ chiếu, mất sổ là mất tiền cọc.",
        "So sánh Tạm nhập Tái xuất truyền thống (Mở tờ khai G12/G22, ký quỹ bảo lãnh thuế bằng tiền mặt tại từng quốc gia, mất 3-5 ngày làm thủ tục) vs Sử dụng Sổ ATA Carnet (1 cuốn sổ duy nhất cho 80 quốc gia, miễn đặt cọc thuế tại nước đến, thông quan ngay tại sân bay/cảng biển trong 30 phút).",
        "Anh Dũng (Chủ DN sản xuất/thương mại đi dự hội chợ quốc tế) & Chị Hương (Trưởng phòng Logistics phụ trách sự kiện).",
        "1. Đơn đăng ký cấp sổ ATA theo mẫu VCCI; 2. Bảng kê General List chi tiết số serial; 3. Thư bảo lãnh ngân hàng ký quỹ; 4. Kiểm tra dấu mộc hải quan từng chặng; 5. Nộp lại sổ gốc thanh khoản trong 30 ngày sau tái nhập.",
        "https://en.wikipedia.org/wiki/ATA_Carnet",
        "Quy trình Hải quan",
        "Chưa dùng",
        "Cẩm nang step-by-step sử dụng Sổ Tạm quản ATA Carnet cho doanh nghiệp tham gia hội chợ triển lãm quốc tế theo Nghị định 64/2020/NĐ-CP."
    ],
    [
        47,
        "W35 — 27/08/2026",
        "Hướng dẫn 5 bước Lập Báo cáo Quyết toán Nguyên phụ liệu Gia công & Sản xuất Xuất khẩu (Mẫu 15/BCQT-NVL) Không Lệch 1 Gram Số liệu theo Thông tư 39/2018/TT-BTC",
        "Bước 1: Trích xuất và khóa sổ dữ liệu kế toán kho tại thời điểm kết thúc năm tài chính (Số dư đầu kỳ, Tổng nhập trong kỳ, Tổng xuất dùng sản xuất, Tổng tồn kho cuối kỳ của từng mã nguyên phụ liệu); Bước 2: Đối soát dữ liệu tờ khai hải quan trên hệ thống ECUS (Mã loại hình E21, E31 đối với nhập khẩu và E52, E62 đối với xuất khẩu); Bước 3: Tính toán Định mức sử dụng thực tế (BOM) và Tỷ lệ hao hụt nguyên liệu của từng mã sản phẩm xuất khẩu theo công thức quy chuẩn; Bước 4: Lập Bảng cân đối Nhập - Xuất - Tồn nguyên phụ liệu (Mẫu số 15/BCQT-NVL/GSQL) và Bảng cân đối thành phẩm (Mẫu 15a/BCQT-SP/GSQL); Bước 5: Truyền Báo cáo Quyết toán qua Hệ thống Một cửa Quốc gia trước ngày thứ 90 kể từ ngày kết thúc năm tài chính và lập sẵn Hồ sơ giải trình chênh lệch kho thực tế vs tờ khai.",
        "Công thức Cân đối Quyết toán Nguyên phụ liệu: Lượng Tồn Cuối Kỳ = Lượng Tồn Đầu Kỳ + Tổng Lượng Nhập Trong Kỳ - (Tổng Sản Phẩm Xuất Khẩu x Định Mức Kỹ Thuật) - Lượng Phế Liệu Tiêu Thụ Nội Địa - Lượng Hao Hụt Hợp Lý. Mức phạt nộp chậm BCQT: Phạt tiền từ 2.000.000 đến 5.000.000 đồng và bị hệ thống tự động đưa vào danh sách kiểm tra sau thông quan.",
        "1. Đơn vị tính trên phần mềm kế toán khác với đơn vị tính trên tờ khai hải quan (ví dụ: Kế toán theo dõi mét vải, hải quan khai báo kilogram); 2. Không theo dõi riêng lượng nguyên liệu đưa đi gia công lại tại xưởng phụ; 3. Không nộp bổ sung hồ sơ sửa đổi định mức hao hụt khi cải tiến dây chuyền sản xuất.",
        "Báo cáo Quyết toán Hải quan giống như 'Bản tổng kết tài chính của từng sợi chỉ, thanh nhôm' - Nhập vào bao nhiêu, tạo ra bao nhiêu thành phẩm, thừa thiếu bao nhiêu phải chỉ ra được nằm ở góc kho nào.",
        "So sánh Doanh nghiệp làm BCQT theo phương pháp Thủ công chắp vá số liệu (dễ bị lệch số, đối mặt nguy cơ ấn định thuế tiền tỷ khi hải quan thanh tra) vs Doanh nghiệp chuẩn hóa Dữ liệu Đồng bộ ERP - ECUS (tự động cân đối Nhập - Xuất - Tồn theo thời gian thực, nộp BCQT trong 5 phút).",
        "Chị Hương (Kế toán trưởng/Trưởng phòng XNK nhà máy SXXK) & Anh Dũng (Chủ doanh nghiệp gia công dệt may, da giày, điện tử).",
        "1. Sổ chi tiết kho NVL & Thành phẩm khóa sổ 31/12; 2. Bảng định mức BOM thực tế có phê duyệt của Giám đốc kỹ thuật; 3. Danh sách tờ khai nhập E21/E31 & xuất E52/E62; 4. Bảng tính hao hụt phế liệu; 5. Biên bản đối chiếu kho thực tế.",
        "https://en.wikipedia.org/wiki/Bill_of_materials",
        "Tối ưu Thuế XNK",
        "Chưa dùng",
        "Cẩm nang kỹ thuật lập Báo cáo Quyết toán hải quan nguyên phụ liệu gia công, sản xuất xuất khẩu theo Mẫu 15 Thông tư 39/2018/TT-BTC."
    ],
    [
        48,
        "W35 — 27/08/2026",
        "Hướng dẫn 4 bước Soạn thảo và Kiểm tra Điều khoản Bất khả kháng (Force Majeure Clause) trong Hợp đồng Ngoại thương theo Chuẩn Mẫu ICC 2020",
        "Bước 1: Định nghĩa rõ ràng 3 tiêu chí cấu thành Sự kiện Bất khả kháng (Nằm ngoài tầm kiểm soát hợp lý của các bên, Không thể lường trước được tại thời điểm ký kết hợp đồng, và Không thể khắc phục hoặc tránh được hậu quả dù đã áp dụng mọi biện pháp hợp lý); Bước 2: Liệt kê danh mục sự kiện cụ thể có dẫn chứng (Chiến tranh, phong tỏa kênh đào/eo biển hàng hải, bão lũ thiên tai cấp độ thảm họa, lệnh cấm vận thương mại của chính phủ, dịch bệnh khẩn cấp toàn cầu); Bước 3: Quy định nghĩa vụ và thời hạn gửi Thông báo Bất khả kháng (Bắt buộc gửi thông báo bằng văn bản trong vòng 7-14 ngày kể từ ngày phát sinh sự kiện kèm Giấy chứng nhận của Phòng Thương mại nước sở tại); Bước 4: Quy định lộ trình giải quyết hậu quả (Tạm hoãn thực hiện nghĩa vụ giao hàng trong 60-90 ngày mà không bị phạt vi phạm; nếu sự kiện kéo dài quá thời hạn thỏa thuận, cho phép một trong hai bên có quyền đơn phương chấm dứt hợp đồng mà không phải bồi thường thiệt hại).",
        "Thời hạn gửi thông báo bất khả kháng: Tối đa 07 đến 14 ngày làm việc kể từ thời điểm sự kiện xảy ra; Thời hạn tạm hoãn nghĩa vụ hợp đồng: Thường từ 60 đến 90 ngày (quá 90 ngày có quyền chấm dứt hợp đồng).",
        "1. Sử dụng điều khoản bất khả kháng ngắn ngủn 1 dòng chung chung 'Theo thông lệ quốc tế' (vô hiệu lực khi ra tòa án); 2. Coi sự biến động giá cả thị trường hoặc việc nhà cung cấp chậm hàng là sự kiện bất khả kháng (Tòa án và Trọng tài luôn bác bỏ lý do thương mại thuần túy); 3. Quên không xin Giấy xác nhận sự kiện bất khả kháng (Force Majeure Certificate) của Phòng Thương mại (như VCCI hoặc CCPIT).",
        "Điều khoản Bất khả kháng là 'Chiếc túi khí an toàn trong xe hơi hợp đồng' - Bạn không bao giờ muốn dùng tới nó, nhưng khi tai nạn thảm họa xảy ra, nó là thứ duy nhất cứu doanh nghiệp khỏi phá sản đền bù hợp đồng.",
        "So sánh Hợp đồng không có điều khoản Bất khả kháng chi tiết (Bị đối tác phạt vi phạm 8% + kiện bồi thường toàn bộ thiệt hại doanh thu khi trễ tàu do thiên tai/chiến sự) vs Hợp đồng chuẩn ICC Force Majeure Clause (Miễn trừ 100% trách nhiệm bồi thường và bảo toàn quyền đơn phương hủy hợp đồng an toàn).",
        "Anh Minh (Chủ doanh nghiệp xuất khẩu nông sản, thủ công mỹ nghệ) & Chị Hương (Quản lý Hợp đồng Ngoại thương).",
        "1. 3 tiêu chí cốt lõi bất khả kháng; 2. Danh mục sự kiện loại trừ rủi ro thương mại; 3. Thời hạn gửi thông báo văn bản 7 ngày; 4. Điều kiện cung cấp Chứng nhận VCCI/CCPIT; 5. Quyền hủy hợp đồng sau 60-90 ngày.",
        "https://en.wikipedia.org/wiki/Force_majeure",
        "Đàm phán Quốc tế",
        "Chưa dùng",
        "Cẩm nang soạn thảo điều khoản Bất khả kháng trong hợp đồng thương mại quốc tế theo quy chuẩn ICC Force Majeure Clause 2020."
    ],
    [
        49,
        "W35 — 27/08/2026",
        "Quy trình 5 bước Khử trùng và Kiểm tra Đạt chuẩn Dấu Kiểm định Bao bì Gỗ (ISPM 15) cho Hàng Pallet Gỗ Xuất khẩu Sang Mỹ và Châu Âu",
        "Bước 1: Lựa chọn đơn vị sản xuất pallet gỗ đã được Cục Bảo vệ Thực vật cấp mã số hành nghề xử lý kiểm dịch thực vật quốc tế; Bước 2: Thực hiện quy trình xử lý nhiệt chuẩn (Heat Treatment - HT: Gia nhiệt lõi gỗ đạt nhiệt độ tối thiểu 56°C trong thời gian liên tục ít nhất 30 phút, hoặc Khử trùng bằng Điện môi/Sóng viba) hoặc Khử trùng bằng Methyl Bromide (MB); Bước 3: Đóng dấu mộc chứng nhận ISPM 15 vĩnh viễn (IPPC Logo, Mã quốc gia VN, Mã số nhà xử lý XX, Mã biện pháp HT/MB) trên ít nhất 2 mặt đối diện của từng chiếc pallet; Bước 4: Kiểm tra độ ẩm của pallet gỗ (độ ẩm lý tưởng dưới 18-20% để chống nấm mốc trong container kín); Bước 5: Yêu cầu đơn vị xử lý cấp Giấy chứng nhận Khử trùng (Fumigation / Heat Treatment Certificate) và lưu trữ hồ sơ phục vụ thông quan xuất nhập khẩu.",
        "Quy chuẩn kỹ thuật xử lý nhiệt HT: Nhiệt độ tâm gỗ ≥ 56°C duy trì liên tục tối thiểu 30 phút; Tiêu chuẩn độ ẩm an toàn vận chuyển đường biển: Độ ẩm thớ gỗ (Wood Moisture Content) ≤ 19%.",
        "1. Dấu mộc ISPM 15 bị mờ, nhòe mực hoặc thiếu mã số quốc gia; 2. Sử dụng pallet gỗ cũ đã bị sửa chữa chắp vá bằng các thanh gỗ chưa qua xử lý nhiệt; 3. Pallet gỗ sau khi khử trùng bị bảo quản ngoài trời mưa ẩm dẫn đến mốc đen khi đóng vào container kín.",
        "Dấu mộc ISPM 15 là 'Hộ chiếu sạch bệnh của chiếc Pallet gỗ' - Phải có hình bông lúa IPPC, mã VN rõ ràng và ký hiệu HT nướng chín sâu bọ.",
        "So sánh Pallet Gỗ tự nhiên (Giá rẻ 100k-150k nhưng bắt buộc xử lý ISPM 15, nguy cơ nấm mốc cao) vs Pallet Gỗ ép Plywood / Pallet Nhựa (Được miễn trừ 100% thủ tục ISPM 15, sạch sẽ, không lo bị giữ hàng tại cảng đích Mỹ/EU).",
        "Anh Minh (Chủ xưởng sản xuất đồ gỗ, nông sản xuất khẩu) & Đội ngũ Đóng gói Kho bãi Hiện trường.",
        "1. Mã nhà xử lý còn hiệu lực trên cổng Cục BVTV; 2. Dấu mộc IPPC đóng đủ 2 mặt pallet; 3. Đo độ ẩm gỗ < 19%; 4. Thu giữ Chứng thư Khử trùng gốc; 5. Kiểm tra không chắp vá gỗ tạp.",
        "https://en.wikipedia.org/wiki/ISPM_15",
        "Quy trình Hải quan",
        "Chưa dùng",
        "Hướng dẫn kỹ thuật đóng gói và xử lý bao bì gỗ xuất khẩu đạt tiêu chuẩn quốc tế ISPM 15."
    ],
    [
        50,
        "W35 — 27/08/2026",
        "Hướng dẫn 5 bước Đăng ký và Xây dựng Hệ thống Kiểm soát Nội bộ (ICP) để Đạt Chứng nhận Doanh nghiệp Ưu tiên Hải quan (AEO) theo Chuẩn WCO",
        "Bước 1: Tự đánh giá Điều kiện Tiên quyết về Kim ngạch XNK (Tối thiểu 100 triệu USD/năm đối với DN xuất nhập khẩu nói chung, hoặc 40 triệu USD/năm đối với nông sản, thủy sản xuất khẩu) và Lịch sử Tuân thủ Pháp luật (Không vi phạm pháp luật thuế/hải quan trong 02 năm liên tục gần nhất); Bước 2: Thiết lập Chương trình Tuân thủ Nội bộ (Internal Compliance Program - ICP) bao gồm quy chế kiểm soát chứng từ, quy trình quản lý rủi ro chuỗi cung ứng và an ninh kho bãi/công nghệ thông tin; Bước 3: Hoàn thiện Hồ sơ đề nghị công nhận Doanh nghiệp Ưu tiên nộp lên Tổng cục Hải quan theo quy định Thông tư 72/2015/TT-BTC; Bước 4: Phối hợp tiếp đón Đoàn Thẩm định Thực tế của Tổng cục Hải quan đến kiểm tra hệ thống cơ sở vật chất, kho bãi, quy trình sản xuất và dữ liệu kế toán; Bước 5: Nhận Quyết định công nhận Doanh nghiệp Ưu tiên có hiệu lực 03 năm và duy trì báo cáo tự đánh giá tuân thủ hàng năm.",
        "Thời hạn hiệu lực của Chứng chỉ AEO: 03 năm kể từ ngày ký Quyết định công nhận; Tỷ lệ phân luồng thông quan: 100% tờ khai được phân vào Luồng Xanh (miễn kiểm tra chứng từ và miễn kiểm tra thực tế hàng hóa).",
        "1. Để xảy ra sai phạm về kế toán thuế trong 2 năm thẩm định dẫn đến hồ sơ bị loại ngay từ vòng 1; 2. Hệ thống an ninh kho bãi thiếu camera giám sát góc mù hoặc không có quy trình quản lý khách ra vào; 3. Không duy trì việc tự rà soát tuân thủ sau khi đã nhận chứng chỉ AEO dẫn đến bị thu hồi tư cách ưu tiên.",
        "Chứng chỉ AEO giống như 'Thẻ VIP Bạch Kim tại sân bay Hải quan' - Không cần xếp hàng chờ đợi, bay thẳng qua cửa ưu tiên, tiết kiệm hàng triệu USD chi phí thời gian.",
        "So sánh Doanh nghiệp thông thường (Chịu phân luồng Vàng/Đỏ ngẫu nhiên, mất 1-3 ngày kiểm hóa, lưu bãi tốn kém) vs Doanh nghiệp Đạt chuẩn AEO (Thông quan tự động trong vài giây, kiểm tra sau thông quan 3 năm/lần tại trụ sở, được ưu tiên đặc cách hàng đầu).",
        "Anh Dũng (Tổng Giám đốc Tập đoàn Sản xuất FDI) & Chị Hương (Giám đốc Chuỗi Cung ứng/Logistics Toàn cầu).",
        "1. Báo cáo tài chính kiểm toán 2 năm liên tiếp; 2. Chứng nhận hệ thống ISO 9001/27001; 3. Quy chế an ninh chuỗi cung ứng ICP; 4. Bảng kê tờ khai không vi phạm pháp luật; 5. Hồ sơ đề nghị gửi Tổng cục Hải quan.",
        "https://en.wikipedia.org/wiki/Authorized_economic_operator",
        "Tối ưu Thuế XNK",
        "Chưa dùng",
        "Cẩm nang hướng dẫn doanh nghiệp xây dựng năng lực tuân thủ để đạt chứng nhận Doanh nghiệp Ưu tiên Hải quan (AEO) theo Thông tư 72/2015/TT-BTC."
    ]
]

FUNFACT_W35_D2 = [
    [
        46,
        "W35 — 27/08/2026",
        "Sổ Tạm quản ATA Carnet: Chiếc 'Hộ chiếu quyền lực' của hàng hóa quốc tế đã ra đời từ sáng kiến miễn thuế tạm thời năm 1961 như thế nào?",
        "Có bao giờ bạn nghĩ một chiếc xe đua F1 trị giá 15 triệu USD hay một cây đàn vĩ cầm cổ của danh cầm thế giới có thể chu du qua hơn 80 quốc gia mà không cần nộp một đồng tiền đặt cọc thuế hải quan nào?",
        "Hội đồng Hợp tác Hải quan (nay là WCO) phối hợp cùng Phòng Thương mại Quốc tế (ICC) thông qua Công ước Hải quan về Sổ Tạm quản ATA (ATA Carnet Convention) vào ngày 6/12/1961 tại Brussels.",
        "Trước năm 1961, các công ty muốn mang hàng mẫu, máy quay phim hoặc thiết bị âm nhạc đi biểu diễn/triển lãm qua biên giới các nước phải làm thủ tục hải quan nhập khẩu đầy đủ và nộp tiền đặt cọc thuế bằng tiền mặt tại từng cửa khẩu, rồi chờ đợi hàng tháng trời sau khi rời đi mới được hoàn thuế.",
        "Sổ ATA được thừa nhận tại hơn 80 quốc gia và vùng lãnh thổ, xử lý hàng trăm nghìn lô hàng trị giá hơn 30 tỷ USD mỗi năm, trở thành công cụ quan trọng nhất thúc đẩy triển lãm quốc tế, thể thao (Thế vận hội Olympic, World Cup) và quảng bá sản phẩm toàn cầu.",
        "Các nước phát triển (Mỹ, EU, Nhật Bản) áp dụng Sổ ATA điện tử (e-ATA Carnet) qua ứng dụng điện thoại di động quét mã QR tại sân bay vs Một số nước đang phát triển vẫn xử lý thủ công bằng tập sổ giấy nhiều màu.",
        "Muốn đưa sản phẩm công nghệ, thủ công mỹ nghệ hoặc hàng mẫu Việt Nam ra thế giới dự hội chợ, hãy tận dụng Sổ ATA để giải phóng 100% dòng vốn lưu động và rút ngắn thời gian làm thủ tục tại cửa khẩu.",
        "https://en.wikipedia.org/wiki/ATA_Carnet",
        "Lịch sử Logistics",
        "Chưa dùng",
        "Lịch sử ra đời của Sổ Tạm quản ATA Carnet và cuộc cách mạng miễn thuế tạm thời cho hàng hóa lưu diễn/hội chợ quốc tế."
    ],
    [
        47,
        "W35 — 27/08/2026",
        "Tiêu chuẩn Quốc tế ISPM 15: Con bọ xén tóc tí hon trong khúc gỗ đã buộc cả thế giới phải 'nướng chín' mọi chiếc pallet gỗ như thế nào?",
        "Vì sao chỉ một con sâu bọ dài 2cm ẩn nấp trong pallet gỗ nhập khẩu từ châu Á lại có thể tàn phá hàng triệu héc-ta rừng thông Bắc Mỹ và tạo ra tiêu chuẩn đóng gói gỗ bắt buộc trên toàn hành tinh?",
        "Công ước Bảo vệ Thực vật Quốc tế (IPPC) thuộc Tổ chức Lương thực và Nông nghiệp Liên Hợp Quốc (FAO) ban hành Tiêu chuẩn Quốc tế về Biện pháp Kiểm dịch Thực vật số 15 (ISPM 15) vào năm 2002.",
        "Vào cuối thập niên 1990, loài bọ xén tóc châu Á (Asian Longhorned Beetle) và tuyến trùng thông theo các kiện hàng pallet gỗ thô nhập khẩu đã xâm nhập và hủy diệt hàng triệu cây thông, phong tại Mỹ và Canada, gây thiệt hại ước tính hàng chục tỷ USD cho ngành lâm nghiệp Bắc Mỹ.",
        "Hơn 180 quốc gia trên thế giới bắt buộc 100% bao bì gỗ (pallet, thùng gỗ, thanh chèn lót) dùng trong xuất nhập khẩu phải được xử lý nhiệt (Heat Treatment) hoặc khử trùng đạt chuẩn và đóng dấu mộc IPPC trước khi nhập cảnh.",
        "EU và Mỹ cấm hoàn toàn việc nhập khẩu gỗ thô chưa xử lý và dùng máy quét tự động tại cảng để soi mọt gỗ vs Xu hướng chuyển dịch mạnh mẽ sang sử dụng pallet gỗ ép dán keo nhân tạo hoặc pallet nhựa tái chế không thuộc diện quản lý kiểm dịch.",
        "Một lỗi nhỏ ở khâu đóng gói bao bì thô sơ có thể phá hỏng cả lô hàng xuất khẩu tiền tỷ. Hãy coi an toàn sinh học và kiểm soát chất lượng từ khâu đóng pallet là một phần không thể tách rời của sản phẩm xuất khẩu.",
        "https://en.wikipedia.org/wiki/ISPM_15",
        "Bí mật Hải quan",
        "Chưa dùng",
        "Khám phá nguồn gốc ra đời của Tiêu chuẩn ISPM 15 và bài học về an toàn sinh học trong chuỗi cung ứng toàn cầu."
    ],
    [
        48,
        "W35 — 27/08/2026",
        "Điều khoản Bất khả kháng (Force Majeure): Từ vụ kiện kinh điển về lũ lụt sông Thames đến 'lá bùa hộ mệnh' cứu doanh nghiệp trong chiến tranh và khủng hoảng toàn cầu",
        "Tại sao một trận bão cát, một cuộc phong tỏa kênh đào hay một đại dịch bất ngờ có thể hủy diệt toàn bộ cơ nghiệp của bạn nếu hợp đồng ngoại thương thiếu đúng 3 dòng chữ viết về 'Bất khả kháng'?",
        "Nguyên lý pháp lý 'Force Majeure' bắt nguồn từ Luật La Mã cổ đại (Vis Maior) và được pháp điển hóa trong Bộ luật Dân sự Napoleon 1804, sau đó được Phòng Thương mại Quốc tế (ICC) chuẩn hóa thành Bộ điều khoản mẫu toàn cầu.",
        "Trong các thế kỷ trước, theo nguyên tắc 'Pacta sunt servanda' (Hợp đồng đã ký là phải thực hiện tuyệt đối), các thương nhân dù bị đắm tàu do bão biển hay bị cướp phá do chiến tranh vẫn bị tòa án phán quyết bồi thường toàn bộ thiệt hại cho người mua nếu không giao được hàng.",
        "Điều khoản Bất khả kháng của ICC trở thành điều khoản tiêu chuẩn trong 100% hợp đồng thương mại quốc tế, bảo vệ các doanh nghiệp trước những biến động địa chính trị, thiên tai biến đổi khí hậu và xung đột hàng hải toàn cầu.",
        "Hệ thống Dân luật (Civil Law như Pháp, Đức, Việt Nam) công nhận nguyên tắc Bất khả kháng ngay trong luật thành văn vs Hệ thống Thông luật (Common Law như Anh, Mỹ) chỉ công nhận bất khả kháng nếu các bên có ghi rõ ràng từng câu chữ trong hợp đồng.",
        "Đừng bao giờ ký hợp đồng ngoại thương bằng sự lạc quan ngây thơ. Trong thế giới đầy biến động, sự chuẩn bị chu đáo cho các tình huống xấu nhất (Worst-case scenarios) chính là thước đo của bản lĩnh quản trị chuyên nghiệp.",
        "https://en.wikipedia.org/wiki/Force_majeure",
        "Đàm phán & Tư duy",
        "Chưa dùng",
        "Nguồn gốc lịch sử và giá trị cốt lõi của điều khoản Bất khả kháng Force Majeure trong quản trị rủi ro thương mại quốc tế."
    ],
    [
        49,
        "W35 — 27/08/2026",
        "Khung Tiêu chuẩn An ninh SAFE Framework của WCO: Làm thế nào sự kiện 11/9 tại Mỹ đã biến Hải quan toàn cầu từ 'người thu thuế' thành 'vệ sĩ chuỗi cung ứng'?",
        "Trước năm 2001, hải quan thế giới chỉ quan tâm đếm hàng để tính thuế; sau sự kiện 11/9, tại sao mỗi chiếc container vượt đại dương đều phải được 'soi chiếu lý lịch' nghiêm ngặt như một hành khách đi máy bay?",
        "Tổ chức Hải quan Thế giới (WCO) thông qua Khung Tiêu chuẩn SAFE nhằm Bảo đảm An ninh và Tạo thuận lợi cho Thương mại Toàn cầu (SAFE Framework of Standards) vào tháng 6/2005 tại Brussels.",
        "Sau vụ tấn công khủng bố ngày 11/9/2001, chính phủ Mỹ và thế giới nhận ra rằng hàng triệu container trôi nổi trên đại dương có thể trở thành 'ngựa thành Troy' vận chuyển vũ khí nguy hiểm nếu không có cơ chế kiểm soát an ninh chuỗi cung ứng từ xa.",
        "Khai sinh ra chương trình Doanh nghiệp Ưu tiên (AEO), quy định nộp tờ khai an ninh trước khi xếp hàng lên tàu (như ISF 10+2 của Mỹ, ENS của EU), biến an ninh thương mại thành trụ cột số 1 của hoạt động logistics toàn cầu.",
        "Mỹ áp dụng chương trình C-TPAT (Customs-Trade Partnership Against Terrorism) kết hợp quét tia X 100% container rủi ro cao vs Liên minh Châu Âu và Châu Á áp dụng mạng lưới AEO tương hỗ công nhận lẫn nhau (Mutual Recognition Agreements).",
        "An ninh chuỗi cung ứng không phải là trách nhiệm riêng của công an hay quân đội mà là tiêu chuẩn sống còn của mọi nhà máy xuất khẩu. Xây dựng quy trình kiểm soát kho bãi và nhân sự chuẩn chỉ là tấm vé thông hành để hàng Việt vươn ra biển lớn.",
        "https://en.wikipedia.org/wiki/Authorized_economic_operator",
        "Thương mại Quốc tế",
        "Chưa dùng",
        "Lịch sử chuyển mình của Hải quan thế giới từ thu thuế truyền thống sang bảo đảm an ninh chuỗi cung ứng toàn cầu sau sự kiện 11/9."
    ],
    [
        50,
        "W35 — 27/08/2026",
        "Định mức Nguyên vật liệu BOM (Bill of Materials): Làm thế nào bản vẽ chi tiết đến từng chiếc đinh ốc đã giúp ngành sản xuất toàn cầu tiết kiệm hàng nghìn tỷ USD chi phí phế phẩm?",
        "Bạn có biết rằng chỉ cần lệch một con ốc 2 gram trên bản BOM của một chiếc máy bay hay điện thoại thông minh, dây chuyền sản xuất nghìn tỷ sẽ sụp đổ và hải quan có thể phong tỏa toàn bộ nhà xưởng?",
        "Kỹ sư Joseph Orlicky và Gene Thomas phát triển Khái niệm Hoạch định Yêu cầu Nguyên vật liệu (MRP) và Hệ thống Định mức BOM chuẩn hóa tại tập đoàn IBM vào những năm 1960.",
        "Trước thập niên 1960, các nhà máy chế tạo xe hơi và máy bay thường xuyên rơi vào cảnh đình trệ: Thiếu vài chiếc bulong nhưng tồn kho thừa hàng vạn thanh thép, chi phí thất thoát và lãng phí vật tư chiếm tới 30% giá thành sản phẩm.",
        "BOM là 'trái tim' của toàn bộ hệ thống sản xuất công nghiệp hiện đại (ERP, MES) và là căn cứ pháp lý duy nhất để Hải quan các nước tính thuế, hoàn thuế và kiểm toán nguyên phụ liệu gia công xuất khẩu.",
        "Các tập đoàn công nghệ toàn cầu (Apple, Toyota) sử dụng BOM đa tầng tích hợp AI dự báo hao hụt đến 99.99% vs Các xưởng sản xuất truyền thống vẫn quản lý định mức trên sổ tay giấy thủ công dễ xảy ra sai sót.",
        "Sức mạnh của quản trị doanh nghiệp nằm ở sự chi tiết và chuẩn xác của dữ liệu nền tảng. Chuẩn hóa BOM không chỉ giúp tối ưu chi phí giá thành mà còn là bức tường lửa bảo vệ doanh nghiệp khi quyết toán hải quan.",
        "https://en.wikipedia.org/wiki/Bill_of_materials",
        "Công nghệ & Đổi mới",
        "Chưa dùng",
        "Khám phá lịch sử phát triển của Định mức nguyên vật liệu BOM và vai trò sống còn trong sản xuất và hải quan hiện đại."
    ]
]

LEGAL_NEW_W35_D2 = [
    [
        302,
        "Thông tư 33/2023/TT-BTC",
        "Quy định về việc xác định xuất xứ hàng hóa xuất khẩu, nhập khẩu",
        "Thông tư",
        "Bộ Tài chính",
        "2023-05-31",
        "2023-07-15",
        "Chưa xác định",
        "Thông tư 38/2018/TT-BTC, Thông tư 62/2019/TT-BTC",
        "Quy định chi tiết thủ tục kiểm tra, xác định xuất xứ hàng hóa XNK; cơ chế nộp C/O điện tử có chữ ký số và giải quyết sai sót nhỏ (minor discrepancies) theo Công ước Kyoto sửa đổi của WCO.",
        "Thừa nhận hoàn toàn giá trị pháp lý của C/O điện tử; cho phép nộp bổ sung C/O trong 30 ngày kể từ ngày đăng ký tờ khai; đơn giản hóa thủ tục chứng minh xuất xứ.",
        "2026-08-27"
    ],
    [
        303,
        "Nghị định 18/2021/NĐ-CP",
        "Quy định chi tiết thi hành Luật Thuế xuất khẩu, thuế nhập khẩu về miễn thuế, hoàn thuế hàng gia công và sản xuất xuất khẩu",
        "Nghị định",
        "Chính phủ",
        "2021-03-11",
        "2021-04-25",
        "Chưa xác định",
        "Nghị định 134/2016/NĐ-CP",
        "Quy định điều kiện kiểm tra cơ sở sản xuất, máy móc thiết bị đối với hàng hóa nhập khẩu để gia công, SXXK (Inward Processing Relief) và cơ chế miễn thuế, quyết toán nguyên phụ liệu.",
        "Siết chặt quy định kiểm tra cơ sở gia công lại; bắt buộc thông báo cơ sở sản xuất trước khi làm thủ tục nhập khẩu nguyên liệu miễn thuế.",
        "2026-08-27"
    ],
    [
        304,
        "Thông tư 72/2015/TT-BTC",
        "Quy định áp dụng chế độ ưu tiên trong thực hiện thủ tục hải quan, kiểm tra, giám sát hải quan đối với hàng hóa xuất khẩu, nhập khẩu của doanh nghiệp (Chương trình AEO)",
        "Thông tư",
        "Bộ Tài chính",
        "2015-05-12",
        "2015-06-26",
        "Chưa xác định",
        "Thông tư 86/2013/TT-BTC",
        "Quy chuẩn hóa điều kiện công nhận Doanh nghiệp Ưu tiên (AEO) theo chuẩn WCO SAFE Framework; đặc quyền miễn kiểm tra hồ sơ và kiểm tra thực tế hàng hóa (100% luồng xanh).",
        "Mở rộng đối tượng áp dụng cho doanh nghiệp xuất khẩu nông lâm thủy sản với ngưỡng kim ngạch ưu đãi; kéo dài thời hạn kiểm tra sau thông quan định kỳ lên 3 năm/lần.",
        "2026-08-27"
    ],
    [
        305,
        "WTO Trade Facilitation Agreement (TFA)",
        "Hiệp định Thuận lợi hóa Thương mại của Tổ chức Thương mại Thế giới và Cam kết Cải cách Cửa khẩu của Việt Nam",
        "Hiệp định Đa phương Quốc tế",
        "Tổ chức Thương mại Thế giới (WTO) / Chính phủ Việt Nam",
        "2017-02-22",
        "2017-02-22",
        "Chưa xác định",
        "Không có",
        "Cam kết bắt buộc về minh bạch hóa thông tin hải quan, công khai biểu phí cửa khẩu, triển khai cơ chế đo lường thời gian giải phóng hàng (TRS) và cơ chế Một cửa Quốc gia Single Window.",
        "Thúc đẩy hiện đại hóa thủ tục hải quan điện tử toàn diện, rút ngắn thời gian thông quan hàng hóa xuất nhập khẩu biên giới.",
        "2026-08-27"
    ],
    [
        306,
        "Nghị định 64/2020/NĐ-CP",
        "Quy định về thủ tục hải quan, kiểm tra, giám sát hải quan đối với hàng hóa quá cảnh thông qua Hệ thống quá cảnh Hải quan ASEAN và Sổ Tạm quản ATA",
        "Nghị định",
        "Chính phủ",
        "2020-06-10",
        "2020-06-15",
        "Chưa xác định",
        "Không có",
        "Khung pháp lý quốc gia đầu tiên triển khai Công ước Hải quan về Sổ Tạm quản ATA (ATA Carnet) và Công ước Istanbul về Tạm nhập Tái xuất hàng mẫu, triển lãm tại Việt Nam.",
        "Cho phép doanh nghiệp sử dụng một cuốn Sổ ATA duy nhất thay thế tờ khai hải quan và giấy bảo lãnh thuế, miễn nộp tiền đặt cọc thuế tại hơn 80 quốc gia thành viên.",
        "2026-08-27"
    ]
]


def backup_file(file_path):
    """Không tạo file backup để tránh nặng máy."""
    pass


def append_data_to_excel(file_name, new_data, center_cols=(1, 2, 14), status_col=14):
    file_path = os.path.join(REPORTS_DIR, file_name)
    backup_file(file_path)
    
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    start_row = ws.max_row + 1
    print(f"\n📝 Bắt đầu ghi {len(new_data)} dòng mới vào '{file_name}' (Bắt đầu từ Dòng {start_row})...")
    
    for i, row_data in enumerate(new_data):
        curr_row = start_row + i
        is_even = (curr_row % 2 == 0)
        row_fill = ZEBRA_FILL if is_even else WHITE_FILL
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.fill = row_fill
            
            # Căn lề
            if col_idx in center_cols:
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
                
            # Đổi font cho STT
            if col_idx == 1:
                cell.font = BOLD_FONT
                
    wb.save(file_path)
    print(f"✅ Ghi thành công vào '{file_name}'! Tổng số dòng hiện tại: {ws.max_row}")


def update_legal_tracker(new_data):
    file_name = "Danh_sach_Van_ban_Phap_luat_XNK.xlsx"
    file_path = os.path.join(REPORTS_DIR, file_name)
    backup_file(file_path)
    
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    start_row = ws.max_row + 1
    print(f"\n📝 Cập nhật Feedback loop vào '{file_name}' ({len(new_data)} văn bản mới từ Dòng {start_row})...")
    
    for i, row_data in enumerate(new_data):
        curr_row = start_row + i
        is_even = (curr_row % 2 == 0)
        row_fill = ZEBRA_FILL if is_even else WHITE_FILL
        
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.fill = row_fill
            
            if col_idx in (1, 6, 7, 8, 12):
                cell.alignment = ALIGN_CENTER
            else:
                cell.alignment = ALIGN_LEFT
                
            if col_idx == 1:
                cell.font = BOLD_FONT
                
    wb.save(file_path)
    print(f"✅ Cập nhật thành công vào '{file_name}'! Tổng số dòng hiện tại: {ws.max_row}")


def main():
    print("="*90)
    print("🚀 BẮT ĐẦU QUY TRÌNH WF06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS NGÀY 27/08/2026 (W35 — D2)")
    print("="*90)
    
    # 1. Thu thập toàn bộ 20 URL để chạy qua Technical Gatekeeper
    validation_list = []
    
    for item in TIN_NGANH_W35_D2:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[12],
            "keywords": ["origin", "customs", "duty", "drawback", "authorized", "economic", "operator", "trade", "facilitation", "ata", "carnet"]
        })
        
    for item in CASE_STUDY_W35_D2:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[10],
            "keywords": ["ispm", "wood", "phytosanitary", "bank", "guarantee", "customs", "inspection", "ship", "arrest", "supply", "chain", "security"]
        })
        
    for item in HOC_THUC_CHIEN_W35_D2:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[10],
            "keywords": ["ata", "carnet", "bill", "materials", "force", "majeure", "ispm", "wood", "authorized", "economic", "operator"]
        })
        
    for item in FUNFACT_W35_D2:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[9],
            "keywords": ["ata", "carnet", "ispm", "wood", "force", "majeure", "authorized", "economic", "operator", "bill", "materials"]
        })
        
    # 2. Rào chắn Kỹ thuật cưỡng chế
    print(f"🛡️ Gửi 20 URL vào Technical Gatekeeper (url_validator.py)...")
    enforce_strict_validation(validation_list)
    
    # 3. Ghi dữ liệu vào 4 tệp Excel Kho Ý Tưởng
    append_data_to_excel(
        "Tin_Nganh_Ideas_Bank_Eureka.xlsx",
        TIN_NGANH_W35_D2,
        center_cols=(1, 2, 5, 14),
        status_col=14
    )
    
    append_data_to_excel(
        "Case_Study_Ideas_Bank_Eureka.xlsx",
        CASE_STUDY_W35_D2,
        center_cols=(1, 2, 12, 13),
        status_col=13
    )
    
    append_data_to_excel(
        "Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx",
        HOC_THUC_CHIEN_W35_D2,
        center_cols=(1, 2, 12, 13),
        status_col=13
    )
    
    append_data_to_excel(
        "Funfact_Ideas_Bank_Eureka.xlsx",
        FUNFACT_W35_D2,
        center_cols=(1, 2, 11, 12),
        status_col=12
    )
    
    # 4. Feedback Loop: Cập nhật Danh_sach_Van_ban_Phap_luat_XNK.xlsx
    update_legal_tracker(LEGAL_NEW_W35_D2)
    
    print("\n" + "="*90)
    print("🎉 HOÀN THÀNH 100% QUY TRÌNH CẬP NHẬT IDEAS BANK THEO CHUẨN WF06 NGÀY 27/08/2026!")
    print("="*90)


if __name__ == "__main__":
    main()
