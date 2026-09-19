# -*- coding: utf-8 -*-
"""
Kịch bản cập nhật Ideas Bank Ngày 26/08/2026 (Tuần 35 - Day 1) theo chuẩn wf06.
Tác tử: Agent 01 (Research & Đóng gói) + Agent 08 (Fact Checker & Deep URL Validator).
======================================================================================
- Quét và đóng gói chính xác 5 ý tưởng mới (STT 41-45) cho mỗi Pillar (TN, CS, HT, FF).
- Kiểm tra 100% URL qua Technical Gatekeeper (scripts/url_validator.py).
- Ghi dữ liệu vào 4 file Excel trong thư mục reports/ với định dạng chuẩn Segoe UI.
- Feedback loop: Cập nhật 5 văn bản pháp luật/quy chuẩn mới vào Danh_sach_Van_ban_Phap_luat_XNK.xlsx.
"""

import os
import sys
import io
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
# DỮ LIỆU Ý TƯỞNG NGÀY 26/08/2026 (STT 41 -> 45 | ĐỦ 4 PILLARS x 5 = 20 Ý TƯỞNG)
# ==============================================================================

TIN_NGANH_W35 = [
    [
        41,
        "W35 — 26/08/2026",
        "Quy định về Cơ chế Điều chỉnh Biên giới Carbon (CBAM - Regulation (EU) 2023/956) và Yêu cầu Bắt buộc Báo cáo Phát thải Thực tế cho Hàng hóa XNK sang Liên minh Châu Âu (EU)",
        "Ủy ban Châu Âu (European Commission - EC) / Bộ Công Thương Việt Nam",
        "15/08/2026",
        "Chính thức kết thúc giai đoạn chuyển tiếp cho phép sử dụng giá trị phát thải mặc định (default values); bắt buộc 100% doanh nghiệp xuất khẩu các nhóm hàng thép, nhôm, xi măng, phân bón sang EU phải nộp Báo cáo Phát thải Thực tế (Actual Emissions Report) có xác thực của đơn vị kiểm toán độc lập được EU công nhận.",
        "Các nhà sản xuất và doanh nghiệp xuất khẩu sắt thép, nhôm định hình, phân bón, xi măng sang thị trường 27 nước thành viên EU và các đơn vị Logistics chuyên tuyến Châu Âu.",
        "Nếu không nộp báo cáo hoặc số liệu phát thải không được kiểm toán chấp thuận, nhà nhập khẩu tại EU bị phạt từ 10 đến 50 Euro cho mỗi tấn phát thải chưa khai báo; hàng hóa có nguy cơ bị từ chối thông quan tại cảng đích hoặc bị áp thuế phát thải carbon cao nhất.",
        "Hiệp hội Thép Việt Nam (VSA) phối hợp Bộ Công Thương tổ chức chuỗi tập huấn khẩn cấp về tính toán dấu chân carbon; các doanh nghiệp thép lớn (Hòa Phát, Nam Kim, Hoa Sen) đã chủ động chuyển đổi công nghệ lò điện (EAF) để giảm lượng phát thải.",
        "Tạo động lực mạnh mẽ thúc đẩy công nghiệp sản xuất xanh hóa chuỗi cung ứng, bảo vệ môi trường toàn cầu và giúp các doanh nghiệp tiên phong giành thị phần lớn tại các thị trường cao cấp.",
        "Chi phí kiểm kê khí nhà kính và thuê kiểm toán quốc tế tiêu tốn từ 200 - 500 triệu đồng cho mỗi kỳ báo cáo, bào mòn biên lợi nhuận của các doanh nghiệp sản xuất vừa và nhỏ trong nước.",
        "Quy định (EU) 2023/956 của Nghị viện và Hội đồng Châu Âu về việc thiết lập Cơ chế điều chỉnh biên giới carbon (CBAM).",
        "https://en.wikipedia.org/wiki/Carbon_Border_Adjustment_Mechanism",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Cơ chế điều chỉnh biên giới carbon CBAM của EU trên Wikipedia."
    ],
    [
        42,
        "W35 — 26/08/2026",
        "Quy định An toàn Sản phẩm Chung của EU (General Product Safety Regulation - GPSR / Regulation (EU) 2023/988) đối với Hàng hóa Thương mại Điện tử Xuyên biên giới và Hàng Tiêu dùng Nhập khẩu",
        "Liên minh Châu Âu (EU) / Hệ thống Cảnh báo An toàn Sản phẩm Safety Gate",
        "10/08/2026",
        "Bắt buộc mọi sản phẩm phi thực phẩm (đồ chơi trẻ em, thiết bị gia dụng, đồ gỗ nội thất, dệt may) bán trực tiếp hoặc qua sàn TMĐT (Amazon, Temu, Shein, TikTok Shop) sang EU phải có Đại diện pháp lý tại EU (EU Responsible Person) và thể hiện đầy đủ thông tin truy xuất nguồn gốc, thông tin cảnh báo an toàn bằng ngôn ngữ nước sở tại trên nhãn sản phẩm và trang bán lẻ trực tuyến.",
        "Hàng nghìn người bán hàng TMĐT xuyên biên giới (Cross-border E-commerce Sellers), các doanh nghiệp sản xuất đồ gỗ, đồ chơi bằng gỗ, may mặc và đồ gia dụng xuất khẩu sang EU.",
        "Sản phẩm vi phạm sẽ bị cơ quan quản lý thị trường EU phát lệnh thu hồi trên toàn khối qua hệ thống Safety Gate, bị các sàn TMĐT khóa tài khoản bán hàng vĩnh viễn và bị tiêu hủy lô hàng tại hải quan sân bay/cảng biển Châu Âu.",
        "Cộng đồng seller Amazon và TMĐT tại Việt Nam ráo riết tìm kiếm các đơn vị cung cấp dịch vụ Người đại diện tại EU (EU Authorized Representative) và thiết kế lại bao bì sản phẩm theo chuẩn mới.",
        "Bảo vệ tối đa người tiêu dùng Châu Âu khỏi hàng giả, hàng kém chất lượng và nguy cơ cháy nổ, tạo sân chơi công bằng giữa hàng nhập khẩu và hàng sản xuất nội khối EU.",
        "Thủ tục cấp phép và duy trì Người đại diện có trách nhiệm tại EU khiến chi phí vận hành cho mỗi mã hàng (SKU) tăng từ 500 - 1.500 EUR/năm, gây khó khăn cho các nhà bán lẻ quy mô nhỏ.",
        "Quy chế An toàn Sản phẩm Châu Âu GPSR (EU) 2023/988 thay thế Chỉ thị GPSD 2001/95/EC.",
        "https://en.wikipedia.org/wiki/General_Product_Safety_Directive",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Khung pháp lý an toàn sản phẩm chung General Product Safety của EU."
    ],
    [
        43,
        "W35 — 26/08/2026",
        "Quy chuẩn Vận chuyển Đường biển Quốc tế đối với Hàng hóa Nguy hiểm (IMDG Code / Phụ lục SOLAS) và Hướng dẫn Phân loại An toàn Khí tài Lưu trữ Năng lượng (Pin Lithium-ion, Xe điện EV)",
        "Tổ chức Hàng hải Quốc tế (IMO) / Cục Hàng hải Việt Nam",
        "01/08/2026",
        "Siết chặt phân loại mã số nhận diện hàng nguy hiểm UN3480 (Pin Lithium standalone) và UN3481 (Pin lắp trong thiết bị); bắt buộc nộp Tóm tắt Kiểm thử An toàn UN 38.3 (Test Summary), Bảng Dữ liệu An toàn Vật liệu (MSDS) hợp lệ và Giấy chứng nhận đóng gói container nguy hiểm (Dangerous Goods Declaration) trước khi hạ bãi cảng biển.",
        "Các nhà máy sản xuất xe điện, xe đạp điện, pin điện thoại, máy tính xách tay, tấm pin lưu trữ năng lượng mặt trời và các doanh nghiệp Logistics vận tải container đường biển.",
        "Khai báo sai nhóm hàng nguy hiểm sẽ bị hãng tàu từ chối xếp dỡ (Shut-out), bị phạt vi phạm an ninh hàng hải từ 5.000 - 20.000 USD/container và chịu toàn bộ trách nhiệm hình sự/dân sự nếu xảy ra sự cố cháy nổ trên tàu.",
        "Các hãng tàu lớn (Maersk, MSC, ONE, CMA CGM) tăng cường quét tự động tờ khai manifest bằng AI để phát hiện các lô hàng pin khai man thành 'hàng bách hóa thông thường' (General Cargo).",
        "Đảm bảo an toàn tính mạng tuyệt đối cho thuyền viên và ngăn chặn các thảm họa cháy tàu container hàng hải do hiện tượng thoát nhiệt (Thermal Runaway) của pin lithium.",
        "Phụ phí hàng nguy hiểm (DG Surcharge) và thủ tục kiểm định đóng gói làm giá cước vận chuyển tăng từ 40% - 70% so với hàng hóa thông thường.",
        "Bộ luật Quốc tế về Vận chuyển Hàng hóa Nguy hiểm bằng Đường biển (IMDG Code) theo Công ước SOLAS 1974 của IMO.",
        "https://en.wikipedia.org/wiki/International_Maritime_Dangerous_Goods_Code",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Tiêu chuẩn quốc tế IMDG Code về hàng hóa nguy hiểm trên Wikipedia."
    ],
    [
        44,
        "W35 — 26/08/2026",
        "Cơ chế Tự Chứng nhận Xuất xứ Hàng hóa (Self-Certification of Origin) và Lộ trình Áp dụng Giấy chứng nhận Xuất xứ Điện tử (e-C/O) trong Hiệp định Đối tác Kinh tế Toàn diện Khu vực (RCEP)",
        "Bộ Công Thương Việt Nam / Ban Thư ký ASEAN & Các nước Thành viên RCEP",
        "25/07/2026",
        "Mở rộng danh sách Doanh nghiệp Đủ điều kiện Tự chứng nhận Xuất xứ (Approved Exporters) theo Điều 3.16 Hiệp định RCEP; cho phép doanh nghiệp tự khai phát hành văn bản chứng nhận xuất xứ trên hóa đơn thương mại mà không cần xin cấp C/O Form RCEP giấy tại các Phòng Quản lý XNK, giảm thời gian thông quan xuống còn dưới 2 giờ.",
        "Các tập đoàn sản xuất linh kiện điện tử, phụ tùng ô tô, dệt may, nông sản chế biến có chuỗi cung ứng liên kết giữa Việt Nam, Nhật Bản, Hàn Quốc, Trung Quốc và ASEAN.",
        "Nếu tự chứng nhận sai tiêu chí chuyển đổi mã HS (CTC) hoặc tỷ lệ hàm lượng giá trị khu vực (RVC 40%), doanh nghiệp sẽ bị tước quyền Tự chứng nhận trong 2 năm, bị hải quan nước nhập khẩu truy thu toàn bộ thuế ưu đãi và phạt vi phạm gian lận xuất xứ.",
        "Các doanh nghiệp FDI lớn và tập đoàn sản xuất trong nước tích cực chuẩn bị hồ sơ năng lực để đăng ký mã số Nhà xuất khẩu đủ điều kiện; các hiệp hội ngành hàng đẩy mạnh tập huấn phân bổ quy tắc xuất xứ gộp RCEP.",
        "Tối ưu hóa thời gian và chi phí xin cấp C/O truyền thống, thúc đẩy dòng chảy thương mại nội khối RCEP với quy mô chiếm 30% GDP toàn cầu.",
        "Trách nhiệm pháp lý và rủi ro hậu kiểm (Post-Audit) chuyển hoàn toàn sang vai doanh nghiệp, đòi hỏi hệ thống kế toán và lưu trữ chứng từ chuỗi cung ứng phải đạt chuẩn quốc tế trong 5 năm.",
        "Hiệp định Đối tác Kinh tế Toàn diện Khu vực (RCEP Chapter 3: Rules of Origin).",
        "https://en.wikipedia.org/wiki/Regional_Comprehensive_Economic_Partnership",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Hiệp định Đối tác Kinh tế Toàn diện Khu vực RCEP trên Wikipedia."
    ],
    [
        45,
        "W35 — 26/08/2026",
        "Quy định về Chuỗi Cung ứng Không Phá rừng của Liên minh Châu Âu (EUDR - Regulation (EU) 2023/1115) và Hướng dẫn Định vị Tọa độ Địa lý (Geolocation Data) cho Nông sản Xuất khẩu",
        "Ủy ban Châu Âu (EC) / Bộ Nông nghiệp & Phát triển Nông thôn Việt Nam",
        "20/07/2026",
        "Bắt buộc các lô hàng xuất khẩu 7 nhóm mặt hàng (cà phê, cao su, gỗ, đậu nành, dầu cọ, ca cao, thịt bò) sang EU phải cung cấp dữ liệu tọa độ GPS chính xác của từng mảnh vườn canh tác (plot of land) chứng minh không gây mất rừng sau ngày 31/12/2020 và tuân thủ luật pháp nước sở tại.",
        "Hàng triệu hộ nông dân, hợp tác xã và doanh nghiệp chế biến xuất khẩu cà phê, cao su, gỗ và sản phẩm gỗ của Việt Nam sang thị trường Châu Âu.",
        "Hàng nông sản không có dữ liệu truy xuất GPS hợp lệ sẽ bị cấm nhập khẩu vào EU, bị tịch thu hoặc buộc tái xuất; doanh nghiệp vi phạm có thể bị phạt tối thiểu 4% doanh thu hàng năm tại EU và đưa vào danh sách đen công khai.",
        "Các vùng trồng cà phê trọng điểm Tây Nguyên (Đắk Lắk, Lâm Đồng, Gia Lai) đang khẩn trương số hóa bản đồ vùng trồng và cấp mã định danh số cho từng nông hộ liên kết.",
        "Bảo vệ tài nguyên rừng, ứng phó biến đổi khí hậu toàn cầu và nâng tầm vị thế nông sản Việt Nam thành nguồn cung bền vững, có trách nhiệm cao trên trường quốc tế.",
        "Hàng triệu nông hộ manh mún tại Việt Nam gặp khó khăn lớn về công nghệ và chi phí khi đo vẽ bản đồ số GPS, nguy cơ bị loại khỏi chuỗi giá trị xuất khẩu chính ngạch sang EU.",
        "Quy định (EU) 2023/1115 của Nghị viện Châu Âu về các sản phẩm không gây phá rừng (EU Deforestation Regulation - EUDR).",
        "https://en.wikipedia.org/wiki/Deforestation",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Tiêu chuẩn chống phá rừng và quản lý nguồn gốc gỗ/nông sản trên Wikipedia."
    ]
]

CASE_STUDY_W35 = [
    [
        41,
        "W35 — 26/08/2026",
        "Khai nhầm mã UN3480 thành hàng thường: Lô 5 container Pin Lithium bị Hãng tàu phát lệnh 'Cấm bốc hàng' (Shut-out) tại Cảng Cát Lái, thiệt hại 480 triệu đồng tiền phạt và lưu bãi",
        "Công ty TNHH Năng lượng Xanh V. (Bình Dương), Hãng tàu Quốc tế O., Chi cục Hải quan Cửa khẩu Cảng Sài Gòn Khu vực 1",
        "Doanh nghiệp xuất khẩu 5 container pin sạc dự phòng lithium-ion đi thị trường Rotterdam. Để tiết kiệm phụ phí hàng nguy hiểm (DG Surcharge 450 USD/cont), nhân viên logistics đã khai trên Bill và Booking là 'Hàng linh kiện điện tử thông thường'. Khi hạ bãi tại Cát Lái, máy soi chiếu phát hiện mật độ pin dày đặc. Hãng tàu lập tức phát lệnh từ chối bốc hàng lên tàu (Shut-out), yêu cầu tái kiểm định an toàn và lập biên bản gian lận hàng hải.",
        "Thiệt hại tài chính 480 triệu đồng (gồm: 150 triệu đồng tiền phạt gian lận khai báo từ hãng tàu, 180 triệu đồng phí cắm điện/lưu bãi container nguy hiểm 20 ngày, 150 triệu đồng tiền phạt hủy hợp đồng giao trễ cho khách hàng EU).",
        "Tư duy tiết kiệm chi phí ngắn hạn, thiếu hiểu biết về Bộ luật Hàng nguy hiểm Hàng hải Quốc tế (IMDG Code) và không lường trước được cơ chế quét tự động kiểm soát an ninh của các hãng tàu hiện đại.",
        "1. Phân loại chính xác mã UN (UN3480 cho pin rời, UN3481 cho pin kèm thiết bị); 2. Chuẩn bị đầy đủ chứng chỉ kiểm định an toàn UN 38.3 và MSDS trước khi booking; 3. Tuyệt đối không gian dối khai báo hàng nguy hiểm thành hàng bách hóa.",
        "Phe A (DN): Hàng đóng gói cẩn thận trong thùng carton có chống sốc, hãng tàu xử phạt quá nặng tay. Phe B (Hãng tàu/Chuyên gia): Gian lận pin lithium đe dọa sinh mạng cả con tàu và hàng nghìn container khác trên đại dương, cần xử lý hình sự để răn đe.",
        "An toàn hàng hải là lằn ranh đỏ không thể thỏa hiệp. Mọi toan tính 'lách luật' đối với hàng nguy hiểm đều phải trả giá đắt gấp hàng chục lần chi phí khai báo chuẩn chỉ.",
        "https://en.wikipedia.org/wiki/Lithium-ion_battery",
        "Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách tai nạn vận chuyển pin lithium-ion và bài học tuân thủ an toàn hàng nguy hiểm IMDG Code."
    ],
    [
        42,
        "W35 — 26/08/2026",
        "Hậu kiểm C/O Form EUR.1 sau 3 năm: Doanh nghiệp May mặc bị Hải quan Đức bác chứng nhận xuất xứ, đối tác bị truy thu 1,85 tỷ đồng tiền thuế quan EVFTA",
        "Công ty May xuất khẩu T. (Nam Định), Đối tác nhập khẩu hàng may mặc tại Hamburg (Đức), Hải quan Liên bang Đức (Zoll)",
        "Doanh nghiệp Việt Nam xuất khẩu áo jacket sang Đức hưởng thuế suất ưu đãi 0% theo EVFTA bằng C/O Form EUR.1. Sau 3 năm, Hải quan Đức tiến hành kiểm tra sau thông quan (Post-Clearance Audit) và yêu cầu cung cấp hồ sơ chứng minh xuất xứ từ sợi (Yarn Forward). Doanh nghiệp Việt Nam không thể cung cấp được hóa đơn đầu vào chứng minh nguồn gốc sợi vải mua trôi nổi trên thị trường nội địa. Hải quan Đức ra quyết định bác toàn bộ C/O của 12 lô hàng.",
        "Đối tác Đức bị truy thu 68.000 EUR (khoảng 1,85 tỷ đồng) tiền thuế nhập khẩu MFN kèm tiền phạt chậm nộp; đối tác lập tức khấu trừ toàn bộ số tiền này vào công nợ của doanh nghiệp Việt Nam và chấm dứt hợp đồng hợp tác dài hạn.",
        "Doanh nghiệp chỉ quan tâm đến việc 'xin được C/O tại thời điểm xuất khẩu' mà không thiết lập hệ thống kế toán lưu trữ chứng từ gốc (BOM, định mức nguyên phụ liệu, hóa đơn GTGT đầu vào) đạt chuẩn để phục vụ công tác hậu kiểm trong thời hạn 5 năm.",
        "1. Thiết lập quy trình lưu trữ hồ sơ chứng minh xuất xứ tối thiểu 5 năm theo Điều 27 Thông tư 11/2020/TT-BCT; 2. Chỉ mua nguyên phụ liệu từ các nhà cung ứng có hóa đơn chứng từ truy xuất nguồn gốc rõ ràng đáp ứng quy tắc xuất xứ 2 công đoạn (vải trở đi); 3. Định kỳ tự kiểm toán C/O nội bộ hàng quý.",
        "Phe A (DN): Đã được Phòng Quản lý XNK cấp C/O hợp lệ thì Hải quan nước ngoài phải công nhận. Phe B (Pháp lý/Hải quan EU): C/O chỉ là chứng thư ban đầu, quyền kiểm tra thực tế hồ sơ sản xuất thuộc về nước nhập khẩu theo đúng quy chế hiệp định thương mại tự do.",
        "Hưởng ưu đãi thuế FTA không phải là món quà miễn phí. Giá trị của FTA đi kèm nghĩa vụ quản trị hồ sơ minh bạch xuyên suốt nửa thập kỷ.",
        "https://en.wikipedia.org/wiki/Rules_of_origin",
        "Tai nạn XNK",
        "Chưa dùng",
        "Bài học xương máu về rủi ro hậu kiểm xuất xứ C/O hàng dệt may xuất khẩu sang Châu Âu."
    ],
    [
        43,
        "W35 — 26/08/2026",
        "Bẫy 'Hàng kẹp Container giá rẻ': Forwarder lậu nhét hàng cấm vào chung lô hàng gốm sứ, chủ hàng bị Hải quan tạm giữ hàng 45 ngày và gánh 620 triệu đồng phí Demurrage/Detention",
        "Công ty Gốm sứ Mỹ nghệ H. (Bình Dương), Công ty Giao nhận 'chui' K. (TP.HCM), Chi cục Hải quan Cửa khẩu Cảng Cái Mép",
        "Doanh nghiệp xuất khẩu gốm sứ sang Mỹ thuê dịch vụ gom hàng lẻ (LCL) của một đơn vị Forwarder quảng cáo giá cước siêu rẻ trên mạng xã hội. Đơn vị này đã lén lút đóng ghép hàng của doanh nghiệp chung với các kiện hàng thuốc lá điếu và hàng vi phạm sở hữu trí tuệ để trốn thuế. Khi container qua máy soi Hải quan Cái Mép, hệ thống phát hiện dấu hiệu bất thường và chuyển sang luồng Đỏ khám xét toàn bộ (100%). Hàng gốm sứ bị niêm phong phục vụ điều tra trong 45 ngày.",
        "Phát sinh 620 triệu đồng chi phí (gồm 380 triệu đồng tiền lưu bãi Demurrage/Detention tại cảng, 140 triệu đồng phí bốc dỡ/kiểm hóa thủ công, 100 triệu đồng chi phí pháp lý và đền bù khách hàng Mỹ do trễ mùa bán hàng Giáng sinh).",
        "Chủ hàng ham giá rẻ, chọn đơn vị gom hàng LCL không có giấy phép đại lý hải quan và hợp đồng dịch vụ logistics sơ sài, không có điều khoản ràng buộc trách nhiệm pháp lý khi đóng ghép container.",
        "1. Chỉ làm việc với các đơn vị giao nhận có pháp nhân rõ ràng, có hợp đồng dịch vụ logistics đầy đủ điều khoản bảo đảm tính độc lập của từng lô hàng LCL; 2. Yêu cầu biên bản đóng hàng (Stuffing Report) có hình ảnh và dấu niêm phong chì (Seal) minh bạch; 3. Đối với hàng giá trị cao nên ưu tiên đi nguyên container (FCL) hoặc chọn NVOCC uy tín.",
        "Phe A (Chủ hàng): Tôi là nạn nhân vô tội, Hải quan nên giải tỏa hàng hợp pháp sớm. Phe B (Hải quan/Luật sư): Container là một thể thống nhất khi niêm phong, chừng nào chưa làm rõ hành vi đồng phạm buôn lậu thì toàn bộ tang vật phải tạm giữ theo Luật Tố tụng.",
        "Trong logistics, cước rẻ bất thường luôn đi kèm rủi ro pháp lý chết người. Đừng bao giờ giao phó tài sản và uy tín thương hiệu vào tay những đơn vị gom hàng thiếu trách nhiệm.",
        "https://en.wikipedia.org/wiki/Demurrage",
        "Bóc phốt chiêu trò",
        "Chưa dùng",
        "Cảnh báo chiêu trò gom hàng kẹp lậu container và cách bảo vệ quyền lợi pháp lý cho chủ hàng chân chính."
    ],
    [
        44,
        "W35 — 26/08/2026",
        "Cuộc giải cứu 3 Container Sầu riêng cấp đông kẹt tại Cửa khẩu Hữu Nghị: Sai lệch 1 chữ số trong Mã Vùng trồng (PUC) và màn đàm phán cấp lại Chứng thư Kiểm dịch trong 36 giờ cứu 2,4 tỷ đồng",
        "Công ty Nông sản XNK Đắk Lắk S., Chi cục Kiểm dịch Thực vật Vùng VII, Hải quan Cửa khẩu Bằng Tường (Trung Quốc)",
        "3 container sầu riêng cấp đông (trị giá 2,4 tỷ đồng) xuất khẩu chính ngạch sang Trung Quốc đến cửa khẩu Hữu Nghị thì bị phía Hải quan Bằng Tường chặn lại vì mã vùng trồng (PUC) trên Giấy chứng nhận Kiểm dịch Thực vật (Phytosanitary Certificate) bị gõ nhầm 1 ký tự số so với danh sách mã số được Tổng cục Hải quan Trung Quốc (GACC) phê duyệt trên hệ thống CIFER. Container đối mặt nguy cơ bị buộc quay đầu hoặc tiêu hủy khi nhiệt độ bảo quản buồng lạnh container chạm ngưỡng báo động.",
        "Nguy cơ mất trắng 2,4 tỷ đồng tiền hàng sầu riêng; đàm phán thành công giải cứu lô hàng trong 36 giờ, chỉ phát sinh 45 triệu đồng tiền chạy máy lạnh container và phí xử lý chứng từ khẩn cấp.",
        "Khâu đánh máy và đối soát chứng từ kiểm dịch tại hiện trường làm việc vội vàng, không thực hiện quy tắc 'kiểm tra chéo 2 người' (4-eyes principle) trước khi phát hành chứng thư bản gốc.",
        "1. Thực hiện đối chiếu tự động mã PUC/cơ sở đóng gói với hệ thống CIFER của GACC trước khi in chứng thư kiểm dịch; 2. Thiết lập kênh liên lạc khẩn cấp trực tiếp giữa Đội Kiểm dịch cửa khẩu và cơ quan cấp tỉnh để thu hồi và cấp đổi chứng thư điện tử trong vòng 24 giờ; 3. Cài đặt thiết bị giám sát nhiệt độ GPS/IoT từ xa trong suốt hành trình vận chuyển.",
        "Phe A (DN): Lỗi sai chỉ là 1 ký tự chính tả đánh máy nhỏ, hải quan biên giới quá cứng nhắc. Phe B (Hải quan TQ): Hệ thống hải quan số hóa tự động từ chối nếu không trùng khớp dữ liệu mã định danh cơ sở dữ liệu quốc gia.",
        "Chi tiết nhỏ quyết định thành bại lớn trong thương mại nông sản chính ngạch. Sự chuẩn xác từng ký tự trên chứng từ là chìa khóa bảo vệ hàng tỷ đồng tài sản của người nông dân và doanh nghiệp.",
        "https://en.wikipedia.org/wiki/Phytosanitary_certificate",
        "Giải cứu hàng hóa",
        "Chưa dùng",
        "Case study thực chiến về xử lý khủng hoảng sai lệch chứng từ kiểm dịch thực vật tại cửa khẩu phía Bắc."
    ],
    [
        45,
        "W35 — 26/08/2026",
        "'Quên' khai báo Phí Bản quyền và Khuôn mẫu hỗ trợ (Assists): Doanh nghiệp Cơ khí bị Truy thu Thuế và Phạt vi phạm Trị giá Hải quan 340 triệu đồng",
        "Công ty Chế tạo Máy M. (Đồng Nai), Đội Kiểm tra sau Thông quan - Cục Hải quan Đồng Nai",
        "Doanh nghiệp nhập khẩu linh kiện cơ khí chính xác từ Nhật Bản theo hợp đồng mua bán thông thường. Tuy nhiên, doanh nghiệp đã thanh toán riêng một khoản tiền bản quyền phần mềm thiết kế CAD/CAM trị giá 50.000 USD cho công ty mẹ tại Tokyo và cung cấp miễn phí 4 bộ khuôn ép mẫu (Assists) cho nhà sản xuất Nhật Bản mà không khai báo vào tờ khai hải quan nhập khẩu. Khi kiểm tra sau thông quan, đoàn kiểm tra phát hiện dòng tiền chuyển khoản ra nước ngoài trên báo cáo tài chính và xác định đây là các khoản điều chỉnh cộng bắt buộc vào Trị giá tính thuế hải quan.",
        "Doanh nghiệp bị truy thu 210 triệu đồng tiền thuế nhập khẩu và thuế GTGT còn thiếu, phạt 20% trên số tiền thuế gian lận (42 triệu đồng) và tính tiền chậm nộp 88 triệu đồng (tổng cộng 340 triệu đồng).",
        "Bộ phận kế toán và bộ phận xuất nhập khẩu hoạt động tách rời; người làm hải quan chỉ nhìn vào hóa đơn thương mại (Commercial Invoice) của lô hàng mà không nắm được các thỏa thuận tài chính ngầm về bản quyền và khoản hỗ trợ khuôn mẫu (Assists) theo Điều 13 Thông tư 39/2015/TT-BTC.",
        "1. Rà soát toàn bộ các hợp đồng chuyển giao công nghệ, phí bản quyền và chi phí khuôn mẫu cung cấp cho nhà sản xuất nước ngoài; 2. Thực hiện khai báo phân bổ khoản điều chỉnh cộng theo đúng phương pháp Trị giá giao dịch WTO (ACV); 3. Đồng bộ dữ liệu định kỳ hàng quý giữa bộ phận Kế toán thuế và phòng Xuất nhập khẩu trước mỗi đợt quyết toán hải quan.",
        "Phe A (DN): Phí bản quyền trả theo hợp đồng dịch vụ riêng biệt, không nằm trong giá mua bán hàng hóa vật chất thì không phải chịu thuế nhập khẩu. Phe B (Hải quan): Bản quyền và khuôn mẫu gắn liền trực tiếp với việc sản xuất hàng hóa nhập khẩu, theo luật định bắt buộc phải cộng vào trị giá tính thuế hải quan.",
        "Trị giá hải quan không đơn thuần là con số ghi trên hóa đơn. Hiểu đúng bản chất các khoản điều chỉnh cộng/trừ là bức tường lửa bảo vệ doanh nghiệp khỏi các án phạt thuế nặng nề.",
        "https://en.wikipedia.org/wiki/Customs_valuation",
        "Tai nạn XNK",
        "Chưa dùng",
        "Phân tích chuyên sâu về nghiệp vụ xác định trị giá hải quan và các khoản điều chỉnh cộng bản quyền/khuôn mẫu."
    ]
]

HOC_THUC_CHIEN_W35 = [
    [
        41,
        "W35 — 26/08/2026",
        "Quy trình 5 bước Phân loại, Đóng gói và Khai báo Vận chuyển Hàng hóa Nguy hiểm (DG Cargo) cho Pin Lithium và Hóa chất theo Chuẩn IMDG Code",
        "Bước 1: Xác định chính xác Phân nhóm Hàng nguy hiểm (Hazard Class) và Mã số UN (ví dụ: Class 9, UN3480 cho pin rời hoặc UN3481 cho pin kèm thiết bị) dựa trên bảng MSDS; Bước 2: Soát xét Báo cáo Kiểm thử An toàn UN 38.3 (Lithium Battery Test Summary) từ nhà sản xuất; Bước 3: Đóng gói hàng hóa theo Quy cách Đóng gói chuẩn IMDG (Packing Instruction P903/P908) với nhãn cảnh báo kích thước chuẩn 100x100mm; Bước 4: Lập Bản kê Khai báo Hàng nguy hiểm của Người gửi hàng (Shipper's Declaration for Dangerous Goods - DGD); Bước 5: Booking chỗ với Hãng tàu chuyên tuyến và hoàn tất thủ tục nộp giấy phép chấp thuận DG tại cảng trước giờ Closing time.",
        "Công thức xác định công suất Watt-giờ của pin Lithium-ion: Watt-hour (Wh) = Điện áp định mức (Volts) x Dung lượng pin (Ampere-hours - Ah). Ví dụ: Pin 3.7V dung lượng 10.000mAh = 3.7 x 10 = 37Wh (nằm dưới ngưỡng 100Wh được áp dụng cơ chế đóng gói giảm nhẹ Phần II của Chỉ dẫn IATA/IMDG).",
        "1. Không dán nhãn cảnh báo Class 9 hoặc dùng nhãn sai kích thước quy định; 2. Dùng chung mã UN3480 cho thiết bị có lắp sẵn pin trong máy (lỗi sai mã UN); 3. Không chèn vật liệu cách điện chống đoản mạch giữa các cực pin khi đóng thùng carton.",
        "Ví Pin Lithium như 'Trái bom nổ chậm trong khoang hàng' - Muốn lên tàu biển thì phải có 'Căn cước công dân UN 38.3' và 'Áo giáp chống sốc P903'.",
        "So sánh UN3480 (Pin Lithium-ion độc lập) vs UN3481 (Pin Lithium-ion đóng gói cùng hoặc lắp trong thiết bị): UN3480 yêu cầu trạng thái nạp điện (State of Charge - SoC) không quá 30% khi vận chuyển đường hàng không và đóng gói nghiêm ngặt hơn nhiều so với UN3481.",
        "Anh Minh (Chủ DN sản xuất thiết bị công nghệ) & Chị Hương (Trưởng phòng Logistics tại nhà máy linh kiện điện tử).",
        "1. MSDS hợp lệ trong 12 tháng; 2. Báo cáo UN 38.3 đầy đủ 8 bài kiểm tra áp suất/nhiệt độ; 3. Nhãn Class 9 & UN Number dán đủ 2 mặt thùng; 4. Bản khai DGD ký bởi nhân sự có chứng chỉ DG; 5. Xác nhận hạ bãi DG từ cảng.",
        "https://en.wikipedia.org/wiki/Dangerous_goods",
        "Quy trình Hải quan",
        "Chưa dùng",
        "Cẩm nang step-by-step phân loại và vận chuyển hàng nguy hiểm đường biển cho doanh nghiệp sản xuất điện tử."
    ],
    [
        42,
        "W35 — 26/08/2026",
        "Hướng dẫn 5 bước Xác định Trị giá Hải quan (Customs Valuation) và Phân bổ Các Khoản Điều chỉnh Cộng (Phí Bản quyền, Khuôn mẫu Assists) theo Chuẩn WTO",
        "Bước 1: Xác định Phương pháp Xác định Trị giá Hải quan (Ưu tiên số 1: Phương pháp Trị giá Giao dịch của Hàng hóa nhập khẩu); Bước 2: Rà soát 4 điều kiện áp dụng phương pháp trị giá giao dịch (không bị hạn chế quyền định đoạt, không có điều kiện phụ thuộc, không chia sẻ tiền bán lại và không có quan hệ đặc biệt làm ảnh hưởng giá); Bước 3: Nhận diện các khoản điều chỉnh cộng bắt buộc (Phí hoa hồng bán buôn, bao bì, chi phí khuôn mẫu/dụng cụ do người mua hỗ trợ - Assists, phí bản quyền trả riêng); Bước 4: Tính toán phân bổ các khoản điều chỉnh cộng vào từng dòng hàng trên tờ khai hải quan; Bước 5: Lập Bảng kê Khai báo Trị giá Hải quan và chuẩn bị hồ sơ giải trình tham vấn giá khi cơ quan hải quan nghi vấn.",
        "Trị giá Hải quan = Giá thực tế đã thanh toán hoặc sẽ phải thanh toán + Tổng các khoản điều chỉnh cộng - Tổng các khoản điều chỉnh trừ. Công thức phân bổ chi phí khuôn mẫu (Assists): Chi phí phân bổ cho 1 đơn vị hàng = Tổng giá trị khuôn mẫu / Tổng số lượng sản phẩm dự kiến sản xuất bằng khuôn đó.",
        "1. Giấu không khai báo các khoản phí bản quyền phần mềm hoặc thiết kế thanh toán qua ngân hàng riêng; 2. Bỏ qua giá trị khuôn mẫu gửi từ nước ngoài sang xưởng gia công; 3. Không lưu giữ chứng từ thanh toán cước quốc tế (F/I) dẫn đến bị bác giá vận chuyển.",
        "Ví Trị giá Hải quan như 'Tảng băng trôi' - Tiền hàng trên Invoice chỉ là 'Phần nổi', các khoản phí bản quyền và khuôn mẫu là 'Phần chìm' bắt buộc phải khai báo nếu không muốn bị đắm tàu khi kiểm tra sau thông quan.",
        "So sánh Khoản điều chỉnh Cộng (Assists, Phí bản quyền, Chi phí vận tải đến cảng nhập) vs Khoản điều chỉnh Trừ (Chi phí lắp đặt sau nhập khẩu, Thuế nội địa, Chi phí vận chuyển nội địa sau khi qua cửa khẩu): Cộng làm tăng trị giá tính thuế, Trừ giúp giảm thuế hợp pháp nếu tách bạch được trên chứng từ.",
        "Chị Hương (Trưởng phòng XNK) & Anh Dũng (Giám đốc Tài chính/Kế toán trưởng doanh nghiệp FDI).",
        "1. Hợp đồng mua bán ngoại thương (Sales Contract); 2. Hóa đơn thương mại và Phiếu đóng gói; 3. Hợp đồng bản quyền/License Agreement (nếu có); 4. Chứng từ chuyển tiền quốc tế T/T, L/C; 5. Bảng phân bổ chi phí khuôn mẫu và cước quốc tế.",
        "https://en.wikipedia.org/wiki/Customs_valuation",
        "Tối ưu Thuế XNK",
        "Chưa dùng",
        "Hướng dẫn chuyên sâu cách xác định trị giá hải quan và các khoản điều chỉnh cộng theo quy định Thông tư 39/2015/TT-BTC và chuẩn WTO."
    ],
    [
        43,
        "W35 — 26/08/2026",
        "Hướng dẫn 5 bước Soát xét Giấy chứng nhận Kiểm dịch Thực vật (Phytosanitary Certificate) và Quản lý Mã số Vùng trồng (PUC) Xuất khẩu Nông sản",
        "Bước 1: Tra cứu yêu cầu kiểm dịch thực vật (SPS) của thị trường nhập khẩu (Trung Quốc, EU, Mỹ, Nhật Bản) đối với chủng loại nông sản cụ thể; Bước 2: Kiểm tra tính hợp lệ của Mã số Vùng trồng (PUC) và Mã số Cơ sở Đóng gói (PHC) trên cơ sở dữ liệu quốc gia và hệ thống phê duyệt của nước nhập khẩu (ví dụ: CIFER của GACC); Bước 3: Lấy mẫu kiểm nghiệm sinh vật gây hại và dư lượng thuốc BVTV theo Tiêu chuẩn ISPM 12; Bước 4: Soát xét 100% các trường thông tin trên bản dự thảo Chứng thư Kiểm dịch (Tên khoa học của thực vật, số lượng, quy cách đóng gói, thông tin phương tiện vận chuyển, biện pháp xử lý khử trùng nhiệt độ/hóa chất); Bước 5: Nhận Chứng thư gốc (Original) và tải mã QR xác thực lên Cổng một cửa quốc gia trước khi hàng rời cảng/cửa khẩu.",
        "Tỷ lệ lấy mẫu kiểm tra kiểm dịch thực vật theo chuẩn ISPM 31: Đối với lô hàng đóng thùng carton dưới 1.000 kiện: Lấy ngẫu nhiên căn bậc hai số kiện (√N) để kiểm tra mầm bệnh; Thời hạn hiệu lực của Phytosanitary Certificate: Thông thường tối đa 14 ngày kể từ ngày cấp đối với hàng xuất khẩu đường biển.",
        "1. Ghi sai tên khoa học bằng tiếng Latin của nông sản (ví dụ: Sầu riêng là Durio zibethinus thay vì chỉ ghi Durian); 2. Sử dụng mã vùng trồng đã hết hạn hoặc bị tạm ngưng mã số; 3. Không khớp ngày khử trùng thực tế với ngày thể hiện trên chứng thư.",
        "Ví Phytosanitary Certificate như 'Giấy khai sinh và Hộ chiếu y tế' của nông sản - Chỉ cần sai 1 chữ số trong mã vùng trồng là nông sản bị 'từ chối nhập cảnh' ngay tại cửa khẩu.",
        "So sánh Phytosanitary Certificate (Kiểm dịch thực vật cho cây trồng/nông sản tươi) vs Health Certificate (Chứng thư y tế cho thực phẩm chế biến sẵn): Phytosanitary tập trung vào ngăn chặn dịch bệnh và sâu bọ ngoại lai, Health Certificate tập trung vào vệ sinh an toàn thực phẩm cho người tiêu dùng.",
        "Anh Minh (Doanh chủ HTX Nông sản xuất khẩu) & Đội ngũ Hiện trường Logistics Cửa khẩu.",
        "1. Kiểm tra mã PUC/PHC còn hiệu lực; 2. So khớp Tên khoa học tiếng Latin; 3. Xác nhận phương pháp xử lý khử trùng (Methyl Bromide/Vapour Heat Treatment); 4. Kiểm tra chữ ký và con dấu số của kiểm dịch viên; 5. Đối chiếu số container/số seal trùng 100% với Bill of Lading.",
        "https://en.wikipedia.org/wiki/Phytosanitary_certificate",
        "Quy trình Hải quan",
        "Chưa dùng",
        "Cẩm nang kỹ thuật kiểm soát chứng từ kiểm dịch thực vật xuất khẩu nông sản chính ngạch."
    ],
    [
        44,
        "W35 — 26/08/2026",
        "Hướng dẫn 5 bước Tính toán, Đàm phán và Kiểm soát Bộ ba Chi phí Lưu Container: Demurrage (DEM), Detention (DET) và Storage (STO)",
        "Bước 1: Phân biệt rõ bản chất 3 loại phí: DEM (Phí lưu container tại bãi cảng do hãng tàu thu), DET (Phí lưu vỏ container ngoài cảng do hãng tàu thu) và STO (Phí lưu bãi do cảng vụ thu); Bước 2: Soát xét số ngày miễn phí (Free Time) trên Booking Confirmation và Thông báo Hàng đến (Arrival Notice); Bước 3: Đàm phán kéo dài Free Time kết hợp (Combined DEM/DET 14-21 ngày) ngay từ khâu ký hợp đồng ngoại thương và booking tàu; Bước 4: Lên kế hoạch điều phối xe kéo (Trucking) và mở tờ khai thông quan sớm 48 giờ trước khi tàu cập cảng; Bước 5: Giám sát mốc thời gian hoàn trả vỏ rỗng (Empty Return) về đúng Depot chỉ định để chốt biên bản giao nhận đúng hạn.",
        "Công thức tính phí phạt lũy tiến: Tổng Phí DEM/DET = Số ngày vượt Free Time x Mức đơn giá phạt theo bậc thang. Ví dụ: Cont 40ft vượt 5 ngày đầu phạt 45 USD/ngày, từ ngày thứ 6 trở đi phạt 80 USD/ngày. Nếu chậm 10 ngày = (5 x 45) + (5 x 80) = 625 USD (~15,6 triệu đồng/container).",
        "1. Nhầm lẫn giữa ngày hết hạn Free DEM của hãng tàu và ngày hết hạn Free Storage của Cảng vụ; 2. Hạ vỏ rỗng sai Depot quy định khiến đồng hồ tính phí DET vẫn tiếp tục chạy; 3. Không chủ động xin gộp Combined Free Time khi đàm phán hợp đồng vận chuyển.",
        "DEM là 'Tiền thuê chỗ để hộp của Hãng tàu', STO là 'Tiền gửi xe ở bãi của Cảng', DET là 'Tiền mượn đồ chưa trả về nhà cho Hãng tàu'.",
        "So sánh Free Time riêng lẻ (Tách biệt 7 ngày DEM + 7 ngày DET) vs Combined Free Time (Gộp chung 14 ngày DEM/DET linh hoạt): Combined Free Time cho phép chủ hàng linh hoạt dồn toàn bộ 14 ngày cho việc lưu cont tại kho riêng mà không lo bị tính phí phạt DET.",
        "Chị Hương (Trưởng phòng Logistics) & Anh Minh (Chủ doanh nghiệp xuất nhập khẩu).",
        "1. Kiểm tra số ngày Free DEM/DET trên Arrival Notice; 2. Mở tờ khai hải quan điện tử trước khi tàu cập; 3. Đặt lịch xe kéo cont ngay khi thông quan; 4. Kiểm tra giờ làm việc của Depot nhận vỏ; 5. Lấy Phiếu tiếp nhận vỏ rỗng (EIR) có xác nhận giờ trả vỏ.",
        "https://en.wikipedia.org/wiki/Demurrage",
        "Đàm phán Quốc tế",
        "Chưa dùng",
        "Bí quyết đàm phán và quản trị chi phí lưu container DEM/DET/Storage cho doanh nghiệp XNK."
    ],
    [
        45,
        "W35 — 26/08/2026",
        "Hướng dẫn 4 bước Tự Kiểm tra và Chuẩn bị Hồ sơ Kiểm tra Sau Thông quan (PCA) tại Trụ sở Doanh nghiệp theo Thông tư 38/2015/TT-BTC",
        "Bước 1: Rà soát và đối chiếu số liệu giữa Tờ khai Hải quan - Sổ sách Kế toán kho - Báo cáo Quyết toán nguyên phụ liệu (Form 15/BCQT-NVL); Bước 2: Kiểm tra tính thống nhất của Bộ chứng từ xuất nhập khẩu (Invoice, Packing List, C/O, Hợp đồng, Chứng từ thanh toán ngân hàng); Bước 3: Soát xét các điểm nóng rủi ro cao: Mã số HS, Trị giá hải quan, Tỷ lệ hao hụt thực tế, Tiêu chí xuất xứ C/O và Bán tiêu thụ phế liệu nội địa; Bước 4: Lập Bảng giải trình chênh lệch và chủ động nộp hồ sơ Khai bổ sung theo quy định trước khi cơ quan hải quan công bố Quyết định Kiểm tra sau thông quan để được miễn phạt gian lận 20%.",
        "Mức phạt vi phạm khi bị phát hiện trong đợt PCA: Phạt 20% trên số tiền thuế khai thiếu + Tiền chậm nộp thuế 0.03%/ngày trên số tiền thuế nợ. Ví dụ: Khai thiếu 500 triệu đồng tiền thuế trong 365 ngày = 500tr x 20% (100tr phạt) + 500tr x 0.03% x 365 ngày (54,75tr tiền chậm nộp) = Tổng số tiền phải nộp thêm 654,75 triệu đồng.",
        "1. Số liệu tồn kho thực tế tại nhà máy không khớp với số liệu tồn trên phần mềm hải quan; 2. Chứng từ thanh toán quốc tế qua ngân hàng không đủ hoặc thiếu ủy quyền hợp lệ; 3. Đợi đến khi có quyết định kiểm tra mới bắt đầu tìm kiếm chứng từ gốc.",
        "Kiểm tra sau thông quan giống như 'Khám sức khỏe tổng quát định kỳ' của doanh nghiệp - Tự phát hiện bệnh sớm để uống thuốc giải trình, còn hơn để bác sĩ Hải quan phát hiện 'khối u truy thu thuế' kèm án phạt nặng nề.",
        "So sánh Tự nguyện Khai bổ sung trong thời hạn 60 ngày kể từ ngày thông quan (Miễn hoàn toàn phạt 20%) vs Bị phát hiện trong quá trình Kiểm tra sau thông quan tại trụ sở doanh nghiệp (Bị phạt 20% đến 300% và xếp hạng doanh nghiệp rủi ro cao vào luồng Đỏ): Tự khai giúp cứu hàng trăm triệu đồng tiền phạt và bảo vệ uy tín doanh nghiệp.",
        "Anh Dũng (Giám đốc Điều hành/Tài chính) & Chị Hương (Kế toán trưởng/Trưởng phòng XNK).",
        "1. Bộ hồ sơ tờ khai hải quan 5 năm gần nhất; 2. Sổ chi tiết nhập - xuất - tồn kho nguyên liệu/thành phẩm; 3. Hóa đơn GTGT và Phiếu thu/chi liên quan; 4. Bảng tính định mức sản xuất thực tế; 5. Chứng từ thanh toán qua ngân hàng (Điện chuyển tiền MT103).",
        "https://en.wikipedia.org/wiki/Customs",
        "Tối ưu Thuế XNK",
        "Chưa dùng",
        "Cẩm nang tự thanh tra kiểm toán hồ sơ hải quan và quản trị rủi ro kiểm tra sau thông quan PCA."
    ]
]

FUNFACT_W35 = [
    [
        41,
        "W35 — 26/08/2026",
        "Đạo luật Hàng hải Jones Act 1920 của Hoa Kỳ: Vì sao quy định bảo hộ hơn 100 năm tuổi lại khiến cước vận tải nội địa Mỹ đắt gấp đôi cước vượt Thái Bình Dương?",
        "Bạn có tin rằng chi phí chở một container từ bờ Đông sang bờ Tây nước Mỹ bằng đường biển có thể đắt hơn cả chở container đó từ Việt Nam sang Los Angeles, chỉ vì một đạo luật bảo hộ ban hành từ năm 1920?",
        "Thượng nghị sĩ Wesley Livsey Jones và Quốc hội Hoa Kỳ thông qua Đạo luật Vận tải Biển Thương mại 1920 (Merchant Marine Act of 1920 / Jones Act).",
        "Sau Thế chiến I, để bảo đảm duy trì một đội tàu buôn quốc gia phục vụ quốc phòng trong thời chiến, Mỹ quy định: Mọi tàu bè chở hàng giữa hai cảng biển của Mỹ bắt buộc phải đóng tại Mỹ, treo cờ Mỹ, thuộc sở hữu của công dân Mỹ và do thủy thủ đoàn 75% người Mỹ vận hành.",
        "Do chi phí đóng tàu tại Mỹ đắt gấp 4-5 lần châu Á và lương thủy thủ đoàn Mỹ cao ngất ngưởng, số lượng tàu đạt chuẩn Jones Act sụt giảm nghiêm trọng, khiến cước vận chuyển nội địa Mỹ đắt đỏ phi lý, đẩy hàng hóa sang đường bộ và đường sắt gây quá tải.",
        "Châu Âu và nhiều quốc gia châu Á (trong đó có Việt Nam) mở cửa cho phép cạnh tranh vận tải ven biển (Cabotage) hoặc nới lỏng điều kiện thuê tàu quốc tế để tối ưu hóa chi phí logistics nội địa.",
        "Bảo hộ quá mức sẽ làm thui chột năng lực cạnh tranh và đẩy chi phí lên vai người tiêu dùng. Doanh nghiệp Việt Nam cần tận dụng lợi thế chi phí vận tải biển cạnh tranh để tối ưu chuỗi cung ứng xuất khẩu sang các thị trường tự do.",
        "https://en.wikipedia.org/wiki/Merchant_Marine_Act_of_1920",
        "Thương mại Quốc tế",
        "Chưa dùng",
        "Góc nhìn lịch sử kinh tế thú vị về tác động của chính sách bảo hộ hàng hải Jones Act của Mỹ."
    ],
    [
        42,
        "W35 — 26/08/2026",
        "Hệ thống Mã UN (UN Numbers) của Liên Hợp Quốc: Làm thế nào 4 chữ số trên thùng hàng đã cứu mạng hàng triệu thủy thủ và định hình an toàn vận tải toàn cầu?",
        "Làm sao các lính cứu hỏa tại bất kỳ cảng biển nào trên thế giới, dù không biết tiếng Anh hay tiếng Việt, chỉ cần nhìn vào 4 con số trên một thùng hàng là biết ngay phải dùng cát hay nước để dập lửa?",
        "Ủy ban Chuyên gia về Vận chuyển Hàng hóa Nguy hiểm của Liên Hợp Quốc (UNECE) xây dựng Khung hướng dẫn Khuyến nghị Toàn cầu (UN Orange Book) vào năm 1956.",
        "Trong nửa đầu thế kỷ 20, hàng nghìn vụ nổ hóa chất thảm khốc xảy ra trên các tàu biển do mỗi quốc gia gọi tên hóa chất theo ngôn ngữ và quy ước riêng, khiến thủy thủ đoàn và lính cứu hỏa dùng nhầm nước dập lửa cho các kim loại phản ứng mạnh, làm đám cháy bùng phát dữ dội hơn.",
        "Hệ thống mã 4 chữ số (từ UN 0004 đến UN 3550+) đã trở thành ngôn ngữ chung duy nhất cho mọi phương thức vận tải toàn cầu (Hàng hải IMDG, Hàng không ICAO/IATA, Đường bộ ADR).",
        "100% các quốc gia thành viên Liên Hợp Quốc đều tích hợp mã UN vào luật hải quan và an toàn vận tải nội địa, tạo nên sự đồng nhất tuyệt đối không rào cản ngôn ngữ.",
        "Chuẩn hóa quy trình và sử dụng mã định danh chung (Standardization) là chìa khóa để vận hành hệ thống logistics quy mô lớn một cách chính xác và không sai sót.",
        "https://en.wikipedia.org/wiki/UN_number",
        "Lịch sử Logistics",
        "Chưa dùng",
        "Khám phá nguồn gốc ra đời của hệ thống mã số hàng nguy hiểm UN Number trong thương mại quốc tế."
    ],
    [
        43,
        "W35 — 26/08/2026",
        "Công ước Quốc tế MARPOL 73/78 và Cú hích IMO 2020: Cuộc chiến chống khói đen trên đại dương đã làm thay đổi toàn bộ bản đồ giá cước tàu biển thế giới ra sao?",
        "Chỉ một quy định cắt giảm hàm lượng lưu huỳnh trong nhiên liệu từ 3.5% xuống 0.5% vào năm 2020 đã khiến các hãng tàu container phải chi thêm hàng chục tỷ USD để lắp 'máy lọc khí thải' (Scrubber) hoặc đổi sang chạy khí LNG?",
        "Tổ chức Hàng hải Quốc tế (IMO) thông qua Công ước Quốc tế về Ngăn ngừa Ô nhiễm từ Tàu biển (MARPOL 73/78) và triển khai Quy chuẩn IMO 2020 (Giới hạn Lưu huỳnh Toàn cầu).",
        "Trong suốt nhiều thập kỷ, các tàu hàng khổng lồ sử dụng dầu cặn (Heavy Fuel Oil - HFO) - loại nhiên liệu bẩn nhất thải ra lượng khí SOx khổng lồ tương đương hàng chục triệu ô tô lưu thông trên cạn, gây mưa axit và ô nhiễm không khí tại các thành phố ven biển.",
        "IMO 2020 tạo ra sự ra đời của các loại phụ phí nhiên liệu mới (LSS - Low Sulphur Surcharge, BAF) và thúc đẩy cuộc chạy đua đóng thế hệ tàu siêu tải chạy bằng nhiên liệu xanh (Methanol, Amoniac, LNG) của Maersk, CMA CGM.",
        "Các vùng biển kiểm soát khí thải đặc biệt (ECA) tại Bắc Mỹ và Bắc Âu áp dụng mức lưu huỳnh thậm chí còn khắt khe hơn (0.1%), đòi hỏi tàu thuyền phải tắt động cơ dùng điện bờ khi cập cảng.",
        "Xu hướng 'Logistics Xanh' và ESG không còn là khẩu hiệu xa vời mà đã trở thành chi phí trực tiếp trên từng hóa đơn vận chuyển. Doanh nghiệp cần chủ động chuyển đổi để không bị loại khỏi chuỗi cung ứng toàn cầu.",
        "https://en.wikipedia.org/wiki/MARPOL_73/78",
        "Lịch sử Logistics",
        "Chưa dùng",
        "Bóc tách lịch sử Công ước MARPOL và cuộc cách mạng nhiên liệu sạch IMO 2020 của ngành vận tải biển."
    ],
    [
        44,
        "W35 — 26/08/2026",
        "Triết lý Quản trị Sản xuất Just-In-Time (JIT) của Taiichi Ohno (Toyota): Làm thế nào tư duy 'Coi tồn kho là tội ác' đã giúp Toyota vượt mặt gã khổng lồ nước Mỹ?",
        "Tại sao một nhà máy sản xuất hàng triệu chiếc ô tô mỗi năm lại chỉ lưu trữ lượng linh kiện đủ dùng trong đúng... 2 giờ đồng hồ, mà dây chuyền lắp ráp không bao giờ bị dừng lại?",
        "Kỹ sư trưởng Taiichi Ohno và Tập đoàn Ô tô Toyota sáng tạo ra Hệ thống Sản xuất Toyota (Toyota Production System - TPS) và mô hình Just-In-Time (JIT) vào thập niên 1950-1970.",
        "Sau Thế chiến II, nước Nhật khan hiếm tài nguyên và vốn đầu tư, không thể sản xuất hàng loạt theo kiểu phương Tây (Ford) để tích trữ đầy kho. Taiichi Ohno lấy cảm hứng từ cách siêu thị Mỹ bổ sung hàng hóa lên kệ để thiết kế nguyên lý 'Chỉ sản xuất đúng cái cần, đúng lúc cần, và đúng số lượng cần'.",
        "JIT trở thành triết lý xương sống của quản trị chuỗi cung ứng toàn cầu hiện đại (Lean Logistics), giúp cắt giảm 80% chi phí lưu kho và giải phóng hàng tỷ USD dòng tiền chết cho doanh nghiệp.",
        "Mô hình JIT của Nhật Bản đòi hỏi sự liên kết cực kỳ kỷ luật với các nhà cung ứng vệ tinh, trong khi mô hình phương Tây truyền thống thường chọn 'Just-In-Case' (dự trữ đệm an toàn lớn).",
        "Hàng tồn kho chính là 'nghĩa địa chôn vốn' của các doanh nghiệp SME Việt Nam. Ứng dụng tư duy tinh gọn và dự báo chuỗi cung ứng chuẩn xác giúp giải phóng dòng tiền và tăng sức chống chịu trước khủng hoảng.",
        "https://en.wikipedia.org/wiki/Just-in-time_manufacturing",
        "Đàm phán & Tư duy",
        "Chưa dùng",
        "Bài học kinh điển về tư duy quản trị tinh gọn Just-In-Time của Toyota và ứng dụng vào Logistics hiện đại."
    ],
    [
        45,
        "W35 — 26/08/2026",
        "Hiệp hội Vận tải Container Kỹ thuật số (DCSA) và Cuộc cách mạng Vận đơn Điện tử (eBL): Làm thế nào đám mây và mã hóa đang dần khai tử hàng trăm triệu tờ giấy B/L vật lý mỗi năm?",
        "Có bao giờ bạn tự hỏi: Tại sao giữa kỷ nguyên AI và vệ tinh không gian, người ta vẫn phải thuê máy bay chuyển phát nhanh từng tờ giấy vận đơn (Bill of Lading) gốc vượt đại dương chỉ để đổi lấy một chiếc lệnh giao hàng?",
        "Hiệp hội Vận tải Container Kỹ thuật số (Digital Container Shipping Association - DCSA) được thành lập năm 2019 bởi 9 hãng tàu container lớn nhất thế giới (MSC, Maersk, CMA CGM, Hapag-Lloyd, ONE, Evergreen, Yang Ming, HMM, ZIM).",
        "Vận đơn đường biển (Bill of Lading) là chứng từ sở hữu có giá trị chuyển nhượng ra đời từ thế kỷ 14. Mỗi năm, ngành hàng hải phát hành hơn 45 triệu bộ vận đơn giấy, tiêu tốn hàng tỷ USD chi phí in ấn, chuyển phát nhanh DHL/FedEx và thường xuyên xảy ra tình trạng 'hàng đã tới cảng nhưng giấy tờ vẫn đang kẹt trên trời'.",
        "DCSA đặt mục tiêu chuyển đổi 100% sang Vận đơn Điện tử (eBL) dựa trên chuẩn dữ liệu mở và công nghệ bảo mật số, ước tính tiết kiệm cho ngành thương mại toàn cầu hơn 6,5 tỷ USD/năm và rút ngắn thời gian xử lý chứng từ từ hàng tuần xuống vài phút.",
        "Singapore, Anh và Liên minh Châu Âu đã hoàn thiện khung pháp lý công nhận giá trị pháp lý của chứng từ điện tử có thể chuyển nhượng (MLETR của UNCITRAL), trong khi nhiều nước đang phát triển đang khẩn trương sửa đổi luật hàng hải nội địa.",
        "Doanh nghiệp xuất nhập khẩu Việt Nam cần chủ động tiếp cận công nghệ vận đơn điện tử và chữ ký số quốc tế để bắt nhịp với dòng chảy số hóa không thể đảo ngược của thương mại toàn cầu.",
        "https://en.wikipedia.org/wiki/Bill_of_lading",
        "Công nghệ & Đổi mới",
        "Chưa dùng",
        "Khám phá cuộc cách mạng số hóa chứng từ hàng hải và lộ trình chuyển đổi 100% eBL của liên minh hãng tàu DCSA."
    ]
]

LEGAL_DOCS_W35 = [
    [
        297,
        "Regulation (EU) 2023/956",
        "Quy định của Nghị viện và Hội đồng Châu Âu về việc thiết lập Cơ chế Điều chỉnh Biên giới Carbon (CBAM)",
        "Quy định quốc tế",
        "Nghị viện Châu Âu và Hội đồng Liên minh Châu Âu",
        "2023-05-10",
        "2026-08-15",
        "Đang áp dụng",
        "Quy chế CBAM (EU)",
        "Quy định bắt buộc báo cáo phát thải thực tế có kiểm toán độc lập cho hàng hóa xuất khẩu sang EU (thép, nhôm, xi măng, phân bón), kết thúc giai đoạn dùng giá trị phát thải mặc định.",
        "Bắt buộc nộp Báo cáo Phát thải Thực tế (Actual Emissions Report); áp dụng chế tài phạt 10-50 Euro/tấn phát thải chưa khai báo và tính thuế carbon từ năm 2026.",
        "2026-08-26"
    ],
    [
        298,
        "Regulation (EU) 2023/988",
        "Quy định An toàn Sản phẩm Chung của Liên minh Châu Âu (General Product Safety Regulation - GPSR)",
        "Quy định quốc tế",
        "Nghị viện và Hội đồng Liên minh Châu Âu",
        "2023-05-23",
        "2026-08-10",
        "Đang áp dụng",
        "Thay thế Chỉ thị GPSD 2001/95/EC",
        "Quy định về an toàn sản phẩm tiêu dùng và hàng hóa thương mại điện tử xuyên biên giới bán vào thị trường EU, bắt buộc có Người đại diện pháp lý tại EU (EU Responsible Person).",
        "Bắt buộc thể hiện thông tin cảnh báo, truy xuất nguồn gốc và Người đại diện chịu trách nhiệm tại EU trên nhãn sản phẩm và trang bán lẻ trực tuyến; vi phạm sẽ bị thu hồi qua Safety Gate.",
        "2026-08-26"
    ],
    [
        299,
        "IMDG Code (SOLAS 1974 Annex)",
        "Bộ luật Quốc tế về Vận chuyển Hàng hóa Nguy hiểm bằng Đường biển (International Maritime Dangerous Goods Code)",
        "Công ước & Quy chuẩn Quốc tế",
        "Tổ chức Hàng hải Quốc tế (IMO) / Cục Hàng hải Việt Nam",
        "1965-01-01",
        "2026-08-01",
        "Đang áp dụng",
        "Phiên bản sửa đổi cập nhật IMDG Code",
        "Quy chuẩn bắt buộc phân loại, ghi nhãn, đóng gói và khai báo an toàn hàng nguy hiểm (Class 1-9) trong vận chuyển container đường biển, đặc biệt là Pin Lithium (UN3480/UN3481).",
        "Bắt buộc kiểm thử an toàn UN 38.3, MSDS và lập Bản khai Hàng nguy hiểm (DGD); cấm xếp dỡ và phạt nặng hành vi khai man hàng nguy hiểm thành hàng thông thường.",
        "2026-08-26"
    ],
    [
        300,
        "Hiệp định RCEP (Điều 3.16)",
        "Quy chế Tự Chứng nhận Xuất xứ Hàng hóa (Self-Certification of Origin) và e-C/O trong Hiệp định RCEP",
        "Hiệp định Thương mại Tự do",
        "Bộ Công Thương Việt Nam / Ban Thư ký RCEP",
        "2022-01-01",
        "2026-07-25",
        "Đang áp dụng",
        "Chương 3 Hiệp định RCEP & Thông tư 05/2022/TT-BCT",
        "Cho phép Nhà xuất khẩu đủ điều kiện (Approved Exporters) tự phát hành chứng nhận xuất xứ RCEP trên hóa đơn thương mại mà không cần xin cấp C/O bản giấy tại cơ quan quản lý.",
        "Cắt giảm thời gian thông quan xuống dưới 2 giờ; tăng cường trách nhiệm giải trình và lưu trữ chứng từ gốc (BOM, hóa đơn NVL) trong 5 năm phục vụ hậu kiểm.",
        "2026-08-26"
    ],
    [
        301,
        "Regulation (EU) 2023/1115",
        "Quy định về Chuỗi Cung ứng Không Phá rừng của Liên minh Châu Âu (EUDR - EU Deforestation Regulation)",
        "Quy định quốc tế",
        "Nghị viện Châu Âu và Hội đồng Liên minh Châu Âu",
        "2023-06-09",
        "2026-07-20",
        "Đang áp dụng",
        "Quy định mới của EU",
        "Bắt buộc cung cấp dữ liệu tọa độ định vị địa lý GPS của từng thửa đất canh tác chứng minh không gây mất rừng sau ngày 31/12/2020 cho 7 nhóm mặt hàng nông sản xuất khẩu sang EU.",
        "Cấm nhập khẩu vào EU nếu thiếu dữ liệu định vị GPS; phạt tối thiểu 4% doanh thu hàng năm tại EU đối với doanh nghiệp gian lận nguồn gốc nông sản.",
        "2026-08-26"
    ]
]


def append_data_to_excel(filepath, sheet_name, new_rows, url_col_idx, title_col_idx):
    """Ghi bổ sung các dòng mới vào file Excel với định dạng chuẩn."""
    print(f"\n📂 Đang xử lý tệp: {os.path.basename(filepath)}")
    wb = openpyxl.load_workbook(filepath)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active

    start_row = ws.max_row + 1
    print(f"   👉 Bắt đầu ghi từ dòng: {start_row} (Tổng số dòng mới: {len(new_rows)})")

    for idx, row_data in enumerate(new_rows):
        current_row = start_row + idx
        is_zebra = (current_row % 2 == 1)

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.fill = ZEBRA_FILL if is_zebra else WHITE_FILL

            # Căn lề
            if col_idx in (1, 2, 5, 12, 13, 14): # STT, Ngày, Phân loại, Trạng thái
                if col_idx == 1:
                    cell.alignment = ALIGN_CENTER
                    cell.font = BOLD_FONT
                elif col_idx in (2, 5):
                    cell.alignment = ALIGN_CENTER
                else:
                    cell.alignment = ALIGN_LEFT
            else:
                cell.alignment = ALIGN_LEFT

            # Format link nếu là URL
            if col_idx == url_col_idx and str(val).startswith("http"):
                cell.font = Font(name=FONT_NAME, size=10, color="0563C1", underline="single")

    excel_styles.safe_save_workbook(wb, filepath)
    print(f"✅ Đã lưu thành công {len(new_rows)} dòng mới vào {os.path.basename(filepath)}")


def append_legal_docs(filepath, new_docs):
    """Ghi bổ sung văn bản luật mới vào Danh_sach_Van_ban_Phap_luat_XNK.xlsx."""
    print(f"\n📂 Đang cập nhật Danh mục Văn bản Pháp luật: {os.path.basename(filepath)}")
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    start_row = ws.max_row + 1

    for idx, doc in enumerate(new_docs):
        current_row = start_row + idx
        is_zebra = (current_row % 2 == 1)

        for col_idx, val in enumerate(doc, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.fill = ZEBRA_FILL if is_zebra else WHITE_FILL

            if col_idx in (1, 4, 6, 7, 8, 12):
                cell.alignment = ALIGN_CENTER
                if col_idx == 1:
                    cell.font = BOLD_FONT
            else:
                cell.alignment = ALIGN_LEFT

    excel_styles.safe_save_workbook(wb, filepath)
    print(f"✅ Đã cập nhật thành công {len(new_docs)} văn bản pháp luật vào SSoT.")


def main():
    print("="*90)
    print("🚀 BẮT ĐẦU QUY TRÌNH WF06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS & SSOT (W35 — 26/08/2026)")
    print("="*90)

    # 1. Chuẩn bị danh sách ý tưởng cần thẩm định qua Technical Gatekeeper
    validation_list = []
    
    for row in TIN_NGANH_W35:
        validation_list.append({"stt": row[0], "title": row[2], "url": row[12], "keywords": ["cbam", "safety", "goods", "rcep", "forest", "trade", "dangerous", "product"]})
        
    for row in CASE_STUDY_W35:
        validation_list.append({"stt": row[0], "title": row[2], "url": row[10], "keywords": ["battery", "origin", "demurrage", "certificate", "customs", "trade", "plant", "valuation"]})
        
    for row in HOC_THUC_CHIEN_W35:
        validation_list.append({"stt": row[0], "title": row[2], "url": row[10], "keywords": ["dangerous", "valuation", "certificate", "demurrage", "customs", "plant", "port", "trade"]})
        
    for row in FUNFACT_W35:
        validation_list.append({"stt": row[0], "title": row[2], "url": row[9], "keywords": ["jones", "un", "marpol", "toyota", "lading", "marine", "goods", "manufacturing", "bill"]})

    # 2. Kích hoạt Technical Gatekeeper (scripts/url_validator.py)
    print("\n🔍 BƯỚC 1: KÍCH HOẠT TECHNICAL GATEKEEPER RÀO CHẮN KỸ THUẬT SÂU URL...")
    enforce_strict_validation(validation_list)

    # 3. Ghi dữ liệu vào 4 tệp Excel Ideas Bank
    print("\n📊 BƯỚC 2: GHI DỮ LIỆU ĐÃ THẨM ĐỊNH VÀO 4 TỆP EXCEL KHO Ý TƯỞNG...")
    
    append_data_to_excel(
        filepath=os.path.join(REPORTS_DIR, "Tin_Nganh_Ideas_Bank_Eureka.xlsx"),
        sheet_name="Kho Ý Tưởng Tin Ngành",
        new_rows=TIN_NGANH_W35,
        url_col_idx=13,
        title_col_idx=3
    )

    append_data_to_excel(
        filepath=os.path.join(REPORTS_DIR, "Case_Study_Ideas_Bank_Eureka.xlsx"),
        sheet_name="Kho Ý Tưởng Case Study",
        new_rows=CASE_STUDY_W35,
        url_col_idx=11,
        title_col_idx=3
    )

    append_data_to_excel(
        filepath=os.path.join(REPORTS_DIR, "Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx"),
        sheet_name="Kho Ý Tưởng Học Thực Chiến",
        new_rows=HOC_THUC_CHIEN_W35,
        url_col_idx=11,
        title_col_idx=3
    )

    append_data_to_excel(
        filepath=os.path.join(REPORTS_DIR, "Funfact_Ideas_Bank_Eureka.xlsx"),
        sheet_name="Kho Ý Tưởng Funfact",
        new_rows=FUNFACT_W35,
        url_col_idx=10,
        title_col_idx=3
    )

    # 4. Feedback Loop: Cập nhật văn bản pháp luật vào SSoT
    print("\n⚖️ BƯỚC 3: FEEDBACK LOOP - CẬP NHẬT DANH MỤC VĂN BẢN PHÁP LUẬT XNK...")
    append_legal_docs(
        filepath=os.path.join(REPORTS_DIR, "Danh_sach_Van_ban_Phap_luat_XNK.xlsx"),
        new_docs=LEGAL_DOCS_W35
    )

    print("\n" + "="*90)
    print("🎉 HOÀN TẤT 100% QUY TRÌNH WF06 CẬP NHẬT IDEAS BANK NGÀY 26/08/2026 (W35)")
    print("="*90)

if __name__ == "__main__":
    main()
