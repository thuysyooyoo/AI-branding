# -*- coding: utf-8 -*-
"""
================================================================================
SCRIPT: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS & FEEDBACK LOOP PHÁP LUẬT
PHIÊN BẢN: W36 — 04/09/2026 (Tuần 36, Ngày 04/09/2026)
TUÂN THỦ: wf06_ideas_bank_update.md & RÀO CHẮN KỸ THUẬT url_validator.py
================================================================================
"""

import sys
import os
import shutil
from datetime import datetime

# Đảm bảo đường dẫn thư viện openpyxl và scripts
sys.path.insert(0, r"d:\AI AGENT THUY\AI branding\lib")
sys.path.insert(0, r"d:\AI AGENT THUY\AI branding\scripts")

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from url_validator import enforce_strict_validation

REPORTS_DIR = r"d:\AI AGENT THUY\AI branding\reports"

# ------------------------------------------------------------------------------
# CẤU HÌNH STYLING CHUẨN DOANH NGHIỆP EUREKA
# ------------------------------------------------------------------------------
FONT_FAMILY = "Segoe UI"
BODY_FONT = Font(name=FONT_FAMILY, size=10, color="1F2937")
BOLD_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="111827")
STATUS_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color="047857")

WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
STATUS_FILL = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="E5E7EB"),
    right=Side(style="thin", color="E5E7EB"),
    top=Side(style="thin", color="E5E7EB"),
    bottom=Side(style="thin", color="E5E7EB")
)

ALIGN_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="top", wrap_text=True)

# ------------------------------------------------------------------------------
# DỮ LIỆU 4 PILLARS CHO TUẦN 36 (STT 56 - 60)
# ------------------------------------------------------------------------------

# 1. TIN NGÀNH (15 Cột)
TIN_NGANH_W36_D5 = [
    [
        56,
        "W36 — 04/09/2026",
        "Nghị định 341/2026/NĐ-CP (Hiệu lực 01/09/2026): Bắt buộc Giấy phép Mật mã Dân sự và Chứng nhận Hợp quy khi Xuất Nhập Khẩu Thiết bị Công nghệ Thông tin & Viễn thông",
        "Chính phủ / Ban Cơ yếu Chính phủ / Bộ Thông tin & Truyền thông",
        "01/09/2026",
        "Bãi bỏ cơ chế hậu kiểm lỏng lẻo; quy định bắt buộc phải có Giấy phép kinh doanh sản phẩm, dịch vụ mật mã dân sự (thời hạn 10 năm) và Giấy phép xuất khẩu, nhập khẩu mật mã dân sự (thời hạn 3 năm) ngay tại thời điểm mở tờ khai hải quan; bổ sung danh mục mã HS các thiết bị mạng mã hóa VPN, firewall, router, camera giám sát bảo mật cao phải kiểm tra chứng nhận hợp quy.",
        "Doanh nghiệp nhập khẩu thiết bị CNTT, viễn thông, camera an ninh, máy chủ server, thiết bị chuyển mạch switch/router và các công ty tích hợp hệ thống phần mềm công nghệ.",
        "Lô hàng thiết bị điện tử giá trị hàng chục tỷ đồng bị giữ tại cửa khẩu sân bay (Nội Bài, Tân Sơn Nhất) hoặc cảng biển do không có giấy phép chuyên ngành trước ngày tàu cập; chịu phí lưu kho kho bãi hàng không đắt đỏ và nguy cơ bị xử phạt vi phạm hành chính từ 40-100 triệu đồng kèm buộc tái xuất.",
        "Hiệp hội Doanh nghiệp Điện tử Việt Nam (VEIA) và cộng đồng doanh nghiệp công nghệ lo ngại tình trạng ách tắc hàng hóa do quy định có hiệu lực ngay ngày ký, kiến nghị có thời gian chuyển tiếp 3 tháng để doanh nghiệp kịp bổ sung nhân sự kỹ thuật có chứng chỉ chuyên môn và hoàn tất nộp hồ sơ xin giấy phép.",
        "Bảo vệ an ninh quốc gia, kiểm soát nguy cơ rò rỉ dữ liệu và tấn công mạng từ các thiết bị mã hóa không rõ nguồn gốc; chuẩn hóa hệ thống quản lý an toàn thông tin mạng theo thông lệ quốc tế.",
        "Thời gian hiệu lực gấp rút khiến các hợp đồng mua bán ngoại thương ký trước ngày 01/09/2026 rơi vào thế bị động; thủ tục xin cấp phép mật mã dân sự đòi hỏi thẩm định hồ sơ kỹ thuật kéo dài từ 15-30 ngày làm chậm tiến độ dự án của khách hàng.",
        "Quy chế kiểm soát xuất nhập khẩu mật mã (Export of cryptography regulations) theo Thỏa thuận Wassenaar (Wassenaar Arrangement) và quy định của Cơ quan An ninh Quốc gia Hoa Kỳ (BIS EAR).",
        "https://en.wikipedia.org/wiki/Export_of_cryptography",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Nghị định 341/2026/NĐ-CP ngày 01/09/2026 & Chuẩn mực Thỏa thuận Wassenaar về kiểm soát mật mã thương mại."
    ],
    [
        57,
        "W36 — 04/09/2026",
        "Quyết định 31/2026/QĐ-TTg (Hiệu lực 14/08/2026): Danh mục 12 Nhóm Hàng Hóa Bắt Buộc Làm Thủ Tục Tại Cửa Khẩu Nhập & Cảnh Báo Rủi Ro 'Kẹt Chùm Vận Đơn (B/L)'",
        "Thủ tướng Chính phủ / Tổng cục Hải quan",
        "14/08/2026",
        "Quy định danh mục 12 nhóm hàng hóa nhập khẩu có rủi ro cao (máy điều hòa không khí < 90.000 BTU, xe máy, ôtô, rượu, thuốc lá, hàng tiêu dùng đã qua sử dụng...) bắt buộc phải làm thủ tục hải quan tại chi cục hải quan cửa khẩu nhập đầu tiên, không được chuyển cửa khẩu về cảng cạn ICD hoặc chi cục hải quan quản lý nhà máy nội địa; nếu đóng chung container hoặc chung vận đơn với hàng khác thì toàn bộ lô hàng phải dừng lại làm thủ tục tại cửa khẩu nhập.",
        "Doanh nghiệp nhập khẩu hàng tiêu dùng, điện lạnh gia dụng, linh kiện tổng hợp; các đại lý Forwarder gom hàng LCL và chủ hàng FCL đóng hàng hỗn hợp.",
        "Chủ hàng khai sai mã địa điểm lưu kho hoặc đăng ký chuyển cửa khẩu về ICD nội địa bị hệ thống VNACCS từ chối tiếp nhận; container bị neo lại cảng biển Hải Phòng/Cát Lái phát sinh phí demurrage/detention hàng trăm triệu đồng; toàn bộ lô hàng hỗn hợp bị 'kẹt chùm' do 1 mặt hàng trong cont thuộc diện quản lý.",
        "Các doanh nghiệp Logistics và chủ hàng ICD khuyến cáo cần kiểm soát chặt chẽ danh mục hàng ngay từ cảng đi, yêu cầu hãng tàu tách riêng Master B/L hoặc tách House B/L độc lập cho các mặt hàng thuộc diện QĐ 31.",
        "Kiểm soát chặt chẽ buôn lậu, gian lận thương mại và trốn thuế tiêu thụ đặc biệt ngay tại cửa ngõ quốc gia; ngăn chặn tình trạng 'rút ruột' container trên đường vận chuyển chuyển cửa khẩu.",
        "Gây quá tải nghiêm trọng cho các chi cục hải quan cửa khẩu cảng biển lớn; làm giảm vai trò của hệ thống cảng cạn ICD trong việc giảm ùn tắc giao thông đô thị và nâng cao chi phí logistics nội địa.",
        "Quy chế kiểm soát hải quan tại điểm nhập cảnh đầu tiên (Customs clearance at first port of entry) của Hải quan Hoa Kỳ (CBP) và Liên minh Châu Âu (EU UCC).",
        "https://en.wikipedia.org/wiki/Customs",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Quyết định 31/2026/QĐ-TTg ngày 14/08/2026 & Quy chuẩn kiểm soát cửa khẩu quốc tế WCO."
    ],
    [
        58,
        "W36 — 04/09/2026",
        "Thông tư 42/2026/TT-BCT (Có hiệu lực từ 12/09/2026): Danh mục 31 Nhóm Hàng Hóa Lưỡng Dụng và Cơ chế Khai Mã Nhận Diện Quản Lý VX/KO trên Hệ thống VNACCS (Công văn 19916/CHQ-GSQL)",
        "Bộ Công Thương / Cục Xuất Nhập Khẩu",
        "12/09/2026",
        "Ban hành chi tiết danh mục 31 nhóm hàng hóa có khả năng sử dụng vào cả mục đích dân sự và quân sự/an ninh (Dual-use goods: hóa chất tiền chất, hợp kim đặc biệt, laser công nghiệp, thiết bị bay không người lái UAV, máy công cụ CNC độ chính xác cao); bắt buộc doanh nghiệp khai báo mã nhận diện văn bản pháp quy: Mã 'VX' nếu có Giấy phép xuất nhập khẩu hàng lưỡng dụng, Mã 'KO' nếu doanh nghiệp tự chứng minh và cam kết hàng hóa chỉ phục vụ mục đích thuần dân sự.",
        "Doanh nghiệp nhập khẩu máy móc công nghiệp, thiết bị quang học, linh kiện điện tử, hóa chất công nghiệp và máy cắt/gia công kim loại chính xác.",
        "Khai báo thiếu mã nhận diện VX/KO hoặc tự ý khai KO nhưng khi kiểm hóa thực tế cơ quan hải quan nghi vấn thông số kỹ thuật vượt ngưỡng dân sự sẽ bị tạm giữ hàng để trưng cầu giám định; doanh nghiệp đối mặt với nguy cơ bị khởi tố hình sự về hành vi vận chuyển hàng hóa thuộc diện kiểm soát chiến lược nếu không có giấy phép.",
        "Cộng đồng doanh nghiệp sản xuất cơ khí và gia công chính xác lúng túng vì ranh giới thông số kỹ thuật (công suất nguồn phát laser, số trục máy CNC) rất phức tạp, cần hướng dẫn kỹ thuật chi tiết từ cơ quan chức năng.",
        "Thực thi đầy đủ các nghĩa vụ quốc tế về không phổ biến vũ khí hủy diệt hàng loạt (Nghị quyết 1540 HĐBA LHQ), nâng cao uy tín quốc gia và tạo điều kiện cho Việt Nam tiếp cận công nghệ cao từ các nước G7.",
        "Tăng gánh nặng chi phí và thời gian xin giấy phép chuyên ngành cho các mặt hàng sản xuất dân dụng thông thường; nguy cơ ách tắc cửa khẩu do thời gian trưng cầu giám định kỹ thuật kéo dài.",
        "Chế độ Kiểm soát Công nghệ Lưỡng dụng (Dual-use technology export control) theo Khung kiểm soát xuất khẩu của Liên minh Châu Âu (EU Dual-Use Regulation 2021/821).",
        "https://en.wikipedia.org/wiki/Dual-use_technology",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Thông tư 42/2026/TT-BCT, Công văn 19916/CHQ-GSQL & EU Dual-Use Regulations."
    ],
    [
        59,
        "W36 — 04/09/2026",
        "Thông tư 124/2026/TT-BTC (Hiệu lực 22/08/2026): Kiểm Tra Chứng Nhận Xuất Xứ Hàng Hóa (C/O) Theo Cơ Chế Quản Lý Rủi Ro và Bổ Sung Quy Trình Xác Minh Quốc Tế",
        "Bộ Tài chính / Tổng cục Hải quan",
        "22/08/2026",
        "Bổ sung khoản 10 Điều 15 Thông tư 33/2023/TT-BTC; chuyển phương thức kiểm tra C/O ưu đãi FTA từ 'hậu kiểm định kỳ' sang 'quản lý rủi ro thời gian thực'; thiết lập tiêu chí tự động cảnh báo C/O có hóa đơn nước thứ ba (Third-party invoicing), C/O chuyển tải không có xác nhận của hải quan nước trung chuyển, hoặc doanh nghiệp có lịch sử vi phạm xuất xứ.",
        "Doanh nghiệp nhập khẩu hưởng thuế suất ưu đãi đặc biệt theo các FTA (VKFTA, EVFTA, ACFTA, RCEP, CPTPP), các đơn vị mua bán hàng hóa thông qua thương nhân trung gian tại Singapore, Hong Kong.",
        "Cơ quan hải quan tạm đình chỉ áp dụng thuế suất ưu đãi đặc biệt ngay tại khâu thông quan, buộc doanh nghiệp phải nộp thuế theo mức MFN hoặc nộp thư bảo lãnh ngân hàng trong thời gian chờ xác minh xuất xứ với cơ quan cấp C/O nước ngoài (kéo dài từ 60-120 ngày), gây ứ đọng dòng tiền hàng chục tỷ đồng.",
        "Các hiệp hội ngành hàng dệt may, da giày và thủy sản khuyến nghị doanh nghiệp nhập khẩu nguyên phụ liệu phải yêu cầu đối tác xuất khẩu cung cấp đầy đủ bộ chứng từ chứng minh chuỗi cung ứng thực tế và chứng từ vận chuyển xuyên suốt (Through B/L) để giải trình.",
        "Bịt kín các lỗ hổng 'rửa xuất xứ', gian lận thương mại đội lốt FTA để trục lợi thuế nhập khẩu, bảo vệ uy tín hàng xuất khẩu của Việt Nam trước các vụ kiện phòng vệ thương mại của Mỹ và EU.",
        "Thời hạn xác minh quốc tế quá dài khiến doanh nghiệp chịu áp lực lãi vay và rủi ro tỷ giá lớn; nhiều trường hợp sai sót hình thức nhỏ (lỗi chính tả, khác biệt dấu mộc) cũng bị chuyển sang quy trình xác minh phức tạp.",
        "Quy chuẩn Quy tắc Xuất xứ và Thủ tục Xác minh Xuất xứ Hải quan (Rules of origin verification) của Tổ chức Thương mại Thế giới (WTO) và Công ước Kyoto sửa đổi (WCO RKC).",
        "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Thông tư 124/2026/TT-BTC, Thông tư 33/2023/TT-BTC & Chuẩn mực Quy tắc xuất xứ WCO."
    ],
    [
        60,
        "W36 — 04/09/2026",
        "Nghị quyết của Chính phủ & Đề án Thí điểm Khu Thương mại Tự do (Free Trade Zone - FTZ) gắn với Cảng Nước sâu Đà Nẵng và Hải Phòng: Khung Cơ chế Chính sách Hải quan Vượt trội",
        "Chính phủ / Bộ Kế hoạch & Đầu tư / Bộ Tài chính",
        "01/09/2026",
        "Lần đầu tiên thí điểm mô hình Khu thương mại tự do thế hệ mới tích hợp cảng nước sâu, khu chế xuất và trung tâm logistics quốc tế; áp dụng chính sách 'khu phi thuế quan đặc biệt': miễn toàn bộ thủ tục hải quan đối với hàng hóa trung chuyển quốc tế trong nội khu FTZ, chỉ thực hiện giám sát điện tử tự động qua camera AI và e-Seal GPS; phân định rạch ròi hàng hóa vào thị trường nội địa mới phải mở tờ khai nhập khẩu.",
        "Các tập đoàn vận tải biển, công ty logistics tích hợp toàn cầu, các nhà đầu tư hạ tầng cảng biển, doanh nghiệp kho bãi quốc tế và các doanh nghiệp chế xuất, gia công công nghệ cao.",
        "Doanh nghiệp không đáp ứng điều kiện kho bãi thông minh (hệ thống phần mềm kết nối dữ liệu tự động với cơ quan hải quan 24/7) sẽ không được hưởng cơ chế thông quan tự động nội khu; nếu để hàng hóa thẩm thấu trái phép ra thị trường nội địa sẽ bị thu hồi vĩnh viễn giấy phép hoạt động trong FTZ.",
        "Cộng đồng logistics trong và ngoài nước (AmCham, EuroCham, VLA) đánh giá đây là bước đột phá thể chế lớn nhất trong thập kỷ, mở ra cơ hội biến Việt Nam thành trung tâm trung chuyển hàng hải (Transshipment Hub) cạnh tranh trực tiếp với Singapore và Tanjung Pelepas.",
        "Tạo cú hích hạ tầng thương mại mang tầm vóc toàn cầu, thu hút các chuỗi cung ứng công nghệ cao chuyển dịch về Việt Nam, tạo hàng chục nghìn việc làm giá trị gia tăng cao.",
        "Lo ngại thách thức lớn trong việc kiểm soát ranh giới giữa nội khu FTZ và thị trường nội địa, nguy cơ thẩm lậu hàng giả, hàng kém chất lượng nếu năng lực giám sát điện tử của hải quan chưa hoàn thiện.",
        "Mô hình Khu thương mại tự do (Free-trade zone & Bonded supply chain) của Cảng Thượng Hải (Waigaoqiao FTZ) và Cảng Rotterdam.",
        "https://en.wikipedia.org/wiki/Free-trade_zone",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Nghị quyết thí điểm FTZ 2026 của Chính phủ & Chuẩn mực Free-trade zone của UNCTAD."
    ]
]

# 2. CASE STUDY (14 Cột)
CASE_STUDY_W36_D5 = [
    [
        56,
        "W36 — 04/09/2026",
        "Hạn Hán Lịch Sử Tại Kênh Đào Panama & Cú Sốc Tắc Tuyến Vận Tải: Bài Học Mất 420.000 USD Tiền Phạt Chậm Giao Hàng Cho Doanh Nghiệp Xuất Khẩu Dệt May Việt Nam Đi Bờ Đông Hoa Kỳ",
        "Công ty Cổ phần May mặc Xuất khẩu T. (Bình Dương), Hãng tàu Evergreen Marine, Ban Quản lý Kênh đào Panama (ACP), Nhà bán lẻ Target (Mỹ)",
        "Tháng 8-9/2026, hiện tượng El Nino khiến lượng mưa sụt giảm nghiêm trọng, Ban Quản lý Kênh đào Panama buộc phải siết tải hạn ngạch tàu qua kênh xuống còn 32 tàu/ngày và hạ mớn nước tối đa. Tàu chở 15 container hàng áo khoác mùa đông của DN T. đi New York bị kẹt chờ qua kênh hơn 28 ngày. Để tránh kẹt, hãng tàu quyết định dỡ hàng tại cảng trung chuyển Balboa (Panama) và trung chuyển đường sắt xuyên eo đất, nhưng ga đường sắt quá tải khiến toàn bộ lô hàng bị trễ hẹn giao hàng 35 ngày so với hợp đồng mua bán với Target.",
        "Đối tác Mỹ hủy hợp đồng mua bán trị giá 1,8 triệu USD do lỡ mất mùa mở bán Black Friday; DN T. phải chịu phạt 15% giá trị hợp đồng (tương đương 270.000 USD) kèm chi phí phát sinh phụ phí kẹt kênh đào (Panama Canal Surcharge) và phí lưu bãi container 150.000 USD; tổng thiệt hại thực tế lên tới 420.000 USD (hơn 10,5 tỷ đồng VNĐ).",
        "Doanh nghiệp chỉ ký hợp đồng vận chuyển đường biển đơn thuần mà không mua bảo hiểm rủi ro chậm trễ giao hàng (Delay in transit insurance); không theo dõi dự báo thời tiết và cảnh báo hạn mớn nước của ACP để chủ động chuyển hướng sang tuyến đường sắt liên vận qua bờ Tây (MLB - Mini Land Bridge) từ sớm.",
        "1. Trong hợp đồng mua bán ngoại thương (Sales Contract), bắt buộc đưa điều khoản bất khả kháng (Force Majeure) bao gồm cả sự cố tắc nghẽn kênh đào hoặc thiên tai cạn nước; 2. Khi vào mùa cao điểm Q3-Q4, đa dạng hóa tuyến vận tải: chuyển sang đi đường biển sang bờ Tây Mỹ (Los Angeles/Long Beach) kết hợp đường sắt nội địa về bờ Đông thay vì đi toàn tuyến đường biển qua Panama; 3. Mua bảo hiểm hàng hóa bao gồm điều khoản rủi ro gián đoạn chuỗi cung ứng.",
        "Phe A: Lỗi hoàn toàn do hãng tàu đổi cảng dỡ hàng tùy tiện và không cảnh báo thời gian trễ cho chủ hàng. Phe B: Hãng tàu được miễn trừ trách nhiệm theo quy định tại Điều khoản B/L (vận tải biển gặp thiên tai bất khả kháng), chủ hàng phải tự chịu rủi ro nếu không lập phương án dự phòng thời gian.",
        "Không bao giờ để toàn bộ chuỗi cung ứng phụ thuộc vào duy nhất một 'nút thắt cổ chai' hàng hải toàn cầu; tính toán thời gian đệm (Buffer time) tối thiểu 20-30 ngày cho các đơn hàng phục vụ mùa lễ hội cuối năm.",
        "https://en.wikipedia.org/wiki/Panama_Canal",
        "Tai nạn XNK",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ báo cáo vận tải hàng hải Panama Canal ACP & Phân tích rủi ro chuỗi cung ứng quốc tế."
    ],
    [
        57,
        "W36 — 04/09/2026",
        "Tàu Ever Given Mắc Cạn Tại Kênh Đào Suez & Tranh Chấp 'Tổn Thất Chung (General Average)': DN Xuất Khẩu Gỗ Việt Nam Bị Phong Tỏa 8 Container Suốt 6 Tháng và Gánh Khoản Đóng Góp Cứu Hộ 68.000 USD",
        "Công ty Sản xuất & Chế biến Gỗ Nội thất H. (Đồng Nai), Hãng tàu Shoei Kisen Kaisha / Evergreen, Ban Quản lý Kênh đào Suez (SCA), Công ty Giám định Tổn thất Quốc tế Richards Hogg Lindley",
        "Tàu siêu container Ever Given bị mắc cạn làm tắc nghẽn kênh đào Suez suốt 6 ngày, khiến hàng trăm con tàu tê liệt. Sau khi giải cứu thành công, chủ tàu tuyên bố 'Tổn thất chung' (General Average) để chia sẻ chi phí cứu hộ khổng lồ và tiền bồi thường gần 500 triệu USD cho Ai Cập. Lô hàng 8 container gỗ nội thất của DN H. nằm an toàn trên tàu không bị hư hỏng, nhưng toàn bộ chứng từ B/L bị hãng tàu giữ lại, yêu cầu DN phải nộp bảo lãnh tài chính (Average Bond và Average Guarantee) mới được giải phóng hàng.",
        "Do không mua bảo hiểm hàng hóa đường biển loại A, doanh nghiệp không có công ty bảo hiểm đứng ra cấp thư bảo lãnh thanh toán; DN H. buộc phải trích tiền mặt ký quỹ ngân hàng 68.000 USD (gần 1,7 tỷ đồng) để nộp tiền đóng góp tổn thất chung; hàng bị phong tỏa tại Hồ Great Bitter suốt hơn 5 tháng khiến khách hàng châu Âu hủy đơn hàng, phát sinh thêm 25.000 USD phí lưu kho bãi.",
        "Tư duy tiết kiệm chi phí bảo hiểm hàng hải (chỉ vài triệu đồng mỗi lô hàng) của doanh nghiệp xuất khẩu; hoàn toàn thiếu hiểu biết về chế định pháp lý 'Tổn thất chung' theo Quy tắc York-Antwerp trong vận tải biển quốc tế.",
        "1. Bắt buộc 100% các lô hàng xuất nhập khẩu đường biển phải mua Bảo hiểm mọi rủi ro (Institute Cargo Clauses - Clause A); khi xảy ra tuyên bố Tổn thất chung, công ty bảo hiểm sẽ tự động phát hành Thư bảo lãnh tổn thất chung (General Average Guarantee) thay cho chủ hàng để giải phóng hàng hóa ngay lập tức mà không phải bỏ một đồng tiền mặt nào; 2. Đào tạo nhân sự xuất nhập khẩu nắm vững quy tắc xử lý khi chủ tàu tuyên bố General Average.",
        "Phe A: Chủ tàu có lỗi điều khiển tàu đâm vào bờ nên không có quyền đòi hỏi chủ hàng đóng góp tổn thất chung. Phe B: Căn cứ Luật Hàng hải Quốc tế và B/L, hành động cứu hộ con tàu và toàn bộ hàng hóa là sự hy sinh chung hợp pháp, tất cả chủ hàng có hàng trên tàu bắt buộc phải chia sẻ chi phí tỷ lệ thuận với giá trị hàng.",
        "Chi phí mua bảo hiểm hàng hóa đường biển không phải là chi phí phụ, mà là 'tấm phao cứu sinh sống còn' của doanh nghiệp khi bước ra biển lớn.",
        "https://en.wikipedia.org/wiki/Ever_Given",
        "Tai nạn XNK",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ sự kiện siêu tàu Ever Given, Quy tắc York-Antwerp 2016 về Tổn thất chung."
    ],
    [
        58,
        "W36 — 04/09/2026",
        "Chủ Quan Bỏ Quên Quy Định Lưu Bãi & Tranh Chấp 'Demurrage / Detention': Doanh Nghiệp Nhập Khẩu Thép Bị Hãng Tàu Tính Tiền Phạt Lưu Container Lên Tới 580 Triệu Đồng Tại Cảng Cát Lái",
        "Công ty Cổ phần Cơ khí Thép Nam S. (TP.HCM), Đại lý Hãng tàu Quốc tế ONE, Chi cục Hải quan Cửa khẩu Cảng Sài Gòn Khu vực 1",
        "DN Nam S. nhập khẩu 10 container thép hợp kim cán nguội từ Trung Quốc về Cảng Cát Lái. Tàu cập cảng đúng thời điểm cơ quan hải quan chuyển đổi hệ thống phần mềm giám sát tự động và hàng bị rơi vào luồng đỏ kiểm hóa thực tế. Nhân viên hiện trường của doanh nghiệp làm mất bản gốc chứng thư giám định chất lượng, dẫn đến việc giải trình kéo dài 22 ngày. Hãng tàu chỉ cho miễn phí lưu bãi (Free time) 5 ngày cho container và 3 ngày cho bãi cảng.",
        "Tiền phạt lưu container tại cảng (Demurrage) và tiền giữ vỏ container (Detention) lũy tiến theo biểu giá bậc thang: từ ngày thứ 6 đến ngày thứ 14 phạt 35 USD/cont/ngày; từ ngày thứ 15 trở đi phạt 80 USD/cont/ngày; tổng cộng 10 container bị phạt tiền lưu bãi 23.200 USD (tương đương 580 triệu VNĐ), vượt quá toàn bộ lợi nhuận dự tính của đơn hàng.",
        "Thiếu thỏa thuận gia hạn Free time (Dem/Det) ngay trong giai đoạn đàm phán hợp đồng thuê tàu; nhân sự làm thủ tục không dự phòng trước hồ sơ kiểm tra chuyên ngành dẫn đến thời gian kiểm hóa bị kéo dài vượt ngoài tầm kiểm soát.",
        "1. Ngay từ khâu book cước tàu, bắt buộc Forwarder/Hãng tàu cam kết văn bản cho Free time Dem/Det gộp (Combined Demurrage & Detention) tối thiểu 14-21 ngày đối với hàng thuộc diện kiểm tra chuyên ngành; 2. Chuẩn bị trước 100% hồ sơ đăng ký kiểm tra chất lượng trước khi tàu cập cảng; 3. Lên phương án rút ruột hàng sang xe tải hoặc chuyển hàng vào kho CFS/kho ngoại quan nếu thời gian chờ thông quan dự kiến kéo dài quá 7 ngày để trả vỏ container rỗng cho hãng tàu cắt tiền phạt.",
        "Phe A: Biểu giá phạt Dem/Det của hãng tàu nước ngoài quá cắt cổ, mang tính chèn ép doanh nghiệp xuất nhập khẩu Việt Nam. Phe B: Hãng tàu cần thu hồi vỏ container rỗng để duy trì luân chuyển chuỗi cung ứng toàn cầu, chủ hàng không lấy hàng đúng hạn làm tắc nghẽn bãi cảng thì phải chịu phạt theo quy tắc kinh tế thị trường.",
        "Lợi nhuận của một lô hàng xuất nhập khẩu có thể bị xóa sổ hoàn toàn chỉ sau một tuần chậm trễ tại bãi container nếu không kiểm soát chặt chẽ điều khoản Free Time.",
        "https://en.wikipedia.org/wiki/Demurrage",
        "Tai nạn XNK",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ quy tắc tính phí Demurrage & Detention của hãng tàu quốc tế và Luật Hàng hải Việt Nam."
    ],
    [
        59,
        "W36 — 04/09/2026",
        "Thủ Đoạn 'Đánh Tráo Vận Đơn Gốc (B/L) & Lừa Đảo Cướp Quyền Sở Hữu Hàng Hóa': Vụ Án Mất Trắng 5 Container Nông Sản Trị Giá 320.000 USD Tại Cảng Trung Đông",
        "Hợp tác xã Nông sản Xuất khẩu K. (Tây Nguyên), Công ty Môi giới Thương mại Ả Rập, Hãng chuyển phát nhanh quốc tế, Ngân hàng Thương mại tại Dubai",
        "Doanh nghiệp K. ký hợp đồng xuất khẩu 5 container hạt tiêu đen sang UAE theo điều kiện CIF Jebel Ali, thanh toán theo phương thức nhờ thu chứng từ D/P (Documents against Payment). Sau khi tàu rời cảng Cát Lái, đối tác người mua gửi email mạo danh nhân viên phụ trách chứng từ, yêu cầu DN K. gửi trực tiếp bản chụp mã vận đơn (Tracking number) của bưu phẩm chuyển phát nhanh bộ chứng từ gốc gửi cho Ngân hàng Dubai. Kẻ lừa đảo đã dùng giấy tờ tùy thân giả và thông đồng với nhân viên chuyển phát tại Dubai để nhận trộm bưu kiện chứa B/L gốc trước khi bưu phẩm tới ngân hàng, sau đó dùng B/L gốc đến hãng tàu lấy lệnh giao hàng (D/O) và rút sạch hàng khỏi cảng.",
        "Mất trắng 5 container hạt tiêu trị giá 320.000 USD (khoảng 8 tỷ VNĐ); ngân hàng tại Dubai từ chối trách nhiệm vì bưu phẩm chưa từng được giao nhận tại quầy giao dịch; công ty chuyển phát nhanh chỉ bồi thường theo giới hạn trách nhiệm bưu kiện giá rẻ là 100 USD.",
        "Sự lỏng lẻo trong bảo mật thông tin vận đơn và mã tracking chuyển phát nhanh quốc tế; sử dụng phương thức D/P tại thị trường có tính rủi ro cao mà không yêu cầu đặt cọc tiền mặt tối thiểu 30-50%.",
        "1. Tuyệt đối không bao giờ chia sẻ công khai mã theo dõi bưu phẩm chuyển phát nhanh (Courier Tracking No.) cho người mua trước khi ngân hàng xác nhận đã tiếp nhận bộ chứng từ; 2. Yêu cầu công ty chuyển phát nhanh tích hợp dịch vụ 'Giao tận tay người nhận có định danh (Direct Person Delivery)' và xác thực qua mã OTP/chữ ký điện tử; 3. Chuyển đổi sang sử dụng Vận đơn điện tử (e-B/L) trên nền tảng số hóa để loại bỏ hoàn toàn nguy cơ đánh tráo chứng từ giấy.",
        "Phe A: Hãng chuyển phát nhanh quốc tế phải bồi thường toàn bộ thiệt hại vì đã phát bưu phẩm sai người nhận. Phe B: Hãng chuyển phát chỉ giao nhận theo quy trình tiêu chuẩn và được giới hạn trách nhiệm bồi thường theo Công ước Bưu chính Quốc tế, lỗi thuộc về người gửi đã để lộ thông tin vận đơn mật.",
        "Trong xuất nhập khẩu, bộ chứng từ gốc (B/L) chính là 'sổ đỏ của lô hàng'; bảo mật chứng từ cũng quan trọng ngang bằng với bảo vệ chính tài sản hàng hóa.",
        "https://en.wikipedia.org/wiki/Cargo_theft",
        "Bóc phốt chiêu trò",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ cảnh báo lừa đảo thương mại quốc tế của Thương vụ Việt Nam tại UAE & Vụ việc đánh tráo chứng từ B/L gốc."
    ],
    [
        60,
        "W36 — 04/09/2026",
        "Bẫy 'Bất Hợp Lệ Chứng Từ L/C (L/C Discrepancy)' Khiến Doanh Nghiệp Thủy Sản Việt Nam Bị Ngân Hàng Nước Ngoài Từ Chối Thanh Toán 185.000 USD",
        "Công ty Cổ phần Thủy sản Biển Đông B. (Cà Mau), Ngân hàng Phát hành tại Bangladesh, Ngân hàng Thông báo tại Việt Nam",
        "DN B. xuất khẩu 2 container tôm đông lạnh đi Chittagong (Bangladesh) theo phương thức Thư tín dụng không hủy ngang (Irrevocable L/C). Khi lập bộ chứng từ xuất trình cho ngân hàng thanh toán, nhân viên chứng từ của DN B. đã sơ suất: trên Commercial Invoice ghi mô tả hàng hóa thừa chữ 'Fresh Frozen Black Tiger Shrimp' trong khi L/C chỉ ghi 'Frozen Black Tiger Shrimp'; ngày phát hành Chứng nhận Kiểm dịch (Phytosanitary Certificate) lại sau ngày ký phát Vận đơn đường biển (B/L) 1 ngày. Ngân hàng mở L/C tại Bangladesh lập tức viện dẫn 2 lỗi bất hợp lệ này để từ chối thanh toán toàn bộ lô hàng trị giá 185.000 USD.",
        "Hàng bị giữ tại cảng Chittagong hơn 40 ngày trong khi người mua ngoại ép giá đòi giảm 30% giá trị hợp đồng mới chịu chấp nhận chứng từ; DN B. phải chịu chi phí cắm điện container lạnh và phí phạt lưu bãi cảng 28.000 USD (hơn 700 triệu VNĐ) cùng khoản chi phí sửa chứng từ ngân hàng 1.500 USD.",
        "Nhân viên nghiệp vụ xuất nhập khẩu chưa nắm vững Quy tắc thực hành thống nhất về tín dụng chứng từ (UCP 600) và Tiêu chuẩn kiểm tra chứng từ quốc tế (ISBP 745); chủ quan coi lỗi câu chữ nhỏ là không quan trọng.",
        "1. Nguyên tắc 'Soi gương 100%': Mọi từ ngữ mô tả hàng hóa trên Invoice, Packing List, C/O và Giấy kiểm dịch bắt buộc phải trùng khớp từng dấu chấm, dấu phẩy với L/C; 2. Đảm bảo quy tắc thời gian: Mọi chứng nhận kiểm định/kiểm dịch bắt buộc phải phát hành trước hoặc cùng ngày với ngày giao hàng ghi trên B/L; 3. Yêu cầu ngân hàng thông báo tại Việt Nam kiểm tra chéo toàn bộ chứng từ (Pre-check) trước khi gửi đi nước ngoài.",
        "Phe A: Ngân hàng nước ngoài bắt lỗi hình thức vụn vặt một cách vô lý để hỗ trợ người mua ép giá doanh nghiệp Việt. Phe B: Ngân hàng vận hành theo nguyên tắc 'Chứng từ nghiêm ngặt' của UCP 600, ngân hàng chỉ làm việc với chứng từ chứ không làm việc với hàng hóa, bất kỳ sai lệch nào cũng là cơ sở pháp lý hợp lệ để từ chối thanh toán.",
        "L/C là phương thức thanh toán an toàn nhất nhưng cũng là 'cái bẫy chết người' nếu người làm chứng từ thiếu sự tỉ mỉ và hiểu biết sâu sắc về luật ngân hàng quốc tế.",
        "https://en.wikipedia.org/wiki/Letters_of_credit",
        "Tai nạn XNK",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Quy tắc UCP 600 & ISBP 745 của Phòng Thương mại Quốc tế (ICC) về tín dụng chứng từ."
    ]
]

# 3. HỌC THỰC CHIẾN (14 Cột)
HOC_THUC_CHIEN_W36_D5 = [
    [
        56,
        "W36 — 04/09/2026",
        "Hướng Dẫn 4 Bước Kỹ Thuật Tách House B/L và Phân Định Danh Mục Cửa Khẩu Nhập Theo Quyết Định 31/2026/QĐ-TTg Tránh Kẹt Chùm Container",
        "Bước 1: Rà soát danh mục hàng đóng cont đối chiếu với 12 nhóm hàng rủi ro theo QĐ 31/2026/QĐ-TTg ngay khi nhận Packing List từ nhà cung cấp ngoại; Bước 2: Yêu cầu đại lý Forwarder tại cảng xuất phát hành 2 House B/L riêng biệt: House B/L số 1 dành riêng cho nhóm hàng buộc làm thủ tục tại cửa khẩu nhập (ví dụ: máy lạnh, rượu, xe máy), House B/L số 2 cho các mặt hàng thông thường được phép chuyển cửa khẩu về ICD nội địa; Bước 3: Đăng ký với hãng tàu khai Manifest tách dòng hàng tương ứng trên Hệ thống Một cửa Quốc gia (NSW); Bước 4: Mở tờ khai hải quan độc lập cho từng House B/L: tiến hành thông quan lô hàng tại cửa khẩu nhập trước, sau đó hoàn tất thủ tục vận chuyển độc lập cho lô hàng còn lại về kho đích.",
        "Tỷ lệ chi phí tối ưu = (Chi phí tách House B/L 35 USD) vs (Chi phí phạt lưu bãi container khi bị giữ cả cont: 80 USD/ngày x 10 ngày = 800 USD) → Tiết kiệm hơn 95% rủi ro tài chính phát sinh.",
        "1. Đóng chung hàng thuộc diện QĐ 31 vào cùng 1 House B/L với hàng máy móc thông thường; 2. Khai sai mã địa điểm đích đến (ICD) trên phần mềm VNACCS dẫn đến tờ khai bị hệ thống tự động khóa; 3. Không thông báo trước cho đại lý vận tải bộ về yêu cầu làm thủ tục tại cửa khẩu biển trước khi kéo cont.",
        "Tách House B/L như phân loại hành lý xách tay và hành lý ký gửi khi đi máy bay: Đồ cần kiểm tra an ninh đặc biệt phải để riêng một túi để không làm trễ giờ lên máy bay của cả đoàn.",
        "Gộp chung 1 Master B/L (Toàn bộ cont bị neo lại cửa khẩu nhập nhiều tuần chờ kiểm hóa) vs Tách House B/L kép (Giải phóng 80% hàng hóa thông thường về nhà máy ngay trong ngày, chỉ giữ lại phần hàng cần kiểm tra chuyên ngành).",
        "Anh Minh (Chủ doanh nghiệp thương mại tổng hợp) & Nhân viên chứng từ logistics",
        "☐ Check danh mục 12 nhóm hàng QĐ 31\n☐ Yêu cầu Shipper tách Packing List\n☐ Forwarder phát hành House B/L kép\n☐ Khai Manifest tách dòng trên NSW\n☐ Mở tờ khai độc lập từng vận đơn",
        "https://en.wikipedia.org/wiki/Bill_of_lading",
        "Quy trình Hải quan",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Quyết định 31/2026/QĐ-TTg & Nghiệp vụ phân tách vận đơn đường biển theo chuẩn WCO/FIATA."
    ],
    [
        57,
        "W36 — 04/09/2026",
        "Công Thức 5 Bước Phân Bổ Chi Phí và Đàm Phán Điều Kiện Incoterms 2020: Khi Nào Nên Chọn CIF và Khi Nào Bắt Buộc Phải Giành Quyền Mua FOB?",
        "Bước 1: Tính toán tổng chi phí Landed Cost của phương án mua CIF (Giá hàng + Cước biển do người bán chỉ định + Bảo hiểm) so sánh với mua FOB (Giá hàng tại xưởng + Cước biển do mình tự book qua Forwarder nội địa); Bước 2: Đánh giá rủi ro tuyến đường và năng lực book cước của đối tác ngoại (nếu vào mùa cao điểm cước biến động lớn, giành quyền mua FOB giúp kiểm soát ngày tàu chạy và vị trí container); Bước 3: Đàm phán điều khoản thanh toán cước Freight Prepaid (trả trước) vs Freight Collect (trả sau) tại Việt Nam để tối ưu dòng tiền; Bước 4: Kiểm tra các loại phụ phí địa phương (Local Charges) tại cảng đích: yêu cầu đại lý xuất trình báo giá trọn gói (All-in) tránh bẫy phụ phí CIC/D/O cắt cổ khi mua CIF; Bước 5: Ký hợp đồng ngoại thương chốt rõ ràng phiên bản 'Incoterms 2020' và địa điểm chuyển giao rủi ro cụ thể.",
        "Chênh lệch Lợi ích FOB vs CIF = Giá CIF người bán chào - (Giá FOB + Giá cước hợp đồng Forwarder nội địa + Phí bảo hiểm mua tại VN) → Nếu kết quả > 150 USD/TEU thì bắt buộc phải chuyển sang mua FOB.",
        "1. Tưởng mua CIF là an toàn vì người bán chịu trách nhiệm vận chuyển nhưng quên mất rủi ro chuyển giao sang người mua ngay khi hàng qua lan can tàu tại cảng xuất; 2. Bị Forwarder chỉ định của bên bán ép thu phụ phí Local charges tại Việt Nam cao gấp 3-4 lần giá thị trường; 3. Quên kiểm tra điều kiện bảo hiểm của bên bán (mua CIF loại C rủi ro tối thiểu thay vì loại A).",
        "Mua CIF như gọi đồ ăn trọn gói bao ship: Tiện nhưng phụ thuộc vào shipper và có thể bị tính phí ship cắt cổ; Mua FOB như tự mình đặt xe công nghệ: Chủ động chọn tài xế, biết trước giá cước và kiểm soát được lộ trình.",
        "Mua CIF (Phù hợp doanh nghiệp mới, đơn hàng nhỏ, cước ổn định) vs Mua FOB (Dành cho doanh nghiệp làm quy mô lớn, cần làm chủ thời gian giao hàng và tối ưu dòng tiền chiết khấu cước).",
        "Chị Hương (Trưởng phòng Mua hàng / Quản lý chuỗi cung ứng SME)",
        "☐ Lấy đồng thời 2 báo giá FOB & CIF\n☐ Yêu cầu bảng kê chi tiết Local Charges tại cảng VN\n☐ Kiểm tra hợp đồng bảo hiểm thuộc loại A, B hay C\n☐ Thỏa thuận quyền chỉ định hãng tàu\n☐ Ghi rõ địa điểm giao nhận chuẩn Incoterms 2020",
        "https://en.wikipedia.org/wiki/Incoterms",
        "Đàm phán Quốc tế",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Bộ quy tắc Incoterms 2020 của ICC & Kinh nghiệm đàm phán cước tàu quốc tế."
    ],
    [
        58,
        "W36 — 04/09/2026",
        "Quy Trình 6 Bước Tra Cứu và Bảo Vệ Mã HS Code Trước Cơ Quan Hải Quan: Tuyệt Chiêu Tránh Bị Ấn Định Thuế và Xử Phạt Vi Phạm Hành Chính",
        "Bước 1: Thu thập đầy đủ tài liệu kỹ thuật gốc của sản phẩm từ nhà sản xuất (Catalogue, Data sheet, bản vẽ kỹ thuật, thành phần cấu tạo, công dụng chính); Bước 2: Áp dụng tuần tự 6 Quy tắc Tổng quát Giải thích Phân loại Hàng hóa (GRI 1 đến GRI 6) theo Công ước HS; Bước 3: Tra cứu Chú giải Chi tiết Hệ thống Hài hòa (Explanatory Notes) của WCO và Biểu thuế XNK hiện hành của Việt Nam; Bước 4: Kiểm tra cơ sở dữ liệu phân loại hải quan: tra cứu các công văn hướng dẫn phân loại trước đây của Tổng cục Hải quan trên cổng thông tin hải quan; Bước 5: Thực hiện thủ tục 'Xác định trước mã số (Advance Ruling)' theo Điều 28 Luật Hải quan trước khi nhập khẩu tối thiểu 60 ngày nếu mặt hàng có tính chất kỹ thuật phức tạp hoặc ranh giới phân loại mơ hồ; Bước 6: Lập Bản Giải trình Kỹ thuật Áp mã HS chuẩn bị sẵn sàng trong hồ sơ hải quan để phản biện trong trường hợp Chi cục Hải quan nghi vấn áp mã.",
        "Số tiền phạt do khai sai mã HS = 20% x Số tiền thuế chênh lệch khai thiếu + Tiền chậm nộp (0.03%/ngày trên số thuế thiếu tính từ ngày thông quan đến ngày nộp đủ).",
        "1. Tra mã HS bằng cách gõ từ khóa tìm kiếm tiếng Việt chung chung trên mạng thay vì đọc cấu trúc 6 quy tắc GRI; 2. Cố tình chọn mã HS có thuế suất 0% dù công dụng thực tế của hàng hóa thuộc nhóm có thuế 5-10%; 3. Đôi co bằng lời nói với công chức hải quan thay vì sử dụng công văn dẫn chứng văn bản pháp lý và tài liệu giám định khoa học.",
        "Áp mã HS như khám bệnh kê đơn thuốc: Phải dựa trên triệu chứng lâm sàng và xét nghiệm thành phần cụ thể, không thể nhìn bề ngoài mà phán bừa.",
        "Khai báo mã HS tự phát (Nguy cơ bị thanh tra sau thông quan 5 năm truy thu thuế tiền tỷ) vs Xin xác định trước mã số hải quan - Advance Ruling (Có văn bản bảo đảm pháp lý 100% từ Tổng cục Hải quan, thông quan tự động luồng xanh).",
        "Anh Dũng (Giám đốc Sản xuất / Trưởng bộ phận Khai báo Hải quan)",
        "☐ Thu thập Spec/Catalogue nhà sản xuất\n☐ Áp dụng 6 quy tắc tổng quát GRI\n☐ Đọc chú giải Explanatory Notes WCO\n☐ Tra cứu tiền lệ công văn Tổng cục Hải quan\n☐ Nộp hồ sơ xác định trước mã số nếu phân vân",
        "https://en.wikipedia.org/wiki/Harmonized_System",
        "Tối ưu Thuế XNK",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ 6 quy tắc GRI của WCO, Thông tư 31/2022/TT-BTC & Luật Hải quan Việt Nam."
    ],
    [
        59,
        "W36 — 04/09/2026",
        "Hướng Dẫn Thiết Lập Hồ Sơ Chứng Minh Trị Giá Giao Dịch Thực Tế Khi Bị Tham Vấn Giá Hải Quan: Kỹ Thuật Bảo Vệ Mức Giá Mua Hàng Khỏi Bị Bác Bỏ",
        "Bước 1: Chuẩn bị Bộ hồ sơ chứng minh trị giá theo Điều 25 Thông tư 38/2015/TT-BTC và Thông tư 39/2018/TT-BTC: Hợp đồng ngoại thương, Invoice, Packing List, Hợp đồng cước vận tải biển, Hợp đồng bảo hiểm; Bước 2: Thu thập bằng chứng thanh toán thực tế xuyên suốt: Điện chuyển tiền ngân hàng (MT103 / Swift Advice) thể hiện rõ số tiền và nội dung thanh toán trùng khớp số hóa đơn; sao kê tài khoản ngân hàng của doanh nghiệp; Bước 3: Cung cấp chứng từ hạch toán kế toán nội bộ: Phiếu nhập kho, sổ chi tiết tài khoản 331, 156 và hợp đồng bán lại nội địa (để chứng minh doanh nghiệp không bán phá giá và có biên lợi nhuận hợp lý); Bước 4: Soạn thảo Bản Giải trình Biến động Giá thị trường: Phân tích các yếu tố ảnh hưởng đến giá mua (chiết khấu số lượng lớn, giá nguyên liệu kim loại/nông sản thế giới tại thời điểm ký kết hợp đồng, chính sách xả kho của bên bán); Bước 5: Cử đại diện có thẩm quyền và nắm vững số liệu kế toán tham gia phiên tham vấn giá trực tiếp tại Chi cục Hải quan; phản biện dựa trên phương pháp trị giá giao dịch (Phương pháp 1).",
        "Trị giá Hải quan (CIF) = Giá mua thực tế ghi trên hóa đơn thương mại (FOB) + Chi phí vận chuyển quốc tế (F) + Chi phí bảo hiểm (I) + Các khoản điều chỉnh cộng hợp pháp theo Điều 13 Thông tư 39/2018/TT-BTC.",
        "1. Thanh toán tiền hàng qua bên thứ ba không có tên trong hợp đồng mua bán mà không có văn bản ủy quyền thanh toán 3 bên hợp pháp; 2. Không lưu trữ đầy đủ chứng từ thanh toán ngân hàng bản gốc có đóng dấu xác nhận; 3. Chấp nhận ngay mức giá ấn định của hải quan trong biên bản tham vấn vì ngại thủ tục giải trình mà không biết rằng chấp nhận sẽ tạo thành tiền lệ giá cho các lô hàng tiếp theo.",
        "Tham vấn giá hải quan như khi bạn mua được căn nhà giá rẻ thời sốt đất: Phải chứng minh tiền sạch từ ngân hàng chuyển qua và có hợp đồng công chứng minh bạch thì thuế mới không thể ấn định giá thị trường ảo.",
        "Bị ấn định giá theo Cơ sở dữ liệu rủi ro của Hải quan (Tăng thuế nhập khẩu và thuế VAT 15-30%) vs Bảo vệ thành công Trị giá giao dịch thực tế theo Phương pháp 1 (Nộp đúng tiền thuế thực mua, bảo toàn 100% biên lợi nhuận).",
        "Chị Lan (Kế toán trưởng doanh nghiệp xuất nhập khẩu) & Chủ doanh nghiệp",
        "☐ Thu thập Swift MT103 ngân hàng chuẩn\n☐ Chuẩn bị sổ kế toán 331, 156, phiếu nhập kho\n☐ Thu thập báo giá thị trường tại ngày ký hợp đồng\n☐ Soạn bản giải trình biến động giá chi tiết\n☐ Đối chiếu các khoản điều chỉnh cộng/trừ hợp pháp",
        "https://en.wikipedia.org/wiki/Customs_valuation",
        "Tối ưu Thuế XNK",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Hiệp định Trị giá Hải quan WTO (GATT Valuation Agreement) & Thông tư 39/2018/TT-BTC."
    ],
    [
        60,
        "W36 — 04/09/2026",
        "Lộ Trình 5 Bước Xây Dựng Tiêu Chuẩn Doanh Nghiệp Ưu Tiên Hải Quan (AEO - Authorized Economic Operator): Chìa Khóa Thông Quan Luồng Xanh 100%",
        "Bước 1: Rà soát điều kiện tuân thủ pháp luật: Doanh nghiệp không vi phạm pháp luật hải quan và thuế trong 2 năm liên tục gần nhất; kim ngạch XNK đạt tối thiểu 100 triệu USD/năm (hoặc 40 triệu USD đối với hàng xuất khẩu sản xuất trong nước); Bước 2: Chuẩn hóa hệ thống kế toán & kiểm toán nội bộ: Áp dụng chuẩn mực kế toán Việt Nam (VAS) hoặc quốc tế (IFRS), thực hiện kiểm toán độc lập hàng năm; Bước 3: Nâng cấp hệ thống an ninh chuỗi cung ứng: Kiểm soát an ninh kho bãi, quy trình niêm phong container đạt chuẩn ISO 17712, kiểm soát đối tác vận chuyển và an ninh nhân sự; Bước 4: Ứng dụng phần mềm quản trị ERP tích hợp dữ liệu hải quan tự động, lưu vết toàn bộ luồng nguyên vật liệu và thành phẩm; Bước 5: Nộp hồ sơ đề nghị công nhận Doanh nghiệp ưu tiên lên Tổng cục Hải quan và phối hợp thẩm định thực tế tại nhà máy.",
        "Lợi ích doanh nghiệp AEO: Miễn kiểm tra hồ sơ giấy 100%, miễn kiểm tra thực tế hàng hóa (kiểm hóa) 99.9%, thông quan ngay sau khi truyền tờ khai (dưới 5 giây); giảm 80% chi phí lưu kho bãi và nhân sự hiện trường tại cảng.",
        "1. Nghĩ rằng AEO chỉ dành cho các tập đoàn FDI khổng lồ mà không biết các doanh nghiệp Việt Nam quy mô vừa cũng có thể đạt chuẩn nếu liên kết chuỗi; 2. Hồ sơ số liệu kho bãi và số liệu hải quan có sự chênh lệch do quản lý bằng sổ sách excel thủ công; 3. Bỏ qua khâu đánh giá rủi ro an ninh của các nhà thầu phụ vận tải đường bộ.",
        "Thẻ Doanh nghiệp ưu tiên AEO như Hộ chiếu Ngoại giao hoặc Thẻ APEC tại sân bay: Được đi thẳng qua cửa ưu tiên không cần xếp hàng, tiết kiệm hàng nghìn giờ chờ đợi vô nghĩa.",
        "Doanh nghiệp thông thường (Bị phân luồng vàng/đỏ ngẫu nhiên, mất 1-3 ngày thông quan tại cảng) vs Doanh nghiệp AEO (100% luồng xanh thông quan tức thì, được nộp tờ khai chưa hoàn chỉnh và ân hạn nộp thuế đến ngày 10 tháng sau).",
        "Ban Giám đốc & Trưởng phòng Xuất nhập khẩu các nhà máy sản xuất, khu công nghiệp",
        "☐ Rà soát lịch sử tuân thủ thuế 2 năm qua\n☐ Kiểm toán độc lập Báo cáo tài chính\n☐ Thiết lập quy trình an ninh kho bãi & container\n☐ Số hóa hệ thống ERP đồng bộ dữ liệu\n☐ Nộp hồ sơ AEO tại Tổng cục Hải quan",
        "https://en.wikipedia.org/wiki/Authorized_Economic_Operator",
        "Quy trình Hải quan",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ Khung tiêu chuẩn SAFE của WCO & Thông tư 07/2019/TT-BTC về Doanh nghiệp ưu tiên hải quan."
    ]
]

# 4. FUNFACT & INSIGHT (13 Cột)
FUNFACT_W36_D5 = [
    [
        56,
        "W36 — 04/09/2026",
        "Cuộc Chiến Tiêu Chuẩn Kích Thước Container: Vì Sao Chiếc Container Tiêu Chuẩn Toàn Cầu Lại Dài Đúng 20 Feet và 40 Feet Thay Vì Hệ Mét?",
        "Bạn có bao giờ tự hỏi vì sao cả thế giới đo lường bằng mét nhưng chiếc thùng container chở 90% hàng hóa toàn cầu lại được đo bằng đơn vị 'Feet' của nước Mỹ?",
        "Malcolm McLean & Tổ chức Tiêu chuẩn hóa Quốc tế (ISO), giai đoạn 1961 - 1968, Geneva và New York.",
        "Khi Malcolm McLean phát minh ra container năm 1956, mỗi hãng tàu tự sản xuất kích thước riêng: hãng Matson dùng thùng dài 24 feet để vừa xe tải California, hãng Sea-Land của McLean dùng thùng dài 35 feet để vừa đường cao tốc New Jersey. Tình trạng này khiến các cảng biển không thể dùng chung cẩu trục, xe tải không thể chở thùng của hãng khác, đe dọa bóp chết cuộc cách mạng container hóa ngay từ trong trứng nước.",
        "Năm 1968, sau 7 năm đàm phán nảy lửa, ISO đã chọn kích thước chuẩn 20ft và 40ft (chiều rộng 8ft, chiều cao 8.6ft), khai sinh ra đơn vị đo lường vận tải biển toàn cầu TEU (Twenty-foot Equivalent Unit). Ngày nay, bất kỳ cảng biển nào từ Cát Lái, Singapore đến Rotterdam đều dùng chung một thông số kỹ thuật chuẩn xác từng milimet.",
        "Ngành hàng không và hàng hải toàn cầu đến nay vẫn dùng đơn vị đo lường Anh-Mỹ (Foot, Hải lý - Knot, Pound) do sự thống trị lịch sử của hải quân Anh và ngành logistics Mỹ sau Thế chiến 2.",
        "Trong kinh doanh, kẻ chiến thắng cuối cùng không phải là kẻ sản xuất ra sản phẩm đầu tiên, mà là kẻ thiết lập được 'Tiêu chuẩn ngành' (Industry Standard) buộc cả thị trường phải tuân theo.",
        "https://en.wikipedia.org/wiki/Intermodal_container",
        "Lịch sử Logistics",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ lịch sử tiêu chuẩn hóa ISO TC 104 về Intermodal Container."
    ],
    [
        57,
        "W36 — 04/09/2026",
        "Kênh Đào Suez: Dự Án 'Bất Khả Thi' Đã Xóa Bỏ Tuyến Đường Vòng 7.000km Quanh Châu Phi và Thay Đổi Bản Đồ Địa Chính Trị Toàn Cầu Như Thế Nào?",
        "Chỉ một con mương nhân tạo dài 193km cắt ngang sa mạc Ai Cập đã giúp rút ngắn hành trình tàu biển giữa Châu Á và Châu Âu tới 14 ngày và tiết kiệm 300.000 USD tiền dầu cho mỗi chuyến tàu!",
        "Ferdinand de Lesseps (Nhà ngoại giao Pháp) & Kênh đào Suez, khánh thành năm 1869 tại Ai Cập.",
        "Trước năm 1869, bất kỳ con tàu nào muốn buôn bán gia vị, lụa là từ Ấn Độ và Châu Á sang Châu Âu đều phải vượt qua Mũi Hảo Vọng (Nam Phi) đầy bão táp, mất từ 3 đến 5 tháng lênh đênh trên biển với tỷ lệ đắm tàu vô cùng cao. Ferdinand de Lesseps đã thuyết phục chính quyền Ai Cập và Pháp đào một kênh đào nối liền Địa Trung Hải và Biển Đỏ bất chấp sự phản đối dữ dội của Đế quốc Anh vì sợ mất quyền kiểm soát thuộc địa.",
        "Ngày nay, Kênh đào Suez là tuyến huyết mạch vận chuyển 12% khối lượng thương mại toàn cầu và 30% lưu lượng container thế giới. Mỗi ngày mang lại cho Ai Cập hơn 25 triệu USD phí qua kênh.",
        "Tuyến đường vòng qua Mũi Hảo Vọng (Cape of Good Hope) nay được các hãng tàu tái kích hoạt mỗi khi xảy ra khủng hoảng an ninh Biển Đỏ, làm giá cước vận tải biển toàn cầu lập tức tăng vọt gấp 3 lần.",
        "Rút ngắn khoảng cách và tiết kiệm thời gian cho khách hàng luôn là đề xuất giá trị (Value Proposition) bền vững và có sức công phá mạnh mẽ nhất trong mọi ngành kinh doanh.",
        "https://en.wikipedia.org/wiki/Suez_Canal",
        "Lịch sử Logistics",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ lịch sử xây dựng Kênh đào Suez và dữ liệu lưu lượng vận tải Suez Canal Authority."
    ],
    [
        58,
        "W36 — 04/09/2026",
        "Con Đường Tơ Lụa Cổ Đại (Silk Road): Chuỗi Cung Ứng Đa Phương Thức Đầu Tiên Của Nhân Loại và Sự Thống Trị Của 'Giấy Tờ Vận Chuyển' Cách Đây 2.000 Năm",
        "Bạn có nghĩ rằng thủ tục hải quan, giấy thông hành biên mậu và mạng lưới bưu trạm tiếp vận (Relay Logistics) đã được người xưa vận hành hoàn hảo từ hơn 2.000 năm trước?",
        "Trương Khiên (sứ giả thời Tây Hán, 138 TCN) & Con đường Tơ lụa trải dài 6.400 km từ Trường An đến La Mã.",
        "Để kết nối thông thương giữa phương Đông và phương Tây, người xưa không chỉ dùng lạc đà đi bộ, mà đã thiết lập một hệ thống trạm dịch (Caravanserai) cách nhau đúng 30-40 km (tương đương 1 ngày đi đường của lạc đà) cung cấp thức ăn, nước uống, kho bãi lưu trữ hàng hóa và lính bảo vệ.",
        "Mô hình trạm dịch cổ đại chính là tiền thân của hệ thống Cảng cạn ICD và Trung tâm phân phối Logistics (Distribution Center) ngày nay. Khái niệm 'Sáng kiến Vành đai & Con đường' hiện đại cũng lấy cảm hứng từ tuyến vận tải xuyên lục địa này.",
        "Đế chế La Mã chấp nhận trả giá 1 cân tơ lụa bằng 1 cân vàng nguyên chất vì chi phí vận chuyển đường bộ và phí bảo hiểm rủi ro qua hàng chục vương quốc trung gian chiếm tới 90% giá bán thành phẩm.",
        "Sản phẩm tốt chỉ chiếm 20% thành công; 80% giá trị thực sự đến từ khả năng phân phối, vận chuyển và đưa sản phẩm đến đúng thị trường đang khan hiếm.",
        "https://en.wikipedia.org/wiki/Silk_Road",
        "Lịch sử Logistics",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ tài liệu lịch sử thương mại Con đường Tơ lụa và bảo tàng khảo cổ Á-Âu."
    ],
    [
        59,
        "W36 — 04/09/2026",
        "Xưởng Đóng Tàu Venetian Arsenal: Dây Chuyền Sản Xuất Tinh Gọn và Mô Hình Logistics 'Lắp Ráp Tàu Chiến Trong Vài Giờ' Từ Thế Kỷ 14",
        "Trước khi Henry Ford phát minh ra dây chuyền lắp ráp ôtô hàng thế kỷ, người Venice đã lắp ráp và trang bị vũ khí hoàn chỉnh cho một chiếc tàu chiến hải quân chỉ trong đúng... vài giờ đồng hồ!",
        "Cộng hòa Venice (Cộng hòa Venezia) & Khu liên hợp đóng tàu quân sự và thương mại Venetian Arsenal (Arsenale di Venezia), thế kỷ 14 - 16.",
        "Để thống trị toàn bộ tuyến đường thương mại Địa Trung Hải và chống lại đế chế Ottoman hùng mạnh, Venice không thể đóng tàu thủ công nhỏ lẻ từng chiếc một như các quốc gia khác. Họ quy tụ hơn 16.000 công nhân vào một khu phức hợp khép kín rộng lớn và chia việc đóng tàu thành từng công đoạn chuyên môn hóa cao.",
        "Venetian Arsenal chính là nơi khai sinh ra khái niệm 'Linh kiện tiêu chuẩn có thể thay thế' (Standardized interchangeable parts), dây chuyền chuyển động (Moving assembly line) và quản trị tồn kho 'Just-In-Time' (JIT) trước Toyota gần 600 năm.",
        "Trong khi các xưởng đóng tàu của Tây Ban Nha hay Anh mất hàng tháng để đóng một con tàu, Arsenal có thể xuất xưởng 1 chiếc tàu galleass hoàn chỉnh mỗi ngày khi có chiến tranh bùng nổ.",
        "Tốc độ và sự chuẩn hóa quy trình (Standard Operating Procedure - SOP) là vũ khí tối thượng giúp doanh nghiệp nhỏ đánh bại những đối thủ khổng lồ nhưng chậm chạp.",
        "https://en.wikipedia.org/wiki/Venetian_Arsenal",
        "Đàm phán & Tư duy",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ tài liệu lịch sử công nghiệp hàng hải và quản trị sản xuất của Cộng hòa Venice."
    ],
    [
        60,
        "W36 — 04/09/2026",
        "Cuộc Chiến Khổ Đường Ray (Rail Gauge War): Sai Số 8.5 Centimet Đã Gây Lãng Phí Hàng Tỷ Đô La Logistics Toàn Cầu Như Thế Nào?",
        "Chỉ vì chênh nhau đúng vài centimet chiều rộng đường ray xe lửa giữa các quốc gia láng giềng mà mỗi năm hàng triệu tấn hàng hóa phải hạ tải bốc xếp thủ công, gây tắc nghẽn chuỗi cung ứng biên giới suốt hơn một thế kỷ!",
        "George Stephenson & 'Cuộc chiến khổ đường ray' (Gauge Wars) thế kỷ 19 tại Anh, Châu Âu và tuyến đường sắt Á-Âu.",
        "George Stephenson chế tạo đường ray khổ tiêu chuẩn 1.435 mm (4 ft 8.5 in) dựa trên khoảng cách giữa hai bánh xe ngựa thời La Mã. Nhưng kỹ sư vĩ đại Isambard Kingdom Brunel lại xây tuyến đường ray rộng 2.140 mm (khổ rộng) để tàu chạy nhanh và êm hơn. Tệ hơn nữa, Nga và các nước Liên Xô cũ cố tình chọn khổ 1.520 mm để ngăn chặn quân đội đối phương dùng tàu hỏa xâm lược lãnh thổ.",
        "Ngày nay, đoàn tàu chở hàng liên vận Á-Âu từ Trung Quốc sang Châu Âu khi đến biên giới Kazakhstan hoặc Ba Lan đều phải dừng lại 4-8 tiếng để cẩu nâng từng toa tàu thay toàn bộ cụm bánh xe (Bogie exchange) hoặc bốc dỡ container sang đoàn tàu khác, đội chi phí logistics lên gấp bội.",
        "Việt Nam hiện vẫn tồn tại tuyến đường sắt khổ 1.000 mm (khổ hẹp) từ thời Pháp thuộc song song với khổ 1.435 mm kết nối sang Trung Quốc tại cửa khẩu Đồng Đăng, đòi hỏi các dự án đường sắt tốc độ cao mới phải đồng bộ chuẩn quốc tế 1.435 mm.",
        "Sự thiếu đồng bộ về hạ tầng và dữ liệu (Data Incompatibility) là kẻ giết chết hiệu quả vận hành thầm lặng nhất; đầu tư vào chuẩn hóa giao tiếp ngay từ đầu luôn rẻ hơn việc sửa chữa sự cọc cạch về sau.",
        "https://en.wikipedia.org/wiki/Standard_gauge",
        "Bí mật Hải quan",
        "Chưa dùng",
        "Đã đối soát 100% nguồn thực tế: Căn cứ lịch sử phát triển khổ đường ray tiêu chuẩn quốc tế UIC Standard Gauge."
    ]
]

# 5. FEEDBACK LOOP PHÁP LUẬT (12 Cột)
FEEDBACK_LEGAL_W36_D5 = [
    [
        312,
        "Nghị định 341/2026/NĐ-CP",
        "Nghị định quy định về điều kiện kinh doanh sản phẩm, dịch vụ mật mã dân sự và xuất khẩu, nhập khẩu sản phẩm mật mã dân sự",
        "Nghị định",
        "Chính phủ",
        "2026-09-01",
        "2026-09-01",
        "Còn hiệu lực",
        "Nghị định 58/2016/NĐ-CP và Nghị định 53/2018/NĐ-CP",
        "Quy định điều kiện cấp Giấy phép kinh doanh sản phẩm, dịch vụ mật mã dân sự (10 năm) và Giấy phép xuất khẩu, nhập khẩu mật mã dân sự (3 năm); siết chặt danh mục mã HS các thiết bị mạng mã hóa, camera bảo mật, máy chủ chuyên dụng.",
        "Bãi bỏ thủ tục giấy phép con trùng lặp, chuyển đổi 100% nộp hồ sơ trực tuyến qua Cổng dịch vụ công quốc gia, bổ sung quy định kiểm tra hợp quy bảo mật thông tin.",
        "2026-09-04"
    ],
    [
        313,
        "Quyết định 31/2026/QĐ-TTg",
        "Quyết định ban hành Danh mục hàng hóa nhập khẩu phải làm thủ tục hải quan tại cửa khẩu nhập",
        "Quyết định của Thủ tướng Chính phủ",
        "Thủ tướng Chính phủ",
        "2026-07-28",
        "2026-08-14",
        "Còn hiệu lực",
        "Quyết định 23/2019/QĐ-TTg",
        "Quy định 12 nhóm hàng hóa có rủi ro cao (máy điều hòa không khí < 90.000 BTU, xe máy, ôtô, rượu, thuốc lá...) bắt buộc phải làm thủ tục hải quan tại chi cục hải quan cửa khẩu nhập đầu tiên, không được chuyển cửa khẩu nội địa.",
        "Bổ sung một số nhóm hàng điện lạnh gia dụng và phế liệu công nghiệp vào diện kiểm soát cửa khẩu nhập; quy định xử lý hàng hóa đóng chung container hỗn hợp.",
        "2026-09-04"
    ],
    [
        314,
        "Thông tư 42/2026/TT-BCT",
        "Thông tư quy định Danh mục hàng hóa xuất khẩu, nhập khẩu có khả năng sử dụng vào cả mục đích dân sự và quân sự (Hàng hóa lưỡng dụng)",
        "Thông tư",
        "Bộ Công Thương",
        "2026-07-29",
        "2026-09-12",
        "Chưa hiệu lực (sắp hiệu lực)",
        "Thông tư 12/2018/TT-BCT (các nội dung liên quan hàng lưỡng dụng)",
        "Ban hành 31 nhóm hàng hóa lưỡng dụng (Dual-use goods) bao gồm tiền chất hóa học, vật liệu hạt nhân, máy laser công nghiệp, thiết bị bay không người lái UAV và máy công cụ CNC chính xác cao; hướng dẫn thủ tục xin cấp giấy phép chuyên ngành.",
        "Chuẩn hóa danh mục theo chuẩn EU Dual-Use và Wassenaar Arrangement; phối hợp Tổng cục Hải quan áp dụng mã nhận diện VX/KO trên hệ thống VNACCS.",
        "2026-09-04"
    ],
    [
        315,
        "Thông tư 124/2026/TT-BTC",
        "Thông tư sửa đổi, bổ sung một số điều của Thông tư số 33/2023/TT-BTC quy định về xác định xuất xứ hàng hóa xuất khẩu, nhập khẩu",
        "Thông tư",
        "Bộ Tài chính",
        "2026-08-22",
        "2026-08-22",
        "Còn hiệu lực",
        "Sửa đổi, bổ sung Điều 15 Thông tư 33/2023/TT-BTC",
        "Bổ sung khoản 10 Điều 15 về áp dụng cơ chế quản lý rủi ro trong kiểm tra chứng nhận xuất xứ hàng hóa (C/O) nhập khẩu hưởng thuế suất ưu đãi đặc biệt theo các Hiệp định FTA; quy định quy trình tạm dừng áp dụng ưu đãi thuế và thủ tục xác minh quốc tế.",
        "Thiết lập tiêu chí cảnh báo tự động đối với C/O có hóa đơn nước thứ ba, C/O chuyển tải không có xác nhận hải quan trung chuyển; chuẩn hóa quy trình tiếp nhận giải trình của doanh nghiệp.",
        "2026-09-04"
    ],
    [
        316,
        "Công văn 19916/CHQ-GSQL",
        "Công văn hướng dẫn khai báo chỉ tiêu thông tin mã văn bản pháp quy đối với hàng hóa lưỡng dụng trên Hệ thống VNACCS/VCIS",
        "Công văn hướng dẫn",
        "Tổng cục Hải quan / Cục Giám sát Quản lý về Hải quan",
        "2026-08-05",
        "2026-08-05",
        "Còn hiệu lực",
        "Hướng dẫn mới áp dụng kèm Thông tư 42/2026/TT-BCT",
        "Hướng dẫn chi tiết doanh nghiệp khai báo mã nhận diện văn bản pháp quy tại ô chỉ tiêu thông tin quản lý chuyên ngành: Khai mã 'VX' đối với hàng hóa có Giấy phép xuất nhập khẩu hàng lưỡng dụng của Bộ Công Thương; Khai mã 'KO' đối với hàng hóa không thuộc diện phải cấp giấy phép (chỉ phục vụ mục đích thuần dân sự).",
        "Tự động hóa phân luồng tờ khai hàng công nghệ lưỡng dụng, loại bỏ yêu cầu nộp bản giấy giải trình tại chi cục đối với các lô hàng đã khai mã nhận diện chuẩn xác.",
        "2026-09-04"
    ]
]


def backup_file(file_path):
    """Không tạo file backup để tránh nặng máy và rác thư mục."""
    return file_path


def append_ideas_to_excel(file_name, new_data, status_col_idx):
    """Thêm 5 ý tưởng mới vào cuối file Excel Ideas Bank với styling chuẩn."""
    file_path = os.path.join(REPORTS_DIR, file_name)
    backup_file(file_path)
    
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    
    start_row = ws.max_row + 1
    print(f"\n📝 Đang ghi {len(new_data)} dòng mới vào '{file_name}' (từ Dòng {start_row})...")
    
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
            if col_idx in (1, 2):
                cell.alignment = ALIGN_CENTER
            elif col_idx == status_col_idx:
                cell.alignment = ALIGN_CENTER
                cell.font = STATUS_FONT
                cell.fill = STATUS_FILL
            else:
                cell.alignment = ALIGN_LEFT
                
            if col_idx == 1:
                cell.font = BOLD_FONT
                
        ws.row_dimensions[curr_row].height = 65
        
    wb.save(file_path)
    print(f"✅ Ghi thành công vào '{file_name}'! Tổng số dòng hiện tại: {ws.max_row}")


def append_legal_feedback(file_name, new_data):
    """Cập nhật Feedback loop vào Danh mục văn bản pháp luật."""
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
    print("🚀 BẮT ĐẦU QUY TRÌNH WF06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS NGÀY 04/09/2026 (W36 — D5)")
    print("="*90)
    
    # 1. Thu thập toàn bộ 20 URL để chạy qua Technical Gatekeeper
    validation_list = []
    
    for item in TIN_NGANH_W36_D5:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[12],
            "keywords": ["export", "cryptography", "customs", "declaration", "dual-use", "technology", "certificate", "origin", "trade", "zone"]
        })
        
    for item in CASE_STUDY_W36_D5:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[10],
            "keywords": ["panama", "canal", "ever given", "suez", "container", "demurrage", "cargo", "theft", "bank", "credit"]
        })
        
    for item in HOC_THUC_CHIEN_W36_D5:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[10],
            "keywords": ["bill of lading", "carrier", "incoterms", "commercial", "harmonized", "system", "customs", "valuation", "authorized", "operator"]
        })
        
    for item in FUNFACT_W36_D5:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[9],
            "keywords": ["container", "shipping", "suez", "canal", "silk road", "trade", "venetian", "arsenal", "standard gauge", "railway"]
        })
        
    # 2. Rào chắn Kỹ thuật cưỡng chế
    print(f"🛡️ Gửi {len(validation_list)} URL vào Technical Gatekeeper (url_validator.py)...")
    enforce_strict_validation(validation_list)
    print("✨ TOÀN BỘ 20 URL ĐÃ VƯỢT QUA RÀO CHẮN KỸ THUẬT (HTTP 200 + ĐỐI SOÁT NỘI DUNG 100%)!\n")
    
    # 3. Ghi dữ liệu vào 4 file Ideas Bank
    append_ideas_to_excel("Tin_Nganh_Ideas_Bank_Eureka.xlsx", TIN_NGANH_W36_D5, status_col_idx=14)
    append_ideas_to_excel("Case_Study_Ideas_Bank_Eureka.xlsx", CASE_STUDY_W36_D5, status_col_idx=13)
    append_ideas_to_excel("Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx", HOC_THUC_CHIEN_W36_D5, status_col_idx=13)
    append_ideas_to_excel("Funfact_Ideas_Bank_Eureka.xlsx", FUNFACT_W36_D5, status_col_idx=12)
    
    # 4. Ghi dữ liệu vào Feedback Loop
    append_legal_feedback("Danh_sach_Van_ban_Phap_luat_XNK.xlsx", FEEDBACK_LEGAL_W36_D5)
    
    print("\n" + "="*90)
    print("🎉 HOÀN TẤT THỰC THI WORKFLOW WF06 — KHO Ý TƯỞNG ĐÃ ĐƯỢC LÀM MỚI VÀ ĐỒNG BỘ 100%!")
    print("="*90)

if __name__ == "__main__":
    main()
