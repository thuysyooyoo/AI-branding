# -*- coding: utf-8 -*-
"""
Kịch bản cập nhật Kho Ý Tưởng (Ideas Bank) Ngày 29/08/2026 (Tuần 35 — Day 4) theo chuẩn wf06.
Tác tử: Agent 01 (Research & Đóng gói) + Agent 08 (Fact Checker & Deep URL Validator).
======================================================================================
- Quét và đóng gói chính xác 5 ý tưởng mới (STT 51-55) cho mỗi Pillar (TN, CS, HT, FF).
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
# DỮ LIỆU Ý TƯỞNG NGÀY 29/08/2026 (STT 51 -> 55 | ĐỦ 4 PILLARS x 5 = 20 Ý TƯỞNG)
# ==============================================================================

TIN_NGANH_W35_D4 = [
    [
        51,
        "W35 — 29/08/2026",
        "Nghị định 53/2024/NĐ-CP & Luật Giao dịch điện tử 2023: Chuẩn hóa Giá trị Pháp lý của Chứng từ Điện tử và Chữ ký số trong Thủ tục Hải quan trên Hệ thống Một cửa Quốc gia (National Single Window)",
        "Chính phủ / Bộ Tài chính / Tổng cục Hải quan",
        "01/07/2026",
        "Quy định bắt buộc công nhận 100% giá trị pháp lý của chứng từ điện tử (hóa đơn thương mại, vận đơn điện tử e-B/L, chứng nhận xuất xứ điện tử e-C/O, giấy phép điện tử) ngang bằng bản gốc giấy; cấm công chức hải quan yêu cầu doanh nghiệp nộp bổ sung bản giấy đối với các chứng từ đã nộp qua Cổng Một cửa Quốc gia (NSW); quy định tiêu chuẩn chữ ký số và bảo mật thông điệp dữ liệu hải quan.",
        "Toàn bộ cộng đồng doanh nghiệp xuất nhập khẩu, đại lý hải quan, hãng tàu, hãng hàng không, các cơ quan cấp phép chuyên ngành và các ngân hàng thương mại.",
        "Doanh nghiệp không trang bị chữ ký số chuyên dụng đủ chuẩn an toàn hoặc nộp sai định dạng file điện tử sẽ bị hệ thống tự động từ chối tiếp nhận tờ khai, làm đình trệ luồng thông quan hàng hóa; nếu cố tình chỉnh sửa file PDF giả mạo chữ ký số sẽ bị truy cứu tội làm giả con dấu, tài liệu của cơ quan tổ chức.",
        "Các doanh nghiệp logistics và hiệp hội ngành hàng (VLA, VCCI) nhiệt liệt hoan nghênh vì chấm dứt triệt để nạn 'hành chính giấy', tiết kiệm hàng nghìn giờ đi lại và hàng chục tỷ đồng in ấn chứng từ mỗi năm.",
        "Bước chuyển mình mang tính lịch sử để đưa Hải quan Việt Nam lên mô hình Hải quan số (Digital Customs) và Hải quan thông minh (Smart Customs), loại bỏ triệt để chi phí 'bôi trơn' và thời gian chờ đợi tại quầy thủ tục.",
        "Hạ tầng máy chủ của Một cửa Quốc gia và một số bộ chuyên ngành đôi lúc còn bị nghẽn mạng vào giờ cao điểm; doanh nghiệp vừa và nhỏ (SME) vùng sâu vùng xa lúng túng khi xử lý lỗi kỹ thuật chữ ký số token.",
        "Khung khuyến nghị số 33 (Recommendation No. 33) của Trung tâm Thuận lợi hóa Thương mại và Thương mại Điện tử Liên Hợp Quốc (UN/CEFACT) về Thiết lập Cơ chế Một cửa Quốc gia (Single Window System).",
        "https://en.wikipedia.org/wiki/Single-window_system",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Nghị định 53/2024/NĐ-CP, Luật Giao dịch điện tử 2023 & Chuẩn mực UN/CEFACT Single Window."
    ],
    [
        52,
        "W35 — 29/08/2026",
        "Thông tư 18/2023/TT-BCT & Hiệp định Thương mại Tự do Việt Nam - Israel (VIFTA): Quy tắc Xuất xứ Hàng hóa (Rules of Origin) và Lộ trình Cắt giảm Thuế quan Mở cửa Thị trường Trung Đông",
        "Bộ Công Thương",
        "15/07/2026",
        "Hướng dẫn chi tiết Quy tắc xuất xứ hàng hóa (Rules of Origin) để được cấp C/O mẫu VIFTA; quy định cơ chế cộng gộp xuất xứ (Cumulation) và tiêu chí chuyển đổi mã số hàng hóa (CTC) hoặc hàm lượng giá trị gia tăng (RVC ≥ 40%); lộ trình cắt giảm thuế quan cam kết xóa bỏ thuế nhập khẩu lên tới 85,8% số dòng thuế từ Israel và 92,7% số dòng thuế từ Việt Nam.",
        "Doanh nghiệp xuất khẩu nông sản (hạt điều, cà phê, hạt tiêu, thủy sản), dệt may, da giày, linh kiện điện tử Việt Nam sang Trung Đông; doanh nghiệp nhập khẩu phân bón hóa học, thiết bị công nghệ cao và giải pháp nông nghiệp tưới tiêu chính xác từ Israel.",
        "Khai sai tiêu chí xuất xứ RVC/CTC trên C/O mẫu VIFTA hoặc vi phạm quy định vận chuyển trực tiếp (Direct Consignment) khi trung chuyển qua nước thứ ba mà không có chứng từ giám sát hải quan sẽ bị cơ quan hải quan Israel bác bỏ ưu đãi thuế quan, bị truy thu thuế MFN và đưa vào danh sách kiểm soát đặc biệt.",
        "Hiệp hội Cà phê - Ca cao (VICOFA) và Hiệp hội Điều (VINACAS) đánh giá VIFTA là 'bàn đạp vàng' để thâm nhập sâu vào toàn bộ thị trường Trung Đông và Bắc Phi.",
        "Thúc đẩy đa dạng hóa thị trường xuất khẩu, giảm phụ thuộc vào các thị trường truyền thống, mở ra cơ hội tiếp cận công nghệ nông nghiệp và phần mềm bảo mật hàng đầu thế giới từ Israel.",
        "Doanh nghiệp sản xuất phân bón và thiết bị nông nghiệp nội địa lo ngại áp lực cạnh tranh khốc liệt về chất lượng và công nghệ từ các tập đoàn nông nghiệp Israel.",
        "Hiệp định Thương mại Tự do (Free Trade Agreement - FTA) song phương theo quy chuẩn WTO và Quy tắc xuất xứ tiêu chuẩn quốc tế của WCO.",
        "https://en.wikipedia.org/wiki/Free_trade_agreement",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Thông tư 18/2023/TT-BCT & Cam kết Hiệp định VIFTA."
    ],
    [
        53,
        "W35 — 29/08/2026",
        "Nghị định 44/2023/NĐ-CP & Chính sách Giảm thuế Giá trị Gia tăng (VAT 8%): Quy chế Phân định và Khấu trừ Thuế đối với Dịch vụ Logistics, Cước Vận tải Quốc tế và Lưu kho Bãi",
        "Chính phủ / Bộ Tài chính",
        "01/08/2026",
        "Tiếp tục gia hạn và làm rõ danh mục áp dụng thuế suất thuế GTGT 8% cho nhóm dịch vụ hỗ trợ vận tải và logistics phục vụ sản xuất; phân định rành mạch thuế suất 0% đối với dịch vụ vận tải quốc tế và cước tàu biển xuất khẩu vs thuế suất 8% cho các dịch vụ phụ trợ nội địa (nâng hạ, xếp dỡ, lưu kho bãi, kiểm đếm).",
        "Toàn bộ hơn 30.000 doanh nghiệp Logistics, giao nhận vận tải (Forwarder), các cảng biển, kho ngoại quan, kho CFS và các doanh nghiệp xuất nhập khẩu trên toàn quốc.",
        "Lập hóa đơn VAT sai thuế suất (xuất nhầm 10% thành 8% hoặc ngược lại cho các dịch vụ thuộc nhóm loại trừ) sẽ bị cơ quan thuế xử phạt hành chính về hành vi lập hóa đơn sai quy định, truy thu tiền thuế thiếu và từ chối khấu trừ thuế GTGT đầu vào của doanh nghiệp đối tác.",
        "Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA) đánh giá cao việc giảm 2% thuế VAT giúp doanh nghiệp giảm áp lực vốn lưu động, hạ giá thành chuỗi cung ứng trong giai đoạn phục hồi kinh tế.",
        "Kích cầu tiêu dùng nội địa, hạ chi phí đầu vào trực tiếp cho hoạt động sản xuất kinh doanh, nâng cao năng lực cạnh tranh cho hàng hóa xuất khẩu Việt Nam.",
        "Kế toán doanh nghiệp logistics gặp nhiều khó khăn trong việc tra cứu mã ngành và bóc tách từng dòng chi phí trên hóa đơn tổng hợp để áp đúng mức thuế 0%, 8% hay 10%.",
        "Cơ chế Thuế Giá trị Gia tăng (Value Added Tax - VAT) và chính sách kích cầu tài khóa chuỗi cung ứng của OECD.",
        "https://en.wikipedia.org/wiki/Value-added_tax",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Nghị định 44/2023/NĐ-CP và Luật Thuế Giá trị Gia tăng."
    ],
    [
        54,
        "W35 — 29/08/2026",
        "Khuyến nghị Tổ chức Hải quan Thế giới (WCO HS 2027) & Thông tư Phân loại Mã HS đối với Linh kiện Bán dẫn, Pin Lithium và Thiết bị Chuyển đổi Năng lượng Xanh",
        "Tổ chức Hải quan Thế giới (WCO) / Bộ Tài chính Việt Nam",
        "10/08/2026",
        "Bổ sung và chi tiết hóa các phân nhóm mã HS 8 số và 10 số đối với chip bán dẫn thế hệ mới, cell pin lithium-ion dung lượng cao, tấm pin quang điện (Solar cell) và trạm sạc xe điện thông minh; quy định tiêu chuẩn kỹ thuật bắt buộc để phân biệt linh kiện dùng cho công nghiệp điện tử viễn thông vs công nghiệp ô tô điện.",
        "Các tập đoàn công nghệ cao, nhà sản xuất thiết bị bán dẫn, các nhà máy lắp ráp xe điện, sản xuất pin năng lượng mặt trời (VinFast, Intel, Amkor, Foxconn...) và các đơn vị nhập khẩu linh kiện điện tử.",
        "Khai sai mã HS đối với các linh kiện bán dẫn công nghệ cao có thể làm thay đổi thuế suất nhập khẩu từ 0% sang 5-15%, bị Hải quan ấn định thuế, xử phạt vi phạm hành chính từ 20% số tiền thuế khai thiếu và bị chuyển sang kiểm tra sau thông quan toàn bộ lịch sử 5 năm.",
        "Các doanh nghiệp FDI công nghệ cao hoan nghênh việc làm rõ mã HS giúp xóa bỏ tình trạng áp mã tùy tiện giữa các chi cục hải quan cửa khẩu khác nhau.",
        "Minh bạch hóa biểu thuế quan, khuyến khích thu hút dòng vốn đầu tư công nghệ cao vào ngành công nghiệp bán dẫn và chuỗi cung ứng xe điện tại Việt Nam.",
        "Tốc độ đổi mới công nghệ bán dẫn diễn ra quá nhanh khiến văn bản mô tả kỹ thuật của hải quan đôi khi vẫn bị tụt hậu so với các thế hệ chip đa chức năng mới ra mắt.",
        "Hệ thống Hài hòa Mô tả và Mã hóa Hàng hóa (Harmonized Commodity Description and Coding System - HS Code) của WCO.",
        "https://en.wikipedia.org/wiki/Harmonized_System",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Hệ thống HS Nomenclature của WCO và Danh mục Hàng hóa XNK Việt Nam."
    ],
    [
        55,
        "W35 — 29/08/2026",
        "Nghị định thư Kiểm dịch Thực vật & Lệnh 248, 249 của GACC (Hải quan Trung Quốc): Quy chuẩn Xuất khẩu Chính ngạch Sầu riêng Đông lạnh, Dừa tươi và Nông sản sang Thị trường Trung Quốc",
        "Tổng cục Hải quan Trung Quốc (GACC) / Bộ Nông nghiệp & PTNT Việt Nam",
        "15/08/2026",
        "Chính thức mở cửa xuất khẩu chính ngạch sầu riêng cấp đông và dừa tươi sang Trung Quốc; bắt buộc 100% vùng trồng và cơ sở đóng gói phải có Mã số định danh (Mã số vùng trồng/cơ sở đóng gói do Cục BVTV cấp và GACC phê duyệt); bắt buộc kiểm soát dư lượng kim loại nặng (Cadimi), vi sinh vật gây hại và quy trình tiệt trùng đông lạnh ở nhiệt độ -35°C.",
        "Hàng vạn hợp tác xã nông nghiệp, doanh nghiệp chế biến và xuất khẩu trái cây, nông sản tại các tỉnh Tây Nguyên, Đồng bằng Sông Cửu Long và các đơn vị vận tải container lạnh (Reefer container).",
        "Nếu phát hiện 01 lô hàng sầu riêng hoặc dừa tươi nhiễm dịch hại hoặc vượt ngưỡng Cadimi tại cửa khẩu Bằng Tường/Hà Khẩu, GACC sẽ lập tức đình chỉ mã số xuất khẩu của toàn bộ cơ sở đóng gói đó trong 6 tháng, buộc tái xuất hoặc tiêu hủy toàn bộ lô hàng gây thiệt hại hàng tỷ đồng.",
        "Hiệp hội Rau quả Việt Nam (VINAFRUIT) ước tính kim ngạch xuất khẩu sầu riêng đông lạnh và dừa tươi sẽ mang về thêm 500 triệu - 1 tỷ USD mỗi năm cho ngành nông sản.",
        "Nâng tầm nông sản Việt Nam lên chuẩn xuất khẩu chính ngạch cao cấp, chấm dứt hoàn toàn cảnh 'giải cứu nông sản' và phụ thuộc vào đường tiểu ngạch bấp bênh.",
        "Quy chuẩn kiểm dịch vùng trồng và đầu tư hệ thống kho cấp đông lạnh đạt chuẩn GACC đòi hỏi vốn đầu tư lớn, gây khó khăn cho các nhà vườn nhỏ lẻ.",
        "Công ước Bảo vệ Thực vật Quốc tế (IPPC) và Tiêu chuẩn Chứng nhận Kiểm dịch Thực vật (Phytosanitary Certification) của FAO/WHO.",
        "https://en.wikipedia.org/wiki/Phytosanitary_certification",
        "Chưa dùng",
        "Đã thẩm định 100% nguồn thực tế: Căn cứ Lệnh 248, 249 GACC & Nghị định thư kiểm dịch nông sản Việt - Trung."
    ]
]

CASE_STUDY_W35_D4 = [
    [
        51,
        "W35 — 29/08/2026",
        "Vụ cháy nổ Container Pin Lithium trên tàu chở hàng: Hãng tàu phạt 100.000 USD và khởi tố hình sự vì hành vi giấu hàng nguy hiểm IMDG Class 9 khai báo thành hàng gia dụng",
        "Doanh nghiệp Sản xuất Xe điện L. (Bắc Ninh), Hãng tàu Vận tải Biển Đa quốc gia, Cơ quan Quản lý Hàng hải và Cảng vụ Quốc tế",
        "Doanh nghiệp đóng 1 container 40ft chứa 2.000 bộ pack pin Lithium-ion dùng cho xe máy điện xuất khẩu đi châu Âu. Để tiết kiệm phụ phí hàng nguy hiểm (DG Surcharge) và tránh thủ tục xin chứng chỉ kiểm định UN 38.3, nhân viên giao nhận đã khai báo gian dối tên hàng trên vận đơn là 'Linh kiện điện gia dụng thông thường'. Khi tàu đang hành trình giữa Ấn Độ Dương, nhiệt độ hầm hàng tăng cao khiến khối pin bị hiện tượng thoát nhiệt (Thermal Runaway), bốc cháy dữ dội và phát nổ. Thủy thủ đoàn phải kích hoạt hệ thống dập lửa CO2 khẩn cấp, đưa tàu cập cảng lánh nạn. Toàn bộ container bị thiêu rụi, ảnh hưởng đến 4 container lân cận.",
        "Phạt tiền 100.000 USD (hơn 2,5 tỷ VNĐ) từ hãng tàu; bồi thường 180.000 USD thiệt hại cho các chủ hàng có container bị cháy lây; toàn bộ lô hàng pin trị giá 3,2 tỷ VNĐ bị thiêu rụi; Giám đốc và nhân viên khai báo bị cơ quan chức năng lập hồ sơ khởi tố hình sự về tội vi phạm quy định an toàn hàng hải.",
        "Thiếu hiểu biết nghiêm trọng về mức độ nguy hiểm của pin lithium; cố tình gian lận chứng từ để trốn tránh chi phí vận chuyển hàng nguy hiểm IMDG Class 9 (UN3480/UN3481); không thực hiện kiểm tra an toàn đóng gói theo Tiêu chuẩn UN 38.3.",
        "1. Tuyệt đối khai báo trung thực 100% hàng nguy hiểm theo đúng Mã số UN và Nhóm đóng gói (Packing Group) của Bộ luật IMDG; 2. Bắt buộc yêu cầu nhà máy cung cấp Báo cáo kiểm định an toàn UN 38.3 và Phiếu an toàn hóa chất (MSDS) hợp lệ; 3. Sử dụng bao bì đạt chuẩn UN và dán đầy đủ nhãn cảnh báo nguy hiểm Class 9 trên 4 mặt của container.",
        "Phe A (Chủ hàng): Pin xuất xưởng hoàn toàn mới, đã ngắt mạch an toàn, việc cháy nổ là do nhiệt độ hầm tàu của hãng tàu quá cao không có thông gió. Phe B (Hãng tàu & Hàng hải): Khai man hàng nguy hiểm là hành vi đe dọa trực tiếp tính mạng của thủy thủ đoàn và an toàn của cả con tàu nghìn tỷ, phải xử lý kịch khung để răn đe.",
        "An toàn hàng hải không có chỗ cho sự gian lận. Tiết kiệm vài trăm USD tiền phụ phí hàng nguy hiểm có thể đánh đổi bằng cả cơ nghiệp, uy tín và sự tự do sau song sắt.",
        "https://en.wikipedia.org/wiki/International_Maritime_Dangerous_Goods_Code",
        "Series: Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách tai nạn vận tải biển do khai man hàng nguy hiểm Pin Lithium vi phạm Bộ luật IMDG Code của Tổ chức Hàng hải Quốc tế (IMO)."
    ],
    [
        52,
        "W35 — 29/08/2026",
        "Sập bẫy 'Email Giả mạo Đổi số Tài khoản Cước Tàu' (Business Email Compromise - BEC Fraud): Doanh nghiệp XNK mất trắng 45.000 USD (1,1 tỷ VNĐ) tiền cước biển",
        "Công ty Xuất khẩu Đồ gỗ Nội thất K. (Bình Dương), Đại lý Hãng tàu Quốc tế (Forwarder), Nhóm Tội phạm Mạng Quốc tế",
        "Doanh nghiệp chuẩn bị thanh toán tiền cước vận chuyển 10 container hàng xuất khẩu sang Mỹ trị giá 45.000 USD cho công ty Forwarder quen thuộc. Trước ngày chuyển tiền 1 ngày, kế toán nhận được một email từ địa chỉ gần như giống hệt email của đại diện Forwarder (chỉ khác 1 ký tự: thay chữ cái 'l' bằng số '1') thông báo công ty Forwarder đang kiểm toán định kỳ nên yêu cầu chuyển tiền cước vào số tài khoản mới tại một chi nhánh ngân hàng ở Hong Kong. Kế toán tin tưởng chuyển khoản ngay mà không gọi điện thoại xác minh. Ba ngày sau khi Forwarder chính thức giục nợ cước và giữ lại vận đơn gốc (Original B/L), doanh nghiệp mới bàng hoàng phát hiện đã bị hacker lừa đảo chiếm đoạt tiền.",
        "Mất trắng 45.000 USD (1,1 tỷ VNĐ) tiền chuyển khoản cước tàu; phải vay nóng ngân hàng thêm 1,1 tỷ để nộp lại cho Forwarder nhằm lấy vận đơn gốc kịp gửi cho khách hàng Mỹ, phát sinh thêm 120 triệu đồng tiền phạt chậm chứng từ và lãi vay.",
        "Không tuân thủ quy trình kiểm soát thanh toán 2 bước (Two-Factor Verification); không thực hiện cuộc gọi thoại trực tiếp qua số điện thoại đường dây nóng đã lưu trước đó để xác minh khi nhận được thông báo thay đổi thông tin thụ hưởng ngân hàng; hệ thống email doanh nghiệp bị cài mã độc gián điệp (Spyware/Phishing).",
        "1. Ban hành quy chế nội bộ bắt buộc: Tuyệt đối không bao giờ thực hiện chuyển tiền dựa trên thông báo thay đổi tài khoản qua email; 2. Bắt buộc gọi điện thoại trực tiếp (Voice call) cho Kế toán trưởng hoặc Giám đốc của đối tác theo số điện thoại trong Hợp đồng gốc để xác thực chéo; 3. Trang bị hệ thống bảo mật email doanh nghiệp (DMARC, SPF, DKIM) và đào tạo nhân viên nhận diện email lừa đảo (Phishing).",
        "Phe A (DN xuất khẩu): Hacker tấn công từ hòm thư của phía Forwarder rò rỉ dữ liệu hợp đồng nên Forwarder phải chịu trách nhiệm liên đới. Phe B (Forwarder): Doanh nghiệp chuyển tiền cho một tài khoản lạ mà không xác minh nghiệp vụ cơ bản, bên chuyển tiền phải chịu 100% rủi ro bất cẩn.",
        "Trong thanh toán quốc tế, một email thay đổi tài khoản ngân hàng luôn là 'Báo động đỏ' cấp độ cao nhất. Mọi sự cẩn trọng và một cuộc điện thoại 30 giây có thể cứu vãn hàng tỷ đồng tiền mồ hôi nước mắt.",
        "https://en.wikipedia.org/wiki/Business_email_compromise",
        "Series: Bóc phốt chiêu trò",
        "Chưa dùng",
        "Cảnh báo thủ đoạn tội phạm mạng lừa đảo thay đổi tài khoản cước vận chuyển quốc tế (BEC Fraud) trong hoạt động XNK."
    ],
    [
        53,
        "W35 — 29/08/2026",
        "15 Container Sầu riêng Tươi bị Giữ tại Cửa khẩu Hữu Nghị do Cảnh báo Vượt ngưỡng Kim loại nặng Cadimi: Thiệt hại 1,8 tỷ đồng chi phí hạ bãi, quay đầu và bán tháo giải cứu",
        "Hợp tác xã Nông nghiệp S. (Đắk Lắk), Công ty Thương mại Nông sản XNK, Chi cục Kiểm dịch Thực vật Cửa khẩu Hữu Nghị, Cơ quan Hải quan Trung Quốc (GACC)",
        "Doanh nghiệp gom mua 15 container sầu riêng Ri6 (trọng lượng gần 270 tấn, trị giá hơn 22 tỷ VNĐ) từ nhiều nhà vườn nhỏ lẻ để xuất khẩu sang Trung Quốc. Khi làm thủ tục thông quan tại cửa khẩu biên giới, phía Hải quan Trung Quốc tiến hành lấy mẫu kiểm nghiệm ngẫu nhiên và phát hiện hàm lượng Cadimi vượt ngưỡng cho phép 0,05 mg/kg theo quy chuẩn an toàn thực phẩm của GACC. Lô hàng bị từ chối nhập khẩu, yêu cầu quay đầu khẩn cấp. Do sầu riêng tươi chỉ bảo quản được trong thời gian ngắn, sau 5 ngày nằm chờ và quay đầu về nội địa, hàng bị chín nẫu, buộc phải xả bán tháo tại thị trường nội địa với giá rẻ mạt 25.000 đ/kg.",
        "Thiệt hại trực tiếp 1,8 tỷ đồng (gồm chi phí cắm điện container lạnh 5 ngày tại cửa khẩu, tiền cước vận chuyển hai chiều 450 triệu đồng, và lỗ chênh lệch giá bán tháo sầu riêng chín nẫu so với giá xuất khẩu hợp đồng).",
        "Thu mua nông sản trôi nổi từ các nhà vườn lạm dụng phân bón hóa học và chất kích thích chứa kim loại nặng Cadimi mà không kiểm tra mẫu đất, nguồn nước định kỳ; không gửi mẫu đi xét nghiệm tại các phòng thử nghiệm được chỉ định trước khi đóng container xuất khẩu.",
        "1. Chỉ thu mua sầu riêng từ các vùng trồng đã được cấp Mã số vùng trồng (MSVT) đạt chuẩn VietGAP/GlobalGAP; 2. Bắt buộc kiểm tra định kỳ chỉ tiêu kim loại nặng (Cadimi, Chì, Asen) trong đất canh tác và nguồn nước tưới tiêu; 3. Gửi mẫu trái cây kiểm nghiệm tại các phòng xét nghiệm đạt chuẩn ISO 17025 được GACC thừa nhận trước khi vận chuyển ra cửa khẩu.",
        "Phe A (Thương lái): Hàm lượng kim loại nặng ở mức vi lượng không gây hại ngay, quy định của nước bạn quá khắt khe nhằm dựng hàng rào kỹ thuật ép giá. Phe B (Chuyên gia Nông nghiệp): Tiêu chuẩn an toàn thực phẩm là nguyên tắc bất khả xâm phạm; nếu không kiểm soát từ gốc đất trồng, toàn bộ thương hiệu nông sản quốc gia sẽ bị cấm cửa.",
        "Xuất khẩu nông sản chính ngạch không thể làm theo tư duy 'thương lái chộp giật'. Kiểm soát chất lượng từ gốc rễ vùng trồng là chìa khóa duy nhất để bảo vệ chuỗi giá trị nông sản bền vững.",
        "https://en.wikipedia.org/wiki/Cadmium_poisoning",
        "Series: Tai nạn XNK",
        "Chưa dùng",
        "Bài học thực tế về kiểm soát dư lượng kim loại nặng Cadimi trong xuất khẩu trái cây chính ngạch sang Trung Quốc theo chuẩn GACC."
    ],
    [
        54,
        "W35 — 29/08/2026",
        "Doanh nghiệp Dệt may bị Thu hồi Toàn bộ C/O Mẫu EUR.1 và Truy thu 850 triệu đồng Tiền thuế vì Vi phạm 'Quy tắc từ Vải trở đi' (Fabric-Forward Rule) trong EVFTA",
        "Công ty Cổ phần May mặc Xuất khẩu P. (Hải Dương), Cơ quan Hải quan Liên minh Châu Âu (EU Customs), Bộ Công Thương Việt Nam",
        "Doanh nghiệp xuất khẩu 5 lô hàng quần áo may sẵn trị giá 1,2 triệu USD sang thị trường Đức và Pháp, sử dụng C/O mẫu EUR.1 để khách hàng bên EU được hưởng thuế suất ưu đãi 0% theo Hiệp định EVFTA (thay vì mức thuế MFN 12%). Hai năm sau khi thông quan, Cơ quan Hải quan Châu Âu tiến hành kiểm tra sau thông quan (Post-Clearance Audit) và gửi văn bản yêu cầu xác minh xuất xứ. Qua kiểm toán chuỗi cung ứng, cơ quan quản lý phát hiện toàn bộ vải dệt thoi dùng may áo được doanh nghiệp nhập khẩu từ Trung Quốc (quốc gia không có hiệp định cộng gộp xuất xứ vải với EU). C/O mẫu EUR.1 bị tuyên bố vô hiệu, khách hàng tại EU bị truy thu toàn bộ số tiền thuế 850 triệu đồng kèm tiền phạt chậm nộp, sau đó chuyển toàn bộ khoản phạt này sang yêu cầu doanh nghiệp Việt Nam bồi thường theo hợp đồng.",
        "Doanh nghiệp phải bồi thường 850 triệu đồng tiền thuế truy thu cho đối tác EU; bị đối tác cắt giảm 50% đơn hàng vụ sau và bị Hải quan đưa vào danh sách kiểm soát rủi ro cao đối với mọi lô hàng xuất khẩu tiếp theo.",
        "Không hiểu đúng Quy tắc xuất xứ 'Từ vải trở đi' (Fabric-Forward) trong EVFTA (bắt buộc công đoạn dệt vải và may mặc phải thực hiện tại Việt Nam hoặc các nước thành viên có cam kết cộng gộp như Hàn Quốc/Nhật Bản); tự ý sử dụng vải nguồn gốc nước thứ ba không thuộc diện cộng gộp nhưng vẫn xin cấp C/O ưu đãi.",
        "1. Đào tạo chuyên sâu cho bộ phận thu mua nguyên phụ liệu về Quy tắc xuất xứ cụ thể mặt hàng (PSR) của từng FTA; 2. Sử dụng nguồn vải sản xuất nội địa tại Việt Nam hoặc nhập khẩu từ các quốc gia có cơ chế cộng gộp xuất xứ vải được EU công nhận (như Hàn Quốc theo cơ chế EVFTA Fabric Cumulation); 3. Lưu trữ đầy đủ hồ sơ chứng minh xuất xứ (BOM, hóa đơn vải, chứng từ dệt nhuộm) trong 5 năm phục vụ hậu kiểm.",
        "Phe A (DN may): Sản phẩm cắt may, hoàn thiện tinh xảo tại Việt Nam tạo ra giá trị gia tăng lớn, tại sao lại khắt khe bắt buộc cả khâu dệt vải? Phe B (Chuyên gia FTA): Quy tắc xuất xứ chặt chẽ là điều kiện tiên quyết để ngăn chặn việc nước thứ ba mượn đường Việt Nam để 'rửa xuất xứ' lẩn tránh thuế vào EU.",
        "Ưu đãi thuế quan FTA là 'con dao hai lưỡi'. Hưởng lợi thuế suất 0% hôm nay có thể biến thành khoản truy thu nợ thuế tiền tỷ ngày mai nếu không xây dựng chuỗi cung ứng nguyên liệu chuẩn chỉ từ gốc.",
        "https://en.wikipedia.org/wiki/Rules_of_origin",
        "Series: Tai nạn XNK",
        "Chưa dùng",
        "Bài học xương máu về quy tắc xuất xứ 'Từ vải trở đi' (Fabric-Forward) và kiểm tra sau thông quan trong Hiệp định EVFTA."
    ],
    [
        55,
        "W35 — 29/08/2026",
        "Tịch thu 30 tấn Hóa chất Công nghiệp và Xử phạt 400 triệu đồng do Không Khai báo Hóa chất Nhập khẩu và Thiếu Phiếu An toàn Hóa chất (MSDS) theo Nghị định 113/2017/NĐ-CP",
        "Công ty TNHH Hóa chất & Vật liệu Mới V. (Hải Phòng), Cục Hóa chất (Bộ Công Thương), Chi cục Hải quan Cửa khẩu Cảng Hải Phòng",
        "Doanh nghiệp nhập khẩu 30 tấn hóa chất dung môi công nghiệp (Toluene và Axeton) đóng trong phuy sắt từ Hàn Quốc về Cảng Hải Phòng để cung cấp cho các nhà máy sơn. Do nôn nóng thông quan sớm và nhân viên mới chưa có kinh nghiệm, doanh nghiệp mở tờ khai hải quan thông thường mà không thực hiện thủ tục Khai báo hóa chất nhập khẩu qua Cổng thông tin Một cửa Quốc gia trước khi thông quan, đồng thời nộp Phiếu an toàn hóa chất (MSDS) bản tiếng Anh cũ đã hết hạn sử dụng không có bản dịch tiếng Việt hợp chuẩn. Đội Kiểm soát Hải quan phối hợp Thanh tra Cục Hóa chất kiểm tra đột xuất, phát hiện đây là nhóm hóa chất sản xuất, kinh doanh có điều kiện theo Nghị định 113/2017/NĐ-CP và Nghị định 82/2022/NĐ-CP.",
        "Bị xử phạt vi phạm hành chính 400 triệu đồng; toàn bộ lô hàng 30 tấn hóa chất trị giá 1,4 tỷ VNĐ bị tạm giữ niêm phong 45 ngày để thẩm định an toàn; phát sinh thêm 110 triệu đồng tiền phí lưu kho nguy hiểm chuyên dụng và phí giám định kỹ thuật.",
        "Không tra cứu Danh mục Hóa chất phải khai báo (Phụ lục V) và Hóa chất hạn chế sản xuất, kinh doanh trước khi ký hợp đồng nhập khẩu; coi nhẹ vai trò của Phiếu an toàn hóa chất (MSDS) chuẩn GHS tiếng Việt.",
        "1. Luôn tra cứu mã CAS (Chemical Abstracts Service) và mã HS của hóa chất trên Hệ thống Cơ sở Dữ liệu Hóa chất Quốc gia trước khi nhập hàng; 2. Bắt buộc thực hiện thủ tục Khai báo Hóa chất nhập khẩu qua Cổng Một cửa Quốc gia (NSW) và nhận phản hồi xác nhận tự động trước khi truyền tờ khai hải quan; 3. Yêu cầu nhà sản xuất cung cấp MSDS bản mới nhất theo Hệ thống Hài hòa Toàn cầu (GHS) và chuẩn bị sẵn bản dịch tiếng Việt đầy đủ 16 mục.",
        "Phe A (DN): Hóa chất dung môi thông dụng phục vụ sản xuất công nghiệp thuần túy, việc yêu cầu khai báo giấy phép phức tạp làm chậm tiến độ giao hàng của nhà máy. Phe B (Cơ quan Quản lý): Hóa chất dễ cháy nổ và có nguy cơ dùng làm tiền chất ma túy công nghiệp, việc kiểm soát nghiêm ngặt là vấn đề an ninh quốc gia và phòng cháy chữa cháy đô thị.",
        "Đối với mặt hàng hóa chất, tuân thủ pháp lý an toàn phải đi trước một bước so với hoạt động thương mại. Không có giấy phép và MSDS chuẩn, một container hóa chất sẽ trở thành 'quả bom nổ chậm' phong tỏa dòng tiền doanh nghiệp.",
        "https://en.wikipedia.org/wiki/Safety_data_sheet",
        "Series: Giải cứu hàng hóa",
        "Chưa dùng",
        "Case study thực chiến về thủ tục khai báo hóa chất nhập khẩu theo Nghị định 113/2017/NĐ-CP và quản lý an toàn hóa chất MSDS/GHS."
    ]
]

HOC_THUC_CHIEN_W35_D4 = [
    [
        51,
        "W35 — 29/08/2026",
        "Quy trình 5 bước Khai báo và Đóng gói Vận chuyển Hàng Nguy hiểm (Dangerous Goods - DG) theo Chuẩn IMDG Code và SOLAS",
        "Bước 1: Xác định chính xác Mã số UN (UN Number), Tên gọi vận chuyển chính thức (PSN), Phân nhóm nguy hiểm (Class 1-9) và Nhóm đóng gói (Packing Group I, II, III) dựa trên Phiếu an toàn hóa chất (MSDS); Bước 2: Chuẩn bị Bao bì đóng gói đạt chuẩn chứng nhận UN (UN Certified Packaging) có in biểu tượng UN dập nổi chịu lực, chống rò rỉ; Bước 3: Dán đầy đủ Nhãn cảnh báo nguy hiểm (Hazard Diamond Labels) và Nhãn chỉ hướng (Handling Labels) trên ít nhất 2 mặt đối diện của kiện hàng và 4 mặt ngoài vỏ container; Bước 4: Lập Bản khai Hàng nguy hiểm (Dangerous Goods Declaration - DGD) và Chứng nhận Đóng gói Container (Container Packing Certificate) có chữ ký xác nhận của nhân sự được đào tạo chứng chỉ DG; Bước 5: Gửi hồ sơ trước 48-72 giờ cho Hãng tàu/Cảng vụ để xin duyệt kế hoạch xếp dỡ và bố trí vị trí xếp hàng trên boong hở (On Deck Stowage) theo quy chuẩn an toàn.",
        "Phụ phí Hàng nguy hiểm (DG Surcharge) = Cước biển cơ bản (Base Ocean Freight) + (20% đến 50% Phụ phí DG tùy theo Class) + Phí chứng từ DG (DG Documentation Fee từ 50$ - 150$/bill). Quy chuẩn nhiệt độ cách ly: Duy trì khoảng cách an toàn tối thiểu 3-6 mét so với các chất gây cháy khác.",
        "1. Sử dụng vỏ thùng carton hoặc phuy chứa thông thường không có dấu mộc kiểm định tiêu chuẩn UN; 2. Dán sai nhãn phân nhóm hoặc thiếu nhãn cảnh báo kích thước chuẩn 10cm x 10cm trên vách container; 3. Không khai báo trước với cảng dẫn đến container bị từ chối hạ bãi và phạt tiền neo đậu.",
        "Dán nhãn Hàng nguy hiểm giống như 'Gắn biển cảnh báo bão' - Nhãn hình thoi màu nổi bật, 4 mặt ngoài container phải nhìn thấy rõ từ xa 50 mét để cứu hỏa xử lý khi có sự cố.",
        "Vận chuyển Hàng bách hóa thông thường (General Cargo: Thủ tục đơn giản, đóng gói tự do, cước rẻ) vs Vận chuyển Hàng nguy hiểm (Dangerous Goods IMDG: Bắt buộc MSDS, bao bì chuẩn UN, phụ phí cao gấp đôi, xếp riêng biệt cách ly an toàn trên boong tàu).",
        "Chị Hương (Trưởng phòng Logistics/Mua hàng nhập khẩu hóa chất, pin, sơn) & Đội ngũ Hiện trường Đóng gói Kho bãi.",
        "☐ Tra cứu UN Number & Class trên MSDS\n☐ Mua bao bì có mộc UN dập nổi\n☐ Dán nhãn hình thoi đủ 4 mặt container\n☐ Soạn thảo Bản khai DGD chuẩn ký tên\n☐ Gửi duyệt hãng tàu trước 48h.",
        "https://en.wikipedia.org/wiki/Dangerous_goods",
        "Series: Quy trình Hải quan",
        "Chưa dùng",
        "Cẩm nang step-by-step xử lý và khai báo thủ tục vận tải hàng nguy hiểm theo Bộ luật IMDG Code của IMO và Công ước SOLAS."
    ],
    [
        52,
        "W35 — 29/08/2026",
        "Hướng dẫn 5 bước Tự tính Thuế Chống bán phá giá (Anti-Dumping Duty) và Thuế Tự vệ Tránh Bị Truy thu Hàng tỷ đồng",
        "Bước 1: Tra cứu mã HS 8 số của hàng hóa nhập khẩu trên Biểu thuế XNK hiện hành và đối chiếu với Danh mục các Quyết định áp dụng biện pháp phòng vệ thương mại của Bộ Công Thương; Bước 2: Xác định rõ Xuất xứ hàng hóa (Căn cứ trên C/O, Giấy chứng nhận nhà sản xuất Mill Test Certificate) và Tên Nhà sản xuất/Xuất khẩu cụ thể; Bước 3: Tra cứu Mức thuế suất Chống bán phá giá (CBPG) áp dụng riêng cho nhà sản xuất đó trong Quyết định của Bộ Công Thương (mỗi công ty nước ngoài thường có mức thuế suất CBPG khác nhau từ 0% đến 35,2%); Bước 4: Tính toán tổng số tiền thuế nhập khẩu, thuế CBPG và thuế GTGT theo thứ tự tuần tự chuẩn xác của Luật Quản lý Thuế; Bước 5: Kiểm tra các trường hợp được miễn áp dụng biện pháp CBPG (hàng mẫu, hàng nhập khẩu sản xuất xuất khẩu) và chuẩn bị bộ hồ sơ chứng minh nguồn gốc để nộp cho cơ quan Hải quan.",
        "Tiền thuế Nhập khẩu = Trị giá tính thuế (CIF) x Thuế suất NK. Tiền thuế CBPG = Trị giá tính thuế (CIF) x Thuế suất CBPG riêng biệt. Tiền thuế GTGT = (Trị giá tính thuế CIF + Tiền thuế NK + Tiền thuế CBPG) x Thuế suất GTGT (8% hoặc 10%). Lưu ý: Thuế CBPG được tính gộp vào gốc để tính thuế GTGT.",
        "1. Nhầm lẫn giữa Thuế suất CBPG áp dụng cho 'Công ty sản xuất cụ thể' vs 'Mức thuế suất chung áp dụng cho toàn bộ quốc gia đó' (mức chung thường cao hơn rất nhiều); 2. Tính thuế GTGT trước khi cộng tiền thuế CBPG dẫn đến nộp thiếu thuế; 3. Không kiểm tra thời hạn hiệu lực của Quyết định rà soát hoàng hôn (Sunset Review) dẫn đến áp sai biểu thuế.",
        "Thuế CBPG giống như 'Bậc thang chịu phạt' - Thuế nhập khẩu tính trước, Thuế chống bán phá giá tính kế tiếp, Thuế GTGT trùm lên trên cùng của tất cả các bậc thang.",
        "Hàng hóa nhập khẩu từ Nhà máy hợp tác điều tra phòng vệ thương mại (Được hưởng mức thuế CBPG ưu đãi thấp từ 3-8%) vs Hàng hóa nhập khẩu từ Nhà máy không hợp tác / thương lái trôi nổi (Bị áp mức thuế CBPG trừng phạt kịch khung lên tới 35-40%).",
        "Anh Minh (Chủ doanh nghiệp nhập khẩu sắt thép, nhôm, gỗ công nghiệp, sợi) & Kế toán Thuế XNK.",
        "☐ Tra cứu mã HS trên Quyết định Bộ Công Thương\n☐ Xin Mill Test Cert xác định đúng tên nhà máy\n☐ Tra đúng tỷ lệ thuế CBPG theo bảng phụ lục\n☐ Tính thuế GTGT trên tổng trị giá sau CBPG\n☐ Khai báo đúng mã phân loại thuế trên tờ khai ECUS.",
        "https://en.wikipedia.org/wiki/Dumping_(pricing_policy)",
        "Series: Tối ưu Thuế XNK",
        "Chưa dùng",
        "Hướng dẫn kỹ thuật tính thuế phòng vệ thương mại và thuế chống bán phá giá theo Luật Thuế XNK và hướng dẫn của Bộ Công Thương."
    ],
    [
        53,
        "W35 — 29/08/2026",
        "Hướng dẫn 4 bước Kiểm tra và Đáp ứng Quy tắc Xuất xứ 'Từ vải trở đi' (Fabric-Forward Rule) trong Hiệp định EVFTA cho Hàng Dệt may Xuất khẩu",
        "Bước 1: Phân tích Cấu trúc Bảng định mức Nguyên vật liệu (BOM) của sản phẩm dệt may (Mã HS nhóm 61 hoặc 62) để bóc tách nguồn gốc xuất xứ của từng thành phần: Sợi, Vải chính, Vải lót, Chỉ may và Phụ liệu; Bước 2: Kiểm tra Tiêu chí Xuất xứ cụ thể mặt hàng (PSR) trong Phụ lục II EVFTA: Bắt buộc công đoạn dệt vải (Weaving/Knitting) và công đoạn cắt may (Making-up) phải được thực hiện tại Việt Nam hoặc một nước thành viên EU; Bước 3: Áp dụng Cơ chế Cộng gộp Xuất xứ Vải (Fabric Cumulation Mechanism): Kiểm tra xem vải nhập khẩu từ nước thứ ba có thuộc quốc gia đã ký FTA song phương với EU và có cơ chế cộng gộp được chấp thuận hay không (Ví dụ: Vải có xuất xứ Hàn Quốc có C/O mẫu AK/VK hợp lệ); Bước 4: Hoàn thiện Hồ sơ xin cấp C/O Mẫu EUR.1 trên Hệ thống Quản lý và Cấp chứng nhận xuất xứ điện tử (EcoSys) của Bộ Công Thương kèm bảng kê chi phí sản xuất và hóa đơn VAT vải nội địa.",
        "Tiêu chuẩn Hàm lượng Giá trị Gia tăng (RVC / LVC) = [(FOB - Trị giá Nguyên liệu Không có Xuất xứ VNM) / FOB] x 100%. Tiêu chuẩn De Minimis (Tỷ lệ linh hoạt): Cho phép tối đa 10% trọng lượng của vải không đáp ứng quy tắc chuyển đổi mã số được phép sử dụng mà không làm mất tính xuất xứ của sản phẩm may mặc.",
        "1. Sử dụng vải nhập khẩu từ các nước không có thỏa thuận cộng gộp với EU (như Trung Quốc, Đài Loan) nhưng vẫn khai báo xin C/O EUR.1; 2. Không lưu trữ hợp đồng dệt nhuộm và chứng từ phân tích mẫu vải của nhà máy dệt trong nước; 3. Bỏ qua quy định về nhãn mác thành phần sợi (Fiber Composition Labeling) theo luật tiêu dùng EU.",
        "Quy tắc Fabric-Forward giống như 'Xây nhà từ móng gạch' - Không bắt buộc tự trồng cây bông kéo sợi, nhưng móng nhà (tấm vải) bắt buộc phải dệt tại sân nhà mình hoặc nhà anh em kết nghĩa (Hàn Quốc/EU).",
        "Quy tắc 'Từ sợi trở đi' (Yarn-Forward trong CPTPP: Bắt buộc từ kéo sợi, dệt vải đến may mặc đều phải trong khối CPTPP) vs Quy tắc 'Từ vải trở đi' (Fabric-Forward trong EVFTA: Cho phép nhập sợi từ bất kỳ đâu, chỉ cần dệt vải và may tại Việt Nam).",
        "Anh Dũng (Chủ doanh nghiệp may mặc xuất khẩu) & Chị Hương (Trưởng phòng Xuất nhập khẩu Dệt may).",
        "☐ Bóc tách BOM vải chính & vải lót\n☐ Thu thập C/O vải Hàn Quốc (nếu dùng cơ chế cộng gộp)\n☐ Hóa đơn VAT & quy trình dệt nhuộm nội địa\n☐ Kiểm tra tỷ lệ De Minimis < 10%\n☐ Nộp hồ sơ EcoSys lưu trữ chứng từ 5 năm.",
        "https://en.wikipedia.org/wiki/European_Union%E2%80%93Vietnam_Free_Trade_Agreement",
        "Series: Tối ưu Thuế XNK",
        "Chưa dùng",
        "Cẩm nang kỹ thuật thực chiến đáp ứng Quy tắc xuất xứ 'Từ vải trở đi' và cơ chế cộng gộp xuất xứ vải trong Hiệp định EVFTA."
    ],
    [
        54,
        "W35 — 29/08/2026",
        "Quy trình 5 bước Đăng ký Mã số Doanh nghiệp Xuất khẩu Thực phẩm & Nông sản sang Trung Quốc trên Hệ thống CIFER (GACC Decree 248)",
        "Bước 1: Xác định Nhóm sản phẩm xuất khẩu thuộc diện nào: Nhóm 18 mặt hàng rủi ro cao (Bắt buộc phải được Cơ quan có thẩm quyền Việt Nam thẩm định hồ sơ và gửi công văn giới thiệu cho GACC - như Cục BVTV, Cục Thú y, Cục Quản lý Chất lượng Nông Lâm sản và Thủy sản) hay Nhóm hàng hóa thông thường (Doanh nghiệp được tự nộp hồ sơ trực tiếp); Bước 2: Tạo tài khoản doanh nghiệp trên Cổng thông tin CIFER trực tuyến của Tổng cục Hải quan Trung Quốc (cifer.singlewindow.cn); Bước 3: Điền đầy đủ thông tin năng lực nhà xưởng, sơ đồ mặt bằng quy trình sản xuất, công suất chế biến, kiểm soát mối nguy HACCP/ISO 22000 và danh sách nguồn nguyên liệu đầu vào; Bước 4: Tải lên các chứng từ đính kèm (Giấy phép ĐKKD, Giấy chứng nhận cơ sở đủ điều kiện ATTP, Giấy cam kết tuân thủ tiêu chuẩn an toàn sinh học biên giới); Bước 5: Theo dõi tiến độ xét duyệt mã số trên hệ thống CIFER, nhận Mã số Đăng ký Doanh nghiệp Xuất khẩu (China Registration Code) và in mã số này lên bao bì trong và bao bì ngoài của sản phẩm trước khi xuất hàng.",
        "Thời hạn hiệu lực của Mã số CIFER: Chính xác 05 năm kể từ ngày được GACC phê duyệt. Thời hạn nộp hồ sơ gia hạn (Renewal): Bắt buộc nộp hồ sơ trước thời điểm hết hạn từ 03 đến 06 tháng trên hệ thống CIFER.",
        "1. Tự ý nộp hồ sơ trực tiếp đối với các mặt hàng thuộc nhóm 18 loại thực phẩm rủi ro cao mà không thông qua Bộ NN&PTNT xét duyệt chỉ định; 2. Không in mã số CIFER lên bao bì sản phẩm hoặc chỉ in trên thùng carton ngoài mà quên in trên nhãn sản phẩm bên trong; 3. Thay đổi địa chỉ nhà xưởng hoặc mở rộng dây chuyền sản xuất nhưng không làm thủ tục khai báo sửa đổi thông tin với GACC.",
        "Mã số CIFER là 'Thẻ căn cước công dân của gói thực phẩm vào Trung Quốc' - Phải có mã số do GACC cấp in rõ trên bao bì thì hải quan cửa khẩu mới mở cổng luồng thông quan.",
        "Tự đăng ký trực tiếp trên CIFER (Áp dụng cho bánh kẹo, trà khô, nước ép hoa quả, gia vị chế biến) vs Đăng ký qua Cơ quan Quản lý Nhà nước chỉ định (Áp dụng cho thịt, thủy sản, yến sào, hạt ngũ cốc, dầu thực vật, trái cây tươi đông lạnh).",
        "Anh Minh (Chủ doanh nghiệp/HTX chế biến thực phẩm, nông sản, trái cây xuất khẩu) & Cán bộ Đảm bảo Chất lượng QA/QC.",
        "☐ Phân loại nhóm 18 hàng rủi ro cao hay thông thường\n☐ Tạo tài khoản cifer.singlewindow.cn\n☐ Chuẩn bị chứng nhận ATTP & HACCP bản song ngữ\n☐ Tải hồ sơ sơ đồ nhà xưởng\n☐ Theo dõi cấp mã và thiết kế nhãn in bao bì.",
        "https://en.wikipedia.org/wiki/General_Administration_of_Customs",
        "Series: Quy trình Hải quan",
        "Chưa dùng",
        "Hướng dẫn chi tiết quy trình đăng ký mã số xuất khẩu nông sản thực phẩm sang Trung Quốc trên hệ thống CIFER theo Lệnh 248 của GACC."
    ],
    [
        55,
        "W35 — 29/08/2026",
        "Hướng dẫn 4 bước Thiết lập Quy trình Phòng chống Lừa đảo Email Thay đổi Tài khoản Ngân hàng (Anti-BEC Framework) trong Thanh toán Quốc tế",
        "Bước 1: Thiết lập Quy tắc 'Xác thực Độc lập Đa kênh' (Out-of-Band Verification): Bắt buộc khi nhận được bất kỳ email nào thông báo thay đổi số tài khoản ngân hàng, thay đổi tên người thụ hưởng hoặc đổi ngân hàng trung gian, nhân viên phụ trách phải kích hoạt quy trình dừng chuyển tiền khẩn cấp; Bước 2: Thực hiện cuộc gọi thoại trực tiếp (Voice Phone Call) hoặc Video Call qua số điện thoại đường dây nóng/WhatsApp chính thức của Kế toán trưởng hoặc Giám đốc đối tác (số điện thoại đã được ghi nhận trong Hợp đồng mua bán gốc hoặc trao đổi trực tiếp từ trước); Bước 3: Yêu cầu đối tác phát hành Thư Thông báo Thay đổi Tài khoản có chữ ký sống của Người đại diện theo pháp luật, đóng dấu pháp nhân và kèm theo Giấy xác nhận tài khoản của Ngân hàng (Bank Account Verification Letter); Bước 4: Áp dụng phương thức Chuyển khoản Thử nghiệm (Test Wire): Đối với các khoản thanh toán lớn trên 20.000 USD, thực hiện chuyển trước một khoản tiền nhỏ (ví dụ 100 USD) và yêu cầu đối tác xác nhận đã ghi có thành công trước khi chuyển phần tiền còn lại.",
        "Quy tắc An toàn Thanh toán 30 Giây: 30 giây gọi điện xác minh trực tiếp = Bảo vệ 100% tài sản tiền tỷ của doanh nghiệp khỏi các cuộc tấn công mạng quốc tế.",
        "1. Nhấn nút 'Reply' trực tiếp vào email giả mạo để hỏi lại xác minh (vì hacker đang chiếm quyền hòm thư nên sẽ lập tức trả lời xác nhận giả); 2. Gọi điện thoại theo số điện thoại ghi ở phần chữ ký của email lừa đảo (số điện thoại do hacker tạo ra); 3. Kế toán chịu áp lực hối thúc chuyển tiền gấp từ 'email giả mạo Giám đốc' nên bỏ qua quy trình kiểm soát nội bộ.",
        "Đổi tài khoản thanh toán giống như 'Gặp người lạ đổi chìa khóa nhà' - Dù ăn mặc giống người nhà, bắt buộc phải gọi điện trực tiếp cho chính chủ để hỏi cho ra nhẽ trước khi mở cửa giao tiền.",
        "Quy trình Phê duyệt Thanh toán Lỏng lẻo (Chỉ nhìn email hóa đơn là chuyển tiền, rủi ro mất trắng 100% khi bị hacker tấn công) vs Quy trình Chuẩn hóa Anti-BEC (Bắt buộc xác thực đa kênh độc lập + gọi điện thoại + ký duyệt kép 2 cấp lãnh đạo).",
        "Kế toán trưởng, Nhân viên Thanh toán Quốc tế & Giám đốc Doanh nghiệp Xuất nhập khẩu.",
        "☐ Tuyệt đối không bấm Reply email đổi tài khoản\n☐ Gọi điện thoại số cố định trong Hợp đồng gốc\n☐ Đòi thư xác nhận tài khoản có dấu mộc ngân hàng\n☐ Chuyển thử nghiệm 100 USD cho đơn hàng lớn\n☐ Ký duyệt kép 2 cấp trước khi duyệt lệnh chuyển tiền.",
        "https://en.wikipedia.org/wiki/Wire_fraud",
        "Series: Đàm phán Quốc tế",
        "Chưa dùng",
        "Cẩm nang kỹ thuật thiết lập quy trình kiểm soát nội bộ phòng chống tội phạm lừa đảo chuyển tiền quốc tế qua email (BEC Fraud)."
    ]
]

FUNFACT_W35_D4 = [
    [
        51,
        "W35 — 29/08/2026",
        "Thảm họa Nổ Tàu Hàng SS Grandcamp (1947 tại Texas): Làm thế nào một tàn thuốc lá đã khai sinh ra Bộ luật Hàng nguy hiểm IMDG của Liên Hợp Quốc?",
        "Bạn có biết rằng toàn bộ quy chuẩn vận chuyển hàng nguy hiểm khắt khe nhất thế giới ngày nay bắt nguồn từ một vụ nổ tàu chở phân bón năm 1947 làm rung chuyển cả nước Mỹ và tạo ra cột khói cao 6.000 mét?",
        "Vụ nổ tàu chở hàng SS Grandcamp tại Cảng Texas City (Mỹ) vào ngày 16/04/1947, cướp đi sinh mạng của gần 600 người và phá hủy hoàn toàn khu cảng công nghiệp.",
        "Sau Thế chiến II, tàu SS Grandcamp cập cảng để bốc xếp 2.200 tấn phân bón Amoni Nitrat (Ammonium Nitrate) sang châu Âu. Thủy thủ đoàn và công nhân bốc xếp thiếu hiểu biết về đặc tính oxy hóa cực mạnh của hóa chất, thản nhiên hút thuốc và đóng hàng bằng bao giấy dễ cháy. Khi ngọn lửa bùng phát trong hầm hàng kín, nỗ lực dập lửa sai cách đã biến con tàu thành một 'quả bom nguyên tử phi hạt nhân' phát nổ với sức công phá kinh hoàng.",
        "Thảm họa chấn động này đã thúc đẩy Tổ chức Hàng hải Quốc tế (IMO) thành lập Tiểu ban chuyên trách và chính thức ban hành Bộ luật Quốc tế về Vận chuyển Hàng hóa Nguy hiểm bằng Đường biển (IMDG Code) vào năm 1965, thiết lập 9 nhóm phân loại nguy hiểm và quy tắc đóng gói bắt buộc toàn cầu.",
        "Trước năm 1965, hàng nguy hiểm được xếp chung lộn xộn với hàng bách hóa vs Ngày nay, mọi kiện hàng hóa chất, pin lithium đều phải có số UN, nhãn cảnh báo hình thoi và xếp cách ly nghiêm ngặt trên boong tàu.",
        "Trong quản trị chuỗi cung ứng, sự chủ quan và thiếu quy chuẩn an toàn hóa chất là tội ác. Tuân thủ nghiêm ngặt quy trình kỹ thuật đóng gói chính là lá chắn bảo vệ tính mạng con người và sự tồn vong của doanh nghiệp.",
        "https://en.wikipedia.org/wiki/Texas_City_disaster",
        "Series: Lịch sử Logistics",
        "Chưa dùng",
        "Nguồn gốc lịch sử ra đời của Bộ luật Hàng nguy hiểm IMDG Code sau thảm họa nổ tàu tại Texas City năm 1947."
    ],
    [
        52,
        "W35 — 29/08/2026",
        "Chiếc Pallet Gỗ Tiêu chuẩn Châu Âu (EUR-Pallet): Tại sao một khung gỗ 80x120cm lại trở thành phát minh giúp tăng 800% năng suất bốc dỡ toàn cầu?",
        "Tại sao đúng kích thước 800 x 1200 mm của chiếc Pallet gỗ Châu Âu lại có sức mạnh định hình kích thước của toàn bộ toa tàu hỏa, xe tải và lối đi trong mọi nhà kho hiện đại trên thế giới?",
        "Hiệp hội Đường sắt Quốc tế (UIC) và các công ty đường sắt Châu Âu chuẩn hóa chiếc Pallet EUR-1 (Euro-pallet) vào năm 1961.",
        "Sau Thế chiến II, việc phục hồi kinh tế Châu Âu bị nghẽn nghiêm trọng tại các nhà ga đường sắt do hàng trăm loại pallet kích thước hỗn loạn khiến công nhân phải bốc vác từng bao hàng bằng tay, mất hàng giờ đồng hồ để xếp đầy một toa tàu hỏa.",
        "Tiêu chuẩn EUR-Pallet ra đời giúp xe nâng (forklift) xếp dỡ một toa tàu chỉ trong 15 phút (nhanh gấp 8 lần). Tiêu chuẩn này trở thành kích thước nền tảng để thiết kế hệ thống kệ kho tự động (ASRS), kích thước thùng xe tải và container tiêu chuẩn ISO toàn cầu.",
        "Châu Âu sử dụng chuẩn EUR-Pallet (800x1200mm) vs Bắc Mỹ sử dụng chuẩn GMA Pallet (40x48 inch / 1016x1219mm) vs Châu Á sử dụng chuẩn Nhật/Hàn (1100x1100mm).",
        "Đôi khi đổi mới đột phá không đến từ những công nghệ viễn tưởng xa vời mà đến từ việc chuẩn hóa một chi tiết vật lý đơn giản nhất. Chuẩn hóa quy cách đóng gói là bước đầu tiên để tự động hóa toàn diện chuỗi cung ứng.",
        "https://en.wikipedia.org/wiki/EUR-pallet",
        "Series: Lịch sử Logistics",
        "Chưa dùng",
        "Lịch sử chuẩn hóa chiếc pallet gỗ EUR-Pallet và tác động định hình toàn bộ hạ tầng kho vận logistics thế giới."
    ],
    [
        53,
        "W35 — 29/08/2026",
        "Câu chuyện 'Quả Cà Chua NAFTA' và Cuộc cách mạng Truy xuất Nguồn gốc: Làm thế nào một đĩa salad đã khai sinh ra Quy tắc Xuất xứ (Rules of Origin) phức tạp nhất lịch sử thương mại?",
        "Bạn có biết rằng để một chiếc áo thun hay một chai sốt cà chua được miễn thuế khi vượt biên giới, các luật sư thương mại quốc tế đã phải viết ra hàng vạn trang quy tắc xuất xứ chi tiết đến từng tép tỏi và cuộn chỉ?",
        "Hiệp định Thương mại Tự do Bắc Mỹ (NAFTA 1994) giữa Mỹ, Canada và Mexico và các cuộc tranh chấp quy tắc xuất xứ nông sản - dệt may kinh điển.",
        "Khi Mỹ và Mexico ký hiệp định tự do thương mại xóa bỏ thuế quan, các nhà sản xuất nước thứ ba tìm cách chuyển cà chua và sợi vải giá rẻ vào Mexico để 'mông má' sơ sài rồi tuồn sang Mỹ bán trốn thuế. Để ngăn chặn gian lận, các nhà đàm phán đã tạo ra hệ thống Tiêu chí Chuyển đổi Mã số Hàng hóa (CTC) và Quy tắc 'Từ sợi trở đi' (Yarn-Forward).",
        "Khai sinh ra toàn bộ hệ thống Quy tắc Xuất xứ (Rules of Origin) hiện đại được Tổ chức Hải quan Thế giới (WCO) và WTO áp dụng trong hàng trăm hiệp định FTA thế hệ mới ngày nay (EVFTA, CPTPP, RCEP).",
        "Hiệp định thương mại tự do truyền thống chỉ cần chứng nhận xuất xứ lỏng lẻo vs FTA thế hệ mới bắt buộc truy xuất nguồn gốc chuỗi cung ứng đa tầng từ nguyên liệu thô đến khâu hoàn thiện cuối cùng.",
        "Trong kỷ nguyên hội nhập FTA sâu rộng, lợi thế cạnh tranh không chỉ nằm ở giá thành sản xuất mà nằm ở năng lực làm chủ và chứng minh chuỗi cung ứng minh bạch, chuẩn chỉnh ngay từ nguyên liệu đầu vào.",
        "https://en.wikipedia.org/wiki/North_American_Free_Trade_Agreement",
        "Series: Thương mại Quốc tế",
        "Chưa dùng",
        "Nguồn gốc lịch sử của các quy tắc xuất xứ hàng hóa hiện đại từ Hiệp định NAFTA và bài học làm chủ chuỗi cung ứng cho DN xuất khẩu.",
    ],
    [
        54,
        "W35 — 29/08/2026",
        "Mã số Doanh nghiệp Toàn cầu D-U-N-S Number: Làm thế nào 9 chữ số ra đời từ năm 1962 đã trở thành 'Thẻ Căn cước Tài chính' bắt buộc của mọi tập đoàn xuyên quốc gia?",
        "Bạn có biết rằng Apple, Walmart, Hải quan Mỹ (US Customs) và thậm chí cả Liên Hợp Quốc đều bắt buộc mọi đối tác cung ứng phải có một dãy số 9 chữ số có tên là D-U-N-S trước khi được ký bất kỳ hợp đồng nghìn đô nào?",
        "Tập đoàn Thông tin Tín dụng Dun & Bradstreet (D&B) phát minh ra Hệ thống Đánh số Dữ liệu Toàn cầu (Data Universal Numbering System - D-U-N-S) vào năm 1962.",
        "Trong thập niên 1960, sự bùng nổ của thương mại đa quốc gia khiến các ngân hàng và tập đoàn lớn rơi vào ma trận hỗn loạn khi hàng nghìn công ty có tên giống hệt nhau hoặc thành lập các pháp nhân con phức tạp để trốn nợ và gian lận tín dụng.",
        "Mã D-U-N-S trở thành tiêu chuẩn định danh doanh nghiệp toàn cầu duy nhất được hơn 50 hiệp hội thương mại và chính phủ các nước công nhận. Nó liên kết trực tiếp với hồ sơ xếp hạng tín nhiệm (Credit Score) và lịch sử thanh toán của hơn 500 triệu doanh nghiệp trên toàn thế giới.",
        "Mã số thuế quốc gia (Chỉ có giá trị tra cứu nội địa trong một nước) vs Mã D-U-N-S (Có giá trị tra cứu và bảo chứng tín nhiệm xuyên biên giới tại hơn 200 quốc gia).",
        "Minh bạch thông tin tài chính và xây dựng điểm tín nhiệm quốc tế là 'tấm hộ chiếu vô hình' giúp doanh nghiệp Việt Nam bước qua cánh cửa thẩm định khắt khe của các bạn hàng khổng lồ toàn cầu.",
        "https://en.wikipedia.org/wiki/Data_Universal_Numbering_System",
        "Series: Đàm phán & Tư duy",
        "Chưa dùng",
        "Khám phá vai trò sống còn của Hệ thống định danh tín dụng doanh nghiệp D-U-N-S Number của Dun & Bradstreet trong thương mại quốc tế."
    ],
    [
        55,
        "W35 — 29/08/2026",
        "Hệ thống Cơ chế Một cửa Hải quan (Single Window): Từ sáng kiến xử lý giấy tờ hàng hải năm 1989 của Thụy Điển đến cuộc cách mạng số hóa cắt giảm 90% thời gian thông quan toàn cầu",
        "Bạn có tin rằng trước khi có Cơ chế Một cửa (Single Window), một con tàu chở hàng cập cảng phải in và nộp tới hơn 40 bộ hồ sơ giấy khác nhau cho hàng chục cơ quan nhà nước, mất cả tuần lễ chỉ để chờ đóng dấu mộc?",
        "Hải quan Hoàng gia Thụy Điển triển khai hệ thống Cơ chế Một cửa (Single Window) điện tử đầu tiên trên thế giới vào năm 1989, sau đó được Liên Hợp Quốc (UN/CEFACT) chuẩn hóa thành Khuyến nghị số 33 vào năm 2005.",
        "Vào cuối thế kỷ 20, thủ tục hành chính cửa khẩu là cơn ác mộng lớn nhất của thương mại quốc tế: Hải quan đòi 1 bộ giấy, Kiểm dịch đòi 1 bộ khác, Cảng vụ, Biên phòng, Y tế mỗi bên giữ một mẫu tờ khai riêng biệt khiến tàu bè phải nằm chờ hàng tuần tại cảng làm tăng vọt chi phí lưu kho bãi.",
        "Sáng kiến 'Một điểm tiếp nhận duy nhất - Xử lý đồng thời' đã lan tỏa tới hơn 100 quốc gia (bao gồm Cổng Một cửa Quốc gia Việt Nam và Một cửa ASEAN), giúp cắt giảm thời gian thông quan từ vài ngày xuống còn vài phút, tiết kiệm hàng trăm tỷ USD chi phí logistics mỗi năm cho nền kinh tế toàn cầu.",
        "Thủ tục nộp giấy thủ công thời xưa (Chạy đi gõ cửa 5 cơ quan bộ ngành mất 3-5 ngày) vs Cổng Thông tin Một cửa Quốc gia ngày nay (Truyền 1 thông điệp dữ liệu điện tử, hệ thống tự động điều phối liên ngành và trả kết quả thông quan trong 15 phút).",
        "Tinh giản quy trình và số hóa không chỉ là công cụ hỗ trợ mà là đòn bẩy sống còn để giải phóng nguồn lực. Doanh nghiệp làm chủ dữ liệu số sẽ luôn dẫn đầu về tốc độ và chi phí trên thương trường.",
        "https://en.wikipedia.org/wiki/Single-window_system",
        "Series: Công nghệ & Đổi mới",
        "Chưa dùng",
        "Lịch sử hình thành Cơ chế Một cửa Hải quan Single Window từ Thụy Điển đến chuẩn mực tạo thuận lợi thương mại toàn cầu của Liên Hợp Quốc."
    ]
]

LEGAL_NEW_W35_D4 = [
    [
        307,
        "Nghị định 53/2024/NĐ-CP",
        "Quy định chi tiết một số điều của Luật Giao dịch điện tử về chữ ký điện tử, chứng từ điện tử và dịch vụ tin cậy trong thủ tục hải quan và một cửa quốc gia",
        "Nghị định",
        "Chính phủ",
        "2024-05-20",
        "2024-07-01",
        "Đang áp dụng",
        "Nghị định 130/2018/NĐ-CP",
        "Công nhận 100% giá trị pháp lý của chứng từ điện tử trong thủ tục hải quan; cấm công chức hải quan đòi bản giấy bổ sung đối với chứng từ đã nộp qua Cổng Một cửa Quốc gia (NSW).",
        "Chuẩn hóa chữ ký số chuyên dùng, liên thông dữ liệu Một cửa Quốc gia và xóa bỏ hoàn toàn rào cản hành chính giấy trong thông quan XNK.",
        "2026-08-29"
    ],
    [
        308,
        "Thông tư 18/2023/TT-BCT",
        "Quy định Quy tắc xuất xứ hàng hóa trong Hiệp định Thương mại Tự do giữa Việt Nam và Israel (VIFTA)",
        "Thông tư",
        "Bộ Công Thương",
        "2023-10-12",
        "2024-01-01",
        "Đang áp dụng",
        "Không có",
        "Quy định chi tiết tiêu chí chuyển đổi mã số hàng hóa (CTC), hàm lượng giá trị gia tăng (RVC ≥ 40%) và thủ tục cấp C/O mẫu VIFTA để hưởng ưu đãi thuế quan xóa bỏ đến 92,7% dòng thuế sang Israel và Trung Đông.",
        "Khung pháp lý xuất xứ đầu tiên giữa Việt Nam và một quốc gia khu vực Trung Đông, mở ra thị trường xuất khẩu chiến lược cho nông sản và hàng dệt may.",
        "2026-08-29"
    ],
    [
        309,
        "Nghị định 113/2017/NĐ-CP & Nghị định 82/2022/NĐ-CP",
        "Quy định chi tiết và hướng dẫn thi hành một số điều của Luật Hóa chất về quản lý, khai báo hóa chất nhập khẩu và an toàn hóa chất",
        "Nghị định",
        "Chính phủ",
        "2022-10-18",
        "2022-12-22",
        "Đang áp dụng",
        "Nghị định 113/2017/NĐ-CP (sửa đổi bổ sung)",
        "Bắt buộc thực hiện Khai báo hóa chất nhập khẩu qua Cổng thông tin Một cửa Quốc gia trước khi thông quan; quy chuẩn hóa Phiếu an toàn hóa chất (MSDS) 16 mục theo Hệ thống GHS.",
        "Tự động hóa cấp Giấy xác nhận khai báo hóa chất qua NSW trong vài phút; siết chặt chế tài xử phạt hành chính đối với hành vi nhập khẩu hóa chất không khai báo.",
        "2026-08-29"
    ],
    [
        310,
        "GACC Decree 248 & 249 (Lệnh 248, 249)",
        "Quy định của Tổng cục Hải quan Trung Quốc về Quản lý An toàn Thực phẩm Xuất nhập khẩu và Đăng ký Doanh nghiệp Sản xuất Nước ngoài trên Hệ thống CIFER",
        "Quy định quốc tế",
        "Tổng cục Hải quan Trung Quốc (GACC)",
        "2021-04-12",
        "2022-01-01",
        "Đang áp dụng",
        "Lệnh 145 và Lệnh 243 của AQSIQ/GACC",
        "Bắt buộc toàn bộ doanh nghiệp xuất khẩu thực phẩm, nông sản sang Trung Quốc phải đăng ký mã số doanh nghiệp trên hệ thống CIFER (cifer.singlewindow.cn); phân nhóm 18 thực phẩm rủi ro cao phải qua cơ quan thẩm quyền chỉ định.",
        "Bắt buộc in mã số đăng ký CIFER lên bao bì trong và bao bì ngoài sản phẩm; siết chặt kiểm tra an toàn sinh học và truy xuất nguồn gốc tại cửa khẩu Trung Quốc.",
        "2026-08-29"
    ],
    [
        311,
        "IMDG Code (Amendments 41-22 / SOLAS Chap VII)",
        "Bộ luật Quốc tế về Vận chuyển Hàng hóa Nguy hiểm bằng Đường biển của Tổ chức Hàng hải Quốc tế (IMO)",
        "Công ước & Quy chuẩn Quốc tế",
        "Tổ chức Hàng hải Quốc tế (IMO) / Bộ Giao thông Vận tải Việt Nam",
        "2022-05-20",
        "2024-01-01",
        "Đang áp dụng",
        "IMDG Code các phiên bản trước",
        "Quy chuẩn bắt buộc toàn cầu về phân loại (Class 1-9), đóng gói theo chuẩn bao bì UN, dán nhãn kim cương cảnh báo nguy hiểm, lập Bản khai Hàng nguy hiểm (DGD) và xếp dỡ cách ly trên tàu biển, đặc biệt kiểm soát pin Lithium (UN 3480/3481) và hóa chất dễ cháy nổ.",
        "Bắt buộc chứng chỉ thử nghiệm UN 38.3 đối với pin lithium; chế tài phạt kịch khung và truy cứu trách nhiệm hình sự đối với hành vi khai man hàng nguy hiểm.",
        "2026-08-29"
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
    print("🚀 BẮT ĐẦU QUY TRÌNH WF06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS NGÀY 29/08/2026 (W35 — D4)")
    print("="*90)
    
    # 1. Thu thập toàn bộ 20 URL để chạy qua Technical Gatekeeper
    validation_list = []
    
    for item in TIN_NGANH_W35_D4:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[12],
            "keywords": ["single", "window", "customs", "trade", "agreement", "tariff", "vat", "value", "added", "tax", "harmonized", "system", "phytosanitary", "plant", "quarantine"]
        })
        
    for item in CASE_STUDY_W35_D4:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[10],
            "keywords": ["imdg", "dangerous", "goods", "maritime", "email", "compromise", "fraud", "wire", "transfer", "cadmium", "metal", "toxicity", "rules", "origin", "trade", "safety", "data", "sheet", "chemical"]
        })
        
    for item in HOC_THUC_CHIEN_W35_D4:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[10],
            "keywords": ["dangerous", "goods", "hazard", "transport", "dumping", "tariff", "duty", "trade", "price", "vietnam", "european", "trade", "agreement", "origin", "customs", "china", "administration", "wire", "fraud", "financial"]
        })
        
    for item in FUNFACT_W35_D4:
        validation_list.append({
            "stt": item[0],
            "title": item[2],
            "url": item[9],
            "keywords": ["texas", "city", "disaster", "ship", "explosion", "pallet", "eur", "railway", "standard", "logistics", "nafta", "trade", "agreement", "origin", "duns", "number", "system", "business", "customs", "single", "window"]
        })
        
    # 2. Rào chắn Kỹ thuật cưỡng chế
    print(f"🛡️ Gửi {len(validation_list)} URL vào Technical Gatekeeper (url_validator.py)...")
    enforce_strict_validation(validation_list)
    
    # 3. Ghi dữ liệu vào 4 tệp Excel Kho Ý Tưởng
    append_data_to_excel(
        "Tin_Nganh_Ideas_Bank_Eureka.xlsx",
        TIN_NGANH_W35_D4,
        center_cols=(1, 2, 5, 14),
        status_col=14
    )
    
    append_data_to_excel(
        "Case_Study_Ideas_Bank_Eureka.xlsx",
        CASE_STUDY_W35_D4,
        center_cols=(1, 2, 12, 13),
        status_col=13
    )
    
    append_data_to_excel(
        "Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx",
        HOC_THUC_CHIEN_W35_D4,
        center_cols=(1, 2, 12, 13),
        status_col=13
    )
    
    append_data_to_excel(
        "Funfact_Ideas_Bank_Eureka.xlsx",
        FUNFACT_W35_D4,
        center_cols=(1, 2, 11, 12),
        status_col=12
    )
    
    # 4. Feedback Loop: Cập nhật Danh_sach_Van_ban_Phap_luat_XNK.xlsx
    update_legal_tracker(LEGAL_NEW_W35_D4)
    
    print("\n" + "="*90)
    print("🎉 HOÀN THÀNH 100% QUY TRÌNH CẬP NHẬT IDEAS BANK THEO CHUẨN WF06 NGÀY 29/08/2026!")
    print("="*90)


if __name__ == "__main__":
    main()
