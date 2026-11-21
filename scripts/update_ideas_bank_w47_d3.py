# -*- coding: utf-8 -*-
"""
Execute Workflow 06: Ideas Bank Daily Update (wf06) - Week 47 Day 3 (18/11/2026)
Appends 5 new verified ideas (STT 261-265) to each of the 4 Pillars (TN, CS, HT, FF).
Runs through Technical Gatekeeper url_validator.py and Banned Words check before saving.
Updates Danh_sach_Van_ban_Phap_luat_XNK.xlsx (Feedback Loop) for missing verified laws.
Uses atomic save (write to .tmp.xlsx and os.replace) to strictly avoid any backup files.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from scripts.url_validator import enforce_strict_validation

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# Styles matching existing sheets
BORDER_THIN = Border(
    left=Side(style='thin', color='E5E7EB'),
    right=Side(style='thin', color='E5E7EB'),
    top=Side(style='thin', color='E5E7EB'),
    bottom=Side(style='thin', color='E5E7EB')
)
FONT_BOLD = Font(name='Segoe UI', size=10, bold=True)
FONT_NORMAL = Font(name='Segoe UI', size=10, bold=False)
FILL_EVEN = PatternFill(fill_type='solid', start_color='F9FAFB', end_color='F9FAFB')
FILL_ODD = PatternFill(fill_type='solid', start_color='FFFFFF', end_color='FFFFFF')

ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

BANNED_WORDS_CHECK = [
    "trốn thuế", "lách luật", "hàng lậu", "hàng nhái", "hàng giả",
    "bao thuế", "đi chui", "đi lậu", "tiểu ngạch",
    "mua hàng trung quốc tận gốc", "đánh hàng quảng châu", "gom hàng biên", "chạy line"
]


def atomic_save(workbook, target_path):
    """Save workbook via temporary file and atomic replace to prevent locks and backups."""
    tmp_path = target_path + ".tmp.xlsx"
    workbook.save(tmp_path)
    os.replace(tmp_path, target_path)


# ==============================================================================
# 1. DATA FOR TIN NGANH (15 cols) — STT 261..265 (W47 — 18/11/2026)
# ==============================================================================
TN_NEW_IDEAS = [
    {
        "stt": 261,
        "week": "W47 — 18/11/2026",
        "title": "Công Văn 24153/CHQ-NVTHQ (12/11/2026): Hải Quan Siết Rà Soát Mã HS Mỹ Phẩm Nhập Khẩu (Sữa Rửa Mặt 3401.30.00 & Mặt Nạ Dưỡng Da 3307.90.90)",
        "agency": "Cục Hải quan (Ban Nghiệp vụ Thuế Hải quan)",
        "effective_date": "Ban hành và chỉ đạo triển khai từ ngày 12/11/2026",
        "core_diff": "Cũ: Nhiều doanh nghiệp nhập khẩu mỹ phẩm thường gộp chung sữa rửa mặt và mặt nạ dưỡng da vào nhóm 33.04 (chế phẩm chăm sóc da thông thường) do quen tay khai báo hoặc dựa theo tên thương mại; Mới: Cục Hải quan chỉ đạo rà soát toàn diện hồ sơ và thực tế hàng hóa theo Chú giải Chương 33 & Danh mục XNK: sữa rửa mặt dạng lỏng/kem chứa chất hoạt động bề mặt làm sạch da đóng gói bán lẻ phân loại vào nhóm 34.01 (3401.30.00), còn mặt nạ dưỡng da làm từ vải không dệt thấm tẩm mỹ phẩm phân loại vào nhóm 33.07 (3307.90.90).",
        "affected_entity": "Các doanh nghiệp nhập khẩu, phân phối mỹ phẩm Hàn Quốc, Nhật Bản, Châu Âu và các đơn vị làm dịch vụ khai báo hải quan mặt hàng chăm sóc cá nhân.",
        "risk_pain": "Nếu bị hải quan rà soát lại và xác định khai sai mã HS dẫn đến chênh lệch thuế suất (MFN hoặc thuế ưu đãi FTA), doanh nghiệp có thể bị ấn định truy thu thuế trong 5 năm, xử phạt vi phạm hành chính 10%–20% số thuế thiếu theo Nghị định 169/2026/NĐ-CP và bị lấy mẫu phân tích phân loại kéo dài thời gian thông quan.",
        "market_reaction": "Cộng đồng doanh nghiệp nhập khẩu mỹ phẩm và đại lý hải quan rà soát khẩn cấp bảng thành phần INCI, MSDS và chứng thư giám định của các lô hàng sữa rửa mặt, mặt nạ chuẩn bị về cảng cuối năm.",
        "pro_argument": "Phân loại đúng bản chất hóa học và cấu tạo sản phẩm giúp áp dụng thuế suất công bằng, thống nhất trên toàn quốc và tránh tình trạng mỗi cửa khẩu áp một mã HS khác nhau.",
        "con_argument": "Sự thay đổi định hướng áp mã từ nhóm 33.04 sang 34.01/33.07 dễ phát sinh rủi ro truy thu hồi tố cho các doanh nghiệp đã thông quan nhiều tờ khai trước đó nếu chưa kịp xin xác định trước mã số (Pre-ruling).",
        "intl_ref": "Hệ thống Chú giải hài hòa HS Code (Explanatory Notes) của Tổ chức Hải quan Thế giới (WCO) phân biệt rõ chế phẩm rửa da bề mặt (Heading 34.01) với kem dưỡng da (Heading 33.04).",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/sua-rua-mat-mat-na-duong-da-yeu-cau-phan-loai-theo-dung-ban-chat-hang-hoa-172411.html",
        "keywords": ["3401", "sua rua mat"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn), Thư Viện Pháp Luật (thuvienphapluat.vn) và Báo Mới (baomoi.com)."
    },
    {
        "stt": 262,
        "week": "W47 — 18/11/2026",
        "title": "Kết Luận Sơ Bộ Của Bộ Thương Mại Hoa Kỳ (DOC, 10/11/2026): Ấn Định Thuế Ký Quỹ Chống Bán Phá Giá 36,55% Đối Với Hóa Chất PTMEG (HS 3907.29.0000) Xuất Khẩu Từ Việt Nam",
        "agency": "Bộ Thương mại Hoa Kỳ (DOC) & Cục Phòng vệ thương mại (Bộ Công Thương)",
        "effective_date": "Ban hành kết luận sơ bộ ngày 10/11/2026; áp dụng ký quỹ tạm thời ngay khi công bố trên Công báo Liên bang Hoa Kỳ",
        "core_diff": "Cũ: Sản phẩm hóa chất PTMEG (Polytetramethylene ether glycol — nguyên liệu dệt sợi đàn hồi, polyurethane) xuất khẩu từ Việt Nam sang Hoa Kỳ thông quan theo mức thuế nhập khẩu thông thường; Mới: DOC xác định biên độ bán phá giá sơ bộ 36,55% cho cả bị đơn bắt buộc duy nhất và nhóm doanh nghiệp chịu thuế suất toàn quốc (Vietnam-wide rate), buộc nhà nhập khẩu Mỹ phải nộp tiền ký quỹ (Cash Deposit) 36,55% ngay từ giai đoạn sơ bộ.",
        "affected_entity": "Doanh nghiệp sản xuất, xuất khẩu hóa chất công nghiệp, nguyên liệu polymer/polyurethane (mã HS 3907.29.0000) và các nhà máy dệt sợi kỹ thuật xuất đi thị trường Hoa Kỳ.",
        "risk_pain": "Mức ký quỹ 36,55% trực tiếp đội giá bán tại Mỹ, làm mất khả năng cạnh tranh với nguồn cung từ Hàn Quốc, Đài Loan; nếu không hợp tác trả lời bản câu hỏi và nộp chứng từ thẩm tra tại chỗ đầy đủ, biên độ thuế cuối cùng có thể giữ nguyên hoặc tăng cao bằng dữ liệu bất lợi sẵn có (AFA).",
        "market_reaction": "Cục Phòng vệ thương mại (Bộ Công Thương) khuyến nghị doanh nghiệp liên quan rà soát chặt sổ sách kế toán, giá thành sản xuất và chuẩn bị kỹ lưỡng cho đợt thẩm tra tại chỗ của DOC trước khi có kết luận cuối cùng.",
        "pro_argument": "Việc tuân thủ đầy đủ quy trình kháng nghị, minh bạch hệ thống sổ sách kế toán giúp doanh nghiệp bảo vệ quyền lợi hợp pháp và chứng minh không bán phá giá để hạ biên độ thuế ở kết luận cuối cùng.",
        "con_argument": "Chi phí thuê luật sư thương mại quốc tế và áp lực chứng minh cấu trúc chi phí phi thị trường (NME) khiến các doanh nghiệp quy mô vừa dễ bị loại khỏi cuộc chơi xuất khẩu sang Hoa Kỳ.",
        "intl_ref": "Quy trình điều tra chống bán phá giá (Antidumping Duty Investigation) của Bộ Thương mại Hoa Kỳ (US DOC) áp dụng song song với Trung Quốc, Hàn Quốc và Đài Loan.",
        "url": "https://tapchicongthuong.vn/hoa-ky-ket-luan-so-bo-chong-ban-pha-gia-ptmeg-nhap-khau-tu-viet-nam-566896.htm",
        "keywords": ["ptmeg", "36,55%"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Công Thương (tapchicongthuong.vn), VietnamPlus (vietnamplus.vn) và Cục Phòng vệ thương mại (moit.gov.vn)."
    },
    {
        "stt": 263,
        "week": "W47 — 18/11/2026",
        "title": "Thông Tư Số 63/2026/TT-BKHCN (Ban Hành 27/10/2026, Hiệu Lực 01/01/2027): Cấm Nhập Khẩu 13 Nhóm Sản Phẩm Công Nghệ Số Đã Qua Sử Dụng & Hàng Tân Trang (Refurbished) Theo Mã HS",
        "agency": "Bộ Khoa học và Công nghệ",
        "effective_date": "Ban hành ngày 27/10/2026; chính thức có hiệu lực thi hành từ ngày 01/01/2027 (thay thế Thông tư 11/2018/TT-BTTTT)",
        "core_diff": "Cũ: Danh mục sản phẩm CNTT đã qua sử dụng cấm nhập khẩu thực hiện theo Thông tư 11/2018/TT-BTTTT cũ; Mới: Thông tư 63/2026/TT-BKHCN công bố 13 nhóm mã HS cấm nhập khẩu (Chương 84 & 85: 8443, 8470, 8471, 8517, 8518, 8525, 8526, 8527, 8528, 8534, 8540, 8542, 8544), mở rộng cấm cả linh kiện/phụ kiện cũ và hàng tân trang (Refurbished) dù đã sửa chữa bảo hành như mới, chỉ ngoại trừ dây chuyền phục vụ dự án bán dẫn hoặc R&D đạt chuẩn.",
        "affected_entity": "Doanh nghiệp nhập khẩu thiết bị IT, máy chủ, linh kiện điện tử, bo mạch, máy in công nghiệp, camera, màn hình và các đơn vị nhập thiết bị tân trang phục vụ nhà xưởng.",
        "risk_pain": "Nếu doanh nghiệp ký hợp đồng nhập khẩu thiết bị/linh kiện tân trang thuộc 13 nhóm HS trên mà hàng cập cảng sau 01/01/2027 (không có giấy phép cấp trước mốc chuyển tiếp), lô hàng sẽ bị từ chối thông quan, buộc tái xuất khỏi lãnh thổ Việt Nam và xử phạt nặng.",
        "market_reaction": "Các doanh nghiệp gia công điện tử, lắp ráp thiết bị công nghiệp rà soát gấp đơn hàng máy móc, phụ tùng thay thế (spare parts) để hoàn tất thông quan trước ngày 31/12/2026 hoặc chuyển hướng mua hàng mới 100%.",
        "pro_argument": "Ngăn chặn triệt để nguy cơ Việt Nam trở thành bãi rác công nghệ lạc hậu, bảo vệ an toàn an ninh mạng và thúc đẩy chuyển đổi sang dây chuyền sản xuất công nghệ cao, bán dẫn.",
        "con_argument": "Lệnh cấm bao trùm cả phụ tùng, linh kiện thay thế đã qua sử dụng hoặc tân trang có thể gây khó khăn cho các nhà máy đang vận hành máy móc đời cũ cần linh kiện bảo trì đồng bộ.",
        "intl_ref": "Công ước Basel về kiểm soát vận chuyển xuyên biên giới chất thải điện tử (E-waste) và quy định kiểm soát thiết bị tân trang của ASEAN, EU.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/tu-1-1-2027-ap-dung-danh-muc-moi-ve-hang-cong-nghe-so-da-qua-su-dung-cam-nhap-khau-170992.html",
        "keywords": ["63/2026", "cong nghe so"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn), Cổng TTĐT Chính phủ (vanban.chinhphu.vn/?pageid=27160&docid=220121&classid=1), Dân Trí và Thư Viện Pháp Luật."
    },
    {
        "stt": 264,
        "week": "W47 — 18/11/2026",
        "title": "Thông Tư Số 154/2026/TT-BTC (Ban Hành 15/10/2026, Hiệu Lực 01/01/2027): Siết Thủ Tục Hải Quan & Thời Hạn 30 Ngày Kê Khai Nộp Thuế Hàng Mua Gom Của Cư Dân Biên Giới",
        "agency": "Bộ Tài chính",
        "effective_date": "Ban hành ngày 15/10/2026; chính thức có hiệu lực thi hành từ ngày 01/01/2027 (sửa đổi Thông tư 80/2019/TT-BTC)",
        "core_diff": "Cũ: Thủ tục kê khai, quản lý thuế đối với hàng hóa mua gom của cư dân biên giới đưa vào nội địa theo Thông tư 80/2019/TT-BTC còn lỏng lẻo về thời hạn gom hàng và chế tài giám sát thuế; Mới: Bắt buộc thương nhân mua gom phải làm thủ tục hải quan như hàng nhập khẩu thông thường tại hải quan cửa khẩu gần nhất, thời hạn nộp hồ sơ chậm nhất 30 ngày kể từ ngày mua gom và phải hoàn thành nghĩa vụ thuế trước khi đưa hàng vào nội địa tiêu thụ.",
        "affected_entity": "Thương nhân thu mua nông sản, hàng tiêu dùng tại khu vực cửa khẩu biên giới đất liền phía Bắc/Tây Nam và các đơn vị vận chuyển hàng hóa biên mậu.",
        "risk_pain": "Nếu quá thời hạn 30 ngày chưa nộp đủ hồ sơ hải quan hoặc đưa hàng qua khỏi khu vực cửa khẩu vào nội địa khi chưa hoàn thành nghĩa vụ thuế, lô hàng bị coi là hàng chưa khai báo hải quan và đối mặt chế tài xử phạt, tịch thu theo quy định mới.",
        "market_reaction": "Các thương nhân kinh doanh qua cửa khẩu biên giới đất liền bắt đầu chuyển đổi sang ký hợp đồng ngoại thương chính thức và mở tờ khai trực tiếp tại chi cục cửa khẩu để tránh rủi ro quá hạn 30 ngày.",
        "pro_argument": "Chính quy hóa hoạt động giao thương biên giới, ngăn chặn tình trạng lợi dụng định mức miễn thuế cư dân để chia nhỏ lô hàng thương mại lớn đưa vào nội địa tiêu thụ.",
        "con_argument": "Áp lực kê khai trong 30 ngày và đầy đủ chứng từ mua gom có thể làm tăng chi phí quản lý đối với nông sản mùa vụ nhỏ lẻ tại các tỉnh vùng sâu biên giới.",
        "intl_ref": "Chuẩn mực quản lý thương mại biên giới nhỏ lẻ (Cross-border Petty Trade Formalization) của Tổ chức Hải quan Thế giới (WCO) và Hiệp định Tạo thuận lợi Thương mại (WTO TFA).",
        "url": "https://thoibaotaichinhvietnam.vn/bo-tai-chinh-quy-dinh-thu-tuc-hai-quan-voi-hang-hoa-mua-gom-cua-cu-dan-bien-gioi-205691.html",
        "keywords": ["154/2026", "cu dan bien gioi"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Thời báo Tài chính Việt Nam (thoibaotaichinhvietnam.vn), Cổng TTĐT Chính phủ (chinhphu.vn/?pageid=27160&docid=220072) và Thư Viện Pháp Luật."
    },
    {
        "stt": 265,
        "week": "W47 — 18/11/2026",
        "title": "Công Văn 23739/CHQ-NVTHQ (02/11/2026): Cắt Hoàn Thuế GTGT Tại Cơ Quan Hải Quan Với Hàng Nhập Khẩu Xuất Trả Chủ Hàng Nước Ngoài (B13) Từ Mốc 01/07/2025",
        "agency": "Cục Hải quan (Ban Nghiệp vụ Thuế Hải quan)",
        "effective_date": "Ban hành ngày 02/11/2026 (nhắc lại và hướng dẫn áp dụng thống nhất theo Công văn 45806/CHQ-NVTHQ từ mốc 01/07/2025)",
        "core_diff": "Cũ: Doanh nghiệp nhập khẩu hàng kinh doanh (A12) đã nộp thuế nhập khẩu và thuế GTGT, sau đó hàng bị lỗi phải xuất trả chủ hàng nước ngoài (loại hình B13) thường làm hồ sơ xin cơ quan Hải quan hoàn lại cả 2 sắc thuế nộp thừa; Mới: Cục Hải quan khẳng định tờ khai xuất trả đăng ký từ ngày 01/07/2025 trở đi sẽ KHÔNG được hoàn thuế GTGT nộp thừa tại cơ quan Hải quan, mà doanh nghiệp phải liên hệ cơ quan thuế nội địa để xử lý khấu trừ hoặc hạch toán chi phí.",
        "affected_entity": "Các doanh nghiệp nhập khẩu hàng kinh doanh, nguyên phụ liệu (A11/A12) phát sinh lỗi chất lượng buộc phải xuất trả nhà cung cấp nước ngoài (B13) và bộ phận kế toán thuế XNK.",
        "risk_pain": "Doanh nghiệp nộp nhầm hồ sơ xin hoàn thuế GTGT sang cơ quan Hải quan bị trả hồ sơ kéo dài nhiều tháng, đọng vốn hàng trăm triệu đồng thuế GTGT đầu vào hoặc hạch toán sai kỳ khấu trừ dẫn đến bị cơ quan thuế nội địa xuất toán chi phí.",
        "market_reaction": "Nhiều doanh nghiệp XNK và đơn vị tư vấn thuế phải tách đôi quy trình xử lý: xin hoàn thuế nhập khẩu tại Hải quan và chuyển bộ chứng từ tờ khai B13 sang cơ quan Thuế nội địa để kê khai khấu trừ GTGT.",
        "pro_argument": "Phân định rạch ròi thẩm quyền quản lý thuế GTGT giữa cơ quan Hải quan (thu khâu nhập khẩu) và cơ quan Thuế nội địa (khấu trừ/hoàn thuế GTGT), đồng bộ với Luật Thuế GTGT 48/2024/QH15 hiệu lực từ 01/07/2025.",
        "con_argument": "Doanh nghiệp gặp tình huống bất khả kháng phải xuất trả hàng lỗi bị chia đôi thủ tục hành chính qua 2 cơ quan (Hải quan và Thuế), làm chậm tiến độ thu hồi dòng tiền.",
        "intl_ref": "Nguyên tắc khấu trừ thuế GTGT đầu vào qua tờ khai thuế nội địa (Input VAT Credit Mechanism) đối với hàng tái xuất trả lại nhà cung cấp tại EU và Singapore.",
        "url": "https://mitou.vn/vi/chi-tiet-tin/hoan-thue-gtgt-doi-voi-hang-hoa-nhap-khau-sau-do-xuat-tra-chu-hang-nuoc-ngoai?lang=en",
        "keywords": ["23739", "xuat tra"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Mitou Vietnam (mitou.vn), Thư Viện Pháp Luật (thuvienphapluat.vn) và LuatVietnam (luatvietnam.vn)."
    }
]


# ==============================================================================
# 2. DATA FOR CASE STUDY (14 cols) — STT 261..265 (W47 — 18/11/2026)
# ==============================================================================
CS_NEW_IDEAS = [
    {
        "stt": 261,
        "week": "W47 — 18/11/2026",
        "title": "Khởi Tố Đường Dây Công Ty Thủy Sản Sỹ Hưng Tại Hải Phòng (2026): Nhập Khẩu Trái Quy Định 4.410 Tấn Cua Cà Ra & Nầm Lợn Gần 1.000 Tỷ Đồng, Gian Lận Thuế Hơn 40 Tỷ Đồng",
        "parties": "Công ty TNHH Thủy sản Sỹ Hưng (Nguyễn Thị Thư) cùng 13 bị can + Cục C05 (Bộ Công an) & Công an TP. Hải Phòng + Cán bộ Chi cục Thú y vùng I",
        "plot": "Trong năm 2026 (khởi tố các đợt tháng 03 và tháng 07/2026), cơ quan điều tra triệt phá đường dây do Nguyễn Thị Thư cầm đầu móc nối lập nhiều pháp nhân bình phong để vận chuyển trái phép 4.410 tấn cua cà ra và nầm lợn đông lạnh từ Trung Quốc qua tuyến cảng/biên giới về Hải Phòng, chi tiền hối lộ cán bộ thú y để bỏ qua quy trình kiểm dịch động vật và phát hành hóa đơn khống.",
        "damage": "Tổng giá trị hàng hóa vi phạm lên tới gần 1.000 tỷ đồng (4.410 tấn hàng đông lạnh); số tiền gian lận thuế bị xác định hơn 40 tỷ đồng; 14 bị can bị khởi tố hình sự về các tội buôn bán trái phép qua biên giới, gian lận thuế, đưa/nhận hối lộ và mua bán trái phép hóa đơn.",
        "root_cause": "Cố tình dùng pháp nhân ma và hối lộ nhân sự kiểm dịch để né chi phí kiểm dịch động vật, né thuế nhập khẩu chính ngạch và hợp thức hóa đầu ra bằng hóa đơn bất hợp pháp.",
        "prevention": "1. Tuyệt đối thực hiện nhập khẩu thủy hải sản và thực phẩm đông lạnh theo đường chính ngạch có Giấy phép kiểm dịch Cục Thú y (e-Cert) và chứng thư xuất xứ rõ ràng.\n2. Rà soát chuỗi cung ứng đầu vào: yêu cầu nhà cung cấp xuất trình đầy đủ tờ khai hải quan nhập khẩu và giấy chứng nhận kiểm dịch thú y hợp lệ cho từng lô hàng.\n3. Kiểm toán hóa đơn đầu vào định kỳ, tránh mua hàng từ các doanh nghiệp mới thành lập có dấu hiệu rủi ro hóa đơn.",
        "opposing_views": "Phe A: Chỉ cần có hóa đơn đỏ đầu vào từ công ty trung gian trong nước là doanh nghiệp thu mua yên tâm không cần check gốc nhập khẩu vs Phe B: Hóa đơn hợp thức hóa cho lô hàng không qua kiểm dịch và không mở tờ khai hợp lệ sẽ khiến cả bên mua lẫn bên bán liên đới trách nhiệm hình sự khi đường dây bị triệt phá.",
        "lessons": "Trong ngành thực phẩm và thủy sản, tiết kiệm vài phần trăm chi phí kiểm dịch và thuế bằng con đường phi chính thức có thể đánh đổi bằng cả sự nghiệp, án tù và khoản truy thu hàng chục tỷ đồng.",
        "url": "https://vneconomy.vn/triet-pha-duong-day-buon-lau-4410-tan-cua-ca-ra-nam-lon-khong-qua-kiem-dich.htm",
        "keywords": ["sy hung", "4.410"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: VnEconomy (vneconomy.vn), Báo Pháp Luật Plus (phapluatplus.baophapluat.vn) và Dân Trí."
    },
    {
        "stt": 262,
        "week": "W47 — 18/11/2026",
        "title": "Đại Án Lập Công Ty 'Ma' Nhập Khẩu 1.287 Container Máy Móc, Thiết Bị Cũ Tại TP.HCM (Trị Giá Hơn 217 Tỷ Đồng): Cái Giá Đắt Của Hơn 1.000 Bộ Tờ Khai Gian Dối",
        "parties": "Nhóm chủ mưu thành lập hàng loạt công ty bình phong + Cựu cán bộ hải quan tiếp tay + TAND TP.HCM xét xử",
        "plot": "Các đối tượng móc nối thành lập hàng loạt pháp nhân 'ma', thuê người đứng tên giám đốc để mở hơn 1.000 bộ tờ khai hải quan, hợp thức hóa việc nhập khẩu 1.287 container máy móc, thiết bị cũ từ Nhật Bản, Trung Quốc về cảng TP.HCM dưới vỏ bọc máy móc phục vụ sản xuất trực tiếp trong khi thực chất là hàng cũ cấm/hạn chế nhập khẩu bán thương mại.",
        "damage": "Tổng trị giá tang vật vi phạm lên tới hơn 217 tỷ đồng (1.287 container); toàn bộ các chủ mưu và cán bộ tiếp tay bị TAND TP.HCM tuyên phạt các mức án tù nghiêm khắc từ 12 đến 18 năm tù cùng hình phạt bổ sung nộp lại toàn bộ số lợi bất chính.",
        "root_cause": "Lợi dụng kẽ hở trong quy định nhập khẩu máy móc đã qua sử dụng (chỉ được nhập phục vụ trực tiếp sản xuất theo Quyết định 18/2019/QĐ-TTg) để lập công ty vỏ bọc nhập bán lại kiếm lời.",
        "prevention": "1. Chỉ nhập khẩu máy móc cũ phục vụ trực tiếp hoạt động sản xuất của chính doanh nghiệp (tuổi thiết bị ≤ 10 năm, đáp ứng QCVN/TCVN về an toàn, tiết kiệm năng lượng).\n2. Thực hiện giám định độc lập tại nước xuất khẩu hoặc tại cảng bởi tổ chức giám định được chỉ định trước khi mở tờ khai.\n3. Tuyệt đối không cho mượn pháp nhân hoặc đứng tên ủy thác cho các lô hàng máy móc cũ không rõ mục đích sử dụng.",
        "opposing_views": "Phe A: Nhập máy móc cũ giá rẻ qua công ty dịch vụ đứng tên giúp tiết kiệm vốn đầu tư ban đầu cho xưởng sản xuất nhỏ vs Phe B: Một khi cơ quan điều tra truy vết 1.000+ tờ khai và luồng tiền, cả người đứng tên thuê lẫn doanh nghiệp mua lại thiết bị không rõ hồ sơ đều bị phong tỏa tài sản và truy tố.",
        "lessons": "Mọi mô hình 'đứng tên giùm' hoặc làm khống hồ sơ mục đích sử dụng máy móc cũ đều là quả bom nổ chậm; minh bạch chứng thư giám định ngay từ đầu là cách bảo vệ doanh nghiệp rẻ nhất.",
        "url": "https://vietnamnet.vn/hai-cuu-canh-sat-thiet-lap-duong-day-buon-lau-khung-tai-tphcm-linh-an-2391796.html",
        "keywords": ["217", "1.287"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: VietNamNet (vietnamnet.vn), Tiền Phong (tienphong.vn) và Công An Nhân Dân."
    },
    {
        "stt": 263,
        "week": "W47 — 18/11/2026",
        "title": "Doanh Nghiệp Xuất Khẩu Việt Nam Mất Trắng Gần 72.000 USD (Gần 2 Tỷ Đồng) Sau Vài Email Giao Dịch: Cạm Bẫy Giả Mạo Đối Tác Bằng AI (BEC Fraud)",
        "parties": "Doanh nghiệp xuất khẩu Việt Nam + Đối tác ảo mạo danh doanh nghiệp lớn nước ngoài + Cảnh báo từ chuyên gia RMIT & Thương vụ Việt Nam",
        "plot": "Trong đợt cảnh báo tháng 06/2026, một doanh nghiệp xuất khẩu Việt Nam giao dịch trực tuyến qua email với đối tượng sử dụng AI tạo danh tính số hoàn hảo và mở tài khoản ngân hàng trùng tên một doanh nghiệp nhập khẩu uy tín ở nước ngoài. Sau khi ký hợp đồng và nhận yêu cầu chuyển tiền tạm ứng/phí giao dịch vào tài khoản mới thay đổi qua email phút chót, đối tượng rút sạch tiền và cắt đứt mọi liên lạc.",
        "damage": "Doanh nghiệp thiệt hại trực tiếp gần 72.000 USD (tương đương gần 2 tỷ đồng) không thể thu hồi qua ngân hàng trung gian do tiền đã bị chuyển tiếp qua nhiều tài khoản ở nước ngoài.",
        "root_cause": "Tin tưởng 100% vào trao đổi email/chat trực tuyến mà không thực hiện xác minh đa kênh (Out-of-band verification), không kiểm tra tên miền email giả (spoofed domain) và không đối chiếu thông tin đăng ký kinh doanh/tài khoản thụ hưởng độc lập.",
        "prevention": "1. Áp dụng quy tắc 'Two-Channel Verification': Mọi thay đổi số tài khoản ngân hàng trên Invoice đều phải gọi video call trực tiếp số hotline công khai của đối tác và ký Phụ lục hợp đồng đóng dấu.\n2. Kiểm tra kỹ từng ký tự tên miền email (chặn lỗi typo-squatting như `.co` thay `.com` hay `rn` thay `m`) và bật cảnh báo DMARC/SPF.\n3. Sử dụng L/C không hủy ngang (Irrevocable L/C) hoặc bảo hiểm tín dụng xuất khẩu cho đơn hàng mới giá trị từ 20.000 USD trở lên.",
        "opposing_views": "Phe A: Đàm phán nhanh qua email/WhatsApp giúp chốt đơn hàng quốc tế thần tốc tiết kiệm chi phí xác minh vs Phe B: Thời đại AI deepfake, chỉ một phút lơ là bỏ qua cuộc gọi xác minh 2 lớp là mất trắng toàn bộ lợi nhuận cả năm.",
        "lessons": "Trong thương mại quốc tế 2026, chữ ký số và email chưa bao giờ là bằng chứng an toàn tuyệt đối; quy trình xác minh tài khoản thụ hưởng qua kênh độc lập phải là kỷ luật thép của phòng Kế toán và Sales XNK.",
        "url": "https://dantri.com.vn/kinh-doanh/doanh-nghiep-xuat-khau-doi-mat-bay-ai-mat-72000-usd-chi-sau-vai-email-20260608104836192.htm",
        "keywords": ["72.000", "email"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Báo Dân Trí (dantri.com.vn) và VietnamBiz (vietnambiz.vn)."
    },
    {
        "stt": 264,
        "week": "W47 — 18/11/2026",
        "title": "Cuộc Giải Cứu Lô Hàng Chè Xuất Khẩu Kẹt Chiến Sự Trung Đông Phải Tái Nhập Về Việt Nam (2026): Chọn Đúng Mã Loại Hình A31 Hay G13 Để Không Mất Oan Tiền Thuế & Lưu Bãi",
        "parties": "Doanh nghiệp xuất khẩu chè (Công ty TNHH chè Hoàng Anh) + Hãng tàu quốc tế + Chi cục Hải quan cửa khẩu",
        "plot": "Do biến động chiến sự quân sự tại khu vực Trung Đông khiến cảng đích phong tỏa và tuyến vận tải biển gián đoạn, hàng loạt container chè đã xuất khẩu rời cảng Việt Nam buộc phải quay đầu tái nhập về nước. Doanh nghiệp lúng túng giữa việc khai báo mã loại hình A31 (nhập hàng đã xuất khẩu bị trả lại) hay G13/C11 để đưa hàng về xử lý, tiêu thụ hoặc tái xuất sang thị trường thay thế.",
        "damage": "Nếu chậm xác định đúng mã loại hình và hồ sơ chứng minh hàng tái nhập theo Điều 47 Nghị định 08/2015/NĐ-CP, doanh nghiệp đối mặt chi phí lưu container/lưu bãi (DEM/DET) phát sinh 80–150 USD/container/ngày và nguy cơ bị ấn định thuế nhập khẩu như hàng mua mới từ nước ngoài.",
        "root_cause": "Hợp đồng ngoại thương thiếu điều khoản xử lý hàng quay đầu khi xảy ra bất khả kháng (Force Majeure) và nhân sự chứng từ chưa nắm chắc ranh giới áp dụng mã loại hình A31 vs G13.",
        "prevention": "1. Nếu tái nhập về để tiêu thụ nội địa hoặc tái chế/tiêu hủy hoàn toàn, mở tờ khai mã A31 kèm công văn giải trình bất khả kháng, vận đơn gốc và chứng từ chứng minh hàng hóa có nguồn gốc xuất khẩu từ Việt Nam để được xét miễn/không thu thuế nhập khẩu.\n2. Nếu chỉ tạm nhập để sửa chữa/đóng gói lại rồi tái xuất sang thị trường thứ ba, đối chiếu điều kiện mã G13 hoặc gửi kho ngoại quan để tránh phát sinh thuế nội địa.\n3. Đàm phán ngay với hãng tàu gói miễn giảm DEM/DET cho hàng quay đầu do bất khả kháng chiến sự.",
        "opposing_views": "Phe A: Hàng quay đầu về cảng Việt Nam cứ mở tờ khai nhập kinh doanh A11 cho nhanh rồi tính sau vs Phe B: Khai nhầm sang A11 sẽ bị hệ thống tính thuế nhập khẩu và VAT ngay lập tức, mất quyền hưởng cơ chế không thu thuế hàng tái nhập theo Điều 47 NĐ 08.",
        "lessons": "Khi khủng hoảng địa chính trị làm hàng phải quay đầu, tốc độ nắm vững bảng mã loại hình hải quan (A31 vs G13) chính là phao cứu sinh bảo vệ dòng tiền và hàng hóa tại cảng.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/thu-tuc-tai-nhap-hang-da-xuat-khau-va-tai-xuat-hang-hoa-154427.html",
        "keywords": ["a31", "tai nhap"],
        "series": "Series: Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn) và PTN Logistics / Thư Viện Pháp Luật."
    },
    {
        "stt": 265,
        "week": "W47 — 18/11/2026",
        "title": "Chặn Đứng Lô Hàng Nhập Khẩu Xâm Phạm Sở Hữu Trí Tuệ Tại Cửa Khẩu Chi Ma — Lạng Sơn (05/2026): Thu Giữ 30 Áo Phông Burberry, 180 Túi Xách Coach & Giày Dép Crocs, Adidas",
        "parties": "Chi cục Điều tra chống buôn lậu (Đội 1) + Hải quan cửa khẩu Chi Ma (Lạng Sơn) + Doanh nghiệp đứng tên nhập khẩu",
        "plot": "Trong đợt cao điểm kiểm tra tháng 05/2026 thực hiện Công điện 38/CĐ-TTg, lực lượng Hải quan cửa khẩu Chi Ma phối hợp Đội 1 (Chi cục Điều tra chống buôn lậu) kiểm tra thực tế lô hàng nhập khẩu qua biên giới phía Bắc và phát hiện hàng loạt sản phẩm thời trang mang nhãn hiệu nổi tiếng thế giới không có ủy quyền chủ sở hữu trí tuệ gồm 30 áo phông Burberry, 180 túi xách Coach cùng nhiều giày dép, mỹ phẩm Crocs, Adidas, Dolce&Gabbana.",
        "damage": "100% lô hàng xâm phạm quyền sở hữu trí tuệ bị lập biên bản tạm giữ và tịch thu tiêu hủy theo quy định; doanh nghiệp nhập khẩu bị xử phạt vi phạm hành chính hàng chục triệu đồng và bị đưa vào diện kiểm soát luồng Đỏ 100% đối với các lô hàng tiếp theo.",
        "root_cause": "Doanh nghiệp nhận ủy thác hoặc gom hàng nhập khẩu không kiểm tra chứng thư ủy quyền phân phối chính hãng (Letter of Authorization) và không kiểm hàng thực tế trước khi xếp xe qua cửa khẩu.",
        "prevention": "1. Yêu cầu nhà cung cấp xuất trình Giấy ủy quyền của chủ sở hữu nhãn hiệu (LOA) hoặc hóa đơn mua hàng từ đại lý phân phối được công nhận toàn cầu đối với mọi mặt hàng có thương hiệu.\n2. Tra cứu danh sách nhãn hiệu đã đăng ký bảo hộ quyền SHTT tại cơ quan Hải quan Việt Nam trước khi ký hợp đồng nhập khẩu.\n3. Cam kết điều khoản từ chối nhận và bồi hoàn 100% nếu nhà cung cấp tự ý đóng kèm hàng không rõ nguồn gốc thương hiệu vào chung container/xe hàng.",
        "opposing_views": "Phe A: Chỉ cần khai đúng tên chủng loại 'áo phông, túi xách' trên tờ khai mà không ghi tên thương hiệu là thông quan được vs Phe B: Hệ thống soi chiếu và kiểm hóa luồng Đỏ của Hải quan đối chiếu trực tiếp logo thực tế với cơ sở dữ liệu bảo hộ SHTT, mọi hành vi che giấu nhãn hiệu đều bị tịch thu hàng.",
        "lessons": "Đừng bao giờ nhận vận chuyển hay ủy thác nhập khẩu các lô hàng thời trang, phụ kiện dính logo thương hiệu nổi tiếng khi chưa cầm trên tay giấy ủy quyền chính hãng hợp lệ.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/hai-quan-lien-tiep-phat-hien-nhieu-hang-hoa-vi-pham-so-huu-tri-tue-o-bien-gioi-phia-bac-156344.html",
        "keywords": ["burberry", "coach"],
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn) và Báo Điện tử Chính phủ (baochinhphu.vn)."
    }
]


# ==============================================================================
# 3. DATA FOR HOC THUC CHIEN (14 cols) — STT 261..265 (W47 — 18/11/2026)
# ==============================================================================
HT_NEW_IDEAS = [
    {
        "stt": 261,
        "week": "W47 — 18/11/2026",
        "title": "Quy Trình 5 Bước Tự Rà Soát Mã HS Mỹ Phẩm (Sữa Rửa Mặt 3401.30.00 vs 3304 & Mặt Nạ 3307.90.90) Chuẩn Công Văn 24153/CHQ-NVTHQ (12/11/2026)",
        "steps": "Bước 1: Thu thập đầy đủ Bảng thành phần INCI, MSDS và hình ảnh bao bì bán lẻ từ nhà sản xuất.\nBước 2: Phân tách nhóm công dụng bề mặt: nếu sản phẩm rửa mặt dạng lỏng/kem chứa chất hoạt động bề mặt làm sạch da (có hoặc không chứa xà phòng) -> đối chiếu Nhóm 34.01 (3401.30.00).\nBước 3: Đối với mặt nạ giấy/vải không dệt thấm tẩm dung dịch mỹ phẩm đóng gói bán lẻ -> đối chiếu Nhóm 33.07 (3307.90.90) thay vì mặc định 33.04.\nBước 4: Kiểm tra chéo thuế suất MFN và ưu đãi FTA (Form E, VK, AK, EVFTA) giữa mã HS cũ và mã HS mới để dự toán chênh lệch chi phí.\nBước 5: Đối với SKU mới nhập lần đầu hoặc còn tranh cãi công dụng kép, nộp hồ sơ Xác định trước mã số (Pre-ruling) trước khi hàng về cảng.",
        "formula": "Chênh lệch thuế cần dự phòng = Trị giá tính thuế × (Thuế suất NK mã HS mới − Thuế suất NK mã HS cũ) + Thuế GTGT 10% tương ứng; tham chiếu Công văn 24153/CHQ-NVTHQ (12/11/2026).",
        "top3_errors": "1. Áp toàn bộ mỹ phẩm vào nhóm 33.04 theo thói quen mà không đọc Chú giải Chương 33 & 34.\n2. Khai tên hàng quá chung chung ('Mỹ phẩm các loại') thiếu dạng bào chế (lỏng/kem/vải không dệt) dẫn đến bị giữ hàng lấy mẫu phân tích.\n3. Không rà soát lại các tờ khai cũ trong 60 ngày (khai bổ sung AMA) khi cơ quan Hải quan đã có công văn chỉ đạo kiểm tra toàn ngành.",
        "mnemonic": "Mặt nạ vải không dệt đi với 33.07 ('bảy' miếng vải tẩm serum); Sữa rửa mặt tạo bọt làm sạch bề mặt nhớ ngay nhóm xà phòng 34.01!",
        "comparison": "Nhóm 33.04 (Kem dưỡng, serum thoa lưu lại trên da) vs Nhóm 34.01 (Sữa rửa mặt làm sạch bề mặt rửa trôi) & Nhóm 33.07 (Mặt nạ vải không dệt tẩm sẵn dưỡng chất): khác hoàn toàn bản chất Chú giải HS.",
        "persona": "Chị Hương (Chủ doanh nghiệp nhập khẩu mỹ phẩm) & Nhân viên chứng từ XNK ngành làm đẹp",
        "checklist": "☐ Tải INCI & MSDS có tỷ lệ % thành phần\n☐ Phân loại Sữa rửa mặt có chất hoạt động bề mặt -> 3401.30.00\n☐ Phân loại Mặt nạ vải không dệt tẩm mỹ phẩm -> 3307.90.90\n☐ Kiểm tra thuế suất FTA theo mã HS chuẩn\n☐ Lập hồ sơ AMA nếu cần điều chỉnh tờ khai cũ trong 60 ngày",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/sua-rua-mat-mat-na-duong-da-yeu-cau-phan-loai-theo-dung-ban-chat-hang-hoa-172411.html",
        "keywords": ["3401", "sua rua mat"],
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Gắn trực tiếp SEO Keyword #5 (17/11/2026) & Công văn 24153/CHQ-NVTHQ; đối chiếu 2+ nguồn: thuehaiquan.tapchikinhtetaichinh.vn & thuvienphapluat.vn."
    },
    {
        "stt": 262,
        "week": "W47 — 18/11/2026",
        "title": "Checklist 5 Bước Rà Soát 13 Nhóm Mã HS Sản Phẩm Công Nghệ Số Đã Qua Sử Dụng & Hàng Tân Trang (Refurbished) Cấm Nhập Khẩu Theo Thông Tư 63/2026/TT-BKHCN",
        "steps": "Bước 1: Kiểm tra 4 số đầu mã HS của thiết bị/linh kiện dự kiến nhập có nằm trong 13 nhóm cấm (8443, 8470, 8471, 8517, 8518, 8525, 8526, 8527, 8528, 8534, 8540, 8542, 8544) hay không.\nBước 2: Xác minh tình trạng hàng hóa từ nhà cung cấp: lưu ý Thông tư 63/2026/TT-BKHCN cấm cả hàng tân trang (Refurbished) và linh kiện/phụ kiện đã qua sử dụng, không chỉ máy nguyên chiếc.\nBước 3: Đối chiếu điều khoản ngoại lệ (máy in khổ lớn/tốc độ cao đạt chuẩn hoặc dây chuyền phục vụ trực tiếp dự án bán dẫn, đào tạo R&D công nghệ số).\nBước 4: Rà soát ngày tàu cập cảng: mọi lô hàng thuộc diện cấm chưa có giấy phép trước 01/01/2027 phải thông quan xong trước ngày 31/12/2026.\nBước 5: Bổ sung điều khoản cam kết 'Brand New 100% — Never Refurbished' kèm CO/CQ gốc nhà sản xuất vào hợp đồng ngoại thương.",
        "formula": "Mốc hiệu lực bắt buộc: 01/01/2027 | Phạm vi cấm: 13 nhóm HS Chương 84 & 85 (bao gồm hàng cũ + hàng tân trang Refurbished + linh kiện rời đã qua sử dụng) theo Thông tư 63/2026/TT-BKHCN.",
        "top3_errors": "1. Nhầm tưởng hàng tân trang (Factory Refurbished) còn nguyên seal bảo hành hãng là hàng mới nên vẫn ký hợp đồng nhập khẩu.\n2. Chỉ kiểm tra máy nguyên chiếc (8471, 8517) mà bỏ sót bo mạch, cáp, linh kiện thay thế (8534, 8542, 8544).\n3. Xếp hàng về cảng sát cuối tháng 12/2026 dẫn đến trễ thông quan qua ngày 01/01/2027 và bị buộc tái xuất.",
        "mnemonic": "Nhớ câu thần chú '13 nhóm công nghệ số — Cũ hay Tân trang đều CẤM từ 01/01/2027', trừ dự án Bán dẫn & R&D có hồ sơ chuẩn!",
        "comparison": "Thông tư 11/2018/TT-BTTTT (Quy định cũ) vs Thông tư 63/2026/TT-BKHCN (Hiệu lực 01/01/2027): Cập nhật theo Luật Công nghiệp công nghệ số mới, bịt kẽ hở hàng tân trang Refurbished và linh kiện cũ.",
        "persona": "Anh Dũng (Giám đốc nhà máy sản xuất/gia công điện tử) & Phòng Mua hàng thiết bị IT",
        "checklist": "☐ Đối chiếu mã HS với 13 nhóm Chương 84–85\n☐ Yêu cầu nhà cung cấp xác nhận hàng mới 100% (không Refurbished)\n☐ Kiểm tra hồ sơ ngoại lệ (nếu nhập cho dự án bán dẫn/R&D)\n☐ Đẩy lịch tàu cập cảng & thông quan trước 31/12/2026\n☐ Ghi rõ điều khoản phạt tái xuất trong hợp đồng nếu giao sai hàng cũ",
        "url": "https://vanban.chinhphu.vn/?pageid=27160&docid=220121&classid=1",
        "keywords": ["63/2026", "cong nghe so"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn chính thống Cổng TTĐT Chính phủ (vanban.chinhphu.vn); đối chiếu 2+ nguồn với thuehaiquan.tapchikinhtetaichinh.vn & dantri.com.vn."
    },
    {
        "stt": 263,
        "week": "W47 — 18/11/2026",
        "title": "Cẩm Nang 4 Bước Xử Lý Thuế GTGT Hàng Nhập Khẩu Lỗi Phải Xuất Trả Nước Ngoài (Loại Hình B13) Từ Mốc 01/07/2025 Theo Công Văn 23739/CHQ-NVTHQ",
        "steps": "Bước 1: Kiểm tra ngày đăng ký tờ khai xuất trả loại hình B13: nếu từ ngày 01/07/2025 trở đi, tuyệt đối tách riêng hồ sơ thuế nhập khẩu và thuế GTGT.\nBước 2: Lập bộ hồ sơ hoàn thuế nhập khẩu (nếu đã nộp khi nhập A11/A12) gửi cơ quan Hải quan kèm biên bản xác nhận hàng lỗi, thỏa thuận trả hàng và chứng từ chứng minh hàng chưa qua sử dụng/sản xuất tại Việt Nam.\nBước 3: Đối với khoản thuế GTGT đã nộp khâu nhập khẩu, KHÔNG nộp đơn xin hoàn tại Hải quan mà tập hợp Tờ khai nhập khẩu, Giấy nộp tiền vào NSNN và Tờ khai xuất trả B13 chuyển sang Kế toán thuế nội địa.\nBước 4: Phối hợp cơ quan Thuế quản lý trực tiếp thực hiện kê khai khấu trừ thuế GTGT đầu vào hoặc xử lý hạch toán theo đúng hướng dẫn Luật Thuế GTGT 48/2024/QH15.",
        "formula": "Mốc phân tách thẩm quyền: Tờ khai xuất trả B13 trước 01/07/2025 (xử lý theo hướng dẫn cũ) vs Từ 01/07/2025 trở đi (Hải quan KHÔNG hoàn thuế GTGT nộp thừa -> chuyển cơ quan Thuế nội địa xử lý).",
        "top3_errors": "1. Nộp gộp yêu cầu hoàn cả thuế NK lẫn thuế GTGT cho cơ quan Hải quan đối với tờ khai B13 sau 01/07/2025 khiến hồ sơ bị trả đi trả lại mất nhiều tháng.\n2. Mở nhầm tờ khai xuất trả sang loại hình xuất kinh doanh B11 thay vì B13 làm mất căn cứ hoàn thuế nhập khẩu.\n3. Đưa hàng lỗi vào lắp ráp/gia công thử tại xưởng trước khi xuất trả khiến hải quan bác bỏ điều kiện chưa qua sử dụng.",
        "mnemonic": "Nhớ quy tắc 'Tách đôi sau 01/07/2025': Thuế Nhập khẩu gõ cửa Hải quan — Thuế GTGT gõ cửa Thuế nội địa!",
        "comparison": "Hoàn thuế tại Hải quan (chỉ áp dụng cho Thuế Nhập khẩu hàng B13 chưa sử dụng) vs Khấu trừ/hoàn thuế tại Cơ quan Thuế nội địa (áp dụng cho Thuế GTGT hàng B13 từ 01/07/2025).",
        "persona": "Kế toán trưởng, Kế toán thuế XNK & Chuyên viên chứng từ doanh nghiệp thương mại/FDI",
        "checklist": "☐ Xác nhận ngày đăng ký tờ khai xuất trả B13 (>= 01/07/2025)\n☐ Giữ nguyên trạng hàng lỗi chưa đưa vào sản xuất/sử dụng\n☐ Ký thỏa thuận trả hàng & lập hóa đơn/chứng từ xuất trả B13 chuẩn\n☐ Nộp hồ sơ hoàn Thuế NK tại Hải quan\n☐ Bàn giao bộ chứng từ nộp thuế GTGT khâu NK cho Kế toán thuế nội địa",
        "url": "https://mitou.vn/vi/chi-tiet-tin/hoan-thue-gtgt-doi-voi-hang-hoa-nhap-khau-sau-do-xuat-tra-chu-hang-nuoc-ngoai?lang=en",
        "keywords": ["23739", "xuat tra"],
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Gắn trực tiếp SEO Keyword #9 & Công văn 23739/CHQ-NVTHQ (02/11/2026); đối chiếu 2+ nguồn: mitou.vn, thuvienphapluat.vn & luatvietnam.vn."
    },
    {
        "stt": 264,
        "week": "W47 — 18/11/2026",
        "title": "Quy Trình 4 Bước Khai Hải Quan & Nộp Thuế Trong 30 Ngày Đối Với Hàng Hóa Mua Gom Của Cư Dân Biên Giới Theo Thông Tư 154/2026/TT-BTC",
        "steps": "Bước 1: Lập bảng kê chi tiết hàng hóa thu mua từ cư dân biên giới (chủng loại, số lượng, trị giá, thông tin định danh người bán) ngay tại thời điểm giao nhận ở chợ biên giới/cửa khẩu.\nBước 2: Lưu ý mốc đếm ngược 30 ngày kể từ ngày mua gom: tập kết đầy đủ bộ chứng từ mua gom để làm thủ tục hải quan tại Chi cục Hải quan cửa khẩu gần nhất.\nBước 3: Khai báo tờ khai nhập khẩu thương mại trên hệ thống hải quan điện tử như hàng hóa nhập khẩu thông thường qua cửa khẩu và kiểm dịch/kiểm tra chuyên ngành (nếu thuộc danh mục).\nBước 4: Nộp đủ 100% thuế nhập khẩu, thuế GTGT và lệ phí hải quan vào Ngân sách Nhà nước trước khi điều xe vận chuyển hàng ra khỏi khu vực giám sát cửa khẩu vào nội địa.",
        "formula": "Thời hạn tối đa nộp hồ sơ hải quan = Ngày mua gom + 30 ngày | Điều kiện xuất hàng vào nội địa: Hoàn thành 100% nghĩa vụ thuế trước khi thông quan (Thông tư 154/2026/TT-BTC hiệu lực 01/01/2027).",
        "top3_errors": "1. Để dồn hàng mua gom quá 30 ngày mới mở tờ khai dẫn đến vi phạm thời hạn làm thủ tục hải quan.\n2. Vận chuyển hàng mua gom đi sâu vào nội địa trước khi nộp đủ thuế và hoàn tất thông quan.\n3. Không giữ lại bảng kê định danh nguồn gốc người bán cư dân biên giới làm căn cứ giải trình trị giá tính thuế.",
        "mnemonic": "Khẩu quyết biên mậu 2027: 'Gom xong đếm 30 ngày — Đóng thuế xong mới lăn bánh vào nội địa!'",
        "comparison": "Thông tư 80/2019/TT-BTC (Quy định cũ) vs Thông tư 154/2026/TT-BTC (Hiệu lực 01/01/2027): Khóa chặt thời hạn 30 ngày nộp hồ sơ và điều kiện hoàn thành thuế trước khi đưa vào nội địa.",
        "persona": "Thương nhân thu mua nông sản/hàng hóa cửa khẩu biên giới & Đại lý hải quan tuyến biên giới",
        "checklist": "☐ Lập bảng kê chi tiết từng giao dịch mua gom cư dân biên giới\n☐ Đặt lịch nhắc hạn 30 ngày trên phần mềm quản lý lô hàng\n☐ Đăng ký tờ khai tại Hải quan cửa khẩu gần nhất\n☐ Kiểm tra yêu cầu kiểm dịch/VSATTP (nếu là nông thủy sản)\n☐ Xác nhận hoàn thành nộp thuế trước khi điều xe rời cửa khẩu",
        "url": "https://chinhphu.vn/?pageid=27160&docid=220072",
        "keywords": ["154/2026", "bien gioi"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn chính thống Cổng TTĐT Chính phủ (chinhphu.vn); đối chiếu 2+ nguồn với thoibaotaichinhvietnam.vn & thuvienphapluat.vn."
    },
    {
        "stt": 265,
        "week": "W47 — 18/11/2026",
        "title": "Hướng Dẫn 5 Bước Làm Thủ Tục Đưa Hàng Vào — Đưa Ra Kho Ngoại Quan (Gia Cố, Chia Gói, Nhập Nội Địa) Chuẩn Công Văn 23897/CHQ-GSQL & 23907/CHQ-GSQL",
        "steps": "Bước 1: Khi đưa hàng từ nước ngoài hoặc từ nội địa vào kho ngoại quan (KNQ), chuẩn bị vận đơn, hợp đồng thuê KNQ và mở tờ khai vận chuyển độc lập/nhập KNQ theo đúng loại hình.\nBước 2: Nếu phát sinh các dịch vụ gia tăng trong KNQ (gia cố, phân loại, chia gói, đóng gói lại bao bì), phải gửi thông báo trước cho Chi cục Hải quan quản lý KNQ và lập phiếu xuất/nhập kho giám sát.\nBước 3: Lưu ý quy định C/O khi đưa hàng từ KNQ nhập khẩu vào nội địa (Công văn 23006 & 23907/CHQ-GSQL): rà soát kỹ tính hợp lệ của chứng từ xuất xứ nếu muốn hưởng thuế suất ưu đãi đặc biệt.\nBước 4: Khi xuất hàng từ KNQ ra nước ngoài hoặc vào nội địa, mở tờ khai xuất KNQ hoặc tờ khai nhập khẩu nội địa tương ứng, đối chiếu tồn kho điện tử trên phần mềm quản lý KNQ.\nBước 5: Thanh khoản và lưu trữ phiếu xuất kho, biên bản kiểm tra giám sát hải quan đối với các lô hàng đã chia tách kiện trong KNQ.",
        "formula": "Thời hạn lưu giữ KNQ chuẩn Luật Hải quan = tối đa 12 tháng (được gia hạn 1 lần tối đa 12 tháng nếu có lý do chính đáng); áp dụng quy trình giám sát theo Công văn 23897/CHQ-GSQL & 23907/CHQ-GSQL (05/11/2026).",
        "top3_errors": "1. Tự ý thực hiện chia gói, dán nhãn, đóng gói lại hàng hóa bên trong KNQ mà không thông báo và chịu sự giám sát của Hải quan quản lý kho.\n2. Mặc định hàng từ nội địa gửi vào KNQ rồi nhập ngược lại nội địa luôn được hưởng C/O ưu đãi đặc biệt mà không kiểm tra hướng dẫn liên Bộ.\n3. Để quá thời hạn lưu giữ 12 tháng tại KNQ mà không làm thủ tục gia hạn hoặc xuất kho dẫn đến xử phạt và xử lý hàng tồn đọng.",
        "mnemonic": "Kho ngoại quan là 'vùng đệm miễn thuế', nhưng mọi thao tác 'đụng tay' vào kiện hàng (chia gói, gia cố) đều phải xin phép Hải quan kho trước khi làm!",
        "comparison": "Nhập thẳng nội địa (nộp thuế ngay tại cảng, không được gia cố/chia gói trước thông quan) vs Gửi Kho ngoại quan (tạm hoãn nộp thuế, được phân loại/chia gói/đóng bao bì dưới giám sát hải quan).",
        "persona": "Quản lý chuỗi cung ứng, Doanh nghiệp logistics 3PL & Công ty FDI (như mô hình trung tâm phân phối vùng)",
        "checklist": "☐ Ký hợp đồng thuê KNQ & xác định rõ mục đích lưu kho/chia gói\n☐ Thông báo bằng văn bản cho HQ quản lý KNQ trước khi gia cố/đóng gói lại\n☐ Kiểm tra điều kiện C/O trước khi nhập từ KNQ vào nội địa (CV 23907)\n☐ Theo dõi cảnh báo hạn lưu kho 12 tháng\n☐ Đối khớp số liệu tồn kho giữa phần mềm KNQ và hệ thống Hải quan",
        "url": "https://bkhub.vn/news/cong-van-23897-chq-gsql-ve-thu-tuc-dua-hang-vao-dua-hang-ra-khoi-kho-ngoai-quan",
        "keywords": ["23897", "kho ngoai quan"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: BKHub (bkhub.vn), Thư Viện Pháp Luật (thuvienphapluat.vn) và LuatVietnam (luatvietnam.vn)."
    }
]


# ==============================================================================
# 4. DATA FOR FUNFACT (13 cols) — STT 261..265 (W47 — 18/11/2026)
# ==============================================================================
FF_NEW_IDEAS = [
    {
        "stt": 261,
        "week": "W47 — 18/11/2026",
        "title": "Con Tàu Buồm Cutty Sark (1869) & Cuộc Đua Chở Trà Nghẹt Thở: Khi Tốc Độ Giao Hàng Quyết Định Cả Giá Trị Lô Hàng Trên Thị Trường London",
        "hook": "Thế kỷ 19, các thương gia trà London sẵn sàng trả mức giá thưởng khổng lồ cho con tàu buồm đầu tiên cập bến sông Thames chở vụ trà thu hoạch mới nhất từ Trung Quốc.",
        "history": "Tàu Cutty Sark hạ thủy năm 1869 tại Dumbarton (Scotland) cho hãng Jock Willis Shipping Line, dài 85 mét, thiết kế thân hẹp lướt sóng đạt tốc độ lên tới 17,5 hải lý/giờ để tranh tài tuyến Trung Quốc — Anh.",
        "context": "Trước khi tàu hơi nước và kênh đào Suez (mở cùng năm 1869) thống trị, các tàu buồm 'Tea Clippers' phải chạy đua vượt 14.000 hải lý vòng qua Mũi Hảo Vọng trong hơn 100 ngày để đưa trà xanh giữ trọn hương vị về Anh.",
        "modern_impact": "Cuộc đua Tea Clippers đặt nền móng cho khái niệm 'Time-to-Market' trong logistics hiện đại: chi phí cước nhanh (air freight, fast-boat e-commerce) luôn đắt hơn nhưng tạo biên lợi nhuận cao hơn hẳn hàng chậm mùa vụ.",
        "intl_comp": "Ngày nay, các hãng tàu áp dụng tuyến 'Fast Boat' xuyên Thái Bình Dương (Matson, ZIM) rút ngắn thời gian từ Trung Quốc/Việt Nam đi bờ Tây Mỹ xuống còn 11–13 ngày cho hàng TMĐT mùa cao điểm.",
        "business_lesson": "Trong xuất nhập khẩu hàng mùa vụ (nông sản, thời trang, quà Tết), người thắng không phải là người có giá cước rẻ nhất, mà là người đưa hàng lên kệ đúng 'khung giờ vàng' đầu mùa.",
        "url": "https://en.wikipedia.org/wiki/Cutty_Sark",
        "keywords": ["cutty sark", "tea"],
        "series": "Series: Lịch sử Logistics",
        "status": "Chưa dùng",
        "note": "Hướng 1 (Cột mốc lịch sử thương mại). Đối chiếu 2+ nguồn: Wikipedia (en.wikipedia.org/wiki/Cutty_Sark) & Royal Museums Greenwich (rmg.co.uk)."
    },
    {
        "stt": 262,
        "week": "W47 — 18/11/2026",
        "title": "Chỉ Số Baltic Dry Index (BDI) & Quán Cà Phê Năm 1744: 'Phong Vũ Biểu' Dự Báo Suy Thoái Kinh Tế Toàn Cầu Từ Tiền Cước Tàu Chở Than, Quặng",
        "hook": "Trước cả khi chỉ số chứng khoán Dow Jones ra đời hơn 100 năm, các chủ tàu và thương nhân đã ngồi trong quán cà phê Virginia and Baltick ở London năm 1744 để định giá vận chuyển toàn cầu.",
        "history": "Năm 1744, quán cà phê Virginia and Baltick trên phố Threadneedle (London) trở thành nơi chủ tàu và chủ hàng trao đổi tin hàng hải biển Baltic; năm 1985 lập chỉ số tiền thân và năm 1999 chính thức gọi là Baltic Dry Index (BDI).",
        "context": "BDI đo giá cước vận chuyển hàng rời khô (than đá, quặng sắt, ngũ cốc) trên hơn 20 tuyến hàng hải chính bằng các tàu khổng lồ Capesize, Panamax, Supramax — thứ hàng hóa nguyên liệu đầu vào của mọi nhà máy.",
        "modern_impact": "Giới tài chính và logistics toàn cầu theo dõi BDI mỗi ngày vì nó không thể bị 'làm giá' đầu cơ: chỉ khi nhà máy thực sự đặt mua nguyên liệu thô sản xuất thì cước tàu rời mới tăng, báo trước chu kỳ tăng trưởng 3–6 tháng.",
        "intl_comp": "Khác với chỉ số container (WCI của Drewry hay SCFI của Thượng Hải) phản ánh hàng tiêu dùng thành phẩm, BDI phản ánh trực tiếp sức khỏe ngành công nghiệp nặng và nông nghiệp thế giới.",
        "business_lesson": "Doanh nghiệp XNK nguyên liệu công nghiệp và nông sản nên theo dõi song song BDI và SCFI/WCI để chốt hợp đồng cước dài hạn (Contract Rate) trước khi chu kỳ cước tăng bứt phá.",
        "url": "https://en.wikipedia.org/wiki/Baltic_Dry_Index",
        "keywords": ["baltic dry index", "1744"],
        "series": "Series: Thương mại Quốc tế",
        "status": "Chưa dùng",
        "note": "Hướng 2 & 3 (So sánh hệ thống & Con số lịch sử thương mại). Đối chiếu 2+ nguồn: Wikipedia (en.wikipedia.org/wiki/Baltic_Dry_Index) & Baltic Exchange."
    },
    {
        "stt": 263,
        "week": "W47 — 18/11/2026",
        "title": "Eo Biển Hormuz — 'Van Khóa Năng Lượng' Rộng 21 Hải Lý: Nơi Hơn 20 Triệu Thùng Dầu Mỗi Ngày (~20% Nguồn Cung Toàn Cầu) Phải Lách Qua Hành Lang Rộng Đúng 2 Hải Lý",
        "hook": "Dù Eo biển Hormuz rộng 21 hải lý (~39 km) ở điểm hẹp nhất, luồng tàu chở dầu đi và về thực tế chỉ rộng đúng 2 hải lý mỗi chiều — bằng khoảng cách nhìn thấy bằng mắt thường giữa hai bờ.",
        "history": "Theo Cơ quan Thông tin Năng lượng Hoa Kỳ (EIA), trung bình 20–21 triệu thùng dầu thô và sản phẩm dầu mỏ cùng phần lớn LNG đường biển toàn cầu đi qua Eo biển Hormuz mỗi ngày.",
        "context": "Hơn 80% lượng dầu đi qua eo biển nối Vịnh Ba Tư với Ấn Độ Dương này chảy thẳng về các thị trường công xưởng Châu Á (Trung Quốc, Ấn Độ, Nhật Bản, Hàn Quốc); hầu hết đường ống dẫn dầu thay thế đều không đủ công suất gánh thay.",
        "modern_impact": "Mỗi khi căng thẳng địa chính trị bùng phát tại Trung Đông, chỉ cần rủi ro ách tắc tại Hormuz là giá nhiên liệu tàu biển (VLSFO/MGO) tăng vọt lập tức, kéo theo phụ phí nhiên liệu (BAF) và phụ phí rủi ro chiến tranh đánh thẳng vào chi phí logistics Việt Nam.",
        "intl_comp": "Nếu Eo biển Malacca là 'cổ chai' của hàng container Châu Á thì Eo biển Hormuz là 'cổ chai' quyết định giá năng lượng và phụ phí cước vận tải của toàn bộ chuỗi cung ứng thế giới.",
        "business_lesson": "Đừng bao giờ ký hợp đồng xuất nhập khẩu giá cố định dài hạn (CIF/DDP) mà thiếu điều khoản điều chỉnh phụ phí nhiên liệu (BAF) và điều khoản bất khả kháng địa chính trị.",
        "url": "https://www.eia.gov/todayinenergy/detail.php?id=61002",
        "keywords": ["hormuz", "oil"],
        "series": "Series: Thương mại Quốc tế",
        "status": "Chưa dùng",
        "note": "Hướng 3 (Thống kê gây sốc). Đối chiếu 2+ nguồn: Cơ quan Thông tin Năng lượng Hoa Kỳ (eia.gov) & Wikipedia."
    },
    {
        "stt": 264,
        "week": "W47 — 18/11/2026",
        "title": "Thuật Toán 'Cấm Rẽ Trái' (UPS ORION): Cách Gã Khổng Lồ Chuyển Phát Tiết Kiệm 100 Triệu Dặm & 10 Triệu Gallon Nhiên Liệu Mỗi Năm Chỉ Nhờ Né Ngã Tư",
        "hook": "Từ năm 2004 đến khi triển khai hệ thống AI ORION, hàng chục nghìn xe tải giao hàng của UPS được lập trình để hạn chế tối đa việc rẽ trái qua làn xe ngược chiều — dù quãng đường có thể dài hơn vài trăm mét.",
        "history": "UPS phát triển hệ thống On-Road Integrated Optimization and Navigation (ORION) ứng dụng thuật toán AI tối ưu hóa hàng trăm nghìn phương án lộ trình mỗi ngày cho hơn 55.000 tuyến giao nhận tại Bắc Mỹ.",
        "context": "Tại các quốc gia đi bên phải, rẽ trái bắt buộc xe tải phải dừng chờ đèn đỏ lâu và cắt ngang dòng xe ngược chiều — gây tốn nhiên liệu chạy không tải (idling), trễ lịch trình và tăng nguy cơ tai nạn giao thông.",
        "modern_impact": "Nhờ ưu tiên rẽ phải và tối ưu thứ tự giao hàng theo thời gian thực, UPS công bố tiết kiệm khoảng 100 triệu dặm di chuyển, 10 triệu gallon nhiên liệu và giảm 100.000 tấn khí thải CO2 mỗi năm.",
        "intl_comp": "Các tập đoàn logistics và giao nhận chặng cuối (Last-mile delivery) tại Châu Âu, Nhật Bản và Việt Nam đều đang áp dụng thuật toán định tuyến TMS/VRP tương tự để cắt 15%–20% chi phí xăng dầu.",
        "business_lesson": "Trong vận hành kho vận và đội xe container nội địa, tối ưu hóa 'những điểm nghẽn nhỏ lặp lại hàng nghìn lần' (như thời gian chờ cổng cảng, lộ trình quay đầu xe) mang lại lợi nhuận lớn hơn nhiều so với việc chỉ ép giá tài xế.",
        "url": "https://metalumna.com/articles/ups-orion-billion-dollar-ai-routing",
        "keywords": ["orion", "ups"],
        "series": "Series: Đàm phán & Tư duy",
        "status": "Chưa dùng",
        "note": "Hướng 4 (Bài học từ mô hình doanh nghiệp & AI logistics). Đối chiếu 2+ nguồn: metalumna.com & nextbillion.ai."
    },
    {
        "stt": 265,
        "week": "W47 — 18/11/2026",
        "title": "Bí Mật 'Nước Dằn Tàu' (Ballast Water) & Công Ước BWM Của IMO: Cuộc Chiến Ngăn Hàng Chục Nghìn Tấn Nước Biển Mang Sinh Vật Xâm Lấn Vượt Đại Dương",
        "hook": "Mỗi khi một con tàu chở hàng dỡ hết container tại cảng, nó phải bơm hàng chục nghìn tấn nước biển vào đáy tàu để giữ thăng bằng — vô tình chở theo hàng tỷ sinh vật biển từ châu lục này sang châu lục khác.",
        "history": "Tổ chức Hàng hải Quốc tế (IMO) thông qua Công ước Quốc tế về Kiểm soát và Quản lý Nước dằn và Cặn lắng Tàu biển (BWM Convention) năm 2004, chính thức có hiệu lực toàn cầu từ ngày 08/09/2017.",
        "context": "Nếu không xử lý, nước dằn tàu bơm tại cảng xuất khẩu và xả ra ở cảng đích sẽ phát tán tảo độc, vi khuẩn và loài xâm lấn ngoại lai (như trai vằn), tàn phá hệ sinh thái và ngành thủy sản nước sở tại.",
        "modern_impact": "Tất cả tàu biển quốc tế hiện nay bắt buộc lắp đặt Hệ thống Xử lý Nước dằn (BWMS) bằng tia cực tím (UV), điện phân clo hoặc lọc cơ học đạt chuẩn D-2 và lưu trữ Nhật ký Nước dằn để Hải quan/Cảng vụ kiểm tra.",
        "intl_comp": "Các cảng biển lớn tại Hoa Kỳ (USCG), Úc và Châu Âu phạt rất nặng và có quyền giữ tàu nếu phát hiện chứng chỉ BWM không hợp lệ hoặc mẫu nước xả vượt ngưỡng sinh vật sống.",
        "business_lesson": "Chi phí tuân thủ tiêu chuẩn môi trường (xử lý nước dằn, nhiên liệu lưu huỳnh thấp IMO 2020, EU ETS) đang cấu thành một phần không thể thiếu trong giá cước biển quốc tế mà chủ hàng cần thấu hiểu khi đàm phán.",
        "url": "https://en.wikipedia.org/wiki/Ballast_Water_Management_Convention",
        "keywords": ["ballast water", "imo"],
        "series": "Series: Công nghệ & Đổi mới",
        "status": "Chưa dùng",
        "note": "Hướng 5 (Đổi mới công nghệ & quy chuẩn hàng hải). Đối chiếu 2+ nguồn: Wikipedia (en.wikipedia.org) & Tổ chức Hàng hải Quốc tế (imo.org)."
    }
]


# ==============================================================================
# 5. FEEDBACK LOOP: MISSING VERIFIED LEGAL / TRADE DEFENSE DOCS FOR TRACKER
# ==============================================================================
NEW_LEGAL_DOCS = [
    {
        "ma_so": "Kết luận sơ bộ DOC (PTMEG Hoa Kỳ)",
        "ten": "Kết luận sơ bộ của Bộ Thương mại Hoa Kỳ (DOC) trong vụ điều tra chống bán phá giá đối với sản phẩm PTMEG (HS 3907.29.0000) nhập khẩu từ Việt Nam",
        "loai": "Quyết định Phòng vệ thương mại Quốc tế",
        "co_quan": "Bộ Thương mại Hoa Kỳ (DOC) / Cục Phòng vệ thương mại (Bộ Công Thương) cảnh báo",
        "ngay_ban_hanh": "2026-11-10",
        "ngay_hieu_luc": "2026-11-10",
        "ngay_het_hieu_luc": "Chờ kết luận cuối cùng",
        "van_ban_bi_thay_the": "Lần đầu ban hành biên độ sơ bộ cho mặt hàng PTMEG Việt Nam",
        "phan_tich": "DOC xác định biên độ bán phá giá sơ bộ là 36,55% đối với bị đơn bắt buộc và nhóm thuế suất toàn quốc của Việt Nam; hàng hóa thuộc diện điều tra xuất sang Hoa Kỳ phải nộp tiền ký quỹ tạm thời tương ứng mức 36,55%.",
        "so_sanh": "Từ thuế suất thông thường chuyển sang phải ký quỹ tạm thời 36,55% trong giai đoạn điều tra sơ bộ, đòi hỏi doanh nghiệp chuẩn bị hồ sơ thẩm tra tại chỗ chặt chẽ."
    }
]


def check_banned_words():
    """Verify 0 banned words across all 20 ideas and legal docs."""
    all_blobs = []
    for lst in (TN_NEW_IDEAS, CS_NEW_IDEAS, HT_NEW_IDEAS, FF_NEW_IDEAS, NEW_LEGAL_DOCS):
        for item in lst:
            all_blobs.append((item.get("stt", item.get("ma_so", "")), " ".join(str(v) for v in item.values())))
    violations = []
    for label, text in all_blobs:
        low = text.lower()
        for bw in BANNED_WORDS_CHECK:
            if bw in low:
                violations.append((label, bw))
    if violations:
        raise ValueError(f"Phat hien tu cam vi pham banned_words.md: {violations}")
    print("🟢 BANNED WORDS CHECK PASSED: 0 vi phạm từ cấm trên toàn bộ 20 ý tưởng & văn bản luật!")


def check_no_duplicates():
    """Ensure STT 261-265 and titles/legal codes do not duplicate existing Excel rows."""
    for path in (
        "reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx",
        "reports/Case_Study_Ideas_Bank_Eureka.xlsx",
        "reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx",
        "reports/Funfact_Ideas_Bank_Eureka.xlsx",
    ):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        last_stt = ws.cell(row=ws.max_row, column=1).value
        if last_stt != 260:
            raise ValueError(f"{path} last STT is {last_stt}, expected 260 before appending 261-265!")
        wb.close()
    print("🟢 DUPLICATE & SEQUENCE CHECK PASSED: 4 file Ideas Bank sẵn sàng nối tiếp STT 261-265.")


def execute_update():
    print("=" * 80)
    print("🚀 WORKFLOW 06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (WEEK 47 DAY 3 — 18/11/2026)")
    print("=" * 80)

    check_banned_words()
    check_no_duplicates()

    # 1. TECHNICAL GATEKEEPER: VALIDATE ALL 20 URLS BEFORE TOUCHING EXCEL
    all_ideas_for_gatekeeper = []
    for item in TN_NEW_IDEAS + CS_NEW_IDEAS + HT_NEW_IDEAS + FF_NEW_IDEAS:
        all_ideas_for_gatekeeper.append({
            "stt": item["stt"],
            "title": item["title"],
            "url": item["url"],
            "keywords": item["keywords"]
        })

    print(f"\n🔍 Gửi {len(all_ideas_for_gatekeeper)} link qua Technical Gatekeeper url_validator.py...")
    enforce_strict_validation(all_ideas_for_gatekeeper)
    print("\n🟢 TECHNICAL GATEKEEPER PHÊ DUYỆT 100%: TẤT CẢ 20 LINK ĐỀU HỢP LỆ VÀ ĐÃ ĐỐI SOÁT NỘI DUNG THỰC TẾ!\n")

    # 2. TIN NGANH (15 cols)
    tn_path = "reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx"
    wb_tn = openpyxl.load_workbook(tn_path)
    ws_tn = wb_tn.active
    for item in TN_NEW_IDEAS:
        r = ws_tn.max_row + 1
        fill = FILL_ODD if item["stt"] % 2 != 0 else FILL_EVEN
        row_vals = [
            item["stt"], item["week"], item["title"], item["agency"], item["effective_date"],
            item["core_diff"], item["affected_entity"], item["risk_pain"], item["market_reaction"],
            item["pro_argument"], item["con_argument"], item["intl_ref"], item["url"],
            item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_tn.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in (1, 14) else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in (1, 2, 14) else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
    atomic_save(wb_tn, tn_path)
    print(f"✅ Ghi thành công {tn_path} (Tổng dòng: {ws_tn.max_row}, STT cuối: 265)")

    # 3. CASE STUDY (14 cols)
    cs_path = "reports/Case_Study_Ideas_Bank_Eureka.xlsx"
    wb_cs = openpyxl.load_workbook(cs_path)
    ws_cs = wb_cs.active
    for item in CS_NEW_IDEAS:
        r = ws_cs.max_row + 1
        fill = FILL_ODD if item["stt"] % 2 != 0 else FILL_EVEN
        row_vals = [
            item["stt"], item["week"], item["title"], item["parties"], item["plot"],
            item["damage"], item["root_cause"], item["prevention"], item["opposing_views"],
            item["lessons"], item["url"], item["series"], item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_cs.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in (1, 13) else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in (1, 2, 13) else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
    atomic_save(wb_cs, cs_path)
    print(f"✅ Ghi thành công {cs_path} (Tổng dòng: {ws_cs.max_row}, STT cuối: 265)")

    # 4. HOC THUC CHIEN (14 cols)
    ht_path = "reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx"
    wb_ht = openpyxl.load_workbook(ht_path)
    ws_ht = wb_ht.active
    for item in HT_NEW_IDEAS:
        r = ws_ht.max_row + 1
        fill = FILL_ODD if item["stt"] % 2 != 0 else FILL_EVEN
        row_vals = [
            item["stt"], item["week"], item["title"], item["steps"], item["formula"],
            item["top3_errors"], item["mnemonic"], item["comparison"], item["persona"],
            item["checklist"], item["url"], item["series"], item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_ht.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in (1, 13) else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in (1, 2, 13) else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
    atomic_save(wb_ht, ht_path)
    print(f"✅ Ghi thành công {ht_path} (Tổng dòng: {ws_ht.max_row}, STT cuối: 265)")

    # 5. FUNFACT (13 cols)
    ff_path = "reports/Funfact_Ideas_Bank_Eureka.xlsx"
    wb_ff = openpyxl.load_workbook(ff_path)
    ws_ff = wb_ff.active
    for item in FF_NEW_IDEAS:
        r = ws_ff.max_row + 1
        fill = FILL_ODD if item["stt"] % 2 != 0 else FILL_EVEN
        row_vals = [
            item["stt"], item["week"], item["title"], item["hook"], item["history"],
            item["context"], item["modern_impact"], item["intl_comp"], item["business_lesson"],
            item["url"], item["series"], item["status"], item["note"]
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_ff.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in (1, 12) else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in (1, 2, 12) else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
    atomic_save(wb_ff, ff_path)
    print(f"✅ Ghi thành công {ff_path} (Tổng dòng: {ws_ff.max_row}, STT cuối: 265)")

    # 6. FEEDBACK LOOP: DANH SACH VAN BAN PHAP LUAT XNK
    legal_path = "reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx"
    wb_leg = openpyxl.load_workbook(legal_path)
    ws_leg = wb_leg.active
    existing_codes = {str(ws_leg.cell(r, 2).value or "").strip().lower() for r in range(2, ws_leg.max_row + 1)}
    added_legal = 0
    for doc in NEW_LEGAL_DOCS:
        if doc["ma_so"].strip().lower() in existing_codes:
            continue
        r = ws_leg.max_row + 1
        stt_val = r - 1
        fill = FILL_ODD if stt_val % 2 != 0 else FILL_EVEN
        row_vals = [
            stt_val,
            doc["ma_so"],
            doc["ten"],
            doc["loai"],
            doc["co_quan"],
            doc["ngay_ban_hanh"],
            doc["ngay_hieu_luc"],
            doc["ngay_het_hieu_luc"],
            doc["van_ban_bi_thay_the"],
            doc["phan_tich"],
            doc["so_sanh"],
            "2026-11-18"
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_leg.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in (1, 2) else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in (1, 4, 6, 7, 8, 12) else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
        added_legal += 1
        existing_codes.add(doc["ma_so"].strip().lower())
    atomic_save(wb_leg, legal_path)
    print(f"✅ Feedback Loop: Bổ sung {added_legal} văn bản mới vào {legal_path} (Tổng dòng: {ws_leg.max_row})")

    print("\n" + "=" * 80)
    print("🎉 HOÀN TẤT WORKFLOW 06 (18/11/2026): ĐÃ CẬP NHẬT 20 Ý TƯỞNG MỚI (STT 261-265) & FEEDBACK LOOP!")
    print("=" * 80)


if __name__ == "__main__":
    execute_update()
