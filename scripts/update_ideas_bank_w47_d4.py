# -*- coding: utf-8 -*-
"""
Execute Workflow 06: Ideas Bank Daily Update (wf06) - Week 47 Day 4 (19/11/2026)
Appends 5 new verified ideas (STT 266-270) to each of the 4 Pillars (TN, CS, HT, FF).
Runs through Technical Gatekeeper url_validator.py and Banned Words check before saving.
Updates Danh_sach_Van_ban_Phap_luat_XNK.xlsx (Feedback Loop) for missing verified regulations.
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
# 1. DATA FOR TIN NGANH (15 cols) — STT 266..270 (W47 — 19/11/2026)
# ==============================================================================
TN_NEW_IDEAS = [
    {
        "stt": 266,
        "week": "W47 — 19/11/2026",
        "title": "Quyết Định Số 3318/QĐ-BTC (Ban Hành 12/11/2026): Bộ Tài Chính Công Bố 133 Thủ Tục Hải Quan Thực Hiện 'Phi Địa Giới Hành Chính' Trên Toàn Quốc",
        "agency": "Bộ Tài chính",
        "effective_date": "Ban hành và triển khai từ ngày 12/11/2026 (trong tổng số 255 TTHC thuộc 6 lĩnh vực quản lý của Bộ Tài chính)",
        "core_diff": "Cũ: Doanh nghiệp xuất nhập khẩu thường bị ràng buộc nộp hồ sơ tại đúng cơ quan hải quan quản lý trực tiếp địa bàn đặt trụ sở hoặc cửa khẩu nơi lưu giữ hàng hóa; Mới: Quyết định 3318/QĐ-BTC công bố 133 thủ tục hành chính lĩnh vực hải quan được tiếp nhận, giải quyết không phụ thuộc địa giới hành chính trên phạm vi cả nước (gồm xem hàng trước khi khai, hủy tờ khai, đăng ký tờ khai một lần, thủ tục đưa hàng vào/ra khu phi thuế quan và hoàn thuế, miễn thuế).",
        "affected_entity": "Các doanh nghiệp XNK có nhiều nhà máy, chi nhánh tại nhiều tỉnh thành, đại lý làm thủ tục hải quan, doanh nghiệp trong khu phi thuế quan và phòng chứng từ XNK.",
        "risk_pain": "Nếu bộ phận chứng từ không cập nhật danh mục 133 TTHC phi địa giới và quy trình nộp qua Cổng dịch vụ công / bưu chính, doanh nghiệp vẫn phải cử nhân sự di chuyển xa tốn kém chi phí đi lại, chậm tiến độ hủy tờ khai hoặc xem hàng trước khi khai làm phát sinh lưu cont, lưu bãi.",
        "market_reaction": "Cộng đồng doanh nghiệp logistics và đại lý hải quan đánh giá đây là bước cải cách mạnh mẽ giúp gỡ bỏ rào cản hành chính địa phương, tối ưu nhân sự khai báo tập trung tại một đầu mối.",
        "pro_argument": "Giảm đáng kể thời gian, chi phí tuân thủ thủ tục hành chính và thúc đẩy liên thông dữ liệu số giữa các Chi cục Hải quan khu vực trên toàn quốc.",
        "con_argument": "Giai đoạn đầu triển khai cần đồng bộ hạ tầng tiếp nhận và phân công trách nhiệm phối hợp giữa nơi tiếp nhận hồ sơ với chi cục nơi kiểm tra thực tế hàng hóa để tránh đùn đẩy xử lý.",
        "intl_ref": "Mô hình Hải quan Một cửa Quốc gia & Thông quan tập trung (Centralized Customs Clearance) đang áp dụng tại Liên minh Châu Âu (UCC) và Singapore TradeNet.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/133-thu-tuc-hai-quan-duoc-thuc-hien-khong-phu-thuoc-dia-gioi-hanh-chinh-172492.html",
        "keywords": ["3318", "133"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn), Thời báo Tài chính Việt Nam (thoibaotaichinhvietnam.vn), LuatVietnam (luatvietnam.vn) và Thư Viện Pháp Luật."
    },
    {
        "stt": 267,
        "week": "W47 — 19/11/2026",
        "title": "Quy Định Thực Thi (EU) 2026/2547 (Ban Hành 13/11/2026, Hiệu Lực 17/11/2026): Ủy Ban Châu Âu (EC) Chính Thức Áp Thuế Chống Bán Phá Giá 16% Đối Với Thép Cán Nguội Dạng Phẳng Việt Nam",
        "agency": "Ủy ban Châu Âu (EC) & Cục Phòng vệ thương mại (Bộ Công Thương)",
        "effective_date": "Ban hành ngày 13/11/2026, công bố Công báo EU ngày 16/11/2026, chính thức có hiệu lực từ ngày 17/11/2026",
        "core_diff": "Cũ: Thép cán nguội dạng phẳng (Cold-rolled flat steel, nhóm mã HS 7209, 7211, 7225, 7226) xuất khẩu từ Việt Nam sang EU chỉ chịu hạn ngạch tự vệ thép thông thường; Mới: EC chính thức áp mức thuế chống bán phá giá 16,0% chung cho các doanh nghiệp Việt Nam (thấp hơn Nhật Bản 27,9% và Đài Loan 20,7–27%, nhưng cao hơn Ấn Độ 9,5% và Thổ Nhĩ Kỳ 3–7,8%), đồng thời bác bỏ cáo buộc 'tình hình thị trường đặc biệt' (PMS) đối với Việt Nam.",
        "affected_entity": "Các doanh nghiệp sản xuất, gia công và xuất khẩu thép cán nguội dạng phẳng (trừ thép không gỉ, thép điện và thép lá đen dưới 0,35mm) sang thị trường 27 nước EU.",
        "risk_pain": "Biên độ thuế 16% trực tiếp bào mòn lợi thế giá thép Việt Nam trước đối thủ Ấn Độ và Thổ Nhĩ Kỳ; các lô hàng đang trên biển hoặc ký hợp đồng giao hàng DDP/CIF chưa dự phòng điều khoản thuế phòng vệ thương mại đối mặt nguy cơ thua lỗ nặng.",
        "market_reaction": "Cục Phòng vệ thương mại (Bộ Công Thương) và Thương vụ Việt Nam tại EU khuyến nghị doanh nghiệp rà soát kỹ danh mục mã HS loại trừ và chuẩn hóa dữ liệu truy xuất 'Melt & Pour' sang EU.",
        "pro_argument": "Việc EC bác bỏ cáo buộc PMS là tiền đề pháp lý tích cực giúp Việt Nam tránh bị áp biên độ thuế cực đoan theo giá trị thay thế bất lợi, giữ được mức thuế thấp hơn Nhật Bản và Đài Loan.",
        "con_argument": "Áp lực thuế 16% cộng hưởng với cơ chế CBAM và quy tắc 'nấu chảy và đúc' (Melt & Pour) khiến cánh cửa xuất khẩu thép cán nguội sang Châu Âu ngày càng thu hẹp biên lợi nhuận.",
        "intl_ref": "Quy định thực thi (EU) 2026/2547 trong cuộc điều tra song song với Ấn Độ, Nhật Bản, Đài Loan và Thổ Nhĩ Kỳ khởi xướng từ tháng 09/2025.",
        "url": "https://vietstock.vn/2026/11/eu-ap-thue-chong-ban-pha-gia-16-voi-thep-can-nguoi-viet-nam-742-1515476.htm",
        "keywords": ["2026/2547", "16%"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Vietstock (vietstock.vn), Trung tâm WTO VCCI (chongbanphagia.vn), Công Thương (congthuong.vn) và VnEconomy."
    },
    {
        "stt": 268,
        "week": "W47 — 19/11/2026",
        "title": "Công Văn Số 24313/CHQ-GSQL (Ban Hành 16/11/2026): Cục Hải Quan Triển Khai Hệ Thống Thủ Tục Hải Quan Chuyên Biệt Cho Hàng Hóa XNK Qua Thương Mại Điện Tử",
        "agency": "Cục Hải quan (Ban Giám sát quản lý về hải quan)",
        "effective_date": "Ban hành ngày 16/11/2026; tổ chức phổ biến kế hoạch triển khai trực tiếp vào ngày 19/11/2026",
        "core_diff": "Cũ: Hàng hóa giao dịch qua sàn thương mại điện tử xuyên biên giới chủ yếu khai báo ghép theo luồng chuyển phát nhanh/bưu chính thông thường trên hệ thống VNACCS cũ; Mới: Cục Hải quan triển khai Hệ thống xử lý dữ liệu điện tử hải quan chuyên biệt dành riêng cho hàng hóa XNK giao dịch qua thương mại điện tử để đồng bộ chuẩn dữ liệu đơn hàng, giá trị giao dịch và quản lý rủi ro từ nền tảng sàn.",
        "affected_entity": "Các doanh nghiệp chuyển phát nhanh, kho ngoại quan phục vụ TMĐT, sàn thương mại điện tử xuyên biên giới và các chủ shop nhập khẩu tần suất cao.",
        "risk_pain": "Đơn vị vận hành và chủ hàng không kịp kết nối chuẩn dữ liệu đơn hàng TMĐT theo hệ thống mới sẽ bị kẹt luồng thông quan tự động tại các cửa khẩu sân bay Nội Bài, Tân Sơn Nhất và cửa khẩu đường bộ phía Bắc.",
        "market_reaction": "Các doanh nghiệp chuyển phát nhanh và đại lý hải quan khẩn trương tham gia tập huấn và nâng cấp API kết nối hệ thống khai báo hải quan TMĐT riêng biệt.",
        "pro_argument": "Tạo luồng xử lý tách biệt giúp thông quan siêu tốc hàng triệu bưu kiện TMĐT hợp pháp mỗi ngày, đồng thời ngăn chặn gian lận giá trị và xé nhỏ đơn hàng.",
        "con_argument": "Yêu cầu đầu tư đồng bộ phần mềm API kết nối giữa sàn TMĐT, đơn vị vận chuyển và hải quan tạo áp lực chi phí công nghệ cho các đơn vị logistics vừa và nhỏ.",
        "intl_ref": "Khung tiêu chuẩn Hải quan về Thương mại điện tử xuyên biên giới (WCO Framework of Standards on Cross-Border E-Commerce).",
        "url": "https://luatvietnam.vn/xuat-nhap-khau/cong-van-24313-chq-gsql-2026-cua-cuc-hai-quan-trien-khai-thu-tuc-hai-quan-qua-thuong-mai-dien-tu-456556-d6.html",
        "keywords": ["24313", "thuong mai dien tu"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: LuatVietnam (luatvietnam.vn), Thư Viện Pháp Luật (thuvienphapluat.vn) và Thời báo Tài chính Việt Nam."
    },
    {
        "stt": 269,
        "week": "W47 — 19/11/2026",
        "title": "Quyết Định Số 2927/QĐ-BCT (04/11/2026) & Triển Khai VNACCS Từ 19/11/2026: Cập Nhật Thuế Chống Bán Phá Giá 4,47% Cáp Thép Dự Ứng Lực & Siết Cấu Kiện Thép (CV 23969–23970)",
        "agency": "Bộ Công Thương & Cục Hải quan",
        "effective_date": "QĐ 2927/QĐ-BCT ban hành 04/11/2026; Cục Hải quan áp dụng biểu thuế mới trên VNACCS thống nhất từ ngày 19/11/2026",
        "core_diff": "Cũ: Thuế chống bán phá giá đối với cáp thép dự ứng lực (AR01.AD17) từ Malaysia, Thái Lan, Trung Quốc áp dụng theo mức trước rà soát và một số lô cấu kiện thép còn khai lẫn mã HS 7308/9406; Mới: Từ ngày 19/11/2026, hệ thống VNACCS cập nhật mức thuế CBPG mới (ví dụ 4,47% cho nhóm Boxing Victory đối với HS 7312.10.91, 7312.10.99), đồng thời Công văn 23969 và 23970/CHQ-NVTHQ yêu cầu kiểm tra gắt gao trị giá, C/O và mã HS cấu kiện thép/nhà tiền chế.",
        "affected_entity": "Doanh nghiệp nhập khẩu cáp thép dự ứng lực xây dựng, cấu kiện kết cấu thép (nhóm 7308), nhà thép tiền chế (nhóm 9406) và nhà thầu dự án FDI.",
        "risk_pain": "Khai sai mã thuế CBPG trên VNACCS từ mốc 19/11/2026 hoặc áp nhầm mã HS từ 7308 sang 9406 để xin miễn thuế tài sản cố định dễ bị hải quan bác miễn thuế, tham vấn giá và ấn định truy thu hàng tỷ đồng.",
        "market_reaction": "Hiệp hội Doanh nghiệp Cơ khí Việt Nam (VAMI) ủng hộ siết chặt kiểm tra cấu kiện thép nhập khẩu để bảo vệ sản xuất nội địa, trong khi các nhà thầu rà soát lại chứng từ C/O và tên nhà sản xuất miễn trừ.",
        "pro_argument": "Ngăn chặn tình trạng khai báo trị giá thấp bất thường và chuyển dịch mã HS để né thuế phòng vệ thương mại, tạo sân chơi bình đẳng cho ngành cơ khí – kết cấu thép trong nước.",
        "con_argument": "Thời gian kiểm tra thực tế bản vẽ kỹ thuật và xác minh tiêu chí 'trong nước chưa sản xuất được' kéo dài có thể ảnh hưởng tiến độ thi công các dự án công nghiệp.",
        "intl_ref": "Cơ chế rà soát định kỳ thuế chống bán phá giá (Administrative Review) theo Hiệp định Chống bán phá giá của WTO (ADA).",
        "url": "https://congthuong.vn/hai-quan-tang-kiem-soat-thep-nhap-khau-481514.html",
        "keywords": ["2927", "cap thep"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Báo Công Thương (congthuong.vn), LuatVietnam (cong-van-23970-chq-nvthq), ASL Law (aslgate.com) và Thư Viện Pháp Luật."
    },
    {
        "stt": 270,
        "week": "W47 — 19/11/2026",
        "title": "Thông Báo Số 150/2026 Của Tổng Cục Hải Quan Trung Quốc (GACC — Triển Khai Tháng 11/2026): Cắt Giảm Hồ Sơ Đăng Ký Xuất Khẩu Thực Phẩm Từ 20 Xuống 10 Mục & Duyệt Trong 3 Ngày",
        "agency": "Tổng cục Hải quan Trung Quốc (GACC) & Bộ Công Thương Việt Nam thông báo",
        "effective_date": "Hiệu lực từ ngày 23/10/2026; Cục Xuất nhập khẩu (Bộ Công Thương) & SPS Việt Nam triển khai phổ biến rộng rãi trong tháng 11/2026",
        "core_diff": "Cũ: Doanh nghiệp xuất khẩu và đại lý nước ngoài đăng ký mã GACC phải kê khai 20 mục thông tin chi tiết và chờ xét duyệt kéo dài; Mới: Thông báo 150/2026 của GACC tinh giản biểu mẫu đăng ký xuống còn 10 mục thông tin, rút ngắn thời gian xét duyệt hồ sơ hợp lệ xuống tối đa 03 ngày làm việc qua Cổng một cửa trực tuyến, nhưng bắt buộc cập nhật trong vòng 60 ngày khi có bất kỳ thay đổi thông tin nào.",
        "affected_entity": "Doanh nghiệp sản xuất, chế biến, xuất khẩu nông sản, thủy sản, thực phẩm và các đại lý ủy thác xuất khẩu chính ngạch sang thị trường Trung Quốc.",
        "risk_pain": "Doanh nghiệp thay đổi địa chỉ, người đại diện hoặc danh mục sản phẩm nhưng quên cập nhật lên hệ thống GACC trong vòng 60 ngày có nguy cơ bị tạm đình chỉ mã số xuất khẩu, khiến container nông sản tươi sống ùn ứ tại cửa khẩu.",
        "market_reaction": "Các hiệp hội nông sản và doanh nghiệp xuất khẩu sang Trung Quốc hoan nghênh việc rút ngắn thời gian duyệt còn 3 ngày làm việc trước mùa cao điểm Tết Nguyên đán 2027.",
        "pro_argument": "Giảm 50% khối lượng giấy tờ kê khai ban đầu và tăng tốc độ cấp mã xuất khẩu giúp doanh nghiệp mới tiếp cận thị trường Trung Quốc nhanh hơn.",
        "con_argument": "Trách nhiệm pháp lý tự khai tự chịu và mốc thời hạn 60 ngày cập nhật thông tin đòi hỏi doanh nghiệp phải có nhân sự chuyên trách theo dõi sát tài khoản CIFER/Một cửa GACC.",
        "intl_ref": "Cải cách quản lý đăng ký nhà xuất khẩu thực phẩm xuyên biên giới theo chuẩn số hóa của GACC (sửa đổi bổ sung cơ chế Lệnh 248/249).",
        "url": "https://congthuong.vn/trung-quoc-sua-quy-dinh-dang-ky-doanh-nghiep-xuat-khau-thuc-pham-479439.html",
        "keywords": ["GACC", "150"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Báo Công Thương (congthuong.vn), Cổng TTĐT Bộ Công Thương (moit.gov.vn) và Chi cục Trồng trọt & BVTV TP.HCM."
    }
]


# ==============================================================================
# 2. DATA FOR CASE STUDY (14 cols) — STT 266..270 (W47 — 19/11/2026)
# ==============================================================================
CS_NEW_IDEAS = [
    {
        "stt": 266,
        "week": "W47 — 19/11/2026",
        "title": "Vụ Kiện Hành Chính Hiếm Có Tại Hải Phòng (2026): Công Ty Phúc Thành Kiện Đòi Hải Quan Bồi Thường Hơn 3,3 Tỷ Đồng Vì Giữ Kéo Dài 883 Tấn Lúa Mì Dù Đã Nộp Phạt 40 Triệu",
        "parties": "Công ty TNHH XNK và Dịch vụ Phúc Thành + Chi cục Hải quan khu vực III & Hải quan cửa khẩu cảng Đình Vũ (TAND khu vực 4 TP. Hải Phòng xét xử sơ thẩm 2026)",
        "plot": "Đầu tháng 07/2025, Công ty Phúc Thành nhập khẩu 31 container gồm hơn 883 tấn lúa mì từ Kazakhstan qua cảng Hải Phòng nhưng sơ suất khai nhầm mục đích sử dụng từ nguyên liệu thức ăn chăn nuôi sang thực phẩm dùng cho người. Doanh nghiệp đã chủ động khai bổ sung, hoàn tất kiểm dịch và nộp phạt vi phạm hành chính 40 triệu đồng. Tuy nhiên, Đội Kiểm soát Hải quan vẫn yêu cầu giữ lô hàng hơn 1 tháng để xác minh dấu hiệu hình sự rồi mới giải phóng cuối tháng 08/2025 khi không có dấu hiệu tội phạm; năm 2026 doanh nghiệp khởi kiện hành chính đòi bồi thường.",
        "damage": "Chi phí lưu container, lưu bãi kéo dài hơn 1 tháng và thiệt hại hủy hợp đồng khiến doanh nghiệp khởi kiện yêu cầu bồi thường hơn 3,3 tỷ đồng (cùng 40 triệu đồng tiền phạt hành chính ban đầu cho 31 container / 883 tấn lúa mì).",
        "root_cause": "Khâu kiểm tra chứng từ đầu vào để nhầm mục đích sử dụng giữa 'thức ăn chăn nuôi' và 'thực phẩm cho người' (hai luồng quản lý kiểm tra chuyên ngành khác nhau), kích hoạt quy trình xác minh kiểm soát kéo dài của cơ quan hải quan.",
        "prevention": "1. Lập checklist đối chiếu 2 lớp mục đích sử dụng hàng hóa trên Hợp đồng, Giấy phép nhập khẩu, Đăng ký kiểm dịch thực vật/chất lượng trước khi truyền tờ khai chính thức.\n2. Nếu phát hiện sai mục đích sử dụng, nộp ngay văn bản giải trình kèm chứng từ chứng minh không cố ý gian lận để xin áp dụng khai bổ sung nhanh và nhận hàng về kho bảo quản.\n3. Khi bị giữ hàng kéo dài quá thời hạn xác minh hành chính, ghi nhận đầy đủ biên bản, hóa đơn DEM/DET và gửi kiến nghị bằng văn bản lên Cục Hải quan cấp trên để bảo vệ quyền lợi hợp pháp.",
        "opposing_views": "Phe A: Doanh nghiệp đã nộp phạt 40 triệu đồng và sửa tờ khai thì phải được giải phóng hàng ngay, giữ tiếp hơn 1 tháng là gây thiệt hại nặng vs Phe B: Khi phát hiện dấu hiệu nghi vấn chuyển đổi mục đích sử dụng số lượng lớn (883 tấn), cơ quan kiểm soát có quyền xác minh dấu hiệu hình sự theo quy trình pháp luật.",
        "lessons": "Một chữ nhầm về mục đích sử dụng trên tờ khai hải quan có thể khiến 31 container nằm cảng hơn 1 tháng và đẩy doanh nghiệp vào cuộc chiến pháp lý nhiều tỷ đồng. Phòng chứng từ phải rà soát mục đích sử dụng song song với mã HS.",
        "url": "https://phapluatplus.baophapluat.vn/tand-khu-vuc-4-hai-phong-dua-ra-xet-xu-vu-doanh-nghiep-khieu-kien-hanh-chinh-trong-linh-vuc-hai-quan.html",
        "keywords": ["883", "lua mi"],
        "series": "Series: Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Pháp luật Plus / Báo Pháp Luật Việt Nam (phapluatplus.baophapluat.vn) và VietnamFinance (vietnamfinance.vn)."
    },
    {
        "stt": 267,
        "week": "W47 — 19/11/2026",
        "title": "Xét Xử 13 Bị Cáo Đường Dây Nhập Lậu Hơn 8,65 Tỷ Đồng Qua Cửa Khẩu Bắc Luân 2 (10/2026): Thủ Đoạn Trà Trộn Hàng Vào Container Nguyên Liệu Sản Xuất Doanh Nghiệp FDI",
        "parties": "13 bị cáo trong đường dây vận chuyển trái phép từ Trung Quốc qua Cửa khẩu Bắc Luân 2 (Móng Cái, Quảng Ninh) + Hải quan & Công an tỉnh Quảng Ninh + TAND xét xử sơ thẩm 10/2026",
        "plot": "Các đối tượng móc nối với doanh nghiệp bên Trung Quốc lợi dụng chính sách ưu đãi miễn kiểm tra thực tế đối với container nguyên liệu sản xuất da giày, máy móc công nghiệp của doanh nghiệp FDI để cất giấu, xen kẹp máy móc, linh kiện, đồ điện tử không khai báo chung vào container hợp pháp đi qua cửa khẩu Bắc Luân 2, sau đó chở về kho trung chuyển tại Hải Phòng bóc tách tiêu thụ.",
        "damage": "Tổng trị giá hàng hóa nhập khẩu trái phép bị truy tố trong vụ án được xác định hơn 8,65 tỷ đồng; toàn bộ đường dây gồm 13 bị cáo phải hầu tòa về tội danh hình sự qua biên giới trong phiên xét xử tháng 10/2026.",
        "root_cause": "Lòng tham lợi dụng luồng xanh và ưu đãi kiểm tra của doanh nghiệp gia công/FDI để ghép hàng không khai báo vào container chính ngạch nhằm né thuế và kiểm tra chuyên ngành.",
        "prevention": "1. Doanh nghiệp FDI và công ty logistics tuyệt đối không cho bên thứ ba ghép hàng ngoài danh mục Packing List vào container nguyên liệu sản xuất.\n2. Lắp camera giám sát quá trình đóng hàng tại kho xuất và kiểm tra đối chiếu số seal container thực tế trước khi tàu/xe rời bãi.\n3. Thiết lập quy trình rà soát bảng kê chi tiết từng kiện hàng trước khi đại lý hải quan truyền tờ khai qua cửa khẩu đường bộ.",
        "opposing_views": "Phe A: Ghép thêm vài kiện máy móc vào container FDI luồng xanh là cách tiết kiệm chi phí vận chuyển nhanh gọn vs Phe B: Một kiện hàng không khai báo nằm chung container sẽ kéo cả chủ hàng FDI, nhà xe và người khai hải quan vào vòng lao lý hình sự.",
        "lessons": "Luồng xanh hải quan là phần thưởng cho sự tuân thủ, không phải vỏ bọc để ghép hàng. Doanh nghiệp FDI cần kiểm soát chặt khâu đóng hàng tại đầu nước ngoài để tránh liên đới oan.",
        "url": "https://congly.vn/buon-lau-qua-cua-khau-13-bi-cao-hau-toa-545035.html",
        "keywords": ["Bac Luan", "8,65"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Báo Công Lý (congly.vn), Báo Nhân Dân (nhandan.vn) và Báo Quảng Ninh."
    },
    {
        "stt": 268,
        "week": "W47 — 19/11/2026",
        "title": "Triệt Phá 3 Container Gỗ Xẻ Từ Congo Giấu 1.193 Kg Ngà Voi & 79,3 Kg Xương Thú Quý Hiếm Cập Cảng Quy Nhơn (Khởi Tố 08/2026)",
        "parties": "Công ty CP Dịch vụ Thương mại và Dịch thuật Thành Công + Chi cục Điều tra chống buôn lậu (Cục Hải quan) + Cục C03, C05 (Bộ Công an) & Công an tỉnh Gia Lai",
        "plot": "Ngày 09/07/2026, 3 container xuất phát từ Cộng hòa Congo cập cảng Quy Nhơn được doanh nghiệp khai báo là gỗ xẻ nhập khẩu. Qua máy soi chiếu container kết hợp khoan cắt kiểm tra thực tế đầu tháng 08/2026, lực lượng liên ngành phát hiện các đối tượng đã khoét rỗng lõi các thanh gỗ xẻ, nhét ngà voi và xương động vật vào bên trong rồi đổ sáp bịt kín ngụy trang như gỗ nguyên khối. Ngày 12/08/2026, cơ quan điều tra đã khởi tố vụ án hình sự.",
        "damage": "Tịch thu toàn bộ 3 container gỗ tang vật chứa 1.193 kg ngà voi (221 khúc thuộc họ Voi) và 79,3 kg xương động vật họ Mèo lớn thuộc Danh mục CITES cấm buôn bán; khởi tố vụ án hình sự xử lý nghiêm theo pháp luật.",
        "root_cause": "Tội phạm xuyên quốc gia sử dụng thủ đoạn ngụy trang công nghệ cao trong lõi gỗ xẻ, lợi dụng pháp nhân doanh nghiệp dịch vụ thương mại đứng tên mở tờ khai nhập khẩu qua cảng biển.",
        "prevention": "1. Doanh nghiệp làm dịch vụ ủy thác nhập khẩu gỗ, nông sản từ Châu Phi tuyệt đối từ chối đứng tên mở tờ khai cho các chủ hàng không minh bạch hồ sơ CITES và giấy phép khai thác hợp pháp.\n2. Yêu cầu chứng thư kiểm định độc lập (SGS/Bureau Veritas) tại cảng xếp hàng trước khi tàu chạy đối với hàng gỗ xẻ hộp lớn.\n3. Kiểm tra lý lịch pháp nhân đối tác và không nhận ủy thác trọn gói không rõ nguồn gốc hàng hóa bên trong container.",
        "opposing_views": "Phe A: Công ty nhận ủy thác chỉ mở tờ khai theo Invoice/Packing List do khách gửi nên không biết hàng giấu trong lõi gỗ vs Phe B: Theo pháp luật hải quan và hình sự, đơn vị đứng tên chủ hàng trên tờ khai phải chịu trách nhiệm pháp lý trực tiếp đối với toàn bộ hàng hóa thực tế trong container.",
        "lessons": "Đừng bao giờ bán rẻ chữ ký số của doanh nghiệp để đứng tên ủy thác nhập khẩu gỗ hoặc khoáng sản từ tuyến rủi ro cao mà không kiểm chứng thực tế tại đầu xếp hàng.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/hai-quan-phat-hien-nhieu-vu-vi-pham-co-dau-hieu-hinh-su-166527.html",
        "keywords": ["1.193", "nga voi"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn), Báo Tuổi Trẻ (tuoitre.vn), Báo Nhân Dân và Cổng TTĐT Chính phủ."
    },
    {
        "stt": 269,
        "week": "W47 — 19/11/2026",
        "title": "Gỗ Đức Thành (GDT) Bị Hải Quan KCX Tân Thuận Xử Phạt & Truy Thu Hơn 1,13 Tỷ Đồng Sau Kiểm Tra Sau Thông Quan Do Khai Sai Thuế Suất Nhập Khẩu & GTGT",
        "parties": "Công ty CP Chế biến Gỗ Đức Thành (mã chứng khoán GDT) + Chi cục Hải quan Khu chế xuất Tân Thuận (TP.HCM)",
        "plot": "Qua công tác kiểm tra sau thông quan đối với các tờ khai nhập khẩu nguyên vật liệu, thiết bị đã thông quan trước đó, Chi cục Hải quan KCX Tân Thuận phát hiện Công ty CP Chế biến Gỗ Đức Thành có hành vi khai sai thuế suất thuế nhập khẩu và thuế giá trị gia tăng (GTGT) dẫn đến thiếu số tiền thuế phải nộp. Doanh nghiệp đã nhận quyết định xử phạt vi phạm hành chính và công bố thông tin chính thức.",
        "damage": "Tổng số tiền bị ấn định truy thu, tiền chậm nộp và xử phạt vi phạm hành chính lên tới hơn 1,13 tỷ đồng (1.132 triệu đồng), ảnh hưởng trực tiếp đến dòng tiền và chi phí tài chính trong kỳ của doanh nghiệp niêm yết.",
        "root_cause": "Áp dụng mức thuế suất nhập khẩu và thuế suất GTGT chưa khớp với biểu thuế hiện hành tại thời điểm mở tờ khai và thiếu cơ chế rà soát lại biểu thuế định kỳ sau khi thông quan.",
        "prevention": "1. Thiết lập bảng đối chiếu mã HS – thuế suất NK – thuế suất GTGT chuẩn hóa trên hệ thống ERP trước khi mở tờ khai, cập nhật ngay khi có Nghị định biểu thuế mới.\n2. Thực hiện tự kiểm tra nội bộ các tờ khai trong vòng 60 ngày kể từ ngày thông quan để kịp thời khai bổ sung (AMA) và hưởng cơ chế không xử phạt theo Nghị định 169/2026/NĐ-CP.\n3. Kiểm soát chặt chẽ các mặt hàng có sự chuyển đổi thuế suất GTGT 8% vs 10% theo từng giai đoạn hiệu lực.",
        "opposing_views": "Phe A: Hàng đã được hệ thống VNACCS thông quan từ lâu thì lỗi áp thuế suất là sơ suất hành chính bình thường vs Phe B: Cơ quan hải quan có quyền kiểm tra sau thông quan trong 5 năm và phạt 20% số thuế khai thiếu nếu phát hiện qua thanh tra.",
        "lessons": "Ngay cả doanh nghiệp niêm yết uy tín cũng có thể mất hơn 1,13 tỷ đồng chỉ vì sai lệch thuế suất trên tờ khai cũ. Quy tắc 'tự kiểm tra trong 60 ngày vàng' giúp doanh nghiệp né án phạt 20%.",
        "url": "https://thuongtruong.com.vn/print/ke-khai-thue-sai-go-duc-thanh-bi-phat-hon-11-ty-dong-171267.html",
        "keywords": ["1,13", "Duc Thanh"],
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Thương Trường (thuongtruong.com.vn) và công bố thông tin chính thức của doanh nghiệp."
    },
    {
        "stt": 270,
        "week": "W47 — 19/11/2026",
        "title": "Bắt Giữ Lô Máy Móc Cũ Khai Mới 100% Trị Giá Hơn 10 Tỷ Đồng Tại Hữu Nghị & Doanh Nghiệp Khai Sai Mã HS Xe Chở Hàng Thiếu Thuế Hơn 2,19 Tỷ Đồng Tại Tân Thanh (2026)",
        "parties": "Công ty TNHH Vật liệu mới BBL HOME, Công ty CP TM&KT Công nghệ Tân Thời Đại + Chi cục Hải quan khu vực VI (Cửa khẩu Hữu Nghị & Tân Thanh, Lạng Sơn)",
        "plot": "Trong chiến dịch kiểm soát nhập khẩu chính ngạch năm 2026 tại Lạng Sơn, Hải quan khu vực VI phát hiện Công ty BBL HOME khai báo lô máy móc thiết bị đã qua sử dụng thành hàng mới 100% tại Cửa khẩu Hữu Nghị; cùng thời điểm, Công ty Tân Thời Đại bị phát hiện khai sai tên hàng, mã HS và thuế suất xe chở hàng tự hành phòng nổ tại Tân Thanh, còn 2 doanh nghiệp khác nhập đầu kéo SANY và bơm bê tông HOWO không đạt quy chuẩn.",
        "damage": "Lô máy móc BBL HOME trị giá hơn 10 tỷ đồng bị phạt 40 triệu đồng và buộc tái xuất trong 30 ngày (cùng lô SANY 2,6 tỷ và HOWO 7,5 tỷ bị buộc tái xuất); Công ty Tân Thời Đại bị truy thu hơn 2,19 tỷ đồng tiền thuế thiếu và phạt hơn 438,8 triệu đồng.",
        "root_cause": "Cố tình khai máy móc đã qua sử dụng thành hàng mới 100% để né Quyết định 18/2019/QĐ-TTg, hoặc thiếu kiểm tra kỹ thuật và mã HS xe chuyên dùng dẫn đến áp sai thuế suất hàng tỷ đồng.",
        "prevention": "1. Tuyệt đối kiểm tra năm sản xuất, tem nhãn gốc (Nameplate) và chứng thư giám định máy cũ đạt chuẩn QCVN/G7 trước khi đưa hàng lên cửa khẩu theo Quyết định 18/2019/QĐ-TTg.\n2. Đối với xe chuyên dùng, xe tự hành, máy công trình, phải lập hồ sơ xin xác định trước mã số (Pre-ruling) và đăng kiểm an toàn kỹ thuật trước khi nhập.\n3. Kiểm tra kỹ tỷ lệ thuế suất giữa xe chuyên dùng và xe vận tải thông thường trước khi ký hợp đồng mua bán.",
        "opposing_views": "Phe A: Máy móc sơn sửa lại nhìn như mới thì khai hàng mới cho nhanh thông quan khỏi tốn tiền giám định vs Phe B: Hải quan kiểm hóa thực tế và soi chiếu hồ sơ kỹ thuật sẽ phát hiện ngay máy cũ, dẫn đến buộc tái xuất lô hàng hơn 10 tỷ đồng và mất trắng cước vận chuyển 2 chiều.",
        "lessons": "Đừng bao giờ đánh cược lô hàng máy móc 10 tỷ đồng bằng cách khai 'hàng mới' hoặc ước lượng mã HS xe chuyên dùng. Chi phí giám định và xác định trước mã số chỉ bằng một phần nhỏ so với án phạt và lệnh tái xuất.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/hai-quan-khu-vuc-vi-xu-ly-nhieu-vi-pham-trong-nhap-khau-chinh-ngach-170078.html",
        "keywords": ["BBL", "Huu Nghi"],
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn) và Thương hiệu & Công luận (thuonghieucongluan.com.vn)."
    }
]


# ==============================================================================
# 3. DATA FOR HOC THUC CHIEN (14 cols) — STT 266..270 (W47 — 19/11/2026)
# ==============================================================================
HT_NEW_IDEAS = [
    {
        "stt": 266,
        "week": "W47 — 19/11/2026",
        "title": "Quy Trình 4 Bước Nộp Hồ Sơ Hải Quan 'Phi Địa Giới Hành Chính' Theo Quyết Định 3318/QĐ-BTC (133 TTHC): Xem Hàng, Hủy Tờ Khai & Khu Phi Thuế Quan Từ Xa",
        "steps": "Bước 1: Đối chiếu thủ tục cần làm (hủy tờ khai, xem hàng trước khi khai, gia hạn, hoàn thuế, khu phi thuế quan) xem có nằm trong Danh mục 133 TTHC Hải quan tại Quyết định 3318/QĐ-BTC.\nBước 2: Chuẩn bị bộ hồ sơ điện tử ký số đầy đủ (đơn đề nghị, vận đơn, công văn giải trình) theo mẫu chuẩn.\nBước 3: Chọn kênh nộp linh hoạt: nộp trực tuyến trên Cổng Dịch vụ công Quốc gia/Hải quan, gửi bưu chính hoặc nộp tại điểm tiếp nhận một cửa thuận tiện nhất không phụ thuộc tỉnh/thành.\nBước 4: Theo dõi mã hồ sơ trực tuyến và phối hợp lịch kiểm tra thực tế (nếu có) tại nơi lưu giữ hàng hóa.",
        "formula": "Phạm vi cải cách QĐ 3318/QĐ-BTC = 255 TTHC toàn Bộ Tài chính (trong đó 133 TTHC Hải quan + 51 TTHC Thuế + 58 TTHC Doanh nghiệp) | Giảm 100% chi phí di chuyển liên tỉnh nộp hồ sơ giấy.",
        "top3_errors": "1. Tưởng nhầm mọi thủ tục hải quan đều phi địa giới nên bỏ qua quy định kiểm tra thực tế hàng hóa tại nơi hàng đang lưu giữ.\n2. Nộp hồ sơ giấy bản cứng không đúng quy chuẩn ký số điện tử khi nộp qua Cổng dịch vụ công.\n3. Quên đối chiếu thời hạn giải quyết từng loại thủ tục (đặc biệt thủ tục hủy tờ khai quá hạn 15 ngày).",
        "mnemonic": "Nhớ quy tắc '133 Phi Địa Giới': Hồ sơ nộp bất cứ đâu trên toàn quốc — Hàng ở cảng nào thì kiểm thực tế ở cảng đó!",
        "comparison": "Trước 12/11/2026 (Nộp hồ sơ theo địa giới: trụ sở ở đâu/hàng ở đâu phải cử người đến đúng Chi cục đó) vs Sau QĐ 3318/QĐ-BTC (133 TTHC Hải quan tiếp nhận & xử lý phi địa giới toàn quốc qua Cổng dịch vụ công & điểm tiếp nhận thuận tiện).",
        "persona": "Anh Dũng (Giám đốc nhà máy có nhiều chi nhánh khác tỉnh) & Nhân viên phòng chứng từ XNK, đại lý hải quan.",
        "checklist": "☐ Tra cứu tên thủ tục trong danh mục 133 TTHC QĐ 3318/QĐ-BTC\n☐ Chuẩn bị hồ sơ điện tử gắn chữ ký số hợp lệ\n☐ Nộp qua Cổng Dịch vụ công hoặc điểm tiếp nhận thuận tiện\n☐ Lưu mã tiếp nhận và theo dõi kết quả trả về",
        "url": "https://thoibaotaichinhvietnam.vn/doanh-nghiep-huong-loi-tu-viec-bo-tai-chinh-cong-bo-133-thu-tuc-hai-quan-va-51-thu-tuc-thue-207032.html",
        "keywords": ["3318", "133"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Gắn trực tiếp SEO Keyword #2 (18/11/2026) & QĐ 3318/QĐ-BTC. Đối chiếu 3+ nguồn: Thời báo Tài chính Việt Nam, Tạp chí Kinh tế - Tài chính, LuatVietnam và Thư Viện Pháp Luật."
    },
    {
        "stt": 267,
        "week": "W47 — 19/11/2026",
        "title": "Cẩm Nang 5 Bước Ứng Phó Thuế Chống Bán Phá Giá 16% Của EU (Quy Định 2026/2547) & 4,47% Cáp Thép (QĐ 2927/QĐ-BCT): Rà Soát Mã HS & Quy Tắc 'Melt & Pour'",
        "steps": "Bước 1: Kiểm tra đối chiếu chính xác mã HS sản phẩm xuất/nhập khẩu (thép cán nguội EU: nhóm 7209, 7211, 7225, 7226; cáp thép VN: 7312.10.91, 7312.10.99) và danh mục loại trừ (thép không gỉ, thép điện, thép lá đen <0,35mm).\nBước 2: Xác định tên nhà sản xuất thực tế trên Mill Test Certificate (MTC) và kiểm tra mức thuế riêng theo QĐ 2927/QĐ-BCT (ví dụ 4,47% với Boxing Victory từ 19/11/2026).\nBước 3: Chuẩn hóa hồ sơ truy xuất nguồn gốc phôi thép 'nấu chảy và đúc' (Melt & Pour) theo yêu cầu mới của EU.\nBước 4: Cài điều khoản điều chỉnh giá / phân chia rủi ro thuế phòng vệ thương mại trong hợp đồng ngoại thương.\nBước 5: Lưu trữ sổ sách kế toán giá thành theo chuẩn quốc tế để sẵn sàng tham gia rà soát hành chính hằng năm.",
        "formula": "Giá nhập khẩu sau thuế EU (từ 17/11/2026) = Giá CIF × (1 + Thuế CBPG 16,0% + Phụ phí CBAM nếu áp dụng) | Biên độ so sánh EU: VN 16,0% vs Ấn Độ 9,5% vs Nhật 27,9%.",
        "top3_errors": "1. Không kiểm tra danh mục sản phẩm loại trừ (độ dày <0,35mm, thép không gỉ) nên tính sai chi phí dự toán hoặc chào giá mất đơn.\n2. Thiếu chứng từ 'Melt & Pour' chứng minh xuất xứ phôi thép khiến hải quan EU giữ hàng dù đã nộp thuế.\n3. Nhập cáp thép dự ứng lực sau mốc 19/11/2026 nhưng dùng tên thương nhân trung gian không khớp tên nhà sản xuất được hưởng thuế suất riêng 4,47%.",
        "mnemonic": "Nhớ bộ 3 lá chắn thép 2026: 'Check mã HS loại trừ — Khớp đúng tên Nhà sản xuất MTC — Cầm chắc chứng từ Melt & Pour'!",
        "comparison": "Thị trường EU áp thuế CBPG 16% với thép cán nguội VN (Quy định 2026/2547 hiệu lực 17/11/2026) vs Việt Nam áp thuế CBPG 4,47% cáp thép dự ứng lực nhập khẩu (QĐ 2927/QĐ-BCT áp dụng VNACCS 19/11/2026).",
        "persona": "Anh Minh (Chủ doanh nghiệp xuất nhập khẩu thép & vật liệu xây dựng) và Bộ phận mua hàng / xuất khẩu thép.",
        "checklist": "☐ Đối chiếu mã HS với danh mục chịu thuế & loại trừ\n☐ Kiểm tra tên nhà sản xuất gốc trên Mill Test Certificate\n☐ Chuẩn bị chứng từ truy xuất Melt & Pour (với hàng đi EU)\n☐ Cập nhật biểu thuế mới trên VNACCS từ 19/11/2026",
        "url": "https://chongbanphagia.vn/quy-dinh-moi-cua-eu-gia-tang-suc-ep-voi-thep-viet-n31398.html",
        "keywords": ["EU", "thep"],
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Gắn trực tiếp SEO Keyword #1 (18/11/2026) & Quy định (EU) 2026/2547 + QĐ 2927/QĐ-BCT. Đối chiếu 3+ nguồn: chongbanphagia.vn, vietstock.vn, congthuong.vn và Thư Viện Pháp Luật."
    },
    {
        "stt": 268,
        "week": "W47 — 19/11/2026",
        "title": "Quy Trình 5 Bước Phân Định Mã HS Cấu Kiện Thép (Nhóm 7308) vs Nhà Lắp Ghép (Nhóm 9406) & Hồ Sơ Miễn Thuế Tài Sản Cố Định Theo Công Văn 23970/CHQ-NVTHQ",
        "steps": "Bước 1: Thu thập bản vẽ kỹ thuật, danh mục cấu kiện đồng bộ và công năng thiết kế của công trình trước khi ký hợp đồng.\nBước 2: Phân định mã HS theo Chú giải HS: nếu là cột, kèo, xà gồ, kết cấu thép rời rạc thì áp nhóm 7308; chỉ áp nhóm 9406 khi là nhà lắp ghép hoàn chỉnh hoặc đồng bộ khung + vách + mái theo quy chuẩn.\nBước 3: Nếu xin miễn thuế nhập khẩu tạo tài sản cố định cho dự án ưu đãi đầu tư, kiểm tra Danh mục vật tư xây dựng trong nước đã sản xuất được của Bộ KH&ĐT.\nBước 4: Đăng ký Danh mục miễn thuế dự kiến trước khi mở tờ khai nhập khẩu.\nBước 5: Chuẩn bị bộ hồ sơ chứng minh trị giá giao dịch thực tế để sẵn sàng giải trình nếu bị tham vấn giá theo Công văn 23969/CHQ-NVTHQ.",
        "formula": "Điều kiện miễn thuế tài sản cố định (K11 Đ16 Luật Thuế XNK 2016 + NĐ 134/2016): Hàng hóa thuộc dự án ưu đãi đầu tư VÀ KHÔNG nằm trong Danh mục vật tư xây dựng trong nước đã sản xuất được.",
        "top3_errors": "1. Cố tình gộp các cấu kiện thép rời (nhóm 7308) khai thành 'Nhà lắp ghép đồng bộ' (nhóm 9406) để hưởng thuế thấp hoặc né kiểm tra.\n2. Quên kiểm tra Danh mục hàng trong nước đã sản xuất được khiến toàn bộ lô hàng bị bác miễn thuế khi kiểm tra sau thông quan.\n3. Khai trị giá gia công cấu kiện thép thấp bất thường so với giá thép nguyên liệu thế giới khiến tờ khai bị chuyển luồng đỏ tham vấn giá.",
        "mnemonic": "Thép rời là 7308 (tám mảnh ghép rời) — Nhà nguyên căn là 9406 (sáu vách khép kín) — Xin miễn thuế phải soi Danh mục trong nước chưa sản xuất được!",
        "comparison": "Nhóm 7308 (Cấu kiện, cột, kèo, xà gồ bằng sắt thép riêng lẻ) vs Nhóm 9406 (Nhà lắp ghép hoàn chỉnh đồng bộ hoặc ở dạng chưa lắp ráp đầy đủ bộ phận cấu thành ngôi nhà).",
        "persona": "Anh Dũng (Giám đốc dự án nhà xưởng FDI) & Chuyên viên khai báo hải quan vật tư dự án.",
        "checklist": "☐ Đọc kỹ bản vẽ kỹ thuật & mức độ đồng bộ lô hàng\n☐ Áp đúng nhóm 7308 vs 9406 trước khi mở tờ khai\n☐ Tra cứu Danh mục trong nước chưa sản xuất được\n☐ Chuẩn bị chứng từ chứng minh trị giá phòng tham vấn giá",
        "url": "https://luatvietnam.vn/xuat-nhap-khau/cong-van-23970-chq-nvthq-2026-cua-cuc-hai-quan-thao-go-kho-khan-va-bao-dam-canh-tranh-nganh-ket-cau-thep-455345-d6.html",
        "keywords": ["23970", "cau kien thep"],
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Căn cứ Công văn 23969 & 23970/CHQ-NVTHQ (06/11/2026). Đối chiếu 3+ nguồn: LuatVietnam, Báo Công Thương và Thư Viện Pháp Luật."
    },
    {
        "stt": 269,
        "week": "W47 — 19/11/2026",
        "title": "Checklist 4 Bước Chuẩn Hóa Hồ Sơ Xuất Khẩu Thực Phẩm Sang Trung Quốc Theo Thông Báo 150/2026 Của GACC: Từ 20 Xuống 10 Mục, Duyệt 3 Ngày & Mốc Khóa 60 Ngày",
        "steps": "Bước 1: Đăng nhập Cổng thông tin một cửa thương mại quốc tế / Dịch vụ công của Hải quan Trung Quốc để kiểm tra biểu mẫu đăng ký mới (rút gọn còn 10 mục thông tin).\nBước 2: Chuẩn hóa chính xác 10 trường thông tin cốt lõi của nhà xuất khẩu/đại lý nước ngoài (tên pháp nhân, mã số thuế/đăng ký kinh doanh, địa chỉ, loại hình, người liên hệ) và nộp trực tuyến.\nBước 3: Theo dõi kết quả xét duyệt trong thời hạn tối đa 03 ngày làm việc kể từ ngày nộp đủ hồ sơ hợp lệ.\nBước 4: Thiết lập lịch nhắc nội bộ: nếu phát sinh bất kỳ thay đổi nào về pháp nhân, địa chỉ hoặc thông tin đại lý, BẮT BUỘC cập nhật trên hệ thống GACC trong vòng 60 ngày.",
        "formula": "Cải cách GACC 150/2026: Số mục kê khai giảm 50% (20 -> 10 mục) | Thời gian xét duyệt = tối đa 03 ngày làm việc | Thời hạn bắt buộc cập nhật thay đổi = tối đa 60 ngày.",
        "top3_errors": "1. Dùng mẫu biểu cũ 20 mục hoặc kê khai sai tên tiếng Anh của doanh nghiệp không khớp giấy phép kinh doanh.\n2. Thay đổi trụ sở công ty hoặc người đại diện tại Việt Nam nhưng quên cập nhật lên hệ thống GACC quá hạn 60 ngày.\n3. Nhầm lẫn giữa mã số đăng ký nhà sản xuất (CIFER theo Lệnh 248) với mã đăng ký nhà xuất khẩu/đại lý thương mại.",
        "mnemonic": "Khẩu quyết GACC 2026: '10 mục gọn nhẹ — 3 ngày cấp mã — 60 ngày phải báo khi đổi thông tin'!",
        "comparison": "Quy định GACC cũ (kê khai 20 mục phức tạp, thời gian chờ duyệt kéo dài) vs Thông báo 150/2026 của GACC (chỉ 10 mục kê khai, xét duyệt trong 3 ngày làm việc, siết hạn cập nhật 60 ngày).",
        "persona": "Chị Hương (Chủ doanh nghiệp xuất khẩu thực phẩm, nông sản chế biến) & Nhân viên chứng từ xuất khẩu tuyến Trung Quốc.",
        "checklist": "☐ Đối chiếu tên pháp nhân & địa chỉ chuẩn tiếng Anh\n☐ Điền đủ 10 mục theo Phụ lục Thông báo 150/2026\n☐ Kiểm tra kết quả cấp mã sau 03 ngày làm việc\n☐ Ghi lịch nhắc cập nhật thông tin trong 60 ngày nếu có thay đổi",
        "url": "https://congthuong.vn/trung-quoc-sua-quy-dinh-dang-ky-doanh-nghiep-xuat-khau-thuc-pham-479439.html",
        "keywords": ["GACC", "150"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Căn cứ Thông báo 150/2026 của GACC triển khai 11/2026. Đối chiếu 2+ nguồn: Báo Công Thương, Cổng TTĐT Bộ Công Thương (moit.gov.vn) và Chi cục Trồng trọt & BVTV TP.HCM."
    },
    {
        "stt": 270,
        "week": "W47 — 19/11/2026",
        "title": "Kỹ Thuật 5 Bước Chuẩn Bị Hồ Sơ Phân Loại & Tiền Chất Hóa Chất (CY-9, Potassium Gold Cyanide 28.43) Theo Công Văn 24288/CHQ-GSQL & Nghị Định 24/2026/NĐ-CP",
        "steps": "Bước 1: Yêu cầu nhà sản xuất nước ngoài cung cấp MSDS đầy đủ 16 mục và Bảng thành phần tỷ lệ % chính xác đến từng chất (COA/Composition Sheet) trước khi tàu cập cảng.\nBước 2: Kiểm tra xem thành phần hóa chất có chứa hợp chất vàng/kim loại quý (nhóm 28.43 như Potassium Gold Cyanide) hoặc nhóm chất xyanua/tiền chất bị kiểm soát theo Nghị định 24/2026/NĐ-CP hay không.\nBước 3: Xin Giấy phép nhập khẩu / Khai báo hóa chất qua Cổng Một cửa Quốc gia trước khi mở tờ khai.\nBước 4: Với chất hỗn hợp điều chế (như chất xúc tiến lưu hóa cao su CY-9) mà phòng kiểm định trong nước chưa đủ thiết bị tách hàm lượng copolymer, chuẩn bị sẵn tài liệu kỹ thuật gốc của nhà sản xuất để nộp bổ sung giải trình.\nBước 5: Theo dõi thông báo phân tích phân loại để xử lý kịp thời, tránh lưu container hóa chất nguy hiểm phí bãi cao.",
        "formula": "Ngưỡng rủi ro hóa chất hỗn hợp: Khi cơ quan kiểm định chưa đủ máy tách hàm lượng polymer/copolymer -> Thời gian giải phóng hàng phụ thuộc 100% vào Bộ tài liệu kỹ thuật gốc từ nhà sản xuất + Giấy phép NĐ 24/2026/NĐ-CP.",
        "top3_errors": "1. Chỉ cầm Invoice ghi tên thương mại (ví dụ 'CY-9') mà không có Bảng thành phần % chi tiết của nhà sản xuất khiến hải quan không thể phân loại mã HS.\n2. Bỏ sót thành phần gốc Xyanua hoặc tiền chất trong hợp chất mạ/xử lý bề mặt nên không xin giấy phép theo Nghị định 24/2026/NĐ-CP.\n3. Không dự trù thời gian lấy mẫu phân tích phân loại, dẫn đến container hóa chất nằm bãi phát sinh phí lưu bãi hàng nguy hiểm (DG storage) gấp 2-3 lần hàng thường.",
        "mnemonic": "Nhập hóa chất nhớ '3 Tấm Vé Thông Hành': MSDS 16 mục + Bảng % thành phần gốc Nhà sản xuất + Giấy phép NĐ 24/2026!",
        "comparison": "Hóa chất đơn chất tinh khiết (dễ định danh nhóm 28/29 như Potassium Gold Cyanide nhóm 28.43) vs Hóa chất hỗn hợp điều chế (như chất xúc tiến cao su CY-9 cần tài liệu kỹ thuật nhà sản xuất bổ trợ khi kiểm định chưa đủ điều kiện tách chất).",
        "persona": "Anh Dũng (Giám đốc nhà máy sản xuất cao su, xi mạ, điện tử) & Nhân viên khai báo hải quan hàng hóa chất.",
        "checklist": "☐ Thu thập MSDS 16 mục & Bảng % thành phần có dấu NSX\n☐ Kiểm tra danh mục hóa chất kiểm soát theo NĐ 24/2026/NĐ-CP\n☐ Khai báo hóa chất / xin giấy phép trước khi hàng về\n☐ Dự phòng tài liệu kỹ thuật gốc để giải trình nếu lấy mẫu kiểm định",
        "url": "https://luatvietnam.vn/xuat-nhap-khau/cong-van-24288-chq-gsql-2026-cua-cuc-hai-quan-ve-vuong-mac-thu-tuc-hai-quan-hoa-chat-456557-d6.html",
        "keywords": ["24288", "Cyanide"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Căn cứ Công văn 24288/CHQ-GSQL (16/11/2026) & Nghị định 24/2026/NĐ-CP. Đối chiếu 3+ nguồn: LuatVietnam, BKHUB và Thư Viện Pháp Luật."
    }
]


# ==============================================================================
# 4. DATA FOR FUNFACT (13 cols) — STT 266..270 (W47 — 19/11/2026)
# ==============================================================================
FF_NEW_IDEAS = [
    {
        "stt": 266,
        "week": "W47 — 19/11/2026",
        "title": "Liên Minh Thương Mại Hanseatic (Thế Kỷ 12–17): Mạng Lưới 'Cảng Tự Do & Luật Biển Riêng' Của Các Thương Hội Bắc Âu Trước Khi Có Luật Thương Mại Quốc Tế",
        "hook": "Từ thế kỷ 12 đến thế kỷ 17, khi Châu Âu còn chưa có bất kỳ hiệp định FTA hay tổ chức WTO nào, một liên minh gồm các thương hội và thành phố cảng Đức đã tự lập hạm đội, tự đặt 'luật biển chung' và kiểm soát gần như toàn bộ thương mại Biển Bắc – Biển Baltic suốt 300 năm!",
        "history": "Liên minh Hanseatic (Hanseatic League), hình thành từ cuối thế kỷ 12 quanh trung tâm Lübeck và Hamburg (Đức), kết nối hàng trăm thành phố thương mại và các thương điếm nước ngoài (Kontor) từ London đến Novgorod.",
        "context": "Thời Trung cổ, tàu buôn đi biển đối mặt cướp biển, thuế quan tùy tiện của lãnh chúa phong kiến và sự thiếu vắng tòa án thương mại xuyên biên giới. Các thương nhân Đức đã liên kết thành hiệp hội phòng vệ chung, cử tàu hộ tống và thống nhất tiêu chuẩn cân đo, tiền tệ, bảo hiểm hàng hải.",
        "modern_impact": "Luật hàng hải Hanseatic và mô hình thương điếm miễn thuế (Kontor) chính là tiền thân trực tiếp của các hiệp định tạo thuận lợi thương mại, cảng trung chuyển tự do và mạng lưới đại lý forwarder quốc tế ngày nay.",
        "intl_comp": "Nếu Con đường Tơ lụa đường bộ dựa vào các đoàn lạc đà qua sa mạc, thì Liên minh Hanseatic chứng minh sức mạnh của 'chuỗi cảng biển liên kết' (Maritime Port Network) dùng luật chơi chung để giảm chi phí giao dịch xuyên quốc gia.",
        "business_lesson": "Doanh nghiệp XNK Việt Nam không thể đi xa nếu đơn độc: tham gia hiệp hội ngành hàng uy tín, chọn đối tác forwarder có mạng lưới đại lý chuẩn mực toàn cầu giúp bảo vệ lô hàng khi xảy ra tranh chấp ở cảng ngoại.",
        "url": "https://en.wikipedia.org/wiki/Hanseatic_League",
        "keywords": ["Hanseatic", "League"],
        "series": "Series: Lịch sử Logistics",
        "status": "Chưa dùng",
        "note": "Hướng 1 (Cột mốc lịch sử thương mại). Nguồn xác minh Wikipedia (Hanseatic League), đối chiếu Britannica."
    },
    {
        "stt": 267,
        "week": "W47 — 19/11/2026",
        "title": "Tàu Ro-Ro (Roll-on/Roll-off) & 'Cầu Nối Nổi' Biển Cả: Vì Sao Tàu Chở Ô Tô Không Dùng Cần Cẩu Lại Xếp Dỡ Nhanh Gấp Nhiều Lần Tàu Container?",
        "hook": "Trong khi tàu container cần hàng loạt cẩu giàn khổng lồ nhấc từng thùng thép lên xuống (Lo-Lo), tàu Ro-Ro chỉ cần hạ một chiếc cầu dẫn (Ramp) là hàng nghìn chiếc ô tô và máy công trình tự lăn bánh lên xuống boong tàu như đi qua một bãi đỗ xe nổi!",
        "history": "Khái niệm tàu Roll-on/Roll-off (Ro-Ro) phát triển từ các phà chở toa tàu hỏa thế kỷ 19 và tàu đổ bộ xe tăng Thế chiến II, sau đó bùng nổ thành đội tàu Pure Car Carrier (PCC) phục vụ xuất khẩu ô tô toàn cầu từ thập niên 1960–1970.",
        "context": "Đưa từng chiếc ô tô hay xe tải nặng vào container hoặc dùng cần cẩu cẩu từng chiếc cực kỳ tốn thời gian và dễ va quẹt móp méo. Kỹ sư hàng hải đã thiết kế các boong tàu nhiều tầng nối với cầu cảng bằng dốc thủy lực chịu tải hàng trăm tấn, kết hợp đầu kéo chuyên dụng và rơ-moóc lùn (Mafi trailer).",
        "modern_impact": "Ngày nay, tàu Ro-Ro là xương sống của ngành xuất nhập khẩu ô tô nguyên chiếc (CBU), máy xúc, thiết bị quá khổ quá tải (OOG) cập cảng Hải Phòng, Cái Mép và Hiệp Phước (TP.HCM).",
        "intl_comp": "Khác với tàu Lo-Lo (Lift-on/Lift-off) tính công suất theo TEU container, tàu Ro-Ro đo sức chở bằng làn xe (Lane meters) hoặc số lượng xe tiêu chuẩn (CEU — Car Equivalent Unit), rút ngắn thời gian tàu nằm cảng xuống chỉ còn vài giờ.",
        "business_lesson": "Chọn đúng phương thức vận tải quyết định trực tiếp chi phí logistics: với máy móc tự hành, xe cơ giới hoặc hàng siêu trường siêu trọng nguyên khối, phương án Ro-Ro hoặc Break-bulk thường an toàn và kinh tế hơn cố nhồi nhét container.",
        "url": "https://en.wikipedia.org/wiki/Roll-on/roll-off",
        "keywords": ["Roll-on", "RORO"],
        "series": "Series: Lịch sử Logistics",
        "status": "Chưa dùng",
        "note": "Hướng 1 & 5 (Lịch sử vận tải & kỹ thuật xếp dỡ). Nguồn xác minh Wikipedia (Roll-on/roll-off), đối chiếu IMO / Maritime manuals."
    },
    {
        "stt": 268,
        "week": "W47 — 19/11/2026",
        "title": "Tuyến Đường Biển Phương Bắc (Northern Sea Route): Hành Lang Bắc Cực Rút Ngắn Gần 40% Hải Trình Á – Âu So Với Kênh Đào Suez",
        "hook": "Một con tàu đi từ Đông Á sang Bắc Âu qua Bắc Cực có thể tiết kiệm tới khoảng 40% quãng đường so với đi qua Kênh đào Suez, nhưng lại đòi hỏi tàu phá băng hộ tống và vỏ tàu gia cường chịu băng dày hàng mét!",
        "history": "Tuyến đường biển Phương Bắc (Northern Sea Route — NSR) chạy dọc bờ biển Siberia thuộc vùng biển Bắc Cực của Nga từ eo biển Kara đến eo biển Bering, là phần khai thác chính của Hành lang Đông Bắc (Northeast Passage).",
        "context": "Khi biến đổi khí hậu làm băng Bắc Cực tan mỏng hơn vào mùa hè và căng thẳng địa chính trị tại Biển Đỏ – Suez gia tăng, các hãng vận tải năng lượng và container bắt đầu thử nghiệm tuyến NSR để cắt giảm thời gian hành trình từ ~35–40 ngày xuống còn ~20–25 ngày.",
        "modern_impact": "Dù rút ngắn quãng đường, tuyến NSR vẫn đối mặt chi phí bảo hiểm băng giá đắt đỏ, phí thuê tàu phá băng hạt nhân, cửa sổ thời tiết ngắn và rủi ro môi trường cực cao, nên chưa thể thay thế hoàn toàn tuyến truyền thống qua Singapore – Malacca – Suez.",
        "intl_comp": "Đặt cạnh 2 'huyết mạch' nhân tạo là Kênh đào Suez và Kênh đào Panama, tuyến NSR cho thấy bản đồ logistics toàn cầu luôn biến động theo cả địa chính trị lẫn khí hậu tự nhiên.",
        "business_lesson": "Đừng bao giờ phụ thuộc duy nhất vào một tuyến vận chuyển hay một cảng trung chuyển: khi lập kế hoạch giao hàng dài hạn sang Châu Âu, doanh nghiệp Việt cần luôn dự phòng phương án đường biển vòng Mũi Hảo Vọng, đường sắt liên vận hoặc kết hợp Sea-Air.",
        "url": "https://en.wikipedia.org/wiki/Northern_Sea_Route",
        "keywords": ["Northern Sea Route", "Arctic"],
        "series": "Series: Thương mại Quốc tế",
        "status": "Chưa dùng",
        "note": "Hướng 2 & 3 (So sánh hành lang thương mại quốc tế). Nguồn xác minh Wikipedia (Northern Sea Route), đối chiếu Arctic Portal."
    },
    {
        "stt": 269,
        "week": "W47 — 19/11/2026",
        "title": "Bí Mật Kích Thước 800 × 1.200 mm Của Pallet Chuẩn Châu Âu (EUR/EPAL Pallet, 1961): Cách 78 Chiếc Đinh Chuẩn Hóa Cắt Giảm Thời Gian Bốc Xếp Toàn Cầu",
        "hook": "Trước năm 1961, mỗi nhà máy và toa tàu Châu Âu dùng một cỡ kệ gỗ khác nhau khiến xe nâng gần như vô dụng khi đổi phương tiện — cho đến khi chuẩn EUR-pallet (800 × 1.200 × 144 mm) với đúng 78 chiếc đinh được thống nhất trên toàn lục địa!",
        "history": "Năm 1961, Liên minh Đường sắt Quốc tế (UIC) và các quốc gia Châu Âu chính thức ký thỏa thuận chuẩn hóa EUR-pallet (nay thuộc quản lý của Hiệp hội Pallet Châu Âu EPAL), nặng khoảng 25 kg và chịu tải động 1.500 kg.",
        "context": "Nhờ kích thước 800 × 1.200 mm ăn khớp tuyệt đối với chiều rộng toa tàu hỏa, thùng xe tải tiêu chuẩn Châu Âu và các dòng thùng carton theo hệ mô-đun (ISO 3394), hàng trăm triệu chiếc EUR-pallet có thể luân chuyển đổi ngang (Pallet Exchange) giữa nhà máy, kho lạnh và siêu thị mà không cần dỡ hàng sang kệ khác.",
        "modern_impact": "Khi xuất khẩu hàng hóa sang thị trường EU, việc đóng gói thùng carton vừa khít chuẩn EUR-pallet (và xử lý hun trùng nhiệt ISPM 15) giúp doanh nghiệp Việt tối ưu thể tích container và tránh bị siêu thị đối tác phạt phí dỡ hàng thủ công.",
        "intl_comp": "Trong khi Châu Âu chuộng chuẩn EUR/EPAL 800 × 1.200 mm, Bắc Mỹ lại phổ biến chuẩn GMA Pallet 48 × 40 inch (1.219 × 1.016 mm), còn Châu Á và container biển hay dùng chuẩn 1.100 × 1.100 mm để xếp vừa khít 2 hàng trong container khô 20ft/40ft.",
        "business_lesson": "Chuẩn hóa kích thước đóng gói ngay từ nhà máy (Carton -> Pallet -> Container) là bài toán tiết kiệm cước thầm lặng: xếp khéo vừa khít chuẩn pallet điểm đến có thể tiết kiệm 10–15% số lượng container cả năm.",
        "url": "https://en.wikipedia.org/wiki/EUR-pallet",
        "keywords": ["EUR-pallet", "800"],
        "series": "Series: Đàm phán & Tư duy",
        "status": "Chưa dùng",
        "note": "Hướng 4 (Chuẩn hóa vận hành & tư duy tối ưu chuỗi cung ứng). Nguồn xác minh Wikipedia (EUR-pallet), đối chiếu EPAL."
    },
    {
        "stt": 270,
        "week": "W47 — 19/11/2026",
        "title": "Mạng Lưới Điện Tín Liên Ngân Hàng SWIFT (1973): 'Hệ Thần Kinh Tài Chính' Truyền Tải Điện L/C (MT700) & Chuyển Tiền (MT103) Nhưng Không Giữ Một Đồng Nào",
        "hook": "Hiệp hội Viễn thông Tài chính Liên ngân hàng Toàn cầu (SWIFT) kết nối hơn 11.000 ngân hàng và định chế tại trên 200 quốc gia để luân chuyển hàng nghìn tỷ USD giao dịch ngoại thương mỗi ngày — nhưng bản thân hệ thống SWIFT không hề cầm giữ một xu tiền mặt nào!",
        "history": "SWIFT được thành lập năm 1973 tại Bỉ bởi 239 ngân hàng từ 15 quốc gia nhằm thay thế hệ thống điện tín Telex thủ công chậm chạp và dễ sai sót trong thanh toán quốc tế.",
        "context": "Thay vì chuyển tiền vật lý, SWIFT hoạt động như một 'hệ thống tin nhắn mã hóa chuẩn mực toàn cầu' (chuẩn ISO 20022 & các điện MT700 phát hành L/C, MT103 chuyển tiền T/T). Mọi giao dịch xuất nhập khẩu qua ngân hàng đều đi qua mã định danh BIC/SWIFT 8 hoặc 11 ký tự để xác thực giữa ngân hàng người mua và ngân hàng người bán.",
        "modern_impact": "Tại Việt Nam, hiểu đúng bản chất điện MT103 và điện phát hành thư tín dụng MT700 giúp doanh nghiệp XNK phân biệt ngay chứng từ chuyển tiền thật hay lệnh chuyển tiền cắt ghép lừa đảo (BEC scam).",
        "intl_comp": "Bên cạnh SWIFT của phương Tây, Trung Quốc phát triển hệ thống CIPS phục vụ thanh toán nhân dân tệ xuyên biên giới, cho thấy hạ tầng thông điệp thanh toán là huyết mạch sống còn của thương mại quốc tế.",
        "business_lesson": "Trong thanh toán T/T hoặc L/C quốc tế, một bức ảnh chụp màn hình 'đã chuyển tiền' hoàn toàn có thể bị làm giả; chỉ khi ngân hàng Việt Nam xác nhận nhận được điện SWIFT MT103/MT700 hợp lệ hoặc tiền thực sự báo Có vào tài khoản mới an toàn giao vận đơn gốc.",
        "url": "https://en.wikipedia.org/wiki/SWIFT",
        "keywords": ["SWIFT", "interbank"],
        "series": "Series: Công nghệ & Đổi mới",
        "status": "Chưa dùng",
        "note": "Hướng 5 (Công nghệ tài chính thương mại toàn cầu). Nguồn xác minh Wikipedia (SWIFT), đối chiếu swift.com."
    }
]


# ==============================================================================
# 5. FEEDBACK LOOP: MISSING LEGAL DOCS -> Danh_sach_Van_ban_Phap_luat_XNK.xlsx
# ==============================================================================
NEW_LEGAL_DOCS = [
    {
        "ma_so": "Quyết định 2927/QĐ-BCT",
        "ten": "Quyết định về kết quả rà soát lần thứ nhất việc áp dụng biện pháp chống bán phá giá đối với một số sản phẩm cáp thép dự ứng lực có xuất xứ từ Malaysia, Thái Lan và Trung Quốc (mã vụ việc AR01.AD17)",
        "loai": "Quyết định",
        "co_quan": "Bộ Công Thương",
        "ngay_ban_hanh": "2026-11-04",
        "ngay_hieu_luc": "2026-11-19",
        "ngay_het_hieu_luc": "Chưa xác định",
        "van_ban_bi_thay_the": "Điều chỉnh kết quả rà soát AD17",
        "phan_tich": "Công bố kết quả rà soát lần thứ nhất biện pháp chống bán phá giá cáp thép dự ứng lực (HS 7312.10.91, 7312.10.99); ấn định mức thuế 4,47% đối với nhóm công ty Boxing Victory và Cục Hải quan triển khai áp dụng trên VNACCS từ 19/11/2026.",
        "so_sanh": "Cập nhật điều chỉnh mức thuế chống bán phá giá theo từng nhà sản xuất xuất khẩu cụ thể trên hệ thống khai báo hải quan tự động VNACCS."
    },
    {
        "ma_so": "Quy định thực thi (EU) 2026/2547",
        "ten": "Quy định thực thi (EU) 2026/2547 của Ủy ban Châu Âu áp thuế chống bán phá giá chính thức đối với thép cán nguội dạng phẳng nhập khẩu từ Việt Nam, Nhật Bản, Ấn Độ, Đài Loan và Thổ Nhĩ Kỳ",
        "loai": "Quy định Phòng vệ thương mại Quốc tế",
        "co_quan": "Ủy ban Châu Âu (EC)",
        "ngay_ban_hanh": "2026-11-13",
        "ngay_hieu_luc": "2026-11-17",
        "ngay_het_hieu_luc": "2031-11-17",
        "van_ban_bi_thay_the": "Lần đầu áp thuế CBPG chính thức vụ việc thép cán nguội",
        "phan_tich": "EC ấn định thuế chống bán phá giá chính thức 16,0% đối với thép cán nguội dạng phẳng xuất khẩu từ Việt Nam (nhóm 7209, 7211, 7225, 7226, loại trừ thép không gỉ, thép điện và thép lá đen dưới 0,35mm); đồng thời bác bỏ cáo buộc tình hình thị trường đặc biệt (PMS).",
        "so_sanh": "Chuyển từ trạng thái điều tra (khởi xướng 09/2025) sang áp thuế chống bán phá giá chính thức 16,0% tại cửa khẩu EU kể từ 17/11/2026."
    },
    {
        "ma_so": "Thông báo 150/2026/GACC",
        "ten": "Thông báo số 150 năm 2026 của Tổng cục Hải quan Trung Quốc (GACC) về cải cách, đơn giản hóa thủ tục đăng ký doanh nghiệp xuất khẩu, nhập khẩu thực phẩm nước ngoài",
        "loai": "Quy định Hải quan Quốc tế (Trung Quốc)",
        "co_quan": "Tổng cục Hải quan Trung Quốc (GACC)",
        "ngay_ban_hanh": "2026-10-23",
        "ngay_hieu_luc": "2026-10-23",
        "ngay_het_hieu_luc": "Chưa xác định",
        "van_ban_bi_thay_the": "Sửa đổi, tinh giản thủ tục đăng ký nhà xuất khẩu thực phẩm GACC",
        "phan_tich": "Giảm mẫu kê khai đăng ký nhà xuất khẩu và đại lý thực phẩm nước ngoài từ 20 mục xuống còn 10 mục; rút ngắn thời hạn xét duyệt cấp mã xuống còn 03 ngày làm việc qua Cổng một cửa; bắt buộc cập nhật thông tin trong 60 ngày khi có thay đổi.",
        "so_sanh": "Cắt giảm 50% trường thông tin kê khai và rút ngắn thời gian xét duyệt xuống 03 ngày làm việc, đồng thời siết trách nhiệm tự khai và hạn cập nhật 60 ngày."
    }
]


def check_banned_words():
    """Verify zero prohibited words from banned_words.md appear anywhere."""
    all_records = TN_NEW_IDEAS + CS_NEW_IDEAS + HT_NEW_IDEAS + FF_NEW_IDEAS + NEW_LEGAL_DOCS
    violations = []
    for idx, rec in enumerate(all_records, start=1):
        combined = " ".join(str(v) for v in rec.values() if v is not None).lower()
        for bw in BANNED_WORDS_CHECK:
            if bw in combined:
                violations.append((idx, rec.get("title") or rec.get("ma_so"), bw))
        if "chính ngạch" in combined and "tiểu ngạch" in combined:
            violations.append((idx, rec.get("title") or rec.get("ma_so"), "so sánh chính ngạch vs tiểu ngạch"))
    if violations:
        raise ValueError(f"Phát hiện từ cấm vi phạm banned_words.md: {violations}")
    print("🟢 BANNED WORDS CHECK PASSED: 0 vi phạm từ cấm trên toàn bộ 20 ý tưởng & văn bản luật!")


def check_no_duplicates():
    """Ensure STT 266-270 and titles do not duplicate existing Excel rows."""
    all_new_titles = {
        "reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx": [x["title"].strip().lower() for x in TN_NEW_IDEAS],
        "reports/Case_Study_Ideas_Bank_Eureka.xlsx": [x["title"].strip().lower() for x in CS_NEW_IDEAS],
        "reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx": [x["title"].strip().lower() for x in HT_NEW_IDEAS],
        "reports/Funfact_Ideas_Bank_Eureka.xlsx": [x["title"].strip().lower() for x in FF_NEW_IDEAS],
    }
    for path, new_titles in all_new_titles.items():
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        last_stt = ws.cell(row=ws.max_row, column=1).value
        if last_stt != 265:
            wb.close()
            raise ValueError(f"{path} last STT is {last_stt}, expected 265 before appending 266-270!")
        existing_titles = {str(ws.cell(r, 3).value or "").strip().lower() for r in range(5, ws.max_row + 1)}
        for nt in new_titles:
            if nt in existing_titles:
                wb.close()
                raise ValueError(f"Trùng tiêu đề trong {path}: {nt}")
        wb.close()
    print("🟢 DUPLICATE & SEQUENCE CHECK PASSED: 4 file Ideas Bank sẵn sàng nối tiếp STT 266-270.")


def execute_update():
    print("=" * 80)
    print("🚀 WORKFLOW 06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (WEEK 47 DAY 4 — 19/11/2026)")
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
    print(f"✅ Ghi thành công {tn_path} (Tổng dòng: {ws_tn.max_row}, STT cuối: 270)")

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
    print(f"✅ Ghi thành công {cs_path} (Tổng dòng: {ws_cs.max_row}, STT cuối: 270)")

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
    print(f"✅ Ghi thành công {ht_path} (Tổng dòng: {ws_ht.max_row}, STT cuối: 270)")

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
    print(f"✅ Ghi thành công {ff_path} (Tổng dòng: {ws_ff.max_row}, STT cuối: 270)")

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
            "2026-11-19"
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
    print("🎉 HOÀN TẤT WORKFLOW 06 (19/11/2026): ĐÃ CẬP NHẬT 20 Ý TƯỞNG MỚI (STT 266-270) & FEEDBACK LOOP!")
    print("=" * 80)


if __name__ == "__main__":
    execute_update()
