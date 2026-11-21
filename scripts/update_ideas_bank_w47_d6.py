# -*- coding: utf-8 -*-
"""
Execute Workflow 06: Ideas Bank Daily Update (wf06) - Week 47 Day 6 (21/11/2026)
Appends 5 new verified ideas (STT 271-275) to each of the 4 Pillars (TN, CS, HT, FF).
Runs through Technical Gatekeeper url_validator.py and Banned Words check before saving.
Updates Danh_sach_Van_ban_Phap_luat_XNK.xlsx (Feedback Loop) for missing verified regulations.
Uses atomic save (write to .tmp.xlsx and os.replace) to strictly avoid any backup files.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from scripts.url_validator import enforce_strict_validation

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

BORDER_THIN = Border(
    left=Side(style="thin", color="E5E7EB"),
    right=Side(style="thin", color="E5E7EB"),
    top=Side(style="thin", color="E5E7EB"),
    bottom=Side(style="thin", color="E5E7EB"),
)
FONT_BOLD = Font(name="Segoe UI", size=10, bold=True)
FONT_NORMAL = Font(name="Segoe UI", size=10, bold=False)
FILL_EVEN = PatternFill(fill_type="solid", start_color="F9FAFB", end_color="F9FAFB")
FILL_ODD = PatternFill(fill_type="solid", start_color="FFFFFF", end_color="FFFFFF")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

BANNED_WORDS_CHECK = [
    "trốn thuế", "lách luật", "hàng lậu", "hàng nhái", "hàng giả",
    "bao thuế", "đi chui", "đi lậu", "tiểu ngạch",
    "mua hàng trung quốc tận gốc", "đánh hàng quảng châu", "gom hàng biên", "chạy line",
]


def atomic_save(workbook, target_path):
    """Save workbook via temporary file and atomic replace to prevent locks and backups."""
    tmp_path = target_path + ".tmp.xlsx"
    workbook.save(tmp_path)
    os.replace(tmp_path, target_path)


# ==============================================================================
# 1. DATA FOR TIN NGANH (15 cols) — STT 271..275 (W47 — 21/11/2026)
# ==============================================================================
TN_NEW_IDEAS = [
    {
        "stt": 271,
        "week": "W47 — 21/11/2026",
        "title": "Thông Tư Số 71/2026/TT-BCT (Ban Hành 19/11/2026, Hiệu Lực 08/01/2027): Chuẩn Hóa 5 Loại Hình & Phân Hạng 3 Cấp Trung Tâm Logistics Toàn Quốc Trên TradeNode.vn",
        "agency": "Bộ Công Thương",
        "effective_date": "Ban hành ngày 19/11/2026; chính thức có hiệu lực thi hành từ ngày 08/01/2027",
        "core_diff": "Cũ: Chưa có bộ tiêu chí pháp lý thống nhất để phân loại, phân hạng các trung tâm logistics và chưa có nền tảng tự công bố tập trung; Mới: Thông tư 71/2026/TT-BCT phân chia 5 loại hình trung tâm logistics (theo phương thức vận tải, ngành hàng, khu chế xuất/FTZ, cửa khẩu biên giới, thương mại điện tử) và 3 hạng (Hạng I >= 20 ha, xử lý >= 3 triệu tấn hoặc 200.000 TEU/năm, kết nối >= 2 phương thức vận tải, bãi đỗ >= 200 xe container; Hạng II từ 10–<20 ha; Hạng III từ 3–<10 ha), đồng thời yêu cầu tự công bố trên TradeNode.vn, tích hợp hạ tầng xanh & trạm sạc điện.",
        "affected_entity": "Chủ đầu tư phát triển hạ tầng khu công nghiệp, cảng cạn ICD, trung tâm logistics, doanh nghiệp 3PL/4PL và các sàn thương mại điện tử vận hành kho trung chuyển lớn.",
        "risk_pain": "Các trung tâm kho vận chưa nâng cấp phần mềm WMS, camera 24/7, kiểm soát thẻ thông minh hoặc chưa đáp ứng tiêu chí hạ tầng năng lượng sạch có nguy cơ bị rớt hạng khi tự công bố trên TradeNode.vn, mất lợi thế đấu thầu với các tập đoàn FDI.",
        "market_reaction": "Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA) và giới đầu tư hạ tầng đánh giá cao chuẩn hóa 3 hạng giúp minh bạch năng lực thực chất, chấm dứt tình trạng kho bãi nhỏ lẻ tự phong 'trung tâm logistics quốc tế'.",
        "pro_argument": "Tạo thước đo minh bạch chuẩn quốc tế để quy hoạch kết nối đa phương thức, thu hút dòng vốn xanh và nâng vị thế chuỗi cung ứng Việt Nam.",
        "con_argument": "Áp lực chi phí đầu tư hệ thống trạm sạc xe điện, tái chế bao bì và phần mềm quản lý số hóa đồng bộ cho các kho bãi vừa và nhỏ để đạt chuẩn Hạng II–III.",
        "intl_ref": "Hệ thống phân hạng Trung tâm Logistics & Làng hàng hóa (Freight Village / GVZ) tại Đức và mô hình phân tầng Hub logistics của Trung Quốc & Singapore.",
        "url": "https://m.thuvienphapluat.vn/van-ban/Thuong-mai/Thong-tu-71-2026-TT-BCT-tieu-chi-phan-loai-phan-hang-trung-tam-logistics-734832.aspx",
        "keywords": ["71/2026", "logistics"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Thư Viện Pháp Luật (m.thuvienphapluat.vn), Cổng TTĐT Chính phủ (vanban.chinhphu.vn) và Tin Nhanh Chứng Khoán (tinnhanhchungkhoan.vn)."
    },
    {
        "stt": 272,
        "week": "W47 — 21/11/2026",
        "title": "Công Văn Số 24381/CHQ-GSQL (17/11/2026) & 24449/CHQ-GSQL (18/11/2026): Tháo Gỡ Điểm Nghẽn Dừng Cấp C/O Cho Doanh Nghiệp Chế Xuất (DNCX) Bán Hàng Vào Nội Địa",
        "agency": "Cục Hải quan (phối hợp trao đổi cùng Cục Xuất nhập khẩu — Bộ Công Thương)",
        "effective_date": "Công văn 24381/CHQ-GSQL ban hành ngày 17/11/2026; Công văn 24449/CHQ-GSQL ban hành ngày 18/11/2026",
        "core_diff": "Cũ: Doanh nghiệp chế xuất (DNCX), khu phi thuế quan và hàng gửi kho ngoại quan khi xuất bán vào nội địa được cấp C/O (hoặc C/O giáp lưng/cấp sau) để doanh nghiệp nội địa hưởng thuế suất ưu đãi đặc biệt hoặc miễn trừ phòng vệ thương mại; Mới: Sau khi một số thủ tục cấp C/O bị bãi bỏ theo tinh thần cắt giảm TTHC (Nghị quyết 19/2026/NQ-CP), một số tổ chức cấp C/O dừng tiếp nhận hồ sơ từ tháng 11/2026 khiến doanh nghiệp nội địa mua hàng từ DNCX (như trường hợp LG Electronics Việt Nam Hải Phòng) đối mặt nguy cơ chịu thuế MFN cao.",
        "affected_entity": "Các Doanh nghiệp chế xuất (DNCX), doanh nghiệp FDI bán hàng vào nội địa Việt Nam, doanh nghiệp nội địa mua nguyên liệu/linh kiện từ khu phi thuế quan và kho ngoại quan.",
        "risk_pain": "Không có C/O nộp cho cơ quan hải quan khi mở tờ khai nhập khẩu tại chỗ từ DNCX, doanh nghiệp nội địa bị áp thuế suất MFN thông thường thay vì thuế suất FTA 0%, đội chi phí vốn hàng tỷ đồng và phát sinh tranh chấp điều khoản giá hợp đồng.",
        "market_reaction": "Cộng đồng doanh nghiệp FDI và xuất nhập khẩu tại chỗ bày tỏ lo ngại sâu sắc về 'độ vênh' giữa cắt giảm thủ tục cấp C/O và yêu cầu bắt buộc xuất trình C/O tại khâu thông quan hải quan.",
        "pro_argument": "Việc bãi bỏ các thủ tục C/O trung gian giúp giảm gánh nặng giấy tờ hành chính cho cơ quan cấp phép nếu có cơ chế tự chứng nhận xuất xứ hoặc xác nhận điện tử thay thế tương đương.",
        "con_argument": "Dừng cấp C/O đột ngột khi quy định pháp luật về biểu thuế ưu đãi của Bộ Tài chính vẫn yêu cầu nộp C/O đang đẩy rủi ro tài chính và ách tắc thông quan về phía doanh nghiệp.",
        "intl_ref": "Cơ chế Tự chứng nhận xuất xứ (REX / Approved Exporter) của EU và khu phi thuế quan ASEAN cho phép chứng từ thương mại thay thế C/O giấy khi giao dịch qua ranh giới hải quan nội khối.",
        "url": "https://m.thuvienphapluat.vn/cong-van/Xuat-nhap-khau/Cong-van-24449-CHQ-GSQL-2026-vuong-mac-C-O-734756.aspx",
        "keywords": ["24449", "C/O"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Thư Viện Pháp Luật (m.thuvienphapluat.vn) và Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn)."
    },
    {
        "stt": 273,
        "week": "W47 — 21/11/2026",
        "title": "Công Văn Số 24424/CHQ-NVTHQ (Ban Hành 18/11/2026): Hướng Dẫn Miễn Thuế Chống Bán Phá Giá Đối Với Thiết Bị Tạo Tài Sản Cố Định Cho Dự Án Ưu Đãi Đầu Tư",
        "agency": "Cục Hải quan (Ban Nghiệp vụ thuế hải quan)",
        "effective_date": "Ban hành ngày 18/11/2026 (trước thềm kỳ rà soát AD18 tháp điện gió Trung Quốc hạn chót 06/12/2026)",
        "core_diff": "Cũ: Nhiều chủ đầu tư dự án năng lượng tái tạo và nhà thầu lúng túng không rõ thuế chống bán phá giá (như thuế AD18 đối với tháp điện gió xuất xứ Trung Quốc theo QĐ 3453/QĐ-BCT) có được miễn cùng thuế nhập khẩu khi tạo tài sản cố định hay phải nộp riêng; Mới: Công văn 24424/CHQ-NVTHQ khẳng định thuế chống bán phá giá là 'thuế nhập khẩu bổ sung' (Khoản 5 Điều 4 Luật Thuế XNK 107/2016/QH13), nên việc kê khai, miễn thuế được thực hiện như thuế nhập khẩu theo Khoản 11 Điều 16 và Nghị định 134/2016/NĐ-CP nếu thông báo Danh mục miễn thuế đúng Điều 13 & Điều 28 Thông tư 86/2026/TT-BTC.",
        "affected_entity": "Chủ đầu tư các dự án điện gió, năng lượng tái tạo, khu công nghiệp công nghệ cao được hưởng ưu đãi đầu tư và nhà thầu nhập khẩu máy móc thiết bị tạo tài sản cố định.",
        "risk_pain": "Nếu không đăng ký Danh mục hàng hóa miễn thuế dự kiến nhập khẩu trước khi làm thủ tục hải quan hoặc khai sai mã miễn giảm trên VNACCS, doanh nghiệp có thể phải nộp tạm/nộp nhầm hàng chục tỷ đồng thuế chống bán phá giá làm đứt gãy dòng tiền dự án.",
        "market_reaction": "Các nhà đầu tư năng lượng và nhà thầu EPC đánh giá cao hướng dẫn kịp thời của Cục Hải quan giúp tháo gỡ ách tắc tài chính cho các lô thiết bị tháp điện gió siêu trường siêu trọng.",
        "pro_argument": "Thống nhất áp dụng quy chế miễn thuế nhập khẩu bổ sung giúp bảo toàn cam kết ưu đãi đầu tư cho các dự án trọng điểm quốc gia và năng lượng sạch.",
        "con_argument": "Cần giám sát chặt khâu hậu kiểm sử dụng đúng mục đích tài sản cố định để tránh tình trạng lợi dụng miễn thuế nhập thiết bị chịu phòng vệ thương mại rồi chuyển nhượng sai đối tượng.",
        "intl_ref": "Cơ chế miễn trừ thuế phòng vệ thương mại theo mục đích sử dụng cuối cùng (End-use Exemption) tại một số dự án hạ tầng chiến lược của EU và ASEAN.",
        "url": "https://cpahcm.com.vn/cong-van-24424-chq-nvthq-thap-dien-gio-nhap-khau-tao-tai-san-co-dinh-co-duoc-mien-thue-chong-ban-pha-gia.html",
        "keywords": ["24424", "dien gio"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: CPA HCM (cpahcm.com.vn), Kế toán Bảo Lộc (ketoanbaoloc.com) và Thư Viện Pháp Luật."
    },
    {
        "stt": 274,
        "week": "W47 — 21/11/2026",
        "title": "Quy Định Thực Thi (EU) 2026/2497 (Ban Hành 10/11/2026, Hiệu Lực 12/11/2026): EC Áp Thuế Chống Bán Phá Giá Tạm Thời 12,3% – 14,4% Đối Với Ống Đồng Việt Nam",
        "agency": "Ủy ban Châu Âu (EC) & Thương vụ Việt Nam tại EU thông báo",
        "effective_date": "Ban hành ngày 10/11/2026, có hiệu lực từ ngày 12/11/2026 (áp dụng tạm thời tối đa 6 tháng trước phán quyết cuối cùng)",
        "core_diff": "Cũ: Sản phẩm ống đồng tinh luyện dạng cuộn phẳng (level wound coils, mã CN ex 7411 10 90 / TARIC 7411 10 90 10) dùng cho ngành HVACR và năng lượng tái tạo xuất khẩu từ Việt Nam sang EU chỉ chịu thuế nhập khẩu thông thường; Mới: EC áp thuế chống bán phá giá tạm thời từ 12,3% (Hailiang Vietnam), 13,3% (các doanh nghiệp hợp tác như Toàn Phát, LS Metal Vilna), 13,8% (Jintian Copper) đến 14,4% (các doanh nghiệp còn lại), đồng thời áp mức cao hơn đối với Trung Quốc (22,9%–27,6%), Mexico (37,8%) và Uzbekistan (15,2%–18,7%).",
        "affected_entity": "Các doanh nghiệp sản xuất và xuất khẩu ống đồng, linh kiện điều hòa không khí (HVACR), thiết bị tản nhiệt và năng lượng tái tạo sang thị trường 27 nước EU.",
        "risk_pain": "Doanh nghiệp xuất khẩu chưa đăng ký mã TARIC riêng hoặc không chứng minh được nguồn gốc nguyên liệu đồng tinh luyện minh bạch sẽ bị áp mức thuế toàn quốc cao nhất 14,4%, làm mất đơn hàng vào tay đối thủ hoặc phải ký quỹ bảo lãnh ngay tại cảng EU.",
        "market_reaction": "Các doanh nghiệp ống đồng lớn tại Việt Nam tích cực phối hợp cung cấp số liệu bản câu hỏi cho EC để bảo vệ biên độ phá giá thấp nhất trước kỳ điều tra chính thức.",
        "pro_argument": "Mức thuế tạm thời 12,3%–14,4% của Việt Nam vẫn thấp hơn đáng kể so với Trung Quốc (lên tới 27,6%) và Mexico (37,8%), giữ lại khe cửa cạnh tranh nếu doanh nghiệp kiểm soát tốt chi phí sản xuất.",
        "con_argument": "Chuỗi các vụ kiện phòng vệ thương mại dồn dập từ EU (cả thép cán nguội 16% và ống đồng 12,3–14,4% trong tháng 11/2026) đang gia tăng áp lực tuân thủ kế toán giá thành cho ngành công nghiệp chế tạo.",
        "intl_ref": "Quy định thực thi (EU) 2026/2497 điều tra đồng thời 4 nước Việt Nam, Trung Quốc, Mexico và Uzbekistan do Hiệp hội các nhà sản xuất đồng Châu Âu đệ đơn.",
        "url": "https://vinanet.vn/chinh-sach-the-gioi/eu-ap-thue-chong-ban-pha-gia-tam-thoi-doi-voi-san-pham-ong-dong-nhap-khau-tu-viet-nam-trung-quoc-mexico-va-uzbekistan-817390.html",
        "keywords": ["2026/2497", "12,3%"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Vinanet (vinanet.vn), Doanh nghiệp & Cuộc sống (doanhnghiepcuocsong.vn) và Cục Phòng vệ thương mại (Bộ Công Thương)."
    },
    {
        "stt": 275,
        "week": "W47 — 21/11/2026",
        "title": "Quyết Định Tự Vệ Nhựa Polypropylene (PP) Của Thái Lan (12/11/2026): Áp Thuế Tự Vệ 3 Năm Mức 7,73% CIF Từ Ngày 28/12/2026 Đối Với 40 Quốc Gia Gồm Việt Nam",
        "agency": "Cục Ngoại thương Thái Lan (DFT) & Bộ Công Thương Việt Nam thông báo",
        "effective_date": "Ban hành ngày 12/11/2026; thời hạn áp dụng 3 năm (28/09/2026 – 27/09/2029), bắt đầu thu thuế thực tế từ ngày 28/12/2026",
        "core_diff": "Cũ: Hạt nhựa nguyên sinh Polypropylene (PP) xuất khẩu từ Việt Nam sang Thái Lan được hưởng thuế suất ưu đãi nội khối ASEAN (ATIGA 0%); Mới: Thái Lan quyết định áp thuế tự vệ chính thức (Safeguard) trong 3 năm đối với sản phẩm nhựa PP nhập khẩu từ 40 quốc gia/vùng lãnh thổ (không loại trừ Việt Nam), mức năm đầu là 7,73% tính trên giá CIF và bắt đầu thực thu từ ngày 28/12/2026.",
        "affected_entity": "Các nhà máy lọc hóa dầu, sản xuất hạt nhựa PP, doanh nghiệp xuất khẩu nguyên liệu nhựa và bao bì nhựa sang thị trường Thái Lan.",
        "risk_pain": "Thuế tự vệ 7,73% CIF đánh thẳng vào lợi thế thuế ATIGA 0%, khiến hợp đồng xuất khẩu hạt nhựa PP sang Thái Lan giao sau mốc 28/12/2026 mất khả năng cạnh tranh giá nếu không đàm phán lại điều khoản Incoterms hoặc chuyển hướng thị trường.",
        "market_reaction": "Cục Phòng vệ thương mại và Thương vụ Việt Nam tại Thái Lan cảnh báo doanh nghiệp hóa dầu sớm tính toán lại phương án giao hàng trước mốc 28/12/2026 và rà soát danh mục mã HS nhóm 3902 bị áp thuế.",
        "pro_argument": "Việc Thái Lan lùi thời điểm bắt đầu thu thuế đến 28/12/2026 giúp doanh nghiệp hai nước có khoảng đệm hơn 1 tháng để giải phóng các lô hàng đã ký hợp đồng.",
        "con_argument": "Biện pháp tự vệ nội khối ASEAN cho thấy xu hướng bảo hộ ngành hóa dầu khu vực đang gia tăng mạnh trước áp lực dư cung công suất toàn cầu.",
        "intl_ref": "Hiệp định về Biện pháp Tự vệ của WTO (Agreement on Safeguards) cho phép nước nhập khẩu áp thuế tạm thời/chính thức tối đa 4 năm khi lượng nhập khẩu tăng đột biến gây thiệt hại nghiêm trọng.",
        "url": "https://moit.gov.vn/tin-tuc/thai-lan-quyet-dinh-ap-thue-tu-ve-chinh-thuc-doi-voi-san-pham-nhua-polypropylen-nhap-khau.html",
        "keywords": ["polypropylen", "7,73"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Bộ Công Thương (moit.gov.vn) và Cục Phòng vệ thương mại."
    },
]


# ==============================================================================
# 2. DATA FOR CASE STUDY (14 cols) — STT 271..275 (W47 — 21/11/2026)
# ==============================================================================
CS_NEW_IDEAS = [
    {
        "stt": 271,
        "week": "W47 — 21/11/2026",
        "title": "Đại Án Xuất Khẩu Trái Phép Hơn 28.000 Tấn Quặng Titan (1.033 Container) Tại Cảng Sài Gòn KV1 (Xét Xử 07/2026): Đánh Tráo Mẫu Giám Định & Bản Án 12 Năm Tù",
        "parties": "28 bị cáo tại Công ty CP Tập đoàn Khoáng sản Đầu tư Hưng Thịnh + Thương nhân nước ngoài + Cựu công chức kiểm hóa Chi cục Hải quan Cảng Sài Gòn KV1 (TAND TP.HCM xét xử sơ thẩm 07/2026)",
        "plot": "Trong vụ án khai thác trái phép hơn 459.000 tấn quặng trị giá hơn 1.508 tỷ đồng, nhóm doanh nghiệp thu mua quặng Ilmenite thô chỉ đạt hàm lượng TiO2 51–52% (không đủ điều kiện xuất khẩu tối thiểu 56%). Để thông quan trót lọt 1.033 container chứa gần 28.000 tấn quặng sang Trung Quốc, doanh nghiệp đã móc nối với cán bộ kiểm hóa Hải quan Cảng Sài Gòn KV1 chi khoảng 8,6 tỷ đồng để đánh tráo mẫu quặng đạt chuẩn đi giám định thay cho hàng thực tế trong container.",
        "damage": "Tổng trị giá quặng vi phạm trong đại án hơn 1.508 tỷ đồng; riêng cựu công chức hải quan lĩnh án 12 năm tù về tội nhận hối lộ, các chủ mưu doanh nghiệp nhận mức án 10–15 năm tù và bị truy thu toàn bộ khoản thu lợi bất chính.",
        "root_cause": "Cố tình xuất khẩu khoáng sản thô chưa đạt hàm lượng TiO2 tối thiểu theo quy định pháp luật và dùng tiền hối lộ để gian lận quy trình lấy mẫu giám định thực tế tại cảng.",
        "prevention": "1. Chỉ xuất khẩu khoáng sản đáp ứng đúng tiêu chuẩn kỹ thuật, hàm lượng chế biến sâu và giấy phép theo danh mục của Bộ Công Thương.\n2. Quy trình lấy mẫu kiểm hóa, giám định phải có biên bản niêm phong mẫu đối chứng, chữ ký đa bên (chủ hàng, đơn vị giám định độc lập và hải quan) kèm hình ảnh định vị.\n3. Thiết lập cơ chế kiểm soát nội bộ và từ chối tuyệt đối mọi đề nghị 'tác động hồ sơ giám định' từ đối tác hoặc môi giới.",
        "opposing_views": "Phe A: Quặng chênh lệch vài phần trăm hàm lượng thì nhờ 'linh hoạt lấy mẫu' cho kịp lịch tàu xuất khẩu vs Phe B: Đánh tráo mẫu giám định là hành vi cấu thành tội phạm hình sự kéo cả doanh nghiệp và cán bộ kiểm hóa vào vòng tù tội 10–15 năm.",
        "lessons": "Trong xuất nhập khẩu hàng khoáng sản và hàng kiểm tra chuyên ngành, mẫu giám định là bằng chứng pháp lý tối thượng. Mọi sự 'linh hoạt' sai quy chuẩn đều có giá trả bằng tự do và toàn bộ sự nghiệp.",
        "url": "https://www.sggp.org.vn/xet-xu-28-bi-cao-trong-vu-khai-thac-trai-phep-titan-tri-gia-hon-1500-ty-dong-post860104.html",
        "keywords": ["titan", "1.500"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Báo Sài Gòn Giải Phóng (sggp.org.vn), Tuổi Trẻ (tuoitre.vn) và VnExpress (vnexpress.net)."
    },
    {
        "stt": 272,
        "week": "W47 — 21/11/2026",
        "title": "Bắt Giữ 24.930 Phụ Kiện 'Đội Lốt' Nhãn Hiệu Harley-Davidson Trong Lô Hàng Quá Cảnh Tại Cửa Khẩu Tà Lùng — Cao Bằng (05/2026)",
        "parties": "Chủ lô hàng quá cảnh qua Cửa khẩu Quốc tế Tà Lùng + Chi cục Hải quan cửa khẩu Tà Lùng & Đội Kiểm soát chống buôn lậu số 1 (Hải quan khu vực XVI)",
        "plot": "Tháng 05/2026, khi kiểm tra thực tế một lô hàng đăng ký loại hình quá cảnh qua Cửa khẩu Quốc tế Tà Lùng (Cao Bằng), lực lượng Hải quan phát hiện số lượng lớn hàng hóa không khai báo trên tờ khai và có dấu hiệu xâm phạm quyền sở hữu trí tuệ nhãn hiệu mô tô nổi tiếng Harley-Davidson, được đóng ghép lẫn trong lô hàng quá cảnh.",
        "damage": "Toàn bộ lô hàng gồm 24.930 sản phẩm phụ kiện bằng vải in/thêu logo Harley-Davidson (12.140 khăn vải, 11.830 mũ và 960 khẩu trang) bị tạm giữ khẩn cấp để xác minh xâm phạm quyền sở hữu trí tuệ, kéo theo đình trệ cả lô hàng quá cảnh và xử phạt theo Thông tư 06/2026/TT-BTC.",
        "root_cause": "Nhiều chủ hàng và đơn vị vận tải ngộ nhận rằng 'hàng quá cảnh không nhập vào nội địa Việt Nam thì không bị kiểm tra sở hữu trí tuệ', dẫn đến chủ quan không rà soát nhãn hiệu hàng hóa trước khi nhận ủy thác vận chuyển.",
        "prevention": "1. Theo Thông tư 06/2026/TT-BTC (hiệu lực 01/03/2026), Hải quan có quyền chủ động tạm dừng thông quan đối với CẢ hàng quá cảnh nếu nghi ngờ xâm phạm quyền SHTT.\n2. Đơn vị logistics nhận vận chuyển hàng quá cảnh bắt buộc yêu cầu khách hàng cung cấp Thư ủy quyền sử dụng nhãn hiệu (Letter of Authorization) hoặc giấy tờ chứng minh nguồn gốc chính hãng.\n3. Kiểm đếm đối chiếu 100% số lượng thực tế với Packing List để không bị gài thêm hàng không khai báo vào container quá cảnh.",
        "opposing_views": "Phe A: Hàng quá cảnh chỉ đi nhờ đường sang nước thứ ba nên cơ quan chức năng Việt Nam không cần soi kỹ nhãn hiệu vs Phe B: Theo luật SHTT và Thông tư 06/2026/TT-BTC, hàng quá cảnh vẫn thuộc diện kiểm soát SHTT nghiêm ngặt tại cửa khẩu, vi phạm là bị chặn giữ ngay.",
        "lessons": "Đừng bao giờ tin lời môi giới rằng 'hàng quá cảnh đi qua Việt Nam không kiểm nhãn hiệu'. Một lô phụ kiện 24.930 chiếc vi phạm SHTT đủ khiến cả doanh nghiệp vận tải mất uy tín và kẹt xe tại biên giới.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/phat-hien-hang-chuc-nghin-san-pham-nghi-van-vi-pham-so-huu-tri-tue-156697.html",
        "keywords": ["Harley", "24.930"],
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn) và Văn bản Thông tư 06/2026/TT-BTC."
    },
    {
        "stt": 273,
        "week": "W47 — 21/11/2026",
        "title": "Khởi Tố Vụ Án Vận Chuyển Trái Phép Tinh Dầu & Linh Kiện Thuốc Lá Điện Tử Trị Giá 2,48 Tỷ Đồng Tại Cửa Khẩu Hữu Nghị (QĐ 01/QĐ-HQHN, 05/2026)",
        "parties": "Đối tượng vận chuyển hàng qua biên giới + Hải quan Cửa khẩu Quốc tế Hữu Nghị (Chi cục Hải quan khu vực VI, Lạng Sơn)",
        "plot": "Ngày 20/05/2026, qua công tác soi chiếu và kiểm tra thực tế từ khâu làm thủ tục hải quan tại Cửa khẩu Quốc tế Hữu Nghị, lực lượng Hải quan khu vực VI phát hiện các đối tượng cất giấu, trà trộn hàng loạt bộ phận chứa tinh dầu và tinh dầu thuốc lá điện tử (mặt hàng thuộc diện cấm kinh doanh/lưu hành tại Việt Nam) không khai báo vào bên trong các lô hàng khai báo là hàng hóa thông thường. Hải quan đã ban hành Quyết định số 01/QĐ-HQHN khởi tố vụ án hình sự.",
        "damage": "Toàn bộ tang vật trị giá 2,48 tỷ đồng bị tịch thu và vụ án bị khởi tố hình sự về tội vận chuyển trái phép hàng hóa qua biên giới (nằm trong tổng số 1.590 vụ vi phạm trị giá gần 146 tỷ đồng tại Hải quan khu vực VI tính đến giữa tháng 11/2026).",
        "root_cause": "Thủ đoạn lợi dụng luồng hàng ghép/hàng tiêu dùng thông thường để trà trộn hàng cấm (thuốc lá điện tử) nhằm qua mặt kiểm tra cửa khẩu đường bộ.",
        "prevention": "1. Doanh nghiệp logistics và đại lý hải quan tuyệt đối từ chối vận chuyển ghép kiện khi không được mở kiểm tra thực tế ruột kiện hàng tại kho tập kết.\n2. Cập nhật thường xuyên danh mục hàng cấm nhập khẩu (bao gồm thuốc lá điện tử, khí cười N2O, hàng công nghệ cũ theo TT 63/2026/TT-BKHCN) để cảnh báo cho bộ phận Sales.\n3. Ký hợp đồng dịch vụ vận chuyển có điều khoản miễn trừ trách nhiệm rõ ràng nếu chủ hàng cố tình giấu hàng cấm trong kiện hàng niêm phong.",
        "opposing_views": "Phe A: Đơn vị giao nhận chỉ chở theo tờ khai khách đưa, không biết bên trong thùng hàng có giấu tinh dầu thuốc lá điện tử vs Phe B: Khi tang vật trị giá 2,48 tỷ đồng bị phát hiện trên xe qua cửa khẩu, cả chủ hàng lẫn đơn vị vận tải đều phải giải trình liên đới hình sự.",
        "lessons": "Nhận chở hàng ghép không rõ nguồn gốc mà không kiểm soát nội dung bên trong là đánh cược sinh mạng pháp lý của doanh nghiệp logistics ngay tại cổng máy soi cửa khẩu.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/hai-quan-khu-vuc-vi-ngan-buon-lau-gian-lan-thuong-mai-tu-khau-lam-thu-tuc-173104.html",
        "keywords": ["2,48", "Huu Nghi"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn) và Thương hiệu & Công luận (thuonghieucongluan.com.vn)."
    },
    {
        "stt": 274,
        "week": "W47 — 21/11/2026",
        "title": "Truy Tố Đường Dây Cắt Seal Hải Quan Nhập Sợi Polyester Né Thuế CBPG 17,45% Tại Chi Cục Hải Quan Chơn Thành (Bình Phước, Truy Tố 08/2025)",
        "parties": "Bạch Tấn Cường (Phó GĐ Công ty XNK Tân Phong Vận) cùng 11 bị can (trong đó có 5 cựu công chức Chi cục Hải quan Chơn Thành) + Công an & VKSND TP.HCM",
        "plot": "Để né thuế chống bán phá giá 17,45% đối với sợi Polyester từ Trung Quốc, Bạch Tấn Cường lập các công ty vỏ bọc mở tờ khai luồng Đỏ chuyển cửa khẩu về Chi cục Hải quan Chơn Thành, khai sai tên hàng thành loại sợi thuế suất 0%. Khi container được niêm phong kẹp seal từ cảng TP.HCM vận chuyển đi kiểm hóa, các đối tượng chở thẳng về kho ở TP.HCM tự ý cắt seal hải quan rút ruột hàng đưa đi tiêu thụ và chi hối lộ 3,5 triệu đồng/container (tổng cộng 266 triệu đồng) để 5 công chức hải quan bỏ qua kiểm hóa thực tế.",
        "damage": "Lô hàng tang vật bị bắt quả tang khi cắt seal tại Cảng Phú Định ước tính gần 150 tỷ đồng (riêng 1 đợt hàng bị truy tố gần 30 tỷ đồng); 12 bị can bị truy tố các tội danh buôn lậu và nhận hối lộ vào tháng 08/2025.",
        "root_cause": "Cấu kết lập doanh nghiệp vỏ bọc khai sai chủng loại hàng chịu thuế phòng vệ thương mại (17,45%) và tự ý phá niêm phong hải quan (seal) trên đường vận chuyển luồng Đỏ.",
        "prevention": "1. Tuyệt đối tuân thủ nguyên tắc bất khả xâm phạm niêm phong hải quan (seal) khi hàng đang vận chuyển chuyển cửa khẩu hoặc chờ kiểm hóa.\n2. Khi nhập khẩu các mặt hàng dệt may, sợi tổng hợp, thép nằm trong danh mục điều tra phòng vệ thương mại, phải khai đúng mã HS và thành phần kỹ thuật, kiểm tra chứng thư giám định.\n3. Kiểm soát chặt quy trình bàn giao GPS và hành trình xe container từ cảng về địa điểm kiểm tra ngoài cửa khẩu.",
        "opposing_views": "Phe A: Thuế chống bán phá giá 17,45% quá cao nên doanh nghiệp tìm cách 'đổi tên sợi' để hạ giá thành vs Phe B: Khai sai chủng loại hàng hóa kèm hành vi tự ý cắt seal hải quan là tội danh hình sự nghiêm trọng khiến toàn bộ bộ máy vận hành sụp đổ.",
        "lessons": "Thuế phòng vệ thương mại 17,45% có thể làm biên lợi nhuận mỏng đi, nhưng cắt seal hải quan và khai sai tên hàng sẽ xóa sổ hoàn toàn doanh nghiệp và đẩy người đứng đầu vào vòng lao lý.",
        "url": "https://tuoitre.vn/truy-to-5-cuu-cong-chuc-hai-quan-chon-thanh-nhan-hoi-lo-266-trieu-dong-de-tiep-tay-cho-nhom-buon-lau-20250803094146063.htm",
        "keywords": ["Chon Thanh", "266"],
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 3+ nguồn: Báo Tuổi Trẻ (tuoitre.vn), VTV (vtv.vn) và VietnamPlus (vietnamplus.vn)."
    },
    {
        "stt": 275,
        "week": "W47 — 21/11/2026",
        "title": "Nộp Thông Báo Gia Công Lại (Mẫu 18A) Suốt 3–4 Năm Không Lấy Số Tiếp Nhận Văn Thư: Doanh Nghiệp SXXK Đối Mặt Nguy Cơ Phạt & Bác Miễn Thuế Khi Hậu Kiểm",
        "parties": "Doanh nghiệp Sản xuất xuất khẩu (SXXK) thuê gia công lại (ghi nhận thực tế từ Daily Notes 20/11/2026) + Chi cục Hải quan quản lý",
        "plot": "Một doanh nghiệp sản xuất xuất khẩu thuê đối tác bên ngoài gia công lại công đoạn sản xuất suốt 3–4 năm. Mỗi lần đi nộp Thông báo hợp đồng gia công lại (Mẫu 18a/TB-HĐGCL/GSQL theo Phụ lục V Thông tư 39/2018/TT-BTC), nhân viên chỉ đưa hồ sơ giấy trực tiếp mà không qua bộ phận văn thư lấy số tiếp nhận hoặc không lưu biên nhận điện tử. Đến kỳ kiểm tra sau thông quan, cơ quan hải quan đối soát không thấy dữ liệu tiếp nhận đã thông báo trước khi giao nguyên liệu, đặt doanh nghiệp vào nguy cơ bị phạt và xem xét lại chính sách miễn thuế.",
        "damage": "Mức phạt hành chính cho hành vi thông báo hợp đồng gia công lại không đúng thời hạn là 2.000.000 – 5.000.000 đồng/lần (Điểm e khoản 3 Điều 7 Nghị định 128/2020/NĐ-CP & khung phạt Nghị định 169/2026/NĐ-CP); rủi ro lớn nhất là nguy cơ bị bác miễn thuế nhập khẩu và truy thu toàn bộ tiền thuế nguyên liệu nhiều năm kèm tiền chậm nộp 0,03%/ngày.",
        "root_cause": "Thói quen làm thủ tục hành chính 'giao tay không biên nhận', bỏ qua bước lấy số tiếp nhận văn thư hoặc nộp qua hệ thống điện tử theo quy định trước thời điểm xuất kho nguyên liệu đi gia công lại.",
        "prevention": "1. Bắt buộc nộp Mẫu 18a/TB-HĐGCL/GSQL qua hệ thống hải quan điện tử (theo Công văn 17658/CHQ-GSQL) hoặc lấy đầy đủ số tiếp nhận, dấu ngày của bộ phận văn thư trước khi giao nguyên liệu.\n2. Khi bị hậu kiểm thiếu số tiếp nhận cũ, lập ngay bộ hồ sơ chứng minh tính chân thật của chuỗi sản xuất (hợp đồng gia công lại, phiếu xuất-nhập kho nguyên liệu/thành phẩm, chứng từ thanh toán qua ngân hàng, vận chuyển thực tế, báo cáo quyết toán Mẫu 15) để giải trình xin giữ quyền miễn thuế.\n3. Kiểm tra định kỳ hồ sơ pháp lý gia công 6 tháng/lần trước khi làm báo cáo quyết toán năm.",
        "opposing_views": "Phe A: Đã mất công mang giấy lên nộp cho cán bộ phụ trách rồi thì coi như xong trách nhiệm, mất số tiếp nhận là lỗi nhỏ vs Phe B: Trong kiểm tra sau thông quan, chỉ có số tiếp nhận hệ thống hoặc dấu văn thư mới là bằng chứng pháp lý chứng minh đã khai báo trước khi xuất nguyên liệu.",
        "lessons": "Trong nghiệp vụ SXXK và gia công, 'đã làm' nhưng 'không có dấu tiếp nhận' đồng nghĩa với 'chưa chứng minh được'. Hãy xây dựng hàng rào chứng từ thực tế song hành cùng số tiếp nhận điện tử để bảo vệ khoản miễn thuế hàng tỷ đồng.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/xu-phat-hanh-chinh-neu-khong-thong-bao-hop-dong-gia-cong-lai-dung-thoi-han-87964.html",
        "keywords": ["gia cong lai", "128/2020"],
        "series": "Series: Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn thực tế từ scratch/daily_notes.md (20/11/2026). Đối chiếu 3+ nguồn pháp lý: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn), LuatVietnam (luatvietnam.vn) và Thư Viện Pháp Luật."
    },
]


# ==============================================================================
# 3. DATA FOR HOC THUC CHIEN (14 cols) — STT 271..275 (W47 — 21/11/2026)
# ==============================================================================
HT_NEW_IDEAS = [
    {
        "stt": 271,
        "week": "W47 — 21/11/2026",
        "title": "Quy Trình 5 Bước Làm Sạch & Cập Nhật Tài Khoản VNACCS/VCIS Tại dksd.customs.gov.vn Trước Hạn Chót 31/12/2026 Theo Công Văn 23474/CHQ-CNTT",
        "steps": "Bước 1: Đăng nhập Hệ thống Đăng ký và quản lý người sử dụng tại dksd.customs.gov.vn bằng chữ ký số (Token) còn hiệu lực của doanh nghiệp.\nBước 2: Kiểm tra lại toàn bộ thông tin định danh pháp nhân (tên doanh nghiệp, mã số thuế, địa chỉ sau sáp nhập địa giới hành chính, người đại diện pháp luật, email/SĐT nhận OTP).\nBước 3: Rà soát danh sách User ID khai báo hải quan VNACCS/VCIS đang hoạt động; thu hồi hoặc khóa ngay các tài khoản của nhân sự chứng từ đã nghỉ việc.\nBước 4: Cập nhật bổ sung thông tin chuẩn xác và thực hiện ký số xác nhận gửi Chi cục Hải quan quản lý phê duyệt.\nBước 5: Lưu lại biên nhận phê duyệt điện tử và kiểm tra thử truyền tờ khai nháp trước ngày 31/12/2026 để tránh bị khóa tài khoản tự động từ 01/01/2027.",
        "formula": "Mốc khóa tài khoản tự động: 01/01/2027 (Hạn chót hoàn thành: 31/12/2026). Cổng thao tác chính thức: https://dksd.customs.gov.vn | Căn cứ: Công văn 23474/CHQ-CNTT (26/10/2026) & 21692/CHQ-CNTT.",
        "top3_errors": "1. Để nguyên thông tin địa chỉ hành chính cũ hoặc đại diện pháp luật cũ khiến hệ thống đối soát không khớp dữ liệu thuế.\n2. Quên xóa User ID của nhân viên chứng từ cũ đã nghỉ việc, gây rủi ro bảo mật hoặc vượt định mức tài khoản.\n3. Để sát tuần cuối tháng 12/2026 mới thao tác khiến hệ thống nghẽn mạng hoặc chờ cán bộ hải quan duyệt quá hạn, dẫn đến bị khóa tài khoản ngay đầu năm mới.",
        "mnemonic": "Ví tài khoản VNACCS như 'thẻ căn cước thông quan điện tử': Quá hạn 31/12 mà chưa định danh lại là cổng cảng đóng ngay đầu năm 2027.",
        "comparison": "Tài khoản VNACCS đã làm sạch dữ liệu (Truyền tờ khai thông suốt 24/7 từ 01/01/2027) vs Tài khoản chưa làm sạch (Bị hệ thống tự động khóa quyền truy cập, hàng cập cảng không mở được tờ khai, phát sinh DEM/DET).",
        "persona": "Chị Hương (Trưởng phòng Chứng từ XNK) & Nhân viên khai báo hải quan tại doanh nghiệp XNK/đại lý hải quan.",
        "checklist": "☐ Kiểm tra Token chữ ký số còn hạn\n☐ Đăng nhập dksd.customs.gov.vn đối soát địa chỉ & người đại diện\n☐ Rà soát & xóa User nhân sự đã nghỉ việc\n☐ Ký số gửi cập nhật & theo dõi trạng thái 'Đã phê duyệt' trước 31/12/2026",
        "url": "https://luatvietnam.vn/doanh-nghiep/cong-van-23474-chq-cntt-2026-cua-cuc-hai-quan-ve-trien-khai-lam-sach-du-lieu-doanh-nghiep-452790-d6.html",
        "keywords": ["23474", "VNACCS"],
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Từ khóa SEO Top 8 Sheet 1. Đối chiếu 3+ nguồn: LuatVietnam (luatvietnam.vn), BKHUB (bkhub.vn) và Kế toán Bảo Lộc (ketoanbaoloc.com)."
    },
    {
        "stt": 272,
        "week": "W47 — 21/11/2026",
        "title": "Cẩm Nang 4 Bước Xử Lý Vướng Mắc Dừng Cấp C/O Cho Hàng Hóa Từ DNCX Bán Vào Nội Địa Theo Công Văn 24381/CHQ-GSQL & 24449/CHQ-GSQL",
        "steps": "Bước 1: Xác định xem lô hàng mua từ Doanh nghiệp chế xuất (DNCX), khu phi thuế quan hoặc gửi kho ngoại quan có thuộc diện áp thuế suất ưu đãi đặc biệt (FTA) hay hàng không chịu thuế phòng vệ thương mại khi nhập vào nội địa hay không.\nBước 2: Kiểm tra phản hồi bằng văn bản của tổ chức cấp C/O tại địa phương về việc tạm dừng cấp C/O theo danh mục cắt giảm TTHC (Nghị quyết 19/2026/NQ-CP).\nBước 3: Gửi công văn đề nghị hướng dẫn trực tiếp tới Cục Xuất nhập khẩu (Bộ Công Thương) theo chỉ dẫn tại Công văn 24449/CHQ-GSQL (18/11/2026), đồng thời rà soát khả năng áp dụng cơ chế tự chứng nhận xuất xứ (nếu đáp ứng điều kiện FTA).\nBước 4: Đàm phán bổ sung phụ lục hợp đồng mua bán giữa DNCX và doanh nghiệp nội địa về phương án xử lý phần thuế MFN chênh lệch hoặc bảo lưu quyền khai bổ sung hoàn thuế khi có hướng dẫn tháo gỡ liên bộ.",
        "formula": "Chênh lệch thuế khi thiếu C/O = Trị giá hải quan × (Thuế suất MFN − Thuế suất FTA ưu đãi đặc biệt) × (1 + Thuế GTGT). Căn cứ: CV 24381/CHQ-GSQL (17/11/2026) & CV 24449/CHQ-GSQL (18/11/2026).",
        "top3_errors": "1. Mặc định đinh ninh hàng sản xuất tại DNCX Việt Nam bán vào nội địa luôn được cấp C/O mẫu D/E như trước nên báo giá bán cố định không dự phòng thuế MFN.\n2. Tự ý nộp hồ sơ thiếu C/O để thông quan trước mà không ghi chú bảo lưu hoặc kiểm tra quy định thời hạn nộp bổ sung C/O (nếu có).\n3. Chỉ hỏi cơ quan hải quan mà không gửi văn bản sang cơ quan có thẩm quyền cấp C/O thuộc Bộ Công Thương.",
        "mnemonic": "Ví C/O hàng DNCX bán nội địa như 'giấy thông hành giảm thuế': Khi cơ quan cấp tạm dừng in vé, phải kiểm tra ngay vé điện tử thay thế trước khi xe qua trạm.",
        "comparison": "Mua hàng DNCX có C/O hợp lệ / Tự chứng nhận xuất xứ (Thuế nhập khẩu FTA 0–5%) vs Thiếu C/O khi mở tờ khai tại chỗ (Áp thuế MFN 10–25%, đội giá vốn đột biến).",
        "persona": "Anh Minh (Giám đốc Doanh nghiệp nội địa mua linh kiện từ DNCX) & Phòng Xuất nhập khẩu doanh nghiệp FDI.",
        "checklist": "☐ Xác nhận tình trạng cấp C/O tại phòng quản lý XNK khu vực trước khi ký hợp đồng\n☐ Kiểm tra khả năng áp dụng Tự chứng nhận xuất xứ theo FTA\n☐ Rà soát điều khoản phân chia rủi ro thuế MFN trong hợp đồng với DNCX\n☐ Theo dõi công văn phúc đáp của Cục Xuất nhập khẩu (Bộ Công Thương)",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/can-thong-nhat-huong-xu-ly-khi-dung-cap-c-o-173080.html",
        "keywords": ["24381", "C/O"],
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Từ khóa SEO Top 1 Sheet 1 & Daily Notes 20/11/2026. Đối chiếu 2+ nguồn: Tạp chí Kinh tế - Tài chính (thuehaiquan.tapchikinhtetaichinh.vn) và Thư Viện Pháp Luật."
    },
    {
        "stt": 273,
        "week": "W47 — 21/11/2026",
        "title": "Hướng Dẫn 5 Bước Lập Hồ Sơ Miễn Thuế Chống Bán Phá Giá (AD18 Tháp Điện Gió) Tạo Tài Sản Cố Định Theo Công Văn 24424/CHQ-NVTHQ",
        "steps": "Bước 1: Xác minh dự án thuộc đối tượng hưởng ưu đãi đầu tư (Giấy chứng nhận đăng ký đầu tư) và đối chiếu mặt hàng (như tháp điện gió Trung Quốc chịu thuế CBPG theo QĐ 3453/QĐ-BCT) có thuộc danh mục máy móc, thiết bị tạo tài sản cố định trong nước chưa sản xuất được.\nBước 2: Lập và thông báo Danh mục hàng hóa miễn thuế dự kiến nhập khẩu lên hệ thống hải quan đúng quy định tại Điều 13 Thông tư 86/2026/TT-BTC trước khi hàng về cảng.\nBước 3: Chuẩn bị bộ hồ sơ miễn thuế theo Điều 28 Thông tư 86/2026/TT-BTC, viện dẫn rõ Khoản 5 Điều 4 và Khoản 11 Điều 16 Luật Thuế XNK 107/2016/QH13 cùng Công văn 24424/CHQ-NVTHQ (thuế CBPG là thuế nhập khẩu bổ sung nên được áp dụng cơ chế miễn thuế tương tự thuế nhập khẩu).\nBước 4: Khai báo đúng mã miễn giảm thuế trên tờ khai VNACCS (lưu ý tách biệt rõ với hàng tiêu dùng/nguyên liệu sản xuất không thuộc TSCĐ).\nBước 5: Lưu trữ đầy đủ chứng từ lắp đặt, bàn giao, hạch toán tăng Tài sản cố định và theo dõi báo cáo tình hình sử dụng hàng miễn thuế hàng năm (đồng thời chuẩn bị dữ liệu nếu tham gia rà soát AD18 trước hạn 06/12/2026).",
        "formula": "Thuế CBPG được miễn = Trị giá tính thuế CIF × Thuế suất CBPG áp dụng (ví dụ theo QĐ 3453/QĐ-BCT). Căn cứ: Khoản 5 Điều 4 & Khoản 11 Điều 16 Luật 107/2016/QH13; Điều 13, 28 TT 86/2026/TT-BTC; CV 24424/CHQ-NVTHQ.",
        "top3_errors": "1. Nhập khẩu hàng về cảng rồi mới làm thủ tục thông báo Danh mục hàng hóa miễn thuế dự kiến.\n2. Nghĩ rằng thuế chống bán phá giá là thuế riêng biệt không được miễn cho tài sản cố định nên nộp oan hàng tỷ đồng hoặc bỏ lỡ ưu đãi.\n3. Đưa lẫn vật tư tiêu hao, phụ tùng thay thế không đủ tiêu chuẩn tài sản cố định vào danh mục xin miễn thuế CBPG.",
        "mnemonic": "Nhớ quy tắc 'Thuế bổ sung đi cùng thuế gốc': Đã được miễn thuế nhập khẩu tạo TSCĐ hợp lệ thì thuế CBPG cũng được miễn theo đúng Khoản 5 Điều 4 Luật 107/2016.",
        "comparison": "Khai báo Danh mục miễn thuế TSCĐ trước khi nhập (Miễn cả thuế nhập khẩu & thuế CBPG) vs Không đăng ký danh mục miễn thuế (Nộp đủ thuế CBPG ngay tại cảng, tắc dòng tiền dự án).",
        "persona": "Anh Dũng (Giám đốc Dự án / Chủ đầu tư điện gió, năng lượng tái tạo) & Kế toán trưởng XNK.",
        "checklist": "☐ Kiểm tra Giấy chứng nhận ưu đãi đầu tư & tiêu chí hàng trong nước chưa sản xuất được\n☐ Thông báo Danh mục miễn thuế trước khi mở tờ khai theo Điều 13 TT 86/2026\n☐ Đối chiếu mã miễn thuế trên VNACCS cho thuế NK bổ sung (CBPG)\n☐ Hạch toán TSCĐ và nộp báo cáo sử dụng hàng miễn thuế định kỳ",
        "url": "https://ketoanbaoloc.com/tin-tuc-ke-toan/24424-chq-nvthq-v-v-chinh-sach-thue-nk-bo-sung-cong-ty-co-phan-dau-tu-xay-dung-truong-son",
        "keywords": ["24424", "Truong Son"],
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Từ khóa SEO Top 7 Sheet 1. Đối chiếu 3+ nguồn: Kế toán Bảo Lộc (ketoanbaoloc.com), CPA HCM (cpahcm.com.vn) và Thư Viện Pháp Luật."
    },
    {
        "stt": 274,
        "week": "W47 — 21/11/2026",
        "title": "Quy Trình 5 Bước Tự Đánh Giá Tiêu Chí & Công Bố Phân Hạng Trung Tâm Logistics (Hạng I, II, III) Trên TradeNode.vn Theo Thông Tư 71/2026/TT-BCT",
        "steps": "Bước 1: Xác định đúng 1 trong 5 loại hình trung tâm logistics của doanh nghiệp (Cảng biển/Hàng không, Ngành hàng, KCX/FTZ, Biên giới/Cửa khẩu hoặc Thương mại điện tử).\nBước 2: Tự chấm điểm quy mô diện tích & công suất xử lý (Hạng I >= 20 ha & >= 3 triệu tấn/200.000 TEU/35 triệu đơn hàng/năm; Hạng II 10–<20 ha & 1–<3 triệu tấn; Hạng III 3–<10 ha).\nBước 3: Rà soát hệ thống kết nối giao thông (tối thiểu 2 phương thức vận tải đối với Hạng I, bãi đỗ container >= 200 vị trí) và điều kiện kỹ thuật số (phần mềm quản lý kho WMS, camera 24/7, kiểm soát thẻ thông minh).\nBước 4: Hoàn thiện tiêu chuẩn hạ tầng xanh bắt buộc (xử lý nước thải, quy trình thu gom tái chế bao bì gỗ/carton, trạm sạc năng lượng sạch cho xe vận tải).\nBước 5: Chuẩn bị hồ sơ tự công bố phân hạng và cập nhật dữ liệu lên Cơ sở dữ liệu hạ tầng thương mại quốc gia (TradeNode.vn) khi Thông tư có hiệu lực từ 08/01/2027.",
        "formula": "Ngưỡng chuẩn 3 Hạng: Hạng I (>= 20 ha; >= 3 triệu tấn hoặc 200.000 TEU hoặc 35 triệu đơn/năm) | Hạng II (10–<20 ha; 1–<3 triệu tấn hoặc 70.000–<200.000 TEU) | Hạng III (3–<10 ha). Căn cứ: TT 71/2026/TT-BCT.",
        "top3_errors": "1. Chỉ tính diện tích đất kho mà quên tiêu chí bãi đỗ xe container tối thiểu và kết nối đa phương thức.\n2. Chưa trang bị hệ thống trạm sạc xe điện/năng lượng sạch và quy trình tái chế bao bì theo chuẩn phát triển xanh mới.\n3. Chậm tự công bố trên TradeNode.vn khiến khách hàng lớn/FDI không tra cứu được xếp hạng năng lực chính thức.",
        "mnemonic": "Nhớ bộ ba chỉ số phân hạng '20-10-3': Hạng I từ 20 ha, Hạng II từ 10 ha, Hạng III từ 3 ha — cộng thêm bộ đôi 'Số hóa + Xanh hóa'.",
        "comparison": "Trung tâm Logistics Hạng I chuẩn Thông tư 71/2026 (Kết nối >= 2 phương thức, >= 20 ha, chuẩn xanh & WMS) vs Kho bãi truyền thống chưa phân hạng (Khó tiếp cận gói thầu chuỗi cung ứng FDI & TMĐT xuyên biên giới).",
        "persona": "Anh Dũng (Giám đốc Điều hành kho bãi / Chủ đầu tư hạ tầng logistics) & Bộ phận Phát triển kinh doanh 3PL.",
        "checklist": "☐ Đo đạc lại diện tích thực tế & năng lực TEU/tấn/đơn hàng năm\n☐ Kiểm tra điều kiện kết nối cao tốc/quốc lộ & 2 phương thức vận tải\n☐ Nâng cấp WMS, camera 24/7, kiểm soát thẻ thông minh\n☐ Lắp đặt trạm sạc sạch & chuẩn bị hồ sơ công bố TradeNode.vn trước 08/01/2027",
        "url": "https://www.tinnhanhchungkhoan.vn/quy-dinh-moi-ve-tieu-chi-phan-loai-phan-hang-trung-tam-logistics-post401780.html",
        "keywords": ["71/2026", "TradeNode"],
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Từ khóa SEO Top 2 Sheet 1. Đối chiếu 3+ nguồn: Tin Nhanh Chứng Khoán (tinnhanhchungkhoan.vn), Thư Viện Pháp Luật và Cổng TTĐT Chính phủ."
    },
    {
        "stt": 275,
        "week": "W47 — 21/11/2026",
        "title": "Kỹ Thuật 5 Bước Ứng Phó Điều Tra Chống Lẩn Tránh & Phòng Vệ Thương Mại Kép (Ukraine AD-615 Dây Thép, EU 2026/2497 Ống Đồng, Thái Lan PP 7,73%)",
        "steps": "Bước 1: Kiểm tra ngay mã HS và thị trường đích của lô hàng xuất khẩu (dây thép 7229.20 sang Ukraine; ống đồng 7411.10 sang EU; hạt nhựa PP nhóm 3902 sang Thái Lan) để xác định tình trạng bị điều tra hoặc áp thuế tạm thời/tự vệ.\nBước 2: Tách bạch chuỗi truy xuất nguyên liệu đầu vào (hóa đơn mua phôi/nguyên liệu nội địa hoặc nhập khẩu ngoài nước bị áp thuế gốc như Trung Quốc) để chứng minh chuyển đổi bản chất tại Việt Nam, chống cáo buộc lẩn tránh thuế.\nBước 3: Đăng ký tham gia trả lời bản câu hỏi điều tra đúng hạn (hoặc phối hợp nhà nhập khẩu xử lý khoản tiền đặt cọc/ký quỹ bảo lãnh tại cửa khẩu nước nhập khẩu như yêu cầu của Ukraine và EU).\nBước 4: Rà soát điều khoản thuế phòng vệ thương mại trong hợp đồng ngoại thương (tránh ký DDP giá cố định mà không có điều khoản điều chỉnh nếu phát sinh thuế CBPG/tự vệ bất ngờ).\nBước 5: Tận dụng 'cửa sổ vàng' trước ngày thu thuế (ví dụ trước 28/12/2026 đối với nhựa PP sang Thái Lan) để hoàn tất giao hàng các đơn đã ký.",
        "formula": "Biên độ tác động giá = Giá CIF × (Thuế suất CBPG/Tự vệ + Thuế NK thông thường). Ví dụ: Ống đồng sang EU chịu 12,3–14,4% (2026/2497); Nhựa PP sang Thái Lan chịu 7,73% từ 28/12/2026; Dây thép sang Ukraine chịu ký quỹ đặt cọc theo AD-615/2026/441-01.",
        "top3_errors": "1. Không hợp tác trả lời bản câu hỏi điều tra nên bị áp mức thuế toàn quốc cao nhất (như 14,4% đối với ống đồng sang EU thay vì 12,3%).\n2. Mua nguyên liệu từ nước bị áp thuế gốc gia công sơ sài rồi xuất đi mà không tính hàm lượng giá trị gia tăng, dẫn đến dính điều tra chống lẩn tránh.\n3. Ký hợp đồng giá DDP cố định khiến bên bán Việt Nam phải gánh toàn bộ khoản thuế ký quỹ bất ngờ khi tàu đang trên biển.",
        "mnemonic": "Nhớ nguyên tắc '3 Minh Bạch Khi Xuất Khẩu': Minh bạch nguồn phôi đầu vào, Minh bạch giá thành kế toán và Minh bạch điều khoản thuế trong hợp đồng.",
        "comparison": "Doanh nghiệp chủ động hợp tác điều tra & tách chuỗi phôi sạch (Hưởng thuế thấp nhất 12,3% hoặc được miễn trừ chống lẩn tránh) vs Doanh nghiệp phớt lờ điều tra (Nhận mức thuế trừng phạt cao nhất 14,4% và mất thị trường).",
        "persona": "Anh Minh (Chủ doanh nghiệp sản xuất xuất khẩu thép, đồng, nhựa) & Giám đốc Kinh doanh xuất khẩu.",
        "checklist": "☐ Rà soát cảnh báo PVTM của Cục Phòng vệ thương mại hàng tuần\n☐ Lưu hồ sơ truy xuất nguồn gốc nguyên liệu thô tối thiểu 5 năm\n☐ Tránh ký hợp đồng DDP không có điều khoản bất khả kháng về thuế PVTM\n☐ Phối hợp trả lời bản câu hỏi đúng hạn khi có khởi xướng điều tra",
        "url": "https://tapchicongthuong.vn/ukraine-khoi-xuong-dieu-tra-chong-lan-tranh-thue-chong-ban-pha-gia-doi-voi-day-thep-viet-nam-567435.htm",
        "keywords": ["AD-615", "day thep"],
        "series": "Series: Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Từ khóa SEO Top 4 & 5 Sheet 1. Đối chiếu 3+ nguồn: Tạp chí Công Thương (tapchicongthuong.vn), ASL Law (aslgate.com) và Bộ Công Thương (moit.gov.vn)."
    },
]


# ==============================================================================
# 4. DATA FOR FUNFACT & INSIGHT (13 cols) — STT 271..275 (W47 — 21/11/2026)
# ==============================================================================
FF_NEW_IDEAS = [
    {
        "stt": 271,
        "week": "W47 — 21/11/2026",
        "title": "Công Ty Đông Ấn Hà Lan (VOC, 1602) & Sàn Chứng Khoán Đầu Tiên Thế Giới: Vì Sao Cổ Phiếu Ra Đời Từ Những Chuyến Tàu Chở Gia Vị?",
        "hook": "Sàn chứng khoán hiện đại đầu tiên của nhân loại không sinh ra từ công nghệ hay ngân hàng — mà ra đời năm 1602 tại Amsterdam chỉ để chia sẻ rủi ro cho những con tàu gỗ chở tiêu và quế từ Châu Á vượt đại dương!",
        "history": "Ngày 20/03/1602, Quốc hội Hà Lan hợp nhất các công ty thương mại cạnh tranh thành Công ty Đông Ấn Hà Lan (Vereenigde Oostindische Compagnie — VOC), trao độc quyền thương mại Châu Á trong 21 năm và thực hiện đợt phát hành cổ phiếu ra công chúng (IPO) đầu tiên trong lịch sử.",
        "context": "Mỗi chuyến tàu viễn dương thế kỷ 17 mất nhiều năm lênh đênh, chi phí đóng tàu khổng lồ và rủi ro đắm tàu hay cướp biển cực cao. Để huy động vốn dài hạn thay vì góp vốn theo từng chuyến đơn lẻ, VOC phát hành cổ phần vĩnh viễn và tạo ra thị trường thứ cấp mua bán lại cổ phiếu ngay tại Sàn Chứng khoán Amsterdam.",
        "modern_impact": "Cơ chế công ty cổ phần và thị trường vốn mà VOC khai sinh chính là nền móng tài chính giúp các tập đoàn vận tải biển và cảng biển toàn cầu ngày nay huy động hàng tỷ USD đóng siêu tàu 24.000 TEU.",
        "intl_comp": "Cùng thời kỳ, Công ty Đông Ấn Anh (EIC, thành lập 1600) ban đầu vẫn góp vốn theo từng chuyến rồi mới chuyển sang vốn cố định, cho thấy sáng kiến tài chính hàng hải của Hà Lan đã đi trước thời đại và định hình chuẩn mực vốn toàn cầu.",
        "business_lesson": "Muốn đi đường dài trong thương mại quốc tế có rủi ro cao, doanh nghiệp phải chuyển từ tư duy 'đánh quả từng chuyến hàng' sang xây dựng cấu trúc vốn bền vững và chia sẻ rủi ro minh bạch.",
        "url": "https://en.wikipedia.org/wiki/Dutch_East_India_Company",
        "keywords": ["VOC", "1602"],
        "series": "Series: Lịch sử Logistics",
        "status": "Chưa dùng",
        "note": "Hướng 1 (Cột mốc lịch sử thương mại). Nguồn xác minh Wikipedia (Dutch East India Company), đối chiếu World History & Amsterdam Stock Exchange."
    },
    {
        "stt": 272,
        "week": "W47 — 21/11/2026",
        "title": "Cảng Tự Do Hải Nam (Hainan Free Trade Port) — Cột Mốc Đóng Cửa Hải Quan Độc Lập 18/12/2025: Biến Cả Hòn Đảo Thành Đặc Khu Miễn Thuế",
        "hook": "Thay vì chỉ lập vài khu kho ngoại quan nhỏ quanh cảng, Trung Quốc đã vận hành quy chế Hải quan độc lập trên toàn bộ đảo Hải Nam từ ngày 18/12/2025 theo triết lý 'Mở cửa ở tuyến 1, kiểm soát ở tuyến 2'!",
        "history": "Sau Luật Cảng Thương mại Tự do Hải Nam (2021), ngày 18/12/2025 đánh dấu cột mốc hòn đảo này chính thức vận hành cơ chế đóng cửa hải quan độc lập (independent customs closure operations), trở thành đặc khu thương mại tự do có diện tích lớn nhất khu vực.",
        "context": "Khác với mô hình cửa khẩu truyền thống, Hải Nam áp dụng mô hình 'giám sát hàng hóa, thông thoáng cho con người' (supervision of goods, not people): hàng hóa từ nước ngoài vào đảo được miễn thuế rộng rãi tại 'tuyến 1', và chỉ tính thuế nhập khẩu khi đưa qua 'tuyến 2' vào thị trường đại lục — thậm chí được miễn thuế nếu đạt tỷ lệ gia tăng giá trị chế biến >= 30% tại đảo.",
        "modern_impact": "Sự trỗi dậy của các khu thương mại tự do thế hệ mới tại Châu Á đặt ra bài học cạnh tranh trực diện cho Việt Nam khi chuẩn hóa các trung tâm logistics khu phi thuế quan theo Thông tư 71/2026/TT-BCT.",
        "intl_comp": "Nếu Singapore hay Dubai JAFZA mạnh nhờ vị trí trung chuyển cảng biển nhỏ gọn, Hải Nam lại tận dụng quy mô cả một tỉnh đảo kết hợp ưu đãi thuế doanh nghiệp, gia công giá trị gia tăng và tự do hóa dòng vốn xuyên biên giới.",
        "business_lesson": "Doanh nghiệp XNK Việt Nam xuất hàng sang Trung Quốc hoặc nhập nguyên liệu qua các khu chế xuất, phi thuế quan cần nắm chắc ranh giới hải quan 'tuyến 1 – tuyến 2' để tận dụng ưu đãi thuế hợp pháp.",
        "url": "https://en.wikipedia.org/wiki/Hainan_Free_Trade_Port",
        "keywords": ["Hainan", "Free Trade"],
        "series": "Series: Thương mại Quốc tế",
        "status": "Chưa dùng",
        "note": "Hướng 2 (So sánh hệ thống XNK quốc tế). Nguồn xác minh Wikipedia (Hainan Free Trade Port), đối chiếu haikou.gov.cn & trade.gov."
    },
    {
        "stt": 273,
        "week": "W47 — 21/11/2026",
        "title": "Báo Cáo 'Containers Lost at Sea' (Cập Nhật 2026) Của WSC: Chỉ 0,0005% Rơi Xuống Biển Trong 280 Triệu Container — Nhưng Vì Sao Không Thể Chủ Quan?",
        "hook": "Trong khoảng 280 triệu container vận chuyển đường biển toàn cầu năm 2025, chỉ có khoảng 1.478 chiếc bị rơi mất tích ngoài đại dương (tương đương 0,0005%) — nhưng một vụ đổi tuyến qua Mũi Hảo Vọng hay khai sai hàng nguy hiểm đủ khiến con số thiệt hại nhân lên gấp bội!",
        "history": "Ngày 25/06/2026, Hội đồng Vận tải Biển Thế giới (World Shipping Council — WSC) công bố báo cáo thường niên 'Containers Lost at Sea Report: 2026 Update' tổng hợp dữ liệu trực tiếp từ các hãng tàu container lớn nhất toàn cầu.",
        "context": "Dù tỷ lệ an toàn đạt tới 99,9995% trên tổng lượng ~280 triệu container luân chuyển năm 2025, WSC cảnh báo các yếu tố thời tiết cực đoan, hành trình vòng qua Mũi Hảo Vọng (Cape of Good Hope) nhiều sóng dữ và tình trạng chủ hàng khai báo sai hàng nguy hiểm (như than củi, pin lithium) vẫn là nguyên nhân hàng đầu gây rơi đổ chồng container và cháy nổ trên tàu.",
        "modern_impact": "Áp lực từ các sự cố rơi container đã thúc đẩy Tổ chức Hàng hải Quốc tế (IMO) ban hành quy định bắt buộc báo cáo container rơi mất trên biển và siết chặt kiểm soát khai báo hàng nguy hiểm (IMDG Code).",
        "intl_comp": "Cùng với Chương trình Top Tier và WSC Cargo Safety Program, các cảng lớn Châu Âu và Mỹ hiện quét dữ liệu booking bằng AI để chặn đứng container khai sai tên hàng ngay từ cảng xếp.",
        "business_lesson": "Xác suất mất container chỉ 0,0005% khiến nhiều chủ hàng chủ quan bỏ qua mua bảo hiểm hàng hải (Cargo Insurance) — cho đến khi 1 cơn bão ngoài Mũi Hảo Vọng cuốn trôi toàn bộ lô hàng trị giá hàng tỷ đồng.",
        "url": "https://www.worldshipping.org/news/world-shipping-council-releases-containers-lost-at-sea-report-2026-update",
        "keywords": ["Containers Lost at Sea", "World Shipping Council"],
        "series": "Series: Bí mật Hải quan",
        "status": "Chưa dùng",
        "note": "Hướng 3 (Con số bất ngờ / Thống kê). Nguồn xác minh trực tiếp World Shipping Council (worldshipping.org)."
    },
    {
        "stt": 274,
        "week": "W47 — 21/11/2026",
        "title": "Kỳ Tích Tàu Chở Hàng Liberty Ship (1941–1945): Khi Henry J. Kaiser Đóng Xong Một Con Tàu Viễn Dương Trong Đúng 4 Ngày 15 Giờ!",
        "hook": "Giữa Thế chiến II, nhà công nghiệp Henry J. Kaiser đã biến quy trình đóng một con tàu chở hàng mất cả năm trời thành dây chuyền lắp ráp module — hoàn thiện chiếc tàu SS Robert E. Peary chỉ trong 4 ngày 15 giờ 29 phút!",
        "history": "Từ năm 1941 đến 1945, 18 xưởng đóng tàu dọc bờ biển Hoa Kỳ đã xuất xưởng tổng cộng 2.710 chiếc tàu chở hàng Liberty Ship — mẫu tàu vận tải được sản xuất hàng loạt lớn nhất trong lịch sử hàng hải nhân loại.",
        "context": "Thay vì gò từng tấm thép theo kiểu thủ công, Henry J. Kaiser áp dụng triệt để tư duy chuẩn hóa bản vẽ, chế tạo sẵn các khối module (prefabrication) tại nhiều phân xưởng khác nhau và dùng hàn điện thay thế tán đinh truyền thống để lắp ghép siêu tốc trên triền đà.",
        "modern_impact": "Tư duy 'chuẩn hóa module và lắp ghép song song' của dự án Liberty Ship chính là tiền thân trực tiếp cho công nghệ đóng tàu container hiện đại và cách chuẩn hóa kiện hàng pallet/container trong chuỗi cung ứng ngày nay.",
        "intl_comp": "Ngày nay, các cường quốc đóng tàu như Hàn Quốc, Trung Quốc và Nhật Bản vẫn kế thừa triết lý chia nhỏ block thân tàu chế tạo sẵn rồi cẩu ghép tại ụ khô để rút ngắn thời gian giao tàu xuống kỷ lục.",
        "business_lesson": "Trong vận hành doanh nghiệp xuất nhập khẩu và logistics, tốc độ xử lý chứng từ không đến từ việc ép nhân viên thức đêm, mà đến từ việc 'module hóa' quy trình kiểm tra và checklist chuẩn trước khi hàng về.",
        "url": "https://en.wikipedia.org/wiki/Liberty_ship",
        "keywords": ["Liberty", "2,710"],
        "series": "Series: Đàm phán & Tư duy",
        "status": "Chưa dùng",
        "note": "Hướng 4 (Bài học từ nhân vật lịch sử). Nguồn xác minh Wikipedia (Liberty ship), đối chiếu usmm.org."
    },
    {
        "stt": 275,
        "week": "W47 — 21/11/2026",
        "title": "FedEx Express World Hub Tại Memphis (1973): Vì Sao Một Thành Phố Giữa Đất Liền Lại Trở Thành 'Siêu Cảng Hàng Hóa' Bận Rộn Nhất Bắc Mỹ?",
        "hook": "Không nằm cạnh biển lớn cũng chẳng phải thủ đô tài chính, thành phố Memphis (bang Tennessee) vẫn trở thành một trong những sân bay chở hàng bận rộn nhất hành tinh chỉ nhờ một ý tưởng điên rồ năm 1973: gom mọi chuyến bay về một điểm lúc nửa đêm!",
        "history": "Năm 1973, FedEx Express chọn Sân bay Quốc tế Memphis (nay mang tên Frederick W. Smith International Airport) làm trung tâm phân loại toàn cầu (Superhub), vận hành mô hình trục-nan (Hub-and-Spoke) tập trung cao điểm vào khung giờ đêm từ 22h đến 5h sáng.",
        "context": "Thay vì bay thẳng từ mọi thành phố này sang thành phố kia gây lãng phí tải trọng, toàn bộ máy bay FedEx từ khắp Bắc Mỹ, Châu Âu và Châu Á hội tụ về Memphis trong đêm, băng chuyền tự động phân loại hàng trăm nghìn bưu kiện mỗi giờ rồi máy bay lại cất cánh tỏa đi trước bình minh.",
        "modern_impact": "Mô hình Hub-and-Spoke của FedEx tại Memphis đã định hình toàn bộ ngành chuyển phát nhanh hàng không (Express Air Freight), chuỗi lạnh y tế và thương mại điện tử xuyên biên giới hiện đại.",
        "intl_comp": "Mô hình siêu hub hàng không Memphis được DHL (tại Leipzig/Cincinnati), UPS (Worldport tại Louisville) và các trung tâm logistics hàng không tại Châu Á như Hong Kong, Incheon học hỏi và phát triển quy mô khổng lồ.",
        "business_lesson": "Một trung tâm trung chuyển hiệu quả không nhất thiết phải nằm ở nơi đắt đỏ nhất, mà nằm ở điểm tối ưu hóa được luồng gom hàng (consolidation) và tốc độ phân loại qua đêm.",
        "url": "https://en.wikipedia.org/wiki/Memphis_International_Airport",
        "keywords": ["FedEx", "Memphis"],
        "series": "Series: Công nghệ & Đổi mới",
        "status": "Chưa dùng",
        "note": "Hướng 5 (Đổi mới mô hình & công nghệ vận tải). Nguồn xác minh Wikipedia (Memphis International Airport), đối chiếu FedEx Express."
    },
]


# ==============================================================================
# 5. FEEDBACK LOOP: MISSING VERIFIED LEGAL DOCS
# ==============================================================================
NEW_LEGAL_DOCS = [
    {
        "ma_so": "Nghị quyết 19/2026/NQ-CP",
        "ten": "Nghị quyết về việc cắt giảm, phân cấp, đơn giản hóa thủ tục hành chính, điều kiện kinh doanh thuộc phạm vi quản lý của Bộ Công Thương",
        "loai": "Nghị quyết",
        "co_quan": "Chính phủ",
        "ngay_ban_hanh": "2026-04-29",
        "ngay_hieu_luc": "2026-04-29",
        "ngay_het_hieu_luc": "Còn hiệu lực",
        "van_ban_bi_thay_the": "Cắt giảm một số TTHC thuộc lĩnh vực Công Thương (dẫn đến vướng mắc dừng cấp C/O cho DNCX/kho ngoại quan tại CV 24381 & 24449/CHQ-GSQL)",
        "phan_tich": "Bãi bỏ một số thủ tục hành chính thuộc phạm vi quản lý của Bộ Công Thương, trong đó có thủ tục cấp C/O cấp sau và cấp C/O cho hàng hóa từ DNCX, khu phi thuế quan, kho ngoại quan bán vào nội địa.",
        "so_sanh": "Tinh giản TTHC tại Bộ Công Thương nhưng đang phát sinh điểm vênh với yêu cầu xuất trình C/O khi làm thủ tục hải quan nhập khẩu tại chỗ vào nội địa."
    },
    {
        "ma_so": "Thông tư 06/2026/TT-BTC",
        "ten": "Thông tư sửa đổi, bổ sung một số điều của Thông tư số 13/2015/TT-BTC về kiểm tra, giám sát, tạm dừng làm thủ tục hải quan đối với hàng hóa XNK có yêu cầu bảo vệ quyền sở hữu trí tuệ",
        "loai": "Thông tư",
        "co_quan": "Bộ Tài chính",
        "ngay_ban_hanh": "2026-01-26",
        "ngay_hieu_luc": "2026-03-01",
        "ngay_het_hieu_luc": "Còn hiệu lực",
        "van_ban_bi_thay_the": "Sửa đổi, bổ sung Thông tư 13/2015/TT-BTC",
        "phan_tich": "Cho phép nộp hồ sơ bảo vệ quyền SHTT trực tuyến qua hệ thống hải quan điện tử và trao quyền cho cơ quan Hải quan chủ động tạm dừng thông quan (kể cả hàng quá cảnh) khi có căn cứ rõ ràng nghi ngờ xâm phạm quyền SHTT.",
        "so_sanh": "Mở rộng quyền chủ động ngăn chặn tại cửa khẩu của cơ quan Hải quan thay vì chỉ chờ đơn yêu cầu của chủ thể quyền."
    },
    {
        "ma_so": "Quyết định 3453/QĐ-BCT",
        "ten": "Quyết định về việc áp dụng biện pháp chống bán phá giá chính thức đối với một số sản phẩm tháp điện gió có xuất xứ từ nước Cộng hòa Nhân dân Trung Hoa (mã vụ việc AD18)",
        "loai": "Quyết định",
        "co_quan": "Bộ Công Thương",
        "ngay_ban_hanh": "2024-12-24",
        "ngay_hieu_luc": "2025-01-08",
        "ngay_het_hieu_luc": "2030-01-08",
        "van_ban_bi_thay_the": "Quyết định áp thuế chống bán phá giá chính thức AD18 (thời hạn 5 năm)",
        "phan_tich": "Áp dụng thuế chống bán phá giá chính thức đối với sản phẩm tháp điện gió xuất xứ Trung Quốc (nhóm mã HS 7308.20.11, 7308.20.19...); được Cục Hải quan hướng dẫn cơ chế miễn thuế khi tạo TSCĐ dự án ưu đãi đầu tư tại CV 24424/CHQ-NVTHQ.",
        "so_sanh": "Doanh nghiệp thuộc dự án ưu đãi đầu tư có thể được miễn thuế CBPG AD18 khi nhập khẩu tạo TSCĐ nếu đăng ký danh mục miễn thuế đúng quy định."
    },
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
    """Ensure STT 271-275 and titles do not duplicate existing Excel rows."""
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
        if last_stt != 270:
            wb.close()
            raise ValueError(f"{path} last STT is {last_stt}, expected 270 before appending 271-275!")
        existing_titles = {str(ws.cell(r, 3).value or "").strip().lower() for r in range(5, ws.max_row + 1)}
        for nt in new_titles:
            if nt in existing_titles:
                wb.close()
                raise ValueError(f"Trùng tiêu đề trong {path}: {nt}")
        wb.close()
    print("🟢 DUPLICATE & SEQUENCE CHECK PASSED: 4 file Ideas Bank sẵn sàng nối tiếp STT 271-275.")


def append_ideas_to_sheet(path, ideas, row_builder, bold_cols, center_cols, expected_last_stt=275):
    """Load workbook, append formatted rows, and save atomically."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    for item in ideas:
        r = ws.max_row + 1
        fill = FILL_ODD if item["stt"] % 2 != 0 else FILL_EVEN
        row_vals = row_builder(item)
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in bold_cols else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in center_cols else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
    atomic_save(wb, path)
    max_row = ws.max_row
    wb.close()
    print(f"✅ Ghi thành công {path} (Tổng dòng: {max_row}, STT cuối: {expected_last_stt})")


def execute_update():
    print("=" * 80)
    print("🚀 WORKFLOW 06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (WEEK 47 DAY 6 — 21/11/2026)")
    print("=" * 80)

    check_banned_words()
    check_no_duplicates()

    # 1. TECHNICAL GATEKEEPER: VALIDATE ALL 20 URLS BEFORE TOUCHING EXCEL
    all_ideas_for_gatekeeper = [
        {
            "stt": item["stt"],
            "title": item["title"],
            "url": item["url"],
            "keywords": item["keywords"],
        }
        for item in TN_NEW_IDEAS + CS_NEW_IDEAS + HT_NEW_IDEAS + FF_NEW_IDEAS
    ]

    print(f"\n🔍 Gửi {len(all_ideas_for_gatekeeper)} link qua Technical Gatekeeper url_validator.py...")
    enforce_strict_validation(all_ideas_for_gatekeeper)
    print("\n🟢 TECHNICAL GATEKEEPER PHÊ DUYỆT 100%: TẤT CẢ 20 LINK ĐỀU HỢP LỆ VÀ ĐÃ ĐỐI SOÁT NỘI DUNG THỰC TẾ!\n")

    # 2. TIN NGANH (15 cols)
    append_ideas_to_sheet(
        "reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx",
        TN_NEW_IDEAS,
        lambda item: [
            item["stt"], item["week"], item["title"], item["agency"], item["effective_date"],
            item["core_diff"], item["affected_entity"], item["risk_pain"], item["market_reaction"],
            item["pro_argument"], item["con_argument"], item["intl_ref"], item["url"],
            item["status"], item["note"],
        ],
        bold_cols=(1, 14),
        center_cols=(1, 2, 14),
    )

    # 3. CASE STUDY (14 cols)
    append_ideas_to_sheet(
        "reports/Case_Study_Ideas_Bank_Eureka.xlsx",
        CS_NEW_IDEAS,
        lambda item: [
            item["stt"], item["week"], item["title"], item["parties"], item["plot"],
            item["damage"], item["root_cause"], item["prevention"], item["opposing_views"],
            item["lessons"], item["url"], item["series"], item["status"], item["note"],
        ],
        bold_cols=(1, 13),
        center_cols=(1, 2, 13),
    )

    # 4. HOC THUC CHIEN (14 cols)
    append_ideas_to_sheet(
        "reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx",
        HT_NEW_IDEAS,
        lambda item: [
            item["stt"], item["week"], item["title"], item["steps"], item["formula"],
            item["top3_errors"], item["mnemonic"], item["comparison"], item["persona"],
            item["checklist"], item["url"], item["series"], item["status"], item["note"],
        ],
        bold_cols=(1, 13),
        center_cols=(1, 2, 13),
    )

    # 5. FUNFACT (13 cols)
    append_ideas_to_sheet(
        "reports/Funfact_Ideas_Bank_Eureka.xlsx",
        FF_NEW_IDEAS,
        lambda item: [
            item["stt"], item["week"], item["title"], item["hook"], item["history"],
            item["context"], item["modern_impact"], item["intl_comp"], item["business_lesson"],
            item["url"], item["series"], item["status"], item["note"],
        ],
        bold_cols=(1, 12),
        center_cols=(1, 2, 12),
    )

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
            "2026-11-21",
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
    max_leg_row = ws_leg.max_row
    wb_leg.close()
    print(f"✅ Feedback Loop: Bổ sung {added_legal} văn bản mới vào {legal_path} (Tổng dòng: {max_leg_row})")

    print("\n" + "=" * 80)
    print("🎉 HOÀN TẤT WORKFLOW 06 (21/11/2026): ĐÃ CẬP NHẬT 20 Ý TƯỞNG MỚI (STT 271-275) & FEEDBACK LOOP!")
    print("=" * 80)


if __name__ == "__main__":
    execute_update()
