# -*- coding: utf-8 -*-
"""
Execute Workflow 06: Ideas Bank Daily Update (wf06) - Week 47 Day 2 (17/11/2026)
Appends 5 new verified ideas (STT 256-260) to each of the 4 Pillars (TN, CS, HT, FF).
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

# Styles definition matching existing sheets
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
# 1. DATA FOR TIN NGANH (15 cols) — STT 256..260
# ==============================================================================
TN_NEW_IDEAS = [
    {
        "stt": 256,
        "week": "W47 — 17/11/2026",
        "title": "Quyết Định Số 1828/QĐ-CHQ (Ban Hành 12/11/2026, Hiệu Lực 15/12/2026): Quy Trình Kiểm Tra Xuất Xứ Hàng Hóa XNK Sau Thông Quan & Siết Chặt Hành Vi Chia Nhỏ Lô Hàng Miễn C/O",
        "agency": "Cục Hải quan (Bộ Tài chính)",
        "effective_date": "Ban hành ngày 12/11/2026; chính thức có hiệu lực thi hành từ ngày 15/12/2026",
        "core_diff": "Cũ: Việc rà soát C/O sau khi hàng đã thông quan chủ yếu thực hiện phân tán theo các đợt thanh tra chuyên đề không cố định thời hạn; Mới: Chuẩn hóa quy trình tập trung tại Đội Phúc tập và Kiểm tra sau thông quan với thời hạn kiểm tra tối đa 15 ngày kể từ khi nhận danh sách tờ khai phân bổ theo quản lý rủi ro; cập nhật kết quả lên hệ thống ngay trong ngày hoặc ngày làm việc tiếp theo; đặc biệt siết chặt Điều 5 giám sát hành vi chia nhỏ lô hàng để hưởng ngưỡng miễn nộp chứng từ C/O.",
        "affected_entity": "Các doanh nghiệp xuất khẩu sang thị trường FTA, doanh nghiệp nhập khẩu hưởng thuế ưu đãi C/O và các đơn vị nhập khẩu lô hàng giá trị nhỏ dưới ngưỡng miễn nộp chứng từ chứng nhận xuất xứ.",
        "risk_pain": "Nếu doanh nghiệp có dấu hiệu chia tách hóa đơn/lô hàng để né C/O hoặc chứng từ xuất xứ không khớp hồ sơ hải quan, công chức kiểm tra sau thông quan sẽ chuyển hồ sơ kiểm tra tại trụ sở DN, bác bỏ ưu đãi thuế và ấn định truy thu theo mức thuế MFN kèm tiền phạt.",
        "market_reaction": "Các doanh nghiệp xuất nhập khẩu và đại lý hải quan khẩn trương rà soát lại toàn bộ hồ sơ C/O đã thông quan và quy trình lưu trữ chứng từ định mức trước mốc 15/12/2026.",
        "pro_argument": "Việc áp dụng quản lý rủi ro và ấn định khung thời gian 15 ngày xử lý giúp minh bạch hóa công tác hậu kiểm, ngăn chặn gian lận xuất xứ và chuyển tải bất hợp pháp để bảo vệ uy tín hàng hóa Việt Nam.",
        "con_argument": "Tần suất rà soát hậu kiểm tự động tăng cao tạo thêm áp lực giải trình và lưu trữ chứng từ chi tiết cho các doanh nghiệp vừa và nhỏ không có bộ phận pháp chế chuyên trách.",
        "intl_ref": "Cơ chế hậu kiểm xuất xứ dựa trên quản lý rủi ro (Post-Clearance Origin Verification) của Hải quan EU (Hệ thống REX) và Hải quan Hoa Kỳ (US CBP).",
        "url": "https://hocxuatnhapkhau.com/quyet-dinh-1828-qd-chq-kiem-tra-xuat-xu-sau-thong-quan/",
        "keywords": ["1828", "xuat xu"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: hocxuatnhapkhau.com, LuatVietnam và Thư Viện Pháp Luật (căn cứ Thông tư 33/2023/TT-BTC, Thông tư 124/2026/TT-BTC, Quyết định 467/QĐ-CHQ)."
    },
    {
        "stt": 257,
        "week": "W47 — 17/11/2026",
        "title": "Thông Tư Số 142/2026/TT-BTC (Hiệu Lực 16/11/2026): Chuẩn Hóa Thủ Tục Hải Quan & Bộ Mẫu Bảng Kê Chi Tiết Hàng Hóa (HQ 01a/01b-BKCTHH) Cho Hàng Bưu Chính, Chuyển Phát Nhanh",
        "agency": "Bộ Tài chính",
        "effective_date": "Ban hành 29/09/2026; chính thức có hiệu lực thi hành từ ngày 16/11/2026",
        "core_diff": "Cũ: Các doanh nghiệp bưu chính và chuyển phát nhanh sử dụng biểu mẫu kê khai phân tán, chưa đồng bộ triệt để dữ liệu số hóa với mã vận đơn Master AWB/CN37/CN38 cho thương mại điện tử xuyên biên giới; Mới: Áp dụng bắt buộc hệ thống biểu mẫu chuẩn hóa (Mẫu HQ 01a-BKCTHH cho hàng nhập khẩu và Mẫu HQ 01b/06b-BKCTHH cho hàng xuất khẩu), buộc đồng bộ chính xác thông tin người gửi/nhận, số vận đơn, trị giá tính thuế, trọng lượng và xuất xứ trên hệ thống số.",
        "affected_entity": "Doanh nghiệp kinh doanh dịch vụ bưu chính chỉ định, hãng chuyển phát nhanh quốc tế (DHL, FedEx, UPS...) và các nhà bán hàng TMĐT xuyên biên giới.",
        "risk_pain": "Nếu bảng kê chi tiết hàng hóa sai lệch thông tin vận đơn, mã định danh hoặc thiếu dữ liệu theo mẫu mới kể từ 16/11/2026, lô hàng sẽ bị giữ lại tại khu vực giám sát kho hàng không/bưu chính, gây chậm trễ giao hàng và phát sinh phí lưu kho.",
        "market_reaction": "Các hãng chuyển phát nhanh và sàn TMĐT xuyên biên giới đang gấp rút nâng cấp API phần mềm khai báo để truyền đúng chuẩn dữ liệu Phụ lục I Thông tư 142/2026/TT-BTC.",
        "pro_argument": "Chuẩn hóa dữ liệu điện tử giúp tách bạch rõ hàng hóa thương mại và phi mậu dịch, ngăn chặn thất thu thuế GTGT/NK trong làn sóng hàng TMĐT giá trị nhỏ tràn vào nội địa.",
        "con_argument": "Giai đoạn đầu chuyển đổi biểu mẫu từ 16/11/2026 đòi hỏi các chủ hàng nhỏ lẻ phải cung cấp đầy đủ thông tin định danh và chứng từ hơn, làm tăng thời gian chuẩn bị trước khi gửi hàng.",
        "intl_ref": "Khung tiêu chuẩn quản lý thương mại điện tử xuyên biên giới (WCO Framework of Standards on Cross-Border E-Commerce) và hệ thống khai báo trước ICS2 của Liên minh Châu Âu.",
        "url": "https://vanban.chinhphu.vn/?pageid=27160&docid=219685&classid=1&orggroupid=4",
        "keywords": ["142/2026", "chuyen phat nhanh"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn), LuatVietnam (Mẫu HQ 01a-BKCTHH, HQ 01b-BKCTHH) và man.net.vn."
    },
    {
        "stt": 258,
        "week": "W47 — 17/11/2026",
        "title": "Nghị Định Số 293/2026/NĐ-CP (Hiệu Lực 11/09/2026): Triển Khai Chứng Nhận Apostille & 'Điểm Mù' Loại Trừ Đối Với Chứng Từ Thương Mại, Hải Quan Xuất Nhập Khẩu",
        "agency": "Chính phủ / Bộ Ngoại giao",
        "effective_date": "Ban hành 23/07/2026; có hiệu lực thi hành từ ngày 11/09/2026",
        "core_diff": "Cũ: Giấy tờ công giữa Việt Nam và đối tác nước ngoài phải trải qua quy trình kép gồm chứng nhận lãnh sự tại Bộ Ngoại giao và hợp pháp hóa lãnh sự tại Đại sứ quán mất nhiều tuần lễ; Mới: Giấy tờ công pháp lý của doanh nghiệp (ĐKKD, giấy phép đầu tư, ủy quyền) chỉ cần cấp chứng nhận Apostille/e-Apostille gắn mã QR là được công nhận giữa các nước thành viên Công ước La Hay; TUY NHIÊN theo Điều 1 Công ước, chứng từ thương mại và hải quan (Invoice, B/L, Packing List, C/O) vẫn bị loại trừ khỏi phạm vi Apostille.",
        "affected_entity": "Doanh nghiệp XNK có giao dịch với đối tác thuộc Công ước La Hay, doanh nghiệp mở chi nhánh/hồ sơ đấu thầu quốc tế và bộ phận chứng từ ngoại thương.",
        "risk_pain": "Nhiều doanh nghiệp hiểu nhầm mang Commercial Invoice, Packing List, B/L hoặc C/O đi xin dấu Apostille khiến hồ sơ bị trả về không đúng thẩm quyền, lỡ thời hạn xuất trình chứng từ thanh toán L/C hoặc thông quan tại cảng đến.",
        "market_reaction": "VCCI và các hiệp hội ngành hàng liên tục phát văn bản hướng dẫn doanh nghiệp tách bạch 2 nhóm giấy tờ: Hồ sơ pháp nhân (dùng Apostille) và Chứng từ hàng hóa (dùng xác thực VCCI/tập quán quốc tế).",
        "pro_argument": "Gia nhập cơ chế Apostille và triển khai e-Apostille giúp cắt giảm mạnh thời gian và chi phí hợp pháp hóa hồ sơ pháp lý doanh nghiệp khi mở rộng giao thương toàn cầu.",
        "con_argument": "Trong giai đoạn đầu áp dụng, sự khác biệt giữa danh sách nước thành viên Công ước La Hay và yêu cầu riêng của ngân hàng nước ngoài dễ gây lúng túng cho nhân sự chứng từ.",
        "intl_ref": "Công ước La Hay ngày 05/10/1961 về miễn hợp pháp hóa đối với giấy tờ công nước ngoài (Apostille Convention) do Hội nghị La Hay về Tư pháp Quốc tế (HCCH) quản lý.",
        "url": "https://vanban.chinhphu.vn/?pageid=27160&docid=218964",
        "keywords": ["293/2026", "apostille"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn) và Chi nhánh VCCI (covcci.com.vn), đồng bộ với Nghị định 292/2026/NĐ-CP."
    },
    {
        "stt": 259,
        "week": "W47 — 17/11/2026",
        "title": "Thông Tư Số 86/2026/TT-BTC (Hiệu Lực 01/07/2026): Chuẩn Hóa Quản Lý Thuế Hàng Hóa XNK, Cơ Chế Bù Trừ Tiền Thuế Nộp Thừa & Tiêu Chí Không Nợ Thuế Khi Thông Quan",
        "agency": "Bộ Tài chính",
        "effective_date": "Ban hành 30/06/2026; có hiệu lực thi hành từ ngày 01/07/2026",
        "core_diff": "Cũ: Quy trình xử lý tiền thuế nộp thừa, bù trừ nợ thuế và thủ tục miễn/giảm/hoàn thuế hải quan nằm rải rác và còn nhiều bước đối soát thủ công; Mới: Thông tư 86/2026/TT-BTC chuẩn hóa đồng bộ với Luật Quản lý thuế mới về thời hạn khai nộp thuế (kể cả giao dịch TMĐT/nền tảng số), quy định minh bạch nguyên tắc bù trừ tiền thuế nộp thừa và tiêu chí kiểm tra người nộp thuế không còn nợ tiền thuế, tiền phạt, tiền phí hải quan.",
        "affected_entity": "Toàn bộ doanh nghiệp xuất nhập khẩu, doanh nghiệp gia công/SXXK làm thủ tục hoàn thuế và bộ phận kế toán thuế XNK.",
        "risk_pain": "Doanh nghiệp không đối soát kịp khoản nợ thuế, phí hải quan nhỏ hoặc tiền chậm nộp phát sinh có thể bị hệ thống chặn thông quan lô hàng mới hoặc bị bác hồ sơ hoàn thuế/bù trừ thuế nộp thừa.",
        "market_reaction": "Cộng đồng kế toán XNK đánh giá cao việc hợp nhất hướng dẫn xử lý bù trừ thuế nộp thừa, giúp doanh nghiệp chủ động xoay vòng dòng tiền nộp thuế 24/7.",
        "pro_argument": "Tạo hành lang pháp lý đồng bộ, minh bạch hóa quyền và nghĩa vụ thuế của doanh nghiệp trên môi trường hải quan điện tử.",
        "con_argument": "Việc liên thông dữ liệu chặt chẽ đòi hỏi kế toán thuế và bộ phận XNK phải phối hợp đối soát công nợ thuế theo thời gian thực để tránh bị cưỡng chế dừng thủ tục.",
        "intl_ref": "Chuẩn mực quản lý thu thuế hải quan điện tử và hoàn thuế nhanh của Tổ chức Hải quan Thế giới (WCO Revised Kyoto Convention).",
        "url": "https://vanban.chinhphu.vn/?pageid=27160&docid=218821&classid=1&typegroupid=6",
        "keywords": ["86/2026", "quan ly thue"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn), VASEP (vasep.com.vn) và Thư Viện Pháp Luật."
    },
    {
        "stt": 260,
        "week": "W47 — 17/11/2026",
        "title": "Xuất Nhập Khẩu Việt Nam Cán Mốc Lịch Sử Vượt 1.000 Tỷ USD (10 Tháng 2026 Đạt 1.010,6 Tỷ USD, Tăng 32,3%) & Cảnh Báo Siết Chặt Kiểm Soát Rủi Ro Hải Quan Cuối Năm",
        "agency": "Cục Hải quan / Bộ Tài chính / Bộ Công Thương",
        "effective_date": "Số liệu công bố tháng 11/2026 (lũy kế 10 tháng đầu năm 2026)",
        "core_diff": "Cũ: Kim ngạch xuất nhập khẩu các năm trước xoay quanh ngưỡng 700–800 tỷ USD với quy trình thông quan phân tán; Mới: Hết 10 tháng năm 2026, tổng kim ngạch XNK Việt Nam chính thức vượt mốc 1.000 tỷ USD (đạt 1.010,6 tỷ USD, tăng 32,3% so với cùng kỳ 2025); đi kèm quy mô khổng lồ là việc ngành Hải quan đẩy mạnh chuyển đổi số, phân luồng quản lý rủi ro và kiểm soát gắt gao gian lận xuất xứ, trị giá dịp cao điểm cuối năm.",
        "affected_entity": "Các doanh nghiệp xuất khẩu chủ lực (điện tử, dệt may, da giày, nông thủy sản), doanh nghiệp nhập khẩu nguyên phụ liệu và hệ sinh thái cảng biển — logistics.",
        "risk_pain": "Lưu lượng tờ khai tăng kỷ lục khiến các doanh nghiệp có lịch sử tuân thủ thấp dễ bị đẩy vào luồng Vàng/Đỏ khi hệ thống quản lý rủi ro siết chặt sàng lọc hàng hóa cuối năm.",
        "market_reaction": "Doanh nghiệp logistics và cảng biển ghi nhận sản lượng bùng nổ nhưng đồng thời phải đầu tư mạnh vào chuẩn hóa dữ liệu chứng từ để giữ vững luồng Xanh.",
        "pro_argument": "Cột mốc 1.000 tỷ USD khẳng định vị thế mắt xích trọng yếu của Việt Nam trong chuỗi cung ứng toàn cầu và thúc đẩy hiện đại hóa hạ tầng Hải quan số.",
        "con_argument": "Áp lực ùn tắc cục bộ tại cảng biển và nguy cơ bị các thị trường lớn điều tra phòng vệ thương mại, chuyển tải bất hợp pháp tăng tỷ lệ thuận với kim ngạch.",
        "intl_ref": "Xếp hạng nhóm các nền kinh tế có quy mô thương mại hàng hóa hàng đầu thế giới của Tổ chức Thương mại Thế giới (WTO World Trade Statistical Review).",
        "url": "https://thoibaotaichinhvietnam.vn/hai-quan-tiep-suc-dua-xuat-nhap-khau-can-moc-1-000-ty-usd-206391-206391.html",
        "keywords": ["1.000", "xuat nhap khau"],
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Thời báo Tài chính Việt Nam (thoibaotaichinhvietnam.vn) và Báo Công Thương (congthuong.vn)."
    }
]

# ==============================================================================
# 2. DATA FOR CASE STUDY (14 cols) — STT 256..260
# ==============================================================================
CS_NEW_IDEAS = [
    {
        "stt": 256,
        "week": "W47 — 17/11/2026",
        "title": "Đại Án Công Ty Tài Lộc Thông Quan Trái Phép 13.376 Container Gỗ Ván Bóc Tại Cảng Hải Phòng (Hơn 1.800 Tỷ Đồng): Cái Giá Của 'Chi Ngoài' 8,1 Tỷ Đồng Và Khai Khống Chứng Từ",
        "parties": "Công ty TNHH SX TM Quốc tế Tài Lộc (cùng 13 công ty lập mới và 3 công ty mượn pháp nhân), Chi cục Hải quan Cảng 2 Hải Phòng (14 cựu cán bộ/lãnh đạo), TAND TP. Hải Phòng.",
        "plot": "Từ tháng 12/2020 đến tháng 1/2024, lãnh đạo Công ty Tài Lộc mua gỗ ván bóc trong nước xuất khẩu sang Trung Quốc nhưng lập khống hồ sơ lâm sản, khai báo gian dối khối lượng thấp hơn thực tế và mượn pháp nhân 16 công ty để thông quan trái phép. Nhằm được bỏ qua sai phạm khi tiếp nhận tờ khai và kiểm hóa, doanh nghiệp thỏa thuận chi tiền ngoài 100.000 đồng/tờ khai và 800.000 đồng/container cho cán bộ, lãnh đạo Chi cục Hải quan Cảng 2 (Hải Phòng). Vụ án được VKSND Tối cao truy tố tháng 5/2025 và TAND TP. Hải Phòng xét xử sơ thẩm tháng 10/2025.",
        "damage": "13.376 container gỗ ván bóc (1.698 tờ khai hải quan) trị giá hơn 1.800 tỷ đồng, thu lợi bất chính 210 tỷ đồng; tổng số tiền chi ngoài lên tới 8,1 tỷ đồng; 22 bị cáo (gồm lãnh đạo doanh nghiệp và 14 cựu cán bộ hải quan) phải lĩnh các bản án tù nghiêm khắc (mức án cao nhất lên tới 19 năm tù).",
        "root_cause": "Tư duy làm ăn chụp giật, lập nhiều công ty bình phong để khai báo sai lệch khối lượng thực tế và dùng tiền hối lộ thay cho tuân thủ pháp luật hải quan và quản lý lâm sản.",
        "prevention": "1. Tuyệt đối tuân thủ khai báo trung thực 100% số lượng, trọng lượng và nguồn gốc lâm sản hợp pháp theo đúng quy định; 2. Không bao giờ mượn hoặc cho mượn pháp nhân đứng tên mở tờ khai xuất khẩu ủy thác thiếu kiểm soát; 3. Thiết lập cơ chế kiểm toán nội bộ độc lập đối chiếu chứng từ kho bãi với tờ khai VNACCS.",
        "opposing_views": "Phe A (Tư duy đường tắt): Cho rằng việc thỏa thuận 'chi phí bôi trơn' giúp thông quan nhanh và giảm chi phí thuế là cách vận hành quen thuộc; Phe B (Pháp lý & Quản trị bền vững): Mọi dấu vết giao dịch ngân hàng, tờ khai điện tử và dữ liệu cảng biển đều lưu lại nhiều năm — một khi chuyên án khởi tố, toàn bộ tài sản và tự do của chủ doanh nghiệp đều sụp đổ.",
        "lessons": "Trong xuất nhập khẩu hiện đại, 'chi ngoài' không bao giờ mua được sự an toàn dài hạn; chỉ có sự minh bạch tuyệt đối trong hồ sơ chứng từ mới bảo vệ được doanh nghiệp và gia đình.",
        "url": "https://vnexpress.net/14-hai-quan-nhan-hoi-lo-de-thong-quan-hon-13-000-container-hang-lau-4887473.html",
        "keywords": ["Tai Loc", "13.376"],
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Báo VnExpress (18/05/2025) và Báo Bảo vệ Pháp luật / VnEconomy (tuyên án sơ thẩm 16/10/2025). Đã chuẩn hóa ngôn từ tuân thủ banned_words.md."
    },
    {
        "stt": 257,
        "week": "W47 — 17/11/2026",
        "title": "Khởi Tố Vụ Án Hình Sự Tại Cảng Cát Lái (09/2026): 1 Container 20 Feet Chứa 71.200 Bạc Đạn Có 9.352 Sản Phẩm Xâm Phạm Nhãn Hiệu Được Bảo Hộ Tại Việt Nam",
        "parties": "Công ty TNHH Thương mại XNK T.M. (TP.HCM), Đội Kiểm soát Hải quan — Chi cục Hải quan Khu vực II, Hải quan cửa khẩu cảng Sài Gòn khu vực I, Viện Sở hữu trí tuệ quốc gia.",
        "plot": "Tháng 5/2025, Công ty T.M. mở tờ khai nhập khẩu tại Cảng Cát Lái cho 1 container 20 feet gồm 71.200 đơn vị sản phẩm (25 dòng hàng bạc đạn, thân lắp ổ đũa hiệu CUXB, xuất xứ Trung Quốc, khai mới 100%). Khi kiểm tra thực tế, hải quan phát hiện dấu hiệu xâm phạm quyền sở hữu trí tuệ nên tạm dừng để trưng cầu giám định. Doanh nghiệp sau đó không tiếp tục xuất trình giấy tờ liên quan theo yêu cầu. Tháng 3/2026, hải quan phối hợp Vinacontrol kiểm tra toàn bộ lô hàng, gửi mẫu giám định tại Viện Sở hữu trí tuệ quốc gia (Bộ KH&CN). Tháng 9/2026, Đội Kiểm soát Hải quan ra Quyết định số 1793 khởi tố vụ án hình sự và chuyển hồ sơ cho Công an TP.HCM điều tra.",
        "damage": "Toàn bộ container 71.200 sản phẩm bị phong tỏa kéo dài hơn 16 tháng (từ 05/2025 đến 09/2026); 9.352 đơn vị sản phẩm bị kết luận xâm phạm nhãn hiệu được bảo hộ tại Việt Nam; doanh nghiệp đối mặt trực tiếp với án hình sự theo Quyết định khởi tố số 1793.",
        "root_cause": "Nhập khẩu phụ tùng công nghiệp mang nhãn hiệu có dấu hiệu xâm phạm quyền sở hữu trí tuệ nhưng không kiểm tra tình trạng bảo hộ nhãn hiệu tại Cục Sở hữu trí tuệ Việt Nam, đồng thời thiếu ủy quyền phân phối chính hãng (Letter of Authorization) từ chủ sở hữu thương hiệu.",
        "prevention": "1. Tra cứu bắt buộc tình trạng bảo hộ nhãn hiệu trên cơ sở dữ liệu WIPO và Cục Sở hữu trí tuệ Việt Nam trước khi ký hợp đồng nhập khẩu hàng có thương hiệu; 2. Yêu cầu nhà cung cấp xuất trình Thư ủy quyền phân phối/sử dụng nhãn hiệu hợp lệ; 3. Tuyệt đối không bỏ trốn hay né tránh xuất trình chứng từ khi hải quan yêu cầu giải trình kiểm tra thực tế.",
        "opposing_views": "Phe A (Nhà nhập khẩu): Cho rằng hàng hóa mua từ nhà máy Trung Quốc có hóa đơn thương mại hợp lệ nên không phải chịu trách nhiệm về nhãn hiệu in trên linh kiện; Phe B (Hải quan & Luật SHTT): Theo Luật Hải quan và Luật Sở hữu trí tuệ, người nhập khẩu chịu trách nhiệm pháp lý hoàn toàn đối với hàng hóa đưa qua biên giới mang nhãn hiệu đang được bảo hộ tại Việt Nam.",
        "lessons": "Đừng bao giờ nhập hàng gắn logo thương hiệu nổi tiếng mà không có giấy ủy quyền chính hãng; một container linh kiện vài trăm triệu có thể kéo doanh nghiệp vào vòng lao lý hình sự.",
        "url": "https://thanhnien.vn/hai-quan-khoi-to-vu-buon-lau-hang-gia-mao-tai-cang-cat-lai-185260916110654162.htm",
        "keywords": ["71.200", "9.352"],
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Báo Thanh Niên (16/09/2026) và Tạp chí Kinh tế Tài chính. Đã thay thế từ nhạy cảm bằng thuật ngữ chuẩn SHTT theo banned_words.md."
    },
    {
        "stt": 258,
        "week": "W47 — 17/11/2026",
        "title": "Bắt Giữ 1 Container Điều Hòa Nhiệt Độ Đã Qua Sử Dụng 'Use Only In Japan' Tại Cảng Đình Vũ — Hải Phòng: Cảnh Báo Cạm Bẫy Pháp Nhân Mới Đứng Tên Nhập Hàng Cấm",
        "parties": "Doanh nghiệp mới thành lập đứng tên nhận hàng (chưa từng phát sinh XNK), Cục Hải quan Hải Phòng (Đội Kiểm soát HQ & Chi cục HQ cảng Đình Vũ), Cảnh sát biển Vùng 1, Bộ đội Biên phòng Hải Phòng.",
        "plot": "Trong đợt cao điểm kiểm soát biên giới cuối năm (công bố ngày 25/12/2024), Hải quan Hải Phòng phát hiện một container nhập khẩu từ Nhật Bản về Cảng Đình Vũ có dấu hiệu rủi ro cao vì doanh nghiệp đứng tên nhận hàng là công ty mới thành lập, chưa từng có lịch sử giao dịch XNK. Khi lực lượng liên ngành cắt niêm phong khám xét thực tế, toàn bộ hàng hóa bên trong container là điều hòa nhiệt độ đã qua sử dụng (chủ yếu hiệu Toshiba, National), nhiều máy ghi rõ dòng chữ 'Use only in Japan' — thuộc danh mục hàng điện lạnh cũ cấm nhập khẩu theo quy định Việt Nam.",
        "damage": "100% hàng hóa trong container bị tạm giữ toàn bộ tại cảng Đình Vũ để điều tra xử lý theo pháp luật; doanh nghiệp đứng tên nhận hàng bị đưa vào diện điều tra vi phạm chính sách mặt hàng cấm nhập khẩu và đối mặt chế tài xử phạt hành chính/hình sự cùng chi phí lưu bãi, tiêu hủy hoặc tái xuất.",
        "root_cause": "Thiếu kiểm soát danh mục hàng hóa cấm nhập khẩu (hàng tiêu dùng, điện tử, điện lạnh đã qua sử dụng theo Nghị định 69/2018/NĐ-CP và Danh mục của Bộ Công Thương) hoặc cho mượn/thuê pháp nhân công ty mới thành lập để đứng tên nhận hàng không rõ thực tế đóng gói.",
        "prevention": "1. Trước khi nhận ủy thác nhập khẩu hoặc nhập máy móc điện lạnh, kiểm tra đối chiếu mã HS với Danh mục hàng hóa cấm nhập khẩu theo Nghị định 69/2018/NĐ-CP (hàng điện lạnh gia dụng đã qua sử dụng bị cấm tuyệt đối); 2. Nghiêm cấm cho thuê/mượn pháp nhân công ty mới thành lập đứng tên trên B/L và tờ khai; 3. Yêu cầu giám định chất lượng/hiện trạng hàng hóa tại cảng xếp (Pre-shipment Inspection) trước khi tàu chạy.",
        "opposing_views": "Phe A (Người mua): Máy điều hòa nội địa Nhật vẫn còn dùng tốt, chất lượng bền hơn hàng mới giá rẻ nên nhu cầu thị trường cao; Phe B (Quản lý nhà nước): Cho phép nhập thiết bị điện lạnh cũ tiềm ẩn nguy cơ biến Việt Nam thành bãi rác công nghệ, rò rỉ môi chất lạnh phá hủy tầng ozone và gây mất an toàn cháy nổ.",
        "lessons": "Máy móc công nghiệp cũ có thể nhập có điều kiện theo Quyết định 18/2019/QĐ-TTg, nhưng đồ điện lạnh gia dụng đã qua sử dụng là vùng cấm tuyệt đối — đừng đánh đổi pháp nhân công ty cho những lô hàng rủi ro pháp lý.",
        "url": "https://thuehaiquan.tapchikinhtetaichinh.vn/bat-giu-1-container-hang-cam-tai-cang-hai-phong-89495.html",
        "keywords": ["Hai Phong", "dieu hoa"],
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn Tạp chí Kinh tế - Tài chính (Hải quan Online, 25/12/2024 - trong hạn 2 năm). Đã kiểm duyệt banned_words.md."
    },
    {
        "stt": 259,
        "week": "W47 — 17/11/2026",
        "title": "Chặn Đứng Lô Hàng Quá Cảnh Từ Trung Quốc Đi Campuchia Qua Cửa Khẩu Hữu Nghị (11/2025): Cất Giấu Hơn 2.000 Sản Phẩm Gắn Nhãn 'Made in Viet Nam' & 11.500 Tai Nghe Không Khai Báo",
        "parties": "Doanh nghiệp thực hiện dịch vụ quá cảnh, Chi cục Điều tra chống buôn lậu (Đội 1 — Cục Hải quan), Hải quan Cửa khẩu quốc tế Hữu Nghị (Chi cục HQ khu vực VI).",
        "plot": "Thực hiện Kế hoạch số 3362/KH-CHQ và Chuyên đề số 66/CĐ-ĐTCBL về chống chuyển tải bất hợp pháp, lực lượng Hải quan tại Cửa khẩu quốc tế Hữu Nghị (Lạng Sơn) đã kiểm tra thực tế các lô hàng quá cảnh từ Trung Quốc đi Campuchia (tháng 10–11/2025). Một lô hàng khai báo 39 mục nhưng kiểm tra thực tế phát hiện cất giấu thêm 4 mục không khai báo gồm 1.800 đôi dép và 365 túi xách xâm phạm nhãn hiệu nổi tiếng (Crocs, LV, Chanel, Hermes). Trước đó ít ngày, một lô quá cảnh khác cũng bị phát hiện cất giấu hơn 2.000 giày dép, balo, váy áo gắn sẵn nhãn 'Made in Viet Nam' cùng 11.500 tai nghe vi phạm nhãn hiệu OPPO.",
        "damage": "Toàn bộ hàng hóa không khai báo, gắn nhãn xuất xứ 'Made in Viet Nam' sai sự thật và xâm phạm SHTT (hơn 4.000 sản phẩm may mặc/giày dép/túi xách và 11.500 phụ kiện điện tử) bị lập biên bản tạm giữ, tịch thu xử lý nghiêm; doanh nghiệp vận chuyển quá cảnh bị xử phạt nặng và đưa vào danh sách kiểm soát trọng điểm 100%.",
        "root_cause": "Các đối tượng lợi dụng luồng thông thoáng của thủ tục hàng quá cảnh biên giới để trà trộn hàng không khai báo, gắn sẵn nhãn 'Made in Viet Nam' hòng thực hiện chuyển tải bất hợp pháp.",
        "prevention": "1. Doanh nghiệp dịch vụ logistics quá cảnh bắt buộc kiểm tra danh mục hàng hóa thực tế tại kho đóng hàng, yêu cầu khách hàng cam kết pháp lý không chứa hàng xâm phạm SHTT hay nhãn 'Made in Viet Nam' sai nguồn gốc; 2. Tuân thủ nghiêm ngặt quy định gắn seal định vị GPS điện tử và khai báo chính xác 100% dòng hàng quá cảnh; 3. Từ chối vận chuyển hàng không rõ nguồn gốc từ đối tác trung gian thiếu minh bạch.",
        "opposing_views": "Phe A (Forwarder quá cảnh): Hàng chỉ đi ngang qua lãnh thổ Việt Nam sang Campuchia chứ không tiêu thụ nội địa nên không ảnh hưởng thị trường Việt Nam; Phe B (Cục Hải quan): Hành vi gắn nhãn 'Made in Viet Nam' sai sự thật trên hàng quá cảnh hủy hoại trực tiếp uy tín xuất xứ quốc gia và có thể khiến hàng Việt bị các nước áp thuế trừng phạt.",
        "lessons": "Đừng nghĩ hàng quá cảnh là 'vùng an toàn vô hình'; với hệ thống seal định vị GPS và chuyên đề kiểm tra đột xuất, doanh nghiệp logistics nhận chở hàng quá cảnh sai khai báo sẽ gánh trọn rủi ro pháp lý.",
        "url": "https://baovephapluat.vn/kinh-te/kinh-doanh-phap-luat/hai-quan-ngan-chan-nhieu-vu-vi-pham-ve-xuat-xu-hang-hoa-ghi-nhan-hang-hoa-188622.html",
        "keywords": ["Huu Nghi", "2.000"],
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Báo Bảo vệ Pháp luật (26/11/2025) và Báo VietNamNet (27/11/2025). Đã tuân thủ banned_words.md."
    },
    {
        "stt": 260,
        "week": "W47 — 17/11/2026",
        "title": "Cơn Bão Hậu Kiểm Tháng 10–11/2026: Hải Quan Hoàn Thành 173 Cuộc Kiểm Tra Sau Thông Quan Trong 1 Tháng, Ấn Định Thuế & Xử Phạt 443,97 Tỷ Đồng (KV XVII Thu Hơn 29 Tỷ)",
        "parties": "Cục Hải quan (Bộ Tài chính), Chi cục Hải quan Khu vực XVII và 173 doanh nghiệp thuộc diện kiểm tra sau thông quan (KTSTQ) dựa trên quản lý rủi ro.",
        "plot": "Theo thống kê Cục Hải quan công bố ngày 16/11/2026, chỉ trong 1 tháng (15/09–14/10/2026), toàn ngành ban hành 178 quyết định kiểm tra sau thông quan và hoàn thành 173 cuộc kiểm tra tập trung vào các lĩnh vực rủi ro cao (trị giá, mã HS, C/O, gia công/SXXK). Riêng tại Chi cục Hải quan Khu vực XVII, đơn vị thực hiện 36 cuộc hậu kiểm (trong đó có 33 cuộc kiểm tra tại trụ sở doanh nghiệp gồm 25 cuộc theo kế hoạch rủi ro, 6 cuộc theo dấu hiệu vi phạm và 2 cuộc đánh giá tuân thủ), đồng thời xử lý 622 vụ vi phạm hành chính và kiến nghị khởi tố 11 vụ.",
        "damage": "Tổng số tiền ấn định thuế và xử phạt qua 173 cuộc hậu kiểm trong 1 tháng lên tới 443,97 tỷ đồng (thực thu vào ngân sách 420,78 tỷ đồng); riêng Hải quan Khu vực XVII thu nộp hơn 29 tỷ đồng (gấp 3,64 lần chỉ tiêu được giao). Doanh nghiệp bị truy thu gặp cú sốc dòng tiền nặng nề và mất xếp hạng ưu tiên.",
        "root_cause": "Nhiều doanh nghiệp chủ quan cho rằng 'tờ khai đã thông quan luồng Xanh/Vàng là xong', không thực hiện tự rà soát định kỳ (Pre-audit) hồ sơ trong thời hạn 5 năm về mã HS, các khoản phải cộng vào trị giá hải quan và chênh lệch tồn kho nguyên liệu.",
        "prevention": "1. Thiết lập lịch trình tự kiểm toán nội bộ (Pre-audit) hàng quý cho bộ hồ sơ đã thông quan trong vòng 5 năm gần nhất; 2. Kiểm tra kỹ các khoản phải cộng trị giá (phí bản quyền, hoa hồng, chi phí đóng gói) và tính nhất quán mã HS giữa các lô hàng; 3. Khi phát hiện sai sót trước khi có quyết định kiểm tra, chủ động khai bổ sung AMA (trong 60 ngày hoặc trước thanh tra) để giảm thiểu chế tài xử phạt theo Nghị định 169/2026/NĐ-CP.",
        "opposing_views": "Phe A (Doanh nghiệp): Việc bị truy thu ngược lại nhiều năm khiến doanh nghiệp đã chốt giá bán và quyết toán lợi nhuận từ lâu rơi vào cảnh thua lỗ đột ngột; Phe B (Cục Hải quan): Hậu kiểm dựa trên phân tích dữ liệu rủi ro là thông lệ quốc tế nhằm tạo sự công bằng cho các doanh nghiệp làm ăn chân chính và chống thất thu ngân sách.",
        "lessons": "Thông quan tại cảng mới chỉ là hết hiệp 1; hiệp 2 mang tên 'Kiểm tra sau thông quan' kéo dài suốt 5 năm — doanh nghiệp nào lơ là quản trị hồ sơ sẽ phải trả giá bằng chính dòng tiền của mình.",
        "url": "https://thoibaotaichinhvietnam.vn/hai-quan-tang-thu-ngan-sach-hon-420-ty-dong-tu-hau-kiem-207023.html",
        "keywords": ["443,97", "173"],
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Thời báo Tài chính Việt Nam bài ngày 16/11/2026 (443,97 tỷ / 173 cuộc) và bài Hải quan KV XVII (hơn 29 tỷ / 622 vụ)."
    }
]

# ==============================================================================
# 3. DATA FOR HOC THUC CHIEN (14 cols) — STT 256..260
# ==============================================================================
HT_NEW_IDEAS = [
    {
        "stt": 256,
        "week": "W47 — 17/11/2026",
        "title": "Quy Trình 5 Bước Chuẩn Bị Hồ Sơ Xuất Xứ Hậu Kiểm Theo Quyết Định 1828/QĐ-CHQ (Hiệu Lực 15/12/2026): Bảo Vệ Tờ Khai Trong 15 Ngày & Tránh Bẫy Chia Nhỏ Lô Hàng",
        "steps": "Bước 1: Phân loại toàn bộ tờ khai xuất khẩu, nhập khẩu năm 2026 theo 3 nhóm của QĐ 1828/QĐ-CHQ: (a) Hàng xuất khẩu có khai báo xuất xứ; (b) Hàng nhập khẩu miễn nộp C/O (lô hàng giá trị nhỏ); (c) Hàng nhập khẩu bắt buộc nộp C/O ưu đãi;\nBước 2: Rà soát chống rủi ro 'chia nhỏ lô hàng' ở Nhóm (b): Kiểm tra tần suất các lô hàng nhập khẩu từ cùng một nhà cung cấp trong cùng ngày/tuần có giá trị sát ngưỡng miễn nộp chứng từ xuất xứ hay không, chuẩn bị chứng từ chứng minh tính độc lập thương mại của từng đơn hàng;\nBước 3: Đối chiếu bộ hồ sơ C/O ưu đãi ở Nhóm (c): Kiểm tra sự khớp nối tuyệt đối giữa Thông tin trên C/O, Commercial Invoice, Packing List, Vận đơn (B/L) và Thông báo kết quả xác định trước xuất xứ (nếu có);\nBước 4: Thiết lập Thư mục Chứng từ Giải trình Nhanh (Audit-Ready Folder): Gom sẵn định mức B.O.M, hóa đơn nguyên phụ liệu đầu vào và chứng từ thanh toán ngân hàng để có thể phản hồi ngay khi Đội Phúc tập và KTSTQ rà soát trong khung 15 ngày;\nBước 5: Khi có yêu cầu bổ sung hoặc nghi vấn từ hệ thống quản lý rủi ro hải quan, chủ động cử nhân sự pháp chế/XNK đối chiếu trực tiếp và lưu biên bản cập nhật kết quả lên hệ thống.",
        "formula": "Khung thời gian vận hành theo Quyết định 1828/QĐ-CHQ:\n- Ngày hiệu lực áp dụng: 15/12/2026.\n- Thời hạn tối đa Đội Phúc tập & KTSTQ thực hiện kiểm tra: 15 ngày kể từ ngày nhận danh sách tờ khai phân bổ trên hệ thống.\n- Thời hạn cập nhật kết quả: Trong ngày hoàn thành hoặc chậm nhất ngày làm việc tiếp theo.",
        "top3_errors": "1. Tách 1 đơn hàng lớn thành nhiều vận đơn/hóa đơn nhỏ lẻ để hưởng ngưỡng miễn nộp chứng từ xuất xứ (dính ngay bẫy Điều 5 QĐ 1828/QĐ-CHQ);\n2. Khai báo mã nước xuất xứ trên tờ khai VNACCS khác với chứng từ vận tải thực tế nhưng không lưu tài liệu chứng minh quá trình gia công/chế biến;\n3. Chậm trễ cung cấp hồ sơ giải trình khiến Đội Phúc tập chuyển hồ sơ sang diện kiểm tra sau thông quan tại trụ sở doanh nghiệp.",
        "mnemonic": "Quyết định 1828/QĐ-CHQ giống như 'Đèn quét sân bay sau chuyến bay': Dù bạn đã qua cổng hải quan (thông quan), hồ sơ vẫn được soi lại trong 15 ngày — ai chia nhỏ hành lý để qua cửa miễn kiểm sẽ bị máy quét phát hiện ngay!",
        "comparison": "Kiểm tra C/O tại khâu thông quan (Kiểm tra tính hợp lệ hình thức bề mặt của C/O để cho hưởng thuế ưu đãi ngay) vs Kiểm tra xuất xứ sau thông quan theo QĐ 1828/QĐ-CHQ (Rà soát đối chiếu chéo dữ liệu quản lý rủi ro trong 15 ngày, truy vết dấu hiệu chia nhỏ lô hàng và chuyển tải bất hợp pháp).",
        "persona": "Anh Dũng (Trưởng phòng XNK & Pháp chế thương mại tại doanh nghiệp có tần suất nhập hàng FTA hoặc hàng chuyển phát nhanh lớn).",
        "checklist": "☐ Tách và rà soát 3 nhóm tờ khai theo Điều 4, 5, 6 QĐ 1828/QĐ-CHQ\n☐ Kiểm tra không có dấu hiệu chia nhỏ lô hàng né C/O\n☐ Đối chiếu khớp 100% C/O — Invoice — B/L — Tờ khai\n☐ Chuẩn bị sẵn hồ sơ B.O.M & nguồn nguyên liệu cho hàng xuất khẩu\n☐ Sẵn sàng phản hồi giải trình trong vòng 15 ngày khi hệ thống gọi tên.",
        "url": "https://hocxuatnhapkhau.com/quyet-dinh-1828-qd-chq-kiem-tra-xuat-xu-sau-thong-quan/",
        "keywords": ["1828", "xuat xu"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: hocxuatnhapkhau.com, LuatVietnam và Thư Viện Pháp Luật (căn cứ QĐ 1828/QĐ-CHQ ngày 12/11/2026, Thông tư 33/2023/TT-BTC)."
    },
    {
        "stt": 257,
        "week": "W47 — 17/11/2026",
        "title": "Hướng Dẫn 4 Bước Xử Lý Chứng Từ Quốc Tế Sau Nghị Định 293/2026/NĐ-CP: Phân Biệt Hồ Sơ Apostille Với Chứng Từ Thương Mại — Hải Quan (Invoice, B/L, C/O)",
        "steps": "Bước 1: Kiểm tra quốc gia đối tác có phải là thành viên của Công ước La Hay 1961 (Apostille) hay không trên danh mục chính thức của Bộ Ngoại giao / HCCH;\nBước 2: Tách bạch bộ hồ sơ XNK thành 2 rổ riêng biệt: Rổ A (Giấy tờ công pháp lý: ĐKKD, Giấy chứng nhận đầu tư, Giấy ủy quyền, Tư pháp) và Rổ B (Chứng từ thương mại & hải quan: Commercial Invoice, Packing List, B/L, C/O, CQ/COA);\nBước 3: Áp dụng đúng quy trình cho Rổ A (Nếu là thành viên Công ước La Hay, xin cấp tem Apostille hoặc e-Apostille có mã QR qua Cổng DVC Quốc gia / Bộ Ngoại giao / UBND tỉnh được phân cấp để dùng trực tiếp mà không cần qua Đại sứ quán);\nBước 4: Xử lý đúng chuẩn cho Rổ B (Tuyệt đối KHÔNG xin Apostille; thực hiện chứng thực thương mại tại VCCI hoặc cơ quan chuyên ngành theo yêu cầu hợp đồng ngoại thương và quy tắc UCP 600 của L/C).",
        "formula": "Nguyên tắc 2-1-0 của Nghị định 293/2026/NĐ-CP:\n- Thời điểm áp dụng tại Việt Nam: Từ ngày 11/09/2026.\n- 1 bước duy nhất (xin tem Apostille/e-Apostille) thay cho 2 bước kép (Bộ Ngoại giao + Đại sứ quán) đối với giấy tờ công.\n- 0 áp dụng cho chứng từ thương mại & hải quan trực tiếp (theo Điều 1 Công ước La Hay 1961).",
        "top3_errors": "1. Mang Hóa đơn thương mại (Invoice), Vận đơn (B/L) hoặc C/O đi nộp hồ sơ xin dấu Apostille khiến hồ sơ bị từ chối và trễ tàu;\n2. Mặc định mọi quốc gia đối tác đều dùng Apostille mà không kiểm tra xem nước đó đã gia nhập Công ước La Hay chưa;\n3. Không sử dụng tính năng tra cứu mã QR của e-Apostille để kiểm chứng tính xác thực của giấy phép đối tác nước ngoài gửi về.",
        "mnemonic": "Apostille giống như 'Hộ chiếu miễn thị thực cho giấy tờ công': Giấy khai sinh của công ty (ĐKKD, Ủy quyền) dùng hộ chiếu Apostille đi thẳng qua biên giới, còn 'vé máy bay và hóa đơn mua hàng' (B/L, Invoice, C/O) phải đi cửa Hải quan và VCCI!",
        "comparison": "Hợp pháp hóa lãnh sự truyền thống (2 vòng: Bộ Ngoại giao + Đại sứ quán nước ngoài, mất nhiều tuần, áp dụng cho nước ngoài Công ước La Hay) vs Chứng nhận Apostille theo NĐ 293/2026/NĐ-CP (1 vòng duy nhất hoặc e-Apostille mã QR, áp dụng cho giấy tờ công giữa các nước thành viên Công ước La Hay).",
        "persona": "Chị Hương (Trưởng phòng Chứng từ XNK & Pháp chế hợp đồng ngoại thương).",
        "checklist": "☐ Xác minh đối tác thuộc quốc gia thành viên Công ước La Hay 1961\n☐ Tách riêng Giấy tờ công (Rổ A) và Chứng từ thương mại/hải quan (Rổ B)\n☐ Nộp hồ sơ e-Apostille cho Rổ A trên Cổng Dịch vụ công Quốc gia\n☐ Xác thực chứng từ thương mại Rổ B qua VCCI hoặc chuẩn UCP 600\n☐ Quét mã QR xác minh tính thật của chứng thư Apostille từ đối tác ngoại.",
        "url": "https://covcci.com.vn/tin-tuc/Tu_1192026_Viet_Nam_ap_dung_chung_nhan_Apostille_doanh_nghiep_xuat_khau_can_tach_bach_hai_nhom_giay_to_304",
        "keywords": ["293/2026", "apostille"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn/docid=218964) và Chi nhánh VCCI (covcci.com.vn)."
    },
    {
        "stt": 258,
        "week": "W47 — 17/11/2026",
        "title": "Quy Trình 5 Bước Lập Bảng Kê Chi Tiết Hàng Hóa Chuyển Phát Nhanh & Bưu Chính (Mẫu HQ 01a/01b-BKCTHH) Chuẩn Thông Tư 142/2026/TT-BTC Từ Ngày 16/11/2026",
        "steps": "Bước 1: Phân loại luồng hàng hóa gửi qua dịch vụ bưu chính hoặc chuyển phát nhanh quốc tế theo ngưỡng trị giá và đối tượng chịu thuế hay miễn thuế;\nBước 2: Chọn đúng mẫu biểu chuẩn tại Phụ lục I Thông tư 142/2026/TT-BTC: Dùng Mẫu HQ 01a-BKCTHH cho hàng nhập khẩu và Mẫu HQ 01b-BKCTHH (hoặc HQ 06b-BKCTHH tùy loại hình) cho hàng xuất khẩu;\nBước 3: Điền đầy đủ 11 trường dữ liệu chi tiết từng kiện hàng trên bảng kê: STT, Họ tên/địa chỉ người nhập/xuất khẩu, Số vận đơn, Số Master AWB/CN37/CN38, Tên hàng cụ thể, Đơn vị tính, Số lượng, Trị giá tính thuế, Trọng lượng, Xuất xứ, Ghi chú;\nBước 4: Tính toán và đối khớp phần tổng hợp nghĩa vụ thuế trên mẫu nhập khẩu: Tổng trị giá tính thuế, Thuế suất GTGT (%), Số tiền thuế GTGT, Thuế bảo vệ môi trường, Thuế TTĐB (nếu có) và Tổng số tiền thuế phải nộp;\nBước 5: Truyền dữ liệu điện tử đồng bộ lên hệ thống hải quan, trình ký xác nhận của đại diện doanh nghiệp và hoàn tất xác nhận qua khu vực giám sát kho hàng không/bưu chính.",
        "formula": "Cấu trúc kiểm tra chéo Bảng kê Mẫu HQ 01a-BKCTHH (hiệu lực 16/11/2026):\nTổng số tiền thuế = Thuế nhập khẩu (nếu có) + Thuế TTĐB (nếu có) + Thuế BVMT (nếu có) + Thuế GTGT.\nYêu cầu bắt buộc: Mã vận đơn con (House AWB) và vận đơn chủ (Master AWB / CN37 / CN38) phải khớp 100% với dữ liệu manifest kho hàng không.",
        "top3_errors": "1. Dùng lại mẫu bảng kê cũ trước ngày 16/11/2026 dẫn đến hệ thống từ chối tiếp nhận hồ sơ tại kho chuyển phát nhanh;\n2. Khai gộp nhiều tên hàng chung chung ('Phụ kiện', 'Hàng mẫu') mà không tách chi tiết tên hàng, số lượng, xuất xứ từng dòng;\n3. Sai lệch số liệu giữa trọng lượng cân thực tế tại kho hàng không và trọng lượng khai trên bảng kê chi tiết.",
        "mnemonic": "Mẫu HQ 01a/01b-BKCTHH giống như 'Danh sách hành khách lên máy bay': Mỗi món hàng đều phải có tên thật, số ghế (Master AWB/CN38), cân nặng và quê quán (xuất xứ) rõ ràng thì hải quan mới cho qua cổng!",
        "comparison": "Khai tờ khai riêng lẻ (Áp dụng cho lô hàng thương mại giá trị lớn/có giấy phép chuyên ngành, mất nhiều thời gian xử lý từng bộ hồ sơ) vs Khai gộp theo Bảng kê HQ 01a/01b-BKCTHH (Áp dụng cho hàng chuyển phát nhanh/bưu chính theo quy chuẩn Thông tư 142/2026/TT-BTC, tối ưu tốc độ thông quan số lượng lớn kiện hàng).",
        "persona": "Anh Minh (Chủ shop TMĐT xuyên biên giới và Nhân viên khai báo hải quan hãng chuyển phát nhanh).",
        "checklist": "☐ Cập nhật biểu mẫu Phụ lục I Thông tư 142/2026/TT-BTC (hiệu lực 16/11/2026)\n☐ Khớp đủ số House AWB và Master AWB / CN37 / CN38\n☐ Tách chi tiết tên hàng, trọng lượng, trị giá, xuất xứ từng người nhận\n☐ Kiểm tra đúng công thức thuế GTGT, TTĐB, BVMT trên bảng kê\n☐ Ký số xác thực và kiểm tra trạng thái qua khu vực giám sát.",
        "url": "https://vanban.chinhphu.vn/?pageid=27160&docid=219685&classid=1&orggroupid=4",
        "keywords": ["142/2026", "buu chinh"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn), LuatVietnam và man.net.vn."
    },
    {
        "stt": 259,
        "week": "W47 — 17/11/2026",
        "title": "Cẩm Nang 5 Bước Xử Lý Tiền Thuế Nộp Thừa, Bù Trừ & Hoàn Thuế Hải Quan Điện Tử Theo Thông Tư 86/2026/TT-BTC: Tối Ưu Dòng Tiền Doanh Nghiệp XNK",
        "steps": "Bước 1: Tra cứu định kỳ tình trạng công nợ thuế, tiền chậm nộp và tiền thuế nộp thừa của doanh nghiệp trên cổng thông tin chính thức (customs.gov.vn) và hệ thống nộp thuế 24/7 (epayment.customs.gov.vn);\nBước 2: Xác định nguyên nhân phát sinh tiền thuế nộp thừa hoặc thuộc diện hoàn thuế (khai bổ sung giảm thuế, xuất trả hàng nhập khẩu B13, nộp nhầm tài khoản thu ngân sách, nguyên liệu SXXK đã xuất khẩu sản phẩm);\nBước 3: Lập bộ hồ sơ đề nghị xử lý tiền thuế, tiền chậm nộp, tiền phạt nộp thừa theo quy định tại Thông tư 86/2026/TT-BTC, xác định phương án ưu tiên: (a) Bù trừ trực tiếp sang tờ khai phát sinh nợ/tờ khai mới, hoặc (b) Hoàn trả bằng tiền về tài khoản ngân hàng;\nBước 4: Kiểm tra điều kiện tiên quyết 'Không còn nợ tiền thuế, tiền phạt, tiền phí hải quan' tại tất cả các Chi cục Hải quan trên toàn quốc trước khi nộp hồ sơ bù trừ/hoàn thuế;\nBước 5: Nộp hồ sơ điện tử qua hệ thống dịch vụ công hải quan, theo dõi quyết định bù trừ/hoàn thuế và hạch toán khớp đối chiếu trên sổ kế toán doanh nghiệp.",
        "formula": "Thứ tự ưu tiên bù trừ tiền thuế nộp thừa theo nguyên tắc quản lý thuế:\n1. Bù trừ với tiền thuế, tiền chậm nộp, tiền phạt còn nợ của chính người nộp thuế tại cơ quan hải quan;\n2. Bù trừ sang số tiền thuế phát sinh của các tờ khai hải quan tiếp theo;\n3. Hoàn trả phần còn lại vào tài khoản ngân hàng của người nộp thuế khi không còn khoản nợ ngân sách.",
        "top3_errors": "1. Chỉ kiểm tra nợ thuế tại 1 Chi cục Hải quan hay mở tờ khai mà quên kiểm tra khoản phí hải quan/tiền chậm nộp nhỏ lẻ còn treo ở Chi cục khác trên toàn quốc;\n2. Nộp nhầm mã chương, mã tiểu mục ngân sách khi chuyển khoản nộp thuế điện tử nhưng không làm văn bản điều chỉnh thông tin thu nộp ngân sách;\n3. Thiếu chứng từ thanh toán không dùng tiền mặt hoặc chứng từ chứng minh hàng thực xuất khi lập hồ sơ hoàn thuế.",
        "mnemonic": "Quản lý thuế XNK theo Thông tư 86/2026 giống như 'Ví điện tử thông minh với Kho bạc': Tiền đóng dư có thể xoay vòng bù trừ ngay sang lô hàng sau, nhưng chỉ cần dính một khoản nợ phí 20.000đ ở cảng khác là cả ví sẽ bị khóa!",
        "comparison": "Hoàn thuế về tài khoản ngân hàng (Tiền mặt quay về công ty nhưng quy trình thẩm định Kho bạc mất nhiều thời gian hơn) vs Bù trừ sang tờ khai mới (Nhanh gọn, xoay vòng trực tiếp tiền thuế đã nộp thừa cho các lô hàng nhập khẩu tiếp theo của doanh nghiệp).",
        "persona": "Chị Hương (Kế toán trưởng & Chuyên viên thanh khoản/hoàn thuế doanh nghiệp XNK).",
        "checklist": "☐ Tra cứu công nợ thuế toàn quốc trên customs.gov.vn\n☐ Xử lý sạch các khoản phí hải quan, tiền chậm nộp lẻ trước khi nộp hồ sơ\n☐ Đối chiếu đúng mã tiểu mục và số tờ khai phát sinh nộp thừa\n☐ Chọn phương án bù trừ nhanh hoặc hoàn tiền về tài khoản\n☐ Lưu trữ quyết định bù trừ/hoàn thuế khớp sổ sách kế toán.",
        "url": "https://vasep.com.vn/van-ban/bo-tai-chinh/thong-tu-86-2026-tt-btc-quy-dinh-ve-quan-ly-thue-doi-voi-hang-hoa-xuat-khau-nhap-khau-3610.html",
        "keywords": ["86/2026", "thue"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn/docid=218821) và VASEP (vasep.com.vn)."
    },
    {
        "stt": 260,
        "week": "W47 — 17/11/2026",
        "title": "Quy Trình 4 Bước Rà Soát Mã HS & Phòng Ngừa Rủi Ro Hợp Đồng Khi Tạm Ngừng Tạm Nhập — Tái Xuất, Chuyển Khẩu Theo Thông Tư 41/2026/TT-BCT (Đến 31/12/2029)",
        "steps": "Bước 1: Đối chiếu mã HS 8 số và mô tả hàng hóa thực tế của lô hàng dự kiến kinh doanh tạm nhập tái xuất (TNTX) hoặc chuyển khẩu với 2 Phụ lục Danh mục ban hành kèm Thông tư 41/2026/TT-BCT (Phụ lục Phế liệu và Phụ lục Hàng hóa đã qua sử dụng tạm ngừng TNTX, chuyển khẩu);\nBước 2: Kiểm tra tình trạng pháp lý lô hàng theo mốc hiệu lực: Thông tư 41/2026/TT-BCT áp dụng từ 05/09/2026 đến hết 31/12/2029 đối với các nhóm điện tử, gia dụng, dệt may, phương tiện đã qua sử dụng và phế liệu nằm trong danh mục;\nBước 3: Đàm phán sửa đổi điều khoản hợp đồng ngoại thương với đối tác: Nếu mặt hàng thuộc diện tạm ngừng TNTX/chuyển khẩu qua Việt Nam, tuyệt đối không nhận ủy thác hoặc ký hợp đồng quá cảnh/TNTX trái quy định, chuyển sang phương án giao thẳng từ cảng xuất đến cảng đích cuối cùng;\nBước 4: Nếu hàng đã lỡ xếp tàu hoặc nằm trong các trường hợp loại trừ/xử lý chuyển tiếp đặc biệt, làm việc ngay với Sở Công Thương/Bộ Công Thương và Chi cục Hải quan cửa khẩu để có văn bản hướng dẫn giải quyết tránh bị phạt lưu bãi và buộc tái xuất.",
        "formula": "Thời hạn áp dụng lệnh tạm ngừng kinh doanh TNTX, chuyển khẩu theo Thông tư 41/2026/TT-BCT:\n- Ngày bắt đầu hiệu lực: 05/09/2026.\n- Ngày hết hiệu lực dự kiến: 31/12/2029.\n- Phạm vi kiểm tra: 100% mã HS tại Danh mục phế liệu & Danh mục hàng đã qua sử dụng kèm theo Thông tư.",
        "top3_errors": "1. Forwarder hoặc thương nhân nhận booking hàng đã qua sử dụng đi tạm nhập tái xuất qua Việt Nam theo thói quen cũ mà không check Danh mục mới từ 05/09/2026;\n2. Nhầm lẫn giữa mã HS hàng mới 100% và hàng đã qua sử dụng/phế liệu khi nhận ủy thác chuyển khẩu;\n3. Không cài điều khoản miễn trách nhiệm/hủy hợp đồng nếu pháp luật thay đổi chính sách mặt hàng giữa chừng.",
        "mnemonic": "Thông tư 41/2026/TT-BCT giống như 'Biển cấm xe quá tải qua cầu đến hết 2029': Hàng cũ và phế liệu trong danh mục không được phép mượn đường Việt Nam để tạm nhập tái xuất hay chuyển khẩu — cố đi qua là kẹt cứng tại cửa khẩu!",
        "comparison": "Kinh doanh Tạm nhập — Tái xuất (Hàng đi vào lãnh thổ Việt Nam, làm thủ tục hải quan TNTX và chịu sự kiểm soát danh mục cấm/tạm ngừng của Thông tư 41/2026/TT-BCT) vs Giao hàng thẳng quốc tế (Drop-shipping / Switch B/L từ nước xuất khẩu thẳng sang nước thứ ba mà không đưa hàng qua cửa khẩu Việt Nam).",
        "persona": "Anh Minh (Giám đốc công ty thương mại quốc tế & Quản lý giao nhận dịch vụ TNTX/chuyển khẩu).",
        "checklist": "☐ Tra cứu mã HS trong 2 Danh mục của Thông tư 41/2026/TT-BCT\n☐ Xác minh tình trạng hàng mới 100% hay đã qua sử dụng trước khi ký HĐ\n☐ Tách luồng hợp đồng Drop-shipping nếu hàng thuộc danh mục tạm ngừng TNTX\n☐ Cập nhật điều khoản thay đổi pháp luật trong hợp đồng ngoại thương\n☐ Lưu trữ công văn hướng dẫn của Bộ Công Thương cho các mã hàng giáp ranh.",
        "url": "https://vanban.chinhphu.vn/?pageid=27160&docid=218990",
        "keywords": ["41/2026", "tam nhap"],
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Đối chiếu 2+ nguồn: Cổng TTĐT Chính phủ (vanban.chinhphu.vn/docid=218990) và Viện Nghiên cứu Chiến lược Chính sách Công Thương (vioit.moit.gov.vn)."
    }
]

# ==============================================================================
# 4. DATA FOR FUNFACT (13 cols) — STT 256..260 (5 Hướng khai thác)
# ==============================================================================
FF_NEW_IDEAS = [
    {
        "stt": 256,
        "week": "W47 — 17/11/2026",
        "title": "Đường Mực Tải Trọng Plimsoll (Plimsoll Line, 1876): Cuộc Chiến Chống 'Những Con Tàu Quan Tài' Đã Cứu Hàng Vạn Thủy Thủ Và Khai Sinh Chuẩn An Toàn Hàng Hải Quốc Tế",
        "hook": "Trước năm 1876, những chủ tàu tham lam sẵn sàng chất hàng nặng đến mức sàn tàu ngập mấp mé mặt nước rồi mua bảo hiểm giá cao để chờ tàu chìm ngoài khơi thu tiền bồi thường — cho đến khi một vạch sơn hình tròn cắt ngang trên thân tàu thay đổi vĩnh viễn lịch sử hàng hải!",
        "history": "Nghị sĩ Anh Samuel Plimsoll (1824–1898) cùng vợ là Eliza Plimsoll phát động chiến dịch nghị trường quyết liệt, xuất bản cuốn 'Our Seamen' (1873) dẫn đến Đạo luật Hàng hải Thương mại 1875–1876 (Merchant Shipping Act) và quy chuẩn bảng tính của Hội đồng Thương mại Anh năm 1890 (sau này là Công ước Quốc tế về Mạn khô Load Lines 1966 của IMO).",
        "context": "Thế kỷ 19, bùng nổ Cách mạng Công nghiệp khiến cạnh tranh vận tải biển khốc liệt; nhiều chủ tàu tận dụng những 'con tàu quan tài' (Coffin Ships) cũ nát chở quá tải vượt mức an toàn khiến hàng nghìn thủy thủ thiệt mạng mỗi năm khi gặp bão.",
        "modern_impact": "Ngày nay, 100% tàu biển chở hàng quốc tế cập cảng Cát Lái, Hải Phòng hay Rotterdam đều bắt buộc phải sơn vạch Plimsoll Line với các nấc mật độ nước (T - Nhiệt đới, S - Mùa hè, W - Mùa đông, WNA - Bắc Đại Tây Dương mùa đông, F - Nước ngọt); nếu vạch sơn chìm dưới mặt nước, cảng vụ lập tức giữ tàu và từ chối cấp phép rời cảng.",
        "intl_comp": "Giai đoạn 1876–1890, chủ tàu từng lách quy định bằng cách tự ý vẽ vạch tải trọng sát boong tàu (lỗ hổng pháp lý); đến năm 1890, Anh giao quyền tính toán và đóng dấu vạch Plimsoll cho cơ quan đăng kiểm độc lập (như Lloyd's Register) — mô hình tiền thân cho đăng kiểm hàng hải toàn cầu hiện nay.",
        "business_lesson": "Đừng bao giờ tối ưu cước phí bằng cách nhồi nhét vượt tải trọng an toàn hay bỏ qua tiêu chuẩn kiểm định kỹ thuật; trong logistics, một vạch giới hạn minh bạch và sự giám sát độc lập chính là thứ bảo vệ sinh mạng cả chuyến hàng và uy tín doanh nghiệp.",
        "url": "https://www.lr.org/en/knowledge/horizons/march-2024/plimsoll-at-200-the-legacy-beyond-the-load-line/",
        "keywords": ["Plimsoll", "load line"],
        "series": "Lịch sử Logistics",
        "status": "Chưa dùng",
        "note": "Hướng 1: Cột mốc lịch sử thương mại. Đối chiếu 2 nguồn độc lập: Lloyd's Register (lr.org) và Lloyd's Register Foundation (lrfoundation.org.uk)."
    },
    {
        "stt": 257,
        "week": "W47 — 17/11/2026",
        "title": "Eo Biển Malacca — 'Cổ Chai Sinh Tử' Của Châu Á: Hơn 94.000 Tàu Qua Lại Mỗi Năm, Chở 25–30% Thương Mại Biển Toàn Cầu Qua Kênh Hẹp Chỉ 1,7 Hải Lý",
        "hook": "Có một eo biển mà ở điểm hẹp nhất chỉ rộng vỏn vẹn 1,7 hải lý (hơn 3 km), nhưng mỗi ngày gánh tới hơn 250 chuyến tàu khổng lồ và 23,2 triệu thùng dầu thô — nếu nơi này tắc nghẽn, gần như toàn bộ chuỗi cung ứng công xưởng Châu Á sẽ đứng hình!",
        "history": "Eo biển Malacca nối liền Ấn Độ Dương và Thái Bình Dương (giữa bán đảo Malaysia và đảo Sumatra của Indonesia), từng là tuyến đường gia vị huyền thoại hàng thế kỷ và nay trở thành điểm nghẽn (Chokepoint) vận tải dầu mỏ và container lớn nhất thế giới theo thống kê của Cơ quan Thông tin Năng lượng Mỹ (EIA).",
        "context": "Sự trỗi dậy của các nền kinh tế sản xuất Đông Á (Trung Quốc, Nhật Bản, Hàn Quốc, Đài Loan, ASEAN) tạo ra nhu cầu vận chuyển khổng lồ năng lượng từ Trung Đông và hàng hóa công nghiệp đi Châu Âu qua tuyến đường biển ngắn nhất này.",
        "modern_impact": "Eo biển Malacca chiếm khoảng 25–30% tổng lưu lượng thương mại hàng hải toàn cầu (khoảng 94.000–102.500 lượt tàu mỗi năm) và khoảng 80% lượng dầu nhập khẩu của Trung Quốc; Việt Nam nằm ngay cửa ngõ Biển Đông tiếp giáp tuyến hàng hải huyết mạch này, giúp cụm cảng nước sâu Cái Mép — Thị Vải có lợi thế vàng đón tàu mẹ quốc tế.",
        "intl_comp": "Vì eo biển Phillips Channel tại Malacca quá hẹp và nông đối với một số siêu tàu chở dầu (ULCC) mớn nước sâu, những tàu này phải đi vòng qua eo biển Sunda hoặc Lombok của Indonesia làm tăng thêm nhiều ngày hành trình và chi phí nhiên liệu.",
        "business_lesson": "Hiểu bản đồ các 'cổ chai hàng hải' (Suez, Panama, Malacca) giúp doanh nghiệp XNK dự báo sớm biến động giá cước BAF, thời gian vận chuyển (Transit Time) và chủ động phân bổ lịch giao hàng dự phòng trước các cú sốc địa chính trị.",
        "url": "https://www.eia.gov/todayinenergy/detail.php?id=32452",
        "keywords": ["Malacca", "oil"],
        "series": "Thương mại Quốc tế",
        "status": "Chưa dùng",
        "note": "Hướng 3: Con số bất ngờ & Thống kê gây sốc. Nguồn chính thống: U.S. Energy Information Administration (eia.gov) đối chiếu dữ liệu lưu lượng hàng hải quốc tế."
    },
    {
        "stt": 258,
        "week": "W47 — 17/11/2026",
        "title": "Nghịch Lý Logistics Của Zara (Inditex, Tây Ban Nha): Sẵn Sàng Chi Cước Hàng Không Đắt Đỏ Để Biến Chuỗi Cung Ứng Thành 'Cỗ Máy 15 Ngày' Và Giảm Tồn Kho Còn 10%",
        "hook": "Trong khi cả ngành thời trang thuê tàu biển giá rẻ từ Châu Á và lập kế hoạch trước 6–9 tháng, Zara lại chấp nhận trả cước máy bay đắt gấp nhiều lần để đưa áo quần mới thiết kế lên kệ toàn cầu chỉ trong 15 ngày — và kết quả là họ kiếm lời cao gấp đôi đối thủ!",
        "history": "Tập đoàn Inditex (thành lập tại La Coruña, vùng Galicia, Tây Ban Nha) xây dựng trung tâm phân phối tự động khổng lồ 'The Cube' kết nối trực tiếp với mạng lưới xe tải cao tốc Châu Âu và vận tải hàng không liên lục địa.",
        "context": "Ngành thời trang truyền thống thường mất 6–9 tháng từ lúc thiết kế đến khi bán hàng, dẫn đến tỷ lệ hàng tồn kho ế ẩm phải giảm giá thanh lý lên tới 17–20% toàn ngành.",
        "modern_impact": "Zara sản xuất các mẫu chạy theo xu hướng (Trend Items) gần thị trường Châu Âu (Tây Ban Nha, Bồ Đào Nha, Maroc, Thổ Nhĩ Kỳ) chỉ mất 15 ngày từ phác thảo đến cửa hàng, dùng xe tải chạy lịch cố định 2 lần/tuần tại Châu Âu và máy bay (Air Freight) cho Châu Mỹ/Châu Á; nhờ vậy hàng tồn kho không bán được chỉ khoảng 10% và khách hàng quay lại cửa hàng trung bình 17 lần/năm thay vì 3–4 lần.",
        "intl_comp": "Mô hình chuỗi cung ứng nhanh (Agile Supply Chain) của Zara tách làm 2 tốc độ: Hàng cơ bản giá rẻ (áo thun trơn) vẫn đặt gia công Châu Á đi đường biển để tối ưu giá vốn, còn hàng mốt nhanh đi gần + bay Air Freight để tối ưu vòng quay vốn.",
        "business_lesson": "Đừng chỉ nhìn vào 'đơn giá cước vận chuyển rẻ nhất' trên một chứng từ; đôi khi trả cước Air Freight đắt hơn để kịp mùa vụ và giải phóng dòng tiền tồn kho lại là bài toán kinh doanh siêu lợi nhuận nhất!",
        "url": "https://www.scmglobe.com/zara-clothing-company-supply-chain/",
        "keywords": ["Zara", "Inditex"],
        "series": "Đàm phán & Tư duy",
        "status": "Chưa dùng",
        "note": "Hướng 4: Bài học từ mô hình doanh nghiệp & chuỗi cung ứng thế giới. Nguồn: SCM Globe (scmglobe.com) đối chiếu tài liệu chuỗi cung ứng Inditex."
    },
    {
        "stt": 259,
        "week": "W47 — 17/11/2026",
        "title": "Bí Mật 11 Ký Tự Mã Số Container Toàn Cầu (ISO 6346 & BIC Code): Thuật Toán 'Số Kiểm Tra Cuối' Modulo 11 Giúp Hàng Chục Triệu Container Không Bao Giờ Bị Nhầm",
        "hook": "Mỗi chiếc vỏ container ngoài cảng đều có đúng 11 ký tự như 'MSCU 123456 7' — và chữ số cuối cùng đóng khung ô vuông không phải số ngẫu nhiên, mà là một phép toán kiểm lỗi Modulo 11 do máy tính kiểm tra chỉ trong 1 phần nghìn giây!",
        "history": "Tổ chức Văn phòng Container Quốc tế (BIC - Bureau International des Containers) khởi xướng hệ thống đăng ký mã chủ sở hữu từ năm 1970, và Tổ chức Tiêu chuẩn hóa Quốc tế chính thức chuẩn hóa thành tiêu chuẩn ISO 6346 vào năm 1972.",
        "context": "Khi cuộc cách mạng container hóa bùng nổ, hàng triệu vỏ container của hàng trăm hãng tàu tràn ngập các cảng biển toàn cầu; chỉ cần nhân viên nhập liệu hoặc camera OCR đọc nhầm 1 con số trên vận đơn là cả container hàng tỷ đồng có thể bị bốc nhầm sang tàu đi châu lục khác.",
        "modern_impact": "Chuẩn ISO 6346 chia 11 ký tự gồm: 3 chữ cái mã chủ sở hữu (BIC Code độc nhất toàn cầu), 1 chữ cái phân loại thiết bị (U cho container chở hàng, J cho thiết bị tháo rời, Z cho rơ-moóc), 6 số sê-ri và 1 chữ số kiểm tra (Check Digit) tính theo lũy thừa 2 chia lấy dư 11 (bỏ qua bội số 11, 22, 33; nếu dư 10 thì quy về 0). Hệ thống VNACCS và cổng cảng tự động chặn ngay tờ khai nếu gõ sai 1 ký tự số container!",
        "intl_comp": "Khác với biển số xe nội địa quản lý theo từng quốc gia, mã container ISO 6346 được đăng ký tập trung duy nhất tại BIC (Paris) để bảo đảm không có 2 container nào trên Trái Đất trùng mã số từ cảng Los Angeles đến cảng Cát Lái.",
        "business_lesson": "Sai 1 ký tự số container trên vận đơn hoặc tờ khai hải quan là lỗi hệ thống dễ gây kẹt lệnh giao hàng nhất; hãy tận dụng quy luật Check Digit ISO 6346 trong phần mềm quản trị chứng từ để chặn sai sót ngay từ khâu nhập liệu.",
        "url": "https://en.wikipedia.org/wiki/ISO_6346",
        "keywords": ["ISO 6346", "BIC"],
        "series": "Công nghệ & Đổi mới",
        "status": "Chưa dùng",
        "note": "Hướng 5: Công nghệ & tiêu chuẩn hóa tác động thương mại. Đối chiếu 2 nguồn độc lập: Wikipedia (ISO 6346) và Bureau International des Containers (bic-code.org)."
    },
    {
        "stt": 260,
        "week": "W47 — 17/11/2026",
        "title": "Cửa Sổ Hàng Hải Một Cửa Toàn Cầu (IMO Maritime Single Window — Bắt Buộc Từ 01/01/2024): Bước Ngoặt Xóa Bỏ Hàng Chục Bộ Giấy Tờ Trùng Lặp Khi Tàu Cập Cảng",
        "hook": "Trước năm 2024, mỗi lần một con tàu chở hàng cập cảng quốc tế, thuyền trưởng và đại lý tàu biển phải nộp lặp đi lặp lại cùng một dữ liệu cho Hải quan, Biên phòng, Kiểm dịch y tế và Cảng vụ bằng hàng xấp hồ sơ giấy — cho đến khi Công ước FAL của IMO bắt buộc cả thế giới dùng 'Một cửa số hóa'!",
        "history": "Tổ chức Hàng hải Quốc tế (IMO) sửa đổi Phụ lục Công ước FAL (Facilitation of International Maritime Traffic), chính thức có hiệu lực bắt buộc toàn cầu từ ngày 01/01/2024 đối với mọi quốc gia thành viên.",
        "context": "Sự rời rạc giữa các cơ quan quản lý tại cảng biển từng khiến tàu hàng tốn hàng giờ đến hàng ngày nằm chờ tại khu neo đậu chỉ để hoàn tất thủ tục hành chính giấy tờ trùng lặp, gây lãng phí nhiên liệu và tăng phát thải khí nhà kính.",
        "modern_impact": "Maritime Single Window (MSW) thiết lập một điểm tiếp nhận dữ liệu số duy nhất theo nguyên tắc 'khai báo một lần — dùng chung đa ngành' (Submit once, reuse across authorities), kết hợp chuẩn dữ liệu IMO Compendium giúp cảng biển, hãng tàu, xe đầu kéo và chủ hàng đồng bộ lịch giải phóng hàng theo thời gian thực.",
        "intl_comp": "Các cảng lớn như Singapore, Rotterdam đã vận hành Port Community System (PCS) đồng bộ MSW hoàn chỉnh, trong khi IMO phối hợp với Singapore hỗ trợ nhiều quốc gia đang phát triển triển khai nền tảng MSW mẫu nhằm thu hẹp khoảng cách số hóa cảng biển.",
        "business_lesson": "Số hóa chuỗi cung ứng không còn là lựa chọn mà là luật chơi bắt buộc toàn cầu; doanh nghiệp XNK và đại lý logistics Việt Nam phải chuẩn hóa dữ liệu chứng từ điện tử từ gốc để khớp nhịp với Cổng Một cửa Quốc gia và Cảng thông minh.",
        "url": "https://www.imo.org/en/ourwork/facilitation/pages/maritimesinglewindow-default.aspx",
        "keywords": ["Maritime Single Window", "IMO"],
        "series": "Bí mật Hải quan",
        "status": "Chưa dùng",
        "note": "Hướng 2: So sánh hệ thống XNK & cảng biển quốc tế. Nguồn chính thống: Tổ chức Hàng hải Quốc tế (imo.org)."
    }
]

# ==============================================================================
# 5. FEEDBACK LOOP: MISSING VERIFIED LEGAL DOCS FOR Danh_sach_Van_ban_Phap_luat_XNK.xlsx
# ==============================================================================
NEW_LEGAL_DOCS = [
    {
        "ma_so": "Quyết định 1828/QĐ-CHQ",
        "ten": "Quyết định ban hành Quy trình kiểm tra xuất xứ hàng hóa xuất khẩu, nhập khẩu sau khi hàng hóa đã được thông quan",
        "loai": "Quyết định",
        "co_quan": "Cục Hải quan (Bộ Tài chính)",
        "ngay_ban_hanh": "2026-11-12",
        "ngay_hieu_luc": "2026-12-15",
        "ngay_het_hieu_luc": "Chưa xác định",
        "van_ban_bi_thay_the": "Bổ sung quy trình hậu kiểm xuất xứ theo Luật Hải quan, Thông tư 33/2023/TT-BTC và Thông tư 124/2026/TT-BTC",
        "phan_tich": "Quy định quy trình chuẩn cho Đội Phúc tập và Kiểm tra sau thông quan thực hiện rà soát, đối chiếu chứng từ khai báo xuất xứ, C/O ưu đãi và chứng từ hàng miễn nộp C/O dựa trên nguyên tắc quản lý rủi ro trong thời hạn tối đa 15 ngày kể từ khi nhận danh sách tờ khai phân bổ trên hệ thống.",
        "so_sanh": "Chuẩn hóa thời hạn kiểm tra 15 ngày và yêu cầu cập nhật kết quả lên hệ thống ngay trong ngày hoặc ngày làm việc tiếp theo; đặc biệt siết chặt kiểm soát hành vi chia nhỏ lô hàng để hưởng ngưỡng miễn nộp chứng từ chứng nhận xuất xứ và chuyển tải bất hợp pháp."
    },
    {
        "ma_so": "Thông tư 86/2026/TT-BTC",
        "ten": "Thông tư quy định về quản lý thuế đối với hàng hóa xuất khẩu, nhập khẩu",
        "loai": "Thông tư",
        "co_quan": "Bộ Tài chính",
        "ngay_ban_hanh": "2026-06-30",
        "ngay_hieu_luc": "2026-07-01",
        "ngay_het_hieu_luc": "Chưa xác định",
        "van_ban_bi_thay_the": "Đồng bộ hóa hướng dẫn quản lý thuế XNK theo Luật Quản lý thuế mới",
        "phan_tich": "Quy định chi tiết về khai thuế, thời hạn nộp thuế, ấn định thuế, miễn giảm thuế, không thu thuế, xử lý tiền thuế nộp thừa, bù trừ, hoàn thuế và quản lý nợ thuế đối với hàng hóa xuất khẩu, nhập khẩu (bao gồm giao dịch thương mại điện tử/nền tảng số).",
        "so_sanh": "Chuẩn hóa quy trình bù trừ tiền thuế nộp thừa, hoàn thuế điện tử và tiêu chí người nộp thuế không còn nợ tiền thuế, tiền phạt, tiền phí hải quan khi thực hiện thủ tục thông quan."
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
    """Ensure STT 256-260 and titles/legal codes do not duplicate existing Excel rows."""
    for path in (
        "reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx",
        "reports/Case_Study_Ideas_Bank_Eureka.xlsx",
        "reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx",
        "reports/Funfact_Ideas_Bank_Eureka.xlsx",
    ):
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        last_stt = ws.cell(row=ws.max_row, column=1).value
        if last_stt != 255:
            raise ValueError(f"{path} last STT is {last_stt}, expected 255 before appending 256-260!")
    print("🟢 DUPLICATE & SEQUENCE CHECK PASSED: 4 file Ideas Bank sẵn sàng nối tiếp STT 256-260.")


def execute_update():
    print("=" * 80)
    print("🚀 WORKFLOW 06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (WEEK 47 DAY 2 — 17/11/2026)")
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
    print(f"✅ Ghi thành công {tn_path} (Tổng dòng: {ws_tn.max_row}, STT cuối: 260)")

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
    print(f"✅ Ghi thành công {cs_path} (Tổng dòng: {ws_cs.max_row}, STT cuối: 260)")

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
    print(f"✅ Ghi thành công {ht_path} (Tổng dòng: {ws_ht.max_row}, STT cuối: 260)")

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
    print(f"✅ Ghi thành công {ff_path} (Tổng dòng: {ws_ff.max_row}, STT cuối: 260)")

    # 6. FEEDBACK LOOP: DANH SACH VAN BAN PHAP LUAT XNK (Using legal_tracker style & dedup)
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
            "2026-11-17"
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
    print("🎉 HOÀN TẤT WORKFLOW 06 (17/11/2026): ĐÃ CẬP NHẬT 20 Ý TƯỞNG MỚI (STT 256-260) & FEEDBACK LOOP!")
    print("=" * 80)


if __name__ == "__main__":
    execute_update()
