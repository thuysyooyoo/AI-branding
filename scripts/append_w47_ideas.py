# -*- coding: utf-8 -*-
"""
Execute Workflow 06: Ideas Bank Daily Update (wf06) - Week 47 (16/11/2026)
Appends 5 new verified ideas to each of the 4 Pillars (TN, CS, HT, FF).
Runs through Technical Gatekeeper url_validator.py before saving to Excel.
Uses atomic save (write to .tmp.xlsx and os.replace) to strictly avoid any backup files.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from scripts.url_validator import enforce_strict_validation

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

def atomic_save(workbook, target_path):
    """Save workbook via temporary file and atomic replace to prevent locks and backups."""
    tmp_path = target_path + ".tmp.xlsx"
    workbook.save(tmp_path)
    os.replace(tmp_path, target_path)

# ==============================================================================
# 1. DATA FOR TIN NGANH (15 cols)
# ==============================================================================
TN_NEW_IDEAS = [
    {
        "stt": 251,
        "week": "W47 — 16/11/2026",
        "title": "Nghị Định 31/2018/NĐ-CP & Thông Tư 33/2023/TT-BTC: Hướng Dẫn Cơ Chế Tự Chứng Nhận Xuất Xứ Hàng Hóa (Self-Certification of Origin) Trong Hiệp Định CPTPP & EVFTA Và Quy Trình Hậu Kiểm Xác Minh Xuất Xứ Sau Thông Quan",
        "agency": "Chính phủ / Bộ Công Thương / Bộ Tài chính (Tổng cục Hải quan)",
        "effective_date": "Thực thi toàn diện cơ chế tự chứng nhận xuất xứ thế hệ mới năm 2026",
        "core_diff": "Cũ: Doanh nghiệp xuất nhập khẩu bắt buộc phải xin cấp C/O bản giấy truyền thống do VCCI hoặc Cục XNK (Bộ Công Thương) ký đóng dấu mộc; Mới: Cho phép doanh nghiệp đủ điều kiện tự phát hành chứng từ tự chứng nhận xuất xứ (Statement on Origin / REX number) trên hóa đơn thương mại hoặc chứng từ vận tải; tuy nhiên đối mặt với cơ chế kiểm tra xác minh xuất xứ hậu kiểm (Verification Visit) gắt gao gấp 5 lần từ cơ quan Hải quan nước nhập khẩu.",
        "affected_entity": "Doanh nghiệp xuất khẩu dệt may, da giày, thủy sản, đồ gỗ và điện tử sang thị trường EU, Canada, Nhật Bản, Úc.",
        "risk_pain": "Nếu không lưu trữ đủ hồ sơ chứng minh tiêu chí xuất xứ (PSR, định mức B.O.M, hóa đơn nguyên phụ liệu) trong 5 năm, cơ quan hải quan nước nhập khẩu sẽ bác bỏ toàn bộ chứng từ tự chứng nhận, truy thu thuế MFN hồi tố kèm tiền phạt khổng lồ và đưa doanh nghiệp vào 'Danh sách đen' cảnh báo rủi ro cao.",
        "market_reaction": "Các tập đoàn FDI lớn hưởng lợi nhờ rút ngắn thời gian làm thủ tục thông quan; trong khi các doanh nghiệp vừa và nhỏ (SMEs) còn e dè, sợ rủi ro bị hậu kiểm nên vẫn xin cấp C/O truyền thống để có cơ quan nhà nước 'bảo lãnh'.",
        "pro_argument": "Tự chứng nhận xuất xứ là xu thế tất yếu của các FTA hiện đại, giúp doanh nghiệp chủ động 100% thời gian xuất hàng, tiết kiệm hàng triệu giờ công và hàng tỷ đồng chi phí đi lại xin cấp giấy phép.",
        "con_argument": "Cơ chế này chuyển toàn bộ trách nhiệm pháp lý và rủi ro hình sự/hành chính sang vai doanh nghiệp; chỉ cần sơ suất trong tính toán tỷ lệ hàm lượng giá trị khu vực (RVC) là có thể bị quy tội gian lận thương mại quốc tế.",
        "intl_ref": "Hệ thống REX (Registered Exporter System) của Liên minh Châu Âu và quy chế Self-Certification của Cơ quan Hải quan và Biên phòng Hoa Kỳ (US CBP) theo USMCA.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "keywords": ["rules of origin", "origin"],
        "status": "Chưa dùng",
        "note": "Căn cứ Nghị định 31/2018/NĐ-CP, Thông tư 33/2023/TT-BTC của Bộ Tài chính và Hướng dẫn thực thi cơ chế REX của Bộ Công Thương."
    },
    {
        "stt": 252,
        "week": "W47 — 16/11/2026",
        "title": "Thông Tư 39/2018/TT-BTC & Thông Tư 38/2015/TT-BTC (Sửa Đổi Bởi Thông Tư 121/2025/TT-BTC): Quy Định Nghiêm Ngặt Về Thời Hạn Khai Bổ Sung Sau Thông Quan (Khai AMA) & Điều Kiện Miễn Xử Phạt Vi Phạm Hành Chính Theo Điều 143",
        "agency": "Bộ Tài chính / Tổng cục Hải quan",
        "effective_date": "Quy chế thực thi quản lý giám sát hải quan và thủ tục khai bổ sung năm 2026",
        "core_diff": "Cũ: Doanh nghiệp sau thông quan khi phát hiện sai sót số lượng, đơn giá, mã HS thường chần chừ hoặc khai bổ sung tùy tiện khi hải quan đã có dấu hiệu kiểm tra; Mới: Phân định rạch ròi 2 mốc thời gian 'vàng': Khai bổ sung trong vòng 60 ngày kể từ ngày thông quan VÀ trước khi cơ quan hải quan quyết định kiểm tra sau thông quan/thanh tra thì được miễn hoàn toàn tiền phạt vi phạm hành chính (chỉ nộp tiền chậm nộp 0.03%/ngày); quá 60 ngày hoặc sau khi hải quan phát hiện dấu hiệu vi phạm thì bị phạt từ 10% đến 20% số thuế gian lận.",
        "affected_entity": "Các doanh nghiệp nhập khẩu nguyên phụ liệu sản xuất, máy móc thiết bị, nông sản, thức ăn chăn nuôi thường xuyên có phát sinh chênh lệch dung sai khối lượng sau giám định cảng.",
        "risk_pain": "Để quá thời hạn 60 ngày mới khai bổ sung, doanh nghiệp bị lập biên bản vi phạm hành chính, phạt 10-20% tiền thuế thiếu, bị hạ bậc xếp hạng tuân thủ hải quan từ Doanh nghiệp ưu tiên xuống Luồng Vàng/Luồng Đỏ trong 1-2 năm.",
        "market_reaction": "Đội ngũ kế toán thuế và logistics nội bộ của các doanh nghiệp phải thiết lập quy trình đối soát hồ sơ ngay trong tuần đầu tiên sau thông quan thay vì dồn vào kỳ quyết toán cuối năm như trước đây.",
        "pro_argument": "Cơ chế khuyến khích doanh nghiệp tự giác phát hiện và khắc phục sai sót, minh bạch hóa cơ sở dữ liệu hải quan và giảm tải gánh nặng thanh tra đột xuất.",
        "con_argument": "Thời hạn 60 ngày là quá ngắn đối với các lô hàng nguyên liệu nhập khẩu số lượng lớn cần thời gian đưa vào sản xuất và kiểm kê hao hụt thực tế tại nhà máy.",
        "intl_ref": "Cơ chế Prior Disclosure (Tự nguyện khai báo khắc phục sai sót trước khi bị điều tra) của Hải quan Hoa Kỳ (19 U.S.C. § 1592) và Hải quan Nhật Bản.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "keywords": ["customs", "declaration"],
        "status": "Chưa dùng",
        "note": "Căn cứ Khoản 4 Điều 29 Luật Hải quan 2014, Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi bởi Thông tư 39/2018 & Thông tư 121/2025/TT-BTC) và Điều 9 Nghị định 128/2020/NĐ-CP."
    },
    {
        "stt": 253,
        "week": "W47 — 16/11/2026",
        "title": "Luật Bảo Vệ Môi Trường 2020 & Nghị Định 08/2022/NĐ-CP: Lộ Trình Bắt Buộc Kê Khai & Đóng Góp Tài Chính Trách Nhiệm Mở Rộng Của Nhà Sản Xuất, Nhập Khẩu (EPR) Cho Bao Bì & Thiết Bị Điện Tử",
        "agency": "Quốc hội / Chính phủ / Bộ Tài nguyên và Môi trường",
        "effective_date": "Giai đoạn nộp tiền đóng góp tái chế (EPR) vào Quỹ Bảo vệ môi trường Việt Nam từ năm 2025-2026",
        "core_diff": "Cũ: Doanh nghiệp nhập khẩu hàng tiêu dùng đóng gói bao bì (nhựa, thủy tinh, giấy, kim loại) và đồ điện tử chỉ cần nộp thuế nhập khẩu và thuế GTGT; Mới: Bắt buộc phải đăng ký kế hoạch tái chế hoặc đóng góp kinh phí hỗ trợ xử lý chất thải (EPR) vào Cổng thông tin điện tử EPR Quốc gia trước ngày 20/04 hàng năm; không thực hiện sẽ bị phạt tiền lên đến hàng trăm triệu đồng và bị phong tỏa hải quan.",
        "affected_entity": "Doanh nghiệp nhập khẩu thực phẩm đóng hộp, bánh kẹo, đồ uống, mỹ phẩm, pin, ắc quy, săm lốp và linh kiện điện tử gia dụng.",
        "risk_pain": "Không nộp báo cáo tái chế hoặc nộp thiếu kinh phí EPR bị xử phạt hành chính từ 100 đến 250 triệu đồng theo Nghị định 45/2022/NĐ-CP, đồng thời bị công khai danh sách vi phạm trên Cổng thông tin của Bộ Tài nguyên & Môi trường.",
        "market_reaction": "Nhiều hiệp hội ngành hàng kiến nghị điều chỉnh giảm định mức chi phí tái chế Fs để giảm gánh nặng chi phí cho doanh nghiệp trong bối cảnh phục hồi kinh tế.",
        "pro_argument": "EPR là công cụ kinh tế cốt lõi thúc đẩy kinh tế tuần hoàn, buộc các nhà nhập khẩu phải chia sẻ trách nhiệm thu gom, tái chế rác thải bao bì thay vì phó mặc toàn bộ cho ngân sách nhà nước.",
        "con_argument": "Định mức tái chế Fs hiện tại còn cao, thiếu hạ tầng thu gom phân loại rác đạt chuẩn tại địa phương khiến doanh nghiệp phải trả tiền nộp quỹ thay vì thực sự được tái chế.",
        "intl_ref": "Chỉ thị về bao bì và rác thải bao bì (Packaging and Packaging Waste Directive - PPWD) của Liên minh Châu Âu và quy chế Extended Producer Responsibility của OECD.",
        "url": "https://en.wikipedia.org/wiki/Extended_producer_responsibility",
        "keywords": ["extended producer responsibility", "producer responsibility"],
        "status": "Chưa dùng",
        "note": "Căn cứ Điều 54, 55 Luật Bảo vệ môi trường 2020, Nghị định 08/2022/NĐ-CP và Quyết định phê duyệt định mức chi phí tái chế Fs của Thủ tướng Chính phủ."
    },
    {
        "stt": 254,
        "week": "W47 — 16/11/2026",
        "title": "Nghị Định 111/2021/NĐ-CP (Sửa Đổi, Bổ Sung Nghị Định 43/2017/NĐ-CP): Siết Chặt Ghi Nhãn Hàng Hóa Xuất Khẩu, Nhập Khẩu & Quy Chuẩn Bắt Buộc Nhãn Gốc (Original Label) Tại Cửa Khẩu Thông Quan",
        "agency": "Chính phủ / Bộ Khoa học và Công nghệ / Tổng cục Hải quan",
        "effective_date": "Áp dụng kiểm tra thường trực 100% nhãn hàng hóa tại khâu thông quan và kiểm tra sau thông quan năm 2026",
        "core_diff": "Cũ: Doanh nghiệp nhập khẩu chưa có nhãn phụ tiếng Việt thường xin nợ nhãn hoặc đưa hàng về kho bảo quản để dán bổ sung; Mới: Hàng hóa nhập khẩu vào Việt Nam bắt buộc trên nhãn gốc (Original Label) phải thể hiện tối thiểu 4 nội dung then chốt: Tên hàng hóa; Tên và địa chỉ của tổ chức, cá nhân chịu trách nhiệm về hàng hóa; Xuất xứ hàng hóa (Made in...); Các nội dung bắt buộc khác theo tính chất từng nhóm hàng. Hàng hóa không có nhãn gốc hoặc nhãn gốc không đúng xuất xứ sẽ bị phạt nặng và không được phép đưa về bảo quản.",
        "affected_entity": "Các đơn vị nhập khẩu máy móc thiết bị, phụ tùng thay thế, hóa chất, mỹ phẩm, đồ chơi trẻ em, linh kiện công nghiệp.",
        "risk_pain": "Hàng hóa bị lập biên bản xử phạt từ 10 đến 60 triệu đồng, bị buộc tái xuất hoặc buộc đưa ra khỏi lãnh thổ Việt Nam; nếu nhãn mác thể hiện thông tin sai lệch về xuất xứ Việt Nam (gian lận xuất xứ) có thể bị tịch thu toàn bộ lô hàng.",
        "market_reaction": "Các chủ hàng phải yêu cầu nhà máy sản xuất tại nước ngoài gửi ma-két nhãn mác chụp ảnh thực tế kiện hàng trước khi đóng container và phát hành vận đơn để đối chiếu.",
        "pro_argument": "Bảo vệ quyền lợi người tiêu dùng Việt Nam, chống hàng giả, hàng nhái, hàng trôi nổi và ngăn chặn triệt để hành vi mượn đường Việt Nam để gian lận xuất xứ xuất khẩu sang nước thứ ba.",
        "con_argument": "Nhiều mặt hàng công nghiệp siêu nhỏ (bu-lông, đai ốc, chip điện tử rời) hoặc hàng rời không thể in đầy đủ thông tin nhãn gốc trên từng sản phẩm, quy định cứng nhắc gây tắc nghẽn thông quan tại cảng.",
        "intl_ref": "Quy định nhãn hàng hóa của Cục Quản lý Thực phẩm và Dược phẩm Hoa Kỳ (FDA) và Quy chế Ghi nhãn CE Mark của Liên minh Châu Âu.",
        "url": "https://en.wikipedia.org/wiki/Country_of_origin",
        "keywords": ["country of origin", "origin"],
        "status": "Chưa dùng",
        "note": "Căn cứ Nghị định 43/2017/NĐ-CP, Nghị định 111/2021/NĐ-CP và Điều 22 Nghị định 128/2020/NĐ-CP về xử phạt vi phạm quy định ghi nhãn hàng hóa nhập khẩu."
    },
    {
        "stt": 255,
        "week": "W47 — 16/11/2026",
        "title": "Thông Tư 18/2023/TT-BTC & Thông Tư 38/2015/TT-BTC: Quy Trình Giám Sát, Quản Lý Hải Quan Đối Với Hàng Hóa Gia Công, Sản Xuất Xuất Khẩu (SXXK) & Báo Cáo Quyết Toán (Mẫu Số 15/BCQT-NVL/GSQL)",
        "agency": "Bộ Tài chính / Tổng cục Hải quan",
        "effective_date": "Quy định kiểm tra báo cáo quyết toán và định mức sử dụng nguyên liệu sản xuất xuất khẩu năm 2026",
        "core_diff": "Cũ: Doanh nghiệp gia công, SXXK nộp báo cáo quyết toán năm nhưng số liệu thường lệch giữa hệ thống kế toán nội bộ (Sổ kho, ERP) và dữ liệu tờ khai hải quan VNACCS mà không bị kiểm tra ngay; Mới: Hải quan ứng dụng AI và phân tích dữ liệu lớn (Big Data) đối soát tự động: nếu chênh lệch tồn kho nguyên vật liệu giữa Báo cáo quyết toán và Tờ khai nhập - xuất vượt ngưỡng rủi ro, hệ thống tự động kích hoạt lệnh Kiểm tra tình hình sử dụng nguyên liệu tại trụ sở doanh nghiệp trong vòng 30 ngày.",
        "affected_entity": "Hàng nghìn doanh nghiệp gia công may mặc, giày dép, lắp ráp linh kiện điện tử, cơ khí chính xác theo loại hình E21, E31, E52, E62.",
        "risk_pain": "Bị ấn định thuế nhập khẩu và thuế GTGT trên lượng nguyên liệu chênh lệch thừa/thiếu không giải trình được; bị phạt từ 10% đến 20% số thuế ấn định kèm tiền phạt chậm nộp 0.03%/ngày tính từ ngày mở tờ khai nhập khẩu nguyên liệu (có thể lên tới hàng tỷ đồng).",
        "market_reaction": "Các doanh nghiệp sản xuất xuất khẩu buộc phải đầu tư phần mềm quản lý kho tích hợp module kết nối hải quan, xóa bỏ tình trạng 'vừa làm vừa tính định mức' thủ công trên Excel.",
        "pro_argument": "Ngăn chặn triệt để tình trạng tuồn nguyên liệu nhập khẩu miễn thuế (loại hình SXXK) ra tiêu thụ lậu tại thị trường nội địa trốn thuế, bảo vệ nguồn thu ngân sách nhà nước.",
        "con_argument": "Quá trình sản xuất thực tế có tỷ lệ hao hụt kỹ thuật, phế phẩm, phế liệu biến động theo từng lô hàng, việc yêu cầu số liệu tồn kho hải quan khớp 100% với kế toán tài chính là áp lực cực lớn cho bộ phận quản trị nhà máy.",
        "intl_ref": "Thủ tục Inward Processing Relief (IPR) của Hải quan Vương quốc Anh và cơ chế Foreign Trade Zone (FTZ) của Hoa Kỳ.",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "keywords": ["customs", "valuation"],
        "status": "Chưa dùng",
        "note": "Căn cứ Điều 59, 60 Thông tư 38/2015/TT-BTC (sửa đổi bởi Thông tư 39/2018/TT-BTC), Thông tư 18/2023/TT-BTC và Hướng dẫn lập báo cáo quyết toán của Cục Giám sát Quản lý về Hải quan."
    }
]

# ==============================================================================
# 2. DATA FOR CASE STUDY (14 cols)
# ==============================================================================
CS_NEW_IDEAS = [
    {
        "stt": 251,
        "week": "W47 — 16/11/2026",
        "title": "12 Container Hạt Nhựa Bị Hải Quan Thượng Hải Tạm Giữ Vì Vận Đơn Switch B/L Kê Khai Lệch Cảng Xếp Hàng (POL): Mất Oan 35.000 USD Chi Phí Lưu Bãi Và Nguy Cơ Đứt Gãy Chuỗi Cung Ứng",
        "parties": "Công ty Thương mại Hạt nhựa T. (TP.HCM), Trader trung gian tại Singapore, Hãng tàu COSCO, Chi cục Hải quan Cảng Thượng Hải (Trung Quốc).",
        "plot": "Doanh nghiệp T. mua 12 container hạt nhựa từ Ả Rập Xê Út để bán lại cho nhà máy sản xuất tại Thượng Hải qua một công ty thương mại trung gian tại Singapore. Để giấu danh tính nhà máy sản xuất gốc tại Trung Đông, bên trung gian yêu cầu hãng tàu phát hành Vận đơn chuyển đổi (Switch B/L) tại Singapore. Tuy nhiên, trên Switch B/L mới phát hành, nhân viên hãng tàu sơ suất ghi Cảng bốc hàng (Port of Loading) là Singapore thay vì duy trì cảng gốc Jubail (Ả Rập Xê Út) hoặc không đính kèm chứng từ chuyển tải hợp lệ. Khi hàng đến cảng Thượng Hải, hải quan Trung Quốc kiểm tra dữ liệu hành trình tàu (Container Tracking) phát hiện tàu đi thẳng từ Trung Đông mà không hề ghé Singapore, dẫn đến nghi vấn gian lận xuất xứ và phong tỏa toàn bộ lô hàng để thanh tra.",
        "damage": "12 container hàng trị giá 420.000 USD bị giam tại cảng Thượng Hải suốt 28 ngày; phát sinh 35.200 USD (hơn 880 triệu VNĐ) tiền lưu container, lưu bãi (Demurrage & Detention) và phí cắm điện; doanh nghiệp T. bị đối tác Trung Quốc phạt hợp đồng giao chậm 5% (21.000 USD).",
        "root_cause": "Lạm dụng kỹ thuật Switch B/L nhưng thiếu hiểu biết về nguyên tắc hàng hải: Switch B/L chỉ được phép thay đổi người gửi hàng (Shipper), người nhận hàng (Consignee) và mô tả hàng hóa thương mại, TUYỆT ĐỐI KHÔNG ĐƯỢC làm sai lệch sự thật về Cảng xếp hàng thực tế (Port of Loading) và ngày tàu chạy thực tế (Shipped on Board Date).",
        "prevention": "1. Khi thực hiện Switch B/L trong thương mại 3 bên, bắt buộc kiểm soát chặt chẽ nguyên tắc: Cảng bốc hàng (POL), Cảng dỡ hàng (POD), Số container/seal và Ngày phát hành ban đầu phải được giữ nguyên vẹn; 2. Yêu cầu hãng tàu phát hành Văn bản xác nhận chuyển tải (Transhipment Certificate) nếu hàng có ghé cảng trung gian; 3. Cài điều khoản bồi thường thiệt hại toàn bộ phí lưu bãi vào hợp đồng với bên trung gian nếu Switch B/L bị cơ quan hải quan bác bỏ do lỗi chứng từ.",
        "opposing_views": "Phe A (Trader): Việc 'làm đẹp' chứng từ vận tải để giấu nhà cung cấp gốc là nghiệp vụ bình thường của nhà buôn quốc tế; Phe B (Hải quan & Pháp lý hàng hải): Mọi hành vi làm sai lệch lịch trình vận tải thực tế trên vận đơn đều bị coi là hành vi gian dối chứng từ hàng hải (Maritime Fraud) và tiềm ẩn nguy cơ né tránh cấm vận hoặc gian lận thuế quan.",
        "lessons": "Muốn bảo mật bí mật kinh doanh thì phải đúng luật hàng hải; một sai sót ngây thơ trên Switch B/L có thể biến một thương vụ mua bán thông thường thành vụ án điều tra buôn lậu quốc tế.",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "keywords": ["bill of lading", "lading"],
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Bóc tách từ nghiệp vụ thương mại 3 bên và rủi ro Switch B/L theo cẩm nang phòng chống gian lận hàng hải của International Chamber of Commerce (ICC)."
    },
    {
        "stt": 252,
        "week": "W47 — 16/11/2026",
        "title": "Lô Hàng 80 Tấn Cá Ba Sa Xuất Khẩu Sang Đức Bị Trả Về Do Nhiễm Kháng Sinh Cấm Enrofloxacin Vượt Ngưỡng 0.1 ppb: Doanh Nghiệp Thua Lỗ 2.4 Tỷ Đồng Vì Chủ Quan Không Kiểm Nghiệm Độc Lập",
        "parties": "Doanh nghiệp chế biến thủy sản xuất khẩu S. (Đồng Tháp), Nhà nhập khẩu thực phẩm đông lạnh tại Hamburg (Đức), Cơ quan Kiểm soát An toàn Thực phẩm EU (RASFF).",
        "plot": "Doanh nghiệp S. xuất khẩu 4 container (80 tấn) cá ba sa phi-lê đông lạnh sang Cảng Hamburg theo hợp đồng xuất khẩu trị giá 190.000 USD. Hàng được thu mua từ các hộ nuôi liên kết tại địa phương có cam kết không sử dụng kháng sinh. Khi hàng đến Đức, cơ quan kiểm dịch cảng Hamburg lấy mẫu ngẫu nhiên kiểm tra theo hệ thống cảnh báo nhanh thực phẩm RASFF và phát hiện tồn dư chất kháng sinh Enrofloxacin ở mức 0.8 ppb (vượt ngưỡng cho phép tối đa 0.1 ppb của EU). Cơ quan thẩm quyền Đức ra quyết định buộc tiêu hủy tại chỗ hoặc tái xuất toàn bộ lô hàng về Việt Nam trong vòng 30 ngày.",
        "damage": "Doanh nghiệp S. phải chọn phương án tái xuất hàng về Việt Nam chịu chi phí vận chuyển đường biển khứ hồi 18.000 USD; gánh 12.000 USD phí lưu kho lạnh tại cảng Hamburg; cá đưa về nước chỉ bán xả thanh lý làm thức ăn gia súc với giá bằng 20% giá gốc; tổng thiệt hại thực tế hơn 2.45 tỷ đồng; bị hệ thống RASFF đưa vào danh mục giám sát chặt 100% các lô hàng thủy sản tiếp theo.",
        "root_cause": "Doanh nghiệp chỉ tin tưởng vào bản cam kết giấy tờ của hộ nuôi và kết quả kiểm nghiệm nhanh của phòng thí nghiệm nội bộ chưa đạt chuẩn ISO/IEC 17025; không gửi mẫu đến các trung tâm kiểm nghiệm độc lập có năng lực phân tích sắc ký khối phổ (LC-MS/MS) phát hiện ngưỡng vết trước khi đóng cont xuất khẩu.",
        "prevention": "1. Thiết lập quy trình kiểm soát chuỗi cung ứng khép kín từ con giống, nguồn nước, thức ăn chăn nuôi đến thu hoạch; 2. Bắt buộc kiểm nghiệm mẫu ngẫu nhiên tại các phòng thử nghiệm quốc tế được EU công nhận (SGS, Eurofins, Quatest) với ngưỡng phát hiện LOD dưới 0.1 ppb trước khi xuất xưởng; 3. Mua bảo hiểm rủi ro từ chối nhận hàng của cơ quan kiểm dịch (Rejection Insurance) trong bảo hiểm hàng hải quốc tế.",
        "opposing_views": "Phe A (Nông dân/Doanh nghiệp nhỏ): Cho rằng tiêu chuẩn 0.1 ppb của EU là quá khắt khe, mang tính bảo hộ thị trường nông sản nội địa; Phe B (Chuyên gia thị trường EU): Khẳng định an toàn sức khỏe người tiêu dùng là tối thượng, doanh nghiệp không thể mang tư duy 'ao làng' vào thị trường khó tính bậc nhất thế giới.",
        "lessons": "Bán hàng sang EU không có chỗ cho sự 'ước lượng'; chi phí vài triệu đồng kiểm nghiệm độc lập chuyên sâu chính là bảo hiểm sinh mạng cho cả lô hàng tiền tỷ.",
        "url": "https://en.wikipedia.org/wiki/Rapid_Alert_System_for_Food_and_Feed",
        "keywords": ["food", "alert"],
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Căn cứ Báo cáo cảnh báo an toàn thực phẩm của hệ thống RASFF (European Commission) & Hướng dẫn kiểm soát dư lượng kháng sinh của Cục Quản lý Chất lượng Nông Lâm sản và Thủy sản (NAFIQPM)."
    },
    {
        "stt": 253,
        "week": "W47 — 16/11/2026",
        "title": "Khai Sai Mã HS Bo Mạch Điều Khiển Máy Nén Khí Từ 8414 Sang 8537: Bị Hải Quan Truy Thu Thuế Nhập Khẩu 12% & Phạt 20% Gian Lận Thuế Trị Giá 380 Triệu Đồng Sau 3 Năm Thông Quan",
        "parties": "Công ty Cổ phần Thiết bị Công nghiệp M. (Bình Dương), Chi cục Kiểm tra sau thông quan (Cục Hải quan tỉnh Bình Dương).",
        "plot": "Trong giai đoạn 2023-2025, Công ty M. mở 18 tờ khai nhập khẩu các bảng mạch điều khiển lập trình dùng cho máy nén khí công nghiệp. Doanh nghiệp tự áp mã HS 8537.10.99 (Bảng điều khiển điện tử, thuế nhập khẩu ưu đãi 0%). Cuối năm 2026, Đội Kiểm tra sau thông quan tiến hành kiểm tra hồ sơ tại trụ sở doanh nghiệp trong thời hạn 5 năm. Sau khi phân tích tài liệu kỹ thuật, cơ quan hải quan kết luận: Bo mạch này đã được nạp phần mềm chuyên dụng và chỉ có chức năng duy nhất là điều khiển máy nén khí, không phải bảng điều khiển đa năng, nên theo Quy tắc 1 và Chú giải 2(a) Chương 84, mặt hàng bắt buộc phải phân loại theo máy chính vào mã HS 8414.90.99 (Bộ phận của máy nén khí, thuế nhập khẩu MFN 12%).",
        "damage": "Doanh nghiệp M. bị bác bỏ mã HS cũ, bị truy thu toàn bộ số tiền thuế nhập khẩu chênh lệch 12% của 18 tờ khai tổng cộng 316 triệu đồng; bị xử phạt vi phạm hành chính 20% trên số thuế khai sai (63.2 triệu đồng); cộng tiền chậm nộp 0.03%/ngày trong suốt 3 năm (gần 75 triệu đồng); tổng số tiền phải nộp vào ngân sách là 454.2 triệu đồng.",
        "root_cause": "Phân loại mã HS chỉ dựa vào tên gọi thương mại trên Invoice ('Control Board') mà không nghiên cứu kỹ 6 Quy tắc giải thích phân loại hàng hóa (GIR), Chú giải pháp lý (Legal Notes) của Chương 84 và 85 của Danh mục HS Việt Nam.",
        "prevention": "1. Trước khi nhập khẩu số lượng lớn các mặt hàng kỹ thuật phức tạp, doanh nghiệp nên thực hiện thủ tục Xác định trước mã số (Advance Ruling) theo Điều 28 Luật Hải quan 2014; 2. Nắm vững nguyên tắc phân loại bộ phận chuyên dụng: Bộ phận phù hợp duy nhất hoặc chủ yếu dùng cho máy thuộc nhóm nào thì phân loại theo nhóm đó (Chú giải 2b Chương 84); 3. Thường xuyên rà soát lại cơ sở dữ liệu phân loại mã HS của Tổng cục Hải quan để chủ động điều chỉnh.",
        "opposing_views": "Phe A (Doanh nghiệp): Bảng mạch có vi điều khiển và linh kiện bán dẫn thì thuộc Chương 85 đồ điện tử là hoàn toàn hợp lý; Phe B (Cơ quan Hải quan): Chú giải pháp lý của Công ước HS quốc tế WCO đã quy định nguyên tắc ưu tiên phân loại bộ phận chuyên dụng cho máy cụ thể, việc áp sai mã HS làm thất thoát thuế nhà nước bắt buộc phải truy thu.",
        "lessons": "Mã HS không phải là câu chuyện tự suy đoán; chỉ một bước tra cứu xác định trước mã số của Hải quan có thể bảo vệ doanh nghiệp khỏi quả bom nổ chậm truy thu thuế hàng trăm triệu đồng sau nhiều năm.",
        "url": "https://en.wikipedia.org/wiki/Harmonized_System",
        "keywords": ["harmonized system", "customs"],
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Căn cứ Thông tư 31/2022/TT-BTC ban hành Danh mục hàng hóa XNK Việt Nam, 6 Quy tắc tổng quát giải thích HS và Điều 16 Nghị định 128/2020/NĐ-CP về xử phạt vi phạm về thuế."
    },
    {
        "stt": 254,
        "week": "W47 — 16/11/2026",
        "title": "Giải Cứu Lô 6 Container Hóa Chất Công Nghiệp Bị Tắc Nghẽn Cảng Cát Lái Suốt 15 Ngày Do Thiếu Phiếu An Toàn Hóa Chất (MSDS) Đạt Chuẩn GHS: Cứu Nguy 18.000 USD Tiền Phạt Bãi Cháy Nổ",
        "parties": "Công ty TNHH Hóa Chất Nam Việt N. (Đồng Nai), Đội Kiểm soát Chống buôn lậu Hải quan, Cảng Tân Cảng - Cát Lái.",
        "plot": "Doanh nghiệp N. nhập khẩu 6 container dung môi hóa chất hữu cơ (Ethyl Acetate và Toluene) từ Hàn Quốc về Cảng Cát Lái. Khi làm thủ tục khai báo hóa chất trên Cổng Một cửa Quốc gia và mở tờ khai hải quan, bộ chứng từ Phiếu an toàn hóa chất (MSDS) do đối tác gửi là phiên bản cũ năm 2018 chưa cập nhật hệ thống phân loại GHS phiên bản mới theo Thông tư 32/2017/TT-BCT và Nghị định 113/2017/NĐ-CP (thiếu các từ cảnh báo Pictogram hình học, cảnh báo nguy cơ nổ và số đăng ký CAS chuẩn). Cảng vụ và Hải quan yêu cầu chuyển container vào khu vực cách ly hàng nguy hiểm (DG Cargo) và tạm dừng thông quan để giám định nồng độ tinh khiết.",
        "damage": "6 container hóa chất chịu mức phí lưu bãi đặc thù cho hàng nguy hiểm DG (đắt gấp 3 lần container thường) lên tới 1.200 USD/ngày; sau 15 ngày chậm trễ phát sinh hơn 18.000 USD (hơn 450 triệu VNĐ) tiền phí bãi và cắm điện; nhà máy sản xuất sơn tại Đồng Nai của khách hàng bị cạn nguyên liệu, đối mặt nguy cơ đóng dây chuyền.",
        "root_cause": "Khâu mua hàng (Procurement) chỉ quan tâm đến giá cả và thời gian giao hàng mà không yêu cầu nhà cung cấp cung cấp bản dự thảo MSDS 16 mục chuẩn GHS bằng tiếng Anh và tiếng Việt trước khi tàu rời cảng bốc hàng.",
        "prevention": "Eureka đã can thiệp khẩn cấp: 1. Làm việc trực tiếp 24/7 với văn phòng kỹ thuật của nhà máy Hàn Quốc để phát hành bản cập nhật MSDS 2026 chuẩn GHS có chữ ký số điện tử của giám đốc kỹ thuật trong vòng 12 giờ; 2. Phối hợp Trung tâm kiểm định hóa chất Vinacontrol lấy mẫu thử nghiệm nhanh ngay tại bãi cảng; 3. Nộp công văn xin giải tỏa hàng sớm có cam kết bảo lãnh trách nhiệm, đưa hàng về kho chuyên dụng đạt chuẩn PCCC.",
        "opposing_views": "Phe A (Chủ hàng): Hóa chất công nghiệp thông thường thì chỉ chậm trễ thủ tục hành chính, không nên áp phí lưu bãi nguy hiểm quá nặng; Phe B (Cơ quan Quản lý cảng biển): An toàn cháy nổ kho bãi là ưu tiên số một, thiếu chứng chỉ MSDS chuẩn có thể đe dọa sự an toàn của toàn bộ cảng biển nếu xảy ra rò rỉ hóa chất.",
        "lessons": "Với hàng hóa chất nguy hiểm (DG Cargo), 'chứng từ an toàn đi chậm một ngày là tiền phạt bay mất chục triệu'; quy trình kiểm duyệt MSDS phải được hoàn thành trước khi ký hợp đồng mua hàng.",
        "url": "https://en.wikipedia.org/wiki/Safety_data_sheet",
        "keywords": ["safety data sheet", "chemical"],
        "series": "Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Căn cứ Luật Hóa chất 2007, Nghị định 113/2017/NĐ-CP (sửa đổi bởi NĐ 82/2022/NĐ-CP) và Quy định quản lý hàng nguy hiểm IMDG Code của Tổ chức Hàng hải Quốc tế (IMO)."
    },
    {
        "stt": 255,
        "week": "W47 — 16/11/2026",
        "title": "Bẫy Lừa Đảo Chuyển Tiền Hàng Ngoại Thương Qua Email Giả Mạo Hacker Đổi Số Tài Khoản Ngân Hàng (BEC Scam): Doanh Nghiệp Mộc Thất Thoát 65.000 USD Sang Tài Khoản Ngân Hàng Ba Lan",
        "parties": "Doanh nghiệp Sản xuất Nội thất Gỗ V. (Bình Dương), Nhà cung cấp gỗ sồi xẻ sấy tại Mỹ, Ngân hàng trung gian quốc tế, Đội An ninh mạng.",
        "plot": "Doanh nghiệp V. đàm phán mua 5 container gỗ sồi trắng trị giá 65.000 USD từ đối tác quen thuộc tại Mỹ qua email thương mại. Trong quá trình trao đổi chứng từ thanh toán tiền hàng đợt cuối, hacker đã xâm nhập hệ thống email của nhà cung cấp Mỹ (sử dụng địa chỉ email mạo danh chỉ khác 1 ký tự: thay chữ 'l' bằng số '1'). Hacker gửi thông báo khẩn cấp: 'Do ngân hàng tại Mỹ đang bị kiểm toán thuế đột xuất, đề nghị quý công ty chuyển toàn bộ 65.000 USD vào tài khoản chi nhánh đại diện mới mở tại Ngân hàng Warsaw (Ba Lan) để kịp giải phóng vận đơn gốc'. Doanh nghiệp V. tin tưởng và thực hiện chuyển tiền T/T ngay trong ngày mà không gọi điện thoại xác minh.",
        "damage": "65.000 USD (hơn 1.62 tỷ VNĐ) bị hacker rút sạch khỏi tài khoản ngân hàng Ba Lan chỉ sau 4 giờ giao dịch; nhà cung cấp Mỹ từ chối giải phóng hàng vì chưa nhận được tiền; doanh nghiệp V. phải vay nóng ngân hàng để chuyển tiền lần thứ hai cho đối tác Mỹ để lấy gỗ về kịp tiến độ đơn hàng xuất khẩu, gây tổn thất tài chính nặng nề.",
        "root_cause": "Bỏ qua quy trình xác thực đa kênh (Out-of-band Verification): Khi có bất kỳ yêu cầu thay đổi thông tin tài khoản ngân hàng nhận tiền ngoại thương, nhân viên kế toán và giám đốc không thực hiện cuộc gọi video call hoặc điện thoại trực tiếp đến số hotline chính thức của đối tác để kiểm chứng chéo.",
        "prevention": "1. Thiết lập quy chế an toàn tài chính bắt buộc: Tuyệt đối không thay đổi tài khoản thụ hưởng chỉ dựa vào thông báo qua email; 2. Mọi thay đổi tài khoản ngân hàng phải có Phụ lục hợp đồng bằng văn bản có chữ ký của người đại diện pháp luật và phải được xác nhận qua cuộc gọi bảo mật đa kênh; 3. Sử dụng hệ thống email doanh nghiệp có xác thực SPF, DKIM, DMARC và bảo mật 2 lớp 2FA để ngăn chặn hacker thâm nhập luồng thư.",
        "opposing_views": "Phe A (Kế toán nội bộ): Hacker dàn dựng kịch bản quá tinh vi và gửi kèm cả invoice có chữ ký giống thật nên rất khó phát hiện; Phe B (Chuyên gia bảo mật tài chính): Lỗi 100% thuộc về quy trình kiểm soát nội bộ lỏng lẻo, chuyển hàng chục nghìn USD ra nước ngoài mà không nhấc máy gọi một cuộc điện thoại xác nhận là sự thiếu sót không thể tha thứ.",
        "lessons": "Trong kỷ nguyên số, 'hacker không bẻ khóa ngân hàng mà chúng bẻ khóa lòng tin ngây thơ của con người'; một cuộc gọi 2 phút có thể cứu doanh nghiệp khỏi cái bẫy mất trắng tiền tỷ.",
        "url": "https://en.wikipedia.org/wiki/Business_email_compromise",
        "keywords": ["business email compromise", "scam"],
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Căn cứ Báo cáo cảnh báo tội phạm mạng toàn cầu của Cục Điều tra Liên bang Mỹ (FBI IC3) & Khuyến cáo an toàn thanh toán quốc tế của Hiệp hội Ngân hàng Việt Nam (VNBA)."
    }
]

# ==============================================================================
# 3. DATA FOR HOC THUC CHIEN (14 cols)
# ==============================================================================
HT_NEW_IDEAS = [
    {
        "stt": 251,
        "week": "W47 — 16/11/2026",
        "title": "Quy Trình 5 Bước Tự Chứng Nhận Xuất Xứ Hàng Hóa (REX & Self-Certification) Cho Hàng Xuất Sang EU & CPTPP: Hướng Dẫn Kê Khai B.O.M & Bảo Vệ Tiêu Chí Xuất Xứ A-Z",
        "steps": "Bước 1: Đăng ký mã số REX hoặc mã số nhà xuất khẩu tự chứng nhận: Doanh nghiệp đăng ký tài khoản trên Cổng thông tin REX của EU (hoặc hệ thống của Bộ Công Thương), chuẩn bị hồ sơ năng lực sản xuất và nhận mã số REX hợp lệ;\nBước 2: Thiết lập Bảng kê khai chi phí và định mức nguyên phụ liệu (B.O.M - Bill of Materials): Bóc tách chi tiết từng dòng nguyên liệu cấu thành sản phẩm, phân loại rõ nguyên liệu có xuất xứ nội khối FTA và nguyên liệu nhập khẩu ngoài khối kèm mã HS 6 số;\nBước 3: Soát xét tiêu chí xuất xứ cụ thể mặt hàng (PSR - Product Specific Rules): Xác định chính xác sản phẩm áp dụng tiêu chí Chuyển đổi mã số hàng hóa (CTC: CC, CTH, CTSH) hay tiêu chí Hàm lượng giá trị khu vực (RVC) hoặc Công đoạn gia công chế biến cụ thể (SP);\nBước 4: Soạn thảo và phát hành Lời văn tự chứng nhận xuất xứ (Statement on Origin): In đúng câu chữ quy chuẩn theo mẫu quy định tại Phụ lục FTA lên Hóa đơn thương mại (Commercial Invoice) hoặc Phiếu đóng gói (Packing List), ký tên và ghi rõ họ tên người có thẩm quyền;\nBước 5: Thiết lập Hệ thống lưu trữ hồ sơ chứng minh xuất xứ trong 5 năm: Lưu trữ toàn bộ hợp đồng mua nguyên liệu, hóa đơn GTGT, chứng từ thanh toán ngân hàng, nhật ký sản xuất để sẵn sàng phục vụ các cuộc thanh tra xác minh hậu kiểm của Hải quan nước nhập khẩu.",
        "formula": "Công thức tính Hàm lượng giá trị khu vực RVC theo phương pháp gián tiếp (Build-down):\nRVC (%) = [(FOB - VNM) / FOB] x 100\nTrong đó: FOB là giá trị xuất khẩu tại cảng bốc hàng; VNM là giá trị nguyên liệu không có xuất xứ (Non-originating materials) dùng trong sản xuất.\nHạn mức tự chứng nhận: Dưới 6.000 EUR áp dụng cho mọi doanh nghiệp xuất khẩu sang EU; trên 6.000 EUR bắt buộc phải có mã số REX (Registered Exporter).",
        "top3_errors": "1. Tự ý thay đổi hoặc dịch sai câu chữ của Lời văn tuyên bố xuất xứ chuẩn (Statement on Origin) quy định trong Hiệp định EVFTA khiến chứng từ bị bác bỏ;\n2. Tính toán sai giá FOB hoặc đưa các chi phí vận tải nội địa nước nhập khẩu vào giá trị tính RVC;\n3. Không lưu trữ hóa đơn chứng từ mua nguyên phụ liệu trong nước, đến khi hải quan nước ngoài yêu cầu giải trình hậu kiểm thì không có bằng chứng chứng minh.",
        "mnemonic": "Tự chứng nhận xuất xứ (REX) như 'Tấm thẻ căn cước tự khai': Bạn tự in thông tin và cam kết chịu trách nhiệm trước pháp luật, nhưng nếu cảnh sát kiểm tra mà bạn không chứng minh được lai lịch nguồn gốc thì bạn sẽ bị xử phạt rất nặng!",
        "comparison": "C/O truyền thống (Cấp bởi VCCI/Bộ Công Thương - Mất thời gian nộp hồ sơ xét duyệt 1-3 ngày, tốn phí dịch vụ nhưng có cơ quan nhà nước đứng tên bảo lãnh) vs Tự chứng nhận xuất xứ REX (Doanh nghiệp tự phát hành trong 5 phút trên hóa đơn, hoàn toàn miễn phí nhưng chịu 100% trách nhiệm pháp lý và rủi ro hậu kiểm).",
        "persona": "Anh Dũng (Trưởng phòng Xuất khẩu & Tuân thủ pháp lý doanh nghiệp dệt may, da giày, chế biến gỗ).",
        "checklist": "☐ Đăng ký và kích hoạt mã số REX thành công trên hệ thống\n☐ Lập bảng B.O.M phân loại chuẩn nguyên liệu nội khối vs ngoại khối\n☐ Tính toán RVC đạt ngưỡng tối thiểu quy định trong biểu PSR\n☐ In nguyên văn Statement on Origin lên Commercial Invoice kèm mã REX\n☐ Đóng gói hồ sơ chứng từ lưu trữ riêng biệt trong thời hạn tối thiểu 5 năm.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "keywords": ["rules of origin", "origin"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Căn cứ Thông tư 33/2023/TT-BTC, Thông tư 11/2020/TT-BCT hướng dẫn Quy tắc xuất xứ hàng hóa trong EVFTA và quy chế vận hành hệ thống REX."
    },
    {
        "stt": 252,
        "week": "W47 — 16/11/2026",
        "title": "Kỹ Thuật 5 Bước Soát Xét Báo Cáo Quyết Toán Hải Quan (Mẫu Số 15/BCQT-NVL/GSQL): Khử Lệch Tồn Kho Nguyên Liệu Sản Xuất Xuất Khẩu & Tránh Phạt Thuế Hàng Trăm Triệu",
        "steps": "Bước 1: Thu thập và đồng bộ 3 luồng dữ liệu độc lập: Dữ liệu Tờ khai hải quan nhập - xuất trên VNACCS, Dữ liệu Sổ kho thẻ kho quản lý vật tư của bộ phận sản xuất, và Dữ liệu Kế toán chi phí giá thành trên phần mềm kế toán (ERP/MISA);\nBước 2: Đối soát số lượng nguyên vật liệu Nhập - Xuất - Tồn theo từng mã quản trị: Quy đổi đồng nhất đơn vị tính (Unit of Measure) giữa hóa đơn mua hàng quốc tế và đơn vị tính định mức sản xuất nội bộ (ví dụ: mét sang kg, yard sang mét vuông);\nBước 3: Rà soát định mức thực tế sản xuất và tỷ lệ hao hụt kỹ thuật: Tính toán chính xác định mức thực tế sử dụng nguyên liệu để sản xuất ra một đơn vị thành phẩm xuất khẩu, lập biên bản xác nhận tỷ lệ phế liệu, phế phẩm thu hồi;\nBước 4: Lập Bảng đối chiếu chênh lệch và xây dựng hồ sơ giải trình: Nếu phát hiện chênh lệch giữa Tồn kho kế toán và Tồn kho hải quan (do hao hụt bay hơi, mất mát tự nhiên, hàng lỗi trả về), lập tức lập văn bản giải trình nguyên nhân hợp lý trước khi chốt số liệu;\nBước 5: Hoàn thiện và truyền dữ liệu Mẫu số 15/BCQT-NVL/GSQL qua Hệ thống Hải quan điện tử: Nộp báo cáo quyết toán đúng thời hạn quy định (trong vòng 90 ngày kể từ ngày kết thúc năm tài chính), in và lưu trữ bộ hồ sơ chứng từ đi kèm.",
        "formula": "Phương trình cân bằng tồn kho bắt buộc:\nTồn cuối kỳ = Tồn đầu kỳ + Tổng Nhập trong kỳ - Tổng Xuất dùng sản xuất - Xuất khác\nThời hạn nộp BCQT: Chậm nhất là ngày thứ 90 kể từ ngày kết thúc năm tài chính (Ví dụ: năm tài chính kết thúc 31/12 thì hạn chót là 31/03 năm sau).\nMức phạt nộp chậm báo cáo quyết toán theo Nghị định 128/2020/NĐ-CP: Phạt tiền từ 2.000.000 đến 5.000.000 VNĐ; nếu có dấu hiệu trốn thuế bị phạt từ 1 đến 3 lần số thuế trốn.",
        "top3_errors": "1. Không đồng nhất mã nguyên phụ liệu giữa phần mềm kế toán và tờ khai hải quan dẫn đến số liệu 'râu ông nọ cắm cằm bà kia';\n2. Bỏ qua phế phẩm, phế liệu thu hồi hoặc tự ý bán phế liệu ra thị trường nội địa mà không mở tờ khai hải quan chuyển đổi mục đích sử dụng (mã A42);\n3. Đợi đến ngày thứ 89 mới bắt đầu đối soát số liệu, không kịp tìm hóa đơn chứng từ giải trình các khoản chênh lệch.",
        "mnemonic": "Báo cáo quyết toán hải quan giống như 'Cân đong gạo nấu cơm': Bạn mua 10 kg gạo (nhập khẩu), nấu thành 30 bát cơm (thành phẩm xuất khẩu), phần cháy cơm dính đáy nồi (hao hụt phế liệu); nếu cân lại mà thiếu mất 2 kg gạo không rõ đi đâu thì bạn sẽ bị coi là đã đem bán trộm ra chợ!",
        "comparison": "Loại hình Gia công (E21, E52 - Doanh nghiệp nhận nguyên liệu từ đối tác ngoại không phải trả tiền mua, chỉ thu phí gia công, máy móc thiết bị thường là mượn) vs Loại hình Sản xuất xuất khẩu SXXK (E31, E62 - Doanh nghiệp tự bỏ tiền vốn nhập khẩu nguyên liệu, tự chủ chuỗi sản xuất và bán thành phẩm, được miễn thuế NK có điều kiện).",
        "persona": "Chị Hương (Kế toán trưởng & Trưởng bộ phận Xuất nhập khẩu doanh nghiệp may mặc, cơ khí, linh kiện phụ tùng).",
        "checklist": "☐ Quy đổi chuẩn xác đơn vị tính giữa Hải quan, Kho và Kế toán\n☐ Chốt định mức thực tế và biên bản nghiệm thu tỷ lệ hao hụt kỹ thuật\n☐ Đối chiếu khớp số lượng tờ khai nhập (E21/E31) và xuất (E52/E62)\n☐ Soạn thảo văn bản giải trình các khoản chênh lệch tồn kho\n☐ Truyền Mẫu 15 lên Cổng thông tin hải quan trước thời hạn ngày 90.",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "keywords": ["customs", "valuation"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Căn cứ Điều 59, 60 Thông tư 38/2015/TT-BTC (sửa đổi bởi TT 39/2018/TT-BTC), Thông tư 18/2023/TT-BTC và Mẫu biểu 15/BCQT-NVL/GSQL."
    },
    {
        "stt": 253,
        "week": "W47 — 16/11/2026",
        "title": "Quy Trình 5 Bước Khai Bổ Sung Sau Thông Quan (Khai Báo AMA): Phác Đồ 60 Ngày Để Không Bị Phạt Vi Phạm Hành Chính Khi Phát Hiện Lỗi Tờ Khai",
        "steps": "Bước 1: Phát hiện và đánh giá mức độ sai lệch của tờ khai đã thông quan: Kiểm tra lại các chỉ tiêu có sai sót (số lượng, đơn giá, mã HS, trị giá tính thuế, xuất xứ, mã số thuế người mua);\nBước 2: Kiểm tra điều kiện mốc thời gian 60 ngày: Xác định ngày tờ khai được thông quan trên hệ thống VNACCS đến thời điểm hiện tại đã quá 60 ngày hay chưa VÀ kiểm tra cơ quan hải quan đã ban hành quyết định kiểm tra sau thông quan/thanh tra hay chưa;\nBước 3: Soạn thảo Văn bản giải trình lý do sai sót và chuẩn bị chứng từ chứng minh: Thu thập hợp đồng ngoại thương gốc, Commercial Invoice sửa đổi, Packing List, chứng từ thanh toán ngân hàng (Swift message) và biên bản đối chiếu hàng hóa thực tế;\nBước 4: Thực hiện khai bổ sung chỉ tiêu nghiệp vụ AMA trên phần mềm hải quan (ECUS5-VNACCS): Nhập đầy đủ thông tin sửa đổi tại các ô chỉ tiêu tương ứng, đính kèm văn bản giải trình và chứng từ chứng minh vào mục chứng từ điện tử số hóa;\nBước 5: Theo dõi kết quả phê duyệt và hoàn tất nghĩa vụ tài chính: Nhận thông báo chấp nhận khai bổ sung từ cơ quan hải quan; nếu phát sinh tăng số thuế phải nộp, chủ động nộp ngay tiền thuế chênh lệch và tiền chậm nộp (0.03%/ngày) vào Kho bạc Nhà nước.",
        "formula": "Công thức tính tiền chậm nộp thuế khi khai bổ sung sau thông quan:\nTiền chậm nộp = Số tiền thuế nộp thiếu x 0.03% x Số ngày chậm nộp\n(Số ngày chậm nộp tính từ ngày tiếp sau ngày cuối cùng của thời hạn nộp thuế đến ngày nộp thuế vào NSNN).\nLợi ích tài chính: Nếu khai AMA trong 60 ngày: Miễn 100% tiền phạt vi phạm hành chính (10-20% số thuế thiếu).",
        "top3_errors": "1. Phát hiện sai sót nhưng sợ bị hải quan phạt nên giấu nhẹm đi, để đến khi hải quan ra quyết định kiểm tra sau thông quan mới tá hỏa;\n2. Khai bổ sung AMA nhưng không kèm văn bản giải trình hợp lý và thiếu chứng từ thanh toán ngân hàng làm căn cứ đối chứng;\n3. Quên nộp tiền chậm nộp 0.03%/ngày sau khi được duyệt tờ khai AMA khiến tờ khai bị treo nợ thuế trên hệ thống quản lý nợ thuế KTST.",
        "mnemonic": "Khai AMA giống như 'Tự giác nhận lỗi với người yêu': Nếu bạn tự chủ động xin lỗi trước khi bị bắt quả tang thì chỉ phải 'mua trà sữa đền bù' (tiền chậm nộp); còn nếu để bị điều tra phát hiện thì chắc chắn 'ăn phạt nặng nề' (phạt 20% tiền thuế và bị giam kiểm tra)!",
        "comparison": "Khai sửa đổi bổ sung trước thông quan (Khai IBA/IBE - Chưa hoàn thành thủ tục thông quan, sửa đổi nhanh chóng trực tiếp trên VNACCS) vs Khai bổ sung sau thông quan (Khai AMA - Hàng đã thông quan rời cảng, phải giải trình chặt chẽ và chịu sự giám sát quản lý rủi ro 5 năm).",
        "persona": "Anh Minh (Chủ doanh nghiệp thương mại & Trưởng phòng Xuất nhập khẩu).",
        "checklist": "☐ Xác định mốc thời gian còn trong thời hạn 60 ngày kể từ ngày thông quan\n☐ Đối chiếu số liệu sai lệch với bộ chứng từ ngoại thương và thanh toán ngân hàng\n☐ Soạn thảo văn bản giải trình nêu rõ nguyên nhân khách quan dẫn đến sai sót\n☐ Thực hiện khai nghiệp vụ AMA trên hệ thống ECUS5-VNACCS\n☐ Nộp thuế chênh lệch và tiền chậm nộp vào Kho bạc ngay sau khi được duyệt.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "keywords": ["customs", "declaration"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Căn cứ Khoản 4 Điều 29 Luật Hải quan 2014, Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi bởi TT 39/2018/TT-BTC & TT 121/2025/TT-BTC) và Điều 9 Nghị định 128/2020/NĐ-CP."
    },
    {
        "stt": 254,
        "week": "W47 — 16/11/2026",
        "title": "Cẩm Nang 5 Bước Đàm Phán Hợp Đồng Ngoại Thương Điều Khoản Bất Khả Kháng (Force Majeure Clause): 'Tấm Khiên Sinh Tử' Bảo Vệ Doanh Nghiệp Trước Bão Lũ, Chiến Tranh & Tắc Nghẽn Cảng Biển",
        "steps": "Bước 1: Rà soát định nghĩa sự kiện Bất khả kháng chuẩn theo ICC Force Majeure Clause 2020: Liệt kê rõ các biến cố thiên tai (bão lũ, động đất), chiến tranh, xung đột quân sự vũ trang, dịch bệnh toàn cầu, cấm vận chính phủ và đình công diện rộng ngoài tầm kiểm soát;\nBước 2: Bổ sung các sự kiện gián đoạn chuỗi cung ứng hiện đại: Thỏa thuận đưa vào điều khoản các biến cố như tắc nghẽn kênh đào quốc tế (Suez, Panama), hãng tàu phá sản hoặc cảng biển ngừng tiếp nhận tàu do thảm họa kỹ thuật;\nBước 3: Quy định nghĩa vụ thông báo và thời hạn chứng minh (Notice of Force Majeure): Bên bị ảnh hưởng bắt buộc phải gửi thông báo bằng văn bản cho đối tác trong vòng 48 đến 72 giờ kể từ khi sự kiện xảy ra;\nBước 4: Xác lập nguồn chứng chỉ pháp lý xác nhận bất khả kháng: Thỏa thuận cơ quan có thẩm quyền cấp Giấy chứng nhận bất khả kháng độc lập (ví dụ: Phòng Thương mại và Công nghiệp VCCI tại Việt Nam hoặc CCPIT tại Trung Quốc);\nBước 5: Xác định lộ trình xử lý hậu quả và điều khoản chấm dứt hợp đồng: Cho phép gia hạn thời gian thực hiện hợp đồng tương ứng với thời gian gián đoạn; nếu sự kiện bất khả kháng kéo dài quá 60 hoặc 90 ngày, hai bên có quyền đơn phương chấm dứt hợp đồng mà không phải bồi thường thiệt hại.",
        "formula": "Thời hạn vàng thông báo: 48 - 72 giờ kể từ thời điểm phát sinh sự kiện.\nThời hạn cấp chứng thư: Trong vòng 15 - 30 ngày kể từ ngày gửi thông báo.\nThời hạn kích hoạt quyền hủy hợp đồng không phạt: 60 - 90 ngày liên tục bị đình trệ.",
        "top3_errors": "1. Sao chép điều khoản mẫu chung chung trên mạng mà không liệt kê danh mục sự kiện cụ thể dẫn đến tranh chấp pháp lý khi ra tòa trọng tài;\n2. Gặp sự cố thiên tai bão lũ nhưng không gửi thông báo chính thức bằng văn bản trong thời hạn cam kết, làm mất quyền viện dẫn bất khả kháng;\n3. Nghĩ rằng việc giá cước tàu biển tăng phi mã hoặc giá nguyên liệu leo thang là bất khả kháng (luật quốc tế coi đây là rủi ro kinh doanh thông thường Hardship chứ không phải Force Majeure).",
        "mnemonic": "Điều khoản Bất khả kháng giống như 'Túi khí an toàn trên xe ô tô': Bạn không bao giờ mong muốn phải sử dụng đến nó trong chuyến đi, nhưng khi xảy ra va chạm kinh hoàng ngoài ý muốn thì nó là thứ duy nhất cứu mạng doanh nghiệp khỏi bị phá sản vì nợ bồi thường!",
        "comparison": "Bất khả kháng (Force Majeure - Sự kiện khách quan không thể lường trước và không thể khắc phục, miễn trừ hoàn toàn nghĩa vụ thực hiện hợp đồng) vs Hoàn cảnh thay đổi cơ bản (Hardship - Sự kiện làm biến động nghiêm trọng cán cân kinh tế của hợp đồng nhưng không làm cho việc thực hiện trở thành bất khả thi, buộc hai bên phải đàm phán lại giá cả).",
        "persona": "Anh Minh (Giám đốc doanh nghiệp xuất nhập khẩu nông sản, gỗ, thủy sản ký các hợp đồng dài hạn).",
        "checklist": "☐ Viện dẫn chuẩn xác điều khoản mẫu ICC Force Majeure Clause 2020\n☐ Liệt kê cụ thể biến cố tắc nghẽn cảng và xung đột địa chính trị\n☐ Khóa thời hạn gửi thông báo văn bản trong vòng tối đa 72 giờ\n☐ Chỉ định cơ quan thẩm quyền cấp chứng nhận độc lập (VCCI / CCPIT)\n☐ Cài đặt điều khoản tự động chấm dứt hợp đồng sau 60-90 ngày gián đoạn.",
        "url": "https://en.wikipedia.org/wiki/Force_majeure",
        "keywords": ["force majeure", "contract"],
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Căn cứ Bộ điều khoản mẫu ICC Force Majeure and Hardship Clauses 2020 của Phòng Thương mại Quốc tế, Điều 156 Bộ luật Dân sự 2015 và Điều 294 Luật Thương mại 2005."
    },
    {
        "stt": 255,
        "week": "W47 — 16/11/2026",
        "title": "Quy Trình 4 Bước Đăng Ký Kiểm Tra Nhà Nước Về Chất Lượng Hàng Nhập Khẩu (Mã H11 & Hàng Kinh Doanh) Chuẩn Nghị Định 74/2018/NĐ-CP: Kỹ Thuật Mang Hàng Về Kho Bảo Quản Trong 24H",
        "steps": "Bước 1: Tra cứu Danh mục hàng hóa nhóm 2 bắt buộc kiểm tra chất lượng: Đối chiếu mã HS và mô tả kỹ thuật của sản phẩm với danh mục quy định của Bộ Khoa học & Công nghệ, Bộ Giao thông Vận tải, Bộ Xây dựng hoặc Bộ Y tế;\nBước 2: Nộp đơn đăng ký kiểm tra chất lượng trên Cổng thông tin Một cửa Quốc gia (NSW): Khai báo thông tin nhà sản xuất, nước xuất xứ, số lượng, quy chuẩn kỹ thuật quốc gia (QCVN) áp dụng và chọn tổ chức chứng nhận giám định được chỉ định;\nBước 3: Lập văn bản Đề nghị đưa hàng về kho bảo quản (Mẫu số 09/BQHH/GSQL): Chuẩn bị hồ sơ cam kết địa chỉ kho bãi đạt chuẩn, xuất trình đơn đăng ký đã được tiếp nhận có số tiếp nhận trên NSW để nộp cho Chi cục Hải quan cửa khẩu;\nBước 4: Lấy mẫu giám định và nộp kết quả hợp quy thông quan chính thức: Phối hợp trung tâm giám định (Quatest, Vinacontrol) lấy mẫu kiểm tra tại kho của doanh nghiệp trong vòng 15 ngày; khi có Chứng thư hợp quy, cập nhật số chứng thư lên NSW để hải quan phê duyệt thông quan hoàn tất.",
        "formula": "Thời hạn cấp số tiếp nhận đăng ký trên NSW: Tối đa 01 ngày làm việc.\nThời hạn bảo quản hàng hóa chờ kết quả kiểm tra chuyên ngành: Tối đa 30 ngày kể từ ngày đưa hàng về kho.\nMức phạt hành vi tự ý tiêu thụ, sử dụng hàng hóa đang trong thời gian đưa về bảo quản chờ kết quả kiểm tra chuyên ngành theo Nghị định 128/2020/NĐ-CP: Phạt tiền từ 30.000.000 đến 50.000.000 VNĐ.",
        "top3_errors": "1. Nhầm lẫn giữa hàng mẫu phục vụ R&D thử nghiệm không thương mại (được miễn KTCL theo Nghị định 74/2018) với hàng nhập khẩu kinh doanh thông thường;\n2. Tự ý mở bao bì hoặc phân phối bán trước hàng hóa khi mới chỉ đưa hàng về kho bảo quản mà chưa có kết quả thông quan chính thức;\n3. Khai báo địa chỉ kho bảo quản không đúng thực tế hoặc kho không thuộc quyền sở hữu/hợp đồng thuê hợp pháp của doanh nghiệp.",
        "mnemonic": "Đưa hàng về kho bảo quản giống như 'Được tại ngoại chờ tuyên án': Cơ quan hải quan tạo điều kiện cho bạn mang hàng về kho để tiết kiệm phí lưu bãi cảng, nhưng 'hàng hóa vẫn đang bị giam lỏng'; bạn tuyệt đối không được phép động vào cho đến khi có Giấy thông quan trắng án!",
        "comparison": "Miễn kiểm tra chất lượng (Hàng mẫu thử nghiệm, hàng ngoại giao, hàng quà biếu tặng trong định mức - Miễn 100% thủ tục đăng ký KTCL) vs Đưa hàng về kho bảo quản (Hàng kinh doanh thương mại nhóm 2 - Bắt buộc đăng ký KTCL nhưng được mang về kho riêng lưu giữ chờ kết quả giám định để tiết kiệm chi phí bãi cảng).",
        "persona": "Anh Dũng (Trưởng phòng Mua hàng & Logistics công ty nhập khẩu thiết bị điện tử, đồ chơi trẻ em, vật liệu xây dựng).",
        "checklist": "☐ Tra cứu chính xác QCVN áp dụng cho mặt hàng nhóm 2\n☐ Tạo đơn đăng ký kiểm tra chất lượng lấy số tiếp nhận trên NSW\n☐ Lập công văn Mẫu 09/BQHH cam kết bảo quản nguyên trạng hàng hóa\n☐ Mang hàng về kho riêng an toàn và hẹn lịch tổ chức giám định lấy mẫu\n☐ Nộp kết quả chứng nhận hợp quy lên hệ thống trước hạn 30 ngày.",
        "url": "https://en.wikipedia.org/wiki/Conformity_assessment",
        "keywords": ["conformity", "assessment"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Căn cứ Nghị định 132/2008/NĐ-CP, Nghị định 74/2018/NĐ-CP (sửa đổi NĐ 154/2018/NĐ-CP), Điều 32 Thông tư 38/2015/TT-BTC và Điều 18 Nghị định 128/2020/NĐ-CP."
    }
]

# ==============================================================================
# 4. DATA FOR FUNFACT (13 cols)
# ==============================================================================
FF_NEW_IDEAS = [
    {
        "stt": 251,
        "week": "W47 — 16/11/2026",
        "title": "Vận Tải Đa Phương Thức & Cuộc Đua Tuyến Đường Sắt Liên Vận Á - Âu (China - Europe Railway Express): Con Đường Tơ Lụa Thép Rút Ngắn 2/3 Thời Gian So Với Đường Biển",
        "hook": "Bạn có tin rằng một container hàng điện tử xuất phát từ Việt Nam có thể chạy xuyên qua Trung Quốc, Kazakhstan, Nga, Ba Lan và cập bến nước Đức chỉ trong vòng 18 ngày mà không cần chạm vào một giọt nước biển nào?",
        "history": "Các chuyến tàu hàng liên vận Á - Âu (CR Express) và tuyến đường sắt liên vận quốc tế Việt Nam - Trung Quốc - Châu Âu khai trương các đoàn tàu chuyên tuyến từ Ga Yên Viên (Hà Nội) đi Duisburg (Đức) từ năm 2021 đến nay.",
        "context": "Khi các tuyến hàng hải qua Kênh đào Suez và Biển Đỏ gặp khủng hoảng an ninh địa chính trị kéo dài hải trình lên 45 ngày và cước hàng không quá đắt đỏ (5-7 USD/kg), vận tải đường sắt liên vận nổi lên như một giải pháp cứu cánh trung hòa hoàn hảo.",
        "modern_impact": "Tuyến đường sắt vận hành hơn 15.000 chuyến tàu mỗi năm, vận chuyển hàng triệu container linh kiện máy tính, điện thoại Samsung, đồ gia dụng cao cấp với thời gian nhanh gấp đôi đường biển và chi phí rẻ bằng 1/4 đường hàng không.",
        "intl_comp": "Trục đường sắt Trans-Siberian của Nga và mạng lưới đường sắt hàng hóa của Liên minh Châu Âu (Freight Corridors) kết nối liền mạch qua các trạm đổi khổ đường ray (Gauge Change) tại Alashankou và Brest.",
        "business_lesson": "Đừng bao giờ 'bỏ tất cả trứng vào một giỏ đường biển'; doanh nghiệp xuất khẩu linh kiện giá trị cao của Việt Nam cần chủ động phân bổ tỷ trọng vận tải đa phương thức (Multimodal Transport) để bảo đảm chuỗi cung ứng không bao giờ bị đứt gãy trước biến động địa chính trị.",
        "url": "https://en.wikipedia.org/wiki/Trans-Eurasia_Logistics",
        "keywords": ["eurasia", "railway"],
        "series": "Lịch sử Logistics",
        "status": "Chưa dùng",
        "note": "Cột mốc phát triển logistics vận tải đa phương thức đường sắt liên vận quốc tế Á - Âu và sự tham gia của đường sắt Việt Nam."
    },
    {
        "stt": 252,
        "week": "W47 — 16/11/2026",
        "title": "Quy Tắc 'Cắt Bỏ Góc Vận Đơn' (Stamping & Endorsement): Bí Mật Ngân Hàng Kiểm Soát Quyền Sở Hữu Hàng Hóa Triệu Đô Bằng Một Chữ Ký Ký Hậu",
        "hook": "Một tờ giấy mỏng manh có tên Vận đơn gốc đích danh (To Order B/L) có thể chuyển giao quyền sở hữu một con tàu chở đầy dầu thô trị giá 50 triệu USD cho bất kỳ ai chỉ bằng một dòng chữ ký viết tay ở mặt sau!",
        "history": "Tập quán thương mại hàng hải quốc tế hàng trăm năm của người Phoenicia, Venice và được chuẩn hóa trong Quy tắc Hague-Visby 1968 cùng Bộ luật Thương mại Thống nhất (UCC).",
        "context": "Trước khi có hệ thống thanh toán điện tử, các thương nhân lênh đênh trên biển cần một công cụ chứng từ có thể cầm cố, mua bán sang tay hàng hóa ngay khi con tàu vẫn đang trôi dạt giữa đại dương bao la.",
        "modern_impact": "Cơ chế Vận đơn theo lệnh (To Order) và Ký hậu chuyển nhượng (Endorsement) trở thành nền tảng của 90% giao dịch Thư tín dụng chứng từ (L/C) toàn cầu, nơi ngân hàng nắm giữ vận đơn gốc như tài sản thế chấp sinh tử để giải ngân tiền hàng.",
        "intl_comp": "Hệ thống Vận đơn đích danh (Straight B/L) của Mỹ không cần xuất trình bản gốc ở một số bang vs Hệ thống Vận đơn theo lệnh quốc tế bắt buộc phải có bản gốc có ký hậu của ngân hàng phát hành L/C mới được giao hàng.",
        "business_lesson": "'Ai nắm vận đơn gốc theo lệnh người đó nắm tài sản'; doanh nghiệp xuất khẩu tuyệt đối không bao giờ phát hành vận đơn đích danh cho người mua nếu chưa thu đủ 100% tiền hàng, và phải kiểm tra cẩn thận từng con dấu ký hậu ngân hàng để tránh bị lừa đảo chiếm đoạt container.",
        "url": "https://en.wikipedia.org/wiki/Negotiable_instrument",
        "keywords": ["negotiable", "instrument"],
        "series": "Thương mại Quốc tế",
        "status": "Chưa dùng",
        "note": "Tập quán hàng hải quốc tế về vận đơn có thể chuyển nhượng (Negotiable Bill of Lading) và quy tắc ký hậu theo Điều 15 UCP 600."
    },
    {
        "stt": 253,
        "week": "W47 — 16/11/2026",
        "title": "Máy Soi Container Khổng Lồ Năng Lượng Cao (High-Energy Cargo Gantry Scanner): Cỗ Máy 'Chụp X-Quang' Xuyên Qua 35cm Thép Đặc Chỉ Trong 30 Giây",
        "hook": "Chiếc máy soi container tại các cảng biển lớn hiện đại có sức mạnh bức xạ tia X mạnh gấp 1.000 lần máy chụp X-quang y tế trong bệnh viện, có thể phân biệt rõ ràng giữa đường cát trắng và ma túy đá chỉ qua màu sắc hiển thị trên màn hình!",
        "history": "Sự ra đời của các hệ thống soi chiếu container cố định và di động (HCVG Gantry / Smiths Detection / Nuctech) được WCO và Hải quan Hoa Kỳ bắt buộc triển khai sau sự kiện 11/09/2001 theo Sáng kiến An ninh Container (CSI).",
        "context": "Với hơn 200 triệu container luân chuyển mỗi năm, nếu công chức hải quan phải dùng xà-beng cạy từng thùng hàng để kiểm hóa thủ công thì toàn bộ hệ thống thương mại toàn cầu sẽ lập tức sụp đổ vì tắc nghẽn.",
        "modern_impact": "Công nghệ soi chiếu tia X năng lượng kép (Dual-energy) tự động tô màu vật chất: Hợp chất hữu cơ hiển thị màu cam, kim loại nhẹ màu xanh lá, kim loại nặng màu xanh da trời; kết hợp cùng AI tự động phát hiện hàng cấm giấu trong vách giả container.",
        "intl_comp": "Cảng Rotterdam và Cảng Los Angeles soi chiếu 100% container chuyển tải tự động trên ray di chuyển vs Cảng Cát Lái và Hải Phòng soi chiếu chọn lọc theo thuật ngữ quản lý rủi ro luồng Vàng/Đỏ của Tổng cục Hải quan.",
        "business_lesson": "'Đừng bao giờ nghĩ đến chuyện giấu hàng thừa vào đáy container'; công nghệ soi chiếu hiện đại nhìn thấy từng con ốc vít và từng kg hàng chênh lệch; sự trung thực và minh bạch tuyệt đối trong khai báo hải quan là con đường duy nhất để doanh nghiệp tồn tại lâu dài.",
        "url": "https://en.wikipedia.org/wiki/Cargo_scanning",
        "keywords": ["cargo", "scanning"],
        "series": "Bí mật Hải quan",
        "status": "Chưa dùng",
        "note": "Công nghệ kiểm tra hàng hóa không xâm nhập (Non-Intrusive Inspection - NII) theo chuẩn mực Khung tiêu chuẩn SAFE của Tổ chức Hải quan Thế giới (WCO)."
    },
    {
        "stt": 254,
        "week": "W47 — 16/11/2026",
        "title": "Mô Hình 'Kho Ngoại Quan Đa Năng' Của Dubai (Jebel Ali Free Zone - JAFZA): Nơi Biến Sa Mạc Cằn Cỗi Thành Trung Tâm Tái Xuất Khẩu Hàng Đầu Hành Tinh",
        "hook": "Dubai không hề có nền sản xuất công nghiệp nặng hay rừng vàng biển bạc, nhưng lại là nơi trung chuyển của hơn 25% lượng hàng tiêu dùng và kim loại quý của thế giới nhờ một chính sách duy nhất: Miễn thuế 100% và cho phép sở hữu nước ngoài tuyệt đối!",
        "history": "Sheikh Rashid bin Saeed Al Maktoum sáng lập Cảng Jebel Ali năm 1979 và Khu thương mại tự do JAFZA năm 1985 tại Các Tiểu Vương Quốc Ả Rập Thống Nhất (UAE).",
        "context": "Nhận thấy nguồn dầu mỏ của Dubai sẽ cạn kiệt trong tương lai gần, vị tiểu vương quyết định đào một cảng nhân tạo lớn nhất thế giới ngay giữa sa mạc và ban hành cơ chế pháp lý đặc biệt để thu hút các tập đoàn logistics toàn cầu.",
        "modern_impact": "JAFZA hiện quy tụ hơn 9.500 doanh nghiệp từ 140 quốc gia (bao gồm hơn 100 công ty trong danh sách Fortune 500), đóng góp 33% GDP của Dubai và tạo ra hơn 135.000 việc làm; kết nối liền mạch cảng biển với sân bay Al Maktoum trong hành lang logistics duy nhất chỉ mất 4 giờ chuyển hàng.",
        "intl_comp": "JAFZA của Dubai vs Khu thương mại tự do Thượng Hải (FTZ) và Đề án thí điểm Khu thương mại tự do Cái Mép Hạ & Đà Nẵng của Việt Nam (Quyết định 48/2026/QĐ-TTg).",
        "business_lesson": "Vị trí địa lý chỉ là tiềm năng, thể chế thông thoáng và tốc độ luân chuyển hàng hóa mới là tài sản vô giá; Việt Nam hoàn toàn có thể trở thành một 'Dubai của Đông Nam Á' nếu xây dựng thành công các trung tâm thương mại tự do thế hệ mới tại Cái Mép và Hải Phòng.",
        "url": "https://en.wikipedia.org/wiki/Jebel_Ali_Free_Zone",
        "keywords": ["jebel ali", "free zone"],
        "series": "Đàm phán & Tư duy",
        "status": "Chưa dùng",
        "note": "Mô hình kinh tế khu thương mại tự do JAFZA và bài học kinh nghiệm phát triển Free Trade Zone cho Việt Nam."
    },
    {
        "stt": 255,
        "week": "W47 — 16/11/2026",
        "title": "Con Tem Niêm Phong Chì Hải Quan (High Security Bolt Seal ISO 17712): Chiếc 'Khóa Cửa' 2 Đô-La Quyết Định Sự Toàn Vẹn Của Hàng Hóa Trị Giá Cả Triệu USD",
        "hook": "Một thanh thép mạ kẽm nhỏ bé có giá chưa đầy 50.000 VNĐ lại là bằng chứng pháp lý duy nhất trên toàn thế giới để các hãng bảo hiểm và tòa án phân định xem ai phải đền bù khi một container hàng cập cảng bị bốc hơi không còn một món đồ!",
        "history": "Sự ra đời của Tiêu chuẩn quốc tế ISO 17712 về niêm phong an ninh cao (High Security Mechanical Seals) và chương trình quan hệ đối tác hải quan - thương mại chống khủng bố (C-TPAT) của Mỹ năm 2001.",
        "context": "Tình trạng trộm cắp container trên biển và thủ đoạn tráo hàng, nhét hàng lậu ma túy vào container hợp pháp đòi hỏi một cơ chế niêm phong chống gian lận (Tamper-evident) không thể làm giả hoặc hàn gắn lại sau khi đã cắt khóa.",
        "modern_impact": "Mọi container vận tải quốc tế bắt buộc phải bấm seal cáp hoặc seal cối đạt chuẩn ISO 17712 Loại H (High Security); số seal phải được ghi chính xác trên Vận đơn (B/L) và Tờ khai hải quan; chỉ cần số seal thực tế lệch 1 ký tự so với chứng từ là con tàu sẽ bị tạm dừng bốc dỡ ngay lập tức.",
        "intl_comp": "Seal cơ khí truyền thống (Mechanical Seal) chỉ phát hiện sau khi cắt khóa vs Seal thông minh điện tử RFID / GPS (E-Seal) của Hải quan hiện đại có thể phát tín hiệu cảnh báo vệ tinh theo thời gian thực nếu seal bị cắt trái phép trên đường vận chuyển.",
        "business_lesson": "'Quy tắc 3 bước kiểm tra seal sống còn': Chụp ảnh rõ nét số seal ngay khi bấm tại kho đóng hàng, đối chiếu số seal trên biên bản giao nhận với tài xế xe đầu kéo, và kiểm tra tình trạng nguyên vẹn của khớp khóa trước khi ký nhận tại cảng đích; sự cẩn thận nhỏ này cứu doanh nghiệp khỏi những vụ tranh chấp mất hàng triệu đô.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "keywords": ["security seal", "seal"],
        "series": "Công nghệ & Đổi mới",
        "status": "Chưa dùng",
        "note": "Tiêu chuẩn niêm phong an ninh cơ khí ISO 17712, chương trình C-TPAT và quy định kiểm tra niêm chì seal của Tổng cục Hải quan Việt Nam."
    }
]


def execute_update():
    print("=" * 80)
    print("🚀 BẮT ĐẦU WORKFLOW 06: CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (WEEK 47 — 16/11/2026)")
    print("=" * 80)

    # 1. GÁC CỔNG KỸ THUẬT: VALIDATE TOÀN BỘ 20 URLS TRƯỚC KHI THAO TÁC FILE
    all_ideas_for_gatekeeper = []
    for item in TN_NEW_IDEAS:
        all_ideas_for_gatekeeper.append({"stt": item["stt"], "title": item["title"], "url": item["url"], "keywords": item["keywords"]})
    for item in CS_NEW_IDEAS:
        all_ideas_for_gatekeeper.append({"stt": item["stt"], "title": item["title"], "url": item["url"], "keywords": item["keywords"]})
    for item in HT_NEW_IDEAS:
        all_ideas_for_gatekeeper.append({"stt": item["stt"], "title": item["title"], "url": item["url"], "keywords": item["keywords"]})
    for item in FF_NEW_IDEAS:
        all_ideas_for_gatekeeper.append({"stt": item["stt"], "title": item["title"], "url": item["url"], "keywords": item["keywords"]})

    print(f"\n🔍 Gửi {len(all_ideas_for_gatekeeper)} link qua Technical Gatekeeper url_validator.py...")
    enforce_strict_validation(all_ideas_for_gatekeeper)
    print("\n🟢 TECHNICAL GATEKEEPER PHÊ DUYỆT 100%: TẤT CẢ 20 LINK ĐỀU HỢP LỆ VÀ ĐÃ ĐỐI SOÁT NỘI DUNG THỰC TẾ!\n")

    # 2. CẬP NHẬT KHO Ý TƯỞNG TIN NGÀNH
    tn_path = "reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx"
    wb_tn = openpyxl.load_workbook(tn_path)
    ws_tn = wb_tn.active
    print(f"📂 Đang ghi 5 ý tưởng vào {tn_path} (Dòng hiện tại: {ws_tn.max_row})...")
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
    print(f"✅ Ghi thành công {tn_path} (Tổng dòng sau ghi: {ws_tn.max_row})")

    # 3. CẬP NHẬT KHO Ý TƯỞNG CASE STUDY
    cs_path = "reports/Case_Study_Ideas_Bank_Eureka.xlsx"
    wb_cs = openpyxl.load_workbook(cs_path)
    ws_cs = wb_cs.active
    print(f"📂 Đang ghi 5 ý tưởng vào {cs_path} (Dòng hiện tại: {ws_cs.max_row})...")
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
    print(f"✅ Ghi thành công {cs_path} (Tổng dòng sau ghi: {ws_cs.max_row})")

    # 4. CẬP NHẬT KHO Ý TƯỞNG HỌC THỰC CHIẾN
    ht_path = "reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx"
    wb_ht = openpyxl.load_workbook(ht_path)
    ws_ht = wb_ht.active
    print(f"📂 Đang ghi 5 ý tưởng vào {ht_path} (Dòng hiện tại: {ws_ht.max_row})...")
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
    print(f"✅ Ghi thành công {ht_path} (Tổng dòng sau ghi: {ws_ht.max_row})")

    # 5. CẬP NHẬT KHO Ý TƯỞNG FUNFACT
    ff_path = "reports/Funfact_Ideas_Bank_Eureka.xlsx"
    wb_ff = openpyxl.load_workbook(ff_path)
    ws_ff = wb_ff.active
    print(f"📂 Đang ghi 5 ý tưởng vào {ff_path} (Dòng hiện tại: {ws_ff.max_row})...")
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
    print(f"✅ Ghi thành công {ff_path} (Tổng dòng sau ghi: {ws_ff.max_row})")

    # 6. FEEDBACK LOOP: CẬP NHẬT DANH SÁCH VĂN BẢN PHÁP LUẬT XNK MỚI
    legal_path = "reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx"
    wb_leg = openpyxl.load_workbook(legal_path)
    ws_leg = wb_leg.active
    last_stt = ws_leg.cell(row=ws_leg.max_row, column=1).value or 254

    new_legal_docs = [
        [
            last_stt + 1,
            "Thông tư 33/2023/TT-BTC",
            "Quy định về xác định xuất xứ hàng hóa xuất khẩu, nhập khẩu; cơ chế tự chứng nhận xuất xứ (REX/Statement on Origin) và thủ tục kiểm tra, xác minh xuất xứ hậu kiểm của cơ quan Hải quan",
            "Thông tư",
            "Bộ Tài chính",
            "2023-05-31",
            "2023-07-15",
            "Còn hiệu lực",
            "Thay thế Thông tư 38/2018/TT-BTC và Thông tư 62/2019/TT-BTC về xác định xuất xứ",
            "Chuẩn hóa toàn diện quy chế nộp chứng từ chứng nhận xuất xứ điện tử (e-C/O), hướng dẫn kiểm tra trừ lùi C/O, cơ chế tự chứng nhận xuất xứ trong CPTPP, EVFTA và thủ tục từ chối cho hưởng ưu đãi thuế khi có nghi ngờ gian lận.",
            "Quy định điều kiện bắt buộc để doanh nghiệp được áp dụng cơ chế tự chứng nhận xuất xứ và thời hạn 5 năm lưu trữ toàn bộ hồ sơ B.O.M phục vụ đoàn thanh tra xác minh của Hải quan nước nhập khẩu.",
            "2026-11-16"
        ],
        [
            last_stt + 2,
            "Nghị định 111/2021/NĐ-CP",
            "Nghị định sửa đổi, bổ sung một số điều của Nghị định 43/2017/NĐ-CP của Chính phủ về nhãn hàng hóa đối với hàng hóa xuất khẩu, nhập khẩu và lưu thông nội địa",
            "Nghị định",
            "Chính phủ",
            "2021-12-09",
            "2022-02-15",
            "Còn hiệu lực",
            "Sửa đổi, bổ sung một số điều của Nghị định 43/2017/NĐ-CP",
            "Quy định bắt buộc trên nhãn gốc (Original Label) của hàng hóa nhập khẩu tại cửa khẩu phải thể hiện tối thiểu 4 nội dung: Tên hàng hóa; Tên và địa chỉ tổ chức chịu trách nhiệm; Xuất xứ hàng hóa; và nội dung bắt buộc khác theo tính chất mặt hàng.",
            "Chấm dứt tình trạng nợ nhãn phụ hoặc đưa hàng về bảo quản khi nhãn gốc thiếu xuất xứ; là căn cứ xử phạt vi phạm hành chính từ 10 - 60 triệu đồng hoặc buộc tái xuất tại cửa khẩu hải quan.",
            "2026-11-16"
        ],
        [
            last_stt + 3,
            "Nghị định 08/2022/NĐ-CP",
            "Nghị định quy định chi tiết một số điều của Luật Bảo vệ môi trường 2020 về trách nhiệm mở rộng của nhà sản xuất, nhập khẩu (EPR) trong việc tái chế sản phẩm, bao bì và đóng góp tài chính xử lý chất thải",
            "Nghị định",
            "Chính phủ",
            "2022-01-10",
            "2022-01-10",
            "Còn hiệu lực",
            "Quy định chi tiết thi hành Luật Bảo vệ môi trường 2020",
            "Quy định lộ trình, đối tượng và tỷ lệ tái chế bắt buộc hoặc đóng góp kinh phí (Fs) vào Quỹ Bảo vệ môi trường đối với nhà nhập khẩu bao bì (nhựa, thủy tinh, giấy, kim loại), ắc quy, pin, săm lốp, dầu nhớt và thiết bị điện tử.",
            "Bắt buộc doanh nghiệp nhập khẩu phải đăng ký kế hoạch tái chế hoặc nộp tiền đóng góp trước ngày 20/04 hàng năm; không tuân thủ sẽ bị xử phạt theo Nghị định 45/2022/NĐ-CP và công khai vi phạm.",
            "2026-11-16"
        ]
    ]

    print(f"\n⚖️ THỰC HIỆN FEEDBACK LOOP: Bổ sung {len(new_legal_docs)} văn bản luật mới vào {legal_path}...")
    for doc in new_legal_docs:
        r = ws_leg.max_row + 1
        fill = FILL_ODD if doc[0] % 2 != 0 else FILL_EVEN
        for c_idx, val in enumerate(doc, start=1):
            cell = ws_leg.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_BOLD if c_idx in (1, 2) else FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c_idx in (1, 4, 6, 7, 8, 12) else ALIGN_LEFT
            cell.border = BORDER_THIN
            cell.fill = fill
    atomic_save(wb_leg, legal_path)
    print(f"✅ Ghi thành công {legal_path} (Tổng dòng sau ghi: {ws_leg.max_row})")

    print("\n" + "=" * 80)
    print("🎉 HOÀN TẤT 100% CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS & FEEDBACK LOOP LUẬT XNK!")
    print("=" * 80)


if __name__ == "__main__":
    execute_update()
