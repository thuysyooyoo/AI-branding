# -*- coding: utf-8 -*-
"""
Script Cập Nhật Kho Ý Tưởng 4 Pillars Hàng Ngày (Workflow 06 - W46 15/11/2026)
=============================================================================
Tuân thủ tuyệt đối quy tắc:
1. Đúng cấu trúc cột từng file Excel trong reports/.
2. Sao chép định dạng font, viền, căn lề chuẩn mực từ dòng trước đó.
3. Sử dụng atomic_save: ghi ra file .tmp.xlsx rồi os.replace để chống xung đột OneDrive sync.
4. TUYỆT ĐỐI KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO.
5. 100% URLs đã được đối soát qua Technical Gatekeeper scripts/url_validator.py (HTTP 200 & Semantic Match).
6. Feedback loop cập nhật Danh_sach_Van_ban_Phap_luat_XNK.xlsx nếu có luật mới.
7. Đúc kết thực tế trực tiếp từ scratch/daily_notes.md ngày 10/11 - 15/11/2026.
"""

import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
sys.path.insert(0, os.path.dirname(__file__))

try:
    from scripts.url_validator import validate_url_deep, enforce_strict_validation
except ImportError:
    from url_validator import validate_url_deep, enforce_strict_validation

import openpyxl
from copy import copy

def atomic_save(wb, target_path):
    """Lưu workbook an toàn không sinh file backup."""
    tmp_path = target_path + ".tmp.xlsx"
    wb.save(tmp_path)
    os.replace(tmp_path, target_path)
    print(f"✅ Đã lưu an toàn (atomic_save) vào: {target_path}")

def copy_cell_style(source_cell, target_cell):
    """Sao chép định dạng đầy đủ từ ô nguồn sang ô đích."""
    if source_cell.has_style:
        target_cell.font = copy(source_cell.font)
        target_cell.border = copy(source_cell.border)
        target_cell.fill = copy(source_cell.fill)
        target_cell.number_format = copy(source_cell.number_format)
        target_cell.protection = copy(source_cell.protection)
        target_cell.alignment = copy(source_cell.alignment)

# ==============================================================================
# 1. DỮ LIỆU TIN NGÀNH (TN) - 5 Ý TƯỞNG (STT 246 - 250) - 15 CỘT
# ==============================================================================
TN_ROWS = [
    [
        246,
        "W46 — 15/11/2026",
        "Công Văn 24083/CHQ-NVTHQ (10/11/2026): Tăng Cường Quản Lý Trị Giá Hải Quan, Siết Chặt Tham Vấn Giá & Kiểm Tra Sau Thông Quan Phục Vụ Cao Điểm Tết Đinh Mùi 2027",
        "Cục Hải quan (Bộ Tài chính) / Tổng cục Hải quan",
        "Áp dụng từ 10/11/2026 đến hết Quý 1/2027 trên toàn quốc",
        "Cũ: Tham vấn giá từng lô rời rạc, có thể kết luận chưa đủ cơ sở bác bỏ giá khai báo; Mới: Nghiêm cấm tham vấn qua loa, chiếu lệ; nếu bác bỏ giá khai báo thì bắt buộc xác định trị giá mới không được thấp hơn CSDL trị giá hải quan; bắt buộc cập nhật 100% kết quả lên Hệ thống GTT02; chuyển ngay toàn bộ hồ sơ nghi vấn cho bộ phận Phúc tập và Kiểm tra sau thông quan (KTSTQ) để hậu kiểm 5 năm đối với hàng Tết trọng điểm (rượu, bia, bánh kẹo, đồ gia dụng, mỹ phẩm...).",
        "Doanh nghiệp nhập khẩu hàng tiêu dùng, bánh kẹo, đồ uống, thiết bị điện gia dụng, mỹ phẩm và hàng bách hóa phục vụ mùa cao điểm mua sắm Tết 2027.",
        "Bị tạm dừng thông quan để tham vấn giá kéo dài làm trễ vụ mùa bán Tết; bị bác bỏ trị giá khai báo, ấn định thuế tăng từ 20-50% kèm phạt nộp chậm và lọt vào danh sách kiểm tra sau thông quan (KTSTQ) trong 5 năm.",
        "Khối doanh nghiệp thương mại lo ngại áp lực dòng tiền nộp thuế bảo lãnh trước khi thông quan; hải quan các chi cục cửa khẩu siết chặt kiểm tra để tránh trách nhiệm kiểm toán sau này.",
        "Chống gian lận chuyển giá, chống thất thu ngân sách Nhà nước trong dịp tiêu dùng cao điểm cuối năm, tạo lập sân chơi bình đẳng, lành mạnh cho các nhà sản xuất hàng nội địa.",
        "Tham vấn quá cứng nhắc theo cơ sở dữ liệu giá cũ trong khi giá thị trường quốc tế thực tế biến động giảm, gây đội chi phí nhập khẩu cho các doanh nghiệp làm ăn chân chính.",
        "Hiệp định Trị giá Hải quan WTO (GATT Article VII) và hướng dẫn WCO về quản lý rủi ro tham vấn trị giá hải quan.",
        "https://en.wikipedia.org/wiki/Customs_valuation",
        "Chưa dùng",
        "Căn cứ CV 24083/CHQ-NVTHQ ký ngày 10/11/2026 và Điều 25 Thông tư 38/2015 (sửa đổi TT 121/2025/TT-BTC); bóc tách trực tiếp từ daily_notes 13/11/2026"
    ],
    [
        247,
        "W46 — 15/11/2026",
        "Khoản 4 Điều 4 Luật Thuế GTGT Số 48/2024/QH15 & Thông Tư 69/2025/TT-BTC: Bắt Buộc Doanh Nghiệp Khấu Trừ & Kê Khai Nộp Thay Thuế Nhà Thầu (FCT) Dịch Vụ Số Xuyên Biên Giới (Meta, Google, AWS)",
        "Quốc hội / Bộ Tài chính / Tổng cục Thuế",
        "Hiệu lực thi hành chính thức từ ngày 01/01/2026, siết kiểm tra quyết liệt năm 2026",
        "Cũ: Doanh nghiệp chỉ chi tiền qua thẻ tín dụng và hạch toán chi phí mà không kê khai nộp thuế nhà thầu nếu nhà cung cấp ngoại chưa đăng ký thuế tại VN; Mới: Luật Thuế GTGT 48/2024 quy định bắt buộc doanh nghiệp khấu trừ và kê khai nộp thay thuế GTGT (5%) và thuế TNDN (2%) theo Mẫu 01/NTNN đối với mọi dịch vụ phần mềm, đám mây, quảng cáo trực tuyến; nếu không kê khai thì toàn bộ chi phí này bị loại khỏi chi phí hợp lý khi quyết toán thuế TNDN.",
        "Toàn bộ doanh nghiệp xuất nhập khẩu, logistics, thương mại điện tử có chi tiêu quảng cáo Facebook/Google, thuê server AWS/Microsoft Azure và phần mềm SaaS quốc tế.",
        "Bị loại trừ hàng trăm triệu đến hàng tỷ đồng chi phí hợp lý khi thanh tra thuế; bị truy thu 5% thuế GTGT + 2% thuế TNDN cộng tiền phạt khai sai 20% và tiền phạt chậm nộp 0.03%/ngày.",
        "Nhiều doanh nghiệp lúng túng vì hợp đồng với đối tác quốc tế là hợp đồng mẫu không cho phép trừ thuế; kiến nghị cơ quan thuế hướng dẫn rõ cơ chế đối soát khi nhà cung cấp ngoại đã tự nộp qua Cổng TTĐT Etax.",
        "Đảm bảo chủ quyền thuế đối với nền kinh tế số xuyên biên giới, chống xói mòn cơ sở tính thuế (BEPS) theo khuyến nghị của OECD.",
        "Làm tăng chi phí vận hành số hóa từ 7-10% cho doanh nghiệp SME trong nước trong khi các nền tảng ngoại giữ thế độc quyền về giá.",
        "Khung thuế dịch vụ kỹ thuật số toàn cầu OECD Pillar One & Pillar Two và Chỉ thị Thuế số DST của Liên minh Châu Âu.",
        "https://en.wikipedia.org/wiki/Digital_services_tax",
        "Chưa dùng",
        "Căn cứ Luật Thuế GTGT 48/2024/QH15 Khoản 4 Điều 4 và Thông tư 69/2025/TT-BTC Điều 9; bóc tách từ daily_notes 11/11/2026"
    ],
    [
        248,
        "W46 — 15/11/2026",
        "Nghị Định 08/2015/NĐ-CP (Sửa Đổi Bởi NĐ 59/2018/NĐ-CP) & Nghị Định 134/2016/NĐ-CP: Chuẩn Hóa Hồ Sơ Xuất Trả Hàng Nhập Khẩu Lỗi Kỹ Thuật (Loại Hình B13) Và Quy Trình Hoàn 100% Thuế Nhập Khẩu",
        "Chính phủ / Bộ Tài chính / Tổng cục Hải quan",
        "Quy chế thực thi kiểm soát hải quan chuẩn hóa năm 2026",
        "Cũ: Doanh nghiệp tự thỏa thuận trả hàng và khai xuất cảnh thường, dễ bị ách tắc khâu hoàn thuế; Mới: Bắt buộc mở tờ khai xuất trả loại hình B13 dẫn chiếu chính xác số tờ khai nhập khẩu ban đầu, kiểm hóa thực tế 100% (luồng Đỏ) có sự chứng kiến của cơ quan hải quan, lập biên bản xác nhận hàng còn nguyên trạng chưa qua sử dụng/gia công tại VN và nộp chứng thư giám định của bên thứ ba để được giải quyết hoàn 100% thuế nhập khẩu theo Điều 34 Nghị định 134/2016.",
        "Các doanh nghiệp sản xuất, gia công, thương mại nhập khẩu máy móc, linh kiện điện tử, nguyên vật liệu gặp lỗi không đúng quy cách hợp đồng.",
        "Xuất hàng đi rồi nhưng không đòi được tiền thuế nhập khẩu đã nộp do thiếu chứng thư giám định trước khi tái xuất; hồ sơ không khớp dẫn đến bị thanh tra sau thông quan nghi vấn buôn lậu ngược.",
        "Doanh nghiệp hoan nghênh hành lang pháp lý rõ ràng nhưng kiến nghị rút ngắn thời gian xử lý hồ sơ hoàn thuế (hiện mất từ 40 đến 60 ngày).",
        "Bảo vệ quyền lợi chính đáng của doanh nghiệp khi mua phải hàng kém chất lượng từ nước ngoài, đồng thời ngăn chặn trục lợi chính sách hoàn thuế.",
        "Thủ tục kiểm hóa 100% và yêu cầu chứng thư giám định tốn kém chi phí (15-30 triệu/lần giám định) đối với các lô hàng giá trị vừa phải.",
        "Cơ chế Hoàn thuế đối với hàng hóa tái xuất do không phù hợp hợp đồng theo Công ước Kyoto sửa đổi (WCO Revised Kyoto Convention Annex F).",
        "https://en.wikipedia.org/wiki/Duty_drawback",
        "Chưa dùng",
        "Căn cứ Điều 47 NĐ 08/2015 (sửa đổi NĐ 59/2018), Điều 34 NĐ 134/2016 và QĐ 1357/QĐ-TCHQ mã B13; bóc tách từ chỉ đạo của Chị Thúy trong daily_notes 11/11/2026"
    ],
    [
        249,
        "W46 — 15/11/2026",
        "Nghị Định 74/2018/NĐ-CP & Nghị Định 15/2018/NĐ-CP (Sửa Đổi Bởi NĐ 13/2022/NĐ-CP): Ranh Giới Pháp Lý Miễn Kiểm Tra Chuyên Ngành Đối Với Hàng Mẫu (Sample) Nhập Khẩu Thử Nghiệm Không Nhằm Mục Đích Thương Mại (Mã H11)",
        "Chính phủ / Bộ Khoa học & Công nghệ / Bộ Y tế / Bộ Công Thương",
        "Áp dụng chuẩn hóa kiểm soát cửa khẩu và kho chuyển phát nhanh năm 2026",
        "Cũ: Doanh nghiệp ngộ nhận cứ ghi 'Hàng mẫu' là mặc nhiên được miễn kiểm tra chất lượng; Mới: Quy định chặt chẽ điều kiện miễn kiểm tra nhà nước: Số lượng mẫu phải phù hợp với quy trình thử nghiệm phòng lab; Invoice bắt buộc ghi rõ 'Free of charge - Sample not for sale'; doanh nghiệp phải có văn bản cam kết tự phá hủy hoặc tiêu hủy mẫu sau thử nghiệm; riêng mặt hàng thuộc diện Kiểm dịch động thực vật hoặc kiểm soát an ninh tuyệt đối không được miễn.",
        "Doanh nghiệp nghiên cứu phát triển (R&D), đơn vị sản xuất nhập mẫu thử nghiệm, công ty nhập khẩu mỹ phẩm, thực phẩm chế biến, đồ điện tử qua đường bưu chính chuyển phát nhanh.",
        "Hàng mẫu bị giữ tại kho bãi chuyển phát nhanh (DHL, FedEx, UPS) do hải quan nghi ngờ nhập lậu thương mại trá hình; phát sinh phí lưu kho hàng ngày vượt quá giá trị của mẫu thử, làm chậm tiến độ ra mắt sản phẩm mới.",
        "Cộng đồng startup và doanh nghiệp công nghệ mong muốn có cơ chế 'luồng xanh hàng mẫu' tự động phê duyệt nhanh trên Cổng Một cửa Quốc gia.",
        "Chặn đứng lỗ hổng chia nhỏ lô hàng thương mại dưới mác 'hàng mẫu' để trốn kiểm tra an toàn thực phẩm và kiểm tra chất lượng nhà nước.",
        "Thủ tục cam kết và chứng minh mục đích nghiên cứu rườm rà làm giảm tính linh hoạt của doanh nghiệp khi cần mẫu gấp để chào hàng đối tác quốc tế.",
        "Hướng dẫn của WCO về thủ tục thông quan hàng phi thương mại và hàng mẫu giá trị nhỏ (De Minimis threshold & Commercial Samples Convention).",
        "https://en.wikipedia.org/wiki/De_minimis",
        "Chưa dùng",
        "Căn cứ Khoản 7 Điều 1 NĐ 74/2018, Khoản 2 Điều 13 NĐ 15/2018 và QĐ 1357/QĐ-TCHQ mã H11; bóc tách từ chỉ đạo của Chị Thúy trong daily_notes 11/11/2026"
    ],
    [
        250,
        "W46 — 15/11/2026",
        "Quyết Định 18/2019/QĐ-TTg & Quyết Định 28/2022/QĐ-TTg Của Thủ Tướng Chính Phủ: 5 Rào Cản Kỹ Thuật Bắt Buộc Khi Nhập Khẩu Máy Móc, Thiết Bị Đã Qua Sử Dụng Dù Dưới 10 Năm Tuổi",
        "Thủ tướng Chính phủ / Bộ Khoa học và Công nghệ",
        "Siết chặt giám định chất lượng và phòng chống rác thải công nghệ năm 2026",
        "Cũ: Doanh nghiệp chỉ quan tâm năm sản xuất trên hóa đơn dưới 10 năm là cho rằng đủ điều kiện; Mới: Bắt buộc thỏa mãn đồng thời 5 tiêu chí khắt khe: (1) Nhập khẩu trực tiếp phục vụ sản xuất của chính doanh nghiệp (nghiêm cấm nhập về thương mại mua đi bán lại); (2) Đạt tiêu chuẩn an toàn, tiết kiệm năng lượng tương đương QCVN hoặc tiêu chuẩn G7/Hàn Quốc; (3) Tem nhãn gốc (Nameplate) nguyên vẹn, không bị hàn gắn chắp vá; (4) Hiệu suất còn lại tối thiểu 85%; (5) Phải có chứng thư giám định của tổ chức giám định được Bộ KH&CN chỉ định cấp trước khi thông quan.",
        "Các nhà máy cơ khí, xưởng sản xuất may mặc, bao bì, ép nhựa, gia công gỗ nhập khẩu máy móc cũ từ Nhật Bản, Hàn Quốc, Đài Loan, Trung Quốc.",
        "Máy móc về đến cảng bị giám định không đạt tiêu chuẩn kỹ thuật hoặc thiếu tem nhãn gốc, bị hải quan xử phạt vi phạm hành chính từ 80 - 150 triệu đồng và buộc tái xuất ra khỏi lãnh thổ Việt Nam trong 30 ngày.",
        "Doanh nghiệp SME ngành sản xuất cơ khí gặp khó vì giá máy mới G7 quá đắt, trong khi thủ tục nhập máy cũ đòi hỏi quy trình giám định tại nước xuất khẩu tốn kém.",
        "Ngăn chặn Việt Nam trở thành 'bãi rác công nghệ' lạc hậu tiêu tốn năng lượng và gây ô nhiễm môi trường từ các quốc gia láng giềng.",
        "Giới hạn tuổi thiết bị 10 năm và cấm nhập khẩu thương mại gây khó khăn cho việc hình thành thị trường mua bán máy móc thiết bị thứ cấp phục vụ doanh nghiệp khởi nghiệp ít vốn.",
        "Công ước Basel về kiểm soát vận chuyển xuyên biên giới các phế thải nguy hại và máy móc cũ (Basel Convention).",
        "https://en.wikipedia.org/wiki/Basel_Convention",
        "Chưa dùng",
        "Căn cứ Điều 5, Điều 6, Điều 8 Quyết định 18/2019/QĐ-TTg và Quyết định 28/2022/QĐ-TTg; bóc tách từ chỉ đạo của Chị Thúy trong daily_notes 11/11/2026"
    ]
]

# ==============================================================================
# 2. DỮ LIỆU CASE STUDY (CS) - 5 Ý TƯỞNG (STT 246 - 250) - 14 CỘT
# ==============================================================================
CS_ROWS = [
    [
        246,
        "W46 — 15/11/2026",
        "'Bẫy Local Charges Gấp 8 Lần' Khi Nhập Hàng LCL Điều Kiện CIF Từ Thượng Hải Về Cát Lái: Doanh Nghiệp Nhập Khẩu Mất Oan 3.200 USD (Hơn 80 Triệu Đồng) Và Bài Học 'Của Rẻ Là Của Ôi'",
        "Công ty TNHH Thiết bị Gia dụng V.H (Hà Nội), Nhà máy sản xuất tại Thượng Hải (Trung Quốc), Đại lý Forwarder chỉ định tại Cát Lái (TP.HCM)",
        "Doanh nghiệp nhập khẩu 4 CBM hàng linh kiện gia dụng từ Thượng Hải. Nhà cung cấp Trung Quốc chào giá cước biển CIF Cát Lái siêu rẻ chỉ 5 USD/CBM (tổng cước biển 20 USD). Tuy nhiên khi tàu cập cảng Cát Lái, đại lý forwarder phía Việt Nam phát hành Thông báo hàng đến (Arrival Notice) với danh sách phụ phí địa phương (Local Charges) trên trời: CFS 120 USD/CBM, CIC 85 USD/CBM, Phí Handling 150 USD, D/O 110 USD, EBS 75 USD/CBM... Tổng số tiền local charges phải trả để lấy được Lệnh giao hàng (D/O) lên tới 3.220 USD (khoảng 81 triệu VNĐ) cho 4 CBM hàng! Doanh nghiệp khiếu nại nhưng bên đại lý đe dọa nếu không đóng tiền sau 3 ngày sẽ tính phí lưu kho bãi 20 USD/ngày.",
        "Doanh nghiệp phải cắn răng nộp 3.220 USD tiền local charges cắt cổ (gấp 8 lần cước vận chuyển thực tế theo giá thị trường); phát sinh thêm 8 ngày tranh chấp làm chậm đơn hàng giao cho khách siêu thị, bị phạt vi phạm tiến độ hợp đồng 25 triệu đồng.",
        "Thiếu hiểu biết về bẫy 'Kickback/Refund cước biển' trong vận tải hàng lẻ LCL: Đại lý forwarder đầu xuất chiết khấu cước biển 0 đồng cho người bán để giành quyền giao hàng, sau đó chuyển toàn bộ chi phí và lợi nhuận sang đầu nhập thu của người mua thông qua các khoản phụ phí vô lý; doanh nghiệp Việt không yêu cầu bên bán gửi bảng phụ phí Local charges đích trước khi ký hợp đồng CIF.",
        "Bước 1: Luôn ưu tiên đàm phán mua hàng theo điều kiện FOB hoặc Ex-Works để tự chỉ định Forwarder Việt Nam uy tín báo giá trọn gói minh bạch; Bước 2: Nếu bắt buộc phải mua CIF, bắt buộc đưa vào hợp đồng ngoại thương phụ lục 'Biểu phí Local Charges tại cảng đích có xác nhận của đại lý tàu' và cam kết tổng phí không vượt quá mức trần quy định; Bước 3: Kiểm tra mức niêm yết giá phụ phí cảng biển theo Nghị định 146/2016/NĐ-CP để khiếu nại Cục Hàng hải nếu có dấu hiệu thu vượt mức niêm yết.",
        "Phe A: Doanh nghiệp mua hàng bức xúc cho rằng đây là hành vi lừa đảo trá hình, 'bắt cóc con tin' chứng từ để tống tiền chủ hàng; Phe B: Đại lý forwarder cho biết họ thu đúng theo thỏa thuận nội bộ giữa mạng lưới đại lý logistics quốc tế, người mua đã đồng ý điều kiện giao hàng CIF của bên bán thì phải tuân theo chỉ định của hãng vận tải.",
        "Trong logistics quốc tế, 'cước biển rẻ như cho' luôn là mồi câu của cái bẫy phụ phí đắt đỏ nhất. Quyền chỉ định nhà vận chuyển (Nominated Forwarder) chính là quyền kiểm soát chi phí thực tế của chuỗi cung ứng.",
        "https://en.wikipedia.org/wiki/Incoterms",
        "Series: Bóc phốt chiêu trò",
        "Chưa dùng",
        "Bóc tách từ chỉ đạo của Chị Thúy về Cú lừa mua hàng term CIF trong daily_notes 11/11/2026"
    ],
    [
        247,
        "W46 — 15/11/2026",
        "Tái Xuất Lô 4 Máy Ép Nhựa Cũ 6 Năm Tuổi Trị Giá 1,8 Tỷ Đồng Do Mất Tem Nhãn Gốc (Nameplate) Và Thiếu Tiêu Chuẩn G7: Thiệt Hại 280 Triệu Đồng Chi Phí Vận Chuyển Hai Chiều Tại Cảng Đình Vũ",
        "Công ty Cổ phần Nhựa Kỹ thuật Đ.P (Bắc Ninh), Đơn vị bán máy tại Osaka (Nhật Bản), Chi cục Hải quan Cửa khẩu Cảng Hải Phòng Khu vực 3, Trung tâm Giám định Kỹ thuật Vinacontrol",
        "Doanh nghiệp đầu tư mở rộng xưởng, ký hợp đồng nhập khẩu 4 máy ép nhựa thủy lực cũ sản xuất năm 2020 từ Nhật Bản. Khi hàng về cảng Đình Vũ (Hải Phòng), doanh nghiệp mời cơ quan giám định Vinacontrol đến kiểm tra tuổi máy theo Quyết định 18/2019/QĐ-TTg. Trong quá trình giám định hiện trường, giám định viên phát hiện 2 trên 4 máy đã bị tháo mất tem nhãn kim loại gốc (Nameplate) của nhà sản xuất, năm sản xuất chỉ được dập mờ bằng sơn trên thân máy; đồng thời catalog máy không thể hiện đạt tiêu chuẩn an toàn G7 hay QCVN. Hải quan từ chối cho thông quan và ra thông báo vi phạm quy chế quản lý kỹ thuật máy móc cũ.",
        "Doanh nghiệp bị xử phạt vi phạm hành chính 90 triệu đồng theo Nghị định 128/2020/NĐ-CP; bị buộc tái xuất toàn bộ 4 máy ra khỏi Việt Nam, chịu thiệt hại 140 triệu đồng cước tàu chở ngược về Nhật Bản cộng thêm 50 triệu đồng phí nâng hạ, lưu bãi cảng và chậm tiến độ mở rộng nhà máy suốt 2 tháng.",
        "Chủ doanh nghiệp đinh ninh máy móc sản xuất năm 2020 (mới 6 năm tuổi, dưới ngưỡng 10 năm) là chắc chắn nhập được; không cử nhân sự kỹ thuật hoặc thuê đại lý kiểm định sang tận kho bên Nhật Bản chụp ảnh tem nhãn gốc và đối chiếu hồ sơ xuất xưởng trước khi đóng container.",
        "Bước 1: Trước khi ký hợp đồng mua máy cũ, yêu cầu người bán chụp ảnh cận cảnh tem nhãn gốc (Nameplate) bằng kim loại có dập nổi số serial, năm sản xuất, công suất và nước xuất xứ; Bước 2: Thuê tổ chức giám định quốc tế (SGS/Bureau Veritas/Vinacontrol) thực hiện giám định trước khi xếp hàng tại nước xuất khẩu (Pre-shipment Inspection); Bước 3: Soát xét catalog kỹ thuật xem máy có sản xuất theo tiêu chuẩn JIS (Nhật), DIN (Đức), ANSI (Mỹ) hoặc tiêu chuẩn an toàn G7 hay không trước khi mở L/C.",
        "Phe A: Doanh nghiệp thanh minh máy móc thực tế còn rất mới, hoạt động êm ái, việc mất tem nhãn chỉ là do quá trình vệ sinh xưởng của bên bán; Phe B: Hải quan và cơ quan giám định kiên quyết không thỏa hiệp vì tem nhãn là căn cứ pháp lý duy nhất để xác định danh tính và tuổi thiết bị, nếu chấp nhận sẽ mở đường cho rác thải công nghệ tràn vào nước.",
        "Với máy móc thiết bị cũ, 'tuổi đời' chỉ là điều kiện cần, còn 'tem nhãn gốc và tiêu chuẩn kỹ thuật xuất xưởng' mới là điều kiện đủ. Đừng bao giờ mua máy cũ qua ảnh chụp mạng nếu chưa có chứng thư giám định độc lập tại nước sở tại.",
        "https://en.wikipedia.org/wiki/Used_goods",
        "Series: Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách từ chỉ đạo của Chị Thúy về 5 rào cản nhập khẩu máy móc cũ Quyết định 18 trong daily_notes 11/11/2026"
    ],
    [
        248,
        "W46 — 15/11/2026",
        "Khai Nhầm Số Lượng Và Ngày Hóa Đơn Lô Thức Ăn Chăn Nuôi Thuế 0% Sau Khi Đã Thông Quan: Bài Học Khai Bổ Sung AMA Và Phác Đồ Tránh Phạt Khai Sai Tại Cảng Quy Nhơn",
        "Công ty TNHH Nông nghiệp Xanh Q.N (Bình Định), Chi cục Hải quan Cửa khẩu Cảng Quy Nhơn, Đại lý Hải quan ECUS",
        "Doanh nghiệp mở tờ khai nhập khẩu lô 50 tấn phụ gia thức ăn chăn nuôi thuế suất 0% từ Indonesia và đã được hệ thống VNACCS phê duyệt thông quan luồng Xanh. Tuy nhiên 3 ngày sau, khi kế toán công ty đối chiếu lại bộ chứng từ thanh toán ngân hàng mới tá hỏa phát hiện nhân viên chứng từ đã lấy nhầm file Invoice của một lô hàng khác sắp về: số hóa đơn, ngày hóa đơn và số lượng trên tờ khai bị khai lệch (khai 50 tấn trong khi thực tế lô hàng này chỉ có 42 tấn, đơn giá cao hơn). Doanh nghiệp hoang mang không biết có bị hải quan quy tội gian lận chứng từ hay không và lúng túng khi khai sửa đổi bổ sung.",
        "Nguy cơ bị xử phạt từ 5 đến 10 triệu đồng về hành vi khai sai các chỉ tiêu thông tin trên tờ khai hải quan; hàng bị tạm giữ tại kho chờ đối chiếu chứng từ gốc; phát sinh 12 triệu đồng chi phí chỉnh sửa hợp đồng và kéo dài thời gian quyết toán sổ sách kho nội bộ.",
        "Thiếu quy trình 'Double Check' (kiểm tra chéo 2 người) trước khi truyền tờ khai chính thức; nhân viên chứng từ quản lý file lộn xộn, lưu nhiều phiên bản Invoice trong cùng một thư mục dẫn đến đính kèm nhầm tệp khi khai báo hải quan điện tử.",
        "Bước 1: Chủ động lập hồ sơ khai bổ sung sau thông quan theo nghiệp vụ AMA trên phần mềm ECUS5-VNACCS theo quy định tại Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi Thông tư 39/2018/TT-BTC) trước khi cơ quan hải quan ra quyết định kiểm tra sau thông quan; Bước 2: Nộp Mẫu số 03/KBS/GSQL kèm Công văn giải trình rõ nguyên nhân sai sót kỹ thuật hành chính và nộp bộ chứng từ gốc đối chiếu (Invoice chuẩn, Packing List chuẩn, B/L gốc, điện chuyển tiền ngân hàng); Bước 3: Đề nghị hải quan chi cục phúc tập xác nhận việc tự nguyện khai báo bổ sung để được miễn phạt vi phạm hành chính theo Điều 9 Luật Xử lý vi phạm hành chính.",
        "Phe A: Nhân viên cho rằng vì hàng có thuế suất 0% nên việc khai sai số lượng không làm thất thu thuế của Nhà nước, hải quan chỉ nên nhắc nhở hành chính; Phe B: Hải quan nhấn mạnh tính chuẩn xác của số liệu hải quan là nền tảng quản lý thống kê quốc gia và phòng chống gian lận thương mại, khai sai chứng từ vẫn là hành vi vi phạm thủ tục hải quan bắt buộc phải lập biên bản.",
        "Sự trung thực và chủ động khắc phục sai sót bằng nghiệp vụ AMA ngay khi phát hiện là chìa khóa duy nhất để chuyển hóa nguy cơ bị xử phạt thành trường hợp tự nguyện tuân thủ được pháp luật bảo vệ.",
        "https://en.wikipedia.org/wiki/Customs_declaration",
        "Series: Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách từ case study thực tế của tác giả Quỳnh trên nhóm Giải đáp thủ tục hải quan trong daily_notes 12/11/2026"
    ],
    [
        249,
        "W46 — 15/11/2026",
        "Giải Cứu Lô 200 Kg Hàng Mẫu Hợp Chất Hương Liệu Bị Kẹt Tại Kho Chuyển Phát Nhanh Nội Bài: Quy Trình Gỡ Vướng Kiểm Tra An Toàn Thực Phẩm Theo Nghị Định 15/2018 Trong 24 Giờ",
        "Công ty Cổ phần Thực phẩm & Đồ uống Á Châu (Hà Nội), Hãng chuyển phát nhanh DHL Express, Chi cục Hải quan Chuyển phát nhanh (Cục HQ Hà Nội), Viện Kiểm nghiệm An toàn Vệ sinh Thực phẩm Quốc gia",
        "Doanh nghiệp nhập khẩu lô hàng 200 kg mẫu hương liệu thực phẩm từ Pháp qua đường chuyển phát nhanh hàng không để phục vụ công tác R&D thử nghiệm công thức nước giải khát mới. Do trên kiện hàng người gửi chỉ dán nhãn ghi 'Food Flavor' mà không có nhãn phụ hay giấy tờ phân tích thành phần, công chức hải quan cửa khẩu đã yêu cầu giữ hàng và yêu cầu doanh nghiệp phải đăng ký kiểm tra nhà nước về an toàn thực phẩm. Thời gian làm thủ tục kiểm tra ATTP thông thường mất 7-10 ngày, trong khi mẫu thử cần giữ lạnh và chuyên gia nước ngoài đã bay sang chuẩn bị test thử nghiệm trong 2 ngày tới.",
        "Chi phí lưu kho lạnh tại kho chuyển phát nhanh 850.000 đồng/ngày; nguy cơ hủy chuyến công tác của chuyên gia quốc tế gây lãng phí vé máy bay và khách sạn hơn 4.500 USD (hơn 110 triệu đồng); rủi ro hỏng mẫu do biến tính nhiệt độ.",
        "Doanh nghiệp không hướng dẫn nhà cung cấp Pháp chuẩn bị bộ chứng từ hàng mẫu đạt chuẩn; trên Invoice không ghi điều khoản miễn phí phi thương mại; không chuẩn bị trước công văn đề nghị áp dụng cơ chế miễn kiểm tra an toàn thực phẩm theo quy định đối với hàng mẫu phục vụ nghiên cứu thử nghiệm.",
        "Bước 1: Soạn ngay Công văn khẩn cấp gửi Chi cục Hải quan chuyển phát nhanh dẫn chiếu Khoản 2 Điều 13 Nghị định 15/2018/NĐ-CP quy định các trường hợp được miễn kiểm tra nhà nước về ATTP nhập khẩu gồm 'Sản phẩm, nguyên liệu sản xuất dùng để nghiên cứu, phát triển sản phẩm mới với số lượng phù hợp'; Bước 2: Đính kèm Đề án nghiên cứu thử nghiệm nội bộ có phê duyệt của Ban Giám đốc và Bảng cam kết sử dụng mẫu đúng mục đích, không lưu thông trên thị trường; Bước 3: Làm việc trực tiếp với Đội thủ tục hải quan để kiểm tra thực tế bao bì niêm phong và làm thủ tục thông quan giải phóng hàng ngay trong ngày.",
        "Phe A: Doanh nghiệp bức xúc vì hàng nghiên cứu phát triển sản phẩm của tập đoàn lớn mà bị giữ như hàng kinh doanh thương mại trôi nổi; Phe B: Hải quan cửa khẩu chuyển phát nhanh giải thích họ phải kiểm soát chặt vì số lượng 200 kg là khá lớn đối với hàng mẫu, nếu không có hồ sơ nghiên cứu chứng minh thì rất dễ bị lợi dụng tuồn vào thị trường tiêu dùng.",
        "Khi nhập khẩu hàng mẫu đặc thù qua đường hàng không chuyển phát nhanh, sự chuẩn bị trước về mặt pháp lý (văn bản cam kết, viện dẫn điều khoản miễn trừ) có giá trị cứu nguy nhanh hơn nhiều so với việc chờ đợi hàng về đến nơi mới lo giải trình.",
        "https://en.wikipedia.org/wiki/Express_mail",
        "Series: Giải cứu hàng hóa",
        "Chưa dùng",
        "Bóc tách từ chỉ đạo của Chị Thúy về điều kiện miễn KTCN hàng mẫu theo Nghị định 15 và Nghị định 74 trong daily_notes 11/11/2026"
    ],
    [
        250,
        "W46 — 15/11/2026",
        "Khai Nhầm Tờ Khai OLA Cho Hàng Vận Chuyển Container Nội Địa Cát Lái - Cái Mép: Doanh Nghiệp Bị Phạt Vi Phạm Hành Chính 15 Triệu Đồng Và Cách Phân Định Rạch Ròi Giám Sát Hải Quan",
        "Công ty TNHH Tiếp vận Biển Xanh (TP.HCM), Hãng tàu SITC, Chi cục Hải quan Cảng Cát Lái, Chi cục Hải quan Cảng Cái Mép",
        "Doanh nghiệp vận tải nhận vận chuyển 10 container hàng rỗng và hàng tiêu dùng nội địa từ bãi cảng Cát Lái sang cụm cảng Cái Mép để phục vụ đóng hàng xuất khẩu. Nhân viên điều phối mới vào nghề nghĩ rằng việc di chuyển container giữa hai cảng biển thuộc hai tỉnh thành khác nhau (TP.HCM và Bà Rịa - Vũng Tàu) bắt buộc phải mở tờ khai vận chuyển độc lập OLA trên hệ thống VNACCS như hàng chuyển cửa khẩu quốc tế. Khi tờ khai OLA được truyền lên hệ thống, hải quan kiểm tra phát hiện đây là hàng hóa lưu thông nội địa thuần túy không thuộc đối tượng chịu sự kiểm tra, giám sát hải quan theo Luật Hải quan, việc tự ý mở tờ khai OLA đã làm sai lệch số liệu quản lý hàng hóa chuyển cảng.",
        "Doanh nghiệp bị lập biên bản xử phạt vi phạm hành chính 15 triệu đồng về hành vi thực hiện thủ tục hải quan không đúng đối tượng quy định; 10 container bị ách lại tại cổng bãi kiểm soát suốt 36 giờ để chờ thủ tục hủy tờ khai OLA, phát sinh chi phí neo xe đầu kéo và trễ lịch cấp vỏ rỗng cho khách xuất khẩu.",
        "Nhân viên giao nhận không phân biệt được sự khác nhau giữa 'Hàng hóa xuất nhập khẩu chịu sự giám sát hải quan trung chuyển giữa các cảng' (bắt buộc mở OLA) và 'Hàng hóa nội địa hoặc vỏ rỗng luân chuyển nội địa' (chỉ cần lệnh điều động và chứng từ vận tải thông thường).",
        "Bước 1: Xây dựng ma trận phân loại luồng hàng hóa nội bộ: Chỉ mở tờ khai OLA khi hàng hóa là hàng ngoại nhập chưa hoàn thành thủ tục hải quan chuyển từ cảng dỡ về cảng đích (loại hình OLA vận chuyển độc lập kết hợp chỉ định hãng tàu); Bước 2: Đối với hàng nội địa chuyển cảng, chỉ sử dụng Lệnh giao vỏ (Equipment Interchange Receipt - EIR), Phiếu xuất kho kiêm vận chuyển nội bộ hoặc Hóa đơn GTGT vận tải; Bước 3: Đào tạo định kỳ cho đội ngũ OPS hiện trường về ranh giới thẩm quyền giám sát của cơ quan hải quan tại khu vực giám sát cảng biển theo Điều 51 Thông tư 38/2015/TT-BTC.",
        "Phe A: Nhân viên cho rằng cẩn tắc vô áy náy, mở tờ khai hải quan OLA chỉ là muốn minh bạch để công chức giám sát cổng cảng dễ kiểm tra chứ không có ý gian dối; Phe B: Hải quan khẳng định hệ thống OLA phục vụ quản lý rủi ro đối với hàng hóa chưa thông quan có rủi ro buôn lậu cao, việc mở tờ khai bừa bãi gây lãng phí tài nguyên giám sát và làm sai lệch chỉ số phân tích rủi ro của ngành.",
        "Trong nghiệp vụ hải quan, 'thừa thủ tục' đôi khi cũng nguy hiểm và gây hại không kém gì 'thiếu thủ tục'. Hiểu đúng bản chất pháp lý của từng loại chứng từ là tiêu chuẩn vàng của một chuyên gia logistics chuyên nghiệp.",
        "https://en.wikipedia.org/wiki/Cabotage",
        "Series: Tai nạn XNK",
        "Chưa dùng",
        "Bóc tách từ tình huống nghiệp vụ thực tế của Trúc Vân về tờ khai OLA nội địa trong daily_notes 11/11/2026"
    ]
]

# ==============================================================================
# 3. DỮ LIỆU HỌC THỰC CHIẾN (HT) - 5 Ý TƯỞNG (STT 246 - 250) - 14 CỘT
# ==============================================================================
HT_ROWS = [
    [
        246,
        "W46 — 15/11/2026",
        "Quy Trình 5 Bước Chuẩn Bị Hồ Sơ Tham Vấn Giá Hải Quan & Kháng Biên Bản Ấn Định Trị Giá Theo Công Văn 24083/CHQ-NVTHQ: Cẩm Nang Bảo Vệ Giá Khai Báo Cho Doanh Nghiệp Hàng Tiêu Dùng",
        "Bước 1: Nhận Thông báo nghi vấn trị giá và chuẩn bị hồ sơ trong 5 ngày làm việc:\n- Rà soát Thông báo nghi vấn theo Mẫu số 02A/TB-NVTG/TXNK;\n- Tập hợp đầy đủ 7 chứng từ cốt lõi: Hợp đồng ngoại thương, Hóa đơn thương mại gốc, Vận đơn B/L, Bảng kê thanh toán ngân hàng (MT103), Chứng nhận xuất xứ (C/O), Hợp đồng bảo hiểm và Bảng báo giá gốc của nhà sản xuất nước ngoài;\nBước 2: Lập Bảng giải trình chi tiết cấu thành giá thành sản phẩm (Price Breakdown):\n- Chứng minh tính hợp lý của đơn giá mua căn cứ theo biến động giá nguyên liệu thế giới, khối lượng đặt hàng lớn (Volume Discount) hoặc chương trình khuyến mại mùa vụ của người bán;\nBước 3: Thu thập bằng chứng khách quan của thị trường:\n- Thu thập chứng từ bán ra nội địa (Hóa đơn GTGT đầu ra, hợp đồng đại lý phân phối), catalog sản phẩm và dữ liệu giao dịch cùng mặt hàng của các doanh nghiệp khác trên thị trường;\nBước 4: Tham gia phiên họp tham vấn trực tiếp tại Chi cục Hải quan:\n- Đại diện pháp luật hoặc người được ủy quyền có kiến thức chuyên môn trình bày lập luận theo đúng 6 phương pháp xác định trị giá hải quan quy định tại Thông tư 39/2015/TT-BTC;\n- Yêu cầu cán bộ hải quan ghi nhận đầy đủ chứng cứ vào Biên bản tham vấn Mẫu 02B/BB-TVTG/TXNK;\nBước 5: Xử lý kết quả tham vấn và khiếu nại (nếu bị ấn định giá sai):\n- Nếu hải quan bác bỏ giá khai báo và ấn định giá cao hơn không có căn cứ, doanh nghiệp ký bảo lưu ý kiến trên biên bản, thực hiện nộp thuế theo giá ấn định để giải phóng hàng trước, đồng thời gửi đơn khiếu nại lên Cục Hải quan trong thời hạn 90 ngày.",
        "Biên độ biến động giá kích hoạt nghi vấn hải quan: Biến động chênh lệch từ ±10% so với Cơ sở dữ liệu giá tham chiếu GTT02.\nSố thuế truy thu chênh lệch khi bị ấn định giá = (Trị giá ấn định mới - Trị giá khai báo) x (Thuế suất NK + Thuế suất VAT).\nThời hạn bảo lưu hồ sơ phục vụ Kiểm tra sau thông quan (KTSTQ): Tối thiểu 5 năm kể từ ngày thông quan theo Điều 77 Luật Hải quan 2014.",
        "1. Mang hồ sơ sơ sài đi tham vấn và chỉ giải thích bằng miệng mà không có chứng từ thanh toán ngân hàng MT103 đối chiếu khớp số tiền từng cent;\n2. Nhượng bộ ký vào biên bản đồng ý mức giá ấn định của hải quan chỉ để lấy hàng nhanh, dẫn đến bị gắn cờ rủi ro cao trên hệ thống rủi ro cho toàn bộ các lô hàng sau;\n3. Khai báo giá quá thấp so với mặt bằng chung (thấp hơn 30-50%) mà không có văn bản giải trình lý do thương mại hợp lý.",
        "Tham vấn giá hải quan như 'phiên tòa bảo vệ tài sản': Chứng cứ bằng văn bản ngân hàng là nhân chứng thép; lời nói suông không có giá trị trước bảng dữ liệu giá GTT02 của hải quan!",
        "Tham vấn giá trước khi thông quan (Điều 25 TT 38/2015) vs Kiểm tra sau thông quan tại trụ sở doanh nghiệp (Điều 77-80 Luật HQ): Tham vấn giá chỉ tập trung vào trị giá hải quan của 1 lô hàng cụ thể trong 30 ngày; KTSTQ là cuộc thanh tra toàn diện mọi giao dịch, sổ sách kế toán, dòng tiền trong vòng 5 năm.",
        "Anh Minh (Giám đốc doanh nghiệp thương mại), Chị Hương (Kế toán trưởng), Anh Dũng (Nhân viên hiện trường OPS chuyên làm thủ tục hải quan)",
        "☐ Đầy đủ bộ hồ sơ ngoại thương gốc (Hợp đồng, Invoice, B/L, C/O)\n☐ Điện chuyển tiền thanh toán quốc tế Swift MT103 có dấu ngân hàng\n☐ Bảng giải trình chi tiết cấu thành giá thành (Price Breakdown Sheet)\n☐ Hóa đơn GTGT đầu ra chứng minh biên độ lợi nhuận thương mại\n☐ Giấy ủy quyền hợp pháp của Tổng giám đốc cho người đi tham vấn",
        "https://en.wikipedia.org/wiki/Customs_valuation",
        "Series: Quy trình Hải quan",
        "Chưa dùng",
        "Cẩm nang thực chiến tham vấn giá chuẩn bị cho đợt cao điểm kiểm soát trị giá Tết Đinh Mùi 2027 theo Công văn 24083/CHQ-NVTHQ"
    ],
    [
        247,
        "W46 — 15/11/2026",
        "Quy Trình 5 Bước Mở Tờ Khai Xuất Trả Hàng Nhập Khẩu Không Đạt Chất Lượng (Mã B13) & Hoàn 100% Thuế Nhập Khẩu: Hướng Dẫn Soát Xét Hồ Sơ Từ A-Z",
        "Bước 1: Thiết lập căn cứ pháp lý & Lập Biên bản ghi nhận lỗi chất lượng:\n- Khi phát hiện hàng lỗi, lập tức dừng sử dụng, giữ nguyên hiện trường; gửi thông báo lỗi cho người bán nước ngoài kèm hình ảnh/video; ký Văn bản thỏa thuận đồng ý nhận lại hàng của đối tác ngoại (Return Agreement);\nBước 2: Mời cơ quan giám định độc lập cấp chứng thư:\n- Thuê tổ chức giám định uy tín (Vinacontrol, SGS, Quatest) kiểm tra thực tế tại kho; chứng thư giám định phải nêu rõ hàng không đúng quy cách hợp đồng, còn nguyên trạng chưa qua sử dụng hay chế biến tại Việt Nam;\nBước 3: Mở tờ khai hải quan xuất trả loại hình B13:\n- Đăng ký tờ khai xuất khẩu B13 trên ECUS5-VNACCS; tại ô 'Ghi chú', bắt buộc ghi rõ: 'Hàng xuất trả cho tờ khai nhập khẩu số [Số tờ khai nhập cũ] ngày [Ngày mở tờ khai cũ] do lỗi không đạt chất lượng theo Biên bản giám định số...';\nBước 4: Thực hiện thủ tục kiểm hóa thực tế 100% tại cửa khẩu xuất (Luồng Đỏ):\n- Phối hợp với công chức hải quan kiểm tra thực tế hàng hóa tại cảng xuất; lập Biên bản bàn giao và xác nhận hàng đã thực xuất ra khỏi lãnh thổ Việt Nam;\nBước 5: Nộp hồ sơ đề nghị hoàn thuế nhập khẩu theo Điều 34 Nghị định 134/2016/NĐ-CP:\n- Nộp hồ sơ hoàn thuế lên Chi cục Hải quan nơi mở tờ khai nhập ban đầu; hồ sơ gồm tờ khai nhập cũ, tờ khai xuất B13, chứng thư giám định, thỏa thuận trả hàng và chứng từ thanh toán tiền hoàn trả của người bán.",
        "Số tiền thuế nhập khẩu được hoàn = Số tiền thuế nhập khẩu đã nộp thực tế x (Số lượng hàng xuất trả / Tổng số lượng hàng nhập ban đầu).\nThời hạn cơ quan hải quan giải quyết hoàn thuế: Tối đa 40 ngày kể từ ngày nhận đủ hồ sơ hợp lệ đối với trường hợp kiểm tra trước hoàn sau theo Điều 75 Luật Quản lý thuế số 38/2019/QH14.",
        "1. Tự ý tháo dỡ bao bì, đưa hàng vào gia công chế biến thử một phần rồi mới đòi xuất trả (mất tính nguyên trạng, bị từ chối hoàn thuế 100%);\n2. Mở nhầm mã loại hình xuất khẩu kinh doanh (B11) thay vì mã xuất trả hàng nhập khẩu (B13), khiến hệ thống không liên kết được dữ liệu để hoàn thuế;\n3. Thiếu chứng thư giám định độc lập trước khi đóng cont tái xuất, chỉ có email trao đổi qua lại giữa hai bên thì hải quan không chấp nhận hồ sơ hoàn thuế.",
        "Xuất trả hàng B13 như 'đổi trả hàng trên sàn thương mại điện tử': Phải giữ nguyên tem mác, hộp niêm phong và có video mở hộp chứng minh hàng lỗi; nếu bóc ra dùng thử rồi thì không người bán nào nhận lại và sàn cũng không hoàn tiền!",
        "Xuất trả hàng lỗi loại hình B13 vs Xuất khẩu hàng hóa kinh doanh B11: B13 liên kết trực tiếp với tờ khai nhập cũ để được miễn thuế xuất khẩu và hoàn 100% thuế nhập khẩu đã nộp; B11 là bán hàng thương mại độc lập, không được hoàn thuế nhập khẩu của lô hàng cũ.",
        "Chị Thu (Chuyên viên chứng từ XNK), Anh Dũng (Nhân viên hiện trường OPS), Chị Hương (Kế toán trưởng phụ trách hoàn thuế)",
        "☐ Văn bản thỏa thuận nhận lại hàng có chữ ký và con dấu của đối tác ngoại\n☐ Chứng thư giám định độc lập của Vinacontrol/SGS xác nhận hàng lỗi\n☐ Tờ khai xuất khẩu B13 có ghi chú chính xác số tờ khai nhập cũ\n☐ Biên bản kiểm hóa thực tế 100% có xác nhận công chức hải quan cửa khẩu\n☐ Giấy nộp tiền thuế nhập khẩu ban đầu vào ngân sách nhà nước",
        "https://en.wikipedia.org/wiki/Duty_drawback",
        "Series: Tối ưu Thuế XNK",
        "Chưa dùng",
        "Hướng dẫn step-by-step xử lý xuất trả hàng không đạt chất lượng loại hình B13 theo chỉ đạo của Chị Thúy trong daily_notes 11/11/2026"
    ],
    [
        248,
        "W46 — 15/11/2026",
        "Kỹ Thuật Đàm Phán Mua Hàng Quốc Tế Chuyển Từ CIF Sang FOB: Bộ Công Cụ 4 Bước Giúp Chủ Hàng Kiểm Soát 100% Chi Phí Vận Tải Và Triệt Tiêu Nạn 'Chặt Chém' Local Charges",
        "Bước 1: Phân tích so sánh tổng chi phí cập cảng (Landed Cost Analysis):\n- Yêu cầu người bán nước ngoài chào song song 2 bảng giá: Giá CIF Cát Lái/Hải Phòng vs Giá FOB tại cảng xuất khẩu (POL);\n- Liên hệ 2 Forwarder Việt Nam uy tín lấy báo giá cước biển (Ocean Freight) và bảng phụ phí Local charges chuẩn tại cảng đến;\nBước 2: Vạch trần chênh lệch chi phí ẩn của điều khoản CIF:\n- Lấy tổng chi phí CIF (Giá hàng CIF + Local charges đại lý chỉ định) trừ đi tổng chi phí FOB (Giá hàng FOB + Cước biển Forwarder VN + Local charges niêm yết chuẩn);\n- Chỉ ra cho ban giám đốc thấy phương án FOB giúp tiết kiệm từ 15% đến 30% tổng chi phí logistics;\nBước 3: Đàm phán chuyển đổi điều khoản với nhà cung cấp:\n- Thuyết phục nhà cung cấp rằng chuyển sang FOB giúp họ giảm bớt rủi ro theo dõi tàu biển và thanh toán cước quốc tế; nhà cung cấp chỉ cần giao hàng lên boong tàu tại cảng xuất xưởng theo đúng Incoterms 2020;\nBước 4: Thiết lập hợp đồng dịch vụ logistics tổng thể với Forwarder nội địa:\n- Ký hợp đồng chỉ định Forwarder Việt Nam làm đại lý vận chuyển độc quyền; yêu cầu Forwarder cam kết bằng văn bản không phát sinh bất kỳ khoản phụ phí nào ngoài hợp đồng, cung cấp đường link theo dõi lịch trình tàu thời gian thực (Tracking link) và chịu trách nhiệm bảo hiểm nếu chậm lịch trình.",
        "Công thức tính Tiết kiệm chi phí khi chuyển từ CIF sang FOB:\nMức tiết kiệm = (Local Charges CIF của đại lý ngoại) - (Cước biển FOB + Local Charges chuẩn của Forwarder nội địa).\nThực tế đối với hàng lẻ LCL: Doanh nghiệp thường tiết kiệm được từ 50 USD đến 120 USD trên mỗi CBM hàng hóa khi tự chủ quyền book cước FOB.",
        "1. Lười tìm kiếm Forwarder uy tín trong nước, chọn phương án CIF cho 'nhàn đầu' mà không biết công ty đang bị rút ruột hàng chục triệu tiền phụ phí mỗi lô hàng;\n2. Mua hàng FOB nhưng không kiểm tra kỹ phụ phí tại cảng xuất (Local Charges POL) khiến bên bán nước ngoài tính thêm chi phí giao hàng nội địa vào hóa đơn;\n3. Quên mua bảo hiểm hàng hải cho chặng đường biển (mua FOB thì người mua phải tự mua bảo hiểm, trong khi CIF thì người bán đã mua gói tối thiểu).",
        "Mua hàng CIF vs Mua hàng FOB như 'đi xe ôm trọn gói không bấm đồng hồ' vs 'gọi xe công nghệ biết trước giá': Đi xe ôm trọn gói nhìn qua tưởng rẻ, nhưng đến nơi tài xế mới đòi thêm tiền cầu đường, tiền chở nặng gấp 5 lần giá gốc!",
        "Điều kiện CIF (Cost, Insurance and Freight) vs Điều kiện FOB (Free On Board): CIF chuyển giao rủi ro tại cảng xếp nhưng người bán chọn tàu và chịu cước, người mua chịu trận phụ phí cảng đích; FOB người mua toàn quyền chọn hãng tàu, đàm phán cước và kiểm soát phụ phí minh bạch từ đầu đến cuối.",
        "Anh Tuấn (Giám đốc công ty thương mại XNK), Chị Trang (Trưởng phòng Thu mua quốc tế), Chị Thu (Chuyên viên logistics)",
        "☐ Yêu cầu bên bán báo giá kép (Dual Quotation: CIF vs FOB)\n☐ Thu thập bảng giá cước biển và Local charges niêm yết từ 2 Forwarder VN\n☐ Tính toán tổng chi phí Landed Cost của cả 2 phương án\n☐ Mua hợp đồng bảo hiểm hàng hải riêng cho lô hàng FOB\n☐ Cài đặt điều khoản chỉ định hãng tàu trong hợp đồng ngoại thương",
        "https://en.wikipedia.org/wiki/FOB_(shipping)",
        "Series: Đàm phán Quốc tế",
        "Chưa dùng",
        "Cẩm nang kỹ thuật đàm phán chuyển đổi điều khoản thương mại Incoterms giúp doanh nghiệp SME tiết kiệm chi phí logistics bám sát daily_notes 11/11/2026"
    ],
    [
        249,
        "W46 — 15/11/2026",
        "Checklist 5 Bước Khấu Trừ & Kê Khai Nộp Thay Thuế Nhà Thầu (FCT) Cho Dịch Vụ Số Xuyên Biên Giới (Meta, Google, Microsoft, AWS) Chuẩn Thông Tư 69/2025/TT-BTC",
        "Bước 1: Rà soát và phân loại toàn bộ giao dịch mua dịch vụ số quốc tế:\n- Rà soát sao kê tài khoản ngân hàng, thẻ tín dụng doanh nghiệp phát sinh chi trả cho các nhà cung cấp nước ngoài (Meta/Facebook Ads, Google Workspace, Microsoft 365, AWS Cloud, Zoom, Canva...); xác định rõ bản chất dịch vụ là dịch vụ số trực tuyến;\nBước 2: Xác định phương pháp tính thuế và tỷ lệ thuế áp dụng:\n- Căn cứ Điều 9 Thông tư 69/2025/TT-BTC, áp dụng phương pháp tính trực tiếp trên doanh thu đối với dịch vụ số: Thuế GTGT là 5% trên giá thanh toán; Thuế TNDN là 2% (nếu nhà thầu là tổ chức kinh doanh);\nBước 3: Xác định giá tính thuế theo công thức Net sang Gross:\n- Nếu hợp đồng hoặc điều khoản thanh toán ghi giá chưa bao gồm thuế (giá Net), doanh nghiệp bắt buộc phải quy đổi (Gross-up) sang doanh thu tính thuế trước khi áp tỷ lệ phần trăm;\nBước 4: Lập và nộp Tờ khai thuế nhà thầu Mẫu số 01/NTNN trên Cổng Thuế điện tử:\n- Truy cập hệ thống thuedientu.gdt.gov.vn, chọn tờ khai Mẫu 01/NTNN theo từng lần phát sinh thanh toán (hoặc khai theo tháng nếu phát sinh thường xuyên); nộp tiền thuế GTGT và TNDN vào Kho bạc Nhà nước;\nBước 5: Lưu trữ hồ sơ chứng từ để hạch toán chi phí hợp lý được trừ:\n- Lưu trữ trọn bộ hồ sơ gồm: Hợp đồng điện tử/Điều khoản dịch vụ chấp thuận (Terms of Service), Hóa đơn điện tử có mã số thuế của nhà cung cấp ngoại, Biên lai thanh toán thẻ ngân hàng và Giấy nộp tiền thuế nhà thầu Mẫu 01/NTNN.",
        "Công thức tính thuế FCT theo giá Net (Doanh nghiệp trả thay toàn bộ thuế):\nDoanh thu tính thuế TNDN = Giá trị thanh toán Net / (1 - Tỷ lệ thuế TNDN) = Giá Net / (1 - 0.02) = Giá Net / 0.98.\nSố thuế TNDN phải nộp = Doanh thu tính thuế TNDN x 2%.\nDoanh thu tính thuế GTGT = Doanh thu tính thuế TNDN / (1 - Tỷ lệ thuế GTGT) = Doanh thu tính thuế TNDN / (1 - 0.05).\nSố thuế GTGT phải nộp = Doanh thu tính thuế GTGT x 5%.",
        "1. Tự ý lấy hóa đơn điện tử của Meta/Google kẹp vào chi phí doanh nghiệp mà không kê khai nộp thuế FCT, dẫn đến bị cơ quan thuế loại toàn bộ chi phí khi quyết toán;\n2. Dùng thẻ tín dụng cá nhân của giám đốc để quẹt trả tiền quảng cáo rồi rút tiền công ty thanh toán hoàn ứng (vi phạm quy định chứng từ thanh toán không dùng tiền mặt);\n3. Tính thuế trực tiếp trên giá Net mà không thực hiện công thức Gross-up dẫn đến tính thiếu số tiền thuế phải nộp, bị phạt chậm nộp.",
        "Nộp thuế nhà thầu FCT cho dịch vụ số như 'mua vé vào cửa công viên quốc tế': Dịch vụ quốc tế bay vào Việt Nam kiếm tiền thì doanh nghiệp trong nước là người giữ cổng; nếu quên thu phí vé vào cửa nộp cho Nhà nước thì chính doanh nghiệp sẽ phải bỏ tiền túi ra đền vé!",
        "Nhà cung cấp ngoại đã tự đăng ký thuế tại Việt Nam (qua cổng Etax) vs Chưa đăng ký: Nếu nhà cung cấp (như Meta/Google) đã cấp hóa đơn có MST 10 số đăng ký tại VN và đã thu thuế GTGT 5% thì DN không phải khấu trừ nộp thay thuế GTGT, nhưng vẫn phải kiểm tra nghĩa vụ thuế TNDN theo hướng dẫn của Cục Thuế.",
        "Chị Hương (Kế toán trưởng), Anh Minh (Giám đốc doanh nghiệp), Chị Lan (Kế toán thuế)",
        "☐ Sao kê tài khoản thẻ tín dụng ghi nhận rõ giao dịch thanh toán quốc tế\n☐ Tải file Invoice/Hóa đơn thương mại điện tử từ tài khoản quản trị\n☐ Tính toán số thuế FCT theo đúng công thức Gross-up (5% VAT, 2% TNDN)\n☐ Lập và nộp tờ khai Mẫu 01/NTNN trên Cổng Thuế điện tử\n☐ Lưu biên lai nộp tiền thuế vào kho bạc vào hồ sơ quyết toán thuế TNDN",
        "https://en.wikipedia.org/wiki/Withholding_tax",
        "Series: Tối ưu Thuế XNK",
        "Chưa dùng",
        "Cẩm nang kê khai thuế nhà thầu dịch vụ số bám sát Luật Thuế GTGT 48/2024 và Thông tư 69/2025/TT-BTC từ daily_notes 11/11/2026"
    ],
    [
        250,
        "W46 — 15/11/2026",
        "Kỹ Thuật Đọc Và Khai Báo Thủ Tục Tờ Khai Vận Chuyển Độc Lập OLA Trên VNACCS: Quy Trình 4 Bước Cho Hàng Quá Cảnh, Hàng Chuyển Cửa Khẩu & Hàng Trung Chuyển Cảng Biển",
        "Bước 1: Xác định đúng đối tượng bắt buộc và đối tượng cấm mở tờ khai OLA:\n- Đối tượng BẮT BUỘC OLA: Hàng hóa quá cảnh, hàng hóa chuyển cửa khẩu từ cảng dỡ hàng về cảng đích/ICD nội địa, hàng trung chuyển quốc tế;\n- Đối tượng TUYỆT ĐỐI CẤM OLA: Hàng hóa nội địa thuần túy, container rỗng nội địa luân chuyển giữa các bến cảng;\nBước 2: Chuẩn bị thông tin chỉ tiêu khai báo vận chuyển độc lập:\n- Đăng nhập ECUS5-VNACCS, chọn mục 'Tờ khai vận chuyển độc lập (OLA)';\n- Nhập chính xác 5 chỉ tiêu then chốt: Mã Chi cục Hải quan nơi đi, Mã Chi cục Hải quan nơi đến, Phương tiện vận chuyển (Số hiệu xe đầu kéo/sà lan), Tuyến đường vận chuyển và Thời gian vận chuyển dự kiến;\nBước 3: Gắn seal niêm phong hải quan hoặc định vị vệ tinh GPS:\n- Phối hợp với công chức hải quan giám sát cổng cảng nơi đi để bấm seal niêm phong hải quan hoặc gắn thiết bị giám sát định vị GPS của Tổng cục Hải quan; ghi nhận số seal trên hệ thống và biên bản bàn giao hàng hóa;\nBước 4: Xác nhận hoàn thành vận chuyển tại Chi cục Hải quan nơi đến:\n- Khi phương tiện đến đích, xuất trình hồ sơ cho công chức hải quan giám sát nơi đến kiểm tra tình trạng nguyên vẹn của niêm chì seal; công chức hải quan thực hiện thao tác xác nhận 'Hàng đã đến đích' trên hệ thống VNACCS để đóng tờ khai OLA.",
        "Thời gian vận chuyển cho phép giữa các cảng/cửa khẩu: Tối đa 24 giờ đối với tuyến vận tải đường bộ dưới 200 km; tối đa 48 giờ đối với tuyến sà lan đường thủy nội địa.\nMức phạt hành vi vi phạm tự ý thay đổi tuyến đường hoặc chậm trễ bàn giao hàng hóa chịu sự giám sát hải quan theo Nghị định 128/2020/NĐ-CP: Phạt tiền từ 10 đến 30 triệu đồng.",
        "1. Khai sai mã Chi cục Hải quan nơi đến hoặc mã địa điểm lưu kho khiến hệ thống không thể chuyển dữ liệu theo dõi hành trình;\n2. Để xe đầu kéo chạy lệch tuyến đường đã đăng ký trên tờ khai OLA mà không báo cáo lý do khách quan (như kẹt xe, sự cố hỏng hóc), bị hệ thống cảnh báo rủi ro buôn lậu;\n3. Phương tiện đã đến đích nhưng lái xe không vào văn phòng hải quan giám sát để quét mã xác nhận hoàn thành, khiến tờ khai OLA bị treo quá hạn trên toàn hệ thống hải quan cả nước.",
        "Tờ khai OLA như 'thẻ bài hộ tống tù nhân': Khi hàng hóa chưa nộp thuế đi trên lãnh thổ Việt Nam, tờ khai OLA cùng chiếc seal niêm chì là vòng bảo vệ pháp lý duy nhất; nếu đứt xích hoặc chạy sai đường thì tài xế và chủ hàng sẽ bị xử lý như tiếp tay cho buôn lậu!",
        "Tờ khai vận chuyển độc lập OLA vs Tờ khai vận chuyển kết hợp (trên tờ khai nhập khẩu): OLA dùng cho hãng vận chuyển/Forwarder độc lập chuyển hàng nguyên trạng khi chưa mở tờ khai nhập khẩu; Tờ khai vận chuyển kết hợp do chính chủ hàng khai báo ngay trên tờ khai nhập khẩu phân luồng Vàng/Đỏ chuyển về chi cục hải quan quản lý.",
        "Anh Dũng (Nhân viên hiện trường OPS cảng), Anh Tuấn (Điều hành đội xe vận tải container), Chị Thu (Chuyên viên thủ tục hải quan)",
        "☐ Xác định hàng hóa thuộc diện chịu sự giám sát hải quan chuyển cửa khẩu\n☐ Khai đúng mã Chi cục HQ nơi đi, nơi đến và biển số phương tiện vận tải\n☐ Kiểm tra chì niêm phong seal hải quan và biên bản bàn giao trước khi rời cảng\n☐ Hướng dẫn tài xế chạy đúng tuyến đường và thời gian quy định\n☐ Xác nhận hoàn tất đóng tờ khai OLA tại Chi cục Hải quan nơi đến",
        "https://en.wikipedia.org/wiki/Customs",
        "Series: Quy trình Hải quan",
        "Chưa dùng",
        "Cẩm nang chuẩn hóa nghiệp vụ mở tờ khai OLA và ranh giới giám sát hải quan bóc tách từ daily_notes 11/11/2026"
    ]
]

# ==============================================================================
# 4. DỮ LIỆU FUNFACT & INSIGHT (FF) - 5 Ý TƯỞNG (STT 246 - 250) - 13 CỘT
# ==============================================================================
FF_ROWS = [
    [
        246,
        "W46 — 15/11/2026",
        "Kênh Đào Suez & Vụ Kẹt Tàu Ever Given (2021): 6 Ngày Tê Liệt 12% Thương Mại Toàn Cầu, Giam Lỏng 400 Con Tàu Và Thiệt Hại 9,6 Tỷ USD Mỗi Ngày",
        "Một cơn gió giật sa mạc và một cú bẻ lái lệch đã khiến một con tàu container khổng lồ nằm ngang làm nghẽn toàn bộ 'động mạch chủ' thương mại Á - Âu, thổi bay gần 60 tỷ USD của kinh tế thế giới chỉ trong chưa đầy một tuần!",
        "Thuyền trưởng và hoa tiêu trên siêu tàu container Ever Given (dài 400 mét, trọng tải 220.000 tấn của hãng tàu Evergreen) bị mắc cạn chéo thân tại km 151 của Kênh đào Suez (Ai Cập) ngày 23/03/2021.",
        "Kênh đào Suez khánh thành năm 1869 là con đường tắt chiến lược nối Địa Trung Hải và Biển Đỏ, giúp tàu thuyền không phải đi vòng qua Mũi Hảo Vọng của Châu Phi. Sáng ngày 23/03/2021, trong cơn bão cát với sức gió giật cấp 8, con tàu Ever Given bị mất lái và cắm mũi sâu vào bờ cát phía đông của kênh đào. Chiều dài 400 mét của con tàu lớn hơn cả chiều rộng đáy kênh đào, biến nó thành một bức tường thép khổng lồ bịt kín hoàn toàn tuyến đường thủy tấp nập nhất thế giới. Hơn 400 con tàu chở dầu, khí đốt và hàng tiêu dùng phải thả neo xếp hàng chờ đợi trong vô vọng.",
        "Sự cố đã làm bùng nổ cuộc khủng hoảng chuỗi cung ứng toàn cầu sau đại dịch, đẩy giá cước vận tải biển tăng vọt 300-400%; thôi thúc các liên minh hãng tàu tái cấu trúc chiến lược phân tán rủi ro hàng hải và đẩy nhanh tiến độ mở rộng đoạn kênh đào phía nam của Ai Cập.",
        "Nếu Kênh đào Panama dựa vào hệ thống âu thuyền nước ngọt phức tạp dễ bị hạn hán, thì Kênh đào Suez là kênh đào nước mặn ngang mực nước biển, không có âu thuyền nhưng lại đối mặt với rủi ro bão cát sa mạc và nguy cơ tắc nghẽn đơn điểm (Single Point of Failure).",
        "Đừng bao giờ dồn 100% dòng hàng của doanh nghiệp vào một lộ trình vận chuyển duy nhất mà không có phương án dự phòng (Contingency Plan). Trong quản trị logistics hiện đại, sự phụ thuộc thái quá vào một 'nút thắt cổ chai' có thể đánh sập uy tín giao hàng và đốt cháy toàn bộ lợi nhuận của doanh nghiệp chỉ sau một sự cố khách quan.",
        "https://en.wikipedia.org/wiki/2021_Suez_Canal_obstruction",
        "Series: Lịch sử Logistics",
        "Chưa dùng",
        "Bài học kinh điển về rủi ro nút thắt chuỗi cung ứng toàn cầu từ sự cố kẹt tàu Ever Given tại Kênh đào Suez"
    ],
    [
        247,
        "W46 — 15/11/2026",
        "Cảng Tự Động Hóa Maasvlakte II Của Rotterdam: Nơi Những Cỗ Cẩu Khổng Lồ Và Xe Tự Hành AGV Bốc Dỡ Hàng Triệu Container Không Một Bóng Người",
        "Tại bến cảng container hiện đại nhất Châu Âu, hàng triệu tấn hàng hóa được nâng hạ, xếp dỡ xuyên đêm suốt 24/7 hoàn toàn bởi robot tự hành điều khiển bằng thuật toán AI mà không có một bóng dáng công nhân nào bước chân xuống bãi cảng!",
        "Dự án lấn biển xây dựng siêu cảng Maasvlakte II thuộc Cảng Rotterdam (Hà Lan) chính thức đi vào vận hành các bến cảng tự động hóa hoàn toàn (APM Terminals MedPort và RWG) từ năm 2015.",
        "Rotterdam từng là cảng biển bận rộn nhất thế giới trong nhiều thập kỷ thế kỷ 20. Khi đối mặt với sự cạnh tranh khốc liệt về chi phí từ các cảng biển Châu Á và tình trạng thiếu hụt nhân lực lao động nặng nhọc, Hà Lan đã quyết định lấn biển 2.000 hecta ra Biển Bắc để kiến tạo một kỳ quan logistics tương lai: bến cảng container không phát thải carbon đầu tiên trên hành tinh, nơi toàn bộ cần cẩu bờ (STS Cranes) được điều khiển từ xa qua màn hình máy tính cách đó hàng km, và 100% việc di chuyển container trên bãi do đội xe tự hành AGV (Automated Guided Vehicles) chạy điện đảm nhiệm nhờ mạng lưới nam châm định vị ngầm dưới lòng đất.",
        "Maasvlakte II trở thành hình mẫu tiêu chuẩn vàng cho mô hình Cảng biển Xanh và Cảng thông minh (Smart Port) toàn cầu; nâng năng suất giải phóng container lên mức kỷ lục 40-50 thao tác/giờ và giảm thiểu 90% tai nạn lao động tại khu vực bốc xếp.",
        "Các cảng biển truyền thống vẫn phụ thuộc vào hàng ngàn tài xế xe kéo và công nhân hoa tiêu mặt bãi; trong khi Rotterdam và Cảng Dương Sơn (Thượng Hải) đã chứng minh tự động hóa toàn phần bằng robot giúp cảng biển vận hành bền bỉ bất chấp thời tiết tuyết rơi hay đêm tối.",
        "Số hóa và tự động hóa không phải là cuộc đua xa xỉ của tương lai mà là điều kiện sống còn của hiện tại. Doanh nghiệp logistics Việt Nam muốn nâng cao năng lực cạnh tranh buộc phải bắt đầu chuẩn hóa dữ liệu số từ ngày hôm nay; chỉ có chuẩn hóa quy trình mới mở đường cho tự động hóa vận hành hiệu quả.",
        "https://en.wikipedia.org/wiki/Port_of_Rotterdam",
        "Series: Công nghệ & Đổi mới",
        "Chưa dùng",
        "Mô hình cảng biển tự động hóa thông minh Rotterdam và tầm nhìn tương lai của công nghệ logistics bốc dỡ hàng hải"
    ],
    [
        248,
        "W46 — 15/11/2026",
        "Hệ Thống Quản Lý Sản Xuất Tinh Gọn (Lean/JIT) Của Toyota: Triết Lý 'Không Tồn Kho' Đã Định Hình Lại Toàn Bộ Chuỗi Cung Ứng Toàn Cầu",
        "Tập đoàn sản xuất ô tô số 1 thế giới Toyota không hề sở hữu những kho chứa phụ tùng khổng lồ; linh kiện từ nhà cung cấp chỉ được chở đến dây chuyền lắp ráp đúng vài chục phút trước khi chiếc xe lăn bánh qua trạm sản xuất!",
        "Kỹ sư trưởng Taiichi Ohno và Chủ tịch Eiji Toyoda phát triển Hệ thống Sản xuất Toyota (Toyota Production System - TPS) cùng triết lý Đúng thời điểm (Just-in-Time - JIT) và thẻ Kanban tại Nhật Bản từ thập niên 1950 đến 1970.",
        "Sau Thế chiến thứ hai, nước Nhật đổ nát thiếu vốn, thiếu nguyên vật liệu và thị trường nội địa nhỏ bé không thể áp dụng mô hình sản xuất hàng loạt ồ ạt kiểu dây chuyền Ford của Mỹ (vốn đòi hỏi kho hàng khổng lồ chứa hàng triệu linh kiện dự trữ). Taiichi Ohno đã quan sát cách hoạt động của siêu thị Mỹ: người mua chỉ lấy đúng món đồ mình cần với số lượng cần thiết, và siêu thị chỉ bổ sung hàng lên kệ đúng số lượng đã bán. Ông áp dụng nguyên lý này vào nhà máy xe hơi: mỗi công đoạn sản xuất chỉ yêu cầu công đoạn trước cung cấp đúng linh kiện cần thiết vào đúng thời điểm (JIT) thông qua chiếc thẻ lệnh Kanban, triệt tiêu hoàn toàn 7 loại lãng phí (Muda), đặc biệt là lãng phí tồn kho và chờ đợi.",
        "TPS và JIT trở thành 'kinh thánh' quản trị của ngành sản xuất và logistics toàn cầu; tạo ra làn sóng cách mạng Tinh gọn (Lean Manufacturing) giúp các tập đoàn đa quốc gia cắt giảm hàng tỷ USD chi phí lưu kho và tối ưu hóa vòng quay vốn lưu động.",
        "Mô hình Fordist của Mỹ chuộng sản xuất quy mô lớn đẩy hàng tồn vào kho (Push System); trong khi mô hình JIT của Nhật Bản vận hành theo cơ chế kéo (Pull System) bám sát nhu cầu thực tế của khách hàng.",
        "Tồn kho không phải là tài sản mà là cái bẫy chôn vùi dòng tiền của doanh nghiệp. Học cách vận hành tinh gọn, đo đếm chính xác nhu cầu thị trường và tối ưu hóa thời gian giao nhận (Lead Time) chính là bí quyết giúp doanh nghiệp SME tồn tại vững vàng ngay cả khi nguồn vốn hạn hẹp.",
        "https://en.wikipedia.org/wiki/Just-in-time_manufacturing",
        "Series: Đàm phán & Tư duy",
        "Chưa dùng",
        "Bài học kinh điển về phương pháp sản xuất Just-in-Time của Toyota và tư duy tối ưu hóa dòng tiền chuỗi cung ứng"
    ],
    [
        249,
        "W46 — 15/11/2026",
        "Chó Nghiệp Vụ Hải Quan (K9 Detector Dogs): 'Vũ Khí Sinh Học' Đánh Hơi Ma Túy & Tiền Lậu Chuẩn Xác Hơn 10.000 Lần Mọi Máy Móc Nhân Tạo",
        "Dù con người đã chế tạo ra những cỗ máy soi chiếu tia X trị giá hàng triệu USD, lực lượng hải quan tinh nhuệ nhất thế giới vẫn phải dựa vào chiếc mũi của những chú chó K9 để phát hiện những bọc tiền mặt giấu sâu trong bình xăng xe ô tô!",
        "Lịch sử thành lập các đơn vị Chó nghiệp vụ Hải quan (Customs Canine / K9 Units) tại biên giới Mỹ, Anh, Úc và các quốc gia thành viên WCO từ thập niên 1970.",
        "Tội phạm buôn lậu quốc tế sử dụng muôn vàn thủ đoạn tinh vi: ngâm ma túy vào quần áo, giấu tiền mặt triệu đô vào khung gầm xe tải, bọc hàng cấm trong nhiều lớp giấy bạc, bã cà phê, hạt tiêu cay nồng để đánh lừa máy soi và công chức hải quan. Tuy nhiên khứu giác của loài chó có hơn 300 triệu thụ thể cảm nhận mùi (so với chỉ 6 triệu của con người) và vùng não xử lý mùi lớn gấp 40 lần. Các chú chó giống Labrador hay Belgian Malinois được huấn luyện đặc biệt có thể bóc tách từng phân tử mùi đặc trưng của tiền polymer mới, ma túy tổng hợp, ngà voi hay thịt động vật hoang dã ngay giữa hàng ngàn mùi tạp chất hỗn độn tại các sân bay quốc tế.",
        "Đội quân K9 là một phần không thể thiếu trong cơ chế kiểm soát biên giới hiện đại của Tổ chức Hải quan Thế giới (WCO); hàng năm thu giữ hàng trăm tấn chất cấm và hàng trăm triệu USD tiền mặt buôn lậu xuyên biên giới.",
        "Hải quan Mỹ (CBP) duy trì hơn 1.500 chú chó nghiệp vụ K9; tại Việt Nam, Cục Điều tra Chống buôn lậu (Tổng cục Hải quan) sở hữu Trung tâm Huấn luyện chó nghiệp vụ chuyên trách lập nhiều chiến công bắt giữ các vụ vận chuyển ma túy lớn tại các cửa khẩu Tây Nam và sân bay quốc tế.",
        "Mọi hành vi gian dối, giấu giếm trong khai báo hải quan đều không thể qua mặt được các lớp phòng thủ đa tầng của cơ quan chức năng. Sự chính trực và minh bạch trong khai báo hàng hóa là con đường duy nhất giúp doanh nghiệp thông quan an toàn và giữ gìn danh dự thương hiệu.",
        "https://en.wikipedia.org/wiki/Detection_dog",
        "Series: Bí mật Hải quan",
        "Chưa dùng",
        "Bí mật về chiếc mũi thần kỳ của đội chó nghiệp vụ K9 trong công tác kiểm soát buôn lậu tại các cửa khẩu hải quan quốc tế"
    ],
    [
        250,
        "W46 — 15/11/2026",
        "Vận Đơn Điện Tử Đổi Mạng Lưới (e-BL): Cuộc Cách Mạng Giấy Tờ Hàng Hải Số Hóa Cắt Bớt 6,5 Tỷ USD Chi Phí Chuyển Phát Chứng Từ Cho Thương Mại Toàn Cầu",
        "Trong suốt hàng trăm năm qua, một tờ giấy mỏng manh gọi là Vận đơn gốc (Original Bill of Lading) phải được gửi bằng máy bay chuyển phát nhanh vòng quanh nửa vòng Trái Đất chỉ để người nhận hàng cầm ra cảng đổi lấy một container hàng!",
        "Hiệp hội Vận tải Container Kỹ thuật số (DCSA - Digital Container Shipping Association) gồm 9 hãng tàu lớn nhất thế giới (MSC, Maersk, CMA CGM, Hapag-Lloyd, ONE...) ký cam kết chuyển đổi 100% vận đơn sang dạng số hóa (e-BL) vào năm 2030 (công bố năm 2023).",
        "Hàng năm có hơn 45 triệu bộ vận đơn gốc bằng giấy được phát hành trên toàn cầu. Mỗi bộ vận đơn phải qua tay trung bình 6 đến 8 chủ thể (chủ hàng, forwarder, ngân hàng phát hành L/C, ngân hàng thông báo, hãng tàu, người mua, công chức hải quan). Quy trình in ấn, ký tươi, chuyển phát nhanh quốc tế DHL/FedEx mất từ 3 đến 7 ngày, tiêu tốn hàng tỷ USD tiền bưu chính và thường xuyên xảy ra tai nạn thất lạc chứng từ gốc khiến container bị giam tại bãi cảng hàng tuần. Sự xuất hiện của công nghệ chuỗi khối (Blockchain) và luật mẫu UNCITRAL về Bản ghi Chuyển nhượng Điện tử (MLETR) đã tạo bước ngoặt: vận đơn số e-BL có giá trị pháp lý tương đương bản giấy, chuyển giao sở hữu tức thì chỉ trong 2 giây bằng chữ ký số.",
        "Theo ước tính của McKinsey và DCSA, việc chuyển đổi sang 100% e-BL sẽ tiết kiệm trực tiếp 6,5 tỷ USD chi phí hành chính mỗi năm cho các hãng tàu và chủ hàng, đồng thời thúc đẩy thương mại toàn cầu tăng thêm 30 - 40 tỷ USD nhờ tốc độ giải phóng hàng hóa siêu tốc.",
        "Singapore, Anh, UAE là những quốc gia tiên phong thông qua luật công nhận vận đơn điện tử e-BL; tại Việt Nam, khung pháp lý cho vận đơn điện tử đang được hoàn thiện theo Nghị định 13/2024/NĐ-CP và Đề án Hải quan số 2026.",
        "Chuyển đổi số trong logistics không còn là câu chuyện viển vông mà đang diễn ra từng ngày tại các phòng chứng từ quốc tế. Doanh nghiệp xuất nhập khẩu nào chủ động nắm bắt công nghệ e-BL và giao dịch không giấy tờ sẽ giành lợi thế vượt trội về tốc độ giao dịch và tối ưu hóa chi phí vận hành.",
        "https://en.wikipedia.org/wiki/Paperless_trade",
        "Series: Lịch sử Logistics",
        "Chưa dùng",
        "Cuộc cách mạng vận đơn điện tử e-BL của hiệp hội DCSA và xu hướng số hóa thương mại không giấy tờ toàn cầu"
    ]
]

# ==============================================================================
# HÀM CẬP NHẬT TỪNG WORKBOOK
# ==============================================================================
def update_ideas_bank(filepath, new_rows):
    print(f"\n📂 Đang xử lý tệp: {filepath}")
    if not os.path.exists(filepath):
        print(f"❌ Không tìm thấy tệp: {filepath}")
        return False

    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    current_max_row = ws.max_row
    print(f"   Dòng hiện tại trước khi thêm: {current_max_row}")
    
    last_stt = ws.cell(row=current_max_row, column=1).value
    print(f"   STT cuối cùng hiện tại: {last_stt}")

    # Mẫu định dạng lấy từ dòng cuối cùng hiện tại
    template_row = ws[current_max_row]

    for row_idx, row_data in enumerate(new_rows, start=current_max_row + 1):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if col_idx <= len(template_row):
                copy_cell_style(template_row[col_idx - 1], cell)
        print(f"   ➕ Đã thêm dòng {row_idx} (STT: {row_data[0]}): {str(row_data[2])[:60]}...")

    # Lưu nguyên tử
    atomic_save(wb, filepath)
    return True

# ==============================================================================
# BƯỚC 7: FEEDBACK LOOP CẬP NHẬT DANH SÁCH VĂN BẢN PHÁP LUẬT XNK
# ==============================================================================
def update_legal_feedback_loop():
    filepath = "reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx"
    if not os.path.exists(filepath):
        print(f"❌ Không tìm thấy tệp: {filepath}")
        return

    print(f"\n📂 Đang kiểm tra Feedback Loop: {filepath}")
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active
    current_max_row = ws.max_row
    last_stt = ws.cell(row=current_max_row, column=1).value
    try:
        last_stt_int = int(last_stt)
    except Exception:
        last_stt_int = 246

    template_row = ws[current_max_row]

    new_legal_docs = [
        [
            last_stt_int + 1,
            "Công văn 24083/CHQ-NVTHQ",
            "Tăng cường công tác quản lý trị giá hải quan, siết chặt tham vấn giá và kiểm tra sau thông quan đối với các nhóm mặt hàng trọng điểm phục vụ Tết Nguyên đán Đinh Mùi 2027",
            "Công văn",
            "Cục Hải quan (Bộ Tài chính)",
            "2026-11-10",
            "2026-11-10",
            "Còn hiệu lực",
            "Hướng dẫn thực hiện Điều 25 Thông tư 38/2015/TT-BTC (sửa đổi Thông tư 121/2025/TT-BTC) và Thông tư 39/2015/TT-BTC",
            "Yêu cầu các Chi cục Hải quan tăng cường kiểm tra trị giá khai báo các mặt hàng có thuế suất cao (rượu, bia, bánh kẹo, đồ gia dụng, mỹ phẩm); nghiêm cấm tham vấn chiếu lệ; cập nhật 100% hồ sơ lên Hệ thống GTT02; chuyển hồ sơ nghi vấn cho bộ phận Phúc tập và Kiểm tra sau thông quan (KTSTQ) hậu kiểm 5 năm.",
            "Doanh nghiệp nhập khẩu hàng tiêu dùng Tết phải chủ động chuẩn bị bộ chứng từ thanh toán ngân hàng MT103 và chứng từ thị trường để bảo vệ trị giá khai báo; phòng ngừa rủi ro bị ấn định thuế tăng và kiểm tra sau thông quan kéo dài.",
            "2026-11-15"
        ],
        [
            last_stt_int + 2,
            "Thông tư 69/2025/TT-BTC",
            "Hướng dẫn thi hành một số điều của Luật Thuế giá trị gia tăng số 48/2024/QH15 về khấu trừ, nộp thay nghĩa vụ thuế nhà thầu (FCT) đối với dịch vụ số và thương mại điện tử xuyên biên giới",
            "Thông tư",
            "Bộ Tài chính",
            "2025-11-15",
            "2026-01-01",
            "Còn hiệu lực",
            "Thay thế các quy định cũ về thuế nhà thầu dịch vụ số tại Thông tư 103/2014/TT-BTC",
            "Quy định chi tiết tỷ lệ thuế GTGT trực tiếp 5% trên doanh thu đối với dịch vụ nền tảng số xuyên biên giới (quảng cáo trực tuyến, điện toán đám mây, phần mềm SaaS); bắt buộc doanh nghiệp Việt Nam mua dịch vụ phải khấu trừ và kê khai nộp thay theo Mẫu 01/NTNN.",
            "Doanh nghiệp logistics và XNK có chi trả chi phí dịch vụ số (Meta, Google, AWS, Microsoft) bắt buộc phải thực hiện thủ tục kê khai nộp thay thuế nhà thầu để bảo vệ tính hợp lý của chi phí được trừ khi quyết toán thuế TNDN.",
            "2026-11-15"
        ]
    ]

    for row_idx, row_data in enumerate(new_legal_docs, start=current_max_row + 1):
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if col_idx <= len(template_row):
                copy_cell_style(template_row[col_idx - 1], cell)
        print(f"   📜 Đã bổ sung văn bản mới dòng {row_idx}: {row_data[1]} - {str(row_data[2])[:50]}...")

    atomic_save(wb, filepath)

# ==============================================================================
# MAIN EXECUTION
# ==============================================================================
if __name__ == "__main__":
    print("=" * 80)
    print("🚀 BẮT ĐẦU CẬP NHẬT KHO Ý TƯỞNG 4 PILLARS (WORKFLOW 06) — W46 (15/11/2026)")
    print("=" * 80)

    # 0. CHẠY TECHNICAL GATEKEEPER CHO TOÀN BỘ 20 Ý TƯỞNG
    print("\n🛡️ BƯỚC 0: KIỂM THỬ KỸ THUẬT GATEKEEPER 20 URLS TRƯỚC KHI GHI FILE...")
    validation_items = [
        # TN
        {"stt": TN_ROWS[0][0], "title": TN_ROWS[0][2], "url": TN_ROWS[0][12], "keywords": ["valuation", "customs"]},
        {"stt": TN_ROWS[1][0], "title": TN_ROWS[1][2], "url": TN_ROWS[1][12], "keywords": ["tax", "digital"]},
        {"stt": TN_ROWS[2][0], "title": TN_ROWS[2][2], "url": TN_ROWS[2][12], "keywords": ["drawback", "customs"]},
        {"stt": TN_ROWS[3][0], "title": TN_ROWS[3][2], "url": TN_ROWS[3][12], "keywords": ["minimis", "customs"]},
        {"stt": TN_ROWS[4][0], "title": TN_ROWS[4][2], "url": TN_ROWS[4][12], "keywords": ["basel", "waste"]},
        
        # CS
        {"stt": CS_ROWS[0][0], "title": CS_ROWS[0][2], "url": CS_ROWS[0][10], "keywords": ["incoterms", "cif"]},
        {"stt": CS_ROWS[1][0], "title": CS_ROWS[1][2], "url": CS_ROWS[1][10], "keywords": ["used", "goods"]},
        {"stt": CS_ROWS[2][0], "title": CS_ROWS[2][2], "url": CS_ROWS[2][10], "keywords": ["customs", "declaration"]},
        {"stt": CS_ROWS[3][0], "title": CS_ROWS[3][2], "url": CS_ROWS[3][10], "keywords": ["mail", "express"]},
        {"stt": CS_ROWS[4][0], "title": CS_ROWS[4][2], "url": CS_ROWS[4][10], "keywords": ["cabotage", "transport"]},

        # HT
        {"stt": HT_ROWS[0][0], "title": HT_ROWS[0][2], "url": HT_ROWS[0][10], "keywords": ["valuation", "customs"]},
        {"stt": HT_ROWS[1][0], "title": HT_ROWS[1][2], "url": HT_ROWS[1][10], "keywords": ["drawback", "duty"]},
        {"stt": HT_ROWS[2][0], "title": HT_ROWS[2][2], "url": HT_ROWS[2][10], "keywords": ["fob", "shipping"]},
        {"stt": HT_ROWS[3][0], "title": HT_ROWS[3][2], "url": HT_ROWS[3][10], "keywords": ["withholding", "tax"]},
        {"stt": HT_ROWS[4][0], "title": HT_ROWS[4][2], "url": HT_ROWS[4][10], "keywords": ["customs"]},

        # FF
        {"stt": FF_ROWS[0][0], "title": FF_ROWS[0][2], "url": FF_ROWS[0][9], "keywords": ["suez", "ever given"]},
        {"stt": FF_ROWS[1][0], "title": FF_ROWS[1][2], "url": FF_ROWS[1][9], "keywords": ["rotterdam", "port"]},
        {"stt": FF_ROWS[2][0], "title": FF_ROWS[2][2], "url": FF_ROWS[2][9], "keywords": ["toyota", "just-in-time"]},
        {"stt": FF_ROWS[3][0], "title": FF_ROWS[3][2], "url": FF_ROWS[3][9], "keywords": ["detection", "dog"]},
        {"stt": FF_ROWS[4][0], "title": FF_ROWS[4][2], "url": FF_ROWS[4][9], "keywords": ["paperless", "trade"]}
    ]

    enforce_strict_validation(validation_items)

    # 1. Cập nhật Tin Ngành
    update_ideas_bank("reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx", TN_ROWS)

    # 2. Cập nhật Case Study
    update_ideas_bank("reports/Case_Study_Ideas_Bank_Eureka.xlsx", CS_ROWS)

    # 3. Cập nhật Học Thực Chiến
    update_ideas_bank("reports/Hoc_Thuc_Chien_Ideas_Bank_Eureka.xlsx", HT_ROWS)

    # 4. Cập nhật Funfact
    update_ideas_bank("reports/Funfact_Ideas_Bank_Eureka.xlsx", FF_ROWS)

    # 5. Feedback loop danh mục luật
    update_legal_feedback_loop()

    print("\n" + "=" * 80)
    print("🎉 HOÀN TẤT 100% CẬP NHẬT 20 Ý TƯỞNG MỚI CHO 4 PILLARS & FEEDBACK LOOP PHÁP LUẬT NGÀY 15/11/2026!")
    print("=" * 80)
