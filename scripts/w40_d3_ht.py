# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W40 DAY 3 — 30/09/2026)
Số lượng: 5 ý tưởng (STT 116 - 120)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, quy trình step-by-step, công thức/link, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W40_D3 = [
    {
        "stt": 116,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Cẩm Nang 4 Bước Soát Xét C/O Form AI Hóa Đơn Bên Thứ Ba (Third-Party Invoicing) Chuẩn Tuyệt Đối Để Tránh Bị Bác Bỏ Ưu Đãi Thuế",
        "workflow": "Bước 1: Kiểm tra ô số 7 & ô số 13: Đảm bảo ô số 7 ghi đầy đủ tên và địa chỉ công ty phát hành hóa đơn bên thứ ba (trader), ô số 13 bắt buộc phải được đánh dấu tích (tick) vào 'Third-party invoicing'; Bước 2: Kiểm tra ô số 10: Số và ngày hóa đơn bắt buộc phải là số hóa đơn thương mại của bên thứ ba (hóa đơn dùng để khai báo hải quan Việt Nam), không ghi chồng chéo hóa đơn nội bộ của nhà máy; Bước 3: So khớp ô số 9: Trị giá FOB ghi trên C/O phải tương thích với quy định của hiệp định AIFTA; Bước 4: Đối chiếu với Cổng thông tin eCOSys: Truy cập hệ thống tra cứu C/O của cơ quan cấp nước xuất khẩu để xác nhận dữ liệu điện tử trùng khớp hoàn toàn trước khi mở tờ khai.",
        "formula": "Thuế nhập khẩu tiết kiệm được nhờ C/O Form AI = Trị giá tính thuế x (Thuế suất MFN - Thuế suất ưu đãi đặc biệt AIFTA); Ví dụ lô hàng 100.000 USD, MFN 10%, AIFTA 0% -> Tiền thuế tiết kiệm = 100.000 x (10% - 0%) x 25.500 = 255.000.000 VNĐ.",
        "top3_mistakes": "1. Để ô số 10 ghi số hóa đơn của nhà sản xuất trong khi nộp cho hải quan hóa đơn của công ty thương mại Singapore; 2. Quên tích ô số 13 'Third-party invoicing'; 3. Ghi sai tên công ty bên thứ ba tại ô số 7 so với hợp đồng mua bán.",
        "mnemonic_tip": "C/O 3 bên như chiếc vé máy bay mua qua đại lý: Tên hãng bay là nhà sản xuất, tên đại lý xuất vé là bên thứ ba, số vé trên tay bạn (hóa đơn thương mại) bắt buộc phải khớp với số hiển thị trên cổng check-in (C/O ô số 10).",
        "comparison": "C/O 2 bên thông thường vs C/O 3 bên (Third-Party Invoicing): C/O 2 bên bên bán và nhà sản xuất là một; C/O 3 bên đòi hỏi kiểm soát chặt chẽ tính độc lập của hóa đơn thương mại và bắt buộc tích ô số 13 để hải quan công nhận luồng tiền hợp lệ.",
        "persona": "Anh Minh (Chủ doanh nghiệp XNK nguyên liệu) & Kế toán xuất nhập khẩu, nhân viên chứng từ Forwarder.",
        "checklist": "☐ Ô 7: Tên & địa chỉ công ty bên thứ ba\n☐ Ô 10: Đúng số & ngày hóa đơn thương mại bên thứ ba\n☐ Ô 13: Đã tích 'Third-party invoicing'\n☐ Dữ liệu tra cứu trên eCOSys trùng khớp 100%\n☐ Hợp đồng thể hiện rõ điều khoản mua bán 3 bên.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 15/2010/TT-BCT & 08/2024/TT-BCT) & daily_notes.md (case C/O AI 3 bên ngày 30/09/2026).",
        "keywords": ["certificate", "origin"]
    },
    {
        "stt": 117,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Quy Trình 4 Bước Xác Định Xuất Xứ Hàng Hóa Gia Công Biến Tính & Hướng Dẫn Kê Khai Đúng Khi RVC Không Đạt Tiêu Chuẩn",
        "workflow": "Bước 1: So sánh mã HS đầu vào vs mã HS thành phẩm sau gia công: Kiểm tra ở cấp 2 chữ số (CC), 4 chữ số (CTH) hoặc 6 chữ số (CTSH) theo quy tắc cụ thể mặt hàng (PSR); Bước 2: Tính toán hàm lượng giá trị gia tăng (RVC): Tính theo phương pháp gián tiếp hoặc trực tiếp căn cứ chi phí nguyên liệu nội địa, chi phí nhân công và chi phí quản lý; Bước 3: Đánh giá tiêu chí công đoạn gia công đơn giản: Rà soát Điều 9 Nghị định 31/2018/NĐ-CP để xác định công đoạn pha trộn, thay đổi độ cứng có bị coi là gia công giản đơn hay không; Bước 4: Hướng dẫn khai báo xuất xứ và ghi nhãn: Nếu không đổi mã HS và RVC < 30%, trên nhãn hàng hóa bắt buộc ghi xuất xứ nguyên liệu ban đầu (Made in China) hoặc 'Xuất xứ: Không xác định', tuyệt đối không khai Made in Vietnam.",
        "formula": "RVC (Hàm lượng giá trị khu vực) = [(FOB - VNM) / FOB] x 100% hoặc RVC = (VOM / FOB) x 100%; Trong đó: FOB là giá xuất xưởng/FOB của thành phẩm, VNM là trị giá nguyên liệu không có xuất xứ (nhập khẩu), VOM là trị giá nguyên liệu và chi phí nội địa có xuất xứ Việt Nam.",
        "top3_mistakes": "1. Nghĩ rằng bỏ chi phí gia công biến tính tại Việt Nam là tự động được cấp xuất xứ Made in Vietnam; 2. Tự ý thay đổi mã HS thành phẩm khi cơ quan hải quan giám định chỉ ra bản chất hóa học không đổi; 3. Khai báo mâu thuẫn giữa hồ sơ hải quan và nhãn phụ dán trên bao bì sản phẩm.",
        "mnemonic_tip": "Xuất xứ như quốc tịch con người: Bố mẹ là người nước ngoài sinh con ở nước ngoài, đưa về Việt Nam học tập và rèn luyện thì người đó vẫn mang quốc tịch gốc trừ khi hoàn tất thủ tục nhập tịch chính thức (vượt qua bài kiểm tra CTC/RVC).",
        "comparison": "Tiêu chí Chuyển đổi mã số hàng hóa (CTC) vs Tiêu chí Hàm lượng giá trị (RVC): CTC dựa trên sự biến đổi bản chất vật lý/hóa học làm thay đổi chương/nhóm/phân nhóm HS; RVC dựa trên tỷ lệ đóng góp chi phí tài chính trong nước. Khi CTC không đạt, RVC là chiếc phao cứu sinh duy nhất.",
        "persona": "Chị Hương (Giám đốc công ty sản xuất gia công) & Trưởng phòng Xuất nhập khẩu các nhà máy DNCX/FDI.",
        "checklist": "☐ Tra cứu tiêu chí PSR mặt hàng tại Thông tư 05/2018/TT-BCT\n☐ Lập bảng tính toán định mức chi phí RVC có xác nhận kế toán\n☐ Kiểm tra danh mục các công đoạn gia công đơn giản\n☐ Thống nhất ghi nhãn mác xuất xứ trong hợp đồng gia công\n☐ Lưu trữ trọn bộ hồ sơ chứng minh nguồn gốc nguyên liệu.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP & Thông tư 05/2018/TT-BCT) & daily_notes.md (hạt nhựa gia công biến tính ngày 30/09/2026).",
        "keywords": ["rules", "origin"]
    },
    {
        "stt": 118,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Quy Trình 3 Bước Xử Lý Khi Container Bị Đổi / Gãy Seal Trong Quá Trình Vận Chuyển Chuyển Cảng (Port Transfer Seal Protocol)",
        "workflow": "Bước 1: Lập Biên bản ghi nhận hiện trường sự cố gãy seal: Ngay khi phát hiện seal bị gãy trong quá trình bốc xếp/vận chuyển, yêu cầu đại diện bến cảng hoặc tài xế lập biên bản có chữ ký của bộ phận điều độ cảng và chụp ảnh hiện trạng vị trí tay khóa cont; Bước 2: Bấm seal mới và nhận Chứng từ thay seal: Bộ phận giám sát cảng tiến hành bấm seal mới (seal cảng) và cấp Giấy xác nhận thay seal ghi rõ: số cont, số seal cũ bị gãy, số seal mới bấm, lý do bấm lại; Bước 3: Cập nhật tờ khai hải quan & Giải trình trên hệ thống VASSCM: Khi truyền danh sách cont lấy hàng tại cảng đến, truyền khai báo theo SỐ SEAL MỚI và tải bản scan Biên bản sự cố lên hệ thống Một cửa quốc gia trước khi xe tiến vào cổng cảng.",
        "formula": "Chi phí rủi ro phát sinh nếu không xử lý đúng quy trình = Phí nâng hạ kiểm hóa (1.500.000 - 3.000.000 VNĐ/cont) + Phí lưu bãi bến (500.000 - 1.200.000 VNĐ/ngày x số ngày lưu) + Phí phạt vi phạm quy chế giám sát hải quan (2.000.000 - 5.000.000 VNĐ).",
        "top3_mistakes": "1. Để tài xế tự ý mua seal chợ đen bấm bừa mà không có biên bản của Cảng; 2. Khai số seal cũ của hãng tàu trong khi thực tế trên cont là seal mới của cảng; 3. Không chụp ảnh hiện trường lúc seal bị đứt gãy để làm bằng chứng giải trình với hải quan giám sát.",
        "mnemonic_tip": "Seal container như niêm phong bài thi đại học: Muốn bóc niêm phong hay dán lại bắt buộc phải có biên bản của Hội đồng coi thi, tự ý xé dán là bị đình chỉ thi ngay lập tức!",
        "comparison": "Seal Hãng tàu (Carrier Seal) vs Seal Niêm phong Cảng (Terminal Seal): Seal hãng tàu xác nhận trách nhiệm từ cảng xếp hàng; Seal cảng dùng thay thế tạm thời khi có sự cố nội địa và chỉ có giá trị pháp lý khi đi kèm Biên bản ghi nhận sự cố có dấu cảng.",
        "persona": "Anh Dũng (Chủ doanh nghiệp vận tải bộ/sà lan) & Đội ngũ điều phối hiện trường (Ops), tài xế xe đầu kéo container.",
        "checklist": "☐ Kiểm tra tình trạng nguyên vẹn của seal trước khi kéo cont ra khỏi bãi\n☐ Yêu cầu Cảng lập Biên bản sự cố ngay khi seal bị va đập gãy\n☐ Lưu giữ số seal mới và ảnh chụp rõ nét 4 góc khóa cont\n☐ Cập nhật số seal mới trên phần mềm khai báo hải quan ECUS5\n☐ Đính kèm biên bản sự cố khi làm thủ tục qua cổng cảng (in-gate/out-gate).",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 38/2015 & Thông tư 39/2018/TT-BTC) & daily_notes.md (sự cố gãy seal chuyển cảng Cát Lái - Hiệp Phước ngày 25/09/2026).",
        "keywords": ["security", "seal"]
    },
    {
        "stt": 119,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Bộ Checklist 5 Điểm Tự Bảo Vệ Trước Rủi Ro Mua Hàng Từ Doanh Nghiệp Bỏ Trốn & Hóa Đơn Công Ty Ma (Anti-Ghost Company Protocol)",
        "workflow": "Bước 1: Thẩm định pháp lý thời gian thực: Tra cứu trạng thái mã số thuế đối tác trên Cổng thông tin Tổng cục Thuế (tracuunnt.gdt.gov.vn) và Cổng đăng ký kinh doanh quốc gia trước khi ký hợp đồng và trước mỗi đợt chuyển tiền; Bước 2: Kiểm tra thực địa cơ sở vật chất (KYC Site Visit): Yêu cầu ảnh chụp nhà xưởng, văn phòng, kho hàng hoặc cử nhân sự thị sát trực tiếp đối với các hợp đồng giá trị lớn (> 100 triệu); Bước 3: Chuẩn hóa phương thức thanh toán: 100% thanh toán qua tài khoản ngân hàng đã đăng ký với cơ quan thuế của đúng pháp nhân trên hóa đơn, tuyệt đối không chuyển tiền vào tài khoản cá nhân giám đốc; Bước 4: Thiết lập 'Hồ sơ chứng minh logistics thực tế': Bắt buộc lưu trữ trọn bộ biên bản giao nhận hàng, phiếu cân, hợp đồng vận chuyển, biển số xe tải và số điện thoại tài xế giao hàng; Bước 5: Bổ sung điều khoản bảo lãnh thuế trong hợp đồng: Quy định bên bán chịu phạt 100% số thuế và chi phí phát sinh nếu hóa đơn bị cơ quan thuế kết luận là rủi ro hoặc bên bán bỏ trốn trong vòng 3 năm.",
        "formula": "Tổng thiệt hại tài chính khi dính hóa đơn công ty ma = Tiền thuế VAT bị truy thu + Tiền thuế TNDN bị truy thu (20% giá trị trước thuế) + Tiền phạt khai sai 20% + Tiền chậm nộp (0.03%/ngày x Số ngày từ ngày nộp tờ khai đến ngày thanh tra).",
        "top3_mistakes": "1. Ỷ lại vào việc 'tiền đã chuyển khoản qua ngân hàng và hóa đơn có mã của cơ quan thuế' là an toàn tuyệt đối; 2. Mua hàng giao tận nơi nhưng không lập biên bản giao nhận hoặc không ghi biển số xe tải vận chuyển; 3. Không kiểm tra định kỳ tình trạng mã số thuế của đối tác cũ sau khi đã hoàn tất hợp đồng.",
        "mnemonic_tip": "Mua hàng thời hóa đơn điện tử giống như mua đất: Không chỉ xem sổ đỏ (hóa đơn), mà bắt buộc phải đến tận nơi xem đất thực tế (kho hàng) và có biên bản cắm mốc ranh giới (chứng từ logistics vận tải).",
        "comparison": "Chứng từ thanh toán ngân hàng vs Hồ sơ chứng minh logistics thực tế: Ngân hàng chỉ chứng minh dòng tiền di chuyển; Hồ sơ logistics (phiếu cân, xe vận tải, kho bãi) chứng minh hàng hóa có thực - yếu tố quyết định để cơ quan thuế không bác chi phí.",
        "persona": "Anh Minh (Chủ doanh nghiệp thương mại/sản xuất) & Kế toán trưởng, Giám đốc tài chính (CFO), Trưởng phòng Thu mua.",
        "checklist": "☐ Tra cứu trạng thái MST bên bán còn 'Đang hoạt động'\n☐ Hợp đồng kinh tế có điều khoản cam kết tính hợp pháp hóa đơn\n☐ Ủy nhiệm chi ngân hàng khớp 100% thông tin tài khoản bên bán\n☐ Trọn bộ biên bản giao nhận kho có chữ ký thủ kho và tài xế\n☐ Lưu trữ phiếu cân xe hoặc vận đơn vận chuyển hàng thực tế.",
        "url": "https://en.wikipedia.org/wiki/Tax_evasion",
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Luật Quản lý thuế số 38/2019/QH14 & Nghị định 126/2020/NĐ-CP) & daily_notes.md (rủi ro hóa đơn công ty ma ngày 25/09/2026).",
        "keywords": ["tax", "evasion"]
    },
    {
        "stt": 120,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Hướng Dẫn 4 Nguyên Tắc Quản Trị Đơn Hàng Chính Ngạch Bằng Mã Khách Hàng (Customer Code SOP) Giúp Tối Ưu Tốc Độ Kho Bãi & Thông Quan",
        "workflow": "Bước 1: Khởi tạo Mã Khách Hàng định danh duy nhất (Customer Unique ID): Thiết lập trên hệ thống ERP cú pháp chuẩn gồm mã phân loại + tên viết tắt + mã số thuế (Ví dụ: ERK-HN-001); Bước 2: Bắt buộc dán Shipping Mark chứa Mã Khách Hàng lên 100% kiện hàng: In tem nhãn chuẩn dán lên tối thiểu 2 mặt của từng thùng carton trước khi chuyển về kho ngoại quan/kho quốc tế; Bước 3: Bắn mã vạch quét Barcode tự động khi nhập kho: Nhân viên kho chỉ tiếp nhận kiện hàng có mã định danh, hệ thống tự động đồng bộ trọng lượng, kích thước lên tài khoản khách hàng trong 15 phút; Bước 4: Khớp nối tự động với Dữ liệu Khai báo Hải quan: Khi hàng đủ điều kiện xuất khẩu, hệ thống ERP tự động gom nhóm hàng theo mã khách hàng để xuất dữ liệu làm Invoice, Packing List và tờ khai hải quan chính ngạch mà không sợ thất lạc.",
        "formula": "Thời gian tiết kiệm được nhờ quản trị ERP mã khách hàng = Thời gian kiểm đếm thủ công (3-5 ngày) - Thời gian quét barcode tự động (15-30 phút); Tỷ lệ thất lạc hàng hóa giảm từ 3-5% xuống 0%.",
        "top3_mistakes": "1. Khách gửi hàng thẳng từ nhà máy Trung Quốc về kho mà không dán tem mã khách hàng, khiến hàng thành 'hàng vô thừa nhận'; 2. Gửi hàng xong không gửi mã vận đơn nội địa cho bộ phận CSKH để theo dõi; 3. Không khai báo chi tiết danh mục hàng hóa trước khi hàng về kho ngoại quan.",
        "mnemonic_tip": "Mã khách hàng như số hộ chiếu đi máy bay: Không có hộ chiếu thì dù bạn có đứng trước cửa máy bay cũng không nhân viên nào cho bạn lên; không có mã khách hàng thì hàng triệu kiện hàng trong kho không thể bay về Việt Nam!",
        "comparison": "Quản trị kho thủ công (Manual Tracking) vs Quản trị ERP bằng Mã định danh (Barcode ERP): Quản lý thủ công phụ thuộc vào trí nhớ cá nhân sale, dễ thất lạc và nhầm lẫn hàng khi cao điểm; Quản trị ERP số hóa toàn diện, truy xuất lịch sử trong 5 giây, bảo vệ uy tín doanh nghiệp trước mọi khiếu nại.",
        "persona": "Chị Thúy (CEO Eureka) & Trưởng phòng Vận hành Logistics, Nhân viên Sales, Khách hàng nhập khẩu chính ngạch.",
        "checklist": "☐ Cấp mã khách hàng định danh ngay khi ký hợp đồng dịch vụ\n☐ Gửi hướng dẫn dán Shipping Mark bằng 2 thứ tiếng (Việt - Trung/Anh)\n☐ Yêu cầu khách nhập mã vận đơn nội địa lên cổng portal trước khi hàng đến\n☐ Quét mã vạch kiểm đếm tự động ngay tại cửa kho\n☐ Tự động cảnh báo các kiện hàng thiếu mã sau 24h lưu kho.",
        "url": "https://en.wikipedia.org/wiki/Enterprise_resource_planning",
        "series": "Series: Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (quy chuẩn quản lý đơn hàng ERP Eureka ngày 24/09/2026) & Brand Persona Eureka.",
        "keywords": ["enterprise", "resource"]
    }
]
