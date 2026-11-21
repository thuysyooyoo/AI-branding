# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W41 DAY 2 — 06/10/2026)
Số lượng: 5 ý tưởng (STT 126 - 130)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W41_D2 = [
    {
        "stt": 126,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Hướng Dẫn 4 Bước Lựa Chọn Chuẩn Điều Kiện Incoterms 2020 Trong Xuất Nhập Khẩu Tại Chỗ (FCA vs EXW vs DAP) Để Không Bị Ấn Định Thuế",
        "workflow": "Bước 1: Rà soát Hợp đồng Ngoại thương 3 bên để xác định chủ thể thanh toán cước vận tải nội địa (Bên bán nước ngoài, Bên bán VN hay Bên mua VN); Bước 2: So sánh Điều kiện Incoterms phù hợp: Ưu tiên chọn DAP tại kho người nhận nếu bên bán chịu cước và muốn gộp cước vào giá bán; chọn FCA nếu bên mua cử xe đến lấy tại kho bên bán; Bước 3: Phân bổ và Khai báo Chi phí Vận chuyển trên Tờ khai Hải quan: Nếu chọn EXW/FCA, bên mua bắt buộc phải khai báo khoản điều chỉnh cộng chi phí vận tải nội địa vào ô Trị giá tính thuế; Bước 4: Chuẩn hóa Bộ chứng từ Kế toán: Lưu trữ hợp đồng vận chuyển, hóa đơn cước và phiếu xuất kho để sẵn sàng đối soát khi hải quan kiểm tra sau thông quan.",
        "formula": "Trị giá tính thuế NK tại chỗ = Giá hàng hóa ghi trên hóa đơn thương mại + Chi phí vận chuyển và bốc dỡ nội địa đến kho người nhận chỉ định (nếu chưa bao gồm trong giá mua).",
        "top3_mistakes": "1. Chọn điều kiện EXW/FCA nhưng không khai báo chi phí vận chuyển nội địa vào trị giá tính thuế hải quan; 2. Tự ý thỏa thuận miệng chi phí cước xe trả hộ mà không thể hiện trên hợp đồng mua bán; 3. Khai báo điều kiện DDP nhưng bên bán trong nước không có chức năng nộp thay thuế nhập khẩu cho bên mua nội địa.",
        "mnemonic_tip": "Ví von Incoterms trong XKTC như 'Chiếc vé xe buýt': Ai là người trả tiền vé thì chi phí đó vẫn phải tính vào tổng giá trị chuyến đi; trốn tiền vé thì thanh tra hải quan sẽ phạt gấp đôi.",
        "comparison": "So sánh FCA vs DAP trong XKTC: FCA thì người mua tự điều xe và phải tự cộng cước vào tờ khai thuế; DAP thì người bán lo trọn gói vận tải và xuất hóa đơn gộp, thủ tục thông quan người mua đơn giản hơn 50%.",
        "persona": "Anh Minh (Chủ doanh nghiệp phụ trợ) & Chị Hương (Kế toán trưởng công ty XNK)",
        "checklist": "[ ] Hợp đồng ghi rõ địa điểm giao hàng DAP/FCA; [ ] Thể hiện rõ ai thanh toán cước nội địa; [ ] Đã cộng chi phí vận tải vào tờ khai nếu đi EXW/FCA; [ ] Lưu trữ hóa đơn cước xe khớp với tờ khai.",
        "url": "https://en.wikipedia.org/wiki/Incoterms",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 60/2019/TT-BTC) & daily_notes.md (ghi nhận ngày 06/10/2026 từ case study Stella Pham).",
        "keywords": ["incoterms", "fca", "exw"]
    },
    {
        "stt": 127,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Cẩm Nang 4 Bước Tự Soát C/O Form AI 3 Bên (Third-Party Invoicing) Chuẩn Chỉnh Tránh Bị Tạm Thu Thuế Hải Quan",
        "workflow": "Bước 1: Kiểm tra Tính Hợp Pháp của Hóa Đơn Thương Mại Bên Thứ Ba: Xác định hóa đơn do công ty thương mại phát hành (Third party invoice) và lấy số hiệu, ngày phát hành; Bước 2: Đối soát Ô số 9 (Trị giá FOB) và Ô số 10 (Số Invoice): Yêu cầu cơ quan cấp chỉ thể hiện 1 số invoice thương mại thực tế tại ô số 10; nếu bắt buộc ghi cả 2 số thì phải ghi chú rõ ràng 'Invoice của bên bán thứ ba' và 'Invoice của nhà sản xuất'; Bước 3: Kiểm tra Ô số 13: Bắt buộc phải được tích chọn dấu 'Third-Party Invoicing' và tên, quốc gia của công ty thương mại phải được ghi rõ ràng tại ô số 7; Bước 4: Tra cứu trên Cổng Thông Tin C/O Quốc Gia của Nước Cấp: Nhập mã số tham chiếu C/O để kiểm tra tính toàn vẹn của dữ liệu điện tử trước khi truyền tờ khai Một cửa.",
        "formula": "Tỷ lệ tiết kiệm thuế nhờ C/O Form AI = (Thuế suất MFN thông thường - Thuế suất ưu đãi đặc biệt AIFTA 0%) x Trị giá tính thuế CIF của lô hàng.",
        "top3_mistakes": "1. Ô số 10 ghi số hóa đơn của nhà sản xuất nhưng tờ khai hải quan lại đính kèm hóa đơn của công ty thương mại; 2. Quên không tích chọn ô số 13 'Third-Party Invoicing'; 3. Trị giá FOB ô số 9 ghi theo giá nhà sản xuất nhưng hợp đồng ngoại thương thanh toán theo giá công ty thương mại mà không có giải trình chênh lệch.",
        "mnemonic_tip": "Nhớ quy tắc '3 Trùng' của C/O bên thứ ba: Số hóa đơn ô 10 trùng với Hóa đơn mở tờ khai - Ô 13 trùng với Tên công ty phát hành hóa đơn - Hàng hóa ô 7 trùng với Thực tế nhập khẩu.",
        "comparison": "C/O 2 bên vs C/O 3 bên (Third-Party): C/O 2 bên thì nhà sản xuất bán trực tiếp cho người mua Việt Nam; C/O 3 bên có sự tham gia của trung gian thương mại nước ngoài, thủ tục soát xét ô số 7, 10, 13 khắt khe gấp 3 lần.",
        "persona": "Chị Hương (Trưởng phòng XNK) & Anh Dũng (Chuyên viên chứng từ Forwarder)",
        "checklist": "[ ] Ô số 10 ghi đúng số hóa đơn thương mại kèm tờ khai; [ ] Đã tích ô số 13 Third-Party Invoicing; [ ] Tên công ty bên thứ ba thể hiện tại ô số 7; [ ] Tra cứu mã tham chiếu C/O thành công trên cổng điện tử.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 05/2018/TT-BCT) & daily_notes.md (ghi nhận ngày 30/09/2026 từ case study Bùi Văn Khẩn).",
        "keywords": ["origin", "goods"]
    },
    {
        "stt": 128,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Quy Trình 4 Bước Đánh Giá Tính Hợp Lệ Của Xuất Xứ Hàng Hóa Gia Công (Tiêu Chí CTH & RVC) Theo Nghị Định 31/2018/NĐ-CP",
        "workflow": "Bước 1: Tra cứu Bảng Tiêu chí Xuất xứ cụ thể cho Mặt hàng (PSR - Product Specific Rules): Xác định mặt hàng thuộc diện áp dụng Chuyển đổi mã số (CTC/CTH/CTSH) hay Tỷ lệ hàm lượng giá trị gia tăng (RVC); Bước 2: Phân tích Quy trình Công nghệ & Mã HS: So sánh mã HS của nguyên liệu đầu vào với mã HS của thành phẩm đầu ra; loại trừ ngay nếu quy trình chỉ là gia công đơn giản theo Điều 9 Nghị định 31/2018; Bước 3: Tính toán Hàm lượng Giá trị Gia tăng RVC: Áp dụng công thức tính trực tiếp hoặc gián tiếp để đo lường tỷ lệ % giá trị tạo ra tại Việt Nam; Bước 4: Kết luận Xuất xứ & Thiết kế Nhãn Hàng Hóa: Nếu đạt RVC >= 40% hoặc chuyển đổi mã HS hợp lệ -> Ghi 'Made in Vietnam'; Nếu không đạt -> Giữ nguyên xuất xứ ban đầu hoặc ghi 'Sản xuất tại VN từ nguyên liệu nhập khẩu'.",
        "formula": "Công thức RVC gián tiếp = [(Giá FOB thành phẩm - Trị giá nguyên liệu không có xuất xứ VNM) / Giá FOB thành phẩm] x 100%. Tiêu chuẩn chuẩn thường yêu cầu RVC >= 40%.",
        "top3_mistakes": "1. Tự ý ghi 'Made in Vietnam' khi sản phẩm chỉ qua công đoạn pha trộn, đóng gói đơn giản; 2. Đưa chi phí khấu hao máy móc khống vào công thức tính RVC để làm tròn lên 40%; 3. Không lưu trữ bảng định mức sản xuất và hồ sơ chứng minh xuất xứ nguyên liệu đầu vào suốt 5 năm.",
        "mnemonic_tip": "Ví von quy tắc xuất xứ như 'Nấu bát phở': Mua bánh phở và thịt bò nấu chín thì là Phở Việt Nam; nhưng mua lon nước ngọt về dán tem tiếng Việt thì vĩnh viễn không thể gọi là nước ngọt Việt Nam.",
        "comparison": "Tiêu chí CTH (Chuyển đổi 4 số) vs Tiêu chí RVC (Hàm lượng 40%): CTH dựa trên bản chất kỹ thuật thay đổi mã số hàng hóa; RVC dựa trên giá trị kinh tế gia tăng hạch toán bằng tiền tệ.",
        "persona": "Anh Minh (Chủ xưởng sản xuất phụ trợ) & Anh Dũng (Giám đốc nhà máy FDI)",
        "checklist": "[ ] Kiểm tra mã HS nguyên liệu vs thành phẩm; [ ] Loại trừ 10 công đoạn gia công đơn giản Điều 9; [ ] Tính toán công thức RVC đạt >= 40%; [ ] Thiết kế nhãn mác đúng quy định Nghị định 111/2021.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP, Điều 9) & daily_notes.md (ghi nhận ngày 30/09/2026 từ case study Minh Thu).",
        "keywords": ["origin", "customs"]
    },
    {
        "stt": 129,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Quy Trình 3 Bước Xử Lý Sự Cố Thay Seal Container Khi Chuyển Cảng Đích Tránh Bị Tắc Cổng Giám Sát Tự Động (VASSCM)",
        "workflow": "Bước 1: Tiếp nhận Thông tin và Yêu cầu Biên bản Ghi nhận Sự cố: Ngay khi nhận được thông báo đứt/gãy seal tại cảng chuyển, yêu cầu Cảng và Hãng tàu lập Biên bản bàn giao/thay seal có ký nhận của Hải quan giám sát bãi; Bước 2: Cập nhật Số Seal Mới vào Phần Mềm Khai Báo Hải Quan: Truy cập phần mềm VNACCS/VCIS, nhập số seal mới của cảng hoặc seal hãng tàu cấp lại vào chỉ tiêu khai báo danh sách container của tờ khai vận chuyển kết hợp (OLA) hoặc tờ khai nhập khẩu; Bước 3: Đính kèm Hồ sơ Giải trình Điện tử lên Hệ thống V5: Tải bản scan Biên bản thay seal lên hệ thống Một cửa quốc gia trước khi xe cont tiếp cận cổng cảng đích để bộ phận giám sát tự động phê duyệt giải phóng xe.",
        "formula": "Chi phí tắc cổng do sai seal = Phí neo xe đầu kéo (2.000.000 - 3.500.000 VNĐ/ngày) + Phí nâng hạ bãi kiểm hóa luồng Đỏ (1.500.000 - 3.000.000 VNĐ/cont) + Phạt vi phạm thủ tục hải quan.",
        "top3_mistakes": "1. Nhắm mắt khai theo số seal cũ trên Bill of Lading mà không kiểm tra tình trạng seal thực tế khi hàng về cảng đích; 2. Không lấy Biên bản sự cố của Cảng trước khi cho xe lăn bánh rời cảng trung chuyển; 3. Đợi xe bị chặn barie cổng cảng mới cuống cuồng liên hệ hãng tàu xin xác nhận.",
        "mnemonic_tip": "Nhớ câu thần chú: 'Seal trên cont là Seal trên web' - Thực tế kẹp chì số mấy thì tờ khai điện tử phải gõ đúng số đó, sai một ký tự là đứng bánh xe.",
        "comparison": "Tờ khai OLA vận chuyển độc lập vs Tờ khai chuyển cửa khẩu thông thường: OLA giám sát chặt chẽ hành trình bằng định vị GPS và seal điện tử/seal chì chuẩn ISO 17712, quy định xử lý sự cố đứt seal ngặt nghèo gấp đôi.",
        "persona": "Anh Dũng (Điều hành đội xe vận tải) & Chị Hương (Chuyên viên chứng từ Hải quan)",
        "checklist": "[ ] Có ảnh chụp thực tế seal mới trên container; [ ] Cầm trong tay Biên bản thay seal của Cảng; [ ] Đã gõ số seal mới lên hệ thống VNACCS; [ ] Hệ thống VASSCM báo Barcode xanh hợp lệ.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Điều 51 Thông tư 38/2015/TT-BTC & Quyết định 2722/QĐ-TCHQ) & daily_notes.md (ghi nhận ngày 25/09/2026 từ case study Bùi Văn Khẩn).",
        "keywords": ["seal", "container"]
    },
    {
        "stt": 130,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Hướng Dẫn 4 Bước Chuẩn Hóa Quản Trị Chuỗi Cung Ứng Bằng Mã Định Danh (Barcoding/ERP) Cho Hàng Ủy Thác Nhập Khẩu Chính Ngạch",
        "workflow": "Bước 1: Khởi tạo Mã Định Danh Khách Hàng (Customer ID) & Mã Đơn Hàng (Order ID) trên Phần mềm ERP: Mỗi khách hàng và lô hàng được cấp một mã vạch/mã QR duy nhất trước khi giao dịch phát sinh; Bước 2: Thiết lập Quy chuẩn Đóng gói & Dán Tem Nội Địa (Shipping Mark): Hướng dẫn nhà cung cấp nước ngoài in tem định danh kích thước chuẩn 10x15cm dán tối thiểu 2 mặt góc của từng thùng hàng; Bước 3: Bắn Mã Quét Tự Động (Barcode Scanning) Tại Điểm Tập Kết Kho Ngoại: Khi xe hàng về kho biên giới, nhân viên dùng máy quét cầm tay đối soát tự động; kiện nào không có mã lập tức bị cách ly vào khu vực chờ xác minh; Bước 4: Đồng bộ Dữ liệu với Hồ sơ Tờ Khai Hải Quan: Số kiện, trọng lượng và mã kiện được tự động đẩy vào packing list điện tử phục vụ mở tờ khai hải quan chính ngạch chỉ trong 5 phút.",
        "formula": "Tỷ lệ giảm thiểu rủi ro thất lạc hàng hóa = 100% - (Số kiện mất mã / Tổng số kiện hàng vận chuyển) x 100%. Với hệ thống ERP chuẩn, tỷ lệ sai sót giảm từ 8% xuống dưới 0.01%.",
        "top3_mistakes": "1. Cho phép nhà cung cấp nước ngoài gửi hàng về kho mà không dán mã định danh khách hàng; 2. Quản lý hàng hóa bằng sổ sách ghi tay hoặc tin nhắn Zalo/WeChat phân tán thay vì cơ sở dữ liệu tập trung; 3. Không kiểm tra đối soát quy cách đóng thùng (pallet/kiện lẻ) dẫn đến rách nát mất tem khi bốc dỡ qua nhiều chặng.",
        "mnemonic_tip": "Ví von mã định danh như 'Căn cước công dân của thùng hàng': Hàng không có căn cước thì là hàng vô danh, không thể nào bước qua cửa khẩu biên giới hợp pháp.",
        "comparison": "Quản trị thủ công tiểu ngạch vs Quản trị số hóa chính ngạch (ERP): Quản trị thủ công phụ thuộc 100% vào trí nhớ của bốc vác, rủi ro mất hàng cực cao; Quản trị ERP có log vết điện tử từng giây, minh bạch tuyệt đối trước khách hàng.",
        "persona": "Anh Minh (Chủ doanh nghiệp thương mại) & Chị Thúy (Nhà sáng lập điều hành chuỗi logistics)",
        "checklist": "[ ] Tạo mã khách hàng trên hệ thống trước khi đặt cọc; [ ] Nhà cung cấp chụp ảnh tem dán trên thùng trước khi gửi; [ ] Kho ngoại quét mã thành công trong vòng 2 giờ kể từ khi nhận; [ ] Hồ sơ tờ khai khớp 100% với dữ liệu ERP.",
        "url": "https://en.wikipedia.org/wiki/Supply_chain_management",
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Hệ thống quản trị nội bộ Eureka ERP & daily_notes.md (ghi nhận ngày 24/09/2026 từ case study thực tế Jasmine Nguyen).",
        "keywords": ["supply", "chain"]
    }
]
