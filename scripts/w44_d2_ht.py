# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W44 DAY 2 — 30/10/2026)
Số lượng: 5 ý tưởng (STT 181 - 185)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W44_D2 = [
    {
        "stt": 181,
        "week": "W44 — 30/10/2026",
        "title": "4 Bước Rà Soát Xuất Xứ Hàng Hóa Thuần Túy (WO) & Hàm Lượng Giá Trị Khu Vực (RVC): Kỹ Thuật Tự Kiểm Tra Khử Trùng Rủi Ro Gian Lận Xuất Xứ Trước Khi Đăng Ký C/O",
        "workflow": "Bước 1: Tra cứu Bảng Tiêu chí Xuất xứ cụ thể cho từng mã HS (Product Specific Rules - PSR) trong phụ lục Hiệp định thương mại hoặc Nghị định 31/2018/NĐ-CP để xác định mặt hàng áp dụng tiêu chí nào: Chuyển đổi mã số hàng hóa (CTC), Hàm lượng giá trị khu vực (RVC) hay Công đoạn gia công chế biến cụ thể (SP); Bước 2: Bóc tách Bảng kê khai chi phí nguyên vật liệu (BOM - Bill of Materials): Phân định rõ ràng nguyên liệu có xuất xứ nội địa (VNM) có hóa đơn VAT hợp pháp với nguyên liệu nhập khẩu (Non-originating Materials); Bước 3: Áp dụng công thức tính RVC trực tiếp hoặc gián tiếp, kiểm tra tỷ lệ đạt ngưỡng (thường là RVC >= 40% theo giá FOB); Bước 4: Kiểm tra điều kiện loại trừ 'Các công đoạn gia công chế biến đơn giản' theo Điều 9 Nghị định 31/2018/NĐ-CP (bảo quản hàng khi vận chuyển, lau bụi, phân loại, đóng gói lại, gắn nhãn mác, lắp ráp đơn giản các bộ phận) để đảm bảo quy trình sản xuất vượt qua ngưỡng tối thiểu.",
        "formula": "Công thức tính Hàm lượng giá trị khu vực RVC theo phương pháp gián tiếp: RVC = [(FOB - VNM) / FOB] x 100%, trong đó FOB là giá xuất xưởng đã tính cước đến mạn tàu xuất khẩu, VNM là trị giá nguyên liệu đầu vào không có xuất xứ (hoặc không xác định được xuất xứ).",
        "top3_mistakes": "1. Lầm tưởng rằng cứ mua nguyên vật liệu từ công ty thương mại tại Việt Nam là tự động được tính là nguyên liệu có xuất xứ Việt Nam (trong khi công ty thương mại đó nhập khẩu nguyên liệu từ Trung Quốc); 2. Đưa các chi phí không hợp lệ như chi phí vận chuyển chặng quốc tế hoặc lợi nhuận dự kiến chưa thực hiện vào công thức tính RVC; 3. Không lưu trữ bảng định mức sản xuất thực tế và hóa đơn chứng từ mua nguyên phụ liệu theo từng lô hàng, dẫn đến bị bác C/O khi hải quan kiểm tra sau thông quan.",
        "mnemonic_tip": "Nhớ quy tắc '3 KHÔNG khi chứng minh C/O': KHÔNG nhập cắt sẵn chỉ để bắt vít - KHÔNG mua hóa đơn khống nguyên liệu - KHÔNG vứt bỏ hồ sơ chứng từ sản xuất trong vòng 5 năm.",
        "comparison": "So sánh Tiêu chí Chuyển đổi mã HS (CTC) vs Hàm lượng giá trị khu vực (RVC): CTC tập trung vào việc nguyên liệu đầu vào và thành phẩm đầu ra khác nhau ở cấp độ 2 số (CC), 4 số (CTH) hoặc 6 số (CTSH), kiểm tra nhanh và rõ ràng về mặt kỹ thuật; trong khi RVC dựa trên tỷ lệ giá trị tiền tệ gia tăng thực tế, linh hoạt hơn nhưng đòi hỏi hệ thống kế toán chi phí cực kỳ chặt chẽ.",
        "persona": "Anh Minh - Giám đốc nhà máy sản xuất linh kiện cơ khí và đồ gỗ xuất khẩu đang chuẩn bị mở rộng sang thị trường Mỹ và Châu Âu.",
        "checklist": "1. Bản định mức nguyên vật liệu (BOM) có chữ ký giám đốc và dấu giáp lai; 2. Hóa đơn GTGT và tờ khai hải quan nhập khẩu nguyên vật liệu đầu vào; 3. Bảng tính RVC chi tiết đính kèm hóa đơn xuất khẩu FOB; 4. Sơ đồ quy trình sản xuất chứng minh có công đoạn chế biến sâu vượt ngưỡng gia công đơn giản.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP quy định chi tiết Luật Quản lý ngoại thương về xuất xứ hàng hóa, Thông tư 05/2018/TT-BCT) & Cục Xuất nhập khẩu (Bộ Công Thương).",
        "keywords": ["trade", "origin"]
    },
    {
        "stt": 182,
        "week": "W44 — 30/10/2026",
        "title": "Hướng Dẫn 5 Bước Lập Bảng Dữ Liệu Phát Thải Khí Nhà Kính CBAM (Phạm Vi 1 & 2): Quy Trình Chuẩn Hóa Số Liệu Cho Doanh Nghiệp Kim Loại Xuất Khẩu Sang EU",
        "workflow": "Bước 1: Xác định ranh giới hệ thống sản xuất (System Boundary) và phân loại mã số hàng hóa CN Code (HS 8 chữ số của EU) thuộc phạm vi quản lý của CBAM (như mã 7208 đối với thép cuộn cán nóng, mã 7318 đối với bulong ốc vít thép); Bước 2: Thu thập số liệu hoạt động phát thải trực tiếp Scope 1 (Direct Emissions): Đo lường lượng nhiên liệu hóa thạch tiêu thụ thực tế tại nhà máy (xăng, dầu diesel, khí gas LPG, than đá, than cốc) trong kỳ báo cáo quý; Bước 3: Thu thập số liệu hoạt động phát thải gián tiếp Scope 2 (Indirect Emissions): Ghi nhận tổng sản lượng điện năng (kWh) tiêu thụ phục vụ riêng cho các máy móc, thiết bị sản xuất lô hàng; Bước 4: Áp dụng hệ số phát thải chuẩn (Emission Factors): Nhân lượng tiêu thụ nhiên liệu với hệ số phát thải theo hướng dẫn GHG Protocol và nhân sản lượng điện với Hệ số phát thải lưới điện quốc gia của Việt Nam; Bước 5: Điền dữ liệu vào Mẫu bảng tính CBAM Communication Template của Ủy ban Châu Âu và trích xuất file định dạng chuẩn nộp cho nhà nhập khẩu EU.",
        "formula": "Tổng cường độ phát thải riêng (Specific Embedded Emissions - SEE) trên 1 tấn sản phẩm: SEE = (DirEm + IndirEm) / AL, trong đó DirEm là lượng phát thải trực tiếp (tấn CO2e), IndirEm là lượng phát thải gián tiếp từ điện tiêu thụ (tấn CO2e), và AL là tổng sản lượng hàng hóa sản xuất trong kỳ (tấn).",
        "top3_mistakes": "1. Nhầm lẫn giữa tổng lượng phát thải của toàn bộ nhà máy với lượng phát thải phân bổ riêng cho số lượng sản phẩm xuất khẩu sang thị trường EU; 2. Sử dụng hệ số phát thải điện của Châu Âu thay vì áp dụng hệ số phát thải lưới điện quốc gia chính thức của Việt Nam do Bộ Tài nguyên và Môi trường ban hành; 3. Bỏ quên dữ liệu phát thải tích hợp trong các nguyên liệu đầu vào tiền chất (Precursors) như phôi thép, cuộn thép cán nóng mua từ nhà cung cấp bên ngoài.",
        "mnemonic_tip": "Nhớ công thức '2 NGUỒN 1 CHUẨN': Nguồn 1 đốt tại lò (Scope 1) + Nguồn 2 điện kéo lưới (Scope 2) = Chuẩn hóa báo cáo CBAM đúng hẹn 31/10.",
        "comparison": "So sánh Báo cáo Phát thải Scope 1 vs Scope 2: Scope 1 là phát thải trực tiếp từ các nguồn thuộc quyền sở hữu hoặc kiểm soát của nhà máy (lò hơi, ống khói, xe nâng chở phôi), doanh nghiệp hoàn toàn chủ động đo đếm bằng hóa đơn mua nhiên liệu; Scope 2 là phát thải gián tiếp sinh ra từ quá trình phát điện của nhà máy điện bán vào lưới quốc gia, phụ thuộc hoàn toàn vào cơ cấu năng lượng quốc gia (than, khí, thủy điện, điện gió).",
        "persona": "Chị Hương - Kế toán trưởng kiêm phụ trách tuân thủ ESG tại nhà máy sản xuất thép và nhôm định hình xuất khẩu.",
        "checklist": "1. Bảng kê hóa đơn mua dầu FO/DO, khí gas phục vụ lò nung; 2. Hóa đơn tiền điện và biên bản chốt chỉ số công tơ điện phân xưởng; 3. Văn bản xác nhận hệ số phát thải lưới điện quốc gia Việt Nam mới nhất; 4. File Excel CBAM Communication Template đã điền đầy đủ thông số SEE và Precursors.",
        "url": "https://en.wikipedia.org/wiki/Carbon_accounting",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin CBAM Ủy ban Châu Âu, Bộ Tài nguyên và Môi trường Việt Nam & Wikipedia (Carbon accounting).",
        "keywords": ["greenhouse", "carbon"]
    },
    {
        "stt": 183,
        "week": "W44 — 30/10/2026",
        "title": "Hướng Dẫn 4 Bước Chuyển Đổi & Kiểm Thử API Khai Báo Hải Quan Chuẩn JSON DataFormat: Cách Phòng Ngừa Lỗi Truyền Tờ Khai Khi Nâng Cấp Hệ Thống",
        "workflow": "Bước 1: Tải tài liệu đặc tả kỹ thuật API (API Specification Document) và lược đồ dữ liệu JSON Schema chính thức ban hành kèm theo Quyết định 1434/QĐ-CHQ trên Cổng thông tin điện tử Tổng cục Hải quan; Bước 2: Thiết lập bảng ánh xạ dữ liệu (Data Mapping): Chuyển đổi toàn bộ các trường thông tin tờ khai từ định dạng XML/EDIFACT cũ sang các cặp khóa-giá trị (Key-Value Pairs) chuẩn JSON, đặc biệt chú ý định dạng ngày giờ chuẩn ISO 8601 (YYYY-MM-DDTHH:mm:ssZ) và mã hóa ký tự UTF-8 không có BOM; Bước 3: Đăng ký tài khoản và thực hiện kiểm thử trên môi trường Sandbox của Hải quan: Gửi các gói tin mẫu (Mock Requests) cho từng loại hình tờ khai (nhập kinh doanh, xuất kinh doanh, gia công, chế xuất) để kiểm tra mã phản hồi HTTP 200 OK và xác thực chữ ký số điện tử (Digital Signature); Bước 4: Thiết lập cơ chế tự động ghi nhật ký lỗi (Error Logging) và cơ chế chuyển đổi dự phòng (Fallback Mechanism) sang cổng Web nếu đường truyền API gặp sự cố.",
        "formula": "Cấu trúc gói tin JSON hợp lệ gửi đến NSW API: { 'Header': { 'MessageId': 'UUID', 'Sender': 'TaxCode', 'Recipient': 'Customs', 'Timestamp': 'ISO8601' }, 'Body': { 'DeclarationNumber': 'String', 'Items': [ { 'HSCode': '8digits', 'Quantity': Number, 'InvoiceValue': Number } ] } }.",
        "top3_mistakes": "1. Để sót lỗi cú pháp JSON kinh điển như thừa dấu phẩy ở phần tử cuối cùng của mảng (Trailing Comma) hoặc dùng dấu ngoặc kép sai quy chuẩn; 2. Không kiểm tra giới hạn kích thước gói tin (Payload Limit), dẫn đến việc gửi file đính kèm PDF quá dung lượng bị máy chủ hải quan ngắt kết nối; 3. Không cập nhật chứng thư số số hóa trên cổng API khiến gói tin bị từ chối xác thực chữ ký điện tử.",
        "mnemonic_tip": "Nhớ nguyên tắc '3 ĐÚNG khi truyền JSON': Đúng khóa (Keys) - Đúng kiểu dữ liệu (Types) - Đúng giờ quốc tế (Timestamp ISO).",
        "comparison": "So sánh Chuẩn XML vs Chuẩn JSON trong khai báo hải quan: XML có cấu trúc thẻ đóng mở chặt chẽ nhưng dung lượng file rất nặng (chứa nhiều ký tự thẻ dư thừa) và xử lý tốn tài nguyên máy chủ; trong khi JSON có cú pháp tinh gọn, nhẹ hơn từ 40 - 60%, tốc độ phân tích cú pháp (parsing) cực nhanh và dễ dàng tích hợp với các ứng dụng Web/Mobile hiện đại.",
        "persona": "Anh Dũng - Trưởng phòng Công nghệ thông tin kiêm Quản trị hệ thống ERP tại doanh nghiệp sản xuất và xuất nhập khẩu lớn.",
        "checklist": "1. Tài liệu JSON Schema chuẩn QĐ 1434/QĐ-CHQ; 2. Khóa xác thực API Key và Chứng thư số còn hạn sử dụng; 3. Biên bản kiểm thử thành công trên môi trường Sandbox hải quan; 4. Quy trình vận hành dự phòng xử lý sự cố rớt mạng API.",
        "url": "https://en.wikipedia.org/wiki/REST",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cục CNTT & Thống kê Hải quan (customs.gov.vn), WCO Data Model Guidelines & Wikipedia (REST).",
        "keywords": ["rest", "architecture"]
    },
    {
        "stt": 184,
        "week": "W44 — 30/10/2026",
        "title": "Quy Trình 4 Bước Đăng Ký Danh Mục Máy Móc Thiết Bị Đồng Bộ & Thủ Tục Xác Định Trước Mã HS (Advance Ruling): Bí Quyết Khóa Chặt Thuế 0% Hợp Pháp",
        "workflow": "Bước 1: Thu thập bộ chứng từ kỹ thuật gốc từ nhà chế tạo nước ngoài: Bản vẽ sơ đồ nguyên lý hoạt động, sơ đồ tháo lắp chi tiết từng cụm chi tiết (Exploded View), thông số công suất và hợp đồng mua bán nguyên chuyến; Bước 2: Lập bộ hồ sơ Đăng ký Danh mục thiết bị đồng bộ dạng tháo rời theo Điều 7 Thông tư 14/2015/TT-BTC: Nộp hồ sơ điện tử qua Cổng Một cửa Quốc gia tới Chi cục Hải quan nơi mở tờ khai trước ngày nhập khẩu lô hàng đầu tiên; Bước 3: Trong trường hợp mã HS có yếu tố rủi ro tranh cãi, tiến hành nộp hồ sơ xin Xác định trước mã số (Advance Ruling) theo Điều 28 Luật Hải quan 2014 gửi Tổng cục Hải quan ít nhất 60 ngày trước ngày xuất nhập khẩu; Bước 4: Khai báo số văn bản tiếp nhận danh mục đồng bộ hoặc văn bản xác định trước mã số vào chỉ tiêu ghi chú trên tờ khai điện tử VNACCS và thực hiện quy trình theo dõi trừ lùi phiếu theo dõi lượng hàng hóa qua từng chuyến nhập khẩu.",
        "formula": "Nguyên tắc phân loại Quy tắc 2a (GRI 2a): Mặt hàng chưa hoàn chỉnh hoặc chưa lắp ráp có mã HS = Mã HS của sản phẩm hoàn chỉnh NẾU đã có đặc trưng cơ bản (essential character) của sản phẩm hoàn chỉnh và đã hoàn tất thủ tục đăng ký danh mục đồng bộ theo quy định pháp luật.",
        "top3_mistakes": "1. Nhập hàng về đến cảng rồi mới nộp hồ sơ đăng ký danh mục đồng bộ (bị hải quan bác bỏ quyền áp dụng Quy tắc 2a); 2. Khai báo tên hàng trên tờ khai chung chung là 'linh kiện máy móc' thay vì ghi rõ 'Dây chuyền máy ép nhựa đồng bộ ở dạng tháo rời (chuyến 1/3) theo Danh mục số...'; 3. Không đối chiếu danh sách linh kiện thực tế trong container với bản vẽ kỹ thuật, dẫn đến việc thừa thiếu chi tiết so với hồ sơ đăng ký ban đầu.",
        "mnemonic_tip": "Nhớ câu khẩu quyết 'TIỀN TRẢM HẬU TẤU THÌ CHẾT - ĐĂNG KÝ TRƯỚC HÀNG VỀ MỚI SỐNG': Đăng ký danh mục đồng bộ trước ngày mở tờ khai là nguyên tắc sinh tử để khóa thuế 0%.",
        "comparison": "So sánh Xác định trước mã số (Advance Ruling) vs Phân loại thông thường khi mở tờ khai: Phân loại thông thường tiềm ẩn rủi ro bị bác mã số, chuyển luồng đỏ kiểm hóa và truy thu thuế bất cứ lúc nào trong thời hạn 5 năm kiểm tra sau thông quan; trong khi Xác định trước mã số mang lại văn bản phán quyết có giá trị pháp lý ràng buộc toàn diện đối với mọi cơ quan hải quan trên cả nước trong vòng 3 năm.",
        "persona": "Anh Minh - Giám đốc nhà máy cơ khí kiêm chủ đầu tư dự án nhập khẩu dây chuyền tự động hóa.",
        "checklist": "1. Đơn đề nghị xác định trước mã số (Mẫu số 01/XĐTMS/TXNK); 2. Bản vẽ thiết kế kỹ thuật, hình ảnh thực tế và bản thuyết minh công nghệ có dịch thuật công chứng; 3. Giấy đăng ký danh mục máy móc tháo rời có số tiếp nhận của hải quan; 4. Phiếu theo dõi trừ lùi hàng nhập khẩu theo từng chuyến tàu.",
        "url": "https://en.wikipedia.org/wiki/Harmonized_System",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Luật Hải quan 2014 Điều 28, Thông tư 38/2015/TT-BTC Điều 24) & Cục Thuế XNK - Tổng cục Hải quan.",
        "keywords": ["harmonized", "system"]
    },
    {
        "stt": 185,
        "week": "W44 — 30/10/2026",
        "title": "Công Thức Căn Điểm Rơi Truyền Tờ Khai Hải Quan Tối Ưu Hóa Tỷ Giá Ngoại Tệ: Ma Trận Ra Quyết Định Chiều Thứ Năm Cho Lô Hàng Nhập Khẩu Giá Trị Lớn",
        "workflow": "Bước 1: Vào lúc 17h00 chiều thứ Năm hàng tuần, truy cập Cổng thông tin Vietcombank lấy tỷ giá mua chuyển khoản của đồng tiền khai báo trên hợp đồng (USD, EUR, JPY, CNY) làm tỷ giá dự kiến cho tuần tiếp theo; Bước 2: So sánh tỷ giá tuần mới (E_new) với tỷ giá hải quan đang áp dụng trong tuần hiện tại (E_current) trên Cổng thông tin Tổng cục Hải quan; Bước 3: Tính toán khoảng chênh lệch thuế tiềm năng: Delta_Tax = Trị giá ngoại tệ x (E_new - E_current) x (Thuế suất NK + Thuế suất VAT); Bước 4: Chạy Ma trận ra quyết định thời điểm truyền tờ khai: (a) Nếu Delta_Tax > 0 (Tỷ giá tuần mới tăng): Bắt buộc hoàn thiện hồ sơ và truyền tờ khai trước 23h59 ngày thứ Năm; (b) Nếu Delta_Tax < 0 (Tỷ giá tuần mới giảm): Kiểm tra phí lưu kho bãi DEM/DET dự kiến trong 1-2 ngày; nếu Số tiền thuế tiết kiệm được > Chi phí lưu kho bãi, quyết định giữ tờ khai chờ qua 00h05 sáng thứ Sáu mới truyền; Bước 5: Kiểm tra xác nhận số tờ khai và tỷ giá tính thuế chính thức hiển thị trên ô [Tỷ giá tính thuế] của tờ khai phân luồng VNACCS.",
        "formula": "Công thức Ma trận Điểm rơi Quyết định: Giá trị Lợi ích Ròng (Net Savings) = [Trị giá CIF (Ngoại tệ) x |E_new - E_current| x Tổng thuế suất] - [Phí DEM/DET phát sinh nếu chờ đợi (VND)]. Nếu Net Savings > 5.000.000 VND -> Kích hoạt lệnh điều chỉnh thời điểm truyền tờ khai.",
        "top3_mistakes": "1. Nhầm lẫn giữa tỷ giá bán ra của ngân hàng với tỷ giá mua chuyển khoản được cơ quan hải quan sử dụng làm căn cứ tính thuế; 2. Cố tình chờ đợi sang tuần mới để tiết kiệm vài triệu tiền thuế nhưng không tính toán chi phí cắm điện kho lạnh container hoặc phí phạt lưu bãi cảng hàng chục triệu đồng; 3. Truyền tờ khai vào lúc 23h58 đêm thứ Năm khi đường truyền mạng chập chờn, dẫn đến việc tờ khai nhận dấu thời gian (timestamp) vào 00h02 sáng thứ Sáu và bị áp tỷ giá mới.",
        "mnemonic_tip": "Nhớ câu thần chú 'TỶ GIÁ TĂNG THÌ CHẠY ĐÊM THỨ NĂM - TỶ GIÁ GIẢM THÌ NẰM CHỜ SÁNG THỨ SÁU': Căn giờ chuẩn xác giúp giữ lại tiền thật trong tài khoản doanh nghiệp.",
        "comparison": "So sánh Tỷ giá Tính thuế Hải quan vs Tỷ giá Hạch toán Kế toán Doanh nghiệp: Tỷ giá tính thuế hải quan là tỷ giá cố định áp dụng theo tuần do Tổng cục Hải quan công bố theo Điều 35 Thông tư 38/2015/TT-BTC nhằm mục đích duy nhất là tính số tiền thuế phải nộp vào ngân sách nhà nước; trong khi Tỷ giá hạch toán kế toán là tỷ giá giao dịch thực tế tại thời điểm thanh toán theo Thông tư 200/2014/TT-BTC dùng để ghi nhận nợ phải trả và doanh thu/chi phí của doanh nghiệp.",
        "persona": "Chị Hương - Kế toán trưởng kiêm phụ trách thanh toán quốc tế và quản trị dòng tiền doanh nghiệp nhập khẩu.",
        "checklist": "1. Bảng chốt tỷ giá ngoại tệ Vietcombank chiều thứ Năm; 2. Bảng tính Ma trận Net Savings so sánh thuế và chi phí lưu bãi; 3. Chữ ký số USB Token sẵn sàng trước 23h00; 4. Bản in xác nhận tờ khai hải quan có hiển thị tỷ giá tính thuế đã chốt.",
        "url": "https://en.wikipedia.org/wiki/Exchange_rate",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 38/2015/TT-BTC, Thông tư 200/2014/TT-BTC) & Ngân hàng TMCP Ngoại thương Việt Nam (Vietcombank).",
        "keywords": ["rate", "exchange"]
    }
]
