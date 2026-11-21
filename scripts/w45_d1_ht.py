# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W45 DAY 1 — 04/11/2026)
Số lượng: 5 ý tưởng (STT 201 - 205)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W45_D1 = [
    {
        "stt": 201,
        "week": "W45 — 04/11/2026",
        "title": "5 Bước Tự Kiểm Tra & Khắc Phục Lỗi 'Con Lăn Chuột Nhảy Dữ Liệu' Trên Phần Mềm Khai Báo Hải Quan ECUS5-VNACCS: Tránh Lỗi Khai Sai Đơn Giá, Số Lượng Khiến Tờ Khai Bị Luồng Đỏ Kiểm Hóa",
        "workflow": "Quy trình 5 bước kiểm soát thao tác khai hải quan an toàn tuyệt đối: Bước 1: Trước khi nhập dữ liệu tờ khai mới, tắt tính năng Scroll wheel của chuột trong giao diện hệ thống hoặc cài đặt chuột chuyển sang chế độ click từng nấc; Bước 2: Khi nhập liệu vào các ô Danh mục dropdown (như Đơn vị tính, Mã số thuế, Mã nước xuất xứ, Điều kiện giao hàng), tuyệt đối dùng phím Tab hoặc Enter để chuyển trường, không lăn con chuột khi con trỏ đang nằm trong ô nhập; Bước 3: Sau khi hoàn thành nhập liệu, bấm nút 'In tờ khai nháp' ra bản PDF, áp dụng nguyên tắc kiểm tra chéo 4 mắt (Four-eyes check): Đối chiếu số lượng, đơn vị tính, tổng trọng lượng (Gross/Net) và tổng trị giá Invoice với bộ chứng từ gốc; Bước 4: Chạy tính năng 'Kiểm tra lỗi tờ khai' tích hợp sẵn trên ECUS5 để phần mềm quét các sai lệch logic về đơn giá và thuế suất; Bước 5: Bấm nút 'Truyền tờ khai' và lưu lại file sao lưu (.xml) trên ổ đĩa bảo mật nội bộ.",
        "formula": "Công thức kiểm tra logic bắt buộc: Tỷ số Trọng lượng trung bình = Tổng Gross Weight / Tổng Số lượng kiện. Nếu con số chênh lệch quá 200% so với Invoice -> Dừng truyền tờ khai ngay lập tức để rà soát đơn vị tính.",
        "top3_mistakes": "1. Quá tự tin vào tốc độ gõ phím mà bỏ qua bước in tờ khai nháp kiểm tra chéo; 2. Thao tác vừa lăn chuột vừa gõ phím làm nhảy ô dropdown mà không hề hay biết; 3. Không đối chiếu trường Đơn vị tính hải quan (Customs Unit) với Đơn vị tính thương mại của nhà sản xuất.",
        "mnemonic_tip": "Ví con lăn chuột trên ECUS như 'cò súng đang mở chốt' - Chạm nhẹ một li, đi tong cả tỷ bạc!",
        "comparison": "Tờ khai truyền trực tiếp (Tự tin mù quáng, dễ vướng lỗi sơ đẳng) vs Tờ khai qua quy trình kiểm tra chéo 5 bước (Chắc chắn 100% không bao giờ văng luồng đỏ oan).",
        "persona": "Anh Minh (Chủ doanh nghiệp sản xuất phụ trợ) & Chị Hương (Kế toán kiêm chứng từ XNK mới vào nghề).",
        "checklist": "[ ] Đã dùng phím Tab chuyển trường? [ ] Đã in bản nháp PDF? [ ] Đơn vị tính khớp Invoice? [ ] Tỷ lệ Gross Weight/Số lượng hợp lý? [ ] Đã chạy kiểm tra logic phần mềm?",
        "url": "https://en.wikipedia.org/wiki/Electronic_data_interchange",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Sổ tay hướng dẫn vận hành phần mềm ECUS5-VNACCS (Thái Sơn), Cổng thông tin Tổng cục Hải quan & daily_notes.md (04/11/2026).",
        "keywords": ["electronic", "data"]
    },
    {
        "stt": 202,
        "week": "W45 — 04/11/2026",
        "title": "Quy Trình 6 Bước Đưa Hàng Hóa Nhập Khẩu Về Kho Riêng Bảo Quản Theo Điều 32 Thông Tư 38/2015/TT-BTC: Cách Tiết Kiệm Hàng Trăm Triệu Tiền Lưu Bãi Khi Chờ Kết Quả Kiểm Tra Chuyên Ngành",
        "workflow": "Quy trình 6 bước đưa hàng về kho riêng bảo quản: Bước 1: Nộp tờ khai hải quan và xuất trình giấy tờ chứng minh hàng hóa thuộc diện phải kiểm tra chuyên ngành (kiểm dịch, ATTP, hợp quy); Bước 2: Chuẩn bị Bộ hồ sơ đề nghị đưa hàng về bảo quản gồm: Giấy đề nghị đưa hàng về bảo quản theo Mẫu số 09/BQHH/GSQL, Hợp đồng thuê kho hoặc Giấy chứng nhận quyền sở hữu kho bãi đáp ứng điều kiện bảo quản nguyên trạng hàng hóa; Bước 3: Nộp kèm Giấy đăng ký kiểm tra chuyên ngành có xác nhận đã tiếp nhận của cơ quan kiểm tra có thẩm quyền; Bước 4: Cơ quan Hải quan cửa khẩu kiểm tra điều kiện kho bãi và ra Quyết định cho phép đưa hàng về bảo quản trong vòng 02 giờ làm việc; Bước 5: Thực hiện niêm phong hải quan (seal chì cơ học hoặc seal điện tử) và lập biên bản bàn giao cho người vận chuyển áp tải hàng về địa điểm bảo quản đã đăng ký; Bước 6: Sau khi có kết quả kiểm tra chuyên ngành đạt chuẩn, nộp kết quả cho Hải quan trong vòng 30 ngày để hoàn tất thủ tục thông quan chính thức.",
        "formula": "Khoản tiết kiệm thực tế = (Phí lưu bãi cảng/ngày - Chi phí kho riêng/ngày) x Số ngày chờ kiểm định (thường từ 7 đến 21 ngày) = Tiết kiệm từ 50 triệu đến 200 triệu đồng/lô hàng.",
        "top3_mistakes": "1. Tự ý bốc dỡ, phá niêm phong chì hoặc bán hàng hóa trước khi có quyết định thông quan chính thức (Bị xử phạt từ 40 - 80 triệu đồng theo Điều 13 Nghị định 128/2020/NĐ-CP); 2. Địa chỉ kho bảo quản trên đơn không khớp với địa chỉ thực tế khi hải quan kiểm tra đột xuất; 3. Quá thời hạn 30 ngày không nộp kết quả kiểm tra chuyên ngành mà không có văn bản giải trình lý do chính đáng.",
        "mnemonic_tip": "Đưa hàng về kho riêng giống như 'tại ngoại có bảo lãnh' - Hàng được về nhà nhưng vòng chì niêm phong của pháp luật vẫn phải giữ nguyên!",
        "comparison": "Để hàng lưu bãi cảng chờ kiểm định (Chi phí đắt đỏ, dễ hư hỏng, phụ thuộc thời tiết) vs Đưa hàng về kho riêng bảo quản (Chủ động bảo quản kho lạnh, tiết kiệm chi phí, giải phóng vỏ cont nhanh).",
        "persona": "Anh Dũng (Giám đốc Logistics & Supply Chain) & Anh Minh (Chủ xưởng sản xuất nhập khẩu linh kiện/thực phẩm).",
        "checklist": "[ ] Có Mẫu 09/BQHH hợp lệ? [ ] Có Giấy đăng ký kiểm tra chuyên ngành dấu đỏ? [ ] Kho bảo quản đủ điều kiện an toàn? [ ] Niêm phong hải quan nguyên vẹn? [ ] Hẹn lịch lấy kết quả trong 30 ngày?",
        "url": "https://en.wikipedia.org/wiki/Customs",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Thư Viện Pháp Luật (Điều 32 Thông tư 38/2015/TT-BTC, sửa đổi bổ sung tại Thông tư 39/2018/TT-BTC, Điều 13 Nghị định 128/2020/NĐ-CP) & daily_notes.md (04/11/2026).",
        "keywords": ["customs", "duty"]
    },
    {
        "stt": 203,
        "week": "W45 — 04/11/2026",
        "title": "Hướng Dẫn 4 Bước Tra Cứu & Xác Minh C/O Điện Tử Trên Cổng Một Cửa Quốc Gia (NSW) & Hệ Thống Hải Quan Nước Cấp: Tuyệt Chiêu Chống Bị Bác Bỏ C/O Ưu Đãi Đặc Biệt Khi Tham Vấn Xuất Xứ",
        "workflow": "Quy trình 4 bước thẩm tra tính xác thực của C/O điện tử: Bước 1: Xác định hệ thống tra cứu tương ứng với từng form C/O: Form D (Cổng Một cửa ASEAN - ASW), Form E điện tử (Cổng kiểm tra C/O Hải quan Trung Quốc origin.customs.gov.cn), Form AK/VK (Hải quan Hàn Quốc customspass.go.kr); Bước 2: Nhập chính xác 3 trường thông tin then chốt: Số tham chiếu C/O (Reference Number), Ngày cấp (Date of Issuance), và Mã số xác thực (Verification Code/FOB Invoice Number); Bước 3: Đối chiếu chéo 100% dữ liệu hiển thị trên màn hình với bản in giấy hoặc file PDF: Tên người xuất khẩu (Exporter), Tên người nhận hàng (Consignee), Mô tả hàng hóa, Mã HS 6 số, Tiêu chí xuất xứ (WO, PE, RVC, CTH) và Số lượng/Trọng lượng; Bước 4: Lưu ảnh chụp màn hình kết quả tra cứu thành công có dấu thời gian (Timestamp) đính kèm vào hồ sơ hải quan điện tử để làm bằng chứng bảo vệ trước đoàn kiểm tra tham vấn giá và xuất xứ.",
        "formula": "Công thức thẩm định C/O hợp lệ: Số tham chiếu thật trên cổng nước cấp + Tiêu chí xuất xứ đúng quy định FTA + Mã HS 6 số trùng khớp tờ khai = 100% Thuế suất ưu đãi đặc biệt 0%.",
        "top3_mistakes": "1. Sử dụng C/O scan từ file trung gian của bên bán mà không chủ động tự tra cứu trên website chính thức của hải quan nước cấp; 2. Tiêu chí xuất xứ trên C/O ghi chung chung hoặc không phù hợp với quy tắc cụ thể mặt hàng (PSR) của hiệp định; 3. Ngày phát hành C/O sau ngày tàu chạy quá 03 ngày làm việc mà không có tích chọn ô 'Issued Retroactively' (Cấp sau).",
        "mnemonic_tip": "C/O không tra cứu được trên mạng cũng giống như 'tiền polymer không có hình chìm' - Đẹp mắt đến mấy hải quan cũng gạt ra ngoài!",
        "comparison": "Tin tưởng mù quáng vào bản scan do đối tác gửi vs Chủ động tra cứu trực tiếp trên Cổng cơ quan thẩm quyền quốc tế (An toàn tuyệt đối).",
        "persona": "Chị Hương (Kế toán kho kiêm chứng từ xuất nhập khẩu) & Anh Minh (Chủ doanh nghiệp muốn tối ưu thuế).",
        "checklist": "[ ] Đã tra cứu trên website hải quan nước cấp? [ ] Số tham chiếu khớp 100%? [ ] Tiêu chí xuất xứ (RVC/CTH) hợp lệ? [ ] Ô Issued Retroactively đã tích nếu cấp sau? [ ] Đã lưu bằng chứng tra cứu có timestamp?",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Cổng Một cửa Quốc gia (vnsw.gov.vn), Cổng tra cứu C/O Hải quan các nước ASEAN, Thư Viện Pháp Luật (Thông tư 05/2018/TT-BCT) & daily_notes.md (04/11/2026).",
        "keywords": ["origin", "certificate"]
    },
    {
        "stt": 204,
        "week": "W45 — 04/11/2026",
        "title": "Nghệ Thuật Sử Dụng Bảo Lãnh Thuế Hải Quan Qua Ngân Hàng (Bank Guarantee for Customs Duty): Cách Giải Phóng Hàng Trong 30 Phút Mà Không Bị Đọng Vốn Tiền Mặt Hàng Tỷ Đồng",
        "workflow": "Quy trình 5 bước vận hành bảo lãnh thuế hải quan: Bước 1: Doanh nghiệp đàm phán ký kết Hợp đồng cấp hạn mức bảo lãnh thuế hải quan với ngân hàng thương mại (Vietcombank, BIDV, Techcombank...); Bước 2: Căn cứ vào số tiền thuế nhập khẩu và thuế GTGT ước tính trên tờ khai, ngân hàng phát hành Thư bảo lãnh điện tử (Bảo lãnh chung theo hạn mức hoặc Bảo lãnh riêng theo từng tờ khai) truyền trực tiếp sang Cổng thanh toán điện tử của Tổng cục Hải quan; Bước 3: Khi mở tờ khai trên phần mềm ECUS, người khai tích chọn phương thức 'Nộp thuế qua bảo lãnh ngân hàng' và điền Mã bảo lãnh do ngân hàng cung cấp; Bước 4: Hệ thống Kế toán thuế tập trung (KTTT) của Hải quan tự động kiểm tra số dư hạn mức bảo lãnh và ghi nhận hoàn thành nghĩa vụ thuế chỉ trong vòng 3 - 5 phút; Bước 5: Tờ khai được giải phóng thông quan ngay lập tức, doanh nghiệp có thêm từ 30 đến 60 ngày ân hạn vốn lưu động để thanh toán tiền thuế thực tế cho ngân hàng.",
        "formula": "Lợi ích tài chính: Lợi nhuận dòng vốn lưu động = Số tiền thuế tạm hoãn x Tỷ suất sinh lời vốn kinh doanh/tháng - Phí bảo lãnh ngân hàng (thường chỉ 0.15% - 0.25%/tháng) = Lãi ròng hàng chục đến hàng trăm triệu đồng dòng tiền.",
        "top3_mistakes": "1. Để thời hạn hiệu lực của Thư bảo lãnh ngắn hơn thời hạn quy định thông quan của tờ khai; 2. Sai lệch 1 ký tự giữa Tên doanh nghiệp/Mã số thuế trên thư bảo lãnh với dữ liệu đăng ký hải quan; 3. Không thanh toán tiền thuế đúng hạn bảo lãnh khiến ngân hàng trích tiền ký quỹ tự động và bị xếp hạng tín nhiệm xấu.",
        "mnemonic_tip": "Bảo lãnh thuế giống như 'thẻ tín dụng VIP của doanh nghiệp tại cửa khẩu' - Hàng cứ việc đi trước, tiền thuế thanh toán thong thả sau!",
        "comparison": "Nộp tiền mặt chuyển khoản thủ công (Đọng vốn lưu động, chờ gạch nợ lâu, dễ trễ chuyến) vs Sử dụng Bảo lãnh thuế ngân hàng điện tử (Thông quan tức thì, tối ưu dòng tiền kinh doanh).",
        "persona": "Anh Minh (Chủ doanh nghiệp cần quay vòng vốn nhanh) & Anh Dũng (Giám đốc Tài chính/Logistics).",
        "checklist": "[ ] Hạn mức bảo lãnh ngân hàng còn đủ? [ ] Thư bảo lãnh đã truyền lên cổng Hải quan? [ ] Mã số thuế trên bảo lãnh chuẩn xác 100%? [ ] Thời hạn bảo lãnh còn tối thiểu 45 ngày? [ ] Đã đặt lịch nhắc thanh toán trước ngày đáo hạn?",
        "url": "https://en.wikipedia.org/wiki/Bank_guarantee",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Luật Hải quan 2014, Luật Quản lý Thuế 2019, Thư Viện Pháp Luật (Điều 9 Thông tư 184/2015/TT-BTC) & daily_notes.md (04/11/2026).",
        "keywords": ["bank", "guarantee"]
    },
    {
        "stt": 205,
        "week": "W45 — 04/11/2026",
        "title": "Tuyệt Kỹ Áp Mã HS Cho Hàng Hóa Dạng Bộ & Hàng Đóng Gói Đồng Bộ Theo Quy Tắc 3b WCO: Cách Xác Định 'Đặc Trưng Cơ Bản' Để Không Bị Cơ Quan Hải Quan Bác Mã Truy Thu Thuế Sau Thông Quan",
        "workflow": "Quy trình 5 bước xác định mã HS theo Quy tắc 3b (GIR 3b): Bước 1: Xác định hàng hóa có thỏa mãn định nghĩa 'Hàng hóa đóng thành bộ để bán lẻ' hay không (phải gồm ít nhất 2 món khác nhau, đóng chung bao bì, cùng phục vụ một nhu cầu hoặc hoạt động cụ thể); Bước 2: Lập bảng phân tích công năng, cấu tạo, trọng lượng, thể tích và giá trị thành phần của từng món trong bộ sản phẩm; Bước 3: Xác định món hàng nào tạo nên 'Đặc trưng cơ bản' (Essential Character) của toàn bộ sản phẩm: Đánh giá xem nếu bỏ món đó ra thì sản phẩm có còn thực hiện được chức năng chính hay không; Bước 4: Áp mã HS của toàn bộ cho mã HS của thành phần tạo nên đặc trưng cơ bản đó theo biểu thuế XNK hiện hành; Bước 5: Soạn thảo Hồ sơ giải trình cơ sở kỹ thuật Quy tắc 3b đính kèm catalogue, bảng phân tích giá thành (Cost Breakdown) và gửi công văn xin xác định trước mã số (nếu mặt hàng có giá trị lớn hoặc tính chất phức tạp).",
        "formula": "Công thức Quy tắc 3b: Mã HS của Bộ sản phẩm = Mã HS của Thành phần chi phối chức năng cốt lõi (Chiếm ưu thế về công năng vận hành hoặc giá trị cấu thành trên 60%).",
        "top3_mistakes": "1. Khai tách rời từng món thành các dòng hàng riêng biệt dù hàng đóng gói nguyên bộ để bán lẻ, dẫn đến bị phạt lỗi khai sai cấu trúc hàng hóa; 2. Nhầm lẫn giữa món hàng có giá tiền cao nhất với món hàng tạo nên 'đặc trưng cơ bản'; 3. Đóng gói các sản phẩm không cùng phục vụ một mục đích sử dụng chung nhưng vẫn cố tình gộp bộ để hưởng thuế suất thấp.",
        "mnemonic_tip": "Quy tắc 3b giống như 'chọn nhạc trưởng của dàn nhạc' - Ai giữ linh hồn và giai điệu chính của tác phẩm, cả dàn nhạc sẽ mang tên người đó!",
        "comparison": "Tách rời từng linh kiện khai báo (Tăng dòng hàng, thuế suất phân tán, rủi ro phân loại) vs Áp mã đồng bộ theo Quy tắc 3b (Tối ưu thủ tục, minh bạch thuế suất theo chuẩn WCO).",
        "persona": "Anh Minh (Chủ DN sản xuất, gia công thiết bị) & Chị Hương (Nhân viên chứng từ tra mã HS).",
        "checklist": "[ ] Hàng có đóng chung bao bì bán lẻ? [ ] Các món có cùng phục vụ 1 mục đích cụ thể? [ ] Đã xác định rõ thành phần tạo đặc trưng cơ bản? [ ] Có bảng kê giá thành chi tiết từng món? [ ] Đã đối chiếu 6 Quy tắc tổng quát WCO?",
        "url": "https://en.wikipedia.org/wiki/Harmonized_System",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: 6 Quy tắc Tổng quát Giải thích Phân loại Hàng hóa theo Danh mục HS của Tổ chức Hải quan Thế giới (WCO), Thông tư 31/2022/TT-BTC, Thư Viện Pháp Luật & daily_notes.md (04/11/2026).",
        "keywords": ["harmonized", "system"]
    }
]
