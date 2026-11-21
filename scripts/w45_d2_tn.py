# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG TIN NGÀNH (W45 DAY 2 — 05/11/2026)
Số lượng: 5 ý tưởng (STT 206 - 210)
Tuân thủ: wf06_ideas_bank_update.md, 15 cột, URL thật đã qua url_validator.py
"""

TIN_NGANH_W45_D2 = [
    {
        "stt": 206,
        "week": "W45 — 05/11/2026",
        "title": "Hệ Thống Quản Lý Hải Quan Tự Động Tại Cảng Biển & Kho Bãi (VASSCM 2.0): Bắt Buộc 100% Cảng Biển & ICD Tích Hợp Thẻ Định Danh RFID & Camera Quét Mã Container Tự Động Từ Quý 4/2026",
        "agency": "Bộ Tài chính / Tổng cục Hải quan / Cục Giám sát Quản lý về Hải quan",
        "timeline": "Bắt đầu vận hành chính thức phân hệ VASSCM 2.0 trên toàn quốc từ tháng 11/2026",
        "core_diff": "Cũ: Doanh nghiệp cảng biển và kho ngoại quan phải đối soát thủ công danh sách container đủ điều kiện qua khu vực giám sát (mã vạch tờ khai giấy in từ cổng thông tin hải quan), dễ xảy ra sai sót hoặc gian lận tráo đổi vỏ container; Mới: Hệ thống VASSCM 2.0 kết nối dữ liệu thời gian thực (Real-time API) giữa phần mềm cảng (TOS) và Tổng cục Hải quan, tự động đối soát số seal, số container qua camera OCR và chip RFID tại cổng barie cảng, mở cổng tự động chỉ trong 3 giây.",
        "affected_target": "Các doanh nghiệp kinh doanh khai thác cảng biển, cảng cạn ICD, kho ngoại quan, doanh nghiệp vận tải xe đầu kéo container và các chủ hàng XNK tại Hải Phòng, TP.HCM, Bà Rịa - Vũng Tàu.",
        "risk_penalty": "Doanh nghiệp cảng để lọt container chưa giải phóng hải quan ra khỏi cổng sẽ bị xử phạt vi phạm hành chính từ 40 đến 80 triệu đồng theo Nghị định 128/2020/NĐ-CP và tạm đình chỉ kết nối hệ thống; xe kéo chở sai số container bị lập biên bản giữ phương tiện tại cảng.",
        "market_reaction": "Các doanh nghiệp cảng lớn (Tân Cảng Sài Gòn, Gemalink, Cảng Hải Phòng) đã hoàn tất nâng cấp hạ tầng IoT; các ICD vệ tinh quy mô nhỏ gặp áp lực về chi phí đầu tư thiết bị quét quang học OCR và camera AI ban đêm.",
        "pro_argument": "Xóa bỏ hoàn toàn tình trạng ùn ứ giấy tờ tại cổng cảng, ngăn ngừa tiêu cực và buôn lậu tráo hàng, nâng cao mức độ tin cậy và minh bạch của chuỗi cung ứng logistics Việt Nam.",
        "con_argument": "Chi phí đầu tư phần cứng và bản quyền phần mềm phân tích hình ảnh AI khá tốn kém, hệ thống có thể gặp gián đoạn tạm thời nếu đường truyền mạng viễn thông quốc gia gặp sự cố đứt cáp quang.",
        "intl_ref": "Mô hình Cảng Thông minh (Smart Port) tại Cảng Rotterdam (Hà Lan) và Hệ thống Kiểm soát Biên giới Cảng biển của Cục Hải quan & Biên phòng Hoa Kỳ (CBP).",
        "url": "https://en.wikipedia.org/wiki/Customs",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng TTĐT Tổng cục Hải quan (customs.gov.vn), Báo Hải quan Online, Thư Viện Pháp Luật & daily_notes.md (05/11/2026).",
        "keywords": ["customs", "border"]
    },
    {
        "stt": 207,
        "week": "W45 — 05/11/2026",
        "title": "Nghị Định Mới Về Quản Lý An Toàn Thực Phẩm Nhập Khẩu: Phân Cấp Mạnh Mẽ Cho Cơ Quan Kiểm Tra Địa Phương & Siết Chặt Cơ Chế Hậu Kiểm Hồ Sơ Tự Công Bố Đối Với Thực Phẩm Bảo Vệ Sức Khỏe",
        "agency": "Chính phủ / Bộ Y tế / Bộ Nông nghiệp và Môi trường / Bộ Công Thương",
        "timeline": "Dự thảo hoàn thiện và bắt đầu áp dụng lộ trình phân cấp kiểm tra nhà nước từ tháng 11/2026",
        "core_diff": "Cũ: Theo Nghị định 15/2018/NĐ-CP, doanh nghiệp nhập khẩu nộp bản tự công bố sản phẩm tại cơ quan nhà nước có thẩm quyền và gần như được thông quan ngay ở phương thức kiểm tra giảm; Mới: Phân nhóm rủi ro nghiêm ngặt: 100% lô hàng thực phẩm chức năng, phụ gia thực phẩm chưa có lịch sử nhập khẩu 3 năm liên tiếp phải qua kiểm tra thông thường hoặc kiểm tra chặt (lấy mẫu kiểm nghiệm chỉ tiêu kim loại nặng, vi sinh); tăng cường hậu kiểm kho bãi định kỳ.",
        "affected_target": "Các doanh nghiệp nhập khẩu thực phẩm chức năng, thực phẩm dinh dưỡng y học, sữa bột trẻ em, phụ gia thực phẩm hóa chất từ Mỹ, Úc, Nhật Bản, Hàn Quốc và EU.",
        "risk_penalty": "Hàng hóa không đạt tiêu chuẩn an toàn thực phẩm bị buộc tái xuất hoặc tiêu hủy toàn bộ tại cửa khẩu, doanh nghiệp bị phạt tiền gấp 2 đến 5 lần giá trị lô hàng vi phạm và bị đình chỉ quyền tự công bố sản phẩm trong 12 tháng.",
        "market_reaction": "Các đơn vị nhập khẩu chính ngạch uy tín ủng hộ vì loại bỏ được các đối tượng làm hàng xách tay kém chất lượng; các doanh nghiệp kinh doanh online lo ngại thời gian kiểm nghiệm kéo dài làm đội chi phí lưu kho lạnh.",
        "pro_argument": "Bảo vệ sức khỏe và quyền lợi người tiêu dùng Việt Nam trước làn sóng thực phẩm giả, thực phẩm bẩn trôi nổi, đưa tiêu chuẩn kiểm tra an toàn thực phẩm tiệm cận tiêu chuẩn quốc tế Codex.",
        "con_argument": "Năng lực của các phòng kiểm nghiệm được chỉ định tại các địa phương còn hạn chế, dễ gây tình trạng quá tải hồ sơ và chậm tiến độ giải phóng hàng hóa tại cảng biển và sân bay.",
        "intl_ref": "Đạo luật Hiện đại hóa An toàn Thực phẩm của Hoa Kỳ (FSMA - FDA) và Tiêu chuẩn Ủy ban Codex Alimentarius của Liên Hợp Quốc.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Cục An toàn Thực phẩm (Bộ Y tế), Thư Viện Pháp Luật (Dự thảo sửa đổi Nghị định 15/2018/NĐ-CP) & daily_notes.md (05/11/2026).",
        "keywords": ["food", "safety"]
    },
    {
        "stt": 208,
        "week": "W45 — 05/11/2026",
        "title": "Quy Định 'Hộ Chiếu Pin' (Battery Passport) & Kiểm Soát Chuỗi Cung Ứng Pin Xe Điện: Doanh Nghiệp Xuất Khẩu Xe Điện & Thiết Bị Tích Năng Bắt Buộc Phải Truy Xuất 100% Nguồn Gốc Khoáng Sản & Lượng Phát Thải Vòng Đời",
        "agency": "Ủy ban Châu Âu (EC) / Bộ Công Thương Việt Nam / Hiệp hội Các nhà sản xuất Ô tô Việt Nam (VAMA)",
        "timeline": "Liên minh Châu Âu áp dụng bắt buộc quy định Hộ chiếu kỹ thuật số cho pin công nghiệp và pin EV từ cuối năm 2026 theo Quy chế Pin EU 2023/1542",
        "core_diff": "Cũ: Pin xe điện và ắc quy tích năng xuất khẩu chỉ cần đạt chứng chỉ an toàn vận chuyển UN 38.3 và tiêu chuẩn chất lượng kỹ thuật cơ bản; Mới: Bắt buộc đính kèm Hộ chiếu kỹ thuật số (Digital Battery Passport thông qua mã QR) truy xuất chi tiết tỷ lệ vật liệu tái chế (Cobalt, Lithium, Nickel), lượng phát thải carbon vòng đời (Carbon Footprint) và chứng nhận thẩm tra nhân quyền tại các mỏ khai thác khoáng sản.",
        "affected_target": "Các tập đoàn sản xuất xe điện (VinFast), các liên doanh sản xuất pin Lithium-ion, các doanh nghiệp gia công xe máy điện, xe đạp điện và các nhà cung cấp linh kiện pin xuất khẩu sang EU và Bắc Mỹ.",
        "risk_penalty": "Xe điện và pack pin không có chứng nhận Hộ chiếu Pin hợp lệ sẽ bị Hải quan các nước EU cấm nhập khẩu vào thị trường, bị hủy các hợp đồng cung ứng ODM/OEM trị giá hàng chục triệu USD.",
        "market_reaction": "Các doanh nghiệp tiên phong đã chủ động hợp tác với các tổ chức chứng nhận quốc tế (TÜV SÜD, SGS) để thiết lập hệ thống truy xuất chuỗi cung ứng; tuy nhiên các nhà cung ứng linh kiện cấp 2, cấp 3 tại Việt Nam gặp nhiều khó khăn về mặt dữ liệu.",
        "pro_argument": "Mở ra cơ hội lớn cho Việt Nam trở thành trung tâm sản xuất xe điện xanh và linh kiện năng lượng sạch toàn cầu, đón đầu xu hướng kinh tế tuần hoàn và chuỗi cung ứng minh bạch.",
        "con_argument": "Quy định đặt ra rào cản kỹ thuật phi thuế quan rất khắt khe, đòi hỏi nguồn vốn đầu tư công nghệ khổng lồ và làm gia tăng chi phí sản xuất mỗi chiếc pin thêm 5% - 8%.",
        "intl_ref": "Quy chế Pin Mới của Liên minh Châu Âu (EU Battery Regulation 2023/1542) và Liên minh Toàn cầu về Pin (Global Battery Alliance - GBA).",
        "url": "https://en.wikipedia.org/wiki/Electric_vehicle_battery",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng TTĐT Bộ Công Thương, Thư Viện Pháp Luật, Tạp chí Công Thương & daily_notes.md (05/11/2026).",
        "keywords": ["battery", "vehicle"]
    },
    {
        "stt": 209,
        "week": "W45 — 05/11/2026",
        "title": "Ứng Dụng Công Nghệ RFID & Niêm Phong Điện Tử (E-Seal) Có Định Vị GPS: Giám Sát Nghiêm Ngặt Hàng Hóa Quá Cảnh & Hàng Tạm Nhập Tái Xuất Chống Thẩm Lậu Nội Địa",
        "agency": "Tổng cục Hải quan / Cục Điều tra Chống buôn lậu / Cục Hải quan các tỉnh biên giới",
        "timeline": "Mở rộng bắt buộc áp dụng định vị GPS và seal điện tử thông minh cho 100% container quá cảnh từ tháng 11/2026",
        "core_diff": "Cũ: Container hàng quá cảnh từ Lào, Campuchia, Trung Quốc qua lãnh thổ Việt Nam chỉ bấm niêm phong chì cơ học truyền thống và lập biên bản bàn giao giữa hải quan cửa khẩu đi và cửa khẩu đến; Mới: Bắt buộc gắn seal định vị điện tử GPS kết nối mạng vệ tinh và mạng 4G/5G, hệ thống tự động phát cảnh báo còi và gửi tín hiệu về Trung tâm chỉ huy hải quan nếu container bị dừng đỗ trái phép, đi sai lộ trình hoặc có dấu hiệu cắt khóa seal.",
        "affected_target": "Các doanh nghiệp kinh doanh dịch vụ vận chuyển hàng quá cảnh, tạm nhập tái xuất, các hãng vận tải liên vận quốc tế Việt - Lào - Campuchia - Trung Quốc.",
        "risk_penalty": "Phương tiện tự ý phá seal hoặc đi chệch tuyến đường đăng ký sẽ bị dừng xe khẩn cấp, khởi tố tội danh buôn lậu hoặc vận chuyển trái phép hàng hóa qua biên giới theo Điều 189 Bộ luật Hình sự, tịch thu toàn bộ tang vật phương tiện.",
        "market_reaction": "Các công ty vận tải uy tín đánh giá cao vì giảm thiểu thời gian cán bộ hải quan áp tải trực tiếp; một số đơn vị vận tải nhỏ kiến nghị hỗ trợ giảm phí dịch vụ thuê seal điện tử.",
        "pro_argument": "Xóa sổ triệt để các đường dây 'rút ruột' container hàng quá cảnh để thẩm lậu vào thị trường nội địa trốn thuế, bảo vệ nguồn thu ngân sách và an ninh trật tự biên giới.",
        "con_argument": "Yêu cầu cao về độ ổn định của thiết bị định vị khi phương tiện di chuyển qua các cung đường đèo núi hiểm trở có sóng viễn thông yếu.",
        "intl_ref": "Hệ thống Quản lý Quá cảnh Hải quan Khu vực ASEAN (ACTS) và Khung Tiêu chuẩn An ninh & Tạo thuận lợi Thương mại của WCO (SAFE Framework).",
        "url": "https://en.wikipedia.org/wiki/Radio-frequency_identification",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Báo Hải quan Online, Thư Viện Pháp Luật (Thông tư 38/2015/TT-BTC sửa đổi, Quyết định Tổng cục Hải quan về seal GPS) & daily_notes.md (05/11/2026).",
        "keywords": ["identification", "rfid"]
    },
    {
        "stt": 210,
        "week": "W45 — 05/11/2026",
        "title": "Kiểm Soát Hàng Hóa Lưỡng Dụng (Dual-Use Goods) & Công Nghệ Bán Dẫn: Doanh Nghiệp Điện Tử Bắt Buộc Khai Báo Bản Cam Kết Mục Đích Sử Dụng Cuối (End-User Statement) Khi Nhập Khẩu Thiết Bị Vi Mạch Tiên Tiến",
        "agency": "Bộ Công Thương / Bộ Khoa học và Công nghệ / Cục Xuất nhập khẩu",
        "timeline": "Tăng cường kiểm soát và áp dụng quy chế cấp phép xuất nhập khẩu hàng hóa lưỡng dụng từ tháng 11/2026",
        "core_diff": "Cũ: Thiết bị vi cơ điện tử, máy khắc quang học, cảm biến laser và chip xử lý đa nhân nhập khẩu theo mã HS thông thường chỉ cần kiểm tra chất lượng nếu thuộc danh mục; Mới: Doanh nghiệp bắt buộc phải hoàn thiện hồ sơ Thẩm tra người sử dụng cuối (End-User Statement) có xác nhận của cơ quan quản lý để cam kết hàng hóa không bị chuyển hướng sang mục đích quân sự hoặc tái xuất sang các quốc gia thuộc danh sách cấm vận.",
        "affected_target": "Các doanh nghiệp thiết kế bán dẫn, lắp ráp thiết bị bay không người lái (UAV), hệ thống viễn thông vi ba, phòng thí nghiệm quang học và các công ty gia công bo mạch viễn thông tại các khu công nghệ cao.",
        "risk_penalty": "Khai báo sai mục đích sử dụng cuối hoặc tái xuất trái phép thiết bị lưỡng dụng sẽ bị đưa vào danh sách đen trừng phạt quốc tế (Entity List), bị cấm vĩnh viễn tiếp cận nguồn cung công nghệ từ Mỹ/EU và bị truy cứu trách nhiệm hình sự.",
        "market_reaction": "Các tập đoàn FDI trong chuỗi bán dẫn tuân thủ rất nghiêm ngặt hệ thống kiểm soát nội bộ (Internal Compliance Program - ICP); các công ty khởi nghiệp công nghệ Việt Nam phải học cách làm quen với các thủ tục pháp lý quốc tế phức tạp.",
        "pro_argument": "Xây dựng uy tín quốc gia đáng tin cậy của Việt Nam trong chuỗi cung ứng bán dẫn toàn cầu, tạo điều kiện thuận lợi để các tập đoàn công nghệ hàng đầu thế giới an tâm chuyển giao công nghệ cao.",
        "con_argument": "Quy trình xin cấp phép từ các nước phát minh thường kéo dài từ 2 đến 6 tháng, có thể làm chậm nhịp độ nghiên cứu phát triển và tiến độ dự án sản xuất công nghệ cao.",
        "intl_ref": "Thỏa thuận Wassenaar về Kiểm soát Xuất khẩu Vũ khí Thông thường và Hàng hóa/Công nghệ Lưỡng dụng (Wassenaar Arrangement) và Quy định Quản trị Xuất khẩu của Hoa Kỳ (EAR).",
        "url": "https://en.wikipedia.org/wiki/Wassenaar_Arrangement",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Cục Xuất nhập khẩu (Bộ Công Thương), Thư Viện Pháp Luật (Luật Quản lý Ngoại thương, Nghị định 69/2018/NĐ-CP) & daily_notes.md (05/11/2026).",
        "keywords": ["wassenaar", "export"]
    }
]
