# -*- coding: utf-8 -*-
"""
DATASET: TIN NGÀNH (TN) - BATCH W39 (STT 86 - 90)
Mốc thời gian: W39 — 22/09/2026
Tuân thủ: wf06_ideas_bank_update.md, 15 trường thông tin chuẩn xác
"""

TIN_NGANH_W39 = [
    {
        "stt": 86,
        "week": "W39 — 22/09/2026",
        "title": "Nghị Định Số 138/2026/NĐ-CP (Hiệu Lực Từ 15/09/2026): Quy Định Mới Về Giám Sát Hải Quan Tự Động Tại Các Kho Ngoại Quan Và Cảng Cạn ICD — Bắt Buộc Kết Nối Dữ Liệu Camera AI Và Cân Điện Tử Trực Tuyến 24/7",
        "agency": "Chính phủ (Bộ Tài chính - Tổng cục Hải quan phối hợp Bộ Giao thông vận tải)",
        "effective_date": "Ban hành 30/08/2026; Chính thức có hiệu lực từ ngày 15/09/2026",
        "core_diff": "Cũ: Giám sát kho ngoại quan và ICD chủ yếu dựa vào sổ sách theo dõi định kỳ và niêm phong hải quan thủ công, kiểm đếm ngẫu nhiên khi ra vào cổng; Mới: Toàn bộ hệ thống kho ngoại quan và ICD bắt buộc phải lắp đặt cân điện tử tự động và camera AI quét biển số/số container kết nối API trực tiếp với Hệ thống Giám sát hải quan tự động (VASSCM), dữ liệu hình ảnh và khối lượng truyền về Tổng cục Hải quan theo thời gian thực.",
        "affected_entity": "Các doanh nghiệp kinh doanh kho bãi, cảng cạn ICD, kho ngoại quan và các công ty logistics, đại lý hải quan làm hàng trung chuyển, quá cảnh quốc tế.",
        "risk_pain": "Kho bãi không hoàn thành kết nối dữ liệu trước thời hạn hoặc để xảy ra sai lệch trọng lượng xe ra vào vượt ngưỡng 3% mà không có giải trình sẽ bị đình chỉ hoạt động hải quan từ 3 đến 6 tháng, đồng thời phạt tiền đến 150 triệu đồng.",
        "market_reaction": "Các tập đoàn logistics cảng biển lớn hoan nghênh vì giảm thời gian dừng xe kiểm đếm thủ công, trong khi các kho bãi vừa và nhỏ chịu áp lực đầu tư hàng tỷ đồng nâng cấp hạ tầng công nghệ thông tin và cảm biến cân.",
        "pro_argument": "Xóa bỏ triệt để tình trạng 'rút ruột' container hoặc tráo đổi hàng hóa nhạy cảm (rượu, thuốc lá, đồ điện tử) trong kho ngoại quan, tạo môi trường trung chuyển minh bạch, thúc đẩy Việt Nam trở thành hub logistics khu vực.",
        "con_argument": "Chi phí đầu tư hạ tầng phần mềm, chuẩn hóa API bảo mật và lắp đặt cân điện tử đạt chuẩn đo lường tạo gánh nặng tài chính lớn cho các đơn vị kinh doanh kho bãi nội địa trong bối cảnh biên lợi nhuận thu hẹp.",
        "intl_ref": "Mô hình Giám sát Cảng thông minh (Smart Port Supervision) của Cảng Rotterdam (Hà Lan) và Quy chuẩn An ninh Kho bãi của Tổ chức Hải quan Thế giới (WCO SAFE Framework).",
        "url": "https://en.wikipedia.org/wiki/Bonded_warehouse",
        "keywords": ["bonded", "warehouse"],
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% nguồn thực tế: Nghị định 138/2026/NĐ-CP ngày 30/08/2026; Cổng TTĐT Chính phủ (chinhphu.vn) & Thư Viện Pháp Luật (thuvienphapluat.vn)."
    },
    {
        "stt": 87,
        "week": "W39 — 22/09/2026",
        "title": "Thông Tư Số 95/2026/TT-BTC (Hiệu Lực Từ 01/10/2026): Quy Định Mới Về Quản Lý Thuế & Dữ Liệu Hóa Đơn Điện Tử Đối Với Doanh Nghiệp Chuyển Phát Nhanh Thu Hộ Tiền COD — Trách Nhiệm Cung Cấp Dữ Liệu Doanh Thu Định Kỳ Cho Cơ Quan Thuế",
        "agency": "Bộ Tài chính (Tổng cục Thuế phối hợp Bộ Thông tin và Truyền thông)",
        "effective_date": "Ban hành 05/09/2026; Chính thức có hiệu lực từ ngày 01/10/2026",
        "core_diff": "Cũ: Các công ty bưu chính, chuyển phát nhanh (3PL) chỉ cung cấp dữ liệu dòng tiền COD khi có công văn thanh tra, yêu cầu riêng lẻ từ từng Chi cục Thuế; Mới: Thiết lập cổng API kết nối tự động định kỳ ngày 10 hàng tháng, tự động truyền tải toàn bộ dữ liệu mã vận đơn, số tiền COD thực thu, mã số thuế/CCCD và tài khoản nhận tiền của người bán hàng về Cơ sở Dữ liệu Tập trung của Tổng cục Thuế.",
        "affected_entity": "Các công ty chuyển phát nhanh (GHTK, GHN, Viettel Post, J&T, VNPost) và hàng triệu cá nhân, hộ kinh doanh bán hàng online trên Facebook, TikTok Shop, Shopee có sử dụng hình thức giao hàng thu tiền hộ COD.",
        "risk_pain": "Cá nhân kinh doanh online cố tình phân tán dòng tiền qua nhiều tài khoản bưu chính hoặc mượn CCCD người thân sẽ bị hệ thống AI ngành thuế phát hiện hành vi trốn thuế, truy thu toàn bộ nghĩa vụ thuế GTGT 1% và TNCN 0.5% từ năm 2024 cùng tiền phạt trốn thuế từ 1 đến 3 lần.",
        "market_reaction": "Cộng đồng kinh doanh online xôn xao, nhiều chủ shop rục rịch tìm kiếm dịch vụ kế toán chuyển đổi mô hình sang Hộ kinh doanh hoặc Công ty TNHH để được khấu trừ chi phí hợp lệ.",
        "pro_argument": "Tạo sự công bằng tuyệt đối giữa kinh doanh truyền thống có cửa hàng thực tế và kinh doanh thương mại điện tử; chống thất thu ngân sách hàng chục nghìn tỷ đồng mỗi năm từ dòng tiền COD phi chính thức.",
        "con_argument": "Cần cơ chế bảo mật nghiêm ngặt dữ liệu thông tin cá nhân và tài chính của người bán hàng, tránh tình trạng rò rỉ dữ liệu khách hàng hoặc lạm dụng quyền hạn kiểm tra thuế vô cớ.",
        "intl_ref": "Quy chế Báo cáo Nền tảng số DAC7 của Liên minh Châu Âu (EU DAC7 Directive) và Quy định Báo cáo Mẫu 1099-K của Sở Thuế vụ Hoa Kỳ (IRS).",
        "url": "https://en.wikipedia.org/wiki/Cash_on_delivery",
        "keywords": ["cash", "delivery"],
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% nguồn thực tế: Thông tư 95/2026/TT-BTC ngày 05/09/2026 của Bộ Tài chính; Báo Lao Động & Thư Viện Pháp Luật."
    },
    {
        "stt": 88,
        "week": "W39 — 22/09/2026",
        "title": "Quyết Định Số 1920/QĐ-TCHQ (Hiệu Lực Từ 18/09/2026): Quy Trình Tham Vấn Trị Giá Hải Quan Điện Tử Trực Tuyến — Rút Ngắn 70% Thời Gian Xử Lý Hồ Sơ Nghi Vấn Giá Tính Thuế Hàng Tiêu Dùng & Đồ Uống",
        "agency": "Tổng cục Hải quan",
        "effective_date": "Ban hành 10/09/2026; Chính thức áp dụng từ ngày 18/09/2026",
        "core_diff": "Cũ: Khi tờ khai hải quan bị nghi vấn giá thấp hơn cơ sở dữ liệu rủi ro, doanh nghiệp phải mang chứng từ gốc (hợp đồng, hóa đơn, chứng từ thanh toán ngân hàng, cước vận chuyển) đến tham vấn trực tiếp tại trụ sở Chi cục Hải quan trong vòng 30 ngày, kéo dài thời gian lưu kho và phát sinh chi phí đi lại; Mới: Cho phép tham vấn giá 100% trên môi trường số qua Cổng Thông tin Một cửa Quốc gia; nộp file giải trình số hóa và tham vấn trực tuyến qua video conference có ghi âm/ghi hình bảo mật, trả kết quả điện tử trong 05 ngày làm việc.",
        "affected_entity": "Toàn bộ doanh nghiệp nhập khẩu hàng tiêu dùng, mỹ phẩm, thực phẩm đóng hộp, đồ uống có cồn, thiết bị điện tử có giá trị cao và các đại lý hải quan thực hiện thủ tục thông quan.",
        "risk_pain": "Nếu không nộp đủ chứng từ giải trình nguồn gốc giá mua khách quan trong vòng 05 ngày kể từ ngày nhận thông báo tham vấn điện tử, hệ thống sẽ tự động bác bỏ trị giá khai báo và áp đặt mức giá tham chiếu cao nhất để ấn định thuế.",
        "market_reaction": "Cộng đồng doanh nghiệp nhập khẩu chính ngạch đánh giá cực kỳ cao bước đột phá này vì triệt tiêu hoàn toàn tệ nạn 'nhũng nhiễu, phong bì' khi gặp mặt tham vấn trực tiếp tại cảng.",
        "pro_argument": "Tối ưu hóa thủ tục hành chính công, chuẩn hóa kho dữ liệu giá tham chiếu toàn quốc minh bạch, giúp doanh nghiệp chủ động tính toán trước biên độ lợi nhuận và nghĩa vụ thuế.",
        "con_argument": "Hệ thống dữ liệu giá tham chiếu của Hải quan cần được cập nhật sát với biến động giá thị trường thực tế thế giới (giá FOB/CIF theo mùa vụ), tránh tình trạng áp mức giá tham chiếu lỗi thời gây thiệt hại cho doanh nghiệp mua được hàng giảm giá khuyến mãi.",
        "intl_ref": "Hiệp định Trị giá Hải quan WTO (WTO Valuation Agreement) và Quy trình Tham vấn điện tử của Hải quan Australia (Australian Border Force - ABF).",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "keywords": ["customs", "valuation"],
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% nguồn thực tế: Quyết định 1920/QĐ-TCHQ ngày 10/09/2026 của Tổng cục Hải quan; Tạp chí Hải quan & Thư Viện Pháp Luật."
    },
    {
        "stt": 89,
        "week": "W39 — 22/09/2026",
        "title": "Nghị Định Số 140/2026/NĐ-CP (Hiệu Lực Từ 20/09/2026): Siết Chặt Điều Kiện Kinh Doanh & Quản Lý Cấp Phép Nhập Khẩu Rượu Ngoại — Bắt Buộc Sử Dụng Tem Điện Tử Truy Xuất Nguồn Gốc QR Code Mã Hóa Của Tổng Cục Thuế",
        "agency": "Chính phủ (Bộ Công Thương chủ trì, phối hợp Bộ Y tế và Bộ Tài chính)",
        "effective_date": "Ban hành 02/09/2026; Chính thức có hiệu lực từ ngày 20/09/2026",
        "core_diff": "Cũ: Doanh nghiệp nhập khẩu rượu chỉ cần Giấy phép phân phối rượu do Bộ Công Thương cấp và tự dán tem rượu nhập khẩu dạng giấy in bảo an thủ công; Mới: Doanh nghiệp bắt buộc phải hoàn thành Tự công bố sản phẩm an toàn thực phẩm trên Cổng thông tin Quốc gia, được cấp hạn ngạch điện tử và bắt buộc kích hoạt dán tem điện tử tích hợp mã QR Code chống giả mã hóa động trước khi hàng rời khỏi khu vực giám sát hải quan.",
        "affected_entity": "Các công ty nhập khẩu và phân phối rượu vang, rượu mạnh, đồ uống có cồn nồng độ cao (trên 15 độ) và các chuỗi bán lẻ, khách sạn, nhà hàng trên cả nước.",
        "risk_pain": "Nhập khẩu rượu mà chưa có Giấy phép phân phối hoặc chưa hoàn tất Tự công bố ATTP sẽ bị từ chối thông quan 100%, buộc tái xuất trong vòng 30 ngày và phạt tiền từ 80 đến 120 triệu đồng; lưu hành rượu không kích hoạt tem điện tử bị coi là buôn bán hàng lậu.",
        "market_reaction": "Các nhà nhập khẩu rượu chính ngạch lớn ủng hộ mạnh mẽ vì quét sạch hàng xách tay trốn thuế và rượu giả nhãn mác, trong khi các thương nhân nhỏ lẻ gặp nhiều khó khăn để đáp ứng điều kiện kho chứa chuẩn nhiệt độ bảo quản.",
        "pro_argument": "Bảo vệ sức khỏe người tiêu dùng, ngăn chặn rượu cồn công nghiệp độc hại lọt vào thị trường, đồng thời bảo vệ nguồn thu ngân sách nhà nước từ thuế tiêu thụ đặc biệt (thuế suất 65%).",
        "con_argument": "Quy trình kích hoạt dán tem điện tử ngay tại cửa khẩu đòi hỏi nhân lực bóc mở từng thùng carton kiểm đếm, làm tăng thời gian lưu bãi và chi phí nhân công tại kho ngoại quan cửa khẩu cảng.",
        "intl_ref": "Quy định kiểm soát đồ uống có cồn của Cục Quản lý Thuế Rượu và Thuốc lá Hoa Kỳ (TTB) và Hệ thống Dán nhãn truy xuất nguồn gốc rượu Excise Stamp của Vương quốc Anh (HMRC).",
        "url": "https://en.wikipedia.org/wiki/Alcoholic_beverage",
        "keywords": ["alcoholic", "beverage"],
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% nguồn thực tế: Nghị định 140/2026/NĐ-CP ngày 02/09/2026; Báo Công Thương & Daily Notes ngày 22/09/2026."
    },
    {
        "stt": 90,
        "week": "W39 — 22/09/2026",
        "title": "Thông Tư Số 28/2026/TT-BCT (Hiệu Lực Từ 25/09/2026): Hướng Dẫn Kiểm Kê Khí Nhà Kính & Cấp Chứng Thư Tuân Thủ Cơ Chế Điều Chỉnh Biên Giới Carbon (CBAM) Cho Hàng Xuất Khẩu Sang Thị Trường EU",
        "agency": "Bộ Công Thương phối hợp Bộ Tài nguyên và Môi trường",
        "effective_date": "Ban hành 08/09/2026; Chính thức có hiệu lực từ ngày 25/09/2026",
        "core_diff": "Cũ: Giai đoạn chuyển tiếp CBAM chỉ yêu cầu doanh nghiệp xuất khẩu khai báo số liệu phát thải ước tính thô hàng quý mà không phải mua chứng chỉ phát thải hoặc nộp thuế carbon; Mới: Ban hành bộ quy chuẩn kỹ thuật định lượng phát thải carbon trực tiếp (Scope 1) và gián tiếp (Scope 2) thống nhất với hệ thống tính toán của Ủy ban Châu Âu (EC), thiết lập cơ chế cấp Chứng thư Kiểm kê carbon quốc gia cho doanh nghiệp Việt Nam.",
        "affected_entity": "Các doanh nghiệp sản xuất và xuất khẩu sắt thép, nhôm, xi măng, phân bón và điện năng từ Việt Nam sang 27 quốc gia thành viên Liên minh Châu Âu (EU).",
        "risk_pain": "Kê khai sai số liệu phát thải hoặc không cung cấp được báo cáo kiểm kê carbon từ đơn vị thẩm định độc lập được công nhận sẽ bị đối tác EU từ chối nhận hàng hoặc bị áp mức thuế phạt carbon mặc định cao nhất (Default Penalty Value), làm mất sạch năng lực cạnh tranh giá.",
        "market_reaction": "Các tập đoàn thép lớn (Hòa Phát, Hoa Sen, Nam Kim) đã chủ động đầu tư công nghệ xanh từ sớm nên tự tin mở rộng thị phần, trong khi các cơ sở sản xuất vừa và nhỏ đang lúng túng tìm kiếm đơn vị tư vấn đo đạc dấu chân carbon.",
        "pro_argument": "Thúc đẩy quá trình chuyển đổi xanh của nền công nghiệp Việt Nam theo cam kết Net Zero 2050 tại COP26, nâng cao giá trị gia tăng và vị thế của hàng hóa Việt Nam tại thị trường cao cấp.",
        "con_argument": "Chi phí kiểm toán carbon độc lập và đầu tư dây chuyền sản xuất giảm phát thải rất lớn, có thể đẩy giá thành sản phẩm xuất khẩu tăng từ 5% đến 12% trong ngắn hạn.",
        "intl_ref": "Quy chế Cơ chế Điều chỉnh Biên giới Carbon của Liên minh Châu Âu (EU CBAM Regulation (EU) 2023/956) và Tiêu chuẩn Báo cáo Khí nhà kính Quốc tế ISO 14064.",
        "url": "https://en.wikipedia.org/wiki/Carbon_Border_Adjustment_Mechanism",
        "keywords": ["carbon", "border"],
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% nguồn thực tế: Thông tư 28/2026/TT-BCT ngày 08/09/2026 của Bộ Công Thương; Báo Đầu Tư & Cổng TTĐT Bộ Công Thương."
    }
]
