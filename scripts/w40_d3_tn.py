# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG TIN NGÀNH (W40 DAY 3 — 30/09/2026)
Số lượng: 5 ý tưởng (STT 116 - 120)
Tuân thủ: wf06_ideas_bank_update.md, 15 cột, URL thật đã qua url_validator.py
"""

TIN_NGANH_W40_D3 = [
    {
        "stt": 116,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Thông Tư 15/2010/TT-BCT & Thông Tư 08/2024/TT-BCT: Cơ Chế Khai Báo Hóa Đơn Bên Thứ Ba (Third-Party Invoicing) Trên C/O Mẫu AI (ASEAN - Ấn Độ)",
        "agency": "Bộ Công Thương & Tổng cục Hải quan",
        "timeline": "Có hiệu lực từ 15/06/2024, tiếp tục siết chặt kiểm tra đối chiếu dữ liệu điện tử eCOSys năm 2026",
        "core_diff": "Quy định rõ trường hợp hóa đơn bên thứ ba phát hành bởi công ty có trụ sở tại nước thứ ba: Ô số 7 phải thể hiện tên và địa chỉ công ty phát hành hóa đơn bên thứ ba; Ô số 10 ghi số và ngày hóa đơn bên thứ ba; Ô số 13 bắt buộc tích chọn 'Third-party invoicing'. Trường hợp ô số 9 và 10 thể hiện cả 2 trị giá FOB và 2 hóa đơn (nhà sản xuất + bên thứ ba) thì hệ thống hải quan bắt buộc đối soát số hóa đơn bên thứ ba với hóa đơn thương mại thực tế nộp kèm tờ khai.",
        "affected_target": "Doanh nghiệp nhập khẩu nguyên liệu hóa chất, dược phẩm, dệt may từ Ấn Độ qua các công ty thương mại trung gian tại Singapore, Hong Kong.",
        "risk_penalty": "Bị Hải quan cửa khẩu từ chối áp dụng thuế suất ưu đãi đặc biệt AIFTA (thuế suất từ 0% nhảy lên 10-20% thuế MFN), phát sinh nợ thuế hàng trăm triệu đồng và tiền chậm nộp 0.03%/ngày.",
        "market_reaction": "Các doanh nghiệp nhập khẩu phản ánh nhiều đối tác Ấn Độ và trung gian thương mại khai báo lúng túng giữa hóa đơn nhà máy và hóa đơn thương mại dẫn đến hải quan nghi vấn gian lận trị giá.",
        "pro_argument": "Siết chặt đối chiếu hóa đơn bên thứ ba giúp ngăn chặn hành vi chuyển giá, khai man trị giá tính thuế hoặc lợi dụng C/O của nhà máy để hợp thức hóa hàng hóa của nước thứ ba không thuộc khối AIFTA.",
        "con_argument": "Chuỗi cung ứng quốc tế qua nhiều tầng thương mại là tập quán thông thường; việc máy móc bắt bẻ từng câu chữ tại ô số 9 và ô số 10 làm kéo dài thời gian thông quan, tăng chi phí lưu bãi cho doanh nghiệp tuân thủ.",
        "intl_ref": "Hướng dẫn Operational Certification Procedures (OCP) của Hiệp định AIFTA và quy định Third-Party Invoicing của Hiệp định CPTPP và RCEP.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 08/2024/TT-BCT & Thông tư 15/2010/TT-BCT) & daily_notes.md (vấn đề C/O AI 3 bên ngày 30/09/2026).",
        "keywords": ["certificate", "origin"]
    },
    {
        "stt": 117,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Nghị Định 31/2018/NĐ-CP & Thông Tư 05/2018/TT-BCT: Tiêu Chí Chuyển Đổi Mã Số Hàng Hóa (CTC) & Hàm Lượng Giá Trị Khu Vực (RVC) Đối Với Hàng Nhập Khẩu Đi Gia Công Biến Tính Trong Nước",
        "agency": "Chính phủ & Bộ Công Thương",
        "timeline": "Thi hành theo Nghị định 31/2018/NĐ-CP và Thông tư 05/2018/TT-BCT, trọng tâm kiểm tra xuất xứ hàng hóa phục vụ sản xuất gia công năm 2026",
        "core_diff": "Quy định rõ hàng hóa nhập khẩu gia công nội địa chỉ được coi là có xuất xứ Việt Nam nếu đáp ứng tiêu chí chuyển đổi mã số hàng hóa (CTC) ở cấp 2, 4 hoặc 6 chữ số hoặc đạt tỷ lệ RVC tối thiểu (thường 30-40%). Nếu công đoạn gia công chỉ làm thay đổi cơ lý tính (tăng độ cứng, chống cháy) mà không đổi mã HS và không đạt RVC, sản phẩm bắt buộc phải giữ nguyên xuất xứ của nguyên liệu ban đầu (ví dụ: Made in China) hoặc khai 'Xuất xứ không xác định', tuyệt đối không được ghi Made in Vietnam.",
        "affected_target": "Doanh nghiệp chế xuất (DNCX) và doanh nghiệp FDI nhập khẩu nguyên liệu hạt nhựa, linh kiện điện tử thuê gia công xử lý bề mặt, pha trộn tại các xưởng nội địa.",
        "risk_penalty": "Bị xử phạt hành vi gian lận ghi nhãn xuất xứ hàng hóa theo Nghị định 98/2020/NĐ-CP (sửa đổi bởi Nghị định 17/2022/NĐ-CP), tước quyền ưu đãi thuế quan và đình chỉ hoạt động gia công.",
        "market_reaction": "Nhiều nhà máy gia công lúng túng khi soạn thảo hợp đồng gia công, tranh cãi gay gắt với đối tác về việc ghi nhãn mác thành phẩm và trách nhiệm pháp lý khi kiểm tra sau thông quan.",
        "pro_argument": "Ngăn chặn triệt để tình trạng mượn danh gia công đơn giản để 'rửa xuất xứ', bảo vệ uy tín hàng xuất khẩu Việt Nam trước các đợt điều tra chống lẩn tránh thuế của Mỹ và EU.",
        "con_argument": "Các công đoạn gia công biến tính đòi hỏi công nghệ cao và phụ gia đắt đỏ nhưng vì quy tắc HS không đổi mã nên doanh nghiệp không được ghi nhận giá trị gia tăng tại Việt Nam, gây thiệt thòi trong chuỗi giá trị.",
        "intl_ref": "Khái niệm Substantial Transformation (Chuyển đổi thực chất) theo Công ước Kyoto sửa đổi của Tổ chức Hải quan Thế giới (WCO) và Quy tắc xuất xứ không ưu đãi của WTO.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP & Thông tư 05/2018/TT-BCT) & daily_notes.md (vướng mắc gia công hạt nhựa biến tính ngày 30/09/2026).",
        "keywords": ["rules", "origin"]
    },
    {
        "stt": 118,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Thông Tư 38/2015/TT-BTC (Điều 20 & Điều 51) & Thông Tư 39/2018/TT-BTC: Quy Trình Giám Sát Hải Quan & Thủ Tục Bấm Lại Seal Đối Với Hàng Hóa Chuyển Cửa Khẩu / Chuyển Cảng Biển",
        "agency": "Bộ Tài chính (Tổng cục Hải quan)",
        "timeline": "Thực hiện theo Thông tư 38/2015/TT-BTC và Thông tư 39/2018/TT-BTC, siết chặt giám sát container trung chuyển cảng năm 2026",
        "core_diff": "Quy định bắt buộc: Khi container chuyển cảng (như từ Cát Lái sang Hiệp Phước) bị đứt/gãy seal hãng tàu hoặc seal niêm phong hải quan, đơn vị vận tải và doanh nghiệp cảng phải lập Biên bản ghi nhận hiện trường có xác nhận của đại diện cảng; Cảng tiến hành bấm seal mới và cấp Biên bản thay seal; Doanh nghiệp khi truyền khai báo lấy hàng phải cập nhật số seal mới và đính kèm Biên bản sự cố lên hệ thống điện tử VASSCM để hải quan đối chiếu tự động.",
        "affected_target": "Các hãng tàu, đại lý giao nhận (Forwarder), doanh nghiệp vận tải bộ/sà lan và chủ hàng có cont chuyển tải giữa các cụm cảng nước sâu và cảng cạn (ICD).",
        "risk_penalty": "Số seal trên tờ khai không khớp thực tế bị hệ thống VASSCM chặn tại cổng cảng, bị cơ quan hải quan chuyển luồng Đỏ kiểm hóa 100%, phát sinh chi phí rút ruột, phí lưu bãi (demurrage/detention) từ 15-50 triệu đồng/cont.",
        "market_reaction": "Các tài xế xe kéo và doanh nghiệp logistics bức xúc vì thủ tục phối hợp giữa các cảng còn rườm rà, biên bản giấy thường chậm trễ khiến xe nằm chờ cả ngày tại cổng cảng.",
        "pro_argument": "Giám sát seal nghiêm ngặt là yêu cầu cốt tử để chống buôn lậu, đánh tráo ruột container chứa hàng cấm hoặc rút bớt hàng giá trị cao trong quá trình di chuyển trên đường bộ/sông.",
        "con_argument": "Các va chạm cơ học trong quá trình cẩu bốc xếp làm gãy seal nhựa là sự cố bất khả kháng thường xuyên; cần có cơ chế quét mã QR hoặc cập nhật điện tử tức thì giữa hai cảng thay vì bắt tài xế đi xin xác nhận giấy.",
        "intl_ref": "Tiêu chuẩn an toàn niêm phong container quốc tế ISO 17712:2013 và quy chuẩn giám sát container chuyển tiếp (In-transit Cargo Security) của WCO SAFE Framework.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 38/2015 & 39/2018) & daily_notes.md (sự cố cont gãy seal chuyển cảng Cát Lái - Hiệp Phước ngày 25/09/2026).",
        "keywords": ["security", "seal"]
    },
    {
        "stt": 119,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Luật Quản Lý Thuế Số 38/2019/QH14 & Nghị Định 126/2020/NĐ-CP: Cơ Chế Quét Dữ Liệu Tự Động Hóa Đơn Điện Tử Rủi Ro & Chế Tài Xử Phạt Bên Mua Hàng Từ Doanh Nghiệp Bỏ Trốn",
        "agency": "Quốc hội & Chính phủ (Tổng cục Thuế)",
        "timeline": "Thi hành theo Luật Quản lý thuế 38/2019/QH14, tăng cường ứng dụng AI phân tích dữ liệu hóa đơn điện tử năm 2026",
        "core_diff": "Quy định trách nhiệm giải trình của doanh nghiệp mua hàng: Khi đối tác bán hàng bị cơ quan thuế thông báo 'Không hoạt động tại địa chỉ đăng ký' (công ty bỏ trốn/công ty ma), bên mua dù có chứng từ thanh toán ngân hàng vẫn bị cơ quan thuế yêu cầu chứng minh tính có thực của giao dịch (hợp đồng, biên bản giao nhận, phiếu cân xe, chứng từ vận chuyển logistics). Nếu không chứng minh được, toàn bộ số thuế VAT đầu vào bị loại khấu trừ và chi phí mua hàng bị loại khỏi chi phí hợp lý khi tính thuế TNDN, bị xử phạt 20% hành vi khai sai kèm tiền chậm nộp 0.03%/ngày.",
        "affected_target": "Doanh nghiệp phân phối thương mại, sản xuất cơ khí, vật liệu xây dựng, nông sản thu mua từ nhiều nguồn cung nhỏ lẻ.",
        "risk_penalty": "Bị truy thu hàng tỷ đồng tiền thuế sau 2-3 năm quyết toán, bị phong tỏa tài khoản ngân hàng và thậm chí bị chuyển hồ sơ sang cơ quan công an điều tra nếu giá trị hóa đơn trên 100 triệu đồng.",
        "market_reaction": "Cộng đồng doanh nghiệp kêu trời vì bên bán lúc giao dịch vẫn có mã số thuế hoạt động bình thường, xuất hóa đơn hợp pháp trên Cổng Thuế, việc bên bán bỏ trốn sau đó là rủi ro khách quan nằm ngoài tầm kiểm soát của bên mua.",
        "pro_argument": "Siết chặt truy thu bên mua là biện pháp hữu hiệu nhất để triệt phá thị trường mua bán hóa đơn bất hợp pháp, buộc doanh nghiệp phải có trách nhiệm thẩm định đối tác cung ứng.",
        "con_argument": "Đẩy toàn bộ rủi ro quản lý nhà nước sang vai doanh nghiệp làm ăn chân chính; cơ quan thuế nắm hệ thống quản lý nhưng không cảnh báo kịp thời mà lại đi phạt bên mua thiện chí.",
        "intl_ref": "Quy định về VAT Fraud / Carousel Fraud của Liên minh Châu Âu (EU) và nguyên tắc thẩm định nhà cung cấp 'Know Your Supplier' (KYS) trong quản trị rủi ro thuế quốc tế.",
        "url": "https://en.wikipedia.org/wiki/Tax_evasion",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Luật Quản lý thuế 38/2019/QH14 & Nghị định 126/2020) & daily_notes.md (vướng mắc hóa đơn công ty ma ngày 25/09/2026).",
        "keywords": ["tax", "evasion"]
    },
    {
        "stt": 120,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Quyết Định 1357/QĐ-TCHQ & Công Văn 1478/TCHQ-GSQL: Hướng Dẫn Phân Loại Mã Loại Hình Xuất Nhập Khẩu (A11, A12, A41) Theo Đúng Bản Chất Hoạt Động Của Doanh Nghiệp",
        "agency": "Tổng cục Hải quan (Chi cục Hải quan Khu vực II)",
        "timeline": "Thực hiện theo Quyết định 1357/QĐ-TCHQ và Công văn 1478/TCHQ-GSQL, hướng dẫn nghiệp vụ mới nhất tại Công văn 21939/HQKV2-NVHQ",
        "core_diff": "Quy định bắt buộc: Doanh nghiệp khi mở tờ khai phải xác định mã loại hình căn cứ vào bản chất giao dịch và quyền xuất nhập khẩu, không được chọn tùy tiện: Mã A11 áp dụng cho nhập khẩu tiêu dùng thông thường của doanh nghiệp thuần túy trong nước; Mã A12 áp dụng cho doanh nghiệp FDI nhập khẩu để phục vụ sản xuất hoặc bán buôn theo quyền phân phối; Mã A41 áp dụng khi doanh nghiệp nhập khẩu hàng hóa sau đó chuyển đổi mục đích sử dụng hoặc bán lại theo hình thức khác. Khai sai mã loại hình dẫn đến sai lệch chính sách thuế và báo cáo quyết toán hải quan.",
        "affected_target": "Doanh nghiệp FDI, doanh nghiệp chế xuất (DNCX), kho ngoại quan và các trung tâm phân phối logistics (Distribution Centre).",
        "risk_penalty": "Không thể hoàn thuế nhập khẩu, không nộp được báo cáo quyết toán nguyên phụ liệu năm, bị phạt vi phạm hành chính từ 2-5 triệu đồng/tờ khai và bị đưa vào diện nghi vấn gian lận mã loại hình.",
        "market_reaction": "Doanh nghiệp chế xuất gặp nhiều lúng túng khi nhập khẩu hàng hóa vừa dùng cho sản xuất vừa xuất bán nội địa hoặc tiêu hủy, các chi cục hải quan địa phương có cách hiểu không đồng nhất.",
        "pro_argument": "Chuẩn hóa mã loại hình theo bản chất giúp dữ liệu hải quan điện tử minh bạch, phục vụ phân tích rủi ro tự động bằng AI và ngăn chặn tình trạng lợi dụng quyền phân phối để trốn thuế.",
        "con_argument": "Các mô hình kinh doanh hiện đại rất linh hoạt (vừa sản xuất vừa thương mại vừa logistics); việc phân chia cứng nhắc theo mã loại hình truyền thống làm phát sinh thủ tục hủy tờ khai, sửa chữa phức tạp.",
        "intl_ref": "Hướng dẫn về Customs Procedure Codes (CPC) của Tổ chức Hải quan Thế giới (WCO) và Công ước Kyoto sửa đổi về thủ tục hải quan đơn giản hóa.",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Quyết định 1357/QĐ-TCHQ) & daily_notes.md (Công văn 21939/HQKV2-NVHQ hướng dẫn mã loại hình A11, A12, A41 ngày 23/09/2026).",
        "keywords": ["customs", "valuation"]
    }
]
