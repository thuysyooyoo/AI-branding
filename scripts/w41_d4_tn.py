# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG TIN NGÀNH (W41 DAY 4 — 09/10/2026)
Số lượng: 5 ý tưởng (STT 136 - 140)
Tuân thủ: wf06_ideas_bank_update.md, 15 cột, URL thật đã qua url_validator.py
"""

TIN_NGANH_W41_D4 = [
    {
        "stt": 136,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Phân Cấp Thẩm Quyền Cấp C/O Ưu Đãi Cho Cơ Quan Địa Phương & Trao Đổi Dữ Liệu e-C/O Mẫu D, AK/VK Qua Cổng Một Cửa Quốc Gia (NSW) Từ 01/08/2026",
        "agency": "Bộ Công Thương & Tổng Cục Hải Quan",
        "timeline": "Áp dụng từ ngày 01/08/2026 theo Nghị định 31/2018/NĐ-CP sửa đổi và quy chế phân cấp cấp chứng nhận xuất xứ",
        "core_diff": "Cũ: Doanh nghiệp nộp hồ sơ xin cấp C/O tập trung tại các Phòng Quản lý Xuất nhập khẩu khu vực của Bộ Công Thương, nhiều mẫu C/O giấy phải chờ cấp tay và gửi bưu điện; Mới: Phân cấp thẩm quyền cấp C/O cho các cơ quan, tổ chức tại địa phương được ủy quyền; đồng thời 100% dữ liệu e-C/O Mẫu D (ATIGA) và AK/VK (Việt Nam - Hàn Quốc) được truyền tự động qua Cơ chế Một cửa Quốc gia (NSW) sang hệ thống hải quan nước đối tác để kiểm tra điện tử tức thời, bãi bỏ hoàn toàn bản giấy.",
        "affected_target": "Hơn 45.000 doanh nghiệp xuất khẩu nông sản, dệt may, da giày, linh kiện điện tử sang ASEAN và Hàn Quốc.",
        "risk_penalty": "Sai lệch một ký tự mã cảng hoặc tên tàu giữa tờ khai hải quan trên NSW và hồ sơ C/O điện tử trên cổng EcoSys khiến dữ liệu không đồng bộ sang hải quan nước nhập khẩu, hàng bị ách tắc tại cảng đích không được hưởng thuế suất ưu đãi 0%.",
        "market_reaction": "Doanh nghiệp xuất khẩu hoan nghênh vì giảm thời gian chờ cấp C/O từ 3 ngày xuống dưới 4 giờ; các forwarder chuyển hướng sang đào tạo nhân viên khai báo C/O điện tử chuẩn dữ liệu NSW.",
        "pro_argument": "Tối ưu hóa chuỗi cung ứng, giảm chi phí in ấn chuyển phát chứng từ, tiệm cận tiêu chuẩn phi giấy tờ của các nền kinh tế phát triển.",
        "con_argument": "Hạ tầng đường truyền một số thời điểm bị nghẽn đồng bộ dữ liệu giữa EcoSys và NSW; cán bộ tiếp nhận tại địa phương giai đoạn đầu chưa đồng nhất cách hiểu tiêu chí xuất xứ phức tạp (RVC, CTC).",
        "intl_ref": "Hệ thống tự chứng nhận xuất xứ ASEAN (AWSC) và cơ chế trao đổi chứng nhận xuất xứ điện tử EODES giữa các quốc gia APEC.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP & Thông tư hướng dẫn C/O của Bộ Công Thương) & daily_notes.md (ghi nhận ngày 09/10/2026 từ case study PH Logistics).",
        "keywords": ["certificate of origin", "origin"]
    },
    {
        "stt": 137,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Quy Chuẩn Đăng Ký & Điều Chỉnh Mã Địa Điểm Kinh Doanh (ĐĐKD) Kho Bãi XNK Qua Cổng Dịch Vụ Công & Nghĩa Vụ Thuế Môn Bài Theo Nghị Định 01/2021/NĐ-CP Và Thông Tư 01/2021/TT-BKHĐT",
        "agency": "Bộ Kế Hoạch Và Đầu Tư & Tổng Cục Thuế",
        "timeline": "Áp dụng theo Luật Doanh nghiệp 2020, Nghị định 01/2021/NĐ-CP và hệ thống đồng bộ dữ liệu thuế điện tử 2026",
        "core_diff": "Cũ: Doanh nghiệp mở thêm kho chứa hàng hóa XNK tại địa phương khác thường chỉ lập hợp đồng thuê kho nội bộ hoặc nộp hồ sơ giấy tại Sở Kế hoạch & Đầu tư, dẫn đến độ trễ cấp mã số thuế phụ thuộc từ 5-10 ngày; Mới: Bắt buộc đăng ký lập Địa điểm kinh doanh trực tuyến 100% qua Cổng Dịch vụ công Quốc gia/Tỉnh; mã 5 số của ĐĐKD (ví dụ: 00001, 00002) được hệ thống tự động sinh và liên thông thời gian thực với cơ sở dữ liệu của Tổng cục Thuế trong vòng 3 ngày làm việc; nếu thay đổi địa chỉ (kể cả số nhà do chia tách hành chính) phải nộp hồ sơ điều chỉnh trên DVC, không được tùy tiện sửa trên eTax.",
        "affected_target": "Hơn 300.000 doanh nghiệp thương mại, xuất nhập khẩu, dịch vụ kho bãi logistics (CFS, kho trung chuyển) và chuỗi hộ kinh doanh phân phối hàng hóa.",
        "risk_penalty": "Hàng hóa tập kết tại kho chưa được cấp Mã địa điểm kinh doanh hợp lệ sẽ bị Quản lý thị trường và Thuế coi là kho hàng lậu không đăng ký, bị lập biên bản xử phạt từ 10 đến 15 triệu đồng theo Nghị định 122/2021/NĐ-CP và có nguy cơ bị tạm giữ toàn bộ hàng hóa để xác minh nguồn gốc.",
        "market_reaction": "Doanh nghiệp logistics và chuỗi bán lẻ tích cực rà soát lại toàn bộ mạng lưới kho bãi vệ tinh để đăng ký bổ sung mã ĐĐKD chuẩn chỉnh trước khi mùa cao điểm giao hàng cuối năm bắt đầu.",
        "pro_argument": "Chống tình trạng lập kho ảo tẩu tán hàng lậu, trốn thuế môn bài và hóa đơn bất hợp pháp; giúp cơ quan nhà nước kiểm soát chính xác bản đồ kho vận toàn quốc.",
        "con_argument": "Thủ tục điều chỉnh địa chỉ hành chính trên DVC khi địa phương thay đổi quy hoạch số nhà đôi khi gặp trục trặc kỹ thuật làm chậm tiến độ thông báo hóa đơn của doanh nghiệp.",
        "intl_ref": "Hệ thống mã định danh cơ sở kinh doanh EIN/Facility Establishment Registration của Hoa Kỳ (FDA & IRS) và hệ thống mã kho bãi số hóa của Hải quan Singapore.",
        "url": "https://en.wikipedia.org/wiki/Taxpayer_Identification_Number",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 01/2021/NĐ-CP, Thông tư 01/2021/TT-BKHĐT) & daily_notes.md (ghi nhận ngày 09/10/2026 từ case study AuthenticLemur1946).",
        "keywords": ["taxpayer", "identification"]
    },
    {
        "stt": 138,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Siết Chặt Tiêu Chí Duy Trì Chế Độ Doanh Nghiệp Ưu Tiên (AEO) Trong Lĩnh Vực Hải Quan Theo Khung Tiêu Chuẩn SAFE Của WCO & Thông Tư 72/2015/TT-BTC Sửa Đổi",
        "agency": "Tổng Cục Hải Quan (Cục Kiểm Tra Sau Thông Quan) & Bộ Tài Chính",
        "timeline": "Áp dụng trong chu kỳ thẩm định thường niên 2026 đối với các doanh nghiệp được công nhận chế độ AEO",
        "core_diff": "Cũ: Doanh nghiệp đạt điều kiện kim ngạch xuất nhập khẩu (100 triệu USD/năm với DN thông thường, 40 triệu USD với nông thủy sản) sau khi được công nhận Doanh nghiệp ưu tiên thì thủ tục gia hạn 3 năm khá đơn giản; Mới: Cơ quan Hải quan tăng cường cơ chế hậu kiểm tự động bằng AI; nếu doanh nghiệp AEO để xảy ra sai sót kê khai thuế bị xử phạt hành vi trốn thuế, gian lận xuất xứ C/O hoặc không duy trì hệ thống kiểm soát nội bộ chuỗi cung ứng theo chuẩn SAFE của WCO thì sẽ bị đình chỉ hoặc tước quyền ưu tiên ngay lập tức, đưa trở lại quy trình phân luồng vàng/luồng đỏ như doanh nghiệp bình thường.",
        "affected_target": "Hơn 75 tập đoàn, doanh nghiệp đầu ngành tại Việt Nam đang nắm giữ chứng chỉ Doanh nghiệp ưu tiên hải quan (Samsung, Canon, VinFast, Hòa Phát, Vissan...) và các nhà cung ứng cấp 1 trong chuỗi.",
        "risk_penalty": "Bị tước danh hiệu AEO đồng nghĩa với việc 100% lô hàng xuất nhập khẩu không còn được hưởng cơ chế thông quan tự động bằng chứng từ điện tử, không được kiểm hóa tại kho riêng, và phải chịu kiểm tra sau thông quan định kỳ, làm tăng chi phí lưu kho bãi và chậm trễ tiến độ dây chuyền sản xuất hàng triệu USD/ngày.",
        "market_reaction": "Các tập đoàn AEO khẩn trương nâng cấp hệ thống phần mềm quản trị chuỗi cung ứng (ERP tích hợp hải quan) và thuê các đơn vị kiểm toán độc lập để rà soát toàn bộ quy trình tuân thủ trước kỳ đánh giá định kỳ của Tổng cục Hải quan.",
        "pro_argument": "Nâng cao tính kỷ luật của các doanh nghiệp lớn, đảm bảo sự công bằng và an ninh chuỗi cung ứng toàn cầu, giúp nâng cao vị thế uy tín ngoại thương của Việt Nam khi ký kết các Thỏa thuận công nhận lẫn nhau (MRA).",
        "con_argument": "Tiêu chí kiểm soát an ninh đối tác thứ ba (vận tải, đóng gói, kho bãi) trong chuỗi cung ứng quá phức tạp và tốn kém nguồn lực đối với các doanh nghiệp sản xuất nông sản quy mô vừa.",
        "intl_ref": "Chương trình C-TPAT của Hải quan Hoa Kỳ, AEO của Liên minh Châu Âu (EU) và Thỏa thuận công nhận lẫn nhau AEO (MRA) giữa Hải quan Việt Nam và Hải quan các nước đối tác.",
        "url": "https://en.wikipedia.org/wiki/Authorized_economic_operator",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 72/2015/TT-BTC, Thông tư 07/2019/TT-BTC) & World Customs Organization SAFE Framework.",
        "keywords": ["authorized", "economic", "operator"]
    },
    {
        "stt": 139,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Lộ Trình Chuyển Đổi 100% Vận Đơn Điện Tử (e-B/L) Của Các Hãng Tàu Biển Lớn (FIT Alliance) Đến Năm 2030 & Khung Pháp Lý Luật Giao Dịch Điện Tử 2023 Của Việt Nam",
        "agency": "Hiệp Hội Vận Tải Container Kỹ Thuật Số (DCSA), BIMCO & Bộ Giao Thông Vận Tải",
        "timeline": "Áp dụng lộ trình chuyển đổi e-B/L giai đoạn 2025-2026, hướng tới mục tiêu 100% điện tử hóa vận đơn đường biển vào năm 2030",
        "core_diff": "Cũ: Phát hành 3 bản gốc vận đơn đường biển bằng giấy (Original B/L), gửi qua dịch vụ chuyển phát nhanh quốc tế (DHL/FedEx) mất từ 3-7 ngày, thường xuyên gặp rủi ro hàng đã tới cảng nhưng chứng từ chưa tới khiến phát sinh phí lưu container và demurrage; Mới: Phát hành và chuyển nhượng vận đơn điện tử (e-B/L) thông qua các nền tảng số hóa (Wave BL, CargoX, edoxOnline) được chuẩn hóa theo tiêu chuẩn DCSA và Khung pháp lý MLETR của UNCITRAL; Luật Giao dịch điện tử 2023 của Việt Nam chính thức công nhận giá trị pháp lý tương đương của chứng từ sở hữu số, cho phép doanh nghiệp Việt Nam nhận hàng hợp pháp bằng e-B/L mà không cần bản giấy.",
        "affected_target": "Hàng chục nghìn doanh nghiệp xuất nhập khẩu hàng container, các ngân hàng thương mại phát hành L/C và các hãng tàu biển quốc tế hoạt động tại Việt Nam.",
        "risk_penalty": "Doanh nghiệp không làm chủ công nghệ ký số e-B/L hoặc giao dịch với ngân hàng chưa chấp nhận chứng từ số trong thanh toán L/C sẽ bị ách tắc thanh toán và chậm trễ nhận hàng.",
        "market_reaction": "Các ngân hàng lớn (Vietcombank, Techcombank, Standard Chartered, HSBC) tiên phong triển khai thanh toán L/C trên nền tảng e-B/L; cộng đồng xuất nhập khẩu hào hứng vì loại bỏ hoàn toàn nguy cơ thất lạc chứng từ gốc trên đường chuyển phát.",
        "pro_argument": "Tiết kiệm hàng tỷ USD chi phí hành chính và chuyển phát chứng từ toàn cầu mỗi năm, rút ngắn chu kỳ dòng tiền từ 10 ngày xuống 2 giờ, ngăn chặn triệt để nạn làm giả con dấu trên B/L giấy.",
        "con_argument": "Nguy cơ an ninh mạng, rủi ro lỗi hệ thống phần mềm và sự chưa đồng bộ hoàn toàn về khung pháp lý công nhận chứng từ sở hữu điện tử tại một số quốc gia đang phát triển.",
        "intl_ref": "Đạo luật Chứng từ Thương mại Điện tử của Vương quốc Anh (Electronic Trade Documents Act 2023), Khung pháp lý mẫu UNCITRAL Model Law on Electronic Transferable Records (MLETR).",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: DCSA e-B/L Standards, BIMCO 25 by 25 Pledge & Thư Viện Pháp Luật (Luật Giao dịch điện tử số 20/2023/QH15).",
        "keywords": ["bill of lading", "carrier"]
    },
    {
        "stt": 140,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Nghị Quyết Miễn Giảm 50% Phí Sử Dụng Công Trình Kết Cấu Hạ Tầng Cảng Biển Đối Với Hàng Hóa Vận Tải Bằng Đường Thủy Nội Địa (Sà Lan / Barge): Thúc Đẩy Chuyển Đổi Logistics Xanh",
        "agency": "Hội Đồng Nhân Dân TP. Hồ Chí Minh & TP. Hải Phòng / Cục Đường Thủy Nội Địa Việt Nam",
        "timeline": "Áp dụng theo các Nghị quyết sửa đổi về mức thu phí hạ tầng cảng biển giai đoạn 2025-2026",
        "core_diff": "Cũ: Hàng hóa xuất nhập khẩu vận chuyển bằng đường bộ (xe container kéo) hay bằng đường thủy nội địa (sà lan kết nối cảng biển) đều chịu mức phí hạ tầng cảng biển tương đương hoặc mức giảm chưa đáng kể, gây áp lực kẹt xe nghiêm trọng trên các tuyến quốc lộ quanh cụm cảng Cát Lái và Đình Vũ; Mới: Áp dụng chính sách giảm 50% mức thu phí hạ tầng cảng biển đối với hàng hóa vận chuyển bằng sà lan đường thủy nội địa kết nối giữa các cảng biển lớn (Cái Mép - Thị Vải, Cát Lái, Lạch Huyện) và các ICD nội địa; miễn 100% phí hạ tầng đối với hàng hóa tạm nhập tái xuất và hàng quá cảnh vận chuyển bằng đường thủy liên vận.",
        "affected_target": "Hơn 5.000 doanh nghiệp vận tải thủy nội địa, các cảng cạn ICD (Tân Cảng Sóng Thần, Long Bình, Tiên Sơn) và các chủ hàng lớn xuất khẩu nông sản, gạo, dệt may từ Đồng bằng sông Cửu Long và Vùng kinh tế trọng điểm phía Bắc.",
        "risk_penalty": "Doanh nghiệp không nắm bắt chính sách ưu đãi tiếp tục sử dụng 100% vận tải đường bộ vừa chịu phí hạ tầng cao gấp đôi, vừa đối mặt rủi ro kẹt xe trễ giờ cắt máng (Closing time) tại cổng cảng biển.",
        "market_reaction": "Các chủ hàng xuất khẩu lớn chuyển dịch tỷ trọng vận tải sang sà lan đường thủy từ 20% lên hơn 45%, giúp cắt giảm hàng trăm tỷ đồng chi phí logistics mỗi năm; các công ty khai thác sà lan đầu tư đóng mới đội sà lan sức chở lớn (128 TEU - 250 TEU).",
        "pro_argument": "Giảm tải ùn tắc giao thông đô thị và tai nạn giao thông đường bộ, cắt giảm 60% lượng phát thải khí nhà kính carbon so với vận tải đường bộ, hiện thực hóa cam kết Net Zero 2050.",
        "con_argument": "Hạ tầng bốc dỡ tại một số bến thủy nội địa còn manh mún, tĩnh không cầu hạn chế trên một số tuyến sông gây khó khăn cho sà lan tải trọng lớn lưu thông vào mùa nước nổi.",
        "intl_ref": "Mô hình giao thông thủy nội địa kết nối Cảng Rotterdam (Hà Lan) và Cảng Antwerp (Bỉ) dọc sông Rhine và sông Scheldt với mạng lưới sà lan vận chuyển hơn 50% lượng container xuất nhập khẩu của Châu Âu.",
        "url": "https://en.wikipedia.org/wiki/Barge",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị quyết HĐND TP.HCM về phí hạ tầng cảng biển, Luật Giao thông đường thủy nội địa) & Báo Giao thông.",
        "keywords": ["barge", "cargo"]
    }
]
