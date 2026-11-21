# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W41 DAY 4 — 09/10/2026)
Số lượng: 5 ý tưởng (STT 136 - 140)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

CASE_STUDY_W41_D4 = [
    {
        "stt": 136,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Lô Hàng 8 Container Thủy Sản Đi Busan Bị Bác e-C/O Form AK Do Lệch Mã Cảng Trên Hệ Thống Một Cửa Quốc Gia: Doanh Nghiệp Mất Hơn 650 Triệu Đồng Thuế Ưu Đãi",
        "stakeholders": "Công ty Cổ phần Thủy sản Biển Đông T. (Cà Mau) vs Hải quan Cảng Busan (Hàn Quốc) vs Phòng Quản lý XNK khu vực",
        "context": "Công ty T. xuất khẩu 8 container tôm đông lạnh trị giá 12 tỷ đồng sang Hàn Quốc theo hiệp định VKFTA/AKFTA. Doanh nghiệp nộp hồ sơ xin cấp e-C/O Form AK trên hệ thống EcoSys và được cấp mã số C/O điện tử truyền sang NSW. Tuy nhiên, trên hệ thống EcoSys nhân viên khai cảng dỡ hàng là 'Pusan Old Port' trong khi trên tờ khai hải quan điện tử và B/L là 'Busan New Port (PNIT)'. Khi hàng cập cảng Busan, hệ thống tự động của Hải quan Hàn Quốc từ chối mã e-C/O vì không khớp mã cảng đích giữa dữ liệu NSW Việt Nam và hải quan sở tại.",
        "damage": "Doanh nghiệp đối tác tại Hàn Quốc bị buộc tạm nộp thuế nhập khẩu MFN 20% thay vì 0% (tổng cộng 2.4 tỷ đồng bảo lãnh tạm nộp); phát sinh 650 triệu đồng chi phí phạt trễ hạn hợp đồng và phí lưu container lạnh suốt 12 ngày chờ xác minh văn bản cấp lại; công ty T. phải giảm giá 5% đơn hàng tiếp theo để giữ khách.",
        "root_cause": "Nhân viên chứng từ bất cẩn chọn sai mã cảng gợi ý trên phần mềm EcoSys; không kiểm tra đối soát 1:1 giữa bản truyền dữ liệu C/O điện tử và Tờ khai hải quan xuất khẩu trước khi bấm gửi lên NSW.",
        "solution": "1. Áp dụng quy tắc kiểm soát 'Cặp bài trùng dữ liệu': Bắt buộc copy chính xác mã UN/LOCODE của cảng bốc và cảng dỡ từ B/L và Tờ khai thông quan vào hồ sơ e-C/O; 2. Kiểm tra trạng thái 'Đã tiếp nhận và đối soát thành công' trên Cổng một cửa quốc gia NSW trước khi tàu cập cảng; 3. Nắm vững quy trình xin cấp C/O sửa đổi (Issued in substitution) qua cơ chế phân cấp địa phương trong vòng 24 giờ.",
        "opposing_views": "Phe A (Chủ hàng): Cảng Busan cũ hay mới đều thuộc cùng một cụm cảng thành phố Busan, lỗi khác biệt tên gọi kỹ thuật không làm thay đổi xuất xứ tôm Việt Nam; Phe B (Hải quan Hàn Quốc): Hệ thống kiểm soát rủi ro hải quan tự động chạy bằng thuật toán mã hóa số, sai lệch trường dữ liệu định danh bắt buộc phải từ chối ưu đãi để chống gian lận thương mại.",
        "lesson": "Trong kỷ nguyên hải quan số và C/O điện tử, 'chữ nghĩa là dữ liệu số'; sai một ký tự trên hệ thống NSW có sức tàn phá không kém gì làm mất toàn bộ bộ chứng từ gốc bằng giấy.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 09/10/2026 từ case study PH Logistics về phân cấp C/O & NSW) & Thư Viện Pháp Luật (Thông tư hướng dẫn quy tắc xuất xứ VKFTA/AKFTA).",
        "keywords": ["certificate of origin", "origin"]
    },
    {
        "stt": 137,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Kho Hàng Linh Kiện Nhập Khẩu 3.2 Tỷ Đồng Tại Bắc Ninh Bị Tạm Giữ Vì Chủ Doanh Nghiệp Nhầm Lẫn Giữa 'Tra Cứu Mã Số Thuế' Và 'Đăng Ký Địa Điểm Kinh Doanh'",
        "stakeholders": "Công ty TNHH Thương mại Công nghệ Đ. (Hà Nội) vs Đội Quản lý Thị trường & Chi cục Thuế khu vực Bắc Ninh",
        "context": "Công ty Đ. có trụ sở chính tại Cầu Giấy (Hà Nội), thuê thêm một kho hàng rộng 500m2 tại Từ Sơn (Bắc Ninh) để làm điểm tập kết linh kiện máy tính nhập khẩu chính ngạch từ cảng Hải Phòng. Khi thuê kho, kế toán chỉ vào trang web tra cứu mã số thuế gõ tên chi nhánh và nộp tờ khai thuế môn bài online mà không làm thủ tục nộp hồ sơ Đăng ký Địa điểm kinh doanh trên Cổng Dịch vụ công tỉnh Bắc Ninh. Khi Đội Quản lý thị trường phối hợp cơ quan thuế kiểm tra đột xuất, kho hàng không có Giấy chứng nhận đăng ký địa điểm kinh doanh và mã ĐĐKD trên hệ thống DVC vẫn ở trạng thái 'Chưa đăng ký'.",
        "damage": "Bị xử phạt vi phạm hành chính 15 triệu đồng về hành vi kinh doanh tại địa điểm không thông báo với cơ quan đăng ký kinh doanh; toàn bộ lô hàng linh kiện trị giá 3.2 tỷ đồng bị niêm phong tạm giữ suốt 18 ngày để đối soát tờ khai hải quan và hóa đơn; thiệt hại 180 triệu đồng tiền bồi thường giao hàng trễ cho 3 đối tác lắp ráp linh kiện tại KCN Yên Phong.",
        "root_cause": "Chủ doanh nghiệp và kế toán nhầm tưởng việc nộp thuế môn bài đồng nghĩa với việc đã hoàn thành đăng ký kinh doanh; thao tác sai trên cổng thông tin tra cứu thay vì nộp hồ sơ dịch vụ công mở mã ĐĐKD theo quy định.",
        "solution": "1. Trước khi nhập hàng vào bất kỳ kho bãi mới nào, bắt buộc phải có Giấy chứng nhận đăng ký địa điểm kinh doanh cấp bởi Sở KH&ĐT (có mã số dạng 0000X gắn liền MST công ty); 2. Kiểm tra trạng thái 'Đã duyệt và cấp mã số' trên Cổng Dịch vụ công quốc gia; 3. Nộp tờ khai lệ phí môn bài cho địa điểm kinh doanh mới (1 triệu đồng/năm) đúng hạn chậm nhất là ngày 30/01 năm sau.",
        "opposing_views": "Phe A (Chủ doanh nghiệp): Hàng hóa có đầy đủ tờ khai hải quan thông quan và hóa đơn VAT đầu vào thì kho chứa ở đâu cũng là tài sản hợp pháp của công ty; Phe B (Cơ quan chức năng): Luật quy định địa điểm kinh doanh là nơi diễn ra hoạt động lưu trữ buôn bán hàng hóa cụ thể, không đăng ký là vi phạm pháp luật về quản lý thương mại và tiềm ẩn rủi ro buôn lậu.",
        "lesson": "Hàng hóa chính ngạch nhưng chứa trong một cái kho 'chưa chính danh' thì vẫn có thể bị coi là hàng lậu; sự cẩu thả trong thủ tục hành chính ban đầu luôn phải trả giá bằng sự tê liệt của cả chuỗi cung ứng.",
        "url": "https://en.wikipedia.org/wiki/Taxpayer_Identification_Number",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 09/10/2026 từ case study AuthenticLemur1946) & Thư Viện Pháp Luật (Nghị định 122/2021/NĐ-CP, Nghị định 01/2021/NĐ-CP).",
        "keywords": ["taxpayer", "identification"]
    },
    {
        "stt": 138,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Tập Đoàn Sản Xuất Điện Tử Bị Đình Chỉ Tạm Thời Tư Cách 'Doanh Nghiệp Ưu Tiên AEO' Do Đối Tác Vận Tải Lén Lút Buôn Lậu Linh Kiện Trên Xe Chuyển Cửa Khẩu",
        "stakeholders": "Công ty TNHH Điện tử Toàn Cầu G. (Bắc Ninh - Doanh nghiệp AEO) vs Đội Kiểm soát Chống buôn lậu Tổng cục Hải quan vs Nhà xe vận chuyển đường bộ T.",
        "context": "Công ty G. là doanh nghiệp ưu tiên hải quan hàng đầu, được hưởng đặc quyền miễn kiểm tra thực tế hàng hóa và đưa container thẳng từ Cảng Đình Vũ về nhà máy Bắc Ninh mở niêm phong hải quan. Tuy nhiên, một tài xế của nhà xe vệ tinh T. được công ty G. thuê vận chuyển container kẹp chì đã thông đồng với đối tượng buôn lậu bên ngoài, dừng xe trên đường cao tốc Hà Nội - Hải Phòng, dùng kỹ thuật rút bản lề cửa container để nhét thêm 1.500 chiếc điện thoại thông minh lậu vào thùng hàng. Lực lượng kiểm soát chống buôn lậu Hải quan bắt quả tang vụ việc khi xe chuẩn bị vào cổng nhà máy.",
        "damage": "Công ty G. bị Tổng cục Hải quan ra quyết định đình chỉ chế độ Doanh nghiệp ưu tiên AEO trong 12 tháng; toàn bộ các lô hàng nguyên liệu nhập khẩu của công ty bị chuyển từ luồng xanh sang luồng đỏ kiểm hóa 100%; phát sinh chi phí nâng hạ kiểm hóa và chậm trễ dây chuyền lắp ráp ước tính hơn 1.8 triệu USD (hơn 44 tỷ đồng); cổ phiếu công ty mẹ trên sàn quốc tế sụt giảm 3.5%.",
        "root_cause": "Doanh nghiệp quá tin tưởng vào danh hiệu AEO của mình mà bỏ qua quy trình kiểm toán an ninh chuỗi cung ứng đối với nhà thầu vận tải phụ trợ (Third-party Logistics Security Audit) theo Tiêu chuẩn an ninh WCO SAFE; không lắp đặt hệ thống GPS hành trình giám sát độc lập đối với các xe chở hàng từ cảng về kho.",
        "solution": "1. Bắt buộc lắp đặt thiết bị e-Seal định vị vệ tinh GPS có cảnh báo mở cửa trái phép đối với 100% chuyến xe vận chuyển container về nhà máy; 2. Đưa điều khoản trách nhiệm bồi thường vô điều kiện và sa thải vĩnh viễn vào hợp đồng với các đơn vị vận tải nếu tài xế vi phạm an ninh hải quan; 3. Tổ chức đội thanh tra độc lập kiểm tra ngẫu nhiên lộ trình xe hàng trên đường.",
        "opposing_views": "Phe A (Doanh nghiệp AEO): Hành vi buôn lậu là do cá nhân tài xế nhà xe bên ngoài lén lút thực hiện, doanh nghiệp hoàn toàn không hay biết và không vụ lợi; Phe B (Cơ quan Hải quan): Bản chất của chế độ AEO là 'Doanh nghiệp chịu trách nhiệm toàn diện về tính an toàn của toàn bộ chuỗi cung ứng', nếu không kiểm soát được khâu vận tải nghĩa là lỗ hổng an ninh đã xuất hiện và bắt buộc phải thu hồi quyền ưu tiên.",
        "lesson": "'Đặc quyền lớn đi kèm trách nhiệm lớn'; một tấm thẻ Doanh nghiệp ưu tiên có thể mất hàng chục năm để xây dựng nhưng có thể sụp đổ chỉ sau một chuyến xe vận tải bị cài cắm hàng lậu.",
        "url": "https://en.wikipedia.org/wiki/Authorized_economic_operator",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Báo Hải quan Online & Thư Viện Pháp Luật (Thông tư 72/2015/TT-BTC về Doanh nghiệp ưu tiên).",
        "keywords": ["authorized", "economic", "operator"]
    },
    {
        "stt": 139,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Thất Lạc Bộ Vận Đơn Gốc 3 Bản Trên Đường Chuyển Phát Nhanh Sang Dubai: Doanh Nghiệp Bị Ngân Hàng Yêu Cầu Ký Quỹ 110% Trị Giá Lô Hàng Suốt 2 Năm Để Lấy Hàng",
        "stakeholders": "Công ty Cổ phần Nông sản Xuất khẩu V. (Đắk Lắk) vs Đơn vị Chuyển phát Nhanh Quốc tế vs Hãng tàu Biển M. vs Ngân hàng N.",
        "context": "Công ty V. xuất khẩu 5 container hạt tiêu đen sang Cảng Jebel Ali (Dubai) trị giá 450.000 USD (hơn 11 tỷ đồng) thanh toán bằng Thư tín dụng (L/C). Sau khi nhận bộ 3 bản gốc Vận đơn đường biển (Full set 3/3 Original B/L) từ hãng tàu M., ngân hàng xuất trình bộ chứng từ gửi qua bưu cục chuyển phát nhanh quốc tế sang ngân hàng phát hành tại Dubai. Tuy nhiên, gói bưu phẩm bị thất lạc trong quá trình trung chuyển tại sân bay Doha. Tàu đã cập cảng Jebel Ali nhưng không ai có bản gốc B/L để nhận hàng; hãng tàu M. từ chối cấp lại B/L và từ chối giao hàng nếu không có Thư bảo lãnh của ngân hàng (Bank Guarantee / Letter of Indemnity - LOI) kèm điều kiện ngân hàng phải cam kết chịu trách nhiệm vô điều kiện nếu có bên thứ ba xuất trình B/L gốc đòi hàng.",
        "damage": "Để phát hành Thư bảo lãnh ngân hàng LOI lấy hàng cho khách, ngân hàng yêu cầu Công ty V. phải ký quỹ phong tỏa bằng tiền mặt 110% giá trị lô hàng (495.000 USD - hơn 12.3 tỷ đồng) bị đóng băng vốn suốt 24 tháng (thời hiệu khởi kiện hàng hải); phát sinh 68.000 USD tiền phí lưu kho bãi demurrage tại Cảng Jebel Ali trong 25 ngày chờ giải quyết; công ty cạn kiệt dòng tiền lưu động suýt phá sản.",
        "root_cause": "Sử dụng vận đơn giấy truyền thống có tính chất chứng từ sở hữu (Document of Title) nhưng phụ thuộc vào rủi ro vận chuyển bưu chính; không áp dụng phương thức Vận đơn điện tử (e-B/L) hoặc phương thức làm Vận đơn giao hàng bằng điện (Telex Release / Surrendered B/L) ngay từ đầu.",
        "solution": "1. Đàm phán chuyển đổi sang sử dụng Vận đơn điện tử (e-B/L) trên các nền tảng số hóa được hãng tàu và ngân hàng chấp thuận; 2. Nếu bắt buộc dùng B/L giấy, chia bộ chứng từ gốc gửi làm 2 chuyến thư riêng biệt cách nhau 24 giờ (Tách 2/3 gửi trước, 1/3 gửi sau); 3. Yêu cầu bưu điện mua bảo hiểm trách nhiệm chứng từ có giá trị cao.",
        "opposing_views": "Phe A (Chủ hàng): Bưu điện làm mất thư thì bưu điện và hãng tàu phải phối hợp cấp lại cho doanh nghiệp nhận hàng, sao bắt phong tỏa cả chục tỷ đồng; Phe B (Hãng tàu & Ngân hàng): Bản gốc B/L là chứng từ có giá trị vô danh (Bearer instrument), bất kỳ ai cầm bản gốc đều có quyền đòi hàng hợp pháp, nếu hãng tàu giao hàng bừa bãi sẽ phải bồi thường gấp đôi nên bắt buộc phải có bảo lãnh ký quỹ 110%.",
        "lesson": "Một tờ giấy vận đơn gốc mang sức nặng bằng cả cơ đồ của doanh nghiệp; hiểu rõ bản chất pháp lý nguy hiểm của B/L giấy và chuyển dịch sang e-B/L là con đường sống còn để bảo vệ dòng tiền.",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cảnh báo rủi ro thất lạc B/L gốc của Hiệp hội Bảo trợ Chủ tàu quốc tế (The Standard Club P&I) & Bộ luật Hàng hải Việt Nam 2015.",
        "keywords": ["bill of lading", "carrier"]
    },
    {
        "stt": 140,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Doanh Nghiệp Gạo Miền Tây Tiết Kiệm Gần 1.2 Tỷ Đồng Chi Phí Vận Chuyển Mỗi Năm Nhờ Chuyển Đổi Tuyến Logistics Từ Xe Container Sang Tuyến Sà Lan Xanh Đi Cái Mép",
        "stakeholders": "Công ty Cổ phần Nông sản Xuất khẩu An Giang M. vs Đội Sà lan Container Tân Cảng vs Cụm Cảng Quốc tế Cái Mép (CMIT)",
        "context": "Công ty M. xuất khẩu bình quân 150 container gạo đóng bao 20 feet mỗi tháng từ nhà máy tại Thoại Sơn (An Giang) sang các thị trường Châu Âu và Châu Phi. Trước đây, doanh nghiệp ký hợp đồng với các đội xe đầu kéo đường bộ vận chuyển container từ An Giang chạy hơn 240km đường bộ lên Cảng Cát Lái và Cái Mép. Chi phí đường bộ đội lên cực cao do cước xe, phí cầu đường BOT và phí hạ tầng cảng biển đường bộ (500.000 VNĐ/cont 20ft). Sau khi làm việc với chuyên gia tư vấn giải pháp logistics Eureka, công ty chuyển đổi phương thức vận tải sang tập kết hàng tại Cảng Thốt Nốt (Cần Thơ) rồi đưa lên sà lan chuyên dụng chạy thẳng đường thủy sông Hậu ra Cụm cảng nước sâu Cái Mép.",
        "damage": "So sánh chi phí trước chuyển đổi: Doanh nghiệp từng chịu lãng phí hơn 100 triệu đồng mỗi tháng tiền chênh lệch cước vận tải đường bộ và phí hạ tầng cảng biển; ngoài ra từng có 3 lô hàng bị rớt tàu do xe container kẹt đường hơn 8 tiếng tại cửa ngõ TP.HCM vào giờ cao điểm, phát sinh 140 triệu đồng tiền phí lưu bãi và đổi chuyến tàu.",
        "root_cause": "Thói quen ngại thay đổi phương thức vận tải truyền thống của bộ phận mua hàng; chưa tận dụng lợi thế tự nhiên sông nước của Đồng bằng sông Cửu Long và chính sách miễn giảm phí hạ tầng cảng biển đường thủy nội địa của Nhà nước.",
        "solution": "1. Tái cấu trúc chuỗi cung ứng theo mô hình Hub-and-Spoke: Gom hàng tại kho nhà máy, chuyển ngắn ra cảng sông địa phương và đi bằng sà lan gom chuyến; 2. Tận dụng chính sách giảm 50% phí hạ tầng cảng biển đối với đường thủy nội địa; 3. Lập kế hoạch đóng hàng trước giờ tàu mẹ cập cảng 36-48 giờ để đảm bảo hành trình sà lan an toàn tuyệt đối.",
        "opposing_views": "Phe A (Bộ phận kinh doanh cũ): Vận chuyển đường bộ linh hoạt hơn, muốn đi lúc nào đi lúc đó; Phe B (Giám đốc Chuỗi cung ứng): Đường bộ đắt đỏ và tiềm ẩn rủi ro kẹt xe khó kiểm soát, trong khi sà lan có lịch trình cố định, chở khối lượng lớn cùng lúc và tiết kiệm 30% chi phí.",
        "lesson": "'Logistics xanh không phải là khẩu hiệu vì môi trường, mà là công cụ trực tiếp cắt giảm chi phí sản xuất'; doanh nghiệp biết tận dụng hạ tầng thủy nội địa sẽ nắm giữ lợi thế giá thành vượt trội trên thị trường xuất khẩu.",
        "url": "https://en.wikipedia.org/wiki/Barge",
        "series": "Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Báo Công Thương & Thư Viện Pháp Luật (Nghị quyết HĐND TP.HCM về thu phí hạ tầng cảng biển đường thủy).",
        "keywords": ["barge", "cargo"]
    }
]
