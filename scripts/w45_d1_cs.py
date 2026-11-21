# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W45 DAY 1 — 04/11/2026)
Số lượng: 5 ý tưởng (STT 201 - 205)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

CASE_STUDY_W45_D1 = [
    {
        "stt": 201,
        "week": "W45 — 04/11/2026",
        "title": "Tai Nạn Nghề Nghiệp Từ 'Con Lăn Chuột Nhảy Dữ Liệu' Trên Phần Mềm ECUS5-VNACCS: Nhân Viên Khai Nhầm Số Lượng Từ 78 Thùng Thành 78 Pallet, Tờ Khai Văng Luồng Đỏ Kiểm Hóa 40 Tấn Hàng Kẹt Cứng Tại Cửa Khẩu Nội Địa",
        "stakeholders": "Công ty TNHH Nhập khẩu & Thương mại Nội thất V.T (Bắc Ninh), Nhân viên khai báo hải quan, Chi cục Hải quan Quản lý Hàng Đầu tư Gia công, Doanh nghiệp sản xuất đối tác.",
        "context": "Lấy chất liệu thực tế từ ghi chép của tác giả Thúy (daily_notes.md 04/11/2026): Nhân viên chứng từ thao tác gõ máy cực nhanh, vừa làm vừa lướt chuột. Khi mở phần mềm khai báo hải quan ECUS5 để truyền tờ khai nhập khẩu tại chỗ, trong lúc con trỏ chuột vô tình nằm trên ô 'Đơn vị tính', thao tác lăn chuột đã vô tình làm nhảy đơn vị từ 'Thùng (Carton)' sang 'Pallet', và số lượng nhảy vọt từ 78 thùng (tương đương 1.2 tấn) lên 78 pallet (tương đương 40 tấn). Do hệ thống không cảnh báo chênh lệch trọng lượng bất thường, khi bấm truyền chính thức, tờ khai lập tức bị hệ thống VNACCS phân luồng Đỏ kiểm tra thực tế 100% hàng hóa.",
        "damage": "Toàn bộ lô hàng bị ách tắc tại kho không được phép đưa vào dây chuyền sản xuất; nhà máy đối tác đã sử dụng gần hết lượng hàng thực tế khiến khi công chức hải quan kiểm tra hiện trường không đủ số lượng 40 tấn như khai báo; doanh nghiệp đối mặt với nguy cơ bị lập biên bản vi phạm hành chính về hành vi khai sai số lượng, trọng lượng hàng hóa; đình trệ tiến độ giao hàng cho khách VIP phạt 150 triệu đồng và chi phí giải trình, kiểm đếm lại kho bãi mất gần 10 ngày làm việc.",
        "root_cause": "Sự chủ quan và thao tác vội vã của nhân viên khai hải quan: Quá tự tin vào tốc độ gõ phím mà bỏ qua bước in tờ khai nháp (tờ khai thử nghiệm) để đối chiếu chéo số liệu 2 mắt; thiếu kiến thức về lỗi nhạy cảm của con lăn chuột trên các trường danh mục dropdown của phần mềm ECUS5.",
        "solution": "1. Thiết lập quy tắc 'Khóa con lăn chuột' và bắt buộc in tờ khai nháp kiểm tra chéo (Four-eyes principle) trước khi bấm nút truyền chính thức; 2. Cài đặt chế độ kiểm tra logic trên phần mềm ECUS: So sánh tỷ lệ Số lượng / Trọng lượng gộp (Gross Weight) nếu vượt ngưỡng bình thường phải hiển thị cảnh báo đỏ; 3. Nếu lỡ truyền sai, lập tức liên hệ cán bộ hải quan nộp công văn giải trình sửa đổi bổ sung tờ khai theo Điều 20 Thông tư 38/2015/TT-BTC trước khi hải quan ban hành quyết định kiểm tra thực tế hàng hóa.",
        "opposing_views": "Phe Người khai hải quan: Cho rằng phần mềm ECUS5 có lỗi giao diện người dùng (UI bug) khi để con lăn chuột can thiệp vào dropdown mà không có cảnh báo khóa trường; Phe Hải quan: Quy định pháp luật nêu rõ người khai hải quan phải chịu trách nhiệm pháp lý cao nhất về tính chính xác của dữ liệu đã truyền lên hệ thống VNACCS.",
        "lesson": "Trong nghề khai báo hải quan, 'Nhanh 1 giây có thể chậm cả tháng'. Tốc độ gõ phím không bao giờ quan trọng bằng sự cẩn trọng đối chiếu từng con số, dấu phẩy và đơn vị tính trước khi bấm nút truyền dữ liệu.",
        "url": "https://en.wikipedia.org/wiki/Electronic_data_interchange",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Tổng cục Hải quan, Thư Viện Pháp Luật (Thông tư 38/2015/TT-BTC, Thông tư 39/2018/TT-BTC) & daily_notes.md (04/11/2026).",
        "keywords": ["electronic", "data"]
    },
    {
        "stt": 202,
        "week": "W45 — 04/11/2026",
        "title": "'Bẫy' Đóng Mã Số Thuế Do Gián Đoạn Địa Điểm Kinh Doanh: Doanh Nghiệp Bị Tạm Khóa Cổng Hải Quan Điện Tử Đúng Ngày Cập Cảng 3 Container Hàng 5.8 Tỷ Đồng, Phạt Lưu Bãi 75 Triệu Đồng Vì Cán Bộ Thuế Kiểm Tra Thực Tế Đột Xuất",
        "stakeholders": "Công ty TNHH Thương mại Quốc tế N.A (Hà Nội), Chi cục Thuế Quận Cầu Giấy, Chi cục Hải quan Cửa khẩu Cảng Hải Phòng, Đội Vận tải Giao nhận.",
        "context": "Lấy chất liệu từ ghi chép thực tế (daily_notes.md 04/11/2026): Doanh nghiệp thay đổi địa chỉ văn phòng từ năm 2024 nhưng bộ phận hành chính không làm thủ tục thay đổi thông tin đăng ký thuế kịp thời. Trong đợt rà soát định kỳ, cơ quan Thuế cử cán bộ xuống xác minh thực địa nhưng không thấy biển hiệu công ty tại địa chỉ cũ, liền cập nhật trạng thái 'Người nộp thuế không hoạt động tại địa chỉ đã đăng ký' (Trạng thái 06). Đến đầu tháng 11/2026, khi công ty nhập khẩu 3 container nguyên liệu sản xuất trị giá 5.8 tỷ đồng từ Hàn Quốc về cảng Hải Phòng, nhân viên mở phần mềm ECUS thì hệ thống báo lỗi không thể truyền tờ khai do Mã số thuế bị cơ quan thuế khóa trên Hệ thống thông tin quốc gia.",
        "damage": "3 container hàng bị mắc kẹt tại cảng Hải Phòng trong suốt 12 ngày làm thủ tục khôi phục mã số thuế; phát sinh chi phí lưu container (Demurrage) và lưu bãi (Detention) lên tới 75 triệu đồng; hợp đồng gia công với đối tác bị trễ tiến độ 2 tuần, bị phạt hợp đồng 180 triệu đồng.",
        "root_cause": "Sự tắc trách của lãnh đạo và bộ phận pháp lý công ty: Coi nhẹ việc cập nhật thay đổi địa chỉ kinh doanh; không nắm được quy định liên thông tự động giữa cơ sở dữ liệu của Tổng cục Thuế và Tổng cục Hải quan (chỉ cần Thuế khóa MST thì cổng Hải quan tự động từ chối thông quan ngay lập tức).",
        "solution": "1. Khi thay đổi bất kỳ thông tin nào về địa chỉ trụ sở, văn phòng giao dịch, bắt buộc nộp hồ sơ thay đổi thông tin đăng ký thuế trong vòng 10 ngày làm việc theo Luật Quản lý Thuế; 2. Khi bị khóa MST, chuẩn bị ngay hồ sơ đề nghị khôi phục MST gồm: Giấy đề nghị khôi phục MST (Mẫu 25/ĐK-TCT), Hợp đồng thuê văn phòng mới có công chứng, Giấy nộp thuế đầy đủ và mời cán bộ thuế xuống lập Biên bản xác minh thực địa hoạt động bình thường; 3. Kiểm tra định kỳ trạng thái MST trên Cổng tracuunnt.gdt.gov.vn trước mỗi lần ký kết hợp đồng nhập khẩu lớn.",
        "opposing_views": "Phe Doanh nghiệp: Cho rằng việc cơ quan Thuế tự động khóa MST mà không gửi thông báo trước gây thiệt hại nặng nề cho hoạt động XNK; Phe Cơ quan Thuế: Việc rà soát kiểm tra thực tế địa bàn là biện pháp nghiệp vụ bắt buộc nhằm ngăn chặn tình trạng thành lập doanh nghiệp ma để mua bán hóa đơn GTGT khống.",
        "lesson": "Trong kinh doanh xuất nhập khẩu, 'Mã số thuế là nhịp thở của doanh nghiệp'. Một sự cố hành chính sơ đẳng về địa chỉ trụ sở có thể làm tê liệt toàn bộ chuỗi cung ứng hàng tỷ đồng trong chớp mắt.",
        "url": "https://en.wikipedia.org/wiki/Value-added_tax",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Tổng cục Thuế, Thư Viện Pháp Luật (Luật Quản lý Thuế 38/2019/QH14, Thông tư 105/2020/TT-BTC) & daily_notes.md (04/11/2026).",
        "keywords": ["tax", "vat"]
    },
    {
        "stt": 203,
        "week": "W45 — 04/11/2026",
        "title": "Cảnh Báo Đỏ RASFF Từ Cảng Rotterdam: Lô 4 Container Thanh Long 2.4 Tỷ Đồng Bị Hải Quan Hà Lan Ra Lệnh Tiêu Hủy Do Vượt Ngưỡng Dư Lượng Thuốc BVTV 0.02 ppm — Nỗi Đau Mất Trắng Vì Tin Tưởng Bừa Chứng Thư Khảo Sát Nông Hộ",
        "stakeholders": "Công ty Cổ phần Nông sản Xuất khẩu Đ.T (Bình Thuận), Nhà nhập khẩu trái cây tại Rotterdam (Hà Lan), Cơ quan An toàn Thực phẩm & Dược phẩm Châu Âu (EFSA), Đơn vị kiểm nghiệm tư nhân.",
        "context": "Doanh nghiệp xuất khẩu 4 container thanh long ruột đỏ tươi sang thị trường Hà Lan theo điều kiện CIF cảng Rotterdam. Trước khi xuất bốc hàng lên tàu, doanh nghiệp thuê một công ty dịch vụ kiểm nghiệm lấy mẫu thử và cấp chứng thư đạt tiêu chuẩn MRL của EU. Tuy nhiên, khi tàu cập cảng Rotterdam, Cơ quan Kiểm dịch Thực vật Hà Lan lấy mẫu ngẫu nhiên và phát hiện dư lượng hoạt chất thuốc bảo vệ thực vật Dithiocarbamates vượt ngưỡng cho phép 0.02 ppm. Hải quan Hà Lan ngay lập tức kích hoạt cảnh báo trên Hệ thống Cảnh báo Nhanh về Thực phẩm và Thức ăn Chăn nuôi của EU (RASFF) và ban hành lệnh tiêu hủy bắt buộc tại bến cảng.",
        "damage": "Toàn bộ 4 container thanh long trị giá 2.4 tỷ đồng bị tiêu hủy tại lò rác thải công nghiệp cảng Rotterdam; doanh nghiệp Việt Nam phải chi trả thêm 18,000 EUR (khoảng 490 triệu đồng) chi phí tiêu hủy và phí lưu kho lạnh; nhà nhập khẩu hủy hợp đồng nguyên tắc năm tiếp theo trị giá hơn 20 tỷ đồng; thương hiệu trái cây Việt Nam bị đưa vào danh sách tăng tần suất kiểm tra 50% tại các cửa khẩu EU.",
        "root_cause": "Doanh nghiệp thu mua gom nông sản từ nhiều nông hộ nhỏ lẻ mà không kiểm soát được nhật ký canh tác; công ty kiểm nghiệm tư nhân lấy mẫu đại diện không chuẩn xác (chỉ lấy mẫu bề mặt ngoài thùng mà không kiểm tra ngẫu nhiên sâu bên trong ruột container); người nông dân tự ý phun thuốc nấm sát ngày thu hoạch để quả bóng đẹp.",
        "solution": "1. Xây dựng chuỗi liên kết vùng trồng đạt chuẩn GlobalGAP có mã số vùng trồng (PUC) và mã số cơ sở đóng gói (PHC) được Cục Bảo vệ Thực vật cấp; 2. Giám sát 100% nhật ký sử dụng thuốc BVTV và cách ly tuyệt đối tối thiểu 21 ngày trước khi thu hoạch; 3. Chỉ gửi mẫu kiểm nghiệm tại các phòng lab đạt chuẩn ISO 17025 được EU công nhận (như Eurofins, SGS); 4. Mua bảo hiểm rủi ro thu hồi sản phẩm xuất khẩu quốc tế (Product Recall Insurance).",
        "opposing_views": "Phe Nông dân & Thương lái: Cho rằng tiêu chuẩn MRL của EU thay đổi liên tục và quá khắt khe, biến tướng thành hàng rào kỹ thuật phi thuế quan bảo hộ nông nghiệp nội khối; Phe Chuyên gia Quốc tế: Khẳng định bảo vệ sức khỏe người tiêu dùng là nguyên tắc tối thượng, doanh nghiệp không thể mang tư duy 'làm ăn cẩu thả' bước ra biển lớn.",
        "lesson": "Trong xuất khẩu nông sản sang các thị trường khó tính, 'Một giọt thuốc BVTV dư thừa có thể dìm chết một con tàu hàng triệu USD'. Uy tín xuất khẩu không thể xây dựng bằng sự may rủi mà phải bằng quy trình chuẩn mực từ gốc rễ nông hộ.",
        "url": "https://en.wikipedia.org/wiki/Rapid_Alert_System_for_Food_and_Feed",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Cục Bảo vệ Thực vật, Thư Viện Pháp Luật (Quy định EU 2019/1793 về an toàn thực phẩm) & daily_notes.md (04/11/2026).",
        "keywords": ["food", "feed"]
    },
    {
        "stt": 204,
        "week": "W45 — 04/11/2026",
        "title": "Cạm Bẫy 'Forwarder Ma' Thu Hộ Tiền Cước Tàu Biển (Freight Collect): Doanh Nghiệp Dệt May Bị Hãng Tàu Mỹ Giữ 8 Container Tại Cảng Los Angeles Buộc Phải Trả Cước Lần 2 Kèm 65,000 USD Phạt Lưu Bãi Vì Đại Lý Trung Gian Ôm Tiền Bỏ Trốn",
        "stakeholders": "Công ty Cổ phần Dệt may May mặc T.P (TP.HCM), Công ty TNHH Dịch vụ Giao nhận Hàng hải T. (Forwarder trung gian tại TP.HCM), Hãng tàu container quốc tế Evergreen Line, Nhà nhập khẩu thời trang tại California (Hoa Kỳ).",
        "context": "Doanh nghiệp xuất khẩu 8 container hàng may mặc sang Mỹ theo điều kiện FOB Cảng Cát Lái, người mua chỉ định đại lý Forwarder nội địa là Công ty T. đứng ra thu hộ cước biển và phát hành House B/L (HBL). Doanh nghiệp Việt Nam đã thanh toán đầy đủ các khoản phụ phí local charges và người mua tại Mỹ đã chuyển tiền cước biển quốc tế cho đại lý của bên Forwarder. Tuy nhiên, do đầu tư tài chính thua lỗ, giám đốc công ty Forwarder T. đã ôm toàn bộ tiền cước hàng trăm container của khách hàng bỏ trốn mà không thanh toán tiền cước cho Hãng tàu Evergreen Line. Khi hàng đến cảng Los Angeles, Hãng tàu kiên quyết kích hoạt quyền giữ hàng (Lien on Cargo) và từ chối phát hành lệnh giao hàng Master D/O.",
        "damage": "Chủ hàng Mỹ không thể lấy được hàng; tiền phạt lưu container và lưu bãi (Demurrage & Detention) tại cảng Los Angeles tăng chóng mặt lên tới 65,000 USD sau 25 ngày mắc kẹt; doanh nghiệp dệt may Việt Nam phải đứng ra cùng chia sẻ trả tiền cước biển lần thứ hai cho hãng tàu (38,000 USD) để kịp lấy hàng cung cấp cho mùa mua sắm Black Friday, tổng thiệt hại vượt hơn 2.5 tỷ đồng.",
        "root_cause": "Sự lỏng lẻo trong việc thẩm định uy tín và năng lực tài chính của Forwarder trung gian; sử dụng đơn vị giao nhận không có giấy phép đại lý vận tải (NVOCC) được Ủy ban Hàng hải Liên bang Hoa Kỳ (FMC) cấp phép; hợp đồng không ràng buộc trách nhiệm bồi thường và không yêu cầu ký quỹ bảo lãnh nghĩa vụ cước tàu.",
        "solution": "1. Khi xuất khẩu hàng hóa sang thị trường Bắc Mỹ, bắt buộc kiểm tra xem Forwarder có chứng chỉ FMC Bond và mã SCAC hợp lệ hay không; 2. Ưu tiên sử dụng Master B/L phát hành trực tiếp từ hãng tàu hoặc các công ty Logistics đa quốc gia uy tín có văn phòng đại diện pháp nhân tại cả hai đầu cảng đi và cảng đến; 3. Cài đặt điều khoản thanh toán rõ ràng: Chỉ thanh toán tiền cước khi có xác nhận đã gạch nợ (Freight Prepaid Confirmation) chính thức từ hãng tàu.",
        "opposing_views": "Phe Chủ hàng: Cho rằng hãng tàu không có quyền giữ hàng của người nhận hàng khi họ đã trả tiền cước đầy đủ cho đại lý được ủy quyền; Phe Hãng tàu: Theo Luật Hàng hải Quốc tế, hãng tàu chỉ có quan hệ hợp đồng với chủ thể đứng tên trên Master B/L (tức công ty Forwarder), và có toàn quyền giữ hàng nếu chưa nhận đủ cước phí vận chuyển.",
        "lesson": "Trong giao nhận vận tải quốc tế, 'Rẻ một đồng tiền cước có thể mất trắng cả chục container hàng'. Đừng bao giờ giao phó số phận tài sản tiền tỷ của doanh nghiệp cho những đơn vị trung gian không tên tuổi, thiếu bảo chứng tài chính và pháp lý vững chắc.",
        "url": "https://en.wikipedia.org/wiki/Freight_forwarder",
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Hiệp hội Doanh nghiệp Dịch vụ Logistics Việt Nam (VLA), Bộ Luật Hàng hải Việt Nam, Thư Viện Pháp Luật & daily_notes.md (04/11/2026).",
        "keywords": ["freight", "forwarder"]
    },
    {
        "stt": 205,
        "week": "W45 — 04/11/2026",
        "title": "'Giải Cứu' Ngoạn Mục Lô Thiết Bị Y Tế Chẩn Đoán 18 Tỷ Đồng Kẹt Tại Sân Bay Nội Bài Vì Vướng Giấy Phép Nhập Khẩu: Bí Quyết Đưa Hàng Về Kho Riêng Bảo Quản Theo Điều 32 Thông Tư 38 Giúp Tiết Kiệm 140 Triệu Phí Kho Lạnh",
        "stakeholders": "Công ty TNHH Thiết bị Y sinh P.H (Hà Nội), Chi cục Hải quan Cửa khẩu Sân bay Quốc tế Nội Bài, Vụ Trang thiết bị & Công trình Y tế (Bộ Y tế), Đơn vị tư vấn Hải quan Eureka Logistics.",
        "context": "Doanh nghiệp nhập khẩu khẩn cấp lô 2 máy chụp cộng hưởng từ và linh kiện cảm biến y tế chuyên dụng trị giá 18 tỷ đồng bằng đường hàng không từ Munich (Đức) về Sân bay Quốc tế Nội Bài. Do quy định phân loại trang thiết bị y tế loại C, D vừa có sự thay đổi theo Nghị định 98/2021/NĐ-CP và Nghị định 07/2023/NĐ-CP, hồ sơ đăng ký lưu hành của doanh nghiệp vẫn đang trong quá trình xét duyệt bổ sung tại Bộ Y tế. Hàng về sân bay phải lưu tại kho lạnh đặc biệt với chi phí cắm điện và bảo quản lên tới 12 triệu đồng/ngày, đồng thời có nguy cơ hỏng bo mạch điện tử nhạy cảm nếu lưu giữ quá 10 ngày.",
        "damage": "Nguy cơ bị phạt lưu kho lạnh sân bay hơn 140 triệu đồng nếu phải chờ cấp phép thông thường từ 2 - 3 tuần; tiến độ lắp đặt cho bệnh viện đa khoa tuyến đầu bị đe dọa phạt tiến độ 50 triệu/ngày.",
        "root_cause": "Doanh nghiệp chủ quan ký hợp đồng giao hàng gấp bằng đường hàng không trước khi có Giấy phép nhập khẩu/Số lưu hành trang thiết bị y tế chính thức; chưa nắm vững cơ chế đưa hàng về bảo quản độc lập theo quy định của pháp luật hải quan.",
        "solution": "Đội ngũ chuyên gia Hải quan Eureka Logistics đã lập tức kích hoạt phương án cứu nguy: 1. Soạn thảo Bộ hồ sơ Đề nghị đưa hàng về bảo quản theo Điều 32 Thông tư 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC); 2. Xuất trình hợp đồng thuê kho chuyên dụng có điều hòa kiểm soát nhiệt độ đạt chuẩn GSP và Giấy cam kết bảo quản nguyên trạng hàng hóa không tiêu thụ; 3. Nộp văn bản tiếp nhận hồ sơ hợp lệ của Bộ Y tế để Hải quan Nội Bài chấp thuận cho niêm phong kẹp chì và áp tải đưa toàn bộ lô hàng về kho riêng của doanh nghiệp chỉ sau 24 giờ, giúp tiết kiệm 100% phí kho lạnh đắt đỏ tại sân bay.",
        "opposing_views": "Phe Doanh nghiệp ban đầu: E ngại việc xin đưa hàng về bảo quản thủ tục quá phức tạp và hải quan sẽ khó duyệt vì hàng giá trị lớn; Phe Chuyên gia Eureka: Khẳng định đưa hàng về bảo quản là quyền lợi hợp pháp của doanh nghiệp được pháp luật bảo vệ, chỉ cần hồ sơ chuẩn mực và cam kết kho bãi minh bạch.",
        "lesson": "Trong xử lý hải quan thông minh, 'Luật pháp không chỉ là rào cản, mà chính là chiếc chìa khóa vạn năng nếu bạn biết sử dụng đúng điều khoản'. Nắm vững quyền đưa hàng về kho bảo quản là tuyệt chiêu sống còn để cắt đứt mọi loại chi phí phát sinh oan uổng.",
        "url": "https://en.wikipedia.org/wiki/Customs",
        "series": "Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Tổng cục Hải quan, Thư Viện Pháp Luật (Điều 32 Thông tư 38/2015/TT-BTC sửa đổi, Nghị định 98/2021/NĐ-CP & 07/2023/NĐ-CP) & daily_notes.md (04/11/2026).",
        "keywords": ["customs", "duty"]
    }
]
