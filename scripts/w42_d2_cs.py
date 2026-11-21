# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W42 DAY 2 — 15/10/2026)
Số lượng: 5 ý tưởng (STT 151 - 155)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

CASE_STUDY_W42_D2 = [
    {
        "stt": 151,
        "week": "W42 — 15/10/2026",
        "title": "Lô Hàng 800 Bộ Switch & Router Văn Phòng Bị Hải Quan Sân Bay Nội Bài Giữ 12 Ngày Vì Không Khai Mã 'KO': Phạt 35 Triệu Đồng Suýt Bị Khép Lỗi Thiếu Giấy Phép Mật Mã Dân Sự",
        "stakeholders": "Công ty TNHH Công nghệ Viễn thông T. (Hà Nội), Chi cục Hải quan Cửa khẩu Sân bay Quốc tế Nội Bài, Cục Quản lý Mật mã Dân sự & Kiểm định Sản phẩm Mật mã (Ban Cơ yếu Chính phủ).",
        "context": "Công ty T. nhập khẩu lô hàng 800 bộ thiết bị phát wifi công nghiệp và switch mạng văn phòng từ Đài Loan, trị giá lô hàng 1.8 tỷ đồng. Mặt hàng có mã HS 8517.62.99. Trong danh mục quản lý mật mã dân sự theo Nghị định 58/2016/NĐ-CP và Nghị định 53/2018/NĐ-CP, mã HS 8517.62.99 thuộc diện phải có Giấy phép xuất nhập khẩu sản phẩm mật mã dân sự. Nhân viên chứng từ của công ty T. không biết đến hướng dẫn tại Công văn 22233/CHQ-GSQL nên mở tờ khai hải quan thông thường mà không khai mã nghiệp vụ 'KO' để xác nhận đặc tính kỹ thuật không có chức năng mã hóa bảo mật chuyên dụng. Hệ thống VNACCS lập tức cảnh báo luồng đỏ, Hải quan cửa khẩu đình chỉ thông quan và yêu cầu xuất trình Giấy phép của Ban Cơ yếu.",
        "damage": "Lô hàng bị lưu giữ tại kho hàng sân bay 12 ngày; Phát sinh chi phí lưu kho 28 triệu đồng; Bị phạt vi phạm hành chính 35 triệu đồng theo Nghị định 128/2020/NĐ-CP; Bị đối tác dự án phạt chậm tiến độ lắp đặt hệ thống mạng 60 triệu đồng.",
        "root_cause": "Doanh nghiệp không cập nhật văn bản nghiệp vụ mới (Công văn 22233/CHQ-GSQL); Không hiểu quy tắc phân định giữa sản phẩm mật mã dân sự chuyên dụng và thiết bị viễn thông dân dụng thông thường có mã HS trùng nhau; Thiếu chuẩn bị tài liệu kỹ thuật (Datasheet/Catalogue) mô tả thuật toán mã hóa khi mở tờ khai.",
        "solution": "1. Trước khi nhập khẩu thiết bị viễn thông, tra cứu mã HS đối chiếu danh mục mật mã dân sự và an ninh mạng; 2. Kiểm tra Datasheet: Nếu thiết bị chỉ dùng chuẩn bảo mật wifi thông thường (WPA2/WPA3 dân dụng) không thuộc diện quản lý chuyên ngành, BẮT BUỘC khai mã nghiệp vụ 'KO' tại ô chỉ tiêu thông tin quy định; 3. Chuẩn bị sẵn Bảng giải trình kỹ thuật và cam kết chịu trách nhiệm trước pháp luật để nộp kèm tờ khai.",
        "opposing_views": "Phe Doanh nghiệp: Thiết bị wifi văn phòng bán đầy ở siêu thị điện máy mà hải quan cứ thấy trùng mã HS là đòi giấy phép con của Ban Cơ yếu làm tắc nghẽn kinh doanh; Phe Cơ quan Quản lý: Không gian mạng tiềm ẩn nguy cơ gián điệp và tấn công mạng, nếu không kiểm soát chặt các thiết bị truyền dẫn có mã hóa thì an ninh quốc gia bị đe dọa nghiêm trọng.",
        "lesson": "Trong thời đại công nghệ số, một chiếc switch hay router thông thường cũng có thể trở thành 'hàng hóa nhạy cảm' nếu trùng mã HS với danh mục an ninh mạng. Nắm vững mã nghiệp vụ hải quan (như mã KO) chính là chiếc chìa khóa vàng để tự giải cứu hàng hóa mà không cần chạy vạy giấy phép con.",
        "url": "https://en.wikipedia.org/wiki/Cryptography",
        "series": "Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 58/2016/NĐ-CP, Công văn 22233/CHQ-GSQL, Công văn 19916/CHQ-GSQL) & daily_notes.md (ghi nhận ngày 15/10/2026 từ case study Thợ Khai HQ).",
        "keywords": ["cryptography", "security"]
    },
    {
        "stt": 152,
        "week": "W42 — 15/10/2026",
        "title": "Đoàn 4 Sà Lan Chở Hàng Quá Cảnh Đi Campuchia Bị Đình Chỉ Tại Cảng Sài Gòn: Quên Gắn Thiết Bị Giám Sát GPS Cho Container Xếp Chồng Theo Thông Tư 121/2025/TT-BTC, Thiệt Hại 185 Triệu Đồng",
        "stakeholders": "Công ty Cổ phần Vận tải Thủy Quốc tế M. (Long An), Chi cục Hải quan Cảng Sài Gòn Khu vực 3, Hãng tàu biển quốc tế Evergreen.",
        "context": "Công ty M. nhận vận chuyển quá cảnh 32 container hàng tiêu dùng và vải cuộn từ tàu mẹ cập cảng Cát Lái, bốc sang 4 sà lan đường thủy nội địa để vận chuyển theo tuyến sông Tiền sang cảng Phnom Penh (Campuchia). Do các container được xếp chồng 2 tầng trên sà lan để tối ưu tải trọng, theo quy định mới tại khoản 22, 23 Điều 1 Thông tư 121/2025/TT-BTC (sửa đổi Điều 51a, 51b Thông tư 38/2015/TT-BTC), đối với container xếp chồng đường thủy nội địa bắt buộc phải gắn thiết bị giám sát hành trình (e-Seal GPS) có kết nối truyền dữ liệu trực tiếp với hệ thống giám sát của Tổng cục Hải quan. Tuy nhiên, đơn vị vận tải chủ quan chỉ bấm chì cơ khí hải quan thông thường như trước đây và nộp Bản kê quá cảnh cũ (không dùng Mẫu 09 chuẩn mới). Khi sà lan chuẩn bị nhổ neo, Đội Kiểm soát Hải quan đã kiểm tra đột xuất và đình chỉ xuất bến toàn bộ đoàn sà lan.",
        "damage": "4 sà lan bị giữ tại phao neo cảng 5 ngày chờ bổ sung thiết bị định vị và làm lại thủ tục điện tử; Phát sinh chi phí neo sà lan và lưu bãi 110 triệu đồng; Hãng tàu phạt chậm trả rỗng container 75 triệu đồng (tổng thiệt hại 185 triệu đồng); Khách hàng Campuchia khiếu nại đứt gãy chuyền may mặc tại Phnom Penh.",
        "root_cause": "Đơn vị vận tải đường thủy nội địa không nghiên cứu kỹ Thông tư 121/2025/TT-BTC; Nghĩ rằng thủ tục đường thủy đơn giản hơn đường bộ; Không đầu tư đồng bộ thiết bị e-Seal GPS chuyên dụng cho container tầng trên và tầng dưới trên sà lan.",
        "solution": "1. Trước khi lập phương án xếp dỡ container quá cảnh lên phương tiện thủy, phải kiểm tra danh mục thiết bị định vị GPS đạt chuẩn tương thích với phần mềm quản lý của Hải quan; 2. Áp dụng chuẩn Mẫu Bản kê hàng hóa quá cảnh số 09 theo Thông tư 121/2025/TT-BTC; 3. Cắt cử nhân sự kỹ thuật kiểm tra dung lượng pin và tín hiệu phát sóng của từng e-Seal GPS trước khi trình hải quan ký phê duyệt xuất bến.",
        "opposing_views": "Phe Vận tải thủy: Sà lan chạy trên tuyến sông độc đạo có trạm kiểm soát biên phòng liên tục, việc bắt buộc mỗi container xếp chồng phải gắn thiết bị GPS làm tăng chi phí vận tải lên 30%; Phe Quản lý: Tuyến đường thủy nội địa có vô số nhánh sông, kênh rạch nhỏ; nếu không giám sát GPS từng container thì nguy cơ hàng quá cảnh bị rút ruột hoặc thẩm lậu vào nội địa là cực kỳ lớn.",
        "lesson": "Không có sự phân biệt 'dễ dãi' giữa vận tải đường bộ và vận tải đường thủy trong giám sát hải quan hiện đại. Số hóa định vị (GPS tracking) là luật chơi bắt buộc của toàn bộ chuỗi cung ứng quá cảnh xuyên biên giới.",
        "url": "https://en.wikipedia.org/wiki/Inland_waterway",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 121/2025/TT-BTC, Thông tư 38/2015/TT-BTC, Luật Hải quan 2014) & daily_notes.md (ghi nhận ngày 15/10/2026 từ case study Triệu Hưng).",
        "keywords": ["waterway", "transport"]
    },
    {
        "stt": 153,
        "week": "W42 — 15/10/2026",
        "title": "Chủ Hộ Kinh Doanh Bán Hàng Online Bị Treo Mã Số Thuế Gần 1 Tháng Trên eTax Mobile: Mất Trắng Đơn Hàng Tết 450 Triệu Vì Trùng Số CCCD Với Người Phụ Thuộc Cũ Từ Năm 2018",
        "stakeholders": "Chị Hoàng Mai Lan (Chủ hộ kinh doanh phụ kiện thời trang tại TP.HCM), Chi cục Thuế Khu vực Quận 7 - Nhà Bè, Trung tâm Dữ liệu Thuế Quốc gia.",
        "context": "Chị Lan nộp hồ sơ xin cấp mã số thuế hộ kinh doanh cá thể trên ứng dụng eTax Mobile để chuẩn bị xuất hóa đơn điện tử cho các đơn hàng sỉ nhân dịp cao điểm cuối năm. Tuy nhiên, sau 25 ngày gửi hồ sơ, ứng dụng vẫn báo trạng thái 'Đang xử lý / Lỗi dữ liệu'. Khi trực tiếp đến Chi cục Thuế hỏi, cán bộ đối soát phát hiện vào năm 2018, bố của chị Lan làm việc tại một công ty xây dựng ở tỉnh khác đã đăng ký chị làm 'Người phụ thuộc' giảm trừ gia cảnh bằng số CMND 9 số cũ. Khi triển khai Đề án 06 và Thông tư 86/2024/TT-BTC, hệ thống tự động quét và ghép nối số CCCD 12 số của chị với MST người phụ thuộc cũ, dẫn đến tình trạng 'Trùng lặp hồ sơ định danh thuế'. Hệ thống tự động khóa quy trình cấp MST hộ kinh doanh mới cho đến khi hoàn tất thủ tục hủy MST người phụ thuộc.",
        "damage": "Mất hợp đồng cung cấp 2.000 set quà tặng doanh nghiệp trị giá 450 triệu đồng do không thể xuất hóa đơn điện tử; Bị khách sỉ hủy cọc 50 triệu đồng; Mất gần 1 tháng đi lại giữa hai địa phương để xin giấy xác nhận cắt giảm trừ gia cảnh.",
        "root_cause": "Người nộp thuế không quản lý lịch sử thuế cá nhân; Không biết rằng mình từng có MST người phụ thuộc phát sinh từ quá khứ; Thiếu quy trình tự kiểm tra thông tin định danh thuế trước khi nộp hồ sơ mở hộ kinh doanh.",
        "solution": "1. Trước khi đăng ký MST hộ kinh doanh hoặc doanh nghiệp, bắt buộc tra cứu trên eTax Mobile hoặc tracuunnt.gdt.gov.vn xem CCCD của mình đã từng được gắn với MST nào (kể cả MST người phụ thuộc); 2. Nếu phát hiện có MST cũ, làm thủ tục cập nhật số CCCD vào MST cũ hoặc nộp Mẫu 08-MST xin hủy MST người phụ thuộc trước; 3. Sử dụng tài khoản định danh VNeID Mức 2 để xác thực đồng bộ thông tin dân cư trước khi bấm nộp hồ sơ đăng ký kinh doanh.",
        "opposing_views": "Phe Người dân: Lỗi do hệ thống thuế tự sinh mã người phụ thuộc ngày xưa mà không báo cho đương sự, nay chuyển đổi số lại bắt người dân chạy vạy khắp nơi để minh oan; Phe Cơ quan Thuế: Quy tắc một người chỉ có một mã định danh duy nhất là nguyên tắc nền tảng của quản lý thuế hiện đại, nếu không làm sạch dữ liệu thì không thể ngăn chặn gian lận giảm trừ gia cảnh.",
        "lesson": "Trong thời đại dữ liệu số liên thông, 'quá khứ thuế' của bạn không bao giờ biến mất. Hãy chủ động rà soát và làm sạch thông tin cá nhân trên eTax Mobile trước khi bắt tay vào bất kỳ kế hoạch kinh doanh lớn nào.",
        "url": "https://en.wikipedia.org/wiki/Taxpayer",
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 86/2024/TT-BTC, Thông tư 105/2020/TT-BTC, Luật Quản lý thuế 2019) & daily_notes.md (ghi nhận ngày 15/10/2026 từ case study Kim Anh).",
        "keywords": ["taxpayer", "tax"]
    },
    {
        "stt": 154,
        "week": "W42 — 15/10/2026",
        "title": "5 Container Sầu Riêng Đông Lạnh Xuất Khẩu Sang Trung Quốc Bị Trả Về Tại Cửa Khẩu Hữu Nghị: Mã Vùng Trồng Trên Thùng Carton Sai 1 Ký Tự So Với Hệ Thống GACC, Thiệt Hại 3.2 Tỷ Đồng",
        "stakeholders": "Công ty Cổ phần Nông sản XNK Tây Nguyên (Đắk Lắk), Nhà nhập khẩu tại Quảng Tây (Trung Quốc), Chi cục Hải quan Cửa khẩu Quốc tế Hữu Nghị (Lạng Sơn).",
        "context": "Doanh nghiệp xuất khẩu 5 container 40 feet sầu riêng múi cấp đông sang Trung Quốc theo đường bộ qua Cửa khẩu Hữu Nghị. Hàng hóa đã được cấp Giấy chứng nhận kiểm dịch thực vật tại nội địa. Tuy nhiên, khi lực lượng chức năng phía Trung Quốc (Hải quan Nam Ninh) quét mã QR và đối chiếu mã cơ sở đóng gói (PHC) in trên bao bì thùng carton với danh sách doanh nghiệp được cấp phép trên Hệ thống đăng ký doanh nghiệp nước ngoài của Tổng cục Hải quan Trung Quốc (GACC - cifer.singlewindow.cn), phát hiện mã PHC in trên nhãn bị gõ nhầm ký tự 'VN-DLPH-089' thành 'VN-DLPH-098'. Hải quan Trung Quốc lập tức ra thông báo từ chối thông quan và buộc tái xuất toàn bộ 5 container về Việt Nam theo Nghị định thư về kiểm dịch thực vật giữa hai nước.",
        "damage": "Lô hàng bị trả về phải hạ tải bán tháo nội địa với giá rẻ mạt, lỗ vốn 2.8 tỷ đồng; Chi phí vận chuyển hai chiều và tiền lưu cắm điện container lạnh suốt 14 ngày hết 420 triệu đồng (tổng thiệt hại hơn 3.2 tỷ đồng); Bị đình chỉ mã số đóng gói trong 03 tháng để khắc phục.",
        "root_cause": "Bộ phận in ấn bao bì carton của doanh nghiệp cẩu thả, không có quy trình kiểm soát chất lượng (QC) đối chiếu chéo maquette in với mã số chuẩn được GACC phê duyệt; Nhân viên kiểm hóa trước khi đóng container bỏ qua bước quét thử mã QR và đối chiếu từng ký tự mã định danh.",
        "solution": "1. Thiết lập quy trình 'Kiểm tra 3 điểm chạm': Đối chiếu mã PUC (vùng trồng) và PHC (đóng gói) giữa Hợp đồng, Giấy kiểm dịch và bao bì thực tế; 2. Dùng điện thoại quét thử mã QR trên từng pallet để xác nhận link dẫn về đúng trang phê duyệt của GACC; 3. Chỉ xuất xưởng container khi có biên bản đối soát chữ ký sống của Quản đốc đóng gói và Nhân viên xuất nhập khẩu.",
        "opposing_views": "Phe Doanh nghiệp: Chỉ nhầm lẫn đảo vị trí 2 chữ số do lỗi đánh máy mà bác bỏ 5 container hàng trị giá nhiều tỷ đồng là quá khắt khe, không có sự linh hoạt thương mại; Phe Cơ quan Kiểm dịch: Thị trường tỷ dân quản lý an toàn sinh học bằng mã vạch tự động hóa; một ký tự sai đồng nghĩa với việc nguồn gốc hàng hóa không được xác thực, có nguy cơ mang mầm bệnh vào nội địa.",
        "lesson": "Xuất khẩu chính ngạch sang Trung Quốc thời 4.0 không còn chỗ cho sự 'ước chừng'. Một lỗi chính tả nhỏ trên vỏ thùng carton có thể xóa sạch toàn bộ lợi nhuận của cả niên vụ sầu riêng.",
        "url": "https://en.wikipedia.org/wiki/Phytosanitary_certification",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Cục Bảo vệ Thực vật (Bộ NN&PTNT), Tổng cục Hải quan Trung Quốc (GACC Lệnh 248/249) & Thư Viện Pháp Luật (Nghị định thư xuất khẩu sầu riêng).",
        "keywords": ["phytosanitary", "plant"]
    },
    {
        "stt": 155,
        "week": "W42 — 15/10/2026",
        "title": "Tranh Chấp Hợp Đồng Thuê Tàu Chở Thép Cuộn Nhập Khẩu Từ Ấn Độ: Tàu Chờ 16 Ngày Ngoài Phao Số 0, Tranh Cãi 45.000 USD Tiền Phạt Phạt Lưu Tàu (Demurrage) Vì Thông Báo NOR Vô Hiệu",
        "stakeholders": "Công ty Cổ phần Cơ khí Thép Việt Nhật (Hải Phòng), Chủ tàu hàng rời (Shipowner) tại Mumbai (Ấn Độ), Tòa Trọng tài Hàng hải London (LMAA).",
        "context": "Doanh nghiệp Việt Nhật ký hợp đồng mua 15.000 tấn thép cuộn cán nguội theo điều kiện CFR Cảng Hải Phòng. Theo Hợp đồng thuê tàu (Charter Party), thời gian bốc dỡ hàng (Laytime) cho phép là 5 ngày làm việc liên tục. Khi tàu đến luồng phao số 0 Hải Phòng vào lúc 22h đêm Thứ Sáu, thuyền trưởng phát điện tín Thông báo sẵn sàng dỡ hàng (Notice of Readiness - NOR). Tuy nhiên, do luồng hàng hải bị cạn sau bão và cầu cảng đang kẹt tàu khác dỡ hàng xá, tàu phải thả neo ngoài biển suốt 16 ngày mới được hoa tiêu dẫn vào cập cầu. Khi dỡ xong hàng, Chủ tàu Ấn Độ gửi hóa đơn đòi phạt tiền lưu tàu (Demurrage) lên tới 45.000 USD (hơn 1.1 tỷ đồng) dựa trên nguyên tắc 'Once on Demurrage, Always on Demurrage'. Phía doanh nghiệp Việt Nam từ chối thanh toán với lập luận tàu chưa thực sự sẵn sàng cập cầu và thông báo NOR phát ngoài giờ hành chính là vô giá trị.",
        "damage": "Vụ việc bị chủ tàu đưa ra Tòa Trọng tài Hàng hải London (LMAA); Doanh nghiệp Việt Nam phải chi hơn 30.000 USD (khoảng 750 triệu đồng) thuê luật sư hàng hải quốc tế giải trình; Bị phong tỏa vận đơn cho các lô hàng tiếp theo, ảnh hưởng nghiêm trọng đến chuỗi cung ứng phôi thép của nhà máy.",
        "root_cause": "Hợp đồng mua bán ngoại thương không quy định rõ điều khoản 'Whether In Berth Or Not - WIBON' (Dù đã vào cầu hay chưa); Không thống nhất thời điểm NOR bắt đầu có hiệu lực và các trường hợp bất khả kháng do luồng hàng hải cạn; Doanh nghiệp nhập khẩu thiếu kinh nghiệm đàm phán hợp đồng thuê tàu chuyến (Voyage Charter).",
        "solution": "1. Khi mua hàng rời theo điều kiện CFR/CIF, bắt buộc yêu cầu người bán cung cấp hợp đồng thuê tàu chuyến để rà soát kỹ các điều khoản về Laytime, Demurrage và Despatch; 2. Cài cắm rõ điều khoản: 'Thời gian laytime chỉ bắt đầu tính khi tàu đã cập cầu an toàn và được cơ quan kiểm dịch/hải quan cấp phép dỡ hàng'; 3. Đàm phán mức phạt lưu tàu hợp lý theo ngày và có giới hạn trần chi phí tối đa (Demurrage Cap).",
        "opposing_views": "Phe Chủ hàng VN: Tàu chưa vào đến cầu cảng, không có cần cẩu dỡ hàng thì sao có thể coi là 'sẵn sàng dỡ hàng' mà tính phạt tiền của người mua; Phe Chủ tàu: Tàu đã đến vùng nước cảng theo đúng hợp đồng, việc cảng kẹt hay luồng cạn là rủi ro thương mại của bên thuê tàu, mỗi ngày tàu nằm im là chủ tàu mất hàng nghìn USD chi phí nhiên liệu và lương thuyền viên.",
        "lesson": "Trong vận tải hàng hải quốc tế, điều khoản tính thời gian bốc dỡ (Laytime) là cái bẫy chết người. Một dòng chữ 'WIBON' bị bỏ quên trong hợp đồng có thể làm bốc hơi toàn bộ biên lợi nhuận của lô hàng thép nhập khẩu.",
        "url": "https://en.wikipedia.org/wiki/Demurrage",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: London Maritime Arbitrators Association (LMAA Terms 2021), Công ước Liên hợp quốc về Chuyên chở hàng hóa bằng đường biển (Quy tắc Rotterdam) & Bộ luật Hàng hải Việt Nam 2015.",
        "keywords": ["demurrage", "laytime"]
    }
]
