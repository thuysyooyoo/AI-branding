# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W40 DAY 4 — 01/10/2026)
Số lượng: 5 ý tưởng (STT 121 - 125)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, số liệu thiệt hại cụ thể, URL thật đã qua url_validator.py
"""

CASE_STUDY_W40_D4 = [
    {
        "stt": 121,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "15 Container Linh Kiện Bị Bác Thủ Tục XNK Tại Chỗ: Doanh Nghiệp Điện Tử Bị Đòi 3.5 Tỷ Đồng Thuế VAT Do Thiếu Văn Bản Chỉ Định Giao Hàng Của Thương Nhân Nước Ngoài (Inland Export Delivery Order Trap)",
        "stakeholders": "Công ty TNHH Công Nghệ Cao H.K (Bắc Giang) vs Nhà mua hàng tại Đài Loan vs Doanh nghiệp Chế xuất L.G (Hải Phòng) vs Chi cục Hải quan Quản lý Hàng Gia công",
        "context": "H.K ký hợp đồng cung cấp linh kiện nhựa và bo mạch cho một tập đoàn tại Đài Loan, theo thỏa thuận miệng hàng được giao thẳng cho nhà máy DNCX tại KCN Tràng Duệ (Hải Phòng) để lắp ráp màn hình. Khi mở tờ khai xuất khẩu tại chỗ, H.K không cung cấp được văn bản chỉ định giao hàng chính thức bằng văn bản hoặc phụ lục hợp đồng từ đối tác Đài Loan mà chỉ có tin nhắn Zalo/WeChat trao đổi giữa nhân viên bán hàng. Hải quan kiểm tra sau thông quan kết luận giao dịch không đáp ứng điều kiện XNK tại chỗ theo Điều 47a Luật Hải quan.",
        "damage": "Bị bác tư cách XNK tại chỗ, toàn bộ lô hàng 15 container trị giá 35 tỷ đồng bị coi là giao dịch mua bán nội địa; doanh nghiệp bị truy thu 3.5 tỷ đồng tiền thuế GTGT 10%; xử phạt 20% hành vi khai sai (700 triệu đồng) và tiền chậm nộp 180 triệu đồng.",
        "root_cause": "Thói quen làm việc tùy tiện, dựa vào thỏa thuận miệng và tin nhắn mạng xã hội; nhân viên xuất nhập khẩu không nắm vững nguyên tắc pháp lý cốt tử của hợp đồng ngoại thương 3 bên (Hợp đồng mua bán gốc + Văn bản chỉ định người nhận hàng tại Việt Nam).",
        "solution": "Quy trình 3 bước chuẩn: 1. Bắt buộc cài điều khoản chỉ định giao hàng (Designated Consignee Clause) ngay trong Hợp đồng mua bán ngoại thương hoặc phát hành Phụ lục Hợp đồng (Addendum) có chữ ký số/con dấu của đối tác ngoại; 2. Yêu cầu bên nhận hàng ký Biên bản bàn giao hàng hóa có đối chiếu số tờ khai hải quan đối ứng; 3. Kiểm tra chéo mã loại hình và thông tin người nhập khẩu tại chỗ trước khi truyền tờ khai xuất.",
        "opposing_views": "Phe A (Doanh nghiệp): Hàng hóa thực tế đã đưa vào khu chế xuất để lắp ráp xuất khẩu, không tiêu thụ nội địa một gram nào, phạt truy thu thuế VAT 10% là máy móc và làm kiệt quệ doanh nghiệp. Phe B (Hải quan): Pháp luật quy định điều kiện XNK tại chỗ rất nghiêm ngặt để chống gian lận thương mại; nếu không có văn bản chỉ định hợp pháp thì không thể xác định được quyền sở hữu hàng hóa và nghĩa vụ thuế của thương nhân nước ngoài.",
        "lesson": "Trong nghiệp vụ hải quan, 'thỏa thuận miệng là con số 0, tin nhắn mạng xã hội không có giá trị pháp lý'. Một dòng chữ chỉ định giao hàng hợp pháp trên hợp đồng ngoại thương đáng giá hàng tỷ đồng tiền thuế.",
        "url": "https://en.wikipedia.org/wiki/Customs",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (Công văn 22411/CHQ-GSQL ngày 30/09/2026 về XNK tại chỗ) & Thư Viện Pháp Luật (Điều 47a Luật Hải quan).",
        "keywords": ["customs"]
    },
    {
        "stt": 122,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "4 Container Túi Vô Trùng Nhập Khẩu Bị 'Giam' 18 Ngày Tại Cảng Cát Lái: Bài Học Đắt Giá Khi Nhầm Lẫn 'Tự Công Bố' Là 'Miễn Kiểm Tra An Toàn Thực Phẩm' (Aseptic Bag Food Contact Crisis)",
        "stakeholders": "Công ty CP Chế Biến Thực Phẩm & Nước Giải Khát Á Châu (Bình Dương) vs Chi cục Hải quan Cửa khẩu Cảng Sài Gòn KV1 vs Chi cục An toàn Vệ sinh Thực phẩm TP.HCM",
        "context": "Á Châu nhập khẩu 4 container túi vô trùng chuyên dụng (Aseptic bag - HS 3923.21.99) từ Đức dùng đóng gói nước ép xoài xuất khẩu. Đọc Thông tư 27/2026/TT-BYT thấy bao bì tiếp xúc thực phẩm thuộc nhóm rủi ro trung bình thực hiện 'Tự công bố theo Nghị định 15/2018/NĐ-CP', giám đốc công ty yên tâm cho rằng đã tự công bố trên mạng thì được miễn kiểm tra khi hàng về. Tuy nhiên khi tàu cập cảng Cát Lái, Hải quan yêu cầu xuất trình Giấy thông báo kết quả kiểm tra nhà nước về ATTP theo Danh mục Thông tư 15/2024/TT-BYT. Doanh nghiệp không đăng ký kiểm tra trước nên hệ thống không cho thông quan.",
        "damage": "4 container bị lưu bãi suốt 18 ngày; phát sinh 165 triệu đồng chi phí lưu container và bãi (demurrage/detention); nhà máy phải dừng chạy dây chuyền đóng gói nước ép 1 tuần, chịu phạt hợp đồng giao hàng trễ 320 triệu đồng cho đối tác Hàn Quốc.",
        "root_cause": "Nhầm lẫn tai hại giữa hai khái niệm pháp lý: 'Hồ sơ tự công bố sản phẩm lưu hành nội địa' (theo Nghị định 15) và 'Thủ tục kiểm tra nhà nước về an toàn thực phẩm nhập khẩu' (theo Thông tư 15). Không cập nhật kịp thời văn bản giải thích nghiệp vụ giữa Cục Hải quan và Cục ATTP (Công văn 22082/CHQ-GSQL).",
        "solution": "Cẩm nang 3 bước kiểm tra bao bì tiếp xúc thực phẩm: 1. Tra cứu chéo đồng thời Thông tư 27/2026/TT-BYT và Thông tư 15/2024/TT-BYT; 2. Đăng ký kiểm tra nhà nước về ATTP trên Cổng Một cửa Quốc gia (NSW) tối thiểu 3 ngày trước khi tàu cập cảng; 3. Chuẩn bị sẵn trọn bộ hồ sơ gồm: Bản tự công bố + Chứng nhận kiểm nghiệm thành phần thôi nhiễm (Migration Test) từ nhà sản xuất châu Âu để nộp lấy mẫu nhanh trong 24 giờ.",
        "opposing_views": "Phe A (Doanh nghiệp): Bao bì đã qua kiểm định đạt chuẩn an toàn của Đức và đã nộp bản tự công bố, bắt kiểm tra lại từng lô hàng là lãng phí và cản trở sản xuất. Phe B (Cơ quan Kiểm dịch & Hải quan): Bao bì tiếp xúc trực tiếp đồ ăn uống nếu làm từ nhựa tái sinh độc hại sẽ nguy hiểm khôn lường cho sức khỏe cộng đồng; pháp luật quy định phải kiểm tra nhà nước thì cơ quan thừa hành bắt buộc phải kiểm soát.",
        "lesson": "Đọc văn bản pháp luật không được 'cắt ngọn bớt cành'. Sự khác biệt giữa 'tự công bố' và 'miễn kiểm tra chuyên ngành' là ranh giới sống còn giữa một lô hàng thông quan trong 2 giờ và một lô hàng bị chôn chân gần 1 tháng tại cảng.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (Công văn 22082/CHQ-GSQL & 1913/ATTP-KN ngày 24/09/2026) & Thư Viện Pháp Luật (Thông tư 27/2026/TT-BYT & Thông tư 15/2024/TT-BYT).",
        "keywords": ["food", "safety"]
    },
    {
        "stt": 123,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Lô Hàng Thiết Bị Điện Tử Đưa Về Kho Bảo Quản Quá Hạn 30 Ngày Do Phòng Lab Chậm Trả Kết Quả: Doanh Nghiệp Bị Phạt 40 Triệu Đồng & Tước Quyền Bảo Quản Suốt 1 Năm (Storage Extension Trap)",
        "stakeholders": "Công ty TNHH Công Nghệ Viễn Thông Tân Phát (TP.HCM) vs Trung tâm Kỹ thuật Đo lường Chất lượng 3 (Quatest 3) vs Chi cục Hải quan Cửa khẩu Sân bay Quốc tế Tân Sơn Nhất",
        "context": "Tân Phát nhập khẩu lô hàng 500 bộ thiết bị phát sóng không dây thuộc danh mục quản lý rủi ro cao theo Thông tư 121/2025/TT-BTC. Doanh nghiệp làm thủ tục đưa hàng về kho riêng tại KCN Tân Bình để bảo quản chờ kết quả thử nghiệm hiệu năng và tương thích điện từ (EMC). Do thiết bị phức tạp và phòng lab của Quatest 3 quá tải, thời gian thử nghiệm kéo dài quá 30 ngày. Khi Tân Phát đề nghị Quatest 3 ký xác nhận văn bản gia hạn theo mẫu của Hải quan, bên Quatest từ chối với lý do quy chế nội bộ không ký giấy tờ bảo lãnh tiến độ cho khách. Ngày thứ 32, Hải quan phát hiện quá hạn 30 ngày chưa nộp kết quả kiểm tra, lập tức ra quyết định xử phạt vi phạm hành chính.",
        "damage": "Doanh nghiệp bị phạt tiền 40 triệu đồng; bị tước quyền đưa hàng về kho bảo quản trong vòng 12 tháng (mọi lô hàng nhập khẩu trong 1 năm tiếp theo phải lưu giữ tại kho sân bay, chịu phí lưu kho 250 USD/ngày cho đến khi có kết quả thử nghiệm); chi phí đội lên ước tính hơn 300 triệu đồng.",
        "root_cause": "Doanh nghiệp chủ quan trông chờ vào phòng lab bên ngoài mà không theo dõi thời hạn 'cửa sổ 30 ngày' của cơ quan Hải quan; không nộp văn bản giải trình kịp thời kèm chứng từ tiến độ thử nghiệm thực tế lên hệ thống Một cửa trước khi hết hạn 30 ngày.",
        "solution": "Chiến thuật tác chiến 3 bước: 1. Đặt lịch kiểm nghiệm và chốt cam kết tiến độ bằng văn bản với phòng lab trước khi mở tờ khai đưa hàng về bảo quản; 2. Thiết lập hệ thống đếm ngược (countdown) tại ngày thứ 20: nếu phòng lab chưa có kết quả, lập tức phát hành công văn báo cáo tiến độ và xin gia hạn nộp cho Hải quan kèm phiếu tiếp nhận mẫu thử nghiệm gốc; 3. Tuyệt đối không để vượt qua mốc 30 ngày mà không có dấu vết giải trình trên hệ thống Một cửa.",
        "opposing_views": "Phe A (Doanh nghiệp): Doanh nghiệp không hề gian lận hay tẩu tán hàng, hàng vẫn nằm nguyên niêm phong trong kho; việc chậm trễ do phòng thí nghiệm nhà nước quá tải sao lại đổ lỗi phạt doanh nghiệp? Phe B (Hải quan): Thời hạn 30 ngày là quy định pháp luật bất di bất dịch trong Thông tư 121; nếu thả lỏng không kiểm soát thì việc quản lý hàng đưa về bảo quản sẽ hoàn toàn mất kiểm soát.",
        "lesson": "Mang hàng về kho bảo quản là 'đặc ân', không phải quyền đương nhiên. Chậm một ngày giải trình có thể tước đi đặc quyền kinh doanh của doanh nghiệp trong suốt cả năm trời.",
        "url": "https://en.wikipedia.org/wiki/Conformity_assessment",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (case vướng mắc gia hạn kiểm tra chất lượng ngày 25/09/2026) & Thư Viện Pháp Luật (Thông tư 121/2025/TT-BTC & Nghị định 128/2020/NĐ-CP).",
        "keywords": ["conformity", "assessment"]
    },
    {
        "stt": 124,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Nhập Khẩu 2 Container Hành Lá Sấy Nhiệt Áp Mã Miễn Thuế VK010: Doanh Nghiệp Gia Vị Bị Bác Bỏ, Truy Thu & Phạt Gần 280 Triệu Đồng Thuế VAT (Dehydrated Agricultural VAT Trap)",
        "stakeholders": "Công ty TNHH Chế Biến Nông Sản & Gia Vị Nam An (Hải Dương) vs Nhà cung cấp Sơn Đông (Trung Quốc) vs Chi cục Hải quan Cửa khẩu Cảng Hải Phòng",
        "context": "Nam An nhập khẩu 40 tấn hành lá sấy nhiệt, đã cắt nhỏ đóng gói kín (HS 0712.90.90) từ Trung Quốc về làm gia vị cho gói mì ăn liền. Kế toán đọc Khoản 1 Điều 4 Nghị định 181/2025/NĐ-CP thấy nông sản chỉ qua sơ chế thông thường ở khâu nhập khẩu không chịu VAT nên đã khai báo mã miễn giảm thuế VK010 (thuế VAT 0%). Tuy nhiên, khi kiểm hóa và gửi mẫu giám định, Chi cục Kiểm định Hải quan kết luận: Hành lá đã qua xử lý nhiệt độ cao trên 80°C bằng luồng khí nóng công nghiệp và cắt vụn theo quy chuẩn sản phẩm ăn liền, đây là sản phẩm đã qua chế biến công nghiệp sâu, không phải sơ chế thông thường.",
        "damage": "Hải quan bác bỏ mã VK010, ấn định áp mức thuế suất GTGT 10% (hoặc 5%); doanh nghiệp bị truy thu 215 triệu đồng tiền thuế VAT; phạt 20% hành vi khai sai (43 triệu đồng) và tiền chậm nộp 18 triệu đồng.",
        "root_cause": "Doanh nghiệp tự suy diễn quy trình sản xuất theo cảm tính, không phân biệt được bản chất kỹ thuật giữa 'sấy khô tự nhiên/phơi gió sơ chế' và 'sấy nhiệt công nghiệp khử nước'; không tham vấn chuyên môn hoặc xin giám định trước khi truyền tờ khai.",
        "solution": "Quy trình 3 bước tự bảo vệ khi nhập khẩu nông sản khô: 1. Yêu cầu nhà xưởng nước ngoài cung cấp bản mô tả quy trình công nghệ chi tiết (Flowchart & Technical Specification) ghi rõ nhiệt độ và phương pháp làm khô; 2. Tra cứu tiền lệ phân loại và đối chiếu với Thông tư 219/2013/TT-BTC và Nghị định 181/2025/NĐ-CP; 3. Nếu còn nghi ngờ ranh giới sơ chế vs chế biến, làm văn bản tham vấn hoặc xin Xác định trước mã số/chính sách thuế từ Tổng cục Hải quan trước khi nhập khẩu số lượng lớn.",
        "opposing_views": "Phe A (Chủ hàng): Hành lá sấy khô chỉ mất nước chứ không cho thêm bất kỳ hóa chất hay đường muối nào, bản chất vẫn là cây nông sản thì phải được miễn thuế theo chính sách khuyến khích nông nghiệp. Phe B (Hải quan & Cơ quan Thuế): Máy sấy nhiệt công nghiệp cắt vụn đã làm chín một phần và biến đổi tế bào thực vật, tạo ra giá trị gia tăng công nghiệp, nên bắt buộc phải điều tiết thuế GTGT bình đẳng với hàng chế biến trong nước.",
        "lesson": "Đừng bao giờ 'nhìn mặt đặt tên' cho chính sách thuế. Một chút nhiệt độ của máy sấy công nghiệp có thể làm bốc hơi toàn bộ lợi nhuận của cả chuyến hàng nếu doanh nghiệp chọn sai mã thuế suất VAT.",
        "url": "https://en.wikipedia.org/wiki/Value-added_tax",
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (case mã thuế suất VAT VK010 hành lá sấy khô ngày 25/09/2026) & Thư Viện Pháp Luật (Nghị định 181/2025/NĐ-CP & Thông tư 219/2013/TT-BTC).",
        "keywords": ["tax", "value-added"]
    },
    {
        "stt": 125,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Nhầm Lẫn Giữa Mã Loại Hình A11 Và A12: Doanh Nghiệp FDI Nhập Khẩu Bán Buôn Bị Treo Hoàn Thuế 1.8 Tỷ Đồng Suốt 8 Tháng (Customs Procedure Code Discrepancy)",
        "stakeholders": "Công ty TNHH Công Nghiệp Cơ Khí V. (KCN VSIP Bình Dương - Doanh nghiệp 100% vốn nước ngoài) vs Chi cục Hải quan Quản lý Đầu tư vs Cục Thuế Tỉnh Bình Dương",
        "context": "Công ty V. có giấy chứng nhận đầu tư vừa sản xuất vừa thực hiện quyền phân phối xuất nhập khẩu hàng hóa. Tháng 11/2025, doanh nghiệp nhập khẩu lô linh kiện van thủy lực trị giá 18 tỷ đồng để bán buôn lại cho các nhà máy tại Việt Nam theo quyền phân phối. Nhân viên khai thuê hải quan quen tay chọn mã loại hình A11 (Nhập khẩu tiêu dùng của DN trong nước) thay vì mã A12 (Nhập khẩu của doanh nghiệp có vốn đầu tư nước ngoài theo quyền nhập khẩu/phân phối). Doanh nghiệp đã nộp đủ thuế nhập khẩu và VAT. Tuy nhiên, đến kỳ hoàn thuế và quyết toán tài chính năm 2026, cơ quan thuế từ chối hoàn 1.8 tỷ đồng thuế VAT đầu vào do tờ khai mang mã A11 không tương thích với quyền xuất nhập khẩu đăng ký của pháp nhân FDI trên hệ thống Một cửa.",
        "damage": "Khoản tiền hoàn thuế 1.8 tỷ đồng bị 'treo' suốt 8 tháng gây ách tắc dòng tiền lưu động; doanh nghiệp phải thực hiện thủ tục xin sửa đổi bổ sung sau thông quan phức tạp, nộp hồ sơ giải trình qua 3 cấp phòng ban và chịu phạt vi phạm thủ tục hải quan.",
        "root_cause": "Khai thuê hải quan thiếu chuyên nghiệp, không nắm vững Quyết định 1357/QĐ-TCHQ về bảng mã loại hình XNK; không phân biệt được sự khác biệt pháp lý giữa doanh nghiệp nội địa (A11) và doanh nghiệp có vốn FDI thực hiện quyền phân phối (A12).",
        "solution": "Bảng quy chuẩn 3 bước: 1. Đối chiếu Giấy chứng nhận đăng ký đầu tư (IRC) và quyền kinh doanh của doanh nghiệp trước khi truyền tờ khai; 2. Lập ma trận kiểm soát mã loại hình nội bộ: Hàng sản xuất dùng E21/E31, hàng phân phối thương mại FDI bắt buộc dùng A12, chỉ doanh nghiệp 100% vốn Việt Nam mới dùng A11; 3. Bắt buộc kiểm tra chéo (Cross-check) giữa kế toán thuế và nhân viên hiện trường trước khi ký số truyền dữ liệu VNACCS.",
        "opposing_views": "Phe A (Doanh nghiệp): Doanh nghiệp đã nộp đủ 100% tiền thuế vào kho bạc nhà nước, bản chất hàng hóa không đổi thì sai sót mã ký hiệu A11 hay A12 chỉ là lỗi đánh máy hành chính nhỏ, không nên gây khó khăn việc hoàn thuế. Phe B (Hải quan & Cơ quan Thuế): Hệ thống quản trị thuế và hải quan hiện đại vận hành tự động theo thuật toán quản lý rủi ro; sai mã loại hình làm méo mó số liệu thống kê vĩ mô và tiềm ẩn nguy cơ doanh nghiệp FDI lấn sân kinh doanh ngoài phạm vi giấy phép đầu tư.",
        "lesson": "Trong khai báo hải quan điện tử, mã loại hình là 'mã định danh linh hồn' của tờ khai. Chọn sai một chữ số trong mã loại hình có thể làm 'đóng băng' hàng tỷ đồng vốn hoàn thuế của doanh nghiệp.",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "series": "Series: Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (vướng mắc phân loại mã loại hình A11, A12, A41 theo Công văn 21939/HQKV2-NVHQ) & Thư Viện Pháp Luật (Quyết định 1357/QĐ-TCHQ).",
        "keywords": ["customs"]
    }
]
