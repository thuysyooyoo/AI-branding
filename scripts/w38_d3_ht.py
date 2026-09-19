# -*- coding: utf-8 -*-
"""
DATASET: HỌC THỰC CHIẾN (HT) - BATCH W38 DAY 3 (STT 81 - 85)
Mốc thời gian: W38 — 19/09/2026
Tuân thủ: wf06_ideas_bank_update.md, 14 trường thông tin chuẩn xác
"""

HOC_THUC_CHIEN_W38_D3 = [
    {
        "stt": 81,
        "week": "W38 — 19/09/2026",
        "title": "Quy Trình 4 Bước Khai Bổ Sung Tờ Khai Hải Quan Sau Thông Quan (AMA) Cho Kiện Hàng Bị Bỏ Sót Hoặc Sai Lệch Số Lượng Mà Không Bị Phạt Nặng",
        "steps": "Bước 1: Lập Biên bản bất thường nội bộ ngay khi dỡ hàng hoặc kiểm tra chứng từ phát hiện dôi dư/thiếu sót kiện hàng; yêu cầu shipper nước ngoài phát hành Thư giải trình sai sót (Explanation Letter) có dấu mộc xác nhận lỗi đóng hàng nhầm.\nBước 2: Chuẩn bị bộ chứng từ khai bổ sung điện tử: Packing List sửa đổi, Commercial Invoice điều chỉnh, vận đơn có xác nhận của hãng vận chuyển và bản giải trình lý do sai sót.\nBước 3: Truy cập phần mềm ECUS5-VNACCS, tạo nghiệp vụ AMA (Khai bổ sung sau thông quan), nhập các chỉ tiêu thông tin cần sửa đổi (số lượng kiện, trọng lượng gross weight, trị giá tính thuế bổ sung), ký số token và truyền dữ liệu lên hệ thống.\nBước 4: Nộp hồ sơ chứng từ điện tử đính kèm qua Cổng Thông tin Một cửa Quốc gia (VNSW) theo Quyết định 1856/QĐ-TCHQ; theo dõi phản hồi phê duyệt của Chi cục Hải quan nơi mở tờ khai trong vòng 2 giờ làm việc và nộp số thuế phát sinh (nếu có).",
        "formula": "Số thuế nộp bổ sung = (Trị giá tính thuế kiện hàng sót x Tỷ giá tính thuế) x (Thuế suất XNK + Thuế GTGT). Miễn hoàn toàn mức phạt 20% số thuế khai thiếu nếu người nộp thuế chủ động gửi AMA tự nguyện trước khi cơ quan hải quan ban hành quyết định kiểm tra/thanh tra theo Điều 142 Luật Hải quan.",
        "top3_errors": "1. Tự ý vận chuyển kiện hàng sót ra khỏi kho bãi cảng khi chưa hoàn tất thủ tục khai bổ sung (bị quy vào tội buôn lậu).\n2. Không xin được Thư giải trình đóng hàng nhầm của bên bán khiến hồ sơ AMA bị cơ quan hải quan nghi vấn gian lận.\n3. Chậm trễ gửi hồ sơ quá 60 ngày kể từ ngày thông quan đối với các trường hợp tự phát hiện sai sót.",
        "mnemonic": "Biết sai khai trước là quân tử (được miễn phạt), để thuế soi ra hóa tiểu nhân (phạt 20% kèm hồ sơ rủi ro luồng đỏ).",
        "comparison": "Chủ động gửi tờ khai AMA tự nguyện theo Quyết định 1856/QĐ-TCHQ (thủ tục online 2 giờ, miễn phạt 20%) vs Bị Hải quan soi chiếu/hậu kiểm phát hiện dôi dư hàng hóa (phạt 20% đến 100% số thuế thiếu, đình chỉ luồng xanh trong 1 năm).",
        "persona": "Anh Minh (Giám đốc công ty thiết bị phụ tùng, gặp rắc rối vì xưởng Trung Quốc đóng kèm 1 thùng linh kiện mẫu mà không ghi vào tờ khai).",
        "checklist": "- [ ] Đã có Thư giải trình nguyên nhân sai sót của Shipper.\n- [ ] Đã cập nhật Packing List và Invoice sửa đổi.\n- [ ] Truyền nghiệp vụ AMA chuẩn xác trên VNACCS.\n- [ ] Nộp chứng từ giải trình trên Cổng Một cửa Quốc gia.\n- [ ] Hoàn thành nộp số thuế chênh lệch vào kho bạc ngay khi có phê duyệt.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "keywords": ["customs", "declaration"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Căn cứ Điều 20 Thông tư 38/2015/TT-BTC (sửa đổi bởi Thông tư 39/2018/TT-BTC), Quyết định 1856/QĐ-TCHQ & Thư Viện Pháp Luật."
    },
    {
        "stt": 82,
        "week": "W38 — 19/09/2026",
        "title": "Cẩm Nang 5 Bước Chuyển Đổi Từ Bán Hàng Online Cá Nhân Thu COD Sang Mô Hình Hộ Kinh Doanh Chuẩn Pháp Lý Để Tránh Bị Thuế Truy Thu Hàng Trăm Triệu",
        "steps": "Bước 1: Nộp hồ sơ đăng ký thành lập Hộ kinh doanh trực tuyến tại Cổng Dịch vụ công Quốc gia hoặc UBND Quận/Huyện nơi cư trú; nhận Giấy chứng nhận đăng ký kinh doanh và Mã số thuế 10 số liên thông với số định danh cá nhân CCCD 12 số trong vòng 3 ngày làm việc.\nBước 2: Mở tài khoản thanh toán chuyên dùng cho Hộ kinh doanh tại ngân hàng thương mại, đăng ký thông tin tài khoản này với Chi cục Thuế quản lý trực tiếp.\nBước 3: Đăng ký sử dụng Hóa đơn điện tử có mã của cơ quan thuế khởi tạo từ máy tính tiền theo Thông tư 78/2021/TT-BTC kết nối với phần mềm bán hàng đa kênh.\nBước 4: Cập nhật thông tin Mã số thuế HKD lên toàn bộ các tài khoản shop trên sàn TMĐT (Shopee, TikTok Shop) và hợp đồng ký với các đơn vị vận chuyển giao hàng thu COD (GHTK, Viettel Post, GHN, J&T).\nBước 5: Thực hiện chế độ sổ sách kế toán giản đơn theo Thông tư 88/2021/TT-BTC (ghi chép doanh thu, chi phí mua hàng) và kê khai nộp thuế định kỳ theo quý qua Cổng Thuế điện tử (thuedientu.gdt.gov.vn).",
        "formula": "Nghĩa vụ thuế Hộ kinh doanh TMĐT = Doanh thu bán lẻ x 1.5% (trong đó: Thuế GTGT 1% + Thuế TNCN 0.5%). Nếu không đăng ký: Số tiền phạt = Truy thu 1.5% tổng doanh thu COD + Phạt tiền chậm nộp 0.03%/ngày + Phạt hành chính trốn thuế từ 1 đến 3 lần số thuế thiếu.",
        "top3_errors": "1. Dùng lẫn lộn tài khoản ngân hàng cá nhân với tài khoản kinh doanh nhận tiền COD khiến không thể bóc tách số liệu giải trình với cơ quan thuế.\n2. Mua hàng trôi nổi không có hóa đơn chứng từ đầu vào, khi bị thanh tra không chứng minh được nguồn gốc hàng hóa hợp pháp.\n3. Quên không nộp tờ khai thuế định kỳ hàng quý dù doanh thu trong quý đó bằng 0.",
        "mnemonic": "Một chấm năm phần trăm (1.5%) nộp thuế ngủ ngon - Trốn thuế vài năm ăn mòn tài sản: Tuân thủ sớm là con đường duy nhất để phát triển kinh doanh bền vững.",
        "comparison": "Kinh doanh online tự do thu tiền COD trốn thuế (luôn nơm nớp lo sợ bị eTax Mobile quét dữ liệu truy thu, rủi ro vướng tội trốn thuế) vs Hộ kinh doanh chuẩn mực nộp thuế 1.5% (danh chính ngôn thuận, mở rộng quy mô, được vay vốn ngân hàng lãi suất ưu đãi).",
        "persona": "Chị Hương (Chủ shop bán hàng online thời trang và mỹ phẩm, đang hoang mang khi nhận được thông báo rà soát dòng tiền thu COD trên ứng dụng eTax Mobile).",
        "checklist": "- [ ] Nhận Giấy phép ĐKKD Hộ kinh doanh và MST 10 số.\n- [ ] Đăng ký tài khoản ngân hàng kinh doanh riêng biệt.\n- [ ] Kích hoạt Hóa đơn điện tử khởi tạo từ máy tính tiền.\n- [ ] Cập nhật MST trên app GHTK/Viettel Post/TikTok Shop.\n- [ ] Lưu trữ hóa đơn VAT nguyên liệu/hàng hóa mua vào.",
        "url": "https://en.wikipedia.org/wiki/Electronic_commerce",
        "keywords": ["electronic", "commerce"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Căn cứ Nghị định 132/2026/NĐ-CP, Thông tư 40/2021/TT-BTC & Thông tư 88/2021/TT-BTC; Thư Viện Pháp Luật & Cổng thông tin Tổng cục Thuế."
    },
    {
        "stt": 83,
        "week": "W38 — 19/09/2026",
        "title": "Bí Quyết 4 Bước Phân Biệt Mã HS Phụ Tùng Máy Chuyên Dụng (Chương 84) Với Cơ Cấu Độc Lập (Chương 84/85) Theo Chú Giải 2 Phần XVI",
        "steps": "Bước 1: Rà soát Chú giải 2(a) Phần XVI: Kiểm tra xem mặt hàng có được định danh cụ thể bằng tên gọi tại bất kỳ nhóm nào của Chương 84 hoặc Chương 85 hay không (ví dụ: máy bơm nhóm 8413, động cơ xi lanh khí nén nhóm 8412, van vòi nhóm 8481, vòng bi nhóm 8482). Nếu có, BẮT BUỘC phân loại vào nhóm đó, bất kể nó được thiết kế riêng cho loại máy nào.\nBước 2: Rà soát Chú giải 2(b) Phần XVI: Nếu mặt hàng KHÔNG thuộc diện nêu tại Chú giải 2(a), hãy chứng minh nó có phải là bộ phận chỉ dùng hoặc chủ yếu dùng cho một loại máy cụ thể nào hay không. Nếu chỉ lắp vừa vào một dòng máy duy nhất, áp mã bộ phận của nhóm máy đó.\nBước 3: Rà soát Chú giải 2(c) Phần XVI: Nếu bộ phận có thể dùng cho nhiều loại máy thuộc nhiều nhóm khác nhau của Chương 84/85, phân loại vào các nhóm bộ phận chung (8487 hoặc 8548).\nBước 4: Chuẩn bị Bộ hồ sơ kỹ thuật bảo vệ mã HS: Bản vẽ bóc tách kích thước hình học, hình ảnh vị trí lắp đặt trên máy tổng thể, thư xác nhận của hãng chế tạo khẳng định tính chuyên dụng độc quyền; nộp hồ sơ xin Xác định trước mã số trước khi nhập hàng số lượng lớn.",
        "formula": "Tiền thuế tiết kiệm = Trị giá tính thuế lô hàng x (Thuế suất nhóm chức năng độc lập 10-15% - Thuế suất nhóm bộ phận chuyên dùng 0-5%). Ví dụ: Lô linh kiện 2 tỷ đồng nếu áp đúng mã bộ phận thuế 0% sẽ tiết kiệm ngay 200-300 triệu đồng tiền thuế.",
        "top3_errors": "1. Nhầm lẫn giữa 'bộ phận cấu thành' và 'thiết bị có chức năng độc lập' (ví dụ coi lò xo khí là bộ phận máy nén khí thay vì xi lanh khí nén).\n2. Căn cứ hoàn toàn vào mô tả trên hợp đồng ngoại thương của bên bán ('compressor parts') mà bỏ qua Chú giải pháp lý HS.\n3. Không làm thủ tục Xác định trước mã số hải quan để có văn bản pháp lý bảo hộ giá trị trong 3 năm.",
        "mnemonic": "Chức năng riêng - Tên gọi riêng vào nhóm riêng (Chú giải 2a); Chuyên dùng một máy - Chung thân theo máy về nhà (Chú giải 2b).",
        "comparison": "Tự suy đoán mã HS theo thói quen (dễ bị kiểm tra sau thông quan bác mã, truy thu thuế 5 năm kèm phạt 20%) vs Làm thủ tục Xác định trước mã số HS tại Tổng cục Hải quan (được cấp văn bản phê duyệt mã HS có giá trị toàn quốc trong 3 năm).",
        "persona": "Anh Dũng (Kỹ sư trưởng phòng kỹ thuật và vật tư của nhà máy sản xuất linh kiện, thường xuyên phải đối chất về mã HS phụ tùng với hải quan cửa khẩu).",
        "checklist": "- [ ] Đã tra cứu bảng phân loại Chú giải 2(a) Phần XVI.\n- [ ] Đã có bản vẽ kỹ thuật chứng minh tính chuyên dụng Chú giải 2(b).\n- [ ] Có catalogue gốc đóng dấu giáp lai của nhà sản xuất.\n- [ ] Nộp hồ sơ xin Xác định trước mã số trước khi nhập hàng thương mại.\n- [ ] Viện dẫn Thông tư 85/2026/TT-BTC để bảo vệ tính đồng nhất mã số.",
        "url": "https://en.wikipedia.org/wiki/Harmonized_System",
        "keywords": ["harmonized", "system"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Căn cứ Chú giải 2 Phần XVI Danh mục Hàng hóa XNK Việt Nam, Công văn 2358/TCHQ-TXNK & Thông tư 85/2026/TT-BTC; Thư Viện Pháp Luật."
    },
    {
        "stt": 84,
        "week": "W38 — 19/09/2026",
        "title": "Quy Trình 5 Bước Kiểm Soát Container Lạnh (Reefer Container) Xuất Khẩu Từ Kho Ra Cảng: Bí Kíp Triệt Tiêu 100% Rủi Ro Tụt Nhiệt Biến Chất Hàng Hóa",
        "steps": "Bước 1: Kiểm tra Biên bản kiểm tra kỹ thuật container trước chuyến đi (Pre-Trip Inspection - PTI) của hãng tàu cấp; từ chối nhận vỏ cont nếu máy lạnh có cảnh báo lỗi cảm biến hoặc gioăng cao su cửa bị rách hở.\nBước 2: Cắm điện làm lạnh buồng container rỗng (Pre-cooling) đạt nhiệt độ cài đặt tiêu chuẩn (-20°C đối với tôm cá đông lạnh; +4°C đối với trái cây tươi) trước khi mở cửa bốc hàng vào.\nBước 3: Lắp đặt thiết bị ghi nhận dữ liệu nhiệt độ độc lập (IoT Real-time Temperature & Humidity Data Logger) tại 3 vị trí trọng yếu: Gần cửa cont, giữa sàn gió và trên nóc kiện hàng sâu bên trong buồng lạnh.\nBước 4: Đóng cửa container, kiểm tra thông gió (Ventilation) và bấm Seal hãng tàu + Seal hải quan trong vòng tối đa 45 phút kể từ lúc hoàn tất xếp hàng để tránh thất thoát nhiệt.\nBước 5: Giám sát lộ trình xe đầu kéo qua GPS và kiểm tra Phiếu theo dõi nhiệt độ cắm điện bãi cảng (Reefer Monitoring Sheet) mỗi 6 tiếng một lần cho đến khi container được bốc xếp an toàn lên tàu mẹ.",
        "formula": "Ngưỡng nhiệt độ an toàn hàng thủy sản đông lạnh xuất khẩu: t <= -18°C (dung sai tối đa cho phép ± 1.5°C trong thời gian chu kỳ xả đá không quá 45 phút). Nếu nhiệt độ ruột sản phẩm vượt quá -10°C, lô hàng có nguy cơ bị thị trường Nhật Bản từ chối tiếp nhận.",
        "top3_errors": "1. Đóng hàng vào container nóng khi buồng lạnh chưa được chạy Pre-cooling đạt chuẩn.\n2. Bác tài xe container tắt máy phát điện (GenSet) trên đường kéo từ kho ra cảng để bớt xén tiền dầu.\n3. Giao phó hoàn toàn trách nhiệm kiểm tra cắm điện tại bãi cảng Cát Lái/Hải Phòng cho hãng tàu mà không có nhân sự hiện trường giám sát.",
        "mnemonic": "Lạnh buồng trước - Đóng hàng sau - Gắn thiết bị đọc - Giám sát cắm sâu: 4 câu khẩu quyết bảo vệ trọn vẹn giá trị lô hàng container lạnh.",
        "comparison": "Quản lý container lạnh truyền thống phó mặc cho hãng tàu (rủi ro mất trắng 15-20 tỷ đồng khi tụt nhiệt) vs Giám sát nhiệt độ thời gian thực qua cảm biến IoT độc lập (nhận cảnh báo tức thời qua app điện thoại ngay khi nhiệt độ lệch 1°C).",
        "persona": "Anh Tuấn (Giám đốc HTX xuất khẩu thủy sản và trái cây Cà Mau, từng chịu bài học đắng cay vì container hàng đông lạnh bị cúp điện bãi cảng).",
        "checklist": "- [ ] Đã nhận phiếu PTI máy lạnh hoạt động hoàn hảo.\n- [ ] Đã chạy Pre-cooling đạt âm 20 độ C trước khi bốc hàng.\n- [ ] Đã kích hoạt 03 thiết bị IoT Data Logger bên trong cont.\n- [ ] Đã kiểm tra thông số cắm điện GenSet trên xe kéo container.\n- [ ] Nhận phiếu xác nhận cắm điện bãi cảng có chữ ký kỹ thuật viên cảng.",
        "url": "https://en.wikipedia.org/wiki/Cold_chain",
        "keywords": ["cold", "chain"],
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Cẩm nang quản lý vận hành chuỗi cung ứng lạnh tiêu chuẩn quốc tế (Cold Chain Logistics Best Practices) & Quy trình giám sát Reefer Eureka Logistics."
    },
    {
        "stt": 85,
        "week": "W38 — 19/09/2026",
        "title": "Chiến Lược 4 Vòng Đàm Phán Hợp Đồng Ngoại Thương Đưa Điều Khoản Phạt Vi Phạm Kiểm Tra Chuyên Ngành (MEPS / ATTP / Kiểm Dịch) Về Phía Nhà Cung Cấp Ngoại",
        "steps": "Vòng 1: Đưa Bộ Tiêu chuẩn Kỹ thuật bắt buộc của Việt Nam (QCVN, TCVN, Mức hiệu suất năng lượng tối thiểu MEPS) vào Phụ lục Kỹ thuật không thể tách rời của Hợp đồng mua bán ngoại thương (Sales Contract).\nVòng 2: Thiết kế Cơ chế Thanh toán Ràng buộc (Retained Payment Clause): Chỉ thanh toán trước 70-80% giá trị hợp đồng qua L/C hoặc T/T; giữ lại 20-30% còn lại thanh toán sau khi hàng hóa nhận được Giấy chứng nhận đạt chuẩn kiểm tra chuyên ngành và thông quan tại Việt Nam.\nVòng 3: Cài đặt Điều khoản Chế tài Trách nhiệm Buộc Tái Xuất (Rejection & Repatriation Liability): Quy định rõ: 'Nếu hàng hóa không đạt chứng nhận hợp quy/kiểm tra chuyên ngành do lỗi chất lượng từ phía bên bán, bên bán có trách nhiệm hoàn trả 100% số tiền đã nhận và chịu toàn bộ chi phí vận chuyển tái xuất ra khỏi lãnh thổ Việt Nam trong vòng 15 ngày'.\nVòng 4: Lựa chọn Cơ quan Giải quyết Tranh chấp Trọng tài thuận lợi: Thỏa thuận lựa chọn Trung tâm Trọng tài Quốc tế Việt Nam (VIAC) hoặc Trọng tài Singapore (SIAC) thay vì chấp nhận điều khoản kiện tụng tại tòa án nội địa của bên bán.",
        "formula": "Số tiền cọc an toàn = Trị giá hợp đồng - (Chi phí vận chuyển khứ hồi ước tính + Chi phí lưu bãi cảng dự kiến + 20% dự phòng rủi ro hủy hợp đồng).",
        "top3_errors": "1. Chấp nhận thanh toán 100% tiền hàng trước khi tàu cập cảng Việt Nam mà không có cam kết bảo lãnh chất lượng.\n2. Tin tưởng vào chứng chỉ xuất xưởng nội bộ (Factory Test Certificate) của nhà sản xuất mà không yêu cầu chứng chỉ kiểm định độc lập của bên thứ ba (SGS, TUV, Quatest).\n3. Không quy định rõ ràng nghĩa vụ chi trả phí lưu kho lưu bãi (Demurrage/Detention) khi hàng bị cơ quan quản lý nhà nước tạm giữ lấy mẫu thử nghiệm.",
        "mnemonic": "Tiền trao khi hàng đạt chuẩn - Điều khoản phạt chặt thì ngoại bang mới kiêng nể: Pháp lý chặt chẽ là vũ khí bảo vệ tiền tài cho doanh nhân Việt.",
        "comparison": "Ký hợp đồng ngoại thương sơ sài theo mẫu đối tác đưa (DN Việt lãnh trọn hậu quả khi hàng bị buộc tái xuất) vs Cài đặt điều khoản ràng buộc tiêu chuẩn chuyên ngành theo chuẩn CISG (bảo đảm an toàn 100% vốn đầu tư, đối tác ngoại tự giác kiểm định từ xưởng).",
        "persona": "Chị Mai (Trưởng phòng Thu mua của tập đoàn bán lẻ thiết bị điện máy, phụ trách đàm phán các hợp đồng nhập khẩu động cơ và thiết bị điện từ đối tác nước ngoài).",
        "checklist": "- [ ] Đính kèm Phụ lục Tiêu chuẩn QCVN/MEPS vào Hợp đồng.\n- [ ] Giữ lại tối thiểu 20% trị giá thanh toán sau thông quan kiểm định.\n- [ ] Cài đặt điều khoản bên bán chịu 100% chi phí nếu bị buộc tái xuất.\n- [ ] Yêu cầu cấp chứng chỉ kiểm định chất lượng phòng lab quốc tế.\n- [ ] Thỏa thuận điều khoản trọng tài thương mại quốc tế độc lập.",
        "url": "https://en.wikipedia.org/wiki/United_Nations_Convention_on_Contracts_for_the_International_Sale_of_Goods",
        "keywords": ["international", "goods"],
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Căn cứ Công ước Viên 1980 về Hợp đồng Mua bán Hàng hóa Quốc tế (CISG), Luật Thương mại Việt Nam & Cẩm nang soạn thảo hợp đồng thương mại quốc tế ICC."
    }
]
