# -*- coding: utf-8 -*-
"""
DATASET: CASE STUDY (CS) - BATCH W38 DAY 3 (STT 81 - 85)
Mốc thời gian: W38 — 19/09/2026
Tuân thủ: wf06_ideas_bank_update.md, 14 trường thông tin chuẩn xác
"""

CASE_STUDY_W38_D3 = [
    {
        "stt": 81,
        "week": "W38 — 19/09/2026",
        "title": "Quên Mở Tờ Khai Riêng Cho Kiện Hàng Phụ Tùng Nhỏ Đi Kèm Lô Hàng Lớn: Doanh Nghiệp Bị Tạm Giữ Hàng, Phạt 140 Triệu Đồng Vì Nghi Vấn Buôn Lậu & 12 Ngày Tắc Nghẽn Tại Cảng Hải Phòng",
        "parties": "Doanh nghiệp sản xuất cơ khí chính xác tại Hải Dương + Nhà cung cấp máy móc Quảng Đông (Trung Quốc) + Chi cục Hải quan Cửa khẩu Cảng Hải Phòng",
        "plot": "Tháng 08/2026, doanh nghiệp nhập khẩu lô dây chuyền máy đột dập CNC trị giá 420.000 USD từ Quảng Đông. Phía nhà cung cấp Trung Quốc 'nhiệt tình' tặng kèm và đóng thêm 1 kiện gỗ nhỏ chứa linh kiện dự phòng trị giá 8.500 USD vào đáy container nhưng lại không tách dòng riêng trên Bill of Lading và Packing List. Nhân viên làm thủ tục hải quan của công ty chỉ khai báo theo hóa đơn thương mại ban đầu mà không kiểm đếm thực tế số kiện đóng gói. Khi container đi qua máy soi chiếu tại Cảng Hải Phòng, hình ảnh soi phát hiện dôi dư 1 kiện hàng có cấu tạo kim loại khác biệt hoàn toàn với thân máy chính. Cơ quan hải quan lập tức chuyển luồng đỏ kiểm hóa 100%, niêm phong tạm giữ toàn bộ container với nghi vấn buôn lậu, gian lận thương mại trốn thuế.",
        "damage": "Bị xử phạt vi phạm hành chính 140 triệu đồng do hành vi nhập khẩu hàng hóa không khai báo hải quan; phát sinh 68 triệu đồng tiền lưu container và lưu bãi (Dem/Det) trong 12 ngày giải trình; chậm tiến độ bàn giao xưởng bị phạt hợp đồng 350 triệu đồng.",
        "root_cause": "Thiếu quy trình kiểm soát đối chiếu số kiện thực tế (Package Count) giữa Packing List chi tiết của xưởng đóng hàng và tờ khai hải quan VNACCS; chủ quan tin tưởng vào đối tác nước ngoài mà không yêu cầu xác nhận danh mục đóng gói cuối cùng trước khi bấm seal cont.",
        "prevention": "Thiết lập quy trình kiểm soát 3 lớp chứng từ: 1. Yêu cầu shipper gửi hình ảnh đóng hàng và Packing List cuối cùng ngay khi đóng seal; 2. Nếu phát hiện hàng gửi kèm ngoài hợp đồng, phải yêu cầu shipper gửi Thư giải trình sai sót (Explanation Letter) và mở tờ khai phi mậu dịch hoặc khai bổ sung AMA tự nguyện ngay trước khi truyền tờ khai chính thức; 3. Tuyệt đối không để hàng dôi dư không khai báo đi qua khu vực giám sát hải quan.",
        "opposing_views": "Doanh nghiệp thanh minh đây chỉ là hàng phụ tùng khuyến mãi tặng kèm của nhà sản xuất, không có động cơ buôn lậu; Phía Hải quan khẳng định nguyên tắc quản lý biên giới: mọi hàng hóa vào lãnh thổ Việt Nam không khai báo đều bị coi là hành vi nhập lậu và phải xử lý nghiêm theo khung pháp luật.",
        "lessons": "Trong xuất nhập khẩu, không có khái niệm 'tiện tay gửi kèm'; một kiện hàng ốc vít nhỏ không khai báo cũng đủ sức làm tê liệt cả dây chuyền triệu đô và hủy hoại uy tín tuân thủ hải quan của doanh nghiệp.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "keywords": ["customs", "declaration"],
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Trực quan: Thư giải trình sai sót của Shipper, Biên bản kiểm tra thực tế hàng hóa của Hải quan Cảng Hải Phòng & Daily Notes ngày 19/09/2026."
    },
    {
        "stt": 82,
        "week": "W38 — 19/09/2026",
        "title": "Bán Hàng Online Thu Hộ COD Bị eTax Mobile Quét Dữ Liệu Truy Thu Thuế & Phạt Chậm Nộp Gần 480 Triệu Đồng: Bài Học Đắt Giá Về Ảo Tưởng 'Doanh Thu Ẩn Không Ai Biết'",
        "parties": "Chủ shop kinh doanh thiết bị gia dụng và đồ chơi thông minh tại Hà Nội + Đơn vị vận chuyển (GHTK & Viettel Post) + Chi cục Thuế Quận Cầu Giấy",
        "plot": "Suốt từ năm 2024 đến giữa năm 2026, chủ shop tổ chức bán hàng qua fanpage Facebook và livestream cá nhân với doanh thu bình quân 700 - 800 triệu đồng/tháng (tổng doanh số tích lũy hơn 18 tỷ đồng). Toàn bộ đơn hàng đều được giao qua các đơn vị chuyển phát nhanh và thu tiền hộ COD đổ về tài khoản ngân hàng cá nhân. Cho rằng bán hàng không mở cửa hiệu và không xuất hóa đơn VAT thì cơ quan thuế không quản lý được, chủ shop không làm thủ tục đăng ký kinh doanh và không kê khai thuế. Đầu tháng 09/2026, sau khi Nghị định 132/2026/NĐ-CP có hiệu lực, ứng dụng eTax Mobile của chủ shop bất ngờ nhận thông báo rà soát nghĩa vụ thuế kèm bảng kê chi tiết từng mã vận đơn, số tiền COD nhận được từ GHTK và Viettel Post, yêu cầu lên cơ quan thuế giải trình trong vòng 10 ngày.",
        "damage": "Bị truy thu 1.5% thuế GTGT và TNCN trên tổng doanh thu rà soát (270 triệu đồng) + phạt hành chính hành vi trốn thuế 1 lần (165 triệu đồng) + tiền phạt chậm nộp tính theo mức 0.03%/ngày (45 triệu đồng) = Tổng cộng phải nộp ngân sách gần 480 triệu đồng.",
        "root_cause": "Thiếu hiểu biết pháp luật quản lý thuế hiện đại; ảo tưởng dòng tiền COD và tài khoản ngân hàng cá nhân là bí mật riêng tư; không nhận thức được việc Tổng cục Thuế đã triển khai Big Data kết nối trực tiếp với doanh nghiệp bưu chính và ngân hàng theo Đề án 06.",
        "prevention": "Chủ động đăng ký thành lập Hộ kinh doanh hoặc Công ty TNHH ngay khi doanh thu bán lẻ trực tuyến vượt mốc 100 triệu đồng/năm; hạch toán chi phí đầu vào có hóa đơn chứng từ đầy đủ; chủ động rà soát nộp hồ sơ kê khai bổ sung tự nguyện trước khi cơ quan thuế gửi trát thanh tra để được miễn giảm tối đa khung phạt vi phạm.",
        "opposing_views": "Chủ shop bức xúc cho rằng chi phí nhập hàng, phí quảng cáo Facebook và tỷ lệ hoàn đơn COD chiếm tới 85% doanh thu, nếu thu thuế trên tổng dòng tiền COD thì coi như mất sạch lợi nhuận; Cơ quan thuế viện dẫn Điều 51 Luật Quản lý thuế: cá nhân không thực hiện chế độ kế toán thì bắt buộc ấn định thuế trên toàn bộ doanh thu thực nhận.",
        "lessons": "Kỷ nguyên kinh doanh online trốn thuế đã chính thức chấm dứt; dòng tiền COD từ các đơn vị vận chuyển chính là 'vết chân số' rõ ràng nhất chứng minh doanh thu của bạn trước pháp luật.",
        "url": "https://en.wikipedia.org/wiki/Electronic_commerce",
        "keywords": ["electronic", "commerce"],
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Dẫn chứng thực tế: Thông báo rà soát nghĩa vụ thuế trên ứng dụng eTax Mobile & Dữ liệu thực tế từ scratch/daily_notes.md ngày 18/09 - 19/09/2026."
    },
    {
        "stt": 83,
        "week": "W38 — 19/09/2026",
        "title": "Tranh Chấp Áp Mã HS Lò Xo Khí Máy Nén Khí Bị Hải Quan Chuyển Luồng Đỏ, Nguy Cơ Truy Thu Thuế 15%: Bài Học Về Quy Tắc Áp Mã Bộ Phận Máy Chuyên Dụng",
        "parties": "Doanh nghiệp nhập khẩu phụ tùng máy công nghiệp Hải Phòng + Hãng thiết bị Ingersoll Rand + Chi cục Kiểm tra sau thông quan Cục Hải quan Hải Phòng",
        "plot": "Tháng 07/2026, doanh nghiệp nhập khẩu 1.200 chiếc lò xo khí (ty hơi - Gas Spring, hiệu Ingersoll Rand) dùng để tạo và duy trì lực căng cho dây đai máy nén khí công nghiệp. Doanh nghiệp áp mã HS 8414.90.90 (bộ phận của máy nén khí, thuế suất thuế nhập khẩu 0%). Tuy nhiên, khi kiểm tra hồ sơ sau thông quan, công chức hải quan viện dẫn Chú giải 2(a) Phần XVI và Công văn số 2358/TCHQ-TXNK, lập luận rằng ty hơi bản chất là cơ cấu xi lanh khí nén hoạt động độc lập có chức năng riêng, yêu cầu chuyển sang mã 8412.31.00 hoặc 8479.89.70 với mức thuế nhập khẩu lên tới 15% kèm thuế GTGT 10%.",
        "damage": "Nguy cơ bị ấn định truy thu thuế 320 triệu đồng cho toàn bộ các lô hàng đã thông quan trong 2 năm trước đó, cộng thêm khoản tiền phạt khai sai 20% (64 triệu đồng) và tiền chậm nộp thuế.",
        "root_cause": "Doanh nghiệp chỉ nhìn vào tên gọi thương mại trên catalogue ('part of compressor') mà không nắm vững nguyên tắc phân loại theo Chú giải pháp lý 2 Phần XVI (hàng hóa có chức năng riêng thuộc nhóm định danh cụ thể phải ưu tiên phân loại vào nhóm đó, không được áp vào mã bộ phận chung chung); không chủ động làm thủ tục Xác định trước mã HS.",
        "prevention": "Thực hiện thủ tục nộp hồ sơ xin Xác định trước mã số (Advance Ruling) tại Tổng cục Hải quan trước khi nhập khẩu thương mại quy mô lớn; lập bản phân tích kỹ thuật chi tiết chứng minh ty hơi này có kích thước và lực nén đặc thù chỉ lắp đặt được vào khung máy nén Ingersoll Rand mà không thể dùng cho mục đích cơ khí nào khác.",
        "opposing_views": "Doanh nghiệp khẳng định linh kiện này là bộ phận chuyên dùng không thể tách rời của máy nén khí (áp dụng Chú giải 2b); Cơ quan hải quan giữ vững quan điểm cơ cấu khí nén xi lanh đã được nêu danh cụ thể ở nhóm 8412 (áp dụng Chú giải 2a).",
        "lessons": "Phân loại mã HS linh kiện máy móc không thể dựa vào cảm tính 'lắp trên máy nào thì là bộ phận của máy đó'; Chú giải pháp lý của Danh mục HS mới là thước đo duy nhất quyết định số tiền thuế doanh nghiệp phải nộp.",
        "url": "https://en.wikipedia.org/wiki/Piston",
        "keywords": ["piston"],
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Dẫn chứng thực tế: Công văn 2358/TCHQ-TXNK, Chú giải 2 Phần XVI Danh mục HS & Vướng mắc thực tế từ scratch/daily_notes.md ngày 18/09/2026."
    },
    {
        "stt": 84,
        "week": "W38 — 19/09/2026",
        "title": "Khai Báo Sai Bản Chất Bộ Báo Nghẹt Lọc Gió Từ Thiết Bị Đo Áp Suất Thành Đồ Nhựa Phụ Tùng: Bị Ấn Định Thuế & Giữ Lô Hàng 15 Ngày Tại Cảng Hàng Không Nội Bài",
        "parties": "Công ty dịch vụ kỹ thuật bảo trì nhà xưởng Hà Nội + Hãng thiết bị lọc khí Ingersoll Rand + Chi cục Hải quan Cửa khẩu Sân bay Quốc tế Nội Bài",
        "plot": "Công ty nhập khẩu một lô gồm 500 chiếc Bộ báo nghẹt lọc gió (Air Filter Indicator) dùng cho máy nén khí. Thấy sản phẩm có lớp vỏ làm hoàn toàn bằng nhựa tổng hợp, nhân viên hiện trường khai báo mã HS 3926.90.99 (sản phẩm bằng plastic khác, thuế 10%) để truyền tờ khai luồng vàng thông quan nhanh. Tuy nhiên, khi hải quan chuyển luồng đỏ kiểm tra thực tế, công chức kiểm hóa phát hiện bên trong lớp vỏ nhựa là một cụm màng áp suất vi sai (Differential Pressure Diaphragm) tự động kích hoạt vạch chỉ thị màu đỏ khi chênh lệch áp suất lọc gió vượt ngưỡng 0.5 bar. Đây là một thiết bị đo kiểm áp suất vi sai thuộc Chương 90 (mã 9026.20), thuộc diện phải có chứng nhận kiểm tra đo lường chuyên ngành.",
        "damage": "Lô hàng bị đình chỉ thông quan 15 ngày chờ bổ sung kết quả giám định kỹ thuật; phạt vi phạm hành chính khai sai mã HS và thiếu giấy phép chuyên ngành 45 triệu đồng; chi phí lưu kho cảng hàng không ALS phát sinh 38 triệu đồng; bị nhà máy đối tác phạt hợp đồng bảo trì 280 triệu đồng do trễ hạn cung ứng vật tư.",
        "root_cause": "Nhân viên chứng từ nhìn bề ngoài vật liệu (nhựa) để chọn mã mà không đọc hiểu nguyên lý vận hành cốt lõi của thiết bị; vi phạm Chú giải 1(m) Chương 39 (loại trừ các dụng cụ, thiết bị đo lường thuộc Chương 90).",
        "prevention": "Bắt buộc nhân viên làm chứng từ xuất nhập khẩu phải tra cứu 6 Quy tắc tổng quát giải thích phân loại hàng hóa (GIR), đặc biệt Quy tắc 1 và Quy tắc 3(a): mặt hàng có mô tả cụ thể nhất (thiết bị đo áp suất Chương 90) phải được ưu tiên áp dụng hơn mô tả chung chung theo chất liệu (đồ nhựa Chương 39).",
        "opposing_views": "Doanh nghiệp giải thích đây chỉ là phụ kiện cảnh báo cơ học đơn giản, không dùng điện nên nghĩ là đồ nhựa; Hải quan khẳng định nguyên lý đo chênh áp cơ học vẫn là chức năng đo lường chuẩn xác thuộc phạm vi điều chỉnh của Chương 90.",
        "lessons": "Đừng bao giờ 'trông mặt bắt hình dong' khi phân loại mã hải quan; một linh kiện vỏ ngoài bằng nhựa nhưng bên trong có cơ chế đo lường vẫn phải tuân thủ nghiêm ngặt quy chuẩn kỹ thuật của Chương 90.",
        "url": "https://en.wikipedia.org/wiki/Pressure_measurement",
        "keywords": ["pressure", "measurement"],
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Dẫn chứng thực tế: Chú giải 1(m) Chương 39, Nhóm 9026 Danh mục HS & Bài học nghiệp vụ từ scratch/daily_notes.md ngày 18/09/2026."
    },
    {
        "stt": 85,
        "week": "W38 — 19/09/2026",
        "title": "Giải Cứu Thần Tốc 8 Container Thủy Sản Đông Lạnh Đi Nhật Bản Bị Hãng Tàu Cấp Nhầm Container Cài Sai Dải Nhiệt Độ Âm 5°C Thay Vì Âm 20°C Tại Cảng Cát Lái",
        "parties": "Doanh nghiệp chế biến thủy sản Cà Mau + Hãng tàu container quốc tế tại Cảng Cát Lái + Đội ngũ Kỹ thuật Giám định Độc lập SGS + Eureka Logistics",
        "plot": "Tháng 08/2026, công ty hoàn tất đóng gói 8 container lạnh 40 feet (Reefer Container) chứa 160 tấn tôm sú và cá tra phi lê đông lạnh xuất khẩu đi Cảng Yokohama (Nhật Bản), với yêu cầu nhiệt độ bảo quản nghiêm ngặt là âm 20°C (-20°C). Sau khi kéo hàng về hạ bãi Cát Lái chờ xếp lên tàu mẹ, chuyên viên giám sát chuỗi cung ứng lạnh của Eureka đã trích xuất dữ liệu cảm biến IoT Data Logger độc lập gắn bên trong container và phát hiện bảng điều khiển cắm điện bãi cảng của hãng tàu đang cài đặt ở mức âm 5°C (-5°C) do lỗi nhập liệu của nhân viên kỹ thuật bãi cont hãng tàu. Nếu không phát hiện kịp thời và tàu xuất bến, sau 7 ngày hải trình, toàn bộ 160 tấn thủy sản sẽ bị rã đông, tái kết tinh đá phá vỡ cấu trúc cơ thịt và bị Hải quan kiểm dịch Nhật Bản tiêu hủy 100% ngay khi cập cảng.",
        "damage": "Nguy cơ mất trắng toàn bộ lô hàng trị giá 18.5 tỷ đồng (khoảng 750.000 USD), đền bù hợp đồng thương mại và bị thị trường Nhật Bản đưa vào danh sách đen kiểm soát chất lượng.",
        "root_cause": "Lỗi cẩu thả trong thao tác cài đặt nhiệt độ bãi cont (Set Temperature) của kỹ thuật viên hãng tàu; sự phụ thuộc thụ động của các chủ hàng Việt Nam vào báo cáo của hãng tàu mà không có công cụ giám sát nhiệt độ độc lập.",
        "prevention": "Kích hoạt quy trình xử lý khủng hoảng chuỗi lạnh chuẩn Eureka: 1. Lập Biên bản bất thường tại bãi có sự chứng kiến và ký xác nhận của đại diện hãng tàu và giám định viên độc lập SGS; 2. Yêu cầu hãng tàu điều chỉnh ngay nhiệt độ về -20°C và kích hoạt chế độ làm lạnh cưỡng bức (Quick Freeze); 3. Đo nhiệt độ lõi sản phẩm (Core Temperature) bằng nhiệt kế đâm xuyên chuyên dụng xác nhận thịt tôm vẫn giữ mức an toàn -18.5°C trước khi ký lệnh cho phép bốc hàng lên tàu.",
        "opposing_views": "Hãng tàu ban đầu thoái thác cho rằng nhiệt độ hiển thị -5°C chỉ là chu kỳ rã đông dàn lạnh tự động (Defrost Cycle); Eureka và đơn vị giám định dùng biểu đồ ghi nhận liên tục từ Data Logger độc lập chứng minh nhiệt độ đã duy trì sai lệch suốt 6 giờ liên tục, buộc hãng tàu phải thừa nhận sai sót và bồi thường chi phí giám định.",
        "lessons": "Đối với hàng thủy sản xuất khẩu đi các thị trường khó tính như Nhật Bản hay EU, thiết bị giám sát nhiệt độ độc lập gắn bên trong container chính là chiếc 'phao cứu sinh' bảo vệ tài sản hàng chục tỷ đồng của doanh nghiệp.",
        "url": "https://en.wikipedia.org/wiki/Reefer_container",
        "keywords": ["container", "refrigerated"],
        "series": "Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Hồ sơ thực chiến Eureka Logistics: Biên bản giám định nhiệt độ container lạnh tại Cảng Cát Lái & Báo cáo kỹ thuật chuỗi cung ứng lạnh ISO 22000."
    }
]
