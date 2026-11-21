# -*- coding: utf-8 -*-
"""
DỮ LIỆU Ý TƯỞNG PILLAR: CASE STUDY (CS) — TUẦN W40 (28/09/2026)
STT: 106 - 110
Tuân thủ đầy đủ 14 trường dữ liệu theo wf06_ideas_bank_update.md & skills/pillar-cs-casestudy/SKILL.md
"""

CASE_STUDY_W40 = [
    {
        "stt": 106,
        "week": "W40 — 28/09/2026",
        "title": "Container Nhập Khẩu Bị Bấm Seal Mới Do Gãy Seal Khi Chuyển Cảng: Khai Nhầm Số Seal Cũ Bị Dừng Thông Quan Khẩn Cấp & Phạt 28 Triệu Đồng (Customs Seal Mishap)",
        "parties": "Doanh nghiệp thiết bị may mặc Long An, Hãng tàu Maersk, Cảng Tân Cảng Cát Lái, Cảng Tân Cảng Hiệp Phước, Chi cục Hải quan Cửa khẩu Cảng Hiệp Phước",
        "plot": "Tháng 09/2026, doanh nghiệp nhập khẩu 1 container 40HC nguyên phụ liệu dệt may từ Thượng Hải về cảng Cát Lái và đăng ký chuyển cảng về Tân Cảng Hiệp Phước làm thủ tục thông quan. Trong quá trình cẩu hạ bãi tại Cát Lái, chốt khóa seal hãng tàu bị va chạm gãy rời. Đơn vị điều độ cảng Cát Lái đã lập biên bản sự cố kỹ thuật nội bộ và bấm bổ sung seal cảng mới để bảo vệ an toàn hàng hóa. Khi làm thủ tục nhận cont tại Hiệp Phước, nhân viên chứng từ của DN vội vàng lấy số seal hãng tàu trên Bill of Lading (B/L) để truyền danh sách container tờ khai. Khi xe đầu kéo đến cổng nhận hàng, hệ thống cổng giám sát tự động VASSCM phát hiện số seal thực tế trên cont khác hoàn toàn với số seal khai báo trên tờ khai điện tử.",
        "damage": "Hải quan phát lệnh dừng thông quan khẩn cấp; chuyển phân luồng từ luồng Vàng sang luồng Đỏ kiểm hóa 100%; DN bị xử phạt 5.000.000 đồng hành vi khai sai thông tin quản lý hải quan, chịu phí nâng hạ rút ruột kiểm hóa 8.500.000 đồng, và phát sinh 14.500.000 đồng tiền lưu bãi/lưu cont suốt 6 ngày chờ xin trích lục biên bản thay seal từ Cát Lái. Tổng thiệt hại 28.000.000 đồng.",
        "root_cause": "Nhân viên hiện trường và người khai hải quan chủ quan không kiểm tra tình trạng thực tế của container tại bãi cảng đích trước khi truyền tờ khai nhận cont; không nắm được quy định bắt buộc phải truyền theo số seal mới nhất kèm biên bản sự cố của cảng chuyển tải.",
        "prevention": "1. Trước khi truyền danh sách cont lấy hàng, luôn yêu cầu nhân viên hiện trường hoặc lái xe chụp ảnh đối chiếu trực tiếp chốt seal thực tế trên cánh cửa container; 2. Nếu phát hiện số seal khác B/L, liên hệ ngay bộ phận Điều độ cảng để lấy 'Biên bản ghi nhận sự cố thay seal'; 3. Truyền tờ khai theo số seal mới và đính kèm bản scan biên bản cảng lên hệ thống v5 tại mục Chứng từ đính kèm (HYS).",
        "opposing_views": "Phe A: Doanh nghiệp hoàn toàn bị động vì cảng tự bấm seal mới mà không gửi thông báo kịp thời cho chủ hàng; Phe B: Hải quan thực hiện đúng nguyên tắc nghiệp vụ, bất kỳ sự sai lệch số seal nào giữa tờ khai và hàng hóa thực tế đều phải đặt trong diện nghi vấn tráo hàng lậu và buộc phải kiểm hóa 100%.",
        "lessons": "Trong logistics chính ngạch, 'Seal là con dấu sinh mệnh của container'. Mọi sự xê dịch, gãy đứt hoặc đổi số seal đều phải được bảo đảm bằng biên bản có chữ ký của cảng và hãng tàu, không bao giờ được phép khai báo theo thói quen giấy tờ.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Case study người thật việc thật ghi nhận từ tình huống thảo luận nghiệp vụ chuyển cảng Cát Lái - Hiệp Phước tại daily_notes.md ngày 25/09/2026 & đối chiếu Thông tư 38/2015/TT-BTC.",
        "keywords": ["seal", "security"]
    },
    {
        "stt": 107,
        "week": "W40 — 28/09/2026",
        "title": "3 Container Túi Vô Trùng Aseptic Bag Nhập Khẩu Bị Ách Tắc 35 Ngày Tại Cảng Do Xung Đột Thủ Tục Kiểm Tra ATTP: Thiệt Hại 195 Triệu Đồng (Aseptic Bag Import Crisis)",
        "parties": "Công ty CP Chế biến Nông sản Xuất khẩu Miền Nam, Chi cục Hải quan Cửa khẩu Cảng Sài Gòn KV1, Viện Kiểm nghiệm An toàn Vệ sinh Thực phẩm Quốc gia",
        "plot": "Doanh nghiệp nhập khẩu 3 container 40 feet túi vô trùng (Aseptic bag, HS 3923.21.99) từ Tây Ban Nha dùng để chiết rót nước ép thanh long xuất khẩu sang EU. Doanh nghiệp đã chủ động hoàn thành thủ tục Bản Tự công bố sản phẩm theo Điều 4 Nghị định 15/2018/NĐ-CP và nộp cho Hải quan. Tuy nhiên, khi mở tờ khai, Chi cục Hải quan yêu cầu doanh nghiệp phải nộp thêm 'Giấy đăng ký kiểm tra nhà nước về ATTP' có xác nhận của cơ quan kiểm tra được Bộ Y tế chỉ định theo Thông tư 15/2024/TT-BYT. Khi doanh nghiệp mang hồ sơ sang trung tâm kiểm nghiệm thì trung tâm lúng túng từ chối tiếp nhận vì cho rằng Thông tư 27/2026/TT-BYT mới đã xếp bao bì tiếp xúc thực phẩm vào nhóm rủi ro trung bình chỉ cần tự công bố.",
        "damage": "3 container hàng nằm phơi nắng tại bãi cảng Cát Lái suốt 35 ngày để chờ văn bản xin ý kiến chỉ đạo giữa Cục Hải quan và Cục ATTP (Bộ Y tế); doanh nghiệp phát sinh 195.000.000 đồng chi phí lưu bãi, lưu container (DEM/DET) và suýt bị đối tác quốc tế hủy hợp đồng xuất khẩu nước ép trị giá 4.2 tỷ đồng do thiếu bao bì chiết rót.",
        "root_cause": "Sự giao thoa và áp dụng thiếu nhất quán giữa Thông tư 27/2026/TT-BYT và Thông tư 15/2024/TT-BYT khiến cơ quan kiểm tra chuyên ngành và hải quan không thống nhất được quy trình; doanh nghiệp không chủ động tham vấn văn bản nghiệp vụ hướng dẫn trước khi hàng về cảng.",
        "prevention": "1. Trước khi nhập khẩu bao bì tiếp xúc thực phẩm, gửi công văn xin xác nhận cụ thể của cơ quan kiểm tra chuyên ngành và Chi cục Hải quan mở tờ khai; 2. Chuẩn bị song song cả Bản Tự công bố và Giấy đăng ký kiểm tra nhà nước về ATTP nhập khẩu; 3. Sử dụng Công văn 22082/CHQ-GSQL và Công văn 1913/ATTP-KN làm căn cứ pháp lý để yêu cầu cơ quan kiểm tra tiếp nhận hồ sơ ngay từ ngày đầu.",
        "opposing_views": "Phe A: Doanh nghiệp bị oan do các cơ quan nhà nước 'đá bóng trách nhiệm' và văn bản chồng chéo; Phe B: Hải quan phải giữ vững nguyên tắc quản lý an toàn thực phẩm, bao bì trực tiếp không qua kiểm nghiệm có thể gây ngộ độc hàng loạt cho người tiêu dùng.",
        "lessons": "Nhập khẩu hàng hóa liên quan đến sức khỏe người tiêu dùng (ATTP, bao bì thực phẩm) là 'vùng biển có nhiều đá ngầm pháp lý'. Doanh nghiệp không được suy đoán văn bản theo hướng có lợi cho mình mà phải luôn chuẩn bị sẵn sàng cho mức độ kiểm soát cao nhất.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "series": "Series: Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Tình huống thực tế xử lý tranh chấp kiểm tra ATTP bao bì túi Aseptic bag từ daily_notes.md ngày 24/09/2026 & Công văn 22082/CHQ-GSQL.",
        "keywords": ["food", "safety"]
    },
    {
        "stt": 108,
        "week": "W40 — 28/09/2026",
        "title": "Phòng Lab Thử Nghiệm Quá Tải Khiến Lô Hàng Điện Tử Quá Hạn 30 Ngày Bảo Quản: Doanh Nghiệp Bị Phạt 20 Triệu & Mất Quyền Đưa Hàng Về Kho 1 Năm (Storage Deadline Trap)",
        "parties": "Công ty Thiết bị Điện Gia Dụng Hải Phòng, Chi cục Hải quan Quản lý Hàng Đầu tư Gia công, Chi cục Tiêu chuẩn Đo lường Chất lượng, Trung tâm Thử nghiệm Kỹ thuật 1 (Quatest 1)",
        "plot": "Tháng 08/2026, doanh nghiệp nhập khẩu 2 container lò nướng và máy xay sinh tố, làm thủ tục xin đưa hàng về kho riêng bảo quản theo Điều 32 Thông tư 38 để chờ kết quả chứng nhận hợp quy của Quatest 1. Do đợt cao điểm phòng lab quá tải và phương pháp thử nghiệm độ bền nhiệt đòi hỏi vận hành liên tục 500 giờ, Quatest 1 thông báo phải mất 40 ngày mới trả phiếu kết quả. Đến ngày thứ 28, doanh nghiệp làm công văn xin Chi cục TCĐLCL xác nhận thời gian kiểm tra kéo dài nhưng bị cán bộ từ chối ký vì 'Chi cục chỉ tiếp nhận hồ sơ, việc thử nghiệm là của Quatest'. Quá ngày thứ 30, Chi cục Hải quan lập biên bản vi phạm hành chính đối với doanh nghiệp vì quá thời hạn nộp kết quả kiểm tra chuyên ngành.",
        "damage": "Doanh nghiệp bị xử phạt vi phạm hành chính 20.000.000 đồng theo điểm d khoản 2 Điều 18 Nghị định 128/2020/NĐ-CP; nghiêm trọng hơn, doanh nghiệp bị hệ thống tự động khóa quyền đưa hàng về kho bảo quản trong suốt 12 tháng tiếp theo. 5 lô hàng nhập khẩu sau đó trị giá 8.5 tỷ đồng bắt buộc phải nằm tại cảng Đình Vũ chờ kết quả test, phát sinh thêm 210.000.000 đồng tiền lưu container và lưu bãi.",
        "root_cause": "Doanh nghiệp chủ quan nghĩ rằng nộp giấy đăng ký có mộc là đủ; không lường trước thời gian thử nghiệm thực tế của phòng lab; không chủ động yêu cầu phòng lab cấp văn bản cam kết tiến độ thử nghiệm để kẹp vào công văn giải trình với cơ quan hải quan trước thời hạn ngày thứ 30.",
        "prevention": "1. Trước khi xin mang hàng về bảo quản, phải ký phụ lục cam kết thời gian trả kết quả với phòng thử nghiệm dưới 25 ngày; 2. Nếu thời gian thử nghiệm kéo dài > 30 ngày, phải nộp văn bản đề nghị gia hạn cho Hải quan ngay từ ngày thứ 20, đính kèm Hợp đồng thử nghiệm và Thông báo tiến độ của phòng lab; 3. Nếu nghi ngờ phòng lab quá tải, chủ động phân bổ mẫu sang các phòng thử nghiệm tư nhân được chỉ định khác có tiến độ nhanh hơn.",
        "opposing_views": "Phe A: Doanh nghiệp là nạn nhân của sự quá tải phòng lab và sự vô cảm của cơ quan cấp phép; Phe B: Hải quan buộc phải xử phạt nghiêm minh theo luật định vì nếu nới lỏng thời hạn 30 ngày mà không có kiểm soát thì kho bảo quản của doanh nghiệp sẽ biến thành kho tẩu tán hàng lậu.",
        "lessons": "Đưa hàng về kho bảo quản là một 'đặc ân pháp lý' chứ không phải quyền mặc định. Một khi đã sử dụng cơ chế này, doanh nghiệp phải quản trị thời gian chặt chẽ như đếm ngược từng ngày, vì cái giá của việc mất quyền bảo quản đắt gấp hàng chục lần tiền phạt hành chính.",
        "url": "https://en.wikipedia.org/wiki/Bonded_warehouse",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Case study điển hình về rủi ro thủ tục chuyên ngành đúc kết từ tình huống Thông tư 121 tại daily_notes.md ngày 25/09/2026 & đối chiếu Nghị định 128/2020/NĐ-CP.",
        "keywords": ["bonded", "warehouse"]
    },
    {
        "stt": 109,
        "week": "W40 — 28/09/2026",
        "title": "Nhập Khẩu 20 Tấn Hành Sấy Nhiệt Áp Nhầm Mã Miễn Thuế VK010: Doanh Nghiệp Bị Hải Quan Truy Thu Thuế VAT & Phạt Gần 95 Triệu Đồng (Dried Scallion VAT Audit)",
        "parties": "Công ty TNHH Gia Vị & Nông Sản Thực Phẩm Hà Nội, Chi cục Hải quan Cửa khẩu Cảng Hải Phòng KV3, Cục Kiểm định Hải quan",
        "plot": "Doanh nghiệp nhập khẩu 20 tấn 'Hành lá sấy nhiệt, đã cắt vụn, đóng gói kín' (HS 0712.90.90) từ Sơn Đông (Trung Quốc) về làm nguyên liệu gia vị mì tôm. Căn cứ khoản 1 Điều 4 Nghị định 181/2025/NĐ-CP (quy định nông sản chỉ qua sơ chế thông thường ở khâu nhập khẩu thuộc đối tượng không chịu thuế VAT), người khai hải quan đã chọn mã thuế suất 'V' và chọn mã đối tượng không chịu thuế là 'VK010'. Khi hồ sơ vào luồng Vàng, cán bộ Hải quan yêu cầu giải trình quy trình kỹ thuật. Doanh nghiệp cung cấp quy trình của nhà sản xuất ghi rõ: 'Rửa sạch -> Cắt vụn -> Sấy nhiệt lò đối lưu ở 95°C trong 45 phút'. Hải quan kết luận: Việc xử lý nhiệt độ cao 95°C đã làm thay đổi trạng thái hóa sinh của tế bào thực vật, biến hành từ nông sản tươi sống sơ chế thành sản phẩm chế biến công nghiệp, do đó không được áp mã VK010 mà phải chịu thuế suất VAT 10%.",
        "damage": "Doanh nghiệp bị bác bỏ mã VK010, bị truy thu 10% thuế VAT trên trị giá lô hàng 720 triệu đồng (tương đương 72.000.000 đồng); bị xử phạt vi phạm hành chính 20% số thuế khai thiếu (14.400.000 đồng) và chịu tiền chậm nộp 7.800.000 đồng. Tổng số tiền phải nộp thêm là 94.200.000 đồng, đồng thời toàn bộ các lô hàng sau bị phân luồng Đỏ kiểm tra thực tế.",
        "root_cause": "Doanh nghiệp đánh đồng khái niệm 'sơ chế thông thường' trong luật thuế với mọi phương pháp sấy khô; không tìm hiểu kỹ hướng dẫn kỹ thuật của Tổng cục Hải quan về ranh giới nhiệt độ và công nghệ sấy nông sản trước khi truyền tờ khai.",
        "prevention": "1. Trước khi nhập khẩu nông sản dạng sấy khô/khử nước, yêu cầu nhà cung cấp nước ngoài cấp Bản mô tả chi tiết quy trình công nghệ sản xuất (Production Flowchart); 2. Đối chiếu quy trình với Thông tư 219/2013/TT-BTC: nếu có gia nhiệt > 60°C hoặc dùng máy sấy công nghiệp, chủ động xin ý kiến hướng dẫn nghiệp vụ bằng văn bản từ Chi cục Hải quan; 3. Nếu chưa chắc chắn, tạm thời nộp thuế VAT 5% hoặc 10% để thông quan nhanh, sau đó làm thủ tục xin hoàn thuế hoặc hoàn nhập chi phí đầu vào.",
        "opposing_views": "Phe A: Doanh nghiệp cho rằng hành sấy khô chỉ rút nước chứ không tẩm ướp phụ gia nào, bản chất vẫn là nông sản thô nên xứng đáng được hưởng ưu đãi thuế; Phe B: Hải quan khẳng định sấy nhiệt lò công nghiệp là công đoạn chế biến tạo giá trị gia tăng, phải chịu thuế VAT để bình đẳng với các nhà máy sấy nông sản nội địa.",
        "lessons": "Ưu đãi thuế là con dao hai lưỡi. Khai sai một mã miễn thuế (như VK010) không chỉ khiến doanh nghiệp bị truy thu số thuế gốc mà còn kéo theo án phạt trốn/thiếu thuế và làm hỏng toàn bộ lịch sử tuân thủ hải quan của công ty.",
        "url": "https://en.wikipedia.org/wiki/Value-added_tax",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Tình huống thực chiến từ case study hành sấy nhiệt HS 07129090 tại daily_notes.md ngày 25/09/2026 & đối chiếu Thông tư 219/2013/TT-BTC.",
        "keywords": ["value", "tax"]
    },
    {
        "stt": 110,
        "week": "W40 — 28/09/2026",
        "title": "Có Đầy Đủ Ủy Nhiệm Chi Ngân Hàng Nhưng Vẫn Bị Truy Thu 420 Triệu Đồng Thuế Vì Đối Tác Bán Hàng Trở Thành 'Công Ty Ma' (Ghost Company Invoice Nightmare)",
        "parties": "Công ty TNHH Cơ Khí & Thương Mại Vận Tải Đông Đô, Chi cục Thuế Thành phố Bắc Ninh, Nhà cung cấp vật tư kim khí X (đã giải thể bỏ trốn)",
        "plot": "Năm 2024 - 2025, Công ty Đông Đô mua hơn 3.5 tỷ đồng sắt thép và phụ tùng từ Nhà cung cấp X. Đông Đô thực hiện thanh toán 100% qua tài khoản ngân hàng của Công ty X, nhận đủ hóa đơn điện tử hợp lệ có mã của cơ quan thuế và kê khai khấu trừ thuế VAT bình thường. Đến tháng 09/2026, Chi cục Thuế gửi thông báo yêu cầu Đông Đô giải trình vì Nhà cung cấp X nằm trong danh sách 524 doanh nghiệp rủi ro cao bị công an điều tra tội mua bán trái phép hóa đơn và người đại diện pháp luật đã bỏ trốn khỏi địa phương. Khi cơ quan thuế yêu cầu xuất trình hồ sơ chứng minh giao dịch thực tế, Đông Đô chỉ đưa ra được Hợp đồng mua bán sơ sài và Giấy báo Nợ ngân hàng; toàn bộ phiếu cân hàng, biên bản giao nhận kho, biển số xe tải vận chuyển và log theo dõi hàng hóa đều không được lưu trữ.",
        "damage": "Cơ quan thuế bác bỏ toàn bộ 3.5 tỷ đồng chi phí đầu vào và 350.000.000 đồng thuế VAT đầu vào; truy thu 350 triệu VAT và 70 triệu thuế TNDN (tổng cộng 420.000.000 đồng), cộng thêm khoản tiền phạt vi phạm và tiền chậm nộp hơn 85.000.000 đồng. Tổng số tiền thiệt hại lên tới 505.000.000 đồng, làm suy kiệt dòng tiền lưu động của doanh nghiệp.",
        "root_cause": "Doanh nghiệp mắc sai lầm nghiêm trọng khi tin rằng 'Chỉ cần có hóa đơn điện tử và chuyển khoản ngân hàng là an toàn 100%'. Doanh nghiệp hoàn toàn xem nhẹ việc lưu trữ chuỗi bằng chứng logistics thực tế (phiếu cân, biên bản giao nhận, thông tin lái xe, camera kho) để chứng minh dòng hàng vật chất tương ứng với dòng tiền.",
        "prevention": "1. Trước khi ký hợp đồng lớn, tra cứu tình trạng hoạt động của người nộp thuế trên tracuunnt.gdt.gov.vn và kiểm tra thực tế địa chỉ kho bãi bên bán; 2. Thiết lập quy trình 'Hồ sơ 4 lớp' cho mọi lô hàng mua vào: Hợp đồng + Hóa đơn + Chuyển khoản ngân hàng + Trọn bộ chứng từ logistics (Biên bản giao nhận có chữ ký lái xe, phiếu cân điện tử, hình ảnh hàng tại kho, biển số xe vận chuyển); 3. Số hóa quy trình quản lý kho và giao nhận trên phần mềm ERP như cách Eureka quản lý vết đơn hàng để sẵn sàng trích xuất đối chất bất kỳ lúc nào.",
        "opposing_views": "Phe A: Doanh nghiệp bức xúc vì đã thanh toán tiền thật qua ngân hàng và hóa đơn có mã của cơ quan thuế mà vẫn bị phạt; Phe B: Cơ quan thuế khẳng định nhiều đường dây mua bán hóa đơn sử dụng thủ đoạn 'rửa tiền' chuyển khoản vòng vo qua ngân hàng rồi rút tiền mặt trả lại, do đó bắt buộc phải có bằng chứng chứng minh hàng hóa thực sự di chuyển mới được công nhận chi phí.",
        "lessons": "Trong quản trị hiện đại, 'Hóa đơn và Dòng tiền chỉ là cái bóng, Dòng hàng hóa thực tế mới là cái cây'. Doanh nghiệp nào xây dựng được hệ thống dữ liệu logistics và log vận hành ERP minh bạch sẽ sở hữu 'tấm áo giáp chống đạn' bảo vệ công ty trước mọi cơn bão thanh tra thuế hay bóc phốt truyền thông.",
        "url": "https://en.wikipedia.org/wiki/Invoice",
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Case study sâu sắc về rủi ro hóa đơn công ty ma đối chiếu với triết lý quản trị bằng chứng ERP của Eureka tại daily_notes.md ngày 24-25/09/2026.",
        "keywords": ["invoice"]
    }
]
