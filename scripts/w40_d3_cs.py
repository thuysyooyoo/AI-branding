# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W40 DAY 3 — 30/09/2026)
Số lượng: 5 ý tưởng (STT 116 - 120)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, số liệu thiệt hại cụ thể, URL thật đã qua url_validator.py
"""

CASE_STUDY_W40_D3 = [
    {
        "stt": 116,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Gia Công Biến Tính Hạt Nhựa Nhập Khẩu Nhưng Không Đổi Mã HS: Doanh Nghiệp Chế Xuất Bị Truy Thu 420 Triệu Đồng Thuế Do Khai Sai Xuất Xứ Hàng Hóa (Substantial Transformation Violation Trap)",
        "stakeholders": "Doanh nghiệp Chế xuất (DNCX) Nhựa Kỹ thuật KCN Quế Võ (Bắc Ninh) vs Xưởng gia công nội địa vs Chi cục Hải quan Bắc Ninh",
        "context": "Doanh nghiệp chế xuất nhập khẩu 100 tấn hạt nhựa ABS (HS 390330) xuất xứ Trung Quốc theo loại hình E11 để phục vụ sản xuất. Do công suất nhà máy hạn chế, DN thuê xưởng FDI nội địa gia công biến tính (pha trộn phụ gia tăng độ chống cháy và độ cứng). Sau gia công, sản phẩm vẫn mang bản chất hạt nhựa ABS, mã HS không thay đổi (vẫn 390330), giá trị gia tăng RVC chỉ đạt 18% (< 30%). Khi lập hồ sơ thanh khoản và chuyển thành phẩm về kho DNCX, kế toán tự ý khai báo thành phẩm là 'Made in Vietnam' để xuất đi nhà máy lắp ráp điện tử. Tháng 09/2026, Hải quan kiểm tra sau thông quan phát hiện công đoạn gia công không đáp ứng tiêu chí chuyển đổi cơ bản.",
        "damage": "Bị truy thu 420 triệu đồng tiền thuế nhập khẩu ưu đãi bị hưởng sai; xử phạt vi phạm hành chính 84 triệu đồng (20% số tiền thuế khai thiếu); tiền chậm nộp 55 triệu đồng; hạ bậc tuân thủ doanh nghiệp từ hạng 2 xuống hạng 4 (buộc kiểm hóa 100% tất cả các lô hàng tiếp theo).",
        "root_cause": "Doanh nghiệp nhầm lẫn tai hại giữa việc 'gia công có phát sinh chi phí tại Việt Nam' với việc 'đạt tiêu chuẩn xuất xứ hàng hóa Việt Nam'; thiếu kiến thức về quy tắc chuyển đổi mã số hàng hóa (CTC) và hàm lượng giá trị gia tăng (RVC) theo Nghị định 31/2018/NĐ-CP.",
        "solution": "Checklist 3 bước: 1. Đối chiếu mã HS đầu vào vs đầu ra: Nếu không đổi mã HS cấp 4-6 số, bắt buộc tính RVC; 2. Nếu RVC < 30%, tuyệt đối không được ghi Made in Vietnam, phải giữ nguyên xuất xứ nguyên liệu gốc (Made in China) hoặc ghi 'Xuất xứ không xác định'; 3. Đưa điều khoản trách nhiệm khai báo xuất xứ chuẩn vào Hợp đồng gia công nội địa.",
        "opposing_views": "Phe A (Chủ hàng & Bên gia công): DN đã bỏ tiền đầu tư công nghệ, phụ gia đắt đỏ để biến đổi cơ lý tính sản phẩm thì sản phẩm phải được công nhận một phần giá trị tại VN. Phe B (Hải quan): Quy định xuất xứ phải dựa trên tiêu chí định lượng minh bạch (CTC, RVC); nếu ai cũng tự nhận xuất xứ thì Việt Nam sẽ thành bãi đáp trung chuyển 'rửa xuất xứ' cho hàng ngoại.",
        "lesson": "Xuất xứ hàng hóa là bài toán kỹ thuật pháp lý khắt khe, không thể suy diễn theo cảm tính. Gia công dù tốn kém đến đâu nhưng không vượt qua rào cản CTC/RVC thì về mặt pháp lý vẫn mang quốc tịch của nguyên liệu gốc.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (vướng mắc gia công hạt nhựa biến tính ngày 30/09/2026) & Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP).",
        "keywords": ["rules", "origin"]
    },
    {
        "stt": 117,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "C/O Form AI 3 Bên Bị Bác Bỏ Tại Cảng Cát Lái: Doanh Nghiệp Mất 650 Triệu Đồng Thuế Ưu Đãi Vì Ô Số 9 & Ô Số 10 Thể Hiện Cả 2 Trị Giá FOB Và 2 Số Hóa Đơn (Third-Party Invoicing Discrepancy)",
        "stakeholders": "Công ty TNHH Nhập Khẩu Hóa Chất Nam Việt (TP.HCM) vs Nhà xuất khẩu Ấn Độ vs Trader trung gian Singapore vs Chi cục Hải quan Cửa khẩu Cảng Sài Gòn KV1 (Cát Lái)",
        "context": "Nam Việt nhập khẩu 5 container hóa chất tẩy rửa công nghiệp từ Ấn Độ trị giá 180.000 USD, ký hợp đồng mua bán qua một công ty thương mại tại Singapore. Cơ quan cấp C/O tại Ấn Độ cấp C/O Form AI điện tử, tại ô số 9 ghi cả trị giá FOB của nhà sản xuất (150.000 USD) và trị giá của bên thứ ba (180.000 USD); ô số 10 ghi đồng thời số hóa đơn của nhà máy Ấn Độ và số hóa đơn thương mại Singapore. Tuy nhiên, khi hải quan Cát Lái tra cứu trên Cổng eCOSys thì dữ liệu điện tử chỉ truyền 1 số hóa đơn của nhà máy Ấn Độ. Hải quan nghi ngờ tính xác thực của hóa đơn bên thứ ba, tạm dừng áp dụng thuế ưu đãi AIFTA 0% và yêu cầu nộp bảo lãnh theo thuế MFN 10%.",
        "damage": "Doanh nghiệp phải nộp 650 triệu đồng tiền thuế MFN và bảo lãnh để mang hàng về; phát sinh 45 triệu đồng phí lưu container và bãi (demurrage/detention) trong 12 ngày chờ xác minh từ cơ quan cấp C/O nước ngoài.",
        "root_cause": "Nhà xuất khẩu và công ty thương mại Singapore không hiểu rõ chuẩn khai báo C/O bên thứ ba; tích chọn Third-party invoicing nhưng lại thể hiện chồng chéo cả thông tin nội bộ giữa nhà máy và trader lên cùng một form C/O, gây xung đột với dữ liệu truyền điện tử trên hệ thống hải quan.",
        "solution": "Quy trình chuẩn: 1. Với C/O 3 bên, ô số 10 bắt buộc phải ghi ĐÚNG số và ngày hóa đơn thương mại của bên thứ ba (invoice nộp cho hải quan); 2. Ô số 7 phải thể hiện rõ tên và địa chỉ của công ty phát hành hóa đơn bên thứ ba; 3. Ô số 13 bắt buộc tích 'Third-party invoicing'; 4. Yêu cầu bên bán gửi bản draft C/O để Forwarder/đại lý hải quan chuyên nghiệp kiểm tra trước khi cấp bản chính thức.",
        "opposing_views": "Phe A (Doanh nghiệp): Hàng hóa đi thẳng từ Ấn Độ về Cát Lái, C/O do cơ quan có thẩm quyền Ấn Độ cấp thật 100%, việc ghi thêm thông tin nhà máy là để minh bạch chứ không gian lận. Phe B (Hải quan): Dữ liệu C/O điện tử phải tuyệt đối đồng nhất với chứng từ thanh toán; nếu ô 10 có 2 số hóa đơn nhưng hệ thống chỉ có 1 thì hải quan có quyền nghi vấn chuyển giá hoặc hồ sơ giả mạo.",
        "lesson": "Một chi tiết thừa trên C/O còn nguy hiểm hơn một chi tiết thiếu. Đừng bao giờ để hai số hóa đơn khác nhau cùng xuất hiện trên ô số 10 nếu không có hướng dẫn cụ thể từ hiệp định thương mại.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Series: Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (case C/O Form AI 3 bên ngày 30/09/2026) & Thông tư 15/2010/TT-BCT sửa đổi bởi Thông tư 08/2024/TT-BCT.",
        "keywords": ["certificate", "origin"]
    },
    {
        "stt": 118,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Container Rời Cảng Bị Gãy Seal Phải Bấm Lại Tại Bãi: Doanh Nghiệp Bị Dừng Thông Quan 4 Ngày, Tốn 38 Triệu Đồng Phí Lưu Bãi Do Không Có Biên Bản Sự Cố Của Cảng (Broken Seal Handover Failure)",
        "stakeholders": "Công ty CP Logistics Toàn Cầu vs Cảng Cát Lái vs Cảng Tân Cảng Hiệp Phước vs Hải quan Giám sát Tân Cảng Hiệp Phước",
        "context": "Doanh nghiệp vận chuyển 3 container hàng máy móc chuyển cảng từ Cát Lái sang Tân Cảng Hiệp Phước bằng sà lan. Quá trình cẩu hạ tại bãi Cát Lái, dây cẩu va quệt làm gãy seal gốc của hãng tàu trên 1 container. Nhân viên điều độ cảng đã bấm seal mới của Cảng Tân Cảng vào để niêm phong lại cont nhưng tài xế và đơn vị vận tải không yêu cầu Cảng lập 'Biên bản ghi nhận sự cố thay seal'. Khi hàng đến Hiệp Phước, chủ hàng truyền danh sách container theo số seal gốc của hãng tàu trên tờ khai. Hệ thống VASSCM tại cổng phát hiện số seal thực tế khác với số seal tờ khai, lập tức báo động chặn xe, dừng thông quan.",
        "damage": "Lô hàng bị ách tắc 4 ngày tại bãi; phát sinh 38 triệu đồng chi phí lưu bãi, lưu xe cẩu và tiền phạt chậm tiến độ giao máy cho nhà máy; doanh nghiệp phải trả thêm 15 triệu đồng phí bốc xếp kiểm hóa 100% khi hải quan nghi ngờ hàng bị đánh tráo.",
        "root_cause": "Đơn vị vận tải tắc trách, thiếu quy trình chuẩn (SOP) xử lý sự cố đứt gãy seal; không lấy Biên bản sự cố của đơn vị bấm seal thay thế; người làm thủ tục truyền tờ khai theo thói quen cũ mà không kiểm tra số seal thực tế của container tại bãi đến.",
        "solution": "Checklist 3 bước khi chuyển cảng: 1. Mọi trường hợp thay seal bắt buộc phải có Biên bản sự cố có chữ ký đóng dấu của Cảng hoặc Hải quan giám sát; 2. Khi truyền tờ khai lấy cont, bắt buộc khai theo SỐ SEAL MỚI đang bấm trên cont; 3. Nộp đính kèm Biên bản sự cố lên Cổng thông tin một cửa quốc gia trước khi cho xe ra khỏi cổng.",
        "opposing_views": "Phe A (Chủ hàng & Vận tải): Hàng được sà lan cảng chuyên trách vận chuyển nguyên niêm phong, seal gãy do va đập ngoài ý muốn chứ hàng hóa bên trong còn nguyên vẹn 100%. Phe B (Hải quan): Seal là căn cứ pháp lý duy nhất để bảo đảm tính nguyên trạng hàng hóa; seal khác mà không có biên bản chứng minh thì hải quan bắt buộc phải kiểm tra thực tế để đề phòng buôn lậu.",
        "lesson": "Trong vận tải container chính ngạch, chiếc seal kim loại nhỏ bé mang giá trị pháp lý tương đương với tờ khai hải quan. Sai một con số seal có thể khiến cả lô hàng trị giá hàng chục tỷ bị giam giữ vô thời hạn.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Series: Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (sự cố gãy seal chuyển cảng ngày 25/09/2026) & Thông tư 39/2018/TT-BTC.",
        "keywords": ["security", "seal"]
    },
    {
        "stt": 119,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Mua Hàng Thật Có Chuyển Khoản Nhưng Nhà Cung Cấp Bỏ Trốn: Doanh Nghiệp Phân Phối Thiết Bị Điện Bị Truy Thu 1,2 Tỷ Đồng Thuế VAT & TNDN (Shell Company Ghost Invoice Trap)",
        "stakeholders": "Công ty TNHH Phân Phối Thiết Bị Điện H.T (Hà Nội) vs Công ty TNHH Thương mại Dịch vụ Nam Long (Hải Phòng - Công ty ma) vs Cục Thuế TP. Hà Nội",
        "context": "Trong giai đoạn 2024-2025, Công ty H.T mua 12 tỷ đồng dây cáp và thiết bị đóng cắt từ Nam Long. Hàng hóa giao nhận đầy đủ, hóa đơn điện tử tra cứu hợp lệ trên Cổng Thuế, tiền thanh toán 100% qua tài khoản ngân hàng của Nam Long. Đến tháng 09/2026, cơ quan thuế thanh tra phát hiện Nam Long là 'công ty ma' được thành lập bởi các đối tượng mua bán hóa đơn, giám đốc đứng tên là người chạy xe ôm và công ty đã bỏ trốn khỏi địa chỉ kinh doanh. Cơ quan thuế ra quyết định không công nhận chi phí và loại bỏ toàn bộ số thuế VAT đầu vào của các hóa đơn này.",
        "damage": "Công ty H.T bị truy thu 1,2 tỷ đồng thuế VAT; truy thu 2,4 tỷ đồng thuế TNDN do bị loại chi phí; phạt vi phạm hành chính 20% (720 triệu đồng) và tiền chậm nộp 380 triệu đồng. Tổng tổn thất tài chính lên tới 4,7 tỷ đồng, đẩy doanh nghiệp vào bờ vực phá sản.",
        "root_cause": "Doanh nghiệp quá tin tưởng vào việc 'hóa đơn có trên hệ thống thuế và có ủy nhiệm chi ngân hàng'; không thực hiện quy trình thẩm định thực tế đối tác (KYC/KYS); không lưu trữ hồ sơ chứng minh logistics thực tế (vận đơn xe tải chở hàng, biên bản giao nhận kho, danh tính tài xế giao hàng).",
        "solution": "Bộ cẩm nang phòng hộ 4 lớp: 1. Trước khi ký hợp đồng, cử nhân viên đến tận kho/trụ sở đối tác chụp ảnh xác minh cơ sở vật chất; 2. Định kỳ hàng tháng tra cứu trạng thái hoạt động của đối tác trên Cổng thông tin Tổng cục Thuế; 3. Bắt buộc lưu trữ trọn bộ 'Hồ sơ 4 điểm': Hợp đồng + Hóa đơn điện tử + Ủy nhiệm chi ngân hàng + Chứng từ vận tải giao nhận thực tế có biển số xe và chữ ký thủ kho; 4. Giữ lại bảo lãnh bảo hành 5-10% thanh toán sau 6 tháng.",
        "opposing_views": "Phe A (Doanh nghiệp): Chúng tôi mua hàng thật, bán thật, tiền chuyển qua ngân hàng đàng hoàng; cơ quan thuế thu thuế của bên bán không được lại đi đè bên mua ra phạt là bất công. Phe B (Cơ quan Thuế): Nhiều doanh nghiệp dùng thủ đoạn mua hàng trôi nổi ngoài thị trường chợ đen rồi thông đồng mua hóa đơn của công ty ma hợp thức hóa đầu vào; nếu không có chứng từ giao nhận logistics thực tế thì không thể công nhận giao dịch là có thật.",
        "lesson": "Hóa đơn điện tử và ủy nhiệm chi ngân hàng không còn là 'kim bài miễn tử'. Trong thời đại kiểm soát thuế số hóa, hồ sơ chứng từ logistics thực tế chính là lá chắn thép duy nhất bảo vệ doanh nghiệp trước rủi ro công ty ma.",
        "url": "https://en.wikipedia.org/wiki/Tax_evasion",
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (case rủi ro hóa đơn công ty ma ngày 25/09/2026) & Luật Quản lý thuế 38/2019/QH14.",
        "keywords": ["tax", "evasion"]
    },
    {
        "stt": 120,
        "week": "W40 Day 3 — 30/09/2026",
        "title": "Khủng Hoảng 'Hàng Về Lâu Lắc' Do Quên Mã Khách Hàng (Customer Code): Eureka Phản Pháo Bằng Dữ Liệu ERP Hệ Thống & Cú Lội Ngược Dòng Khi Khách Ruột Thành Đại Sứ Bảo Vệ (ERP System Turnaround)",
        "stakeholders": "Khách hàng Phan Tran vs Jasmine Nguyen (CEO Eureka) vs Khách hàng ruột (Ẩn danh 558) trên nhóm Cộng Đồng Nhập Khẩu Chính Ngạch Trung Quốc",
        "context": "Một khách hàng đăng bài bóc phốt trên nhóm 150.000 thành viên tố Eureka vận chuyển hàng chậm chạp và tìm đối tác khác. Thay vì đôi co, Chị Thúy (CEO Eureka) bình tĩnh tiếp nhận, cam kết bồi thường 100% nếu Eureka sai, đồng thời truy xuất trực tiếp lịch sử đơn hàng trên hệ thống ERP: Phát hiện khách gửi 2 đơn hàng về kho Trung Quốc nhưng không ghi Mã Khách Hàng, hàng nằm chờ xác nhận 5 ngày khách mới báo. Chị Thúy công khai ảnh chụp vận đơn gốc không mã và log hệ thống, chỉ rõ nguyên tắc vận hành chính ngạch. Ngay sau đó, khách hàng ruột lâu năm đã vào đăng 'tâm thư' chia sẻ trải nghiệm thực tế về sự tận tâm của Sếp Thúy và sự chuẩn chỉ của nhân viên Eureka, biến khủng hoảng thành cơ hội xây dựng uy tín đỉnh cao.",
        "damage": "Nguy cơ mất hợp đồng vận chuyển định kỳ 300 triệu đồng/tháng và khủng hoảng uy tín thương hiệu Eureka trên mạng xã hội nếu không xử lý minh bạch; doanh nghiệp chủ hàng bị chậm hàng 7 ngày dẫn đến trễ hẹn bàn giao cho khách lẻ.",
        "root_cause": "Khách hàng quen thói quen đi tiểu ngạch vô tổ chức, không tuân thủ quy chuẩn mã định danh khách hàng (Customer Code); không báo trước thông tin hàng hóa để kho làm thủ tục kiểm đếm và hồ sơ xuất khẩu chính ngạch.",
        "solution": "Bộ quy chuẩn 3 bước: 1. Tuyệt đối không gửi bất kỳ kiện hàng nào về kho nước ngoài nếu chưa dán tem Mã Khách Hàng (Customer Code) và thông báo mã vận đơn nội địa; 2. Hệ thống ERP tự động gửi tin nhắn cảnh báo khi kiện hàng về kho không có mã định danh sau 24h; 3. Lập cam kết dịch vụ (SLA) quy định rõ trách nhiệm thời gian chỉ tính từ khi khách hoàn thiện đủ thông tin tờ khai.",
        "opposing_views": "Phe A (Chủ hàng đi gom cont): Tôi đã gửi hàng về địa chỉ kho của bạn thì việc tìm hàng là trách nhiệm của công ty dịch vụ; bắt ghi đủ thứ mã là làm khó khách hàng. Phe B (Đơn vị Chính ngạch): Mỗi ngày kho tiếp nhận hàng chục tấn hàng với hàng nghìn kiện; không có mã định danh thì không hệ thống nào quản lý nổi; làm chính ngạch bắt buộc phải chuẩn chỉnh từng quy trình.",
        "lesson": "Dữ liệu hệ thống ERP và sự chuẩn chỉ quy trình là vũ khí mạnh nhất để bảo vệ uy tín trước khủng hoảng. Nhưng tài sản quý giá nhất của một doanh nghiệp dịch vụ chính là niềm tin của khách hàng cũ - những người sẽ tự nguyện trở thành 'bức tường lửa' bảo vệ bạn khi có bão tố.",
        "url": "https://en.wikipedia.org/wiki/Enterprise_resource_planning",
        "series": "Series: Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (case study thực chiến xử lý khủng hoảng bóc phốt ngày 24/09/2026) & Brand Persona Eureka.",
        "keywords": ["enterprise", "resource"]
    }
]
