# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W42 DAY 1 — 14/10/2026)
Số lượng: 5 ý tưởng (STT 146 - 150)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W42_D1 = [
    {
        "stt": 146,
        "week": "W42 — 14/10/2026",
        "title": "Cẩm Nang 5 Bước Xác Định Tỷ Lệ Thành Phần Thịt Trong Bánh Thực Phẩm: Chốt Chuẩn Mã HS Chương 16 Hay Chương 19 Né Truy Thu Thuế Hàng Trăm Triệu",
        "workflow": "Bước 1: Yêu cầu Bảng công thức định lượng chuẩn từ Nhà máy sản xuất (Product Formulation & Specification Sheet): Yêu cầu nhà cung cấp liệt kê chi tiết tỷ lệ % khối lượng của từng thành phần trước và sau khi gia nhiệt; Bước 2: Bóc tách thành phần theo Chú giải 2 Chương 16: Tổng hợp khối lượng của tất cả các thành phần xúc xích, thịt, phụ phẩm giết mổ, tiết, cá, giáp xác (gọi chung là nhóm thịt); Bước 3: Áp dụng công thức tính tỷ lệ % thịt trên tổng khối lượng thành phẩm; Bước 4: Lập Bảng giải trình phân loại kỹ thuật gửi kèm hồ sơ hải quan: Trích dẫn rõ Chú giải 2 Chương 16 và Quy tắc phân loại 1, 3(b) hoặc 6 của WCO; Bước 5: Tham vấn phân loại trước (Advance Ruling) hoặc gửi mẫu thử nghiệm tại Chi nhánh Trung tâm Kiểm định Hải quan trước khi nhập khẩu hàng loạt.",
        "formula": "Tỷ lệ % Thịt = (Tổng khối lượng thành phần thịt và phụ phẩm / Khối lượng tịnh của 1 chiếc bánh hoàn chỉnh) x 100%. Nếu Tỷ lệ > 20% -> Bắt buộc áp Chương 16 (nhóm 16.02, thuế suất 30%); Nếu Tỷ lệ <= 20% -> Được áp Chương 19 (nhóm 19.05, thuế suất 15%).",
        "top3_mistakes": "1. Chỉ nhìn tên gọi trên vỏ hộp ('Bánh mì hamburger') mà tự tiện áp mã 19.05 mà không kiểm tra ruột nhân; 2. Tính tỷ lệ % theo trọng lượng thịt tươi trước chế biến thay vì khối lượng thực tế trong sản phẩm hoàn chỉnh; 3. Bỏ qua các thành phần phụ phẩm như mỡ động vật, da gà, chất độn protein động vật được xay lẫn trong nhân xúc xích.",
        "mnemonic_tip": "Nhớ quy tắc '20 Bánh - 80 Thịt': Dưới 20% thịt coi là Bánh (Chương 19 nhẹ thuế); Vượt 20% thịt coi là Thịt (Chương 16 thuế cao + kiểm dịch gắt).",
        "comparison": "Chương 19 (Nhóm 19.05): Thuế NK 15%, kiểm tra ATTP thông thường, thủ tục hồ sơ gọn nhẹ VS Chương 16 (Nhóm 16.02): Thuế NK 30%, kiểm dịch động vật chuyên ngành, lấy mẫu xét nghiệm định lượng 100% tại cảng.",
        "persona": "Anh Minh (Chủ doanh nghiệp chuỗi nhà hàng fast-food) và Chị Hương (Trưởng phòng Xuất nhập khẩu công ty thực phẩm đông lạnh).",
        "checklist": "1. Đã có Formulation Sheet có dấu giáp lai xưởng sản xuất; 2. Tỷ lệ thịt tính toán dưới 20.0%; 3. Đã có Giấy chứng nhận phân tích C/A độc lập; 4. Bản dịch công chứng nhãn gốc; 5. Bảng kê giải trình mã HS gửi kèm tờ khai VNACCS.",
        "url": "https://en.wikipedia.org/wiki/Food_composition_data",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 31/2022/TT-BTC, Chú giải phân loại HS WCO 2022) & daily_notes.md (ghi nhận ngày 14/10/2026 từ case study Công văn 22883/CHQ-NVTHQ).",
        "keywords": ["composition", "food"]
    },
    {
        "stt": 147,
        "week": "W42 — 14/10/2026",
        "title": "Quy Trình 6 Bước Xử Lý Khẩn Cấp Khi Bị Cơ Quan Thuế Khóa Mã Số Thuế Vì 'Không Hoạt Động Tại Địa Chỉ': Mở Lại Hóa Đơn Điện Tử Trong 48 Giờ",
        "workflow": "Bước 1: Tra cứu trạng thái người nộp thuế trên Cổng thông tin Tổng cục Thuế (tracuunnt.gdt.gov.vn) để xác định chính xác ngày ban hành Thông báo Mẫu 02/TB-ĐKT và lý do ghi nhận vắng mặt; Bước 2: Thu thập bằng chứng hiện diện pháp lý thực tế: Chụp ảnh biển hiệu công ty gắn định vị GPS ngày/giờ, văn phòng làm việc có máy tính bàn ghế, sổ chấm công, hợp đồng thuê trụ sở và biên lai đóng tiền điện nước/tiền thuê; Bước 3: Xin Giấy xác nhận địa điểm kinh doanh từ Ban Quản lý tòa nhà hoặc Tổ trưởng tổ dân phố / UBND phường nơi đặt trụ sở; Bước 4: Soạn thảo Công văn đề nghị khôi phục mã số thuế (Mẫu số 25/ĐK-TCT theo Thông tư 105/2020/TT-BTC) giải trình rõ nguyên nhân vắng mặt khách quan; Bước 5: Trực tiếp nộp hồ sơ tại Bộ phận Một cửa Chi cục Thuế quản lý và liên hệ Đội Kiểm tra Thuế phụ trách địa bàn để hẹn lịch phúc tra tại chỗ; Bước 6: Ký Biên bản ghi nhận hoạt động thực tế với cán bộ thuế để hệ thống mở lại trạng thái MST 'Đang hoạt động' và kích hoạt lại chữ ký số.",
        "formula": "Thời hạn giải quyết theo quy định: Tối đa 03 ngày làm việc kể từ ngày cơ quan thuế nhận đủ hồ sơ đề nghị khôi phục MST hợp lệ. Mức phạt chậm nộp hồ sơ thay đổi thông tin: Phạt từ 1.000.000đ - 3.000.000đ nếu quá hạn dưới 30 ngày; Phạt 3.000.000đ - 6.000.000đ nếu quá hạn 31-90 ngày theo Điều 11 Nghị định 125/2020/NĐ-CP.",
        "top3_mistakes": "1. Đóng cửa văn phòng làm việc online hoàn toàn mà không nhờ bảo vệ/lễ tân tòa nhà nhận thư từ của cơ quan nhà nước; 2. Không kiểm tra hòm thư điện tử đăng ký với thuế để nhận thông báo cảnh báo trước kiểm tra; 3. Vội vàng đi thành lập công ty mới thay vì xử lý tận gốc mã số thuế cũ (bị cấm mở công ty mới do người đại diện bị khóa trên hệ thống quốc gia).",
        "mnemonic_tip": "Quy tắc '3 Có cứu MST': Có Biển hiệu - Có Hợp đồng thuê - Có Con người hiện diện. Thiếu 1 trong 3 là Thuế gõ cửa khóa ngay.",
        "comparison": "Để mặc bị đóng MST: Mất hết hóa đơn, bị cưỡng chế tài khoản, người đại diện bị cấm xuất cảnh và đưa vào danh sách đen liên ngành VS Xử lý khẩn trong 48h: Tốn 1-2 ngày hoàn thiện hồ sơ chứng minh thực địa, nộp phạt hành chính nhẹ và bảo toàn nguyên vẹn lịch sử tín dụng doanh nghiệp.",
        "persona": "Anh Dũng (Founder công ty khởi nghiệp công nghệ) và Chị Lan (Kế toán trưởng doanh nghiệp thương mại).",
        "checklist": "1. Mẫu đơn 25/ĐK-TCT có chữ ký giám đốc; 2. Hợp đồng thuê văn phòng công chứng/hóa đơn VAT; 3. Ảnh biển hiệu + phòng làm việc có định vị tọa độ; 4. Giấy xác nhận của Ban quản lý tòa nhà; 5. Biên bản làm việc tại chỗ với Đội thuế địa bàn.",
        "url": "https://en.wikipedia.org/wiki/Tax_compliance",
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 105/2020/TT-BTC, Nghị định 125/2020/NĐ-CP, Luật Quản lý thuế 2019) & daily_notes.md (ghi nhận ngày 12/10/2026 từ case study Trần Yến Nhật).",
        "keywords": ["tax", "compliance"]
    },
    {
        "stt": 148,
        "week": "W42 — 14/10/2026",
        "title": "Hướng Dẫn Thủ Tục Xin Giấy Phép Quá Cảnh Hàng Hóa Chất & Khoáng Sản Qua Lãnh Thổ Việt Nam (Mã Loại Hình G61): Chuẩn Chỉnh Từ Cảng Biển Đến Cửa Khẩu Xuất",
        "workflow": "Bước 1: Tra cứu Danh mục kiểm soát chuyên ngành: Tra cứu mã HS và mã CAS trên Bảng MSDS để xác định mặt hàng có thuộc diện phải xin phép Bộ Công Thương theo Nghị định 69/2018/NĐ-CP (như than chì Graphite HS 3801.90.00, hóa chất công nghiệp); Bước 2: Nộp hồ sơ xin cấp Giấy phép quá cảnh trực tuyến trên Cổng Dịch vụ công Quốc gia (dichvucong.gov.vn) hoặc Cổng thông tin Một cửa Bộ Công Thương: Hồ sơ gồm Đơn đề nghị, Bản sao Hợp đồng vận tải, Bản sao Hợp đồng mua bán giữa hai nước đối tác; Bước 3: Mở tờ khai hải quan vận chuyển độc lập (mã loại hình OLA/G61) trên phần mềm VNACCS: Đính kèm số Giấy phép quá cảnh điện tử vào ô giấy phép chuyên ngành; Bước 4: Kiểm tra điều kiện phương tiện vận tải và niêm phong: Đăng ký container đạt chuẩn, gắn thiết bị định vị vệ tinh GPS kết nối với hệ thống giám sát của Tổng cục Hải quan tại cảng nhập (Hải Phòng/Cái Mép); Bước 5: Vận chuyển theo đúng tuyến đường và thời hạn quy định, giám sát bàn giao niêm phong chì tại Chi cục Hải quan cửa khẩu xuất (Cầu Treo/Cha Lo/Lao Bảo).",
        "formula": "Thời hạn quá cảnh tối đa: Không quá 30 ngày kể từ ngày hoàn thành thủ tục hải quan tại cửa khẩu nhập, trừ trường hợp được gia hạn theo quy định tại Điều 38 Luật Quản lý ngoại thương. Công thức bảo lãnh thuế quá cảnh = Trị giá tính thuế CIF x (Thuế suất NK thông thường + Thuế suất GTGT) (nếu thuộc diện yêu cầu bảo lãnh ngân hàng).",
        "top3_mistakes": "1. Cho rằng hàng chỉ đi mượn đường thì Hải quan không kiểm tra giấy phép chuyên ngành của Bộ Công Thương; 2. Phương tiện vận tải tự ý rẽ khỏi tuyến đường quốc lộ đã đăng ký trong hồ sơ quá cảnh dẫn đến hệ thống GPS cảnh báo vi phạm; 3. Chậm trễ làm thủ tục gia hạn thời hạn quá cảnh khi xe container gặp sự cố tai nạn hoặc sạt lở đường đèo.",
        "mnemonic_tip": "Quy tắc 'Phép trước - Khai sau - Giữ nguyên chì': Phải có Giấy phép Bộ duyệt trước khi tàu cập cảng, khai tờ khai G61 chuẩn và tuyệt đối không làm đứt chì seal GPS suốt dọc đường.",
        "comparison": "Quá cảnh hàng bách hóa thông thường (G61): Không cần giấy phép, chỉ cần khai tờ khai OLA và kẹp chì điện tử hải quan VS Quá cảnh hàng hóa chất/khoáng sản đặc thù: Bắt buộc xin Giấy phép Bộ Công Thương, kiểm tra giấy phép môi trường và tuyến đường vận tải hạn chế nghiêm ngặt.",
        "persona": "Anh Tuấn (Giám đốc công ty vận tải xuyên biên giới Việt - Lào) và Chị Thảo (Chuyên viên chứng từ Forwarder).",
        "checklist": "1. Giấy phép quá cảnh của Bộ Công Thương phê duyệt; 2. Bảng chỉ dẫn an toàn hóa chất MSDS có mã CAS; 3. Hợp đồng vận chuyển quá cảnh đa phương thức; 4. Thiết bị định vị GPS hành trình hoạt động tốt; 5. Tờ khai vận chuyển kết hợp OLA đã phân luồng xanh.",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 69/2018/NĐ-CP, Thông tư 39/2018/TT-BTC, Luật Quản lý ngoại thương 2017) & daily_notes.md (ghi nhận ngày 13/10/2026 từ case study Thanh Tùng).",
        "keywords": ["bill", "lading"]
    },
    {
        "stt": 149,
        "week": "W42 — 14/10/2026",
        "title": "Bí Quyết Nâng Cấp Từ 'Kiểm Tra Thông Thường' Lên Diện 'Kiểm Tra Giảm' ATTP Nhập Khẩu Theo Nghị Định 15/2018/NĐ-CP: Thông Quan Hàng Thực Phẩm Trong 24 Giờ",
        "workflow": "Bước 1: Rà soát điều kiện đạt chuẩn diện Kiểm tra giảm (Điều 16 Nghị định 15/2018/NĐ-CP): Mặt hàng đã qua 03 lần kiểm tra thông thường liên tiếp đạt yêu cầu HOẶC được sản xuất tại cơ sở có chứng nhận GMP, HACCP, ISO 22000, IFS, BRC còn hiệu lực; Bước 2: Chuẩn bị bộ hồ sơ đề nghị áp dụng phương thức kiểm tra giảm: Bản sao chứng thực chứng nhận hệ thống quản lý an toàn thực phẩm hoặc 03 Thông báo kết quả xác nhận thực phẩm đạt yêu cầu nhập khẩu liên tiếp; Bước 3: Đăng ký kiểm tra nhà nước về ATTP trên Cổng thông tin Một cửa Quốc gia (NSW), tích chọn phương thức 'Kiểm tra giảm'; Bước 4: Hệ thống cơ quan kiểm tra chuyên ngành tự động đối soát hồ sơ điện tử trong vòng 02 giờ làm việc (không lấy mẫu thử nghiệm thực tế); Bước 5: Cơ quan kiểm tra truyền Thông báo tiếp nhận kiểm tra giảm qua NSW để Hải quan cửa khẩu làm căn cứ thông quan hàng hóa ngay tại bốc dỡ.",
        "formula": "Xác suất lấy mẫu ngẫu nhiên diện kiểm tra giảm: Tối đa 5% tổng số lô hàng nhập khẩu trong vòng 01 năm do cơ quan quản lý chỉ định ngẫu nhiên. Thời gian thông quan giảm từ 7-10 ngày xuống còn dưới 24 giờ; Tiết kiệm 100% chi phí kiểm nghiệm phòng lab (từ 2 - 5 triệu đồng/mẫu).",
        "top3_mistakes": "1. Chứng nhận HACCP/ISO của nhà máy sản xuất nước ngoài hết hạn hiệu lực mà không kiểm tra gia hạn trước khi xuất xưởng; 2. Thay đổi tên gọi thương mại của sản phẩm hoặc thay đổi nhà máy đóng gói khiến chuỗi 3 lần kiểm tra đạt yêu cầu bị tính lại từ đầu; 3. Khai sai thông tin nhà sản xuất trên hồ sơ NSW so với Giấy chứng nhận lưu hành tự do (CFS).",
        "mnemonic_tip": "Công thức '3 Đạt hoặc 1 Chứng chỉ': Đạt 3 lần liên tiếp HOẶC có chứng chỉ HACCP/ISO là tự động được bay vào luồng giảm.",
        "comparison": "Kiểm tra thông thường: Kiểm tra 100% hồ sơ, thời gian chờ từ 3-5 ngày làm việc VS Kiểm tra giảm: Chỉ kiểm tra hồ sơ xác suất tối đa 5%, thông quan tự động qua NSW trong vòng 2 giờ làm việc.",
        "persona": "Chị Trang (Trưởng phòng Thu mua chuỗi siêu thị bán lẻ) và Anh Long (Giám đốc công ty phân phối đồ uống nhập khẩu).",
        "checklist": "1. Chứng nhận HACCP/ISO 22000 công chứng hợp pháp hóa lãnh sự; 2. Ba bản Thông báo kết quả kiểm tra đạt liên tiếp; 3. Hồ sơ tự công bố sản phẩm hợp lệ; 4. Nhãn phụ tiếng Việt đầy đủ cảnh báo dị ứng; 5. Tài khoản NSW đã đồng bộ chữ ký số.",
        "url": "https://en.wikipedia.org/wiki/Hazard_analysis_and_critical_control_points",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 15/2018/NĐ-CP, Luật An toàn thực phẩm 2010) & daily_notes.md (ghi nhận ngày 13/10/2026 từ case study Thu Tran).",
        "keywords": ["safety", "haccp"]
    },
    {
        "stt": 150,
        "week": "W42 — 14/10/2026",
        "title": "4 Bước Tự Kiểm Tra & Gỡ Bỏ Khoản Thu Nhập Khống Trên Ứng Dụng eTax Mobile: Bảo Vệ Dòng Vốn Hoàn Thuế Và Quyền Lợi Của Người Lao Động",
        "workflow": "Bước 1: Hướng dẫn người lao động tải và đăng nhập ứng dụng eTax Mobile bằng tài khoản Định danh điện tử VNeID Mức 2; Bước 2: Truy cập phân hệ 'Tra cứu thông tin quyết toán thuế' -> 'Nguồn thu nhập chi trả': Rà soát danh sách toàn bộ các doanh nghiệp, tổ chức đã nộp hồ sơ quyết toán khấu trừ thuế TNCN dưới mã số thuế cá nhân trong năm tài chính; Bước 3: Phát hiện doanh nghiệp lạ khai khống: Bấm trực tiếp vào nút 'Gửi phản ánh thông tin người nộp thuế không chính xác' ngay trên màn hình ứng dụng; chọn lý do 'Không có quan hệ lao động / Không phát sinh thu nhập thực tế'; Bước 4: Lập Biên bản giải trình nội bộ gửi Chi cục Thuế quản lý doanh nghiệp: Kèm theo bản cam kết của người lao động và đề nghị cơ quan thuế tiến hành thanh tra độc lập đơn vị chi trả khống, đồng thời loại trừ cảnh báo rủi ro liên đới để cơ quan thuế tiếp tục giải ngân hồ sơ hoàn thuế GTGT của công ty.",
        "formula": "Mức phạt hành vi khai khống bảng lương theo Điều 16, Điều 17 Nghị định 125/2020/NĐ-CP: Phạt 20% trên số thuế khai thiếu HOẶC phạt từ 1 đến 3 lần số thuế trốn đối với hành vi trốn thuế; Bị truy thu 100% thuế TNDN và phạt nộp chậm 0.03%/ngày trên số chi phí bị bóc tách.",
        "top3_mistakes": "1. Doanh nghiệp thờ ơ, nghĩ rằng việc người lao động bị lộ CCCD ở đâu là việc riêng cá nhân của họ; 2. Người lao động thấy bị khai khống nhưng bỏ qua không bấm nút phản ánh trên eTax Mobile khiến hệ thống mặc định thừa nhận thu nhập 2 nơi; 3. Không đối chiếu dữ liệu giữa sổ lương kế toán và tài khoản đóng BHXH trên ứng dụng VssID.",
        "mnemonic_tip": "Khẩu quyết 'Quét eTax - Bấm Phản ánh - Tách Rủi ro': Nhận diện doanh nghiệp ma trên app, bấm phản ánh online để cơ quan thuế chuyển hồ sơ sang công an xử lý.",
        "comparison": "Chủ động kiểm tra eTax Mobile định kỳ: Phát hiện sớm, phản ánh online chỉ mất 3 phút, bảo vệ 100% chi phí lương hợp lý của công ty VS Bị động chờ Thuế thanh tra: Bị đình chỉ hoàn thuế hàng tỷ đồng, bị truy thu thuế TNCN oan và mất hàng tháng giải trình sao kê ngân hàng.",
        "persona": "Chị Thảo (Kế toán trưởng phụ trách hoàn thuế) và Anh Hoàng (Chuyên viên nhân sự tiền lương).",
        "checklist": "1. 100% nhân sự cài đặt và đối soát eTax Mobile; 2. Danh sách phản ánh nguồn thu nhập sai lệch; 3. Hợp đồng lao động và bảng chấm công lưu trữ số hóa; 4. Sao kê tài khoản ngân hàng chi trả lương; 5. Công văn đề nghị thuế loại trừ cảnh báo rủi ro chéo.",
        "url": "https://en.wikipedia.org/wiki/Payroll_tax",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Luật Quản lý thuế 2019, Thông tư 111/2013/TT-BTC, Nghị định 125/2020/NĐ-CP) & daily_notes.md (ghi nhận ngày 13/10/2026 từ case study Trung Anh).",
        "keywords": ["tax", "payroll"]
    }
]
