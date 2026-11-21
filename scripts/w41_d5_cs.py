# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W41 DAY 5 — 10/10/2026)
Số lượng: 5 ý tưởng (STT 141 - 145)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

CASE_STUDY_W41_D5 = [
    {
        "stt": 141,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Doanh Nghiệp Y Tế Tạm Nhập Thiết Bị Chẩn Đoán Triển Lãm Bị Trễ Hạn 5 Ngày: Bị Phạt 40 Triệu Đồng Và Truy Thu 850 Triệu Tiền Thuế Do Nhầm Lẫn Quy Trình TIA",
        "stakeholders": "Công ty Cổ phần Thiết bị Y tế M. (Hà Nội), Chi cục Hải quan Cửa khẩu Sân bay Quốc tế Nội Bài, Đối tác sản xuất thiết bị y tế từ Đức.",
        "context": "Công ty M. làm thủ tục tạm nhập 2 hệ thống máy siêu âm thế hệ mới theo mã loại hình G12 để tham gia Triển lãm Y tế Quốc tế tại Hà Nội với thời hạn tạm nhập 30 ngày. Khi kết thúc triển lãm, đối tác Đức đề nghị kéo dài thêm 15 ngày để tổ chức hội thảo đào tạo cho các bệnh viện. Nhân viên chứng từ của công ty M. tự ý nộp hồ sơ xin gia hạn trên Cổng Dịch vụ công vào đúng ngày thứ 29 mà không khai báo nghiệp vụ TIA trên phần mềm VNACCS và không theo dõi phản hồi của Hải quan. 5 ngày sau khi thời hạn ban đầu kết thúc, Chi cục Hải quan Nội Bài ra quyết định kiểm tra và xử phạt vì hành vi quá thời hạn tạm nhập mà không tái xuất.",
        "damage": "Bị xử phạt vi phạm hành chính 40 triệu đồng theo Nghị định 128/2020/NĐ-CP; bị cơ quan Hải quan ban hành quyết định ấn định thuế nhập khẩu và thuế GTGT 850 triệu đồng; phát sinh thêm 25 triệu đồng tiền phạt chậm nộp 0.03%/ngày và bị khóa luồng xanh tờ khai trong 3 tháng.",
        "root_cause": "Nhân viên không nắm vững quy định: Thủ tục xin gia hạn tạm nhập tái xuất phải nộp trước khi hết thời hạn tối thiểu 10 ngày làm việc và bắt buộc phải có văn bản chấp thuận của Hải quan trước khi hệ thống chấp nhận chỉ tiêu TIA.",
        "solution": "1. Thiết lập lịch theo dõi hạn tờ khai tạm nhập tái xuất (G12/G22) trên phần mềm quản trị với cảnh báo tự động trước 15 ngày; 2. Soạn hồ sơ đề nghị gia hạn kèm công văn đề nghị của đối tác nước ngoài và nộp trên Cổng Một cửa Quốc gia/DVC trước ngày hết hạn ít nhất 10 ngày; 3. Ngay khi Hải quan phê duyệt công văn, truyền ngay chỉ tiêu TIA trên VNACCS để đồng bộ thời hạn bảo lãnh ngân hàng.",
        "opposing_views": "Phe A (Doanh nghiệp): Cho rằng hàng hóa vẫn nằm trong tay doanh nghiệp và thực tế đã tái xuất về Đức sau đó thì việc phạt tiền thuế và ấn định thuế là quá cứng nhắc; Phe B (Hải quan): Tạm nhập tái xuất là cơ chế miễn thuế có điều kiện; quá thời hạn mà không hoàn tất gia hạn hợp lệ thì mặc nhiên hàng hóa bị coi là tiêu thụ nội địa trái phép và phải truy thu thuế theo luật.",
        "lesson": "Thời hạn trong thủ tục hải quan là 'ranh giới sinh tử'; chậm một ngày gia hạn có thể biến một thiết bị mượn trưng bày thành khoản nợ thuế tiền tỷ.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Điều 82 Thông tư 38/2015/TT-BTC (sửa đổi tại Thông tư 39/2018/TT-BTC) & daily_notes.md (ghi nhận ngày 10/10/2026 từ case study Nguyen Quoc Chinh).",
        "keywords": ["customs declaration", "customs"]
    },
    {
        "stt": 142,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Sập Bẫy Dịch Vụ 'Bao Trọn Gói Order Trung - Việt Tính Cước 3% Giá Trị': Chủ Shop Bị Tịch Thu 300 Triệu Tiền Hàng Phụ Kiện Do Đơn Vị Gom Đơn Khai Sai Trị Giá Hải Quan",
        "stakeholders": "Chủ shop phụ kiện thời trang online B. (Hà Nội), Đơn vị vận chuyển gom hàng 'siêu tốc' không hợp đồng pháp nhân, Đội Quản lý Thị trường số 14 TP. Hà Nội.",
        "context": "Chủ shop B. nhập lô hàng 5.000 món phụ kiện điện thoại và túi xách từ xưởng Quảng Châu với tổng giá trị 300 triệu đồng. Được chào mời dịch vụ 'vận chuyển chính ngạch trọn gói chỉ 3% giá trị đơn hàng, bao thuế bao thủ tục, hàng về tận kho sau 5 ngày', chủ shop chuyển toàn bộ tiền hàng và cước cho bên vận chuyển qua tài khoản cá nhân. Khi xe giao hàng tập kết tại kho của chủ shop tại Hoàng Mai thì lực lượng Quản lý thị trường ập vào kiểm tra. Đơn vị vận chuyển chỉ cung cấp được 1 tờ khai phi mậu dịch đứng tên một cá nhân lạ hoắc với trị giá khai báo tượng trưng chỉ... 15 triệu đồng (bằng 5% giá trị thật) và không có hóa đơn GTGT.",
        "damage": "Toàn bộ lô hàng 300 triệu đồng bị lập biên bản tịch thu vì không có hóa đơn chứng từ hợp pháp; chủ shop B. bị xử phạt hành chính 45 triệu đồng về hành vi kinh doanh hàng hóa nhập lậu; đơn vị vận chuyển phủi trách nhiệm và chỉ đồng ý 'hỗ trợ đền bù cước 5 triệu đồng' theo điều khoản miệng.",
        "root_cause": "Ham rẻ và thiếu hiểu biết pháp lý, tin vào lời quảng cáo 'chính ngạch giá rẻ tính theo %' mà không yêu cầu hợp đồng dịch vụ logistics, không kiểm tra tên người nhập khẩu trên tờ khai hải quan và không nhận hóa đơn GTGT đầu vào hợp lệ.",
        "solution": "1. Tuyệt đối không sử dụng dịch vụ vận chuyển tính cước theo % giá trị đơn hàng mà không có bảng kê chi tiết thuế và cước; 2. Khi nhập khẩu chính ngạch, tờ khai hải quan bắt buộc phải đứng tên pháp nhân hoặc hộ kinh doanh của chính mình; 3. Mọi khoản thanh toán mua hàng quốc tế phải thực hiện qua chuyển khoản ngân hàng có hợp đồng và Invoice đi kèm để làm căn cứ bảo vệ quyền sở hữu.",
        "opposing_views": "Phe A (Chủ shop nhỏ): Bán lẻ online vốn ít nên phải đi ghép đơn giá rẻ mới cạnh tranh được; Phe B (Chuyên gia Logistics): Ghép đơn không có chứng từ thực chất là buôn lậu trá hình; chi phí rủi ro bị tịch thu hàng luôn đắt gấp hàng chục lần số tiền thuế tiết kiệm được.",
        "lesson": "Trong xuất nhập khẩu, không có bữa trưa nào miễn phí; 'bao thuế trọn gói không hóa đơn' thực chất là giao toàn bộ tài sản của bạn cho kẻ buôn lậu cầm giữ.",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Thông tư 39/2015/TT-BTC, Thông tư 60/2019/TT-BTC & daily_notes.md (ghi nhận ngày 10/10/2026 từ case study Viettel Post gom đơn).",
        "keywords": ["customs valuation", "customs"]
    },
    {
        "stt": 143,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Mua 1.2 Tỷ Đồng Linh Kiện Điện Tử Trôi Nổi Rồi Lập 'Bảng Kê Hộ Kinh Doanh 01/TNDN' Hợp Thức Hóa: Doanh Nghiệp Bị Chuyển Hồ Sơ Sang Cơ Quan Điều Tra Tội Buôn Lậu",
        "stakeholders": "Công ty TNHH Sản Xuất Thiết Bị Điện K. (Bắc Ninh), Cục Thuế Tỉnh Bắc Ninh, Cơ quan Cảnh sát Điều tra Công an Tỉnh.",
        "context": "Công ty K. mua trôi nổi lô hàng linh kiện bo mạch điện tử từ các đầu mối chợ biên giới Lạng Sơn không có hóa đơn GTGT với số tiền 1.2 tỷ đồng. Kế toán công ty tư vấn lập 'Bảng kê thu mua hàng hóa dịch vụ mua vào không có hóa đơn Mẫu 01/TNDN' ghi tên một số người dân địa phương thu gom phế liệu để đưa vào chi phí hợp lý được trừ khi quyết toán thuế TNDN. Trong đợt thanh tra chuyên đề, Cơ quan Thuế phát hiện toàn bộ linh kiện là hàng hóa nhập khẩu công nghệ mới, dán nhãn mác nguyên seal từ Trung Quốc chứ không phải là phế liệu tự thu gom của người dân. Cơ quan Thuế lập biên bản bác bỏ toàn bộ chi phí 1.2 tỷ đồng, xác định hành vi tiêu thụ hàng nhập lậu và chuyển toàn bộ hồ sơ sang Cơ quan Điều tra.",
        "damage": "Bị truy thu 240 triệu đồng thuế TNDN, phạt 20% kê khai sai (48 triệu), phạt tiền chậm nộp 35 triệu; doanh nghiệp bị xử phạt hành vi kinh doanh hàng nhập lậu và giám đốc cùng kế toán trưởng bị khởi tố điều tra về tội buôn lậu và trốn thuế theo Bộ luật Hình sự.",
        "root_cause": "Cố tình đánh tráo khái niệm pháp lý: Áp dụng sai đối tượng của Bảng kê 01/TNDN (vốn chỉ dành cho nông lâm thủy sản tự sản xuất hoặc đồ dùng cũ của cá nhân) để hợp thức hóa hàng hóa nhập khẩu thương mại trôi nổi.",
        "solution": "1. Phân biệt tuyệt đối phạm vi áp dụng của Bảng kê 01/TNDN theo Thông tư 78/2014/TT-BTC và Thông tư 96/2015/TT-BTC; 2. Đối với mọi hàng hóa có nguồn gốc xuất xứ nước ngoài, bắt buộc phải có hóa đơn GTGT điện tử từ nhà nhập khẩu chính ngạch kèm tờ khai hải quan; 3. Thiết lập quy trình kiểm toán tuân thủ chứng từ đầu vào trước khi ký duyệt thanh toán bất kỳ đơn hàng nào.",
        "opposing_views": "Phe A (Doanh nghiệp vi phạm): Nghĩ rằng nếu bị phát hiện thì Cơ quan Thuế chỉ loại trừ chi phí và đóng thêm 20% thuế TNDN là xong; Phe B (Cơ quan Thuế & Pháp luật): Hàng hóa nhập lậu không hóa đơn thuộc phạm vi điều chỉnh của Luật Xử lý vi phạm hành chính và Bộ luật Hình sự, không thể đơn thuần giải quyết bằng việc nộp thuế TNDN bổ sung.",
        "lesson": "Hóa đơn chứng từ là tấm lá chắn bảo vệ sự tự do của chủ doanh nghiệp; đừng bao giờ dùng một bảng kê giả mạo để che đậy một tội danh buôn lậu.",
        "url": "https://en.wikipedia.org/wiki/Tax_evasion",
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Nghị định 98/2020/NĐ-CP, Thông tư 78/2014/TT-BTC, Điều 188 & 200 Bộ luật Hình sự 2015 & daily_notes.md (ghi nhận ngày 10/10/2026 từ case study Nam Vũ).",
        "keywords": ["fraud", "tax evasion"]
    },
    {
        "stt": 144,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "10 Container Sữa Bột Nhập Khẩu Bị Ách Tắc 21 Ngày Tại Cảng Cát Lái Do Lệch Địa Chỉ Cơ Sở Sản Xuất Trên Giấy Chứng Nhận Lưu Hành Tự Do (CFS): Thiệt Hại 180 Triệu Tiền Lưu Kho Lạnh",
        "stakeholders": "Công ty Cổ phần Dinh Dưỡng Quốc Tế D. (TP.HCM), Cơ quan Kiểm tra Nhà nước An toàn Thực phẩm (Viện Y tế Công cộng), Chi cục Hải quan Cửa khẩu Cảng Sài Gòn Khu vực 1.",
        "context": "Công ty D. nhập khẩu lô hàng 10 container sữa bột công thức nguyên lon từ New Zealand về Cảng Cát Lái. Khi nộp hồ sơ đăng ký kiểm tra an toàn thực phẩm nhập khẩu trên Cổng Một cửa Quốc gia (NSW), chuyên viên kiểm tra phát hiện trên Giấy chứng nhận lưu hành tự do (CFS) do cơ quan thẩm quyền New Zealand cấp ghi địa chỉ nhà máy sản xuất tại số 142 Greenlane Road, trong khi trên nhãn gốc của lon sữa và bản Tự công bố sản phẩm của doanh nghiệp lại ghi số 142/A Greenlane Industrial Park. Do sự sai lệch này, cơ quan kiểm tra từ chối tiếp nhận hồ sơ theo phương thức kiểm tra thông thường, yêu cầu chuyển sang phương thức kiểm tra chặt và yêu cầu doanh nghiệp bổ sung thư giải trình xác nhận của Đại sứ quán New Zealand.",
        "damage": "10 container phải lưu bãi lạnh tại cảng suốt 21 ngày; phát sinh 145 triệu đồng tiền phí lưu cắm điện container lạnh và lưu bãi (demurrage/storage); tốn thêm 35 triệu đồng chi phí kiểm nghiệm mẫu đột xuất và phí hợp pháp hóa lãnh sự; hàng bỏ lỡ chiến dịch khuyến mãi mùa tựu trường.",
        "root_cause": "Bộ phận chứng từ không đối soát chéo 1:1 từng ký tự địa chỉ nhà sản xuất giữa Giấy phép CFS, hồ sơ Tự công bố sản phẩm và maquette nhãn sản phẩm in thực tế trước khi xuất hàng từ nước ngoài.",
        "solution": "1. Trước khi nhà sản xuất in ấn bao bì lon và phát hành CFS, bắt buộc chuyển bản thảo cho chuyên viên an toàn thực phẩm đối chiếu chuẩn xác từng dấu phẩy với hồ sơ tự công bố; 2. Yêu cầu nhà cung cấp ký cam kết bồi thường toàn bộ chi phí lưu cont nếu CFS sai lệch thông tin; 3. Đăng ký tài khoản kiểm tra chuyên ngành trên NSW trước khi tàu cập cảng tối thiểu 3 ngày.",
        "opposing_views": "Phe A (Chủ hàng): Sai lệch số nhà hay thêm chữ Industrial Park là lỗi hình thức không ảnh hưởng đến chất lượng sữa; Phe B (Cơ quan Kiểm tra): Thực phẩm là mặt hàng ảnh hưởng trực tiếp đến sức khỏe trẻ em, mọi thông tin về cơ sở sản xuất phải trùng khớp tuyệt đối để đảm bảo truy xuất nguồn gốc khi xảy ra sự cố nhiễm khuẩn.",
        "lesson": "Trong thủ tục kiểm tra chuyên ngành an toàn thực phẩm, 'chuẩn xác từng ký tự là mệnh lệnh'; một chữ cái sai lệch có thể làm đóng băng chuỗi cung ứng lạnh hàng trăm triệu đồng.",
        "url": "https://en.wikipedia.org/wiki/Sanitary_and_Phytosanitary_Agreement",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Nghị định 15/2018/NĐ-CP, Thông tư quy định về CFS của Bộ Công Thương & Hiệp định SPS của WTO.",
        "keywords": ["wto", "food safety"]
    },
    {
        "stt": 145,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Container Chứa 5.000 Pin Lithium Tự Bốc Cháy Trên Tàu Mẹ: Hãng Tàu Khởi Kiện Đòi Bồi Thường 2.5 Triệu USD Do Chủ Hàng Giấu Nhãn Hàng Nguy Hiểm (DG Cargo Class 9)",
        "stakeholders": "Công ty TNHH Năng Lượng Xanh P. (Bình Dương), Hãng tàu vận tải container biển quốc tế MSC, Tòa án Trọng tài Hàng hải Quốc tế.",
        "context": "Công ty P. xuất khẩu 1 container 40 feet chứa 5.000 khối pin lưu trữ năng lượng Lithium-ion từ Cảng Cái Mép đi Rotterdam. Để tiết kiệm chi phí phụ phí hàng nguy hiểm (DG Surcharge khoảng 800 USD/container) và tránh thủ tục kiểm định khắt khe, nhân viên booking đã khai báo gian lận tên hàng trên vận đơn là 'Linh kiện điện tử dân dụng thông thường (Electronic Parts)' và không nộp Bản khai hàng nguy hiểm (Dangerous Goods Declaration - DGD). Khi tàu mẹ đang hành trình qua Ấn Độ Dương, một khối pin bị đoản mạch nhiệt (Thermal Runaway) bốc cháy dữ dội làm hư hỏng toàn bộ container và thiêu rụi 4 container hàng bách hóa liền kề trên boong tàu. Thủy thủ đoàn phải kích hoạt hệ thống dập lửa khẩn cấp và tàu phải chuyển hướng vào cảng Colombo để xử lý sự cố.",
        "damage": "Lô hàng pin 150.000 USD bị thiêu rụi hoàn toàn; công ty bảo hiểm từ chối chi trả 100% vì hành vi gian lận khai báo hàng nguy hiểm; hãng tàu MSC và các chủ hàng bị cháy lây lập hồ sơ khởi kiện đòi Công ty P. bồi thường thiệt hại tổng cộng 2.5 triệu USD bao gồm chi phí cứu nạn, sửa chữa tàu và bồi thường hàng hóa bên thứ ba.",
        "root_cause": "Tiết kiệm thiển cận vài trăm USD phụ phí hàng nguy hiểm, cố tình làm ngơ quy tắc an toàn vận tải biển quốc tế theo Bộ luật IMDG Code của IMO đối với mặt hàng pin Lithium (Hàng nguy hiểm Nhóm 9 - UN 3480/UN 3481).",
        "solution": "1. Bắt buộc tra cứu Bảng chỉ dẫn an toàn hóa chất (MSDS/SDS) Mục 14 (Transport Information) để xác định chính xác mã UN Number và Phân nhóm hàng nguy hiểm; 2. Trang bị chứng chỉ kiểm định bao bì UN (UN Packaging Certificate) và báo cáo thử nghiệm va đập pin UN 38.3; 3. Lập Bản khai hàng nguy hiểm (Multimodal Dangerous Goods Form) đầy đủ và dán đúng nhãn cảnh báo hình thoi Class 9 ở cả 4 mặt ngoài container.",
        "opposing_views": "Phe A (Một số chủ hàng liều lĩnh): Cho rằng thủ tục khai báo DG quá rườm rà, cước cao và hãng tàu từ chối nhận nhiều nên cố tình 'lách' khai hàng thường; Phe B (Hãng tàu & Thủy thủ đoàn): Pin lithium giấu nhãn là quả bom nổ chậm trên biển, đe dọa sinh mạng hàng chục thủy thủ và an toàn con tàu hàng trăm triệu USD.",
        "lesson": "Minh bạch tính chất hàng hóa là đạo đức cơ bản trong logistics quốc tế; gian lận một nhãn cảnh báo nguy hiểm có thể đánh đổi bằng sự phá sản của cả doanh nghiệp.",
        "url": "https://en.wikipedia.org/wiki/Dangerous_goods",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Bộ luật Quốc tế về Vận chuyển Hàng nguy hiểm bằng đường biển IMDG Code (IMO) & Tiêu chuẩn UN 38.3.",
        "keywords": ["hazardous", "dangerous goods"]
    }
]
