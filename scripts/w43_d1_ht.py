# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W43 DAY 1 — 19/10/2026)
Số lượng: 5 ý tưởng (STT 161 - 165)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W43_D1 = [
    {
        "stt": 161,
        "week": "W43 — 19/10/2026",
        "title": "Cẩm Nang 4 Bước Khai Báo Phương Tiện Vận Tải Xuất Nhập Cảnh & Quá Cảnh Qua Cổng Một Cửa Quốc Gia (NSW) Chuẩn Thông Tư 127/2026/TT-BTC & Kỹ Thuật Dự Phòng Khi Hệ Thống Lỗi",
        "workflow": "Bước 1: Chuẩn bị bộ chỉ tiêu thông tin điện tử theo Mẫu quy định tại Thông tư 127/2026 gồm: Giấy phép vận tải quốc tế, danh sách hành khách/tổ lái, bản kê hàng hóa vận chuyển, thông tin số khung số máy hoặc số đăng kiểm phương tiện; Bước 2: Đăng nhập Cổng Một cửa Quốc gia (vnsw.gov.vn) bằng Chữ ký số doanh nghiệp, chọn phân hệ 'Phương tiện vận tải xuất nhập cảnh'; Bước 3: Nhập dữ liệu và kiểm tra mã xác thực định danh phương tiện; Bước 4: Nhận phản hồi cấp mã số tiếp nhận điện tử và xuất trình mã QR code cho lực lượng kiểm soát biên phòng và hải quan tại barie cửa khẩu; Nếu hệ thống lỗi, lập tức kích hoạt Phương án dự phòng bằng văn bản giấy kèm Công văn 23007/CHQ-GSQL.",
        "formula": "Thời hạn gửi hồ sơ khai báo phương tiện: Đối với đường biển tối thiểu 12 giờ trước khi tàu cập cảng; Hàng không tối thiểu 02 giờ trước khi cất cánh/hạ cánh; Đường bộ tối thiểu 01 giờ trước khi phương tiện đến trạm kiểm soát cửa khẩu biên giới.",
        "top3_mistakes": "1. Để đến sát giờ phương tiện đến cửa khẩu mới đăng nhập hệ thống NSW dẫn đến quá tải đường truyền; 2. Khai báo sai số khung, số máy của rơ-moóc hoặc container kéo theo khiến hệ thống từ chối cấp mã định danh; 3. Không lưu lại mã số tiếp nhận hồ sơ điện tử hoặc không in bản sao giấy dự phòng khi qua khu vực biên giới sóng viễn thông yếu.",
        "mnemonic_tip": "Quy tắc 'Chìa khóa mở cửa trước 1 giờ': Xe muốn lăn bánh qua biên giới thì giấy tờ điện tử phải 'bay' đến bàn hải quan trước ít nhất 60 phút!",
        "comparison": "Khai báo điện tử qua NSW vs Khai báo giấy truyền thống: Khai qua NSW thông quan tự động chỉ mất 3-5 phút, không cần tiếp xúc trực tiếp; khai báo giấy mất hàng giờ xếp hàng, dễ phát sinh tiêu cực và ách tắc tại cửa khẩu.",
        "persona": "Anh Dũng (Giám đốc công ty vận tải container xuyên biên giới / Logistics Forwarder)",
        "checklist": "☐ Chữ ký số USB Token còn hạn trên NSW\n☐ Dữ liệu đăng kiểm xe và rơ-moóc khớp 100%\n☐ Bản kê hàng hóa quá cảnh/hàng trên xe\n☐ Gửi hồ sơ trước giờ xe đến cửa khẩu 60 phút\n☐ In sẵn bản QR code và lưu mã tiếp nhận hồ sơ",
        "url": "https://en.wikipedia.org/wiki/Single-window_system",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Thông tư 127/2026/TT-BTC, Công văn hỏa tốc 23007/CHQ-GSQL & kinh nghiệm thực chiến từ daily_notes.md (17/10/2026).",
        "keywords": ["single", "window"]
    },
    {
        "stt": 162,
        "week": "W43 — 19/10/2026",
        "title": "Tuyệt Chiêu 4 Bước Tra Cứu & Khai Báo Mã Phụ 'KO' Trên Hệ Thống VNACCS: Cách Hóa Giải Triệt Để Nguy Cơ Bị 'Giam' Hàng Nghi Ngờ Thiết Bị Quân Sự & An Ninh Mạng",
        "workflow": "Bước 1: Tra cứu mã HS 8 số của hàng hóa đối chiếu với 3 danh mục kiểm soát: Danh mục Mật mã dân sự (NĐ 58/2016), An ninh mạng (NĐ 53/2018) và Hàng hóa lưỡng dụng (CV 19916 & 22233); Bước 2: Bóc tách tài liệu kỹ thuật (Catalog, Datasheet) của nhà sản xuất, chứng minh hàng hóa không có chức năng mã hóa dữ liệu quân sự hoặc tính năng an ninh thông tin bí mật; Bước 3: Tại ô 'Mã phân loại kiểm tra' hoặc 'Mã quản lý chuyên ngành' trên phần mềm VNACCS, nhập chính xác ký tự 'KO' (Không thuộc danh mục quản lý chuyên ngành); Bước 4: Đính kèm Công văn cam kết mục đích sử dụng dân sự thuần túy vào hồ sơ hải quan điện tử (HYS).",
        "formula": "Công thức thẩm định: Hàng hóa hợp lệ mã KO = [Mã HS nằm trong biểu thuế kiểm soát] + [Đặc tính kỹ thuật thực tế KHÔNG trùng mô tả quản lý chuyên ngành] + [Cam kết mục đích dân sự 100% của Giám đốc doanh nghiệp].",
        "top3_mistakes": "1. Thấy mã HS trùng danh mục là tự ý bỏ trống ô khai báo vì nghĩ 'mình không làm hàng cấm'; 2. Nhầm lẫn giữa mã 'KO' (không thuộc diện quản lý) với việc xin giấy phép miễn trừ; 3. Không chuẩn bị sẵn bản dịch kỹ thuật tiếng Việt của bảng thông số vi mạch để giải trình khi hải quan yêu cầu kiểm tra hồ sơ.",
        "mnemonic_tip": "Bí kíp 'Hai chữ KO cứu cả lô': Khi mã HS giống như súng nhưng ruột chỉ là bật lửa, hãy gõ ngay mã 'KO' để hải quan biết hàng vô hại!",
        "comparison": "Khai mã KO minh bạch vs Không khai để hệ thống tự quét: Khai mã KO giúp hệ thống VNACCS phân luồng chính xác, thông quan mượt mà; không khai sẽ bị trí tuệ nhân tạo hải quan 'bắt lỗi' tự động chuyển luồng đỏ kiểm tra STQ.",
        "persona": "Chị Hương (Kế toán trưởng / Trưởng phòng Xuất nhập khẩu doanh nghiệp linh kiện điện tử)",
        "checklist": "☐ Tra cứu đối chiếu 3 danh mục an ninh/lưỡng dụng\n☐ Tải Datasheet gốc từ website nhà sản xuất\n☐ Soạn công văn cam kết mục đích sử dụng dân sự\n☐ Điền mã 'KO' vào chỉ tiêu quản lý chuyên ngành VNACCS\n☐ Nộp kèm hồ sơ điện tử HYS",
        "url": "https://en.wikipedia.org/wiki/Dual-use_technology",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Công văn 22233/CHQ-GSQL, Công văn 19916/CHQ-GSQL, Nghị định 58/2016/NĐ-CP & kinh nghiệm thực chiến từ daily_notes.md (15/10/2026).",
        "keywords": ["dual-use", "technology"]
    },
    {
        "stt": 163,
        "week": "W43 — 19/10/2026",
        "title": "Hướng Dẫn 5 Bước Vận Hành Hàng Quá Cảnh & Trung Chuyển Cảng Biển Chuẩn Thông Tư 121/2025/TT-BTC: Kỹ Thuật Bảo Vệ Niêm Phong GPS Seal Chống Bị Phạt Nặng",
        "workflow": "Bước 1: Khai báo Bản kê hàng hóa quá cảnh điện tử (Mẫu số 09 cho đường bộ/đường thủy hoặc Mẫu 08 cho hàng không) trên Cổng Một cửa Quốc gia trước khi bốc dỡ hàng khỏi cảng; Bước 2: Nhận bàn giao thiết bị niêm phong điện tử định vị GPS (E-Seal) từ công chức hải quan giám sát cảng, kiểm tra đèn báo tín hiệu pin và sóng kết nối vệ tinh; Bước 3: Kẹp seal đúng vị trí tay khóa cửa container bên phải và chụp ảnh góc rộng hiển thị rõ số container và số seri seal; Bước 4: Quản trị hành trình: Xe/tàu chạy đúng tuyến đường và thời gian đăng ký; theo dõi ứng dụng hành trình hải quan để phát hiện sớm sự cố mất sóng viễn thông; Bước 5: Bàn giao container và chì seal nguyên vẹn tại Chi cục Hải quan cửa khẩu xuất trong vòng thời hạn quy định để kết thúc hành trình quá cảnh.",
        "formula": "Thời hạn vận chuyển hàng quá cảnh tối đa: Đường bộ không quá 30 ngày; Đường sắt không quá 30 ngày; Đường thủy nội địa không quá 15 ngày; Quá cảnh chuyển tải lưu kho không quá 30 ngày (trừ trường hợp được Cục trưởng Hải quan gia hạn).",
        "top3_mistakes": "1. Tự ý thay đổi tuyến đường vận chuyển hoặc dừng đỗ qua đêm tại các điểm không được phép khi chưa báo hải quan; 2. Làm mất, đứt hoặc va đập làm hỏng vỏ seal định vị GPS mà không lập biên bản sự cố khẩn cấp; 3. Không theo dõi dung lượng pin của E-Seal đối với các chuyến vận chuyển đường thủy nội địa dài ngày.",
        "mnemonic_tip": "Thần chú 'Đúng đường - Đúng giờ - Nguyên seal': Hàng quá cảnh chỉ là 'vị khách qua đường', đi đúng đường vạch sẵn, đúng giờ hẹn và giữ nguyên chiếc áo khóa thì không ai làm khó!",
        "comparison": "Vận chuyển quá cảnh kẹp seal cơ học truyền thống vs Kẹp E-Seal định vị GPS: Seal cơ học dễ bị gian lận cắt tráo ở góc khuất; E-Seal phát tín hiệu cảnh báo lập tức về trung tâm chỉ huy nếu bị cạy phá hoặc đi chệch tọa độ tuyến đường.",
        "persona": "Anh Minh (Chủ doanh nghiệp vận tải logistics / Đại lý hải quan đường thủy & đường bộ)",
        "checklist": "☐ Khai bản kê Mẫu 09/NSW trước khi xếp hàng\n☐ Kiểm tra đèn tín hiệu pin và mã QR của E-Seal\n☐ Chụp ảnh biên bản bàn giao seal tại cảng đi\n☐ Theo dõi hải trình GPS trên phần mềm di động\n☐ Bàn giao đúng hẹn tại trạm hải quan cửa khẩu đích",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Thông tư 121/2025/TT-BTC, Thông tư 38/2015/TT-BTC & bài học thực tế từ daily_notes.md (15/10/2026).",
        "keywords": ["security", "seal"]
    },
    {
        "stt": 164,
        "week": "W43 — 19/10/2026",
        "title": "Công Thức & 4 Bước Tính Toán Thuế Suất Thực Tế (ETR) Chuẩn Bị Cho Thuế Tối Thiểu Toàn Cầu (Pillar 2): Cẩm Nang Cho Doanh Nghiệp Phụ Trợ & Vệ Tinh Chuỗi Cung Ứng FDI",
        "workflow": "Bước 1: Thu thập Báo cáo tài chính đã kiểm toán của năm tài chính, bóc tách Chỉ tiêu Tổng chi phí thuế TNDN hiện hành và thuế TNDN hoãn lại hợp lệ; Bước 2: Xác định Lợi nhuận ròng tài chính (GloBE Income) sau khi loại trừ các khoản thu nhập được miễn giảm và điều chỉnh chênh lệch vĩnh viễn/tạm thời; Bước 3: Tính toán Thuế suất Thực tế (ETR): ETR = [Tổng số thuế TNDN hợp lệ đã ghi nhận] / [Lợi nhuận kế toán thuần GloBE]; Bước 4: So sánh ETR với mức sàn 15%: Nếu ETR < 15%, tính toán số thuế bổ sung nội địa tối thiểu đạt chuẩn (QDMTT) cần nộp tại Việt Nam; Bước 5: Lập Báo cáo giải trình nghĩa vụ thuế cung cấp cho khách hàng FDI để phục vụ hồ sơ báo cáo toàn cầu GIR.",
        "formula": "Thuế suất thực tế (Effective Tax Rate - ETR) = (Covered Taxes / Net GloBE Income) x 100%; Thuế bổ sung (Top-up Tax %) = max(0, 15% - ETR); Số thuế TNDN bổ sung phải nộp = Top-up Tax % x (Net GloBE Income - Giảm trừ theo chất chất tài sản & lao động SBIE).",
        "top3_mistakes": "1. Lấy luôn thuế suất danh nghĩa trên Giấy chứng nhận đầu tư (10% hoặc 20%) làm ETR mà không tính theo lợi nhuận kế toán thực tế; 2. Bỏ qua các khoản giảm trừ dựa trên tài sản hữu hình và tiền lương nhân công (SBIE) dẫn đến nộp thừa thuế bổ sung; 3. Không chuẩn bị trước Hồ sơ giao dịch liên kết (Transfer Pricing) để chứng minh tính thị trường của giá bán linh kiện cho đối tác FDI.",
        "mnemonic_tip": "Quy tắc 'Đo vạch nước 15%': Mọi con thuyền lợi nhuận đều phải đóng đủ vạch nước 15% thuế; nếu ở Việt Nam mới đóng 8% thì phải bù thêm 7% trước khi con thuyền cập bến quốc tế!",
        "comparison": "Ưu đãi thuế truyền thống vs Quy tắc Thuế tối thiểu toàn cầu: Ưu đãi truyền thống giúp doanh nghiệp giảm thuế xuống 0-5%; Quy tắc Pillar 2 vô hiệu hóa ưu đãi này nếu không nộp bổ sung tại nước sở tại thì tập đoàn mẹ sẽ bị thu thuế tại nước ngoài.",
        "persona": "Chị Hương (Kế toán trưởng doanh nghiệp sản xuất phụ trợ cơ khí, điện tử, bao bì xuất khẩu)",
        "checklist": "☐ Báo cáo tài chính đã kiểm toán độc lập\n☐ Bảng bóc tách chi phí thuế hoãn lại và hiện hành\n☐ Bảng tính ETR theo công thức OECD GloBE\n☐ Tính toán giảm trừ tài sản hữu hình và quỹ lương (SBIE)\n☐ Hồ sơ giá chuyển nhượng (Local File) sẵn sàng",
        "url": "https://en.wikipedia.org/wiki/Global_minimum_corporate_tax_rate",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Nghị quyết 107/2023/QH15, Thông tư hướng dẫn thuế TNDN & Khung kỹ thuật OECD Pillar 2.",
        "keywords": ["minimum", "tax"]
    },
    {
        "stt": 165,
        "week": "W43 — 19/10/2026",
        "title": "Cẩm Nang 5 Bước Tự Kiểm Kê Khí Nhà Kính & Lập Báo Cáo Phát Thải Carbon Thực Tế Đạt Chuẩn CBAM Châu Âu Cho Doanh Nghiệp Xuất Khẩu Nhôm Thép",
        "workflow": "Bước 1: Xác định ranh giới cơ sở sản xuất (Installation boundaries) và các công đoạn phát sinh phát thải trực tiếp (Scope 1: đốt than, khí gas, dầu DO) và gián tiếp (Scope 2: điện lưới tiêu thụ); Bước 2: Thu thập số liệu hoạt động (Activity data): Hóa đơn điện lực, phiếu xuất kho nhiên liệu, sản lượng nguyên liệu đầu vào (quặng sắt, phế liệu, vôi); Bước 3: Áp dụng hệ số phát thải (Emission Factors) đạt chuẩn quốc tế IPCC hoặc hệ số lưới điện quốc gia do Bộ Tài nguyên và Môi trường công bố; Bước 4: Tính toán phát thải riêng tích hợp (Specific Embedded Emissions - SEE) trên mỗi tấn sản phẩm thép/nhôm xuất khẩu; Bước 5: Điền dữ liệu vào Bảng khai CBAM Communication Template của EU và thuê đơn vị kiểm định độc lập cấp chứng nhận xác thực (Verification Report).",
        "formula": "Phát thải tích hợp riêng (SEE) = (Tổng phát thải trực tiếp Scope 1 + Tổng phát thải gián tiếp Scope 2) / Tổng sản lượng hàng hóa sản xuất (tCO2e/tấn sản phẩm).",
        "top3_mistakes": "1. Sử dụng hệ số phát thải mặc định ước tính cũ của EU đã hết hạn áp dụng; 2. Không thu thập chứng chỉ năng lượng tái tạo (I-REC) hợp lệ khi tuyên bố sử dụng điện mặt trời để giảm phát thải Scope 2; 3. Bỏ quên lượng phát thải carbon tích hợp trong nguyên liệu tiền chất mua ngoài (như phôi thép hoặc khí argon kỹ thuật).",
        "mnemonic_tip": "Công thức 'Cân đĩa carbon': Mỗi tấn thép xuất xưởng phải đặt lên bàn cân cùng lượng khói thải ra; khói càng nhẹ, giấy thông quan châu Âu càng bay nhanh!",
        "comparison": "Báo cáo phát thải tự lập nội bộ vs Báo cáo kiểm định độc lập được EU công nhận: Báo cáo nội bộ chỉ có giá trị tham khảo, bị hải quan EU từ chối; Báo cáo được kiểm định độc lập theo ISO 14065 là 'tấm hộ chiếu xanh' giúp hàng hóa thông quan thẳng vào 27 nước EU.",
        "persona": "Anh Minh (Chủ doanh nghiệp xuất khẩu sắt thép, nhôm, cơ khí chế tạo)",
        "checklist": "☐ Sơ đồ công nghệ ranh giới cơ sở sản xuất\n☐ Hóa đơn tiêu thụ điện, than, dầu 12 tháng\n☐ Chứng chỉ hợp chuẩn của phôi nguyên liệu đầu vào\n☐ Hợp đồng với đơn vị kiểm định carbon độc lập\n☐ File mẫu khai báo CBAM XML của Ủy ban Châu Âu",
        "url": "https://en.wikipedia.org/wiki/Carbon_Border_Adjustment_Mechanism",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Quy định EU CBAM 2023/956, Tiêu chuẩn ISO 14064-1 & Quyết định 01/2022/QĐ-TTg.",
        "keywords": ["carbon", "adjustment"]
    }
]
