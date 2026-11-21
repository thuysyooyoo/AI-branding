# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W42 DAY 3 — 16/10/2026)
Số lượng: 5 ý tưởng (STT 156 - 160)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W42_D3 = [
    {
        "stt": 156,
        "week": "W42 — 16/10/2026",
        "title": "Cẩm Nang 4 Bước Thẩm Định & Thiết Lập Nhãn Gốc Hàng Nhập Khẩu Đạt Chuẩn Nghị Định 37/2026/NĐ-CP: Tuyệt Chiêu 'Khóa' Nhà Cung Cấp Ngoại Trước Khi Xuất Xưởng",
        "workflow": "Bước 1: Soạn thảo Maquette tem nhãn chuẩn Nghị định 37/2026 gồm 3 thông tin bắt buộc bằng tiếng Anh/Việt: Tên hàng hóa, Tên & địa chỉ nhà sản xuất, Xuất xứ hàng hóa (ví dụ: Made in China); Bước 2: Đưa điều khoản 'Ràng buộc quy chuẩn tem nhãn' vào Hợp đồng ngoại thương: Shipper chịu 100% chi phí phạt và lưu kho nếu nhãn sai quy chuẩn; Bước 3: Giữ lại 30-40% tiền thanh toán trước khi xuất hàng, yêu cầu Shipper gửi video/ảnh chụp cận cảnh Nameplate kim loại bắn trên thân máy và tem dán vỏ thùng carton; Bước 4: Kiểm tra tính ăn khớp 100% giữa thông tin trên nhãn vật lý với Hóa đơn thương mại (Invoice), Bảng kê đóng gói (Packing List) và Chứng nhận xuất xứ (C/O Form E).",
        "formula": "Công thức kiểm tra nhãn chuẩn: Nhãn hợp lệ = [Tên sản phẩm] + [Xuất xứ: Made in...] + [Tên & địa chỉ nhà sản xuất/chịu trách nhiệm] + [Thông số kỹ thuật/Model khớp 100% với Invoice & Packing List].",
        "top3_mistakes": "1. Nhẹ dạ tin lời Shipper 'tao xuất sang VN nhiều rồi không sao' mà không kiểm tra tem trước khi đóng cont; 2. Nghĩ rằng có C/O Form E là được quyền nợ xuất xứ trên tem máy; 3. In nhãn dán bằng giấy mỏng dễ bong tróc trong quá trình vận chuyển đường biển dẫn đến mất nhãn khi kiểm hóa.",
        "mnemonic_tip": "Quy tắc 'Bất động 3 điểm': Như chiếc chứng minh thư dán trên trán, máy móc nhập khẩu phải có 3 điểm cố định không bao giờ được thiếu: Tên máy - Tên xưởng - Nước làm ra!",
        "comparison": "Dán nhãn chuẩn từ xưởng nước ngoài vs Về cảng mới xin đưa vào kho khắc phục: Dán chuẩn từ xưởng chi phí chỉ vài trăm nghìn tiền in ấn; để về cảng bị dừng thông quan phát sinh lưu bãi, tiền phạt vi phạm và phí cẩu kéo từ 30-60 triệu đồng!",
        "persona": "Anh Minh (Chủ doanh nghiệp sản xuất cơ khí / nhập khẩu máy móc mới)",
        "checklist": "☐ Maquette nhãn duyệt trước khi sản xuất\n☐ Điều khoản phạt nhãn trong Hợp đồng mua bán\n☐ Giữ 40% tiền cọc chờ ảnh chụp tem thực tế\n☐ Đối chiếu Nameplate khớp Invoice & Packing List\n☐ Tem in bằng chất liệu bền (kim loại/decal nhựa chống nước)",
        "url": "https://en.wikipedia.org/wiki/Country_of_origin",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Nghị định 37/2026/NĐ-CP, Nghị định 128/2020/NĐ-CP & kinh nghiệm thực chiến từ daily_notes.md (16/10/2026).",
        "keywords": ["country", "origin"]
    },
    {
        "stt": 157,
        "week": "W42 — 16/10/2026",
        "title": "Quy Trình 5 Bước Vận Hành Lô Hàng Thương Mại 3 Bên (Triangular Trade) & Kỹ Thuật Ẩn Danh Bằng Switch Bill Of Lading (Switch B/L) Không Lộ Giá Vốn",
        "workflow": "Bước 1: Ký 2 hợp đồng ngoại thương độc lập: Hợp đồng 1 giữa Bên bán A (Việt Nam) và Bên trung gian B; Hợp đồng 2 giữa Bên trung gian B và Bên mua C; Bước 2: Phát hành Vận đơn gốc thứ nhất (B/L 1) tại cảng đi: Shipper A, Consignee To Order của Ngân hàng bên B (hoặc bên B); Bước 3: Bên trung gian B nộp lại toàn bộ 3/3 bản gốc B/L 1 cho hãng tàu/forwarder tại điểm trung chuyển (Singapore/Hong Kong) và yêu cầu phát hành Vận đơn chuyển đổi (Switch B/L - B/L 2); Bước 4: Lập B/L 2 với thông tin điều chỉnh: Shipper là bên B (hoặc đại diện của B), Consignee là bên C, thay đổi mô tả hàng hóa phù hợp và kèm Hóa đơn thương mại mới (Invoice 2) có giá bán B-C; Bước 5: Gửi bộ chứng từ mới (B/L 2 + Invoice 2 + Packing List 2) cho người mua C làm thủ tục nhập khẩu và thanh toán.",
        "formula": "Biên lợi nhuận trung gian (Arbitrage Margin) = Trị giá Hóa đơn B-C - Trị giá Hóa đơn A-B - Phí phát hành Switch B/L - Chi phí giao dịch ngoại tệ.",
        "top3_mistakes": "1. Chưa thu hồi đủ 3/3 bản gốc B/L 1 mà hãng tàu đã vội phát hành B/L 2 dẫn đến lưu hành 2 bộ vận đơn song song; 2. Forwarder gửi nhầm hóa đơn giá gốc A-B cho người mua C làm lộ bí mật giá vốn; 3. Không kiểm tra sự tương thích của C/O giáp lưng (Back-to-Back C/O) khiến hàng hóa bị từ chối ưu đãi thuế quan tại cảng đến.",
        "mnemonic_tip": "Kỹ thuật Switch B/L như 'thay áo cho hàng hóa': Hàng đi một mạch thẳng đường biển, nhưng giấy tờ phải 'thay áo đổi tên' ở giữa đường để người mua không biết ai là người may áo!",
        "comparison": "Giao dịch trực tiếp 2 bên vs Thương mại 3 bên dùng Switch B/L: Giao dịch trực tiếp đơn giản nhưng lộ nguồn cung ứng; Switch B/L phức tạp về chứng từ nhưng bảo vệ tuyệt đối chuỗi khách hàng và biên lợi nhuận của nhà trung gian.",
        "persona": "Anh Dũng (Giám đốc công ty thương mại XNK trung gian / Forwarder quốc tế)",
        "checklist": "☐ Ký 2 hợp đồng ngoại thương riêng biệt\n☐ Chỉ định Forwarder có uy tín xử lý Switch B/L\n☐ Thu hồi đủ 3/3 bản gốc B/L 1 trước khi in B/L 2\n☐ Kiểm tra kỹ thông tin Shipper/Consignee trên B/L 2\n☐ Tách biệt tuyệt đối 2 bộ Invoice và Packing List",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Incoterms 2020, Hướng dẫn UCP 600 của ICC & tình huống thực chiến từ daily_notes.md (16/10/2026).",
        "keywords": ["bill", "lading"]
    },
    {
        "stt": 158,
        "week": "W42 — 16/10/2026",
        "title": "Quy Trình 4 Bước Tự Kê Khai & Khấu Trừ Thuế Nhà Thầu (FCT) Khi Mua Phần Mềm SaaS & Công Cụ AI Quốc Tế (ChatGPT, Midjourney, Phần Mềm Logistics)",
        "workflow": "Bước 1: Phân loại bản chất khoản chi phí thanh toán cho đối tác ngoại: Bản quyền phần mềm (chỉ chịu 10% TNDN, miễn GTGT) hay Dịch vụ phần mềm / Điện toán đám mây / AI (chịu 5% GTGT và 5% TNDN); Bước 2: Xác định phương pháp tính thuế: Áp dụng phương pháp Gross-up (nếu hợp đồng quy định giá Net và người mua Việt Nam chịu thuế) hoặc phương pháp khấu trừ trực tiếp tại nguồn; Bước 3: Lập Tờ khai thuế nhà thầu Mẫu 01/NTNN trên Cổng Thuế điện tử (thuedientu.gdt.gov.vn) trong thời hạn 10 ngày kể từ ngày phát sinh thanh toán; Bước 4: Nộp tiền thuế vào Kho bạc Nhà nước, lưu trữ trọn bộ hồ sơ gồm Hóa đơn điện tử PDF của nhà cung cấp, Giấy nộp tiền vào NSNN và Sao kê thẻ tín dụng doanh nghiệp để đưa 100% vào chi phí được trừ khi quyết toán thuế TNDN.",
        "formula": "Doanh thu tính thuế TNDN (Gross-up) = Số tiền thanh toán thực tế / (1 - 5% thuế TNDN); Thuế TNDN nộp thay = Doanh thu tính thuế TNDN x 5%; Thuế GTGT nộp thay = Doanh thu tính thuế GTGT x 5%.",
        "top3_mistakes": "1. Tưởng quẹt thẻ tín dụng công ty có sao kê ngân hàng là mặc nhiên được tính chi phí hợp lý mà không nộp thuế nhà thầu; 2. Đánh đồng toàn bộ phần mềm SaaS đều là 'chuyển giao bản quyền' để áp thuế suất GTGT 0%; 3. Không lưu vết hợp đồng điện tử (Terms of Service) và hóa đơn điện tử PDF làm căn cứ giải trình khi thanh tra thuế.",
        "mnemonic_tip": "Quy tắc 'Đi qua cầu phải trả phí': Tiền chảy ra nước ngoài mua dịch vụ là chiếc xe qua cầu; Thuế nhà thầu chính là vé trạm thu phí để được đóng dấu lưu hành hợp pháp vào chi phí công ty!",
        "comparison": "Tự kê khai nộp thuế nhà thầu vs Mua qua đại lý ủy quyền trong nước: Tự kê khai tiết kiệm chi phí dịch vụ nhưng kế toán phải thạo thủ tục tờ khai 01/NTNN; mua qua đại lý trong nước giá cao hơn 10-15% nhưng nhận ngay hóa đơn GTGT đầu vào hợp pháp.",
        "persona": "Chị Hương (Kế toán trưởng doanh nghiệp Logistics & Xuất nhập khẩu)",
        "checklist": "☐ Phân loại dịch vụ chịu thuế (Bản quyền vs Dịch vụ đám mây)\n☐ Tính toán số thuế nộp thay theo công thức Gross-up\n☐ Nộp tờ khai Mẫu 01/NTNN trong vòng 10 ngày\n☐ Nộp tiền thuế vào Kho bạc Nhà nước\n☐ Lưu trữ trọn bộ hóa đơn PDF, sao kê và chứng từ nộp thuế",
        "url": "https://en.wikipedia.org/wiki/Withholding_tax",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Thông tư 103/2014/TT-BTC, Thông tư 80/2021/TT-BTC & bài học kế toán thực chiến từ daily_notes.md (16/10/2026).",
        "keywords": ["withholding", "tax"]
    },
    {
        "stt": 159,
        "week": "W42 — 16/10/2026",
        "title": "Checklist 6 Điểm Vàng Tự Rà Soát & Bảo Vệ Mã Số Thuế Cá Nhân: Quy Trình Xử Lý Khi Bị Kẻ Xấu Mạo Danh CCCD Thành Lập Doanh Nghiệp 'Ma' Nợ Thuế",
        "workflow": "Bước 1: Mở ứng dụng eTax Mobile và Cổng tra cứu người nộp thuế (tracuunnt.gdt.gov.vn), đăng nhập bằng VNeID Mức 2 để tra cứu danh sách toàn bộ các mã số thuế và doanh nghiệp đang gắn với số CCCD của bạn; Bước 2: Nếu phát hiện có doanh nghiệp lạ đứng tên mình làm Đại diện pháp luật, lập tức chụp màn hình và in toàn bộ thông tin đăng ký kinh doanh; Bước 3: Soạn Đơn kiến nghị & tố giác hành vi giả mạo danh tính gửi đồng thời đến Phòng Đăng ký kinh doanh (Sở KH&ĐT) và Chi cục Thuế quản lý doanh nghiệp đó; Bước 4: Làm việc trực tiếp với Đội Quản lý nợ & cưỡng chế thuế, cung cấp chứng cứ ngoại phạm và chữ ký mẫu để chứng minh chữ ký trên hồ sơ thành lập là giả mạo; Bước 5: Nộp đơn trình báo sang Cơ quan Cảnh sát điều tra (PC03) để thụ lý hồ sơ tội phạm làm giả con dấu, tài liệu của cơ quan tổ chức; Bước 6: Theo dõi và nhận Văn bản tạm hoãn các biện pháp cưỡng chế thuế và hủy bỏ quyết định tạm hoãn xuất cảnh.",
        "formula": "Hồ sơ pháp lý giải cứu danh tính thuế = [Đơn tố giác giả mạo chữ ký] + [Chứng cứ ngoại phạm/chữ ký mẫu] + [Biên bản làm việc với Sở KH&ĐT] + [Giấy xác nhận thụ lý tin báo tội phạm của Công an PC03].",
        "top3_mistakes": "1. Thờ ơ, bỏ qua khi nhận được thông báo nợ thuế vì nghĩ 'mình không mở công ty thì không việc gì phải sợ'; 2. Chỉ gửi đơn một nơi (chỉ gửi công an mà không thông báo cho thuế dẫn đến bị ban hành trát cấm xuất cảnh); 3. Tùy tiện cho người thân, bạn bè mượn ảnh chụp CCCD hoặc quét khuôn mặt trên ứng dụng tài chính.",
        "mnemonic_tip": "Chiếc chìa khóa nhà không bao giờ cho mượn, thì số CCCD và dữ liệu khuôn mặt cũng là 'chìa khóa số' của cuộc đời bạn; đừng bao giờ để người khác mở cửa rước họa vào nhà!",
        "comparison": "Chủ động rà soát định kỳ trên eTax Mobile vs Đợi đến khi bị chặn ở cửa khẩu sân bay mới biết: Rà soát định kỳ mất 5 phút trên điện thoại; bị chặn ở sân bay mất hàng trăm triệu tiền vé máy bay, lỡ việc làm ăn và mất cả năm đi minh oan!",
        "persona": "Anh Minh (Chủ shop kinh doanh / Cá nhân có hoạt động thương mại tự do)",
        "checklist": "☐ Cài đặt eTax Mobile và đăng nhập bằng VNeID Mức 2\n☐ Tra cứu danh mục mã số thuế cá nhân và pháp nhân\n☐ Kiểm tra thông báo thuế định kỳ hàng quý\n☐ Tuyệt đối không gửi ảnh CCCD 2 mặt lên mạng xã hội\n☐ Nộp đơn tố giác ngay khi phát hiện thông tin bất thường",
        "url": "https://en.wikipedia.org/wiki/Tax_identity_theft",
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Luật Quản lý thuế 2019, Luật Doanh nghiệp 2020, Bộ luật Hình sự 2015 & case study thực chiến từ daily_notes.md (16/10/2026).",
        "keywords": ["identity", "tax"]
    },
    {
        "stt": 160,
        "week": "W42 — 16/10/2026",
        "title": "Công Thức & 5 Bước Phân Loại Mã HS Hàng Hóa Thực Phẩm Chế Biến Hỗn Hợp Dựa Trên Tiêu Chí 20% Hàm Lượng Thịt (Chương 16 vs Chương 19)",
        "workflow": "Bước 1: Thu thập Bảng phân tích thành phần công thức (Recipe Formulation / Product Specification) có xác nhận đóng dấu của nhà sản xuất nước ngoài; Bước 2: Tính toán chính xác tỷ lệ khối lượng của toàn bộ các thành phần có nguồn gốc động vật (thịt heo, thịt bò, xúc xích, mỡ, tiết, phụ phẩm) trên tổng khối lượng tịnh của sản phẩm thành phẩm; Bước 3: Áp dụng Quy tắc 20% theo Chú giải 2 Chương 16: Nếu hàm lượng thịt > 20%, định danh vào Chương 16 (nhóm 16.01 hoặc 16.02); nếu hàm lượng thịt <= 20% và bản chất là sản phẩm từ bột ngũ cốc, định danh vào nhóm 19.05; Bước 4: Đối chiếu chính sách quản lý chuyên ngành: Mặt hàng thuộc Chương 16 bắt buộc đăng ký Kiểm tra vệ sinh thú y và an toàn thực phẩm nhập khẩu tại Cục Thú y; Bước 5: Nếu có tranh chấp kỹ thuật với hải quan, lập hồ sơ đề nghị Xác định trước mã số (Advance Ruling) gửi Tổng cục Hải quan trước khi nhập hàng.",
        "formula": "Tỷ lệ hàm lượng thịt (%) = (Tổng khối lượng thành phần thịt + xúc xích + phụ phẩm động vật / Tổng khối lượng tịnh của sản phẩm thành phẩm) x 100%. (Nếu Tỷ lệ > 20% -> Chương 16; Nếu Tỷ lệ <= 20% -> Chương 19).",
        "top3_mistakes": "1. Khai báo theo tên gọi thương mại (ví dụ khai 'Bánh mì sandwich' để vào 19.05) trong khi bên trong kẹp miếng thịt bò chiếm 35% khối lượng; 2. Thiếu chứng từ Certificate of Analysis (C/A) chứng minh tỷ lệ % khối lượng thành phần từ xưởng sản xuất; 3. Bỏ qua khâu kiểm dịch động vật của Cục Thú y khi hàng bị cơ quan hải quan chuyển từ Chương 19 sang Chương 16.",
        "mnemonic_tip": "Quy tắc 'Cái cân 20%': Trên chiếc cân phân loại hải quan, chỉ cần miếng thịt vượt quá 20% sức nặng thì chiếc bánh mì lập tức biến thành tảng thịt đông lạnh chịu kiểm tra thú y gắt gao!",
        "comparison": "Khai báo vào Chương 19 (Bánh mì 19.05) vs Bị ấn định sang Chương 16 (Chế phẩm thịt 16.02): Chương 19 thủ tục thông quan nhanh, thuế suất ưu đãi; Chương 16 phải lấy mẫu phân tích thú y, lưu kho chờ kết quả từ 7-10 ngày và thuế suất thường cao hơn.",
        "persona": "Anh Dũng (Trưởng phòng Xuất nhập khẩu doanh nghiệp thực phẩm / F&B)",
        "checklist": "☐ Yêu cầu xưởng cung cấp Spec công thức chi tiết từng gam\n☐ Tính toán tỷ lệ % thành phần thịt theo công thức chuẩn\n☐ Chuẩn bị hồ sơ đăng ký kiểm tra thú y nếu thịt > 20%\n☐ Lưu trữ mẫu thử nghiệm thực tế đối chiếu\n☐ Nộp hồ sơ xác định trước mã số nếu mặt hàng nhạy cảm",
        "url": "https://en.wikipedia.org/wiki/Tariff",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn tham khảo: Danh mục HS WCO, Công văn 22883/CHQ-NVTHQ, Thông tư 31/2022/TT-BTC & bài học thực chiến từ daily_notes.md (14/10/2026).",
        "keywords": ["tariff", "customs"]
    }
]
