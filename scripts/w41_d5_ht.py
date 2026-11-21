# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W41 DAY 5 — 10/10/2026)
Số lượng: 5 ý tưởng (STT 141 - 145)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W41_D5 = [
    {
        "stt": 141,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Hướng Dẫn 4 Bước Chuẩn Chỉ Gia Hạn Tờ Khai Tạm Nhập - Tái Xuất (Mã G12/G22): Trình Tự Song Trùng DVC Và Chỉ Tiêu TIA Trên VNACCS Chống Bị Ấn Định Thuế",
        "workflow": "Bước 1: Rà soát điều kiện pháp lý và chuẩn bị hồ sơ: Xác định lý do gia hạn khách quan (đối tác hoãn tiến độ, hội thảo kéo dài, hỏng hóc cần sửa chữa...), chuẩn bị Phụ lục hợp đồng gia hạn hoặc văn bản xác nhận của bên nước ngoài; Bước 2: Nộp văn bản đề nghị gia hạn trên Cổng Dịch vụ công trực tuyến / Một cửa Quốc gia (NSW): Gửi hồ sơ trước ngày hết hạn tạm nhập tối thiểu 10 ngày làm việc để Chi cục Hải quan nơi đăng ký tờ khai xem xét phê duyệt; Bước 3: Nhận văn bản chấp thuận gia hạn của Chi cục Hải quan: Theo dõi trạng thái trên DVC để nhận công văn chấp thuận có ghi rõ thời hạn gia hạn mới; Bước 4: Khai báo nghiệp vụ TIA trên phần mềm VNACCS/VCIS: Nhập mã tờ khai tạm nhập gốc, số công văn chấp thuận gia hạn của Hải quan, nhập ngày hết hạn mới tại chỉ tiêu gia hạn và gửi thông điệp TIA để hệ thống tự động cập nhật thời hạn bảo lãnh thuế.",
        "formula": "Thời hạn nộp hồ sơ xin gia hạn: T_nop <= T_hethan - 10 ngày làm việc. Thời hạn gia hạn tối đa: Thường không quá thời hạn tạm nhập ban đầu và phù hợp với mục đích sử dụng. Phạt chậm nộp nếu quá hạn: Tiền chậm nộp = Số thuế ấn định x 0.03%/ngày x Số ngày quá hạn.",
        "top3_mistakes": "1. Khai báo chỉ tiêu TIA trên VNACCS trước khi có văn bản chấp thuận của Chi cục Hải quan (dẫn đến bị hệ thống từ chối hoặc treo tờ khai); 2. Nộp công văn xin gia hạn quá sát ngày hết hạn (dưới 3 ngày) khiến hồ sơ không kịp duyệt và rơi vào trạng thái quá hạn nợ thuế; 3. Quên không làm thủ tục gia hạn thư bảo lãnh ngân hàng đối ứng tương ứng với thời hạn tạm nhập mới.",
        "mnemonic_tip": "'DVC duyệt trước - TIA khai sau - Bảo lãnh nối hạn - Thuế khỏi lo âu': Không bao giờ làm ngược quy trình khai TIA trước khi có công văn chấp thuận.",
        "comparison": "Tờ khai Tạm nhập tái xuất G12/G22 (Được miễn thuế tạm thời, quản lý theo thời hạn nghiêm ngặt, phải gia hạn hoặc tái xuất) vs Tờ khai Nhập kinh doanh tiêu dùng A11 (Nộp đủ thuế ngay khi thông quan, tự do lưu thông vĩnh viễn không giới hạn thời gian).",
        "persona": "Chị Hương (Quản lý XNK) & Anh Dũng (Giám đốc kỹ thuật phụ trách máy móc thi công dự án, triển lãm).",
        "checklist": "☐ Kiểm tra hạn tờ khai gốc trước 15 ngày\n☐ Soạn phụ lục hợp đồng gia hạn có ký đóng dấu\n☐ Nộp đơn đề nghị gia hạn trên Cổng Dịch vụ công trước hạn 10 ngày\n☐ Nhận công văn phê duyệt chính thức của Chi cục Hải quan\n☐ Truyền chỉ tiêu TIA trên phần mềm VNACCS và đồng bộ bảo lãnh ngân hàng.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Điều 82 Thông tư 38/2015/TT-BTC, Thông tư 39/2018/TT-BTC & daily_notes.md (ghi nhận ngày 10/10/2026 từ case study Nguyen Quoc Chinh).",
        "keywords": ["customs declaration", "customs"]
    },
    {
        "stt": 142,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Checklist 5 Tiêu Chuẩn Pháp Lý Khi Thuê Dịch Vụ Mua Hộ, Order & Vận Chuyển Hàng Trung Quốc: Cách Kiểm Soát Tờ Khai Chính Ngạch Và Chống Mất Trắng Tiền Hàng",
        "workflow": "Bước 1: Thẩm định pháp nhân đơn vị cung cấp dịch vụ: Yêu cầu cung cấp Giấy chứng nhận ĐKKD có ngành nghề Đại lý làm thủ tục hải quan và Vận tải quốc tế, có trụ sở văn phòng và kho bãi thực tế tại Việt Nam; Bước 2: Ký Hợp đồng Dịch vụ Ủy thác Nhập khẩu / Logistics minh bạch: Ràng buộc rõ ràng trách nhiệm pháp lý, cam kết cung cấp Tờ khai hải quan thông quan chính ngạch và Hóa đơn GTGT đầu vào hợp lệ; Bước 3: Minh bạch cơ cấu chi phí (No Hidden Fees): Yêu cầu bảng báo giá chi tiết gồm 3 phần riêng biệt: Tiền hàng xưởng gốc (theo tỷ giá ngân hàng niêm yết), Tiền cước vận chuyển (tính theo CBM hoặc KG thực tế), và Tiền thuế XNK (tính đúng theo biểu thuế của mã HS); Bước 4: Kiểm tra và đối soát bộ chứng từ tờ khai sau khi thông quan: Yêu cầu gửi bản sao tờ khai hải quan điện tử có mã vạch thông quan, đối chiếu tên người xuất khẩu, người nhập khẩu và giá trị khai báo; Bước 5: Cài điều khoản bồi thường thiệt hại 100%: Cam kết đền bù toàn bộ giá trị đơn hàng nếu hàng hóa bị thu giữ do lỗi kê khai sai của bên vận chuyển.",
        "formula": "Tổng chi phí nhập khẩu chính ngạch chuẩn = Trị giá tiền hàng CIF (Quy đổi theo tỷ giá thanh toán ngân hàng) + Thuế nhập khẩu + Thuế GTGT hàng nhập khẩu + Cước vận chuyển nội địa & Phí dịch vụ khai thuê hải quan. Tuyệt đối không chấp nhận công thức 'Cước = % x Trị giá hàng'.",
        "top3_mistakes": "1. Tin vào lời chào mời 'Bao trọn gói cước chỉ 3-5% giá trị đơn hàng' (thực chất là gộp đơn lậu); 2. Chuyển tiền cọc và tiền mua hàng vào số tài khoản cá nhân của môi giới mà không có hợp đồng ủy thác; 3. Không yêu cầu xuất hóa đơn GTGT đầu vào khiến hàng hóa lưu thông nội địa bị coi là hàng lậu.",
        "mnemonic_tip": "'Pháp nhân chuẩn - Hợp đồng rõ - Tờ khai riêng - Hóa đơn đỏ - Đền bù 100%': 5 ngón tay vàng để không bao giờ bị mất tiền oan khi nhập hàng Trung Quốc.",
        "comparison": "Dịch vụ Gom đơn phi mậu dịch trôi nổi (Rủi ro mất trắng hàng, bị phạt buôn lậu, không có hóa đơn khấu trừ thuế) vs Dịch vụ Nhập khẩu ủy thác chính ngạch (Tờ khai hải quan hợp pháp, hóa đơn GTGT đầy đủ, an toàn tuyệt đối trước Quản lý thị trường).",
        "persona": "Anh Minh (Chủ shop kinh doanh TMĐT, phụ kiện, đồ gia dụng chuyển đổi sang chính ngạch).",
        "checklist": "☐ Kiểm tra mã số thuế và giấy phép đại lý hải quan của forwarder\n☐ Ký hợp đồng ủy thác nhập khẩu có dấu pháp nhân\n☐ Nhận báo giá bóc tách rõ cước vận chuyển (kg/cbm) và tiền thuế riêng\n☐ Yêu cầu cung cấp tờ khai hải quan thông quan bản điện tử\n☐ Lấy hóa đơn GTGT hợp lệ trong vòng 7 ngày sau khi nhận hàng.",
        "url": "https://en.wikipedia.org/wiki/Freight_forwarder",
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Luật Thương mại 2005, Luật Hải quan 2014 & daily_notes.md (ghi nhận ngày 10/10/2026 từ case study Viettel Post gom đơn).",
        "keywords": ["freight forwarder", "cargo"]
    },
    {
        "stt": 143,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Cẩm Nang 4 Bước Lập Bảng Kê Thu Mua Mẫu 01/TNDN Chuẩn Pháp Lý: Điều Kiện Đưa Vào Chi Phí Hợp Lý Được Trừ Và Ranh Giới Chống Bị Quy Tội Buôn Lậu",
        "workflow": "Bước 1: Xác định chính xác đối tượng được phép lập Bảng kê 01/TNDN: Chỉ áp dụng khi mua nông, lâm, thủy sản của người sản xuất, đánh bắt trực tiếp bán ra; mua sản phẩm thủ công mỹ nghệ làm bằng đay, cói, tre, nứa của người sản xuất trực tiếp; mua phế liệu của người trực tiếp thu nhặt; mua tài sản của hộ, cá nhân không kinh doanh; Bước 2: Kiểm tra ngưỡng doanh thu và nhân thân người bán: Thu thập bản photocopy CCCD của người bán trực tiếp, xác nhận địa chỉ cư trú; đối với hộ cá nhân kinh doanh doanh thu bán hàng phải dưới ngưỡng 100 triệu đồng/năm; Bước 3: Lập đầy đủ Bảng kê thu mua theo Mẫu 01/TNDN (Thông tư 78/2014/TT-BTC): Ghi rõ họ tên, địa chỉ, số CCCD người bán; tên hàng hóa, số lượng, đơn giá, tổng số tiền thanh toán; chữ ký trực tiếp của người bán và người đại diện theo pháp luật của doanh nghiệp ký duyệt; Bước 4: Hoàn thiện chứng từ thanh toán đi kèm: Lập Phiếu chi tiền mặt hoặc Ủy nhiệm chi ngân hàng khớp từng đồng với số tiền trên bảng kê; kèm theo Phiếu nhập kho (Mẫu 01-VT) có chữ ký của thủ kho và người giao hàng.",
        "formula": "Chi phí được trừ hợp lý thuế TNDN = Tổng giá trị trên Bảng kê 01/TNDN (nếu đáp ứng đủ 4 điều kiện: Đúng đối tượng quy định + Có CCCD người bán + Có Bảng kê ký thật + Có Phiếu chi & Phiếu nhập kho hợp lệ). Nếu vi phạm: Chi phí bị loại trừ 100% + Truy thu 20% thuế TNDN + Phạt 20% khai sai.",
        "top3_mistakes": "1. Lập Bảng kê 01/TNDN để mua hàng hóa có xuất xứ nước ngoài (đồ gia dụng, máy móc, linh kiện Trung Quốc trôi nổi) - đây là hành vi tiếp tay cho buôn lậu; 2. Bảng kê không có chữ ký của người bán thực tế hoặc dùng số CCCD giả; 3. Không có phiếu nhập kho thực tế chứng minh hàng hóa đã nhập vào kho doanh nghiệp.",
        "mnemonic_tip": "'Đúng người trực tiếp làm ra - Đủ thẻ căn cước - Ký tên thật thà - Nhập kho rõ ràng - Thuế mới cho qua': 5 câu khẩu quyết khi lập Bảng kê 01/TNDN.",
        "comparison": "Bảng kê Mẫu 01/TNDN (Áp dụng hợp pháp cho nông sản tự sản xuất, phế liệu nhặt trực tiếp, tài sản cá nhân thanh lý) vs Mua hàng nhập khẩu không hóa đơn (Bị coi là hành vi kinh doanh hàng nhập lậu theo Nghị định 98/2020/NĐ-CP, tuyệt đối không được dùng Bảng kê 01 để hợp thức hóa).",
        "persona": "Chị Hương (Kế toán trưởng / Quản lý XNK) & Anh Dũng (Chủ doanh nghiệp chế biến nông sản, thương mại).",
        "checklist": "☐ Đối chiếu mặt hàng có thuộc danh mục được lập bảng kê Điều 6 Thông tư 78\n☐ Thu thập bản sao CCCD của người trực tiếp sản xuất/bán hàng\n☐ Lập bảng kê 01/TNDN đầy đủ chữ ký 2 bên và đại diện pháp luật\n☐ Lập phiếu chi tiền mặt hoặc ủy nhiệm chi khớp đúng số tiền\n☐ Lập phiếu nhập kho có chữ ký thủ kho lưu hồ sơ kế toán.",
        "url": "https://en.wikipedia.org/wiki/Tax_deduction",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Thông tư 78/2014/TT-BTC, Thông tư 96/2015/TT-BTC & daily_notes.md (ghi nhận ngày 10/10/2026 từ case study Nam Vũ).",
        "keywords": ["tax deduction", "expense"]
    },
    {
        "stt": 144,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Quy Trình 5 Bước Đăng Ký Kiểm Tra An Toàn Thực Phẩm Nhập Khẩu Trên Cổng Một Cửa Quốc Gia (NSW): Bí Quyết Đạt Tiêu Chí Kiểm Tra Giảm Trong 24 Giờ",
        "workflow": "Bước 1: Rà soát hồ sơ công bố và chuẩn hóa chứng từ: Kiểm tra Bản tự công bố sản phẩm còn hiệu lực, đối chiếu từng chữ trên Giấy chứng nhận lưu hành tự do (CFS) và kết quả kiểm nghiệm (Certificate of Analysis - COA) khớp chính xác với nhãn hàng hóa; Bước 2: Tạo hồ sơ đăng ký kiểm tra an toàn thực phẩm trên Cổng Một cửa Quốc gia (national.singlewindow.gov.vn): Chọn đúng Bộ ngành quản lý (Bộ Y tế, Bộ Nông nghiệp & PTNT hoặc Bộ Công Thương tùy theo nhóm mặt hàng) và đơn vị kiểm tra được chỉ định; Bước 3: Đăng ký áp dụng phương thức Kiểm tra giảm (nếu đủ điều kiện): Đính kèm chứng chỉ ISO 22000/HACCP/GMP còn hiệu lực hoặc 3 Thông báo kết quả kiểm tra đạt yêu cầu của 3 lô hàng nhập khẩu liền kề trước đó; Bước 4: Nhận Thông báo kết quả xác nhận đạt yêu cầu nhập khẩu điện tử: Đối với diện kiểm tra giảm, cơ quan kiểm tra chỉ kiểm tra xác suất hồ sơ và trả kết quả tự động trên NSW trong vòng 2 ngày làm việc mà không phải lấy mẫu thực tế; Bước 5: Đồng bộ thông quan tờ khai trên VNACCS: Hệ thống NSW tự động đẩy số Thông báo đạt yêu cầu sang phần mềm Hải quan để giải phóng hàng hóa ngay tại cảng.",
        "formula": "Tỷ lệ kiểm tra thực tế: Phương thức Kiểm tra giảm = Tối đa 5% tổng số lô hàng ngẫu nhiên do Hải quan chỉ định; Phương thức Kiểm tra thông thường = 100% kiểm tra hồ sơ; Phương thức Kiểm tra chặt = 100% kiểm tra hồ sơ kết hợp lấy mẫu kiểm nghiệm hóa lý/vi sinh vật tại hiện trường.",
        "top3_mistakes": "1. Đăng ký sai cơ quan quản lý chuyên ngành được phân công theo Nghị định 15/2018/NĐ-CP (ví dụ: thực phẩm chức năng thuộc Bộ Y tế nhưng lại nộp sang Bộ Công Thương); 2. Địa chỉ nhà sản xuất trên CFS khác với địa chỉ trên bản Tự công bố sản phẩm; 3. Nộp hồ sơ NSW sau khi tàu đã cập cảng khiến phát sinh tiền lưu bãi container.",
        "mnemonic_tip": "'Công bố chuẩn trước - Đúng Bộ ngành giao - Đính kèm HACCP - Kiểm tra giảm mau - Thông quan một cửa': 5 bước để giải phóng hàng thực phẩm trong 24 giờ.",
        "comparison": "Kiểm tra an toàn thực phẩm giảm (Chỉ duyệt hồ sơ xác suất, không lấy mẫu, giải phóng hàng trong 24h) vs Kiểm tra chặt (Lấy mẫu kiểm nghiệm tại cảng, chờ kết quả phòng lab 5-7 ngày, tốn kém chi phí lưu kho lạnh).",
        "persona": "Chị Hương (Quản lý XNK chuỗi F&B, sữa, thực phẩm dinh dưỡng) & Anh Minh (Chủ shop nhập khẩu bánh kẹo).",
        "checklist": "☐ Kiểm tra mã số tự công bố sản phẩm trên hệ thống an toàn thực phẩm\n☐ Đối chiếu CFS, COA trùng khớp 100% với nhãn gốc và nhãn phụ tiếng Việt\n☐ Đính kèm chứng chỉ HACCP/ISO 22000 của nhà máy ngoại\n☐ Nộp đơn đăng ký trên Cổng Một cửa Quốc gia trước khi tàu cập cảng 3 ngày\n☐ Theo dõi trạng thái tiếp nhận và nhận giấy chứng nhận điện tử trên NSW.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Nghị định 15/2018/NĐ-CP, Luật An toàn thực phẩm 2010 & Hệ thống Một cửa Quốc gia NSW.",
        "keywords": ["standard", "food safety"]
    },
    {
        "stt": 145,
        "week": "W41 Day 5 — 10/10/2026",
        "title": "Cẩm Nang 6 Bước Khai Báo Và Đóng Gói Hàng Nguy Hiểm (Dangerous Goods - DG) Đường Biển Theo Bộ Luật IMDG Code: An Toàn Tuyệt Đối Cho Hàng Pin Và Hóa Chất",
        "workflow": "Bước 1: Tra cứu Bảng chỉ dẫn an toàn hóa chất (MSDS/SDS): Kiểm tra Mục 14 (Transport Information) để xác định 4 thông số cốt lõi: Số hiệu UN (UN Number), Tên gọi vận chuyển chuẩn (Proper Shipping Name - PSN), Phân nhóm nguy hiểm (DG Class từ 1 đến 9), và Nhóm đóng gói (Packing Group I, II hoặc III); Bước 2: Sử dụng bao bì đóng gói đạt chuẩn UN (UN Certified Packaging): Thùng carton, phuy sắt hoặc IBC phải có dấu chứng nhận kiểm định UN (ví dụ: UN 4G/Y...) chịu được áp lực rơi và rò rỉ; Bước 3: Dán nhãn cảnh báo nguy hiểm (DG Labels) và đánh dấu (Marking): Dán nhãn hình thoi biểu trưng nguy hiểm và nhãn cảnh báo ô nhiễm biển (Marine Pollutant) ở tối thiểu 2 mặt đối diện của kiện hàng và cả 4 mặt ngoài container; Bước 4: Lập Bản khai hàng nguy hiểm đa phương thức (Multimodal Dangerous Goods Form - DGD): Điền đầy đủ thông tin kỹ thuật, ký tên cam kết người gửi hàng có chứng chỉ huấn luyện DG; Bước 5: Nộp hồ sơ xin chấp thuận vận chuyển (DG Approval) từ Hãng tàu: Gửi bản DGD và MSDS cho hãng tàu trước khi đóng hàng tối thiểu 3-5 ngày; Bước 6: Giám sát đóng hàng và chèn lót (Lashing & Securing): Kiểm tra container khô ráo, không thủng rách, chèn lót cố định chống va đập và bấm niêm phong chì seal chuyên dụng.",
        "formula": "Phân loại 9 nhóm hàng nguy hiểm IMDG: Class 1 (Chất nổ) -> Class 2 (Khí ga) -> Class 3 (Chất lỏng dễ cháy) -> Class 4 (Chất rắn dễ cháy) -> Class 5 (Chất oxy hóa) -> Class 6 (Chất độc & lây nhiễm) -> Class 7 (Chất phóng xạ) -> Class 8 (Chất ăn mòn) -> Class 9 (Hàng nguy hiểm khác, bao gồm Pin Lithium UN 3480/3481).",
        "top3_mistakes": "1. Khai man tên hàng thành hàng bách hóa thông thường để né phụ phí DG (bị phạt hình sự nếu xảy ra hỏa hoạn); 2. Dùng bao bì thường không có dấu chứng nhận kiểm định UN Packaging; 3. Không dán đủ 4 nhãn cảnh báo hình thoi ở 4 mặt bên ngoài thùng container.",
        "mnemonic_tip": "'Tra mã UN - Chuẩn thùng chứng chỉ - Dán nhãn 4 bên - Ký bản khai DGD - Hãng tàu duyệt mới đi': Quy tắc sinh tử khi xuất khẩu hàng nguy hiểm.",
        "comparison": "Hàng nguy hiểm khai báo chính ngạch theo IMDG Code (Được bố trí xếp trên boong hoặc khoang cách ly an toàn, có hệ thống chữa cháy chuyên dụng, được bảo hiểm đền bù 100%) vs Gian lận giấu nhãn DG (Bị xếp sâu dưới hầm hàng bí khí, nguy cơ cháy nổ thảm họa, bảo hiểm từ chối bồi thường và chủ hàng bị kiện phá sản).",
        "persona": "Anh Dũng (Giám đốc công ty sản xuất pin, hóa chất, sơn, vật liệu công nghiệp) & Chị Hương (Quản lý logistics).",
        "checklist": "☐ Lấy MSDS bản cập nhật 16 mục từ nhà sản xuất\n☐ Kiểm tra mã UN Number và DG Class tại Mục 14\n☐ Đóng hàng bằng bao bì có dấu mộc UN Certified\n☐ Điền và ký Bản khai hàng nguy hiểm DGD chuẩn mẫu IMO\n☐ Dán 4 nhãn cảnh báo Class tương ứng lên 4 mặt ngoài container.",
        "url": "https://en.wikipedia.org/wiki/International_Maritime_Dangerous_Goods_Code",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Căn cứ Bộ luật IMDG Code của IMO, Công ước SOLAS 1974 & Nghị định 34/2024/NĐ-CP về vận chuyển hàng nguy hiểm.",
        "keywords": ["dangerous goods", "maritime"]
    }
]
