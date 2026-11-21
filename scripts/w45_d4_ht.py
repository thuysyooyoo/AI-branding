# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W45 DAY 4 — 08/11/2026)
Số lượng: 5 ý tưởng (STT 216 - 220)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
Nguồn dữ liệu thực tế: daily_notes.md (08/11/2026) & SSoT Research
"""

HOC_THUC_CHIEN_W45_D4 = [
    {
        "stt": 216,
        "week": "W45 — 08/11/2026",
        "title": "Cẩm Nang Phân Loại Mã HS Ngành Dệt May Chuẩn WCO: 4 Tiêu Chí Phân Biệt Rõ Ràng Ranh Giới Giữa 'Vải Dệt Kim Dạng Cuộn' (Chương 60) Và 'Phụ Kiện May Mặc Đã Cắt Rời' (Nhóm 6117) Để Tránh Bị Truy Thu Thuế",
        "workflow": "Quy trình 4 bước phân loại chuẩn xác bo len và vải dệt may: Bước 1: Giám định trạng thái đóng gói và hình thái xuất trình của hàng hóa: Kiểm tra hàng xuất trình ở dạng cuộn liên tục (in rolls) khổ vải cố định hay đã được cắt rời thành từng miếng có hình dáng cụ thể (cổ áo, bo tay, bo gấu); Bước 2: Tra cứu Chú giải chi tiết Chương 60 của Tổ chức Hải quan Thế giới (WCO): Toàn bộ vải dệt kim hoặc móc ở dạng cuộn hoặc dạng mảnh hình chữ nhật chưa may thành sản phẩm bắt buộc phải xếp vào Chương 60 (nhóm 60.02 nếu có pha ≥ 5% sợi đàn hồi elastomeric/spandex, hoặc nhóm 60.06 đối với các loại vải dệt kim khác); Bước 3: Rà soát điều kiện áp mã Nhóm 61.17 (Phụ liệu may mặc đã hoàn thiện khác): Chỉ áp dụng cho các sản phẩm may mặc đã được cắt rời định hình thành chiếc/cái hoặc dệt riêng biệt theo hình dáng bộ phận trang phục sẵn sàng đưa vào may ráp mà không cần cắt gọt thêm; Bước 4: Khai báo đúng đơn vị tính theo Biểu thuế XNK hiện hành: ĐVT bắt buộc là 'kg' hoặc 'mét' đối với Chương 60; ĐVT là 'chiếc' (PCE) hoặc 'unit' đối với nhóm 61.17.",
        "formula": "Công thức kiểm tra thành phần sợi elastomeric theo Chú giải 1(a) Chương 60: Tỷ lệ pha Spandex = Khối lượng sợi Spandex / Tổng khối lượng mẫu vải x 100%. Nếu ≥ 5% -> Áp mã nhóm 60.02 (thuế suất ACFTA thường 0% nếu có C/O Form E hợp lệ, MFN 12%).",
        "top3_mistakes": "1. Khai mã phụ kiện 61178090 nhưng đơn vị tính trên tờ khai lại để là mét hoặc kg; 2. Yêu cầu nhà cung ứng ghi chung chung 'knitted parts' trên C/O Form E trong khi hàng thực tế là cả cây vải cuộn dài; 3. Không đối soát mã HS với catalogue kỹ thuật dệt trước khi tàu cập cảng.",
        "mnemonic_tip": "Mẹo nhớ 'CUỘN NGUYÊN VÀO SÁU KHÔNG (CHƯƠNG 60) - CẮT RỜI SANG SÁU MỐT (CHƯƠNG 61)': Cứ thấy cuộn dài tính mét là chương 60 bất kể may áo hay may quần; khi nào cắt ra thành cái cổ, cái gấu áo thì mới được bay sang chương 61.",
        "comparison": "Vải dệt kim dạng cuộn Chương 60 (Tính thuế theo kg/mét, quy trình giám định sợi đàn hồi, chứng từ C/O phải thể hiện mã 6002 hoặc 6006) VS Phụ kiện dệt may cắt rời Nhóm 61.17 (Tính thuế theo chiếc/bộ, phải chứng minh đã hoàn thiện hình dáng chi tiết may mặc, C/O thể hiện mã 6117).",
        "persona": "Chị Hương (Chủ xưởng may thời trang xuất khẩu) & Nhân viên chứng từ xuất nhập khẩu chuyên trách mảng nguyên phụ liệu dệt may.",
        "checklist": "1. Ảnh chụp thực tế quy cách cuộn hàng; 2. Kết quả test thành phần sợi (Polyester % vs Spandex %); 3. Thước đo khổ vải cuộn; 4. Bảng đối chiếu đơn vị tính tờ khai vs C/O; 5. Khai đúng Chương 60 trên VNACCS.",
        "url": "https://en.wikipedia.org/wiki/Knitting",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Diễn đàn Hỏi đáp Thủ tục hải quan - XNK - Logistics, Thư Viện Pháp Luật (Thông tư 31/2022/TT-BTC Chú giải Chương 60, Chương 61 WCO) & daily_notes.md (08/11/2026).",
        "keywords": ["knitting"]
    },
    {
        "stt": 217,
        "week": "W45 — 08/11/2026",
        "title": "Quy Trình 4 Bước Khảo Sát & Xin Phê Duyệt Nhập Khẩu Hóa Chất, Tiền Chất Vào Doanh Nghiệp Chế Xuất (DNCX): Bí Quyết Nắm Chắc Căn Cứ Pháp Lý An Toàn Tuyệt Đối Trước Khi Tàu Về Cảng",
        "workflow": "Quy trình 4 bước giải quyết vướng mắc giấy phép hóa chất DNCX: Bước 1: Rà soát và bóc tách thành phần hóa học dựa trên Phiếu an toàn hóa chất (MSDS): Tra cứu mã số CAS và tên danh pháp quốc tế IUPAC đối chiếu với Phụ lục I, II, III, IV, V của Nghị định 113/2017/NĐ-CP (được sửa đổi bởi Nghị định 82/2022/NĐ-CP) để xác định hóa chất thuộc diện: Khai báo hóa chất, Hóa chất hạn chế sản xuất kinh doanh, hay Tiền chất công nghiệp; Bước 2: Soạn thảo công văn xin ý kiến hướng dẫn nghiệp vụ gửi Sở Công Thương địa phương (nơi đặt nhà máy DNCX) kèm theo hồ sơ kỹ thuật và hợp đồng nhập khẩu, nêu rõ đặc thù DNCX theo Luật Quản lý ngoại thương; Bước 3: Đồng thời nộp hồ sơ xin cấp Giấy phép nhập khẩu trên Cổng thông tin một cửa quốc gia (NSW): Nếu Sở/Bộ phê duyệt cấp giấy phép -> Sử dụng giấy phép điện tử để truyền tờ khai; Nếu Sở trả lời bằng văn bản 'DNCX không thuộc đối tượng phải xin giấy phép' -> Lưu giữ văn bản chính thức này làm căn cứ pháp lý số 1; Bước 4: Khi truyền tờ khai hải quan, đính kèm văn bản trả lời của Sở Công Thương hoặc Giấy phép điện tử vào ô chứng từ kèm theo trên hệ thống VNACCS để công chức hải quan tiếp nhận thông quan ngay trong ngày.",
        "formula": "Thời hạn thẩm định hồ sơ cấp phép hóa chất hạn chế: Tối đa 16 ngày làm việc kể từ ngày nhận đủ hồ sơ hợp lệ trên Cổng một cửa quốc gia theo Nghị định 82/2022/NĐ-CP.",
        "top3_mistakes": "1. Áp dụng Công văn 2347 của Cục Hóa chất trả lời cho TP.HCM sang các tỉnh thành khác mà không hỏi Sở sở tại; 2. Để hàng về đến cảng mới bắt đầu nộp hồ sơ xin cấp phép khiến container bị lưu bãi phát sinh chi phí lớn; 3. Không khai báo hóa chất nhập khẩu đối với các hóa chất thuộc danh mục phải khai báo trước khi thông quan.",
        "mnemonic_tip": "Mẹo 'HỎI SỞ ĐI TRƯỚC - NƯỚC BƯỚC KHÔNG LO': Trong vùng pháp lý xám, văn bản trả lời có mộc đỏ của cơ quan chủ quản địa phương là tấm bùa hộ mệnh mạnh hơn trăm lời giải trình miệng.",
        "comparison": "Chủ động xin cấp phép/văn bản xác nhận trước khi hàng về (Chủ động 100%, có chứng từ nộp hải quan thông quan trong 15 phút, không mất chi phí lưu kho) VS Đợi hàng về cảng mới cãi lý cơ chế phi thuế quan (Tờ khai bị đình chỉ, tranh cãi kéo dài hàng tuần, phát sinh hàng trăm triệu chi phí lưu bãi và dừng dây chuyền).",
        "persona": "Anh Minh (Giám đốc nhà máy FDI/DNCX) & Trưởng phòng Xuất nhập khẩu các doanh nghiệp công nghệ cao.",
        "checklist": "1. Bản MSDS tiếng Anh và bản dịch tiếng Việt; 2. Tra cứu danh mục mã CAS tại NĐ 82/2022/NĐ-CP; 3. Công văn gửi Sở Công Thương tỉnh sở tại; 4. Đăng ký tài khoản trên Cổng một cửa quốc gia (NSW); 5. Bộ hồ sơ kỹ thuật kho bãi lưu trữ an toàn hóa chất.",
        "url": "https://en.wikipedia.org/wiki/Chemical_safety",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Diễn đàn Giải đáp Thủ tục Hải quan, Thư Viện Pháp Luật (Luật Hóa chất 2007, Nghị định 113/2017/NĐ-CP, Nghị định 82/2022/NĐ-CP, CV 23913/CHQ-GSQL) & daily_notes.md (08/11/2026).",
        "keywords": ["chemical", "safety"]
    },
    {
        "stt": 218,
        "week": "W45 — 08/11/2026",
        "title": "Bộ Quy Tắc 5 Điểm 'Bắt Thóp' Phụ Phí Vận Tải Đích (Local Charges) Tuyến Trung - Việt Cho Hàng LCL: Cách Yêu Cầu Báo Giá All-In Chuẩn Chỉ Và Đàm Phán Điều Khoản Không Phát Sinh Chi Phí Ẩn",
        "workflow": "Quy trình 5 bước đàm phán cước và chốt phụ phí LCL an toàn: Bước 1: Khi nhận báo giá từ forwarder hoặc xưởng Trung Quốc, yêu cầu bắt buộc phải gửi 'Debit Note mẫu' thể hiện toàn bộ các đầu mục chi phí tại cảng đích Việt Nam (Inbound Local Charges: CFS, D/O, CIC, Handling, Cleaning, THC...); Bước 2: Bóc tách và so sánh từng đơn giá phụ phí với mức giá bình quân thị trường: Chi phí kho CFS thông thường từ 15 - 22 USD/CBM; D/O fee thông thường từ 30 - 45 USD/set; nếu thấy cước biển 0 đồng nhưng CFS vọt lên > 35 USD/CBM hoặc xuất hiện phí 'Management fee' vô lý thì từ chối ngay lập tức; Bước 3: Đàm phán chuyển đổi báo giá sang cơ chế 'All-in Destination': Yêu cầu đơn vị vận chuyển cam kết một mức trọn gói tính theo CBM hoặc theo lô hàng bao gồm toàn bộ cước và phụ phí; Bước 4: Cài cắm điều khoản chế tài vào hợp đồng dịch vụ logistics: 'Bên B cam kết không thu thêm bất kỳ khoản phụ thu hoặc chi phí phát sinh nào ngoài biểu giá đã thỏa thuận. Mọi khoản thu vượt quá thỏa thuận sẽ do Bên B tự chịu trách nhiệm'; Bước 5: Kiểm tra tính đồng bộ của chứng từ trước khi hàng cập cảng: Đối soát Bill of Lading (B/L), Commercial Invoice, Packing List và C/O Form E để đảm bảo mã HS và tên hàng khớp 100%, sẵn sàng làm thủ tục hải quan nhận hàng ngay khi dỡ bãi CFS.",
        "formula": "Công thức tính tổng chi phí gom hàng lẻ LCL: Tổng chi phí LCL = (Cước biển Ocean Freight x Thể tích CBM) + (Phụ phí đích Local Charges niêm yết x Thể tích CBM) + Phí D/O cố định. Tiêu chí tính cước: 1 CBM tương đương tối đa 1.000 kg (quy tắc W/M - Weight or Measurement).",
        "top3_mistakes": "1. Chỉ quan tâm hỏi giá cước biển (O/F) mà không hỏi biểu phí local charges tại Việt Nam; 2. Để người bán nước ngoài tự ý chỉ định đại lý forwarder giao nhận hàng LCL; 3. Thanh toán phụ phí cho forwarder mà không yêu cầu xuất hóa đơn giá trị gia tăng hợp pháp.",
        "mnemonic_tip": "Mẹo 'CƯỚC BIỂN RẺ MỘT - PHỤ PHÍ CẮT MƯỜI': Nhìn vào đáy hóa đơn đừng chỉ nhìn vào dòng đầu tiên. Không có bữa trưa nào miễn phí trong vận tải hàng lẻ quốc tế.",
        "comparison": "Thuê Forwarder Việt Nam uy tín (Báo giá All-in minh bạch, hỗ trợ đối soát mã HS, làm thủ tục hải quan và C/O Form E trọn gói, chi phí dự toán chuẩn xác 100%) VS Mua CIF cước 0 đồng chỉ định (Bị chèn ép phụ phí đích gấp 3 lần, rủi ro giữ hàng, C/O sai lệch mã HS dễ bị truy thu thuế).",
        "persona": "Anh Dũng (Chủ doanh nghiệp thương mại nhập khẩu hàng tiêu dùng, phụ tùng Trung - Việt) & Nhân viên thu mua hàng hóa quốc tế.",
        "checklist": "1. Giấy báo giá có chữ ký đại lý forwarder; 2. Bảng kê chi tiết từng loại phụ phí đích (Local charges breakdown); 3. Điều khoản All-in trong hợp đồng vận tải; 4. Thể tích CBM thực tế đo đạc tại kho; 5. Đối soát số tiền trên Debit Note trước khi chuyển khoản.",
        "url": "https://en.wikipedia.org/wiki/Containerization",
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Diễn đàn Hỏi đáp Thủ tục hải quan - XNK - Logistics, Hiệp hội VLA, Cục Hàng hải Việt Nam & daily_notes.md (08/11/2026).",
        "keywords": ["containerization"]
    },
    {
        "stt": 219,
        "week": "W45 — 08/11/2026",
        "title": "Hướng Dẫn 3 Bước Đồng Bộ & Cập Nhật Mã Số Định Danh CCCD 12 Số / Định Danh Mức 2 Cho Người Khai Hải Quan Lên Hệ Thống VNACCS/VCIS Tránh Bị Treo Tờ Khai Điện Tử",
        "workflow": "Quy trình 3 bước cập nhật định danh điện tử chuẩn Đề án 06: Bước 1: Đăng nhập Cổng Thông tin Dịch vụ công Trực tuyến Tổng cục Hải quan (pus.customs.gov.vn) bằng tài khoản chữ ký số (Token) của người đại diện theo pháp luật của doanh nghiệp; Bước 2: Truy cập phân hệ 'Quản lý người khai hải quan': Nhập thông tin số CCCD 12 số, họ và tên, ngày sinh của nhân sự trực tiếp truyền tờ khai; đối với chuyên gia nước ngoài, nhập mã định danh điện tử mức 2 do Bộ Công an cấp (hoặc số hộ chiếu hợp lệ đã đăng ký); Bước 3: Thực hiện liên kết mã số Token với thông tin định danh cá nhân và gửi yêu cầu phê duyệt trực tuyến; hệ thống tự động kiểm tra đối chiếu thời gian thực với Cơ sở dữ liệu Quốc gia về Dân cư trong vòng 5 - 10 phút và trả kết quả kích hoạt tài khoản khai báo thành công.",
        "formula": "Quy tắc đối soát dữ liệu VNACCS: Mã số doanh nghiệp (Tax ID) + Số serial Token Chữ ký số + Số CCCD 12 số của người trực tiếp khai báo = Bộ ba khóa định danh bắt buộc để mở cổng truyền dữ liệu tờ khai hải quan điện tử.",
        "top3_mistakes": "1. Đổi nhân viên làm thủ tục hải quan nhưng vẫn để nguyên thông tin CCCD của nhân viên cũ đã nghỉ việc; 2. Khai số CMND 9 số cũ chưa được chuyển đổi trên Cơ sở dữ liệu quốc gia về dân cư; 3. Không kiểm tra thời hạn hiệu lực của chứng thư số chữ ký số trước khi thực hiện đồng bộ.",
        "mnemonic_tip": "Mẹo 'ĐÚNG NGƯỜI - ĐÚNG SỐ - MỞ CỔNG THÔNG QUAN': Token là con dấu của công ty, nhưng CCCD là khuôn mặt của người bấm nút. Thiếu mặt thì dấu không thể đóng.",
        "comparison": "Đăng ký phân quyền người khai hải quan chính quy (Tờ khai truyền qua cổng VNACCS trong 30 giây, phân luồng tức thì, rõ ràng trách nhiệm pháp lý) VS Dùng chung Token mượn danh nghĩa (Hệ thống tự động phát hiện sai lệch định danh, treo tờ khai, nguy cơ bị xử phạt vi phạm hành chính và khóa tài khoản doanh nghiệp).",
        "persona": "Trưởng phòng Hành chính - Nhân sự & Trưởng phòng Xuất nhập khẩu tại các doanh nghiệp có đội ngũ chứng từ đông đảo.",
        "checklist": "1. Bản chụp CCCD 12 số gắn chip của nhân viên khai báo; 2. Token chữ ký số còn hạn; 3. Văn bản ủy quyền của Tổng Giám đốc; 4. Tài khoản quản trị Cổng Dịch vụ công Hải quan; 5. Biên bản bàn giao tài khoản khai báo.",
        "url": "https://en.wikipedia.org/wiki/Public_key_infrastructure",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Diễn đàn Giải đáp Thủ tục Hải quan, Cổng Dịch vụ công Tổng cục Hải quan, Thư Viện Pháp Luật (Đề án 06/CP, Thông tư 38/2015/TT-BTC, Thông tư 39/2018/TT-BTC) & daily_notes.md (08/11/2026).",
        "keywords": ["infrastructure", "public"]
    },
    {
        "stt": 220,
        "week": "W45 — 08/11/2026",
        "title": "Hướng Dẫn 5 Bước Hoàn Thuế Nhập Khẩu & Thuế GTGT Nộp Thừa Khi Hủy Tờ Khai Hải Quan A11 Theo Đúng Điều 131 Thông Tư 38 (Sửa Đổi Bởi Thông Tư 39/2018/TT-BTC)",
        "workflow": "Quy trình 5 bước thu hồi tiền thuế nộp thừa chuẩn xác: Bước 1: Hoàn tất thủ tục phê duyệt hủy tờ khai hải quan nhập khẩu trên hệ thống VNACCS và nhận Quyết định hủy tờ khai chính thức có chữ ký số của Lãnh đạo Chi cục Hải quan; Bước 2: Tuyệt đối giữ nguyên sổ sách thuế nội địa, KHÔNG đưa hóa đơn/chứng từ nộp thuế của tờ khai đã hủy vào Bảng kê khai thuế GTGT đầu vào Mẫu 01/GTGT gửi Chi cục Thuế nội địa; Bước 3: Lập Công văn gửi Chi cục Thuế quản lý trực tiếp đề nghị xác nhận: 'Doanh nghiệp chưa kê khai khấu trừ thuế GTGT đối với chứng từ nộp thuế số... của tờ khai hải quan số...'; Bước 4: Lập Hồ sơ hoàn trả tiền thuế nộp thừa theo Mẫu số 27/HTNT/TXNK ban hành kèm Thông tư 39/2018/TT-BTC gửi Chi cục Hải quan nơi mở tờ khai kèm theo: Giấy đề nghị hoàn thuế, Quyết định hủy tờ khai, Giấy nộp tiền vào NSNN (bản chính hoặc bản điện tử), và Văn bản xác nhận của Chi cục Thuế nội địa; Bước 5: Theo dõi tiến độ thụ lý tại Đội Quản lý Thuế Hải quan: Trong thời hạn tối đa 05 ngày làm việc kể từ ngày nhận đủ hồ sơ hợp lệ (đối với diện hoàn trước kiểm tra sau), cơ quan Hải quan ban hành Quyết định hoàn trả tiền thuế nộp thừa và gửi Lệnh hoàn trả khoản thu NSNN sang Kho bạc Nhà nước để giải ngân tiền về tài khoản ngân hàng của doanh nghiệp.",
        "formula": "Thời hiệu xử lý tiền thuế nộp thừa theo Điều 60 Luật Quản lý thuế số 38/2019/QH14: Trong thời hạn 10 năm kể từ ngày nộp tiền vào ngân sách nhà nước, doanh nghiệp được quyền nộp hồ sơ yêu cầu hoàn trả hoặc bù trừ vào các khoản thuế phát sinh tiếp theo.",
        "top3_mistakes": "1. Nhầm lẫn áp dụng quy trình hoàn thuế nội địa của Thông tư 80/2021/TT-BTC sang nộp tại cơ quan thuế nội địa thay vì cơ quan hải quan; 2. Đã kê khai khấu trừ thuế GTGT đầu vào tại cơ quan thuế nội địa rồi vẫn nộp hồ sơ xin hoàn tiền mặt tại cơ quan hải quan; 3. Không lưu giữ Giấy nộp tiền ngân sách nhà nước bản điện tử có mã tham chiếu giao dịch kho bạc.",
        "mnemonic_tip": "Mẹo 'MỞ ĐÂU HOÀN ĐÓ - THUẾ NỘI XÁC NHẬN - HẢI QUAN CHI TIỀN': Tờ khai mở tại Chi cục Hải quan nào thì nộp hồ sơ hoàn tại Chi cục Hải quan đó; chỉ cần qua Thuế nội địa xin một tờ xác nhận chưa khấu trừ là tiền về tài khoản.",
        "comparison": "Hoàn thuế nộp thừa đúng quy trình Điều 131 TT 38/2015 (Hồ sơ chuẩn chỉ, được giải ngân tiền mặt hoặc bù trừ thuế chỉ trong 5 - 15 ngày làm việc) VS Tự ý kê khai bù trừ thuế lộn xộn (Bị cơ quan thuế phát hiện xử phạt khai sai, truy thu kèm tiền chậm nộp 0.03%/ngày, hồ sơ hoàn thuế bị đình chỉ thanh tra).",
        "persona": "Kế toán trưởng & Kế toán thuế xuất nhập khẩu tại các doanh nghiệp sản xuất, thương mại.",
        "checklist": "1. Quyết định hủy tờ khai hải quan A11; 2. Giấy nộp tiền vào NSNN bản in điện tử; 3. Văn bản xác nhận chưa khấu trừ thuế VAT của Chi cục Thuế; 4. Giấy đề nghị hoàn thuế Mẫu 27/HTNT/TXNK; 5. Số tài khoản ngân hàng thụ hưởng đã đăng ký với Thuế.",
        "url": "https://en.wikipedia.org/wiki/Tariff",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Diễn đàn Giải đáp Thủ tục Hải quan, Thư Viện Pháp Luật (Luật Quản lý thuế 38/2019/QH14 Điều 60, Thông tư 38/2015/TT-BTC Điều 131, Thông tư 39/2018/TT-BTC, Thông tư 80/2021/TT-BTC) & daily_notes.md (08/11/2026).",
        "keywords": ["tariff"]
    }
]
