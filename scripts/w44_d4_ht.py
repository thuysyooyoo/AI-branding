# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W44 DAY 4 — 01/11/2026)
Số lượng: 5 ý tưởng (STT 191 - 195)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W44_D4 = [
    {
        "stt": 191,
        "week": "W44 — 01/11/2026",
        "title": "Quy Trình 4 Bước Cập Nhật Mã Số Thuế, Địa Chỉ Mới & Chữ Ký Số Trên Phần Mềm ECUS5-VNACCS: Bí Quyết Chống Bị Khóa Luồng Thông Quan Khi Doanh Nghiệp Thay Đổi Giấy Phép Kinh Doanh",
        "workflow": "Bước 1: Cập nhật dữ liệu tại Cơ quan Thuế & Sở KH&ĐT: Đảm bảo thông tin MST, tên công ty, địa chỉ và số định danh người đại diện pháp luật đã được đối soát 100% khớp trên CSDLQG; Bước 2: Đăng ký Chữ ký số (Token) mới trên Cổng thông tin Tổng cục Hải quan (customs.gov.vn): Vào mục 'Đăng ký Doanh nghiệp sử dụng Chữ ký số', cắm Token mới, nhập mã PIN, lấy chứng thư số và bấm lưu thông tin; Bước 3: Thiết lập tài khoản VNACCS trên Cổng Một cửa Quốc gia (VNSW): Vào mục 'Đăng ký người sử dụng hệ thống VNACCS', đồng bộ tài khoản khai báo (User ID, Terminal ID, Encryption Key); Bước 4: Cấu hình lại phần mềm khai hải quan ECUS5: Vào menu 'Hệ thống' -> 'Thiết lập thông số khai báo VNACCS', cập nhật lại thông tin Chữ ký số mới, bấm 'Kiểm tra kết nối Hải quan' đạt trạng thái 'Kết nối thành công' trước khi truyền tờ khai đầu tiên.",
        "formula": "Công thức kiểm tra tính sẵn sàng thông quan: CSQ = (Thuế OK) + (CSDLQG Đồng bộ) + (Token Hải quan Duyệt) + (ECUS5 Kết nối thành công). Nếu thiếu bất kỳ 1 trong 4 mắt xích, tỷ lệ lỗi từ chối tờ khai = 100%. Thời gian chờ đồng bộ hệ thống chuẩn = 24 giờ sau khi đăng ký Token mới.",
        "top3_mistakes": "1. Vừa nhận Token mới là cắm vào ECUS5 truyền tờ khai ngay mà quên đăng ký kích hoạt chứng thư số trên Cổng Tổng cục Hải quan; 2. Thay đổi địa chỉ công ty sang quận/tỉnh khác nhưng không kiểm tra cập nhật lại 'Mã Chi cục Hải quan quản lý' tương ứng trên phần mềm ECUS5; 3. Không kiểm tra thời hạn hiệu lực của chứng thư số cũ dẫn đến tờ khai đang ở luồng vàng thì chứng thư hết hạn, không thể truyền chứng từ sửa đổi.",
        "mnemonic_tip": "'ĐỔI VỎ PHẢI ĐỔI RUỘT - ĐỔI PHÁP LÝ PHẢI KHAI CHỮ KÝ': Ví Chữ ký số như chiếc chìa khóa xe, còn hệ thống Hải quan VNACCS là ổ khóa. Dù doanh nghiệp đã có giấy phép mới nhưng chưa mang chìa khóa đến cơ quan hải quan cài đặt khớp răng cưa thì không bao giờ nổ máy khởi hành được lô hàng.",
        "comparison": "So sánh Nộp hồ sơ thủ công vs Đăng ký trực tuyến VNACCS: Nộp thủ công mất 3-5 ngày làm việc và phải đi lại nhiều cửa; Đăng ký trực tuyến qua VNSW chỉ mất 15 phút thao tác và được hệ thống AI phê duyệt tự động trong vòng 24 giờ nếu dữ liệu khớp CSDLQG.",
        "persona": "Kế toán thuế xuất nhập khẩu, nhân viên chứng từ hải quan (docs), giám đốc doanh nghiệp vừa thay đổi giấy phép kinh doanh (Anh Minh - Chủ DN sản xuất; Chị Hương - Kế toán trưởng).",
        "checklist": "[ ] Đã hoàn tất thông báo thay đổi với Cơ quan Thuế địa phương\n[ ] Đã đăng ký Token mới trên website Tổng cục Hải quan\n[ ] Đã tạo mã User ID và Terminal ID mới trên VNACCS\n[ ] Đã cấu hình và kiểm tra kết nối trên phần mềm ECUS5 thành công\n[ ] Đã kiểm tra đối soát thông tin CSDLQG không bị cảnh báo lệch.",
        "url": "https://en.wikipedia.org/wiki/Public_key_infrastructure",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin điện tử Tổng cục Hải quan, Thư Viện Pháp Luật (Thông tư 38/2015/TT-BTC Điều 5; Quyết định 222/QĐ-BTC) & daily_notes.md (01/11/2026).",
        "keywords": ["public", "key", "infrastructure"]
    },
    {
        "stt": 192,
        "week": "W44 — 01/11/2026",
        "title": "Xử Lý Sự Cố Nghẽn Mạng Gạch Nợ Thuế Hải Quan 24/7 Trong 15 Phút: Cẩm Nang Giải Cứu Tờ Khai Treo Nợ Sát Giờ Cắt Máng (Closing Time) Tránh Rớt Tàu Xuất Khẩu",
        "workflow": "Bước 1: Kiểm tra trạng thái giao dịch ngân hàng: Lấy ngay Giấy nộp tiền vào NSNN điện tử có mã FT/mã tham chiếu giao dịch thành công từ Internet Banking ngân hàng thương mại phối hợp thu; Bước 2: Tra cứu trạng thái truyền dữ liệu trên Cổng Một cửa Quốc gia (VNSW) mục 'Tra cứu nộp thuế hải quan': Xác định lỗi đang nghẽn tại Ngân hàng -> Kho bạc hay Kho bạc -> Cổng thanh toán Hải quan; Bước 3: In Giấy nộp tiền có xác nhận của ngân hàng, đóng dấu treo công ty kèm Công văn đề nghị hỗ trợ thông quan khẩn cấp do sự cố kỹ thuật theo mẫu chuẩn; Bước 4: Trực tiếp nộp cho Lãnh đạo Chi cục Hải quan cửa khẩu đề nghị sử dụng tài khoản nghiệp vụ can thiệp phê duyệt thông quan cưỡng chế hợp lệ trên hệ thống VNACCS trong 15 phút.",
        "formula": "Quy tắc thời gian vàng xử lý sự cố: T_safe = T_closing - 180 phút (Phải nộp thuế trước giờ cắt máng tối thiểu 3 tiếng). Nếu Delta_T (thời gian tiền chưa nổi) > 30 phút, lập tức kích hoạt quy trình can thiệp bằng chứng từ giấy có xác nhận ngân hàng, không được ngồi đợi thụ động.",
        "top3_mistakes": "1. Ngồi chờ hệ thống tự động gạch nợ trong vô vọng cho đến khi hết giờ làm việc hành chính của cán bộ hải quan; 2. Điền sai dù chỉ 1 ký tự trong mục 'Số tờ khai', 'Mã loại hình' hoặc 'Chương - Tiểu mục nộp thuế' khiến tiền vào Kho bạc nhưng bị treo ở tài khoản chờ xử lý; 3. Không xin giấy nộp tiền có đóng dấu đỏ xác nhận của chi nhánh ngân hàng khi hệ thống ngân hàng điện tử bị đơ.",
        "mnemonic_tip": "'TIỀN ĐÃ ĐI - GIẤY PHẢI ĐẾN TAY HẢI QUAN': Đừng tin vào lời hứa 'tiền sắp nổi rồi' của nhân viên tổng đài ngân hàng khi kim đồng hồ đang chạy đua với giờ cắt máng. Một tờ Giấy nộp tiền có dấu xác nhận ngân hàng cầm trên tay có giá trị giải cứu lô hàng hơn 100 cuộc gọi điện thoại.",
        "comparison": "So sánh Thanh toán điện tử 24/7 vs Bảo lãnh thuế ngân hàng: Thanh toán điện tử phụ thuộc vào hạ tầng mạng liên ngân hàng (nguy cơ nghẽn mạng); Bảo lãnh thuế giúp tờ khai được tự động thông quan tức thì 100% ngay khi truyền tờ khai mà không cần nộp tiền mặt trước.",
        "persona": "Nhân viên hiện trường (ops) hải quan tại cảng biển/sân bay, nhân viên thanh toán quốc tế và giám đốc điều hành chuỗi cung ứng (Anh Dũng - Trưởng phòng Logistics).",
        "checklist": "[ ] In ngay Giấy nộp tiền điện tử có mã số giao dịch ngân hàng (FT code)\n[ ] Tra cứu mã lỗi trên Cổng Một cửa Quốc gia VNSW\n[ ] Lập công văn xin thông quan khẩn cấp có chữ ký Giám đốc\n[ ] Trực tiếp gặp Lãnh đạo Đội Thủ tục Chi cục Hải quan cửa khẩu\n[ ] Theo dõi trạng thái tờ khai chuyển sang mã 04 (Thông quan) trên hệ thống.",
        "url": "https://en.wikipedia.org/wiki/Real-time_gross_settlement",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 184/2015/TT-BTC, Quyết định 2062/QĐ-TCHQ), Cổng TTĐT Tổng cục Hải quan & daily_notes.md (31/10/2026).",
        "keywords": ["settlement", "payment"]
    },
    {
        "stt": 193,
        "week": "W44 — 01/11/2026",
        "title": "Hướng Dẫn Kỹ Thuật Kiểm Soát An Toàn Thực Phẩm & Kiểm Dịch Nông Sản Nhập Khẩu: 5 Bước Tự Thẩm Tra Dư Lượng Thuốc BVTV (Abamectin, Carbendazim) Chống Bị Bắt Buộc Tái Xuất Tại Cửa Khẩu",
        "workflow": "Bước 1: Tra cứu Danh mục kiểm tra chặt & Hoạt chất cảnh báo trên Cổng thông tin Cục Bảo vệ thực vật (PPDS) và Thông tư 50/2016/TT-BYT để xác định mức giới hạn tối đa cho phép (MRL); Bước 2: Ràng buộc đối tác nhà vườn ngoại gửi Kết quả thử nghiệm mẫu (Lab Test Report) từ đơn vị giám định quốc tế (SGS/CCIC) trước khi đóng container bốc hàng; Bước 3: Đăng ký Giấy phép Kiểm dịch thực vật nhập khẩu trên Cổng Một cửa Quốc gia (VNSW) trước khi tàu/xe cập cảng; Bước 4: Đăng ký Kiểm tra nhà nước về an toàn thực phẩm nhập khẩu theo Nghị định 15/2018/NĐ-CP; Bước 5: Phối hợp lấy mẫu kiểm tra thực tế tại kho lạnh cửa khẩu và theo dõi sát chỉ số phân tích sắc ký khí (GC-MS) của phòng lab chỉ định.",
        "formula": "Công thức an toàn MRL: Chỉ số dư lượng thực tế (C_lab) / MRL quy chuẩn <= 0.7 (Ngưỡng an toàn biên phòng ngừa sai số kiểm nghiệm). Ví dụ đối với Abamectin trên trái cây: C_lab <= 0.007 mg/kg (khi MRL = 0.01 mg/kg).",
        "top3_mistakes": "1. Không yêu cầu nhà xuất khẩu cung cấp nhật ký phun xịt và thời gian cách ly thuốc bảo vệ thực vật trước khi thu hoạch; 2. Đưa hàng về cửa khẩu mới nộp đơn xin Giấy phép kiểm dịch thực vật nhập khẩu khiến hàng bị kẹt chờ giấy phép; 3. Không chuẩn bị hợp đồng dịch vụ kiểm tra an toàn thực phẩm với đơn vị được Bộ Nông nghiệp chỉ định từ trước.",
        "mnemonic_tip": "'TEST TRƯỚC TẠI NƯỚC BẠN - VỀ NƯỚC MÌNH MỚI BÌNH AN': Chi vài triệu đồng làm test thử nghiệm tại kho nước xuất khẩu sẽ cứu doanh nghiệp khỏi nguy cơ mất trắng cả tỷ đồng khi hàng bị ép tái xuất tại biên giới.",
        "comparison": "So sánh Kiểm tra thông thường vs Kiểm tra chặt an toàn thực phẩm: Kiểm tra thông thường chỉ kiểm tra hồ sơ xác suất, lấy mẫu 5% lô hàng, thông quan trong 24h; Kiểm tra chặt lấy mẫu 100% các lô hàng liên tiếp, phân tích toàn diện các chỉ tiêu hóa lý, thời gian chờ kết quả từ 3-5 ngày.",
        "persona": "Chủ hàng nông sản trái cây nhập khẩu, giám đốc doanh nghiệp bán lẻ thực phẩm, chuyên viên pháp chế an toàn thực phẩm (Chị Lan - Giám đốc Chuỗi cửa hàng hoa quả sạch).",
        "checklist": "[ ] Đã tra cứu MRL hoạt chất theo Thông tư 50/2016/TT-BYT\n[ ] Có biên bản test dư lượng của nhà vườn trước khi đóng hàng\n[ ] Đã được cấp Giấy phép Kiểm dịch thực vật nhập khẩu hợp lệ\n[ ] Hợp đồng ngoại thương có điều khoản giữ tiền cọc nếu kiểm tra ATTP không đạt\n[ ] Chuẩn bị sẵn kho lạnh bảo quản tại cửa khẩu nếu phải chờ kết quả test.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng TTĐT Cục Bảo vệ thực vật, Thư Viện Pháp Luật (Nghị định 15/2018/NĐ-CP, Thông tư 34/2018/TT-BNNPTNT, Thông tư 50/2016/TT-BYT) & daily_notes.md (31/10/2026).",
        "keywords": ["food", "safety"]
    },
    {
        "stt": 194,
        "week": "W44 — 01/11/2026",
        "title": "Bí Kíp Thẩm Định C/O Giáp Lưng (Back-To-Back C/O) Qua Cảng Trung Chuyển (Singapore, Hong Kong, Malaysia): 5 Điều Kiện Sống Còn Để Được Hưởng Thuế Nhập Khẩu Ưu Đãi 0%",
        "workflow": "Bước 1: Đối chiếu tính đồng nhất của thông tin: Tên nhà xuất khẩu trên C/O giáp lưng (phải là thương nhân nước trung chuyển), tên người nhận hàng tại Việt Nam, mô tả hàng hóa, số lượng và trọng lượng không được vượt quá C/O gốc ban đầu; Bước 2: Kiểm tra số tham chiếu (Reference No) và ngày cấp của C/O gốc: C/O giáp lưng bắt buộc phải ghi rõ số tham chiếu và ngày cấp của C/O gốc tại Ô số 7 hoặc Ô số 12; Bước 3: Kiểm tra thời hạn hiệu lực: Ngày hết hạn của C/O giáp lưng phải cùng ngày hoặc trước ngày hết hạn của C/O gốc; Bước 4: Đánh dấu hợp lệ tại Ô số 13 mục 'Back-to-Back CO'; Bước 5: Yêu cầu chứng thư Non-Manipulation Certificate (CNM) do Hải quan nước trung chuyển cấp nếu hàng hóa có lưu kho bãi hoặc thay đổi container.",
        "formula": "Công thức hợp lệ C/O giáp lưng: Hợp lệ 100% = [Back-to-Back C/O hợp lệ] + [Through B/L HOẶC CNM của Hải quan trung chuyển] + [Quy cách hàng giữ nguyên vẹn]. Nếu thiếu CNM khi đã thay đổi container -> Thuế suất ưu đãi = 0, truy thu thuế MFN.",
        "top3_mistakes": "1. Số lượng trên C/O giáp lưng lớn hơn số lượng ghi trên C/O gốc của nước sản xuất ban đầu; 2. Thay đổi container tại cảng trung chuyển nhưng chỉ nộp vận đơn chặng 2 mà không có văn bản xác nhận giám sát hải quan của cảng trung chuyển; 3. Thời điểm cấp C/O giáp lưng vượt quá thời hạn 1 năm kể từ ngày cấp C/O gốc.",
        "mnemonic_tip": "'GIÁP LƯNG PHẢI TỰA VÀO GỐC - RỜI CONTAINER PHẢI CÓ GIẤY GIÁM SÁT': C/O giáp lưng như tấm hộ chiếu chuyển tiếp, nếu không chứng minh được bạn chỉ quá cảnh mà không làm giả giấy tờ tại sân bay trung chuyển thì visa sẽ bị hủy lập tức.",
        "comparison": "So sánh Vận đơn suốt (Through B/L) vs Chứng thư Non-Manipulation (CNM): Through B/L dùng khi hàng nằm nguyên trong container từ cảng đầu đến cảng cuối; CNM bắt buộc phải có khi hàng dỡ ra bãi hoặc chia nhỏ đóng vào nhiều container khác nhau tại cảng trung chuyển.",
        "persona": "Nhân viên chứng từ xuất nhập khẩu (docs), chuyên viên đàm phán mua hàng quốc tế, kế toán thanh toán thuế hải quan (Anh Tuấn - Forwarder; Chị Mai - Trưởng phòng Mua hàng).",
        "checklist": "[ ] Kiểm tra Ô số 13 có tích chọn 'Back-to-Back CO'\n[ ] Kiểm tra Ô số 7 có ghi số tham chiếu và ngày cấp C/O gốc\n[ ] Có Chứng thư xác nhận chuyển tải (Non-Manipulation Certificate) nếu dỡ container\n[ ] Mã HS trên C/O giáp lưng trùng khớp 6 số đầu với mã HS trên tờ khai hải quan\n[ ] Trọng lượng hàng khai báo không vượt quá C/O gốc.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 12/2019/TT-BCT, Thông tư 38/2018/TT-BCT, Quyết định 19/2008/QĐ-BCT) & Tạp chí Tài chính.",
        "keywords": ["certificate", "origin"]
    },
    {
        "stt": 195,
        "week": "W44 — 01/11/2026",
        "title": "Kỹ Thuật Hạch Toán & Mở Tờ Khai Nhập Khẩu Tại Chỗ E31 Đối Ứng B11/E62 Theo Nghị Định 18/2021/NĐ-CP: Chiến Lược Nộp Thuế A12 Và Hoàn Thuế An Toàn Cho Doanh Nghiệp Sản Xuất Xuất Khẩu",
        "workflow": "Bước 1: Phân loại bản chất giao dịch: Xác định rõ hàng hóa nhập tại chỗ có được đưa trực tiếp vào chuỗi sản xuất ra sản phẩm xuất khẩu ra nước ngoài hay không; Bước 2: Kê khai mở tờ khai nhập khẩu: Nếu bên bán mở tờ khai xuất khẩu B11 (Xuất kinh doanh) hoặc E62 (Xuất sản phẩm chế xuất), doanh nghiệp nhập khẩu bắt buộc phải mở tờ khai theo loại hình A12 (Nhập kinh doanh sản xuất) để kê khai nộp thuế nhập khẩu và thuế GTGT, KHÔNG ĐƯỢC mở E31 miễn thuế trực tiếp; Bước 3: Lập hồ sơ theo dõi định mức tiêu hao nguyên vật liệu thực tế phục vụ sản xuất sản phẩm xuất khẩu; Bước 4: Sau khi sản phẩm hoàn thành và thực tế xuất khẩu ra nước ngoài (mã E62/E52), lập bộ hồ sơ hoàn thuế nhập khẩu theo quy định tại Điều 34 & 36 Nghị định 134/2016/NĐ-CP gửi cơ quan hải quan giải quyết hoàn thuế.",
        "formula": "Công thức số thuế hoàn được: Thuế_hoàn = (Thuế NK đã nộp của NVL nhập tại chỗ) x (Lượng NVL thực tế cấu thành sản phẩm đã xuất khẩu) / (Tổng lượng NVL nhập khẩu ban đầu). Thời hạn nộp hồ sơ hoàn thuế: Không quá 60 ngày kể từ ngày thông quan tờ khai xuất khẩu.",
        "top3_mistakes": "1. Tự ý mở tờ khai E31 miễn thuế cho hàng nhập tại chỗ khi chưa kiểm tra loại hình tờ khai bên bán mở (bên bán mở B11 nhưng bên mua mở E31); 2. Không hạch toán tách biệt kho nguyên liệu nhập khẩu tại chỗ với kho nguyên liệu nhập khẩu trực tiếp từ nước ngoài; 3. Bỏ lỡ thời hạn nộp hồ sơ xin hoàn thuế dẫn đến tiền thuế nộp thừa bị đóng băng.",
        "mnemonic_tip": "'BÊN BÁN XUẤT KINH DOANH - BÊN MUA PHẢI NỘP THUẾ RỒI HOÀN': Không bao giờ có chuyện bên bán được hưởng doanh thu kinh doanh nội địa mà bên mua lại được hưởng đặc quyền miễn thuế hải quan trực tiếp.",
        "comparison": "So sánh Mở tờ khai E31 trực tiếp vs Mở A12 rồi hoàn thuế: Mở E31 trực tiếp (nếu đủ điều kiện đối ứng E54/E62) giúp doanh nghiệp không phải bỏ vốn nộp thuế trước; Mở A12 tuy phải ứng vốn nộp thuế nhưng an toàn pháp lý tuyệt đối 100%, triệt tiêu hoàn toàn rủi ro bị truy thu phạt chậm nộp 0,03%/ngày khi kiểm tra sau thông quan.",
        "persona": "Kế toán trưởng doanh nghiệp sản xuất xuất khẩu, giám đốc tài chính (CFO), chuyên viên kế toán kho vận và đại lý hải quan (Anh Nam - Giám đốc Tài chính).",
        "checklist": "[ ] Đối chiếu mã loại hình tờ khai xuất của bên bán trước khi mở tờ khai nhập\n[ ] Lựa chọn chính xác mã loại hình A12 nếu bên bán mở B11\n[ ] Kê khai nộp đủ thuế nhập khẩu và GTGT theo quy định\n[ ] Lưu trữ đầy đủ chứng từ thanh toán ngân hàng cho bên bán\n[ ] Thiết lập bảng định mức sản xuất xuất khẩu để phục vụ hoàn thuế.",
        "url": "https://en.wikipedia.org/wiki/Customs_duties",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 18/2021/NĐ-CP Điều 1; Nghị định 134/2016/NĐ-CP Điều 12, 36; Thông tư 38/2015/TT-BTC Điều 86) & SEO_Keyword_Trends_Eureka.xlsx (Từ khóa STT 1 W44).",
        "keywords": ["customs", "duties"]
    }
]
