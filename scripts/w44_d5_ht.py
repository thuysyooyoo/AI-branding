# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W44 DAY 5 — 02/11/2026)
Số lượng: 5 ý tưởng (STT 196 - 200)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W44_D5 = [
    {
        "stt": 196,
        "week": "W44 — 02/11/2026",
        "title": "Hướng Dẫn Kỹ Thuật Đọc Và Soát 7 Chỉ Tiêu Sống Còn Trên Vận Đơn Switch B/L Cho Giao Dịch 3 Bên: Cách Bảo Mật 100% Thông Tin Xưởng Gốc & Giá Mua Đầu Vào Không Để Khách Hàng Cuối Phát Hiện",
        "workflow": "Quy trình 5 bước thực hiện Switch B/L an toàn tuyệt đối: Bước 1: Ký hợp đồng dịch vụ vận chuyển chỉ định rõ đại lý Forwarder uy tín có văn phòng tại cảng trung chuyển (Singapore, Hong Kong); Bước 2: Nhận bộ B/L 1 từ xưởng sản xuất gốc (Shipper: Xưởng gốc, Consignee: Công ty Trading); Bước 3: Công ty Trading nộp lại và Surrendered toàn bộ 3 bản B/L gốc thứ nhất cho văn phòng Hãng tàu/Forwarder; Bước 4: Yêu cầu phát hành bộ B/L 2 (Switch B/L) với Shipper đổi thành Công ty Trading, Consignee đổi thành Người mua cuối cùng tại nước nhập khẩu; Bước 5: Kiểm tra khớp tuyệt đối 7 chỉ tiêu sống còn: Tên hàng, Số lượng kiện, Trọng lượng Gross/Net, Thể tích CBM, Số cont/Số chì, Cảng bốc/Cảng dỡ, và Ngày Onboard B/L 2 phải trùng khớp 100% với ngày Onboard của B/L 1.",
        "formula": "Nguyên tắc '3 Đổi - 4 Khóa': Đổi Shipper, Đổi Consignee, Đổi Notify Party — Khóa Ngày Onboard, Khóa Cảng Đi/Đến, Khóa Trọng Lượng, Khóa Số Container/Seal.",
        "top3_mistakes": "1. Để ngày xếp hàng lên tàu (Shipped on Board Date) trên B/L 2 khác với B/L 1, dẫn đến bị ngân hàng mở L/C từ chối thanh toán vì lỗi Discrepancy; 2. Để sót thông tin tên xưởng sản xuất gốc trong phần mô tả hàng hóa (Description of Goods) hoặc trên ký mã hiệu bao bì (Shipping Marks); 3. Chưa thu hồi triệt để bộ 3 bản B/L gốc thứ nhất mà đã cho phát hành B/L thứ hai, tạo ra rủi ro pháp lý 2 bộ vận đơn sở hữu cùng 1 lô hàng ngoài thị trường.",
        "mnemonic_tip": "'B/L Một thu trọn về tay - B/L Hai mới được phép bay sang người'.",
        "comparison": "B/L Thông thường (Direct B/L): Thể hiện trực tiếp người bán và người mua thực tế, không bảo mật được giá gốc và thông tin nhà máy vs Switch B/L: Nghiệp vụ nghiệp vụ hàng hải trung gian cho phép che giấu xưởng sản xuất và bảo toàn biên lợi nhuận của công ty thương mại.",
        "persona": "Chuyên viên Chứng từ Xuất nhập khẩu (Docs/CS Logistics), Giám đốc Kinh doanh Công ty Thương mại Quốc tế, Trưởng phòng Mua hàng Quốc tế.",
        "checklist": "Checklist 6 điểm soát trước khi phát hành Switch B/L: 1. Đã có xác nhận bằng văn bản thu hồi hủy 3/3 B/L gốc 1 chưa?; 2. Ngày Onboard trên B/L 2 có trùng từng ngày với B/L 1 không?; 3. Shipping Marks có bị lộ tên xưởng gốc không?; 4. Commercial Invoice và Packing List mới đã thay thế giá bán cho Consignee cuối chưa?; 5. C/O giáp lưng (nếu có) đã cấp đúng theo B/L 2 chưa?; 6. Hãng tàu đã nhận đủ phí Switch B/L và cam kết không liên hệ trực tiếp Consignee chưa?",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cẩm nang Incoterms 2020 & UCP 600, Thư Viện Pháp Luật (Bộ luật Hàng hải Việt Nam Điều 86, 87) & daily_notes.md (02/11/2026).",
        "keywords": ["bill", "lading"]
    },
    {
        "stt": 197,
        "week": "W44 — 02/11/2026",
        "title": "Quy Trình 5 Bước Xử Lý Tình Huống Tờ Khai Hải Quan Bị Báo Lỗi 'Chưa Đối Chiếu Cơ Sở Dữ Liệu Quốc Gia (CSDLQG)': Hướng Dẫn Đồng Bộ VNeID Mức 2 Và Chữ Ký Số Khai Báo Hải Quan Nhanh Nhất",
        "workflow": "Quy trình 5 bước gỡ lỗi đồng bộ CSDLQG: Bước 1: Kiểm tra thông tin CCCD gắn chip của Người đại diện theo pháp luật trên ứng dụng VNeID (đảm bảo đã định danh Mức 2 và không có thông báo sai lệch cư trú); Bước 2: Đăng nhập Cổng Dịch vụ công Quốc gia (dichvucong.gov.vn) tra cứu mã số thuế doanh nghiệp xem trạng thái liên thông với Bộ Công an đã đạt tích xanh chưa; Bước 3: Liên hệ nhà cung cấp Chữ ký số (Viettel-CA, VNPT-CA, FPT-CA...) yêu cầu cập nhật lại chứng thư số số hóa theo số CCCD mới 12 số (thay vì CMND 9 số cũ); Bước 4: Đăng nhập Cổng Thông tin điện tử Tổng cục Hải quan (customs.gov.vn), vào mục 'Đăng ký Doanh nghiệp sử dụng Chữ ký số', cập nhật lại Serial Token và gửi yêu cầu phê duyệt tự động; Bước 5: Mở phần mềm ECUS5, vào Hệ thống -> Đăng ký thông tin tài khoản VNACCS, chọn đồng bộ thông tin người khai hải quan và đợi 15-30 phút để hệ thống cập nhật trước khi truyền tờ khai mới.",
        "formula": "Quy tắc '3 Sạch': Thông tin VNeID cá nhân sạch - Dữ liệu Thuế doanh nghiệp sạch - Chữ ký số hải quan đồng bộ sạch.",
        "top3_mistakes": "1. Vội vàng truyền đi truyền lại tờ khai nhiều lần khi bị báo lỗi, khiến hệ thống VNACCS đưa tài khoản vào danh sách nghi ngờ spam và khóa tạm thời 24h; 2. Thay đổi giám đốc trên giấy phép ĐKKD nhưng dùng chữ ký số cũ mang thông tin giám đốc tiền nhiệm để truyền tờ khai; 3. Quên không làm thủ tục cập nhật thông tin chữ ký số trên hệ thống VNACCS mà chỉ cắm Token mới vào máy tính rồi truyền luôn.",
        "mnemonic_tip": "'Dân cư sạch trước - Thuế thông sau - Chữ ký số đồng bộ thì tàu mới đi'.",
        "comparison": "Xử lý thủ công truyền thống: Kế toán phải chạy lên Chi cục Thuế và Chi cục Hải quan xin công văn xác nhận mất 3-5 ngày vs Xử lý số hóa chuẩn mực: Tự động đối soát và kích hoạt trực tuyến trên VNeID và VNSW trong vòng 2 giờ.",
        "persona": "Kế toán thuế doanh nghiệp XNK, Nhân viên khai báo Hải quan hiện trường, Quản trị viên hệ thống công nghệ thông tin công ty Logistics.",
        "checklist": "Checklist 5 bước gỡ nghẽn: 1. Đã kiểm tra VNeID Mức 2 của Giám đốc chưa?; 2. Token đã cập nhật số CCCD 12 số chưa?; 3. Đã cập nhật Serial trên Cổng thông tin Tổng cục Hải quan chưa?; 4. Đã chạy đồng bộ lại tài khoản trên ECUS5 chưa?; 5. Có tờ khai gấp cần xin văn bản can thiệp thủ công từ Đội thủ tục Hải quan không?",
        "url": "https://en.wikipedia.org/wiki/Authentication",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng Dịch vụ công Quốc gia, Thư Viện Pháp Luật (Quyết định 06/QĐ-TTg, Nghị định 59/2022/NĐ-CP) & daily_notes.md (02/11/2026).",
        "keywords": ["authentication", "identity"]
    },
    {
        "stt": 198,
        "week": "W44 — 02/11/2026",
        "title": "Cẩm Nang Rà Soát Tên Hàng & Mã HS Code Giữa Hóa Đơn Thương Mại (Commercial Invoice) Và Hóa Đơn Giá Trị Gia Tăng (VAT): Bí Quyết Tránh Bẫy Bị Bác Khấu Trừ Thuế GTGT Khi Kiểm Tra Sau Thông Quan",
        "workflow": "Quy trình 4 bước rà soát chuẩn hóa tên hàng XNK: Bước 1: Lập Bảng đối chiếu chéo (Cross-check Matrix) gồm 5 cột: Tên hàng trên Invoice ngoại (Tiếng Anh), Tên hàng dịch nghĩa tiếng Việt trên Tờ khai Hải quan, Mã HS Code khai báo, Tên hàng trên Hóa đơn GTGT đầu ra/đầu vào và Tên mã vật tư trên phần mềm ERP; Bước 2: Kiểm tra tính thống nhất của đơn vị tính (Unit): Đảm bảo đơn vị tính trên hóa đơn thuế (cái, bộ, kg, mét) phải khớp hoặc có công thức quy đổi chính xác với đơn vị tính chuẩn của biểu thuế hải quan; Bước 3: Rà soát mô tả hàng hóa theo Nghị định 123/2020/NĐ-CP: Tuyệt đối không chấp nhận các hóa đơn chỉ ghi chung chung 'vật tư phụ tùng', bắt buộc phải có tên model, nhãn hiệu và quy cách kỹ thuật; Bước 4: Lưu trữ bộ chứng từ hoàn chỉnh kèm tài liệu kỹ thuật (Catalogue, TDS) trong cùng 1 folder điện tử theo số tờ khai để phục vụ giải trình thanh tra.",
        "formula": "Nguyên tắc '3 Khớp': Khớp Tên Hàng - Khớp Mã HS - Khớp Định Mức Kỹ Thuật.",
        "top3_mistakes": "1. Hóa đơn đầu vào ghi tên tiếng Việt theo cách gọi dân dã của xưởng nhưng tờ khai hải quan ghi theo thuật ngữ chuyên ngành dẫn đến giải trình lệch nhau; 2. Để nhà cung cấp xuất hóa đơn gộp nhiều linh kiện vào một dòng duy nhất để giảm độ dài hóa đơn; 3. Không lưu lại tài liệu kỹ thuật chứng minh sự liên quan giữa vật tư mua vào và sản phẩm xuất khẩu hoàn thành.",
        "mnemonic_tip": "'Tên trên hóa đơn - Khớp tên tờ khai - Rõ từng model - Thuế soi chẳng ngại'.",
        "comparison": "Cách làm lỏng lẻo: Mua hàng chỉ cần nhận hóa đơn có tổng tiền và MST đúng -> Hậu quả bị thanh tra truy thu và phạt 20% vs Cách làm chuẩn mực quản trị Eureka: Kiểm soát từng chỉ tiêu hàng hóa ngay tại thời điểm ký duyệt đề nghị thanh toán.",
        "persona": "Kế toán trưởng doanh nghiệp sản xuất xuất khẩu, Giám đốc Tài chính (CFO), Chuyên viên tư vấn kiểm toán thuế.",
        "checklist": "Checklist 6 điểm kiểm soát hóa đơn thuế XNK: 1. Tên hàng có ghi rõ model và nhãn hiệu không?; 2. Đơn vị tính có khớp với tờ khai hải quan không?; 3. Số lượng có khớp với phiếu nhập kho thực tế không?; 4. Thuế suất GTGT áp dụng có đúng quy định giảm thuế hiện hành không?; 5. Có bảng kê chi tiết đính kèm cho các lô vật tư nhiều món không?; 6. Chứng từ thanh toán ngân hàng ủy nhiệm chi có ghi rõ số hóa đơn thanh toán không?",
        "url": "https://en.wikipedia.org/wiki/Harmonized_System",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thông tin Tổng cục Thuế, Thư Viện Pháp Luật (Nghị định 123/2020/NĐ-CP, Thông tư 38/2015/TT-BTC Điều 16) & daily_notes.md (02/11/2026).",
        "keywords": ["harmonized", "system"]
    },
    {
        "stt": 199,
        "week": "W44 — 02/11/2026",
        "title": "Hướng Dẫn Chi Tiết Thủ Tục Nhập Khẩu Máy Móc Thiết Bị Đã Qua Sử Dụng Theo Quyết Định 18/2019/QĐ-TTg: Cách Tra Cứu Tuổi Thiết Bị, Tiêu Chuẩn An Toàn & Chuẩn Bị Hồ Sơ Giám Định Không Bị Giữ Tại Cảng",
        "workflow": "Quy trình 5 bước nhập khẩu máy móc cũ an toàn: Bước 1: Tính toán tuổi thiết bị: Tuổi thiết bị (năm) = Năm nhập khẩu - Năm sản xuất. Đảm bảo tuổi thiết bị không quá 10 năm (đối với một số ngành đặc thù như cơ khí, dệt may, in ấn có thể lên đến 15-20 năm theo phụ lục); Bước 2: Yêu cầu người bán nước ngoài cung cấp đầy đủ: Bản chụp tem máy (Nameplate) nguyên bản, số Serial/Chassis, Catalogue gốc và Giấy xác nhận năm sản xuất của nhà chế tạo; Bước 3: Đăng ký giám định tình trạng máy móc cũ tại một trong các tổ chức giám định được Bộ Khoa học & Công nghệ chỉ định (Quatest, Vinacontrol, SGS...) trước hoặc ngay khi hàng về cảng; Bước 4: Mở tờ khai hải quan kèm văn bản đăng ký giám định và xin đưa hàng về kho bảo quản theo Điều 32 Thông tư 38/2015/TT-BTC; Bước 5: Mời giám định viên đến kho nhà máy tiến hành kiểm tra thực tế, lấy chứng thư giám định kết luận 'Thiết bị sản xuất theo tiêu chuẩn phù hợp quy chuẩn kỹ thuật quốc gia (QCVN) và tuổi dưới 10 năm' nộp bổ sung cho Hải quan để thông quan chính thức.",
        "formula": "Công thức kiểm soát tuổi máy: Tuổi máy = (Năm Tờ khai Hải quan - Năm Sản xuất ghi trên Nameplate) <= 10 năm.",
        "top3_mistakes": "1. Mua máy móc cũ không có tem mác gốc hoặc tem mác bị mờ rỉ sét không đọc được năm sản xuất, dẫn đến bị Hải quan từ chối cho đưa về kho bảo quản; 2. Tin vào năm sản xuất ghi trên hóa đơn của đại lý bán máy cũ mà không kiểm tra số khung thực tế của nhà sản xuất; 3. Tự ý tháo dỡ niêm phong hải quan hoặc đưa máy móc vào vận hành sản xuất trước khi có chứng thư giám định đạt yêu cầu và được thông quan chính thức.",
        "mnemonic_tip": "'Dưới mười năm tuổi mới mua - Tem nhãn nguyên bản chẳng thừa một ly'.",
        "comparison": "Máy móc mới 100%: Chỉ cần kiểm tra hồ sơ thông thường và công bố hợp quy (nếu thuộc nhóm 2) vs Máy móc đã qua sử dụng: Bắt buộc giám định niên hạn, giám định hiệu suất năng lượng và tiêu chuẩn an toàn môi trường nghiêm ngặt.",
        "persona": "Trưởng phòng Kỹ thuật Cơ điện nhà máy, Trưởng phòng Xuất nhập khẩu doanh nghiệp sản xuất, Giám đốc điều hành doanh nghiệp cơ khí chế tạo.",
        "checklist": "Checklist 6 tài liệu bắt buộc trong bộ hồ sơ máy cũ: 1. Hợp đồng mua bán và Commercial Invoice ghi rõ model, serial?; 2. Bản chụp nameplate rõ nét năm sản xuất?; 3. Giấy chứng nhận của nhà sản xuất về năm chế tạo (nếu máy không ghi năm)?; 4. Đơn đăng ký giám định được tổ chức giám định tiếp nhận?; 5. Văn bản đề nghị đưa hàng về kho bảo quản mẫu 09/BQHH?; 6. Giấy phép kinh doanh ngành nghề phù hợp với thiết bị nhập khẩu?",
        "url": "https://en.wikipedia.org/wiki/Depreciation",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng TTĐT Bộ KH&CN, Thư Viện Pháp Luật (Quyết định 18/2019/QĐ-TTg, Thông tư 23/2015/TT-BKHCN) & daily_notes.md (02/11/2026).",
        "keywords": ["depreciation", "asset"]
    },
    {
        "stt": 200,
        "week": "W44 — 02/11/2026",
        "title": "6 Bước Tự Kiểm Tra Và Khắc Phục Lỗi Hệ Thống VNACCS Chậm Gạch Nợ Thuế Hải Quan 24/7: Cách Xuất Trình Giấy Nộp Tiền Có Xác Nhận Ngân Hàng Để Đề Nghị Giải Phóng Hàng Khẩn Cấp",
        "workflow": "Quy trình 6 bước xử lý sự cố chậm gạch nợ thuế: Bước 1: Đăng nhập Cổng thông tin Tổng cục Hải quan tra cứu tình trạng nợ thuế của tờ khai (kiểm tra xem trạng thái nợ thuế là 'Chưa nộp', 'Đã truyền thông tin' hay 'Chờ xác nhận Kho bạc'); Bước 2: Kiểm tra lại lệnh nộp tiền trên hệ thống Internet Banking của ngân hàng: Đảm bảo khớp tuyệt đối 4 thông tin: Mã số thuế, Số tờ khai, Mã Chi cục Hải quan mở tờ khai và Tiểu mục nộp thuế (Ví dụ: 1901 cho thuế nhập khẩu, 1702 cho thuế GTGT hàng nhập khẩu); Bước 3: Nếu sau 60 phút hệ thống hải quan chưa gạch nợ tự động do nghẽn cổng truyền tin, yêu cầu ngân hàng in Giấy nộp tiền vào ngân sách nhà nước có đóng dấu đỏ tròn và chữ ký xác nhận của Giao dịch viên/Kiểm soát viên; Bước 4: Soạn Công văn đề nghị hỗ trợ thông quan giải phóng hàng khẩn cấp do lỗi truyền nhận dữ liệu điện tử theo mẫu chuẩn; Bước 5: Nộp trực tiếp Giấy nộp tiền dấu đỏ và công văn cho bộ phận Quản lý thuế của Chi cục Hải quan nơi mở tờ khai để cán bộ thuế kiểm tra và nhập gạch nợ thủ công trên hệ thống KTTT (Kế toán thuế tập trung); Bước 6: Nhận thông báo chấp nhận thông quan trên phần mềm ECUS5 và tiến hành in mã vạch lấy hàng ra khỏi cảng.",
        "formula": "Công thức cấp cứu thông quan: Giấy nộp tiền dấu đỏ Ngân hàng + Công văn giải trình + Đề nghị can thiệp thủ công KTTT = Giải phóng hàng trong 30 phút.",
        "top3_mistakes": "1. Điền sai mã Chi cục Hải quan hoặc sai số tờ khai trên lệnh nộp thuế điện tử khiến tiền chuyển vào tài khoản treo của Kho bạc khác; 2. Nộp tiền sát giờ hành chính thứ Sáu dẫn đến tiền bị giữ tại ngân hàng chuyển lệnh và phải đợi sang sáng thứ Hai tuần sau mới nổi nợ; 3. Không lưu lại mã giao dịch tham chiếu (FT Code) của ngân hàng thương mại để tra cứu phối hợp thu.",
        "mnemonic_tip": "'Bốn đúng nộp thuế chẳng sai - Cổng nghẽn xin dấu gạch ngay tại quầy'.",
        "comparison": "Nộp thuế qua cổng thanh toán điện tử 24/7 tự động: Gạch nợ tức thì trong vòng 5-10 phút vs Nộp thuế qua ngân hàng chưa liên kết trực tiếp: Mất từ 2-4 tiếng hoặc chuyển sang ngày hôm sau nếu có sự cố gián đoạn đường truyền Kho bạc.",
        "persona": "Nhân viên làm thủ tục hải quan hiện trường, Kế toán thanh toán xuất nhập khẩu, Trưởng phòng Logistics.",
        "checklist": "Checklist 5 điểm đối soát nộp thuế hải quan 24/7: 1. Mã Chi cục Hải quan đã đúng mã 4 ký tự chưa (Ví dụ: 02CI, 01NV)?; 2. Số tờ khai đã gõ đúng 12 chữ số chưa?; 3. Tiểu mục nộp thuế đã tách riêng thuế XNK và thuế GTGT chưa?; 4. Số dư tài khoản thanh toán tại ngân hàng có đủ trừ phí giao dịch không?; 5. Đã in sẵn Giấy nộp tiền dấu đỏ của ngân hàng phòng trường hợp nghẽn mạng chưa?",
        "url": "https://en.wikipedia.org/wiki/Electronic_funds_transfer",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Cổng thanh toán điện tử Hải quan 24/7, Thư Viện Pháp Luật (Thông tư 184/2015/TT-BTC, Thông tư 38/2015/TT-BTC Điều 34) & daily_notes.md (02/11/2026).",
        "keywords": ["funds", "transfer"]
    }
]
