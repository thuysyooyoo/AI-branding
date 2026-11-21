# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W43 DAY 3 — 23/10/2026)
Số lượng: 5 ý tưởng (STT 171 - 175)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W43_D3 = [
    {
        "stt": 171,
        "week": "W43 — 23/10/2026",
        "title": "Cẩm Nang 4 Bước Xử Lý Khi Phát Hiện Khai Sai Mã HS Code Hàng Nhập Kho Ngoại Quan (KNQ): Tránh Bị Phạt Gian Lận & Phong Tỏa Lô Hàng Tại Cảng",
        "workflow": "Bước 1: Rà soát và thu thập toàn bộ tài liệu kỹ thuật gốc (Datasheet, Catalog, Bản vẽ phân tích thành phần của nhà sản xuất) và đối chiếu với 6 quy tắc phân loại mã HS để xác định chính xác mã HS đúng; Bước 2: Lập văn bản Đề nghị sửa đổi, bổ sung thông tin tờ khai hải quan (Mẫu số 01/KBS/GSQL ban hành kèm Thông tư 39/2018/TT-BTC) giải trình rõ lý do nhầm lẫn khách quan; Bước 3: Nộp hồ sơ xin sửa đổi trên Hệ thống VNACCS TRƯỚC KHI công chức hải quan tiếp nhận kiểm tra hồ sơ hoặc kiểm tra thực tế hàng hóa, nộp kèm bản sao chứng từ giải trình; Bước 4: Làm việc trực tiếp với Chi cục Hải quan quản lý kho ngoại quan để phối hợp kiểm tra đối chiếu thực tế (nếu thuộc diện luồng Vàng/Đỏ) và nhận kết quả phê duyệt cho phép đưa hàng vào kho.",
        "formula": "Thời hạn khai sửa tờ khai an toàn = Thời điểm nộp Mẫu 01/KBS/GSQL < Thời điểm Hải quan thông báo kiểm tra thực tế hàng hóa hoặc cho phép đưa hàng vào kho ngoại quan (Điều 20 Thông tư 38/2015/TT-BTC sửa đổi tại Thông tư 39/2018/TT-BTC).",
        "top3_mistakes": "1. Nghĩ rằng hàng vào kho ngoại quan chưa chịu thuế nhập khẩu nên khai mã HS thế nào cũng được; 2. Đợi đến khi hải quan phát hiện ra lỗi sai và ra quyết định kiểm hóa mới làm đơn xin sửa chữa; 3. Không lưu giữ chứng từ kỹ thuật gốc dẫn đến khi bị yêu cầu phân tích giám định thì mất hàng tuần chờ đợi đối tác nước ngoài gửi bổ sung.",
        "mnemonic_tip": "'HS Sai Phải Sửa Liền Tay - Đợi Khi Kiểm Hóa Đắng Cay Mất Tiền' (Tờ khai kho ngoại quan vẫn là hồ sơ pháp lý, phát hiện sai phải chủ động nộp đơn KBS ngay trước khi bị kiểm tra).",
        "comparison": "Chủ động sửa trước khi kiểm tra (Chỉ nộp Mẫu 01/KBS, không bị xử phạt vi phạm hành chính, thông quan trong 24h) VS Bị động để hải quan phát hiện (Bị lập biên bản xử phạt 1-5 triệu, chuyển luồng Đỏ kiểm hóa 100%, hàng ách tắc 7-15 ngày tại bãi cảng).",
        "persona": "Anh Minh - Chủ doanh nghiệp phân phối linh kiện điện tử và thiết bị công nghiệp; Chuyên viên chứng từ XNK tại các công ty logistics và chủ kho ngoại quan.",
        "checklist": "[ ] 1. Có Datasheet và hình ảnh thực tế hàng hóa chưa? | [ ] 2. Đã xác định mã HS đúng theo 6 quy tắc HS chưa? | [ ] 3. Tờ khai đã bị thông báo kiểm tra thực tế chưa? | [ ] 4. Đã hoàn thiện văn bản Mẫu 01/KBS/GSQL ký đóng dấu chưa? | [ ] 5. Đã gửi hồ sơ điện tử qua hệ thống VNACCS chưa?",
        "url": "https://en.wikipedia.org/wiki/Harmonized_System",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 38/2015/TT-BTC Điều 20 & 91, Thông tư 39/2018/TT-BTC) & daily_notes.md (22/10/2026).",
        "keywords": ["harmonized", "system"]
    },
    {
        "stt": 172,
        "week": "W43 — 23/10/2026",
        "title": "Lộ Trình 3 Bước Chuyển Đổi Địa Chỉ Hộ Kinh Doanh Khác Quận/Tỉnh Chuẩn Luật: Bảo Toàn 100% Mã Số Thuế & Tránh Bị Khóa Gian Hàng Sàn TMĐT",
        "workflow": "Bước 1: Nộp hồ sơ thay đổi thông tin đăng ký thuế (Tờ khai Mẫu 08-MST ban hành kèm Thông tư 105/2020/TT-BTC) tại Chi cục Thuế nơi chuyển đi để thực hiện đối soát và hoàn thành các nghĩa vụ thuế môn bài, thuế GTGT, TNCN còn nợ; Bước 2: Cơ quan thuế nơi đi lập và ban hành Thông báo người nộp thuế chuyển địa điểm (Mẫu số 09-MST) trong thời hạn 03 ngày làm việc; Bước 3: Nộp hồ sơ thay đổi địa chỉ trụ sở hộ kinh doanh tại Phòng Tài chính - Kế hoạch (UBND quận mới) và làm thủ tục kích hoạt tiếp nhận nghĩa vụ thuế tại Chi cục Thuế nơi chuyển đến mà không làm thay đổi mã số thuế định danh.",
        "formula": "Thời hạn hoàn tất thủ tục chuyển địa điểm = Tối đa 03 ngày làm việc tại cơ quan thuế nơi đi (nhận Mẫu 09-MST) + 03 ngày làm việc tại UBND quận mới (cấp đổi Giấy chứng nhận ĐKKD theo Nghị định 01/2021/NĐ-CP).",
        "top3_mistakes": "1. Nghe lời dịch vụ kế toán 'đóng hộ cũ ở quận này để sang quận khác mở hộ mới' làm mất mã số thuế và bị sàn TMĐT xóa gian hàng; 2. Không kiểm tra trước các khoản nợ thuế lũy kế hoặc thuế đất phi nông nghiệp gán nhầm trên eTax Mobile dẫn đến bị ngâm hồ sơ chốt nợ; 3. Tự ý thay đổi biển hiệu và kinh doanh tại địa chỉ mới khi chưa hoàn tất thủ tục thuế dẫn đến bị xử phạt trốn thuế từ 2 đến 5 triệu đồng.",
        "mnemonic_tip": "'Giữ Mã Giữ Nghiệp - Đổi Chỗ Giữ MST' (Mã số thuế hộ kinh doanh gắn liền với định danh cá nhân người đại diện, chuyển đi đâu cũng chỉ dùng Mẫu 08-MST chuyển quận, tuyệt đối không bao giờ được xóa sổ đóng mã).",
        "comparison": "Chuyển địa điểm theo Mẫu 08-MST (Giữ nguyên MST, giữ nguyên lịch sử bán hàng TMĐT, không phải quyết toán đóng mã) VS Đóng hộ cũ mở hộ mới (Mất sạch shop online, mất thương hiệu, dính thủ tục giải thể phức tạp, bị rà soát nợ thuế toàn diện).",
        "persona": "Chị Hương - Chủ shop kinh doanh thời trang và đồ gia dụng trên Shopee/TikTok Shop; Các chủ hộ kinh doanh cá thể đang tìm mặt bằng mới.",
        "checklist": "[ ] 1. Đã tra cứu tình trạng nợ thuế cá nhân trên eTax Mobile chưa? | [ ] 2. Đã kê khai và nộp đủ thuế quý gần nhất chưa? | [ ] 3. Đã nộp Mẫu 08-MST tại Chi cục Thuế cũ chưa? | [ ] 4. Đã nhận Thông báo Mẫu 09-MST chưa? | [ ] 5. Đã nộp hồ sơ đổi giấy phép kinh doanh tại UBND quận mới chưa?",
        "url": "https://en.wikipedia.org/wiki/Tax",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 105/2020/TT-BTC Điều 10, Thông tư 80/2021/TT-BTC, Nghị định 01/2021/NĐ-CP) & daily_notes.md (22/10/2026).",
        "keywords": ["tax"]
    },
    {
        "stt": 173,
        "week": "W43 — 23/10/2026",
        "title": "Hướng Dẫn 3 Bước Khấu Trừ & Kê Khai Thuế TNCN Tạm Nộp Theo Quý Chuẩn Xác: Cách Tính Lũy Tiến Từng Tháng Tránh Bị Truy Thu Phạt Chậm Nộp",
        "workflow": "Bước 1: Tính toán số thuế TNCN phải khấu trừ theo Biểu thuế lũy tiến từng phần riêng biệt cho từng tháng trong quý (Tháng nào người lao động có thu nhập chịu thuế vượt mức giảm trừ gia cảnh 15.5 triệu đồng thì trích khấu trừ ngay trong tháng đó, dù doanh nghiệp kê khai theo quý); Bước 2: Lập Bảng kê chi tiết khấu trừ thuế TNCN hàng tháng của toàn bộ nhân viên và tổng hợp số thuế khấu trừ của cả 3 tháng lên Tờ khai thuế TNCN Mẫu số 05/KK-TNCN; Bước 3: Nộp Tờ khai và tiền thuế TNCN đã khấu trừ vào ngân sách nhà nước chậm nhất vào ngày cuối cùng của tháng đầu tiên của quý tiếp theo (Ví dụ quý 3 hạn nộp là 31/10).",
        "formula": "Thuế TNCN tạm nộp Quý = Thuế TNCN khấu trừ Tháng 1 + Thuế TNCN khấu trừ Tháng 2 + Thuế TNCN khấu trừ Tháng 3 (Trong đó Thuế từng tháng = Thu nhập tính thuế từng tháng x Thuế suất lũy tiến, KHÔNG ĐƯỢC lấy tổng thu nhập quý chia 3).",
        "top3_mistakes": "1. Lấy tổng thu nhập 3 tháng chia bình quân để so sánh với mức giảm trừ gia cảnh 15.5 triệu đồng; 2. Tháng thu nhập vượt ngưỡng không khấu trừ tiền của người lao động mà đợi đến cuối quý mới tính bù, khi người lao động đã nghỉ việc không đòi lại được tiền; 3. Không phân biệt giữa nghĩa vụ 'khấu trừ tạm nộp tại nguồn từng tháng' với thủ tục 'quyết toán năm'.",
        "mnemonic_tip": "'Khai Theo Quý - Tính Theo Tháng' (Nhớ nguyên tắc vàng: Tờ khai nộp theo quý nhưng dòng tiền tính thuế phải chặt chẽ theo từng tháng phát sinh trả lương).",
        "comparison": "Khấu trừ đúng từng tháng (Đúng quy định Điều 25 Thông tư 111, an toàn tuyệt đối khi thanh tra, không bao giờ bị truy thu) VS Gộp quý chia bình quân (Kế toán tự tiện vi phạm luật, doanh nghiệp bị truy thu toàn bộ số thuế thiếu kèm phạt 20% và phạt chậm nộp 0.03%/ngày).",
        "persona": "Anh Dũng - Giám đốc doanh nghiệp vận tải & kho vận; Kế toán trưởng, chuyên viên tiền lương và nhân sự (C&B) tại các doanh nghiệp sản xuất, xuất nhập khẩu.",
        "checklist": "[ ] 1. Bảng lương có tách bạch thu nhập từng tháng 7, 8, 9 riêng biệt không? | [ ] 2. Thuế TNCN từng tháng đã tính theo biểu lũy tiến riêng chưa? | [ ] 3. Tổng số thuế trên 3 tháng có khớp với Tờ khai Mẫu 05/KK-TNCN không? | [ ] 4. Đã lập chứng từ khấu trừ thuế điện tử cho lao động thời vụ chưa? | [ ] 5. Đã nộp tiền thuế trước hạn chót ngày cuối tháng đầu quý sau chưa?",
        "url": "https://en.wikipedia.org/wiki/Tax_withholding",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 111/2013/TT-BTC Điều 25, Nghị định 126/2020/NĐ-CP Điều 8, Luật Thuế TNCN) & daily_notes.md (22/10/2026).",
        "keywords": ["withholding", "tax"]
    },
    {
        "stt": 174,
        "week": "W43 — 23/10/2026",
        "title": "Quy Trình 5 Bước Xin Cấp C/O Thay Thế (Replacement C/O) Khi Hàng Đã Cập Cảng Nước Ngoài: Giải Cứu Ưu Đãi Thuế FTA Trong 48 Giờ",
        "workflow": "Bước 1: Nhận văn bản từ chối hoặc thông báo nghi vấn sai lệch C/O từ Hải quan nước nhập khẩu, đối chiếu và xác định cụ thể thông tin bị sai (tên nhà máy, mã số duyệt, số vận đơn B/L); Bước 2: Thu hồi bản chính C/O bị lỗi (nếu là bản giấy) hoặc làm đơn hủy C/O điện tử trên hệ thống EcoSys/VCCI kèm công văn giải trình nguyên nhân nhầm lẫn; Bước 3: Lập hồ sơ xin cấp C/O thay thế, tích chọn ô 'Issued in substitution for C/O No... Dated...' ghi rõ số tham chiếu và ngày cấp của C/O cũ; Bước 4: Cơ quan cấp C/O thụ lý, đối soát và cấp bộ C/O thay thế mới có ngày cấp thực tế mới nhưng ghi nhận hiệu lực hồi tố theo C/O cũ; Bước 5: Truyền dữ liệu số hoặc chuyển phát hỏa tốc bản gốc cho đối tác nước ngoài nộp lại cho hải quan cảng đến để giải phóng hàng và hoàn trả tiền ký quỹ.",
        "formula": "Thời hạn hợp lệ xin cấp C/O thay thế = Trong vòng 01 năm kể từ ngày cấp C/O ban đầu và trước thời điểm cơ quan hải quan nước nhập khẩu ra quyết định bác bỏ ưu đãi thuế vĩnh viễn (Theo các Hiệp định RCEP, VKFTA, EVFTA).",
        "top3_mistakes": "1. Tự ý gạch xóa, sửa chữa viết đè lên bản chính C/O cũ (Hành vi này làm C/O mất hiệu lực vĩnh viễn và bị nghi ngờ làm giả chứng từ); 2. Không nộp lại bản gốc C/O cũ dẫn đến việc cơ quan chức năng từ chối cấp C/O thay thế vì vi phạm nguyên tắc 'chỉ tồn tại một C/O duy nhất cho một lô hàng'; 3. Quên không ghi chú dòng chữ 'Issued in substitution' trên form mới.",
        "mnemonic_tip": "'C/O Lỗi Không Gạch Xóa - Xin Cấp Thay Thế Cứu Ngay Khoản Thuế' (Tuyệt đối không sửa tay, phải đi đúng quy trình cấp C/O thay thế có hồi tố để bảo vệ quyền lợi thuế FTA).",
        "comparison": "Quy trình C/O thay thế chuẩn (Được chấp nhận 100%, bảo lưu nguyên vẹn thuế ưu đãi 0%, hải quan nước ngoài hoàn tiền ký quỹ trong 3 ngày) VS Tự ý sửa chữa trên C/O cũ (Bị hải quan từ chối thẳng thừng, phạt gian lận xuất xứ, áp thuế kịch khung MFN).",
        "persona": "Anh Minh - Giám đốc doanh nghiệp xuất khẩu thủy sản và nông sản chế biến; Trưởng phòng xuất nhập khẩu phụ trách thủ tục C/O.",
        "checklist": "[ ] 1. Đã thu hồi bản gốc C/O bị lỗi chưa? | [ ] 2. Đã xác định chính xác số tham chiếu và ngày cấp của C/O cũ chưa? | [ ] 3. Đơn giải trình đã ghi rõ lỗi sai và cam kết tính trung thực chưa? | [ ] 4. Ô 'Issued in substitution' trên C/O mới đã điền đủ thông tin chưa? | [ ] 5. Đã gửi công hàm xác nhận pháp nhân nhà máy chưa?",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 05/2018/TT-BCT, Thông tư 33/2023/TT-BTC) & Sổ tay hướng dẫn Quy tắc xuất xứ Bộ Công Thương.",
        "keywords": ["rules", "origin"]
    },
    {
        "stt": 175,
        "week": "W43 — 23/10/2026",
        "title": "Checklist 6 Bước Nhập Khẩu Thực Phẩm Nguồn Gốc Động Vật & Thủy Sản Đông Lạnh Chuẩn Luật: Bí Quyết Thông Quan 100% Không Bị Giữ Hàng",
        "workflow": "Bước 1: Tra cứu Danh sách các cơ sở chế biến nước ngoài được phép xuất khẩu thịt và thủy sản vào Việt Nam trên cổng thông tin Cục Thú y (cucthuy.gov.vn) để đảm bảo nhà cung cấp có tên và mã phê duyệt còn hiệu lực; Bước 2: Nộp hồ sơ xin cấp Giấy phép kiểm dịch động vật nhập khẩu trực tuyến trên Cổng Một cửa Quốc gia trước khi ký hợp đồng và xếp hàng lên tàu; Bước 3: Nhận Giấy phép kiểm dịch động vật từ Cục Thú y và gửi mẫu Health Certificate chuẩn cho bên bán làm việc với cơ quan thú y nước xuất khẩu; Bước 4: Khi hàng cập cảng Việt Nam, mở tờ khai hải quan và đăng ký kiểm tra vệ sinh thú y/an toàn thực phẩm trên Cổng Một cửa Quốc gia; Bước 5: Phối hợp Chi cục Thú y vùng lấy mẫu xét nghiệm vi sinh, tồn dư hóa chất tại bãi kiểm hóa cảng hoặc kho lạnh đủ điều kiện; Bước 6: Nhận Chứng thư kiểm dịch an toàn thực phẩm đạt yêu cầu và nộp lên hệ thống VNACCS để thông quan giải phóng hàng.",
        "formula": "Chi phí kiểm soát rủi ro = 0 đồng (Chỉ cần 5 phút kiểm tra danh sách nhà máy tại Bước 1) VS Chi phí thiệt hại nếu bỏ qua = Hàng bị buộc tái xuất + 40-80 triệu tiền phạt + 2.000 USD tiền điện lạnh lưu bãi.",
        "top3_mistakes": "1. Ký hợp đồng và chuyển cọc trước khi kiểm tra mã số nhà máy xuất khẩu có trong danh sách phê duyệt của Cục Thú y hay không; 2. Nhập hàng về cảng rồi mới nộp đơn xin Giấy phép kiểm dịch động vật (hàng sẽ bị xử phạt vi phạm hành chính ngay tại cửa khẩu); 3. Thông tin tên nhà máy hoặc loài động vật trên Health Certificate không khớp từng chữ với Giấy phép kiểm dịch nhập khẩu.",
        "mnemonic_tip": "'Check Nhà Máy Trước - Hàng Xuôi Nước Sau' (Quy tắc sống còn của dân nhập khẩu thịt đông lạnh: Chưa thấy mã duyệt của Cục Thú y trên web thì tuyệt đối không mở L/C hay chuyển tiền cọc).",
        "comparison": "Quy trình tuân thủ 6 bước (Thông quan trơn tru trong 48 giờ sau khi có kết quả mẫu test, đưa hàng về kho phân phối ngay) VS Quy trình đốt cháy giai đoạn (Hàng bị giam lỏng tại cảng 20-30 ngày, tiền điện lạnh ăn mòn hết vốn, đối mặt nguy cơ bị tiêu hủy hoặc tái xuất).",
        "persona": "Chị Lan - Giám đốc thu mua chuỗi nhà hàng ẩm thực và thực phẩm cao cấp; Doanh nghiệp mới gia nhập thị trường nhập khẩu thịt bò, cá hồi, thịt heo đông lạnh.",
        "checklist": "[ ] 1. Mã số nhà máy nước ngoài có trong danh sách Cục Thú y phê duyệt chưa? | [ ] 2. Đã có Giấy phép kiểm dịch động vật của Cục Thú y cấp chưa? | [ ] 3. Bản gốc Health Certificate của nước xuất khẩu đã về tay chưa? | [ ] 4. Đã đăng ký kiểm tra trên Cổng Một cửa Quốc gia chưa? | [ ] 5. Kho lạnh bảo quản đã được cấp mã đủ điều kiện vệ sinh thú y chưa? | [ ] 6. Đã hẹn lịch cán bộ thú y lấy mẫu kiểm hóa chưa?",
        "url": "https://en.wikipedia.org/wiki/Quarantine",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 15/2018/NĐ-CP, Thông tư 26/2016/TT-BNNPTNT, Luật Thú y 2015) & Cổng thông tin Cục Thú y.",
        "keywords": ["quarantine"]
    }
]
