# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W41 DAY 4 — 09/10/2026)
Số lượng: 5 ý tưởng (STT 136 - 140)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W41_D4 = [
    {
        "stt": 136,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Cẩm Nang 4 Bước Khai Báo & Truyền Dữ Liệu e-C/O Mẫu D & AK/VK Qua Cổng Một Cửa Quốc Gia (NSW) Tránh Lỗi Lệch Dữ Liệu Cửa Khẩu",
        "workflow": "Bước 1: Chuẩn Hóa Hồ Sơ Gốc & Mã UN/LOCODE: Tập hợp Tờ khai hải quan xuất khẩu đã thông quan, Commercial Invoice, Packing List và Vận đơn B/L; tra cứu mã định danh quốc tế chuẩn UN/LOCODE của cảng bốc (POL) và cảng dỡ (POD); Bước 2: Kê Khai Dữ Liệu Trên Hệ Thống Quản Lý C/O: Nhập chính xác từng trường thông tin trên hệ thống cấp C/O điện tử của cơ quan được phân cấp tại địa phương; kiểm tra tiêu chí xuất xứ (WO, PE, RVC, CTC); Bước 3: Đồng Bộ & Kiểm Tra Mã Tra Cứu Trên Cổng Một Cửa Quốc Gia: Nhập số tờ khai hải quan và số C/O lên Cổng thông tin một cửa quốc gia (national.singlewindow.gov.vn), kiểm tra trạng thái truyền dữ liệu sang hải quan nước đối tác; Bước 4: Lưu Trữ Bản Điện Tử & Cung Cấp Mã QR/Tham Chiếu Cho Đối Tác: Gửi số tham chiếu C/O (Reference Number) và link tra cứu xác thực trực tuyến cho nhà nhập khẩu nước ngoài để mở tờ khai ưu đãi thuế quan.",
        "formula": "Tỷ lệ hàm lượng giá trị khu vực (RVC) theo công thức gián tiếp = [(Trị giá FOB - Trị giá nguyên liệu không có xuất xứ VNM) / Trị giá FOB] x 100%. Tiêu chuẩn RVC >= 40% đối với Form D (ATIGA) và Form AK/VK.",
        "top3_mistakes": "1. Gõ sai mã số định danh cảng biển quốc tế (UN/LOCODE) dẫn đến hải quan nước nhập khẩu không nhận diện được chuyến tàu; 2. Số lượng và trọng lượng gross/net trên C/O lệch với Tờ khai hải quan thông quan; 3. Không kiểm tra trạng thái tiếp nhận thành công trên NSW trước khi tàu cập cảng.",
        "mnemonic_tip": "Quy tắc '3 Khớp Một Cửa': Số tờ khai khớp - Tên cảng khớp - Trọng lượng khớp. Ba khớp thì thông quan trong 3 nốt nhạc!",
        "comparison": "C/O giấy truyền thống (Phải đến cơ quan nộp hồ sơ trực tiếp, chờ ký đóng dấu mực, gửi phát chuyển nhanh quốc tế mất 3-5 ngày) vs e-C/O qua NSW (Khai báo trực tuyến, duyệt điện tử trong 4 giờ, dữ liệu tự động truyền qua cổng hải quan hai nước).",
        "persona": "Chị Hương (Chuyên viên chứng từ XNK) & Anh Minh (Chủ doanh nghiệp chế biến nông thủy sản xuất khẩu)",
        "checklist": "[ ] Kiểm tra mã UN/LOCODE cảng biển; [ ] Đối chiếu số tờ khai thông quan trên NSW; [ ] Kiểm tra tiêu chí RVC/CTC; [ ] Xác nhận trạng thái 'Đã đồng bộ sang nước nhập khẩu' trên portal.",
        "url": "https://en.wikipedia.org/wiki/Single-window_system",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP, Cổng một cửa quốc gia NSW) & daily_notes.md (ghi nhận ngày 09/10/2026).",
        "keywords": ["single-window", "customs"]
    },
    {
        "stt": 137,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Quy Trình 4 Bước Đăng Ký & Điều Chỉnh Địa Điểm Kinh Doanh (ĐĐKD) Kho Hàng XNK Trên Cổng DVC Tránh Bị Phạt 15 Triệu Đồng",
        "workflow": "Bước 1: Chuẩn Bị Hồ Sơ Điện Tử Theo Mẫu Phụ Lục II-7 (Thông tư 01/2021/TT-BKHĐT): Thông báo lập/thay đổi địa điểm kinh doanh, hợp đồng thuê kho/mặt bằng, giấy ủy quyền thực hiện thủ tục và chữ ký số token (hoặc tài khoản VNeID cấp độ 2); Bước 2: Nộp Hồ Sơ Trực Tuyến Trên Cổng Thông Tin Đăng Ký Doanh Nghiệp Quốc Gia: Đăng nhập dangkykinhdoanh.gov.vn hoặc Cổng DVC tỉnh, tạo hồ sơ đăng ký doanh nghiệp trực tuyến, nhập chính xác địa chỉ chi tiết (số nhà, đường, phường/xã, quận/huyện); Bước 3: Nhận Kết Quả Giấy Chứng Nhận Đăng Ký ĐĐKD & Kiểm Tra Mã 5 Số: Sau 3 ngày làm việc, tải Giấy xác nhận có mã số địa điểm kinh doanh (Mã dạng MST-0000X); đối chiếu mã số tự động liên thông với Cổng thông tin Tổng cục Thuế; Bước 4: Kê Khai Lệ Phí Môn Bài & Thông Báo Địa Điểm Lưu Kho Với Hải Quan: Nộp tờ khai lệ phí môn bài Mẫu 01/LPMB (mức 1.000.000 VNĐ/năm) và cập nhật mã địa điểm kho vào danh mục quản lý kho bãi phục vụ kiểm tra sau thông quan.",
        "formula": "Lệ phí môn bài cho Địa điểm kinh doanh / Chi nhánh = 1.000.000 VNĐ/năm (nếu thành lập trong 6 tháng đầu năm nộp 100%, 6 tháng cuối năm nộp 50% = 500.000 VNĐ). Mức phạt không thông báo địa điểm kinh doanh: Phạt tiền từ 10.000.000 đến 15.000.000 VNĐ theo Điều 45 Nghị định 122/2021/NĐ-CP.",
        "top3_mistakes": "1. Vào nhầm mục 'Tra cứu thông tin người nộp thuế' trên web Thuế rồi tưởng là đã nộp đơn đăng ký; 2. Thay đổi số nhà/địa chỉ kho bãi nhưng không làm thủ tục thông báo thay đổi trên Cổng DVC trong vòng 10 ngày; 3. Không nộp tờ khai môn bài dẫn đến bị cơ quan thuế khóa mã số địa điểm kinh doanh và cưỡng chế hóa đơn.",
        "mnemonic_tip": "'Mở kho là mở số - 3 ngày có mã 5 số': Cứ lập kho mới là phải xin mã 5 số (00001, 00002...), không có mã thì đừng vội đưa hàng vào!",
        "comparison": "Địa điểm kinh doanh (Chỉ là nơi chứa hàng/giao dịch phụ thuộc, không có con dấu riêng, không có MST riêng, thuế hạch toán phụ thuộc công ty mẹ) vs Chi nhánh (Có quyền tự ký hợp đồng, có thể hạch toán độc lập, có con dấu và MST 13 số riêng).",
        "persona": "Chị Hương (Kế toán trưởng phụ trách tài sản) & Anh Minh (Chủ doanh nghiệp xuất nhập khẩu mở rộng mạng lưới phân phối)",
        "checklist": "[ ] Soạn Thông báo Phụ lục II-7 chuẩn form; [ ] Nộp online qua dangkykinhdoanh.gov.vn; [ ] Nhận giấy xác nhận mã ĐĐKD 5 số; [ ] Khai lệ phí môn bài 1 triệu đồng; [ ] Đăng ký địa điểm kho với chi cục hải quan quản lý.",
        "url": "https://en.wikipedia.org/wiki/Company_register",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 01/2021/NĐ-CP, Nghị định 122/2021/NĐ-CP) & daily_notes.md (ghi nhận ngày 09/10/2026).",
        "keywords": ["register", "business"]
    },
    {
        "stt": 138,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Cẩm Nang 4 Bước Thiết Lập Hệ Thống Kiểm Soát An Ninh Chuỗi Cung Ứng Đạt Chuẩn Doanh Nghiệp Ưu Tiên AEO / WCO SAFE",
        "workflow": "Bước 1: Đánh Giá Rủi Ro Điểm Nghẽn An Ninh (Threat & Vulnerability Assessment): Rà soát toàn bộ các mắt xích từ nhà máy sản xuất, nhà cung cấp nguyên liệu, đơn vị kho bãi đến đơn vị vận tải đường bộ/biển; phát hiện các điểm hở có nguy cơ bị cài cắm hàng cấm hoặc đánh tráo hàng hóa; Bước 2: Thiết Lập Tiêu Chuẩn An Ninh Phương Tiện Vận Tải 7 Điểm (7-Point Container Inspection): Kiểm tra thành trước, bên trái, bên phải, sàn, trần/nóc, cửa trong/ngoài và gầm xe container trước khi đóng hàng và bấm niêm phong chì seal ISO 17712; Bước 3: Kiểm Toán Nhà Thầu Phụ Trợ Vận Tải & Kho Bãi (Vendor Due Diligence): Yêu cầu các nhà xe vệ tinh cam kết tuân thủ quy chế an ninh AEO, kiểm tra lý lịch tư pháp tài xế và tích hợp hệ thống định vị GPS hành trình giám sát thời gian thực; Bước 4: Xây Dựng Quy Trình Ứng Phó Sự Cố Khẩn Cấp & Báo Cáo Hải Quan: Soạn thảo kịch bản phản ứng nhanh khi chì seal bị đứt hoặc phương tiện đi lệch hành trình, báo cáo ngay cho cơ quan hải quan trong vòng 2 giờ kể từ khi phát hiện dấu hiệu bất thường.",
        "formula": "Điều kiện kim ngạch để nộp hồ sơ Doanh nghiệp ưu tiên (AEO): Kim ngạch XNK tối thiểu 100 triệu USD/năm (đối với DN thông thường); tối thiểu 40 triệu USD/năm (đối với DN nông, thủy sản Việt Nam); hoặc tối thiểu 30 triệu USD/năm (đối với DN xuất khẩu hàng hóa sản xuất tại Việt Nam).",
        "top3_mistakes": "1. Xem nhẹ khâu kiểm tra vật lý container 7 điểm trước khi đóng hàng; 2. Không kiểm toán quy trình an ninh của đối tác vận tải phụ trợ mà chỉ dựa vào hợp đồng giấy tờ; 3. Giấu giếm sự cố đứt seal hoặc mất hàng thay vì chủ động khai báo ngay với cơ quan hải quan.",
        "mnemonic_tip": "'Kiểm tra 7 điểm - Bấm seal chuẩn ISO': Trước khi cho hàng vào thùng, soi kỹ 7 điểm (Trước - Trái - Phải - Sàn - Trần - Cửa - Gầm), bấm seal chì chuẩn thì an tâm ngủ ngon!",
        "comparison": "Doanh nghiệp thông thường (Bị phân luồng ngẫu nhiên Vàng/Đỏ, kiểm tra chứng từ và kiểm hóa thực tế tại cảng, nộp thuế trước khi thông quan) vs Doanh nghiệp AEO (Thông quan bằng dữ liệu điện tử, miễn kiểm tra chứng từ và hàng hóa, được nộp thuế chậm đến ngày 10 của tháng kế tiếp).",
        "persona": "Anh Dũng (Giám đốc Chuỗi cung ứng & Logistics) & Chị Hương (Trưởng phòng Pháp chế tuân thủ doanh nghiệp lớn)",
        "checklist": "[ ] Thực hiện biên bản kiểm tra container 7 điểm; [ ] Sử dụng chì niêm phong đạt chuẩn ISO 17712 High Security Seal; [ ] Giám sát lộ trình xe bằng GPS thời gian thực; [ ] Định kỳ kiểm toán an ninh nhà thầu vận tải 6 tháng/lần.",
        "url": "https://en.wikipedia.org/wiki/Authorized_economic_operator",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 72/2015/TT-BTC) & WCO SAFE Framework of Standards.",
        "keywords": ["authorized", "economic", "operator"]
    },
    {
        "stt": 139,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Cẩm Nang 4 Bước Xử Lý Khẩn Cấp Khi Bị Thất Lạc Bộ Vận Đơn Gốc (Original B/L) & Quy Trình Lập Thư Bảo Lãnh Giao Hàng (LOI)",
        "workflow": "Bước 1: Thông Báo Khẩn Cấp Bằng Văn Bản & Phong Tỏa Vận Đơn: Gửi ngay văn bản thông báo mất B/L gốc cho Hãng tàu, Đại lý giao nhận, Ngân hàng thanh toán và Cơ quan cảng biển yêu cầu dừng ngay việc giao hàng nếu có bên thứ ba xuất trình; Bước 2: Thu Thập Xác Nhận Thất Lạc Của Đơn Vị Bưu Chính: Yêu cầu hãng chuyển phát nhanh (DHL/FedEx/Bưu điện) cấp Biên bản xác nhận thất lạc bưu phẩm chính thức có chữ ký và dấu mộc; Bước 3: Đàm Phán Phát Hành Thư Bảo Lãnh Ngân Hàng (LOI - Letter of Indemnity): Làm việc với ngân hàng để phát hành thư bảo lãnh theo mẫu chuẩn của Hãng tàu/Câu lạc bộ P&I (thường yêu cầu ký quỹ 100% - 110% giá trị hàng hóa trong thời hạn 1 đến 2 năm); Bước 4: Yêu Cầu Hãng Tàu Giải Phóng Hàng & Thu Hồi Tiền Ký Quỹ: Sau khi hãng tàu chấp thuận LOI và cấp lệnh giao hàng (D/O) cho người nhận hàng, theo dõi hết thời hiệu khởi kiện tranh chấp hàng hải (thường là 1 năm theo Quy tắc Hague-Visby hoặc 2 năm theo Bộ luật Hàng hải VN) để làm thủ tục hoàn trả tiền ký quỹ.",
        "formula": "Số tiền ký quỹ bảo lãnh ngân hàng LOI khi mất B/L gốc = Trị giá hóa đơn lô hàng (Commercial Invoice) x 110% (hoặc 120% tùy quy định từng hãng tàu). Thời gian phong tỏa: Tối thiểu 12 đến 24 tháng.",
        "top3_mistakes": "1. Giấu thông tin mất B/L để tự tìm kiếm làm trễ thời hạn, để hàng lưu bãi phát sinh tiền demurrage hàng chục nghìn USD; 2. Gửi trọn bộ 3 bản gốc B/L trong cùng 1 phong bì chuyển phát; 3. Không kiểm tra thời hiệu bảo lãnh khiến tiền ký quỹ bị ngân hàng giữ quá hạn nhiều năm.",
        "mnemonic_tip": "'Mất B/L như mất vàng - Báo ngay hãng tàu chặn hàng đầu tiên': Báo khẩn cấp để phong tỏa hàng, lập LOI giải cứu hàng trước, tính chuyện bồi thường bưu điện sau!",
        "comparison": "Vận đơn đích danh (Straight B/L - Không thể chuyển nhượng, giao đúng người nhận ghi tên) vs Vận đơn theo lệnh (To Order B/L - Có thể chuyển nhượng bằng ký hậu, nếu mất thì rủi ro cực lớn).",
        "persona": "Anh Minh (Giám đốc công ty xuất khẩu nông sản) & Chị Hương (Kế toán trưởng quản lý dòng tiền)",
        "checklist": "[ ] Gửi văn bản khẩn phong tỏa B/L cho hãng tàu; [ ] Lấy biên bản xác nhận mất của bưu chính; [ ] Đàm phán giảm tỷ lệ ký quỹ LOI tại ngân hàng; [ ] Cấp lệnh D/O giải phóng hàng tránh demurrage; [ ] Lên lịch giải tỏa tiền phong tỏa sau 1 năm.",
        "url": "https://en.wikipedia.org/wiki/Bill_of_lading",
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: BIMCO Standard LOI Clauses & Bộ luật Hàng hải Việt Nam 2015 (Điều 159 về trả hàng khi mất vận đơn).",
        "keywords": ["bill of lading", "carrier"]
    },
    {
        "stt": 140,
        "week": "W41 Day 4 — 09/10/2026",
        "title": "Cẩm Nang 4 Bước Tối Ưu Hóa Tuyến Vận Tải Sà Lan Đường Thủy Kết Nối Cảng Biển Giúp Tiết Kiệm 30% Chi Phí Logistics",
        "workflow": "Bước 1: Khảo Sát & Lựa Chọn Cảng Thủy Nội Địa / Cảng Cạn ICD Phù Hợp: Xác định ICD hoặc bến thủy nội địa gần nhà máy nhất có đủ năng lực cẩu nâng hạ container (ví dụ: ICD Phúc Long, ICD Sóng Thần, Cảng Thốt Nốt, Cảng Tiên Sa); Bước 2: Đàm Phán Hợp Đồng Cước Sà Lan Trọn Gói (All-in Barge Rate): Ký hợp đồng với hãng sà lan bao gồm cước vận chuyển thủy, phí xếp dỡ hai đầu cảng (THC barge) và cam kết thời gian hành trình (Transit time); Bước 3: Đăng Ký Thủ Tục Giảm 50% Phí Hạ Tầng Cảng Biển: Khi kê khai nộp phí hạ tầng cảng biển trên hệ thống thu phí trực tuyến, chọn phương thức vận chuyển 'Đường thủy nội địa' và đính kèm số hiệu sà lan để được hệ thống tự động áp mức giảm 50%; Bước 4: Kiểm Soát Thời Gian Cắt Máng (Closing Time) Của Tàu Mẹ: Thiết lập kế hoạch đóng hàng và hạ sà lan trước giờ cắt máng của tàu mẹ tối thiểu 24-36 giờ để phòng ngừa rủi ro thủy triều và thời tiết trên sông.",
        "formula": "Chi phí tiết kiệm trên mỗi container = (Cước xe đầu kéo đường bộ - Cước sà lan trọn gói) + (Mức chênh lệch giảm 50% phí hạ tầng cảng biển). Ví dụ tuyến An Giang - Cái Mép: Tiết kiệm bình quân 1.500.000 - 2.200.000 VNĐ / container 40 feet.",
        "top3_mistakes": "1. Canh thời gian giao hàng sà lan quá sát giờ Closing time khiến container bị rớt tàu do gặp nước ròng hoặc thủy triều thấp; 2. Quên tích chọn phương thức 'Đường thủy nội địa' khi nộp phí hạ tầng trên cổng thanh toán khiến bị thu 100% mức phí đường bộ; 3. Không kiểm tra tải trọng tối đa của sà lan và tĩnh không cầu vào mùa mưa lũ.",
        "mnemonic_tip": "'Đi thủy giảm phí - Tính nước ròng trước': Đi sà lan giảm nửa tiền phí hạ tầng, nhưng nhớ trừ hao 1 ngày con nước thủy triều!",
        "comparison": "Vận chuyển đường bộ (Nhanh hơn 12-24h, cước cao gấp 1.5 lần, dễ kẹt xe giờ cao điểm, chịu 100% phí hạ tầng) vs Vận chuyển sà lan đường thủy (Chở số lượng lớn từ 50-200 TEU/chuyến, cước rẻ hơn 30%, không kẹt xe, giảm 50% phí hạ tầng cảng biển).",
        "persona": "Anh Minh (Chủ doanh nghiệp xuất khẩu hàng rời, nông sản, gạch men) & Anh Dũng (Trưởng phòng Logistics điều phối vận tải)",
        "checklist": "[ ] Đặt chỗ sà lan trước ngày xuất hàng 48 giờ; [ ] Khai báo giảm 50% phí hạ tầng cảng biển đường thủy; [ ] Kiểm tra lịch thủy triều và giờ tàu mẹ cập cảng; [ ] Giám sát thời điểm sà lan cập mạn tàu mẹ hoặc hạ bãi Cái Mép.",
        "url": "https://en.wikipedia.org/wiki/Barge",
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Luật Giao thông đường thủy nội địa) & Biểu cước vận tải thủy Tổng công ty Tân Cảng Sài Gòn.",
        "keywords": ["barge", "cargo"]
    }
]
