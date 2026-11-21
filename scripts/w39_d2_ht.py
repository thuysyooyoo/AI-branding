# -*- coding: utf-8 -*-
"""
DATASET: HỌC THỰC CHIẾN (HT) - BATCH W39 DAY 2 (STT 91 - 95)
Mốc thời gian: W39 — 23/09/2026
Tuân thủ: wf06_ideas_bank_update.md, 14 trường thông tin chuẩn xác
"""

HOC_THUC_CHIEN_W39_D2 = [
    {
        "stt": 91,
        "week": "W39 — 23/09/2026",
        "title": "4 Bước Thẩm Định & Áp Đúng Mã Loại Hình Hải Quan A11, A12 Và A41: Nguyên Tắc Bất Di Bất Dịch 'Bản Chất Đi Trước, Tên Hàng Đi Sau'",
        "steps": "Bước 1: Thẩm tra tư cách pháp nhân (DNCX hay Doanh nghiệp FDI thương mại hay DN 100% vốn trong nước); Bước 2: Bóc tách bản chất giao dịch và quyền xuất nhập khẩu (Hàng đưa vào dây chuyền sản xuất nội bộ hay bán phân phối ra thị trường theo quyền nhập khẩu?); Bước 3: Đối chiếu bảng mã chuẩn tại Quyết định 1357/QĐ-TCHQ và Công văn giải đáp 1478/TCHQ-GSQL; Bước 4: Soạn hồ sơ giải trình mục đích sử dụng nếu phát sinh chuyển đổi mục đích (chuyển loại hình từ E31/A12 sang A41/A11).",
        "formula": "Thuế Nhập Khẩu Phải Nộp Khi Chuyển Mục Đích = Trị Giá Hải Quan Tính Thuế Tại Thời Điểm Thay Đổi x Mức Thuế Suất Quy Định Tại Thời Điểm Đăng Ký Tờ Khai Mới (Điều 21 Thông tư 38/2015/TT-BTC sửa đổi tại TT 39/2018/TT-BTC).",
        "top3_errors": "1. Nhìn tên hàng hóa để phán đoán mã loại hình mà quên xem xét quyền phân phối của công ty; 2. Tự ý mở mã A11 để lấy hàng nhanh cho DNCX mà không làm thủ tục nộp thuế chuyển đổi mục đích; 3. Không lưu trữ hồ sơ chứng minh dòng tiền thanh toán và hợp đồng bán lẻ nội địa để phục vụ kiểm tra sau thông quan.",
        "mnemonic": "'Tư cách - Mục đích - Quyền phân phối': Nhớ quy tắc kiềng 3 chân, thiếu 1 chân là rụng tờ khai vào luồng Đỏ.",
        "comparison": "A11 (Nhập kinh doanh tiêu dùng: Mua về bán thương mại ngay hoặc dùng nội bộ DN nội địa) vs A12 (Nhập kinh doanh sản xuất: Đưa trực tiếp vào nhà máy tạo thành phẩm) vs A41 (Nhập theo quyền nhập khẩu của DN FDI: Hàng không đưa vào sản xuất mà bán buôn/lẻ theo giấy phép kinh doanh phân phối).",
        "persona": "Chị Hương (Kế toán trưởng XNK phụ trách thuế và quyết toán) & Nhân viên chứng từ đại lý hải quan.",
        "checklist": "[ ] Kiểm tra Giấy chứng nhận đầu tư có quyền nhập khẩu phân phối chưa; [ ] Xác định hàng dùng để sản xuất hay bán lại; [ ] Đối chiếu mã theo QĐ 1357 & CV 1478; [ ] Lưu hợp đồng mua bán và giải trình mục đích sử dụng.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "keywords": ["customs", "declaration"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Thực chiến xây dựng từ Công văn 21939/HQKV2-NVHQ ngày 21/09/2026 và kiến thức chuẩn Tổng cục Hải quan."
    },
    {
        "stt": 92,
        "week": "W39 — 23/09/2026",
        "title": "SOP 5 Bước Chuẩn Hóa 100% Trường Dữ Liệu Khai Báo Hải Quan Điện Tử: Triệt Tiêu Lỗi Lệch Số Khung, Số Máy, Số Container Và Mã HS Trước Khi Bấm Truyền",
        "steps": "Bước 1: Thu thập và số hóa dữ liệu nguồn từ B/L gốc, Commercial Invoice, Packing List và Chứng thư giám định chuyên ngành; Bước 2: Sử dụng công cụ OCR/Excel Validator đối soát tự động trùng khớp giữa số container, số seal trên Manifest hãng tàu và tờ khai; Bước 3: Nhập dữ liệu định danh riêng biệt (Số CCCD, Biển số xe, Số khung/máy, Số IMO) vào đúng trường dữ liệu số hóa quy định tại Nghị định 142/2026/NĐ-CP; Bước 4: Kiểm tra chéo mã HS 8 chữ số và mô tả hàng hóa theo Danh mục Biểu thuế XNK hiện hành; Bước 5: Chạy lệnh tiền kiểm tra (Pre-audit check) trên phần mềm VNACCS trước khi ký số truyền chính thức.",
        "formula": "Tỷ Lệ Sai Sót Dữ Liệu Trường Khai Báo = (Số Lỗi Định Dạng Bị Phần Mềm Reject / Tổng Số Trường Thông Tin Bắt Buộc) x 100%. Tiêu chuẩn chuẩn hóa Eureka: Target 0% lỗi kỹ thuật.",
        "top3_errors": "1. Viết tắt hoặc nhồi nhét số khung/số máy vào ô mô tả hàng hóa hoặc ô ghi chú chung; 2. Nhập sai định dạng mã container (bỏ quên mã chủ vỏ 4 ký tự chữ cái đầu hoặc số kiểm tra cuối); 3. Khai sai số định danh người nhận hàng (nhập mã nội bộ thay vì số CCCD/MST chuẩn quốc gia).",
        "mnemonic": "'Tách biệt - Đúng ô - Khớp nguồn': Mỗi thông tin một trường riêng, không bao giờ nhồi nhét như viết văn xuôi.",
        "comparison": "Khai báo truyền thống (Ghi chung trong ô mô tả, kiểm tra thủ công bằng mắt) vs Chuẩn hóa trường số hóa (Dữ liệu bóc tách độc lập, máy chủ AI tự động quét và duyệt thông quan tự động).",
        "persona": "Anh Minh (Chủ doanh nghiệp muốn tối ưu tốc độ thông quan) & Đội ngũ chuyên viên chứng từ Logistics.",
        "checklist": "[ ] Đối soát B/L và Manifest hãng tàu; [ ] Kiểm tra số container đủ 11 ký tự; [ ] Nhập số khung/số máy vào trường dedicated; [ ] Pre-check VNACCS trước khi ký số Token.",
        "url": "https://en.wikipedia.org/wiki/Single-window_system",
        "keywords": ["single", "window"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "SOP chuẩn hóa vận hành bám sát Nghị định 142/2026/NĐ-CP và chia sẻ thực tế từ Lê Phương Logistics trên daily_notes.md."
    },
    {
        "stt": 93,
        "week": "W39 — 23/09/2026",
        "title": "Cẩm Nang 4 Bước Xử Lý Khôi Phục Mã Số Thuế Bị Khóa Trạng Thái 06 & Giải Thể Doanh Nghiệp Hợp Pháp Không Lo Bị Phạt Chồng Phạt",
        "steps": "Bước 1: Tra cứu thực trạng mã số thuế trên Hệ thống Thông tin Đăng ký Doanh nghiệp Quốc gia (dangkykinhdoanh.gov.vn) và ứng dụng eTax Mobile để xác định chính xác ngày bị chuyển Trạng thái 06; Bước 2: Chuẩn bị hồ sơ đề nghị khôi phục MST gồm: Mẫu 25/ĐK-TCT, Hợp đồng thuê trụ sở hợp pháp mới (hoặc xác nhận của địa phương), Bản giải trình lý do không hoạt động; Bước 3: Làm việc với Đội Kiểm tra Thuế để đối chiếu toàn bộ tờ khai còn thiếu, lập biên bản vi phạm hành chính và nộp phạt nghĩa vụ thuế môn bài/tờ khai tích lũy; Bước 4: Sau khi MST được mở lại Trạng thái 00 (Đang hoạt động), tiến hành quyết toán thuế đầy đủ, nộp hồ sơ đóng MST và nộp hồ sơ giải thể tại Sở KH&ĐT.",
        "formula": "Tổng Chi Phí Nộp Phạt Do Bỏ Địa Chỉ = Tiền Phạt Chậm Nộp Tờ Khai (tối đa 25 triệu/kỳ nếu quá hạn >90 ngày) + Tiền Thuế Môn Bài Nợ + Tiền Chậm Nộp 0.03%/ngày trên số thuế nợ (Điều 59 Luật Quản lý Thuế).",
        "top3_errors": "1. Bỏ mặc công ty không hoạt động mà không nộp tờ khai trắng, nghĩ rằng 'không có doanh thu thì không bị phạt'; 2. Thuê dịch vụ 'ma' nhận tiền rồi bỏ trốn, không lấy Giấy xác nhận hoàn tất nghĩa vụ thuế từ cơ quan thuế; 3. Cố tình mở công ty mới khi vẫn đang là người đại diện pháp luật của công ty bị khóa Trạng thái 06.",
        "mnemonic": "'Khôi phục trước - Nộp phạt đủ - Quyết toán xong - Mới được đóng': Không có đường tắt cho việc đóng mã số thuế.",
        "comparison": "Bỏ mặc công ty (Khóa MST vĩnh viễn, phạt lũy kế hàng trăm triệu, cấm xuất cảnh) vs Giải thể chính ngạch (Tất toán sổ sách, đóng MST sạch sẽ, bảo vệ danh dự và tự do di chuyển của chủ DN).",
        "persona": "Anh Dũng (Nhà quản trị muốn dọn dẹp pháp nhân sạch sẽ) & Kế toán trưởng.",
        "checklist": "[ ] Tra cứu trạng thái MST trên eTax; [ ] Lập mẫu 25/ĐK-TCT khôi phục MST; [ ] Đối chiếu và nộp tiền phạt hành chính; [ ] Xin giấy xác nhận hoàn thành nghĩa vụ thuế tại Chi cục Thuế.",
        "url": "https://en.wikipedia.org/wiki/Corporate_tax",
        "keywords": ["tax", "corporate"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Cẩm nang pháp lý xử lý case nợ thuế thực tế từ Diễn đàn Thuế chuyên sâu và daily_notes.md ngày 23/09/2026."
    },
    {
        "stt": 94,
        "week": "W39 — 23/09/2026",
        "title": "5 Bước Tự Xin Giấy Xác Nhận Hoàn Thành Nghĩa Vụ Thuế Hải Quan Cấp Tốc Tại Chi Cục Hải Quan KV2 (02 Hàm Nghi) Cho Lô Hàng Dự Án",
        "steps": "Bước 1: Đăng nhập Cổng Thông tin Hải quan điện tử (customs.gov.vn) vào mục 'Tra cứu nợ thuế của doanh nghiệp' để kiểm tra xem có khoản nợ thuế, tiền phạt chậm nộp hoặc lệ phí hải quan phát sinh chưa nộp; Bước 2: Soạn Công văn đề nghị xác nhận hoàn thành nghĩa vụ thuế theo mẫu chuẩn của Tổng cục Hải quan, nêu rõ mục đích (phục vụ thông quan lô hàng dự án, tham gia đấu thầu hoặc giải thể); Bước 3: Nộp trực tiếp tại Bộ phận Một cửa / Văn thư Chi cục Hải quan KV2 (Số 02 Hàm Nghi, Bến Nghé, Quận 1, TP.HCM), lấy phiếu hẹn tiếp nhận và số công văn đến; Bước 4: Cung cấp mã số thuế cho cán bộ văn thư tra cứu trên hệ thống nội bộ để xử lý luân chuyển hồ sơ nhanh; Bước 5: Nhận văn bản trả lời chính thức trong vòng 03-05 ngày làm việc và lưu bản scan có đóng dấu đỏ để đính kèm tờ khai thông quan.",
        "formula": "Thời Gian Xử Lý = Tra Cứu Hệ Thống Số (05 phút) + Nộp Hồ Sơ Một Cửa (01 ngày) + Trả Kết Quả Bằng Văn Bản (03-05 ngày làm việc theo quy chế).",
        "top3_errors": "1. Để sót khoản nợ lệ phí hải quan nhỏ (chỉ 20.000 VNĐ/tờ khai) khiến hệ thống tự động treo nợ xấu; 2. Đi thẳng vào phòng chuyên môn gặp cán bộ hải quan thay vì vào bộ phận Văn thư Một cửa tại số 02 Hàm Nghi; 3. Không đối soát trước chứng từ nộp thuế qua ngân hàng với cơ sở dữ liệu kho bạc khiến dữ liệu bị trễ.",
        "mnemonic": "'Tra online trước - Lấy số văn thư - Đối chiếu kho bạc': Đi đúng luồng Một cửa, tránh lòng vòng mất việc.",
        "comparison": "Xin xác nhận thủ công (Mất 7-10 ngày, đi lại nhiều lần) vs Quy trình chuẩn bị dữ liệu đồng bộ (Chỉ mất 3 ngày, nhận văn bản dấu đỏ chuẩn chỉnh).",
        "persona": "Chị Hương (Kế toán XNK) & Chuyên viên giao nhận hiện trường (Ops).",
        "checklist": "[ ] Tra nợ thuế online trên customs.gov.vn; [ ] Soạn công văn chuẩn form; [ ] Nộp văn thư 02 Hàm Nghi lấy số tiếp nhận; [ ] Theo dõi trả kết quả.",
        "url": "https://en.wikipedia.org/wiki/Customs_declaration",
        "keywords": ["customs", "declaration"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đúc kết từ chỉ dẫn thực địa của Cẩm Huyền tại Chi cục HQ KV2 và daily_notes.md ngày 23/09/2026."
    },
    {
        "stt": 95,
        "week": "W39 — 23/09/2026",
        "title": "Kỹ Thuật Đàm Phán Giảm 30% - 50% Phí Lưu Bãi DEM/DET Với Hãng Tàu Quốc Tế Khi Hàng Hóa Bị Giữ Kiểm Tra Chuyên Ngành",
        "steps": "Bước 1: Thu thập đầy đủ chứng cứ khách quan chứng minh nguyên nhân chậm trễ không do lỗi chủ quan của chủ hàng (Biên bản lấy mẫu kiểm định chuyên ngành của Hải quan, Quyết định trưng cầu giám định, Giấy xác nhận bất khả kháng do thiên tai/dịch bệnh); Bước 2: Làm việc ngay với hãng tàu tại thời điểm phát sinh sự cố (trước khi thời hạn Free time kết thúc 03 ngày) để thông báo tình trạng container và đề xuất gia hạn Free time bổ sung; Bước 3: Soạn Công văn đề nghị miễn giảm phí lưu container/lưu bãi (Demurrage / Detention Waiver Letter) gửi trực tiếp Trưởng đại diện hãng tàu, cam kết lịch rút ruột vỏ và thanh toán cước đầy đủ; Bước 4: Đàm phán áp dụng mức phí ưu đãi theo bậc lũy thoái hoặc xin chuyển container về bãi cạn ICD có đơn giá lưu bãi rẻ hơn 60% so với cảng chính; Bước 5: Ký phụ lục thỏa thuận và thanh toán quyết toán dứt điểm để lấy D/O (Delivery Order) điện tử.",
        "formula": "Tỷ Lệ Tiết Kiệm Chi Phí DEM/DET = [(Tổng Phí Biểu Cước Gốc - Phí Đã Đàm Phán Giảm) / Tổng Phí Biểu Cước Gốc] x 100%. Mức chiết khấu thông thường đạt từ 30% đến 50% tổng số tiền phát sinh.",
        "top3_errors": "1. Chờ đến khi lấy được hàng mới gửi đơn xin giảm phí khi hóa đơn đã bị hệ thống chốt số tiền phạt; 2. Đổ lỗi cho hãng tàu thay vì cung cấp hồ sơ khách quan từ cơ quan nhà nước; 3. Không xin chuyển container về bãi Depot/ICD ngoài cảng để cắt đứt mức phí lũy tiến khổng lồ tại cầu cảng.",
        "mnemonic": "'Thông báo sớm - Bằng chứng chuẩn - Xin chuyển bãi - Cắt lũy tiến': Đàm phán là chạy đua với thời gian tính phí.",
        "comparison": "Phí DEM tại cầu cảng (Lũy tiến gấp đôi từ ngày thứ 8: $80 -> $160/ngày/FEU) vs Phí lưu bãi tại ICD vệ tinh (Giá cố định chỉ $15 - $25/ngày/FEU).",
        "persona": "Anh Minh (Chủ doanh nghiệp) & Trưởng phòng Logistics.",
        "checklist": "[ ] Thu thập biên bản hải quan kiểm tra; [ ] Gửi văn bản thông báo cho hãng tàu trước khi hết Free time; [ ] Xin chuyển container về ICD nội địa; [ ] Đàm phán waiver trực tiếp cấp quản lý.",
        "url": "https://en.wikipedia.org/wiki/Demurrage",
        "keywords": ["demurrage", "detention"],
        "series": "Đàm phán Quốc tế",
        "status": "Chưa dùng",
        "note": "Chiến thuật đàm phán thực tế từ chuyên gia Logistics Eureka xử lý giải cứu hàng trăm container tồn bãi."
    }
]
