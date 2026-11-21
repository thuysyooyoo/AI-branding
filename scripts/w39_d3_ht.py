# -*- coding: utf-8 -*-
"""
DATASET: HỌC THỰC CHIẾN (HT) - BATCH W39 DAY 3 (STT 96 - 100)
Mốc thời gian: W39 — 25/09/2026
Tuân thủ: wf06_ideas_bank_update.md, 14 trường thông tin chuẩn xác
"""

HOC_THUC_CHIEN_W39_D3 = [
    {
        "stt": 96,
        "week": "W39 — 25/09/2026",
        "title": "Quy Trình 4 Bước Khai Báo Hải Quan Chuẩn Khi Container Bị Gãy Seal / Đổi Seal Trong Quá Trình Chuyển Cảng (Cát Lái -> Hiệp Phước / Cái Mép)",
        "steps": "Bước 1: Lập Biên bản ghi nhận sự cố gãy seal ngay tại bãi cảng (có xác nhận chữ ký của Đội trưởng kiểm soát bãi và Đại diện hãng tàu/cảng); Bước 2: Bấm seal niêm phong mới của đơn vị vận chuyển/cảng và chụp ảnh rõ nét 4 góc container; Bước 3: Khai báo số seal mới vào Hệ thống thông quan điện tử ECUS5-VNACCS kèm đính kèm scan biên bản lên phần Quản lý văn bản đính kèm (v5); Bước 4: Đến bàn giám sát hải quan tại cổng cảng đích đối chiếu dữ liệu VASSCM để giải phóng hàng.",
        "formula": "Phí phạt khai sai thông tin giám sát hải quan: 1.000.000 - 3.000.000 VNĐ/tờ khai (Khoản 1 Điều 7 NĐ 128/2020); Tiền lưu bãi trung bình khi bị treo barie: 1.500.000 - 2.500.000 VNĐ/cont/ngày.",
        "top3_errors": "1. Tài xế tự mua seal chợ bấm bừa không báo ai; 2. Khai theo số seal cũ của hãng tàu dù thực tế trên container đã đổi seal mới; 3. Không lưu giữ Biên bản thay seal của Cảng làm bằng chứng đối chất.",
        "mnemonic": "'Seal cũ trên giấy - Seal mới trên thùng - Lệch một con số - Bãi cảng nằm chung' (Nhớ nguyên tắc: Thùng bấm seal nào thì giấy phải mang đúng số đó kèm biên bản).",
        "comparison": "Khai theo Seal Hãng Tàu cũ (Bị máy quét cổng cảng từ chối, kẹt xe ở cổng) vs Khai theo Seal Cảng Mới kèm Biên bản (Hệ thống VASSCM tự động đồng bộ, barie mở trong 5 giây).",
        "persona": "Chị Hương (Kế toán kiêm chứng từ XNK) & Anh Dũng (Quản trị vận tải logistics container).",
        "checklist": "[ ] Biên bản gãy seal có mộc cảng; [ ] Ảnh chụp seal mới rõ số; [ ] Đã sửa số seal trên phần mềm ECUS; [ ] Đã đính kèm v5; [ ] Kiểm tra trạng thái VASSCM cổng.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "keywords": ["security", "seal"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% quy trình thực tế từ Quyết định 2065/QĐ-BTC và cẩm nang xử lý trên daily_notes.md ngày 25/09/2026."
    },
    {
        "stt": 97,
        "week": "W39 — 25/09/2026",
        "title": "Cẩm Nang 5 Bước 'Chạy Đua 30 Ngày' Đưa Hàng Về Kho Bảo Quản An Toàn Khi Chờ Kết Quả Kiểm Tra Chuyên Ngành (Thông Tư 121)",
        "steps": "Bước 1: Rà soát điều kiện kho bãi (phải có hợp đồng thuê/sở hữu kho hợp pháp, hệ thống camera, tường rào biệt lập); Bước 2: Nộp Giấy đăng ký kiểm tra chuyên ngành có mộc tiếp nhận của cơ quan kiểm tra kèm Đơn xin mang hàng về bảo quản (Mẫu số 09/BQHH); Bước 3: Đưa hàng về kho riêng, niêm phong và giữ nguyên trạng thái hàng hóa (tuyệt đối không mở bán); Bước 4: Đôn đốc cơ quan thử nghiệm trước ngày thứ 15 và ngày thứ 25; Bước 5: Nộp Giấy chứng nhận đạt chất lượng lên cổng thông tin Một cửa Quốc gia trước ngày thứ 30.",
        "formula": "Thời hạn tối đa: 30 ngày kể từ ngày đưa hàng về kho; Mức phạt chậm nộp kết quả: 20.000.000 - 40.000.000 VNĐ + Đình chỉ quyền mang hàng về bảo quản 12 tháng (Điều 17 NĐ 128/2020/NĐ-CP).",
        "top3_errors": "1. Tự ý xé niêm phong tiêu thụ hàng hóa trước khi có kết quả thử nghiệm; 2. Để quá hạn 30 ngày mà không chủ động nộp công văn gia hạn; 3. Không kiểm tra năng lực thử nghiệm của phòng lab trước khi chọn gửi mẫu.",
        "mnemonic": "'Mang hàng về kho - Giữ nguyên không mở - Đếm ngược từng ngày - Đúng ba mươi chớ lỡ' (Nhớ quy tắc đèn đỏ 30 ngày).",
        "comparison": "Để hàng tại cảng chờ kết quả (An toàn tuyệt đối nhưng tốn hàng chục triệu tiền DEM/DET) vs Đưa hàng về kho bảo quản (Tiết kiệm tiền lưu bãi nhưng gánh rủi ro pháp lý nếu quá hạn 30 ngày).",
        "persona": "Anh Minh (Chủ doanh nghiệp sản xuất/nhập khẩu) & Chị Hương (Phụ trách thủ tục hải quan).",
        "checklist": "[ ] Kho đạt chuẩn an toàn; [ ] Nộp mẫu 09/BQHH; [ ] Check tiến độ lab ngày 15; [ ] Xin gia hạn nếu chậm trước ngày 25; [ ] Nộp kết quả trước ngày 30.",
        "url": "https://en.wikipedia.org/wiki/Customs",
        "keywords": ["customs", "goods"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% Thông tư 121/2025/TT-BTC, Nghị định 128/2020/NĐ-CP và hướng dẫn thực tế tại daily_notes.md ngày 25/09/2026."
    },
    {
        "stt": 98,
        "week": "W39 — 25/09/2026",
        "title": "Phương Pháp 3 Lớp Tự Phân Định Nông Sản 'Sơ Chế Thông Thường' (Mã VK010 - Không Chịu VAT) vs 'Chế Biến Công Nghiệp' (VAT 5%-10%)",
        "steps": "Lớp 1: Kiểm tra mô tả quy trình sản xuất (Tách vỏ, phơi khô, sấy mát = Sơ chế; Sấy nhiệt độ cao, tẩm ướp, nướng chín = Chế biến); Lớp 2: Kiểm tra hình thái vật lý (Hàng nguyên củ/nguyên lá/nguyên hạt = Sơ chế; Đã xay nhuyễn, ép viên, cắt vụn vi mô = Có nguy cơ là chế biến); Lớp 3: Tra cứu Biểu thuế VAT và chú giải chi tiết chương 07, 08, 09, 12 để chọn đúng mã miễn giảm VK010 hoặc mã thuế suất 5%/10%.",
        "formula": "Số thuế GTGT phải nộp = Trị giá tính thuế NK x Thuế suất GTGT (0% nếu VK010, hoặc 5%, hoặc 10%); Mức phạt khai sai mã miễn thuế: 20% trên số thuế khai thiếu + 0.03%/ngày tiền chậm nộp.",
        "top3_errors": "1. Mặc định 'cứ nông sản là không chịu VAT'; 2. Không xin bản thuyết minh quy trình sản xuất từ nhà cung cấp nước ngoài; 3. Không gửi công văn xin xác định trước mã miễn giảm tới Cục Hải quan khi gặp mặt hàng giao thoa.",
        "mnemonic": "'Phơi sấy tự nhiên: VK010 không thuế - Nhiệt luyện băm nghiền: Mười phần trăm nộp ngay' (Sơ chế thì nhẹ thuế, công nghiệp chế biến thì gánh thuế).",
        "comparison": "Mã VK010 Không Chịu Thuế (Chỉ áp dụng hàng sơ chế đơn giản bán cho DN sản xuất) vs Mã Thuế Suất 5%-10% (Áp dụng cho nông sản đã qua chế biến hoặc bán thẳng cho người tiêu dùng cuối).",
        "persona": "Chị Hương (Kế toán thuế XNK) & Anh Minh (Chủ doanh nghiệp nông sản).",
        "checklist": "[ ] Lấy bản mô tả quy trình nhà sản xuất; [ ] Check chú giải chương HS; [ ] So sánh với Khoản 1 Điều 4 NĐ 181; [ ] Tham vấn trước Hải quan nếu sấy nhiệt; [ ] Khai mã thuế chính xác.",
        "url": "https://en.wikipedia.org/wiki/Value-added_tax",
        "keywords": ["value-added", "tax"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% Khoản 1 Điều 4 Nghị định 181/2025/NĐ-CP, Thông tư 219/2013/TT-BTC và tình huống mã HS 07129090 tại daily_notes.md ngày 25/09/2026."
    },
    {
        "stt": 99,
        "week": "W39 — 25/09/2026",
        "title": "Quy Trình 5 Bước Thẩm Định Hóa Đơn Đầu Vào Chống Dính 'Công Ty Ma' Và Kỹ Thuật Lập Bộ Hồ Sơ Chứng Từ Logistics Chứng Minh Giao Dịch Thật",
        "steps": "Bước 1: Tra cứu thời gian thực trạng thái mã số thuế của người bán trên Cổng tracuunnt.gdt.gov.vn (tránh Trạng thái 06); Bước 2: Đối chiếu địa chỉ trụ sở và lịch sử hoạt động, kiểm tra tài khoản ngân hàng thụ hưởng phải đúng tên pháp nhân xuất hóa đơn; Bước 3: Thu thập trọn bộ chứng từ logistics (Hợp đồng vận chuyển, Biên bản giao nhận hàng hóa có ký nhận của tài xế và thủ kho, Phiếu cân xe, Ảnh chụp biển số xe); Bước 4: Lưu trữ sao kê chuyển khoản ngân hàng UNC khớp đúng số tiền và nội dung thanh toán; Bước 5: Đóng gói và số hóa hồ sơ theo từng hóa đơn để sẵn sàng giải trình khi cơ quan thuế thanh tra.",
        "formula": "Chi phí bị loại trừ khi dính hóa đơn ma: 100% giá trị hóa đơn; Thuế TNDN bị truy thu: 20% x Giá trị hóa đơn; Thuế GTGT bị loại: 10% x Giá trị hóa đơn; Phạt hành vi trốn thuế: 1 - 3 lần số thuế trốn.",
        "top3_errors": "1. Chỉ có hóa đơn và ủy nhiệm chi ngân hàng mà không có chứng từ logistics chứng minh dòng hàng; 2. Thanh toán tiền mặt trên 20 triệu hoặc ủy nhiệm chi ghi nội dung chung chung; 3. Không lưu ảnh chụp thực tế giao nhận hàng hóa tại kho.",
        "mnemonic": "'Tiền đi ngân hàng - Hàng đi có xe - Kho có biên bản - Thuế nào ép phê' (Có hóa đơn phải có dòng hàng thực tế đi kèm chứng từ logistics).",
        "comparison": "Chỉ Có Hóa Đơn + Chuyển Khoản (Dễ bị cơ quan thuế bác bỏ chi phí nếu bên bán bỏ trốn) vs Bộ Hồ Sơ 4 Lớp Đầy Đủ Logistics (Được bảo vệ quyền lợi hợp pháp là bên mua ngay tình theo CV 5120/TCT-CS).",
        "persona": "Anh Minh (Chủ doanh nghiệp) & Chị Hương (Kế toán trưởng).",
        "checklist": "[ ] Check MST bên bán còn hoạt động; [ ] Hợp đồng kinh tế đầy đủ; [ ] Vận đơn/Biên bản giao hàng có ký nhận; [ ] UNC ngân hàng đúng tên; [ ] Phiếu xuất nhập kho có chữ ký.",
        "url": "https://en.wikipedia.org/wiki/Tax_compliance_software",
        "keywords": ["tax", "compliance"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% Công văn 5120/TCT-CS của Tổng cục Thuế và cẩm nang phòng tránh rủi ro hóa đơn công ty ma trên daily_notes.md ngày 25/09/2026."
    },
    {
        "stt": 100,
        "week": "W39 — 25/09/2026",
        "title": "Chiến Lược Hợp Pháp Hóa Doanh Thu Affiliate TMĐT Cho Chủ Hộ Kinh Doanh: Cách Đăng Ký Mã Ngành Tiếp Thị & Tối Ưu Thuế 7% So Với 10% Cá Nhân",
        "steps": "Bước 1: Rà soát giấy chứng nhận đăng ký hộ kinh doanh hiện tại; Bước 2: Nộp hồ sơ bổ sung mã ngành kinh doanh 'Dịch vụ quảng cáo, tiếp thị và môi giới thương mại' tại UBND quận/huyện; Bước 3: Cập nhật thông tin mã số thuế hộ kinh doanh lên hệ thống thanh toán của sàn TMĐT (Shopee, TikTok Shop, Lazada); Bước 4: Kê khai doanh thu hoa hồng Affiliate vào tờ khai thuế khoán của hộ kinh doanh định kỳ; Bước 5: Lưu trữ sao kê tài khoản ngân hàng kinh doanh khớp đúng với đối soát doanh thu sàn.",
        "formula": "Thuế suất Cá nhân kinh doanh không đăng ký HKD: 10% (5% GTGT + 5% TNCN); Thuế suất Hộ kinh doanh ngành dịch vụ tiếp thị: 7% (5% GTGT + 2% TNCN); Mức tiết kiệm thuế: 3% trên tổng doanh thu hoa hồng Affiliate.",
        "top3_errors": "1. Để sàn tự khấu trừ 10% thu nhập cá nhân mà không đưa vào hộ kinh doanh; 2. Tự ý gộp doanh thu Affiliate vào HKD khi chưa đăng ký bổ sung mã ngành nghề; 3. Dùng tài khoản cá nhân nhận tiền hoa hồng nhưng không kê khai quyết toán thuế hàng năm.",
        "mnemonic": "'Cá nhân chịu mười - Hộ kinh doanh bảy - Thêm ngành tiếp thị - Đúng luật không lo' (Đưa vào hộ kinh doanh vừa rẻ thuế vừa minh bạch tài chính).",
        "comparison": "Khai Thuế Dạng Cá Nhân Độc Lập (Chịu thuế suất 10%, dễ bị truy thu nếu không quyết toán) vs Khai Thuế Dạng Hộ Kinh Doanh (Chịu thuế suất 7%, chi phí hạch toán rõ ràng, an toàn pháp lý).",
        "persona": "Anh Minh (Chủ hộ kinh doanh online) & Chị Hương (Kế toán tư vấn).",
        "checklist": "[ ] Bổ sung mã ngành tiếp thị tại UBND; [ ] Đổi thông tin thuế trên sàn TMĐT; [ ] Nhận tiền qua tài khoản HKD; [ ] Kê khai thuế khoán 7%; [ ] Lưu bảng đối soát hoa hồng.",
        "url": "https://en.wikipedia.org/wiki/Affiliate_marketing",
        "keywords": ["affiliate", "marketing"],
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Đã đối soát 100% Thông tư 40/2021/TT-BTC, Luật Quản lý Thuế và hướng dẫn thực tế cho chủ shop Affiliate trên daily_notes.md ngày 25/09/2026."
    }
]
