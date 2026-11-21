# -*- coding: utf-8 -*-
"""
DATASET: HỌC THỰC CHIẾN (HT) - BATCH W39 (STT 86 - 90)
Mốc thời gian: W39 — 22/09/2026
Tuân thủ: wf06_ideas_bank_update.md, 14 trường thông tin chuẩn xác
"""

HOC_THUC_CHIEN_W39 = [
    {
        "stt": 86,
        "week": "W39 — 22/09/2026",
        "title": "Quy Trình 4 Bước Đóng Mã Số Thuế Phụ (Đuôi 888) Cho Cá Nhân Kinh Doanh & Mẫu Văn Bản Giải Trình Giảm Phạt Chậm Nộp Tờ Khai Quá Hạn",
        "steps": "Bước 1: Rà soát & Lập tờ khai bổ sung: Đăng nhập thuedientu.gdt.gov.vn nộp toàn bộ các tờ khai trắng mẫu 01/CNKD còn thiếu của các quý quá hạn để hoàn tất nghĩa vụ báo cáo; Bước 2: Soạn văn bản giải trình tình tiết giảm nhẹ: Làm Đơn xin giảm nhẹ mức phạt theo Điều 9 Luật Xử lý vi phạm hành chính, viện dẫn hoàn cảnh cá nhân không phát sinh doanh thu thực tế, hành vi vi phạm lần đầu do thiếu hiểu biết; Bước 3: Nộp hồ sơ đóng MST: Nộp Văn bản thông báo chấm dứt hiệu lực mã số thuế (Mẫu 24/ĐK-TCT) kèm bảng kê doanh thu và đơn giải trình lên Đội Đăng ký thuế; Bước 4: Nộp tiền phạt & Nhận Thông báo hoàn tất đóng MST: Sau khi có quyết định xử phạt đã giảm nhẹ, nộp tiền phạt vào Kho bạc và nhận Thông báo chấm dứt hiệu lực MST (Mẫu 18/TB-ĐKT) trong vòng 03 ngày.",
        "formula": "Số tiền phạt tối ưu = Khung phạt tối thiểu x 50% (nếu có 02 tình tiết giảm nhẹ: tự nguyện khắc phục nộp tờ khai + vi phạm lần đầu không có động cơ trốn thuế).",
        "top3_errors": "1. Để mặc MST không đóng vì nghĩ không dùng thì tự hết hạn; 2. Nộp đơn đóng MST trước khi nộp bổ sung các tờ khai còn nợ; 3. Không làm văn bản giải trình xin tình tiết giảm nhẹ nên bị áp mức phạt kịch khung trung bình.",
        "mnemonic": "KHAI TRƯỚC - GIẢI TRÌNH SAU - NỘP PHẠT NHẸ - ĐÓNG AN TOÀN",
        "comparison": "Để mặc tài khoản: Bị phạt tích lũy ngày chậm nộp, khóa CCCD vĩnh viễn, không hoàn được thuế TNCN vs Chủ động đóng bài bản: Tốn 1 lần tiền phạt tối thiểu, hồ sơ thuế cá nhân sạch 100%.",
        "persona": "Dành cho cá nhân làm Affiliate Shopee/TikTok, Reviewer, Freelancer, người bán hàng online từng mở MST phụ nay muốn chuyển nghề hoặc về làm văn phòng.",
        "checklist": "[ ] Đã nộp đủ tờ khai 01/CNKD các quý; [ ] Có file Đơn giải trình giảm nhẹ ký tên; [ ] Tải Mẫu 24/ĐK-TCT; [ ] Biên lai nộp phạt Kho bạc; [ ] Quyết định đóng MST Mẫu 18/TB-ĐKT.",
        "url": "https://en.wikipedia.org/wiki/Taxpayer_Identification_Number",
        "keywords": ["taxpayer", "identification"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Cẩm nang đúc kết từ case study thực tế daily_notes.md ngày 22/09/2026 về xử lý phạt tiểu mục 4254 và đóng MST đuôi 888."
    },
    {
        "stt": 87,
        "week": "W39 — 22/09/2026",
        "title": "Cẩm Nang 5 Bước Nhập Khẩu Rượu Ngoại Chính Ngạch Từ A-Z: Điều Kiện Giấy Phép, Tự Công Bố ATTP, Tham Vấn Trị Giá Hải Quan & Kích Hoạt Tem Điện Tử",
        "steps": "Bước 1: Chuẩn bị Pháp nhân & Giấy phép: Xin Giấy phép phân phối rượu (Bộ Công Thương) hoặc Giấy phép bán buôn rượu (Sở Công Thương) có kho chứa đạt chuẩn PCCC và bảo quản nhiệt độ; Bước 2: Tự công bố An toàn thực phẩm: Yêu cầu nhà sản xuất gửi mẫu kiểm nghiệm, CFS và bản dịch nhãn gốc để làm hồ sơ Tự công bố trên Cổng thông tin ATTP; Bước 3: Đăng ký Kiểm tra chất lượng nhà nước: Đăng ký kiểm tra an toàn thực phẩm lô hàng nhập khẩu qua Cổng Một cửa Quốc gia trước khi tàu cập cảng; Bước 4: Mở tờ khai & Tham vấn giá tính thuế: Mở tờ khai hải quan, nộp hồ sơ chứng minh trị giá thanh toán ngân hàng (L/C, T/T) để vượt qua khâu tham vấn giá rượu nhập khẩu; Bước 5: Kích hoạt dán tem điện tử & Thông quan: Đăng ký mua tem rượu nhập khẩu điện tử của cơ quan thuế, tổ chức dán tem QR Code tại kho ngoại quan cửa khẩu trước khi đưa hàng ra lưu thông.",
        "formula": "Tổng thuế rượu = Thuế Nhập Khẩu (CIF x 50%) + Thuế Tiêu Thụ Đặc Biệt [(CIF + Thuế NK) x 65%] + Thuế GTGT [(CIF + Thuế NK + Thuế TTĐB) x 10%].",
        "top3_errors": "1. Nhập hàng trước khi có Giấy phép kinh doanh có điều kiện; 2. Thiếu chứng nhận CFS hoặc kết quả kiểm nghiệm kim loại nặng/methanol; 3. Bị bác bỏ trị giá khai báo do không chứng minh được quan hệ đặc biệt với nhà sản xuất khi mua giá thấp.",
        "mnemonic": "PHÉP CHUẨN - MẪU KIỂM - GIÁ CHẶT - TEM ĐỦ - THÔNG NHANH",
        "comparison": "Đi đường tiểu ngạch/xách tay: Dễ bị tịch thu toàn bộ, phạt tội buôn lậu và mất thương hiệu vs Đi chính ngạch 100%: Pháp lý vững chắc, xuất hóa đơn VAT cho khách sạn/nhà hàng lớn và phát triển bền vững.",
        "persona": "Dành cho chủ doanh nghiệp F&B, giám đốc công ty phân phối đồ uống, nhà nhập khẩu thực phẩm cao cấp và nhân viên chứng từ logistics.",
        "checklist": "[ ] Giấy phép phân phối rượu còn hạn; [ ] Bản Tự công bố ATTP đã duyệt; [ ] Giấy chứng nhận CFS hợp pháp hóa lãnh sự; [ ] Chứng từ thanh toán ngân hàng MT103; [ ] Hạn ngạch tem điện tử được cấp.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "keywords": ["food", "safety"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Giải pháp toàn diện từ vướng mắc thủ tục nhập khẩu rượu ngoại tại Hải Phòng/Nội Bài ghi nhận trong daily_notes.md ngày 22/09/2026."
    },
    {
        "stt": 88,
        "week": "W39 — 22/09/2026",
        "title": "Hướng Dẫn Tra Cứu & Lấy Đúng Mã Định Danh Khoản Nộp (ID Khoản Nộp) Trên eTax Cho Doanh Nghiệp Trạng Thái 06 Để Gạch Nợ Dứt Điểm",
        "steps": "Bước 1: Tra cứu Bảng nợ thuế chi tiết: Đăng nhập thuedientu.gdt.gov.vn bằng tài khoản chữ ký số, vào menu 'Tra cứu' -> 'Tra cứu nghĩa vụ thuế' -> Xem chi tiết từng tiểu mục nợ gốc và nợ phạt; Bước 2: Sinh Mã định danh khoản nộp (ID): Tích chọn chính xác dòng nợ thuế môn bài, thuế TNDN hoặc tiền chậm nộp cần thanh toán, bấm nút 'Lập giấy nộp tiền theo ID' để hệ thống tự động sinh chuỗi mã số ID khoản nộp duy nhất; Bước 3: Thực hiện thanh toán điện tử: Thanh toán trực tiếp qua cổng eTax liên kết ngân hàng thương mại (hoặc quét mã VietQR có mã hóa ID); Bước 4: Đối soát & Lấy Bảng kê mẫu 01/BK-NT: Sau 24-48 giờ, kiểm tra lại tài khoản eTax đảm bảo trạng thái khoản nợ đã chuyển sang 'Đã thanh toán (Nợ = 0)'; in Giấy nộp tiền đến Đội Quản lý Nợ để xin xác nhận hoàn thành nghĩa vụ tài chính.",
        "formula": "Khoản nợ gạch thành công = Đúng Mã Số Thuế + Đúng Mã Kho Bạc + Đúng ID Khoản Nộp Đích Danh (Tuyệt đối không chuyển khoản tự do).",
        "top3_errors": "1. Chuyển khoản ngân hàng thường rồi tự gõ tay nội dung mà không qua ID; 2. Nộp đủ tiền nợ gốc nhưng quên không tích chọn dòng nợ phạt chậm nộp 0.03%/ngày; 3. Không kiểm tra lại sau khi nộp để tiền bị treo ở tài khoản tạm thu.",
        "mnemonic": "TRA NỢ - CHỌN DÒNG - TẠO ID - QUÉT QR - CHỐT SỔ",
        "comparison": "Nộp tự do ghi tay: Tiền vào Kho bạc nhưng nợ trên app vẫn tăng vì sai mã tiểu mục vs Nộp theo Mã ID eTax: Tiền khớp tự động trong 5 phút, gạch nợ tức thì trên hệ thống TMS toàn quốc.",
        "persona": "Dành cho kế toán trưởng, giám đốc doanh nghiệp đang làm thủ tục khôi phục mã số thuế, giải thể công ty hoặc dọn dẹp nghĩa vụ thuế tồn đọng.",
        "checklist": "[ ] Chữ ký số USB Token đang hoạt động; [ ] Bảng kê nợ thuế in từ TMS; [ ] Mã ID khoản nộp đã sinh; [ ] Thông báo trích nợ ngân hàng thành công; [ ] Giấy xác nhận nợ bằng 0 từ Chi cục Thuế.",
        "url": "https://en.wikipedia.org/wiki/Electronic_billing",
        "keywords": ["electronic", "billing"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Đúc kết giải pháp nghiệp vụ từ câu hỏi của thành viên Kho Buôn Hân Hân trong daily_notes.md ngày 22/09/2026."
    },
    {
        "stt": 89,
        "week": "W39 — 22/09/2026",
        "title": "6 Bước Tự Khai Quyết Toán Thuế Hộ Cá Nhân Kinh Doanh Đa Nguồn Thu Nhập (Mẫu 01/TKN-CNKD) & Cách Lưu Giữ Chứng Từ Khấu Trừ Thuế Điện Tử",
        "steps": "Bước 1: Thu thập chứng từ khấu trừ: Gửi yêu cầu đến tất cả các đối tác chi trả thu nhập (Shopee, TikTok, Behome, Ahamove...) cấp Chứng từ khấu trừ thuế TNCN điện tử có mã của cơ quan thuế; Bước 2: Tổng hợp doanh thu thực tế: Lập bảng Excel đối soát tổng doanh thu phát sinh từng nguồn và số tiền thuế 7% (5% GTGT + 2% TNCN) đã bị đối tác khấu trừ giữ lại; Bước 3: Đăng nhập eTax cá nhân: Truy cập canhan.gdt.gov.vn hoặc app eTax Mobile bằng tài khoản định danh điện tử VNeID Mức 2; Bước 4: Lập Tờ khai 01/TKN-CNKD: Điền chỉ tiêu Doanh thu ngành nghề dịch vụ tương ứng; Bước 5: Kê khai số thuế đã nộp: Nhập chính xác số thuế đã được tổ chức khấu trừ nộp thay dựa trên số hiệu chứng từ khấu trừ ở Bước 1 vào mục 'Số thuế đã khấu trừ/nộp thừa'; Bước 6: Ký số & Nộp hồ sơ: Gửi tờ khai kèm bản PDF chứng từ khấu trừ lên cổng thuế, nhận Thông báo chấp nhận hồ sơ khai thuế điện tử.",
        "formula": "Số thuế phải nộp thêm = (Tổng doanh thu dịch vụ x 7%) - Tổng số thuế đã bị tổ chức chi trả khấu trừ tại nguồn (Nếu kết quả <= 0 thì số thuế phải nộp bằng 0).",
        "top3_errors": "1. Khai doanh thu bằng 0 vì nghĩ tiền đã bị trừ rồi thì không cần báo; 2. Không lấy chứng từ khấu trừ thuế điện tử từ đối tác nên không có căn cứ giảm trừ; 3. Khai gộp thu nhập từ tiền lương/tiền công vào bảng doanh thu cá nhân kinh doanh.",
        "mnemonic": "GOM CHỨNG TỪ - BẢNG ĐỐI SOÁT - KHAI ĐỦ DOANH THU - TRỪ THUẾ ĐÃ KHẤU TRỪ",
        "comparison": "Tự ý khai man 0đ: Hệ thống tự động báo đỏ sai lệch số liệu và bị phạt 20% tiền trốn thuế vs Khai trung thực và đối trừ: Nộp 0 đồng hợp pháp, hồ sơ thuế hoàn toàn minh bạch.",
        "persona": "Dành cho cá nhân kinh doanh tự do, tài xế công nghệ, người làm tiếp thị liên kết, KOLs/KOCs và cộng tác viên nhận hoa hồng dịch vụ.",
        "checklist": "[ ] File PDF Chứng từ khấu trừ thuế điện tử; [ ] Bảng sao kê dòng tiền ngân hàng; [ ] Tài khoản VNeID Mức 2 đã kích hoạt; [ ] Tờ khai 01/TKN-CNKD đã điền chỉ tiêu; [ ] Thông báo tiếp nhận hồ sơ Mẫu 01-1/TB-TĐT.",
        "url": "https://en.wikipedia.org/wiki/Tax_return",
        "keywords": ["tax", "return"],
        "series": "Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Hướng dẫn thực chiến chuẩn xác giải quyết nỗi đau kê khai thuế cá nhân đa nguồn trong daily_notes.md ngày 21/09/2026."
    },
    {
        "stt": 90,
        "week": "W39 — 22/09/2026",
        "title": "Quy Trình 5 Bước Mở Tờ Khai Khai Bổ Sung Sau Thông Quan (AMA) Qua Phần Mềm VNACCS Cho Kiện Hàng Bị Sót Không Bị Phạt Nặng",
        "steps": "Bước 1: Giữ nguyên hiện trạng lô hàng: Ngay khi phát hiện dôi dư kiện hàng hoặc sai sót trọng lượng, lập tức cho dừng xe tại bãi tập kết, tuyệt đối không đưa hàng qua cổng kiểm soát giám sát barie; Bước 2: Thu thập Thư giải trình sai sót (Explanation Letter): Yêu cầu người gửi hàng (Shipper nước ngoài) gửi công văn có đóng dấu xác nhận sự cố gửi kèm nhầm kiện hàng do sơ suất của bộ phận đóng gói tại xưởng; Bước 3: Khởi tạo nghiệp vụ AMA trên VNACCS: Mở phần mềm khai hải quan điện tử (ECUS5-VNACCS), chọn phân hệ 'Khai bổ sung sau thông quan (AMA)', nhập số tờ khai gốc cần sửa đổi; Bước 4: Nhập thông tin kiện hàng bổ sung: Khai bổ sung dòng hàng phụ tùng/kiện dôi dư, điều chỉnh tổng số kiện và trọng lượng Gross Weight khớp với thực tế cân xe; đính kèm bản scan Thư giải trình và Packing List sửa đổi; Bước 5: Truyền thông điệp & Nộp lệ phí: Truyền tờ khai AMA đến Chi cục hải quan quản lý, gặp công chức trực kiểm hóa để giải trình lý do khách quan và nhận phê duyệt thông quan trong ngày.",
        "formula": "Mức xử phạt tự nguyện = Phạt hành vi khai bổ sung chậm (0,5 - 2 triệu đồng) << Phạt buôn lậu/gian lận thương mại (20 - 100 triệu đồng + tịch thu tang vật nếu để hải quan tự phát hiện).",
        "top3_errors": "1. Cố tình giấu nhẹm kiện hàng thừa để xe đi qua cổng soi chiếu hải quan; 2. Làm tờ khai mới hoàn toàn thay vì làm nghiệp vụ sửa đổi AMA cho tờ khai gốc; 3. Không yêu cầu Shipper nước ngoài xác nhận bằng văn bản trước khi nộp giải trình.",
        "mnemonic": "DỪNG XE - XIN THƯ - MỞ AMA - KHAI BỔ SUNG - THÔNG QUAN AN TOÀN",
        "comparison": "Chủ động mở AMA trước soi chiếu: Được coi là hành vi tự nguyện khai báo, chỉ phạt thủ tục nhẹ vs Bị máy soi hải quan phát hiện: Bị lập biên bản vi phạm buôn lậu, phạt kịch khung và niêm phong hàng hóa.",
        "persona": "Dành cho nhân viên hiện trường (Ops), nhân viên chứng từ xuất nhập khẩu (Docs) và các đại lý dịch vụ hải quan tại các cửa khẩu biên giới và cảng biển.",
        "checklist": "[ ] Biên bản bàn giao hiện trường; [ ] Thư giải trình sai sót của Shipper đóng dấu; [ ] Commercial Invoice & Packing List sửa đổi; [ ] Tờ khai AMA đã truyền thành công có mã tiếp nhận; [ ] Biên lai nộp phạt thủ tục hành chính.",
        "url": "https://en.wikipedia.org/wiki/Single-window_system",
        "keywords": ["single", "window"],
        "series": "Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Quy trình chuẩn hóa xử lý tình huống thực tế kiện hàng nhỏ bỏ sót từ daily_notes.md ngày 19/09/2026."
    }
]
