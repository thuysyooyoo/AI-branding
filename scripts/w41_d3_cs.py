# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W41 DAY 3 — 08/10/2026)
Số lượng: 5 ý tưởng (STT 131 - 135)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

CASE_STUDY_W41_D3 = [
    {
        "stt": 131,
        "week": "W41 Day 3 — 08/10/2026",
        "title": "Doanh Nghiệp Phụ Trợ Bán Linh Kiện Vào Khu Chế Xuất Quên Mở Tờ Khai: Nhà Cung Cấp Bị Truy Thu 420 Triệu Đồng Thuế GTGT 10% Suốt 2 Năm",
        "stakeholders": "Công ty TNHH Cơ khí Chính xác V. (Bắc Ninh) vs Tập đoàn Điện tử FDI H. (KCX Bắc Thăng Long) vs Cục Thuế tỉnh Bắc Ninh",
        "context": "Công ty V. ký hợp đồng cung ứng đồ gá và linh kiện cơ khí chính xác cho Tập đoàn FDI H. nằm trong Khu Chế Xuất với tổng giá trị hơn 4.2 tỷ đồng trong 2 năm. Do nghĩ rằng bên mua nằm trong KCX nên Công ty V. xuất hóa đơn thuế suất GTGT 0% mà không mở tờ khai hải quan xuất khẩu tại chỗ (do bên mua FDI từ chối mở tờ khai nhập khẩu đối ứng để tránh thủ tục). Khi Cục Thuế thanh tra thuế định kỳ, toàn bộ hóa đơn VAT 0% không có tờ khai hải quan thông quan đối ứng bị loại bỏ điều kiện áp dụng thuế suất 0% theo Điều 75 Thông tư 38/2015/TT-BTC.",
        "damage": "Công ty V. bị ấn định truy thu 420 triệu đồng tiền thuế GTGT (tính lại theo thuế suất 10%), phạt vi phạm hành chính 20% do khai sai thuế (84 triệu đồng), và tính tiền chậm nộp 58 triệu đồng. Tổng thiệt hại tài chính lên tới 562 triệu đồng, xóa sạch toàn bộ lợi nhuận hợp đồng.",
        "root_cause": "Hiểu sai bản chất pháp lý của Khu phi thuế quan: Nhầm tưởng rằng chỉ cần địa chỉ giao hàng ghi 'Khu Chế Xuất' là mặc nhiên được xuất hóa đơn VAT 0%, bỏ qua điều kiện tiên quyết là phải có tờ khai hải quan xuất khẩu đã thông quan.",
        "solution": "Thiết lập quy trình kiểm soát bán hàng cho DNCX: Đưa điều khoản bắt buộc mở tờ khai hải quan 2 đầu (xuất khẩu tại chỗ B11 và nhập khẩu tại chỗ E11/E15) vào hợp đồng; chỉ phát hành hóa đơn điện tử GTGT 0% sau khi đã nhận được bản sao tờ khai thông quan của bên mua.",
        "opposing_views": "Phe A (Bên bán): Cho rằng hàng hóa đã thực tế chuyển giao vào khu chế xuất phục vụ sản xuất thì bản chất là hàng xuất khẩu; Phe B (Cơ quan thuế): Luật quy định tờ khai hải quan là chứng từ duy nhất chứng minh hàng hóa đã thực sự rời khỏi thị trường nội địa vào khu phi thuế quan.",
        "lesson": "Trong xuất nhập khẩu, không có chứng từ hải quan thì không có khái niệm xuất khẩu; sự thỏa hiệp bỏ qua thủ tục hải quan với khách hàng lớn luôn là cái bẫy tự sát về thuế.",
        "url": "https://en.wikipedia.org/wiki/Free-trade_zone",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 07/10/2026 từ case study Ung Dung Ô Mai) & Thư Viện Pháp Luật (Điều 75 Thông tư 38/2015/TT-BTC).",
        "keywords": ["customs", "zone"]
    },
    {
        "stt": 132,
        "week": "W41 Day 3 — 08/10/2026",
        "title": "Áp Nhầm Mã HS 'Đồng Hồ Đeo Tay' Cho Lô 2.000 Chiếc Smartwatch: Doanh Nghiệp Công Nghệ Bị Tạm Nộp Thừa 380 Triệu Đồng Thuế Nhập Khẩu",
        "stakeholders": "Công ty Cổ phần Công nghệ K. (Hà Nội) vs Chi cục Hải quan Cửa khẩu Sân bay Quốc tế Nội Bài",
        "context": "Công ty K. nhập khẩu lô hàng 2.000 chiếc đồng hồ thông minh thế hệ mới (màn hình AMOLED 1.47 inch, chip Bluetooth 6.0, định vị GPS, micro đàm thoại) trị giá 2.5 tỷ đồng từ Thâm Quyến. Do nhân viên chứng từ mới thấy tên hàng là 'Smartwatch' nên tự ý áp mã HS nhóm 91.02 (đồng hồ đeo tay, thuế suất nhập khẩu 15%). Khi làm việc với chuyên gia Eureka và đối chiếu Thông báo xác định trước mã số 22555/TB-CHQ của Cục Hải quan, công ty mới nhận ra bản chất kỹ thuật của thiết bị là máy truyền nhận dữ liệu thuộc nhóm 8517.62.59 với thuế nhập khẩu ưu đãi 0%.",
        "damage": "Doanh nghiệp bị nộp thừa 375 triệu đồng tiền thuế nhập khẩu ngay tại khâu thông quan; dòng vốn lưu động bị đóng băng suốt 4 tháng để làm hồ sơ giải trình kỹ thuật và thủ tục hoàn thuế nộp thừa theo Thông tư 38/2015/TT-BTC.",
        "root_cause": "Phân loại mã HS theo tên gọi thương mại bề ngoài thay vì dựa vào cấu tạo kỹ thuật, chức năng chính và các Chú giải bắt buộc của Danh mục hàng hóa xuất nhập khẩu Việt Nam.",
        "solution": "Áp dụng thủ tục 'Xác định trước mã số' (Advance Ruling) tại Tổng cục Hải quan trước khi nhập khẩu các mặt hàng công nghệ mới; phân tích kỹ Chú giải 3 Phần XVI để xác định chức năng chính là truyền dữ liệu qua Bluetooth.",
        "opposing_views": "Phe A (Nhân viên chứng từ): Thấy có quai đeo và hiển thị giờ thì phân loại vào chương Đồng hồ 91; Phe B (Chuyên gia Hải quan): Cấu tạo vi mạch, loa micro và truyền sóng Bluetooth là chức năng cốt lõi vượt trội so với máy đồng hồ cơ học.",
        "lesson": "Mã HS Code quyết định trực tiếp biên lợi nhuận của doanh nghiệp; hiểu sâu bản chất kỹ thuật của sản phẩm là kỹ năng tối thượng giúp tiết kiệm hàng trăm triệu tiền thuế hợp pháp.",
        "url": "https://en.wikipedia.org/wiki/Smartwatch",
        "series": "Giải cứu hàng hóa",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 07/10/2026 từ case study Cục Hải quan ban hành Thông báo số 22555/TB-CHQ xác định mã HS Huawei GT 6) & Thư Viện Pháp Luật.",
        "keywords": ["smartwatch", "wearable"]
    },
    {
        "stt": 133,
        "week": "W41 Day 3 — 08/10/2026",
        "title": "Quên Kê Khai 1 Hóa Đơn Thuế Nhà Thầu Khi Mua Dịch Vụ SaaS Nước Ngoài: Doanh Nghiệp Công Nghệ Bị Phạt Truy Thu Gần 90 Triệu Đồng Sau 3 Năm",
        "stakeholders": "Công ty Cổ phần Giải pháp Số A. (TP.HCM) vs Công ty Dịch vụ Cloud S. (Singapore) vs Cục Thuế TP. Hồ Chí Minh",
        "context": "Công ty A. thanh toán phí dịch vụ hạ tầng đám mây và bản quyền phần mềm 25.000 USD (hơn 620 triệu đồng) cho đối tác Singapore qua điện chuyển tiền quốc tế ngày 21/07. Kế toán đã lập tờ khai thuế nhà thầu cho hóa đơn thanh toán ngày 13/07 nhưng bỏ sót không kê khai hóa đơn ngày 21/07 do cùng ngày phát sinh với thời điểm lập tờ khai cũ. Ba năm sau khi Cục Thuế kiểm tra quyết toán thuế, cơ quan thuế đối soát sao kê ngân hàng và phát hiện khoản thanh toán không có tờ khai thuế nhà thầu đối ứng.",
        "damage": "Bị ấn định truy thu 62.5 triệu đồng tiền thuế FCT (5% GTGT và 5% TNDN), phạt vi phạm hành chính chậm nộp hồ sơ khai thuế 6.5 triệu đồng, và tính tiền chậm nộp 0.03%/ngày suốt 1.050 ngày (19.7 triệu đồng). Tổng số tiền phải nộp phạt là 88.7 triệu đồng.",
        "root_cause": "Thiếu quy trình kiểm soát chéo giữa bộ phận Kế toán thanh toán quốc tế và Kế toán thuế; ngộ nhận rằng các khoản thanh toán phần mềm quốc tế được miễn trừ hoặc gộp chung vào kỳ khai thuế tháng tiếp theo.",
        "solution": "Xây dựng Checklist 'Khóa sổ điện chuyển tiền MT103': Bất kỳ lệnh chuyển tiền quốc tế nào ra nước ngoài đều phải gắn liền với mã hồ sơ khai thuế FCT trong vòng 10 ngày; nếu phát hiện sót lọt phải lập ngay tờ khai bổ sung Mẫu 01/KHBS theo Thông tư 80/2021/TT-BTC.",
        "opposing_views": "Phe A (Doanh nghiệp): Cho rằng dịch vụ phần mềm không chịu thuế GTGT nên chỉ cần nộp thuế TNDN và có thể kê khai bù cuối năm; Phe B (Cơ quan thuế): Luật Quản lý thuế quy định thời hạn nộp hồ sơ từng lần là 10 ngày, vi phạm thời hạn là vi phạm pháp luật bất kể số thuế lớn hay nhỏ.",
        "lesson": "Dòng tiền ra nước ngoài luôn để lại dấu vết rõ ràng trên sao kê ngân hàng; trốn hoặc quên thuế nhà thầu chỉ là 'giấu đầu hở đuôi', để càng lâu tiền phạt chậm nộp càng phình to.",
        "url": "https://en.wikipedia.org/wiki/Withholding_tax",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 07/10/2026 từ case study Khoa Nguyen) & Thư Viện Pháp Luật (Thông tư 80/2021/TT-BTC, Nghị định 125/2020/NĐ-CP).",
        "keywords": ["withholding", "tax"]
    },
    {
        "stt": 134,
        "week": "W41 Day 3 — 08/10/2026",
        "title": "Chuyển Nhượng Vốn Góp Công Ty TNHH Bằng Mệnh Giá 1 Tỷ Trong Khi Giá Trị Sổ Sách 6 Tỷ: Cổ Đông Sáng Lập Bị Ấn Định Thuế TNCN Phạt 365 Triệu Đồng",
        "stakeholders": "Ông Trần Văn V. (Cổ đông sáng lập Công ty Vận tải H.) vs Cục Thuế tỉnh Đồng Nai",
        "context": "Ông V. sở hữu 40% phần vốn góp (vốn điều lệ danh nghĩa 1 tỷ đồng) trong công ty logistics. Khi rút vốn, ông V. ký hợp đồng chuyển nhượng cho người quen với đúng giá 1 tỷ đồng để kê khai không phát sinh chênh lệch lợi nhuận, nhằm không phải nộp thuế TNCN 20%. Tuy nhiên, tại thời điểm chuyển nhượng, Báo cáo tài chính cho thấy công ty có quỹ đầu tư phát triển và lợi nhuận chưa phân phối lũy kế lên tới 12.5 tỷ đồng (giá trị sổ sách của 40% vốn góp là gần 6 tỷ đồng). Cơ quan thuế bác bỏ giá hợp đồng và thực hiện quyền ấn định giá theo Điều 50 Luật Quản lý thuế.",
        "damage": "Cơ quan thuế ấn định giá chuyển nhượng thực tế là 5 tỷ đồng, truy thu thuế TNCN 20% trên khoản chênh lệch 4 tỷ đồng, số thuế phải nộp là 320 triệu đồng, cộng thêm tiền phạt khai man và tiền chậm nộp 45 triệu đồng. Tổng số tiền phải nộp phạt là 365 triệu đồng.",
        "root_cause": "Ngộ nhận rằng các bên có toàn quyền tự do thỏa thuận giá chuyển nhượng phần vốn góp công ty TNHH; không hiểu nguyên tắc ấn định thuế theo giá trị tài sản ròng và giá thị trường của cơ quan thuế.",
        "solution": "Thực hiện định giá phần vốn góp dựa trên Báo cáo tài chính đã được kiểm toán trước khi chuyển nhượng; trường hợp có chênh lệch phải chuẩn bị hồ sơ giải trình biến động tài sản hoặc điều chỉnh cổ tức trước khi ký hợp đồng chuyển nhượng vốn.",
        "opposing_views": "Phe A (Người chuyển nhượng): Hợp đồng dân sự là sự thỏa thuận tự nguyện giữa hai bên, nhà nước không thể ép bán giá cao; Phe B (Cơ quan thuế): Luật Quản lý thuế trao quyền ấn định giá khi giao dịch có dấu hiệu không phù hợp với giá trị thực tế nhằm mục đích trốn nghĩa vụ thuế.",
        "lesson": "Bán vốn kinh doanh không đơn giản như bán hàng hóa thông thường; giá trị sổ sách kế toán là chiếc gương soi phản chiếu nghĩa vụ thuế mà không thỏa thuận ngầm nào che giấu được.",
        "url": "https://en.wikipedia.org/wiki/Capital_gains_tax",
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 07/10/2026 từ case study Hoàng Thị Mai Anh) & Thư Viện Pháp Luật (Thông tư 111/2013/TT-BTC, Thông tư 78/2014/TT-BTC).",
        "keywords": ["capital", "gains", "tax"]
    },
    {
        "stt": 135,
        "week": "W41 Day 3 — 08/10/2026",
        "title": "Kê Khai Thuế Trực Tuyến Bị Treo Đêm Cuối Kỳ Quyết Toán: Kế Toán Trưởng Bị Phạt 15 Triệu Đồng Vì Quên Kỹ Năng Xuất File XML Dự Phòng Từ HTKK",
        "stakeholders": "Công ty TNHH Thương mại Dịch vụ M. (Hà Nội) vs Hệ thống Thuế điện tử (thuedientu.gdt.gov.vn) vs Chi cục Thuế Quận Cầu Giấy",
        "context": "Vào 23h ngày 31/03 (hạn chót nộp Báo cáo tài chính và Quyết toán thuế TNDN năm), kế toán công ty M. mới đăng nhập vào Cổng eTax để 'Kê khai trực tuyến'. Do hàng vạn doanh nghiệp cùng truy cập, trình duyệt bị đơ liên tục, plugin ký số Java báo lỗi cấu trúc khiến toàn bộ số liệu nhập trên web bị biến mất. Kế toán lúng túng mất 2 tiếng tìm cách khắc phục, đến 01h15 sáng hôm sau mới tìm được cách mở phần mềm HTKK xuất file XML nộp lên thì hệ thống đã ghi nhận nộp trễ hạn 1 ngày.",
        "damage": "Doanh nghiệp bị xử phạt vi phạm hành chính 15 triệu đồng do nộp hồ sơ khai thuế quá thời hạn từ 1 đến 5 ngày có tình tiết không có lý do bất khả kháng theo Nghị định 125/2020/NĐ-CP; bị hạ điểm tín nhiệm tuân thủ thuế doanh nghiệp.",
        "root_cause": "Phụ thuộc hoàn toàn vào tính năng kê khai online trực tiếp trên trình duyệt; không nắm vững quy trình thao tác chuẩn là làm tờ khai offline trên HTKK, xuất file XML và nộp tờ khai định dạng XML.",
        "solution": "Ban hành quy chế nộp báo cáo thuế trước hạn chót tối thiểu 48 giờ; chuẩn hóa 100% việc nhập liệu trên phần mềm HTKK offline, kết xuất file XML sẵn sàng trong máy tính và kiểm tra chữ ký số trước ngày cao điểm.",
        "opposing_views": "Phe A (Kế toán): Hệ thống mạng của cơ quan thuế bị nghẽn là lỗi khách quan của hạ tầng nhà nước; Phe B (Cơ quan thuế): Doanh nghiệp có cả kỳ tính thuế 90 ngày để chuẩn bị, việc dồn vào giờ chót không được coi là sự kiện bất khả kháng.",
        "lesson": "Trong công tác kế toán và thủ tục công quyền, nộp sớm là an toàn, nộp sát giờ là mạo hiểm; luôn chuẩn bị sẵn kịch bản nộp file XML offline để không bị động trước sự cố đường truyền.",
        "url": "https://en.wikipedia.org/wiki/Electronic_tax_filing",
        "series": "Đơn lẻ",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 07/10/2026 từ case study Milan Nguyễn) & Thư Viện Pháp Luật (Nghị định 125/2020/NĐ-CP).",
        "keywords": ["electronic", "tax", "filing"]
    }
]
