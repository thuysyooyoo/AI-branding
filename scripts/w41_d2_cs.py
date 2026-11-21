# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG CASE STUDY (W41 DAY 2 — 06/10/2026)
Số lượng: 5 ý tưởng (STT 126 - 130)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, URL thật đã qua url_validator.py
"""

CASE_STUDY_W41_D2 = [
    {
        "stt": 126,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Hiểu Nhầm Về Điều Kiện FCA Và EXW Trong Xuất Nhập Khẩu Tại Chỗ: Doanh Nghiệp Phụ Trợ Bị Phạt Truy Thu Thuế Do Bỏ Quên 250 Triệu Đồng Cước Xe Nội Địa",
        "stakeholders": "Công ty Cổ phần Cơ khí Chính xác M. (Hưng Yên) vs Công ty Điện tử FDI (KCN Thăng Long) vs Chi cục Hải quan Bắc Hà Nội",
        "context": "Công ty M. ký hợp đồng bán 50.000 bộ vỏ kim loại cho khách hàng Nhật Bản, khách hàng chỉ định giao hàng cho nhà máy FDI tại KCN Thăng Long (Hà Nội). Hợp đồng ghi điều kiện FCA nhà máy Hưng Yên. Công ty M. thuê xe tải chở hàng sang giao và xuất hóa đơn trị giá 3.2 tỷ đồng (không thể hiện phí xe vì tự thỏa thuận ngầm trả hộ). Khi Chi cục Hải quan kiểm tra sau thông quan, Hải quan xác định theo Thông tư 60/2019/TT-BTC, khoản cước vận chuyển nội địa 250 triệu đồng phát sinh chưa được phân bổ vào trị giá tính thuế của tờ khai nhập khẩu tại chỗ bên FDI.",
        "damage": "Bên mua FDI bị bác trị giá khai báo, ấn định truy thu 25 triệu đồng thuế VAT, phạt vi phạm hành chính 20% (5 triệu đồng) và tính tiền chậm nộp 12 triệu đồng. Bên bán M. phải chịu trách nhiệm hoàn trả toàn bộ số tiền phạt và chi phí luật sư giải trình 45 triệu đồng để giữ quan hệ đối tác.",
        "root_cause": "Lỗi hiểu sai bản chất Incoterms trong XKTC: Ngộ nhận rằng giao FCA tại cổng nhà máy thì các chi phí vận tải sau đó đến kho nhận chỉ định không liên quan đến nghĩa vụ thuế hải quan của tờ khai tại chỗ.",
        "solution": "Soạn thảo hợp đồng XKTC chuẩn hóa: Thỏa thuận rõ điều kiện giao hàng là DAP tại kho bên nhận; tách bạch hoặc gộp hẳn chi phí vận chuyển nội địa vào đơn giá sản phẩm; có phụ lục trách nhiệm phân bổ chi phí rõ ràng cho khâu mở tờ khai hải quan 2 đầu.",
        "opposing_views": "Phe A (Bên bán): Cho rằng hợp đồng ngoại thương FCA thì trách nhiệm giao hàng chấm dứt tại cổng kho; Phe B (Hải quan): Khẳng định giao dịch XKTC bản chất là hàng hóa chuyển giao nội địa Việt Nam, mọi chi phí đưa hàng đến điểm giao chỉ định đều cấu thành trị giá hải quan.",
        "lesson": "Đừng bao giờ bê nguyên xi tư duy xuất khẩu viễn dương sang xuất nhập khẩu tại chỗ; mọi thỏa thuận Incoterms trong XKTC đều phải soi chiếu dưới lăng kính quy định trị giá hải quan Việt Nam.",
        "url": "https://en.wikipedia.org/wiki/Incoterms",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 06/10/2026 từ case study Stella Pham) & Thư Viện Pháp Luật (Thông tư 60/2019/TT-BTC).",
        "keywords": ["incoterms", "fca", "exw"]
    },
    {
        "stt": 127,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "C/O Form AI 3 Bên Thể Hiện Cả Hai Số Hóa Đơn Và Hai Trị Giá FOB: Doanh Nghiệp Hóa Chất Bị Giam Lô Hàng 8.5 Tỷ Suốt 45 Ngày Chờ Xác Minh AIFTA",
        "stakeholders": "Công ty TNHH Hóa Chất Nam Việt (Hải Phòng) vs Công ty Thương mại H. (Singapore) vs Nhà máy Sản xuất P. (Mumbai, Ấn Độ) vs Chi cục Hải quan Cảng Đình Vũ",
        "context": "Nam Việt nhập khẩu 10 container hạt phụ gia hóa chất trị giá 8.5 tỷ đồng từ Ấn Độ thông qua đối tác thương mại tại Singapore. Nhà xuất khẩu cung cấp C/O Form AI do Cơ quan Xuất khẩu Ấn Độ (EIC) cấp. Tuy nhiên, tại ô số 9, EIC ghi cả 2 trị giá FOB (giá xuất xưởng Ấn Độ 320.000 USD và giá bán Singapore 380.000 USD), ô số 10 ghi cả 2 số hóa đơn của Ấn Độ và Singapore. Công chức hải quan tiếp nhận nghi ngờ tính hợp lệ vì trên Cổng Một cửa chỉ thể hiện 1 số invoice và cho rằng C/O vi phạm quy cách Third-Party Invoicing.",
        "damage": "Lô hàng bị ách tắc tại cảng suốt 45 ngày; phát sinh phí lưu bãi (demurrage/detention) và tiền cắm điện cont lạnh lên tới 185 triệu đồng; doanh nghiệp phải nộp ký quỹ bảo lãnh thuế chênh lệch 680 triệu đồng để giải phóng hàng cho nhà máy sản xuất kịp tiến độ.",
        "root_cause": "Doanh nghiệp không kiểm tra đối soát quy cách phát hành C/O Form AI theo Thông tư 05/2018/TT-BCT trước khi tàu cập cảng; sự bất đồng bộ trong thói quen ghi chú của cơ quan cấp C/O Ấn Độ với quy chế kiểm tra của Hải quan Việt Nam.",
        "solution": "Thiết lập quy trình 'Pre-check C/O Draft 48h': Luôn yêu cầu shipper gửi bản scan C/O nháp trước khi cấp bản gốc chính thức; yêu cầu chỉ thể hiện duy nhất số hóa đơn của bên bán thứ ba tại ô số 10 và tích rõ ô số 13 Third-Party Invoicing theo đúng cẩm nang AIFTA.",
        "opposing_views": "Phe A (Doanh nghiệp): Chứng từ hoàn toàn minh bạch, việc ghi thêm thông tin nhà sản xuất chứng minh hàng không bị gian lận; Phe B (Hải quan): C/O phải tuân thủ nghiêm ngặt định dạng mẫu chuẩn đã thỏa thuận giữa các quốc gia thành viên để hệ thống đối soát tự động không báo lỗi.",
        "lesson": "Minh bạch quá mức không đúng quy chuẩn cũng là một cái bẫy kỹ thuật; trong thủ tục hải quan, đúng khuôn mẫu pháp lý mới là lá chắn an toàn nhất.",
        "url": "https://en.wikipedia.org/wiki/Certificate_of_origin",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 30/09/2026 từ case study C/O Form AI Bùi Văn Khẩn) & Thư Viện Pháp Luật (Thông tư 05/2018/TT-BCT).",
        "keywords": ["origin", "goods"]
    },
    {
        "stt": 128,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Hạt Nhựa DNCX Gửi Gia Công Biến Tính Tự Ý Đóng Nhãn 'Made In Vietnam': Lô Hàng 30 Tấn Bị Hải Quan Tạm Giữ Vì Vi Phạm Quy Tắc Xuất Xứ Không Đổi Mã HS",
        "stakeholders": "Doanh nghiệp Chế xuất K. (KCN Đồng Nai) vs Công ty Gia công Nhựa P. (Bình Dương) vs Chi cục Hải quan Quản lý Hàng Gia công",
        "context": "Doanh nghiệp K. nhập khẩu hạt nhựa ABS nguyên sinh từ Trung Quốc, sau đó gửi hợp đồng gia công sang Công ty P. để phối trộn hạt phụ gia chống cháy và sợi thủy tinh làm tăng độ cứng. Sau khi gia công xong, Công ty P. đóng bao bì mới và in nhãn 'Made in Vietnam'. Khi DNCX K. làm thủ tục tái xuất lô hàng 30 tấn sang thị trường xuất khẩu, cơ quan hải quan kiểm tra hồ sơ định mức và quy trình sản xuất, xác định mã HS trước và sau gia công vẫn giữ nguyên ở nhóm 3903, chi phí gia công tại Việt Nam chỉ chiếm 18% tổng giá trị sản phẩm (tiêu chí RVC < 40%).",
        "damage": "Toàn bộ lô hàng 30 tấn (trị giá 1.4 tỷ đồng) bị tạm giữ xử lý; doanh nghiệp bị phạt 65 triệu đồng về hành vi gian lận ghi nhãn mác xuất xứ theo Nghị định 128/2020/NĐ-CP; buộc phải bóc gỡ nhãn mác vi phạm, khai báo lại xuất xứ 'Không xác định' và nộp phí phạt chậm hợp đồng xuất khẩu 120 triệu đồng.",
        "root_cause": "Lẫn lộn giữa 'biến tính lý hóa' và 'chuyển đổi bản chất tạo xuất xứ' (Substantial Transformation); nhầm tưởng hễ có phát sinh công đoạn gia công đóng gói tại Việt Nam là được phép ghi xuất xứ Việt Nam.",
        "solution": "Rà soát Bộ quy tắc xuất xứ không ưu đãi tại Nghị định 31/2018/NĐ-CP trước khi thiết kế nhãn mác; nếu công đoạn chỉ là gia công đơn giản không đổi HS code và RVC không đạt 40%, bắt buộc giữ nguyên xuất xứ nguyên liệu gốc hoặc ghi 'Gia công tại Việt Nam từ nguyên liệu Trung Quốc'.",
        "opposing_views": "Phe A (Doanh nghiệp): Sản phẩm đã thay đổi hoàn toàn đặc tính kỹ thuật từ dễ cháy thành chống cháy cao cấp thì phải là sản phẩm mới; Phe B (Hải quan): Quy tắc xuất xứ quốc tế phân định bằng mã HS và tỷ lệ giá trị thặng dư định lượng, không dựa trên cảm tính kỹ thuật.",
        "lesson": "Gia công kỹ thuật cao chưa chắc đã tạo ra xuất xứ mới; xuất xứ hàng hóa là một bài toán toán học và pháp lý nghiêm ngặt, không thể làm theo trực giác.",
        "url": "https://en.wikipedia.org/wiki/Rules_of_origin",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 30/09/2026 từ case study Minh Thu) & Thư Viện Pháp Luật (Nghị định 31/2018/NĐ-CP, Điều 9).",
        "keywords": ["origin", "customs"]
    },
    {
        "stt": 129,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Container Chuyển Cảng Bị Đứt Seal Hãng Tàu Khai Báo Sai Số Niêm Phong: Doanh Nghiệp Thức Trắng Đêm Giải Trình Tại Cảng Tân Cảng Hiệp Phước",
        "stakeholders": "Doanh nghiệp Nhập khẩu Máy móc Thiết bị T. (TP.HCM) vs Hãng tàu viễn dương O. vs Cảng Cát Lái & Cảng Hiệp Phước vs Hải quan Giám sát",
        "context": "Lô hàng 2 container linh kiện thiết bị của Công ty T. cập cảng Cát Lái, đăng ký chuyển cảng đích về Tân Cảng Hiệp Phước bằng sà lan. Trong quá trình cẩu bốc xếp tại bãi Cát Lái, 1 container bị va quệt làm gãy seal chì của hãng tàu. Đơn vị cảng đã tự động bấm bổ sung seal nhựa định vị của cảng để tiếp tục vận chuyển. Tuy nhiên, nhân viên khai hải quan của công ty T. không biết sự việc, vẫn dùng số seal hãng tàu trên Bill of Lading để truyền tờ khai lấy container trên hệ thống Một cửa. Khi xe kéo ra cổng cảng Hiệp Phước, máy quét camera Barcode phát hiện số seal thực tế và số seal tờ khai không trùng khớp, lập tức còi báo động khóa barie.",
        "damage": "Xe đầu kéo bị giữ tại cổng bãi suốt 24 giờ, phát sinh phí neo xe 3.5 triệu đồng; container bị chuyển sang bãi kiểm hóa luồng Đỏ; chi phí nâng hạ bốc xếp kiểm hóa và tiền phạt hành chính khai sai chỉ tiêu giám sát hải quan tổng cộng 28 triệu đồng.",
        "root_cause": "Thiếu cơ chế đồng bộ thông tin thời gian thực giữa bộ phận hiện trường (OPS cảng), hãng tàu và nhân viên khai tờ khai văn phòng; không kiểm tra tình trạng vật lý của container trước khi truyền dữ liệu giám sát.",
        "solution": "Xây dựng Checklist kiểm tra seal tại cảng đích: Nhân viên hiện trường bắt buộc chụp ảnh đối soát số seal thực tế trên tay khóa container và gửi về phòng chứng từ trước khi truyền tờ khai lấy hàng; nếu có seal thay thế, bắt buộc yêu cầu Cảng cấp Biên bản ghi nhận sự cố đính kèm tờ khai OLA.",
        "opposing_views": "Phe A (Doanh nghiệp): Do lỗi khách quan của cảng làm gãy seal, hàng hóa bên trong còn nguyên vẹn 100%; Phe B (Hải quan): Quy định giám sát VASSCM quản lý bằng số seal định danh tự động, bất kỳ sai lệch nào đều tiềm ẩn nguy cơ thẩm lậu hàng lậu.",
        "lesson": "Trong kỷ nguyên hải quan số và camera AI soi tự động, một con số seal khai lệch một ký tự cũng đủ làm tê liệt cả một chuyến xe hàng bạc tỷ.",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Tai nạn XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 25/09/2026 từ case study Bùi Văn Khẩn) & Thư Viện Pháp Luật (Điều 51 Thông tư 38/2015/TT-BTC).",
        "keywords": ["seal", "container"]
    },
    {
        "stt": 130,
        "week": "W41 Day 2 — 06/10/2026",
        "title": "Bóc Phốt Hàng Về Chậm Hóa Ra Quên Dán Mã Khách Hàng: Eureka Phản Biện Bằng Dữ Liệu ERP Vạch Trần Lỗ Hổng Quy Trình Của Chủ Hàng",
        "stakeholders": "Khách hàng Phan Tran vs Ban Điều hành Eureka (Jasmine Nguyen) & Bạn Linh CSKH vs Hệ thống ERP Quản lý Đơn hàng",
        "context": "Khách hàng Phan Tran đăng bài lên diễn đàn 50.000 thành viên bóc phốt dịch vụ ủy thác của Eureka 'hàng về lâu lắc' và tìm đối tác thay thế. Chị Thúy (CEO Eureka) không đôi co chợ búa mà lập tức cam kết đền bù 100% nếu có thiệt hại, đồng thời trích xuất toàn bộ dữ liệu lịch sử trên hệ thống ERP nội bộ: 2 kiện hàng gần nhất bên bán gửi về kho Đông Hưng (Trung Quốc) hoàn toàn không ghi mã định danh khách hàng, khách cũng không báo mã vận đơn nội địa cho sale Linh. Hàng nằm ở kho trung chuyển suốt 5 ngày do hệ thống tự động khóa không cho xuất khẩu vì không có thông tin chủ hàng khai báo hải quan.",
        "damage": "Chi phí cơ hội và uy tín suýt bị tổn hại nếu Eureka không có hệ thống quản trị dữ liệu chặt chẽ; vụ việc kết thúc với việc khách hàng ruột vào làm chứng minh bạch, bảo vệ thương hiệu và mang lại thêm 3 hợp đồng ủy thác mới cho Eureka nhờ uy tín minh bạch.",
        "root_cause": "Khách hàng quen thói quen làm hàng tiểu ngạch trôi nổi (vứt hàng vào kho là xe chạy) mà không hiểu nguyên tắc vận hành chính ngạch bắt buộc phải có mã định danh và thông tin tờ khai xuất khẩu trước khi bốc xếp lên xe.",
        "solution": "Chuẩn hóa quy trình 3 chạm chính ngạch cho khách hàng: Chạm 1 - Tạo mã định danh khách hàng trên hệ thống trước khi đặt hàng; Chạm 2 - Yêu cầu nhà cung cấp Trung Quốc in tem dán mã khách hàng lên từng kiện; Chạm 3 - Bắn mã vận đơn nội địa lên App Eureka để kích hoạt quy trình thông quan tự động.",
        "opposing_views": "Phe A (Khách hàng): Cho rằng bên vận chuyển phải tự đi tìm hàng và chủ động giục khách; Phe B (Đơn vị ủy thác chính ngạch): Làm hàng chính ngạch chuẩn chỉ mỗi ngày xử lý hàng chục nghìn kiện, hệ thống ERP quản lý bằng dữ liệu chứ không thể quản lý bằng trí nhớ thủ công.",
        "lesson": "Dữ liệu số hóa và quy trình chuẩn mực là 'bùa hộ mệnh' mạnh nhất bảo vệ doanh nghiệp trước mọi khủng hoảng truyền thông; phục vụ tận tâm kết hợp với nguyên tắc đanh thép sẽ biến khủng hoảng thành cơ hội xây dựng niềm tin thương hiệu.",
        "url": "https://en.wikipedia.org/wiki/Supply_chain_management",
        "series": "Bóc phốt chiêu trò",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: daily_notes.md (ghi nhận ngày 24/09/2026 từ case study thực tế Jasmine Nguyen vs Phan Tran) & Hệ thống quản trị nội bộ Eureka ERP.",
        "keywords": ["supply", "chain"]
    }
]
