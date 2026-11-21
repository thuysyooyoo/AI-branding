# -*- coding: utf-8 -*-
"""
DATASET: KHO Ý TƯỞNG HỌC THỰC CHIẾN (W40 DAY 4 — 01/10/2026)
Số lượng: 5 ý tưởng (STT 121 - 125)
Tuân thủ: wf06_ideas_bank_update.md, 14 cột, quy trình step-by-step, công thức/link, URL thật đã qua url_validator.py
"""

HOC_THUC_CHIEN_W40_D4 = [
    {
        "stt": 121,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Hướng Dẫn 4 Bước Thiết Lập Hồ Sơ Xuất Nhập Khẩu Tại Chỗ Hợp Chuẩn Theo Công Văn 22411/CHQ-GSQL & Điều 47a Luật Hải Quan",
        "workflow": "Bước 1: Soát xét Hợp đồng ngoại thương 3 bên: Bắt buộc cài đặt điều khoản 'Chỉ định giao nhận hàng tại Việt Nam' (Designation Clause) ghi rõ tên, địa chỉ và mã số thuế của Doanh nghiệp nhận hàng tại Việt Nam; Bước 2: Thiết lập Bộ chứng từ xuất khẩu tại chỗ: Lập Commercial Invoice, Packing List và Thông báo giao hàng có xác nhận của bên mua nước ngoài; Bước 3: Mở tờ khai xuất khẩu tại chỗ và phối hợp mở tờ khai nhập khẩu đối ứng: Doanh nghiệp xuất mở tờ khai loại hình E62 hoặc B11, thông báo số tờ khai xuất cho Doanh nghiệp nhập để mở tờ khai nhập đối ứng (A11, E21, A42) trong thời hạn quy định 15 ngày; Bước 4: Lập Biên bản giao nhận hàng hóa nội địa: Cử đại diện hai bên ký biên bản giao nhận có đối chiếu biển số xe vận tải và số tờ khai hải quan để lưu trữ hồ sơ hoàn thuế.",
        "formula": "Thuế GTGT tiết kiệm hợp lệ nhờ XNK tại chỗ = Trị giá tính thuế x 10% thuế suất GTGT; Ví dụ lô hàng linh kiện trị giá 5 tỷ VNĐ, nếu làm đúng chuẩn XNK tại chỗ giúp doanh nghiệp được hưởng thuế suất VAT 0% cho hàng xuất khẩu và bên mua được khấu trừ thuế hợp pháp, không bị truy thu 500.000.000 VNĐ.",
        "top3_mistakes": "1. Không có văn bản chỉ định giao hàng hợp lệ từ thương nhân nước ngoài mà chỉ dựa vào thỏa thuận miệng; 2. Để quá thời hạn 15 ngày bên nhập không mở tờ khai đối ứng khiến tờ khai xuất bị hủy trên hệ thống; 3. Tên hàng hóa và số lượng trên tờ khai xuất và tờ khai nhập không trùng khớp 100%.",
        "mnemonic_tip": "XNK tại chỗ như 'chuyền bóng tay ba': Người ném bóng ở nước ngoài (thương nhân ngoại), người chuyền bóng ở Việt Nam (DN xuất), người nhận bóng cũng ở Việt Nam (DN nhập); nhưng đường bóng chỉ được trọng tài công nhận khi có 'hợp đồng chỉ định bằng văn bản'!",
        "comparison": "XNK tại chỗ (hàng hóa luân chuyển thực tế trong lãnh thổ VN theo chỉ định của đối tác ngoại, có thủ tục hải quan) vs Mua bán nội địa thông thường (giao dịch giữa hai pháp nhân VN không có yếu tố thương nhân nước ngoài, chỉ xuất hóa đơn GTGT nội địa).",
        "persona": "Anh Minh (Chủ doanh nghiệp phụ trợ, gia công cơ khí) & Chị Hương (Quản lý XNK các công ty DNCX/FDI).",
        "checklist": "☐ Hợp đồng có điều khoản chỉ định giao nhận tại VN\n☐ Tờ khai xuất và nhập mở đúng thời hạn 15 ngày\n☐ Tên hàng, mã HS, số lượng khớp 100% giữa 2 tờ khai\n☐ Biên bản bàn giao hàng hóa có chữ ký hai bên\n☐ Chứng từ thanh toán ngân hàng quốc tế hợp lệ.",
        "url": "https://en.wikipedia.org/wiki/Customs",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Công văn 22411/CHQ-GSQL ngày 30/09/2026 & Điều 47a Luật Hải quan) & daily_notes.md (ngày 01/10/2026).",
        "keywords": ["customs"]
    },
    {
        "stt": 122,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Cẩm Nang 4 Bước Rà Soát Hồ Sơ Nhập Khẩu Bao Bì & Dụng Cụ Tiếp Xúc Thực Phẩm Chuẩn Công Văn 22082/CHQ-GSQL",
        "workflow": "Bước 1: Tra cứu Danh mục kiểm tra chuyên ngành: Đối chiếu mã HS của sản phẩm (ví dụ túi vô trùng HS 3923.21.99, hộp nhựa định hình HS 3923.10.90) với Danh mục ban hành kèm Thông tư 15/2024/TT-BYT; Bước 2: Chuẩn bị Hồ sơ Tự công bố sản phẩm: Lập Bản tự công bố theo Mẫu số 01 Nghị định 15/2018/NĐ-CP kèm Phiếu kết quả kiểm nghiệm thôi nhiễm an toàn thực phẩm (Migration Test) còn hạn 12 tháng; Bước 3: Đăng ký Kiểm tra Nhà nước về ATTP trên Cổng Một cửa Quốc gia (NSW): Truyền bộ hồ sơ điện tử trước khi hàng về cảng tối thiểu 3 ngày làm việc; Bước 4: Làm việc với Cơ quan Kiểm tra được chỉ định lấy mẫu tại cảng: Cử nhân viên hiện trường phối hợp cán bộ kiểm dịch lấy mẫu thử nghiệm hoặc kiểm tra hồ sơ giảm kiểm tra để nhận Thông báo kết quả đạt yêu cầu nộp cho Hải quan thông quan.",
        "formula": "Chi phí kiểm tra ATTP nhập khẩu = Lệ phí kiểm tra nhà nước (theo biểu phí Bộ Y tế) + Phí thử nghiệm thôi nhiễm mẫu lab (1.500.000 - 3.500.000 VNĐ/chỉ tiêu); Tiết kiệm chi phí lưu kho bãi: Đăng ký NSW trước giúp giải phóng hàng trong 24h, tiết kiệm 15-30 triệu đồng tiền phí lưu container so với làm thủ công.",
        "top3_mistakes": "1. Nhầm tưởng đã nộp Bản tự công bố lên cơ quan quản lý là nghiễm nhiên được miễn kiểm tra ATTP khi mở tờ khai hải quan; 2. Sử dụng kết quả kiểm nghiệm quá hạn 12 tháng hoặc thử nghiệm không đủ các chỉ tiêu thôi nhiễm kim loại nặng, phẩm màu; 3. Tên sản phẩm trên Giấy tự công bố không khớp từng từ với tên trên Commercial Invoice và Bill of Lading.",
        "mnemonic_tip": "Bao bì tiếp xúc thực phẩm như 'chiếc bát ăn cơm': Muốn bán vào nhà hàng (thị trường), bạn phải tự cam kết bát sạch (Tự công bố), nhưng khi nhập khẩu qua cổng hải quan, nhân viên y tế bắt buộc phải kiểm tra chiếc bát đó có chứa chất độc hay không (Kiểm tra ATTP)!",
        "comparison": "Hàng rủi ro trung bình thực hiện Tự công bố (quản lý lưu hành nội địa) vs Kiểm tra nhà nước về ATTP nhập khẩu (quản lý rào cản kỹ thuật tại biên giới): Phải đáp ứng cả hai lớp rào cản thì lô hàng mới được thông quan hợp pháp.",
        "persona": "Chị Hương (Quản lý thu mua và XNK ngành thực phẩm đóng hộp, F&B, bao bì đóng gói) & Nhân viên chứng từ logistics.",
        "checklist": "☐ Tra cứu mã HS trên Danh mục Thông tư 15/2024/TT-BYT\n☐ Phiếu kết quả kiểm nghiệm thôi nhiễm còn hạn 12 tháng\n☐ Bản tự công bố đã đăng ký thành công trên cổng quản lý\n☐ Đăng ký hồ sơ kiểm tra ATTP trên Cổng Một cửa Quốc gia\n☐ Đối chiếu tên hàng trùng khớp 100% trên bộ chứng từ.",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 27/2026/TT-BYT & Thông tư 15/2024/TT-BYT) & daily_notes.md (Công văn 22082/CHQ-GSQL ngày 24/09/2026).",
        "keywords": ["food", "safety"]
    },
    {
        "stt": 123,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Quy Trình 3 Bước Quản Trị Rủi Ro 'Hạn 30 Ngày' Khi Đưa Hàng Về Kho Bảo Quản Theo Thông Tư 121/2025/TT-BTC",
        "workflow": "Bước 1: Thẩm định năng lực và tiến độ của Trung tâm Kiểm nghiệm: Ký thỏa thuận cam kết thời gian trả kết quả (SLA) bằng văn bản với phòng thử nghiệm độc lập (Quatest, Vinacontrol, SGS) trước khi nộp đơn xin mang hàng về bảo quản; Bước 2: Thiết lập Mốc cảnh báo 'T-10' (Ngày thứ 20): Cài đặt nhắc việc tự động trên hệ thống ERP tại ngày thứ 20 kể từ ngày đưa hàng về kho; nếu phòng lab chưa cấp chứng thư, nhân viên chuyên trách lập tức soạn thảo Công văn giải trình tiến độ; Bước 3: Nộp Công văn đề nghị gia hạn kiểm tra chuyên ngành: Gửi văn bản chính thức kèm Phiếu tiếp nhận mẫu thử nghiệm có dấu của phòng lab nộp lên Chi cục Hải quan quản lý trước ngày thứ 28, đảm bảo ghi nhận dấu vết pháp lý trên hệ thống Một cửa quốc gia.",
        "formula": "Tổn thất rủi ro nếu quá hạn 30 ngày = Mức phạt vi phạm hành chính (30.000.000 - 50.000.000 VNĐ theo Nghị định 128/2020) + Chi phí lưu kho cảng trong 12 tháng tiếp theo (ước tính 150 - 300 triệu VNĐ) + Rủi ro bị đưa vào luồng Đỏ kiểm hóa 100%.",
        "top3_mistakes": "1. Chờ đến ngày thứ 30 mới liên hệ phòng lab thì phát hiện mẫu thử nghiệm bị hỏng hoặc chưa có kết quả; 2. Tự ý mở bao bì đưa hàng vào sử dụng hoặc bán ra thị trường khi chưa có kết quả thông quan chính thức của Hải quan; 3. Không lưu giữ biên bản bàn giao nguyên niêm phong hải quan tại kho bảo quản.",
        "mnemonic_tip": "Mang hàng về kho bảo quản như 'mượn sách thư viện': Hạn trả sách là 30 ngày; nếu chưa đọc xong muốn mượn tiếp bắt buộc phải đến xin gia hạn trước khi hết hạn, để quá hạn một ngày là bị phạt tiền và tước thẻ thư viện cả năm!",
        "comparison": "Hàng lưu tại kho bãi cảng (tốn chi phí lưu container đắt đỏ nhưng an toàn về hạn chứng từ) vs Hàng đưa về kho riêng bảo quản (tiết kiệm chi phí kho bãi nhưng chịu áp lực thời hạn 30 ngày nộp kết quả rất khắt khe).",
        "persona": "Anh Dũng (Chủ doanh nghiệp cơ khí, thiết bị viễn thông) & Trưởng phòng Thủ tục Hải quan, Thủ kho doanh nghiệp.",
        "checklist": "☐ Ký cam kết tiến độ trả kết quả với phòng thử nghiệm\n☐ Cài chuông báo động kiểm soát mốc thời gian ngày thứ 20\n☐ Giữ nguyên trạng niêm phong hàng hóa tại kho riêng\n☐ Soạn công văn giải trình gia hạn nộp trước ngày thứ 28\n☐ Nộp chứng thư kiểm định ngay trong vòng 24h sau khi nhận kết quả.",
        "url": "https://en.wikipedia.org/wiki/Conformity_assessment",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Thông tư 121/2025/TT-BTC & Nghị định 128/2020/NĐ-CP) & daily_notes.md (vướng mắc gia hạn kiểm tra chất lượng ngày 25/09/2026).",
        "keywords": ["conformity", "assessment"]
    },
    {
        "stt": 124,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Kỹ Thuật 4 Bước Phân Biệt Nông Sản Sơ Chế Thông Thường (Mã VK010) vs Chế Biến Sâu Để Tối Ưu Thuế VAT Nhập Khẩu",
        "workflow": "Bước 1: Thu thập Hồ sơ Quy trình Sản xuất (Manufacturing Process Dossier): Yêu cầu nhà cung cấp nước ngoài cấp bản giải trình các công đoạn gia công từ nông sản tươi đến thành phẩm khô, ghi rõ mức nhiệt độ làm khô và bảng phân tích thành phần; Bước 2: Đối chiếu tiêu chí Khoản 1 Điều 4 Nghị định 181/2025/NĐ-CP: Xác định xem công đoạn làm khô có thuộc các phương pháp sơ chế thông thường (phơi nắng, sấy gió tự nhiên, sấy mát dưới 60°C) hay đã dùng nhiệt độ cao làm chín tế bào; Bước 3: Soát xét tiêu chí Phụ gia và Tẩm ướp: Kiểm tra kỹ xem hàng có bị tẩm ướp muối, đường, chất điều vị hay chất bảo quản nhân tạo hay không; Bước 4: Lựa chọn Mã Thuế suất & Soạn thảo Bản Giải trình Kỹ thuật: Nếu đạt tiêu chuẩn sơ chế thông thường thì khai mã VK010 (không chịu thuế GTGT); nếu có công đoạn chế biến nhiệt độ cao thì chủ động khai mã thuế VAT 5% hoặc 10% để tránh rủi ro phạt 20% khai sai sau thông quan.",
        "formula": "Chênh lệch thuế GTGT đầu vào = Trị giá tính thuế CIF x (5% hoặc 10% - 0%); Ví dụ lô hàng gia vị 2 tỷ VNĐ: Nếu áp mã VK010 hợp pháp, doanh nghiệp không phải bỏ ra 200.000.000 VNĐ tiền vốn nộp thuế GTGT tại khâu nhập khẩu, tối ưu hóa dòng tiền lưu động.",
        "top3_mistakes": "1. Thấy hàng có mã HS thuộc Chương 07, 08, 09 là tự ý tích mã VK010 mà không kiểm tra quy trình công nghệ; 2. Khai sai bản chất hàng hóa (như hành sấy nhiệt độ cao khai thành hành phơi khô thông thường); 3. Không có tài liệu chứng minh kỹ thuật từ nhà máy sản xuất khi cơ quan hải quan yêu cầu tham vấn giá và chính sách thuế.",
        "mnemonic_tip": "Phân biệt sơ chế vs chế biến giống như 'quả chuối': Chuối bóc vỏ phơi khô là sơ chế thông thường (thuế 0%); chuối đem chiên giòn rụm hoặc sấy nhiệt độ cao tẩm mật ong là chế biến công nghiệp (phải nộp thuế GTGT)!",
        "comparison": "Mã VK010 (Đối tượng không chịu thuế GTGT khâu nhập khẩu - dành cho nông sản thô sơ chế thông thường) vs Mã V05/V10 (Chịu thuế GTGT 5% hoặc 10% - dành cho nông sản thực phẩm đã qua chế biến công nghiệp sâu).",
        "persona": "Anh Minh (Chủ shop/doanh nghiệp nhập khẩu thực phẩm khô, gia vị) & Kế toán thuế XNK.",
        "checklist": "☐ Xin quy trình công nghệ có dấu của nhà máy nước ngoài\n☐ Kiểm tra nhiệt độ sấy không vượt ngưỡng làm biến tính sản phẩm\n☐ Cam kết sản phẩm không tẩm ướp phụ gia nhân tạo\n☐ Tra cứu danh mục mã miễn giảm thuế hải quan ban hành mới nhất\n☐ Chuẩn bị mẫu thực tế sẵn sàng giám định khi kiểm hóa.",
        "url": "https://en.wikipedia.org/wiki/Value-added_tax",
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Nghị định 181/2025/NĐ-CP & Thông tư 219/2013/TT-BTC) & daily_notes.md (vướng mắc mã thuế VK010 hành lá sấy khô ngày 25/09/2026).",
        "keywords": ["tax", "value-added"]
    },
    {
        "stt": 125,
        "week": "W40 Day 4 — 01/10/2026",
        "title": "Hướng Dẫn 4 Bước Tra Cứu & Khai Báo Chuẩn Bảng Mã Loại Hình Xuất Nhập Khẩu (A11, A12, A41, E21, E31) Theo Quyết Định 1357/QĐ-TCHQ",
        "workflow": "Bước 1: Xác định Tư cách Pháp nhân và Phạm vi Giấy phép Đầu tư: Kiểm tra doanh nghiệp là 100% vốn trong nước hay doanh nghiệp có vốn đầu tư nước ngoài (FDI), có đăng ký quyền phân phối/xuất nhập khẩu hay không; Bước 2: Xác định Mục đích Sử dụng Hàng hóa: Hàng dùng để tiêu dùng nội địa, hàng làm nguyên phụ liệu sản xuất gia công hay hàng để bán buôn thương mại; Bước 3: Tra cứu Bảng Mã Loại hình chuẩn Quyết định 1357/QĐ-TCHQ: Doanh nghiệp nội địa nhập tiêu dùng chọn A11; Doanh nghiệp FDI nhập khẩu theo quyền phân phối bắt buộc chọn A12; Nhập khẩu chuyển đổi mục đích sử dụng chọn A41; Hàng sản xuất xuất khẩu chọn E21/E31; Bước 4: Kiểm tra tính Tương thích với Hệ thống Báo cáo Quyết toán: Đảm bảo mã loại hình khai báo sẽ khớp nối tự động với phần mềm kế toán và số liệu nộp Báo cáo quyết toán định kỳ năm cho Hải quan.",
        "formula": "Rủi ro khai sai mã loại hình = Số tiền thuế bị ấn định do áp sai chính sách ưu đãi + Phạt vi phạm thủ tục hải quan (2.000.000 - 5.000.000 VNĐ/tờ khai) + Nguy cơ tắc nghẽn hồ sơ hoàn thuế VAT trị giá hàng tỷ đồng.",
        "top3_mistakes": "1. Doanh nghiệp FDI nhập hàng phân phối thương mại nhưng lại quen tay chọn mã A11 của doanh nghiệp nội địa; 2. Doanh nghiệp gia công nhập nguyên liệu nhưng khai nhầm mã kinh doanh tiêu dùng A11 khiến không làm được thanh khoản; 3. Tự ý chuyển đổi mục đích sử dụng hàng hóa (như bán hàng miễn thuế vào nội địa) mà không mở tờ khai chuyển mục đích mã A41.",
        "mnemonic_tip": "Mã loại hình như 'biển số xe': Xe cá nhân đeo biển trắng (A11), xe liên doanh nước ngoài đeo biển vàng/NN (A12), xe tải chuyên chở hàng sản xuất đeo biển riêng (E21); gắn sai biển số xe là bị cảnh sát giao thông hải quan 'thổi còi' xử phạt ngay lập tức!",
        "comparison": "Mã A11 (Nhập tiêu dùng thông thường của pháp nhân Việt Nam) vs Mã A12 (Nhập khẩu của doanh nghiệp FDI thực hiện quyền nhập khẩu/phân phối): Phân định ranh giới pháp lý cực kỳ nghiêm ngặt đối với luồng tiền và quyền hoàn thuế GTGT.",
        "persona": "Anh Dũng (Trưởng phòng Logistics doanh nghiệp FDI) & Nhân viên khai báo hải quan hiện trường, Kế toán trưởng.",
        "checklist": "☐ Kiểm tra Giấy phép đăng ký kinh doanh và quyền phân phối FDI\n☐ Xác định đúng dòng luân chuyển hàng hóa (thương mại vs sản xuất)\n☐ Tra cứu văn bản Quyết định 1357/QĐ-TCHQ đối chiếu từng dòng mã\n☐ Kiểm tra chéo với bộ phận kế toán thuế trước khi ký số tờ khai\n☐ Lưu trữ hồ sơ giải trình mục đích sử dụng hàng hóa sau thông quan.",
        "url": "https://en.wikipedia.org/wiki/Customs_valuation",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Nguồn đối soát chéo: Thư Viện Pháp Luật (Quyết định 1357/QĐ-TCHQ & Công văn 21939/HQKV2-NVHQ) & daily_notes.md (hướng dẫn mã loại hình XNK).",
        "keywords": ["customs"]
    }
]
