# -*- coding: utf-8 -*-
"""
DỮ LIỆU Ý TƯỞNG PILLAR: HỌC THỰC CHIẾN (HT) — TUẦN W40 (28/09/2026)
STT: 106 - 110
Tuân thủ đầy đủ 14 trường dữ liệu theo wf06_ideas_bank_update.md & skills/pillar-ht-tutorial/SKILL.md
"""

HOC_THUC_CHIEN_W40 = [
    {
        "stt": 106,
        "week": "W40 — 28/09/2026",
        "title": "Quy Trình 4 Bước Xử Lý Tờ Khai Khi Container Bị Thay Seal / Gãy Seal Trong Quá Trình Vận Chuyển Chuyển Cảng (Customs Seal Protocol)",
        "steps": "Bước 1: Kiểm tra thực tế tình trạng seal tại bãi cảng đích (chụp ảnh rõ nét số seal đang gắn trên cửa cont, vị trí bấm seal, tình trạng niêm phong);\nBước 2: Yêu cầu Bộ phận Điều độ/Giao nhận của cảng cung cấp 'Biên bản xác nhận gãy/thay seal' (nêu rõ số seal cũ bị gãy, lý do thay thế, số seal mới được bấm thay thế và thời điểm bấm);\nBước 3: Khai báo số seal mới vào chỉ tiêu 'Số seal hãng vận chuyển' hoặc mục ghi chú trên tờ khai hải quan điện tử v5, đồng thời upload file scan Biên bản thay seal của Cảng vào chỉ tiêu HYS (chứng từ đính kèm);\nBước 4: Làm việc trực tiếp với Hải quan giám sát cổng/bãi cảng đích để đối soát số seal thực tế và gỡ cảnh báo trên hệ thống VASSCM trước khi cho xe nâng cont lên rơ-moóc.",
        "formula": "Chi phí phát sinh nếu khai sai số seal = Tiền phạt vi phạm hành chính (khoản 1 Điều 7 Nghị định 128/2020/NĐ-CP: từ 1.000.000 - 5.000.000 VNĐ) + Phí nâng hạ kiểm hóa (5.000.000 - 10.000.000 VNĐ) + Phí lưu bãi DEM/DET (trung bình 1.500.000 - 2.500.000 VNĐ/ngày).",
        "top3_errors": "1. Thấy số seal trên cont khác B/L nhưng vẫn truyền tờ khai theo B/L vì nghĩ Hải quan chỉ quản lý theo giấy tờ;\n2. Tự ý cắt seal bị hỏng khi cont chưa có sự giám sát của Hải quan cổng;\n3. Không yêu cầu cảng đóng dấu xác nhận lên biên bản thay seal, khiến hồ sơ giải trình bị Hải quan từ chối.",
        "mnemonic": "Seal container như vân tay bảo mật điện thoại: Chỉ cần sai một nét là hệ thống lập tức khóa máy và báo động xâm nhập trái phép.",
        "comparison": "Truyền tờ khai theo số seal mới (có biên bản cảng) vs Truyền tờ khai theo số seal cũ trên B/L: Khai số seal mới được thông quan trôi chảy tại cổng giám sát VASSCM; khai số seal cũ sẽ bị hệ thống dừng khẩn cấp và chuyển kiểm hóa 100%.",
        "persona": "Anh Minh (Chủ doanh nghiệp XNK vừa và nhỏ), Chị Hương (Trưởng phòng Logistics/Chứng từ hiện trường)",
        "checklist": "☐ Chụp ảnh kiểm tra seal thực tế tại cảng đích\n☐ Lấy Biên bản thay seal có mộc của Cảng\n☐ Khai đúng số seal mới lên hệ thống v5\n☐ Upload biên bản sự cố vào mục HYS\n☐ Đối chiếu VASSCM trước khi kéo xe ra cổng",
        "url": "https://en.wikipedia.org/wiki/Security_seal",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Hướng dẫn nghiệp vụ đúc kết từ case thực tế chuyển cảng Cát Lái - Hiệp Phước tại daily_notes.md ngày 25/09/2026 & Thông tư 38/2015/TT-BTC.",
        "keywords": ["seal", "security"]
    },
    {
        "stt": 107,
        "week": "W40 — 28/09/2026",
        "title": "Cẩm Nang 4 Bước Tự Công Bố & Thông Quan Bao Bì, Dụng Cụ Tiếp Xúc Thực Phẩm Nhập Khẩu A-Z (Food Contact Material Import Guide)",
        "steps": "Bước 1: Gửi mẫu thử nghiệm thôi nhiễm tại phòng lab đạt chuẩn ISO 17025 được Bộ Y tế chỉ định (kiểm tra giới hạn chì, cadmi, formaldehyde, thôi nhiễm tổng số theo QCVN 12-1:2011/BYT đối với bao bì nhựa tổng hợp);\nBước 2: Lập và nộp Hồ sơ Bản Tự công bố sản phẩm (Mẫu số 01 Phụ lục I Nghị định 15/2018/NĐ-CP) lên Ban Quản lý ATTP hoặc Chi cục An toàn vệ sinh thực phẩm tỉnh/thành phố;\nBước 3: Đăng ký Kiểm tra nhà nước về ATTP trên Cổng thông tin một cửa quốc gia (NSW) đối với từng lô hàng nhập khẩu (đính kèm Bản Tự công bố đã đăng tải);\nBước 4: Mở tờ khai hải quan (khai mã số hồ sơ kiểm tra ATTP trên NSW), phối hợp lấy mẫu kiểm tra tại cảng hoặc xin đưa hàng về kho bảo quản và nộp Thông báo kết quả đạt yêu cầu ATTP để thông quan chính thức.",
        "formula": "Tổng chi phí thông quan 1 lô bao bì thực phẩm = Chi phí thử nghiệm mẫu phòng lab (3.000.000 - 6.000.000 VNĐ/mẫu) + Phí kiểm tra nhà nước ATTP (theo Thông tư 67/2021/TT-BTC: 300.000 VNĐ/lô) + Thuế nhập khẩu (mã HS 3923: 5% - 10%, có C/O Form E/VK thuế 0%) + Thuế VAT (8% hoặc 10%).",
        "top3_errors": "1. Chỉ làm Tự công bố sản phẩm nhưng không đăng ký kiểm tra nhà nước ATTP cho từng lô hàng về cảng;\n2. Sử dụng Phiếu kết quả kiểm nghiệm của nhà sản xuất nước ngoài chưa được phòng kiểm nghiệm chỉ định của Việt Nam công nhận;\n3. Không đối chiếu danh mục miễn trừ theo Điều 13 Nghị định 15/2018/NĐ-CP (bao bì nhập khẩu chỉ để sản xuất nội bộ xuất khẩu được miễn kiểm tra ATTP).",
        "mnemonic": "Tự công bố như 'Giấy khai sinh' của sản phẩm (làm 1 lần dùng mãi mãi), còn Kiểm tra nhà nước ATTP như 'Tấm vé lên máy bay' (mỗi chuyến hàng về bắt buộc phải mua vé riêng).",
        "comparison": "Bao bì nhập khẩu để bao gói hàng tiêu thụ nội địa vs Bao bì nhập khẩu để gia công/sản xuất hàng xuất khẩu: Tiêu thụ nội địa bắt buộc phải làm Tự công bố và Kiểm tra ATTP từng lô; sản xuất xuất khẩu được miễn kiểm tra nhà nước ATTP theo khoản 7 Điều 13 Nghị định 15/2018/NĐ-CP.",
        "persona": "Chị Hương (Trưởng phòng Thu mua nguyên vật liệu), Anh Dũng (Giám đốc nhà máy chế biến nông sản thực phẩm)",
        "checklist": "☐ Test thôi nhiễm mẫu theo QCVN 12-1/2/3/4\n☐ Nộp hồ sơ Tự công bố lên NSW/Sở Y tế\n☐ Đăng ký Kiểm tra nhà nước ATTP từng lô\n☐ Khai báo số tiếp nhận NSW lên tờ khai v5\n☐ Nộp kết quả đạt yêu cầu để thông quan",
        "url": "https://en.wikipedia.org/wiki/Food_safety",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Quy trình chi tiết đối soát Nghị định 15/2018/NĐ-CP, Thông tư 27/2026/TT-BYT, Công văn 22082/CHQ-GSQL & phản ánh thực tế tại daily_notes.md ngày 24/09/2026.",
        "keywords": ["food", "safety"]
    },
    {
        "stt": 108,
        "week": "W40 — 28/09/2026",
        "title": "Quy Trình 4 Bước Xin Đưa Hàng Về Bảo Quản & Xử Lý Quá Hạn 30 Ngày Kiểm Tra Chuyên Ngành An Toàn (Supervised Storage Compliance)",
        "steps": "Bước 1: Chuẩn bị hồ sơ đề nghị mang hàng về bảo quản (Văn bản đề nghị theo Mẫu 09/BQHH/GSQL, Giấy đăng ký kiểm tra chuyên ngành có xác nhận của cơ quan kiểm tra, Hợp đồng thuê kho hoặc Giấy chứng nhận quyền sử dụng kho đáp ứng điều kiện niêm phong bảo quản);\nBước 2: Nộp hồ sơ và ký cam kết giữ nguyên trạng hàng hóa tại kho (không được tự ý mở niêm phong, sử dụng hoặc tiêu thụ sản phẩm trước khi thông quan);\nBước 3: Đôn đốc phòng thử nghiệm và thiết lập mốc cảnh báo ngày thứ 20: Nếu phòng lab chưa có kết quả, lập tức yêu cầu phòng lab phát hành 'Biên bản xác nhận tiến độ thử nghiệm' nêu rõ phương pháp và ngày dự kiến trả kết quả;\nBước 4: Soạn Công văn báo cáo tiến độ và đề nghị kéo dài thời gian nộp kết quả gửi Chi cục Hải quan quản lý trước ngày thứ 27, đính kèm biên bản xác nhận của phòng lab để bảo lưu quyền bảo quản và tránh bị xử phạt.",
        "formula": "Thời hạn tối đa lưu giữ hàng bảo quản = 30 ngày (kể từ ngày bàn giao hàng tại cửa khẩu). Khung phạt vi phạm quá hạn nộp kết quả theo Điều 18 Nghị định 128/2020/NĐ-CP: Phạt tiền từ 10.000.000 - 30.000.000 VNĐ đối với hành vi nộp kết quả kiểm tra quá thời hạn quy định, kèm biện pháp tước quyền đưa hàng về kho trong 01 năm.",
        "top3_errors": "1. Để trôi qua ngày thứ 30 mới bắt đầu đi nộp công văn giải trình với cơ quan Hải quan;\n2. Tự ý bóc tem niêm phong hoặc xuất một phần hàng trong kho đi bán cho khách để kịp tiến độ hợp đồng;\n3. Khai báo địa chỉ kho bảo quản không đúng với thực tế hoặc kho không có tường rào, biển bảng nhận diện theo quy định.",
        "mnemonic": "Hàng đưa về bảo quản như 'Tù nhân được tại ngoại chờ xét xử': Được ở nhà riêng nhưng tuyệt đối không được đi lại tự do và phải trình diện đúng ngày giờ hẹn.",
        "comparison": "Để hàng tại cảng chờ kết quả test vs Đưa hàng về kho riêng bảo quản: Để hàng tại cảng an toàn pháp lý tuyệt đối nhưng tốn tiền DEM/DET (1 - 3 triệu/cont/ngày); Đưa về kho riêng tiết kiệm 80% chi phí lưu bãi nhưng chịu trách nhiệm pháp lý hình sự nếu để thất thoát hoặc trễ hạn 30 ngày.",
        "persona": "Anh Dũng (Giám đốc Sản xuất/Nhà máy), Chị Hương (Trưởng phòng Xuất Nhập Khẩu)",
        "checklist": "☐ Nộp mẫu 09/BQHH/GSQL kèm Giấy kiểm tra chuyên ngành\n☐ Chụp ảnh niêm phong kho bảo quản\n☐ Theo dõi tiến độ lab và cảnh báo ngày thứ 20\n☐ Xin văn bản tiến độ nếu test kéo dài > 25 ngày\n☐ Nộp kết quả đạt chuẩn lên NSW để giải phóng hàng",
        "url": "https://en.wikipedia.org/wiki/Bonded_warehouse",
        "series": "Series: Quy trình Hải quan",
        "status": "Chưa dùng",
        "note": "Hướng dẫn chuyên sâu bảo vệ quyền mang hàng về bảo quản cho DN XNK, đối chiếu thực tế daily_notes.md ngày 25/09/2026 & Điều 32 Thông tư 38/2015/TT-BTC.",
        "keywords": ["bonded", "warehouse"]
    },
    {
        "stt": 109,
        "week": "W40 — 28/09/2026",
        "title": "Bộ Tiêu Chí 3 Bước Phân Biệt Nông Sản Sơ Chế vs Chế Biến Để Áp Đúng Mã VAT & Mã Miễn Thuế VK010 (Agricultural VAT Classification Guide)",
        "steps": "Bước 1: Thu thập Bản mô tả quy trình chế biến sản phẩm (Flowchart & Spec Sheet) từ nhà cung cấp nước ngoài (xác định rõ các công đoạn: rửa, bóc vỏ, tách hạt, phơi nắng tự nhiên, sấy gió lạnh, sấy nhiệt, tẩm ướp);\nBước 2: Đối chiếu bộ tiêu chí phân loại của Thông tư 219/2013/TT-BTC và Thông tư 83/2014/TT-BTC: Nếu sản phẩm chỉ qua bóc vỏ, tách hạt, cắt lát, phơi khô tự nhiên hoặc sấy gió ở nhiệt độ môi trường -> Xếp vào 'Sơ chế thông thường' (thuộc đối tượng không chịu VAT ở khâu nhập khẩu, áp mã định danh VK010); Nếu sản phẩm đã qua sấy nhiệt công nghiệp cao, nướng, chiên, tẩm ướp gia vị -> Xếp vào 'Chế biến công nghiệp' (chịu thuế suất VAT 5% hoặc 10%);\nBước 3: Khai báo chuẩn xác trên phần mềm Hải quan điện tử: Nếu thuộc diện sơ chế thô, chọn mã thuế suất 'V' và nhập mã miễn giảm 'VK010', đồng thời đính kèm Quy trình sản xuất vào hồ sơ hải quan để sẵn sàng đối soát khi hải quan kiểm tra hồ sơ.",
        "formula": "Số tiền thuế VAT nộp ở khâu nhập khẩu = (Trị giá tính thuế NK + Thuế NK + Thuế TTĐB nếu có) x Thuế suất VAT (0%, 5%, 10%). Nếu đủ điều kiện áp mã VK010: Thuế VAT phải nộp = 0 VNĐ.",
        "top3_errors": "1. Cho rằng nông sản nguyên vẹn hình thức sấy khô thì mặc định không chịu thuế VAT;\n2. Không lưu trữ bản mô tả công nghệ sản xuất của nhà cung cấp nước ngoài để giải trình nguồn nhiệt sấy;\n3. Không khấu trừ được thuế VAT đầu vào do khai nhầm mã không chịu thuế thay vì khai chịu thuế 5% để được khấu trừ cho nhà máy chế biến.",
        "mnemonic": "Nông sản sơ chế như quả chuối bóc vỏ phơi nắng (tự nhiên nguyên chất); nông sản chế biến như miếng chuối sấy giòn trong lò (đã qua công nghiệp hóa).",
        "comparison": "Áp mã miễn thuế VK010 (Không chịu VAT) vs Áp mã thuế suất VAT 5%: Mã VK010 không phải nộp thuế ở khâu nhập khẩu nhưng không được khấu trừ thuế đầu vào của các chi phí liên quan; Áp thuế suất 5% phải nộp tiền thuế tại cảng nhưng được khấu trừ 100% thuế GTGT đầu vào khi bán hàng hoặc sản xuất.",
        "persona": "Anh Minh (Chủ doanh nghiệp nông sản/thực phẩm), Chị Hương (Kế toán trưởng/Phụ trách thuế XNK)",
        "checklist": "☐ Lấy Specification và Flowchart từ bên bán\n☐ Kiểm tra nhiệt độ và phương pháp làm khô\n☐ Đối chiếu Thông tư 219/2013/TT-BTC\n☐ Chọn đúng mã VAT (V kèm VK010 hay 5%/10%)\n☐ Lưu trữ hồ sơ kỹ thuật phục vụ hậu kiểm",
        "url": "https://en.wikipedia.org/wiki/Value-added_tax",
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Cẩm nang thực chiến chuyên sâu giải quyết vướng mắc mã thuế VK010 nông sản nhập khẩu từ daily_notes.md ngày 25/09/2026.",
        "keywords": ["value", "tax"]
    },
    {
        "stt": 110,
        "week": "W40 — 28/09/2026",
        "title": "Quy Trình 4 Bước Thiết Lập 'Bộ Hồ Sơ 4 Lớp' Chống Rủi Ro Hóa Đơn Doanh Nghiệp Bỏ Trốn & Bảo Vệ Chi Phí Thuế (Tax Audit Defense Kit)",
        "steps": "Bước 1: Thẩm định pháp lý đối tác trước khi giao dịch (Tra cứu trạng thái MST trên Cổng tracuunnt.gdt.gov.vn, kiểm tra ngày thành lập, vốn điều lệ, danh sách ngành nghề và xác thực địa điểm trụ sở/kho bãi thực tế);\nBước 2: Chuẩn hóa hợp đồng kinh tế và điều khoản thanh toán (Quy định rõ tài khoản ngân hàng thụ hưởng phải đứng tên chính doanh nghiệp bán, cấm thanh toán qua tài khoản cá nhân giám đốc; bổ sung điều khoản cam kết bồi thường 100% nếu hóa đơn bị cơ quan thuế từ chối);\nBước 3: Thu thập và lưu trữ đồng bộ 'Bộ chứng từ logistics vật chất' (Gồm: Biên bản bàn giao hàng hóa có chữ ký người giao nhận, Phiếu cân tải trọng xe, Phiếu xuất kho của bên bán kèm Phiếu nhập kho của bên mua, Ảnh chụp biển số xe và kiện hàng thực tế tại kho);\nBước 4: Định kỳ đối soát tự động danh sách hóa đơn đầu vào với hệ thống rủi ro hóa đơn của cơ quan thuế: Nếu phát hiện đối tác bị cảnh báo ngừng hoạt động, lập tức đóng gói toàn bộ hồ sơ 4 lớp và nộp văn bản giải trình thiện chí trước khi đoàn thanh tra thuế công bố quyết định.",
        "formula": "Rủi ro tài chính khi bị loại hóa đơn = Số thuế VAT bị truy thu (8% - 10% x Trị giá hóa đơn) + Số thuế TNDN bị truy thu (20% x Trị giá hóa đơn) + Tiền phạt khai sai (20% tổng số thuế thiếu) + Tiền chậm nộp (0.03%/ngày x Số ngày chậm nộp). Tổng chi phí mất thêm có thể lên tới 35% - 40% giá trị lô hàng!",
        "top3_errors": "1. Chỉ lưu trữ hóa đơn điện tử dạng PDF/XML và UNC ngân hàng mà không lưu chứng từ giao nhận vận tải;\n2. Thanh toán hóa đơn cho bên thứ ba hoặc tài khoản cá nhân thay vì tài khoản công ty đã đăng ký với thuế;\n3. Mua hàng của các đơn vị thành lập dưới 3 tháng có doanh thu đột biến hàng chục tỷ đồng nhưng không có kho hàng thực tế.",
        "mnemonic": "Hồ sơ thuế 4 lớp như bộ áo giáp hiệp sĩ: Hợp đồng là mũ bảo hiểm, Hóa đơn là khiên, Chuyển khoản ngân hàng là giáp ngực, và Chứng từ logistics thực tế là thanh kiếm hộ thân.",
        "comparison": "Doanh nghiệp chỉ có Hóa đơn + Chuyển khoản vs Doanh nghiệp có đủ Hồ sơ 4 lớp (kèm log logistics): Nhóm chỉ có Hóa đơn bị loại 100% chi phí và bị chuyển hồ sơ sang cơ quan công an điều tra nếu đối tác bỏ trốn; Nhóm có đủ Hồ sơ 4 lớp được công nhận chi phí hợp lý vì chứng minh được bản chất giao dịch mua hàng là có thật.",
        "persona": "Anh Minh (Giám đốc điều hành), Chị Hương (Kế toán trưởng & Trưởng phòng Mua hàng)",
        "checklist": "☐ Tra cứu trạng thái MST đối tác trên cổng GDT\n☐ Chuyển khoản đúng tài khoản công ty mở tại ngân hàng\n☐ Lưu phiếu cân, biên bản giao nhận có ký nhận lái xe\n☐ Lưu ảnh chụp kiện hàng và biển số xe tải tại kho\n☐ Rà soát định kỳ danh sách hóa đơn rủi ro",
        "url": "https://en.wikipedia.org/wiki/Invoice",
        "series": "Series: Tối ưu Thuế XNK",
        "status": "Chưa dùng",
        "note": "Hướng dẫn nghiệp vụ quản trị thuế thực chiến đúc kết từ daily_notes.md ngày 24-25/09/2026 & Thông tư 78/2021/TT-BTC.",
        "keywords": ["invoice"]
    }
]
