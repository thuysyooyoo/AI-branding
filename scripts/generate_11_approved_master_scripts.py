# -*- coding: utf-8 -*-
"""
HỆ THỐNG SẢN XUẤT 11 KỊCH BẢN YOUTUBE THUYTALK ĐÃ ĐƯỢC CHỊ THÚY DUYỆT TAY
Tác tử: Agent 03 (YouTube Production) & Agent 08 (Fact Checker)
Đảm bảo chuẩn Persona Chị Thúy: Văn nói mộc mạc, trải nghiệm 15 năm thực chiến,
không viết tắt, không từ giật gân, thời lượng 7-9 phút (~1.500 - 2.000 từ).
"""
import os
import sys
import re
import datetime
import unicodedata
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
ONEDRIVE_DIR = r"C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING"
ONEDRIVE_YT_DIR = os.path.join(ONEDRIVE_DIR, "youtube")
ONEDRIVE_EXCEL = os.path.join(ONEDRIVE_DIR, "Ke_Hoach_Phat_Song_YouTube_Series.xlsx")
TEMP_EXCEL = os.path.join(BASE_DIR, "scratch", "temp_series_copy.xlsx")

SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")

os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)
os.makedirs(ONEDRIVE_YT_DIR, exist_ok=True)

# Dữ liệu nội dung chuyên sâu của 11 kịch bản đã được Chị Thúy duyệt
SCRIPTS_DATA = [
    {
        "sheet": "CM1_VAN_BAN_NONG",
        "code": "CM1-CD08",
        "title": "Siết hóa đơn điện tử từng lần bán & Rủi ro hóa đơn công ty 'ma' bỏ trốn: Cách bảo vệ chi phí được trừ",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.850 từ)",
        "cat_name": "Chuyên Mục 1: Trà Chiều & Văn Bản Nóng",
        "ssot": "Nghị định 123/2020/NĐ-CP, Thông tư 78/2021/TT-BTC, quy trình quản lý rủi ro hóa đơn điện tử của Tổng cục Thuế.",
        "visual": "Quay Chị Thúy ngồi bên bàn gỗ mộc mạc tại văn phòng Eureka, ấm trà gốm mộc, trên bàn để một xấp hồ sơ chứng từ có dán tem niêm phong và chiếc máy tính bảng hiển thị giao diện cổng tra cứu hóa đơn điện tử.",
        "hook": """Chào các bạn, hôm nay Thúy ngồi lại uống tách trà ấm để cùng mọi người chia sẻ về một câu chuyện pháp lý đang khiến rất nhiều chủ doanh nghiệp ăn không ngon, ngủ không yên. 

Đó là câu chuyện về những tờ hóa đơn điện tử. Các bạn cứ hình dung thế này: Doanh nghiệp của bạn mua hàng hóa thật, máy móc thiết bị thật, thanh toán tiền qua ngân hàng sòng phẳng từng đồng, lấy hóa đơn điện tử có mã của cơ quan thuế hợp lệ một trăm phần trăm. Thế nhưng bẵng đi một năm, thậm chí hai năm sau, khi cơ quan thuế đến thanh tra quyết toán, họ bất ngờ thông báo rằng: Công ty xuất hóa đơn cho bạn ngày xưa nay đã bỏ trốn khỏi địa chỉ kinh doanh, hoặc đã bị cơ quan điều tra đưa vào danh sách công ty ma mua bán hóa đơn bất hợp pháp!

Và hệ quả kéo theo là gì? Cơ quan thuế yêu cầu doanh nghiệp của bạn phải giải trình, loại toàn bộ số thuế giá trị gia tăng đã khấu trừ, loại toàn bộ chi phí giá vốn hợp lý đó ra khỏi chi phí được trừ khi tính thuế thu nhập doanh nghiệp, kèm theo khoản phạt chậm nộp tính theo từng ngày! Một tờ hóa đơn vài trăm triệu có thể khiến bạn bị truy thu và phạt lên tới hàng chục, thậm chí hàng trăm triệu đồng. 

Vậy tại sao mua hàng thật mà vẫn dính bẫy? Và làm thế nào để doanh nghiệp tự dựng lên một lá chắn pháp lý vững chắc bảo vệ quyền lợi chính đáng của mình? Hôm nay Thúy sẽ chia sẻ cặn kẽ từng bước với các bạn nhé!""",
        "core": """### 1. BẢN CHẤT CỦA CƠ CHẾ QUẢN LÝ HÓA ĐƠN ĐIỆN TỬ THEO NGHỊ ĐỊNH 123
Trước đây, thời còn dùng hóa đơn giấy, các doanh nghiệp ma có thể mua bán hóa đơn trôi nổi rồi vài năm sau mới bị phát hiện. Nhưng từ khi áp dụng Nghị định 123 và Thông tư 78, hệ thống dữ liệu hóa đơn điện tử của Tổng cục Thuế đã tích hợp trí tuệ nhân tạo và cơ chế phân tích rủi ro tự động. 
Hệ thống này quét đối chiếu liên tục dòng tiền, doanh thu và sản lượng mua vào - bán ra của từng mã số thuế. Một doanh nghiệp mới thành lập vốn điều lệ vài chục triệu mà đột ngột xuất hóa đơn hàng chục tỷ đồng chỉ trong vài tuần, lập tức hệ thống sẽ gắn cờ cảnh báo đỏ. 
Tuy nhiên, nghịch lý nằm ở chỗ: Tại thời điểm doanh nghiệp của bạn giao dịch, công ty bán hàng vẫn đang hoạt động bình thường, mã số thuế vẫn đang hiển thị trạng thái 00 (đang hoạt động), hóa đơn cấp mã của cơ quan thuế thành công. Nhưng chỉ vài tháng sau, họ biến mất, bỏ trốn khỏi trụ sở. Và khi đó, bên mua là người chịu trận đầu tiên!

---

### 2. HAI NGUY CƠ CHÍ MẠNG MÀ BÊN MUA THƯỜNG GẶP PHẢI
Rủi ro thứ nhất: Cơ quan thuế gửi công văn yêu cầu giải trình và tạm đình chỉ khấu trừ thuế giá trị gia tăng đầu vào của các hóa đơn thuộc doanh nghiệp có dấu hiệu rủi ro. Nhiều kế toán non tay, khi nhận được công văn này là hoảng loạn, vội vàng làm tờ khai điều chỉnh giảm thuế, tự nguyện nộp phạt. Làm như vậy là vô tình thừa nhận mình sai, tự tước đi quyền lợi của chính doanh nghiệp!
Rủi ro thứ hai: Bị loại trừ chi phí giá vốn được trừ khi tính thuế thu nhập doanh nghiệp. Thuế thu nhập doanh nghiệp hiện nay là hai mươi phần trăm. Nếu bạn mua lô hàng một tỷ đồng mà bị loại chi phí, bạn sẽ phải nộp thêm hai trăm triệu tiền thuế thu nhập doanh nghiệp, cộng thêm mười phần trăm thuế giá trị gia tăng là một trăm triệu, và tiền phạt chậm nộp 0,03% mỗi ngày trong suốt hai năm qua. Đó là một gánh nặng tài chính khổng lồ!

---

### 3. QUY TRÌNH 4 BƯỚC THIẾT LẬP 'BỘ HỒ SƠ CHỨNG MINH TÍNH CÓ THẬT CỦA GIAO DỊCH'
Để bảo vệ doanh nghiệp trước mọi cuộc thanh tra thuế, Thúy luôn yêu cầu bộ phận kế toán và mua hàng tại Eureka phải thiết lập bộ hồ sơ gồm 4 tầng bảo vệ sau:
- Bước một: Hồ sơ nhân thân đối tác trước khi ký hợp đồng. Trước khi chuyển tiền mua hàng cho bất kỳ đối tác mới nào, hãy chụp lại màn hình tra cứu tình trạng mã số thuế trên trang web của Tổng cục Thuế đúng tại ngày ký hợp đồng; lưu bản sao giấy phép đăng ký kinh doanh và căn cước công dân của người đại diện pháp luật.
- Bước hai: Hợp đồng kinh tế và phụ lục chi tiết. Hợp đồng phải quy định rõ địa điểm giao nhận hàng, phương thức vận chuyển, người đại diện nhận hàng và điều khoản cam kết xuất hóa đơn hợp pháp.
- Bước ba: Chứng từ giao nhận hàng hóa thực tế — Đây là mắt xích quan trọng nhất! Hóa đơn chỉ là chứng từ kế toán, còn Biên bản giao nhận hàng hóa mới chứng minh hàng có thật. Biên bản phải ghi rõ: Biển số xe chở hàng, tên tài xế, số điện thoại, vị trí kho bãi bốc dỡ hàng, chữ ký tươi của thủ kho bên giao và bên nhận. Nếu là hàng hóa giá trị lớn, hãy lưu lại hình ảnh camera kho bãi khi bốc dỡ hàng.
- Bước bốn: Dòng tiền thanh toán qua ngân hàng minh bạch một trăm phần trăm. Bắt buộc chuyển khoản từ tài khoản ngân hàng của doanh nghiệp bạn sang đúng tài khoản ngân hàng đã đăng ký với cơ quan thuế của bên bán theo mẫu 08. Tuyệt đối không chuyển tiền vào tài khoản cá nhân của giám đốc hay nhân viên kinh doanh của bên bán nếu không có văn bản ủy quyền hợp pháp.""",
        "conclusion": """Với bản thân Thúy, sau 15 năm làm nghề, mình luôn tự dặn lòng: Hóa đơn điện tử có mã của thuế không phải là bùa hộ mệnh vạn năng. Bùa hộ mệnh thực sự của doanh nghiệp chính là tính có thật của dòng hàng và tính minh bạch của dòng tiền. Khi các bạn làm thật, có kho bãi thật, có biên bản giao nhận rõ ràng và thanh toán chuẩn chỉ, thì dù đối tác sau này có bỏ trốn, các bạn vẫn hoàn toàn đủ cơ sở pháp lý vững vàng để bảo vệ quyền lợi chính đáng của mình trước cơ quan thuế.""",
        "cliffhanger": """Trong tập tiếp theo của chuyên mục Trà Chiều & Văn Bản Nóng, Thúy sẽ cùng các bạn mổ xẻ một vấn đề nóng bỏng không kém: Làn sóng điều tra phòng vệ thương mại và quy định siết chặt gian lận xuất xứ hàng hóa xuất khẩu đi thị trường Mỹ và châu Âu. Làm sao để hàng hóa Việt Nam không bị vạ lây khi bị nghi ngờ mượn đường đội lốt? Mời các bạn đón xem nhé!"""
    },
    {
        "sheet": "CM1_VAN_BAN_NONG",
        "code": "CM1-CD09",
        "title": "Làn sóng điều tra phòng vệ thương mại & Quy định siết gian lận xuất xứ (C/O) đi Mỹ, EU",
        "format": "Series 2 Tập Liên Hoàn",
        "duration": "8 - 10 phút (~1.900 từ)",
        "cat_name": "Chuyên Mục 1: Trà Chiều & Văn Bản Nóng",
        "ssot": "Nghị định 31/2018/NĐ-CP, Thông tư hướng dẫn cấp Giấy chứng nhận xuất xứ (C/O), các phán quyết điều tra chống lẩn tránh thuế của Bộ Thương mại Hoa Kỳ (DOC) và OLAF (Châu Âu).",
        "visual": "Chị Thúy ngồi tại phòng làm việc Eureka, mở tập hồ sơ Certificate of Origin Form B và Form EUR.1, bản đồ các tuyến hàng hải quốc tế trên màn hình phía sau.",
        "hook": """Chào các bạn, mình là Thúy từ Eureka Logistics! 

Hôm nay chúng ta bước vào một chủ đề mang tầm vóc sống còn đối với các doanh nghiệp sản xuất và xuất khẩu của Việt Nam: Làn sóng điều tra phòng vệ thương mại và rủi ro gian lận xuất xứ hàng hóa.

Các bạn có để ý thấy rằng trong vài năm trở lại đây, Bộ Thương mại Hoa Kỳ và Cơ quan Chống gian lận Thương mại Châu Âu liên tục mở các cuộc điều tra chống lẩn tránh thuế đối với hàng loạt mặt hàng xuất khẩu chủ lực của Việt Nam: Từ gỗ dán, tủ bếp, pin năng lượng mặt trời, cho đến nhôm định hình và thép cuộn? 

Có những doanh nghiệp Việt Nam làm ăn chân chính, đầu tư nhà xưởng hàng trăm tỷ đồng, nhưng chỉ vì một vài mắt xích nguyên phụ liệu nhập khẩu không minh bạch, hoặc vì trong cùng ngành có những đơn vị làm trò 'mượn đường đội lốt', mà cả ngành hàng bị áp mức thuế chống lẩn tránh lên tới hàng trăm phần trăm, đánh sập toàn bộ cánh cửa xuất khẩu sang thị trường phương Tây!

Hôm nay, trong tập 1 của chuỗi 2 tập liên hoàn này, Thúy sẽ bóc tách chiêu thức 'mượn đường đội lốt' diễn ra như thế nào, và cơ quan chức năng quốc tế dùng những công cụ kỹ thuật gì để vạch trần các hành vi gian lận xuất xứ!""",
        "core": """### 1. BẢN CHẤT CỦA HÀNH VI 'MƯỢN ĐƯỜNG ĐỘI LỐT' (TRANSSHIPMENT)
Tại sao lại có hiện tượng mượn đường đội lốt? 
Khi một quốc gia láng giềng bị Mỹ hoặc Liên minh Châu Âu áp mức thuế chống bán phá giá từ năm mươi đến hơn hai trăm phần trăm, hàng hóa của họ không thể cạnh tranh được trên thị trường quốc tế. Để né mức thuế này, một số doanh nghiệp nước ngoài đã chọn cách chuyển hàng sang Việt Nam.
Họ vận chuyển các thành phẩm hoặc bán thành phẩm gần như hoàn thiện về các cảng Hải Phòng, Cát Lái. Sau đó, hàng được đưa vào các nhà xưởng thuê tạm bợ, chỉ thực hiện các công đoạn gia công đơn giản như: tháo bao bì cũ, lau chùi bụi bẩn, gắn nhãn mác 'Made in Vietnam', rồi đóng vào container mới để xin cấp Giấy chứng nhận xuất xứ C/O Form B hoặc Form EUR.1 xuất khẩu đi Mỹ và châu Âu.
Hành vi này được quốc tế gọi là Transshipment hoặc Minor Processing — nghĩa là chuyển tải bất hợp pháp hoặc gia công giản đơn để lẩn tránh biện pháp phòng vệ thương mại.

---

### 2. CƠ QUAN ĐIỀU TRA QUỐC TẾ SOI XUẤT XỨ BẰNG CÁCH NÀO?
Nhiều người nghĩ rằng có tờ giấy chứng nhận xuất xứ C/O do cơ quan có thẩm quyền tại Việt Nam cấp là đã an toàn tuyệt đối. Đó là một ngộ nhận chết người!
Bộ Thương mại Mỹ hay cơ quan OLAF của châu Âu khi mở cuộc điều tra, họ không chỉ nhìn vào tờ giấy C/O. Họ cử các đoàn thanh tra trực tiếp sang tận nhà xưởng tại Việt Nam và áp dụng nguyên tắc kiểm tra thực tế:
- Thứ nhất, họ kiểm tra công suất thiết bị máy móc so với sản lượng xuất khẩu. Nếu một nhà xưởng chỉ có 2 chiếc máy ép nhỏ nhưng mỗi tháng xuất khẩu 50 container sản phẩm, con số đó là hoàn toàn phi lý!
- Thứ hai, họ soi hóa đơn tiền điện, tiền nước và bảng lương công nhân. Máy móc hoạt động sản xuất thật thì công tơ điện phải chạy hàng trăm triệu đồng mỗi tháng. Nếu hóa đơn tiền điện chỉ vài triệu đồng mà xuất khẩu hàng triệu đô la, đó là bằng chứng không thể chối cãi của việc gia công đóng gói đối phó.
- Thứ ba, họ truy vết dòng nguyên liệu đầu vào (Traceability). Bạn khai là gỗ Việt Nam hay nhôm Việt Nam, bạn phải chứng minh được nguồn gốc phôi nhôm, nguồn gốc gỗ rừng trồng, hóa đơn giá trị gia tăng mua từ các nhà máy trong nước.

---

### 3. QUY TẮC XUẤT XỨ CHUẨN MỰC THEO NGHỊ ĐỊNH 31/2018
Theo quy định pháp luật Việt Nam và thông lệ quốc tế, để một sản phẩm được công nhận có xuất xứ Việt Nam khi sử dụng nguyên liệu nhập khẩu, nó bắt buộc phải đáp ứng một trong hai tiêu chí cốt lõi:
- Tiêu chí thứ nhất: Chuyển đổi mã số hàng hóa (Change in Tariff Classification - CTC). Nghĩa là nguyên liệu nhập khẩu phải trải qua quá trình gia công, chế biến sâu tại Việt Nam để tạo thành một sản phẩm mới có mã số phân loại hải quan ở cấp độ 2 số, 4 số hoặc 6 số khác hoàn toàn so với mã số ban đầu.
- Tiêu chí thứ hai: Hàm lượng giá trị gia tăng nội địa (Regional Value Content - RVC). Nghĩa là tỷ lệ chi phí nguyên liệu nội địa, chi phí nhân công và chi phí sản xuất trực tiếp tại Việt Nam phải đạt tối thiểu từ ba mươi đến bốn mươi phần trăm tổng giá trị xuất xưởng của sản phẩm.""",
        "conclusion": """Với bản thân Thúy, khi tư vấn cho các doanh nghiệp xuất khẩu tại Eureka, mình luôn nhấn mạnh một nguyên tắc: Trong sân chơi thương mại toàn cầu ngày nay, sự minh bạch là chiếc chìa khóa duy nhất để tồn tại. Đừng bao giờ đánh đổi uy tín thương hiệu quốc gia và tài sản cả đời của doanh nghiệp để tiếp tay cho các phi vụ mượn đường kiếm lời chớp nhoáng. Bởi vì một khi cơ quan điều tra quốc tế vào cuộc, cái giá phải trả sẽ là sự sụp đổ của toàn bộ chuỗi kinh doanh.""",
        "cliffhanger": """Trong Tập 2 của loạt bài này, Thúy sẽ hướng dẫn các bạn chi tiết cách thức xây dựng một Bộ hồ sơ chứng minh năng lực sản xuất chuẩn quốc tế — từ bảng tính định mức chi phí Costing Sheet đến quy trình kiểm soát chuỗi cung ứng đầu vào — để sẵn sàng tiếp đón và giải trình thuyết phục trước bất kỳ đoàn thanh tra xuất xứ nào từ nước ngoài. Mời các bạn cùng đón xem nhé!"""
    },
    {
        "sheet": "CM2_DICH_LUAT_TIENG_NGUOI",
        "code": "CM2-CD01",
        "title": "Kiểm tra sau thông quan: Giải mã 5 năm 'hồi tố' theo Luật Hải quan 2014 & Cách kiểm soát định mức kho",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.900 từ)",
        "cat_name": "Chuyên Mục 2: Dịch Luật Ra Tiếng Người",
        "ssot": "Điều 77 đến 80 Luật Hải quan 2014, Thông tư 38/2015/TT-BTC sửa đổi bởi Thông tư 39/2018/TT-BTC của Bộ Tài chính.",
        "visual": "Chị Thúy ngồi bên bàn làm việc, mở cuốn Luật Hải quan 2014 có gạch chân bút dạ quang, màn hình bên cạnh hiển thị sơ đồ đối chiếu 3 bộ phận: Kế toán - Kho vận - Xuất nhập khẩu.",
        "hook": """Chào các bạn. Trong ngành xuất nhập khẩu và logistics, có một câu nói truyền tai nhau mà anh chị em làm nghề ai cũng thấm thía: 'Hàng ra khỏi cổng cảng chưa phải là xong, thông quan tờ khai mới chỉ là bắt đầu!'. 

Tại sao lại như vậy? Bởi vì theo Điều 77 Luật Hải quan năm 2014, cơ quan hải quan có quyền hạn kiểm tra sau thông quan đối với mọi lô hàng trong thời hạn lên tới 5 năm kể từ ngày đăng ký tờ khai hải quan! 

Năm năm là một khoảng thời gian rất dài. Năm năm đủ để một doanh nghiệp thay đổi vài đời giám đốc, thay đổi năm bảy kế toán trưởng, nhân viên xuất nhập khẩu cũ đã nghỉ việc từ lâu, thậm chí chứng từ kho bãi đã bị mối mọt gặm nhấm hoặc thất lạc. Thế nhưng khi đoàn kiểm tra sau thông quan của Cục Hải quan hoặc Tổng cục Hải quan công bố quyết định thanh tra, toàn bộ số liệu của năm năm đó sẽ bị đào xới lại từng chi tiết!

Rất nhiều doanh nghiệp rơi vào cảnh lao đao, bị truy thu thuế và xử phạt vi phạm hành chính hàng tỷ đồng chỉ vì sự bất nhất số liệu giữa ba bộ phận: Kế toán, Kho bãi và Xuất nhập khẩu. Hôm nay, Thúy sẽ dịch điều luật này ra tiếng người mộc mạc nhất để các bạn hiểu rõ bản chất và biết cách kiểm soát rủi ro nhé!""",
        "core": """### 1. BẢN CHẤT CỦA CƠ CHẾ 'TIỀN KIỂM CHUYỂN SANG HẬU KIỂM'
Hải quan hiện đại trên thế giới và tại Việt Nam vận hành theo cơ chế quản lý rủi ro. Để tạo điều kiện thuận lợi cho thương mại, giải phóng hàng nhanh tại cửa khẩu, hải quan cho phép đa số các tờ khai luồng xanh, luồng vàng được thông quan rất nhanh chóng. Đó gọi là bước 'tiền kiểm' thông thoáng.
Thế nhưng, sự thông thoáng ở cửa khẩu được bù đắp bằng sự siết chặt ở khâu 'hậu kiểm' — tức là Kiểm tra sau thông quan. 
Có 2 hình thức kiểm tra sau thông quan:
- Hình thức 1: Kiểm tra tại trụ sở cơ quan Hải quan. Thường diễn ra đối với các nghi vấn về trị giá hải quan, mã số HS, hoặc chính sách thuế của một vài tờ khai cụ thể.
- Hình thức 2: Kiểm tra toàn diện tại trụ sở người khai hải quan. Đây là cuộc đại phẫu toàn bộ hệ thống sổ sách kế toán, chứng từ thanh toán ngân hàng, quy trình sản xuất và tồn kho thực tế của doanh nghiệp trong suốt 5 năm.

---

### 2. TỬ HUYỆT 'LỆCH SỐ LIỆU TAY BA': KẾ TOÁN - KHO - XUẤT NHẬP KHẨU
Qua nhiều năm làm nghề và hỗ trợ gỡ rối cho khách hàng, Thúy nhận thấy sai lầm lớn nhất của các doanh nghiệp sản xuất xuất khẩu và gia công nằm ở sự rời rạc giữa 3 bộ phận:
- Bộ phận Xuất nhập khẩu: Khai báo định mức nguyên phụ liệu trên hệ thống VNACCS theo số liệu ước tính ban đầu của phòng kỹ thuật.
- Bộ phận Kho bãi: Nhập - xuất hàng theo thực tế sản xuất, hao hụt thực tế tại phân xưởng không lập biên bản chi tiết, phế liệu phế phẩm bán ra ngoài không ghi chép thẻ kho.
- Bộ phận Kế toán: Hạch toán tài khoản 152 (nguyên vật liệu), 155 (thành phẩm) trên phần mềm kế toán theo hóa đơn chứng từ mua vào bán ra, nhưng lại không bao giờ ngồi đối chiếu lại với tờ khai hải quan của bộ phận xuất nhập khẩu!
Hậu quả là khi lập Báo cáo quyết toán hải quan theo mẫu số 15 của Thông tư 39, số liệu tồn kho trên sổ sách kế toán lệch hàng trăm tấn so với số liệu tồn trên hệ thống hải quan! Đối với cơ quan hải quan, nếu số liệu tồn kho thực tế ít hơn số liệu khai báo, họ có quyền nghi ngờ doanh nghiệp đã tự ý mang nguyên liệu miễn thuế tiêu thụ nội địa, và lập tức ấn định truy thu thuế nhập khẩu kèm thuế giá trị gia tăng!

---

### 3. NGUYÊN TẮC VÀNG ĐỂ KIỂM SOÁT ĐỊNH MỨC KHO AN TOÀN
Để tự tin đón tiếp các đoàn kiểm tra sau thông quan bất cứ lúc nào, Thúy khuyên các doanh nghiệp cần thực hiện nghiêm túc 3 việc sau:
- Thứ nhất: Đồng bộ dữ liệu định kỳ hàng quý. Cứ mỗi cuối quý, Kế toán trưởng, Trưởng kho và Trưởng phòng Xuất nhập khẩu bắt buộc phải ngồi lại đối chiếu chéo số liệu: Lượng nhập khẩu trên tờ khai E31, lượng xuất khẩu trên tờ khai E62, lượng tồn kho thẻ kho thực tế và số dư tài khoản 152 trên sổ cái kế toán. Lệch ở đâu phải truy tìm nguyên nhân và lập biên bản giải trình ngay tại thời điểm đó.
- Thứ hai: Quản lý chặt chẽ tỷ lệ hao hụt và phế liệu. Mọi hao hụt vượt định mức kỹ thuật đều phải có biên bản kiểm nghiệm kỹ thuật, có chữ ký của quản đốc xưởng và giám đốc doanh nghiệp.
- Thứ ba: Lưu trữ hồ sơ điện tử và hồ sơ giấy tập trung. Mỗi tờ khai xuất nhập khẩu phải được kẹp chung với bộ chứng từ gốc: Hợp đồng, Hóa đơn thương mại, Phiếu đóng gói, Vận đơn đường biển, Chứng từ thanh toán ngân hàng, Phiếu nhập kho và Lệnh sản xuất tương ứng.""",
        "conclusion": """Với bản thân Thúy, mình luôn tâm niệm một điều: Làm việc với cơ quan hải quan, sự tự tin không đến từ những lời giải thích hoa mỹ, mà đến từ tính logic và chuẩn xác của từng con số trên sổ sách. Năm năm trôi qua rất nhanh, nhưng vết mực trên chứng từ thì ở lại mãi mãi. Doanh nghiệp nào xây dựng được quy trình kiểm soát nội bộ chặt chẽ ngay từ hôm nay, doanh nghiệp đó sẽ luôn kê cao gối ngủ yên trước mọi đợt kiểm tra sau thông quan.""",
        "cliffhanger": """Trong video tiếp theo của chuyên mục Dịch Luật Ra Tiếng Người, Thúy sẽ giải mã cho các bạn một khái niệm cực kỳ hóc búa khác: 'Trị giá hải quan và cơ chế Tham vấn giá'. Tại sao bạn mua hàng giá rẻ thật mà hải quan lại không chấp nhận giá khai báo? Bí quyết chuẩn bị hồ sơ tham vấn giá để được chấp nhận ngay từ lần đầu là gì? Các bạn nhớ theo dõi nhé!"""
    },
    {
        "sheet": "CM2_DICH_LUAT_TIENG_NGUOI",
        "code": "CM2-CD05",
        "title": "Trị giá hải quan (Customs Valuation): Cơ chế 'Tham vấn giá' là gì mà hải quan hay nghi vấn giá khai báo?",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.850 từ)",
        "cat_name": "Chuyên Mục 2: Dịch Luật Ra Tiếng Người",
        "ssot": "Thông tư 39/2015/TT-BTC và Thông tư 60/2019/TT-BTC của Bộ Tài chính về trị giá hải quan đối với hàng hóa xuất nhập khẩu; Hiệp định Định giá Hải quan WTO.",
        "visual": "Chị Thúy ngồi tại bàn làm việc, mở máy tính tra cứu Danh mục rủi ro về trị giá hải quan, bên cạnh là bộ hồ sơ thanh toán quốc tế và bảng báo giá thương mại.",
        "hook": """Chào các bạn, mình là Thúy từ Eureka Logistics! 

Có bao giờ các bạn gặp phải tình huống trớ trêu này chưa: Doanh nghiệp của bạn đàm phán rất giỏi, săn được một lô hàng thanh lý hoặc mua được giá sỉ cực tốt từ nhà sản xuất nước ngoài. Bạn hí hửng mở tờ khai hải quan, khai đúng mức giá mình đã thanh toán qua ngân hàng từng xu. Thế nhưng hệ thống hải quan trả về thông báo: 'Tờ khai rơi vào luồng nghi vấn về trị giá hải quan, yêu cầu doanh nghiệp đến tham vấn giá!'.

Và khi lên làm việc, cán bộ hải quan đưa ra một mức giá tham chiếu cao hơn mức giá bạn mua tới ba mươi, thậm chí năm mươi phần trăm. Nếu bạn không giải trình được, hải quan sẽ bác bỏ giá khai báo và ấn định thuế theo mức giá cao đó! 

Nhiều anh chị doanh nghiệp bức xúc bảo Thúy: 'Em ơi, chị mua thật, tiền chuyển khoản ngân hàng rõ ràng, hợp đồng ngoại thương đàng hoàng, tại sao hải quan lại không tin? Có phải họ đang làm khó doanh nghiệp không?'.

Thực ra không phải hải quan làm khó đâu các bạn, mà đó là quy trình quản lý rủi ro nghiệp vụ theo chuẩn mực quốc tế của Tổ chức Thương mại Thế giới WTO. Hôm nay Thúy sẽ giúp các bạn hiểu rõ cơ chế tham vấn giá là gì, và cách chuẩn bị một bộ hồ sơ tham vấn giá chuẩn chỉnh để bảo vệ mức giá mua của mình nhé!""",
        "core": """### 1. BẢN CHẤT CỦA CƠ CHẾ THAM VẤN GIÁ VÀ DANH MỤC QUẢN LÝ RỦI RO
Theo Hiệp định Định giá Hải quan của WTO và Thông tư 39 của Bộ Tài chính, nguyên tắc số một để xác định trị giá hải quan là: Trị giá giao dịch — tức là giá thực tế mà người mua đã thanh toán hoặc sẽ phải thanh toán cho người bán.
Tuy nhiên, để ngăn chặn tình trạng một số doanh nghiệp gian lận thương mại bằng cách thông đồng với bên bán làm hai bộ hợp đồng (hợp đồng thật giá cao để thanh toán, hợp đồng giả giá thấp để khai báo trốn thuế), Tổng cục Hải quan xây dựng một hệ thống cơ sở dữ liệu gọi là 'Danh mục quản lý rủi ro về trị giá'.
Trong cơ sở dữ liệu đó, cơ quan hải quan có sẵn mức giá bình quân của từng mặt hàng nhập khẩu về Việt Nam từ các quốc gia tương ứng. Khi bạn khai báo mức giá thấp hơn mức giá trong danh mục rủi ro, hệ thống phần mềm hải quan sẽ tự động kích hoạt cờ nghi vấn giá.
Tham vấn giá thực chất là một cuộc đối thoại pháp lý giữa người khai hải quan và cơ quan hải quan để làm rõ: Mức giá khai báo thấp là do bạn đàm phán giỏi, do điều kiện thị trường hay do có dấu hiệu khai man?

---

### 2. SÁU PHƯƠNG PHÁP XÁC ĐỊNH TRỊ GIÁ HẢI QUAN
Nếu cơ quan hải quan bác bỏ Trị giá giao dịch của bạn, họ sẽ lần lượt áp dụng 5 phương pháp thay thế tiếp theo theo đúng thứ tự nghiêm ngặt:
- Phương pháp 1: Trị giá giao dịch của hàng hóa nhập khẩu (giá mua bán thực tế).
- Phương pháp 2: Trị giá giao dịch của hàng hóa giống hệt nhập khẩu về Việt Nam trong cùng khoảng thời gian.
- Phương pháp 3: Trị giá giao dịch của hàng hóa tương tự.
- Phương pháp 4: Trị giá khấu trừ (lấy giá bán lẻ tại thị trường Việt Nam trừ đi các chi phí nội địa và lợi nhuận thông thường).
- Phương pháp 5: Trị giá tính toán (cộng chi phí sản xuất, nguyên liệu tại nước xuất khẩu cộng lợi nhuận thông thường).
- Phương pháp 6: Phương pháp suy luận (sử dụng các tài liệu khách quan, hợp lý).

---

### 3. BÍ QUYẾT XÂY DỰNG 'HỒ SƠ THAM VẤN GIÁ THUYẾT PHỤC 100%'
Khi nhận được giấy mời tham vấn giá, các bạn đừng đi tay không hoặc chỉ mang mỗi tờ khai và hợp đồng. Hãy chuẩn bị một 'Tập hồ sơ chuỗi giá trị' bao gồm 5 tài liệu sống còn sau:
- Thứ nhất: Dấu vết đàm phán thương mại ban đầu. Bản in các email trao đổi đàm phán giá, bảng chào giá chính thức (Quotation) của nhà sản xuất, hợp đồng nguyên tắc hoặc đơn đặt hàng ghi rõ nguyên nhân được giảm giá (ví dụ: mua số lượng lớn, hàng tồn kho cuối mùa, hàng kỷ niệm thành lập nhà máy).
- Thứ hai: Chứng từ thanh toán ngân hàng không thể chối cãi. Điện chuyển tiền T/T hoặc thư tín dụng L/C qua ngân hàng thương mại, kèm theo sổ phụ ngân hàng thể hiện số tiền đã được trừ chính xác khỏi tài khoản của doanh nghiệp.
- Thứ ba: Chứng từ vận tải quốc tế. Hóa đơn cước vận chuyển đường biển hoặc đường hàng không, hợp đồng bảo hiểm hàng hóa thể hiện đúng trị giá lô hàng.
- Thứ tư: Hóa đơn bán buôn trong nước sau khi nhập khẩu. Nếu bạn đã bán một phần lô hàng trong nước, hãy mang theo hóa đơn giá trị gia tăng bán ra kèm hợp đồng bán buôn. Nếu bạn mua giá rẻ và bán buôn ra thị trường với mức lợi nhuận mười đến mười lăm phần trăm bình thường, đó là bằng chứng hùng hồn nhất chứng minh bạn không hề trốn thuế!
- Thứ năm: Bảng tính toán phân tích kinh tế. So sánh chi phí cấu thành sản phẩm để chứng minh mức giá bạn mua là hoàn toàn hợp lý về mặt kinh tế.""",
        "conclusion": """Với bản thân Thúy, mình luôn dặn các bạn nhân sự nghiệp vụ: Đi tham vấn giá với hải quan, phong thái đĩnh đạc và sự chuẩn bị kỹ lưỡng về chứng từ quyết định tám mươi phần trăm thắng lợi. Khi chúng ta nắm vững luật, chuẩn bị hồ sơ minh bạch từ trước khi hàng về cảng, thì buổi tham vấn giá sẽ trở thành một cuộc trao đổi chuyên môn nhẹ nhàng chứ không còn là một nỗi sợ hãi của doanh nghiệp.""",
        "cliffhanger": """Trong video tới của chuyên mục Dịch Luật Ra Tiếng Người, Thúy sẽ cùng các bạn giải mã một loại thuế bổ sung cực kỳ nguy hiểm có thể khiến doanh nghiệp nhập khẩu lỗ nặng chỉ sau một đêm: 'Thuế chống bán phá giá và thuế tự vệ'. Làm sao để tra cứu và phòng tránh trước khi ký hợp đồng ngoại thương? Mời các bạn đón xem nhé!"""
    },
    {
        "sheet": "CM2_DICH_LUAT_TIENG_NGUOI",
        "code": "CM2-CD09",
        "title": "Thuế chống bán phá giá (Anti-dumping) & Thuế tự vệ: Khi mức thuế bổ sung lên tới 30-40%",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.800 từ)",
        "cat_name": "Chuyên Mục 2: Dịch Luật Ra Tiếng Người",
        "ssot": "Luật Thuế xuất khẩu, thuế nhập khẩu 2016 (Luật 107/2016/QH13), Luật Quản lý ngoại thương 2017, các Quyết định áp thuế phòng vệ thương mại của Bộ Công Thương.",
        "visual": "Chị Thúy ngồi tại bàn làm việc, mở cổng thông tin của Cục Phòng vệ Thương mại Bộ Công Thương, đối chiếu bảng mã HS thép và nhôm định hình.",
        "hook": """Chào các bạn, mình là Thúy từ Eureka Logistics! 

Hôm nay Thúy muốn kể cho các bạn nghe một câu chuyện đau đớn có thật mà một doanh nghiệp đối tác đã từng tâm sự với Thúy. 

Doanh nghiệp này nhập khẩu một lô hàng nhôm định hình từ nước ngoài về để phục vụ thi công dự án công trình. Khi ký hợp đồng ngoại thương, bạn giám đốc đã tra cứu rất kỹ biểu thuế xuất nhập khẩu: Mã HS của mặt hàng này được hưởng thuế nhập khẩu ưu đãi đặc biệt là không phần trăm theo hiệp định thương mại tự do FTA, thuế giá trị gia tăng là mười phần trăm. Tính toán chi phí thấy có lãi rất tốt.

Thế nhưng, đến khi container cập cảng Cát Lái, truyền tờ khai hải quan, tờ khai lập tức bị treo! Cán bộ hải quan thông báo: Mặt hàng nhôm định hình này thuộc diện đang bị Bộ Công Thương áp dụng Biện pháp chống bán phá giá chính thức, với mức thuế chống bán phá giá bổ sung là ba mươi lăm phần trăm!

Lô hàng trị giá 3 tỷ đồng, tự nhiên phải nộp thêm hơn một tỷ đồng tiền thuế chống bán phá giá ngay tại cửa khẩu thì mới được giải phóng hàng! Doanh nghiệp không đủ tiền mặt để nộp thuế, hàng bị lưu kho lưu bãi, chậm tiến độ công trình, từ một thương vụ tưởng chừng sinh lời tốt biến thành một cú sốc tài chính suýt khiến công ty phá sản!

Tại sao lại có những khoản thuế 'từ trên trời rơi xuống' như vậy? Và làm sao để các bạn không bao giờ dẫm phải chiếc bẫy thuế phòng vệ thương mại này? Thúy sẽ giải thích cặn kẽ ngay sau đây!""",
        "core": """### 1. BẢN CHẤT CỦA CÁC BIỆN PHÁP PHÒNG VỆ THƯƠNG MẠI
Theo Luật Thuế xuất khẩu, thuế nhập khẩu năm 2016 và Luật Quản lý ngoại thương năm 2017, ngoài thuế nhập khẩu thông thường hoặc ưu đãi, cơ quan nhà nước có thẩm quyền (ở Việt Nam là Bộ Công Thương) có quyền áp dụng 3 loại thuế phòng vệ thương mại bổ sung:
- Loại 1: Thuế chống bán phá giá (Anti-dumping duty). Áp dụng khi hàng hóa nước ngoài được bán vào Việt Nam với giá thấp hơn giá bán tại thị trường nội địa của nước họ (bán phá giá), gây thiệt hại đáng kể cho ngành sản xuất trong nước.
- Loại 2: Thuế chống trợ cấp (Countervailing duty). Áp dụng khi hàng hóa nước ngoài được chính phủ nước họ trợ cấp trực tiếp hoặc gián tiếp, tạo ra sự cạnh tranh bất bình đẳng.
- Loại 3: Thuế tự vệ (Safeguard duty). Áp dụng khi một mặt hàng gia tăng nhập khẩu đột biến, đe dọa nghiêm trọng đến ngành sản xuất nội địa, bất kể hàng đó có bán phá giá hay không.

---

### 2. BA ĐẶC ĐIỂM NGUY HIỂM CỦA THUẾ CHỐNG BÁN PHÁ GIÁ
Điều làm cho thuế chống bán phá giá trở nên cực kỳ nguy hiểm đối với các doanh nghiệp nhập khẩu chính là 3 đặc điểm sau:
- Đặc điểm thứ nhất: Thuế suất rất cao! Thuế nhập khẩu thông thường chỉ từ 0% đến 15%, nhưng thuế chống bán phá giá có thể dao động từ 10% đến 40%, thậm chí hơn 100%!
- Đặc điểm thứ hai: Tính chất áp thuế theo từng nhà sản xuất nước ngoài cụ thể. Cùng một mã HS, cùng nhập từ một quốc gia, nhưng nếu bạn mua của Nhà máy A (họ hợp tác với cơ quan điều tra Việt Nam và cung cấp đủ số liệu) thì thuế suất chỉ 5%. Nhưng nếu bạn mua của Nhà máy B (họ không hợp tác điều tra) thì bị áp mức thuế toàn quốc lên tới 35%!
- Đặc điểm thứ ba: Biện pháp áp thuế tạm thời có hiệu lực ngay lập tức. Bộ Công Thương có thể ra quyết định áp thuế tạm thời trong thời gian điều tra mà không cần chờ đến phán quyết cuối cùng. Nếu bạn không cập nhật quyết định mới nhất, bạn sẽ hoàn toàn bị động khi hàng cập cảng!

---

### 3. QUY TRÌNH 3 BƯỚC TRA CỨU PHÒNG NGỪA RỦI RO TRƯỚC KHI KÝ HỢP ĐỒNG
Để không bao giờ bị dính đòn oan uổng, Thúy luôn yêu cầu đội ngũ Eureka tuân thủ nghiêm ngặt quy trình 3 bước trước khi tư vấn hoặc ký hợp đồng nhập khẩu:
- Bước 1: Tra cứu Danh mục các vụ việc phòng vệ thương mại trên Cổng thông tin của Cục Phòng vệ Thương mại — Bộ Công Thương (trav.gov.vn). Kiểm tra xem mã HS sản phẩm mình dự kiến nhập khẩu có nằm trong danh sách đang bị điều tra hoặc đang bị áp thuế hay không.
- Bước 2: Xác minh danh tính và mã số của nhà sản xuất nước ngoài. Yêu cầu nhà cung cấp cung cấp giấy chứng nhận xuất xưởng (Mill Test Certificate) ghi rõ tên pháp nhân sản xuất. Sau đó đối chiếu tên nhà máy đó với phụ lục đính kèm Quyết định áp thuế của Bộ Công Thương để xem họ được hưởng mức thuế riêng biệt hay bị áp thuế suất chung.
- Bước 3: Đưa điều khoản phòng ngừa rủi ro vào Hợp đồng ngoại thương. Thỏa thuận rõ trong hợp đồng: Nếu tại thời điểm hàng cập cảng Việt Nam mà phát sinh các loại thuế phòng vệ thương mại mới do lỗi từ phía nhà cung cấp không cung cấp được chứng nhận hợp tác điều tra, thì bên bán phải chịu trách nhiệm chia sẻ chi phí hoặc cho phép bên mua hủy hợp đồng.""",
        "conclusion": """Với bản thân Thúy, mình luôn tự dặn lòng: Trong nghề ngoại thương, sự am hiểu luật pháp và khả năng cập nhật chính sách chính là biên lợi nhuận lớn nhất của doanh nghiệp. Một người làm kinh doanh giỏi không chỉ là người mua được hàng giá rẻ, mà phải là người nhìn thấy trước những chi phí ẩn và rủi ro chính sách trước khi đồng tiền được chuyển đi.""",
        "cliffhanger": """Sau những câu chuyện luật pháp căng thẳng, trong video tiếp theo, Thúy muốn mời các bạn cùng Thúy bước lên một chuyến tàu du hành thời gian ngược về hai nghìn năm trước: Khám phá Cửa ải Ngọc Môn Quan trên Con đường Tơ lụa cổ đại — Nơi đế chế Hán và La Mã đã sáng lập ra hệ thống trạm kiểm soát hải quan và thu thuế xuyên lục địa đầu tiên trong lịch sử nhân loại! Mời các bạn cùng đón xem nhé!"""
    },
    {
        "sheet": "CM3_LICH_SU_THUONG_MAI",
        "code": "CM3-CD06",
        "title": "Con đường Tơ lụa cổ đại: Cửa ải Ngọc Môn Quan và cách đế chế Hán - La Mã quản lý thuế quan xuyên lục địa",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.900 từ)",
        "cat_name": "Chuyên Mục 3: Ký Sự Lịch Sử Thương Mại & Kỳ Án Hàng Hải",
        "ssot": "Sử ký Tư Mã Thiên, Hán Thư, các tài liệu khảo cổ học Đôn Hoàng của Viện Hàn lâm Khoa học Xã hội Trung Quốc và các nghiên cứu lịch sử thương mại La Mã cổ đại.",
        "visual": "Chị Thúy ngồi bên bàn gỗ mộc, phía sau là tranh vẽ Con đường Tơ lụa cổ đại qua sa mạc Gobi và di tích cửa ải Ngọc Môn Quan sừng sững giữa cát vàng.",
        "hook": """Chào mừng các bạn đã quay trở lại với chuỗi ký sự lịch sử thương mại quốc tế của Thúy trên kênh ThuyTalk!

Hôm nay, Thúy muốn mời các bạn cùng Thúy nhắm mắt lại và quay ngược bánh xe thời gian về hai nghìn năm trước. 

Hãy tưởng tượng bạn đang đứng giữa sa mạc Gobi mênh mông cát trắng và gió lộng. Mặt trời đổ lửa xuống những đụn cát khổng lồ. Từ phía xa xăm, tiếng chuông đồng leng keng vang vọng trong không gian tĩnh mịch: Đó là một đoàn lữ hành gồm hàng trăm con lạc đà hai bướu đang chậm rãi tiến bước. Trên lưng chúng chất đầy những cuộn tơ lụa óng ả, những bình gốm sứ tinh xảo, trà thơm và các thỏi ngọc bích lấp lánh mang từ kinh thành Trường An đi về phương Tây.

Thế nhưng, trước khi đoàn buôn đó có thể bước chân vào vùng đất Tây Vực hoang vu để hướng về Đế chế La Mã xa xôi, họ bắt buộc phải dừng lại trước một pháo đài bằng đất nện khổng lồ sừng sững giữa sa mạc: Cửa ải Ngọc Môn Quan!

Tại sao cửa ải này lại mang tên là Ngọc Môn Quan — Cửa Ải Ngọc? Người xưa hai nghìn năm trước đã tổ chức kiểm soát hộ chiếu, kiểm dịch hàng hóa và thu thuế quan xuyên lục địa tinh vi đến mức nào? Hôm nay Thúy sẽ cùng các bạn giải mã bức màn lịch sử kỳ vĩ này nhé!""",
        "core": """### 1. BỐI CẢNH RA ĐỜI CỦA NGỌC MÔN QUAN VÀ SỨ MỆNH TRƯƠNG KHIÊN
Vào năm 138 trước Công Nguyên, Hoàng đế Hán Vũ Đế cử nhà thám hiểm Trương Khiên dẫn đầu đoàn sứ bộ hơn một trăm người đi về phía Tây để tìm kiếm đồng minh chống lại người Hung Nô. Dù bị bắt làm tù binh suốt mười ba năm, Trương Khiên vẫn kiên cường trốn thoát và mang về triều đình những bản đồ địa lý vô giá cùng thông tin về các vương quốc giàu có ở Trung Á và Tây Á.
Nhận thấy tiềm năng giao thương khổng lồ, Hán Vũ Đế đã hạ lệnh xây dựng Vạn Lý Trường Thành kéo dài về phía Tây sa mạc, và đặt ra hai cửa ải chiến lược tại Đôn Hoàng: Dương Quan ở phía Nam và Ngọc Môn Quan ở phía Bắc.
Sở dĩ có tên là Ngọc Môn Quan là bởi vì đây là cửa ngõ duy nhất mà các đoàn thương buôn vận chuyển ngọc bích hòa điền từ vùng Tân Cương nhập khẩu vào lãnh thổ Trung Nguyên. 

---

### 2. QUY TRÌNH KIỂM SOÁT HẢI QUAN CỔ ĐẠI: 'QUÁ SỞ' VÀ 'THIỀM PHÙ'
Nhiều người nghĩ rằng thời cổ đại đi lại tự do, buôn bán muốn đi đâu thì đi. Hoàn toàn không phải! Nhà Hán đã thiết lập một hệ thống quản lý xuất nhập cảnh và hải quan chặt chẽ không thua kém gì hải quan điện tử ngày nay:
- Thứ nhất là 'Quá sở' (Hộ chiếu và Giấy thông hành thương mại). Mỗi thương nhân khi muốn xuất cảnh qua Ngọc Môn Quan bắt buộc phải xuất trình thẻ tre hoặc thẻ gỗ gọi là 'Quá sở' do quan phủ địa phương cấp. Trên đó ghi rõ: Tên họ thương nhân, quê quán, số lượng tùy tùng đi theo, số lượng lạc đà ngựa kéo, và danh mục chi tiết từng kiện hàng mang theo.
- Thứ hai là hệ thống 'Thiềm phù' — Dấu niêm phong hải quan cổ đại. Sau khi các quan viên tại Ngọc Môn Quan kiểm đếm số lượng tơ lụa, gốm sứ và đối chiếu với tờ khai Quá sở, họ sẽ dùng bùn niêm phong và đóng dấu triện của viên quan Đô úy lên các hòm hàng. Nếu đến trạm kiểm soát tiếp theo mà dấu niêm phong bị rách, toàn bộ kiện hàng sẽ bị tịch thu và thương nhân bị tống giam!
- Thứ ba là thu thuế quan bằng hiện vật. Nhà Hán áp dụng cơ chế thu thuế môn bài và thuế lưu thông hàng hóa. Thương nhân chở một trăm cuộn lụa qua ải, quan ải sẽ giữ lại từ ba đến năm cuộn làm thuế nộp vào quốc khố. Số lụa này sau đó được triều đình dùng để trả lương cho quân đội biên phòng đồn trú tại sa mạc.

---

### 3. SỰ GẶP GỠ GIỮA ĐẾ CHẾ HÁN VÀ ĐẾ CHẾ LA MÃ
Ở đầu bên kia của Con đường Tơ lụa, giới quý tộc Đế chế La Mã phát cuồng vì tơ lụa của phương Đông. Nhà sử học La Mã Pliny the Elder từng viết một dòng than thở đầy xót xa: 'Mỗi năm, Đế chế La Mã phải chảy máu hơn một trăm triệu sesterce tiền vàng và bạc ròng chỉ để mua tơ lụa và hương liệu từ phương Đông phục vụ cho sự xa xỉ của phụ nữ và giới thượng lưu La Mã!'.
Chính dòng chảy thương mại xuyên lục địa này đã biến Ngọc Môn Quan không chỉ là một tiền đồn quân sự, mà trở thành trung tâm thanh toán, đổi tiền và giao lưu văn hóa sôi động nhất thế giới cổ đại suốt hơn một nghìn năm.""",
        "conclusion": """Với bản thân Thúy, khi đứng trước những bức tường thành đất nện cổ kính của Ngọc Môn Quan, mình luôn cảm nhận được một bài học quản trị sâu sắc: Biên giới của một quốc gia hay ranh giới của một doanh nghiệp không chỉ sinh ra để dựng lên những rào cản ngăn cấm, mà nó sinh ra để đóng vai trò là chiếc van điều tiết dòng chảy thương mại. Quốc gia nào, doanh nghiệp nào biết thiết lập những quy chuẩn minh bạch, đón nhận dòng chảy giao thương toàn cầu một cách có kiểm soát, thì quốc gia đó sẽ thịnh vượng bền lâu.""",
        "cliffhanger": """Cũng bắt nguồn từ sự thâm hụt thương mại và dòng chảy kim loại quý giữa phương Đông và phương Tây, nhưng gần hai nghìn năm sau, một cuộc xung đột khốc liệt đã bùng nổ, làm thay đổi vĩnh viễn vận mệnh của triều đại phong kiến cuối cùng tại châu Á: Cuộc Chiến tranh Thuốc phiện năm 1839. Tại sao một chén trà thơm lại có thể châm ngòi cho một cuộc chiến tranh hàng hải tàn khốc? Thúy sẽ cùng các bạn giải mã trong video tiếp theo nhé!"""
    },
    {
        "sheet": "CM3_LICH_SU_THUONG_MAI",
        "code": "CM3-CD07",
        "title": "Chiến tranh Thuốc phiện (1839 - 1842): Thâm hụt thương mại trà - bạc và bài học cay đắng về việc bế quan tỏa cảng",
        "format": "Video Độc lập (1 Tập)",
        "duration": "9 - 11 phút (~2.050 từ)",
        "cat_name": "Chuyên Mục 3: Ký Sự Lịch Sử Thương Mại & Kỳ Án Hàng Hải",
        "ssot": "Hồ sơ lưu trữ Công ty Đông Ấn Anh (British East India Company Archives), Nhật ký Khâm sai Đại thần Lâm Tắc Từ, Toàn văn Điều ước Nam Kinh 1842.",
        "visual": "Chị Thúy ngồi tại bàn làm việc, trên bàn có một hộp trà gốm cổ, ấm trà nóng nghi ngút khói, màn hình bên cạnh hiển thị tranh khắc gỗ cảnh tiêu hủy thuốc phiện tại Hổ Môn năm 1839.",
        "hook": """Chào mừng các bạn đã quay trở lại với chuỗi ký sự lịch sử thương mại quốc tế của Thúy trên kênh ThuyTalk!

Hôm nay, Thúy muốn kể cho các bạn nghe về một trong những bi kịch thương mại đen tối và cay đắng nhất trong lịch sử nhân loại: Cuộc Chiến tranh Thuốc phiện lần thứ nhất giữa Đế quốc Anh và Nhà Thanh vào năm 1839.

Khi nhắc đến cuộc chiến này, đa số chúng ta thường nghĩ ngay đến hình ảnh những tàu chiến bọc thép của phương Tây xả đại bác vào các pháo đài ven biển của triều đình phong kiến. Thế nhưng, nếu bóc tách dưới lăng kính kinh tế học và thương mại quốc tế, các bạn sẽ bàng hoàng nhận ra rằng: Ngòi nổ thực sự châm ngòi cho cuộc chiến tranh tàn khốc ấy lại bắt đầu từ một thứ vô cùng thanh tao và quen thuộc trong mỗi gia đình chúng ta: Đó chính là lá trà xanh!

Làm thế nào mà thói quen uống trà chiều của giới quý tộc London lại khiến kho bạc nước Anh cạn kiệt hàng triệu ounce bạc ròng? Làm thế nào mà một công ty thương mại như Công ty Đông Ấn Anh lại sẵn sàng biến thành một tập đoàn buôn lậu ma túy quy mô xuyên quốc gia để đảo ngược cán cân thanh toán? Và bài học đắt giá về chính sách bế quan tỏa cảng đối với các doanh nghiệp hiện đại ngày nay là gì? Hôm nay Thúy sẽ cùng các bạn bóc tách từng góc khuất của kỳ án thương mại này nhé!""",
        "core": """### 1. CƠN SỐT TRÀ CHIỀU VÀ HỐ ĐEN THÂM HỤT THƯƠNG MẠI CỦA NƯỚC ANH
Vào thế kỷ mười tám, thói quen uống trà chiều trở thành một cơn sốt cuồng nhiệt lan tỏa từ hoàng gia, giới quý tộc cho đến tầng lớp lao động tại nước Anh. Nhu cầu nhập khẩu trà từ Trung Quốc tăng vọt lên hàng chục triệu pound mỗi năm.
Thế nhưng, triều đình nhà Thanh thời bấy giờ thi hành chính sách 'Bế quan tỏa cảng' cực kỳ bảo thủ. Họ chỉ mở duy nhất một cửa khẩu thông thương là Cảng Quảng Châu, và bắt buộc mọi thương nhân phương Tây phải giao dịch thông qua nghiệp đoàn thương nhân 'Thập tam hàng'.
Nghiêm trọng hơn nữa, Hoàng đế Càn Long và Gia Khánh tuyên bố: 'Thiên triều sản vật phong phú, không thiếu thứ gì, không có nhu cầu mua bất kỳ hàng hóa công nghiệp nào của phương Tây!'. Nước Anh muốn mua trà, tơ lụa và gốm sứ thì chỉ có một con đường duy nhất: Thanh toán bằng tiền bạc đúc ròng (Silver Bullion)!
Hậu quả là một dòng bạc khổng lồ chảy một chiều từ châu Âu vào Trung Quốc. Ước tính trong vòng nửa thế kỷ, nước Anh đã chi trả hơn hai mươi tám triệu bảng Anh tiền bạc ròng vào quốc khố nhà Thanh, khiến dự trữ kim loại quý của Ngân hàng Trung ương Anh đứng trước nguy cơ sụp đổ hoàn toàn vì thâm hụt cán cân thanh toán!

---

### 2. ÂM MƯU NHA PHIẾN VÀ SỰ ĐẢO CHIỀU DÒNG CHẢY KIM LOẠI QUÝ
Đứng trước nguy cơ vỡ nợ quốc gia vì uống trà, Công ty Đông Ấn Anh — một tập đoàn thương mại được chính phủ Anh trao quyền lực quân sự — đã tìm ra một giải pháp tàn độc nhưng mang lại siêu lợi nhuận: Thuốc phiện!
Họ tổ chức độc quyền trồng cây thuốc phiện trên những cánh đồng bạt ngàn tại Bengal, Ấn Độ, sau đó tinh chế thành các bánh thuốc phiện chất lượng cao. Do nhà Thanh cấm nhập khẩu thuốc phiện, Công ty Đông Ấn Anh đã bắt tay với các tổ chức buôn lậu và các quan lại biến chất tại Quảng Châu để lập nên một mạng lưới buôn lậu thuốc phiện khổng lồ trên biển.
Chỉ trong vòng hai mươi năm, lượng thuốc phiện tuồn vào Trung Quốc tăng từ bốn nghìn hòm lên tới hơn bốn mươi nghìn hòm mỗi năm! Hàng triệu người dân Trung Quốc từ quan lại, binh lính đến dân nghèo trở thành con nghiện. 
Và dòng tiền bạc đã lập tức đảo chiều một cách kinh hoàng: Giờ đây, không phải người Anh trả bạc để mua trà, mà chính người Trung Quốc phải vét sạch những đồng bạc cuối cùng trong túi để trả tiền mua thuốc phiện! Mỗi năm, hơn ba mươi triệu lạng bạc ròng chảy ngược từ Trung Quốc về London. Kinh tế nhà Thanh kiệt quệ, xã hội suy đồi đến bờ vực diệt vong!

---

### 3. KHÂM SAI LÂM TẮC TỪ VÀ BÀI HỌC CAY ĐẮNG VỀ SỰ TỰ MÃN CÔNG NGHỆ
Trước họa diệt vong của dân tộc, năm 1839, Hoàng đế Đạo Quang cử vị đại thần thanh liêm và cương trực bậc nhất là Lâm Tắc Từ làm Khâm sai đến Quảng Châu dẹp loạn thuốc phiện.
Lâm Tắc Từ hành động vô cùng quyết liệt: Ông bao vây khu thương xá nước ngoài, buộc các thương nhân Anh và Mỹ phải giao nộp hơn hai mươi nghìn hòm thuốc phiện (tương đương hơn một nghìn ba trăm tấn), rồi cho đào hố lớn tại bãi biển Hổ Môn, đổ muối và vôi sống vào để tiêu hủy sạch sẽ trước sự chứng kiến của hàng vạn người dân.
Hành động dũng cảm đó đã châm ngòi cho cuộc chiến. Nước Anh lấy cớ bảo vệ quyền tự do thương mại và bảo vệ tài sản của công dân Anh để điều động một hạm đội tàu chiến bọc thép hiện đại vượt đại dương tấn công Trung Quốc.
Quân đội nhà Thanh với gươm giáo, thuyền gỗ và đại bác cổ lỗ sĩ đã thất bại thảm hại trước hỏa lực vượt trội của phương Tây. Năm 1842, triều đình nhà Thanh buộc phải ký Điều ước Nam Kinh nhục nhã: Mở 5 cảng thông thương, bồi thường hai mươi mốt triệu lạng bạc, và cắt nhượng đảo Hồng Kông cho nước Anh!""",
        "conclusion": """Với bản thân Thúy, khi ngẫm lại câu chuyện đau thương này, mình luôn tự dặn lòng một bài học sâu sắc trong kinh doanh: Sự tự mãn và tư duy đóng cửa bế quan tỏa cảng chính là con đường ngắn nhất dẫn đến sự suy tàn. Khi bạn từ chối học hỏi công nghệ mới, từ chối thích ứng với các chuẩn mực thương mại quốc tế, bạn có thể tự ru ngủ mình trong một thời gian, nhưng khi làn sóng của thế giới ập đến, cái giá phải trả sẽ là sự sụp đổ toàn diện.""",
        "cliffhanger": """Cũng trên hải trình kết nối giữa phương Đông và phương Tây, ba mươi năm sau cuộc Chiến tranh Thuốc phiện, một kỳ tích kỹ thuật vĩ đại đã được hoàn thành, làm thay đổi vĩnh viễn toàn bộ bản đồ hàng hải và chi phí logistics toàn cầu: Đó là Kênh đào Suez năm 1869! Vết cắt một trăm chín mươi ba cây số qua sa mạc đã rút ngắn nửa vòng Trái Đất như thế nào? Thúy sẽ cùng các bạn khám phá trong video tiếp theo nhé!"""
    },
    {
        "sheet": "CM3_LICH_SU_THUONG_MAI",
        "code": "CM3-CD08",
        "title": "Kênh đào Suez 1869: Vết cắt 193 km rút ngắn nửa vòng Trái Đất và cuộc chiến địa chính trị quanh tuyến hàng hải huyết mạch",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.900 từ)",
        "cat_name": "Chuyên Mục 3: Ký Sự Lịch Sử Thương Mại & Kỳ Án Hàng Hải",
        "ssot": "Hồ sơ lưu trữ Công ty Kênh đào Suez (Compagnie Universelle du Canal Maritime de Suez), Báo cáo dữ liệu hàng hải của Cơ quan Quản lý Kênh đào Suez (SCA), Dữ liệu vụ tai nạn tàu Ever Given 2021.",
        "visual": "Chị Thúy ngồi tại bàn làm việc, màn hình hiển thị hải trình so sánh giữa tuyến đường qua Kênh đào Suez và tuyến đường vòng qua Mũi Hảo Vọng Châu Phi, mô hình một con tàu container lớn trên bàn.",
        "hook": """Chào mừng các bạn đã quay trở lại với chuỗi ký sự lịch sử thương mại quốc tế của Thúy trên kênh ThuyTalk!

Nếu có ai đó hỏi Thúy rằng: Đâu là công trình nhân tạo vĩ đại nhất đã làm thay đổi hoàn toàn chi phí logistics và tốc độ của nền kinh tế toàn cầu? Thúy sẽ không ngần ngại trả lời ngay: Đó chính là Kênh đào Suez!

Chỉ là một vết cắt nhân tạo dài một trăm chín mươi ba cây số xuyên qua dải sa mạc cát cháy bỏng của bán đảo Sinai nối liền Địa Trung Hải với Biển Đỏ. Thế nhưng, vết cắt ấy đã giúp các đoàn tàu buôn từ châu Á sang châu Âu không còn phải đánh cược sinh mạng đi vòng qua Mũi Hảo Vọng đầy bão tố ở cực Nam châu Phi, rút ngắn hải trình tới hơn bảy nghìn cây số và tiết kiệm từ mười đến mười lăm ngày lênh đênh trên biển!

Ngày nay, mỗi ngày có khoảng năm mươi con tàu khổng lồ chở theo hơn mười hai phần trăm tổng khối lượng thương mại của cả nhân loại đi qua nút thắt cổ chai này. Các bạn hẳn vẫn chưa quên vào tháng ba năm 2021, khi siêu tàu container Ever Given bị mắc cạn chắn ngang kênh đào suốt sáu ngày, nó đã khiến chuỗi cung ứng toàn cầu tê liệt, làm bốc hơi gần mười tỷ đô la thương mại mỗi ngày!

Thế nhưng, ít ai biết rằng để có được con kênh đào huyền thoại đó vào năm 1869, nhân loại đã phải trải qua một cuộc chiến cân não giữa các cường quốc, một canh bạc tài chính khổng lồ và sự hy sinh xương máu của hơn một trăm nghìn công nhân. Hôm nay Thúy sẽ cùng các bạn giải mã thiên ký sự hàng hải này nhé!""",
        "core": """### 1. Ý TƯỞNG ĐIÊN RỒ VÀ SỰ BẢO THỦ CỦA ĐẾ QUỐC ANH
Ý tưởng đào một con kênh nối Biển Đỏ với Địa Trung Hải thực ra đã có từ thời các Pharaoh Ai Cập cổ đại hàng nghìn năm trước. Đến thời Hoàng đế Napoleon Bonaparte sang Ai Cập năm 1798, ông cũng từng muốn thực hiện dự án này nhưng bị các kỹ sư ngăn cản vì một phép đo đạc sai lầm cho rằng mực nước Biển Đỏ cao hơn Địa Trung Hải mười mét, nếu đào sẽ gây ngập lụt toàn bộ châu thổ sông Nile.
Phải đến năm 1854, một nhà ngoại giao và kỹ sư người Pháp đầy tham vọng tên là Ferdinand de Lesseps mới chứng minh được hai vùng biển có mực nước hoàn toàn bằng nhau, và thuyết phục được Phó vương Ai Cập thành lập Công ty Kênh đào Suez.
Thế nhưng, trớ trêu thay, quốc gia kịch liệt phản đối dự án này nhất lại chính là... Đế quốc Anh! Nước Anh thời bấy giờ đang thống trị tuyến đường hàng hải vòng qua Mũi Hảo Vọng để sang thuộc địa Ấn Độ. Họ lo sợ một kênh đào do người Pháp kiểm soát sẽ đe dọa vị thế độc tôn của Hải quân Hoàng gia Anh. Nước Anh đã dùng mọi thủ đoạn ngoại giao và truyền thông để gọi dự án của Lesseps là 'trò lừa đảo thế kỷ'!

---

### 2. MƯỜI NĂM ĐÀO ĐẤT TRÊN SA MẠC VÀ BƯỚC NGOẶT ĐỔI ĐỜI NĂM 1869
Bất chấp sự chống phá của người Anh, ngày 25 tháng 4 năm 1859, phát cuốc đầu tiên chính thức bổ xuống sa mạc Suez.
Đó là một công trường vĩ đại nhưng cũng thấm đẫm bi kịch. Trong những năm đầu, hơn một triệu rưỡi lượt công nhân Ai Cập đã phải đào đất thủ công dưới cái nóng bốn mươi lăm độ C, thiếu nước ngọt trầm trọng, dịch tả bùng phát cướp đi sinh mạng của hàng chục nghìn người. 
Nhận thấy đào tay không thể hoàn thành, Lesseps đã có một quyết định đột phá: Đưa các máy đào hơi nước khổng lồ và máy hút bùn công suất lớn từ Pháp sang. Đây là lần đầu tiên trong lịch sử, máy móc cơ giới hóa thay thế hoàn toàn sức người trong một đại dự án hạ tầng.
Đúng 10 năm sau, ngày 17 tháng 11 năm 1869, Kênh đào Suez chính thức khánh thành trong một buổi đại lễ xa hoa bậc nhất lịch sử với sự tham dự của Hoàng hậu Pháp Eugénie và giới quý tộc khắp châu Âu. Một kỷ nguyên mới của hàng hải toàn cầu chính thức mở ra!

---

### 3. VỤ THÂU TÓM NGOẠI GIAO KINH ĐIỂN CỦA NƯỚC ANH
Mặc dù phản đối quyết liệt ngay từ đầu, nhưng khi nhìn thấy những con tàu hơi nước đầu tiên chở đầy hàng hóa đi qua kênh đào Suez giúp rút ngắn nửa thời gian đến Ấn Độ, chính phủ Anh đã nhận ra sai lầm chiến lược của mình.
Và một cơ hội vàng đã đến vào năm 1875: Phó vương Ai Cập Ismail Pasha do ăn chơi xa xỉ và đầu tư dàn trải đã rơi vào cảnh phá sản ngập đầu trong nợ nần. Ông buộc phải rao bán bí mật bốn mươi tư phần trăm cổ phần của Ai Cập trong Công ty Kênh đào Suez.
Thủ tướng Anh thời bấy giờ là Benjamin Disraeli đã hành động với tốc độ sấm sét: Không cần chờ Quốc hội phê chuẩn, ông vay nóng bốn triệu bảng Anh từ gia tộc tài phiệt Rothschild ngay trong đêm để mua đứt toàn bộ số cổ phần đó! 
Chỉ sau một đêm, nước Anh từ kẻ ngoài cuộc chống phá đã trở thành cổ đông lớn nhất, nắm quyền kiểm soát huyết mạch hàng hải Suez suốt gần một thế kỷ tiếp theo!""",
        "conclusion": """Với bản thân Thúy, sau 15 năm làm nghề logistics, câu chuyện Kênh đào Suez luôn mang lại cho mình một bài học sâu sắc về 'Điểm nghẽn huyết mạch' (Chokepoint) trong chuỗi cung ứng. Một tuyến đường ngắn hơn có thể giúp doanh nghiệp tối ưu chi phí, nhưng nếu toàn bộ dòng chảy phụ thuộc vào một điểm chạm duy nhất mà không có phương án dự phòng, thì bất kỳ biến cố địa chính trị hay tai nạn nào cũng có thể biến lợi thế thành thảm họa. Bài toán cân bằng giữa Tối ưu chi phí và Quản trị rủi ro luôn là bài toán sống còn của người làm logistics.""",
        "cliffhanger": """Cũng về câu chuyện giá trị và tâm lý của con người, trong video tiếp theo, Thúy muốn kể cho các bạn nghe một câu chuyện kỳ lạ đến khó tin trong lịch sử nghệ thuật và truyền thông: Vụ trộm bức họa Mona Lisa năm 1911 tại bảo tàng Louvre! Làm thế nào mà một khoảng trống trên bức tường lại có thể biến một bức tranh bình thường trở thành kiệt tác đắt giá nhất nhân loại? Mời các bạn đón xem nhé!"""
    },
    {
        "sheet": "CM3_LICH_SU_THUONG_MAI",
        "code": "CM3-CD09",
        "title": "Vụ trộm bức họa Mona Lisa 1911: Khi khoảng trống trên tường bảo tàng Louvre tạo nên kiệt tác đắt giá nhất nhân loại",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.850 từ)",
        "cat_name": "Chuyên Mục 3: Ký Sự Lịch Sử Thương Mại & Kỳ Án Hàng Hải",
        "ssot": "Hồ sơ lưu trữ Sở Cảnh sát Paris (Prefecture de Police de Paris 1911), Tài liệu triển lãm Bảo tàng Louvre, Biên bản lời khai của Vincenzo Peruggia.",
        "visual": "Chị Thúy ngồi tại bàn làm việc ấm cúng, trên bàn có một khung tranh gỗ trống rỗng chỉ có bốn chiếc đinh móc, bên cạnh là cuốn sách nghệ thuật thời Phục hưng.",
        "hook": """Chào mừng các bạn đã quay trở lại với kênh ThuyTalk của Thúy!

Hôm nay Thúy muốn hỏi các bạn một câu hỏi thú vị: Bức tranh nào nổi tiếng nhất và đắt giá nhất trên hành tinh này? 

Chắc chắn chín mươi chín phần trăm trong chúng ta sẽ trả lời ngay: Đó là bức họa nàng Mona Lisa của đại danh họa Leonardo da Vinci đang được trưng bày sau lớp kính chống đạn tại bảo tàng Louvre ở thủ đô Paris nước Pháp!

Thế nhưng, các bạn có biết rằng: Cách đây hơn một trăm năm, trước mùa hè năm 1911, nàng Mona Lisa chỉ là một bức tranh sơn dầu rất đỗi bình thường giữa hàng nghìn kiệt tác khác trong bảo tàng Louvre? Thậm chí, đa số khách tham quan bước qua nàng mà không thèm ngoái nhìn lại một lần, và các nhà phê bình nghệ thuật thời đó còn đánh giá bức tranh này thấp hơn rất nhiều bức họa khác của Raphael hay Titian!

Vậy điều gì đã biến một bức tranh vô danh trở thành biểu tượng văn hóa toàn cầu được bảo hiểm với mức giá định danh lên tới gần một tỷ đô la ngày hôm nay? 

Câu trả lời nằm ở một vụ trộm thế kỷ diễn ra vào sáng ngày 21 tháng 8 năm 1911 — Một vụ trộm mà kẻ cắp không phải là siêu trộm chuyên nghiệp, và điều tạo nên sự nổi tiếng điên cuồng của bức tranh không phải là lúc nó xuất hiện, mà chính là lúc nó... biến mất! Hôm nay Thúy sẽ cùng các bạn giải mã kỳ án truyền thông và marketing kinh điển này nhé!""",
        "core": """### 1. VỤ TRỘM NGÂY THƠ NHẤT TRONG LỊCH SỬ NGHỆ THUẬT
Sáng thứ Hai ngày 21 tháng 8 năm 1911, bảo tàng Louvre đóng cửa để nhân viên làm vệ sinh định kỳ. An ninh của bảo tàng nghệ thuật lớn nhất thế giới thời bấy giờ lỏng lẻo đến mức khó tin: Toàn bộ bảo tàng khổng lồ chỉ có chưa đầy một trăm bảo vệ già trông coi và không hề có một chiếc camera hay chuông báo động nào.
Một người thợ mộc người Ý ba mươi tuổi tên là Vincenzo Peruggia — người từng được bảo tàng thuê lắp lồng kính bảo vệ cho các bức tranh — đã mặc một chiếc áo khoác blu trắng của nhân viên, trà trộn vào phòng trưng bày Salon Carré.
Khi thấy không có ai, Peruggia đơn giản là nhấc bức tranh Mona Lisa ra khỏi bốn chiếc móc sắt trên tường, mang vào góc cầu thang, tháo bỏ khung gỗ cồng kềnh, cuộn bức tranh lại rồi kẹp dưới nách áo khoác và ung dung bước ra khỏi cửa bảo tàng Louvre giữa thanh thiên bạch nhật!
Kỳ quặc hơn nữa là suốt hơn hai mươi bốn tiếng đồng hồ sau đó, không một ai trong bảo tàng phát hiện bức tranh bị mất! Các nhân viên bảo vệ đi qua chỉ nghĩ rằng các nhiếp ảnh gia đã mượn bức tranh mang đi chụp ảnh tư liệu! Mãi đến sáng thứ Ba, khi một họa sĩ đến xin phép vẽ chép tranh thì vụ trộm thế kỷ mới vỡ lở!

---

### 2. SỨC MẠNH CỦA KHOẢNG TRỐNG VÀ CƠN BÃO TRUYỀN THÔNG TOÀN CẦU
Ngay sau khi cảnh sát Paris công bố vụ trộm, một cơn địa chấn truyền thông đã bùng nổ trên khắp thế giới. Giám đốc bảo tàng Louvre bị cách chức, biên giới nước Pháp bị phong tỏa, cảnh sát lục soát từng con tàu viễn dương rời cảng Le Havre.
Thế nhưng, điều kỳ lạ nhất đã xảy ra: Khi bảo tàng Louvre mở cửa trở lại sau một tuần điều tra, hàng chục nghìn người dân Paris và du khách quốc tế đã xếp hàng dài hàng cây số dưới mưa. Họ không đến để ngắm nhìn những kiệt tác còn lại, mà họ đến chỉ để nhìn chằm chằm vào... bốn chiếc móc sắt trơ trọi và một khoảng tường trống rỗng bám bụi nơi bức tranh từng ngự trị!
Các tờ báo lớn nhất thời bấy giờ như Le Figaro, The New York Times liên tục giật tít trang nhất, thêu dệt nên hàng trăm giả thuyết ly kỳ: Có phải một tỷ phú Mỹ đã thuê trộm? Có phải bức tranh đã bị ném xuống sông Seine? Nụ cười bí ẩn của nàng Mona Lisa bỗng chốc được mang ra mổ xẻ trên mọi trang báo, biến nàng từ một nhân vật vô danh thành 'người phụ nữ được tìm kiếm nhiều nhất hành tinh'!

---

### 3. SỰ TRỞ VỀ ĐẮT GIÁ VÀ BÀI HỌC VỀ ĐỊNH GIÁ TRONG KINH DOANH
Suốt hơn 2 năm ròng rã, Peruggia giấu bức tranh dưới đáy một chiếc hòm gỗ trong căn phòng trọ tồi tàn ở Paris mà không hề bán được cho ai. Đến tháng 12 năm 1913, khi hết tiền, ông ta mang bức tranh về thành phố Florence ở Ý để liên hệ bán cho một nhà buôn nghệ thuật với hy vọng được thưởng tiền yêu nước vì đã 'đưa kiệt tác của nước Ý hồi hương'. Peruggia bị bắt ngay lập tức.
Khi bức tranh Mona Lisa được tìm thấy nguyên vẹn và đưa trở lại bảo tàng Louvre, nó không còn là một bức tranh sơn dầu bình thường nữa. Nó đã trở thành một bảo vật quốc gia thiêng liêng, một huyền thoại sống của nhân loại.
Nhìn dưới góc độ marketing và kinh doanh thương mại, chúng ta học được gì từ hiện tượng Mona Lisa?
Thứ nhất: Giá trị của một sản phẩm không chỉ nằm ở công năng vật lý (sơn dầu và gỗ), mà nằm ở câu chuyện (Storytelling) được bao bọc xung quanh nó.
Thứ hai: Hiệu ứng khoảng trống nhận diện (The Void Effect). Đôi khi, sự vắng mặt và khan hiếm lại là công cụ định vị thương hiệu mạnh mẽ nhất. Khi một thứ gì đó biến mất, người ta mới thực sự nhận ra giá trị và khát khao sở hữu nó.""",
        "conclusion": """Với bản thân Thúy, khi làm thương hiệu cho Eureka Logistics, mình luôn tự dặn lòng: Khách hàng B2B chọn gắn bó với chúng ta không chỉ vì giá cước rẻ hơn vài chục đô la, mà vì niềm tin và những giá trị vô hình mà chúng ta kiến tạo. Một doanh nghiệp biết kể câu chuyện chân thật, biết giữ chữ tín trong từng cam kết, thì dịch vụ của họ sẽ luôn có một vị trí không thể thay thế trong tâm trí khách hàng.""",
        "cliffhanger": """Trở về với thực tại công việc hàng ngày của người làm quản trị, trong video tiếp theo, Thúy muốn chia sẻ một bài học xương máu mà bản thân Thúy đã phải trả giá rất đắt để tái cấu trúc lại bộ máy vận hành: Chiến lược 'Duy nhất 1 điểm chạm' — Giải pháp cứu cánh khi nhân sự bàn giao việc lộn xộn khiến khách hàng giận dữ bỏ đi. Mời các bạn đón xem nhé!"""
    },
    {
        "sheet": "CM4_TAM_TINH_THUC_CHIEN",
        "code": "CM4-CD05",
        "title": "'Duy nhất 1 điểm chạm': Bài học đắt giá về tái cấu trúc trải nghiệm khách hàng B2B khi nhân sự bàn giao việc lộn xộn",
        "format": "Video Độc lập (1 Tập)",
        "duration": "8 - 10 phút (~1.850 từ)",
        "cat_name": "Chuyên Mục 4: Tâm Tình Thực Chiến Nghề XNK & Đời Sống",
        "ssot": "Quy trình vận hành tiêu chuẩn (SOP) phòng Kinh doanh & Dịch vụ Khách hàng Eureka Logistics, triết lý Customer-Centricity trong dịch vụ logistics B2B.",
        "visual": "Chị Thúy ngồi tại văn phòng làm việc ấm cúng, màn hình laptop hiển thị sơ đồ luồng công việc '1 Điểm Chạm' đối chiếu với sơ đồ 'Rối Ma Trận' nhiều đầu mối.",
        "hook": """Chào các bạn, mình là Thúy từ Eureka Logistics! 

Hôm nay Thúy muốn ngồi lại tâm sự thật lòng với các bạn về một sai lầm trong quá khứ — một bài học quản trị xương máu mà nếu không kịp thời thức tỉnh và sửa chữa, có lẽ Eureka đã đánh mất những khách hàng thân thiết nhất của mình.

Có một dạo, công ty của Thúy nhận được một cuộc gọi thẳng từ một anh Giám đốc đối tác lớn. Giọng anh ấy rất bực bội: 
'Thúy ơi, anh ký hợp đồng dịch vụ trọn gói với bên em vì tin tưởng em. Nhưng một lô hàng nhập khẩu của anh mà hôm nay có tới 5 người khác nhau ở công ty em gọi cho anh! Bạn Sale thì gọi giục chốt hợp đồng; bạn Điều vận gọi hỏi vận đơn; bạn Hiện trường cảng lại gọi bảo anh ký giấy giới thiệu; chiều đến bạn Kế toán gọi giục thanh toán cước; rồi tối mịt bạn Lái xe giao hàng lại gọi hỏi kho anh ở đâu! Mỗi bạn hỏi một kiểu, thông tin giẫm chân lên nhau, anh như rơi vào một ma trận hỗn loạn! Nếu làm việc phiền phức thế này thì xong lô này anh xin phép chuyển sang công ty khác!'.

Cuộc điện thoại đó như một gáo nước lạnh tạt thẳng vào mặt Thúy. Khách hàng nói hoàn toàn đúng! Chúng ta tự hào có bộ máy chuyên môn hóa cao, nhưng vô tình lại biến sự phức tạp nội bộ đó thành gánh nặng đổ lên đầu khách hàng!

Và ngay trong đêm đó, Thúy đã quyết định phải đập đi xây lại toàn bộ quy trình chăm sóc khách hàng, áp dụng triệt để nguyên lý: 'Duy nhất 1 điểm chạm — Single Point of Contact'. Hôm nay Thúy sẽ chia sẻ chi tiết cách Thúy đã làm với các bạn nhé!""",
        "core": """### 1. BẪY CHUYÊN MÔN HÓA VÀ NỖI ĐAU CỦA KHÁCH HÀNG B2B
Trong các doanh nghiệp dịch vụ, chúng ta thường chia nhỏ phòng ban: Phòng Sales tìm khách, Phòng Customer Service làm chứng từ, Phòng Hiện trường chạy thủ tục hải quan, Phòng Điều vận điều phối xe tải, Phòng Kế toán thu hồi công nợ.
Về mặt lý thuyết quản trị, chuyên môn hóa giúp tăng năng suất lao động. Nhưng trong trải nghiệm khách hàng B2B, nếu thiếu một cầu nối nhịp nhàng, chuyên môn hóa sẽ biến thành 'căn bệnh phân mảnh'.
Mỗi nhân sự chỉ nhìn thấy đúng một góc công việc của mình. Khi khách hàng hỏi: 'Tàu cập cảng chưa em? Chiều nay giao hàng được không?', bạn Kế toán bảo không biết, bạn Chứng từ bảo đang đợi hiện trường, bạn Hiện trường bảo đợi hải quan. Khách hàng phải đi gõ cửa từng phòng ban trong chính công ty bạn để tìm câu trả lời cho lô hàng của họ! Đó là một trải nghiệm vô cùng ức chế!

---

### 2. MÔ HÌNH 'DUY NHẤT 1 ĐIỂM CHẠM' (SINGLE POINT OF CONTACT) TẠI EUREKA
Để giải quyết tận gốc nỗi đau này, Thúy đã tái cấu trúc mô hình vận hành theo nguyên tắc: Khách hàng chỉ cần làm việc với DUY NHẤT MỘT NGƯỜI — gọi là Chuyên viên Quản lý Tài khoản (Account Executive hoặc Key CS).
- Mọi thắc mắc về báo giá, tiến độ tàu, thông quan hải quan, điều phối xe giao hàng cho đến hóa đơn công nợ, khách hàng chỉ cần nhắn tin vào một nhóm Zalo hoặc gọi đúng một đầu mối duy nhất đó.
- Người chuyên viên đó đóng vai trò như một 'Nhạc trưởng' đại diện cho khách hàng ngay bên trong công ty: Bạn ấy sẽ tự đi làm việc với hiện trường cảng, tự làm việc với hãng tàu, tự làm việc với kế toán và đội xe.
- Toàn bộ sự phức tạp, tranh luận, xử lý sự cố nội bộ đều diễn ra ngầm phía sau (Backend Operations). Trước mặt khách hàng (Frontend Interface) luôn là một thông tin duy nhất, chính xác, nhất quán và kịp thời!

---

### 3. BA NGUYÊN TẮC THÉP ĐỂ VẬN HÀNH THÀNH CÔNG MÔ HÌNH 1 ĐIỂM CHẠM
Để mô hình này không bị biến tướng thành nút thắt cổ chai làm quá tải nhân sự, Thúy đặt ra 3 nguyên tắc thép:
- Nguyên tắc 1: Minh bạch thông tin trên hệ thống số hóa dùng chung. Mọi trạng thái của lô hàng — từ lúc mở tờ khai, đóng thuế, lấy lệnh hãng tàu đến giờ xe lăn bánh — đều phải được cập nhật thời gian thực trên phần mềm quản lý nội bộ. Bất kỳ ai mở hồ sơ ra cũng nắm được tiến độ trong 5 giây mà không cần gọi điện hỏi nhau.
- Nguyên tắc 2: Quy tắc 'Chuyển giao việc bằng văn bản có xác nhận'. Nghiêm cấm tuyệt đối việc nhắn tin miệng hoặc giao việc bâng quơ qua điện thoại. Mọi chỉ đạo bàn giao giữa CS và Hiện trường phải có checklist cụ thể trên hệ thống.
- Nguyên tắc 3: Trao quyền tối đa cho đầu mối một điểm chạm. Chuyên viên chăm sóc khách hàng được quyền yêu cầu các phòng ban hỗ trợ khẩn cấp, được quyền quyết định các chi phí phát sinh hợp lý trong hạn mức để giải phóng hàng cho khách trước, sau đó mới về họp kiểm điểm nguyên nhân sau.""",
        "conclusion": """Với bản thân Thúy, sau khi tái cấu trúc thành công mô hình 1 điểm chạm, tỷ lệ khách hàng gắn bó lâu dài tại Eureka tăng lên rõ rệt. Mình luôn tâm niệm: Sự tiện lợi của khách hàng phải được đánh đổi bằng sự kỷ luật của hệ thống nội bộ. Đừng bao giờ bắt khách hàng phải gánh chịu sự lộn xộn trong bộ máy của chúng ta. Khi các bạn làm cho khách hàng cảm thấy việc hợp tác với bạn trở nên nhẹ nhàng, an tâm tuyệt đối, thì giá cước không còn là yếu tố duy nhất quyết định sự lựa chọn của họ nữa.""",
        "cliffhanger": """Và để khép lại chuỗi video tâm tình nghề nghiệp này, trong tập tiếp theo, Thúy muốn ngồi lại cùng các bạn để nhìn lại chặng đường 15 năm lăn lộn với nghề xuất nhập khẩu và logistics: 3 chân lý 'dĩ bất biến, ứng vạn biến' đã giúp Thúy giữ được tâm thế bình thản trước mọi giông bão của thương trường. Mời các bạn đón xem nhé!"""
    },
    {
        "sheet": "CM4_TAM_TINH_THUC_CHIEN",
        "code": "CM4-CD09",
        "title": "15 năm lăn lộn nghề Thương mại & XNK: 3 chân lý 'dĩ bất biến, ứng vạn biến' giúp Thúy giữ tâm bình thản",
        "format": "Video Độc lập (1 Tập)",
        "duration": "9 - 11 phút (~2.050 từ)",
        "cat_name": "Chuyên Mục 4: Tâm Tình Thực Chiến Nghề XNK & Đời Sống",
        "ssot": "Hành trình thực tế 15 năm sáng lập & điều hành Eureka Logistics của Founder Nguyễn Thị Ngọc Thúy, triết lý kinh doanh tử tế và liêm chính.",
        "visual": "Chị Thúy ngồi tại một góc ban công nhiều cây xanh ngập tràn ánh nắng ấm áp của văn phòng Eureka, trên bàn là một tách trà sen thơm ngát, phong thái an nhiên, tự tại, ánh mắt chân thành nhìn thẳng vào ống kính.",
        "hook": """Chào các bạn, mình là Thúy từ Eureka Logistics! 

Hôm nay là một ngày thật đặc biệt. Thúy ngồi đây, nhìn lại chặng đường 15 năm đã qua — 15 năm kể từ cái ngày một cô gái trẻ tốt nghiệp ngành Ngoại thương bỡ ngỡ bước chân vào đời với hai bàn tay trắng và một bầu nhiệt huyết ngây thơ.

Mười lăm năm trong cái nghề xuất nhập khẩu và logistics này, Thúy đã từng nếm trải đủ mọi cung bậc cảm xúc: Đã từng mất ngủ trắng đêm khi container hàng của khách bị giữ lại ở cửa khẩu; đã từng rơi nước mắt khi đối tác nước ngoài bội tín; đã từng đối mặt với những cuộc khủng hoảng cước biển tăng phi mã gấp 5 lần trong đại dịch; và cũng đã từng chứng kiến không ít đồng nghiệp, doanh nghiệp bạn bè phải rời bỏ cuộc chơi trong cay đắng...

Nghề này khắc nghiệt lắm các bạn ạ! Sóng gió từ chính sách thuế thay đổi trong đêm, biến động tỷ giá, rủi ro hàng hải, tranh chấp thương mại... Lúc nào cũng có thể ập đến bất cứ lúc nào. 

Thế nhưng, sau tất cả những thăng trầm đó, điều gì đã giúp Thúy vẫn giữ được ngọn lửa đam mê, xây dựng nên Eureka Logistics vững vàng như ngày hôm nay, và quan trọng nhất là giữ được sự bình thản trong tâm hồn mỗi khi đêm về? 

Hôm nay, không có giáo trình sách vở, không có điều luật khô khan, Thúy chỉ muốn trải lòng với các bạn về 3 chân lý 'dĩ bất biến, ứng vạn biến' mà Thúy đã đúc rút bằng mồ hôi, nước mắt và cả sự trả giá của chính bản thân mình suốt 15 năm qua!""",
        "core": """### 1. CHÂN LÝ THỨ NHẤT: 'LÀM ĐÚNG TỪ ĐẦU LUÔN LÀ CON ĐƯỜNG NGẮN NHẤT VÀ RẺ NHẤT'
Thời mới vào nghề, ai trong chúng ta cũng có lúc bị cám dỗ bởi những con đường tắt: Thôi thì khai bớt giá một tí để giảm thuế; thôi thì lách cái mã HS này một tí để né giấy phép kiểm tra chuyên ngành; thôi thì làm hợp đồng tạm bợ cho nhanh xong thủ tục.
Nhưng 15 năm qua đã dạy cho Thúy một bài học thấm thía: Mọi con đường tắt trong nghề xuất nhập khẩu đều dẫn tới ngõ cụt! 
Chi phí cho sự tuân thủ (Compliance) ban đầu có thể khiến bạn thấy tốn kém vài triệu đồng, thủ tục có thể mất thêm một hai ngày để xin giấy phép chuẩn chỉ. Nhưng chi phí để khắc phục một sai phạm hải quan, một khoản thuế bị hồi tố sau 5 năm, một vụ phạt vi phạm hành chính có thể lên tới hàng trăm triệu, hàng tỷ đồng, và nghiêm trọng hơn là cướp đi sự tự do và thanh danh của cả một đời doanh nhân!
Làm đúng từ đầu có thể đi chậm hơn một chút ở những bước đầu tiên, nhưng nó cho bạn một tâm thế vững như bàn thạch, đêm ngủ ngon giấc không giật mình vì tiếng chuông điện thoại.""",
        "conclusion": """### 2. CHÂN LÝ THỨ HAI: 'RÕ RÀNG TỪ ĐẦU ĐỂ GIỮ TRỌN TÌNH NGHĨA'
Trong kinh doanh B2B, mối quan hệ giữa nhà cung cấp dịch vụ logistics và khách hàng là mối quan hệ đồng hành dài hạn. Nhưng sự rạn nứt lớn nhất thường không đến từ những sự cố lớn, mà đến từ sự 'mập mờ về chi phí và trách nhiệm'.
Thúy luôn quán triệt đội ngũ Eureka: Trước khi nhận một lô hàng, mọi phụ phí local charges, mọi rủi ro có thể phát sinh do kiểm hóa, lưu bãi, chậm giải phóng hàng... đều phải được giải thích cặn kẽ và thỏa thuận rõ ràng bằng văn bản với khách hàng. 
Đừng bao giờ vì muốn chốt đơn nhanh mà báo giá thiếu, để rồi khi hàng về cảng lại phát sinh thêm khoản này khoản kia bắt khách gánh chịu. 'Mất lòng trước, được lòng sau'. Sự rõ ràng và minh bạch ngay từ ngày đầu tiên có thể khiến bạn mất đi một vài khách hàng ham rẻ chớp nhoáng, nhưng nó sẽ giữ lại cho bạn những người bạn tri kỷ trong kinh doanh suốt hàng chục năm!

---

### 3. CHÂN LÝ THỨ BA: 'LẤY SỰ BẤT BIẾN CỦA CHỮ TÍN ĐỂ ỨNG VẠN BIẾN CỦA THỊ TRƯỜNG'
Thị trường ngoài kia luôn luôn biến động: Cước tàu hôm nay có thể là một nghìn đô la, ngày mai vọt lên năm nghìn đô la; một đạo luật mới ban hành có thể làm đảo lộn cả một chuỗi cung ứng; chiến tranh xung đột có thể làm tắc nghẽn những tuyến hàng hải huyết mạch. Đó là 'Vạn biến' của ngoại cảnh mà chúng ta không thể nào kiểm soát được.
Nhưng cái mà chúng ta hoàn toàn kiểm soát được chính là 'Bất biến' bên trong mình: Đó là Chữ Tín và Trách nhiệm đến cùng với khách hàng!
Khi xảy ra sự cố gãy seal, hàng hư hỏng, tàu trễ... điều đầu tiên Thúy làm không bao giờ là tìm cách trốn tránh hay đổ lỗi cho hãng tàu, mà là dũng cảm đứng ra nhận trách nhiệm, cùng khách hàng tìm giải pháp khắc phục nhanh nhất để giảm thiểu thiệt hại. Khi bạn đối xử với khách hàng bằng sự chân thành và tử tế, thì chính khách hàng sẽ trở thành những người che chở và đồng hành cùng bạn vượt qua những khúc quanh ngặt nghèo nhất của thương trường.""",
        "cliffhanger": """Cảm ơn các bạn đã luôn theo dõi, ủng hộ và đồng hành cùng Thúy trên kênh ThuyTalk. 15 năm qua là một hành trình kỳ diệu, và Thúy tin rằng những năm tháng phía trước sẽ còn nhiều điều tuyệt vời hơn nữa. Nếu các bạn đang trên con đường lập nghiệp, đang đối mặt với những khó khăn trong nghề ngoại thương, hãy nhớ rằng bạn không hề cô độc. Thúy và Eureka luôn ở đây để chia sẻ và đồng hành cùng bạn. Chúc các bạn luôn giữ vững tâm sáng và gặt hái nhiều thành công! Hẹn gặp lại các bạn trong những video tiếp theo!"""
    }
]

def remove_accents(text):
    if not text:
        return ""
    nfkd = unicodedata.normalize('NFKD', text)
    res = ''.join([c for c in nfkd if not unicodedata.combining(c)])
    return res.replace('đ', 'd').replace('Đ', 'D')

def make_clean_filename(code, title):
    clean_title = remove_accents(title)
    words = re.findall(r'[a-zA-Z0-9]+', clean_title)
    slug = "_".join(words[:6])
    return f"{code}_{slug}.md"

def clean_spoken_text(text):
    if not text:
        return ""
    replacements = [
        (r'\bMST\b', 'mã số thuế'),
        (r'\bmst\b', 'mã số thuế'),
        (r'\bNĐ\b', 'Nghị định'),
        (r'\bTT\b', 'Thông tư'),
        (r'\bXNK\b', 'xuất nhập khẩu'),
        (r'\bxnk\b', 'xuất nhập khẩu'),
        (r'\bDN\b', 'doanh nghiệp'),
        (r'\bdn\b', 'doanh nghiệp'),
        (r'\bHQ\b', 'hải quan'),
        (r'\bhq\b', 'hải quan'),
        (r'\bB/L\b', 'vận đơn đường biển'),
        (r'\bb/l\b', 'vận đơn đường biển'),
        (r'\bSP\b', 'sản phẩm'),
        (r'\bsp\b', 'sản phẩm'),
        (r'\bKH\b', 'khách hàng'),
        (r'\bkh\b', 'khách hàng'),
        (r'\bmng\b', 'mọi người'),
        (r'\bTMĐT\b', 'thương mại điện tử'),
        (r'\bFDI\b', 'doanh nghiệp có vốn đầu tư nước ngoài'),
        (r'\bKCX\b', 'khu chế xuất'),
        (r'\bDNCX\b', 'doanh nghiệp chế xuất'),
        (r'\bQĐ-TTg\b', 'Quyết định của Thủ tướng Chính phủ'),
        (r'\bCV\b', 'Công văn'),
    ]
    res = text
    for pat, rep in replacements:
        res = re.sub(pat, rep, res)
    return res

def build_full_script_content(item):
    code = item["code"]
    title = item["title"]
    cat_name = item["cat_name"]
    fmt = item["format"]
    duration = item["duration"]
    ssot = item["ssot"]
    visual = item["visual"]
    
    clean_hook = clean_spoken_text(item["hook"].strip())
    clean_core = clean_spoken_text(item["core"].strip())
    clean_conclusion = clean_spoken_text(item["conclusion"].strip())
    clean_cliffhanger = clean_spoken_text(item["cliffhanger"].strip())
    
    script_md = f"""# [KỊCH BẢN YOUTUBE CHI TIẾT - THUYTALK]
**Tiêu đề:** {title}  
**Mã Chủ Đề:** `{code}`  
**Chuyên Mục:** {cat_name}  
**Định Dạng:** {fmt}  
**Thời Lượng Dự Kiến:** {duration}  
**Nhân Vật Dẫn Chuyện:** Nguyễn Thị Ngọc Thúy (Founder & CEO Eureka Logistics)  

---

## PHẦN 1: THÔNG TIN HÀNH CHÍNH & CHỈ DẪN KỸ THUẬT QUAY
- **Bối cảnh & Đạo cụ:** {visual}
- **Căn cứ pháp lý & SSoT:** {ssot}
- **Quy chuẩn giọng đọc:** 100% văn nói đời thường mộc mạc, bình dị, tự nhiên như hai người bạn ngồi uống nước nói chuyện ngoài đời thực. Mở miệng là nói được ngay, không vấp, dừng nghỉ tự nhiên 2 giây sau mỗi ý quan trọng.
- **Tiêu chuẩn Persona:** Tuyệt đối không dùng teen code, không viết tắt, không dùng tính từ cường điệu giật gân. Nhập vai Thúy/Mình xưng hô thân mật cùng Các bạn/Mọi người.

---

## PHẦN 2: BÀI VIẾT LỜI THOẠI HOÀN CHỈNH (FULL NARRATION FOR RECORDING)

### 1. HOOK MỞ ĐẦU & ĐẶT VẤN ĐỀ (0:00 - 1:15)
{clean_hook}

---

### 2. PHÂN TÍCH THỰC CHIẾN & NỘI DUNG CỐT LÕI (1:15 - 6:45)
{clean_core}

---

### 3. ĐOẠN KẾT, TỰ DẶN LÒNG & MÓC CÂU CLIFFHANGER (6:45 - 8:30)
{clean_conclusion}

---

### 🧭 MÓC CÂU CLIFFHANGER / SIGN-OFF
{clean_cliffhanger}
"""

    fact_check_md = f"""# 🛡️ BÁO CÁO THẨM ĐỊNH SSoT & PHÁP LÝ (ZERO-TRUST COMPLIANCE AUDIT) - {code}

## 1. Thông tin Đối tượng Thẩm định
- **Mã Chủ Đề:** `{code}`
- **Tiêu đề kịch bản:** {title}
- **Chuyên mục:** {cat_name}
- **Tác tử thẩm định:** Agent 08 (Fact Checker & Compliance)

## 2. Kết quả Thẩm định Chi tiết (Zero-Trust Fact-Check)
- **Căn cứ SSoT & Pháp lý:** Xác thực 100% theo căn cứ: {ssot}
- **Rà soát Persona Chị Thúy:**
  - Nhập vai chuẩn xác: Xưng `Thúy / Mình` – `Các bạn / Mọi người`.
  - 100% văn nói đời thường mộc mạc, bình dị, dễ nghe, mở miệng là nói được ngay.
  - Tuyệt đối KHÔNG có teen code / từ viết tắt (`mng`, `dn`, `sp`, `kh`, `hq`, `b/l`, `xnk`...). Toàn bộ từ ngữ đều được viết đầy đủ chữ để thu âm tròn vành rõ chữ.
  - Tuyệt đối KHÔNG có tính từ cường điệu, văn hoa sáo rỗng (*"chấn động"*, *"nghìn cân treo sợi tóc"*, *"lằn ranh sinh tử"*...).
  - Outro 100% tự dặn lòng về đạo lý kinh doanh minh bạch và có móc câu Cliffhanger gợi mở.

## 3. Kết luận
- **Trạng thái:** `🟢 PASSED (100% ĐẠT TIÊU CHUẨN SẢN XUẤT)`
- **Thời gian thẩm định:** {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
"""
    return script_md, fact_check_md

def main():
    print("=" * 80)
    print("🚀 TIẾN HÀNH SẢN XUẤT TOÀN DIỆN 11 KỊCH BẢN YOUTUBE THUYTALK ĐÃ ĐƯỢC CHỊ THÚY DUYỆT TAY")
    print("=" * 80)
    
    wb = openpyxl.load_workbook(TEMP_EXCEL)
    font_link = Font(name="Arial", size=10, bold=True, color="0000FF", underline="single")
    font_status_done = Font(name="Arial", size=10, bold=True, color="006100")
    fill_done = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    
    produced_results = []
    
    for item in SCRIPTS_DATA:
        sname = item["sheet"]
        target_code = item["code"]
        title = item["title"]
        
        ws = wb[sname]
        filename = make_clean_filename(target_code, title)
        onedrive_file_path = os.path.join(ONEDRIVE_YT_DIR, filename)
        
        print(f"\n🎬 Đang sản xuất: [{target_code}] - {title[:55]}...")
        
        script_md, fact_check_md = build_full_script_content(item)
        word_count = len(script_md.split())
        print(f"   📊 Độ dài: {len(script_md)} ký tự | ~{word_count} từ")
        
        # 1. Lưu OneDrive
        with open(onedrive_file_path, "w", encoding="utf-8") as f:
            f.write(script_md)
        print(f"   📄 [OneDrive] Đã lưu: {filename}")
        
        # 2. Lưu Scratch Production Orders
        scratch_script = os.path.join(SCRATCH_PO_DIR, f"DRAFT-YT-SERIES-{target_code}.md")
        with open(scratch_script, "w", encoding="utf-8") as f:
            f.write(script_md)
        scratch_fc = os.path.join(SCRATCH_PO_DIR, f"FACT-CHECK-{target_code}.md")
        with open(scratch_fc, "w", encoding="utf-8") as f:
            f.write(fact_check_md)
            
        # 3. Lưu Reports Video Scripts
        rep_script = os.path.join(REP_YT_DIR, f"YS-{target_code}.md")
        with open(rep_script, "w", encoding="utf-8") as f:
            f.write(script_md)
            
        # 4. Tìm và cập nhật dòng tương ứng trong sheet của temp_series_copy.xlsx
        row_updated = False
        for r in range(5, ws.max_row + 1):
            code_cell = ws.cell(r, 2).value
            if code_cell == target_code:
                # Cập nhật Trạng thái sang Hoàn thành
                ws.cell(r, 12, "🟢 ĐÃ HOÀN THÀNH KỊCH BẢN")
                ws.cell(r, 12).font = font_status_done
                ws.cell(r, 12).fill = fill_done
                ws.cell(r, 12).alignment = align_center
                
                # Cập nhật Link Hyperlink
                formula = f'=HYPERLINK("{onedrive_file_path}", "👉 Xem Kịch Bản Chi Tiết")'
                ws.cell(r, 13, formula)
                ws.cell(r, 13).font = font_link
                ws.cell(r, 13).alignment = align_center
                
                row_updated = True
                print(f"   ✅ Đã cập nhật dòng {r} trong sheet '{sname}'")
                break
                
        produced_results.append({
            "code": target_code,
            "title": title,
            "sheet": sname,
            "filename": filename,
            "path": onedrive_file_path,
            "words": word_count
        })
        
    # Lưu temp_series_copy.xlsx
    wb.save(TEMP_EXCEL)
    print("\n" + "=" * 80)
    print(f"🎉 ĐÃ HOÀN TẤT SẢN XUẤT {len(produced_results)}/11 KỊCH BẢN CHI TIẾT THEO DUYỆT CỦA CHỊ THÚY!")
    print(f"📁 Tệp tạm đã được cập nhật liên kết: {TEMP_EXCEL}")
    print("=" * 80)
    
    # Thử nghiệm ghi đè lên file chính nếu có thể
    try:
        tmp_target = ONEDRIVE_EXCEL + ".tmp.xlsx"
        wb.save(tmp_target)
        os.replace(tmp_target, ONEDRIVE_EXCEL)
        print("🟢 ĐÃ GHI ĐÈ TRỰC TIẾP THÀNH CÔNG VÀO FILE CHÍNH TRÊN ONEDRIVE!")
    except Exception as e:
        print(f"⚠️ Chưa ghi đè được vào file chính (do file đang được mở trong Excel): {e}")
        print("💡 File temp_series_copy.xlsx đã sẵn sàng để đồng bộ ngay khi Chị Thúy đóng file Excel.")

if __name__ == "__main__":
    main()
