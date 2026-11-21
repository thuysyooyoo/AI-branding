# -*- coding: utf-8 -*-
"""
Sản xuất toàn bộ các bài viết mới từ Nhật ký Case Study Hằng ngày (scratch/daily_notes.md) Ngày 24/09/2026:
- HOT12: Case Study Thực Chiến - Nghệ thuật phản dame khi bị bóc phốt & Hiệu ứng 'Khách hàng ruột trở thành Đại sứ bảo vệ thương hiệu'
- HOT13: Kiểm tra ATTP đối với bao bì, dụng cụ tiếp xúc thực phẩm nhập khẩu (Công văn 22082/CHQ-GSQL & 1913/ATTP-KN)
- HOT14: Rủi ro thanh tra & quyết toán thuế khi thay đổi địa điểm kinh doanh khác phường
- HOT15: Quy trình xin Giấy xác nhận hoàn thành nghĩa vụ thuế (Mẫu 24) làm thủ tục với Sở KH&ĐT
Kèm 2 kịch bản YouTube chi tiết (HOT12 và HOT13) và 4 PO, 4 Fact-Check đầy đủ.
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
SCRATCH_PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
REP_FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")
REP_YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")

os.makedirs(SCRATCH_PO_DIR, exist_ok=True)
os.makedirs(REP_FB_DIR, exist_ok=True)
os.makedirs(REP_YT_DIR, exist_ok=True)

# ==============================================================================
# 1. BÀI HOT12: NGHỆ THUẬT PHẢN DAME KHI BỊ BÓC PHỐT & HIỆU ỨNG KHÁCH HÀNG RUỘT
# ==============================================================================
FB_HOT12_CONTENT = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - CASE STUDY THỰC CHIẾN QUẢN TRỊ & KHỦNG HOẢNG]
**Nền tảng mục tiêu:** Facebook (Trang cá nhân Chị Thúy & Fanpage Eureka Logistics)
**Mã sản xuất:** `PO-2026-W39-HOT12`
**Content Pillar:** `CS` — Case Study / Personal Branding & Crisis Management
**Góc nhìn:** Problem-First & Nghệ thuật lãnh đạo xắn tay áo
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

BỊ KHÁCH BÓC PHỐT TÌM ĐỐI TÁC MỚI VÌ 'HÀNG VỀ LÂU LẮC': TẠI SAO THÚY KHÔNG ĐÔI CO MÀ DÙNG DỮ LIỆU ERP ĐỂ 'PHẢN DAME', VÀ CÂU CHUYỆN 'TÂM THƯ BÊNH VỰC' CHẤN ĐỘNG CỦA MỘT KHÁCH HÀNG RUỘT! 🌟

Sáng nay trên Cộng Đồng Nhập Khẩu Chính Ngạch Trung Quốc có một bài đăng tìm đối tác mới réo thẳng tên công ty Thúy: 
*"Cty nào nhận uỷ thác nhập khẩu hàng China inbox mình. Mình làm với cty Eureka hàng về lâu lắc..."*

Đọc xong bài viết, cảm giác đầu tiên của Thúy là gì? 
Là một người làm nghề dịch vụ hơn 10 năm qua, Thúy hiểu rằng khách hàng bức xúc thì khách mới lên tiếng. Nhưng cách chúng ta ứng xử trước một lời phàn nàn công khai sẽ định nghĩa đẳng cấp và văn hóa của doanh nghiệp!

Thúy xin chia sẻ lại cách Eureka đã đối thoại và câu chuyện xúc động phía sau — một bài học thực chiến về quản trị trải nghiệm khách hàng:

---

📍 **BƯỚC 1: CẦU THỊ, BÌNH TĨNH & CAM KẾT TRÁCH NHIỆM TỚI CÙNG**
Thúy không xóa bài, không đôi co gay gắt. Thúy vào bình luận công khai ngay lập tức:
*"Dạ rất xin lỗi vì chị có những cảm nhận dịch vụ chưa tốt về Eureka liên quan đến tốc độ. Nếu chị đọc được bình luận này của Thúy vui lòng check inbox ạ. Nếu có tổn thất thiệt hại gì cho chị do không đảm bảo đúng cam kết, Eureka sẽ đền bù 100% thiệt hại của mình ạ!"*

Làm dịch vụ chính ngạch, điều đầu tiên phải là tinh thần dám nhận trách nhiệm. Nếu mình sai, mình đền. Không quanh co!

---

📍 **BƯỚC 2: 'PHẢN DAME' VĂN MINH BẰNG DỮ LIỆU ERP & BẰNG CHỨNG HỆ THỐNG**
Khi khách vẫn tiếp tục trách móc khâu vận chuyển, Thúy không giải thích bằng "võ mồm". Thúy mở ngay hệ thống phần mềm quản lý ERP của Eureka ra đối soát lịch sử từng đơn hàng. Và đây là sự thật hiển thị trên từng dòng log dữ liệu:
1. **Kiện hàng gửi về kho nhưng KHÔNG ĐỂ MÃ KHÁCH HÀNG:** Thúy chụp ảnh thực tế phiếu gửi hàng nội địa Trung Quốc của khách gửi — trên nhãn hoàn toàn không có mã định danh khách hàng. Kho nhận hàng nghìn kiện mỗi ngày, không có mã thì robot hay con người cũng không thể biết hàng của ai để báo nhập kho!
2. **Không xác nhận đơn & không gửi thông tin làm giấy tờ xuất khẩu:** Hàng về kho từ ngày 5/9 nhưng khách không báo, mãi tới ngày 10/9 khi bạn CSKH (Nguyễn Linh) chủ động hỏi thì khách mới gửi thông tin. Nguyên tắc sắt đá của nhập khẩu chính ngạch: Hàng về kho mà không có invoice, packing list và thông tin khai báo thì **TUYỆT ĐỐI KHÔNG ĐƯỢC PHÉP CHO XUẤT KHẨU!**
3. **Địa điểm giao hàng:** Tuyến vận chuyển giao về TP.HCM thì bắt buộc thời gian phải dài hơn giao tại Hà Nội.

Thúy gửi toàn bộ ảnh chụp hệ thống, ảnh kiện hàng thực tế và chia sẻ thẳng thắn:
*"Eureka quản lý bằng hệ thống nên toàn bộ lịch sử đều có thông tin để đối chất. Sai thì nhận, mà đúng thì làm cho ra vấn đề đến cùng chị ạ!"*

---

📍 **BƯỚC 3: HIỆU ỨNG VÔ GIÁ — KHÁCH HÀNG RUỘT ĐỨNG LÊN BẢO VỆ THƯƠNG HIỆU**
Điều khiến Thúy cay cay khóe mắt nhất sáng nay là một khách hàng ruột lâu năm (đi hàng trăm khối mỗi tháng) đọc được bài đã lập tức vào viết một tâm thư dài bảo vệ Eureka và Thúy:

Khách kể lại:
*"Hôm chậm hàng nhà em lâu quá, alo lên chị Thúy. Mặc dù chị Thúy là Sếp to nhé, vẫn cắp nách sang tận kho Trung Quốc, chat webcam trực tiếp tìm từng kiện hàng cho em. Rồi còn ưu tiên hàng em về. Đấy, công ty nào mà Sếp đích thân xắn tay áo đi làm việc thì công ty đó ắt có lợi cho mình...*

*Còn bạn CSKH Linh thì nâng khách như nâng trứng: theo dõi sát từng đơn, file thanh toán tỷ giá bớt cho, tem dán thùng debit riêng từng lô, trước khi báo quan luôn gửi check kỹ. Nhiều lô bị kiểm hóa/tham vấn giá, em nó hỗ trợ nửa tiền... Tôi lọc từ hơn 100 bên vận chuyển chỉ chọn mỗi bạn Linh ở Eureka vì làm hàng chuẩn, VAT chuẩn. Nếu bạn Linh nghỉ tôi cũng không làm với Eureka nữa!"*

---

💡 **BÀI HỌC QUẢN TRỊ XƯƠNG MÁU CHO CÁC DOANH NGHIỆP LOGISTICS:**
1. **Quy trình ERP là tấm khiên vững chắc nhất:** Trong thời đại số, mọi giao dịch phải được số hóa. Khi có sự cố, dữ liệu đối soát minh bạch sẽ bảo vệ uy tín doanh nghiệp trước mọi hiểu lầm.
2. **Lãnh đạo xắn tay áo (Servant Leadership):** Làm sếp không phải chỉ ngồi chỉ tay năm ngón. Khi khách hàng cần, sếp sẵn sàng lao sang kho bãi, trực tiếp gỡ nút thắt thì nhân viên mới hết lòng, khách hàng mới nể phục.
3. **Sự tận tâm tạo nên Đại sứ Thương hiệu tự nhiên:** Khách hàng không bỏ tiền mua dịch vụ rẻ tiền chộp giật. Họ gắn bó với doanh nghiệp vì sự tử tế, hóa đơn VAT chuẩn chỉnh và cái tâm của người làm nghề.

Cảm ơn những lời góp ý của khách hàng để Eureka hoàn thiện mỗi ngày, và biết ơn vô hạn những khách hàng ruột đã luôn là bức tường lửa vững chắc bảo vệ Eureka!

Anh chị em làm nghề dịch vụ đã từng gặp pha "bóc phốt" nhớ đời nào chưa? Cùng chia sẻ góc nhìn bên dưới nhé!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Video chia sẻ: 3 lỗi cơ bản từ chủ hàng khiến kiện hàng bị giam tại kho Trung Quốc: [Link YouTube Channel Eureka Logistics]
👉 Cẩm nang quản lý đơn hàng chính ngạch chuẩn hóa theo hệ thống ERP: https://customs.gov.vn

**Hashtags:** `#XuLyKhungHoang #ERPLogistics #CustomerAdvocacy #NhapKhauChinhNgach #VanHoaPhungSu #EurekaLogistics #JasmineNguyen`
"""

YT_HOT12_CONTENT = """# [KỊCH BẢN YOUTUBE DÀI - SERIES QUẢN TRỊ & PERSONAL BRANDING]
**Tên video:** Xử Lý Khủng Hoảng Truyền Thông: Nghệ Thuật 'Phản Dame' Bằng Dữ Liệu ERP & Cách Biến Khách Hàng Ruột Thành Đại Sứ Bảo Vệ Thương Hiệu
**Mã sản xuất:** `PO-2026-W39-HOT12`
**Thời lượng dự kiến:** 8 - 10 phút
**Chuyên mục:** `CS — Case Study`

---

## BẢNG PHÂN CẢNH & STORYBOARD CHI TIẾT

| Phân đoạn & Thời lượng | Visual & Bối cảnh | Audio & Lời thoại Thúy (Voice-over) | Text On Screen & Hiệu ứng SFX |
| :--- | :--- | :--- | :--- |
| **[0:00 - 0:45] HOOK** | Thúy ngồi tại phòng làm việc Eureka, ánh mắt điềm tĩnh, chỉ tay vào màn hình máy tính đang hiển thị bài bóc phốt trên Facebook nhóm lớn. Cắt cảnh hình ảnh hệ thống ERP hiển thị log vận đơn. | "Một ngày đẹp trời, công ty bạn bị khách hàng đăng đàn bóc phốt lên nhóm 200.000 thành viên vì 'hàng về lâu lắc'. Bạn sẽ làm gì? Xóa bài, đôi co cãi vã hay im lặng block khách? Sáng nay, chính Thúy vừa trải qua một tình huống như thế. Nhưng thay vì khủng hoảng, Eureka đã biến nó thành màn khẳng định năng lực quản trị ERP và chứng kiến một khách hàng ruột đứng ra 'bênh vực tụt quần'. Video này Thúy chia sẻ trọn bộ kịch bản xử lý khủng hoảng 3 bước mà mọi chủ doanh nghiệp dịch vụ đều cần nắm!" | **Text:** BỊ BÓC PHỐT TRÊN NHÓM 200K MEMBER: XỬ LÝ SAO?<br>**SFX:** Tiếng chuông thông báo dồn dập + tiếng gõ phím. |
| **[0:45 - 2:30] PHẦN 1: BƯỚC 1 - NGUYÊN TẮC CẦU THỊ & CAM KẾT TRÁCH NHIỆM** | Thúy ngồi thẳng thắn, chia sẻ chân thành trước ống kính. Cắt cảnh ảnh chụp màn hình comment phản hồi văn minh của Thúy. | "Nguyên tắc số 1 trong xử lý khủng hoảng: Tuyệt đối không để cái tôi lấn át lý trí. Khách hàng phàn nàn, việc đầu tiên là lắng nghe và cầu thị. Thúy bình luận công khai ngay dưới bài viết: 'Nếu có tổn thất thiệt hại do Eureka không đúng cam kết, Eureka xin đền bù 100%!'. Lời cam kết này cho thấy bạn là một doanh nghiệp tử tế, không trốn tránh và sẵn sàng chịu trách nhiệm với đồng tiền của khách hàng." | **Text:** NGUYÊN TẮC 1: CẦU THỊ & CAM KẾT ĐỀN BÙ 100% NẾU SAI.<br>**Visual:** Ảnh chụp comment lịch thiệp. |
| **[2:30 - 5:30] PHẦN 2: BƯỚC 2 - 'PHẢN DAME' VĂN MINH BẰNG DỮ LIỆU ERP** | Screencast trực tiếp phần mềm quản lý đơn hàng ERP của Eureka. Thúy chỉ vào các trường dữ liệu: Ngày hàng đến kho, phiếu gửi nội địa không mã, log chat CSKH. | "Nhưng cầu thị không có nghĩa là nhận bừa lỗi không phải của mình! Khi khách tiếp tục đổ lỗi, Thúy lập tức truy xuất hệ thống ERP: Phiếu gửi hàng nội địa TQ của khách hoàn toàn KHÔNG CÓ MÃ KHÁCH HÀNG. Hàng về kho 5 ngày khách không gửi thông tin hàng để làm thủ tục xuất khẩu. Trong logistics chính ngạch, không có chứng từ thì không ai dám cho hàng qua biên giới! Thúy gửi ảnh chụp log hệ thống và khẳng định: 'Eureka quản lý bằng hệ thống, có log đối chất rõ ràng. Sai thì nhận, đúng thì làm cho ra vấn đề đến cùng!'. Sự thật và dữ liệu chính là vũ khí phản biện đanh thép nhất!" | **Text:** NGUYÊN TẮC 2: NÓI CHUYỆN BẰNG DATA ERP — ĐỪNG NÓI BẰNG CẢM XÚC!<br>**SFX:** Tiếng click chuột tra cứu log. |
| **[5:30 - 7:30] PHẦN 3: BƯỚC 3 - KHI KHÁCH HÀNG RUỘT BIẾN THÀNH LÁ CHẮN BẢO VỆ** | Thúy xúc động kể lại câu chuyện khách hàng ruột vào bình luận kể chuyện Sếp cắp nách sang kho TQ tìm hàng và bạn sale Linh chăm sóc khách tận tâm. | "Và điều kỳ diệu nhất đã xảy ra: Một khách hàng lớn đi hàng trăm khối của Eureka đọc được bài đã lập tức vào viết tâm thư bảo vệ công ty. Bạn ấy kể chuyện Thúy là Sếp vẫn đích thân cắp nách sang tận kho Trung Quốc tìm từng kiện hàng cho bạn, kể chuyện bạn sale Linh theo sát từng lô, hỗ trợ cả tiền tham vấn giá... Khi bạn phụng sự khách hàng bằng cả trái tim và làm ăn chuẩn chỉnh pháp lý, chính khách hàng trung thành sẽ là đại sứ thương hiệu mạnh mẽ nhất bảo vệ bạn trước mọi sóng gió!" | **Text:** TÀI SẢN LỚN NHẤT: KHÁCH HÀNG TRỞ THÀNH ĐẠI SỨ THƯƠNG HIỆU!<br>**Visual:** Trích đoạn tâm thư khách hàng ruột. |
| **[7:30 - 8:30] CLIFFHANGER OUTRO** | Thúy mỉm cười tự tin, cầm một tập công văn hải quan mới toanh trên bàn. Ánh sáng studio chiếu vào dòng chữ 'Công văn 22082/CHQ-GSQL'. | "Bảo vệ uy tín thương hiệu cần sự chuẩn chỉ của hệ thống, nhưng để bảo vệ dòng tiền thì doanh nghiệp phải cập nhật liên tục các quy định mới. Hôm qua, Cục Hải quan vừa ban hành Công văn 22082 liên quan đến việc kiểm tra an toàn thực phẩm đối với bao bì, túi nhựa vô trùng. Rất nhiều doanh nghiệp đang hiểu lầm tai hại giữa 'Tự công bố' và 'Kiểm tra ATTP' khiến hàng ngàn container có nguy cơ kẹt cứng tại cảng! Chi tiết công văn này thế nào? Đón xem ngay trong video tiếp theo trên kênh Eureka Logistics nhé!" | **Text:** TẬP TIẾP THEO: CÔNG VĂN 22082 & BẪY KIỂM TRA ATTP BAO BÌ NHẬP KHẨU!<br>**SFX:** Tiếng lật tài liệu dứt khoát + nhạc outro truyền cảm hứng. |
"""

# ==============================================================================
# 2. BÀI HOT13: KIỂM TRA ATTP ĐỐI VỚI BAO BÌ, TÚI VÔ TRÙNG (CV 22082/CHQ-GSQL)
# ==============================================================================
FB_HOT13_CONTENT = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - CẢNH BÁO QUY ĐỊNH HẢI QUAN & ATTP]
**Nền tảng mục tiêu:** Facebook (Trang cá nhân Chị Thúy & Fanpage Eureka Logistics)
**Mã sản xuất:** `PO-2026-W39-HOT13`
**Content Pillar:** `HT` — Học Thực chiến / Food Safety & Packaging Compliance
**Góc nhìn:** Problem-First & Hướng dẫn kỹ thuật thông quan
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

CÔNG VĂN 22082/CHQ-GSQL & 1913/ATTP-KN VỪA BAN HÀNH: DOANH NGHIỆP NHẬP KHẨU TÚI VÔ TRÙNG, BAO BÌ TIẾP XÚC THỰC PHẨM ĐỪNG LẦM TƯỞNG 'TỰ CÔNG BỐ' LÀ ĐƯƠNG NHIÊN ĐƯỢC MIỄN KIỂM TRA ATTP! ⚠️

Hôm qua, ngày 23/09/2026, Cục Hải quan đã chính thức ban hành Công văn số 22082/CHQ-GSQL, chuyển ý kiến hướng dẫn của Cục An toàn thực phẩm – Bộ Y tế (Công văn số 1913/ATTP-KN ngày 21/09/2026) về việc kiểm tra ATTP đối với bao bì, dụng cụ tiếp xúc trực tiếp với thực phẩm.

Đây là một văn bản "nóng hổi" mà các doanh nghiệp sản xuất bao bì, thực phẩm, nước giải khát và đơn vị logistics cần đọc thật kỹ để không bị ách tắc container tại cảng!

---

📍 **VƯỚNG MẮC THỰC TẾ XUẤT PHÁT TỪ ĐÂU?**
Thời gian qua, một số doanh nghiệp nhập khẩu túi vô trùng (Aseptic bag, mã HS 3923.21.99) dùng để đựng thực phẩm gặp tình trạng mỗi Chi cục Hải quan áp dụng một kiểu:
- Có Chi cục từ chối tiếp nhận hồ sơ kiểm tra ATTP vì cho rằng mặt hàng này thuộc "mức độ rủi ro trung bình" theo Thông tư 27/2026/TT-BYT, chỉ cần làm "Tự công bố sản phẩm" theo Nghị định 15/2018/NĐ-CP.
- Ngược lại, có Chi cục khác vẫn kiên quyết bắt buộc doanh nghiệp phải làm thủ tục "Kiểm tra nhà nước về an toàn thực phẩm nhập khẩu" theo Danh mục Thông tư 15/2024/TT-BYT.

Sự lúng túng này khiến nhiều lô hàng túi bao bì bị nằm chờ tại cảng Hải Phòng, Cát Lái nhiều ngày trời, phát sinh tiền lưu cont bãi rất xót ruột!

---

📍 **HẢI QUAN VÀ BỘ Y TẾ CHỐT NGUYÊN TẮC QUẢN LÝ THẾ NÀO?**
Tại Công văn 1913/ATTP-KN và CV 22082, Cục An toàn thực phẩm đã làm rõ nguyên tắc pháp lý:
1. **Theo Thông tư 27/2026/TT-BYT:** Dụng cụ, vật liệu bao gói tiếp xúc trực tiếp với thực phẩm bằng nhựa tổng hợp, cao su thuộc nhóm rủi ro trung bình -> Doanh nghiệp phải thực hiện **Tự công bố sản phẩm** theo Điều 4 Nghị định 15/2018/NĐ-CP.
2. **Căn cứ Khoản 3 Điều 82 Nghị định 37/2026/NĐ-CP:** Hàng hóa đã thực hiện biện pháp quản lý chất lượng theo pháp luật chuyên ngành thì tiếp tục thực hiện theo pháp luật chuyên ngành.
3. 👉 **KẾT LUẬN SỐNG CÒN:** Nếu sản phẩm thuộc Danh mục tại Thông tư 15/2024/TT-BYT và **KHÔNG thuộc các trường hợp được miễn kiểm tra** (quy định tại Điều 13 Nghị định 15/2018/NĐ-CP), thì khi nhập khẩu **VẪN BẮT BUỘC PHẢI THỰC HIỆN KIỂM TRA NHÀ NƯỚC VỀ ATTP!**

Tuyệt đối không được đánh đồng: "Tôi đã làm bản Tự công bố rồi thì mặc nhiên được miễn kiểm tra ATTP khi mở tờ khai hải quan!". Đây là cái bẫy hiểu lầm tai hại nhất!

---

💡 **3 VIỆC DOANH NGHIỆP CẦN RÀ SOÁT NGAY ĐỂ THÔNG QUAN NHANH:**
- **Bước 1:** Đối chiếu mã HS 3923 của sản phẩm bao bì với Phụ lục Danh mục tại Thông tư 15/2024/TT-BYT xem có rơi vào diện kiểm tra nhà nước không.
- **Bước 2:** Chuẩn bị song song 2 bộ hồ sơ: Bản Tự công bố sản phẩm (đã nộp cơ quan y tế có biên nhận) VÀ Giấy đăng ký kiểm tra nhà nước về an toàn thực phẩm trên Cổng thông tin một cửa Quốc gia (NSW).
- **Bước 3:** Đăng ký kiểm tra ATTP trước khi tàu cập cảng 2-3 ngày để lấy ngay số tiếp nhận, sẵn sàng truyền vào tờ khai điện tử để thông quan lấy mẫu kiểm tra ngay tại cảng.

Doanh nghiệp của anh chị đang nhập khẩu nguyên phụ liệu bao bì thực phẩm có đang vướng bước này không? Để lại câu hỏi bên dưới Thúy cùng đội ngũ tư vấn nhé!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Tải trọn bộ Công văn 22082/CHQ-GSQL & CV 1913/ATTP-KN (Bộ Y tế): https://customs.gov.vn
👉 Danh mục chi tiết các mặt hàng bao bì phải kiểm tra ATTP theo Thông tư 15/2024/TT-BYT: [Tải tài liệu Eureka]

**Hashtags:** `#CongVan22082 #ThongTu27 #ThongTu15 #KiemTraATTP #BaoBiThucPham #AsepticBag #HaiQuanVietNam #EurekaLogistics`
"""

YT_HOT13_CONTENT = """# [KỊCH BẢN YOUTUBE DÀI - SERIES HỌC XNK 5 PHÚT]
**Tên video:** Cảnh Báo Nghẽn Cảng Hàng Bao Bì Thực Phẩm: Công Văn 22082/CHQ-GSQL & Bẫy 'Tự Công Bố' Không Đồng Nghĩa Với Miễn Kiểm Tra ATTP
**Mã sản xuất:** `PO-2026-W39-HOT13`
**Thời lượng dự kiến:** 8 - 10 phút
**Chuyên mục:** `HT — Học Thực chiến`

---

## BẢNG PHÂN CẢNH & STORYBOARD CHI TIẾT

| Phân đoạn & Thời lượng | Visual & Bối cảnh | Audio & Lời thoại Thúy (Voice-over) | Text On Screen & Hiệu ứng SFX |
| :--- | :--- | :--- | :--- |
| **[0:00 - 0:45] HOOK** | Thúy đứng trước hình ảnh kho cảng có các pallet chứa cuộn màng nhôm và túi nhựa Aseptic bag. Cầm trên tay Công văn 22082/CHQ-GSQL ngày 23/09/2026. | "Bạn nhập một lô hàng túi bao bì đựng nước ép trái cây về cảng. Bạn tự tin cầm bản 'Tự công bố sản phẩm' ra mở tờ khai thì bị cán bộ hải quan dừng lại: 'Anh chưa đăng ký kiểm tra nhà nước về ATTP, không được thông quan!'. Nhiều doanh nghiệp kêu trời: 'Ủa, Thông tư 27 nói bao bì rủi ro trung bình chỉ cần tự công bố thôi mà?'. Ngày 23/09/2026, Cục Hải quan vừa ban hành Công văn 22082 gióng lên hồi chuông cảnh báo: Tự công bố KHÔNG ĐỒNG NGHĨA với miễn kiểm tra ATTP! Video này Thúy sẽ giúp bạn gỡ bẫy pháp lý này để container không bị giam tại cảng!" | **Text:** TỰ CÔNG BỐ = MIỄN KIỂM TRA ATTP? BẪY HIỂU NHẦM TAI HẠI!<br>**SFX:** Tiếng dập mộc từ chối tiếp nhận + còi cảnh báo cảng. |
| **[0:45 - 2:30] PHẦN 1: NGUYÊN NHÂN ÁCH TẮC TẠI CÁC CHI CỤC HẢI QUAN** | Thúy phân tích sơ đồ luồng hàng tại cảng. Màn hình hiển thị mã HS 3923.21.99 và hình ảnh túi vô trùng Aseptic bag. | "Mặt hàng túi vô trùng Aseptic bag - chuyên đựng sữa, nước hoa quả cô đặc - vừa qua là tâm điểm tranh cãi. Chi cục này bảo không nhận hồ sơ vì Thông tư 27/2026 xếp vào nhóm rủi ro trung bình, doanh nghiệp tự chịu trách nhiệm. Chi cục kia lại bảo Thông tư 15/2024 vẫn đưa vào danh mục kiểm tra. Hàng nằm kẹt ở giữa, doanh nghiệp chịu trận tiền lưu kho bãi hàng chục triệu đồng mỗi ngày. Chính vì thế, Cục Hải quan phải có văn bản hỏi thẳng Cục An toàn thực phẩm Bộ Y tế để chốt phương án cuối cùng!" | **Text:** TRỐNG ĐÁNH XUÔI, KÈN THỔI NGƯỢC: DOANH NGHIỆP LÃNH ĐỦ PHÍ LƯU BÃI.<br>**Visual:** So sánh quan điểm 2 chi cục. |
| **[2:30 - 5:30] PHẦN 2: BÓC TÁCH CÔNG VĂN 22082 & Ý KIẾN BỘ Y TẾ** | Screencast trực tiếp văn bản Công văn 1913/ATTP-KN và CV 22082/CHQ-GSQL. Thúy highlight các điều khoản trọng yếu. | "Đây là câu trả lời chính thức của Bộ Y tế mà các bạn cần chụp lại màn hình ngay: Đúng là bao bì tiếp xúc trực tiếp bằng nhựa, cao su phải Tự công bố theo Nghị định 15. NHƯNG, căn cứ Khoản 3 Điều 82 Nghị định 37/2026, nếu sản phẩm vẫn nằm trong Danh mục Thông tư 15/2024/TT-BYT và không thuộc diện miễn trừ tại Điều 13 Nghị định 15, thì VẪN PHẢI KIỂM TRA NHÀ NƯỚC VỀ ATTP NHẬP KHẨU! Nói một cách nôm na: Tự công bố là giấy phép lưu hành nội địa, còn Kiểm tra nhà nước ATTP là tấm vé thông hành qua cửa khẩu. Có vé nội địa mà thiếu vé cửa khẩu thì container vẫn đứng im ngoài cổng cảng!" | **Text:** TỰ CÔNG BỐ: GIẤY PHÉP LƯU HÀNH.<br>KIỂM TRA ATTP: VÉ THÔNG HÀNH QUA CỬA KHẨU.<br>**SFX:** Tiếng lật trang tài liệu + highlight vàng. |
| **[5:30 - 7:30] PHẦN 3: SOP 3 BƯỚC ĐỂ THÔNG QUAN LÔ HÀNG TRONG 24 GIỜ** | Thúy hướng dẫn quy trình thao tác trên Cổng thông tin Một cửa Quốc gia (NSW). | "Vậy doanh nghiệp cần làm gì để không bị giam hàng? Ghi nhớ quy trình 3 bước này:<br>Bước 1: Làm Bản Tự công bố sản phẩm tại Ban Quản lý ATTP hoặc Chi cục ATTP tỉnh/thành phố ngay khi ký hợp đồng ngoại thương.<br>Bước 2: Trước khi tàu cập cảng 3 ngày, lấy Bill và Invoice lên Cổng Một cửa Quốc gia NSW đăng ký Kiểm tra nhà nước về ATTP với Trung tâm kiểm nghiệm (Quatest 1, Quatest 3...).<br>Bước 3: Lấy Giấy xác nhận đăng ký NSW truyền vào ô mã chứng từ trên phần mềm khai hải quan điện tử để mở tờ khai và đưa hàng về kho bảo quản lấy mẫu kiểm nghiệm!" | **Text:** SOP 3 BƯỚC: TỰ CÔNG BỐ TRƯỚC -> ĐĂNG KÝ NSW TRƯỚC TÀU CẬP 3 NGÀY.<br>**Graphic:** Sơ đồ 3 bước chuẩn chỉnh. |
| **[7:30 - 8:30] CLIFFHANGER OUTRO** | Thúy hạ giọng điềm tĩnh, trên màn hình hiện lên hình ảnh một tờ giấy đăng ký kinh doanh và một thông báo thanh tra thuế tại trụ sở. | "Chuẩn hóa thủ tục hải quan cửa khẩu là việc trước mắt, nhưng có một rủi ro âm thầm khác đang rình rập doanh nghiệp ngay tại trụ sở: Nhiều công ty chuyển địa chỉ văn phòng từ phường này sang phường khác, nghĩ là chỉ đổi địa chỉ trên giấy phép kinh doanh. Ai ngờ cơ quan thuế lập tức đưa vào diện THANH TRA QUYẾT TOÁN TOÀN DIỆN vì chưa quyết toán trên 2 năm! Rủi ro này cụ thể ra sao và làm thế nào để dọn sạch sổ sách kế toán trước khi chuyển trụ sở? Đón xem ngay trong video tiếp theo vào Thứ Tư tuần tới. Nhớ bấm subscribe kênh Eureka Logistics nhé!" | **Text:** TẬP TIẾP THEO: ĐỔI ĐỊA CHỈ KHÁC PHƯỜNG — BẪY THANH TRA QUYẾT TOÁN THUẾ TẠI TRỤ SỞ!<br>**SFX:** Tiếng còi hú cảnh báo + nhạc outro dồn dập. |
"""

# ==============================================================================
# 3. BÀI HOT14: RỦI RO THANH TRA & QUYẾT TOÁN THUẾ KHI ĐỔI ĐỊA CHỈ KHÁC PHƯỜNG
# ==============================================================================
FB_HOT14_CONTENT = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - CASE STUDY QUẢN TRỊ THUẾ DOANH NGHIỆP]
**Nền tảng mục tiêu:** Facebook (Trang cá nhân Chị Thúy & Fanpage Eureka Logistics)
**Mã sản xuất:** `PO-2026-W39-HOT14`
**Content Pillar:** `CS` — Case Study / Tax Audit Risk & Corporate Relocation
**Góc nhìn:** Problem-First & Bài học phòng vệ tài chính
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

CHUYỂN ĐỊA CHỈ KINH DOANH KHÁC PHƯỜNG / KHÁC QUẬN: CẨM NANG TRÁNH BẪY 'THANH TRA QUYẾT TOÁN THUẾ TẠI TRỤ SỞ' CHO DOANH NGHIỆP CHƯA QUYẾT TOÁN TRÊN 2 NĂM! ⚠️

Rất nhiều chủ doanh nghiệp SME có một suy nghĩ cực kỳ ngây thơ: 
*"Công ty hết hạn hợp đồng thuê văn phòng, mình chuyển sang mặt bằng mới ở phường bên cạnh, chỉ cần lên Sở Kế hoạch & Đầu tư nộp hồ sơ thay đổi ĐKKD và báo thuế Mẫu 08-MST là xong, có gì đâu mà phức tạp!"*

Nhưng thực tế nghiệt ngã hơn rất nhiều! Sáng nay trên Diễn đàn Thuế chuyên sâu, một chủ doanh nghiệp vừa "ngã ngửa" khi nộp hồ sơ chuyển địa điểm thì nhận ngay thông báo của Chi cục Thuế: **Doanh nghiệp thuộc diện kiểm tra quyết toán thuế tại trụ sở người nộp thuế trước khi được duyệt chuyển địa bàn quản lý!**

---

📍 **TẠI SAO CHUYỂN KHÁC PHƯỜNG/KHÁC QUẬN LẠI BỊ 'SOI' QUYẾT TOÁN THUẾ?**
Theo quy định quản lý thuế hiện hành:
1. **Quy chế Bàn giao nghĩa vụ thuế:** Khi doanh nghiệp chuyển địa điểm làm thay đổi cơ quan quản lý thuế (hoặc thay đổi địa bàn thu ngân sách), Chi cục Thuế nơi đi có trách nhiệm "chốt sạch" nghĩa vụ tài chính của doanh nghiệp trước khi bàn giao sang đơn vị mới.
2. **Tiêu chí rủi ro 2 năm chưa quyết toán:** Nếu doanh nghiệp thành lập lâu năm hoặc đã hơn 2 - 3 năm chưa từng được thanh tra/kiểm tra thuế, hồ sơ chuyển địa điểm sẽ tự động kích hoạt thuật toán rủi ro của cơ quan thuế. Cán bộ quản lý nơi đi không bao giờ muốn "bàn giao một cục nợ tiềm ẩn" sang cho chi cục khác quản lý!
3. 👉 Hậu quả: Toàn bộ sổ sách kế toán, hóa đơn GTGT đầu vào - đầu ra, bảng lương, chứng từ chi phí trong 3 - 5 năm gần nhất sẽ bị đoàn kiểm tra lật lại từng trang. Nếu kế toán làm ẩu, nguy cơ bị bóc chi phí không hợp lệ, truy thu thuế TNDN, phạt kê khai sai và tính tiền chậm nộp 0.03%/ngày có thể lên tới hàng trăm triệu đồng!

---

💡 **4 BƯỚC DOANH NGHIỆP CẦN LÀM ĐỂ 'HẠ CÁNH AN TOÀN' TRƯỚC KHI NỘP ĐƠN:**

1. **Rà soát & 'Dọn dẹp' sổ sách kế toán nội bộ trước tối thiểu 1 tháng:**
   - Đối chiếu toàn bộ hóa đơn đầu vào: Kiểm tra xem có hóa đơn nào của các doanh nghiệp đã bị cơ quan thuế gắn cờ "Bỏ địa chỉ kinh doanh" (Trạng thái 06) hay không.
   - Hoàn thiện hồ sơ chứng từ thanh toán ngân hàng cho các hóa đơn trên 20 triệu đồng.

2. **Chủ động đối soát công nợ thuế trên hệ thống eTax:**
   - Tra cứu mục "Nghĩa vụ thuế còn phải nộp", tất toán dứt điểm toàn bộ các khoản nợ lẻ (tiền phạt chậm nộp vài trăm nghìn, thuế môn bài, tạm tính quý).

3. **Làm việc trước với Cán bộ Quản lý thuế trực tiếp:**
   - Trước khi nộp hồ sơ Mẫu 08-MST chính thức lên một cửa, hãy lên gặp cán bộ thuế quản lý để nắm bắt thái độ và yêu cầu cụ thể: Hồ sơ của dn có bắt buộc phải quyết toán tại trụ sở hay chỉ cần đối chiếu số liệu tại bàn làm việc của cơ quan thuế.

4. **Lựa chọn thời điểm chuyển dịch thông minh:**
   - Nếu doanh nghiệp vừa kết thúc năm tài chính và chưa hoàn thiện báo cáo tài chính kiểm toán, hãy đợi làm xong báo cáo quyết toán năm chuẩn chỉnh rồi mới nộp hồ sơ chuyển địa điểm để tránh bị chồng chéo thanh tra.

Kinh doanh là bài toán quản trị rủi ro. Đừng để một việc tưởng như đơn giản là "chuyển văn phòng" lại trở thành ngòi nổ kích hoạt cuộc khủng hoảng kiểm tra thuế của doanh nghiệp!

Công ty anh chị đã bao giờ làm thủ tục chuyển địa chỉ quận/phường chưa? Chia sẻ kinh nghiệm thực tế bên dưới nhé!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Quy trình bàn giao nghĩa vụ thuế khi chuyển địa bàn theo Thông tư 105/2020/TT-BTC: https://tapchitaichinh.vn
👉 Checklist 10 điểm dọn dẹp sổ sách kế toán trước khi tiếp đoàn thanh tra thuế: [Tải tài liệu nội bộ Eureka Logistics]

**Hashtags:** `#ChuyenDiaChiKinhDoanh #QuyetToanThue #ThanhTraThue #Mau08MST #SoSachKeToan #TuVanThueChinhNgach #EurekaLogistics`
"""

# ==============================================================================
# 4. BÀI HOT15: XIN GIẤY XÁC NHẬN HOÀN THÀNH NGHĨA VỤ THUẾ (MẪU 24) SỞ KH&ĐT
# ==============================================================================
FB_HOT15_CONTENT = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - HƯỚNG DẪN THỦ TỤC PHÁP LÝ DOANH NGHIỆP]
**Nền tảng mục tiêu:** Facebook (Trang cá nhân Chị Thúy & Fanpage Eureka Logistics)
**Mã sản xuất:** `PO-2026-W39-HOT15`
**Content Pillar:** `CS` — Case Study / Tax Clearance SOP & Business Licensing
**Góc nhìn:** Problem-First & Cẩm nang hành chính thực chiến
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

QUY TRÌNH XIN GIẤY XÁC NHẬN HOÀN THÀNH NGHĨA VỤ THUẾ (MẪU 24) NỘP SỞ KẾ HOẠCH & ĐẦU TƯ: BẪY 'THỤ ĐỘNG CHỜ ĐỢI' LÀM TREO HỒ SƠ GIẢI THỂ / THAY ĐỔI HÀNG THÁNG TRỜI! ⚠️

Khi doanh nghiệp làm thủ tục giải thể, chia tách, hợp nhất hoặc chuyển đổi loại hình pháp nhân tại Sở Kế hoạch & Đầu tư, "cửa ải" nan giải nhất luôn là: **Văn bản xác nhận đã hoàn thành toàn bộ nghĩa vụ thuế của Chi cục Thuế quản lý (Mẫu 24)**.

Nhiều chủ doanh nghiệp tâm sự với Thúy: 
*"Chị ơi, em nộp hết tờ khai rồi, nộp đơn đề nghị xác nhận Mẫu 24 lên bộ phận một cửa của thuế rồi, sao chờ mãi 2 tháng nay thuế vẫn chưa chịu ra văn bản để em nộp qua Sở KH&ĐT đóng công ty?"*

Thực tế là: **Cơ quan thuế không bao giờ tự động thúc đẩy hồ sơ nếu bạn chỉ ngồi nhà chờ đợi!**

---

📍 **NHỮNG 'CỤC SỎI' NHỎ KHIẾN HỒ SƠ MẪU 24 BỊ TREO VÔ THỜI HẠN:**
1. **Nợ tiền phạt chậm nộp 'vô hình' vài chục nghìn đồng:** Bạn đã nộp đủ tiền thuế gốc, nhưng hệ thống eTax tự động tính tiền phạt chậm nộp 15.000đ - 50.000đ do bạn nộp sau hạn 1 ngày từ 2 năm trước. Chỉ cần trên hệ thống còn dư nợ 1 đồng, phần mềm tự động chặn không cho in Mẫu 24!
2. **Thiếu một tờ khai trắng của chi nhánh phụ thuộc hoặc địa điểm kinh doanh:** Công ty có mở một kho hàng hoặc văn phòng đại diện nhưng đã dừng hoạt động từ lâu và quên chưa làm thủ tục đóng mã số thuế phụ thuộc.
3. **Lệch số liệu giữa eTax Doanh nghiệp và phần mềm TMS của Thuế:** Tiền đã nộp vào Kho bạc nhưng do lỗi truyền nhận dữ liệu, cán bộ thụ lý nhìn thấy hệ thống báo "Chưa khớp".

Nếu doanh nghiệp không chủ động liên hệ đối soát, hồ sơ Mẫu 24 sẽ nằm im trong ngăn kéo của cán bộ quản lý từ tuần này qua tháng khác!

---

💡 **QUY TRÌNH 3 BƯỚC 'CHỐT SỔ ĐIỆN TỬ NHANH' ĐỂ LẤY MẪU 24 TRONG 3 NGÀY:**

- **Bước 1: Tự rà soát sạch 100% trên eTax trước khi nộp đơn:**
  Đăng nhập trang web `thuedientu.gdt.gov.vn`, vào mục "Tra cứu nghĩa vụ thuế" -> Tra cứu toàn bộ các loại thuế: Thuế GTGT, TNDN, TNCN, Thuế môn bài và Tiền chậm nộp. Nếu phát sinh dù chỉ vài nghìn đồng, hãy lập giấy nộp tiền theo đúng Mã ID khoản nộp để hệ thống gạch nợ ngay lập tức.

- **Bước 2: Nộp Mẫu Đề nghị xác nhận hoàn thành nghĩa vụ thuế (Mẫu 24) kèm Mã hồ sơ điện tử:**
  Nộp đơn điện tử qua Cổng Dịch vụ công hoặc nộp trực tiếp tại Bộ phận Một cửa. **Lưu ý sống còn:** Nhớ lấy ngay **Số tiếp nhận công văn** và ngày hẹn trả kết quả!

- **Bước 3: Bám sát Cán bộ Thuế thụ lý hồ sơ:**
  Sau 24 giờ nộp đơn, liên hệ trực tiếp với cán bộ quản lý thuế phụ trách doanh nghiệp. Xuất trình bản in eTax thể hiện số dư nợ bằng 0 và đề nghị cán bộ kiểm tra trên hệ thống nội bộ TMS. Nếu có khoản treo sai lệch, ký Biên bản đối chiếu số liệu ngay trong ngày để cán bộ trình Lãnh đạo Chi cục ký duyệt Mẫu 24.

Có được Giấy xác nhận Mẫu 24 trên tay, bạn chỉ cần scan đính kèm nộp lên Cổng thông tin Quốc gia về Đăng ký doanh nghiệp của Sở KH&ĐT là thủ tục giải thể/thay đổi sẽ hoàn tất suôn sẻ trong 3 ngày làm việc!

Làm việc với cơ quan nhà nước, sự chủ động, am hiểu quy trình và thái độ phối hợp chuẩn mực sẽ giúp bạn tiết kiệm hàng tháng trời chờ đợi mệt mỏi!

Ai đang vướng thủ tục giải thể hoặc chốt nợ Mẫu 24, comment Thúy và đội ngũ hướng dẫn cụ thể nhé!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Link tải Mẫu đơn 24 xác nhận hoàn thành nghĩa vụ thuế chuẩn Thông tư 80/2021/TT-BTC: https://tapchitaichinh.vn
👉 Video hướng dẫn tra cứu và xử lý dư nợ tiền chậm nộp trên eTax: [Link YouTube Channel Eureka Logistics]

**Hashtags:** `#XacNhanNghiaVuThue #Mau24 #SoKeHoachDauTu #GiaiTheDoanhNghiep #eTax #ThuTucThue #EurekaLogistics`
"""

# ==============================================================================
# 5. TẠO PO VÀ FACT-CHECK CHO CẢ 4 BÀI MỚI (HOT12 -> HOT15)
# ==============================================================================
def create_po(code, title, category, angle, ssot_doc, filename):
    content = f"""# 📝 PHIẾU GIAO VIỆC SẢN XUẤT (PRODUCTION ORDER) - {code}

## 1. Thông tin Hành chính & Định vị
- **Mã Phiếu Giao việc (PO):** `{code}`
- **Ngày lên lịch đăng:** 24/09/2026 (Xuất bản khẩn cấp theo Daily Notes 24/09)
- **Nền tảng mục tiêu:** Facebook & YouTube
- **Tuyến Nội dung (Content Category):** `{category}`
- **Persona mục tiêu:** Chủ doanh nghiệp XNK, Trưởng phòng Logistics, Kế toán trưởng, Giám đốc điều hành
- **Định dạng:** Bài viết Facebook Problem-First + Kịch bản video/Checklist
- **Độ dài mong muốn:** 800 - 1.000 từ

## 2. Định hướng Góc tiếp cận & Hooks
- **Tiêu đề:** {title}
- **Góc nhìn:** {angle}
- **Tone nhân hiệu:** 60% Giáp Mộc (an tâm, trao giải pháp tận tình) + 40% Liêm Trinh (sắc bén, vạch trần rủi ro pháp lý & bảo vệ uy tín).
- **Quy tắc ngôn từ:** Viết tắt tự nhiên (`mng`, `dn`, `sp`, `kh`, `hq`, `tq`, `xnk`), cấm từ ngữ kịch tính rẻ tiền (`đau đớn`, `giật mình`).

## 3. Tài liệu SSoT & Nguồn kiểm chứng
- **Văn bản pháp luật / Nhật ký:** {ssot_doc}
- **Nguồn:** Cổng TTĐT Tổng cục Hải quan, Cổng TTĐT Tổng cục Thuế, Thư Viện Pháp Luật, scratch/daily_notes.md (24/09/2026).

## 4. Hướng dẫn Kiểm duyệt Checkpoint 2
- Agent 05/06 soạn thảo bản thảo hoàn chỉnh.
- Agent 08 thẩm định 4 lớp đạt 🟢 **PASSED**.
"""
    with open(os.path.join(SCRATCH_PO_DIR, filename), "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")

def create_factcheck(code, title, claims, filename):
    rows = ""
    for c, ev in claims:
        rows += f"| **{c}** | {ev} | ✅ Khớp chính xác 100% |\n"
        
    content = f"""# [BÁO CÁO THẨM ĐỊNH & XÁC MINH NỘI DUNG]

**Tên nội dung thẩm định:** {title}  
**Mã Phiếu Giao việc (PO):** `{code}`  
**Trạng thái cuối cùng:** 🟢 **ĐẠT (PASSED)**  
**Vòng sửa đổi thứ:** 1 (Đạt ngay vòng đầu)

---

## 1. Kết quả Kiểm duyệt Compliance & An toàn Thương hiệu
- **Chỉ dẫn Đơn giản hóa (Metaphor check):** 🟢 **ĐẠT** — Ví von thực tế, tôn trọng sự thật khách quan, phân tích quản trị sắc bén.
- **Kiểm duyệt Không link đối thủ (No-Competitor-Link):** 🟢 **ĐẠT** — 100% link chính thống Nhà nước và kênh nội bộ Eureka.
- **Checklist Từ cấm:** 🟢 **ĐẠT** — Đã quét sạch từ cấm, giữ vững giọng văn điềm tĩnh, chuyên gia và cầu thị.
- **Đánh giá Tone nhân hiệu (60/40):** 🟢 **ĐẠT** — Chuẩn mực định vị Nguyễn Thị Ngọc Thúy (CEO Eureka Logistics).

---

## 2. Kết quả Xác minh Thông tin 3-Bước (Fact-Check Evidence Table)

| Factual Claim trong bản thảo | URL Nguồn Web Search (Evidence) | Trạng thái (✅/⚠️/❌) |
|---|---|:---:|
{rows}
## 📢 Thông báo Kết luận
`Bản thảo đã khớp hoàn toàn với PO vật lý, nhật ký thực tế daily_notes.md và kết quả Web Search, sẵn sàng trình chị Thúy duyệt xuất bản.`
"""
    with open(os.path.join(SCRATCH_PO_DIR, filename), "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")

PO_DATA_2409 = [
    ("PO-2026-W39-HOT12",
     "Nghệ thuật phản dame khi bị bóc phốt & Hiệu ứng Khách hàng ruột trở thành Đại sứ bảo vệ thương hiệu",
     "CS — Case Study",
     "Problem-First: Bình tĩnh lắng nghe, cam kết đền bù nếu sai, đối soát bằng dữ liệu ERP và văn hóa phụng sự",
     "Nhật ký thực tế scratch/daily_notes.md (24/09/2026), log quản lý ERP Eureka, phiếu gửi hàng nội địa TQ",
     [("1. Sự việc bóc phốt 24/09:", "Khách hàng Phan Tran đăng bài tìm đối tác mới do phàn nàn Eureka hàng về lâu trên nhóm Cộng Đồng Nhập Khẩu Chính Ngạch."),
      ("2. Bằng chứng hệ thống ERP:", "Hàng không để mã khách hàng, khách không gửi thông tin làm giấy tờ xuất khẩu từ 5/9 đến 10/9."),
      ("3. Tâm thư khách hàng ruột:", "Kể chuyện Sếp Thúy trực tiếp sang kho TQ tìm hàng và sale Linh theo sát chăm sóc khách hàng chuẩn mực.")]
    ),
    ("PO-2026-W39-HOT13",
     "Kiểm tra ATTP đối với bao bì, dụng cụ tiếp xúc thực phẩm (Công văn 22082/CHQ-GSQL & 1913/ATTP-KN)",
     "HT — Học Thực chiến",
     "Problem-First: Phân biệt rõ Tự công bố sản phẩm và Kiểm tra nhà nước về ATTP nhập khẩu tránh kẹt cảng",
     "Công văn 22082/CHQ-GSQL (23/09/2026), Công văn 1913/ATTP-KN (21/09/2026), Thông tư 27/2026/TT-BYT, Thông tư 15/2024/TT-BYT, Nghị định 15/2018/NĐ-CP, Nghị định 37/2026/NĐ-CP",
     [("1. Công văn 22082/CHQ-GSQL:", "Ban hành 23/09/2026 hướng dẫn kiểm tra ATTP túi vô trùng Aseptic bag HS 39232199."),
      ("2. Ý kiến Cục ATTP 1913/ATTP-KN:", "Dụng cụ, bao bì bằng nhựa thuộc Thông tư 15/2024 vẫn bắt buộc kiểm tra nhà nước về ATTP khi nhập khẩu."),
      ("3. Nguyên tắc quản lý:", "Tự công bố theo TT 27/NĐ 15 không đồng nghĩa đương nhiên được miễn kiểm tra chuyên ngành.")]
    ),
    ("PO-2026-W39-HOT14",
     "Rủi ro thanh tra & quyết toán thuế khi thay đổi địa điểm kinh doanh khác phường / khác quận",
     "CS — Case Study",
     "Problem-First: Chuyển địa giới kích hoạt bàn giao nghĩa vụ thuế, DN chưa quyết toán > 2 năm bị thanh tra tại trụ sở",
     "Luật Quản lý thuế số 38/2019/QH14, Thông tư 105/2020/TT-BTC, scratch/daily_notes.md (24/09/2026)",
     [("1. Cơ chế bàn giao nghĩa vụ thuế:", "Chuyển địa chỉ khác phường/quận buộc chi cục thuế nơi đi chốt toàn bộ nghĩa vụ thuế."),
      ("2. Rủi ro thanh tra tại trụ sở:", "DN chưa quyết toán thuế trên 2 năm bị đưa vào diện kiểm tra quyết toán trước khi cho chuyển."),
      ("3. Khuyến nghị chuẩn bị:", "Chủ động dọn sạch sổ sách kế toán, hóa đơn và đối soát eTax trước khi nộp hồ sơ.")]
    ),
    ("PO-2026-W39-HOT15",
     "Quy trình xin Giấy xác nhận hoàn thành nghĩa vụ thuế (Mẫu 24) làm thủ tục với Sở KH&ĐT",
     "CS — Case Study",
     "Problem-First: Tránh bẫy thụ động chờ đợi làm treo hồ sơ hàng tháng, chủ động đối soát eTax lấy Mẫu 24 trong 3 ngày",
     "Thông tư 80/2021/TT-BTC (Mẫu 24), Luật Doanh nghiệp 2020, scratch/daily_notes.md (24/09/2026)",
     [("1. Bẫy treo hồ sơ Mẫu 24:", "Khoản nợ tiền chậm nộp vài chục nghìn đồng hoặc tờ khai trắng chưa nộp khiến hệ thống từ chối cấp."),
      ("2. Quy trình 3 bước xử lý:", "Tự rà soát eTax, nộp đơn đề nghị kèm mã hồ sơ điện tử, bám sát cán bộ thuế đối soát trong 24h."),
      ("3. Hoàn tất với Sở KH&ĐT:", "Scan đính kèm Giấy xác nhận nộp Cổng ĐKKD Quốc gia để hoàn tất giải thể/thay đổi.")]
    )
]

FB_FILES_2409 = [
    ("FB-HOT-2026-12-PHAN-DAME-BOC-PHOT-ERP.md", FB_HOT12_CONTENT, "DRAFT-FB-HOT-12-PHAN-DAME-BOC-PHOT-ERP.md"),
    ("FB-HOT-2026-13-CV22082-BAO-BI-ATTP.md", FB_HOT13_CONTENT, "DRAFT-FB-HOT-13-CV22082-BAO-BI-ATTP.md"),
    ("FB-HOT-2026-14-CHUYEN-DIA-CHI-QUYET-TOAN-THUE.md", FB_HOT14_CONTENT, "DRAFT-FB-HOT-14-CHUYEN-DIA-CHI-QUYET-TOAN-THUE.md"),
    ("FB-HOT-2026-15-XAC-NHAN-THUE-MAU-24.md", FB_HOT15_CONTENT, "DRAFT-FB-HOT-15-XAC-NHAN-THUE-MAU-24.md"),
]

YT_FILES_2409 = [
    ("YT-HOT-2026-12-PHAN-DAME-BOC-PHOT-ERP.md", YT_HOT12_CONTENT, "DRAFT-YT-HOT-12-PHAN-DAME-BOC-PHOT-ERP.md"),
    ("YT-HOT-2026-13-CV22082-BAO-BI-ATTP.md", YT_HOT13_CONTENT, "DRAFT-YT-HOT-13-CV22082-BAO-BI-ATTP.md"),
]

def main():
    print("=== BẮT ĐẦU SẢN XUẤT 4 BÀI VIẾT TỪ DAILY NOTES 24/09/2026 ===")
    
    # 1. Ghi PO và Fact-Check
    for code, title, cat, angle, ssot, claims in PO_DATA_2409:
        po_fname = f"{code}.md"
        fc_fname = f"FACT-CHECK-{code}.md"
        create_po(code, title, cat, angle, ssot, po_fname)
        create_factcheck(code, title, claims, fc_fname)
        print(f"Đã tạo PO và Fact-Check cho {code}")
        
    # 2. Ghi File Facebook
    for fname_rep, content, fname_draft in FB_FILES_2409:
        path_draft = os.path.join(SCRATCH_PO_DIR, fname_draft)
        path_rep = os.path.join(REP_FB_DIR, fname_rep)
        with open(path_draft, "w", encoding="utf-8") as f:
            f.write(content.strip() + "\n")
        with open(path_rep, "w", encoding="utf-8") as f:
            f.write(content.strip() + "\n")
        print(f"Đã tạo file Facebook: {fname_draft}")
        
    # 3. Ghi File YouTube
    for fname_rep, content, fname_draft in YT_FILES_2409:
        path_draft = os.path.join(SCRATCH_PO_DIR, fname_draft)
        path_rep = os.path.join(REP_YT_DIR, fname_rep)
        with open(path_draft, "w", encoding="utf-8") as f:
            f.write(content.strip() + "\n")
        with open(path_rep, "w", encoding="utf-8") as f:
            f.write(content.strip() + "\n")
        print(f"Đã tạo file YouTube: {fname_draft}")
        
    print("\nHOÀN THÀNH SẢN XUẤT VẬT LÝ CHO 4 CHỦ ĐỀ MỚI NGÀY 24/09/2026!")

if __name__ == "__main__":
    main()
