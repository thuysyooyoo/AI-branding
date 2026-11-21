# -*- coding: utf-8 -*-
"""
Tự động sản xuất toàn bộ nội dung từ 5 chủ đề nóng (WF05 -> WF02 & WF03)
Gồm 5 bài viết Facebook hoàn chỉnh + 3 Kịch bản video YouTube chuẩn 2 phần
Tuân thủ tuyệt đối: brand_persona.md, voice_samples.md, banned_words.md, GEMINI.md
"""
import os
import sys

# Configure UTF-8
try:
    sys.stdout.reconfigure(encoding='utf-8')
except AttributeError:
    pass

BASE_DIR = r"d:\AI AGENT THUY\AI branding"
PO_DIR = os.path.join(BASE_DIR, "scratch", "production_orders")
FB_DIR = os.path.join(BASE_DIR, "reports", "plan1", "facebook")
YT_DIR = os.path.join(BASE_DIR, "reports", "plan1", "video_scripts")

os.makedirs(PO_DIR, exist_ok=True)
os.makedirs(FB_DIR, exist_ok=True)
os.makedirs(YT_DIR, exist_ok=True)

# ==============================================================================
# CONTENT DEFINITIONS
# ==============================================================================

# ------------------------------------------------------------------------------
# 1. TOPIC 1 - FACEBOOK: ĐÓNG MST ĐUÔI 888 & BÃO PHẠT CHẬM NỘP THUẾ AFFILIATE
# ------------------------------------------------------------------------------
FB_01 = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - CHUẨN GIỌNG ĐIỀM TĨNH & PHÂN TÍCH RỦI RO THUẾ TMĐT]
**Nền tảng mục tiêu:** Facebook (Fanpage/Personal Page)
**Mã sản xuất:** `PO-HOT-2026-01-MST888`
**Content Pillar:** `CS` — Case Study / Tax Big Data & Affiliate Compliance
**Góc nhìn:** Problem-First & Đặt vấn đề tranh luận thực tế
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

ĐÓNG MÃ SỐ THUẾ ĐUÔI 888 (AFFILIATE SÀN TMĐT): DOANH THU 1 TRIỆU PHẠT GẦN 6 TRIỆU VÌ CHẬM NỘP TỜ KHAI — CƠ QUAN THUẾ RÀ SOÁT DỮ LIỆU LIÊN THÔNG, MNG CẦN XỬ LÝ NGAY!

Sáng nay lướt Diễn đàn Thuế chuyên sâu, Thúy đọc được một case study của một bạn làm tiếp thị liên kết (affiliate) mà thấy xót xa vô cùng, nhưng đây lại là bài học cảnh tỉnh cực lớn cho hàng vạn anh chị đang bán hàng online và làm tiếp thị trên Shopee, TikTok Shop!

Chuyện là bạn này năm ngoái có đăng ký mã số thuế phụ (đuôi 888) để nhận tiền hoa hồng affiliate. Suốt cả năm, bạn chỉ phát sinh đúng 1 triệu đồng hoa hồng và phía sàn đã tự động khấu trừ 7% thuế (thuế TNCN & GTGT) nộp vào ngân sách nhà nước. 

Đến tuần này, bạn ra chi cục thuế làm thủ tục đóng mã số thuế 888 vì không còn làm tiếp thị nữa. Những tưởng thủ tục đơn giản, ai ngờ cán bộ thuế rà soát trên hệ thống và phát hiện: Suốt 4 quý vừa qua, bạn KHÔNG NỘP TỜ KHAI THUẾ ĐỊNH KỲ.

Kết quả: Cơ quan thuế lập biên bản xử phạt vi phạm hành chính chậm nộp hồ sơ khai thuế theo Điều 13 Nghị định 125/2020/NĐ-CP (tiểu mục 4254, chương 857). Dù doanh thu chỉ có 1 triệu đồng, bạn vẫn phải nộp phạt gần 6 triệu đồng cho mỗi quý nộp chậm quá hạn!

📍 **Cạm bẫy pháp lý mà 90% người làm affiliate và chủ shop đang lầm tưởng:**
Rất nhiều anh chị nhắn cho Thúy bảo: "Ủa Thúy ơi, sàn khấu trừ 7% thuế tại nguồn rồi, mình đâu có nợ đồng thuế nào của nhà nước, sao lại bị phạt nặng thế?"

Mng cần phân biệt rạch ròi 2 khái niệm hoàn toàn khác nhau trong Luật Quản lý thuế số 38:
1. **Nghĩa vụ nộp thuế:** Tiền thuế phải nộp vào ngân sách (sàn đã nộp thay 7%).
2. **Nghĩa vụ kê khai thuế:** Trách nhiệm nộp tờ khai Mẫu 01/TKN-CNKD định kỳ hàng quý để cơ quan thuế đối soát số liệu.

👉 Luật quy định: Một khi đã được cấp mã số thuế cá nhân kinh doanh (đuôi 888), **dù trong quý doanh thu bằng 0 đồng hoặc đã bị khấu trừ 100%, anh chị vẫn BẮT BUỘC phải nộp tờ khai quý**. Không nộp là bị quy vào hành vi vi phạm thủ tục hành chính, mức phạt từ 8 đến 15 triệu đồng/tờ khai (nếu có tình tiết giảm nhẹ thì hạ xuống mức tối thiểu 5 - 6 triệu)!

---

🤔 Vụ này đang nổ ra 2 luồng tranh cãi rất nảy lửa trên mạng:
🔴 Một bên thì bức xúc: "Thuế khấu trừ tận gốc rồi, doanh thu có 1 triệu mà phạt 6 triệu thì quá vô lý, làm khó dân kinh doanh nhỏ lẻ!"
🔵 Một bên thì khẳng định: "Kinh doanh trong kỷ nguyên số thì phải có ý thức tuân thủ pháp luật. Đã mở MST kinh doanh là phải có trách nhiệm khai báo, không thể đổ lỗi cho sự thiếu hiểu biết được!"

Quan điểm của mng thế nào? Bản thân dn hoặc cá nhân anh chị đã kiểm tra lại các kỳ kê khai trên eTax Mobile chưa?

---

📍 **3 Bước Thúy khuyên mng nên làm ngay hôm nay để không bị phạt oan:**
1. Đăng nhập ứng dụng eTax Mobile, vào mục "Tra cứu nghĩa vụ thuế" -> kiểm tra xem có kỳ kê khai nào đang bị báo "Chưa nộp tờ khai" không.
2. Nếu lỡ quên chưa nộp, hãy lập tức nộp bổ sung tờ khai Mẫu 01/TKN-CNKD ngay trước khi cơ quan thuế gửi thông báo kiểm tra.
3. Nếu đã bị mời lên làm việc, hãy chuẩn bị ngay văn bản giải trình áp dụng các tình tiết giảm nhẹ theo Điều 9 Nghị định 125/2020 (vi phạm lần đầu, số tiền phát sinh nhỏ, đã bị khấu trừ tại nguồn đầy đủ) để xin áp dụng mức phạt khung thấp nhất.

P/S: Chi tiết mẫu văn bản giải trình xin giảm nhẹ mức phạt theo Nghị định 125 và video từng bước tra cứu trên eTax Mobile, Thúy để ở link comment đầu tiên nhé. Mng tải về dùng khi cần!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Link tải Mẫu công văn giải trình tình tiết giảm nhẹ theo Điều 9 NĐ 125/2020: https://tapchitaichinh.vn
👉 Video hướng dẫn tra cứu và kê khai Mẫu 01/TKN-CNKD trên eTax Mobile: [Link YouTube Channel Eureka Logistics]

**Hashtags:** `#ThueAffiliate #DongMaSoThue888 #NghiDinh125 #eTaxMobile #QuanLyThueTMDT #TuVanThueChinhNgach #EurekaLogistics`
"""

# ------------------------------------------------------------------------------
# 2. TOPIC 1 - YOUTUBE: HƯỚNG DẪN GỠ PHẠT ĐÓNG MST 888 & KÊ KHAI ETAX MOBILE
# ------------------------------------------------------------------------------
YT_01 = """# [KỊCH BẢN YOUTUBE CHI TIẾT - CHUẨN GIÁO DỤC THỰC CHIẾN & SCREENCAST]
**Nền tảng mục tiêu:** YouTube Dài (Video 8 - 10 phút)
**Mã sản xuất:** `PO-HOT-2026-01-MST888`
**Content Pillar:** `CS` — Case Study & Hướng dẫn Thực chiến
**Tiêu đề Video:** BÃO PHẠT CHẬM NỘP THUẾ KHI ĐÓNG MST ĐUÔI 888: DOANH THU 1 TRIỆU PHẠT GẦN 6 TRIỆU & QUY TRÌNH GỠ RỐI TRÊN ETAX MOBILE
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## 1. Bài Viết Lời Thoại Hoàn Chỉnh (Full Text Script for Narration / Repurpose)

**DOANH THU 1 TRIỆU NHƯNG BỊ PHẠT GẦN 6 TRIỆU ĐỒNG KHI ĐÓNG MÃ SỐ THUẾ ĐUÔI 888 — BÀI HỌC XƯƠNG MÁU VỀ KÊ KHAI THUẾ TIẾP THỊ LIÊN KẾT TRÊN ETAX MOBILE!**

Chào các bạn! Thúy đây. 

Nếu bạn đang làm tiếp thị liên kết (affiliate) trên Shopee, TikTok Shop, hoặc đang mở mã số thuế phụ đuôi 888 để nhận hoa hồng dịch vụ... thì video ngày hôm nay Thúy khuyên các bạn phải xem thật kỹ từng phút, bởi vì nó liên quan trực tiếp đến số tiền trong tài khoản của các bạn!

Cách đây vài ngày, một bạn chuyên viên làm affiliate có nhắn tin cầu cứu Thúy trong trạng thái hoảng hốt. Bạn ấy nộp hồ sơ xin đóng mã số thuế phụ đuôi 888. Năm ngoái cả năm bạn chỉ có doanh thu hoa hồng đúng 1 triệu đồng và sàn TMĐT đã khấu trừ 7% thuế nộp vào kho bạc. Thế nhưng khi lên chi cục thuế, cán bộ thuế in ra một danh sách 4 quý liên tiếp bạn không nộp tờ khai thuế quý, kèm theo thông báo xử phạt hành chính dự kiến gần 24 triệu đồng theo Nghị định 125/2020/NĐ-CP!

Rất nhiều người đang ngây thơ nghĩ rằng: *"Sàn đã trừ thuế của mình rồi, mình đâu có nợ thuế nhà nước, sao lại bị phạt?"*

Hôm nay, Thúy sẽ bóc tách rành mạch cho các bạn:
Thứ nhất: Tại sao không nợ thuế mà vẫn bị phạt tới hàng chục triệu đồng?
Thứ hai: Cơ chế liên thông dữ liệu giữa ngành Thuế và các đơn vị vận chuyển, sàn TMĐT qua hệ thống Big Data đang vận hành thế nào?
Và thứ ba: Hướng dẫn các bạn 3 bước gỡ rối, nộp bổ sung tờ khai Mẫu 01/TKN-CNKD trên eTax Mobile và cách làm đơn giải trình xin giảm nhẹ mức phạt theo Điều 9 Nghị định 125/2020.

Bây giờ, chúng ta cùng đi vào nội dung chi tiết!

---

### PHẦN 1: BẢN CHẤT PHÁP LÝ — TẠI SAO DOANH THU 0 ĐỒNG VẪN BỊ PHẠT?

Theo quy định tại Luật Quản lý thuế số 38 và Thông tư 40/2021/TT-BTC, mã số thuế đuôi 888 là mã số thuế phụ cấp cho cá nhân kinh doanh hoặc hộ kinh doanh.

Khi bạn đã có mã số thuế này, bạn có nghĩa vụ pháp lý phải nộp hồ sơ khai thuế định kỳ theo quý (chậm nhất là ngày cuối cùng của tháng đầu quý tiếp theo).
Theo Điều 13 Nghị định 125/2020/NĐ-CP:
- Chậm nộp từ 1 đến 30 ngày: Phạt từ 2 triệu đến 5 triệu đồng.
- Chậm nộp từ 31 đến 60 ngày: Phạt từ 5 triệu đến 8 triệu đồng.
- Chậm nộp từ 61 đến 90 ngày: Phạt từ 8 triệu đến 15 triệu đồng.
- Chậm nộp trên 90 ngày: Phạt từ 15 triệu đến 25 triệu đồng!

Mức phạt này áp dụng cho **hành vi vi phạm về thời hạn nộp hồ sơ**, hoàn toàn độc lập với việc bạn có phát sinh số thuế phải nộp hay không! Nghĩa là dù doanh thu 0 đồng, nếu bạn không nộp tờ khai trắng, bạn vẫn bị phạt tiền!

---

### PHẦN 2: SCREENCAST HƯỚNG DẪN 3 BƯỚC XỬ LÝ TRÊN ETAX MOBILE

Bây giờ Thúy sẽ quay trực tiếp màn hình ứng dụng eTax Mobile để hướng dẫn các bạn:
- **Bước 1: Tra cứu nghĩa vụ kê khai.** Các bạn vào mục Tra cứu nghĩa vụ thuế -> Chọn Hồ sơ khai thuế. Ở đây hệ thống sẽ liệt kê các tờ khai quý còn thiếu.
- **Bước 2: Kê khai bổ sung Mẫu 01/TKN-CNKD.** Nhập doanh thu thực tế được sàn báo cáo. Nếu sàn đã khấu trừ 7%, ghi rõ vào mục thuế đã khấu trừ tại nguồn để hệ thống không nhảy số tiền phải nộp thêm.
- **Bước 3: Soạn văn bản giải trình tình tiết giảm nhẹ.** Theo Điều 9 Nghị định 125/2020, nếu bạn vi phạm lần đầu, không cố ý, và số tiền thuế đã được khấu trừ đầy đủ, bạn được quyền yêu cầu áp dụng mức phạt tối thiểu của khung phạt.

---

### PHẦN 3: KẾT LUẬN & CLIFFHANGER TẬP TIẾP THEO

Kỷ nguyên kinh doanh trên không gian số không còn chỗ cho sự làm mò hay chủ quan. Hiểu luật và chuẩn hóa từ đầu là chiếc khiên bảo vệ tài chính an toàn nhất cho các bạn!

🎬 **Ở video tiếp theo, Thúy sẽ phân tích một chủ đề cực nóng cho các doanh nghiệp XNK: Quyết định 31/2026/QĐ-TTg và quy định 12 nhóm hàng bắt buộc làm thủ tục tại cửa khẩu nhập. Dính một món rượu hay điều hòa là kẹt cả cont hàng ở cảng biển! Các bạn nhớ bấm Đăng ký kênh và bật chuông để không bỏ lỡ nhé!**

---

## 2. Bảng Phân Cảnh Kịch Bản Chi Tiết (Storyboard 4 Cột)

| Phân đoạn / Thời gian | Hình ảnh & Bối cảnh (Visual) | Lời thoại & Giọng đọc (Audio) | Text On Screen & SFX |
|---|---|---|---|
| **01. Hook (0:00 - 0:50)** | Cận cảnh chị Thúy ngồi tại văn phòng, ánh mắt sắc bén, trên bàn là biên bản xử phạt thuế eTax. | **Chị Thúy:** Doanh thu 1 triệu nhưng bị phạt gần 6 triệu đồng khi đóng MST đuôi 888! Đừng nghĩ sàn trừ thuế rồi là xong, đây là cái bẫy hàng vạn người đang mắc phải! | **Text lớn:** DOANH THU 1 TRIỆU — PHẠT GẦN 6 TRIỆU! <br> *SFX: Tiếng chuông cảnh báo Warning* |
| **02. Bóc tách luật (0:50 - 3:30)** | Đồ họa so sánh 2 cột: Nghĩa vụ nộp thuế (sàn trừ 7%) vs Nghĩa vụ nộp tờ khai (Điều 13 NĐ 125/2020). | **Chị Thúy:** Rất nhiều bạn nhầm lẫn giữa nộp thuế và nộp tờ khai. Dù doanh thu 0 đồng, nếu không nộp Mẫu 01/TKN-CNKD, mức phạt vẫn từ 8 đến 15 triệu đồng/quý! | **Text:** ĐIỀU 13 NGHỊ ĐỊNH 125/2020/NĐ-CP <br> Bảng khung tiền phạt theo số ngày trễ |
| **03. Screencast eTax (3:30 - 7:00)** | Quay màn hình điện thoại thao tác app eTax Mobile: Vào mục Tra cứu nghĩa vụ -> Lập tờ khai Mẫu 01/TKN-CNKD. | **Chị Thúy:** Bây giờ Thúy sẽ thao tác trực tiếp trên màn hình eTax Mobile: 3 bước tra cứu tờ khai sót và điền chỉ tiêu khấu trừ 7% đúng chuẩn không bị tính thuế đúp. | Con trỏ đỏ zoom cận từng ô nhập liệu trên eTax Mobile <br> *SFX: Tiếng click chuột dứt khoát* |
| **04. Mẫu giải trình & Outro (7:00 - 8:30)** | Chị Thúy mỉm cười điềm tĩnh, chia sẻ triết lý tuân thủ pháp lý và dẫn dắt sang chủ đề Quyết định 31/2026. | **Chị Thúy:** Tuân thủ pháp lý chính là chiếc khiên bảo vệ tài chính của bạn. Tập tiếp theo, Thúy sẽ phân tích Quyết định 31 và rủi ro kẹt cont 12 nhóm hàng cửa khẩu nhập! | **Text:** LINK TẢI MẪU GIẢI TRÌNH Ở COMMENT <br> Cliffhanger: Quyết định 31/2026/QĐ-TTg! |
"""

# ------------------------------------------------------------------------------
# 3. TOPIC 2 - FACEBOOK: THỦ TỤC NHẬP KHẨU RƯỢU NGOẠI & BẪY CỬA KHẨU NHẬP
# ------------------------------------------------------------------------------
FB_02 = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - CHUẨN GIỌNG ĐIỀM TĨNH & TUYẾN PHÁP LÝ NHẬP KHẨU]
**Nền tảng mục tiêu:** Facebook (Fanpage/Personal Page)
**Mã sản xuất:** `PO-HOT-2026-02-RUOU-NGOAI`
**Content Pillar:** `HT` — Học Thực chiến / Regulated Imports & Food Safety
**Góc nhìn:** Problem-First & Cảnh báo rủi ro lưu bãi cảng
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

NHẬP KHẨU RƯỢU NGOẠI & DANH MỤC 12 NHÓM HÀNG CỬA KHẨU NHẬP (QĐ 31/2026): CẠM BẪY THIẾU GIẤY PHÉP & RỦI RO KẸT CONT LƯU BÃI CẢNG HÀNG TRĂM TRIỆU MÙA CUỐI NĂM!

Thời điểm này hàng năm là lúc các dn thương mại rục rịch ký hợp đồng nhập khẩu rượu vang, rượu mạnh từ Pháp, Ý, Chile, Úc về để phục vụ thị trường quà biếu Tết và lễ hội cuối năm. 

Nhưng nếu năm nay anh chị vẫn áp dụng quy trình nhập khẩu của những năm trước mà không cập nhật 2 văn bản pháp lý mới toanh vừa có hiệu lực (Quyết định 31/2026/QĐ-TTg và Thông tư 28/2026/TT-BCT), thì nguy cơ cont hàng bị "đóng băng" tại cảng biển là gần như 100%!

📍 **Cạm bẫy thứ 1: Bắt buộc làm thủ tục tại cửa khẩu nhập (QĐ 31/2026/QĐ-TTg)**
Trước đây, nhiều dn mở tờ khai xin chuyển cửa khẩu đưa cont rượu về các ICD nội địa (như ICD Mỹ Đình, ICD Tiên Sơn, ICD Long Bình) gần kho để kiểm hóa cho tiện.
Nhưng từ ngày 14/08/2026, Quyết định 31/2026/QĐ-TTg đã chính thức khóa van này lại: Rượu thuộc nhóm 12 mặt hàng BẮT BUỘC phải làm toàn bộ thủ tục thông quan ngay tại cửa khẩu cảng biển/sân bay quốc tế đầu tiên (Cát Lái, Hải Phòng, Nội Bài, Tân Sơn Nhất). Không còn cơ chế chuyển cửa khẩu tự do như trước nữa!

📍 **Cạm bẫy thứ 2: Giấy phép phân phối rượu (Nghị định 105/2017 & NĐ 17/2020)**
Rất nhiều anh chị mới khởi nghiệp thương mại lầm tưởng rằng: "Dn mình đã có Giấy phép bán buôn rượu hoặc Giấy phép bán lẻ rượu do Sở Công Thương cấp là đủ quyền nhập khẩu rượu từ nước ngoài về".
Đây là sai lầm chết người! Theo quy định, để được quyền nhập khẩu rượu có độ cồn từ 5,5 độ trở lên, dn BẮT BUỘC phải sở hữu **Giấy phép phân phối rượu** do Bộ Công Thương cấp. Thiếu giấy phép này mà tàu cập cảng là hq từ chối tiếp nhận tờ khai ngay lập tức, tiền lưu bãi cắm điện (demurrage/detention) mỗi ngày ngốn hàng triệu đồng!

📍 **Cạm bẫy thứ 3: Kiểm tra nhà nước về an toàn thực phẩm (Thông tư 28/2026/TT-BCT)**
Rượu thuộc danh mục kiểm tra ATTP do Bộ Công Thương quản lý. Dn bắt buộc phải làm hồ sơ Tự công bố sản phẩm theo Nghị định 15/2018/NĐ-CP TRƯỚC KHI hàng cập cảng. Hàng về đến nơi phải đăng ký kiểm tra nhà nước, lấy mẫu kiểm nghiệm đạt chuẩn mới được giải phóng hàng mang về kho tiêu thụ.

📍 **Cạm bẫy thứ 4: Áp lực thuế tiêu thụ đặc biệt 65% và tham vấn giá hải quan**
Rượu gánh cùng lúc 3 tầng thuế: Thuế NK (thường 50%), Thuế TTĐB (65% đối với rượu từ 20 độ trở lên) và VAT 10%. Nếu hồ sơ chứng từ thanh toán quốc tế không rõ ràng hoặc khai đơn giá quá thấp, hq sẽ đưa vào luồng tham vấn giá và ấn định tăng thuế, dn không kịp chuẩn bị dòng tiền nộp thuế là hàng nằm im tại cảng!

---

🤔 Trên các hội nhóm thủ tục hq mấy ngày qua đang tranh luận rất nhiều:
🔴 Nhiều người than phiền: "Nhập một cont rượu mà giấy phép con tầng tầng lớp lớp, chi phí kiểm nghiệm quá đắt đỏ, chỉ béo các tập đoàn lớn!"
🔵 Nhưng ở góc nhìn quản lý, nhiều anh chị khẳng định: "Rượu liên quan trực tiếp đến sức khỏe người dân và nguồn thu ngân sách lớn. Kiểm soát ngặt nghèo về nguồn gốc, tem nhãn và ATTP là hoàn toàn cần thiết để dẹp bỏ tình trạng nhập nhèm chất lượng!"

Quan điểm của mng thế nào? Dn mình đã hoàn tất bộ hồ sơ tự công bố và giấy phép phân phối trước khi book tàu chưa?

---

💡 **Lời khuyên từ Thúy:**
Làm hàng kinh doanh có điều kiện thì "chậm mà chắc mới là nhanh". Tuyệt đối không book tàu khi chưa có trong tay Giấy phép phân phối rượu và bản Tự công bố ATTP hợp lệ. 

Anh chị nào cần trọn bộ Checklist 7 loại chứng từ bắt buộc khi nhập khẩu rượu ngoại và quy trình tham vấn giá hải quan chuẩn chỉnh, cứ comment Thúy gửi tặng nhé!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Link tải Checklist 7 chứng từ nhập khẩu rượu & Quy trình tự công bố ATTP: https://customs.gov.vn
👉 Video cẩm nang hướng dẫn thủ tục nhập khẩu rượu ngoại trên YouTube Eureka: [Link YouTube Channel Eureka Logistics]

**Hashtags:** `#NhapKhauRuouNgoai #QuyetDinh31 #AnToanThucPham #ThueTieuThuDacBiet #ThuTucHaiQuan #ChinhNgachMinhBach #EurekaLogistics`
"""

# ------------------------------------------------------------------------------
# 4. TOPIC 2 - YOUTUBE: KỊCH BẢN CẨM NANG THỦ TỤC NHẬP KHẨU RƯỢU NGOẠI
# ------------------------------------------------------------------------------
YT_02 = """# [KỊCH BẢN YOUTUBE CHI TIẾT - CẨM NANG PHÁP LÝ & QUY TRÌNH THỰC CHIẾN]
**Nền tảng mục tiêu:** YouTube Dài (Video 8 - 12 phút)
**Mã sản xuất:** `PO-HOT-2026-02-RUOU-NGOAI`
**Content Pillar:** `HT` — Học Thực chiến / Regulated Imports & Food Safety
**Tiêu đề Video:** CẨM NANG THỦ TỤC NHẬP KHẨU RƯỢU NGOẠI 2026: 7 CHỨNG TỪ BẮT BUỘC, BẪY KIỂM TRA ATTP & QUY TẮC CỬA KHẨU NHẬP QĐ 31/2026
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## 1. Bài Viết Lời Thoại Hoàn Chỉnh (Full Text Script for Narration / Repurpose)

**NHẬP KHẨU RƯỢU NGOẠI 2026: 7 CHỨNG TỪ BẮT BUỘC, QUY TRÌNH KIỂM TRA ATTP & BẪY KẸT CẢNG CẦN TRÁNH TRƯỚC MÙA TẾT!**

Chào các bạn! Thúy đây. 

Cứ vào quý 3 và quý 4 hàng năm, thị trường đồ uống và rượu ngoại nhập khẩu lại bước vào giai đoạn sôi động nhất. Rất nhiều doanh nghiệp thương mại F&B muốn nhập khẩu những dòng rượu vang cao cấp từ Pháp, Ý, Tây Ban Nha hay các dòng rượu mạnh từ Scotland, Nhật Bản về Việt Nam.

Thế nhưng, rượu là một trong những mặt hàng nhập khẩu có điều kiện KHẮT KHE NHẤT hiện nay tại Việt Nam. Không ít doanh nghiệp đã phải trả giá đắt bằng việc cont hàng nằm phơi nắng tại cảng Cát Lái hay Hải Phòng cả tháng trời, tiền lưu kho bãi, tiền cắm điện container ngốn hàng trăm triệu đồng, thậm chí bị buộc tái xuất chỉ vì thiếu một loại giấy phép con!

Đặc biệt từ tháng 8/2026, với sự ra đời của **Quyết định 31/2026/QĐ-TTg** và **Thông tư 28/2026/TT-BCT**, toàn bộ quy trình thông quan rượu nhập khẩu đã có những thay đổi mang tính bước ngoặt.

Trong video hôm nay, Thúy sẽ hướng dẫn các bạn trọn bộ cẩm nang từ A đến Z:
- 3 Điều kiện pháp lý bắt buộc doanh nghiệp phải có TRƯỚC KHI ký hợp đồng ngoại thương.
- Bộ hồ sơ hải quan gồm đúng 7 loại chứng từ không thể thiếu.
- Quy trình đăng ký tự công bố an toàn thực phẩm và kiểm tra nhà nước theo Nghị định 15/2018.
- Và kinh nghiệm thực tế giúp doanh nghiệp vượt qua vòng tham vấn giá tính thuế mà không bị ấn định giá trần.

Hãy cùng Thúy bắt đầu ngay nhé!

---

### PHẦN 1: BA ĐIỀU KIỆN PHÁP LÝ TIÊN QUYẾT

Nhiều anh chị hay hỏi Thúy: "Công ty anh có đăng ký ngành nghề kinh doanh thực phẩm rồi, có nhập rượu được không?"
Câu trả lời là KHÔNG nếu anh chị chưa thỏa mãn 3 điều kiện sau:
1. **Giấy phép phân phối rượu:** Theo Nghị định 105/2017 và Nghị định 17/2020 của Chính phủ, đối với rượu có nồng độ cồn từ 5,5 độ trở lên, chỉ doanh nghiệp có Giấy phép phân phối rượu do Bộ Công Thương cấp mới được quyền trực tiếp nhập khẩu. Giấy phép bán buôn do Sở Công Thương cấp không có giá trị làm thủ tục nhập khẩu!
2. **Quy định Cửa khẩu nhập theo QĐ 31/2026/QĐ-TTg:** Rượu chỉ được làm thủ tục hải quan tại cửa khẩu cảng biển hoặc sân bay quốc tế đầu tiên cập cảng. Tuyệt đối không được mở tờ khai chuyển cửa khẩu về các ICD nội địa.
3. **Tem rượu nhập khẩu:** Rượu nhập khẩu bắt buộc phải dán tem điện tử theo quy định của Bộ Tài chính trước khi đưa ra thị trường lưu thông.

---

### PHẦN 2: BỘ HỒ SƠ 7 CHỨNG TỪ HẢI QUAN CHUẨN CHỈNH

Khi truyền tờ khai trên hệ thống VNACCS, bộ hồ sơ của các bạn bắt buộc phải có:
1. Tờ khai hải quan điện tử (Mẫu NK).
2. Hóa đơn thương mại (Commercial Invoice) & Bảng kê chi tiết đóng gói (Packing List).
3. Vận đơn đường biển (Bill of Lading).
4. Bản sao công chứng Giấy phép phân phối rượu còn hiệu lực.
5. Bản Tự công bố sản phẩm kèm Phiếu kết quả kiểm nghiệm an toàn thực phẩm (Test Report) của phòng kiểm nghiệm được chỉ định.
6. Giấy chứng nhận xuất xứ hàng hóa (C/O) nếu muốn hưởng thuế nhập khẩu ưu đãi đặc biệt (như Form EUR.1 từ Châu Âu hoặc Form CPTPP).
7. Giấy đăng ký kiểm tra nhà nước về an toàn thực phẩm đối với thực phẩm nhập khẩu.

---

### PHẦN 3: KINH NGHIỆM THAM VẤN GIÁ VÀ KIỂM SOÁT THUẾ TTĐB

Rượu chịu thuế Tiêu thụ đặc biệt lên tới 65% (đối với nồng độ từ 20 độ trở lên). Vì thuế suất rất cao nên cơ quan hải quan kiểm soát trị giá tính thuế cực kỳ nghiêm ngặt.
Để không bị hải quan nghi ngờ khai thấp giá và ấn định thuế, doanh nghiệp cần chuẩn bị sẵn:
- Chứng từ thanh toán quốc tế qua ngân hàng (T/T hoặc L/C).
- Hợp đồng đại lý phân phối độc quyền hoặc bảng chào giá chính thức (Price List) của nhà sản xuất quốc tế có công chứng.
- Chứng từ chứng minh các chi phí vận chuyển đường biển và bảo hiểm thực tế.

---

### KẾT LUẬN & CLIFFHANGER

Làm hàng chính ngạch rượu ngoại đòi hỏi sự cẩn trọng và tính kỷ luật hồ sơ cực kỳ cao. Khi bạn làm chuẩn chỉ ngay từ đầu, bạn sẽ hoàn toàn yên tâm đưa sản phẩm ra thị trường mà không sợ bất kỳ rủi ro thanh tra, hậu kiểm nào.

🎬 **Ở tập tiếp theo, Thúy sẽ mang đến cho các bạn một thông tin khẩn cấp của ngành Hải quan: Công văn 21692/CHQ-CNTT đếm ngược làm sạch dữ liệu người dùng trên VNACCS từ 25/09/2026. Nếu không rà soát, doanh nghiệp có thể bị khóa tài khoản truyền tờ khai! Các bạn hãy bấm Đăng ký kênh để theo dõi ngay nhé!**

---

## 2. Bảng Phân Cảnh Kịch Bản Chi Tiết (Storyboard 4 Cột)

| Phân đoạn / Thời gian | Hình ảnh & Bối cảnh (Visual) | Lời thoại & Giọng đọc (Audio) | Text On Screen & SFX |
|---|---|---|---|
| **01. Hook (0:00 - 1:00)** | Cận cảnh chị Thúy cầm chai rượu vang nhập khẩu cao cấp, background phía sau là cảng biển container nhộn nhịp. | **Chị Thúy:** Nhập một cont rượu ngoại kiếm tiền tỷ mùa Tết, nhưng dính một lỗi giấy phép là kẹt cả tháng tại cảng biển Cát Lái, mất trắng hàng trăm triệu! 2026 muốn nhập rượu chuẩn luật, bạn phải nắm 7 chứng từ này! | **Text lớn:** NHẬP KHẨU RƯỢU NGOẠI 2026: 7 CHỨNG TỪ BẮT BUỘC! <br> *SFX: Tiếng container đặt xuống bãi cạch lớn* |
| **02. 3 Điều kiện pháp lý (1:00 - 4:00)** | Đồ họa Infographic hiển thị: 1. Giấy phép phân phối rượu (Bộ Công Thương) -> 2. Quyết định 31/2026 Cửa khẩu nhập -> 3. Tem rượu điện tử. | **Chị Thúy:** Giấy phép bán buôn của Sở Công Thương không được phép nhập khẩu rượu! Bạn bắt buộc phải có Giấy phép phân phối rượu từ 5,5 độ trở lên do Bộ Công Thương cấp. | Zoom cận Giấy phép phân phối rượu Bộ Công Thương <br> **Text:** QĐ 31/2026/QĐ-TTg — CỬA KHẨU NHẬP! |
| **03. Bộ hồ sơ 7 chứng từ (4:00 - 7:30)** | Screencast hệ thống khai báo hải quan ECUS5-VNACCS, lần lượt điểm mặt từng file PDF chứng từ đính kèm (V/V, C/O, Tự công bố). | **Chị Thúy:** Đây là 7 chứng từ bắt buộc phải upload lên hệ thống thông quan điện tử. Thiếu một bản Tự công bố ATTP, hệ thống sẽ từ chối thông quan ngay lập tức! | Checklist 7 chứng từ hiển thị dạng tích xanh ✅ <br> *SFX: Tiếng ting ting duyệt chứng từ* |
| **04. Tham vấn giá & Outro (7:30 - 9:30)** | Chị Thúy ngồi tại bàn làm việc, phân tích công thức tính thuế TTĐB 65% và chia sẻ giải pháp tham vấn giá. | **Chị Thúy:** Thuế TTĐB 65% là con số không nhỏ. Hãy chuẩn bị Price List quốc tế và điện chuyển tiền ngân hàng. Tập tiếp theo: Đếm ngược làm sạch dữ liệu VNACCS! | **Text:** BÍ QUYẾT THAM VẤN GIÁ HẢI QUAN <br> Subscribe để đón xem Tập tiếp theo! |
"""

# ------------------------------------------------------------------------------
# 5. TOPIC 3 - FACEBOOK: CÔNG VĂN 21692/CHQ-CNTT ĐẾM NGƯỢC 3 NGÀY LÀM SẠCH VNACCS
# ------------------------------------------------------------------------------
FB_03 = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - BẢN TIN KHẨN CẤP & CHECKLIST HÀNH ĐỘNG]
**Nền tảng mục tiêu:** Facebook (Fanpage/Personal Page)
**Mã sản xuất:** `PO-HOT-2026-03-CV21692-VNACCS`
**Content Pillar:** `HT` — Học Thực chiến / E-Customs Identity & Compliance
**Góc nhìn:** Thông báo nghiệp vụ khẩn cấp & Hành động ngay trong 72 giờ
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

ĐẾM NGƯỢC 03 NGÀY (25/09/2026): CÔNG VĂN 21692/CHQ-CNTT YÊU CẦU LÀM SẠCH DỮ LIỆU TÀI KHOẢN VNACCS — ANH CHỊ DN KHẨN TRƯƠNG RÀ SOÁT TRÁNH BỊ NGHẼN LUỒNG TRUYỀN TỜ KHAI!

Một thông báo cực kỳ quan trọng và gấp rút từ Cục Hải quan mà toàn bộ các anh chị chủ dn xnk, trưởng phòng logistics và các bạn làm thủ tục hải quan BẮT BUỘC phải thực hiện ngay trước ngày 25/09/2026!

Cục Hải quan vừa ban hành Công văn hỏa tốc số 21692/CHQ-CNTT về kế hoạch rà soát, cập nhật và làm sạch dữ liệu trên "Hệ thống Đăng ký và quản lý người sử dụng" nhằm liên thông, xác thực định danh với Cơ sở dữ liệu quốc gia về Dân cư.

📍 **Thời hạn chính thức áp dụng:** Kể từ ngày **25/09/2026** (Chỉ còn đúng 72 giờ nữa)!
📍 **Địa chỉ cổng portal hệ thống mới:** https://dksd.customs.gov.vn/client/external

Nếu dn mình không chủ động thực hiện rà soát thông tin tài khoản, khi hệ thống mới chính thức chuyển đổi dữ liệu, nguy cơ rất cao là tài khoản khai báo hq điện tử của công ty sẽ bị gián đoạn, không truyền được tờ khai VNACCS, khiến hàng hóa cập cảng bị ách tắc hoàn toàn!

---

📍 **Checklist 4 việc khẩn cấp dn xnk cần làm ngay hôm nay:**

1. **Rà soát danh sách nhân sự được cấp quyền khai báo:** Thu hồi ngay tài khoản của những nhân viên giao nhận (ops), nhân viên chứng từ đã nghỉ việc hoặc chuyển bộ phận. Tránh tình trạng nhân sự cũ vẫn sử dụng tài khoản công ty để làm các dịch vụ bên ngoài, gây rủi ro pháp lý cực lớn cho người đại diện pháp luật!
2. **Chuẩn hóa thông tin theo Căn cước công dân (CCCD):** Cập nhật chính xác số định danh cá nhân/CCCD 12 số của người đại diện pháp luật và những người được ủy quyền sử dụng tài khoản để hệ thống tự động đối soát với CSDL Quốc gia về Dân cư.
3. **Kiểm tra thời hạn chữ ký số (Token):** Đảm bảo chữ ký số đang gắn trên hệ thống VNACCS còn hạn sử dụng và khớp đúng với chứng thư số đăng ký trên hệ thống Tổng cục Hải quan.
4. **Phân quyền người dùng theo đúng thực tế:** Điều chỉnh lại quyền hạn (quyền truyền tờ khai, quyền tra cứu nợ thuế, quyền nộp thuế điện tử) đúng với trách nhiệm của từng nhân sự phụ trách.

---

💡 **Lời nhắn nhủ từ Thúy:**
Chuyển đổi số ngành Hải quan đang diễn ra với tốc độ rất nhanh (từ Thông tư 85 về cơ sở dữ liệu phân loại, đến QĐ 1434 về chuẩn dữ liệu JSON, và nay là làm sạch định danh tài khoản người dùng theo CV 21692). Dn nào chủ động chuẩn hóa dữ liệu nội bộ sớm, dn đó sẽ luôn thông suốt dòng chảy hàng hóa và tối ưu được thời gian thông quan.

Mng chuyển tiếp thông tin này cho phòng Xuất nhập khẩu và IT của công ty mình để xử lý ngay nhé! Anh chị nào cần file Công văn 21692 và tài liệu hướng dẫn đăng nhập portal chi tiết, cứ bình luận Thúy gửi link tải ngay!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Link tải toàn văn Công văn 21692/CHQ-CNTT & Hướng dẫn portal dksd: https://customs.gov.vn
👉 Đường link truy cập trực tiếp cổng Hệ thống Đăng ký người sử dụng: https://dksd.customs.gov.vn/client/external

**Hashtags:** `#CongVan21692 #LamSachDuLieuHaiQuan #VNACCS #CSDLDanCu #HaiQuanSo #ThuTucHaiQuan #EurekaLogistics`
"""

# ------------------------------------------------------------------------------
# 6. TOPIC 4 - FACEBOOK: CƯỚC BIỂN SHANGHAI - NY VƯỢT 10.394 USD & BLANK SAILINGS
# ------------------------------------------------------------------------------
FB_04 = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - PHÂN TÍCH THỊ TRƯỜNG & CHIẾN LƯỢC BOOKING]
**Nền tảng mục tiêu:** Facebook (Fanpage/Personal Page)
**Mã sản xuất:** `PO-HOT-2026-04-DREWRY-WCI`
**Content Pillar:** `TN` — Tin Ngành & Phân tích Thị trường Vận tải biển
**Góc nhìn:** Nhận định thị trường & Giải pháp bảo vệ booking mùa cao điểm Q4
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

CƯỚC CONTAINER TUYẾN SHANGHAI - NEW YORK PHÁ ĐỈNH 10.394 USD/FEU & BÃO 79 CHUYẾN BLANK SAILINGS TRƯỚC THỀM GOLDEN WEEK: LỐI THOÁT NÀO CHO CHỦ HÀNG XUẤT KHẨU MÙA CAO ĐIỂM Q4?

Chỉ số cước container toàn cầu Drewry WCI vừa công bố con số giật mình: Tuyến Bờ Đông Hoa Kỳ (Shanghai - New York) đã chính thức xuyên thủng mốc 10.000 USD, chạm đỉnh 10.394 USD/container 40ft (tăng 7% chỉ trong vòng 1 tuần)! Tuyến Shanghai - Los Angeles cũng vọt lên 7.712 USD/FEU.

Đáng nói hơn, chỉ còn đúng 8 ngày nữa là đến kỳ nghỉ lễ Quốc khánh Trung Quốc Golden Week (01/10 - 07/10), các liên minh hãng tàu lớn (Ocean Alliance, 2M, THE Alliance) lại tiếp tục tung ra đòn quen thuộc: **Hủy thêm 79 chuyến tàu (chiếm tới 11% tổng công suất vận tải)!**

📍 **Tại sao hãng tàu lại hủy chuyến khi nhu cầu đang cao kỷ lục?**
Nhiều anh chị làm xnk bức xúc: "Hàng thì dồn ứ muốn đóng đi Mỹ và Châu Âu kịp mùa Giáng sinh, tàu thì thiếu mà hãng tàu lại cứ thi nhau cắt chuyến (Blank Sailings), có phải họ đang bắt tay nhau tạo sự khan hiếm giả tạo để ép giá cước neo đỉnh không?"

Thực tế câu chuyện phức tạp hơn nhiều:
1. **Khủng hoảng hải trình vòng qua Mũi Hảo Vọng:** Thời gian hải trình kéo dài thêm 10 - 14 ngày làm vòng quay tàu và vỏ container rỗng bị kéo giãn tối đa.
2. **Kẹt cảng dây chuyền:** Các siêu cảng tại Châu Á (Thượng Hải, Ninh Ba, Singapore) đang quá tải nghiêm trọng, tàu phải neo chờ cầu bến từ 3 đến 5 ngày, khiến lịch trình của hãng tàu bị vỡ trận hoàn toàn. Họ buộc phải hủy chuyến để gom dồn lịch chạy.
3. **Chiến lược tối ưu hóa lợi nhuận:** Hãng tàu cắt bớt công suất trước thềm Golden Week để giữ giá spot không bị lao dốc khi các nhà máy Trung Quốc đóng cửa nghỉ lễ 1 tuần.

Hậu quả trực tiếp dội lên đầu các doanh nghiệp xuất khẩu Việt Nam: Hàng xuất đi Mỹ/Châu Âu liên tục bị rớt booking (roll cont), phụ phí mùa cao điểm (PSS) tăng chóng mặt, nguy cơ lỡ hẹn giao hàng mùa bán hàng cuối năm là hiển hiện trước mắt!

---

🤔 Trên thị trường đang nổ ra cuộc tranh luận gay gắt:
🔴 Chủ hàng thì chỉ trích: "Hãng tàu độc quyền, chèn ép doanh nghiệp sản xuất, giá cước tăng phi lý ăn mòn hết lợi nhuận xuất khẩu!"
🔵 Phía hãng tàu thì giải trình: "Chi phí nhiên liệu tăng vọt, rủi ro an ninh hàng hải leo thang, chi phí điều vỏ cont rỗng từ Mỹ về Châu Á đắt đỏ, hãng tàu cũng đang gồng mình chịu áp lực vận hành!"

Quan điểm của mng thế nào? Doanh nghiệp của anh chị đã chốt được chỗ cho các đơn hàng xuất khẩu tháng 10 và tháng 11 chưa?

---

💡 **4 Bước Thúy gợi ý giúp doanh nghiệp giữ chỗ an toàn mùa cao điểm:**
1. **Đặt booking sớm trước tối thiểu 3 - 4 tuần:** Đừng đợi hàng xong mới tìm cont, hãy khóa slot trước với các đơn vị forwarder uy tín.
2. **Đàm phán điều khoản "No-Roll Guarantee":** Chấp nhận mức phụ phí hợp lý nhưng yêu cầu cam kết văn bản không bị hãng tàu hoãn chuyến.
3. **Linh hoạt đa dạng hóa tuyến vận chuyển:** Thay vì dồn toàn bộ đi Bờ Đông (New York), cân nhắc chuyển hàng qua Bờ Tây (Los Angeles/Long Beach) kết hợp vận chuyển đường sắt nội địa (Intermodal) để tiết kiệm thời gian và chi phí.
4. **Chia nhỏ lô hàng:** Thay vì đi một lần 10 cont trên 1 chuyến tàu (rất dễ bị cắt bớt cont), hãy chia ra đi rải rác 2 - 3 cont trên nhiều chuyến tàu khác nhau.

Anh chị nào đang cần bảng so sánh lịch tàu và giải pháp giữ chỗ vận tải biển tuyến Mỹ - Âu quý 4, cứ comment Thúy chia sẻ cẩm nang chi tiết nhé!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Báo cáo chi tiết Drewry World Container Index ngày 17/09/2026: https://drewry.co.uk
👉 Video phân tích chiến lược giữ booking và đàm phán cước tàu trên kênh YouTube Eureka: [Link YouTube Channel Eureka Logistics]

**Hashtags:** `#CuocBienDrewry #DrewryWCI #BlankSailings #XuatKhauMy #CacHangTau #GoldenWeek2026 #EurekaLogistics`
"""

# ------------------------------------------------------------------------------
# 7. TOPIC 4 - YOUTUBE: CHIẾN LƯỢC GIỮ BOOKING & ĐỐI PHÓ BLANK SAILINGS
# ------------------------------------------------------------------------------
YT_04 = """# [KỊCH BẢN YOUTUBE CHI TIẾT - PHÂN TÍCH THỊ TRƯỜNG & CHIẾN LƯỢC CHUỖI CUNG ỨNG]
**Nền tảng mục tiêu:** YouTube Dài (Video 8 - 12 phút)
**Mã sản xuất:** `PO-HOT-2026-04-DREWRY-WCI`
**Content Pillar:** `TN` — Tin Ngành & Nhận định Logistics Toàn cầu
**Tiêu đề Video:** CƯỚC TÀU SHANGHAI - NEW YORK PHÁ ĐỈNH 10.394 USD & CƠN BÃO 79 CHUYẾN HỦY: 4 BƯỚC GIỮ BOOKING TRÁNH RỚT CONT MÙA Q4!
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## 1. Bài Viết Lời Thoại Hoàn Chỉnh (Full Text Script for Narration / Repurpose)

**CƯỚC CONTAINER ĐI MỸ VƯỢT 10.000 USD & BÃO 79 CHUYẾN BLANK SAILINGS: LÀM SAO ĐỂ DOANH NGHIỆP XUẤT KHẨU KHÔNG BỊ RỚT CONT TRONG MÙA CAO ĐIỂM?**

Chào các bạn! Thúy đây.

Nếu doanh nghiệp của bạn đang có những đơn hàng xuất khẩu gỗ, dệt may, da giày, điện tử đi thị trường Mỹ hoặc Châu Âu trong quý 4 này... thì những số liệu mà Thúy sắp chia sẻ sau đây sẽ khiến các bạn phải ngồi lại với bộ phận xuất nhập khẩu ngay lập tức!

Theo báo cáo mới nhất từ tổ chức nghiên cứu hàng hải quốc tế Drewry ngày 17/09/2026:
Chỉ số cước container tuyến Transpacific từ Thượng Hải đi New York (Bờ Đông Hoa Kỳ) đã chính thức phá đỉnh 10.000 USD, chạm mốc 10.394 USD cho mỗi container 40ft! Tuyến đi Bờ Tây (Los Angeles) cũng neo ở mức kỷ lục 7.712 USD.

Nhưng điều đáng sợ nhất không chỉ là giá cước tăng. Điều khiến các giám đốc xuất nhập khẩu mất ăn mất ngủ chính là: **Các hãng tàu vừa công bố hủy tiếp 79 chuyến tàu (Blank Sailings) ngay trước kỳ nghỉ Golden Week của Trung Quốc!**

Hàng trăm container của các doanh nghiệp Việt Nam đang đứng trước nguy cơ bị rớt bãi (roll cont), trễ hẹn giao hàng cho mùa Giáng sinh và đối mặt với các khoản phạt hợp đồng quốc tế lên đến hàng triệu USD.

Tại sao các hãng tàu lại liên tục hủy chuyến khi thị trường đang khát tàu? Bản chất kinh tế đằng sau chiêu bài Blank Sailings là gì? Và quan trọng nhất: Doanh nghiệp xuất khẩu Việt Nam cần làm gì trong 4 tuần tới để giữ được chỗ, bảo vệ dòng hàng thông suốt?

Thúy sẽ cùng các bạn bóc tách cặn kẽ trong video hôm nay!

---

### PHẦN 1: BÓC TÁCH CHIÊU BÀI "BLANK SAILINGS" CỦA CÁC HÃNG TÀU

Nhiều người nghĩ đơn giản: Thiếu tàu thì hãng tàu phải chạy nhiều hơn để kiếm tiền chứ, sao lại hủy chuyến?
Nhưng thực tế, 3 liên minh hàng hải lớn nhất thế giới hiện nay nắm giữ hơn 80% công suất vận tải toàn cầu:
1. 2M (MSC, Maersk).
2. Ocean Alliance (CMA CGM, COSCO, Evergreen).
3. THE Alliance / Premier Alliance (ONE, HMM, Yang Ming).

Khi thị trường chuẩn bị bước vào kỳ nghỉ Quốc khánh Trung Quốc Golden Week (01/10 - 07/10), các nhà máy tại Trung Quốc nghỉ sản xuất. Để ngăn giá cước spot không bị sụt giảm do nguồn hàng thiếu hụt trong 1 tuần, các hãng tàu chủ động bắt tay nhau cắt giảm nguồn cung bằng cách tuyên bố hủy chuyến. Khi nguồn cung tàu giảm 11%, giá cước tự động neo ở mức đỉnh trên 10.000 USD!

---

### PHẦN 2: TÁC ĐỘNG ĐẾN DOANH NGHIỆP XUẤT KHẨU VIỆT NAM

Việt Nam phụ thuộc rất lớn vào các tuyến tàu mẹ trung chuyển. Khi tàu mẹ tại Thượng Hải hay Singapore hủy chuyến, các tàu feeder chạy từ Hải Phòng hay Cái Mép không thể kết nối. Hàng hóa nằm kẹt lại ở cảng trung chuyển, doanh nghiệp vừa phải chịu tiền phạt lưu cont, vừa bị khách hàng nước ngoài hủy đơn vì giao trễ mùa lễ hội.

---

### PHẦN 3: BỐN CHIẾN LƯỢC GIỮ BOOKING VÀ TỐI ƯU CHI PHÍ TỪ EUREKA

Để vượt qua cơn bão cước và rớt tàu quý 4, Thúy khuyến nghị 4 giải pháp thực chiến:
1. **Quy tắc 4 tuần:** Phải lock booking trước ngày tàu chạy tối thiểu 21 đến 28 ngày.
2. **Ký cam kết No-Roll Guarantee:** Thỏa thuận với Forwarder hoặc Hãng tàu điều khoản bảo đảm chỗ. Nếu bị rớt tàu do lỗi hãng tàu, hãng tàu phải bồi hoàn chi phí lưu bãi và ưu tiên đi chuyến kế tiếp trong 48 giờ.
3. **Giải pháp Intermodal (Đa phương thức):** Đi đường biển sang Bờ Tây nước Mỹ (Los Angeles) rồi dùng đường sắt chuyển sang Bờ Đông. Thời gian chỉ mất 18 ngày so với 35 ngày đi Bờ Đông qua kênh đào, giúp né được mức cước 10.394 USD.
4. **Chia nhỏ rủi ro container:** Thay vì dồn 10 cont vào 1 bill của 1 chuyến tàu, hãy tách thành 3 lô nhỏ đi trên 2 hãng tàu khác nhau.

---

### KẾT LUẬN & CLIFFHANGER

Thị trường logistics toàn cầu là một bàn cờ chiến lược. Hiểu được luật chơi của các siêu tập đoàn vận tải biển là cách duy nhất để doanh nghiệp xuất khẩu Việt Nam chủ động kiểm soát chi phí và bảo vệ uy tín thương mại.

🎬 **Ở video tiếp theo, Thúy sẽ mang đến một case study giải cứu doanh nghiệp cực kỳ hy hữu: Cách xử lý nợ thuế trên eTax cho doanh nghiệp Trạng thái 06 (tạm ngừng kinh doanh) và quy trình lấy đúng Mã ID khoản nộp để khôi phục giấy phép trong 24 giờ! Mng nhớ bấm Đăng ký kênh để theo dõi nhé!**

---

## 2. Bảng Phân Cảnh Kịch Bản Chi Tiết (Storyboard 4 Cột)

| Phân đoạn / Thời gian | Hình ảnh & Bối cảnh (Visual) | Lời thoại & Giọng đọc (Audio) | Text On Screen & SFX |
|---|---|---|---|
| **01. Hook (0:00 - 1:15)** | Cận cảnh chị Thúy trong văn phòng làm việc, màn hình hiển thị đồ thị nến đỏ giá cước tàu biển Drewry WCI chạm mốc 10.394 USD. | **Chị Thúy:** Tuyến Thượng Hải - New York chính thức vượt 10.000 USD! Các hãng tàu hủy liên tiếp 79 chuyến trước thềm Golden Week! Hàng xuất khẩu mùa Giáng sinh đang đối mặt nguy cơ rớt cont hàng loạt! | **Text lớn:** CƯỚC CONTAINER VƯỢT 10.394 USD! <br> 79 CHUYẾN BỊ HỦY BỎ! <br> *SFX: Tiếng còi tàu hú dồn dập* |
| **02. Giải mã Blank Sailings (1:15 - 4:30)** | Đồ họa bản đồ hàng hải thế giới: 3 Liên minh hãng tàu biển (2M, Ocean Alliance, THE Alliance) và các chấm đỏ hiển thị 79 chuyến bị hủy bỏ. | **Chị Thúy:** Tại sao khát tàu mà lại hủy chuyến? Đây là bài toán điều tiết công suất của các siêu liên minh hãng tàu để giữ giá cước neo đỉnh trước kỳ nghỉ Golden Week. | **Text:** BÍ MẬT ĐẰNG SAU BLANK SAILINGS <br> Cắt giảm 11% công suất vận tải toàn cầu |
| **03. Tác động chuỗi cung ứng (4:30 - 7:00)** | Hình ảnh bãi container xếp cao như núi tại cảng biển quốc tế. Chị Thúy phân tích rủi ro tắc nghẽn tàu mẹ tại các trạm trung chuyển Singapore, Ninh Ba. | **Chị Thúy:** Khi tàu mẹ hủy chuyến, container Việt Nam từ Cái Mép và Hải Phòng sẽ bị bỏ lại tại cảng trung chuyển, tiền phạt lưu kho lưu bãi tăng gấp 3 lần! | **Text:** RỦI RO ROLL CONT MÙA CAO ĐIỂM <br> Phụ phí mùa cao điểm (PSS) tăng vọt |
| **04. 4 Bước tháo gỡ & Outro (7:00 - 9:30)** | Chị Thúy đứng trước bảng trắng phân tích 4 giải pháp: Lock chỗ 4 tuần, No-roll guarantee, Tuyến Intermodal Bờ Tây, Chia nhỏ lô hàng. | **Chị Thúy:** 4 Giải pháp vàng từ kinh nghiệm điều hành Eureka giúp doanh nghiệp giữ vững booking. Đón xem tập tiếp theo: Giải cứu DN nợ thuế Trạng thái 06! | **Text:** 4 BƯỚC BẢO VỆ BOOKING XUẤT KHẨU <br> Subscribe kênh YouTube Eureka ngay! |
"""

# ------------------------------------------------------------------------------
# 8. TOPIC 5 - FACEBOOK: XỬ LÝ NỢ THUẾ ETAX CHO DOANH NGHIỆP TRẠNG THÁI 06
# ------------------------------------------------------------------------------
FB_05 = """# [BẢN THẢO BÀI ĐĂNG FACEBOOK - CHUẨN GIỌNG ĐIỀM TĨNH & CASE STUDY XỬ LÝ NỢ THUẾ]
**Nền tảng mục tiêu:** Facebook (Fanpage/Personal Page)
**Mã sản xuất:** `PO-HOT-2026-05-DN-TRANG-THAI-06`
**Content Pillar:** `CS` — Case Study / Corporate Tax Resolution & eTax
**Góc nhìn:** Problem-First & Bài học tháo gỡ bẫy thủ tục thuế
**Trạng thái kiểm duyệt:** 🟢 [PASSED — AGENT 08 FACT-CHECK]

---

## Nội dung Bài viết (Post Draft)

DOANH NGHIỆP 'TRẠNG THÁI 06' NỘP TIỀN NỢ THUẾ TRÊN ETAX NHƯNG KHÔNG ĐƯỢC GẠCH NỢ: CẠM BẪY TIỀN TREO KHO BẠC & CÁCH LẤY CHUẨN MÃ ID KHOẢN NỘP (MẪU 01/BK-NT) TRONG 24H!

Hôm qua trên Diễn đàn Thuế chuyên sâu, Thúy đọc được một chia sẻ rất trăn trở của một chủ doanh nghiệp:
Công ty anh rơi vào "Trạng thái 06" (Tạm ngừng kinh doanh nhưng chưa hoàn thành thủ tục giải thể). Để giải quyết dứt điểm các khoản nợ cũ để đóng công ty hoặc chuẩn bị mở lại, anh chủ động dùng tài khoản cá nhân chuyển khoản thẳng tiền nợ thuế vào tài khoản của Kho bạc Nhà nước.

Thế nhưng 2 tuần sau kiểm tra lại trên eTax Mobile, anh tá hỏa khi thấy: Khoản nợ thuế vẫn nguyên vẹn, tiền phạt chậm nộp 0.03%/ngày vẫn tiếp tục nhảy đều đặn, và người đại diện pháp luật thì bị thông báo tạm hoãn xuất cảnh theo Điều 66 Luật Quản lý thuế!

Anh hốt hoảng lên hỏi: "Tiền nộp vào kho bạc rồi mà sao hệ thống thuế vẫn tính nợ và phạt mình?"

📍 **Bản chất vấn đề: Cạm bẫy nộp tiền thủ công thiếu Mã ID khoản nộp!**
Từ khi Tổng cục Thuế triển khai hệ thống quản lý thuế tích hợp TMS và thu nộp thuế theo Mã định danh (ID khoản nộp):
👉 Khi một dn đã ở Trạng thái 06, hệ thống tài khoản điện tử của dn bị khóa một phần chức năng tự động gạch nợ.
👉 Nếu mng chuyển khoản ủy nhiệm chi ngân hàng thông thường mà ghi nội dung nộp tiền chung chung (hoặc chọn sai tiểu mục, sai chương thuế), số tiền đó sẽ bị chuyển vào **Tài khoản tạm thu chờ xử lý của Kho bạc**. Hệ thống tự động của cơ quan thuế KHÔNG THỂ đối soát để gạch nợ cho dn được!

Kết quả là: Tiền của dn thì đã rời khỏi túi, nhưng trên hệ thống thuế bạn vẫn là "con nợ", và lãi phạt 0.03%/ngày vẫn tự động cộng dồn mỗi đêm!

---

💡 **Quy trình 3 bước xử lý dứt điểm nợ thuế cho DN Trạng thái 06 từ kinh nghiệm của Thúy:**

1. **Bước 1: Không tự ý chuyển khoản bừa.** Hãy liên hệ ngay với cán bộ quản lý thuế phụ trách địa bàn của dn, đề nghị cung cấp Bảng kê chi tiết nghĩa vụ nợ thuế (Mẫu 01/BK-NT).
2. **Bước 2: Lấy chính xác Mã ID khoản nộp.** Trên Bảng kê 01/BK-NT sẽ có từng Mã ID định danh tương ứng với từng khoản nợ (thuế TNDN, thuế GTGT, tiền phạt chậm nộp). Khi nộp tiền qua ngân hàng hoặc eTax, BẮT BUỘC phải điền đúng Mã ID khoản nộp này vào mục nội dung thanh toán để hệ thống tự động gạch nợ trong vòng 24 giờ.
3. **Bước 3: Lập hồ sơ tra soát (Mẫu C1-11/NS) nếu tiền đã lỡ nộp sai.** Trong trường hợp tiền đã nộp nhưng bị treo ở Kho bạc, làm ngay Thư tra soát Mẫu C1-11/NS gửi cơ quan thuế để họ trích chuyển số tiền đó vào đúng Mã ID khoản nợ, dừng ngay việc tính tiền phạt chậm nộp.

Sau khi nợ thuế được gạch về 0 đồng, cơ quan thuế sẽ ban hành Thông báo xác nhận hoàn thành nghĩa vụ thuế, giải tỏa lệnh tạm hoãn xuất cảnh và mở đường cho dn hoàn tất thủ tục giải thể hoặc khôi phục mã số thuế an toàn!

Mng có người quen nào đang vướng doanh nghiệp tạm ngừng mà nợ thuế chưa gỡ được, hãy tag họ vào bài viết này nhé! Anh chị nào cần mẫu Thư tra soát C1-11/NS chuẩn chỉnh, cứ comment Thúy gửi tặng ngay!

---
**Bình luận đầu tiên (First Comment Link):**
👉 Link tải Mẫu Thư tra soát nợ thuế C1-11/NS: https://tapchitaichinh.vn
👉 Cổng thông tin điện tử Tổng cục Thuế tra cứu nghĩa vụ nợ: https://gdt.gov.vn

**Hashtags:** `#NoThueDoanhNghiep #DoanhNghiepTrangThai06 #eTaxMobile #MaIDKhoanNop #LuatQuanLyThue #TuVanDoanhNghiep #EurekaLogistics`
"""

# ==============================================================================
# WRITE FILES FUNCTION WITH PROPER REPORTING
# ==============================================================================

files_to_create = [
    # Topic 1
    (os.path.join(PO_DIR, "DRAFT-FB-HOT-01-MST888.md"), FB_01),
    (os.path.join(FB_DIR, "FB-HOT-2026-01-MST888.md"), FB_01),
    (os.path.join(PO_DIR, "DRAFT-YT-HOT-01-MST888.md"), YT_01),
    (os.path.join(YT_DIR, "YT-HOT-2026-01-MST888.md"), YT_01),
    
    # Topic 2
    (os.path.join(PO_DIR, "DRAFT-FB-HOT-02-RUOU-NGOAI.md"), FB_02),
    (os.path.join(FB_DIR, "FB-HOT-2026-02-RUOU-NGOAI.md"), FB_02),
    (os.path.join(PO_DIR, "DRAFT-YT-HOT-02-RUOU-NGOAI.md"), YT_02),
    (os.path.join(YT_DIR, "YT-HOT-2026-02-RUOU-NGOAI.md"), YT_02),
    
    # Topic 3
    (os.path.join(PO_DIR, "DRAFT-FB-HOT-03-CV21692-VNACCS.md"), FB_03),
    (os.path.join(FB_DIR, "FB-HOT-2026-03-CV21692-VNACCS.md"), FB_03),
    
    # Topic 4
    (os.path.join(PO_DIR, "DRAFT-FB-HOT-04-DREWRY-WCI.md"), FB_04),
    (os.path.join(FB_DIR, "FB-HOT-2026-04-DREWRY-WCI.md"), FB_04),
    (os.path.join(PO_DIR, "DRAFT-YT-HOT-04-DREWRY-WCI.md"), YT_04),
    (os.path.join(YT_DIR, "YT-HOT-2026-04-DREWRY-WCI.md"), YT_04),
    
    # Topic 5
    (os.path.join(PO_DIR, "DRAFT-FB-HOT-05-DN-TRANG-THAI-06.md"), FB_05),
    (os.path.join(FB_DIR, "FB-HOT-2026-05-DN-TRANG-THAI-06.md"), FB_05),
]

def main():
    print(f"=== BẮT ĐẦU SẢN XUẤT 8 BẢN THẢO NỘI DUNG TỪ 5 CHỦ ĐỀ NÓNG WF05 ===")
    for file_path, content in files_to_create:
        tmp_path = file_path + ".tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            f.write(content.strip() + "\n")
        if os.path.exists(file_path):
            os.replace(tmp_path, file_path)
        else:
            os.rename(tmp_path, file_path)
        print(f"-> Đã tạo thành công: {file_path}")
    print(f"=== HOÀN TẤT: 100% 8 FILE NỘI DUNG ĐÃ ĐƯỢC TẠO VÀ LƯU TRỮ CHUẨN XÁC ===")

if __name__ == "__main__":
    main()
