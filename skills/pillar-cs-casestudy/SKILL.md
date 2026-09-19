---
name: pillar-cs-casestudy
description: Kỹ năng chuyên sâu quét và đóng gói 5 ý tưởng Case Study (Bóc tách tình huống thực tế) mỗi tuần vào Kho Ý tưởng Excel, phục vụ Agent 03/05/06 viết bài phân tích sâu.
dependencies:
  - reports/Case_Study_Ideas_Bank_Eureka.xlsx
  - scratch/daily_notes.md
  - knowledge/content_pillars.md
---

# 🔍 Kỹ năng: Quét & Đóng gói Ý tưởng Case Study (CS Pillar Research)

## 🎯 Mục tiêu (Goal)
Mỗi tuần, Agent 01 phải **chủ động quét và đóng gói chính xác 5 ý tưởng Case Study (Tình huống thực tế)** vào Kho Ý tưởng [Case_Study_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Case_Study_Ideas_Bank_Eureka.xlsx). Những câu chuyện này là chất liệu vàng để Agent 03/05/06 viết bài phân tích nghiệp vụ, vạch trần chiêu trò/rủi ro và rút ra bài học kinh nghiệm sâu sắc cho doanh nghiệp xuất nhập khẩu Việt Nam.

## 📥 Đầu vào (Inputs) — Nguồn Research
- **Nhật ký thực tế**: [daily_notes.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/daily_notes.md) — nguồn ưu tiên số 1 để tìm kiếm các câu chuyện người thật việc thật từ kinh nghiệm của tác giả Thúy.
- **Web Search**: Tìm kiếm các vụ tranh chấp thương mại quốc tế, tai nạn hàng hóa, lừa đảo biên mậu, hoặc xử phạt hải quan từ các trang tin uy tín.

## 📤 Đầu ra (Outputs) — Ghi vào Kho Ý tưởng Excel
- **5 Thẻ Ý tưởng Case Study** chuẩn hóa, ghi trực tiếp vào [Case_Study_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Case_Study_Ideas_Bank_Eureka.xlsx).
- Mỗi thẻ phải điền đầy đủ 14 trường dữ liệu (xem Bước 2).

---

## 📊 Định ngạch & Phân loại Series (Quota & Series Classification)

### Định ngạch Tuần (Weekly Quota)
- **Quét: 5 ý tưởng** (các case study bóc tách sâu, có con số thiệt hại rõ ràng).
- **Phân loại Series:** Phân bổ các ý tưởng theo đúng các Series có sẵn trong dropdown:

| Tên Series | Trọng tâm nội dung |
|---|---|
| **Tai nạn XNK** | Các sự cố khách quan (bão lũ, hãng tàu phá sản) hoặc chủ quan (mất chứng từ, khai sai thuế) |
| **Bóc phốt chiêu trò** | Vạch trần các thủ đoạn lừa đảo của nhà cung cấp ngoại, Forwarder lậu, C/O giả |
| **Giải cứu hàng hóa** | Các ca giải phóng hàng bị giữ tại cảng hoặc xử lý thành công tranh chấp thương mại |
| **Đơn lẻ** | Các bài phân tích tình huống độc lập khác |

---

## 🛠️ Quy trình Thực hiện (Workflow)

### Bước 1: Thu thập Ý tưởng (Web Search & Daily Notes)
Tìm kiếm chất liệu theo **2 hướng chính**:

| Hướng tìm kiếm | Phương pháp | Trọng tâm tìm kiếm |
|---|---|---|
| **Câu chuyện thực tế từ chị Thúy** | Đọc `daily_notes.md` | Các ghi chép về khách hàng bị giữ hàng, đối tác đàm phán thất bại, lỗi chứng từ |
| **Các vụ án/tranh chấp quốc tế** | Web Search | Các vụ lừa đảo mua hàng CIF/FOB, hàng bị FDA tiêu hủy, tranh chấp tại tòa án trọng tài thương mại |

### Bước 2: Đóng gói Thẻ Ý tưởng (BẮT BUỘC đầy đủ 14 trường)
Mỗi ý tưởng phải trả lời đầy đủ 7 câu hỏi Research Kit + các trường quản lý:

| # | Trường dữ liệu | Yêu cầu |
|---|---|---|
| 1 | **STT** | Số thứ tự tăng dần |
| 2 | **Tuần/Ngày Thêm** | Ví dụ: "W33 — 13/08/2026" |
| 3 | **Tên Case Study (Title)** | Tiêu đề giật gân, tò mò, nói thẳng vào tổn thất |
| 4 | **Bên liên quan** | Các chủ thể tham gia (Ví dụ: DN VN + Hải quan cảng Hải Phòng) |
| 5 | **Diễn biến tình huống** | Tóm tắt sự việc xảy ra theo trình tự thời gian |
| 6 | **Thiệt hại / Hệ quả cụ thể** | **Bắt buộc phải có con số** (tiền phạt, tiền lưu kho, thời gian chậm trễ) |
| 7 | **Nguyên nhân gốc rễ** | Lỗi sai cốt lõi dẫn đến tai nạn (sự ngây thơ, thiếu kiến thức, v.v.) |
| 8 | **Giải pháp phòng ngừa / Cách tránh** | Checklist hành động cụ thể để các DN khác không mắc lại |
| 9 | **Góc nhìn đối lập (Phe A vs Phe B)** | 2 luồng ý kiến trái chiều để kích thích tranh luận mạng xã hội |
| 10 | **Bài học rút ra cho DN Việt** | Đúc kết triết lý kinh doanh hoặc kinh nghiệm xương máu |
| 11 | **Nguồn URL xác minh / Trực quan** | Link bài báo chính thống hoặc ghi nhận từ `daily_notes.md` |
| 12 | **Phân loại Series** | Chọn từ danh mục (Tai nạn XNK / Bóc phốt chiêu trò / Giải cứu hàng hóa / Đơn lẻ) |
| 13 | **Trạng thái** | Mặc định: "Chưa dùng" |
| 14 | **Ghi chú** | Các thông tin bổ trợ khác |

### Bước 3: Kiểm tra Trùng lặp & Ghi vào Excel
- Đọc file [Case_Study_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Case_Study_Ideas_Bank_Eureka.xlsx) để tránh lặp câu chuyện.
- Ghi 5 ý tưởng mới vào các dòng tiếp theo trong Sheet "Kho Ý Tưởng Case Study".

---

## 🚫 Quy định Nghiêm ngặt (Anti-Patterns)
- **KHÔNG** tự bịa ra tên doanh nghiệp thật (sử dụng cụm từ chung như "Doanh nghiệp X", "DN nhập khẩu đồ gia dụng VN") để bảo mật danh tính khách hàng, trừ khi vụ việc đã được báo chí đăng tải công khai.
- **BẮT BUỘC** phải có số liệu thiệt hại cụ thể. Các case study chung chung kiểu "bị phạt nhiều tiền" không có giá trị truyền thông.
- **KHÔNG** đề xuất giải pháp phòng ngừa hời hợt (ví dụ: "DN cần chú ý hơn"). Phải chỉ rõ hành động: check mã gì, sửa hợp đồng điều khoản nào, dùng phương thức thanh toán gì.
- **CẤM** chỉ tham chiếu 1 nguồn duy nhất. Nguồn tham khảo phải từ **tối thiểu 2 nguồn trở lên** để đối chiếu chéo độc lập.
- **KHUYẾN KHÍCH** tham chiếu chéo với **Thư Viện Pháp Luật (thuvienphapluat.vn)** khi vụ việc có liên quan đến xử phạt hải quan hoặc vi phạm pháp lý XNK.

## 🟢 Checklist Xác minh (Verification Checklist)
- [ ] Đã quét đủ 5 câu chuyện tình huống thực tế.
- [ ] Điền đầy đủ 14 cột của bảng tính Excel.
- [ ] Mọi case study đều có con số thiệt hại cụ thể (tiền phạt/chi phí/ngày trễ).
- [ ] Đã xây dựng giải pháp phòng ngừa mang tính kỹ thuật/nghiệp vụ rõ ràng.
- [ ] Có link URL kiểm chứng hoạt động bình thường cho từng ý tưởng từ tối thiểu 2 nguồn độc lập.
- [ ] Đã đối chiếu tránh trùng lặp với kho dữ liệu cũ.
