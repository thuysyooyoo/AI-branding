---
name: pillar-tn-news
description: Kỹ năng chuyên sâu quét và đóng gói 5 ý tưởng Tin Ngành (Chính sách & Pháp luật) mỗi tuần vào Kho Ý tưởng Excel, phục vụ Agent 03/05/06 viết bài tranh luận chính sách.
dependencies:
  - reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx
  - reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx
  - scratch/daily_notes.md
  - knowledge/content_pillars.md
---

# 📰 Kỹ năng: Quét & Đóng gói Ý tưởng Tin Ngành (TN Pillar Research)

## 🎯 Mục tiêu (Goal)
Mỗi tuần, Agent 01 phải **chủ động quét và đóng gói chính xác 5 ý tưởng Tin Ngành (Chính sách & Pháp luật mới)** vào Kho Ý tưởng [Tin_Nganh_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx). Bản tin này cung cấp cốt liệu pháp lý thực tế và 2 luồng quan điểm đối lập để Agent 03/05/06 dựng nội dung tranh luận trên Facebook.

## 📥 Đầu vào (Inputs) — Nguồn Research
- **Văn bản pháp luật có sẵn**: [Danh_sach_Van_ban_Phap_luat_XNK.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx) — nguồn văn bản gốc.
- **Web Search** (nguồn thời sự cập nhật): Tổng cục Hải quan (customs.gov.vn), Cổng TTĐT Chính phủ (chinhphu.vn), báo Hải quan Online (haiquanonline.com.vn), báo Công Thương (congthuong.vn).
- **Nhật ký thực tế**: [daily_notes.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/daily_notes.md) — thu thập phản ứng thực tế hoặc ghi nhận của chị Thúy.

## 📤 Đầu ra (Outputs) — Ghi vào Kho Ý tưởng Excel
- **5 Thẻ Ý tưởng Tin Ngành** chuẩn hóa, ghi trực tiếp vào [Tin_Nganh_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx).
- Mỗi thẻ phải điền đầy đủ 15 trường dữ liệu (xem Bước 2).

---

## 📊 Định ngạch & Phân loại Series (Quota & Series Classification)

### Định ngạch Tuần (Weekly Quota)
- **Quét: 5 ý tưởng** (các chính sách mới ban hành, thay đổi mức thuế, hạn ngạch hoặc quy định hải quan cửa khẩu).
- **Trọng tâm truyền thông:** Tập trung tạo tranh luận văn minh trên mạng xã hội bằng cách vạch rõ mâu thuẫn lợi ích/góc nhìn.

---

## 🛠️ Quy trình Thực hiện (Workflow)

### Bước 1: Quét Nguồn Tin mới (Web Search BẮT BUỘC)
Tìm kiếm chính sách/tin tức theo **3 nhóm nguồn** sau:

| Nhóm nguồn | Địa chỉ Web Search | Nội dung cần tìm |
|---|---|---|
| **Chính phủ & Bộ ngành** | customs.gov.vn, chinhphu.vn, mof.gov.vn | Nghị định, Thông tư mới ban hành liên quan XNK, Thuế, Hải quan |
| **Báo chí chuyên ngành** | haiquanonline.com.vn, congthuong.vn, tapchitaichinh.vn | Các bài phân tích ảnh hưởng của chính sách mới, phản ứng của doanh nghiệp |
| **Xu hướng toàn cầu** | UNCTAD, WTO, WCO | Các hiệp định thương mại mới ký kết hoặc thay đổi chính sách từ thị trường lớn (Mỹ, EU, Trung Quốc) |

### Bước 2: Đóng gói Thẻ Ý tưởng (BẮT BUỘC đầy đủ 15 trường)
Mỗi ý tưởng phải trả lời đầy đủ 8 câu hỏi Research Kit + các trường quản lý:

| # | Trường dữ liệu | Yêu cầu |
|---|---|---|
| 1 | **STT** | Số thứ tự tăng dần |
| 2 | **Tuần/Ngày Thêm** | Ví dụ: "W33 — 13/08/2026" |
| 3 | **Số hiệu văn bản / Tên chính sách** | Tên chính xác của văn bản pháp lý hoặc chính sách |
| 4 | **Cơ quan ban hành** | Cơ quan nhà nước ban hành (Chính phủ, Bộ Công Thương, v.v.) |
| 5 | **Ngày hiệu lực / Timeline** | Thời gian áp dụng chính thức hoặc lộ trình quan trọng |
| 6 | **Điểm khác biệt cốt lõi (Cũ vs Mới)** | So sánh rõ ràng thay đổi lớn nhất |
| 7 | **Đối tượng bị ảnh hưởng mạnh nhất** | Nhóm doanh nghiệp hoặc mặt hàng cụ thể |
| 8 | **Nỗi đau / Rủi ro nếu sai phạm** | Thiệt hại tài chính, pháp lý cụ thể cho doanh nghiệp |
| 9 | **Phản ứng thị trường** | Phản hồi của hiệp hội, doanh nghiệp hoặc chuyên gia |
| 10 | **Phe Ủng hộ (Lập luận)** | Quan điểm ủng hộ sự thay đổi (Vì sao chính sách này tốt?) |
| 11 | **Phe Phản đối (Lập luận)** | Quan điểm phản đối/lo ngại từ phía doanh nghiệp bị ảnh hưởng |
| 12 | **Tham chiếu quốc tế** | Cách thế giới giải quyết vấn đề tương tự |
| 13 | **Nguồn URL xác minh** | Link chính thức dẫn tới văn bản hoặc bài báo uy tín (BẮT BUỘC) |
| 14 | **Trạng thái** | Mặc định: "Chưa dùng" |
| 15 | **Ghi chú** | Các lưu ý thêm |

### Bước 3: Kiểm tra Trùng lặp & Ghi vào Excel
- Đọc file [Tin_Nganh_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Tin_Nganh_Ideas_Bank_Eureka.xlsx) để đảm bảo không ghi đè hoặc trùng lặp văn bản pháp luật.
- Ghi 5 ý tưởng mới vào các dòng tiếp theo trong Sheet "Kho Ý Tưởng Tin Ngành".

---

## 🚫 Quy định Nghiêm ngặt (Anti-Patterns)
- **TUYỆT ĐỐI KHÔNG** trích dẫn sai số hiệu, điều khoản hoặc ngày hiệu lực của văn bản. Mọi thông tin phải dựa trên nguồn chính thống.
- **KHÔNG** thiên vị một phía khi viết 2 luồng quan điểm đối lập. Phải đảm bảo tính trung lập và logic cho cả Phe Ủng hộ và Phe Phản đối.
- **KHÔNG** đưa tin đồn đoán chưa được xác thực từ mạng xã hội làm nguồn tin chính sách.
- **CẤM** chỉ lấy nguồn từ 1 địa chỉ duy nhất. Phải đối soát chéo trên **tối thiểu 2 nguồn độc lập**.
- **BẮT BUỘC** tham chiếu từ **Thư Viện Pháp Luật (thuvienphapluat.vn)** khi phân tích các thông tư, nghị định mới ban hành.

## 🟢 Checklist Xác minh (Verification Checklist)
- [ ] Đã tìm đủ 5 chính sách/tin ngành mới.
- [ ] Điền đầy đủ 15 cột trong file Excel.
- [ ] Đã ghi nhận rõ ràng lập luận logic của cả 2 phe (Ủng hộ và Phản đối).
- [ ] Có link URL kiểm chứng hoạt động bình thường cho từng ý tưởng từ tối thiểu 2 nguồn độc lập.
- [ ] Đã đối chiếu với danh sách văn bản cũ để tránh trùng lặp.
