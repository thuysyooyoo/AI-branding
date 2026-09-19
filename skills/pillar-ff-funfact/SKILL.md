---
name: pillar-ff-funfact
description: Kỹ năng chuyên sâu quét và đóng gói 5 ý tưởng Funfact & Insight mỗi tuần vào Kho Ý tưởng Excel, phục vụ Agent 03/06 chọn lọc 2-3 ý tưởng viết bài (Series hoặc Đơn lẻ).
dependencies:
  - reports/Funfact_Ideas_Bank_Eureka.xlsx
  - scratch/daily_notes.md
  - knowledge/content_pillars.md
---

# 📚 Kỹ năng: Quét & Đóng gói Ý tưởng Funfact & Insight (FF Pillar Research)

## 🎯 Mục tiêu (Goal)
Mỗi tuần, Agent 01 phải **chủ động tìm kiếm và đóng gói chính xác 5 ý tưởng Funfact** vào Kho Ý tưởng [Funfact_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Funfact_Ideas_Bank_Eureka.xlsx). Từ kho này, Agent 03/06 sẽ chọn lọc ra 2-3 ý tưởng chất lượng nhất để viết bài theo Series hoặc Đơn lẻ trên YouTube và Facebook.

## 📥 Đầu vào (Inputs) — Nguồn Research
- **Web Search** (nguồn chính): Quét từ các trang quốc tế uy tín theo 5 hướng khai thác ở Bước 1.
- **Nhật ký thực tế**: [daily_notes.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/daily_notes.md) — tìm câu chuyện thật từ công việc hàng ngày để liên hệ bài học.

## 📤 Đầu ra (Outputs) — Ghi vào Kho Ý tưởng Excel
- **5 Thẻ Ý tưởng Funfact** chuẩn hóa, ghi trực tiếp vào [Funfact_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Funfact_Ideas_Bank_Eureka.xlsx).
- Mỗi thẻ phải đầy đủ 10 trường dữ liệu (xem Bước 2).

---

## 📊 Định ngạch & Phân loại Series (Quota & Series Classification)

### Định ngạch Tuần (Weekly Quota)
- **Quét: 5 ý tưởng** (đủ cho 2-3 bài + dự phòng thay thế).
- **Tỷ lệ Series vs Đơn lẻ:** Khuyến khích tối thiểu 2-3 ý tưởng thuộc cùng Series, 2 ý tưởng Đơn lẻ.

### Danh mục Series có sẵn
| Series | Mô tả | Ví dụ chủ đề |
|---|---|---|
| **Lịch sử Logistics** | Các cột mốc lịch sử vận chuyển, thương mại quốc tế | Container đầu tiên, Kênh đào Suez, Silk Road |
| **Thương mại Quốc tế** | So sánh hệ thống XNK giữa các quốc gia | HS Code toàn cầu, Free Trade Zone Singapore vs VN |
| **Bí mật Hải quan** | Sự thật bất ngờ về quy trình hải quan | Máy soi X-ray hải quan, K9 phát hiện hàng cấm |
| **Đàm phán & Tư duy** | Bài học kinh doanh từ nhân vật nổi tiếng **và mô hình thành công của các nước trên thế giới** | Toyota (Nhật) & JIT, Kaizen; Đức & Industrie 4.0; Singapore FTZ; Hàn Quốc & chiến lược xuất khẩu |
| **Công nghệ & Đổi mới** | Đổi mới công nghệ tác động tới thương mại | Blockchain trong B/L, AI phân loại hàng hóa |
| **Đơn lẻ** | Ý tưởng không thuộc series nào | Bất kỳ funfact độc lập hấp dẫn |

---

## 🛠️ Quy trình Thực hiện (Workflow)

### Bước 1: Quét Nguồn Ý tưởng (Web Search BẮT BUỘC)
Tìm kiếm ý tưởng từ **5 hướng khai thác** sau (mỗi hướng ít nhất 1 ý tưởng):

| # | Hướng khai thác | Nguồn Web Search ưu tiên |
|---|---|---|
| 1 | **Cột mốc lịch sử thương mại** | World Shipping Council, IMO, Wikipedia (kiểm chứng chéo) |
| 2 | **So sánh hệ thống XNK quốc tế** | WCO, WTO, UNCTAD, Customs websites các nước |
| 3 | **Con số bất ngờ / Thống kê gây sốc** | Drewry, Freightos, Statista, World Bank Trade Data |
| 4 | **Bài học từ nhân vật nổi tiếng & mô hình các nước** | Harvard Business Review, Forbes, Business Insider, McKinsey, Nikkei Asia, The Economist |
| 5 | **Đổi mới công nghệ tác động thương mại** | TechCrunch, MIT Technology Review, Official Documentation |

### Bước 2: Đóng gói Thẻ Ý tưởng (BẮT BUỘC đầy đủ 10 trường)
Mỗi ý tưởng phải trả lời đầy đủ 7 câu hỏi Research Kit + 3 trường quản lý:

| # | Trường dữ liệu | Yêu cầu |
|---|---|---|
| 1 | **Chủ đề (Title)** | Tên chủ đề ngắn gọn, gây tò mò |
| 2 | **Sự thật bất ngờ (Hook)** | 1 câu duy nhất khiến người đọc phải dừng lại |
| 3 | **Nhân vật / Sự kiện** | Ai, năm nào, ở đâu |
| 4 | **Bối cảnh câu chuyện** | Vì sao sự kiện đó xảy ra |
| 5 | **Tác động đến ngày nay** | Liên hệ với thương mại hiện đại VN |
| 6 | **So sánh quốc tế** | Các nước khác có gì khác biệt |
| 7 | **Liên hệ bài học kinh doanh** | Rút ra gì cho DN Việt (liên hệ `daily_notes.md` nếu có) |
| 8 | **Nguồn URL xác minh** | Link nguồn chính thống (BẮT BUỘC) |
| 9 | **Phân loại Series** | Thuộc Series nào hoặc Đơn lẻ |
| 10 | **Trạng thái** | Mặc định: "Chưa dùng" |

### Bước 3: Kiểm tra Trùng lặp & Ghi vào Excel
- Đọc file [Funfact_Ideas_Bank_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Funfact_Ideas_Bank_Eureka.xlsx) để kiểm tra ý tưởng mới KHÔNG trùng với ý tưởng đã có.
- Ghi 5 ý tưởng mới vào các dòng tiếp theo trong Sheet "Kho Ý Tưởng Funfact".

---

## 🚫 Quy định Nghiêm ngặt (Anti-Patterns)
- **KHÔNG** tự bịa ra số liệu, năm tháng hoặc nhân vật. Mọi thông tin phải có URL nguồn xác minh.
- **KHÔNG** lấy các funfact quá phổ biến (kiểu "Trái Đất quay quanh Mặt Trời"). Phải liên quan trực tiếp tới XNK, Logistics, Thương mại, Quản trị hoặc Công nghệ.
- **KHÔNG** lặp lại chủ đề đã có trong Kho Ý tưởng Excel.

## 🟢 Checklist Xác minh (Verification Checklist)
- [ ] Đã quét đủ 5 hướng khai thác và tìm được 5 ý tưởng.
- [ ] Mỗi ý tưởng có đầy đủ 10 trường dữ liệu.
- [ ] Mỗi ý tưởng có ít nhất 1 URL nguồn xác minh hoạt động bình thường.
- [ ] Đã kiểm tra không trùng lặp với Kho Ý tưởng hiện có.
- [ ] Có ít nhất 2-3 ý tưởng thuộc cùng 1 Series.
