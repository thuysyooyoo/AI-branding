---
name: seo-compliance
description: Kỹ năng tích hợp từ khóa SEO hiệu quả từ báo cáo xu hướng và rà soát, lọc bỏ các từ cấm/từ nhạy cảm dễ bị bóp tương tác trên các nền tảng mạng xã hội.
dependencies:
  - knowledge/banned_words.md
  - reports/SEO_Keyword_Trends_Eureka.xlsx
---

# 📈 Kỹ năng: SEO & Rà soát Từ Cấm (SEO & Compliance)

## 🎯 Mục tiêu (Goal)
Tối ưu hóa khả năng hiển thị tự nhiên của bài viết (SEO) bằng cách chèn hợp lý các từ khóa hạt giống và từ khóa phái sinh từ báo cáo xu hướng, đồng thời bảo vệ tài khoản khỏi bị quét vi phạm chính sách hoặc giảm tương tác (bóp reach) trên TikTok, Facebook, LinkedIn, YouTube.

## 📥 Đầu vào (Inputs)
- Bản thảo bài viết (FB/LinkedIn) hoặc Kịch bản video (TikTok/YouTube).
- Báo cáo từ khóa SEO từ [SEO_Keyword_Trends_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx)
- Danh mục từ cấm và từ thay thế từ [banned_words.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/knowledge/banned_words.md)

## 📤 Đầu ra (Outputs)
- Bản thảo hoàn chỉnh đã được chèn từ khóa và làm sạch từ nhạy cảm.
- Báo cáo rà soát nhanh (Liệt kê các từ khóa đã sử dụng, từ cấm phát hiện và cách xử lý).

---

## 🛠️ Quy trình Thực hiện (Workflow & Rules)

### Bước 1: Thu thập từ khóa mục tiêu
1. Đọc báo cáo [SEO_Keyword_Trends_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) để lấy các từ khóa nóng của tuần.
2. Tự động sinh ra 3-5 từ khóa phái sinh có lượng tìm kiếm cao dựa trên chủ đề bài viết.

### Bước 2: Tích hợp từ khóa vào nội dung
- Phân bổ từ khóa tự nhiên vào:
  - Tiêu đề chính (Hook).
  - 100 chữ đầu tiên của bài viết.
  - Phần thân bài (mật độ từ khóa đạt 1-2%).
  - Bộ hashtags ở cuối bài/video.

### Bước 3: Rà soát & Lọc từ cấm (Compliance Check)
Đối chiếu bản thảo với danh sách trong [banned_words.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/knowledge/banned_words.md). Đặc biệt chú ý các nhóm từ sau:
1. **Từ khóa bị quét bóp reach trực tiếp:** *lách luật, trốn thuế, cam kết 100%, chắc chắn thành công, kiếm tiền nhanh*.
2. **Từ ngữ mang tính thần thánh hóa AI:** *AI thay thế con người, tự động hoàn toàn không cần làm gì, cỗ máy in tiền*.
3. **Thay thế bằng:** Các cụm từ trung tính, chính trực hơn (ví dụ: thay "lách luật" bằng "tối ưu quy trình chuẩn chỉnh", thay "cam kết chắc chắn" bằng "kiểm soát rủi ro", thay "tự động hoàn toàn" bằng "hỗ trợ tăng năng suất").

---

## 🚫 Quy định Nghiêm ngặt (Anti-Patterns)
- **KHÔNG** nhồi nhét từ khóa (keyword stuffing) làm mất đi giọng văn tự nhiên của tác giả.
- **KHÔNG** cố tình lách từ cấm bằng cách dùng ký tự đặc biệt (ví dụ: t.h.u.ế, l-á-c-h) vì thuật toán quét ngữ nghĩa hiện đại vẫn sẽ phát hiện và phạt nặng hơn.

## 🟢 Checklist Xác minh (Verification Checklist)
- [ ] Đã quét toàn bộ bài viết qua danh mục từ cấm và thay thế thành công.
- [ ] Có tối thiểu 1-2 từ khóa SEO chính xuất hiện ở 3 dòng đầu tiên.
- [ ] Phần kịch bản video không sử dụng từ ngữ nhạy cảm dễ bị AI của TikTok quét vi phạm cộng đồng.
