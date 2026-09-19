---
name: ssot-research
description: Hướng dẫn quy trình 3 lớp SSoT nhằm tra cứu văn bản pháp luật, chính sách hải quan từ nguồn lực cục bộ và các kênh hỗ trợ trực tuyến để đảm bảo tính xác thực thông tin.
dependencies:
  - scripts/legal_tracker.py
  - reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx
---

# 🔍 Kỹ năng: Kiểm chứng Tri thức Gốc (SSoT Research)

## 🎯 Mục tiêu (Goal)
Đảm bảo tất cả các bài viết, kịch bản video hay đề xuất có liên quan tới pháp luật xuất nhập khẩu, hải quan, và giải pháp kỹ thuật đều được xác minh độ tin cậy thông qua quy trình chuẩn 3 lớp, loại bỏ hoàn toàn việc AI tự tạo (hallucination) dữ liệu giả.

## 📥 Đầu vào (Inputs)
- Ý tưởng thô từ [daily_notes.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/daily_notes.md).
- Yêu cầu nghiên cứu từ người dùng hoặc các agent trước trong pipeline.

## 📤 Đầu ra (Outputs)
- Báo cáo nghiên cứu chứa bảng đối chiếu SSoT đạt chuẩn 3 lớp:
  - ✅ Đã xác minh (Đạt 3 lớp)
  - ⚠️ Cần xác minh thêm (Không dùng trực tiếp)
  - ❌ Không xác minh được (Hủy bỏ)

---

## 🛠️ Quy trình Thực hiện (Workflow & Rules)

### Bước 1: Tra cứu Dữ liệu Cục bộ (Local SSoT First)
Trước tiên, bạn bắt buộc phải kiểm tra cơ sở dữ liệu có sẵn tại máy:
1. Đọc tệp Excel cập nhật luật: [Danh_sach_Van_ban_Phap_luat_XNK.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/Danh_sach_Van_ban_Phap_luat_XNK.xlsx)
2. Tra cứu từ khóa xu hướng để hiểu nỗi đau độc giả: [SEO_Keyword_Trends_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx)

### Bước 2: Áp dụng Quy chuẩn Kiểm chứng SSoT 3 Lớp
Khi sử dụng nguồn tin từ bên ngoài (Internet), bạn phải áp dụng kiểm chứng:

#### 1. Lớp 1 - Kiểm định tính hiệu lực thời gian:
- Xác định văn bản quy phạm pháp luật hoặc hướng dẫn kỹ thuật đó được ban hành ngày nào.
- Kiểm tra xem đã có nghị định, thông tư hoặc phiên bản cập nhật nào thay thế/bãi bỏ nội dung đó hay chưa.
- **Quy tắc:** Chỉ sử dụng văn bản đang *Có hiệu lực* hiện hành.

#### 2. Lớp 2 - Đối chiếu chéo độc lập:
- Tìm kiếm và xác thực nội dung trên tối thiểu **2 nguồn uy tín độc lập** (không chấp nhận thông tin chỉ từ 1 nguồn duy nhất).
- *Đối với hải quan/pháp luật:* Cổng thông tin Tổng cục Hải quan, Thư viện Pháp luật (thuvienphapluat.vn - bắt buộc ưu tiên hàng đầu để đối soát hiệu lực văn bản), các cổng báo chí chính thống của Chính phủ.
- *Đối với giải pháp AI/Công nghệ:* Trang tài liệu chính thức (Official Documentation) của nhà phát triển (như OpenAI, n8n, make.com), không lấy tin tức tổng hợp không rõ nguồn.

#### 3. Lớp 3 - Định dạng minh bạch nguồn (Trích dẫn chuẩn):
Mọi thông tin trích dẫn phải được viết theo cấu trúc:
`[Tên văn bản pháp lý/Số hiệu] - Điều/Khoản [nếu có] - Ban hành ngày: [DD/MM/YYYY] - Link nguồn: [URL trực tiếp]`

---

## 🚫 Quy định Nghiêm ngặt (Anti-Patterns)
- **KHÔNG** sử dụng bài viết của đối thủ cạnh tranh (Ví dụ: các bài viết ở cột Top 1 Google Search được liệt kê trong file SEO Excel) làm nguồn SSoT chính cho luật hoặc hướng dẫn kỹ thuật.
- **KHÔNG** tự ý đoán số hiệu thông tư, ngày ban hành hoặc link liên kết.
- **KHÔNG** sử dụng tài liệu lỗi thời đã bị thay thế.

## 🟢 Checklist Xác minh (Verification Checklist)
- [ ] Số hiệu văn bản pháp lý được ghi chính xác (ví dụ: Nghị định 08/2015/NĐ-CP).
- [ ] Toàn bộ đường dẫn URL nguồn là link trực tiếp và hoạt động bình thường.
- [ ] Dữ liệu số liệu hoặc quy trình kỹ thuật đã được chạy thử nghiệm hoặc đối chiếu với tài liệu chính thống của hãng.
