# HƯỚNG DẪN PHỐI HỢP HỆ THỐNG AGENT - GIAI ĐOẠN 1: LẬP LỊCH & PHÊ DUYỆT Ý TƯỞNG

> [!CAUTION]
> ### ⛔ [DEPRECATED] — File này đã được chuyển đổi sang Kiến trúc Modular Workflows
> Toàn bộ quy trình phối hợp Lập lịch tuần đã được chuẩn hóa và nâng cấp tại: [**`workflows/wf01_weekly_planning.md`**](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf01_weekly_planning.md).  
> File này được lưu trữ lại chỉ nhằm mục đích tham chiếu lịch sử.

Tài liệu này hướng dẫn quy trình phối hợp giữa các Agent từ lúc tiếp nhận ý tưởng thô cho đến khi hoàn thành **Lịch nội dung tuần** trên **Sheet 1** và chờ phê duyệt từ chị **Nguyễn Thị Ngọc Thúy**.

---

## 🧭 Các Agent Tham Gia Giai Đoạn 1

*   **Agent 01 (Market Researcher):** Nghiên cứu chủ đề, thiết lập dữ liệu chính sách/kỹ thuật chuẩn xác (SSoT).
*   **Agent 02 (Policy & SEO Consultant):** Định hướng từ khóa SEO, tìm rủi ro chính sách và từ cấm.
*   **Agent 03 (Creative Director):** Sáng tạo 3-4 góc tiếp cận (Hooks) và định hướng hình ảnh thiết kế.
*   **Agent 04 (Content Planner):** Cân đối lịch tuần, điền các phương án A, B, C vào Sheet 1 để chờ chị duyệt.

---

## 🔄 Quy trình Phối hợp Chi tiết

### Bước 1: Nghiên cứu chuyên sâu & Thiết lập SSoT (Agent 01)
*   **Đầu vào (Input):** 
    *   Ý tưởng thô được chọn từ kho ý tưởng [content_pillars.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/knowledge/content_pillars.md).
    *   Câu chuyện thực tế hoặc nhật ký hoạt động ngày gần nhất tại [daily_notes.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/daily_notes.md) (nếu thuộc Pillar Đời sống).
*   **Hành động (Action):**
    *   Quét dữ liệu trực tuyến và dữ liệu nội bộ để thu thập số hiệu văn bản pháp luật, mã HS, quy trình kỹ thuật thực tế.
    *   Xây dựng nguồn thông tin xác thực duy nhất (SSoT). **Tuyệt đối không sử dụng link của đối thủ cạnh tranh.**
*   **Đầu ra (Output):** Bản báo cáo Nghiên cứu & SSoT thô.

### Bước 2: Phân bổ Từ khóa & Compliance Check (Agent 02)
*   **Đầu vào (Input):** Bản báo cáo SSoT thô từ `Agent 01`.
*   **Hành động (Action):**
    *   Đối chiếu với bảng xu hướng từ khóa tại [SEO_Keyword_Trends_Eureka.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/SEO_Keyword_Trends_Eureka.xlsx) để chọn ra 1 từ khóa chính và 3-5 từ khóa phụ phù hợp.
    *   Xác định danh sách từ cấm cần tránh dựa theo [banned_words.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/knowledge/banned_words.md) và phản hồi từ chu kỳ trước của `Agent 07`.
*   **Đầu ra (Output):** Bản thiết kế SEO & Compliance Brief.

### Bước 3: Sáng tạo Góc Tiếp Cận & Hooks (Agent 03)
*   **Đầu vào (Input):** Đầu ra của `Agent 01` & `Agent 02`.
*   **Hành động (Action):**
    *   Xây dựng 3-4 góc tiếp cận khác nhau dựa trên 2 hình mẫu nhân hiệu cốt lõi: **Giáp Mộc** (nâng đỡ, chỉ dẫn dễ hiểu) và **Liêm Trinh** (cảnh báo rủi ro sắc bén).
    *   Mỗi góc tiếp cận phải đi kèm tiêu đề đề xuất (Title), đoạn mở đầu (Hook) và ý tưởng hình ảnh (Visual Brief).
*   **Đầu ra (Output):** Concept Angles & Hooks.

### Bước 4: Lập Lịch Tuần & Điền Sheet 1 (Agent 04)
*   **Đầu vào (Input):** Tổng hợp dữ liệu từ `Agent 01`, `Agent 02`, `Agent 03`.
*   **Hành động (Action):**
    *   Sắp xếp các bài viết vào lịch tuần, cân đối tỷ lệ **60% Giáp Mộc** (Góc 2 - Giải pháp & Góc 4 - Đơn giản hóa) và **40% Liêm Trinh** (Góc 1 - Cảnh báo rủi ro).
    *   Cung cấp đủ 3 phương án lựa chọn: **Option A (Chủ đề chính)**, **Option B (Thay thế 1)**, **Option C (Thay thế 2)** cho mỗi slot đăng bài.
    *   **Cập nhật Excel:** Ghi toàn bộ dữ liệu này vào **Sheet 1: Weekly Calendar** của tệp [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx).
*   **Đầu ra (Output):** Sheet 1 được điền đầy đủ dữ liệu lịch trình tuần.

---

## 👥 Checkpoint 1: Chờ Chị Thúy Phê Duyệt Concept

Khi Giai đoạn 1 hoàn thành, hệ thống sẽ dừng lại để chờ chị phê duyệt ý tưởng tại **Sheet 1 (Weekly Calendar)**:

1.  Chị mở tệp Excel [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx), truy cập **Sheet 1**.
2.  Đọc các phương án tiêu đề Option A, B, C được đề xuất cho từng ngày.
3.  Tại cột **Trạng thái Duyệt** (Cột L), nhấp chọn dropdown để quyết định phê duyệt:
    *   `Duyệt Option A` (Nếu đồng ý với chủ đề chính)
    *   `Duyệt Option B` (Nếu muốn đổi sang phương án dự phòng 1)
    *   `Duyệt Option C` (Nếu muốn đổi sang phương án dự phòng 2)
    *   `Không duyệt` (Nếu muốn hủy bỏ slot đăng đó)

> 💡 *Sau khi chị chọn duyệt xong, hệ thống sẽ tự động quét trạng thái này để kích hoạt tiếp Giai đoạn 2 (Sản xuất chi tiết).*
