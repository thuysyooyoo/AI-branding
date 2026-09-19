# HƯỚNG DẪN PHỐI HỢP HỆ THỐNG AGENT — GIAI ĐOẠN 2: SẢN XUẤT, THẨM ĐỊNH & DUYỆT ĐĂNG (CHIẾN LƯỢC 2.0 - TẬP TRUNG FACEBOOK & YOUTUBE)

> [!CAUTION]
> ### ⛔ [DEPRECATED] — File này đã được chuyển đổi sang Kiến trúc Modular Workflows
> Toàn bộ quy trình phối hợp Sản xuất & Thẩm định đã được chuẩn hóa và tách biệt tại:
> - Bài viết Facebook: [**`workflows/wf02_facebook_production.md`**](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf02_facebook_production.md)
> - Kịch bản YouTube: [**`workflows/wf03_youtube_production.md`**](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf03_youtube_production.md)  
> File này được lưu trữ lại chỉ nhằm mục đích tham chiếu lịch sử.

Tài liệu này hướng dẫn quy trình phối hợp ở **Giai đoạn 2** theo Chiến lược Content 2.0 (tinh gọn, loại bỏ các kênh phụ, tập trung 100% vào **Facebook** và **YouTube**), bắt đầu ngay sau khi chị **Nguyễn Thị Ngọc Thúy** đã phê duyệt lựa chọn ý tưởng (Option) tại **Sheet 1**.

---

## 🧭 Các Agent Tham Gia Giai Đoạn 2

*   **Agent 04 (Content Planner):** Đóng gói Phiếu Giao việc (PO) vật lý làm nguồn dữ liệu xác thực (SSoT).
*   **Agent 05 (Copywriter):** Soạn thảo bài viết Facebook (giao tiếp mộc mạc, từ viết tắt tự nhiên, đặt vấn đề tranh luận) và cập nhật trực tiếp vào **Sheet 2 (Facebook)**.
*   **Agent 06 (Video Scriptwriter):** Soạn thảo kịch bản YouTube dài/Series nhiều tập (phân cảnh mốc thời gian, Outro Cliffhanger gợi mở kịch tính) và cập nhật trực tiếp vào **Sheet 3 (YouTube)**.
*   **Agent 08 (Fact Checker):** Rà soát pháp lý, số liệu, từ cấm, tính khả thi của bài viết và kịch bản video (Checkpoint 2).
*   **Agent 07 (Performance Analyst):** Phân tích hiệu suất tương tác để cập nhật từ cấm động và phản hồi tối ưu hóa chu kỳ sau.

---

## 🔄 Quy trình Phối hợp Chi tiết (Chiến lược 2.0)

### Bước 4.5: Đóng gói Phiếu Giao việc Vật lý (Agent 04)
*   **Hành động (Action):** 
    *   Hệ thống quét Sheet 1 của tệp [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) và lọc ra các dòng có trạng thái `Duyệt Option A/B/C`.
    *   Agent 04 tự động xuất một file **Phiếu Giao việc (PO)** vật lý dạng markdown tại `scratch/production_orders/PO-YYYY-WXX-XX.md` làm căn cứ viết bài và kịch bản video.

### Bước 5: Soạn thảo & Ghi nhận Bản thảo/Kịch bản vào Excel (Agent 05 & Agent 06)
*   **Đầu vào (Input):** Phiếu Giao việc PO vật lý tương ứng.
*   **Hành động (Action):**
    *   **Kênh bài viết Facebook (Agent 05):** Soạn thảo bài viết theo phong cách Problem-First, mộc mạc đời thường, sử dụng từ viết tắt chuẩn MXH (`mng`, `dn`, `sp`, `kh`, `hq`, `b/l`, `tq`, `hq`, `xnk`). Ghi nhận bản thảo hoàn chỉnh vào **Sheet 2 (Facebook)** và đồng thời tạo file nháp `.md` tại `scratch/production_orders/DRAFT-FB-PO-YYYY-WXX-XX.md`, chèn hyperlink vào cột F (`Link Bài viết Chi tiết`).
    *   **Kênh video YouTube (Agent 06):** 
        *   Thiết kế kịch bản trao giá trị thực chiến hoặc **Series nối tiếp (Tập 1, Tập 2...)** kèm **Cliffhanger Outro (đoạn kết kịch tính)** để giữ chân khán giả.
        *   Ghi toàn bộ kịch bản phân cảnh mốc thời gian (`[0:00 - 0:45] HOOK`, `[0:45 - 3:30] BODY`, `[7:00 - 8:30] CLIFFHANGER OUTRO`) trực tiếp vào cột G của **Sheet 3 (YouTube)**.
        *   Tạo file kịch bản `.md` chi tiết tương ứng tại `scratch/production_orders/DRAFT-YT-SERIES-PO-YYYY-WXX-XX.md` và dán hyperlink vào cột F (`Link Kịch bản Chi tiết`).
*   **Đầu ra (Output):** Bản thảo/Kịch bản được điền vào Sheet tương ứng (Sheet 2 hoặc Sheet 3) kèm link file `.md` với trạng thái mặc định là `Chờ thẩm định`.

### Bước 6: Thẩm định & Cập nhật Kết quả trên Excel (Agent 08)
*   **Đầu vào (Input):** Bản thảo/Kịch bản tại Sheet 2/3 và PO vật lý từ `Agent 04`.
*   **Hành động (Action):**
    *   *Đối chiếu chéo:* Tính toán lại các công thức số liệu; đối chiếu các số hiệu nghị định/thông tư hải quan; kiểm tra sự trùng khớp của cốt truyện thực tế từ [daily_notes.md](file:///d:/AI%20AGENT%20THUY/AI%20branding/scratch/daily_notes.md).
    *   *Compliance Audit:* Quét sạch các từ cấm, các tính từ cảm xúc kịch tính quá đà (*"đau đớn"*, *"giật mình"*) và đảm bảo tuân thủ định vị tuyến Trung Quốc/Hàn Quốc - Việt Nam.
    *   *GHI VÀO EXCEL:* `Agent 08` ghi trực tiếp kết quả thẩm định vào cột **Kết quả Fact-Check** (Cột H/I) và nội dung phản hồi cụ thể vào cột **Báo cáo Thẩm định SSoT**:
        *   `🟢 PASSED`: Đạt chuẩn hoàn chỉnh, 100% khớp với kết quả Web Search và PO, sẵn sàng trình chị Thúy duyệt.
        *   `🔴 REVISION REQUIRED`: REJECT bản thảo nếu phát hiện thông tin không có trong kết quả tìm kiếm/sai quy chuẩn; chỉ rõ sai lệch để Agent 05/06 tự chạy lại và sửa chữa trực tiếp trên sheet và file `.md`.
        *   `🔴 ESCALATED`: Chuyển tiếp Delta Report (dưới 150 từ) để trình chị Thúy nhận xét trực tiếp sau 2 vòng sửa đổi.
*   **Đầu ra (Output):** Cột Kết quả và Báo cáo Thẩm định trên Excel được cập nhật.

### Bước 7: Phân tích & Tối ưu hóa chu kỳ sau (Agent 07)
*   **Đầu vào (Input):** Số liệu tương tác thực tế từ các bài viết Facebook và video YouTube đã đăng.
*   **Hành động (Action):**
    *   Định kỳ phân tích và đưa ra đề xuất tối ưu khung giờ đăng, độ dài, nhóm từ khóa hoặc định dạng video hiệu quả.
    *   Bổ sung các "từ cấm động" nếu phát hiện thuật ngữ bị bóp reach.
*   **Đầu ra (Output):** Báo cáo tối ưu hóa hiệu suất gửi cho `Agent 04` & `Agent 08`.

---

## 📊 Cấu trúc Lưu trữ Tập trung của Tệp Weekly_Content_Calendar.xlsx (Chiến lược 2.0)

Tệp [Weekly_Content_Calendar.xlsx](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) được tinh gọn thành **3 Sheets duy nhất** tập trung 100% vào Facebook và YouTube:

### 1. Danh sách các Sheets:
*   **Sheet 1: Weekly Calendar:** Lịch tuần tổng 50:50 và duyệt ý tưởng concept A/B/C (Checkpoint 1).
*   **Sheet 2: Facebook:** Bài viết trang cá nhân/fanpage & kiểm duyệt chi tiết (Agent 05).
*   **Sheet 3: YouTube:** Kịch bản video dài/Series/Hướng dẫn thực chiến & kiểm duyệt chi tiết (Agent 06).

### 2. Cấu trúc Cột chuẩn hóa (Áp dụng cho Sheet 2 & Sheet 3):

Để đồng bộ lập trình và hiển thị trực quan, Sheet 2 (Facebook) và Sheet 3 (YouTube) áp dụng cấu trúc 11 cột từ **A đến K**:

| Cột | Tên Cột | Mô tả cho Sheet 2 (Facebook) | Mô tả cho Sheet 3 (YouTube) | Người thực hiện |
| :---: | :--- | :--- | :--- | :--- |
| **A** | **STT** | Số thứ tự bài đăng trong tuần. | Số thứ tự video trong tuần. | Tự động |
| **B** | **Mã PO** | Mã Phiếu Giao việc (Ví dụ: `PO-2026-W33-01`). | Mã Phiếu Giao việc (Ví dụ: `PO-2026-W33-02`). | `Agent 04` |
| **C** | **Ngày đăng/quay** | Ngày đăng dự kiến. | Ngày dự kiến quay / xuất bản video. | `Agent 04` |
| **D** | **Chủ đề / Tên Series**| Tiêu đề & Góc tranh luận chính. | Tên Series & Số tập video (`Tập 1`, `Tập 2`...). | `Agent 04 / 06` |
| **E** | **Từ khóa SEO & Hashtags**| Từ khóa SEO và hashtags bài viết. | Từ khóa SEO YouTube và Tags đi kèm. | `Agent 05 / 06` |
| **F** | **Link Bài viết / Kịch bản (.md)** | **[MỚI]** Hyperlink dẫn tới file `.md` bài viết chi tiết tại `scratch/production_orders/`. | **[MỚI]** Hyperlink dẫn tới file `.md` kịch bản chi tiết tại `scratch/production_orders/`. | `Agent 05 / 06` |
| **G** | **Nội dung / Kịch bản Chi tiết** | Bản thảo hoàn chỉnh của bài viết Facebook. | Kịch bản phân cảnh mốc thời gian chi tiết (`Hook`, `Body`, `Cliffhanger Outro`). | `Agent 05 / 06` |
| **H** | **Visual Brief / Screencast** | Chỉ dẫn thiết kế Infographic / Ảnh chụp mộc. | Mô tả bối cảnh quay, chuyển cảnh, Screencast ECUS, đồ họa. | `Agent 05 / 06` |
| **I** | **Kết quả Fact-Check** | Dropdown: `🟢 PASSED`, `🔴 REVISION REQUIRED`, `🔴 ESCALATED`. | Dropdown: `🟢 PASSED`, `🔴 REVISION REQUIRED`, `🔴 ESCALATED`. | `Agent 08` |
| **J** | **Báo cáo Thẩm định SSoT** | Ghi nhận chi tiết kết quả rà soát luật hải quan, từ cấm, số liệu. | Ghi nhận lỗi sai về điều luật hải quan, cốt truyện lịch sử, từ cấm. | `Agent 08` |
| **K** | **Trạng thái Xuất bản** | Trạng thái phê duyệt (`Sẵn sàng đăng`, `Đã duyệt quay`...). | Trạng thái phê duyệt (`Sẵn sàng dựng`, `Đã duyệt quay`...). | **Chị Thúy (User)** |

---

## 🚫 Bộ Quy tắc Chống Lỗi Phối hợp (Anti-Patterns)

1.  **Cấm sản xuất trước Checkpoint 1:** Không viết bài hoặc viết kịch bản khi chưa được chị Thúy duyệt ý tưởng tại Sheet 1.
2.  **Cấm dùng từ ngữ cảm xúc kịch tính thái quá:** Tuyệt đối không dùng các từ như *"đau đớn"*, *"giật mình"*, *"hoảng loạn"* trong bài viết hay kịch bản video.
3.  **Cấm đóng vai sai định vị tuyến đường:** Ví dụ thực tế về khách hàng chỉ lấy trên tuyến Trung Quốc - Việt Nam & Hàn Quốc - Việt Nam. Các xu hướng Châu Âu/Châu Mỹ chỉ trích dẫn dưới dạng thông tin thị trường khách quan.
4.  **Cấm bỏ qua Link File .md:** Mọi dòng trên Sheet 2 và Sheet 3 bắt buộc phải có hyperlink tại Cột F trỏ tới file nháp `.md` tương ứng để quản lý mã nguồn vật lý đồng bộ.
5.  **Cấm thiếu Cliffhanger Outro ở Kịch bản YouTube:** Kịch bản YouTube dài/Series bắt buộc phải có đoạn kết Cliffhanger ở Outro gợi mở kịch tính cho tập tiếp theo.
