# 🔄 HỆ THỐNG QUY TRÌNH PHỐI HỢP ĐIỀU PHỐI (PBAN WORKFLOWS 2.0)

Thư mục này chứa toàn bộ các kịch bản phối hợp liên tác tử (Multi-Agent Workflows) của Hệ thống Thương hiệu Cá nhân PBAN dành cho chị **Nguyễn Thị Ngọc Thúy**.

---

## 🧭 1. Ma trận Phân công & Điều phối (Orchestration Matrix)

Khi tiếp nhận yêu cầu từ chị Thúy, **Agent 00 (Chief Orchestrator)** sẽ tự động tra cứu bảng dưới đây để kích hoạt Workflow tương ứng:

| Mã Workflow | Tên Quy trình | Ý định / Câu lệnh kích hoạt mẫu | Tác tử tham gia | Đầu ra chính |
|:---:|---|---|:---:|---|
| **`wf01`** | **Lập lịch nội dung tuần** | *"Lên lịch tuần mới", "Lập content calendar tuần 34"* | `01`, `02`, `03`, `04`, `07` | File Excel Sheet 1 + Các file PO vật lý |
| **`wf02`** | **Sản xuất bài viết Facebook** | *"Viết bài Facebook theo PO-2026-W34-01", "Soạn bài FB"* | `05`, `08` | File `.md` bài viết + Cập nhật Sheet 2 |
| **`wf03`** | **Sản xuất kịch bản YouTube** | *"Làm kịch bản YouTube theo PO-2026-W34-02", "Soạn series YT"* | `06`, `08` | File `.md` kịch bản + Cập nhật Sheet 3 |
| **`wf04`** | **Nghiên cứu văn bản luật mới** | *"Nghiên cứu Nghị định 36/2026", "Tra cứu thông tư mới"* | `01`, `02`, `08` | Báo cáo SSoT + Cập nhật Excel Luật |
| **`wf05`** | **Quét tin nóng & Gợi ý chủ đề** | *"Quét tin nóng tuần này", "Đề xuất chủ đề tuần mới"* | `01` | Bản tin đề xuất 3-5 chủ đề chấm điểm |
| **`wf06`** | **Cập nhật kho ý tưởng 4 Pillars**| *"Cập nhật kho ý tưởng", "Quét ý tưởng Pillar Tin Ngành"* | `01` | Cập nhật 4 file Excel Ideas Banks |
| **`wf07`** | **Rà soát & Thẩm định nội dung** | *"Rà soát bài viết này", "Kiểm tra tính chính xác của bản thảo"*| `08`, `02` | Báo cáo Fact-Check `🟢 PASSED` / Cần sửa |
| **`wf08`** | **Đo lường & Cập nhật từ cấm** | *"Đánh giá hiệu suất tuần qua", "Cập nhật từ cấm động"* | `07`, `02` | Báo cáo tối ưu + Cập nhật từ cấm |
| **`wf09`** | **Sản xuất video ngắn Viral Reels** | *"/viral-reels", "/fb-video", "viết kịch bản hài văn phòng"* | `09`, `02`, `08` | File kịch bản + Cập nhật Excel OneDrive |

---

## 🎯 2. Hai Điểm Chặn Quyết Định Của Con Người (Human-in-the-Loop Checkpoints)

Tất cả các workflow sản xuất đều tuân thủ nghiêm ngặt 2 trạm kiểm soát của chị Thúy:
1. **Checkpoint 1 (Duyệt Ý tưởng & Lịch tuần):** Diễn ra tại cuối quy trình `wf01`. Chị Thúy mở file Excel [`Weekly_Content_Calendar.xlsx`](file:///d:/AI%20AGENT%20THUY/AI%20branding/reports/plan1/Weekly_Content_Calendar.xlsx) tại **Sheet 1**, chọn `Duyệt Option A/B/C` tại Cột L.
2. **Checkpoint 2 (Duyệt Xuất bản Nội dung):** Diễn ra tại cuối quy trình `wf02` và `wf03` sau khi **Agent 08** đã thẩm định `🟢 PASSED`. Chị Thúy kiểm tra bản thảo lần cuối và chọn trạng thái tại Cột K (Sheet 2/3).

---

## 📁 3. Danh mục Tệp Workflow Chi tiết
- [`wf01_weekly_planning.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf01_weekly_planning.md) — Quy trình Lập lịch tuần & Tự động hóa Pipeline
- [`wf02_facebook_production.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf02_facebook_production.md) — Quy trình Sản xuất bài viết Facebook
- [`wf03_youtube_production.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf03_youtube_production.md) — Quy trình Sản xuất kịch bản YouTube & Series
- [`wf04_legal_research.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf04_legal_research.md) — Quy trình Nghiên cứu nhanh & Phản biện Văn bản Pháp luật
- [`wf05_hot_news_scan.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf05_hot_news_scan.md) — Quy trình Quét tin nóng chủ động
- [`wf06_ideas_bank_update.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf06_ideas_bank_update.md) — Quy trình Cập nhật kho ý tưởng 4 Trụ cột
- [`wf07_content_audit_verification.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf07_content_audit_verification.md) — Quy trình Rà soát & Thẩm định tính chính xác/hợp lệ
- [`wf08_performance_feedback.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf08_performance_feedback.md) — Quy trình Đánh giá hiệu suất & Vòng lặp phản hồi
- [`wf09_viral_reels_production.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/workflows/wf09_viral_reels_production.md) — Quy trình Sản xuất video ngắn Viral Reels hài văn phòng cho Eureka
