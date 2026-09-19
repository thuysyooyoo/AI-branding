# BỘ QUY TẮC & TRÍ NHỚ VĨNH VIỄN (PROJECT MEMORY & RULES)
# DỰ ÁN: CÀO & PHÂN TÍCH BÀI ĐĂNG GROUP FACEBOOK VÀO ONEDRIVE EXCEL

> **LƯU Ý DÀNH CHO AI AGENT (ANTIGRAVITY):**
> Tài liệu này là kim chỉ nam tối cao cho mọi hành động quét, phân loại, ghi file và báo cáo dữ liệu Facebook Group của Marketing & Ban Giám Đốc Eureka. AI BẮT BUỘC tuân thủ nghiêm ngặt các quy tắc dưới đây trong MỌI phiên làm việc.

---

## 1. NGUYÊN TẮC PHÂN LOẠI BÀI VIẾT (TUYỆT ĐỐI KHÔNG NHẦM LẪN)
Mỗi bài viết hợp lệ chỉ được xếp vào **DUY NHẤT 1 TRONG 2 FILE EXCEL**, không bao giờ được xuất hiện ở cả hai:

### A. File 1: `01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx`
* **Tiêu chí:**
  - Người đăng nêu câu hỏi thực tế, case study vướng mắc nghiệp vụ (hải quan, thuế, C/O, mã HS, nợ thuế, kiểm tra chuyên ngành...).
  - Các bài nhờ tư vấn, hỏi kinh nghiệm xử lý tranh chấp, gỡ khoản nợ gán nhầm, hỏi quy trình nộp hồ sơ.
  - Ví dụ: Bị thuế truy thu hộ kinh doanh không mở, bị gán nhầm nghĩa vụ thuế đất phi nông nghiệp, hỏi cách gỡ hóa đơn sai đơn giá.

### B. File 2: `02_BAI_CHIA_SE_KIEN_THUC_GROUP_FACEBOOK.xlsx`
* **Tiêu chí:**
  - Bài viết mang tính chia sẻ công văn mới, thông tư, nghị định, chính sách cập nhật từ cơ quan chức năng.
  - Bài đúc kết bài học kinh nghiệm, cẩm nang nghiệp vụ, hướng dẫn chi tiết quy trình cho cộng đồng.
  - Ví dụ: Công văn tháo gỡ kiểm tra ATTP, hướng dẫn tỷ giá theo CV 6810/CT, cẩm nang tra mã HS bo mạch LED.

### C. Bài loại bỏ (Spam / Rác):
* Tuyệt đối loại bỏ: Tuyển dụng, tìm việc, chào giá cước/xe/cont, làm C/O thuê, dịch vụ hải quan, bán khóa học, mua bán via/tài khoản FB, bài chỉ để lại số điện thoại/Zalo.

---

## 2. QUY TẮC XÁC ĐỊNH TÁC GIẢ (CHỐNG NHẬM LẪN)
* **Tuyệt đối KHÔNG lấy tên tài khoản đăng nhập tool (ví dụ: Jasmine Nguyen)** làm tác giả bài viết.
* **Cách xác định tác giả chuẩn 100%:**
  1. Đọc tiêu đề popup/modal: `h2` có dạng `"Bài viết của [Tên tác giả]"`.
  2. Bóc tách tên tác giả từ thẻ `a[role="link"] strong` trong phần header của bài viết.
  3. Nếu bài đăng ẩn danh: Ghi đúng là `"Thành viên ẩn danh"`.

---

## 3. QUY TẮC KIỂM TRA HÌNH ẢNH / VIDEO (KHÔNG PHÁN ĐOÁN BỪA)
* **Phương tiện phải được xác thực trực tiếp trên DOM của bài viết:**
  - **Có ảnh:** Phải có thẻ `img[src*="fbcdn"]` với kích thước hiển thị thực tế `> 150px` bên trong vùng chứa bài viết hoặc có link ảnh. Ghi rõ mô tả: `Có (1 ảnh chụp màn hình eTax Mobile)`, `Có (ảnh chứng từ)`, `Có (2 ảnh)`.
  - **Có video:** Chỉ ghi khi thực sự có thẻ `<video>` hoặc player video đính kèm bài viết. KHÔNG BAO GIỜ nhận diện nhầm video reels/quảng cáo ngoài sidebar thành video của bài.
  - **Chỉ có chữ:** Ghi rõ ràng là `Không`.

---

## 4. QUY TẮC ĐẶT TIÊU ĐỀ & CHỦ ĐỀ TỪ KHÓA
* **Tiêu đề / tóm tắt:** Phải mang tính nghiệp vụ chuyên sâu, tóm lược bản chất vấn đề (ví dụ: *Kinh nghiệm gỡ bỏ khoản nợ thuế đất phi nông nghiệp bị gán nhầm trên eTax Mobile*). Tuyệt đối không cắt cụt ngủn hoặc chép y nguyên dòng đầu tiên của người đăng nếu câu đó quá sơ sài.
* **Chủ đề / từ khóa:** Phải gồm 3–5 cụm từ khóa nghiệp vụ chuẩn xác, cách nhau bằng dấu chấm phẩy (ví dụ: *Thuế đất phi nông nghiệp; Gán nhầm nghĩa vụ thuế; eTax Mobile; Trùng mã số thuế*).

---

## 5. QUY TẮC GHI FILE EXCEL & TRÁNH XUNG ĐỘT ONEDRIVE
* **Vị trí file:**
  `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\.13.Dữ liệu - bài viết - Hình ảnh - video\TỔNG HỢP DỮ LIỆU QUÉT GROUP FB`
* **Định dạng:** Giữ nguyên font `Arial 10pt`, STT số nguyên căn giữa, Ngày quét/đăng format `dd/mm/yyyy hh:mm`, số liệu format `#,##0`.
* **Tránh lỗi khóa file OneDrive (PermissionError [Errno 13]):**
  - Khi lưu file, BẮT BUỘC lưu ra file tạm `.tmp.xlsx` sau đó dùng lệnh đổi tên nguyên tử `os.replace` để tránh xung đột với tiến trình đồng bộ của OneDrive.

---

## 6. SCRIPT THỰC THI CHUẨN
* Mọi lần cào quét tự động phải sử dụng script: `d:\AI AGENT THUY\AI branding\scripts\scan_fb_groups_daily.py`.

---

## 7. TẦNG 4: BÁO CÁO ĐIỀU HÀNH QUA ZALO BOT (/onedrive-to-zalo)
* **Kích hoạt tự động & Theo yêu cầu:**
  - Ngay sau khi cào quét bài mới và ghi thành công vào OneDrive Excel, hệ thống sẽ tự động gửi một bản tin tóm tắt điều hành qua Zalo Bot (phòng chat: `"Cloud của tôi"`).
  - Hoặc khi người dùng gõ lệnh `/onedrive-to-zalo`, AI sẽ quét các file tổng hợp trên OneDrive và bắn ngay báo cáo phân tích vào Zalo của Sếp.
* **Cấu trúc bản tin Zalo chuẩn mực:**
  1. Header: 📊 **BẢN TIN ĐIỀU HÀNH: TỔNG HỢP DỮ LIỆU GROUP FACEBOOK 24H**
  2. Số lượng bài mới thu thập: Số bài Hỏi đáp + Số bài Chia sẻ kiến thức.
  3. Chi tiết từng case study / quy định nổi bật:
     - 📌 *Tiêu đề nghiệp vụ* | Nhóm | Tác giả
     - 🔍 *Vấn đề cốt lõi & Khuyến nghị*
     - 📎 *Phương tiện:* Ảnh eTax Mobile / Chứng từ / Không
     - 🔗 *Link bài viết gốc*
  4. Đích gửi: Zalo `"Cloud của tôi"` (qua Playwright session: `D:\AI AGENT THUY\ZALO REPORT\zalo_meal_bot_package\zalo_profile`).

---

## 8. TASK TỰ ĐỘNG 09:00 SÁNG: GIÁM SÁT KHO MARKETING ONEDRIVE (24H QUA)
* **Lịch trình:** Chạy tự động hàng ngày lúc **09:00:00 AM** qua Windows Task Scheduler (`Eureka_Marketing_Daily_9AM_Zalo`).
* **Phạm vi kiểm tra:** Toàn bộ thư mục `Phòng Marketing - Eureka - Marketing - Eureka` trong **24 giờ gần nhất**.
* **Mục tiêu giám sát:**
  - Kiểm tra xem các bạn Marketing (Khánh Linh, Nguyễn Thùy Trang, Thu Ánh, Hằng, Hương...) có cập nhật/nộp báo cáo hay tài liệu mới không.
  - Phân loại rõ ràng: Báo cáo & Số liệu, Kế hoạch & Chiến dịch, Thiết kế & Ấn phẩm, Video & Media.
  - Cảnh báo ngay nếu trong 24h qua phòng Marketing hoặc nhân sự nào không có bất kỳ cập nhật nào.
* **Kênh báo cáo:** Tự động gửi Bản tin điều hành 9h sáng vào Zalo `"Cloud của tôi"`.
* **Script thực thi:** `D:\AI AGENT THUY\ZALO REPORT\core\marketing_daily_monitor.py` (File kích hoạt: `D:\AI AGENT THUY\ZALO REPORT\chay_tu_dong_9h_marketing.bat`).

---

## 9. TASK TỰ ĐỘNG 08:30 SÁNG: CÀO BÀI GROUP FACEBOOK VÀO ONEDRIVE & BÁO CÁO ZALO (24H QUA)
* **Lịch trình:** Chạy tự động hàng ngày lúc **08:30:00 AM** qua Windows Task Scheduler (`Eureka_Facebook_Groups_Daily_Zalo`).
* **Phạm vi quét:** 6 Group Facebook chuyên ngành XNK, Logistics, Hải quan, Thuế trong **24 giờ gần nhất**.
* **Hành động tự động:**
  1. Quét và lọc bài viết công khai trong 24h qua, loại bỏ bài rác/spam.
  2. Bóc tách tác giả chuẩn 100%, kiểm tra trực tiếp ảnh/video đính kèm.
  3. Phân loại chuẩn vào `01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx` hoặc `02_BAI_CHIA_SE_KIEN_THUC_GROUP_FACEBOOK.xlsx`.
  4. Lưu vào Daily Note hàng ngày: Tự động đúc kết các case study thực tế, bài học nghiệp vụ vào `scratch/daily_notes.md` (chỉ lưu cục bộ để làm chất liệu cho Agent 01/03 sản xuất content branding / truyền thông, không cần đồng bộ lên OneDrive).
  5. Tự động kích hoạt Báo cáo Zalo: Soạn bản tin điều hành gửi vào Zalo `"Cloud của tôi"`.
* **Cơ chế dự phòng:** Có bật cờ `StartWhenAvailable`, nếu máy tính bật sau 08:30 sáng thì hệ thống sẽ tự động chạy bù ngay lập tức.
* **Script thực thi:** `D:\AI AGENT THUY\AI branding\scripts\scan_fb_groups_daily.py` (File kích hoạt: `D:\AI AGENT THUY\AI branding\chay_tu_dong_facebook.bat`).

---

## 10. QUY TẮC VĨNH VIỄN: TUYỆT ĐỐI KHÔNG TẠO BẤT KỲ FILE BACKUP NÀO (.XLSX)
* **Quy định tối cao:** Toàn bộ hệ thống và mọi AI Agent/Script **TUYỆT ĐỐI KHÔNG ĐƯỢC TẠO FILE BACKUP** (`*_backup_*.xlsx`, `_backup_w*.xlsx`, `Weekly_Content_Calendar_backup_*.xlsx`...) trong `reports/`, `reports/plan1/`, `scratch/dumps/` hay bất kỳ thư mục nào khác.
* **Lý do:** Tránh rác thư mục, nặng ổ đĩa và gây rối loạn cấu trúc tệp của người dùng.
* **Cơ chế cập nhật dữ liệu an toàn không cần backup:**
  - Cập nhật trực tiếp vào file chính thức.
  - Sử dụng kỹ thuật `atomic_save`: ghi workbook ra file tạm `.tmp.xlsx`, sau đó dùng lệnh đổi tên nguyên tử `os.replace` để ghi đè vào file chính nhằm tránh xung đột OneDrive sync, và đảm bảo không để lại bất kỳ file sao lưu nào.
* **Tự động dọn dẹp:** Bất kỳ quy trình nào nếu lỡ sinh ra file backup phải lập tức tự xóa bỏ ngay sau khi kết thúc.

