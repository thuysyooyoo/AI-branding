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

---

## 11. TASK TỰ ĐỘNG 11:00 SÁNG THỨ HAI: DỌN DẸP MÁY TÍNH HÀNG TUẦN (WEEKLY CLEANUP)
* **Lịch trình:** Chạy tự động hàng tuần vào lúc **11:00:00 AM Thứ Hai** qua Windows Task Scheduler (`WeeklyCleanupTask`).
* **Mục tiêu:** Tối ưu hóa dung lượng ổ C & D, xóa rác hệ thống, dọn cache trình duyệt và các phiên làm việc Antigravity cũ.
* **Các hạng mục dọn dẹp tự động:**
  1. Kiểm tra tính toàn vẹn của Junction Roaming (`C:\Users\thuyn\AppData\Roaming` -> `D:\AppData\Roaming`).
  2. Dọn sạch thư mục tạm `%TEMP%`, `C:\Windows\Temp`, `Prefetch`.
  3. Xóa cache Windows Update Download (`SoftwareDistribution\Download`).
  4. Quét & xóa file cài đặt cũ trong thư mục Downloads (> 30 ngày: `.exe`, `.msi`, `.zip`...).
  5. Dọn dẹp bản cũ của file trùng lặp & file phiên bản (V1, V2, V3...) trong thư mục Downloads:
     - **Chỉ áp dụng trong thư mục Downloads** (không đụng đến thư mục làm việc hay OneDrive).
     - **Quy tắc giữ:** Giữ lại duy nhất **BẢN CUỐI CÙNG** (file có version cao nhất như `V4 > V3 > V2` hoặc bản tải về sau cùng theo `LastWriteTime`).
     - **Xóa:** Toàn bộ các bản thảo cũ hơn (`V1`, `V2`, các bản copy tải trước `(1)`, `(2)`).
     - **An toàn tuyệt đối:** Với các tên file chung chung (`document`, `image`, `scan`...), chỉ xóa nếu trùng 100% mã băm nội dung SHA-256 để tránh xóa nhầm chứng từ khác nhau.
  6. Dọn dẹp cache trình duyệt: Cốc Cốc, Google Chrome, Microsoft Edge.
  7. Làm rỗng Thùng rác (Recycle Bin) và xóa bộ nhớ đệm ảnh thu nhỏ (Thumbnail Cache).
  8. Dọn Antigravity Cache: Xóa các phiên chat và brain không hoạt động > 2 ngày để giải phóng dung lượng ổ C.
* **Cơ chế dự phòng:** Bật cờ `StartWhenAvailable`, nếu máy tính tắt vào sáng Thứ Hai thì khi khởi động lại sẽ tự chạy bù ngay.
* **Script thực thi:** `D:\AI AGENT THUY\AI branding\scripts\weekly_cleanup.ps1`
* **File kích hoạt thủ công:** `D:\AI AGENT THUY\AI branding\chay_don_dep_may_tinh.bat`
* **Nhật ký dọn dẹp:** Lưu tại `C:\Users\thuyn\CleanupLogs\cleanup_YYYY-MM.log`.

---

## 12. QUY ĐỊNH CHỐNG CÀO TRÙNG BÀI VIẾT (DEDUPLICATION PROTOCOL - 4 LỚP BẢO VỆ)
* **Mục tiêu:** Tuyệt đối không cào lại bài viết đã thu thập ở các ngày hôm trước hoặc các phiên trước đó, ngay cả khi bài đăng vẫn còn hiển thị trên bảng tin ("hôm qua", "20 giờ trước"), hoặc tác giả copy-paste đăng chéo nhiều nhóm.
* **4 Lớp bảo vệ chống trùng lặp bắt buộc:**
  1. **Lớp 1 - Chuẩn hóa & Trích xuất Facebook Post ID độc nhất:**
     - Facebook URL có nhiều định dạng: `/posts/<id>`, `/permalink/<id>`, `/share/p/<id>`, có/không có query tracking (`?__cft__`, `?ref`, `?mibextid`), bản mobile (`m.facebook.com`) vs bản máy tính (`www.facebook.com`).
     - BẮT BUỘC bóc tách `post_id` duy nhất từ URL. So khớp trực tiếp `post_id` trước khi click/truy cập bài viết. Nếu `post_id` đã có trong lịch sử thì bỏ qua ngay lập tức.
  2. **Lớp 2 - Sổ cái lịch sử bền vững cục bộ (`scratch/scraped_posts_ledger.json`):**
     - Mọi bài viết đã cào thành công phải được lưu lại vào sổ cái JSON cục bộ (`scratch/scraped_posts_ledger.json`) kèm `post_id`, `url`, `content_hash`, `title`, `author`, `group_name`, `scan_date`.
     - Đầu mỗi phiên quét, hệ thống tự động nạp toàn bộ Post ID và URL từ cả 3 nguồn:
       + Sổ cái cục bộ `scratch/scraped_posts_ledger.json`.
       + File OneDrive `01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx`.
       + File OneDrive `02_BAI_CHIA_SE_KIEN_THUC_GROUP_FACEBOOK.xlsx`.
     - Cơ chế này giúp chống trùng 100% kể cả khi file Excel trên OneDrive đang bị lock hoặc sync trễ.
  3. **Lớp 3 - Băm nội dung chống trùng bài chéo nhóm (Content Hash MD5):**
     - Chuẩn hóa nội dung bài viết (loại bỏ toàn bộ khoảng trắng thừa, đưa về chữ thường, bỏ ký tự đặc biệt) và tính toán mã băm MD5 cho 250 ký tự đặc trưng đầu tiên (`content_hash`).
     - Nếu một bài viết được tác giả đăng ở nhóm A hôm trước, hôm sau lại đăng ở nhóm B (khác URL/Post ID nhưng cùng nội dung) -> Hệ thống phát hiện trùng mã băm nội dung và tự động bỏ qua.
  4. **Lớp 4 - Lọc mốc thời gian 24h & Khử vùng giao thoa (Overlap Filter):**
     - Chỉ thu thập các bài viết có mốc thời gian thực sự trong vòng 24 giờ.
     - Các bài viết hiển thị "hôm qua" hoặc số giờ 20–24h bắt buộc phải đối chiếu qua Lớp 1 (Post ID) và Lớp 3 (Content Hash) trước khi phân tích để tránh cào lại bài của phiên quét 08:30 sáng hôm trước.

---

## 13. QUY CHUẨN BẢNG PHÂN CẢNH VIDEO TẢI LÊN GOOGLE FLOW (.CSV & .XLSX)
* **Quy tắc vĩnh viễn:** Mọi kịch bản phân cảnh video AI (khi dùng `/ai-video`, `/thuytalk-cm3` hoặc thiết kế video) **BẮT BUỘC TUÂN THỦ ĐÚNG 100% CẤU TRÚC 6 CỘT CHUẨN (TÍCH HỢP GÓC QUAY)**:
  1. `STT`: Số thứ tự nguyên (1, 2, 3...)
  2. `Tiêu đề`: Tên phân cảnh định dạng chuẩn (`Scene 01`, `Scene 02`...)
  3. `Mô tả phân cảnh (Prompt)`: Đoạn prompt tiếng Anh tinh gọn tập trung chuyển động thị giác và âm thanh: [Visual Animation & Camera Movement] + [Audio & SFX]. (LƯU Ý CỐT LÕI: TUYỆT ĐỐI KHÔNG ghi mô tả prompt lời thoại/Voiceover vào cột này vì toàn văn lời thoại đã được quản lý riêng biệt tại Cột 5: `Lời thoại (Dialogue)`).
  4. `Góc quay (Shot Type)`: Cỡ cảnh / góc máy điện ảnh chuẩn theo dropdown: `Trung cảnh (Medium Shot)`, `Toàn cảnh (Wide Shot)`, `Cận cảnh (Close-up)`.
  5. `Lời thoại (Dialogue)`: Lời thoại thuyết minh tiếng Việt nguyên văn đầy đủ.
  6. `Thời lượng (giây)`: Số giây của cảnh (ví dụ: 8 hoặc 10).
* **Quy cách xuất file bắt buộc:**
  - **File 1 (.csv):** `<Ten_Du_An>_Kich_Ban_Flow.csv` chuẩn UTF-8 có BOM, phân cách bằng dấu phẩy `,`, đúng 6 cột để người dùng tải thẳng lên Google Flow hoặc nạp vào workflow AI sinh video hàng loạt.
  - **File 2 (.xlsx):** `<Ten_Du_An>_Studio.xlsx` trong đó **Sheet 1 chính là `KICH_BAN_FLOW` chuẩn 6 cột nói trên**, kèm các sheet phụ trợ (Keyframes, Voiceover full, Profile phong cách).
* **Hiển thị trên Chat/Markdown:** Bảng phân cảnh trong tin nhắn chat hoặc tài liệu `STORYBOARD.md` phải trình bày đúng cấu trúc 6 cột này để người dùng đối chiếu 1:1.

---

## 14. QUY CHUẨN BẢNG PROMPT TẠO ẢNH HÀNG LOẠT (.CSV & .XLSX)
* **Quy tắc vĩnh viễn:** Mọi bảng prompt tạo ảnh hàng loạt (Keyframes, ảnh Start/End Frame, ảnh minh họa video) **BẮT BUỘC TUÂN THỦ ĐÚNG 100% CẤU TRÚC 5 CỘT** theo mẫu chuẩn (`C:\Users\thuyn\Downloads\Mau_Excel_Tao_Anh_Hang_Loat.csv`):
  1. `STT`: Số thứ tự nguyên (1, 2, 3...)
  2. `Tiêu đề`: Tên phân đoạn / chủ đề bức ảnh (`Khởi đầu`, `Gặp tường`, `Frame 01`...)
  3. `Prompt tạo ảnh (Image Prompt)`: Đoạn prompt Text-to-Image tiếng Anh chi tiết, giàu hình tượng bám sát cốt truyện:
     - **CẤM TUYỆT ĐỐI TÊN GÓC QUAY TRONG PROMPT:** Nghiêm cấm ghi các từ ngữ góc máy (`Toàn cảnh`, `Cận cảnh`, `Trung cảnh`, `Wide Shot`, `Medium Shot`, `Close-up`...) vào trong nội dung cột Prompt. Cỡ cảnh / góc máy chỉ được phép nằm DUY NHẤT ở Cột 4: `Góc quay (Shot Type)`.
     - **MÔ TẢ KHUNG CẢNH THỰC TẾ & SẮC NÉT (ZERO GENERIC):** Ảnh phải thực sự mô tả khung cảnh, nhân vật, hành động, đạo cụ cụ thể của câu chuyện. Tuyệt đối cấm các prompt chung chung lặp lại nhàm chán.
     - **TRỰC QUAN HÓA 100% SỐ LIỆU & DỮ KIỆN THỰC TẾ (DATA-INFUSED):** Nếu phân cảnh có số liệu, tiền tệ, phần trăm, số lượng, kích thước, thời gian (ví dụ: `$5,000`, `24.000 USD`, `27.000 USD`, `186 lá thư`, `14 máy bay`, `10 tỷ USD`, `193 km`, `40 container/giờ`, `90%`, `36 lần`...), bức ảnh BẮT BUỘC phải thể hiện các con số/dữ kiện đó một cách trực quan ngộ nghĩnh (trên bảng phấn doodle, biển neon, hóa đơn, két sắt, thước đo, đồng hồ đếm ngược, biểu ngữ, biểu đồ hình quạt...).
     - **ĐỒNG NHẤT MÀU SẮC TUYỆT ĐỐI (CONSISTENT PALETTE):** 100% cảnh vẽ nét mực đen vector trên nền trắng tinh khiết (`pure white background`), màu nhấn đồng nhất vàng rực (`vibrant yellow #FFD700`) và xanh cyan (`cyan blue #00B4D8`), điểm xuyết đỏ cảnh báo (`alert red`) khi nguy cấp/thâm hụt. Tuyệt đối cấm lệch sang các dải màu tím, cam, chàm làm vỡ tính thống nhất của video.
  4. `Góc quay (Shot Type)`: Cỡ cảnh / góc máy điện ảnh chuẩn theo dropdown: `Trung cảnh (Medium Shot)`, `Toàn cảnh (Wide Shot)`, `Cận cảnh (Close-up)`.
  5. `Ghi chú`: Mục đích, diễn biến hoặc trạng thái nhân vật / vai trò khung hình bằng tiếng Việt (`Nhân vật xuất hiện`, `Đối mặt thử thách`, `Start frame Scene 01`...).
* **Quy cách xuất file bắt buộc:**
  - **File CSV:** `<Ten_Du_An>_Tao_Anh_Hang_Loat.csv` chuẩn UTF-8 có BOM, phân cách bằng dấu phẩy `,`, đúng 5 cột để người dùng tải lên các công cụ tạo ảnh AI hàng loạt.
  - **File Excel:** `<Ten_Du_An>_Studio.xlsx` trong đó **Sheet 2 chính là `TAO_ANH_HANG_LOAT` đúng 5 cột nói trên**.
* **Hiển thị trên Chat/Markdown:** Bảng prompt tạo ảnh trong tin nhắn chat hoặc tài liệu phải trình bày đúng cấu trúc 5 cột này để người dùng đối chiếu 1:1.

---

## 15. TASK TỰ ĐỘNG 17:30 CHIỀU: NHẮC NHỞ CẬP NHẬT DAILY NOTES (CÔNG VIỆC & CUỘC SỐNG)
* **Lịch trình:** Chạy tự động hàng ngày lúc **17:30:00 PM** qua Windows Task Scheduler (`Eureka_Daily_Notes_Reminder_Zalo`).
* **Mục tiêu:** Nhắc Chị Thúy ghi lại các câu chuyện thực tế, bài học quản trị, đàm phán, sự cố hoặc chiêm nghiệm trong công việc & đời sống vào `scratch/daily_notes.md` để làm chất liệu gốc (SSoT) sống động cho hệ thống PBAN 2.0 (Agent 01, 03, 05, 06) sản xuất content thương hiệu cá nhân.
* **Cấu trúc 4 trường thông tin bắt buộc trong file:**
  1. `Tiêu đề`: Tên sự việc / tình huống trong ngày.
  2. `Sự kiện / Câu chuyện nổi bật trong ngày`: Diễn biến thực tế, bối cảnh, nhân vật, nút thắt (viết mộc mạc, không cần trau chuốt câu từ).
  3. `Bài học rút ra hoặc suy ngẫm cá nhân`: Góc nhìn Eureka, bài học quản trị, tư duy lãnh đạo kết hợp năng lượng Giáp Mộc (vững chãi, bao bọc, phụng sự) & Liêm Trinh (minh bạch, liêm chính, sắc bén).
  4. `Thông tin liên quan khác (nếu có)`: Hình ảnh chứng từ, tin nhắn đối chất, quyết định điều hành, lưu ý cho đội ngũ.
* **Kênh truyền thông:** Gửi tin nhắn qua Zalo Bot vào `"Cloud của tôi"` + hiển thị thông báo màn hình (Windows Alert).
* **Cơ chế tự động:** Nếu ngày hôm đó chưa có ghi chép, script tự động chèn khung mẫu `## Ngày dd/mm/yyyy` lên đầu `daily_notes.md`.
* **Script thực thi:** `D:\AI AGENT THUY\AI branding\scripts\daily_notes_reminder.py` (File kích hoạt: `D:\AI AGENT THUY\AI branding\chay_nhac_nho_daily_notes.bat`).

---

## 16. BỘ QUY CHUẨN ĐỒNG BỘ NGHỆ THUẬT & ĐO ĐẾM THOẠI VIDEO AI TOÀN NĂNG (UNIVERSAL AI VIDEO MASTER STANDARD)
* **Mục tiêu:** Đảm bảo 100% video AI sản xuất ra (cho Google Flow, Kling, Runway, Midjourney, Flux...) đạt tính đồng nhất tuyệt đối về thị giác, nhân vật, bối cảnh, nhịp điệu thời lượng và phong cách giọng đọc.

### A. Quy tắc Khóa Phong Cách Nghệ Thuật (Master Visual Anchor):
1. Trước khi viết prompt cho bất kỳ cảnh nào, AI BẮT BUỘC định hình và khóa chặt **Master Visual Anchor** (Bao gồm: Trường phái mỹ thuật, Chất liệu thể hiện, Phong cách ánh sáng, Nền bối cảnh và Bảng màu chủ đạo).
2. **Nguyên tắc "100% Cảnh Như Một":** Cụm Master Visual Anchor này phải xuất hiện đồng nhất trong toàn bộ prompt ảnh (`ảnh.xlsx`) và prompt video (`phân cảnh.xlsx`).
   - *Ví dụ Doodle 2D:* `Minimalist 2D doodle animation art style, clean vector line art, crisp bold black ink outlines, clean pure white background, selective vibrant accent colors, 8k resolution --ar 16:9`.
   - *Ví dụ 3D Pixar:* `3D stylized Pixar animation style, soft subsurface scattering, vibrant warm studio lighting, clean solid pastel background, smooth clay render, 8k resolution --ar 16:9`.
   - *Ví dụ Điện ảnh Chân thực:* `Cinematic realistic live-action, 35mm film grain, anamorphic lens flare, Kodak Portra 400 color grading, 8k resolution --ar 16:9`.
3. **Tuyệt đối cấm nhảy phong cách (Zero Visual Clash):** Nếu dự án chọn 2D Doodle nền trắng, cấm tuyệt đối việc tự ý xen kẽ cảnh 3D isometric, ảnh chụp macro sản phẩm, nền đen rọi đèn hay mặt bàn gỗ giữa chừng.

### B. Quy tắc Khóa DNA Nhân Vật (Master Character Anchor):
1. **Bộ gen mô tả bất biến (Fixed Character DNA):** Đối với bất kỳ nhân vật đại diện nào xuất hiện (người que, mascot, hay người thật), phải dùng chung 100% cụm từ mô tả ngoại hình đặc trưng cho mọi phân cảnh A-roll.
   - *Ví dụ Cô bé Thúy:* `Cute 2D female stick figure Thuy with signature high ponytail tied with vibrant yellow ribbon, minimalist black dress silhouette, expressive large doodle eyes with eyelashes`.
2. Không bao giờ thay đổi kiểu tóc, phụ kiện nhận diện hay trang phục giữa các cảnh trừ khi kịch bản có yêu cầu biến hình có chủ đích.

### C. Quy tắc Nhịp Cầu Start-Frame Continuity (Ảnh Tĩnh -> Video Flow):
1. Bức ảnh tạo ra từ `ảnh.xlsx` chính là **Start Frame (Khung hình bắt đầu)** tải lên Google Flow.
2. Mô tả chuyển động trong `phân cảnh.xlsx` bắt buộc phải xuất phát trực tiếp từ các chủ thể tĩnh, đạo cụ và bố cục đã được sinh ra trong ảnh Start Frame đó, sau đó phát triển hành động mượt mà bám theo lời đọc để triệt tiêu hiện tượng AI hallucination hoặc biến dạng hình ảnh.

### D. Quy tắc Chuẩn Hóa Nhãn Dropdown Góc Máy (`ảnh.xlsx`):
* Cột `Góc quay (Shot Type)` trong `ảnh.xlsx` BẮT BUỘC ghi đúng 100% một trong các nhãn hệ thống dropdown của web tool (không thêm tiền tố A-roll, B-roll hay mô tả phụ):
  - `Trung cảnh (Medium Shot)`
  - `Toàn cảnh (Wide Shot)`
  - `Cận cảnh (Close-up)`

### E. Quy tắc Đo Đếm Lời Thoại Theo Thời Lượng (Audio-to-Time Pacing):
1. **Chuẩn tốc độ nói tiếng Việt tự nhiên, diễn cảm:** **2.5 – 3.3 từ/giây** (tương đương 140 – 160 từ/phút).
2. **Bảng định mức số từ tối đa bắt buộc:**
   - **Cảnh 10 giây:** Tối đa **28 – 32 từ tiếng Việt**.
   - **Cảnh 8 giây:** Tối đa **22 – 26 từ tiếng Việt**.
   - **Cảnh 6 giây:** Tối đa **16 – 19 từ tiếng Việt**.
   - **Cảnh 5 giây:** Tối đa **12 – 15 từ tiếng Việt**.
3. **Quy tắc chia tách phân cảnh khi thoại dài:** Tuyệt đối KHÔNG nhồi nhét lời thoại dài vào 1 cảnh (ví dụ 86 từ ép nói trong 10s là sai phạm nghiêm trọng). Khi người dùng có nội dung dài mà không muốn cắt xén lời thoại, AI **BẮT BUỘC CHỦ ĐỘNG TÁCH THÀNH NHIỀU PHÂN CẢNH LIÊN TIẾP (Scene 01, Scene 02...)**, giữ nguyên vẹn 100% nội dung gốc của người dùng.

### F. Quy tắc Định Danh Giọng Đọc & Phong Cách Thoại (Vocal Identity & Persona):
1. **Một giọng đọc duy nhất (Single Vocal Persona):** Xác định rõ nhân vật thuyết minh ngay từ đầu (Ví dụ: *Nữ chuyên gia logistics năng lượng cao, dí dỏm, sắc sảo* hoặc *Nam chuyên gia phân tích điều tra trầm tĩnh*).
2. Với kịch bản phân cảnh Flow 6 cột (RULE 13), toàn văn lời thoại được tách riêng sang cột `Lời thoại (Dialogue)`. Cột `Mô tả phân cảnh (Prompt)` tuyệt đối KHÔNG lặp lại mô tả prompt Voiceover/lời thoại nữa để tránh thừa thãi khi nạp vào AI sinh video. Nhân vật giọng đọc được khai báo tổng thể ở Sheet `TONG_QUAN_PROJECT_STYLE` hoặc kịch bản chi tiết.

3. Đồng nhất ngôi xưng (ví dụ: "Thúy - bạn", "mình - các bạn") và văn phong xuyên suốt từ Scene đầu đến Scene cuối.

### G. Quy tắc Xuất File & Tích Hợp Server MCP `flow-storyboard-exporter`:
1. Mọi kịch bản phân cảnh video và tạo ảnh hàng loạt khi hoàn thành BẮT BUỘC xuất đồng thời:
   - File `.csv` UTF-8 có BOM đúng 5 cột tải thẳng lên Google Flow / Web tool tạo ảnh.
   - File `.xlsx` chuẩn 5 cột lưu trữ và kiểm soát.
2. Tự động sao lưu và đồng bộ vào thư mục MCP `C:\Users\thuyn\Documents\Flow_Storyboards`.
3. Tuân thủ tuyệt đối **RULE 10**: Không tạo bất kỳ file backup rác nào trong hệ thống.

### H. Quy Tắc Giới Hạn Trần Tối Đa Phân Cảnh (Hard Ceiling - Tối Đa 40 Cảnh / Video AI):
1. **Quy định trần cứng tối cao (Hard Ceiling):** **1 VIDEO AI TUYỆT ĐỐI KHÔNG QUÁ 40 PHÂN CẢNH** (Tương đương thời lượng tối đa **400 giây ~ 6 phút 40 giây** ở định mức 10s/cảnh).
2. **Quy chuẩn tách tập bắt buộc (Multi-Episode Series Protocol):**
   - Khi chủ đề hoặc kịch bản có nội dung dài (> 40 phân cảnh / > 400 giây), AI **BẮT BUỘC CHỦ ĐỘNG TÁCH THÀNH CÁC TẬP LIÊN HOÀN (Tập 1, Tập 2...)**, mỗi tập **tối đa đúng 40 phân cảnh**.
   - Mỗi tập phải là một dự án hoàn chỉnh, độc lập và có trọn bộ file riêng:
     + `<Ten_Du_An>_Tap_1_Studio.xlsx` (4 sheet quản trị)
     + `<Ten_Du_An>_Tap_1_Kich_Ban_Flow.csv` (Chuẩn 6 cột Google Flow)
     + `<Ten_Du_An>_Tap_1_Tao_Anh_Hang_Loat.csv` (Chuẩn 5 cột Start Frame)
     + `<Ten_Du_An>_Tap_1.md` (Toàn văn kịch bản nói & phân cảnh)
     *(Và tương tự cho Tập 2, Tập 3...)*
   - Đánh số phân cảnh của từng tập độc lập từ `Scene 01` đến tối đa `Scene 40` (STT 1 đến 40) để người dùng tải thẳng lên Google Flow hoặc các công cụ sinh video mà không bị lỗi tràn hàng hoặc quá tải thời lượng.
   - Cuối mỗi tập giữa (Tập 1, Tập 2...) bắt buộc có **Móc câu Cliffhanger** kịch tính để kéo khán giả sang tập tiếp theo; Tập cuối cùng chốt bằng **Outro tự dặn lòng mộc mạc** và lời chào kết.

---

## 17. QUY TRÌNH TỰ ĐỘNG ĐĂNG BÀI FACEBOOK (CHẾ ĐỘ CHỈ MÌNH TÔI -> KIỂM TRA -> CÔNG KHAI)
* **Mục tiêu:** Tự động lấy bài viết đã được kiểm duyệt từ Lịch nội dung Excel, đăng lên Facebook cá nhân ở chế độ `Chỉ mình tôi` (Only Me 🔒) để Chị Thúy kiểm tra trực quan giao diện bài viết trước, sau đó kích hoạt chuyển sang `Công khai` (Public 🌍) và tự động thả First Comment.
* **Quy trình 3 giai đoạn:**
  1. **Giai đoạn 1 - Đăng bài kiểm tra (Chỉ mình tôi):**
     - Đọc bài có trạng thái `Sẵn sàng đăng` từ `reports/plan1/Weekly_Content_Calendar.xlsx` (sheet `Facebook`).
     - Tự động mở Playwright với profile `fb_profile`, vào Facebook, chọn quyền riêng tư `Chỉ mình tôi`, nhập nội dung bài viết và bấm Đăng.
     - Lấy permalink bài viết, lưu vào `scratch/pending_facebook_post.json`.
     - Cập nhật cột Trạng thái trong Excel thành `Đang kiểm tra Chỉ mình tôi (dd/mm/yyyy hh:mm)` qua hàm `atomic_save` (không sinh file backup).
     - Gửi tin nhắn Zalo vào phòng `"Cloud của tôi"` kèm link bài viết để Chị Thúy mở trên app kiểm tra.
  2. **Giai đoạn 2 - Phê duyệt & Kiểm tra:**
     - Chị Thúy xem trực tiếp bài đăng trên Facebook (kiểm tra ngắt dòng, tiêu đề, icon, ảnh đính kèm).
  3. **Giai đoạn 3 - Xuất bản Công khai (Public):**
     - Chị Thúy bấm chạy `chay_dang_cong_khai_facebook.bat` (hoặc nhắn trong chat Antigravity "đăng công khai đi em").
     - Bot Playwright tự động truy cập bài viết, mở menu `...` -> `Chỉnh sửa đối tượng` -> Chọn `Công khai` -> Lưu.
     - Tự động thả First Comment chứa toàn bộ link nguồn tài liệu/văn bản pháp luật.
     - Cập nhật trạng thái Excel sang `Đã đăng Công khai (dd/mm/yyyy hh:mm)`.
     - Gửi tin nhắn Zalo thông báo hoàn tất vào `"Cloud của tôi"`.
* **File thực thi chuẩn:**
  - Script điều khiển: `scripts/fb_auto_poster.py`
  - Đăng bản nháp Chỉ mình tôi: `chay_dang_bai_chi_minh_toi.bat` (hoặc lệnh `python scripts/fb_auto_poster.py --post-draft`)
  - Chuyển sang Công khai: `chay_dang_cong_khai_facebook.bat` (hoặc lệnh `python scripts/fb_auto_poster.py --publish`)

---

## 18. QUY CHUẨN KỊCH BẢN PHÂN CẢNH & ẢNH YOUTUBE THUYTALK CHUYÊN MỤC 3 (/thuytalk-cm3)
* **Mục tiêu:** Tự động hóa sản xuất trọn gói kịch bản phân cảnh (Storyboard Google Flow) và bộ prompt tạo ảnh hàng loạt cho Kênh YouTube `ThuyTalk` của Chị Thúy (Chuyên mục 3: Ký sự Lịch sử Thương mại, Kỳ án Hàng hải & Logistics Quốc tế).
* **Kích hoạt lệnh:** Gõ `/thuytalk-cm3`, `/thuytalk-story` hoặc yêu cầu làm kịch bản phân cảnh & ảnh ThuyTalk chuyên mục 3.
* **Hệ thống lưu trữ đồng bộ (Storage Locations):**
  - **Vị trí Skill File:** `d:\AI AGENT THUY\AI branding\skills\thuytalk-cm3\SKILL.md` (Workspace) & `C:\Users\thuyn\.gemini\config\skills\thuytalk-cm3\SKILL.md` (Global Config).
  - **Vị trí lưu trữ đám mây (OneDrive):** `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\youtube\`
  - **Vị trí lưu trữ cục bộ:** `d:\AI AGENT THUY\AI branding\reports\ThuyTalk_CM3\`
  - **Vị trí đồng bộ MCP:** `C:\Users\thuyn\Documents\Flow_Storyboards\`
  - **File Master Series SSoT duy nhất:** `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx`
* **Bộ quy tắc bắt buộc:**
  1. **Nhập vai Chị Thúy 100%:** Xưng hô `Thúy / Mình` – `Các bạn / Mọi người`. 100% văn nói đời thường, mộc mạc, tự nhiên. Cấm viết tắt (`mng`, `dn`, `xnk`...), cấm tính từ cường điệu sáo rỗng (*"chấn động"*, *"nghìn cân treo sợi tóc"*...). Outro tự dặn lòng hoặc Cliffhanger kịch tính cho Series.
  2. **Bảng Phân cảnh 6 Cột Google Flow (RULE 13):** `STT`, `Tiêu đề`, `Mô tả phân cảnh (Prompt)`, `Góc quay (Shot Type)`, `Lời thoại (Dialogue)`, `Thời lượng (giây)` (chuẩn block 8s – 10s).
  3. **Bảng Prompt Tạo ảnh 5 Cột (RULE 14):** `STT`, `Tiêu đề`, `Prompt tạo ảnh (Image Prompt)`, `Góc quay (Shot Type)`, `Ghi chú`.
     - Tuyệt đối cấm chèn chữ góc máy (`Toàn cảnh`, `Cận cảnh`, `Trung cảnh`...) vào prompt.
     - Bắt buộc trực quan hóa số liệu nếu có trong câu chuyện (bảng phấn, hóa đơn, két sắt, biển hiệu neon, đồng hồ...).
     - Màu sắc đồng nhất 100% (nền trắng, nét đen, nhấn vàng `#FFD700` & cyan `#00B4D8`, cảnh báo đỏ).
     - Góc quay ở Cột 4 chỉ dùng 1 trong 3 nhãn dropdown: `Trung cảnh (Medium Shot)`, `Toàn cảnh (Wide Shot)`, `Cận cảnh (Close-up)`.
  4. **Kết nối 3 Trục hình ảnh (Không rời rạc):**
     - *Trục 1 - Host Thúy:* Master Character Anchor của Chị Thúy (nữ chuyên gia XNK hiện đại, thanh lịch, ấm áp, ánh mắt sắc bén, tương tác đạo cụ).
     - *Trục 2 - Góc nhìn Người nghe Thúy:* Các góc máy Over-the-shoulder, góc nhìn người nghe chăm chú, suy ngẫm, hoặc tương tác Eye-Level 1-1.
     - *Trục 3 - Bối cảnh Lịch sử:* Tàu buồm cổ, cảng biển xưa và nay, bản đồ hải trình, kho hàng, container... đồng bộ 100% Master Visual Anchor.
  5. **Đo đếm lời thoại (RULE 16E):** Tốc độ 2.5 – 3.3 từ/giây (Cảnh 10s tối đa 28–32 từ tiếng Việt). Tự động chia tách cảnh khi thoại dài.
  6. **Xuất file đồng thời (RULE 16G & RULE 10):** Tự động xuất `<Ten_Du_An>_Kich_Ban_Flow.csv`, `<Ten_Du_An>_Tao_Anh_Hang_Loat.csv`, `<Ten_Du_An>_Studio.xlsx` (4 sheet) qua script `scripts/export_thuytalk_storyboard.py`. Tuyệt đối không tạo file backup rác.





---

## 19. TASK TỰ ĐỘNG 09:00 SÁNG & 17:00 CHIỀU: TƯƠNG TÁC BẢNG TIN BẠN BÈ FACEBOOK CÁ NHÂN (50% THẢ TIM, 50% LIKE & BÁO CÁO ZALO)
* **Lịch trình:** Chạy tự động hàng ngày 2 lần vào lúc **09:00:00 AM** và **17:00:00 PM (5h chiều)** qua Windows Task Scheduler (`Eureka_Facebook_Personal_Engagement_Daily`).
* **Mục tiêu tương tác:**
  - Tăng tương tác bạn bè cá nhân tự nhiên (nurture / seeding quan hệ bạn bè của Chị Thúy).
  - Tương tác đúng **20 bài viết cá nhân** mỗi đợt.
  - Phân bổ cảm xúc: **50% Thả tim (❤️ Love)** và **50% Thích (👍 Like)** đan xen tự nhiên.
* **Bộ lọc loại trừ bắt buộc (Chống nhầm lẫn):**
  - **Loại bỏ 100% Group / Nhóm:** Phát hiện link `/groups/` trong bài viết -> Bỏ qua.
  - **Loại bỏ 100% Fanpage / Doanh nghiệp:** Phát hiện nút "Theo dõi" / "Follow" hoặc link `/pages/` -> Bỏ qua.
  - **Loại bỏ 100% Quảng cáo (Sponsored):** Nhận diện thẻ ads, "Được tài trợ", "Sponsored" -> Bỏ qua.
  - **Chỉ tương tác bạn bè cá nhân:** Chỉ nhận diện bài viết từ profile bạn bè (`/profile.php` hoặc `<username>`).
  - **TUYỆT ĐỐI KHÔNG LIKE COMMENT (BÌNH LUẬN):** Chỉ tương tác đúng duy nhất nút Thích trên thanh công cụ chính của bài viết (Post Action Bar cạnh nút Bình luận và Chia sẻ). Tuyệt đối loại bỏ 100% mọi nút Thích bên trong khu vực bình luận, câu trả lời hay danh sách comment bên dưới (`el.closest('ul, li, [aria-label*="bình luận"]')`).
  - **Chống lặp lại & Hủy like nhầm:** Tự động kiểm tra trạng thái nút tương tác, nếu bài đã thả tim/like trước đó thì bỏ qua ngay.
* **Giao thức an toàn (Chống checkpoint / Chặn tính năng Facebook):**
  - Cuộn chuột mượt mà (smooth scroll).
  - Dừng đọc ngẫu nhiên **4.5 – 7.5 giây** ở mỗi bài trước khi tương tác.
  - Nghỉ tự nhiên **3 – 5 giây** sau mỗi lần like/thả tim.
* **Kênh báo cáo kép (Zalo & Desktop Excel):**
  1. Tự động gửi Bản tin điều hành kèm danh sách 20 bài và trích dẫn vào Zalo **"Cloud của tôi"**.
  2. Tự động ghi nhận một dòng nhật ký thực thi chi tiết (thời gian, ca chạy, số tim, số like, tên bạn bè) vào Sheet `04_LOG_TUONG_TAC_FB` của file Excel Desktop: `C:\Users\thuyn\Desktop\TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx`.
  3. Tự động đồng bộ và làm mới Bảng điều hành HTML Desktop (`BANG_CHIDAN_TASK_VA_KETQUA_HANGNGAY.html`).
* **Script thực thi:** `D:\AI AGENT THUY\AI branding\scripts\fb_personal_interactor.py` (File chạy nhanh: `D:\AI AGENT THUY\AI branding\chay_tuong_tac_fb.bat`, Task Scheduler: `D:\AI AGENT THUY\AI branding\chay_tu_dong_fb_tuong_tac.bat`).

---

## 20. AGENT 09: TÁC TỬ BIÊN KỊCH & ĐẠO DIỄN VIDEO HÀI VĂN PHÒNG TRIỆU VIEW EUREKA (VIRAL REELS DIRECTOR)
* **Vị trí tệp Agent:** `d:\AI AGENT THUY\AI branding\agents\09_viral_reels_director.md`
* **Vị trí Skill nội bộ:** `d:\AI AGENT THUY\AI branding\skills\eureka-viral-reels\SKILL.md`
* **Quy trình kích hoạt:** `d:\AI AGENT THUY\AI branding\workflows\wf09_viral_reels_production.md`
* **Lệnh kích hoạt nhanh:** `/viral-reels`, `/fb-video`, `/hai-van-phong`
* **Nhiệm vụ tối cao:** Phụ trách sản xuất trọn bộ kịch bản video ngắn (Facebook Reels, TikTok, Shorts) hài văn phòng triệu view, bối cảnh 100% tại văn phòng Eureka 44 Triều Khúc, lồng ghép sắc sảo mô hình Canvas và thế mạnh cốt lõi của Eureka.
* **Bộ quy chuẩn bắt buộc:**
  1. **6 Hạng mục cốt lõi:** Deadpan mặt lạnh 100% trong tình huống phi lý, Thoại Ping-Pong dồn dập 2–5 từ/câu, Công thức 4 nhịp 30s (Hook 3s -> Escalation 12s -> Punchline 10s -> Deflation 5s), 100% văn phòng 44 Triều Khúc, Quảng cáo tự nhiên Native Branding, Đóng gói kỹ thuật (Top banner + Sub vàng `#FFE600` viền đen + SFX viral).
  2. **Luân phiên 6 Motif đa dạng:** Slapstick âm thanh/hành động, Drama nội bộ công sở (100% không có khách ngoài), Trào phúng tâm linh dân buôn, Nỗi sợ vợ thực tế của chủ hàng, Khách vốn nhỏ & Chơi chữ nghĩa đen, Nhại điệp viên Hollywood.
  3. **Chốt chặn Kiểm duyệt SSoT (`knowledge/banned_words.md`):** BẮT BUỘC rà soát 100% nội dung qua tệp từ cấm trước khi hoàn thiện kịch bản. Tuyệt đối cấm: trốn thuế, buôn lậu, hàng lậu, bao thuế, đi chui, tiểu ngạch, đánh hàng Quảng Châu, mua hàng trung tận gốc, dở à, đếch... Thay thế bằng ngôn từ an toàn, chính thống.
  4. **Tập trung 100% mặt hàng thế mạnh Canvas:** Máy móc công nghiệp, dây chuyền cơ khí, linh kiện phụ tùng, thiết bị điện, gia dụng thông minh, hàng mẫu hỏa tốc. Tuyệt đối không đưa mặt hàng cấm/ngoài thế mạnh (như đồ chơi trẻ em) vào kịch bản.
  5. **Ngôn từ Sếp Thúy:** Chuẩn phong thái Nữ CEO B2B sắc bén, văn minh, đĩnh đạc, châm biếm duyên dáng chuẩn quyền uy lãnh đạo.
  6. **Quy tắc Kịch bản Mới (Chống Trùng & Tự Động Xóa Kịch Bản Đã Sản Xuất):**
     - **Không trùng/tương tự:** Kịch bản mới tuyệt đối không trùng hoặc tương tự về tình huống drama, ngòi nổ hay cú lật so với các kịch bản cũ.
     - **Tự động xóa kịch bản đã sản xuất:** Nếu kịch bản nào được note **"Đã sản xuất"** (hoặc đã quay xong khi đến lịch) -> Hệ thống tự động xóa khỏi danh sách chờ quay để giữ file luôn tinh gọn, cập nhật.
     - **Chỉ tiêu & Tham khảo MBBank:** Thực hiện đúng theo [`HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md`](file:///d:/AI%20AGENT%20THUY/AI%20branding/HUONG_DAN_LAP_KICH_BAN_HANG_NGAY.md) và `wf09`. Mỗi ngày viết đúng 3 kịch bản mới. Tham khảo kênh MBBank Hải Dương qua script `scripts/scan_mbhaiduong_daily.py` (ngưỡng **>= 30,000 views**, bóc băng audio 100%).
  7. **Quy chuẩn Phân Luồng Đúng 2 File Excel Trên OneDrive (9 Cột, atomic_save):**
     - **Luồng 1 (Đào kịch bản của MB):** Cập nhật vào `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_KICH_BAN_EUREKA_CHUAN_THOAI_MB.xlsx.tmp.xlsx` *(hoặc `.xlsx`)* (lưu toàn bộ thoại bóc băng audio 100% và kịch bản đối ứng).
     - **Luồng 2 (Tạo kịch bản mới Eureka):** Cập nhật vào `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\video content\BANG_10_KICH_BAN_VIRAL_EUREKA_DA_DANG_6_MOTIF.xlsx` (Sheet 1 lưu 3 kịch bản mới/ngày; Sheet 4 lưu ý tưởng tham khảo). Tuyệt đối không tạo file backup rác theo RULE 10.
  8. **Quy tắc Báo cáo Kết quả vào Bảng Chỉ Dẫn Desktop:**
     - Sau khi hoàn thành đào kịch bản MB hoặc tạo kịch bản mới Eureka, Agent BẮT BUỘC cập nhật / đồng bộ kết quả vào file điều hành của Giám Đốc tại `C:\Users\thuyn\Desktop\TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx` (bằng cách kích hoạt script `python D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py`).

---

## 21. QUY TẮC ĐỒNG BỘ BÁO CÁO VÀO BẢNG CHỈ DẪN KẾT QUẢ DESKTOP (MASTER TASK INDEX)
* **Vị trí file:** `C:\Users\thuyn\Desktop\TONG_HOP_CHI_DAN_TASK_VA_FILE_KET_QUA_HANG_NGAY.xlsx` & `C:\Users\thuyn\Desktop\BANG_CHIDAN_TASK_VA_KETQUA_HANGNGAY.html`
* **Script kích hoạt:** `python "D:\AI AGENT THUY\QUAN_TRI_TASK_HANG_NGAY\generator_chi_dan.py"` (File chạy 1-click Desktop: `C:\Users\thuyn\Desktop\CAP_NHAT_NGAY_GIO_KET_QUA.bat`)
* **Nguyên tắc bất biến:** Mọi tác vụ tự động hoặc thủ công khi hoàn thành (cào Group FB, bóc băng MBBank Reels, viết kịch bản viral mới, cập nhật Daily Notes...) BẮT BUỘC phải gọi script `generator_chi_dan.py` để quét lại ngày giờ thực tế và cập nhật trạng thái thời gian thực vào Bảng Điều Hành Desktop cho Giám Đốc Thúy.

---

## 22. QUY ĐỊNH BẤT BIẾN: QUY CHUẨN MASTER SERIES YOUTUBE THUYTALK & CƠ CHẾ XOAY VÒNG HÀNG TUẦN, ĐÀO THẢI TỒN ĐỌNG 2 TUẦN (WEEKLY ROLLING PIPELINE & 2-WEEK STALE PRUNING PROTOCOL)
* **Nguyên tắc cốt lõi về phát triển hàng tuần & đào thải tồn đọng (Không nhân đôi máy móc):**
  - **Mỗi tuần phát triển thêm 9 chủ đề mới mỗi CM (+36 chủ đề/tuần):** Ở mỗi chu kỳ tuần sản xuất (W45, W46, W47...), hệ thống tự động nghiên cứu, bóc tách từ 4 nguồn tri thức sống (kho kịch bản YouTube `Weekly_Content_Calendar.xlsx`, Cổng thông tin Hải quan, nhóm Facebook `01_BAI_HOI_DAP_GROUP_FACEBOOK.xlsx` & `02_BAI_CHIA_SE_KIEN_THUC_GROUP_FACEBOOK.xlsx`, nhật ký `daily_notes.md`) và **bổ sung thêm đúng 9 chủ đề mới cho mỗi chuyên mục** (CM1: 9 mã, CM2: 9 mã, CM3: 9 mã, CM4: 9 mã) vào `Ke_Hoach_Phat_Song_YouTube_Series.xlsx`.
  - **Cơ chế Đào thải Tự động Chủ đề Tồn đọng (2-Week Stale Pruning):** Trong mỗi phiên quét tuần, hệ thống tự động rà soát toàn bộ các chủ đề đã được đề xuất từ **2 tuần trước đó** (tức là đã nằm trong bảng qua 2 chu kỳ tuần xem xét, >= 14 ngày):
    + **XÓA BỎ HOÀN TOÀN** các chủ đề mà sau 2 tuần **KHÔNG ĐƯỢC CHỊ THÚY BÁO DUYỆT VIẾT KỊCH BẢN CỤ THỂ** (vẫn ở trạng thái 'Chờ Chị Thúy duyệt' / chưa được chuyển sang trạng thái duyệt viết kịch bản hoặc chưa có link kịch bản chi tiết).
    + **Lý do:** Giữ cho bảng kế hoạch luôn tinh gọn, mang tính thời sự cao, đào thải triệt để các ý tưởng cũ không trúng định hướng hoặc đã nguội.
  - **Cơ chế Bảo lưu Vĩnh viễn (Approval Retention):**
    + Mọi chủ đề đã được Chị Thúy báo duyệt viết kịch bản cụ thể (trạng thái `🟢 ĐÃ HOÀN THÀNH KỊCH BẢN`, `🟢 ĐÃ DUYỆT VIẾT KỊCH BẢN`...) ➔ **ĐƯỢC BẢO LƯU VĨNH VIỄN 100%**, tuyệt đối không bao giờ bị xóa.
    + Các chủ đề mới đưa vào trong vòng 2 tuần gần nhất (vẫn còn trong cửa sổ xem xét 2 tuần) ➔ Tiếp tục được giữ lại để Chị Thúy cân nhắc duyệt.
  - **Quản lý Tuần Đề Xuất (Week Tracking):** Trên file Excel ghi rõ cột `Tuần Đề Xuất` (ví dụ: `W45/2026`, `W46/2026`...) để theo dõi chính xác tuổi thọ của từng chủ đề.
  - **TIÊU CHUẨN THỜI GIAN KHẮT KHE ĐỐI VỚI ĐỀ TÀI CM1 (TRONG VÒNG 1 THÁNG TRỞ LẠI ĐÂY - ZERO STALE LAWS RULE):**
    + Chuyên mục 1 (`CM1_VAN_BAN_NONG`) là 'Trà chiều & Văn bản nóng', bản chất BẮT BUỘC 100% các đề tài đề xuất mới phải lấy từ các thông tư, công văn hỏa tốc, nghị định, quyết định hoặc dự thảo ban hành/có hiệu lực **TRONG VÒNG 1 THÁNG (30 NGÀY) TRỞ LẠI ĐÂY** tính từ thời điểm tuần hiện tại.
    + **TUYỆT ĐỐI CẤM LẤY VĂN BẢN CŨ CỦA NHỮNG NĂM TRƯỚC (2025, 2024...)** đưa vào CM1 làm tin nóng.
    + Nếu là văn bản nền tảng giải thích bản chất ngộ nhận, hướng dẫn quy trình làm đúng thì bắt buộc phải đưa vào **CM2: Dịch luật ra tiếng người**, tuyệt đối cấm đưa vào CM1.
* **Nguồn SSoT & Quy trình phối hợp:**
  - **File Master Series SSoT duy nhất:** `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx`.
  - Hệ thống tự động điền đầy đủ metadata cho mọi dòng mới: Mã chủ đề, Tiêu đề SEO, Định dạng & Số tập, Thời lượng, Nguồn SSoT Fact-Check, Hook mở đầu, Cốt lõi kịch bản & Luận điểm chi tiết, Chỉ dẫn Visual/Góc máy, Móc câu Cliffhanger, Kết luận tự dặn lòng của Thúy.
  - Khi Chị Thúy xem xét ghi chú `"đã duyệt"` (hoặc các biến thể duyệt: `duyệt quay`, `OK`, `Thúy duyệt`...) hoặc khi kích hoạt sản xuất theo batch, hệ thống tự động sinh kịch bản chi tiết (.md) chuẩn Persona Thúy, thẩm định Fact-Check qua Agent 08, lưu vào `BRANDING/youtube/` và gắn công thức Hyperlink 1-click vào cột 13.
* **Quy chuẩn sản xuất bắt buộc (Theo đúng tiêu chuẩn của Chị Thúy):**
  1. **Persona Chị Thúy:** 100% văn nói đời thường mộc mạc, bình dị, tự nhiên, mở miệng là nói được ngay, phong thái người làm nghề thực chiến 15 năm (Founder & CEO Eureka Logistics).
  2. **Thời lượng:** 8 – 10 phút (~1.400 – 1.600 từ), phân đoạn rõ ràng: Hook mở đầu (0:00 - 1:15), Thân bài phân tích thực chiến & tình huống sâu sắc (1:15 - 6:45), Outro tự dặn lòng và móc câu Cliffhanger dẫn sang tập tiếp theo (6:45 - 8:30).
  3. **Không từ viết tắt & Không giật gân (Agent 08 Audit):** Tuyệt đối cấm teen code, cấm từ viết tắt (`xnk`, `mst`, `dn`, `sp`, `kh`, `hq`, `b/l`, `mng`...), cấm từ giật gân sáo rỗng (*chấn động, kinh hoàng, rúng động, lằn ranh sinh tử...*).
  4. **Cơ chế đọc/ghi file an toàn (Bypass Lock & Atomic Save):** Bắt buộc sử dụng Windows API Shared Read (`FILE_SHARE_READ`) và kỹ thuật `atomic_save` (ghi ra file tạm cùng thư mục rồi `os.replace`) để luôn đọc và ghi an toàn ngay cả khi Chị Thúy đang mở file trong Excel/SharePoint Web.
  5. **Gắn link Hyperlink 1-click:** Cập nhật công thức `=HYPERLINK("<path_onedrive_md>", "👉 Xem Kịch Bản Chi Tiết")` tại Cột 13 của dòng tương ứng.
* **Script thực thi chuẩn:** `scripts/auto_generate_youtube_weekly.py`.

---

## 23. QUY CHUẨN KỊCH BẢN YOUTUBE THUYTALK CM1, CM2, CM4 (5-STAGE RETENTION FORMULA & FULL SPOKEN SCRIPT)
* **Mục tiêu tối cao:** Đảm bảo mọi video dài (5 – 9 phút) thuộc Chuyên mục 1 (Văn bản nóng), Chuyên mục 2 (Dịch luật ra tiếng người), và Chuyên mục 4 (Tâm tình thực chiến) đạt tỷ lệ giữ chân khán giả (Audience Retention) tối đa, người nghe nghe từ đầu đến cuối không bấm thoát.
* **Kích hoạt lệnh:** Gõ `/thuytalk-cm124`, `/thuytalk-script`, `/thuytalk-spoken`, `/thuytalk-talk`.
* **Hệ thống lưu trữ đồng bộ (Storage Locations):**
  - **Vị trí Skill File:** `d:\AI AGENT THUY\AI branding\skills\thuytalk-cm124\SKILL.md` (Workspace) & `C:\Users\thuyn\.gemini\config\skills\thuytalk-cm124\SKILL.md` (Global Config).
  - **Vị trí lưu trữ đám mây (OneDrive):** `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\youtube\` (Đồng bộ 100% cùng thư mục với CM3).
  - **Vị trí lưu trữ cục bộ:** 
    + CM1: `d:\AI AGENT THUY\AI branding\reports\ThuyTalk_CM1\`
    + CM2: `d:\AI AGENT THUY\AI branding\reports\ThuyTalk_CM2\`
    + CM4: `d:\AI AGENT THUY\AI branding\reports\ThuyTalk_CM4\`
  - **File Master Series SSoT duy nhất:** `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\BRANDING\Ke_Hoach_Phat_Song_YouTube_Series.xlsx`
* **5 Nguyên tắc sản xuất bắt buộc vĩnh viễn:**
  1. **Phong cách trò chuyện & tâm tình đồng nhất:**
     - Xưng hô chuẩn: `Thúy / Mình` – `Các bạn / Mọi người` (hoặc `các anh chị`).
     - Giọng văn mộc mạc, bình dị, tự nhiên như hai người bạn ngồi uống nước ngoài đời thực, nói chuyện chân thành, mở miệng là nói được ngay, không gượng gạo.
  2. **Phân tích bắt buộc có ví dụ cụ thể để liên tưởng (Zero Khoe Khoang):**
     - Tuyệt đối KHÔNG nói theo kiểu "mình có kinh nghiệm xử lý", "đội ngũ bên Thúy từng hỗ trợ xử lý" hay khoe khoang thành tích doanh nghiệp.
     - BẮT BUỘC dùng các cấu trúc dẫn chuyện giả định, gợi mở hoặc case study khách quan từ cộng đồng:
       * *"Thử tưởng tượng một ngày đẹp trời..."*
       * *"Ví dụ như một doanh nghiệp đang nhập khẩu..."*
       * *"Giả sử trường hợp các bạn nhập một lô hàng..."*
       * *"Nếu như ngày mai doanh nghiệp của các bạn nhận được thông báo..."*
       * *"Hãy xem case study thực tế này: Mới đây trên một hội nhóm xuất nhập khẩu, có một bạn đăng bài hỏi cầu cứu về việc..."* (nếu là case study thật cào từ các group Facebook).
     - Giúp người nghe "mắt thấy tai nghe" và liên tưởng ngay đến công việc thực tế của họ mà không cảm thấy người nói đang lên mặt dạy đời hay tự khen mình.
  3. **Tách riêng 1 phần kịch bản nói (Full Spoken Script):**
     - Trong file Markdown kịch bản, phần lời thoại phải được tách riêng thành một khối độc lập hoàn chỉnh: `## 🎙️ PHẦN 2: KỊCH BẢN THUYẾT MINH THUẦN VĂN NÓI (FULL SPOKEN SCRIPT - DÀNH RIÊNG ĐỂ THU ÂM & PHÂN TÍCH THOẠI)`.
     - Phần này không chứa bảng biểu hay ghi chú kỹ thuật rườm rà, giúp Chị Thúy tập trung 100% vào việc đọc, phân tích và thu âm trơn tru.
  4. **Cấu trúc 5 phần giữ chân khán giả (5-Stage Retention Formula):**
     - **Hồi 1: Hook mở đầu (0:00 - 0:45):** Đặt vấn đề giật mình, con số thiệt hại thực tế ➔ *[Câu dẫn chuyển tiếp 1]*
     - **Hồi 2: Phân tích thực trạng & Ví dụ cụ thể (0:45 - 3:00):** Bóc tách bản chất + Case study thực tế ➔ *[Câu dẫn chuyển tiếp 2]*
     - **Hồi 3: Giải pháp thực chiến Step-by-Step (3:00 - 5:30):** 3 bước tháo gỡ, quy trình làm đúng ➔ *[Câu dẫn chuyển tiếp 3]*
     - **Hồi 4: Liên tưởng mở rộng & Tác động thị trường (5:30 - 7:00):** Tác động dòng tiền, luồng xanh/đỏ, xu hướng siết chặt 1–2 năm tới ➔ *[Câu dẫn chuyển tiếp 4]*
     - **Hồi 5: Kết thúc mở & Outro tự dặn lòng (7:00 - 8:30):** Chiêm nghiệm tự dặn lòng mộc mạc (không dạy đời) + Câu hỏi gợi mở tương tác.
  5. **Móc câu chuyển tiếp bắt buộc (Transitional Bridges) - Biến hóa theo đề tài:**
     - **Ý nghĩa chữ "Bắt buộc":** Bắt buộc về mặt KỸ THUẬT CẤU TRÚC (Script Architecture) để giữ nhịp kịch bản, kích thích tò mò, không được nhảy bổ sang ý mới mà không có cầu nối.
     - **Tuyệt đối CẤM rập khuôn / văn mẫu (Zero Cookie-Cutter):** Nghiêm cấm copy-paste một câu dẫn y chang nhau từ kịch bản này sang kịch bản khác. Câu từ phải được "đo ni đóng giày" riêng cho từng chủ đề, bám sát từ khóa và nỗi đau cụ thể (CM1 chính sách gắt gao, CM2 kỹ thuật nghiệp vụ ách tắc cảng, CM4 tâm sự chuyện nghề và áp lực).
     - **4 Vị trí câu dẫn biến hóa linh hoạt:**
       * *Cuối Hook:* Kích thích tò mò về bản chất đằng sau quy định / chủ đề đó.
       * *Cuối Phân tích:* Đặt câu hỏi tháo gỡ rủi ro, bảo vệ quyền lợi doanh nghiệp.
       * *Cuối Giải pháp:* Nâng tầm nhìn rộng ra dòng tiền, chuỗi cung ứng và xu hướng 1-2 năm tới.
       * *Cuối Liên tưởng:* Lắng đọng vào tâm thế, đạo lý và bài học của người làm nghề.
      - **Ví dụ tương phản chống hiểu nhầm cho mọi phiên làm việc mới:**
        * ❌ **SAI (Rập khuôn máy móc - CẤM):** Video nào cuối Hook cũng ghi: "Nhiều người nghĩ đây chỉ là một sự cố giấy tờ nhỏ nhặt... Thúy sẽ chỉ rõ cho các bạn thấy ngay sau đây."
        * ✅ **ĐÚNG (Đo ni đóng giày theo từng đề tài - BẮT BUỘC):**
          - *CM1 (Hóa đơn ma):* "Nhiều người vẫn nghĩ quy định về siết hóa đơn điện tử chỉ là câu chuyện hành chính xa vời. Nhưng tại sao rủi ro từ công ty 'ma' lại có thể khiến doanh nghiệp mất trắng chi phí được trừ? Thúy sẽ bóc tách ngay sau đây:"
          - *CM2 (Sau thông quan):* "Nhiều người cứ nghĩ hàng về kho là xong xuôi. Nhưng tại sao kiểm tra sau thông quan 5 năm lại là 'cơn ác mộng' tiền tỷ? Những điểm tử huyệt nào bị soi kỹ nhất?"
          - *CM4 (1 điểm chạm):* "Người ngoài nhìn vào logistics thường chỉ thấy tàu to hàng nhiều. Nhưng những đêm mất ngủ đằng sau sự hỗn loạn 'đá bóng trách nhiệm' giữa các phòng ban là gì?"


