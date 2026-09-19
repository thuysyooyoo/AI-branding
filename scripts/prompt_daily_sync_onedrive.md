# PROMPT HƯỚNG DẪN AI TẠO TASK ĐỒNG BỘ REPORT SANG ONEDRIVE HÀNG NGÀY (11:00 AM)

---

## 1. Prompt copy trực tiếp ra lệnh cho AI Agent (Antigravity / Gemini / Custom Agent)

```text
[YÊU CẦU THIẾT LẬP TASK ĐỒNG BỘ FILE REPORT SANG ONEDRIVE HÀNG NGÀY LÚC 11:00 AM]

Hãy tạo một task định kỳ (Daily Recurring Task) chạy tự động vào lúc 11:00 AM hàng ngày với nhiệm vụ:

1. Thư mục nguồn (Source Folder):
   "D:\AI AGENT THUY\AI branding\reports"

2. Thư mục đích OneDrive Online (Target Local Sync Folder):
   "C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh"

3. Danh sách file Excel cần đồng bộ:
   - "Danh_sach_Van_ban_Phap_luat_XNK.xlsx"
   - "SEO_Keyword_Trends_Eureka.xlsx"

4. Quy trình xử lý:
   - Bước 1: Kiểm tra sự tồn tại của 2 file Excel nguồn trong thư mục D:\AI AGENT THUY\AI branding\reports.
   - Bước 2: Sao chép ghi đè (Overwrite) bản mới nhất của 2 file này sang thư mục đích OneDrive.
   - Bước 3: OneDrive Client trên máy tính sẽ tự động đẩy các file mới này lên OneDrive Cloud/Online.
   - Bước 4: Ghi log nhật ký đồng bộ vào "D:\AI AGENT THUY\AI branding\reports\sync_log.txt" ghi rõ thời gian và trạng thái (Thành công / Thất bại).

5. Cách thức thực thi:
   - Có thể kích hoạt chạy trực tiếp qua kịch bản PowerShell tại: "D:\AI AGENT THUY\AI branding\scripts\sync_reports_onedrive.ps1"
```

---

## 2. Cách đặt lịch bằng Lệnh `/schedule` trong Antigravity AI Agent

Nếu bạn muốn Antigravity AI Agent tự động lên lịch chạy tác vụ này hàng ngày:
1. Gõ `/schedule` trong ô chat.
2. Nhập biểu thức Cron: `0 11 * * *` (Mỗi ngày lúc 11:00 AM).
3. Nội dung Prompt cho Schedule:
   > "Chạy script PowerShell `D:\AI AGENT THUY\AI branding\scripts\sync_reports_onedrive.ps1` để đồng bộ 2 file Excel báo cáo sang OneDrive lúc 11h sáng hàng ngày."

---

## 3. Lệnh cài đặt tự động vào Windows Task Scheduler (Khuyên dùng - Chạy ngầm độc lập)

Để task chạy **100% tự động** kể cả khi ứng dụng AI Agent không mở, chạy lệnh PowerShell sau dưới quyền Admin:

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument '-ExecutionPolicy Bypass -File "D:\AI AGENT THUY\AI branding\scripts\sync_reports_onedrive.ps1"'
$trigger = New-ScheduledTaskTrigger -Daily -At 11:00AM
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -TaskName "DailySyncOneDriveReports_11AM" -Description "Đồng bộ file report Excel sang OneDrive lúc 11h sáng hàng ngày" -Force
```
