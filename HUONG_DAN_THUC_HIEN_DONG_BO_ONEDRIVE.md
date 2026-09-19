# CHỈ DẪN AGENT: ĐỒNG BỘ REPORT SANG ONEDRIVE (11:00 AM HÀNG NGÀY)

Hàng ngày vào lúc **11:00 AM**, Agent thực hiện nhiệm vụ đồng bộ 2 file báo cáo Excel theo quy trình sau:

## 1. Thông tin đồng bộ
* **Thư mục nguồn:** `D:\AI AGENT THUY\AI branding\reports`
* **Thư mục đích (OneDrive):** `C:\Users\thuyn\OneDrive - EUREKA TRADE AND IMPORT EXPORT COMPANY LIMITED\Phòng Marketing - Eureka - Marketing - Eureka\thuanh`
* **Các file đồng bộ:**
  - `Danh_sach_Van_ban_Phap_luat_XNK.xlsx`
  - `SEO_Keyword_Trends_Eureka.xlsx`

## 2. Hành động của Agent
Agent chạy kịch bản PowerShell bên dưới để copy đè file mới nhất sang OneDrive:

```powershell
powershell -ExecutionPolicy Bypass -File "D:\AI AGENT THUY\AI branding\scripts\sync_reports_onedrive.ps1"
```

## 3. Kiểm tra nhật ký
Sau khi chạy, Agent kiểm tra file nhật ký để xác nhận kết quả: `D:\AI AGENT THUY\AI branding\reports\sync_log.txt`
