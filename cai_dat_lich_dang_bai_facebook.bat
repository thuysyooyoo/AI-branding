@echo off
chcp 65001 >nul
title CÀI ĐẶT TASK TỰ ĐỘNG ĐĂNG BÀI CHỈ MÌNH TÔI 08:00 SÁNG - EUREKA AI AGENT

echo =====================================================================
echo    CÀI ĐẶT LỊCH TỰ ĐỘNG ĐĂNG BÀI FACEBOOK (CHẾ ĐỘ CHỈ MÌNH TÔI)
echo    Thời gian: 08:00 Sáng hàng ngày qua Windows Task Scheduler
echo =====================================================================
echo.

schtasks /create /tn "Eureka_Facebook_Draft_Post_Daily" /tr "\"d:\AI AGENT THUY\AI branding\chay_dang_bai_chi_minh_toi.bat\"" /sc daily /st 08:00 /f

if %ERRORLEVEL% equ 0 (
    echo.
    echo [✓] ĐÃ CÀI ĐẶT THÀNH CÔNG TASK "Eureka_Facebook_Draft_Post_Daily"!
    echo     Từ ngày mai, đúng 08:00 sáng hệ thống sẽ tự động:
    echo     1. Lấy bài viết sẵn sàng từ Excel.
    echo     2. Đăng lên Facebook cá nhân ở chế độ "Chỉ mình tôi".
    echo     3. Gửi link bài viết vào Zalo "Cloud của tôi" để Chị Thúy kiểm tra!
) else (
    echo.
    echo [!] Cài đặt không thành công. Hãy bấm chuột phải chọn 'Run as administrator'!
)

echo.
pause
