@echo off
chcp 65001 >nul
title TỰ ĐỘNG ĐĂNG BÀI FACEBOOK (CHẾ ĐỘ CHỈ MÌNH TÔI) - EUREKA AI AGENT

echo =====================================================================
echo    HỆ THỐNG TỰ ĐỘNG ĐĂNG BÀI FACEBOOK DẠNG "CHỈ MÌNH TÔI" ĐỂ DUYỆT
echo    Dự án: AI Personal Branding - Jasmine Nguyen / Eureka Logistics
echo =====================================================================
echo.

cd /d "d:\AI AGENT THUY\AI branding"

echo [*] Đang lấy bài viết sẵn sàng từ Lịch nội dung Excel...
echo [*] Đang đăng bài ở chế độ 'Chỉ mình tôi' và gửi link vào Zalo...
echo.

python "scripts\fb_auto_poster.py" --post-draft

echo.
echo =====================================================================
echo    HOÀN TẤT! Chị Thúy mở Zalo "Cloud của tôi" để xem link bài viết nhé!
echo =====================================================================
timeout /t 10
