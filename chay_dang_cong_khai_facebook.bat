@echo off
chcp 65001 >nul
title XÁC NHẬN ĐĂNG BÀI FACEBOOK CÔNG KHAI - EUREKA AI AGENT

echo =====================================================================
echo    HỆ THỐNG XÁC NHẬN CHUYỂN BÀI FACEBOOK SANG CÔNG KHAI (PUBLIC)
echo    Dự án: AI Personal Branding - Jasmine Nguyen / Eureka Logistics
echo =====================================================================
echo.

cd /d "d:\AI AGENT THUY\AI branding"

echo [*] Đang chuyển quyền riêng tư bài viết sang 'Công khai' (Public)...
echo [*] Đang tự động thả First Comment chứa link nguồn tài liệu...
echo.

python "scripts\fb_auto_poster.py" --publish

echo.
echo =====================================================================
echo    XUẤT BẢN THÀNH CÔNG! Đã cập nhật Lịch nội dung và báo Zalo.
echo =====================================================================
timeout /t 10
