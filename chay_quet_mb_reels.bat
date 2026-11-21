@echo off
chcp 65001 > nul
echo =====================================================================
echo    EUREKA LOGISTICS - QUET CAP NHAT REELS MBBANK HAI DUONG HANG NGAY
echo =====================================================================
cd /d "d:\AI AGENT THUY\AI branding"
python scripts\scan_mbhaiduong_daily.py
echo [*] Hoan tat quet va cap nhat y tuong reels MBBank!
pause
