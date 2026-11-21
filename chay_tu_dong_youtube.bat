@echo off
chcp 65001 >nul
echo ========================================================
echo  KÍCH HOẠT QUY TRÌNH TỰ ĐỘNG SẢN XUẤT YOUTUBE THUYTALK (WF03)
echo  (Cơ chế: Event-driven / Automation Trigger / On-demand)
echo ========================================================
cd /d "d:\AI AGENT THUY\AI branding"
python "scripts\auto_generate_youtube_weekly.py" %*
echo.
echo ========================================================
echo  Hoàn tất thực thi kịch bản YouTube ThuyTalk.
echo ========================================================
pause
