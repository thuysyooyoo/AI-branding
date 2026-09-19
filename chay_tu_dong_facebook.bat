@echo off
chcp 65001 >nul
cd /d "D:\AI AGENT THUY\AI branding"

if not exist logs md logs

echo === BAT DAU CAO GROUP FACEBOOK 24H: %date% %time% === >> logs\facebook_daily_scan.log
python -u scripts\scan_fb_groups_daily.py >> logs\facebook_daily_scan.log 2>&1
echo === HOAN TAT CAO GROUP FACEBOOK: %date% %time% === >> logs\facebook_daily_scan.log
