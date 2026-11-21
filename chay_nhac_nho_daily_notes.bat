@echo off
chcp 65001 >nul
cd /d "D:\AI AGENT THUY\AI branding"

if not exist logs md logs

echo ================================================================================ >> logs\daily_notes_reminder.log
echo [BAT DAU NHAC NHO DAILY NOTES]: %date% %time% >> logs\daily_notes_reminder.log
echo ================================================================================ >> logs\daily_notes_reminder.log

python -u scripts\daily_notes_reminder.py >> logs\daily_notes_reminder.log 2>&1

echo [HOAN TAT NHAC NHO DAILY NOTES]: %date% %time% >> logs\daily_notes_reminder.log
echo. >> logs\daily_notes_reminder.log
