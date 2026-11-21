@echo off
chcp 65001 >nul
cd /d "D:\AI AGENT THUY\AI branding"

if not exist logs md logs

echo ================================================================================ >> logs\wf06_ideas_bank_update.log
echo [BAT DAU CAP NHAT KHO Y TUONG WF06]: %date% %time% >> logs\wf06_ideas_bank_update.log
echo ================================================================================ >> logs\wf06_ideas_bank_update.log

python -u scripts\update_ideas_bank_w45_d4.py >> logs\wf06_ideas_bank_update.log 2>&1

echo [HOAN TAT CAP NHAT KHO Y TUONG WF06]: %date% %time% >> logs\wf06_ideas_bank_update.log
echo. >> logs\wf06_ideas_bank_update.log
