@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===================================================================
echo     DANG NHAP FACEBOOK CHO AI AGENT QUET GROUP (1-CLICK)
echo ===================================================================
python scripts\fb_auth.py
pause
