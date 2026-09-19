@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===================================================================
echo     CHAY CAO DU LIEU 6 GROUP FACEBOOK TU DONG VAO ONEDRIVE
echo ===================================================================
python scripts\scan_fb_groups_daily.py
pause
