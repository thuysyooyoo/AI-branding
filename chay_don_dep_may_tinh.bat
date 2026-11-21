@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ====================================================
echo   TIEN TRINH DON DEP MAY TINH HANG TUAN (CLEANUP)
echo ====================================================
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\weekly_cleanup.ps1"
echo.
echo ====================================================
echo   DA HOAN TAT DON DEP!
echo ====================================================
pause
