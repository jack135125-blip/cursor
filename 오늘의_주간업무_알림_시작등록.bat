@echo off
chcp 65001 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0오늘의_주간업무_알림_시작등록.ps1"
if errorlevel 1 (
  echo 등록에 실패했습니다.
  pause
  exit /b 1
)
echo.
pause
