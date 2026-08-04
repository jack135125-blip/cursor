@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] 의존성 설치 중...
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Python을 찾을 수 없습니다. Python 3.10+ 설치 후 PATH에 추가해 주세요.
  pause
  exit /b 1
)

echo.
echo [2/2] 웹 서버 시작: http://127.0.0.1:5000
echo 브라우저에서 위 주소로 접속하세요. 종료하려면 Ctrl+C
echo.
python app.py
pause
