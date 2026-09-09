@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PYW=%LOCALAPPDATA%\Python\bin\pythonw.exe"
if not exist "%PYW%" set "PYW=%LOCALAPPDATA%\Python\pythoncore-3.14-64\pythonw.exe"
if exist "%PYW%" (
  start "" "%PYW%" "%~dp0오늘의_주간업무_알림.py" %*
  exit /b 0
)

where pythonw >nul 2>&1
if %errorlevel%==0 (
  start "" pythonw "%~dp0오늘의_주간업무_알림.py" %*
) else (
  start "" python "%~dp0오늘의_주간업무_알림.py" %*
)
