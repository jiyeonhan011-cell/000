@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0설치.ps1"
if %errorlevel% neq 0 (
    echo.
    echo 오류가 발생했습니다. 위 메시지를 캘포쳐서 알려주세요.
    pause
)
