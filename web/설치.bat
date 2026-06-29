@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if %errorlevel% neq 0 (
    echo.
    echo 오류가 발생했습니다. 위 메시지를 캡처해서 알려주세요.
    pause
)
