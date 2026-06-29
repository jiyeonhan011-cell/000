@echo off
powershell -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Error: Please capture the message above.
    pause
)
