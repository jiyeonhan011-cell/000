@echo off
python --version >nul 2>&1
if errorlevel 1 (
    echo Python이 필요합니다. python.org 에서 설치 후 다시 실행하세요.
    echo 설치시 Add Python to PATH 체크 필수!
    pause
    exit /b
)
pip install PyQt5 pandas openpyxl xlsxwriter -q
python "%~dp0gui.py"
