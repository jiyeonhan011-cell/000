@echo off
chcp 65001 > nul
echo 창고이동 검수 웹 프로그램 시작 중...

python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 Python 3.10 이상을 설치해주세요.
    pause & exit /b
)

echo 패키지 설치 중...
pip install flask xlrd openpyxl -q

echo.
echo ========================================
echo  브라우저에서 접속하세요:
echo  http://localhost:5000
echo ========================================
echo.
start http://localhost:5000
python app.py
pause
