@echo off
chcp 65001 > nul
title 창고이동 검수 웹 서버

echo.
echo  ╔══════════════════════════════════════╗
echo  ║   창고이동 3단계 검수 웹 프로그램      ║
echo  ╚══════════════════════════════════════╝
echo.

:: Python 설치 확인
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo  https://www.python.org/downloads/ 에서
    echo  Python 3.10 이상을 설치한 후 다시 실행하세요.
    echo  (설치 시 "Add Python to PATH" 체크 필수)
    echo.
    pause
    exit /b 1
)

echo [1/2] 필요한 패키지 설치 중...
pip install flask xlrd openpyxl -q
IF ERRORLEVEL 1 (
    echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
    pause
    exit /b 1
)

echo [2/2] 서버 시작 중...
echo.
echo  브라우저가 자동으로 열립니다.
echo  열리지 않으면 주소창에 http://localhost:5000 입력
echo.
echo  [종료하려면 이 창을 닫거나 Ctrl+C 누르세요]
echo.

python app.py
IF ERRORLEVEL 1 (
    echo.
    echo [오류] 서버 실행 중 문제가 발생했습니다.
    pause
)
