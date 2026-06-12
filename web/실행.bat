@echo off
chcp 65001 > nul
title 창고이동 검수 프로그램

echo.
echo  ╔══════════════════════════════════════╗
echo  ║   창고이동 3단계 검수 프로그램         ║
echo  ╚══════════════════════════════════════╝
echo.

:: Python 확인
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo.
    echo  https://www.python.org/downloads/ 에서
    echo  Python 3.10 이상을 설치하세요.
    echo  (설치 시 "Add Python to PATH" 체크 필수!)
    echo.
    pause
    exit /b 1
)

echo 패키지 설치 중 (최초 1회만 시간이 걸립니다)...
pip install streamlit xlrd openpyxl -q
IF ERRORLEVEL 1 (
    echo [오류] 패키지 설치 실패. 인터넷 연결을 확인하세요.
    pause
    exit /b 1
)

echo.
echo  브라우저가 자동으로 열립니다.
echo  열리지 않으면: http://localhost:8501
echo.
echo  [종료하려면 이 창을 닫거나 Ctrl+C]
echo.

streamlit run app.py --server.headless false --browser.gatherUsageStats false
IF ERRORLEVEL 1 (
    echo.
    echo [오류] 실행 중 문제 발생. 위 오류 메시지를 확인하세요.
    pause
)
