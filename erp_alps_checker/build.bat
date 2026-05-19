@echo off
title ERP-Alps 빌드

echo.
echo ======================================
echo  ERP-Alps 검증 프로그램 빌드 시작
echo ======================================
echo.

:: Python 확인
python --version
if errorlevel 1 (
    echo.
    echo [오류] Python이 없습니다.
    echo python.org/downloads 에서 설치 후 다시 실행하세요.
    echo 설치 시 "Add Python to PATH" 체크 필수!
    echo.
    pause
    exit /b 1
)

echo.
echo [1/3] 패키지 설치 중...
pip install PyQt5 pandas openpyxl xlsxwriter pyinstaller
if errorlevel 1 (
    echo.
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

echo.
echo [2/3] exe 빌드 중... (3~5분 소요)
echo.
pyinstaller --onefile --windowed --name ERP_Alps_Check --add-data "logic.py;." gui.py
if errorlevel 1 (
    echo.
    echo [오류] 빌드 실패
    pause
    exit /b 1
)

echo.
echo ======================================
echo  완료! dist\ERP_Alps_Check.exe 실행
echo ======================================
echo.
pause
