@echo off
chcp 65001 > nul
echo 창고이동 검수 프로그램 시작 중...

:: Python 설치 확인
python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo https://www.python.org 에서 Python 3.10 이상을 설치해주세요.
    pause
    exit /b
)

:: 패키지 설치
echo 필요한 패키지를 설치합니다...
pip install PyQt5 xlrd openpyxl -q

:: 프로그램 실행
echo 프로그램을 시작합니다...
python 검수프로그램.py

pause
