@echo off
chcp 65001 >nul
echo.
echo ========================================
echo  ERP-Alps 검증 프로그램 빌드 시작
echo ========================================
echo.

:: Python 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되지 않았습니다.
    echo https://www.python.org/downloads/ 에서 Python 3.10 이상을 설치하세요.
    pause
    exit /b 1
)

echo [1/3] 필수 패키지 설치 중...
pip install -r requirements.txt -q
if errorlevel 1 (
    echo [오류] 패키지 설치 실패
    pause
    exit /b 1
)

echo [2/3] 실행 파일 빌드 중... (수 분 소요)
pyinstaller --onefile ^
    --windowed ^
    --name "ERP_Alps_검증" ^
    --add-data "logic.py;." ^
    gui.py

if errorlevel 1 (
    echo [오류] 빌드 실패
    pause
    exit /b 1
)

echo [3/3] 완료!
echo.
echo ========================================
echo  dist\ERP_Alps_검증.exe 생성 완료!
echo  해당 파일을 실행하세요.
echo ========================================
echo.
pause
