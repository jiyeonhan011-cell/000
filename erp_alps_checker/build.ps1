Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "  ERP-Alps 검증 프로그램 빌드 시작" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

# Python 확인
try {
    $ver = python --version 2>&1
    Write-Host "[OK] $ver 감지됨" -ForegroundColor Green
} catch {
    Write-Host "[오류] Python이 설치되지 않았습니다." -ForegroundColor Red
    Write-Host "https://www.python.org/downloads/ 에서 설치 후 다시 실행하세요." -ForegroundColor Yellow
    Write-Host "설치 시 'Add Python to PATH' 반드시 체크!" -ForegroundColor Yellow
    Read-Host "엔터를 누르면 종료"
    exit 1
}

# 패키지 설치
Write-Host ""
Write-Host "[1/3] 패키지 설치 중..." -ForegroundColor Yellow
pip install PyQt5 pandas openpyxl xlsxwriter pyinstaller
if ($LASTEXITCODE -ne 0) {
    Write-Host "[오류] 패키지 설치 실패" -ForegroundColor Red
    Read-Host "엔터를 누르면 종료"
    exit 1
}

# PyInstaller 빌드
Write-Host ""
Write-Host "[2/3] exe 빌드 중... (3~5분 소요)" -ForegroundColor Yellow
Write-Host ""
pyinstaller --onefile --windowed --name ERP_Alps_Check --add-data "logic.py;." gui.py
if ($LASTEXITCODE -ne 0) {
    Write-Host "[오류] 빌드 실패" -ForegroundColor Red
    Read-Host "엔터를 누르면 종료"
    exit 1
}

Write-Host ""
Write-Host "======================================" -ForegroundColor Green
Write-Host "  완료! dist\ERP_Alps_Check.exe 실행" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Green
Write-Host ""
Read-Host "엔터를 누르면 종료"
