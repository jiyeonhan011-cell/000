# 창고이동 검수 설치 스크립트
$ErrorActionPreference = "Stop"
$AppName    = "창고이동검수"
$InstallDir = "$env:LOCALAPPDATA\$AppName"
$GithubBase = "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web"

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  창고이동 검수 시스템 설치" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Python 확인
$pyExe = $null
try   { $pyExe = (Get-Command py     -ErrorAction Stop).Source } catch {}
if (-not $pyExe) {
    try { $pyExe = (Get-Command python -ErrorAction Stop).Source } catch {}
}
if (-not $pyExe) {
    Write-Host "[오류] Python이 설치되어 있지 않습니다." -ForegroundColor Red
    Write-Host "https://www.python.org/downloads/ 에서 설치 후 다시 실행해주세요." -ForegroundColor Yellow
    Write-Host "'Add Python to PATH' 체크 필수!" -ForegroundColor Yellow
    Read-Host "`n엔터를 눌러 닫기"
    exit 1
}
Write-Host "[1/5] Python 확인됨: $pyExe" -ForegroundColor Green

# 설치 폴더 생성
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
New-Item -ItemType Directory -Force -Path "$InstallDir\.streamlit" | Out-Null
foreach ($sub in "이동처리","라벨발행","작업내역","선작업") {
    New-Item -ItemType Directory -Force -Path "$InstallDir\검수파일\$sub" | Out-Null
}
Write-Host "[2/5] 설치 폴더 생성: $InstallDir" -ForegroundColor Green

# 파일 다운로드
Write-Host "[3/5] 최신 파일 다운로드 중..." -ForegroundColor Yellow
Invoke-WebRequest -Uri "$GithubBase/app.py" -OutFile "$InstallDir\app.py" -UseBasicParsing
Invoke-WebRequest -Uri "$GithubBase/.streamlit/config.toml" -OutFile "$InstallDir\.streamlit\config.toml" -UseBasicParsing
Write-Host "      완료" -ForegroundColor Green

# 패키지 설치
Write-Host "[4/5] 패키지 설치 중 (streamlit xlrd openpyxl)..." -ForegroundColor Yellow
& $pyExe -m pip install streamlit xlrd openpyxl -q
Write-Host "      완료" -ForegroundColor Green

# 자동업데이트 + 실행 VBScript 생성 (콘솔 창 없이 실행)
$launchVbs = @"
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "$($InstallDir -replace '\\','\\\\')"

' 최신 버전 자동 다운로드 (숨김)
WshShell.Run "cmd /c curl -L -o app.py \"$GithubBase/app.py\" 2>nul", 0, True
WshShell.Run "cmd /c curl -L -o .streamlit\\config.toml \"$GithubBase/.streamlit/config.toml\" 2>nul", 0, True

' Streamlit 실행
WshShell.Run "cmd /c \"$($pyExe -replace '\\','\\\\')\" -m streamlit run app.py --server.headless false --browser.gatherUsageStats false", 0
"@
$launchVbs | Out-File -Encoding ASCII "$InstallDir\실행.vbs"

# 바탕화면 + 시작 메뉴 바로가기
$WshShell = New-Object -ComObject WScript.Shell

function New-Shortcut($path) {
    $sc = $WshShell.CreateShortcut($path)
    $sc.TargetPath       = "wscript.exe"
    $sc.Arguments        = "`"$InstallDir\실행.vbs`""
    $sc.WorkingDirectory = $InstallDir
    $sc.IconLocation     = "shell32.dll,20"
    $sc.Description      = "창고이동 검수 시스템"
    $sc.Save()
}

New-Shortcut "$env:USERPROFILE\Desktop\창고이동 검수.lnk"
$startDir = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
New-Shortcut "$startDir\창고이동 검수.lnk"

Write-Host "[5/5] 바로가기 생성 완료" -ForegroundColor Green
Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  설치 완료!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "바탕화면의 '창고이동 검수' 아이콘을 더블클릭하면 실행됩니다." -ForegroundColor White
Write-Host "실행 시마다 자동으로 최신 버전으로 업데이트됩니다." -ForegroundColor White
Write-Host ""
Read-Host "엔터를 눌러 닫기"
