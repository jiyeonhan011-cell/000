﻿﻿﻿﻿$AppName    = "WMSInspect"
$InstallDir = "$env:LOCALAPPDATA\$AppName"
$GithubBase = "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web"

Write-Host "====================================="  -ForegroundColor Cyan
Write-Host "  창고이동 검수 시스템 설치"            -ForegroundColor Cyan
Write-Host "====================================="  -ForegroundColor Cyan
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
Invoke-WebRequest -Uri "$GithubBase/app.py"      -OutFile "$InstallDir\app.py"      -UseBasicParsing
Invoke-WebRequest -Uri "$GithubBase/launcher.py" -OutFile "$InstallDir\launcher.py" -UseBasicParsing
Invoke-WebRequest -Uri "$GithubBase/icon.ico"    -OutFile "$InstallDir\icon.ico"    -UseBasicParsing
Invoke-WebRequest -Uri "$GithubBase/.streamlit/config.toml" -OutFile "$InstallDir\.streamlit\config.toml" -UseBasicParsing
Write-Host "      완료" -ForegroundColor Green

# 패키지 설치
Write-Host "[4/5] 패키지 설치 중 (streamlit xlrd openpyxl pywebview)..." -ForegroundColor Yellow
& $pyExe -m pip install streamlit xlrd openpyxl pywebview -q
Write-Host "      완료" -ForegroundColor Green

# 업데이트 스크립트 생성 (PowerShell — curl보다 안정적, 실패 시 update.log에 기록)
$update = @'
$ErrorActionPreference = "Continue"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $dir
$base = "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web"
$log  = Join-Path $dir "update.log"
"[$(Get-Date)] 업데이트 시작" | Out-File $log
$files = @{
    "app.py"                  = "$base/app.py"
    "launcher.py"              = "$base/launcher.py"
    "icon.ico"                 = "$base/icon.ico"
    ".streamlit\config.toml"   = "$base/.streamlit/config.toml"
}
foreach ($f in $files.Keys) {
    try {
        Invoke-WebRequest -Uri $files[$f] -OutFile (Join-Path $dir $f) -UseBasicParsing -TimeoutSec 15
        "  OK: $f" | Out-File $log -Append
    } catch {
        "  FAIL: $f - $($_.Exception.Message)" | Out-File $log -Append
    }
}
"[$(Get-Date)] 업데이트 완료" | Out-File $log -Append
'@
$update | Out-File -Encoding UTF8 "$InstallDir\update.ps1"

# VBScript 런처 생성 (업데이트 → 실행)
$py  = $pyExe    -replace '\\', '\\'
$dir = $InstallDir -replace '\\', '\\'
$vbs = @"
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = "$dir"
sh.Run "powershell -ExecutionPolicy Bypass -File ""$dir\update.ps1""", 0, True
sh.Run "cmd /c ""$py"" launcher.py", 0
"@
$vbs | Out-File -Encoding Unicode "$InstallDir\run.vbs"

# 바탕화면 + 시작 메뉴 바로가기
$ws = New-Object -ComObject WScript.Shell

function Make-Shortcut($dest) {
    $sc = $ws.CreateShortcut($dest)
    $sc.TargetPath       = "wscript.exe"
    $sc.Arguments        = "`"$InstallDir\run.vbs`""
    $sc.WorkingDirectory = $InstallDir
    $sc.IconLocation     = "$InstallDir\icon.ico"
    $sc.Description      = "창고이동 검수 시스템"
    $sc.Save()
}

Make-Shortcut "$env:USERPROFILE\Desktop\창고이동 검수.lnk"
Make-Shortcut "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\창고이동 검수.lnk"

Write-Host "[5/5] 바로가기 생성 완료" -ForegroundColor Green
Write-Host ""
Write-Host "====================================="  -ForegroundColor Cyan
Write-Host "  설치 완료!"                           -ForegroundColor Green
Write-Host "====================================="  -ForegroundColor Cyan
Write-Host ""
Write-Host "바탕화면의 '창고이동 검수' 아이콘을 더블클릭하면" -ForegroundColor White
Write-Host "브라우저 없이 자체 창으로 실행됩니다."           -ForegroundColor White
Write-Host ""
Read-Host "엔터를 눌러 닫기"
