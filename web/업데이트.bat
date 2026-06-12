@echo off
chcp 65001 > nul
echo Updating...

curl -L -o app.py "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/app.py"
curl -L -o 실행.bat "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/%EC%8B%A4%ED%96%89.bat"

echo.
echo Update complete. Run 실행.bat to start.
pause
