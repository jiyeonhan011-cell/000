@echo off
cd /d "%~dp0"
echo Downloading latest app.py...
curl -L -o app.py "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/app.py"
echo Done. Press any key to close.
pause
