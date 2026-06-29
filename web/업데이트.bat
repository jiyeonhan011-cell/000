@echo off
cd /d "%~dp0"
echo Downloading latest files...
curl -L -o app.py "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/app.py"
curl -L -o 실행.bat "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/%EC%8B%A4%ED%96%89.bat"
if not exist .streamlit mkdir .streamlit
curl -L -o .streamlit\config.toml "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/.streamlit/config.toml"
echo Done. Press any key to close.
pause
