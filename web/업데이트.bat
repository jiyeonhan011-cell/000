@echo off
cd /d "%~dp0"
echo Downloading latest app.py...
curl -L -o app.py "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/app.py"
if not exist .streamlit mkdir .streamlit
curl -L -o .streamlit\config.toml "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/.streamlit/config.toml"
echo Done. Press any key to close.
pause
