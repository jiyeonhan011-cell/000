@echo off
cd /d "%~dp0"

echo Downloading latest version...
curl -L -o app.py "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/app.py"
if not exist .streamlit mkdir .streamlit
curl -L -o .streamlit\config.toml "https://raw.githubusercontent.com/jiyeonhan011-cell/000/main/web/.streamlit/config.toml"
echo Update complete.

py --version >nul 2>&1
IF NOT ERRORLEVEL 1 (
    py -m pip install streamlit xlrd openpyxl -q
    py -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
    pause
    exit /b
)

python --version >nul 2>&1
IF NOT ERRORLEVEL 1 (
    python -m pip install streamlit xlrd openpyxl -q
    python -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
    pause
    exit /b
)

echo Python not found. Install from https://www.python.org/downloads/
echo Check "Add Python to PATH" during install.
pause
