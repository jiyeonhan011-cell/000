@echo off
cd /d "%~dp0"

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
