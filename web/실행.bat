@echo off
chcp 65001 > nul

python --version >nul 2>&1
IF ERRORLEVEL 1 (
    echo Python not found. Install from https://www.python.org/downloads/
    echo Check "Add Python to PATH" during install.
    pause
    exit /b 1
)

pip install streamlit xlrd openpyxl -q

streamlit run app.py --server.headless false --browser.gatherUsageStats false
IF ERRORLEVEL 1 (
    pause
)
