@echo off
chcp 65001 > nul

:: Try py launcher first, then python
py --version >nul 2>&1
IF NOT ERRORLEVEL 1 (
    SET PYTHON=py
    GOTO FOUND
)

python --version >nul 2>&1
IF NOT ERRORLEVEL 1 (
    SET PYTHON=python
    GOTO FOUND
)

python3 --version >nul 2>&1
IF NOT ERRORLEVEL 1 (
    SET PYTHON=python3
    GOTO FOUND
)

echo Python not found. Install from https://www.python.org/downloads/
echo Check "Add Python to PATH" during install.
pause
exit /b 1

:FOUND
%PYTHON% -m pip install streamlit xlrd openpyxl -q
streamlit run app.py --server.headless false --browser.gatherUsageStats false
IF ERRORLEVEL 1 (
    pause
)
