@echo off
cd /d "%~dp0"
title Candles_Amazon

echo ============================================
echo   Candles_Amazon - Scent Classifier
echo ============================================
echo.

python --version >nul 2>&1 && goto :found
py --version >nul 2>&1 && goto :found

echo [!] Python not found.
echo     Install Python 3.12 from https://www.python.org/downloads/
echo     *** Tick "Add Python to PATH" during install ***
pause
exit /b

:found
echo [OK] Python found

echo [1/2] Installing packages (Tsinghua mirror)...
python -m pip install streamlit openpyxl curl_cffi beautifulsoup4 playwright -q --disable-pip-version-check -i https://pypi.tuna.tsinghua.edu.cn/simple

echo [2/2] Starting...
echo     http://localhost:8501
echo     Ctrl+C to stop
echo.
start http://localhost:8501
python -m streamlit run app.py --server.headless true --browser.gatherUsageStats false

pause
