@echo off
chcp 65001 >nul
echo ============================================
echo   Candles_Amazon 香型分类器 - 一键安装启动
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] 安装 Python 依赖...
pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo [错误] 依赖安装失败，请检查网络连接
    pause
    exit /b 1
)

echo [2/3] 安装 Playwright 浏览器（Chromium）...
python -m playwright install chromium
if %errorlevel% neq 0 (
    echo [警告] Playwright 浏览器安装失败，curl_cffi 优先模式仍可运行
)

echo [3/3] 启动应用...
echo.
echo 浏览器打开 http://localhost:8501
echo 按 Ctrl+C 停止
echo.
streamlit run app.py

pause
