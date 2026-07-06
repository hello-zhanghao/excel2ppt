@echo off
cd /d "%~dp0"
title Excel Tool

echo ========================================
echo   Excel Analysis Tool
echo ========================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found
    echo Please install Python: https://www.python.org/downloads/
    goto :end
)

python --version
echo.

python -c "import tkinter" 2>nul
if errorlevel 1 (
    echo [ERROR] tkinter not available
    echo Please reinstall Python and check "tcl/tk" option
    goto :end
)

python -c "import openpyxl, matplotlib, pptx, pandas, flask, PIL" 2>nul
if errorlevel 1 (
    echo [INFO] Installing dependencies...
    python -m pip install -r app\requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 python -m pip install -r app\requirements.txt
    echo.
)

echo [START] Launching GUI...
python app\main.py
echo.
echo [DONE] Program exited.

:end
echo.
echo Press any key to close...
pause >nul
