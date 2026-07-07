@echo off
cd /d "%~dp0"
set LOGFILE=%~dp0logs\run_%date:~0,4%%date:~5,2%%date:~8,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log
set LOGFILE=%LOGFILE: =0%
if not exist "%~dp0logs" mkdir "%~dp0logs"

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
echo Log: %LOGFILE%
python app\main.py >> "%LOGFILE%" 2>&1
echo [DONE] Program exited. >> "%LOGFILE%"

:end
pause >nul
