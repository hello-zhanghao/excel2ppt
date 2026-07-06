@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Excel 统一分析工具

:: ===== 自动检测 Python =====
set PYTHON=
where python >nul 2>&1 && set PYTHON=python
where python3 >nul 2>&1 && set PYTHON=python3
where py >nul 2>&1 && set PYTHON=py
if "%PYTHON%"=="" (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo         下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [检测] Python: %PYTHON%
%PYTHON% --version

:: ===== 检查依赖 =====
%PYTHON% -c "import openpyxl, matplotlib, pptx, pandas, flask, PIL" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [依赖] 正在安装缺失的依赖包...
    %PYTHON% -m pip install -r app\requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo [警告] 清华源安装失败，尝试默认源...
        %PYTHON% -m pip install -r app\requirements.txt
    )
    echo.
)

:: ===== 启动 GUI =====
echo [启动] 正在启动 Excel 统一分析工具...
%PYTHON% app\main.py
if errorlevel 1 (
    echo.
    echo [错误] 程序异常退出
    pause
)
