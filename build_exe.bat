@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title Excel 统一分析工具 - 打包脚本

echo ============================================================
echo   Excel 统一分析工具 - 一键打包脚本
echo ============================================================
echo.

REM 切换到脚本所在目录的 app 子目录
cd /d "%~dp0app"

REM ===== 1. 检查 Python =====
echo [1/4] 检查 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo   [错误] 未找到 Python，请先安装 Python 3.8+ 并添加到 PATH
    echo   下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version') do set PYVER=%%i
echo   Python 版本: %PYVER%  [OK]
echo.

REM ===== 2. 安装项目依赖 =====
echo [2/4] 安装项目依赖...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo   [错误] 依赖安装失败，请检查 requirements.txt
    pause
    exit /b 1
)
echo   依赖安装完成  [OK]
echo.

REM ===== 3. 安装 PyInstaller =====
echo [3/4] 安装 PyInstaller 打包工具...
python -m pip install pyinstaller
if errorlevel 1 (
    echo   [错误] PyInstaller 安装失败
    pause
    exit /b 1
)
echo   PyInstaller 安装完成  [OK]
echo.

REM ===== 4. 开始打包 =====
echo [4/4] 开始打包（单文件 GUI 模式，约需 2-5 分钟）...
echo.

python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "Excel统一分析工具" ^
    --add-data "static;static" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.scrolledtext" ^
    --hidden-import "openpyxl" ^
    --hidden-import "matplotlib" ^
    --hidden-import "matplotlib.backends.backend_agg" ^
    --hidden-import "pptx" ^
    --hidden-import "pptx.chart.data" ^
    --hidden-import "pptx.enum.chart" ^
    --hidden-import "pptx.enum.shapes" ^
    --hidden-import "pptx.enum.text" ^
    --hidden-import "pptx.oxml.ns" ^
    --hidden-import "lxml.etree" ^
    --hidden-import "pandas" ^
    --hidden-import "PIL" ^
    --hidden-import "cartopy" ^
    --hidden-import "numpy" ^
    main.py

if errorlevel 1 (
    echo.
    echo   [错误] 打包失败，请查看上方错误信息
    pause
    exit /b 1
)

REM ===== 移动 exe 到项目根目录 =====
if exist "dist\Excel统一分析工具.exe" (
    move /y "dist\Excel统一分析工具.exe" "%~dp0Excel统一分析工具.exe" >nul
    echo.
    echo ============================================================
    echo   打包成功！
    echo ============================================================
    echo   输出文件: %~dp0Excel统一分析工具.exe
    echo   文件大小:
    for %%A in ("%~dp0Excel统一分析工具.exe") do echo     %%~zA 字节
    echo.
    echo   使用方法: 双击 exe 文件即可启动 GUI 界面
    echo ============================================================
) else (
    echo   [错误] 未找到生成的 exe 文件
)

REM 清理临时文件
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Excel统一分析工具.spec" del /q "Excel统一分析工具.spec"

echo.
pause
