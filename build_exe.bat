@echo off
setlocal enabledelayedexpansion
title Excel ͳһ�������� - ����ű�

echo ============================================================
echo   Excel ͳһ�������� - һ������ű�
echo ============================================================
echo.

REM �л����ű�����Ŀ¼�� app ��Ŀ¼
cd /d "%~dp0app"

REM ===== 1. ��� Python =====
echo [1/4] ��� Python ����...
python --version >nul 2>&1
if errorlevel 1 (
    echo   [����] δ�ҵ� Python�����Ȱ�װ Python 3.8+ �����ӵ� PATH
    echo   ���ص�ַ: https://www.python.org/downloads/
    pause
    exit /b 1
)
for /f "tokens=2" %%i in ('python --version') do set PYVER=%%i
echo   Python �汾: %PYVER%  [OK]
echo.

REM ===== 2. ��װ��Ŀ���� =====
echo [2/4] ��װ��Ŀ����...
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo   [����] ������װʧ�ܣ����� requirements.txt
    pause
    exit /b 1
)
echo   ������װ���  [OK]
echo.

REM ===== 3. ��װ PyInstaller =====
echo [3/4] ��װ PyInstaller �������...
python -m pip install pyinstaller
if errorlevel 1 (
    echo   [����] PyInstaller ��װʧ��
    pause
    exit /b 1
)
echo   PyInstaller ��װ���  [OK]
echo.

REM ===== 4. ��ʼ��� =====
echo [4/4] ��ʼ��������ļ� GUI ģʽ��Լ�� 2-5 ���ӣ�...
echo.

python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "Excelͳһ��������" ^
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
    echo   [����] ���ʧ�ܣ���鿴�Ϸ�������Ϣ
    pause
    exit /b 1
)

REM ===== �ƶ� exe ����Ŀ��Ŀ¼ =====
if exist "dist\Excelͳһ��������.exe" (
    move /y "dist\Excelͳһ��������.exe" "%~dp0Excelͳһ��������.exe" >nul
    echo.
    echo ============================================================
    echo   ����ɹ���
    echo ============================================================
    echo   ����ļ�: %~dp0Excelͳһ��������.exe
    echo   �ļ���С:
    for %%A in ("%~dp0Excelͳһ��������.exe") do echo     %%~zA �ֽ�
    echo.
    echo   ʹ�÷���: ˫�� exe �ļ��������� GUI ����
    echo ============================================================
) else (
    echo   [����] δ�ҵ����ɵ� exe �ļ�
)

REM ������ʱ�ļ�
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "Excelͳһ��������.spec" del /q "Excelͳһ��������.spec"

echo.
pause
