@echo off
chcp 65001 >nul
echo 正在打包 Excel 统一分析工具...

cd /d "%~dp0app"

REM 打包为单文件模式（GUI窗口模式，无控制台）
pyinstaller --noconfirm --onefile --windowed ^
    --name "Excel统一分析工具" ^
    --add-data "static;static" ^
    --hidden-import "tkinter" ^
    --hidden-import "tkinter.filedialog" ^
    --hidden-import "tkinter.ttk" ^
    --hidden-import "tkinter.scrolledtext" ^
    --hidden-import "openpyxl" ^
    --hidden-import "matplotlib" ^
    --hidden-import "pptx" ^
    --hidden-import "pptx.chart.data" ^
    --hidden-import "pptx.enum.chart" ^
    --hidden-import "pptx.enum.shapes" ^
    --hidden-import "pptx.enum.text" ^
    --hidden-import "pptx.oxml.ns" ^
    --hidden-import "lxml.etree" ^
    --hidden-import "pandas" ^
    --hidden-import " PIL" ^
    --hidden-import "cartopy" ^
    --hidden-import "numpy" ^
    main.py

echo.
echo 打包完成！
echo 输出文件: dist\Excel统一分析工具.exe
echo.
pause