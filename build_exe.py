"""
Excel 统一分析工具 - 一键打包脚本
双击 run_build.bat 或执行 python build_exe.py 即可打包
"""

import os
import sys
import shutil
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
APP_DIR = os.path.join(BASE_DIR, "app")
EXE_NAME = "Excel统一分析工具"

def check_python():
    print("[1/4] 检查 Python 环境...")
    print(f"  Python 版本: {sys.version.split()[0]}  [OK]")
    print()

def install_requirements():
    print("[2/4] 安装项目依赖...")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "-r", "requirements.txt"],
        cwd=APP_DIR,
    )
    if result.returncode != 0:
        print("  [错误] 依赖安装失败，请检查 requirements.txt")
        input("\n按回车键退出...")
        sys.exit(1)
    print("  依赖安装完成  [OK]")
    print()

def install_pyinstaller():
    print("[3/4] 安装 PyInstaller 打包工具...")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "pyinstaller"],
    )
    if result.returncode != 0:
        print("  [错误] PyInstaller 安装失败")
        input("\n按回车键退出...")
        sys.exit(1)
    print("  PyInstaller 安装完成  [OK]")
    print()

def build_exe():
    print("[4/4] 开始打包（单文件 GUI 模式，约需 2-5 分钟）...")
    print()

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", EXE_NAME,
        "--add-data", f"static{os.pathsep}static",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.ttk",
        "--hidden-import", "tkinter.scrolledtext",
        "--hidden-import", "openpyxl",
        "--hidden-import", "matplotlib",
        "--hidden-import", "matplotlib.backends.backend_agg",
        "--hidden-import", "pptx",
        "--hidden-import", "pptx.chart.data",
        "--hidden-import", "pptx.enum.chart",
        "--hidden-import", "pptx.enum.shapes",
        "--hidden-import", "pptx.enum.text",
        "--hidden-import", "pptx.oxml.ns",
        "--hidden-import", "lxml.etree",
        "--hidden-import", "pandas",
        "--hidden-import", "PIL",
        "--hidden-import", "cartopy",
        "--hidden-import", "numpy",
        "main.py",
    ]

    result = subprocess.run(cmd, cwd=APP_DIR)
    if result.returncode != 0:
        print("\n  [错误] 打包失败，请查看上方错误信息")
        input("\n按回车键退出...")
        sys.exit(1)

    dist_exe = os.path.join(APP_DIR, "dist", f"{EXE_NAME}.exe")
    output_exe = os.path.join(BASE_DIR, f"{EXE_NAME}.exe")

    if os.path.exists(dist_exe):
        shutil.move(dist_exe, output_exe)
        size_mb = os.path.getsize(output_exe) / (1024 * 1024)
        print()
        print("=" * 60)
        print("  打包成功！")
        print("=" * 60)
        print(f"  输出文件: {output_exe}")
        print(f"  文件大小: {size_mb:.2f} MB")
        print()
        print("  使用方法: 双击 exe 文件即可启动 GUI 界面")
        print("=" * 60)
    else:
        print("  [错误] 未找到生成的 exe 文件")

    build_dir = os.path.join(APP_DIR, "build")
    dist_dir = os.path.join(APP_DIR, "dist")
    spec_file = os.path.join(APP_DIR, f"{EXE_NAME}.spec")
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir, ignore_errors=True)
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir, ignore_errors=True)
    if os.path.exists(spec_file):
        os.remove(spec_file)

def main():
    print("=" * 60)
    print("  Excel 统一分析工具 - 一键打包脚本")
    print("=" * 60)
    print()

    check_python()
    install_requirements()
    install_pyinstaller()
    build_exe()

    print()
    input("按回车键退出...")

if __name__ == "__main__":
    main()
