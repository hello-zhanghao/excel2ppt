"""
防护用例一键测试脚本
用法: python run_test.py

执行流程:
  1. 重建测试数据
  2. 运行透视分析（pivot）
  3. 运行PPT生成（ppt），引用透视结果
  4. 读取并展示输出结果，供人工检查
  5. PPT转图片 + 生成HTML报告（手机可查看）
"""
import os
import sys
import subprocess
import openpyxl
import glob
import base64

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# SCRIPT_DIR = .../cases/00_防护用例
# PROJECT_DIR = .../excel2ppt  (项目根目录，向上退两级)
PROJECT_DIR = os.path.dirname(os.path.dirname(SCRIPT_DIR))
PYTHON = sys.executable

# 颜色输出
GREEN = "\033[92m"
RED = "\033[91m"
YELLOW = "\033[93m"
CYAN = "\033[96m"
RESET = "\033[0m"
BOLD = "\033[1m"


def print_header(title):
    print(f"\n{CYAN}{'='*60}")
    print(f"  {BOLD}{title}")
    print(f"{'='*60}{RESET}")


def run_cmd(cmd, cwd=PROJECT_DIR):
    """运行命令，返回是否成功"""
    print(f"  {YELLOW}$ {cmd}{RESET}")
    result = subprocess.run(cmd, shell=True, cwd=cwd, capture_output=True, text=True, encoding="utf-8")
    if result.stdout.strip():
        print(result.stdout.rstrip())
    if result.returncode != 0:
        print(f"  {RED}[FAIL] 退出码={result.returncode}{RESET}")
        if result.stderr.strip():
            print(f"  {RED}stderr: {result.stderr.rstrip()}{RESET}")
        return False
    return True


def find_latest_output():
    """查找最新的输出目录"""
    output_dirs = glob.glob(os.path.join(SCRIPT_DIR, "output_*"))
    if not output_dirs:
        return None
    return max(output_dirs, key=os.path.getmtime)


def inspect_excel(excel_path):
    """读取并展示透视分析Excel输出"""
    print_header("Excel 输出内容检查")
    if not excel_path or not os.path.exists(excel_path):
        print(f"  {RED}Excel 文件不存在{RESET}")
        return

    print(f"  文件: {os.path.basename(excel_path)}\n")
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    print(f"  Sheet 列表: {wb.sheetnames}\n")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        print(f"  {BOLD}--- {sheet_name} ({len(rows)-1} 行数据) ---{RESET}")

        # 找到真正的表头行（跳过区块标题行）
        header_row_idx = 0
        for i, row in enumerate(rows[:3]):
            non_none = [c for c in row if c is not None]
            if len(non_none) >= 2 and any(isinstance(c, str) and c.strip() for c in row):
                header_row_idx = i
                break

        headers = [str(c) if c is not None else "" for c in rows[header_row_idx]]
        # 只保留非空表头
        col_indices = [i for i, h in enumerate(headers) if h.strip()]
        headers = [headers[i] for i in col_indices]

        # 打印表头
        print(f"    列名: {' | '.join(headers)}")

        # 打印数据行
        for row in rows[header_row_idx + 1:]:
            vals = []
            for i in col_indices:
                v = row[i] if i < len(row) else None
                if v is None:
                    vals.append("")
                elif isinstance(v, float):
                    vals.append(f"{v:.4f}" if v != int(v) else str(int(v)))
                else:
                    vals.append(str(v))
            # 跳过全空行
            if any(v.strip() for v in vals):
                print(f"    {' | '.join(vals)}")
        print()
    wb.close()


def inspect_ppt(ppt_path):
    """检查PPT输出"""
    print_header("PPT 输出检查")
    if not ppt_path or not os.path.exists(ppt_path):
        print(f"  {RED}PPT 文件不存在{RESET}")
        return

    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    print(f"  文件: {os.path.basename(ppt_path)}")
    print(f"  文件大小: {os.path.getsize(ppt_path) / 1024:.1f} KB\n")

    prs = Presentation(ppt_path)
    print(f"  总页数: {len(prs.slides)}\n")

    for idx, slide in enumerate(prs.slides, 1):
        print(f"  {BOLD}--- 第 {idx} 页 ---{RESET}")

        # 收集文本
        texts = []
        charts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    texts.append(text)
            if shape.has_chart:
                chart = shape.chart
                charts.append(chart)

        if texts:
            print(f"    文本内容:")
            for t in texts:
                # 截断过长的文本
                t_display = t.replace("\n", " | ") if len(t) < 100 else t[:100] + "..."
                print(f"      - {t_display}")

        if charts:
            print(f"    图表数量: {len(charts)}")
            for ci, chart in enumerate(charts, 1):
                chart_type = str(chart.chart_type) if chart.chart_type else "未知"
                print(f"      图表{ci}: 类型={chart_type}")
                if chart.has_title and chart.chart_title:
                    print(f"               标题={chart.chart_title.text_frame.text}")

                # 打印图表数据
                try:
                    plot = chart.plots[0]
                    categories = list(plot.categories)
                    print(f"               X轴类别: {categories}")

                    for si, series in enumerate(plot.series):
                        vals = list(series.values)
                        name = series.name if hasattr(series, "name") else f"系列{si+1}"
                        # 截断显示
                        vals_str = ", ".join([f"{v:.2f}" if isinstance(v, float) else str(v) for v in vals])
                        if len(vals_str) > 80:
                            vals_str = vals_str[:80] + "..."
                        print(f"               系列[{name}]: {vals_str}")
                except Exception as e:
                    print(f"               (数据读取异常: {e})")

        if not texts and not charts:
            print(f"    (空页)")
        print()


def pptx_to_images(ppt_path, output_dir):
    """用 PowerPoint COM 接口将 PPT 每页导出为 PNG 图片"""
    import win32com.client
    import pythoncom

    os.makedirs(output_dir, exist_ok=True)
    images = []
    abs_ppt = os.path.abspath(ppt_path)

    pythoncom.CoInitialize()
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        # 不可见，避免弹窗
        try:
            ppt_app.Visible = 0
        except Exception:
            pass
        pres = ppt_app.Presentations.Open(abs_ppt, WithWindow=False)

        for i, slide in enumerate(pres.Slides, 1):
            img_path = os.path.join(output_dir, f"slide_{i:02d}.png")
            # 导出为 PNG，分辨率 1280x720
            slide.Export(img_path, "PNG", 1280, 720)
            images.append(img_path)

        pres.Close()
        ppt_app.Quit()
    except Exception as e:
        print(f"  {YELLOW}PPT转图片失败: {e}{RESET}")
        # 退出时确保进程关闭
        try:
            ppt_app.Quit()
        except Exception:
            pass
    finally:
        pythoncom.CoUninitialize()

    return images


def generate_html_report(excel_path, ppt_path, output_dir):
    """使用 html_builder 生成 HTML 报告"""
    print_header("生成 HTML 报告（手机可查看）")
    
    sys.path.insert(0, PROJECT_DIR)
    from app.src.html_builder import generate_html_report as build_html
    
    html_path = build_html(
        excel_path=excel_path,
        ppt_path=ppt_path,
        output_dir=output_dir,
        report_title="防护用例测试报告",
        report_subtitle="透视分析结果 + PPT预览",
    )
    return html_path


def start_preview_server(html_path):
    """使用 html_builder 启动预览服务器"""
    from app.src.html_builder import start_preview_server as start_server
    return start_server(html_path)


def main():
    print_header("防护用例测试 - 开始")

    # Step 0: 确认数据文件存在
    config_path = os.path.join(SCRIPT_DIR, "项目配置.xlsx")
    data_path = os.path.join(SCRIPT_DIR, "测试数据.xlsx")
    if not os.path.exists(config_path) or not os.path.exists(data_path):
        print(f"  {YELLOW}数据/配置文件不存在，先运行 创建测试数据.py{RESET}")
        run_cmd(f'"{PYTHON}" "{os.path.join(SCRIPT_DIR, "创建测试数据.py")}"')

    # Step 1: 透视分析
    print_header("Step 1: 透视分析 (pivot)")
    ok = run_cmd(
        f'"{PYTHON}" "{os.path.join(PROJECT_DIR, "app", "main.py")}" pivot '
        f'-c "{config_path}"'
    )
    if not ok:
        print(f"\n{RED}{BOLD}透视分析失败！请检查上方错误信息。{RESET}")
        sys.exit(1)

    # Step 2: 查找透视分析输出
    output_dir = find_latest_output()
    if not output_dir:
        print(f"  {RED}未找到输出目录{RESET}")
        sys.exit(1)

    # 找到透视结果 Excel 文件
    excel_files = glob.glob(os.path.join(output_dir, "*_分析_*.xlsx"))
    if not excel_files:
        # 尝试其他命名模式
        excel_files = glob.glob(os.path.join(output_dir, "*.xlsx"))
    pivot_excel = excel_files[0] if excel_files else None

    # Step 3: 生成 PPT
    print_header("Step 2: PPT 生成 (ppt)")
    ok = run_cmd(
        f'"{PYTHON}" "{os.path.join(PROJECT_DIR, "app", "main.py")}" ppt '
        f'-c "{config_path}" --pivot-file "{pivot_excel}"'
    )
    if not ok:
        print(f"\n{RED}{BOLD}PPT 生成失败！请检查上方错误信息。{RESET}")
        sys.exit(1)

    # Step 4: 查找 PPT 文件（可能在不同的时间戳目录中）
    all_output_dirs = sorted(glob.glob(os.path.join(SCRIPT_DIR, "output_*")), key=os.path.getmtime)
    ppt_path = None
    for d in reversed(all_output_dirs):
        ppt_files = glob.glob(os.path.join(d, "*.pptx"))
        if ppt_files:
            ppt_path = ppt_files[0]
            break

    # Step 5: 检查输出
    inspect_excel(pivot_excel)
    inspect_ppt(ppt_path)

    # Step 6: 生成HTML报告（手机可查看）
    ppt_dir = os.path.dirname(ppt_path) if ppt_path else SCRIPT_DIR
    html_path = generate_html_report(pivot_excel, ppt_path, ppt_dir)

    # 总结
    print_header("测试总结")
    print(f"  Excel: {pivot_excel}" if pivot_excel else "  Excel: N/A")
    print(f"  PPT:   {ppt_path}" if ppt_path else "  PPT:   N/A")
    print(f"  HTML:  {html_path}" if html_path else "  HTML:  N/A")

    pivot_ok = pivot_excel is not None and os.path.exists(pivot_excel)
    ppt_ok = ppt_path is not None and os.path.exists(ppt_path)

    print()
    if pivot_ok and ppt_ok:
        print(f"  {GREEN}{BOLD}✓ 全部通过 - 请人工检查输出格式{RESET}")
    else:
        print(f"  {RED}{BOLD}✗ 存在失败项{RESET}")
        if not pivot_ok:
            print(f"    - Excel 生成失败")
        if not ppt_ok:
            print(f"    - PPT 生成失败")

    # 启动本地服务器供预览
    if html_path and os.path.exists(html_path):
        print(f"\n  {CYAN}启动预览服务器...{RESET}")
        start_preview_server(html_path)

    print(f"\n  {CYAN}提示: 请打开输出文件检查格式是否正常{RESET}")
    print(f"  {CYAN}  Excel: {pivot_excel}{RESET}" if pivot_excel else "")
    print(f"  {CYAN}  PPT:   {ppt_path}{RESET}" if ppt_path else "")


if __name__ == "__main__":
    main()
