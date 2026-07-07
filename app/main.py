"""
Excel 统一分析工具 — PPT 生成 + 透视分析 + HTML报告
用法：
  python main.py ppt 文件夹路径           ← PPT 生成（自动找配置和数据）
  python main.py ppt -c 配置.xlsx         ← PPT 生成（指定配置）
  python main.py ppt -c 配置.xlsx -o out.pptx
  python main.py pivot 文件夹路径         ← 透视分析（自动找配置和数据）
  python main.py pivot -c 配置.xlsx       ← 透视分析（指定配置）
  python main.py pivot -c 配置.xlsx -o out.xlsx
  python main.py html -c 配置.xlsx        ← 生成HTML报告（自动模式）
  python main.py html --pivot-file 分析.xlsx --ppt-file 报告.pptx
  python main.py template 模板.pptx --pivot 透视结果.xlsx   ← 基于PPT模板填充数据
  python main.py 文件夹路径               ← 自动检测配置类型，分发到对应模式
"""
import os
import sys
import argparse
import glob
from datetime import datetime

# 版本信息
__VERSION__ = "2.16.1"
__UPDATE_DATE__ = "2026-07-07"

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def find_config_file(folder):
    xlsx_files = glob.glob(os.path.join(folder, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]
    if not xlsx_files:
        return None

    # 优先：文件名含「配置」/「config」的候选
    config_candidates = [f for f in xlsx_files
                         if "配置" in os.path.basename(f) or "config" in os.path.basename(f).lower()]

    # 无「配置」候选时，若只有一个 xlsx 则直接用
    if not config_candidates:
        if len(xlsx_files) == 1:
            return xlsx_files[0]
        # 多个 xlsx 但都不含「配置」→ 也纳入候选让用户选
        config_candidates = xlsx_files

    # 只有一个候选，直接返回
    if len(config_candidates) == 1:
        return config_candidates[0]

    # 多个候选：交互式选择
    return _prompt_select_config(config_candidates)


def _prompt_select_config(candidates):
    """多个配置文件时让用户选择。非交互环境取第一个并警告。"""
    # 非交互式（管道/重定向），不能 input()，取第一个并警告
    if not sys.stdin.isatty():
        chosen = candidates[0]
        print(f"[警告] 检测到 {len(candidates)} 个配置文件，非交互环境自动选择: {os.path.basename(chosen)}")
        print(f"        其他配置文件: {', '.join(os.path.basename(f) for f in candidates[1:])}")
        return chosen

    print(f"\n[信息] 检测到 {len(candidates)} 个配置文件，请选择:")
    for i, f in enumerate(candidates, 1):
        print(f"  {i}. {os.path.basename(f)}")

    while True:
        try:
            choice = input(f"请输入序号 (1-{len(candidates)})，回车默认选 1: ").strip()
            if not choice:
                return candidates[0]
            idx = int(choice)
            if 1 <= idx <= len(candidates):
                return candidates[idx - 1]
            print(f"[错误] 序号超出范围，请输入 1-{len(candidates)}")
        except ValueError:
            print("[错误] 请输入数字")
        except (EOFError, KeyboardInterrupt):
            print("\n[信息] 已取消选择")
            return None


def _auto_find_data_file(config_dir, config_path):
    """
    自动查找数据文件，支持 Excel 和 CSV。
    优先查找与配置同目录下的非配置数据文件。
    """
    from src.excel_reader import get_candidate_data_files
    candidates = get_candidate_data_files(config_dir)
    return candidates[0] if candidates else None


def _ensure_output_dir(config_dir):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(config_dir, f"output_{timestamp}")
    os.makedirs(out_dir, exist_ok=True)
    return out_dir


def _resolve_config_path(args):
    config_path = args.config
    folder_or_config = getattr(args, "folder_or_config", None)
    if not config_path and folder_or_config:
        if os.path.isdir(folder_or_config):
            config_path = find_config_file(folder_or_config)
            if config_path:
                print(f"[信息] 使用配置文件: {os.path.basename(config_path)}")
            else:
                print(f"[错误] 在 {folder_or_config} 里没找到 Excel 文件")
                sys.exit(1)
        elif os.path.isfile(folder_or_config):
            config_path = folder_or_config
    if not config_path:
        print("[错误] 请指定文件夹路径或配置文件路径")
        sys.exit(1)
    if not os.path.exists(config_path):
        print(f"[错误] 配置文件不存在: {config_path}")
        sys.exit(1)
    return os.path.abspath(config_path)


def _detect_mode(config_path):
    import openpyxl
    try:
        wb = openpyxl.load_workbook(config_path, read_only=True)
    except Exception as e:
        print(f"[错误] 无法读取配置文件: {e}")
        sys.exit(1)
    ppt_keywords = {"页码", "页面类型", "页面标题", "图表类型"}
    pivot_keywords = {"数据源", "行维度", "列维度", "值字段", "聚合方式"}
    ppt_found = False
    pivot_found = False
    for name in wb.sheetnames:
        ws = wb[name]
        # 检查前 5 行（PPT 配置第一行可能是主题设置行）
        all_texts = set()
        for row in ws.iter_rows(min_row=1, max_row=5):
            for cell in row:
                v = cell.value
                if v is not None:
                    all_texts.add(str(v).strip())
        if len(ppt_keywords & all_texts) >= 2:
            ppt_found = True
        if len(pivot_keywords & all_texts) >= 2:
            pivot_found = True
    wb.close()
    if ppt_found and pivot_found:
        return "all"
    if pivot_found:
        return "pivot"
    if ppt_found:
        return "ppt"
    return "unknown"


def _append_chart(chart_map, chart_title, chart_def, page_title):
    """追加图表到 chart_map（list），同名时打印警告但不丢弃。"""
    same_name = [c for c in chart_map if c.get("图表标题") == chart_title]
    if same_name:
        print(f"    [警告] [{page_title}] 图表标题 '{chart_title}' 与已有图表重名，两者均保留")
    chart_map.append(chart_def)


def _run_ppt_mode(config_path, output_path=None, pivot_data_file=None, validate_only=False):
    """
    生成PPT。
    :param pivot_data_file: 透视分析结果文件路径，用于 {pivot} 数据源引用
    :param validate_only: 仅校验配置不执行生成
    """
    from src.excel_reader import (
        read_config, read_data, read_geo_data, find_data_file, get_data_file_info
    )
    from src.ppt_builder import build_ppt, validate_ppt_config, print_ppt_validation_results

    config_dir = os.path.dirname(config_path)
    print(f"[PPT/1] 读取配置: {config_path}")
    config = read_config(config_path)
    general = config.get("general", {})
    pages = config.get("pages", [])
    colors = config.get("colors", {})
    print(f"    → PPT页面: {len(pages)} 页")

    # 配置校验
    print(f"[PPT/校验] 检查配置...")
    val_results, all_ok = validate_ppt_config(config, config_dir, pivot_data_file)
    print_ppt_validation_results(val_results)

    if validate_only:
        print("[校验] 仅校验模式，不执行生成")
        return None

    if not all_ok:
        print("[错误] 配置校验未通过，请修正后再执行")
        sys.exit(1)

    # 跳过「是否生成=否」的页面
    skip_keywords = ("否", "no", "false", "0", "不生成", "跳过", "skip")
    filtered_pages = []
    for page in pages:
        should_gen = str(page.get("是否生成", "是")).strip().lower()
        if should_gen in skip_keywords:
            print(f"    [SKIP] 第{page.get('页码','?')}页：是否生成=否")
            continue
        filtered_pages.append(page)
    pages = filtered_pages
    config["pages"] = pages

    if pivot_data_file:
        print(f"    → 透视结果数据源: {os.path.basename(pivot_data_file)}")
    else:
        # 检测是否引用了透视结果但未提供文件，提前给出明确提示
        pivot_keywords = ("{pivot}", "pivot", "透视结果", "透视分析")
        refs_pivot = any(
            str(cd.get("数据源", "")).strip().lower() in pivot_keywords
            for page in pages for cd in page.get("charts", [])
        )
        if refs_pivot:
            print(f"    [警告] 配置引用了 {{pivot}} 数据源但未提供 --pivot-file，相关图表将被跳过")
            print(f"           建议：使用 'python main.py 配置.xlsx' 自动模式（先透视后PPT），")
            print(f"                 或先用 pivot 子命令生成结果，再用 'ppt --pivot-file 结果.xlsx' 指定")

    # 全局数据文件路径
    default_data_file = general.get("数据文件", general.get("excel_path", ""))
    if not default_data_file:
        default_data_file = _auto_find_data_file(config_dir, config_path)
        if default_data_file:
            print(f"    → 自动找到数据文件: {os.path.basename(default_data_file)}")
    elif not os.path.isabs(str(default_data_file)):
        default_data_file = os.path.join(config_dir, str(default_data_file))

    print(f"    → 数据文件: {os.path.basename(default_data_file) if default_data_file else '无'}")

    print(f"[PPT/2] 读取图表数据...")
    total_charts = 0
    chart_map = []  # 使用 list 避免同名图表键覆盖

    for page_def in pages:
        page_title = str(page_def.get("页面标题", f"第{page_def.get('页码', '?')}页"))
        for chart_def in page_def.get("charts", []):
            chart_title = chart_def.get("图表标题", "")
            chart_type = str(chart_def.get("图表类型", "")).strip().lower()
            data_sheet = chart_def.get("数据Sheet", "Sheet1")
            data_source = chart_def.get("数据源", "")
            x_range = chart_def.get("X轴范围", "")
            y_range = chart_def.get("Y轴范围", "")
            block_name = chart_def.get("区块名", "")

            # 查找数据文件：根据图表指定的数据源或全局默认
            if data_source:
                ds_lower = str(data_source).strip().lower()
                # 支持 {pivot} / pivot / 透视结果 引用透视分析结果
                if ds_lower in ("{pivot}", "pivot", "透视结果", "透视分析"):
                    if pivot_data_file and os.path.exists(pivot_data_file):
                        file_path = pivot_data_file
                        print(f"    [信息] [{page_title}] {chart_title} 使用透视结果数据")
                    else:
                        print(f"    ! [{page_title}] {chart_title} 引用透视结果但未找到透视输出文件")
                        continue
                else:
                    file_path = find_data_file(data_source, config_dir)
                    if file_path:
                        print(f"    [信息] [{page_title}] {chart_title} 使用数据源: {os.path.basename(file_path)}")
                    else:
                        print(f"    ! [{page_title}] {chart_title} 数据源 '{data_source}' 未找到")
                        continue
            else:
                file_path = default_data_file
            
            if not file_path or not os.path.exists(str(file_path)):
                print(f"    ! [{page_title}] {chart_title} 数据文件不存在: {file_path}")
                continue

            try:
                if chart_type in ("map", "heatmap"):
                    geo_df = read_geo_data(file_path, data_sheet, x_range, y_range)
                    if geo_df is not None and len(geo_df) > 0:
                        chart_def["_geo_df"] = geo_df
                        chart_def["_is_map"] = True
                        _append_chart(chart_map, chart_title, chart_def, page_title)
                        total_charts += 1
                        print(f"    [OK] [{page_title}] {chart_title} (地图)")
                    else:
                        print(f"    - [{page_title}] {chart_title} (无地理数据)")
                    continue

                x_values, y_values = read_data(file_path, data_sheet, x_range, y_range, block_name)
                if x_values and y_values:
                    chart_def["_categories"] = x_values
                    chart_def["_values"] = y_values
                    if isinstance(x_values, list) and x_values and isinstance(x_values[0], (tuple, list)):
                        chart_def["_is_hierarchical"] = True
                    _append_chart(chart_map, chart_title, chart_def, page_title)
                    total_charts += 1
                    print(f"    [OK] [{page_title}] {chart_title}")
                else:
                    print(f"    - [{page_title}] {chart_title} (无数据)")
            except KeyError as e:
                print(f"    ! [{page_title}] {chart_title} 列名未找到: {e}")
                print(f"      建议: 检查 X轴='{x_range}' Y轴='{y_range}' 是否与数据 Sheet '{data_sheet}' 的列名完全一致")
            except (FileNotFoundError, PermissionError) as e:
                print(f"    ! [{page_title}] {chart_title} 数据文件访问失败: {e}")
            except Exception as e:
                msg = str(e).lower()
                if "sheet" in msg or "worksheet" in msg:
                    print(f"    ! [{page_title}] {chart_title} Sheet不存在: {e}")
                    print(f"      建议: 检查 数据Sheet='{data_sheet}' 是否存在于 {os.path.basename(file_path)}")
                elif "column" in msg or "key" in msg:
                    print(f"    ! [{page_title}] {chart_title} 列名匹配失败: {e}")
                    print(f"      建议: 检查 X轴='{x_range}' Y轴='{y_range}' 列名是否精确匹配")
                else:
                    print(f"    ! [{page_title}] {chart_title} (错误: {e})")

    print(f"    → 共 {total_charts} 个图表")
    if total_charts == 0 and not pages:
        print("[错误] 没有有效内容。")
        sys.exit(1)

    if not output_path:
        # 有透视结果文件时，复用其所在目录，避免 Excel 和 PPT 分散到不同目录
        if pivot_data_file and os.path.exists(str(pivot_data_file)):
            output_dir = os.path.dirname(os.path.abspath(str(pivot_data_file)))
        else:
            output_dir = _ensure_output_dir(config_dir)
        base_name = os.path.splitext(os.path.basename(config_path))[0]
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"{base_name}_报告_{timestamp_str}.pptx")
    else:
        output_dir = os.path.dirname(output_path) or "."
        os.makedirs(output_dir, exist_ok=True)

    if os.path.exists(output_path):
        print(f"    [警告] 输出文件已存在，将被覆盖: {os.path.basename(output_path)}")
    print(f"[PPT/3] 生成PPT: {output_path}")
    build_ppt(config, chart_map, output_path)
    print(f"\n[OK] 完成！PPT已保存至: {output_path}")
    print(f"   共 {len(pages)} 页, {total_charts} 个原生图表")
    print(f"   提示: 双击PPT中的图表可查看和编辑数据")


def _run_pivot_mode(config_path, output_path=None, validate_only=False):
    from src.pivot_analyzer import (
        read_pivot_config, run_analysis, find_data_files,
        validate_pivot_config, print_validation_results,
        collect_task_scalars
    )
    from src.excel_writer import write_results

    config_dir = os.path.dirname(config_path)
    print(f"[Pivot/1] 读取配置: {config_path}")
    tasks = read_pivot_config(config_path)
    print(f"    → 共 {len(tasks)} 个分析任务")

    if not tasks:
        print("[错误] 没有有效的分析任务。")
        sys.exit(1)

    data_files = find_data_files(config_dir, config_path)
    if data_files:
        print(f"    → 找到数据文件: {[os.path.basename(f) for f in data_files]}")

    # 配置校验
    print(f"[Pivot/校验] 检查配置...")
    val_results = validate_pivot_config(tasks, config_dir)
    all_ok = print_validation_results(val_results)
    
    if validate_only:
        print("[校验] 仅校验模式，不执行分析")
        return None
    
    if not all_ok:
        print("[错误] 配置校验未通过，请修正后再执行")
        sys.exit(1)

    print(f"[Pivot/2] 执行透视分析...")
    results = []
    errors = []
    skipped = 0
    scalar_context: dict = {}

    for task in tasks:
        seq = task.get("序号", "?")
        sheet_name = task.get("结果Sheet", f"结果{seq}")
        
        should_calc = str(task.get("是否计算", "是")).strip()
        if should_calc.lower() in ("否", "no", "false", "0", "不计算", "跳过", "skip"):
            print(f"    [SKIP] [任务{seq}] 已设置为不计算，跳过")
            results.append(None)
            skipped += 1
            continue
        
        try:
            result, error = run_analysis(task, config_dir, scalar_context)
            if error:
                print(f"    [FAIL] [任务{seq}] {error}")
                errors.append({"序号": seq, "错误": error})
                results.append(None)
            else:
                if isinstance(result, dict):
                    for key, df in result.items():
                        if hasattr(df, "shape"):
                            print(f"    [OK] [任务{seq}] {sheet_name} -> {df.shape[0]}行 x {df.shape[1]}列")
                        else:
                            print(f"    [OK] [任务{seq}] {sheet_name} (标量)")
                else:
                    print(f"    [OK] [任务{seq}] {sheet_name}")
                results.append(result)
                # 收集无行维度任务产生的标量，供后续任务公式引用
                task_scalars = collect_task_scalars(result)
                if task_scalars:
                    scalar_context.update(task_scalars)
        except Exception as e:
            print(f"    [FAIL] [任务{seq}] 异常: {e}")
            errors.append({"序号": seq, "错误": str(e)})
            results.append(None)

    valid_results = [r for r in results if r is not None]
    if not valid_results:
        print("[错误] 所有任务均失败，未生成输出。")
        sys.exit(1)

    if not output_path:
        output_dir = _ensure_output_dir(config_dir)
        base_name = os.path.splitext(os.path.basename(config_path))[0]
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"{base_name}_分析_{timestamp_str}.xlsx")
    else:
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    if os.path.exists(output_path):
        print(f"    [警告] 输出文件已存在，将被覆盖: {os.path.basename(output_path)}")
    print(f"[Pivot/3] 输出结果: {output_path}")
    valid_tasks = [t for t, r in zip(tasks, results) if r is not None]
    write_results(valid_tasks, valid_results, errors, output_path)

    print(f"\n[OK] 完成！分析结果已保存至: {output_path}")
    print(f"   共 {len(tasks)} 个任务: {len(valid_tasks)} 个成功" + (f", {skipped} 个跳过" if skipped else "") + (f", {len(errors)} 个失败" if errors else ""))
    if errors:
        print(f"   失败详情见「错误信息」Sheet")
    return output_path


def _run_html_mode(config_path, output_path=None, pivot_file=None, ppt_file=None):
    """
    生成HTML报告。
    :param config_path: 配置文件路径
    :param output_path: 输出HTML路径
    :param pivot_file: 透视分析结果文件路径
    :param ppt_file: PPT文件路径（可选）
    """
    from src.html_builder import generate_html_report, start_preview_server
    from src.excel_reader import read_ppt_config

    config_dir = os.path.dirname(config_path)
    print(f"[HTML/1] 读取配置: {config_path}")

    ppt_pages = []
    try:
        ppt_pages = read_ppt_config(config_path)
        print(f"    → PPT 配置: {len(ppt_pages)} 页")
    except Exception as e:
        print(f"    [警告] 读取 PPT 配置失败: {e}")

    pivot_excel = None
    if pivot_file:
        if os.path.isabs(pivot_file):
            pivot_excel = pivot_file
        else:
            pivot_excel = os.path.join(config_dir, pivot_file)
        if os.path.exists(pivot_excel):
            print(f"    → 透视结果文件: {os.path.basename(pivot_excel)}")
        else:
            print(f"    [警告] 指定的透视结果文件不存在: {pivot_excel}")
            pivot_excel = None

    if not pivot_excel:
        print(f"[HTML/2] 自动查找透视分析结果...")
        output_dirs = glob.glob(os.path.join(config_dir, "output_*"))
        if output_dirs:
            latest_dir = max(output_dirs, key=os.path.getmtime)
            excel_files = glob.glob(os.path.join(latest_dir, "*_分析_*.xlsx"))
            if excel_files:
                pivot_excel = excel_files[0]
                print(f"    → 找到透视结果: {os.path.basename(pivot_excel)}")
            else:
                excel_files = glob.glob(os.path.join(latest_dir, "*.xlsx"))
                if excel_files:
                    pivot_excel = excel_files[0]
                    print(f"    → 使用最新Excel: {os.path.basename(pivot_excel)}")

    if not pivot_excel:
        print(f"    未找到透视结果，自动运行透视分析...")
        pivot_excel = _run_pivot_mode(config_path)

    ppt_path = None
    if ppt_file:
        if os.path.isabs(ppt_file):
            ppt_path = ppt_file
        else:
            ppt_path = os.path.join(config_dir, ppt_file)
        if os.path.exists(ppt_path):
            print(f"    → PPT文件: {os.path.basename(ppt_path)}")
        else:
            print(f"    [警告] 指定的PPT文件不存在: {ppt_path}")
            ppt_path = None

    if not output_path:
        out_dir = os.path.dirname(pivot_excel) if pivot_excel else config_dir
        os.makedirs(out_dir, exist_ok=True)
        base_name = os.path.splitext(os.path.basename(config_path))[0]
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(out_dir, f"{base_name}_报告_{timestamp_str}.html")

    print(f"[HTML/3] 生成HTML报告...")
    html_path = generate_html_report(
        excel_path=pivot_excel,
        ppt_path=ppt_path,
        ppt_config=ppt_pages if ppt_pages else None,
        output_dir=os.path.dirname(output_path),
        report_title=f"{os.path.splitext(os.path.basename(config_path))[0]} - 数据分析报告",
        report_subtitle="透视分析结果",
    )

    if html_path:
        print(f"\n[OK] 完成！HTML报告已保存至: {html_path}")
        print(f"[HTML] 启动预览服务器...")
        start_preview_server(html_path)
        print(f"\n[提示] 报告包含：")
        print(f"  • 目录导航（点击跳转）")
        print(f"  • 数据摘要卡片")
        print(f"  • 图表切换（柱状图/折线图/饼图/表格）")
        print(f"  • 数据详情表格")
        if ppt_path:
            print(f"  • PPT页面预览（点击缩放）")
    else:
        print(f"[错误] HTML报告生成失败")
        sys.exit(1)


class _GuiLogRedirector:
    """重定向 stdout 到 GUI Text 控件"""
    def __init__(self, text_widget, tag=""):
        self.text_widget = text_widget
        self.tag = tag
        self._lock = False

    def write(self, s):
        if self._lock or not s:
            return
        self._lock = True
        try:
            self.text_widget.after(0, lambda: self._append(s))
        finally:
            self._lock = False

    def _append(self, s):
        self.text_widget.insert("end", s, self.tag)
        self.text_widget.see("end")

    def flush(self):
        pass


def _scan_files(folder):
    """扫描文件夹，返回检测到的配置和数据文件信息"""
    if not folder or not os.path.isdir(folder):
        return []

    xlsx_files = glob.glob(os.path.join(folder, "*.xlsx"))
    csv_files = glob.glob(os.path.join(folder, "*.csv"))
    all_data_files = [f for f in (xlsx_files + csv_files)
                      if not os.path.basename(f).startswith("~$")]

    results = []

    # 查找配置文件
    config_candidates = []
    for f in all_data_files:
        name = os.path.basename(f)
        if "配置" in name or "config" in name.lower():
            config_candidates.append(f)

    if not config_candidates and xlsx_files:
        config_candidates = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]

    import openpyxl
    ppt_keywords = {"页码", "页面类型", "页面标题", "图表类型"}
    pivot_keywords = {"数据源", "行维度", "列维度", "值字段", "聚合方式"}

    for cfg in config_candidates:
        name = os.path.basename(cfg)
        try:
            wb = openpyxl.load_workbook(cfg, read_only=True)
            ppt_found = False
            pivot_found = False
            sheet_names = wb.sheetnames
            for sn in sheet_names:
                ws = wb[sn]
                try:
                    row = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
                except StopIteration:
                    continue
                row_set = set(row)
                if len(ppt_keywords & row_set) >= 2:
                    ppt_found = True
                if len(pivot_keywords & row_set) >= 2:
                    pivot_found = True
            wb.close()

            if ppt_found and pivot_found:
                mode = "综合配置 (PPT + 透视)"
            elif pivot_found:
                mode = "透视分析配置"
            elif ppt_found:
                mode = "PPT报告配置"
            else:
                continue

            results.append({"type": "config", "name": name, "mode": mode, "path": cfg})
        except Exception:
            continue

    # 查找数据文件
    data_candidates = []
    for f in all_data_files:
        name = os.path.basename(f)
        if name.startswith("~$"):
            continue
        is_config = any("配置" in r["name"] or "config" in r["name"].lower() for r in results)
        if results and any(name == r["name"] for r in results):
            continue
        if "配置" in name or "config" in name.lower():
            continue
        data_candidates.append(f)

    for df in data_candidates:
        name = os.path.basename(df)
        ftype = "CSV" if name.lower().endswith(".csv") else "Excel"
        results.append({"type": "data", "name": name, "mode": f"数据文件 ({ftype})", "path": df})

    return results


def _draw_gradient_banner(canvas, height, top_hex, bottom_hex):
    """在 Canvas 上绘制从上到下的渐变矩形"""
    r1, g1, b1 = int(top_hex[1:3], 16), int(top_hex[3:5], 16), int(top_hex[5:7], 16)
    r2, g2, b2 = int(bottom_hex[1:3], 16), int(bottom_hex[3:5], 16), int(bottom_hex[5:7], 16)
    w = canvas.winfo_reqwidth() or 900
    for i in range(height):
        t = i / max(height - 1, 1)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        color = f"#{r:02x}{g:02x}{b:02x}"
        canvas.create_rectangle(0, i, w + 200, i + 1, fill=color, outline="")
    canvas.bind("<Configure>", lambda e: _redraw_gradient(canvas, height, top_hex, bottom_hex))


def _redraw_gradient(canvas, height, top_hex, bottom_hex):
    canvas.delete("all")
    _draw_gradient_banner_content(canvas, height, top_hex, bottom_hex)


def _draw_gradient_banner_content(canvas, height, top_hex, bottom_hex):
    w = canvas.winfo_width() or 900
    for i in range(height):
        r1, g1, b1 = int(top_hex[1:3], 16), int(top_hex[3:5], 16), int(top_hex[5:7], 16)
        r2, g2, b2 = int(bottom_hex[1:3], 16), int(bottom_hex[3:5], 16), int(bottom_hex[5:7], 16)
        t = i / max(height - 1, 1)
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        color = f"#{r:02x}{g:02x}{b:02x}"
        canvas.create_rectangle(0, i, w + 200, i + 1, fill=color, outline="")
    canvas.create_text(16, 18, text="Excel 统一分析工具", anchor="w",
                       font=("Microsoft YaHei", 15, "bold"), fill="#FFFFFF")
    canvas.create_text(16, 38, text=f"PPT 报告生成 + 透视分析  |  v{__VERSION__}  |  更新: {__UPDATE_DATE__}",
                       anchor="w", font=("Microsoft YaHei", 8), fill="#C0D4F0")


def _select_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog, ttk, scrolledtext
        import threading

        root = tk.Tk()
        root.title("Excel 统一分析工具")
        root.geometry("720x720")
        root.minsize(600, 560)
        root.configure(bg="#f7f8fa")

        # 样式
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f7f8fa")
        style.configure("Header.TLabel", font=("Microsoft YaHei", 16, "bold"), foreground="#182B49", background="#f7f8fa")
        style.configure("Sub.TLabel", font=("Microsoft YaHei", 11, "bold"), foreground="#182B49", background="#f7f8fa")
        style.configure("Info.TLabel", font=("Microsoft YaHei", 9), foreground="#666666", background="#f7f8fa")
        style.configure("Guide.TButton", font=("Microsoft YaHei", 10), padding=6)
        style.configure("Run.TButton", font=("Microsoft YaHei", 11, "bold"), padding=8)
        style.configure("Pick.TButton", font=("Microsoft YaHei", 10), padding=6)
        style.configure("Scan.TButton", font=("Microsoft YaHei", 10), padding=6)

        # 主容器
        main_frame = tk.Frame(root, bg="#f7f8fa", padx=16, pady=12)
        main_frame.pack(fill="both", expand=True)

        # 渐变色标题横幅（Canvas 绘制）
        header_canvas = tk.Canvas(main_frame, height=56, bg="#f7f8fa", highlightthickness=0)
        header_canvas.pack(fill="x", pady=(0, 8))
        _draw_gradient_banner(header_canvas, 56, "#182B49", "#2E75B6")
        header_canvas.create_text(16, 18, text="Excel 统一分析工具", anchor="w",
                                  font=("Microsoft YaHei", 15, "bold"), fill="#FFFFFF")
        header_canvas.create_text(16, 38, text=f"PPT 报告生成 + 透视分析  |  v{__VERSION__}  |  更新: {__UPDATE_DATE__}",
                                  anchor="w", font=("Microsoft YaHei", 8), fill="#C0D4F0")

        # 路径选择行（卡片内）
        path_outer = tk.Frame(main_frame, bg="#ffffff", highlightbackground="#e0e4ea", highlightthickness=1, padx=10, pady=10)
        path_outer.pack(fill="x", pady=(0, 8))
        path_frame = tk.Frame(path_outer, bg="#ffffff")
        path_frame.pack(fill="x")

        tk.Label(path_frame, text="📁 选择文件夹", font=("Microsoft YaHei", 11, "bold"),
                 fg="#182B49", bg="#ffffff").pack(anchor="w", pady=(0, 4))

        path_row = tk.Frame(path_frame, bg="#ffffff")
        path_row.pack(fill="x")
        path_var = tk.StringVar(value="")
        path_entry = tk.Entry(path_row, textvariable=path_var, font=("Microsoft YaHei", 10), state="readonly",
                              relief="solid", borderwidth=1, bg="#f5f7fa")
        path_entry.pack(side="left", fill="x", expand=True, ipady=4)

        ttk.Button(path_row, text="选择文件夹", style="Pick.TButton",
                   command=lambda: _browse_folder(path_var, file_list_text, root)).pack(side="left", padx=(8, 0))
        ttk.Button(path_row, text="扫描文件", style="Scan.TButton",
                   command=lambda: _scan_and_display(path_var.get(), file_list_text)).pack(side="left", padx=(6, 0))

        # 检测到的文件区域（卡片内）
        file_outer = tk.Frame(main_frame, bg="#ffffff", highlightbackground="#e0e4ea", highlightthickness=1, padx=10, pady=10)
        file_outer.pack(fill="x", pady=(0, 8))
        file_frame = tk.Frame(file_outer, bg="#ffffff")
        file_frame.pack(fill="x")
        tk.Label(file_frame, text="📄 检测到的文件", font=("Microsoft YaHei", 11, "bold"),
                 fg="#182B49", bg="#ffffff").pack(anchor="w", pady=(0, 4))

        file_list_text = tk.Text(file_frame, height=6, wrap="word", font=("Microsoft YaHei", 10),
                                  bg="#ffffff", fg="#333333", relief="solid", borderwidth=1)
        file_list_text.pack(fill="x", expand=True)

        # 配置文件选择区域
        config_select_frame = tk.Frame(file_frame, bg="#ffffff")
        config_combo_var = tk.StringVar()
        ttk.Label(config_select_frame, text="选择配置文件:", style="Info.TLabel").pack(side="left", padx=(0, 6))
        config_combo = ttk.Combobox(config_select_frame, textvariable=config_combo_var,
                                    state="readonly", width=45, font=("Microsoft YaHei", 10))
        config_combo.pack(side="left")
        file_list_text._config_select_frame = config_select_frame
        file_list_text._config_combo = config_combo
        file_list_text._config_combo_var = config_combo_var
        file_list_text._config_paths = []

        # 按钮行（放在日志区域之前 pack，确保始终可见）——卡片包裹
        btn_outer = tk.Frame(main_frame, bg="#ffffff", highlightbackground="#e0e4ea", highlightthickness=1, padx=10, pady=10)
        btn_outer.pack(fill="x", pady=(4, 0))
        tk.Label(btn_outer, text="🎯 操作", font=("Microsoft YaHei", 11, "bold"),
                 fg="#182B49", bg="#ffffff").pack(anchor="w", pady=(0, 6))
        btn_frame = tk.Frame(btn_outer, bg="#ffffff")
        btn_frame.pack(fill="x")

        run_state = {"running": False}

        def do_run():
            folder = path_var.get()
            if not folder:
                log_text.insert("end", "[错误] 请先选择文件夹\n")
                log_text.see("end")
                return
            if run_state["running"]:
                log_text.insert("end", "[警告] 分析正在运行中，请等待完成\n")
                log_text.see("end")
                return
            run_state["running"] = True
            run_btn.config(state="disabled", text="运行中...")
            log_text.delete("1.0", "end")

            # 读取选中的配置文件（多配置文件时由用户在下拉框选择）
            raw_args = [folder]
            selected_idx = getattr(file_list_text, "_config_combo", None)
            if selected_idx is not None:
                combo = file_list_text._config_combo
                idx = combo.current()
                paths = getattr(file_list_text, "_config_paths", [])
                if idx >= 0 and idx < len(paths):
                    raw_args = ["-c", paths[idx]]

            threading.Thread(target=_run_analysis_thread,
                             args=(raw_args, log_text, run_btn, run_state),
                             daemon=True).start()

        ttk.Button(btn_frame, text="操作指南", style="Guide.TButton",
                   command=_open_guide).pack(side="left")

        run_btn = ttk.Button(btn_frame, text="开始分析", style="Run.TButton", command=do_run)
        run_btn.pack(side="left", fill="x", expand=True, padx=(8, 0))

        ttk.Button(btn_frame, text="🌐 在浏览器中打开", style="Guide.TButton",
                   command=_open_web).pack(side="left", padx=(8, 0))

        ttk.Label(main_frame, text="提示：选择文件夹后点击「扫描文件」检测合规文件，再点击「开始分析」",
                  style="Info.TLabel").pack(anchor="w", pady=(8, 0))

        # 日志区域（放在最后 pack，expand=True 填满剩余空间）——卡片包裹
        log_outer = tk.Frame(main_frame, bg="#ffffff", highlightbackground="#e0e4ea", highlightthickness=1, padx=10, pady=10)
        log_outer.pack(fill="both", expand=True, pady=(8, 0))
        tk.Label(log_outer, text="📋 运行日志", font=("Microsoft YaHei", 11, "bold"),
                 fg="#182B49", bg="#ffffff").pack(anchor="w", pady=(0, 4))
        log_frame = tk.Frame(log_outer, bg="#ffffff")
        log_frame.pack(fill="both", expand=True)

        log_text = scrolledtext.ScrolledText(log_frame, wrap="word", font=("Consolas", 10),
                                              bg="#1d2b3a", fg="#c9d1d9", insertbackground="#fff",
                                              relief="solid", borderwidth=1)
        log_text.pack(fill="both", expand=True)

        # 尝试绑定拖拽（需要 tkinterdnd2，非必需）
        try:
            root.drop_target_register("DND_Files")
            root.dnd_bind("<<Drop>>", lambda e: _on_drop(e, path_var, file_list_text, root))
        except Exception:
            pass

        root.mainloop()
    except Exception as e:
        print(f"GUI 启动失败: {e}")
        import traceback
        traceback.print_exc()
        return None
    return None


def _browse_folder(path_var, file_list_text, root):
    from tkinter import filedialog
    folder = filedialog.askdirectory(title="选择包含配置 Excel 和数据文件的文件夹", parent=root)
    if folder:
        path_var.set(folder)
        _scan_and_display(folder, file_list_text)


def _scan_and_display(folder, file_list_text):
    file_list_text.delete("1.0", "end")
    if not folder:
        file_list_text.insert("end", "请先选择文件夹\n")
        return
    results = _scan_files(folder)
    if not results:
        file_list_text.insert("end", f"在 {folder} 中未检测到合规的配置或数据文件\n")
        file_list_text.insert("end", "提示：配置文件需要包含「页码」「数据源」等列头\n")
        return

    config_count = sum(1 for r in results if r["type"] == "config")
    data_count = sum(1 for r in results if r["type"] == "data")
    file_list_text.insert("end", f"检测到 {config_count} 个配置文件, {data_count} 个数据文件\n\n")
    for r in results:
        icon = "[配置]" if r["type"] == "config" else "[数据]"
        file_list_text.insert("end", f"{icon} {r['name']}  —  {r['mode']}\n")

    # 多配置文件时显示下拉选择框
    config_files = [r for r in results if r["type"] == "config"]
    config_frame = getattr(file_list_text, "_config_select_frame", None)
    config_combo = getattr(file_list_text, "_config_combo", None)
    if config_frame is not None and config_combo is not None:
        if len(config_files) > 1:
            file_list_text._config_paths = [r["path"] for r in config_files]
            config_combo["values"] = [r["name"] for r in config_files]
            config_combo.current(0)
            config_frame.pack(fill="x", pady=(6, 0))
        else:
            file_list_text._config_paths = []
            config_combo["values"] = []
            config_combo_var = getattr(file_list_text, "_config_combo_var", None)
            if config_combo_var:
                config_combo_var.set("")
            config_frame.pack_forget()


def _on_drop(event, path_var, file_list_text, root):
    """处理拖拽事件"""
    data = event.data
    if not data:
        return
    # 拖拽可能包含引号，去掉
    folder = data.strip().strip('"').strip("'")
    if os.path.isdir(folder):
        path_var.set(folder)
        _scan_and_display(folder, file_list_text)


def _open_guide():
    import webbrowser
    guide_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "guide.html")
    webbrowser.open(f"file:///{guide_path.replace(os.sep, '/')}")


def _open_web():
    import webbrowser, subprocess, threading
    base_dir = os.path.dirname(os.path.abspath(__file__))
    server_path = os.path.join(base_dir, "server.py")
    if not os.path.exists(server_path):
        log_text.insert("end", "[错误] Web 服务文件不存在\n")
        log_text.see("end")
        return

    log_text.insert("end", "🌐 正在启动 Web 服务...\n")
    log_text.see("end")

    def launch():
        subprocess.run([sys.executable, server_path, "--no-browser"], cwd=base_dir)
    threading.Thread(target=launch, daemon=True).start()

    def open_browser():
        import time
        time.sleep(1.5)
        webbrowser.open("http://localhost:8899")
        log_text.insert("end", "   → 浏览器已打开，访问 http://localhost:8899\n")
        log_text.see("end")
    threading.Thread(target=open_browser, daemon=True).start()


def _run_analysis_thread(raw_args, log_text, run_btn, run_state):
    """在子线程中运行分析，不阻塞 GUI"""
    import threading
    redirector = _GuiLogRedirector(log_text)
    old_stdout = sys.stdout
    sys.stdout = redirector

    try:
        _dispatch_from_gui(raw_args)
    except SystemExit:
        pass
    except Exception as e:
        print(f"\n[错误] 分析过程异常: {e}")
        import traceback
        traceback.print_exc()
    finally:
        sys.stdout = old_stdout
        run_state["running"] = False
        log_text.after(0, lambda: run_btn.config(state="normal", text="开始分析"))
        log_text.after(0, lambda: log_text.insert("end", "\n[完成] 分析结束，可选择其他文件夹再次运行\n"))
        log_text.after(0, lambda: log_text.see("end"))


def _dispatch_from_gui(raw_args):
    """从 GUI 调用的分析入口"""
    mode = "auto"
    legacy = argparse.ArgumentParser(add_help=False)
    legacy.add_argument("folder_or_config", nargs="?", default=None)
    legacy.add_argument("-c", "--config", default=None)
    legacy.add_argument("-o", "--output", default=None)
    args = legacy.parse_args(raw_args)
    config_path = _resolve_config_path(args)
    detected = _detect_mode(config_path)
    if detected == "all":
        print("[信息] 检测到综合配置（PPT + 透视分析），执行全部模式")
        config_dir = os.path.dirname(config_path)
        out_dir = _ensure_output_dir(config_dir)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.basename(config_path).rsplit(".", 1)[0]
        ppt_out = os.path.join(out_dir, f"{base}_报告_{ts}.pptx")
        pivot_out = os.path.join(out_dir, f"{base}_分析_{ts}.xlsx")
        _run_pivot_mode(config_path, pivot_out)
        print()
        _run_ppt_mode(config_path, ppt_out, pivot_data_file=pivot_out)
        return
    print(f"[信息] 自动检测配置类型: {detected}")
    if detected == "unknown":
        print("[错误] 未识别的配置类型，请确认 Excel 包含正确的配置 Sheet")
        sys.exit(1)
    if detected == "pivot":
        _run_pivot_mode(config_path, args.output)
    else:
        _run_ppt_mode(config_path, args.output)


def _run_template_mode(template_path, pivot_file=None, output_path=None):
    """基于 PPT 模板和透视结果进行数据替换"""
    from src.template_filler import fill_template

    if not os.path.exists(template_path):
        print(f"[错误] 模板文件不存在: {template_path}")
        sys.exit(1)

    template_dir = os.path.dirname(os.path.abspath(template_path))

    # 自动查找透视结果文件
    if not pivot_file:
        # 优先查找同目录下 _分析_ 命名的文件
        candidates = []
        for f in os.listdir(template_dir):
            if f.endswith(".xlsx") and "_分析_" in f and not f.startswith("~$"):
                candidates.append(os.path.join(template_dir, f))
        if len(candidates) == 1:
            pivot_file = candidates[0]
            print(f"[信息] 自动找到透视结果: {os.path.basename(pivot_file)}")
        elif len(candidates) > 1:
            print(f"[信息] 找到多个透视结果文件，请用 --pivot 指定:")
            for c in candidates:
                print(f"  - {os.path.basename(c)}")
            sys.exit(1)
        else:
            print("[错误] 未找到透视结果文件，请用 --pivot 指定")
            sys.exit(1)
    elif not os.path.exists(pivot_file):
        # 相对路径处理
        abs_pivot = os.path.join(template_dir, pivot_file) if not os.path.isabs(pivot_file) else pivot_file
        if os.path.exists(abs_pivot):
            pivot_file = abs_pivot
        else:
            print(f"[错误] 透视结果文件不存在: {pivot_file}")
            sys.exit(1)

    if not output_path:
        base_name = os.path.splitext(os.path.basename(template_path))[0]
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(template_dir, f"{base_name}_填充_{ts}.pptx")

    fill_template(template_path, pivot_file, output_path)


def main():
    if len(sys.argv) == 1:
        # 无参数：启动 GUI 模式
        _select_folder()
        return

    raw_args = sys.argv[1:]
    mode = None

    if raw_args and raw_args[0] in ("ppt", "pivot", "html", "template"):
        mode = raw_args[0]
        raw_args = raw_args[1:]
    elif raw_args and raw_args[0] in ("-h", "--help"):
        mode = "help"
    else:
        mode = "auto"

    if mode == "help":
        parser = argparse.ArgumentParser(description="Excel 统一分析工具 — PPT生成 & 透视分析 & HTML报告")
        sub = parser.add_subparsers(dest="mode", help="子命令")
        ppt_p = sub.add_parser("ppt", help="生成PPT报告")
        ppt_p.add_argument("folder_or_config", nargs="?", default=None)
        ppt_p.add_argument("-c", "--config", default=None)
        ppt_p.add_argument("-o", "--output", default=None)
        ppt_p.add_argument("--pivot-file", dest="pivot_file", default=None,
                           help="透视分析结果文件路径，用于 {pivot} 数据源引用")
        ppt_p.add_argument("--check", action="store_true", help="仅校验配置，不执行")
        pivot_p = sub.add_parser("pivot", help="透视分析")
        pivot_p.add_argument("folder_or_config", nargs="?", default=None)
        pivot_p.add_argument("-c", "--config", default=None)
        pivot_p.add_argument("-o", "--output", default=None)
        pivot_p.add_argument("--check", action="store_true", help="仅校验配置，不执行")
        html_p = sub.add_parser("html", help="生成HTML报告")
        html_p.add_argument("folder_or_config", nargs="?", default=None)
        html_p.add_argument("-c", "--config", default=None)
        html_p.add_argument("-o", "--output", default=None)
        html_p.add_argument("--pivot-file", dest="pivot_file", default=None,
                           help="透视分析结果文件路径")
        html_p.add_argument("--ppt-file", dest="ppt_file", default=None,
                           help="PPT文件路径（可选，用于预览截图）")
        tpl_p = sub.add_parser("template", help="基于PPT模板填充数据")
        tpl_p.add_argument("template_path", help="PPT模板文件路径 (.pptx)")
        tpl_p.add_argument("--pivot", dest="pivot_file", default=None,
                           help="透视分析结果文件路径 (.xlsx)")
        tpl_p.add_argument("-o", "--output", default=None, help="输出PPT路径")
        parser.parse_args()
        return

    legacy = argparse.ArgumentParser(add_help=False)
    legacy.add_argument("folder_or_config", nargs="?", default=None, help="文件夹路径或配置文件")
    legacy.add_argument("-c", "--config", default=None, help="配置文件路径")
    legacy.add_argument("-o", "--output", default=None, help="输出路径")
    legacy.add_argument("--pivot-file", dest="pivot_file", default=None,
                        help="透视分析结果文件路径，用于 {pivot} 数据源引用")
    legacy.add_argument("--check", action="store_true", help="仅校验配置，不执行")

    if mode == "auto":
        args = legacy.parse_args(raw_args)
        config_path = _resolve_config_path(args)
        detected = _detect_mode(config_path)
        if detected == "all":
            print("[信息] 检测到综合配置（PPT + 透视分析），执行全部模式")
            config_dir = os.path.dirname(config_path)
            if not args.output:
                out_dir = _ensure_output_dir(config_dir)
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                ppt_out = os.path.join(out_dir, f"{os.path.basename(config_path).rsplit('.',1)[0]}_报告_{ts}.pptx")
                pivot_out = os.path.join(out_dir, f"{os.path.basename(config_path).rsplit('.',1)[0]}_分析_{ts}.xlsx")
            else:
                out_dir = os.path.dirname(args.output) or "."
                os.makedirs(out_dir, exist_ok=True)
                base_name = os.path.basename(args.output).rsplit(".", 1)[0]
                ppt_out = os.path.join(out_dir, f"{base_name}.pptx")
                pivot_out = os.path.join(out_dir, f"{base_name}_分析.xlsx")
            _run_pivot_mode(config_path, pivot_out, validate_only=getattr(args, 'check', False))
            if not getattr(args, 'check', False):
                print()
                _run_ppt_mode(config_path, ppt_out, pivot_data_file=pivot_out)
                # 自动生成 HTML 报告（供手机预览）
                print()
                try:
                    _run_html_mode(config_path, None, pivot_file=pivot_out, ppt_file=ppt_out)
                except Exception as e:
                    print(f"[警告] HTML 报告生成失败: {e}")
            return
        print(f"[信息] 自动检测配置类型: {detected}")
        if detected == "unknown":
            print("[错误] 未识别的配置类型，请确认 Excel 包含正确的配置 Sheet")
            sys.exit(1)
        mode = detected
    else:
        args = legacy.parse_args(raw_args)

    # ===== template 模式：独立处理，不走 _resolve_config_path =====
    if mode == "template":
        tpl_parser = argparse.ArgumentParser(add_help=False)
        tpl_parser.add_argument("template_path", help="PPT模板文件路径")
        tpl_parser.add_argument("--pivot", dest="pivot_file", default=None)
        tpl_parser.add_argument("-o", "--output", default=None)
        tpl_args = tpl_parser.parse_args(raw_args)
        _run_template_mode(tpl_args.template_path, tpl_args.pivot_file, tpl_args.output)
        return

    config_path = _resolve_config_path(args)
    output_path = args.output
    validate_only = getattr(args, 'check', False)

    if mode == "pivot":
        _run_pivot_mode(config_path, output_path, validate_only=validate_only)
    elif mode == "html":
        _run_html_mode(config_path, output_path,
                       pivot_file=getattr(args, "pivot_file", None),
                       ppt_file=getattr(args, "ppt_file", None))
    else:
        pivot_file = getattr(args, "pivot_file", None)
        _run_ppt_mode(config_path, output_path, pivot_data_file=pivot_file, validate_only=validate_only)


if __name__ == "__main__":
    main()
