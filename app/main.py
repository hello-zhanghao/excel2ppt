"""
Excel 统一分析工具 — PPT 生成 + 透视分析
用法：
  python main.py ppt 文件夹路径           ← PPT 生成（自动找配置和数据）
  python main.py ppt -c 配置.xlsx         ← PPT 生成（指定配置）
  python main.py ppt -c 配置.xlsx -o out.pptx
  python main.py pivot 文件夹路径         ← 透视分析（自动找配置和数据）
  python main.py pivot -c 配置.xlsx       ← 透视分析（指定配置）
  python main.py pivot -c 配置.xlsx -o out.xlsx
  python main.py 文件夹路径               ← 自动检测配置类型，分发到对应模式
"""
import os
import sys
import argparse
import glob
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def find_config_file(folder):
    xlsx_files = glob.glob(os.path.join(folder, "*.xlsx"))
    xlsx_files = [f for f in xlsx_files if not os.path.basename(f).startswith("~$")]
    if not xlsx_files:
        return None
    for f in xlsx_files:
        name = os.path.basename(f)
        if "配置" in name or "config" in name.lower():
            return f
    return xlsx_files[0]


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
                print(f"[信息] 自动找到配置文件: {os.path.basename(config_path)}")
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
    wb = openpyxl.load_workbook(config_path, read_only=True)
    ppt_keywords = {"页码", "页面类型", "页面标题", "图表类型"}
    pivot_keywords = {"数据源", "行维度", "列维度", "值字段", "聚合方式"}
    ppt_found = False
    pivot_found = False
    for name in wb.sheetnames:
        ws = wb[name]
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
        return "all"
    if pivot_found:
        return "pivot"
    return "ppt"


def _run_ppt_mode(config_path, output_path=None):
    from src.excel_reader import (
        read_config, read_data, read_geo_data, find_data_file, get_data_file_info
    )
    from src.ppt_builder import build_ppt

    config_dir = os.path.dirname(config_path)
    print(f"[PPT/1] 读取配置: {config_path}")
    config = read_config(config_path)
    general = config.get("general", {})
    pages = config.get("pages", [])
    colors = config.get("colors", {})
    print(f"    → PPT页面: {len(pages)} 页")

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
    chart_map = {}

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
                        chart_map[chart_title] = chart_def
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
                        chart_def["_hierarchical"] = {}
                    chart_map[chart_title] = chart_def
                    total_charts += 1
                    print(f"    [OK] [{page_title}] {chart_title}")
                else:
                    print(f"    - [{page_title}] {chart_title} (无数据)")
            except Exception as e:
                print(f"    ! [{page_title}] {chart_title} (错误: {e})")

    print(f"    → 共 {total_charts} 个图表")
    if total_charts == 0 and not pages:
        print("[错误] 没有有效内容。")
        sys.exit(1)

    if not output_path:
        output_dir = _ensure_output_dir(config_dir)
        base_name = os.path.splitext(os.path.basename(config_path))[0]
        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_dir, f"{base_name}_报告_{timestamp_str}.pptx")
    else:
        output_dir = os.path.dirname(output_path) or "."
        os.makedirs(output_dir, exist_ok=True)

    print(f"[PPT/3] 生成PPT: {output_path}")
    build_ppt(config, chart_map, output_path)
    print(f"\n[OK] 完成！PPT已保存至: {output_path}")
    print(f"   共 {len(pages)} 页, {total_charts} 个原生图表")
    print(f"   提示: 双击PPT中的图表可查看和编辑数据")


def _run_pivot_mode(config_path, output_path=None):
    from src.pivot_analyzer import read_pivot_config, run_analysis, _auto_find_data_files
    from src.excel_writer import write_results

    config_dir = os.path.dirname(config_path)
    print(f"[Pivot/1] 读取配置: {config_path}")
    tasks = read_pivot_config(config_path)
    print(f"    → 共 {len(tasks)} 个分析任务")

    if not tasks:
        print("[错误] 没有有效的分析任务。")
        sys.exit(1)

    data_files = _auto_find_data_files(config_dir, config_path)
    if data_files:
        print(f"    → 找到数据文件: {[os.path.basename(f) for f in data_files]}")

    print(f"[Pivot/2] 执行透视分析...")
    results = []
    errors = []

    for task in tasks:
        seq = task.get("序号", "?")
        sheet_name = task.get("结果Sheet", f"结果{seq}")
        try:
            result, error = run_analysis(task, config_dir)
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

    print(f"[Pivot/3] 输出结果: {output_path}")
    valid_tasks = [t for t, r in zip(tasks, results) if r is not None]
    write_results(valid_tasks, valid_results, errors, output_path)

    print(f"\n[OK] 完成！分析结果已保存至: {output_path}")
    print(f"   共 {len(valid_tasks)} 个任务成功" + (f", {len(errors)} 个失败" if errors else ""))
    if errors:
        print(f"   失败详情见「错误信息」Sheet")


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


def _select_folder():
    try:
        import tkinter as tk
        from tkinter import filedialog, ttk, scrolledtext
        import threading

        root = tk.Tk()
        root.title("Excel 统一分析工具")
        root.geometry("720x680")
        root.minsize(600, 500)
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
        main_frame = ttk.Frame(root, padding="16 12 16 12")
        main_frame.pack(fill="both", expand=True)

        # 标题
        ttk.Label(main_frame, text="Excel 统一分析工具", style="Header.TLabel").pack(anchor="w", pady=(0, 2))
        ttk.Label(main_frame, text="PPT 报告生成 + 透视分析", style="Info.TLabel").pack(anchor="w", pady=(0, 12))

        # 路径选择行
        path_frame = ttk.Frame(main_frame)
        path_frame.pack(fill="x", pady=(0, 8))

        path_var = tk.StringVar(value="")
        path_entry = ttk.Entry(path_frame, textvariable=path_var, font=("Microsoft YaHei", 10), state="readonly")
        path_entry.pack(side="left", fill="x", expand=True, ipady=2)

        ttk.Button(path_frame, text="选择文件夹", style="Pick.TButton",
                   command=lambda: _browse_folder(path_var, file_list_text, root)).pack(side="left", padx=(8, 0))
        ttk.Button(path_frame, text="扫描文件", style="Scan.TButton",
                   command=lambda: _scan_and_display(path_var.get(), file_list_text)).pack(side="left", padx=(6, 0))

        # 检测到的文件区域
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill="x", pady=(4, 8))
        ttk.Label(file_frame, text="检测到的文件", style="Sub.TLabel").pack(anchor="w", pady=(0, 4))

        file_list_text = tk.Text(file_frame, height=6, wrap="word", font=("Microsoft YaHei", 10),
                                  bg="#ffffff", fg="#333333", relief="solid", borderwidth=1)
        file_list_text.pack(fill="x", expand=True)

        # 日志区域
        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill="both", expand=True, pady=(4, 8))
        ttk.Label(log_frame, text="运行日志", style="Sub.TLabel").pack(anchor="w", pady=(0, 4))

        log_text = scrolledtext.ScrolledText(log_frame, wrap="word", font=("Consolas", 10),
                                              bg="#1d2b3a", fg="#c9d1d9", insertbackground="#fff",
                                              relief="solid", borderwidth=1)
        log_text.pack(fill="both", expand=True)

        # 按钮行
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill="x", pady=(4, 0))

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
            threading.Thread(target=_run_analysis_thread,
                             args=(folder, log_text, run_btn, run_state),
                             daemon=True).start()

        ttk.Button(btn_frame, text="操作指南", style="Guide.TButton",
                   command=_open_guide).pack(side="left")

        run_btn = ttk.Button(btn_frame, text="开始分析", style="Run.TButton", command=do_run)
        run_btn.pack(side="left", fill="x", expand=True, padx=(8, 0))

        ttk.Label(main_frame, text="提示：选择文件夹后点击「扫描文件」检测合规文件，再点击「开始分析」",
                  style="Info.TLabel").pack(anchor="w", pady=(8, 0))

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


def _run_analysis_thread(folder, log_text, run_btn, run_state):
    """在子线程中运行分析，不阻塞 GUI"""
    import threading
    redirector = _GuiLogRedirector(log_text)
    old_stdout = sys.stdout
    sys.stdout = redirector

    try:
        raw_args = [folder]
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
        _run_ppt_mode(config_path, ppt_out)
        print()
        _run_pivot_mode(config_path, pivot_out)
        return
    print(f"[信息] 自动检测配置类型: {detected}")
    if detected == "pivot":
        _run_pivot_mode(config_path, args.output)
    else:
        _run_ppt_mode(config_path, args.output)


def main():
    if len(sys.argv) == 1:
        # 无参数：启动 GUI 模式
        _select_folder()
        return

    raw_args = sys.argv[1:]
    mode = None

    if raw_args and raw_args[0] in ("ppt", "pivot"):
        mode = raw_args[0]
        raw_args = raw_args[1:]
    elif raw_args and raw_args[0] in ("-h", "--help"):
        mode = "help"
    else:
        mode = "auto"

    if mode == "help":
        parser = argparse.ArgumentParser(description="Excel 统一分析工具 — PPT生成 & 透视分析")
        sub = parser.add_subparsers(dest="mode", help="子命令")
        ppt_p = sub.add_parser("ppt", help="生成PPT报告")
        ppt_p.add_argument("folder_or_config", nargs="?", default=None)
        ppt_p.add_argument("-c", "--config", default=None)
        ppt_p.add_argument("-o", "--output", default=None)
        pivot_p = sub.add_parser("pivot", help="透视分析")
        pivot_p.add_argument("folder_or_config", nargs="?", default=None)
        pivot_p.add_argument("-c", "--config", default=None)
        pivot_p.add_argument("-o", "--output", default=None)
        parser.parse_args()
        return

    legacy = argparse.ArgumentParser(add_help=False)
    legacy.add_argument("folder_or_config", nargs="?", default=None, help="文件夹路径或配置文件")
    legacy.add_argument("-c", "--config", default=None, help="配置文件路径")
    legacy.add_argument("-o", "--output", default=None, help="输出路径")

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
                ppt_out = args.output
                pivot_out = args.output.replace(".pptx", ".xlsx").replace("_报告_", "_分析_")
            _run_ppt_mode(config_path, ppt_out)
            print()
            _run_pivot_mode(config_path, pivot_out)
            return
        print(f"[信息] 自动检测配置类型: {detected}")
        mode = detected
    else:
        args = legacy.parse_args(raw_args)

    config_path = _resolve_config_path(args)
    output_path = args.output

    if mode == "pivot":
        _run_pivot_mode(config_path, output_path)
    else:
        _run_ppt_mode(config_path, output_path)


if __name__ == "__main__":
    main()
