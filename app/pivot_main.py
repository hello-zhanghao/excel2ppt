"""
Excel 透视分析工具（独立入口）
用法：
  python pivot_main.py 文件夹路径      ← 自动找配置和数据
  python pivot_main.py -c 配置.xlsx    ← 指定配置文件
  python pivot_main.py -c 配置.xlsx -o 结果.xlsx
"""
import os
import sys
import argparse
import glob
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.pivot_analyzer import read_pivot_config, run_analysis, _auto_find_data_files
from src.excel_writer import write_results


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


def main():
    parser = argparse.ArgumentParser(description="Excel透视分析")
    parser.add_argument("folder_or_config", nargs="?", default=None,
                        help="文件夹或配置文件路径")
    parser.add_argument("-c", "--config", default=None, help="配置Excel文件路径")
    parser.add_argument("-o", "--output", default=None, help="输出Excel路径")
    args = parser.parse_args()

    config_path = args.config

    if not config_path and args.folder_or_config:
        if os.path.isdir(args.folder_or_config):
            config_path = find_config_file(args.folder_or_config)
            if not config_path:
                print(f"[错误] 在 {args.folder_or_config} 里没找到 Excel 文件")
                sys.exit(1)
            print(f"[信息] 自动找到配置文件: {os.path.basename(config_path)}")
        elif os.path.isfile(args.folder_or_config):
            config_path = args.folder_or_config

    if not config_path:
        print("[错误] 请指定文件夹路径或配置文件路径")
        print("  用法1: python pivot_main.py cases/pivot_demo")
        print("  用法2: python pivot_main.py -c 配置.xlsx")
        sys.exit(1)

    if not os.path.exists(config_path):
        print(f"[错误] 配置文件不存在: {config_path}")
        sys.exit(1)

    config_path = os.path.abspath(config_path)
    config_dir = os.path.dirname(config_path)

    print(f"[1/3] 读取配置: {config_path}")
    tasks = read_pivot_config(config_path)
    print(f"    → 共 {len(tasks)} 个分析任务")

    if not tasks:
        print("[错误] 没有有效的分析任务。")
        sys.exit(1)

    data_files = _auto_find_data_files(config_dir, config_path)
    if data_files:
        print(f"    → 找到数据文件: {[os.path.basename(f) for f in data_files]}")

    print(f"[2/3] 执行透视分析...")
    results = []
    errors = []
    scalar_context = {}
    block_results = {}

    for task in tasks:
        seq = task.get("序号", "?")
        结果Sheet = task.get("结果Sheet", f"结果{seq}")
        try:
            result, error = run_analysis(task, config_dir, scalar_context, block_results)
            if error:
                print(f"    ✗ [任务{seq}] {error}")
                errors.append({"序号": seq, "错误": error})
                results.append(None)
            else:
                if isinstance(result, dict):
                    for key, df in result.items():
                        if hasattr(df, "shape"):
                            print(f"    ✓ [任务{seq}] {结果Sheet} → {df.shape[0]}行 x {df.shape[1]}列")
                        else:
                            print(f"    ✓ [任务{seq}] {结果Sheet} (标量)")
                else:
                    print(f"    ✓ [任务{seq}] {结果Sheet}")
                results.append(result)

                block_name = task.get("区块名", "") or 结果Sheet
                if isinstance(result, dict):
                    for key, df in result.items():
                        if hasattr(df, "shape") and not key.startswith("_JOIN中间表_"):
                            block_results[block_name] = df
                            break

                from src.pivot_analyzer import collect_task_scalars
                task_scalars = collect_task_scalars(result)
                if task_scalars:
                    scalar_context.update(task_scalars)
        except Exception as e:
            print(f"    ✗ [任务{seq}] 异常: {e}")
            errors.append({"序号": seq, "错误": str(e)})
            results.append(None)

    valid_results = [r for r in results if r is not None]
    if not valid_results:
        print("[错误] 所有任务均失败，未生成输出。")
        sys.exit(1)

    if args.output:
        output_path = args.output
    else:
        base_name = os.path.splitext(os.path.basename(config_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(config_dir, f"{base_name}_分析结果_{timestamp}.xlsx")

    print(f"[3/3] 输出结果: {output_path}")
    valid_tasks = [t for t, r in zip(tasks, results) if r is not None]
    write_results(valid_tasks, valid_results, errors, output_path)

    print(f"\n✅ 完成！分析结果已保存至: {output_path}")
    print(f"   共 {len(valid_tasks)} 个任务成功" + (f", {len(errors)} 个失败" if errors else ""))

    if errors:
        print(f"   失败详情见「错误信息」Sheet")


if __name__ == "__main__":
    main()
