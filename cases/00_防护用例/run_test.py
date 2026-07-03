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


def verify_ppt_features(ppt_path):
    """验证 P0/P1 新功能是否正确实现"""
    print_header("P0/P1 新功能验证")
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    checks = []
    prs = Presentation(ppt_path)
    slides = list(prs.slides)

    # 检查页数（期望11页）
    checks.append(("总页数=11", len(slides) == 11, f"实际{len(slides)}页"))

    def get_texts(slide):
        return [s.text_frame.text for s in slide.shapes if s.has_text_frame]

    # P1: 目录页（第2页）
    if len(slides) >= 2:
        texts = get_texts(slides[1])
        has_toc = any("目录" in t or "CONTENTS" in t for t in texts)
        checks.append(("P1 目录页", has_toc, f"文本: {texts[:3]}"))
    else:
        checks.append(("P1 目录页", False, "页数不足"))

    # P1: 章节页（第3页，深色背景+居中标题）
    if len(slides) >= 3:
        texts = get_texts(slides[2])
        has_section = any("销售汇总" in t for t in texts)
        checks.append(("P1 章节页", has_section, f"文本: {texts[:2]}"))
    else:
        checks.append(("P1 章节页", False, "页数不足"))

    # P0: 左图右文文字区（第7页，应含多行文字内容）
    if len(slides) >= 7:
        texts = get_texts(slides[6])
        has_text = any("透视分析" in t or "图表类型" in t for t in texts)
        checks.append(("P0 左图右文文字区", has_text, f"文本块数: {len(texts)}"))
    else:
        checks.append(("P0 左图右文文字区", False, "页数不足"))

    # P1: 长标签X轴旋转（第10页，8个城市>6触发旋转）
    if len(slides) >= 10:
        has_chart = False
        for shape in slides[9].shapes:
            if shape.has_chart:
                has_chart = True
                try:
                    cats = list(shape.chart.plots[0].categories)
                    checks.append(("P1 长标签旋转(>6类别)", len(cats) > 6, f"{len(cats)}个类别"))
                except Exception as e:
                    checks.append(("P1 长标签旋转(>6类别)", False, f"异常: {e}"))
                break
        if not has_chart:
            checks.append(("P1 长标签旋转(>6类别)", False, "无图表"))
    else:
        checks.append(("P1 长标签旋转(>6类别)", False, "页数不足"))

    # P1: 结论卡片（第4页，应含✦标记的独立卡片）
    if len(slides) >= 4:
        texts = get_texts(slides[3])
        has_conclusion = any("✦" in t for t in texts)
        # 检查是否有圆角矩形背景框
        has_card_bg = any(
            s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
            and hasattr(s, "adjustments")
            and len(s.adjustments) > 0
            for s in slides[3].shapes
        )
        checks.append(("P1 结论卡片", has_conclusion, "✦标记" + ("+背景框" if has_card_bg else "无背景框")))
    else:
        checks.append(("P1 结论卡片", False, "页数不足"))

    # P1: 结尾页（第11页）
    if len(slides) >= 11:
        texts = get_texts(slides[10])
        has_ending = any("谢谢" in t for t in texts)
        checks.append(("P1 结尾页", has_ending, f"文本: {texts[:2]}"))
    else:
        checks.append(("P1 结尾页", False, "页数不足"))

    # 是否生成=否 的页面被跳过（配置12页，实际11页，"跳过测试页"不应出现）
    all_texts = [t for slide in slides for t in get_texts(slide)]
    no_skip_page = not any("跳过测试页" in t or "不应出现" in t for t in all_texts)
    checks.append(("是否生成跳过(页12)", no_skip_page and len(slides) == 11, f"配置12页→实际{len(slides)}页"))

    # P0: 百分比饼图数据为 0~1 小数（第5页饼图，引用占比透视结果）
    # 验证 PPT 拿到的是 0~1 小数，未被 Excel 的 0.0% 格式 ×100 影响
    if len(slides) >= 5:
        pct_ok = False
        pct_detail = "无饼图"
        from pptx.enum.chart import XL_CHART_TYPE
        for shape in slides[4].shapes:
            if shape.has_chart:
                chart = shape.chart
                try:
                    if chart.chart_type == XL_CHART_TYPE.PIE:
                        plot = chart.plots[0]
                        vals = list(plot.series[0].values)
                        # 饼图占比值应为 0~1 小数（如 0.376）
                        pct_ok = all(isinstance(v, (int, float)) and 0 <= v <= 1 for v in vals)
                        pct_detail = f"饼图值={vals}"
                        break
                except Exception as e:
                    pct_detail = f"异常: {e}"
        checks.append(("P0 百分比饼图0~1小数", pct_ok, pct_detail))

    # 打印结果
    all_ok = True
    for name, ok, detail in checks:
        icon = "✓" if ok else "✗"
        color = GREEN if ok else RED
        print(f"  {color}{icon} {name}: {detail}{RESET}")
        if not ok:
            all_ok = False

    return all_ok


def verify_excel_output(excel_path):
    """验证透视分析 Excel 输出的完整性和数据正确性"""
    print_header("Excel 输出验证")
    checks = []

    if not excel_path or not os.path.exists(excel_path):
        print(f"  {RED}✗ Excel 文件不存在{RESET}")
        return False

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(f"  {RED}✗ 无法打开 Excel: {e}{RESET}")
        return False

    # 1. 检查 Sheet 数量（期望 >=10）
    expected_sheets = [
        "按地区汇总", "地区多维统计", "地区产品交叉", "地区占比",
        "产品计数占比", "销售额分箱", "华东华北汇总", "产品系列汇总",
        "地区单价分析", "block合并测试",
    ]
    actual_sheets = wb.sheetnames
    checks.append(("Sheet数量>=10", len(actual_sheets) >= 10, f"实际{len(actual_sheets)}个"))

    # 2. 检查关键 Sheet 是否存在
    for sname in expected_sheets:
        exists = sname in actual_sheets
        checks.append((f"Sheet[{sname}]存在", exists, ""))

    # 2.5 检查区块名是否使用配置的值（首行=区块名）
    expected_titles = {
        "按地区汇总": "简单分组求和",
        "地区多维统计": "多字段多聚合",
        "地区产品交叉": "交叉表",
        "销售额分箱": "分箱统计",
    }
    for sname, expected_title in expected_titles.items():
        if sname in actual_sheets:
            ws = wb[sname]
            first_cell = ws.cell(row=1, column=1).value
            checks.append((f"区块名[{sname}]", first_cell == expected_title, f"期望'{expected_title}', 实际'{first_cell}'"))

    def _read_sheet_data(sname):
        """读取 Sheet 数据（跳过区块标题行，返回表头+数据行）"""
        ws = wb[sname]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return [], []
        # 第1行是区块标题，第2行是表头
        header = [str(c).strip() if c is not None else "" for c in rows[1]]
        data = rows[2:]
        return header, data

    # 3. 按地区汇总 - 验证数据值
    if "按地区汇总" in actual_sheets:
        header, data = _read_sheet_data("按地区汇总")
        checks.append(("按地区汇总表头", "地区" in header and "总销售额" in header, str(header)))
        # 华东=4800, 华北=3300, 华南=4650
        vals = {str(r[0]).strip(): r[1] for r in data if r[0]}
        checks.append(("华东销售额=4800", vals.get("华东") == 4800, str(vals.get("华东"))))
        checks.append(("华北销售额=3300", vals.get("华北") == 3300, str(vals.get("华北"))))
        checks.append(("华南销售额=4650", vals.get("华南") == 4650, str(vals.get("华南"))))

    # 4. 地区多维统计 - 验证多聚合字段
    if "地区多维统计" in actual_sheets:
        header, data = _read_sheet_data("地区多维统计")
        expected_cols = ["地区", "总销售额", "平均销售额", "客户数"]
        has_all_cols = all(c in header for c in expected_cols)
        checks.append(("地区多维统计4列", has_all_cols, str(header)))
        # 华东: 总额4800, 均值1200, 客户数4
        vals = {str(r[0]).strip(): r for r in data if r[0]}
        hd = vals.get("华东")
        if hd:
            checks.append(("华东客户数=4", hd[3] == 4, str(hd[3])))
            checks.append(("华东均值=1200", hd[2] == 1200, str(hd[2])))

    # 5. 地区产品交叉 - 验证交叉表含合计列
    if "地区产品交叉" in actual_sheets:
        header, data = _read_sheet_data("地区产品交叉")
        has_total = "合计" in header
        checks.append(("交叉表含合计列", has_total, str(header)))
        # 华东合计=4800
        vals = {str(r[0]).strip(): r for r in data if r[0]}
        hd_row = vals.get("华东")
        if hd_row and has_total:
            total_idx = header.index("合计")
            checks.append(("华东合计=4800", hd_row[total_idx] == 4800, str(hd_row[total_idx])))

    # 6. 地区占比 - 验证占比值合理（0~1 小数，由 0.0% 格式自动 ×100 显示）
    if "地区占比" in actual_sheets:
        header, data = _read_sheet_data("地区占比")
        pct_vals = [r[1] for r in data if r[0] and r[1] is not None]
        all_valid = all(isinstance(v, (int, float)) and 0 <= v <= 1 for v in pct_vals)
        checks.append(("占比值0~1小数", all_valid, str(pct_vals)))
        # 验证 Excel 单元格 number_format 是百分比格式
        ws_pct = wb["地区占比"]
        # 第3行第2列是第一个占比数据单元格（第1行区块标题，第2行表头）
        pct_cell = ws_pct.cell(row=3, column=2)
        fmt_ok = "%" in str(pct_cell.number_format)
        checks.append(("占比列number_format含%", fmt_ok, f"format={pct_cell.number_format}"))
        # 验证存储值是 0.376 而非 37.6（确认未被 ×100）
        store_ok = pct_cell.value is not None and 0 < float(pct_cell.value) <= 1
        checks.append(("占比存储值未×100", store_ok, f"存储值={pct_cell.value}"))

    # 6.1 产品计数占比 - 验证 count_pct 也走 0~1 小数
    if "产品计数占比" in actual_sheets:
        header, data = _read_sheet_data("产品计数占比")
        pct_vals = [r[1] for r in data if r[0] and r[1] is not None]
        all_valid = all(isinstance(v, (int, float)) and 0 <= v <= 1 for v in pct_vals)
        checks.append(("计数占比0~1小数", all_valid, str(pct_vals)))

    # 7. 销售额分箱 - 验证分箱区间存在
    if "销售额分箱" in actual_sheets:
        header, data = _read_sheet_data("销售额分箱")
        has_bins = len(data) >= 3 and "~" in str(data[0][0])
        checks.append(("分箱区间存在", has_bins, str([r[0] for r in data[:3]])))

    # 8. block合并测试 - 验证多列合并
    if "block合并测试" in actual_sheets:
        header, data = _read_sheet_data("block合并测试")
        has_multi_col = "客户数_A" in header and "客户数_B" in header
        checks.append(("block合并双列", has_multi_col, str(header)))

    # 9. 无行维度汇总 - 横向一行输出（位置1:1对应）
    if "无行维度汇总" in actual_sheets:
        header, data = _read_sheet_data("无行维度汇总")
        # 验证列名：位置1:1 → 总销售额_sum, 平均客户数_avg
        has_col_sum = "总销售额" in header or any("销售额" in h and "求和" in h for h in header)
        has_col_avg = "平均客户数" in header or any("客户数" in h and ("均值" in h or "avg" in h) for h in header)
        # 验证只有1行数据（横向一行）
        is_single_row = len(data) == 1
        checks.append(("无行维度→横向一行", is_single_row, f"data行数={len(data)}"))
        checks.append(("无行维度→列名匹配", has_col_sum and has_col_avg, f"header={header}"))

    # 10. 区块名合并：task1(按地区汇总) + task14(按地区汇总, 同名"简单分组求和") 应合并
    if "按地区汇总" in actual_sheets:
        header, data = _read_sheet_data("按地区汇总")
        # 合并后应有 总销售额 + 总销量 两列（task1 sum销售额 + task14 sum销量，同区块名合并）
        has_merged_cols = "总销售额" in header and "总销量" in header
        checks.append(("区块名不连续合并", has_merged_cols, f"header={header}"))
        # 验证行维度列（地区）头有特殊颜色
        ws_block = wb["按地区汇总"]
        # header在第2行（第1行是区块标题），找到"地区"列的位置
        for ci in range(1, len(header) + 1):
            cell = ws_block.cell(row=2, column=ci)
            if str(cell.value) == "地区":
                dim_fill = cell.fill
                # 行维度列头颜色应区别于普通列头（HEADER_FILL="4472C4" vs DIM_HEADER_FILL="5B9BD5"）
                checks.append(("行维度列头区分色", dim_fill.start_color.rgb != "4472C4",
                               f"地区 fill={dim_fill.start_color.rgb}"))
                break

    # 11. 历史标量引用：任务15生产"总销售额"→任务16公式"销量/总销售额=销量占比"
    if "历史标量" in actual_sheets:
        _, data = _read_sheet_data("历史标量")
        total_sales = float(data.iloc[0]["总销售额"]) if not data.empty and "总销售额" in data.columns else None
        checks.append(("标量生产者_总销售额", total_sales is not None and total_sales > 0,
                       f"total_sales={total_sales}"))
        
        if "按地区汇总" in actual_sheets:
            header2, data2 = _read_sheet_data("按地区汇总")
            has_ratio_col = "销量占比" in header2
            checks.append(("标量消费者_销量占比列", has_ratio_col, f"header={header2}"))
            if has_ratio_col and total_sales:
                # 华北 销量=270，华东=390，华南=370
                expected_ratios = {"华北": 270/total_sales, "华东": 390/total_sales, "华南": 370/total_sales}
                ratio_ok = True
                for idx, row in data2.iterrows():
                    region = str(row["地区"]) if "地区" in data2.columns else ""
                    if region in expected_ratios:
                        actual_ratio = float(row["销量占比"])
                        if abs(actual_ratio - expected_ratios[region]) > 0.001:
                            ratio_ok = False
                            break
                checks.append(("标量消费者_占比正确", ratio_ok,
                               f"expected={expected_ratios}"))

    wb.close()

    # 打印结果
    all_ok = True
    for name, ok, detail in checks:
        icon = "✓" if ok else "✗"
        color = GREEN if ok else RED
        detail_str = f": {detail}" if detail else ""
        print(f"  {color}{icon} {name}{detail_str}{RESET}")
        if not ok:
            all_ok = False

    return all_ok


def verify_theme_config():
    """验证 P0 主题色配置解析（从配置 Excel 到 PptTheme）"""
    print_header("P0 主题色配置解析验证")
    import tempfile
    sys.path.insert(0, PROJECT_DIR)
    from app.src.excel_reader import read_config
    from app.src.ppt_theme import PptTheme

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PPT配置"
    ws.append(["主色", "#1F4E79"])
    ws.append(["深色", "#1F3864"])
    ws.append(["accent1", "#2E75B6"])
    ws.append(["页脚", "测试报告"])
    ws.append(["字体", "思源黑体"])
    ws.append([])
    ws.append(["页码", "页面类型", "页面标题", "布局"])
    ws.append([1, "封面", "测试", ""])

    f = tempfile.mktemp(suffix=".xlsx")
    wb.save(f)

    checks = []
    try:
        cfg = read_config(f)
        t = PptTheme.from_config(cfg["colors"], cfg["general"])
        checks.append(("主色解析", t.primary == "#1F4E79", t.primary))
        checks.append(("深色解析", t.dark == "#1F3864", t.dark))
        checks.append(("强调色解析", t.accent_colors == ["#2E75B6"], str(t.accent_colors)))
        checks.append(("页脚解析", t.footer_text == "测试报告", t.footer_text))
        checks.append(("字体解析", t.font_name == "思源黑体", t.font_name))
    except Exception as e:
        checks.append(("主题色解析", False, str(e)))
    finally:
        if os.path.exists(f):
            os.remove(f)

    all_ok = True
    for name, ok, detail in checks:
        icon = "✓" if ok else "✗"
        color = GREEN if ok else RED
        print(f"  {color}{icon} {name}: {detail}{RESET}")
        if not ok:
            all_ok = False

    return all_ok


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
    from app.src.excel_reader import read_config
    
    config_path = os.path.join(SCRIPT_DIR, "项目配置.xlsx")
    data_path = os.path.join(SCRIPT_DIR, "测试数据.xlsx")
    ppt_pages = []
    try:
        config = read_config(config_path)
        ppt_pages = config.get("pages", [])
        print(f"  PPT 配置: {len(ppt_pages)} 页")
    except Exception as e:
        print(f"  [警告] 读取 PPT 配置失败: {e}")
    
    html_path = build_html(
        excel_path=excel_path,
        ppt_path=ppt_path,
        ppt_config=ppt_pages if ppt_pages else None,
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

    # Step 5.5: 自动化验证（Excel + PPT功能 + 主题色）
    excel_ok = verify_excel_output(pivot_excel)
    features_ok = verify_ppt_features(ppt_path)
    theme_ok = verify_theme_config()

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
    if pivot_ok and ppt_ok and excel_ok and features_ok and theme_ok:
        print(f"  {GREEN}{BOLD}✓ 全部通过 - 请人工检查输出格式{RESET}")
    else:
        print(f"  {RED}{BOLD}✗ 存在失败项{RESET}")
        if not pivot_ok:
            print(f"    - Excel 生成失败")
        if not ppt_ok:
            print(f"    - PPT 生成失败")
        if not excel_ok:
            print(f"    - Excel 输出验证未通过")
        if not features_ok:
            print(f"    - P0/P1 功能验证未通过")
        if not theme_ok:
            print(f"    - 主题色配置解析未通过")

    # 启动本地服务器供预览
    if html_path and os.path.exists(html_path):
        print(f"\n  {CYAN}启动预览服务器...{RESET}")
        start_preview_server(html_path)

    print(f"\n  {CYAN}提示: 请打开输出文件检查格式是否正常{RESET}")
    print(f"  {CYAN}  Excel: {pivot_excel}{RESET}" if pivot_excel else "")
    print(f"  {CYAN}  PPT:   {ppt_path}{RESET}" if ppt_path else "")


if __name__ == "__main__":
    main()
