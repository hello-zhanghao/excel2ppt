import openpyxl
import pandas as pd
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter


HEADER_FONT = Font(name="Microsoft YaHei", bold=True, size=11, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

TOTAL_FONT = Font(name="Microsoft YaHei", bold=True, size=11)
TOTAL_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")

BLOCK_TITLE_FONT = Font(name="Microsoft YaHei", bold=True, size=11, color="FFFFFF")
BLOCK_TITLE_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
BLOCK_TITLE_ALIGNMENT = Alignment(horizontal="center", vertical="center")

DATA_FONT = Font(name="Microsoft YaHei", size=10)
DATA_ALIGNMENT = Alignment(horizontal="center", vertical="center")

# 行维度列配色（区分分组维度列与数值列）
DIM_HEADER_FILL = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
DIM_DATA_FILL = PatternFill(start_color="EBF5FB", end_color="EBF5FB", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


def write_results(tasks, results, errors, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    if len(tasks) != len(results):
        print(f"    [警告] 任务数({len(tasks)})与结果数({len(results)})不一致，取较小值对齐")

    groups = {}
    for task, result in zip(tasks, results):
        if result is None:
            continue
        sheet_name = _safe_sheet_name(task.get("结果Sheet", "分析结果"))
        if sheet_name not in groups:
            groups[sheet_name] = []
        groups[sheet_name].append((task, result))

    for sheet_name, group_items in groups.items():
        _write_multi_result_sheet(wb, sheet_name, group_items)

    if errors:
        ws = wb.create_sheet("错误信息")
        ws.append(["任务", "错误详情"])
        for err in errors:
            ws.append([err.get("序号", ""), err.get("错误", "")])
        _style_header_row(ws)

    wb.save(output_path)
    return output_path


def _write_multi_result_sheet(wb, sheet_name, group_items):
    ws = wb.create_sheet(sheet_name)

    # 按「区块名」分组（非连续的相同区块名自动合并），按首次出现顺序输出
    block_groups = []
    seen_blocks = set()
    for task, result in group_items:
        block_name = _get_block_title(task)
        if block_name not in seen_blocks:
            block_groups.append((block_name, []))
            seen_blocks.add(block_name)
        for name, items in block_groups:
            if name == block_name:
                items.append((task, result))
                break

    current_row = 1
    for block_name, items in block_groups:
        row_dims = _get_row_dims(items[0][0])

        if not isinstance(items[0][1], dict):
            for task, result in items:
                if current_row > 1:
                    current_row += 1
                _write_block_title(ws, block_name, current_row, 2)
                current_row += 1
                _write_scalar_block(ws, task, result, block_name, current_row,
                                    task.get("_pct_columns", []), task.get("_number_formats", {}))
                current_row = ws.max_row + 1
            continue

        if current_row > 1:
            current_row += 1

        if len(items) > 1:
            merged_df, merged_pct_cols, merged_number_formats = _merge_same_dim_results(items, row_dims)
            if merged_df is None:
                continue
            _write_block_title(ws, block_name, current_row, len(merged_df.columns))
            current_row += 1
            _write_df_block(ws, merged_df, current_row, merged_pct_cols, row_dims, merged_number_formats)
            current_row = ws.max_row + 1
        else:
            task, result = items[0]
            pct_cols = task.get("_pct_columns", [])
            num_fmts = task.get("_number_formats", {})
            for key, df in result.items():
                if current_row > 1:
                    current_row += 1
                _write_block_title(ws, block_name, current_row, len(df.columns))
                current_row += 1
                _write_df_block(ws, df, current_row, pct_cols, row_dims, num_fmts)
                current_row = ws.max_row + 1

    if ws.max_column and ws.max_row:
        _auto_fit_columns(ws)
        ws.freeze_panes = ws.cell(row=1, column=1)


def _get_block_title(task):
    """获取区块标题：优先用区块名，否则用任务序号"""
    title = task.get("区块名", task.get("备注"))
    if title is not None and str(title).strip():
        return str(title).strip()
    seq = task.get("序号", "?")
    return f"任务{seq}"


def _write_block_title(ws, title, row, col_count):
    """在指定行写入区块标题（合并单元格，深蓝底白字）"""
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = BLOCK_TITLE_FONT
    cell.fill = BLOCK_TITLE_FILL
    cell.alignment = BLOCK_TITLE_ALIGNMENT
    cell.border = THIN_BORDER
    if col_count > 1:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=col_count)
    for ci in range(2, col_count + 1):
        c = ws.cell(row=row, column=ci)
        c.fill = BLOCK_TITLE_FILL
        c.border = THIN_BORDER


def _get_row_dims(task):
    raw = str(task.get("行维度", "")).strip()
    if not raw:
        return []
    return [d.strip() for d in raw.split(",") if d.strip()]


def _merge_same_dim_results(items, row_dims):
    merged = None
    merged_pct_cols = set()
    merged_number_formats = {}  # {列名: Excel数字格式}

    for task, result in items:
        if not isinstance(result, dict):
            continue
        task_pct_cols = set(task.get("_pct_columns", []))
        task_number_formats = task.get("_number_formats", {}) or {}
        for key, df in result.items():
            if df is None:
                continue
            df = df.copy()

            if merged is None:
                merged = df
                # 记录本 df 中属于 pct 的列
                for c in df.columns:
                    col_str = str(c)
                    if col_str in task_pct_cols:
                        merged_pct_cols.add(col_str)
                    if col_str in task_number_formats:
                        merged_number_formats[col_str] = task_number_formats[col_str]
            else:
                dup = [c for c in df.columns if c in merged.columns]
                df_rest = df.drop(columns=dup, errors="ignore")
                for c in df_rest.columns:
                    col_str = str(c)
                    if col_str in task_pct_cols:
                        merged_pct_cols.add(col_str)
                    if col_str in task_number_formats:
                        merged_number_formats[col_str] = task_number_formats[col_str]
                merged = pd.concat([merged.reset_index(drop=True), df_rest.reset_index(drop=True)], axis=1)

    return merged, list(merged_pct_cols), merged_number_formats


def _write_df_block(ws, df, start_row, pct_columns=None, row_dims=None, number_formats=None):
    pct_columns = set(pct_columns or [])
    row_dims = set(row_dims or [])
    number_formats = number_formats or {}  # {列名: PPT格式串}，需转换为 Excel number_format
    headers = list(df.columns)
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=start_row, column=ci, value=str(h))
    _style_header_row(ws, start_row, len(headers), row_dims)

    for ri, (idx, row_data) in enumerate(df.iterrows()):
        row_num = start_row + 1 + ri
        for ci, val in enumerate(row_data, 1):
            col_name = headers[ci - 1] if ci <= len(headers) else None
            col_name_str = str(col_name) if col_name is not None else ""
            is_pct = _is_pct_col(col_name, pct_columns)
            cell = ws.cell(row=row_num, column=ci, value=_format_cell_value(val, col_name, is_pct))
            # 优先级：自定义格式（聚合方式|PPT格式）> pct 默认 > 0.00 默认
            # number_formats 存的是 PPT 格式串（如 .2f），需转换为 Excel number_format
            excel_fmt = _ppt_fmt_to_excel_fmt(number_formats.get(col_name_str, ""))
            if excel_fmt:
                cell.number_format = excel_fmt
            elif is_pct:
                cell.number_format = PCT_NUMBER_FORMAT
            elif isinstance(val, (int, float)):
                cell.number_format = VALID_NUMBER_FORMAT

        is_total = str(idx) == "合计" or str(idx) == "总计"
        for ci in range(1, len(headers) + 1):
            col_name = headers[ci - 1] if ci <= len(headers) else ""
            is_dim = col_name in row_dims
            cell = ws.cell(row=row_num, column=ci)
            cell.alignment = DATA_ALIGNMENT
            cell.font = TOTAL_FONT if is_total else DATA_FONT
            if is_total:
                cell.fill = TOTAL_FILL
            elif is_dim:
                cell.fill = DIM_DATA_FILL
            else:
                cell.fill = PatternFill()
            cell.border = THIN_BORDER


def _write_scalar_block(ws, task, result, remark, start_row, pct_columns=None, number_formats=None):
    pct_columns = set(pct_columns or [])
    number_formats = number_formats or {}  # {指标名: PPT格式串}，需转换为 Excel number_format
    ws.cell(row=start_row, column=1, value="指标")
    ws.cell(row=start_row, column=2, value="值")
    _style_header_row(ws, start_row, 2)

    row = start_row + 1
    for key, val in result.items():
        key_str = str(key)
        c1 = ws.cell(row=row, column=1, value=key)
        is_pct = _is_pct_col(key, pct_columns)
        c2 = ws.cell(row=row, column=2, value=_format_cell_value(val, key, is_pct))
        # 优先级：自定义格式 > pct 默认 > 0.00 默认
        excel_fmt = _ppt_fmt_to_excel_fmt(number_formats.get(key_str, ""))
        if excel_fmt:
            c2.number_format = excel_fmt
        elif is_pct:
            c2.number_format = PCT_NUMBER_FORMAT
        elif isinstance(val, (int, float)):
            c2.number_format = VALID_NUMBER_FORMAT
        c1.alignment = DATA_ALIGNMENT
        c2.alignment = DATA_ALIGNMENT
        c1.border = THIN_BORDER
        c2.border = THIN_BORDER
        row += 1
def _safe_sheet_name(name):
    if not name or not str(name).strip():
        return "分析结果"
    name = str(name)
    illegal = r"[]:*?/\\"
    for ch in illegal:
        name = name.replace(ch, "_")
    return name[:31]


def _style_header_row(ws, row=1, max_col=None, dim_cols=None):
    dim_cols = set(dim_cols or [])
    if max_col is None:
        for cell in ws[row]:
            col_name = str(cell.value) if cell.value else ""
            cell.font = HEADER_FONT
            cell.fill = DIM_HEADER_FILL if col_name in dim_cols else HEADER_FILL
            cell.alignment = HEADER_ALIGNMENT
            cell.border = THIN_BORDER
        return
    for ci in range(1, max_col + 1):
        cell = ws.cell(row=row, column=ci)
        col_name = str(cell.value) if cell.value else ""
        cell.font = HEADER_FONT
        cell.fill = DIM_HEADER_FILL if col_name in dim_cols else HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


# 百分比列名兜底关键词（仅当 task 未携带 _pct_columns 时使用）
# 注意："率"过于宽泛（功率/速率/频率等非百分比字段），已移除
PCT_KEYWORDS = ["占比", "pct", "百分比", "比例"]


def _is_pct_col(col_name, pct_columns=None):
    """判断列是否为百分比列。
    优先用 task 携带的 pct_columns 元信息（精确），列名关键词仅作兜底。
    """
    if col_name is None:
        return False
    name = str(col_name)
    if pct_columns and name in pct_columns:
        return True
    return any(kw in name for kw in PCT_KEYWORDS)


def _format_cell_value(val, col_name=None, is_pct=False):
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    if isinstance(val, float) and val == int(val):
        return int(val)
    return val

VALID_NUMBER_FORMAT = '0.00'
PCT_NUMBER_FORMAT = '0.0%'


def _ppt_fmt_to_excel_fmt(ppt_fmt):
    """将 PPT 占位符的格式串（Python format spec 风格）转换为 Excel number_format。

    与 template_filler.py 的 _KNOWN_FMTS 语法一致，聚合方式的 |格式 后缀复用同一套语法。

    映射规则：
        .Nf       → 0.00（N 位小数，如 .2f → 0.00，.0f → 0）
        int / d   → 0（整数）
        .N%       → 0.00%（N 位小数百分比，如 .2% → 0.00%）
        ,.Nf      → #,##0.00（千分位 + N 位小数）
        ,.N%      → #,##0.00%（千分位 + 百分比）
        .Ne/.NE   → 0.00E+00（科学计数法）

    :param ppt_fmt: PPT 格式串（如 ".2f"、".1%"、",.0f"）
    :return: Excel number_format 字符串。无法识别时返回 None（用默认）。
    """
    if not ppt_fmt:
        return None
    fmt = ppt_fmt.strip()

    # 整数格式
    if fmt in ("int", "d"):
        return "0"

    # 千分位 + 百分比：,.N%
    if fmt.startswith(",.") and fmt.endswith("%"):
        try:
            digits = int(fmt[2:-1])
            return f"#,##0.{'0' * digits}%"
        except ValueError:
            return None

    # 千分位 + 小数：,.Nf
    if fmt.startswith(",.") and fmt.endswith("f"):
        try:
            digits = int(fmt[2:-1])
            return f"#,##0.{'0' * digits}" if digits > 0 else "#,##0"
        except ValueError:
            return None

    # 百分比：.N%
    if fmt.startswith(".") and fmt.endswith("%"):
        try:
            digits = int(fmt[1:-1])
            return f"0.{'0' * digits}%" if digits > 0 else "0%"
        except ValueError:
            return None

    # 小数：.Nf
    if fmt.startswith(".") and fmt.endswith("f"):
        try:
            digits = int(fmt[1:-1])
            return f"0.{'0' * digits}" if digits > 0 else "0"
        except ValueError:
            return None

    # 科学计数法：.Ne / .NE
    if fmt.startswith(".") and (fmt.endswith("e") or fmt.endswith("E")):
        try:
            digits = int(fmt[1:-1])
            zero_str = "0" * digits
            return f"0.{zero_str}E+00" if digits > 0 else "0E+00"
        except ValueError:
            return None

    return None


def _auto_fit_columns(ws):
    for col_cells in ws.columns:
        if not col_cells or len(col_cells) == 0:
            continue
        col_letter = get_column_letter(col_cells[0].column)
        max_length = 0
        for cell in col_cells:
            if cell.value:
                val_str = str(cell.value)
                length = 0
                for ch in val_str:
                    if "\u4e00" <= ch <= "\u9fff" or "\u3000" <= ch <= "\u303f" or "\uff00" <= ch <= "\uffef":
                        length += 2
                    else:
                        length += 1
                max_length = max(max_length, length)
        adjusted = min(max_length + 3, 40)
        ws.column_dimensions[col_letter].width = max(adjusted, 8)
