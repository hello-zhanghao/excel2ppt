import openpyxl
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

    dim_groups = []
    current_dim_key = None
    for task, result in group_items:
        row_dims = _get_row_dims(task)
        dim_key = tuple(row_dims) if row_dims else ()
        if dim_key != current_dim_key or not dim_groups:
            dim_groups.append((dim_key, row_dims, []))
            current_dim_key = dim_key
        dim_groups[-1][2].append((task, result))

    current_row = 1
    for dim_key, row_dims, items in dim_groups:
        if not isinstance(items[0][1], dict):
            for task, result in items:
                if current_row > 1:
                    current_row += 1
                title = _get_block_title(task)
                _write_block_title(ws, title, current_row, 2)
                current_row += 1
                _write_scalar_block(ws, task, result, task.get("区块名", ""), current_row)
                current_row = ws.max_row + 1
            continue

        if current_row > 1:
            current_row += 1

        if len(items) > 1:
            merged_df = _merge_same_dim_results(items, row_dims)
            if merged_df is None:
                continue
            title = _get_block_title(items[0][0])
            _write_block_title(ws, title, current_row, len(merged_df.columns))
            current_row += 1
            _write_df_block(ws, merged_df, current_row)
            current_row = ws.max_row + 1
        else:
            task, result = items[0]
            for key, df in result.items():
                if current_row > 1:
                    current_row += 1
                title = _get_block_title(task)
                _write_block_title(ws, title, current_row, len(df.columns))
                current_row += 1
                _write_df_block(ws, df, current_row)
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

    for task, result in items:
        if not isinstance(result, dict):
            continue
        for key, df in result.items():
            if df is None:
                continue
            df = df.copy()

            if merged is None:
                merged = df
            else:
                dup = [c for c in df.columns if c in merged.columns]
                df_rest = df.drop(columns=dup, errors="ignore")
                merged = pd.concat([merged.reset_index(drop=True), df_rest.reset_index(drop=True)], axis=1)

    return merged


def _write_df_block(ws, df, start_row):
    headers = list(df.columns)
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=start_row, column=ci, value=str(h))
    _style_header_row(ws, start_row, len(headers))

    for ri, (idx, row_data) in enumerate(df.iterrows()):
        row_num = start_row + 1 + ri
        if isinstance(row_data, pd.core.series.Series):
            for ci, val in enumerate(row_data, 1):
                col_name = headers[ci - 1] if ci <= len(headers) else None
                cell = ws.cell(row=row_num, column=ci, value=_format_cell_value(val, col_name))
                if col_name and _is_pct_column(col_name):
                    cell.number_format = '0.0"%"'
        else:
            for ci, val in enumerate(row_data, 1):
                col_name = headers[ci - 1] if ci <= len(headers) else None
                cell = ws.cell(row=row_num, column=ci, value=_format_cell_value(val, col_name))
                if col_name and _is_pct_column(col_name):
                    cell.number_format = '0.0"%"'

        is_total = str(idx) == "合计" or str(idx) == "总计"
        for ci in range(1, len(headers) + 1):
            cell = ws.cell(row=row_num, column=ci)
            cell.alignment = DATA_ALIGNMENT
            cell.font = TOTAL_FONT if is_total else DATA_FONT
            cell.fill = TOTAL_FILL if is_total else PatternFill()
            cell.border = THIN_BORDER


def _write_scalar_block(ws, task, result, remark, start_row):
    ws.cell(row=start_row, column=1, value="指标")
    ws.cell(row=start_row, column=2, value="值")
    _style_header_row(ws, start_row, 2)

    row = start_row + 1
    for key, val in result.items():
        c1 = ws.cell(row=row, column=1, value=key)
        c2 = ws.cell(row=row, column=2, value=_format_cell_value(val))
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


def _style_header_row(ws, row=1, max_col=None):
    if max_col is None:
        for cell in ws[row]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
            cell.alignment = HEADER_ALIGNMENT
            cell.border = THIN_BORDER
        return
    for ci in range(1, max_col + 1):
        cell = ws.cell(row=row, column=ci)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER


PCT_KEYWORDS = ["占比", "率", "pct", "百分比", "比例"]


def _is_pct_column(col_name):
    if col_name is None:
        return False
    name = str(col_name)
    return any(kw in name for kw in PCT_KEYWORDS)


def _pct_already_scaled(col_name):
    return "%" in str(col_name) if col_name else False


def _format_cell_value(val, col_name=None):
    if val is None or (isinstance(val, float) and str(val) == "nan"):
        return ""
    if isinstance(val, (int, float)):
        if col_name and _is_pct_column(col_name):
            if _pct_already_scaled(col_name):
                return round(float(val), 1)
            else:
                return round(float(val) * 100, 1)
        if isinstance(val, float):
            if abs(val) >= 1000:
                return round(val, 1)
            if val == int(val):
                return int(val)
            return round(val, 2)
    return val


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


import pandas as pd
