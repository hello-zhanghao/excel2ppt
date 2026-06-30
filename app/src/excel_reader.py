import os
import openpyxl
import pandas as pd


def read_config(config_path, sheet_name=None):
    wb = openpyxl.load_workbook(config_path, data_only=True)
    ws = _find_or_get_sheet(wb, sheet_name, _PPT_KEYWORDS)
    config = _parse_config(ws)
    wb.close()
    return config


_PPT_KEYWORDS = {"页码", "页面类型", "页面标题", "图表类型", "布局"}
_PIVOT_KEYWORDS = {"数据源", "行维度", "列维度", "值字段", "聚合方式"}


def _find_or_get_sheet(wb, sheet_name, keywords):
    if sheet_name and sheet_name in wb.sheetnames:
        return wb[sheet_name]
    for name in wb.sheetnames:
        ws = wb[name]
        row = [str(c.value).strip() if c.value else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
        if len(keywords & set(row)) >= 2:
            return ws
    return wb[wb.sheetnames[0]]


def _parse_config(ws):
    all_rows = list(ws.iter_rows(values_only=True))

    if not all_rows:
        return {"general": {}, "pages": [], "colors": {}}

    header_idx = None
    for i, row in enumerate(all_rows):
        if row[0] is not None and str(row[0]).strip() in ["页码", "序号", "编号"]:
            header_idx = i
            break

    if header_idx is None:
        header_idx = 0

    headers = [str(c).strip() if c else "" for c in all_rows[header_idx]]

    raw_rows = []
    for row in all_rows[header_idx + 1:]:
        if row[0] is None and all(c is None for c in row):
            continue
        item = {}
        for h, v in zip(headers, row):
            if h:
                item[h] = v
        raw_rows.append(item)

    pages = _group_rows_to_pages(raw_rows)

    return {
        "general": {},
        "colors": {},
        "pages": pages,
    }


def _group_rows_to_pages(raw_rows):
    """按页码分组，相同页码的行合并为一页，每页含charts列表"""
    pages = []
    current_page = None
    current_page_num = None
    last_sheet = None

    for row in raw_rows:
        page_num = row.get("页码")

        if page_num is not None and str(page_num).strip() != "":
            if current_page is not None:
                pages.append(current_page)
            last_sheet = None
            page_title = str(row.get("页面标题", "")).strip() if row.get("页面标题") else ""
            sub_title = str(row.get("副标题", "")).strip() if row.get("副标题") else ""
            if sub_title:
                page_title = page_title + "|" + sub_title if page_title else sub_title
            current_page = {
                "页码": page_num,
                "页面类型": str(row.get("页面类型", "内容")).strip() or "内容",
                "页面标题": page_title,
                "布局": str(row.get("布局", "")).strip() if row.get("布局") else "",
                "charts": [],
            }
            current_page_num = page_num

        if current_page is None:
            continue

        chart_title = str(row.get("图表标题", "")).strip() if row.get("图表标题") else ""
        if chart_title:
            sheet_val = str(row.get("数据Sheet", "")).strip() if row.get("数据Sheet") else ""
            if sheet_val:
                last_sheet = sheet_val
            elif last_sheet:
                sheet_val = last_sheet
            else:
                sheet_val = "Sheet1"
            chart_def = {
                "图表标题": chart_title,
                "图表类型": str(row.get("图表类型", "column")).strip() if row.get("图表类型") else "column",
                "数据Sheet": sheet_val,
                "X轴范围": str(row.get("X轴范围", row.get("X轴", ""))).strip() if (row.get("X轴范围") or row.get("X轴")) else "",
                "Y轴范围": str(row.get("Y轴范围", row.get("Y轴", ""))).strip() if (row.get("Y轴范围") or row.get("Y轴")) else "",
                "颜色": str(row.get("颜色", "")).strip() if row.get("颜色") else "",
                "区块名": str(row.get("区块名", "")).strip() if row.get("区块名") else "",
                "结论模板": str(row.get("结论模板", "")).strip() if row.get("结论模板") else "",
            }
            current_page["charts"].append(chart_def)

    if current_page is not None:
        pages.append(current_page)

    return pages


def read_data(excel_path, sheet_name, x_range, y_range, block_name=None):
    x_col_names = [cn.strip() for cn in str(x_range).split(",") if cn.strip()] if x_range else []
    y_col_names = [cn.strip() for cn in str(y_range).split(",") if cn.strip()] if y_range else []

    if block_name and str(block_name).strip():
        return _read_with_block_name(excel_path, sheet_name, str(block_name).strip(), x_range, y_range)

    if len(x_col_names) >= 2:
        return _read_hierarchical_multi_y(excel_path, sheet_name, x_col_names, y_col_names)

    x_values = _read_axis(excel_path, sheet_name, x_range, is_x=True)
    y_values = _read_axis(excel_path, sheet_name, y_range, is_x=False)
    return x_values, y_values


def _read_with_block_name(excel_path, sheet_name, block_name, x_range, y_range):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name]
    max_row = ws.max_row
    max_col = ws.max_column

    block_row = _find_block_name_row(ws, max_row, max_col, block_name)
    if block_row is None:
        wb.close()
        return [], []

    header_row = block_row + 1
    if header_row > max_row:
        wb.close()
        return [], []

    header_map = {}
    for c in range(1, max_col + 1):
        cell_val = ws.cell(row=header_row, column=c).value
        if cell_val:
            col_name = str(cell_val).strip()
            if col_name not in header_map:
                header_map[col_name] = c

    x_col_names = [cn.strip() for cn in str(x_range).split(",") if cn.strip()]
    y_col_names = [cn.strip() for cn in str(y_range).split(",") if cn.strip()]

    x_cols = [header_map[cn] for cn in x_col_names if cn in header_map]
    y_cols = [header_map[cn] for cn in y_col_names if cn in header_map]

    if not x_cols:
        wb.close()
        return [], []

    if len(x_cols) >= 2:
        result = _read_hierarchical_multi_from_ws(ws, header_row, max_row, max_col, x_cols, y_cols, x_col_names, y_col_names)
        wb.close()
        return result

    x_values = []
    y_values = {}
    for ycn in y_col_names:
        y_values[ycn] = []

    for r in range(header_row + 1, max_row + 1):
        all_cells = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        if len(set(all_cells)) == 1 and all_cells[0] is None:
            break

        if _looks_like_header(ws, r, x_cols, x_col_names):
            break

        if len(x_cols) == 1:
            v = ws.cell(row=r, column=x_cols[0]).value
            if v is not None:
                x_values.append(str(v))
        else:
            parts = [str(ws.cell(row=r, column=xc).value) for xc in x_cols
                     if ws.cell(row=r, column=xc).value is not None]
            if parts:
                x_values.append(" - ".join(parts))

        for ycn, yc in zip(y_col_names, y_cols):
            v = ws.cell(row=r, column=yc).value
            if v is not None:
                y_values[ycn].append(v)

    wb.close()

    if len(y_cols) == 1:
        y_values = y_values.get(y_col_names[0], [])

    return x_values, y_values


def _read_grouped_from_ws(ws, header_row, max_row, max_col, x_cols, y_col, x_col_names):
    cat_col = x_cols[0]
    group_col = x_cols[1]

    categories = []
    cat_set = set()
    raw_map = {}

    for r in range(header_row + 1, max_row + 1):
        all_cells = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        if len(set(all_cells)) == 1 and all_cells[0] is None:
            break

        if _looks_like_header(ws, r, x_cols, x_col_names):
            break

        cat_val = ws.cell(row=r, column=cat_col).value
        grp_val = ws.cell(row=r, column=group_col).value
        y_val = ws.cell(row=r, column=y_col).value

        if cat_val is None or y_val is None:
            continue

        cat_str = str(cat_val)
        grp_str = str(grp_val) if grp_val is not None else ""

        if cat_str not in cat_set:
            categories.append(cat_str)
            cat_set.add(cat_str)

        if grp_str not in raw_map:
            raw_map[grp_str] = {}
        raw_map[grp_str][cat_str] = y_val

    y_values = {}
    for grp_str in raw_map:
        series = []
        for cat in categories:
            series.append(raw_map[grp_str].get(cat, None))
        y_values[grp_str] = series

    active_categories = set()
    for grp_str in raw_map:
        active_categories.update(raw_map[grp_str].keys())
    categories = [c for c in categories if c in active_categories]

    cat_series_count = {c: 0 for c in categories}
    for grp_str in raw_map:
        for c in categories:
            if c in raw_map[grp_str]:
                cat_series_count[c] += 1

    categories = [c for c in categories if cat_series_count.get(c, 0) >= 1]

    y_values = {}
    for grp_str in raw_map:
        series = [raw_map[grp_str].get(cat) for cat in categories]
        valid_count = sum(1 for v in series if v is not None)
        if valid_count >= 1:
            y_values[grp_str] = series

    return categories, y_values

    return categories, y_values


def _read_grouped(excel_path, sheet_name, x_col_names, y_col_name):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    cat_col = x_col_names[0]
    group_col = x_col_names[1]

    if cat_col not in df.columns or group_col not in df.columns or y_col_name not in df.columns:
        return [], []

    categories = []
    cat_set = set()
    raw_map = {}

    for _, row in df.iterrows():
        cat_val = row.get(cat_col)
        grp_val = row.get(group_col)
        y_val = row.get(y_col_name)

        if pd.isna(cat_val) or pd.isna(y_val):
            continue

        cat_str = str(cat_val)
        grp_str = str(grp_val) if pd.notna(grp_val) else ""

        if cat_str not in cat_set:
            categories.append(cat_str)
            cat_set.add(cat_str)

        if grp_str not in raw_map:
            raw_map[grp_str] = {}
        raw_map[grp_str][cat_str] = y_val

    y_values = {}
    for grp_str in raw_map:
        series = []
        for cat in categories:
            series.append(raw_map[grp_str].get(cat, None))
        y_values[grp_str] = series

    active_categories = set()
    for grp_str in raw_map:
        active_categories.update(raw_map[grp_str].keys())
    categories = [c for c in categories if c in active_categories]

    cat_series_count = {c: 0 for c in categories}
    for grp_str in raw_map:
        for c in categories:
            if c in raw_map[grp_str]:
                cat_series_count[c] += 1

    categories = [c for c in categories if cat_series_count.get(c, 0) >= 1]

    y_values = {}
    for grp_str in raw_map:
        series = [raw_map[grp_str].get(cat) for cat in categories]
        valid_count = sum(1 for v in series if v is not None)
        if valid_count >= 1:
            y_values[grp_str] = series

    return categories, y_values

    return categories, y_values


def _read_hierarchical_multi_y(excel_path, sheet_name, x_col_names, y_col_names):
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    cat_col = x_col_names[0]
    sub_col = x_col_names[1]
    if cat_col not in df.columns or sub_col not in df.columns:
        return [], {}
    hierarchical = []
    y_values = {ycn: [] for ycn in y_col_names}
    current_parent = None
    current_children = []
    for _, row in df.iterrows():
        cat_val = row.get(cat_col)
        if pd.isna(cat_val):
            continue
        cat_str = str(cat_val)
        if cat_str != current_parent:
            if current_children:
                hierarchical.append((current_parent, current_children))
            current_parent = cat_str
            current_children = []
        sub_val = row.get(sub_col)
        sub_str = str(sub_val) if pd.notna(sub_val) else ""
        children_y = {}
        for ycn in y_col_names:
            if ycn in df.columns:
                v = row.get(ycn)
                if pd.notna(v):
                    children_y[ycn] = v
        current_children.append((sub_str, children_y))
        for ycn in y_col_names:
            if ycn in children_y:
                y_values[ycn].append(children_y[ycn])
            else:
                y_values[ycn].append(None)
    if current_children:
        hierarchical.append((current_parent, current_children))
    if len(y_col_names) == 1:
        y_values = {y_col_names[0]: [v for v in y_values[y_col_names[0]] if v is not None]}
    return hierarchical, y_values


def _read_hierarchical(excel_path, sheet_name, x_col_names, y_col_name):
    hierarchical, y_values = _read_hierarchical_multi_y(excel_path, sheet_name, x_col_names, [y_col_name])
    flat = y_values.get(y_col_name, [])
    return hierarchical, {y_col_name: flat}


def _read_hierarchical_multi_from_ws(ws, header_row, max_row, max_col, x_cols, y_cols, x_col_names, y_col_names):
    cat_col = x_cols[0]
    sub_col = x_cols[1]
    hierarchical = []
    y_values = {ycn: [] for ycn, yc in zip(y_col_names, y_cols) if yc is not None}
    current_parent = None
    current_children = []
    for r in range(header_row + 1, max_row + 1):
        all_cells = [ws.cell(row=r, column=c).value for c in range(1, max_col + 1)]
        if len(set(all_cells)) == 1 and all_cells[0] is None:
            if current_children:
                hierarchical.append((current_parent, current_children))
            break
        if _looks_like_header(ws, r, x_cols, x_col_names):
            if current_children:
                hierarchical.append((current_parent, current_children))
            break
        cat_val = ws.cell(row=r, column=cat_col).value
        if cat_val is None:
            continue
        cat_str = str(cat_val)
        if cat_str != current_parent:
            if current_children:
                hierarchical.append((current_parent, current_children))
            current_parent = cat_str
            current_children = []
        sub_val = ws.cell(row=r, column=sub_col).value
        sub_str = str(sub_val) if sub_val is not None else ""
        children_y = {}
        for ycn, yc in zip(y_col_names, y_cols):
            if yc is not None:
                v = ws.cell(row=r, column=yc).value
                if v is not None:
                    children_y[ycn] = v
        current_children.append((sub_str, children_y))
        for ycn, yc in zip(y_col_names, y_cols):
            if ycn in children_y:
                y_values[ycn].append(children_y[ycn])
            else:
                y_values[ycn].append(None)
    if current_children:
        hierarchical.append((current_parent, current_children))
    if len(y_col_names) == 1:
        y_values = {y_col_names[0]: [v for v in y_values[y_col_names[0]] if v is not None]}
    return hierarchical, y_values


def _read_hierarchical_from_ws(ws, header_row, max_row, max_col, x_cols, y_col, x_col_names, y_col_name):
    hierarchical, y_values = _read_hierarchical_multi_from_ws(ws, header_row, max_row, max_col, x_cols, [y_col], x_col_names, [y_col_name])
    return hierarchical, y_values


def _find_block_name_row(ws, max_row, max_col, block_name):
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value and block_name in str(cell.value).strip():
                return r
    return None


def _looks_like_header(ws, row, x_cols, x_col_names):
    for xc, xn in zip(x_cols, x_col_names):
        val = ws.cell(row=row, column=xc).value
        if val is not None and str(val).strip() == xn:
            return True
    return False


def _is_range_format(s):
    s = str(s).strip()
    return ":" in s and any(c.isalpha() for c in s.split(":")[0])


def _read_axis(excel_path, sheet_name, ref, is_x=False):
    if ref is None or str(ref).strip() == "":
        return []

    ref_str = str(ref).strip()

    if _is_range_format(ref_str):
        return _read_range_by_range(excel_path, sheet_name, ref_str, combine_multi_col=is_x)

    return _read_range_by_columns(excel_path, sheet_name, ref_str, combine_multi_col=is_x)


def _read_range_by_columns(excel_path, sheet_name, col_names_str, combine_multi_col=False):
    col_names = [c.strip() for c in str(col_names_str).split(",") if c.strip()]

    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    available = [c for c in df.columns if c in col_names]
    if not available:
        return []

    if len(available) == 1 or not combine_multi_col:
        series_cols = []
        for cn in col_names:
            if cn in df.columns:
                series_cols.append(cn)
        if not combine_multi_col and len(series_cols) == 1:
            return df[series_cols[0]].dropna().tolist()
        if not combine_multi_col and len(series_cols) > 1:
            result = {}
            for cn in series_cols:
                result[cn] = df[cn].dropna().tolist()
            return result

    combined = []
    for idx, row in df.iterrows():
        parts = []
        all_none = True
        for cn in col_names:
            if cn in df.columns:
                val = row[cn]
                if pd.notna(val):
                    all_none = False
                    parts.append(str(val))
                else:
                    parts.append("")
        if not all_none:
            valid_parts = [p for p in parts if p != ""]
            combined.append(" - ".join(valid_parts))

    return combined


def _read_range_by_range(excel_path, sheet_name, cell_range, combine_multi_col=False):
    range_str = str(cell_range).strip()
    if "!" in range_str:
        sheet_name, range_str = range_str.split("!", 1)

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name]

    start_cell, end_cell = range_str.split(":")
    start_col = _col_letter_to_index(start_cell)
    end_col = _col_letter_to_index(end_cell)
    start_row = int("".join(c for c in start_cell if c.isdigit()))
    end_row = int("".join(c for c in end_cell if c.isdigit()))

    num_cols = end_col - start_col + 1

    if num_cols <= 1 or not combine_multi_col:
        values = []
        for row in ws[range_str]:
            for cell in row:
                values.append(cell.value)
        wb.close()
        return values

    result = []
    for r in range(start_row, end_row + 1):
        parts = []
        all_none = True
        for c in range(start_col, end_col + 1):
            val = ws.cell(row=r, column=c + 1).value
            if val is not None:
                all_none = False
                parts.append(str(val))
            else:
                parts.append("")
        if not all_none:
            valid_parts = [p for p in parts if p != ""]
            result.append(" - ".join(valid_parts))

    wb.close()
    return result


def _col_letter_to_index(cell_ref):
    letters = "".join(c for c in cell_ref if c.isalpha()).upper()
    idx = 0
    for char in letters:
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    return idx - 1


def read_data_as_dataframe(excel_path, sheet_name):
    return pd.read_excel(excel_path, sheet_name=sheet_name)


def read_geo_data(excel_path, sheet_name, x_range, y_range):
    x_cols = [cn.strip() for cn in str(x_range).split(",") if cn.strip()] if x_range else []
    y_cols = [cn.strip() for cn in str(y_range).split(",") if cn.strip()] if y_range else []

    all_cols = x_cols + y_cols
    if not all_cols:
        return None

    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    available = [c for c in all_cols if c in df.columns]
    if not available:
        return None

    geo_df = df[available].copy()
    extra_cols = ["site_id", "cell_id", "siteid", "cellid", "站点ID", "小区ID"]
    for ec in extra_cols:
        if ec in df.columns and ec not in available:
            geo_df[ec] = df[ec]

    return geo_df
