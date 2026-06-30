import os
import re
import openpyxl
import pandas as pd
import numpy as np


AGG_MAP = {
    "sum": "sum",
    "avg": "mean",
    "mean": "mean",
    "count": "count",
    "max": "max",
    "min": "min",
    "nunique": "nunique",
    "去重计数": "nunique",
    "distinct": "nunique",
    "pct": "pct",
    "占比": "pct",
    "percentage": "pct",
}


def read_pivot_config(config_path, sheet_name=None):
    wb = openpyxl.load_workbook(config_path, data_only=True)

    from src.excel_reader import _find_or_get_sheet, _PIVOT_KEYWORDS
    ws = _find_or_get_sheet(wb, sheet_name, _PIVOT_KEYWORDS)

    rows = list(ws.iter_rows(values_only=True))

    if not rows:
        return []

    headers = [str(c).strip() if c else "" for c in rows[0]]

    tasks = []
    for row in rows[1:]:
        if all(c is None for c in row):
            continue
        item = {}
        for h, v in zip(headers, row):
            if h:
                item[h] = v

        task = {}
        task["序号"] = item.get("序号", "")
        task["数据源"] = str(item.get("数据源", "")).strip() if item.get("数据源") else ""
        task["sheet"] = str(item.get("Sheet", "")).strip() if item.get("Sheet") else "Sheet1"
        task["行维度"] = str(item.get("行维度", "")).strip() if item.get("行维度") else ""
        task["列维度"] = str(item.get("列维度", "")).strip() if item.get("列维度") else ""
        task["值字段"] = str(item.get("值字段", "")).strip() if item.get("值字段") else ""
        task["聚合方式"] = str(item.get("聚合方式", "sum")).strip() if item.get("聚合方式") else "sum"
        task["结果Sheet"] = str(item.get("结果Sheet", "")).strip() if item.get("结果Sheet") else f"结果{task.get('序号','')}"
        task["备注"] = str(item.get("备注", "")).strip() if item.get("备注") else ""
        task["行映射"] = str(item.get("行映射", "")).strip() if item.get("行映射") else (str(item.get("行维度映射", "")).strip() if item.get("行维度映射") else str(item.get("映射表", "")).strip() if item.get("映射表") else "")
        task["列映射"] = str(item.get("列映射", "")).strip() if item.get("列映射") else (str(item.get("列维度映射", "")).strip() if item.get("列维度映射") else "")
        task["值映射"] = str(item.get("值映射", "")).strip() if item.get("值映射") else (str(item.get("值字段映射", "")).strip() if item.get("值字段映射") else "")
        task["分箱"] = str(item.get("分箱", "")).strip() if item.get("分箱") else ""
        task["值计算"] = str(item.get("值计算", "")).strip() if item.get("值计算") else ""

        if not task["行维度"] and not task["值字段"]:
            continue

        tasks.append(task)

    wb.close()
    return tasks


def _resolve_data_path(data_source, config_dir):
    if not data_source:
        return None
    if os.path.isabs(data_source):
        return data_source
    return os.path.join(config_dir, data_source)


def _parse_join_spec(data_source):
    if " JOIN " not in data_source.upper():
        return None

    tokens = _tokenize_join(data_source)
    if not tokens:
        return None

    parts = []
    i = 0
    while i < len(tokens):
        if i == 0:
            left_table = tokens[i]
            i += 1
        else:
            left_table = parts[-1]["right"]

        if i >= len(tokens):
            break

        if tokens[i].upper() not in ("JOIN", "INNER JOIN", "LEFT JOIN", "RIGHT JOIN", "OUTER JOIN"):
            i += 1
            if i >= len(tokens):
                break

        join_type = tokens[i]
        i += 1
        if i >= len(tokens):
            break

        right_table = tokens[i]
        i += 1

        if i + 1 <= len(tokens) and i < len(tokens) and tokens[i].upper() == "ON":
            i += 1
            on_expr = tokens[i]
            i += 1
            if "=" in on_expr:
                parts_eq = on_expr.split("=", 1)
                left_key = parts_eq[0].strip()
                right_key = parts_eq[1].strip()
            else:
                left_key = on_expr
                right_key = ""
        else:
            left_key = ""
            right_key = ""

        how = "inner"
        jt = join_type.upper()
        if "LEFT" in jt:
            how = "left"
        elif "RIGHT" in jt:
            how = "right"
        elif "OUTER" in jt:
            how = "outer"

        left_file, left_sheet = _split_table_token(left_table)
        right_file, right_sheet = _split_table_token(right_table)

        parts.append({
            "left": left_file,
            "right": right_file,
            "left_sheet": left_sheet,
            "right_sheet": right_sheet,
            "left_key": left_key,
            "right_key": right_key,
            "how": how,
        })

    return parts


def _split_table_token(token):
    if "@" in token:
        idx = token.index("@")
        return token[:idx], token[idx + 1:]
    return token, None


def _tokenize_join(data_source):
    s = data_source.strip()
    keywords = ["INNER JOIN", "LEFT JOIN", "RIGHT JOIN", "OUTER JOIN", "JOIN", "ON"]
    tokens = []
    remaining = s
    while remaining:
        found = False
        for kw in keywords:
            pattern = re.compile(r"^\s*" + re.escape(kw) + r"\s+", re.IGNORECASE)
            m = pattern.match(remaining)
            if m:
                tokens.append(kw)
                remaining = remaining[m.end():]
                found = True
                break
        if not found:
            m = re.match(r"^\s*(.+?)(?=\s+(?:INNER\s+JOIN|LEFT\s+JOIN|RIGHT\s+JOIN|OUTER\s+JOIN|JOIN|ON)\b|$)", remaining, re.IGNORECASE)
            if m:
                raw = m.group(1).strip()
                tokens.append(raw)
                remaining = remaining[m.end():]
            else:
                remaining = remaining.strip()
                if remaining:
                    tokens.append(remaining)
                break
    return tokens


def _load_joined_dataframe(config_dir, data_source, sheet_name):
    join_parts = _parse_join_spec(data_source)

    if not join_parts:
        file_path = _resolve_data_path(data_source, config_dir)
        if not file_path or not os.path.exists(file_path):
            return None
        return pd.read_excel(file_path, sheet_name=sheet_name)

    candidate_files = _collect_candidate_xlsx(config_dir)

    first_file, first_sheet = _resolve_join_table(join_parts[0]["left"], config_dir, candidate_files)
    if first_file is None:
        return None
    if first_sheet is None:
        first_sheet = join_parts[0].get("left_sheet") or sheet_name

    df = pd.read_excel(first_file, sheet_name=first_sheet)

    for jp in join_parts:
        right_file, right_sheet = _resolve_join_table(jp["right"], config_dir, candidate_files)
        if right_file is None:
            right_file = first_file
            right_sheet = jp["right"]
            if right_sheet.lower().endswith(".xlsx"):
                right_sheet = right_sheet[:-5]
            if not _sheet_exists(right_file, right_sheet):
                continue
        if right_sheet is None:
            right_sheet = jp.get("right_sheet") or "Sheet1"

        df_right = pd.read_excel(right_file, sheet_name=right_sheet)

        left_on = jp["left_key"]
        right_on = jp["right_key"]

        if left_on not in df.columns:
            left_on = left_on.strip()
        if right_on not in df_right.columns:
            right_on = right_on.strip()

        if left_on not in df.columns or right_on not in df_right.columns:
            continue

        df = pd.merge(df, df_right, left_on=left_on, right_on=right_on, how=jp["how"], suffixes=("", "_r"))

    return df


def _resolve_join_table(table_name, config_dir, candidate_files):
    clean = table_name
    if clean.lower().endswith(".xlsx"):
        clean = clean[:-5]

    for f in candidate_files:
        if not f or not os.path.exists(f):
            continue
        for test_name in (table_name, clean):
            if _sheet_exists(f, test_name):
                return f, test_name

    file_path = _resolve_data_path(table_name, config_dir)
    if os.path.exists(file_path):
        return file_path, None
    file_path = _resolve_data_path(table_name + ".xlsx", config_dir)
    if os.path.exists(file_path):
        return file_path, None

    return None, None


def _sheet_exists(file_path, sheet_name):
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        result = sheet_name in wb.sheetnames
        wb.close()
        return result
    except Exception:
        return False


def _collect_candidate_xlsx(config_dir):
    import glob
    return [
        f for f in glob.glob(os.path.join(config_dir, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")
        and "配置" not in os.path.basename(f)
        and "config" not in os.path.basename(f).lower()
    ]


def run_analysis(task, config_dir):
    序号 = task.get("序号", "?")
    data_source = task.get("数据源", "")
    行维度_str = task.get("行维度", "")
    列维度_str = task.get("列维度", "")
    值字段_str = task.get("值字段", "")
    聚合方式_str = task.get("聚合方式", "sum")
    sheet_name = task.get("sheet", "Sheet1")

    if not 值字段_str:
        return None, f"[任务{序号}] 未指定值字段"

    行维度 = [d.strip() for d in 行维度_str.split(",") if d.strip()]
    列维度 = [d.strip() for d in 列维度_str.split(",") if d.strip()]
    值字段 = [v.strip() for v in 值字段_str.split(",") if v.strip()]
    聚合函数 = [a.strip() for a in 聚合方式_str.split(",") if a.strip()]
    聚合函数 = [AGG_MAP.get(a, a) for a in 聚合函数]

    df = _load_joined_dataframe(config_dir, data_source, sheet_name)
    if df is None or df.empty:
        return None, f"[任务{序号}] 数据为空或文件不存在: {data_source}"

    mapping_str = task.get("映射表", "")
    row_map_str = task.get("行映射", "")
    col_map_str = task.get("列映射", "")
    val_map_str = task.get("值映射", "")
    col_map, val_map = _parse_mapping(mapping_str, row_map_str, col_map_str, val_map_str, 行维度, 列维度, 值字段)
    if col_map or val_map:
        df = _apply_mapping(df, col_map, val_map)
        行维度 = [_translate_name(d, col_map) for d in 行维度]
        列维度 = [_translate_name(d, col_map) for d in 列维度]
        值字段 = [_translate_name(v, col_map) for v in 值字段]

    bin_spec = task.get("分箱", "")
    if bin_spec:
        行维度, 列维度, df = _apply_binning(bin_spec, 行维度, 列维度, df, col_map)

    missing_cols = []
    all_dims = 行维度 + 列维度
    for col in all_dims:
        if col not in df.columns:
            missing_cols.append(col)
    for col in 值字段:
        if col not in df.columns:
            missing_cols.append(col)

    if missing_cols:
        available = list(df.columns)
        return None, f"[任务{序号}] 列不存在: {missing_cols}。可用列: {available}"

    for col in 值字段:
        af_for_col = _get_agg_for_col(col, 聚合函数, 值字段)
        if af_for_col not in ("count", "nunique"):
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=值字段, how="all")
    if df.empty:
        return None, f"[任务{序号}] 值字段全为空"

    if 列维度 and 列维度[0]:
        result = _cross_pivot(df, 行维度, 列维度, 值字段, 聚合函数, task)
    else:
        result = _group_aggregate(df, 行维度, 值字段, 聚合函数, task)

    val_calc = task.get("值计算", "")
    if val_calc and result:
        result = _apply_value_calc(result, val_calc, 值字段, 聚合函数)

    return result, None


def _cross_pivot(df, row_dims, col_dims, value_cols, agg_funcs, task):
    results = {}
    for vcol in value_cols:
        for af in agg_funcs:
            is_pct = af in ("pct", "占比", "percentage")
            actual_af = "sum" if is_pct else af
            pivot = pd.pivot_table(
                df,
                values=vcol,
                index=row_dims,
                columns=col_dims,
                aggfunc=actual_af,
                fill_value=0,
            )

            if is_pct:
                grand = pivot.values.sum()
                if grand != 0:
                    pivot = pivot / grand
                    pivot = pivot.map(lambda x: round(x, 4))

            pivot = pivot.reset_index()

            dim_cols = [c for c in pivot.columns if c not in pivot.select_dtypes(include=np.number).columns]
            if not dim_cols:
                dim_cols = [pivot.columns[0]]
            numeric_cols = [c for c in pivot.columns if c in pivot.select_dtypes(include=np.number).columns]

            col_sums = {}
            for nc in numeric_cols:
                col_sums[nc] = pivot[nc].sum()

            total_row_data = {}
            for c in pivot.columns:
                if c in col_sums:
                    total_row_data[c] = col_sums[c]
                elif c == dim_cols[0]:
                    total_row_data[c] = "合计"
                else:
                    total_row_data[c] = ""
            pivot = pd.concat([pivot, pd.DataFrame([total_row_data])], ignore_index=True)

            for c in pivot.columns:
                if c not in dim_cols and c != "合计":
                    try:
                        pivot[c] = pd.to_numeric(pivot[c], errors="coerce")
                    except Exception:
                        pass

            row_sums = pd.Series(0.0, index=pivot.index)
            for nc in numeric_cols:
                row_sums = row_sums + pd.to_numeric(pivot[nc], errors="coerce").fillna(0)
            if "合计" in pivot.columns:
                row_sums = row_sums + pd.to_numeric(pivot["合计"], errors="coerce").fillna(0)
            pivot["合计"] = row_sums

            agg_label = {"sum": "求和", "mean": "均值", "avg": "均值", "count": "计数", "max": "最大值", "min": "最小值", "nunique": "去重计数", "pct": "占比"}.get(af, af)
            if _is_display_name(vcol):
                key = vcol
            elif len(value_cols) * len(agg_funcs) > 1:
                key = f"{vcol}_{agg_label}"
            else:
                key = vcol
            results[key] = pivot

    return results


def _group_aggregate(df, group_cols, value_cols, agg_funcs, task):
    if not group_cols:
        rows = []
        for vcol in value_cols:
            for af in agg_funcs:
                actual_af = "sum" if af in ("pct", "占比", "percentage") else af
                val = df[vcol].agg(actual_af)
                if af in ("pct", "占比", "percentage"):
                    total = float(df[vcol].sum())
                    val = float(val) / total if total != 0 else 0
                agg_label = {"sum": "求和", "mean": "均值", "avg": "均值", "count": "计数", "max": "最大值", "min": "最小值", "nunique": "去重计数", "pct": "占比"}.get(af, af)
                rows.append({"指标": f"{vcol}_{agg_label}", "值": round(float(val), 4) if isinstance(val, (int, float)) else val})
        result_df = pd.DataFrame(rows)
        return {"结果": result_df}

    agg_dict = {}
    has_pct = False
    per_field_funcs = _align_funcs_to_fields(value_cols, agg_funcs)
    for idx, vcol in enumerate(value_cols):
        f_clean = []
        for a in per_field_funcs.get(vcol, agg_funcs):
            a_mapped = AGG_MAP.get(a, a)
            if a_mapped in ("pct", "占比", "percentage"):
                f_clean.append("sum")
                has_pct = True
            else:
                f_clean.append(a_mapped)
        agg_dict[vcol] = f_clean

    grouped = df.groupby(group_cols, as_index=False, observed=True).agg(agg_dict)

    orig_tuples = [col for col in grouped.columns.values]
    grouped.columns = [_flatten_col(col) for col in orig_tuples]

    seen = {}
    for i, name in enumerate(grouped.columns):
        if name in seen and name not in group_cols:
            prev = seen[name]
            rec = lambda t, n: "_".join(str(p).strip() for p in t if str(p).strip()) if isinstance(t, tuple) else n
            grouped.columns.values[i] = rec(orig_tuples[i], name)
            grouped.columns.values[prev] = rec(orig_tuples[prev], name)
        else:
            seen[name] = i

    grouped = grouped.sort_values(group_cols)

    if has_pct:
        for col in grouped.columns:
            if col not in group_cols and pd.api.types.is_numeric_dtype(grouped[col]):
                total = grouped[col].sum()
                if total != 0:
                    grouped[col] = grouped[col] / total
                    grouped[col] = grouped[col].round(4)
    else:
        for col in grouped.columns:
            if col not in group_cols:
                grouped[col] = grouped[col].round(2)

    return {"结果": grouped}


def _flatten_col(col):
    if isinstance(col, tuple):
        parts = [str(p).strip() for p in col if str(p).strip()]
        if len(parts) >= 2 and _is_display_name(parts[0]):
            return parts[0]
        return "_".join(parts) if parts else str(col)
    return str(col).strip()


def _is_display_name(name):
    return any(c in name for c in "()（）") or any("\u4e00" <= c <= "\u9fff" for c in name)


def _get_agg_for_col(col, agg_funcs, value_cols):
    idx = value_cols.index(col) if col in value_cols else 0
    if idx < len(agg_funcs):
        return agg_funcs[idx]
    return agg_funcs[0] if agg_funcs else "sum"


def _align_funcs_to_fields(value_cols, agg_funcs):
    if len(value_cols) == len(agg_funcs):
        return {vc: [agg_funcs[i]] for i, vc in enumerate(value_cols)}
    return {}


def _auto_find_data_files(config_dir, config_path):
    import glob
    config_name = os.path.basename(config_path).lower()
    xlsx_files = glob.glob(os.path.join(config_dir, "*.xlsx"))
    candidates = []
    for f in xlsx_files:
        name = os.path.basename(f)
        if name.startswith("~$"):
            continue
        if name.lower() == config_name:
            continue
        if "配置" in name or "config" in name.lower():
            continue
        candidates.append(f)
    return candidates


def _parse_mapping(mapping_str, row_map_str, col_map_str, val_map_str, 行维度, 列维度, 值字段):
    parts_legacy = [p.strip() for p in mapping_str.split(",") if p.strip()] if mapping_str and mapping_str.strip() else []
    if parts_legacy and any("=" in p for p in parts_legacy):
        return _parse_old_new_mapping(parts_legacy)

    col_map = {}
    if row_map_str or col_map_str or val_map_str:
        col_map = _map_fields(row_map_str, 行维度, col_map)
        col_map = _map_fields(col_map_str, 列维度, col_map)
        col_map = _map_fields(val_map_str, 值字段, col_map)
    elif parts_legacy:
        all_fields = 行维度 + 列维度 + 值字段
        for i, new_name in enumerate(parts_legacy):
            if i < len(all_fields) and new_name:
                col_map[all_fields[i]] = new_name

    return col_map, {}


def _map_fields(map_str, fields, col_map):
    if not map_str or not map_str.strip():
        return col_map
    parts = [p.strip() for p in map_str.split(",") if p.strip()]
    for i, new_name in enumerate(parts):
        if i < len(fields) and new_name:
            col_map[fields[i]] = new_name
    return col_map


def _parse_old_new_mapping(parts):
    col_map = {}
    val_map = {}
    for pair in parts:
        if "=" not in pair:
            continue
        src, dst = pair.split("=", 1)
        src = src.strip()
        dst = dst.strip()
        if not src or not dst:
            continue
        if " " not in dst and not any(c in dst for c in "():（）"):
            col_map[src] = dst
        val_map[src] = dst
    return col_map, val_map


def _translate_name(name, col_map):
    if name in col_map:
        return col_map[name]
    return name


def _apply_mapping(df, col_map, val_map):
    rename = {}
    for src, dst in col_map.items():
        if src in df.columns:
            rename[src] = dst
        else:
            for c in df.columns:
                if c.startswith(src):
                    new_c = dst + c[len(src):]
                    rename[c] = new_c
    df = df.rename(columns=rename)
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col]):
            df[col] = df[col].astype(str).map(lambda x: val_map.get(x, x))
    return df


def _apply_binning(bin_spec, row_dims, col_dims, df, col_map):
    spec = bin_spec.strip()
    if "=" not in spec:
        return row_dims, col_dims, df
    col_name, params_str = spec.split("=", 1)
    col_name = col_name.strip()
    params_str = params_str.strip()
    if col_name in col_map:
        col_name = col_map[col_name]
    if col_name not in df.columns:
        return row_dims, col_dims, df
    parts = [p.strip() for p in params_str.split(",")]
    if len(parts) < 3:
        return row_dims, col_dims, df

    try:
        if len(parts) == 3:
            start = float(parts[0])
            end = float(parts[1])
            bins = int(parts[2])
            if bins < 1:
                return row_dims, col_dims, df
            bin_edges = np.linspace(start, end, bins + 1)
        else:
            bin_edges = np.array([float(p) for p in parts])
            if len(np.unique(bin_edges)) < 2:
                return row_dims, col_dims, df
            bin_edges = np.sort(np.unique(bin_edges))
    except ValueError:
        return row_dims, col_dims, df

    labels = []
    for i in range(len(bin_edges) - 1):
        labels.append(f"{_fmt_num(bin_edges[i])}~{_fmt_num(bin_edges[i+1])}")
    binned_col = col_name + "_区间"
    df[binned_col] = pd.cut(df[col_name], bins=bin_edges, labels=labels, include_lowest=True)
    new_row_dims = [binned_col if d == col_name else d for d in row_dims]
    new_col_dims = [binned_col if d == col_name else d for d in col_dims]
    return new_row_dims, new_col_dims, df


def _fmt_num(val):
    if val == int(val):
        return str(int(val))
    return f"{val:.1f}"


def _apply_value_calc(result, val_calc, value_cols, agg_funcs):
    calcs = [c.strip() for c in val_calc.split(",")]
    calc_map = {}
    for i, expr in enumerate(calcs):
        if not expr:
            continue
        if i < len(value_cols):
            calc_map[value_cols[i]] = expr

    for key, df in result.items():
        if not isinstance(df, pd.DataFrame):
            continue
        for col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                continue
            expr = _find_matching_calc(col, calc_map, value_cols)
            if expr:
                try:
                    if expr.startswith("*"):
                        factor = float(expr[1:])
                        df[col] = df[col] * factor
                    elif expr.startswith("/"):
                        divisor = float(expr[1:])
                        df[col] = df[col] / divisor
                    elif expr.startswith("+"):
                        addend = float(expr[1:])
                        df[col] = df[col] + addend
                    elif expr.startswith("-"):
                        subtrahend = float(expr[1:])
                        df[col] = df[col] - subtrahend
                except Exception:
                    pass
    return result


def _find_matching_calc(col, calc_map, value_cols):
    for vcol, expr in calc_map.items():
        if vcol in col:
            return expr
    return None
