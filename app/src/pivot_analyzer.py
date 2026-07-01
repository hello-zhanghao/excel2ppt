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
        task["是否计算"] = str(item.get("是否计算", "是")).strip() if item.get("是否计算") else "是"

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
    """加载数据，支持 Excel 和 CSV 文件"""
    from src.excel_reader import find_data_file, read_data_file
    
    join_parts = _parse_join_spec(data_source)

    if not join_parts:
        # 非 JOIN 情况：直接查找文件
        file_path = _resolve_data_path(data_source, config_dir)
        if not file_path or not os.path.exists(file_path):
            # 尝试模糊查找
            file_path = find_data_file(data_source, config_dir)
        if not file_path or not os.path.exists(file_path):
            return None
        return read_data_file(file_path, sheet_name)

    candidate_files = _collect_candidate_xlsx(config_dir)

    first_file, first_sheet = _resolve_join_table(join_parts[0]["left"], config_dir, candidate_files)
    if first_file is None:
        return None
    if first_sheet is None:
        first_sheet = join_parts[0].get("left_sheet") or sheet_name

    df = read_data_file(first_file, first_sheet)

    for jp in join_parts:
        right_file, right_sheet = _resolve_join_table(jp["right"], config_dir, candidate_files)
        if right_file is None:
            right_file = first_file
            right_sheet = jp["right"]
            # 移除扩展名
            for ext in (".xlsx", ".csv"):
                if right_sheet.lower().endswith(ext):
                    right_sheet = right_sheet[:-len(ext)]
                    break
            if not _sheet_exists(right_file, right_sheet):
                continue
        if right_sheet is None:
            right_sheet = jp.get("right_sheet")

        df_right = read_data_file(right_file, right_sheet)

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
    """解析 JOIN 表名，支持 Excel 和 CSV 文件"""
    from src.excel_reader import find_data_file, get_data_file_sheets
    
    # 先尝试直接查找
    file_path = find_data_file(table_name, config_dir)
    if file_path:
        # 确定 sheet 名称
        sheets = get_data_file_sheets(file_path)
        if sheets:
            # 尝试匹配 sheet
            clean_name = table_name
            for ext in (".xlsx", ".csv"):
                if clean_name.lower().endswith(ext):
                    clean_name = clean_name[:-len(ext)]
                    break
            # 匹配最相似的 sheet
            for s in sheets:
                if clean_name.lower() in s.lower() or s.lower() in clean_name.lower():
                    return file_path, s
            return file_path, sheets[0]  # 默认第一个 sheet
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
    """保留原有函数名，但改为收集所有数据文件（Excel + CSV）"""
    from src.excel_reader import get_candidate_data_files
    return get_candidate_data_files(config_dir)


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

    行维度_str = 行维度_str or ""
    列维度_str = 列维度_str or ""
    值字段_str = 值字段_str or ""
    聚合方式_str = 聚合方式_str or "sum"

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

    # 聚合后应用列名映射
    # 值映射配置是按值字段顺序的，按顺序映射到聚合后的值列
    raw_val_maps = [m.strip() for m in val_map_str.split(",") if m.strip()] if val_map_str else []
    
    if result and raw_val_maps:
        map_idx = 0  # 值映射的索引
        for key, df in result.items():
            if isinstance(df, pd.DataFrame):
                rename = {}
                for c in df.columns:
                    if c in 行维度:  # 行维度是分组列
                        continue
                    if map_idx < len(raw_val_maps):
                        rename[c] = raw_val_maps[map_idx]
                        map_idx += 1
                    else:
                        break  # 没有更多映射了
                
                if rename:
                    df.rename(columns=rename, inplace=True)

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
    
    # 处理重复字段名：对同一字段的多个聚合合并到一起
    # 正确做法：对于同名字段，把所有聚合函数合并到同一列表
    # pandas groupby().agg() 会返回 tuple 列名，如 ('销售额', 'sum'), ('销售额', 'mean')
    agg_dict = {}
    field_funcs = {}  # 记录每个字段的聚合函数列表
    for idx, vcol in enumerate(value_cols):
        # 直接使用索引获取对应的聚合函数（避免 per_field_funcs 字典重复键问题）
        if idx < len(agg_funcs):
            funcs_for_vcol = [agg_funcs[idx]]
        else:
            funcs_for_vcol = agg_funcs
        
        for a in funcs_for_vcol:
            a_mapped = AGG_MAP.get(a, a)
            if a_mapped in ("pct", "占比", "percentage"):
                if vcol not in field_funcs:
                    field_funcs[vcol] = []
                field_funcs[vcol].append("sum")
                has_pct = True
            else:
                if vcol not in field_funcs:
                    field_funcs[vcol] = []
                field_funcs[vcol].append(a_mapped)
    
    # 构建 agg_dict，每个字段只出现一次，聚合函数合并为列表
    for vcol, funcs in field_funcs.items():
        agg_dict[vcol] = funcs

    grouped = df.groupby(group_cols, as_index=False, observed=True).agg(agg_dict)

    # 检查是否有值映射配置
    val_map_str = task.get("值映射", "") if task else ""
    raw_val_maps = [m.strip() for m in val_map_str.split(",") if m.strip()] if val_map_str else []
    has_val_map = bool(raw_val_maps)
    
    # 如果有值映射，使用原始 key 作为列名（不加聚合后缀），后续步骤会应用映射
    # 如果没有值映射，使用 _flatten_col 添加聚合后缀
    orig_tuples = [col for col in grouped.columns.values]
    if has_val_map:
        # 有值映射时，直接在 _group_aggregate 中应用映射，不走后续重命名逻辑
        # 值映射按顺序对应值列（跳过分组列）
        new_columns = []
        val_map_idx = 0
        for col in orig_tuples:
            # 先判断是否是分组列（分组列的 tuple 第二个元素通常是空字符串）
            if isinstance(col, tuple):
                # 检查是否是分组列
                if len(col) >= 2 and str(col[1]).strip() == '':
                    # 这是分组列，保留原名
                    new_columns.append(str(col[0]).strip())
                elif val_map_idx < len(raw_val_maps):
                    # 值列，使用值映射
                    new_columns.append(raw_val_maps[val_map_idx])
                    val_map_idx += 1
                else:
                    new_columns.append(_agg_key_to_name(col))
            elif col in group_cols:
                new_columns.append(col)
            else:
                new_columns.append(col)
        grouped.columns = new_columns
    else:
        grouped.columns = [_flatten_col(col) for col in grouped.columns.values]

    # 处理重复列名（仅在没有值映射时添加聚合后缀区分）
    if not has_val_map:
        seen = {}
        for i, name in enumerate(grouped.columns):
            if name in seen and name not in group_cols:
                prev = seen[name]
                # 使用原始 tuple 中的聚合信息来区分
                curr_tuple = orig_tuples[i]
                prev_tuple = orig_tuples[prev]
                # 如果 tuple 包含聚合信息，使用它
                if isinstance(curr_tuple, tuple) and len(curr_tuple) >= 2:
                    agg_part = str(curr_tuple[1]).strip()
                    if agg_part:
                        grouped.columns.values[i] = f"{name}_{agg_part}"
                        continue
                if isinstance(prev_tuple, tuple) and len(prev_tuple) >= 2:
                    agg_part = str(prev_tuple[1]).strip()
                    if agg_part:
                        grouped.columns.values[prev] = f"{name}_{agg_part}"
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


def _agg_key_to_name(col):
    """将聚合后的列名 tuple 转为字符串，保留原始字段名（不加聚合后缀）"""
    if isinstance(col, tuple):
        # 取第一个元素（字段名），可能带序号后缀
        parts = [str(p).strip() for p in col if str(p).strip()]
        if parts:
            return parts[0]
        return str(col)
    return str(col).strip()


def _flatten_col(col):
    if isinstance(col, tuple):
        parts = [str(p).strip() for p in col if str(p).strip()]
        if len(parts) >= 2:
            field_name = parts[0]
            agg_func = parts[1] if len(parts) > 1 else ""
            
            # 还原原始字段名（去掉我们添加的 _N 后缀）
            import re
            m = re.match(r'^(.+)_(\d+)$', field_name)
            if m:
                # field_name 是 "销售额_1" 这样的格式，还原为 "销售额"
                field_name = m.group(1)
                suffix = m.group(2)  # 保留序号
            else:
                suffix = None
            
            # 如果聚合函数是 sum/mean 等英文，转为中文
            agg_cn = {"sum": "求和", "mean": "均值", "avg": "均值", "count": "计数",
                      "max": "最大值", "min": "最小值", "nunique": "去重"}.get(agg_func, agg_func)
            
            if suffix is not None:
                return f"{field_name}_{agg_cn}_{suffix}"
            return f"{field_name}_{agg_cn}"
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
            # 处理聚合后的列名（如 "销售额_求和"，需要映射到 "新名称_求和"）
            for c in df.columns:
                # 精确匹配原始列名
                if c == src:
                    rename[c] = dst
                    break
                # 匹配 "原列名_聚合后缀" 格式（如 "销售额_求和" -> "新名称_求和"）
                if c.startswith(src + "_"):
                    suffix = c[len(src):]  # 取 "_求和" 部分
                    rename[c] = dst + suffix
                    break
                # 部分匹配
                if c.startswith(src):
                    suffix = c[len(src):]
                    rename[c] = dst + suffix
                    break
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
