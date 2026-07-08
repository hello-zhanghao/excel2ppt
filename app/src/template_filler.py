"""
PPT 模板填充器 — 基于已有 PPT 模板和透视结果进行数据替换

用法：
    from src.template_filler import fill_template
    fill_template("模板.pptx", "透视结果.xlsx", "输出.pptx")

占位符语法：
    区块查找支持三种形式（图表/表格/图片占位符的"区块名"位置同样适用）：
        区块名                       跨所有 sheet 查找（兼容旧用法）
        Sheet名.区块名               精确指定 sheet 内的区块
        Sheet名                      该 sheet 第一个区块（兼容旧用法）

    文本占位符（在文本框中）：
        {{区块名.列名}}              取该列合计值（数值列）或首行值
        {{区块名.列名.行值}}         精确取某行某列的值
        {{区块名.行数}}              取该区块的数据行数
        {{标量.标量名}}              取无行维度标量
        {{Sheet名.区块名.列名}}      精确指定 sheet 内区块（同名区块不冲突）
        {{Sheet名.区块名.列名.行值}} 精确指定 sheet 内区块的某行

    图表占位符（在图表形状名称/替代文字中）：
        {{图表:区块名}}              用该区块数据替换图表数据源
        {{图表:Sheet名.区块名}}      精确指定 sheet 内区块
        {{图表:区块名|列1,列2}}      只使用指定列（第一列自动作为X轴类别）

    图片占位符（在图片替代文字/名称中，或在文本框中）：
        {{图片:文件路径}}            替换为指定图片（绝对路径或相对模板目录）
        {{图片:区块名.列名.行值}}    取透视数据中的图片路径

    表格占位符（在表格替代文字/名称中）：
        {{表格:区块名}}              用该区块数据整表替换（保留模板表格样式）
        {{表格:Sheet名.区块名}}      精确指定 sheet 内区块
        {{表格:区块名|列1,列2}}      只填指定列（第一列自动保留）

    聚合后缀（可选）：
        {{区块名.列名.sum}}          求和（默认）
        {{区块名.列名.avg}}          平均
        {{区块名.列名.max}}          最大
        {{区块名.列名.min}}          最小

PPT 备注配置（每页备注区可写）：
    数据源=透视结果.xlsx
    区块=按地区汇总                # 声明本页默认区块，占位符可省略前缀
"""
import os
import re
from typing import Dict, Optional, Tuple, List

import openpyxl
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE


# 占位符正则：{{...}}
_PLACEHOLDER_RE = re.compile(r"\{\{([^{}]+)\}\}")
# 图表占位符正则：{{图表:xxx}}
_CHART_PLACEHOLDER_RE = re.compile(r"\{\{图表[:：]([^{}]+)\}\}")
# 图片占位符正则：{{图片:xxx}}
_IMAGE_PLACEHOLDER_RE = re.compile(r"\{\{图片[:：]([^{}]+)\}\}")
# 表格占位符正则：{{表格:xxx}}
_TABLE_PLACEHOLDER_RE = re.compile(r"\{\{表格[:：]([^{}]+)\}\}")


def load_pivot_results(pivot_data_path: str) -> Dict[str, pd.DataFrame]:
    """加载透视结果 xlsx，返回 {区块名或sheet名: DataFrame}。

    一个 sheet 中可能包含多个区块（以区块标题行分隔），每个区块独立解析。
    - 区块标题行：只有第1列有值（区块名），其余列为空
    - 区块标题行的下一行是表头，再下面是数据行

    索引优先级：
    1. 区块名（精确匹配，跨所有 sheet 查找）
    2. sheet 名（指向该 sheet 的第一个区块，兼容旧用法）
    3. "Sheet名.区块名"（精确指定 sheet 内的区块，新增）
    """
    if not os.path.exists(pivot_data_path):
        raise FileNotFoundError(f"透视结果文件不存在: {pivot_data_path}")

    wb = openpyxl.load_workbook(pivot_data_path, data_only=True)
    result = {}
    # 记录每个 sheet 的第一个区块名，用于 sheet 名别名
    sheet_first_block = {}
    # 记录 (sheet_name, block_name) → DataFrame，用于 "Sheet名.区块名" 查找
    sheet_block_map = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        # 扫描所有行，按区块标题行分段
        # 区块标题行：只有第1列有值，其余列为空
        block_starts = []  # [(title_row_idx, block_name), ...]
        for i, row in enumerate(rows):
            non_empty = [c for c in row if c is not None and str(c).strip()]
            if len(non_empty) == 1:
                # 只有第1列有值 → 区块标题行
                block_name = str(rows[i][0]).strip()
                block_starts.append((i, block_name))

        if not block_starts:
            # 没有区块标题行：整个 sheet 当作一个区块，sheet 名作为区块名
            block_starts.append((0, sheet_name))

        first_block_in_sheet = True
        for idx, (title_row_idx, block_name) in enumerate(block_starts):
            # 区块标题行存在时，表头在下一行；否则表头在当前行
            if block_name == sheet_name and not _is_title_row(rows, title_row_idx):
                header_row_idx = title_row_idx
            else:
                header_row_idx = title_row_idx + 1

            # 数据行范围：表头行+1 到 下一个区块标题行-1（或 sheet 末尾）
            data_end = block_starts[idx + 1][0] if idx + 1 < len(block_starts) else len(rows)
            data_rows = rows[header_row_idx + 1:data_end]
            # 过滤空行
            data_rows = [r for r in data_rows if any(c is not None and str(c).strip() for c in r)]
            if not data_rows:
                continue
            if header_row_idx >= len(rows):
                continue

            headers = [str(c).strip() if c is not None else f"col_{i}" for i, c in enumerate(rows[header_row_idx])]
            df = pd.DataFrame(data_rows, columns=headers)

            # 用区块名作为 key（后出现的同名区块覆盖先出现的，与 pivot 输出语义一致）
            result[block_name] = df

            # 记录 (sheet_name, block_name) → DataFrame
            sheet_block_map[f"{sheet_name}.{block_name}"] = df

            # 记录每个 sheet 的第一个区块名
            if first_block_in_sheet:
                sheet_first_block[sheet_name] = block_name
                first_block_in_sheet = False

    # 用 sheet 名作为别名，指向该 sheet 的第一个区块（兼容旧用法）
    for sheet_name, first_block in sheet_first_block.items():
        if sheet_name not in result:
            result[sheet_name] = result.get(first_block)

    # 合并 "Sheet名.区块名" 形式的 key
    result.update(sheet_block_map)

    wb.close()
    return result


def _is_title_row(rows, row_idx):
    """判断指定行是否是区块标题行（只有第1列有值）"""
    if row_idx >= len(rows):
        return False
    row = rows[row_idx]
    non_empty = [c for c in row if c is not None and str(c).strip()]
    return len(non_empty) == 1


def _parse_value(value, default=""):
    """将单元格值转换为字符串，数值保留合理精度"""
    if value is None:
        return default
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return f"{value:.2f}"
    return str(value)


def _aggregate_column(df: pd.DataFrame, col: str, agg: str = "sum") -> Optional[float]:
    """对指定列做聚合"""
    if col not in df.columns:
        return None
    series = df[col]
    # 尝试转为数值
    numeric_series = pd.to_numeric(series, errors="coerce")
    if numeric_series.isna().all():
        return None
    agg = agg.lower().strip()
    if agg == "sum":
        return float(numeric_series.sum())
    elif agg in ("avg", "mean"):
        return float(numeric_series.mean())
    elif agg == "max":
        return float(numeric_series.max())
    elif agg == "min":
        return float(numeric_series.min())
    elif agg == "count":
        return int(numeric_series.count())
    return float(numeric_series.sum())


def _lookup_block(pivot_data: Dict[str, pd.DataFrame], block_expr: str) -> Optional[pd.DataFrame]:
    """按区块表达式查找 DataFrame，支持三种形式：
    1. "区块名"          → 跨 sheet 查找（兼容旧用法）
    2. "Sheet名.区块名"  → 精确指定 sheet 内的区块
    3. "Sheet名"         → 该 sheet 第一个区块（兼容旧用法）
    """
    if not block_expr:
        return None
    # 精确匹配（包括 "Sheet名.区块名" 形式）
    if block_expr in pivot_data:
        return pivot_data[block_expr]
    return None


def _parse_block_and_cols(expr: str) -> Tuple[str, Optional[List[str]]]:
    """解析占位符参数：'区块名|列1,列2,列3' → ('区块名', ['列1','列2','列3'])

    支持的格式：
        区块名                → (区块名, None)  全量列
        区块名|列1,列2        → (区块名, [列1,列2])  指定列
        Sheet名.区块名|列1    → (Sheet名.区块名, [列1])

    列名会去除首尾空白。无 | 时返回 (expr, None) 保持兼容。
    """
    expr = expr.strip()
    if "|" not in expr:
        return expr, None
    block_part, cols_part = expr.split("|", 1)
    block_part = block_part.strip()
    cols = [c.strip() for c in cols_part.split(",") if c.strip()]
    if not cols:
        return block_part, None
    return block_part, cols


def _filter_df_columns(df: pd.DataFrame, cols: Optional[List[str]]) -> pd.DataFrame:
    """按指定列筛选 DataFrame，保留第一列（行维度）+ 指定列。

    无 cols 或 cols 为 None 时返回原 df。
    指定列不存在时跳过，至少保留第一列。
    """
    if cols is None or df.empty:
        return df
    # 第一列是行维度，必须保留
    first_col = df.columns[0]
    keep = [first_col]
    for c in cols:
        if c in df.columns and c != first_col:
            keep.append(c)
    if not keep:
        return df
    return df[keep]


def _resolve_text_placeholder(expr: str, pivot_data: Dict[str, pd.DataFrame],
                              default_block: Optional[str] = None) -> str:
    """解析文本占位符表达式，返回替换值字符串

    支持的格式：
        区块名.列名
        区块名.列名.行值
        区块名.行数
        区块名.列名.聚合(sum/avg/max/min/count)
        标量.标量名
        Sheet名.区块名.列名              （精确指定 sheet 内区块，新增）
        Sheet名.区块名.列名.行值         （新增）
        Sheet名.区块名.列名.聚合         （新增）
        列名                             （使用 default_block）
        列名.行值                        （使用 default_block）

    查找优先级：先尝试 "Sheet名.区块名" 双段精确匹配，
    再回退到 "区块名" 单段跨 sheet 匹配。
    """
    expr = expr.strip()
    parts = expr.split(".")

    # 标量特殊处理
    if parts[0] == "标量":
        if len(parts) < 2:
            return ""
        scalar_name = parts[1]
        # 标量存放在无行维度的结果里（单行），遍历所有 sheet 查找
        for sheet_name, df in pivot_data.items():
            if scalar_name in df.columns and len(df) == 1:
                return _parse_value(df[scalar_name].iloc[0])
        return ""

    # 尝试 "Sheet名.区块名" 双段形式（至少3段：Sheet.区块.列）
    # 先用最长前缀匹配，找出 Sheet.区块 的边界
    block = None
    col = None
    row_val = None
    agg = None

    if len(parts) >= 3:
        # 尝试前两段作为 "Sheet名.区块名"
        candidate_two = f"{parts[0]}.{parts[1]}"
        if candidate_two in pivot_data:
            block = candidate_two
            col = parts[2]
            if len(parts) >= 4:
                fourth = parts[3]
                if fourth in ("sum", "avg", "mean", "max", "min", "count"):
                    agg = fourth
                    row_val = None
                else:
                    row_val = fourth
                    agg = parts[4] if len(parts) > 4 else None

    # 回退到旧逻辑：单段区块名
    if block is None:
        if len(parts) == 1:
            # 只有列名，依赖 default_block
            if not default_block:
                return ""
            col = parts[0]
            block = default_block
            row_val = None
            agg = None
        elif len(parts) == 2:
            # 区块名.列名 或 列名.行值/聚合
            if parts[0] in pivot_data:
                block, col = parts[0], parts[1]
                row_val = None
                agg = None
            elif default_block and parts[1] in ("sum", "avg", "mean", "max", "min", "count"):
                block = default_block
                col = parts[0]
                row_val = None
                agg = parts[1]
            elif default_block:
                block = default_block
                col = parts[0]
                row_val = parts[1]
                agg = None
            else:
                return ""
        elif len(parts) >= 3:
            block = parts[0]
            col = parts[1]
            third = parts[2]
            # 判断第三个是聚合还是行值
            if third in ("sum", "avg", "mean", "max", "min", "count"):
                agg = third
                row_val = None
            else:
                row_val = third
                agg = parts[3] if len(parts) > 3 else None
        else:
            return ""

    df = _lookup_block(pivot_data, block)
    if df is None:
        return ""

    # 特殊列名：行数
    if col == "行数":
        return str(len(df))

    if col not in df.columns:
        return ""

    # 按行值定位
    if row_val:
        # 第一列作为行维度
        row_dim_col = df.columns[0]
        matched = df[df[row_dim_col].astype(str) == row_val]
        if len(matched) == 0:
            return ""
        value = matched[col].iloc[0]
        return _parse_value(value)

    # 聚合
    if agg:
        result = _aggregate_column(df, col, agg)
        if result is None:
            return ""
        if isinstance(result, float):
            if result == int(result):
                return str(int(result))
            return f"{result:.2f}"
        return str(result)

    # 默认：数值列求和，非数值列取首行
    numeric_series = pd.to_numeric(df[col], errors="coerce")
    if not numeric_series.isna().all():
        return _parse_value(float(numeric_series.sum()))
    return _parse_value(df[col].iloc[0])


def _parse_slide_notes(notes_text: str) -> Dict[str, str]:
    """解析幻灯片备注区的配置"""
    config = {}
    for line in notes_text.split("\n"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            key, value = line.split("=", 1)
            config[key.strip()] = value.strip()
    return config


def _replace_in_text_frame(text_frame, pivot_data: Dict[str, pd.DataFrame],
                           default_block: Optional[str],
                           image_collector: Optional[List[str]] = None) -> int:
    """替换文本框中的占位符，返回替换次数。
    如果 image_collector 不为 None，{{图片:...}} 会被收集到列表中并从文本中移除。
    """
    replace_count = 0
    for para in text_frame.paragraphs:
        full_text = "".join(run.text for run in para.runs)
        if not _PLACEHOLDER_RE.search(full_text):
            continue

        def _replacer(m):
            nonlocal replace_count
            expr = m.group(1).strip()
            if expr.startswith("图表") or expr.startswith("图表:"):
                return m.group(0)
            if (expr.startswith("图片") or expr.startswith("图片:")) and image_collector is not None:
                image_collector.append(expr[2:].lstrip(":：").strip())
                replace_count += 1
                return ""
            value = _resolve_text_placeholder(expr, pivot_data, default_block)
            if value:
                replace_count += 1
            return value

        new_text = _PLACEHOLDER_RE.sub(_replacer, full_text)
        if new_text != full_text and para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""
    return replace_count


def _replace_chart_data(slide, pivot_data: Dict[str, pd.DataFrame],
                        default_block: Optional[str]) -> int:
    """替换幻灯片中的图表数据，返回替换次数

    图表占位符只从形状名称/替代文字读取，不修改图表标题（保留模板原标题）。
    """
    replace_count = 0
    for shape in slide.shapes:
        if not shape.has_chart:
            continue

        chart = shape.chart

        # 只从形状名称/替代文字读取占位符（不读取图表标题，保留模板原标题）
        target_expr = None
        try:
            for attr in ("name", "alternative_text"):
                val = (getattr(shape, attr, None) or "")
                m = _CHART_PLACEHOLDER_RE.search(val)
                if m:
                    target_expr = m.group(1).strip()
                    break
        except Exception:
            pass

        # 兼容：如果形状名称没有占位符，再回退到图表标题读取（旧模板兼容）
        if not target_expr:
            try:
                if chart.has_title and chart.chart_title.has_text_frame:
                    title_text = chart.chart_title.text_frame.text
                    m = _CHART_PLACEHOLDER_RE.search(title_text)
                    if m:
                        target_expr = m.group(1).strip()
            except Exception:
                pass

        if not target_expr:
            continue

        # 解析 "区块名|列1,列2" 语法
        target_block, cols = _parse_block_and_cols(target_expr)
        df = _lookup_block(pivot_data, target_block)
        if df is None:
            print(f"    [警告] 图表数据区块 '{target_block}' 未在透视结果中找到")
            continue
        if df.empty:
            continue
        # 按指定列筛选
        df = _filter_df_columns(df, cols)

        try:
            _write_chart_data(chart, df)
            replace_count += 1
            cols_info = f", 仅列: {cols}" if cols else ""
            print(f"    [OK] 图表数据替换: {target_block} ({df.shape[0]}行 x {df.shape[1]}列{cols_info})")
            # 清除形状名称中的占位符（不动图表标题）
            try:
                if shape.name and _CHART_PLACEHOLDER_RE.search(shape.name):
                    shape.name = _CHART_PLACEHOLDER_RE.sub(target_block, shape.name)
            except Exception:
                pass
        except Exception as e:
            print(f"    [警告] 图表数据替换失败 [{target_block}]: {e}")
    return replace_count


def _write_chart_data(chart, df: pd.DataFrame):
    """将 DataFrame 写入图表的内嵌 WorkBook"""
    from pptx.chart.data import CategoryChartData

    # 第一列作为类别（X轴），其余列作为系列（Y轴）
    categories = df.iloc[:, 0].astype(str).tolist()
    series_data = {}
    for col_idx in range(1, len(df.columns)):
        col_name = df.columns[col_idx]
        values = pd.to_numeric(df.iloc[:, col_idx], errors="coerce").fillna(0).tolist()
        series_data[col_name] = values

    chart_data = CategoryChartData()
    chart_data.categories = categories
    for name, values in series_data.items():
        chart_data.add_series(name, values)

    chart.replace_data(chart_data)


def _resolve_image_path(expr: str, pivot_data: Dict[str, pd.DataFrame],
                        default_block: Optional[str], template_dir: str) -> Optional[str]:
    """解析 {{图片:...}} 表达式，返回图片文件绝对路径或 None"""
    expr = expr.strip()
    if not expr:
        return None

    # 1) 绝对路径或相对模板目录的直接路径
    for candidate in (expr, os.path.join(template_dir, expr)):
        if os.path.isfile(candidate):
            return os.path.abspath(candidate)

    # 2) 透视数据解析（区块.列 或 区块.列.行值）
    text_value = _resolve_text_placeholder(expr, pivot_data, default_block)
    if text_value:
        for candidate in (text_value, os.path.join(template_dir, text_value)):
            if os.path.isfile(candidate):
                return os.path.abspath(candidate)

    return None


def _replace_pictures(slide, pivot_data: Dict[str, pd.DataFrame],
                      default_block: Optional[str],
                      template_dir: str,
                      text_image_exprs: Optional[List[str]] = None) -> int:
    """替换幻灯片中的图片，返回替换次数。

    匹配优先级：
    1. 图片形状的 name 或 alternative_text 中含 {{图片:...}}
    2. 文本框中收集到的 {{图片:...}} 表达式 → 匹配同页第一张未被其他方式匹配的图片
    """
    if text_image_exprs is None:
        text_image_exprs = []

    replaced = 0
    matched_shape_ids = set()

    for shape in list(slide.shapes):
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue

        expr = None
        for attr in ("name", "alternative_text"):
            val = (getattr(shape, attr, None) or "")
            m = _IMAGE_PLACEHOLDER_RE.search(val)
            if m:
                expr = m.group(1).strip()
                break

        if not expr:
            continue

        image_path = _resolve_image_path(expr, pivot_data, default_block, template_dir)
        if not image_path:
            print(f"    [警告] 图片路径无效 [{expr}]: 文件不存在")
            continue

        _do_replace_picture(slide, shape, image_path)
        matched_shape_ids.add(shape.shape_id)
        replaced += 1
        print(f"    [OK] 图片替换: {os.path.basename(image_path)}")

    # 文本关联模式：{{图片:...}} 在文本框中 → 替换同页第一张未被匹配的图片
    for expr in text_image_exprs:
        image_path = _resolve_image_path(expr, pivot_data, default_block, template_dir)
        if not image_path:
            continue
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            if shape.shape_id in matched_shape_ids:
                continue
            _do_replace_picture(slide, shape, image_path)
            matched_shape_ids.add(shape.shape_id)
            replaced += 1
            print(f"    [OK] 图片替换(文本关联): {os.path.basename(image_path)}")
            break

    return replaced


def _do_replace_picture(slide, old_shape, image_path: str):
    """删除旧图片，在原位置插入新图片"""
    left, top, width, height = old_shape.left, old_shape.top, old_shape.width, old_shape.height
    sp = old_shape._element
    sp.getparent().remove(sp)
    slide.shapes.add_picture(image_path, left, top, width, height)


def _expand_table_columns(table, target_col_count: int):
    """给表格添加列到目标列数，复制最后一列的样式"""
    current = len(table.columns)
    if target_col_count <= current:
        return
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tbl = table._tbl
    for row_idx, tr in enumerate(tbl.findall(f"{{{ns}}}tr")):
        # 取该行最后一格做模板
        last_tc = tr.findall(f"{{{ns}}}tc")[-1]
        for _ in range(target_col_count - current):
            new_tc = last_tc.__deepcopy__(True)
            # 清空新格的文本段落
            for p in new_tc.findall(f".//{{{ns}}}p"):
                for r in p.findall(f"{{{ns}}}r"):
                    r.text = ""
            tr.append(new_tc)
    # 刷新 table 的网格列定义
    grid = tbl.find(f"{{{ns}}}tblGrid")
    if grid is not None:
        last_gc = grid.findall(f"{{{ns}}}gridCol")[-1]
        for _ in range(target_col_count - current):
            grid.append(last_gc.__deepcopy__(True))


def _expand_table_rows(table, target_row_count: int):
    """给表格添加行到目标行数（含表头），复制最后一行的样式"""
    current = len(table.rows)
    if target_row_count <= current:
        return
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tbl = table._tbl
    tr_list = tbl.findall(f"{{{ns}}}tr")
    last_tr = tr_list[-1]
    for _ in range(target_row_count - current):
        new_tr = last_tr.__deepcopy__(True)
        # 清空新行所有单元格的文本
        for tc in new_tr.findall(f"{{{ns}}}tc"):
            for p in tc.findall(f".//{{{ns}}}p"):
                for r in p.findall(f"{{{ns}}}r"):
                    r.text = ""
        tbl.append(new_tr)


def _clear_table_row(table, row_idx: int):
    """清空表格某行的所有单元格文本"""
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tbl = table._tbl
    tr_list = tbl.findall(f"{{{ns}}}tr")
    if row_idx >= len(tr_list):
        return
    tr = tr_list[row_idx]
    for tc in tr.findall(f"{{{ns}}}tc"):
        for p in tc.findall(f".//{{{ns}}}p"):
            for r in p.findall(f"{{{ns}}}r"):
                r.text = ""


def _set_cell_text_preserve_format(cell, text: str):
    """替换单元格文本但保留模板原有格式（字体/颜色/对齐/填充）。

    python-pptx 的 cell.text setter 会清空段落重建 run，导致格式丢失。
    本函数保留第一个段落的第一个 run，只改其文本；多余 run 清空文本。
    """
    tf = cell.text_frame
    paragraphs = tf.paragraphs
    if not paragraphs:
        cell.text = text
        return
    # 用第一个段落承载文本
    first_para = paragraphs[0]
    runs = first_para.runs
    if runs:
        # 第一个 run 写入新文本
        runs[0].text = text
        # 其余 run 清空（避免拼接出旧文本）
        for r in runs[1:]:
            r.text = ""
    else:
        # 段落没有 run（空段落），直接用 add_run（会继承段落默认格式）
        first_para.add_run().text = text
    # 多余段落清空文本（保留段落本身的格式属性）
    for p in paragraphs[1:]:
        for r in p.runs:
            r.text = ""


def _replace_table_data(slide, pivot_data: Dict[str, pd.DataFrame],
                        default_block: Optional[str]) -> int:
    """替换幻灯片中的表格数据（整表替换），返回替换次数。
    通过表格形状的 name 或 alternative_text 中的 {{表格:区块名}} 匹配。
    自动扩展表格行列以容纳完整数据，数据少于模板时清空多余行。
    """
    replace_count = 0
    for shape in list(slide.shapes):
        if not shape.has_table:
            continue

        target_expr = None
        for attr in ("name", "alternative_text"):
            val = (getattr(shape, attr, None) or "")
            m = _TABLE_PLACEHOLDER_RE.search(val)
            if m:
                target_expr = m.group(1).strip()
                break

        if not target_expr:
            continue

        # 解析 "区块名|列1,列2" 语法
        target_block, cols = _parse_block_and_cols(target_expr)
        df = _lookup_block(pivot_data, target_block)
        if df is None:
            print(f"    [警告] 表格数据区块 '{target_block}' 未在透视结果中找到")
            continue
        if df.empty:
            continue
        # 按指定列筛选
        df = _filter_df_columns(df, cols)

        table = shape.table
        headers = list(df.columns)
        data_rows = df.values.tolist()

        need_rows = 1 + len(data_rows)  # 表头 + 数据行
        need_cols = len(headers)

        try:
            # 自动扩展
            if need_cols > len(table.columns):
                _expand_table_columns(table, need_cols)
            if need_rows > len(table.rows):
                _expand_table_rows(table, need_rows)

            # 填表头（第一行）——保留模板单元格格式（字体/颜色/对齐/填充）
            for col_idx, header in enumerate(headers):
                cell = table.cell(0, col_idx)
                _set_cell_text_preserve_format(cell, str(header))

            # 填数据行
            for row_idx, row_data in enumerate(data_rows):
                table_row = row_idx + 1
                for col_idx, value in enumerate(row_data):
                    cell = table.cell(table_row, col_idx)
                    _set_cell_text_preserve_format(cell, _parse_value(value))

            # 数据少于模板行数时，清空多余行
            for extra_row in range(need_rows, len(table.rows)):
                _clear_table_row(table, extra_row)

            replace_count += 1
            cols_info = f", 仅列: {cols}" if cols else ""
            print(f"    [OK] 表格数据替换: {target_block} ({len(headers)}列 x {len(data_rows)}行{cols_info})")
            # 清除表格形状名称/替代文字中的占位符
            try:
                if shape.name and _TABLE_PLACEHOLDER_RE.search(shape.name):
                    shape.name = _TABLE_PLACEHOLDER_RE.sub(target_block, shape.name)
            except Exception:
                pass
        except Exception as e:
            print(f"    [警告] 表格数据替换失败 [{target_block}]: {e}")
    return replace_count


def fill_template(template_path: str, pivot_data_path: str, output_path: str) -> Dict:
    """填充 PPT 模板

    Args:
        template_path: PPT 模板文件路径
        pivot_data_path: 透视结果 xlsx 路径
        output_path: 输出 PPT 路径

    Returns:
        dict: 替换统计 {slides, text_replacements, chart_replacements, picture_replacements, table_replacements}
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"模板文件不存在: {template_path}")

    template_dir = os.path.dirname(os.path.abspath(template_path))

    print(f"[模板/1] 加载透视结果: {os.path.basename(pivot_data_path)}")
    pivot_data = load_pivot_results(pivot_data_path)
    print(f"    → 共 {len(pivot_data)} 个数据区块: {', '.join(pivot_data.keys())}")

    print(f"[模板/2] 加载模板: {os.path.basename(template_path)}")
    prs = Presentation(template_path)
    print(f"    → 共 {len(prs.slides)} 页")

    total_text = 0
    total_chart = 0
    total_picture = 0
    total_table = 0

    for slide_idx, slide in enumerate(prs.slides, 1):
        default_block = None
        try:
            if slide.has_notes_slide:
                notes_text = slide.notes_slide.notes_text_frame.text
                config = _parse_slide_notes(notes_text)
                default_block = config.get("区块")
        except Exception:
            pass

        image_collector: List[str] = []
        text_count = 0
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text_count += _replace_in_text_frame(shape.text_frame, pivot_data, default_block, image_collector)

        chart_count = _replace_chart_data(slide, pivot_data, default_block)

        picture_count = _replace_pictures(slide, pivot_data, default_block, template_dir, image_collector)

        table_count = _replace_table_data(slide, pivot_data, default_block)

        total_text += text_count
        total_chart += chart_count
        total_picture += picture_count
        total_table += table_count
        print(f"    [页{slide_idx}] 文本替换: {text_count}, 图表替换: {chart_count}, 图片替换: {picture_count}, 表格替换: {table_count}")

    print(f"[模板/3] 保存: {output_path}")
    prs.save(output_path)

    stats = {
        "slides": len(prs.slides),
        "text_replacements": total_text,
        "chart_replacements": total_chart,
        "picture_replacements": total_picture,
        "table_replacements": total_table,
    }
    print(f"\n[OK] 模板填充完成！共 {stats['slides']} 页, 文本替换 {total_text} 处, 图表替换 {total_chart} 处, 图片替换 {total_picture} 处, 表格替换 {total_table} 处")
    return stats
