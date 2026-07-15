import os
import sys
import json
import base64
from datetime import datetime
from typing import List, Dict, Optional, Tuple

HUAWEI_PALETTE = [
    "#C8102E", "#182B49", "#5B9BD5", "#ED7D31", 
    "#70AD47", "#A5A5A5", "#FFC000", "#4472C4"
]

def read_excel_data(excel_path: str) -> List[Dict]:
    try:
        import openpyxl
    except ImportError:
        print("[警告] 缺少 openpyxl，无法读取 Excel 数据")
        return []

    if not os.path.exists(excel_path):
        return []

    sections = []
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        header_idx = 0
        for i, row in enumerate(rows[:3]):
            non_none = [c for c in row if c is not None]
            if len(non_none) >= 2:
                header_idx = i
                break

        headers = [str(c) if c is not None else "" for c in rows[header_idx]]
        col_indices = [i for i, h in enumerate(headers) if h.strip()]
        headers = [headers[i] for i in col_indices]

        data_rows = []
        for row in rows[header_idx + 1:]:
            vals = []
            for i in col_indices:
                v = row[i] if i < len(row) else None
                if v is None:
                    vals.append("")
                elif isinstance(v, float):
                    vals.append(f"{v:.4f}" if v != int(v) else str(int(v)))
                else:
                    vals.append(str(v))
            if any(v.strip() for v in vals):
                data_rows.append(vals)

        sections.append({
            "name": sheet_name,
            "headers": headers,
            "rows": data_rows,
        })
    wb.close()
    return sections


def _detect_geo_columns(headers: List[str]) -> Optional[Tuple[int, int]]:
    """检测经纬度列。返回 (lat_col_idx, lon_col_idx) 或 None。"""
    lat_keys = ["纬度", "lat", "latitude"]
    lon_keys = ["经度", "lon", "lng", "longitude"]
    lat_idx = None
    lon_idx = None
    for i, h in enumerate(headers):
        h_lower = h.lower().strip()
        if lat_idx is None and any(k in h_lower for k in lat_keys):
            lat_idx = i
        if lon_idx is None and any(k in h_lower for k in lon_keys):
            lon_idx = i
    if lat_idx is not None and lon_idx is not None:
        return (lat_idx, lon_idx)
    return None


def _infer_chart_type(headers: List[str], rows: List[List[str]]) -> str:
    # 优先：含经纬度列 → 地图
    if _detect_geo_columns(headers) is not None:
        return "map"

    numeric_cols = 0
    category_cols = 0
    for h in headers:
        h_lower = h.lower()
        if any(key in h_lower for key in ["金额", "数量", "销售额", "利润", "成本", "值", "计数", "总计", "合计"]):
            numeric_cols += 1
        else:
            category_cols += 1

    total_rows = len(rows)

    for h in headers:
        h_lower = h.lower()
        if any(key in h_lower for key in ["占比", "百分比", "pct"]):
            return "pie"
        if any(key in h_lower for key in ["日期", "月份", "时间", "季度", "年"]):
            return "line"

    if total_rows > 20:
        return "table"

    if numeric_cols >= 2 and category_cols == 1:
        return "bar"

    if numeric_cols == 1 and category_cols == 1:
        return "bar"

    if numeric_cols == 1 and category_cols >= 2:
        return "bar"

    return "table"


def _compute_summary(headers: List[str], rows: List[List[str]]) -> Dict:
    summary = {}
    numeric_indices = []
    
    for i, h in enumerate(headers):
        h_lower = h.lower()
        if any(key in h_lower for key in ["金额", "数量", "销售额", "利润", "成本", "值", "占比", "百分比", "pct", "计数"]):
            numeric_indices.append(i)
    
    for idx in numeric_indices:
        col_name = headers[idx]
        values = []
        for row in rows:
            if idx < len(row) and row[idx] != "":
                try:
                    values.append(float(row[idx]))
                except ValueError:
                    pass
        
        if values:
            summary[col_name] = {
                "sum": round(sum(values), 2),
                "avg": round(sum(values) / len(values), 2),
                "max": round(max(values), 2),
                "min": round(min(values), 2),
                "count": len(values),
            }
    
    for row in rows:
        if any("合计" in str(v) or "总计" in str(v) for v in row):
            for i, v in enumerate(row):
                if i < len(headers):
                    try:
                        val = float(v) if v and str(v) != "" else None
                        if val is not None:
                            summary[f"合计_{headers[i]}"] = val
                    except ValueError:
                        pass
    
    return summary


def _escape_html(text: str) -> str:
    if text is None:
        return ""
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def _generate_chart_options(chart_id: str, data: Dict, chart_type: str) -> str:
    headers = data["headers"]
    rows = data["rows"]
    title = data.get("title", "")
    
    if not rows or not headers:
        return ""
    
    category_col = None
    value_cols = []
    
    for i, h in enumerate(headers):
        h_lower = h.lower()
        if any(key in h_lower for key in ["日期", "月份", "时间", "区域", "产品", "类别", "类型", "名称", "城市", "年份", "季度"]):
            if category_col is None:
                category_col = i
        elif any(key in h_lower for key in ["金额", "数量", "销售额", "利润", "成本", "值", "占比", "百分比", "pct", "计数"]):
            value_cols.append(i)
    
    if category_col is None:
        category_col = 0
        value_cols = [i for i in range(1, len(headers))]
    
    categories = []
    series_data = []
    
    for row in rows:
        if category_col < len(row):
            categories.append(str(row[category_col]))
    
    for vc in value_cols:
        values = []
        for row in rows:
            if vc < len(row) and row[vc] != "":
                try:
                    values.append(float(row[vc]))
                except ValueError:
                    values.append(0)
            else:
                values.append(0)
        series_data.append({
            "name": headers[vc],
            "data": values,
        })
    
    if chart_type == "pie":
        if series_data:
            pie_data = []
            for i, cat in enumerate(categories):
                pie_data.append({
                    "name": cat,
                    "value": series_data[0]["data"][i] if i < len(series_data[0]["data"]) else 0,
                })
            return json.dumps({
                "title": {"text": title, "left": "center", "textStyle": {"fontSize": 14}},
                "tooltip": {"trigger": "item", "formatter": "{b}: {c} ({d}%)"},
                "legend": {"orient": "vertical", "left": "left"},
                "series": [{
                    "name": title,
                    "type": "pie",
                    "radius": ["40%", "70%"],
                    "center": ["50%", "55%"],
                    "avoidLabelOverlap": True,
                    "itemStyle": {"borderRadius": 4, "borderColor": "#fff", "borderWidth": 2},
                    "label": {"show": True},
                    "data": pie_data,
                    "color": HUAWEI_PALETTE,
                }],
            })
    
    if chart_type == "line":
        series = []
        for i, sd in enumerate(series_data):
            series.append({
                "name": sd["name"],
                "type": "line",
                "data": sd["data"],
                "smooth": True,
                "symbol": "circle",
                "symbolSize": 8,
                "lineStyle": {"width": 3},
                "itemStyle": {"color": HUAWEI_PALETTE[i % len(HUAWEI_PALETTE)]},
            })
        return json.dumps({
            "title": {"text": title, "left": "center", "textStyle": {"fontSize": 14}},
            "tooltip": {"trigger": "axis"},
            "legend": {"data": [sd["name"] for sd in series_data], "bottom": 0},
            "grid": {"left": "3%", "right": "4%", "bottom": "15%", "top": "15%", "containLabel": True},
            "xAxis": {"type": "category", "data": categories, "axisLabel": {"rotate": 30}},
            "yAxis": {"type": "value"},
            "series": series,
        })
    
    if chart_type == "bar":
        series = []
        for i, sd in enumerate(series_data):
            series.append({
                "name": sd["name"],
                "type": "bar",
                "data": sd["data"],
                "barWidth": "50%",
                "itemStyle": {"color": HUAWEI_PALETTE[i % len(HUAWEI_PALETTE)], "borderRadius": [4, 4, 0, 0]},
            })
        return json.dumps({
            "title": {"text": title, "left": "center", "textStyle": {"fontSize": 14}},
            "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}},
            "legend": {"data": [sd["name"] for sd in series_data], "bottom": 0},
            "grid": {"left": "3%", "right": "4%", "bottom": "15%", "top": "15%", "containLabel": True},
            "xAxis": {"type": "category", "data": categories, "axisLabel": {"rotate": 30}},
            "yAxis": {"type": "value"},
            "series": series,
        })

    return ""


def _generate_geo_chart_options(chart_id: str, data: Dict, chart_type: str) -> str:
    """生成地图类型 chartOptions（ECharts geo + scatter）。

    chart_type: 'map'（散点地图）或 'heatmap'（热力地图）。
    数据需含经纬度列（lat/纬度、lon/经度），其余数值列作为指标。
    """
    headers = data["headers"]
    rows = data["rows"]
    title = data.get("title", "")

    if not rows or not headers:
        return ""

    geo = _detect_geo_columns(headers)
    if not geo:
        return ""

    lat_idx, lon_idx = geo
    # 指标列：除经纬度外的数值列
    metric_indices = []
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx):
            continue
        h_lower = h.lower()
        if any(k in h_lower for k in ["金额", "数量", "销售额", "利润", "成本", "值", "占比", "百分比", "pct", "计数", "总计", "合计"]):
            metric_indices.append(i)
    # 没有识别到指标列时，取经纬度外的第一个数值列
    if not metric_indices:
        for i, h in enumerate(headers):
            if i in (lat_idx, lon_idx):
                continue
            try:
                if rows:
                    float(rows[0][i])
                    metric_indices.append(i)
                    break
            except (ValueError, IndexError):
                continue

    # 名称列（站点名等）：经纬度和指标外的第一列
    name_idx = None
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx) or i in metric_indices:
            continue
        name_idx = i
        break

    # 收集数据点
    points = []
    for row in rows:
        try:
            lat = float(row[lat_idx])
            lon = float(row[lon_idx])
        except (ValueError, IndexError):
            continue
        if lat == 0 and lon == 0:
            continue
        name = str(row[name_idx]) if name_idx is not None and name_idx < len(row) else ""
        values = [lon, lat]
        for mi in metric_indices:
            try:
                values.append(float(row[mi]))
            except (ValueError, IndexError):
                values.append(0)
        points.append({"name": name, "value": values})

    if not points:
        return ""

    # 计算经纬度范围用于 geo 中心点和缩放
    lats = [p["value"][1] for p in points]
    lons = [p["value"][0] for p in points]
    center_lon = (min(lons) + max(lons)) / 2
    center_lat = (min(lats) + max(lats)) / 2
    span = max(max(lons) - min(lons), max(lats) - min(lats), 0.1)
    import math
    zoom = round(max(1, min(15, 8 - math.log10(span + 1))), 2)

    metric_names = [headers[mi] for mi in metric_indices] or ["指标"]
    metric_idx_in_value = 2  # value 数组中指标的起始位置

    # 散点大小归一化（用第一个指标列）
    if metric_indices:
        m_vals = [p["value"][metric_idx_in_value] for p in points]
        m_min, m_max = min(m_vals), max(m_vals)
        if m_max == m_min:
            sizes = [20] * len(points)
        else:
            sizes = [int(10 + (v - m_min) / (m_max - m_min) * 40) for v in m_vals]
    else:
        sizes = [20] * len(points)

    # tooltip 的 formatter 是 JS 函数，不能放进 json，改用模板字符串拼装整个选项
    tooltip_fn_parts = ['function(p){var s=p.name+"<br/>";var lon=p.value[0],lat=p.value[1];',
                        's+="经度:"+lon.toFixed(4)+", 纬度:"+lat.toFixed(4)+"<br/>";']
    for i, mn in enumerate(metric_names):
        tooltip_fn_parts.append(f's+="{mn}:"+p.value[{i+2}]+"<br/>";')
    tooltip_fn_parts.append('return s;}')
    tooltip_fn = "".join(tooltip_fn_parts)

    series_type = "scatter" if chart_type == "map" else "effectScatter"
    item_style = {"color": HUAWEI_PALETTE[0], "opacity": 0.85,
                  "borderColor": "#333", "borderWidth": 0.5} if chart_type == "map" else \
                 {"color": HUAWEI_PALETTE[3], "shadowBlur": 10,
                  "shadowColor": "rgba(237, 125, 49, 0.5)"}
    extra_series = {}
    if chart_type == "heatmap":
        extra_series = {"showEffectOn": "render", "rippleEffect": {"brushType": "stroke"}}

    # 拼装 option JSON，formatter 用 JS 函数字面量（不能 json.dumps）
    opt = {
        "title": {"text": title, "left": "center", "textStyle": {"fontSize": 14}},
        "tooltip": {"trigger": "item"},  # formatter 稍后手动替换
        "geo": {
            "map": "china",
            "roam": True,
            "zoom": zoom,
            "center": [center_lon, center_lat],
            "itemStyle": {"areaColor": "#F2F0EB", "borderColor": "#999"},
            "emphasis": {"itemStyle": {"areaColor": "#DCE6F0"}, "label": {"show": False}},
        },
        "series": [{
            "name": metric_names[0],
            "type": series_type,
            "coordinateSystem": "geo",
            "data": points,
            "symbolSize": sizes,
            "itemStyle": item_style,
            "label": {"show": True, "formatter": "{b}", "position": "right",
                      "fontSize": 10, "color": "#182B49"},
            **extra_series,
        }],
    }
    opt_json = json.dumps(opt, ensure_ascii=False)
    # 在 tooltip 对象内插入 formatter 函数
    opt_json = opt_json.replace('"tooltip": {"trigger": "item"}',
                                f'"tooltip": {{"trigger": "item", "formatter": {tooltip_fn}}}')
    return opt_json


def _build_geo_data_js(sec: Dict) -> str:
    """序列化地图 section 的完整数据，供 JS 端动态构建图表（支持指标列切换+筛选）。

    返回 JS 字符串：geoData['sectionId'] = {...}
    """
    headers = sec["headers"]
    rows = sec["rows"]
    geo = _detect_geo_columns(headers)
    if not geo:
        return ""

    lat_idx, lon_idx = geo
    # 指标列候选：除经纬度外的数值列
    metric_indices = []
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx):
            continue
        h_lower = h.lower()
        if any(k in h_lower for k in ["金额", "数量", "销售额", "利润", "成本", "值", "占比", "百分比", "pct", "计数", "总计", "合计"]):
            metric_indices.append(i)
    if not metric_indices:
        for i, h in enumerate(headers):
            if i in (lat_idx, lon_idx):
                continue
            try:
                if rows:
                    float(rows[0][i])
                    metric_indices.append(i)
            except (ValueError, IndexError):
                continue

    # 筛选列候选：非经纬度、非指标、非名称的列（通常是分类列）
    name_idx = None
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx) or i in metric_indices:
            continue
        if name_idx is None:
            name_idx = i
    # 筛选列：所有非数值列（排除经纬度）
    filter_indices = []
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx) or i in metric_indices:
            continue
        filter_indices.append(i)

    # 收集每个筛选列的唯一值
    filter_values = {}
    for fi in filter_indices:
        vals = set()
        for row in rows:
            if fi < len(row) and row[fi] != "":
                vals.add(str(row[fi]))
        # 唯一值超过 50 个时不作为筛选列（太多无意义）
        if len(vals) <= 50:
            filter_values[fi] = sorted(vals)

    data_obj = {
        "headers": headers,
        "rows": rows,
        "latIdx": lat_idx,
        "lonIdx": lon_idx,
        "nameIdx": name_idx,
        "metricIndices": metric_indices,
        "filterIndices": list(filter_values.keys()),
        "filterValues": {str(k): v for k, v in filter_values.items()},
    }
    return f"  geoData['{sec['id']}'] = {json.dumps(data_obj, ensure_ascii=False)};\n"


def _build_geo_controls_html(sec: Dict) -> str:
    """构建地图 section 的控件面板 HTML（指标列下拉 + 筛选列下拉 + 筛选值）。"""
    headers = sec["headers"]
    geo = _detect_geo_columns(headers)
    if not geo:
        return ""
    lat_idx, lon_idx = geo

    # 指标列选项
    metric_indices = []
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx):
            continue
        h_lower = h.lower()
        if any(k in h_lower for k in ["金额", "数量", "销售额", "利润", "成本", "值", "占比", "百分比", "pct", "计数", "总计", "合计"]):
            metric_indices.append(i)
    if not metric_indices:
        for i, h in enumerate(headers):
            if i in (lat_idx, lon_idx):
                continue
            try:
                if sec["rows"]:
                    float(sec["rows"][0][i])
                    metric_indices.append(i)
            except (ValueError, IndexError):
                continue

    metric_options = "".join(
        f'<option value="{mi}">{_escape_html(headers[mi])}</option>' for mi in metric_indices
    )

    # 筛选列选项（非数值列，排除经纬度）
    filter_indices = []
    for i, h in enumerate(headers):
        if i in (lat_idx, lon_idx) or i in metric_indices:
            continue
        # 唯一值不超过 50 个才作为筛选列
        vals = set()
        for row in sec["rows"]:
            if i < len(row) and row[i] != "":
                vals.add(str(row[i]))
        if len(vals) <= 50:
            filter_indices.append(i)
    filter_options = '<option value="">-- 不筛选 --</option>' + "".join(
        f'<option value="{fi}">{_escape_html(headers[fi])}</option>' for fi in filter_indices
    )

    sid = sec["id"]
    return f'''<div class="geo-controls" id="geo_ctrl_{sid}" style="display:none">
  <div class="geo-ctrl-item"><label>指标列:</label>
    <select id="geo_metric_{sid}" onchange="onGeoMetricChange('{sid}')">{metric_options}</select>
  </div>
  <div class="geo-ctrl-item"><label>筛选列:</label>
    <select id="geo_filter_col_{sid}" onchange="onGeoFilterColChange('{sid}')">{filter_options}</select>
  </div>
  <div class="geo-ctrl-item" id="geo_filter_vals_wrap_{sid}" style="display:none">
    <label>筛选值:</label><div class="geo-filter-vals" id="geo_filter_vals_{sid}"></div>
  </div>
</div>'''


def _ppt_chart_type_to_html(ppt_type: str) -> str:
    type_map = {
        "column": "bar",
        "bar": "bar",
        "line": "line",
        "pie": "pie",
        "doughnut": "pie",
        "area": "line",
        "scatter": "bar",
        "table": "table",
    }
    return type_map.get(ppt_type.lower(), "bar")


def _find_data_in_excel(excel_sections: List[Dict], sheet_name: str, block_name: str = None) -> Optional[Dict]:
    for sec in excel_sections:
        if sec["name"] == sheet_name:
            return sec
    return None


def generate_html_report(
    excel_path: Optional[str] = None,
    ppt_path: Optional[str] = None,
    ppt_config: Optional[List[Dict]] = None,
    output_dir: Optional[str] = None,
    report_title: str = "数据分析报告",
    report_subtitle: str = "透视分析结果",
) -> Optional[str]:
    if not output_dir:
        output_dir = os.path.dirname(excel_path) if excel_path else "."
    os.makedirs(output_dir, exist_ok=True)

    excel_sections = []
    if excel_path and os.path.exists(excel_path):
        print("[HTML] 读取 Excel 数据...")
        excel_sections = read_excel_data(excel_path)

    sections = []

    if ppt_config:
        print("[HTML] 从 PPT 配置生成报告结构...")
        
        for page_idx, page in enumerate(ppt_config):
            page_title_full = page.get("页面标题", "")
            if "|" in page_title_full:
                title_parts = page_title_full.split("|", 1)
                page_title = title_parts[0].strip()
                page_subtitle = title_parts[1].strip()
            else:
                page_title = page_title_full
                page_subtitle = ""

            if page.get("页面类型") == "封面":
                continue

            charts = page.get("charts", [])
            if not charts:
                continue

            for chart_idx, chart in enumerate(charts):
                sheet_name = chart.get("数据Sheet", "")
                block_name = chart.get("区块名", "")
                data_sec = _find_data_in_excel(excel_sections, sheet_name, block_name)

                if not data_sec and excel_sections:
                    if page_idx == 0 and chart_idx == 0:
                        print(f"    [警告] 未找到 Sheet '{sheet_name}'，尝试使用第一个可用数据")
                    data_sec = excel_sections[0] if excel_sections else None

                if not data_sec:
                    continue

                chart_type = _ppt_chart_type_to_html(chart.get("图表类型", "column"))
                
                title = chart.get("图表标题", "") or page_title
                summary = _compute_summary(data_sec["headers"], data_sec["rows"])
                
                sections.append({
                    "id": f"section_{page_idx}_{chart_idx}",
                    "title": title,
                    "subtitle": page_subtitle if chart_idx == 0 else "",
                    "chart_type": chart_type,
                    "available_charts": ["bar", "line", "pie", "table"],
                    "headers": data_sec["headers"],
                    "rows": data_sec["rows"],
                    "summary": summary,
                })

    if not sections and excel_sections:
        print("[HTML] 无 PPT 配置，从 Excel 自动生成...")
        for sec in excel_sections:
            if sec["name"] == "错误信息":
                continue

            chart_type = _infer_chart_type(sec["headers"], sec["rows"])
            summary = _compute_summary(sec["headers"], sec["rows"])

            # 含经纬度列时追加地图类型
            has_geo = _detect_geo_columns(sec["headers"]) is not None
            if has_geo:
                available = ["map", "heatmap", "table"]
            else:
                available = ["bar", "line", "pie", "table"]

            sections.append({
                "id": f"section_{len(sections)}",
                "title": sec["name"],
                "subtitle": "",
                "chart_type": chart_type,
                "available_charts": available,
                "headers": sec["headers"],
                "rows": sec["rows"],
                "summary": summary,
            })

    if ppt_path and os.path.exists(ppt_path):
        print("[HTML] 正在将 PPT 转为图片...")
        try:
            import win32com.client
            import pythoncom
            
            pythoncom.CoInitialize()
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = 0
            pres = ppt_app.Presentations.Open(os.path.abspath(ppt_path), WithWindow=False)
            
            ppt_images = []
            for i, slide in enumerate(pres.Slides, 1):
                img_path = os.path.join(output_dir, f"slide_{i:02d}.png")
                slide.Export(img_path, "PNG", 1280, 720)
                with open(img_path, "rb") as f:
                    img_b64 = base64.b64encode(f.read()).decode("utf-8")
                ppt_images.append(img_b64)
            
            pres.Close()
            ppt_app.Quit()
            pythoncom.CoUninitialize()
            
            if ppt_images:
                ppt_section = {
                    "id": "section_ppt",
                    "title": "PPT 预览",
                    "subtitle": "点击图片可缩放查看",
                    "chart_type": "ppt",
                    "available_charts": [],
                    "headers": [],
                    "rows": [],
                    "summary": {},
                    "ppt_images": ppt_images,
                }
                sections.insert(0, ppt_section)
        except Exception as e:
            print(f"[警告] PPT转图片失败: {e}")

    if not sections:
        print("[错误] 没有可用数据")
        return None

    all_summary = {}
    for sec in sections:
        if sec.get("summary"):
            for key, val in sec["summary"].items():
                if isinstance(val, dict):
                    all_summary[key] = val
                else:
                    all_summary[key] = val

    section_html_parts = []
    toc_html_parts = []
    all_chart_options_js = ""  # 累积所有 section 的 chartOptions，避免循环内重置
    all_geo_data_js = ""  # 累积地图 section 的完整数据，供 JS 动态构建

    for sec in sections:
        toc_html_parts.append(f'<li><a href="#{sec["id"]}">{_escape_html(sec["title"])}</a></li>')
        
        if sec.get("ppt_images"):
            slide_html = ""
            for idx, img_b64 in enumerate(sec["ppt_images"], 1):
                slide_html += f'''
    <div class="slide-card">
      <div class="slide-num">第 {idx} 页</div>
      <img src="data:image/png;base64,{img_b64}" alt="Slide {idx}" onclick="toggleZoom(this)">
    </div>'''
            section_html_parts.append(f'''
<div id="{sec["id"]}" class="section">
  <h2>{_escape_html(sec["title"])}</h2>
  <p class="subtitle">{_escape_html(sec["subtitle"])}</p>
  {slide_html}
</div>''')
            continue
        
        chart_options_js = ""
        is_geo_section = "map" in sec["available_charts"] or "heatmap" in sec["available_charts"]
        for ct in sec["available_charts"]:
            if ct in ("map", "heatmap"):
                # 地图类型：预生成默认 option 作为初始渲染，同时注册 geoData 供动态切换
                opts = _generate_geo_chart_options(sec["id"], sec, ct)
            else:
                opts = _generate_chart_options(sec["id"], sec, ct)
            if opts:
                chart_options_js += f"  chartOptions['{sec['id']}_{ct}'] = {opts};\n"
        all_chart_options_js += chart_options_js

        # 地图 section：注册 geoData 供 JS 动态构建（指标列切换+筛选）
        geo_controls_html = ""
        if is_geo_section:
            geo_js = _build_geo_data_js(sec)
            if geo_js:
                all_geo_data_js += geo_js
                geo_controls_html = _build_geo_controls_html(sec)

        chart_buttons_html = ""
        for ct in sec["available_charts"]:
            active = "active" if ct == sec["chart_type"] else ""
            labels = {"bar": "柱状图", "line": "折线图", "pie": "饼图", "table": "表格",
                      "map": "地图", "heatmap": "热力图"}
            chart_buttons_html += f'<button class="chart-btn {active}" onclick="switchChart(\'{sec["id"]}\', \'{ct}\')">{labels.get(ct, ct)}</button>'
        
        header_cells = "".join(f"<th>{_escape_html(h)}</th>" for h in sec["headers"])
        body_rows = ""
        for row in sec["rows"]:
            is_total = any("合计" in str(v) or "总计" in str(v) for v in row)
            row_class = "total-row" if is_total else ""
            cells = "".join(f"<td>{_escape_html(v)}</td>" for v in row)
            body_rows += f"<tr class=\"{row_class}\">{cells}</tr>"
        
        summary_html = ""
        if sec.get("summary"):
            summary_items = []
            for key, val in sec["summary"].items():
                if isinstance(val, dict):
                    for stat, num in val.items():
                        stat_labels = {"sum": "合计", "avg": "平均", "max": "最大", "min": "最小", "count": "数量"}
                        summary_items.append(f"{stat_labels.get(stat, stat)}: {num}")
                else:
                    summary_items.append(f"{key}: {val}")
            if summary_items:
                summary_html = f'<div class="summary-bar"><span>数据摘要:</span> {" | ".join(summary_items)}</div>'
        
        section_html_parts.append(f'''
<div id="{sec["id"]}" class="section">
  <h2>{_escape_html(sec["title"])}</h2>
  <p class="subtitle">{_escape_html(sec["subtitle"])}</p>
  {summary_html}
  <div class="chart-container">
    <div class="chart-buttons">{chart_buttons_html}</div>
    {geo_controls_html}
    <div id="chart_{sec["id"]}" class="chart-canvas"></div>
  </div>
  <div class="table-container">
    <div class="table-header">数据详情</div>
    <div class="table-wrap">
      <table><thead><tr>{header_cells}</tr></thead><tbody>{body_rows}</tbody></table>
    </div>
  </div>
</div>''')

    summary_cards_html = ""
    top_summary = list(all_summary.items())[:6]
    if top_summary:
        for key, val in top_summary:
            if isinstance(val, dict):
                val_str = str(val.get("sum", ""))
            else:
                val_str = str(val)
            summary_cards_html += f'<div class="summary-card"><div class="card-label">{_escape_html(key)}</div><div class="card-value">{_escape_html(val_str)}</div></div>'

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
<title>{_escape_html(report_title)}</title>
<style>
  * {{ margin: 0; padding: 0; box-sizing: border-box; }}
  body {{ font-family: -apple-system, "Microsoft YaHei", "PingFang SC", sans-serif; background: #f0f2f5; color: #333; line-height: 1.6; }}
  .header {{ background: linear-gradient(135deg, #2E75B6, #1a4a7a); color: #fff; padding: 20px 16px; text-align: center; position: sticky; top: 0; z-index: 100; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }}
  .header h1 {{ font-size: 18px; margin-bottom: 4px; font-weight: 600; }}
  .header p {{ font-size: 12px; opacity: 0.85; }}
  .nav-bar {{ display: flex; justify-content: center; gap: 12px; padding: 10px 0; background: #fff; border-bottom: 1px solid #eee; }}
  .nav-btn {{ padding: 6px 14px; background: #f5f7fa; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; color: #555; cursor: pointer; transition: all 0.2s; }}
  .nav-btn:hover {{ background: #2E75B6; color: #fff; border-color: #2E75B6; }}
  .main-layout {{ display: flex; max-width: 1200px; margin: 0 auto; }}
  .sidebar {{ width: 260px; min-width: 260px; padding: 16px; background: #fff; position: sticky; top: 84px; height: calc(100vh - 84px); overflow-y: auto; }}
  .sidebar h3 {{ font-size: 14px; color: #2E75B6; margin-bottom: 12px; padding-left: 8px; border-left: 3px solid #2E75B6; }}
  .sidebar ul {{ list-style: none; padding-left: 8px; }}
  .sidebar li {{ margin-bottom: 8px; }}
  .sidebar a {{ text-decoration: none; color: #555; font-size: 13px; transition: color 0.2s; display: block; padding: 4px 8px; border-radius: 4px; }}
  .sidebar a:hover {{ color: #2E75B6; background: #f0f5ff; }}
  .summary-cards {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-top: 16px; }}
  .summary-card {{ background: linear-gradient(135deg, #5B9BD5, #2E75B6); color: #fff; padding: 12px; border-radius: 8px; text-align: center; }}
  .summary-card .card-label {{ font-size: 11px; opacity: 0.8; margin-bottom: 4px; }}
  .summary-card .card-value {{ font-size: 16px; font-weight: bold; }}
  .content-area {{ flex: 1; padding: 20px; }}
  .section {{ background: #fff; border-radius: 10px; box-shadow: 0 2px 12px rgba(0,0,0,0.06); padding: 20px; margin-bottom: 20px; }}
  .section h2 {{ font-size: 18px; color: #182B49; margin-bottom: 4px; font-weight: 600; }}
  .section .subtitle {{ font-size: 13px; color: #999; margin-bottom: 12px; }}
  .summary-bar {{ background: #f5f7fa; padding: 10px 14px; border-radius: 6px; margin-bottom: 16px; font-size: 13px; color: #555; }}
  .summary-bar span {{ font-weight: 600; color: #2E75B6; }}
  .chart-container {{ margin-bottom: 20px; }}
  .chart-buttons {{ display: flex; gap: 8px; margin-bottom: 12px; }}
  .chart-btn {{ padding: 6px 14px; background: #f5f7fa; border: 1px solid #ddd; border-radius: 4px; font-size: 12px; color: #555; cursor: pointer; transition: all 0.2s; }}
  .chart-btn:hover {{ background: #e8f0fe; border-color: #2E75B6; color: #2E75B6; }}
  .chart-btn.active {{ background: #2E75B6; color: #fff; border-color: #2E75B6; }}
  .geo-controls {{ display: none; flex-wrap: wrap; gap: 12px; align-items: flex-start; padding: 10px 12px; margin-bottom: 10px; background: #f5f7fa; border: 1px solid #e0e6ed; border-radius: 6px; font-size: 12px; }}
  .geo-ctrl-item {{ display: flex; align-items: center; gap: 6px; }}
  .geo-ctrl-item label {{ color: #555; font-weight: 600; white-space: nowrap; }}
  .geo-ctrl-item select {{ padding: 4px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 12px; max-width: 160px; }}
  .geo-filter-vals {{ display: flex; flex-wrap: wrap; gap: 6px 12px; max-width: 520px; }}
  .geo-fv-item {{ display: inline-flex; align-items: center; gap: 3px; cursor: pointer; color: #444; }}
  .geo-fv-item input {{ margin: 0; }}
  .chart-canvas {{ height: 350px; }}
  .table-container {{ border-top: 1px solid #eee; padding-top: 16px; }}
  .table-header {{ font-size: 14px; font-weight: 600; color: #2E75B6; margin-bottom: 10px; }}
  .table-wrap {{ overflow-x: auto; -webkit-overflow-scrolling: touch; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  th {{ background: #f8f9fa; color: #495057; font-weight: 600; padding: 10px 12px; text-align: left; border-bottom: 2px solid #dee2e6; position: sticky; top: 0; }}
  td {{ padding: 8px 12px; border-bottom: 1px solid #f0f0f0; }}
  tbody tr:nth-child(even) {{ background: #fafbfc; }}
  tbody tr:hover {{ background: #e9f5ff; }}
  tbody tr.total-row {{ background: #D9E2F3; font-weight: bold; }}
  .slide-card {{ background: #fff; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 12px; overflow: hidden; }}
  .slide-num {{ background: #2E75B6; color: #fff; font-size: 12px; padding: 6px 12px; }}
  .slide-card img {{ width: 100%; display: block; cursor: pointer; }}
  .slide-card img.zoomed {{ position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; object-fit: contain; background: rgba(0,0,0,0.9); z-index: 1000; }}
  .footer {{ text-align: center; padding: 24px; color: #999; font-size: 12px; }}
  @media (max-width: 768px) {{
    .main-layout {{ flex-direction: column; }}
    .sidebar {{ width: 100%; min-width: 100%; position: static; height: auto; }}
    .summary-cards {{ grid-template-columns: repeat(3, 1fr); }}
    .chart-canvas {{ height: 280px; }}
  }}
</style>
</head>
<body>
<div class="header">
  <h1>{_escape_html(report_title)}</h1>
  <p>{_escape_html(report_subtitle)}</p>
</div>
<div class="nav-bar">
  <button class="nav-btn" onclick="scrollToTop()">回到顶部</button>
  <button class="nav-btn" onclick="window.print()">打印报告</button>
</div>
<div class="main-layout">
  <div class="sidebar">
    <h3>报告目录</h3>
    <ul>{''.join(toc_html_parts)}</ul>
    <h3>数据摘要</h3>
    <div class="summary-cards">{summary_cards_html}</div>
  </div>
  <div class="content-area">
    {''.join(section_html_parts)}
    <div class="footer">由 Excel 统一分析工具自动生成 | {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>
  </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
<script>
  // 地图数据加载状态：'pending' | 'ready' | 'failed'
  var _chinaMapStatus = 'pending';
  function _loadChinaMap(cb) {{
    if (_chinaMapStatus === 'ready') {{ cb(); return; }}
    if (_chinaMapStatus === 'pending') {{
      _chinaMapStatus = 'loading';
      var s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/echarts@4.9.0/map/js/china.js';
      s.onload = function() {{
        // echarts 4.x 的 china.js 会注册到 echarts.registerMap
        _chinaMapStatus = 'ready';
        cb();
      }};
      s.onerror = function() {{
        // 回退：尝试用 fetch 加载 GeoJSON
        fetch('https://geo.datav.aliyun.com/areas_v3/bound/100000_full.json')
          .then(function(r) {{ return r.json(); }})
          .then(function(data) {{ echarts.registerMap('china', data); _chinaMapStatus='ready'; cb(); }})
          .catch(function() {{ _chinaMapStatus='failed'; cb(); }});
      }};
      document.head.appendChild(s);
    }} else {{
      // loading 中：轮询
      var t = setInterval(function() {{
        if (_chinaMapStatus === 'ready' || _chinaMapStatus === 'failed') {{
          clearInterval(t); cb();
        }}
      }}, 100);
    }}
  }}
</script>
<script>
  var chartInstances = {{}};
  var chartOptions = {{}};
  var geoData = {{}};  // 地图 section 的完整数据，供动态构建
  {all_chart_options_js}
  {all_geo_data_js}
  
  function initCharts() {{
    Object.keys(chartOptions).forEach(function(key) {{
      // key 格式: section_0_bar / section_1_pie，sectionId 含下划线，从末尾分割
      var lastUnderscore = key.lastIndexOf('_');
      var sectionId = key.substring(0, lastUnderscore);
      var chartType = key.substring(lastUnderscore + 1);
      var container = document.getElementById('chart_' + sectionId);
      if (container && chartType === getDefaultChart(sectionId)) {{
        renderChart(sectionId, chartType);
      }}
    }});
  }}
  
  function getDefaultChart(sectionId) {{
    var btns = document.querySelectorAll('#' + sectionId + ' .chart-btn');
    for (var i = 0; i < btns.length; i++) {{
      if (btns[i].classList.contains('active')) {{
        return btns[i].getAttribute('onclick').match(/'([^']+)'/g)[1].replace(/'/g, '');
      }}
    }}
    return 'bar';
  }}
  
  function renderChart(sectionId, chartType) {{
    var container = document.getElementById('chart_' + sectionId);
    if (!container) return;

    // 显示/隐藏地图控件面板
    var geoCtrl = document.getElementById('geo_ctrl_' + sectionId);
    if (geoCtrl) {{
      geoCtrl.style.display = (chartType === 'map' || chartType === 'heatmap') ? 'flex' : 'none';
    }}

    if (chartType === 'table') {{
      container.style.display = 'none';
      return;
    }}

    container.style.display = 'block';

    if (chartInstances[sectionId]) {{
      chartInstances[sectionId].dispose();
    }}

    var doRender = function() {{
      var chart = echarts.init(container);
      // 地图类型且有 geoData → 动态构建（支持指标列切换+筛选）
      var options;
      if ((chartType === 'map' || chartType === 'heatmap') && geoData[sectionId]) {{
        options = buildGeoOptionFromData(sectionId, chartType);
      }} else {{
        options = chartOptions[sectionId + '_' + chartType];
      }}
      if (options) {{
        chart.setOption(options);
        chartInstances[sectionId] = chart;
      }}
    }};

    // 地图类型需先加载 china 地图数据
    if (chartType === 'map' || chartType === 'heatmap') {{
      _loadChinaMap(doRender);
    }} else {{
      doRender();
    }}
  }}

  // 根据当前控件状态动态构建地图 ECharts option
  function buildGeoOptionFromData(sectionId, chartType) {{
    var d = geoData[sectionId];
    if (!d) return null;

    // 读取控件状态
    var metricSel = document.getElementById('geo_metric_' + sectionId);
    var metricIdx = metricSel ? parseInt(metricSel.value) : (d.metricIndices[0] || -1);

    var filterColSel = document.getElementById('geo_filter_col_' + sectionId);
    var filterColIdx = filterColSel && filterColSel.value ? parseInt(filterColSel.value) : -1;

    // 收集选中的筛选值
    var selectedFilterVals = {{}};
    if (filterColIdx >= 0) {{
      var checkboxes = document.querySelectorAll('input[name="geo_fv_' + sectionId + '"]:checked');
      for (var i = 0; i < checkboxes.length; i++) {{
        selectedFilterVals[checkboxes[i].value] = true;
      }}
    }}

    // 过滤+收集数据点
    var points = [];
    var metricVals = [];
    for (var r = 0; r < d.rows.length; r++) {{
      var row = d.rows[r];
      var lat = parseFloat(row[d.latIdx]);
      var lon = parseFloat(row[d.lonIdx]);
      if (isNaN(lat) || isNaN(lon)) continue;
      if (lat === 0 && lon === 0) continue;
      // 筛选
      if (filterColIdx >= 0) {{
        var fv = String(row[filterColIdx] || '');
        if (Object.keys(selectedFilterVals).length > 0 && !selectedFilterVals[fv]) continue;
      }}
      var name = d.nameIdx !== null && d.nameIdx < row.length ? String(row[d.nameIdx]) : '';
      var values = [lon, lat];
      for (var mi = 0; mi < d.metricIndices.length; mi++) {{
        var v = parseFloat(row[d.metricIndices[mi]]) || 0;
        values.push(v);
      }}
      points.push({{name: name, value: values}});
      var mv = metricIdx >= 0 ? (parseFloat(row[metricIdx]) || 0) : 0;
      metricVals.push(mv);
    }}

    if (points.length === 0) {{
      return {{title: {{text: '无符合条件的数据', left: 'center'}}}};
    }}

    // 经纬度范围
    var lats = points.map(function(p) {{ return p.value[1]; }});
    var lons = points.map(function(p) {{ return p.value[0]; }});
    var centerLon = (Math.min.apply(null, lons) + Math.max.apply(null, lons)) / 2;
    var centerLat = (Math.min.apply(null, lats) + Math.max.apply(null, lats)) / 2;
    var span = Math.max(Math.max.apply(null, lons) - Math.min.apply(null, lons),
                        Math.max.apply(null, lats) - Math.min.apply(null, lats), 0.1);
    var zoom = Math.max(1, Math.min(15, 8 - Math.log10(span + 1)));

    // 指标名
    var metricName = metricIdx >= 0 ? d.headers[metricIdx] : '指标';
    // 指标在 value 数组中的位置
    var metricPos = d.metricIndices.indexOf(metricIdx);
    if (metricPos < 0) metricPos = 0;
    var metricValuePos = metricPos + 2;  // +2 因为前两位是 lon, lat

    // 散点大小归一化
    var mMin = Math.min.apply(null, metricVals);
    var mMax = Math.max.apply(null, metricVals);
    var sizes = [];
    for (var i = 0; i < metricVals.length; i++) {{
      if (mMax === mMin) sizes.push(20);
      else sizes.push(Math.round(10 + (metricVals[i] - mMin) / (mMax - mMin) * 40));
    }}

    // tooltip formatter
    var tipParts = ['function(p){{var s=p.name+"<br/>";var lon=p.value[0],lat=p.value[1];'];
    tipParts.push('s+="经度:"+lon.toFixed(4)+", 纬度:"+lat.toFixed(4)+"<br/>";');
    for (var mi = 0; mi < d.metricIndices.length; mi++) {{
      tipParts.push('s+="' + d.headers[d.metricIndices[mi]] + ':"+p.value[' + (mi+2) + ']+"<br/>";');
    }}
    tipParts.push('return s;}}');
    var tipFn = tipParts.join('');

    var seriesType = chartType === 'map' ? 'scatter' : 'effectScatter';
    var itemStyle = chartType === 'map' ?
      {{color: '#C8102E', opacity: 0.85, borderColor: '#333', borderWidth: 0.5}} :
      {{color: '#ED7D31', shadowBlur: 10, shadowColor: 'rgba(237,125,49,0.5)'}};
    var extra = chartType === 'heatmap' ?
      {{showEffectOn: 'render', rippleEffect: {{brushType: 'stroke'}}}} : {{}};

    var opt = {{
      title: {{text: metricName + ' 分布', left: 'center', textStyle: {{fontSize: 14}}}},
      tooltip: {{trigger: 'item'}},
      geo: {{
        map: 'china', roam: true, zoom: zoom,
        center: [centerLon, centerLat],
        itemStyle: {{areaColor: '#F2F0EB', borderColor: '#999'}},
        emphasis: {{itemStyle: {{areaColor: '#DCE6F0'}}, label: {{show: false}}}}
      }},
      series: [{{
        name: metricName, type: seriesType, coordinateSystem: 'geo',
        data: points, symbolSize: sizes, itemStyle: itemStyle,
        label: {{show: true, formatter: '{{b}}', position: 'right', fontSize: 10, color: '#182B49'}},
      }}].concat([extra])
    }};
    // 合并 extra 到 series[0]
    for (var k in extra) {{ opt.series[0][k] = extra[k]; }}
    // 移除多余的空对象
    opt.series = [opt.series[0]];
    // 注入 tooltip formatter（不能 JSON 序列化函数，直接赋值）
    opt.tooltip.formatter = tipFn;
    // eval tipFn 为真实函数
    try {{ opt.tooltip.formatter = eval('(' + tipFn + ')'); }} catch(e) {{}}
    return opt;
  }}

  // 指标列切换
  function onGeoMetricChange(sectionId) {{
    var activeBtn = document.querySelector('#' + sectionId + ' .chart-btn.active');
    var chartType = activeBtn ? (activeBtn.getAttribute('onclick').match(/'([^']+)'$/) || [])[1] : 'map';
    renderChart(sectionId, chartType);
  }}

  // 筛选列切换 → 填充筛选值复选框
  function onGeoFilterColChange(sectionId) {{
    var sel = document.getElementById('geo_filter_col_' + sectionId);
    var valsWrap = document.getElementById('geo_filter_vals_wrap_' + sectionId);
    var valsBox = document.getElementById('geo_filter_vals_' + sectionId);
    if (!sel || !valsBox) return;
    if (!sel.value) {{
      valsWrap.style.display = 'none';
      onGeoMetricChange(sectionId);
      return;
    }}
    var d = geoData[sectionId];
    var vals = d.filterValues[sel.value] || [];
    var html = '';
    // 全选按钮
    html += '<label class="geo-fv-item"><input type="checkbox" name="geo_fv_' + sectionId + '" value="__all__" checked onchange="onGeoFilterAllChange(\\'' + sectionId + '\\')"> 全选</label>';
    for (var i = 0; i < vals.length; i++) {{
      html += '<label class="geo-fv-item"><input type="checkbox" name="geo_fv_' + sectionId + '" value="' + vals[i].replace(/"/g, '&quot;') + '" checked onchange="onGeoFilterValChange(\\'' + sectionId + '\\')"> ' + vals[i] + '</label>';
    }}
    valsBox.innerHTML = html;
    valsWrap.style.display = 'block';
    onGeoMetricChange(sectionId);
  }}

  // 单个筛选值变化
  function onGeoFilterValChange(sectionId) {{
    onGeoMetricChange(sectionId);
  }}

  // 全选/取消全选
  function onGeoFilterAllChange(sectionId) {{
    var allBox = document.querySelector('input[name="geo_fv_' + sectionId + '"][value="__all__"]');
    var boxes = document.querySelectorAll('input[name="geo_fv_' + sectionId + '"]:not([value="__all__"])');
    for (var i = 0; i < boxes.length; i++) {{
      boxes[i].checked = allBox.checked;
    }}
    onGeoMetricChange(sectionId);
  }}

  function switchChart(sectionId, chartType) {{
    var btns = document.querySelectorAll('#' + sectionId + ' .chart-btn');
    for (var i = 0; i < btns.length; i++) {{
      btns[i].classList.remove('active');
    }}
    event.target.classList.add('active');
    renderChart(sectionId, chartType);
  }}
  
  function toggleZoom(img) {{
    img.classList.toggle('zoomed');
    document.body.style.overflow = img.classList.contains('zoomed') ? 'hidden' : '';
  }}
  
  function scrollToTop() {{
    window.scrollTo({{ top: 0, behavior: 'smooth' }});
  }}
  
  window.addEventListener('resize', function() {{
    Object.keys(chartInstances).forEach(function(key) {{
      chartInstances[key].resize();
    }});
  }});
  
  document.addEventListener('DOMContentLoaded', initCharts);
</script>
</body>
</html>"""

    html_path = os.path.join(output_dir, f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[HTML] 报告已生成: {html_path}")
    return html_path


def start_preview_server(html_path: str) -> Tuple[str, object]:
    import http.server
    import threading
    import socket

    for port in range(8765, 8780):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.bind(("", port))
                break
        except OSError:
            continue

    directory = os.path.dirname(html_path)
    filename = os.path.basename(html_path)

    handler = lambda *args: http.server.SimpleHTTPRequestHandler(*args, directory=directory)
    server = http.server.HTTPServer(("0.0.0.0", port), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    url = f"http://localhost:{port}/{filename}"
    print(f"[HTML] 预览服务已启动: {url}")
    return url, server
