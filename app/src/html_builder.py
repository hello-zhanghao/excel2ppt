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


def _infer_chart_type(headers: List[str], rows: List[List[str]]) -> str:
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
            
            sections.append({
                "id": f"section_{len(sections)}",
                "title": sec["name"],
                "subtitle": "",
                "chart_type": chart_type,
                "available_charts": ["bar", "line", "pie", "table"],
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
        for ct in sec["available_charts"]:
            opts = _generate_chart_options(sec["id"], sec, ct)
            if opts:
                chart_options_js += f"  chartOptions['{sec['id']}_{ct}'] = {opts};\n"
        
        chart_buttons_html = ""
        for ct in sec["available_charts"]:
            active = "active" if ct == sec["chart_type"] else ""
            labels = {"bar": "柱状图", "line": "折线图", "pie": "饼图", "table": "表格"}
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
  var chartInstances = {{}};
  var chartOptions = {{}};
  {chart_options_js}
  
  function initCharts() {{
    Object.keys(chartOptions).forEach(function(key) {{
      var parts = key.split('_');
      var sectionId = parts[0];
      var chartType = parts[1];
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
    
    if (chartType === 'table') {{
      container.style.display = 'none';
      return;
    }}
    
    container.style.display = 'block';
    
    if (chartInstances[sectionId]) {{
      chartInstances[sectionId].dispose();
    }}
    
    var chart = echarts.init(container);
    var options = chartOptions[sectionId + '_' + chartType];
    if (options) {{
      chart.setOption(options);
      chartInstances[sectionId] = chart;
    }}
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
