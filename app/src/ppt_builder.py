import os
from datetime import datetime
from time import strftime
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn, nsmap
from lxml import etree

HW_RED = "#C8102E"
HW_DARK = "#182B49"
HW_GRAY = "#8C8C8C"
HW_LIGHT = "#F7F8FA"
HW_BORDER = "#E5E7EB"
FONT_NAME = "Microsoft YaHei"

# 商务配色：主色+辅色，避免紫黄混搭
HUAWEI_PALETTE = [
    "#C8102E",  # 主红
    "#182B49",  # 深蓝
    "#5B9BD5",  # 浅蓝
    "#ED7D31",  # 橙
    "#70AD47",  # 绿
    "#A5A5A5",  # 灰
    "#FFC000",  # 金黄
    "#4472C4",  # 商务蓝
]

# 页脚项目名（可被配置覆盖）
DEFAULT_FOOTER = "数据分析报告"

LAYOUT_POSITIONS = {
    "1图": [
        (Inches(1.2), Inches(1.3), Inches(10.9), Inches(5.6)),
    ],
    "2图上下": [
        (Inches(1.2), Inches(1.0), Inches(10.9), Inches(2.8)),
        (Inches(1.2), Inches(4.2), Inches(10.9), Inches(2.8)),
    ],
    "2图左右": [
        (Inches(0.6), Inches(1.3), Inches(5.8), Inches(5.5)),
        (Inches(6.9), Inches(1.3), Inches(5.8), Inches(5.5)),
    ],
    "4图": [
        (Inches(0.6), Inches(1.15), Inches(5.8), Inches(2.75)),
        (Inches(6.9), Inches(1.15), Inches(5.8), Inches(2.75)),
        (Inches(0.6), Inches(4.2), Inches(5.8), Inches(2.75)),
        (Inches(6.9), Inches(4.2), Inches(5.8), Inches(2.75)),
    ],
    "左图右文": [
        (Inches(0.6), Inches(1.3), Inches(6.3), Inches(5.5)),
    ],
    "上文下图": [
        (Inches(1.2), Inches(2.9), Inches(10.9), Inches(4.1)),
    ],
}


def build_ppt(config, chart_map, output_path):
    """根据页面定义逐页生成PPT（使用PPT原生图表）"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    pages = config.get("pages", [])
    colors = config.get("colors", {})

    if not pages:
        pages = _generate_auto_pages(config, chart_map)

    total_pages = len(pages)
    for idx, page_def in enumerate(pages):
        page_type = str(page_def.get("页面类型", "content")).strip().lower()

        if page_type == "cover" or page_type == "封面":
            title = str(page_def.get("页面标题", ""))
            subtitle = ""
            if "|" in title:
                parts = title.split("|", 1)
                title = parts[0].strip()
                subtitle = parts[1].strip()
            _add_cover_slide(prs, title, subtitle, colors)
        else:
            _add_content_slide_from_def(prs, page_def, chart_map, colors, idx + 1, total_pages)

    prs.save(output_path)
    return output_path


def _generate_auto_pages(config, chart_map):
    """没有「PPT页面」Sheet时，自动按每页图表数分页（兼容旧模式）"""
    general = config.get("general", {})
    charts_per_slide = int(general.get("charts_per_slide", 4))
    chart_list = list(chart_map)
    pages = []
    for i in range(0, len(chart_list), charts_per_slide):
        page_charts = chart_list[i:i + charts_per_slide]
        page = {"页面类型": "content", "布局": f"{min(charts_per_slide, len(page_charts))}图"}
        for j, ch in enumerate(page_charts):
            page[f"图表{j + 1}"] = ch.get("图表标题", "")
        pages.append(page)
    return pages


def _set_slide_bg(slide, hex_color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = _hex_to_rgb(hex_color)


def _set_font_name(paragraph, font_name):
    for run in paragraph.runs:
        run.font.name = font_name
    try:
        pPr = paragraph._pPr
        if pPr is None:
            pPr = paragraph._p.get_or_add_pPr()
        defRPr = pPr.find(qn('a:defRPr'))
        if defRPr is None:
            defRPr = pPr.makeelement(qn('a:defRPr'), {})
            pPr.insert(0, defRPr)
        defRPr.set('latin', font_name)
        defRPr.set('ea', font_name)
    except Exception:
        pass


def _add_cover_slide(prs, title, subtitle, colors):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    _set_slide_bg(slide, "FFFFFF")

    # 左侧深色色块（占宽 1/3，作背景）
    left_block = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(4.5), prs.slide_height
    )
    left_block.fill.solid()
    left_block.fill.fore_color.rgb = _hex_to_rgb(HW_DARK)
    left_block.line.fill.background()

    # 深色块右侧红色竖条（强调色）
    red_accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(4.5), Inches(0), Inches(0.06), prs.slide_height
    )
    red_accent.fill.solid()
    red_accent.fill.fore_color.rgb = _hex_to_rgb(HW_RED)
    red_accent.line.fill.background()

    # 左侧块顶部小红色装饰
    top_deco = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.6), Inches(0.8), Inches(0.06)
    )
    top_deco.fill.solid()
    top_deco.fill.fore_color.rgb = _hex_to_rgb(HW_RED)
    top_deco.line.fill.background()

    # 左侧块上的项目标识
    tag_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.8), Inches(3.5), Inches(0.5)
    )
    ttf = tag_box.text_frame
    ttf.word_wrap = True
    tp = ttf.paragraphs[0]
    tp.text = "DATA REPORT"
    tp.font.size = Pt(11)
    tp.font.bold = True
    tp.font.color.rgb = _hex_to_rgb("#FFFFFF")
    tp.alignment = PP_ALIGN.LEFT
    _set_font_name(tp, FONT_NAME)

    # 主标题（右侧白色区域，左对齐）
    txBox = slide.shapes.add_textbox(
        Inches(5.0), Inches(2.7), Inches(7.8), Inches(1.5)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = _hex_to_rgb(HW_DARK)
    p.alignment = PP_ALIGN.LEFT
    _set_font_name(p, FONT_NAME)

    # 副标题
    if subtitle:
        txBox2 = slide.shapes.add_textbox(
            Inches(5.0), Inches(4.0), Inches(7.8), Inches(0.8)
        )
        tf2 = txBox2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle
        p2.font.size = Pt(16)
        p2.font.color.rgb = _hex_to_rgb(HW_GRAY)
        p2.alignment = PP_ALIGN.LEFT
        _set_font_name(p2, FONT_NAME)

    # 右下日期
    date_str = strftime("%Y.%m.%d")
    date_box = slide.shapes.add_textbox(
        Inches(5.0), Inches(6.5), Inches(5), Inches(0.4)
    )
    dtf = date_box.text_frame
    dp = dtf.paragraphs[0]
    dp.text = date_str
    dp.font.size = Pt(10)
    dp.font.color.rgb = _hex_to_rgb(HW_GRAY)
    dp.alignment = PP_ALIGN.LEFT
    _set_font_name(dp, FONT_NAME)

    # 左侧块底部小标识
    foot_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(6.7), Inches(3.5), Inches(0.4)
    )
    ftf = foot_box.text_frame
    fp = ftf.paragraphs[0]
    fp.text = DEFAULT_FOOTER
    fp.font.size = Pt(9)
    fp.font.color.rgb = _hex_to_rgb("#A0A8B0")
    fp.alignment = PP_ALIGN.LEFT
    _set_font_name(fp, FONT_NAME)


def _add_content_slide_from_def(prs, page_def, chart_map, colors, page_num=0, total_pages=0):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    _set_slide_bg(slide, "FFFFFF")

    layout = str(page_def.get("布局", "4图")).strip()
    title = str(page_def.get("页面标题", ""))
    # 支持"主标题|副标题"语法
    subtitle = ""
    if "|" in title:
        parts = title.split("|", 1)
        title = parts[0].strip()
        subtitle = parts[1].strip()

    if title:
        _add_page_title(slide, title, subtitle, colors)

    positions = LAYOUT_POSITIONS.get(layout, LAYOUT_POSITIONS["4图"])

    active_charts = []
    if "charts" in page_def and page_def["charts"]:
        active_charts = page_def["charts"]
    else:
        for i in range(1, 5):
            chart_def = None
            inline_title = page_def.get(f"图表{i}标题", "")
            if inline_title and str(inline_title).strip():
                chart_def = {
                    "图表标题": str(inline_title).strip(),
                    "图表类型": str(page_def.get(f"图表{i}类型", "column")).strip(),
                    "数据Sheet": str(page_def.get(f"图表{i}Sheet", "Sheet1")).strip(),
                    "X轴范围": str(page_def.get(f"图表{i}X范围", "")).strip(),
                    "Y轴范围": str(page_def.get(f"图表{i}Y范围", "")).strip(),
                    "颜色": str(page_def.get(f"图表{i}颜色", "")).strip(),
                    "_categories": page_def.get(f"_图表{i}_categories"),
                    "_values": page_def.get(f"_图表{i}_values"),
                }
            else:
                chart_name = page_def.get(f"图表{i}", "")
                if chart_name and str(chart_name).strip():
                    chart_def = next(
                        (c for c in chart_map if c.get("图表标题") == str(chart_name).strip()),
                        None,
                    )
            if chart_def:
                active_charts.append(chart_def)

    # A4：给每个图表区加浅灰圆角矩形背景
    chart_rects = []
    for idx, chart_def in enumerate(active_charts):
        if idx >= len(positions):
            break
        left, top, width, height = positions[idx]
        _add_chart_bg(slide, left, top, width, height)
        chart_rects.append((left, top, width, height))

    for idx, chart_def in enumerate(active_charts):
        if idx >= len(chart_rects):
            break
        left, top, width, height = chart_rects[idx]
        _add_native_chart(slide, chart_def, colors, left, top, width, height)

    _add_footer(slide, page_num, total_pages)


def _add_chart_bg(slide, left, top, width, height):
    """图表区背景：浅灰圆角矩形，让图表与白底页面视觉分离"""
    pad = Inches(0.1)
    bg = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left - pad, top - pad,
        width + 2 * pad, height + 2 * pad
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = _hex_to_rgb(HW_LIGHT)
    bg.line.color.rgb = _hex_to_rgb(HW_BORDER)
    bg.line.width = Pt(0.5)
    try:
        bg.adjustments[0] = 0.04
    except Exception:
        pass
    # 移到最底层
    spTree = bg._element.getparent()
    spTree.remove(bg._element)
    spTree.insert(2, bg._element)


def _add_page_title(slide, title, subtitle, colors):
    # 左侧装饰色块
    deco_block = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(0.28), Inches(0.08), Inches(0.5)
    )
    deco_block.fill.solid()
    deco_block.fill.fore_color.rgb = _hex_to_rgb(HW_RED)
    deco_block.line.fill.background()

    # 主标题
    left = Inches(0.85)
    top = Inches(0.2)
    width = Inches(11.85)
    height = Inches(0.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = _hex_to_rgb(HW_DARK)
    p.alignment = PP_ALIGN.LEFT
    _set_font_name(p, FONT_NAME)

    # 副标题
    if subtitle:
        sub_box = slide.shapes.add_textbox(
            Inches(0.85), Inches(0.72), Inches(11.85), Inches(0.3)
        )
        stf = sub_box.text_frame
        stf.word_wrap = True
        sp = stf.paragraphs[0]
        sp.text = subtitle
        sp.font.size = Pt(11)
        sp.font.color.rgb = _hex_to_rgb(HW_GRAY)
        sp.alignment = PP_ALIGN.LEFT
        _set_font_name(sp, FONT_NAME)

    # 红色短线（标题下方）
    line_left = Inches(0.6)
    line_top = Inches(0.78 if not subtitle else 1.05)
    line_width = Inches(1.6)
    line_height = Inches(0.04)
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, line_left, line_top, line_width, line_height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = _hex_to_rgb(HW_RED)
    shape.line.fill.background()


def _add_footer(slide, page_num, total_pages):
    """页脚：左侧项目名 + 右侧页码 + 顶部分隔线"""
    # 分隔细线
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(7.05), Inches(12.1), Inches(0.015)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = _hex_to_rgb(HW_BORDER)
    line.line.fill.background()

    # 左下项目名
    left_box = slide.shapes.add_textbox(
        Inches(0.6), Inches(7.1), Inches(6), Inches(0.3)
    )
    ltf = left_box.text_frame
    lp = ltf.paragraphs[0]
    lp.text = DEFAULT_FOOTER
    lp.font.size = Pt(8)
    lp.font.color.rgb = _hex_to_rgb(HW_GRAY)
    lp.alignment = PP_ALIGN.LEFT
    _set_font_name(lp, FONT_NAME)

    # 右下页码
    if total_pages:
        page_text = f"{page_num} / {total_pages}"
        right_box = slide.shapes.add_textbox(
            Inches(11.0), Inches(7.1), Inches(1.7), Inches(0.3)
        )
        rtf = right_box.text_frame
        rp = rtf.paragraphs[0]
        rp.text = page_text
        rp.font.size = Pt(8)
        rp.font.color.rgb = _hex_to_rgb(HW_GRAY)
        rp.alignment = PP_ALIGN.RIGHT
        _set_font_name(rp, FONT_NAME)


def _add_right_text(slide, text, colors):
    left = Inches(7.2)
    top = Inches(1.2)
    width = Inches(5.6)
    height = Inches(5.5)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(14)
        p.font.color.rgb = _hex_to_rgb("#333333")
        p.space_after = Pt(8)


def _add_top_text(slide, text, colors):
    left = Inches(1.0)
    top = Inches(0.9)
    width = Inches(11.3)
    height = Inches(1.8)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(13)
        p.font.color.rgb = _hex_to_rgb("#333333")
        p.space_after = Pt(4)


def _add_native_chart(slide, chart_info, colors, left, top, width, height):
    chart_type = str(chart_info.get("图表类型", "column")).strip().lower()
    title = str(chart_info.get("图表标题", ""))
    categories = chart_info.get("_categories", [])
    values = chart_info.get("_values", [])
    color = str(chart_info.get("颜色", "")).strip()
    is_hierarchical = bool(chart_info.get("_is_hierarchical"))

    if chart_info.get("_is_map"):
        _add_map_slide(slide, chart_info, left, top, width, height)
        return

    if not color:
        color = HW_RED
    if not categories:
        return

    if _is_combo_chart_type(chart_type):
        _add_combo_chart(slide, chart_info, chart_type, title, categories, values, color, colors, left, top, width, height)
        return

    is_scatter = chart_type == "scatter"
    xl_type = _get_chart_type(chart_type)

    if is_hierarchical and isinstance(values, dict):
        _add_hierarchical_chart(slide, chart_type, title, categories, values, color, colors, xl_type, chart_info, left, top, width, height)
        return

    series_list = []
    if isinstance(values, dict):
        for s_name, s_vals in values.items():
            series_list.append((s_name, s_vals))
    elif isinstance(values, list):
        series_list.append((title or "数据", values))

    if not series_list:
        return

    if is_scatter:
        chart_data = XyChartData()
        for s_name, s_vals in series_list:
            series = chart_data.add_series(s_name)
            for i in range(min(len(categories), len(s_vals))):
                try:
                    x_val = float(i + 1)
                    y_val = float(s_vals[i])
                    series.add_data_point(x_val, y_val)
                except (ValueError, TypeError):
                    continue
    else:
        chart_data = CategoryChartData()
        chart_data.categories = [str(c) for c in categories]
        for s_name, s_vals in series_list:
            chart_data.add_series(s_name, s_vals)

    chart_shape = slide.shapes.add_chart(
        xl_type, left, top, width, height, chart_data
    )
    chart = chart_shape.chart

    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
        for p in chart.chart_title.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.color.rgb = _hex_to_rgb(HW_DARK)

    _apply_series_color(chart, color, chart_type, colors)
    _style_chart(chart, chart_type)
    _render_conclusion(slide, chart_info, left, top - Inches(0.35), width)


def _is_combo_chart_type(chart_type):
    return "," in chart_type or chart_type == "combo"


def _parse_combo_types(chart_type):
    parts = [t.strip() for t in chart_type.split(",") if t.strip()]
    return [t for t in parts if t in ("column", "line", "bar")] or ["column", "line"]


def _add_combo_chart(slide, chart_info, chart_type, title, categories, values, color, colors, left, top, width, height):
    series_list = []
    if isinstance(values, dict):
        for s_name, s_vals in values.items():
            series_list.append((s_name, s_vals))
    elif isinstance(values, list):
        series_list.append((title or "数据", values))

    if not series_list:
        return

    type_map = _parse_combo_types(chart_type)
    while len(type_map) < len(series_list):
        type_map.append("column")

    chart_data = CategoryChartData()
    # 检测层级X轴：categories为 [(parent, [(child, {...}), ...]), ...]
    if categories and isinstance(categories[0], (tuple, list)) and len(categories[0]) == 2 \
            and isinstance(categories[0][1], list):
        cats = chart_data.categories
        for parent, children in categories:
            cat = cats.add_category(str(parent))
            for child_label, _child_data in children:
                cat.add_sub_category(str(child_label))
    else:
        chart_data.categories = [str(c) for c in categories]
    for s_name, s_vals in series_list:
        chart_data.add_series(s_name, s_vals)

    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, chart_data
    )
    chart = chart_shape.chart

    _restructure_as_combo_chart(chart, type_map)

    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
        for p in chart.chart_title.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.color.rgb = _hex_to_rgb(HW_DARK)

    _apply_combo_series_color(chart, chart_type, colors, type_map)
    _style_chart(chart, chart_type)

    _render_conclusion(slide, chart_info, left, top - Inches(0.35), width)


def _restructure_as_combo_chart(chart, type_map):
    chart_cs = chart._chartSpace
    chart_el = chart_cs.find(qn('c:chart'))
    if chart_el is None:
        return
    plot_area = chart_el.find(qn('c:plotArea'))
    if plot_area is None:
        return
    bar_chart = plot_area.find(qn('c:barChart'))
    if bar_chart is None:
        return

    sers = bar_chart.findall(qn('c:ser'))
    if not sers:
        return

    line_sers = []
    bar_sers = []
    for i, ser in enumerate(sers):
        if i < len(type_map) and type_map[i] == "line":
            line_sers.append(ser)
        else:
            bar_sers.append(ser)

    if not line_sers:
        return

    for ser in line_sers:
        bar_chart.remove(ser)

    if not bar_sers:
        axid_elements = bar_chart.findall(qn('c:axId'))
        for axid in axid_elements:
            bar_chart.remove(axid)

    # 清除所有旧的 axId，重新分配
    for axid in bar_chart.findall(qn('c:axId')):
        bar_chart.remove(axid)

    cat_ax = plot_area.find(qn('c:catAx'))
    left_val_ax = plot_area.find(qn('c:valAx'))
    cat_axid = cat_ax.find(qn('c:axId')).get('val') if cat_ax is not None else None
    left_val_axid = left_val_ax.find(qn('c:axId')).get('val') if left_val_ax is not None else None

    bar_axids = bar_chart.findall(qn('c:axId'))
    bar_refs_right = any(a.get('val') == (str(int(left_val_axid) + 1) if left_val_axid else None) for a in bar_axids)

    line_chart = etree.SubElement(plot_area, qn('c:lineChart'))
    grouping = etree.SubElement(line_chart, qn('c:grouping'))
    grouping.set('val', 'standard')
    for ser in line_sers:
        line_chart.append(ser)

    if cat_axid and left_val_axid and left_val_ax is not None and not bar_refs_right:
        new_val_axid = str(int(left_val_axid) + 1)
        axid_cat = etree.SubElement(line_chart, qn('c:axId'))
        axid_cat.set('val', cat_axid)
        axid_val = etree.SubElement(line_chart, qn('c:axId'))
        axid_val.set('val', new_val_axid)

        right_val_ax = etree.SubElement(plot_area, qn('c:valAx'))
        etree.SubElement(right_val_ax, qn('c:axId')).set('val', new_val_axid)
        scaling = etree.SubElement(right_val_ax, qn('c:scaling'))
        etree.SubElement(scaling, qn('c:orientation')).set('val', 'minMax')
        etree.SubElement(right_val_ax, qn('c:delete')).set('val', '0')
        etree.SubElement(right_val_ax, qn('c:axPos')).set('val', 'r')
        etree.SubElement(right_val_ax, qn('c:majorTickMark')).set('val', 'out')
        etree.SubElement(right_val_ax, qn('c:minorTickMark')).set('val', 'none')
        etree.SubElement(right_val_ax, qn('c:tickLblPos')).set('val', 'nextTo')
        etree.SubElement(right_val_ax, qn('c:crossAx')).set('val', cat_axid)
        etree.SubElement(right_val_ax, qn('c:crosses')).set('val', 'max')
        etree.SubElement(right_val_ax, qn('c:auto')).set('val', '1')

        baxid_cat = etree.SubElement(bar_chart, qn('c:axId'))
        baxid_cat.set('val', cat_axid)
        baxid_val = etree.SubElement(bar_chart, qn('c:axId'))
        baxid_val.set('val', left_val_axid)
    elif bar_sers:
        for axid in bar_axids:
            new_axid = etree.SubElement(line_chart, qn('c:axId'))
            new_axid.set('val', axid.get('val'))


def _apply_combo_series_color(chart, chart_type, colors, type_map):
    palette = _get_pie_color_list(colors)
    try:
        chart_cs = chart._chartSpace
        chart_el = chart_cs.find(qn('c:chart'))
        if chart_el is None:
            return
        plot_area = chart_el.find(qn('c:plotArea'))
        if plot_area is None:
            return

        bar_chart = plot_area.find(qn('c:barChart'))
        line_chart = plot_area.find(qn('c:lineChart'))

        sers = []
        if bar_chart is not None:
            sers.extend(bar_chart.findall(qn('c:ser')))
        if line_chart is not None:
            sers.extend(line_chart.findall(qn('c:ser')))

        for idx, ser in enumerate(sers):
            c = _hex_to_rgb(palette[idx % len(palette)])
            is_line = idx < len(type_map) and type_map[idx] == "line"

            if is_line:
                spPr = ser.find(qn('c:spPr'))
                if spPr is None:
                    spPr = etree.SubElement(ser, qn('c:spPr'))
                ln = spPr.find(qn('a:ln'))
                if ln is None:
                    ln = etree.SubElement(spPr, qn('a:ln'))
                ln.set('w', '25400')
                solidFill = ln.find(qn('a:solidFill'))
                if solidFill is None:
                    solidFill = etree.SubElement(ln, qn('a:solidFill'))
                srgb = solidFill.find(qn('a:srgbClr'))
                if srgb is None:
                    for child in list(solidFill):
                        solidFill.remove(child)
                    srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
                srgb.set('val', '{:02X}{:02X}{:02X}'.format(c[0], c[1], c[2]))

                marker = ser.find(qn('c:marker'))
                if marker is None:
                    marker = etree.SubElement(ser, qn('c:marker'))
                symbol = marker.find(qn('c:symbol'))
                if symbol is None:
                    symbol = etree.SubElement(marker, qn('c:symbol'))
                symbol.set('val', 'circle')
            else:
                spPr = ser.find(qn('c:spPr'))
                if spPr is None:
                    spPr = etree.SubElement(ser, qn('c:spPr'))
                solidFill = spPr.find(qn('a:solidFill'))
                if solidFill is None:
                    solidFill = etree.SubElement(spPr, qn('a:solidFill'))
                srgb = solidFill.find(qn('a:srgbClr'))
                if srgb is None:
                    for child in list(solidFill):
                        solidFill.remove(child)
                    srgb = etree.SubElement(solidFill, qn('a:srgbClr'))
                srgb.set('val', '{:02X}{:02X}{:02X}'.format(c[0], c[1], c[2]))
                ln = spPr.find(qn('a:ln'))
                if ln is not None:
                    try:
                        spPr.remove(ln)
                    except Exception:
                        pass
    except Exception as e:
        print(f"[警告] 组合图颜色应用失败: {e}")


def _add_map_slide(slide, chart_info, left, top, width, height):
    from src.map_builder import build_scatter_map, build_heatmap, save_map_image

    geo_df = chart_info.get("_geo_df")
    if geo_df is None or geo_df.empty:
        return

    chart_type = str(chart_info.get("图表类型", "map")).strip().lower()
    title = str(chart_info.get("图表标题", ""))
    x_range = str(chart_info.get("X轴范围", "")).strip()
    y_range = str(chart_info.get("Y轴范围", "")).strip()
    color_hex = str(chart_info.get("颜色", "")).strip() or None

    x_cols = [c.strip() for c in x_range.split(",") if c.strip()] if x_range else []
    y_cols = [c.strip() for c in y_range.split(",") if c.strip()] if y_range else []

    default_lat = [c for c in ["纬度", "lat", "latitude", "Lat"] if c in geo_df.columns]
    default_lon = [c for c in ["经度", "lon", "longitude", "lng", "Lon"] if c in geo_df.columns]

    lon_col = x_cols[0] if x_cols and x_cols[0] in geo_df.columns else (default_lon[0] if default_lon else None)
    lat_col = y_cols[0] if y_cols and y_cols[0] in geo_df.columns else (default_lat[0] if default_lat else None)

    metric_col = y_cols[-1] if len(y_cols) >= 2 and y_cols[-1] in geo_df.columns else None
    if not metric_col:
        num_cols = geo_df.select_dtypes(include=["number"]).columns.tolist()
        exclude = {lat_col, lon_col, "site_id", "cell_id", "站点ID", "小区ID"}
        metric_cols = [c for c in num_cols if c not in exclude and c not in default_lat and c not in default_lon]
        if metric_cols:
            metric_col = metric_cols[0]

    if not lat_col or not lon_col or not metric_col:
        return

    if chart_type == "heatmap":
        fig = build_heatmap(geo_df, lat_col, lon_col, metric_col, title)
    else:
        fig = build_scatter_map(geo_df, lat_col, lon_col, metric_col, title, color_hex)

    if fig is None:
        return

    # 直接用 BytesIO 流插入图片，避免临时文件残留
    png_stream = save_map_image(fig)
    slide.shapes.add_picture(png_stream, left, top, width, height)


def _get_chart_type(chart_type):
    type_map = {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE,
        "pie": XL_CHART_TYPE.PIE,
        "scatter": XL_CHART_TYPE.XY_SCATTER,
        "area": XL_CHART_TYPE.AREA,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
    }
    return type_map.get(chart_type, XL_CHART_TYPE.COLUMN_CLUSTERED)


def _add_hierarchical_chart(slide, chart_type, title, categories, values, color, colors, xl_type, chart_info, left, top, width, height):
    chart_data = CategoryChartData()
    cats = chart_data.categories

    for parent, children in categories:
        cat = cats.add_category(str(parent))
        for child_label, _child_data in children:
            cat.add_sub_category(str(child_label))

    if isinstance(values, dict):
        for series_name, series_vals in values.items():
            clean = [v if v is not None else 0 for v in series_vals]
            chart_data.add_series(str(series_name), clean)
    else:
        chart_data.add_series(title or "数据", values)

    chart_shape = slide.shapes.add_chart(
        xl_type, left, top, width, height, chart_data
    )
    chart = chart_shape.chart

    chart.has_title = bool(title)
    if title:
        chart.chart_title.text_frame.text = title
        for p in chart.chart_title.text_frame.paragraphs:
            for run in p.runs:
                run.font.size = Pt(13)
                run.font.bold = True
                run.font.color.rgb = _hex_to_rgb(HW_DARK)

    _apply_series_color(chart, color, chart_type, colors)
    _style_chart(chart, chart_type)
    _render_conclusion(slide, chart_info, left, top - Inches(0.35), width)


def _apply_series_color(chart, color_hex, chart_type, colors):
    rgb = _hex_to_rgb(color_hex)
    num_series = len(chart.series)

    if chart_type == "pie":
        pie_colors = _get_pie_color_list(colors)
        plot = chart.plots[0]
        series = plot.series[0]
        for idx, point in enumerate(series.points):
            point_color = pie_colors[idx % len(pie_colors)]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = _hex_to_rgb(point_color)
        plot.has_data_labels = True
        data_labels = plot.data_labels
        data_labels.show_percentage = True
        data_labels.show_category_name = True
        data_labels.show_value = False
        data_labels.font.size = Pt(8)
        data_labels.font.color.rgb = _hex_to_rgb(HW_DARK)
    else:
        palette = _get_pie_color_list(colors)
        for idx, series in enumerate(chart.series):
            s_color = _hex_to_rgb(palette[idx % len(palette)]) if num_series > 1 else rgb
            if chart_type == "line":
                series.format.line.color.rgb = s_color
                series.format.line.width = Pt(2)
                series.marker.style = XL_MARKER_STYLE.CIRCLE
                series.marker.size = 6
                series.marker.format.fill.solid()
                series.marker.format.fill.fore_color.rgb = s_color
                series.marker.format.line.color.rgb = s_color
            else:
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = s_color


def _get_pie_color_list(colors):
    accent_keys = sorted([k for k in colors if k.startswith("accent")])
    palette = []
    for k in accent_keys:
        palette.append(colors[k])
    palette += HUAWEI_PALETTE
    return palette


def _style_chart(chart, chart_type):
    try:
        num_series = len(chart.series)
        chart.has_legend = (chart_type == "pie" or num_series > 1)
        if chart.has_legend:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = Pt(9)
            chart.legend.font.color.rgb = _hex_to_rgb(HW_DARK)
    except Exception as e:
        print(f"[警告] 图例样式设置失败: {e}")

    try:
        chart_cs = chart._chartSpace
        chart_el = chart_cs.find(qn('c:chart'))
        if chart_el is not None:
            plot_area = chart_el.find(qn('c:plotArea'))
            if plot_area is not None:
                for ax_tag in ['c:valAx', 'c:catAx']:
                    for axis in plot_area.findall(qn(ax_tag)):
                        major_grid = axis.find(qn('c:majorGridlines'))
                        if major_grid is None:
                            major_grid = etree.SubElement(axis, qn('c:majorGridlines'))
                        spPr = major_grid.find(qn('c:spPr'))
                        if spPr is None:
                            spPr = etree.SubElement(major_grid, qn('c:spPr'))
                        ln = spPr.find(qn('a:ln'))
                        if ln is None:
                            ln = etree.SubElement(spPr, qn('a:ln'))
                        sf = ln.find(qn('a:solidFill'))
                        if sf is None:
                            for child in list(ln):
                                ln.remove(child)
                            sf = etree.SubElement(ln, qn('a:solidFill'))
                        sc = sf.find(qn('a:srgbClr'))
                        if sc is None:
                            for child in list(sf):
                                sf.remove(child)
                            sc = etree.SubElement(sf, qn('a:srgbClr'))
                        sc.set('val', 'D9D9D9')
                        ln.set('w', '9525')
    except Exception as e:
        print(f"[警告] 网格线样式设置失败: {e}")

    try:
        if chart_type not in ("pie", "doughnut", "scatter"):
            value_axis = chart.value_axis
            value_axis.tick_labels.font.size = Pt(9)
            value_axis.tick_labels.font.color.rgb = _hex_to_rgb(HW_GRAY)
            value_axis.has_major_gridlines = True
            if chart.has_legend:
                value_axis.visible = True
            cat_axis = chart.category_axis
            cat_axis.tick_labels.font.size = Pt(9)
            cat_axis.tick_labels.font.color.rgb = _hex_to_rgb(HW_DARK)
            _force_axis_text_horizontal(cat_axis)
    except Exception as e:
        print(f"[警告] 坐标轴样式设置失败: {e}")

    try:
        plot = chart.plots[0]
        # scatter 图不支持数据标签，跳过
        if chart_type != "scatter":
            plot.has_data_labels = True
            data_labels = plot.data_labels
            data_labels.font.size = Pt(9)
            data_labels.font.color.rgb = _hex_to_rgb(HW_DARK)
            try:
                data_labels.show_legend_key = False
                if chart_type in ("pie", "doughnut"):
                    data_labels.number_format = '0.0%'
                else:
                    data_labels.show_value = True
                    # 千分位 + 负数红色（商务报表风格）
                    data_labels.number_format = '#,##0;[Red]-#,##0'
                    # line/area 图表不支持 OUTSIDE_END，Office 会报文件损坏
                    if chart_type in ("line", "area"):
                        data_labels.position = XL_LABEL_POSITION.CENTER
                    else:
                        data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
            except Exception:
                pass
    except Exception as e:
        print(f"[警告] 数据标签样式设置失败: {e}")


def _force_axis_text_horizontal(axis):
    try:
        txPr = axis.tick_labels._txPr
        bodyPr = txPr.find(qn('a:bodyPr'))
        if bodyPr is None:
            bodyPr = txPr.makeelement(qn('a:bodyPr'), {})
            txPr.append(bodyPr)
        bodyPr.set('rot', '0')
        if 'vert' in bodyPr.attrib:
            del bodyPr.attrib['vert']
    except Exception:
        pass


def _hex_to_rgb(hex_color):
    hex_color = str(hex_color).lstrip("#")
    if len(hex_color) == 6:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return RGBColor(r, g, b)
    return RGBColor(0x18, 0x2B, 0x49)


def _render_conclusion(slide, chart_info, left, top, width):
    tmpl = str(chart_info.get("结论模板", "")).strip()
    if not tmpl:
        return
    categories = chart_info.get("_categories", [])
    values = chart_info.get("_values", {})

    flat_cats, flat_vals = _flatten_chart_data(categories, values)
    if not flat_vals:
        return

    ctx = _compute_stats(flat_cats, flat_vals)

    # 多系列时补充每个系列的命名统计；全局 max/min/avg 已基于全量数据计算
    if isinstance(values, dict) and len(values) > 1:
        for s_name in values:
            svals = [v for v in values[s_name] if v is not None]
            if svals:
                ctx[f"max_{s_name}"] = max(svals)
                ctx[f"min_{s_name}"] = min(svals)
                ctx[f"avg_{s_name}"] = round(sum(svals) / len(svals), 1)
                ctx[f"total_{s_name}"] = sum(svals)
            # 标记多系列：占位符 {max_<系列名>} 可精确定位到某一系列

    try:
        text = tmpl
        for key, val in ctx.items():
            text = text.replace("{" + key + "}", str(val))
    except Exception:
        text = tmpl

    if text == tmpl:
        return

    tb = slide.shapes.add_textbox(
        left, top, width, Inches(0.35)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "✦ " + text
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = _hex_to_rgb(HW_DARK)
    p.alignment = PP_ALIGN.CENTER
    p.space_before = Pt(2)
    _set_font_name(p, FONT_NAME)


def _flatten_chart_data(categories, values):
    """
    展平图表数据用于统计。
    多系列时合并所有系列的 (类别, 值)，而非只取第一个系列，
    以确保 _compute_stats 的全局 max/min/avg 反映完整数据。
    """
    # 先展开单系列的类别标签
    if isinstance(categories, list) and categories and isinstance(categories[0], (tuple, list)):
        base_cats = [str(child_label) for parent, children in categories for child_label, _ in children]
    else:
        base_cats = [str(c) for c in categories]

    flat_cats = []
    flat_vals = []

    if isinstance(values, dict):
        # 多系列：每个系列都用相同的 base_cats，合并所有 (cat, val)
        for s_vals in values.values():
            for i, v in enumerate(s_vals):
                if v is not None:
                    flat_vals.append(v)
                    flat_cats.append(base_cats[i] if i < len(base_cats) else f"项{i+1}")
    elif isinstance(values, list):
        for i, v in enumerate(values):
            if v is not None:
                flat_vals.append(v)
                flat_cats.append(base_cats[i] if i < len(base_cats) else f"项{i+1}")

    return flat_cats, flat_vals


def _compute_stats(categories, values):
    if not values:
        return {}
    # 防御非数值数据：过滤掉无法转 float 的值，避免 sort/sum 抛 TypeError
    pairs = []
    for cat, v in zip(categories, values):
        try:
            pairs.append((cat, float(v)))
        except (TypeError, ValueError):
            continue
    if not pairs:
        return {}
    pairs.sort(key=lambda x: x[1], reverse=True)
    total = sum(p[1] for p in pairs)
    avg = round(total / len(pairs), 1)
    ctx = {
        "max_val": pairs[0][1],
        "max_cat": pairs[0][0],
        "min_val": pairs[-1][1],
        "min_cat": pairs[-1][0],
        "avg": avg,
        "total": total,
        "count": len(pairs),
    }
    if len(pairs) >= 3:
        ctx["top3_cats"] = ", ".join(p[0] for p in pairs[:3])
        ctx["top3_vals"] = ", ".join(str(p[1]) for p in pairs[:3])
        ctx["bottom3_cats"] = ", ".join(p[0] for p in pairs[-3:])
        ctx["bottom3_vals"] = ", ".join(str(p[1]) for p in pairs[-3:])
    return ctx


# ==================== PPT 配置校验 ====================

VALID_CHART_TYPES = {"bar", "column", "line", "pie", "scatter", "area", "doughnut", "combo", "map", "heatmap"}
PIVOT_DATA_SOURCE_KEYWORDS = ("{pivot}", "pivot", "透视结果", "透视分析")


def validate_ppt_config(config, config_dir, pivot_data_file=None):
    """
    校验 PPT 配置，返回 (results, all_ok)。
    results: list[dict] 每项含 level(error/warning/info)/page/chart/column/message
    all_ok: True 表示无错误
    """
    results = []
    pages = config.get("pages", [])

    if not pages:
        results.append({
            "level": "warning",
            "page": "-",
            "chart": "-",
            "column": "-",
            "message": "未检测到 PPT 页面配置（pages 为空）",
        })
        return results, True

    # 收集数据文件 Sheet 列名缓存，避免重复 IO
    sheet_columns_cache = {}

    def _get_sheet_columns(file_path, sheet_name):
        key = (file_path, sheet_name)
        if key in sheet_columns_cache:
            return sheet_columns_cache[key]
        try:
            import openpyxl
            ext = os.path.splitext(str(file_path))[1].lower()
            if ext == ".csv":
                import pandas as pd
                df = pd.read_csv(file_path, encoding="utf-8-sig", nrows=0)
                cols = list(df.columns)
            else:
                wb = openpyxl.load_workbook(file_path, read_only=True)
                if sheet_name and sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                else:
                    ws = wb[wb.sheetnames[0]]
                row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
                cols = [str(c).strip() if c is not None else "" for c in row]
                wb.close()
            sheet_columns_cache[key] = cols
            return cols
        except Exception as e:
            sheet_columns_cache[key] = None
            return None

    for page_idx, page in enumerate(pages, 1):
        page_num = page.get("页码", page_idx)
        page_type = str(page.get("页面类型", "内容")).strip().lower()
        layout = str(page.get("布局", "")).strip()
        charts = page.get("charts", [])

        # 封面页不校验图表配置
        if page_type in ("cover", "封面"):
            continue

        # 校验布局名
        if layout and layout not in LAYOUT_POSITIONS:
            results.append({
                "level": "error",
                "page": str(page_num),
                "chart": "-",
                "column": "布局",
                "message": f"布局 '{layout}' 不支持，可用: {', '.join(LAYOUT_POSITIONS.keys())}",
            })

        # 校验图表数量与布局槽位
        if layout and layout in LAYOUT_POSITIONS and charts:
            slot_count = len(LAYOUT_POSITIONS[layout])
            if len(charts) > slot_count:
                results.append({
                    "level": "warning",
                    "page": str(page_num),
                    "chart": "-",
                    "column": "布局",
                    "message": f"图表数 {len(charts)} 超过布局 '{layout}' 的槽位 {slot_count}，多出的将被忽略",
                })

        for chart in charts:
            chart_title = str(chart.get("图表标题", "")).strip()
            chart_type = str(chart.get("图表类型", "column")).strip().lower()
            data_sheet = str(chart.get("数据Sheet", "")).strip()
            data_source = str(chart.get("数据源", "")).strip()
            x_range = str(chart.get("X轴范围", chart.get("X轴", ""))).strip()
            y_range = str(chart.get("Y轴范围", chart.get("Y轴", ""))).strip()
            chart_id = chart_title or f"第{page_idx}页图表"

            # 校验图表标题
            if not chart_title:
                results.append({
                    "level": "warning",
                    "page": str(page_num),
                    "chart": chart_id,
                    "column": "图表标题",
                    "message": "图表标题为空",
                })

            # 校验图表类型
            if chart_type:
                # 组合图支持逗号分隔
                types = [t.strip() for t in chart_type.split(",")]
                for t in types:
                    if t and t not in VALID_CHART_TYPES:
                        results.append({
                            "level": "error",
                            "page": str(page_num),
                            "chart": chart_id,
                            "column": "图表类型",
                            "message": f"图表类型 '{t}' 不支持，可用: {', '.join(sorted(VALID_CHART_TYPES))}",
                        })

            # 跳过地图类型的列名校验（地图列名逻辑较灵活）
            if chart_type in ("map", "heatmap"):
                continue

            # 校验 X/Y 轴必填
            if not x_range:
                results.append({
                    "level": "error",
                    "page": str(page_num),
                    "chart": chart_id,
                    "column": "X轴",
                    "message": "X轴未指定",
                })
            if not y_range:
                results.append({
                    "level": "error",
                    "page": str(page_num),
                    "chart": chart_id,
                    "column": "Y轴",
                    "message": "Y轴未指定",
                })
            if not x_range or not y_range:
                continue

            # 解析数据文件路径
            is_pivot_ref = data_source.lower() in PIVOT_DATA_SOURCE_KEYWORDS
            if is_pivot_ref:
                if not pivot_data_file or not os.path.exists(str(pivot_data_file)):
                    results.append({
                        "level": "warning",
                        "page": str(page_num),
                        "chart": chart_id,
                        "column": "数据源",
                        "message": f"引用透视结果但未提供 --pivot-file，图表将被跳过",
                    })
                file_path = pivot_data_file
            elif data_source:
                from src.excel_reader import find_data_file as _find
                file_path = _find(data_source, config_dir)
                if not file_path:
                    results.append({
                        "level": "warning",
                        "page": str(page_num),
                        "chart": chart_id,
                        "column": "数据源",
                        "message": f"数据源 '{data_source}' 未找到匹配文件",
                    })
            else:
                # 无数据源时无法校验列名，跳过
                continue

            if not file_path or not os.path.exists(str(file_path)):
                continue

            # 校验 Sheet 存在性与列名匹配
            cols = _get_sheet_columns(file_path, data_sheet)
            if cols is None:
                results.append({
                    "level": "warning",
                    "page": str(page_num),
                    "chart": chart_id,
                    "column": "数据Sheet",
                    "message": f"无法读取 Sheet '{data_sheet}' 的列名（文件可能被占用）",
                })
                continue

            # 校验 Sheet 名（Excel 才校验，CSV 无 Sheet）
            ext = os.path.splitext(str(file_path))[1].lower()
            if ext != ".csv" and data_sheet:
                import openpyxl
                try:
                    wb = openpyxl.load_workbook(file_path, read_only=True)
                    if data_sheet not in wb.sheetnames:
                        results.append({
                            "level": "error",
                            "page": str(page_num),
                            "chart": chart_id,
                            "column": "数据Sheet",
                            "message": f"Sheet '{data_sheet}' 不存在，可用: {', '.join(wb.sheetnames)}",
                        })
                    wb.close()
                except Exception:
                    pass

            # 校验 X 轴列名（透视结果文件含区块标题行，列名匹配可能误报，降级为警告）
            is_pivot_result = is_pivot_ref or "分析" in os.path.basename(str(file_path))
            col_check_level = "warning" if is_pivot_result else "error"
            x_cols = [c.strip() for c in x_range.split(",") if c.strip()]
            for xc in x_cols:
                if xc not in cols:
                    results.append({
                        "level": col_check_level,
                        "page": str(page_num),
                        "chart": chart_id,
                        "column": "X轴",
                        "message": f"X轴列名 '{xc}' 在 Sheet '{data_sheet}' 首行中未找到，可用列: {', '.join(c for c in cols[:10] if c)}{'...' if len(cols) > 10 else ''}",
                    })

            # 校验 Y 轴列名
            y_cols = [c.strip() for c in y_range.split(",") if c.strip()]
            for yc in y_cols:
                if yc not in cols:
                    results.append({
                        "level": col_check_level,
                        "page": str(page_num),
                        "chart": chart_id,
                        "column": "Y轴",
                        "message": f"Y轴列名 '{yc}' 在 Sheet '{data_sheet}' 首行中未找到，可用列: {', '.join(c for c in cols[:10] if c)}{'...' if len(cols) > 10 else ''}",
                    })

    all_ok = all(r["level"] != "error" for r in results)
    return results, all_ok


def print_ppt_validation_results(results):
    """格式化打印 PPT 配置校验结果，返回 all_ok"""
    if not results:
        print("[校验] 全部通过 ✓")
        return True

    error_count = sum(1 for r in results if r["level"] == "error")
    warning_count = sum(1 for r in results if r["level"] == "warning")
    info_count = sum(1 for r in results if r["level"] == "info")

    print(f"\n{'='*60}")
    print(f"  PPT 配置校验结果")
    print(f"{'='*60}")
    print(f"  错误: {error_count}  |  警告: {warning_count}  |  提示: {info_count}")
    print(f"{'='*60}")

    for r in results:
        icon = {"error": "✗", "warning": "⚠", "info": "ℹ"}.get(r["level"], " ")
        col_info = f" [{r['column']}]" if r.get("column") and r["column"] != "-" else ""
        chart_info = f" [{r['chart']}]" if r.get("chart") and r["chart"] != "-" else ""
        print(f"  {icon} [页{r['page']}{chart_info}{col_info}] {r['message']}")

    print(f"{'='*60}")

    if error_count > 0:
        print(f"  ❌ 发现 {error_count} 个错误，请修正后再执行")
    elif warning_count > 0:
        print(f"  ⚠  发现 {warning_count} 个警告，可继续执行但可能影响结果")
    print()

    return error_count == 0
