"""
防护用例数据生成脚本
生成「测试数据.xlsx」和「项目配置.xlsx」到同目录
每次代码修改后运行此脚本重建数据，再执行 run_test.py 验证
"""
import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ============================================================
# 1. 生成测试数据
# ============================================================
def create_data_file():
    """创建测试数据文件，包含2个Sheet，覆盖所有核心特性"""
    data_path = os.path.join(SCRIPT_DIR, "测试数据.xlsx")

    # --- Sheet1: 销售明细 ---
    # 设计原则：3个地区×2个产品×2个季度 = 12行，足够覆盖所有聚合但数据量极小
    sales_data = [
        # 地区, 产品, 季度, 销售额, 销量, 客户数
        ("华东", "产品A", "Q1", 1200, 100, 15),
        ("华东", "产品A", "Q2", 1800, 150, 20),
        ("华东", "产品B", "Q1", 800,  60, 10),
        ("华东", "产品B", "Q2", 1000, 80,  12),
        ("华北", "产品A", "Q1", 900,  75,  8),
        ("华北", "产品A", "Q2", 1100, 90,  10),
        ("华北", "产品B", "Q1", 600,  50,  6),
        ("华北", "产品B", "Q2", 700,  55,  7),
        ("华南", "产品A", "Q1", 1500, 120, 18),
        ("华南", "产品A", "Q2", 2000, 160, 25),
        ("华南", "产品B", "Q1", 500,  40,  5),
        ("华南", "产品B", "Q2", 650,  50,  8),
    ]
    df_sales = pd.DataFrame(sales_data, columns=[
        "地区", "产品", "季度", "销售额", "销量", "客户数"
    ])

    # --- Sheet2: 产品目录（用于JOIN演示）---
    catalog_data = [
        ("产品A", "高端系列", 3000),
        ("产品B", "经济系列", 2000),
    ]
    df_catalog = pd.DataFrame(catalog_data, columns=[
        "产品", "产品系列", "基准价"
    ])

    # 写入Excel
    with pd.ExcelWriter(data_path, engine="openpyxl") as writer:
        df_sales.to_excel(writer, sheet_name="销售明细", index=False)
        df_catalog.to_excel(writer, sheet_name="产品目录", index=False)

    print(f"[OK] 数据文件已生成: {data_path}")
    return data_path


# ============================================================
# 2. 生成配置文件
# ============================================================
def create_config_file():
    """创建配置文件，包含透视分析和PPT配置两个Sheet"""
    config_path = os.path.join(SCRIPT_DIR, "项目配置.xlsx")
    wb = openpyxl.Workbook()

    # ---------- 样式定义 ----------
    header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    data_font = Font(name="微软雅黑", size=10)
    data_align = Alignment(vertical="center", wrap_text=True)

    def style_header(ws, ncols):
        for col in range(1, ncols + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

    def style_data(ws, nrows, ncols):
        for row in range(2, nrows + 2):
            for col in range(1, ncols + 1):
                cell = ws.cell(row=row, column=col)
                cell.font = data_font
                cell.alignment = data_align
                cell.border = thin_border

    def auto_width(ws, ncols, min_w=8, max_w=30):
        for col in range(1, ncols + 1):
            max_len = 0
            for row in range(1, ws.max_row + 1):
                val = ws.cell(row=row, column=col).value
                if val and len(str(val)) > max_len:
                    max_len = len(str(val))
            ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 4, min_w), max_w)

    # ============================================================
    # Sheet1: 透视分析
    # ============================================================
    ws_pivot = wb.active
    ws_pivot.title = "透视分析"

    pivot_headers = [
        "序号", "数据源", "Sheet", "行维度", "列维度", "值字段",
        "聚合方式", "结果Sheet", "行映射", "列映射", "值映射",
        "分箱", "值计算", "是否计算", "过滤条件", "备注"
    ]

    # 透视分析任务（8个任务，覆盖所有核心特性）
    pivot_rows = [
        # 任务1: 简单分组求和（基础场景）
        [1, "测试数据.xlsx", "销售明细", "地区", "", "销售额",
         "sum", "按地区汇总", "", "", "总销售额",
         "", "", "是", "", "简单分组求和"],

        # 任务2: 多字段多聚合（单字段多聚合 + 多字段）
        [2, "测试数据.xlsx", "销售明细", "地区", "", "销售额,销售额,客户数",
         "sum,avg,count", "地区多维统计", "", "", "总销售额,平均销售额,客户数",
         "", "", "是", "", "多字段多聚合"],

        # 任务3: 交叉表（行×列维度）
        [3, "测试数据.xlsx", "销售明细", "地区", "产品", "销售额",
         "sum", "地区产品交叉", "", "", "",
         "", "", "是", "", "交叉表"],

        # 任务4: 占比聚合（pct）
        [4, "测试数据.xlsx", "销售明细", "地区", "", "销售额",
         "pct", "地区占比", "", "", "销售额占比",
         "", "", "是", "", "占比聚合"],

        # 任务5: 计数占比（count_pct）
        [5, "测试数据.xlsx", "销售明细", "产品", "", "客户数",
         "count_pct", "产品计数占比", "", "", "客户占比",
         "", "", "是", "", "计数占比"],

        # 任务6: 分箱（自定义列名）
        [6, "测试数据.xlsx", "销售明细", "销售额", "", "客户数",
         "count", "销售额分箱", "", "", "客户数",
         "销售额=500,1000,1500,2000|销售额区间", "", "是", "", "分箱统计"],

        # 任务7: 过滤条件
        [7, "测试数据.xlsx", "销售明细", "地区", "", "销售额",
         "sum", "华东华北汇总", "", "", "总销售额",
         "", "", "是", "地区 = '华东' OR 地区 = '华北'", "过滤条件"],

        # 任务8: JOIN + 值计算
        [8, "测试数据.xlsx@销售明细 JOIN 测试数据.xlsx@产品目录 ON 产品=产品", "销售明细", "产品", "", "销售额",
         "sum", "产品系列汇总", "", "", "总销售额",
         "", "/1000", "是", "", "JOIN+值计算"],

        # 任务9: 跳过任务（是否计算=否）
        [9, "测试数据.xlsx", "销售明细", "季度", "", "销售额",
         "sum", "季度汇总", "", "", "总销售额",
         "", "", "否", "", "跳过任务演示"],
    ]

    # 写入表头
    for col, header in enumerate(pivot_headers, 1):
        ws_pivot.cell(row=1, column=col, value=header)
    # 写入数据
    for row_idx, row_data in enumerate(pivot_rows, 2):
        for col_idx, val in enumerate(row_data, 1):
            ws_pivot.cell(row=row_idx, column=col_idx, value=val)

    style_header(ws_pivot, len(pivot_headers))
    style_data(ws_pivot, len(pivot_rows), len(pivot_headers))
    auto_width(ws_pivot, len(pivot_headers))
    ws_pivot.freeze_panes = "A2"

    # ============================================================
    # Sheet2: PPT配置
    # ============================================================
    ws_ppt = wb.create_sheet("PPT配置")

    ppt_headers = [
        "页码", "页面类型", "页面标题", "副标题", "布局",
        "图表类型", "数据Sheet", "X轴", "Y轴",
        "图表标题", "颜色", "备注", "结论模板", "数据源",
        "HTML生成", "HTML图表类型"
    ]

    # PPT页面（5页，覆盖所有布局和图表类型）
    ppt_rows = [
        # 页1: 封面
        [1, "封面", "销售数据分析报告\n防护用例", "基础防护测试 | 全特性覆盖", "",
         "", "", "", "",
         "", "", "", "", "",
         "是", ""],

        # 页2: 单图柱状图（引用透视结果）
        [2, "内容", "各地区销售汇总", "", "1图",
         "column", "按地区汇总", "地区", "总销售额",
         "各地区总销售额", "auto", "", "销售额最高地区: {max_cat} ({max_val})", "透视结果",
         "是", "bar"],

        # 页3: 2图左右（柱状图+饼图）
        [3, "内容", "多维统计分析", "", "2图左右",
         "bar", "地区多维统计", "地区", "总销售额",
         "各地区销售额对比", "auto", "", "", "透视结果",
         "是", ""],
        # 第二个图表（同页）
        ["", "", "", "", "",
         "pie", "地区占比", "地区", "销售额占比",
         "各地区销售额占比", "auto", "", "", "",
         "是", "pie"],

        # 页4: 折线图（引用透视结果，季度数据）
        [4, "内容", "销售额分箱分布", "", "1图",
         "line", "销售额分箱", "销售额区间", "客户数",
         "销售额区间客户分布", "auto", "", "共 {count} 个区间", "透视结果",
         "是", "line"],

        # 页5: 左图右文（纯文字内容，无图表数据）- HTML生成设置为"否"，用于测试跳过功能
        [5, "内容", "分析结论与说明", "", "左图右文",
         "", "", "", "",
         "", "", "本报告由防护用例自动生成，用于验证PPT输出格式是否正常。", "", "",
         "否", ""],
    ]

    # 写入表头
    for col, header in enumerate(ppt_headers, 1):
        ws_ppt.cell(row=1, column=col, value=header)
    # 写入数据
    for row_idx, row_data in enumerate(ppt_rows, 2):
        for col_idx, val in enumerate(row_data, 1):
            ws_ppt.cell(row=row_idx, column=col_idx, value=val)

    style_header(ws_ppt, len(ppt_headers))
    style_data(ws_ppt, len(ppt_rows), len(ppt_headers))
    auto_width(ws_ppt, len(ppt_headers))
    ws_ppt.freeze_panes = "A2"

    # 保存
    wb.save(config_path)
    print(f"[OK] 配置文件已生成: {config_path}")
    return config_path


# ============================================================
# 主函数
# ============================================================
if __name__ == "__main__":
    print("=" * 60)
    print("  防护用例数据生成")
    print("=" * 60)
    create_data_file()
    create_config_file()
    print("\n生成完成！接下来运行: python run_test.py")
