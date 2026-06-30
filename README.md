# Excel 统一分析工具

> Excel 数据分析 + PPT 报告生成 + 透视分析 三合一

## 快速开始

```bash
cd app

# 统一配置（一个 Excel 含多个 Sheet）
python main.py 项目配置.xlsx          # 自动检测 → 依次执行 PPT + 透视

# 子命令方式
python main.py ppt 项目配置.xlsx      # 只生成 PPT
python main.py pivot 项目配置.xlsx    # 只跑透视分析
```

### 配置格式（统一 Excel）

一个配置 Excel 包含两个 Sheet：

| Sheet 名 | 用途 | 自动匹配关键词 |
|---------|------|--------------|
| **PPT配置** | 页面+图表定义 | 包含「页码」「页面类型」等列头 |
| **透视分析** | 交叉透视+分组聚合任务 | 包含「数据源」「行维度」等列头 |

> 兼容旧格式：只有第一个 Sheet 有 PPT 列头时，自动兼容。

## 安装

```bash
pip install -r requirements.txt
```

### 依赖

| 包 | 用途 |
|---|------|
| openpyxl | Excel 读写 |
| pandas | 数据处理 |
| python-pptx | PPT 原生图表生成 |
| matplotlib | 地图/图表渲染 |
| cartopy | 地理地图底图 |

## 功能概览

### 1. PPT 报告生成 (`ppt`)

从配置 Excel + 数据 Excel 一键生成 PPT，图表为 **原生可编辑 PPT 图表**（非静态图片）。

**支持的图表类型：**

| 类型 | 配置值 | 说明 |
|------|--------|------|
| 柱状图 | `column` | 默认类型 |
| 折线图 | `line` | 带圆形标记点 |
| 饼图 | `pie` | 百分比标签 |
| 散点图 | `scatter` | XY 坐标 |
| 面积图 | `area` | 填充区域 |
| 圆环图 | `doughnut` | 环形占比 |
| 组合图 | `column,line` | 柱状+折线双Y轴 |
| 散点地图 | `map` | 经纬度散点+指标着色 |
| 热力地图 | `heatmap` | 密度热力分布 |

**配置列（12列）：**

```
页码 | 页面类型 | 页面标题 | 布局 | 图表类型 | 数据Sheet | X轴 | Y轴 | 图表标题 | 颜色 | 区块名 | 结论模板
```

### 2. 透视分析 (`pivot`)

配置驱动的交叉透视与分组聚合分析，支持多表 JOIN。

**聚合方式：** sum、avg、count、max、min

**配置列（7列）：**

```
序号 | 数据源 | Sheet | 行维度 | 列维度 | 值字段 | 聚合方式
```

**JOIN 语法示例：** `订单表 JOIN 客户表 ON 客户ID=客户ID`

## 案例

`cases/02_网络指标/` 包含：
- `项目配置.xlsx` — 统一配置（PPT配置 + 透视分析 两个 Sheet）
- `网络指标数据.xlsx` — 数据文件（小区指标 + 地理数据）

```bash
python main.py cases/02_网络指标/项目配置.xlsx
# 自动检测为综合配置，依次输出 .pptx 和 .xlsx
```

## 项目结构

```
excel2ppt/
├── app/
│   ├── main.py              # 统一入口
│   ├── requirements.txt
│   └── src/
│       ├── excel_reader.py   # 配置+数据+地理数据读取
│       ├── ppt_builder.py    # PPT 原生图表构建引擎
│       ├── map_builder.py    # 地图可视化（cartopy + matplotlib）
│       ├── pivot_analyzer.py # 透视分析核心引擎
│       └── excel_writer.py   # 透视结果写入 Excel
└── cases/
    └── 02_网络指标/          # 示例案例
```
