---
name: "excel-pivot"
description: "Configuration-driven Excel cross-pivot and group aggregation analysis engine. Supports multi-table JOINs with multi-key AND conditions, cascaded pivot (output of one task as input to another), and intermediate join table output for verification. Invoke when user wants to perform pivot analysis on Excel data, needs cross-tabulation, or wants to aggregate data by dimensions from Excel files."
---

# Excel Pivot Analysis Engine

This skill performs **configuration-driven** pivot analysis on Excel data. Simply edit the configuration file to change analysis dimensions, metrics, and aggregation methods — no code changes needed.

## When to Use

Invoke this skill when:
- User wants to create cross-pivot tables from Excel data
- User needs group aggregation (sum/avg/count/max/min/nunique) by dimensions
- User wants to JOIN multiple sheets/tables before analysis
- User mentions "透视分析", "交叉表", "分组汇总", "excel-pivot", "数据透视"
- User wants to generate analysis result Excel files from raw data

## How It Works

```
配置.xlsx (define analysis tasks)
    ↓
pivot_analyzer.py (parse JOIN + execute pivot)
    ↓
分析结果.xlsx (formatted output)
```

### Usage

```bash
python app/main.py pivot cases/ 或
python app/main.py pivot -c 配置.xlsx --data-dir 数据目录 -o 输出.xlsx
```

## Configuration File Format

| 列名 | 说明 | 示例 |
|------|------|------|
| 序号 | Task number (1, 2, 3...) | `1` |
| 数据源 | Data file path AND join spec or block reference | `数据.xlsx` or `{区块名}` or `表A JOIN 表B ON key=key` |
| Sheet | Sheet name within the data file | `订单明细` |
| 行维度 | Row dimension column(s), comma-separated | `月份` or `区域,省份` |
| 列维度 | Column dimension (cross-pivot), leave empty for groupby | `产品` |
| 值字段 | Value column(s), comma-separated | `销售额` or `销售额,利润` |
| 聚合方式 | Aggregation function(s), comma-separated | `sum` or `sum,avg` |
| 结果Sheet | Output sheet name | `月度产品交叉表` |
| 区块名 | Block name for cascaded pivot references | `站点级` |
| 备注 | Optional notes | |

## Multi-Table JOIN

Single key: `表A JOIN 表B ON key=key`

Multi-key AND: `表A JOIN 表B ON key1=key1 AND key2=key2`

JOIN types supported: JOIN, LEFT JOIN, RIGHT JOIN, OUTER JOIN

Chain: `A JOIN B ON k=k JOIN C ON k=k`

Use `@` for sheet specification: `文件.xlsx@Sheet名`

## Cascaded Pivot

Output of one task can be input to another via block references:

| 数据源 | 含义 |
|--------|------|
| `{区块名}` | Reference a specific block |
| `{pivot}` / `{上一个}` | Reference the last successful task output |

## Aggregation Functions

sum, avg/mean, count, max, min, nunique, pct/占比, join/拼接

Multi-word join: `拼接|、` (custom separator after |)

## Intermediate JOIN Table Output

When a JOIN is used in the data source, the joined intermediate table is automatically output as a separate sheet (prefixed `_JOIN中间表_`) in the result Excel, so you can verify the join correctness.

## Notes

- Validation errors are printed as warnings but do NOT block execution
- This tool is independent of the excel2ppt skill
