# 统一 Excel 分析工具项目 — 实施计划

## 概要

将当前分散在两个入口文件（`main.py` + `pivot_main.py`）、以 TRAE Skill 机制管理的 excel2ppt 和 excel-pivot 功能，合并为一个统一项目，移除 TRAE 技能包装层。

## 当前状态

```
f:\【1】AI探索\【3】excel2ppt\
├── .trae/skills/          ← 要删除
│   ├── excel2ppt/SKILL.md
│   └── excel-pivot/SKILL.md
├── app/
│   ├── main.py            ← PPT 生成入口
│   ├── pivot_main.py      ← 透视分析入口
│   ├── requirements.txt
│   └── src/
│       ├── excel_reader.py    ← 共享：配置+数据+地理读取
│       ├── ppt_builder.py     ← PPT 专用
│       ├── map_builder.py     ← PPT 专用（地图）
│       ├── pivot_analyzer.py  ← 透视专用
│       └── excel_writer.py    ← 透视专用
├── cases/
│   ├── 01_销售数据/       ← 多余的案例
│   ├── 02_网络指标/       ← 保留（含PPT+透视）
│   ├── 03_网络透视/       ← 多余的案例
│   └── pivot_demo/        ← 多余的案例
└── README.md              ← 不存在，需新建
```

## 目标状态

```
f:\【1】AI探索\【3】excel2ppt\
├── app/
│   ├── main.py                ← 统一入口（子命令 ppt / pivot）
│   ├── requirements.txt       ← 不变
│   └── src/
│       ├── excel_reader.py    ← 不变
│       ├── ppt_builder.py     ← 不变
│       ├── map_builder.py     ← 不变
│       ├── pivot_analyzer.py  ← 不变
│       └── excel_writer.py    ← 不变
├── cases/
│   └── 02_网络指标/           ← 保留（清理临时文件）
├── README.md                  ← 新建：统一项目文档
└── .trae/skills/              ← 已删除
```

## 改动清单

### 1. 统一入口 — `app/main.py`（重写）

**当前：** 两个独立 CLI 入口，各有自己的 `argparse`、`find_config_file()`、错误处理

**改为：** 合并为一个 `main.py`，用子命令区分功能

```bash
# PPT 生成（原 main.py 的行为）
python app/main.py ppt cases/02_网络指标/配置.xlsx
python app/main.py ppt -c 配置.xlsx -o 输出.pptx
python app/main.py ppt cases/02_网络指标/     # 自动找配置

# 透视分析（原 pivot_main.py 的行为）
python app/main.py pivot cases/02_网络指标/配置.xlsx
python app/main.py pivot -c 配置.xlsx -o 结果.xlsx

# 向后兼容：不带子命令时，根据配置文件内容自动判断类型
python app/main.py cases/02_网络指标/配置.xlsx
```

**实现细节：**
- 顶层 `argparse` 定义 `ppt` / `pivot` 子命令
- `ppt` 子命令 → 调用 `_run_ppt_mode()`
- `pivot` 子命令 → 调用 `_run_pivot_mode()`
- 不带子命令时 → 尝试自动检测配置类型（读第一行表头），分发到对应模式
- `find_config_file()` 和 `_auto_find_data_file()` 只定义一次，两个模式共用

### 2. 删除文件

| 文件 | 原因 |
|------|------|
| `.trae/skills/excel2ppt/SKILL.md` | 移除 TRAE 技能包装 |
| `.trae/skills/excel-pivot/SKILL.md` | 移除 TRAE 技能包装 |
| `.trae/skills/` 整个目录 | 清空技能目录 |
| `app/pivot_main.py` | 功能已合并到 `main.py` 子命令 |

### 3. 保留但清理的文件

| 操作 | 文件 | 说明 |
|------|------|------|
| 保留 | `app/src/excel_reader.py` | 共享模块，不变 |
| 保留 | `app/src/ppt_builder.py` | PPT 专用，不变 |
| 保留 | `app/src/map_builder.py` | 地图专用，不变 |
| 保留 | `app/src/pivot_analyzer.py` | 透视专用，不变 |
| 保留 | `app/src/excel_writer.py` | 透视专用，不变 |
| 保留 | `app/src/__init__.py` | 包标识 |
| 保留 | `app/requirements.txt` | 不变 |
| **清理** | `cases/02_网络指标/` | 删除 `test_*.pptx`、`test_*.png`、`配置_报告_*.pptx` 等临时文件，只保留原始配置和数据 |
| **删除** | `cases/01_销售数据/` | 用户要求只保留一个代表性案例 |
| **删除** | `cases/03_网络透视/` | 同上 |
| **删除** | `cases/pivot_demo/` | 同上 |

### 4. 新建文件

#### `README.md`

项目根目录顶层文档，包含：
- 项目简介（Excel 数据分析 + PPT 报告生成 + 透视分析 三合一）
- 安装说明
- 快速开始（PPT 生成 / 透视分析）
- 配置格式（PPT 配置列说明 + 透视配置列说明）
- 图表类型列表
- 地图可视化说明
- 命令行参考

## 实施步骤

| 步骤 | 操作 | 文件 |
|------|------|------|
| 1 | 重写 `app/main.py` 为统一入口（子命令 + 自动检测） | `app/main.py` |
| 2 | 删除 `app/pivot_main.py` | 删除 |
| 3 | 删除 `.trae/skills/` 整个目录 | 删除 |
| 4 | 清理 `cases/02_网络指标/` 临时文件 | 删除 |
| 5 | 删除 `cases/01_销售数据/`、`03_网络透视/`、`pivot_demo/` | 删除 |
| 6 | 新建 `README.md` | `README.md` |
| 7 | 用保留的案例数据跑一遍验证（PPT + Pivot 都要通过） | 测试 |

## 验证标准

```bash
# PPT 生成（子命令方式）
python app/main.py ppt cases/02_网络指标/配置_v3.xlsx
# 期望：生成 .pptx，含 11 页 16 个图表（含 2 个地图）

# PPT 生成（自动检测）
python app/main.py cases/02_网络指标/配置_v3.xlsx
# 期望：同上

# 透视分析（子命令方式）
# 需要有透视配置——如果不保留 03_网络透视 案例，
# 则需要确保 at least one 案例能跑通 pivot 子命令
```

## 决策记录

| 决策 | 选择 | 理由 |
|------|------|------|
| 入口方式 | `python main.py ppt/pivot ...` | 单文件，子命令清晰，用户选择 |
| skills 目录 | 彻底删除 | 用户明确"不是 skill" |
| 案例保留 | 仅 02_网络指标 | 一个案例同时展示 PPT + 透视功能 |
| 向后兼容 | 不带子命令时自动检测 | 对老用户友好 |
