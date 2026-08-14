# 申万行业投资分析系统

基于申万一级/二级行业价格、PE/PB 估值及收益率数据，生成热力图、统计报告和综合投资价值评估。

---

## 目录结构

```
Code估值/
├── 模块 A：数据拉取与映射 ──────────────────────────
├── 2fetch_datayes_data_SWL1.py     # 拉取申万一级行业日线数据 (DataYes API)
├── 2fetch_datayes_data_SWL2.py     # 拉取申万二级行业日线数据 (DataYes API)
├── 2add_sw1_mapping.py             # 为L2数据补充"申万一级行业"列
├── run_datayesfetch_and_map.py     # 一键运行：并行拉取 L1+L2 → L2 映射
│
├── 模块 B：热力图生成 ──────────────────────────────
├── sector_order.py                 # 行业板块分类排序模块（被热力图脚本导入）
├── plot_quarterly_heatmap.py       # 季度收益率热力图 (2022-至今)
├── plot_monthly_heatmap.py         # 月度收益率热力图 (2025-至今)
├── plot_monthly_heatmap_pe_pb.py   # 月度收益率热力图 + PE/PB 混合估值复合图
├── plot_yearly_heatmap.py          # 年度收益率热力图 (2016-2026)
│
├── 模块 C：分析报告与估值评估 ──────────────────────
├── export_quarterly_analysis.py    # 季度收益率分析报告 (Word)
├── export_yearly_analysis.py       # 年度收益率分析报告 (Word)
├── pe_pb_2026_analysis.py          # PE/PB 综合投资价值与风险评估
├── hybrid_analysis.py              # 混合估值版评分报告（部分PB、其余PE）
│
├── 模块 D：一键运行入口 ────────────────────────────
├── run_all.py                      # 一键运行全部分析（热力图+报告+估值）
├── run_datayesfetch_and_map.py     # 一键运行数据拉取+映射
│
├── README_SW.md                    # 本文件
├── 申万行业投资分析系统_产品说明书_v2.0.docx  # 详细产品说明
└── 备份/                           # 历史备份脚本
```

---

## 数据源

### 输入数据

| 数据 | 路径 | 使用者 |
|---|---|---|
| 申万一级行业指数代码 | `申银万国一级行业指数代码.xlsx` | L1 数据拉取 |  
| 申万二级行业指数代码 | `申银万国二级行业指数代码.xlsx` | L2 数据拉取 |
| 申万三级/中信行业分类映射 | `D:\CC\DB\data\申万三级和中信一级行业板块分类参考.xlsx` → "二级" sheet | L2 映射 |
| 申万一级行业价格 | `D:\CC\DB\MKT\PriceSW1.xlsx` | 季度/月度/年度热力图 |
| 申万行业 PE | `D:\CC\DB\MKT\申万行业指数PE_2026.xlsx` | PE/PB 评估 |
| 申万行业 PB | `D:\CC\DB\MKT\申万行业指数PB_2026.xlsx` | PE/PB 评估 |
| 年度收益率矩阵 | `D:\CC\Mid\估值\中信\SW1\申万一级行业年度收益率矩阵.xlsx` | PE/PB 评估、混合评估 |
| 季度收益率矩阵 | `D:\CC\Mid\估值\中信\SW1\申万一级行业季度收益率矩阵.xlsx` | PE/PB 评估、混合评估 |

### 拉取后产生的中间数据

| 数据 | 路径 | 产生方式 |
|---|---|---|
| L1 CSV | `D:\CC\DB\data\datayes_all_SW_Industries_*.csv` | `2fetch_datayes_data_SWL1.py` |
| L2 CSV | `D:\CC\DB\data\datayes_all_SW_Industries_Level2.csv` | `2fetch_datayes_data_SWL2.py` |
| L2 mapped CSV | `D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv` | `2add_sw1_mapping.py` |

---

## 输出

所有输出文件统一写入 `D:\CC\Mid\估值\申万\`：

### 热力图与矩阵

| 产出 | 文件名 |
|---|---|
| 季度收益率热力图 | `申万一级行业季度收益率热力图.png` |
| 月度收益率热力图 | `申万一级行业月度收益率热力图.png` |
| 月度收益率 + PE/PB 复合图 | `申万一级行业月度收益率_含PE_PB.png` |
| 年度收益率热力图 | `申万一级行业年度收益率热力图.png` |
| 季度收益率矩阵 | `申万一级行业季度收益率矩阵.xlsx` |
| 月度收益率矩阵（含估值数据） | `申万一级行业月度收益率矩阵.xlsx` |
| 年度收益率矩阵 | `申万一级行业年度收益率矩阵.xlsx` |

### 分析报告

| 产出 | 文件名 |
|---|---|
| 季度收益率分析报告 | `申万一级行业季度收益率分析报告.docx` |
| 年度收益率分析报告 | `申万一级行业年度收益率分析报告.docx` |

### 估值与投资价值评估

| 产出 | 文件名 |
|---|---|
| PE 统计表 | `申万行业PE统计.xlsx` |
| PB 统计表 | `申万行业PB统计.xlsx` |
| 投资价值综合评估 (PE+PB) | `行业投资价值综合评估_2026.xlsx` |
| 投资价值综合评估 (PB版) | `行业投资价值综合评估_PB_2026.xlsx` |
| 投资价值综合评估 (混合版) | `行业投资价值综合评估_混合版.xlsx` |
| 综合评估报告 (PE+PB版) | `申万一级行业综合投资价值与风险评估报告_PE_PB版.docx` |
| 综合评估报告 (PB版) | `申万一级行业综合投资价值与风险评估报告_PB版.docx` |
| 综合评估报告 (混合版) | `申万一级行业综合投资价值与风险评估报告_混合版.docx` |

---

## 运行方式

### 一键运行全部分析

```bash
python run_all.py
```

该命令按以下顺序依次执行：
1. 季度/月度/年度收益率热力图（生成 PNG + 矩阵 xlsx）
2. 月度收益率 + PE/PB 估值复合图
3. 季度/年度收益率分析报告（Word）
4. PE/PB 综合投资价值评估（Excel + Word）
5. 混合估值版评分报告（Excel + Word）

### 数据拉取流程（需要 DataYes API Token）

数据拉取独立于分析流程，适合在数据源过期时运行：

```bash
# 一键拉取 L1 + L2 + 映射
python run_alldatayesfetch_and_map.py
```

该命令自动：
1. **并行**拉取申万一级/二级行业日线数据（2021-05-01 至今）
2. 二级行业数据拉取完成后，自动补充"申万一级行业"映射列

### 分步手动运行

#### 模块 A：数据拉取（按需）

```bash
# 1. 并行拉取 L1 和 L2 数据（可手动分别运行）
python 2fetch_datayes_data_SWL1.py
python 2fetch_datayes_data_SWL2.py

# 2. L2 数据补充一级行业映射
python 2add_sw1_mapping.py
```

#### 模块 B + C：生成分析结果

```bash
# 1. 生成热力图及收益率矩阵
python plot_quarterly_heatmap.py
python plot_monthly_heatmap.py
python plot_yearly_heatmap.py

# 2. 生成月度收益率 + PE/PB 估值复合图（依赖步骤1的月度矩阵 + PE/PB 原始数据）
python plot_monthly_heatmap_pe_pb.py

# 3. 生成收益率分析报告（依赖步骤1的矩阵文件）
python export_quarterly_analysis.py
python export_yearly_analysis.py

# 4. 生成 PE/PB 投资价值评估（依赖原始 PE/PB 数据和收益率矩阵）
python pe_pb_2026_analysis.py

# 5. 生成混合估值版评分报告（依赖步骤4的 PE/PB 统计表）
python hybrid_analysis.py
```

---

## 脚本间依赖关系

```
┌─────────────────────────────────────────────────────────┐
│ 模块 A：数据拉取与映射                                    │
│                                                          │
│ 2fetch_datayes_data_SWL1.py ──→ datayes_all_SW_*.csv    │
│ 2fetch_datayes_data_SWL2.py ──→ Level2.csv              │
│                                      │                   │
│ 2add_sw1_mapping.py  ←──────────────┘                   │
│   └──→ Level2_mapped.csv                                │
│                                                          │
│ run_datayesfetch_and_map.py（并行调度上述三步）            │
└─────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────┐
│ 模块 B + C：热力图 → 报告 → 估值评估                      │
│                                                          │
│ plot_quarterly_heatmap.py ──→ 季度收益率矩阵.xlsx        │
│   │                                  │                   │
│   ├──→ 季度热力图.png               │                   │
│   └──→ ───────────────────────── export_quarterly_analysis.py
│                                                          │
│ plot_monthly_heatmap.py ──→ 月度收益率矩阵.xlsx          │
│   │                              │                       │
│   └──→ 月度热力图.png            │                       │
│                                  │                       │
│ plot_monthly_heatmap_pe_pb.py ←─┘ + PE/PB_2026.xlsx    │
│   └──→ 月度+PE_PB复合图.png                              │
│                                                          │
│ plot_yearly_heatmap.py ──→ 年度收益率矩阵.xlsx           │
│   │                                  │                   │
│   ├──→ 年度热力图.png               │                   │
│   └──→ ───────────────────────── export_yearly_analysis.py
│                                                          │
│ pe_pb_2026_analysis.py                                    │
│   ├── 输入: PE_2026.xlsx + PB_2026.xlsx                  │
│   │         年度/季度收益率矩阵.xlsx                       │
│   ├──→ 申万行业PE统计.xlsx ────┐                         │
│   ├──→ 申万行业PB统计.xlsx ────┤                         │
│   │                              │                       │
│   │  hybrid_analysis.py ←───────┘                        │
│   │    └──→ 混合版.xlsx + 混合版.docx                     │
│   │                                                      │
│   ├──→ 综合评估_2026.xlsx                                │
│   ├──→ PE_PB版报告.docx                                  │
│   └──→ PB版报告.docx                                     │
│                                                          │
│ run_all.py（顺序调度上述全部8个脚本）                       │
└─────────────────────────────────────────────────────────┘
```

---

## 混合估值规则

`hybrid_analysis.py` 中定义：以下 9 个行业使用 **PB 历史百分位**评分，其余 22 个行业使用 **PE 历史百分位**评分：

| PB 估值行业 | 理由 |
|---|---|
| 银行、非银金融 | 高杠杆行业，净资产价值为核心定价基础 |
| 电力设备、交通运输、建筑装饰 | 重资产行业，资产价值驱动 |
| 有色金属 | 强周期行业，盈利波动剧烈，PE 易失真 |
| 食品饮料、家用电器 | 估值已进入成熟期，PB 辅助参考 |
| 商贸零售 | 行业转型期盈利波动大，PB 反映资产底牌 |

综合评分 = 估值评分 × 40% + 动量评分 × 40% + (3 - 风险评分) × 20%

---

## 行业板块分类

`sector_order.py` 中定义，按上中下游板块分类：

```
上游资源：石油石化、煤炭、有色金属、基础化工、钢铁、农林牧渔
中游制造：电力设备、国防军工、机械设备、建筑装饰、建筑材料、交通运输、轻工制造、公用事业、环保
下游消费：汽车、商贸零售、家用电器、食品饮料、纺织服饰、医药生物、社会服务、美容护理
TMT：     电子、计算机、通信、传媒
金融地产：银行、非银金融、房地产
综合：     综合
```

---

## 环境依赖

- Python 3.x
- pandas / numpy
- matplotlib / seaborn
- python-docx
- openpyxl
- requests（仅数据拉取模块）
- 中文字体：微软雅黑（`msyh.ttc`）或黑体（`simhei.ttf`）

可通过 `FONT_SCALE` 环境变量调整热力图字体大小（默认 1.2~1.3）：

```bash
set FONT_SCALE=1.5 && python plot_quarterly_heatmap.py
```

---

## 相关项目

- `D:\CC\Mid\估值\CodeCC\` — 申万二级行业 PE/PB 蜡烛图与柱状图分析（`run_all_SWcandle.py` + `run_all_SWcandle_pb.py`）
- `D:\CC\Mid\估值\中信\` — 中信一级行业分析工具
