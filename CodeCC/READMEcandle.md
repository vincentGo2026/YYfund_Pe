# 申万二级行业 PE/PB 统计 & 可视化

基于通联数据（DataYes）导出的申万二级行业指数日频数据，进行 PE/PB 历史分位统计、估值蜡烛图绘制、柱状+折线双轴图，以及季度估值 vs 收益率分析。

---

## 目录结构

```
CodeCC/
├── 模块 A：数据拉取 ────────────────────────────────
├── 2fetch_sw2_market_data.py              # 拉取申万二级行业行情数据 (DataYes API)
│
├── 模块 B：PE/PB 统计 & 可视化（蜡烛图流水线）────────
│   ├── PE 版 ────────────────────────────────────
│   ├── pe_stats.py                       # PE 基础统计 + 4 张图表
│   ├── enhance_pe_stats.py               # PE/PB 完整百分位统计
│   ├── draw_candlestick.py               # PE 蜡烛图 3 张（含最新PE菱形标注）
│   ├── draw_barwithline.py               # PE 柱状图 + 百分位折线图 2 张
│   │
│   ├── PB 版 ────────────────────────────────────
│   ├── pb_stats.py                       # PB 基础统计 + 4 张图表
│   ├── enhance_pb_stats.py               # PB 百分位统计扩充
│   ├── draw_candlestick_pb.py            # PB 蜡烛图 3 张（含最新PB菱形标注）
│   ├── draw_barwithline_pb.py            # PB 柱状图 + 百分位折线图 2 张
│   │
│   └── 一键运行入口 ──────────────────────────────
│       run_all_SWcandle.py               # 一键运行 PE 版（4步顺序执行）
│       run_all_SWcandle_pb.py            # 一键运行 PB 版（4步顺序执行）
│
├── 模块 C：季度估值 vs 收益率分析 ──────────────────
├── 2quarterly_pe_vs_return_datayes.py    # 季度末 PE 绝对值 vs 次季度收益率
├── 2quarterly_pe_vs_return_SW2T.py       # 季度末 PE vs 收益率（备用价格源）
├── 2quarterly_pb_vs_return_datayes.py    # 季度末 PB 绝对值 vs 次季度收益率
├── 2quarterly_pb_pct_vs_return_datayes.py # 季度末 PB 历史百分位 vs 次季度收益率
│
├── 模块 D：实盘主动持仓估值分析 ────────────────────
├── plot_sw2_valuation.py                 # 实盘主动 申万二级行业估值分布 4 面板图
├── plot_sw2_scatter.py                   # PB vs PE 估值分位散点气泡图
├── shipan_sw2_assessment.py              # 实盘主动 申万二级行业投资评估（Excel 输出）
├── build_valuation_page.py               # 生成估值分析可视化网页（汇总所有图表&文件链接）
│
├── READMEcandle.md                       # 本文件
└── 新建文件夹/                             # 历史备份
```

---

## 一键运行

### PE 版蜡烛图流水线

```bash
cd D:\CC\Mid\估值\CodeCC
D:\Users\dingd\anaconda3\python.exe run_all_SWcandle.py
```

按顺序执行：`pe_stats` → `enhance_pe_stats` → `draw_candlestick` → `draw_barwithline`

### PB 版蜡烛图流水线

```bash
cd D:\CC\Mid\估值\CodeCC
D:\Users\dingd\anaconda3\python.exe run_all_SWcandle_pb.py
```

按顺序执行：`pb_stats` → `enhance_pb_stats` → `draw_candlestick_pb` → `draw_barwithline_pb`

---

## 数据源

| # | 路径 | 使用者 | 说明 |
|---|------|--------|------|
| ① | `D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv` | 模块B全部 + 模块C全部 | **核心源**：申万二级行业日频 PE/PB 序列（124行业，含 `申万一级行业` 映射） |
| ② | `D:\CC\DB\data\sw2_market_data_20210501_*.csv` | 模块C | 申万二级行业日频收盘指数（由模块A产出） |
| ③ | `D:\CC\DB\data\priceSW2T.xlsx` | `2quarterly_pe_vs_return_SW2T` | 备用价格源（Excel格式） |
| ④ | `D:\CC\DB\data\sw2_ticker_mapping.csv` | `2fetch_sw2_market_data` | 二级行业 ticker 代码映射表 |

数据清洗规则（模块B所有步骤统一）：
- 剔除 PE/PB ≤ 0
- 剔除上下 1% 极端值

---

## 各脚本详细说明

---

### 模块 A：数据拉取

#### `2fetch_sw2_market_data.py` — 拉取行情数据

**输入**：`D:\CC\DB\data\sw2_ticker_mapping.csv`（行业代码 → 名称映射）

**输出**：`D:\CC\DB\data\sw2_market_data_{begin}_{end}.csv`（长格式：tradeDate, closeIndex, 申万二级行业, ticker）

通过 DataYes API（`getMktIdxd.json`，exchangeCD=ZICN）批量拉取申万二级行业指数日线行情，默认从 2021-05-01 至当天。

---

### 模块 B：PE/PB 统计 & 可视化

---

#### PE 版（`run_all_SWcandle.py` 一键运行）

#### B-1. `pe_stats.py` — PE 基础统计

**输入**：数据源 ①

**输出 Excel**：`D:\CC\Mid\估值\PE_Statistics_by_Sector.xlsx`
| Sheet | 内容 |
|-------|------|
| `PE统计` | 124 行业 × 6 项指标（最大值/最小值/中位数/平均数/1/4分位/3/4分位） |
| `说明` | 数据源、时间区间、清洗规则 |

**输出图表**（4 张，均在 `D:\CC\Mid\估值\`）：

| 图表文件 | 内容 |
|----------|------|
| `Chart_PE_Median_Quartiles.png` | Top 30 行业 PE 中位数 + Q1/Q3 区间条形图 |
| `Chart_PE_Mean_Bar.png` | Top 30 行业 PE 平均值条形图 |
| `Chart_PE_Multi_Metrics.png` | Top 15 行业 6 项指标分组柱状图 |
| `Chart_PE_Median_Distribution.png` | 全行业 PE 中位数分布直方图 |

---

#### B-2. `enhance_pe_stats.py` — 扩充 PE/PB 统计

**输入**：数据源 ①

**处理逻辑**：
- 同时统计 PE 和 PB 两个维度
- 计算最新 PE/PB 值、历史百分位、各分位数（min/Q1/median/Q3/max）、样本天数
- 行业名标准化：`申万XXX` → `XXX(申万)`
- 按一级行业分组，PE 中位数降序排列

**输出**：覆盖写入 `PE_Statistics_by_Sector.xlsx`
| Sheet | 内容 |
|-------|------|
| `PE_PB统计` | 124 行业 × 19 列完整统计表 |
| `说明` | 数据源及清洗规则 |

**`PE_PB统计` 列说明**：

| 列名 | 说明 |
|------|------|
| 所属一级行业 | 申万一级行业归属 |
| 申万二级行业_API名 | 标准化行业名（如 `半导体(申万)`） |
| 行业代码 | 如 `801991` |
| 最新日期 | 数据截止日 |
| 最新PE / PE历史百分位(%) | 当前 PE 值及历史分位 |
| 历史最低PE / 1/4分位PE / 中位数PE / 3/4分位PE / 历史最高PE | PE 五档分位 |
| 最新PB / PB历史百分位(%) | 当前 PB 值及历史分位 |
| 历史最低PB / 1/4分位PB / 中位数PB / 3/4分位PB / 历史最高PB | PB 五档分位 |
| PE样本天数 / PB样本天数 | 有效数据天数 |

---

#### B-3. `draw_candlestick.py` — PE 蜡烛图

**输入**：数据源 ①（直接计算一级行业 PE）+ `PE_Statistics_by_Sector.xlsx` sheet `PE_PB统计`（读取二级行业 PE 统计）

**输出图表**（3 张）：

| 图表文件 | 内容 |
|----------|------|
| `Chart_PE_Candlestick_L1.png` | 申万一级行业 PE 蜡烛图（31行业，不标数值） |
| `Chart_PE_Candlestick_L1_with_values.png` | 申万一级行业 PE 蜡烛图（31行业，标注 min/中位数/max） |
| `Chart_PE_Candlestick.png` | 申万二级行业 PE 蜡烛图（124行业，标注数值） |

**蜡烛图说明**（水平箱线图风格）：
- 细线：`历史最低PE → 1/4分位PE` 和 `3/4分位PE → 历史最高PE`
- 彩色箱体：`1/4分位PE → 3/4分位PE`（红=高PE，绿=低PE）
- 黑色竖线：`中位数PE`

---

#### B-4. `draw_barwithline.py` — PE 柱状图 + 百分位折线

**输入**：`PE_Statistics_by_Sector.xlsx` sheet `PE_PB统计`（二级行业 Top 40）+ 数据源 ①（自算一级行业最新PE及百分位）

**输出图表**（2 张）：

| 图表文件 | 内容 |
|----------|------|
| `Chart_PE_L2_BarWithLine.png` | 二级行业（Top 40）：柱状图(最新PE) + 折线图(PE历史百分位) |
| `Chart_PE_L1_BarWithLine.png` | 一级行业（全量31行业）：柱状图(最新PE) + 折线图(PE历史百分位) |

**双轴图说明**：
- 左轴（蓝色柱状图）：最新 PE 值
- 右轴（红色折线+散点）：PE 历史百分位(%)，50% 处虚线标注

---

#### PB 版（`run_all_SWcandle_pb.py` 一键运行）

以下 4 个脚本与 PE 版完全对应，仅将估值指标从 PE 替换为 PB。

| PB 版脚本 | 对应 PE 版 | 功能 |
|-----------|-----------|------|
| `pb_stats.py` | `pe_stats.py` | PB 基础统计 + 4 张图表 |
| `enhance_pb_stats.py` | `enhance_pe_stats.py` | PB 百分位统计扩充 |
| `draw_candlestick_pb.py` | `draw_candlestick.py` | PB 蜡烛图 3 张（一级/二级） |
| `draw_barwithline_pb.py` | `draw_barwithline.py` | PB 柱状图 + 百分位折线图 2 张 |

**PB 版输出 Excel**：`D:\CC\Mid\估值\PB_Statistics_by_Sector.xlsx`

**PB 版输出图表**（9 张）：

| 图表文件 | 内容 |
|----------|------|
| `Chart_PB_Median_Quartiles.png` | PB 中位数 + Q1/Q3 区间 |
| `Chart_PB_Mean_Bar.png` | PB 平均值条形图 |
| `Chart_PB_Multi_Metrics.png` | PB 多指标分组柱状图 |
| `Chart_PB_Median_Distribution.png` | PB 中位数分布直方图 |
| `Chart_PB_Candlestick_L1.png` | 一级行业 PB 蜡烛图 |
| `Chart_PB_Candlestick_L1_with_values.png` | 一级行业 PB 蜡烛图（标数值） |
| `Chart_PB_Candlestick.png` | 二级行业 PB 蜡烛图 |
| `Chart_PB_L2_BarWithLine.png` | 二级行业 PB 柱状+折线 |
| `Chart_PB_L1_BarWithLine.png` | 一级行业 PB 柱状+折线 |

---

### 模块 C：季度估值 vs 收益率分析

以下 5 个脚本分析季度末估值（PE 或 PB）与次季度收益率的关系，生成散点图并输出 Excel。

**公共逻辑**：从 2022-06-30 开始，对每个季度的每个行业，取其季度末估值指标（PE/PB/PB百分位），计算次季度指数收益率，对每个一级行业生成一个 Excel 文件（含散点图 + 数据表），同时输出汇总 CSV。

---

#### C-1. `2quarterly_pe_vs_return_datayes.py` — PE 绝对值 vs 收益率

| 属性 | 值 |
|------|-----|
| 估值指标 | PE 绝对值 |
| X 轴 | 季度末 PE |
| 价格数据源 | `sw2_market_data_*.csv` |
| 输出目录 | `D:\CC\Mid\估值\REStemp\` |
| 汇总 CSV | `quarterly_pe_vs_return_data_v2.csv` |

---

#### C-2. `2quarterly_pe_vs_return_SW2T.py` — PE vs 收益率（备用价格源）

| 属性 | 值 |
|------|-----|
| 估值指标 | PE 绝对值 |
| X 轴 | 季度末 PE |
| 价格数据源 | `priceSW2T.xlsx`（Excel格式，特殊布局） |
| 输出目录 | `D:\CC\Mid\估值\REStemp\` |
| 汇总 CSV | `quarterly_pe_vs_return_data.csv` |
| 备注 | 使用备用价格源，散点图用 viridis 色板 |

---

#### C-3. `2quarterly_pb_vs_return_datayes.py` — PB 绝对值 vs 收益率

| 属性 | 值 |
|------|-----|
| 估值指标 | PB 绝对值 |
| X 轴 | 季度末 PB |
| 价格数据源 | `sw2_market_data_*.csv` |
| 输出目录 | `D:\CC\Mid\估值\REStemp\pb\` |
| 汇总 CSV | `quarterly_pb_vs_return_data_v2.csv` |

---

#### C-4. `2quarterly_pb_pct_vs_return_datayes.py` — PB 历史百分位 vs 收益率

| 属性 | 值 |
|------|-----|
| 估值指标 | PB 历史百分位 |
| X 轴 | 季度末 PB 在自身历史中的百分位（%越高 → 估值越贵） |
| 计算方式 | `历史PB ≤ 当前PB的天数 / 历史总天数 × 100` |
| 特色 | 散点图标注低估/高估分区（绿/红底色），点按时间渐变（蓝→橙） |
| 输出目录 | `D:\CC\Mid\估值\REStemp\pb_pct\` |
| 汇总 CSV | `quarterly_pb_pct_vs_return.csv` |

---

### 模块 C 文件对比

| 脚本 | 估值指标 | X 轴 | 价格数据源 | 色板 |
|------|---------|------|-----------|------|
| `2quarterly_pe_vs_return_datayes` | PE | PE 绝对值 | sw2_market CSV | 蓝→橙 渐变 |
| `2quarterly_pe_vs_return_SW2T` | PE | PE 绝对值 | priceSW2T Excel | viridis |
| `2quarterly_pb_vs_return_datayes` | PB | PB 绝对值 | sw2_market CSV | 蓝→橙 渐变 |
| `2quarterly_pb_pct_vs_return_datayes` | PB | PB 历史百分位 | sw2_market CSV | 蓝→橙 渐变+分区 |

---

### 模块 D：实盘主动持仓估值分析

以下 3 个脚本基于 `Fund_IRG_updated_v2.xlsx` 的实盘主动持仓数据，绘制申万二级行业估值分布图表，并生成带评级/动量/风险的完整评估 Excel。

**公共数据源**：`D:\CC\DB\data\Fund_IRG_updated_v2.xlsx` sheet `申万二级行业分布`（实体盘主动、市值占比>0 的行）

**估值方式规则**：以下 9 个一级行业看 PB，其余看 PE：
> 银行、非银金融、电力设备、交通运输、有色金属、建筑装饰、食品饮料、家用电器、商贸零售

---

#### D-1. `plot_sw2_valuation.py` — 实盘主动 申万二级行业估值分布图

**输入**：`D:\CC\DB\data\Fund_IRG_updated_v2.xlsx` sheet `申万二级行业分布`（筛选 实盘主动 非零持仓）

**输出图表**（1 张 2×2 面板大图）：

| 子图位置 | 内容 |
|----------|------|
| 左上 | **最新PB** 柱状图（蓝色=PB估值行业，灰色=PE估值行业），柱顶红色数字=PB百分位 |
| 右上 | **PB历史百分位(%)** 柱状图，50%/80% 红色虚线参考线 |
| 左下 | **最新PE** 柱状图（橙色=PE估值行业，灰色=PB估值行业），柱顶红色数字=PE百分位 |
| 右下 | **PE历史百分位(%)** 柱状图，50%/80% 红色虚线参考线 |

**输出文件**：`D:\CC\Mid\估值\sw2_valuation_chart.png`

图表按市值占比从大到小排序，底部标注图例说明。

---

#### D-2. `plot_sw2_scatter.py` — PB vs PE 估值分位散点图

**输入**：`D:\CC\DB\data\Fund_IRG_updated_v2.xlsx` sheet `申万二级行业分布`（筛选 实盘主动 非零持仓）

**输出图表**（1 张散点气泡图）：

| 属性 | 说明 |
|------|------|
| X 轴 | PB 历史百分位(%) |
| Y 轴 | PE 历史百分位(%) |
| 气泡大小 | 市值占比（越大=持仓越重） |
| 颜色 | 蓝色=PB估值行业 / 橙色=PE估值行业 |
| 象限分割 | 50% 十字虚线，标注「PB低估/PE高估」等四象限含义 |
| 标注 | 每个点标注二级行业名称 |

**输出文件**：`D:\CC\Mid\估值\sw2_valuation_scatter.png`

---

#### D-3. `shipan_sw2_assessment.py` — 实盘主动 申万二级行业评估

**输入**：
| # | 路径 | 用途 |
|---|------|------|
| ① | `D:\CC\Mid\估值\申万\行业投资价值综合评估_混合版.xlsx` | 一级行业 投资评级/综合评分/动量/风险/收益 映射 |
| ② | `D:\CC\Mid\估值\实盘主动_申万二级行业估值.xlsx` | 实盘主动 45 个二级行业的 PE/PB 及百分位 |

**处理逻辑**：
1. 读取混合版一级行业评估（31 行业），建立 投资评级/综合评分/动量/风险/收益 映射表
2. 读取实盘主动估值明细，按一级行业匹配估值方式（PB 行业 / PE 行业）
3. 百分位 → 估值状态：\<20% 深度低估 / 20-40% 低估 / 40-60% 合理 / 60-80% 偏高 / ≥80% 高估
4. 估值状态 → 估值评分：深度低估=3、低估/合理=2、偏高=1、高估=0
5. 生成两个 Excel 文件

**输出 Excel**（2 个文件，均在 `D:\CC\Mid\估值\`）：

| 文件 | Sheet | 列数 | 内容 |
|------|-------|------|------|
| `实盘主动_申万二级行业评估.xlsx` | 实盘主动评估 | 22 列 | 完整评估表：申万一/二级行业、市值占比、投资评级、综合/估值/动量评分、风险等级、估值状态/依据、PE/PB 最新值/分位/状态、动量状态、2026收益、3年累计、波动率、年胜率、收益风险比 |
| `实盘主动_申万二级行业估值明细.xlsx` | 持仓估值 | 10 列 | 精简估值明细：申万一/二级行业、市值占比、估值依据、最新PE/分位/状态、最新PB/分位/状态（按市值占比降序） |

**着色规则**：
- 投资评级：绿(积极配置) / 黄(标配) / 粉(谨慎观望) / 红(规避)
- 估值状态：深绿(深度低估) / 浅绿(低估) / 黄(合理) / 橙(偏高) / 红(高估)
- 估值依据：蓝底(PB) / 绿底(PE)
- 风险等级：绿(低) / 黄(中) / 粉(高)

---

#### D-4. `build_valuation_page.py` — 估值分析可视化网页

**输入**：模块 B/C/D 生成的所有图表、Excel 及 Word 报告（无需参数，纯静态生成）

**输出**：`valuation_charts.html`（单文件 HTML 网页，可直接浏览器打开或局域网发布）

**网页内容**（6 个板块 + 粘性导航栏）：

| 板块 | ID | 内容 |
|------|-----|------|
| PE/PB 蜡烛图 | `#pe_charts` | PE 一级+二级蜡烛图、PB 一级+二级蜡烛图（4 张） |
| 实盘主动持仓估值 | `#shipan_valuation` | PB 分布图、PE 分布图、PB vs PE 散点图 + 估值明细 Excel 链接 |
| 投资价值评估 | `#valuation` | 混合版/PE版/PB版评估 Excel + PE/PB 统计 + 月度/季度/年度收益率矩阵（8 个文件链接） |
| 分析报告 | `#reports` | 季度/年度收益率报告 + 混合版/PE版/PB版风险评估（5 个 Word 文档链接） |
| 热力图 | `#heatmaps` | 月度/季度/年度收益率热力图（3 张 PNG） |
| 数据说明 | `#data` | 行业分类、数据频率、收益率计算、评估维度、色阶说明 |

**技术特点**：
- 响应式布局（移动端自适应）、图片点击放大
- 文件链接相对路径引用，同级目录引用 `D:\CC\Mid\估值\CodeCC\` 下的 Excel/Word/热力图，模块 B/D 图表通过 `../` 引用上级 `估值\` 目录
- 可直接通过 `python -m http.server` 在局域网发布

---

## 数据流总览

```
┌─────────────────────────────────────────────────────────────┐
│ 模块 A：数据拉取                                              │
│                                                              │
│ sw2_ticker_mapping.csv                                       │
│   └──→ 2fetch_sw2_market_data.py ──→ sw2_market_data_*.csv  │
└─────────────────────────────────────────────────────────────┘
                    │                              │
                    ▼                              │
┌─────────────────────────────────────────────────────────────┐
│ 核心数据源                                                    │
│                                                              │
│ datayes_all_SW_Industries_Level2_mapped.csv ─────────────┐  │
│      │                                                   │  │
│      ▼                                                   │  │
│ ┌────────────── 模块 B：PE 版 ──────────────────┐        │  │
│ │ pe_stats.py ──→ PE_Statistics_by_Sector.xlsx  │        │  │
│ │ enhance_pe_stats.py ──→ (覆盖写入 PE_PB统计)   │        │  │
│ │ draw_candlestick.py ──→ 3张 PE 蜡烛图 PNG     │        │  │
│ │ draw_barwithline.py ──→ 2张 PE 柱状+折线 PNG  │        │  │
│ │ run_all_SWcandle.py（一键串联）                │        │  │
│ └───────────────────────────────────────────────┘        │  │
│                                                           │  │
│ ┌────────────── 模块 B：PB 版 ──────────────────┐        │  │
│ │ pb_stats.py ──→ PB_Statistics_by_Sector.xlsx │        │  │
│ │ enhance_pb_stats.py ──→ (覆盖写入 PB_PB统计)  │        │  │
│ │ draw_candlestick_pb.py ──→ 3张 PB 蜡烛图 PNG │        │  │
│ │ draw_barwithline_pb.py ──→ 2张 PB 柱状+折线  │        │  │
│ │ run_all_SWcandle_pb.py（一键串联）             │        │  │
│ └───────────────────────────────────────────────┘        │  │
│                                                           │  │
│ ┌────────────── 模块 C：季度分析 ────────────────────────────┘  │
│ │                                                              │
│ │ datayes Level2 + sw2_market_data ──→ PE绝对值 vs 收益        │
│ │                                   ──→ PB绝对值 vs 收益        │
│ │                                   ──→ PB百分位 vs 收益        │
│ │                                                              │
│ │ datayes Level2 + priceSW2T ────────→ PE绝对值 vs 收益(备用)    │
│ │                                                              │
│ │ 输出：每个一级行业一个 .xlsx + 汇总 CSV                         │
│ └──────────────────────────────────────────────────────────────┘
│
│ ┌────────────── 模块 D：实盘主动持仓估值 ─────────────────────┐
│ │                                                              │
│ │ Fund_IRG_updated_v2.xlsx (实盘主动 二级行业分布)               │
│ │   ├──→ plot_sw2_valuation.py ──→ sw2_valuation_chart.png    │
│ │   ├──→ plot_sw2_scatter.py   ──→ sw2_valuation_scatter.png  │
│ │   └──→ shipan_sw2_assessment.py ──→ 评估/明细 2个 Excel     │
│ │        (结合 混合版评估 一级行业映射)                           │
│ │                                                              │
│ │ 输出：2张 PNG + 2个 Excel                                    │
│ └──────────────────────────────────────────────────────────────┘
│
│ ┌────────────── 汇总网页（消费全模块产物） ────────────────────┐
│ │                                                              │
│ │ 模块 B 图表 + 模块 C 文件 + 模块 D 图表/Excel                │
│ │   └──→ build_valuation_page.py ──→ valuation_charts.html   │
│ │        (单文件 HTML，聚合所有图表预览 + 文件链接)               │
│ └──────────────────────────────────────────────────────────────┘
```

---

## 输出文件总览

### 模块 B（PE + PB）

`D:\CC\Mid\估值\` 目录下：

| 类型 | 文件 |
|------|------|
| PE 统计 Excel | `PE_Statistics_by_Sector.xlsx` |
| PB 统计 Excel | `PB_Statistics_by_Sector.xlsx` |
| PE 图表 ×4 | `Chart_PE_Median_Quartiles.png`, `Chart_PE_Mean_Bar.png`, `Chart_PE_Multi_Metrics.png`, `Chart_PE_Median_Distribution.png` |
| PE 蜡烛图 ×3 | `Chart_PE_Candlestick_L1.png`, `Chart_PE_Candlestick_L1_with_values.png`, `Chart_PE_Candlestick.png` |
| PE 柱状+折线 ×2 | `Chart_PE_L2_BarWithLine.png`, `Chart_PE_L1_BarWithLine.png` |
| PB 图表 ×4 | `Chart_PB_Median_Quartiles.png`, `Chart_PB_Mean_Bar.png`, `Chart_PB_Multi_Metrics.png`, `Chart_PB_Median_Distribution.png` |
| PB 蜡烛图 ×3 | `Chart_PB_Candlestick_L1.png`, `Chart_PB_Candlestick_L1_with_values.png`, `Chart_PB_Candlestick.png` |
| PB 柱状+折线 ×2 | `Chart_PB_L2_BarWithLine.png`, `Chart_PB_L1_BarWithLine.png` |

### 模块 C

`D:\CC\Mid\估值\REStemp\` 目录下：

| 类型 | 路径 | 说明 |
|------|------|------|
| 汇总 CSV | `quarterly_pe_vs_return_data_v2.csv` | PE 绝对值 vs 收益 |
| 汇总 CSV | `quarterly_pe_vs_return_data.csv` | PE vs 收益（备用价格源） |
| 汇总 CSV | `pb\quarterly_pb_vs_return_data_v2.csv` | PB 绝对值 vs 收益 |
| 汇总 CSV | `pb_pct\quarterly_pb_pct_vs_return.csv` | PB 百分位 vs 收益 |
| 行业 Excel | `REStemp\*.xlsx` | PE 版每个一级行业一个文件 |
| 行业 Excel | `REStemp\pb\*.xlsx` | PB 版每个一级行业一个文件 |
| 行业 Excel | `REStemp\pb_pct\*.xlsx` | PB 百分位版每个一级行业一个文件 |

### 模块 D（实盘主动估值）

`D:\CC\Mid\估值\` 目录下：

| 类型 | 文件 |
|------|------|
| 估值分布图 | `sw2_valuation_chart.png` |
| 散点气泡图 | `sw2_valuation_scatter.png` |
| 评估 Excel | `实盘主动_申万二级行业评估.xlsx`（22列，含评级/评分/动量/风险） |
| 估值明细 Excel | `实盘主动_申万二级行业估值明细.xlsx`（10列，PE/PB/状态，按市值占比降序） |
| 汇总网页 | `valuation_charts.html`（单文件，聚合所有图表+报告链接，可局域网发布） |

---

## 依赖

```bash
pip install pandas numpy matplotlib openpyxl requests
```

- 中文字体：微软雅黑（`msyh.ttc`）或黑体（`simhei.ttf`）
- DataYes API Token（仅模块 A、C 需要）

---

## 相关项目

- `D:\CC\Mid\估值\Code估值\` — 申万一级行业热力图 & PE/PB 投资价值评估（`run_all.py`）
- `D:\CC\Mid\估值\Code估值\` — DataYes 数据拉取 + SW1 映射（`run_datayesfetch_and_map.py`）
