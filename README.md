# YYfund_Pe — 申万行业估值分析系统

基于申万一级/二级行业的价格、PE/PB 估值与收益率数据，进行统计分析、可视化与投资价值评估。本仓库包含两个子项目：

| 子目录 | 说明 | 详细文档 |
|--------|------|----------|
| `CodeCC/` | 申万二级行业 PE/PB 统计、蜡烛图/柱状图可视化、季度估值 vs 收益率分析、实盘持仓估值评估 | [READMEcandle.md](CodeCC/READMEcandle.md) |
| `Code估值/` | 申万一级/二级行业数据拉取、收益率热力图、PE/PB 综合投资价值与风险评估报告 | [README_SW.md](Code估值/README_SW.md) |

> 注：脚本中的数据路径（如 `D:\CC\DB\data\...`）为本地环境路径，请按需调整。

---

## 环境变量配置

部分脚本通过环境变量读取敏感配置，**不要**将真实密钥提交到仓库。

在运行前设置以下环境变量（或参考 `.env.example`）：

| 变量 | 用途 | 必需 |
|------|------|------|
| `DATAYES_TOKEN` | 通联数据（DataYes）API Token，数据拉取脚本使用 | 是（数据拉取） |
| `AUTH_USER` / `AUTH_PASS` | 公网发布服务器（`5publish_public.py`）的 Basic Auth 凭据 | 可选（有默认值） |

Windows PowerShell 示例：

```powershell
$env:DATAYES_TOKEN = "你的_token"
$env:AUTH_PASS = "你的密码"
python Code估值/2fetch_datayes_data_SWL1.py
```

---

## 依赖

```bash
pip install pandas numpy matplotlib seaborn openpyxl requests python-docx flask
```

- 中文字体：微软雅黑（`msyh.ttc`）或黑体（`simhei.ttf`）
- DataYes API Token（仅数据拉取脚本需要）

---

## 快速开始

### Code估值 — 一键分析

```bash
cd Code估值
python run_all.py                 # 热力图 + 报告 + 估值评估
python run_alldatayesfetch_and_map.py   # 数据拉取 + 映射（需 DATAYES_TOKEN）
```

### CodeCC — PE/PB 蜡烛图流水线

```bash
cd CodeCC
python run_all_SWcandle.py        # PE 版（4 步）
python run_all_SWcandle_pb.py     # PB 版（4 步）
```

各子项目的完整脚本说明、数据源与输出文件，见各自的 README。
