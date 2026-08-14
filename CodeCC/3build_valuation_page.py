"""生成 valuation_charts.html — 申万行业估值精简版网页"""
import os

OUTPUT = os.path.join(os.path.dirname(__file__), "valuation_charts.html")

html = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>申万一级行业估值分析</title>
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { background: #f5f5f5; font-family: -apple-system, "Microsoft YaHei", sans-serif; color: #333; }
  .header { background: linear-gradient(135deg, #0d2137, #1a3a5c); color: #fff; padding: 30px 20px; text-align: center; }
  .header h1 { font-size: 24px; margin-bottom: 6px; }
  .header p { font-size: 14px; opacity: 0.8; }
  .nav { position: sticky; top: 0; z-index: 10; background: #fff; padding: 12px 20px; box-shadow: 0 1px 6px rgba(0,0,0,0.1); display: flex; gap: 8px; flex-wrap: wrap; border-radius: 0 0 12px 12px; }
  .nav a { display: inline-block; padding: 6px 14px; background: #e8f0fe; color: #1a73e8; border-radius: 16px; text-decoration: none; font-size: 13px; transition: all 0.2s; }
  .nav a:hover { background: #1a73e8; color: #fff; }
  .container { max-width: 1400px; margin: 0 auto; padding: 20px; }
  .chart-card { background: #fff; border-radius: 12px; box-shadow: 0 2px 12px rgba(0,0,0,0.08); margin-bottom: 30px; overflow: hidden; }
  .chart-title { padding: 16px 24px; background: #fafafa; border-bottom: 1px solid #eee; font-size: 16px; font-weight: 600; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 8px; }
  .chart-title .badge { font-size: 12px; background: #e8f0fe; color: #1a73e8; padding: 2px 10px; border-radius: 10px; font-weight: 400; }
  .chart-body { padding: 16px; text-align: center; }
  .chart-body img { max-width: 100%; height: auto; border-radius: 6px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); cursor: pointer; }
  .file-list { padding: 16px 24px; }
  .file-list a { display: inline-block; margin: 4px; padding: 8px 16px; background: #f0f7ff; color: #1a73e8; border-radius: 8px; text-decoration: none; font-size: 13px; border: 1px solid #d0e3ff; }
  .file-list a:hover { background: #e0eeff; }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
  @media (max-width: 900px) { .grid2 { grid-template-columns: 1fr; } }
  footer { text-align: center; padding: 30px; color: #999; font-size: 12px; }
</style>
</head>
<body>
<div class="header">
  <h1>申万一级行业估值分析</h1>
  <p>PE/PB 估值图表 &amp; 实盘主动持仓估值分布</p>
</div>
<div class="nav">
  <a href="#pe_charts">PE/PB蜡烛图</a>
  <a href="#shipan_valuation">实盘主动持仓估值</a>
  <a href="#valuation">投资价值评估</a>
  <a href="#reports">分析报告</a>
  <a href="#heatmaps">热力图</a>
  <a href="#data">数据说明</a>
</div>
<div class="container">

<div id="pe_charts" class="chart-card">
<div class="chart-title"><span>PE / PB 蜡烛图</span><span class="badge">10年数据</span></div>
<div class="chart-body">

<h3 style="text-align:left;padding:8px 0;color:#1a3a5c;font-size:15px;border-bottom:1px solid #eee;">PE 蜡烛图</h3>
<div><img src="../Chart_PE_Candlestick.png" alt="PE二级蜡烛图" onclick="window.open(this.src)" style="width:100%"></div>
<div style="margin-top:12px;"><img src="../Chart_PE_Candlestick_L1_with_values.png" alt="PE一级蜡烛图" onclick="window.open(this.src)" style="width:100%"></div>

<h3 style="text-align:left;padding:16px 0 8px;color:#1a3a5c;font-size:15px;border-bottom:1px solid #eee;">PB 蜡烛图</h3>
<div><img src="../Chart_PB_Candlestick.png" alt="PB二级蜡烛图" onclick="window.open(this.src)" style="width:100%"></div>
<div style="margin-top:12px;"><img src="../Chart_PB_Candlestick_L1_with_values.png" alt="PB一级蜡烛图" onclick="window.open(this.src)" style="width:100%"></div>

</div></div>

<div id="shipan_valuation" class="chart-card">
<div class="chart-title"><span>实盘主动持仓 · 申万二级行业估值分布</span><span class="badge">PB+PE</span></div>
<div class="chart-body">

<h3 style="text-align:left;padding:8px 0;color:#1a3a5c;font-size:15px;border-bottom:1px solid #eee;">最新PB & PB历史百分位</h3>
<div><img src="../sw2_pb_combined.png" alt="PB估值分布" onclick="window.open(this.src)" style="width:100%"></div>

<h3 style="text-align:left;padding:16px 0 8px;color:#1a3a5c;font-size:15px;border-bottom:1px solid #eee;">最新PE & PE历史百分位</h3>
<div><img src="../sw2_pe_combined.png" alt="PE估值分布" onclick="window.open(this.src)" style="width:100%"></div>

<h3 style="text-align:left;padding:16px 0 8px;color:#1a3a5c;font-size:15px;border-bottom:1px solid #eee;">PB vs PE 估值分位矩阵</h3>
<div><img src="../sw2_valuation_scatter.png" alt="PB vs PE估值散点图" onclick="window.open(this.src)" style="width:100%"></div>

<h3 style="text-align:left;padding:16px 0 8px;color:#1a3a5c;font-size:15px;border-bottom:1px solid #eee;">持仓行业估值明细</h3>
<div class="file-list" style="padding:12px 0 0 0;">
<a href="../实盘主动_申万二级行业估值明细.xlsx">实盘主动_申万二级行业估值明细.xlsx</a>
</div>

</div></div>

<div id="valuation" class="chart-card">
<div class="chart-title"><span>行业投资价值综合评估</span><span class="badge">Excel</span></div>
<div class="file-list">
<a href="行业投资价值综合评估_混合版.xlsx">混合版评估</a>
<a href="行业投资价值综合评估_2026.xlsx">PE版评估</a>
<a href="行业投资价值综合评估_PB_2026.xlsx">PB版评估</a>
<a href="申万行业PE统计.xlsx">PE统计</a>
<a href="申万行业PB统计.xlsx">PB统计</a>
<a href="申万一级行业月度收益率矩阵.xlsx">月度收益率矩阵</a>
<a href="申万一级行业季度收益率矩阵.xlsx">季度收益率矩阵</a>
<a href="申万一级行业年度收益率矩阵.xlsx">年度收益率矩阵</a>
</div></div>

<div id="reports" class="chart-card">
<div class="chart-title"><span>分析报告文档</span><span class="badge">Word</span></div>
<div class="file-list">
<a href="申万一级行业季度收益率分析报告.docx">季度收益率分析报告</a>
<a href="申万一级行业年度收益率分析报告.docx">年度收益率分析报告</a>
<a href="申万一级行业综合投资价值与风险评估报告_混合版.docx">混合版风险评估</a>
<a href="申万一级行业综合投资价值与风险评估报告_PE_PB版_2026.docx">PE/PB版风险评估</a>
<a href="申万一级行业综合投资价值与风险评估报告_PB版_2026.docx">PB版风险评估</a>
</div></div>

<div id="heatmaps" class="chart-card">
<div class="chart-title"><span>收益率热力图</span><span class="badge">三频</span></div>
<div class="chart-body">
<div id="heatmap_monthly"><img src="申万一级行业月度收益率热力图.png" alt="月度" onclick="window.open(this.src)" style="width:100%"></div>
<div style="margin-top:12px;" id="heatmap_quarterly"><img src="申万一级行业季度收益率热力图.png" alt="季度" onclick="window.open(this.src)" style="width:100%"></div>
<div style="margin-top:12px;" id="heatmap_yearly"><img src="申万一级行业年度收益率热力图.png" alt="年度" onclick="window.open(this.src)" style="width:100%"></div>
</div></div>

<div id="data" class="chart-card">
<div class="chart-title"><span>数据说明</span><span class="badge">方法论</span></div>
<div class="file-list" style="font-size:13px;line-height:1.8;">
<p> 行业分类：申万一级行业（31个行业）</p>
<p> 数据频率：月度 / 季度 / 年度</p>
<p> 收益率计算：申万行业指数收盘价环比</p>
<p> 综合评估维度：PE估值分位 + PB估值分位 + 动量 + 风险等级</p>
<p> 热力图色阶：红色=上涨，绿色=下跌，颜色深度代表涨跌幅大小</p>
</div></div>

</div>
<footer>数据来源：Wind / 申万研究所 | 生成日期：2026-07-28 | 仅供参考</footer>
</body>
</html>'''

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(html)
print(f"OK -> {OUTPUT} ({len(html)} bytes)")
