# -*- coding: utf-8 -*-
"""
按照 Chart_PE_Candlestick 样式绘制 PE 蜡烛图（箱线图风格）
输出3张图：
  Chart_PE_Candlestick_L1.png          (一级行业，无数值)
  Chart_PE_Candlestick_L1_with_values.png (一级行业，含数值)
  Chart_PE_Candlestick.png             (二级行业，含数值)
"""

import pandas as pd, numpy as np, matplotlib, os
matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.ticker as mticker
from matplotlib import font_manager

# 字体
for fp in [r'C:\Windows\Fonts\msyh.ttc', r'C:\Windows\Fonts\simhei.ttf']:
    if os.path.exists(fp):
        font_manager.fontManager.addfont(fp)
        plt.rcParams['font.family'] = font_manager.FontProperties(fname=fp).get_name()
        break
plt.rcParams['axes.unicode_minus'] = False

def draw_chart(data, name_col, title, filename, show_vals, latest_pe=None):
    """水平蜡烛图：min |== Q1 |中位数| Q3 ==| max
    可选 latest_pe: 与 data 对应的最新 PE 序列，以红色菱形标注在图中"""
    d = data.sort_values('中位数PE', ascending=False).reset_index(drop=True)
    n = len(d)
    fig, ax = plt.subplots(figsize=(18, max(6, n*0.28)))
    names = d[name_col].tolist()
    lo, q1, med, q3, hi = [d[c].values for c in
        ['历史最低PE','1/4分位PE','中位数PE','3/4分位PE','历史最高PE']]
    y = range(n)

    for i in range(n):
        ax.plot([lo[i],q1[i]],[i,i], color='#555', lw=1)
        ax.plot([q3[i],hi[i]],[i,i], color='#555', lw=1)
    colors = plt.cm.RdYlGn_r(1-(med-med.min())/(med.max()-med.min()+1e-10))
    for i in range(n):
        ax.barh(i, q3[i]-q1[i], left=q1[i], height=0.6,
                color=colors[i], edgecolor='#333', lw=0.8)
    ax.scatter(med, y, marker='|', s=200, color='black', zorder=4, lw=2)

    if show_vals:
        pad = (hi.max()-lo.min())*0.015
        for i in range(n):
            ax.text(lo[i]-pad, i, f'{lo[i]:.1f}', ha='right', va='center', fontsize=8, color='#666')
            ax.text(med[i], i+0.35, f'{med[i]:.1f}', ha='center', va='bottom', fontsize=8, fontweight='bold')
            ax.text(hi[i]+pad, i, f'{hi[i]:.1f}', ha='left', va='center', fontsize=8, color='#666')

    # 红色菱形标注最新 PE
    if latest_pe is not None:
        # 用原始 data 的 name_col 和 latest_pe 建映射，再按排序后的 names 对齐
        lp_map = dict(zip(data[name_col], latest_pe))
        lp_vals = [lp_map.get(n, np.nan) for n in names]
        lp_arr = np.array(lp_vals, dtype=float)
        valid = ~np.isnan(lp_arr)
        ax.scatter(lp_arr[valid], np.array(y)[valid], marker='D', s=40,
                   color='#e63946', edgecolor='#fff', lw=0.8, zorder=5, label='最新PE')
        # 标注数值
        pad_lp = (hi.max()-lo.min())*0.02
        for i in range(n):
            if not np.isnan(lp_arr[i]):
                val = lp_arr[i]
                ax.text(val+pad_lp, i, f'{val:.1f}', ha='left', va='center',
                        fontsize=8, color='#e63946', fontweight='bold')
        ax.legend(loc='lower right', fontsize=10)

    ax.set_yticks(list(y))
    ax.set_yticklabels(names, fontsize=9)
    ax.invert_yaxis()
    ax.set_xlabel('PE', fontsize=14)
    ax.set_title(title, fontsize=16, fontweight='bold')
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f'))
    ax.grid(axis='x', alpha=0.3, ls='--')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(filename, dpi=250, bbox_inches='tight')
    plt.close()
    print(f"  {filename} ({n}行业)")

# L1 stats from raw CSV
raw = pd.read_csv(r'D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv')
raw['pe'] = pd.to_numeric(raw['pe'], errors='coerce')
raw = raw.dropna(subset=['pe'])
raw = raw[raw['pe'] > 0]
lo, hi = raw['pe'].quantile(0.01), raw['pe'].quantile(0.99)
raw = raw[(raw['pe']>=lo) & (raw['pe']<=hi)]

l1 = raw.groupby('申万一级行业')['pe'].agg([
    ('历史最低PE','min'),('1/4分位PE',lambda x:x.quantile(0.25)),
    ('中位数PE','median'),('3/4分位PE',lambda x:x.quantile(0.75)),('历史最高PE','max')
]).round(2).reset_index()

# 各一级行业最新日期 PE（B列 secShortName 为二级行业，按一级行业取最新日期 pe 中位数）
raw['tradeDate'] = pd.to_datetime(raw['tradeDate'])
latest_date = raw['tradeDate'].max()
latest_raw = raw[raw['tradeDate'] == latest_date]
l1_latest = latest_raw.groupby('申万一级行业')['pe'].median().round(2).reset_index()
l1_latest_map = dict(zip(l1_latest['申万一级行业'], l1_latest['pe']))

os.chdir(r'D:\CC\Mid\估值')
draw_chart(l1, '申万一级行业',
    '申万一级行业 PE 蜡烛图（2021-2026）','Chart_PE_Candlestick_L1.png', False)
# L1 带最新PE标记
draw_chart(l1, '申万一级行业',
    '申万一级行业 PE 蜡烛图（2021-2026）','Chart_PE_Candlestick_L1_with_values.png', True,
    latest_pe=[l1_latest_map.get(n, np.nan) for n in l1['申万一级行业']])

# L2 from enhanced Excel
l2 = pd.read_excel('PE_Statistics_by_Sector.xlsx', sheet_name='PE_PB统计')
draw_chart(l2, '申万二级行业_API名',
    '申万二级行业 PE 蜡烛图（2021-2026）','Chart_PE_Candlestick.png', True,
    latest_pe=l2['最新PE'].values)
print("完成！")
