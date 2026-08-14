# -*- coding: utf-8 -*-
"""
PB 基础统计 + 4 张图表
输出: D:\CC\Mid\估值\PB_Statistics_by_Sector.xlsx
"""
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib import font_manager
import os

# ============================================================
# 0. 字体设置（支持中文）
# ============================================================
font_paths = [
    r'C:\Windows\Fonts\msyh.ttc',
    r'C:\Windows\Fonts\simhei.ttf',
    r'C:\Windows\Fonts\simsun.ttc',
    r'C:\Windows\Fonts\SIMKAI.TTF',
]
for fp in font_paths:
    if os.path.exists(fp):
        font_manager.fontManager.addfont(fp)
        font_prop = font_manager.FontProperties(fname=fp)
        plt.rcParams['font.family'] = font_prop.get_name()
        break

plt.rcParams['axes.unicode_minus'] = False

# ============================================================
# 1. 读取数据
# ============================================================
csv_path = r"D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv"
df = pd.read_csv(csv_path)
print(f"读取完成：{len(df)} 行, {len(df.columns)} 列")

# ============================================================
# 2. 清理PB数据
# ============================================================
df = df.dropna(subset=['pb'])
df['pb'] = pd.to_numeric(df['pb'], errors='coerce')
df = df.dropna(subset=['pb'])

before = len(df)
df = df[df['pb'] > 0]
print(f"删除 {before - len(df)} 行 pb<=0 的数据")

# 去掉上下1%极端值
lower = df['pb'].quantile(0.01)
upper = df['pb'].quantile(0.99)
df = df[(df['pb'] >= lower) & (df['pb'] <= upper)]
print(f"清洗后剩余 {len(df)} 行")

# ============================================================
# 3. 按 secShortName 分组统计
# ============================================================
stats = df.groupby('secShortName')['pb'].agg([
    ('最大值', 'max'),
    ('最小值', 'min'),
    ('中位数', 'median'),
    ('1/4分位数', lambda x: x.quantile(0.25)),
    ('3/4分位数', lambda x: x.quantile(0.75)),
    ('平均数', 'mean'),
]).reset_index()

# 所有数值保留2位小数
for col in ['最大值', '最小值', '中位数', '1/4分位数', '3/4分位数', '平均数']:
    stats[col] = stats[col].round(2)

stats = stats.sort_values('中位数', ascending=False).reset_index(drop=True)
stats.index = stats.index + 1
stats.index.name = '序号'
print(f"统计完成：共 {len(stats)} 个行业")

# 打印前5行验证
print("\n前5行预览：")
print(stats.head().to_string())

# ============================================================
# 4. 导出 Excel
# ============================================================
excel_path = r"D:\CC\Mid\估值\PB_Statistics_by_Sector.xlsx"
with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    stats.to_excel(writer, sheet_name='PB统计', index=True)
    ws = writer.sheets['PB统计']
    for col_cells in ws.columns:
        col_letter = col_cells[0].column_letter
        max_len = max((len(str(c.value)) if c.value else 0) for c in col_cells)
        ws.column_dimensions[col_letter].width = min(max_len + 4, 22)

    info_df = pd.DataFrame({'说明': [
        '数据文件: datayes_all_SW_Industries_Level2_mapped.csv',
        '统计区间: 2021-05-01 至 2026-05-21',
        f'行业数量: {len(stats)}',
        '清洗规则: 剔除PB<=0及上下1%极端值',
        '全部数值保留2位小数',
    ]})
    info_df.to_excel(writer, sheet_name='说明', index=False)
    writer.sheets['说明'].column_dimensions['A'].width = 70

print(f"Excel导出: {excel_path}")

# ============================================================
# 5. 画图
# ============================================================
top_n = min(30, len(stats))
plot_data = stats.head(top_n).copy()

# 图1: PB中位数 + Q1/Q3区间条形图
fig, ax = plt.subplots(figsize=(16, 10))
y_pos = range(len(plot_data))
medians = plot_data['中位数'].values
q1 = plot_data['1/4分位数'].values
q3 = plot_data['3/4分位数'].values

norm = (medians - medians.min()) / (medians.max() - medians.min() + 1e-10)
colors = plt.cm.viridis(norm)

ax.barh(y_pos, medians, color=colors, edgecolor='grey', linewidth=0.5)
ax.errorbar(medians, y_pos, xerr=[medians - q1, q3 - medians], fmt='none',
            ecolor='grey', capsize=3, capthick=1, elinewidth=1)
ax.set_yticks(y_pos)
ax.set_yticklabels(plot_data['secShortName'], fontsize=9)
ax.invert_yaxis()
ax.set_xlabel('PB（市净率）', fontsize=12)
ax.set_title('申万二级行业 PB 中位数及四分位区间（2021-2026）', fontsize=14, fontweight='bold')
ax.xaxis.set_major_formatter(mticker.FormatStrFormatter('%.2f'))

for i, v in enumerate(medians):
    ax.text(v + 0.03, i, f'{v:.2f}', va='center', fontsize=7.5, color='darkred')

plt.tight_layout()
p1 = r"D:\CC\Mid\估值\Chart_PB_Median_Quartiles.png"
plt.savefig(p1, dpi=200, bbox_inches='tight')
plt.close()
print(f"图1: {p1}")

# 图2: PB平均值条形图
fig, ax = plt.subplots(figsize=(16, 10))
means = plot_data['平均数'].values
norm2 = (means - means.min()) / (means.max() - means.min() + 1e-10)
colors2 = plt.cm.plasma(norm2)
ax.barh(y_pos, means, color=colors2, edgecolor='grey', linewidth=0.5)
ax.set_yticks(y_pos)
ax.set_yticklabels(plot_data['secShortName'], fontsize=9)
ax.invert_yaxis()
ax.set_xlabel('PB（市净率）', fontsize=12)
ax.set_title('申万二级行业 PB 平均值（2021-2026）', fontsize=14, fontweight='bold')
ax.xaxis.set_major_formatter(mticker.FormatStrFormatter('%.2f'))
for i, v in enumerate(means):
    ax.text(v + 0.03, i, f'{v:.2f}', va='center', fontsize=7.5, color='darkred')
plt.tight_layout()
p2 = r"D:\CC\Mid\估值\Chart_PB_Mean_Bar.png"
plt.savefig(p2, dpi=200, bbox_inches='tight')
plt.close()
print(f"图2: {p2}")

# 图3: 多项指标分组柱状图（Top 15）
top15 = plot_data.head(15).copy()
fig, ax = plt.subplots(figsize=(18, 8))
x = range(len(top15))
width = 0.13
metrics = ['最小值', '1/4分位数', '中位数', '平均数', '3/4分位数', '最大值']
clrs = ['#2ecc71', '#3498db', '#f39c12', '#e74c3c', '#9b59b6', '#2c3e50']
for i, (m, c) in enumerate(zip(metrics, clrs)):
    vals = top15[m].values
    off = (i - 2.5) * width
    ax.bar([xi + off for xi in x], vals, width, label=m, color=c, alpha=0.85)
ax.set_xticks(x)
ax.set_xticklabels(top15['secShortName'], fontsize=8, rotation=45, ha='right')
ax.set_ylabel('PB', fontsize=12)
ax.set_title('申万二级行业 PB 多项指标对比（Top 15）', fontsize=14, fontweight='bold')
ax.legend(fontsize=9)
ax.yaxis.set_major_formatter(mticker.FormatStrFormatter('%.2f'))
plt.tight_layout()
p3 = r"D:\CC\Mid\估值\Chart_PB_Multi_Metrics.png"
plt.savefig(p3, dpi=200, bbox_inches='tight')
plt.close()
print(f"图3: {p3}")

# 图4: PB中位数分布直方图
fig, ax = plt.subplots(figsize=(12, 6))
all_m = stats['中位数'].values
ax.hist(all_m, bins=30, color='#3498db', edgecolor='white', alpha=0.8)
ax.axvline(all_m.mean(), color='red', ls='--', lw=2, label=f'均值={all_m.mean():.2f}')
ax.axvline(np.median(all_m), color='green', ls='--', lw=2, label=f'中位数={np.median(all_m):.2f}')
ax.set_xlabel('PB 中位数', fontsize=12)
ax.set_ylabel('行业数', fontsize=12)
ax.set_title('申万二级行业 PB 中位数分布', fontsize=14, fontweight='bold')
ax.legend(fontsize=10)
plt.tight_layout()
p4 = r"D:\CC\Mid\估值\Chart_PB_Median_Distribution.png"
plt.savefig(p4, dpi=200, bbox_inches='tight')
plt.close()
print(f"图4: {p4}")

print("\n=== 全部完成 ===")
print(f"Excel: {excel_path}")
print(f"图1(中位数+区间): {p1}")
print(f"图2(平均值): {p2}")
print(f"图3(多指标对比): {p3}")
print(f"图4(分布直方图): {p4}")
