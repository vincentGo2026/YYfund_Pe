# -*- coding: utf-8 -*-
"""申万一级行业月度收益率热力图 + PE/PB 混合估值复合图"""
import pandas as pd, numpy as np, matplotlib, os
matplotlib.use('Agg')
import matplotlib.pyplot as plt, seaborn as sns
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap
from sector_order import reindex_by_sector
import warnings; warnings.filterwarnings('ignore')

FONT_SCALE = float(os.environ.get('FONT_SCALE', '1.2'))
out_dir = r'D:\CC\Mid\估值\申万'
for fp in [r'C:\Windows\Fonts\msyh.ttc', r'C:\Windows\Fonts\simhei.ttf']:
    if os.path.exists(fp):
        font_manager.fontManager.addfont(fp)
        plt.rcParams['font.family'] = font_manager.FontProperties(fname=fp).get_name(); break
plt.rcParams['axes.unicode_minus'] = False

# ═══════════════════════════════════════════════════
# 1. 读取月度收益率矩阵
# ═══════════════════════════════════════════════════
dr = pd.read_excel(os.path.join(out_dir, '申万一级行业月度收益率矩阵.xlsx'), index_col=0)
dr = reindex_by_sector(dr, 'sw')
print(f"月度矩阵: {dr.shape[0]}行业 x {dr.shape[1]}月")

# ═══════════════════════════════════════════════════
# 2. 读取PE/PB原始数据，按混合规则合成
# ═══════════════════════════════════════════════════
PB_INDS = {'银行', '非银金融', '电力设备', '交通运输',
           '有色金属', '建筑装饰', '食品饮料', '家用电器', '商贸零售'}

def read_pe_pb(sheet, suffix):
    fp = os.path.join(out_dir, f'申万二级行业指数{suffix}_2026d.xlsx')
    raw = pd.read_excel(fp, sheet_name=sheet, header=None)
    names = [str(raw.iloc[1, j]).split('(')[0] for j in range(1, raw.shape[1])]
    dates = pd.to_datetime(raw.iloc[2:, 0])
    data = raw.iloc[2:, 1:].copy().reset_index(drop=True)
    data.columns = names
    for c in names:
        data[c] = pd.to_numeric(data[c], errors='coerce')
    data.index = dates
    return data[data.index.notna()]

pe_data = read_pe_pb('PEtxt', 'PE')
pb_data = read_pe_pb('PBtxt', 'PB')
print(f"PE数据: {pe_data.shape[0]}行, PB数据: {pb_data.shape[0]}行")

def calc_stats(data, label):
    latest = data.iloc[-1]
    rows = []
    for c in data.columns:
        vals = data[c].dropna()
        if len(vals) < 2: continue
        pct = (vals.sort_values() < latest[c]).mean() * 100
        rows.append({'行业': c,
            f'最新{label}': round(latest[c], 2),
            f'{label}历史百分位(%)': round(pct, 1),
            f'最高{label}': round(vals.max(), 2),
            f'最低{label}': round(vals.min(), 2),
            f'{label}中位数': round(vals.median(), 2)})
    return pd.DataFrame(rows).set_index('行业')

pe_stats = calc_stats(pe_data, 'PE')
pb_stats = calc_stats(pb_data, 'PB')

merged_stats = pe_stats.join(pb_stats, how='inner')
merged_stats['估值类型'] = merged_stats.index.map(lambda x: 'PB' if x in PB_INDS else 'PE')

def pick_val(row, col_pe, col_pb):
    return row[col_pb] if row['估值类型'] == 'PB' else row[col_pe]

merged_stats['最新值'] = merged_stats.apply(lambda r: pick_val(r, '最新PE', '最新PB'), axis=1)
merged_stats['历史百分位(%)'] = merged_stats.apply(lambda r: pick_val(r, 'PE历史百分位(%)', 'PB历史百分位(%)'), axis=1)
merged_stats['最高'] = merged_stats.apply(lambda r: pick_val(r, '最高PE', '最高PB'), axis=1)
merged_stats['最低'] = merged_stats.apply(lambda r: pick_val(r, '最低PE', '最低PB'), axis=1)
merged_stats['中位数'] = merged_stats.apply(lambda r: pick_val(r, 'PE中位数', 'PB中位数'), axis=1)

val_cols = ['估值类型', '最新值', '历史百分位(%)', '最高', '最低', '中位数']
print(f"统计完成: {len(merged_stats)}行业")
print(f"PB行业({sum(merged_stats['估值类型']=='PB')}个):", list(merged_stats.index[merged_stats['估值类型']=='PB']))

# ═══════════════════════════════════════════════════
# 3. 合并到月度矩阵并导出Excel
# ═══════════════════════════════════════════════════
dr_ex = dr.copy()
for col in val_cols:
    dr_ex[col] = merged_stats[col]

ex_path = os.path.join(out_dir, '申万一级行业月度收益率矩阵.xlsx')
dr_ex.to_excel(ex_path)
print(f"导出增强矩阵: {ex_path}")

# ═══════════════════════════════════════════════════
# 4. 绘图：热力图 + 混合估值表
# ═══════════════════════════════════════════════════
nc = len(dr.columns)
n_ind = len(dr)

hm_w = max(8, nc * 0.5 * FONT_SCALE)
info_w = 4.2
fw = hm_w + info_w
fh = max(8, n_ind * 0.38 * FONT_SCALE)

fig = plt.figure(figsize=(fw, fh))
gs = fig.add_gridspec(1, 2, width_ratios=[hm_w, info_w], wspace=0.06,
                       left=0.02, right=0.98, bottom=0.06, top=0.94)

# 左侧：月度收益率热力图
ax_hm = fig.add_subplot(gs[0, 0])
sns.heatmap(dr, annot=True, fmt=".1f", cmap='RdYlGn_r', linewidths=0.5, center=0,
            cbar_kws={'label': '月度收益率(%)', 'shrink': 0.6},
            annot_kws={'size': max(7, 9 * FONT_SCALE), 'weight': 'bold'}, ax=ax_hm)
ax_hm.set_title('申万一级行业月度收益率热力图 (2025.01-至今)',
                fontsize=max(11, 13 * FONT_SCALE), fontweight='bold', pad=12)
ax_hm.set_xlabel('月份', fontsize=10 * FONT_SCALE, labelpad=8)
ax_hm.set_ylabel('')

# 右侧：混合估值数据表
ax_info = fig.add_subplot(gs[0, 1])
ax_info.axis('off')

info_df = merged_stats.reindex(dr.index)
col_labels = ['行业', '类型', '最新值', '百分位\n(%)', '最高', '最低', '中位数']

def fmt_val(row):
    if row['估值类型'] == 'PB':
        return f"{row['最新值']:.2f}" if pd.notna(row['最新值']) else ''
    else:
        return f"{row['最新值']:.1f}" if pd.notna(row['最新值']) else ''

cell_text = []
for idx_name, row in info_df.iterrows():
    cell_text.append([
        idx_name,
        row['估值类型'],
        fmt_val(row),
        f"{row['历史百分位(%)']:.0f}" if pd.notna(row.get('历史百分位(%)')) else '',
        f"{row['最高']:.1f}" if pd.notna(row.get('最高')) else '',
        f"{row['最低']:.1f}" if pd.notna(row.get('最低')) else '',
        f"{row['中位数']:.1f}" if pd.notna(row.get('中位数')) else '',
    ])

nrows = len(cell_text)
ncols = len(col_labels)

pct_cmap = LinearSegmentedColormap.from_list('pct_cmap', ['#27ae60', '#f1c40f', '#e74c3c'])

table = ax_info.table(cellText=cell_text, colLabels=col_labels,
                      loc='center', cellLoc='center',
                      colWidths=[0.16, 0.09, 0.12, 0.12, 0.12, 0.12, 0.12])
table.auto_set_font_size(False)
table.set_fontsize(max(7, 8 * FONT_SCALE))
table.scale(1, 1.02)

for i in range(nrows):
    for j in range(ncols):
        cell = table[(i + 1, j)]
        cell.set_edgecolor('#cccccc'); cell.set_linewidth(0.5)
        if j == 3:  # 百分位列，着色
            try:
                v = float(cell_text[i][j])
                cell.set_facecolor(pct_cmap(v / 100))
                cell.get_text().set_fontweight('bold')
            except:
                cell.set_facecolor('#f5f5f5')
        elif j == 1:  # 估值类型列
            t = cell_text[i][j]
            cell.set_facecolor('#e8f0fe' if t == 'PE' else '#fce8e8')
            cell.get_text().set_fontweight('bold')
        elif j == 0:  # 行业名列
            cell.set_facecolor('#f0f0f0')
            cell.get_text().set_fontweight('bold')
        else:
            cell.set_facecolor('#fafafa')

for j in range(ncols):
    cell = table[(0, j)]
    cell.set_facecolor('#1B365D')
    cell.get_text().set_color('white')
    cell.get_text().set_fontweight('bold')
    cell.set_edgecolor('#1B365D')

ax_info.set_title('估值数据（混合）', fontsize=max(10, 12 * FONT_SCALE), fontweight='bold', pad=12, color='#1B365D')

cp = os.path.join(out_dir, '申万一级行业月度收益率_含PE_PB.png')
plt.savefig(cp, dpi=300, bbox_inches='tight'); plt.close()
print(f"合成图: {cp}\n完成！")
