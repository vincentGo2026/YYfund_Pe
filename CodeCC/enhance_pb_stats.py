# -*- coding: utf-8 -*-
"""
PB百分位统计扩充，写入 PB_Statistics_by_Sector.xlsx
"""
import pandas as pd
import numpy as np

CSV_PATH      = r"D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv"
OUTPUT_EXCEL  = r"D:\CC\Mid\估值\PB_Statistics_by_Sector.xlsx"

# 1. Read original CSV
raw = pd.read_csv(CSV_PATH)
print(f"原始数据: {len(raw)} rows")

# Normalize industry name: 申万半导体 -> 半导体(申万)
raw['name_norm'] = raw['secShortName'].apply(
    lambda n: (n[2:] if str(n).startswith('申万') else str(n)) + '(申万)'
    if not str(n).endswith('(申万)') else str(n)
)

# 2. Industry code from secID (e.g. 801081.ZICN -> 801081)
raw['industry_code'] = raw['secID'].str.split('.').str[0].astype(int)
code_map     = dict(raw[['name_norm', 'industry_code']].drop_duplicates('name_norm').values)
l1_map       = dict(raw[['name_norm', '申万一级行业']].drop_duplicates('name_norm').values)

# 3. Clean numeric data
df = raw.dropna(subset=['pb']).copy()
df['pb'] = pd.to_numeric(df['pb'], errors='coerce')
df = df.dropna(subset=['pb'])
df = df[df['pb'] > 0]

lo, hi = df['pb'].quantile(0.01), df['pb'].quantile(0.99)
df = df[(df['pb'] >= lo) & (df['pb'] <= hi)]

print(f"清洗后: {len(df)} rows")

# 4. Per-industry PB percentile statistics
def calc_pb_stats(grp):
    vals = grp['pb'].values
    latest_date = grp['tradeDate'].max()
    latest_val  = grp.loc[grp['tradeDate'] == latest_date, 'pb'].iloc[0]
    pct = (vals < latest_val).sum() / len(vals) * 100
    return pd.Series({
        '最新日期': latest_date,
        '最新PB': round(latest_val, 2),
        'PB历史百分位(%)': round(pct, 2),
        '历史最低PB': round(float(vals.min()), 2),
        '1/4分位PB': round(float(np.percentile(vals, 25)), 2),
        '中位数PB': round(float(np.median(vals)), 2),
        '3/4分位PB': round(float(np.percentile(vals, 75)), 2),
        '历史最高PB': round(float(vals.max()), 2),
        'PB样本天数': len(vals),
    })

pb_stat = df.groupby('name_norm').apply(calc_pb_stats, include_groups=False).reset_index()
pb_stat['行业代码'] = pb_stat['name_norm'].map(code_map).astype(int)
pb_stat['所属一级行业'] = pb_stat['name_norm'].map(l1_map)

# 5. Reorder columns
result = pb_stat[[
    '所属一级行业', 'name_norm', '行业代码', '最新日期',
    '最新PB', 'PB历史百分位(%)', '历史最低PB', '1/4分位PB', '中位数PB', '3/4分位PB', '历史最高PB',
    'PB样本天数',
]].copy()
result.columns = [
    '所属一级行业', '申万二级行业_API名', '行业代码', '最新日期',
    '最新PB', 'PB历史百分位(%)', '历史最低PB', '1/4分位PB', '中位数PB', '3/4分位PB', '历史最高PB',
    'PB样本天数',
]
result = result.sort_values(['所属一级行业', '中位数PB'], ascending=[True, False]).reset_index(drop=True)
result.index = result.index + 1
result.index.name = '序号'

print(f"\n输出: {len(result)} 行, 列数: {len(result.columns)}")
print("列名:", list(result.columns))
print(result.head(10).to_string())

# 6. Export
with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as w:
    result.to_excel(w, sheet_name='PB_PB统计', index=True)
    ws = w.sheets['PB_PB统计']
    for cc in ws.columns:
        cl = cc[0].column_letter
        ml = max((len(str(c.value)) if c.value else 0) for c in cc)
        ws.column_dimensions[cl].width = min(ml + 4, 22)
    pd.DataFrame({'说明': [
        '数据: datayes_all_SW_Industries_Level2_mapped.csv',
        '清洗: 剔除 PB<=0 及上下1%极端值',
        f'行业数: {len(result)}',
    ]}).to_excel(w, sheet_name='说明', index=False)
    w.sheets['说明'].column_dimensions['A'].width = 70

print(f"\n导出: {OUTPUT_EXCEL}")
