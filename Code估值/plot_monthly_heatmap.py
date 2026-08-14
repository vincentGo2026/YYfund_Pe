# -*- coding: utf-8 -*-
"""PriceSW1 月度收益率热力图（2025-01-01至今）"""
import pandas as pd, numpy as np, matplotlib, os
matplotlib.use('Agg')
import matplotlib.pyplot as plt, seaborn as sns
from matplotlib import font_manager
from sector_order import reindex_by_sector

FONT_SCALE = float(os.environ.get('FONT_SCALE','1.2'))
data_dir = r'D:\CC\DB\MKT'
out_dir = r'D:\CC\Mid\估值\申万'
for fp in [r'C:\Windows\Fonts\msyh.ttc',r'C:\Windows\Fonts\simhei.ttf']:
    if os.path.exists(fp):
        font_manager.fontManager.addfont(fp)
        plt.rcParams['font.family']=font_manager.FontProperties(fname=fp).get_name(); break
plt.rcParams['axes.unicode_minus']=False

df = pd.read_excel(os.path.join(data_dir,'PriceSW1.xlsx'),header=None)
h=list(df.iloc[1].values);h[0]='Date'
d=df.iloc[2:].copy();d.columns=h
d['Date']=pd.to_datetime(d['Date']);d=d.dropna(subset=['Date']).sort_values('Date').reset_index(drop=True)
ind=[c for c in d.columns if c!='Date']
for i in ind:d[i]=pd.to_numeric(d[i],errors='coerce')
print(f"数据: {len(d)}行 {len(ind)}行业 {d['Date'].min()}~{d['Date'].max()}")

d=d[d['Date']>='2025-01-01'].reset_index(drop=True)
d['YM']=d['Date'].dt.to_period('M')
rows=[]
for ym,g in d.groupby('YM'):
    g=g.sort_values('Date');f=g.iloc[0];l=g.iloc[-1]
    rv={ind:(l[ind]/f[ind]-1)*100 for ind in ind if pd.notna(f[ind]) and f[ind]>0 and pd.notna(l[ind])}
    rv['Month']=str(ym);rows.append(rv)
dr=pd.DataFrame(rows).set_index('Month').T
dr.index=dr.index.map(lambda x: str(x).split('\n')[0].replace('(申万)','').strip())
dr=reindex_by_sector(dr, 'sw')
print(f"矩阵: {dr.shape[0]}行业 x {dr.shape[1]}月")
print("月份:",list(dr.columns))
print(f"范围: {dr.min().min():.1f}%~{dr.max().max():.1f}%")

ex=os.path.join(out_dir,'申万一级行业月度收益率矩阵.xlsx')
dr.to_excel(ex);print(f"导出: {ex}")

nc=len(dr.columns);fw=max(10,nc*0.65*FONT_SCALE);fh=max(8,dr.shape[0]*0.38*FONT_SCALE)
plt.figure(figsize=(fw,fh))
sns.heatmap(dr,annot=True,fmt=".1f",cmap='RdYlGn_r',linewidths=0.5,center=0,
    cbar_kws={'label':'月度收益率(%)'},annot_kws={'size':max(8,10*FONT_SCALE),'weight':'bold'})
plt.title('申万一级行业月度收益率热力图 (2025.01-至今)',fontsize=max(12,14*FONT_SCALE),fontweight='bold',pad=15)
plt.xlabel('月份',fontsize=11*FONT_SCALE,labelpad=10)
plt.ylabel('申万一级行业',fontsize=11*FONT_SCALE,labelpad=10)
plt.xticks(rotation=45 if nc>6 else 0,fontsize=max(8,10*FONT_SCALE),fontweight='bold')
plt.yticks(fontsize=max(8,10*FONT_SCALE),fontweight='bold',rotation=0)
plt.tight_layout()
cp=os.path.join(out_dir,'申万一级行业月度收益率热力图.png')
plt.savefig(cp,dpi=300,bbox_inches='tight');plt.close()
print(f"热力图: {cp}\n完成！")
