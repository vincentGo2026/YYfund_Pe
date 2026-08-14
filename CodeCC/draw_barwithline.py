# -*- coding: utf-8 -*-
"""柱状图(最新PE)+折线图(百分位). 字号 FONT_SCALE 整体缩放"""
import pandas as pd, numpy as np, matplotlib, os
matplotlib.use('Agg')
import matplotlib.pyplot as plt, matplotlib.ticker as mticker
from matplotlib import font_manager

FONT_SCALE = float(os.environ.get('FONT_SCALE','1.3'))
for fp in [r'C:\Windows\Fonts\msyh.ttc',r'C:\Windows\Fonts\simhei.ttf']:
    if os.path.exists(fp):
        font_manager.fontManager.addfont(fp)
        plt.rcParams['font.family']=font_manager.FontProperties(fname=fp).get_name(); break
plt.rcParams['axes.unicode_minus']=False
plt.rcParams.update({'font.size':11*FONT_SCALE,'axes.labelsize':12*FONT_SCALE,'axes.titlesize':14*FONT_SCALE,'xtick.labelsize':9*FONT_SCALE,'ytick.labelsize':9*FONT_SCALE})
def fs(s): return s*FONT_SCALE

def draw_dual(data,name_col,title,filename,bar_col='最新PE',pct_col='PE历史百分位(%)',top_n=None):
    d=data.sort_values(bar_col,ascending=False).reset_index(drop=True)
    if top_n:d=d.head(top_n)
    n=len(d); fig_w=max(12,n*0.35*FONT_SCALE); fig_h=max(6,n*0.25*FONT_SCALE)
    fig,ax1=plt.subplots(figsize=(fig_w,fig_h))
    names=d[name_col].tolist(); vals=d[bar_col].values; pcts=d[pct_col].values; y=range(n)
    bh=0.7*FONT_SCALE; vfs=max(fs(7),9); nfs=max(fs(7),9)
    cols=plt.cm.viridis(vals/vals.max())
    ax1.barh(y,vals,color=cols,edgecolor='grey',lw=0.5,height=bh,zorder=3)
    ax1.set_xlabel('最新PE',fontsize=fs(11),color='#2c3e50')
    ax1.xaxis.set_major_formatter(mticker.FormatStrFormatter('%.1f'))
    ax1.tick_params(axis='x',colors='#2c3e50')
    for i,v in enumerate(vals):
        ax1.text(v+vals.max()*0.005,i,f'{v:.1f}',va='center',fontsize=vfs,color='#2c3e50')
    ax2=ax1.twiny()
    ax2.scatter(pcts,y,color='#e74c3c',s=30*FONT_SCALE,zorder=5,marker='o')
    pfs = max(fs(7), 8)
    for i, p in enumerate(pcts):
        ax2.text(p + 3, i, f'{p:.1f}% {names[i]}', va='center', fontsize=pfs, color='black')
    ax2.set_xlabel('PE历史百分位(%)',fontsize=fs(11),color='#e74c3c')
    ax2.xaxis.set_major_formatter(mticker.FormatStrFormatter('%.0f%%'))
    ax2.tick_params(axis='x',colors='#e74c3c'); ax2.set_xlim(-5,105)
    ax2.axvline(50,color='#e74c3c',ls=':',alpha=0.3,lw=0.8)
    ax1.set_yticks(list(y)); ax1.set_yticklabels(names,fontsize=nfs); ax1.invert_yaxis()
    ax1.grid(axis='x',alpha=0.25,ls='--')
    ax1.set_title(title,fontsize=fs(14),fontweight='bold')
    ax1.spines['top'].set_visible(False);ax1.spines['right'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    plt.tight_layout(); plt.savefig(filename,dpi=250,bbox_inches='tight'); plt.close()
    print(f'  {filename} ({n}行业)')

os.chdir(r'D:\CC\Mid\估值')
draw_dual(pd.read_excel('PE_Statistics_by_Sector.xlsx',sheet_name='PE_PB统计'),'申万二级行业_API名','申万二级行业 最新PE及历史百分位（2021-2026）','Chart_PE_L2_BarWithLine.png',top_n=40)
raw=pd.read_csv(r'D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv')
raw['pe']=pd.to_numeric(raw['pe'],errors='coerce');raw=raw.dropna(subset=['pe']);raw=raw[raw['pe']>0]
l,h=raw['pe'].quantile(0.01),raw['pe'].quantile(0.99);raw=raw[(raw['pe']>=l)&(raw['pe']<=h)]
def cl1(d):
    lt=d.groupby('申万一级行业').apply(lambda g:g.sort_values('tradeDate').iloc[-1]['pe'],include_groups=False).reset_index()
    lt.columns=['申万一级行业','最新PE']
    pc=d.groupby('申万一级行业').apply(lambda g:(g['pe']<g.sort_values('tradeDate').iloc[-1]['pe']).sum()/len(g)*100,include_groups=False).reset_index()
    pc.columns=['申万一级行业','PE历史百分位(%)']
    return lt.merge(pc,on='申万一级行业').sort_values('最新PE',ascending=False).reset_index(drop=True)
l1=cl1(raw);l1['PE历史百分位(%)']=l1['PE历史百分位(%)'].round(2)
draw_dual(l1,'申万一级行业','申万一级行业 最新PE及历史百分位（2021-2026）','Chart_PE_L1_BarWithLine.png')
print(f'完成！FONT_SCALE={FONT_SCALE:.1f}')
