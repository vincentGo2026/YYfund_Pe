# -*- coding: utf-8 -*-
"""一键运行中信文件夹下所有分析脚本"""
import subprocess, sys, os, time

PYTHON = r'D:\Users\dingd\anaconda3\python.exe'
OUT_DIR = os.path.dirname(os.path.abspath(__file__))

scripts = [
    ('plot_quarterly_heatmap.py',     '季度收益率热力图'),
    ('plot_monthly_heatmap.py',           '月度收益率热力图'),
    ('plot_monthly_heatmap_pe_pb.py',    '月度收益率热力图（含PE/PB估值）'),
    ('plot_yearly_heatmap.py',            '年度收益率热力图'),
    ('export_quarterly_analysis.py',  '季度收益率分析报告'),
    ('export_yearly_analysis.py',     '年度收益率分析报告'),
    ('pe_pb_2026_analysis.py',        'PE/PB 综合投资价值评估'),
    ('hybrid_analysis.py',            '混合估值版评分报告'),
]

def run(script, desc):
    path = os.path.join(OUT_DIR, script)
    if not os.path.exists(path):
        print(f'  ! 文件不存在，跳过: {script}')
        return True
    print(f'\n{"="*60}')
    print(f'  [{desc}] 运行中...')
    t0 = time.time()
    r = subprocess.run([PYTHON, path], capture_output=True, text=True, encoding='utf-8', errors='replace')
    elapsed = time.time() - t0
    if r.returncode == 0:
        print(f'  OK 完成 ({elapsed:.1f}s)')
        return True
    else:
        print(f'  FAIL (exit {r.returncode}, {elapsed:.1f}s)')
        for line in r.stderr.strip().split('\n')[-3:]:
            print(f'     {line}')
        return False

def main():
    print('申万一级行业投资分析 -- 一键运行')
    print(f'目录: {OUT_DIR}')
    ok = fail = 0
    for script, desc in scripts:
        if run(script, desc):
            ok += 1
        else:
            fail += 1
    print(f'\n成功 {ok}/{len(scripts)}，失败 {fail}')
    return 0 if fail == 0 else 1

if __name__ == '__main__':
    sys.exit(main())
