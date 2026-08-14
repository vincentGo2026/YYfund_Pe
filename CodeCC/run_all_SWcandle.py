# -*- coding: utf-8 -*-
"""一键运行 PE/PB 统计分析三件套（按依赖顺序）

  1. pe_stats.py           → PE 基础统计 + 4 张图表
  2. enhance_pe_stats.py   → PE/PB 百分位统计，追加到 Excel
  3. draw_candlestick.py   → PE 蜡烛图 3 张
  4. draw_barwithline.py   → PE 柱状图 + 百分位折线图 2 张

输出目录: D:\CC\Mid\估值\
"""
import subprocess, sys, os, time

PYTHON = r'D:\Users\dingd\anaconda3\python.exe'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

steps = [
    ('pe_stats.py',            'PE 基础统计 + 4张图'),
    ('enhance_pe_stats.py',    'PE/PB 百分位统计扩充'),
    ('draw_candlestick.py',    'PE 蜡烛图 (一级/二级)'),
    ('..\\draw_barwithline.py', 'PE 柱状图+百分位折线 (一级/二级)'),
]

def run(script, desc):
    path = os.path.join(BASE_DIR, script)
    if not os.path.exists(path):
        print(f'  ! 文件不存在，跳过: {script}')
        return True
    print(f'\n{"=" * 60}')
    print(f'  [{desc}] 运行中 ...')
    t0 = time.time()
    r = subprocess.run(
        [PYTHON, path],
        cwd=BASE_DIR,
        capture_output=True, text=True, encoding='utf-8', errors='replace',
    )
    elapsed = time.time() - t0
    if r.returncode == 0:
        print(f'  OK 完成 ({elapsed:.1f}s)')
        # 打印脚本最后几行输出
        *_, last = (r.stdout.strip().split('\n')[-5:])
        for line in last:
            print(f'     {line}')
        return True
    else:
        print(f'  FAIL (exit {r.returncode}, {elapsed:.1f}s)')
        for line in r.stderr.strip().split('\n')[-5:]:
            print(f'     {line}')
        return False

def main():
    print('PE/PB 统计分析 — 一键运行')
    print(f'目录: {BASE_DIR}')
    print(f'Python: {PYTHON}')

    ok = fail = 0
    for script, desc in steps:
        if run(script, desc):
            ok += 1
        else:
            fail += 1
            print(f'\n  ! 因 {script} 失败，后续步骤可能受影响')

    print(f'\n{"=" * 60}')
    print(f'完成: {ok}/{len(steps)} 成功, {fail} 失败')
    return 0 if fail == 0 else 1

if __name__ == '__main__':
    sys.exit(main())
