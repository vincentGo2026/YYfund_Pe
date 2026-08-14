# -*- coding: utf-8 -*-
r"""一键运行 PB 统计分析三件套（按依赖顺序）

  1. pb_stats.py           → PB 基础统计 + 4 张图表
  2. enhance_pb_stats.py    → PB 百分位统计，追加到 Excel
  3. draw_candlestick_pb.py → PB 蜡烛图 3 张
  4. draw_barwithline_pb.py → PB 柱状图 + 百分位折线图 2 张

输出目录: D:\CC\Mid\估值\
"""
import subprocess, sys, os, time

PYTHON = r'D:\Users\dingd\anaconda3\python.exe'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

steps = [
    ('pb_stats.py',            'PB 基础统计 + 4张图'),
    ('enhance_pb_stats.py',    'PB 百分位统计扩充'),
    ('draw_candlestick_pb.py', 'PB 蜡烛图 (一级/二级)'),
    ('draw_barwithline_pb.py', 'PB 柱状图+百分位折线 (一级/二级)'),
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
        lines = r.stdout.strip().split('\n')[-5:]
        for line in lines:
            safe_line = line.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding)
            print(f'     {safe_line}')
        return True
    else:
        print(f'  FAIL (exit {r.returncode}, {elapsed:.1f}s)')
        for line in r.stderr.strip().split('\n')[-5:]:
            safe_line = line.encode(sys.stdout.encoding, errors='replace').decode(sys.stdout.encoding)
            print(f'     {safe_line}')
        return False

def main():
    print('PB 统计分析 — 一键运行')
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
