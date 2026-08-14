# -*- coding: utf-8 -*-
"""一键运行 DataYes 数据拉取 + SW1 映射

  1. 2fetch_datayes_data_SWL1.py  (并行) → 申万一级行业数据
  2. 2fetch_datayes_data_SWL2.py  (并行) → 申万二级行业数据
  3. 2add_sw1_mapping.py           (等待L2完成) → 补充申万一级行业列
"""
import subprocess, sys, os, time, threading

PYTHON = r'D:\Users\dingd\anaconda3\python.exe'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def run(script, desc):
    """同步执行一个脚本，返回 (success, elapsed)"""
    path = os.path.join(BASE_DIR, script)
    if not os.path.exists(path):
        print(f'  ! 文件不存在，跳过: {script}')
        return True, 0
    t0 = time.time()
    r = subprocess.run(
        [PYTHON, path],
        cwd=BASE_DIR,
        capture_output=True, text=True, encoding='utf-8', errors='replace',
    )
    elapsed = time.time() - t0
    if r.returncode == 0:
        print(f'  OK ({elapsed:.1f}s)')
        *_, tail = (r.stdout.strip().split('\n')[-3:])
        for line in tail:
            print(f'     {line}')
        return True, elapsed
    else:
        print(f'  FAIL (exit {r.returncode}, {elapsed:.1f}s)')
        for line in r.stderr.strip().split('\n')[-3:]:
            print(f'     {line}')
        return False, elapsed


def run_async(script, desc, results):
    """异步执行（放入线程），结果写入 results dict"""
    ok, elapsed = run(script, desc)
    results[script] = (ok, elapsed)


def main():
    print('=' * 60)
    print('DataYes 数据拉取 + SW1 映射 — 一键运行')
    print(f'目录: {BASE_DIR}')
    print('=' * 60)

    # ---------- 第1步：并行拉取 L1 和 L2 ----------
    print('\n[1/2] 并行拉取 申万一级 + 申万二级数据 ...\n')

    results = {}
    t1 = threading.Thread(target=run_async, args=(
        '2fetch_datayes_data_SWL1.py', '申万一级行业', results))
    t2 = threading.Thread(target=run_async, args=(
        '2fetch_datayes_data_SWL2.py', '申万二级行业', results))
    t1.start()
    t2.start()
    t1.join()
    t2.join()

    l1_ok, l1_time = results.get('2fetch_datayes_data_SWL1.py', (True, 0))
    l2_ok, l2_time = results.get('2fetch_datayes_data_SWL2.py', (True, 0))

    l1_status = 'OK' if l1_ok else 'FAIL'
    l2_status = 'OK' if l2_ok else 'FAIL'
    print(f'\n  一级行业: {l1_status} ({l1_time:.1f}s) | 二级行业: {l2_status} ({l2_time:.1f}s)')

    if not l2_ok:
        print('\n  ! 申万二级数据拉取失败，跳过后续映射步骤')
        return 1

    # ---------- 第2步：SW1 映射 ----------
    print(f'\n[2/2] 补充 申万一级行业 映射 ...\n')
    m_ok, m_time = run('2add_sw1_mapping.py', 'SW1 映射')

    # ---------- 汇总 ----------
    ok = sum([l1_ok, l2_ok, m_ok])
    fail = 3 - ok
    print(f'\n{"=" * 60}')
    print(f'完成: {ok}/3 成功, {fail} 失败')
    return 0 if fail == 0 else 1


if __name__ == '__main__':
    sys.exit(main())
