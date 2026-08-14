# -*- coding: utf-8 -*-
"""
局域网发布脚本 — 服务 sw_charts.html (申万行业分析报告)
- 根目录: D:/CC/Mid/估值  (使 sw_charts.html 的 ../ 引用可访问)
- 首页:   /CodeCC/sw_charts.html
- IP白名单: 局域网(192.168./10./172.16.) + 本机
用法:
  python 4publish_lan.py            # 默认端口 8888
  python 4publish_lan.py 9000       # 指定端口
  python 4publish_lan.py 8888 --public   # 允许所有IP(慎用)
"""
import os
import sys
from flask import Flask, request, abort, send_from_directory, redirect

BASE_DIR = r'D:\CC\Mid\估值'
INDEX_HTML = 'CodeCC/sw_charts.html'   # URL 路径必须用前向斜杠

app = Flask(__name__)

ALLOWED_IPS = ["127.0.0.1", "::1", "192.168.", "10.", "172.16."]


def is_allowed(addr):
    for rule in ALLOWED_IPS:
        if addr.startswith(rule):
            return True
    return False


@app.before_request
def check_ip():
    if not is_allowed(request.remote_addr):
        return f"<h3>拒绝访问</h3><p>IP {request.remote_addr} 不在白名单</p>", 403


@app.route('/')
def index():
    return send_from_directory(BASE_DIR, INDEX_HTML)


@app.route('/<path:p>')
def static_files(p):
    return send_from_directory(BASE_DIR, p)


if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8888
    public = '--public' in sys.argv
    if public:
        ALLOWED_IPS.clear()
        ALLOWED_IPS.append('')
        print(f'公网模式: 允许所有IP, http://localhost:{port}/')
    else:
        print(f'局域网发布: http://<本机IP>:{port}/')
        print(f'  首页: http://<本机IP>:{port}/  → 跳转 /CodeCC/sw_charts.html')
        print(f'  白名单: {len(ALLOWED_IPS)} 条规则 (局域网+本机)')
    app.run(host='0.0.0.0', port=port, debug=False)
