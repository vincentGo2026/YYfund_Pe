# -*- coding: utf-8 -*-
"""
Render 部署服务器 — 服务申万行业估值分析公网版页面 (sw_charts_public.html)

- 根目录: 仓库根目录 (对应本地 D:\CC\Mid\估值)
- 首页:   /CodeCC/sw_charts_public.html
- 端口:   Render 注入的 $PORT (本地默认 8890)
- 认证:   HTTP Basic Auth。凭据来自环境变量 AUTH_USER / AUTH_PASS,
          由 Render 控制台手动设置 (render.yaml 中 sync: false), 不写入仓库。

本地运行 (Windows PowerShell):
  $env:AUTH_USER='sw'; $env:AUTH_PASS='你的密码'; python app.py
  浏览器打开 http://localhost:8890/ 并输入凭据
"""
import os
from functools import wraps

from flask import Flask, Response, redirect, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = 'CodeCC/sw_charts_public.html'

# Basic Auth 凭据 (两者缺一不可; render.yaml 中 sync: false 由控制台填入)
AUTH_USER = os.environ.get('AUTH_USER')
AUTH_PASS = os.environ.get('AUTH_PASS')

app = Flask(__name__)


def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or auth.username != AUTH_USER or auth.password != AUTH_PASS:
            return Response('Unauthorized', status=401,
                            headers={'WWW-Authenticate': 'Basic realm="YYfund_Pe"'})
        return f(*args, **kwargs)
    return decorated


@app.route('/healthz')
def healthz():
    """Render 健康检查 (不要求认证)"""
    return 'ok', 200


@app.route('/')
@requires_auth
def index():
    # 页面相对路径按 /CodeCC/ 基准编写 (../ 指向仓库根, 同级指向 CodeCC/)
    # 必须重定向到该路径, 否则同级引用会解析到根目录而 404
    return redirect('/' + INDEX_HTML)


@app.route('/<path:p>')
@requires_auth
def static_files(p):
    return send_from_directory(BASE_DIR, p)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '8890'))
    if not AUTH_USER or not AUTH_PASS:
        print('[警告] AUTH_USER / AUTH_PASS 未设置, 页面将无法通过认证', flush=True)
    print(f'服务器: http://localhost:{port}/  [用户: {AUTH_USER or "?"}]', flush=True)
    app.run(host='0.0.0.0', port=port, debug=False)
