# -*- coding: utf-8 -*-
"""
公网发布脚本 — 服务 sw_charts_public.html (去除实盘主动估值) + 密码认证
- 根目录: D:/CC/Mid/估值
- 首页:   sw_charts_public.html (公网版，已去除实盘主动估值区块)
- 密码:   HTTP Basic Auth
用法:
  python 5publish_public.py              # 默认端口 8890
  python 5publish_public.py 9000         # 指定端口
配合 cloudflared: cloudflared tunnel --url http://127.0.0.1:8890
"""
import os
import sys
from functools import wraps
from flask import Flask, request, Response, send_from_directory

BASE_DIR = r'D:\CC\Mid\估值'
INDEX_HTML = 'CodeCC/sw_charts_public.html'

# ====== 访问密码 (从环境变量读取, 避免密钥泄露到仓库) ======
AUTH_USER = os.environ.get("AUTH_USER", "sw")
AUTH_PASS = os.environ.get("AUTH_PASS", "CHANGE_ME")


def check_auth(username, password):
    return username == AUTH_USER and password == AUTH_PASS


def authenticate():
    # realm 必须用 ASCII (HTTP头 latin-1 编码, 中文会报错)
    return Response('Unauthorized', status=401,
                    headers={'WWW-Authenticate': 'Basic realm="SW Industry Analysis"'})


def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return authenticate()
        return f(*args, **kwargs)
    return decorated


app = Flask(__name__)


@app.route('/')
@requires_auth
def index():
    return send_from_directory(BASE_DIR, INDEX_HTML)


@app.route('/<path:p>')
@requires_auth
def static_files(p):
    return send_from_directory(BASE_DIR, p)


if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8890
    print(f'公网版服务器(密码认证): http://127.0.0.1:{port}/')
    print(f'  用户名: {AUTH_USER}  密码: {AUTH_PASS}  (请在脚本中修改)')
    print(f'  首页: {INDEX_HTML}')
    print(f'  cloudflared: cloudflared tunnel --url http://127.0.0.1:{port}')
    app.run(host='0.0.0.0', port=port, debug=False)
