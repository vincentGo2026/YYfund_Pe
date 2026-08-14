"""PE图表发布服务 — Flask + IP白名单"""
import os, sys
from flask import Flask, request, abort, send_from_directory

app = Flask(__name__)

# ====== IP白名单 ======
ALLOWED_IPS = [
    "127.0.0.1", "::1",
    "192.168.", "10.", "172.16.",
]

def is_allowed(addr):
    for rule in ALLOWED_IPS:
        if addr.startswith(rule): return True
    return False

@app.before_request
def check_ip():
    if not is_allowed(request.remote_addr):
        return f"<h3>拒绝访问</h3><p>IP {request.remote_addr} 不在白名单</p>", 403

@app.route('/')
def index():
    return send_from_directory(os.path.dirname(__file__), 'pe_charts.html')

@app.route('/<path:p>')
def static_files(p):
    return send_from_directory(os.path.dirname(__file__), p)

if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5000
    public = "--public" in sys.argv
    if not public:
        print(f"服务: http://localhost:{port}/")
        print(f"白名单: {len(ALLOWED_IPS)}条规则")
        print("外网需配置隧道(ngrok/cloudflared) — 加 --public 允许所有IP")
    else:
        print(f"公网模式: http://localhost:{port}/ (允许所有IP)")
        # 清空白名单允许所有
        ALLOWED_IPS.clear()
        ALLOWED_IPS.append("")
    app.run(host='0.0.0.0', port=port, debug=False)
