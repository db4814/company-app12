from http.server import BaseHTTPRequestHandler
import json
import os

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(b'''
            <html>
                <body>
                    <h1>企业管理系统 - Vercel 测试</h1>
                    <p>Flask 应用正在运行！</p>
                    <a href="/login">登录</a>
                </body>
            </html>
        ''')

# Vercel 需要的函数入口
def handler(request):
    return Handler()