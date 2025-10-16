from http.server import BaseHTTPRequestHandler
import json

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        
        html_content = '''
        <html>
            <head><title>企业管理系统</title></head>
            <body>
                <h1>🚀 部署成功！</h1>
                <p>Flask 应用在 Vercel 上运行正常</p>
                <p>路径: {}</p>
                <a href="/login">测试登录页面</a>
            </body>
        </html>
        '''.format(self.path)
        
        self.wfile.write(html_content.encode())

def handler(request, context):
    return Handler()