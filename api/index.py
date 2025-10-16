from flask import Flask
app = Flask(__name__)

# 这里复制您 app.py 中的所有路由和函数
@app.route('/')
def index():
    return '企业管理系统首页'

@app.route('/login')
def login():
    return '登录页面'

# 复制您 app.py 中所有的路由...

if __name__ == '__main__':
    app.run()