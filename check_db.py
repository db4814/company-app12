import sqlite3
import os

def check_projects_table():
    if not os.path.exists('enterprise.db'):
        print("数据库文件不存在")
        return
    
    conn = sqlite3.connect('enterprise.db')
    cursor = conn.cursor()
    
    # 查看projects表结构
    print("=== projects表结构 ===")
    cursor.execute("PRAGMA table_info(projects)")
    columns = cursor.fetchall()
    for col in columns:
        print(f"字段名: {col[1]}, 类型: {col[2]}, 允许空: {'是' if col[3] == 0 else '否'}")
    
    print("\n=== 表创建SQL ===")
    cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='projects'")
    create_sql = cursor.fetchone()
    if create_sql:
        print(create_sql[0])
    
    conn.close()

if __name__ == '__main__':
    check_projects_table()