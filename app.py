#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import webbrowser

# 设置项目根目录
ROOT_DIR = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, ROOT_DIR)

# 导入试用管理器
from utils.trial_manager import check_dwg_to_dxf_usage

def check_trial():
    """检查应用程序试用状态"""
    can_use, message, usage_count, remaining = check_dwg_to_dxf_usage()
    
    if not can_use:
        return False, message
    
    return True, f"已使用{usage_count}次，剩余{remaining}次"

from src.web import app

# 添加API路由获取试用状态
@app.route('/api/trial-status')
def trial_status():
    from flask import jsonify
    valid, message = check_trial()
    return jsonify({'valid': valid, 'message': message})

if __name__ == '__main__':
    # 检查试用期状态
    valid, message = check_trial()
    
    if not valid:
        # 创建并显示试用结束页面
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>CADtoExcel - 试用已结束</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }}
                .container {{ max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }}
                h1 {{ color: #d9534f; }}
                .contact {{ background-color: #f9f9f9; padding: 15px; border-radius: 5px; margin-top: 30px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>试用已结束</h1>
                <p>{message}</p>
                
                <div class="contact">
                    <h2>如何购买完整版</h2>
                    <p>请通过以下方式联系我们购买完整版：</p>
                    <ul>
                        <li>电话: 18689628301</li>
                        <li>邮箱: <a href="mailto:reef@doohuan.com">reef@doohuan.com</a></li>
                        <li>公司: 深圳市多焕智能科技有限公司</li>
                    </ul>
                </div>
            </div>
        </body>
        </html>
        """
        
        # 创建一个临时HTML文件
        if os.name == 'nt':  # Windows
            temp_dir = os.path.join(os.getenv('TEMP', os.path.expanduser('~/temp')), 'CADtoExcel')
        else:  # Linux/macOS
            temp_dir = os.path.join('/tmp', 'CADtoExcel')
        
        os.makedirs(temp_dir, exist_ok=True)
        html_path = os.path.join(temp_dir, 'trial_ended.html')
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        # 打开浏览器显示试用期结束页面
        webbrowser.open('file://' + html_path)
        
        print(message)
        input("按Enter键退出...")
        sys.exit(0)
    
    # 如果试用有效，显示剩余次数并继续启动应用
    print(f"CADtoExcel - {message}")
    
    # 在生产环境中应使用WSGI服务器（如gunicorn或uwsgi）
    # 这里仅为开发环境使用
    app.run(host='0.0.0.0', port=5001, debug=False)
