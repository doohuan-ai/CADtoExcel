#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import datetime
import requests
import webbrowser

# 设置项目根目录
ROOT_DIR = os.path.abspath(os.path.dirname(__file__))
sys.path.insert(0, ROOT_DIR)

def get_network_time():
    """从时间服务器获取网络时间，防止用户修改本地时间"""
    # 使用国内可访问的时间服务器
    time_servers = [
        "https://api.m.taobao.com/rest/api3.do?api=mtop.common.getTimestamp",  # 淘宝API
        "http://quan.suning.com/getSysTime.do",  # 苏宁API
        "https://sapi.k780.com/?app=life.time&appkey=10003&sign=b59bc3ef6191eb9f747dd4e83c99f2a4&format=json"  # K780
    ]
    
    for server in time_servers:
        try:
            response = requests.get(server, timeout=10)  # 增加超时时间
            if response.status_code == 200:
                data = response.json()
                # 根据不同API解析时间
                if "taobao.com" in server:
                    timestamp = int(data["data"]["t"]) / 1000
                    return datetime.datetime.fromtimestamp(timestamp)
                elif "suning.com" in server:
                    time_str = data["sysTime2"]
                    return datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                elif "k780.com" in server:
                    time_str = data["result"]["datetime_2"]
                    return datetime.datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except Exception as e:
            print(f"时间服务器 {server} 请求失败: {e}")
    
    # 如果所有服务器都失败，再尝试国际服务器
    try:
        response = requests.get("http://worldtimeapi.org/api/ip", timeout=5)
        if response.status_code == 200:
            data = response.json()
            return datetime.datetime.fromisoformat(data["datetime"].replace("Z", "+00:00"))
    except Exception as e:
        print(f"国际时间服务器请求失败: {e}")
    
    # 所有服务器都失败时才使用本地时间
    print("警告: 无法获取网络时间，使用本地时间")
    return datetime.datetime.now()

def get_machine_id():
    """获取机器唯一ID，防止删除配置文件重新开始试用"""
    import uuid
    import hashlib
    
    try:
        # 获取MAC地址作为基础ID
        machine_id = str(uuid.getnode())
        
        # 加入CPU信息
        try:
            import cpuinfo
            cpu_info = cpuinfo.get_cpu_info()
            if 'serial_number' in cpu_info:
                machine_id += cpu_info['serial_number']
            elif 'processor_id' in cpu_info:
                machine_id += cpu_info['processor_id']
        except:
            pass
            
        # 加入磁盘序列号 - 不同操作系统采用不同方法
        try:
            if os.name == 'nt':  # Windows
                import subprocess
                disk_info = subprocess.check_output('wmic diskdrive get SerialNumber', shell=True).decode()
                serial_lines = [line.strip() for line in disk_info.split('\n') if line.strip() and line.strip() != 'SerialNumber']
                if serial_lines:
                    machine_id += serial_lines[0]
            elif sys.platform == 'darwin':  # macOS
                import subprocess
                result = subprocess.check_output('ioreg -l | grep IOPlatformSerialNumber', shell=True).decode()
                if result:
                    serial = result.split('=')[-1].strip().replace('"', '')
                    machine_id += serial
            else:  # Linux
                try:
                    with open('/etc/machine-id', 'r') as f:
                        machine_id += f.read().strip()
                except:
                    pass
        except:
            pass
        
        # 计算哈希值
        hash_obj = hashlib.md5(machine_id.encode())
        return hash_obj.hexdigest()
    except:
        # 如果出错，返回一个随机ID
        return str(uuid.uuid4())

def check_trial():
    """检查试用期状态"""
    # 根据操作系统选择适当的用户数据目录
    if os.name == 'nt':  # Windows
        app_data = os.getenv('APPDATA')
    elif sys.platform == 'darwin':  # macOS
        app_data = os.path.expanduser('~/Library/Application Support')
    else:  # Linux
        app_data = os.path.expanduser('~/.config')
    
    license_dir = os.path.join(app_data, 'CADtoExcel')
    license_file = os.path.join(license_dir, 'trial.json')
    os.makedirs(license_dir, exist_ok=True)
    
    machine_id = get_machine_id()
    
    if not os.path.exists(license_file):
        # 首次运行，记录安装时间和机器ID
        install_time = get_network_time()
        expiry_time = install_time + datetime.timedelta(days=3)
        
        data = {
            'install_date': install_time.isoformat(),
            'expiry_date': expiry_time.isoformat(),
            'machine_id': machine_id,
            'trial': True
        }
        
        with open(license_file, 'w') as f:
            json.dump(data, f)
        
        return True, f"试用期开始，将于 {expiry_time.strftime('%Y-%m-%d %H:%M')} 结束"
    
    # 读取试用期信息
    with open(license_file, 'r') as f:
        data = json.load(f)
    
    # 检查机器ID是否匹配
    stored_machine_id = data.get('machine_id')
    if stored_machine_id and stored_machine_id != machine_id:
        return False, "试用许可与此计算机不匹配，请联系供应商。"
    
    expiry_date = datetime.datetime.fromisoformat(data['expiry_date'])
    current_time = get_network_time()
    
    if current_time > expiry_date:
        return False, "试用期已结束，请联系供应商购买完整版。"
    
    # 计算剩余时间
    remaining = expiry_date - current_time
    days = remaining.days
    hours = remaining.seconds // 3600
    
    return True, f"试用期剩余: {days}天{hours}小时"

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
            <title>CADtoExcel - 试用期已结束</title>
            <style>
                body {{ font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }}
                .container {{ max-width: 800px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }}
                h1 {{ color: #d9534f; }}
                .contact {{ background-color: #f9f9f9; padding: 15px; border-radius: 5px; margin-top: 30px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>试用期已结束</h1>
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
    
    # 如果试用期有效，显示剩余时间并继续启动应用
    print(f"CADtoExcel - {message}")
    
    # 在生产环境中应使用WSGI服务器（如gunicorn或uwsgi）
    # 这里仅为开发环境使用
    app.run(host='0.0.0.0', port=5001, debug=False)
