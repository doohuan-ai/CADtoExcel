#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import datetime
import hashlib
import uuid

def get_machine_id():
    """获取机器唯一ID，防止删除配置文件重新开始试用"""
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

def get_trial_info():
    """获取试用信息文件路径和机器ID"""
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
    
    return license_file, machine_id

def check_dwg_to_dxf_usage():
    """
    检查DWG转DXF功能的使用次数
    
    Returns:
        tuple: (是否可以使用, 消息, 当前使用次数, 剩余次数)
    """
    license_file, machine_id = get_trial_info()
    
    if not os.path.exists(license_file):
        # 首次运行，记录机器ID和使用次数
        data = {
            'install_date': datetime.datetime.now().isoformat(),
            'machine_id': machine_id,
            'trial': True,
            'dxf_usage_count': 0,
            'max_usage': 15
        }
        
        with open(license_file, 'w') as f:
            json.dump(data, f)
        
        return True, "可以使用该功能", 0, 15
    
    # 读取试用信息
    with open(license_file, 'r') as f:
        data = json.load(f)
    
    # 检查机器ID是否匹配
    stored_machine_id = data.get('machine_id')
    if stored_machine_id and stored_machine_id != machine_id:
        return False, "试用许可与此计算机不匹配，请联系供应商。", 0, 0
    
    # 获取当前使用次数
    usage_count = data.get('dxf_usage_count', 0)
    max_usage = data.get('max_usage', 15)
    
    # 检查是否超过最大使用次数
    if usage_count >= max_usage:
        return False, "本软件试用次数已用完，请联系供应商购买完整版。", usage_count, 0
    
    # 仍然可以使用
    remaining = max_usage - usage_count
    return True, f"本软件可以使用，已使用{usage_count}次，剩余{remaining}次", usage_count, remaining

def increment_dwg_to_dxf_usage():
    """
    增加DWG转DXF功能的使用次数计数
    
    Returns:
        tuple: (是否可以使用, 消息, 更新后的使用次数, 剩余次数)
    """
    license_file, machine_id = get_trial_info()
    
    if not os.path.exists(license_file):
        # 首次使用，初始化计数为1
        data = {
            'install_date': datetime.datetime.now().isoformat(),
            'machine_id': machine_id,
            'trial': True,
            'dxf_usage_count': 1,
            'max_usage': 15
        }
        
        with open(license_file, 'w') as f:
            json.dump(data, f)
        
        return True, "软件首次使用", 1, 14
    
    # 读取试用信息
    with open(license_file, 'r') as f:
        data = json.load(f)
    
    # 检查机器ID是否匹配
    stored_machine_id = data.get('machine_id')
    if stored_machine_id and stored_machine_id != machine_id:
        return False, "试用许可与此计算机不匹配，请联系供应商。", 0, 0
    
    # 获取当前使用次数
    usage_count = data.get('dxf_usage_count', 0)
    max_usage = data.get('max_usage', 15)
    
    # 检查是否超过最大使用次数
    if usage_count >= max_usage:
        return False, "本软件试用次数已用完，请联系供应商购买完整版。", usage_count, 0
    
    # 更新使用次数
    usage_count += 1
    data['dxf_usage_count'] = usage_count
    
    with open(license_file, 'w') as f:
        json.dump(data, f)
    
    remaining = max_usage - usage_count
    
    if usage_count >= max_usage:
        return False, "本软件试用次数已用完，请联系供应商购买完整版。", usage_count, 0
    
    return True, f"本软件使用成功，已使用{usage_count}次，剩余{remaining}次", usage_count, remaining 