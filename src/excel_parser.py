#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import json
import pandas as pd
from typing import Dict, List, Any, Optional

def parse_excel_file(file_path: str) -> Dict[str, Any]:
    """
    解析工艺流程卡Excel文件，提取工序信息
    
    Args:
        file_path: Excel文件路径
        
    Returns:
        Dict[str, Any]: 提取的工序数据
    """
    # 检查文件扩展名，使用相应的引擎
    file_ext = os.path.splitext(file_path)[1].lower()
    
    if file_ext == '.xls':
        engine = 'xlrd'
    elif file_ext == '.xlsx':
        engine = 'openpyxl'
    else:
        raise ValueError(f"不支持的文件格式: {file_ext}，仅支持 .xls 或 .xlsx")
    
    # 读取Excel文件中的第一个sheet
    xlsx = pd.ExcelFile(file_path, engine=engine)
    sheet_names = xlsx.sheet_names
    
    if not sheet_names:
        raise ValueError(f"Excel文件 {file_path} 不包含任何工作表")
    
    # 获取第一个工作表（通常是最左边最新的）
    first_sheet_name = sheet_names[0]
    
    result = {
        "file_name": os.path.basename(file_path),
        "processes": []  # 存储工序数据
    }
    
    # 读取第一个工作表的数据
    df = pd.read_excel(xlsx, sheet_name=first_sheet_name, header=None)
    
    # 查找工序行（通常是第7行，索引为6）
    process_header_row = None
    for i in range(min(10, len(df))):  # 只在前10行中查找
        if df.iloc[i, 0] == "工序":
            process_header_row = i
            break
    
    if process_header_row is None:
        raise ValueError(f"在Excel工作表 {first_sheet_name} 中未找到工序标题行")
    
    # 从工序行的下一行开始提取数据
    start_row = process_header_row + 1
    
    # 记录上一个有效的工序名称和代码
    last_valid_process_name = ""
    last_valid_process_code = ""
    
    # 标记上一行是否被跳过
    previous_row_skipped = False
    
    # 构建工序数据列表
    for i in range(start_row, len(df)):
        # 获取当前行
        row = df.iloc[i]
        
        # 获取工序名称、工序代码和工序说明
        process_name = row[0] if not pd.isna(row[0]) else ""
        process_code = row[1] if not pd.isna(row[1]) else ""
        process_desc = row[2] if not pd.isna(row[2]) else ""
        
        # 将所有字段转换为字符串并去除空格
        process_name = str(process_name).strip()
        process_code = str(process_code).strip()
        process_desc = str(process_desc).strip()
        
        # 如果已经遇到了"以上全程保护外观无划痕伤"，则认为工序部分结束
        if "以上全程保护外观无划痕伤" in process_desc:
            break
        
        # 增加新的规则：如果上一行被跳过且当前行的process_code为空，则当前行也跳过
        if previous_row_skipped and process_code == "":
            previous_row_skipped = True  # 当前行也被跳过
            continue
        
        # 应用处理逻辑
        if process_name == "":
            if process_code != "":
                # 如果process_name为空但process_code有值，使用上一个有效的process_name
                process_name = last_valid_process_name
                previous_row_skipped = False  # 当前行不跳过
            elif process_desc != "":
                # 如果process_name和process_code都为空但process_desc有值，使用上一个有效的process_name和process_code
                process_name = last_valid_process_name
                process_code = last_valid_process_code
                previous_row_skipped = False  # 当前行不跳过
            else:
                # 如果所有字段都为空，跳过该条目
                previous_row_skipped = True  # 当前行被跳过
                continue
        else:
            previous_row_skipped = False  # 当前行不跳过
        
        # 创建工序记录
        process = {
            "process_name": process_name,
            "process_code": process_code,
            "process_desc": process_desc
        }
        
        # 更新最后有效的工序名称和代码
        if process_name != "":
            last_valid_process_name = process_name
        if process_code != "":
            last_valid_process_code = process_code
            
        # 添加到结果列表
        result["processes"].append(process)
    
    return result

def save_excel_data_to_json(excel_data: Dict[str, Any], output_path: str) -> None:
    """
    将工序数据保存为JSON文件
    
    Args:
        excel_data: 解析后的工序数据
        output_path: 输出JSON文件路径
    """
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 写入JSON文件
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(excel_data, f, ensure_ascii=False, indent=2)
    
    print(f"工序数据已保存至: {output_path}")

if __name__ == "__main__":
    # 测试代码
    test_file = os.path.join('test', '81206851-03.xls')
    excel_data = parse_excel_file(test_file)
    
    # 获取文件名（不含扩展名）用于生成输出文件名
    basename = os.path.splitext(os.path.basename(test_file))[0]
    output_path = os.path.join('outputs', f'{basename}_excel.json')
    
    save_excel_data_to_json(excel_data, output_path)
    
    # 打印摘要信息
    print(f"文件名: {excel_data['file_name']}")
    print(f"工序数量: {len(excel_data['processes'])}")
    for process in excel_data['processes']:
        print(f"工序: {process['process_name']}, 代码: {process['process_code']}, 说明: {process['process_desc']}") 