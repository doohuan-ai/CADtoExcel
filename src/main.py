#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import logging
import argparse
from typing import Dict, List, Any, Optional
import traceback

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.abspath(os.path.dirname(os.path.dirname(__file__))))
# 添加src目录到路径
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

# 导入自定义模块
from excel_parser import parse_excel_file, save_excel_data_to_json
from dwg_parser import parse_dwg_file, save_dwg_data_to_json

# 确保日志目录存在
logs_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
os.makedirs(logs_dir, exist_ok=True)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(logs_dir, 'cadtoexcel.log'), encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

def load_appearance_map(file_path: str = 'maps/appearance_map.json') -> Dict[str, Any]:
    """
    从JSON文件加载外观要求对照表
    
    Args:
        file_path: 外观要求对照表的JSON文件路径
        
    Returns:
        Dict[str, Any]: 外观要求对照表
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            appearance_map = json.load(f)
        logger.info(f"成功从 {file_path} 加载外观要求对照表")
        return appearance_map
    except Exception as e:
        logger.error(f"加载外观要求对照表时出错: {e}")
        logger.error(traceback.format_exc())
        return {}

def process_files(dwg_file: str, excel_file: str, output_dir: str = 'outputs') -> Dict[str, str]:
    """
    处理DWG和Excel文件，生成JSON输出
    
    Args:
        dwg_file: DWG文件路径
        excel_file: Excel文件路径
        output_dir: 输出目录
        
    Returns:
        Dict[str, str]: 包含输出文件路径的字典
    """
    logger.info("开始处理文件...")
    
    # 确保输出目录存在
    os.makedirs(output_dir, exist_ok=True)
    
    # 生成输出文件名（基于输入文件的基本名称）
    base_name = os.path.splitext(os.path.basename(dwg_file))[0]
    output_files = {
        "dwg_data": os.path.join(output_dir, f"{base_name}_dwg.json"),
        "excel_data": os.path.join(output_dir, f"{base_name}_excel.json")
    }
    
    # 解析DWG文件
    logger.info(f"解析DWG文件: {dwg_file}")
    try:
        dwg_data, dwg_raw_data = parse_dwg_file(dwg_file)
        save_dwg_data_to_json(dwg_data, output_files["dwg_data"])
        
        # 保存原始DWG数据
        if dwg_raw_data:
            raw_data_path = os.path.join(output_dir, f"{base_name}_dwg_raw_data.json")
            with open(raw_data_path, 'w', encoding='utf-8') as f:
                json.dump(dwg_raw_data, f, ensure_ascii=False, indent=2)
                logger.info(f"原始DWG数据已保存至: {raw_data_path}")
    except Exception as e:
        logger.error(f"解析DWG文件时出错: {e}")
        logger.error(traceback.format_exc())
        # 使用基本信息创建一个简化版DWG数据
        dwg_data = {
            "file_name": os.path.basename(dwg_file),
            "metadata": {
                "version": "",
                "header_variables": {},
                "drawing_info": {
                    "filename": os.path.basename(dwg_file),
                    "drawing_number": base_name,
                }
            },
            "entities": {
                "lines": [],
                "circles": [],
                "arcs": [],
                "texts": [],
                "polylines": [],
                "blocks": [],
                "dimensions": [],
                "other": []
            },
            "layers": []
        }
        save_dwg_data_to_json(dwg_data, output_files["dwg_data"])
        logger.info("已创建简化版DWG数据")
    
    # 解析Excel文件
    logger.info(f"解析Excel文件: {excel_file}")
    try:
        excel_data = parse_excel_file(excel_file)
        save_excel_data_to_json(excel_data, output_files["excel_data"])
    except Exception as e:
        logger.error(f"解析Excel文件时出错: {e}")
        logger.error(traceback.format_exc())
        raise
    
    logger.info(f"所有数据已保存到目录: {output_dir}")
    return output_files

def main():
    """
    主函数，解析命令行参数并处理文件
    """
    parser = argparse.ArgumentParser(description='CAD与Excel文件处理工具')
    parser.add_argument('--dwg', required=True, help='输入DWG文件路径')
    parser.add_argument('--excel', required=True, help='输入Excel文件路径')
    parser.add_argument('--output', default='outputs', help='输出目录路径')
    
    args = parser.parse_args()
    
    # 检查文件是否存在
    if not os.path.exists(args.dwg):
        logger.error(f"DWG文件不存在: {args.dwg}")
        return 1
    
    if not os.path.exists(args.excel):
        logger.error(f"Excel文件不存在: {args.excel}")
        return 1
    
    try:
        output_files = process_files(args.dwg, args.excel, args.output)
        logger.info("处理完成！输出文件:")
        for file_type, file_path in output_files.items():
            logger.info(f"- {file_type}: {file_path}")
        return 0
    except Exception as e:
        logger.error(f"处理文件时发生错误: {e}")
        logger.error(traceback.format_exc())
        return 1

if __name__ == "__main__":
    sys.exit(main()) 