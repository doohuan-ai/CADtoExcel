#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import logging
import argparse
from typing import Dict, List, Any, Optional
import traceback

# 设置项目根目录
ROOT_DIR = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
sys.path.insert(0, ROOT_DIR)

# 导入自定义模块
from src.excel_parser import parse_excel_file, save_excel_data_to_json
from src.dwg_parser import parse_dwg_file, save_dwg_data_to_json, convert_dwg_to_dxf
from src.report_generator import generate_template_report

# 确保日志目录存在
logs_dir = os.path.join(ROOT_DIR, 'logs')
os.makedirs(logs_dir, exist_ok=True)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(logs_dir, 'cli.log'), encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

def parse_args():
    parser = argparse.ArgumentParser(description='CADtoExcel: 解析CAD文件和Excel文件，生成结构化数据和检验报告')
    
    # 创建子命令分组
    subparsers = parser.add_subparsers(dest='command', help='选择命令')
    
    # 处理命令
    process_parser = subparsers.add_parser('process', help='处理DWG和Excel文件')
    process_parser.add_argument('--dwg', required=True, help='DWG文件路径')
    process_parser.add_argument('--excel', required=True, help='Excel文件路径')
    process_parser.add_argument('--output', default='outputs', help='输出目录')
    
    # 转换命令
    convert_parser = subparsers.add_parser('convert', help='将DWG文件转换为DXF格式')
    convert_parser.add_argument('--dwg', required=True, help='要转换的DWG文件路径')
    convert_parser.add_argument('--output', default='outputs', help='输出目录')
    
    # 保留的兼容性命令行参数
    parser.add_argument('--dwg', help='DWG文件路径(兼容旧版)')
    parser.add_argument('--excel', help='Excel文件路径(兼容旧版)')
    parser.add_argument('--output', default='outputs', help='输出目录(兼容旧版)')
    parser.add_argument('--convert-dwg', help='将DWG文件转换为DXF格式(兼容旧版)')
    
    return parser.parse_args()

def main():
    """命令行主入口点"""
    args = parse_args()
    
    # 确保输出目录存在
    os.makedirs(args.output, exist_ok=True)
    
    # 处理子命令
    if args.command == 'convert':
        # 处理DWG转DXF命令
        convert_dwg_to_dxf_command(args)
    elif args.command == 'process':
        # 处理DWG和Excel文件
        process_files_command(args)
    elif args.convert_dwg:
        # 兼容性模式：使用--convert-dwg参数
        args.dwg = args.convert_dwg
        convert_dwg_to_dxf_command(args)
    elif args.dwg and args.excel:
        # 兼容性模式：同时传递了--dwg和--excel参数
        process_files_command(args)
    else:
        print("错误: 未提供所需参数。请使用 --help 查看帮助信息。")
        sys.exit(1)

def convert_dwg_to_dxf_command(args):
    """将DWG文件转换为DXF格式的命令"""
    dwg_file = args.dwg
    
    # 检查DWG文件是否存在
    if not os.path.exists(dwg_file):
        logger.error(f"DWG文件不存在: {dwg_file}")
        sys.exit(1)
    
    # 创建输出文件路径
    base_name = os.path.splitext(os.path.basename(dwg_file))[0]
    dxf_file = os.path.join(args.output, f"{base_name}.dxf")
    
    # 转换文件
    logger.info(f"开始将DWG文件 {dwg_file} 转换为DXF格式")
    try:
        result_path = convert_dwg_to_dxf(dwg_file, dxf_file)
        if result_path:
            logger.info(f"DWG转DXF成功: {result_path}")
            print(f"DWG文件已成功转换为DXF格式: {result_path}")
        else:
            logger.error("DWG转DXF失败")
            print("DWG转DXF失败，请检查日志获取详细信息")
            sys.exit(1)
    except Exception as e:
        logger.error(f"转换过程出错: {str(e)}")
        print(f"转换过程出错: {str(e)}")
        sys.exit(1)

def process_files_command(args):
    """处理DWG和Excel文件的命令"""
    dwg_file = args.dwg
    excel_file = args.excel
    
    # 检查两个文件是否都存在
    if not os.path.exists(dwg_file):
        logger.error(f"DWG文件不存在: {dwg_file}")
        sys.exit(1)
    if not os.path.exists(excel_file):
        logger.error(f"Excel文件不存在: {excel_file}")
        sys.exit(1)
    
    # 解析DWG文件
    logger.info(f"解析DWG文件: {args.dwg}")
    dwg_basename = os.path.splitext(os.path.basename(args.dwg))[0]
    dwg_output_path = os.path.join(args.output, f"{dwg_basename}_dwg.json")
    dwg_raw_output_path = os.path.join(args.output, f"{dwg_basename}_dwg_raw_data.json")
    
    dwg_data, dwg_raw_data = parse_dwg_file(args.dwg)
    save_dwg_data_to_json(dwg_data, dwg_output_path)
    
    # 保存原始DWG数据
    if dwg_raw_data:
        with open(dwg_raw_output_path, 'w', encoding='utf-8') as f:
            json.dump(dwg_raw_data, f, ensure_ascii=False, indent=2)
            logger.info(f"原始DWG数据已保存至: {dwg_raw_output_path}")
    
    logger.info(f"DWG文件解析完成: {dwg_output_path}")
    
    # 解析Excel文件
    logger.info(f"解析Excel文件: {args.excel}")
    excel_basename = os.path.splitext(os.path.basename(args.excel))[0]
    excel_output_path = os.path.join(args.output, f"{excel_basename}_excel.json")
    
    excel_data = parse_excel_file(args.excel)
    save_excel_data_to_json(excel_data, excel_output_path)
    logger.info(f"Excel文件解析完成: {excel_output_path}")
    
    # 生成检验报告
    try:
        # 定义映射文件和模板文件路径
        appearance_map_file = os.path.join(ROOT_DIR, 'maps', 'appearance_map.json')
        template_file = os.path.join(ROOT_DIR, 'maps', 'report_map.xlsx')
        
        # 确保外观要求对照表存在
        if not os.path.exists(appearance_map_file):
            logger.error(f"外观要求对照表不存在: {appearance_map_file}")
            sys.exit(1)
        
        # 确保报告模板文件存在
        if not os.path.exists(template_file):
            logger.error(f"报告模板文件不存在: {template_file}")
            sys.exit(1)
        
        logger.info("开始生成检验报告")
        report_path = generate_template_report(
            dwg_file=dwg_output_path, 
            excel_file=excel_output_path,
            appearance_map_file=appearance_map_file,
            template_file=template_file,
            output_dir=args.output,
            original_excel_file=args.excel
        )
        
        logger.info(f"检验报告已生成: {report_path}")
        print(f"处理完成！检验报告已生成: {report_path}")
        
    except Exception as e:
        logger.error(f"生成检验报告时出错: {str(e)}")
        print(f"生成检验报告时出错: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main() 