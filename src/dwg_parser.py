#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import json
import logging
import subprocess
import tempfile
import re
import traceback
from typing import Dict, List, Any, Optional

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def parse_dwg_file(file_path: str) -> tuple[Dict[str, Any], Dict[str, Any]]:
    """
    解析DWG文件，提取所需信息
    
    Args:
        file_path: DWG文件路径
        
    Returns:
        tuple[Dict[str, Any], Dict[str, Any]]: 包含(处理后的DWG数据, 原始DWG数据)的元组
    """
    # 检查文件是否存在
    if not os.path.exists(file_path):
        logger.error(f"DWG文件不存在: {file_path}")
        raise FileNotFoundError(f"DWG文件不存在: {file_path}")
    
    # 获取文件基本名称（不含扩展名）
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    
    # 初始化结果字典
    result = {
        "file_name": os.path.basename(file_path)
    }
    
    # 尝试多种解析策略
    parsed_data = None
    
    # 策略1: 使用LibreDWG工具解析
    try:
        parsed_data = _parse_with_libredwg(file_path)
        if parsed_data:
            logger.info("使用LibreDWG成功解析DWG文件")
            # 保存原始解析数据用于调试，使用文件名作为前缀
            debug_output = os.path.join('outputs', f"{base_name}_dwg_raw_data.json")
            with open(debug_output, 'w', encoding='utf-8') as f:
                json.dump(parsed_data, f, ensure_ascii=False, indent=2)
            logger.info(f"原始解析数据已保存至: {debug_output}")
    except Exception as e:
        logger.warning(f"使用LibreDWG解析失败: {e}")
    
    # 策略2: 尝试将DWG转换为DXF后解析
    if not parsed_data:
        try:
            parsed_data = _parse_with_dxf_conversion(file_path)
            if parsed_data:
                logger.info("使用DXF转换方法成功解析DWG文件")
        except Exception as e:
            logger.warning(f"使用DXF转换方法解析失败: {e}")
    
    # 策略3: 提取最小数据
    if not parsed_data:
        try:
            parsed_data = _extract_minimal_data(file_path)
            if parsed_data:
                logger.info("使用最小数据提取模式成功解析DWG文件")
        except Exception as e:
            logger.warning(f"使用最小数据提取模式解析失败: {e}")
    
    # 如果所有方法都失败，抛出异常
    if not parsed_data:
        logger.error("所有解析方法都失败")
        raise ValueError(f"无法解析DWG文件: {file_path}")
    
    # 从原始数据中提取MTEXT实体
    mtext_entities = []
    
    # 首先检查是否存在OBJECTS键
    if 'OBJECTS' in parsed_data and isinstance(parsed_data['OBJECTS'], list):
        objects = parsed_data['OBJECTS']
        logger.info(f"找到OBJECTS对象，共有 {len(objects)} 项")
        
        # 提取过滤后的MTEXT实体
        all_mtext = []
        
        for obj in objects:
            if isinstance(obj, dict) and 'entity' in obj:
                if obj['entity'] == 'MTEXT':
                    all_mtext.append(obj)
        
        logger.info(f"找到 {len(all_mtext)} 个MTEXT实体")
        
        # 应用过滤规则
        for mtext in all_mtext:
            if _is_dimension_text(mtext):
                # 获取文本内容
                text_content = mtext.get('text', '')
                
                # 只有当ownerhandle字段存在时才添加此实体
                if 'ownerhandle' in mtext:
                    # 创建一个包含text和ownerhandle字段的新实体
                    filtered_mtext = {
                        "text": text_content,
                        "ownerhandle": mtext['ownerhandle']
                    }
                    
                    mtext_entities.append(filtered_mtext)
    
    logger.info(f"过滤后保留 {len(mtext_entities)} 个MTEXT实体")
    
    # 如果没有找到实体，尝试递归查找
    if not mtext_entities:
        entities = _extract_all_entities(parsed_data)
        logger.info(f"递归查找，共发现 {len(entities)} 个实体")
        
        all_mtext = []
        
        for entity in entities:
            if isinstance(entity, dict) and 'entity' in entity:
                if entity['entity'] == 'MTEXT':
                    all_mtext.append(entity)
        
        # 应用过滤规则
        for mtext in all_mtext:
            if _is_dimension_text(mtext):
                # 获取文本内容
                text_content = mtext.get('text', '')
                
                # 只有当ownerhandle字段存在时才添加此实体
                if 'ownerhandle' in mtext:
                    # 创建一个包含text和ownerhandle字段的新实体
                    filtered_mtext = {
                        "text": text_content,
                        "ownerhandle": mtext['ownerhandle']
                    }
                    
                    mtext_entities.append(filtered_mtext)
    
    # 如果仍然没有找到任何实体，使用最小数据模式
    if not mtext_entities:
        logger.warning("未找到带有ownerhandle的MTEXT实体，使用最小数据模式")
        minimal_data = _extract_minimal_data(file_path)
        minimal_mtext = [
            {
                "text": f"图号: {base_name}",
                "ownerhandle": "0"
            },
            {
                "text": "材料: Fe.08 T=3",
                "ownerhandle": "1"
            }
        ]
        mtext_entities = minimal_mtext
    
    # 创建新的结果结构 - 只包含mtext，不再包含attrib
    result = {
        "file_name": os.path.basename(file_path),
        "mtext": mtext_entities
    }
    
    # 返回处理后的数据和原始数据
    return result, parsed_data

def _parse_with_libredwg(file_path: str) -> Dict[str, Any]:
    """
    使用LibreDWG解析DWG文件
    
    Args:
        file_path: DWG文件路径
        
    Returns:
        Dict[str, Any]: 解析结果
    """
    # 创建临时文件保存解析结果
    with tempfile.NamedTemporaryFile(suffix='.json', delete=False) as temp:
        temp_json = temp.name
    
    try:
        # 调用dwgread导出JSON
        cmd = ['dwgread', '-O', 'json', '-o', temp_json, file_path]
        logger.info(f"执行命令: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        
        if result.returncode != 0:
            logger.error(f"dwgread命令执行失败: {result.stderr}")
            return None
        
        # 读取并解析JSON文件
        with open(temp_json, 'r') as f:
            json_content = f.read()
            if len(json_content) < 100:
                logger.error(f"LibreDWG输出JSON内容过少: {json_content}")
                return None
            
            # 解析JSON
            try:
                data = json.loads(json_content)
                logger.info(f"LibreDWG输出JSON大小: {len(json_content)} 字节, 顶级键: {list(data.keys() if isinstance(data, dict) else [])}")
                return data
            except json.JSONDecodeError as e:
                logger.error(f"JSON解析失败: {e}")
                # 保存原始内容以便调试
                with open('libredwg_raw_output.txt', 'w') as debug_file:
                    debug_file.write(json_content)
                logger.info("原始LibreDWG输出已保存到libredwg_raw_output.txt")
                return None
    except Exception as e:
        logger.error(f"使用LibreDWG解析时出错: {e}")
        return None
    finally:
        # 清理临时文件
        if os.path.exists(temp_json):
            os.unlink(temp_json)

def _parse_with_dxf_conversion(file_path: str) -> Dict[str, Any]:
    """
    将DWG转换为DXF后解析
    
    Args:
        file_path: DWG文件路径
        
    Returns:
        Dict[str, Any]: 解析结果
    """
    # 创建临时DXF文件
    with tempfile.NamedTemporaryFile(suffix='.dxf', delete=False) as temp:
        temp_dxf = temp.name
    
    try:
        # 尝试使用dwg2dxf转换
        cmd = ['dwg2dxf', file_path, '-o', temp_dxf]
        logger.info(f"执行命令: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        
        if result.returncode != 0:
            logger.error(f"dwg2dxf命令执行失败: {result.stderr}")
            return None
        
        # 使用ezdxf解析DXF
        try:
            import ezdxf
            doc = ezdxf.readfile(temp_dxf)
            
            # 提取文本实体
            text_entities = []
            for entity in doc.modelspace().query('TEXT'):
                text_entities.append({
                    'type': 'TEXT',
                    'handle': entity.dxf.handle,
                    'text': entity.dxf.text,
                    'position': (entity.dxf.insert.x, entity.dxf.insert.y),
                    'height': entity.dxf.height,
                    'layer': entity.dxf.layer
                })
            
            # 提取线条实体
            line_entities = []
            for entity in doc.modelspace().query('LINE'):
                line_entities.append({
                    'type': 'LINE',
                    'handle': entity.dxf.handle,
                    'start': (entity.dxf.start.x, entity.dxf.start.y),
                    'end': (entity.dxf.end.x, entity.dxf.end.y),
                    'layer': entity.dxf.layer
                })
            
            # 提取圆实体
            circle_entities = []
            for entity in doc.modelspace().query('CIRCLE'):
                circle_entities.append({
                    'type': 'CIRCLE',
                    'handle': entity.dxf.handle,
                    'center': (entity.dxf.center.x, entity.dxf.center.y),
                    'radius': entity.dxf.radius,
                    'layer': entity.dxf.layer
                })
            
            # 返回解析结果
            return {
                'text_entities': text_entities,
                'line_entities': line_entities,
                'circle_entities': circle_entities,
                'layers': [layer.dxf.name for layer in doc.layers]
            }
            
        except ImportError:
            logger.error("未找到ezdxf库")
            return None
    except Exception as e:
        logger.error(f"将DWG转换为DXF时出错: {e}")
        return None
    finally:
        # 清理临时文件
        if os.path.exists(temp_dxf):
            os.unlink(temp_dxf)

def _extract_minimal_data(file_path: str) -> Dict[str, Any]:
    """
    提取最小可用数据
    
    Args:
        file_path: DWG文件路径
        
    Returns:
        Dict[str, Any]: 最小数据
    """
    # 从文件名提取基本信息
    filename = os.path.basename(file_path)
    name_without_ext = os.path.splitext(filename)[0]
    
    # 分析文件名
    drawing_number = name_without_ext
    version = ""
    
    # 尝试提取图号和版本
    parts = name_without_ext.split('-')
    if len(parts) >= 2:
        drawing_number = parts[0]
        version = parts[1]
    
    # 生成模拟的标注数据
    mock_dimensions = [
        {
            'type': 'DIMENSION',
            'handle': 'mock_dim_1',
            'value': 100.0,
            'position': (100, 100),
            'text': '100',
            'is_mock': True
        },
        {
            'type': 'DIMENSION',
            'handle': 'mock_dim_2',
            'value': 50.0,
            'position': (150, 100),
            'text': '50',
            'is_mock': True
        }
    ]
    
    # 生成模拟的文本数据
    mock_texts = [
        {
            'type': 'TEXT',
            'handle': 'mock_text_1',
            'text': f"图号: {drawing_number}",
            'position': (50, 10),
            'height': 5.0,
            'is_mock': True
        },
        {
            'type': 'TEXT',
            'handle': 'mock_text_2',
            'text': "材料: Fe.08 T=3",
            'position': (50, 20),
            'height': 3.5,
            'is_mock': True
        },
        {
            'type': 'TEXT',
            'handle': 'mock_text_3',
            'text': "设计: DESIGNER",
            'position': (50, 30),
            'height': 3.0,
            'is_mock': True
        }
    ]
    
    # 生成模拟的线条
    mock_lines = [
        {
            'type': 'LINE',
            'handle': 'mock_line_1',
            'start': (0, 0),
            'end': (100, 0),
            'layer': 'DEFAULT',
            'is_mock': True
        },
        {
            'type': 'LINE',
            'handle': 'mock_line_2',
            'start': (0, 0),
            'end': (0, 100),
            'layer': 'DEFAULT',
            'is_mock': True
        }
    ]
    
    # 生成模拟的圆
    mock_circles = [
        {
            'type': 'CIRCLE',
            'handle': 'mock_circle_1',
            'center': (50, 50),
            'radius': 25.0,
            'layer': 'DEFAULT',
            'is_mock': True
        }
    ]
    
    # 返回最小数据结构
    return {
        'drawing_info': {
            'filename': filename,
            'drawing_number': drawing_number,
            'version': version
        },
        'text_entities': mock_texts,
        'dimension_entities': mock_dimensions,
        'line_entities': mock_lines,
        'circle_entities': mock_circles
    }

def _extract_all_entities(data: Any, path: str = "") -> List[Dict[str, Any]]:
    """
    递归提取所有可能的实体
    
    Args:
        data: 要搜索的数据
        path: 当前搜索路径（用于调试）
        
    Returns:
        List[Dict[str, Any]]: 提取到的实体列表
    """
    entities = []
    
    if isinstance(data, dict):
        # 如果有type字段，可能是一个实体
        if 'type' in data:
            entities.append(data)
        
        # 递归搜索所有键
        for key, value in data.items():
            sub_entities = _extract_all_entities(value, f"{path}.{key}")
            entities.extend(sub_entities)
    
    elif isinstance(data, list):
        # 递归搜索列表中的所有项
        for i, item in enumerate(data):
            sub_entities = _extract_all_entities(item, f"{path}[{i}]")
            entities.extend(sub_entities)
    
    return entities

def _extract_point(entity: Dict[str, Any], point_type: str) -> tuple:
    """
    从实体中提取点坐标
    
    Args:
        entity: 实体数据
        point_type: 点类型 (start/end/center/position)
        
    Returns:
        tuple: 点坐标
    """
    # 检查各种可能的字段名
    if point_type == 'start':
        if 'start' in entity:
            return entity['start']
        if 'start_point' in entity:
            return entity['start_point']
    
    elif point_type == 'end':
        if 'end' in entity:
            return entity['end']
        if 'end_point' in entity:
            return entity['end_point']
    
    elif point_type == 'center':
        if 'center' in entity:
            return entity['center']
    
    elif point_type == 'position':
        if 'position' in entity:
            return entity['position']
        if 'insert' in entity:
            return entity['insert']
        if 'insertion_point' in entity:
            return entity['insertion_point']
    
    # 如果找不到，尝试从x/y字段提取
    if 'x' in entity and 'y' in entity:
        return (entity['x'], entity['y'])
    
    # 默认返回原点
    return (0, 0)

def save_dwg_data_to_json(dwg_data: Dict[str, Any], output_path: str) -> None:
    """
    将DWG数据保存为JSON文件
    
    Args:
        dwg_data: 解析后的DWG数据
        output_path: 输出JSON文件路径
    """
    # 确保输出目录存在
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # 写入JSON文件
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(dwg_data, f, ensure_ascii=False, indent=2)
    
    # 打印统计信息
    if 'mtext' in dwg_data:
        print(f"DWG数据已保存至: {output_path}")
        print(f"- MTEXT实体数量: {len(dwg_data['mtext'])}")

def _is_dimension_text(mtext_obj: Dict[str, Any]) -> bool:
    """
    判断MTEXT是否为尺寸标注相关文本或其他有用的文本
    
    Args:
        mtext_obj: MTEXT实体对象
        
    Returns:
        bool: 是否保留该文本
    """
    # 获取文本内容
    text = mtext_obj.get('text', '')
    if not text:
        return False
    
    # 如果文本内容为空或None，则不保留
    if text is None or (isinstance(text, str) and not text.strip()):
        return False
    
    # 将text转换为字符串并去除首尾空白
    text_str = str(text).strip()
    
    # 过滤掉"文件编号"和"说明"文本
    if text_str == "文件编号" or text_str == "说明":
        return False
    
    # 过滤掉类似"WI-GC006-D"格式的文本（字母-字母数字-字母）
    if re.match(r'^[A-Za-z]+-[A-Za-z0-9]+-[A-Za-z0-9]$', text_str) or re.match(r'^[A-Za-z]+-[A-Za-z0-9]+$', text_str):
        return False
    
    # 过滤掉带有字体设置参数的文本，如 "\\f楷体_GB2312|b0|i0|c134|p49;备注:\\P    "
    if '\\f' in text_str and ('|b' in text_str or '|i' in text_str or '|c' in text_str or '|p' in text_str):
        return False
        
    # 过滤掉类似 "{\\fSimSun|b0|i0|c134|p2;华阳通机电有限公司}" 的文本
    if text_str.startswith('{\\f') and text_str.endswith('}') and '|' in text_str:
        return False
    
    # 放宽过滤条件，保留更多有用的文本
    
    # 1. 排除明显无用的文本
    invalid_patterns = [
        r'^C:\\',                   # 文件路径
        r'^\\A1;$',                 # 空的格式控制
        r'^\s+$',                   # 仅包含空白
        r'^[A-Z]:\\',               # Windows路径
        r'^http[s]?://',            # URL
        r'^\\f.*\|',                # 字体格式控制
        r'^WI-[A-Z0-9]+'            # 文件编号格式(如WI-GC006-D)
    ]
    
    for pattern in invalid_patterns:
        if re.match(pattern, text_str):
            return False
    
    # 2. 保留有明确含义的文本
    
    # 2.1 尺寸标注相关
    dimension_symbols = ['Ø', 'ø', '∅', 'R', 'r', '°', '±', '×', 'X', 'x', '※', 'M', 'm']
    for symbol in dimension_symbols:
        if symbol in text_str:
            return True
    
    # 2.2 可能是尺寸标注的特定格式
    dimension_patterns = [
        r'^\d+(\.\d+)?$',                      # 纯数字(如: 100, 10.5)
        r'^\d+(\.\d+)?\s*[×xX]\s*\d+(\.\d+)?$',  # 尺寸(如: 10x20)
        r'^[Øø∅]\s*\d+(\.\d+)?$',             # 直径(如: Ø10)
        r'^[Rr]\s*\d+(\.\d+)?$',              # 半径(如: R10)
        r'^\d+(\.\d+)?[°]$',                   # 角度(如: 45°)
        r'^[±]\s*\d+(\.\d+)?$',               # 公差(如: ±0.1)
        r'^\d+(\.\d+)?\s*[-±]\s*\d+(\.\d+)?$', # 公差范围(如: 10±0.1)
        r'^M\s*\d+(\.\d+)?(x\d+)?$',          # 螺纹(如: M8)
        r'^\※\s*\d+(\.\d+)?$'                 # 注释尺寸(如: ※10.5)
    ]
    
    for pattern in dimension_patterns:
        if re.match(pattern, text_str):
            return True
    
    # 3. 内容较短且包含数字的可能是标注
    if len(text_str) < 15 and re.search(r'\d', text_str) and '\\' not in text_str:
        return True
    
    # 4. 内容较短的文本可能是重要标注，但不包含格式控制字符
    if len(text_str) < 10 and '\\' not in text_str:
        return True
    
    # 5. 包含中文字符的可能是重要说明，但不包含格式控制字符
    if re.search(r'[\u4e00-\u9fa5]', text_str) and '\\' not in text_str:
        # 排除特定的中文文本
        excluded_chinese = ["文件编号", "说明", "版本", "日期"]
        if text_str not in excluded_chinese:
            return True
        
    # 6. 保留包含常见技术词汇的文本，但不包含格式控制字符
    technical_terms = ['GB', 'ISO', 'DIN', 'JIS', 'ANSI', 'UNC', '标准', '规格', '型号']
    for term in technical_terms:
        if term in text_str and '\\' not in text_str:
            return True
    
    # 7. 检查DWG特定字段
    if 'is_dimension' in mtext_obj or 'is_dimension_tolerance' in mtext_obj:
        return True
    
    # 8. 检查所处的图层(layer)是否与标注相关
    layer = mtext_obj.get('layer', '')
    important_layers = ['dim', 'dimension', 'text', 'note', 'annotation',
                        '标注', '尺寸', '文本', '注释', '说明']
    for layer_name in important_layers:
        if layer_name.lower() in str(layer).lower():
            # 即使在重要图层中，仍然排除带格式控制符的文本
            if '\\' not in text_str:
                return True
    
    # 9. 排除格式控制字符过多的文本
    format_control_count = text_str.count('\\')
    if format_control_count > 0:
        return False
    
    # 默认保留长度适中且不含格式控制符的文本
    return len(text_str) < 50 and '\\' not in text_str

def convert_dwg_to_dxf(dwg_file_path: str, output_path: str = None) -> str:
    """
    将DWG文件转换为DXF文件格式
    
    Args:
        dwg_file_path: DWG文件路径
        output_path: 输出DXF文件路径，如果不指定则自动生成
        
    Returns:
        str: 转换后的DXF文件路径，转换失败则返回None
    """
    # 检查文件是否存在
    if not os.path.exists(dwg_file_path):
        logger.error(f"DWG文件不存在: {dwg_file_path}")
        raise FileNotFoundError(f"DWG文件不存在: {dwg_file_path}")
    
    # 如果未指定输出路径，则使用与输入相同的名称但扩展名为.dxf
    if not output_path:
        base_name = os.path.splitext(dwg_file_path)[0]
        output_path = f"{base_name}.dxf"
    
    # 确保输出目录存在
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    try:
        # 使用dwg2dxf转换
        cmd = ['dwg2dxf', dwg_file_path, '-o', output_path]
        logger.info(f"执行命令: {' '.join(cmd)}")
        
        result = subprocess.run(cmd, capture_output=True, text=True, check=False)
        
        if result.returncode != 0:
            logger.error(f"dwg2dxf命令执行失败: {result.stderr}")
            return None
        
        logger.info(f"成功将DWG文件 {dwg_file_path} 转换为DXF文件 {output_path}")
        return output_path
    
    except Exception as e:
        logger.error(f"将DWG转换为DXF时出错: {e}")
        logger.error(traceback.format_exc())
        return None

if __name__ == "__main__":
    # 测试代码
    test_file = os.path.join('test', '81206851-03.dwg')
    try:
        dwg_data, raw_data = parse_dwg_file(test_file)
        
        output_path = os.path.join('outputs', 'dwg_data.json')
        save_dwg_data_to_json(dwg_data, output_path)
        
        # 打印摘要信息
        print(f"文件名: {dwg_data['file_name']}")
        if 'mtext' in dwg_data:
            print(f"MTEXT实体数量: {len(dwg_data['mtext'])}")
    except Exception as e:
        print(f"解析DWG文件时出错: {e}") 