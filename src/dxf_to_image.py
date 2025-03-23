#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import logging
import ezdxf
from ezdxf.addons.drawing import Frontend, RenderContext
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
import matplotlib.pyplot as plt
from PIL import Image
import numpy as np
import matplotlib.font_manager as fm
import traceback
import codecs

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def parse_acad_text_format(text: str) -> tuple[str, dict]:
    """
    解析AutoCAD文本格式化字符串
    
    Args:
        text: AutoCAD格式的文本字符串
        
    Returns:
        tuple[str, dict]: (纯文本, 格式信息字典)
    """
    try:
        if '\\f' not in text:
            return text, {}
        
        # 分离格式和文本内容
        format_part = text[:text.index(';')]
        content = text[text.index(';')+1:]
        
        # 解析格式参数
        format_info = {}
        parts = format_part.split('|')
        for part in parts:
            if part.startswith('\\f'):
                format_info['font'] = part[2:]
            elif part.startswith('b'):
                format_info['bold'] = part[1] == '1'
            elif part.startswith('i'):
                format_info['italic'] = part[1] == '1'
            elif part.startswith('c'):
                format_info['color'] = part[1:]
            elif part.startswith('p'):
                format_info['paragraph'] = part[1:]
        
        return content, format_info
    except:
        return text, {}

def setup_matplotlib_fonts():
    """
    设置Matplotlib的中文字体
    """
    try:
        # 尝试加载系统中的字体
        font_paths = {
            'SimSun': [
                'C:/Windows/Fonts/simsun.ttc',             # Windows
                '/usr/share/fonts/truetype/simsun.ttc',    # Linux
                '/System/Library/Fonts/SimSun.ttf'         # macOS
            ],
            'STHeiti': [
                '/System/Library/Fonts/STHeiti Light.ttc',  # macOS
                '/usr/share/fonts/truetype/stheiti.ttc'    # Linux
            ],
            'Microsoft YaHei': [
                'C:/Windows/Fonts/msyh.ttc',               # Windows
                '/usr/share/fonts/truetype/msyh.ttc'       # Linux
            ]
        }
        
        # 记录找到的字体
        found_fonts = {}
        
        # 检查每种字体
        for font_name, paths in font_paths.items():
            for path in paths:
                if os.path.exists(path):
                    found_fonts[font_name] = path
                    logger.info(f"找到字体 {font_name}: {path}")
                    break
        
        if found_fonts:
            # 设置默认字体
            plt.rcParams['font.family'] = ['sans-serif']
            plt.rcParams['font.sans-serif'] = list(found_fonts.keys()) + ['SimSun', 'Microsoft YaHei', 'STHeiti']
            logger.info(f"成功设置字体列表: {plt.rcParams['font.sans-serif']}")
        else:
            # 如果找不到预设字体，使用系统默认
            plt.rcParams['font.family'] = ['sans-serif']
            plt.rcParams['font.sans-serif'] = ['SimSun', 'Microsoft YaHei', 'STHeiti', 'SimHei', 'NSimSun', 'FangSong']
            logger.warning("未找到系统预设字体，使用默认字体列表")
        
        # 解决负号显示问题
        plt.rcParams['axes.unicode_minus'] = False
        
        return found_fonts
        
    except Exception as e:
        logger.error(f"设置中文字体时出错: {e}")
        logger.error(traceback.format_exc())
        return {}

def fix_dxf_encoding(dxf_path: str) -> str:
    """
    修复DXF文件的编码问题
    
    Args:
        dxf_path: DXF文件路径
        
    Returns:
        str: 修复后的DXF文件路径（如果需要修复则返回新文件路径，否则返回原路径）
    """
    try:
        # 读取文件内容
        with open(dxf_path, 'rb') as f:
            content = f.read()
        
        # 尝试不同的编码方式解码
        encodings = ['gbk', 'gb2312', 'gb18030', 'utf-8']  # 优先尝试中文编码
        detected_encoding = None
        decoded_text = None
        
        for enc in encodings:
            try:
                decoded_text = content.decode(enc)
                detected_encoding = enc
                break
            except UnicodeDecodeError:
                continue
        
        if not detected_encoding:
            logger.warning("无法检测到正确的编码，将使用原始文件")
            return dxf_path
        
        # 如果检测到编码，创建新的DXF文件
        new_path = dxf_path + '.fixed.dxf'
        
        # 处理特殊的中文编码标记
        if detected_encoding in ['gbk', 'gb2312', 'gb18030']:
            # 替换MTEXT中的Unicode转义序列
            decoded_text = decoded_text.replace('\\U+', '\\u')
            
            # 处理可能的其他编码问题
            lines = decoded_text.split('\n')
            processed_lines = []
            for line in lines:
                if 'MTEXT' in line or 'TEXT' in line:
                    # 对于文本实体，确保编码正确
                    try:
                        # 尝试处理Unicode转义序列
                        line = bytes(line, 'utf-8').decode('unicode_escape')
                    except:
                        pass
                processed_lines.append(line)
            decoded_text = '\n'.join(processed_lines)
        
        # 写入新文件
        with open(new_path, 'w', encoding='utf-8') as f:
            f.write(decoded_text)
        
        logger.info(f"已修复DXF文件编码 (从 {detected_encoding} 到 UTF-8): {new_path}")
        return new_path
        
    except Exception as e:
        logger.error(f"修复DXF文件编码时出错: {e}")
        logger.error(traceback.format_exc())
        return dxf_path

def convert_dxf_to_image_direct(dxf_path: str, output_path: str = None, layout_name: str = 'A3', dpi: int = 300) -> str:
    """
    使用直接渲染方式将DXF文件转换为黑底图像
    
    Args:
        dxf_path: DXF文件路径
        output_path: 输出图像路径，如果不指定则自动生成
        layout_name: 布局名称，默认为'A3'
        dpi: 图像分辨率，默认300dpi
        
    Returns:
        str: 生成的图像文件路径
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(dxf_path):
            logger.error(f"DXF文件不存在: {dxf_path}")
            raise FileNotFoundError(f"DXF文件不存在: {dxf_path}")
        
        # 设置matplotlib中文字体
        setup_matplotlib_fonts()
        
        # 如果未指定输出路径，则使用与输入相同的名称但扩展名为.png
        if not output_path:
            base_name = os.path.splitext(dxf_path)[0]
            output_path = f"{base_name}.png"
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 使用subprocess调用外部命令将DXF转换为PNG
        temp_path = output_path + '.temp.png'
        
        # 尝试使用LibreDWG工具
        import subprocess
        try:
            logger.info(f"尝试使用LibreDWG工具转换 {dxf_path}")
            cmd = ['dwg2svg', dxf_path, '-o', temp_path + '.svg']
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0 and os.path.exists(temp_path + '.svg'):
                # 使用cairosvg将SVG转换为PNG
                import cairosvg
                cairosvg.svg2png(url=temp_path + '.svg', write_to=temp_path, 
                               background_color='black', scale=2.0)
                logger.info("使用LibreDWG和cairosvg转换成功")
            else:
                raise Exception("LibreDWG转换失败")
        except:
            # 如果上面的方法失败，尝试使用Python内置的方法
            logger.info("尝试使用matplotlib直接转换DXF")
            
            # 使用matplotlib加载DXF并渲染
            from matplotlib import pyplot as plt
            from ezdxf.addons.drawing import RenderContext, Frontend
            from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
            
            # 加载DXF文件
            doc = ezdxf.readfile(dxf_path)
            
            # 获取指定布局
            try:
                layout = doc.layouts.get(layout_name)
                if not layout:
                    # 如果找不到指定布局，尝试使用模型空间
                    logger.warning(f"未找到布局 {layout_name}，使用模型空间")
                    layout = doc.modelspace()
            except Exception as e:
                logger.warning(f"获取布局失败: {e}，使用模型空间")
                layout = doc.modelspace()
            
            # 创建图形对象
            fig = plt.figure(figsize=(16, 11.3))  # A3比例
            ax = fig.add_axes([0, 0, 1, 1])
            
            # 设置背景色
            ax.set_facecolor('black')
            fig.patch.set_facecolor('black')
            
            # 创建渲染上下文
            ctx = RenderContext(doc)
            ctx.set_current_layout(layout)
            
            # 将实体绘制到图形上
            out = MatplotlibBackend(ax)
            Frontend(ctx, out).draw_layout(layout, finalize=True)
            
            # 设置图形属性
            ax.set_axis_off()
            plt.margins(0, 0)
            
            # 保存图形
            plt.savefig(temp_path, dpi=dpi, facecolor='black', 
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        
        # 使用PIL处理图像
        with Image.open(temp_path) as img:
            # 转换为RGBA
            img = img.convert('RGBA')
            
            # 获取图像数据
            data = np.array(img)
            
            # 将白色转换为亮蓝色
            white_areas = (data[:, :, 0] > 200) & (data[:, :, 1] > 200) & (data[:, :, 2] > 200)
            data[white_areas] = [0, 255, 255, 255]  # 亮蓝色
            
            # 创建新图像
            new_img = Image.fromarray(data)
            
            # 保存最终图像
            new_img.save(output_path, 'PNG')
        
        # 清理临时文件
        if os.path.exists(temp_path):
            os.remove(temp_path)
        if os.path.exists(temp_path + '.svg'):
            os.remove(temp_path + '.svg')
        
        logger.info(f"成功将DXF文件转换为图像: {output_path}")
        return output_path
    
    except Exception as e:
        logger.error(f"直接转换DXF为图像时出错: {e}")
        logger.error(traceback.format_exc())
        return None

def convert_dxf_using_external_tool(dxf_path: str, output_path: str = None) -> str:
    """
    使用外部工具将DXF文件转换为图像
    
    Args:
        dxf_path: DXF文件路径
        output_path: 输出图像路径，如果不指定则自动生成
        
    Returns:
        str: 生成的图像文件路径
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(dxf_path):
            logger.error(f"DXF文件不存在: {dxf_path}")
            raise FileNotFoundError(f"DXF文件不存在: {dxf_path}")
        
        # 如果未指定输出路径，则使用与输入相同的名称但扩展名为.png
        if not output_path:
            base_name = os.path.splitext(dxf_path)[0]
            output_path = f"{base_name}.png"
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 尝试使用ODA File Converter (如果安装了)
        import subprocess
        import platform
        
        # 检测操作系统
        system = platform.system()
        
        if system == "Windows":
            # Windows路径
            converter_paths = [
                r"C:\Program Files\ODA\ODAFileConverter\ODAFileConverter.exe",
                r"C:\Program Files (x86)\ODA\ODAFileConverter\ODAFileConverter.exe"
            ]
            
            for converter_path in converter_paths:
                if os.path.exists(converter_path):
                    # 创建一个临时目录来保存输出
                    temp_dir = os.path.join(os.path.dirname(output_path), "temp_convert")
                    os.makedirs(temp_dir, exist_ok=True)
                    
                    # 准备命令行参数
                    cmd = [
                        converter_path,
                        os.path.dirname(dxf_path),  # 输入目录
                        temp_dir,                    # 输出目录
                        "ACAD2018",                  # 输出DWG版本
                        "PNG",                       # 输出格式
                        "0",                         # 递归级别
                        "1"                          # 覆盖已存在的文件
                    ]
                    
                    # 运行转换器
                    subprocess.run(cmd)
                    
                    # 寻找生成的PNG文件
                    base_name = os.path.splitext(os.path.basename(dxf_path))[0]
                    png_file = os.path.join(temp_dir, f"{base_name}.png")
                    
                    if os.path.exists(png_file):
                        # 复制到最终位置
                        with Image.open(png_file) as img:
                            # 转换为RGBA
                            img = img.convert('RGBA')
                            
                            # 反转颜色 (黑底白字变为白底黑字)
                            img = Image.fromarray(255 - np.array(img))
                            
                            # 保存最终图像
                            img.save(output_path)
                        
                        # 清理临时文件
                        import shutil
                        shutil.rmtree(temp_dir)
                        
                        logger.info(f"成功使用ODA转换DXF为图像: {output_path}")
                        return output_path
        
        # 如果ODA转换器不可用或失败，使用其他方法
        logger.info("ODA转换器不可用，使用内置方法转换")
        return convert_dxf_to_image_direct(dxf_path, output_path)
    
    except Exception as e:
        logger.error(f"使用外部工具转换DXF时出错: {e}")
        logger.error(traceback.format_exc())
        return convert_dxf_to_image_direct(dxf_path, output_path)  # 回退到内置方法

def convert_dxf_using_inkscape(dxf_path: str, output_path: str = None) -> str:
    """
    使用Inkscape将DXF文件转换为图像（适用于无头Linux服务器）
    
    Args:
        dxf_path: DXF文件路径
        output_path: 输出图像路径，如果不指定则自动生成
        
    Returns:
        str: 生成的图像文件路径
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(dxf_path):
            logger.error(f"DXF文件不存在: {dxf_path}")
            raise FileNotFoundError(f"DXF文件不存在: {dxf_path}")
        
        # 如果未指定输出路径，则使用与输入相同的名称但扩展名为.png
        if not output_path:
            base_name = os.path.splitext(dxf_path)[0]
            output_path = f"{base_name}.png"
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
        
        # 创建临时SVG文件路径
        temp_svg = f"{dxf_path}.temp.svg"
        temp_png = f"{dxf_path}.temp.png"
        
        import subprocess
        import shutil
        import platform
        
        # 检查是否安装了Inkscape
        try:
            subprocess.run(['which', 'inkscape'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            inkscape_installed = True
        except:
            inkscape_installed = False
            logger.warning("未找到Inkscape，将尝试其他方法")
        
        if inkscape_installed:
            logger.info("使用Inkscape转换DXF到SVG")
            
            # 创建命令
            # 在无头服务器上运行Inkscape需要Xvfb
            xvfb_run = 'xvfb-run -a ' if shutil.which('xvfb-run') else ''
            
            # Inkscape命令
            # 检查Inkscape版本，命令行参数格式不同
            inkscape_version = subprocess.run(['inkscape', '--version'], 
                                             stdout=subprocess.PIPE, 
                                             stderr=subprocess.PIPE, 
                                             text=True)
            
            if '1.' in inkscape_version.stdout:  # Inkscape 1.x版本
                # 使用新版命令行语法
                cmd_svg = f"{xvfb_run}inkscape --export-filename={temp_svg} {dxf_path}"
                cmd_png = f"{xvfb_run}inkscape --export-filename={temp_png} --export-background=black --export-dpi=300 {temp_svg}"
            else:  # Inkscape 0.9x版本
                # 使用旧版命令行语法
                cmd_svg = f"{xvfb_run}inkscape -f {dxf_path} -l {temp_svg}"
                cmd_png = f"{xvfb_run}inkscape -f {temp_svg} -e {temp_png} -b black -d 300"
            
            # 执行DXF到SVG的转换
            process = subprocess.run(cmd_svg, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if process.returncode != 0:
                logger.error(f"Inkscape转换失败: {process.stderr.decode()}")
                raise Exception("Inkscape转换失败")
            
            # 执行SVG到PNG的转换
            process = subprocess.run(cmd_png, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if process.returncode != 0:
                # 如果Inkscape转换PNG失败，尝试使用librsvg
                logger.warning("Inkscape转PNG失败，尝试使用librsvg")
                try:
                    # 检查是否安装了rsvg-convert
                    subprocess.run(['which', 'rsvg-convert'], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    cmd_rsvg = f"rsvg-convert -f png -o {temp_png} -b black -w 1200 {temp_svg}"
                    process = subprocess.run(cmd_rsvg, shell=True, check=True)
                except:
                    logger.error("librsvg未安装或转换失败")
                    raise Exception("librsvg转换失败")
            
            # 处理图像
            with Image.open(temp_png) as img:
                # 转换为RGBA
                img = img.convert('RGBA')
                
                # 获取图像数据
                data = np.array(img)
                
                # 将白色转换为亮蓝色
                white_areas = (data[:, :, 0] > 200) & (data[:, :, 1] > 200) & (data[:, :, 2] > 200)
                data[white_areas] = [0, 255, 255, 255]  # 亮蓝色
                
                # 创建新图像
                new_img = Image.fromarray(data)
                
                # 保存最终图像
                new_img.save(output_path, 'PNG')
            
            # 清理临时文件
            if os.path.exists(temp_svg):
                os.remove(temp_svg)
            if os.path.exists(temp_png):
                os.remove(temp_png)
            
            logger.info(f"成功使用Inkscape将DXF文件转换为图像: {output_path}")
            return output_path
        
        # 如果Inkscape未安装，回退到其他方法
        return None
    
    except Exception as e:
        logger.error(f"使用Inkscape转换DXF时出错: {e}")
        logger.error(traceback.format_exc())
        return None

def convert_dxf_to_image(dxf_path: str, output_path: str = None, layout_name: str = 'A3', dpi: int = 300) -> str:
    """
    将DXF文件转换为黑底图像
    
    Args:
        dxf_path: DXF文件路径
        output_path: 输出图像路径，如果不指定则自动生成
        layout_name: 布局名称，默认为'A3'
        dpi: 图像分辨率，默认300dpi
        
    Returns:
        str: 生成的图像文件路径
    """
    # 尝试不同的方法来转换DXF文件
    methods = [
        # 第一种方法：使用Inkscape（适合Linux无头服务器）
        lambda: convert_dxf_using_inkscape(dxf_path, output_path),
        
        # 第二种方法：使用外部工具（如果可用）
        lambda: convert_dxf_using_external_tool(dxf_path, output_path),
        
        # 第三种方法：使用直接渲染的方式
        lambda: convert_dxf_to_image_direct(dxf_path, output_path),
        
        # 第四种方法：使用原来的方法（最后尝试）
        lambda: _convert_dxf_to_image_original(dxf_path, output_path, layout_name, dpi)
    ]
    
    for method in methods:
        try:
            result = method()
            if result and os.path.exists(result):
                return result
        except Exception as e:
            logger.warning(f"转换方法失败: {e}")
            continue
    
    logger.error("所有转换方法都失败")
    return None

# 重命名原来的函数为_convert_dxf_to_image_original
_convert_dxf_to_image_original = convert_dxf_to_image_direct

if __name__ == "__main__":
    # 测试代码
    if len(sys.argv) > 1:
        dxf_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        try:
            result = convert_dxf_to_image(dxf_file, output_file)
            if result:
                print(f"转换成功: {result}")
            else:
                print("转换失败")
        except Exception as e:
            print(f"转换过程出错: {e}")
    else:
        print("使用方法: python dxf_to_image.py <dxf文件路径> [输出图像路径]") 