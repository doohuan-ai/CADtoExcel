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
import traceback

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


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

        # 创建渲染上下文
        fig = plt.figure(figsize=(16, 11.3))  # A3比例
        ax = fig.add_axes([0, 0, 1, 1])
        ctx = RenderContext(doc)
        ctx.set_current_layout(layout)
        logger.info(f"打印ctx: {ctx}")

        # 设置黑色背景
        ax.set_facecolor('black')
        fig.patch.set_facecolor('black')

        # 创建前端和后端
        out = MatplotlibBackend(ax)
        Frontend(ctx, out).draw_layout(layout, finalize=True)

        # 调整视图
        ax.set_axis_off()
        plt.margins(0, 0)

        # 保存为临时文件
        temp_path = output_path + '.temp.png'
        plt.savefig(temp_path, dpi=dpi, bbox_inches='tight', pad_inches=0,
                    facecolor='black', edgecolor='none')
        plt.close()

        # 使用Pillow处理图像
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

        # 删除临时文件
        if os.path.exists(temp_path):
            os.remove(temp_path)

        logger.info(f"成功将DXF文件转换为图像: {output_path}")
        return output_path

    except Exception as e:
        logger.error(f"转换DXF为图像时出错: {e}")
        logger.error(traceback.format_exc())
        return None


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