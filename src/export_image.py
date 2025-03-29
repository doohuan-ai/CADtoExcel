# -*- coding: utf-8 -*-
import win32com.client
import pythoncom
import time
import os
import logging
from dataclasses import dataclass
from tenacity import (
    retry,
    stop_after_attempt,
    wait_fixed,
    retry_if_result)


logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - [%(filename)s:%(lineno)d- %(message)s',
)


@dataclass
class CADConfig:
    printer_name: str = "PublishToWeb PNG (Transparent).pc3"  # 透明背景
    media_name: str = "FHD_(1920.00_x_1080.00_Pixels)"  # 打印尺寸
    style_sheet: str = "acad.ctb"  # 黑白：mochrome.ctb
    plot_rotation: int = 0
    plot_scale: int = 0
    plot_with_lineweights: bool = True  # 打印线宽
    center_plot: bool = True  # 居中打印
    plot_with_plottyles: bool = True  # 依照样式打印
    plot_hidden: bool = False  # 隐藏图纸空间对象


class CADExporter:
    def __init__(self, config: CADConfig):
        self.config = config
        self.acad = None
        self.doc = None

    def __enter__(self):
        self.connect()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def connect(self):
        """创建AutoCAD连接"""
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.acad.Visible = False
        except Exception as e:
            logging.error(f"AutoCAD连接失败: {str(e)}")
            raise RuntimeError(f"AutoCAD连接失败: {str(e)}") from e

    def open_document(self, file_path):
        """打开DWG文档"""
        if not os.path.exists(file_path):
            logging.error(f"文件不存在: {file_path}")
            raise FileNotFoundError(f"文件不存在: {file_path}")

        try:
            self.doc = self.acad.Documents.Open(file_path)
            time.sleep(2)  # 等待文件加载
            return self.doc
        except Exception as e:
            logging.error(f"文档设置失败: {str(e)}")
            raise RuntimeError(f"文档打开失败: {str(e)}") from e

    def setup_layout(self):
        """配置打印参数"""
        # layout = self.doc.layouts.item('A3')  # 连接layout对象“A3”
        layout = self.doc.ActiveLayout
        layout.ConfigName = self.config.printer_name
        layout.CanonicalMediaName = self.config.media_name
        layout.StyleSheet = self.config.style_sheet
        layout.PlotRotation = self.config.plot_rotation
        layout.StandardScale = self.config.plot_scale
        layout.PlotWithLineweights = self.config.plot_with_lineweights
        layout.CenterPlot = self.config.center_plot
        layout.PlotWithPlotStyles = self.config.plot_with_plottyles
        layout.PlotHidden = self.config.plot_hidden

        # 设置打印窗口
        # start = self._apoint(0, 0)
        # start = self._apoint(420, 297) # A3尺寸
        # layout.SetWindowToPlot(start, end)

        # 必要参数设置
        self.doc.SetVariable('BACKGROUNDPLOT', 0)


    def export_image(self, output_dir):
        """执行导出操作"""
        base_name = os.path.splitext(self.doc.Name)[0]
        output_path = os.path.join(output_dir, base_name + ".png")

        try:
            plot = self.doc.Plot
            plot.PlotToFile(output_path, self.config.printer_name)
            return output_path
        except Exception as e:
            logging.error(f"导出失败: {str(e)}")
            raise RuntimeError(f"导出失败: {str(e)}") from e

    def _apoint(self, x, y):
        """坐标点转换辅助方法"""
        return win32com.client.VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))

    def close(self):
        """安全关闭资源"""
        try:
            # 强制刷新AutoCAD命令队列
            if self.doc:
                self.doc.SendCommand('(command "_.QSAVE") ')  # 发送保存命令
                time.sleep(0.5)

            # 分阶段关闭（先文档后应用）
            if self.doc:
                self.doc.Close(False)  # 参数False不保存修改
                time.sleep(0.3)  # 等待文档关闭
            if self.acad:
                self.acad.Quit()
                time.sleep(0.5)  # 等待进程退出
        except Exception as e:
            logging.error(f"资源关闭异常: {str(e)}")

def retry_if_result_false(value):
    return value is False


@retry(stop=stop_after_attempt(2),
       wait=wait_fixed(5),
       retry=retry_if_result(retry_if_result_false),
       reraise=True)
def process_dwg(file_path, output_dir, config=CADConfig()):
    """单文件导出函数（供进程池调用）"""
    try:
        with CADExporter(config) as exporter:
            exporter.open_document(file_path)
            exporter.setup_layout()
            output_path = exporter.export_image(output_dir)
            logging.info(f"成功导出: {output_path}")
            return True
    except Exception as e:
        logging.error(f"处理失败 [{file_path}]: {str(e)}")
        return False