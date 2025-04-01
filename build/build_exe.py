#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
CADtoExcel 打包脚本
使用PyInstaller将应用打包为EXE文件
"""

import os
import sys
import shutil
import subprocess
import platform
from pathlib import Path

# 获取项目根目录
PROJECT_ROOT = Path(__file__).parent.parent.absolute()
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
SPEC_FILE = BUILD_DIR / "CADtoExcel.spec"

def clean_build():
    """清理旧的构建文件"""
    print("清理旧的构建文件...")
    
    # 删除dist文件夹
    if DIST_DIR.exists():
        shutil.rmtree(DIST_DIR)
    
    # 删除构建临时文件
    for path in BUILD_DIR.glob("*"):
        if path.name != "setup.iss" and path.name != "build_exe.py":
            if path.is_dir():
                shutil.rmtree(path)
            else:
                os.remove(path)
    
    print("清理完成。")

def build_exe():
    """使用PyInstaller构建EXE"""
    print("开始构建EXE...")
    
    # 根据操作系统确定分隔符
    # Windows用分号(;)，Linux/macOS用冒号(:)
    separator = ";" if platform.system() == "Windows" else ":"
    
    # 构建命令
    cmd = [
        "pyinstaller",
        "--name=CADtoExcel",
        "--onefile",
        f"--workpath={BUILD_DIR}",
        f"--distpath={DIST_DIR}",
        f"--add-data=maps{separator}maps",
        f"--add-data=fonts{separator}fonts",
        f"--add-data=web/templates{separator}web/templates",
        f"--add-data=web/static{separator}web/static",
        "--hidden-import=engineio.async_drivers.threading",
        "--hidden-import=engineio.async_drivers.eventlet",
        "--hidden-import=eventlet.hubs.poll",
        "--hidden-import=eventlet.hubs.selects",
        "--hidden-import=eventlet.hubs.epolls",
        "--clean",
        "--noconfirm",
        str(PROJECT_ROOT / "app.py")
    ]
    
    # 执行打包命令
    subprocess.run(cmd, check=True)
    
    print("EXE构建完成。")

def build_installer():
    """使用Inno Setup构建安装程序"""
    print("开始构建安装程序...")
    
    # 检查当前系统是否为Windows
    if platform.system() != "Windows":
        print("警告: 安装程序只能在Windows系统上构建。")
        print("EXE文件已生成，可以将其复制到Windows系统上使用Inno Setup构建安装程序。")
        return
    
    # 检查Inno Setup是否安装
    iscc_path = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
    if not os.path.exists(iscc_path):
        print("警告: 未找到Inno Setup，请安装后再运行此脚本。")
        print("您可以从 https://jrsoftware.org/isdl.php 下载Inno Setup。")
        return
    
    # 创建输出目录
    output_dir = BUILD_DIR / "output"
    output_dir.mkdir(exist_ok=True)
    
    # 构建安装程序
    iss_file = BUILD_DIR / "setup.iss"
    subprocess.run([iscc_path, str(iss_file)], check=True)
    
    print("安装程序构建完成。")

def main():
    """主函数"""
    print("===== CADtoExcel 打包脚本 =====")
    print(f"当前操作系统: {platform.system()}")
    
    # 切换到项目根目录
    os.chdir(PROJECT_ROOT)
    
    try:
        # 清理旧的构建文件
        clean_build()
        
        # 构建EXE
        build_exe()
        #
        # # 构建安装程序
        build_installer()
        
        print("\n===== 打包成功 =====")
        if platform.system() == "Windows":
            print(f"安装程序已保存到: {BUILD_DIR / 'output' / 'CADtoExcel_Trial_Setup.exe'}")
        else:
            print(f"EXE文件已保存到: {DIST_DIR / 'CADtoExcel.exe'}")
            print("注意: 在非Windows系统上构建的EXE只能在Windows系统上运行")
    except Exception as e:
        print(f"\n错误: {e}")
        print("\n===== 打包失败 =====")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main()) 