# AI智能CAD工艺数据解析系统

<div align="center">
  <img src="https://img.shields.io/badge/版本-1.0.0-blue.svg" alt="版本">
  <img src="https://img.shields.io/badge/开发-深圳市多焕智能科技-brightgreen.svg" alt="开发">
  <img src="https://img.shields.io/badge/客户-华阳通机电有限公司-orange.svg" alt="客户">
</div>

## 🚀 产品概述

**AI智能CAD工艺数据解析系统** 是一款基于人工智能技术的高级智能处理工具，专为工业制造企业打造。本系统利用先进的计算机视觉和数据分析算法，实现CAD图纸与工艺卡数据的智能化解析、匹配和报告生成，显著提升制造工艺流程的自动化水平和效率。

### 核心智能功能

- **智能识别技术**：采用先进的计算机视觉算法自动识别和解析DWG格式CAD图纸
- **智能匹配系统**：智能将CAD图纸与对应的Excel工艺数据进行自动匹配和关联
- **智能批处理引擎**：支持批量文件的智能化自动处理，无需人工干预
- **智能报告生成**：自动生成结构化的工序检验报告，符合工业标准要求
- **智能文件转换**：内置智能格式转换引擎，实现DWG到DXF的无损转换

## ✨ 主要特性

- **多维度数据解析**：智能解析固定的外观要求对照表和动态工艺卡数据
- **智能格式处理**：支持多种格式的智能处理，包括DWG、Excel(xls/xlsx)等
- **结构化输出**：生成标准化JSON数据，便于系统集成和二次开发
- **友好的Web界面**：直观的操作界面，支持拖放上传和进度实时显示
- **高效批量处理**：内置智能匹配算法，一次处理多个文件对
- **数据安全保障**：所有数据处理在本地完成，确保企业核心数据安全

## 💻 系统要求

- **操作系统**：Windows 10/11
- **运行环境**：Python 3.6+
- **依赖组件**：详见requirements.txt

## 🔧 安装与部署

```bash
# 克隆仓库
git clone https://github.com/yourusername/CADtoExcel.git
cd CADtoExcel

# 创建虚拟环境
python -m venv env
env\Scripts\activate  # Windows系统

# 安装依赖
pip install -r requirements.txt

# 启动应用
python app.py
```

## 📊 使用方法

### Web智能界面（推荐）

启动Web应用后，在浏览器中访问：`http://localhost:5000`

Web界面提供三种智能处理模式：
1. **单文件智能处理** - 上传单个DWG文件和Excel文件进行智能匹配和处理
2. **批量智能处理** - 一次上传多个DWG文件和Excel文件，系统自动匹配并智能处理
3. **智能格式转换** - 利用智能算法将AutoCAD的DWG文件转换为更通用的DXF格式

### 命令行调用

```bash
python src/cli.py --dwg <DWG文件路径> --excel <Excel文件路径> --output <输出目录>
```

## 🔄 智能处理流程

1. **数据导入**：上传CAD图纸和工艺卡Excel文件
2. **智能解析**：系统自动解析图纸中的实体信息和Excel中的工艺数据
3. **智能匹配**：通过智能算法匹配图纸和工艺数据的对应关系
4. **智能生成**：自动生成标准化的工序检验报告
5. **数据导出**：提供JSON格式的结构化数据和Excel格式的报告文件

## 📁 输出文件说明

系统会在指定的输出目录智能生成以下文件：

1. `<basename>_dwg.json` - DWG文件智能解析结果
2. `<basename>_excel.json` - Excel文件智能解析结果
3. `<basename>_report.xlsx` - 智能生成的外委工序检验报告
4. `<basename>.dxf` - 智能转换的DXF文件（使用转换功能时）
5. 批量处理报告打包 `batch_<job_id>_reports.zip`

## 📞 联系与支持

- **开发方**：深圳市多焕智能科技有限公司
- **联系电话**：18689628301
- **电子邮箱**：reef@doohuan.com

## 📄 许可说明

本软件为试用版，试用期为3天。如需完整版，请联系我们获取商业授权。