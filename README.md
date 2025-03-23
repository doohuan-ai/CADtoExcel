# CADtoExcel

CADtoExcel 是一个用于解析 CAD 文件（DWG格式）和 Excel 文件，并生成相应 JSON 格式数据以及外委工序检验报告的工具。

## 功能概述

- 解析固定的外观要求对照表（maps/appearances.xlsx）
- 解析任意 DWG 文件（CAD 图纸）
- 解析任意 Excel 文件（xls 或 xlsx 格式）
- 生成结构化的 JSON 输出
- 生成外委工序检验报告
- 提供Web界面，支持单文件和批量处理
- 支持DWG文件转换为DXF格式

## 环境要求

- Python 3.6+
- 依赖项已在 requirements.txt 中列出

## 安装

```bash
# 克隆仓库
git clone https://github.com/yourusername/CADtoExcel.git
cd CADtoExcel

# 创建虚拟环境
python -m venv env
source env/bin/activate  # 在Windows上使用: env\Scripts\activate

# 安装依赖
pip install -r requirements.txt
```

## 使用方法

### Web界面（推荐）

启动Web应用：

```bash
python app.py
```

然后在浏览器中访问：`http://localhost:5000`

Web界面提供三种处理方式：
1. **单文件处理** - 上传单个DWG文件和Excel文件进行处理
2. **批量处理** - 一次上传多个DWG文件和Excel文件，系统自动匹配并处理
3. **DWG转DXF** - 将AutoCAD的DWG文件格式转换为更通用的DXF格式

### 命令行调用

```bash
python src/cli.py --dwg <DWG文件路径> --excel <Excel文件路径> --output <输出目录>
```

例如：

```bash
python src/cli.py --dwg test/81206851-03.dwg --excel test/81206851-03.xls --output outputs
```

### 作为库调用

```python
from src.main import process_files

# 处理文件并获取输出路径
output_files = process_files('test/81206851-03.dwg', 'test/81206851-03.xls', 'outputs')
```

## 批量处理功能

批量处理功能可以一次性处理多个DWG和Excel文件对，特别适合于需要处理大量文件的场景。

### 工作流程：

1. 用户上传多个DWG文件和Excel文件
2. 系统根据文件名自动匹配相应的DWG和Excel文件对
3. 系统在后台处理每一对文件，生成各自的报告
4. 所有报告会打包成一个ZIP文件供用户下载

### 匹配规则：

系统会自动查找名称相同或相似的DWG和Excel文件进行匹配。例如：
- `81206851-03.dwg` 会与 `81206851-03.xlsx` 匹配
- `ABC部件.dwg` 会与 `ABC部件工艺卡.xlsx` 匹配（包含关系）

### 状态跟踪：

批量处理过程中，用户可以看到：
- 总任务进度
- 已成功处理的文件数量
- 处理失败的文件数量
- 详细的处理日志

## 输出文件

程序会在指定的输出目录生成以下文件：

1. `<basename>_dwg.json` - DWG 文件解析结果
2. `<basename>_dwg_raw_data.json` - DWG 文件解析的原始数据
3. `<basename>_excel.json` - Excel 文件解析结果（包含工序信息）
4. `<basename>_report.xlsx` - 外委工序检验报告Excel文件
5. `<basename>.dxf` - 从DWG转换的DXF文件（使用转换功能时）
6. 批量处理时会生成 `batch_<job_id>_reports.zip` 包含所有报告

## 目录结构

项目使用以下目录结构：

- `uploads/` - 存放用户上传的原始文件
- `outputs/` - 存放所有处理结果
  - `outputs/temp/` - 临时处理目录
- `maps/` - 存放配置文件和模板
  - `maps/appearance_map.json` - 外观要求对照表
  - `maps/thickness_map.json` - 板厚公差对照表
  - `maps/report_map.xlsx` - 报告模板文件

## 数据格式

### 外观要求对照表 (appearance_map.json)

```json
{
  "工序名1": [
    {"id": 1, "description": "外观要求描述1"},
    {"id": 2, "description": "外观要求描述2"}
  ],
  "工序名2": [
    {"id": 1, "description": "外观要求描述1"}
  ]
}
```

### DWG 数据 (<basename>_dwg.json)

```json
{
  "file_name": "文件名.dwg",
  "metadata": {
    "version": "DXF版本",
    "header_variables": {}
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
```

### Excel 数据 (<basename>_excel.json)

```json
{
  "file_name": "文件名.xls",
  "sheets": [
    {
      "name": "工作表名",
      "headers": ["列标题1", "列标题2"],
      "data": [
        {"列标题1": "值1", "列标题2": "值2"},
        {"列标题1": "值3", "列标题2": "值4"}
      ]
    }
  ]
}
```

## 开发者信息

- 作者：[您的姓名]
- 联系方式：[您的联系方式]

## 许可证

MIT

## DWG转DXF功能

系统提供了将AutoCAD的DWG文件转换为更通用的DXF格式的功能：

### Web界面操作流程

1. 在首页点击"转换DWG"按钮或直接访问`/convert`路径
2. 上传DWG文件（支持拖放上传）
3. 点击"转换为DXF"按钮
4. 等待转换完成后下载DXF文件

### 命令行使用

```bash
python src/cli.py --convert-dwg <DWG文件路径> --output <输出目录>
```

### 作为库调用

```python
from src.dwg_parser import convert_dwg_to_dxf

# 转换DWG文件为DXF
dxf_file_path = convert_dwg_to_dxf('test/drawing.dwg', 'outputs/drawing.dxf')
```

### 注意事项

- 转换过程依赖于LibreDWG工具集中的`dwg2dxf`命令
- 转换结果将保留7天，请及时下载
- 当前支持的最大文件大小为100MB