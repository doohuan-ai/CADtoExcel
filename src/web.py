#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import uuid
import logging
import threading
import time
import zipfile
import glob
import json
import shutil
import traceback
from typing import Dict, List, Any, Tuple, Optional
from flask import Flask, render_template, request, send_file, redirect, url_for, flash, jsonify, session, send_from_directory
from werkzeug.utils import secure_filename

# 设置项目根目录
ROOT_DIR = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
sys.path.insert(0, ROOT_DIR)

# 确保日志目录存在
logs_dir = os.path.join(ROOT_DIR, 'logs')
os.makedirs(logs_dir, exist_ok=True)

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(logs_dir, 'web_app.log'), encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)

# 创建Flask应用
app = Flask(__name__, template_folder='../web/templates', static_folder='../web/static')
app.secret_key = 'CADtoExcel-secret-key'
app.config['SESSION_TYPE'] = 'filesystem'

# 配置文件上传
UPLOAD_FOLDER = os.path.join(ROOT_DIR, 'uploads')
ALLOWED_EXTENSIONS_DWG = {'dwg'}
ALLOWED_EXTENSIONS_EXCEL = {'xls', 'xlsx'}
OUTPUT_FOLDER = os.path.join(ROOT_DIR, 'outputs')
TEMP_FOLDER = os.path.join(ROOT_DIR, 'outputs', 'temp')  # 临时存储批处理任务文件

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 限制上传文件大小为100MB

# 存储批处理任务状态
batch_jobs = {}

# 导入自定义模块
from src.excel_parser import parse_excel_file, save_excel_data_to_json
from src.dwg_parser import parse_dwg_file, save_dwg_data_to_json, convert_dwg_to_dxf
from src.report_generator import generate_template_report
from src.extract_excel_cell import extract_product_info_direct

def allowed_file_dwg(filename):
    """检查是否是允许的DWG文件扩展名"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_DWG

def allowed_file_excel(filename):
    """检查是否是允许的Excel文件扩展名"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_EXCEL

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/product')
def product_intro():
    """产品介绍页面"""
    return render_template('product_intro.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理单个文件上传"""
    # 检查是否有文件
    if 'dwg_file' not in request.files or 'excel_file' not in request.files:
        flash('没有选择文件')
        return redirect(request.url)
    
    dwg_file = request.files['dwg_file']
    excel_file = request.files['excel_file']
    
    # 如果用户未选择文件
    if dwg_file.filename == '' or excel_file.filename == '':
        flash('没有选择文件')
        return redirect(request.url)
    
    # 创建一个唯一的会话ID
    session_id = str(uuid.uuid4())
    session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    os.makedirs(session_folder, exist_ok=True)
    
    if dwg_file and allowed_file_dwg(dwg_file.filename) and excel_file and allowed_file_excel(excel_file.filename):
        # 安全地获取文件名
        dwg_filename = secure_filename(dwg_file.filename)
        excel_filename = secure_filename(excel_file.filename)
        
        # 保存上传的文件
        dwg_path = os.path.join(session_folder, dwg_filename)
        excel_path = os.path.join(session_folder, excel_filename)
        
        dwg_file.save(dwg_path)
        excel_file.save(excel_path)
        
        logger.info(f"文件保存成功: {dwg_path} 和 {excel_path}")
        
        try:
            # 处理文件
            dwg_json_path, excel_json_path, report_path = process_files(dwg_path, excel_path)
            
            # 获取结果文件名并传递给模板
            dwg_json_name = os.path.basename(dwg_json_path)
            excel_json_name = os.path.basename(excel_json_path)
            report_name = os.path.basename(report_path)
            
            logger.info(f"生成的文件: {dwg_json_name}, {excel_json_name}, {report_name}")
            
            # 渲染结果页面
            return render_template(
                'result.html',
                success=True,
                message="文件处理成功！已生成检验报告。",
                dwg_file=dwg_filename,
                excel_file=excel_filename, 
                dwg_json=dwg_json_name, 
                excel_json=excel_json_name, 
                report_file=report_name
            )
        except Exception as e:
            logger.error(f"处理文件时出错: {e}")
            logger.error(traceback.format_exc())
            flash(f"处理文件时出错: {str(e)}")
            return redirect(url_for('index'))
    else:
        flash('不支持的文件类型，仅支持.dwg和.xls/.xlsx')
        return redirect(url_for('index'))

@app.route('/batch')
def batch_page():
    """批量处理页面"""
    # 获取正在进行的批处理任务
    active_jobs = {job_id: job for job_id, job in batch_jobs.items() if job['status'] != 'completed'}
    completed_jobs = {job_id: job for job_id, job in batch_jobs.items() if job['status'] == 'completed'}
    
    return render_template('batch.html', active_jobs=active_jobs, completed_jobs=completed_jobs)

@app.route('/batch/upload', methods=['POST'])
def batch_upload():
    """处理批量文件上传"""
    # 检查是否有文件
    if 'dwg_files' not in request.files or 'excel_files' not in request.files:
        flash('没有选择文件')
        return redirect(url_for('batch_page'))
    
    dwg_files = request.files.getlist('dwg_files')
    excel_files = request.files.getlist('excel_files')
    
    # 如果用户未选择文件
    if not dwg_files or not excel_files:
        flash('请同时上传DWG文件和Excel文件')
        return redirect(url_for('batch_page'))
    
    # 创建一个唯一的批处理作业ID
    job_id = str(uuid.uuid4())
    job_folder = os.path.join(app.config['TEMP_FOLDER'], job_id)
    dwg_folder = os.path.join(job_folder, 'dwg')
    excel_folder = os.path.join(job_folder, 'excel')
    
    os.makedirs(dwg_folder, exist_ok=True)
    os.makedirs(excel_folder, exist_ok=True)
    
    # 保存DWG文件
    dwg_filenames = []
    for dwg_file in dwg_files:
        if dwg_file and allowed_file_dwg(dwg_file.filename):
            filename = secure_filename(dwg_file.filename)
            file_path = os.path.join(dwg_folder, filename)
            dwg_file.save(file_path)
            dwg_filenames.append(filename)
    
    # 保存Excel文件
    excel_filenames = []
    for excel_file in excel_files:
        if excel_file and allowed_file_excel(excel_file.filename):
            filename = secure_filename(excel_file.filename)
            file_path = os.path.join(excel_folder, filename)
            excel_file.save(file_path)
            excel_filenames.append(filename)
    
    if not dwg_filenames or not excel_filenames:
        flash('没有有效的文件被上传')
        return redirect(url_for('batch_page'))
    
    # 创建批处理作业
    batch_jobs[job_id] = {
        'id': job_id,
        'status': 'pending',
        'start_time': time.time(),
        'end_time': None,
        'total': 0,
        'processed': 0,
        'success': 0,
        'error': 0,
        'logs': []
    }
    
    # 启动后台线程处理批处理作业
    thread = threading.Thread(target=process_batch_job, args=(job_id,))
    thread.daemon = True
    thread.start()
    
    return redirect(url_for('batch_status', job_id=job_id))

def process_files(dwg_file: str, excel_file: str) -> Tuple[str, str, str]:
    """
    处理DWG和Excel文件，生成报告
    
    Args:
        dwg_file: DWG文件路径
        excel_file: Excel文件路径
        
    Returns:
        Tuple[str, str, str]: 包含DWG JSON路径、Excel JSON路径和报告路径的元组
    """
    logger.info(f"开始处理文件: {dwg_file} 和 {excel_file}")
    
    # 获取基础文件名，用于生成输出文件
    base_name = os.path.splitext(os.path.basename(dwg_file))[0]
    
    # 确保输出目录存在
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'])
    os.makedirs(output_dir, exist_ok=True)
    
    # 解析DWG文件并生成JSON
    logger.info(f"解析DWG文件: {dwg_file}")
    try:
        dwg_data, dwg_raw_data = parse_dwg_file(dwg_file)
        dwg_json_path = os.path.join(output_dir, f"{base_name}_dwg.json")
        dwg_raw_json_path = os.path.join(output_dir, f"{base_name}_dwg_raw_data.json")
        
        # 保存DWG数据
        with open(dwg_json_path, 'w', encoding='utf-8') as f:
            json.dump(dwg_data, f, ensure_ascii=False, indent=2)
        
        # 保存原始DWG数据 - 如果解析成功
        if dwg_raw_data:
            with open(dwg_raw_json_path, 'w', encoding='utf-8') as f:
                json.dump(dwg_raw_data, f, ensure_ascii=False, indent=2)
        else:
            # 如果原始数据为空，创建一个简化的数据结构
            logger.warning("原始DWG数据为空，将创建简化的数据结构")
            with open(dwg_raw_json_path, 'w', encoding='utf-8') as f:
                json.dump({"source": "raw_data_not_available", "parsed_data": dwg_data}, f, ensure_ascii=False, indent=2)
        
        logger.info(f"DWG JSON数据已保存: {dwg_json_path} 和 {dwg_raw_json_path}")
    except Exception as e:
        logger.error(f"解析DWG文件时出错: {e}")
        logger.error(traceback.format_exc())
        raise
    
    # 解析Excel文件并生成JSON
    logger.info(f"解析Excel文件: {excel_file}")
    try:
        excel_data = parse_excel_file(excel_file)
        excel_json_path = os.path.join(output_dir, f"{base_name}_excel.json")
        
        # 保存Excel数据
        with open(excel_json_path, 'w', encoding='utf-8') as f:
            json.dump(excel_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"Excel JSON数据已保存: {excel_json_path}")
    except Exception as e:
        logger.error(f"解析Excel文件时出错: {e}")
        logger.error(traceback.format_exc())
        raise
    
    # 生成报告
    logger.info(f"生成检验报告")
    try:
        # 获取外观要求对照表
        appearance_map_file = os.path.join(ROOT_DIR, 'maps', 'appearance_map.json')
        template_file = os.path.join(ROOT_DIR, 'maps', 'report_map.xlsx')
        
        # 生成报告
        report_path = generate_template_report(
            dwg_json_path, 
            excel_json_path, 
            appearance_map_file, 
            template_file,
            output_dir,
            original_excel_file=excel_file
        )
        
        logger.info(f"报告已生成: {report_path}")
    except Exception as e:
        logger.error(f"生成报告时出错: {e}")
        logger.error(traceback.format_exc())
        raise
    
    return dwg_json_path, excel_json_path, report_path

def create_zip_file(folder_path, file_list, zip_path):
    """
    创建包含指定文件的ZIP文件
    
    Args:
        folder_path: 文件夹路径，如果为None则file_list包含完整的文件路径
        file_list: 要包含的文件列表
        zip_path: 输出ZIP文件路径
    """
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file in file_list:
            if folder_path:
                # 如果提供了文件夹路径，则file是相对路径
                file_path = os.path.join(folder_path, file)
                arc_name = file
            else:
                # 否则file是完整路径
                file_path = file
                arc_name = os.path.basename(file)
            
            try:
                zipf.write(file_path, arc_name)
                logger.info(f"已添加文件到ZIP: {file_path}")
            except Exception as e:
                logger.error(f"添加文件到ZIP时出错: {file_path}, 错误: {e}")

def process_batch_job(job_id):
    """
    处理批处理作业
    
    Args:
        job_id: 作业ID
    """
    if job_id not in batch_jobs:
        logger.error(f"无效的作业ID: {job_id}")
        return
    
    job = batch_jobs[job_id]
    job['status'] = 'processing'
    job['processed'] = 0
    job['success'] = 0
    job['error'] = 0
    job['start_time'] = time.time()
    job['logs'] = []
    
    # 创建输出目录
    job_output_dir = os.path.join(app.config['OUTPUT_FOLDER'], f"batch_{job_id}")
    os.makedirs(job_output_dir, exist_ok=True)
    job['output_folder'] = job_output_dir
    
    def log_message(message, level='info'):
        """记录日志消息"""
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        log_entry = f"[{timestamp}] {message}"
        job['logs'].append(log_entry)
        if level == 'info':
            logger.info(message)
        elif level == 'error':
            logger.error(message)
        elif level == 'warning':
            logger.warning(message)
    
    try:
        # 获取文件路径
        dwg_folder = os.path.join(app.config['TEMP_FOLDER'], job_id, 'dwg')
        excel_folder = os.path.join(app.config['TEMP_FOLDER'], job_id, 'excel')
        output_folder = job['output_folder']
        
        # 获取所有DWG和Excel文件
        dwg_files = glob.glob(os.path.join(dwg_folder, '*.dwg'))
        excel_files = glob.glob(os.path.join(excel_folder, '*.xls')) + glob.glob(os.path.join(excel_folder, '*.xlsx'))
        
        log_message(f"找到 {len(dwg_files)} 个DWG文件和 {len(excel_files)} 个Excel文件")
        
        # 为每对文件找到匹配项
        file_pairs = []
        report_files = []
        
        # 按文件名前缀匹配文件
        for dwg_file in dwg_files:
            dwg_basename = os.path.splitext(os.path.basename(dwg_file))[0]
            matched = False
            
            # 查找匹配的Excel文件
            for excel_file in excel_files:
                excel_basename = os.path.splitext(os.path.basename(excel_file))[0]
                
                # 检查是否匹配
                if dwg_basename == excel_basename or dwg_basename in excel_basename or excel_basename in dwg_basename:
                    file_pairs.append((dwg_file, excel_file))
                    matched = True
                    log_message(f"匹配到文件对: {os.path.basename(dwg_file)} 和 {os.path.basename(excel_file)}")
                    break
            
            if not matched:
                log_message(f"DWG文件 {os.path.basename(dwg_file)} 未找到匹配的Excel文件", level='warning')
        
        # 处理每对文件
        job['total'] = len(file_pairs)
        
        for dwg_file, excel_file in file_pairs:
            try:
                # 增加处理计数
                job['processed'] += 1
                
                # 生成报告
                log_message(f"处理文件对: {os.path.basename(dwg_file)} 和 {os.path.basename(excel_file)}")
                dwg_json_path, excel_json_path, report_path = process_files(dwg_file, excel_file)
                
                # 将报告路径添加到列表
                report_files.append(report_path)
                
                # 增加成功计数
                job['success'] += 1
                log_message(f"成功生成报告: {os.path.basename(report_path)}")
                
            except Exception as e:
                # 增加错误计数
                job['error'] += 1
                error_message = f"处理文件对 {os.path.basename(dwg_file)} 和 {os.path.basename(excel_file)} 时出错: {str(e)}"
                log_message(error_message, level='error')
                logger.error(traceback.format_exc())
        
        # 创建报告包
        if report_files:
            # 创建ZIP包
            zip_file_path = os.path.join(output_folder, f"batch_{job_id}_reports.zip")
            create_zip_file(None, report_files, zip_file_path)
            job['result_zip'] = zip_file_path
            log_message(f"已创建报告包: {os.path.basename(zip_file_path)}")
        
        # 处理完成
        job['status'] = 'completed'
        job['end_time'] = time.time()
        job['duration'] = job['end_time'] - job['start_time']
        log_message(f"批处理作业已完成，总计: {job['total']}，成功: {job['success']}，失败: {job['error']}，耗时: {job['duration']:.2f}秒")
        
        # 清理临时目录
        shutil.rmtree(os.path.join(app.config['TEMP_FOLDER'], job_id), ignore_errors=True)
        
    except Exception as e:
        # 作业失败
        job['status'] = 'failed'
        job['end_time'] = time.time()
        job['duration'] = job['end_time'] - job['start_time']
        error_message = f"批处理作业失败: {str(e)}"
        log_message(error_message, level='error')
        logger.error(traceback.format_exc())

@app.route('/batch/status/<job_id>')
def batch_status(job_id):
    """批处理状态页面"""
    # 检查作业是否存在
    if job_id not in batch_jobs:
        flash('无效的作业ID')
        return redirect(url_for('batch_page'))
    
    # 获取作业信息
    job = batch_jobs[job_id]
    
    return render_template('batch_status.html', job=job, job_id=job_id)

@app.route('/batch/status/<job_id>/json')
def batch_status_json(job_id):
    """获取批处理作业状态的JSON数据"""
    if job_id not in batch_jobs:
        return jsonify({'error': '无效的作业ID'})
    
    return jsonify(batch_jobs[job_id])

@app.route('/batch/download/<job_id>')
def batch_download(job_id):
    """下载批处理作业结果ZIP包"""
    # 检查作业是否存在且完成
    if job_id not in batch_jobs or batch_jobs[job_id]['status'] != 'completed':
        flash('作业未完成或不存在')
        return redirect(url_for('batch_page'))
    
    job = batch_jobs[job_id]
    
    # 检查ZIP文件是否存在
    if 'result_zip' not in job or not os.path.exists(job['result_zip']):
        flash('结果文件不存在')
        return redirect(url_for('batch_status', job_id=job_id))
    
    return send_file(job['result_zip'], as_attachment=True)

@app.route('/download/<filename>')
def download_file(filename):
    """下载单个报告文件"""
    try:
        return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename),
                         as_attachment=True,
                         download_name=filename)
    except Exception as e:
        flash(f'下载文件时发生错误: {e}')
        return redirect(url_for('index'))

@app.route('/index')
def index_alias():
    """主页别名"""
    return redirect(url_for('index'))

@app.route('/convert')
def convert_page():
    """DWG转DXF页面"""
    return render_template('convert.html')

@app.route('/convert/upload', methods=['POST'])
def convert_upload():
    """处理DWG转DXF的文件上传"""
    # 检查是否有文件
    if 'dwg_file' not in request.files:
        flash('没有选择DWG文件')
        return redirect(url_for('convert_page'))
    
    dwg_file = request.files['dwg_file']
    
    # 如果用户未选择文件
    if dwg_file.filename == '':
        flash('没有选择DWG文件')
        return redirect(url_for('convert_page'))
    
    # 检查文件类型
    if not allowed_file_dwg(dwg_file.filename):
        flash('请上传.dwg格式的文件')
        return redirect(url_for('convert_page'))
    
    # 创建一个唯一的会话ID
    session_id = str(uuid.uuid4())
    session_folder = os.path.join(app.config['UPLOAD_FOLDER'], session_id)
    os.makedirs(session_folder, exist_ok=True)
    
    # 安全地保存文件
    dwg_filename = secure_filename(dwg_file.filename)
    dwg_file_path = os.path.join(session_folder, dwg_filename)
    dwg_file.save(dwg_file_path)
    logger.info(f"DWG文件已保存: {dwg_file_path}")
    
    try:
        # 转换DWG为DXF
        base_name = os.path.splitext(dwg_filename)[0]
        output_folder = os.path.join(app.config['OUTPUT_FOLDER'], session_id)
        os.makedirs(output_folder, exist_ok=True)
        
        dxf_file_path = os.path.join(output_folder, f"{base_name}.dxf")
        result_path = convert_dwg_to_dxf(dwg_file_path, dxf_file_path)
        
        if not result_path or not os.path.exists(result_path):
            logger.error(f"转换失败: {dwg_file_path}")
            flash('DWG转DXF失败，请检查文件格式或重试')
            return redirect(url_for('convert_page'))
        
        logger.info(f"DWG已成功转换为DXF: {result_path}")
        
        # 返回结果信息
        result = {
            'success': True,
            'message': 'DWG文件已成功转换为DXF格式',
            'dxf_file': os.path.basename(result_path),
            'session_id': session_id
        }
        
        # 将结果存储在会话中
        session['conversion_result'] = result
        
        # 重定向到结果页面
        return redirect(url_for('convert_result'))
        
    except Exception as e:
        logger.error(f"处理DWG转DXF时出错: {str(e)}")
        logger.error(traceback.format_exc())
        
        # 特别处理试用次数用完的情况
        error_message = str(e)
        if "试用次数已用完" in error_message:
            flash('本软件试用次数已用完，请联系供应商购买完整版')
        else:
            flash(f'处理失败: {error_message}')
        
        return redirect(url_for('convert_page'))

@app.route('/convert/result')
def convert_result():
    """显示DWG转DXF的结果页面"""
    # 从会话中获取结果
    result = session.get('conversion_result', None)
    if not result:
        flash('没有找到转换结果')
        return redirect(url_for('convert_page'))
    
    return render_template('convert_result.html', result=result)

@app.route('/convert/download/<session_id>/<filename>')
def download_dxf(session_id, filename):
    """下载转换后的DXF文件"""
    output_folder = os.path.join(app.config['OUTPUT_FOLDER'], session_id)
    if not os.path.exists(output_folder) or not os.path.exists(os.path.join(output_folder, filename)):
        flash('文件不存在或已过期')
        return redirect(url_for('convert_page'))
    
    return send_from_directory(output_folder, filename, as_attachment=True, download_name=filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001) 