from flask import Flask, render_template, request, jsonify, send_file
import os
import sys
import json
import threading
import time
import logging
from datetime import datetime
from werkzeug.utils import secure_filename
from digikey import DigiKeyClient
from write_excel import read_excel_data, write_excel_data, write_multiple_columns

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# 配置日志
log_dir = 'logs'
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f'digikey_{datetime.now().strftime("%Y%m%d")}.log')

# 创建日志记录器
logger = logging.getLogger('digikey_app')
logger.setLevel(logging.INFO)

# 创建文件处理器
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.INFO)

# 创建控制台处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# 创建格式化器
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# 添加处理器到日志记录器
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# 记录应用启动日志
logger.info("Digi-Key 产品状态查询工具启动")

# 确保上传目录存在
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# 全局变量存储处理状态
processing_status = {
    'is_processing': False,
    'progress': 0,
    'current_product': '',
    'total_products': 0,
    'message': '',
    'results': {}
}

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xls'}

@app.route('/')
def index():
    """渲染主页"""
    return render_template('index.html')

@app.route('/test')
def test():
    """渲染测试页面"""
    return render_template('test.html')

@app.route('/full_test')
def full_test():
    """渲染完整流程测试页面"""
    return render_template('full_test.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传"""
    logger.info("收到文件上传请求")
    
    if 'file' not in request.files:
        logger.warning("文件上传请求中没有文件部分")
        return jsonify({'status': 'error', 'message': '没有文件部分'})
    
    file = request.files['file']
    if file.filename == '':
        logger.warning("文件上传请求中没有选择文件")
        return jsonify({'status': 'error', 'message': '没有选择文件'})
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        logger.info(f"保存上传文件: {filename} 到 {filepath}")
        file.save(filepath)
        
        logger.info(f"文件 {filename} 上传成功")
        return jsonify({
            'status': 'success', 
            'message': '文件上传成功',
            'filename': filename
        })
    
    logger.warning(f"不支持的文件类型: {file.filename}")
    return jsonify({'status': 'error', 'message': '不支持的文件类型'})

@app.route('/start_processing', methods=['POST'])
def start_processing():
    """开始处理产品数据"""
    global processing_status
    
    logger.info("收到开始处理请求")
    
    if processing_status['is_processing']:
        logger.warning("已有任务正在处理中，拒绝新请求")
        return jsonify({'status': 'error', 'message': '已有任务正在处理中'})
    
    data = request.json
    filename = data.get('filename')
    sheet_name = data.get('sheet_name')
    column_name = data.get('column_name')
    result_column_name = data.get('result_column_name')
    selected_fields = data.get('selected_fields', [])
    custom_headers = data.get('custom_headers', {})
    
    logger.info(f"处理参数 - 文件: {filename}, 工作表: {sheet_name}, 列名: {column_name}, 结果列名: {result_column_name}")
    logger.info(f"选择的数据字段: {selected_fields}")
    logger.info(f"自定义表头: {custom_headers}")
    
    if not all([filename, sheet_name, column_name]):
        logger.warning("开始处理请求参数不完整")
        return jsonify({'status': 'error', 'message': '参数不完整'})
    
    # 在新线程中运行处理任务
    thread = threading.Thread(
        target=process_products_task,
        args=(filename, sheet_name, column_name, result_column_name, selected_fields, custom_headers)
    )
    thread.daemon = True
    thread.start()
    
    logger.info("处理任务已启动")
    return jsonify({'status': 'success', 'message': '处理任务已启动'})

def process_products_task(filename, sheet_name, column_name, result_column_name=None, selected_fields=None, custom_headers=None):
    """处理产品数据的任务函数"""
    global processing_status
    
    # 如果没有提供选择字段，则默认选择所有字段
    if selected_fields is None:
        selected_fields = ['status', 'description', 'manufacturer', 'product_url', 'datasheet_url', 'quantity_available']
    
    # 如果没有提供自定义表头，则使用默认表头
    if custom_headers is None:
        custom_headers = {}
    
    logger.info(f"开始处理产品数据，文件: {filename}, 工作表: {sheet_name}, 列名: {column_name}, 结果列名: {result_column_name}")
    logger.info(f"选择的数据字段: {selected_fields}")
    logger.info(f"自定义表头: {custom_headers}")
    
    try:
        processing_status['is_processing'] = True
        processing_status['progress'] = 0
        processing_status['message'] = '正在读取产品数据...'
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        client = DigiKeyClient()
        # 读取数据并获取表头行号
        data, header_row, header_column = read_excel_data(filepath, sheet_name, column_name, return_header_info=True)
        
        if not data:
            logger.warning("未获取到有效的产品数据")
            processing_status['message'] = '未获取到有效的产品数据'
            processing_status['is_processing'] = False
            return
        
        results = {}
        total = len(data)
        logger.info(f"共读取到 {total} 个产品数据")
        processing_status['total_products'] = total
        processing_status['message'] = f'开始处理 {total} 个产品...'
        
        success_count = 0
        failure_count = 0
        
        for i, product_number in enumerate(data, 1):
            processing_status['current_product'] = product_number
            processing_status['progress'] = i / total * 100
            
            logger.info(f"正在处理第 {i}/{total} 个产品: {product_number}")
            
            details = client.get_product_details(product_number)
            if isinstance(details, dict) and details.get('Product'):
                # 获取完整的产品信息
                product = details.get('Product', {})
                product_status = product.get('ProductStatus', {}).get('Status')
                product_description = product.get('Description', {}).get('ProductDescription', '')
                manufacturer = product.get('Manufacturer', {}).get('Name', '')
                product_url = product.get('ProductUrl', '')
                datasheet_url = product.get('DatasheetUrl', '')
                quantity_available = product.get('QuantityAvailable', 0)
                
                if product_status:
                    results[product_number] = {
                        'status': product_status,
                        'description': product_description,
                        'manufacturer': manufacturer,
                        'product_url': product_url,
                        'datasheet_url': datasheet_url,
                        'quantity_available': quantity_available
                    }
                    success_count += 1
                    logger.info(f"产品 {product_number} 状态查询成功: {product_status}")
                else:
                    results[product_number] = {
                        'status': "查询失败: 未找到状态信息",
                        'description': '',
                        'manufacturer': '',
                        'product_url': '',
                        'datasheet_url': '',
                        'quantity_available': 0
                    }
                    failure_count += 1
                    logger.warning(f"产品 {product_number} 未找到状态信息")
            else:
                results[product_number] = {
                    'status': f"查询失败: {details if isinstance(details, str) else '未知错误'}",
                    'description': '',
                    'manufacturer': '',
                    'product_url': '',
                    'datasheet_url': '',
                    'quantity_available': 0
                }
                failure_count += 1
                logger.error(f"产品 {product_number} 查询失败: {details if isinstance(details, str) else '未知错误'}")
        
        logger.info(f"产品处理完成，成功: {success_count}, 失败: {failure_count}")
        processing_status['message'] = '产品处理完成！正在保存结果...'
        
        # 保存结果到JSON文件
        data_file = os.path.join(os.path.dirname(__file__), 'product_details.json')
        with open(data_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        logger.info(f"结果已保存到JSON文件: {data_file}")
        
        # 写入Excel，使用用户指定的列名或自动生成
        output_column = result_column_name if result_column_name else f"{column_name}_状态"
        
        # 准备多列数据，只包含用户选择的字段
        columns_data = {}
        
        # 字段映射
        field_mapping = {
            'status': '状态',
            'description': '描述',
            'manufacturer': '制造商',
            'product_url': '产品链接',
            'datasheet_url': '数据手册',
            'quantity_available': '可用数量'
        }
        
        # 为每个选择的字段创建列
        for field in selected_fields:
            if field in field_mapping:
                # 使用自定义表头或默认表头
                header = custom_headers.get(field, f"{output_column}_{field_mapping[field]}")
                columns_data[header] = [results.get(p, {}).get(field, '') for p in data]
        
        write_result = write_multiple_columns(filepath, sheet_name, columns_data, max_search_rows=10, reference_header=column_name, reference_header_row=header_row)
        
        if isinstance(write_result, dict) and write_result.get('status') == 'error':
            error_msg = f'写入Excel失败: {write_result.get("message")}'
            logger.error(error_msg)
            processing_status['message'] = error_msg
        else:
            logger.info("数据已成功写入Excel文件")
            processing_status['message'] = f'数据已成功写入Excel文件'
        
        processing_status['results'] = results
        processing_status['is_processing'] = False
        
    except Exception as e:
        error_msg = f'处理过程中发生错误: {str(e)}'
        logger.error(error_msg, exc_info=True)
        processing_status['message'] = error_msg
        processing_status['is_processing'] = False

@app.route('/processing_status')
def get_processing_status():
    """获取处理状态"""
    logger.debug("请求获取处理状态")
    return jsonify(processing_status)

@app.route('/download_result')
def download_result():
    """下载处理后的Excel文件"""
    filename = request.args.get('filename')
    if not filename:
        logger.warning("下载结果请求缺少文件名参数")
        return jsonify({'status': 'error', 'message': '参数不完整'})
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        logger.warning(f"下载的文件不存在: {filepath}")
        return jsonify({'status': 'error', 'message': '文件不存在'})
    
    logger.info(f"下载Excel文件: {filename}")
    return send_file(filepath, as_attachment=True)

@app.route('/download_json')
def download_json():
    """下载JSON结果文件"""
    data_file = os.path.join(os.path.dirname(__file__), 'product_details.json')
    if not os.path.exists(data_file):
        logger.warning("下载的JSON结果文件不存在")
        return jsonify({'status': 'error', 'message': '结果文件不存在'})
    
    logger.info("下载JSON结果文件")
    return send_file(data_file, as_attachment=True, download_name='product_details.json')

if __name__ == '__main__':
    logger.info("启动Flask应用服务器")
    app.run(debug=True, host='0.0.0.0', port=5000)