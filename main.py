from digikey import DigiKeyClient
from write_excel import read_excel_data, write_multiple_columns
import json
import os
import sys
import logging
from datetime import datetime

# 创建logs目录（如果不存在）
if not os.path.exists('logs'):
    os.makedirs('logs')

# 设置日志文件路径（按日期）
log_file = os.path.join('logs', f'main_{datetime.now().strftime("%Y%m%d")}.log')

# 创建日志记录器
logger = logging.getLogger('main_processor')
logger.setLevel(logging.DEBUG)

# 创建文件处理器
file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)

# 创建控制台处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# 创建格式化器
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# 添加处理器到记录器
logger.addHandler(file_handler)
logger.addHandler(console_handler)

def process_products(excel_path, sheet_name, product_number_column, output_column):
    logger.info(f"开始处理产品数据: 文件={excel_path}, 工作表={sheet_name}, 产品编号列={product_number_column}, 输出列={output_column}")
    
    try:
        client = DigiKeyClient()
        logger.info("成功创建DigiKey客户端")
        
        data = read_excel_data(excel_path, sheet_name, product_number_column)
        if not data:
            error_msg = '未获取到有效的产品数据'
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}
            
        logger.info(f"成功读取到 {len(data)} 条产品数据")
        
        results = {}
        total = len(data)
        logger.info(f"开始处理 {total} 个产品...")
        
        success_count = 0
        failure_count = 0
        
        for i, product_number in enumerate(data, 1):
            progress = i / total * 100
            sys.stdout.write(f"\r处理进度: {i}/{total} ({progress:.1f}%) - 当前产品: {product_number}")
            sys.stdout.flush()
            
            logger.debug(f"正在处理产品 {i}/{total}: {product_number}")
            
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
                    logger.debug(f"成功获取产品状态: {product_number} -> {product_status}")
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
                logger.warning(f"产品 {product_number} 查询失败: {details if isinstance(details, str) else '未知错误'}")
                
        print("\n产品处理完成！")
        logger.info(f"产品处理完成！成功: {success_count}, 失败: {failure_count}")
        
        # 保存结果到JSON文件
        data_file = os.path.join(os.path.dirname(__file__), 'product_details.json')
        with open(data_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        logger.info(f"结果已保存到 {data_file}")
        
        # 准备多列数据
        columns_data = {
            output_column: [results.get(p, {}).get('status', '') for p in data],
            f"{output_column}_描述": [results.get(p, {}).get('description', '') for p in data],
            f"{output_column}_制造商": [results.get(p, {}).get('manufacturer', '') for p in data],
            f"{output_column}_产品链接": [results.get(p, {}).get('product_url', '') for p in data],
            f"{output_column}_数据手册": [results.get(p, {}).get('datasheet_url', '') for p in data],
            f"{output_column}_可用数量": [results.get(p, {}).get('quantity_available', 0) for p in data]
        }
        
        # 写入Excel（多列数据）
        write_result = write_multiple_columns(excel_path, sheet_name, columns_data)
        if isinstance(write_result, dict) and write_result.get('status') == 'error':
            logger.error(f"写入Excel失败: {write_result.get('message')}")
            return write_result
            
        logger.info(f"数据已成功写入Excel文件: {excel_path}")
        return {'status': 'success', 'message': f"成功处理 {len(results)} 个产品", 'data': results}
        
    except Exception as e:
        error_msg = f"处理过程中发生错误: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return {'status': 'error', 'message': error_msg}

if __name__ == "__main__":
    logger.info("启动主程序")
    
    try:
        file_path = input("请输入文件路径:")
        sheet_name = input("请输入工作表名:")
        product_number_column = input("请输入产品编号列名:")
        output_column = input("请输入输出列名:")

        logger.info(f"用户输入参数: 文件={file_path}, 工作表={sheet_name}, 产品编号列={product_number_column}, 输出列={output_column}")

        result = process_products(file_path, sheet_name, product_number_column, output_column)
        logger.info(f"处理结果: {result['status']}")
        logger.info(f"消息: {result['message']}")
        
        print(f"处理结果: {result['status']}")
        print(f"消息: {result['message']}")

    except Exception as e:
        error_msg = f"发生错误: {e}"
        logger.error(error_msg, exc_info=True)
        print(error_msg)
        
    logger.info("主程序结束")


