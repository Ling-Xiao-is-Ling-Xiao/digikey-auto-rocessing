import openpyxl
import logging
import os
from datetime import datetime

# 配置日志
log_dir = 'logs'
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f'digikey_{datetime.now().strftime("%Y%m%d")}.log')

# 创建日志记录器
logger = logging.getLogger('excel_handler')
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

def read_excel_data(excel_path, sheet_name, header_name):
    logger.info(f"开始读取Excel文件: {excel_path}, 工作表: {sheet_name}, 表头: {header_name}")
    
    try:
        # 加载Excel文件
        workbook = openpyxl.load_workbook(excel_path)
        logger.info(f"成功加载Excel文件: {excel_path}")
        
        # 检查工作表是否存在
        if sheet_name not in workbook.sheetnames:
            error_msg = f"工作表 '{sheet_name}' 不存在"
            logger.error(error_msg)
            raise Exception(error_msg)
            
        # 选择工作表
        sheet = workbook[sheet_name]
        logger.info(f"成功选择工作表: {sheet_name}")
        
        # 查找表头所在列
        header_column = None
        for cell in sheet[1]:  # 假设表头在第一行
            if cell.value == header_name:
                header_column = cell.column
                logger.info(f"找到表头 '{header_name}' 在第 {header_column} 列")
                break
                
        if not header_column:
            error_msg = f"表头 '{header_name}' 未找到"
            logger.error(error_msg)
            raise Exception(error_msg)
        
        # 从表头下方读取所有数据
        data = []
        for row in sheet.iter_rows(min_row=2, min_col=header_column, max_col=header_column):
            if row[0].value:
                data.append(str(row[0].value).strip())
        
        if not data:
            logger.warning("未找到有效数据")
            return []
        
        logger.info(f"成功读取 {len(data)} 条数据")
        return data
        
    except Exception as e:
        error_msg = f"读取Excel文件时发生错误: {str(e)}"
        logger.error(error_msg, exc_info=True)
        raise Exception(error_msg)


def write_excel_data(excel_path, sheet_name, header_name, data):
    logger.info(f"开始写入Excel文件: {excel_path}, 工作表: {sheet_name}, 表头: {header_name}, 数据量: {len(data) if data else 0}")
    
    try:
        if not data:
            error_msg = '数据列表不能为空'
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}
            
        workbook = openpyxl.load_workbook(excel_path)
        logger.info(f"成功加载Excel文件: {excel_path}")
        
        if sheet_name not in workbook.sheetnames:
            error_msg = f"工作表 '{sheet_name}' 不存在"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}
            
        sheet = workbook[sheet_name]
        logger.info(f"成功选择工作表: {sheet_name}")
        
        # 查找表头所在列
        header_column = None
        for cell in sheet[1]:
            if cell.value == header_name:
                header_column = cell.column
                logger.info(f"找到现有表头 '{header_name}' 在第 {header_column} 列")
                break
        
        if not header_column:
            header_column = sheet.max_column + 1
            sheet.cell(row=1, column=header_column, value=header_name)
            logger.info(f"创建新表头 '{header_name}' 在第 {header_column} 列")
        
        # 写入数据（不清空原有数据，直接覆盖/追加）
        for idx, value in enumerate(data, start=2):
            sheet.cell(row=idx, column=header_column, value=value)
        
        workbook.save(excel_path)
        logger.info(f"成功保存Excel文件: {excel_path}")
        
        success_msg = f"成功写入 {len(data)} 条数据到Excel文件"
        logger.info(success_msg)
        return {'status': 'success', 'message': success_msg}
        
    except Exception as e:
        error_msg = f"写入Excel文件时发生错误: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return {'status': 'error', 'message': error_msg}

def write_multiple_columns(excel_path, sheet_name, columns_data):
    """
    写入多列数据到Excel文件
    
    参数:
        excel_path: Excel文件路径
        sheet_name: 工作表名称
        columns_data: 字典，键为列名，值为数据列表
    """
    logger.info(f"开始写入多列数据到Excel文件: {excel_path}, 工作表: {sheet_name}, 列数: {len(columns_data) if columns_data else 0}")
    
    try:
        if not columns_data:
            error_msg = '列数据字典不能为空'
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}
            
        workbook = openpyxl.load_workbook(excel_path)
        logger.info(f"成功加载Excel文件: {excel_path}")
        
        if sheet_name not in workbook.sheetnames:
            error_msg = f"工作表 '{sheet_name}' 不存在"
            logger.error(error_msg)
            return {'status': 'error', 'message': error_msg}
            
        sheet = workbook[sheet_name]
        logger.info(f"成功选择工作表: {sheet_name}")
        
        # 处理每一列数据
        for header_name, data in columns_data.items():
            logger.info(f"处理列 '{header_name}', 数据量: {len(data) if data else 0}")
            
            # 查找表头所在列
            header_column = None
            for cell in sheet[1]:
                if cell.value == header_name:
                    header_column = cell.column
                    logger.info(f"找到现有表头 '{header_name}' 在第 {header_column} 列")
                    break
            
            if not header_column:
                header_column = sheet.max_column + 1
                sheet.cell(row=1, column=header_column, value=header_name)
                logger.info(f"创建新表头 '{header_name}' 在第 {header_column} 列")
            
            # 写入数据（从第2行开始）
            for idx, value in enumerate(data, start=2):
                sheet.cell(row=idx, column=header_column, value=value)
        
        workbook.save(excel_path)
        logger.info(f"成功保存Excel文件: {excel_path}")
        
        success_msg = f"成功写入 {len(columns_data)} 列数据到Excel文件"
        logger.info(success_msg)
        return {'status': 'success', 'message': success_msg}
        
    except Exception as e:
        error_msg = f"写入多列数据到Excel文件时发生错误: {str(e)}"
        logger.error(error_msg, exc_info=True)
        return {'status': 'error', 'message': error_msg}

if __name__ == '__main__':
    # 示例用法 - 读取数据
    excel_path = input("请输入Excel文件路径: ")
    sheet_name = input("请输入工作表名称: ")
    header_name = input("请输入表头名称: ")
    mode = input("请输入操作模式 (读取/写入/多列写入): ")
    if mode == '读取':
        result = read_excel_data(excel_path, sheet_name, header_name)
        print(f"读取状态: {result['status']}")
        print(f"读取消息: {result['message']}")
        if 'data' in result and result['data']:
            print("数据列表:")
            for item in result['data']:
                print(f"- {item}")
    elif mode == '写入':
        data = input("请输入要写入的数据 (逗号分隔): ").split(',')
        data = [item.strip() for item in data]
        result = write_excel_data(excel_path, sheet_name, header_name, data)
        print(f"写入状态: {result['status']}")
        print(f"写入消息: {result['message']}")
    elif mode == '多列写入':
        columns_data = {}
        while True:
            col_name = input("请输入列名 (留空结束): ")
            if not col_name:
                break
            col_data = input(f"请输入列 '{col_name}' 的数据 (逗号分隔): ").split(',')
            col_data = [item.strip() for item in col_data]
            columns_data[col_name] = col_data
        
        result = write_multiple_columns(excel_path, sheet_name, columns_data)
        print(f"多列写入状态: {result['status']}")
        print(f"多列写入消息: {result['message']}")
    else:
        print("无效的操作模式")


    
    
        
