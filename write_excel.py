import openpyxl

def read_excel_data(excel_path, sheet_name, header_name):
    try:
        # 加载Excel文件
        workbook = openpyxl.load_workbook(excel_path)
        
        # 检查工作表是否存在
        if sheet_name not in workbook.sheetnames:
            raise Exception(f"工作表 '{sheet_name}' 不存在")
            
        # 选择工作表
        sheet = workbook[sheet_name]
        
        # 查找表头所在列
        header_column = None
        for cell in sheet[1]:  # 假设表头在第一行
            if cell.value == header_name:
                header_column = cell.column
                break
                
        if not header_column:
            raise Exception(f"表头 '{header_name}' 未找到")
        
        # 从表头下方读取所有数据
        data = []
        for row in sheet.iter_rows(min_row=2, min_col=header_column, max_col=header_column):
            if row[0].value:
                data.append(str(row[0].value).strip())
        
        if not data:
            print("警告: 未找到有效数据")
            return []
        
        print(f"成功读取 {len(data)} 条数据")
        return data
        
    except Exception as e:
        raise Exception(f"读取Excel文件时发生错误: {str(e)}")


def write_excel_data(excel_path, sheet_name, header_name, data):
    try:
        if not data:
            return {'status': 'error', 'message': '数据列表不能为空'}
        workbook = openpyxl.load_workbook(excel_path)
        if sheet_name not in workbook.sheetnames:
            return {'status': 'error', 'message': f"工作表 '{sheet_name}' 不存在"}
        sheet = workbook[sheet_name]
        # 查找表头所在列
        header_column = None
        for cell in sheet[1]:
            if cell.value == header_name:
                header_column = cell.column
                break
        if not header_column:
            header_column = sheet.max_column + 1
            sheet.cell(row=1, column=header_column, value=header_name)
        # 写入数据（不清空原有数据，直接覆盖/追加）
        for idx, value in enumerate(data, start=2):
            sheet.cell(row=idx, column=header_column, value=value)
        workbook.save(excel_path)
        return {'status': 'success', 'message': f"成功写入 {len(data)} 条数据到Excel文件"}
    except Exception as e:
        return {'status': 'error', 'message': f"写入Excel文件时发生错误: {str(e)}"}

if __name__ == '__main__':
    # 示例用法 - 读取数据
    excel_path = input("请输入Excel文件路径: ")
    sheet_name = input("请输入工作表名称: ")
    header_name = input("请输入表头名称: ")
    mode = input("请输入操作模式 (读取/写入): ")
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
    else:
        print("无效的操作模式")


    
    
        
