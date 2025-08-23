from digikey import DigiKeyClient
from write_excel import read_excel_data
from write_excel import write_excel_data
import json
import os
import sys

def process_products(excel_path, sheet_name, product_number_column, output_column):
    try:
        client = DigiKeyClient()
        data = read_excel_data(excel_path, sheet_name, product_number_column)
        if not data:
            return {'status': 'error', 'message': '未获取到有效的产品数据'}
        results = {}
        total = len(data)
        print(f"\n开始处理 {total} 个产品...")
        for i, product_number in enumerate(data, 1):
            progress = i / total * 100
            sys.stdout.write(f"\r处理进度: {i}/{total} ({progress:.1f}%) - 当前产品: {product_number}")
            sys.stdout.flush()
            details = client.get_product_details(product_number)
            if isinstance(details, dict):
                # 尝试从嵌套结构中获取状态信息
                status = details.get('Product', {}).get('ProductStatus', {}).get('Status')
                if status:
                    results[product_number] = status
                else:
                    results[product_number] = f"查询失败: 未找到状态信息"
            else:
                results[product_number] = f"查询失败: {details if isinstance(details, str) else '未知错误'}"
        print("\n产品处理完成！")
        # 保存结果到JSON文件
        data_file = os.path.join(os.path.dirname(__file__), 'product_details.json')
        with open(data_file, 'w', encoding='utf-8') as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"结果已保存到 {data_file}")
        # 写入Excel（只写状态，与excel_to_digikey.py一致）
        status_list = [results.get(p, '') for p in data]
        write_result = write_excel_data(excel_path, sheet_name, output_column, status_list)
        if isinstance(write_result, dict) and write_result.get('status') == 'error':
            return write_result
        print(f"数据已成功写入Excel文件: {excel_path}")
        return {'status': 'success', 'message': f"成功处理 {len(results)} 个产品", 'data': results}
    except Exception as e:
        return {'status': 'error', 'message': f"处理过程中发生错误: {str(e)}"}

if __name__ == "__main__":
    try:
        file_path = input("请输入文件路径:")
        sheet_name = input("请输入工作表名:")
        product_number_column = input("请输入产品编号列名:")
        output_column = input("请输入输出列名:")

        result = process_products(file_path, sheet_name, product_number_column, output_column)
        print(f"处理结果: {result['status']}")
        print(f"消息: {result['message']}")

    except Exception as e:
        print(f"发生错误: {e}")


