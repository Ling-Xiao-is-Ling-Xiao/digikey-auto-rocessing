import os
import requests
import logging
from time import time
from typing import Optional, Dict
from datetime import datetime

# 配置日志
log_dir = 'logs'
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f'digikey_{datetime.now().strftime("%Y%m%d")}.log')

# 创建日志记录器
logger = logging.getLogger('digikey_client')
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

class DigiKeyClient:
    def __init__(self):
        self.token_cache = {
            'access_token': None,
            'expires_at': 0
        }
        self.client_id = os.getenv('DIGIKEY_CLIENT_ID', 'MNaKOUFqcfvGlRASTDApOcLEs0v5Y34FlaBJvJIfU0IJTQdb')

        self.client_secret = os.getenv('DIGIKEY_CLIENT_SECRET', 'GD0EfkONCW3xl0hZuyGqDpRHNH9s6DLzn3U8yx2kIYyCRflsA6kSHYOGKmGQSGPK')


    def get_access_token(self) -> str:
        if self.token_cache['access_token'] and time() < self.token_cache['expires_at']:
            return self.token_cache['access_token']

        token_data = self._request_new_token()
        self.token_cache = {
            'access_token': f"Bearer {token_data['access_token']}",
            'expires_at': time() + token_data['expires_in'] - 60
        }
        return self.token_cache['access_token']

    def _request_new_token(self) -> Dict:
        """请求新的访问令牌"""
        token_url = "https://api.digikey.com/v1/oauth2/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials"
        }
        
        try:
            response = requests.post(token_url, headers=headers, data=data)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            logger.error(f"HTTP错误: {e.response.status_code} {e.response.text}")
            raise
        except requests.exceptions.RequestException as e:
            logger.error(f"请求失败: {e}")
            raise

    def get_product_details(self, product_number: str, manufacturer_id: Optional[str] = None) -> Optional[Dict]:
        """
        使用ProductSearch API获取产品详细信息
        :param product_number: Digi-Key或制造商产品编号
        :param manufacturer_id: 可选制造商ID，用于精确匹配
        :return: 产品详细信息字典
        """
        access_token = self.get_access_token()
        headers = {
            "Authorization": access_token,
            "X-DIGIKEY-Client-Id": self.client_id,
            "X-DIGIKEY-Locale-Site": "US",
            "X-DIGIKEY-Locale-Language": "EN",
            "X-DIGIKEY-Locale-Currency": "USD",
            "X-DIGIKEY-Customer-Id": "0",
            "Content-Type": "application/json"
        }

        params = {}
        if manufacturer_id:
            params["manufacturerId"] = manufacturer_id

        max_retries = 3
        retry_delay = 1  # 初始重试延迟1秒
        
        for attempt in range(max_retries):
            try:
                url = f"https://api.digikey.com/products/v4/search/{product_number}/productdetails"
                response = requests.get(url, headers=headers, params=params, timeout=10)
                response.raise_for_status()
                return response.json()
            except requests.exceptions.HTTPError as e:
                logger.warning(f"API请求失败(尝试 {attempt + 1}/{max_retries}): {e.response.status_code} {e.response.text}")
                if attempt == max_retries - 1:
                    logger.error(f"API请求最终失败: {e.response.status_code} {e.response.text}")
                    return None
            except requests.exceptions.RequestException as e:
                logger.warning(f"API请求失败(尝试 {attempt + 1}/{max_retries}): {e}")
                if attempt == max_retries - 1:
                    logger.error(f"API请求最终失败: {e}")
                    return None
                
            # 指数退避策略
            import time
            time.sleep(retry_delay)
            retry_delay *= 2
            
        return None

    def get_product_info(self, input_str: str) -> Dict:
        """综合获取产品信息（支持URL或直接产品编号）"""
        result = {}
        try:
            # 判断输入类型
            if input_str.startswith("http"):
                product_number = input_str.split('/')[-1].split('?')[0]
                if not product_number:
                    return {"success": False, "error": "无法从URL提取产品编号"}
            else:
                # 确保输入是字符串且去除前后空格
                product_number = str(input_str).strip()
                if not product_number:
                    return {"success": False, "error": "产品编号不能为空"}

            # 获取产品详情
            product_info = self.get_product_details(product_number)
            if not product_info:
                return {"success": False, "error": "未找到产品信息"}

            # 解析产品信息
            result_data = {
                "product_description": product_info.get('Product', {}).get('Description', {}).get('ProductDescription', 'N/A'),
                "manufacturer": product_info.get('Product', {}).get('Manufacturer', {}).get('Name', 'N/A'),
                "product_url": product_info.get('Product', {}).get('ProductUrl', 'N/A'),
                "datasheet_url": product_info.get('Product', {}).get('DatasheetUrl', 'N/A'),
                "quantity_available": product_info.get('Product', {}).get('QuantityAvailable', 0),
                "product_status": product_info.get('Product', {}).get('ProductStatus', {}).get('Status', 'Status Unknown')
            }
            result = {"success": True, "data": result_data}

        except Exception as e:
            logger.error(f"处理过程中发生错误: {str(e)}")
            result = {"success": False, "error": str(e)}
        
        return result

    def get_product_info_interactive(self):
        """交互式终端操作"""
        print("\nDigiKey 产品信息查询工具（输入 exit 退出）")
        print("="*50)
        
        while True:
            user_input = input("\n请输入产品URL或编号: ").strip()
            
            if user_input.lower() in ('exit', 'quit', 'q'):
                print("退出程序。")
                break
                
            if not user_input:
                print("输入不能为空，请重新输入。")
                continue

            # 获取并展示结果
            result = self.get_product_info(user_input)
            
            if result.get("success"):
                data = result["data"]
                print(f"{data['product_status']}")  # 新增状态显示
                
                # 添加停产状态警告
                if data['product_status'].lower() == "obsolete":
                    pass
            else:
                print(f"\n错误: {result.get('error', '未知错误')}")


def digikey_api(product_number: str) -> Dict[str, str]:

    client = DigiKeyClient()
    result = client.get_product_info(product_number)
    if result.get('success'):
        return {
            'status': 'success',
            'product_status': result['data']['product_status'],
        }
    else:
        return {'status': 'error', 'message': result.get('error', '未知错误')}


if __name__ == "__main__":
    client = DigiKeyClient()
    client.get_product_info_interactive()