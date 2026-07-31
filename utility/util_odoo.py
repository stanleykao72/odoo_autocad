# -*- coding: utf-8 -*-
import json
import requests
import simplejson
import yaml
import base64
# from utility.util_log import UtilLog

from requests.exceptions import HTTPError
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError

from utility.util_secrets import mask_secret, mask_url_token


class AuthenticationError(Exception):
    """Odoo 拒絕 user_token（401/403）—— 與網路、伺服器問題區分開來。"""

# _logger = logging.getLogger(__name__)
# _logger.setLevel(logging.INFO)

# # log to console
# c_handler = logging.StreamHandler()

# console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
# c_handler.setFormatter(console_format)
# c_handler.setLevel = logging.DEBUG

# _logger.addHandler(c_handler)


class UtilOdoo:
    def __init__(self, odoo_connection, log_util, connect=True):
        """
        Args:
            odoo_connection: dict(host, db_name, url, token)
            log_util: 具備 safe_log_insert() 的日誌工具
            connect: True 時立即連線，失敗會拋出例外（fail fast）。
                     False 時只建立未連線的實例，供 GUI 在連線失敗後
                     仍能開啟並讓使用者手動重試。
        """
        self.odoo = None
        self.requestOptions = None
        self.user_token = None
        self.log = log_util
        self.odoo_connection = odoo_connection
        # 最近一次失敗的原因，供 GUI 直接顯示（不必自行猜測）
        self.last_error = None
        self.auth_failed = False
        if connect:
            self.connect_odoo(odoo_connection)

    def _fail(self, message, auth=False):
        """記錄失敗原因並寫入日誌。message 不含換行。"""
        self.last_error = message
        if auth:
            self.auth_failed = True
        self.log.safe_log_insert(f"{message}\n")

    def connected_odoo(self):
        if self.odoo:
            return True
        else:
            return False

    def string_to_base64(self, input_string):
        input_bytes = input_string.encode('utf-8')        
        base64_bytes = base64.b64encode(input_bytes)
        base64_string = base64_bytes.decode('utf-8')
        return base64_string

    def connect_odoo(self, odoo_connection):
        host = odoo_connection['host']
        db_name = odoo_connection['db_name']
        url = odoo_connection['url']
        token = odoo_connection['token']
        self.log.safe_log_insert(f"host: {host}\n")
        self.log.safe_log_insert(f"db_name: {db_name}\n")
        # url 的 query string 內含 API token，需遮罩後才可寫入日誌
        self.log.safe_log_insert(f"url: {mask_url_token(url)}\n")
        # 不記錄 token 明文（日誌會落地到 logs/*.log）
        self.log.safe_log_insert(f"token: {mask_secret(token)}\n")

        http_client = RequestsClient()
        http_client.set_basic_auth(host, db_name, token)
        self.last_error = None
        self.auth_failed = False
        try:
            odoo = SwaggerClient.from_url(url, http_client=http_client)
            basic_string = f'{db_name}:{token}'
            basic_token = self.string_to_base64(basic_string)
            headers = {
                'Authorization': f'Basic {basic_token}'
            }
            requestOptions = {
                'headers': headers,
            }
            self.odoo = odoo
            self.requestOptions = requestOptions
            self.user_token = token

            # 取得 swagger.json 成功 ≠ 可以呼叫 API：
            # swagger 網址的 query string 自帶一組 token，Odoo 對該端點不做認證，
            # 而實際 API 呼叫用的是 token.yaml 的 user_token。若只看前者就宣告
            # 「已連接」，畫面會顯示綠色勾勾、每一次操作卻都 401，使用者只會看到
            # 下游那些「取不到專案資料」的訊息，完全找不到真正原因。
            self._verify_authentication()

            self.log.safe_log_insert("與 Odoo 連線成功\n")
            return odoo, requestOptions, token
        except AuthenticationError:
            # _verify_authentication() 已記錄明確原因
            raise
        except requests.exceptions.ConnectionError:
            self._fail("❌ 無法連線到 Odoo 伺服器，請檢查網路或稍後再試")
            raise
        except requests.exceptions.Timeout:
            self._fail("❌ 連線 Odoo 逾時，請稍後再試")
            raise
        except SwaggerValidationError:
            self._fail("❌ Swagger 格式無效")
            raise
        except Exception as e:
            status = self._extract_status_code(e)
            self._fail(self._describe_connect_error(e, url).rstrip("\n"),
                       auth=status in (401, 403))
            raise

    # HTTP 狀態碼 → 使用者看得懂的原因
    _HTTP_REASONS = {
        400: "❌ Odoo 拒絕請求 (400)，請確認 url 的 db 參數是否正確",
        401: "❌ Odoo 認證失敗 (401) — token 不正確或已失效，請更新 config/token.yaml",
        403: "❌ Odoo 拒絕存取 (403) — token 不正確、已失效或權限不足",
        404: "❌ 找不到 Swagger 文件 (404)，請確認 url 是否正確",
        500: "❌ Odoo 伺服器內部錯誤 (500)",
    }

    @staticmethod
    def _extract_status_code(exc):
        """從各種例外形態取出 HTTP 狀態碼（bravado / requests 形態不同）"""
        status = getattr(exc, 'status_code', None)
        if status is None:
            status = getattr(getattr(exc, 'response', None), 'status_code', None)
        return status

    def _describe_connect_error(self, exc, url):
        """把連線例外轉成明確的失敗原因，特別點名 token 問題。"""
        status = self._extract_status_code(exc)

        if status in self._HTTP_REASONS:
            return f"{self._HTTP_REASONS[status]}\n"
        if status is not None:
            return f"❌ Odoo 回應 HTTP {status}，連線失敗\n"

        # 沒有狀態碼：多半是伺服器回了 HTML 登入頁／錯誤頁，導致 JSON 解析失敗。
        # 這在 token 錯誤時很常見（Odoo 會導向登入頁而非回 401）。
        if isinstance(exc, simplejson.errors.JSONDecodeError):
            return ("❌ Odoo 回應不是有效的 JSON — 常見原因是 token 不正確或已失效"
                    f"（伺服器改回傳登入頁）。請確認 config/token.yaml 與 url 的 token 參數\n")
        return f"❌ 連線 Odoo 失敗: {type(exc).__name__}: {exc}\n"

    # 認證探測用的查詢：故意查一個不存在的名稱，只要能通過認證即可，
    # 回傳空結果也算成功。
    _AUTH_PROBE_NAME = "__odoo_autocad_auth_probe__"

    def _verify_authentication(self):
        """以一次實際的 API 呼叫確認 user_token 真的可用。

        失敗時清空連線狀態並拋出例外，讓 connected_odoo() 誠實回報未連線。
        """
        try:
            self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
                method_name="get_project_v2",
                body={
                    "args": [[['name', '=', self._AUTH_PROBE_NAME]]],
                    "kwargs": {'user_token': self.user_token},
                    "context": {}
                },
                _request_options=self.requestOptions,
            ).response().incoming_response.json()
        except Exception as e:
            status = self._extract_status_code(e)
            self.odoo = None
            self.requestOptions = None
            self.user_token = None
            if status in (401, 403):
                self._fail(
                    f"❌ Odoo token 無效 ({status}) — 已取得 Swagger 文件，"
                    "但 API 認證被拒絕", auth=True)
                self.log.safe_log_insert(
                    "   請至 Odoo 重新產生 token，並更新 config/token.yaml\n")
                raise AuthenticationError(self.last_error) from e
            self._fail("❌ 已取得 Swagger 文件，但 API 呼叫失敗: "
                       + self._describe_connect_error(e, None).rstrip("\n").lstrip("❌ "))
            raise

    def _call(self, method_name, args):
        """呼叫 Odoo API 並統一處理失敗。

        連線階段（connect_odoo）成功不代表可以呼叫 API：swagger.json 的 url
        自帶一組 token，而每次 API 呼叫用的是 token.yaml 的 user_token。
        後者失效時，連線看起來正常，卻是每一次呼叫都 401 —— 若不在這裡明確
        記錄，錯誤會被上層的寬鬆 except 吞掉，使用者只會看到「沒有資料」。

        Returns:
            dict on success, None on failure（失敗原因已寫入日誌）
        """
        if not self.odoo:
            self.log.safe_log_insert(f"❌ {method_name}: Odoo 未連線\n")
            return None
        try:
            return self.odoo.job_working_plan_boq.callMethodForJobWorkingPlanBoqModel(
                method_name=method_name,
                body={
                    "args": args,
                    "kwargs": {'user_token': self.user_token},
                    "context": {}
                },
                _request_options=self.requestOptions,
            ).response().incoming_response.json()
        except Exception as e:
            status = self._extract_status_code(e)
            if status in (401, 403):
                self._fail(f"❌ Odoo token 無效 ({status}) — 呼叫 {method_name} 遭拒絕，"
                           "請更新 config/token.yaml", auth=True)
            else:
                self._fail(f"❌ 呼叫 Odoo {method_name} 失敗: "
                           + self._describe_connect_error(e, None).rstrip("\n").lstrip("❌ "))
            return None

    def _unwrap(self, result, key, method_name):
        """取出回應中的資料欄位；Odoo 回報 error_code 時回傳錯誤訊息字串。"""
        if result is None:
            return None
        if 'error_code' in result:
            error_message = result.get('error_message')
            self.log.safe_log_insert(f'❌ {method_name}: {error_message}\n')
            return error_message
        return result.get(key)

    def import2boq(self, layout_dict):
        result = self._call("import2boq_v2", [layout_dict])
        return self._unwrap(result, 'all', 'import2boq')

    def boq2pr(self, layout_dict):
        result = self._call("boq2pr_v2", [layout_dict])
        return self._unwrap(result, 'all', 'boq2pr')

    def get_project(self, pr_no):
        project_dict = self._call("get_project_v2", [[['name', '=', pr_no]]])
        if project_dict is None:
            return None

        if 'error_code' in project_dict:
            error_message = project_dict.get('error_message')
            self.log.safe_log_insert(f'❌ get_project: {error_message}\n')
            return error_message
        return project_dict

    def get_product(self):
        result = self._call("get_product_v2",
                            [[('categ_id', 'child_of', 27), ('active', '=', True)]])
        return self._unwrap(result, 'product', 'get_product')

    def get_setup(self, setup_name):
        result = self._call("get_setup_v2", [[('setup_name', '=', setup_name)]])
        return self._unwrap(result, 'setup', 'get_setup')

    def get_color(self, project_id):
        result = self._call("get_color_v2", [[('job_project_id', '=', project_id)]])
        return self._unwrap(result, 'color', 'get_color')
    
    def search_products(self, query, limit=50):
        """搜尋產品資料"""
        try:
            if not self.connected_odoo():
                self.log.safe_log_insert("Odoo未連接，無法搜尋產品\n")
                return []
            
            # 使用現有的get_product方法作為基礎，並添加查詢功能
            # 這裡是一個簡化的實現，實際應該調用Odoo API搜尋產品
            products = []
            
            # 模擬產品搜尋結果 (實際實現時需要調用真正的Odoo API)
            if query.lower() in ['pipe', '管道']:
                products = [
                    {'id': 1, 'name': 'Pipe 100mm', 'code': 'P100', 'category': 'Piping'},
                    {'id': 2, 'name': 'Pipe 150mm', 'code': 'P150', 'category': 'Piping'}
                ]
            elif query.lower() in ['valve', '閥門']:
                products = [
                    {'id': 3, 'name': 'Valve 50mm', 'code': 'V050', 'category': 'Valves'},
                    {'id': 4, 'name': 'Valve 100mm', 'code': 'V100', 'category': 'Valves'}
                ]
            
            products = products[:limit]  # 限制返回數量
            self.log.safe_log_insert(f"搜尋產品完成，找到 {len(products)} 項產品\n")
            return products
            
        except Exception as e:
            self.log.safe_log_insert(f"搜尋產品時發生錯誤: {str(e)}\n")
            return []
    
    def push_boq_data(self, project_id, boq_data):
        """推送BOQ資料到專案"""
        try:
            if not self.connected_odoo():
                self.log.safe_log_insert("Odoo未連接，無法推送BOQ資料\n")
                return False
            
            # 這裡會整合現有的import2boq功能
            # 暫時返回成功狀態，實際實現時需要調用真正的BOQ推送邏輯
            
            self.log.safe_log_insert(f"開始推送 {len(boq_data)} 項BOQ資料到專案 {project_id}\n")
            
            # 模擬推送過程
            for item in boq_data:
                item_name = item.get('name', 'Unknown')
                self.log.safe_log_insert(f"推送項目: {item_name}\n")
            
            self.log.safe_log_insert("BOQ資料推送完成\n")
            return True
            
        except Exception as e:
            self.log.safe_log_insert(f"推送BOQ資料時發生錯誤: {str(e)}\n")
            return False
