# -*- coding: utf-8 -*-

import logging
import simplejson
import yaml
import base64
import requests
from requests.exceptions import HTTPError

from kivy.lang import Builder
from kivy.metrics import dp
from kivy.uix.popup import Popup
from kivy.uix.recycleview import RecycleView
from kivy.uix.recyclegridlayout import RecycleGridLayout
from kivy.uix.recycleview.layout import LayoutSelectionBehavior

from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.gridlayout import GridLayout
from kivy.properties import BooleanProperty, StringProperty

from kivymd.app import MDApp
from kivymd.uix.menu import MDDropdownMenu
# from kivymd.uix.snackbar import MDSnackbar
from kivymd.uix.label import MDLabel
from kivy.clock import Clock
# from kivy.uix.textinput import TextInput
# from kivy.uix.button import Button
from kivymd.uix.textfield import MDTextField
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.recyclegridlayout import MDRecycleGridLayout
from kivy.core.text import LabelBase
from kivymd.font_definitions import theme_font_styles
from kivy.properties import ObjectProperty


from utility.util_odoo import UtilOdoo
from bravado.requests_client import RequestsClient
from bravado.client import SwaggerClient
from swagger_spec_validator.common import SwaggerValidationError



# 註冊支持中文的黑體字體
LabelBase.register(name='JhengHei', fn_regular='fonts/msjh.ttc')
LabelBase.register(name='JhengHei-Bold', fn_regular='fonts/msjhbd.ttc')

_logger = logging.getLogger(__name__)
_logger.setLevel(logging.INFO)

# log to console
c_handler = logging.StreamHandler()

console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
c_handler.setFormatter(console_format)
c_handler.setLevel = logging.DEBUG

_logger.addHandler(c_handler)

KV = '''
<SelectableLabel>:
    # orientation: 'horizontal'
    canvas.before:
        Color:
            rgba: (0, 0, 0, 0.1) if self.selected else (0, 0, 0, 0)
        Rectangle:
            size: self.size
            pos: self.pos
    BoxLayout:
        orientation: 'horizontal'
        size_hint_x: 1  # 確保 BoxLayout 佔據整個寬度
        size_hint_y: 0.8  # 確保 BoxLayout 佔據整個高度
        size_hint_y: None  # 取消垂直方向上的大小提示
        height: dp(56)  # 設置固定高度
        Label:
            text: root.name
            font_name: root.font_name
            color: 1, 1, 1, 1  # Set text color to white
            size_hint_x: 0.8  # 第一欄位佔 80%
            text_size: self.width, None  # 設置文本大小以啟用自動換行
            halign: 'left'  # 設置文本水平對齊方式
        Label:
            text: root.uom
            font_name: root.font_name
            color: 1, 1, 1, 1  # Set text color to white
            size_hint_x: 0.2  # 第二欄位佔 20%
            haligh: 'right'  # 設置文本水平對齊方式

<RV>:
    viewclass: 'SelectableLabel'
    # RecycleGridLayout:
    #     cols: 2  # 設置列數
    #     default_size: None, dp(56)
    #     default_size_hint: 1, None
    #     size_hint_y: None
    #     height: self.minimum_height
    #     orientation: 'tb-lr'        
    MDRecycleGridLayout:
        cols: 2  # 設置列數
        default_size: None, dp(56)
        default_size_hint: 1, None
        size_hint_y: None
        height: self.minimum_height
        # width: self.width  # 設置寬度為父容器的寬度
        orientation: 'tb-lr'
        padding: dp(10)  # Add padding to avoid overlap
        spacing: dp(10)  # Add spacing between items
    
MDBoxLayout:
    orientation: "vertical"

    MDTopAppBar:
        id: toolbar
        title: "材料及加工方式"
        left_action_items: [["menu", lambda x: app.callback(x)]]
        right_action_items: [["dots-vertical", lambda x: app.callback(x)]]
        font_name: "JhengHei"

    ScrollView:
        size_hint_y: 1
        MDBoxLayout:
            id: content_box
            orientation: "vertical"
            padding: [20, 20]  # 左邊距 20 像素，上下邊距 20 像素
            size_hint_y: None
            height: self.minimum_height
            font_name: "JhengHei"
'''

class KivyLoggerHandler(logging.Handler):
    def __init__(self, update_func):
        super().__init__()
        self.update_func = update_func

    def emit(self, record):
        log_entry = self.format(record)
        self.update_func(log_entry)

class SelectableLabel(RecycleDataViewBehavior, Label):
# class SelectableLabel(RecycleDataViewBehavior, GridLayout):
    index = None
    selected = BooleanProperty(False)
    selectable = BooleanProperty(True)
    # text = StringProperty('')
    name = StringProperty('')
    uom = StringProperty('')
    font_name = StringProperty('JhengHei')

    def refresh_view_attrs(self, rv, index, data):
        print(f"refresh_view_attrs called with index: {index}, data: {data}")
        self.index = index
        # self.text = data['text']
        self.name = data['name']
        self.uom = data['uom']
        return super(SelectableLabel, self).refresh_view_attrs(rv, index, data)

    def on_touch_down(self, touch):
        print(f"on_touch_down called with touch: {touch}")
        if super(SelectableLabel, self).on_touch_down(touch):
            print("super on_touch_down returned True")
            return True
        if self.collide_point(*touch.pos) and self.selectable:
            print(f"collide_point: {self.collide_point(*touch.pos)}, selectable: {self.selectable}")
            self.parent.parent.select_with_touch(self.index, touch)
            app = MDApp.get_running_app()
            print(f"Selected name: {self.name}")
            app.set_item(self.name)
            app.popup.dismiss()
            return True
        return False

class RV(RecycleView):
    def __init__(self, **kwargs):
        super(RV, self).__init__(**kwargs)
        self.data = []

    def select_with_touch(self, index, touch):
        print(f"select_with_touch called with index: {index}, touch: {touch}")
        # 在這裡添加選擇項目的邏輯
        for item in self.data:
            item['selected'] = False
        self.data[index]['selected'] = True
        print('Updated RV data:', self.data)
        self.refresh_from_data()

class FormMain(MDApp):
    odoo = ObjectProperty(None)
    requestOptions = ObjectProperty(None)
    token = ObjectProperty(None)
    odoo_connection = ObjectProperty(None)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        print("kwargs:", kwargs)
        _logger.info("kwargs: %s", kwargs)
        self.odoo = kwargs.get('odoo')
        self.requestOptions = kwargs.get('requestOptions')
        self.token = kwargs.get('token')
        self.odoo_connection = kwargs.get('odoo_connection')
        self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
        self.filtered_items = []  # 初始化 filtered_items

    def set_toolbar_font_name(self, *args):
        self.root.ids.toolbar.ids.label_title.font_name = "JhengHei-Bold"

    def set_toolbar_font_size(self, *args):
        self.root.ids.toolbar.ids.label_title.font_size = '28sp'

    def build(self):
        theme_font_styles.append('JhengHei')  # 將新字體添加到 theme_font_styles 中
        self.theme_cls.font_styles['JhengHei'] =  [
            "JhengHei",
            16, #大小
            True, #粗體
            0.15, #行高
        ]
        Clock.schedule_once(self.set_toolbar_font_name)
        Clock.schedule_once(self.set_toolbar_font_size)

        self.theme_cls.primary_palette = "Orange"
        self.theme_cls.theme_style = "Dark"
        self.log_messages = []
        self.dropdown_items = [
            "Option 1",
            "Option 2",
            "Option 3",
        ]
        menu_items = [
            {
                "text": "Connecting to Odoo",
                "on_release": lambda: self.connect_odoo(self.odoo_connection),
            },
            {
                "text": "Connecting to AutoCAD",
                "on_release": lambda: self.connect_autocad(),
            },
            {
                "text": "Getting parameters from Odoo",
                "on_release": lambda: self.get_parameters_from_odoo(),
            },
            {
                "text": "Pushing to BOQ",
                "on_release": lambda: self.push_to_boq(),
            },
            {
                "text": "Transferring BOQ to PR",
                "on_release": lambda: self.transfer_boq_to_pr(),
            },
        ]
        self.menu = MDDropdownMenu(items=menu_items)
        self.menu.bind(on_open=self.on_menu_open, on_dismiss=self.on_menu_dismiss)
        self.menu_open = False
        self.setup_logger()
        return Builder.load_string(KV)

    def on_start(self):
        super().on_start()
        print("on_start called", self.root.ids)
        Clock.schedule_once(self.set_toolbar_font_name)
        Clock.schedule_once(self.set_toolbar_font_size)

    def on_menu_open(self, instance_menu):
        self.menu_open = True

    def on_menu_dismiss(self, instance_menu):
        self.menu_open = False

    def setup_logger(self):
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)

        # Create custom handler
        kivy_handler = KivyLoggerHandler(self._logger)
        console_format = logging.Formatter("%(asctime)s: %(name)-18s [%(levelname)s] %(message)s")
        kivy_handler.setFormatter(console_format)
        kivy_handler.setLevel(logging.DEBUG)

        # Add handler to the logger
        self.logger.addHandler(kivy_handler)

    def callback(self, button):
        self.menu.caller = button
        self.menu.open()

    def clear_content(self):
        self.log_messages = []
        self.update_content()

    def get_odoo_product(self):
        product_list = self.odoo_util.get_product()
        print(f'product_list:{product_list}\n')
        return product_list

    def string_to_base64(self, input_string):
        # 將字串轉換成 UTF-8 編碼的位元組序列
        input_bytes = input_string.encode('utf-8')        
        # 使用 Base64 編碼位元組序列
        base64_bytes = base64.b64encode(input_bytes)
        # 將 Base64 編碼後的位元組序列轉換成字串
        base64_string = base64_bytes.decode('utf-8')
        return base64_string

    def connect_odoo(self, odoo_conn):
        if not self.odoo:
            self.menu.dismiss()
            self.clear_content()
            self.logger.info("Connecting to Odoo...")
            # 模擬一些操作
            host = odoo_conn['host']
            db_name = odoo_conn['db_name']
            url = odoo_conn['url']
            token = odoo_conn['token']
            http_client = RequestsClient()
            http_client.set_basic_auth(host, db_name, token)
            try:
                odoo = SwaggerClient.from_url(url, http_client=http_client)
                # prompt(acaduti, f"與Odoo連線成功\n")
                # print(f"與Odoo連線成功\n")

                basic_string = f'{db_name}:{token}'
                basic_token = self.string_to_base64(basic_string)
                print(f'basic_token:{basic_token}\n')
                headers = {
                    'Authorization': f'Basic {basic_token}'
                }
                print(f'headers:{headers}\n')

                requestOptions = {
                    # === bravado config ===
                    'headers': headers,
                }

                _logger.info(f"與Odoo連線成功\n")
                # return odoo, requestOptions, token
                self.odoo = odoo
                self.requestOptions = requestOptions
                self.token = token
                self.odoo_util = UtilOdoo(self.odoo, self.requestOptions, self.token)
            except requests.exceptions.ConnectionError:
                # print(f"無法與Odoo，通常多試幾次會成功\n")
                _logger.info(f"無法與Odoo，通常多試幾次會成功\n")
                # prompt(acaduti, f"無法與Odoo，通常多試幾次會成功\n")
                return
            except (
                simplejson.errors.JSONDecodeError,      # type: ignore
                yaml.YAMLError,
                HTTPError,
                ):
                _logger.info(f"Invalid swagger file. Please check to make sure the swagger file can be found at: {url}.\n")

                return
            except SwaggerValidationError:
                # print('Invalid swagger format.\n')
                _logger.info(f'Invalid swagger format.')
                return

    def connect_autocad(self):
        self.menu.dismiss()
        self.clear_content()
        self.logger.info("Connecting to AutoCAD...")
        # 模擬一些操作
        Clock.schedule_once(lambda dt: self.logger.info("Connected to AutoCAD successfully."), 1)

    def get_parameters_from_odoo(self):
        self.menu.dismiss()
        self.clear_content()
        # self.logger.info("Getting parameters from Odoo...")
        # 創建一個新的表單佈局
        form_layout = MDBoxLayout(orientation='vertical', padding=[20, 20], size_hint_y=None)
        form_layout.bind(minimum_height=form_layout.setter('height'))

        # 第一行
        row1 = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(40))
        row1.add_widget(MDLabel(text="材料:", size_hint_y=None, height=dp(30), size_hint_x=0.3, font_style='JhengHei'))
        
        self.selected_item = TextInput(size_hint_y=None, height=dp(30), size_hint_x=0.7, multiline=False, font_name='JhengHei')
        self.selected_item.bind(on_text_validate=self.open_product_menu)

        row1.add_widget(self.selected_item)
        form_layout.add_widget(row1)

        # 第二行
        row2 = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(40))
        row2.add_widget(MDLabel(text="材質:", size_hint_y=None, height=dp(30), size_hint_x=0.3, font_style='JhengHei'))

        row2.add_widget(TextInput(size_hint_y=None, height=dp(30), size_hint_x=0.7, multiline=False, font_name='JhengHei'))
        form_layout.add_widget(row2)

        # 添加一個空間來推動按鈕到最下面
        form_layout.add_widget(MDBoxLayout(size_hint_y=None, height=dp(350)))

        # 添加按鈕
        form_layout.add_widget(Button(text="Submit", size_hint_y=None, height=dp(40)))

        # 將表單佈局添加到 content_box 中
        self.root.ids.content_box.add_widget(form_layout)

    def open_product_menu(self, instance):
        # 獲取產品列表
        self.product_menu_items = self.get_odoo_product()
        
        # 獲取搜索文本並轉換為小寫
        search_text = instance.text.lower()
        print('search_text:', search_text)
        
        # 過濾產品列表
        self.filtered_items = [
            item for item in self.product_menu_items 
            if search_text in item['name'].lower()
        ]

        # 如果沒有過濾到任何項目，關閉彈出窗口
        if not self.filtered_items:
            if hasattr(self, 'popup') and self.popup:
                self.popup.dismiss()
            return

        # 顯示過濾後的結果
        self.show_popup()

    # def show_popup(self):
    #     content = BoxLayout(orientation='vertical')
        
    #     # 添加搜尋欄位
    #     search_input = TextInput(size_hint_y=None, height=dp(40), multiline=False, font_name='JhengHei')
    #     search_input.bind(text=self.update_filter)
    #     content.add_widget(search_input)

    #     # 初始化 self.filtered_items
    #     if not hasattr(self, 'filtered_items') or not self.filtered_items:
    #         self.filtered_items = self.product_menu_items

    #     # 創建 RV 並設置數據
    #     self.rv = RV()
    #     self.rv.viewclass = 'SelectableLabel'
        
    #     grid_layout = RecycleGridLayout(cols=2, default_size=(None, dp(40)), size_hint_y=None)
    #     grid_layout.bind(minimum_height=grid_layout.setter('height'))
    #     self.rv.add_widget(grid_layout)  # 確保 grid_layout 是 rv 的子級
    #     self.rv.layout_manager = grid_layout
        
    #     self.rv.data = [{'text': f"{item['name']}", 'font_name': 'JhengHei'} for item in self.filtered_items]
    #     self.rv.data.extend([{'text': f"{item['uom']}", 'font_name': 'JhengHei'} for item in self.filtered_items])
        
    #     content.add_widget(self.rv)

    #     # 添加調試輸出
    #     print('self.filtered_items:', self.filtered_items)
    #     print('self.rv.data:', self.rv.data)

    #     # 每次打開 popup 前重置它
    #     if hasattr(self, 'popup') and self.popup:
    #         self.popup.dismiss()
    #     self.popup = Popup(title='Select an option', content=content, size_hint=(0.8, 0.8))
    #     self.popup.open()

    # def update_filter(self, instance, value):
    #     search_text = value.lower()
    #     self.filtered_items = [item for item in self.product_menu_items if search_text in item['name'].lower()]
        
    #     # 更新 RV 的數據
    #     self.rv.data = [{'text': f"{item['name']}", 'font_name': 'JhengHei'} for item in self.filtered_items]
    #     self.rv.data.extend([{'text': f"{item['uom']}", 'font_name': 'JhengHei'} for item in self.filtered_items])
    #     self.rv.refresh_from_data()

    #     # 添加調試輸出
    #     print('Updated self.filtered_items:', self.filtered_items)
    #     print('Updated self.rv.data:', self.rv.data)

    def show_popup(self):
        content = BoxLayout(orientation='vertical')
        
        # 添加搜尋欄位
        search_input = TextInput(size_hint_y=None, height=dp(40), multiline=False, font_name='JhengHei')
        search_input.bind(text=self.update_filter)
        content.add_widget(search_input)
        
        # 創建 RV 並設置數據
        self.rv = RV()
        self.rv.viewclass = 'SelectableLabel'

        # grid_layout = MDRecycleGridLayout(cols=2, default_size=(None, dp(40)), size_hint_y=None)
        # grid_layout.bind(minimum_height=grid_layout.setter('height'))
        # self.rv.layout_manager = grid_layout

        # Print filtered items before setting data
        # print('Filtered items:', self.filtered_items)

        # self.rv.data = [{'text': f"{item['name']} ({item['uom']})", 'font_name': 'JhengHei'} for item in self.filtered_items]
        self.rv.data = [{'name': item['name'], 'uom': item['uom'], 'font_name': 'JhengHei'} for item in self.filtered_items]
        print('RV data:', self.rv.data)

        content.add_widget(self.rv)

        # print('rv.data:', self.rv.data)  # 確認 rv.data 的內容
        # print('filtered_items:', self.filtered_items)  # 確認 filtered_items 的內容
        # print('content:', content)  # 確認 content 的內容

        # 每次打開 popup 前重置它
        if hasattr(self, 'popup') and self.popup:
            self.popup.dismiss()
        self.popup = Popup(title='Select an option', content=content, size_hint=(0.8, 0.8))
        self.popup.open()

    def update_filter(self, instance, value):
        search_text = value.lower()
        self.filtered_items = [item for item in self.product_menu_items if search_text in item['name'].lower()]
        
        # 更新 RV 的數據
        # self.rv.data = [{'text': f"{item['name']} ({item['uom']})", 'font_name': 'JhengHei'} for item in self.filtered_items]
        self.rv.data = [{'name': item['name'], 'uom': item['uom'], 'font_name': 'JhengHei'} for item in self.filtered_items]
        self.rv.refresh_from_data()

    def set_item(self, text_item):
        self.selected_item.text = text_item
        # self.dropdown_menu.dismiss()

    def push_to_boq(self):
        self.menu.dismiss()
        self.clear_content()
        self.logger.info("Pushing to BOQ...")
        # 模擬一些操作
        Clock.schedule_once(lambda dt: self.logger.info("Pushed to BOQ successfully."), 1)

    def transfer_boq_to_pr(self):
        self.menu.dismiss()
        self.clear_content()
        self.logger.info("Transferring BOQ to PR...")
        # 模擬一些操作
        Clock.schedule_once(lambda dt: self.logger.info("Transferred BOQ to PR successfully."), 1)

    def _logger(self, message):
        self.log_messages.append(message)
        self.update_content()

    def update_content(self):
        self.root.ids.content_box.clear_widgets()
        for message in self.log_messages:
            label = MDLabel(text=message, halign="left", padding=[20, 0], size_hint_y=None, size_hint_x=1, font_style='JhengHei')
            label.bind(texture_size=lambda instance, value: setattr(instance, 'height', value[1]))
            self.root.ids.content_box.add_widget(label)




if __name__ == '__main__':
    FormMain().run()