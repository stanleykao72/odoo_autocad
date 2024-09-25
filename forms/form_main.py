# -*- coding: utf-8 -*-

import logging
from kivy.lang import Builder
from kivy.metrics import dp
from kivy.uix.popup import Popup
from kivy.uix.recycleview import RecycleView
from kivy.uix.recycleview.views import RecycleDataViewBehavior
from kivy.uix.label import Label
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
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
from kivy.core.text import LabelBase

# 註冊支持中文的黑體字體
# LabelBase.register(name='SimHei',
#                    fn_regular='SimHei.ttf')

KV = '''
MDBoxLayout:
    orientation: "vertical"

    MDTopAppBar:
        title: "MDTopAppBar"
        left_action_items: [["menu", lambda x: app.callback(x)]]
        right_action_items: [["dots-vertical", lambda x: app.callback(x)]]

    ScrollView:
        size_hint_y: 1
        MDBoxLayout:
            id: content_box
            orientation: "vertical"
            padding: [20, 20]  # 左邊距 20 像素，上下邊距 20 像素
            size_hint_y: None
            height: self.minimum_height
'''

class KivyLoggerHandler(logging.Handler):
    def __init__(self, update_func):
        super().__init__()
        self.update_func = update_func

    def emit(self, record):
        log_entry = self.format(record)
        self.update_func(log_entry)

class SelectableLabel(RecycleDataViewBehavior, Label):
    index = None
    selected = BooleanProperty(False)
    selectable = BooleanProperty(True)
    text = StringProperty('')

    def refresh_view_attrs(self, rv, index, data):
        self.index = index
        self.text = data['text']
        return super(SelectableLabel, self).refresh_view_attrs(rv, index, data)

    def on_touch_down(self, touch):
        if super(SelectableLabel, self).on_touch_down(touch):
            return True
        if self.collide_point(*touch.pos) and self.selectable:
            self.parent.parent.select_with_touch(self.index, touch)
            app = MDApp.get_running_app()
            app.set_item(self.text)
            app.popup.dismiss()
            return True

class RV(RecycleView):
    def __init__(self, **kwargs):
        super(RV, self).__init__(**kwargs)
        self.data = []

class FormMain(MDApp):
    def build(self):
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
                "on_release": lambda: self.connect_odoo(),
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

    def connect_odoo(self):
        self.menu.dismiss()
        self.clear_content()
        self.logger.info("Connecting to Odoo...")
        # 模擬一些操作
        Clock.schedule_once(lambda dt: self.logger.info("Connected to Odoo successfully."), 1)

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
        row1.add_widget(MDLabel(text="Parameter 1:", size_hint_y=None, height=dp(30), size_hint_x=0.3))
        
        # self.selected_item = MDTextField(size_hint_y=None, height=dp(30), size_hint_x=0.7)
        self.selected_item = TextInput(size_hint_y=None, height=dp(30), size_hint_x=0.7)
        # self.dropdown_menu = MDDropdownMenu(
        #     caller=self.selected_item,
        #     items=self.dropdown_items,
        #     width_mult=4,
        # )
        # self.selected_item.bind(on_focus=self.open_menu)
        # self.selected_item.bind(on_touch_down=self.open_menu)
        self.selected_item.bind(text=self.open_menu)
        # self.selected_item.bind(on_text_validate=self.open_menu)

        row1.add_widget(self.selected_item)
        form_layout.add_widget(row1)

        # 第二行
        row2 = MDBoxLayout(orientation='horizontal', size_hint_y=None, height=dp(40))
        row2.add_widget(MDLabel(text="Parameter 2:", size_hint_y=None, height=dp(30), size_hint_x=0.3))
        row2.add_widget(TextInput(size_hint_y=None, height=dp(30), size_hint_x=0.7))
        form_layout.add_widget(row2)

        # 添加一個空間來推動按鈕到最下面
        form_layout.add_widget(MDBoxLayout(size_hint_y=None, height=dp(350)))

        # 添加按鈕
        form_layout.add_widget(Button(text="Submit", size_hint_y=None, height=dp(40)))

        # 將表單佈局添加到 content_box 中
        self.root.ids.content_box.add_widget(form_layout)

    # def open_menu(self, instance, value):
    #     # self.logger.info(f"open_menu called with instance: {instance}, value: {value}")
    #     search_text = instance.text.lower()
    #     filtered_items = [item for item in self.dropdown_items if search_text in item["text"].lower()]
    #     self.dropdown_menu.items = filtered_items
    #     self.dropdown_menu.caller = instance
    #     if not self.dropdown_menu.parent:
    #         self.dropdown_menu.open()

    def open_menu(self, instance, value):
        # self.logger.info(f"open_menu called with instance: {instance}, value: {value}")
        search_text = instance.text.lower()
        filtered_items = [item for item in self.dropdown_items if search_text in item.lower()]

        if not filtered_items:
            if hasattr(self, 'popup') and self.popup:
                self.popup.dismiss()
            return

        content = BoxLayout(orientation='vertical')
        rv = RV()
        rv.data = [{'text': item} for item in filtered_items]
        rv.viewclass = 'SelectableLabel'
        content.add_widget(rv)

        if hasattr(self, 'popup') and self.popup:
            self.popup.content = content
        else:
            self.popup = Popup(title='Select an option', content=content, size_hint=(0.8, 0.8))
            self.popup.open()

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
            label = MDLabel(text=message, halign="left", padding=[20, 0], size_hint_y=None, size_hint_x=1)
            label.bind(texture_size=lambda instance, value: setattr(instance, 'height', value[1]))
            self.root.ids.content_box.add_widget(label)



FormMain().run()