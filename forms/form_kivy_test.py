from kivy.lang import Builder
from kivy.metrics import dp

from kivymd.app import MDApp
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.snackbar import MDSnackbar
from kivymd.uix.label import MDLabel

KV = '''
MDBoxLayout:
    orientation: "vertical"

    MDTopAppBar:
        title: "MDTopAppBar"
        left_action_items: [["menu", lambda x: app.callback(x)]]
        right_action_items: [["dots-vertical", lambda x: app.callback(x)]]

    MDLabel:
        text: "Content"
        halign: "center"
'''


class Test(MDApp):
    def build(self):
        self.theme_cls.primary_palette = "Orange"
        self.theme_cls.theme_style = "Dark"
        menu_items = [
            {
                "text": "connect odoo",
                "on_release": lambda: self.connect_odoo(),
            },
            {
                "text": "connect autocad",
                "on_release": lambda: self.connect_autocad(),
            },
            {
                "text": "get parameters from odoo",
                "on_release": lambda: self.get_parameters_from_odoo(),
            },
            {
                "text": "push to BOQ",
                "on_release": lambda: self.push_to_boq(),
            },
            {
                "text": "transfer BOQ to PR",
                "on_release": lambda: self.transfer_boq_to_pr(),
            },
        ]
        self.menu = MDDropdownMenu(items=menu_items)
        return Builder.load_string(KV)

    def callback(self, button):
        self.menu.caller = button
        self.menu.open()

    def connect_odoo(self):
        self.menu.dismiss()
        self.show_snackbar("Connecting to Odoo...")

    def connect_autocad(self):
        self.menu.dismiss()
        self.show_snackbar("Connecting to AutoCAD...")

    def get_parameters_from_odoo(self):
        self.menu.dismiss()
        self.show_snackbar("Getting parameters from Odoo...")

    def push_to_boq(self):
        self.menu.dismiss()
        self.show_snackbar("Pushing to BOQ...")

    def transfer_boq_to_pr(self):
        self.menu.dismiss()
        self.show_snackbar("Transferring BOQ to PR...")

    def show_snackbar(self, text_item):
        snackbar = MDSnackbar(MDLabel(text=text_item), duration=1)
        snackbar.open()


Test().run()
