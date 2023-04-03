import json
import os
import shutil

from openpyxl import load_workbook

import shop_copy_gui as SCG
import shop_copy as SC
import config

class ShopCopyManager:
    def __init__(self, parent, shop_copy_form, config):
        self.load_config()
        self.config = config

        self.shop_copy_form = shop_copy_form(parent, self.organize_shop_copy, self.print_shop_copy)
        self.shop_copy = SC.ShopCopy()

    def organize_shop_copy(self):
        # Get number from GUI
        order_number = self.shop_copy_form.get_order_number()
        
        # Set number in the model
        self.shop_copy.set_order_number(order_number)
        organized_shop_copy_data = self.shop_copy.organize_shop_copy_data()

        # Update shop copy form
        self.shop_copy_form.display_shop_copy_data(organized_shop_copy_data)

    def print_shop_copy(self):
        self.shop_copy.print_shop_copy()

    def load_config(self):
        with open("config.json", "r", encoding="utf-8") as config_file:
            config = json.load(config_file)

        file_paths = config["file_paths"]
        self.drawings_path = file_paths["drawings_folder"]
