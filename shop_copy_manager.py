import json
import os
import shutil

from openpyxl import load_workbook

import shop_copy_gui as SCG
import shop_copy as SC
import config

class ShopCopyManager:
    def __init__(self, parent, shop_copy_form, config):
        self.config = config

        self.shop_copy_form = shop_copy_form(parent, self.organize_shop_copy, self.print_shop_copy)
        self.shop_copy = SC.ShopCopy()
       
        self.shop_copy.set_progress_callback(self.update_progress)

    def organize_shop_copy(self, event=None):
        # Get number from GUI
        order_number = self.shop_copy_form.get_order_number()
        
        # Set number in the model
        self.shop_copy.set_order_number(order_number)
        
        query_results = self.shop_copy.query_customer_order_table(self.config.get_server(), self.config.get_database(), self.config.get_user_id(), self.config.get_user_pwd())
        organized_shop_copy_data = self.shop_copy.organize_shop_copy_data(query_results)

        self.shop_copy.make_compression_list()

        # Update shop copy form
        self.shop_copy_form.display_shop_copy_data(organized_shop_copy_data, self.shop_copy.get_compression_list(), self.shop_copy.get_comp_code_chart())

    def print_shop_copy(self, event=None):
        drawings_path = self.config.get_drawings_folder()

        compression_list = self.shop_copy_form.get_selected_compression_sizes()
        self.shop_copy.print_shop_copy(self.config.get_drawings_folder(), compression_list)
        self.shop_copy_form.focus_entry_field()

    def update_progress(self, progress):
        self.shop_copy_form.update_progress_bar(progress)

