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
        
        query_results, drawing_missing_list, not_shipped_not_ordered_list = self.shop_copy.query_customer_order_table(self.config.get_server(), self.config.get_database(), self.config.get_user_id(), self.config.get_user_pwd(), self.shop_copy_form.get_include_shipped_items_var())
        organized_shop_copy_data = self.shop_copy.organize_shop_copy_data(query_results)

        self.shop_copy.make_compression_list()

        if drawing_missing_list or not_shipped_not_ordered_list:
            self.shop_copy_form.display_organize_error_message(drawing_missing_list, not_shipped_not_ordered_list)

        # Update shop copy form
        self.shop_copy_form.display_shop_copy_data(organized_shop_copy_data, self.shop_copy.get_compression_list(), self.shop_copy.get_comp_code_chart())

    def print_shop_copy(self, event=None):
        drawings_path = self.config.get_drawings_folder()

        compression_list = self.shop_copy_form.get_selected_compression_sizes()
        print_list = self.shop_copy_form.get_print_vars()
        drawing_not_found_list, unable_to_print_drawing_list = self.shop_copy.print_shop_copy(self.config.get_drawings_folder(), compression_list, print_list)
        self.shop_copy_form.focus_entry_field()
        self.shop_copy_form.display_print_error_message(drawing_not_found_list, unable_to_print_drawing_list)
        self.shop_copy.reset_error_lists()

    def update_progress(self, progress):
        self.shop_copy_form.update_progress_bar(progress)

