import json
import os
import shutil

from openpyxl import load_workbook

import engineering_change_order_gui as ECOG
import engineering_change_order as ECO
import config

class EngineeringChangeOrderManager:
    def __init__(self, parent, engineering_change_order_form, config):
        self.config = config

        self.engineering_change_order_form = engineering_change_order_form(parent, self.organize_engineering_change_order, self.print_engineering_change_order)

        self.engineering_change_order = ECO.EngineeringChangeOrder()
        self.engineering_change_order.set_progress_callback(self.update_progress)

    def organize_engineering_change_order(self, event=None):
        # Get number from GUI
        part_number = self.engineering_change_order_form.get_part_number()
        
        # Set number in the model
        self.engineering_change_order.set_part_number(part_number)
        
        query_results, drawing_missing_list = self.engineering_change_order.query_customer_order_table(self.config.get_server(), self.config.get_database(), self.config.get_user_id(), self.config.get_user_pwd())
        organized_engineering_change_order_data = self.engineering_change_order.organize_engineering_change_order_data(query_results)

        self.engineering_change_order.make_compression_list()

        if drawing_missing_list:
            self.engineering_change_order_form.display_organize_error_message(drawing_missing_list, not_shipped_not_ordered_list)

        # Update engineering change order form
        self.engineering_change_order_form.display_engineering_change_order_data(organized_engineering_change_order_data, self.engineering_change_order.get_compression_list(), self.engineering_change_order.get_comp_code_chart())

    def print_engineering_change_order(self, event=None):
        drawings_path = self.config.get_drawings_folder()

        compression_list = self.engineering_change_order_form.get_selected_compression_sizes()
        print_list = self.engineering_change_order_form.get_print_vars()
        drawing_not_found_list, unable_to_print_drawing_list = self.engineering_change_order.print_shop_copy(drawings_path, compression_list, print_list)
        self.engineering_change_order_form.focus_entry_field()
        self.engineering_change_order_form.display_print_error_message(drawing_not_found_list, unable_to_print_drawing_list)
        self.engineering_change_order.reset_error_lists()
        
        #Create engineering change order draft email
        self.engineering_change_order.create_email()

    def update_progress(self, progress):
        self.engineering_change_order_form.update_progress_bar(progress)
