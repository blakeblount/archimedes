# Archimedes GUI by Blake Blount

import tkinter as tk
from tkinter import ttk

from pattern_work_order_gui import PatternWorkOrderForm
from pattern_receipt_gui import PatternReceiptForm
from shop_copy_gui import ShopCopyForm

class ArchimedesGUI:
    def __init__(self, archimedes_manager):
        self.archimedes_manager = archimedes_manager
        
        self.root = tk.Tk()
        self.root.title('Archimedes')
        self.root.geometry("235x125")
        self.root.resizable(False, False)

        self.create_widgets()

    def create_widgets(self):
        # Create buttons for each form
        self.pattern_work_order_form_button = ttk.Button(self.root, text=" Create Pattern Work Order", command=self.archimedes_manager.initialize_pattern_work_order_tool, width=35)
        self.pattern_receipt_form_button = ttk.Button(self.root, text="Create Pattern Receipt", command=self.archimedes_manager.initialize_pattern_receipt_tool, width=35)
        self.shop_copy_form_button = ttk.Button(self.root, text="Create Shop Copy", command=self.archimedes_manager.initialize_shop_copy_tool, width=35)
        self.tasks_form = ttk.Button(self.root, text="Tasks", command=self.archimedes_manager.initialize_tasks_tool, width=35)

        # Pack the buttons
        self.pattern_work_order_form_button.pack(padx=10, pady=3)
        self.pattern_receipt_form_button.pack(padx=10, pady=3)
        self.shop_copy_form_button.pack(padx=10, pady=3)
        self.tasks_form.pack(padx=10, pady=3)

    def open_pattern_work_order_form(self):
        pattern_work_order_form = PatternWorkOrderForm(self.root)
        self.root.wait_window(pattern_work_order_form.top)
        return pattern_work_order_form

    def open_pattern_receipt_form(self):
        pattern_receipt_form = PatternReceiptForm(self.root)
        self.root.wait_window(pattern_receipt_form.top)
        return pattern_receipt_form

    def open_shop_copy_form(self, shop_copy_manager):
        shop_copy_form = ShopCopyForm(self.root, shop_copy_manager)
        self.root.wait_window(shop_copy_form.top)
        return shop_copy_form

    def open_tasks_form(self):
        pass
