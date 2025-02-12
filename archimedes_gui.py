# Archimedes GUI by Blake Blount

import tkinter as tk
from tkinter import ttk

#from pattern_work_order_gui import PatternWorkOrderForm
#from pattern_receipt_gui import PatternReceiptForm
from shop_copy_gui import ShopCopyForm

class ArchimedesGUI:
    def __init__(self, archimedes_manager):
        self.archimedes_manager = archimedes_manager
        
        self.root = tk.Tk()
        self.root.title('Archimedes')
        self.root.geometry("235x90")
        #self.root.resizable(False, False)

        self.create_widgets()

    def create_widgets(self):
        # Create buttons for each form
        self.shop_copy_form_button = ttk.Button(self.root, text="Create Shop Copy", command=self.archimedes_manager.initialize_shop_copy_tool, width=35)
        self.engineering_change_order_form_button = ttk.Button(self.root, text="Create ECO", command=self.archimedes_manager.initialize_engineering_change_order_tool, width=35)

        # Pack the buttons
        self.shop_copy_form_button.grid(padx=10, pady=10)
        self.engineering_change_order_form_button.grid(padx=10, pady=10)
        #self.shop_copy_form_button.bind('<Return>', self.archimedes_manager.initialize_shop_copy_tool)
        #self.shop_copy_form_button.focus_set()
        #self.tasks_form.pack(padx=10, pady=3)

#    def open_shop_copy_form(self, shop_copy_manager):
#        shop_copy_form = ShopCopyForm(self.root, shop_copy_manager)
#        self.root.wait_window(shop_copy_form.top)
#        print("Open shop copy form method ran")
#        return shop_copy_form
#
#    def open_engineering_change_order_form(self):
#        pass
