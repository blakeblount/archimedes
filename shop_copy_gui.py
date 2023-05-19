import tkinter as tk
from tkinter import ttk

from shop_copy_manager import ShopCopyManager

class ShopCopyForm:
    def __init__(self, parent, organize_shop_copy_callback, print_shop_copy_callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Shop Copy Form")
        #self.top.geometry("305x70")
        #self.top.resizable(False, False)

        # Assign callbacks
        self.organize_shop_copy_callback = organize_shop_copy_callback
        self.print_shop_copy_callback = print_shop_copy_callback

        self.create_widgets()

    def create_widgets(self): 
        # Customer Order Number input
        ttk.Label(self.top, text="Customer Order Number:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.customer_order_number_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.customer_order_number_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # Create Organize and Print buttons
        self.organize_button = ttk.Button(self.top, text="Organize Shop Copy", command=self.organize_shop_copy_callback)
        self.organize_button.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.print_button = ttk.Button(self.top, text="Print Shop Copy", command=self.print_shop_copy_callback, state=tk.DISABLED)
        self.print_button.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    def display_shop_copy_data(self, shop_copy_data_table):
        # Display table headers
        table_label = ttk.Label(self.top, text="Part Number")
        table_label.grid(row=2, column=0)
        table_label = ttk.Label(self.top, text="Line Item(s)")
        table_label.grid(row=2, column=1)
        table_label = ttk.Label(self.top, text="Quantity")
        table_label.grid(row=2, column=2)

        # Populate data table
        for i, row in enumerate(shop_copy_data_table):
            for j, cell in enumerate(row):
                table_label = ttk.Label(self.top, text=cell)
                table_label.grid(row=i + 3, column=j)

        # Make print button active
        self.print_button.config(state=tk.NORMAL)
        

    def get_order_number(self):
        return self.customer_order_number_var.get()
