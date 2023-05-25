import tkinter as tk
from tkinter import ttk

from shop_copy_manager import ShopCopyManager

class ShopCopyForm:
    def __init__(self, parent, organize_shop_copy_callback, print_shop_copy_callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Shop Copy Form")

        # Assign callbacks
        self.organize_shop_copy_callback = organize_shop_copy_callback
        self.print_shop_copy_callback = print_shop_copy_callback

        # Initialize table label list
        self.table_labels = []

        # Selected values for compressions (if applicable)
        self.selected_compression_sizes = {}

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

    def display_shop_copy_data(self, shop_copy_data_table, compression_list=None):
        # Destroy existing labels
        for label in self.table_labels:
            label.destroy()

        # Reset the list
        self.table_labels = []

        # Reset compression values
        self.selected_compression_sizes = {}

        # Display table headers
        headers = ["Part Number", "Line Item(s)", "Quantity"]
        if compression_list is not None:
            headers.append("Cable Info")
        for column, header in enumerate(headers):
            table_label = ttk.Label(self.top, text=header)
            table_label.grid(row=2, column=column)
            self.table_labels.append(table_label)

        # Populate data table
        for i, row in enumerate(shop_copy_data_table):
            for j, cell in enumerate(row):
                table_label = ttk.Label(self.top, text=cell)
                table_label.grid(row=i + 3, column=j)
                self.table_labels.append(table_label)
            if compression_list is not None and row[0] in compression_list:
                var = tk.StringVar()
                dropdown = ttk.Combobox(self.top, textvariable=var)
                dropdown.grid(row=i+3, column = len(row))
                self.selected_compression_sizes[row[0]] = var

        # Make print button active
        self.print_button.config(state=tk.NORMAL)

        # Update window size
        self.top.geometry("")
        self.top.update_idletasks()
        self.top.geometry(self.top.geometry())

    def get_order_number(self):
        return self.customer_order_number_var.get()
    
    def get_selected_compression_sizes(self):
        return {part_number: var.get() for part_number, var in self.selected_compression_sizes.items()}
