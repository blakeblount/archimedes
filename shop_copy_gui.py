import tkinter as tk
from tkinter import ttk

from shop_copy_manager import ShopCopyManager


class ShopCopyForm:
    def __init__(self, parent, organize_shop_copy_callback, print_shop_copy_callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Shop Copy Form")
#        self.top.iconbitmap('archimedes.ico')

        # Create a canvas and a scrollbar and connect them
        self.canvas = tk.Canvas(self.top)
        self.scrollbar = ttk.Scrollbar(self.top, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Assign callbacks
        self.organize_shop_copy_callback = organize_shop_copy_callback
        self.print_shop_copy_callback = print_shop_copy_callback

        # Initialize table label list
        self.table_labels = []

        # Initialize compression combobox list
        self.comboboxes = []
        self.headers = []

        # Selected values for compressions (if applicable)
        self.selected_compression_sizes = {}

        # Initialize progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.top, length=200, mode='determinate', variable=self.progress_var)

        self.create_widgets()

    def create_widgets(self): 
        # Customer Order Number input
        ttk.Label(self.top, text="Customer Order Number:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.customer_order_number_var = tk.StringVar()
        self.customer_order_field = ttk.Entry(self.top, textvariable=self.customer_order_number_var)
        self.customer_order_field.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.customer_order_field.focus_set()
        self.customer_order_field.bind('<Return>', self.organize_shop_copy_callback)

        # Create Organize and Print buttons
        self.organize_button = ttk.Button(self.top, text="Organize Shop Copy", command=self.organize_shop_copy_callback)
        self.organize_button.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.organize_button.bind('<Return>', self.organize_shop_copy_callback)
        self.print_button = ttk.Button(self.top, text="Print Shop Copy", command=self.print_shop_copy_callback, state=tk.DISABLED)
        self.print_button.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.print_button.bind('<Return>', self.print_shop_copy_callback)
        
        # Pack the canvas and the scrollbar
        self.canvas.grid(row=2, column=0, columnspan=2, sticky='nsew') # Changed from pack() to grid()
        self.scrollbar.grid(row=2, column=2, sticky='ns') # Changed from pack() to grid()

        # Add mouse wheel scrolling
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        # Create and pack the progress bar
        self.progress_bar.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky='ew')

        # Configure column and row weights so the canvas expands to fill space
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grid_rowconfigure(2, weight=1)

    def display_shop_copy_data(self, shop_copy_data_table, compression_list, comp_code_chart):
        # Destroy existing labels and comboboxes
        for label in self.table_labels:
            label.destroy()
        for combobox in self.comboboxes:
            combobox.destroy()
        for header in self.headers:
            header.destroy()

        self.update_progress_bar(0)

        # Reset the lists
        self.table_labels = []
        self.comboboxes = []
        self.headers = []

        # Reset compression values
        self.selected_compression_sizes = {}

        # Display table headers
        headers = ["Part Number", "Line Item(s)", "Quantity"]
        if compression_list is not None:
            headers.append("Cable Info")
            is_list = any(isinstance(v, list) and len(v) == 2 for v in compression_list.values())
            if is_list:
                headers.append("Tap Cable Info")
        for column, header in enumerate(headers):
            table_header = ttk.Label(self.scrollable_frame, text=header)
            table_header.grid(row=2, column=column, padx=3, pady=3)
            self.headers.append(table_header)

        # Populate data table
        for i, row in enumerate(shop_copy_data_table):
            for j, cell in enumerate(row):
                table_label = ttk.Label(self.scrollable_frame, text=cell)
                table_label.grid(row=i + 3, column=j)
                self.table_labels.append(table_label)
            if compression_list is not None and row[0] in compression_list:
                var = tk.StringVar()
                dropdown = ttk.Combobox(self.scrollable_frame, textvariable=var)
                if isinstance(compression_list[row[0]], list):
                    dropdown['values'] = comp_code_chart[compression_list[row[0]][0]]
                else:
                    dropdown['values'] = comp_code_chart[compression_list[row[0]]]
                dropdown.grid(row=i+3, column=len(row))
                self.comboboxes.append(dropdown)
                self.selected_compression_sizes[row[0]] = var

                if isinstance(compression_list[row[0]], list):
                    var_tap = tk.StringVar()
                    dropdown_tap = ttk.Combobox(self.scrollable_frame, textvariable=var_tap)
                    dropdown_tap['values'] = comp_code_chart[compression_list[row[0]][1]]
                    dropdown_tap.grid(row=i+3, column=len(row)+1)
                    self.comboboxes.append(dropdown_tap)
                    self.selected_compression_sizes[row[0] + '_tap'] = var_tap

        # Make print button active
        self.print_button.config(state=tk.NORMAL)
        self.print_button.focus_set()

        # Update window size
        self.top.geometry("")
        self.top.update_idletasks()
        self.top.geometry(self.top.geometry())

    def get_order_number(self):
        return self.customer_order_number_var.get()
    
    def get_selected_compression_sizes(self):
        result = {}
        for part_number, var in self.selected_compression_sizes.items():
            value = var.get()
            # Check if this part_number was from a second dropdown
            if '_tap' in part_number:
                original_part_number = part_number.replace('_tap', '')
                # If the original part_number is already in the result, append the value to it
                if original_part_number in result:
                    if isinstance(result[original_part_number], list):
                        result[original_part_number].append(value)
                    else:
                        result[original_part_number] = [result[original_part_number], value]
                # If the original part_number is not in the result, add it as a single value
                else:
                    result[original_part_number] = value
            # If this part_number was from a first dropdown
            else:
                # If the part_number is already in the result as a list, it means the second dropdown value is already there. So, add this value to the start of the list
                if part_number in result and isinstance(result[part_number], list):
                    result[part_number].insert(0, value)
                # If the part_number is not in the result, add it as a single value
                else:
                    result[part_number] = value
        return result

    def update_progress_bar(self, progress):
        self.progress_var.set(progress)
        self.top.update_idletasks()

    def focus_entry_field(self):
        self.customer_order_field.focus_set()

    def on_mousewheel(self, event):
        self.canvas.yview_scroll(-1*(event.delta//120), "units")
