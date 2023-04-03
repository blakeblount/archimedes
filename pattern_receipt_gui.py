from datetime import datetime

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

from pattern_receipt_manager import PatternReceiptManager

class PatternReceiptForm:
    def __init__(self, parent):
        self.top = tk.Toplevel(parent)
        self.top.title("Pattern Receipt Form")
        self.top.geometry("305x320")
        self.top.resizable(False, False)
        
        self.pattern_receipt_manager = PatternReceiptManager()

        self.create_widgets()

    def create_widgets(self):
        # Pattern Name input
        ttk.Label(self.top, text="Pattern Name:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.pattern_name_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.pattern_name_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        # Drawing Number input
        ttk.Label(self.top, text="Drawing Number:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.drawing_number_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.drawing_number_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        # Impressions input
        ttk.Label(self.top, text="Impressions:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.impressions_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.impressions_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        # Pattern Type dropdown
        ttk.Label(self.top, text="Pattern Type:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.pattern_type_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Board Mounted", "Matchplate", "Sinto", "Loose", "Other"), textvariable=self.pattern_type_var
        ).grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        # Core Box Doolean dropdown
        ttk.Label(self.top, text="Core Box:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.core_box_boolean_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Yes", "No"), textvariable=self.core_box_boolean_var
        ).grid(row=4, column=1, sticky="ew", padx=5, pady=5)

        # Core Box Name input
        ttk.Label(self.top, text="Core Box Name:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.core_box_name_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.core_box_name_var).grid(row=5, column=1, sticky="ew", padx=5, pady=5)

        # Core Box Type dropdown
        ttk.Label(self.top, text="Core Box Type:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        self.core_box_type_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Dump Box", "Shell Box", "Other"), textvariable=self.core_box_type_var
        ).grid(row=6, column=1, sticky="ew", padx=5, pady=5)

        # Shipping Party dropdown
        ttk.Label(self.top, text="Shipping Party:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        self.shipping_party_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Doug Smith", "Blake Blount", "Craig Light", "Mike Willford", "D'Andrea Howard"), textvariable=self.shipping_party_var
        ).grid(row=7, column=1, sticky="ew", padx=5, pady=5)

        # Receiving Party dropdown
        ttk.Label(self.top, text="Receiving Party:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        self.receiving_party_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Talladega Pattern and Aluminum Works (V#120)", "D and J Manufacturing (V#31)", "Alumicast (V#663)", "B R Metals (V#952)", "Intertool (V#66)", "Texas Metal Casting (V#382)"), textvariable=self.receiving_party_var
        ).grid(row=7, column=1, sticky="ew", padx=5, pady=5)

        # Notes input
        ttk.Label(self.top, text="Notes:").grid(row=8, column=0, sticky="w", padx=5, pady=5)
        self.notes_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.notes_var).grid(row=8, column=1, sticky="ew", padx=5, pady=5)

        # Create PWO button
        ttk.Button(self.top, text="Create Pattern Receipt", command=self.create_pattern_receipt).grid(row=9, column=1, padx=5, pady=5, sticky="e")


    def create_pattern_receipt(self):
        user_input = {
            "pattern_name": self.pattern_name_var.get(),
            "drawing_number": self.drawing_number_var.get(),
            "impressions": self.impressions_var.get(),
            "pattern_type": self.pattern_type_var.get(),
            "core_box_boolean": self.core_box_boolean_var.get(),
            "core_box_name": self.core_box_name_var.get(),
            "core_box_type": self.core_box_type_var.get(),
            "shipping_party": self.shipping_party_var.get(),
            "receiving_party": self.receiving_party_var.get(),
            "notes": self.notes_var.get()
        }
        
        self.pattern_receipt_manager.create_pattern_receipt(user_input)
        tk.messagebox.showinfo("Success", "Pattern Receipt created successfully")
        self.top.destroy()
