from datetime import datetime

import tkinter as tk
from tkinter import ttk
from tkinter import messagebox

from pattern_work_order_manager import PatternWorkOrderManager


class PatternWorkOrderForm():
    def __init__(self, parent, create_pattern_work_order_callback):
        self.top = tk.Toplevel(parent)
        self.top.title("Pattern Work Order Form")
        self.top.geometry("305x285")
        self.top.resizable(False, False)

        self.create_pattern_work_order_callback = create_pattern_work_order_callback

        self.create_widgets()

    def create_widgets(self):
        # Pattern Maker dropdown
        ttk.Label(self.top, text="Pattern Maker:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.pattern_maker_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Intertool", "Pro Pattern"), textvariable=self.pattern_maker_var
        ).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        
        # Orderer dropdown
        ttk.Label(self.top, text="Orderer:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.orderer_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Doug Smith", "Blake Blount", "Craig Light", "Mike Willford", "D'Andrea Howard"), textvariable=self.orderer_var
        ).grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        # Part Name input
        ttk.Label(self.top, text="Part Name:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.part_name_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.part_name_var).grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        # Drawing Number input
        ttk.Label(self.top, text="Drawing Number:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.drawing_number_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.drawing_number_var).grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        # Work Requested dropdown
        ttk.Label(self.top, text="Work Requested:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.work_requested_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("New", "Modify", "Repair"), textvariable=self.work_requested_var
        ).grid(row=4, column=1, sticky="ew", padx=5, pady=5)

        # Type of Pattern dropdown
        ttk.Label(self.top, text="Type of Pattern:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.pattern_type_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Matchplate", "Sinto", "Board Mount", "Other"), textvariable=self.pattern_type_var
        ).grid(row=5, column=1, sticky="ew", padx=5, pady=5)

        # Intended Foundry dropdown
        ttk.Label(self.top, text="Intended Foundry:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        self.intended_foundry_var = tk.StringVar()
        ttk.Combobox(
            self.top, values=("Talladega", "D and J", "Alumicast", "B R Metals"), textvariable=self.intended_foundry_var
        ).grid(row=6, column=1, sticky="ew", padx=5, pady=5)

        # Notes input
        ttk.Label(self.top, text="Notes:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        self.notes_var = tk.StringVar()
        ttk.Entry(self.top, textvariable=self.notes_var).grid(row=7, column=1, sticky="ew", padx=5, pady=5)

        # Create PWO button
        ttk.Button(self.top, text="Create Pattern Work Order", command=self.create_pattern_work_order_callback).grid(row=8, column=1, padx=5, pady=5, sticky="e")

    def get_user_inputs(self):
        user_inputs = {
            "pattern_maker": self.pattern_maker_var.get(),
            "orderer": self.orderer_var.get(),
            "part_name": self.part_name_var.get(),
            "drawing_number": self.drawing_number_var.get(),
            "work_requested": self.work_requested_var.get(),
            "pattern_type": self.pattern_type_var.get(),
            "intended_foundry": self.intended_foundry_var.get(),
            "notes": self.notes_var.get()
        }

        return user_inputs

    def alert_success(self):
        tk.messagebox.showinfo("Success", "Pattern Work Order created successfully")
        self.top.destroy()
