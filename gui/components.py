import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime

class GUIComponents:
    @staticmethod
    def create_labeled_entry(parent, label_text, variable, row, hint_text=None):
        """Create a labeled entry widget with optional hint"""
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=5)
        entry = ttk.Entry(parent, textvariable=variable, width=40)
        entry.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        if hint_text:
            hint_label = ttk.Label(
                parent, 
                text=hint_text, 
                foreground="gray",
                font=("Arial", 8, "italic")
            )
            hint_label.grid(row=row+1, column=1, columnspan=2, sticky=tk.W, padx=2)
            
        return entry

    @staticmethod
    def create_date_picker(parent, label_text, row):
        """Create a date picker with label"""
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=5)
        
        date_frame = ttk.Frame(parent)
        date_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        date_entry = DateEntry(
            date_frame, 
            width=18, 
            background='darkblue',
            foreground='white', 
            borderwidth=2, 
            date_pattern='dd/mm/yyyy'
        )
        date_entry.set_date(datetime.now())
        date_entry.pack(side=tk.LEFT)
        
        today_btn = ttk.Button(date_frame, text="Today", command=lambda: date_entry.set_date(datetime.now()))
        today_btn.pack(side=tk.LEFT, padx=5)
        
        return date_entry

    @staticmethod
    def create_phone_input(parent, label_text, code_var, number_var, row):
        """Create phone number input with country code"""
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=5)
        
        phone_frame = ttk.Frame(parent)
        phone_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        country_codes = ["+60", "+1", "+44", "+91", "+65", "+86", "+81", "+49"]
        
        code_combo = ttk.Combobox(
            phone_frame, 
            textvariable=code_var, 
            values=country_codes, 
            width=5, 
            state="readonly"
        )
        code_combo.pack(side=tk.LEFT, padx=(0, 5))
        
        number_entry = ttk.Entry(phone_frame, textvariable=number_var, width=33)
        number_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        return code_combo, number_entry

    @staticmethod
    def create_file_browser(parent, label_text, variable, row, filetypes):
        """Create file browser with entry and browse button"""
        ttk.Label(parent, text=label_text).grid(row=row, column=0, sticky=tk.W, pady=5)
        
        file_frame = ttk.Frame(parent)
        file_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        entry = ttk.Entry(file_frame, textvariable=variable, width=30)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        return file_frame, entry

    @staticmethod
    def create_checkbox(parent, text, variable, row, columnspan=3):
        """Create a checkbox"""
        checkbox = ttk.Checkbutton(parent, text=text, variable=variable)
        checkbox.grid(row=row, column=0, columnspan=columnspan, sticky=tk.W, pady=5)
        return checkbox