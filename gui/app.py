import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pathlib
from datetime import datetime

from gui.components import GUIComponents
from core.pdf_processor import PDFProcessor
from core.excel_generator import ExcelGenerator
from core.utils import validate_po_number_format, extract_project_number
from config.settings import load_user_settings, save_user_settings

class POGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Purchase Order Generator")
        self.root.geometry("600x500")
        
        # Initialize processors
        self.pdf_processor = PDFProcessor()
        self.excel_generator = ExcelGenerator()
        
        # Variables
        self.po_number = tk.StringVar()
        self.project_name = tk.StringVar()
        self.purchaser_name = tk.StringVar()
        self.phone_code = tk.StringVar(value="+60")
        self.phone_number_only = tk.StringVar()
        self.director_manager = tk.StringVar()
        self.quotation_file = tk.StringVar()
        self.remember_details = tk.BooleanVar(value=False)
        
        # Track if we're currently loading settings (to avoid auto-save during load)
        self.is_loading_settings = False
        
        # Load saved settings
        self.user_settings = load_user_settings()
        self.load_saved_settings()
        
        # Set up auto-save tracking
        self.setup_auto_save()
        
        self.create_widgets()
        
    def setup_auto_save(self):
        """Set up trace to auto-save when fields change and remember_details is checked"""
        # Track changes to form fields
        self.po_number.trace_add('write', self.auto_save_if_enabled)
        self.project_name.trace_add('write', self.auto_save_if_enabled)
        self.purchaser_name.trace_add('write', self.auto_save_if_enabled)
        self.phone_code.trace_add('write', self.auto_save_if_enabled)
        self.phone_number_only.trace_add('write', self.auto_save_if_enabled)
        self.director_manager.trace_add('write', self.auto_save_if_enabled)
        self.remember_details.trace_add('write', self.on_remember_details_changed)
        
    def auto_save_if_enabled(self, *args):
        """Auto-save current settings if remember_details is checked and not loading"""
        if self.remember_details.get() and not self.is_loading_settings:
            self.save_current_settings(silent=True)
            
    def on_remember_details_changed(self, *args):
        """Handle remember_details checkbox changes"""
        if self.remember_details.get():
            # Checkbox checked - save current settings with remember_details=True
            if self.save_current_settings(silent=True):
                self.status_label.config(text="Auto-save enabled - details will be remembered", foreground="green")
        else:
            # Checkbox unchecked - save current settings with remember_details=False
            # This ensures the setting is persisted as false
            if self.save_current_settings(silent=True):
                self.status_label.config(text="Auto-save disabled - details will not be remembered", foreground="orange")
            # Clear the status message after a delay
            self.root.after(3000, lambda: self.status_label.config(text=""))
        
    def load_saved_settings(self):
        """Load saved settings into the form"""
        self.is_loading_settings = True
        
        if self.user_settings.get('remember_details', False):
            self.po_number.set(self.user_settings.get('po_number', ''))
            self.project_name.set(self.user_settings.get('project_name', ''))
            self.purchaser_name.set(self.user_settings.get('purchaser_name', ''))
            self.phone_code.set(self.user_settings.get('phone_code', '+60'))
            self.phone_number_only.set(self.user_settings.get('phone_number_only', ''))
            self.director_manager.set(self.user_settings.get('director_manager', ''))
            self.remember_details.set(True)
        else:
            # Ensure the checkbox reflects the false state from settings
            self.remember_details.set(False)
            self.save_current_settings(silent=True)
            
        self.is_loading_settings = False
        
    def save_current_settings(self, silent=False):
        """Save current form settings"""
        settings = {
            'po_number': self.po_number.get(),
            'project_name': self.project_name.get(),
            'purchaser_name': self.purchaser_name.get(),
            'phone_code': self.phone_code.get(),
            'phone_number_only': self.phone_number_only.get(),
            'director_manager': self.director_manager.get(),
            'remember_details': self.remember_details.get()  # This will be false when unchecked
        }
        
        if save_user_settings(settings):
            self.user_settings = settings
            if not silent:
                if self.remember_details.get():
                    self.status_label.config(text="Details saved successfully! Auto-save enabled.", foreground="green")
                else:
                    self.status_label.config(text="Details saved. Auto-save disabled.", foreground="orange")
            return True
        else:
            if not silent:
                self.status_label.config(text="Error saving details", foreground="red")
            return False
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Purchase Order Generator", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Form components
        row_counter = 1
        
        # PO Number
        GUIComponents.create_labeled_entry(
            main_frame, "PO Number:", self.po_number, row_counter,
            "Format: P-######-###M (e.g., P-250719-001M)"
        )
        row_counter += 2
        
        # Project Name
        GUIComponents.create_labeled_entry(main_frame, "Project Name:", self.project_name, row_counter)
        row_counter += 1
        
        # PO Issue Date
        self.date_entry = GUIComponents.create_date_picker(main_frame, "PO Issue Date:", row_counter)
        row_counter += 1
        
        # Purchaser Name
        GUIComponents.create_labeled_entry(main_frame, "Purchaser Name:", self.purchaser_name, row_counter)
        row_counter += 1
        
        # Phone Number
        GUIComponents.create_phone_input(
            main_frame, "Purchaser Telephone:", 
            self.phone_code, self.phone_number_only, row_counter
        )
        row_counter += 1
        
        # Director/Manager
        GUIComponents.create_labeled_entry(main_frame, "Manager Name:", self.director_manager, row_counter)
        row_counter += 1
        
        # Quotation File
        file_frame, _ = GUIComponents.create_file_browser(
            main_frame, "Quotation PDF:", self.quotation_file, row_counter,
            [("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.RIGHT, padx=5)
        row_counter += 1
        
        # Remember Details Checkbox
        remember_frame = ttk.Frame(main_frame)
        remember_frame.grid(row=row_counter, column=0, columnspan=3, sticky=tk.W, pady=10)
        row_counter += 1
        
        ttk.Checkbutton(
            remember_frame, 
            text="Remember details for next time (auto-save)", 
            variable=self.remember_details
        ).pack(side=tk.LEFT)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row_counter, column=0, columnspan=3, pady=20)
        row_counter += 1
        
        ttk.Button(button_frame, text="Generate Purchase Order", 
                  command=self.generate_po).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Form", 
                  command=self.clear_form).pack(side=tk.LEFT, padx=5)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="", foreground="blue")
        self.status_label.grid(row=row_counter, column=0, columnspan=3, pady=10)
        row_counter += 1
        
        # Save As button
        self.save_as_button = ttk.Button(
            main_frame, 
            text="Save As...", 
            command=self.save_as_dialog,
            state="disabled"
        )
        self.save_as_button.grid(row=row_counter, column=0, columnspan=3, pady=5)
        self.save_as_button.grid_remove()
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Quotation PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if filename:
            self.quotation_file.set(filename)
            
    def clear_form(self):
        # Clear all fields except the remember details setting
        remember_setting = self.remember_details.get()
        
        self.po_number.set("")
        self.project_name.set("")
        self.date_entry.set_date(datetime.now())
        self.purchaser_name.set("")
        self.phone_code.set("+60")
        self.phone_number_only.set("")
        self.director_manager.set("")
        self.quotation_file.set("")
        self.remember_details.set(remember_setting)  # Keep the remember setting
        
        # Auto-save if remember is enabled
        if remember_setting:
            self.save_current_settings(silent=True)
            
        self.status_label.config(text="Form cleared")
        self.hide_save_button()
        
    def hide_save_button(self):
        self.save_as_button.grid_remove()
        self.save_as_button.config(state="disabled")
        
    def show_save_button(self):
        self.save_as_button.grid()
        self.save_as_button.config(state="normal")
        
    def validate_inputs(self):
        required_fields = [
            ("PO Number", self.po_number.get()),
            ("Project Name", self.project_name.get()),
            ("Purchaser Name", self.purchaser_name.get()),
            ("Quotation PDF", self.quotation_file.get())
        ]
        
        missing_fields = []
        for field_name, field_value in required_fields:
            if not field_value.strip():
                missing_fields.append(field_name)
                
        if missing_fields:
            messagebox.showerror(
                "Missing Information", 
                f"Please fill in the following required fields:\n• " + "\n• ".join(missing_fields)
            )
            return False
            
        if not validate_po_number_format(self.po_number.get()):
            messagebox.showerror(
                "Validation Error",
                "The **PO Number** format is incorrect.\n"
                "It must be in the format: P-######-###M (e.g., P-250719-001M)"
            )
            return False
        
        if not os.path.exists(self.quotation_file.get()):
            messagebox.showerror("File Error", "The selected quotation file does not exist.")
            return False
            
        return True
    
    def save_as_dialog(self):
        if not self.excel_generator.temp_filepath or not os.path.exists(self.excel_generator.temp_filepath):
            messagebox.showerror("Error", "No generated file found. Please generate the PO first.")
            return
            
        default_filename = f"{self.po_number.get()}.xlsx" if self.po_number.get() else f"PO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        filepath = filedialog.asksaveasfilename(
            title="Save Purchase Order As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=default_filename
        )
        
        if filepath:
            try:
                import shutil
                shutil.copy2(self.excel_generator.temp_filepath, filepath)
                
                self.status_label.config(
                    text=f"Purchase Order saved successfully!\nLocation: {filepath}", 
                    foreground="green"
                )
                messagebox.showinfo("Success", f"Purchase Order saved successfully!\n\nLocation: {filepath}")
                
                self.excel_generator.cleanup_temp_file()
                self.hide_save_button()
                
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save file: {str(e)}")
        else:
            self.status_label.config(text="Save cancelled - file ready for download", foreground="orange")

    def generate_po(self):
        if not self.validate_inputs():
            return
            
        try:
            self.status_label.config(text="Generating Purchase Order...", foreground="blue")
            self.hide_save_button()
            self.root.update()
            
            # Prepare GUI data
            gui_data = {
                'po_number': self.po_number.get(),
                'project_number': extract_project_number(self.po_number.get()),
                'project_name': self.project_name.get(),
                'po_issue_date': self.date_entry.get_date().strftime("%d/%m/%Y"),
                'purchaser_name': self.purchaser_name.get(),
                'purchaser_phone': self.phone_code.get().strip() + self.phone_number_only.get().strip() if self.phone_number_only.get() else "",
                'director_manager': self.director_manager.get(),
                'quotation_file': self.quotation_file.get()
            }
            
            # Process PDF and generate Excel
            filepath = pathlib.Path(gui_data['quotation_file'])
            extracted_data = self.pdf_processor.extract_po_data(filepath, gui_data)
            self.excel_generator.generate_po_excel(extracted_data)
            
            # Show success and save button
            self.status_label.config(text="Purchase Order generated successfully!", foreground="green")
            self.show_save_button()

        except Exception as e:
            self.status_label.config(text="Error generating PO", foreground="red")
            messagebox.showerror("Error", f"An error occurred while generating the PO:\n{str(e)}")