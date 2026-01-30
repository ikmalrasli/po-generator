import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import pathlib
import subprocess
import sys
import time
import threading
from datetime import datetime

from gui.components import GUIComponents
from core.pdf_processor import PDFProcessor
from core.excel_generator import ExcelGenerator
from core.utils import validate_po_number_format, extract_project_number
from config.settings import load_user_settings, save_user_settings

class CustomSuccessDialog:
    def __init__(self, parent, time_text=""):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Success")
        self.dialog.geometry("400x150")
        self.dialog.resizable(False, False)
        
        # Center the dialog on parent
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Make dialog modal
        self.result = None
        
        self.create_widgets(time_text)
        
        # Center on screen
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (400 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (150 // 2)
        self.dialog.geometry(f"400x150+{x}+{y}")
        
    def create_widgets(self, time_text):
        # Main frame with padding
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Icon and message frame
        message_frame = ttk.Frame(main_frame)
        message_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Information icon (using Unicode symbol)
        icon_label = ttk.Label(message_frame, text="ℹ", font=("Arial", 16))
        icon_label.pack(side=tk.LEFT, padx=(0, 10))
        
        # Success message
        message_text = f"Purchase Order generated successfully!"
        if time_text:
            message_text += f"\n(Time: {time_text})"
        
        message_label = ttk.Label(message_frame, text=message_text, font=("Arial", 10))
        message_label.pack(side=tk.LEFT)
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        # Style buttons
        style = ttk.Style()
        style.configure("Success.TButton", font=("Arial", 9))
        
        # Open Folder button
        ttk.Button(
            button_frame, 
            text="Open Folder", 
            command=self.open_folder,
            style="Success.TButton",
            width=12
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # Save As button
        ttk.Button(
            button_frame, 
            text="Save As...", 
            command=self.save_as,
            style="Success.TButton",
            width=12
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # OK button
        ttk.Button(
            button_frame, 
            text="OK", 
            command=self.ok_clicked,
            style="Success.TButton",
            width=8
        ).pack(side=tk.RIGHT)
        
    def open_folder(self):
        """Open the temp folder where the file is saved"""
        import subprocess
        import sys
        
        temp_dir = "temp"
        if sys.platform == "win32":
            subprocess.run(["explorer", temp_dir])
        elif sys.platform == "darwin":
            subprocess.run(["open", temp_dir])
        else:
            subprocess.run(["xdg-open", temp_dir])
            
    def save_as(self):
        """Trigger save as action"""
        self.result = "save_as"
        self.dialog.destroy()
        
    def ok_clicked(self):
        """Close dialog"""
        self.result = "ok"
        self.dialog.destroy()

class POGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Purchase Order Generator")
        
        # Load saved settings
        self.user_settings = load_user_settings()
        
        # Initialize processors after loading settings
        self.pdf_processor = PDFProcessor(api_key=self.user_settings.get('google_api_key', ''))
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
        self.google_api_key = tk.StringVar()
        
        # Track saved file path
        self.saved_filepath = None
        
        # Track generation start time and state
        self.generation_start_time = None
        self.is_generating = False
        
        # Track if we're currently loading settings (to avoid auto-save during load)
        self.is_loading_settings = False
        
        # Load saved settings into the form
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
        self.google_api_key.trace_add('write', self.auto_save_if_enabled)
        
    def auto_save_if_enabled(self, *args):
        """Auto-save current settings if remember_details is checked and not loading"""
        if self.remember_details.get() and not self.is_loading_settings:
            self.save_current_settings(silent=True)
            
    def on_remember_details_changed(self, *args):
        """Handle remember_details checkbox changes"""
        if self.remember_details.get():
            # Checkbox checked - save current settings with remember_details=True
            if self.save_current_settings(silent=True):
                messagebox.showinfo("Auto-save", "Auto-save enabled - details will be remembered")
        else:
            # Checkbox unchecked - save current settings with remember_details=False
            if self.save_current_settings(silent=True):
                messagebox.showinfo("Auto-save", "Auto-save disabled - details will not be remembered")
                
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
        
        # Load API key separately (not dependent on remember_details)
        self.google_api_key.set(self.user_settings.get('google_api_key', ''))
        self.update_api_key_display()
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
            'remember_details': self.remember_details.get(),
            'google_api_key': self.google_api_key.get()
        }
        
        if save_user_settings(settings):
            self.user_settings = settings
            if not silent:
                if self.remember_details.get():
                    messagebox.showinfo("Success", "Details saved successfully! Auto-save enabled.")
                else:
                    messagebox.showinfo("Success", "Details saved. Auto-save disabled.")
            return True
        else:
            if not silent:
                messagebox.showerror("Error", "Error saving details")
            return False
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # Header with title and API status
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(0, weight=1)
        
        # Title
        title_label = ttk.Label(header_frame, text="Purchase Order Generator", 
                               font=("Arial", 18, "bold"))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # API Status in top-right
        api_status_frame = ttk.Frame(header_frame)
        api_status_frame.grid(row=0, column=1, sticky=tk.E)
        
        self.api_status_indicator = ttk.Label(api_status_frame, text="●", font=("Arial", 12))
        self.api_status_indicator.pack(side=tk.LEFT)
        
        self.api_status_text = ttk.Label(api_status_frame, text="API Key Missing", font=("Arial", 10))
        self.api_status_text.pack(side=tk.LEFT, padx=(5, 10))
        
        self.api_settings_button = ttk.Button(api_status_frame, text="⚙ Settings", 
                                             command=self.open_api_key_dialog, width=10)
        self.api_settings_button.pack(side=tk.RIGHT)
        
        # Form container with consistent padding
        form_container = ttk.Frame(main_frame)
        form_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        form_container.columnconfigure(0, weight=1)
        
        row_counter = 0
        
        # Group 1: Project Information
        project_group = self.create_group_frame(form_container, "Project Information", row_counter)
        row_counter += 1
        
        # PO Number
        po_entry, po_hint, next_row = self.create_labeled_entry(project_group, "PO Number:", self.po_number, 0,
                                 "Format: P-######-###M (e.g., P-250719-001M)")
        
        # Project Name
        self.create_labeled_entry(project_group, "Project Name:", self.project_name, next_row)
        
        # PO Issue Date
        self.date_entry = GUIComponents.create_date_picker(project_group, "PO Issue Date:", next_row + 1)
        
        # Group 2: Personnel Information
        personnel_group = self.create_group_frame(form_container, "Personnel Information", row_counter)
        row_counter += 1
        
        # Purchaser Name
        _, next_row = self.create_labeled_entry(personnel_group, "Purchaser Name:", self.purchaser_name, 0)
        
        # Phone Number
        ttk.Label(personnel_group, text="Purchaser Telephone:", font=("Arial", 9)).grid(row=next_row, column=0, sticky=tk.W, pady=(0, 2))
        
        # Create phone input frame for the entry field
        phone_input_frame = ttk.Frame(personnel_group)
        phone_input_frame.grid(row=next_row, column=1, sticky="ew", pady=(0, 2), padx=(5, 0))
        phone_input_frame.columnconfigure(1, weight=1)
        personnel_group.columnconfigure(1, weight=1)  # Ensure phone column expands
        
        self.phone_code_entry = ttk.Entry(phone_input_frame, textvariable=self.phone_code, width=5)
        self.phone_code_entry.pack(side=tk.LEFT)
        
        ttk.Label(phone_input_frame, text="-").pack(side=tk.LEFT, padx=(2, 2))
        
        self.phone_entry = ttk.Entry(phone_input_frame, textvariable=self.phone_number_only)
        self.phone_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Director/Manager
        self.create_labeled_entry(personnel_group, "Manager Name:", self.director_manager, next_row + 1)
        
        # Group 3: Attachments
        attachments_group = self.create_group_frame(form_container, "Attachments", row_counter)
        row_counter += 1
        
        # Quotation File - Single row layout
        ttk.Label(attachments_group, text="Quotation PDF:", font=("Arial", 9)).grid(row=0, column=0, sticky=tk.W, pady=(0, 2))
        
        # File entry and browse button on same row
        file_input_frame = ttk.Frame(attachments_group)
        file_input_frame.grid(row=0, column=1, sticky="ew", pady=(0, 2), padx=(5, 0))
        file_input_frame.columnconfigure(0, weight=1)
        attachments_group.columnconfigure(1, weight=1)  # Ensure file column expands
        
        self.file_entry = ttk.Entry(file_input_frame, textvariable=self.quotation_file)
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(file_input_frame, text="Browse", command=self.browse_file).pack(side=tk.RIGHT, padx=(5, 0))
        
        # Options section
        options_frame = ttk.Frame(form_container)
        options_frame.grid(row=row_counter, column=0, sticky=(tk.W, tk.E), pady=2)
        
        ttk.Checkbutton(
            options_frame, 
            text="Remember details for next time (auto-save)", 
            variable=self.remember_details
        ).pack(side=tk.LEFT)
        row_counter += 1
        
        # Action buttons - right aligned
        button_frame = ttk.Frame(form_container)
        button_frame.grid(row=row_counter, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Style the buttons with primary/secondary hierarchy
        style = ttk.Style()
        style.configure("Primary.TButton", font=("Arial", 9, "bold"))
        style.configure("Secondary.TButton", font=("Arial", 9))
        
        self.generate_button = ttk.Button(
            button_frame, 
            text="Generate Purchase Order", 
            command=self.generate_po,
            style="Primary.TButton",
            width=25
        )
        self.generate_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Clear Form button (left)
        ttk.Button(button_frame, text="Clear Form", 
                  command=self.clear_form,
                  style="Secondary.TButton",
                  width=12).pack(side=tk.LEFT)
        
        # Save As button (center, initially disabled)
        self.save_as_button = ttk.Button(
            button_frame, 
            text="Save As...", 
            command=self.save_as_dialog,
            style="Secondary.TButton",
            width=12,
            state="disabled"
        )
        self.save_as_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Open File button
        self.open_file_button = ttk.Button(
            main_frame,
            text="Open File",
            command=self.open_saved_file,
            state="disabled"
        )
        self.open_file_button.grid(row=4, column=0, pady=5)
        self.open_file_button.grid_remove()
        
        # Update API key display after all widgets are created
        self.update_api_key_display()
        
        # Auto-size window to fit content
        self.root.update_idletasks()
        self.root.geometry("")
        
    def create_group_frame(self, parent, title, row):
        """Create a visually grouped frame with title"""
        group_frame = ttk.LabelFrame(parent, text=title, padding="3")
        group_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=3)
        group_frame.columnconfigure(0, weight=1)
        return group_frame
    
    def create_labeled_entry(self, parent, label_text, variable, row, help_text=""):
        """Create a labeled entry with consistent styling"""
        ttk.Label(parent, text=label_text, font=("Arial", 9), width=15, anchor=tk.W).grid(row=row, column=0, sticky=tk.W, pady=(0, 2))
        
        entry = ttk.Entry(parent, textvariable=variable, font=("Arial", 9))
        entry.grid(row=row, column=1, sticky="ew", pady=(0, 2), padx=(5, 0))
        parent.columnconfigure(1, weight=1)  # Ensure entry column expands
        
        if help_text:
            help_label = ttk.Label(parent, text=help_text, foreground="gray", font=("Arial", 7))
            help_label.grid(row=row+1, column=1, sticky=tk.W, padx=(5, 0), pady=(0, 0))
            return entry, help_label, row + 2  # Return next available row
        
        return entry, row + 1  # Return next available row
        
    def update_elapsed_time(self):
        """Update the elapsed time during generation"""
        if self.is_generating and self.generation_start_time:
            elapsed_time = time.time() - self.generation_start_time
            minutes = int(elapsed_time // 60)
            seconds = int(elapsed_time % 60)
            
            if minutes > 0:
                time_text = f"{minutes}m {seconds}s"
            else:
                time_text = f"{seconds}s"
                
            # Update window title with progress
            self.root.title(f"Purchase Order Generator - Generating... ({time_text})")
            
            # Schedule next update if still generating
            if self.is_generating:
                self.root.after(1000, self.update_elapsed_time)
        
    def open_saved_file(self):
        """Open the saved file with the default application"""
        if self.saved_filepath and os.path.exists(self.saved_filepath):
            try:
                if sys.platform == "win32":
                    os.startfile(self.saved_filepath)
                elif sys.platform == "darwin":
                    subprocess.run(["open", self.saved_filepath])
                else:
                    subprocess.run(["xdg-open", self.saved_filepath])
                
                messagebox.showinfo("Success", "Opening file...")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open the file:\n{str(e)}")
        else:
            messagebox.showerror("Error", "No saved file found or file has been moved.")
            
    def hide_open_file_button(self):
        """Hide the Open File button"""
        self.open_file_button.grid_remove()
        
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
        self.remember_details.set(remember_setting)
        
        if remember_setting:
            self.save_current_settings(silent=True)
            
        messagebox.showinfo("Success", "Form cleared")
        self.hide_save_button()
        self.hide_open_file_button()
        
    def hide_save_button(self):
        self.save_as_button.config(state="disabled")
        
    def show_save_button(self):
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
                
                self.saved_filepath = filepath
                short_path = self._shorten_file_path(filepath)
                
                self.show_open_file_button()
                self.excel_generator.cleanup_temp_file()
                self.hide_save_button()
                
                messagebox.showinfo("Success", f"Purchase Order saved successfully!\n\nLocation: {filepath}")
                
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save file: {str(e)}")
                self.hide_open_file_button()
        else:
            messagebox.showinfo("Info", "Save cancelled - file ready for download")
            self.hide_open_file_button()

    def _shorten_file_path(self, filepath, max_length=60):
        """Shorten file path for display in status label"""
        if len(filepath) <= max_length:
            return filepath
        
        path = Path(filepath)
        filename = path.name
        parent = path.parent.name
        
        shortened = f".../{parent}/{filename}"
        
        if len(shortened) > max_length:
            shortened = f".../{filename}"
            
        return shortened

    def _generate_po_thread(self, gui_data):
        """Run the PO generation in a separate thread"""
        try:
            # Process PDF and generate Excel
            filepath = pathlib.Path(gui_data['quotation_file'])
            extracted_data = self.pdf_processor.extract_po_data(filepath, gui_data)
            self.excel_generator.generate_po_excel(extracted_data)
            
            # Stop timing and calculate final elapsed time
            elapsed_time = time.time() - self.generation_start_time
            self.is_generating = False
            
            minutes = int(elapsed_time // 60)
            seconds = int(elapsed_time % 60)
            
            if minutes > 0:
                time_text = f"{minutes}m {seconds}s"
            else:
                time_text = f"{seconds}s"
                
            # Update UI in main thread
            self.root.after(0, lambda t=time_text: self._on_generation_success(t))
            
        except Exception as e:
            self.is_generating = False
            # Update UI in main thread
            self.root.after(0, lambda err=str(e): self._on_generation_error(err))
    
    def _on_generation_success(self, time_text):
        """Handle successful generation in main thread"""
        self.generate_button.config(state="normal")
        # Reset window title
        self.root.title("Purchase Order Generator")
        
        # Show custom success dialog
        dialog = CustomSuccessDialog(self.root, time_text)
        self.root.wait_window(dialog.dialog)
        
        # Handle dialog result
        if dialog.result == "save_as":
            self.save_as_dialog()
        elif dialog.result == "ok":
            # Just close dialog, file remains in temp folder
            pass
        
        # Enable save button in main UI
        self.show_save_button()
    
    def _on_generation_error(self, error_message):
        """Handle generation error in main thread"""
        self.generate_button.config(state="normal")
        # Reset window title
        self.root.title("Purchase Order Generator")
        messagebox.showerror("Error", f"An error occurred while generating the PO:\n{error_message}")

    def generate_po(self):
        if not self.validate_inputs():
            return
            
        try:
            # Disable generate button to prevent multiple clicks
            self.generate_button.config(state="disabled")
            
            # Start timing and set generation state
            self.generation_start_time = time.time()
            self.is_generating = True
            
            # Hide buttons during generation
            self.hide_save_button()
            self.hide_open_file_button()
            self.root.update()
            
            # Start the elapsed time updater
            self.update_elapsed_time()
            
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
            
            # Start generation in a separate thread
            generation_thread = threading.Thread(target=self._generate_po_thread, args=(gui_data,))
            generation_thread.daemon = True
            generation_thread.start()

        except Exception as e:
            self.is_generating = False
            self.generate_button.config(state="normal")
            # Reset window title
            self.root.title("Purchase Order Generator")
            messagebox.showerror("Error", f"An error occurred while generating the PO:\n{str(e)}")
    
    def update_api_key_display(self):
        """Update the API key display based on whether it's set"""
        if self.google_api_key.get():
            if hasattr(self, 'api_status_indicator'):
                self.api_status_indicator.config(text="●", foreground="green")
            if hasattr(self, 'api_status_text'):
                self.api_status_text.config(text="API Active", foreground="green")
            if hasattr(self, 'generate_button'):
                self.generate_button.config(state="normal")
        else:
            if hasattr(self, 'api_status_indicator'):
                self.api_status_indicator.config(text="●", foreground="red")
            if hasattr(self, 'api_status_text'):
                self.api_status_text.config(text="API Key Missing", foreground="red")
            if hasattr(self, 'generate_button'):
                self.generate_button.config(state="disabled")
    
    def open_api_key_dialog(self):
        """Open API key dialog for setting the key"""
        dialog = APIKeyDialog(self.root)
        self.root.wait_window(dialog.dialog)
        
        if dialog.result:
            self.google_api_key.set(dialog.result)
            self.update_api_key_display()
            self.save_current_settings(silent=True)
            self.pdf_processor._api_key = dialog.result
            self.pdf_processor._client = None  # Reset client to force re-initialization
            messagebox.showinfo("Success", "API key set successfully")


class APIKeyDialog:
    def __init__(self, parent):
        self.result = None
        self.test_result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Set Google API Key")
        self.dialog.geometry("480x240")
        self.dialog.resizable(False, False)
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (480 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (240 // 2)
        self.dialog.geometry(f"480x240+{x}+{y}")
        
        self.create_widgets()
        
    def create_widgets(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(main_frame, text="Enter your Google API Key:", font=("Arial", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))
        
        self.api_entry = ttk.Entry(main_frame, width=50, show="*")
        self.api_entry.pack(fill=tk.X, pady=(0, 10))
        self.api_entry.focus()
        
        # Link to Google AI Studio
        link_frame = ttk.Frame(main_frame)
        link_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(link_frame, text="Need an API key? ", font=("Arial", 9)).pack(side=tk.LEFT)
        
        # Create clickable link
        link_label = ttk.Label(link_frame, text="Get one here", font=("Arial", 9, "underline"), foreground="blue", cursor="hand2")
        link_label.pack(side=tk.LEFT)
        link_label.bind("<Button-1>", self.open_api_key_website)
        
        # Test result label
        self.test_result_label = ttk.Label(main_frame, text="", font=("Arial", 9))
        self.test_result_label.pack(fill=tk.X, pady=(5, 10))
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Test Connection button
        self.test_button = ttk.Button(button_frame, text="Test Connection", command=self.test_connection)
        self.test_button.pack(side=tk.LEFT)
        
        # OK and Cancel buttons
        ttk.Button(button_frame, text="OK", command=self.on_ok).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self.on_cancel).pack(side=tk.RIGHT)
        
        # Bind Enter key to OK
        self.dialog.bind('<Return>', lambda e: self.on_ok())
        self.dialog.bind('<Escape>', lambda e: self.on_cancel())
        
    def test_connection(self):
        """Test the API key connection"""
        api_key = self.api_entry.get().strip()
        if not api_key:
            self.test_result_label.config(text="Please enter an API key first", foreground="red")
            return
            
        try:
            # Show testing status
            self.test_result_label.config(text="Testing connection...", foreground="blue")
            self.dialog.update()
            
            # Test the API key by initializing a client
            from google import genai
            client = genai.Client(api_key=api_key)
            
            # Simple test - list models (this will fail if key is invalid)
            models = client.models.list()
            
            self.test_result = True
            self.test_result_label.config(text="✓ Connection successful!", foreground="green")
            
        except Exception as e:
            self.test_result = False
            error_msg = str(e)
            if "API_KEY_INVALID" in error_msg or "permission" in error_msg.lower():
                self.test_result_label.config(text="✗ Invalid API key", foreground="red")
            elif "network" in error_msg.lower() or "connection" in error_msg.lower():
                self.test_result_label.config(text="✗ Network error - check internet connection", foreground="orange")
            else:
                self.test_result_label.config(text=f"✗ Error: {error_msg[:50]}...", foreground="red")
        
    def open_api_key_website(self, event):
        """Open Google AI Studio API keys website"""
        import webbrowser
        webbrowser.open("https://aistudio.google.com/api-keys")
        
    def on_ok(self):
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showerror("Error", "API key cannot be empty")
            return
        
        # If we haven't tested or test failed, ask user to confirm
        if self.test_result is None:
            if not messagebox.askyesno("Confirm", "You haven't tested the API key. Continue anyway?"):
                return
        elif self.test_result is False:
            if not messagebox.askyesno("Confirm", "The API key test failed. Are you sure you want to continue?"):
                return
        
        self.result = api_key
        self.dialog.destroy()
        
    def on_cancel(self):
        self.dialog.destroy()