"""
Runtime hook for PyInstaller to optimize startup performance.
Pre-imports common modules and sets environment variables.
"""

import os
import sys

# Optimize Python startup
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
os.environ['PYTHONUNBUFFERED'] = '1'

# Pre-import commonly used modules to avoid import delays
try:
    import tkinter
    import tkinter.ttk
    import tkinter.messagebox
    import tkinter.filedialog
    import json
    import pathlib
    import openpyxl
    import num2words
except ImportError:
    # If any module fails to import, continue silently
    # The actual import will be handled by the application
    pass

# Set up proper paths for packaged application
if getattr(sys, 'frozen', False):
    # Running as PyInstaller bundle
    application_path = sys._MEIPASS
else:
    # Running as script
    application_path = os.path.dirname(os.path.abspath(__file__))

# Add application path to sys.path if not already there
if application_path not in sys.path:
    sys.path.insert(0, application_path)

print("Runtime hook loaded - optimizing startup...")
