import os
import json
from pathlib import Path

# Base paths
BASE_DIR = Path(__file__).parent.parent
TEMP_DIR = BASE_DIR / "temp"
JSONS_DIR = BASE_DIR / "jsons"
TEMPLATE_DIR = BASE_DIR / "data" / "templates"
CONFIG_DIR = BASE_DIR / "config"

# File paths
TEMPLATE_PATH = TEMPLATE_DIR / "po_template.xlsx"
USER_SETTINGS_PATH = CONFIG_DIR / "user_settings.json"

# Gemini configuration
GEMINI_MODEL = "gemini-2.5-flash"
GEMINI_TEMPERATURE = 0.2
GEMINI_TOP_P = 0.9

# Excel configuration
EXCEL_START_ROW = 31
EXCEL_TABLE_END_ROW = 50
ROWS_PER_ITEM = 2
ADDRESS_MAX_LENGTH = 45

# Default user settings
DEFAULT_USER_SETTINGS = {
    "po_number": "",
    "project_name": "",
    "purchaser_name": "",
    "phone_code": "+60",
    "phone_number_only": "",
    "director_manager": "",
    "remember_details": False
}

# Create necessary directories
for directory in [TEMP_DIR, JSONS_DIR, TEMPLATE_DIR, CONFIG_DIR]:
    directory.mkdir(parents=True, exist_ok=True)

def load_user_settings():
    """Load user settings from JSON file"""
    try:
        if USER_SETTINGS_PATH.exists():
            with open(USER_SETTINGS_PATH, 'r') as f:
                return json.load(f)
    except Exception as e:
        print(f"Error loading user settings: {e}")
    
    return DEFAULT_USER_SETTINGS.copy()

def save_user_settings(settings):
    """Save user settings to JSON file"""
    try:
        with open(USER_SETTINGS_PATH, 'w') as f:
            json.dump(settings, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving user settings: {e}")
        return False