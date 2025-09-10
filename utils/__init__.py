"""
utils/__init__.py
Utilities package for the DataViewer application.
Contains helper functions, configuration, and constants.
"""

# Import utility modules
from .helpers import *
from .config import Config, get_config
from .constants import *
from .validators import *
from .formatters import *
from .debug import debug_print, error_print, success_print

__all__ = [
    # From helpers.py
    'get_resource_path',
    'clean_columns', 
    'get_save_path',
    'is_standard_file',
    'plotting_sheet_test',
    'wrap_text',
    'center_window',
    'show_success_message',
    
    # From config.py
    'Config',
    'get_config',
    
    # From constants.py
    'APP_BACKGROUND_COLOR',
    'BUTTON_COLOR',
    'PLOT_CHECKBOX_TITLE',
    'FONT',
    'DEFAULT_WINDOW_SIZE',
    'SUPPORTED_FILE_TYPES',
    
    # From validators.py
    'validate_file_path',
    'validate_excel_file',
    'validate_data_frame',
    'validate_sheet_data',
    
    # From formatters.py
    'format_number',
    'format_percentage',
    'format_scientific',
    'format_datetime',
    
    # From debug.py
    'debug_print',
    'error_print', 
    'success_print'
]

# Debug output for architecture tracking
print("DEBUG: Utils package initialized successfully")
print(f"DEBUG: Available utilities: {len(__all__)} functions and classes")