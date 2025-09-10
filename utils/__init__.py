# utils/__init__.py
"""
utils/__init__.py
Utilities package for the DataViewer application.
Contains helper functions, configuration, and constants.
Enhanced with comprehensive functionality from refactor_files.
"""

# Import utility modules
from .helpers import *
from .config import Config, get_config
from .constants import *
from .validators import *
from .formatters import *
from .debug import debug_print, error_print, success_print

__all__ = [
    # From helpers.py - Core Functions
    'get_resource_path',
    'get_resource_dir', 
    'resource_exists',
    'resource_path',
    'debug_print',
    'set_debug_mode',
    'is_debug_enabled',
    
    # From helpers.py - Data Processing
    'clean_columns', 
    'remove_empty_columns',
    'round_values',
    'safe_numeric_conversion',
    
    # From helpers.py - File Operations
    'get_save_path',
    'get_save_path_dialog',
    'is_standard_file',
    'is_valid_excel_file',
    'load_excel_file',
    'load_excel_file_with_formulas',
    'read_sheet_with_values_standards',
    'unmerge_all_cells',
    'autofit_columns_in_excel',
    
    # From helpers.py - UI Functions
    'plotting_sheet_test',
    'wrap_text',
    'center_window',
    'show_success_message',
    'generate_temp_image',
    
    # From helpers.py - Validation
    'validate_dataframe',
    'validate_sheet_data',
    'ensure_directory_exists',
    'get_file_extension',
    'is_file_accessible',
    
    # From helpers.py - Processing Helpers
    'extract_meta_data',
    'map_meta_data_to_template', 
    'header_matches',
    
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
    'EXCEL_EXTENSIONS',
    'IMAGE_EXTENSIONS',
    'DEFAULT_PLOT_TYPES',
    
    # From validators.py
    'validate_file_path',
    'validate_excel_file',
    'validate_data_frame',
    'validate_sheet_data',
    'validate_numeric_column',
    'validate_plotting_data',
    'validate_database_connection',
    'validate_image_file',
    
    # From formatters.py
    'format_number',
    'format_percentage',
    'format_scientific',
    'format_datetime',
    'format_currency',
    'format_file_size',
    'format_duration',
    'format_table_cell',
    'format_dataframe_for_display',
    'auto_format_value',
    
    # From debug.py
    'debug_print',
    'error_print', 
    'success_print',
    'log_function_entry',
    'log_function_exit',
    'log_timing_checkpoint',
    'print_dataframe_info',
    'print_exception_details'
]

# Debug output for architecture tracking
print("DEBUG: Utils package initialized successfully with enhanced functionality")
print(f"DEBUG: Available utilities: {len(__all__)} functions and classes")
print("DEBUG: Includes functions from utils.py, resource_utils.py, and enhanced modules")