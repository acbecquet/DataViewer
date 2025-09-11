# utils/constants.py
"""
utils/constants.py
Application constants and configuration values.
Consolidated from various modules in refactor_files.
"""

# UI Constants
FONT = ('Arial', 10)
APP_BACKGROUND_COLOR = '#D3D3D3'
BUTTON_COLOR = '#4169E1'
PLOT_CHECKBOX_TITLE = "Click Checkbox to \nAdd/Remove Item \nFrom Plot"

# Window and Layout Constants
DEFAULT_WINDOW_SIZE = (1500, 800)
DEFAULT_WINDOW_WIDTH_RATIO = 0.8
DEFAULT_WINDOW_HEIGHT_RATIO = 0.7

# File Type Constants
SUPPORTED_FILE_TYPES = [
    ("Excel files", "*.xlsx *.xls"),
    ("All files", "*.*")
]

EXCEL_EXTENSIONS = ['.xlsx', '.xls']
IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff']

# Database Constants
DEFAULT_DATABASE_TIMEOUT = 30
CACHE_DURATION_HOURS = 24

# Processing Constants
DEFAULT_PLOT_TYPES = [
    "TPM (mg/puff)",
    "Normalized TPM (mg/s)" 
    "Draw Pressure (kPa)",
    "Resistance (Ohm)",
    "Power Efficiency (%)",
    "Usage Efficiency (%)"
]

# Validation Constants
MIN_SAMPLE_ROWS = 1
MIN_SAMPLE_COLUMNS = 2
MAX_COLUMN_WIDTH = 50

# Debug Constants
DEBUG_LOG_FORMAT = "DEBUG: {message}"
ERROR_LOG_FORMAT = "ERROR: {message}"
SUCCESS_LOG_FORMAT = "SUCCESS: {message}"

# Report Constants
DEFAULT_REPORT_FORMAT = "xlsx"
TEMP_FILE_PREFIX = "dataviewer_temp_"

print("DEBUG: constants.py - All constants loaded successfully")