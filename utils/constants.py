# utils/constants.py
"""
utils/constants.py
Application constants and configuration values.
These will replace constants from the existing utils.py.
"""

import tkinter as tk
from pathlib import Path

# === APPLICATION INFO ===
APP_NAME = "DataViewer"
APP_VERSION = "3.0.0"
APP_AUTHOR = "Charlie Becquet"
APP_DESCRIPTION = "Standardized Testing Data Analysis Tool"

# === UI CONSTANTS ===
APP_BACKGROUND_COLOR = '#EFEFEF'
BUTTON_COLOR = '#E0E0E0'
FONT = ('Arial', 10)
TITLE_FONT = ('Arial', 12, 'bold')
LARGE_FONT = ('Arial', 14, 'bold')

# === WINDOW SETTINGS ===
DEFAULT_WINDOW_SIZE = (1200, 800)
MIN_WINDOW_SIZE = (1000, 600)
DEFAULT_WINDOW_TITLE = f"{APP_NAME} v{APP_VERSION}"

# === PLOT SETTINGS ===
PLOT_CHECKBOX_TITLE = "Plot Lines"
DEFAULT_PLOT_TYPE = "TPM"
PLOT_FIGURE_SIZE = (10, 6)
PLOT_DPI = 100
PLOT_LINE_WIDTH = 2.0
PLOT_MARKER_SIZE = 6.0

# === FILE HANDLING ===
SUPPORTED_FILE_TYPES = [
    ("Excel files", "*.xlsx *.xls"),
    ("VAP3 files", "*.vap3"),
    ("CSV files", "*.csv"),
    ("All files", "*.*")
]

EXCEL_EXTENSIONS = ['.xlsx', '.xls']
VAP3_EXTENSIONS = ['.vap3']
CSV_EXTENSIONS = ['.csv']

MAX_FILE_SIZE_MB = 100
MAX_CACHE_SIZE_MB = 500

# === DATABASE SETTINGS ===
DEFAULT_DATABASE_NAME = "dataviewer.db"
DATABASE_TIMEOUT = 30
MAX_DATABASE_CONNECTIONS = 10

# === PROCESSING SETTINGS ===
DEFAULT_PROCESSING_TIMEOUT = 60
MAX_PROCESSING_THREADS = 4
CHUNK_SIZE_ROWS = 1000

# === PLOT OPTIONS ===
PLOT_OPTIONS = [
    "TPM",
    "Draw Pressure", 
    "Resistance",
    "Power Efficiency",
    "TPM (Bar)"
]

# === REPORT SETTINGS ===
REPORT_FORMATS = ["excel", "powerpoint", "pdf"]
DEFAULT_REPORT_FORMAT = "excel"
REPORT_DPI = 300
REPORT_OUTPUT_DIR = "reports"

# === IMAGE SETTINGS ===
SUPPORTED_IMAGE_FORMATS = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']
THUMBNAIL_SIZE = (150, 150)
MAX_IMAGE_SIZE_MB = 50

# === DIRECTORIES ===
CONFIG_DIR = Path("config")
DATA_DIR = Path("data")
CACHE_DIR = Path("cache")
LOGS_DIR = Path("logs")
TEMP_DIR = Path("temp")
BACKUP_DIR = Path("backups")

# === VALIDATION CONSTANTS ===
MIN_DATA_ROWS = 2
MIN_DATA_COLUMNS = 2
MAX_COLUMN_NAME_LENGTH = 100
MAX_SHEET_NAME_LENGTH = 31  # Excel limit

# === NETWORK SETTINGS ===
REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
UPDATE_CHECK_INTERVAL_HOURS = 24

# === ERROR MESSAGES ===
ERROR_MESSAGES = {
    'file_not_found': "File not found: {path}",
    'file_access_denied': "Access denied: {path}",
    'invalid_file_format': "Invalid file format: {format}",
    'processing_failed': "Data processing failed: {error}",
    'database_error': "Database error: {error}",
    'network_error': "Network error: {error}",
    'insufficient_data': "Insufficient data for operation",
    'invalid_configuration': "Invalid configuration: {details}"
}

# === SUCCESS MESSAGES ===
SUCCESS_MESSAGES = {
    'file_loaded': "File loaded successfully: {filename}",
    'file_saved': "File saved successfully: {filename}",
    'report_generated': "Report generated: {filename}",
    'data_processed': "Data processed successfully",
    'model_trained': "Model trained successfully",
    'database_updated': "Database updated successfully"
}

# === PROGRESS MESSAGES ===
PROGRESS_MESSAGES = {
    'loading_file': "Loading file...",
    'processing_data': "Processing data...",
    'generating_plot': "Generating plot...",
    'creating_report': "Creating report...",
    'saving_file': "Saving file...",
    'updating_database': "Updating database..."
}

# === SHEET TYPE KEYWORDS ===
PLOTTING_SHEET_KEYWORDS = [
    'test', 'data', 'measurement', 'results', 'analysis',
    'tpm', 'resistance', 'pressure', 'efficiency', 'performance'
]

SUMMARY_SHEET_KEYWORDS = [
    'summary', 'overview', 'report', 'conclusion', 'abstract'
]

# === VISCOSITY CALCULATION CONSTANTS ===
VISCOSITY_MODEL_TYPES = ['polynomial', 'arrhenius', 'linear', 'exponential']
DEFAULT_POLYNOMIAL_DEGREE = 2
GAS_CONSTANT = 8.314  # J/(mol·K)
ABSOLUTE_ZERO = 273.15  # Kelvin

# === COLOR SCHEMES ===
PLOT_COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'
]

STATUS_COLORS = {
    'success': '#28a745',
    'warning': '#ffc107', 
    'error': '#dc3545',
    'info': '#17a2b8'
}

# === KEYBOARD SHORTCUTS ===
KEYBOARD_SHORTCUTS = {
    'new_file': 'Ctrl+N',
    'open_file': 'Ctrl+O',
    'save_file': 'Ctrl+S',
    'exit_app': 'Ctrl+Q',
    'help': 'F1',
    'refresh': 'F5'
}

# === LOGGING SETTINGS ===
LOG_LEVELS = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
DEFAULT_LOG_LEVEL = 'INFO'
MAX_LOG_FILE_SIZE_MB = 10
LOG_BACKUP_COUNT = 5

print(f"DEBUG: Constants loaded - {APP_NAME} v{APP_VERSION}")
print(f"DEBUG: Default plot options: {', '.join(PLOT_OPTIONS)}")
print(f"DEBUG: Supported file types: {len(SUPPORTED_FILE_TYPES)} types")