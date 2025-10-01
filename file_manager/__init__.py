"""
File Management Package for DataViewer Application

This package handles all file-related operations including:
- Loading Excel files (standard and legacy formats)
- Saving and loading .vap3 files (custom format)
- Opening Excel files for editing and monitoring changes
- Managing file selections and UI updates related to files
- Database operations and batch processing
"""

# Import all module classes
from .core_file_operations import CoreFileOperations
from .database_operations import DatabaseOperations
from .ui_management import UIManager
from .header_data_processor import HeaderDataProcessor
from .vap3_file_handler import Vap3FileHandler
from .batch_operations import BatchOperations
from .excel_integration import ExcelIntegration
from .test_workflow import TestWorkflow
from .data_collection_integration import DataCollectionIntegration

# Local imports
from database_manager import DatabaseManager


class FileManager:
    """File Management Module for DataViewer.

    This class coordinates all file-related operations by delegating to
    specialized modules while maintaining shared state and caches.
    """
    
    def __init__(self, gui):
        """Initialize FileManager with GUI reference and all sub-modules."""
        self.gui = gui
        self.root = gui.root
        self.db_manager = DatabaseManager()

        # Add cache to prevent redundant operations
        self.loaded_files_cache = {}  # Cache for loaded file data
        self.stored_files_cache = set()  # Track files already stored in database

        # Initialize all operational modules
        self.core_ops = CoreFileOperations(self)
        self.db_ops = DatabaseOperations(self)
        self.ui_manager = UIManager(self)
        self.header_processor = HeaderDataProcessor(self)
        self.vap3_handler = Vap3FileHandler(self)
        self.batch_ops = BatchOperations(self)
        self.excel_integration = ExcelIntegration(self)
        self.test_workflow = TestWorkflow(self)
        self.data_collection_integration = DataCollectionIntegration(self)

    # ==================== CORE FILE OPERATIONS ====================
    
    def load_excel_file(self, *args, **kwargs):
        """Delegate to core operations."""
        return self.core_ops.load_excel_file(*args, **kwargs)
    
    def load_initial_file(self):
        """Delegate to core operations."""
        return self.core_ops.load_initial_file()
    
    def reload_excel_file(self):
        """Delegate to core operations."""
        return self.core_ops.reload_excel_file()
    
    def load_file(self, file_path):
        """Delegate to core operations."""
        return self.core_ops.load_file(file_path)
    
    def ask_open_file(self):
        """Delegate to core operations."""
        return self.core_ops.ask_open_file()
    
    def add_data(self):
        """Delegate to core operations."""
        return self.core_ops.add_data()
    
    def set_active_file(self, file_name):
        """Delegate to core operations."""
        return self.core_ops.set_active_file(file_name)
    
    def ensure_file_is_loaded_in_ui(self, file_path):
        """Delegate to core operations."""
        return self.core_ops.ensure_file_is_loaded_in_ui(file_path)

    def close_current_file(self):
        """Delegate to core operations."""
        return self.core_ops.close_current_file()

    # ==================== DATABASE OPERATIONS ====================
    
    def _store_file_in_database(self, *args, **kwargs):
        """Delegate to database operations."""
        return self.db_ops._store_file_in_database(*args, **kwargs)
    
    def load_from_database(self, *args, **kwargs):
        """Delegate to database operations."""
        return self.db_ops.load_from_database(*args, **kwargs)
    
    def load_multiple_from_database(self, file_ids):
        """Delegate to database operations."""
        return self.db_ops.load_multiple_from_database(file_ids)
    
    def load_from_database_by_id(self, file_id):
        """Delegate to database operations."""
        return self.db_ops.load_from_database_by_id(file_id)
    
    def show_database_browser(self, comparison_mode=False):
        """Delegate to database operations."""
        return self.db_ops.show_database_browser(comparison_mode)

    # ==================== UI MANAGEMENT ====================
    
    def update_ui_for_current_file(self):
        """Delegate to UI management."""
        return self.ui_manager.update_ui_for_current_file()
    
    def ensure_main_gui_initialized(self):
        """Delegate to UI management."""
        return self.ui_manager.ensure_main_gui_initialized()
    
    def update_file_dropdown(self):
        """Delegate to UI management."""
        return self.ui_manager.update_file_dropdown()
    
    def add_or_update_file_dropdown(self):
        """Delegate to UI management."""
        return self.ui_manager.add_or_update_file_dropdown()
    
    def center_window(self, window, width, height):
        """Delegate to UI management."""
        return self.ui_manager.center_window(window, width, height)
    
    def start_file_loading_wrapper(self, startup_menu):
        """Delegate to UI management."""
        return self.ui_manager.start_file_loading_wrapper(startup_menu)
    
    def start_file_loading_database_wrapper(self, startup_menu):
        """Delegate to UI management."""
        return self.ui_manager.start_file_loading_database_wrapper(startup_menu)

    # ==================== HEADER DATA PROCESSING ====================
    
    def extract_existing_header_data(self, file_path, selected_test):
        """Delegate to header processor."""
        return self.header_processor.extract_existing_header_data(file_path, selected_test)
    
    def extract_header_data_from_loaded_sheets(self, selected_test):
        """Delegate to header processor."""
        return self.header_processor.extract_header_data_from_loaded_sheets(selected_test)
    
    def extract_header_data_from_excel_file(self, file_path, selected_test):
        """Delegate to header processor."""
        return self.header_processor.extract_header_data_from_excel_file(file_path, selected_test)
    
    def detect_sheet_format(self, ws):
        """Delegate to header processor."""
        return self.header_processor.detect_sheet_format(ws)
    
    def extract_old_format_header_data(self, ws, selected_test):
        """Delegate to header processor."""
        return self.header_processor.extract_old_format_header_data(ws, selected_test)
    
    def extract_new_format_header_data(self, ws, selected_test):
        """Delegate to header processor."""
        return self.header_processor.extract_new_format_header_data(ws, selected_test)
    
    def extract_header_data_from_excel_file_old(self, file_path, selected_test):
        """Delegate to header processor."""
        return self.header_processor.extract_header_data_from_excel_file_old(file_path, selected_test)
    
    def migrate_header_data_for_backwards_compatibility(self, header_data):
        """Delegate to header processor."""
        return self.header_processor.migrate_header_data_for_backwards_compatibility(header_data)
    
    def validate_header_data(self, header_data):
        """Delegate to header processor."""
        return self.header_processor.validate_header_data(header_data)
    
    def apply_header_data_to_file(self, file_path, header_data):
        """Delegate to header processor."""
        return self.header_processor.apply_header_data_to_file(file_path, header_data)
    
    def determine_sample_count_from_data(self, sheet_data, test_name):
        """Delegate to header processor."""
        return self.header_processor.determine_sample_count_from_data(sheet_data, test_name)

    # ==================== VAP3 FILE OPERATIONS ====================
    
    def save_as_vap3(self, filepath=None):
        """Delegate to VAP3 handler."""
        return self.vap3_handler.save_as_vap3(filepath)
    
    def load_vap3_file(self, *args, **kwargs):
        """Delegate to VAP3 handler."""
        return self.vap3_handler.load_vap3_file(*args, **kwargs)

    # ==================== BATCH OPERATIONS ====================
    
    def batch_load_folder(self):
        """Delegate to batch operations."""
        return self.batch_ops.batch_load_folder()

    # ==================== EXCEL INTEGRATION ====================
    
    def open_raw_data_in_excel(self, sheet_name=None):
        """Delegate to Excel integration."""
        return self.excel_integration.open_raw_data_in_excel(sheet_name)

    # ==================== TEST WORKFLOW ====================
    
    def create_new_template(self, parent_window=None):
        """Delegate to test workflow."""
        return self.test_workflow.create_new_template(parent_window)
    
    def create_new_file_with_tests(self, parent_window=None):
        """Delegate to test workflow."""
        return self.test_workflow.create_new_file_with_tests(parent_window)
    
    def show_test_start_menu(self, file_path, original_filename=None):
        """Delegate to test workflow."""
        return self.test_workflow.show_test_start_menu(file_path, original_filename)
    
    def show_header_data_dialog(self, file_path, selected_test):
        """Delegate to test workflow."""
        return self.test_workflow.show_header_data_dialog(file_path, selected_test)
    
    def start_data_collection_with_header_data(self, file_path, selected_test, header_data):
        """Delegate to test workflow."""
        return self.test_workflow.start_data_collection_with_header_data(file_path, selected_test, header_data)

    # ==================== DATA COLLECTION INTEGRATION ====================
    
    def handle_data_collection_close(self, data_collection_window, result):
        """Delegate to data collection integration."""
        return self.data_collection_integration.handle_data_collection_close(data_collection_window, result)


# Export the main class
__all__ = ['FileManager']