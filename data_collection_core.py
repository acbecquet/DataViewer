"""
data_collection_window.py
Developed by Charlie Becquet.
Interface for rapid test data collection with enhanced saving and menu functionality.
"""

import tkinter as tk
import logging
import os
from utils import debug_print
from data_collection_ui import DataCollectionUI
from data_collection_handlers import DataCollectionHandlers
from data_collection_data import DataCollectionData
from data_collection_file_io import DataCollectionFileIO

# Call the setup function
setup_logging()

class DataCollectionWindow(DataCollectionUI, DataCollectionHandlers, DataCollectionData, DataCollectionFileIO):

"""Main Data Collection window - coordinates all functionality"""

    def __init__(self, parent, file_path, test_name, header_data, original_filename=None):
        """
        Initialize the data collection window.

        Args:
            parent (tk.Tk): The parent window.
            file_path (str): Path to the Excel file.
            test_name (str): Name of the test being conducted.
            header_data (dict): Dictionary containing header data.
            original_filename (str): original filename for .vap3 files
        """
        # Debug flag - set to False to disable all debug output
        self.DEBUG = False
        self.parent = parent
        self.root = parent.root
        self.file_path = file_path
        self.test_name = test_name
        self.header_data = header_data
        self.num_samples = header_data["num_samples"]
        self.result = None
        self.updating_notes = False

        if hasattr(parent, 'root'):
            self.main_window_was_visible = parent.root.winfo_viewable()
            parent.root.withdraw()  # Hide main window
            debug_print("DEBUG: Main GUI window hidden")

        # Store the original filename for saving purposes
        self.original_filename = original_filename
        if self.original_filename:
            debug_print(f"DEBUG: DataCollectionWindow initialized with original filename: {self.original_filename}")

        # Validate num_samples
        if self.num_samples <= 0:
            self.num_samples = len(header_data.get('samples', []))
            if self.num_samples <= 0:
                self.num_samples = 1

        # Auto-save settings
        self.auto_save_interval = 5 * 60 * 1000  # 5 minutes in milliseconds
        self.auto_save_timer = None
        self.has_unsaved_changes = False
        self.last_save_time = None

        # Create the window
        self.window = tk.Toplevel(self.root)
        self.window.title(f"Data Collection - {test_name}")
        self.window.state('zoomed')
        self.window.minsize(1250, 625)

        self.window.lift()  # Bring to top of window stack
        self.window.focus_force()  # Force focus to this window

        # Make it stay on top temporarily
        self.window.attributes('-topmost', True)
        self.window.after(100, lambda: self.window.attributes('-topmost', False))

        self.sample_images = {}
        self.sample_image_crop_states = {}

        # Default puff interval
        self.puff_interval = 10  # Default to 10

        # Initialize tab tracking for notes handling
        self.previous_tab_index = 0
        self.updating_notes = False
        debug_print("DEBUG: Initialized tab change tracking variables")

        # Set up keyboard shortcut flags
        self.hotkeys_enabled = True
        self.hotkey_bindings = {}

        # Create the style for ttk widgets
        self.style = ttk.Style()
        self.setup_styles()

        # Initialize data structures
        self.initialize_data()

        # Create the menu bar first
        self.create_menu_bar()

        # Create the UI
        self.create_widgets()

        self.load_existing_data_from_file()

        self.load_existing_sample_images_from_vap3()

        # Center the window
        self.center_window()

        # Set up event handlers
        self.setup_event_handlers()

        # Start auto-save timer
        self.start_auto_save_timer()

        self.ensure_initial_tpm_calculation()

        self.log(f"DataCollectionWindow initialized for {test_name} with {self.num_samples} samples", "debug")

    def log(self, message, level="info"):
        """Log a message with the specified level."""
        logger = logging.getLogger("DataCollectionWindow")

        if level.lower() == "debug":
            logger.debug(message)
        elif level.lower() == "info":
            logger.info(message)
        elif level.lower() == "warning":
            logger.warning(message)
        elif level.lower() == "error":
            logger.error(message)
        elif level.lower() == "critical":
            logger.critical(message)

    def show(self):
        """
        Show the window and wait for user input.

        Returns:
            str: "load_file" if data was saved and file should be loaded for viewing,
                 "cancel" if the user cancelled.
        """
        debug_print("DEBUG: Showing DataCollectionWindow")
        self.window.lift()
        self.window.focus_force()
        self.window.grab_set()  # Ensure it maintains focus

        self.window.wait_window()

        # Clean up auto-save timer
        if self.auto_save_timer:
            self.window.after_cancel(self.auto_save_timer)

        debug_print(f"DEBUG: DataCollectionWindow closed with result: {self.result}")
        return self.result

