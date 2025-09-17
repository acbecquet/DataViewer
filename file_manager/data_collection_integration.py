"""
Data Collection Integration Module for DataViewer Application

This module handles integration with the data collection window,
including processing results and coordinating window transitions.
"""

# Local imports
from utils import debug_print


class DataCollectionIntegration:
    """Handles data collection window integration and result processing."""
    
    def __init__(self, file_manager):
        """Initialize with reference to parent FileManager."""
        self.file_manager = file_manager
        self.gui = file_manager.gui
        self.root = file_manager.root
        
    def handle_data_collection_close(self, data_collection_window, result):
        """Handle closing of data collection window with sample image support."""
        try:
            if result == "load_file":
                # Check for sample images before loading
                if hasattr(data_collection_window, 'sample_images') and data_collection_window.sample_images:
                    debug_print("DEBUG: Data collection closed with sample images")

                    # Store sample images in main GUI for saving
                    self.gui.pending_sample_images = data_collection_window.sample_images.copy()
                    self.gui.pending_sample_image_crop_states = data_collection_window.sample_image_crop_states.copy()
                    self.gui.pending_sample_header_data = data_collection_window.header_data.copy()

                    # Transfer formatted images
                    data_collection_window.transfer_images_to_main_gui()

                    # Save with sample images
                    if self.gui.file_path and self.gui.file_path.endswith('.vap3'):
                        self.gui.save_with_sample_images()

                # Load the file for viewing
                self.file_manager.load_file(data_collection_window.file_path)

            # Restore main window visibility
            if hasattr(data_collection_window, 'main_window_was_visible') and data_collection_window.main_window_was_visible:
                self.gui.root.deiconify()

        except Exception as e:
            debug_print(f"ERROR: Failed to handle data collection close: {e}")
            import traceback
            traceback.print_exc()