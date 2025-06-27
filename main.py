"""
Main module for the DataViewer Application. Developed by Charlie Becquet.
Provides a Tkinter-based interface for interacting with Excel data, generating reports,
and plotting graphs.
"""

import logging
import os
import sys
import tkinter as tk
from tkinter import messagebox
from utils import debug_print
# Add current directory to path for development mode
if __name__ == "__main__" and not hasattr(sys, '_MEIPASS'):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)

from utils import get_resource_path  # Utility function
from main_gui import TestingGUI
from image_loader import ImageLoader
from file_manager import FileManager
from plot_manager import PlotManager
from report_generator import ReportGenerator
from trend_analysis_gui import TrendAnalysisGUI
from progress_dialog import ProgressDialog

def main():
    """Main entry point for the application."""

    def show_error_message(e):
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

    # Set up logging - use user directory for installed package
    try:
        # Try to use application data directory
        if hasattr(sys, '_MEIPASS'):
            # PyInstaller
            log_dir = os.path.dirname(__file__)
        else:
            # Installed package or development
            log_dir = os.path.expanduser("~/.standardized-testing-gui")
            os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, 'app.log')
    except Exception:
        # Fallback to current directory
        log_file = os.path.join(os.path.dirname(__file__), 'app.log')
    
    logging.basicConfig(
        filename=log_file, 
        level=logging.DEBUG, 
        format='%(asctime)s:%(levelname)s:%(message)s'
    )
    
    logging.info('Application started')
    debug_print(f"DEBUG: Log file location: {log_file}")

    try:
        # Initialize Tkinter root window
        root = tk.Tk()

        # Initialize and launch the GUI application
        app = TestingGUI(root)

        # Start the Tkinter main loop (keeps the application running)
        root.mainloop()

    except Exception as e:
        logging.error("An error occurred", exc_info=True)
        show_error_message(e)

if __name__ == "__main__":
    main()