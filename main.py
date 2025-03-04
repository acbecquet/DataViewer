"""
Main module for the DataViewer Application. Developed by Charlie Becquet.
Provides a Tkinter-based interface for interacting with Excel data, generating reports,
and plotting graphs.
"""

import logging
import os
import tkinter as tk
from tkinter import messagebox
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

    # Set up logging to a filelog_file = os.path.join(os.path.expanduser("~"), 'app.log')
    log_file = os.path.join(os.path.dirname(__file__), 'app.log')
    logging.basicConfig(filename=log_file, level=logging.DEBUG, 
                        format='%(asctime)s:%(levelname)s:%(message)s')

    logging.info('Application started')

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

