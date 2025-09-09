"""
Main module for the DataViewer Application. Developed by Charlie Becquet.
Provides a Tkinter-based interface for interacting with Excel data, generating reports,
and plotting graphs.
"""

import time
startup_timer = {
    'start_time': time.time(),
    'checkpoints': []
}

def log_timing_checkpoint(checkpoint_name):
    """Log a timing checkpoint with elapsed time."""
    elapsed = time.time() - startup_timer['start_time']
    startup_timer['checkpoints'].append((checkpoint_name, elapsed))
    print(f"TIMING: {checkpoint_name}: {elapsed:.3f}s")

# Time each import in main.py individually
import logging
import os
import sys
log_timing_checkpoint("Basic imports complete")

import_start = time.time()
import tkinter as tk
from tkinter import messagebox
import_time = time.time() - import_start
print(f"TIMING: tkinter imports took: {import_time:.3f}s")

# Add current directory to path for development mode
if __name__ == "__main__" and not hasattr(sys, '_MEIPASS'):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    if current_dir not in sys.path:
        sys.path.insert(0, current_dir)

log_timing_checkpoint("Initial setup complete")

import_start = time.time()
from resource_utils import get_resource_path  # Utility function
import_time = time.time() - import_start
print(f"TIMING: from utils import get_resource_path took: {import_time:.3f}s")

import_start = time.time()
from main_gui import TestingGUI
import_time = time.time() - import_start
print(f"TIMING: from main_gui import TestingGUI took: {import_time:.3f}s")

def main():
    """Main entry point for the application."""

    def show_error_message(e):
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

    log_timing_checkpoint("Main function started")

    # Set up logging - always use user directory (writable location)
    try:
        log_dir = os.path.expanduser("~/.standardized-testing-gui")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, 'app.log')
    except Exception:
        # Fallback to current directory
        log_file = 'app.log'

    logging.basicConfig(
        filename=str(log_file), 
        level=logging.DEBUG, 
        format='%(asctime)s:%(levelname)s:%(message)s'
    )
    
    logging.info('Application started')
    print(f"DEBUG: Log file location: {log_file}")
    log_timing_checkpoint("Logging setup complete")

    try:
        # Initialize Tkinter root window
        tkinter_start = time.time()
        root = tk.Tk()
        tkinter_time = time.time() - tkinter_start
        print(f"TIMING: Tkinter root creation took: {tkinter_time:.3f}s")
        log_timing_checkpoint("Tkinter root created")

        # Initialize and launch the GUI application
        gui_start = time.time()
        app = TestingGUI(root)
        gui_time = time.time() - gui_start
        print(f"TIMING: TestingGUI initialization took: {gui_time:.3f}s")
        log_timing_checkpoint("TestingGUI initialized")

        # Calculate and display total startup time
        total_startup_time = time.time() - startup_timer['start_time']
        print(f"\nTIMING SUMMARY:")
        print(f"================")
        for checkpoint, elapsed in startup_timer['checkpoints']:
            print(f"{checkpoint}: {elapsed:.3f}s")
        print(f"TOTAL STARTUP TIME: {total_startup_time:.3f}s")
        print(f"================\n")

        # Log timing to file as well
        logging.info(f"Startup completed in {total_startup_time:.3f}s")
        for checkpoint, elapsed in startup_timer['checkpoints']:
            logging.info(f"Timing: {checkpoint}: {elapsed:.3f}s")

        # Start the Tkinter main loop (keeps the application running)
        log_timing_checkpoint("Starting main loop")
        root.mainloop()

    except Exception as e:
        logging.error("An error occurred", exc_info=True)
        show_error_message(e)

if __name__ == "__main__":
    main()