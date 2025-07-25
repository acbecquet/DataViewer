"""
test_start_menu.py
Developed by Charlie Becquet.
Menu for starting a test or viewing the raw file.
"""

import tkinter as tk
from tkinter import ttk
from utils import APP_BACKGROUND_COLOR, FONT, BUTTON_COLOR

class TestStartMenu:
    def __init__(self, parent, file_path):
        """
        Initialize the test start menu.
        
        Args:
            parent (tk.Tk): The parent window.
            file_path (str): Path to the created file.
        """
        self.parent = parent
        self.file_path = file_path
        self.result = None
        self.selected_test = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Test Menu")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=APP_BACKGROUND_COLOR)
        
        self.create_widgets()
        self.center_window()
    
    def center_window(self):
        """Center the dialog on the screen."""
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    def create_widgets(self):
        """Create the dialog widgets."""
        frame = ttk.Frame(self.dialog, padding=20)
        frame.pack(fill="both", expand=True)
        
        header_label = ttk.Label(
            frame, 
            text="What would you like to do?", 
            font=("Arial", 14),
            foreground="white",
            background=APP_BACKGROUND_COLOR
        )
        header_label.pack(pady=(0, 20))
        
        # Get tests from the file to populate the dropdown
        import openpyxl
        wb = openpyxl.load_workbook(self.file_path)
        self.available_tests = wb.sheetnames
        
        # Test selection dropdown (only shown when starting a test)
        self.test_selection_frame = ttk.Frame(frame)
        self.test_selection_frame.pack(fill="x", pady=(0, 20))
        
        ttk.Label(
            self.test_selection_frame, 
            text="Select test to conduct:",
            foreground="white",
            background=APP_BACKGROUND_COLOR
        ).pack(side="left", padx=(0, 10))
        
        self.test_var = tk.StringVar()
        if self.available_tests:
            self.test_var.set(self.available_tests[0])
        
        self.test_dropdown = ttk.Combobox(
            self.test_selection_frame, 
            textvariable=self.test_var,
            values=self.available_tests,
            state="readonly",
            width=30
        )
        self.test_dropdown.pack(side="left", fill="x", expand=True)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Use fixed width buttons for better layout
        button_width = 15
        
        start_test_btn = ttk.Button(
            button_frame, 
            text="Start Test", 
            command=self.on_start_test,
            width=button_width
        )
        start_test_btn.pack(side="left", padx=(0, 10))
        
        view_file_btn = ttk.Button(
            button_frame, 
            text="View Raw File", 
            command=self.on_view_raw_file,
            width=button_width
        )
        view_file_btn.pack(side="right")
    
    def on_start_test(self):
        """Handle Start Test button click."""
        self.selected_test = self.test_var.get()
        self.result = "start_test"
        self.dialog.destroy()
    
    def on_view_raw_file(self):
        """Handle View Raw File button click."""
        self.result = "view_raw_file"
        self.dialog.destroy()
    
    def show(self):
        """
        Show the dialog and wait for user input.
        
        Returns:
            tuple: (result, selected_test) where result is "start_test" or "view_raw_file"
                  and selected_test is the name of the selected test (if applicable).
        """
        self.dialog.wait_window()
        return self.result, self.selected_test