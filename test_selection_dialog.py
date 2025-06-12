"""
test_selection_dialog.py
Developed by Charlie Becquet.
Dialog for selecting tests to include in a new file.
"""

import tkinter as tk
from tkinter import ttk
from utils import APP_BACKGROUND_COLOR, FONT

class TestSelectionDialog:
    def __init__(self, parent, available_tests):
        """
        Initialize the test selection dialog.
        
        Args:
            parent (tk.Tk): The parent window.
            available_tests (list): List of available test names.
        """
        self.parent = parent
        self.available_tests = available_tests
        self.selected_tests = []
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select Tests")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Set fixed size to ensure everything is visible
        self.dialog.geometry("500x400")
        # Prevent resizing smaller than needed to show buttons
        self.dialog.minsize(500, 400)
        
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
        # Use grid layout for better control
        self.dialog.grid_columnconfigure(0, weight=1)  # Tests column expands
        self.dialog.grid_columnconfigure(1, weight=0)  # Buttons column fixed width
        self.dialog.grid_rowconfigure(0, weight=0)  # Header row fixed height
        self.dialog.grid_rowconfigure(1, weight=1)  # Content row expands
        
        # Header label with white text
        header_frame = ttk.Frame(self.dialog)
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
        
        header_label = ttk.Label(
            header_frame, 
            text="Select tests to include in your new file:",
            font=FONT,
            foreground="white"
        )
        header_label.pack(anchor="w")
        
        # Left side: scrollable tests frame
        test_frame = ttk.Frame(self.dialog)
        test_frame.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=5)
        test_frame.grid_columnconfigure(0, weight=1)
        test_frame.grid_rowconfigure(0, weight=1)
        
        # Scrollable frame for checkboxes
        self.canvas = tk.Canvas(test_frame)
        scrollbar = ttk.Scrollbar(test_frame, orient="vertical", command=self.canvas.yview)
        
        self.checkbox_frame = ttk.Frame(self.canvas)
        self.checkbox_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Create window inside canvas
        self.canvas.create_window((0, 0), window=self.checkbox_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        # Create a variable for each test - no extra padding on right
        self.test_vars = {}
        for i, test in enumerate(sorted(self.available_tests)):
            var = tk.BooleanVar(value=True)  # All tests selected by default
            self.test_vars[test] = var
            cb = ttk.Checkbutton(self.checkbox_frame, text=test, variable=var)
            cb.grid(row=i, column=0, sticky="w", pady=2)
        
        # Right side: buttons frame - fixed width
        button_frame = ttk.Frame(self.dialog, width=120)
        button_frame.grid(row=1, column=1, sticky="ns", padx=(5, 10), pady=5)
        button_frame.grid_propagate(False)  # Keep fixed width
        
        # Stack buttons vertically with fixed width
        select_all_btn = ttk.Button(button_frame, text="Select All", command=self.select_all)
        select_all_btn.pack(side="top", fill="x", pady=5)
        
        deselect_all_btn = ttk.Button(button_frame, text="Deselect All", command=self.deselect_all)
        deselect_all_btn.pack(side="top", fill="x", pady=5)
        
        # Add a separator
        ttk.Separator(button_frame, orient="horizontal").pack(fill="x", pady=10)
        
        ok_btn = ttk.Button(button_frame, text="OK", command=self.on_ok)
        ok_btn.pack(side="top", fill="x", pady=5)
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.on_cancel)
        cancel_btn.pack(side="top", fill="x", pady=5)
        
        # Bind mouse wheel events for scrolling
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
    
    def _on_mousewheel(self, event):
        """Handle mouse wheel scrolling."""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def select_all(self):
        """Select all tests."""
        for var in self.test_vars.values():
            var.set(True)
    
    def deselect_all(self):
        """Deselect all tests."""
        for var in self.test_vars.values():
            var.set(False)
    
    def on_ok(self):
        """Handle OK button click."""
        self.selected_tests = [test for test, var in self.test_vars.items() if var.get()]
        self.result = True
        self.dialog.destroy()
    
    def on_cancel(self):
        """Handle Cancel button click."""
        self.result = False
        self.dialog.destroy()
    
    def show(self):
        """
        Show the dialog and wait for user input.
        
        Returns:
            tuple: (result, selected_tests) where result is True if OK was clicked,
                   and selected_tests is the list of selected test names.
        """
        # Unbind the mousewheel event when dialog closes
        self.dialog.protocol("WM_DELETE_WINDOW", 
                            lambda: (self.canvas.unbind_all("<MouseWheel>"), 
                                     self.on_cancel()))
        
        self.dialog.wait_window()
        return self.result, self.selected_tests