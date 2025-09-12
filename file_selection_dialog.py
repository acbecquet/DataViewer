"""
file_selection_dialog.py
Developed by Charlie Becquet.
Dialog for selecting files to include in sample comparison.
"""

import tkinter as tk
from tkinter import ttk
from utils import APP_BACKGROUND_COLOR, FONT, debug_print

class FileSelectionDialog:
    def __init__(self, parent, available_files):
        """
        Initialize the file selection dialog.

        Args:
            parent (tk.Tk): The parent window.
            available_files (list): List of file data dictionaries.
        """
        debug_print("DEBUG: Initializing FileSelectionDialog")
        self.parent = parent
        self.available_files = available_files
        self.selected_files = []
        self.result = None

        # Initialize references for cleanup
        self.canvas = None
        self.mousewheel_binding_id = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select Files for Comparison")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Set fixed size to ensure everything is visible
        self.dialog.geometry("600x500")
        # Prevent resizing smaller than needed to show buttons
        self.dialog.minsize(600, 500)

        # Set up cleanup when dialog is destroyed
        self.dialog.protocol("WM_DELETE_WINDOW", self.cleanup_and_close)

        self.create_widgets()
        self.center_window()
        debug_print(f"DEBUG: FileSelectionDialog initialized with {len(available_files)} files")

    def center_window(self):
        """Center the dialog on the screen."""
        debug_print("DEBUG: Centering FileSelectionDialog window")
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        """Create the dialog widgets."""
        debug_print("DEBUG: Creating widgets for FileSelectionDialog")

        # Use grid layout for better control
        self.dialog.grid_columnconfigure(0, weight=1)  # Files column expands
        self.dialog.grid_columnconfigure(1, weight=0)  # Buttons column fixed width
        self.dialog.grid_rowconfigure(0, weight=0)  # Header row fixed height
        self.dialog.grid_rowconfigure(1, weight=1)  # Content row expands

        # Header label with white text
        header_frame = ttk.Frame(self.dialog)
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        header_label = ttk.Label(
            header_frame,
            text="Select files to include in your comparison analysis:",
            font=FONT,
            foreground="white"
        )
        header_label.pack(anchor="w")

        # Left side: scrollable files frame
        file_frame = ttk.Frame(self.dialog)
        file_frame.grid(row=1, column=0, sticky="nsew", padx=(10, 5), pady=5)
        file_frame.grid_columnconfigure(0, weight=1)
        file_frame.grid_rowconfigure(0, weight=1)

        # Scrollable frame for checkboxes
        self.canvas = tk.Canvas(file_frame)
        scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=self.canvas.yview)

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

        # Create a variable for each file - all files selected by default
        self.file_vars = {}
        for i, file_data in enumerate(self.available_files):
            # Get display name for the file
            display_name = file_data.get('display_filename', file_data.get('file_name', f'File_{i}'))
            debug_print(f"DEBUG: Adding checkbox for file: {display_name}")

            var = tk.BooleanVar(value=True)  # All files selected by default
            self.file_vars[i] = var  # Use index as key for easier lookup
            cb = ttk.Checkbutton(self.checkbox_frame, text=display_name, variable=var)
            cb.grid(row=i, column=0, sticky="w", pady=2)

        # Right side: buttons frame - fixed width
        button_frame = ttk.Frame(self.dialog, width=120)
        button_frame.grid(row=1, column=1, sticky="ns", padx=(5, 10), pady=5)
        button_frame.grid_propagate(False)  # Keep fixed width

        # Stack buttons vertically with fixed width
        select_all_btn = ttk.Button(button_frame, text="Check All", command=self.select_all)
        select_all_btn.pack(side="top", fill="x", pady=5)

        deselect_all_btn = ttk.Button(button_frame, text="Uncheck All", command=self.deselect_all)
        deselect_all_btn.pack(side="top", fill="x", pady=5)

        # Add a separator
        ttk.Separator(button_frame, orient="horizontal").pack(fill="x", pady=10)

        ok_btn = ttk.Button(button_frame, text="OK", command=self.on_ok)
        ok_btn.pack(side="top", fill="x", pady=5)

        cancel_btn = ttk.Button(button_frame, text="Cancel", command=self.on_cancel)
        cancel_btn.pack(side="top", fill="x", pady=5)

        # Add mouse wheel scrolling with proper error handling
        def _on_mousewheel(event):
            # Check if canvas still exists and is valid
            if self.canvas and self.canvas.winfo_exists():
                try:
                    self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                except tk.TclError as e:
                    debug_print(f"DEBUG: Canvas scrolling error (canvas may be destroyed): {e}")
                    # Unbind the event if canvas is invalid
                    self.cleanup_mousewheel_binding()
            else:
                debug_print("DEBUG: Canvas no longer exists, cleaning up mousewheel binding")
                self.cleanup_mousewheel_binding()

        # Use bind instead of bind_all to limit scope to this canvas
        self.canvas.bind("<MouseWheel>", _on_mousewheel)

        # Also bind to the dialog to capture mousewheel when over the dialog
        self.dialog.bind("<MouseWheel>", _on_mousewheel)

        debug_print("DEBUG: FileSelectionDialog widgets created successfully")

    def cleanup_mousewheel_binding(self):
        """Clean up mousewheel event bindings."""
        try:
            if self.canvas and self.canvas.winfo_exists():
                self.canvas.unbind("<MouseWheel>")
                debug_print("DEBUG: Canvas mousewheel binding cleaned up")
        except tk.TclError:
            debug_print("DEBUG: Canvas already destroyed, binding cleanup not needed")

        try:
            if self.dialog and self.dialog.winfo_exists():
                self.dialog.unbind("<MouseWheel>")
                debug_print("DEBUG: Dialog mousewheel binding cleaned up")
        except tk.TclError:
            debug_print("DEBUG: Dialog already destroyed, binding cleanup not needed")

    def cleanup_and_close(self):
        """Clean up resources and close the dialog when window is closed."""
        debug_print("DEBUG: Window close button clicked - cleaning up FileSelectionDialog resources")
        self.cleanup_mousewheel_binding()
        self.on_cancel()

    def select_all(self):
        """Select all files."""
        debug_print("DEBUG: Check All clicked")
        for var in self.file_vars.values():
            var.set(True)

    def deselect_all(self):
        """Deselect all files."""
        debug_print("DEBUG: Uncheck All clicked")
        for var in self.file_vars.values():
            var.set(False)

    def on_ok(self):
        """Handle OK button click."""
        debug_print("DEBUG: OK button clicked in FileSelectionDialog")
        # Get selected file indices
        selected_indices = [i for i, var in self.file_vars.items() if var.get()]
        # Get actual file data objects
        self.selected_files = [self.available_files[i] for i in selected_indices]
        debug_print(f"DEBUG: Selected {len(self.selected_files)} files for comparison")
        for i, file_data in enumerate(self.selected_files):
            display_name = file_data.get('display_filename', file_data.get('file_name', f'File_{i}'))
            debug_print(f"DEBUG: Selected file {i+1}: {display_name}")

        self.result = True
        self.cleanup_mousewheel_binding()
        self.dialog.destroy()

    def on_cancel(self):
        """Handle Cancel button click."""
        debug_print("DEBUG: FileSelectionDialog cancelled")
        self.result = False
        self.dialog.destroy()

    def show(self):
        """
        Show the dialog and wait for user input.

        Returns:
            tuple: (result, selected_files) where result is True if OK was clicked,
                   and selected_files is the list of selected file data dictionaries.
        """
        debug_print("DEBUG: Showing FileSelectionDialog")

        # Set up proper cleanup for when dialog closes
        def cleanup_on_close():
            self.cleanup_mousewheel_binding()
            self.on_cancel()

        self.dialog.protocol("WM_DELETE_WINDOW", cleanup_on_close)

        self.dialog.wait_window()

        debug_print(f"DEBUG: FileSelectionDialog closed with result: {self.result}")
        if self.result:
            debug_print("DEBUG: Dialog succeeded - proceeding with selected files")
        else:
            debug_print("DEBUG: Dialog was cancelled")

        return self.result, self.selected_files
