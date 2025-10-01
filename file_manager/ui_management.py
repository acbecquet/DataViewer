"""
UI Management Module for DataViewer Application

This module handles user interface updates, initialization, and UI-related helper functions.
"""

# Standard library imports
import os
import copy

# Third party imports
import tkinter as tk
from tkinter import ttk

# Local imports
from utils import debug_print, FONT, APP_BACKGROUND_COLOR


class UIManager:
    """Handles UI updates, initialization, and interface management."""
    
    def __init__(self, file_manager):
        """Initialize with reference to parent FileManager."""
        self.file_manager = file_manager
        self.gui = file_manager.gui
        self.root = file_manager.root
        
    def update_ui_for_current_file(self) -> None:
        """Update UI components to reflect the currently active file."""
        if not self.gui.current_file:
            return

        #debug_print(f"DEBUG: Updating UI for current file: {self.gui.current_file}")

        # Ensure main GUI is properly initialized first
        self.ensure_main_gui_initialized()

        # Update file dropdown
        if hasattr(self.gui, 'file_dropdown_var'):
            self.gui.file_dropdown_var.set(self.gui.current_file)
            #debug_print(f"DEBUG: Set file dropdown to: {self.gui.current_file}")

        # Update sheet dropdown
        self.gui.populate_or_update_sheet_dropdown()
        #debug_print("DEBUG: Updated sheet dropdown")

        # Update displayed sheet
        current_sheet = self.gui.selected_sheet.get()
        if current_sheet not in self.gui.filtered_sheets:
            first_sheet = list(self.gui.filtered_sheets.keys())[0] if self.gui.filtered_sheets else None
            if first_sheet:
                self.gui.selected_sheet.set(first_sheet)
                #debug_print(f"DEBUG: Set selected sheet to first sheet: {first_sheet}")
                # Add a small delay to ensure UI is ready
                self.gui.root.after(100, lambda: self.gui.update_displayed_sheet(first_sheet))
            else:
                debug_print("ERROR: No sheets available to display")
        else:
            debug_print(f"DEBUG: Using current sheet: {current_sheet}")
            # Add a small delay to ensure UI is ready
            self.gui.root.after(100, lambda: self.gui.update_displayed_sheet(current_sheet))

        #debug_print("DEBUG: UI update for current file complete")

    def ensure_main_gui_initialized(self):
        """Ensure the main GUI components are properly initialized before displaying data."""
        debug_print("DEBUG: Ensuring main GUI is properly initialized")

        try:
            # Force the main window to update and initialize
            self.gui.root.update_idletasks()

            # Ensure basic frames exist
            if not hasattr(self.gui, 'top_frame') or not self.gui.top_frame:
                debug_print("DEBUG: Creating missing top_frame")
                self.gui.create_static_frames()

            # Ensure file dropdown exists
            if not hasattr(self.gui, 'file_dropdown') or not self.gui.file_dropdown:
                debug_print("DEBUG: Creating missing file dropdown")
                self.add_or_update_file_dropdown()

            # Ensure sheet dropdown exists
            if not hasattr(self.gui, 'drop_down_menu') or not self.gui.drop_down_menu:
                debug_print("DEBUG: Creating missing sheet dropdown")
                self.gui.populate_or_update_sheet_dropdown()

            # Ensure display frame exists
            if not hasattr(self.gui, 'display_frame') or not self.gui.display_frame:
                debug_print("DEBUG: Creating missing display_frame")
                self.gui.create_static_frames()

            # Force another update
            self.gui.root.update_idletasks()

            debug_print("DEBUG: Main GUI initialization check complete")

        except Exception as e:
            debug_print(f"ERROR: Failed to ensure main GUI initialization: {e}")
            import traceback
            traceback.print_exc()

    def update_file_dropdown(self) -> None:
        """Update the file dropdown with loaded file names."""
        file_names = [file_data["file_name"] for file_data in self.gui.all_filtered_sheets]
        self.gui.file_dropdown["values"] = file_names
        if file_names:
            self.gui.file_dropdown_var.set(file_names[-1])
        else:
            # Clear the dropdown when no files remain
            self.gui.file_dropdown_var.set('')
        self.gui.file_dropdown.update_idletasks()

    def add_or_update_file_dropdown(self) -> None:
        """Add a file selection dropdown or update its values if it already exists."""
        if not hasattr(self.gui, 'file_dropdown') or not self.gui.file_dropdown:
            dropdown_frame = ttk.Frame(self.gui.top_frame, width=1400, height=40)
            dropdown_frame.pack(side="left", pady=2, padx=5)
            file_label = ttk.Label(dropdown_frame, text="Select File:", font=FONT, background=APP_BACKGROUND_COLOR)
            file_label.pack(side="left", padx=(0, 0))
            self.gui.file_dropdown_var = tk.StringVar()
            self.gui.file_dropdown = ttk.Combobox(
                dropdown_frame,
                textvariable=self.gui.file_dropdown_var,
                state="readonly",
                font=FONT,
                width=20
            )
            self.gui.file_dropdown.pack(side="left", fill="x", expand=True, padx=(5, 5))
            self.gui.file_dropdown.bind("<<ComboboxSelected>>", self.gui.on_file_selection)
        self.update_file_dropdown()

    def center_window(self, window, width, height):
        """Center a window on the screen."""
        try:
            # Get screen dimensions
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()

            # Calculate center position
            x = (screen_width - width) // 2
            y = (screen_height - height) // 2

            # Set window geometry with center position
            window.geometry(f"{width}x{height}+{x}+{y}")

            debug_print(f"DEBUG: Centered window {width}x{height} at position ({x}, {y}) on {screen_width}x{screen_height} screen")

        except Exception as e:
            debug_print(f"DEBUG: Error centering window: {e}")
            # Fallback to basic geometry if centering fails
            window.geometry(f"{width}x{height}")

    def start_file_loading_wrapper(self, startup_menu: tk.Toplevel) -> None:
        """Handle the 'Load' button click in the startup menu."""
        startup_menu.destroy()
        self.file_manager.load_initial_file()

    def start_file_loading_database_wrapper(self, startup_menu: tk.Toplevel) -> None:
        """Handle the 'Load from Database' button click in the startup menu."""
        startup_menu.destroy()
        self.file_manager.show_database_browser()