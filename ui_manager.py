"""
ui_manager.py
UI layout and window management for DataViewer application.
Handles frame creation, window sizing, styling, and layout management.
"""

import tkinter as tk

from typing import Optional
from tkinter import ttk
from utils import FONT, APP_BACKGROUND_COLOR, BUTTON_COLOR, debug_print


class UIManager:
    """Manages UI layout, window sizing, and frame organization."""
    
    def __init__(self, gui):
        """Initialize UI manager with reference to main GUI.
        
        Args:
            gui: Reference to main DataViewer instance
        """
        self.gui = gui
        self.root = gui.root
    
    def create_static_frames(self) -> None:
        """Create persistent (static) frames that remain for the lifetime of the UI."""
        # Create top_frame for dropdowns and control buttons.
        if not hasattr(self.gui, 'top_frame') or not self.gui.top_frame:
            self.gui.top_frame = ttk.Frame(self.root, height=30)
            self.gui.top_frame.pack(side="top", fill="x", pady=(10, 0), padx=10)
            self.gui.top_frame.pack_propagate(False)
    
        # Create bottom_frame to hold the image button and image display area.
        if not hasattr(self.gui, 'bottom_frame') or not self.gui.bottom_frame:
            self.gui.bottom_frame = ttk.Frame(self.root, height=150)
            self.gui.bottom_frame.pack(side="bottom", fill="x", padx=10, pady=(0, 10))
            self.gui.bottom_frame.pack_propagate(False)
            self.gui.bottom_frame.grid_propagate(False)
    
        # Create a static frame for the Load Images button within bottom_frame.
        if not hasattr(self.gui, 'image_button_frame') or not self.gui.image_button_frame:
            self.gui.image_button_frame = ttk.Frame(self.gui.bottom_frame)
            self.gui.image_button_frame.pack(side="left", fill="y", padx=5, pady=5)
        
            self.gui.crop_enabled = tk.BooleanVar(value=False)
            self.gui.crop_checkbox = ttk.Checkbutton(
                self.gui.image_button_frame,
                text="Auto-Crop (experimental)",
                variable=self.gui.crop_enabled
            )
            self.gui.crop_checkbox.pack(side="top", padx=5, pady=5)
        
            load_image_button = ttk.Button(
                self.gui.image_button_frame,
                text="Load Images",
                command=lambda: self.gui.image_loader.add_images() if self.gui.image_loader else None
            )
            load_image_button.pack(side="left", padx=5, pady=5)
    
        # Create the dynamic image display frame within bottom_frame.
        if not hasattr(self.gui, 'image_frame') or not self.gui.image_frame:
            self.gui.image_frame = ttk.Frame(self.gui.bottom_frame, height=150)
            self.gui.image_frame.pack(side="left", fill="both", expand=True)
            self.gui.image_frame.pack_propagate(False)
            self.gui.image_frame.grid_propagate(False)
    
        # Create display_frame for the table/plot area.
        if not hasattr(self.gui, 'display_frame') or not self.gui.display_frame:
            self.gui.display_frame = ttk.Frame(self.root)
            self.gui.display_frame.pack(fill="both", expand=True, padx=10, pady=5)
    
        # Create a dynamic subframe inside display_frame for table and plot content.
        if not hasattr(self.gui, 'dynamic_frame') or not self.gui.dynamic_frame:
            self.gui.dynamic_frame = ttk.Frame(self.gui.display_frame)
            self.gui.dynamic_frame.pack(fill="both", expand=True)

    def clear_dynamic_frame(self) -> None:
        """Clear all children widgets from the dynamic frame."""
        for widget in self.gui.dynamic_frame.winfo_children():
            widget.destroy()

    def on_window_resize(self, event):
        """Handle window resize events to maintain layout proportions."""
        if event.widget == self.root:
            self.constrain_plot_width()

    def setup_dynamic_frames(self, is_plotting_sheet: bool = False) -> None:
        """Create frames inside the dynamic_frame based on sheet type."""
        # Clear previous widgets
        for widget in self.gui.dynamic_frame.winfo_children():
            widget.destroy()
    
        # Get dynamic frame height
        window_height = self.root.winfo_height()
        top_height = self.gui.top_frame.winfo_height()
        bottom_height = self.gui.bottom_frame.winfo_height()
        padding = 20
        display_height = window_height - top_height - bottom_height - padding
        display_height = max(display_height, 100)
    
        if is_plotting_sheet:
            # Use grid layout for precise control of width proportions
            self.gui.dynamic_frame.columnconfigure(0, weight=5)  # Table column
            self.gui.dynamic_frame.columnconfigure(1, weight=5)  # Plot column
            self.gui.dynamic_frame.rowconfigure(0, weight=1)
        
            # LEFT SIDE: Table area split into top (table) and bottom (notes)
            left_side_frame = ttk.Frame(self.gui.dynamic_frame)
            left_side_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
        
            # Create a PanedWindow to split table and notes vertically
            left_paned = ttk.PanedWindow(left_side_frame, orient="vertical")
            left_paned.pack(fill="both", expand=True)
        
            # Top pane: Table (60% of height)
            self.gui.table_frame = ttk.Frame(left_paned)
            left_paned.add(self.gui.table_frame, weight=3)
        
            # Bottom pane: Sample Notes (40% of height)
            self.gui.notes_frame = ttk.LabelFrame(left_paned, text="Test Sample Notes", padding=5)
            left_paned.add(self.gui.notes_frame, weight=2)
        
            # Configure the notes display area
            self.gui.create_notes_display_area()
        
            # RIGHT SIDE: Plot takes remaining 50% width
            self.gui.plot_frame = ttk.Frame(self.gui.dynamic_frame)
            self.gui.plot_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)
        
            self.constrain_plot_width()
        else:
            # Non-plotting sheets use the full space (no changes here)
            self.gui.table_frame = ttk.Frame(self.gui.dynamic_frame, height=display_height)
            self.gui.table_frame.pack(fill="both", expand=True, padx=5, pady=5)
    
    def constrain_plot_width(self):
        """Ensure the plot doesn't exceed 50% of the window width."""
        if not hasattr(self.gui, 'plot_frame') or not self.gui.plot_frame:
            return
    
        window_width = self.root.winfo_width()
        max_plot_width = window_width // 2
    
        if max_plot_width > 50:
            self.gui.plot_frame.config(width=max_plot_width)
            self.gui.plot_frame.grid_propagate(False)
        
            if hasattr(self.gui, 'table_frame') and self.gui.table_frame:
                self.gui.table_frame.config(width=max_plot_width)
                self.gui.table_frame.grid_propagate(False)
    
    def center_window(self, window: tk.Toplevel, width: Optional[int] = None, height: Optional[int] = None) -> None:
        """Center a given Tkinter window on the screen."""
        window.update_idletasks()
        window_width = width or window.winfo_width()
        window_height = height or window.winfo_height()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        position_x = (screen_width - window_width) // 2
        position_y = (screen_height - window_height) // 2
        window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
    
    def set_window_size(self, width_ratio: float, height_ratio: float) -> None:
        """Set the window size as a percentage of the screen dimensions."""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * width_ratio)
        window_height = int(screen_height * height_ratio)
        position_x = (screen_width - window_width) // 2
        position_y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
    
    def adjust_window_size(self, fixed_width=1500):
        """Adjust the window height dynamically to fit the content while keeping the width constant."""
        if not isinstance(self.root, (tk.Tk, tk.Toplevel)):
            raise ValueError("Expected 'self.root' to be a tk.Tk or tk.Toplevel instance")

        self.root.update_idletasks()
        required_height = self.root.winfo_reqheight()
        screen_height = self.root.winfo_screenheight()
        screen_width = self.root.winfo_screenwidth()
        final_height = min(required_height, screen_height)

        current_x = self.root.winfo_x()
        current_y = self.root.winfo_y()

        self.root.geometry(f"{fixed_width}x{final_height}+{current_x}+{current_y}")
    
    def set_app_colors(self) -> None:
        """Set consistent color theme and fonts for the application."""
        style = ttk.Style()
        self.root.configure(bg=APP_BACKGROUND_COLOR)
        style.configure('TLabel', background=APP_BACKGROUND_COLOR, font=FONT)
        style.configure('TButton', background=BUTTON_COLOR, font=FONT, padding=6)
        style.configure('TCombobox', font=FONT)
        style.map('TCombobox', background=[('readonly', APP_BACKGROUND_COLOR)])

        for widget in self.root.winfo_children():
            try:
                widget.configure(bg='#EFEFEF')
            except Exception:
                continue
    
    def add_static_controls(self) -> None:
        """Add static controls (Add Data and Trend Analysis buttons) to the top_frame."""
        if not hasattr(self.gui, 'controls_frame') or self.gui.controls_frame is None:
            self.gui.controls_frame = ttk.Frame(self.gui.top_frame)
            self.gui.controls_frame.pack(side="right", fill="x", padx=5, pady=5)