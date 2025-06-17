"""
main_gui.py
Developed By Charlie Becquet.
Main GUI module for the DataViewer Application.
This module instantiates the high-level TestingGUI class which delegates file I/O,
plotting, trend analysis, report generation, and progress dialog management
to separate modules for better modularity and scalability.
"""
import copy
import queue
import os
import threading
import tkinter as tk
import pandas as pd
import processing
import numpy as np
from typing import Optional, Dict, List, Any
import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
from tkintertable import TableCanvas, TableModel
import matplotlib
matplotlib.use('TkAgg')  # Ensure Matplotlib uses TkAgg backend

# Import our new manager classes and utility functions.
from processing import get_valid_plot_options
from plot_manager import PlotManager
from file_manager import FileManager
from report_generator import ReportGenerator
from trend_analysis_gui import TrendAnalysisGUI
from progress_dialog import ProgressDialog
from image_loader import ImageLoader
from viscosity_calculator import ViscosityCalculator
from utils import FONT, get_resource_path, clean_columns, get_save_path, is_standard_file, plotting_sheet_test, APP_BACKGROUND_COLOR,BUTTON_COLOR, PLOT_CHECKBOX_TITLE

class TestingGUI:
    """Main GUI class for the Standardized Testing application."""

    def __init__(self, root):
        self.root = root
        self.root.title("Standardized Testing GUI")
        self.report_thread = None
        self.report_queue = queue.Queue()
        self.initialize_variables()
        self.configure_ui()
        self.show_startup_menu()
        self.image_loader = None # Initialize ImageLoader placeholder
        # Bind the window close event
        self.root.protocol("WM_DELETE_WINDOW", self.on_app_close)


    # Initialization and Configuration
    def initialize_variables(self) -> None:
        """Initialize variables used throughout the GUI."""
        self.sheets: Dict[str, pd.DataFrame] = {}
        self.filtered_sheets: Dict[str, pd.DataFrame] = {}
        self.all_filtered_sheets: List[Dict[str, Any]] = []
        self.current_file: Optional[str] = None
        self.selected_sheet = tk.StringVar()
        self.selected_plot_type = tk.StringVar(value="TPM")
        self.file_path: Optional[str] = None
        self.plot_dropdown = None
        self.threads = [] # track active threads
        self.canvas = None
        self.figure = None
        self.axes = None
        self.lines = None
        self.sheet_images = {}
        self.lock = threading.Lock()
        self.line_labels = []
        self.original_lines_data = []
        self.checkbox_cid = None
    
        # Add different plot options for different tests:
        self.standard_plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        self.user_test_simulation_plot_options = ["TPM", "Draw Pressure", "Power Efficiency", "TPM (Bar)"]  # No Resistance
        self.plot_options = self.standard_plot_options  # Default to standard
    
        self.check_buttons = None
        self.previous_window_geometry = None
        self.is_user_test_simulation = False  
        self.num_columns_per_sample = 12 

        self.crop_enable = tk.BooleanVar(value = False)
    
        # Instantiate manager classes
        self.file_manager = FileManager(self)
        self.plot_manager = PlotManager(self)
        self.report_generator = ReportGenerator(self)
        self.progress_dialog = ProgressDialog(self.root)
        self.viscosity_calculator = ViscosityCalculator(self)
        # (TrendAnalysisGUI will be created when needed)

    def configure_ui(self) -> None:
        """Configure the UI appearance and set application properties."""
        icon_path = get_resource_path('resources/ccell_icon.png')
        self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
        self.set_app_colors()
        self.set_window_size(0.8, 0.8)
        self.root.minsize(1200,800)
        self.center_window(self.root)
        #self.root.bind("<Configure>", lambda e: self.on_window_resize(e))
        self.create_static_frames()
        self.add_menu()
        self.file_manager.add_or_update_file_dropdown()
        self.add_static_controls()

    # === New Centralized Frame Creation Methods ===
    def add_static_controls(self) -> None:
        """Add static controls (Add Data and Trend Analysis buttons) to the top_frame."""
        #print("DEBUG: Adding static controls to top_frame...")
        
        if not hasattr(self, 'controls_frame') or not self.controls_frame:
            self.controls_frame = ttk.Frame(self.top_frame)
            
            self.controls_frame.pack(side="right", fill="x", padx=5, pady=5)
            #print("DEBUG: controls_frame created in top_frame.")
        # Add the "Add Data" button
        #add_data_btn = ttk.Button(self.controls_frame, text="Add Data", command=self.file_manager.add_data)
        #add_data_btn.pack(side="left", padx=(5, 5), pady=(5, 5))
        #print("DEBUG: 'Add Data' button added to static controls.")
        # Add the "Trend Analysis" button
        #trend_button = ttk.Button(self.controls_frame, text="Trend Analysis", command=self.open_trend_analysis_window)
        #trend_button.pack(side="left", padx=(5, 5), pady=(5, 5))
        #print("DEBUG: 'Trend Analysis' button added to static controls.")

        #self.add_report_buttons(self.controls_frame)

    def create_static_frames(self) -> None:
        """Create persistent (static) frames that remain for the lifetime of the UI."""
        #print("DEBUG: Creating static frames...")

        # Create top_frame for dropdowns and control buttons.
        if not hasattr(self, 'top_frame') or not self.top_frame:
            self.top_frame = ttk.Frame(self.root,height = 30)
            self.top_frame.pack(side="top", fill="x", pady=(10, 0), padx=10)
            self.top_frame.pack_propagate(False) # Prevent height changes
            #print("DEBUG: top_frame created.")

        # Create bottom_frame to hold the image button and image display area.
        if not hasattr(self, 'bottom_frame') or not self.bottom_frame:
            self.bottom_frame = ttk.Frame(self.root, height = 150)
            self.bottom_frame.pack(side="bottom", fill = "x", padx=10, pady=(0,10))
            self.bottom_frame.pack_propagate(False)
            self.bottom_frame.grid_propagate(False)
            #print(f"DEBUG: Created bottom_frame with fixed height 150 | "
              #f"Current height: {self.bottom_frame.winfo_height()}")

        # Create a static frame for the Load Images button within bottom_frame.
        if not hasattr(self, 'image_button_frame') or not self.image_button_frame:
            self.image_button_frame = ttk.Frame(self.bottom_frame)
            # Pack it to the left.
            self.image_button_frame.pack(side="left", fill="y", padx=5, pady=5)
            #print("DEBUG: image_button_frame created.")

            self.crop_enabled = tk.BooleanVar(value=False) # Default: Cropping is enabled
            self.crop_checkbox = ttk.Checkbutton(
                self.image_button_frame,
                text = "Auto-Crop (experimental)",
                variable = self.crop_enabled
            )

            self.crop_checkbox.pack(side = "top", padx = 5, pady = 5)

            load_image_button = ttk.Button(
                self.image_button_frame,
                text="Load Images",
                command=lambda: self.image_loader.add_images() if self.image_loader else None
            )
            load_image_button.pack(side="left", padx=5, pady=5)
            #print("DEBUG: Static 'Load Images' button added to image_button_frame.")

        # Create the dynamic image display frame within bottom_frame.
        if not hasattr(self, 'image_frame') or not self.image_frame:
            self.image_frame = ttk.Frame(self.bottom_frame, height = 150)
            self.image_frame.pack(side="left", fill="both", expand=True)
            self.image_frame.pack_propagate(False)
            self.image_frame.grid_propagate(False)
            # Prevent the image_frame from shrinking.
            #self.image_frame.pack_propagate(True)
            #print("DEBUG: Dynamic image_frame created with fixed height 300.")


        # Create display_frame for the table/plot area.
        if not hasattr(self, 'display_frame') or not self.display_frame:
            self.display_frame = ttk.Frame(self.root)
            self.display_frame.pack(fill="both", expand=True, padx=10, pady=5)
            #print("DEBUG: display_frame created.")

        # Create a dynamic subframe inside display_frame for table and plot content.
        if not hasattr(self, 'dynamic_frame') or not self.dynamic_frame:
            self.dynamic_frame = ttk.Frame(self.display_frame)
            self.dynamic_frame.pack(fill="both", expand=True)
            #print("DEBUG: dynamic_frame created.")

    def clear_dynamic_frame(self) -> None:
        """Clear all children widgets from the dynamic frame."""
        #print("DEBUG: Clearing dynamic_frame contents...")
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()

    def on_window_resize(self, event):
        """Handle window resize events to maintain layout proportions."""
        # Only process if this is a window resize, not a child widget configure event
        if event.widget == self.root:
            # Update the width constraints
            self.constrain_plot_width()

    def setup_dynamic_frames(self, is_plotting_sheet: bool = False) -> None:
        """Create frames inside the dynamic_frame based on sheet type."""
        # Clear previous widgets
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()

        # Get dynamic frame height
        window_height = self.root.winfo_height()
        top_height = self.top_frame.winfo_height()
        bottom_height = self.bottom_frame.winfo_height()
        padding = 20
        display_height = window_height - top_height - bottom_height - padding
        display_height = max(display_height, 100)

        if is_plotting_sheet:
            # Use grid layout for precise control of width proportions
            self.dynamic_frame.columnconfigure(0, weight=5)  # Table column
            self.dynamic_frame.columnconfigure(1, weight=5)  # Plot column
        
            self.dynamic_frame.rowconfigure(0, weight = 1)

            # Table takes exactly 50% width
            self.table_frame = ttk.Frame(self.dynamic_frame)
            self.table_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
        
            # Plot takes remaining 50% width
            self.plot_frame = ttk.Frame(self.dynamic_frame)
            self.plot_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)
        
            self.constrain_plot_width()
        else:
            # Non-plotting sheets use the full space
            self.table_frame = ttk.Frame(self.dynamic_frame, height=display_height)
            self.table_frame.pack(fill="both", expand=True, padx=5, pady=5)

    def constrain_plot_width(self):
        """
        Ensure the plot doesn't exceed 50% of the window width by setting
        a maximum width constraint.
        """
        if not hasattr(self, 'plot_frame') or not self.plot_frame:
            return
        
        # Get the current window width
        window_width = self.root.winfo_width()
    
        # Calculate 50% of the window width
        max_plot_width = window_width // 2
    
        if max_plot_width > 50:  # Only apply if we have a reasonable width
            # Set the maximum width for the plot frame
            self.plot_frame.config(width=max_plot_width)
        
            # Prevent the plot from expanding beyond this width
            self.plot_frame.grid_propagate(False)
        
            # Also set the table frame to the same width for balance
            if hasattr(self, 'table_frame') and self.table_frame:
                self.table_frame.config(width=max_plot_width)
                self.table_frame.grid_propagate(False)


    def show_startup_menu(self) -> None:
        """Display a startup menu with 'New' and 'Load' options."""
        startup_menu = Toplevel(self.root)
        startup_menu.title("Welcome")
        startup_menu.geometry("300x150")
        startup_menu.transient(self.root)
        startup_menu.grab_set()  # Make the window modal

        label = ttk.Label(startup_menu, text="Welcome to DataViewer by SDR!", font=FONT,background=APP_BACKGROUND_COLOR, foreground="white")
        label.pack(pady=10)

        # New Button
        new_button = ttk.Button(
            startup_menu,
            text="New",
            command=lambda: self.file_manager.create_new_template(startup_menu)
        )
        new_button.pack(pady=5)

        # Load Button
        load_button = ttk.Button(
            startup_menu,
            text="Load",
            command=lambda: self.file_manager.start_file_loading_wrapper(startup_menu)
        )
        load_button.pack(pady=5)

        self.center_window(startup_menu)

    def center_window(self, window: tk.Toplevel, width: Optional[int] = None, height: Optional[int] = None) -> None:
        """
        Center a given Tkinter window on the screen.
    
        Args:
            window (tk.Toplevel): The Tkinter window to be centered.
            width (Optional[int]): Desired width of the window. If not provided, the current width is used.
            height (Optional[int]): Desired height of the window. If not provided, the current height is used.
        """
        
        window.update_idletasks()  # Ensure all geometry calculations are updated

        # Get the current dimensions of the window if width/height are not specified
        window_width = width or window.winfo_width()
        window_height = height or window.winfo_height()

        # Get the screen dimensions
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        # Calculate the position for centering
        position_x = (screen_width - window_width) // 2
        position_y = (screen_height - window_height) // 2

        # Set the geometry with the calculated position
        window.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    def populate_or_update_sheet_dropdown(self) -> None:
        """Populate or update the dropdown for sheet selection."""
        if not hasattr(self, 'drop_down_menu') or not self.drop_down_menu:
            # Create a dropdown for sheet selection
            sheet_label = ttk.Label(self.top_frame, text = "Select Test:", font=FONT, foreground='white', background = APP_BACKGROUND_COLOR)
            sheet_label.pack(side = "left", padx = (0,5))
            self.drop_down_menu = ttk.Combobox(
                self.top_frame,
                textvariable=self.selected_sheet,
                state="readonly",
                font=FONT
            )
            self.drop_down_menu.pack(side = "left", pady=(5, 5))
            
            # Bind selection event to update the displayed sheet
            self.drop_down_menu.bind(
                "<<ComboboxSelected>>",
                lambda event: self.update_displayed_sheet(self.selected_sheet.get())
            )

        # Update dropdown values with sheet names
        all_sheet_names = list(self.filtered_sheets.keys())
        self.drop_down_menu["values"] = all_sheet_names
        current_selection = self.selected_sheet.get()
        if current_selection not in all_sheet_names and all_sheet_names:
            self.selected_sheet.set(all_sheet_names[0])

    def on_file_selection(self, event) -> None:
        """Handle file selection from the dropdown."""
        selected_file = self.file_dropdown_var.get()
        if not selected_file or selected_file == self.current_file:
            return

        try:
            self.file_manager.set_active_file(selected_file)
            self.file_manager.update_ui_for_current_file()
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def add_trend_analysis_button(self, parent_frame: ttk.Frame) -> None:
        """Add a 'Trend Analysis' button to the UI."""
        trend_button = ttk.Button(
            parent_frame,
            text="Trend Analysis", state = "enabled",
            command=self.open_trend_analysis_window
        )
        trend_button.pack(side="right", padx=(5, 10), pady=(5, 5))  # disabled for now
        
    def on_app_close(self):
        """Handle application shutdown."""
        try:
            for thread in self.threads:
                if thread.is_alive():
                    #print(f"Stopping thread {thread.name}")
                    pass
            self.root.destroy()
            os._exit(0)
        except Exception as e:
            print(f"Error during shutdown: {e}")
            os._exit(1)

    def add_menu(self) -> None:
        """Create a top-level menu with File, Help, and About options."""
        menubar = tk.Menu(self.root)
    
        # File menu
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=self.show_new_template_dialog)
        filemenu.add_command(label="Load Excel", command=self.file_manager.load_initial_file)
        filemenu.add_command(label="Load VAP3", command=self.open_vap3_file)
        filemenu.add_separator()
        filemenu.add_command(label="Save As VAP3", command=self.save_as_vap3)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.on_app_close)
        menubar.add_cascade(label="File", menu=filemenu)
    
        # View menu
        viewmenu = tk.Menu(menubar, tearoff=0)
        viewmenu.add_command(label="View Raw Data", 
                             command=lambda: self.file_manager.open_raw_data_in_excel(self.selected_sheet.get()))
        viewmenu.add_command(label="Trend Analysis", command=self.open_trend_analysis_window)
        viewmenu.add_separator()
        viewmenu.add_command(label="Collect Data", command=self.open_data_collection)
        menubar.add_cascade(label="View", menu=viewmenu)

        # Database menu
        dbmenu = tk.Menu(menubar, tearoff=0)
        dbmenu.add_command(label="Browse Database", command=self.file_manager.show_database_browser)
        dbmenu.add_command(label="Load from Database", command=self.file_manager.load_from_database)
        menubar.add_cascade(label="Database", menu=dbmenu)

       # Calculate menu
        calculatemenu = tk.Menu(menubar, tearoff=0)
        calculatemenu.add_command(label="Viscosity", command=self.open_viscosity_calculator)
        menubar.add_cascade(label="Calculate", menu=calculatemenu)

        # Compare menu (empty for now)
        comparemenu = tk.Menu(menubar, tearoff=0)
        # Future comparison functionality will go here
        menubar.add_cascade(label="Compare", menu=comparemenu)

        # Reports menu
        reportmenu = tk.Menu(menubar, tearoff=0)
        reportmenu.add_command(label="Generate Test Report", command=self.generate_test_report)
        reportmenu.add_command(label="Generate Full Report", command=self.generate_full_report)
        menubar.add_cascade(label="Reports", menu=reportmenu)
    
        # Help menu
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Help", command=self.show_help)
        helpmenu.add_separator()
        helpmenu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)
    
        self.root.config(menu=menubar)

    def show_new_template_dialog(self) -> None:
        """Show a dialog to create a new template file with selected tests."""
        self.file_manager.create_new_file_with_tests()

    def show_new_template_dialog_old(self) -> None:
        """Show a dialog to create a new template file."""
        startup_menu = Toplevel(self.root)
        startup_menu.title("New File")
        startup_menu.geometry("300x100")
        startup_menu.transient(self.root)
        startup_menu.grab_set()  # Make the window modal
    
        new_button = ttk.Button(
            startup_menu,
            text="Create New Template",
            command=lambda: self.file_manager.create_new_template(startup_menu)
        )
        new_button.pack(pady=20)
    
        self.center_window(startup_menu)

    def open_vap3_file(self) -> None:
        """Open a .vap3 file using the file manager."""
        self.file_manager.load_vap3_file()

    def save_as_vap3(self) -> None:
        """Save the current session as a .vap3 file."""
        self.file_manager.save_as_vap3()

    def show_help(self):
        """Display help dialog."""
        messagebox.showinfo("Help", "This program is designed to be used with excel data according to the SDR Standardized Testing Template.\n \nClick 'Generate Test Report' to create an excel report of a single test, or click 'Generate Full Report' to generate both an excel file and powerpoint file of all the contents within the file.")

    def show_about(self) -> None:
        """Display about dialog."""
        messagebox.showinfo("About", "SDR DataViewer Beta Version 3.0\nDeveloped by Charlie Becquet")

    def set_app_colors(self) -> None:
        """Set consistent color theme and fonts for the application."""
        style = ttk.Style()
        self.root.configure(bg=APP_BACKGROUND_COLOR)  # Set background color
        style.configure('TLabel', background=APP_BACKGROUND_COLOR, font=FONT)
        style.configure('TButton', background=BUTTON_COLOR, font=FONT, padding=6)
        style.configure('TCombobox', font=FONT)
        style.map('TCombobox', background=[('readonly', APP_BACKGROUND_COLOR)])
        # Set color for all widgets, loop through children
        # Modified widget configuration loop
        for widget in self.root.winfo_children():
            try:
                # Only works for standard tkinter widgets
                widget.configure(bg='#EFEFEF')
            except Exception:
                # Skip ttk widgets that don't support bg
                continue

    def populate_sheet_dropdown(self) -> None:
        """Set up the dropdown for sheet selection and handle its events."""
        all_sheet_names = list(self.filtered_sheets.keys())

        # Create a dropdown menu for selecting sheets
        self.drop_down_menu = ttk.Combobox(
            self.top_frame,
            textvariable=self.selected_sheet,
            values=list(self.filtered_sheets.keys()),
            state="readonly"  # Prevent users from typing in custom values
        )
        self.drop_down_menu.pack(pady=(10, 10))
        self.drop_down_menu.place(relx=0.25, rely=0.5, anchor="center")

        # Bind selection event
        self.drop_down_menu.bind(
            "<<ComboboxSelected>>",
            lambda event: self.update_displayed_sheet(self.selected_sheet.get())
        )

        # Automatically select the first sheet if available
        if all_sheet_names:
            first_sheet = all_sheet_names[0]
            self.selected_sheet.set(first_sheet)
            self.update_displayed_sheet(first_sheet)
        else:
            messagebox.showerror("Error", "No sheets found in the file.")
       
    def set_window_size(self, width_ratio: float, height_ratio: float) -> None:
        """
        Set the window size as a percentage of the screen dimensions.

        Args:
            width_ratio (float): Width as a fraction of screen width.
            height_ratio (float): Height as a fraction of screen height.
        """
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * width_ratio)
        window_height = int(screen_height * height_ratio)
        position_x = (screen_width - window_width) // 2
        position_y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    def add_report_buttons(self, parent_frame: ttk.Frame) -> None:
        """Add buttons for generating reports and align them in the parent frame."""

        # Add Data Button
        #add_data_btn = ttk.Button(parent_frame, text="Add Data", command=self.file_manager.add_data)
        #add_data_btn.pack(side="right", padx=(5, 5), pady=(5, 5))

        # Generate Full Report button
        full_report_btn = ttk.Button(parent_frame, text="Generate Full Report", command=self.generate_full_report)
        full_report_btn.pack(side="left", padx=(5, 5), pady=(5, 5))

        # Generate Test Report button
        test_report_btn = ttk.Button(parent_frame, text="Generate Test Report", command=self.generate_test_report)
        test_report_btn.pack(side="left", padx=(5, 5), pady=(5, 5))

    def adjust_window_size(self, fixed_width=1500):
        """
        Adjust the window height dynamically to fit the content while keeping the width constant.
        Args:
            fixed_width (int): The fixed width of the window.
        """

        if not isinstance(self.root, (tk.Tk, tk.Toplevel)):
            raise ValueError("Expected 'self.root' to be a tk.Tk or tk.Toplevel instance")

        # Update the geometry manager and calculate the required size
        self.root.update_idletasks()

        # Get the required dimensions for the window height to fit its content
        required_height = self.root.winfo_reqheight()

        # Get the current screen dimensions
        screen_height = self.root.winfo_screenheight()
        screen_width = self.root.winfo_screenwidth()
        # Constrain the height to the screen dimensions if necessary
        final_height = min(required_height, screen_height)

        current_x = self.root.winfo_x()
        current_y = self.root.winfo_y()

        # Set the geometry dynamically with the fixed width
        self.root.geometry(f"{fixed_width}x{final_height}+{current_x}+{current_y}")


    def update_displayed_sheet(self, sheet_name: str) -> None:
        """
        Update the displayed sheet and dynamically manage the plot options and plot type dropdown.
        Enhanced to handle empty sheets for data collection.
        """
        print(f"DEBUG: [update_displayed_sheet] START for sheet: {sheet_name}")

        # Ensure bottom_frame maintains its fixed height
        if hasattr(self, 'bottom_frame') and self.bottom_frame.winfo_exists():
            self.bottom_frame.configure(height=150)
            self.bottom_frame.pack_propagate(False)
            self.bottom_frame.grid_propagate(False)

        if not sheet_name or sheet_name not in self.filtered_sheets:
            print(f"DEBUG: Sheet {sheet_name} not found in filtered_sheets")
            return

        sheet_info = self.filtered_sheets.get(sheet_name)
        if not sheet_info:
            messagebox.showerror("Error", f"Sheet '{sheet_name}' not found.")
            return

        data = sheet_info["data"]
        is_empty = sheet_info.get("is_empty", False)
        is_plotting_sheet = plotting_sheet_test(sheet_name, data)
        
        print(f"DEBUG: Sheet {sheet_name} - is_empty: {is_empty}, is_plotting_sheet: {is_plotting_sheet}, data_shape: {data.shape}")

        # Clear and rebuild only the dynamic content
        self.clear_dynamic_frame()
        self.setup_dynamic_frames(is_plotting_sheet)

        # Handle ImageLoader setup
        print("DEBUG: Setting up ImageLoader...")
        if hasattr(self, 'image_loader'):
            for widget in self.image_frame.winfo_children():
                widget.destroy()
            del self.image_loader

        # Initialize a fresh ImageLoader
        self.image_loader = ImageLoader(self.image_frame, is_plotting_sheet, on_images_selected=lambda paths: self.store_images(sheet_name, paths), main_gui = self)
        self.image_loader.frame.config(height=150)

        # Restore images if they exist
        current_file = self.current_file
        if current_file in self.sheet_images and sheet_name in self.sheet_images[current_file]:
            self.image_loader.load_images_from_list(self.sheet_images[current_file][sheet_name])
            for img_path in self.sheet_images[current_file][sheet_name]:
                if img_path in self.image_crop_states:
                    self.image_loader.image_crop_states[img_path] = self.image_crop_states[img_path]
            print(f"DEBUG: Restored images and crop states for sheet: {sheet_name}")

        self.image_loader.display_images()
        self.image_frame.update_idletasks()

        # Check if sheet is completely empty (for data collection, we want to handle this gracefully)
        if is_empty or data.empty:
            print(f"DEBUG: Sheet {sheet_name} is empty - checking if this is for data collection")
            
            # For data collection, we still want to show the interface
            # Instead of showing an error, create a minimal structure
            if is_plotting_sheet:
                print("DEBUG: Creating minimal structure for empty plotting sheet")
                # Create minimal data for display
                minimal_data = pd.DataFrame([{
                    "Sample Name": "No data - ready for collection",
                    "Media": "",
                    "Viscosity": "",
                    "Voltage, Resistance, Power": "",
                    "Average TPM": "No data",
                    "Standard Deviation": "No data",
                    "Initial Oil Mass": "",
                    "Usage Efficiency": "",
                    "Burn?": "",
                    "Clog?": "",
                    "Leak?": ""
                }])
                
                # Display the minimal table
                self.display_table(self.table_frame, minimal_data, sheet_name, is_plotting_sheet)
                
                # Create empty plot area with message
                if hasattr(self, 'plot_frame') and self.plot_frame:
                    for widget in self.plot_frame.winfo_children():
                        widget.destroy()
                    
                    empty_plot_label = tk.Label(
                        self.plot_frame,
                        text="No data to plot yet.\nUse 'Collect Data' to add measurements.",
                        font=("Arial", 12),
                        fg="gray",
                        justify="center"
                    )
                    empty_plot_label.pack(expand=True)
                    
                print("DEBUG: Minimal plotting sheet structure created")
            else:
                # For non-plotting sheets, show a simple message
                print("DEBUG: Creating minimal structure for empty non-plotting sheet")
                minimal_data = pd.DataFrame([{
                    "Status": "No data - ready for collection",
                    "Instructions": "Use data collection tools to add information"
                }])
                self.display_table(self.table_frame, minimal_data, sheet_name, is_plotting_sheet)
                
            print("DEBUG: Empty sheet handled for data collection")
            self.root.update_idletasks()
            return

        # Process the sheet data
        try:
            print(f"DEBUG: Processing sheet data for {sheet_name}...")
            process_function = processing.get_processing_function(sheet_name)
            processed_data, _, full_sample_data = process_function(data)
            print(f"DEBUG: Processing complete - processed_data: {processed_data.shape}, full_sample_data: {full_sample_data.shape}")
            #SPECIAL HANDLING FOR USER TEST SIMULATION:
            #SPECIAL HANDLING FOR USER TEST SIMULATION:
            if sheet_name in ["User Test Simulation", "User Simulation Test"]:
                print("DEBUG: Detected User Test Simulation - using 8 columns per sample")
                self.plot_options = self.user_test_simulation_plot_options
                # Store the fact that this is User Test Simulation for plotting
                self.is_user_test_simulation = True
                self.num_columns_per_sample = 8
            else:
                self.plot_options = self.standard_plot_options
                self.is_user_test_simulation = False
                self.num_columns_per_sample = 12

        except Exception as e:
            print(f"ERROR: Processing function failed for {sheet_name}: {e}")
            messagebox.showerror("Processing Error", f"Error processing sheet '{sheet_name}': {e}")
            return

        # Display the table
        print(f"DEBUG: Displaying table for {sheet_name}...")
        try:
            self.display_table(self.table_frame, processed_data, sheet_name, is_plotting_sheet)
            print("DEBUG: Table displayed successfully")
        except Exception as e:
            print(f"ERROR: Failed to display table: {e}")

        # Display plot if it's a plotting sheet
        if is_plotting_sheet:
            print(f"DEBUG: Displaying plot for sheet: {sheet_name}")
            try:
                if not full_sample_data.empty:
                    self.display_plot(full_sample_data)
                    print("DEBUG: Plot displayed successfully")
                else:
                    print("DEBUG: Full sample data is empty, showing empty plot message")
                    if hasattr(self, 'plot_frame') and self.plot_frame:
                        for widget in self.plot_frame.winfo_children():
                            widget.destroy()
                        
                        empty_plot_label = tk.Label(
                            self.plot_frame,
                            text="No data to plot yet.\nUse 'Collect Data' to add measurements.",
                            font=("Arial", 12),
                            fg="gray",
                            justify="center"
                        )
                        empty_plot_label.pack(expand=True)
            except Exception as e:
                print(f"ERROR: Failed to display plot: {e}")

        print("DEBUG: Successfully updated displayed sheet.")
        self.root.update_idletasks()
        
        print(f"DEBUG: [update_displayed_sheet] END for sheet: {sheet_name}")

    def load_images(self, is_plotting_sheet):
        """Load and display images in the dynamic image_frame."""
        if not hasattr(self, 'image_frame') or not self.image_frame.winfo_exists():
            #print("Error: image_frame does not exist, recreating...")
            # Recreate image_frame in bottom_frame if needed.
            self.image_frame = ttk.Frame(self.bottom_frame, height=150)
            self.image_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            self.image_frame.pack_propagate(True)
        if not self.image_loader:
            self.image_loader = ImageLoader(self.image_frame, is_plotting_sheet)
        self.image_loader.add_images()

    def store_images(self, sheet_name, paths):
        """Store image paths and their crop states for a specific sheet."""
        if not sheet_name:
            return

        #print(f"DEBUG: Storing images for sheet: {sheet_name}")
        
        # Ensure the file key exists in the sheet_images dictionary
        if self.current_file not in self.sheet_images:
            self.sheet_images[self.current_file] = {}
        # Store images under both the file and sheet name
        self.sheet_images[self.current_file][sheet_name] = paths

        # Store crop states for each image
        if not hasattr(self, 'image_crop_states'):
            self.image_crop_states = {}

        for img_path in paths:
            # Ensure the image crop state is stored
            if img_path not in self.image_crop_states:
                self.image_crop_states[img_path] = self.crop_enabled.get()

        #print(f"DEBUG: Stored image paths: {self.sheet_images[self.current_file][sheet_name]}")
        #print(f"DEBUG: Stored crop states: {self.image_crop_states}")

    def open_data_collection(self):
        """Open data collection interface, handling both loaded and unloaded file states."""
        print("DEBUG: Data collection requested")
        
        # Check if we have a file loaded
        if not hasattr(self, 'filtered_sheets') or not self.filtered_sheets or not hasattr(self, 'file_path') or not self.file_path:
            print("DEBUG: No file loaded - showing data collection startup dialog")
            self.show_data_collection_startup_dialog()
        else:
            print(f"DEBUG: File already loaded ({self.file_path}) - opening test start menu")
            self.file_manager.show_test_start_menu(self.file_path)

    def show_data_collection_startup_dialog(self):
        """Show a startup dialog for data collection when no file is loaded."""
        print("DEBUG: Creating data collection startup dialog")
        
        startup_dialog = tk.Toplevel(self.root)
        startup_dialog.title("Data Collection")
        startup_dialog.geometry("350x200")
        startup_dialog.transient(self.root)
        startup_dialog.grab_set()
        startup_dialog.configure(bg=APP_BACKGROUND_COLOR)

        # Header label
        header_label = ttk.Label(
            startup_dialog, 
            text="Start Data Collection", 
            font=("Arial", 16, "bold"),
            foreground="white",
            background=APP_BACKGROUND_COLOR
        )
        header_label.pack(pady=(20, 10))

        # Instruction label
        instruction_label = ttk.Label(
            startup_dialog,
            text="Choose an option to begin collecting data:",
            font=FONT,
            foreground="white",
            background=APP_BACKGROUND_COLOR
        )
        instruction_label.pack(pady=(0, 20))

        # Button frame
        button_frame = ttk.Frame(startup_dialog)
        button_frame.pack(pady=10)

        # New File Button
        new_button = ttk.Button(
            button_frame,
            text="Create New File",
            command=lambda: self.handle_data_collection_new(startup_dialog),
            width=15
        )
        new_button.pack(side="left", padx=10)

        # Load File Button  
        load_button = ttk.Button(
            button_frame,
            text="Load Existing File",
            command=lambda: self.handle_data_collection_load(startup_dialog),
            width=15
        )
        load_button.pack(side="left", padx=10)

        # Cancel button
        cancel_button = ttk.Button(
            button_frame,
            text="Cancel",
            command=startup_dialog.destroy,
            width=10
        )
        cancel_button.pack(pady=(20, 0))

        self.center_window(startup_dialog)
        print("DEBUG: Data collection startup dialog created and displayed")

    def handle_data_collection_new(self, dialog):
        """Handle 'Create New File' from data collection dialog."""
        print("DEBUG: Create New File selected from data collection dialog")
        dialog.destroy()
        
        # Use the existing new file creation logic
        file_path = self.file_manager.create_new_file_with_tests()
        if file_path:
            print(f"DEBUG: New file created at {file_path} - will proceed to test start menu")
            # The create_new_file_with_tests method already handles showing the test start menu
        else:
            print("DEBUG: New file creation was cancelled")

    def handle_data_collection_load(self, dialog):
        """Handle 'Load Existing File' from data collection dialog."""
        print("DEBUG: Load Existing File selected from data collection dialog")
        dialog.destroy()
        
        # Show file selection dialog
        file_path = self.file_manager.ask_open_file()
        if file_path:
            print(f"DEBUG: File selected for loading: {file_path}")
            try:
                # Load the file
                self.file_manager.load_excel_file(file_path)
                
                # Update the UI state
                self.all_filtered_sheets.append({
                    "file_name": os.path.basename(file_path),
                    "file_path": file_path,
                    "filtered_sheets": copy.deepcopy(self.filtered_sheets)
                })
                
                self.file_manager.update_file_dropdown()
                self.file_manager.set_active_file(os.path.basename(file_path))
                self.file_manager.update_ui_for_current_file()
                
                print(f"DEBUG: File loaded successfully - opening test start menu")
                # Show the test start menu for the loaded file
                self.file_manager.show_test_start_menu(file_path)
                
            except Exception as e:
                print(f"DEBUG: Error loading file: {e}")
                messagebox.showerror("Error", f"Failed to load file: {e}")
        else:
            print("DEBUG: No file selected for loading")

    def display_plot(self, full_sample_data):
        """
        Display the plot in the plot frame based on the current data.
        """
        if not hasattr(self, 'plot_frame') or self.plot_frame is None:
            self.plot_frame = ttk.Frame(self.display_frame)
            self.plot_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)

        # Clear existing plot contents
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        # Determine the number of columns per sample
        num_columns = getattr(self, 'num_columns_per_sample', 12)
    
        print(f"DEBUG: Displaying plot with {num_columns} columns per sample")

        # Generate and embed the plot - force it to respect container width
        self.plot_manager.plot_all_samples(self.plot_frame, full_sample_data, num_columns)

        # Ensure plot stays within its frame
        self.plot_frame.grid_propagate(True)  # Allow the frame to resize based on content
        self.plot_manager.add_plot_dropdown(self.plot_frame)

        # Force immediate update of the UI
        self.plot_frame.update_idletasks()

    def display_table(self, frame, data, sheet_name, is_plotting_sheet=False):
            """Display table with enhanced handling for empty/minimal data."""
            print(f"DEBUG: display_table called for sheet: {sheet_name}, data_shape: {data.shape}, is_plotting_sheet: {is_plotting_sheet}")

            if not frame or not frame.winfo_exists():
                print(f"ERROR: Frame {frame} does not exist! Aborting display_table.")
                return

            for widget in frame.winfo_children():
                widget.destroy()

            print("DEBUG: Cleared existing table widgets.")

            # Handle completely empty data
            if data.empty:
                print("DEBUG: Data is completely empty, creating placeholder")
                # Create a placeholder message instead of showing a warning
                placeholder_label = tk.Label(
                    frame,
                    text=f"Sheet '{sheet_name}' is ready for data collection.\nUse the data collection tools to add measurements.",
                    font=("Arial", 12),
                    fg="gray",
                    justify="center"
                )
                placeholder_label.pack(expand=True, pady=50)
                return

            # Deduplicate columns and clean data
            data = clean_columns(data)
            data.columns = data.columns.map(str)  # Ensure column headers are strings
        
            # Handle data with "No data" entries
            data = data.astype(str)
            data = data.replace([np.nan, pd.NA], '', regex=True)

            print("DEBUG: Data processed successfully.")

            table_frame = ttk.Frame(frame, padding=(2, 1))
            self.table_frame.pack_propagate(False)
            table_frame.pack(fill='both', expand=True)

            print("DEBUG: Table frame created and packed.")

            # Calculate table dimensions
            table_frame.update_idletasks()
            available_width = table_frame.winfo_width()
            num_columns = len(data.columns)
            calculated_cellwidth = max(120, available_width // num_columns)

            # Create scrollbars
            v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical')
            h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal')

            # Create table model
            model = TableModel()
            table_data_dict = data.to_dict(orient='index')
            model.importDict(table_data_dict)

            default_cellwidth = 110
            default_rowheight = 25

            if is_plotting_sheet:
                print("DEBUG: Adjusting row height for plotting sheet.")
                font_height = 16
                char_per_line = 12

                # Calculate optimal row height
                row_heights = []
                for _, row in data.iterrows():
                    max_lines = 1
                    for cell in row:
                        if cell:
                            cell_length = len(str(cell))
                            lines = (cell_length // char_per_line) + 1
                            max_lines = max(max_lines, lines)
                    row_heights.append(max_lines * font_height)

                max_row_height = int(max(row_heights)) if row_heights else default_rowheight
                default_rowheight = max_row_height

            print(f"DEBUG: Final row height: {default_rowheight}")

            # Create table canvas
            table_canvas = TableCanvas(
                table_frame, 
                model=model, 
                cellwidth=calculated_cellwidth, 
                cellbackgr='#4CC9F0',
                thefont=('Arial', 10), 
                rowheight=default_rowheight, 
                rowselectedcolor='#FFFFFF',
                editable=False, 
                yscrollcommand=v_scrollbar.set, 
                xscrollcommand=h_scrollbar.set, 
                showGrid=True
            )

            table_canvas.grid(row=0, column=0, sticky='nsew')
            v_scrollbar.grid(row=0, column=1, sticky='ns')
            h_scrollbar.grid(row=1, column=0, sticky='ew')

            table_canvas.show()

            print("DEBUG: Table displayed successfully.")

    def update_plot_dropdown(self):
        """Update the plot type dropdown with only the valid options and manage visibility."""
        current_sheet = self.selected_sheet.get()
        # If available, use self.full_sample_data; otherwise, fall back to the raw data
        if current_sheet not in self.filtered_sheets:
            self.plot_dropdown.pack_forget()
            return

        sheet_data = self.filtered_sheets[current_sheet]["data"]
        valid_plot_options = get_valid_plot_options(self.plot_options, sheet_data)

        if valid_plot_options:
            # Update the plot type dropdown with available plot options
            self.plot_dropdown['values'] = valid_plot_options

            # Check if the selected plot type is valid, set to the first valid option if not
            if self.selected_plot_type.get() not in valid_plot_options:
                self.selected_plot_type.set(valid_plot_options[0])

            # Ensure the plot dropdown is visible and bind its event
            self.plot_dropdown.pack(fill="x", pady=(5, 5))
            self.plot_dropdown.bind(
                "<<ComboboxSelected>>",
                lambda event: self.plot_manager.plot_all_samples(self.display_frame, self.get_full_sample_data(), 12)
            )
        else:
            # Hide the plot dropdown if no valid options are available
            self.plot_dropdown.pack_forget()

    def open_trend_analysis_window(self) -> None:
        trend_gui = TrendAnalysisGUI(self.root, self.filtered_sheets, self.plot_options)
        trend_gui.show()

    def restore_all(self):
        """Restore all lines or bars to their original state."""
        if self.check_buttons:
            self.check_buttons.disconnect(self.checkbox_cid)

        if self.selected_plot_type.get() == "TPM (Bar)":
            # Restore bars
            for patch, original_data in zip(self.axes.patches, self.original_lines_data):
                patch.set_x(original_data[0])
                patch.set_height(original_data[1])
                patch.set_visible(True)
        else:
            # Restore lines
            for line, original_data in zip(self.lines, self.original_lines_data):
                line.set_xdata(original_data[0])
                line.set_ydata(original_data[1])
                line.set_visible(True)

        # Reset all checkboxes to the default "checked" state (forcefully activating each one)
        if self.check_buttons:
            for i in range(len(self.line_labels)):
                if not self.check_buttons.get_status()[i]:  # Only activate if not already active
                    self.check_buttons.set_active(i)  # Reset each checkbox to "checked"

            # Reconnect the callback for CheckButtons after restoring
            self.checkbox_cid = self.check_buttons.on_clicked(
                self.plot_manager.on_bar_checkbox_click if self.selected_plot_type.get() == "TPM (Bar)" else self.plot_manager.on_checkbox_click
            )

        # Redraw the canvas to reflect the restored state
        if self.canvas:
            self.canvas.draw_idle()

    def generate_full_report(self):
        self.progress_dialog.show_progress_bar("Generating full report...")
        success = False

        try:
            # Force GUI to update progress dialog IMMEDIATELY
            self.root.update_idletasks()  

            # Add detailed logging
            import logging
            logging.basicConfig(filename='report_generation.log', level=logging.DEBUG)
            logging.info("Starting full report generation")

             # Add this before generating the report
            logging.info(f"Sheets: {list(self.filtered_sheets.keys())}")
            for sheet_name, sheet_info in self.filtered_sheets.items():
                logging.info(f"Sheet {sheet_name} - empty: {sheet_info.get('is_empty', 'unknown')}")
                if 'data' in sheet_info:
                    logging.info(f"  Data shape: {sheet_info['data'].shape}")

            # Generate report (blocks UI, but ensures dialog stays visible)
            self.report_generator.generate_full_report(
                self.filtered_sheets, 
                self.plot_options
            )
            success = True

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            logging.error(f"Report generation error: {str(e)}\n{error_details}")
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            # Hide progress AFTER processing
            self.progress_dialog.hide_progress_bar()  

            # Show success message AFTER hiding progress
            if success:  
                messagebox.showinfo("Success", "Full report saved successfully.")

            # Force GUI to process pending events
            self.root.update_idletasks() 

    def generate_test_report(self):
        selected_sheet = self.selected_sheet.get()
        # Show the progress dialog
        self.progress_dialog.show_progress_bar("Generating test report...")
        try:
            # Pass the currently selected sheet, the full sheets dictionary, and the plot options.
            self.report_generator.generate_test_report(selected_sheet, self.filtered_sheets, self.plot_options)
        finally:
            self.progress_dialog.hide_progress_bar()

    def open_viscosity_calculator(self):
        """Open the viscosity calculator as a standalone window with a proper menubar."""
        # Create a new top-level window
        calculator_window = tk.Toplevel(self.root)
        calculator_window.title("Viscosity Calculator")
        calculator_window.geometry("550x500")
        calculator_window.resizable(True, True)
        calculator_window.configure(bg=APP_BACKGROUND_COLOR)
    
        # Make the window modal
        calculator_window.transient(self.root)
        calculator_window.grab_set()
    
        # Center the window
        self.center_window(calculator_window)
    
        # Create a menubar for the calculator window
        menubar = tk.Menu(calculator_window)
        calculator_window.config(menu=menubar)
    
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Upload Viscosity Data", command = self.viscosity_calculator.upload_training_data)
        menubar.add_cascade(label="File", menu=file_menu)

        # Create Calculate menu with all the viscosity-related functions
        calculate_menu = tk.Menu(menubar, tearoff=0)

        calculate_menu.add_command(label="Train Standard Models", 
                                  command=self.train_models_from_data)
        calculate_menu.add_command(label="Train Enhanced Models with Potency", 
                                  command=self.train_models_with_chemistry)
        calculate_menu.add_command(label="Analyze Models", 
                                  command=self.viscosity_calculator.analyze_models)
        calculate_menu.add_command(label="Arrhenius Analysis", 
                                  command=self.viscosity_calculator.filter_and_analyze_specific_combinations)
          
        
        menubar.add_cascade(label="Model", menu=calculate_menu)
    
        # Create main container
        main_container = ttk.Frame(calculator_window)
        main_container.pack(fill='both', expand=True, padx=8, pady=8)
    
        # Create the calculator frame and embed the calculator
        calculator_frame = ttk.Frame(main_container)
        calculator_frame.pack(fill='both', expand=True)
    
        # Embed the calculator with all its tabs
        self.viscosity_calculator.embed_in_frame(calculator_frame)

        self.center_window(calculator_window, 550, 500)

        return calculator_window

    def return_to_previous_view(self):
        """Return to the previous view from the calculator."""
        # Restore the bottom frame if it was hidden
        if hasattr(self, 'bottom_frame') and not self.bottom_frame.winfo_ismapped():
            # Explicitly restore ALL original configuration values
            self.bottom_frame.configure(height=150)  # Reset the explicit height
            self.bottom_frame.pack(side="bottom", fill="x", padx=10, pady=(0,10))
            self.bottom_frame.pack_propagate(False)  # Re-apply propagate settings
            self.bottom_frame.grid_propagate(False)
            self.bottom_frame.update_idletasks()  # Force geometry manager to update
    
        # Restore original minimum size constraint
        if hasattr(self, 'previous_minsize'):
            self.root.minsize(*self.previous_minsize)
        else:
            self.root.minsize(1200, 800)  # Default fallback
    
        # Restore the previous window geometry
        if hasattr(self, 'previous_window_geometry'):
            self.root.geometry(self.previous_window_geometry)
            self.center_window(self.root)
    
        # Get the currently selected sheet
        current_sheet = self.selected_sheet.get()
    
        # Return to the standard view with that sheet
        if current_sheet in self.filtered_sheets:
            self.update_displayed_sheet(current_sheet)
        elif self.filtered_sheets:
            # If current sheet is invalid, select the first available sheet
            first_sheet = list(self.filtered_sheets.keys())[0]
            self.selected_sheet.set(first_sheet)
            self.update_displayed_sheet(first_sheet)
        else:
            # If no sheets available, just clear the display
            self.clear_dynamic_frame()



