"""
main_gui.py
Developed By Charlie Becquet.
Main GUI module for the DataViewer Application.
This module instantiates the high-level TestingGUI class which delegates file I/O,
plotting, trend analysis, report generation, and progress dialog management
to separate modules for better modularity and scalability.
"""
import queue
import os
import threading
import copy
import shutil
import tkinter as tk
import pandas as pd
import processing
import numpy as np
from typing import Optional, Dict, List, Any
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel, Label, Button
from tkintertable import TableCanvas, TableModel
import matplotlib
matplotlib.use('TkAgg')  # Ensure Matplotlib uses TkAgg backend
import matplotlib.pyplot as plt

# Import our new manager classes and utility functions.

from plot_manager import PlotManager
from file_manager import FileManager
from report_generator import ReportGenerator
from trend_analysis_gui import TrendAnalysisGUI
from progress_dialog import ProgressDialog
from image_loader import ImageLoader
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
        self.plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        self.check_buttons = None

        self.crop_enable = tk.BooleanVar(value = False)
        
        # Instantiate manager classes
        self.file_manager = FileManager(self)
        self.plot_manager = PlotManager(self)
        self.report_generator = ReportGenerator(self)
        self.progress_dialog = ProgressDialog(self.root)
        # (TrendAnalysisGUI will be created when needed)

    # === New Centralized Frame Creation Methods ===
    def add_static_controls(self) -> None:
        """Add static controls (Add Data and Trend Analysis buttons) to the top_frame."""
        print("DEBUG: Adding static controls to top_frame...")
        # Create a container frame for the static controls if not already created.
        if not hasattr(self, 'controls_frame') or not self.controls_frame:
            self.controls_frame = ttk.Frame(self.top_frame)
            # Pack the controls_frame to the right so it aligns with the file dropdown on the same row.
            self.controls_frame.pack(side="right", fill="x", padx=5, pady=5)
            print("DEBUG: controls_frame created in top_frame.")
        # Add the "Add Data" button
        add_data_btn = ttk.Button(self.controls_frame, text="Add Data", command=self.file_manager.add_data)
        add_data_btn.pack(side="left", padx=(5, 5), pady=(5, 5))
        print("DEBUG: 'Add Data' button added to static controls.")
        # Add the "Trend Analysis" button
        trend_button = ttk.Button(self.controls_frame, text="Trend Analysis", command=self.open_trend_analysis_window)
        trend_button.pack(side="left", padx=(5, 5), pady=(5, 5))
        print("DEBUG: 'Trend Analysis' button added to static controls.")

        self.add_report_buttons(self.controls_frame)



    def create_static_frames(self) -> None:
        """Create persistent (static) frames that remain for the lifetime of the UI."""
        print("DEBUG: Creating static frames...")

        # Create top_frame for dropdowns and control buttons.
        if not hasattr(self, 'top_frame') or not self.top_frame:
            self.top_frame = ttk.Frame(self.root,height = 80)
            self.top_frame.pack(side="top", fill="x", pady=(10, 5), padx=10)
            self.top_frame.pack_propagate(False) # Prevent height changes
            print("DEBUG: top_frame created.")

        # Create bottom_frame to hold the image button and image display area.
        if not hasattr(self, 'bottom_frame') or not self.bottom_frame:
            self.bottom_frame = ttk.Frame(self.root, height = 150)
            self.bottom_frame.pack(side="bottom", fill = "x", padx=10, pady=(0,10))
            self.bottom_frame.pack_propagate(False)
            self.bottom_frame.grid_propagate(False)
            print(f"DEBUG: Created bottom_frame with fixed height 150 | "
              f"Current height: {self.bottom_frame.winfo_height()}")

        # Create a static frame for the Load Images button within bottom_frame.
        if not hasattr(self, 'image_button_frame') or not self.image_button_frame:
            self.image_button_frame = ttk.Frame(self.bottom_frame)
            # Pack it to the left.
            self.image_button_frame.pack(side="left", fill="y", padx=5, pady=5)
            print("DEBUG: image_button_frame created.")

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
            print("DEBUG: Static 'Load Images' button added to image_button_frame.")

        # Create the dynamic image display frame within bottom_frame.
        if not hasattr(self, 'image_frame') or not self.image_frame:
            self.image_frame = ttk.Frame(self.bottom_frame, height = 150)
            self.image_frame.pack(side="left", fill="both", expand=True)
            self.image_frame.pack_propagate(False)
            self.image_frame.grid_propagate(False)
            # Prevent the image_frame from shrinking.
            #self.image_frame.pack_propagate(True)
            print("DEBUG: Dynamic image_frame created with fixed height 300.")


        # Create display_frame for the table/plot area.
        if not hasattr(self, 'display_frame') or not self.display_frame:
            self.display_frame = ttk.Frame(self.root)
            self.display_frame.pack(fill="both", expand=True, padx=10, pady=10)
            print("DEBUG: display_frame created.")

        # Create a dynamic subframe inside display_frame for table and plot content.
        if not hasattr(self, 'dynamic_frame') or not self.dynamic_frame:
            self.dynamic_frame = ttk.Frame(self.display_frame)
            self.dynamic_frame.pack(fill="both", expand=True)
            print("DEBUG: dynamic_frame created.")

        


    def clear_dynamic_frame(self) -> None:
        """Clear all children widgets from the dynamic frame."""
        print("DEBUG: Clearing dynamic_frame contents...")
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()

    def setup_dynamic_frames(self, is_plotting_sheet: bool = False) -> None:
        """Create frames inside the dynamic_frame based on sheet type."""
        print(f"DEBUG: Setting up dynamic frames for sheet. Plotting: {is_plotting_sheet}")

        # Ensure that any previous frames are cleared before adding new ones
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()

        # Get dynamic frame height
        window_height = self.root.winfo_height()
        top_height = self.top_frame.winfo_height()
        bottom_height = self.bottom_frame.winfo_height()
        padding = 20  # Adjust if needed
        display_height = window_height - top_height - bottom_height - padding
        display_height = max(display_height, 100)  # Ensure it never becomes too small

        print(f"DEBUG: [setup_dynamic_frames] Window: {window_height}px | "
          f"Top: {top_height}px | Bottom: {bottom_height}px | "
          f"Available for display: {window_height - top_height - bottom_height}px")

        if is_plotting_sheet:
            # Table takes 40%, Plot takes 60%
            self.table_frame = ttk.Frame(self.dynamic_frame, height=int(display_height * 0.4))
            self.table_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

            self.plot_frame = ttk.Frame(self.dynamic_frame, height=int(display_height * 0.6))
            self.plot_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

            print("DEBUG: Dynamic frames for plotting sheet created (table_frame, plot_frame).")
        else:
            # Non-plotting sheets should use the full space
            self.table_frame = ttk.Frame(self.dynamic_frame, height=display_height)
            self.table_frame.pack(fill="both", expand=True, padx=5, pady=5)
            print("DEBUG: Dynamic frame for non-plotting sheet created (table_frame).")

    # === End of New Centralized Frame Methods ===


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

    def configure_ui(self) -> None:
        """Configure the UI appearance and set application properties."""
        icon_path = get_resource_path('resources/ccell_icon.png')
        self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
        self.set_app_colors()
        self.set_window_size(0.8, 0.8)
        self.root.minsize(1200,800)
        self.center_window(self.root)

        self.create_static_frames()
        self.add_menu()
        self.file_manager.add_or_update_file_dropdown()
        self.add_static_controls()

    def populate_or_update_sheet_dropdown(self) -> None:
        """Populate or update the dropdown for sheet selection."""
        if not hasattr(self, 'drop_down_menu') or not self.drop_down_menu:
            # Create a dropdown for sheet selection
            self.drop_down_menu = ttk.Combobox(
                self.top_frame,
                textvariable=self.selected_sheet,
                state="readonly",
                font=FONT
            )
            self.drop_down_menu.pack(pady=(10, 10))
            self.drop_down_menu.place(relx=0.25, rely=0.5, anchor="center")

            # Bind selection event to update the displayed sheet
            self.drop_down_menu.bind(
                "<<ComboboxSelected>>",
                lambda event: self.update_displayed_sheet(self.selected_sheet.get())
            )
        self.drop_down_menu
        # Update dropdown values with sheet names
        all_sheet_names = list(self.filtered_sheets.keys())
        self.drop_down_menu["values"] = all_sheet_names
        current_selection = self.selected_sheet.get()
        if current_selection in all_sheet_names:
            self.selected_sheet.set(current_selection)
        elif all_sheet_names:
            # if the current selection is invalid, default to the first sheet
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
                    print(f"Stopping thread {thread.name}")
            self.root.destroy()
            os._exit(0)
        except Exception as e:
            print(f"Error during shutdown: {e}")
            os._exit(1)

    def add_menu(self) -> None:
        """Create a top-level menu with Help and About options."""
        menubar = tk.Menu(self.root)
        
        # Help menu
        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Help", command=self.show_help)
        helpmenu.add_separator()
        helpmenu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)
        
        self.root.config(menu=menubar)

    def show_help(self):
        """Display help dialog."""
        messagebox.showinfo("Help", "This program is designed to be used with excel data according to the SDR Standardized Testing Template.\n \nClick 'Generate Test Report' to create an excel report of a single test, or click 'Generate Full Report' to generate both an excel file and powerpoint file of all the contents within the file.")

    def show_about(self) -> None:
        """Display about dialog."""
        messagebox.showinfo("About", "SDR DataView Alpha Version 1.0\nDeveloped by Charlie Becquet")

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
        Clears the plot area if the sheet is empty.
        """
        print(f"DEBUG: [update_displayed_sheet] START | Current frame heights - "
          f"top: {self.top_frame.winfo_height()}, "
          f"display: {self.display_frame.winfo_height()}, "
          f"bottom: {self.bottom_frame.winfo_height()}")

        if not sheet_name or sheet_name not in self.filtered_sheets:
            return

        print(f"DEBUG: Attempting to update displayed sheet: {sheet_name}")
        if not hasattr(self, 'display_frame') or self.display_frame is None:
            print("ERROR: display_frame is missing! This may cause issues.")
        else:
            print("DEBUG: Existing display_frame widgets:", self.display_frame.winfo_children())

        sheet_info = self.filtered_sheets.get(sheet_name)
        if not sheet_info:
            messagebox.showerror("Error", f"Sheet '{sheet_name}' not found.")
            return

        data = sheet_info["data"]
        is_empty = sheet_info["is_empty"]
        is_plotting_sheet = plotting_sheet_test(sheet_name, data)

        # Instead of destroying static frames, clear and rebuild only the dynamic content.
        self.clear_dynamic_frame()
        self.setup_dynamic_frames(is_plotting_sheet)

        print("DEBUG: Creating ImageLoader...")

         

        # Clear existing ImageLoader properly
        if hasattr(self, 'image_loader'):
            for widget in self.image_frame.winfo_children():
                widget.destroy()
            del self.image_loader

        # initialize a fresh one
        self.image_loader = ImageLoader(self.image_frame, is_plotting_sheet, on_images_selected=lambda paths: self.store_images(sheet_name, paths), main_gui = self)

        self.image_loader.frame.config(height=150)

        if sheet_name in self.sheet_images:
            self.image_loader.load_images_from_list(self.sheet_images[sheet_name])
            for img_path in self.sheet_images[sheet_name]:
                if img_path in self.image_crop_states:
                    self.image_loader.image_crop_states[img_path] = self.image_crop_states[img_path]

            print(f"DEBUG: Restored images and crop states for sheet: {sheet_name}")

        # **Force the frame to refresh**
        self.image_loader.display_images()
        self.image_frame.update_idletasks()  # Ensure UI updates immediately


        if is_empty:
            print("DEBUG: Sheet is empty. Displaying message.")
            empty_label = tk.Label(
                self.display_frame,
                text="This sheet is empty.",
                font=("Arial", 14),
                fg="red"
            )
            empty_label.pack(anchor="center", pady=20)
            return

        try:
            print(f"DEBUG: Retrieving processing function for {sheet_name}...")
            process_function = processing.get_processing_function(sheet_name)
            processed_data, _, full_sample_data = process_function(data)
        except Exception as e:
            messagebox.showerror("Processing Error", f"Error processing sheet '{sheet_name}': {e}")
            print(f"ERROR: Processing function failed for {sheet_name}: {e}")
            return

        print(f"DEBUG: Displaying table for {sheet_name}...")
        self.display_table(self.table_frame, processed_data, sheet_name, is_plotting_sheet)
        print(f"DEBUG: Sheet name: {sheet_name}, is_plotting_sheet: {is_plotting_sheet}")

        if is_plotting_sheet:
            print(f"DEBUG: Displaying plot for sheet: {sheet_name}")
            self.display_plot(full_sample_data)

        print("DEBUG: Successfully updated displayed sheet.")
        self.root.update_idletasks()
        
        print(f"DEBUG: [update_displayed_sheet] END | Current frame heights - "
              f"top: {self.top_frame.winfo_height()}, "
              f"display: {self.display_frame.winfo_height()}, "
              f"bottom: {self.bottom_frame.winfo_height()}\n")

    def load_images(self, is_plotting_sheet):
        """Load and display images in the dynamic image_frame."""
        if not hasattr(self, 'image_frame') or not self.image_frame.winfo_exists():
            print("Error: image_frame does not exist, recreating...")
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

        print(f"DEBUG: Storing images for sheet: {sheet_name}")
    
        # Store image paths
        self.sheet_images[sheet_name] = paths

        # Store crop states for each image
        if not hasattr(self, 'image_crop_states'):
            self.image_crop_states = {}

        for img_path in paths:
            # Ensure the image crop state is stored
            if img_path not in self.image_crop_states:
                self.image_crop_states[img_path] = self.crop_enabled.get()

        print(f"DEBUG: Stored image paths: {self.sheet_images[sheet_name]}")
        print(f"DEBUG: Stored crop states: {self.image_crop_states}")


    def display_plot(self, full_sample_data):
        """
        Display the plot in the plot frame based on the current data.
        """
        if not hasattr(self, 'plot_frame') or self.plot_frame is None:
            self.plot_frame = ttk.Frame(self.display_frame)
            self.plot_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)

        # Generate and embed the plot
        self.plot_manager.plot_all_samples(self.plot_frame, full_sample_data, 12)

        self.plot_manager.add_plot_dropdown(self.plot_frame)

    def display_table(self, frame, data, sheet_name, is_plotting_sheet=False):
        #print(f"DEBUG: Displaying table for sheet: {sheet_name} (is_plotting_sheet={is_plotting_sheet})")
        #print(f"DEBUG: Checking if frame exists before using it... Frame={frame}")

        if not frame or not frame.winfo_exists():
            print(f"ERROR: Frame {frame} does not exist! Aborting display_table.")
            return

        for widget in frame.winfo_children():
            widget.destroy()

        #print("DEBUG: Cleared existing table widgets.")

        # Deduplicate columns
        data = clean_columns(data)
        data.columns = data.columns.map(str)  # Ensure column headers are strings

        if data.empty:
            print("DEBUG: Data is empty. Displaying warning.")
            messagebox.showwarning("Warning", f"Sheet '{sheet_name}' contains no data to display.")
            return

        data = data.astype(str)
        data = data.replace([np.nan, pd.NA], '', regex=True)

        #print("DEBUG: Data processed successfully.")

        table_frame = ttk.Frame(frame, padding=(2, 1))
        self.table_frame.pack_propagate(False)
        table_frame.pack(fill='both', expand=True)

        #print("DEBUG: Table frame created and packed.")

        # Force the frame to update its geometry
        table_frame.update_idletasks()
        available_width = table_frame.winfo_width()
        num_columns = len(data.columns)

        # Calculate a new cell width based on available width (if table would be too narrow, use the default)
        calculated_cellwidth = max(120, available_width // num_columns)

        v_scrollbar = ttk.Scrollbar(table_frame, orient='vertical')
        h_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal')

        model = TableModel()
        table_data_dict = data.to_dict(orient='index')
        model.importDict(table_data_dict)

        default_cellwidth = 120
        default_rowheight = 30

        if is_plotting_sheet:
            #print("DEBUG: Adjusting row height for plotting sheet.")
            font_height = 20  # Taller rows to accommodate larger content
            char_per_line = 12  # Fewer characters per line for wider spacing

            # Calculate optimal row height
            row_heights = []
            for _, row in data.iterrows():
                max_lines = 1  # Default minimum height for single-line cells
                for cell in row:
                    if cell:
                        # Calculate lines needed for this cell
                        cell_length = len(cell)
                        lines = (cell_length // char_per_line) + 1
                        max_lines = max(max_lines, lines)
                row_heights.append(max_lines * font_height)

            # Use the maximum row height across all rows
            max_row_height = int(max(row_heights))
            default_rowheight = max_row_height

        #print(f"DEBUG: Final row height: {default_rowheight}")

        table_canvas = TableCanvas(table_frame, model=model, cellwidth=calculated_cellwidth, cellbackgr='#4CC9F0',
                                   thefont=('Arial', 12), rowheight=default_rowheight, rowselectedcolor='#FFFFFF',
                                   editable=False, yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set, showGrid=True)

        table_canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        table_canvas.show()

        #print("DEBUG: Table displayed successfully.")

        # Add the "View Raw Data" button below the table
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill='x', pady=(10,10))

        #print("DEBUG: Button frame created.")

        view_raw_button = ttk.Button(button_frame, text="View Raw Data", command=lambda: self.file_manager.open_raw_data_in_excel(sheet_name))
        view_raw_button.pack(anchor='center')

        #print("DEBUG: View Raw Data button added.")

        #print("DEBUG: Load Images button added.")

    def update_plot_dropdown(self):
        """Update the plot type dropdown with only the valid options and manage visibility."""
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

            # Generate report (blocks UI, but ensures dialog stays visible)
            self.report_generator.generate_full_report(
                self.filtered_sheets, 
                self.plot_options
            )
            success = True

        except Exception as e:
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

