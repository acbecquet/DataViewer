"""
main_gui.py
Optimized version for better startup time and UI responsiveness.
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
import time
import tkinter as tk
from typing import Optional, Dict, List, Any
from tkinter import ttk, messagebox, Toplevel
from utils import debug_print, wrap_text

# Lazy loading helper functions
def lazy_import_pandas():
    """Lazy import pandas when needed."""
    try:
        import pandas as pd
        return pd
    except ImportError as e:
        debug_print(f"Error importing pandas: {e}")
        return None

def lazy_import_numpy():
    """Lazy import numpy when needed."""
    try:
        import numpy as np
        return np
    except ImportError as e:
        debug_print(f"Error importing numpy: {e}")
        return None

def lazy_import_matplotlib():
    """Lazy import matplotlib when needed."""
    try:
        import matplotlib
        matplotlib.use('TkAgg')
        return matplotlib
    except ImportError as e:
        debug_print(f"Error importing matplotlib: {e}")
        return None

def lazy_import_tkintertable():
    """Lazy import tkintertable when needed."""
    try:
        from tkintertable import TableCanvas, TableModel
        return TableCanvas, TableModel
    except ImportError as e:
        debug_print(f"Error importing tkintertable: {e}")
        return None, None

def lazy_import_tksheet():
    """Lazy import tksheet when needed."""
    try:
        from tksheet import Sheet
        return Sheet
    except ImportError as e:
        debug_print(f"Error importing tksheet: {e}")
        return None



def lazy_import_requests():
    """Lazy import requests when needed."""
    try:
        import requests
        return requests
    except ImportError as e:
        debug_print(f"Error importing requests: {e}")
        return None

def lazy_import_packaging():
    """Lazy import packaging when needed."""
    try:
        from packaging import version
        return version
    except ImportError as e:
        debug_print(f"Error importing packaging: {e}")
        return None

def _lazy_import_processing():
    """Lazy import processing module."""
    try:
        import processing
        from processing import get_valid_plot_options
        return processing, get_valid_plot_options
    except ImportError as e:
        debug_print(f"Error importing processing: {e}")
        return None, None

def lazy_import_viscosity_gui():
    """Lazy import viscosity GUI when needed."""
    try:
        from viscosity_gui import ViscosityGUI
        return ViscosityGUI
    except ImportError as e:
        debug_print(f"Error importing viscosity GUI: {e}")
        return None

# Import utilities after lazy imports
from utils import FONT, clean_columns, get_save_path, is_standard_file, plotting_sheet_test, APP_BACKGROUND_COLOR, BUTTON_COLOR, PLOT_CHECKBOX_TITLE, clean_display_suffixes, show_success_message
from resource_utils import get_resource_path
from update_checker import UpdateChecker

def is_empty_sample(sample_data):
    """
    Check if a sample is empty based only on plotting data:
    - No TPM data
    - No Average TPM
    - No Draw Pressure  
    - No Resistance
    
    Args:
        sample_data (dict or pd.Series): Sample data to check
        
    Returns:
        bool: True if sample is empty, False otherwise
    """
    try:
        # Convert to dict if it's a pandas Series
        if hasattr(sample_data, 'to_dict'):
            sample_dict = sample_data.to_dict()
        else:
            sample_dict = sample_data
            
        # Only check these plotting fields
        plotting_fields = ['Average TPM', 'Draw Pressure', 'Resistance']
        
        for field in plotting_fields:
            value = str(sample_dict.get(field, '')).strip()
            debug_print(f"DEBUG: Checking field '{field}': '{value}'")
            
            # Skip completely empty values
            if not value or value in ['', 'nan', 'No data', 'None']:
                continue
                
            # Check if it's a meaningful numeric value
            try:
                numeric_val = float(value)
                # Remove pd.isna() call and use math.isnan() or simple check
                import math
                if numeric_val != 0 and not math.isnan(numeric_val):
                    debug_print(f"DEBUG: Sample has plotting data in '{field}': {numeric_val}")
                    return False  # Has data, not empty
            except (ValueError, TypeError):
                # If it's not numeric but has meaningful content, it's not empty
                if len(value) > 0 and value not in ['nan', 'None', 'No data', '']:
                    debug_print(f"DEBUG: Sample has non-numeric plotting data in '{field}': '{value}'")
                    return False
        
        
        debug_print("DEBUG: Sample has no plotting data - is empty")
        return True  # No plotting data found
        
    except Exception as e:
        debug_print(f"DEBUG: Error checking sample: {e}")
        return False  # If error, assume not empty

def filter_empty_samples_from_dataframe(df):
    """
    Filter out samples with no plotting data from processed DataFrame.
    """
    if df.empty:
        debug_print("DEBUG: DataFrame is already empty")
        return df
        
    try:
        debug_print(f"DEBUG: Checking {len(df)} samples for plotting data")
        
        # Create mask for samples with plotting data
        has_data_mask = []
        
        for index, row in df.iterrows():
            is_empty = is_empty_sample(row)
            has_data_mask.append(not is_empty)
            debug_print(f"DEBUG: Sample {index} ({'empty' if is_empty else 'has data'}): {row.get('Sample Name', 'Unknown')}")
            
        # Filter the dataframe
        filtered_df = df[has_data_mask].reset_index(drop=True)
        
        debug_print(f"DEBUG: Filtered from {len(df)} to {len(filtered_df)} samples with plotting data")
        return filtered_df
        
    except Exception as e:
        debug_print(f"DEBUG: Error filtering samples: {e}")
        import traceback
        traceback.print_exc()
        return df

def filter_empty_samples_from_full_data(full_sample_data, num_columns_per_sample=12):
    """
    Filter out samples with no TPM plotting data from full sample data.
    """
    pd = lazy_import_pandas()
    if not pd or full_sample_data.empty:
        debug_print("DEBUG: No pandas or empty full_sample_data")
        return full_sample_data
        
    try:
        debug_print(f"DEBUG: Full data shape: {full_sample_data.shape}, columns per sample: {num_columns_per_sample}")
        
        # For User Test Simulation (8 columns), we need to match the processed data filtering
        if num_columns_per_sample == 8:
            debug_print("DEBUG: User Test Simulation detected - preserving all data since processed data was already filtered")
            # The processed data filtering already removed empty samples
            # So we keep the full data as-is for User Test Simulation
            return full_sample_data
        
        # For regular tests (12 columns), do the normal filtering
        num_samples = full_sample_data.shape[1] // num_columns_per_sample
        debug_print(f"DEBUG: Checking {num_samples} samples in full plotting data")
        
        if num_samples == 0:
            return full_sample_data
            
        # Find samples with actual TPM data
        samples_with_data = []
        
        for i in range(num_samples):
            start_col = i * num_columns_per_sample
            end_col = start_col + num_columns_per_sample
            sample_data = full_sample_data.iloc[:, start_col:end_col]
            
            debug_print(f"DEBUG: Checking sample {i+1} columns {start_col}-{end_col-1}")
            
            # Check TPM column (usually column 8 in 12-column format)
            has_tpm_data = False
            if num_columns_per_sample >= 9:
                tpm_col_idx = 8 if num_columns_per_sample == 12 else min(8, num_columns_per_sample - 1)
                if sample_data.shape[1] > tpm_col_idx and sample_data.shape[0] > 3:
                    # Get TPM data from row 3 onwards
                    tpm_data = sample_data.iloc[3:, tpm_col_idx]
                    
                    # Convert to numeric and check for real values
                    numeric_tpm = pd.to_numeric(tpm_data, errors='coerce')
                    valid_tpm = numeric_tpm.dropna()
                    
                    if len(valid_tpm) > 0 and (valid_tpm > 0).any():
                        has_tpm_data = True
                        debug_print(f"DEBUG: Sample {i+1} has TPM data: {valid_tpm.head().tolist()}")
            
            if has_tpm_data:
                samples_with_data.append(i)
        
        debug_print(f"DEBUG: Found {len(samples_with_data)} samples with TPM data out of {num_samples}")
        
        # If no samples have data, return empty DataFrame
        if not samples_with_data:
            debug_print("DEBUG: No samples with plotting data, returning empty DataFrame")
            return pd.DataFrame()
            
        # Reconstruct with only samples that have data
        filtered_columns = []
        for sample_idx in samples_with_data:
            start_col = sample_idx * num_columns_per_sample
            end_col = start_col + num_columns_per_sample
            sample_cols = list(range(start_col, min(end_col, full_sample_data.shape[1])))
            filtered_columns.extend(sample_cols)
            debug_print(f"DEBUG: Including sample {sample_idx+1} columns {start_col}-{end_col-1}")
            
        filtered_data = full_sample_data.iloc[:, filtered_columns]
        debug_print(f"DEBUG: Filtered plotting data from {full_sample_data.shape[1]} to {filtered_data.shape[1]} columns")
        return filtered_data
            
    except Exception as e:
        debug_print(f"DEBUG: Error filtering plotting data: {e}")
        import traceback
        traceback.print_exc()
        return full_sample_data

class TestingGUI:
    """Main GUI class for the Standardized Testing application."""

    def __init__(self, root):
        import main 
        
        init_start = time.time()
        
        self.root = root
        self.root.title("DataViewer")
        self.report_thread = None
        self.report_queue = queue.Queue()

        # Initialize variables first
        self.initialize_variables()

        # Configure UI
        self.configure_ui()

        # Show startup menu
        self.show_startup_menu()

        self.image_loader = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_app_close)

        # Setup update checking - disable when no internet!
        #self.update_checker = UpdateChecker(current_version="3.0.0")
        #self.setup_update_checking()

        init_time = time.time() - init_start

    def setup_update_checking(self):
        """Set up automatic update checking"""
        def check_updates():
            try:
                update_info = self.update_checker.check_for_updates()
                
                if update_info and update_info.get('update_available'):
                    self.show_update_dialog(update_info)
            except Exception:
                pass  # Silently fail for updates
        
        self.root.after(5000, check_updates)
    
    def check_for_updates_background(self):
        """Check for updates without blocking UI"""
        try:
            update_info = self.update_checker.check_for_updates()
            
            if update_info and update_info.get('update_available'):
                self.show_update_dialog(update_info)
                
        except Exception:
            pass
    
    def show_update_dialog(self, update_info):
        """Show update notification to user"""
        from tkinter import messagebox
        
        latest_version = update_info['latest_version']
        current_version = update_info['current_version']
        
        message = f"""Update Available!

Current Version: {current_version}
Latest Version: {latest_version}

Would you like to download and install the update?"""
        
        result = messagebox.askyesno("Update Available", message)
        
        if result:
            self.download_and_install_update(update_info)
    
    def download_and_install_update(self, update_info):
        """Download and install the update"""
        if not update_info.get('download_url'):
            messagebox.showerror("Error", "No download URL available")
            return
        
        try:
            show_success_message("Downloading", "Downloading update... This may take a moment.", self.root)
            
            installer_path = self.update_checker.download_update(
                update_info['download_url'],
                update_info['installer_name']
            )
            
            if installer_path:
                self.update_checker.install_update(installer_path)
            else:
                messagebox.showerror("Error", "Failed to download update")
                
        except Exception as e:
            messagebox.showerror("Error", f"Update failed: {e}")

    def get_viscosity_calculator(self):
        """Lazy load viscosity calculator only when needed."""
        if self.viscosity_calculator is None:
            from viscosity_calculator import ViscosityCalculator
            self.viscosity_calculator = ViscosityCalculator(self)
        return self.viscosity_calculator

    # Lazy loading methods for the class
    def get_pandas(self):
        """Get pandas with lazy loading."""
        if not hasattr(self, '_pandas'):
            self._pandas = lazy_import_pandas()
        return self._pandas

    def get_numpy(self):
        """Get numpy with lazy loading.""" 
        if not hasattr(self, '_numpy'):
            self._numpy = lazy_import_numpy()
        return self._numpy

    def get_matplotlib(self):
        """Get matplotlib with lazy loading."""
        if not hasattr(self, '_matplotlib'):
            self._matplotlib = lazy_import_matplotlib()
        return self._matplotlib

    def get_tkintertable(self):
        """Get tkintertable with lazy loading."""
        if not hasattr(self, '_tkintertable'):
            self._tkintertable = lazy_import_tkintertable()
        return self._tkintertable

    # Update the get_tkintertable method to get_tksheet
    def get_tksheet(self):
        """Get tksheet with lazy loading."""
        if not hasattr(self, '_tksheet'):
            self._tksheet = lazy_import_tksheet()
        return self._tksheet

    def lazy_init_managers(self):
        """Initialize manager classes only when needed."""
        if hasattr(self, '_managers_initialized') and self._managers_initialized:
            return
            
        # Import and initialize managers
        from plot_manager import PlotManager
        from file_manager import FileManager
        from report_generator import ReportGenerator
        from progress_dialog import ProgressDialog
        
        self.file_manager = FileManager(self)
        self.plot_manager = PlotManager(self)
        self.report_generator = ReportGenerator(self)
        self.progress_dialog = ProgressDialog(self.root)
        
        self._managers_initialized = True

    # Initialization and Configuration
    def initialize_variables(self) -> None:
        """Initialize variables used throughout the GUI."""
        pd = self.get_pandas()
        if not pd:
            debug_print("ERROR: Could not load pandas")
            return None
            
        # Basic variables
        self.sheets: Dict[str, pd.DataFrame] = {}
        self.filtered_sheets: Dict[str, pd.DataFrame] = {}
        self.all_filtered_sheets: List[Dict[str, Any]] = []
        self.current_file: Optional[str] = None
        self.selected_sheet = tk.StringVar()
        self.selected_plot_type = tk.StringVar(value="TPM")
        self.file_path: Optional[str] = None
        self.plot_dropdown = None
        self.threads = []
        self.canvas = None
        self.figure = None
        self.axes = None
        self.lines = None
        self.sheet_images = {}
        self.lock = threading.Lock()
        self.line_labels = []
        self.original_lines_data = []
        self.checkbox_cid = None
        

        # Plot options
        self.standard_plot_options = ["TPM", "Draw Pressure", "Resistance", "Power Efficiency", "TPM (Bar)"]
        self.user_test_simulation_plot_options = ["TPM", "Draw Pressure", "Power Efficiency", "TPM (Bar)"]
        self.plot_options = self.standard_plot_options

        self.check_buttons = None
        self.previous_window_geometry = None
        self.is_user_test_simulation = False  
        self.num_columns_per_sample = 12 

        self.crop_enable = tk.BooleanVar(value=False)
        
        # Initialize managers - create basic ones immediately for core functionality
        self.lazy_init_managers()

    def configure_ui(self) -> None:
        """Configure the UI appearance and set application properties."""
        # Icon loading
        icon_path = get_resource_path('resources/ccell_icon.png')
        self.root.iconphoto(False, tk.PhotoImage(file=icon_path))
    
        # Basic UI setup
        self.set_app_colors()
        self.root.state('zoomed')
        self.set_window_size(1, 1)
        self.root.minsize(1200,800)
    
        # Create frames
        self.create_static_frames()
    
        # Create menu immediately - needed for basic functionality
        self.add_menu()
        
        # Setup file dropdown
        self.file_manager.add_or_update_file_dropdown()
        
        # Add static controls
        self.add_static_controls()

    def add_static_controls(self) -> None:
        """Add static controls (Add Data and Trend Analysis buttons) to the top_frame."""
        if not hasattr(self, 'controls_frame') or not self.controls_frame:
            self.controls_frame = ttk.Frame(self.top_frame)
            self.controls_frame.pack(side="right", fill="x", padx=5, pady=5)

    def create_static_frames(self) -> None:
        """Create persistent (static) frames that remain for the lifetime of the UI."""
        # Create top_frame for dropdowns and control buttons.
        if not hasattr(self, 'top_frame') or not self.top_frame:
            self.top_frame = ttk.Frame(self.root,height = 30)
            self.top_frame.pack(side="top", fill="x", pady=(10, 0), padx=10)
            self.top_frame.pack_propagate(False)

        # Create bottom_frame to hold the image button and image display area.
        if not hasattr(self, 'bottom_frame') or not self.bottom_frame:
            self.bottom_frame = ttk.Frame(self.root, height = 150)
            self.bottom_frame.pack(side="bottom", fill = "x", padx=10, pady=(0,10))
            self.bottom_frame.pack_propagate(False)
            self.bottom_frame.grid_propagate(False)

        # Create a static frame for the Load Images button within bottom_frame.
        if not hasattr(self, 'image_button_frame') or not self.image_button_frame:
            self.image_button_frame = ttk.Frame(self.bottom_frame)
            self.image_button_frame.pack(side="left", fill="y", padx=5, pady=5)

            self.crop_enabled = tk.BooleanVar(value=False)
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

        # Create the dynamic image display frame within bottom_frame.
        if not hasattr(self, 'image_frame') or not self.image_frame:
            self.image_frame = ttk.Frame(self.bottom_frame, height = 150)
            self.image_frame.pack(side="left", fill="both", expand=True)
            self.image_frame.pack_propagate(False)
            self.image_frame.grid_propagate(False)

        # Create display_frame for the table/plot area.
        if not hasattr(self, 'display_frame') or not self.display_frame:
            self.display_frame = ttk.Frame(self.root)
            self.display_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Create a dynamic subframe inside display_frame for table and plot content.
        if not hasattr(self, 'dynamic_frame') or not self.dynamic_frame:
            self.dynamic_frame = ttk.Frame(self.display_frame)
            self.dynamic_frame.pack(fill="both", expand=True)

    def clear_dynamic_frame(self) -> None:
        """Clear all children widgets from the dynamic frame."""
        for widget in self.dynamic_frame.winfo_children():
            widget.destroy()

    def on_window_resize(self, event):
        """Handle window resize events to maintain layout proportions."""
        if event.widget == self.root:
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
            self.dynamic_frame.columnconfigure(0, weight=5)
            self.dynamic_frame.columnconfigure(1, weight=5)
        
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
        """Ensure the plot doesn't exceed 50% of the window width."""
        if not hasattr(self, 'plot_frame') or not self.plot_frame:
            return
        
        window_width = self.root.winfo_width()
        max_plot_width = window_width // 2
    
        if max_plot_width > 50:
            self.plot_frame.config(width=max_plot_width)
            self.plot_frame.grid_propagate(False)
        
            if hasattr(self, 'table_frame') and self.table_frame:
                self.table_frame.config(width=max_plot_width)
                self.table_frame.grid_propagate(False)

    def show_startup_menu(self) -> None:
        """Display a startup menu with 'New' and 'Load' options."""
        startup_menu = Toplevel(self.root)
        startup_menu.title("Welcome")
        startup_menu.geometry("400x150")
        startup_menu.transient(self.root)
        startup_menu.grab_set()

        label = ttk.Label(startup_menu, text="Welcome to DataViewer by SDR!", font=FONT,background=APP_BACKGROUND_COLOR)
        label.pack(pady=10)

        # New Button
        new_button = ttk.Button(
            startup_menu,
            text="New",
            command=lambda: self.file_manager.create_new_template(startup_menu)
        )
        new_button.pack(pady=5)

        # Create a frame to hold the load buttons side by side
        load_button_frame = ttk.Frame(startup_menu)
        load_button_frame.pack(pady=5)

        # Load from Database Button
        load_database_button = ttk.Button(
            load_button_frame,
            text="Load from Database",
            command=lambda: self.file_manager.start_file_loading_database_wrapper(startup_menu)
        )
        load_database_button.pack(side="left", padx=5)

        # Load from Excel Button  
        load_excel_button = ttk.Button(
            load_button_frame,
            text="Load from File",
            command=lambda: self.file_manager.start_file_loading_wrapper(startup_menu)
        )
        load_excel_button.pack(side="left", padx=5)

        self.center_window(startup_menu)

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

    def populate_or_update_sheet_dropdown(self) -> None:
        """Populate or update the dropdown for sheet selection."""
        if not hasattr(self, 'drop_down_menu') or not self.drop_down_menu:
            sheet_label = ttk.Label(self.top_frame, text = "Select Test:", font=FONT, background = APP_BACKGROUND_COLOR)
            sheet_label.pack(side = "left", padx = (0,5))
            self.drop_down_menu = ttk.Combobox(
                self.top_frame,
                textvariable=self.selected_sheet,
                state="readonly",
                font=FONT
            )
            self.drop_down_menu.pack(side = "left", pady=(5, 5))
            
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

    def populate_sheet_dropdown(self) -> None:
        """Set up the dropdown for sheet selection and handle its events."""
        all_sheet_names = list(self.filtered_sheets.keys())

        # Create a dropdown menu for selecting sheets
        self.drop_down_menu = ttk.Combobox(
            self.top_frame,
            textvariable=self.selected_sheet,
            values=list(self.filtered_sheets.keys()),
            state="readonly"
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

    def on_file_selection(self, event) -> None:
        """Handle file selection from the dropdown."""
        if not hasattr(self, 'file_dropdown_var'):
            return
            
        selected_file = self.file_dropdown_var.get()
        if not selected_file or selected_file == self.current_file:
            return

        try:
            self.file_manager.set_active_file(selected_file)
            self.file_manager.update_ui_for_current_file()
        except ValueError as e:
            messagebox.showerror("Error", str(e))

    def on_app_close(self):
        """Handle application shutdown."""
        try:
            for thread in self.threads:
                if thread.is_alive():
                    pass
            self.root.destroy()
            os._exit(0)
        except Exception as e:
            debug_print(f"Error during shutdown: {e}")
            os._exit(1)

    def add_menu(self) -> None:
        """Create a top-level menu with File, Help, and About options."""
        menubar = tk.Menu(self.root)
    
        # File menu
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="New", command=self.show_new_template_dialog)
        filemenu.add_command(label="Load Excel", command=lambda: self.file_manager.load_initial_file())
        filemenu.add_command(label="Load VAP3", command=self.open_vap3_file)
        filemenu.add_separator()
        filemenu.add_command(label="Save As VAP3", command=self.save_as_vap3)
        filemenu.add_separator()
        filemenu.add_command(label="Update Database", accelerator="Ctrl+U", command=self.update_database)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.on_app_close)
        menubar.add_cascade(label="File", menu=filemenu)
    
        # View menu
        viewmenu = tk.Menu(menubar, tearoff=0)
        viewmenu.add_command(label="View Raw Data", 
                             command=lambda: self.file_manager.open_raw_data_in_excel(self.selected_sheet.get()))
        viewmenu.add_separator()
        viewmenu.add_command(label="Collect TPM Data", command=self.open_data_collection)
        viewmenu.add_command(label="Collect Sensory Data", command=self.open_sensory_data_collection)
        menubar.add_cascade(label="View", menu=viewmenu)

        # Database menu
        dbmenu = tk.Menu(menubar, tearoff=0)
        dbmenu.add_command(label="Browse Database", command=lambda: self.file_manager.show_database_browser())
        menubar.add_cascade(label="Database", menu=dbmenu)

        # Calculate menu
        calculatemenu = tk.Menu(menubar, tearoff=0)
        calculatemenu.add_command(label="Viscosity (Under Development)", command=self.open_viscosity_calculator)
        menubar.add_cascade(label="Calculate", menu=calculatemenu)

        comparemenu = tk.Menu(menubar, tearoff=0)
        comparemenu.add_command(label="Compare Loaded Samples", command=self.show_sample_comparison)
        comparemenu.add_command(label="Compare From Database", command=self.show_database_comparison)
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

        self.root.bind_all("<Control-u>", lambda e: self.update_database())

    def show_database_comparison(self):
        """Show database browser for file selection, then run comparison analysis."""
        self.file_manager.show_database_browser(comparison_mode=True)

    def load_files_from_database_for_comparison(self, selected_files):
        """Load files from database selection for comparison analysis."""
        loaded_files = []
    
        try:
            self.progress_dialog.show_progress_bar("Loading files from database...")
        
            for i, file_info in enumerate(selected_files):
                try:
                    progress = (i + 1) / len(selected_files) * 100
                    self.progress_dialog.update_progress(progress, f"Loading {file_info.get('filename', 'file')}...")
                
                    if 'filepath' in file_info:
                        self.file_manager.load_excel_file(file_info['filepath'], skip_database_storage=True)
                    elif 'id' in file_info:
                        self.file_manager.load_from_database_by_id(file_info['id'])
                    else:
                        continue
                
                    if self.filtered_sheets:
                        file_data = {
                            'file_name': file_info.get('filename', f'Database_File_{i+1}'),
                            'display_filename': file_info.get('display_name', file_info.get('filename', f'Database_File_{i+1}')),
                            'filtered_sheets': self.filtered_sheets.copy(),
                            'database_created_at': file_info.get('created_at', None),
                            'source': 'database'
                        }
                        loaded_files.append(file_data)
                    
                except Exception as e:
                    debug_print(f"DEBUG: Error loading file {file_info}: {e}")
                    continue
        
            self.progress_dialog.hide_progress_bar()
            return loaded_files
        
        except Exception as e:
            self.progress_dialog.hide_progress_bar()
            debug_print(f"DEBUG: Error in load_files_from_database_for_comparison: {e}")
            return []

    def show_sample_comparison(self):
        """Show the sample comparison window."""
        if not self.all_filtered_sheets:
            show_success_message("Info", "No files are currently loaded. Please load some data files first.", self.root)
            return
    
        from file_selection_dialog import FileSelectionDialog
        file_dialog = FileSelectionDialog(self.root, self.all_filtered_sheets)
        result, selected_files = file_dialog.show()
    
        if not result or not selected_files:
            return
    
        if len(selected_files) < 2:
            messagebox.showwarning("Warning", "Please select at least 2 files for comparison.")
            return
    
        from sample_comparison import SampleComparisonWindow
        comparison_window = SampleComparisonWindow(self, selected_files)
        comparison_window.show()  

    def show_new_template_dialog(self) -> None:
        """Show a dialog to create a new template file with selected tests."""
        self.file_manager.create_new_file_with_tests()

    def open_vap3_file(self) -> None:
        """Open a .vap3 file using the file manager."""
        self.file_manager.load_vap3_file()

    def save_with_sample_images(self):
        """Enhanced save method that includes sample images from data collection."""
        try:
            debug_print("DEBUG: Saving with sample images")
        
            # Check if we have pending sample images from data collection
            sample_images = getattr(self, 'pending_sample_images', {})
            sample_image_crop_states = getattr(self, 'pending_sample_image_crop_states', {})
            sample_header_data = getattr(self, 'pending_sample_header_data', {})
        
            if sample_images:
                debug_print(f"DEBUG: Found {len(sample_images)} sample groups to save")
            
                # Use the enhanced VAP file manager
                from vap_file_manager import VapFileManager
                vap_manager = VapFileManager()
            
                # Get current application state
                image_crop_states = getattr(self, 'image_crop_states', {})
                plot_settings = {
                    'selected_plot_type': getattr(self, 'selected_plot_type', tk.StringVar()).get()
                }
            
                # Save to VAP3 file with sample images
                success = vap_manager.save_to_vap3(
                    self.file_path,
                    self.filtered_sheets,
                    getattr(self, 'sheet_images', {}),
                    getattr(self, 'plot_options', []),
                    image_crop_states,
                    plot_settings,
                    sample_images,  # NEW: Sample images
                    sample_image_crop_states,  # NEW: Sample image crop states
                    sample_header_data  # NEW: Sample header data
                )
            
                if success:
                    # Clear pending sample images after successful save
                    if hasattr(self, 'pending_sample_images'):
                        delattr(self, 'pending_sample_images')
                    if hasattr(self, 'pending_sample_image_crop_states'):
                        delattr(self, 'pending_sample_image_crop_states')
                    if hasattr(self, 'pending_sample_header_data'):
                        delattr(self, 'pending_sample_header_data')
                
                    # Process and display the formatted images
                    self.process_formatted_sample_images()
                
                    debug_print("DEBUG: Successfully saved with sample images")
                    return True
                else:
                    debug_print("ERROR: Failed to save with sample images")
                    return False
            else:
                # No sample images, use regular save
                debug_print("DEBUG: No sample images to save, using regular save")
                return self.file_manager.save_as_vap3()
            
        except Exception as e:
            debug_print(f"ERROR: Failed to save with sample images: {e}")
            import traceback
            traceback.print_exc()
            return False

    def process_pending_sample_images(self):
        """Process and display pending sample images from data collection."""
        try:
            if not hasattr(self, 'pending_formatted_images'):
                debug_print("DEBUG: No pending formatted images to process")
                return
    
            formatted_images = self.pending_formatted_images
            debug_print(f"DEBUG: Processing {len(formatted_images)} pending sample images")
    
            # CRITICAL FIX: Use the target sheet from data collection, not currently selected sheet
            target_sheet = getattr(self, 'pending_images_target_sheet', None)
            if not target_sheet:
                debug_print("ERROR: No target sheet specified for pending images")
                # Fallback to current sheet if no target specified
                target_sheet = self.selected_sheet.get()
                if not target_sheet:
                    debug_print("DEBUG: No target sheet available for sample images")
                    return
        
            debug_print(f"DEBUG: Using target sheet for images: {target_sheet}")
    
            # Initialize image storage if needed
            if not hasattr(self, 'sheet_images'):
                self.sheet_images = {}
    
            if self.current_file not in self.sheet_images:
                self.sheet_images[self.current_file] = {}
    
            if target_sheet not in self.sheet_images[self.current_file]:
                self.sheet_images[self.current_file][target_sheet] = []
    
            # Add the sample images to the target sheet
            for img_info in formatted_images:
                img_path = img_info['path']
        
                # Add to target sheet images if not already present
                if img_path not in self.sheet_images[self.current_file][target_sheet]:
                    self.sheet_images[self.current_file][target_sheet].append(img_path)
        
                # Store crop state
                if not hasattr(self, 'image_crop_states'):
                    self.image_crop_states = {}
                self.image_crop_states[img_path] = img_info.get('crop_state', False)
    
            # Update the image display ONLY if the target sheet is currently displayed
            current_displayed_sheet = self.selected_sheet.get()
            if target_sheet == current_displayed_sheet:
                if hasattr(self, 'image_loader') and self.image_loader:
                    debug_print("DEBUG: Refreshing main GUI image display with sample images")
        
                    # Load the updated image list
                    self.image_loader.load_images_from_list(self.sheet_images[self.current_file][target_sheet])
        
                    # Restore crop states
                    for img_path in self.sheet_images[self.current_file][target_sheet]:
                        if img_path in self.image_crop_states:
                            self.image_loader.image_crop_states[img_path] = self.image_crop_states[img_path]
        
                    # Force refresh display
                    self.image_loader.display_images()
        
                    # Update the image frame
                    self.image_frame.update_idletasks()
            else:
                debug_print(f"DEBUG: Images added to {target_sheet}, but currently displaying {current_displayed_sheet}. Images will appear when {target_sheet} is selected.")
    
            # Clear pending images after processing
            delattr(self, 'pending_formatted_images')
            if hasattr(self, 'pending_images_target_sheet'):
                delattr(self, 'pending_images_target_sheet')
    
            debug_print(f"DEBUG: Successfully processed sample images for sheet {target_sheet}")
    
        except Exception as e:
            debug_print(f"ERROR: Failed to process pending sample images: {e}")
            import traceback
            traceback.print_exc()

    def process_formatted_sample_images(self):
        """Process and display formatted sample images in the main GUI."""
        try:
            formatted_images = getattr(self, 'pending_formatted_images', [])
        
            if not formatted_images:
                debug_print("DEBUG: No formatted images to process")
                return
        
            debug_print(f"DEBUG: Processing {len(formatted_images)} formatted sample images")
        
            # Get current sheet name
            current_sheet = self.selected_sheet.get()
            if not current_sheet:
                debug_print("DEBUG: No current sheet selected")
                return
        
            # Add images to current sheet's image collection
            if not hasattr(self, 'sheet_images'):
                self.sheet_images = {}
        
            if self.current_file not in self.sheet_images:
                self.sheet_images[self.current_file] = {}
        
            if current_sheet not in self.sheet_images[self.current_file]:
                self.sheet_images[self.current_file][current_sheet] = []
        
            # Add the formatted images
            for img_info in formatted_images:
                img_path = img_info['path']
            
                # Add to sheet images if not already present
                if img_path not in self.sheet_images[self.current_file][current_sheet]:
                    self.sheet_images[self.current_file][current_sheet].append(img_path)
            
                # Store crop state
                if not hasattr(self, 'image_crop_states'):
                    self.image_crop_states = {}
                self.image_crop_states[img_path] = img_info.get('crop_state', False)
        
            # Refresh the image display if we have an image loader
            if hasattr(self, 'image_loader') and self.image_loader:
                debug_print("DEBUG: Refreshing image display with sample images")
                self.image_loader.load_images_from_list(self.sheet_images[self.current_file][current_sheet])
            
                # Restore crop states
                for img_path in self.sheet_images[self.current_file][current_sheet]:
                    if img_path in self.image_crop_states:
                        self.image_loader.image_crop_states[img_path] = self.image_crop_states[img_path]
            
                # Force refresh display
                self.image_loader.display_images()
        
            # Clear pending formatted images
            if hasattr(self, 'pending_formatted_images'):
                delattr(self, 'pending_formatted_images')
        
            debug_print(f"DEBUG: Successfully processed sample images for sheet {current_sheet}")
        
        except Exception as e:
            debug_print(f"ERROR: Failed to process formatted sample images: {e}")
            import traceback
            traceback.print_exc()

    def save_as_vap3(self) -> None:
        """Save the current session as a .vap3 file."""
        self.file_manager.save_as_vap3()

    def show_help(self):
        """Display help dialog."""
        show_success_message("Help", "This program is designed to be used with excel data according to the SDR Standardized Testing Template.\n \nClick 'Generate Test Report' to create an excel report of a single test, or click 'Generate Full Report' to generate both an excel file and powerpoint file of all the contents within the file.", self.root)

    def show_about(self) -> None:
        """Display about dialog."""
        show_success_message("About", "SDR DataViewer Beta Version 3.0\nDeveloped by Charlie Becquet", self.root)

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

    def set_window_size(self, width_ratio: float, height_ratio: float) -> None:
        """Set the window size as a percentage of the screen dimensions."""
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * width_ratio)
        window_height = int(screen_height * height_ratio)
        position_x = (screen_width - window_width) // 2
        position_y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

    def add_report_buttons(self, parent_frame: ttk.Frame) -> None:
        """Add buttons for generating reports and align them in the parent frame."""
        # Generate Full Report button
        full_report_btn = ttk.Button(parent_frame, text="Generate Full Report", command=self.generate_full_report)
        full_report_btn.pack(side="left", padx=(5, 5), pady=(5, 5))

        # Generate Test Report button
        test_report_btn = ttk.Button(parent_frame, text="Generate Test Report", command=self.generate_test_report)
        test_report_btn.pack(side="left", padx=(5, 5), pady=(5, 5))

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

    def update_displayed_sheet(self, sheet_name: str) -> None:
        """Update the displayed sheet and dynamically manage the plot options and plot type dropdown."""
        # Use background thread for heavy processing to keep UI responsive
        def update_in_background():
            try:
                processing, get_valid_plot_options = _lazy_import_processing()
                if not processing:
                    return

                pd = self.get_pandas()
                if not pd:
                    return

                # Ensure bottom_frame maintains its fixed height
                if hasattr(self, 'bottom_frame') and self.bottom_frame.winfo_exists():
                    self.bottom_frame.configure(height=150)
                    self.bottom_frame.pack_propagate(False)
                    self.bottom_frame.grid_propagate(False)

                if not sheet_name or sheet_name not in self.filtered_sheets:
                    return

                sheet_info = self.filtered_sheets.get(sheet_name)
                if not sheet_info:
                    self.root.after(0, lambda: messagebox.showerror("Error", f"Sheet '{sheet_name}' not found."))
                    return

                data = sheet_info["data"]
                is_empty = sheet_info.get("is_empty", False)
                is_plotting_sheet = plotting_sheet_test(sheet_name, data)

                # Update UI in main thread
                self.root.after(0, lambda: self._finish_sheet_update(sheet_name, data, is_empty, is_plotting_sheet, processing))

            except Exception as e:
                debug_print(f"Error in background sheet update: {e}")

        # Start background processing
        thread = threading.Thread(target=update_in_background, daemon=True)
        thread.start()

    def _finish_sheet_update(self, sheet_name, data, is_empty, is_plotting_sheet, processing):
        """Finish sheet update in main thread."""
        try:
            # Clear and rebuild frames
            self.clear_dynamic_frame()
            self.setup_dynamic_frames(is_plotting_sheet)

            # Handle ImageLoader setup
            self._setup_image_loader(sheet_name, is_plotting_sheet)

            # Handle empty sheets
            if is_empty or data.empty:
                self._handle_empty_sheet(sheet_name, is_plotting_sheet)
                self.root.update_idletasks()
                return

            # Process the sheet data
            try:        
                process_function = processing.get_processing_function(sheet_name)
                processed_data, _, full_sample_data = process_function(data)
            
                # Apply empty sample filtering ONLY to plotting sheets
                debug_print(f"DEBUG: Before filtering - processed_data shape: {processed_data.shape}, full_sample_data shape: {full_sample_data.shape}")
                debug_print(f"DEBUG: is_plotting_sheet: {is_plotting_sheet}")
            
                if is_plotting_sheet:
                    # Filter empty samples from processed data for plotting sheets
                    filtered_processed_data = filter_empty_samples_from_dataframe(processed_data)
                    debug_print(f"DEBUG: After filtering plotting sheet - processed_data shape: {filtered_processed_data.shape}")
                    processed_data = filtered_processed_data
                else:
                    # For non-plotting sheets, keep all data as-is
                    debug_print(f"DEBUG: Non-plotting sheet - preserving all data as-is")
            
                self.current_sheet_data = processed_data

                # Handle User Test Simulation
                if sheet_name in ["User Test Simulation", "User Simulation Test"]:
                    self.plot_options = self.user_test_simulation_plot_options
                    self.is_user_test_simulation = True
                    self.num_columns_per_sample = 8
                else:
                    self.plot_options = self.standard_plot_options
                    self.is_user_test_simulation = False
                    self.num_columns_per_sample = 12

            except Exception as e:
                debug_print(f"ERROR: Processing function failed for {sheet_name}: {e}")
                messagebox.showerror("Processing Error", f"Error processing sheet '{sheet_name}': {e}")
                return

            # Display the table
            try:
                self.display_table(self.table_frame, processed_data, sheet_name, is_plotting_sheet)
            except Exception as e:
                debug_print(f"ERROR: Failed to display table: {e}")

            # Display plot if it's a plotting sheet

            if is_plotting_sheet:
                try:
                    if not full_sample_data.empty:
                        self.display_plot(full_sample_data)
                    else:
                        self._show_empty_plot_message()
                except Exception as e:
                    debug_print(f"ERROR: Failed to display plot: {e}")

            self.root.update_idletasks()

        except Exception as e:
            debug_print(f"Error finishing sheet update: {e}")
    def _force_plot_redraw(self):
        """Force a complete plot redraw to ensure filtered data displays correctly."""
        try:
            if hasattr(self, 'plot_manager') and hasattr(self.plot_manager, 'canvas'):
                if self.plot_manager.canvas:
                    debug_print("DEBUG: Forcing complete plot redraw")
                    self.plot_manager.canvas.draw_idle()
                    self.plot_frame.update_idletasks()
        except Exception as e:
            debug_print(f"DEBUG: Error in plot redraw: {e}")

    def _setup_image_loader(self, sheet_name, is_plotting_sheet):
        """Setup image loader for current sheet."""
        for widget in self.image_frame.winfo_children():
            widget.destroy()
            
        if hasattr(self, 'image_loader'):
            del self.image_loader

        from image_loader import ImageLoader
        self.image_loader = ImageLoader(self.image_frame, is_plotting_sheet, on_images_selected=lambda paths: self.store_images(sheet_name, paths), main_gui = self)
        self.image_loader.frame.config(height=150)

        # Restore images if they exist
        current_file = self.current_file
        if current_file in self.sheet_images and sheet_name in self.sheet_images[current_file]:
            self.image_loader.load_images_from_list(self.sheet_images[current_file][sheet_name])
            for img_path in self.sheet_images[current_file][sheet_name]:
                if img_path in self.image_crop_states:
                    self.image_loader.image_crop_states[img_path] = self.image_crop_states[img_path]

        self.image_loader.display_images()
        self.image_frame.update_idletasks()

    def _handle_empty_sheet(self, sheet_name, is_plotting_sheet):
        """Handle empty sheet display."""
        pd = self.get_pandas()
        if not pd:
            return

        if is_plotting_sheet:
            # Create completely empty DataFrame with proper columns for empty state
            columns = [
                "Sample Name", "Media", "Viscosity", "Voltage, Resistance, Power",
                "Average TPM", "Standard Deviation", "Initial Oil Mass", 
                "Usage Efficiency", "Burn?", "Clog?", "Leak?"
            ]
            empty_data = pd.DataFrame(columns=columns)
            
            debug_print("DEBUG: Displaying empty plotting sheet with no samples")
            self.display_table(self.table_frame, empty_data, sheet_name, is_plotting_sheet)
            self._show_empty_plot_message()
        else:
            minimal_data = pd.DataFrame([{
                "Status": "No data - ready for collection",
                "Instructions": "Use data collection tools to add information"
            }])
            self.display_table(self.table_frame, minimal_data, sheet_name, is_plotting_sheet)

    def _show_empty_plot_message(self):
        """Show empty plot message."""
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

    def get_modified_files(self):
        """Get list of files that have been modified in the staging area."""
        modified_files = []
        if hasattr(self, 'all_filtered_sheets'):
            for file_data in self.all_filtered_sheets:
                if file_data.get('is_modified', False):
                    modified_files.append(file_data)
        return modified_files

    def clear_modified_flags(self):
        """Clear all modification flags after database update."""
        if hasattr(self, 'all_filtered_sheets'):
            for file_data in self.all_filtered_sheets:
                file_data['is_modified'] = False
                if 'last_modified' in file_data:
                    del file_data['last_modified']
    
        # Update window title to remove modified indicator
        current_title = self.root.title()
        if current_title.endswith(" *"):
            self.root.title(current_title[:-2])

    def update_database(self):
        """Update the database with all staged changes."""
        try:
            modified_files = self.get_modified_files()
        
            if not modified_files:
                show_success_message("No Changes", "No files have been modified. Nothing to update.", self.root)
                return
        
            # Show confirmation dialog
            file_names = [f["file_name"] for f in modified_files]
            message = f"Update database with changes to the following files?\n\n"
            message += "\n".join([f"• {name}" for name in file_names])
        
            if not messagebox.askyesno("Confirm Database Update", message):
                return
        
            # Show progress dialog
            self.progress_dialog.show_progress_bar("Updating database...")
            self.root.update_idletasks()
        
            total_files = len(modified_files)
            successful_updates = 0
            failed_updates = []
        
            for i, file_data in enumerate(modified_files):
                try:
                    # Update progress
                    progress = int(((i + 1) / total_files) * 100)
                    self.progress_dialog.update_progress_bar(progress)
                    self.root.update_idletasks()
                
                    debug_print(f"DEBUG: Updating database for file: {file_data['file_name']}")
                
                    # Save current file state as VAP3 and store in database
                    self._update_file_in_database(file_data)
                    successful_updates += 1
                
                except Exception as e:
                    debug_print(f"ERROR: Failed to update file {file_data['file_name']}: {e}")
                    failed_updates.append(file_data['file_name'])
        
            # Clean up
            self.progress_dialog.hide_progress_bar()
        
            # Show results
            if failed_updates:
                if successful_updates > 0:
                    message = f"Partial success:\n\n"
                    message += f"✓ Successfully updated: {successful_updates} files\n"
                    message += f"✗ Failed to update: {len(failed_updates)} files\n\n"
                    message += "Failed files:\n" + "\n".join([f"• {name}" for name in failed_updates])
                    messagebox.showwarning("Partial Success", message)
                else:
                    message = f"Failed to update all {len(failed_updates)} files:\n\n"
                    message += "\n".join([f"• {name}" for name in failed_updates])
                    messagebox.showerror("Update Failed", message)
            else:
                # Complete success
                message = f"Successfully updated {successful_updates} file(s) in the database."
                show_success_message("Database Updated", message, self.root)
            
                # Clear modification flags
                self.clear_modified_flags()
        
        except Exception as e:
            self.progress_dialog.hide_progress_bar()
            messagebox.showerror("Error", f"Failed to update database: {e}")
            debug_print(f"ERROR: Database update failed: {e}")
            import traceback
            traceback.print_exc()

    def _update_file_in_database(self, file_data):
        """Update a single file in the database."""
        try:
            # Create temporary VAP3 file
            import tempfile
            with tempfile.NamedTemporaryFile(suffix='.vap3', delete=False) as temp_file:
                temp_vap3_path = temp_file.name
        
            # Prepare data for VAP3 save
            filtered_sheets = file_data["filtered_sheets"]
        
            # Get associated images for this file
            sheet_images = {}
            if hasattr(self, 'sheet_images') and file_data["file_name"] in self.sheet_images:
                sheet_images = {file_data["file_name"]: self.sheet_images[file_data["file_name"]]}
        
            # Get image crop states
            image_crop_states = getattr(self, 'image_crop_states', {})
        
            # Plot settings
            plot_settings = {}
            if hasattr(self, 'selected_plot_type'):
                plot_settings['selected_plot_type'] = self.selected_plot_type.get()
        
            # Save as VAP3
            from vap_file_manager import VapFileManager
            vap_manager = VapFileManager()
        
            success = vap_manager.save_to_vap3(
                temp_vap3_path,
                filtered_sheets,
                sheet_images,
                getattr(self, 'plot_options', []),
                image_crop_states,
                plot_settings
            )
        
            if not success:
                raise Exception("Failed to create temporary VAP3 file")
        
            # Update database
            original_filename = file_data.get("original_filename", file_data["file_name"])
            display_filename = file_data["file_name"]
            if not display_filename.endswith('.vap3'):
                display_filename = os.path.splitext(display_filename)[0] + '.vap3'
        
            meta_data = {
                'display_filename': display_filename,
                'original_filename': original_filename,
                'original_path': file_data.get("file_path", ""),
                'creation_date': time.strftime('%Y-%m-%d %H:%M:%S'),
                'last_modified': time.strftime('%Y-%m-%d %H:%M:%S'),
                'sheet_count': len(filtered_sheets),
                'plot_options': getattr(self, 'plot_options', []),
                'plot_settings': plot_settings
            }
        
            # Store/update in database
            file_id = self.file_manager.db_manager.store_vap3_file(temp_vap3_path, meta_data)
        
            # Store sheet metadata
            processing_module, _ = _lazy_import_processing()
            if processing_module:
                for sheet_name, sheet_info in filtered_sheets.items():
                    is_plotting = processing_module.plotting_sheet_test(sheet_name, sheet_info["data"])
                    is_empty = sheet_info.get("is_empty", False)
                
                    self.file_manager.db_manager.store_sheet_info(
                        file_id, 
                        sheet_name, 
                        is_plotting, 
                        is_empty
                    )
            else:
                # Fallback if processing can't be loaded
                for sheet_name, sheet_info in filtered_sheets.items():
                    # Simple heuristic check
                    is_plotting = len(sheet_info["data"]) > 0 and len(sheet_info["data"].columns) > 5
                    is_empty = sheet_info.get("is_empty", False)
                
                    self.file_manager.db_manager.store_sheet_info(
                        file_id, 
                        sheet_name, 
                        is_plotting, 
                        is_empty
                    )
        
            # Clean up temporary file
            try:
                os.unlink(temp_vap3_path)
            except Exception:
                pass
        
            debug_print(f"DEBUG: Successfully updated file {file_data['file_name']} in database")
        
        except Exception as e:
            debug_print(f"ERROR: Failed to update file {file_data['file_name']} in database: {e}")
            raise
   
    def load_images(self, is_plotting_sheet):
        """Load and display images in the dynamic image_frame."""
        if not hasattr(self, 'image_frame') or not self.image_frame.winfo_exists():
            self.image_frame = ttk.Frame(self.bottom_frame, height=150)
            self.image_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            self.image_frame.pack_propagate(True)
        if not self.image_loader:
            from image_loader import ImageLoader
            self.image_loader = ImageLoader(self.image_frame, is_plotting_sheet)
        self.image_loader.add_images()

    def load_sample_images_from_vap3(self, vap_data):
        """Load sample images from VAP3 data and populate both main GUI and metadata structures."""
        try:
            sample_images = vap_data.get('sample_images', {})
            sample_metadata = vap_data.get('sample_images_metadata', {})
            sample_crop_states = vap_data.get('sample_image_crop_states', {})

            if not sample_images:
                debug_print("DEBUG: No sample images found in VAP3 data")
                return

            debug_print(f"DEBUG: Loading sample images from VAP3: {len(sample_images)} samples")

            # Convert sample images to formatted images for main GUI display
            formatted_images = []
            header_data = sample_metadata.get('header_data', {})
            test_name = header_data.get('test', 'Unknown Test')
    
            # Try to find the correct sheet name if test_name doesn't match exactly
            target_sheet = test_name
            if test_name not in self.filtered_sheets:
                # Try to find a matching sheet name
                for sheet_name in self.filtered_sheets.keys():
                    if (sheet_name.lower() == test_name.lower() or 
                        test_name.lower() in sheet_name.lower() or
                        sheet_name.lower() in test_name.lower()):
                        target_sheet = sheet_name
                        debug_print(f"DEBUG: Matched test_name '{test_name}' to sheet '{sheet_name}'")
                        break

            debug_print(f"DEBUG: Using target sheet: {target_sheet}")

            for sample_id, image_paths in sample_images.items():
                try:
                    # Extract sample index
                    sample_index = int(sample_id.split()[-1]) - 1
                except (ValueError, IndexError):
                    sample_index = 0

                # Get sample info from header data
                sample_info = {}
                if 'samples' in header_data and sample_index < len(header_data['samples']):
                    sample_info = header_data['samples'][sample_index]

                # Create labels for each image
                for img_path in image_paths:
                    # Create descriptive label
                    label_parts = [
                        sample_info.get('id', sample_id),
                        target_sheet,  # Use the matched sheet name
                        sample_info.get('media', 'Unknown Media'),
                        f"{sample_info.get('viscosity', 'Unknown')} cP",
                        sample_metadata.get('timestamp', '')[:10]  # Date only
                    ]
        
                    formatted_label = " - ".join(filter(None, label_parts))
        
                    formatted_images.append({
                        'path': img_path,
                        'label': formatted_label,
                        'sample_id': sample_id,
                        'crop_state': sample_crop_states.get(img_path, False)
                    })

            # Store for processing
            self.pending_formatted_images = formatted_images
            self.pending_images_target_sheet = target_sheet

            # CRITICAL: Also populate the sample_image_metadata structure for data collection
            if not hasattr(self, 'sample_image_metadata'):
                self.sample_image_metadata = {}
    
            current_file = getattr(self, 'current_file', 'Unknown File')
            if current_file not in self.sample_image_metadata:
                self.sample_image_metadata[current_file] = {}
    
            if target_sheet not in self.sample_image_metadata[current_file]:
                self.sample_image_metadata[current_file][target_sheet] = {}
    
            # Store the sample-to-image mapping for later retrieval by data collection
            self.sample_image_metadata[current_file][target_sheet] = {
                'sample_images': sample_images.copy(),
                'sample_image_crop_states': sample_crop_states.copy(),
                'header_data': header_data.copy(),
                'test_name': target_sheet
            }
            debug_print(f"DEBUG: Populated sample_image_metadata for {target_sheet} with {len(sample_images)} samples")

            # Process and display immediately
            self.process_pending_sample_images()

            debug_print(f"DEBUG: Successfully loaded {len(formatted_images)} sample images from VAP3")

        except Exception as e:
            debug_print(f"ERROR: Failed to load sample images from VAP3: {e}")
            import traceback
            traceback.print_exc()

    def open_data_collection(self):
        """Open data collection interface, handling both loaded and unloaded file states."""
        if not hasattr(self, 'filtered_sheets') or not self.filtered_sheets or not hasattr(self, 'file_path') or not self.file_path:
            self.show_data_collection_startup_dialog()
        else:
            # Get the original filename for .vap3 files loaded from database
            original_filename = None
        
            # Check if this file was loaded from database and has original filename metadata
            if hasattr(self, 'all_filtered_sheets') and self.all_filtered_sheets:
                current_file_data = None
                for file_data in self.all_filtered_sheets:
                    if file_data.get("file_path") == self.file_path:
                        current_file_data = file_data
                        break
            
                if current_file_data:
                    # Check for original filename in metadata
                    original_filename = current_file_data.get('original_filename')
                    if not original_filename:
                        # Try database filename as fallback
                        original_filename = current_file_data.get('database_filename')
                
                    debug_print(f"DEBUG: Found original filename for data collection: {original_filename}")
        
            self.file_manager.show_test_start_menu(self.file_path, original_filename=original_filename)

    def open_sensory_data_collection(self):
        """Open the sensory data collection window."""
        from sensory_data_collection import SensoryDataCollectionWindow

        # Store current window state before hiding
        print("DEBUG: Storing main window state before opening sensory window")
        current_state = self.root.state()
        print(f"DEBUG: Current window state: {current_state}")
    
        self.root.withdraw()

        def on_sensory_window_close():
            print("DEBUG: Sensory window close callback triggered")
            print("DEBUG: Restoring main window...")
        
            # Restore the main window
            self.root.deiconify()
            print("DEBUG: Main window deiconified")
        
            # Restore fullscreen state 
            self.root.state('zoomed')
            print("DEBUG: Main window restored to fullscreen/zoomed state")
        
            # Restore app colors and styling
            self.set_app_colors()
            print("DEBUG: App colors and styling restored")
        
            # Bring window to front and focus
            self.root.lift()
            self.root.focus_set()
            print("DEBUG: Main window brought to front and focused")
        
            # Force a complete redraw of the interface
            self.root.update_idletasks()
            print("DEBUG: Main window redraw completed")

        sensory_window = SensoryDataCollectionWindow(self.root, close_callback=on_sensory_window_close)
        sensory_window.show()

    def show_data_collection_startup_dialog(self):
        """Show a startup dialog for data collection when no file is loaded."""
        startup_dialog = tk.Toplevel(self.root)
        startup_dialog.title("Data Collection")
        startup_dialog.geometry("350x200")
        startup_dialog.transient(self.root)
        startup_dialog.grab_set()
        startup_dialog.configure(bg=APP_BACKGROUND_COLOR)

        ttk.Label(
            startup_dialog, 
            text="Start Data Collection", 
            font=("Arial", 16, "bold"),
            background=APP_BACKGROUND_COLOR
        ).pack(pady=(20, 10))

        ttk.Label(
            startup_dialog,
            text="Choose an option to begin collecting data:",
            font=FONT,
            background=APP_BACKGROUND_COLOR
        ).pack(pady=(0, 20))

        button_frame = ttk.Frame(startup_dialog)
        button_frame.pack(pady=10)

        ttk.Button(
            button_frame,
            text="Create New File",
            command=lambda: self.handle_data_collection_new(startup_dialog),
            width=15
        ).pack(side="left", padx=10)

        ttk.Button(
            button_frame,
            text="Load Existing File",
            command=lambda: self.handle_data_collection_load(startup_dialog),
            width=15
        ).pack(side="left", padx=10)

        ttk.Button(
            button_frame,
            text="Cancel",
            command=startup_dialog.destroy,
            width=10
        ).pack(pady=(20, 0))

        self.center_window(startup_dialog)

    def handle_data_collection_new(self, dialog):
        """Handle 'Create New File' from data collection dialog."""
        dialog.destroy()
        file_path = self.file_manager.create_new_file_with_tests()

    def handle_data_collection_load(self, dialog):
        """Handle 'Load Existing File' from data collection dialog."""
        dialog.destroy()
        
        file_path = self.file_manager.ask_open_file()
        if file_path:
            try:
                self.file_manager.load_excel_file(file_path)
                
                self.all_filtered_sheets.append({
                    "file_name": os.path.basename(file_path),
                    "file_path": file_path,
                    "filtered_sheets": copy.deepcopy(self.filtered_sheets)
                })
                
                self.file_manager.update_file_dropdown()
                self.file_manager.set_active_file(os.path.basename(file_path))
                self.file_manager.update_ui_for_current_file()
                
                self.file_manager.show_test_start_menu(file_path)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")

    def display_plot(self, full_sample_data):
        """Display the plot in the plot frame based on the current data."""
        if not hasattr(self, 'plot_frame') or self.plot_frame is None:
            self.plot_frame = ttk.Frame(self.display_frame)
            self.plot_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)

        # Clear existing plot contents
        for widget in self.plot_frame.winfo_children():
            widget.destroy()

        # Check if data is empty after filtering
        if full_sample_data.empty or full_sample_data.shape[1] == 0:
            debug_print("DEBUG: No plottable data available after filtering empty samples")
            self._show_empty_plot_message()
            return

        num_columns = getattr(self, 'num_columns_per_sample', 12)
        self.plot_manager.plot_all_samples(self.plot_frame, full_sample_data, num_columns)
        self.plot_frame.grid_propagate(True)
        self.plot_manager.add_plot_dropdown(self.plot_frame)
        self.plot_frame.update_idletasks()

    def display_table(self, frame, data, sheet_name, is_plotting_sheet):
        """Display table data in the given frame using tksheet with robust error handling."""
        pd = self.get_pandas()
        np = self.get_numpy()
        Sheet = self.get_tksheet()

        if not Sheet or not np or not pd:
            debug_print("ERROR: Required libraries not available. Aborting display_table.")
            return

        for widget in frame.winfo_children():
            widget.destroy()

        # Handle completely empty data (after filtering)
        if data.empty:
            debug_print(f"DEBUG: Data is completely empty after filtering for sheet '{sheet_name}'")
        
            # Create placeholder indicating ready for data collection
            placeholder_label = tk.Label(
                frame,
                text=f"No samples to display.\nAll samples were empty or placeholder data.\nReady for data collection.",
                font=("Arial", 12),
                fg="gray",
                justify="center"
            )
            placeholder_label.pack(expand=True, pady=50)
            return

        debug_print("DEBUG: Cleared existing table widgets.")

        # Clean and prepare data with enhanced error handling
        try:
            data = clean_columns(data)
            data.columns = data.columns.map(str)  # Ensure column headers are strings
            debug_print("DEBUG: Column cleaning completed successfully.")

            # Fix: More robust empty check
            try:
                is_data_empty = len(data) == 0 or data.shape[0] == 0
            except Exception as e:
                debug_print(f"DEBUG: Error checking if data is empty, assuming not empty: {e}")
                is_data_empty = False

            if is_data_empty:
                debug_print("DEBUG: Data is empty. Displaying warning.")
                messagebox.showwarning("Warning", f"Sheet '{sheet_name}' contains no data to display.")
                return

            # Fix: More robust data type conversion
            try:
                # Convert to string and handle NaN values in one step
                data = data.astype(str)
                data = data.replace(['nan', 'None', 'NaN'], '', regex=False)
                debug_print("DEBUG: Data type conversion completed successfully.")
            except Exception as type_error:
                debug_print(f"DEBUG: Error in data type conversion: {type_error}")
                # Final fallback - create a clean copy
                try:
                    clean_data = pd.DataFrame()
                    for col in data.columns:
                        clean_data[col] = data[col].apply(lambda x: str(x) if pd.notna(x) else '')
                    data = clean_data
                    debug_print("DEBUG: Used fallback data conversion method.")
                except Exception as fallback_error:
                    debug_print(f"ERROR: Even fallback conversion failed: {fallback_error}")
                    return

            debug_print("DEBUG: Data processed successfully.")

        except Exception as data_prep_error:
            debug_print(f"ERROR: Critical error in data preparation: {data_prep_error}")
            import traceback
            traceback.print_exc()
    
            # Show error message to user
            error_label = tk.Label(
                frame,
                text=f"Error displaying table data for '{sheet_name}'.\nPlease check the console for details.",
                font=("Arial", 12),
                fg="red",
                justify="center"
            )
            error_label.pack(expand=True, pady=50)
            return

        # Continue with table creation using tksheet
        try:
            table_frame = ttk.Frame(frame, padding=(2, 1))
            table_frame.pack(fill='both', expand=True)
        
            # Configure grid weights for proper resizing
            table_frame.grid_rowconfigure(0, weight=1)
            table_frame.grid_columnconfigure(0, weight=1)

            debug_print("DEBUG: Table frame created and packed.")

            # IMPROVED: Dynamic sizing parameters based on font and maximum column width
            font_size = 10  # Font size in points
            char_width_pixels = 8  # Approximate pixels per character for Arial 10pt
            font_height_pixels = 12  # Font height in pixels
        
            # Set maximum column width based on "Test Method" column from your image (1920x1080 screen)
            max_column_width = 380  # Pixels
            min_column_width = 120  # Minimum for readability
        
            # Calculate max characters per line dynamically
            max_chars_per_line = max_column_width // char_width_pixels
            debug_print(f"DEBUG: Using max_column_width={max_column_width}px, char_width={char_width_pixels}px, max_chars_per_line={max_chars_per_line}")
        
            # Row height parameters
            min_row_height = 25  # Minimum row height
            row_padding = 8  # Extra pixels for padding

            # Pre-process data for text wrapping using existing wrap_text function
            debug_print("DEBUG: Pre-processing text for table display with wrapping using utils.wrap_text.")
    
            # Apply text wrapping to all cells and calculate dimensions
            wrapped_data = data.copy()
            row_heights = []
            col_widths = []
    
            debug_print(f"DEBUG: Processing {len(data)} rows and {len(data.columns)} columns for text wrapping")
        
            # First pass: Calculate column widths based on content
            for col_idx, col_name in enumerate(data.columns):
                max_chars_in_column = len(str(col_name))  # Start with header length
            
                # Check all cells in this column
                for row_idx in range(len(data)):
                    try:
                        cell_value = data.iloc[row_idx, col_idx]
                        if pd.isna(cell_value) or cell_value in ['nan', 'None', 'NaN', '']:
                            cell_text = ''
                        else:
                            cell_text = str(cell_value).strip()
                    
                        if cell_text:
                            # Find the longest line after wrapping
                            wrapped_text = wrap_text(cell_text, max_width=max_chars_per_line)
                            if wrapped_text:
                                lines = wrapped_text.split('\n')
                                max_line_length = max(len(line) for line in lines)
                                max_chars_in_column = max(max_chars_in_column, max_line_length)
                        
                    except Exception as cell_error:
                        debug_print(f"DEBUG: Error processing cell at ({row_idx}, {col_idx}) for width: {cell_error}")
                        continue
            
                # Convert character count to pixel width, respecting min/max limits
                calculated_width = max_chars_in_column * char_width_pixels + 5  # 20px padding
                column_width = max(min_column_width, min(max_column_width, calculated_width))
                col_widths.append(column_width)
            
                debug_print(f"DEBUG: Column '{col_name}' - max_chars: {max_chars_in_column}, calculated_width: {calculated_width}px, final_width: {column_width}px")
        
            # Second pass: Apply text wrapping and calculate row heights
            for row_idx in range(len(data)):
                max_lines_in_row = 1
            
                for col_idx, col_name in enumerate(data.columns):
                    try:
                        cell_value = data.iloc[row_idx, col_idx]
                    
                        # Convert to string and handle empty/nan values
                        if pd.isna(cell_value) or cell_value in ['nan', 'None', 'NaN', '']:
                            cell_text = ''
                        else:
                            cell_text = str(cell_value).strip()
                    
                        if cell_text:
                            # Use the existing wrap_text function from utils
                            wrapped_text = wrap_text(cell_text, max_width=max_chars_per_line)
                            wrapped_data.iloc[row_idx, col_idx] = wrapped_text
                        
                            # Calculate lines for row height
                            if wrapped_text:
                                line_count = wrapped_text.count('\n') + 1
                                max_lines_in_row = max(max_lines_in_row, line_count)
                        else:
                            wrapped_data.iloc[row_idx, col_idx] = ''
                        
                    except Exception as cell_error:
                        debug_print(f"DEBUG: Error processing cell at ({row_idx}, {col_idx}): {cell_error}")
                        wrapped_data.iloc[row_idx, col_idx] = str(cell_value) if 'cell_value' in locals() else ''
                        continue
            
                # Calculate row height based on max lines in this row
                row_height = max(min_row_height, max_lines_in_row * font_height_pixels + row_padding)
                row_heights.append(row_height)
    
            debug_print(f"DEBUG: Calculated row heights: min={min(row_heights)}, max={max(row_heights)}, avg={sum(row_heights)/len(row_heights):.1f}")
            debug_print(f"DEBUG: Calculated column widths: min={min(col_widths)}, max={max(col_widths)}, avg={sum(col_widths)/len(col_widths):.1f}")

            # Prepare data for tksheet
            headers = list(data.columns)
            rows_data = wrapped_data.values.tolist()
        
            debug_print(f"DEBUG: Prepared {len(headers)} columns and {len(rows_data)} rows for tksheet")

            # Create tksheet widget with auto-sizing DISABLED
            sheet = Sheet(
                table_frame,
                data=rows_data,
                headers=headers,
                show_table=True,
                show_row_index=True,
                show_header=True,
                show_top_left=True,
                empty_horizontal=0,
                empty_vertical=0,
                selected_rows_to_end_of_window=False,
                horizontal_grid_to_end_of_window=False,
                vertical_grid_to_end_of_window=False,
                show_horizontal_grid=True,
                show_vertical_grid=True,
                auto_resize_default_row_index=False,  # Disable auto-resize
                auto_resize_default_header=False,     # Disable auto-resize
                auto_resize_row_index=False,          # Disable auto-resize
                auto_resize_columns=False,            # Disable auto-resize
                height=400,  # Set reasonable default height
                width=600    # Set reasonable default width
            )

            # Configure tksheet styling with proper 3-tuple font specifications
            sheet.set_options(
                theme="light blue",  # Similar to the #4CC9F0 background
                font=("Arial", font_size, "normal"),  # Use our font_size variable
                header_font=("Arial", font_size, "bold"),
                index_font=("Arial", font_size, "normal"),
                show_dropdown_borders=False,
                redraw_header_grid=True,
                redraw_row_index_grid=True
            )

            # IMPORTANT: Set column widths BEFORE row heights to prevent auto-sizing interference
            debug_print("DEBUG: Setting column widths first to prevent auto-sizing...")
            for col_idx, width in enumerate(col_widths):
                try:
                    sheet.column_width(column=col_idx, width=int(width))
                    debug_print(f"DEBUG: Set column {col_idx} width to {int(width)}px")
                except Exception as e:
                    debug_print(f"DEBUG: Error setting column {col_idx} width: {e}")

            # Set row heights AFTER column widths
            debug_print("DEBUG: Setting individual row heights...")
            for row_idx, height in enumerate(row_heights):
                try:
                    sheet.row_height(row=row_idx, height=int(height))
                except Exception as e:
                    debug_print(f"DEBUG: Error setting row {row_idx} height: {e}")

            # Disable any remaining auto-sizing features that might interfere
            sheet.disable_bindings(
                "edit_cell",
                "edit_header", 
                "edit_index",
                "copy",
                "cut",
                "paste",
                "delete",
                "insert_columns",
                "insert_rows",
                "delete_columns", 
                "delete_rows"
            )

            # Enable only viewing and selection bindings
            sheet.enable_bindings(
                "single_select",
                "drag_select", 
                "select_all",
                "column_select",
                "row_select",
                "column_width_resize",      # Allow manual resizing
                "double_click_column_resize",  # Allow double-click auto-size
                "row_height_resize",       # Allow manual resizing
                "arrowkeys",
                "right_click_popup_menu",
                "rc_select"
            )

            # Grid the sheet widget
            sheet.grid(row=0, column=0, sticky="nsew")

            # FINAL: Force refresh and ensure our sizing sticks
            sheet.refresh()
        
            # Double-check our column widths after refresh (in case tksheet changed them)
            debug_print("DEBUG: Verifying column widths after refresh...")
            for col_idx, expected_width in enumerate(col_widths):
                try:
                    actual_width = sheet.column_width(column=col_idx)
                    if abs(actual_width - expected_width) > 5:  # Allow 5px tolerance
                        debug_print(f"DEBUG: Column {col_idx} width mismatch - expected: {expected_width}, actual: {actual_width}, resetting...")
                        sheet.column_width(column=col_idx, width=int(expected_width))
                except Exception as e:
                    debug_print(f"DEBUG: Error verifying column {col_idx} width: {e}")
        
            debug_print("DEBUG: tksheet table displayed successfully.")

            # Store reference to sheet for potential future use
            if not hasattr(self, 'current_sheet_widget'):
                self.current_sheet_widget = {}
            self.current_sheet_widget[sheet_name] = sheet

        except Exception as table_error:
            debug_print(f"ERROR: Failed to create tksheet table: {table_error}")
            import traceback
            traceback.print_exc()
    
            # Show error message to user
            error_label = tk.Label(
                frame,
                text=f"Error creating table for '{sheet_name}'.\nPlease check the console for details.",
                font=("Arial", 12),
                fg="red",
                justify="center"
            )
            error_label.pack(expand=True, pady=50)

    def store_vap3_data_for_data_collection(self, vap_data):
        """Store VAP3 data for access by data collection windows."""
        try:
            debug_print("DEBUG: Storing VAP3 data for data collection access")
        
            # Store the complete VAP3 data for data collection windows to access
            self.current_vap_data = vap_data
        
            # Log what we're storing
            sample_images = vap_data.get('sample_images', {})
            if sample_images:
                debug_print(f"DEBUG: Stored VAP3 data with {len(sample_images)} sample image groups")
                for sample_id, images in sample_images.items():
                    debug_print(f"DEBUG: Stored {sample_id}: {len(images)} images")
            else:
                debug_print("DEBUG: No sample images in VAP3 data")
            
        except Exception as e:
            debug_print(f"ERROR: Failed to store VAP3 data: {e}")
            import traceback
            traceback.print_exc()

    def store_images(self, sheet_name, paths):
        """Store image paths and their crop states for a specific sheet."""
        if not sheet_name:
            return

        if self.current_file not in self.sheet_images:
            self.sheet_images[self.current_file] = {}
        self.sheet_images[self.current_file][sheet_name] = paths

        if not hasattr(self, 'image_crop_states'):
            self.image_crop_states = {}

        for img_path in paths:
            if img_path not in self.image_crop_states:
                self.image_crop_states[img_path] = self.crop_enabled.get()

    def restore_all(self):
        """Restore all lines or bars to their original state."""
        if self.check_buttons:
            self.check_buttons.disconnect(self.checkbox_cid)

        if self.selected_plot_type.get() == "TPM (Bar)":
            for patch, original_data in zip(self.axes.patches, self.original_lines_data):
                patch.set_x(original_data[0])
                patch.set_height(original_data[1])
                patch.set_visible(True)
        else:
            for line, original_data in zip(self.lines, self.original_lines_data):
                line.set_xdata(original_data[0])
                line.set_ydata(original_data[1])
                line.set_visible(True)

        if self.check_buttons:
            for i in range(len(self.line_labels)):
                if not self.check_buttons.get_status()[i]:
                    self.check_buttons.set_active(i)

            self.checkbox_cid = self.check_buttons.on_clicked(
                self.plot_manager.on_bar_checkbox_click if self.selected_plot_type.get() == "TPM (Bar)" else self.plot_manager.on_checkbox_click
            )

        if self.canvas:
            self.canvas.draw_idle()

    def generate_full_report(self):
        self.progress_dialog.show_progress_bar("Generating full report...")
        success = False

        try:
            self.root.update_idletasks()
            self.report_generator.generate_full_report(
                self.filtered_sheets, 
                self.plot_options
            )
            success = True

        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("Error", f"An error occurred: {e}")

        finally:
            if hasattr(self.report_generator, 'cleanup_temp_files'):
                self.report_generator.cleanup_temp_files()
            self.progress_dialog.hide_progress_bar()

            if success:  
                show_success_message("Success", "Full report saved successfully.", self.root)

            self.root.update_idletasks() 

    def generate_test_report(self):
        selected_sheet = self.selected_sheet.get()
    
        self.progress_dialog.show_progress_bar("Generating test report...")
        try:
            self.report_generator.generate_test_report(selected_sheet, self.filtered_sheets, self.plot_options)
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            raise
        finally:
            self.progress_dialog.hide_progress_bar()

    def open_viscosity_calculator(self):
        """Open the standalone viscosity calculator as a child window."""
        # Lazy load the viscosity GUI
        ViscosityGUI = lazy_import_viscosity_gui()
        if not ViscosityGUI:
            messagebox.showerror("Error", "Could not load viscosity calculator")
            return
        try:
            # Create and show the viscosity calculator as a child window
            viscosity_app = ViscosityGUI(parent=self.root)
            viscosity_app.show()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open viscosity calculator: {e}")

    def train_models_from_data(self):
        """Train models from data."""
        pass

    def train_models_with_chemistry(self):
        """Train models with chemistry."""
        pass

    def return_to_previous_view(self):
        """Return to the previous view from the calculator."""
        if hasattr(self, 'bottom_frame') and not self.bottom_frame.winfo_ismapped():
            self.bottom_frame.configure(height=150)
            self.bottom_frame.pack(side="bottom", fill="x", padx=10, pady=(0,10))
            self.bottom_frame.pack_propagate(False)
            self.bottom_frame.grid_propagate(False)
            self.bottom_frame.update_idletasks()
    
        if hasattr(self, 'previous_minsize'):
            self.root.minsize(*self.previous_minsize)
        else:
            self.root.minsize(1200, 800)
    
        if hasattr(self, 'previous_window_geometry'):
            self.root.geometry(self.previous_window_geometry)
            self.center_window(self.root)
    
        current_sheet = self.selected_sheet.get()
    
        if current_sheet in self.filtered_sheets:
            self.update_displayed_sheet(current_sheet)
        elif self.filtered_sheets:
            first_sheet = list(self.filtered_sheets.keys())[0]
            self.selected_sheet.set(first_sheet)
            self.update_displayed_sheet(first_sheet)
        else:
            self.clear_dynamic_frame()